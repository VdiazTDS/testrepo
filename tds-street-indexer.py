#!/usr/bin/env python3
"""
Build a local indexed streets database for TDS PAK.

Input: large roads GeoJSON FeatureCollection
Output: SQLite database with RTree index
"""

import argparse
import datetime as dt
import json
import os
import re
import sqlite3
import sys
from typing import Dict, Iterable, Iterator, List, Optional, Tuple


FEATURES_RX = re.compile(r'"features"\s*:\s*\[', re.IGNORECASE)
CHUNK_SIZE = 1024 * 1024
COORD_PRECISION = 6


def iter_feature_collection(path: str) -> Iterator[dict]:
    decoder = json.JSONDecoder()
    buf = ""
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        # Find FeatureCollection features array start.
        while True:
            chunk = f.read(CHUNK_SIZE)
            if not chunk:
                raise RuntimeError('Could not find "features" array in GeoJSON.')
            buf += chunk
            match = FEATURES_RX.search(buf)
            if match:
                buf = buf[match.end():]
                break
            if len(buf) > (CHUNK_SIZE * 2):
                buf = buf[-(CHUNK_SIZE * 2):]

        while True:
            # Skip whitespace and commas between features.
            while True:
                buf = buf.lstrip()
                if not buf:
                    chunk = f.read(CHUNK_SIZE)
                    if not chunk:
                        raise RuntimeError("Unexpected EOF while reading features array.")
                    buf += chunk
                    continue
                if buf[0] == ",":
                    buf = buf[1:]
                    continue
                break

            if buf.startswith("]"):
                return

            while True:
                try:
                    feature, consumed = decoder.raw_decode(buf)
                    break
                except json.JSONDecodeError:
                    chunk = f.read(CHUNK_SIZE)
                    if not chunk:
                        raise RuntimeError("Unexpected EOF while decoding a feature.")
                    buf += chunk

            if isinstance(feature, dict):
                yield feature
            buf = buf[consumed:]


def iter_lines(geometry: Dict) -> Iterable[List]:
    gtype = geometry.get("type")
    coords = geometry.get("coordinates")
    if gtype == "LineString" and isinstance(coords, list):
        yield coords
    elif gtype == "MultiLineString" and isinstance(coords, list):
        for line in coords:
            if isinstance(line, list):
                yield line


def normalize_tags(props: Dict) -> Dict:
    def pick(*keys: str) -> str:
        for key in keys:
            raw = props.get(key)
            if raw is None:
                continue
            text = str(raw).strip()
            if text:
                return text
        return "Unknown"

    return {
        "name": pick("name", "NAME"),
        "highway": pick("highway", "HIGHWAY", "road_class", "ROAD_CLASS", "fclass", "FCLASS"),
        "ref": pick("ref", "REF", "ref_name", "REF_NAME"),
        "maxspeed": pick("maxspeed", "MAXSPEED", "max_speed", "MAX_SPEED"),
        "lanes": pick("lanes", "LANES", "num_lanes", "NUM_LANES"),
        "surface": pick("surface", "SURFACE", "surf_type", "SURF_TYPE"),
        "oneway": pick("oneway", "ONEWAY", "one_way", "ONE_WAY"),
    }


def normalize_line_coords(raw_line: List) -> Tuple[List[List[float]], Optional[Tuple[float, float, float, float]]]:
    geom: List[List[float]] = []
    min_lat = 999.0
    max_lat = -999.0
    min_lon = 999.0
    max_lon = -999.0

    for pt in raw_line:
        if not isinstance(pt, (list, tuple)) or len(pt) < 2:
            continue
        try:
            lon = round(float(pt[0]), COORD_PRECISION)
            lat = round(float(pt[1]), COORD_PRECISION)
        except Exception:
            continue

        geom.append([lat, lon])
        min_lat = min(min_lat, lat)
        max_lat = max(max_lat, lat)
        min_lon = min(min_lon, lon)
        max_lon = max(max_lon, lon)

    if len(geom) < 2:
        return [], None
    return geom, (min_lat, max_lat, min_lon, max_lon)


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA temp_store=MEMORY;")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS streets (
          id INTEGER PRIMARY KEY,
          tags_json TEXT NOT NULL,
          geom_json TEXT NOT NULL
        );
        """
    )
    conn.execute(
        """
        CREATE VIRTUAL TABLE IF NOT EXISTS streets_rtree USING rtree(
          id,
          min_lat, max_lat,
          min_lon, max_lon
        );
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS meta (
          key TEXT PRIMARY KEY,
          value TEXT NOT NULL
        );
        """
    )
    conn.commit()


def clear_db(conn: sqlite3.Connection) -> None:
    conn.execute("DELETE FROM streets;")
    conn.execute("DELETE FROM streets_rtree;")
    conn.execute("DELETE FROM meta;")
    conn.commit()


def insert_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        "INSERT OR REPLACE INTO meta(key, value) VALUES(?, ?);",
        (key, value),
    )


def build_index(input_geojson: str, db_path: str, source_name: str, replace: bool = True) -> int:
    conn = sqlite3.connect(db_path)
    try:
        init_db(conn)
        if replace:
            clear_db(conn)

        row_id = 1
        feature_count = 0
        inserted_segments = 0
        batch_streets = []
        batch_rtree = []
        batch_size = 2500

        for feature in iter_feature_collection(input_geojson):
            feature_count += 1
            geometry = feature.get("geometry") or {}
            props = feature.get("properties") or {}
            tags_json = json.dumps(normalize_tags(props), separators=(",", ":"))

            for line in iter_lines(geometry):
                geom, bbox = normalize_line_coords(line)
                if not geom or bbox is None:
                    continue

                geom_json = json.dumps(geom, separators=(",", ":"))
                min_lat, max_lat, min_lon, max_lon = bbox
                batch_streets.append((row_id, tags_json, geom_json))
                batch_rtree.append((row_id, min_lat, max_lat, min_lon, max_lon))
                row_id += 1
                inserted_segments += 1

                if len(batch_streets) >= batch_size:
                    conn.executemany(
                        "INSERT INTO streets(id, tags_json, geom_json) VALUES(?, ?, ?);",
                        batch_streets,
                    )
                    conn.executemany(
                        "INSERT INTO streets_rtree(id, min_lat, max_lat, min_lon, max_lon) VALUES(?, ?, ?, ?, ?);",
                        batch_rtree,
                    )
                    conn.commit()
                    batch_streets.clear()
                    batch_rtree.clear()

            if feature_count % 25000 == 0:
                print(f"[Indexer] Features read: {feature_count:,} | Segments indexed: {inserted_segments:,}")

        if batch_streets:
            conn.executemany(
                "INSERT INTO streets(id, tags_json, geom_json) VALUES(?, ?, ?);",
                batch_streets,
            )
            conn.executemany(
                "INSERT INTO streets_rtree(id, min_lat, max_lat, min_lon, max_lon) VALUES(?, ?, ?, ?, ?);",
                batch_rtree,
            )
            conn.commit()

        now = dt.datetime.utcnow().isoformat() + "Z"
        insert_meta(conn, "source_name", source_name)
        insert_meta(conn, "source_path", os.path.abspath(input_geojson))
        insert_meta(conn, "row_count", str(inserted_segments))
        insert_meta(conn, "updated_at", now)
        conn.commit()

        # Finalize for read-mostly serving:
        # - checkpoint/truncate WAL
        # - switch journal back to DELETE so readers do not touch -wal/-shm files
        # This avoids Live Server auto-reload loops when backend queries the DB.
        try:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        except Exception:
            pass
        try:
            conn.execute("PRAGMA journal_mode=DELETE;")
            conn.commit()
        except Exception:
            pass

        print(f"[Indexer] Completed. Indexed segments: {inserted_segments:,}")
        print(f"[Indexer] SQLite DB: {os.path.abspath(db_path)}")
        return inserted_segments
    finally:
        conn.close()
        for suffix in ("-wal", "-shm"):
            sidecar = f"{db_path}{suffix}"
            if os.path.exists(sidecar):
                try:
                    os.remove(sidecar)
                except OSError:
                    pass


def prompt_for_input_path() -> str:
    print("")
    print("Enter full path to roads GeoJSON file:")
    print("Example: C:\\Users\\you\\Downloads\\texas-latest-free.shp-roads.geojson")
    value = input("> ").strip().strip('"')
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Index streets GeoJSON into SQLite + RTree for TDS PAK")
    parser.add_argument("input_geojson", nargs="?", help="Path to roads GeoJSON file")
    parser.add_argument("--db", default="tds-streets.sqlite", help="Output SQLite DB path")
    parser.add_argument("--source-name", default="", help="Source label stored in DB metadata")
    parser.add_argument("--append", action="store_true", help="Append instead of replacing DB contents")
    args = parser.parse_args()

    input_path = (args.input_geojson or "").strip()
    if not input_path:
        input_path = prompt_for_input_path()
    if not input_path:
        print("No input file provided.", file=sys.stderr)
        return 1

    if not os.path.exists(input_path):
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 1

    source_name = args.source_name.strip() or os.path.basename(input_path)
    try:
        build_index(
            input_geojson=input_path,
            db_path=args.db,
            source_name=source_name,
            replace=not args.append,
        )
        return 0
    except Exception as exc:
        print("")
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""
Serve indexed local street segments to TDS PAK from SQLite + RTree.

Endpoints:
- GET /api/health
- GET /api/streets?south=..&west=..&north=..&east=..&limit=..
"""

import argparse
import json
import os
import sqlite3
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, Optional
from urllib.parse import parse_qs, urlparse


def parse_float(name: str, raw: Optional[str]) -> float:
    if raw is None or str(raw).strip() == "":
        raise ValueError(f"Missing query parameter: {name}")
    try:
        return float(raw)
    except Exception as exc:
        raise ValueError(f"Invalid float for {name}: {raw}") from exc


def parse_int(name: str, raw: Optional[str], default_value: int, min_value: int, max_value: int) -> int:
    if raw is None or str(raw).strip() == "":
        return default_value
    try:
        value = int(raw)
    except Exception as exc:
        raise ValueError(f"Invalid integer for {name}: {raw}") from exc
    if value < min_value:
        return min_value
    if value > max_value:
        return max_value
    return value


def validate_bbox(south: float, west: float, north: float, east: float) -> None:
    if not (-90 <= south <= 90 and -90 <= north <= 90):
        raise ValueError("Latitude must be within [-90, 90].")
    if not (-180 <= west <= 180 and -180 <= east <= 180):
        raise ValueError("Longitude must be within [-180, 180].")
    if north <= south:
        raise ValueError("north must be greater than south.")
    if east <= west:
        raise ValueError("east must be greater than west.")


class StreetBackendContext:
    def __init__(self, db_path: str, max_limit: int):
        self.db_path = os.path.abspath(db_path)
        self.max_limit = max(1000, int(max_limit))

    def _connect(self) -> sqlite3.Connection:
        db_uri_immutable = f"{Path(self.db_path).as_uri()}?mode=ro&immutable=1"
        db_uri_readonly = f"{Path(self.db_path).as_uri()}?mode=ro"
        try:
            conn = sqlite3.connect(db_uri_immutable, uri=True, timeout=30, check_same_thread=False)
        except sqlite3.OperationalError:
            conn = sqlite3.connect(db_uri_readonly, uri=True, timeout=30, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA busy_timeout=5000;")
        conn.execute("PRAGMA query_only=ON;")
        conn.execute("PRAGMA temp_store=MEMORY;")
        return conn

    def _table_exists(self, conn: sqlite3.Connection, name: str) -> bool:
        row = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type IN ('table', 'view') AND name = ? LIMIT 1;",
            (name,),
        ).fetchone()
        return bool(row)

    def read_health(self) -> Dict:
        if not os.path.exists(self.db_path):
            return {
                "ok": True,
                "has_index": False,
                "db_path": self.db_path,
                "source_name": "",
                "row_count": 0,
                "updated_at": "",
            }

        try:
            conn = self._connect()
        except Exception as exc:
            return {
                "ok": False,
                "has_index": False,
                "db_path": self.db_path,
                "source_name": "",
                "row_count": 0,
                "updated_at": "",
                "error": str(exc),
            }

        try:
            has_streets = self._table_exists(conn, "streets")
            has_rtree = self._table_exists(conn, "streets_rtree")
            has_index = bool(has_streets and has_rtree)

            source_name = ""
            updated_at = ""
            row_count = 0

            has_meta = self._table_exists(conn, "meta")
            if has_meta:
                rows = conn.execute("SELECT key, value FROM meta;").fetchall()
                meta = {str(r["key"]): str(r["value"]) for r in rows}
                source_name = meta.get("source_name", "")
                updated_at = meta.get("updated_at", "")
                try:
                    row_count = int(meta.get("row_count", "0"))
                except Exception:
                    row_count = 0

            if has_index and row_count <= 0:
                row = conn.execute("SELECT COUNT(*) AS c FROM streets;").fetchone()
                row_count = int(row["c"] if row and row["c"] is not None else 0)

            return {
                "ok": True,
                "has_index": has_index,
                "db_path": self.db_path,
                "source_name": source_name,
                "row_count": row_count,
                "updated_at": updated_at,
            }
        finally:
            conn.close()

    def stream_streets_for_bbox(self, handler: BaseHTTPRequestHandler, south: float, west: float, north: float, east: float, limit: int) -> None:
        if not os.path.exists(self.db_path):
            raise RuntimeError(f"Database not found: {self.db_path}")

        conn = self._connect()
        try:
            has_streets = self._table_exists(conn, "streets")
            has_rtree = self._table_exists(conn, "streets_rtree")
            if not (has_streets and has_rtree):
                raise RuntimeError("Database is missing streets/streets_rtree tables. Run indexer first.")

            sql = """
                SELECT s.id, s.tags_json, s.geom_json
                FROM streets_rtree r
                JOIN streets s ON s.id = r.id
                WHERE r.max_lat >= ?
                  AND r.min_lat <= ?
                  AND r.max_lon >= ?
                  AND r.min_lon <= ?
                ORDER BY s.id
                LIMIT ?;
            """
            cursor = conn.execute(sql, (south, north, west, east, limit))

            payload_head = (
                '{"ok":true,"elements":['
            ).encode("utf-8")
            handler.wfile.write(payload_head)

            count = 0
            first = True
            fetch_size = 2000

            while True:
                rows = cursor.fetchmany(fetch_size)
                if not rows:
                    break
                for row in rows:
                    if not first:
                        handler.wfile.write(b",")
                    first = False

                    row_id = int(row["id"])
                    tags_json = row["tags_json"] or "{}"
                    geom_json = row["geom_json"] or "[]"
                    item = (
                        '{"id":'
                        + str(row_id)
                        + ',"tags":'
                        + str(tags_json)
                        + ',"geom":'
                        + str(geom_json)
                        + "}"
                    ).encode("utf-8")
                    handler.wfile.write(item)
                    count += 1

            tail = (
                '],"count":'
                + str(count)
                + ',"truncated":'
                + ("true" if count >= limit else "false")
                + "}"
            ).encode("utf-8")
            handler.wfile.write(tail)
        finally:
            conn.close()


class StreetBackendHandler(BaseHTTPRequestHandler):
    server_version = "TDSStreetBackend/1.0"
    protocol_version = "HTTP/1.1"

    def _ctx(self) -> StreetBackendContext:
        return self.server.context  # type: ignore[attr-defined]

    def _send_json(self, status: int, payload: Dict) -> None:
        data = json.dumps(payload, separators=(",", ":")).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(data)

    def _send_stream_headers(self) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Connection", "close")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self._stream_headers_sent = True

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Max-Age", "86400")
        self.end_headers()

    def do_GET(self) -> None:
        self._stream_headers_sent = False
        parsed = urlparse(self.path)
        path = parsed.path or "/"
        if path == "/api/health":
            health = self._ctx().read_health()
            status = HTTPStatus.OK if health.get("ok", False) else HTTPStatus.SERVICE_UNAVAILABLE
            self._send_json(status, health)
            return

        if path == "/api/streets":
            params = parse_qs(parsed.query or "")
            try:
                south = parse_float("south", (params.get("south") or [None])[0])
                west = parse_float("west", (params.get("west") or [None])[0])
                north = parse_float("north", (params.get("north") or [None])[0])
                east = parse_float("east", (params.get("east") or [None])[0])
                validate_bbox(south, west, north, east)
                limit = parse_int(
                    "limit",
                    (params.get("limit") or [None])[0],
                    default_value=120000,
                    min_value=1000,
                    max_value=self._ctx().max_limit,
                )
            except ValueError as exc:
                self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})
                return

            try:
                self._send_stream_headers()
                self._ctx().stream_streets_for_bbox(self, south, west, north, east, limit)
            except BrokenPipeError:
                return
            except ConnectionResetError:
                return
            except Exception as exc:
                if not getattr(self, "_stream_headers_sent", False) and not self.wfile.closed:
                    self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": str(exc)})
                else:
                    print(f"[backend] stream failure: {exc}")
            return

        self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found"})

    def log_message(self, fmt: str, *args) -> None:
        # Keep logs concise for long-running local process.
        print(f"[backend] {self.address_string()} - {fmt % args}")


def main() -> int:
    parser = argparse.ArgumentParser(description="TDS PAK local street backend")
    parser.add_argument("--db", default="tds-streets.sqlite", help="Indexed streets SQLite DB path")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host")
    parser.add_argument("--port", type=int, default=8787, help="Bind port")
    parser.add_argument("--max-limit", type=int, default=300000, help="Upper cap for /api/streets limit")
    args = parser.parse_args()

    context = StreetBackendContext(db_path=args.db, max_limit=args.max_limit)
    server = ThreadingHTTPServer((args.host, int(args.port)), StreetBackendHandler)
    server.context = context  # type: ignore[attr-defined]

    print("")
    print("[TDS Backend] Starting local street backend")
    print(f"[TDS Backend] DB: {os.path.abspath(args.db)}")
    print(f"[TDS Backend] URL: http://{args.host}:{int(args.port)}")
    print("[TDS Backend] Health: /api/health")
    print("[TDS Backend] Streets: /api/streets?south=..&west=..&north=..&east=..&limit=120000")
    print("")

    try:
        server.serve_forever(poll_interval=0.5)
    except KeyboardInterrupt:
        print("\n[TDS Backend] Stopping...")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

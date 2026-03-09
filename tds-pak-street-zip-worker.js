/* eslint-disable no-restricted-globals */

const ROAD_LAYER_RX = /(road|street|highway)/i;
const ROAD_LAYER_STRONG_RX = /(roads?|highway|street|transport|line)/i;
const FCLASS_WHITELIST = new Set([
  "motorway",
  "motorway_link",
  "trunk",
  "trunk_link",
  "primary",
  "primary_link",
  "secondary",
  "secondary_link",
  "tertiary",
  "tertiary_link",
  "residential",
  "unclassified",
  "living_street",
  "service",
  "track",
  "road",
  "pedestrian",
  "footway",
  "cycleway",
  "bridleway",
  "path",
  "steps"
]);

function postProgress(percent, stage) {
  self.postMessage({
    type: "progress",
    percent: Number.isFinite(percent) ? Math.max(0, Math.min(100, Math.round(percent))) : 0,
    stage: String(stage || "Converting streets ZIP...")
  });
}

function collectFeatureCollections(parsedSource) {
  const collections = [];
  const pushCollection = (value, layerName = "") => {
    if (!value || value.type !== "FeatureCollection" || !Array.isArray(value.features)) return;
    collections.push({ layerName: String(layerName || ""), features: value.features });
  };

  if (!parsedSource || typeof parsedSource !== "object") return collections;

  if (parsedSource.type === "FeatureCollection") {
    pushCollection(parsedSource, parsedSource.name || "");
    return collections;
  }

  if (Array.isArray(parsedSource)) {
    parsedSource.forEach((item, idx) => {
      if (item?.type === "FeatureCollection") {
        pushCollection(item, item.name || `layer_${idx + 1}`);
      } else if (item?.type === "Feature") {
        pushCollection({ type: "FeatureCollection", features: [item] }, `layer_${idx + 1}`);
      }
    });
    return collections;
  }

  Object.entries(parsedSource).forEach(([key, value]) => {
    if (value?.type === "FeatureCollection") {
      pushCollection(value, key);
      return;
    }
    if (Array.isArray(value) && value.length && value.every(feature => feature?.type === "Feature")) {
      pushCollection({ type: "FeatureCollection", features: value }, key);
    }
  });

  return collections;
}

function isStreetFeature(feature, layerName = "") {
  const geometryType = feature?.geometry?.type;
  if (geometryType !== "LineString" && geometryType !== "MultiLineString") return false;

  const props = feature?.properties || {};
  const highway = String(props.highway || props.HIGHWAY || "").trim();
  if (highway) return true;

  const fclass = String(props.fclass || props.FCLASS || "").trim().toLowerCase();
  if (fclass) return FCLASS_WHITELIST.has(fclass);

  if (ROAD_LAYER_RX.test(String(layerName || ""))) return true;
  return true;
}

function extractStreetFeatures(parsedSource, sourceLabel) {
  const collections = collectFeatureCollections(parsedSource);
  if (!collections.length) return [];

  const roadCollections = collections.filter(c => ROAD_LAYER_RX.test(c.layerName));
  const sourceCollections = roadCollections.length ? roadCollections : collections;

  let total = 0;
  sourceCollections.forEach(c => {
    total += Array.isArray(c.features) ? c.features.length : 0;
  });

  const features = [];
  let processed = 0;
  const reportEvery = 1000;

  sourceCollections.forEach(collection => {
    const list = Array.isArray(collection.features) ? collection.features : [];
    list.forEach(feature => {
      processed += 1;
      if (isStreetFeature(feature, collection.layerName)) {
        features.push(feature);
      }
      if (processed % reportEvery === 0) {
        const ratio = total > 0 ? (processed / total) : 0;
        postProgress(
          66 + (Math.min(1, ratio) * 22),
          `Extracting road features from ${sourceLabel}...`
        );
      }
    });
  });

  return features;
}

function getShapefileBaseNamesFromZip(zip) {
  const bases = new Set();
  Object.keys(zip.files || {}).forEach((name) => {
    const entry = zip.files[name];
    if (!entry || entry.dir) return;
    if (!/\.shp$/i.test(name)) return;
    bases.add(name.replace(/\.shp$/i, ""));
  });
  return [...bases];
}

function buildZipFileLookup(zip) {
  const map = new Map();
  Object.keys(zip.files || {}).forEach((name) => {
    map.set(String(name || "").toLowerCase(), name);
  });
  return map;
}

function getZipEntryCaseInsensitive(zip, lookup, path) {
  const direct = zip.file(path);
  if (direct) return direct;
  const resolved = lookup.get(String(path || "").toLowerCase());
  if (!resolved) return null;
  return zip.file(resolved);
}

function scoreRoadLayerBaseName(baseName) {
  const lower = String(baseName || "").toLowerCase();
  const short = lower.split("/").pop() || lower;
  let score = 0;
  if (short === "gis_osm_roads_free_1") score += 1200;
  if (short === "gis_osm_highways_free_1") score += 900;
  if (short.includes("roads_free_1")) score += 700;
  if (short.includes("roads_free")) score += 520;
  if (short.endsWith("_roads") || short.endsWith("roads")) score += 300;
  if (/\broads?\b/.test(short) || short.includes("_roads_")) score += 260;
  if (/\bhighway\b/.test(short) || short.includes("_highway_")) score += 200;
  if (/\bstreet\b/.test(short) || short.includes("_street_")) score += 140;
  if (/\btransport\b/.test(lower) || lower.includes("_transport_")) score += 90;
  if (/\bline\b/.test(lower) || lower.includes("_line_")) score += 40;
  if (lower.includes("traffic")) score -= 60;
  if (lower.includes("rail")) score -= 60;
  if (lower.includes("water")) score -= 80;
  if (lower.includes("building")) score -= 80;
  if (short.includes("pois")) score -= 80;
  if (short.includes("places")) score -= 80;
  if (short.includes("landuse")) score -= 80;
  return score;
}

function pickRoadLayerCandidates(rankedItems) {
  const withShort = rankedItems.map(item => ({
    ...item,
    short: String(item.base || "").split("/").pop() || String(item.base || "")
  }));
  const selected = [];
  const seen = new Set();
  const addMatches = (pattern) => {
    withShort.forEach(item => {
      if (!pattern.test(item.short)) return;
      if (seen.has(item.base)) return;
      seen.add(item.base);
      selected.push(item);
    });
  };

  addMatches(/^gis_osm_roads_free_1$/i);
  addMatches(/^gis_osm_highways_free_1$/i);
  addMatches(/roads?_free_1/i);
  addMatches(/\broads?\b/i);
  addMatches(/\bhighways?\b/i);
  addMatches(/\bstreets?\b/i);

  if (selected.length) return selected.slice(0, 3);

  const strong = withShort.filter(item => item.score >= 200);
  if (strong.length) return strong.slice(0, 3);

  return withShort.slice(0, 2);
}

async function parseLayerWithShpLib(shpBuffer, dbfBuffer, prjText = "", cpgText = "") {
  if (typeof self.shp !== "function") {
    throw new Error("shp() converter unavailable.");
  }

  try {
    return await self.shp({
      shp: shpBuffer,
      dbf: dbfBuffer,
      prj: prjText || undefined,
      cpg: cpgText || undefined
    });
  } catch {
    // Fall through to low-level parsing path below.
  }

  if (typeof self.shp.parseShp === "function" &&
      typeof self.shp.parseDbf === "function" &&
      typeof self.shp.combine === "function") {
    const shpPart = self.shp.parseShp(shpBuffer, prjText || undefined);
    const dbfPart = self.shp.parseDbf(dbfBuffer, cpgText || undefined);
    return self.shp.combine([shpPart, dbfPart]);
  }

  throw new Error("No compatible shapefile parsing method available in worker.");
}

async function parseRoadLayersFromZipBuffer(zipBuffer, sourceLabel) {
  if (typeof self.JSZip !== "function") {
    importScripts("https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js");
  }

  postProgress(10, `Opening ${sourceLabel}...`);
  const zip = await self.JSZip.loadAsync(zipBuffer);
  const zipLookup = buildZipFileLookup(zip);
  postProgress(13, `Scanning layers in ${sourceLabel}...`);

  const baseNames = getShapefileBaseNamesFromZip(zip);
  if (!baseNames.length) {
    throw new Error("No .shp layers found in ZIP.");
  }

  const ranked = [...baseNames]
    .map(base => ({ base, score: scoreRoadLayerBaseName(base) }))
    .sort((a, b) => b.score - a.score);

  const likelyRoadLayers = ranked.filter(item => item.score > 0 || ROAD_LAYER_STRONG_RX.test(item.base));
  const candidateItems = pickRoadLayerCandidates(likelyRoadLayers.length ? likelyRoadLayers : ranked);
  let lastError = null;

  for (let i = 0; i < candidateItems.length; i++) {
    const item = candidateItems[i];
    const base = item.base;
    const shpName = `${base}.shp`;
    const dbfName = `${base}.dbf`;
    const shpEntry = getZipEntryCaseInsensitive(zip, zipLookup, shpName);
    const dbfEntry = getZipEntryCaseInsensitive(zip, zipLookup, dbfName);
    if (!shpEntry || !dbfEntry) continue;

    try {
      const phaseStart = 16 + (i * 10);
      postProgress(phaseStart, `Converting road layer ${i + 1}/${candidateItems.length}...`);
      const shpBuffer = await shpEntry.async("arraybuffer");
      const dbfBuffer = await dbfEntry.async("arraybuffer");
      const prjEntry = getZipEntryCaseInsensitive(zip, zipLookup, `${base}.prj`);
      const cpgEntry = getZipEntryCaseInsensitive(zip, zipLookup, `${base}.cpg`);
      let prjText = "";
      let cpgText = "";
      if (prjEntry) {
        prjText = (await prjEntry.async("text")) || "";
      }
      if (cpgEntry) {
        cpgText = (await cpgEntry.async("text")) || "";
      }
      const parsed = await parseLayerWithShpLib(shpBuffer, dbfBuffer, prjText, cpgText);
      const extracted = extractStreetFeatures(parsed, sourceLabel);
      postProgress(Math.min(58, phaseStart + 8), `Converted ${base.split("/").pop()} (${extracted.length.toLocaleString()} roads)`);
      if (extracted.length) {
        // Return immediately on first successful road-like layer for speed/reliability.
        if (item.score >= 120 || /\broads?\b/i.test(base) || /highway/i.test(base)) {
          return extracted;
        }
        return extracted;
      }
    } catch (err) {
      lastError = err;
    }
  }

  if (lastError) throw lastError;
  throw new Error("Could not parse a roads shapefile layer from ZIP.");
}

self.onmessage = async (event) => {
  const data = event?.data || {};
  if (data.type !== "parseZip") return;

  const label = String(data.label || "streets ZIP");
  const zipBytes = Number(data?.buffer?.byteLength || 0);
  const LARGE_ZIP_BYTES = 150 * 1024 * 1024;

  try {
    if (typeof self.shp !== "function") {
      importScripts("https://unpkg.com/shpjs@6.2.0/dist/shp.min.js");
    }

    let features = [];
    try {
      features = await parseRoadLayersFromZipBuffer(data.buffer, label);
    } catch (layerErr) {
      if (zipBytes >= LARGE_ZIP_BYTES) {
        throw new Error(
          `Road-layer parsing failed for a large ZIP (${Math.round(zipBytes / (1024 * 1024))} MB). ` +
          "Compatibility full-ZIP conversion is disabled for large files because it stalls in browsers. " +
          "Please retry with a roads-only file."
        );
      }
      postProgress(18, `Trying compatibility conversion path for ${label}...`);
      const parsed = await self.shp(data.buffer);
      postProgress(66, `Extracting road features from ${label}...`);
      features = extractStreetFeatures(parsed, label);
      if (!features.length) {
        throw layerErr;
      }
    }
    postProgress(89, `Extracted ${features.length.toLocaleString()} road features from ${label}`);

    self.postMessage({
      type: "result",
      features
    });
  } catch (err) {
    self.postMessage({
      type: "error",
      message: err?.message || String(err)
    });
  }
};

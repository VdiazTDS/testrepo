window.addEventListener("error", e => {
  console.error("JS ERROR:", e.message, "at line", e.lineno);
});

// Fallback: always try to wire Print Center after full load.
window.addEventListener("load", () => {
  try {
    if (typeof initPrintCenterControls === "function") {
      initPrintCenterControls();
    }
  } catch (err) {
    console.warn("Print Center init fallback failed:", err);
  }
});

let layerVisibilityState = {};
let selectedLayerKey = null;

// ================= SUPABASE CONFIG =================
// Connection info for cloud file storage
const SUPABASE_URL = "https://lffazhbwvorwxineklsy.supabase.co";
const SUPABASE_KEY = "sb_publishable_Lfh2zlIiTSMB0U-Fe5o6Jg_mJ1qkznh";
const BUCKET = "excel-files";
// ===== CURRENT EXCEL STATE =====
window._currentRows = null;
window._currentWorkbook = null;
window._currentFilePath = null;
window._summaryRows = [];
window._summaryHeaders = [];
window._attributeHeaders = [];

window.streetLabelsEnabled = false;
const attributeRowToId = new WeakMap();
let attributeRowToMarker = new WeakMap();
const attributeMarkerByRowId = new Map();
const attributeState = {
  sortKey: null,
  sortDir: 1,
  filterText: "",
  selectedOnly: false,
  page: 1,
  pageSize: 300,
  selectedRowIds: new Set(),
  lastVisibleRows: []
};
const APP_STORAGE_NS = (() => {
  const path = (window.location.pathname || "").toLowerCase();
  if (path.endsWith("/tds-pak.html") || path.endsWith("tds-pak.html")) return "tds-pak";
  if (path.endsWith("/sales-polygon-viewer.html") || path.endsWith("sales-polygon-viewer.html")) return "sales-polygon-viewer";
  if (path.endsWith("/solution-reviewer.html") || path.endsWith("solution-reviewer.html")) return "solution-reviewer";
  return "cart-delivery";
})();

function storageKey(name) {
  return `${APP_STORAGE_NS}:${name}`;
}

function storageGet(name) {
  return localStorage.getItem(storageKey(name));
}

function storageSet(name, value) {
  localStorage.setItem(storageKey(name), String(value));
}

// Header tools menu links (shared by all app pages).
const HEADER_TOOL_LINKS = [
  { label: "Cart Delivery App", href: "./index.html" },
  { label: "Sales-Polygon Viewer", href: "./sales-polygon-viewer.html" },
  { label: "Solution Reviewer", href: "./solution-reviewer.html" },
  { label: "TDS-PAK", href: "./tds-pak.html" }
];

function setupHeaderToolsMenu() {
  const setupToolsMenuInstance = (btnId, dropdownId, listId) => {
    const menuBtn = document.getElementById(btnId);
    const menuDropdown = document.getElementById(dropdownId);
    const menuList = document.getElementById(listId);

    if (!menuBtn || !menuDropdown || !menuList) return;

    menuList.innerHTML = "";

    HEADER_TOOL_LINKS.forEach(tool => {
      const link = document.createElement("a");
      link.className = "tools-menu-item";
      link.textContent = tool.label;
      link.href = tool.href || "#";

      if (!tool.href || tool.href === "#") {
        link.addEventListener("click", e => e.preventDefault());
      }

      menuList.appendChild(link);
    });

    const closeMenu = () => {
      menuDropdown.classList.remove("open");
      menuBtn.setAttribute("aria-expanded", "false");
    };

    menuBtn.addEventListener("click", e => {
      e.stopPropagation();
      const isOpen = menuDropdown.classList.toggle("open");
      menuBtn.setAttribute("aria-expanded", isOpen ? "true" : "false");
    });

    menuDropdown.addEventListener("click", e => e.stopPropagation());
    document.addEventListener("click", closeMenu);
    document.addEventListener("keydown", e => {
      if (e.key === "Escape") closeMenu();
    });
  };

  setupToolsMenuInstance("toolsMenuBtn", "toolsMenuDropdown", "toolsMenuList");
  setupToolsMenuInstance("toolsMenuBtnMobile", "toolsMenuDropdownMobile", "toolsMenuListMobile");
}

//======
// 🔐 Delete protection password I know this is not secure, I just wanted to make it harder for ppl to accidentally delete files. You can change or remove this as needed.
const DELETE_PASSWORD = "Austin1";  // ← change to whatever you want


document.addEventListener("DOMContentLoaded", () => {
  initApp();
  document.addEventListener("DOMContentLoaded", initApp);
// ================= SUN MODE TOGGLE =================

const sunToggle = document.getElementById("sunModeToggle");
const sunToggleText = document.getElementById("sunToggleText");

function updateSunToggleText() {
  if (!sunToggle || !sunToggleText) return;
  sunToggleText.textContent = sunToggle.checked ? "Light Mode" : "Dark Mode";
}

// Load saved preference
if (storageGet("sunMode") === "on") {
  document.body.classList.add("sun-mode");
  if (sunToggle) sunToggle.checked = true;
}

updateSunToggleText();

if (sunToggle) {
  sunToggle.addEventListener("change", () => {
    if (sunToggle.checked) {
      document.body.classList.add("sun-mode");
      storageSet("sunMode", "on");
    } else {
      document.body.classList.remove("sun-mode");
      storageSet("sunMode", "off");
    }
    updateSunToggleText();
  });
}

setupHeaderToolsMenu();

  });
/* ⭐ Ensures mobile buttons move AFTER full page load */
window.addEventListener("load", placeLocateButton);

// ===== USER GEOLOCATION =====
function locateUser() {
  if (!navigator.geolocation) {
    console.warn("Geolocation not supported");
    return;
  }

  navigator.geolocation.getCurrentPosition(
    pos => {
      const lat = pos.coords.latitude;
      const lon = pos.coords.longitude;

      // Center map on user
      map.setView([lat, lon], 14);
    },
    err => {
      console.warn("Location permission denied or unavailable");
      map.setView([39.5, -98.35], 4);
    },
    {
      enableHighAccuracy: true,
      timeout: 10000,
      maximumAge: 30000
    }
  );
}

// ===== FLOATING "CENTER ON ME" BUTTON =====
let watchId = null;
let userCircle = null;



function startLiveTracking() {
  if (!navigator.geolocation) {
    alert("Geolocation is not supported on this device.");
    return;
  }

  // Stop previous tracking
  if (watchId !== null) {
    navigator.geolocation.clearWatch(watchId);
  }

// Start compass tracking first
startHeadingTracking();

watchId = navigator.geolocation.watchPosition(
  (pos) => {
    const lat = pos.coords.latitude;
    const lng = pos.coords.longitude;
    const accuracy = pos.coords.accuracy;

    const latlng = [lat, lng];

    // ===== Heading Arrow =====
    if (!headingMarker) {
      headingMarker = L.marker(latlng, {
        icon: createHeadingIcon(currentHeading),
        interactive: false
      }).addTo(map);
    } else {
      headingMarker.setLatLng(latlng);
    }

    // Smooth follow
    map.flyTo(latlng, Math.max(map.getZoom(), 16), { duration: 1.2 });

   

    // ===== Accuracy circle =====
    if (!userCircle) {
      userCircle = L.circle(latlng, {
        radius: accuracy,
        color: "#2a93ff",
        fillColor: "#2a93ff",
        fillOpacity: 0.2,
        weight: 2,
      }).addTo(map);
    } else {
      userCircle.setLatLng(latlng);
      userCircle.setRadius(accuracy);
    }
  },
  (err) => {
    console.error("GPS error:", err);
    alert("Unable to get your location.");
  },
  {
    enableHighAccuracy: true,
    maximumAge: 0,
    timeout: 10000,
  }
);

}


//===direction user is facing
let headingMarker = null;
let currentHeading = 0;

function createHeadingIcon(angle) {
  return L.divIcon({
    className: "heading-icon-modern",
    html: `
      <div style="
        transform: rotate(${angle}deg);
        transition: transform 0.12s linear;
        display: flex;
        align-items: center;
        justify-content: center;
      ">
        <svg width="36" height="36" viewBox="0 0 36 36">
          <circle cx="18" cy="18" r="14" fill="rgba(66,165,245,0.15)" />
          <circle cx="18" cy="18" r="10" fill="#ffffff" />
          <path d="M18 6 L24 22 L18 19 L12 22 Z" fill="#42a5f5"/>
        </svg>
      </div>
    `,
    iconSize: [36, 36],
    iconAnchor: [18, 18]
  });
}





function startHeadingTracking() {
  if (typeof DeviceOrientationEvent !== "undefined") {

    if (typeof DeviceOrientationEvent.requestPermission === "function") {
      DeviceOrientationEvent.requestPermission()
        .then(permissionState => {
          if (permissionState === "granted") {
            window.addEventListener("deviceorientation", updateHeading);
          }
        })
        .catch(console.error);
    } else {
      window.addEventListener("deviceorientation", updateHeading);
    }

  }
}




function updateHeading(event) {
  if (event.alpha === null) return;

  currentHeading = 360 - event.alpha; // Convert to compass style

  if (headingMarker) {
    headingMarker.setIcon(createHeadingIcon(currentHeading));
  }
}



// ===== HARD REFRESH BUTTON (SAFE + NO CACHE) =====
const hardRefreshBtn = document.getElementById("hardRefreshBtn");

if (hardRefreshBtn) {
  let refreshArmed = false;

  hardRefreshBtn.addEventListener("click", async () => {

    // Mobile double-tap protection
    if (window.innerWidth <= 900) {
      if (!refreshArmed) {
        refreshArmed = true;
        hardRefreshBtn.textContent = "Tap again to refresh app";

        setTimeout(() => {
          refreshArmed = false;
          hardRefreshBtn.textContent = "Refresh App";
        }, 2000);

        return;
      }
    }

    // Desktop confirmation
    if (window.innerWidth > 900) {
      const confirmed = confirm(
        "Refresh App will clear this app's local cached files and reload the page with fresh data. Continue?"
      );
      if (!confirmed) return;
    }

    // Clear cache storage if supported
    if ("caches" in window) {
      const names = await caches.keys();
      await Promise.all(names.map(n => caches.delete(n)));
    }

    // Unregister service workers so they cannot serve stale assets
    if ("serviceWorker" in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(regs.map(r => r.unregister()));
    }

    // True hard reload (cache-busting URL)
    const url = new URL(window.location.href);
    url.searchParams.set("_cb", String(Date.now()));
    window.location.replace(url.toString());
  });
}

function placeRefreshButton() {
  const refreshBtn = document.getElementById("hardRefreshBtn");
  const desktopTools = document.querySelector(".header-tools-desktop");
  const mobileButtons = document.querySelector(".mobile-header-buttons");

  if (!refreshBtn || !desktopTools || !mobileButtons) return;

  if (window.innerWidth > 900) {
    desktopTools.appendChild(refreshBtn);
  } else {
    mobileButtons.insertBefore(refreshBtn, mobileButtons.firstChild);
  }
}

placeRefreshButton();
window.addEventListener("resize", placeRefreshButton);

//======


// Create Supabase client
const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);


// ================= FILE NAME MATCHING =================
// Makes route files and route summary files match even if
// spacing, punctuation, or "RouteSummary" text is different.
function normalizeName(name) {
  return name
    .toLowerCase()
    .replace(".xlsx", "")
    .replace("route summary", "")   // handles "Route Summary"
    .replace("routesummary", "")    // handles "RouteSummary"
    .replace(/[_\s.-]/g, "")        // ignore spaces, _, ., -
    .trim();
}

function isRouteSummaryFileName(fileName) {
  const lower = String(fileName || "").toLowerCase();
  return lower.includes("routesummary") || lower.includes("route summary");
}

const SUMMARY_ATTACH_STORAGE_KEY = storageKey("summaryAttachments");

function getSummaryAttachments() {
  try {
    const raw = localStorage.getItem(SUMMARY_ATTACH_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (_) {
    return {};
  }
}

function setSummaryAttachments(map) {
  localStorage.setItem(SUMMARY_ATTACH_STORAGE_KEY, JSON.stringify(map || {}));
}

function setRouteSummaryAttachment(routeFileName, summaryFileName) {
  const map = getSummaryAttachments();
  Object.keys(map).forEach(routeKey => {
    if (map[routeKey] === summaryFileName) delete map[routeKey];
  });
  map[routeFileName] = summaryFileName;
  setSummaryAttachments(map);
}

function removeRouteSummaryAttachment(routeFileName) {
  const map = getSummaryAttachments();
  delete map[routeFileName];
  setSummaryAttachments(map);
}

function cleanupSummaryAttachments(existingFileNames) {
  const existing = new Set(existingFileNames || []);
  const map = getSummaryAttachments();
  let dirty = false;

  Object.entries(map).forEach(([routeName, summaryName]) => {
    if (!existing.has(routeName) || !existing.has(summaryName)) {
      delete map[routeName];
      dirty = true;
    }
  });

  if (dirty) setSummaryAttachments(map);
  return map;
}

function resolveSummaryForRoute(routeFileName, filesList) {
  const fileNames = (filesList || []).map(f => f.name);
  const attachmentMap = cleanupSummaryAttachments(fileNames);
  const attachedSummary = attachmentMap[routeFileName];

  if (attachedSummary && fileNames.includes(attachedSummary)) {
    return attachedSummary;
  }

  const normalizedRoute = normalizeName(routeFileName);
  const fallback = (filesList || []).find(f => {
    if (!isRouteSummaryFileName(f.name)) return false;
    return normalizeName(f.name) === normalizedRoute;
  });

  return fallback ? fallback.name : null;
}

function openSummaryAttachModal(summaryFileName, routeFileNames) {
  return new Promise(resolve => {
    const modal = document.getElementById("summaryAttachModal");
    const text = document.getElementById("summaryAttachModalText");
    const selectedLabel = document.getElementById("summaryAttachSelectedRoute");
    const searchInput = document.getElementById("summaryAttachSearch");
    const listBox = document.getElementById("summaryAttachList");
    const cancelBtn = document.getElementById("summaryAttachCancel");
    const confirmBtn = document.getElementById("summaryAttachConfirm");

    if (!modal || !text || !selectedLabel || !searchInput || !listBox || !cancelBtn || !confirmBtn) {
      resolve(null);
      return;
    }

    text.textContent = `Choose which saved route file should attach to: ${summaryFileName}`;
    let selectedRoute = null;
    const sortedRoutes = routeFileNames
      .slice()
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

    const setSelected = routeName => {
      selectedRoute = routeName;
      selectedLabel.textContent = routeName || "None";
      confirmBtn.disabled = !routeName;
      [...listBox.querySelectorAll(".attach-route-item")].forEach(btn => {
        btn.classList.toggle("selected", btn.dataset.routeName === routeName);
      });
    };

    const renderList = filterText => {
      const needle = String(filterText || "").trim().toLowerCase();
      listBox.innerHTML = "";
      const shown = sortedRoutes.filter(name => name.toLowerCase().includes(needle));

      if (!shown.length) {
        const empty = document.createElement("div");
        empty.className = "attach-route-empty";
        empty.textContent = "No matching route files.";
        listBox.appendChild(empty);
        return;
      }

      shown.forEach(routeName => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "attach-route-item";
        btn.dataset.routeName = routeName;
        btn.textContent = routeName;
        btn.addEventListener("click", () => setSelected(routeName));
        listBox.appendChild(btn);
      });

      if (selectedRoute && shown.includes(selectedRoute)) {
        setSelected(selectedRoute);
      } else if (!selectedRoute) {
        setSelected(shown[0]);
      } else {
        confirmBtn.disabled = true;
      }
    };

    searchInput.value = "";
    selectedLabel.textContent = "None";
    confirmBtn.disabled = true;
    renderList("");
    searchInput.oninput = () => renderList(searchInput.value);

    const close = value => {
      modal.style.display = "none";
      searchInput.oninput = null;
      cancelBtn.onclick = null;
      confirmBtn.onclick = null;
      resolve(value);
    };

    cancelBtn.onclick = () => close(null);
    confirmBtn.onclick = () => close(selectedRoute || null);
    modal.style.display = "flex";
    searchInput.focus();
  });
}

// Prevent recursive growth like "..._Downloaded_<ts>_Downloaded_<ts>".
function getDownloadBaseName(filePath) {
  const rawName = (filePath || "Export").replace(/\.[^/.]+$/, "");
  return rawName.replace(
    /(?:_Downloaded_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})+$/i,
    ""
  );
}

function setCurrentFileDisplay(filePath) {
  const label = document.getElementById("currentFileDisplay");
  const name = document.getElementById("currentFileName");
  const displayName = filePath || "None";
  if (!label) return;

  if (name) {
    name.textContent = displayName;
    return;
  }

  label.textContent = `Current file: ${displayName}`;
}

setCurrentFileDisplay(window._currentFilePath);


// ================= MAP SETUP =================
// Create Leaflet map
const map = L.map("map", { preferCanvas: true }).setView([31.0, -99.0], 6);
// Shared Canvas renderer for high-performance drawing
const canvasRenderer = L.canvas({ padding: 0.5 });


// ===== BASE MAP LAYERS =====
const baseMaps = {
  streets: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 19
    }),
  freeStreets: L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 20,
      maxNativeZoom: 19,
      subdomains: ["a", "b", "c"],
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    }),

  satellite: L.tileLayer(
    "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
    {
      maxZoom: 20,
      maxNativeZoom: 19
    }
  )
};
// ===== SATELLITE STREET NAME OVERLAY (LIGHTWEIGHT) =====
const satelliteLabelsLayer = L.tileLayer(
  "https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Transportation/MapServer/tile/{z}/{y}/{x}",
  {
    maxZoom: 20,
    maxNativeZoom: 19,
    opacity: 1
  }
);
const streetAttributeLayerGroup = L.layerGroup();
const streetLoadPolygonLayerGroup = new L.FeatureGroup();
map.addLayer(streetLoadPolygonLayerGroup);
let streetPolygonLoadPending = false;
let lastStreetLoadPolygonSnapshot = null;
const LOCAL_STREET_GRID_SIZE_DEG = 0.025;
const LOCAL_STREET_SAVED_POLYGONS_KEY = "localStreetSavedPolygons";
const LOCAL_STREET_SAVED_POLYGONS_MAX = 60;
const LOCAL_STREET_SYMBOLOGY_KEY = "localStreetSymbology";
const STREET_BASE_LINE_COLOR = "#4ea2f5";
const STREET_BASE_LINE_WEIGHT = 3;
const STREET_BASE_LINE_OPACITY = 0.65;
const STREET_SYMBOLOGY_EMPTY_KEY = "__EMPTY__";
const STREET_NETWORK_MANAGER_TAB_KEY = "streetNetworkManagerTab";
const STREET_SYMBOLOGY_PALETTE = [
  "#4ea2f5",
  "#ef4444",
  "#22c55e",
  "#f59e0b",
  "#a855f7",
  "#06b6d4",
  "#84cc16",
  "#f97316",
  "#ec4899",
  "#14b8a6",
  "#eab308",
  "#3b82f6"
];
const localStreetSourceState = {
  loaded: false,
  sourceName: "",
  cellSizeDeg: LOCAL_STREET_GRID_SIZE_DEG,
  elementsById: new Map(),
  cellIndex: new Map(),
  chunkMode: false,
  chunkBounds: null,
  sourceDescriptor: null
};
let localStreetStatusResetTimer = null;
const TEXAS_STREETS_DOWNLOAD_URL = "https://download.geofabrik.de/north-america/us/texas-latest-free.shp.zip";
const LOCAL_STREET_ZIP_RX = /\.zip$/i;
const LOCAL_STREET_JSON_RX = /\.(geojson|json)$/i;
const LOCAL_STREET_ROAD_LAYER_RX = /(road|street|highway)/i;
const LOCAL_STREET_FCLASS_WHITELIST = new Set([
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
const LOCAL_STREET_SOURCE_META_KEY = "localStreetSourceMeta";
const LOCAL_STREET_HANDLE_DB_NAME = "tdsPakLocalStreetSource";
const LOCAL_STREET_HANDLE_STORE_NAME = "handles";
const LOCAL_STREET_HANDLE_PRIMARY_KEY = "primary";
const LOCAL_STREET_ZIP_WORKER_SCRIPT = "tds-pak-street-zip-worker.js?v=20260306-2";
const LOCAL_STREET_BROWSER_ZIP_LIMIT_MB = 140;
const LOCAL_STREET_JSON_PARSE_WARN_MB = 420;
const LOCAL_STREET_JSON_STREAM_THRESHOLD_MB = 260;
const LOCAL_STREET_STREAM_YIELD_FEATURE_STEP = 300;
const LOCAL_STREET_OFFLINE_CONVERTER_PACKAGE = "tds-streets-offline-converter-package.zip?v=20260306-8";
const LOCAL_STREET_AUTO_SETUP_PACKAGE = "tds-streets-auto-setup-package.zip?v=20260306-6";
const LOCAL_STREET_BACKEND_URL_KEY = "localStreetBackendUrl";
const STREET_NETWORK_LAYER_VISIBLE_KEY = "streetNetworkLayerVisible";
const LOCAL_STREET_BACKEND_URL_DEFAULT = "http://127.0.0.1:8787";
const LOCAL_STREET_BACKEND_HEALTH_TTL_MS = 12000;
const LOCAL_STREET_BACKEND_REQUEST_TIMEOUT_MS = 25000;
const LOCAL_STREET_BACKEND_QUERY_LIMIT = 220000;
const localStreetBackendState = {
  baseUrl: LOCAL_STREET_BACKEND_URL_DEFAULT,
  available: false,
  hasIndex: false,
  sourceName: "",
  lastError: "",
  checking: false,
  lastCheckedAt: 0
};
let streetSymbologyState = {
  enabled: false,
  field: "highway",
  valueColors: {},
  lineWidth: STREET_BASE_LINE_WEIGHT,
  opacity: STREET_BASE_LINE_OPACITY
};

function localStreetCellKey(latIdx, lonIdx) {
  return `${latIdx}:${lonIdx}`;
}

function normalizeLocalStreetBackendUrl(raw) {
  const text = String(raw || "").trim();
  if (!text) return LOCAL_STREET_BACKEND_URL_DEFAULT;
  const withProto = /^https?:\/\//i.test(text) ? text : `http://${text}`;
  return withProto.replace(/\/+$/, "");
}

function getStoredLocalStreetBackendUrl() {
  const stored = storageGet(LOCAL_STREET_BACKEND_URL_KEY);
  return normalizeLocalStreetBackendUrl(stored || LOCAL_STREET_BACKEND_URL_DEFAULT);
}

function setStoredLocalStreetBackendUrl(urlValue) {
  const normalized = normalizeLocalStreetBackendUrl(urlValue);
  storageSet(LOCAL_STREET_BACKEND_URL_KEY, normalized);
  localStreetBackendState.baseUrl = normalized;
}

function localStreetHasProvider() {
  return !!localStreetSourceState.loaded || (!!localStreetBackendState.available && !!localStreetBackendState.hasIndex);
}

function normalizeStreetSymbologyState(raw) {
  const base = {
    enabled: false,
    field: "highway",
    valueColors: {},
    lineWidth: STREET_BASE_LINE_WEIGHT,
    opacity: STREET_BASE_LINE_OPACITY
  };
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return base;

  const field = String(raw.field || "highway").trim() || "highway";
  const valueColors = {};
  if (raw.valueColors && typeof raw.valueColors === "object" && !Array.isArray(raw.valueColors)) {
    Object.keys(raw.valueColors).forEach(key => {
      const normalizedKey = String(key || "").trim();
      const color = String(raw.valueColors[key] || "").trim();
      if (!normalizedKey || !/^#[0-9a-fA-F]{6}$/.test(color)) return;
      valueColors[normalizedKey] = color.toLowerCase();
    });
  }

  const widthNum = Number(raw.lineWidth);
  const opacityNum = Number(raw.opacity);
  return {
    enabled: !!raw.enabled,
    field,
    valueColors,
    lineWidth: Number.isFinite(widthNum) ? Math.max(1, Math.min(8, widthNum)) : STREET_BASE_LINE_WEIGHT,
    opacity: Number.isFinite(opacityNum) ? Math.max(0.2, Math.min(1, opacityNum)) : STREET_BASE_LINE_OPACITY
  };
}

function loadStoredStreetSymbologyState() {
  const raw = storageGet(LOCAL_STREET_SYMBOLOGY_KEY);
  if (!raw) return normalizeStreetSymbologyState(null);
  try {
    return normalizeStreetSymbologyState(JSON.parse(raw));
  } catch {
    return normalizeStreetSymbologyState(null);
  }
}

function saveStoredStreetSymbologyState() {
  storageSet(LOCAL_STREET_SYMBOLOGY_KEY, JSON.stringify(normalizeStreetSymbologyState(streetSymbologyState)));
}

function ensureStreetSymbologyState() {
  streetSymbologyState = normalizeStreetSymbologyState(streetSymbologyState);
  return streetSymbologyState;
}

function normalizeStreetSymbologyValueKey(value) {
  const text = String(value ?? "").trim();
  return text ? text : STREET_SYMBOLOGY_EMPTY_KEY;
}

function formatStreetSymbologyValueLabel(key) {
  if (key === STREET_SYMBOLOGY_EMPTY_KEY) return "(Blank)";
  return key;
}

function getStreetSymbologyPaletteColorByKey(key) {
  const text = String(key || "");
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = ((hash << 5) - hash) + text.charCodeAt(i);
    hash |= 0;
  }
  const idx = Math.abs(hash) % STREET_SYMBOLOGY_PALETTE.length;
  return STREET_SYMBOLOGY_PALETTE[idx];
}

function getStreetSymbologyAvailableFields() {
  const preferredOrder = ["highway", "surface", "lanes", "maxspeed", "oneway", "name", "ref", "id"];
  const fieldSet = new Set(preferredOrder);
  streetAttributeById.forEach(entry => {
    const row = entry?.row;
    if (!row || typeof row !== "object") return;
    Object.keys(row).forEach(key => {
      const normalized = String(key || "").trim();
      if (!normalized) return;
      fieldSet.add(normalized);
    });
  });
  const fields = [...fieldSet];
  fields.sort((a, b) => {
    const aIndex = preferredOrder.indexOf(a);
    const bIndex = preferredOrder.indexOf(b);
    const aRank = aIndex === -1 ? 999 : aIndex;
    const bRank = bIndex === -1 ? 999 : bIndex;
    if (aRank !== bRank) return aRank - bRank;
    return a.localeCompare(b, undefined, { sensitivity: "base" });
  });
  return fields;
}

function getStreetSymbologyClassStats(fieldName) {
  const field = String(fieldName || "").trim();
  const counts = new Map();
  streetAttributeById.forEach(entry => {
    const row = entry?.row || {};
    const valueKey = normalizeStreetSymbologyValueKey(row?.[field]);
    const item = counts.get(valueKey) || { key: valueKey, label: formatStreetSymbologyValueLabel(valueKey), count: 0 };
    item.count += 1;
    counts.set(valueKey, item);
  });
  return [...counts.values()].sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.label.localeCompare(b.label, undefined, { sensitivity: "base" });
  });
}

function resolveStreetSymbologyColorForEntry(entry) {
  const state = ensureStreetSymbologyState();
  if (!state.enabled) return STREET_BASE_LINE_COLOR;
  const field = String(state.field || "highway");
  const valueKey = normalizeStreetSymbologyValueKey(entry?.row?.[field]);
  const explicit = state.valueColors?.[valueKey];
  if (typeof explicit === "string" && /^#[0-9a-fA-F]{6}$/.test(explicit)) {
    return explicit.toLowerCase();
  }
  return getStreetSymbologyPaletteColorByKey(valueKey);
}

function getStreetSegmentBaseStyle(entry) {
  const state = ensureStreetSymbologyState();
  return {
    color: resolveStreetSymbologyColorForEntry(entry),
    weight: state.enabled ? state.lineWidth : STREET_BASE_LINE_WEIGHT,
    opacity: state.enabled ? state.opacity : STREET_BASE_LINE_OPACITY
  };
}

async function fetchJsonWithTimeout(url, options = {}, timeoutMs = LOCAL_STREET_BACKEND_REQUEST_TIMEOUT_MS) {
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timeout = controller
    ? setTimeout(() => controller.abort(), Math.max(500, Number(timeoutMs) || LOCAL_STREET_BACKEND_REQUEST_TIMEOUT_MS))
    : null;
  try {
    const response = await fetch(url, {
      ...options,
      signal: controller ? controller.signal : undefined
    });
    return response;
  } finally {
    if (timeout) clearTimeout(timeout);
  }
}

async function checkLocalStreetBackendAvailability(force = false) {
  const now = Date.now();
  if (!force && !localStreetBackendState.checking && (now - localStreetBackendState.lastCheckedAt) < LOCAL_STREET_BACKEND_HEALTH_TTL_MS) {
    return localStreetBackendState.available;
  }
  if (localStreetBackendState.checking) return localStreetBackendState.available;

  localStreetBackendState.checking = true;
  localStreetBackendState.lastCheckedAt = now;
  try {
    const healthUrl = `${localStreetBackendState.baseUrl}/api/health`;
    const response = await fetchJsonWithTimeout(healthUrl, { cache: "no-store" }, 5000);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const payload = await response.json().catch(() => ({}));
    localStreetBackendState.available = true;
    localStreetBackendState.hasIndex = !!payload?.has_index;
    localStreetBackendState.sourceName = String(payload?.source_name || "");
    localStreetBackendState.lastError = localStreetBackendState.hasIndex
      ? ""
      : "Backend is running but no streets index is loaded.";
  } catch (err) {
    localStreetBackendState.available = false;
    localStreetBackendState.hasIndex = false;
    localStreetBackendState.sourceName = "";
    localStreetBackendState.lastError = String(err?.message || err || "Backend unavailable");
  } finally {
    localStreetBackendState.checking = false;
    updateLocalStreetSourceStatus();
  }
  return localStreetBackendState.available;
}

function resetLocalStreetSourceState(options = {}) {
  const preserveSourceDescriptor = !!options.preserveSourceDescriptor;
  localStreetSourceState.loaded = false;
  localStreetSourceState.sourceName = "";
  localStreetSourceState.elementsById.clear();
  localStreetSourceState.cellIndex.clear();
  localStreetSourceState.chunkMode = false;
  localStreetSourceState.chunkBounds = null;
  if (!preserveSourceDescriptor) {
    localStreetSourceState.sourceDescriptor = null;
  }
}

function shouldUseLocalStreetSource() {
  const toggle = document.getElementById("useLocalStreetSource");
  return !!toggle?.checked && localStreetHasProvider();
}

function isStreetNetworkLayerVisibleEnabled() {
  const toggle = document.getElementById("streetNetworkLayerToggle");
  if (!toggle) return true;
  return !!toggle.checked;
}

function setStreetNetworkManagerBadgeState(state = "off") {
  const badge = document.getElementById("streetNetworkManagerBadge");
  if (!badge) return;
  const states = {
    active: { text: "Active", className: "is-active" },
    hidden: { text: "Hidden", className: "is-hidden" },
    ready: { text: "Ready", className: "is-ready" },
    off: { text: "No Source", className: "is-off" },
    checking: { text: "Checking", className: "is-checking" }
  };
  const next = states[state] || states.off;
  badge.classList.remove("is-active", "is-hidden", "is-ready", "is-off", "is-checking");
  badge.classList.add(next.className);
  badge.textContent = next.text;
}

function resolveStreetNetworkManagerBadgeState(message = "", hasProvider = localStreetHasProvider()) {
  const messageText = String(message || "").toLowerCase();
  if (localStreetBackendState.checking || messageText.includes("checking")) return "checking";
  if (!hasProvider) return "off";
  if (shouldUseLocalStreetSource()) {
    return isStreetNetworkLayerVisibleEnabled() ? "active" : "hidden";
  }
  return "ready";
}

function updateStreetNetworkManagerHint(hasProvider = localStreetHasProvider()) {
  const hintNode = document.getElementById("streetNetworkManagerHint");
  if (!hintNode) return;

  const usingLocal = shouldUseLocalStreetSource();
  const layerVisible = isStreetNetworkLayerVisibleEnabled();

  if (!hasProvider) {
    hintNode.textContent = "Start with Street Setup Wizard: download setup program, run launcher, then check backend.";
    return;
  }

  if (localStreetBackendState.available && localStreetBackendState.hasIndex) {
    if (usingLocal && layerVisible) {
      hintNode.textContent = "Backend source is active. Choose a saved polygon or draw a new one to load streets for this area.";
    } else if (usingLocal) {
      hintNode.textContent = "Street source is on, but map layer is hidden. Turn on Street Network Layer in the sidebar.";
    } else {
      hintNode.textContent = "Backend is ready. Turn on Street Segments, then choose a saved polygon or draw a new one.";
    }
    return;
  }

  const loadedCount = localStreetSourceState.elementsById.size.toLocaleString();
  if (usingLocal && layerVisible) {
    hintNode.textContent = `Local source is active (${loadedCount} indexed). Choose a saved polygon or draw a new one to refresh streets in view.`;
  } else if (usingLocal) {
    hintNode.textContent = "Street source is on, but map layer is hidden. Turn on Street Network Layer in the sidebar.";
  } else {
    hintNode.textContent = `Local source is ready (${loadedCount} indexed). Turn on Street Segments, then choose a saved polygon or draw a new one.`;
  }
}

function updateLocalStreetSourceStatus(message = "") {
  const node = document.getElementById("localStreetsStatus");
  const useLocalToggle = document.getElementById("useLocalStreetSource");
  const hasProvider = localStreetHasProvider();
  updateStreetSetupGuide();
  if (useLocalToggle) {
    useLocalToggle.disabled = !hasProvider;
    if (!hasProvider) useLocalToggle.checked = false;
  }
  setStreetNetworkManagerBadgeState(resolveStreetNetworkManagerBadgeState(message, hasProvider));
  updateStreetNetworkManagerHint(hasProvider);
  if (!node) return;
  if (message) {
    node.textContent = message;
    return;
  }
  if (!hasProvider) {
    node.textContent = "Street layer: Off. Click Street Setup Wizard, then complete steps 1-3.";
    return;
  }

  if (localStreetBackendState.available && localStreetBackendState.hasIndex) {
    const backendName = localStreetBackendState.sourceName ? ` (${localStreetBackendState.sourceName})` : "";
    if (shouldUseLocalStreetSource() && isStreetNetworkLayerVisibleEnabled()) {
      node.textContent = `Street layer: On (Local backend${backendName})`;
    } else if (shouldUseLocalStreetSource()) {
      node.textContent = `Street layer: Hidden (Local backend${backendName})`;
    } else {
      node.textContent = `Street layer: Off (Local backend ready${backendName}). Turn on Street Segments, then choose a saved polygon or draw a new one.`;
    }
    return;
  }

  const usingLocal = shouldUseLocalStreetSource();
  const layerVisible = isStreetNetworkLayerVisibleEnabled();
  const count = localStreetSourceState.elementsById.size.toLocaleString();
  const chunkMode = !!localStreetSourceState.chunkMode;
  if (usingLocal && layerVisible) {
    node.textContent = chunkMode
      ? `Street layer: On (Chunk mode, ${count} segments indexed for current region from ${localStreetSourceState.sourceName})`
      : `Street layer: On (Local file: ${localStreetSourceState.sourceName}, ${count} segments indexed)`;
  } else if (usingLocal) {
    node.textContent = chunkMode
      ? `Street layer: Hidden (Chunk mode, ${count} segments indexed for current region from ${localStreetSourceState.sourceName})`
      : `Street layer: Hidden (Local file: ${localStreetSourceState.sourceName}, ${count} segments indexed)`;
  } else {
    node.textContent = chunkMode
      ? `Street layer: Off (Chunk mode ready: ${localStreetSourceState.sourceName}, ${count} region segments indexed)`
      : `Street layer: Off (Local file loaded: ${localStreetSourceState.sourceName}, ${count} segments indexed)`;
  }
}

function addLocalStreetElementToIndex(element) {
  const id = Number(element?.id);
  if (!Number.isFinite(id)) return false;
  const geom = Array.isArray(element?.geom) ? element.geom : [];
  if (geom.length < 2) return false;

  localStreetSourceState.elementsById.set(id, element);

  let minLat = Infinity;
  let maxLat = -Infinity;
  let minLon = Infinity;
  let maxLon = -Infinity;
  geom.forEach(p => {
    const lat = Number(p?.lat);
    const lon = Number(p?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    if (lat < minLat) minLat = lat;
    if (lat > maxLat) maxLat = lat;
    if (lon < minLon) minLon = lon;
    if (lon > maxLon) maxLon = lon;
  });
  if (!Number.isFinite(minLat) || !Number.isFinite(minLon)) return false;

  const s = localStreetSourceState.cellSizeDeg;
  const minLatIdx = Math.floor((minLat + 90) / s);
  const maxLatIdx = Math.floor((maxLat + 90) / s);
  const minLonIdx = Math.floor((minLon + 180) / s);
  const maxLonIdx = Math.floor((maxLon + 180) / s);

  for (let latIdx = minLatIdx; latIdx <= maxLatIdx; latIdx++) {
    for (let lonIdx = minLonIdx; lonIdx <= maxLonIdx; lonIdx++) {
      const key = localStreetCellKey(latIdx, lonIdx);
      let bucket = localStreetSourceState.cellIndex.get(key);
      if (!bucket) {
        bucket = [];
        localStreetSourceState.cellIndex.set(key, bucket);
      }
      bucket.push(id);
    }
  }
  return true;
}

function getLocalStreetCandidateIds(bounds) {
  const ids = new Set();
  if (!bounds || !localStreetSourceState.loaded) return ids;

  const south = bounds.getSouth();
  const west = bounds.getWest();
  const north = bounds.getNorth();
  const east = bounds.getEast();
  const s = localStreetSourceState.cellSizeDeg;
  const minLatIdx = Math.floor((south + 90) / s);
  const maxLatIdx = Math.floor((north + 90) / s);
  const minLonIdx = Math.floor((west + 180) / s);
  const maxLonIdx = Math.floor((east + 180) / s);

  for (let latIdx = minLatIdx; latIdx <= maxLatIdx; latIdx++) {
    for (let lonIdx = minLonIdx; lonIdx <= maxLonIdx; lonIdx++) {
      const key = localStreetCellKey(latIdx, lonIdx);
      const bucket = localStreetSourceState.cellIndex.get(key);
      if (!bucket?.length) continue;
      bucket.forEach(id => ids.add(id));
    }
  }
  return ids;
}

function getLocalStreetSourceMeta() {
  const raw = storageGet(LOCAL_STREET_SOURCE_META_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    return parsed;
  } catch {
    return null;
  }
}

function setLocalStreetSourceMeta(meta) {
  storageSet(LOCAL_STREET_SOURCE_META_KEY, JSON.stringify(meta || {}));
}

function clearLocalStreetSourceMeta() {
  localStorage.removeItem(storageKey(LOCAL_STREET_SOURCE_META_KEY));
}

function openLocalStreetHandleDb() {
  if (!window.indexedDB) return Promise.resolve(null);
  return new Promise(resolve => {
    const req = window.indexedDB.open(LOCAL_STREET_HANDLE_DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(LOCAL_STREET_HANDLE_STORE_NAME)) {
        db.createObjectStore(LOCAL_STREET_HANDLE_STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null);
  });
}

async function readStoredLocalStreetHandle() {
  const db = await openLocalStreetHandleDb();
  if (!db) return null;
  return new Promise(resolve => {
    const tx = db.transaction(LOCAL_STREET_HANDLE_STORE_NAME, "readonly");
    const store = tx.objectStore(LOCAL_STREET_HANDLE_STORE_NAME);
    const req = store.get(LOCAL_STREET_HANDLE_PRIMARY_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => resolve(null);
    tx.oncomplete = () => db.close();
    tx.onerror = () => db.close();
  });
}

async function writeStoredLocalStreetHandle(handle) {
  const db = await openLocalStreetHandleDb();
  if (!db) return false;
  return new Promise(resolve => {
    const tx = db.transaction(LOCAL_STREET_HANDLE_STORE_NAME, "readwrite");
    const store = tx.objectStore(LOCAL_STREET_HANDLE_STORE_NAME);
    store.put(handle, LOCAL_STREET_HANDLE_PRIMARY_KEY);
    tx.oncomplete = () => {
      db.close();
      resolve(true);
    };
    tx.onerror = () => {
      db.close();
      resolve(false);
    };
  });
}

async function clearStoredLocalStreetHandle() {
  const db = await openLocalStreetHandleDb();
  if (!db) return false;
  return new Promise(resolve => {
    const tx = db.transaction(LOCAL_STREET_HANDLE_STORE_NAME, "readwrite");
    const store = tx.objectStore(LOCAL_STREET_HANDLE_STORE_NAME);
    store.delete(LOCAL_STREET_HANDLE_PRIMARY_KEY);
    tx.oncomplete = () => {
      db.close();
      resolve(true);
    };
    tx.onerror = () => {
      db.close();
      resolve(false);
    };
  });
}

async function rememberLocalStreetSourceForAutoLoad({ sourceName = "", sourcePath = "", handle = null } = {}) {
  const cleanedPath = String(sourcePath || "").trim();
  if (!handle && !cleanedPath) {
    clearLocalStreetSourceMeta();
    await clearStoredLocalStreetHandle();
    return;
  }
  setLocalStreetSourceMeta({
    sourceName: String(sourceName || ""),
    sourcePath: cleanedPath,
    hasHandle: !!handle,
    updatedAt: new Date().toISOString()
  });
  if (handle) {
    await writeStoredLocalStreetHandle(handle);
  } else {
    await clearStoredLocalStreetHandle();
  }
}

async function forgetRememberedLocalStreetSource() {
  clearLocalStreetSourceMeta();
  await clearStoredLocalStreetHandle();
}

function isLikelyLocalFilesystemPath(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (/^[a-z]:[\\/]/i.test(raw)) return true;
  if (raw.startsWith("\\\\")) return true;
  if (raw.startsWith("/")) return true;
  return false;
}

function normalizeLocalStreetSourcePathToUrl(pathInput) {
  const raw = String(pathInput || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw) || /^file:\/\//i.test(raw)) return raw;
  if (/^[a-z]:[\\/]/i.test(raw)) {
    return `file:///${encodeURI(raw.replace(/\\/g, "/"))}`;
  }
  if (raw.startsWith("\\\\")) {
    return `file:${encodeURI(raw.replace(/\\/g, "/"))}`;
  }
  if (raw.startsWith("/")) {
    return `file://${encodeURI(raw)}`;
  }
  return raw;
}

function collectLocalStreetFeatureCollections(parsedSource) {
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

function isLocalStreetFeature(feature, layerName = "") {
  const geometryType = feature?.geometry?.type;
  if (geometryType !== "LineString" && geometryType !== "MultiLineString") return false;

  const props = feature?.properties || {};
  const highway = String(props.highway || props.HIGHWAY || "").trim();
  if (highway) return true;

  const fclass = String(props.fclass || props.FCLASS || "").trim().toLowerCase();
  if (fclass) return LOCAL_STREET_FCLASS_WHITELIST.has(fclass);

  if (LOCAL_STREET_ROAD_LAYER_RX.test(String(layerName || ""))) return true;
  return true;
}

function pickLocalStreetFeaturesFromParsedSource(parsedSource, preferRoadLayers = false) {
  const collections = collectLocalStreetFeatureCollections(parsedSource);
  if (!collections.length) return [];

  const roadCollections = collections.filter(c => LOCAL_STREET_ROAD_LAYER_RX.test(c.layerName));
  const sourceCollections = preferRoadLayers && roadCollections.length ? roadCollections : collections;
  const features = [];

  sourceCollections.forEach(({ layerName, features: layerFeatures }) => {
    layerFeatures.forEach(feature => {
      if (!isLocalStreetFeature(feature, layerName)) return;
      features.push(feature);
    });
  });

  return features;
}

function formatLocalStreetElapsed(ms) {
  const totalSeconds = Math.max(0, Math.floor(Number(ms || 0) / 1000));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  if (hours > 0) {
    return `${hours}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }
  return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

function ensureZipSizeWithinBrowserLimit(byteLength, sourceLabel = "ZIP file") {
  const bytes = Number(byteLength || 0);
  if (!Number.isFinite(bytes) || bytes <= 0) return;
  const maxBytes = LOCAL_STREET_BROWSER_ZIP_LIMIT_MB * 1024 * 1024;
  if (bytes <= maxBytes) return;

  const sizeMb = (bytes / (1024 * 1024)).toFixed(1);
  throw new Error(
    `${sourceLabel} is ${sizeMb} MB, which is above the in-browser conversion limit ` +
    `(${LOCAL_STREET_BROWSER_ZIP_LIMIT_MB} MB).\n\n` +
    "Large statewide ZIPs freeze browser tabs. Use one of these options:\n" +
    "1) Convert offline once with GDAL/mapshaper and load the resulting roads GeoJSON.\n" +
    "2) Create a roads-only ZIP (gis_osm_roads_free_1.*) and use ZIP -> JSON."
  );
}

function parseZipStreetFeaturesWithWorker(sourceZip, label = "streets ZIP") {
  if (typeof Worker !== "function") {
    return Promise.reject(new Error("Web Worker is not supported in this browser."));
  }

  return new Promise((resolve, reject) => {
    let worker;
    let pulseLabel = `Converting ${label} to GeoJSON...`;
    let pulsePercent = 12;
    let pulseCeiling = 69.5;
    let stageKind = "convert";
    const startedAt = Date.now();
    const maxRuntimeMs = 12 * 60 * 1000;
    let lastWorkerMessageAt = startedAt;
    showLocalStreetLoadPercent(pulsePercent, `${pulseLabel} (working 00:00)`);
    const pulseTimer = setInterval(() => {
      if (stageKind === "extract") {
        pulseCeiling = 89.5;
      } else {
        pulseCeiling = 69.5;
      }
      const silentForMs = Date.now() - lastWorkerMessageAt;
      if (stageKind === "convert" && silentForMs > 90000) {
        pulseCeiling = 85;
      }
      if (stageKind === "convert" && silentForMs > 240000) {
        pulseCeiling = 92;
      }
      if (pulsePercent < pulseCeiling) {
        const remaining = pulseCeiling - pulsePercent;
        const step = Math.max(0.15, remaining * 0.08);
        pulsePercent = Math.min(pulseCeiling, pulsePercent + step);
      }
      const elapsed = formatLocalStreetElapsed(Date.now() - startedAt);
      const stale = (Date.now() - lastWorkerMessageAt) > 30000;
      const suffix = stale ? ` (still working ${elapsed})` : ` (working ${elapsed})`;
      showLocalStreetLoadPercent(pulsePercent, `${pulseLabel}${suffix}`);

      if ((Date.now() - startedAt) > maxRuntimeMs) {
        const timeoutMinutes = Math.round(maxRuntimeMs / 60000);
        cleanup();
        reject(new Error(
          `Street ZIP conversion timed out after ${timeoutMinutes} minutes. ` +
          "Use a roads-only ZIP (gis_osm_roads_free_1.*) for reliable loading."
        ));
      }
    }, 1000);
    try {
      worker = new Worker(LOCAL_STREET_ZIP_WORKER_SCRIPT);
    } catch (err) {
      clearInterval(pulseTimer);
      reject(err);
      return;
    }

    const cleanup = () => {
      clearInterval(pulseTimer);
      if (worker) {
        worker.terminate();
      }
    };

    worker.onmessage = (event) => {
      const data = event?.data || {};
      if (data.type === "progress") {
        lastWorkerMessageAt = Date.now();
        const reported = Number(data.percent);
        const stageText = String(data.stage || pulseLabel);
        pulseLabel = stageText;
        if (stageText.toLowerCase().includes("extract")) {
          stageKind = "extract";
          pulseCeiling = 89.5;
        } else if (stageText.toLowerCase().includes("convert")) {
          stageKind = "convert";
          pulseCeiling = 69.5;
        }
        if (Number.isFinite(reported)) {
          pulsePercent = Math.max(pulsePercent, Math.min(reported, 89.9));
        }
        const elapsed = formatLocalStreetElapsed(Date.now() - startedAt);
        showLocalStreetLoadPercent(
          Number.isFinite(reported) ? Math.max(reported, pulsePercent) : pulsePercent,
          `${pulseLabel} (working ${elapsed})`
        );
        return;
      }
      if (data.type === "result") {
        const features = Array.isArray(data.features) ? data.features : [];
        cleanup();
        resolve(features);
        return;
      }
      if (data.type === "error") {
        cleanup();
        reject(new Error(data.message || "ZIP worker conversion failed."));
      }
    };

    worker.onerror = (err) => {
      cleanup();
      reject(new Error(err?.message || "ZIP worker conversion failed."));
    };

    worker.postMessage({
      type: "parseZip",
      buffer: sourceZip,
      label: String(label || "streets ZIP")
    });
  });
}

async function extractStreetFeaturesFromZipBuffer(sourceZip, sourceLabel) {
  const label = String(sourceLabel || "streets ZIP");
  ensureZipSizeWithinBrowserLimit(sourceZip?.byteLength, label);
  showLocalStreetLoadPercent(10, `Preparing ${label} for conversion...`);
  await sleep(20);

  return parseZipStreetFeaturesWithWorker(sourceZip, label);
}

function makeGeoJsonDownloadName(sourceName = "") {
  const raw = String(sourceName || "local-streets");
  const noPath = raw.split(/[\\/]/).pop() || raw;
  const noExt = noPath.replace(/\.[^.]+$/, "");
  const safe = noExt
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  return `${safe || "local-streets"}-roads.geojson`;
}

function triggerGeoJsonDownloadFromFeatures(features, sourceName = "") {
  if (!Array.isArray(features) || !features.length) return false;
  const collection = {
    type: "FeatureCollection",
    features
  };
  const blob = new Blob([JSON.stringify(collection)], { type: "application/geo+json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = makeGeoJsonDownloadName(sourceName);
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return true;
}

async function convertZipToGeoJsonAndDownload(file) {
  if (!file) return false;
  const fileName = file.name || "streets.zip";
  if (!LOCAL_STREET_ZIP_RX.test(fileName)) {
    alert("Select a .zip streets file.");
    return false;
  }
  if (typeof file.arrayBuffer !== "function") {
    alert("Unable to read selected ZIP file.");
    return false;
  }

  showLocalStreetLoadPercent(4, `Reading ${fileName}...`);
  await sleep(20);

  let features = [];
  try {
    const sourceZip = await file.arrayBuffer();
    ensureZipSizeWithinBrowserLimit(sourceZip?.byteLength, fileName);
    features = await extractStreetFeaturesFromZipBuffer(sourceZip, fileName);
  } catch (err) {
    finishLocalStreetLoadProgress("Could not convert ZIP to GeoJSON.", true);
    alert(`Could not convert ZIP to GeoJSON.\n\n${err?.message || err}`);
    return false;
  }

  if (!Array.isArray(features) || !features.length) {
    finishLocalStreetLoadProgress("No street features found in ZIP.", true);
    alert("No street features found in ZIP.");
    return false;
  }

  showLocalStreetLoadPercent(95, `Creating GeoJSON download (${features.length.toLocaleString()} roads)...`);
  await sleep(0);
  const ok = triggerGeoJsonDownloadFromFeatures(features, fileName);
  if (!ok) {
    finishLocalStreetLoadProgress("GeoJSON download failed.", true);
    alert("GeoJSON download failed.");
    return false;
  }

  finishLocalStreetLoadProgress(
    `100% - Downloaded ${features.length.toLocaleString()} roads as GeoJSON`,
    false
  );
  return true;
}

function normalizeLocalStreetTags(props = {}) {
  const normalizeTagValue = (value, fallback = "Unknown") => {
    if (value === null || value === undefined) return fallback;
    const text = String(value).trim();
    return text ? text : fallback;
  };

  return {
    name: normalizeTagValue(props.name ?? props.NAME),
    highway: normalizeTagValue(props.highway ?? props.HIGHWAY ?? props.road_class ?? props.fclass ?? props.FCLASS),
    ref: normalizeTagValue(props.ref ?? props.REF ?? props.ref_name ?? props.REF_NAME),
    maxspeed: normalizeTagValue(props.maxspeed ?? props.MAXSPEED ?? props.max_speed ?? props.MAX_SPEED),
    lanes: normalizeTagValue(props.lanes ?? props.LANES ?? props.num_lanes ?? props.NUM_LANES),
    surface: normalizeTagValue(props.surface ?? props.SURFACE ?? props.surf_type ?? props.SURF_TYPE),
    oneway: normalizeTagValue(props.oneway ?? props.ONEWAY ?? props.one_way ?? props.ONE_WAY)
  };
}

function isUnknownStreetTagValue(value) {
  const text = String(value ?? "").trim();
  return !text || text.toLowerCase() === "unknown";
}

function mergeStreetAttributeRows(existingRow, incomingRow) {
  if (!existingRow) return incomingRow;
  const merged = { ...incomingRow };
  ["name", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"].forEach(key => {
    const nextVal = incomingRow?.[key];
    if (isUnknownStreetTagValue(nextVal) && !isUnknownStreetTagValue(existingRow?.[key])) {
      merged[key] = existingRow[key];
    }
  });
  return merged;
}

function rowHasKnownStreetAttributes(row) {
  return ["name", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"]
    .some(key => !isUnknownStreetTagValue(row?.[key]));
}

function normalizeLocalLineCoords(coords) {
  if (!Array.isArray(coords)) return [];
  const geom = [];
  coords.forEach(pt => {
    if (!Array.isArray(pt) || pt.length < 2) return;
    const lon = Number(pt[0]);
    const lat = Number(pt[1]);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    geom.push({ lat, lon });
  });
  return geom.length >= 2 ? geom : [];
}

function sanitizeLocalStreetJsonText(raw) {
  if (typeof raw !== "string") return "";
  // Remove UTF-8 BOM and stray trailing null bytes from external converters.
  return raw.replace(/^\uFEFF/, "").replace(/\u0000+$/g, "");
}

function tryRepairTruncatedFeatureCollection(raw) {
  if (typeof raw !== "string") return null;
  const trimmed = raw.trimEnd();
  if (!trimmed) return null;
  if (!/^\s*\{/.test(trimmed)) return null;
  if (!/\"type\"\s*:\s*\"FeatureCollection\"/i.test(trimmed)) return null;
  if (!/\"features\"\s*:\s*\[/i.test(trimmed)) return null;
  if (/\]\}\s*$/.test(trimmed)) return null;

  const repaired = `${trimmed.replace(/,\s*$/, "")}]}`;
  try {
    return JSON.parse(repaired);
  } catch {
    return null;
  }
}

function parseLocalStreetJsonPayload(raw, sourceLabel = "street file", sourceSizeBytes = 0) {
  const cleaned = sanitizeLocalStreetJsonText(raw);
  try {
    return JSON.parse(cleaned);
  } catch (err) {
    const message = String(err?.message || err || "Invalid JSON");
    const repaired = /Unexpected end of JSON input/i.test(message)
      ? tryRepairTruncatedFeatureCollection(cleaned)
      : null;
    if (repaired) {
      console.warn(`Recovered truncated FeatureCollection while loading ${sourceLabel}.`);
      return repaired;
    }

    const sizeMb = Number.isFinite(sourceSizeBytes) ? (sourceSizeBytes / (1024 * 1024)) : 0;
    const sizeHint = sizeMb >= LOCAL_STREET_JSON_PARSE_WARN_MB
      ? `\n\nFile size is ${sizeMb.toFixed(1)} MB. Very large GeoJSON files are fragile in browsers.`
      : "";
    const truncHint = /Unexpected end of JSON input/i.test(message)
      ? "\n\nThis file appears incomplete. Re-run the offline converter and use the new output file."
      : "";
    throw new Error(`${message}${truncHint}${sizeHint}`);
  }
}

function isValidLeafletBounds(bounds) {
  return !!bounds &&
    typeof bounds.getSouth === "function" &&
    typeof bounds.getWest === "function" &&
    typeof bounds.getNorth === "function" &&
    typeof bounds.getEast === "function";
}

function cloneLocalStreetBounds(bounds) {
  if (!isValidLeafletBounds(bounds)) return null;
  return L.latLngBounds(
    [bounds.getSouth(), bounds.getWest()],
    [bounds.getNorth(), bounds.getEast()]
  );
}

function boundsContainsBounds(outerBounds, innerBounds) {
  if (!isValidLeafletBounds(outerBounds) || !isValidLeafletBounds(innerBounds)) return false;
  return (
    innerBounds.getSouth() >= outerBounds.getSouth() &&
    innerBounds.getNorth() <= outerBounds.getNorth() &&
    innerBounds.getWest() >= outerBounds.getWest() &&
    innerBounds.getEast() <= outerBounds.getEast()
  );
}

function buildLocalStreetChunkBounds(targetBounds) {
  const base = isValidLeafletBounds(targetBounds)
    ? targetBounds
    : (map?.getBounds?.() || null);
  if (!isValidLeafletBounds(base)) return null;

  const south = base.getSouth();
  const west = base.getWest();
  const north = base.getNorth();
  const east = base.getEast();
  const spanLat = Math.max(0.0001, Math.abs(north - south));
  const spanLng = Math.max(0.0001, Math.abs(east - west));
  const padLat = Math.min(0.35, Math.max(0.015, spanLat * 0.2));
  const padLng = Math.min(0.35, Math.max(0.015, spanLng * 0.2));

  const chunkSouth = Math.max(-89.999999, south - padLat);
  const chunkWest = Math.max(-179.999999, west - padLng);
  const chunkNorth = Math.min(89.999999, north + padLat);
  const chunkEast = Math.min(179.999999, east + padLng);
  return L.latLngBounds([chunkSouth, chunkWest], [chunkNorth, chunkEast]);
}

function localStreetFeatureIntersectsBounds(feature, bounds) {
  if (!isValidLeafletBounds(bounds)) return true;

  const minTargetLat = bounds.getSouth();
  const maxTargetLat = bounds.getNorth();
  const minTargetLon = bounds.getWest();
  const maxTargetLon = bounds.getEast();

  let minLat = Infinity;
  let maxLat = -Infinity;
  let minLon = Infinity;
  let maxLon = -Infinity;

  const scanCoord = (coord) => {
    if (!Array.isArray(coord) || coord.length < 2) return;
    const lon = Number(coord[0]);
    const lat = Number(coord[1]);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    if (lat < minLat) minLat = lat;
    if (lat > maxLat) maxLat = lat;
    if (lon < minLon) minLon = lon;
    if (lon > maxLon) maxLon = lon;
  };

  const geometry = feature?.geometry || {};
  if (geometry.type === "LineString" && Array.isArray(geometry.coordinates)) {
    geometry.coordinates.forEach(scanCoord);
  } else if (geometry.type === "MultiLineString" && Array.isArray(geometry.coordinates)) {
    geometry.coordinates.forEach(line => {
      if (!Array.isArray(line)) return;
      line.forEach(scanCoord);
    });
  } else {
    return false;
  }

  if (!Number.isFinite(minLat) || !Number.isFinite(minLon)) return false;
  return !(
    maxLat < minTargetLat ||
    minLat > maxTargetLat ||
    maxLon < minTargetLon ||
    minLon > maxTargetLon
  );
}

function importLocalStreetFeatureIntoState(feature, nextIdState, filterBounds = null) {
  if (!feature || typeof feature !== "object") return 0;
  if (filterBounds && !localStreetFeatureIntersectsBounds(feature, filterBounds)) return 0;
  const geometry = feature.geometry || {};
  const props = feature.properties || {};
  const lineSets = geometry.type === "LineString"
    ? [geometry.coordinates]
    : (geometry.type === "MultiLineString" ? geometry.coordinates : []);
  if (!lineSets?.length) return 0;

  let imported = 0;
  const idState = nextIdState || { value: 1 };
  if (!Number.isFinite(idState.value) || idState.value < 1) idState.value = 1;
  const baseRawId = feature.id ?? props.id ?? props.osm_id ?? props.osm_way_id ?? props.way_id ?? null;

  lineSets.forEach((coords, partIdx) => {
    const geom = normalizeLocalLineCoords(coords);
    if (geom.length < 2) return;

    let candidateId = Number(baseRawId);
    if (!Number.isFinite(candidateId)) {
      candidateId = idState.value++;
    } else if (partIdx > 0) {
      const composite = Number(`${candidateId}${partIdx}`);
      candidateId = Number.isFinite(composite) ? composite : idState.value++;
    }

    while (localStreetSourceState.elementsById.has(candidateId)) {
      candidateId = idState.value++;
    }

    const element = {
      type: "way",
      id: candidateId,
      tags: normalizeLocalStreetTags(props),
      geom
    };

    if (addLocalStreetElementToIndex(element)) {
      imported += 1;
    }
  });

  return imported;
}

async function parseFeatureCollectionStreamFromReader(reader, options = {}) {
  const sourceLabel = String(options.sourceLabel || "GeoJSON stream");
  const totalBytes = Number(options.totalBytes || 0);
  const onFeature = typeof options.onFeature === "function" ? options.onFeature : null;
  const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;
  const featureYieldStep = Math.max(
    25,
    Number(options.featureYieldStep || LOCAL_STREET_STREAM_YIELD_FEATURE_STEP)
  );
  const decoder = new TextDecoder("utf-8");
  const headerSearchLimit = 1024 * 1024;
  const progressByteStep = 4 * 1024 * 1024;

  let buffer = "";
  let bytesRead = 0;
  let featuresParsed = 0;
  let featuresStarted = false;
  let arrayClosed = false;
  let objectDepth = 0;
  let featureStart = -1;
  let inString = false;
  let escapeNext = false;
  let lastProgressBytes = 0;

  const emitProgress = (force = false) => {
    if (!onProgress) return;
    if (!force && (bytesRead - lastProgressBytes) < progressByteStep) return;
    lastProgressBytes = bytesRead;
    onProgress({
      bytesRead,
      totalBytes,
      featuresParsed
    });
  };

  const emitFeature = async (featureJson) => {
    let parsedFeature;
    try {
      parsedFeature = JSON.parse(featureJson);
    } catch (err) {
      throw new Error(`Invalid feature JSON while streaming ${sourceLabel}: ${err?.message || err}`);
    }
    featuresParsed += 1;
    if (onFeature) {
      // Allow async feature handlers so indexing can yield without locking the page.
      // eslint-disable-next-line no-await-in-loop
      await onFeature(parsedFeature, featuresParsed);
    }
    if ((featuresParsed % featureYieldStep) === 0) {
      emitProgress(false);
      // eslint-disable-next-line no-await-in-loop
      await sleep(0);
    }
  };

  const parseBuffer = async (isFinalChunk) => {
    if (!featuresStarted) {
      const match = /"features"\s*:\s*\[/i.exec(buffer);
      if (!match) {
        if (buffer.length > headerSearchLimit) {
          throw new Error(
            `Could not find a FeatureCollection "features" array in ${sourceLabel}.`
          );
        }
        if (isFinalChunk) {
          throw new Error(`Could not find a FeatureCollection "features" array in ${sourceLabel}.`);
        }
        return;
      }
      const afterMatchIdx = match.index + match[0].length;
      buffer = buffer.slice(afterMatchIdx);
      featuresStarted = true;
    }

    let i = 0;
    while (i < buffer.length) {
      if (arrayClosed) {
        buffer = "";
        return;
      }

      if (objectDepth === 0) {
        while (i < buffer.length) {
          const ch = buffer[i];
          if (ch === "," || ch === " " || ch === "\n" || ch === "\r" || ch === "\t") {
            i += 1;
            continue;
          }
          break;
        }

        if (i >= buffer.length) break;

        const ch = buffer[i];
        if (ch === "]") {
          arrayClosed = true;
          i += 1;
          continue;
        }
        if (ch !== "{") {
          if (isFinalChunk) {
            throw new Error(`Unexpected token "${ch}" while reading features in ${sourceLabel}.`);
          }
          i += 1;
          continue;
        }
        featureStart = i;
        objectDepth = 1;
        inString = false;
        escapeNext = false;
        i += 1;
        continue;
      }

      const ch = buffer[i];
      if (inString) {
        if (escapeNext) {
          escapeNext = false;
        } else if (ch === "\\") {
          escapeNext = true;
        } else if (ch === "\"") {
          inString = false;
        }
      } else if (ch === "\"") {
        inString = true;
      } else if (ch === "{") {
        objectDepth += 1;
      } else if (ch === "}") {
        objectDepth -= 1;
        if (objectDepth === 0) {
          const featureJson = buffer.slice(featureStart, i + 1);
          // eslint-disable-next-line no-await-in-loop
          await emitFeature(featureJson);
          buffer = buffer.slice(i + 1);
          i = 0;
          featureStart = -1;
          continue;
        }
      }
      i += 1;
    }

    if (objectDepth === 0) {
      if (i > 0) {
        buffer = buffer.slice(i);
      }
      if (!arrayClosed && buffer.length > 65536) {
        buffer = buffer.slice(-65536);
      }
    } else if (featureStart > 0) {
      buffer = buffer.slice(featureStart);
      featureStart = 0;
    }

    if (isFinalChunk && !arrayClosed) {
      throw new Error(`Unexpected end of JSON input while reading ${sourceLabel}.`);
    }
  };

  try {
    while (true) {
      // eslint-disable-next-line no-await-in-loop
      const { value, done } = await reader.read();
      if (done) break;
      if (!value || !value.byteLength) continue;
      bytesRead += value.byteLength;
      buffer += decoder.decode(value, { stream: true });
      // eslint-disable-next-line no-await-in-loop
      await parseBuffer(false);
      emitProgress(false);
    }
    buffer += decoder.decode();
    await parseBuffer(true);
    emitProgress(true);
  } finally {
    if (typeof reader.releaseLock === "function") {
      try {
        reader.releaseLock();
      } catch {}
    }
  }

  if (!featuresStarted) {
    throw new Error(`Could not find a FeatureCollection "features" array in ${sourceLabel}.`);
  }
  if (!arrayClosed) {
    throw new Error(`Unexpected end of JSON input while reading ${sourceLabel}.`);
  }

  return { featuresParsed, bytesRead };
}

async function importLocalStreetFeaturesFromGeoJsonStream(reader, sourceName, options = {}) {
  const useLocalToggle = document.getElementById("useLocalStreetSource");
  const persistHandle = options.persistHandle || null;
  let persistPath = String(options.persistPath || "").trim();
  const skipRemember = !!options.skipRemember;
  const quiet = !!options.quiet;
  const filterBounds = isValidLeafletBounds(options.filterBounds) ? options.filterBounds : null;
  const chunkMode = !!filterBounds;
  const sourceSizeBytes = Number(options.sourceSizeBytes || 0);
  const preserveSourceDescriptor = !!options.preserveSourceDescriptor;
  const sourceDescriptor = options.sourceDescriptor || null;
  const preserveToggleState = !!options.preserveToggleState;

  resetLocalStreetSourceState({ preserveSourceDescriptor });
  if (sourceDescriptor) {
    localStreetSourceState.sourceDescriptor = sourceDescriptor;
  }
  const nextIdState = { value: 1 };
  let imported = 0;
  let lastProgressTs = 0;

  showLocalStreetLoadPercent(8, `Streaming ${sourceName}...`);
  await sleep(20);

  await parseFeatureCollectionStreamFromReader(reader, {
    sourceLabel: sourceName,
    totalBytes: sourceSizeBytes,
    onFeature: async (feature) => {
      if (!isLocalStreetFeature(feature, "")) return;
      imported += importLocalStreetFeatureIntoState(feature, nextIdState, filterBounds);
      if ((imported % LOCAL_STREET_STREAM_YIELD_FEATURE_STEP) === 0) {
        // eslint-disable-next-line no-await-in-loop
        await sleep(0);
      }
    },
    onProgress: ({ bytesRead, totalBytes, featuresParsed }) => {
      const now = Date.now();
      if ((now - lastProgressTs) < 250) return;
      lastProgressTs = now;

      const ratio = (Number.isFinite(totalBytes) && totalBytes > 0)
        ? Math.max(0, Math.min(1, bytesRead / totalBytes))
        : 0;
      const progress = 8 + (ratio * 90);
      showLocalStreetLoadPercent(
        progress,
        `Streaming ${sourceName}... ${featuresParsed.toLocaleString()} scanned, ${imported.toLocaleString()} indexed`
      );
    }
  });

  if (!imported) {
    if (chunkMode) {
      localStreetSourceState.loaded = true;
      localStreetSourceState.sourceName = sourceName;
      localStreetSourceState.chunkMode = true;
      localStreetSourceState.chunkBounds = cloneLocalStreetBounds(filterBounds);
      if (useLocalToggle && !preserveToggleState) useLocalToggle.checked = false;
      updateLocalStreetSourceStatus();
      finishLocalStreetLoadProgress("Chunk loaded, but no streets were found in this region.", true);
      return true;
    }
    finishLocalStreetLoadProgress("Street file did not contain valid street lines.", true);
    if (!quiet) alert("No valid street line features were found in this file.");
    updateLocalStreetSourceStatus();
    return false;
  }

  localStreetSourceState.loaded = true;
  localStreetSourceState.sourceName = sourceName;
  localStreetSourceState.chunkMode = chunkMode;
  localStreetSourceState.chunkBounds = chunkMode ? cloneLocalStreetBounds(filterBounds) : null;

  if (useLocalToggle && !preserveToggleState) useLocalToggle.checked = false;
  updateLocalStreetSourceStatus();

  if (!skipRemember && !persistHandle && !persistPath && !quiet) {
    const pastedPath = window.prompt(
      "To auto-open this streets file next time, paste its file path or URL (optional):",
      ""
    );
    persistPath = String(pastedPath || "").trim();
  }

  if (localStreetSourceState.sourceDescriptor && typeof localStreetSourceState.sourceDescriptor === "object") {
    localStreetSourceState.sourceDescriptor.sourceName = sourceName;
    localStreetSourceState.sourceDescriptor.sourcePath = persistPath;
    if (persistHandle) {
      localStreetSourceState.sourceDescriptor.sourceHandle = persistHandle;
    }
  }

  if (!skipRemember) {
    try {
      await rememberLocalStreetSourceForAutoLoad({
        sourceName,
        sourcePath: persistPath,
        handle: persistHandle
      });
    } catch (err) {
      console.warn("Unable to persist local streets source for auto-load:", err);
    }
  }

  showLocalStreetLoadPercent(100, `Loaded ${imported.toLocaleString()} local street segments`);
  finishLocalStreetLoadProgress(`100% - Loaded ${imported.toLocaleString()} local street segments`, false);
  return true;
}

function startTexasStreetsDownload(skipConfirm = false) {
  if (!skipConfirm) {
    const proceed = window.confirm(
      "Download the Geofabrik Texas streets source ZIP now?\n\nhttps://download.geofabrik.de/north-america/us/texas-latest-free.shp.zip"
    );
    if (!proceed) return false;
  }

  const link = document.createElement("a");
  link.href = TEXAS_STREETS_DOWNLOAD_URL;
  link.target = "_blank";
  link.rel = "noopener";
  link.download = "texas-latest-free.shp.zip";
  document.body.appendChild(link);
  link.click();
  link.remove();
  updateLocalStreetSourceStatus("Downloading Texas source ZIP from Geofabrik... then run the setup program on that ZIP.");
  return true;
}

function downloadOfflineStreetConverter() {
  const link = document.createElement("a");
  link.href = LOCAL_STREET_OFFLINE_CONVERTER_PACKAGE;
  link.download = "tds-streets-offline-converter-package.zip";
  document.body.appendChild(link);
  link.click();
  link.remove();
  updateLocalStreetSourceStatus("Downloaded converter package. Extract it, then run tds-streets-offline-converter-launcher.cmd.");
}

function downloadLocalStreetAutoSetupPackage() {
  const link = document.createElement("a");
  link.href = LOCAL_STREET_AUTO_SETUP_PACKAGE;
  link.download = "tds-streets-auto-setup-package.zip";
  document.body.appendChild(link);
  link.click();
  link.remove();
  updateLocalStreetSourceStatus(
    "Downloaded setup program package. Extract it, then run tds-local-streets-auto-setup-launcher.cmd."
  );
}

function showLocalStreetLoadProgress(message) {
  if (localStreetStatusResetTimer) {
    clearTimeout(localStreetStatusResetTimer);
    localStreetStatusResetTimer = null;
  }
  if (typeof window.hideLoading === "function") {
    window.hideLoading();
  }
  setStreetLoadBarVisible(true);
  updateLocalStreetSourceStatus(message || "Loading local streets...");
}

function finishLocalStreetLoadProgress(message = "", isError = false) {
  if (localStreetStatusResetTimer) {
    clearTimeout(localStreetStatusResetTimer);
    localStreetStatusResetTimer = null;
  }
  setStreetLoadBarVisible(false);
  if (!message) {
    updateLocalStreetSourceStatus();
    return;
  }
  updateLocalStreetSourceStatus(message);
  localStreetStatusResetTimer = setTimeout(() => {
    updateLocalStreetSourceStatus();
  }, isError ? 4200 : 2600);
}

function clampLocalStreetPercent(value) {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(100, Math.round(value * 10) / 10));
}

function showLocalStreetLoadPercent(percent, label) {
  const pct = clampLocalStreetPercent(percent);
  const pctText = Number.isInteger(pct) ? String(pct) : pct.toFixed(1);
  const text = String(label || "Loading local streets...");
  showLocalStreetLoadProgress(`${pctText}% - ${text}`);
}

function startLocalStreetProgressPulse(startPercent, endPercent, label) {
  let current = Number.isFinite(startPercent) ? startPercent : 0;
  const cap = Number.isFinite(endPercent) ? endPercent : 95;
  showLocalStreetLoadPercent(current, label);
  const timer = setInterval(() => {
    current = Math.min(cap, current + 1);
    showLocalStreetLoadPercent(current, label);
  }, 320);
  return () => {
    clearInterval(timer);
  };
}

async function importLocalStreetFeatures(features, sourceName, options = {}) {
  const useLocalToggle = document.getElementById("useLocalStreetSource");
  const persistHandle = options.persistHandle || null;
  const persistPath = String(options.persistPath || "").trim();
  const skipRemember = !!options.skipRemember;
  const filterBounds = isValidLeafletBounds(options.filterBounds) ? options.filterBounds : null;
  const chunkMode = !!filterBounds;
  const preserveSourceDescriptor = !!options.preserveSourceDescriptor;
  const sourceDescriptor = options.sourceDescriptor || null;
  const preserveToggleState = !!options.preserveToggleState;
  const progressStartPct = Number.isFinite(options.progressStartPct) ? options.progressStartPct : 70;
  const progressEndPct = Number.isFinite(options.progressEndPct) ? options.progressEndPct : 99;
  const progressLabel = String(options.progressLabel || "Indexing local streets...");

  resetLocalStreetSourceState({ preserveSourceDescriptor });
  if (sourceDescriptor) {
    localStreetSourceState.sourceDescriptor = sourceDescriptor;
  }
  const nextIdState = { value: 1 };
  let imported = 0;
  const reportStep = Math.max(250, Math.floor(features.length / 80));

  for (let i = 0; i < features.length; i++) {
    const feature = features[i];
    imported += importLocalStreetFeatureIntoState(feature, nextIdState, filterBounds);

    if (i % reportStep === 0) {
      const ratio = features.length > 0 ? (i / features.length) : 0;
      const progress = progressStartPct + ((progressEndPct - progressStartPct) * ratio);
      showLocalStreetLoadPercent(
        progress,
        `${progressLabel} ${i.toLocaleString()}/${features.length.toLocaleString()}`
      );
      // Yield to keep UI responsive while indexing large files.
      // eslint-disable-next-line no-await-in-loop
      await sleep(0);
    }
  }

  localStreetSourceState.loaded = imported > 0 || chunkMode;
  localStreetSourceState.sourceName = sourceName;
  localStreetSourceState.chunkMode = chunkMode;
  localStreetSourceState.chunkBounds = chunkMode ? cloneLocalStreetBounds(filterBounds) : null;

  if (!localStreetSourceState.loaded) {
    finishLocalStreetLoadProgress("Street file did not contain valid street lines.", true);
    alert("No valid street line features were found in this file.");
    updateLocalStreetSourceStatus();
    return false;
  }

  if (chunkMode && imported === 0) {
    if (useLocalToggle && !preserveToggleState) useLocalToggle.checked = false;
    updateLocalStreetSourceStatus();
    finishLocalStreetLoadProgress("Chunk loaded, but no streets were found in this region.", true);
    return true;
  }

  if (useLocalToggle && !preserveToggleState) useLocalToggle.checked = false;
  updateLocalStreetSourceStatus();

  if (localStreetSourceState.sourceDescriptor && typeof localStreetSourceState.sourceDescriptor === "object") {
    localStreetSourceState.sourceDescriptor.sourceName = sourceName;
    localStreetSourceState.sourceDescriptor.sourcePath = persistPath;
    if (persistHandle) {
      localStreetSourceState.sourceDescriptor.sourceHandle = persistHandle;
    }
  }

  if (!skipRemember) {
    try {
      await rememberLocalStreetSourceForAutoLoad({
        sourceName,
        sourcePath: persistPath,
        handle: persistHandle
      });
    } catch (err) {
      console.warn("Unable to persist local streets source for auto-load:", err);
    }
  }

  showLocalStreetLoadPercent(100, `Loaded ${imported.toLocaleString()} local street segments`);
  finishLocalStreetLoadProgress(`100% - Loaded ${imported.toLocaleString()} local street segments`, false);
  return true;
}

async function loadLocalStreetSourceFile(file, options = {}) {
  if (!file) return false;

  const fileName = file.name || "local streets file";
  const fileSizeBytes = Number(file.size || 0);
  const streamThresholdBytes = LOCAL_STREET_JSON_STREAM_THRESHOLD_MB * 1024 * 1024;
  const persistHandle = options.persistHandle || null;
  const persistPathValue = String(options.persistPath || "").trim();
  const chunkFilterBounds = isValidLeafletBounds(options.forceChunkBounds)
    ? cloneLocalStreetBounds(options.forceChunkBounds)
    : (fileSizeBytes >= streamThresholdBytes ? buildLocalStreetChunkBounds(map?.getBounds?.()) : null);
  const sourceDescriptor = {
    sourceName: fileName,
    sourceSizeBytes: fileSizeBytes,
    sourcePath: persistPathValue,
    sourceHandle: persistHandle || null,
    sourceFile: file,
    isChunkCapable: true
  };
  const lowerFileName = fileName.toLowerCase();
  const isZipSource = LOCAL_STREET_ZIP_RX.test(lowerFileName);
  const isJsonSource = LOCAL_STREET_JSON_RX.test(lowerFileName);
  const textReader = typeof file.text === "function" ? file.text.bind(file) : null;
  const bufferReader = typeof file.arrayBuffer === "function" ? file.arrayBuffer.bind(file) : null;
  const quiet = !!options.quiet;

  if (!isZipSource && !isJsonSource) {
    if (!quiet) alert("Unsupported file type. Select a .geojson/.json file or a Geofabrik .zip shapefile.");
    return false;
  }
  if (isZipSource && !bufferReader) {
    if (!quiet) alert("Unable to read selected ZIP file.");
    return false;
  }
  if (isJsonSource && !textReader) {
    if (!quiet) alert("Unable to read selected JSON/GeoJSON file.");
    return false;
  }
  if (
    isJsonSource &&
    fileSizeBytes >= streamThresholdBytes &&
    (!file || typeof file.stream !== "function")
  ) {
    if (!quiet) {
      alert(
        "This JSON file is very large and your browser does not support streamed file reads for it.\n\n" +
        "Use the offline converter again, then open the result in a Chromium-based browser."
      );
    }
    return false;
  }

  if (
    isJsonSource &&
    fileSizeBytes >= streamThresholdBytes &&
    file &&
    typeof file.stream === "function"
  ) {
    showLocalStreetLoadPercent(4, `Reading ${fileName}...`);
    await sleep(20);
    try {
      return await importLocalStreetFeaturesFromGeoJsonStream(file.stream().getReader(), fileName, {
        persistHandle,
        persistPath: persistPathValue,
        skipRemember: !!options.skipRemember,
        quiet,
        sourceSizeBytes: fileSizeBytes,
        filterBounds: chunkFilterBounds,
        preserveSourceDescriptor: true,
        sourceDescriptor,
        preserveToggleState: !!options.preserveToggleState
      });
    } catch (err) {
      finishLocalStreetLoadProgress("Could not load streets file.", true);
      if (!quiet) alert(`Could not load local streets file.\n\n${err?.message || err}`);
      return false;
    }
  }

  showLocalStreetLoadPercent(4, `Reading ${fileName}...`);
  await sleep(20);
  let features = [];
  let progressStartPct = 35;
  try {
    if (isZipSource) {
      progressStartPct = 90;
      ensureZipSizeWithinBrowserLimit(file.size, fileName);
      const sourceZip = await bufferReader();
      features = await extractStreetFeaturesFromZipBuffer(sourceZip, fileName);
    } else {
      const raw = await textReader();
      showLocalStreetLoadPercent(14, `Parsing ${fileName}...`);
      await sleep(20);
      const parsed = parseLocalStreetJsonPayload(raw, fileName, fileSizeBytes);
      showLocalStreetLoadPercent(24, `Extracting road features from ${fileName}...`);
      features = pickLocalStreetFeaturesFromParsedSource(parsed, false);
    }
  } catch (err) {
    finishLocalStreetLoadProgress("Could not load streets file.", true);
    if (!quiet) alert(`Could not load local streets file.\n\n${err?.message || err}`);
    return false;
  }

  if (!features.length) {
    finishLocalStreetLoadProgress("No street line features found.", true);
    if (!quiet) alert("No street line features were found in this file.");
    return false;
  }

  let persistPath = persistPathValue;
  if (!options.skipRemember && !persistHandle && !persistPath && !quiet) {
    const pastedPath = window.prompt(
      "To auto-open this streets file next time, paste its file path or URL (optional):",
      ""
    );
    persistPath = String(pastedPath || "").trim();
  }

  return importLocalStreetFeatures(features, fileName, {
    persistHandle,
    persistPath,
    skipRemember: !!options.skipRemember,
    filterBounds: chunkFilterBounds,
    preserveSourceDescriptor: true,
    sourceDescriptor: {
      ...sourceDescriptor,
      sourcePath: persistPath
    },
    preserveToggleState: !!options.preserveToggleState,
    progressStartPct,
    progressEndPct: 99,
    progressLabel: "Indexing local streets..."
  });
}

async function loadLocalStreetSourceFromPath(pathInput, options = {}) {
  const sourcePath = String(pathInput || "").trim();
  if (!sourcePath) {
    throw new Error("No file path was provided.");
  }

  const sourceUrl = normalizeLocalStreetSourcePathToUrl(sourcePath);
  const barePath = sourcePath.split(/[?#]/)[0];
  const bareUrl = sourceUrl.split(/[?#]/)[0];
  const isZipSource = LOCAL_STREET_ZIP_RX.test(barePath) || LOCAL_STREET_ZIP_RX.test(bareUrl);
  const isJsonSource = LOCAL_STREET_JSON_RX.test(barePath) || LOCAL_STREET_JSON_RX.test(bareUrl) || !isZipSource;
  if (!isZipSource && !isJsonSource) {
    throw new Error("Only .zip, .geojson, or .json paths are supported.");
  }

  showLocalStreetLoadPercent(4, `Reading ${sourcePath}...`);
  await sleep(20);

  let response;
  try {
    response = await fetch(sourceUrl, { cache: "no-store" });
  } catch (err) {
    finishLocalStreetLoadProgress("Could not read saved streets path.", true);
    const baseMsg = `Unable to open "${sourcePath}".`;
    const localHint = isLikelyLocalFilesystemPath(sourcePath)
      ? "\n\nBrowsers usually block direct local filesystem paths. Use Street Setup Wizard + backend instead."
      : "";
    throw new Error(`${baseMsg}\n${err?.message || err}${localHint}`);
  }

  if (!response.ok) {
    finishLocalStreetLoadProgress("Could not read saved streets path.", true);
    throw new Error(`File request failed with HTTP ${response.status}.`);
  }
  const responseSizeBytes = Number(response.headers.get("content-length") || 0);
  const streamThresholdBytes = LOCAL_STREET_JSON_STREAM_THRESHOLD_MB * 1024 * 1024;
  const displayName = sourcePath.split(/[\\/]/).pop() || sourcePath;
  const shouldPreferChunkedStream = !Number.isFinite(responseSizeBytes) ||
    responseSizeBytes <= 0 ||
    responseSizeBytes >= streamThresholdBytes ||
    isValidLeafletBounds(options.forceChunkBounds);
  const chunkFilterBounds = isValidLeafletBounds(options.forceChunkBounds)
    ? cloneLocalStreetBounds(options.forceChunkBounds)
    : (
      shouldPreferChunkedStream
        ? buildLocalStreetChunkBounds(map?.getBounds?.())
        : null
    );
  const sourceDescriptor = {
    sourceName: displayName,
    sourceSizeBytes: responseSizeBytes,
    sourcePath,
    sourceHandle: null,
    sourceFile: null,
    isChunkCapable: true
  };

  if (
    isJsonSource &&
    shouldPreferChunkedStream &&
    (!response.body || typeof response.body.getReader !== "function")
  ) {
    finishLocalStreetLoadProgress("Could not stream saved streets file.", true);
    throw new Error(
      "Saved streets JSON is too large for non-stream parsing in this browser. " +
      "Use the backend setup workflow or provide a smaller regional file."
    );
  }

  if (
    isJsonSource &&
    shouldPreferChunkedStream &&
    response.body &&
    typeof response.body.getReader === "function"
  ) {
    try {
      return await importLocalStreetFeaturesFromGeoJsonStream(response.body.getReader(), displayName, {
        persistPath: sourcePath,
        skipRemember: !!options.skipRemember,
        sourceSizeBytes: responseSizeBytes,
        filterBounds: chunkFilterBounds,
        preserveSourceDescriptor: true,
        sourceDescriptor,
        preserveToggleState: !!options.preserveToggleState
      });
    } catch (err) {
      finishLocalStreetLoadProgress("Could not parse saved streets file.", true);
      throw new Error(err?.message || err);
    }
  }

  let features = [];
  let progressStartPct = 35;
  try {
    if (isZipSource) {
      progressStartPct = 90;
      const contentLength = Number(response.headers.get("content-length") || 0);
      if (Number.isFinite(contentLength) && contentLength > 0) {
        ensureZipSizeWithinBrowserLimit(contentLength, sourcePath);
      }
      const sourceZip = await response.arrayBuffer();
      ensureZipSizeWithinBrowserLimit(sourceZip?.byteLength, sourcePath);
      features = await extractStreetFeaturesFromZipBuffer(sourceZip, sourcePath);
    } else {
      const raw = await response.text();
      showLocalStreetLoadPercent(14, `Parsing ${sourcePath}...`);
      await sleep(20);
      const parsed = parseLocalStreetJsonPayload(raw, sourcePath, responseSizeBytes);
      showLocalStreetLoadPercent(24, `Extracting road features from ${sourcePath}...`);
      features = pickLocalStreetFeaturesFromParsedSource(parsed, false);
    }
  } catch (err) {
    finishLocalStreetLoadProgress("Could not parse saved streets file.", true);
    throw new Error(err?.message || err);
  }

  if (!features.length) {
    finishLocalStreetLoadProgress("No street line features found in saved path.", true);
    throw new Error("No street line features were found in that file.");
  }

  return importLocalStreetFeatures(features, displayName, {
    persistPath: sourcePath,
    skipRemember: !!options.skipRemember,
    filterBounds: chunkFilterBounds,
    preserveSourceDescriptor: true,
    sourceDescriptor,
    preserveToggleState: !!options.preserveToggleState,
    progressStartPct,
    progressEndPct: 99,
    progressLabel: "Indexing local streets..."
  });
}

async function promptForMissingSavedStreetSource(error = null) {
  const reason = error?.message
    ? `\n\nSaved streets file could not be opened:\n${error.message}`
    : "";
  const wantsDownload = window.confirm(
    `Saved streets file was not found.${reason}\n\nPress OK to download a new Texas streets file.\nPress Cancel to paste a file path.`
  );
  if (wantsDownload) {
    startTexasStreetsDownload(true);
    return;
  }

  const pastedPath = window.prompt(
    "Paste the streets file path or URL (.zip, .geojson, or .json):",
    ""
  );
  if (!pastedPath) return;

  try {
    await loadLocalStreetSourceFromPath(pastedPath);
  } catch (err) {
    alert(`Could not load streets from that path.\n\n${err?.message || err}\n\nUse Street Setup Wizard instead.`);
  }
}

async function tryRestoreSavedStreetSourceOnStartup() {
  const meta = getLocalStreetSourceMeta();
  if (!meta) return;

  updateLocalStreetSourceStatus("Reopening saved local streets source...");
  let restoreError = null;

  if (meta.hasHandle) {
    try {
      const handle = await readStoredLocalStreetHandle();
      if (!handle) throw new Error("Saved local streets file handle is unavailable.");

      if (typeof handle.queryPermission === "function") {
        const permission = await handle.queryPermission({ mode: "read" });
        if (permission !== "granted") {
          throw new Error("Saved local streets file permission was not granted.");
        }
      }

      const savedFile = await handle.getFile();
      const ok = await loadLocalStreetSourceFile(savedFile, {
        quiet: true,
        persistHandle: handle,
        persistPath: meta.sourcePath || ""
      });
      if (ok) return;
      throw new Error("Saved local streets file could not be loaded.");
    } catch (err) {
      restoreError = err;
    }
  }

  if (meta.sourcePath) {
    try {
      const ok = await loadLocalStreetSourceFromPath(meta.sourcePath, { skipRemember: false });
      if (ok) return;
      throw new Error("Saved local streets path could not be loaded.");
    } catch (err) {
      restoreError = err;
    }
  }

  updateLocalStreetSourceStatus();
  await promptForMissingSavedStreetSource(restoreError);
}

async function openLocalStreetSourcePicker(fileInput) {
  if (typeof window.showOpenFilePicker !== "function") {
    fileInput?.click();
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "Street source files",
          accept: {
            "application/zip": [".zip"],
            "application/json": [".json", ".geojson"]
          }
        }
      ]
    });
    if (!handle) return;
    const file = await handle.getFile();
    await loadLocalStreetSourceFile(file, { persistHandle: handle });
  } catch (err) {
    if (err?.name === "AbortError") return;
    console.warn("showOpenFilePicker failed, falling back to file input.", err);
    fileInput?.click();
  }
}

function updateStreetSetupGuide() {
  const step1 = document.getElementById("streetSetupStep1");
  const step2 = document.getElementById("streetSetupStep2");
  const step3 = document.getElementById("streetSetupStep3");
  const hasSource = !!localStreetSourceState.loaded || (!!localStreetBackendState.available && !!localStreetBackendState.hasIndex);
  const backendReady = !!localStreetBackendState.available && !!localStreetBackendState.hasIndex;
  const toggle = document.getElementById("useLocalStreetSource");
  const streetsVisible = !!toggle?.checked && streetAttributeById.size > 0;

  step1?.classList.toggle("done", hasSource);
  step2?.classList.toggle("done", backendReady);
  step3?.classList.toggle("done", streetsVisible);
}

function setStreetWizardStatus(message = "Ready.") {
  const node = document.getElementById("streetWizardStatus");
  if (!node) return;
  node.textContent = message || "Ready.";
}

function openStreetSetupWizardModal() {
  const modal = document.getElementById("streetSetupWizardModal");
  if (!modal) return;
  updateStreetSetupGuide();
  setStreetWizardStatus(
    localStreetBackendState.available && localStreetBackendState.hasIndex
      ? "Backend is ready. Next: turn on Street Segments, then choose a saved polygon or draw a new one."
      : "Recommended path: Download Setup Program, run launcher, Check Backend, then start polygon load."
  );
  modal.style.display = "flex";
}

function closeStreetSetupWizardModal() {
  const modal = document.getElementById("streetSetupWizardModal");
  if (modal) modal.style.display = "none";
}

function openStreetNetworkManagerModal() {
  const modal = document.getElementById("streetNetworkManagerModal");
  if (!modal) return;
  updateLocalStreetSourceStatus();
  if (typeof window.__refreshStreetNetworkManagerUi === "function") {
    window.__refreshStreetNetworkManagerUi();
  }
  modal.style.display = "flex";
}

function closeStreetNetworkManagerModal() {
  const modal = document.getElementById("streetNetworkManagerModal");
  if (modal) modal.style.display = "none";
}

function initLocalStreetSourceControls() {
  const streetNetworkManagerBtn = document.getElementById("streetNetworkManagerBtn");
  const streetNetworkManagerModal = document.getElementById("streetNetworkManagerModal");
  const streetNetworkManagerClose = document.getElementById("streetNetworkManagerClose");
  const setupTabBtn = document.getElementById("streetNetworkSetupTabBtn");
  const symbologyTabBtn = document.getElementById("streetNetworkSymbologyTabBtn");
  const setupPanel = document.getElementById("streetNetworkSetupPanel");
  const symbologyPanel = document.getElementById("streetNetworkSymbologyPanel");
  const openStreetSetupWizardBtn = document.getElementById("openStreetSetupWizardBtn");
  const downloadTexasZipBtn = document.getElementById("downloadTexasZipBtn");
  const downloadAutoSetupPackageBtn = document.getElementById("downloadAutoSetupPackageBtn");
  const checkLocalBackendBtn = document.getElementById("checkLocalBackendBtn");
  const setLocalBackendUrlBtn = document.getElementById("setLocalBackendUrlBtn");
  const streetWizardModal = document.getElementById("streetSetupWizardModal");
  const streetWizardClose = document.getElementById("streetSetupWizardClose");
  const streetWizardDownloadTexasBtn = document.getElementById("streetWizardDownloadTexasBtn");
  const streetWizardAutoSetupBtn = document.getElementById("streetWizardAutoSetupBtn");
  const streetWizardCheckBackendBtn = document.getElementById("streetWizardCheckBackendBtn");
  const streetWizardStartBtn = document.getElementById("streetWizardStartBtn");
  const streetSymbologyEnabledInput = document.getElementById("streetSymbologyEnabled");
  const streetSymbologyFieldSelect = document.getElementById("streetSymbologyFieldSelect");
  const streetSymbologyRefreshBtn = document.getElementById("streetSymbologyRefreshBtn");
  const streetSymbologyWidthInput = document.getElementById("streetSymbologyWidthInput");
  const streetSymbologyWidthValue = document.getElementById("streetSymbologyWidthValue");
  const streetSymbologyOpacityInput = document.getElementById("streetSymbologyOpacityInput");
  const streetSymbologyOpacityValue = document.getElementById("streetSymbologyOpacityValue");
  const streetSymbologyClassMeta = document.getElementById("streetSymbologyClassMeta");
  const streetSymbologyClassList = document.getElementById("streetSymbologyClassList");
  const streetSymbologyStatus = document.getElementById("streetSymbologyStatus");
  const streetSymbologyApplyBtn = document.getElementById("streetSymbologyApplyBtn");
  const streetSymbologyResetBtn = document.getElementById("streetSymbologyResetBtn");
  const useLocalToggle = document.getElementById("useLocalStreetSource");

  let localStreetSourceFileInput = document.getElementById("localStreetSourceFileInput");
  if (!localStreetSourceFileInput) {
    localStreetSourceFileInput = document.createElement("input");
    localStreetSourceFileInput.type = "file";
    localStreetSourceFileInput.id = "localStreetSourceFileInput";
    localStreetSourceFileInput.accept = ".zip,.json,.geojson";
    localStreetSourceFileInput.hidden = true;
    document.body.appendChild(localStreetSourceFileInput);
  }

  if (!localStreetSourceFileInput.dataset.bound) {
    localStreetSourceFileInput.addEventListener("change", async e => {
      const file = e.target?.files?.[0];
      if (!file) return;
      try {
        await loadLocalStreetSourceFile(file);
        setStreetWizardStatus(`Loaded ${file.name}.`);
      } catch (err) {
        console.error("Failed loading local street source file:", err);
      } finally {
        e.target.value = "";
        updateStreetSetupGuide();
      }
    });
    localStreetSourceFileInput.dataset.bound = "1";
  }

  const openStreetSourceFilePicker = async () => {
    await openLocalStreetSourcePicker(localStreetSourceFileInput);
    updateStreetSetupGuide();
  };

  streetSymbologyState = loadStoredStreetSymbologyState();
  ensureStreetSymbologyState();

  const setStreetSymbologyStatus = message => {
    if (!streetSymbologyStatus) return;
    streetSymbologyStatus.textContent = String(message || "Symbology: Default style.");
  };

  const applyStreetSymbologyToMap = (statusMessage = "") => {
    ensureStreetSymbologyState();
    saveStoredStreetSymbologyState();
    applyStreetSelectionStyles();
    setStreetSymbologyStatus(statusMessage || "Symbology applied to loaded street segments.");
  };

  const setStreetNetworkManagerTab = (tabName, persist = true) => {
    const tab = String(tabName || "setup").toLowerCase() === "symbology" ? "symbology" : "setup";
    const isSymbology = tab === "symbology";
    setupTabBtn?.classList.toggle("active", !isSymbology);
    symbologyTabBtn?.classList.toggle("active", isSymbology);
    setupTabBtn?.setAttribute("aria-selected", !isSymbology ? "true" : "false");
    symbologyTabBtn?.setAttribute("aria-selected", isSymbology ? "true" : "false");
    setupPanel?.classList.toggle("active", !isSymbology);
    symbologyPanel?.classList.toggle("active", isSymbology);
    if (persist) {
      storageSet(STREET_NETWORK_MANAGER_TAB_KEY, tab);
    }
  };

  const renderStreetSymbologyClassList = () => {
    if (!streetSymbologyClassList || !streetSymbologyFieldSelect) return;
    ensureStreetSymbologyState();
    const selectedField = String(streetSymbologyFieldSelect.value || streetSymbologyState.field || "highway");
    const stats = getStreetSymbologyClassStats(selectedField);
    const maxRows = 80;
    const shown = stats.slice(0, maxRows);
    const hiddenCount = Math.max(0, stats.length - shown.length);

    if (streetSymbologyClassMeta) {
      const suffix = hiddenCount ? ` (${hiddenCount.toLocaleString()} more)` : "";
      streetSymbologyClassMeta.textContent = `${stats.length.toLocaleString()} values${suffix}`;
    }

    streetSymbologyClassList.innerHTML = "";
    if (!shown.length) {
      const empty = document.createElement("div");
      empty.className = "street-symbology-empty";
      empty.textContent = "No loaded street segments yet. Load streets, then refresh values.";
      streetSymbologyClassList.appendChild(empty);
      return;
    }

    shown.forEach(item => {
      const row = document.createElement("div");
      row.className = "street-symbology-class-item";

      const colorInput = document.createElement("input");
      colorInput.type = "color";
      colorInput.value = streetSymbologyState.valueColors?.[item.key] || getStreetSymbologyPaletteColorByKey(item.key);
      colorInput.disabled = !streetSymbologyState.enabled;

      const label = document.createElement("span");
      label.className = "street-symbology-class-item-label";
      label.textContent = item.label;

      const count = document.createElement("span");
      count.className = "street-symbology-class-item-count";
      count.textContent = item.count.toLocaleString();

      colorInput.addEventListener("input", () => {
        if (!streetSymbologyState.valueColors || typeof streetSymbologyState.valueColors !== "object") {
          streetSymbologyState.valueColors = {};
        }
        streetSymbologyState.valueColors[item.key] = String(colorInput.value || "").toLowerCase();
        applyStreetSymbologyToMap(`Updated color for ${item.label}.`);
      });

      row.append(colorInput, label, count);
      streetSymbologyClassList.appendChild(row);
    });
  };

  const refreshStreetSymbologyUi = () => {
    if (!streetSymbologyFieldSelect) return;
    ensureStreetSymbologyState();

    const fields = getStreetSymbologyAvailableFields();
    if (!fields.includes(streetSymbologyState.field)) {
      streetSymbologyState.field = fields[0] || "highway";
    }

    streetSymbologyFieldSelect.innerHTML = "";
    fields.forEach(field => {
      const option = document.createElement("option");
      option.value = field;
      option.textContent = field;
      streetSymbologyFieldSelect.appendChild(option);
    });
    streetSymbologyFieldSelect.value = streetSymbologyState.field;

    if (streetSymbologyEnabledInput) streetSymbologyEnabledInput.checked = !!streetSymbologyState.enabled;
    if (streetSymbologyWidthInput) streetSymbologyWidthInput.value = String(streetSymbologyState.lineWidth);
    if (streetSymbologyOpacityInput) streetSymbologyOpacityInput.value = String(streetSymbologyState.opacity);
    if (streetSymbologyWidthValue) streetSymbologyWidthValue.textContent = Number(streetSymbologyState.lineWidth).toFixed(1);
    if (streetSymbologyOpacityValue) streetSymbologyOpacityValue.textContent = Number(streetSymbologyState.opacity).toFixed(2);
    renderStreetSymbologyClassList();

    const symbologyStateLabel = streetSymbologyState.enabled
      ? `Symbology active: ${streetSymbologyState.field}`
      : "Symbology: Default style.";
    setStreetSymbologyStatus(symbologyStateLabel);
  };

  setupTabBtn?.addEventListener("click", () => setStreetNetworkManagerTab("setup"));
  symbologyTabBtn?.addEventListener("click", () => {
    setStreetNetworkManagerTab("symbology");
    refreshStreetSymbologyUi();
  });

  streetSymbologyEnabledInput?.addEventListener("change", () => {
    streetSymbologyState.enabled = !!streetSymbologyEnabledInput.checked;
    applyStreetSymbologyToMap(streetSymbologyState.enabled
      ? `Symbology enabled for [${streetSymbologyState.field}].`
      : "Symbology disabled. Default street style restored.");
    refreshStreetSymbologyUi();
  });

  streetSymbologyFieldSelect?.addEventListener("change", () => {
    streetSymbologyState.field = String(streetSymbologyFieldSelect.value || "highway");
    applyStreetSymbologyToMap(`Symbology field set to [${streetSymbologyState.field}].`);
    renderStreetSymbologyClassList();
  });

  streetSymbologyRefreshBtn?.addEventListener("click", () => {
    refreshStreetSymbologyUi();
    applyStreetSymbologyToMap("Symbology values refreshed.");
  });

  streetSymbologyWidthInput?.addEventListener("input", () => {
    const value = Number(streetSymbologyWidthInput.value);
    streetSymbologyState.lineWidth = Number.isFinite(value) ? Math.max(1, Math.min(8, value)) : STREET_BASE_LINE_WEIGHT;
    if (streetSymbologyWidthValue) streetSymbologyWidthValue.textContent = streetSymbologyState.lineWidth.toFixed(1);
    applyStreetSymbologyToMap(`Street line width set to ${streetSymbologyState.lineWidth.toFixed(1)}.`);
  });

  streetSymbologyOpacityInput?.addEventListener("input", () => {
    const value = Number(streetSymbologyOpacityInput.value);
    streetSymbologyState.opacity = Number.isFinite(value) ? Math.max(0.2, Math.min(1, value)) : STREET_BASE_LINE_OPACITY;
    if (streetSymbologyOpacityValue) streetSymbologyOpacityValue.textContent = streetSymbologyState.opacity.toFixed(2);
    applyStreetSymbologyToMap(`Street opacity set to ${streetSymbologyState.opacity.toFixed(2)}.`);
  });

  streetSymbologyApplyBtn?.addEventListener("click", () => {
    applyStreetSymbologyToMap(`Symbology applied by [${streetSymbologyState.field}].`);
    refreshStreetSymbologyUi();
  });

  streetSymbologyResetBtn?.addEventListener("click", () => {
    streetSymbologyState = normalizeStreetSymbologyState(null);
    applyStreetSymbologyToMap("Street symbology reset to default.");
    refreshStreetSymbologyUi();
  });

  const preferredTab = storageGet(STREET_NETWORK_MANAGER_TAB_KEY) || "setup";
  setStreetNetworkManagerTab(preferredTab, false);
  refreshStreetSymbologyUi();
  window.__refreshStreetNetworkManagerUi = () => {
    refreshStreetSymbologyUi();
    const nextTab = storageGet(STREET_NETWORK_MANAGER_TAB_KEY) || "setup";
    setStreetNetworkManagerTab(nextTab, false);
  };

  const runLocalBackendCheck = async (showAlert = true) => {
    updateLocalStreetSourceStatus("Checking local backend...");
    setStreetWizardStatus("Checking local backend...");
    await checkLocalStreetBackendAvailability(true);
    updateStreetSetupGuide();
    if (localStreetBackendState.available && localStreetBackendState.hasIndex) {
      setStreetWizardStatus("Backend is ready.");
      if (showAlert) alert("Backend is ready.");
      return true;
    }
    if (localStreetBackendState.available && !localStreetBackendState.hasIndex) {
      setStreetWizardStatus("Backend is running but no streets index is loaded.");
      if (showAlert) {
        alert(
          "Local backend is running, but no streets index is loaded.\n\n" +
          "Run the setup program (or indexer) first, then click Check Backend again."
        );
      }
      return false;
    }

    setStreetWizardStatus(`Backend not reachable at ${localStreetBackendState.baseUrl}.`);
    if (showAlert) {
      alert(
        "Local backend is not reachable.\n\n" +
        `URL: ${localStreetBackendState.baseUrl}\n` +
        "Start the backend server locally, then check again."
      );
    }
    return false;
  };

  localStreetBackendState.baseUrl = getStoredLocalStreetBackendUrl();

  if (useLocalToggle) {
    useLocalToggle.checked = false;
    useLocalToggle.disabled = !localStreetHasProvider();
  }

  streetNetworkManagerBtn?.addEventListener("click", () => {
    openStreetNetworkManagerModal();
  });

  streetNetworkManagerClose?.addEventListener("click", () => {
    closeStreetNetworkManagerModal();
  });

  streetNetworkManagerModal?.addEventListener("click", e => {
    if (e.target !== streetNetworkManagerModal) return;
    closeStreetNetworkManagerModal();
  });

  openStreetSetupWizardBtn?.addEventListener("click", () => {
    openStreetSetupWizardModal();
  });

  downloadTexasZipBtn?.addEventListener("click", () => {
    startTexasStreetsDownload(false);
    setStreetWizardStatus("Texas source ZIP download started.");
  });

  downloadAutoSetupPackageBtn?.addEventListener("click", () => {
    downloadLocalStreetAutoSetupPackage();
    setStreetWizardStatus("Auto setup package downloaded.");
  });

  checkLocalBackendBtn?.addEventListener("click", () => {
    runLocalBackendCheck(true).catch(err => {
      console.error("Backend check failed:", err);
    });
  });

  setLocalBackendUrlBtn?.addEventListener("click", async () => {
    const nextUrl = window.prompt(
      "Enter local backend URL:",
      localStreetBackendState.baseUrl || LOCAL_STREET_BACKEND_URL_DEFAULT
    );
    if (!nextUrl) return;
    setStoredLocalStreetBackendUrl(nextUrl);
    updateLocalStreetSourceStatus(`Backend URL set to ${localStreetBackendState.baseUrl}. Checking...`);
    await runLocalBackendCheck(false);
  });

  streetWizardClose?.addEventListener("click", () => {
    closeStreetSetupWizardModal();
  });

  streetWizardModal?.addEventListener("click", e => {
    if (e.target !== streetWizardModal) return;
    closeStreetSetupWizardModal();
  });

  streetWizardDownloadTexasBtn?.addEventListener("click", () => {
    startTexasStreetsDownload(false);
    setStreetWizardStatus("Texas source ZIP download started. Next: run setup program, then check backend.");
  });

  streetWizardAutoSetupBtn?.addEventListener("click", () => {
    downloadLocalStreetAutoSetupPackage();
    setStreetWizardStatus("Auto setup package downloaded. Run the launcher file, then click Check Backend.");
  });

  streetWizardCheckBackendBtn?.addEventListener("click", () => {
    runLocalBackendCheck(true).catch(err => {
      console.error("Backend check failed:", err);
    });
  });

  streetWizardStartBtn?.addEventListener("click", () => {
    if (!localStreetHasProvider()) {
      setStreetWizardStatus("Street source is not ready yet. Complete steps 1 and 2 first.");
      return;
    }
    closeStreetSetupWizardModal();
    if (useLocalToggle && !useLocalToggle.checked) {
      useLocalToggle.checked = true;
      useLocalToggle.dispatchEvent(new Event("change"));
      return;
    }
    if (!streetAttributeById.size) {
      streetLoadPolygonLayerGroup.clearLayers();
      streetPolygonLoadPending = false;
      openStreetSegmentsPromptModal();
    }
  });

  updateLocalStreetSourceStatus();
  checkLocalStreetBackendAvailability(false).catch(err => {
    console.warn("Local backend check failed:", err);
    updateLocalStreetSourceStatus();
    updateStreetSetupGuide();
  });
  updateStreetSetupGuide();
}

function clearRenderedStreetSegmentState() {
  streetAttributeLayerGroup.clearLayers();
  streetAttributeById.clear();
  streetAttributesRows = [];
  streetAttributeSelectedIds.clear();
  if (attributeTableMode === "streets") renderAttributeTable();
  updateStreetSetupGuide();
}

async function rebuildLocalStreetChunkForBounds(targetBounds) {
  const source = localStreetSourceState.sourceDescriptor;
  if (!source || typeof source !== "object") return false;
  const chunkBounds = buildLocalStreetChunkBounds(targetBounds);
  if (!isValidLeafletBounds(chunkBounds)) return false;

  clearRenderedStreetSegmentState();
  const sourceName = source.sourceName || localStreetSourceState.sourceName || "local streets source";
  updateStreetLoadStatus(`Refreshing local street chunk for this region from ${sourceName}...`);

  const runImportFromFile = async (file, handle = null) => {
    if (!file) return false;
    return loadLocalStreetSourceFile(file, {
      quiet: true,
      skipRemember: true,
      preserveToggleState: true,
      persistHandle: handle || source.sourceHandle || null,
      persistPath: String(source.sourcePath || "").trim(),
      forceChunkBounds: chunkBounds,
      preserveSourceDescriptor: true
    });
  };

  if (source.sourceHandle && typeof source.sourceHandle.getFile === "function") {
    try {
      const fileFromHandle = await source.sourceHandle.getFile();
      source.sourceFile = fileFromHandle;
      const ok = await runImportFromFile(fileFromHandle, source.sourceHandle);
      if (ok) return true;
    } catch (err) {
      console.warn("Failed to rebuild chunk from saved handle:", err);
    }
  }

  if (source.sourceFile) {
    try {
      const ok = await runImportFromFile(source.sourceFile, source.sourceHandle || null);
      if (ok) return true;
    } catch (err) {
      console.warn("Failed to rebuild chunk from in-memory file:", err);
    }
  }

  if (source.sourcePath) {
    try {
      const ok = await loadLocalStreetSourceFromPath(source.sourcePath, {
        skipRemember: true,
        preserveToggleState: true,
        forceChunkBounds: chunkBounds
      });
      if (ok) return true;
    } catch (err) {
      console.warn("Failed to rebuild chunk from saved path:", err);
    }
  }

  return false;
}

async function ensureLocalStreetChunkCoversBounds(targetBounds) {
  if (!localStreetSourceState.chunkMode) return true;
  const requested = isValidLeafletBounds(targetBounds)
    ? targetBounds
    : (map?.getBounds?.() || null);
  if (!isValidLeafletBounds(requested)) return true;

  if (
    isValidLeafletBounds(localStreetSourceState.chunkBounds) &&
    boundsContainsBounds(localStreetSourceState.chunkBounds, requested)
  ) {
    return true;
  }

  const rebuilt = await rebuildLocalStreetChunkForBounds(requested);
  if (!rebuilt) return false;

  return (
    isValidLeafletBounds(localStreetSourceState.chunkBounds) &&
    boundsContainsBounds(localStreetSourceState.chunkBounds, requested)
  );
}

function normalizeBackendStreetElement(raw) {
  const id = Number(raw?.id);
  if (!Number.isFinite(id)) return null;
  let tags = {};
  if (raw?.tags && typeof raw.tags === "object") {
    tags = raw.tags;
  } else if (typeof raw?.tags === "string") {
    try {
      const parsed = JSON.parse(raw.tags);
      if (parsed && typeof parsed === "object") tags = parsed;
    } catch {
      tags = {};
    }
  }
  const sourceGeom = Array.isArray(raw?.geom) ? raw.geom : [];
  const geom = [];
  sourceGeom.forEach(pt => {
    if (Array.isArray(pt) && pt.length >= 2) {
      const lat = Number(pt[0]);
      const lon = Number(pt[1]);
      if (Number.isFinite(lat) && Number.isFinite(lon)) geom.push({ lat, lon });
      return;
    }
    const lat = Number(pt?.lat);
    const lon = Number(pt?.lon);
    if (Number.isFinite(lat) && Number.isFinite(lon)) geom.push({ lat, lon });
  });
  if (geom.length < 2) return null;
  return {
    type: "way",
    id,
    tags: normalizeLocalStreetTags(tags),
    geom
  };
}

async function loadStreetAttributesFromLocalBackend(boundsOverride = null, polygonLayer = null) {
  const bounds = boundsOverride || map.getBounds();
  const south = bounds.getSouth();
  const west = bounds.getWest();
  const north = bounds.getNorth();
  const east = bounds.getEast();
  const query = new URLSearchParams({
    south: String(south),
    west: String(west),
    north: String(north),
    east: String(east),
    limit: String(LOCAL_STREET_BACKEND_QUERY_LIMIT)
  });

  const backendUrl = `${localStreetBackendState.baseUrl}/api/streets?${query.toString()}`;
  const response = await fetchJsonWithTimeout(
    backendUrl,
    { cache: "no-store" },
    LOCAL_STREET_BACKEND_REQUEST_TIMEOUT_MS
  );
  if (!response.ok) {
    throw new Error(`Backend HTTP ${response.status}`);
  }

  const payload = await response.json().catch(() => ({}));
  const rows = Array.isArray(payload?.elements) ? payload.elements : [];
  if (!rows.length) {
    return { addedCount: 0, totalCount: streetAttributeById.size, candidateCount: 0, knownCount: 0 };
  }

  let addedCount = 0;
  let knownCount = 0;
  const batchSize = 1200;
  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    batch.forEach(raw => {
      const element = normalizeBackendStreetElement(raw);
      if (!element) return;
      if (polygonLayer && !streetElementIntersectsPolygon(element, polygonLayer)) return;
      const normalizedTags = normalizeLocalStreetTags(element.tags || {});
      element.tags = normalizedTags;
      if (rowHasKnownStreetAttributes(normalizedTags)) knownCount += 1;
      if (upsertStreetElement(element)) addedCount += 1;
    });
    updateStreetLoadStatus(
      `Loading local backend streets... ${Math.min(i + batch.length, rows.length).toLocaleString()}/${rows.length.toLocaleString()}`
    );
    if (i + batchSize < rows.length) {
      // eslint-disable-next-line no-await-in-loop
      await sleep(0);
    }
  }

  streetAttributesRows = [...streetAttributeById.values()].map(v => v.row);
  syncStreetNetworkOverlay();
  if (attributeTableMode === "streets") renderAttributeTable();
  applyStreetSelectionStyles();

  return {
    addedCount,
    totalCount: streetAttributeById.size,
    candidateCount: rows.length,
    knownCount
  };
}

async function loadStreetAttributesFromLocalDataset(boundsOverride = null, polygonLayer = null) {
  const bounds = boundsOverride || map.getBounds();
  const candidates = [...getLocalStreetCandidateIds(bounds)];
  if (!candidates.length) {
    return { addedCount: 0, totalCount: streetAttributeById.size, candidateCount: 0, knownCount: 0 };
  }

  let addedCount = 0;
  let knownCount = 0;
  const batchSize = 1500;
  for (let i = 0; i < candidates.length; i += batchSize) {
    const batch = candidates.slice(i, i + batchSize);
    batch.forEach(id => {
      const element = localStreetSourceState.elementsById.get(id);
      if (!element) return;
      if (polygonLayer && !streetElementIntersectsPolygon(element, polygonLayer)) return;
      const normalizedTags = normalizeLocalStreetTags(element.tags || {});
      element.tags = normalizedTags;
      if (rowHasKnownStreetAttributes(normalizedTags)) knownCount += 1;
      if (upsertStreetElement(element)) addedCount += 1;
    });

    updateStreetLoadStatus(
      `Loading local street segments... ${Math.min(i + batch.length, candidates.length).toLocaleString()}/${candidates.length.toLocaleString()}`
    );
    if (i + batchSize < candidates.length) {
      // Yield between batches for responsiveness on large local datasets.
      // eslint-disable-next-line no-await-in-loop
      await sleep(0);
    }
  }

  streetAttributesRows = [...streetAttributeById.values()].map(v => v.row);
  syncStreetNetworkOverlay();
  if (attributeTableMode === "streets") renderAttributeTable();
  applyStreetSelectionStyles();

  return {
    addedCount,
    totalCount: streetAttributeById.size,
    candidateCount: candidates.length,
    knownCount
  };
}

function syncStreetNetworkOverlay() {
  const sourceToggle = document.getElementById("useLocalStreetSource");
  const sourceEnabled = !!sourceToggle && !!sourceToggle.checked;
  const layerVisible = isStreetNetworkLayerVisibleEnabled();
  const enabled = sourceEnabled && layerVisible;
  if (enabled) {
    if (!map.hasLayer(streetAttributeLayerGroup)) {
      streetAttributeLayerGroup.addTo(map);
    }
  } else {
    map.removeLayer(streetAttributeLayerGroup);
  }
  if (!sourceEnabled) {
    streetLoadPolygonLayerGroup.clearLayers();
    streetPolygonLoadPending = false;
    pendingStreetReload = false;
    if (streetAutoLoadTimer) {
      clearTimeout(streetAutoLoadTimer);
      streetAutoLoadTimer = null;
    }
    setStreetLoadBarVisible(false);
  }

  applyLayerManagerOrder();
  refreshLayerManagerUiIfOpen();
}

function roundStreetPolygonCoord(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return null;
  return Math.round(num * 1000000) / 1000000;
}

function encodeStreetPolygonLatLngs(value) {
  if (!Array.isArray(value) || !value.length) return [];
  const first = value[0];
  if (first && typeof first.lat === "number" && typeof first.lng === "number") {
    const coords = [];
    value.forEach(pt => {
      const lat = roundStreetPolygonCoord(pt?.lat);
      const lng = roundStreetPolygonCoord(pt?.lng);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
      coords.push([lat, lng]);
    });
    return coords;
  }
  if (Array.isArray(first) && first.length >= 2 && Number.isFinite(Number(first[0])) && Number.isFinite(Number(first[1]))) {
    const coords = [];
    value.forEach(pair => {
      if (!Array.isArray(pair) || pair.length < 2) return;
      const lat = roundStreetPolygonCoord(pair[0]);
      const lng = roundStreetPolygonCoord(pair[1]);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
      coords.push([lat, lng]);
    });
    return coords;
  }
  const nested = value
    .map(encodeStreetPolygonLatLngs)
    .filter(part => Array.isArray(part) && part.length > 0);
  return nested;
}

function decodeStreetPolygonLatLngs(value) {
  if (!Array.isArray(value) || !value.length) return [];
  const first = value[0];
  if (Array.isArray(first) && first.length >= 2 && Number.isFinite(Number(first[0])) && Number.isFinite(Number(first[1]))) {
    const latlngs = [];
    value.forEach(pair => {
      if (!Array.isArray(pair) || pair.length < 2) return;
      const lat = Number(pair[0]);
      const lng = Number(pair[1]);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
      latlngs.push(L.latLng(lat, lng));
    });
    return latlngs;
  }
  return value
    .map(decodeStreetPolygonLatLngs)
    .filter(part => Array.isArray(part) && part.length > 0);
}

function cloneStreetPolygonSnapshot(snapshot) {
  try {
    return JSON.parse(JSON.stringify(snapshot || []));
  } catch {
    return [];
  }
}

function normalizeSavedStreetPolygonEntry(entry) {
  if (!entry || typeof entry !== "object") return null;
  const normalizedName = String(entry.name || "").trim().slice(0, 120);
  const rawSnapshot = entry.snapshot ?? entry.latlngs ?? [];
  const encoded = encodeStreetPolygonLatLngs(rawSnapshot);
  const decoded = decodeStreetPolygonLatLngs(encoded);
  if (!Array.isArray(decoded) || !decoded.length) return null;
  const layer = L.polygon(decoded);
  const bounds = layer.getBounds?.();
  if (!bounds || !bounds.isValid?.()) return null;
  return {
    id: String(entry.id || `street_poly_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`),
    name: normalizedName || "Saved Polygon",
    snapshot: cloneStreetPolygonSnapshot(encoded),
    updatedAt: Number(entry.updatedAt || Date.now())
  };
}

function getSavedStreetPolygons() {
  const raw = storageGet(LOCAL_STREET_SAVED_POLYGONS_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    const normalized = parsed
      .map(normalizeSavedStreetPolygonEntry)
      .filter(Boolean);
    return normalized;
  } catch {
    return [];
  }
}

function setSavedStreetPolygons(polygons) {
  const normalized = (Array.isArray(polygons) ? polygons : [])
    .map(normalizeSavedStreetPolygonEntry)
    .filter(Boolean)
    .slice(0, LOCAL_STREET_SAVED_POLYGONS_MAX);
  storageSet(LOCAL_STREET_SAVED_POLYGONS_KEY, JSON.stringify(normalized));
  return normalized;
}

function formatStreetPolygonUpdatedAt(value) {
  const ts = Number(value);
  if (!Number.isFinite(ts)) return "Unknown date";
  try {
    return new Date(ts).toLocaleString();
  } catch {
    return "Unknown date";
  }
}

function createStreetPolygonLayerFromSnapshot(snapshot) {
  const latlngs = decodeStreetPolygonLatLngs(snapshot);
  if (!Array.isArray(latlngs) || !latlngs.length) return null;
  const layer = L.polygon(latlngs);
  const bounds = layer.getBounds?.();
  if (!bounds || !bounds.isValid?.()) return null;
  return layer;
}

async function loadStreetSegmentsFromSavedPolygon(savedPolygon) {
  const normalized = normalizeSavedStreetPolygonEntry(savedPolygon);
  if (!normalized) {
    alert("Saved polygon is invalid.");
    return;
  }

  if (!localStreetHasProvider()) {
    alert("Start the local backend first, or run the setup program.");
    return;
  }

  const polygonLayer = createStreetPolygonLayerFromSnapshot(normalized.snapshot);
  if (!polygonLayer) {
    alert("Saved polygon geometry is invalid.");
    return;
  }

  const toggle = document.getElementById("useLocalStreetSource");
  if (toggle && !toggle.checked) {
    toggle.checked = true;
    storageSet("streetSegmentsVisible", "on");
  }

  closeStreetSegmentsPromptModal();
  closeStreetPolygonLibraryModal();
  streetPolygonLoadPending = false;
  streetLoadPolygonLayerGroup.clearLayers();
  syncStreetNetworkOverlay();
  updateLocalStreetSourceStatus();

  lastStreetLoadPolygonSnapshot = cloneStreetPolygonSnapshot(normalized.snapshot);

  streetAttributeLayerGroup.clearLayers();
  streetAttributeById.clear();
  streetAttributesRows = [];
  streetAttributeSelectedIds.clear();
  if (attributeTableMode === "streets") renderAttributeTable();

  updateStreetLoadStatus(`Loading saved polygon "${normalized.name}"...`);
  await loadStreetAttributesForCurrentView(polygonLayer.getBounds(), polygonLayer);
}

function closeStreetPolygonLibraryModal() {
  const modal = document.getElementById("streetPolygonLibraryModal");
  if (modal) modal.style.display = "none";
}

function renderStreetPolygonLibraryModal() {
  const listNode = document.getElementById("streetPolygonLibraryList");
  if (!listNode) return;
  listNode.innerHTML = "";

  const polygons = getSavedStreetPolygons();
  if (!polygons.length) {
    const empty = document.createElement("div");
    empty.className = "street-polygon-empty";
    empty.textContent = "No saved polygons yet. Draw one, then click Save Polygon.";
    listNode.appendChild(empty);
    return;
  }

  polygons.forEach(poly => {
    const row = document.createElement("div");
    row.className = "street-polygon-library-item";

    const main = document.createElement("div");
    main.className = "street-polygon-library-item-main";

    const name = document.createElement("div");
    name.className = "street-polygon-library-name";
    name.textContent = poly.name;

    const meta = document.createElement("div");
    meta.className = "street-polygon-library-meta";
    meta.textContent = `Updated ${formatStreetPolygonUpdatedAt(poly.updatedAt)}`;

    main.appendChild(name);
    main.appendChild(meta);

    const actions = document.createElement("div");
    actions.className = "street-polygon-library-item-actions";

    const useBtn = document.createElement("button");
    useBtn.type = "button";
    useBtn.className = "mini-btn";
    useBtn.textContent = "Use";
    useBtn.addEventListener("click", () => {
      loadStreetSegmentsFromSavedPolygon(poly).catch(err => {
        console.error("Failed loading saved polygon:", err);
        alert(`Unable to load saved polygon.\n\n${err?.message || err}`);
      });
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "mini-btn";
    deleteBtn.textContent = "Delete";
    deleteBtn.addEventListener("click", () => {
      const confirmed = window.confirm(`Delete saved polygon "${poly.name}"?`);
      if (!confirmed) return;
      const next = getSavedStreetPolygons().filter(item => item.id !== poly.id);
      setSavedStreetPolygons(next);
      renderStreetPolygonLibraryModal();
    });

    actions.appendChild(useBtn);
    actions.appendChild(deleteBtn);
    row.appendChild(main);
    row.appendChild(actions);
    listNode.appendChild(row);
  });
}

function openStreetPolygonLibraryModal() {
  const modal = document.getElementById("streetPolygonLibraryModal");
  if (!modal) return;
  renderStreetPolygonLibraryModal();
  modal.style.display = "flex";
}

function saveCurrentStreetPolygonSnapshot() {
  const snapshot = cloneStreetPolygonSnapshot(lastStreetLoadPolygonSnapshot);
  if (!Array.isArray(snapshot) || !snapshot.length) {
    alert("Draw a street polygon first, then click Save Polygon.");
    return;
  }

  const defaultName = `Polygon ${new Date().toLocaleString()}`;
  const entered = window.prompt("Name this polygon:", defaultName);
  if (entered === null) return;
  const name = String(entered || "").trim();
  if (!name) {
    alert("Polygon name cannot be empty.");
    return;
  }

  const polygons = getSavedStreetPolygons();
  const next = [
    {
      id: `street_poly_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      name: name.slice(0, 120),
      snapshot,
      updatedAt: Date.now()
    },
    ...polygons
  ];
  const saved = setSavedStreetPolygons(next);
  updateLocalStreetSourceStatus(`Saved polygon "${name}" (${saved.length.toLocaleString()} total).`);
}

function openStreetSegmentsPromptModal() {
  const modal = document.getElementById("streetSegmentsPromptModal");
  if (modal) modal.style.display = "flex";
}

function closeStreetSegmentsPromptModal() {
  const modal = document.getElementById("streetSegmentsPromptModal");
  if (modal) modal.style.display = "none";
}

function beginStreetPolygonDraw() {
  streetPolygonLoadPending = true;
  const toolbarHandler = drawControl?._toolbars?.draw?._modes?.polygon?.handler;
  if (toolbarHandler && typeof toolbarHandler.enable === "function") {
    toolbarHandler.enable();
    return;
  }
  const drawHandler = new L.Draw.Polygon(map, {});
  drawHandler.enable();
}

function initStreetNetworkToggle() {
  const sourceToggle = document.getElementById("useLocalStreetSource");
  const layerToggle = document.getElementById("streetNetworkLayerToggle");
  const drawBtn = document.getElementById("streetSegmentsPromptDraw");
  const cancelBtn = document.getElementById("streetSegmentsPromptCancel");
  const promptModal = document.getElementById("streetSegmentsPromptModal");
  const savePolygonBtn = document.getElementById("saveStreetPolygonBtn");
  const openPolygonLibraryBtn = document.getElementById("openStreetPolygonLibraryBtn");
  const choosePolygonBtn = document.getElementById("chooseStreetPolygonBtn");
  const polygonLibraryModal = document.getElementById("streetPolygonLibraryModal");
  const polygonLibraryCloseBtn = document.getElementById("streetPolygonLibraryClose");

  if (layerToggle) {
    const storedLayerPref = storageGet(STREET_NETWORK_LAYER_VISIBLE_KEY);
    layerToggle.checked = storedLayerPref ? storedLayerPref === "on" : true;
    layerToggle.addEventListener("change", () => {
      storageSet(STREET_NETWORK_LAYER_VISIBLE_KEY, layerToggle.checked ? "on" : "off");
      syncStreetNetworkOverlay();
      updateLocalStreetSourceStatus();
    });
  }

  if (!sourceToggle) {
    syncStreetNetworkOverlay();
    updateLocalStreetSourceStatus();
    return;
  }

  sourceToggle.checked = false;
  storageSet("streetSegmentsVisible", "off");
  sourceToggle.addEventListener("change", () => {
    if (sourceToggle.checked && !localStreetHasProvider()) {
      sourceToggle.checked = false;
      storageSet("streetSegmentsVisible", "off");
      updateLocalStreetSourceStatus("Street source is not ready. Open Street Setup Wizard and follow steps 1-3.");
      openStreetSetupWizardModal();
      return;
    }

    storageSet("streetSegmentsVisible", sourceToggle.checked ? "on" : "off");
    syncStreetNetworkOverlay();
    updateLocalStreetSourceStatus();
    if (sourceToggle.checked) {
      if (!streetAttributeById.size) {
        streetLoadPolygonLayerGroup.clearLayers();
        streetPolygonLoadPending = false;
        openStreetSegmentsPromptModal();
      }
    } else {
      closeStreetSegmentsPromptModal();
      closeStreetPolygonLibraryModal();
      closeStreetSetupWizardModal();
    }
  });

  savePolygonBtn?.addEventListener("click", () => {
    saveCurrentStreetPolygonSnapshot();
  });

  const openSavedPolygonPicker = () => {
    closeStreetSegmentsPromptModal();
    openStreetPolygonLibraryModal();
  };

  choosePolygonBtn?.addEventListener("click", openSavedPolygonPicker);
  openPolygonLibraryBtn?.addEventListener("click", openSavedPolygonPicker);

  polygonLibraryCloseBtn?.addEventListener("click", () => {
    closeStreetPolygonLibraryModal();
  });

  polygonLibraryModal?.addEventListener("click", (e) => {
    if (e.target !== polygonLibraryModal) return;
    closeStreetPolygonLibraryModal();
  });

  drawBtn?.addEventListener("click", () => {
    if (!sourceToggle.checked) return;
    closeStreetSegmentsPromptModal();
    streetLoadPolygonLayerGroup.clearLayers();
    beginStreetPolygonDraw();
  });

  cancelBtn?.addEventListener("click", () => {
    closeStreetSegmentsPromptModal();
    closeStreetPolygonLibraryModal();
    sourceToggle.checked = false;
    storageSet("streetSegmentsVisible", "off");
    syncStreetNetworkOverlay();
    updateLocalStreetSourceStatus();
  });

  promptModal?.addEventListener("click", (e) => {
    if (e.target !== promptModal) return;
    closeStreetSegmentsPromptModal();
    closeStreetPolygonLibraryModal();
    sourceToggle.checked = false;
    storageSet("streetSegmentsVisible", "off");
    syncStreetNetworkOverlay();
    updateLocalStreetSourceStatus();
  });
  syncStreetNetworkOverlay();
  updateLocalStreetSourceStatus();
}

let attributeTableMode = "records";
let streetAttributesRows = [];
const streetAttributeSelectedIds = new Set();
const streetAttributeById = new Map();
let streetLoadInFlight = false;
let lastStreetLoadAt = 0;
let streetAutoLoadTimer = null;
let pendingStreetReload = false;
let streetLoadBarHideTimer = null;
let streetLoadBarShownAt = 0;
const STREET_MAX_LOAD_SPAN = 0.9;

function isStreetLoadableView() {
  return true;
}

function getStreetLoadBoundsForView(bounds) {
  const south = bounds.getSouth();
  const west = bounds.getWest();
  const north = bounds.getNorth();
  const east = bounds.getEast();
  const spanLat = Math.abs(north - south);
  const spanLng = Math.abs(east - west);

  if (spanLat <= STREET_MAX_LOAD_SPAN && spanLng <= STREET_MAX_LOAD_SPAN) {
    return { south, west, north, east, wasClamped: false, spanLat, spanLng };
  }

  const c = bounds.getCenter();
  const halfLat = Math.min(spanLat, STREET_MAX_LOAD_SPAN) / 2;
  const halfLng = Math.min(spanLng, STREET_MAX_LOAD_SPAN) / 2;
  return {
    south: c.lat - halfLat,
    west: c.lng - halfLng,
    north: c.lat + halfLat,
    east: c.lng + halfLng,
    wasClamped: true,
    spanLat,
    spanLng
  };
}

function hasStreetSegmentsInView() {
  const view = map.getBounds();
  let found = false;
  streetAttributeLayerGroup.eachLayer(layer => {
    if (found) return;
    if (typeof layer.getBounds === "function" && layer.getBounds().intersects(view)) {
      found = true;
    }
  });
  return found;
}

function getPolygonRings(latlngs) {
  if (!Array.isArray(latlngs) || !latlngs.length) return [];
  if (latlngs[0] && typeof latlngs[0].lat === "number" && typeof latlngs[0].lng === "number") {
    return [latlngs];
  }
  return latlngs.flatMap(getPolygonRings);
}

const polygonIntersectionCache = new WeakMap();

function getPolygonIntersectionData(polygonLayer) {
  if (!polygonLayer || !(polygonLayer instanceof L.Polygon || polygonLayer instanceof L.Rectangle)) {
    return null;
  }
  const cached = polygonIntersectionCache.get(polygonLayer);
  if (cached) return cached;

  const rings = getPolygonRings(polygonLayer.getLatLngs());
  const edges = [];
  rings.forEach(ring => {
    if (!Array.isArray(ring) || ring.length < 2) return;
    for (let i = 0; i < ring.length; i++) {
      const a = ring[i];
      const b = ring[(i + 1) % ring.length];
      edges.push([a, b]);
    }
  });

  const data = {
    bounds: polygonLayer.getBounds?.() || null,
    rings,
    edges
  };
  polygonIntersectionCache.set(polygonLayer, data);
  return data;
}

function isPointInRing(latlng, ring) {
  const x = latlng.lng;
  const y = latlng.lat;
  let inside = false;
  for (let i = 0, j = ring.length - 1; i < ring.length; j = i++) {
    const xi = ring[i].lng;
    const yi = ring[i].lat;
    const xj = ring[j].lng;
    const yj = ring[j].lat;
    const intersects = ((yi > y) !== (yj > y))
      && (x < ((xj - xi) * (y - yi)) / ((yj - yi) || Number.EPSILON) + xi);
    if (intersects) inside = !inside;
  }
  return inside;
}

function isLatLngInsidePolygon(latlng, polygonLayer) {
  if (!polygonLayer) return true;
  const data = getPolygonIntersectionData(polygonLayer);
  const bounds = data?.bounds || polygonLayer.getBounds?.();
  if (bounds && !bounds.contains(latlng)) return false;
  if (polygonLayer instanceof L.Rectangle) return true;
  if (polygonLayer instanceof L.Polygon) {
    const rings = data?.rings || getPolygonRings(polygonLayer.getLatLngs());
    return rings.some(ring => isPointInRing(latlng, ring));
  }
  return true;
}

function cross2d(a, b, c) {
  return (b.lng - a.lng) * (c.lat - a.lat) - (b.lat - a.lat) * (c.lng - a.lng);
}

function isPointOnSegment(a, b, p) {
  const minX = Math.min(a.lng, b.lng) - 1e-10;
  const maxX = Math.max(a.lng, b.lng) + 1e-10;
  const minY = Math.min(a.lat, b.lat) - 1e-10;
  const maxY = Math.max(a.lat, b.lat) + 1e-10;
  if (p.lng < minX || p.lng > maxX || p.lat < minY || p.lat > maxY) return false;
  return Math.abs(cross2d(a, b, p)) <= 1e-10;
}

function segmentsIntersect(a, b, c, d) {
  const ab_c = cross2d(a, b, c);
  const ab_d = cross2d(a, b, d);
  const cd_a = cross2d(c, d, a);
  const cd_b = cross2d(c, d, b);

  if ((ab_c > 0 && ab_d < 0 || ab_c < 0 && ab_d > 0) &&
      (cd_a > 0 && cd_b < 0 || cd_a < 0 && cd_b > 0)) {
    return true;
  }

  if (Math.abs(ab_c) <= 1e-10 && isPointOnSegment(a, b, c)) return true;
  if (Math.abs(ab_d) <= 1e-10 && isPointOnSegment(a, b, d)) return true;
  if (Math.abs(cd_a) <= 1e-10 && isPointOnSegment(c, d, a)) return true;
  if (Math.abs(cd_b) <= 1e-10 && isPointOnSegment(c, d, b)) return true;
  return false;
}

function streetElementIntersectsPolygon(element, polygonLayer) {
  if (!polygonLayer) return true;
  const polygonData = getPolygonIntersectionData(polygonLayer);
  const geometry = Array.isArray(element?.geom)
    ? element.geom
    : (Array.isArray(element?.geometry) ? element.geometry : []);
  if (!geometry.length) return false;
  const points = geometry
    .map(pt => {
      const lat = Number(pt?.lat);
      const lon = Number(pt?.lon);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
      return L.latLng(lat, lon);
    })
    .filter(Boolean);

  if (!points.length) return false;
  if (points.some(point => isLatLngInsidePolygon(point, polygonLayer))) return true;

  const polygonEdges = polygonData?.edges || [];
  if (!polygonEdges.length) return false;

  for (let i = 1; i < points.length; i++) {
    const a = points[i - 1];
    const b = points[i];
    const segmentBounds = L.latLngBounds(a, b);
    if (polygonData?.bounds && !polygonData.bounds.intersects(segmentBounds)) continue;
    for (let j = 0; j < polygonEdges.length; j++) {
      const edge = polygonEdges[j];
      if (segmentsIntersect(a, b, edge[0], edge[1])) return true;
    }
  }

  return false;
}

function maybeAutoLoadStreetSegments() {
  const toggle = document.getElementById("useLocalStreetSource");
  if (!toggle || !toggle.checked) return;
  if (streetLoadInFlight) {
    pendingStreetReload = true;
    setStreetLoadBarVisible(true);
    return;
  }
  const elapsed = Date.now() - lastStreetLoadAt;
  if (elapsed < 900) {
    if (streetAutoLoadTimer) clearTimeout(streetAutoLoadTimer);
    setStreetLoadBarVisible(true);
    const wait = Math.max(120, 900 - elapsed);
    streetAutoLoadTimer = setTimeout(() => {
      if (!streetLoadInFlight) {
        loadStreetAttributesForCurrentView().catch(() => {});
      } else {
        pendingStreetReload = true;
        setStreetLoadBarVisible(true);
      }
    }, wait);
    return;
  }

  if (streetAutoLoadTimer) clearTimeout(streetAutoLoadTimer);
  setStreetLoadBarVisible(true);
  streetAutoLoadTimer = setTimeout(() => {
    if (!streetLoadInFlight) {
      loadStreetAttributesForCurrentView().catch(() => {});
    }
  }, 140);
}

function setStreetLoadBarVisible(visible) {
  const bar = document.getElementById("streetSegmentsLoadBar");
  if (!bar) return;
  if (streetLoadBarHideTimer) {
    clearTimeout(streetLoadBarHideTimer);
    streetLoadBarHideTimer = null;
  }
  if (visible) {
    streetLoadBarShownAt = Date.now();
    bar.classList.add("active");
    return;
  }
  const elapsed = Date.now() - streetLoadBarShownAt;
  const wait = Math.max(180, 520 - elapsed);
  streetLoadBarHideTimer = setTimeout(() => {
    bar.classList.remove("active");
  }, wait);
}

function buildStreetBoundsChunks(south, west, north, east, maxSpan = STREET_MAX_LOAD_SPAN) {
  const latSpan = Math.max(0, north - south);
  const lngSpan = Math.max(0, east - west);
  const rows = Math.max(1, Math.ceil(latSpan / maxSpan));
  const cols = Math.max(1, Math.ceil(lngSpan / maxSpan));
  const latStep = latSpan / rows;
  const lngStep = lngSpan / cols;
  const chunks = [];

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const s = south + (r * latStep);
      const n = r === rows - 1 ? north : (south + ((r + 1) * latStep));
      const w = west + (c * lngStep);
      const e = c === cols - 1 ? east : (west + ((c + 1) * lngStep));
      chunks.push({ south: s, west: w, north: n, east: e });
    }
  }

  return chunks;
}

function sortStreetChunksByFocus(chunks, focusLatLng) {
  if (!Array.isArray(chunks) || !chunks.length || !focusLatLng) return chunks || [];
  return [...chunks].sort((a, b) => {
    const aLat = (a.south + a.north) / 2;
    const aLng = (a.west + a.east) / 2;
    const bLat = (b.south + b.north) / 2;
    const bLng = (b.west + b.east) / 2;
    const da = ((aLat - focusLatLng.lat) ** 2) + ((aLng - focusLatLng.lng) ** 2);
    const db = ((bLat - focusLatLng.lat) ** 2) + ((bLng - focusLatLng.lng) ** 2);
    return da - db;
  });
}

function updateStreetLoadStatus(message = "", isError = false) {
  const node = document.getElementById("streetLoadStatus");
  if (!node) return;
  const show = attributeTableMode === "streets" && !!String(message || "").trim();
  node.textContent = message || "";
  node.classList.toggle("show", show);
  node.classList.toggle("error", !!(show && isError));
}

function setAttributeTableMode(mode) {
  attributeTableMode = mode === "streets" ? "streets" : "records";
  const title = document.querySelector(".attribute-title");
  const streetBtn = document.getElementById("streetAttributesBtn");
  const streetBtnMobile = document.getElementById("streetAttributesBtnMobile");
  const attrBtn = document.getElementById("attributeTableBtn");
  const attrBtnMobile = document.getElementById("attributeTableBtnMobile");
  const selectedOnlyToggle = document.getElementById("attributeSelectedOnly");

  if (title) title.textContent = attributeTableMode === "streets" ? "Street Attributes" : "Attribute Table";
  streetBtn?.classList.toggle("active", attributeTableMode === "streets");
  streetBtnMobile?.classList.toggle("active", attributeTableMode === "streets");
  attrBtn?.classList.toggle("active", attributeTableMode === "records");
  attrBtnMobile?.classList.toggle("active", attributeTableMode === "records");

  if (attributeTableMode === "streets") {
    attributeState.selectedOnly = false;
    if (selectedOnlyToggle) selectedOnlyToggle.checked = false;
  }

  if (attributeTableMode === "streets") updateStreetLoadStatus("Loading street segments...");
  else updateStreetLoadStatus("");

  const selectedCount = attributeTableMode === "streets"
    ? streetAttributeSelectedIds.size
    : attributeState.selectedRowIds.size;
  syncSelectedStopsHeaderCount(selectedCount);
  if (typeof window.__refreshSelectionToolsUi === "function") {
    window.__refreshSelectionToolsUi();
  }
}

function setStreetSegmentStyle(entry, selected) {
  if (!entry?.layer) return;
  const baseStyle = getStreetSegmentBaseStyle(entry);
  const selectedWeight = Math.max((Number(baseStyle.weight) || STREET_BASE_LINE_WEIGHT) + 2, 5);
  entry.layer.setStyle(
    selected
      ? { color: "#ffe066", weight: selectedWeight, opacity: 0.96 }
      : baseStyle
  );
}

function applyStreetSelectionStyles() {
  streetAttributeById.forEach(entry => {
    setStreetSegmentStyle(entry, streetAttributeSelectedIds.has(entry.id));
  });
}

function toggleStreetSegmentSelection(id, selected = null, rerender = true) {
  const wayId = Number(id);
  if (!Number.isFinite(wayId)) return;
  const shouldSelect = selected === null ? !streetAttributeSelectedIds.has(wayId) : !!selected;
  if (shouldSelect) streetAttributeSelectedIds.add(wayId);
  else streetAttributeSelectedIds.delete(wayId);
  const entry = streetAttributeById.get(wayId);
  setStreetSegmentStyle(entry, streetAttributeSelectedIds.has(wayId));
  syncSelectedStopsHeaderCount(streetAttributeSelectedIds.size);
  refreshAttributeStatus();
  if (rerender && attributeTableMode === "streets") renderAttributeTable();
}

function getFilteredStreetAttributeRows() {
  const needle = String(attributeState.filterText || "").trim().toLowerCase();
  const allRows = (Array.isArray(streetAttributesRows) && streetAttributesRows.length)
    ? [...streetAttributesRows]
    : Array.from(streetAttributeById.values()).map(entry => entry?.row).filter(Boolean);
  let rows = attributeState.selectedOnly
    ? allRows.filter(row => streetAttributeSelectedIds.has(Number(row?.id)))
    : allRows;
  if (needle) {
    rows = rows.filter(r =>
      Object.values(r).some(v => String(v ?? "").toLowerCase().includes(needle))
    );
  }
  if (attributeState.sortKey) {
    const key = attributeState.sortKey;
    const dir = attributeState.sortDir;
    rows.sort((a, b) => {
      const av = a[key];
      const bv = b[key];
      const an = Number(av);
      const bn = Number(bv);
      if (Number.isFinite(an) && Number.isFinite(bn)) return (an - bn) * dir;
      return String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true }) * dir;
    });
  }
  return rows;
}

function renderStreetAttributeTable() {
  const table = document.getElementById("attributeTableGrid");
  const empty = document.getElementById("attributeTableEmpty");
  const pageInfo = document.getElementById("attributePageInfo");
  const prevBtn = document.getElementById("attributePrevPageBtn");
  const nextBtn = document.getElementById("attributeNextPageBtn");
  const status = document.getElementById("attributeStatus");
  if (!table || !empty) return;

  if (!streetAttributeById.size) {
    table.innerHTML = "";
    empty.textContent = "No street data loaded. Turn on local streets and draw/load a polygon first.";
    empty.style.display = "block";
    if (pageInfo) pageInfo.textContent = "Page 1/1";
    if (prevBtn) prevBtn.disabled = true;
    if (nextBtn) nextBtn.disabled = true;
    if (status) status.textContent = `${streetAttributeSelectedIds.size} selected`;
    return;
  }

  const headers = ["id", "name", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"];
  const labels = {
    id: "Way ID",
    name: "Road Name",
    highway: "Class",
    ref: "Ref",
    maxspeed: "Max Speed",
    lanes: "Lanes",
    surface: "Surface",
    oneway: "One Way"
  };
  const rows = getFilteredStreetAttributeRows();
  attributeState.lastVisibleRows = rows.map(r => ({ row: r, rowId: Number(r.id) }));
  const totalPages = Math.max(1, Math.ceil(rows.length / attributeState.pageSize));
  if (attributeState.page > totalPages) attributeState.page = totalPages;
  if (attributeState.page < 1) attributeState.page = 1;
  const pageStart = (attributeState.page - 1) * attributeState.pageSize;
  const pageRows = rows.slice(pageStart, pageStart + attributeState.pageSize);

  if (pageInfo) pageInfo.textContent = `Page ${attributeState.page}/${totalPages}`;
  if (prevBtn) prevBtn.disabled = attributeState.page <= 1;
  if (nextBtn) nextBtn.disabled = attributeState.page >= totalPages;
  if (status) {
    const scopeLabel = attributeState.selectedOnly ? "selected scope" : "loaded scope";
    status.textContent = `${streetAttributeSelectedIds.size} selected • ${rows.length} visible (${scopeLabel})`;
  }

  if (!rows.length) {
    table.innerHTML = "";
    empty.textContent = attributeState.selectedOnly
      ? "No selected street segments. Use selection tools or click street lines on the map."
      : "No street segments match the current filter.";
    empty.style.display = "block";
    return;
  }
  empty.style.display = "none";

  const sortIndicator = key => (attributeState.sortKey === key ? (attributeState.sortDir > 0 ? " ▲" : " ▼") : "");
  let html = "<thead><tr><th>Sel</th><th>#</th>";
  headers.forEach(h => {
    html += `<th><button type="button" data-sort="${h}">${labels[h]}${sortIndicator(h)}</button></th>`;
  });
  html += "</tr></thead><tbody>";

  pageRows.forEach((row, idx) => {
    const rowId = Number(row?.id);
    const checked = streetAttributeSelectedIds.has(rowId) ? " checked" : "";
    html += `<tr data-row-id="${rowId}" class="${checked ? "selected" : ""}">`;
    html += `<td><input type="checkbox" data-row-select="${rowId}"${checked}></td>`;
    html += `<td>${pageStart + idx + 1}</td>`;
    headers.forEach(h => {
      const value = row[h];
      html += `<td>${String(value ?? "-").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody>";
  table.innerHTML = html;

  table.querySelectorAll("button[data-sort]").forEach(btn => {
    btn.addEventListener("click", e => {
      const key = e.currentTarget.getAttribute("data-sort");
      if (!key) return;
      if (attributeState.sortKey === key) attributeState.sortDir *= -1;
      else {
        attributeState.sortKey = key;
        attributeState.sortDir = 1;
      }
      attributeState.page = 1;
      renderAttributeTable();
    });
  });

  table.querySelectorAll("input[data-row-select]").forEach(input => {
    input.addEventListener("change", e => {
      const rowId = Number(e.currentTarget.getAttribute("data-row-select"));
      toggleStreetSegmentSelection(rowId, e.currentTarget.checked, true);
    });
  });

  table.querySelectorAll("tbody tr").forEach(tr => {
    tr.addEventListener("click", e => {
      if (e.target.closest("input")) return;
      const id = Number(tr.getAttribute("data-row-id"));
      const entry = streetAttributeById.get(id);
      if (entry?.layer) map.fitBounds(entry.layer.getBounds().pad(0.35));
    });
  });
}

const OVERPASS_ENDPOINTS = [
  "https://overpass-api.de/api/interpreter",
  "https://overpass.kumi.systems/api/interpreter",
  "https://lz4.overpass-api.de/api/interpreter"
];
const STREET_TILE_CACHE_TTL_MS = 5 * 60 * 1000;
const OVERPASS_BASE_COOLDOWN_MS = 60 * 1000;
const OVERPASS_MAX_COOLDOWN_MS = 5 * 60 * 1000;
const OVERPASS_ENDPOINT_TIMEOUT_MS = 9000;
const STREET_DEFAULT_CONCURRENCY = 2;
const STREET_CHUNK_CONCURRENCY = 1;
const STREET_BATCH_GAP_MS = 90;
const STREET_CHUNK_GAP_MS = 120;
const STREET_POLYGON_CHUNK_SPAN = 0.35;
const OSM_SPLIT_MAX_DEPTH = 2;
const OSM_SPLIT_STATUS_CODES = new Set([400, 413]);
const streetTileCache = new Map();
let overpassCooldownUntil = 0;
let overpass429Count = 0;

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function makeStreetTileKey(tile) {
  return [tile.south, tile.west, tile.north, tile.east].map(v => Number(v).toFixed(5)).join("|");
}

async function fetchOverpassJsonWithFallback(queryText) {
  if (Date.now() < overpassCooldownUntil) {
    const waitSec = Math.ceil((overpassCooldownUntil - Date.now()) / 1000);
    throw new Error(`Overpass cooling down (${waitSec}s remaining)`);
  }
  const errors = [];

  for (const endpoint of OVERPASS_ENDPOINTS) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), OVERPASS_ENDPOINT_TIMEOUT_MS);
    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=UTF-8" },
        body: queryText,
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      if (!response.ok) {
        if (response.status === 429) {
          overpass429Count += 1;
          const cooldownMs = Math.min(
            OVERPASS_BASE_COOLDOWN_MS * Math.pow(2, Math.max(0, overpass429Count - 1)),
            OVERPASS_MAX_COOLDOWN_MS
          );
          overpassCooldownUntil = Date.now() + cooldownMs;
          errors.push(`${endpoint} -> HTTP ${response.status}`);
          break;
        }
        errors.push(`${endpoint} -> HTTP ${response.status}`);
        continue;
      }
      overpass429Count = 0;
      overpassCooldownUntil = 0;

      return await response.json();
    } catch (err) {
      clearTimeout(timeoutId);
      const message = err?.name === "AbortError" ? "timeout" : (err?.message || "request failed");
      errors.push(`${endpoint} -> ${message}`);
    }
  }

  throw new Error(`All Overpass endpoints failed: ${errors.join(" | ")}`);
}

function buildStreetTileBounds(south, west, north, east) {
  const spanLat = Math.abs(north - south);
  const spanLng = Math.abs(east - west);
  const area = spanLat * spanLng;

  let grid = 1;
  if (area > 0.02) grid = 2;
  if (area > 0.06) grid = 3;
  if (area > 0.12) grid = 3;

  const latStep = spanLat / grid;
  const lngStep = spanLng / grid;
  const tiles = [];

  for (let r = 0; r < grid; r++) {
    for (let c = 0; c < grid; c++) {
      const s = south + (r * latStep);
      const n = r === grid - 1 ? north : (south + ((r + 1) * latStep));
      const w = west + (c * lngStep);
      const e = c === grid - 1 ? east : (west + ((c + 1) * lngStep));
      tiles.push({ south: s, west: w, north: n, east: e });
    }
  }

  return tiles;
}

function buildStreetOverpassQuery(tile) {
  const s = tile.south;
  const w = tile.west;
  const n = tile.north;
  const e = tile.east;
  return `
[out:json][timeout:20];
(
  way["highway"](${s},${w},${n},${e});
);
out geom tags;
`;
}

async function fetchStreetElementsFromOsmApi(tile) {
  return await fetchStreetElementsFromOsmApiRecursive(tile, 0);
}

function splitStreetTile(tile) {
  const midLat = (tile.south + tile.north) / 2;
  const midLng = (tile.west + tile.east) / 2;
  return [
    { south: tile.south, west: tile.west, north: midLat, east: midLng },
    { south: tile.south, west: midLng, north: midLat, east: tile.east },
    { south: midLat, west: tile.west, north: tile.north, east: midLng },
    { south: midLat, west: midLng, north: tile.north, east: tile.east }
  ];
}

async function fetchStreetElementsFromOsmApiRecursive(tile, depth = 0) {
  const bbox = `${tile.west},${tile.south},${tile.east},${tile.north}`;
  const url = `https://api.openstreetmap.org/api/0.6/map?bbox=${encodeURIComponent(bbox)}`;
  const response = await fetch(url, { method: "GET" });
  if (!response.ok) {
    const detail = (await response.text().catch(() => "")).trim().slice(0, 180);
    const canSplit = OSM_SPLIT_STATUS_CODES.has(response.status) && depth < OSM_SPLIT_MAX_DEPTH;
    if (canSplit) {
      const chunks = splitStreetTile(tile);
      const merged = new Map();
      for (let i = 0; i < chunks.length; i++) {
        const list = await fetchStreetElementsFromOsmApiRecursive(chunks[i], depth + 1);
        list.forEach(e => {
          if (!e || !Number.isFinite(Number(e.id))) return;
          merged.set(Number(e.id), e);
        });
        if (i < chunks.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      }
      return [...merged.values()];
    }
    throw new Error(`OSM API HTTP ${response.status}${detail ? `: ${detail}` : ""}`);
  }

  const xmlText = await response.text();
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  if (doc.querySelector("parsererror")) {
    throw new Error("OSM API XML parse error");
  }

  const nodeMap = new Map();
  doc.querySelectorAll("node").forEach(node => {
    const id = Number(node.getAttribute("id"));
    const lat = Number(node.getAttribute("lat"));
    const lon = Number(node.getAttribute("lon"));
    if (Number.isFinite(id) && Number.isFinite(lat) && Number.isFinite(lon)) {
      nodeMap.set(id, { lat, lon });
    }
  });

  const elements = [];
  doc.querySelectorAll("way").forEach(way => {
    const id = Number(way.getAttribute("id"));
    if (!Number.isFinite(id)) return;

    const tags = {};
    way.querySelectorAll("tag").forEach(tag => {
      const k = tag.getAttribute("k");
      const v = tag.getAttribute("v");
      if (k) tags[k] = v || "";
    });

    if (!tags.highway) return;

    const geom = [];
    way.querySelectorAll("nd").forEach(nd => {
      const ref = Number(nd.getAttribute("ref"));
      const point = nodeMap.get(ref);
      if (point) geom.push({ lat: point.lat, lon: point.lon });
    });

    if (geom.length < 2) return;
    elements.push({ type: "way", id, tags, geom });
  });

  return elements;
}

async function fetchStreetElementsForTile(tile) {
  const key = makeStreetTileKey(tile);
  const cached = streetTileCache.get(key);
  if (cached && (Date.now() - cached.ts) < STREET_TILE_CACHE_TTL_MS) {
    return cached.elements;
  }

  try {
    // OSM API is generally faster/more predictable for this workflow.
    const elements = await fetchStreetElementsFromOsmApi(tile);
    streetTileCache.set(key, { ts: Date.now(), elements });
    return elements;
  } catch (osmErr) {
    const osmMsg = String(osmErr?.message || "");
    if (osmMsg.includes("HTTP 509")) {
      // Back off quickly from OSM API bandwidth throttling.
      await sleep(120);
    }
    const query = buildStreetOverpassQuery(tile);
    const data = await fetchOverpassJsonWithFallback(query);
    const elements = data.elements || [];
    streetTileCache.set(key, { ts: Date.now(), elements });
    return elements;
  }
}

async function collectStreetElementsForBounds(south, west, north, east, options = {}) {
  const tiles = options.singleTile
    ? [{ south, west, north, east }]
    : buildStreetTileBounds(south, west, north, east);
  const mergedById = new Map();
  const tileErrors = [];
  const hw = Number(navigator.hardwareConcurrency || 4);
  const saveData = !!(navigator.connection && navigator.connection.saveData);
  const autoConcurrency = saveData
    ? 1
    : Math.max(1, Math.min(STREET_DEFAULT_CONCURRENCY, Math.floor(hw / 2)));
  const concurrency = Math.max(1, Number(options.concurrency || autoConcurrency));
  const batchGapMs = Math.max(0, Number(options.batchGapMs ?? STREET_BATCH_GAP_MS));

  for (let i = 0; i < tiles.length; i += concurrency) {
    const batch = tiles.slice(i, i + concurrency);
    await Promise.all(batch.map(async (tile, idx) => {
      const tileIndex = i + idx;
      try {
        const elements = await fetchStreetElementsForTile(tile);
        elements.forEach(e => {
          if (!e || e.type !== "way" || !e.tags?.highway || !Array.isArray(e.geom) || e.geom.length < 2) return;
          mergedById.set(e.id, e);
        });
      } catch (tileErr) {
        tileErrors.push(`tile ${tileIndex + 1}/${tiles.length}: ${tileErr?.message || tileErr}`);
      }
    }));
    updateStreetLoadStatus(`Loading street segments... ${mergedById.size} fetched`);
    if (i + concurrency < tiles.length && batchGapMs > 0) {
      await sleep(batchGapMs);
    }
  }

  return { mergedById, tileErrors, tilesCount: tiles.length };
}

function upsertStreetElement(e) {
  const id = Number(e?.id);
  if (!Number.isFinite(id) || !Array.isArray(e.geom) || e.geom.length < 2) return false;
  const latlngs = e.geom.map(g => [g.lat, g.lon]);
  const tags = normalizeLocalStreetTags(e.tags || {});
  const incomingRow = {
    id,
    name: tags.name,
    highway: tags.highway,
    ref: tags.ref,
    maxspeed: tags.maxspeed,
    lanes: tags.lanes,
    surface: tags.surface,
    oneway: tags.oneway
  };
  const existing = streetAttributeById.get(id);
  if (existing?.layer) {
    existing.row = mergeStreetAttributeRows(existing.row, incomingRow);
    if (typeof existing.layer.setLatLngs === "function") existing.layer.setLatLngs(latlngs);
    setStreetSegmentStyle(existing, streetAttributeSelectedIds.has(id));
    return false;
  }
  const row = mergeStreetAttributeRows(null, incomingRow);
  const baseStyle = getStreetSegmentBaseStyle({ row });
  const layer = L.polyline(latlngs, {
    color: baseStyle.color,
    weight: baseStyle.weight,
    opacity: baseStyle.opacity,
    renderer: canvasRenderer,
    smoothFactor: 1.2
  });
  layer.on("click", () => {
    if (!isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) return;
    const attributePanel = document.getElementById("attributeTablePanel");
    if (attributeTableMode !== "streets") {
      setAttributeTableMode("streets");
    }
    if (attributePanel?.classList.contains("closed")) {
      openAttributePanel();
    }
    toggleStreetSegmentSelection(id, null, true);
  });
  streetAttributeLayerGroup.addLayer(layer);
  streetAttributeById.set(id, { id, row, layer });
  return true;
}

async function loadStreetAttributesForCurrentView(boundsOverride = null, polygonLayer = null) {
  if (streetLoadInFlight) {
    pendingStreetReload = true;
    updateStreetLoadStatus("Street attributes are already loading...", false);
    return;
  }
  const b = boundsOverride || map.getBounds();
  const south = b.getSouth();
  const west = b.getWest();
  const north = b.getNorth();
  const east = b.getEast();
  const spanLat = Math.abs(north - south);
  const spanLng = Math.abs(east - west);
  const isPolygonScopedLoad = !!polygonLayer;
  const useLocalSource = shouldUseLocalStreetSource();
  if (!useLocalSource && !isPolygonScopedLoad && (spanLat > 1.6 || spanLng > 1.6)) {
    updateStreetLoadStatus("Zoom in more before loading street segments.", true);
    return;
  }
  if (!useLocalSource) {
    const now = Date.now();
    if (now - lastStreetLoadAt < 900) {
      const waitMs = 900 - (now - lastStreetLoadAt);
      updateStreetLoadStatus(`Please wait ${Math.ceil(waitMs / 1000)}s before reloading (provider rate limit).`, true);
      return;
    }
    lastStreetLoadAt = now;
  }
  streetLoadInFlight = true;
  setStreetLoadBarVisible(true);

  try {
    if (useLocalSource) {
      await checkLocalStreetBackendAvailability(false);
      if (localStreetBackendState.available && localStreetBackendState.hasIndex) {
        updateStreetLoadStatus("Loading street segments from local backend...");
        const backendResult = await loadStreetAttributesFromLocalBackend(b, polygonLayer);
        const { addedCount, totalCount, candidateCount, knownCount } = backendResult;
        if (!candidateCount) {
          updateStreetLoadStatus("No backend street segments found in this area.", true);
        } else if (!totalCount) {
          updateStreetLoadStatus("Backend streets loaded, but none intersect this polygon.", true);
        } else if (!knownCount) {
          updateStreetLoadStatus("Street geometry loaded, but local index attributes are empty. Re-run indexer with the latest converter to populate road fields.", true);
        } else {
          updateStreetLoadStatus(`Loaded ${addedCount} backend street segments (${totalCount} total on map).`, false);
        }
        return;
      }

      const chunkTargetBounds = polygonLayer?.getBounds?.() || b;
      if (localStreetSourceState.chunkMode) {
        updateStreetLoadStatus("Preparing local street chunk for this region...");
        const chunkReady = await ensureLocalStreetChunkCoversBounds(chunkTargetBounds);
        if (!chunkReady) {
          updateStreetLoadStatus("Could not build local chunk for this region. Reload the local streets file.", true);
          return;
        }
      }
      updateStreetLoadStatus("Loading street segments from local file...");
      const localResult = await loadStreetAttributesFromLocalDataset(b, polygonLayer);
      const { addedCount, totalCount, candidateCount, knownCount } = localResult;
      if (!candidateCount) {
        if (localStreetSourceState.chunkMode) {
          updateStreetLoadStatus("Chunk loaded, but no local street segments were found in this area.", true);
        } else {
          updateStreetLoadStatus("No local street segments found in this area.", true);
        }
      } else if (!totalCount) {
        updateStreetLoadStatus("Local streets loaded, but none intersect this polygon.", true);
      } else if (!knownCount) {
        updateStreetLoadStatus("Street geometry loaded, but local file attributes are empty. Use a roads source that includes DBF properties.", true);
      } else {
        updateStreetLoadStatus(`Loaded ${addedCount} local street segments (${totalCount} total on map).`, false);
      }
      return;
    }

    updateStreetLoadStatus("Loading street segments...");
    let mergedById = new Map();
    let tileErrors = [];
    let tilesCount = 0;
    let retriedWithExpandedBounds = false;
    let addedCount = 0;
    let streamedToMap = false;

    const applyElementsToMap = (elementsMap) => {
      let chunkAdded = 0;
      elementsMap.forEach(e => {
        if (polygonLayer && !streetElementIntersectsPolygon(e, polygonLayer)) return;
        if (upsertStreetElement(e)) chunkAdded += 1;
      });
      if (chunkAdded > 0) {
        streetAttributesRows = [...streetAttributeById.values()].map(v => v.row);
        syncStreetNetworkOverlay();
      }
      return chunkAdded;
    };

    if (isPolygonScopedLoad && (spanLat > STREET_MAX_LOAD_SPAN || spanLng > STREET_MAX_LOAD_SPAN)) {
      const focus = map?.getCenter?.() || null;
      const rawChunks = buildStreetBoundsChunks(south, west, north, east, STREET_POLYGON_CHUNK_SPAN);
      const chunks = sortStreetChunksByFocus(rawChunks, focus);
      updateStreetLoadStatus(`Loading large polygon in ${chunks.length} chunks (nearest first)...`);
      streamedToMap = true;

      for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        updateStreetLoadStatus(`Loading street segments... chunk ${i + 1}/${chunks.length}`);
        const chunkResult = await collectStreetElementsForBounds(
          chunk.south,
          chunk.west,
          chunk.north,
          chunk.east,
          {
            concurrency: STREET_CHUNK_CONCURRENCY,
            batchGapMs: STREET_BATCH_GAP_MS,
            singleTile: true
          }
        );
        chunkResult.mergedById.forEach((value, key) => mergedById.set(key, value));
        tileErrors = tileErrors.concat(chunkResult.tileErrors);
        tilesCount += chunkResult.tilesCount;
        addedCount += applyElementsToMap(chunkResult.mergedById);
        updateStreetLoadStatus(
          `Loading street segments... chunk ${i + 1}/${chunks.length} (${streetAttributeById.size} on map)`
        );
        if (attributeTableMode === "streets") {
          renderAttributeTable();
        }
        if (i < chunks.length - 1) {
          await sleep(STREET_CHUNK_GAP_MS);
        }
      }
    } else {
      const first = await collectStreetElementsForBounds(south, west, north, east, {
        concurrency: STREET_DEFAULT_CONCURRENCY,
        batchGapMs: STREET_BATCH_GAP_MS
      });
      mergedById = first.mergedById;
      tileErrors = first.tileErrors;
      tilesCount = first.tilesCount;

      // Retry once with a slightly larger bbox to handle sparse/tile-edge results.
      if (!mergedById.size) {
        retriedWithExpandedBounds = true;
        updateStreetLoadStatus("No segments found. Retrying with expanded area...");
        const padLat = Math.max(spanLat * 0.16, 0.0035);
        const padLng = Math.max(spanLng * 0.16, 0.0035);
        const retryResult = await collectStreetElementsForBounds(
          south - padLat,
          west - padLng,
          north + padLat,
          east + padLng,
          { concurrency: STREET_DEFAULT_CONCURRENCY, batchGapMs: STREET_BATCH_GAP_MS }
        );
        mergedById = retryResult.mergedById;
        tileErrors = tileErrors.concat(retryResult.tileErrors);
        tilesCount += retryResult.tilesCount;
      }
    }

    if (!mergedById.size) {
      const detail = tileErrors.length
        ? ` No segments returned. ${tileErrors.length}/${tilesCount} tiles failed.`
        : " No segments returned for this view after retry.";
      updateStreetLoadStatus(`Street load completed with no data.${detail}`, true);
      return;
    }

    if (!streamedToMap) {
      [...mergedById.values()].forEach(e => {
        if (polygonLayer && !streetElementIntersectsPolygon(e, polygonLayer)) return;
        if (upsertStreetElement(e)) addedCount += 1;
      });
    }
    streetAttributesRows = [...streetAttributeById.values()].map(v => v.row);
    syncStreetNetworkOverlay();
    if (attributeTableMode === "streets") renderAttributeTable();
    applyStreetSelectionStyles();
    const totalCount = streetAttributeById.size;
    if (!totalCount) {
      updateStreetLoadStatus("No street segments returned for current view.", true);
    } else if (tileErrors.length) {
      updateStreetLoadStatus(`Loaded ${addedCount} new segments (${totalCount} total). Some sources throttled (${tileErrors.length} failures).`, false);
      console.warn(`Street attributes loaded with partial tile failures (${tileErrors.length}).`, tileErrors);
    } else if (retriedWithExpandedBounds) {
      updateStreetLoadStatus(`Loaded ${addedCount} new street segments (${totalCount} total, expanded area retry).`, false);
    } else {
      updateStreetLoadStatus(`Loaded ${addedCount} new street segments (${totalCount} total).`, false);
    }
  } catch (err) {
    console.error("STREET ATTRIBUTES LOAD ERROR:", err);
    const msg = String(err?.message || "");
    const rateLimited = msg.includes("429") || msg.includes("509") || msg.toLowerCase().includes("cooling down");
    if (rateLimited) {
      updateStreetLoadStatus("Providers are rate-limiting. Try a smaller polygon or wait 1-2 minutes.", true);
    } else {
      updateStreetLoadStatus("Unable to load street segments from providers.", true);
    }
    alert(`Unable to load street attributes.\n\nData providers are currently busy or rate-limited.\nTry zooming in more (smaller area) and run Street Attributes again.\n\nDetails: ${err.message}`);
  } finally {
    streetLoadInFlight = false;
    setStreetLoadBarVisible(false);
    pendingStreetReload = false;
    updateStreetSetupGuide();
    if (typeof window.__refreshStreetNetworkManagerUi === "function") {
      window.__refreshStreetNetworkManagerUi();
    }
  }
}

function zoomToSelectedStreetSegments() {
  if (!streetAttributeSelectedIds.size) {
    alert("No selected street segments.");
    return;
  }
  const bounds = L.latLngBounds();
  streetAttributeSelectedIds.forEach(id => {
    const entry = streetAttributeById.get(id);
    if (entry?.layer) bounds.extend(entry.layer.getBounds());
  });
  if (!bounds.isValid()) return;
  map.fitBounds(bounds.pad(0.18));
}

function flattenStreetLayerLatLngs(raw, out = []) {
  if (!raw) return out;
  if (Array.isArray(raw)) {
    if (raw.length && raw[0] && typeof raw[0].lat === "number" && typeof raw[0].lng === "number") {
      raw.forEach(pt => {
        const lat = Number(pt?.lat);
        const lng = Number(pt?.lng);
        if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
        out.push({ lat, lon: lng });
      });
      return out;
    }
    raw.forEach(part => flattenStreetLayerLatLngs(part, out));
  }
  return out;
}

function streetEntryIntersectsSelectionPolygon(entry, polygonLayer) {
  const layer = entry?.layer;
  if (!layer || !polygonLayer) return false;
  const geom = flattenStreetLayerLatLngs(layer.getLatLngs?.(), []);
  if (!geom.length) return false;
  return streetElementIntersectsPolygon({ geom }, polygonLayer);
}

function replaceStreetSelection(nextSelectedIds) {
  streetAttributeSelectedIds.clear();
  (nextSelectedIds || []).forEach(id => {
    const wayId = Number(id);
    if (!Number.isFinite(wayId)) return;
    if (!streetAttributeById.has(wayId)) return;
    streetAttributeSelectedIds.add(wayId);
  });
}

const MAP_SELECTION_DRAW_MODE_KEY = "mapSelectionDrawMode";

function normalizeMapSelectionDrawMode(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "add" || normalized === "subtract" || normalized === "intersect") return normalized;
  return "replace";
}

function getMapSelectionModeLabel(modeValue = "replace") {
  const mode = normalizeMapSelectionDrawMode(modeValue);
  if (mode === "add") return "Add";
  if (mode === "subtract") return "Subtract";
  if (mode === "intersect") return "Intersect";
  return "Replace";
}

let mapSelectionDrawMode = normalizeMapSelectionDrawMode(storageGet(MAP_SELECTION_DRAW_MODE_KEY) || "replace");

function setMapSelectionDrawMode(modeValue, persist = true) {
  mapSelectionDrawMode = normalizeMapSelectionDrawMode(modeValue);
  if (persist) {
    storageSet(MAP_SELECTION_DRAW_MODE_KEY, mapSelectionDrawMode);
  }
  if (typeof window.__refreshSelectionToolsUi === "function") {
    window.__refreshSelectionToolsUi();
  }
}

function areNumericIdSetsEqual(a, b) {
  if (!a || !b) return false;
  if (a.size !== b.size) return false;
  for (const id of a) {
    if (!b.has(id)) return false;
  }
  return true;
}

function combineSelectionIds(currentIds, candidateIds, modeValue = mapSelectionDrawMode) {
  const mode = normalizeMapSelectionDrawMode(modeValue);
  const current = new Set(
    Array.from(currentIds || [])
      .map(v => Number(v))
      .filter(v => Number.isFinite(v))
  );
  const candidates = new Set(
    Array.from(candidateIds || [])
      .map(v => Number(v))
      .filter(v => Number.isFinite(v))
  );

  if (mode === "add") {
    candidates.forEach(id => current.add(id));
    return current;
  }
  if (mode === "subtract") {
    candidates.forEach(id => current.delete(id));
    return current;
  }
  if (mode === "intersect") {
    const next = new Set();
    current.forEach(id => {
      if (candidates.has(id)) next.add(id);
    });
    return next;
  }
  return candidates;
}

function getVisibleSelectableRecordRowIds() {
  const out = new Set();
  Object.entries(routeDayGroups).forEach(([key, group]) => {
    if (!isLayerManagerEntrySelectable(key)) return;
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(marker => {
      if (!map.hasLayer(marker)) return;
      const rowId = Number(marker?._rowId);
      if (!Number.isFinite(rowId)) return;
      out.add(rowId);
    });
  });
  return out;
}

function getSelectableRecordRowIds() {
  const out = new Set();
  Object.entries(routeDayGroups).forEach(([key, group]) => {
    if (!isLayerManagerEntrySelectable(key)) return;
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(marker => {
      const rowId = Number(marker?._rowId);
      if (!Number.isFinite(rowId)) return;
      out.add(rowId);
    });
  });
  return out;
}

function getVisibleSelectableStreetIds() {
  const out = new Set();
  if (!isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) return out;
  streetAttributeById.forEach((entry, id) => {
    if (!entry?.layer) return;
    if (!map.hasLayer(entry.layer)) return;
    out.add(id);
  });
  return out;
}

function getRecordRowIdsInCurrentView() {
  const out = new Set();
  const bounds = map.getBounds();
  Object.entries(routeDayGroups).forEach(([key, group]) => {
    if (!isLayerManagerEntrySelectable(key)) return;
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(marker => {
      if (!map.hasLayer(marker)) return;
      const base = marker?._base;
      if (!base) return;
      const rowId = Number(marker?._rowId);
      if (!Number.isFinite(rowId)) return;
      const latlng = L.latLng(base.lat, base.lon);
      if (!bounds.contains(latlng)) return;
      out.add(rowId);
    });
  });
  return out;
}

function getStreetIdsInCurrentView() {
  const out = new Set();
  if (!isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) return out;
  const bounds = map.getBounds();
  streetAttributeById.forEach((entry, id) => {
    const layer = entry?.layer;
    if (!layer) return;
    if (!map.hasLayer(layer)) return;
    const layerBounds = layer.getBounds?.();
    if (!layerBounds?.isValid?.()) return;
    if (!bounds.intersects(layerBounds)) return;
    out.add(id);
  });
  return out;
}

function applySelectionIdsToActiveMode(candidateIds, modeValue = mapSelectionDrawMode) {
  const mode = normalizeMapSelectionDrawMode(modeValue);
  if (attributeTableMode === "streets") {
    if (!isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) {
      return streetAttributeSelectedIds.size;
    }
    const next = combineSelectionIds(streetAttributeSelectedIds, candidateIds, mode);
    replaceStreetSelection(next);
    applyStreetSelectionStyles();
    syncSelectedStopsHeaderCount(streetAttributeSelectedIds.size);
    refreshAttributeStatus();
    if (attributeTableMode === "streets") renderAttributeTable();
    return streetAttributeSelectedIds.size;
  }

  const selectableIds = getSelectableRecordRowIds();
  const base = new Set(
    [...attributeState.selectedRowIds].filter(id => selectableIds.has(id))
  );
  const next = combineSelectionIds(base, candidateIds, mode);
  attributeState.selectedRowIds = next;
  applyAttributeSelectionStyles();
  syncSelectedStopsHeaderCount(next.size);
  refreshAttributeStatus();
  if (attributeTableMode === "records") renderAttributeTable();
  return next.size;
}

function invertVisibleSelectionInActiveMode() {
  if (attributeTableMode === "streets") {
    if (!isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) {
      return streetAttributeSelectedIds.size;
    }
    const next = new Set(streetAttributeSelectedIds);
    const visibleIds = getVisibleSelectableStreetIds();
    visibleIds.forEach(id => {
      if (next.has(id)) next.delete(id);
      else next.add(id);
    });
    replaceStreetSelection(next);
    applyStreetSelectionStyles();
    syncSelectedStopsHeaderCount(streetAttributeSelectedIds.size);
    refreshAttributeStatus();
    if (attributeTableMode === "streets") renderAttributeTable();
    return streetAttributeSelectedIds.size;
  }

  const next = new Set(attributeState.selectedRowIds);
  const visibleIds = getVisibleSelectableRecordRowIds();
  visibleIds.forEach(id => {
    if (next.has(id)) next.delete(id);
    else next.add(id);
  });
  attributeState.selectedRowIds = next;
  applyAttributeSelectionStyles();
  syncSelectedStopsHeaderCount(next.size);
  refreshAttributeStatus();
  if (attributeTableMode === "records") renderAttributeTable();
  return next.size;
}

// ================= POLYGON SELECT =================


// when polygon created
// ================= POLYGON SELECT =================
let drawnLayer = new L.FeatureGroup();
map.addLayer(drawnLayer);

const drawControl = new L.Control.Draw({
  draw: {
    polygon: true,
    rectangle: true,
    circle: false,
    marker: false,
    polyline: false,
    circlemarker: false
  },
  edit: { featureGroup: drawnLayer }
});

map.addControl(drawControl);

function startSelectionDrawTool(shapeType) {
  selectedLayerKey = null;
  const drawToolbar = drawControl?._toolbars?.draw;
  const mode = drawToolbar?._modes?.[shapeType];
  if (mode?.handler && typeof mode.handler.enable === "function") {
    mode.handler.enable();
    return true;
  }
  const fallbackBtn = document.querySelector(`.leaflet-draw-draw-${shapeType}`);
  if (fallbackBtn) {
    fallbackBtn.click();
    return true;
  }
  return false;
}

function clearDrawnSelectionGeometry() {
  selectedLayerKey = null;
  drawnLayer.clearLayers();
  updateSelectionCount();
  updateUndoButtonState();
}

// ===== SELECTION COUNT FUNCTION (GLOBAL & CORRECT) =====
function updateSelectionCount() {
  const polygon = drawnLayer.getLayers()[0] || null;
  const polygonBounds = polygon?.getBounds?.() || null;

  if (attributeTableMode === "streets") {
    const streetSelectable = isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY);
    if (!streetSelectable) {
      const hadSelection = streetAttributeSelectedIds.size > 0;
      if (hadSelection) {
        streetAttributeSelectedIds.clear();
        applyStreetSelectionStyles();
        renderAttributeTable();
      }
      syncSelectedStopsHeaderCount(0);
      refreshAttributeStatus();
      return;
    }

    let nextSelectedStreetIds = new Set(streetAttributeSelectedIds);
    if (polygon) {
      const candidateStreetIds = new Set();
      streetAttributeById.forEach((entry, id) => {
        if (!entry?.layer) return;
        if (!map.hasLayer(entry.layer)) return;
        if (!streetEntryIntersectsSelectionPolygon(entry, polygon)) return;
        candidateStreetIds.add(id);
      });
      nextSelectedStreetIds = combineSelectionIds(nextSelectedStreetIds, candidateStreetIds, mapSelectionDrawMode);
    }
    nextSelectedStreetIds = new Set(
      [...nextSelectedStreetIds].filter(id => streetAttributeById.has(id))
    );

    const prevStreet = streetAttributeSelectedIds;
    const changed = !areNumericIdSetsEqual(prevStreet, nextSelectedStreetIds);

    if (changed) {
      replaceStreetSelection(nextSelectedStreetIds);
      applyStreetSelectionStyles();
      if (attributeTableMode === "streets") renderAttributeTable();
    }

    syncSelectedStopsHeaderCount(streetAttributeSelectedIds.size);
    refreshAttributeStatus();
    return;
  }

  if (selectedLayerKey && !isLayerManagerEntrySelectable(selectedLayerKey)) {
    selectedLayerKey = null;
  }
  const isLayerSelectMode = !!selectedLayerKey;
  const selectableRowIds = new Set();
  let nextSelectedRowIds = new Set(attributeState.selectedRowIds);

  nextSelectedRowIds = new Set(
    [...nextSelectedRowIds].filter(id => Number.isFinite(id))
  );

  const layerCandidateIds = new Set();
  const polygonCandidateIds = new Set();

  Object.entries(routeDayGroups).forEach(([key, group]) => {
    const layerSelectable = isLayerManagerEntrySelectable(key);
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(marker => {
      const rowId = Number(marker?._rowId);
      if (!layerSelectable || !Number.isFinite(rowId)) return;
      selectableRowIds.add(rowId);
      if (isLayerSelectMode && key === selectedLayerKey && map.hasLayer(marker)) {
        layerCandidateIds.add(rowId);
        return;
      }
      if (!isLayerSelectMode && polygonBounds) {
        const base = marker?._base;
        if (!base || !map.hasLayer(marker)) return;
        const latlng = L.latLng(base.lat, base.lon);
        if (polygonBounds.contains(latlng)) polygonCandidateIds.add(rowId);
      }
    });
  });

  nextSelectedRowIds = new Set(
    [...nextSelectedRowIds].filter(id => selectableRowIds.has(id))
  );

  if (isLayerSelectMode) {
    nextSelectedRowIds = layerCandidateIds;
  } else if (polygonBounds) {
    nextSelectedRowIds = combineSelectionIds(nextSelectedRowIds, polygonCandidateIds, mapSelectionDrawMode);
  }

  Object.entries(routeDayGroups).forEach(([key, group]) => {
    const layerSelectable = isLayerManagerEntrySelectable(key);
    const sym = symbolMap[key] || getSymbol(key);
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(marker => {
      const rowId = Number(marker?._rowId);
      const selected = layerSelectable && Number.isFinite(rowId) && nextSelectedRowIds.has(rowId);
      const color = selected ? "#ffff00" : sym.color;
      marker.setStyle?.({ color, fillColor: color });
    });
  });

  const prev = attributeState.selectedRowIds;
  const changed = !areNumericIdSetsEqual(prev, nextSelectedRowIds);
  attributeState.selectedRowIds = nextSelectedRowIds;
  syncSelectedStopsHeaderCount(nextSelectedRowIds.size);
  refreshAttributeStatus();
  if (changed) renderAttributeTable();
}


// ===== COMPLETE SELECTED STOPS =====


  


// ===== WHEN POLYGON IS DRAWN =====
map.on(L.Draw.Event.CREATED, e => {
  const streetToggle = document.getElementById("useLocalStreetSource");
  if (streetPolygonLoadPending && streetToggle?.checked) {
    streetPolygonLoadPending = false;
    closeStreetSegmentsPromptModal();
    lastStreetLoadPolygonSnapshot = encodeStreetPolygonLatLngs(e.layer.getLatLngs());
    streetLoadPolygonLayerGroup.clearLayers();
    streetLoadPolygonLayerGroup.addLayer(e.layer);
    // Keep polygon geometry for this load, but hide the drawn shape from the map.
    setTimeout(() => {
      streetLoadPolygonLayerGroup.clearLayers();
    }, 0);
    // Replace old streets only after user confirms a new polygon.
    streetAttributeLayerGroup.clearLayers();
    streetAttributeById.clear();
    streetAttributesRows = [];
    streetAttributeSelectedIds.clear();
    if (attributeTableMode === "streets") renderAttributeTable();
    updateStreetLoadStatus("Polygon captured. Loading street segments...");
    loadStreetAttributesForCurrentView(e.layer.getBounds(), e.layer).catch(err => {
      console.error("Street polygon load failed:", err);
    });
    return;
  }
  selectedLayerKey = null;
  drawnLayer.clearLayers();
  drawnLayer.addLayer(e.layer);
  updateSelectionCount();
  updateUndoButtonState();   // 🔥 ADD THIS
});

map.on(L.Draw.Event.DRAWSTOP, () => {
  if (!streetPolygonLoadPending) return;
  // Leaflet draw can emit DRAWSTOP before CREATED on some paths.
  // Defer cancel handling so CREATED can clear this flag first.
  setTimeout(() => {
    if (!streetPolygonLoadPending) return;
    if (streetLoadInFlight || streetLoadPolygonLayerGroup.getLayers().length > 0) {
      streetPolygonLoadPending = false;
      return;
    }
    streetPolygonLoadPending = false;
    const streetToggle = document.getElementById("useLocalStreetSource");
    closeStreetSegmentsPromptModal();
    if (streetToggle) {
      streetToggle.checked = false;
      storageSet("streetSegmentsVisible", "off");
    }
    syncStreetNetworkOverlay();
    updateLocalStreetSourceStatus();
  }, 0);
});

// Default map
baseMaps.streets.addTo(map);

// Dropdown to switch map type
document.getElementById("baseMapSelect").addEventListener("change", e => {
  Object.values(baseMaps).forEach(l => map.removeLayer(l));
  map.removeLayer(satelliteLabelsLayer);

  const selected = e.target.value;
  baseMaps[selected].addTo(map);

  if (selected === "satellite" && map.getZoom() >= 15) {
    satelliteLabelsLayer.addTo(map);
  }
  syncStreetNetworkOverlay();
});


// ================= MAP SYMBOL SETTINGS =================
const shapes = ["circle","square","triangle","diamond"];
const MARKER_SIZE_STEPS = [
  [7, 1.5],
  [9, 2.2],
  [11, 3.0],
  [13, 3.6],
  [15, 4.4],
  [Infinity, 5.2]
];

const symbolMap = {};        // stores symbol for each route/day combo
const routeDayGroups = {};   // stores map markers grouped by route/day
const LAYER_MANAGER_ORDER_STORAGE_KEY = "layerManagerOrderTop";
const LAYER_MANAGER_SELECTABLE_STORAGE_KEY = "layerManagerSelectable";
const LAYER_MANAGER_STREET_KEY = "street-network";
let layerManagerOrderTop = [];
let layerManagerSelectableState = {};
// ===== DELIVERED STOPS LAYER =====

function layerManagerDaySortRank(value) {
  const v = String(value ?? "").trim().toLowerCase();
  const byName = {
    "1": 1, mon: 1, monday: 1,
    "2": 2, tue: 2, tues: 2, tuesday: 2,
    "3": 3, wed: 3, wednesday: 3,
    "4": 4, thu: 4, thur: 4, thurs: 4, thursday: 4,
    "5": 5, fri: 5, friday: 5,
    "6": 6, sat: 6, saturday: 6,
    "7": 7, sun: 7, sunday: 7,
    delivered: 99
  };
  if (Object.prototype.hasOwnProperty.call(byName, v)) return byName[v];
  const n = Number(v);
  return Number.isFinite(n) ? n : 98;
}

function getSortedRouteDayKeysForLayerManager() {
  return Object.keys(routeDayGroups).sort((aKey, bKey) => {
    const [aRoute = "", aDay = ""] = String(aKey).split("|");
    const [bRoute = "", bDay = ""] = String(bKey).split("|");
    const routeCmp = aRoute.localeCompare(bRoute, undefined, { numeric: true, sensitivity: "base" });
    if (routeCmp !== 0) return routeCmp;
    return layerManagerDaySortRank(aDay) - layerManagerDaySortRank(bDay);
  });
}

function getLayerManagerDefaultOrder() {
  // Top-to-bottom order: route/day layers on top, street network near bottom by default.
  return [...getSortedRouteDayKeysForLayerManager(), LAYER_MANAGER_STREET_KEY];
}

function isRouteDayLayerManagerEntry(entryId) {
  return entryId !== LAYER_MANAGER_STREET_KEY && Object.prototype.hasOwnProperty.call(routeDayGroups, entryId);
}

function buildLayerManagerOrderWithRouteGrouping(orderCandidate = []) {
  const source = Array.isArray(orderCandidate) ? orderCandidate.map(v => String(v || "")).filter(Boolean) : [];
  const streetOnTop = source.indexOf(LAYER_MANAGER_STREET_KEY) === 0;
  const routeKeys = getSortedRouteDayKeysForLayerManager();

  const routeOrder = [];
  source.forEach(key => {
    if (!isRouteDayLayerManagerEntry(key)) return;
    if (routeOrder.includes(key)) return;
    routeOrder.push(key);
  });
  routeKeys.forEach(key => {
    if (!routeOrder.includes(key)) routeOrder.push(key);
  });

  if (!routeOrder.length) {
    return [LAYER_MANAGER_STREET_KEY];
  }
  return streetOnTop
    ? [LAYER_MANAGER_STREET_KEY, ...routeOrder]
    : [...routeOrder, LAYER_MANAGER_STREET_KEY];
}

function loadStoredLayerManagerOrder() {
  const raw = storageGet(LAYER_MANAGER_ORDER_STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(v => String(v || "")).filter(Boolean);
  } catch (_) {
    return [];
  }
}

function saveLayerManagerOrder(order) {
  storageSet(LAYER_MANAGER_ORDER_STORAGE_KEY, JSON.stringify(order || []));
}

function loadStoredLayerManagerSelectableState() {
  const raw = storageGet(LAYER_MANAGER_SELECTABLE_STORAGE_KEY);
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return {};
    const next = {};
    Object.keys(parsed).forEach(key => {
      const k = String(key || "").trim();
      if (!k) return;
      next[k] = !!parsed[key];
    });
    return next;
  } catch (_) {
    return {};
  }
}

function saveLayerManagerSelectableState() {
  storageSet(LAYER_MANAGER_SELECTABLE_STORAGE_KEY, JSON.stringify(layerManagerSelectableState || {}));
}

function ensureLayerManagerOrder() {
  if (!layerManagerOrderTop.length) {
    layerManagerOrderTop = loadStoredLayerManagerOrder();
  }
  const next = buildLayerManagerOrderWithRouteGrouping(layerManagerOrderTop.length ? layerManagerOrderTop : getLayerManagerDefaultOrder());

  const changed =
    next.length !== layerManagerOrderTop.length ||
    next.some((key, idx) => key !== layerManagerOrderTop[idx]);
  if (changed) {
    layerManagerOrderTop = next;
    saveLayerManagerOrder(layerManagerOrderTop);
  }
  return layerManagerOrderTop;
}

function ensureLayerManagerSelectableState() {
  const keys = ensureLayerManagerOrder();
  if (!layerManagerSelectableState || typeof layerManagerSelectableState !== "object" || Array.isArray(layerManagerSelectableState)) {
    layerManagerSelectableState = {};
  }
  if (!Object.keys(layerManagerSelectableState).length) {
    layerManagerSelectableState = loadStoredLayerManagerSelectableState();
  }

  let changed = false;
  keys.forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(layerManagerSelectableState, key)) {
      layerManagerSelectableState[key] = true;
      changed = true;
    }
  });
  Object.keys(layerManagerSelectableState).forEach(key => {
    if (keys.includes(key)) return;
    delete layerManagerSelectableState[key];
    changed = true;
  });
  if (changed) saveLayerManagerSelectableState();
  return layerManagerSelectableState;
}

function isLayerManagerEntrySelectable(entryId) {
  const state = ensureLayerManagerSelectableState();
  if (!Object.prototype.hasOwnProperty.call(state, entryId)) return true;
  return !!state[entryId];
}

function setLayerManagerEntrySelectable(entryId, selectable) {
  const state = ensureLayerManagerSelectableState();
  const nextValue = !!selectable;
  if (state[entryId] === nextValue) return false;
  state[entryId] = nextValue;
  saveLayerManagerSelectableState();

  if (!nextValue) {
    if (entryId === LAYER_MANAGER_STREET_KEY) {
      streetAttributeSelectedIds.clear();
      applyStreetSelectionStyles();
      if (attributeTableMode === "streets") renderAttributeTable();
    } else if (selectedLayerKey === entryId) {
      selectedLayerKey = null;
    }
  }
  refreshRouteDayLayerSelectButtons();
  return true;
}

function moveLayerManagerEntryBefore(entryId, targetId) {
  const order = [...ensureLayerManagerOrder()];
  const streetOnTop = order[0] === LAYER_MANAGER_STREET_KEY;
  const routeOrder = order.filter(isRouteDayLayerManagerEntry);

  if (entryId === LAYER_MANAGER_STREET_KEY && isRouteDayLayerManagerEntry(targetId)) {
    if (streetOnTop) return false;
    layerManagerOrderTop = [LAYER_MANAGER_STREET_KEY, ...routeOrder];
    saveLayerManagerOrder(layerManagerOrderTop);
    return true;
  }

  if (isRouteDayLayerManagerEntry(entryId) && targetId === LAYER_MANAGER_STREET_KEY) {
    if (!streetOnTop) return false;
    layerManagerOrderTop = [...routeOrder, LAYER_MANAGER_STREET_KEY];
    saveLayerManagerOrder(layerManagerOrderTop);
    return true;
  }

  if (!isRouteDayLayerManagerEntry(entryId) || !isRouteDayLayerManagerEntry(targetId)) return false;
  if (entryId === targetId) return false;

  const fromIndex = routeOrder.indexOf(entryId);
  const toIndex = routeOrder.indexOf(targetId);
  if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return false;
  routeOrder.splice(fromIndex, 1);
  const insertAt = fromIndex < toIndex ? toIndex - 1 : toIndex;
  routeOrder.splice(insertAt, 0, entryId);
  layerManagerOrderTop = streetOnTop
    ? [LAYER_MANAGER_STREET_KEY, ...routeOrder]
    : [...routeOrder, LAYER_MANAGER_STREET_KEY];
  saveLayerManagerOrder(layerManagerOrderTop);
  return true;
}

function moveLayerManagerEntryByOffset(entryId, offset) {
  const order = [...ensureLayerManagerOrder()];
  const streetOnTop = order[0] === LAYER_MANAGER_STREET_KEY;
  const routeOrder = order.filter(isRouteDayLayerManagerEntry);
  const step = Number(offset || 0);
  if (!step) return false;

  if (entryId === LAYER_MANAGER_STREET_KEY) {
    if (!routeOrder.length) return false;
    const moveUp = step < 0;
    if (moveUp && streetOnTop) return false;
    if (!moveUp && !streetOnTop) return false;
    layerManagerOrderTop = moveUp
      ? [LAYER_MANAGER_STREET_KEY, ...routeOrder]
      : [...routeOrder, LAYER_MANAGER_STREET_KEY];
    saveLayerManagerOrder(layerManagerOrderTop);
    return true;
  }

  if (!isRouteDayLayerManagerEntry(entryId)) return false;
  const fromIndex = routeOrder.indexOf(entryId);
  if (fromIndex < 0) return false;
  const toIndex = Math.max(0, Math.min(routeOrder.length - 1, fromIndex + step));
  if (toIndex === fromIndex) return false;
  routeOrder.splice(fromIndex, 1);
  routeOrder.splice(toIndex, 0, entryId);
  layerManagerOrderTop = streetOnTop
    ? [LAYER_MANAGER_STREET_KEY, ...routeOrder]
    : [...routeOrder, LAYER_MANAGER_STREET_KEY];
  saveLayerManagerOrder(layerManagerOrderTop);
  return true;
}

function isRouteDayLayerVisibleOnMap(key) {
  const group = routeDayGroups[key];
  if (!group || !Array.isArray(group.layers)) return false;
  return group.layers.some(layer => map.hasLayer(layer));
}

function applyLayerManagerOrder() {
  const orderTop = ensureLayerManagerOrder();
  for (let i = orderTop.length - 1; i >= 0; i -= 1) {
    const entryId = orderTop[i];
    if (entryId === LAYER_MANAGER_STREET_KEY) {
      if (map.hasLayer(streetAttributeLayerGroup)) {
        streetAttributeLayerGroup.eachLayer(layer => {
          try { layer.bringToFront?.(); } catch (_) {}
        });
      }
      continue;
    }

    const group = routeDayGroups[entryId];
    if (!group || !Array.isArray(group.layers)) continue;
    group.layers.forEach(layer => {
      if (!map.hasLayer(layer)) return;
      try { layer.bringToFront?.(); } catch (_) {}
    });
  }

  // Keep temporary selection/polygon tools above data layers.
  try { drawnLayer?.bringToFront?.(); } catch (_) {}
  try { streetLoadPolygonLayerGroup?.bringToFront?.(); } catch (_) {}
}

function setRouteDayLayerVisibilityFromManager(key, visible) {
  const group = routeDayGroups[key];
  if (!group || !Array.isArray(group.layers)) return;
  const nextVisible = !!visible;
  layerVisibilityState[key] = nextVisible;

  const routeDayCheckbox = [...document.querySelectorAll("#routeDayLayers input[data-key]")]
    .find(node => node.dataset.key === key);
  if (routeDayCheckbox) routeDayCheckbox.checked = nextVisible;

  group.layers.forEach(layer => {
    if (nextVisible) map.addLayer(layer);
    else map.removeLayer(layer);
  });

  applyLayerManagerOrder();
  updateSelectionCount();
  updateStats();
}

function setStreetNetworkLayerVisibilityFromManager(visible) {
  const nextVisible = !!visible;
  const layerToggle = document.getElementById("streetNetworkLayerToggle");
  if (layerToggle) layerToggle.checked = nextVisible;
  storageSet(STREET_NETWORK_LAYER_VISIBLE_KEY, nextVisible ? "on" : "off");
  syncStreetNetworkOverlay();
  updateLocalStreetSourceStatus();
}

function refreshLayerManagerUiIfOpen() {
  if (typeof window.__refreshLayerManagerList === "function") {
    window.__refreshLayerManagerList();
  }
}

function normalizeDayToken(value) {
  const s = String(value ?? "").trim().toLowerCase();
  if (!s) return "";
  const dayMap = {
    "1": "1", "mon": "1", "monday": "1",
    "2": "2", "tue": "2", "tues": "2", "tuesday": "2",
    "3": "3", "wed": "3", "wednesday": "3",
    "4": "4", "thu": "4", "thur": "4", "thurs": "4", "thursday": "4",
    "5": "5", "fri": "5", "friday": "5",
    "6": "6", "sat": "6", "saturday": "6",
    "7": "7", "sun": "7", "sunday": "7"
  };
  return dayMap[s] || s;
}

function getLayerLatLng(layer) {
  if (!layer) return null;
  if (typeof layer.getLatLng === "function") return layer.getLatLng();
  if (typeof layer.getBounds === "function") return layer.getBounds().getCenter();
  if (layer._base && Number.isFinite(layer._base.lat) && Number.isFinite(layer._base.lon)) {
    return L.latLng(layer._base.lat, layer._base.lon);
  }
  return null;
}

// Used by the visualization popup window (window.opener) to focus the matching route+day on the map.
window.highlightRouteDayOnMap = function(routeValue, dayValue) {
  const routeToken = String(routeValue ?? "").trim();
  const dayToken = normalizeDayToken(dayValue);
  if (!routeToken || !dayToken) {
    return { ok: false, message: "Route/day value is missing." };
  }

  let matchingKey = null;
  Object.keys(routeDayGroups).forEach(key => {
    if (matchingKey) return;
    const [kRoute, kDay] = key.split("|");
    if (!kRoute || !kDay) return;
    if (String(kRoute).trim() !== routeToken) return;

    const keyDayToken = normalizeDayToken(kDay);
    if (keyDayToken === dayToken) matchingKey = key;
  });

  if (!matchingKey || !routeDayGroups[matchingKey]) {
    return { ok: false, message: `Could not find ${routeToken} | ${dayValue} on the map.` };
  }

  const group = routeDayGroups[matchingKey];
  const markers = group.layers || [];
  if (!markers.length) {
    return { ok: false, message: `No map points found for ${routeToken} | ${dayValue}.` };
  }

  layerVisibilityState[matchingKey] = true;
  const layerCheckbox = document.querySelector(`input[data-key="${matchingKey}"]`);
  if (layerCheckbox) layerCheckbox.checked = true;

  const bounds = L.latLngBounds();
  markers.forEach(marker => {
    map.addLayer(marker);
    const ll = getLayerLatLng(marker);
    if (ll) bounds.extend(ll);

    const sym = symbolMap[matchingKey] || { color: "#2f89df" };
    marker.setStyle?.({
      color: "#ffd54a",
      fillColor: "#ffd54a",
      fillOpacity: 1,
      opacity: 1,
      weight: 2
    });

    setTimeout(() => {
      marker.setStyle?.({
        color: sym.color,
        fillColor: sym.color,
        fillOpacity: 0.95,
        opacity: 1,
        weight: 1
      });
    }, 2200);
  });

  if (bounds.isValid()) {
    map.fitBounds(bounds, { padding: [40, 40], maxZoom: 15 });
  }

  return {
    ok: true,
    message: `Highlighted ${matchingKey} on the map.`,
    key: matchingKey,
    points: markers.length
  };
};


let symbolIndex = 0;
let globalBounds = L.latLngBounds(); // used to zoom map to all points


// Convert day number → day name
function dayName(n) {
  return ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"][n-1];
}


// Assign a unique color/shape to each route/day
function getSymbol(key) {
  if (!symbolMap[key]) {
    // Generate a distinct color per route+day using golden-angle hue stepping.
    const hue = Math.round((symbolIndex * 137.508) % 360);
    symbolMap[key] = {
      color: `hsl(${hue} 80% 52%)`,
      shape: shapes[symbolIndex % shapes.length]
    };
    symbolIndex++;
  }
  return symbolMap[key];
}


  function getMarkerPixelSize() {
  // Same size for all routes at a given zoom level, but smaller when zoomed out.
  const z = map.getZoom();
  return MARKER_SIZE_STEPS.find(([max]) => z <= max)[1];
}





// Create marker with correct shape
function createMarker(lat, lon, symbol) {
  const size = getMarkerPixelSize();

  // ===== CIRCLE =====
  if (symbol.shape === "circle") {
    const marker = L.circleMarker([lat, lon], {
      radius: size,
      color: symbol.color,
      fillColor: symbol.color,
      fillOpacity: 0.95,
      renderer: canvasRenderer
    });

    marker._base = { lat, lon, symbol };
    return marker;
  }

  function pixelOffset() {
    const zoom = map.getZoom();
    const scale = 40075016.686 / Math.pow(2, zoom + 8);
    const latOffset = size * scale / 111320;
    const lngOffset = latOffset / Math.cos(lat * Math.PI / 180);
    return [latOffset, lngOffset];
  }

  const [dLat, dLng] = pixelOffset();

  let shape;

  if (symbol.shape === "square") {
    shape = L.rectangle([[lat - dLat, lon - dLng], [lat + dLat, lon + dLng]], {
      color: symbol.color,
      fillColor: symbol.color,
      fillOpacity: 0.95,
      weight: 1,
      renderer: canvasRenderer
    });
  }

  if (symbol.shape === "triangle") {
    shape = L.polygon(
      [[lat + dLat, lon], [lat - dLat, lon - dLng], [lat - dLat, lon + dLng]],
      {
        color: symbol.color,
        fillColor: symbol.color,
        fillOpacity: 0.95,
        weight: 1,
        renderer: canvasRenderer
      }
    );
  }

  if (symbol.shape === "diamond") {
    shape = L.polygon(
      [[lat + dLat, lon], [lat, lon + dLng], [lat - dLat, lon], [lat, lon - dLng]],
      {
        color: symbol.color,
        fillColor: symbol.color,
        fillOpacity: 0.95,
        weight: 1,
        renderer: canvasRenderer
      }
    );
  }

  shape._base = { lat, lon, symbol };
  return shape;
}



// ================= FILTER CHECKBOX UI =================
function buildRouteCheckboxes(routes) {
  const c = document.getElementById("routeCheckboxes");
  c.innerHTML = "";

  routes.forEach(route => {
    const label = document.createElement("label");

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = route;
    checkbox.checked = true;
    checkbox.addEventListener("change", applyFilters);

    const text = document.createTextNode(" " + route);

    label.appendChild(checkbox);
    label.appendChild(text);

    c.appendChild(label);
  });
}



function buildDayCheckboxes() {
  const c = document.getElementById("dayCheckboxes");
  c.innerHTML = "";

  [1,2,3,4,5,6,7].forEach(d => {
    const l = document.createElement("label");
    l.innerHTML = `<input type="checkbox" value="${d}" checked> ${dayName(d)}`;
    l.querySelector("input").addEventListener("change", applyFilters);
    c.appendChild(l);
  });
}
buildDayCheckboxes();


// Select/Deselect all checkboxes
function setCheckboxGroup(containerId, checked) {
  document.querySelectorAll(`#${containerId} input`).forEach(b => (b.checked = checked));
  applyFilters();
}

document.getElementById("routesAll").onclick  = () => setCheckboxGroup("routeCheckboxes", true);
document.getElementById("routesNone").onclick = () => setCheckboxGroup("routeCheckboxes", false);
document.getElementById("daysAll").onclick    = () => setCheckboxGroup("dayCheckboxes", true);
document.getElementById("daysNone").onclick   = () => setCheckboxGroup("dayCheckboxes", false);



// ===== Route + Day ALL / NONE =====
document.getElementById("routeDayAll").onclick  = () => {
  document.querySelectorAll("#routeDayLayers input[type='checkbox']")
    .forEach(cb => {
      cb.checked = true;
      cb.dispatchEvent(new Event("change"));
    });
};

document.getElementById("routeDayNone").onclick = () => {
  document.querySelectorAll("#routeDayLayers input[type='checkbox']")
    .forEach(cb => {
      cb.checked = false;
      cb.dispatchEvent(new Event("change"));
    });
};


// ================= APPLY MAP FILTERS =================
function applyFilters() {

  const routeCheckboxes = [...document.querySelectorAll("#routeCheckboxes input")];
  const dayCheckboxes   = [...document.querySelectorAll("#dayCheckboxes input")];

  const routes = routeCheckboxes.filter(i => i.checked).map(i => i.value);
  const days = dayCheckboxes.filter(i => i.checked).map(i => i.value);

  Object.entries(routeDayGroups).forEach(([key, group]) => {
    const [r, d] = key.split("|");
    const show = routes.includes(r) && days.includes(d);
    group.layers.forEach(l => show ? l.addTo(map) : map.removeLayer(l));
  });

  applyLayerManagerOrder();
  updateSelectionCount();
  updateStats();
  refreshLayerManagerUiIfOpen();
}



// ================= ROUTE STATISTICS =================
function updateStats() {
  const list = document.getElementById("statsList");
  if (!list) return;
  list.innerHTML = "";

  Object.entries(routeDayGroups).forEach(([key, group]) => {
    const visible = group.layers.filter(l => map.hasLayer(l)).length;
    if (!visible) return;

    const [r,d] = key.split("|");
    const li = document.createElement("li");
    li.textContent = `Route ${r} – ${dayName(d)}: ${visible}`;
    list.appendChild(li);
  });
}

function refreshRouteDayLayerSelectButtons() {
  document.querySelectorAll("#routeDayLayers button[data-select-layer-key]").forEach(btn => {
    const key = String(btn.dataset.selectLayerKey || "").trim();
    if (!key) return;
    const selectable = isLayerManagerEntrySelectable(key);
    btn.disabled = !selectable;
    btn.title = selectable
      ? "Select this layer"
      : "Enable Selectable in Layer Manager to use layer selection.";
  });
}

function selectEntireLayer(key) {
  if (!isLayerManagerEntrySelectable(key)) return;
  const group = routeDayGroups[key];
  if (!group || !group.layers || !group.layers.length) return;

  layerVisibilityState[key] = true;
  const layerCheckbox = document.querySelector(`input[data-key="${key}"]`);
  if (layerCheckbox) layerCheckbox.checked = true;

  const bounds = L.latLngBounds();
  group.layers.forEach(marker => {
    map.addLayer(marker);
    const ll = getLayerLatLng(marker);
    if (ll) bounds.extend(ll);
  });
  applyLayerManagerOrder();

  if (!bounds.isValid()) return;

  selectedLayerKey = key;
  drawnLayer.clearLayers();

  updateSelectionCount();
  updateUndoButtonState();
  map.fitBounds(bounds, { padding: [30, 30], maxZoom: 16 });
}
  // ===== BUILD ROUTE + DAY LAYER CHECKBOXES =====
// ===== BUILD ROUTE + DAY LAYER CHECKBOXES =====
function buildRouteDayLayerControls() {
  const routeDayContainer = document.getElementById("routeDayLayers");
  if (!routeDayContainer) return;

  routeDayContainer.innerHTML = "";

  const daySortRank = value => {
    const v = String(value ?? "").trim().toLowerCase();
    const byName = {
      "1": 1, mon: 1, monday: 1,
      "2": 2, tue: 2, tues: 2, tuesday: 2,
      "3": 3, wed: 3, wednesday: 3,
      "4": 4, thu: 4, thur: 4, thurs: 4, thursday: 4,
      "5": 5, fri: 5, friday: 5,
      "6": 6, sat: 6, saturday: 6,
      "7": 7, sun: 7, sunday: 7,
      delivered: 99
    };
    if (Object.prototype.hasOwnProperty.call(byName, v)) return byName[v];
    const n = Number(v);
    return Number.isFinite(n) ? n : 98;
  };

  const sortedRouteDayEntries = Object.entries(routeDayGroups).sort(([aKey], [bKey]) => {
    const [aRoute = "", aDay = ""] = aKey.split("|");
    const [bRoute = "", bDay = ""] = bKey.split("|");
    const routeCmp = aRoute.localeCompare(bRoute, undefined, { numeric: true, sensitivity: "base" });
    if (routeCmp !== 0) return routeCmp;
    return daySortRank(aDay) - daySortRank(bDay);
  });

  sortedRouteDayEntries.forEach(([key, group]) => {
    const count = group.layers ? group.layers.length : 0;
    const [route, type] = key.split("|");
    const dayNameMap = {
      1: "Monday",
      2: "Tuesday",
      3: "Wednesday",
      4: "Thursday",
      5: "Friday",
      6: "Saturday",
      7: "Sunday"
    };

    const wrapper = document.createElement("div");
    wrapper.className = "layer-item";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.key = key;

    if (layerVisibilityState.hasOwnProperty(key)) {
      checkbox.checked = layerVisibilityState[key];
    } else {
      checkbox.checked = true;
      layerVisibilityState[key] = true;
    }

    routeDayGroups[key].layers.forEach(marker => {
      if (checkbox.checked) map.addLayer(marker);
      else map.removeLayer(marker);
    });

    checkbox.addEventListener("change", () => {
      layerVisibilityState[key] = checkbox.checked;
      routeDayGroups[key].layers.forEach(marker => {
        if (checkbox.checked) map.addLayer(marker);
        else map.removeLayer(marker);
      });
      applyLayerManagerOrder();
      updateStats();
      refreshLayerManagerUiIfOpen();
    });

    const symbol = getSymbol(key);
    const preview = document.createElement("span");
    preview.className = "layer-preview";
    preview.style.background = symbol.color;

    if (symbol.shape === "circle") preview.style.borderRadius = "50%";
    if (symbol.shape === "square") preview.style.borderRadius = "2px";

    if (symbol.shape === "triangle") {
      preview.style.background = "transparent";
      preview.style.width = "0";
      preview.style.height = "0";
      preview.style.borderLeft = "7px solid transparent";
      preview.style.borderRight = "7px solid transparent";
      preview.style.borderBottom = `14px solid ${symbol.color}`;
    }

    if (symbol.shape === "diamond") {
      preview.style.transform = "rotate(45deg)";
    }

    const labelText = document.createElement("span");
    const dayName = dayNameMap[type] || type;
    labelText.textContent = `Route ${route} - ${dayName} (${count})`;

    const selectBtn = document.createElement("button");
    selectBtn.type = "button";
    selectBtn.className = "mini-btn";
    selectBtn.textContent = "Select";
    selectBtn.dataset.selectLayerKey = key;
    const selectable = isLayerManagerEntrySelectable(key);
    selectBtn.disabled = !selectable;
    selectBtn.title = selectable
      ? "Select this layer"
      : "Enable Selectable in Layer Manager to use layer selection.";
    selectBtn.addEventListener("click", e => {
      e.preventDefault();
      e.stopPropagation();
      selectEntireLayer(key);
    });

    wrapper.appendChild(checkbox);
    wrapper.appendChild(preview);
    wrapper.appendChild(labelText);
    wrapper.appendChild(selectBtn);
    routeDayContainer.appendChild(wrapper);
  });

  ensureLayerManagerOrder();
  ensureLayerManagerSelectableState();
  applyLayerManagerOrder();
  refreshRouteDayLayerSelectButtons();
  refreshLayerManagerUiIfOpen();
}


// ================= COLUMN MAPPING (UPLOAD) =================
const COLUMN_MAPPING_FIELDS = [
  { key: "LATITUDE", label: "Latitude", required: true },
  { key: "LONGITUDE", label: "Longitude", required: true },
  { key: "NEWROUTE", label: "Route Code", required: false },
  { key: "NEWDAY", label: "Day", required: false },
  { key: "del_status", label: "Delivery Status", required: false },
  { key: "CSADR#", label: "Street Number", required: false },
  { key: "CSSDIR", label: "Street Direction", required: false },
  { key: "CSSTRT", label: "Street Name", required: false },
  { key: "CSSFUX", label: "Street Suffix", required: false },
  { key: "SIZE", label: "Container Size", required: false },
  { key: "QTY", label: "Quantity", required: false },
  { key: "BINNO", label: "Bin Number", required: false }
];

const COLUMN_MAPPING_ALIASES = {
  LATITUDE: ["latitude", "lat", "y", "ycoord", "ycoordinate"],
  LONGITUDE: ["longitude", "lon", "lng", "long", "x", "xcoord", "xcoordinate"],
  NEWROUTE: ["newroute", "route", "routecode", "routeid", "rte"],
  NEWDAY: ["newday", "day", "weekday", "serviceday"],
  del_status: ["delstatus", "deliverystatus", "status", "delivered", "stopstatus"],
  "CSADR#": ["csadr", "addressnumber", "streetnumber", "housenumber"],
  CSSDIR: ["cssdir", "streetdir", "direction", "predir"],
  CSSTRT: ["csstrt", "street", "streetname", "road"],
  CSSFUX: ["cssfux", "streetsuffix", "suffix", "sfx"],
  SIZE: ["size", "containersize", "binsize"],
  QTY: ["qty", "quantity", "count"],
  BINNO: ["binno", "bin", "binnumber", "containerid"]
};

function normalizeColumnName(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getColumnMappingStorageKey(headers) {
  const normalized = [...headers]
    .map(normalizeColumnName)
    .filter(Boolean)
    .sort()
    .join("|");
  return `colmap:${APP_STORAGE_NS}:${normalized}`;
}

function loadSavedColumnMapping(headers) {
  try {
    const raw = localStorage.getItem(getColumnMappingStorageKey(headers));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch (_) {
    return null;
  }
}

function saveColumnMapping(headers, mapping) {
  try {
    localStorage.setItem(getColumnMappingStorageKey(headers), JSON.stringify(mapping || {}));
  } catch (_) {
    // Ignore storage failures.
  }
}

function buildColumnMappingGuess(headers) {
  const normalizedHeaders = headers.map(h => ({
    raw: h,
    norm: normalizeColumnName(h)
  }));

  const mapping = {};
  const used = new Set();

  const pickBest = field => {
    const fieldNorm = normalizeColumnName(field.key);
    const aliases = [field.key, ...(COLUMN_MAPPING_ALIASES[field.key] || [])]
      .map(normalizeColumnName);

    const scored = normalizedHeaders
      .filter(h => !used.has(h.raw))
      .map(h => {
        let score = 0;
        if (h.norm === fieldNorm) score = 100;
        else if (aliases.includes(h.norm)) score = 90;
        else if (aliases.some(a => a && (h.norm.includes(a) || a.includes(h.norm)))) score = 60;
        return { header: h.raw, score };
      })
      .filter(x => x.score > 0)
      .sort((a, b) => b.score - a.score);

    return scored.length ? scored[0].header : "";
  };

  COLUMN_MAPPING_FIELDS.forEach(field => {
    const picked = pickBest(field);
    if (picked) used.add(picked);
    mapping[field.key] = picked;
  });

  return mapping;
}

function ensureColumnMappingModal() {
  let modal = document.getElementById("columnMappingModal");
  if (modal) return modal;

  modal = document.createElement("div");
  modal.id = "columnMappingModal";
  modal.style.cssText = [
    "position:fixed",
    "inset:0",
    "background:rgba(10,14,20,0.68)",
    "display:none",
    "align-items:center",
    "justify-content:center",
    "z-index:12000",
    "padding:18px"
  ].join(";");

  modal.innerHTML = `
    <div style="width:min(680px,100%);max-height:86vh;overflow:auto;background:#172230;border:1px solid #32465c;border-radius:12px;padding:14px 14px 12px;color:#eef6ff;">
      <div style="font-size:18px;font-weight:700;margin-bottom:6px;">Map Upload Columns</div>
      <div id="columnMappingDesc" style="font-size:12px;color:#b7c9db;margin-bottom:12px;"></div>
      <div style="display:flex;gap:8px;justify-content:flex-end;margin-bottom:10px;">
        <button id="columnMappingSuggestBtn" type="button" class="primary-btn" style="background:#2a3a4d;color:#eaf4ff;">Use Suggested</button>
      </div>
      <div id="columnMappingRequiredTitle" style="font-size:12px;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#dff0ff;margin:0 0 8px;"></div>
      <div id="columnMappingRequiredFields" style="display:grid;grid-template-columns:1fr 1fr;gap:10px 12px;margin-bottom:12px;"></div>
      <div id="columnMappingOptionalTitle" style="font-size:12px;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:#9eb6cc;margin:0 0 8px;">Optional Fields</div>
      <div id="columnMappingOptionalFields" style="display:grid;grid-template-columns:1fr 1fr;gap:10px 12px;"></div>
      <div id="columnMappingError" style="min-height:18px;color:#ffb4b4;font-size:12px;margin-top:8px;"></div>
      <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:8px;">
        <button id="columnMappingCancelBtn" type="button" class="primary-btn" style="background:#3b4d62;color:#fff;">Cancel</button>
        <button id="columnMappingApplyBtn" type="button" class="primary-btn">Apply Mapping</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);
  return modal;
}

function openColumnMappingPrompt(headers, initialMapping, fileName, sampleByHeader = {}) {
  return new Promise(resolve => {
    const modal = ensureColumnMappingModal();
    const desc = modal.querySelector("#columnMappingDesc");
    const requiredTitle = modal.querySelector("#columnMappingRequiredTitle");
    const requiredBox = modal.querySelector("#columnMappingRequiredFields");
    const optionalBox = modal.querySelector("#columnMappingOptionalFields");
    const errorBox = modal.querySelector("#columnMappingError");
    const suggestBtn = modal.querySelector("#columnMappingSuggestBtn");
    const cancelBtn = modal.querySelector("#columnMappingCancelBtn");
    const applyBtn = modal.querySelector("#columnMappingApplyBtn");

    desc.textContent = `File: ${fileName || "Selected file"}. Only Latitude and Longitude are required. Missing optional fields will use smart defaults.`;
    errorBox.textContent = "";
    requiredBox.innerHTML = "";
    optionalBox.innerHTML = "";

    const selectByKey = {};
    const sortedHeaders = [...headers].sort((a, b) =>
      String(a).localeCompare(String(b), undefined, { sensitivity: "base", numeric: true })
    );
    const updateRequiredTitle = () => {
      const missing = COLUMN_MAPPING_FIELDS
        .filter(f => f.required)
        .filter(f => !selectByKey[f.key]?.value).length;
      requiredTitle.textContent = `Required Fields (${missing} missing)`;
      requiredTitle.style.color = missing ? "#ffd6d6" : "#d6ffd8";
    };

    const renderField = (field, container) => {
      const row = document.createElement("label");
      row.style.display = "flex";
      row.style.flexDirection = "column";
      row.style.gap = "4px";
      row.style.padding = "8px";
      row.style.border = "1px solid #33485d";
      row.style.borderRadius = "9px";
      row.style.background = "rgba(255,255,255,0.03)";

      const title = document.createElement("span");
      title.textContent = `${field.label}${field.required ? " *" : ""}`;
      title.style.fontSize = "12px";
      title.style.color = field.required ? "#dff0ff" : "#b7c9db";

      const select = document.createElement("select");
      select.style.cssText = "height:32px;border-radius:8px;border:1px solid #3f566f;background:#0f1822;color:#eef6ff;padding:0 8px;";
      select.dataset.key = field.key;

      const blank = document.createElement("option");
      blank.value = "";
      blank.textContent = field.required ? "Select a column..." : "Not mapped";
      select.appendChild(blank);

      sortedHeaders.forEach(header => {
        const opt = document.createElement("option");
        opt.value = header;
        opt.textContent = header;
        select.appendChild(opt);
      });

      const expandedSize = Math.min(10, Math.max(4, sortedHeaders.length + 1));
      const collapseSelect = () => {
        select.size = 1;
        select.style.height = "32px";
        select.dataset.expanded = "0";
      };
      const expandSelect = () => {
        select.size = expandedSize;
        select.style.height = "auto";
        select.dataset.expanded = "1";
      };

      collapseSelect();
      if (window.innerWidth > 900) {
        select.addEventListener("mousedown", e => {
          // Keep the list attached to the field so it remains scrollable in-page.
          e.preventDefault();
          if (select.dataset.expanded === "1") collapseSelect();
          else expandSelect();
          select.focus();
        });
      }
      select.addEventListener("blur", collapseSelect);

      const sample = document.createElement("span");
      sample.style.fontSize = "11px";
      sample.style.color = "#9eb6cc";
      sample.style.minHeight = "14px";

      const setSample = () => {
        const picked = select.value;
        const value = picked ? sampleByHeader[picked] : "";
        sample.textContent = picked
          ? `Sample: ${String(value || "(blank)").slice(0, 80)}`
          : "Sample: —";
      };

      select.value = initialMapping[field.key] || "";
      select.addEventListener("change", () => {
        setSample();
        updateRequiredTitle();
        collapseSelect();
      });
      setSample();

      row.appendChild(title);
      row.appendChild(select);
      row.appendChild(sample);
      container.appendChild(row);
      selectByKey[field.key] = select;
    };

    COLUMN_MAPPING_FIELDS.filter(f => f.required).forEach(field => renderField(field, requiredBox));
    COLUMN_MAPPING_FIELDS.filter(f => !f.required).forEach(field => renderField(field, optionalBox));
    updateRequiredTitle();

    suggestBtn.onclick = () => {
      COLUMN_MAPPING_FIELDS.forEach(field => {
        selectByKey[field.key].value = initialMapping[field.key] || "";
        selectByKey[field.key].dispatchEvent(new Event("change"));
      });
      errorBox.textContent = "";
    };

    const close = value => {
      modal.style.display = "none";
      suggestBtn.onclick = null;
      cancelBtn.onclick = null;
      applyBtn.onclick = null;
      resolve(value);
    };

    cancelBtn.onclick = () => close(null);
    applyBtn.onclick = () => {
      const mapping = {};
      const selected = [];

      for (const field of COLUMN_MAPPING_FIELDS) {
        const value = selectByKey[field.key].value;
        if (field.required && !value) {
          errorBox.textContent = `Please map required field: ${field.label}.`;
          return;
        }
        if (value) selected.push(value);
        mapping[field.key] = value;
      }

      const dupes = selected.filter((v, i) => selected.indexOf(v) !== i);
      if (dupes.length) {
        errorBox.textContent = "Each source column can only be mapped once.";
        return;
      }

      saveColumnMapping(headers, mapping);
      close(mapping);
    };

    modal.style.display = "flex";
  });
}

function applyColumnAliasesToRows(rows, mapping) {
  if (!Array.isArray(rows)) return;
  rows.forEach(row => {
    if (!row || typeof row !== "object") return;
    COLUMN_MAPPING_FIELDS.forEach(field => {
      const targetKey = field.key;
      const sourceKey = mapping[targetKey];
      if (!sourceKey || sourceKey === targetKey) return;
      if (Object.prototype.hasOwnProperty.call(row, targetKey)) return;

      Object.defineProperty(row, targetKey, {
        configurable: true,
        enumerable: false,
        get() {
          return row[sourceKey];
        },
        set(value) {
          row[sourceKey] = value;
        }
      });
    });
  });
}

function getAttributeHeaders(rows) {
  const seen = new Set();
  (rows || []).forEach(row => {
    Object.keys(row || {}).forEach(key => {
      if (key && !seen.has(key)) seen.add(key);
    });
  });
  return [...seen];
}

function getAttributeRowId(row) {
  return attributeRowToId.get(row);
}

function getAttributeMarker(rowId) {
  return attributeMarkerByRowId.get(rowId) || null;
}

function refreshAttributeStatus() {
  const status = document.getElementById("attributeStatus");
  if (!status) return;
  const selectedCount = attributeTableMode === "streets"
    ? streetAttributeSelectedIds.size
    : attributeState.selectedRowIds.size;
  const visibleCount = attributeState.lastVisibleRows.length;
  status.textContent = `${selectedCount} selected • ${visibleCount} visible`;
}

function syncSelectedStopsHeaderCount(count) {
  const countNode = document.getElementById("selectionCount");
  if (countNode) countNode.textContent = String(Math.max(0, Number(count) || 0));
  const mobileBtn = document.getElementById("mobileSelectionBtn");
  if (mobileBtn) mobileBtn.textContent = `Selected: ${countNode?.textContent || "0"}`;
}

function applyAttributeSelectionStyles() {
  attributeMarkerByRowId.forEach((marker, rowId) => {
    if (!marker || !marker._base?.symbol) return;
    const baseColor = marker._base.symbol.color;
    const selected = attributeState.selectedRowIds.has(rowId);
    const color = selected ? "#ffe066" : baseColor;
    marker.setStyle?.({
      color,
      fillColor: color,
      fillOpacity: selected ? 1 : 0.95
    });
  });
}

function setAttributeRowSelected(rowId, selected, rerender = true) {
  if (!Number.isFinite(rowId)) return;
  if (selected) {
    attributeState.selectedRowIds.add(rowId);
  } else {
    attributeState.selectedRowIds.delete(rowId);
  }
  applyAttributeSelectionStyles();
  refreshAttributeStatus();
  syncSelectedStopsHeaderCount(attributeState.selectedRowIds.size);
  if (rerender) renderAttributeTable();
}

window.focusAttributeRowOnMap = function(rowId) {
  const marker = getAttributeMarker(Number(rowId));
  if (!marker) return;
  const latlng = getLayerLatLng(marker);
  if (!latlng) return;
  map.setView(latlng, Math.max(map.getZoom(), 16), { animate: true });
  marker.openPopup?.();
};

window.getAttributeSelectedRowIds = function() {
  return [...attributeState.selectedRowIds];
};

window.setAttributeSelectedRowIds = function(rowIds) {
  const ids = Array.isArray(rowIds) ? rowIds : [];
  const next = new Set(
    ids
      .map(v => Number(v))
      .filter(v => Number.isFinite(v))
  );
  const prev = attributeState.selectedRowIds;
  const changed =
    prev.size !== next.size ||
    [...prev].some(id => !next.has(id));

  if (!changed) return false;
  attributeState.selectedRowIds = next;
  applyAttributeSelectionStyles();
  refreshAttributeStatus();
  syncSelectedStopsHeaderCount(next.size);
  renderAttributeTable();
  return true;
};

window.zoomToAttributeRowsOnMap = function(rowIds) {
  const ids = Array.isArray(rowIds) ? rowIds : [];
  if (!ids.length) return false;
  const bounds = L.latLngBounds();
  ids.forEach(id => {
    const marker = getAttributeMarker(Number(id));
    const latlng = getLayerLatLng(marker);
    if (latlng) bounds.extend(latlng);
  });
  if (!bounds.isValid()) return false;
  map.fitBounds(bounds.pad(0.18));
  return true;
};

window.getStreetSelectedSegmentIds = function() {
  return [...streetAttributeSelectedIds];
};

window.setStreetSelectedSegmentIds = function(rowIds) {
  const ids = Array.isArray(rowIds) ? rowIds : [];
  const next = new Set(
    ids
      .map(v => Number(v))
      .filter(v => Number.isFinite(v) && streetAttributeById.has(v))
  );
  const prev = streetAttributeSelectedIds;
  const changed =
    prev.size !== next.size ||
    [...prev].some(id => !next.has(id));

  if (!changed) return false;
  streetAttributeSelectedIds.clear();
  next.forEach(id => streetAttributeSelectedIds.add(id));
  applyStreetSelectionStyles();
  refreshAttributeStatus();
  syncSelectedStopsHeaderCount(next.size);
  renderAttributeTable();
  return true;
};

window.focusStreetSegmentOnMap = function(rowId) {
  const entry = streetAttributeById.get(Number(rowId));
  if (!entry?.layer || typeof entry.layer.getBounds !== "function") return false;
  map.fitBounds(entry.layer.getBounds().pad(0.35));
  return true;
};

window.zoomToStreetSegmentsOnMap = function(rowIds) {
  const ids = Array.isArray(rowIds) ? rowIds : [];
  if (!ids.length) return false;
  const bounds = L.latLngBounds();
  ids.forEach(id => {
    const entry = streetAttributeById.get(Number(id));
    const b = entry?.layer?.getBounds?.();
    if (b?.isValid?.()) bounds.extend(b);
  });
  if (!bounds.isValid()) return false;
  map.fitBounds(bounds.pad(0.15));
  return true;
};

const SELECT_ATTRIBUTES_SAVED_QUERIES_KEY = "selectAttributesSavedQueries";
const SELECT_ATTRIBUTES_MAX_UNIQUE_VALUES = 2000;

function getSelectAttributesSourceLabel(sourceKey) {
  const key = String(sourceKey || "").trim().toLowerCase();
  if (key === "streets") return "Street Attributes";
  if (key === "layers") return "Map Layer Features";
  return "Record Attributes";
}

function getSelectAttributesDataset(sourceKey) {
  const key = String(sourceKey || "records").trim().toLowerCase();

  if (key === "streets") {
    const rows = (Array.isArray(streetAttributesRows) && streetAttributesRows.length)
      ? streetAttributesRows
      : [...streetAttributeById.values()].map(v => v?.row).filter(Boolean);
    return rows
      .map(row => {
        const streetId = Number(row?.id);
        return { row, streetId };
      })
      .filter(item => Number.isFinite(item.streetId));
  }

  if (key === "layers") {
    const out = [];
    Object.entries(routeDayGroups).forEach(([routeDayKey, group]) => {
      const layers = Array.isArray(group?.layers) ? group.layers : [];
      layers.forEach(marker => {
        const latlng = getLayerLatLng(marker);
        const rowId = Number(marker?._rowId);
        out.push({
          row: {
            row_id: Number.isFinite(rowId) ? rowId : "",
            route_day: routeDayKey || "",
            symbol_label: marker?._base?.symbol?.label || marker?._base?.symbol?.name || "",
            symbol_color: marker?._base?.symbol?.color || "",
            latitude: Number.isFinite(latlng?.lat) ? Number(latlng.lat.toFixed(7)) : "",
            longitude: Number.isFinite(latlng?.lng) ? Number(latlng.lng.toFixed(7)) : ""
          },
          rowId: Number.isFinite(rowId) ? rowId : null
        });
      });
    });
    return out;
  }

  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  return rows.map((row, idx) => {
    const mappedRowId = getAttributeRowId(row);
    const rowId = Number.isFinite(mappedRowId) ? mappedRowId : idx;
    return { row, rowId };
  });
}

function getSelectAttributesFieldNames(sourceKey, dataset = null) {
  const key = String(sourceKey || "records").trim().toLowerCase();
  if (key === "records") {
    const fromHeaders = Array.isArray(window._attributeHeaders) ? window._attributeHeaders.filter(Boolean) : [];
    if (fromHeaders.length) return fromHeaders;
  }
  const rows = (Array.isArray(dataset) ? dataset : getSelectAttributesDataset(sourceKey)).map(entry => entry?.row || {});
  return getAttributeHeaders(rows);
}

function getSelectAttributesSavedQueries() {
  const raw = storageGet(SELECT_ATTRIBUTES_SAVED_QUERIES_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map(item => ({
        id: String(item?.id || ""),
        name: String(item?.name || "").trim(),
        source: String(item?.source || "records").trim().toLowerCase(),
        where: String(item?.where || "").trim(),
        updatedAt: Number(item?.updatedAt || 0)
      }))
      .filter(item => item.id && item.name && item.source);
  } catch {
    return [];
  }
}

function setSelectAttributesSavedQueries(queries) {
  const next = Array.isArray(queries) ? queries : [];
  storageSet(SELECT_ATTRIBUTES_SAVED_QUERIES_KEY, JSON.stringify(next.slice(0, 250)));
}

function normalizeSelectAttributesWhereText(input) {
  let text = String(input || "");
  text = text.replace(/\r?\n/g, " ");
  text = text.replace(/\s+/g, " ").trim();
  text = text.replace(/\(\s+/g, "(").replace(/\s+\)/g, ")");
  text = text.replace(/\s+,/g, ",").replace(/,\s*/g, ", ");
  return text;
}

function escapeSelectAttributesHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function quoteSelectAttributesSqlString(value) {
  return `'${String(value ?? "").replace(/'/g, "''")}'`;
}

function formatSelectAttributesSqlLiteral(value) {
  if (value === null || value === undefined) return "NULL";
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  const text = String(value);
  if (!text.length) return "''";
  return quoteSelectAttributesSqlString(text);
}

function normalizeSelectAttributesInputLiteral(raw) {
  const text = String(raw || "").trim();
  if (!text) return "";
  const upper = text.toUpperCase();
  if (upper === "NULL" || upper === "TRUE" || upper === "FALSE") return upper;
  if (/^-?\d+(\.\d+)?$/.test(text)) return text;
  if (text.startsWith("'") && text.endsWith("'") && text.length >= 2) return text;
  return quoteSelectAttributesSqlString(text);
}

function detectLastSelectAttributesField(whereClause) {
  const text = String(whereClause || "");
  const bracketMatches = [...text.matchAll(/\[([^\]]+)\]/g)];
  if (bracketMatches.length) {
    const value = String(bracketMatches[bracketMatches.length - 1]?.[1] || "").trim();
    return value || "";
  }
  const bareMatches = [...text.matchAll(/\b([A-Za-z_][A-Za-z0-9_]*)\b(?=\s*(=|!=|<>|>=|<=|>|<|LIKE|IN|IS))/gi)];
  if (!bareMatches.length) return "";
  return String(bareMatches[bareMatches.length - 1]?.[1] || "").trim();
}

function getSelectAttributesRowFieldValue(row, fieldName) {
  if (!row || typeof row !== "object") return null;
  const name = String(fieldName || "").trim();
  if (!name) return null;
  if (Object.prototype.hasOwnProperty.call(row, name)) return row[name];

  const target = name.toLowerCase();
  const key = Object.keys(row).find(k => String(k).toLowerCase() === target);
  if (!key) return null;
  return row[key];
}

function tokenizeSelectAttributesWhereClause(whereText) {
  const text = String(whereText || "");
  const tokens = [];
  let i = 0;

  const isDigit = ch => /[0-9]/.test(ch);
  const isWord = ch => /[A-Za-z0-9_.]/.test(ch);

  while (i < text.length) {
    const ch = text[i];
    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }

    if (ch === "(" || ch === ")" || ch === ",") {
      tokens.push({ type: "punct", value: ch });
      i += 1;
      continue;
    }

    const two = text.slice(i, i + 2);
    if (two === ">=" || two === "<=" || two === "<>" || two === "!=") {
      tokens.push({ type: "op", value: two });
      i += 2;
      continue;
    }
    if (ch === "=" || ch === ">" || ch === "<") {
      tokens.push({ type: "op", value: ch });
      i += 1;
      continue;
    }

    if (ch === "'") {
      i += 1;
      let value = "";
      let closed = false;
      while (i < text.length) {
        if (text[i] === "'" && text[i + 1] === "'") {
          value += "'";
          i += 2;
          continue;
        }
        if (text[i] === "'") {
          closed = true;
          i += 1;
          break;
        }
        value += text[i];
        i += 1;
      }
      if (!closed) throw new Error("Unclosed string literal in WHERE clause.");
      tokens.push({ type: "string", value });
      continue;
    }

    if (ch === "[") {
      const end = text.indexOf("]", i + 1);
      if (end === -1) throw new Error("Unclosed [field] identifier in WHERE clause.");
      const value = text.slice(i + 1, end).trim();
      if (!value) throw new Error("Empty [field] identifier is not allowed.");
      tokens.push({ type: "identifier", value });
      i = end + 1;
      continue;
    }

    const isSignedNumber = (ch === "-" || ch === "+") && isDigit(text[i + 1]);
    if (isDigit(ch) || isSignedNumber) {
      const start = i;
      i += 1;
      while (i < text.length && /[0-9.]/.test(text[i])) i += 1;
      const raw = text.slice(start, i);
      if (/^[+-]?\d+(\.\d+)?$/.test(raw)) {
        tokens.push({ type: "number", value: Number(raw) });
      } else {
        throw new Error(`Invalid numeric literal "${raw}" in WHERE clause.`);
      }
      continue;
    }

    if (isWord(ch)) {
      const start = i;
      i += 1;
      while (i < text.length && isWord(text[i])) i += 1;
      const raw = text.slice(start, i);
      const upper = raw.toUpperCase();
      if (["AND", "OR", "NOT", "IS", "NULL", "IN", "LIKE", "TRUE", "FALSE"].includes(upper)) {
        tokens.push({ type: "keyword", value: upper });
      } else {
        tokens.push({ type: "identifier", value: raw });
      }
      continue;
    }

    throw new Error(`Unsupported character "${ch}" in WHERE clause.`);
  }

  return tokens;
}

function parseSelectAttributesWhereClause(whereText) {
  const tokens = tokenizeSelectAttributesWhereClause(whereText);
  if (!tokens.length) return { type: "all" };
  let pos = 0;

  const peek = () => tokens[pos] || null;
  const consume = () => tokens[pos++] || null;
  const match = (type, value = null) => {
    const token = peek();
    if (!token) return null;
    if (token.type !== type) return null;
    if (value !== null && token.value !== value) return null;
    pos += 1;
    return token;
  };
  const expect = (type, value = null, message = "Invalid query syntax.") => {
    const token = match(type, value);
    if (token) return token;
    throw new Error(message);
  };

  const parseField = () => {
    const token = expect("identifier", null, "Expected a field name.");
    return String(token.value || "").trim();
  };

  const parseValue = () => {
    const token = consume();
    if (!token) throw new Error("Expected a value.");
    if (token.type === "number" || token.type === "string") return token.value;
    if (token.type === "keyword") {
      if (token.value === "NULL") return null;
      if (token.value === "TRUE") return true;
      if (token.value === "FALSE") return false;
      throw new Error(`Expected a value, got keyword "${token.value}".`);
    }
    if (token.type === "identifier") return token.value;
    throw new Error("Expected a value.");
  };

  const parsePredicate = () => {
    const field = parseField();

    if (match("keyword", "IS")) {
      if (match("keyword", "NOT")) {
        expect("keyword", "NULL", 'Expected "NULL" after "IS NOT".');
        return { type: "is_not_null", field };
      }
      expect("keyword", "NULL", 'Expected "NULL" after "IS".');
      return { type: "is_null", field };
    }

    if (match("keyword", "IN")) {
      expect("punct", "(", 'Expected "(" after "IN".');
      const values = [];
      if (!match("punct", ")")) {
        values.push(parseValue());
        while (match("punct", ",")) values.push(parseValue());
        expect("punct", ")", 'Expected ")" to close "IN" list.');
      }
      return { type: "in", field, values };
    }

    if (match("keyword", "LIKE")) {
      const value = parseValue();
      return { type: "like", field, value };
    }

    const opToken = expect("op", null, "Expected a comparison operator.");
    if (!["=", "!=", "<>", ">", ">=", "<", "<="].includes(opToken.value)) {
      throw new Error(`Unsupported operator "${opToken.value}".`);
    }
    const value = parseValue();
    return { type: "compare", field, op: opToken.value, value };
  };

  const parsePrimary = () => {
    if (match("punct", "(")) {
      const expr = parseOr();
      expect("punct", ")", 'Expected ")" after grouped condition.');
      return expr;
    }
    return parsePredicate();
  };

  const parseNot = () => {
    if (match("keyword", "NOT")) return { type: "not", value: parseNot() };
    return parsePrimary();
  };

  const parseAnd = () => {
    let node = parseNot();
    while (match("keyword", "AND")) {
      node = { type: "and", left: node, right: parseNot() };
    }
    return node;
  };

  const parseOr = () => {
    let node = parseAnd();
    while (match("keyword", "OR")) {
      node = { type: "or", left: node, right: parseAnd() };
    }
    return node;
  };

  const ast = parseOr();
  if (pos < tokens.length) {
    const token = tokens[pos];
    throw new Error(`Unexpected token "${token?.value ?? ""}" near the end of WHERE clause.`);
  }
  return ast;
}

function escapeSelectAttributesRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function compareSelectAttributesValues(left, right, op) {
  const leftIsNull = left === null || left === undefined || left === "";
  const rightIsNull = right === null || right === undefined || right === "";
  if (op === "=" || op === "==" || op === "!=" || op === "<>") {
    if (leftIsNull || rightIsNull) {
      const sameNull = leftIsNull && rightIsNull;
      return (op === "=" || op === "==") ? sameNull : !sameNull;
    }
    const leftNum = Number(left);
    const rightNum = Number(right);
    if (Number.isFinite(leftNum) && Number.isFinite(rightNum)) {
      const eq = leftNum === rightNum;
      return (op === "=" || op === "==") ? eq : !eq;
    }
    const leftText = String(left).toLowerCase();
    const rightText = String(right).toLowerCase();
    const eq = leftText === rightText;
    return (op === "=" || op === "==") ? eq : !eq;
  }

  const leftNum = Number(left);
  const rightNum = Number(right);
  if (Number.isFinite(leftNum) && Number.isFinite(rightNum)) {
    if (op === ">") return leftNum > rightNum;
    if (op === ">=") return leftNum >= rightNum;
    if (op === "<") return leftNum < rightNum;
    if (op === "<=") return leftNum <= rightNum;
    return false;
  }

  const cmp = String(left ?? "").localeCompare(String(right ?? ""), undefined, { numeric: true, sensitivity: "base" });
  if (op === ">") return cmp > 0;
  if (op === ">=") return cmp >= 0;
  if (op === "<") return cmp < 0;
  if (op === "<=") return cmp <= 0;
  return false;
}

function evaluateSelectAttributesAst(node, row) {
  if (!node) return true;
  if (node.type === "all") return true;
  if (node.type === "or") return evaluateSelectAttributesAst(node.left, row) || evaluateSelectAttributesAst(node.right, row);
  if (node.type === "and") return evaluateSelectAttributesAst(node.left, row) && evaluateSelectAttributesAst(node.right, row);
  if (node.type === "not") return !evaluateSelectAttributesAst(node.value, row);

  if (node.type === "is_null") {
    const value = getSelectAttributesRowFieldValue(row, node.field);
    return value === null || value === undefined || value === "";
  }
  if (node.type === "is_not_null") {
    const value = getSelectAttributesRowFieldValue(row, node.field);
    return !(value === null || value === undefined || value === "");
  }

  if (node.type === "compare") {
    const left = getSelectAttributesRowFieldValue(row, node.field);
    return compareSelectAttributesValues(left, node.value, node.op);
  }

  if (node.type === "in") {
    const left = getSelectAttributesRowFieldValue(row, node.field);
    return node.values.some(value => compareSelectAttributesValues(left, value, "="));
  }

  if (node.type === "like") {
    const leftText = String(getSelectAttributesRowFieldValue(row, node.field) ?? "");
    const pattern = String(node.value ?? "");
    const regex = new RegExp(
      `^${escapeSelectAttributesRegex(pattern).replace(/%/g, ".*").replace(/_/g, ".")}$`,
      "i"
    );
    return regex.test(leftText);
  }

  return false;
}

function buildSelectAttributesPredicate(whereClause) {
  const normalized = normalizeSelectAttributesWhereText(whereClause);
  if (!normalized) return () => true;
  const ast = parseSelectAttributesWhereClause(normalized);
  return row => evaluateSelectAttributesAst(ast, row || {});
}

function getSelectAttributesUniqueValues(sourceKey, fieldName) {
  const dataset = getSelectAttributesDataset(sourceKey);
  const uniqueMap = new Map();
  const field = String(fieldName || "").trim();
  if (!field) return [];

  dataset.forEach(entry => {
    const rawValue = getSelectAttributesRowFieldValue(entry?.row || {}, field);
    const type = rawValue === null || rawValue === undefined ? "null" : typeof rawValue;
    const key = `${type}::${String(rawValue)}`;
    if (!uniqueMap.has(key)) uniqueMap.set(key, rawValue);
  });

  const values = [...uniqueMap.values()];
  values.sort((a, b) => {
    const an = Number(a);
    const bn = Number(b);
    if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
    if (a === null || a === undefined) return -1;
    if (b === null || b === undefined) return 1;
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" });
  });
  return values.slice(0, SELECT_ATTRIBUTES_MAX_UNIQUE_VALUES);
}

function initSelectByAttributesControls() {
  const modal = document.getElementById("selectByAttributesModal");
  const openBtn = document.getElementById("selectByAttributesBtn");
  const closeBtn = document.getElementById("selectAttributesCloseBtn");
  const runBtn = document.getElementById("selectAttributesRunBtn");
  const sourceSelect = document.getElementById("selectAttributesSource");
  const fieldSelect = document.getElementById("selectAttributesField");
  const valueInput = document.getElementById("selectAttributesValue");
  const whereArea = document.getElementById("selectAttributesWhereClause");
  const previewNode = document.getElementById("selectAttributesQueryPreview");
  const statusNode = document.getElementById("selectAttributesStatus");
  const insertFieldBtn = document.getElementById("selectAttributesInsertFieldBtn");
  const insertValueBtn = document.getElementById("selectAttributesInsertValueBtn");
  const refreshFieldsBtn = document.getElementById("selectAttributesRefreshFieldsBtn");
  const uniqueValuesBtn = document.getElementById("selectAttributesUniqueValuesBtn");
  const formatBtn = document.getElementById("selectAttributesFormatBtn");
  const clearBtn = document.getElementById("selectAttributesClearBtn");
  const saveBtn = document.getElementById("selectAttributesSaveBtn");
  const loadBtn = document.getElementById("selectAttributesLoadBtn");
  const deleteBtn = document.getElementById("selectAttributesDeleteBtn");
  const queryNameInput = document.getElementById("selectAttributesQueryName");
  const savedList = document.getElementById("selectAttributesSavedList");
  const tokenBar = document.getElementById("selectAttributesTokenBar");

  const uniqueModal = document.getElementById("selectAttributesUniqueModal");
  const uniqueCloseBtn = document.getElementById("selectAttributesUniqueCloseBtn");
  const uniqueMeta = document.getElementById("selectAttributesUniqueMeta");
  const uniqueSearch = document.getElementById("selectAttributesUniqueSearch");
  const uniqueList = document.getElementById("selectAttributesUniqueList");

  if (!modal || !openBtn || !closeBtn || !runBtn || !sourceSelect || !fieldSelect || !whereArea || !previewNode || !statusNode) {
    return;
  }

  const state = {
    savedQueries: getSelectAttributesSavedQueries(),
    uniqueValues: [],
    uniqueField: "",
    uniqueSource: "records"
  };

  const setStatus = (message, type = "") => {
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (type === "error") statusNode.classList.add("error");
    if (type === "success") statusNode.classList.add("success");
  };

  const getWhereClause = () => normalizeSelectAttributesWhereText(whereArea.value || "");

  const updatePreview = () => {
    const sourceLabel = getSelectAttributesSourceLabel(sourceSelect.value);
    const whereClause = getWhereClause();
    previewNode.textContent = whereClause
      ? `SELECT * FROM ${sourceLabel} WHERE ${whereClause}`
      : `SELECT * FROM ${sourceLabel}`;
  };

  const renderSavedList = (selectId = "") => {
    if (!savedList) return;
    savedList.innerHTML = "";
    if (!state.savedQueries.length) {
      const empty = document.createElement("option");
      empty.value = "";
      empty.textContent = "No saved queries";
      savedList.appendChild(empty);
      return;
    }
    const sorted = [...state.savedQueries].sort((a, b) => Number(b.updatedAt || 0) - Number(a.updatedAt || 0));
    sorted.forEach(query => {
      const option = document.createElement("option");
      option.value = query.id;
      option.textContent = `${query.name} (${getSelectAttributesSourceLabel(query.source)})`;
      savedList.appendChild(option);
    });
    if (selectId) savedList.value = selectId;
  };

  const refreshFields = (preferredField = "") => {
    const dataset = getSelectAttributesDataset(sourceSelect.value);
    const fields = getSelectAttributesFieldNames(sourceSelect.value, dataset);
    fieldSelect.innerHTML = "";
    fields.forEach(field => {
      const option = document.createElement("option");
      option.value = field;
      option.textContent = field;
      fieldSelect.appendChild(option);
    });
    if (preferredField && fields.includes(preferredField)) {
      fieldSelect.value = preferredField;
    }
    updatePreview();
  };

  const setWhereClause = next => {
    whereArea.value = normalizeSelectAttributesWhereText(next);
    updatePreview();
  };

  const appendToken = token => {
    const value = String(token || "").trim();
    if (!value) return;
    const current = getWhereClause();
    setWhereClause(current ? `${current} ${value}` : value);
    whereArea.focus();
  };

  const renderUniqueValuesList = () => {
    if (!uniqueList) return;
    const filter = String(uniqueSearch?.value || "").trim().toLowerCase();
    const filtered = state.uniqueValues.filter(value => {
      if (!filter) return true;
      if (value === null || value === undefined) return "null".includes(filter);
      return String(value).toLowerCase().includes(filter);
    });

    uniqueList.innerHTML = "";
    if (!filtered.length) {
      uniqueList.innerHTML = '<div class="select-attributes-unique-empty">No values match the current filter.</div>';
      return;
    }

    filtered.forEach(value => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "select-attributes-unique-item";
      const label = value === null || value === undefined
        ? "NULL"
        : String(value).length
          ? String(value)
          : "(empty string)";
      button.textContent = label;
      button.addEventListener("click", () => {
        appendToken(formatSelectAttributesSqlLiteral(value));
      });
      uniqueList.appendChild(button);
    });
  };

  const openUniqueModalForField = () => {
    if (!uniqueModal || !uniqueMeta) return;
    const whereField = detectLastSelectAttributesField(getWhereClause());
    const field = whereField || String(fieldSelect.value || "").trim();
    if (!field) {
      setStatus("Pick or insert a field before fetching unique values.", "error");
      return;
    }
    const values = getSelectAttributesUniqueValues(sourceSelect.value, field);
    state.uniqueValues = values;
    state.uniqueField = field;
    state.uniqueSource = sourceSelect.value;
    uniqueMeta.textContent = `${values.length.toLocaleString()} unique values for [${field}] from ${getSelectAttributesSourceLabel(sourceSelect.value)}.`;
    if (uniqueSearch) uniqueSearch.value = "";
    renderUniqueValuesList();
    uniqueModal.style.display = "flex";
  };

  const closeUniqueModal = () => {
    if (uniqueModal) uniqueModal.style.display = "none";
  };

  const runQuery = () => {
    const source = String(sourceSelect.value || "records");
    const whereClause = getWhereClause();
    const dataset = getSelectAttributesDataset(source);
    if (!dataset.length) {
      setStatus(`No rows available in ${getSelectAttributesSourceLabel(source)}.`, "error");
      return;
    }

    let predicate = () => true;
    try {
      predicate = buildSelectAttributesPredicate(whereClause);
    } catch (err) {
      setStatus(`Query error: ${String(err?.message || err)}`, "error");
      return;
    }

    const matched = dataset.filter(entry => {
      try {
        return predicate(entry?.row || {});
      } catch {
        return false;
      }
    });

    attributeState.selectedRowIds.clear();
    applyAttributeSelectionStyles();
    streetAttributeSelectedIds.clear();
    applyStreetSelectionStyles();
    syncSelectedStopsHeaderCount(0);

    if (source === "streets") {
      const ids = matched
        .map(item => Number(item?.streetId))
        .filter(id => Number.isFinite(id));
      ids.forEach(id => streetAttributeSelectedIds.add(id));
      applyStreetSelectionStyles();
      syncSelectedStopsHeaderCount(streetAttributeSelectedIds.size);
      setAttributeTableMode("streets");
      openAttributePanel();
      renderAttributeTable();
      setStatus(`Selected ${ids.length.toLocaleString()} street segments.`, "success");
      return;
    }

    const rowIds = matched
      .map(item => Number(item?.rowId))
      .filter(id => Number.isFinite(id));
    setAttributeTableMode("records");
    window.setAttributeSelectedRowIds(rowIds);
    openAttributePanel();
    renderAttributeTable();
    setStatus(`Selected ${rowIds.length.toLocaleString()} records from ${getSelectAttributesSourceLabel(source)}.`, "success");
  };

  const saveQuery = () => {
    const source = String(sourceSelect.value || "records");
    const whereClause = getWhereClause();
    if (!whereClause) {
      setStatus("Enter a WHERE clause before saving a query.", "error");
      return;
    }
    const typedName = String(queryNameInput?.value || "").trim();
    const autoName = `Query ${new Date().toLocaleString()}`;
    const name = typedName || autoName;
    const now = Date.now();
    const existingIndex = state.savedQueries.findIndex(query => query.name.toLowerCase() === name.toLowerCase());

    if (existingIndex >= 0) {
      const current = state.savedQueries[existingIndex];
      state.savedQueries[existingIndex] = {
        ...current,
        source,
        where: whereClause,
        updatedAt: now
      };
      setSelectAttributesSavedQueries(state.savedQueries);
      renderSavedList(current.id);
      setStatus(`Updated saved query "${name}".`, "success");
      return;
    }

    const id = (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function")
      ? crypto.randomUUID()
      : `q_${now}_${Math.random().toString(36).slice(2, 8)}`;
    state.savedQueries.push({
      id,
      name,
      source,
      where: whereClause,
      updatedAt: now
    });
    setSelectAttributesSavedQueries(state.savedQueries);
    renderSavedList(id);
    if (queryNameInput && !typedName) queryNameInput.value = name;
    setStatus(`Saved query "${name}".`, "success");
  };

  const loadSavedQuery = () => {
    const selectedId = String(savedList?.value || "").trim();
    if (!selectedId) {
      setStatus("Choose a saved query to load.", "error");
      return;
    }
    const query = state.savedQueries.find(item => item.id === selectedId);
    if (!query) {
      setStatus("Saved query not found.", "error");
      return;
    }
    if ([...sourceSelect.options].some(option => option.value === query.source)) {
      sourceSelect.value = query.source;
    } else {
      sourceSelect.value = "records";
    }
    refreshFields(detectLastSelectAttributesField(query.where));
    setWhereClause(query.where);
    if (queryNameInput) queryNameInput.value = query.name;
    setStatus(`Loaded query "${query.name}".`, "success");
  };

  const deleteSavedQuery = () => {
    const selectedId = String(savedList?.value || "").trim();
    if (!selectedId) {
      setStatus("Choose a saved query to delete.", "error");
      return;
    }
    const query = state.savedQueries.find(item => item.id === selectedId);
    if (!query) {
      setStatus("Saved query not found.", "error");
      return;
    }
    state.savedQueries = state.savedQueries.filter(item => item.id !== selectedId);
    setSelectAttributesSavedQueries(state.savedQueries);
    renderSavedList("");
    setStatus(`Deleted query "${query.name}".`, "success");
  };

  const openModal = () => {
    refreshFields(detectLastSelectAttributesField(getWhereClause()));
    updatePreview();
    renderSavedList(savedList?.value || "");
    if (!whereArea.value.trim()) setStatus("Build a WHERE clause, then click Run Query.");
    openBtn.classList.add("active");
    modal.style.display = "flex";
  };

  const closeModal = () => {
    openBtn.classList.remove("active");
    modal.style.display = "none";
    closeUniqueModal();
  };

  openBtn.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });

  sourceSelect.addEventListener("change", () => {
    refreshFields(detectLastSelectAttributesField(getWhereClause()));
    updatePreview();
  });
  whereArea.addEventListener("input", updatePreview);
  refreshFieldsBtn?.addEventListener("click", () => refreshFields(detectLastSelectAttributesField(getWhereClause())));

  insertFieldBtn?.addEventListener("click", () => {
    const field = String(fieldSelect.value || "").trim();
    if (!field) {
      setStatus("Pick a field first.", "error");
      return;
    }
    appendToken(`[${field}]`);
  });

  insertValueBtn?.addEventListener("click", () => {
    const literal = normalizeSelectAttributesInputLiteral(valueInput?.value || "");
    if (!literal) {
      setStatus("Enter a value first.", "error");
      return;
    }
    appendToken(literal);
  });

  valueInput?.addEventListener("keydown", event => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    insertValueBtn?.click();
  });

  tokenBar?.addEventListener("click", event => {
    const button = event.target.closest("button[data-token]");
    if (!button) return;
    const token = button.getAttribute("data-token");
    appendToken(token);
  });

  uniqueValuesBtn?.addEventListener("click", openUniqueModalForField);
  formatBtn?.addEventListener("click", () => setWhereClause(whereArea.value || ""));
  clearBtn?.addEventListener("click", () => {
    setWhereClause("");
    setStatus("Query cleared.");
  });

  saveBtn?.addEventListener("click", saveQuery);
  loadBtn?.addEventListener("click", loadSavedQuery);
  deleteBtn?.addEventListener("click", deleteSavedQuery);
  runBtn.addEventListener("click", runQuery);

  uniqueCloseBtn?.addEventListener("click", closeUniqueModal);
  uniqueModal?.addEventListener("click", event => {
    if (event.target === uniqueModal) closeUniqueModal();
  });
  uniqueSearch?.addEventListener("input", renderUniqueValuesList);

  renderSavedList("");
  refreshFields("");
  updatePreview();
}

function escapePrintHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatPrintNumber(value, maximumFractionDigits = 2) {
  const n = Number(value);
  if (!Number.isFinite(n)) return "0";
  return n.toLocaleString(undefined, { maximumFractionDigits });
}

function getCurrentRouteFileNameForPrint() {
  const node = document.getElementById("currentFileName");
  const name = String(node?.textContent || "").trim();
  return name || "None";
}

function getSelectedBasemapLabelForPrint() {
  const select = document.getElementById("baseMapSelect");
  if (!select) return "Unknown";
  const option = select.options?.[select.selectedIndex];
  return String(option?.textContent || option?.value || "Unknown").trim();
}

function getVisibleRouteDayLegendItemsForPrint() {
  return Object.entries(routeDayGroups)
    .map(([key, group]) => {
      const layers = Array.isArray(group?.layers) ? group.layers : [];
      const visibleCount = layers.reduce((count, layer) => count + (map.hasLayer(layer) ? 1 : 0), 0);
      if (!visibleCount) return null;
      const [route = "", dayRaw = ""] = String(key || "").split("|");
      const numericDay = Number(dayRaw);
      const dayLabel = Number.isFinite(numericDay) ? (dayName(numericDay) || String(dayRaw)) : String(dayRaw || "Unknown");
      const symbol = symbolMap[key] || getSymbol(key);
      return {
        key,
        route: String(route || "Unassigned"),
        day: dayLabel,
        color: symbol?.color || "#2f89df",
        shape: symbol?.shape || "circle",
        visibleCount
      };
    })
    .filter(Boolean)
    .sort((a, b) => {
      const routeCmp = String(a.route).localeCompare(String(b.route), undefined, { numeric: true, sensitivity: "base" });
      if (routeCmp !== 0) return routeCmp;
      return String(a.day).localeCompare(String(b.day), undefined, { numeric: true, sensitivity: "base" });
    });
}

function buildMapLegendHtmlForPrint() {
  const basemapLabel = getSelectedBasemapLabelForPrint();
  const routeFileName = getCurrentRouteFileNameForPrint();
  const routeDayItems = getVisibleRouteDayLegendItemsForPrint();
  const shownItems = routeDayItems.slice(0, 42);
  const hiddenCount = Math.max(0, routeDayItems.length - shownItems.length);
  const streetLayerVisible = map.hasLayer(streetAttributeLayerGroup);
  const streetSegmentsLoaded = streetAttributeById.size;

  const itemRows = shownItems.length
    ? shownItems.map(item => `
      <div class="map-print-legend-row">
        <span class="map-print-legend-swatch" style="background:${escapePrintHtml(item.color)};${item.shape === "square" ? "border-radius:2px;" : ""}${item.shape === "diamond" ? "transform:rotate(45deg);" : ""}"></span>
        <span>Route ${escapePrintHtml(item.route)} - ${escapePrintHtml(item.day)} (${item.visibleCount.toLocaleString()})</span>
      </div>
    `).join("")
    : '<div class="map-print-legend-row map-print-legend-muted">No route/day layers currently visible.</div>';

  return `
    <div class="map-print-legend-title">Map Legend</div>
    <p class="map-print-legend-meta">File: ${escapePrintHtml(routeFileName)}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}</p>
    <div class="map-print-legend-section-title">Basemap</div>
    <div class="map-print-legend-row"><span class="map-print-legend-swatch" style="background:#6487ac;"></span><span>${escapePrintHtml(basemapLabel)}</span></div>
    <div class="map-print-legend-section-title">Street Network</div>
    <div class="map-print-legend-row"><span class="map-print-legend-swatch" style="background:${streetLayerVisible ? "#4ea2f5" : "#919aa5"};"></span><span>${streetLayerVisible ? "Visible" : "Hidden"} (${streetSegmentsLoaded.toLocaleString()} loaded)</span></div>
    <div class="map-print-legend-section-title">Visible Route + Day Layers (${routeDayItems.length.toLocaleString()})</div>
    ${itemRows}
    ${hiddenCount ? `<div class="map-print-legend-row map-print-legend-muted">+${hiddenCount.toLocaleString()} more layers not listed</div>` : ""}
  `;
}

function ensureMapPrintLegendPanel() {
  if (!document.body) return null;
  let panel = document.getElementById("mapPrintLegendPanel");
  if (!panel) {
    panel = document.createElement("div");
    panel.id = "mapPrintLegendPanel";
    panel.className = "map-print-legend-panel";
    document.body.appendChild(panel);
  } else if (panel.parentElement !== document.body) {
    document.body.appendChild(panel);
  }
  return panel;
}

const PRINT_MAP_SPAN_OPTIONS = new Set(["full", "wide", "standard", "compact", "legend-right", "legend-left"]);
const PRINT_LEGEND_POSITION_OPTIONS = new Set(["top-left", "top-right", "bottom-left", "bottom-right", "outside-right", "outside-left"]);
const PRINT_OUTPUT_SCALE_OPTIONS = new Set(["1", "0.95", "0.9", "0.85"]);
const PRINT_RESOLUTION_SCALE_OPTIONS = new Set(["1", "1.5", "2"]);

function runMapPrintFlow({
  withLegend = false,
  legendPosition = "outside-right",
  mapSpan = "legend-right",
  outputScale = 1,
  resolutionScale = 1,
  customLegendLeftPx = NaN,
  customLegendTopPx = NaN,
  customLegendLeftRatio = NaN,
  customLegendTopRatio = NaN,
  customMapLeftPx = NaN,
  customMapTopPx = NaN,
  customMapWidthPx = NaN,
  customMapHeightPx = NaN,
  customMapLeftRatio = NaN,
  customMapTopRatio = NaN,
  customMapWidthRatio = NaN,
  customMapHeightRatio = NaN
} = {}) {
  const mapNode = document.getElementById("map");
  if (!mapNode) {
    alert("Map is not available to print.");
    return;
  }

  const finalLegendPosition = PRINT_LEGEND_POSITION_OPTIONS.has(String(legendPosition))
    ? String(legendPosition)
    : "outside-right";
  const finalMapSpan = PRINT_MAP_SPAN_OPTIONS.has(String(mapSpan))
    ? String(mapSpan)
    : "legend-right";
  const clampValue = (value, min, max) => Math.max(min, Math.min(max, value));
  const normalizedOutputScale = Number.isFinite(Number(outputScale))
    ? clampValue(Number(outputScale), 0.55, 1)
    : 1;
  const normalizedResolutionScale = Number.isFinite(Number(resolutionScale))
    ? clampValue(Number(resolutionScale), 1, 2)
    : 1;
  const hasCustomMapRect =
    Number.isFinite(customMapLeftPx) &&
    Number.isFinite(customMapTopPx) &&
    Number.isFinite(customMapWidthPx) &&
    Number.isFinite(customMapHeightPx) &&
    customMapWidthPx > 120 &&
    customMapHeightPx > 120;
  const hasCustomMapRatios =
    Number.isFinite(customMapLeftRatio) &&
    Number.isFinite(customMapTopRatio) &&
    Number.isFinite(customMapWidthRatio) &&
    Number.isFinite(customMapHeightRatio) &&
    customMapWidthRatio > 0.08 &&
    customMapHeightRatio > 0.08;

  let legendPanel = null;
  if (withLegend) {
    legendPanel = ensureMapPrintLegendPanel();
    if (legendPanel) {
      legendPanel.innerHTML = buildMapLegendHtmlForPrint();
      const hasCustomLegendRatios = Number.isFinite(customLegendLeftRatio) && Number.isFinite(customLegendTopRatio);
      if (hasCustomLegendRatios) {
        const legendRect = legendPanel.getBoundingClientRect();
        const legendWidthRatio = window.innerWidth > 0 ? (Math.max(120, legendRect.width || 260) / window.innerWidth) : 0.24;
        const legendHeightRatio = window.innerHeight > 0 ? (Math.max(80, legendRect.height || 220) / window.innerHeight) : 0.22;
        let safeLeftRatio = clampValue(Number(customLegendLeftRatio), 0, 1);
        let safeTopRatio = clampValue(Number(customLegendTopRatio), 0, 1);
        if ((safeLeftRatio + legendWidthRatio) > 1) {
          safeLeftRatio = Math.max(0, 1 - legendWidthRatio - 0.004);
        }
        if ((safeTopRatio + legendHeightRatio) > 1) {
          safeTopRatio = Math.max(0, 1 - legendHeightRatio - 0.004);
        }
        legendPanel.style.left = `${(safeLeftRatio * 100).toFixed(4)}vw`;
        legendPanel.style.top = `${(safeTopRatio * 100).toFixed(4)}vh`;
        legendPanel.style.right = "auto";
        legendPanel.style.bottom = "auto";
      } else {
      const hasCustomLegendPosition = Number.isFinite(customLegendLeftPx) && Number.isFinite(customLegendTopPx);
      if (hasCustomLegendPosition) {
        const legendRect = legendPanel.getBoundingClientRect();
        const legendW = Math.max(120, legendRect.width || 260);
        const legendH = Math.max(80, legendRect.height || 220);
        const safeLeft = Math.max(8, Math.min(Math.round(customLegendLeftPx), Math.max(8, window.innerWidth - legendW - 8)));
        const safeTop = Math.max(8, Math.min(Math.round(customLegendTopPx), Math.max(8, window.innerHeight - legendH - 8)));
        legendPanel.style.left = `${safeLeft}px`;
        legendPanel.style.top = `${safeTop}px`;
        legendPanel.style.right = "auto";
        legendPanel.style.bottom = "auto";
      }
      }
    }
  }

  const previousTitle = document.title;
  document.title = withLegend ? "TDS-PAK Map Print (Legend)" : "TDS-PAK Map Print";
  const pageStyle = document.createElement("style");
  pageStyle.id = "tdsPakMapPrintPageStyle";
  pageStyle.textContent = "@page { size: landscape; margin: 0; }";
  document.head.appendChild(pageStyle);

  let cleaned = false;
  const cleanup = () => {
    if (cleaned) return;
    cleaned = true;
    document.body.classList.remove(
      "print-mode-map",
      "print-mode-map-legend",
      "print-map-span-full",
      "print-map-span-wide",
      "print-map-span-standard",
      "print-map-span-compact",
      "print-map-span-legend-right",
      "print-map-span-legend-left",
      "print-map-span-custom",
      "print-mode-map-legend-pos-top-left",
      "print-mode-map-legend-pos-top-right",
      "print-mode-map-legend-pos-bottom-left",
      "print-mode-map-legend-pos-bottom-right",
      "print-mode-map-legend-pos-outside-right",
      "print-mode-map-legend-pos-outside-left"
    );
    document.body.style.removeProperty("--print-custom-map-left");
    document.body.style.removeProperty("--print-custom-map-top");
    document.body.style.removeProperty("--print-custom-map-width");
    document.body.style.removeProperty("--print-custom-map-height");
    document.body.style.removeProperty("--print-map-render-scale");
    document.body.style.removeProperty("--print-map-render-inverse");
    document.title = previousTitle;
    if (legendPanel) legendPanel.remove();
    if (pageStyle && pageStyle.parentNode) pageStyle.parentNode.removeChild(pageStyle);
    setTimeout(() => {
      try { map.invalidateSize({ pan: false }); } catch (_) {}
    }, 120);
  };

  const cleanupFallbackTimer = window.setTimeout(() => {
    cleanup();
  }, 300000);

  window.addEventListener("afterprint", () => {
    window.clearTimeout(cleanupFallbackTimer);
    cleanup();
  }, { once: true });

  document.body.classList.add("print-mode-map", `print-map-span-${finalMapSpan}`);
  document.body.style.setProperty("--print-map-render-scale", normalizedResolutionScale.toFixed(3));
  document.body.style.setProperty("--print-map-render-inverse", (1 / normalizedResolutionScale).toFixed(6));
  if (hasCustomMapRatios || hasCustomMapRect) {
    document.body.classList.add("print-map-span-custom");
    if (hasCustomMapRatios) {
      let safeLeft = clampValue(Number(customMapLeftRatio), 0, 1);
      let safeTop = clampValue(Number(customMapTopRatio), 0, 1);
      let safeWidth = clampValue(Number(customMapWidthRatio), 0.08, 1);
      let safeHeight = clampValue(Number(customMapHeightRatio), 0.08, 1);
      if (normalizedOutputScale < 0.999) {
        const shrunkWidth = clampValue(safeWidth * normalizedOutputScale, 0.08, 1);
        const shrunkHeight = clampValue(safeHeight * normalizedOutputScale, 0.08, 1);
        safeLeft += (safeWidth - shrunkWidth) * 0.5;
        safeTop += (safeHeight - shrunkHeight) * 0.5;
        safeWidth = shrunkWidth;
        safeHeight = shrunkHeight;
      }
      if ((safeLeft + safeWidth) > 1) safeLeft = Math.max(0, 1 - safeWidth);
      if ((safeTop + safeHeight) > 1) safeTop = Math.max(0, 1 - safeHeight);
      document.body.style.setProperty("--print-custom-map-left", `${(safeLeft * 100).toFixed(4)}vw`);
      document.body.style.setProperty("--print-custom-map-top", `${(safeTop * 100).toFixed(4)}vh`);
      document.body.style.setProperty("--print-custom-map-width", `${(safeWidth * 100).toFixed(4)}vw`);
      document.body.style.setProperty("--print-custom-map-height", `${(safeHeight * 100).toFixed(4)}vh`);
    } else {
      let safeWidthPx = clampValue(Math.round(customMapWidthPx), 120, Math.max(120, window.innerWidth));
      let safeHeightPx = clampValue(Math.round(customMapHeightPx), 120, Math.max(120, window.innerHeight));
      let safeLeftPx = Math.round(customMapLeftPx);
      let safeTopPx = Math.round(customMapTopPx);
      if (normalizedOutputScale < 0.999) {
        const shrunkWidthPx = clampValue(Math.round(safeWidthPx * normalizedOutputScale), 120, Math.max(120, window.innerWidth));
        const shrunkHeightPx = clampValue(Math.round(safeHeightPx * normalizedOutputScale), 120, Math.max(120, window.innerHeight));
        safeLeftPx += Math.round((safeWidthPx - shrunkWidthPx) * 0.5);
        safeTopPx += Math.round((safeHeightPx - shrunkHeightPx) * 0.5);
        safeWidthPx = shrunkWidthPx;
        safeHeightPx = shrunkHeightPx;
      }
      safeLeftPx = clampValue(safeLeftPx, 0, Math.max(0, window.innerWidth - safeWidthPx));
      safeTopPx = clampValue(safeTopPx, 0, Math.max(0, window.innerHeight - safeHeightPx));
      document.body.style.setProperty("--print-custom-map-left", `${safeLeftPx}px`);
      document.body.style.setProperty("--print-custom-map-top", `${safeTopPx}px`);
      document.body.style.setProperty("--print-custom-map-width", `${safeWidthPx}px`);
      document.body.style.setProperty("--print-custom-map-height", `${safeHeightPx}px`);
    }
  }
  if (withLegend) {
    document.body.classList.add("print-mode-map-legend", `print-mode-map-legend-pos-${finalLegendPosition}`);
  }

  try { map.invalidateSize({ pan: false }); } catch (_) {}
  const printDelayMs = clampValue(Math.round(180 + ((normalizedResolutionScale - 1) * 650)), 180, 1300);
  setTimeout(() => {
    try {
      window.print();
    } catch (_) {
      window.clearTimeout(cleanupFallbackTimer);
      cleanup();
    }
  }, printDelayMs);
}

function openPrintDocumentWindow({ title, bodyHtml, extraStyles = "", autoPrint = true }) {
  const win = window.open("", "_blank", "width=1120,height=760,resizable=yes,scrollbars=yes");
  if (!win) return null;

  const autoPrintFlag = autoPrint ? "true" : "false";
  win.document.write(`
    <html>
      <head>
        <title>${escapePrintHtml(title || "Print")}</title>
        <meta charset="UTF-8" />
        <style>
          :root { --bg:#f3f7fb; --panel:#ffffff; --line:#d6e2ee; --head:#e9f2fb; --text:#17334d; --muted:#4f6e8b; --accent:#2f89df; --ok:#1f8a58; --warn:#b06f1e; }
          * { box-sizing: border-box; }
          html, body { margin:0; padding:0; background:var(--bg); color:var(--text); font-family:"Segoe UI", Roboto, Arial, sans-serif; }
          .print-shell { max-width: 1280px; margin: 0 auto; padding: 16px; }
          .print-header { border:1px solid var(--line); background:linear-gradient(180deg,#fafdff 0%,#eef5fd 100%); border-radius:12px; padding:10px 12px; margin-bottom:10px; }
          .print-title { margin:0; font-size:20px; line-height:1.2; }
          .print-meta { margin:5px 0 0; color:var(--muted); font-size:12px; line-height:1.4; }
          .print-section { margin-top: 10px; border:1px solid var(--line); border-radius:12px; background:#fff; overflow:hidden; }
          .print-section h3 { margin:0; padding:10px 12px; font-size:14px; border-bottom:1px solid var(--line); background:#f6faff; }
          .print-section-body { padding:10px 12px; }
          .cards { display:grid; grid-template-columns: repeat(auto-fit,minmax(170px,1fr)); gap:8px; }
          .card { border:1px solid var(--line); border-radius:10px; background:#fbfdff; padding:9px 10px; }
          .card-label { font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.03em; }
          .card-value { margin-top:4px; font-size:20px; font-weight:700; }
          .bar-row { display:grid; grid-template-columns:220px 1fr 110px; gap:8px; align-items:center; margin-bottom:7px; }
          .bar-label { font-size:12px; color:var(--text); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
          .bar-track { height:11px; border-radius:999px; background:#e6eef7; overflow:hidden; }
          .bar-fill { height:100%; background:linear-gradient(90deg,#4ea2f5 0%, var(--accent) 100%); }
          .bar-fill.alt { background:linear-gradient(90deg,#4fc88f 0%, var(--ok) 100%); }
          .bar-fill.warn { background:linear-gradient(90deg,#f0b05f 0%, var(--warn) 100%); }
          .bar-value { text-align:right; font-size:12px; color:var(--muted); font-weight:700; }
          .note { margin:0; color:var(--muted); font-size:12px; line-height:1.4; }
          .list { margin:0; padding-left:18px; }
          .list li { margin:4px 0; font-size:12px; }
          table { width:100%; border-collapse:collapse; font-size:12px; }
          th, td { border:1px solid var(--line); padding:7px 8px; text-align:left; vertical-align:top; white-space:nowrap; }
          th { background:var(--head); font-size:11px; text-transform:uppercase; letter-spacing:.03em; }
          tr:nth-child(even) td { background:#fafcff; }
          .mono { font-family: Consolas, "Courier New", monospace; font-size:12px; }
          @media print {
            html, body { background:#fff; }
            .print-shell { padding:0; max-width:none; }
            .print-section { break-inside: avoid; }
            .print-section table { break-inside:auto; }
            .print-section tr { break-inside: avoid; break-after:auto; }
          }
          ${extraStyles || ""}
        </style>
      </head>
      <body>
        ${bodyHtml}
        <script>
          (function() {
            const shouldPrint = ${autoPrintFlag};
            if (!shouldPrint) return;
            const trigger = function() {
              setTimeout(function() {
                try { window.print(); } catch (_) {}
              }, 260);
            };
            if (document.readyState === "complete") trigger();
            else window.addEventListener("load", trigger, { once: true });
          })();
        <\/script>
      </body>
    </html>
  `);

  win.document.close();
  return win;
}

function buildSummaryTablePrintDocumentHtml() {
  const tableNode = document.getElementById("routeSummaryTable");
  const table = tableNode?.querySelector("table");
  if (!table) return null;
  const tableHtml = table.outerHTML;
  const routeFileName = getCurrentRouteFileNameForPrint();
  return `
    <div class="print-shell">
      <div class="print-header">
        <h1 class="print-title">Route Summary Table</h1>
        <p class="print-meta">Route file: ${escapePrintHtml(routeFileName)}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}</p>
      </div>
      <div class="print-section">
        <h3>Summary Rows</h3>
        <div class="print-section-body">${tableHtml}</div>
      </div>
    </div>
  `;
}

function extractSummaryAnalyticsForPrint(rows, headers) {
  const normalize = value => String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const normalizedHeaders = headers.map(h => ({ original: h, norm: normalize(h) }));

  const findHeader = candidates => {
    const norms = candidates.map(normalize);
    const direct = normalizedHeaders.find(h => norms.includes(h.norm));
    if (direct) return direct.original;
    const fuzzy = normalizedHeaders.find(h => norms.some(n => h.norm.includes(n) || n.includes(h.norm)));
    return fuzzy ? fuzzy.original : null;
  };

  const toNumber = value => {
    if (value === null || value === undefined) return NaN;
    const cleaned = String(value).replace(/,/g, "").trim();
    if (!cleaned) return NaN;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : NaN;
  };

  const toHours = value => {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number" && Number.isFinite(value)) {
      if (value > 0 && value < 1) return value * 24;
      return value;
    }
    const text = String(value).trim();
    if (!text) return NaN;
    const direct = toNumber(text);
    if (Number.isFinite(direct)) return direct > 0 && direct < 1 ? direct * 24 : direct;
    const colonMatch = text.match(/^(-?\d+):(\d{1,2})(?::(\d{1,2}))?$/);
    if (colonMatch) {
      const h = Number(colonMatch[1]) || 0;
      const m = Number(colonMatch[2]) || 0;
      const s = Number(colonMatch[3] || 0) || 0;
      return h + (m / 60) + (s / 3600);
    }
    const lower = text.toLowerCase();
    const hMatch = lower.match(/(-?\d+(?:\.\d+)?)\s*(h|hr|hrs|hour|hours)/);
    const mMatch = lower.match(/(-?\d+(?:\.\d+)?)\s*(m|min|mins|minute|minutes)/);
    if (hMatch || mMatch) {
      const h = hMatch ? Number(hMatch[1]) : 0;
      const m = mMatch ? Number(mMatch[1]) : 0;
      return (Number.isFinite(h) ? h : 0) + ((Number.isFinite(m) ? m : 0) / 60);
    }
    return NaN;
  };

  const fields = {
    route: findHeader(["route", "newroute", "routeid", "rte"]),
    day: findHeader(["day", "newday", "routeday", "dispatchday"]),
    stops: findHeader(["totalstops", "stops", "stopcount", "numberofstops"]),
    miles: findHeader(["miles", "totalmiles", "distancemiles", "distance"]),
    demand: findHeader(["demand", "totaldemand", "volume", "load"]),
    trips: findHeader(["numberoftrips", "trips", "tripcount", "totaltrips"]),
    totalTime: findHeader(["totaltime", "totalroutetime", "routetime", "hours"])
  };

  const routeDayMap = new Map();
  rows.forEach(row => {
    const route = String(fields.route ? row?.[fields.route] : "").trim() || "Unknown Route";
    const day = String(fields.day ? row?.[fields.day] : "").trim() || "Unknown Day";
    const key = `${route} | ${day}`;
    if (!routeDayMap.has(key)) {
      routeDayMap.set(key, { route, day, routeDay: key, stops: 0, miles: 0, demand: 0, trips: 0, totalTime: 0 });
    }
    const bucket = routeDayMap.get(key);
    const stops = toNumber(fields.stops ? row?.[fields.stops] : "");
    const miles = toNumber(fields.miles ? row?.[fields.miles] : "");
    const demand = toNumber(fields.demand ? row?.[fields.demand] : "");
    const trips = toNumber(fields.trips ? row?.[fields.trips] : "");
    const totalTime = toHours(fields.totalTime ? row?.[fields.totalTime] : "");
    bucket.stops += Number.isFinite(stops) ? stops : 0;
    bucket.miles += Number.isFinite(miles) ? miles : 0;
    bucket.demand += Number.isFinite(demand) ? demand : 0;
    bucket.trips += Number.isFinite(trips) ? trips : 0;
    bucket.totalTime += Number.isFinite(totalTime) ? totalTime : 0;
  });

  const routeDayRows = [...routeDayMap.values()];
  const dayTotalsMap = new Map();
  routeDayRows.forEach(row => {
    const day = row.day || "Unknown Day";
    if (!dayTotalsMap.has(day)) {
      dayTotalsMap.set(day, { day, demand: 0, miles: 0, stops: 0, trips: 0, routeDayCount: 0 });
    }
    const dayBucket = dayTotalsMap.get(day);
    dayBucket.routeDayCount += 1;
    dayBucket.demand += row.demand;
    dayBucket.miles += row.miles;
    dayBucket.stops += row.stops;
    dayBucket.trips += row.trips;
  });

  const dayTotals = [...dayTotalsMap.values()].sort((a, b) => b.routeDayCount - a.routeDayCount);
  const totalStops = routeDayRows.reduce((sum, row) => sum + row.stops, 0);
  const totalMiles = routeDayRows.reduce((sum, row) => sum + row.miles, 0);
  const totalDemand = routeDayRows.reduce((sum, row) => sum + row.demand, 0);
  const totalTrips = routeDayRows.reduce((sum, row) => sum + row.trips, 0);
  const totalTime = routeDayRows.reduce((sum, row) => sum + row.totalTime, 0);
  const avgTime = routeDayRows.length ? (totalTime / routeDayRows.length) : 0;
  const overTarget = routeDayRows.filter(row => row.totalTime > 11).length;

  return {
    fields,
    routeDayRows,
    dayTotals,
    totals: {
      totalStops,
      totalMiles,
      totalDemand,
      totalTrips,
      avgTime,
      overTarget,
      routeDayCount: routeDayRows.length
    }
  };
}

function buildBarsHtmlForPrint(items, labelKey, valueKey, fillClass = "") {
  if (!items.length) {
    return '<p class="note">No data available.</p>';
  }
  const max = Math.max(1, ...items.map(item => Number(item?.[valueKey]) || 0));
  return items
    .slice()
    .sort((a, b) => (Number(b?.[valueKey]) || 0) - (Number(a?.[valueKey]) || 0))
    .slice(0, 18)
    .map(item => {
      const label = String(item?.[labelKey] ?? "");
      const value = Number(item?.[valueKey]) || 0;
      const width = Math.max(0, Math.min(100, (value / max) * 100));
      return `
        <div class="bar-row">
          <div class="bar-label">${escapePrintHtml(label)}</div>
          <div class="bar-track"><div class="bar-fill ${fillClass}" style="width:${width}%"></div></div>
          <div class="bar-value">${formatPrintNumber(value)}</div>
        </div>
      `;
    })
    .join("");
}

function buildSummaryVisualizationPrintDocumentHtml() {
  const rows = Array.isArray(window._summaryRows) ? window._summaryRows : [];
  const headers = Array.isArray(window._summaryHeaders) ? window._summaryHeaders : [];
  if (!rows.length || !headers.length) return null;

  const routeFileName = getCurrentRouteFileNameForPrint();
  const analytics = extractSummaryAnalyticsForPrint(rows, headers);
  const dayDemandBars = buildBarsHtmlForPrint(analytics.dayTotals, "day", "demand");
  const dayMilesBars = buildBarsHtmlForPrint(analytics.dayTotals, "day", "miles", "alt");
  const dayStopsBars = buildBarsHtmlForPrint(analytics.dayTotals, "day", "stops", "warn");

  const detailRows = analytics.routeDayRows
    .slice()
    .sort((a, b) => (String(a.day).localeCompare(String(b.day), undefined, { numeric: true, sensitivity: "base" }) || String(a.route).localeCompare(String(b.route), undefined, { numeric: true, sensitivity: "base" })))
    .slice(0, 360)
    .map((row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapePrintHtml(row.route)}</td>
        <td>${escapePrintHtml(row.day)}</td>
        <td>${formatPrintNumber(row.stops)}</td>
        <td>${formatPrintNumber(row.miles)}</td>
        <td>${formatPrintNumber(row.demand)}</td>
        <td>${formatPrintNumber(row.trips)}</td>
        <td>${formatPrintNumber(row.totalTime)}</td>
      </tr>
    `)
    .join("");

  return `
    <div class="print-shell">
      <div class="print-header">
        <h1 class="print-title">Route Summary Visualization Report</h1>
        <p class="print-meta">Route file: ${escapePrintHtml(routeFileName)}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}<br>Summary rows scanned: ${rows.length.toLocaleString()}</p>
      </div>
      <section class="print-section">
        <h3>KPI Snapshot</h3>
        <div class="print-section-body">
          <div class="cards">
            <div class="card"><div class="card-label">Route + Day Units</div><div class="card-value">${analytics.totals.routeDayCount.toLocaleString()}</div></div>
            <div class="card"><div class="card-label">Total Stops</div><div class="card-value">${formatPrintNumber(analytics.totals.totalStops)}</div></div>
            <div class="card"><div class="card-label">Total Miles</div><div class="card-value">${formatPrintNumber(analytics.totals.totalMiles)}</div></div>
            <div class="card"><div class="card-label">Total Demand</div><div class="card-value">${formatPrintNumber(analytics.totals.totalDemand)}</div></div>
            <div class="card"><div class="card-label">Total Trips</div><div class="card-value">${formatPrintNumber(analytics.totals.totalTrips)}</div></div>
            <div class="card"><div class="card-label">Avg Total Time</div><div class="card-value">${formatPrintNumber(analytics.totals.avgTime)}h</div></div>
            <div class="card"><div class="card-label">Over 11 Hours</div><div class="card-value">${analytics.totals.overTarget.toLocaleString()}</div></div>
          </div>
        </div>
      </section>
      <section class="print-section">
        <h3>Demand by Day</h3>
        <div class="print-section-body">${dayDemandBars}</div>
      </section>
      <section class="print-section">
        <h3>Miles by Day</h3>
        <div class="print-section-body">${dayMilesBars}</div>
      </section>
      <section class="print-section">
        <h3>Stops by Day</h3>
        <div class="print-section-body">${dayStopsBars}</div>
      </section>
      <section class="print-section">
        <h3>Route + Day Detail</h3>
        <div class="print-section-body">
          <p class="note">Showing ${Math.min(analytics.routeDayRows.length, 360).toLocaleString()} of ${analytics.routeDayRows.length.toLocaleString()} grouped route/day rows.</p>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Route</th>
                <th>Day</th>
                <th>Stops</th>
                <th>Miles</th>
                <th>Demand</th>
                <th>Trips</th>
                <th>Total Time (h)</th>
              </tr>
            </thead>
            <tbody>
              ${detailRows || '<tr><td colspan="8">No grouped rows available.</td></tr>'}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  `;
}

function buildAttributeTablePrintDocumentHtml() {
  if (attributeTableMode === "streets") {
    const rows = getFilteredStreetAttributeRows();
    if (!rows.length) return null;
    const headers = ["id", "name", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"];
    const maxRows = 4200;
    const slice = rows.slice(0, maxRows);
    const tableRows = slice.map((row, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td>${streetAttributeSelectedIds.has(Number(row?.id)) ? "Yes" : ""}</td>
        ${headers.map(key => `<td>${escapePrintHtml(row?.[key] ?? "")}</td>`).join("")}
      </tr>
    `).join("");

    return `
      <div class="print-shell">
        <div class="print-header">
          <h1 class="print-title">Street Attribute Table</h1>
          <p class="print-meta">Route file: ${escapePrintHtml(getCurrentRouteFileNameForPrint())}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}</p>
        </div>
        <section class="print-section">
          <h3>Rows (${rows.length.toLocaleString()} visible${rows.length > maxRows ? `, showing first ${maxRows.toLocaleString()}` : ""})</h3>
          <div class="print-section-body">
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Selected</th>
                  ${headers.map(h => `<th>${escapePrintHtml(h)}</th>`).join("")}
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    `;
  }

  const headers = Array.isArray(window._attributeHeaders) ? window._attributeHeaders : [];
  const rows = getFilteredAttributeRows();
  if (!headers.length || !rows.length) return null;
  const maxRows = 4200;
  const slice = rows.slice(0, maxRows);
  const tableRows = slice.map((item, idx) => `
    <tr>
      <td>${idx + 1}</td>
      <td>${attributeState.selectedRowIds.has(Number(item?.rowId)) ? "Yes" : ""}</td>
      ${headers.map(h => `<td>${escapePrintHtml(item?.row?.[h] ?? "")}</td>`).join("")}
    </tr>
  `).join("");

  return `
    <div class="print-shell">
      <div class="print-header">
        <h1 class="print-title">Record Attribute Table</h1>
        <p class="print-meta">Route file: ${escapePrintHtml(getCurrentRouteFileNameForPrint())}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}</p>
      </div>
      <section class="print-section">
        <h3>Rows (${rows.length.toLocaleString()} visible${rows.length > maxRows ? `, showing first ${maxRows.toLocaleString()}` : ""})</h3>
        <div class="print-section-body">
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Selected</th>
                ${headers.map(h => `<th>${escapePrintHtml(h)}</th>`).join("")}
              </tr>
            </thead>
            <tbody>
              ${tableRows}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  `;
}

function buildSelectionReportPrintDocumentHtml() {
  const selectedRecordIds = [...attributeState.selectedRowIds].filter(id => Number.isFinite(id)).sort((a, b) => a - b);
  const selectedStreetIds = [...streetAttributeSelectedIds].filter(id => Number.isFinite(id)).sort((a, b) => a - b);
  if (!selectedRecordIds.length && !selectedStreetIds.length) return null;

  const allRows = Array.isArray(window._currentRows) ? window._currentRows : [];
  const headers = Array.isArray(window._attributeHeaders) ? window._attributeHeaders : [];
  const selectedRecordRows = selectedRecordIds
    .map(id => ({ rowId: id, row: allRows[id] }))
    .filter(item => item && item.row && typeof item.row === "object");

  const routeKey = headers.find(h => /route/i.test(String(h || ""))) || null;
  const dayKey = headers.find(h => /day/i.test(String(h || ""))) || null;
  const addressKey = headers.find(h => /(address|csstrt|street|location)/i.test(String(h || ""))) || null;

  const recordRowsHtml = selectedRecordRows.slice(0, 700).map(item => `
    <tr>
      <td>${item.rowId}</td>
      <td>${escapePrintHtml(routeKey ? item.row?.[routeKey] : "")}</td>
      <td>${escapePrintHtml(dayKey ? item.row?.[dayKey] : "")}</td>
      <td>${escapePrintHtml(addressKey ? item.row?.[addressKey] : "")}</td>
    </tr>
  `).join("");

  const selectedStreetRows = selectedStreetIds
    .map(id => streetAttributeById.get(id)?.row)
    .filter(Boolean);

  const streetRowsHtml = selectedStreetRows.slice(0, 700).map(row => `
    <tr>
      <td>${escapePrintHtml(row?.id ?? "")}</td>
      <td>${escapePrintHtml(row?.name ?? "")}</td>
      <td>${escapePrintHtml(row?.highway ?? "")}</td>
      <td>${escapePrintHtml(row?.maxspeed ?? "")}</td>
      <td>${escapePrintHtml(row?.lanes ?? "")}</td>
    </tr>
  `).join("");

  return `
    <div class="print-shell">
      <div class="print-header">
        <h1 class="print-title">Selection Report</h1>
        <p class="print-meta">Route file: ${escapePrintHtml(getCurrentRouteFileNameForPrint())}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}</p>
      </div>

      <section class="print-section">
        <h3>Selection Totals</h3>
        <div class="print-section-body">
          <div class="cards">
            <div class="card"><div class="card-label">Selected Records</div><div class="card-value">${selectedRecordIds.length.toLocaleString()}</div></div>
            <div class="card"><div class="card-label">Selected Street Segments</div><div class="card-value">${selectedStreetIds.length.toLocaleString()}</div></div>
          </div>
        </div>
      </section>

      <section class="print-section">
        <h3>Record Selections ${selectedRecordRows.length > 700 ? `(showing first 700 of ${selectedRecordRows.length.toLocaleString()})` : ""}</h3>
        <div class="print-section-body">
          ${selectedRecordRows.length
            ? `<table>
                <thead><tr><th>Row ID</th><th>Route</th><th>Day</th><th>Address/Location</th></tr></thead>
                <tbody>${recordRowsHtml}</tbody>
              </table>`
            : '<p class="note">No record selections.</p>'}
        </div>
      </section>

      <section class="print-section">
        <h3>Street Segment Selections ${selectedStreetRows.length > 700 ? `(showing first 700 of ${selectedStreetRows.length.toLocaleString()})` : ""}</h3>
        <div class="print-section-body">
          ${selectedStreetRows.length
            ? `<table>
                <thead><tr><th>Way ID</th><th>Name</th><th>Class</th><th>Max Speed</th><th>Lanes</th></tr></thead>
                <tbody>${streetRowsHtml}</tbody>
              </table>`
            : '<p class="note">No street segment selections.</p>'}
        </div>
      </section>
    </div>
  `;
}

function buildOperationalReportPrintDocumentHtml() {
  const routeFileName = getCurrentRouteFileNameForPrint();
  const basemapLabel = getSelectedBasemapLabelForPrint();
  const visibleRouteDayItems = getVisibleRouteDayLegendItemsForPrint();
  const routeSummaryRows = Array.isArray(window._summaryRows) ? window._summaryRows : [];
  const routeSummaryHeaders = Array.isArray(window._summaryHeaders) ? window._summaryHeaders : [];
  const hasSummary = routeSummaryRows.length && routeSummaryHeaders.length;
  const analytics = hasSummary ? extractSummaryAnalyticsForPrint(routeSummaryRows, routeSummaryHeaders) : null;

  const summaryTableRows = hasSummary
    ? routeSummaryRows.slice(0, 280).map((row, idx) => `
      <tr>
        <td>${idx + 1}</td>
        ${routeSummaryHeaders.map(h => `<td>${escapePrintHtml(row?.[h] ?? "")}</td>`).join("")}
      </tr>
    `).join("")
    : "";

  const filterRouteCount = document.querySelectorAll("#routeCheckboxes input:checked").length;
  const filterDayCount = document.querySelectorAll("#dayCheckboxes input:checked").length;
  const selectedCount = attributeTableMode === "streets" ? streetAttributeSelectedIds.size : attributeState.selectedRowIds.size;

  return `
    <div class="print-shell">
      <div class="print-header">
        <h1 class="print-title">Operational Report</h1>
        <p class="print-meta">Route file: ${escapePrintHtml(routeFileName)}<br>Printed: ${escapePrintHtml(new Date().toLocaleString())}</p>
      </div>

      <section class="print-section">
        <h3>Map Context</h3>
        <div class="print-section-body">
          <ul class="list">
            <li>Basemap: ${escapePrintHtml(basemapLabel)}</li>
            <li>Visible route/day layers: ${visibleRouteDayItems.length.toLocaleString()}</li>
            <li>Routes currently checked in filter: ${filterRouteCount.toLocaleString()}</li>
            <li>Days currently checked in filter: ${filterDayCount.toLocaleString()}</li>
            <li>Street network layer: ${map.hasLayer(streetAttributeLayerGroup) ? "Visible" : "Hidden"} (${streetAttributeById.size.toLocaleString()} loaded segments)</li>
            <li>Current selection mode: ${escapePrintHtml(attributeTableMode === "streets" ? "Street Attributes" : "Record Attributes")} (${selectedCount.toLocaleString()} selected)</li>
          </ul>
          <p class="note">For a cartographic sheet, use Print Center -> Map + Legend.</p>
        </div>
      </section>

      <section class="print-section">
        <h3>Summary KPI Snapshot</h3>
        <div class="print-section-body">
          ${analytics
            ? `<div class="cards">
                <div class="card"><div class="card-label">Route + Day Units</div><div class="card-value">${analytics.totals.routeDayCount.toLocaleString()}</div></div>
                <div class="card"><div class="card-label">Total Stops</div><div class="card-value">${formatPrintNumber(analytics.totals.totalStops)}</div></div>
                <div class="card"><div class="card-label">Total Miles</div><div class="card-value">${formatPrintNumber(analytics.totals.totalMiles)}</div></div>
                <div class="card"><div class="card-label">Total Demand</div><div class="card-value">${formatPrintNumber(analytics.totals.totalDemand)}</div></div>
                <div class="card"><div class="card-label">Total Trips</div><div class="card-value">${formatPrintNumber(analytics.totals.totalTrips)}</div></div>
                <div class="card"><div class="card-label">Avg Time</div><div class="card-value">${formatPrintNumber(analytics.totals.avgTime)}h</div></div>
              </div>`
            : '<p class="note">No route summary is currently loaded.</p>'}
        </div>
      </section>

      <section class="print-section">
        <h3>Route Summary Table ${hasSummary ? `(showing first ${Math.min(routeSummaryRows.length, 280).toLocaleString()} of ${routeSummaryRows.length.toLocaleString()})` : ""}</h3>
        <div class="print-section-body">
          ${hasSummary
            ? `<table>
                <thead>
                  <tr><th>#</th>${routeSummaryHeaders.map(h => `<th>${escapePrintHtml(h)}</th>`).join("")}</tr>
                </thead>
                <tbody>${summaryTableRows}</tbody>
              </table>`
            : '<p class="note">No summary table available.</p>'}
        </div>
      </section>
    </div>
  `;
}

function initPrintCenterControls() {
  const modal = document.getElementById("printCenterModal");
  const openBtn = document.getElementById("printCenterBtn");
  const closeBtn = document.getElementById("printCenterClose");
  const statusNode = document.getElementById("printCenterStatus");
  const printMapOnlyBtn = document.getElementById("printMapOnlyBtn");
  const printMapLegendBtn = document.getElementById("printMapLegendBtn");
  const printSummaryTableBtn = document.getElementById("printSummaryTableBtn");
  const printSummaryVizBtn = document.getElementById("printSummaryVizBtn");
  const printAttributeTableBtn = document.getElementById("printAttributeTableBtn");
  const printSelectionReportBtn = document.getElementById("printSelectionReportBtn");
  const printOperationalReportBtn = document.getElementById("printOperationalReportBtn");
  const legendPositionSelect = document.getElementById("printLegendPositionSelect");
  const mapSpanSelect = document.getElementById("printMapSpanSelect");
  const outputScaleSelect = document.getElementById("printOutputScaleSelect");
  const resolutionSelect = document.getElementById("printResolutionSelect");
  const previewOverlay = document.getElementById("mapPrintPreviewOverlay");
  const previewToolbar = document.getElementById("mapPrintPreviewToolbar");
  const previewToolbarHandle = document.getElementById("mapPrintPreviewToolbarHandle");
  const previewPopoutBtn = document.getElementById("mapPrintPreviewPopoutBtn");
  const previewFrame = document.getElementById("mapPrintPreviewFrame");
  const previewLegend = document.getElementById("mapPrintPreviewLegend");
  const previewFrameMoveHandle = document.getElementById("mapPrintPreviewFrameMoveHandle");
  const previewFrameResizeHandle = document.getElementById("mapPrintPreviewFrameResizeHandle");
  const previewMapSpanSelect = document.getElementById("mapPrintPreviewMapSpan");
  const previewLegendPositionSelect = document.getElementById("mapPrintPreviewLegendPosition");
  const previewOutputScaleSelect = document.getElementById("mapPrintPreviewOutputScale");
  const previewResolutionSelect = document.getElementById("mapPrintPreviewResolution");
  const previewLegendToggle = document.getElementById("mapPrintPreviewLegendToggle");
  const previewResetFrameBtn = document.getElementById("mapPrintPreviewResetFrame");
  const previewResetLegendBtn = document.getElementById("mapPrintPreviewResetLegend");
  const previewPrintBtn = document.getElementById("mapPrintPreviewPrintBtn");
  const previewCloseBtn = document.getElementById("mapPrintPreviewCloseBtn");

  if (!modal || !openBtn || !closeBtn) return;
  if (openBtn.dataset.printCenterBound === "1") return;
  openBtn.dataset.printCenterBound = "1";

  const previewState = {
    active: false,
    withLegend: true,
    frameMoved: false,
    frameResized: false,
    frameDragging: false,
    frameResizing: false,
    framePointerId: null,
    frameResizePointerId: null,
    frameDragOffsetX: 0,
    frameDragOffsetY: 0,
    frameStartRect: null,
    frameResizeStartX: 0,
    frameResizeStartY: 0,
    toolbarDragging: false,
    toolbarPointerId: null,
    toolbarDragOffsetX: 0,
    toolbarDragOffsetY: 0,
    popoutWindow: null,
    popoutOpen: false,
    dragging: false,
    dragPointerId: null,
    dragOffsetX: 0,
    dragOffsetY: 0,
    legendMoved: false
  };

  const PRINT_LEGEND_POSITION_STORAGE_KEY = "printLegendPosition";
  const PRINT_MAP_SPAN_STORAGE_KEY = "printMapSpan";
  const PRINT_OUTPUT_SCALE_STORAGE_KEY = "printOutputScale";
  const PRINT_RESOLUTION_SCALE_STORAGE_KEY = "printResolutionScale";

  const getLegendPosition = () => {
    const value = String(legendPositionSelect?.value || storageGet(PRINT_LEGEND_POSITION_STORAGE_KEY) || "outside-right");
    return PRINT_LEGEND_POSITION_OPTIONS.has(value) ? value : "outside-right";
  };

  const getMapSpan = () => {
    const value = String(mapSpanSelect?.value || storageGet(PRINT_MAP_SPAN_STORAGE_KEY) || "legend-right");
    return PRINT_MAP_SPAN_OPTIONS.has(value) ? value : "legend-right";
  };

  const getOutputScaleValue = () => {
    const value = String(outputScaleSelect?.value || storageGet(PRINT_OUTPUT_SCALE_STORAGE_KEY) || "1");
    return PRINT_OUTPUT_SCALE_OPTIONS.has(value) ? value : "1";
  };

  const getResolutionScaleValue = () => {
    const value = String(resolutionSelect?.value || storageGet(PRINT_RESOLUTION_SCALE_STORAGE_KEY) || "1");
    return PRINT_RESOLUTION_SCALE_OPTIONS.has(value) ? value : "1";
  };

  const getOutputScale = () => Number(getOutputScaleValue());
  const getResolutionScale = () => Number(getResolutionScaleValue());

  const syncLayoutControlsFromStorage = () => {
    if (legendPositionSelect) legendPositionSelect.value = getLegendPosition();
    if (mapSpanSelect) mapSpanSelect.value = getMapSpan();
    if (outputScaleSelect) outputScaleSelect.value = getOutputScaleValue();
    if (resolutionSelect) resolutionSelect.value = getResolutionScaleValue();
  };

  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

  const setPreviewToolbarPosition = ({ left, top, right = null }) => {
    if (!previewToolbar) return;
    const toolbarRect = previewToolbar.getBoundingClientRect();
    const width = Math.max(220, toolbarRect.width || 320);
    const height = Math.max(120, toolbarRect.height || 220);
    const safeLeft = clamp(Number(left) || 0, 6, Math.max(6, window.innerWidth - width - 6));
    const safeTop = clamp(Number(top) || 0, 6, Math.max(6, window.innerHeight - height - 6));
    previewToolbar.style.left = `${Math.round(safeLeft)}px`;
    previewToolbar.style.top = `${Math.round(safeTop)}px`;
    previewToolbar.style.right = right === null ? "auto" : String(right);
  };

  const resetPreviewToolbarPosition = () => {
    if (!previewToolbar) return;
    previewToolbar.style.left = "";
    previewToolbar.style.top = "";
    previewToolbar.style.right = "";
  };

  const setPreviewPopoutMode = enabled => {
    if (!previewOverlay) return;
    previewOverlay.classList.toggle("popout-mode", !!enabled);
  };

  const getPreviewFrameRectForSpan = spanValue => {
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const span = PRINT_MAP_SPAN_OPTIONS.has(String(spanValue)) ? String(spanValue) : "legend-right";
    if (span === "full") {
      return { left: 8, top: 8, width: Math.max(140, vw - 16), height: Math.max(140, vh - 16) };
    }
    if (span === "wide") {
      return { left: Math.round(vw * 0.01), top: Math.round(vh * 0.08), width: Math.max(140, Math.round(vw * 0.98)), height: Math.max(140, Math.round(vh * 0.84)) };
    }
    if (span === "standard") {
      return { left: Math.round(vw * 0.06), top: Math.round(vh * 0.1), width: Math.max(140, Math.round(vw * 0.88)), height: Math.max(140, Math.round(vh * 0.8)) };
    }
    if (span === "compact") {
      return { left: Math.round(vw * 0.14), top: Math.round(vh * 0.14), width: Math.max(140, Math.round(vw * 0.72)), height: Math.max(140, Math.round(vh * 0.72)) };
    }
    if (span === "legend-left") {
      const frameWidth = clamp(vw - 420, 240, vw - 24);
      return { left: vw - frameWidth - 14, top: 26, width: frameWidth, height: Math.max(160, vh - 52) };
    }
    const frameWidth = clamp(vw - 420, 240, vw - 24);
    return { left: 14, top: 26, width: frameWidth, height: Math.max(160, vh - 52) };
  };

  const getCurrentPreviewFrameRect = () => {
    if (!previewFrame) return null;
    const rect = previewFrame.getBoundingClientRect();
    return {
      left: rect.left,
      top: rect.top,
      width: rect.width,
      height: rect.height
    };
  };

  const clampPreviewFrameRect = rawRect => {
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const minW = 180;
    const minH = 140;
    const width = clamp(Number(rawRect?.width) || minW, minW, Math.max(minW, vw - 12));
    const height = clamp(Number(rawRect?.height) || minH, minH, Math.max(minH, vh - 12));
    const left = clamp(Number(rawRect?.left) || 0, 6, Math.max(6, vw - width - 6));
    const top = clamp(Number(rawRect?.top) || 0, 6, Math.max(6, vh - height - 6));
    return { left, top, width, height };
  };

  const setPreviewFrameRect = (rawRect, { markMoved = false, markResized = false } = {}) => {
    if (!previewFrame) return null;
    const rect = clampPreviewFrameRect(rawRect);
    previewFrame.style.left = `${Math.round(rect.left)}px`;
    previewFrame.style.top = `${Math.round(rect.top)}px`;
    previewFrame.style.width = `${Math.round(rect.width)}px`;
    previewFrame.style.height = `${Math.round(rect.height)}px`;
    if (markMoved) previewState.frameMoved = true;
    if (markResized) previewState.frameResized = true;
    return rect;
  };

  const applyPreviewFrameRectPreset = () => {
    if (!previewMapSpanSelect) return null;
    const rect = getPreviewFrameRectForSpan(previewMapSpanSelect.value);
    previewState.frameMoved = false;
    previewState.frameResized = false;
    return setPreviewFrameRect(rect, { markMoved: false, markResized: false });
  };

  const setPreviewLegendPosition = (left, top, { markMoved = false } = {}) => {
    if (!previewLegend || !previewFrame) return;
    const frameRect = previewFrame.getBoundingClientRect();
    const legendRect = previewLegend.getBoundingClientRect();
    const legendW = Math.max(90, legendRect.width || 260);
    const legendH = Math.max(70, legendRect.height || 220);
    const minLeft = -frameRect.left + 8;
    const maxLeft = window.innerWidth - frameRect.left - legendW - 8;
    const minTop = -frameRect.top + 8;
    const maxTop = window.innerHeight - frameRect.top - legendH - 8;
    const safeLeft = clamp(Number(left) || 0, minLeft, maxLeft);
    const safeTop = clamp(Number(top) || 0, minTop, maxTop);
    previewLegend.style.left = `${Math.round(safeLeft)}px`;
    previewLegend.style.top = `${Math.round(safeTop)}px`;
    if (markMoved) previewState.legendMoved = true;
  };

  const applyPreviewLegendPreset = (preset, { markMoved = false } = {}) => {
    if (!previewLegend || !previewFrame) return;
    const frameRect = previewFrame.getBoundingClientRect();
    const legendRect = previewLegend.getBoundingClientRect();
    const legendW = Math.max(90, legendRect.width || 260);
    const legendH = Math.max(70, legendRect.height || 220);
    const margin = 12;
    const outsideGap = 16;
    let left = frameRect.width - legendW - margin;
    let top = margin;
    const value = PRINT_LEGEND_POSITION_OPTIONS.has(String(preset)) ? String(preset) : "outside-right";
    if (value === "top-left") {
      left = margin;
      top = margin;
    } else if (value === "top-right") {
      left = frameRect.width - legendW - margin;
      top = margin;
    } else if (value === "bottom-left") {
      left = margin;
      top = frameRect.height - legendH - margin;
    } else if (value === "bottom-right") {
      left = frameRect.width - legendW - margin;
      top = frameRect.height - legendH - margin;
    } else if (value === "outside-left") {
      left = -legendW - outsideGap;
      top = margin;
    } else {
      left = frameRect.width + outsideGap;
      top = margin;
    }
    setPreviewLegendPosition(left, top, { markMoved });
  };

  const setStatus = (message, type = "") => {
    if (!statusNode) return;
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (type === "error") statusNode.classList.add("error");
    if (type === "success") statusNode.classList.add("success");
  };

  const openModal = () => {
    syncLayoutControlsFromStorage();
    setStatus("Ready.");
    openBtn.classList.add("active");
    modal.style.display = "flex";
  };

  const closeModal = () => {
    openBtn.classList.remove("active");
    modal.style.display = "none";
  };

  const openWindowAndPrint = (title, html, emptyMessage) => {
    if (!html) {
      setStatus(emptyMessage || "Nothing available to print right now.", "error");
      return;
    }
    const win = openPrintDocumentWindow({
      title,
      bodyHtml: html,
      autoPrint: true
    });
    if (!win) {
      setStatus("Unable to open print window. Check browser pop-up settings.", "error");
      return;
    }
    setStatus("Print window opened.", "success");
  };

  const getPreviewControlState = () => ({
    mapSpan: String(previewMapSpanSelect?.value || getMapSpan()),
    legendPosition: String(previewLegendPositionSelect?.value || getLegendPosition()),
    outputScale: String(previewOutputScaleSelect?.value || getOutputScaleValue()),
    resolutionScale: String(previewResolutionSelect?.value || getResolutionScaleValue()),
    legendVisible: !!previewLegendToggle?.checked,
    popoutOpen: !!previewState.popoutOpen
  });

  const setPreviewControlState = ({ mapSpan, legendPosition, outputScale, resolutionScale, legendVisible } = {}) => {
    if (previewMapSpanSelect && mapSpan && PRINT_MAP_SPAN_OPTIONS.has(String(mapSpan))) {
      previewMapSpanSelect.value = String(mapSpan);
      if (mapSpanSelect) mapSpanSelect.value = String(mapSpan);
      storageSet(PRINT_MAP_SPAN_STORAGE_KEY, String(mapSpan));
      applyPreviewFrameRectPreset();
      if (!previewState.legendMoved) {
        applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
      }
    }

    if (previewLegendPositionSelect && legendPosition && PRINT_LEGEND_POSITION_OPTIONS.has(String(legendPosition))) {
      previewLegendPositionSelect.value = String(legendPosition);
      if (legendPositionSelect) legendPositionSelect.value = String(legendPosition);
      storageSet(PRINT_LEGEND_POSITION_STORAGE_KEY, String(legendPosition));
      previewState.legendMoved = false;
      applyPreviewLegendPreset(String(legendPosition), { markMoved: false });
    }

    if (previewOutputScaleSelect && outputScale && PRINT_OUTPUT_SCALE_OPTIONS.has(String(outputScale))) {
      previewOutputScaleSelect.value = String(outputScale);
      if (outputScaleSelect) outputScaleSelect.value = String(outputScale);
      storageSet(PRINT_OUTPUT_SCALE_STORAGE_KEY, String(outputScale));
    }

    if (previewResolutionSelect && resolutionScale && PRINT_RESOLUTION_SCALE_OPTIONS.has(String(resolutionScale))) {
      previewResolutionSelect.value = String(resolutionScale);
      if (resolutionSelect) resolutionSelect.value = String(resolutionScale);
      storageSet(PRINT_RESOLUTION_SCALE_STORAGE_KEY, String(resolutionScale));
    }

    if (previewLegendToggle && typeof legendVisible === "boolean") {
      previewLegendToggle.checked = !!legendVisible;
      previewLegend?.classList.toggle("hidden", !previewLegendToggle.checked);
    }
  };

  const notifyPreviewPopoutState = () => {
    const popWin = previewState.popoutWindow;
    if (!popWin || popWin.closed) return;
    try {
      popWin.postMessage(
        {
          type: "tds-pak-map-print-preview-state",
          payload: getPreviewControlState()
        },
        "*"
      );
    } catch (_) {}
  };

  const openPreviewPopoutWindow = () => {
    const existing = previewState.popoutWindow;
    if (existing && !existing.closed) {
      existing.focus();
      setPreviewPopoutMode(true);
      previewState.popoutOpen = true;
      notifyPreviewPopoutState();
      return;
    }

    const pop = window.open("", "tdsPakMapPrintPreviewControls", "width=380,height=650,resizable=yes,scrollbars=yes");
    if (!pop) {
      setStatus("Could not open popout controls window (popup blocked).", "error");
      return;
    }
    previewState.popoutWindow = pop;
    previewState.popoutOpen = true;
    setPreviewPopoutMode(true);

    const stateSeed = JSON.stringify(getPreviewControlState()).replace(/</g, "\\u003c");
    pop.document.write(`
      <html>
        <head>
          <title>Map Print Preview Controls</title>
          <meta charset="UTF-8" />
          <style>
            :root { --bg:#1a1510; --panel:#2a2017; --line:#6f5a40; --text:#f4e3cf; --muted:#dcc3a5; --btn:#4d3a29; --btnLine:#7c6549; --accent:#73bfff; }
            * { box-sizing:border-box; }
            body { margin:0; padding:10px; background:var(--bg); color:var(--text); font-family:"Segoe UI", Roboto, Arial, sans-serif; }
            .shell { border:1px solid var(--line); border-radius:12px; background:linear-gradient(180deg,#302419 0%, #241c14 100%); padding:10px; }
            h3 { margin:0; font-size:16px; }
            p { margin:6px 0 10px; font-size:12px; color:var(--muted); line-height:1.4; }
            .guide { border:1px solid var(--line); border-radius:10px; padding:8px 9px; background:rgba(39,30,21,.75); margin-bottom:10px; }
            .guide-row { display:flex; justify-content:space-between; gap:8px; font-size:11px; margin-bottom:6px; color:#ecd5b7; }
            .guide-row:last-child { margin-bottom:0; }
            label { display:block; margin:8px 0 4px; font-size:11px; font-weight:700; letter-spacing:.02em; text-transform:uppercase; color:#f0d8bb; }
            select { width:100%; min-height:34px; border:1px solid #7b654a; border-radius:8px; background:#2a2117; color:#fff2e2; padding:0 10px; }
            .check { display:inline-flex; align-items:center; gap:8px; margin-top:10px; font-size:12px; font-weight:600; color:#ecd6ba; }
            .actions { margin-top:10px; display:grid; grid-template-columns:1fr 1fr; gap:6px; }
            button { min-height:32px; border-radius:8px; border:1px solid var(--btnLine); background:linear-gradient(165deg,#5f4933 0%, #4a3928 100%); color:#fff3e3; font-size:12px; font-weight:700; cursor:pointer; }
            button:hover { background:linear-gradient(165deg,#6d5338 0%, #59442f 100%); }
            .danger { border-color:#95625f; background:linear-gradient(165deg,#6b3d3a 0%, #55312f 100%); }
            .status { margin-top:8px; border:1px solid rgba(111,90,64,.7); border-radius:8px; background:rgba(34,26,18,.75); color:#ebd7c0; font-size:11px; padding:6px 8px; min-height:16px; }
          </style>
        </head>
        <body>
          <div class="shell">
            <h3>Map Print Preview</h3>
            <p>This control window stays out of the map so you can edit print layout without blocking view.</p>
            <div class="guide">
              <div class="guide-row"><span>Move frame</span><strong>Drag blue header</strong></div>
              <div class="guide-row"><span>Resize frame</span><strong>Drag corner handle</strong></div>
              <div class="guide-row"><span>Move legend</span><strong>Drag legend card</strong></div>
            </div>
            <label for="mapSpan">Map Span</label>
            <select id="mapSpan">
              <option value="legend-right">Legend Right Span</option>
              <option value="legend-left">Legend Left Span</option>
              <option value="full">Full Page</option>
              <option value="wide">Wide</option>
              <option value="standard">Standard</option>
              <option value="compact">Compact</option>
            </select>
            <label for="legendPos">Legend Preset</label>
            <select id="legendPos">
              <option value="outside-right">Outside Right</option>
              <option value="outside-left">Outside Left</option>
              <option value="top-right">Top Right</option>
              <option value="top-left">Top Left</option>
              <option value="bottom-right">Bottom Right</option>
              <option value="bottom-left">Bottom Left</option>
            </select>
            <label for="outputScale">Print Size</label>
            <select id="outputScale">
              <option value="1">100% (Match Preview)</option>
              <option value="0.95">95%</option>
              <option value="0.9">90%</option>
              <option value="0.85">85%</option>
            </select>
            <label for="resolutionScale">Map Resolution</label>
            <select id="resolutionScale">
              <option value="1">Standard</option>
              <option value="1.5">High (1.5x)</option>
              <option value="2">Very High (2x)</option>
            </select>
            <label class="check"><input id="legendVisible" type="checkbox" checked />Include legend in print</label>
            <div class="actions">
              <button id="resetFrame" type="button">Reset Frame</button>
              <button id="resetLegend" type="button">Reset Legend</button>
              <button id="printNow" type="button">Print</button>
              <button id="dock" type="button">Dock Controls</button>
              <button id="closePreview" class="danger" type="button">Close Preview</button>
              <button id="closeWindow" class="danger" type="button">Close Window</button>
            </div>
            <div id="status" class="status">Connected.</div>
          </div>
          <script>
            const stateSeed = ${stateSeed};
            const mapSpan = document.getElementById("mapSpan");
            const legendPos = document.getElementById("legendPos");
            const outputScale = document.getElementById("outputScale");
            const resolutionScale = document.getElementById("resolutionScale");
            const legendVisible = document.getElementById("legendVisible");
            const status = document.getElementById("status");

            function api() {
              return window.opener && window.opener.__mapPrintPreviewApi ? window.opener.__mapPrintPreviewApi : null;
            }

            function setStatus(msg) {
              status.textContent = msg;
            }

            function applyState(s) {
              if (!s) return;
              if (s.mapSpan) mapSpan.value = s.mapSpan;
              if (s.legendPosition) legendPos.value = s.legendPosition;
              if (s.outputScale) outputScale.value = s.outputScale;
              if (s.resolutionScale) resolutionScale.value = s.resolutionScale;
              legendVisible.checked = !!s.legendVisible;
            }

            applyState(stateSeed);

            window.addEventListener("message", ev => {
              if (!ev || !ev.data || ev.data.type !== "tds-pak-map-print-preview-state") return;
              applyState(ev.data.payload || {});
            });

            mapSpan.addEventListener("change", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.setState({ mapSpan: mapSpan.value });
              setStatus("Map span updated.");
            });

            legendPos.addEventListener("change", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.setState({ legendPosition: legendPos.value });
              setStatus("Legend preset updated.");
            });

            outputScale.addEventListener("change", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.setState({ outputScale: outputScale.value });
              setStatus("Print size updated.");
            });

            resolutionScale.addEventListener("change", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.setState({ resolutionScale: resolutionScale.value });
              setStatus("Map resolution updated.");
            });

            legendVisible.addEventListener("change", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.setState({ legendVisible: !!legendVisible.checked });
              setStatus("Legend visibility updated.");
            });

            document.getElementById("resetFrame").addEventListener("click", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.resetFrame();
              setStatus("Frame reset.");
            });

            document.getElementById("resetLegend").addEventListener("click", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.resetLegend();
              setStatus("Legend reset.");
            });

            document.getElementById("printNow").addEventListener("click", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.printNow();
            });

            document.getElementById("dock").addEventListener("click", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.setPopout(false);
              setStatus("Docked controls back in main window.");
            });

            document.getElementById("closePreview").addEventListener("click", () => {
              const a = api();
              if (!a) return setStatus("Connection lost.");
              a.closePreview();
              window.close();
            });

            document.getElementById("closeWindow").addEventListener("click", () => {
              window.close();
            });

            window.addEventListener("beforeunload", () => {
              const a = api();
              if (a) a.onPopoutClosed();
            });
          <\/script>
        </body>
      </html>
    `);
    pop.document.close();
    pop.focus();
    notifyPreviewPopoutState();
  };

  const closeMapPrintPreview = () => {
    if (!previewOverlay) return;
    if (previewToolbarHandle && previewState.toolbarPointerId !== null && typeof previewToolbarHandle.releasePointerCapture === "function") {
      try { previewToolbarHandle.releasePointerCapture(previewState.toolbarPointerId); } catch (_) {}
    }
    if (previewFrameMoveHandle && previewState.framePointerId !== null && typeof previewFrameMoveHandle.releasePointerCapture === "function") {
      try { previewFrameMoveHandle.releasePointerCapture(previewState.framePointerId); } catch (_) {}
    }
    if (previewFrameResizeHandle && previewState.frameResizePointerId !== null && typeof previewFrameResizeHandle.releasePointerCapture === "function") {
      try { previewFrameResizeHandle.releasePointerCapture(previewState.frameResizePointerId); } catch (_) {}
    }
    if (previewLegend && previewState.dragPointerId !== null && typeof previewLegend.releasePointerCapture === "function") {
      try { previewLegend.releasePointerCapture(previewState.dragPointerId); } catch (_) {}
    }
    previewOverlay.classList.remove("show");
    previewOverlay.setAttribute("aria-hidden", "true");
    setPreviewPopoutMode(false);
    previewState.active = false;
    previewState.frameDragging = false;
    previewState.frameResizing = false;
    previewState.framePointerId = null;
    previewState.frameResizePointerId = null;
    previewState.toolbarDragging = false;
    previewState.toolbarPointerId = null;
    previewState.dragging = false;
    previewState.dragPointerId = null;
    previewFrame?.classList.remove("moving", "resizing");
    previewLegend?.classList.remove("dragging");
    const popWin = previewState.popoutWindow;
    if (popWin && !popWin.closed) {
      try { popWin.close(); } catch (_) {}
    }
    previewState.popoutWindow = null;
    previewState.popoutOpen = false;
  };

  const openMapPrintPreview = ({ withLegend }) => {
    if (!previewOverlay || !previewFrame || !previewLegend || !previewMapSpanSelect || !previewLegendPositionSelect || !previewOutputScaleSelect || !previewResolutionSelect || !previewLegendToggle) {
      runMapPrintFlow({
        withLegend: !!withLegend,
        mapSpan: getMapSpan(),
        legendPosition: getLegendPosition(),
        outputScale: getOutputScale(),
        resolutionScale: getResolutionScale()
      });
      return;
    }

    previewState.active = true;
    previewState.withLegend = !!withLegend;
    previewState.frameMoved = false;
    previewState.frameResized = false;
    previewState.legendMoved = false;

    const startSpan = getMapSpan();
    const startLegendPosition = getLegendPosition();
    const startOutputScale = getOutputScaleValue();
    const startResolutionScale = getResolutionScaleValue();

    previewMapSpanSelect.value = startSpan;
    previewLegendPositionSelect.value = startLegendPosition;
    previewOutputScaleSelect.value = startOutputScale;
    previewResolutionSelect.value = startResolutionScale;
    previewLegendToggle.checked = !!withLegend;
    previewLegend.innerHTML = buildMapLegendHtmlForPrint();

    previewOverlay.classList.add("show");
    previewOverlay.setAttribute("aria-hidden", "false");
    setPreviewPopoutMode(!!previewState.popoutOpen);
    previewFrame.classList.remove("moving", "resizing");
    applyPreviewFrameRectPreset();
    previewLegend.classList.toggle("hidden", !previewLegendToggle.checked);

    requestAnimationFrame(() => {
      applyPreviewLegendPreset(previewLegendPositionSelect.value, { markMoved: false });
      try { map.invalidateSize({ pan: false }); } catch (_) {}
      notifyPreviewPopoutState();
    });
  };

  openBtn.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });

  legendPositionSelect?.addEventListener("change", () => {
    storageSet(PRINT_LEGEND_POSITION_STORAGE_KEY, getLegendPosition());
  });

  mapSpanSelect?.addEventListener("change", () => {
    storageSet(PRINT_MAP_SPAN_STORAGE_KEY, getMapSpan());
  });

  outputScaleSelect?.addEventListener("change", () => {
    storageSet(PRINT_OUTPUT_SCALE_STORAGE_KEY, getOutputScaleValue());
    if (previewOutputScaleSelect) previewOutputScaleSelect.value = getOutputScaleValue();
    notifyPreviewPopoutState();
  });

  resolutionSelect?.addEventListener("change", () => {
    storageSet(PRINT_RESOLUTION_SCALE_STORAGE_KEY, getResolutionScaleValue());
    if (previewResolutionSelect) previewResolutionSelect.value = getResolutionScaleValue();
    notifyPreviewPopoutState();
  });

  syncLayoutControlsFromStorage();

  previewMapSpanSelect?.addEventListener("change", () => {
    if (mapSpanSelect) mapSpanSelect.value = previewMapSpanSelect.value;
    storageSet(PRINT_MAP_SPAN_STORAGE_KEY, getMapSpan());
    applyPreviewFrameRectPreset();
    if (!previewState.legendMoved) {
      applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
    }
    notifyPreviewPopoutState();
  });

  previewLegendPositionSelect?.addEventListener("change", () => {
    if (legendPositionSelect) legendPositionSelect.value = previewLegendPositionSelect.value;
    storageSet(PRINT_LEGEND_POSITION_STORAGE_KEY, getLegendPosition());
    previewState.legendMoved = false;
    applyPreviewLegendPreset(previewLegendPositionSelect.value, { markMoved: false });
    notifyPreviewPopoutState();
  });

  previewOutputScaleSelect?.addEventListener("change", () => {
    if (!PRINT_OUTPUT_SCALE_OPTIONS.has(String(previewOutputScaleSelect.value))) {
      previewOutputScaleSelect.value = getOutputScaleValue();
    }
    if (outputScaleSelect) outputScaleSelect.value = previewOutputScaleSelect.value;
    storageSet(PRINT_OUTPUT_SCALE_STORAGE_KEY, previewOutputScaleSelect.value);
    notifyPreviewPopoutState();
  });

  previewResolutionSelect?.addEventListener("change", () => {
    if (!PRINT_RESOLUTION_SCALE_OPTIONS.has(String(previewResolutionSelect.value))) {
      previewResolutionSelect.value = getResolutionScaleValue();
    }
    if (resolutionSelect) resolutionSelect.value = previewResolutionSelect.value;
    storageSet(PRINT_RESOLUTION_SCALE_STORAGE_KEY, previewResolutionSelect.value);
    notifyPreviewPopoutState();
  });

  previewLegendToggle?.addEventListener("change", () => {
    if (!previewLegend) return;
    previewLegend.classList.toggle("hidden", !previewLegendToggle.checked);
    notifyPreviewPopoutState();
  });

  const resetPreviewFrame = () => {
    applyPreviewFrameRectPreset();
    if (!previewState.legendMoved) {
      applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
    }
    notifyPreviewPopoutState();
  };

  const resetPreviewLegend = () => {
    previewState.legendMoved = false;
    applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
    notifyPreviewPopoutState();
  };

  previewResetFrameBtn?.addEventListener("click", resetPreviewFrame);
  previewResetLegendBtn?.addEventListener("click", resetPreviewLegend);

  previewCloseBtn?.addEventListener("click", () => {
    closeMapPrintPreview();
  });

  const printFromPreview = () => {
    if (!previewMapSpanSelect || !previewLegendPositionSelect || !previewOutputScaleSelect || !previewResolutionSelect || !previewLegendToggle || !previewLegend || !previewFrame) {
      closeMapPrintPreview();
      runMapPrintFlow({
        withLegend: !!previewState.withLegend,
        mapSpan: getMapSpan(),
        legendPosition: getLegendPosition(),
        outputScale: getOutputScale(),
        resolutionScale: getResolutionScale()
      });
      return;
    }

    const mapSpan = previewMapSpanSelect.value;
    const legendPosition = previewLegendPositionSelect.value;
    const outputScaleValue = PRINT_OUTPUT_SCALE_OPTIONS.has(String(previewOutputScaleSelect.value))
      ? String(previewOutputScaleSelect.value)
      : getOutputScaleValue();
    const resolutionScaleValue = PRINT_RESOLUTION_SCALE_OPTIONS.has(String(previewResolutionSelect.value))
      ? String(previewResolutionSelect.value)
      : getResolutionScaleValue();
    const outputScale = Number(outputScaleValue);
    const resolutionScale = Number(resolutionScaleValue);
    const withLegend = !!previewLegendToggle.checked;
    if (mapSpanSelect) mapSpanSelect.value = mapSpan;
    if (legendPositionSelect) legendPositionSelect.value = legendPosition;
    if (outputScaleSelect) outputScaleSelect.value = outputScaleValue;
    if (resolutionSelect) resolutionSelect.value = resolutionScaleValue;
    storageSet(PRINT_MAP_SPAN_STORAGE_KEY, getMapSpan());
    storageSet(PRINT_LEGEND_POSITION_STORAGE_KEY, getLegendPosition());
    storageSet(PRINT_OUTPUT_SCALE_STORAGE_KEY, outputScaleValue);
    storageSet(PRINT_RESOLUTION_SCALE_STORAGE_KEY, resolutionScaleValue);

    let customLegendLeftPx = NaN;
    let customLegendTopPx = NaN;
    let customLegendLeftRatio = NaN;
    let customLegendTopRatio = NaN;
    if (withLegend) {
      const legendRect = previewLegend.getBoundingClientRect();
      customLegendLeftPx = legendRect.left;
      customLegendTopPx = legendRect.top;
      if (window.innerWidth > 0 && window.innerHeight > 0) {
        customLegendLeftRatio = legendRect.left / window.innerWidth;
        customLegendTopRatio = legendRect.top / window.innerHeight;
      }
    }

    let customMapLeftPx = NaN;
    let customMapTopPx = NaN;
    let customMapWidthPx = NaN;
    let customMapHeightPx = NaN;
    let customMapLeftRatio = NaN;
    let customMapTopRatio = NaN;
    let customMapWidthRatio = NaN;
    let customMapHeightRatio = NaN;
    const frameRect = previewFrame.getBoundingClientRect();
    if (frameRect && frameRect.width > 120 && frameRect.height > 120) {
      customMapLeftPx = frameRect.left;
      customMapTopPx = frameRect.top;
      customMapWidthPx = frameRect.width;
      customMapHeightPx = frameRect.height;
      if (window.innerWidth > 0 && window.innerHeight > 0) {
        customMapLeftRatio = frameRect.left / window.innerWidth;
        customMapTopRatio = frameRect.top / window.innerHeight;
        customMapWidthRatio = frameRect.width / window.innerWidth;
        customMapHeightRatio = frameRect.height / window.innerHeight;
      }
    }

    closeMapPrintPreview();
    runMapPrintFlow({
      withLegend,
      mapSpan,
      legendPosition,
      customLegendLeftPx,
      customLegendTopPx,
      customLegendLeftRatio,
      customLegendTopRatio,
      customMapLeftPx,
      customMapTopPx,
      customMapWidthPx,
      customMapHeightPx,
      customMapLeftRatio,
      customMapTopRatio,
      customMapWidthRatio,
      customMapHeightRatio,
      outputScale,
      resolutionScale
    });
  };

  previewPrintBtn?.addEventListener("click", printFromPreview);

  const setPreviewPopout = enabled => {
    const next = !!enabled;
    if (!previewState.active) return;
    if (next) {
      openPreviewPopoutWindow();
      return;
    }
    const popWin = previewState.popoutWindow;
    if (popWin && !popWin.closed) {
      try { popWin.close(); } catch (_) {}
    }
    previewState.popoutWindow = null;
    previewState.popoutOpen = false;
    setPreviewPopoutMode(false);
    notifyPreviewPopoutState();
  };

  window.__mapPrintPreviewApi = {
    getState: () => getPreviewControlState(),
    setState: next => {
      setPreviewControlState(next || {});
      notifyPreviewPopoutState();
    },
    resetFrame: () => {
      resetPreviewFrame();
    },
    resetLegend: () => {
      resetPreviewLegend();
    },
    printNow: () => {
      printFromPreview();
    },
    closePreview: () => {
      closeMapPrintPreview();
    },
    setPopout: enabled => {
      setPreviewPopout(!!enabled);
    },
    onPopoutClosed: () => {
      previewState.popoutWindow = null;
      previewState.popoutOpen = false;
      if (previewState.active) {
        setPreviewPopoutMode(false);
      }
    }
  };

  previewPopoutBtn?.addEventListener("click", () => {
    setPreviewPopout(true);
  });

  previewOverlay?.addEventListener("click", event => {
    if (event.target === previewOverlay) closeMapPrintPreview();
  });

  previewToolbarHandle?.addEventListener("pointerdown", event => {
    if (!previewState.active || !previewToolbar || !previewToolbarHandle) return;
    if (event.button !== 0) return;
    if (event.target && event.target.closest("button")) return;
    const toolbarRect = previewToolbar.getBoundingClientRect();
    previewState.toolbarDragging = true;
    previewState.toolbarPointerId = event.pointerId;
    previewState.toolbarDragOffsetX = event.clientX - toolbarRect.left;
    previewState.toolbarDragOffsetY = event.clientY - toolbarRect.top;
    if (typeof previewToolbarHandle.setPointerCapture === "function") {
      try { previewToolbarHandle.setPointerCapture(event.pointerId); } catch (_) {}
    }
    event.preventDefault();
    event.stopPropagation();
  });

  const onPreviewToolbarPointerMove = event => {
    if (!previewState.toolbarDragging || !previewToolbar) return;
    if (previewState.toolbarPointerId !== null && event.pointerId !== previewState.toolbarPointerId) return;
    const left = event.clientX - previewState.toolbarDragOffsetX;
    const top = event.clientY - previewState.toolbarDragOffsetY;
    setPreviewToolbarPosition({ left, top, right: "auto" });
    event.preventDefault();
  };

  const endPreviewToolbarPointer = event => {
    if (!previewState.toolbarDragging) return;
    if (previewState.toolbarPointerId !== null && event && event.pointerId !== undefined && event.pointerId !== previewState.toolbarPointerId) {
      return;
    }
    previewState.toolbarDragging = false;
    previewState.toolbarPointerId = null;
    if (event && previewToolbarHandle && typeof previewToolbarHandle.releasePointerCapture === "function") {
      try { previewToolbarHandle.releasePointerCapture(event.pointerId); } catch (_) {}
    }
  };

  previewToolbarHandle?.addEventListener("pointermove", onPreviewToolbarPointerMove);
  previewToolbarHandle?.addEventListener("pointerup", endPreviewToolbarPointer);
  previewToolbarHandle?.addEventListener("pointercancel", endPreviewToolbarPointer);

  previewFrameMoveHandle?.addEventListener("pointerdown", event => {
    if (!previewState.active || !previewFrame) return;
    if (event.button !== 0) return;
    const rect = getCurrentPreviewFrameRect();
    if (!rect) return;
    previewState.frameResizing = false;
    previewState.frameResizePointerId = null;
    previewState.frameStartRect = null;
    previewFrame.classList.remove("resizing");
    previewState.frameDragging = true;
    previewState.framePointerId = event.pointerId;
    previewState.frameDragOffsetX = event.clientX - rect.left;
    previewState.frameDragOffsetY = event.clientY - rect.top;
    previewFrame.classList.add("moving");
    if (typeof previewFrameMoveHandle.setPointerCapture === "function") {
      try { previewFrameMoveHandle.setPointerCapture(event.pointerId); } catch (_) {}
    }
    event.preventDefault();
    event.stopPropagation();
  });

  previewFrameResizeHandle?.addEventListener("pointerdown", event => {
    if (!previewState.active || !previewFrame) return;
    if (event.button !== 0) return;
    const rect = getCurrentPreviewFrameRect();
    if (!rect) return;
    previewState.frameDragging = false;
    previewState.framePointerId = null;
    previewFrame.classList.remove("moving");
    previewState.frameResizing = true;
    previewState.frameResizePointerId = event.pointerId;
    previewState.frameStartRect = rect;
    previewState.frameResizeStartX = event.clientX;
    previewState.frameResizeStartY = event.clientY;
    previewFrame.classList.add("resizing");
    if (typeof previewFrameResizeHandle.setPointerCapture === "function") {
      try { previewFrameResizeHandle.setPointerCapture(event.pointerId); } catch (_) {}
    }
    event.preventDefault();
    event.stopPropagation();
  });

  const onPreviewFramePointerMove = event => {
    if (!previewState.active || !previewFrame) return;

    if (previewState.frameDragging) {
      if (previewState.framePointerId !== null && event.pointerId !== previewState.framePointerId) return;
      const current = getCurrentPreviewFrameRect();
      if (!current) return;
      const left = event.clientX - previewState.frameDragOffsetX;
      const top = event.clientY - previewState.frameDragOffsetY;
      setPreviewFrameRect({ left, top, width: current.width, height: current.height }, { markMoved: true, markResized: false });
      if (!previewState.legendMoved) {
        applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
      }
      event.preventDefault();
      return;
    }

    if (previewState.frameResizing) {
      if (previewState.frameResizePointerId !== null && event.pointerId !== previewState.frameResizePointerId) return;
      const start = previewState.frameStartRect;
      if (!start) return;
      const width = start.width + (event.clientX - previewState.frameResizeStartX);
      const height = start.height + (event.clientY - previewState.frameResizeStartY);
      setPreviewFrameRect({ left: start.left, top: start.top, width, height }, { markMoved: false, markResized: true });
      if (!previewState.legendMoved) {
        applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
      }
      event.preventDefault();
    }
  };

  const endPreviewFramePointer = event => {
    if (previewState.frameDragging) {
      if (previewState.framePointerId !== null && event && event.pointerId !== undefined && event.pointerId !== previewState.framePointerId) {
        return;
      }
      previewState.frameDragging = false;
      previewState.framePointerId = null;
      previewFrame?.classList.remove("moving");
      if (event && previewFrameMoveHandle && typeof previewFrameMoveHandle.releasePointerCapture === "function") {
        try { previewFrameMoveHandle.releasePointerCapture(event.pointerId); } catch (_) {}
      }
    }

    if (previewState.frameResizing) {
      if (previewState.frameResizePointerId !== null && event && event.pointerId !== undefined && event.pointerId !== previewState.frameResizePointerId) {
        return;
      }
      previewState.frameResizing = false;
      previewState.frameResizePointerId = null;
      previewState.frameStartRect = null;
      previewFrame?.classList.remove("resizing");
      if (event && previewFrameResizeHandle && typeof previewFrameResizeHandle.releasePointerCapture === "function") {
        try { previewFrameResizeHandle.releasePointerCapture(event.pointerId); } catch (_) {}
      }
    }
  };

  previewFrameMoveHandle?.addEventListener("pointermove", onPreviewFramePointerMove);
  previewFrameResizeHandle?.addEventListener("pointermove", onPreviewFramePointerMove);
  previewFrameMoveHandle?.addEventListener("pointerup", endPreviewFramePointer);
  previewFrameResizeHandle?.addEventListener("pointerup", endPreviewFramePointer);
  previewFrameMoveHandle?.addEventListener("pointercancel", endPreviewFramePointer);
  previewFrameResizeHandle?.addEventListener("pointercancel", endPreviewFramePointer);

  previewLegend?.addEventListener("pointerdown", event => {
    if (!previewState.active || !previewFrame || !previewLegend || previewLegend.classList.contains("hidden")) return;
    if (event.button !== 0) return;
    const frameRect = previewFrame.getBoundingClientRect();
    const legendRect = previewLegend.getBoundingClientRect();
    previewState.dragging = true;
    previewState.dragPointerId = event.pointerId;
    previewState.dragOffsetX = event.clientX - legendRect.left;
    previewState.dragOffsetY = event.clientY - legendRect.top;
    previewLegend.classList.add("dragging");
    if (typeof previewLegend.setPointerCapture === "function") {
      try { previewLegend.setPointerCapture(event.pointerId); } catch (_) {}
    }
    const currentLeft = legendRect.left - frameRect.left;
    const currentTop = legendRect.top - frameRect.top;
    setPreviewLegendPosition(currentLeft, currentTop, { markMoved: false });
    event.preventDefault();
    event.stopPropagation();
  });

  const onPreviewLegendPointerMove = event => {
    if (!previewState.dragging || !previewFrame || !previewLegend) return;
    if (previewState.dragPointerId !== null && event.pointerId !== previewState.dragPointerId) return;
    const frameRect = previewFrame.getBoundingClientRect();
    const left = event.clientX - frameRect.left - previewState.dragOffsetX;
    const top = event.clientY - frameRect.top - previewState.dragOffsetY;
    setPreviewLegendPosition(left, top, { markMoved: true });
    event.preventDefault();
  };

  const endPreviewLegendDrag = event => {
    if (!previewState.dragging) return;
    if (previewState.dragPointerId !== null && event && event.pointerId !== undefined && event.pointerId !== previewState.dragPointerId) {
      return;
    }
    previewState.dragging = false;
    previewState.dragPointerId = null;
    previewLegend?.classList.remove("dragging");
    if (event && previewLegend && typeof previewLegend.releasePointerCapture === "function") {
      try { previewLegend.releasePointerCapture(event.pointerId); } catch (_) {}
    }
  };

  previewLegend?.addEventListener("pointermove", onPreviewLegendPointerMove);
  previewLegend?.addEventListener("pointerup", endPreviewLegendDrag);
  previewLegend?.addEventListener("pointercancel", endPreviewLegendDrag);

  window.addEventListener("resize", () => {
    if (!previewState.active) return;
    if (previewToolbar && !previewState.popoutOpen) {
      const toolbarRect = previewToolbar.getBoundingClientRect();
      setPreviewToolbarPosition({
        left: toolbarRect.left,
        top: toolbarRect.top,
        right: "auto"
      });
    }
    if (previewState.frameMoved || previewState.frameResized) {
      const current = getCurrentPreviewFrameRect();
      if (current) {
        setPreviewFrameRect(current, {
          markMoved: previewState.frameMoved,
          markResized: previewState.frameResized
        });
      } else {
        applyPreviewFrameRectPreset();
      }
    } else {
      applyPreviewFrameRectPreset();
    }
    if (!previewState.legendMoved) {
      applyPreviewLegendPreset(previewLegendPositionSelect?.value || getLegendPosition(), { markMoved: false });
    } else if (previewLegend && previewFrame) {
      const frameRect = previewFrame.getBoundingClientRect();
      const legendRect = previewLegend.getBoundingClientRect();
      setPreviewLegendPosition(legendRect.left - frameRect.left, legendRect.top - frameRect.top, { markMoved: true });
    }
  });

  window.addEventListener("keydown", event => {
    if (!previewState.active) return;
    if (event.key === "Escape") {
      closeMapPrintPreview();
    }
  });

  printMapOnlyBtn?.addEventListener("click", () => {
    closeModal();
    openMapPrintPreview({ withLegend: false });
  });

  printMapLegendBtn?.addEventListener("click", () => {
    closeModal();
    openMapPrintPreview({ withLegend: true });
  });

  printSummaryTableBtn?.addEventListener("click", () => {
    const html = buildSummaryTablePrintDocumentHtml();
    openWindowAndPrint("Route Summary Table", html, "No route summary table is loaded.");
  });

  printSummaryVizBtn?.addEventListener("click", () => {
    const html = buildSummaryVisualizationPrintDocumentHtml();
    openWindowAndPrint("Route Summary Visualization", html, "No summary data is loaded for visualization printing.");
  });

  printAttributeTableBtn?.addEventListener("click", () => {
    const html = buildAttributeTablePrintDocumentHtml();
    openWindowAndPrint("Attribute Table", html, "No attribute rows are available to print.");
  });

  printSelectionReportBtn?.addEventListener("click", () => {
    const html = buildSelectionReportPrintDocumentHtml();
    openWindowAndPrint("Selection Report", html, "No selected records or street segments to print.");
  });

  printOperationalReportBtn?.addEventListener("click", () => {
    const html = buildOperationalReportPrintDocumentHtml();
    openWindowAndPrint("Operational Report", html, "Unable to build operational report.");
  });
}

function initLayerManagerControls() {
  const openBtn = document.getElementById("layerManagerBtn");
  const modal = document.getElementById("layerManagerModal");
  const closeBtn = document.getElementById("layerManagerCloseBtn");
  const showAllBtn = document.getElementById("layerManagerShowAllBtn");
  const hideAllBtn = document.getElementById("layerManagerHideAllBtn");
  const listNode = document.getElementById("layerManagerList");
  const statusNode = document.getElementById("layerManagerStatus");

  if (!openBtn || !modal || !closeBtn || !listNode) return;
  if (openBtn.dataset.layerManagerBound === "1") return;
  openBtn.dataset.layerManagerBound = "1";

  ensureLayerManagerOrder();
  ensureLayerManagerSelectableState();
  let draggingEntryId = "";

  const setStatus = message => {
    if (!statusNode) return;
    statusNode.textContent = String(message || "");
  };

  const isModalOpen = () => modal.style.display === "flex";

  const buildRouteDayLabel = key => {
    const [route = "", dayToken = ""] = String(key || "").split("|");
    const dayNumber = Number(dayToken);
    const dayLabel = Number.isFinite(dayNumber) ? (dayName(dayNumber) || String(dayToken)) : String(dayToken || "No Day");
    return `Route ${route || "Unassigned"} - ${dayLabel}`;
  };

  const buildEntryMeta = entryId => {
    if (entryId === LAYER_MANAGER_STREET_KEY) {
      const sourceToggle = document.getElementById("useLocalStreetSource");
      const sourceOn = !!sourceToggle?.checked;
      const visibleOnMap = map.hasLayer(streetAttributeLayerGroup);
      const count = streetAttributeById.size;
      if (!sourceOn) return `Source off - ${count.toLocaleString()} segments loaded`;
      return `${visibleOnMap ? "Visible" : "Hidden"} - ${count.toLocaleString()} segments loaded`;
    }
    const group = routeDayGroups[entryId];
    const count = Array.isArray(group?.layers) ? group.layers.length : 0;
    const visible = isRouteDayLayerVisibleOnMap(entryId);
    return `${visible ? "Visible" : "Hidden"} - ${count.toLocaleString()} stops`;
  };

  const isEntryVisible = entryId => {
    if (entryId === LAYER_MANAGER_STREET_KEY) return isStreetNetworkLayerVisibleEnabled();
    if (Object.prototype.hasOwnProperty.call(layerVisibilityState, entryId)) {
      return !!layerVisibilityState[entryId];
    }
    return isRouteDayLayerVisibleOnMap(entryId);
  };

  const applyEntryVisibility = (entryId, visible) => {
    if (entryId === LAYER_MANAGER_STREET_KEY) {
      setStreetNetworkLayerVisibilityFromManager(visible);
      return;
    }
    setRouteDayLayerVisibilityFromManager(entryId, visible);
  };

  const createSwatchNode = entryId => {
    const swatch = document.createElement("span");
    swatch.className = "layer-manager-swatch";

    if (entryId === LAYER_MANAGER_STREET_KEY) {
      swatch.classList.add("line");
      swatch.style.background = "#4ea2f5";
      return swatch;
    }

    const symbol = symbolMap[entryId] || getSymbol(entryId);
    swatch.style.background = symbol.color;
    if (symbol.shape === "circle") {
      swatch.style.borderRadius = "50%";
    } else if (symbol.shape === "square") {
      swatch.style.borderRadius = "2px";
    } else if (symbol.shape === "triangle") {
      swatch.style.width = "0";
      swatch.style.height = "0";
      swatch.style.borderLeft = "6px solid transparent";
      swatch.style.borderRight = "6px solid transparent";
      swatch.style.borderBottom = `12px solid ${symbol.color}`;
      swatch.style.borderRadius = "0";
      swatch.style.background = "transparent";
      swatch.style.borderTop = "0";
    } else if (symbol.shape === "diamond") {
      swatch.style.borderRadius = "2px";
      swatch.style.transform = "rotate(45deg)";
    }
    return swatch;
  };

  const updateOrderAndRender = message => {
    applyLayerManagerOrder();
    renderLayerManagerList();
    if (message) setStatus(message);
  };

  function renderLayerManagerList() {
    const orderTop = ensureLayerManagerOrder();
    ensureLayerManagerSelectableState();
    listNode.innerHTML = "";

    if (!orderTop.length) {
      const empty = document.createElement("div");
      empty.className = "layer-manager-empty";
      empty.textContent = "No layers available yet. Load a route file to manage map layers.";
      listNode.appendChild(empty);
      return;
    }

    const routeOrderTop = orderTop.filter(isRouteDayLayerManagerEntry);
    const hasStreetEntry = orderTop.includes(LAYER_MANAGER_STREET_KEY);
    const streetOnTop = orderTop[0] === LAYER_MANAGER_STREET_KEY;

    const createRowsForEntries = (entryIds, sectionBody) => {
      entryIds.forEach(entryId => {
        const row = document.createElement("div");
        row.className = "layer-manager-row";
        row.dataset.entryId = entryId;
        row.draggable = true;

        const dragHandle = document.createElement("span");
        dragHandle.className = "layer-manager-drag";
        dragHandle.title = "Drag to reorder";
        dragHandle.textContent = "::";

        const swatch = createSwatchNode(entryId);

        const labelWrap = document.createElement("div");
        labelWrap.className = "layer-manager-label-wrap";
        const label = document.createElement("span");
        label.className = "layer-manager-label";
        label.textContent = entryId === LAYER_MANAGER_STREET_KEY ? "Street Network" : buildRouteDayLabel(entryId);
        const meta = document.createElement("span");
        meta.className = "layer-manager-meta";
        meta.textContent = buildEntryMeta(entryId);
        labelWrap.append(label, meta);

        const visLabel = document.createElement("label");
        visLabel.className = "layer-manager-vis-check";
        const visInput = document.createElement("input");
        visInput.type = "checkbox";
        visInput.checked = isEntryVisible(entryId);
        const visText = document.createElement("span");
        visText.textContent = "Show";
        visLabel.append(visInput, visText);

        visInput.addEventListener("change", () => {
          applyEntryVisibility(entryId, visInput.checked);
          refreshLayerManagerUiIfOpen();
          setStatus(`Updated visibility: ${label.textContent}`);
        });

        const selectLabel = document.createElement("label");
        selectLabel.className = "layer-manager-select-check";
        const selectInput = document.createElement("input");
        selectInput.type = "checkbox";
        selectInput.checked = isLayerManagerEntrySelectable(entryId);
        const selectText = document.createElement("span");
        selectText.textContent = "Selectable";
        selectLabel.append(selectInput, selectText);

        selectInput.addEventListener("change", () => {
          const changed = setLayerManagerEntrySelectable(entryId, selectInput.checked);
          updateSelectionCount();
          if (!changed) return;
          const stateText = selectInput.checked ? "enabled" : "disabled";
          setStatus(`Selection ${stateText}: ${label.textContent}`);
        });

        const orderActions = document.createElement("div");
        orderActions.className = "layer-manager-order-actions";
        const upBtn = document.createElement("button");
        upBtn.type = "button";
        upBtn.className = "layer-manager-order-btn";
        upBtn.title = "Move layer up";
        upBtn.textContent = "^";
        const downBtn = document.createElement("button");
        downBtn.type = "button";
        downBtn.className = "layer-manager-order-btn";
        downBtn.title = "Move layer down";
        downBtn.textContent = "v";

        if (entryId === LAYER_MANAGER_STREET_KEY) {
          upBtn.disabled = streetOnTop || !routeOrderTop.length;
          downBtn.disabled = !streetOnTop || !routeOrderTop.length;
        } else {
          const routeIndex = routeOrderTop.indexOf(entryId);
          upBtn.disabled = routeIndex <= 0;
          downBtn.disabled = routeIndex < 0 || routeIndex >= routeOrderTop.length - 1;
        }

        upBtn.addEventListener("click", () => {
          if (!moveLayerManagerEntryByOffset(entryId, -1)) return;
          updateOrderAndRender(`Moved up: ${label.textContent}`);
        });

        downBtn.addEventListener("click", () => {
          if (!moveLayerManagerEntryByOffset(entryId, 1)) return;
          updateOrderAndRender(`Moved down: ${label.textContent}`);
        });

        orderActions.append(upBtn, downBtn);

        row.addEventListener("dragstart", event => {
          draggingEntryId = entryId;
          row.classList.add("dragging");
          try { event.dataTransfer?.setData("text/plain", entryId); } catch (_) {}
          if (event.dataTransfer) event.dataTransfer.effectAllowed = "move";
        });

        row.addEventListener("dragend", () => {
          draggingEntryId = "";
          row.classList.remove("dragging");
          listNode.querySelectorAll(".layer-manager-row.drop-target").forEach(node => node.classList.remove("drop-target"));
        });

        row.addEventListener("dragover", event => {
          event.preventDefault();
          if (!draggingEntryId || draggingEntryId === entryId) return;
          row.classList.add("drop-target");
        });

        row.addEventListener("dragleave", () => {
          row.classList.remove("drop-target");
        });

        row.addEventListener("drop", event => {
          event.preventDefault();
          row.classList.remove("drop-target");
          if (!draggingEntryId || draggingEntryId === entryId) return;
          if (!moveLayerManagerEntryBefore(draggingEntryId, entryId)) return;
          updateOrderAndRender("Layer draw order updated.");
        });

        row.append(dragHandle, swatch, labelWrap, visLabel, selectLabel, orderActions);
        sectionBody.appendChild(row);
      });
    };

    const appendSection = (title, subtitle, entryIds) => {
      if (!entryIds.length) return;
      const section = document.createElement("section");
      section.className = "layer-manager-section";

      const head = document.createElement("div");
      head.className = "layer-manager-section-head";
      const titleNode = document.createElement("span");
      titleNode.className = "layer-manager-section-title";
      titleNode.textContent = title;
      const subtitleNode = document.createElement("span");
      subtitleNode.className = "layer-manager-section-meta";
      subtitleNode.textContent = subtitle;
      head.append(titleNode, subtitleNode);

      const body = document.createElement("div");
      body.className = "layer-manager-section-body";
      createRowsForEntries(entryIds, body);

      section.append(head, body);
      listNode.appendChild(section);
    };

    if (streetOnTop && hasStreetEntry) {
      appendSection("Map Layers", "Move this above or below the Route + Day group.", [LAYER_MANAGER_STREET_KEY]);
    }

    appendSection(
      "Route + Day Layers",
      "These layers are grouped together. Drag rows here to set top-to-bottom order.",
      routeOrderTop
    );

    if (!streetOnTop && hasStreetEntry) {
      appendSection("Map Layers", "Move this above or below the Route + Day group.", [LAYER_MANAGER_STREET_KEY]);
    }
  }

  const openModal = () => {
    ensureLayerManagerOrder();
    ensureLayerManagerSelectableState();
    renderLayerManagerList();
    setStatus("Drag rows to order layers. Use Show and Selectable toggles per layer.");
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  showAllBtn?.addEventListener("click", () => {
    ensureLayerManagerOrder().forEach(entryId => {
      applyEntryVisibility(entryId, true);
    });
    renderLayerManagerList();
    setStatus("All layers turned on.");
  });

  hideAllBtn?.addEventListener("click", () => {
    ensureLayerManagerOrder().forEach(entryId => {
      applyEntryVisibility(entryId, false);
    });
    renderLayerManagerList();
    setStatus("All layers turned off.");
  });

  openBtn.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });

  window.addEventListener("keydown", event => {
    if (!isModalOpen()) return;
    if (event.key === "Escape") closeModal();
  });

  window.__refreshLayerManagerList = () => {
    if (!isModalOpen()) return;
    renderLayerManagerList();
  };
}

function getFilteredAttributeRows() {
  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  const headers = window._attributeHeaders || [];
  const needle = attributeState.filterText.trim().toLowerCase();

  let data = rows
    .map(row => ({ row, rowId: getAttributeRowId(row) }))
    .filter(item => Number.isFinite(item.rowId));

  if (needle) {
    data = data.filter(({ row }) =>
      headers.some(h => String(row?.[h] ?? "").toLowerCase().includes(needle))
    );
  }

  if (attributeState.selectedOnly) {
    data = data.filter(({ rowId }) => attributeState.selectedRowIds.has(rowId));
  }

  if (attributeState.sortKey) {
    const key = attributeState.sortKey;
    const dir = attributeState.sortDir;
    data.sort((a, b) => {
      const av = a.row?.[key];
      const bv = b.row?.[key];
      const an = Number(av);
      const bn = Number(bv);
      if (Number.isFinite(an) && Number.isFinite(bn)) {
        return (an - bn) * dir;
      }
      return String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true }) * dir;
    });
  }

  attributeState.lastVisibleRows = data;
  return data;
}

function renderAttributeTable() {
  if (attributeTableMode === "streets") {
    renderStreetAttributeTable();
    return;
  }
  const table = document.getElementById("attributeTableGrid");
  const empty = document.getElementById("attributeTableEmpty");
  const pageInfo = document.getElementById("attributePageInfo");
  const prevBtn = document.getElementById("attributePrevPageBtn");
  const nextBtn = document.getElementById("attributeNextPageBtn");
  if (!table || !empty) return;

  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  const headers = window._attributeHeaders || [];

  if (!rows.length || !headers.length) {
    table.innerHTML = "";
    empty.style.display = "block";
    if (pageInfo) pageInfo.textContent = "Page 1/1";
    if (prevBtn) prevBtn.disabled = true;
    if (nextBtn) nextBtn.disabled = true;
    refreshAttributeStatus();
    return;
  }

  const visibleRows = getFilteredAttributeRows();
  const totalRows = visibleRows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / attributeState.pageSize));
  if (attributeState.page > totalPages) attributeState.page = totalPages;
  if (attributeState.page < 1) attributeState.page = 1;
  const pageStart = (attributeState.page - 1) * attributeState.pageSize;
  const pageRows = visibleRows.slice(pageStart, pageStart + attributeState.pageSize);

  if (pageInfo) pageInfo.textContent = `Page ${attributeState.page}/${totalPages}`;
  if (prevBtn) prevBtn.disabled = attributeState.page <= 1;
  if (nextBtn) nextBtn.disabled = attributeState.page >= totalPages;

  empty.style.display = visibleRows.length ? "none" : "block";
  if (!visibleRows.length) {
    table.innerHTML = "";
    if (pageInfo) pageInfo.textContent = "Page 1/1";
    if (prevBtn) prevBtn.disabled = true;
    if (nextBtn) nextBtn.disabled = true;
    refreshAttributeStatus();
    return;
  }

  const sortIndicator = (key) => {
    if (attributeState.sortKey !== key) return "";
    return attributeState.sortDir > 0 ? " ▲" : " ▼";
  };

  let html = "<thead><tr><th>Sel</th><th>#</th>";
  headers.forEach(h => {
    html += `<th><button type="button" data-sort="${h.replace(/"/g, "&quot;")}">${h}${sortIndicator(h)}</button></th>`;
  });
  html += "</tr></thead><tbody>";

  pageRows.forEach(({ row, rowId }, idx) => {
    const checked = attributeState.selectedRowIds.has(rowId) ? " checked" : "";
    html += `<tr data-row-id="${rowId}" class="${checked ? "selected" : ""}">`;
    html += `<td><input type="checkbox" data-row-select="${rowId}"${checked}></td>`;
    html += `<td>${pageStart + idx + 1}</td>`;
    headers.forEach(h => {
      const value = row?.[h];
      html += `<td>${String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody>";
  table.innerHTML = html;

  table.querySelectorAll("button[data-sort]").forEach(btn => {
    btn.addEventListener("click", e => {
      const key = e.currentTarget.getAttribute("data-sort");
      if (!key) return;
      if (attributeState.sortKey === key) {
        attributeState.sortDir = attributeState.sortDir * -1;
      } else {
        attributeState.sortKey = key;
        attributeState.sortDir = 1;
      }
      attributeState.page = 1;
      renderAttributeTable();
    });
  });

  table.querySelectorAll("input[data-row-select]").forEach(input => {
    input.addEventListener("change", e => {
      const rowId = Number(e.currentTarget.getAttribute("data-row-select"));
      setAttributeRowSelected(rowId, e.currentTarget.checked, true);
    });
  });

  table.querySelectorAll("tbody tr").forEach(tr => {
    tr.addEventListener("click", e => {
      if (e.target.closest("input")) return;
      const rowId = Number(tr.getAttribute("data-row-id"));
      if (!Number.isFinite(rowId)) return;
      window.focusAttributeRowOnMap(rowId);
    });
  });

  refreshAttributeStatus();
}

function exportAttributeVisibleRowsToCsv() {
  if (attributeTableMode === "streets") {
    const rows = getFilteredStreetAttributeRows();
    if (!rows.length) {
      alert("No street attributes to export.");
      return;
    }
    const headers = ["id", "name", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"];
    const esc = (v) => {
      const raw = String(v ?? "");
      if (/[\",\\n]/.test(raw)) return `\"${raw.replace(/\"/g, '\"\"')}\"`;
      return raw;
    };
    const lines = [headers.join(",")];
    rows.forEach(r => lines.push(headers.map(h => esc(r[h] ?? "")).join(",")));
    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "street-attributes.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    return;
  }
  const headers = window._attributeHeaders || [];
  const rows = getFilteredAttributeRows();
  if (!rows.length || !headers.length) {
    alert("No attribute rows to export.");
    return;
  }

  const esc = (v) => {
    const raw = String(v ?? "");
    if (/[",\n]/.test(raw)) return `"${raw.replace(/"/g, "\"\"")}"`;
    return raw;
  };

  const lines = [];
  lines.push(headers.map(esc).join(","));
  rows.forEach(({ row }) => {
    lines.push(headers.map(h => esc(row?.[h] ?? "")).join(","));
  });

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "attribute-table.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function getAttributePanelBounds() {
  const mapEl = document.getElementById("map");
  const header = document.querySelector("header");
  const selectionBox = document.getElementById("selectionBox");
  const summary = document.getElementById("bottomSummary");

  const mapRect = mapEl?.getBoundingClientRect?.() || {
    left: 8,
    top: 84,
    right: window.innerWidth - 8,
    bottom: window.innerHeight - 8
  };

  const headerBottom = Math.ceil(header?.getBoundingClientRect?.().bottom || 84);
  let left = Math.max(8, Math.ceil(mapRect.left) + 8);
  let top = Math.max(headerBottom + 8, Math.ceil(mapRect.top) + 8);
  let right = Math.min(window.innerWidth - 8, Math.ceil(mapRect.right) - 8);
  let bottom = Math.min(window.innerHeight - 8, Math.ceil(mapRect.bottom) - 8);

  if (window.innerWidth > 900 && selectionBox && !selectionBox.classList.contains("collapsed")) {
    const selRect = selectionBox.getBoundingClientRect();
    right = Math.min(right, Math.floor(selRect.left) - 8);
  }

  if (summary && !summary.classList.contains("collapsed")) {
    const sumRect = summary.getBoundingClientRect();
    bottom = Math.min(bottom, Math.floor(sumRect.top) - 8);
  }

  return { left, top, right, bottom };
}

function clampAttributePanelRect(left, top, width, height) {
  const bounds = getAttributePanelBounds();
  const minWidth = window.innerWidth > 900 ? 420 : 320;
  const minHeight = 220;
  const maxWidth = Math.max(minWidth, bounds.right - bounds.left);
  const maxHeight = Math.max(minHeight, bounds.bottom - bounds.top);

  const safeWidth = Math.max(minWidth, Math.min(maxWidth, Number(width) || minWidth));
  const safeHeight = Math.max(minHeight, Math.min(maxHeight, Number(height) || minHeight));
  const safeLeft = Math.max(bounds.left, Math.min(bounds.right - safeWidth, Number(left) || bounds.left));
  const safeTop = Math.max(bounds.top, Math.min(bounds.bottom - safeHeight, Number(top) || bounds.top));

  return { left: safeLeft, top: safeTop, width: safeWidth, height: safeHeight, bounds };
}

function applyAttributePanelRect(panel, rect) {
  if (!panel || !rect) return;
  panel.style.left = `${Math.round(rect.left)}px`;
  panel.style.top = `${Math.round(rect.top)}px`;
  panel.style.width = `${Math.round(rect.width)}px`;
  panel.style.height = `${Math.round(rect.height)}px`;
}

function saveAttributePanelRect(panel) {
  if (!panel) return;
  storageSet("attributePanelLeft", Math.round(panel.offsetLeft));
  storageSet("attributePanelTop", Math.round(panel.offsetTop));
  storageSet("attributePanelWidth", Math.round(panel.offsetWidth));
  storageSet("attributePanelHeight", Math.round(panel.offsetHeight));
}

function refreshMapAfterOverlayChange() {
  if (!map || typeof map.invalidateSize !== "function") return;
  requestAnimationFrame(() => {
    try { map.invalidateSize({ pan: false }); } catch (_) {}
  });
  setTimeout(() => {
    try { map.invalidateSize({ pan: false }); } catch (_) {}
  }, 120);
}

function snapAttributePanelToCorner(panel) {
  if (!panel) return;
  const rect = clampAttributePanelRect(panel.offsetLeft, panel.offsetTop, panel.offsetWidth, panel.offsetHeight);
  const threshold = 72;
  const corners = [
    { left: rect.bounds.left, top: rect.bounds.top }, // top-left
    { left: rect.bounds.right - rect.width, top: rect.bounds.top }, // top-right
    { left: rect.bounds.left, top: rect.bounds.bottom - rect.height }, // bottom-left
    { left: rect.bounds.right - rect.width, top: rect.bounds.bottom - rect.height } // bottom-right
  ];

  let closest = null;
  let minDist = Infinity;
  corners.forEach(corner => {
    const dx = rect.left - corner.left;
    const dy = rect.top - corner.top;
    const dist = Math.hypot(dx, dy);
    if (dist < minDist) {
      minDist = dist;
      closest = corner;
    }
  });

  if (closest && minDist <= threshold) {
    applyAttributePanelRect(panel, { ...rect, left: closest.left, top: closest.top });
    return true;
  }
  return false;
}

function syncAttributePanelLayout() {
  const panel = document.getElementById("attributeTablePanel");
  const btnDesktop = document.getElementById("attributeTableBtn");
  const btnMobile = document.getElementById("attributeTableBtnMobile");
  const collapseBtn = document.getElementById("attributeDockCloseBtn");
  if (!panel) return;

  const isClosed = panel.classList.contains("closed");
  const isCollapsed = panel.classList.contains("collapsed");
  panel.setAttribute("aria-hidden", isClosed ? "true" : "false");
  panel.inert = !!isClosed;

  const label = isClosed ? "Attribute Table" : "Hide Table";
  if (btnDesktop) btnDesktop.querySelector("span:last-child").textContent = label;
  if (btnMobile) btnMobile.textContent = isClosed ? "Table" : "Close Table";
  if (collapseBtn) collapseBtn.textContent = isCollapsed ? "Show" : "Hide";
}

function openAttributePanel() {
  const panel = document.getElementById("attributeTablePanel");
  if (!panel) return;
  panel.classList.remove("closed", "collapsed");
  const bounds = getAttributePanelBounds();
  const savedW = Number(storageGet("attributePanelWidth"));
  const savedH = Number(storageGet("attributePanelHeight"));
  const savedL = Number(storageGet("attributePanelLeft"));
  const savedT = Number(storageGet("attributePanelTop"));

  const defaultWidth = Math.min(980, Math.max(420, bounds.right - bounds.left - 120));
  const defaultHeight = Math.min(520, Math.max(240, bounds.bottom - bounds.top - 60));
  const defaultLeft = bounds.left + 16;
  const defaultTop = bounds.top + 8;

  const rect = clampAttributePanelRect(
    Number.isFinite(savedL) ? savedL : defaultLeft,
    Number.isFinite(savedT) ? savedT : defaultTop,
    Number.isFinite(savedW) ? savedW : defaultWidth,
    Number.isFinite(savedH) ? savedH : defaultHeight
  );
  applyAttributePanelRect(panel, rect);
  syncAttributePanelLayout();
  renderAttributeTable();
  refreshMapAfterOverlayChange();
}

function closeAttributePanel() {
  const panel = document.getElementById("attributeTablePanel");
  if (!panel) return;
  const active = document.activeElement;
  if (active && panel.contains(active)) {
    const fallback =
      attributeTableMode === "streets"
        ? (document.getElementById("streetAttributesBtn") || document.getElementById("streetAttributesBtnMobile"))
        : (document.getElementById("attributeTableBtn") || document.getElementById("attributeTableBtnMobile"));
    if (fallback && typeof fallback.focus === "function") fallback.focus();
    if (active && typeof active.blur === "function") active.blur();
  }
  panel.classList.add("closed");
  syncAttributePanelLayout();
  refreshMapAfterOverlayChange();
}

function zoomToSelectedAttributeRows() {
  if (!attributeState.selectedRowIds.size) {
    alert("No selected rows.");
    return;
  }
  const bounds = L.latLngBounds();
  attributeState.selectedRowIds.forEach(rowId => {
    const marker = getAttributeMarker(rowId);
    const latlng = getLayerLatLng(marker);
    if (latlng) bounds.extend(latlng);
  });
  if (!bounds.isValid()) {
    alert("Selected rows are not currently on the map.");
    return;
  }
  map.fitBounds(bounds.pad(0.18));
}

function openStreetAttributeTablePopout() {
  const headers = ["id", "name", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"];
  const rows = (Array.isArray(streetAttributesRows) && streetAttributesRows.length
    ? streetAttributesRows
    : [...streetAttributeById.values()].map(entry => entry?.row).filter(Boolean))
    .map(row => ({ rowId: Number(row?.id), values: row || {} }))
    .filter(item => Number.isFinite(item.rowId));

  if (!rows.length) {
    alert("No street attributes to open.");
    return;
  }

  const win = window.open("", "_blank", "width=1180,height=700,resizable=yes,scrollbars=yes");
  if (!win) {
    alert("Unable to open street attribute table window.");
    return;
  }

  const seed = JSON.stringify({
    headers,
    rows,
    selectedIds: [...streetAttributeSelectedIds]
  }).replace(/</g, "\\u003c");

  win.document.write(`
    <html>
      <head>
        <title>Street Attributes</title>
        <style>
          body { margin:0; font-family: Roboto, Arial, sans-serif; background:#0f1822; color:#e8f2fd; }
          .bar { position:sticky; top:0; display:flex; align-items:center; gap:8px; flex-wrap:wrap; padding:10px; background:#1a2938; border-bottom:1px solid #31495f; z-index:4; }
          .bar button { border:1px solid #5a7ca1; background:#2a4258; color:#edf6ff; border-radius:8px; padding:6px 10px; cursor:pointer; }
          .bar input[type=text] { min-width:220px; height:32px; border-radius:8px; border:1px solid #5a7ca1; background:#111a24; color:#eaf3fc; padding:0 10px; }
          .bar label { display:inline-flex; align-items:center; gap:6px; font-size:12px; border:1px solid #48647e; border-radius:8px; padding:5px 8px; background:#132130; }
          .status { margin-left:auto; font-size:12px; color:#c8dcf1; }
          .wrap { padding:10px; height:calc(100vh - 128px); overflow:auto; }
          table { border-collapse:collapse; min-width:100%; width:max-content; background:#16212b; }
          th, td { border:1px solid #31475b; padding:6px 8px; white-space:nowrap; font-size:12px; }
          th { position:sticky; top:0; background:#26394a; text-align:left; }
          th button { border:0; background:transparent; color:inherit; font:inherit; font-weight:700; cursor:pointer; padding:0; }
          tbody tr { cursor:pointer; }
          tbody tr:nth-child(even) { background:#192633; }
          tbody tr.selected { background: rgba(84,176,255,0.22); }
        </style>
      </head>
      <body>
        <div class="bar">
          <button onclick="window.close()">Close</button>
          <input id="searchInput" type="text" placeholder="Filter streets by any field..." />
          <label><input id="selectedOnly" type="checkbox" /> Selected only</label>
          <button id="selectVisibleBtn">Select Visible</button>
          <button id="clearSelectedBtn">Clear Selection</button>
          <button id="zoomSelectedBtn">Zoom to Selected</button>
          <button id="exportCsvBtn">Export CSV</button>
          <button id="prevBtn">Prev</button>
          <button id="nextBtn">Next</button>
          <span id="pageInfo">Page 1/1</span>
          <span id="status" class="status">0 selected</span>
        </div>
        <div class="wrap">
          <table id="table"></table>
        </div>
        <script>
          const seed = ${seed};
          const headers = seed.headers || [];
          const rows = seed.rows || [];
          let selectedSet = new Set(seed.selectedIds || []);
          const state = { sortKey: null, sortDir: 1, filterText: "", selectedOnly: false, page: 1, pageSize: 300 };

          const table = document.getElementById("table");
          const pageInfo = document.getElementById("pageInfo");
          const status = document.getElementById("status");
          const prevBtn = document.getElementById("prevBtn");
          const nextBtn = document.getElementById("nextBtn");
          const searchInput = document.getElementById("searchInput");
          const selectedOnly = document.getElementById("selectedOnly");
          const selectVisibleBtn = document.getElementById("selectVisibleBtn");
          const clearSelectedBtn = document.getElementById("clearSelectedBtn");
          const zoomSelectedBtn = document.getElementById("zoomSelectedBtn");
          const exportCsvBtn = document.getElementById("exportCsvBtn");

          function setsEqual(a, b) {
            if (a.size !== b.size) return false;
            for (const value of a) {
              if (!b.has(value)) return false;
            }
            return true;
          }

          function pullSelectionFromOpener() {
            if (!window.opener || typeof window.opener.getStreetSelectedSegmentIds !== "function") return false;
            const ids = window.opener.getStreetSelectedSegmentIds() || [];
            const next = new Set(ids.map(v => Number(v)).filter(v => Number.isFinite(v)));
            if (setsEqual(selectedSet, next)) return false;
            selectedSet = next;
            return true;
          }

          function pushSelectionToOpener() {
            if (!window.opener || typeof window.opener.setStreetSelectedSegmentIds !== "function") return;
            window.opener.setStreetSelectedSegmentIds([...selectedSet]);
          }

          function esc(v) {
            return String(v ?? "")
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;");
          }

          function getFilteredRows() {
            const needle = (state.filterText || "").trim().toLowerCase();
            let data = rows.slice();
            if (needle) {
              data = data.filter(item => headers.some(h => String(item.values?.[h] ?? "").toLowerCase().includes(needle)));
            }
            if (state.selectedOnly) {
              data = data.filter(item => selectedSet.has(item.rowId));
            }
            if (state.sortKey) {
              const key = state.sortKey;
              const dir = state.sortDir;
              data.sort((a, b) => {
                const av = a.values?.[key];
                const bv = b.values?.[key];
                const an = Number(av);
                const bn = Number(bv);
                if (Number.isFinite(an) && Number.isFinite(bn)) return (an - bn) * dir;
                return String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true }) * dir;
              });
            }
            return data;
          }

          function exportCsv(items) {
            if (!items.length) return;
            const csvEsc = (v) => {
              const raw = String(v ?? "");
              return /[",\\n]/.test(raw) ? ('"' + raw.replace(/"/g, '""') + '"') : raw;
            };
            const lines = [headers.map(csvEsc).join(",")];
            items.forEach(item => lines.push(headers.map(h => csvEsc(item.values?.[h] ?? "")).join(",")));
            const blob = new Blob([lines.join("\\n")], { type: "text/csv;charset=utf-8;" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "street-attributes.csv";
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
          }

          function render() {
            const filtered = getFilteredRows();
            const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
            state.page = Math.max(1, Math.min(state.page, totalPages));
            const start = (state.page - 1) * state.pageSize;
            const pageRows = filtered.slice(start, start + state.pageSize);

            const sortMarker = (key) => state.sortKey === key ? (state.sortDir > 0 ? " ^" : " v") : "";
            let html = "<thead><tr><th>Sel</th><th>#</th>";
            headers.forEach(h => {
              html += '<th><button type="button" data-sort="' + esc(h) + '">' + esc(h) + sortMarker(h) + "</button></th>";
            });
            html += "</tr></thead><tbody>";
            pageRows.forEach((item, i) => {
              const checked = selectedSet.has(item.rowId) ? " checked" : "";
              html += '<tr data-row-id="' + item.rowId + '" class="' + (checked ? "selected" : "") + '">';
              html += '<td><input type="checkbox" data-row-select="' + item.rowId + '"' + checked + "></td>";
              html += "<td>" + (start + i + 1) + "</td>";
              headers.forEach(h => {
                html += "<td>" + esc(item.values?.[h] ?? "") + "</td>";
              });
              html += "</tr>";
            });
            html += "</tbody>";
            table.innerHTML = html;

            pageInfo.textContent = "Page " + state.page + "/" + totalPages;
            status.textContent = selectedSet.size + " selected • " + filtered.length + " visible";
            prevBtn.disabled = state.page <= 1;
            nextBtn.disabled = state.page >= totalPages;

            table.querySelectorAll("button[data-sort]").forEach(btn => {
              btn.addEventListener("click", () => {
                const key = btn.getAttribute("data-sort");
                if (!key) return;
                if (state.sortKey === key) state.sortDir *= -1;
                else { state.sortKey = key; state.sortDir = 1; }
                state.page = 1;
                render();
              });
            });

            table.querySelectorAll("input[data-row-select]").forEach(input => {
              input.addEventListener("change", () => {
                const rowId = Number(input.getAttribute("data-row-select"));
                if (!Number.isFinite(rowId)) return;
                if (input.checked) selectedSet.add(rowId);
                else selectedSet.delete(rowId);
                pushSelectionToOpener();
                render();
              });
            });

            table.querySelectorAll("tbody tr[data-row-id]").forEach(tr => {
              tr.addEventListener("click", (e) => {
                if (e.target.closest("input")) return;
                const rowId = Number(tr.getAttribute("data-row-id"));
                if (window.opener && Number.isFinite(rowId) && typeof window.opener.focusStreetSegmentOnMap === "function") {
                  window.opener.focusStreetSegmentOnMap(rowId);
                }
              });
            });
          }

          searchInput.addEventListener("input", () => { state.filterText = searchInput.value || ""; state.page = 1; render(); });
          selectedOnly.addEventListener("change", () => {
            pullSelectionFromOpener();
            state.selectedOnly = !!selectedOnly.checked;
            state.page = 1;
            render();
          });
          prevBtn.addEventListener("click", () => { state.page = Math.max(1, state.page - 1); render(); });
          nextBtn.addEventListener("click", () => { state.page += 1; render(); });
          selectVisibleBtn.addEventListener("click", () => {
            getFilteredRows().forEach(item => selectedSet.add(item.rowId));
            pushSelectionToOpener();
            render();
          });
          clearSelectedBtn.addEventListener("click", () => {
            selectedSet.clear();
            pushSelectionToOpener();
            render();
          });
          zoomSelectedBtn.addEventListener("click", () => {
            pullSelectionFromOpener();
            const selectedIds = [...selectedSet];
            if (!selectedIds.length) return;
            if (window.opener && typeof window.opener.zoomToStreetSegmentsOnMap === "function") {
              const ok = window.opener.zoomToStreetSegmentsOnMap(selectedIds);
              if (ok) return;
            }
            if (window.opener && typeof window.opener.focusStreetSegmentOnMap === "function") {
              window.opener.focusStreetSegmentOnMap(selectedIds[0]);
            }
          });
          exportCsvBtn.addEventListener("click", () => exportCsv(getFilteredRows()));

          window.addEventListener("focus", () => {
            if (pullSelectionFromOpener()) render();
          });
          setInterval(() => {
            const changed = pullSelectionFromOpener();
            if (changed) render();
          }, 500);

          pullSelectionFromOpener();
          render();
        <\/script>
      </body>
    </html>
  `);
  win.document.close();
}

function openAttributeTablePopout() {
  if (attributeTableMode === "streets") {
    openStreetAttributeTablePopout();
    return;
  }
  const headers = window._attributeHeaders || [];
  const rows = (window._currentRows || [])
    .map(row => ({ rowId: getAttributeRowId(row), values: row }))
    .filter(item => Number.isFinite(item.rowId));
  if (!headers.length || !rows.length) {
    alert("No attribute data to open.");
    return;
  }

  const win = window.open("", "_blank", "width=1180,height=700,resizable=yes,scrollbars=yes");
  if (!win) {
    alert("Unable to open attribute table window.");
    return;
  }

  const seed = JSON.stringify({
    headers,
    rows,
    selectedIds: [...attributeState.selectedRowIds]
  }).replace(/</g, "\\u003c");

  win.document.write(`
    <html>
      <head>
        <title>Attribute Table</title>
        <style>
          body { margin:0; font-family: Roboto, Arial, sans-serif; background:#0f1822; color:#e8f2fd; }
          .bar { position:sticky; top:0; display:flex; align-items:center; gap:8px; flex-wrap:wrap; padding:10px; background:#1a2938; border-bottom:1px solid #31495f; z-index:4; }
          .bar button { border:1px solid #5a7ca1; background:#2a4258; color:#edf6ff; border-radius:8px; padding:6px 10px; cursor:pointer; }
          .bar input[type=text] { min-width:220px; height:32px; border-radius:8px; border:1px solid #5a7ca1; background:#111a24; color:#eaf3fc; padding:0 10px; }
          .bar label { display:inline-flex; align-items:center; gap:6px; font-size:12px; border:1px solid #48647e; border-radius:8px; padding:5px 8px; background:#132130; }
          .status { margin-left:auto; font-size:12px; color:#c8dcf1; }
          .wrap { padding:10px; height:calc(100vh - 128px); overflow:auto; }
          table { border-collapse:collapse; min-width:100%; width:max-content; background:#16212b; }
          th, td { border:1px solid #31475b; padding:6px 8px; white-space:nowrap; font-size:12px; }
          th { position:sticky; top:0; background:#26394a; text-align:left; }
          th button { border:0; background:transparent; color:inherit; font:inherit; font-weight:700; cursor:pointer; padding:0; }
          tbody tr { cursor:pointer; }
          tbody tr:nth-child(even) { background:#192633; }
          tbody tr.selected { background: rgba(84,176,255,0.22); }
        </style>
      </head>
      <body>
        <div class="bar">
          <button onclick="window.close()">Close</button>
          <input id="searchInput" type="text" placeholder="Filter records by any field..." />
          <label><input id="selectedOnly" type="checkbox" /> Selected only</label>
          <button id="selectVisibleBtn">Select Visible</button>
          <button id="clearSelectedBtn">Clear Selection</button>
          <button id="zoomSelectedBtn">Zoom to Selected</button>
          <button id="exportCsvBtn">Export CSV</button>
          <button id="prevBtn">Prev</button>
          <button id="nextBtn">Next</button>
          <span id="pageInfo">Page 1/1</span>
          <span id="status" class="status">0 selected</span>
        </div>
        <div class="wrap">
          <table id="table"></table>
        </div>
        <script>
          const seed = ${seed};
          const headers = seed.headers || [];
          const rows = seed.rows || [];
          let selectedSet = new Set(seed.selectedIds || []);
          const state = { sortKey: null, sortDir: 1, filterText: "", selectedOnly: false, page: 1, pageSize: 300 };

          const table = document.getElementById("table");
          const pageInfo = document.getElementById("pageInfo");
          const status = document.getElementById("status");
          const prevBtn = document.getElementById("prevBtn");
          const nextBtn = document.getElementById("nextBtn");
          const searchInput = document.getElementById("searchInput");
          const selectedOnly = document.getElementById("selectedOnly");
          const selectVisibleBtn = document.getElementById("selectVisibleBtn");
          const clearSelectedBtn = document.getElementById("clearSelectedBtn");
          const zoomSelectedBtn = document.getElementById("zoomSelectedBtn");
          const exportCsvBtn = document.getElementById("exportCsvBtn");

          function setsEqual(a, b) {
            if (a.size !== b.size) return false;
            for (const value of a) {
              if (!b.has(value)) return false;
            }
            return true;
          }

          function pullSelectionFromOpener() {
            if (!window.opener || typeof window.opener.getAttributeSelectedRowIds !== "function") return false;
            const ids = window.opener.getAttributeSelectedRowIds() || [];
            const next = new Set(ids.map(v => Number(v)).filter(v => Number.isFinite(v)));
            if (setsEqual(selectedSet, next)) return false;
            selectedSet = next;
            return true;
          }

          function pushSelectionToOpener() {
            if (!window.opener || typeof window.opener.setAttributeSelectedRowIds !== "function") return;
            window.opener.setAttributeSelectedRowIds([...selectedSet]);
          }

          function esc(v) {
            return String(v ?? "")
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;");
          }

          function getFilteredRows() {
            const needle = (state.filterText || "").trim().toLowerCase();
            let data = rows.slice();
            if (needle) {
              data = data.filter(item => headers.some(h => String(item.values?.[h] ?? "").toLowerCase().includes(needle)));
            }
            if (state.selectedOnly) {
              data = data.filter(item => selectedSet.has(item.rowId));
            }
            if (state.sortKey) {
              const key = state.sortKey;
              const dir = state.sortDir;
              data.sort((a, b) => {
                const av = a.values?.[key];
                const bv = b.values?.[key];
                const an = Number(av);
                const bn = Number(bv);
                if (Number.isFinite(an) && Number.isFinite(bn)) return (an - bn) * dir;
                return String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true }) * dir;
              });
            }
            return data;
          }

          function exportCsv(items) {
            if (!items.length) return;
            const csvEsc = (v) => {
              const raw = String(v ?? "");
              return /[",\\n]/.test(raw) ? ('"' + raw.replace(/"/g, '""') + '"') : raw;
            };
            const lines = [headers.map(csvEsc).join(",")];
            items.forEach(item => lines.push(headers.map(h => csvEsc(item.values?.[h] ?? "")).join(",")));
            const blob = new Blob([lines.join("\\n")], { type: "text/csv;charset=utf-8;" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "attribute-table.csv";
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
          }

          function render() {
            const filtered = getFilteredRows();
            const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
            state.page = Math.max(1, Math.min(state.page, totalPages));
            const start = (state.page - 1) * state.pageSize;
            const pageRows = filtered.slice(start, start + state.pageSize);

            const sortMarker = (key) => state.sortKey === key ? (state.sortDir > 0 ? " ▲" : " ▼") : "";
            let html = "<thead><tr><th>Sel</th><th>#</th>";
            headers.forEach(h => {
              html += '<th><button type="button" data-sort="' + esc(h) + '">' + esc(h) + sortMarker(h) + "</button></th>";
            });
            html += "</tr></thead><tbody>";
            pageRows.forEach((item, i) => {
              const checked = selectedSet.has(item.rowId) ? " checked" : "";
              html += '<tr data-row-id="' + item.rowId + '" class="' + (checked ? "selected" : "") + '">';
              html += '<td><input type="checkbox" data-row-select="' + item.rowId + '"' + checked + "></td>";
              html += "<td>" + (start + i + 1) + "</td>";
              headers.forEach(h => {
                html += "<td>" + esc(item.values?.[h] ?? "") + "</td>";
              });
              html += "</tr>";
            });
            html += "</tbody>";
            table.innerHTML = html;

            pageInfo.textContent = "Page " + state.page + "/" + totalPages;
            status.textContent = selectedSet.size + " selected • " + filtered.length + " visible";
            prevBtn.disabled = state.page <= 1;
            nextBtn.disabled = state.page >= totalPages;

            table.querySelectorAll("button[data-sort]").forEach(btn => {
              btn.addEventListener("click", () => {
                const key = btn.getAttribute("data-sort");
                if (!key) return;
                if (state.sortKey === key) state.sortDir *= -1;
                else { state.sortKey = key; state.sortDir = 1; }
                state.page = 1;
                render();
              });
            });

            table.querySelectorAll("input[data-row-select]").forEach(input => {
              input.addEventListener("change", () => {
                const rowId = Number(input.getAttribute("data-row-select"));
                if (!Number.isFinite(rowId)) return;
                if (input.checked) selectedSet.add(rowId);
                else selectedSet.delete(rowId);
                pushSelectionToOpener();
                render();
              });
            });

            table.querySelectorAll("tbody tr[data-row-id]").forEach(tr => {
              tr.addEventListener("click", (e) => {
                if (e.target.closest("input")) return;
                const rowId = Number(tr.getAttribute("data-row-id"));
                if (window.opener && Number.isFinite(rowId) && typeof window.opener.focusAttributeRowOnMap === "function") {
                  window.opener.focusAttributeRowOnMap(rowId);
                }
              });
            });
          }

          searchInput.addEventListener("input", () => { state.filterText = searchInput.value || ""; state.page = 1; render(); });
          selectedOnly.addEventListener("change", () => {
            pullSelectionFromOpener();
            state.selectedOnly = !!selectedOnly.checked;
            state.page = 1;
            render();
          });
          prevBtn.addEventListener("click", () => { state.page = Math.max(1, state.page - 1); render(); });
          nextBtn.addEventListener("click", () => { state.page += 1; render(); });
          selectVisibleBtn.addEventListener("click", () => {
            getFilteredRows().forEach(item => selectedSet.add(item.rowId));
            pushSelectionToOpener();
            render();
          });
          clearSelectedBtn.addEventListener("click", () => {
            selectedSet.clear();
            pushSelectionToOpener();
            render();
          });
          zoomSelectedBtn.addEventListener("click", () => {
            pullSelectionFromOpener();
            const selectedIds = [...selectedSet];
            if (!selectedIds.length) return;
            if (window.opener && typeof window.opener.zoomToAttributeRowsOnMap === "function") {
              const ok = window.opener.zoomToAttributeRowsOnMap(selectedIds);
              if (ok) return;
            }
            if (window.opener && typeof window.opener.focusAttributeRowOnMap === "function") {
              window.opener.focusAttributeRowOnMap(selectedIds[0]);
            }
          });
          exportCsvBtn.addEventListener("click", () => exportCsv(getFilteredRows()));

          window.addEventListener("focus", () => {
            if (pullSelectionFromOpener()) render();
          });
          setInterval(() => {
            const changed = pullSelectionFromOpener();
            if (changed) render();
          }, 500);

          pullSelectionFromOpener();
          render();
        <\/script>
      </body>
    </html>
  `);
  win.document.close();
}

async function prepareMappedWorkbookForUpload(buffer, fileName) {
  const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  if (!rows.length) return { wb, rows };

  const headerSet = new Set();
  rows.slice(0, Math.min(rows.length, 25)).forEach(row => {
    Object.keys(row || {}).forEach(k => {
      if (k) headerSet.add(k);
    });
  });
  const headers = [...headerSet];
  const guess = buildColumnMappingGuess(headers);
  const saved = loadSavedColumnMapping(headers) || {};
  const initial = {};
  COLUMN_MAPPING_FIELDS.forEach(field => {
    const savedValue = saved[field.key];
    initial[field.key] = headers.includes(savedValue) ? savedValue : (guess[field.key] || "");
  });
  const sampleByHeader = {};
  headers.forEach(h => {
    const row = rows.find(r => String(r?.[h] ?? "").trim() !== "");
    sampleByHeader[h] = row ? row[h] : "";
  });

  const mapping = await openColumnMappingPrompt(headers, initial, fileName, sampleByHeader);
  if (!mapping) return null;

  applyColumnAliasesToRows(rows, mapping);
  return { wb, rows };
}

// ================= PROCESS ROUTE EXCEL =================
function processExcelBuffer(buffer, preMappedRows = null, preMappedWorkbook = null) {
  const wb = preMappedWorkbook || XLSX.read(new Uint8Array(buffer), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];

  const rows = preMappedRows || XLSX.utils.sheet_to_json(ws);

  // store globally for saving later
  window._currentRows = rows;
  window._currentWorkbook = wb;
  window._attributeHeaders = getAttributeHeaders(rows);
  attributeState.page = 1;
  attributeState.selectedRowIds.clear();
  syncSelectedStopsHeaderCount(0);
  attributeRowToMarker = new WeakMap();
  attributeMarkerByRowId.clear();
  rows.forEach((row, rowIndex) => attributeRowToId.set(row, rowIndex));

  // Clear previous map data
  Object.values(routeDayGroups).forEach(g => g.layers.forEach(l => map.removeLayer(l)));
  Object.keys(routeDayGroups).forEach(k => delete routeDayGroups[k]);
  Object.keys(symbolMap).forEach(k => delete symbolMap[k]);
  symbolIndex = 0;
  globalBounds = L.latLngBounds();

  const routeSet = new Set();

  rows.forEach((row, rowIndex) => {
    const lat = Number(row.LATITUDE);
    const lon = Number(row.LONGITUDE);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;

    const route = String(row.NEWROUTE ?? "").trim() || "Unassigned";
    const day = String(row.NEWDAY ?? "").trim() || "No Day";

    const key = `${route}|${day}`;


    const symbol = getSymbol(key);

    if (!routeDayGroups[key]) routeDayGroups[key] = { layers: [] };

  // Build full street address safely
const fullAddress = [
  row["CSADR#"] || "",
  row["CSSDIR"] || "",
  row["CSSTRT"] || "",
  row["CSSFUX"] || ""
].join(" ").replace(/\s+/g, " ").trim();

// Solution Reviewer popup: route + day only.
const popupContent = `
  <div style="font-size:14px; line-height:1.4;">
    <div><strong>Route:</strong> ${route || "-"}</div>
    <div><strong>Day:</strong> ${dayName(day) || day || "-"}</div>
  </div>
`;

const marker = createMarker(lat, lon, symbol)
  .bindPopup(popupContent)
  .addTo(map);
// ===== STREET LABEL (ZOOM-BASED) =====
const streetNumber = row["CSADR#"] ? String(row["CSADR#"]).trim() : "";
const streetName = row["CSSTRT"] ? String(row["CSSTRT"]).trim() : "";

const labelText = `${streetNumber} ${streetName}`.trim();

if (labelText) {
  marker.bindTooltip(labelText, {
    permanent: false,
    direction: "top",
    offset: [0, -8],
    className: "street-label"
  });

  marker._hasStreetLabel = true;
}


    // 🔥 CRITICAL: link marker to Excel row
    marker._rowRef = row;
    marker._rowId = rowIndex;
    attributeRowToMarker.set(row, marker);
    attributeMarkerByRowId.set(rowIndex, marker);
    routeDayGroups[key].layers.push(marker);
    routeSet.add(route);
    globalBounds.extend([lat, lon]);
  });

  buildRouteCheckboxes([...routeSet]);
  buildRouteDayLayerControls();
  applyFilters();
  renderAttributeTable();
  refreshAttributeStatus();

  if (globalBounds.isValid()) map.fitBounds(globalBounds);
}



// ================= LIST FILES FROM CLOUD =================
const SAVED_FILES_SORT_MODE_KEY = "savedFilesSortMode";

function normalizeSavedFilesSortMode(value) {
  const mode = String(value || "").trim().toLowerCase();
  if (mode === "date-desc" || mode === "date-asc" || mode === "summary-name") return mode;
  return "summary-name";
}

function getSavedFileAddedTimestamp(file) {
  const candidates = [
    file?.created_at,
    file?.createdAt,
    file?.updated_at,
    file?.last_accessed_at,
    file?.metadata?.created_at,
    file?.metadata?.updated_at,
    file?.metadata?.lastModified
  ];
  for (const value of candidates) {
    if (!value) continue;
    const ts = Date.parse(String(value));
    if (Number.isFinite(ts)) return ts;
  }
  return 0;
}

async function listFiles() {
  const { data, error } = await sb.storage.from(BUCKET).list();
  if (error) return console.error(error);

  const ul = document.getElementById("savedFiles");
  ul.innerHTML = "";
  const routeFiles = data.filter(file => !isRouteSummaryFileName(file.name));
  const allFileNames = data.map(f => f.name);
  cleanupSummaryAttachments(allFileNames);
  const sortSelect = document.getElementById("savedFilesSortSelect");
  const storedSortMode = normalizeSavedFilesSortMode(storageGet(SAVED_FILES_SORT_MODE_KEY));
  const activeSortMode = normalizeSavedFilesSortMode(sortSelect?.value || storedSortMode);
  if (sortSelect) sortSelect.value = activeSortMode;
  storageSet(SAVED_FILES_SORT_MODE_KEY, activeSortMode);

  const routeEntries = routeFiles
    .map(file => {
      const routeName = file.name;
      const summaryName = resolveSummaryForRoute(routeName, data);
      const addedTimestamp = getSavedFileAddedTimestamp(file);
      return { routeName, summaryName, addedTimestamp };
    })
    .sort((a, b) => {
      if (activeSortMode === "date-desc") {
        const delta = b.addedTimestamp - a.addedTimestamp;
        if (delta !== 0) return delta;
        return a.routeName.localeCompare(b.routeName, undefined, { numeric: true, sensitivity: "base" });
      }
      if (activeSortMode === "date-asc") {
        const delta = a.addedTimestamp - b.addedTimestamp;
        if (delta !== 0) return delta;
        return a.routeName.localeCompare(b.routeName, undefined, { numeric: true, sensitivity: "base" });
      }
      const aHasSummary = !!a.summaryName;
      const bHasSummary = !!b.summaryName;
      if (aHasSummary !== bHasSummary) return bHasSummary ? 1 : -1;
      return a.routeName.localeCompare(b.routeName, undefined, { numeric: true, sensitivity: "base" });
    });

  routeEntries.forEach(({ routeName, summaryName }) => {
    const li = document.createElement("li");
    li.className = `saved-route-item${summaryName ? " has-summary" : ""}`;

    const infoWrap = document.createElement("div");
    infoWrap.className = "saved-route-info";

    const nameNode = document.createElement("div");
    nameNode.className = "saved-route-name";
    nameNode.textContent = routeName;

    const metaRow = document.createElement("div");
    metaRow.className = "saved-route-meta";
    const badge = document.createElement("span");
    badge.className = `saved-summary-badge${summaryName ? "" : " missing"}`;
    badge.textContent = summaryName ? "Summary Attached" : "No Summary";
    metaRow.appendChild(badge);

    if (summaryName) {
      const summaryFileNode = document.createElement("span");
      summaryFileNode.className = "saved-summary-file";
      summaryFileNode.textContent = summaryName;
      metaRow.appendChild(summaryFileNode);
    }

    infoWrap.append(nameNode, metaRow);

    const actions = document.createElement("div");
    actions.className = "saved-route-actions";

    const openBtn = document.createElement("button");
    openBtn.className = "saved-file-btn open-btn";
    openBtn.textContent = "Open Map";
    openBtn.onclick = async () => {
      try {
        showLoading("Loading Excel file...");

        const { data } = sb.storage.from(BUCKET).getPublicUrl(routeName);
        const urlWithBypass = data.publicUrl + "?v=" + Date.now();
        const r = await fetch(urlWithBypass, {
          cache: "no-store"
        });

        window._currentFilePath = routeName;
        setCurrentFileDisplay(window._currentFilePath);

        processExcelBuffer(await r.arrayBuffer());
        await loadSummaryFor(routeName);

        hideLoading("File Loaded Successfully ✅");
        if (fileManagerModal) fileManagerModal.style.display = "none";
      } catch (err) {
        console.error(err);
        hideLoading();
        alert("Error loading file.");
      }
    };
    actions.appendChild(openBtn);

    if (summaryName) {
      const summaryBtn = document.createElement("button");
      summaryBtn.className = "saved-file-btn summary-btn";
      summaryBtn.textContent = "Summary";
      summaryBtn.onclick = async () => {
        await loadSummaryFor(routeName);
        if (fileManagerModal) fileManagerModal.style.display = "none";
      };
      actions.appendChild(summaryBtn);
    } else {
      const summarySpacer = document.createElement("span");
      summarySpacer.className = "saved-file-btn-spacer";
      summarySpacer.setAttribute("aria-hidden", "true");
      actions.appendChild(summarySpacer);
    }

    const delBtn = document.createElement("button");
    delBtn.className = "saved-file-btn delete-btn";
    delBtn.textContent = "Delete";
    delBtn.onclick = async () => {
      const entered = prompt("Enter password to delete this file:");
      if (entered !== DELETE_PASSWORD) {
        alert("❌ Incorrect password. File not deleted.");
        return;
      }

      const confirmed = confirm("Are you sure you want to permanently delete this file?");
      if (!confirmed) return;

      const toDelete = [routeName];
      if (summaryName) toDelete.push(summaryName);

      await sb.storage.from(BUCKET).remove(toDelete);
      removeRouteSummaryAttachment(routeName);
      if (summaryName) {
        const map = getSummaryAttachments();
        Object.keys(map).forEach(routeKey => {
          if (map[routeKey] === summaryName) delete map[routeKey];
        });
        setSummaryAttachments(map);
      }

      alert("✅ File deleted successfully.");
      listFiles();
    };
    actions.appendChild(delBtn);

    li.append(infoWrap, actions);
    ul.appendChild(li);
  });
}


// ================= UPLOAD FILE =================
async function uploadFile(file) {
  if (!file) return;

  try {

    const fileBuffer = await file.arrayBuffer();
    const mappedWorkbook = await prepareMappedWorkbookForUpload(fileBuffer, file.name);
    if (!mappedWorkbook) {
      return;
    }
    showLoading("Uploading file...");

    const { error } = await sb.storage
      .from(BUCKET)
      .upload(file.name, file, { upsert: true });

    if (error) {
      throw error;
    }

    window._currentFilePath = file.name;
    setCurrentFileDisplay(window._currentFilePath);

    processExcelBuffer(fileBuffer, mappedWorkbook.rows, mappedWorkbook.wb);
    listFiles();

    hideLoading("Upload Complete ✅");

  } catch (error) {
    console.error("UPLOAD ERROR:", error);
    hideLoading();
    alert("Upload failed: " + error.message);
  }
}

async function uploadRouteSummaryAndAttach(file) {
  if (!file) return;

  try {
    showLoading("Uploading route summary...");

    const { error } = await sb.storage
      .from(BUCKET)
      .upload(file.name, file, { upsert: true });

    if (error) throw error;

    const { data: files, error: listErr } = await sb.storage.from(BUCKET).list();
    if (listErr) throw listErr;

    const routeFileNames = files
      .filter(f => !isRouteSummaryFileName(f.name))
      .map(f => f.name);

    hideLoading();

    if (!routeFileNames.length) {
      alert("Summary uploaded, but no route files are saved yet. Upload a route file first, then attach this summary.");
      listFiles();
      return;
    }

    const selectedRoute = await openSummaryAttachModal(file.name, routeFileNames);
    if (!selectedRoute) {
      listFiles();
      return;
    }

    setRouteSummaryAttachment(selectedRoute, file.name);
    alert(`Attached ${file.name} to ${selectedRoute}.`);
    listFiles();
  } catch (error) {
    console.error("SUMMARY UPLOAD ERROR:", error);
    hideLoading();
    alert("Route summary upload failed: " + error.message);
  }
}

// ================= ROUTE SUMMARY DISPLAY =================
function showRouteSummary(rows, headers) {
  const tableBox = document.getElementById("routeSummaryTable");
  const panel = document.getElementById("bottomSummary");
  const btn = document.getElementById("summaryToggleBtn");

  if (!tableBox || !panel || !btn) return;

  tableBox.innerHTML = "";
  window._summaryRows = Array.isArray(rows) ? rows : [];
  window._summaryHeaders = Array.isArray(headers) ? headers : [];

  if (!rows || !rows.length) {
    tableBox.textContent = "No summary data found";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  const headerRow = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h ?? "";
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  rows.forEach(r => {
    const tr = document.createElement("tr");
    headers.forEach(h => {
      const td = document.createElement("td");
      td.textContent = r[h] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  tableBox.appendChild(table);

  const savedHeight = Number(storageGet("summaryHeight"));
  const defaultHeight = window.innerWidth <= 900 ? 300 : 250;
  const targetHeight =
    Number.isFinite(savedHeight) && savedHeight > 60 ? savedHeight : defaultHeight;

  // Always show summary immediately after loading.
  panel.classList.remove("collapsed");
  btn.textContent = "\u25BC";
  panel.style.height = `${targetHeight}px`;

  // iOS standalone occasionally needs a repaint after dynamic table injection.
  if (window.innerWidth <= 900) {
    requestAnimationFrame(() => {
      panel.style.height = `${targetHeight + 1}px`;
      requestAnimationFrame(() => {
        panel.style.height = `${targetHeight}px`;
      });
    });
  }
}
function autoCollapseSidebarsForSummary() {
  const appContainer = document.querySelector(".app-container");
  const toggleSidebarBtn = document.getElementById("toggleSidebarBtn");
  const sidebar = document.querySelector(".sidebar");
  const overlay = document.querySelector(".mobile-overlay");
  const mobileMenuBtn = document.getElementById("mobileMenuBtn");
  const isMobile = window.innerWidth <= 900;

  const selectionBox = document.getElementById("selectionBox");
  const toggleSelectionBtn = document.getElementById("toggleSelectionBtn");

  // Left sidebar: collapse desktop and close mobile drawer.
  if (appContainer) {
    if (isMobile) {
      appContainer.classList.remove("collapsed");
    } else {
      appContainer.classList.add("collapsed");
    }
  }
  if (toggleSidebarBtn) toggleSidebarBtn.setAttribute("aria-expanded", "false");
  if (sidebar) sidebar.classList.remove("open");
  if (overlay) overlay.classList.remove("show");
  if (mobileMenuBtn) mobileMenuBtn.textContent = "☰";

  // Right sidebar: collapse desktop and hide mobile panel.
  if (selectionBox) {
    selectionBox.classList.add("collapsed");
    selectionBox.classList.remove("show");
  }
  if (toggleSelectionBtn) toggleSelectionBtn.textContent = "❮";

  setTimeout(() => map.invalidateSize(), 180);
}



// Load matching summary file
async function loadSummaryFor(routeFileName) {
  const { data, error } = await sb.storage.from(BUCKET).list();
  if (error) {
    console.error("LIST ERROR:", error);
    return;
  }

  const summaryName = resolveSummaryForRoute(routeFileName, data);
  if (!summaryName) {
    document.getElementById("routeSummaryTable").textContent = "No summary available";
    return;
  }

  const { data: urlData } = sb.storage.from(BUCKET).getPublicUrl(summaryName);
  const r = await fetch(urlData.publicUrl);

  const wb = XLSX.read(new Uint8Array(await r.arrayBuffer()), { type: "array" });
const ws = wb.Sheets[wb.SheetNames[0]];

// Read entire sheet as grid
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

// ===== FIND FIRST NON-EMPTY ROW =====
let startRow = raw.findIndex(r =>
  r && r.some(cell => String(cell || "").trim() !== "")
);

if (startRow === -1) {
  showRouteSummary([], []);
  return;
}

// ===== DETECT MULTI-ROW HEADERS (supports 1–3+) =====
let headerRows = [raw[startRow]];
let nextRow = raw[startRow + 1];
let thirdRow = raw[startRow + 2];

function looksLikeHeader(row) {
  if (!row) return false;

  const filled = row.filter(c => String(c || "").trim() !== "").length;
  const numeric = row.filter(c => !isNaN(parseFloat(c))).length;

  return filled > 0 && numeric < filled / 2;
}

if (looksLikeHeader(nextRow)) headerRows.push(nextRow);
if (looksLikeHeader(thirdRow)) headerRows.push(thirdRow);

// ===== BUILD SAFE COLUMN NAMES =====
const columnCount = Math.max(...headerRows.map(r => r.length));

const headers = Array.from({ length: columnCount }, (_, col) => {
  const parts = headerRows
    .map(r => String(r[col] || "").trim())
    .filter(Boolean);

  return parts.join(" ") || `Column ${col + 1}`;
});

// ===== DATA STARTS AFTER HEADER =====
const dataStartIndex = startRow + headerRows.length;

// ===== BUILD ROW OBJECTS =====
const rows = raw.slice(dataStartIndex).map(r => {
  const obj = {};
  headers.forEach((h, i) => {
    obj[h] = r?.[i] ?? "";
  });
  return obj;
});




showRouteSummary(rows, headers);
autoCollapseSidebarsForSummary();


// 🔽 FORCE the panel open when a summary exists
const panel = document.getElementById("bottomSummary");
const btn = document.getElementById("summaryToggleBtn");

const isMobile = window.innerWidth <= 900;

if (panel && btn && !isMobile) {
  panel.classList.remove("collapsed");
  btn.textContent = "▼";
}



}



// ================= START APP =================
// ===== TOGGLE BOTTOM SUMMARY =====
function toggleSummary() {
  const panel = document.getElementById("bottomSummary");
  const btn = document.getElementById("summaryToggleBtn");

  panel.classList.toggle("collapsed");

  // flip arrow direction
  btn.textContent = panel.classList.contains("collapsed") ? "▲" : "▼";
}
// ===== PLACE LOCATE BUTTON BASED ON SCREEN SIZE =====
function placeLocateButton() {
  const locateBtn = document.getElementById("locateMeBtn");
    const headerContainer = document.querySelector(".mobile-header-buttons");
  const desktopContainer = document.getElementById("desktopLocateContainer");
    const streetToggle = document.getElementById("streetLabelToggle");

  if (!locateBtn || !headerContainer || !desktopContainer) return;

  if (window.innerWidth <= 900) {
    // 📱 MOBILE
    headerContainer.appendChild(locateBtn);

    if (streetToggle) {
      headerContainer.appendChild(streetToggle.parentElement);
    }

  } else {
    // 🖥 DESKTOP
    desktopContainer.appendChild(locateBtn);

    if (streetToggle) {
      desktopContainer.appendChild(streetToggle.parentElement);
    }
  }
}

//undo button state
function updateUndoButtonState() {
  // Solution Reviewer does not use undo state.
  return;
}









function initApp() { //begining of initApp=================================================================

// ===== RIGHT SIDEBAR TOGGLE =====

// ===== RIGHT SIDEBAR TOGGLE =====
const selectionBox = document.getElementById("selectionBox");
const toggleSelectionBtn = document.getElementById("toggleSelectionBtn");
const clearBtn = document.getElementById("clearSelectionBtn");
const selectedStopsLabel = document.getElementById("selectedStopsLabel");
const selectionCountNode = document.getElementById("selectionCount");
const desktopSelectionHeader = document.getElementById("desktopSelectionHeader");
const pageHeader = document.querySelector("header");
const selectionDrawModeSelect = document.getElementById("selectionDrawModeSelect");
const selectionModeBadge = document.getElementById("selectionModeBadge");
const selectionToolsStatus = document.getElementById("selectionToolsStatus");
const selectionDrawPolygonBtn = document.getElementById("selectionDrawPolygonBtn");
const selectionDrawRectangleBtn = document.getElementById("selectionDrawRectangleBtn");
const selectionSelectInViewBtn = document.getElementById("selectionSelectInViewBtn");
const selectionSelectVisibleBtn = document.getElementById("selectionSelectVisibleBtn");
const selectionInvertVisibleBtn = document.getElementById("selectionInvertVisibleBtn");
const selectionClearShapeBtn = document.getElementById("selectionClearShapeBtn");

// Start with the right sidebar closed on initial page load.
if (selectionBox) selectionBox.classList.add("collapsed");
if (toggleSelectionBtn) toggleSelectionBtn.textContent = "❮";

// Toggle sidebar open/closed
if (selectionBox && toggleSelectionBtn) {
  toggleSelectionBtn.onclick = () => {
    const collapsed = selectionBox.classList.toggle("collapsed");
    toggleSelectionBtn.textContent = collapsed ? "❮" : "❯";
  };
}

function syncSelectionBoxTop() {
  if (!selectionBox || !pageHeader) return;
  const headerHeight = Math.ceil(pageHeader.getBoundingClientRect().height);
  if (window.innerWidth <= 900) {
    selectionBox.style.top = "";
    selectionBox.style.maxHeight = "";
    if (toggleSelectionBtn) toggleSelectionBtn.style.top = "";
    return;
  }

  // Keep sidebar fully below sticky header on desktop.
  const topOffset = headerHeight + 8;
  selectionBox.style.top = `${topOffset}px`;
  selectionBox.style.maxHeight = `calc(100vh - ${topOffset + 12}px)`;
  if (toggleSelectionBtn) toggleSelectionBtn.style.top = `${topOffset + 8}px`;
}

function placeDesktopSelectionControls() {
  if (
    !selectionBox ||
    !clearBtn ||
    !selectedStopsLabel ||
    !selectionCountNode ||
    !desktopSelectionHeader
  ) return;

  const desktopLocateContainer = document.getElementById("desktopLocateContainer");

  if (window.innerWidth > 900) {
    desktopSelectionHeader.appendChild(selectedStopsLabel);
    desktopSelectionHeader.appendChild(selectionCountNode);
    desktopSelectionHeader.appendChild(clearBtn);
  } else {
    selectionBox.insertBefore(selectedStopsLabel, selectionBox.firstChild);
    selectionBox.insertBefore(selectionCountNode, selectedStopsLabel.nextSibling);
    if (desktopLocateContainer) {
      selectionBox.insertBefore(clearBtn, desktopLocateContainer);
    } else {
      selectionBox.appendChild(clearBtn);
    }
  }
}

syncSelectionBoxTop();
placeDesktopSelectionControls();
window.addEventListener("resize", syncSelectionBoxTop);
window.addEventListener("resize", placeDesktopSelectionControls);

const setSelectionToolsStatus = message => {
  if (!selectionToolsStatus) return;
  selectionToolsStatus.textContent = String(message || "");
};

const refreshSelectionToolsUi = () => {
  const mode = normalizeMapSelectionDrawMode(mapSelectionDrawMode);
  const modeLabel = getMapSelectionModeLabel(mode);
  const targetLabel = attributeTableMode === "streets" ? "Street Segments" : "Record Stops";
  if (selectionDrawModeSelect) selectionDrawModeSelect.value = mode;
  if (selectionModeBadge) selectionModeBadge.textContent = modeLabel;
  setSelectionToolsStatus(`Target: ${targetLabel}. Mode: ${modeLabel}.`);
};

window.__refreshSelectionToolsUi = refreshSelectionToolsUi;
refreshSelectionToolsUi();

selectionDrawModeSelect?.addEventListener("change", () => {
  setMapSelectionDrawMode(selectionDrawModeSelect.value, true);
  refreshSelectionToolsUi();
});

selectionDrawPolygonBtn?.addEventListener("click", () => {
  const started = startSelectionDrawTool("polygon");
  if (!started) {
    setSelectionToolsStatus("Unable to start polygon tool.");
    return;
  }
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  setSelectionToolsStatus(`Polygon draw started (${modeLabel}).`);
});

selectionDrawRectangleBtn?.addEventListener("click", () => {
  const started = startSelectionDrawTool("rectangle");
  if (!started) {
    setSelectionToolsStatus("Unable to start rectangle tool.");
    return;
  }
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  setSelectionToolsStatus(`Rectangle draw started (${modeLabel}).`);
});

selectionSelectInViewBtn?.addEventListener("click", () => {
  const ids = attributeTableMode === "streets"
    ? getStreetIdsInCurrentView()
    : getRecordRowIdsInCurrentView();
  const total = applySelectionIdsToActiveMode(ids, mapSelectionDrawMode);
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  setSelectionToolsStatus(`Select In View applied (${modeLabel}). ${total.toLocaleString()} selected.`);
});

selectionSelectVisibleBtn?.addEventListener("click", () => {
  const ids = attributeTableMode === "streets"
    ? getVisibleSelectableStreetIds()
    : getVisibleSelectableRecordRowIds();
  const total = applySelectionIdsToActiveMode(ids, mapSelectionDrawMode);
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  setSelectionToolsStatus(`Select All Visible applied (${modeLabel}). ${total.toLocaleString()} selected.`);
});

selectionInvertVisibleBtn?.addEventListener("click", () => {
  const total = invertVisibleSelectionInActiveMode();
  setSelectionToolsStatus(`Invert Visible applied. ${total.toLocaleString()} selected.`);
});

selectionClearShapeBtn?.addEventListener("click", () => {
  clearDrawnSelectionGeometry();
  setSelectionToolsStatus("Drawn selection shape cleared.");
});

// Clear selection button (ALWAYS ACTIVE)
if (clearBtn) {
  clearBtn.onclick = () => {
    selectedLayerKey = null;
    // Remove polygon
    drawnLayer.clearLayers();
    attributeState.selectedRowIds.clear();
    streetAttributeSelectedIds.clear();
    applyStreetSelectionStyles();

    


    // Restore original marker colors
    Object.entries(routeDayGroups).forEach(([key, group]) => {
      const sym = symbolMap[key];
      group.layers.forEach(marker => {
        marker.setStyle?.({ color: sym.color, fillColor: sym.color });
      });
    });

    // 🔥 Force counter refresh everywhere (desktop + mobile)
    updateSelectionCount();
    renderAttributeTable();
    updateUndoButtonState();
    setSelectionToolsStatus("Selection cleared.");
  };
}







  
// ===== IMPORT WIZARD + FILE UPLOAD =====
const attributePanel = document.getElementById("attributeTablePanel");
const streetAttributesBtnDesktop = document.getElementById("streetAttributesBtn");
const streetAttributesBtnMobile = document.getElementById("streetAttributesBtnMobile");
const attributeBtnDesktop = document.getElementById("attributeTableBtn");
const attributeBtnMobile = document.getElementById("attributeTableBtnMobile");
const attributeCloseBtn = document.getElementById("attributeDockCloseBtn");
const attributePopoutBtn = document.getElementById("attributePopoutBtn");
const attributeSearchInput = document.getElementById("attributeSearchInput");
const attributeSelectedOnly = document.getElementById("attributeSelectedOnly");
const attributeSelectVisibleBtn = document.getElementById("attributeSelectVisibleBtn");
const attributeClearSelectionBtn = document.getElementById("attributeClearSelectionBtn");
const attributeZoomSelectedBtn = document.getElementById("attributeZoomSelectedBtn");
const attributeExportCsvBtn = document.getElementById("attributeExportCsvBtn");
const attributePrevPageBtn = document.getElementById("attributePrevPageBtn");
const attributeNextPageBtn = document.getElementById("attributeNextPageBtn");
const attributeHeaderBar = attributePanel?.querySelector(".attribute-table-header");
const attributeResizeHandles = attributePanel
  ? [...attributePanel.querySelectorAll(".attribute-resize-handle[data-dir]")]
  : [];

if (attributePanel) {
  attributePanel.classList.add("closed");
  attributePanel.classList.remove("collapsed");
  const interaction = {
    mode: null,
    dir: "",
    startX: 0,
    startY: 0,
    startLeft: 0,
    startTop: 0,
    startW: 0,
    startH: 0,
    popoutIntent: false
  };

  if (attributeHeaderBar) {
    attributeHeaderBar.addEventListener("pointerdown", e => {
      if (attributePanel.classList.contains("closed") || attributePanel.classList.contains("collapsed")) return;
      if (e.target.closest("button, input, label")) return;
      interaction.mode = "move";
      interaction.startX = e.clientX;
      interaction.startY = e.clientY;
      interaction.startLeft = attributePanel.offsetLeft;
      interaction.startTop = attributePanel.offsetTop;
      interaction.startW = attributePanel.offsetWidth;
      interaction.startH = attributePanel.offsetHeight;
      interaction.popoutIntent = false;
      document.body.style.userSelect = "none";
      attributePanel.classList.add("dragging");
      attributeHeaderBar.setPointerCapture?.(e.pointerId);
      e.preventDefault();
    });
  }

  attributeResizeHandles.forEach(handle => {
    handle.addEventListener("pointerdown", e => {
      if (attributePanel.classList.contains("closed") || attributePanel.classList.contains("collapsed")) return;
      interaction.mode = "resize";
      interaction.dir = handle.dataset.dir || "se";
      interaction.startX = e.clientX;
      interaction.startY = e.clientY;
      interaction.startLeft = attributePanel.offsetLeft;
      interaction.startTop = attributePanel.offsetTop;
      interaction.startW = attributePanel.offsetWidth;
      interaction.startH = attributePanel.offsetHeight;
      document.body.style.userSelect = "none";
      handle.setPointerCapture?.(e.pointerId);
      e.preventDefault();
    });
  });

  document.addEventListener("pointermove", e => {
    if (!interaction.mode || attributePanel.classList.contains("closed")) return;
    let left = interaction.startLeft;
    let top = interaction.startTop;
    let width = interaction.startW;
    let height = interaction.startH;

    const dx = e.clientX - interaction.startX;
    const dy = e.clientY - interaction.startY;

    if (interaction.mode === "move") {
      const rawLeft = interaction.startLeft + dx;
      const rawTop = interaction.startTop + dy;
      const rawRight = rawLeft + interaction.startW;
      const rawBottom = rawTop + interaction.startH;
      const bounds = getAttributePanelBounds();
      const detachMargin = 80;

      interaction.popoutIntent =
        rawLeft < bounds.left - detachMargin ||
        rawTop < bounds.top - detachMargin ||
        rawRight > bounds.right + detachMargin ||
        rawBottom > bounds.bottom + detachMargin;

      attributePanel.classList.toggle("detach-ready", interaction.popoutIntent);
      left = rawLeft;
      top = rawTop;
    } else if (interaction.mode === "resize") {
      // Keep resize smooth and direct while still slightly boosted.
      const cornerDrag = interaction.dir.length === 2;
      const speed = cornerDrag ? 1.35 : 1.2;
      const sx = dx * speed;
      const sy = dy * speed;

      if (interaction.dir.includes("e")) {
        width = interaction.startW + sx;
      }
      if (interaction.dir.includes("s")) {
        height = interaction.startH + sy;
      }
      if (interaction.dir.includes("w")) {
        width = interaction.startW - sx;
        left = interaction.startLeft + sx;
      }
      if (interaction.dir.includes("n")) {
        height = interaction.startH - sy;
        top = interaction.startTop + sy;
      }
    }

    if (interaction.mode === "move" && interaction.popoutIntent) {
      // Let the panel follow the cursor outside bounds for intuitive "detach" feel.
      attributePanel.style.left = `${Math.round(left)}px`;
      attributePanel.style.top = `${Math.round(top)}px`;
      attributePanel.style.width = `${Math.round(interaction.startW)}px`;
      attributePanel.style.height = `${Math.round(interaction.startH)}px`;
    } else {
      const rect = clampAttributePanelRect(left, top, width, height);
      applyAttributePanelRect(attributePanel, rect);
    }
  });

  document.addEventListener("pointerup", () => {
    if (!interaction.mode) return;
    if (interaction.mode === "move" && interaction.popoutIntent) {
      interaction.mode = null;
      interaction.dir = "";
      interaction.popoutIntent = false;
      document.body.style.userSelect = "";
      attributePanel.classList.remove("dragging", "detach-ready");
      openAttributeTablePopout();
      closeAttributePanel();
      return;
    }
    if (interaction.mode === "move") {
      snapAttributePanelToCorner(attributePanel);
    }
    interaction.mode = null;
    interaction.dir = "";
    interaction.popoutIntent = false;
    document.body.style.userSelect = "";
    attributePanel.classList.remove("dragging", "detach-ready");
    saveAttributePanelRect(attributePanel);
    refreshMapAfterOverlayChange();
  });
}

const toggleAttributePanel = () => {
  if (!attributePanel) return;
  const isClosed = attributePanel.classList.contains("closed");
  if (isClosed) {
    setAttributeTableMode("records");
    openAttributePanel();
    renderAttributeTable();
    return;
  }
  if (attributeTableMode !== "records") {
    setAttributeTableMode("records");
    renderAttributeTable();
    return;
  }
  closeAttributePanel();
};

const openStreetAttributesPanel = async () => {
  if (!attributePanel) return;
  const layerToggle = document.getElementById("useLocalStreetSource");
  const isClosed = attributePanel.classList.contains("closed");
  // Toggle close when Street Attributes is already the active open panel.
  if (!isClosed && attributeTableMode === "streets") {
    closeAttributePanel();
    return;
  }
  setAttributeTableMode("streets");
  if (isClosed) openAttributePanel();
  if (layerToggle?.checked) {
    if (streetAttributeById.size) {
      updateStreetLoadStatus(`Showing ${streetAttributeById.size.toLocaleString()} loaded street segments.`);
    } else {
      const scopedPolygonFromSnapshot = createStreetPolygonLayerFromSnapshot(lastStreetLoadPolygonSnapshot);
      const selectionPolygonLayer = (drawnLayer?.getLayers?.() || [])
        .find(layer => layer instanceof L.Polygon || layer instanceof L.Rectangle) || null;
      const scopedPolygon = scopedPolygonFromSnapshot || selectionPolygonLayer || null;
      if (scopedPolygon) {
        await loadStreetAttributesForCurrentView(scopedPolygon.getBounds(), scopedPolygon);
      } else {
        updateStreetLoadStatus("No street segments loaded yet. Turn on 'Use Local Streets File', then choose a saved polygon or draw a new one.", true);
      }
    }
  } else {
    updateStreetLoadStatus("Enable 'Street Segments (Local Source)' to load and view street segments.", true);
  }
  renderAttributeTable();
};

streetAttributesBtnDesktop?.addEventListener("click", openStreetAttributesPanel);
streetAttributesBtnMobile?.addEventListener("click", openStreetAttributesPanel);
attributeBtnDesktop?.addEventListener("click", toggleAttributePanel);
attributeBtnMobile?.addEventListener("click", toggleAttributePanel);

attributeCloseBtn?.addEventListener("click", () => {
  if (!attributePanel) return;
  closeAttributePanel();
});

attributePopoutBtn?.addEventListener("click", openAttributeTablePopout);

attributeSearchInput?.addEventListener("input", () => {
  attributeState.filterText = attributeSearchInput.value || "";
  attributeState.page = 1;
  renderAttributeTable();
});

attributeSelectedOnly?.addEventListener("change", () => {
  attributeState.selectedOnly = !!attributeSelectedOnly.checked;
  attributeState.page = 1;
  renderAttributeTable();
});

attributeSelectVisibleBtn?.addEventListener("click", () => {
  if (attributeTableMode === "streets") {
    const visibleStreetIds = (attributeState.lastVisibleRows || [])
      .map(item => Number(item?.rowId))
      .filter(id => Number.isFinite(id) && streetAttributeById.has(id));
    visibleStreetIds.forEach(id => streetAttributeSelectedIds.add(id));
    applyStreetSelectionStyles();
  } else {
    getFilteredAttributeRows().forEach(({ rowId }) => attributeState.selectedRowIds.add(rowId));
    applyAttributeSelectionStyles();
    syncSelectedStopsHeaderCount(attributeState.selectedRowIds.size);
  }
  renderAttributeTable();
});

attributeClearSelectionBtn?.addEventListener("click", () => {
  if (attributeTableMode === "streets") {
    streetAttributeSelectedIds.clear();
    applyStreetSelectionStyles();
  } else {
    attributeState.selectedRowIds.clear();
    applyAttributeSelectionStyles();
    syncSelectedStopsHeaderCount(0);
  }
  renderAttributeTable();
});

attributeZoomSelectedBtn?.addEventListener("click", () => {
  if (attributeTableMode === "streets") zoomToSelectedStreetSegments();
  else zoomToSelectedAttributeRows();
});
attributeExportCsvBtn?.addEventListener("click", exportAttributeVisibleRowsToCsv);
attributePrevPageBtn?.addEventListener("click", () => {
  attributeState.page = Math.max(1, attributeState.page - 1);
  renderAttributeTable();
});
attributeNextPageBtn?.addEventListener("click", () => {
  attributeState.page = attributeState.page + 1;
  renderAttributeTable();
});

window.addEventListener("resize", () => {
  syncAttributePanelLayout();
  if (!attributePanel || attributePanel.classList.contains("closed")) return;
  const rect = clampAttributePanelRect(
    attributePanel.offsetLeft,
    attributePanel.offsetTop,
    attributePanel.offsetWidth,
    attributePanel.offsetHeight
  );
  applyAttributePanelRect(attributePanel, rect);
  saveAttributePanelRect(attributePanel);
  refreshMapAfterOverlayChange();
});
syncAttributePanelLayout();
setAttributeTableMode("records");
renderAttributeTable();

const importWizardBtn = document.getElementById("importWizardBtn");
const importWizardBtnMobile = document.getElementById("importWizardBtnMobile");
const importWizardModal = document.getElementById("importWizardModal");
const importWizardRouteBtn = document.getElementById("importWizardRouteBtn");
const importWizardSummaryBtn = document.getElementById("importWizardSummaryBtn");
const importWizardClose = document.getElementById("importWizardClose");

// create hidden file inputs dynamically (fallback)
let fileInput = document.getElementById("fileInput");
if (!fileInput) {
  fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".xlsx,.xls,.csv";
  fileInput.id = "fileInput";
  fileInput.hidden = true;
  document.body.appendChild(fileInput);
}

let summaryFileInput = document.getElementById("summaryFileInput");
if (!summaryFileInput) {
  summaryFileInput = document.createElement("input");
  summaryFileInput.type = "file";
  summaryFileInput.accept = ".xlsx,.xls,.csv";
  summaryFileInput.id = "summaryFileInput";
  summaryFileInput.hidden = true;
  document.body.appendChild(summaryFileInput);
}

function closeImportWizardModal() {
  if (importWizardModal) importWizardModal.style.display = "none";
}

if (importWizardBtn && importWizardModal) {
  importWizardBtn.addEventListener("click", () => {
    importWizardModal.style.display = "flex";
  });
}

if (importWizardBtnMobile && importWizardModal) {
  importWizardBtnMobile.addEventListener("click", () => {
    importWizardModal.style.display = "flex";
  });
}

if (importWizardClose) {
  importWizardClose.addEventListener("click", closeImportWizardModal);
}

if (importWizardModal) {
  importWizardModal.addEventListener("click", e => {
    if (e.target === importWizardModal) closeImportWizardModal();
  });
}

if (importWizardRouteBtn) {
  importWizardRouteBtn.addEventListener("click", () => {
    closeImportWizardModal();
    fileInput.click();
  });
}

if (importWizardSummaryBtn) {
  importWizardSummaryBtn.addEventListener("click", () => {
    closeImportWizardModal();
    summaryFileInput.click();
  });
}

// FILE SELECTED
fileInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (file) uploadFile(file);
  e.target.value = "";
});

summaryFileInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (file) uploadRouteSummaryAndAttach(file);
  e.target.value = "";
});

// ===== INITIAL MAP LAYER + USER LOCATION =====
baseMaps.streets.addTo(map);
initLocalStreetSourceControls();
tryRestoreSavedStreetSourceOnStartup().catch(err => {
  console.warn("Unable to restore saved local streets source on startup:", err);
  updateLocalStreetSourceStatus();
});
initStreetNetworkToggle();
initSelectByAttributesControls();
initPrintCenterControls();
initLayerManagerControls();



  
  // ===== BASE MAP DROPDOWN =====
  const baseSelect = document.getElementById("baseMapSelect");
if (baseSelect) {
  baseSelect.addEventListener("change", e => {
    Object.values(baseMaps).forEach(l => map.removeLayer(l));
    map.removeLayer(satelliteLabelsLayer);

    const selected = e.target.value;
    baseMaps[selected].addTo(map);

    if (selected === "satellite" && map.getZoom() >= 15) {
      satelliteLabelsLayer.addTo(map);
    }
    syncStreetNetworkOverlay();
  });
}

  // ===== SIDEBAR TOGGLE (DESKTOP) =====
  const toggleSidebarBtn = document.getElementById("toggleSidebarBtn");
  const sidebar = document.querySelector(".sidebar");
  const appContainer = document.querySelector(".app-container");

  if (toggleSidebarBtn && sidebar && appContainer) {
    toggleSidebarBtn.setAttribute(
      "aria-expanded",
      appContainer.classList.contains("collapsed") ? "false" : "true"
    );

    toggleSidebarBtn.addEventListener("click", () => {
      appContainer.classList.toggle("collapsed");
      toggleSidebarBtn.setAttribute(
        "aria-expanded",
        appContainer.classList.contains("collapsed") ? "false" : "true"
      );
      setTimeout(() => map.invalidateSize(), 200);
    });
  }

// ===== MOBILE MENU =====
const mobileMenuBtn = document.getElementById("mobileMenuBtn");

const overlay = document.querySelector(".mobile-overlay");

function syncMobileSidebarLayout() {
  if (!sidebar || !pageHeader) return;
  const isMobile = window.innerWidth <= 900;

  if (isMobile) {
    if (appContainer) appContainer.style.height = "";
    const headerRect = pageHeader.getBoundingClientRect();
    const headerHeight = Math.ceil(headerRect.height);
    const headerBottom = Math.ceil(headerRect.bottom);
    const topOffset = Math.max(headerHeight, headerBottom, 0);

    sidebar.style.top = `${topOffset}px`;
    sidebar.style.height = `calc(100dvh - ${topOffset}px)`;
    if (overlay) overlay.style.top = `${topOffset}px`;
    sidebar.dataset.mobileLayoutApplied = "1";
  } else {
    // Desktop: lock app shell to live header height so sidebar can't drift under header.
    const headerHeight = Math.ceil(pageHeader.getBoundingClientRect().height);
    if (appContainer) appContainer.style.height = `calc(100dvh - ${headerHeight}px)`;
    sidebar.style.top = "";
    sidebar.style.height = "100%";
    if (overlay) overlay.style.top = "";
    delete sidebar.dataset.mobileLayoutApplied;
  }
}

syncMobileSidebarLayout();
window.addEventListener("resize", syncMobileSidebarLayout);
window.addEventListener("scroll", syncMobileSidebarLayout, { passive: true });
if (window.visualViewport) {
  window.visualViewport.addEventListener("resize", syncMobileSidebarLayout);
  window.visualViewport.addEventListener("scroll", syncMobileSidebarLayout);
}
if (map && typeof map.on === "function") {
  map.on("click zoomend moveend", () => {
    requestAnimationFrame(syncMobileSidebarLayout);
  });
}

if (mobileMenuBtn && sidebar && overlay) {

  mobileMenuBtn.addEventListener("click", () => {
    syncMobileSidebarLayout();
    if (window.innerWidth <= 900 && appContainer) {
      appContainer.classList.remove("collapsed");
    }
    const open = sidebar.classList.toggle("open");

    mobileMenuBtn.textContent = open ? "✕" : "☰";
    overlay.classList.toggle("show", open);
    if (open) {
      sidebar.scrollTop = 0;
      requestAnimationFrame(() => {
        sidebar.scrollTop = 0;
      });
    }

    setTimeout(() => map.invalidateSize(), 200);
  });

  overlay.addEventListener("click", () => {
    sidebar.classList.remove("open");
    overlay.classList.remove("show");
    mobileMenuBtn.textContent = "☰";
  });
}
// ===== MOBILE SELECTION TOGGLE =====
const mobileSelBtn = document.getElementById("mobileSelectionBtn");


if (mobileSelBtn && selectionBox) {

  mobileSelBtn.addEventListener("click", () => {
    selectionBox.classList.toggle("show");
  });

  // keep count synced
  const originalUpdate = updateSelectionCount;
  updateSelectionCount = function () {
    originalUpdate();
    mobileSelBtn.textContent =
      "Selected: " + document.getElementById("selectionCount").textContent;
  };
}


  // ===== RESIZABLE BOTTOM SUMMARY PANEL =====
  const panel = document.getElementById("bottomSummary");
  const header = document.querySelector(".bottom-summary-header");
  const toggleBtn = document.getElementById("summaryToggleBtn");

  if (panel && header) {
    let isDragging = false;
    let startY = 0;
    let startHeight = 0;

    // Restore saved height
    const savedHeight = storageGet("summaryHeight");
    if (savedHeight) panel.style.height = savedHeight + "px";

    // Start collapsed on initial load (desktop + mobile)
    panel.classList.add("collapsed");
    panel.style.height = "40px";
    if (toggleBtn) toggleBtn.textContent = "▲";

    // Drag resize (mouse + touch/pen) while preserving button clicks in header
    header.addEventListener("pointerdown", e => {
      if (e.target.closest("button")) return;
      isDragging = true;
      startY = e.clientY;
      startHeight = panel.offsetHeight;
      document.body.style.userSelect = "none";
    });

    document.addEventListener("pointermove", e => {
      if (!isDragging) return;

      const delta = startY - e.clientY;
      let newHeight = startHeight + delta;

      const minHeight = 40;
      const maxHeight = window.innerHeight - 100;

      newHeight = Math.max(minHeight, Math.min(maxHeight, newHeight));
      panel.style.height = newHeight + "px";
    });

    document.addEventListener("pointerup", () => {
      if (!isDragging) return;
      isDragging = false;
      document.body.style.userSelect = "";
      storageSet("summaryHeight", panel.offsetHeight);

      // Hide resize hint after first drag
      const hint = document.querySelector(".resize-hint");
      if (hint) hint.style.display = "none";
    });


    // Collapse toggle
if (toggleBtn) {
  toggleBtn.onclick = () => {
    const isCollapsed = panel.classList.toggle("collapsed");

   if (isCollapsed) {
  storageSet("summaryHeight", panel.offsetHeight);
  panel.style.height = "40px";
  toggleBtn.textContent = "▲";
} else {
  let restored = storageGet("summaryHeight");

  if (!restored || restored <= 60) {
    restored = window.innerWidth <= 900 ? 300 : 250;
  }

  panel.style.height = restored + "px";
  toggleBtn.textContent = "▼";
}

  };
}

  }

  // ===== POP-OUT SUMMARY WINDOW =====
  const popoutBtn = document.getElementById("popoutSummaryBtn");
  const isStandaloneApp =
    window.matchMedia("(display-mode: standalone)").matches ||
    window.navigator.standalone === true;
  const useSameWindowForSummary = isStandaloneApp && window.innerWidth <= 900;

  if (popoutBtn) {
    popoutBtn.onclick = () => {
      const tableHTML = document.getElementById("routeSummaryTable")?.innerHTML;

      if (!tableHTML || tableHTML.includes("No summary")) {
        alert("No route summary loaded.");
        return;
      }

      const mapUrl = window.location.href;
      const win = useSameWindowForSummary
        ? window
        : window.open("", "_blank", "width=900,height=600,resizable=yes,scrollbars=yes");
      if (!win) {
        alert("Unable to open summary window.");
        return;
      }

      win.document.write(`
        <html>
          <head>
            <title>Route Summary</title>
            <style>
              :root {
                --bg: #f3f7fb;
                --panel: #ffffff;
                --line: #d6e2ee;
                --head: #e9f2fb;
                --text: #16324d;
                --muted: #4e6a86;
                --accent: #2f89df;
              }
              * { box-sizing: border-box; }
              body {
                margin: 0;
                padding: 18px;
                background: radial-gradient(circle at 10% 0%, #eaf2fb 0%, var(--bg) 42%);
                color: var(--text);
                font-family: "Segoe UI", Roboto, Arial, sans-serif;
              }
              .summary-shell {
                max-width: 1300px;
                margin: 0 auto;
                background: var(--panel);
                border: 1px solid var(--line);
                border-radius: 14px;
                box-shadow: 0 14px 30px rgba(16, 42, 68, 0.12);
                overflow: hidden;
              }
              .summary-header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                padding: 12px 14px;
                border-bottom: 1px solid var(--line);
                background: linear-gradient(180deg, #f9fcff 0%, #eef5fc 100%);
                gap: 10px;
              }
              .summary-title {
                margin: 0;
                font-size: 18px;
                font-weight: 700;
                letter-spacing: 0.01em;
              }
              .summary-note {
                font-size: 12px;
                color: var(--muted);
              }
              .summary-back-btn {
                border: 1px solid #b8cbdd;
                background: #ffffff;
                color: var(--text);
                border-radius: 10px;
                padding: 8px 10px;
                font-size: 12px;
                font-weight: 600;
                cursor: pointer;
              }
              .summary-back-btn:hover { background: #f2f8ff; }
              .summary-table-wrap {
                max-height: calc(100vh - 140px);
                overflow: auto;
              }
              table {
                border-collapse: separate;
                border-spacing: 0;
                width: 100%;
                font-size: 13px;
              }
              th, td {
                border-right: 1px solid var(--line);
                border-bottom: 1px solid var(--line);
                padding: 8px 10px;
                text-align: left;
                white-space: nowrap;
              }
              th {
                position: sticky;
                top: 0;
                z-index: 1;
                background: var(--head);
                color: var(--text);
                font-size: 12px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 0.03em;
              }
              tr:nth-child(even) td {
                background: #f8fbff;
              }
              tr:hover td {
                background: #edf5ff;
              }
              th:first-child, td:first-child { border-left: 1px solid var(--line); }
            </style>
          </head>
          <body>
            <div class="summary-shell">
              <div class="summary-header">
                <button class="summary-back-btn" onclick="returnToMap()">← Back to Map</button>
                <h2 class="summary-title">Route Summary</h2>
                <span class="summary-note">Scroll to view all columns and rows</span>
              </div>
              <div class="summary-table-wrap">
                ${tableHTML}
              </div>
            </div>
            <script>
              function returnToMap() {
                try {
                  if (window.opener && !window.opener.closed) {
                    window.opener.focus();
                    window.close();
                    return;
                  }
                } catch (e) {}
                window.location.href = ${JSON.stringify(mapUrl)};
              }
            </script>
          </body>
        </html>
      `);

      win.document.close();
    };
  }

  // ===== SUMMARY VISUALIZATION WINDOW =====
  const visualizeBtn = document.getElementById("visualizeSummaryBtn");

  if (visualizeBtn) {
    visualizeBtn.onclick = () => {
      const rows = window._summaryRows || [];
      const headers = window._summaryHeaders || [];

      if (!rows.length || !headers.length) {
        alert("No route summary loaded.");
        return;
      }

      const toNumber = v => {
        if (v === null || v === undefined) return NaN;
        const s = String(v).replace(/,/g, "").trim();
        if (!s) return NaN;
        const n = Number(s);
        return Number.isFinite(n) ? n : NaN;
      };
      const toHours = v => {
        if (v === null || v === undefined) return NaN;
        if (typeof v === "number" && Number.isFinite(v)) {
          // Excel time values may be stored as fractions of a day.
          if (v > 0 && v < 1) return v * 24;
          return v;
        }

        const s = String(v).trim();
        if (!s) return NaN;

        const direct = toNumber(s);
        if (Number.isFinite(direct)) {
          if (direct > 0 && direct < 1) return direct * 24;
          return direct;
        }

        // HH:MM or HH:MM:SS
        const colonMatch = s.match(/^(-?\d+):(\d{1,2})(?::(\d{1,2}))?$/);
        if (colonMatch) {
          const h = Number(colonMatch[1]) || 0;
          const m = Number(colonMatch[2]) || 0;
          const sec = Number(colonMatch[3] || 0) || 0;
          return h + m / 60 + sec / 3600;
        }

        // "11h 30m", "11 hr", "45 min"
        const lower = s.toLowerCase();
        const hMatch = lower.match(/(-?\d+(?:\.\d+)?)\s*(h|hr|hrs|hour|hours)/);
        const mMatch = lower.match(/(-?\d+(?:\.\d+)?)\s*(m|min|mins|minute|minutes)/);
        if (hMatch || mMatch) {
          const h = hMatch ? Number(hMatch[1]) : 0;
          const m = mMatch ? Number(mMatch[1]) : 0;
          return (Number.isFinite(h) ? h : 0) + (Number.isFinite(m) ? m : 0) / 60;
        }

        return NaN;
      };

      const escapeHtml = value =>
        String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/\"/g, "&quot;")
          .replace(/'/g, "&#39;");
      const normalize = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
      const normalizedHeaders = headers.map(h => ({ original: h, norm: normalize(h) }));

      function findHeader(candidates) {
        const direct = normalizedHeaders.find(h => candidates.includes(h.norm));
        if (direct) return direct.original;

        const loose = normalizedHeaders.find(h =>
          candidates.some(c => h.norm.includes(c) || c.includes(h.norm))
        );
        return loose ? loose.original : null;
      }

      const fieldSpec = [
        { key: "route", label: "Route", aliases: ["route", "newroute", "routeid", "rte"] },
        { key: "day", label: "Day", aliases: ["day", "routeday", "dispatchday", "newday"] },
        { key: "start", label: "Start Location", aliases: ["startlocation", "start", "origin", "startdepot", "startfacility"] },
        { key: "end", label: "End Location", aliases: ["endlocation", "end", "destination", "enddepot", "endfacility"] },
        { key: "stops", label: "Total Stops", aliases: ["totalstops", "stops", "stopcount", "numberofstops"] },
        { key: "breakTime", label: "Break Time", aliases: ["breaktime", "totbreaktime", "totalbreaktime", "break"] },
        { key: "facilityTime", label: "Facility Time", aliases: ["facilitytime", "totfacilitytime", "totalfacilitytime"] },
        { key: "totalTime", label: "Total Time", aliases: ["totaltime", "totalroutetime", "routetime", "hours"] },
        { key: "miles", label: "Miles", aliases: ["miles", "totalmiles", "route miles", "distance", "distancemiles"] },
        { key: "demand", label: "Demand", aliases: ["demand", "totaldemand", "volume", "load"] },
        { key: "trips", label: "Number of Trips", aliases: ["numberoftrips", "trips", "tripcount", "totaltrips"] }
      ];

      const selectedHeaders = {};
      fieldSpec.forEach(f => {
        selectedHeaders[f.key] = findHeader(f.aliases.map(normalize));
      });

      const focusedRows = rows.map(r => {
        const out = {};
        fieldSpec.forEach(f => {
          const h = selectedHeaders[f.key];
          out[f.key] = h ? r[h] : "";
        });
        return out;
      });

      // Aggregate at route+day level so the visualization centers on route/day units.
      const routeDayMap = new Map();
      focusedRows.forEach(r => {
        const route = String(r.route || "").trim() || "Unknown Route";
        const day = String(r.day || "").trim() || "Unknown Day";
        const key = `${route} | ${day}`;
        if (!routeDayMap.has(key)) {
          routeDayMap.set(key, {
            route,
            day,
            routeDay: key,
            start: String(r.start || "").trim(),
            end: String(r.end || "").trim(),
            stops: 0,
            breakTime: 0,
            facilityTime: 0,
            totalTime: 0,
            miles: 0,
            demand: 0,
            trips: 0
          });
        }
        const bucket = routeDayMap.get(key);

        const stops = toNumber(r.stops);
        const breakTime = toHours(r.breakTime);
        const facilityTime = toHours(r.facilityTime);
        const totalTime = toHours(r.totalTime);
        const miles = toNumber(r.miles);
        const demand = toNumber(r.demand);
        const trips = toNumber(r.trips);

        bucket.stops += Number.isFinite(stops) ? stops : 0;
        bucket.breakTime += Number.isFinite(breakTime) ? breakTime : 0;
        bucket.facilityTime += Number.isFinite(facilityTime) ? facilityTime : 0;
        bucket.totalTime += Number.isFinite(totalTime) ? totalTime : 0;
        bucket.miles += Number.isFinite(miles) ? miles : 0;
        bucket.demand += Number.isFinite(demand) ? demand : 0;
        bucket.trips += Number.isFinite(trips) ? trips : 0;

        if (!bucket.start && r.start) bucket.start = String(r.start);
        if (!bucket.end && r.end) bucket.end = String(r.end);
      });

      const routeDayRows = Array.from(routeDayMap.values());

      const dayTotalsMap = new Map();
      routeDayRows.forEach(r => {
        const day = r.day || "Unknown Day";
        if (!dayTotalsMap.has(day)) {
          dayTotalsMap.set(day, {
            day,
            routeDayCount: 0,
            demand: 0,
            miles: 0,
            stops: 0,
            trips: 0
          });
        }
        const d = dayTotalsMap.get(day);
        d.routeDayCount += 1;
        d.demand += r.demand;
        d.miles += r.miles;
        d.stops += r.stops;
        d.trips += r.trips;
      });

      const dayTotals = Array.from(dayTotalsMap.values()).sort((a, b) => b.routeDayCount - a.routeDayCount);

      function makeBars(items, labelKey, valueKey, alt, kind) {
        if (!items.length) return '<div class="empty">No data available.</div>';
        const max = Math.max(1, ...items.map(i => i[valueKey]));
        return items
          .slice()
          .sort((a, b) => b[valueKey] - a[valueKey])
          .map(i => `
            <div class="bar-row section-route-row" data-kind="${escapeHtml(kind || "")}" data-key="${encodeURIComponent(String(i[labelKey] ?? ""))}">
              <div class="bar-label">${escapeHtml(i[labelKey])}</div>
              <div class="bar-track ${alt ? "alt" : ""}"><div class="bar-fill ${alt ? "alt" : ""}" style="width:${(i[valueKey] / max) * 100}%"></div></div>
              <div class="bar-value">${i[valueKey].toLocaleString(undefined, { maximumFractionDigits: 2 })}</div>
            </div>
          `)
          .join("");
      }

      const dayDemandBars = makeBars(dayTotals, "day", "demand", false, "day");
      const dayMilesBars = makeBars(dayTotals, "day", "miles", true, "day");
      const dayStopsBars = makeBars(dayTotals, "day", "stops", false, "day");
      const dayTripsBars = makeBars(dayTotals, "day", "trips", true, "day");

      const totalDemand = routeDayRows.reduce((a, r) => a + r.demand, 0);
      const totalMiles = routeDayRows.reduce((a, r) => a + r.miles, 0);
      const totalStops = routeDayRows.reduce((a, r) => a + r.stops, 0);
      const topDemandDay = dayTotals.slice().sort((a, b) => b.demand - a.demand)[0]?.day || "N/A";
      const targetHours = 11;
      const totalTimeSum = routeDayRows.reduce((a, r) => a + r.totalTime, 0);
      const avgTotalTime = routeDayRows.length ? totalTimeSum / routeDayRows.length : 0;

      let overTargetCount = 0;
      let underOrAtTargetCount = 0;
      routeDayRows.forEach(r => {
        if (r.totalTime > targetHours) overTargetCount += 1;
        else underOrAtTargetCount += 1;
      });

      const timeTargetBars = makeBars(
        [
          { label: `Over ${targetHours} Hours`, count: overTargetCount },
          { label: `${targetHours} Hours or Less`, count: underOrAtTargetCount }
        ],
        "label",
        "count",
        true,
        "target"
      );

      const uniqueRoutes = new Set(routeDayRows.map(r => String(r.route || "").trim()).filter(Boolean)).size;
      const scatterDataJson = JSON.stringify(
        routeDayRows.map(r => ({
          routeDay: r.routeDay,
          route: r.route,
          day: r.day,
          stops: Number(r.stops) || 0,
          miles: Number(r.miles) || 0,
          totalTime: Number(r.totalTime) || 0
        }))
      ).replace(/</g, "\\u003c");

      const missingFields = fieldSpec
        .filter(f => !selectedHeaders[f.key])
        .map(f => f.label);

      const routeDayTable = routeDayRows
        .slice()
        .sort((a, b) => (a.day === b.day ? String(a.route).localeCompare(String(b.route)) : String(a.day).localeCompare(String(b.day))))
        .slice(0, 600)
        .map(r => `
          <tr>
            <td>${escapeHtml(r.route)}</td>
            <td>${escapeHtml(r.day)}</td>
            <td>${escapeHtml(r.start)}</td>
            <td>${escapeHtml(r.end)}</td>
            <td>${r.stops.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
            <td>${r.breakTime.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
            <td>${r.facilityTime.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
            <td>${r.totalTime.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
            <td>${r.miles.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
            <td>${r.demand.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
            <td>${r.trips.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
          </tr>
        `)
        .join("");

      const mapUrl = window.location.href;
      const win = useSameWindowForSummary
        ? window
        : window.open("", "_blank", "width=1080,height=760,resizable=yes,scrollbars=yes");
      if (!win) return;

      win.document.write(`
        <html>
          <head>
            <title>Route Summary Visualization</title>
            <style>
              :root { --bg:#f3f7fb; --panel:#ffffff; --line:#d8e3ee; --text:#14314d; --muted:#4f6883; --a:#2f89df; --b:#20a38e; }
              * { box-sizing: border-box; }
              body { margin:0; padding:18px; font-family:"Segoe UI",Roboto,Arial,sans-serif; background:radial-gradient(circle at 10% 0%, #eaf2fb 0%, var(--bg) 45%); color:var(--text); }
              .shell { max-width:1200px; margin:0 auto; background:var(--panel); border:1px solid var(--line); border-radius:14px; box-shadow:0 14px 30px rgba(16,42,68,.12); overflow:hidden; }
              .head { display:flex; justify-content:space-between; align-items:center; padding:12px 14px; border-bottom:1px solid var(--line); background:linear-gradient(180deg,#f9fcff 0%,#eef5fc 100%); }
              .head-left { display:flex; align-items:center; gap:10px; }
              .title { margin:0; font-size:18px; font-weight:700; }
              .meta { font-size:12px; color:var(--muted); }
              .summary-back-btn { border:1px solid #b8cbdd; background:#fff; color:var(--text); border-radius:10px; padding:8px 10px; font-size:12px; font-weight:600; cursor:pointer; }
              .summary-back-btn:hover { background:#f2f8ff; }
              .grid { display:grid; grid-template-columns: repeat(4, minmax(160px, 1fr)); gap:10px; padding:12px; border-bottom:1px solid var(--line); }
              .card { border:1px solid var(--line); border-radius:10px; padding:10px; background:#fbfdff; }
              .card-label { font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.04em; }
              .card-value { margin-top:4px; font-size:22px; font-weight:700; }
              .section { padding:12px; border-bottom:1px solid var(--line); }
              .section h3 { margin:0 0 10px 0; font-size:14px; }
              .bar-row { display:grid; grid-template-columns: 240px 1fr 120px; gap:10px; align-items:center; margin-bottom:8px; }
              .section-route-row { cursor:pointer; border-radius:8px; padding:4px 6px; margin-left:-6px; margin-right:-6px; }
              .section-route-row:hover { background:#f1f7ff; }
              .bar-label { font-size:12px; color:var(--text); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
              .bar-track { height:12px; border-radius:999px; background:#e5eef8; overflow:hidden; }
              .bar-fill { height:100%; background:linear-gradient(90deg,#47a2f8 0%,var(--a) 100%); }
              .bar-track.alt { background:#e8f5f2; }
              .bar-fill.alt { background:linear-gradient(90deg,#36c3ab 0%,var(--b) 100%); }
              .bar-value { text-align:right; font-size:12px; color:var(--muted); font-weight:700; }
              .empty { color:var(--muted); font-size:13px; }
              .note { margin: 0 0 10px 0; color: var(--muted); font-size: 12px; }
              .legend { display:flex; gap:14px; align-items:center; margin: 0 0 10px 0; color:var(--muted); font-size:12px; flex-wrap:wrap; }
              .dot { width:10px; height:10px; border-radius:50%; display:inline-block; margin-right:6px; vertical-align:middle; }
              .dot-blue { background:#2f89df; }
              .dot-red { background:#e25b53; }
              .chart-wrap { border:1px solid var(--line); border-radius:10px; background:#fff; padding:8px; overflow:auto; }
              .chart-controls { display:flex; gap:10px; flex-wrap:wrap; margin:0 0 10px 0; }
              .chart-controls label { font-size:12px; color:var(--muted); display:flex; gap:6px; align-items:center; }
              .chart-controls select, .chart-controls input { border:1px solid #c8d8e8; border-radius:8px; padding:6px 8px; font-size:12px; }
              .scatter-tooltip { position:fixed; pointer-events:none; background:#102a44; color:#fff; padding:8px 10px; border-radius:8px; font-size:12px; z-index:99999; box-shadow:0 10px 24px rgba(0,0,0,.22); display:none; white-space:nowrap; }
              .chart-toast { position:fixed; right:18px; bottom:18px; z-index:99999; background:#1f7a3f; color:#fff; border-radius:10px; padding:9px 12px; font-size:12px; box-shadow:0 10px 24px rgba(0,0,0,.2); opacity:0; transform:translateY(8px); transition:opacity .2s ease, transform .2s ease; pointer-events:none; }
              .chart-toast.show { opacity:1; transform:translateY(0); }
              .chart-toast.error { background:#a83f3a; }
              .section-modal { position:fixed; inset:0; background:rgba(16,42,68,.45); display:flex; align-items:center; justify-content:center; padding:18px; z-index:99998; }
              .section-modal.hidden { display:none; }
              .section-modal-card { width:min(760px, 100%); max-height:80vh; background:#fff; border:1px solid var(--line); border-radius:12px; overflow:hidden; box-shadow:0 18px 36px rgba(16,42,68,.25); }
              .section-modal-head { display:flex; align-items:center; justify-content:space-between; padding:10px 12px; border-bottom:1px solid var(--line); background:#f7fbff; }
              .section-modal-title { margin:0; font-size:14px; }
              .section-modal-close { border:1px solid #b7cce1; background:#fff; border-radius:8px; padding:5px 10px; cursor:pointer; }
              .section-modal-body { padding:10px 12px; max-height:calc(80vh - 52px); overflow:auto; }
              .section-modal-list { margin:0; padding-left:18px; }
              .section-modal-list li { margin:6px 0; font-size:13px; }
              .table-wrap { overflow:auto; border:1px solid var(--line); border-radius:10px; background:#fff; }
              table { width:100%; border-collapse:collapse; min-width:1160px; }
              th, td { border-bottom:1px solid var(--line); border-right:1px solid var(--line); padding:8px 10px; text-align:left; font-size:12px; white-space:nowrap; }
              th { position:sticky; top:0; background:#f4f9ff; z-index:2; text-transform:uppercase; letter-spacing:.03em; font-size:11px; }
              tr:nth-child(even) td { background:#fbfdff; }
              tr:hover td { background:#eef6ff; }
              @media (max-width: 980px) {
                .grid { grid-template-columns: repeat(2, minmax(140px, 1fr)); }
                .bar-row { grid-template-columns: 170px 1fr 90px; }
              }
            </style>
          </head>
          <body>
            <div class="shell">
              <div class="head">
                <div class="head-left">
                  <button class="summary-back-btn" onclick="returnToMap()">← Back to Map</button>
                  <h2 class="title">Route Summary Visualization</h2>
                </div>
                <span class="meta">Rows: ${rows.length.toLocaleString()} | Columns: ${headers.length.toLocaleString()}</span>
              </div>
              <div class="grid">
                <div class="card"><div class="card-label">Total Routes</div><div class="card-value">${uniqueRoutes.toLocaleString()}</div></div>
                <div class="card"><div class="card-label">Avg Total Time</div><div class="card-value">${avgTotalTime.toLocaleString(undefined, { maximumFractionDigits: 2 })}h</div></div>
                <div class="card"><div class="card-label">Over 11h</div><div class="card-value">${overTargetCount.toLocaleString()}</div></div>
                <div class="card"><div class="card-label">11h or Less</div><div class="card-value">${underOrAtTargetCount.toLocaleString()}</div></div>
              </div>
              <div class="section">
                <h3>Stops vs Miles by Route + Day</h3>
                <p class="note">Use filters below and hover any point to see details. Red points are over the 11-hour total-time target.</p>
                <div class="chart-controls">
                  <label>Day
                    <select id="scatterDayFilter">
                      <option value="all">All Days</option>
                    </select>
                  </label>
                  <label><input type="checkbox" id="scatterOverOnly" /> Show only over 11h</label>
                </div>
                <div class="legend">
                  <span><span class="dot dot-blue"></span>At or under ${targetHours}h total time</span>
                  <span><span class="dot dot-red"></span>Over ${targetHours}h total time</span>
                </div>
                <div id="stopsMilesScatterHost" class="chart-wrap"></div>
                <div id="scatterTooltip" class="scatter-tooltip"></div>
                <div id="chartToast" class="chart-toast"></div>
              </div>
              <div class="section">
                <h3>Total Time vs ${targetHours}-Hour Target</h3>
                <p class="note">Average total time per route+day: ${avgTotalTime.toLocaleString(undefined, { maximumFractionDigits: 2 })} hours</p>
                ${timeTargetBars}
              </div>
              <div class="section">
                <h3>Distribution of Demand Per Day</h3>
                ${dayDemandBars}
              </div>
              <div class="section">
                <h3>Distribution of Miles Per Day</h3>
                ${dayMilesBars}
              </div>
              <div class="section">
                <h3>Distribution of Stops Per Day</h3>
                ${dayStopsBars}
              </div>
              <div class="section">
                <h3>Distribution of Trips Per Day</h3>
                ${dayTripsBars}
              </div>
              <div class="section">
                <h3>Route + Day Detail</h3>
                <p class="note">Each row below is grouped by route and day to make day-level performance easier to interpret.</p>
                ${missingFields.length ? `<p class="note">Missing in this file: ${escapeHtml(missingFields.join(", "))}</p>` : ""}
                <p class="note">Totals: Stops ${totalStops.toLocaleString(undefined, { maximumFractionDigits: 2 })}, Demand ${totalDemand.toLocaleString(undefined, { maximumFractionDigits: 2 })}, Miles ${totalMiles.toLocaleString(undefined, { maximumFractionDigits: 2 })}</p>
                <div class="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Route</th>
                        <th>Day</th>
                        <th>Start Location</th>
                        <th>End Location</th>
                        <th>Total Stops</th>
                        <th>Break Time</th>
                        <th>Facility Time</th>
                        <th>Total Time</th>
                        <th>Miles</th>
                        <th>Demand</th>
                        <th>Number of Trips</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${routeDayTable || '<tr><td colspan="11">No route rows found.</td></tr>'}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
            <div id="sectionRouteDayModal" class="section-modal hidden">
              <div class="section-modal-card">
                <div class="section-modal-head">
                  <h4 id="sectionRouteDayTitle" class="section-modal-title">Route+Day List</h4>
                  <button id="sectionRouteDayClose" class="section-modal-close">Close</button>
                </div>
                <div class="section-modal-body">
                  <ul id="sectionRouteDayList" class="section-modal-list"></ul>
                </div>
              </div>
            </div>
            <script>
              function returnToMap() {
                try {
                  if (window.opener && !window.opener.closed) {
                    window.opener.focus();
                    window.close();
                    return;
                  }
                } catch (e) {}
                window.location.href = ${JSON.stringify(mapUrl)};
              }

              (() => {
                const data = ${scatterDataJson};
                const host = document.getElementById("stopsMilesScatterHost");
                const tooltip = document.getElementById("scatterTooltip");
                const chartToast = document.getElementById("chartToast");
                const dayFilter = document.getElementById("scatterDayFilter");
                const overOnly = document.getElementById("scatterOverOnly");
                const targetHours = ${targetHours};
                const sectionRows = Array.from(document.querySelectorAll(".section-route-row"));
                const modal = document.getElementById("sectionRouteDayModal");
                const modalTitle = document.getElementById("sectionRouteDayTitle");
                const modalList = document.getElementById("sectionRouteDayList");
                const modalClose = document.getElementById("sectionRouteDayClose");

                function escapeHtmlLocal(value) {
                  return String(value ?? "")
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;")
                    .replace(/\"/g, "&quot;")
                    .replace(/'/g, "&#39;");
                }

                function openRouteDayModal(title, rows) {
                  modalTitle.textContent = title;
                  if (!rows.length) {
                    modalList.innerHTML = "<li>No route+day entries found.</li>";
                  } else {
                    const sorted = rows.slice().sort((a, b) => String(a.routeDay || "").localeCompare(String(b.routeDay || "")));
                    modalList.innerHTML = sorted.map(r =>
                      "<li><strong>" + escapeHtmlLocal(r.routeDay) + "</strong> - Stops: " + Number(r.stops || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) +
                      ", Miles: " + Number(r.miles || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) +
                      ", Time: " + Number(r.totalTime || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) + "h</li>"
                    ).join("");
                  }
                  modal.classList.remove("hidden");
                }

                function closeRouteDayModal() {
                  modal.classList.add("hidden");
                }

                function showChartToast(message, isError) {
                  if (!chartToast) return;
                  chartToast.textContent = message;
                  chartToast.classList.toggle("error", !!isError);
                  chartToast.classList.add("show");
                  setTimeout(() => {
                    chartToast.classList.remove("show");
                  }, 1900);
                }

                modalClose.addEventListener("click", closeRouteDayModal);
                modal.addEventListener("click", e => {
                  if (e.target === modal) closeRouteDayModal();
                });

                sectionRows.forEach(row => {
                  row.addEventListener("click", () => {
                    const kind = row.getAttribute("data-kind") || "";
                    const rawKey = row.getAttribute("data-key") || "";
                    const key = decodeURIComponent(rawKey);

                    if (kind === "day") {
                      const matched = data.filter(d => String(d.day || "") === key);
                      openRouteDayModal("Route+Day entries for " + key, matched);
                      return;
                    }

                    if (kind === "target") {
                      const isOverBucket = key.toLowerCase().startsWith("over ");
                      const matched = data.filter(d => isOverBucket ? Number(d.totalTime) > targetHours : Number(d.totalTime) <= targetHours);
                      openRouteDayModal("Route+Day entries: " + key, matched);
                    }
                  });
                });

                const days = Array.from(new Set(data.map(d => String(d.day || "").trim()).filter(Boolean))).sort();
                days.forEach(day => {
                  const opt = document.createElement("option");
                  opt.value = day;
                  opt.textContent = day;
                  dayFilter.appendChild(opt);
                });

                function render() {
                  const selectedDay = dayFilter.value;
                  const filtered = data.filter(d => {
                    if (selectedDay !== "all" && String(d.day) !== selectedDay) return false;
                    if (overOnly.checked && !(d.totalTime > targetHours)) return false;
                    return true;
                  });

                  if (!filtered.length) {
                    host.innerHTML = '<div class="empty">No points match the selected filters.</div>';
                    return;
                  }

                  const width = 980;
                  const height = 360;
                  const padL = 56;
                  const padR = 24;
                  const padT = 20;
                  const padB = 46;
                  const plotW = width - padL - padR;
                  const plotH = height - padT - padB;
                  const maxX = Math.max(1, ...filtered.map(p => p.stops));
                  const maxY = Math.max(1, ...filtered.map(p => p.miles));
                  const x = v => padL + (v / maxX) * plotW;
                  const y = v => padT + plotH - (v / maxY) * plotH;

                  const xTicks = Array.from({ length: 6 }, (_, i) => {
                    const val = (maxX * i) / 5;
                    const px = x(val);
                    return '<line x1="' + px + '" y1="' + padT + '" x2="' + px + '" y2="' + (padT + plotH) + '" stroke="#e6edf6" stroke-width="1" />' +
                      '<text x="' + px + '" y="' + (padT + plotH + 18) + '" text-anchor="middle" fill="#627d97" font-size="11">' + val.toFixed(0) + '</text>';
                  }).join("");

                  const yTicks = Array.from({ length: 6 }, (_, i) => {
                    const val = (maxY * i) / 5;
                    const py = y(val);
                    return '<line x1="' + padL + '" y1="' + py + '" x2="' + (padL + plotW) + '" y2="' + py + '" stroke="#e6edf6" stroke-width="1" />' +
                      '<text x="' + (padL - 8) + '" y="' + (py + 4) + '" text-anchor="end" fill="#627d97" font-size="11">' + val.toFixed(0) + '</text>';
                  }).join("");

                  const circles = filtered.map((p, idx) => {
                    const color = p.totalTime > targetHours ? "#e25b53" : "#2f89df";
                    const info = (String(p.routeDay || "") + " | Stops: " + p.stops.toFixed(2) + " | Miles: " + p.miles.toFixed(2) + " | Total Time: " + p.totalTime.toFixed(2) + "h")
                      .replace(/&/g, "&amp;")
                      .replace(/</g, "&lt;")
                      .replace(/>/g, "&gt;");
                    const routeSafe = String(p.route || "").replace(/"/g, "&quot;");
                    const daySafe = String(p.day || "").replace(/"/g, "&quot;");
                    return '<circle class="scatter-point" data-route="' + routeSafe + '" data-day="' + daySafe + '" data-info="' + info + '" cx="' + x(p.stops) + '" cy="' + y(p.miles) + '" r="5" fill="' + color + '" fill-opacity="0.85" stroke="#ffffff" stroke-width="1.2" />';
                  }).join("");

                  host.innerHTML =
                    '<svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Stops vs miles scatter plot by route and day">' +
                    xTicks + yTicks +
                    '<line x1="' + padL + '" y1="' + (padT + plotH) + '" x2="' + (padL + plotW) + '" y2="' + (padT + plotH) + '" stroke="#8ea4bb" stroke-width="1.2" />' +
                    '<line x1="' + padL + '" y1="' + padT + '" x2="' + padL + '" y2="' + (padT + plotH) + '" stroke="#8ea4bb" stroke-width="1.2" />' +
                    circles +
                    '<text x="' + (padL + plotW / 2) + '" y="' + (height - 8) + '" text-anchor="middle" fill="#3e5f7d" font-size="12">Total Stops</text>' +
                    '<text x="16" y="' + (padT + plotH / 2) + '" text-anchor="middle" fill="#3e5f7d" font-size="12" transform="rotate(-90 16 ' + (padT + plotH / 2) + ')">Miles</text>' +
                    '</svg>';

                  host.querySelectorAll(".scatter-point").forEach(el => {
                    el.addEventListener("mousemove", e => {
                      tooltip.style.display = "block";
                      tooltip.style.left = (e.clientX + 12) + "px";
                      tooltip.style.top = (e.clientY + 12) + "px";
                      tooltip.innerHTML = el.getAttribute("data-info") || "";
                    });
                    el.addEventListener("mouseleave", () => {
                      tooltip.style.display = "none";
                    });
                    el.addEventListener("click", () => {
                      const route = el.getAttribute("data-route") || "";
                      const day = el.getAttribute("data-day") || "";

                      if (!window.opener || typeof window.opener.highlightRouteDayOnMap !== "function") {
                        showChartToast("Could not connect to map window.", true);
                        return;
                      }

                      let result = null;
                      try {
                        result = window.opener.highlightRouteDayOnMap(route, day);
                      } catch (err) {
                        showChartToast("Failed to highlight on map.", true);
                        return;
                      }

                      if (result && result.ok) {
                        showChartToast("Route+Day highlighted on map.", false);
                      } else {
                        showChartToast((result && result.message) ? result.message : "No matching Route+Day found.", true);
                      }
                    });
                  });
                }

                dayFilter.addEventListener("change", render);
                overOnly.addEventListener("change", render);
                render();
              })();
            </script>
          </body>
        </html>
      `);
      win.document.close();
    };
  }
//======
// ===== RESET MAP BUTTON (TRUE HARD RESET FOR THIS APP) =====
const resetBtn = document.getElementById("resetMapBtn");

if (resetBtn) {
  resetBtn.addEventListener("click", () => {

    // 1. Reset map view
    map.setView([39.5, -98.35], 4);

    // 2. Clear drawn polygon
    drawnLayer.clearLayers();

    // 3. Remove ALL markers from map
    Object.values(routeDayGroups).forEach(group => {
      group.layers.forEach(marker => map.removeLayer(marker));
    });

    // 4. Clear stored marker groups & symbols
    Object.keys(routeDayGroups).forEach(k => delete routeDayGroups[k]);
    Object.keys(symbolMap).forEach(k => delete symbolMap[k]);

    // 5. Reset counters & stats
    document.getElementById("selectionCount").textContent = "0";

    // 6. Clear route/day checkbox UI
    document.getElementById("routeCheckboxes").innerHTML = "";
    buildDayCheckboxes();

    // 7. Reset bounds tracker
    globalBounds = L.latLngBounds();

    // 8. Clear bottom summary
    const summary = document.getElementById("routeSummaryTable");
    if (summary) summary.innerHTML = "No summary loaded";

    // 9. Collapse summary panel
    const panel = document.getElementById("bottomSummary");
    const btn = document.getElementById("summaryToggleBtn");
    if (panel && btn) {
      panel.classList.add("collapsed");
      panel.style.height = "40px";
      btn.textContent = "▲";
    }
  });
}


// ===== LIVE GPS BUTTON =====
const locateBtn = document.getElementById("locateMeBtn");

if (locateBtn) {
  let tracking = false;

  locateBtn.addEventListener("click", () => {
    if (!tracking) {
      startLiveTracking();
      locateBtn.textContent = "■";
      locateBtn.classList.add("tracking");   // 🔴 turns button red
      tracking = true;
    } else {
      if (watchId !== null) {
        navigator.geolocation.clearWatch(watchId);
        watchId = null;
      }
      if (headingMarker) {
  map.removeLayer(headingMarker);
  headingMarker = null;
}

window.removeEventListener("deviceorientation", updateHeading);


      locateBtn.textContent = "📍";
      locateBtn.classList.remove("tracking"); // 🔵 back to blue
      tracking = false;
    }
  });
}

// ================= FILE MANAGER MODAL =================
const fileManagerModal = document.getElementById("fileManagerModal");
const openFileManagerBtn = document.getElementById("openFileManagerBtn");
const closeFileManagerBtn = document.getElementById("closeFileManager");
const savedFilesSortSelect = document.getElementById("savedFilesSortSelect");
const savedFilesGuide = document.getElementById("savedFilesGuide");

function markSavedFilesGuideSeen() {
  storageSet("savedFilesGuideSeen", "1");
}

if (openFileManagerBtn) {
  if (savedFilesGuide) savedFilesGuide.classList.remove("hidden");

  openFileManagerBtn.classList.add("attention");

  openFileManagerBtn.addEventListener("click", () => {
    fileManagerModal.style.display = "flex";
    markSavedFilesGuideSeen();
  });
}

if (closeFileManagerBtn) {
  closeFileManagerBtn.addEventListener("click", () => {
    fileManagerModal.style.display = "none";
  });
}

if (savedFilesSortSelect) {
  savedFilesSortSelect.value = normalizeSavedFilesSortMode(storageGet(SAVED_FILES_SORT_MODE_KEY));
  savedFilesSortSelect.addEventListener("change", () => {
    storageSet(SAVED_FILES_SORT_MODE_KEY, normalizeSavedFilesSortMode(savedFilesSortSelect.value));
    listFiles();
  });
}

window.addEventListener("click", (e) => {
  if (e.target === fileManagerModal) {
    fileManagerModal.style.display = "none";
  }
});
  

// ===== DOWNLOAD FULL EXCEL (WITH CONFIRM MODAL) =====
const downloadBtn = document.getElementById("downloadFullExcelBtn");
const modal = document.getElementById("downloadConfirmModal");
const confirmBtn = document.getElementById("confirmDownload");
const cancelBtn = document.getElementById("cancelDownload");

if (downloadBtn && modal && confirmBtn && cancelBtn) {

  // Open confirmation modal
  downloadBtn.addEventListener("click", () => {

    if (!window._currentWorkbook) {
      alert("No Excel file loaded.");
      return;
    }

    modal.style.display = "flex";
  });

  // Cancel download
  cancelBtn.addEventListener("click", () => {
    modal.style.display = "none";
  });

  // Confirm download
  confirmBtn.addEventListener("click", () => {

    modal.style.display = "none";

    if (!window._currentWorkbook) {
      alert("No Excel file loaded.");
      return;
    }

    const now = new Date();

    const timestamp =
      now.getFullYear() + "-" +
      String(now.getMonth() + 1).padStart(2, "0") + "-" +
      String(now.getDate()).padStart(2, "0") + "_" +
      String(now.getHours()).padStart(2, "0") + "-" +
      String(now.getMinutes()).padStart(2, "0") + "-" +
      String(now.getSeconds()).padStart(2, "0");

    const baseName = getDownloadBaseName(window._currentFilePath);

    const newFileName = `${baseName}_Downloaded_${timestamp}.xlsx`;

    XLSX.writeFile(window._currentWorkbook, newFileName);
  });
}




// ===== AUTO-RESIZE MARKERS ON ZOOM =====
map.on("zoomend", () => {
  window._labelCount = 0;

  const newSize = getMarkerPixelSize();
  const currentZoom = map.getZoom();
  const maxZoom = map.getMaxZoom();

// ===== AUTO TOGGLE SATELLITE STREET NAMES =====
const currentBase = document.getElementById("baseMapSelect")?.value;

if (currentBase === "satellite") {
  if (map.getZoom() >= 15) {
    map.addLayer(satelliteLabelsLayer);
  } else {
    map.removeLayer(satelliteLabelsLayer);
  }
} else {
  map.removeLayer(satelliteLabelsLayer);
}
syncStreetNetworkOverlay();


  Object.values(routeDayGroups).forEach(group => {
    group.layers.forEach(layer => {
      const base = layer._base;
      if (!base) return;

      // ---- Resize markers (your existing logic) ----
      if (layer.setRadius) {
        layer.setRadius(newSize);
      } else {
        const { lat, lon, symbol } = base;

        const scale = 40075016.686 / Math.pow(2, currentZoom + 8);
        const dLat = newSize * scale / 111320;
        const dLng = dLat / Math.cos(lat * Math.PI / 180);

        if (symbol.shape === "square") {
          layer.setBounds([[lat - dLat, lon - dLng], [lat + dLat, lon + dLng]]);
        }

        if (symbol.shape === "triangle") {
          layer.setLatLngs([
            [lat + dLat, lon],
            [lat - dLat, lon - dLng],
            [lat - dLat, lon + dLng]
          ]);
        }

        if (symbol.shape === "diamond") {
          layer.setLatLngs([
            [lat + dLat, lon],
            [lat, lon + dLng],
            [lat - dLat, lon],
            [lat, lon - dLng]
          ]);
        }
      }

     // ---- STREET LABEL LOGIC (MOBILE SAFE) ----
// ---- STREET LABEL LOGIC (HARD TOGGLE CONTROL) ----
if (layer._hasStreetLabel) {

  // 🚫 If toggle is OFF, force close and skip
  if (!window.streetLabelsEnabled) {
    layer.closeTooltip();
    return;
  }

  const bounds = map.getBounds();
  const isVisible = bounds.contains(layer.getLatLng());

  const MAX_LABELS = 150;
  if (!window._labelCount) window._labelCount = 0;

  if (
    currentZoom >= maxZoom - 3 &&
    map.hasLayer(layer) &&
    isVisible &&
    window._labelCount < MAX_LABELS
  ) {
    layer.openTooltip();
    window._labelCount++;
  } else {
    layer.closeTooltip();
  }
}

    });
  });
});

  
// Position Locate button correctly for desktop/mobile
placeLocateButton();
window.addEventListener("resize", placeLocateButton);

  // Position Locate button correctly for desktop/mobile
placeLocateButton();
window.addEventListener("resize", placeLocateButton);


// ===== STREET LABEL TOGGLE =====
const streetToggle = document.getElementById("streetLabelToggle");

if (streetToggle) {

  // Set initial state based on checkbox
  window.streetLabelsEnabled = streetToggle.checked;

  // Force labels to respect initial state
  map.whenReady(() => {
    map.fire("zoomend");
  });

  streetToggle.addEventListener("change", (e) => {
    window.streetLabelsEnabled = e.target.checked;

    // Immediately refresh labels
    map.fire("zoomend");
  });
}
  
////////////////////////////////////////////////////////////////////
// 🔍 MAP ADDRESS SEARCH (PASTE RIGHT BELOW STREET TOGGLE)
////////////////////////////////////////////////////////////////////

const searchInput = document.getElementById("mapSearchInput");
const searchBtn   = document.getElementById("mapSearchBtn");

function searchMapByAddress() {

  if (!searchInput) return;

  const query = searchInput.value.trim().toLowerCase();
  if (!query) return;

  const resultsPanel = document.getElementById("searchResultsPanel");
  const resultsList  = document.getElementById("searchResultsList");

  resultsList.innerHTML = "";
  let matches = [];

  Object.values(routeDayGroups).forEach(group => {

    group.layers.forEach(marker => {

      const row = marker._rowRef;
      if (!row) return;

      const address = [
        row["CSADR#"] || "",
        row["CSSDIR"] || "",
        row["CSSTRT"] || "",
        row["CSSFUX"] || ""
      ].join(" ").toLowerCase();

      if (address.includes(query)) {
        matches.push({ marker, row });
      }

    });

  });

  if (!matches.length) {
    alert("No matching addresses found.");
    return;
  }

  matches.forEach((item, index) => {

    const div = document.createElement("div");
    div.className = "search-result-item";

    const displayAddress = [
      item.row["CSADR#"] || "",
      item.row["CSSDIR"] || "",
      item.row["CSSTRT"] || "",
      item.row["CSSFUX"] || ""
    ].join(" ");

    div.textContent = displayAddress;

    div.onclick = () => {

      // Remove previous selected styling
      document.querySelectorAll(".search-result-item")
        .forEach(el => el.classList.remove("selected"));

      div.classList.add("selected");

      map.setView(item.marker.getLatLng(), 18);

      item.marker.setStyle?.({
        color: "#ffff00",
        fillColor: "#ffff00",
        fillOpacity: 1
      });

    };

    resultsList.appendChild(div);

  });

  resultsPanel.classList.remove("hidden");
}
  // Hook up search button + Enter key
if (searchBtn) {
  searchBtn.addEventListener("click", searchMapByAddress);
}

if (searchInput) {
  searchInput.addEventListener("keydown", function(e) {
    if (e.key === "Enter") {
      searchMapByAddress();
    }
  });
}
  // ===== CLEAR SEARCH RESULTS =====
const clearSearchBtn = document.getElementById("clearSearchResults");

if (clearSearchBtn) {
  clearSearchBtn.addEventListener("click", () => {

    // Clear search input
    const searchInput = document.getElementById("mapSearchInput");
    if (searchInput) searchInput.value = "";

    // Clear results list
    const resultsList = document.getElementById("searchResultsList");
    if (resultsList) resultsList.innerHTML = "";

    // Hide results panel
    const resultsPanel = document.getElementById("searchResultsPanel");
    if (resultsPanel) {
      resultsPanel.classList.add("hidden");
    }

  });
}
////////////////central save function
async function saveWorkbookToCloud() {

  const newSheet = XLSX.utils.json_to_sheet(window._currentRows);
  window._currentWorkbook.Sheets[
    window._currentWorkbook.SheetNames[0]
  ] = newSheet;

  const bookType = window._currentFilePath.toLowerCase().endsWith(".xlsm")
    ? "xlsm"
    : "xlsx";

  const wbArray = XLSX.write(window._currentWorkbook, {
    bookType,
    type: "array"
  });

  const { error } = await sb.storage
    .from(BUCKET)
    .upload(window._currentFilePath, wbArray, {
      upsert: true,
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

  if (error) {
    console.error("Cloud Save Error:", error);
    return false;
  }

  return true;
}


// Solution Reviewer intentionally excludes complete/undo workflows.
// ================= LOADING OVERLAY =================
window.showLoading = function(message) {
  const loader = document.getElementById("loadingOverlay");
  if (!loader) return;

  loader.classList.remove("hidden");

  const text = loader.querySelector(".loading-text");
  if (text) {
    text.textContent = message || "Loading...";
  }
};

window.hideLoading = function(message) {
  const loader = document.getElementById("loadingOverlay");
  if (!loader) return;

  const text = loader.querySelector(".loading-text");

  if (message && text) {
    text.textContent = message;
    setTimeout(() => {
      loader.classList.add("hidden");
    }, 900);
  } else {
    loader.classList.add("hidden");
  }
};
//////
  
// ===== ROUTE + DAY COLLAPSIBLE =====
const routeDayToggle = document.getElementById("routeDayToggle");
const routeDayContent = document.getElementById("routeDayContent");

if (routeDayToggle && routeDayContent) {

  // Closed by default
  routeDayContent.classList.add("collapsed");

  routeDayToggle.addEventListener("click", (e) => {

    // Prevent clicking All/None from toggling collapse
    if (e.target.id === "routeDayAll" || e.target.id === "routeDayNone") return;

    const isCollapsed = routeDayContent.classList.toggle("collapsed");

    routeDayToggle.classList.toggle("open", !isCollapsed);
  });
}

// ===== DAYS COLLAPSIBLE =====
const daysToggle = document.getElementById("daysToggle");
const daysContent = document.getElementById("daysContent");

if (daysToggle && daysContent) {

  // Closed by default
  daysContent.classList.add("collapsed");

  daysToggle.addEventListener("click", (e) => {

    // Prevent clicking All/None from toggling collapse
    if (e.target.id === "daysAll" || e.target.id === "daysNone") return;

    const isCollapsed = daysContent.classList.toggle("collapsed");

    daysToggle.classList.toggle("open", !isCollapsed);
  });
}

  // ===== ROUTES COLLAPSIBLE =====
const routesToggle = document.getElementById("routesToggle");
const routesContent = document.getElementById("routesContent");

if (routesToggle && routesContent) {

  // Closed by default
  routesContent.classList.add("collapsed");

  routesToggle.addEventListener("click", (e) => {

    // Prevent clicking All/None from toggling collapse
    if (e.target.id === "routesAll" || e.target.id === "routesNone") return;

    const isCollapsed = routesContent.classList.toggle("collapsed");

    routesToggle.classList.toggle("open", !isCollapsed);
  });
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


  listFiles();
}

























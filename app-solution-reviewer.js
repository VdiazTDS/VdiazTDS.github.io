window.addEventListener("error", e => {
  console.error("JS ERROR:", e.message, "at line", e.lineno);
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
window.streetSegmentsEnabled = false;
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
  { label: "Cart Delivery App", href: "https://vdiaztds.github.io/cartdelivery/" },
  { label: "Sales-Polygon Viewer", href: "https://vdiaztds.github.io/Sales-Polygons-Viewer/" },
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
const map = L.map("map").setView([31.0, -99.0], 6);
// Shared Canvas renderer for high-performance drawing
const canvasRenderer = L.canvas({ padding: 0.5 });


// ===== BASE MAP LAYERS =====
const baseMaps = {
  streets: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 19
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

// ================= POLYGON SELECT =================
let streetLoadFilterLayer = null;
let streetFilterDrawPending = false;


// when polygon created
// ================= POLYGON SELECT =================
let drawnLayer = new L.FeatureGroup();
map.addLayer(drawnLayer);
const polygonDrawOptions = {};

const drawControl = new L.Control.Draw({
  draw: {
    polygon: polygonDrawOptions,
    rectangle: true,
    circle: false,
    marker: false,
    polyline: false,
    circlemarker: false
  },
  edit: { featureGroup: drawnLayer }
});

map.addControl(drawControl);

function getPolygonRings(latlngs) {
  if (!Array.isArray(latlngs) || !latlngs.length) return [];
  if (latlngs[0] && typeof latlngs[0].lat === "number" && typeof latlngs[0].lng === "number") {
    return [latlngs];
  }
  return latlngs.flatMap(getPolygonRings);
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

function isInsideStreetLoadFilter(lat, lon) {
  if (!window.streetSegmentsEnabled) return true;
  if (!streetLoadFilterLayer) return false;

  const latlng = L.latLng(lat, lon);
  if (typeof streetLoadFilterLayer.getBounds === "function") {
    const bounds = streetLoadFilterLayer.getBounds();
    if (bounds && !bounds.contains(latlng)) return false;
  }

  if (streetLoadFilterLayer instanceof L.Rectangle) return true;
  if (streetLoadFilterLayer instanceof L.Polygon) {
    const rings = getPolygonRings(streetLoadFilterLayer.getLatLngs());
    if (!rings.length) return false;
    return rings.some(ring => isPointInRing(latlng, ring));
  }

  return true;
}

function reloadMarkersUsingCurrentRows() {
  if (!Array.isArray(window._currentRows) || !window._currentRows.length || !window._currentWorkbook) return;
  processExcelBuffer(null, window._currentRows, window._currentWorkbook);
}

let streetSegmentsLoaderTimer = null;
function showStreetSegmentsLoader() {
  const wrap = document.getElementById("streetSegmentsLoader");
  const bar = document.getElementById("streetSegmentsLoaderBar");
  if (!wrap || !bar) return;

  if (streetSegmentsLoaderTimer) {
    clearInterval(streetSegmentsLoaderTimer);
    streetSegmentsLoaderTimer = null;
  }

  let pct = 8;
  wrap.classList.add("active");
  bar.style.width = `${pct}%`;

  streetSegmentsLoaderTimer = setInterval(() => {
    pct = Math.min(92, pct + (Math.random() * 10 + 3));
    bar.style.width = `${pct}%`;
  }, 120);
}

function hideStreetSegmentsLoader() {
  const wrap = document.getElementById("streetSegmentsLoader");
  const bar = document.getElementById("streetSegmentsLoaderBar");
  if (!wrap || !bar) return;

  if (streetSegmentsLoaderTimer) {
    clearInterval(streetSegmentsLoaderTimer);
    streetSegmentsLoaderTimer = null;
  }

  bar.style.width = "100%";
  setTimeout(() => {
    wrap.classList.remove("active");
    bar.style.width = "0%";
  }, 180);
}

function reloadStreetSegmentsWithLoader() {
  if (!Array.isArray(window._currentRows) || !window._currentRows.length || !window._currentWorkbook) return;
  showStreetSegmentsLoader();
  setTimeout(() => {
    try {
      reloadMarkersUsingCurrentRows();
    } finally {
      hideStreetSegmentsLoader();
    }
  }, 40);
}

function beginStreetFilterPolygonDraw() {
  streetFilterDrawPending = true;

  const toolbarHandler = drawControl?._toolbars?.draw?._modes?.polygon?.handler;
  if (toolbarHandler && typeof toolbarHandler.enable === "function") {
    toolbarHandler.enable();
    return;
  }

  const drawHandler = new L.Draw.Polygon(map, { ...polygonDrawOptions });
  drawHandler.enable();
}

// ===== SELECTION COUNT FUNCTION (GLOBAL & CORRECT) =====
function updateSelectionCount() {
const polygon = drawnLayer.getLayers()[0];
let count = 0;
const nextSelectedRowIds = new Set();

Object.entries(routeDayGroups).forEach(([key, group]) => {
 group.layers.forEach(marker => {
   const base = marker._base;
   if (!base) return;

   const latlng = L.latLng(base.lat, base.lon);

   const isLayerSelectMode = !!selectedLayerKey;
   const selectedByLayer = isLayerSelectMode && key === selectedLayerKey;
   const selectedByPolygon =
     !isLayerSelectMode && polygon && polygon.getBounds().contains(latlng);

   if ((selectedByLayer || selectedByPolygon) && map.hasLayer(marker)) {
     // highlight selected marker
     marker.setStyle?.({ color: "#ffff00", fillColor: "#ffff00" });

     if (Number.isFinite(marker._rowId)) {
       nextSelectedRowIds.add(marker._rowId);
     }
     count++; // ✅ only counting here
   } else {
     // restore original color
     const sym = symbolMap[key];
     marker.setStyle?.({ color: sym.color, fillColor: sym.color });
   }
 });
});

const prev = attributeState.selectedRowIds;
const changed =
  prev.size !== nextSelectedRowIds.size ||
  [...prev].some(id => !nextSelectedRowIds.has(id));

attributeState.selectedRowIds = nextSelectedRowIds;
syncSelectedStopsHeaderCount(count);
refreshAttributeStatus();
if (changed) renderAttributeTable();
}


// ===== COMPLETE SELECTED STOPS =====


  


// ===== WHEN POLYGON IS DRAWN =====
map.on(L.Draw.Event.CREATED, e => {
  selectedLayerKey = null;
  drawnLayer.clearLayers();
  drawnLayer.addLayer(e.layer);
  if (streetFilterDrawPending) {
    streetLoadFilterLayer = e.layer;
    streetFilterDrawPending = false;
    reloadStreetSegmentsWithLoader();
  }
  updateSelectionCount();
  updateUndoButtonState();   // 🔥 ADD THIS
});

map.on(L.Draw.Event.DRAWSTOP, () => {
  if (!streetFilterDrawPending) return;
  streetFilterDrawPending = false;
  const streetToggle = document.getElementById("streetLabelToggle");
  if (streetToggle) streetToggle.checked = false;
  window.streetLabelsEnabled = false;
  window.streetSegmentsEnabled = false;
  reloadMarkersUsingCurrentRows();
  map.fire("zoomend");
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
// ===== DELIVERED STOPS LAYER =====

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

  updateSelectionCount();
  updateStats();
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

function selectEntireLayer(key) {
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
  const selectedCount = attributeState.selectedRowIds.size;
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

function openAttributeTablePopout() {
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
    if (!isInsideStreetLoadFilter(lat, lon)) return;

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
async function listFiles() {
  const { data, error } = await sb.storage.from(BUCKET).list();
  if (error) return console.error(error);

  const ul = document.getElementById("savedFiles");
  ul.innerHTML = "";
  const routeFiles = data.filter(file => !isRouteSummaryFileName(file.name));
  const allFileNames = data.map(f => f.name);
  cleanupSummaryAttachments(allFileNames);

  routeFiles.forEach(file => {
    const routeName = file.name;
    const summaryName = resolveSummaryForRoute(routeName, data);

    const li = document.createElement("li");

    // OPEN MAP
    const openBtn = document.createElement("button");
    openBtn.className = "saved-file-btn";
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



    li.appendChild(openBtn);

    // SUMMARY BUTTON
    if (summaryName) {
      const summaryBtn = document.createElement("button");
      summaryBtn.textContent = "Summary";
      summaryBtn.style.marginLeft = "5px";
      summaryBtn.onclick = async () => {
        await loadSummaryFor(routeName);
        if (fileManagerModal) fileManagerModal.style.display = "none";
      };
      li.appendChild(summaryBtn);
    }

    // DELETE
    const delBtn = document.createElement("button");
    delBtn.textContent = "Delete";
    delBtn.style.marginLeft = "5px";

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


    li.appendChild(delBtn);
    li.appendChild(document.createTextNode(" " + routeName));
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
  if (typeof window._routeSummaryTableLayoutCleanup === "function") {
    window._routeSummaryTableLayoutCleanup();
    window._routeSummaryTableLayoutCleanup = null;
  }

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
  const headerCells = [];
  const columnCells = headers.map(() => []);
  const highlightedColumnIndexes = new Set();
  const frozenColumnIndexes = new Set();

  const updateColumnHighlightState = () => {
    headerCells.forEach((thNode, idx) => {
      const isActive = highlightedColumnIndexes.has(idx);
      thNode.classList.toggle("summary-column-active-header", isActive);
      thNode.setAttribute("aria-pressed", isActive ? "true" : "false");
    });

    columnCells.forEach((cells, idx) => {
      const isActive = highlightedColumnIndexes.has(idx);
      cells.forEach(cell => {
        cell.classList.toggle("summary-column-highlight", isActive);
      });
    });
  };

  const updateFrozenColumnState = () => {
    headerCells.forEach((thNode, idx) => {
      const isFrozen = frozenColumnIndexes.has(idx);
      thNode.classList.toggle("summary-column-frozen-header", isFrozen);
      if (!isFrozen) {
        thNode.style.left = "";
        thNode.style.zIndex = "";
      }
    });

    columnCells.forEach((cells, idx) => {
      const isFrozen = frozenColumnIndexes.has(idx);
      cells.forEach(cell => {
        cell.classList.toggle("summary-column-frozen-cell", isFrozen);
        if (!isFrozen) {
          cell.style.left = "";
          cell.style.zIndex = "";
        }
      });
    });

    let leftOffset = 0;
    [...frozenColumnIndexes]
      .sort((a, b) => a - b)
      .forEach(columnIndex => {
        const thNode = headerCells[columnIndex];
        if (!thNode) return;

        const thisLeft = Math.round(leftOffset);
        const width = thNode.getBoundingClientRect().width || thNode.offsetWidth || 0;

        thNode.style.left = `${thisLeft}px`;
        thNode.style.zIndex = "6";

        columnCells[columnIndex].forEach(cell => {
          cell.style.left = `${thisLeft}px`;
          cell.style.zIndex = "3";
        });

        leftOffset += width;
      });
  };

  const toggleColumnHighlight = columnIndex => {
    if (highlightedColumnIndexes.has(columnIndex)) {
      highlightedColumnIndexes.delete(columnIndex);
    } else {
      highlightedColumnIndexes.add(columnIndex);
    }
    updateColumnHighlightState();
  };

  const toggleColumnFrozen = columnIndex => {
    if (frozenColumnIndexes.has(columnIndex)) {
      frozenColumnIndexes.delete(columnIndex);
    } else {
      frozenColumnIndexes.add(columnIndex);
    }
    updateFrozenColumnState();
  };

  headers.forEach((h, colIndex) => {
    const th = document.createElement("th");
    th.textContent = h ?? "";
    th.classList.add("summary-column-selectable");
    th.setAttribute("tabindex", "0");
    th.setAttribute("role", "button");
    th.setAttribute("aria-pressed", "false");

    const headerLabel = String(h ?? "").trim() || `Column ${colIndex + 1}`;
    th.setAttribute("aria-label", `Highlight ${headerLabel}. Control or Command click to freeze.`);
    th.title = `Click to highlight ${headerLabel}. Ctrl/Cmd+Click to freeze.`;

    th.addEventListener("click", event => {
      if (event.ctrlKey || event.metaKey) {
        toggleColumnFrozen(colIndex);
        return;
      }
      toggleColumnHighlight(colIndex);
    });
    th.addEventListener("keydown", event => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) {
          toggleColumnFrozen(colIndex);
          return;
        }
        toggleColumnHighlight(colIndex);
        return;
      }
      if (String(event.key || "").toLowerCase() === "f") {
        event.preventDefault();
        toggleColumnFrozen(colIndex);
      }
    });

    headerCells.push(th);
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  rows.forEach(r => {
    const tr = document.createElement("tr");
    headers.forEach((h, colIndex) => {
      const td = document.createElement("td");
      td.textContent = r[h] ?? "";
      columnCells[colIndex].push(td);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  tableBox.appendChild(table);

  const syncFrozenColumns = () => {
    if (!frozenColumnIndexes.size) return;
    updateFrozenColumnState();
  };
  const handleWindowResize = () => {
    window.requestAnimationFrame(syncFrozenColumns);
  };

  let tableResizeObserver = null;
  if (typeof ResizeObserver !== "undefined") {
    tableResizeObserver = new ResizeObserver(() => {
      window.requestAnimationFrame(syncFrozenColumns);
    });
    tableResizeObserver.observe(table);
  }
  window.addEventListener("resize", handleWindowResize);
  window._routeSummaryTableLayoutCleanup = () => {
    window.removeEventListener("resize", handleWindowResize);
    if (tableResizeObserver) tableResizeObserver.disconnect();
  };

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

  if (!locateBtn || !headerContainer || !desktopContainer) return;

  if (window.innerWidth <= 900) {
    // 📱 MOBILE
    headerContainer.appendChild(locateBtn);

  } else {
    // 🖥 DESKTOP
    desktopContainer.appendChild(locateBtn);

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

// Clear selection button (ALWAYS ACTIVE)
if (clearBtn) {
  clearBtn.onclick = () => {
    selectedLayerKey = null;
    // Remove polygon
    drawnLayer.clearLayers();
    attributeState.selectedRowIds.clear();

    


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
  };
}







  
// ===== IMPORT WIZARD + FILE UPLOAD =====
const attributePanel = document.getElementById("attributeTablePanel");
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
    openAttributePanel();
    return;
  }
  closeAttributePanel();
};

attributeBtnDesktop?.addEventListener("click", toggleAttributePanel);
attributeBtnMobile?.addEventListener("click", toggleAttributePanel);

attributeCloseBtn?.addEventListener("click", () => {
  if (!attributePanel) return;
  attributePanel.classList.toggle("collapsed");
  syncAttributePanelLayout();
  refreshMapAfterOverlayChange();
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
  getFilteredAttributeRows().forEach(({ rowId }) => attributeState.selectedRowIds.add(rowId));
  applyAttributeSelectionStyles();
  syncSelectedStopsHeaderCount(attributeState.selectedRowIds.size);
  renderAttributeTable();
});

attributeClearSelectionBtn?.addEventListener("click", () => {
  attributeState.selectedRowIds.clear();
  applyAttributeSelectionStyles();
  syncSelectedStopsHeaderCount(0);
  renderAttributeTable();
});

attributeZoomSelectedBtn?.addEventListener("click", zoomToSelectedAttributeRows);
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
    streetLoadFilterLayer = null;
    streetFilterDrawPending = false;
    window.streetLabelsEnabled = false;
    window.streetSegmentsEnabled = false;
    const streetToggle = document.getElementById("streetLabelToggle");
    if (streetToggle) streetToggle.checked = false;

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
const streetSegmentsPromptModal = document.getElementById("streetSegmentsPromptModal");
const streetSegmentsPromptDrawBtn = document.getElementById("streetSegmentsPromptDraw");
const streetSegmentsPromptCancelBtn = document.getElementById("streetSegmentsPromptCancel");

function closeStreetSegmentsPrompt(restoreMap = false) {
  if (streetSegmentsPromptModal) streetSegmentsPromptModal.style.display = "none";
  if (!restoreMap) return;
  streetLoadFilterLayer = null;
  streetFilterDrawPending = false;
  window.streetLabelsEnabled = false;
  window.streetSegmentsEnabled = false;
  if (streetToggle) streetToggle.checked = false;
  reloadMarkersUsingCurrentRows();
  map.fire("zoomend");
}

function openStreetSegmentsPrompt() {
  if (!streetSegmentsPromptModal) {
    beginStreetFilterPolygonDraw();
    return;
  }
  streetSegmentsPromptModal.style.display = "flex";
}

if (streetToggle) {
  streetToggle.checked = false;
  window.streetLabelsEnabled = false;
  window.streetSegmentsEnabled = false;

  // Force labels to respect initial state
  map.whenReady(() => {
    map.fire("zoomend");
  });

  const handleStreetToggle = (enabled) => {
    if (enabled === window.streetSegmentsEnabled && enabled === window.streetLabelsEnabled) return;
    window.streetLabelsEnabled = enabled;
    window.streetSegmentsEnabled = enabled;

    if (enabled) {
      streetLoadFilterLayer = null;
      drawnLayer.clearLayers();
      // Hide all currently rendered records until a polygon is drawn.
      Object.values(routeDayGroups).forEach(group => {
        group.layers.forEach(layer => map.removeLayer(layer));
      });
      reloadMarkersUsingCurrentRows();
      openStreetSegmentsPrompt();
    } else {
      streetLoadFilterLayer = null;
      streetFilterDrawPending = false;
      drawnLayer.clearLayers();
      reloadMarkersUsingCurrentRows();
    }

    // Immediately refresh labels
    map.fire("zoomend");
  };

  streetToggle.addEventListener("change", (e) => handleStreetToggle(!!e.target.checked));
  streetToggle.addEventListener("input", (e) => handleStreetToggle(!!e.target.checked));
  streetToggle.addEventListener("click", () => {
    setTimeout(() => handleStreetToggle(!!streetToggle.checked), 0);
  });
}
  
////////////////////////////////////////////////////////////////////
if (streetSegmentsPromptDrawBtn) {
  streetSegmentsPromptDrawBtn.addEventListener("click", () => {
    closeStreetSegmentsPrompt(false);
    beginStreetFilterPolygonDraw();
  });
}

if (streetSegmentsPromptCancelBtn) {
  streetSegmentsPromptCancelBtn.addEventListener("click", () => {
    closeStreetSegmentsPrompt(true);
  });
}

if (streetSegmentsPromptModal) {
  streetSegmentsPromptModal.addEventListener("click", (e) => {
    if (e.target === streetSegmentsPromptModal) {
      closeStreetSegmentsPrompt(true);
    }
  });
}

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























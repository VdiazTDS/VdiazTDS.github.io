window.addEventListener("error", e => {
  try {
    const location = [
      String(e?.filename || "").trim(),
      Number.isFinite(Number(e?.lineno)) ? `line ${Number(e.lineno)}` : "",
      Number.isFinite(Number(e?.colno)) ? `col ${Number(e.colno)}` : ""
    ].filter(Boolean).join(", ");
    const stack = String(e?.error?.stack || "").trim();
    const detail = [location, stack].filter(Boolean).join("\n");
    recordAppDiagnosticEvent("error", "window.error", String(e?.message || "Unknown JS error"), detail);
  } catch {
    // Keep base error logging active even if diagnostics capture is not ready.
  }
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
const SHARED_FACILITIES_CLOUD_FILE = "__tds_pak_shared_facilities.xlsx";
const SUMMARY_ATTACH_CLOUD_FILE = "__tds_pak_summary_attachments.json";
const SHARED_FACILITIES_SHEET_NAME = "Facilities";
const SHARED_FACILITIES_CLOUD_SAVE_DEBOUNCE_MS = 1200;
const SUMMARY_ATTACH_CLOUD_SAVE_DEBOUNCE_MS = 700;
// ===== CURRENT EXCEL STATE =====
window._currentRows = null;
window._currentWorkbook = null;
window._currentFilePath = null;
window._summaryRows = [];
window._summaryHeaders = [];
window._attributeHeaders = [];
const WORKBOOK_AUTO_SAVE_DEBOUNCE_MS = 1200;
let workbookAutoSaveTimer = null;
let workbookAutoSaveInFlight = null;
let workbookAutoSavePendingReasons = [];
let sharedFacilitiesAutoSaveTimer = null;
let sharedFacilitiesAutoSaveInFlight = null;
let sharedFacilitiesAutoSavePendingReasons = [];
let sharedFacilitiesCloudInitPromise = null;

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
const STREET_ATTRIBUTE_TABLE_HEADERS = ["id", "name", "func_class", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"];
const STREET_ATTRIBUTE_TABLE_LABELS = {
  id: "Way ID",
  name: "Road Name",
  func_class: "Func Class",
  highway: "Class",
  ref: "Ref",
  maxspeed: "Max Speed",
  lanes: "Lanes",
  surface: "Surface",
  oneway: "One Way"
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

const RIGHT_SIDEBAR_ARROW_COLLAPSED = "\u276E";
const RIGHT_SIDEBAR_ARROW_EXPANDED = "\u276F";
const SUMMARY_ARROW_COLLAPSED = "\u25B2";
const SUMMARY_ARROW_EXPANDED = "\u25BC";
const APP_DIAGNOSTICS_MAX_EVENTS = 250;
const APP_DIAGNOSTICS_COPY_LIMIT = 80;
const appDiagnosticsSessionStartedAt = new Date().toISOString();
const appDiagnosticsEvents = [];

function stringifyAppDiagnosticValue(value) {
  if (value == null) return "";
  if (value instanceof Error) {
    const stack = String(value.stack || "").trim();
    if (stack) return stack;
    return `${value.name || "Error"}: ${value.message || ""}`.trim();
  }
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean" || typeof value === "bigint") {
    return String(value);
  }
  try {
    return JSON.stringify(value);
  } catch {
    try {
      return String(value);
    } catch {
      return "[unserializable]";
    }
  }
}

function formatAppDiagnosticArgs(args) {
  if (!Array.isArray(args) || !args.length) return "";
  return args
    .map(value => stringifyAppDiagnosticValue(value))
    .filter(Boolean)
    .join(" | ");
}

function recordAppDiagnosticEvent(level, source, message, detail = "") {
  const entry = {
    at: new Date().toISOString(),
    level: String(level || "info").toUpperCase(),
    source: String(source || "app").trim() || "app",
    message: String(message || "").trim() || "(no message)",
    detail: String(detail || "").trim()
  };
  appDiagnosticsEvents.push(entry);
  if (appDiagnosticsEvents.length > APP_DIAGNOSTICS_MAX_EVENTS) {
    appDiagnosticsEvents.splice(0, appDiagnosticsEvents.length - APP_DIAGNOSTICS_MAX_EVENTS);
  }
}

function getAppDiagnosticStatusText(nodeId) {
  const node = document.getElementById(nodeId);
  if (!node) return "";
  return String(node.textContent || "").replace(/\s+/g, " ").trim();
}

function getRouteDayGroupCountForDiagnostics() {
  try {
    return routeDayGroups && typeof routeDayGroups === "object"
      ? Object.keys(routeDayGroups).length
      : 0;
  } catch {
    return "n/a";
  }
}

function getStreetSegmentCountForDiagnostics() {
  try {
    return streetAttributeById && typeof streetAttributeById.size === "number"
      ? streetAttributeById.size
      : 0;
  } catch {
    return "n/a";
  }
}

function getSelectedStopCountForDiagnostics() {
  try {
    return selectedIds && typeof selectedIds.size === "number"
      ? selectedIds.size
      : 0;
  } catch {
    return "n/a";
  }
}

function buildAppDiagnosticsReport() {
  const nowIso = new Date().toISOString();
  const rowsLoaded = Array.isArray(window._currentRows) ? window._currentRows.length : 0;
  const recentEvents = appDiagnosticsEvents.slice(-APP_DIAGNOSTICS_COPY_LIMIT);
  const lines = [
    "=== TDS-PAK Diagnostics ===",
    `Generated: ${nowIso}`,
    `Session started: ${appDiagnosticsSessionStartedAt}`,
    `URL: ${window.location.href}`,
    `User agent: ${navigator.userAgent}`,
    `Online: ${navigator.onLine}`,
    `Viewport: ${window.innerWidth}x${window.innerHeight}`,
    `Screen: ${window.screen?.width || "?"}x${window.screen?.height || "?"}`,
    `Current file: ${String(window._currentFilePath || "(none)")}`,
    `Rows loaded: ${rowsLoaded}`,
    `Route+Day groups: ${getRouteDayGroupCountForDiagnostics()}`,
    `Selected stops: ${getSelectedStopCountForDiagnostics()}`,
    `Street segments loaded: ${getStreetSegmentCountForDiagnostics()}`,
    "",
    "Status Snapshot:",
    `- Street Load: ${getAppDiagnosticStatusText("streetLoadStatus") || "(blank)"}`,
    `- Route Sequencer: ${getAppDiagnosticStatusText("routeSequencerStatus") || "(blank)"}`,
    `- Set Segment ID + Side: ${getAppDiagnosticStatusText("setSegmentIdSideStatus") || "(blank)"}`,
    `- Update Field: ${getAppDiagnosticStatusText("updateFieldStatus") || "(blank)"}`,
    `- Select By Attributes: ${getAppDiagnosticStatusText("selectAttributesStatus") || "(blank)"}`,
    "",
    `Recent diagnostics (last ${recentEvents.length} of ${appDiagnosticsEvents.length} captured):`
  ];

  if (!recentEvents.length) {
    lines.push("(No warnings/errors captured in this tab session.)");
  } else {
    recentEvents.forEach((entry, index) => {
      lines.push(`${index + 1}. [${entry.at}] ${entry.level} ${entry.source}: ${entry.message}`);
      if (entry.detail) {
        lines.push(`   detail: ${entry.detail}`);
      }
    });
  }

  return lines.join("\n");
}

async function copyAppDiagnosticsToClipboard() {
  const report = buildAppDiagnosticsReport();
  let copied = false;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(report);
      copied = true;
    }
  } catch {
    copied = false;
  }

  if (!copied) {
    try {
      const node = document.createElement("textarea");
      node.value = report;
      node.setAttribute("readonly", "readonly");
      node.style.position = "fixed";
      node.style.opacity = "0";
      node.style.left = "-9999px";
      document.body.appendChild(node);
      node.focus();
      node.select();
      copied = !!document.execCommand("copy");
      node.remove();
    } catch {
      copied = false;
    }
  }

  if (copied) {
    alert("Diagnostics copied. Paste and send this text.");
    return true;
  }

  window.prompt("Copy diagnostics text below:", report);
  return false;
}

(function patchConsoleForDiagnostics() {
  if (window.__tdsPakConsoleDiagnosticsPatched) return;
  window.__tdsPakConsoleDiagnosticsPatched = true;

  const patch = (methodName, level) => {
    const original = console[methodName];
    if (typeof original !== "function") return;
    const bound = original.bind(console);
    console[methodName] = (...args) => {
      try {
        recordAppDiagnosticEvent(level, `console.${methodName}`, formatAppDiagnosticArgs(args));
      } catch {
        // Keep console behavior even if diagnostics capture fails.
      }
      return bound(...args);
    };
  };

  patch("error", "error");
  patch("warn", "warn");
})();

window.addEventListener("unhandledrejection", event => {
  const reason = stringifyAppDiagnosticValue(event?.reason);
  recordAppDiagnosticEvent("error", "window.unhandledrejection", reason || "Unhandled promise rejection");
});

window.getTdsPakDiagnosticsReport = buildAppDiagnosticsReport;

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
// ðŸ” Delete protection password I know this is not secure, I just wanted to make it harder for ppl to accidentally delete files. You can change or remove this as needed.
const DELETE_PASSWORD = "Austin1";  // â† change to whatever you want
// NOTE: This Street Network Manager password check is intentionally not secure.
// It is a temporary placeholder to reduce casual access for now.
const STREET_NETWORK_MANAGER_PASSWORD = "Austin1";


document.addEventListener("DOMContentLoaded", () => {
  initApp();
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
/* â­ Ensures mobile buttons move AFTER full page load */
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
const copyDiagnosticsBtn = document.getElementById("copyDiagnosticsBtn");

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

if (copyDiagnosticsBtn) {
  copyDiagnosticsBtn.addEventListener("click", async () => {
    try {
      await copyAppDiagnosticsToClipboard();
    } catch (error) {
      recordAppDiagnosticEvent(
        "error",
        "copy.diagnostics",
        "Copy Diagnostics failed",
        stringifyAppDiagnosticValue(error)
      );
      alert("Could not copy diagnostics automatically. Please try again.");
    }
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

function placeCopyDiagnosticsButton() {
  const diagnosticsBtn = document.getElementById("copyDiagnosticsBtn");
  const refreshBtn = document.getElementById("hardRefreshBtn");
  const desktopTools = document.querySelector(".header-tools-desktop");
  const mobileButtons = document.querySelector(".mobile-header-buttons");

  if (!diagnosticsBtn || !desktopTools || !mobileButtons) return;

  if (window.innerWidth > 900) {
    desktopTools.appendChild(diagnosticsBtn);
    return;
  }

  if (refreshBtn && refreshBtn.parentElement === mobileButtons) {
    mobileButtons.insertBefore(diagnosticsBtn, refreshBtn.nextSibling);
  } else {
    mobileButtons.insertBefore(diagnosticsBtn, mobileButtons.firstChild);
  }
}

function initHeaderActionOverflowMenu() {
  const row = document.querySelector(".header-import-selection-row");
  const wrapper = document.getElementById("headerActionsOverflow");
  const titleWrap = document.querySelector(".title-with-tools");
  const titleNode = titleWrap ? titleWrap.querySelector("h1") : null;
  if (!row || !titleWrap || !titleNode) return;
  if (row.dataset.inlineOverflowBound === "1") return;
  row.dataset.inlineOverflowBound = "1";

  const groups = [
    {
      label: "Selection & Query",
      items: [
        { id: "selectByAttributesBtn", label: "Select By Attributes" },
        { id: "setSegmentIdSideBtn", label: "Set Segment ID + Side" },
        { id: "updateFieldBtn", label: "Update Field" }
      ]
    },
    {
      label: "Street Tools",
      items: [
        { id: "streetAttributesBtn", label: "Street Attributes" }
      ]
    }
  ];

  const sourceButtons = [];
  const rails = {
    left: document.getElementById("headerOverflowRailLeft"),
    right: document.getElementById("headerOverflowRailRight")
  };

  if (!rails.left) {
    rails.left = document.createElement("div");
    rails.left.id = "headerOverflowRailLeft";
    rails.left.className = "header-overflow-rail header-overflow-rail-left";
  }
  if (!rails.right) {
    rails.right = document.createElement("div");
    rails.right.id = "headerOverflowRailRight";
    rails.right.className = "header-overflow-rail header-overflow-rail-right";
  }

  if (titleNode.previousElementSibling !== rails.left) {
    titleWrap.insertBefore(rails.left, titleNode);
  }
  if (titleNode.nextElementSibling !== rails.right) {
    titleWrap.insertBefore(rails.right, titleNode.nextSibling);
  }

  const buildInlineRails = () => {
    rails.left.innerHTML = "";
    rails.right.innerHTML = "";
    sourceButtons.length = 0;
    let entryCount = 0;
    let placementIndex = 0;

    groups.forEach(group => {
      group.items.forEach(item => {
        const sourceBtn = document.getElementById(item.id);
        if (!sourceBtn) return;
        sourceButtons.push(sourceBtn);
        const sourceClassName = sourceBtn.className
          .split(/\s+/)
          .filter(cls => cls && cls !== "header-overflow-source" && cls !== "header-overflow-source-active")
          .join(" ");
        sourceBtn.classList.add("header-overflow-source");

        const proxyBtn = document.createElement("button");
        proxyBtn.type = "button";
        proxyBtn.className = `${sourceClassName} header-overflow-inline-btn`;
        proxyBtn.innerHTML = sourceBtn.innerHTML;
        if (sourceBtn.title) proxyBtn.title = sourceBtn.title;
        proxyBtn.setAttribute(
          "aria-label",
          sourceBtn.getAttribute("aria-label") || item.label
        );
        proxyBtn.addEventListener("click", () => sourceBtn.click());

        if (placementIndex % 2 === 0) {
          rails.left.appendChild(proxyBtn);
        } else {
          rails.right.appendChild(proxyBtn);
        }
        placementIndex += 1;
        entryCount += 1;
      });
    });

    if (!entryCount) {
      wrapper.hidden = true;
      rails.left.hidden = true;
      rails.right.hidden = true;
      sourceButtons.forEach(btn => {
        btn.classList.remove("header-overflow-source");
        btn.classList.remove("header-overflow-source-active");
      });
      return;
    }

    wrapper.hidden = true;
  };

  const syncDesktopState = () => {
    const desktop = window.innerWidth > 900;
    const showInline = desktop && sourceButtons.length > 0;
    rails.left.hidden = !showInline;
    rails.right.hidden = !showInline;
    if (!showInline) {
      sourceButtons.forEach(btn => btn.classList.remove("header-overflow-source-active"));
      return;
    }
    sourceButtons.forEach(btn => btn.classList.add("header-overflow-source-active"));
  };

  buildInlineRails();
  syncDesktopState();
  window.addEventListener("resize", syncDesktopState);
}

placeRefreshButton();
placeCopyDiagnosticsButton();
initHeaderActionOverflowMenu();
window.addEventListener("resize", placeRefreshButton);
window.addEventListener("resize", placeCopyDiagnosticsButton);

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

function isSystemCloudFileName(fileName) {
  const lower = String(fileName || "").trim().toLowerCase();
  return (
    lower === SHARED_FACILITIES_CLOUD_FILE.toLowerCase() ||
    lower === SUMMARY_ATTACH_CLOUD_FILE.toLowerCase()
  );
}

const SUMMARY_ATTACH_STORAGE_KEY = storageKey("summaryAttachments");
let summaryAttachmentsCache = null;
let summaryAttachmentsInitPromise = null;
let summaryAttachmentsCloudSaveTimer = null;
let summaryAttachmentsCloudSaveInFlight = null;
let summaryAttachmentsCloudSaveQueued = false;

function normalizeSummaryAttachmentsMap(map) {
  const normalized = {};
  if (!map || typeof map !== "object") return normalized;

  Object.entries(map).forEach(([routeName, summaryName]) => {
    const route = String(routeName ?? "").trim();
    const summary = String(summaryName ?? "").trim();
    if (!route || !summary) return;
    normalized[route] = summary;
  });

  return normalized;
}

function readSummaryAttachmentsFromLocalStorage() {
  try {
    const raw = localStorage.getItem(SUMMARY_ATTACH_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return normalizeSummaryAttachmentsMap(parsed);
  } catch (_) {
    return {};
  }
}

function writeSummaryAttachmentsToLocalStorage(map) {
  localStorage.setItem(
    SUMMARY_ATTACH_STORAGE_KEY,
    JSON.stringify(normalizeSummaryAttachmentsMap(map))
  );
}

function summaryAttachmentMapsEqual(a, b) {
  const left = normalizeSummaryAttachmentsMap(a);
  const right = normalizeSummaryAttachmentsMap(b);
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every(key => left[key] === right[key]);
}

async function loadSummaryAttachmentsFromCloud() {
  const { data } = sb.storage.from(BUCKET).getPublicUrl(SUMMARY_ATTACH_CLOUD_FILE);
  const publicUrl = data?.publicUrl;
  if (!publicUrl) return {};

  const response = await fetch(publicUrl, { cache: "no-store" });
  if (!response.ok) {
    if (response.status === 404) return {};

    // Supabase can return 400 for missing objects, depending on bucket policy/settings.
    if (response.status === 400) {
      const text = (await response.text().catch(() => "")).toLowerCase();
      if (text.includes("not found") || text.includes("no such object")) {
        return {};
      }
    }

    throw new Error(`Failed to load summary attachments (HTTP ${response.status}).`);
  }

  const payload = await response.json().catch(() => ({}));
  return normalizeSummaryAttachmentsMap(payload);
}

async function saveSummaryAttachmentsToCloud(map) {
  const payload = JSON.stringify(normalizeSummaryAttachmentsMap(map), null, 2);
  const blob = new Blob([payload], { type: "application/json" });
  const { error } = await sb.storage
    .from(BUCKET)
    .upload(SUMMARY_ATTACH_CLOUD_FILE, blob, {
      upsert: true,
      contentType: "application/json"
    });

  if (error) throw error;
}

function scheduleSummaryAttachmentsCloudSave() {
  if (summaryAttachmentsCloudSaveTimer) clearTimeout(summaryAttachmentsCloudSaveTimer);
  summaryAttachmentsCloudSaveTimer = setTimeout(() => {
    summaryAttachmentsCloudSaveTimer = null;
    void flushSummaryAttachmentsCloudSave();
  }, SUMMARY_ATTACH_CLOUD_SAVE_DEBOUNCE_MS);
}

async function flushSummaryAttachmentsCloudSave() {
  if (summaryAttachmentsCloudSaveInFlight) {
    summaryAttachmentsCloudSaveQueued = true;
    return summaryAttachmentsCloudSaveInFlight;
  }

  const mapToSave = normalizeSummaryAttachmentsMap(
    summaryAttachmentsCache || readSummaryAttachmentsFromLocalStorage()
  );

  summaryAttachmentsCloudSaveInFlight = saveSummaryAttachmentsToCloud(mapToSave)
    .catch(error => {
      console.warn("Unable to save shared summary attachments:", error);
    })
    .finally(() => {
      summaryAttachmentsCloudSaveInFlight = null;
      if (summaryAttachmentsCloudSaveQueued) {
        summaryAttachmentsCloudSaveQueued = false;
        scheduleSummaryAttachmentsCloudSave();
      }
    });

  return summaryAttachmentsCloudSaveInFlight;
}

async function ensureSummaryAttachmentsInitialized() {
  if (summaryAttachmentsCache) return summaryAttachmentsCache;
  if (summaryAttachmentsInitPromise) return summaryAttachmentsInitPromise;

  summaryAttachmentsInitPromise = (async () => {
    const localMap = readSummaryAttachmentsFromLocalStorage();
    let cloudMap = {};

    try {
      cloudMap = await loadSummaryAttachmentsFromCloud();
    } catch (error) {
      console.warn("Unable to load shared summary attachments:", error);
    }

    // Keep cloud values authoritative for conflicts, while still migrating local-only links.
    const mergedMap = normalizeSummaryAttachmentsMap({ ...localMap, ...cloudMap });
    summaryAttachmentsCache = mergedMap;
    writeSummaryAttachmentsToLocalStorage(summaryAttachmentsCache);

    if (!summaryAttachmentMapsEqual(mergedMap, cloudMap)) {
      scheduleSummaryAttachmentsCloudSave();
    }

    return summaryAttachmentsCache;
  })().finally(() => {
    summaryAttachmentsInitPromise = null;
  });

  return summaryAttachmentsInitPromise;
}

function getSummaryAttachments() {
  const map = summaryAttachmentsCache || readSummaryAttachmentsFromLocalStorage();
  return { ...map };
}

function setSummaryAttachments(map) {
  summaryAttachmentsCache = normalizeSummaryAttachmentsMap(map);
  writeSummaryAttachmentsToLocalStorage(summaryAttachmentsCache);
  scheduleSummaryAttachmentsCloudSave();
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
// Allow fractional zoom levels so wheel-step amounts like 0.25 and 0.5 are visible.
map.options.zoomSnap = 0;
map.options.zoomDelta = 0.25;
const MAP_WHEEL_INPUT_AMOUNT_KEY = "mapWheelInputAmount";
const MAP_WHEEL_INPUT_PRESETS = Object.freeze({
  "very-small": {
    step: 0.25,
    title: "Very Small",
    hint: "0.25 zoom levels per wheel input."
  },
  "small": {
    step: 0.5,
    title: "Small",
    hint: "0.5 zoom levels per wheel input."
  },
  "medium": {
    step: 1,
    title: "Medium",
    hint: "1 zoom level per wheel input."
  },
  "large": {
    step: 2,
    title: "Large",
    hint: "2 zoom levels per wheel input."
  }
});
let mapWheelInputStep = 0.5;
let mapWheelZoomHandlerBound = false;
let mapWheelZoomAccumulator = 0;
let mapWheelZoomAccumulatorTimer = null;

function normalizeMapWheelInputAmount(value) {
  const token = String(value || "").trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(MAP_WHEEL_INPUT_PRESETS, token)) return token;
  return "small";
}

function applyMapWheelInputAmount(mode) {
  const normalized = normalizeMapWheelInputAmount(mode);
  const preset = MAP_WHEEL_INPUT_PRESETS[normalized];
  mapWheelInputStep = Math.max(0.1, Number(preset.step) || 0.5);
  return normalized;
}

function bindCustomMouseWheelZoomInput() {
  if (mapWheelZoomHandlerBound) return;
  mapWheelZoomHandlerBound = true;
  try { map.scrollWheelZoom?.disable?.(); } catch (_) {}
  const container = map.getContainer?.();
  if (!container) return;

  container.addEventListener("wheel", event => {
    if (!event || typeof event.deltaY !== "number" || event.deltaY === 0) return;
    event.preventDefault();

    const direction = event.deltaY < 0 ? 1 : -1;
    mapWheelZoomAccumulator += direction;

    if (mapWheelZoomAccumulatorTimer) {
      clearTimeout(mapWheelZoomAccumulatorTimer);
      mapWheelZoomAccumulatorTimer = null;
    }
    mapWheelZoomAccumulatorTimer = setTimeout(() => {
      mapWheelZoomAccumulator = 0;
      mapWheelZoomAccumulatorTimer = null;
    }, 160);

    if (Math.abs(mapWheelZoomAccumulator) < 1) return;
    const stepDirection = mapWheelZoomAccumulator > 0 ? 1 : -1;
    mapWheelZoomAccumulator = 0;

    const step = mapWheelInputStep;
    const targetZoom = map.getZoom() + (stepDirection * step);
    const minZoom = Number.isFinite(map.getMinZoom?.()) ? map.getMinZoom() : 0;
    const maxZoom = Number.isFinite(map.getMaxZoom?.()) ? map.getMaxZoom() : 20;
    const limitedTargetZoom = Math.max(minZoom, Math.min(maxZoom, targetZoom));

    const rect = container.getBoundingClientRect();
    const centerPoint = L.point(
      event.clientX - rect.left,
      event.clientY - rect.top
    );

    map.setZoomAround(centerPoint, limitedTargetZoom, { animate: false });
  }, { passive: false });
}

function initWheelZoomInputControls() {
  const openBtn = document.getElementById("wheelZoomInputBtn");
  const modal = document.getElementById("wheelZoomInputModal");
  const closeBtn = document.getElementById("wheelZoomInputCloseBtn");
  const optionsNode = document.getElementById("wheelZoomInputOptions");
  const statusNode = document.getElementById("wheelZoomInputStatus");
  const storageValue = storageGet(MAP_WHEEL_INPUT_AMOUNT_KEY);
  let activeMode = applyMapWheelInputAmount(storageValue);
  if (!storageValue) storageSet(MAP_WHEEL_INPUT_AMOUNT_KEY, activeMode);
  bindCustomMouseWheelZoomInput();

  const setStatusText = () => {
    if (!statusNode) return;
    const preset = MAP_WHEEL_INPUT_PRESETS[activeMode] || MAP_WHEEL_INPUT_PRESETS.small;
    statusNode.textContent = `Current: ${preset.title} (${preset.hint})`;
  };

  const setMode = mode => {
    const nextMode = applyMapWheelInputAmount(mode);
    if (nextMode === activeMode) return;
    activeMode = nextMode;
    storageSet(MAP_WHEEL_INPUT_AMOUNT_KEY, activeMode);
    renderOptions();
    setStatusText();
  };

  const renderOptions = () => {
    if (!optionsNode) return;
    optionsNode.innerHTML = "";
    Object.entries(MAP_WHEEL_INPUT_PRESETS).forEach(([key, preset]) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "wheel-zoom-option-btn";
      btn.classList.toggle("active", key === activeMode);
      btn.innerHTML = `<span class="wheel-zoom-option-title">${preset.title}</span><span class="wheel-zoom-option-hint">${preset.hint}</span>`;
      btn.addEventListener("click", () => setMode(key));
      optionsNode.appendChild(btn);
    });
  };

  renderOptions();
  setStatusText();

  if (!openBtn || !modal || !closeBtn) return;
  if (openBtn.dataset.wheelZoomBound === "1") return;
  openBtn.dataset.wheelZoomBound = "1";

  const openModal = () => {
    renderOptions();
    setStatusText();
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  openBtn.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });
  window.addEventListener("keydown", event => {
    if (modal.style.display !== "flex") return;
    if (event.key === "Escape") closeModal();
  });
}

initWheelZoomInputControls();
const ROUTE_DAY_PANE = "routeDayPane";
const STREET_NETWORK_PANE = "streetNetworkPane";
const SEQUENCE_PANE = "sequencePane";
const CITY_LIMITS_PANE = "cityLimitsPane";
const routeDayPane = map.createPane(ROUTE_DAY_PANE);
if (routeDayPane) routeDayPane.style.zIndex = "380";
const streetNetworkPane = map.createPane(STREET_NETWORK_PANE);
if (streetNetworkPane) streetNetworkPane.style.zIndex = "370";
const sequencePane = map.createPane(SEQUENCE_PANE);
if (sequencePane) {
  sequencePane.style.zIndex = "379";
  sequencePane.style.pointerEvents = "none";
}
const cityLimitsPane = map.createPane(CITY_LIMITS_PANE);
if (cityLimitsPane) {
  cityLimitsPane.style.zIndex = "365";
  cityLimitsPane.style.pointerEvents = "none";
}
// Shared Canvas renderer for high-performance drawing
const canvasRenderer = L.canvas({ padding: 0.5, pane: ROUTE_DAY_PANE });
// Keep all street casings under all street centerlines to prevent visual cut-outs.
const streetCasingCanvasRenderer = L.canvas({ padding: 0.5, pane: STREET_NETWORK_PANE });
const streetCenterCanvasRenderer = L.canvas({ padding: 0.5, pane: STREET_NETWORK_PANE });


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
  topographic: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 19
    }),
  terrain: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Terrain_Base/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 13
    }),
  natGeo: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/NatGeo_World_Map/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 16
    }),
  lightGray: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 16
    }),
  darkGray: L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      maxNativeZoom: 16
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
const MULTI_DAY_LABEL_PANE = "multiDayLabelPane";
const MULTI_DAY_LABEL_HARD_LIMIT = 12000;
const multiDayLabelPane = map.createPane(MULTI_DAY_LABEL_PANE);
if (multiDayLabelPane) {
  multiDayLabelPane.style.zIndex = "640";
  multiDayLabelPane.style.pointerEvents = "none";
}
const multiDayLabelLayerGroup = L.layerGroup().addTo(map);
const multiDayServiceByLocation = new Map();
const multiDayFrequencyCatalog = new Map();
let multiDayFrequencyVisibilityByKey = new Map();
const multiDayLabelDayCatalog = new Map();
let multiDayLabelDayVisibilityByToken = new Map();
let multiDayLabelFrequencyVisibilityByKey = new Map();
let multiDayServiceEnabled = false;
let multiDayToggleButtonRef = null;
let multiDayManagerUiRenderer = null;
let multiDaySidebarUiRenderer = null;
const cityLimitsLayerGroup = L.layerGroup();
const streetAttributeLayerGroup = L.layerGroup();
const texasLandfillsLayerGroup = L.layerGroup();
const routeSequencerFacilitiesLayerGroup = L.layerGroup().addTo(map);
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
const STREET_FUNC_CLASS_VALUE_ORDER = ["1", "2", "3", "4", "5"];
const STREET_FUNC_CLASS_DEFAULT_COLORS = Object.freeze({
  "1": "#ff3b30",
  "2": "#f4d447",
  "3": "#1f1f1f",
  "4": "#5f6670",
  "5": "#8a9099",
  [STREET_SYMBOLOGY_EMPTY_KEY]: "#b8bec7"
});
const STREET_FUNC_CLASS_DEFAULT_WEIGHTS = Object.freeze({
  "1": 5.0,
  "2": 4.3,
  "3": 3.4,
  "4": 2.7,
  "5": 2.2
});
const STREET_FUNC_CLASS_CASING_COLORS = Object.freeze({
  "1": "#861714",
  "2": "#151515",
  "3": "#101010",
  "4": "#3f454d",
  "5": "#5a6069",
  [STREET_SYMBOLOGY_EMPTY_KEY]: "#636a74"
});
const STREET_FUNC_CLASS_DEFAULT_LINE_WIDTH = STREET_BASE_LINE_WEIGHT;
const STREET_FUNC_CLASS_DEFAULT_OPACITY = 0.96;
const STREET_CASING_FALLBACK_COLOR = "#242a33";
const STREET_SELECTION_INNER_COLOR = "#ffe066";
const STREET_SELECTION_OUTER_COLOR = "#101114";
const STREET_CASING_AUTO_DISABLE_SEGMENT_COUNT = 6000;
const STREET_FUNC_CLASS_PREVIOUS_DEFAULT_COLORS = Object.freeze({
  "1": "#f44336",
  "2": "#f2cc4d",
  "3": "#1f2937",
  "4": "#6b7280",
  "5": "#9ca3af",
  [STREET_SYMBOLOGY_EMPTY_KEY]: "#b0b8c2"
});
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

function darkenStreetHexColor(colorHex, amount = 0.55) {
  const text = String(colorHex || "").trim();
  if (!/^#[0-9a-fA-F]{6}$/.test(text)) return STREET_CASING_FALLBACK_COLOR;
  const factor = 1 - Math.max(0, Math.min(0.9, Number(amount) || 0));
  const r = Math.max(0, Math.min(255, Math.round(parseInt(text.slice(1, 3), 16) * factor)));
  const g = Math.max(0, Math.min(255, Math.round(parseInt(text.slice(3, 5), 16) * factor)));
  const b = Math.max(0, Math.min(255, Math.round(parseInt(text.slice(5, 7), 16) * factor)));
  return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
}
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
  "bus_guideway"
]);
const LOCAL_STREET_NON_MOTOR_CLASS_DENYLIST = new Set([
  "pedestrian",
  "footway",
  "cycleway",
  "bridleway",
  "path",
  "steps",
  "corridor",
  "platform",
  "via_ferrata",
  "proposed",
  "construction"
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
const LOCAL_STREET_AUTO_SETUP_PACKAGE = "tds-streets-auto-setup-package.zip?v=20260311-3";
const LOCAL_STREET_BACKEND_URL_KEY = "localStreetBackendUrl";
const STREET_NETWORK_LAYER_VISIBLE_KEY = "streetNetworkLayerVisible";
const LOCAL_STREET_BACKEND_URL_DEFAULT = "http://127.0.0.1:8787";
const LOCAL_STREET_BACKEND_HEALTH_TTL_MS = 12000;
const LOCAL_STREET_BACKEND_REQUEST_TIMEOUT_MS = 25000;
const LOCAL_STREET_BACKEND_QUERY_LIMIT = 220000;
const LOCAL_STREET_BACKEND_QUERY_TILE_MAX_DEPTH = 4;
const LOCAL_STREET_BACKEND_QUERY_TILE_MIN_SPAN_DEG = 0.01;
const TEXAS_LANDFILLS_LAYER_VISIBLE_KEY = "texasLandfillsLayerVisible";
const TEXAS_LANDFILLS_FEATURE_LAYER_URL = "https://services6.arcgis.com/MhhE5pgCPypps4To/ArcGIS/rest/services/Texas_Landfills_Web_Map/FeatureServer/0";
const TEXAS_LANDFILLS_QUERY_URL = `${TEXAS_LANDFILLS_FEATURE_LAYER_URL}/query`;
const TEXAS_LANDFILLS_FETCH_BATCH_SIZE = 180;
const CITY_LIMITS_LAYER_VISIBLE_KEY = "cityLimitsLayerVisible";
const CITY_LIMITS_FEATURE_LAYER_URL = "https://services.arcgis.com/P3ePLMYs2RVChkJx/ArcGIS/rest/services/USA_Census_Populated_Places/FeatureServer/0";
const CITY_LIMITS_QUERY_URL = `${CITY_LIMITS_FEATURE_LAYER_URL}/query`;
const CITY_LIMITS_FETCH_RECORD_LIMIT = 900;
const CITY_LIMITS_QUERY_WHERE = "ST = 'TX'";
const CITY_LIMITS_MOVE_REFRESH_DEBOUNCE_MS = 320;
const CITY_LIMITS_MIN_ZOOM = 8;
const CITY_LIMITS_LABEL_MIN_ZOOM = 10;
const CITY_LIMITS_BOUNDS_PAD = 0.04;
const TEXAS_LANDFILL_MAJOR_TYPE_ALLOWLIST = new Set(["1", "1AE", "1 AE & 4 AE", "1AE & 4AE", "2"]);
const TEXAS_LANDFILL_EXCLUDED_SITE_STATUSES = new Set(["INACTIVE", "CLOSED", "POST CLOSURE", "NOT CONSTRUCTED", "MISSING"]);
const TEXAS_LANDFILL_EXCLUDED_CEC_STATUSES = new Set(["INACTIVE"]);
const TEXAS_LANDFILL_EXCLUDED_LEGAL_STATUSES = new Set(["DENIED", "REVOKED", "WITHDRAWN", "RETURNED"]);
const texasLandfillsState = {
  loaded: false,
  loading: false,
  loadPromise: null,
  featureCount: 0,
  sourceFeatureCount: 0,
  filteredOutCount: 0,
  lastError: "",
  sourceName: "Texas Landfills Web Map"
};
const cityLimitsState = {
  loaded: false,
  loading: false,
  loadPromise: null,
  featureCount: 0,
  lastError: "",
  lastBoundsKey: "",
  refreshTimer: null,
  moveListenerBound: false,
  zoomListenerBound: false,
  sourceName: "USA Census Populated Places"
};
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
  enabled: true,
  field: "func_class",
  valueColors: { ...STREET_FUNC_CLASS_DEFAULT_COLORS },
  lineWidth: STREET_FUNC_CLASS_DEFAULT_LINE_WIDTH,
  opacity: STREET_FUNC_CLASS_DEFAULT_OPACITY
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

function cloneStreetFuncClassDefaultColors() {
  return { ...STREET_FUNC_CLASS_DEFAULT_COLORS };
}

function normalizeStreetFuncClassValue(value) {
  const text = String(value ?? "").trim().toUpperCase();
  if (!text || text === "UNKNOWN") return "";
  if (/^[1-5]$/.test(text)) return text;
  const firstDigit = text.match(/[1-5]/);
  if (firstDigit) return firstDigit[0];
  return "";
}

function deriveStreetFuncClassFromTags(props = {}) {
  const source = props && typeof props === "object" ? props : {};

  const directCandidates = [
    source.func_class,
    source.funcClass,
    source.FUNC_CLASS,
    source.funclass,
    source.FUNCLASS
  ];
  for (const candidate of directCandidates) {
    const direct = normalizeStreetFuncClassValue(candidate);
    if (direct) return direct;
  }

  const rawHighway = String(
    source.highway ??
    source.HIGHWAY ??
    source.road_class ??
    source.fclass ??
    source.FCLASS ??
    ""
  ).trim().toLowerCase();
  if (!rawHighway || rawHighway === "unknown") return "";

  if (["motorway", "motorway_link", "trunk", "trunk_link"].includes(rawHighway)) return "1";
  if (["primary", "primary_link"].includes(rawHighway)) return "2";
  if (["secondary", "secondary_link"].includes(rawHighway)) return "3";
  if (["tertiary", "tertiary_link"].includes(rawHighway)) return "4";
  return "5";
}

function isLegacyStreetSymbologyDefault(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return false;
  const field = String(raw.field || "").trim().toLowerCase();
  const enabled = !!raw.enabled;
  const colors = raw.valueColors && typeof raw.valueColors === "object" && !Array.isArray(raw.valueColors)
    ? Object.keys(raw.valueColors)
    : [];
  const width = Number(raw.lineWidth);
  const opacity = Number(raw.opacity);

  return (
    field === "highway" &&
    enabled === false &&
    colors.length === 0 &&
    (!Number.isFinite(width) || width === STREET_BASE_LINE_WEIGHT) &&
    (!Number.isFinite(opacity) || opacity === STREET_BASE_LINE_OPACITY)
  );
}

function isPreviousFuncClassSymbologyDefault(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return false;
  const field = String(raw.field || "").trim().toLowerCase();
  if (field !== "func_class") return false;

  const colors = raw.valueColors && typeof raw.valueColors === "object" && !Array.isArray(raw.valueColors)
    ? raw.valueColors
    : null;
  if (!colors) return false;

  const oldKeys = Object.keys(STREET_FUNC_CLASS_PREVIOUS_DEFAULT_COLORS);
  const rawKeys = Object.keys(colors).map(k => String(k || "").trim()).filter(Boolean);
  if (rawKeys.length !== oldKeys.length) return false;
  if (!oldKeys.every(key => rawKeys.includes(key))) return false;

  const allMatch = oldKeys.every(key =>
    String(colors[key] || "").trim().toLowerCase() === String(STREET_FUNC_CLASS_PREVIOUS_DEFAULT_COLORS[key]).toLowerCase()
  );
  if (!allMatch) return false;

  const width = Number(raw.lineWidth);
  const opacity = Number(raw.opacity);
  return (
    (!Number.isFinite(width) || width === STREET_BASE_LINE_WEIGHT) &&
    (!Number.isFinite(opacity) || opacity === STREET_BASE_LINE_OPACITY)
  );
}

function normalizeStreetSymbologyState(raw) {
  const base = {
    enabled: true,
    field: "func_class",
    valueColors: cloneStreetFuncClassDefaultColors(),
    lineWidth: STREET_FUNC_CLASS_DEFAULT_LINE_WIDTH,
    opacity: STREET_FUNC_CLASS_DEFAULT_OPACITY
  };
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return base;
  if (isLegacyStreetSymbologyDefault(raw)) return base;
  if (isPreviousFuncClassSymbologyDefault(raw)) return base;

  const field = String(raw.field || base.field).trim() || base.field;
  const valueColors = String(field).toLowerCase() === "func_class"
    ? cloneStreetFuncClassDefaultColors()
    : {};
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
  const fieldLower = String(field).toLowerCase();
  return {
    enabled: raw.enabled === undefined ? base.enabled : !!raw.enabled,
    field,
    valueColors,
    lineWidth: Number.isFinite(widthNum)
      ? Math.max(1, Math.min(8, widthNum))
      : (fieldLower === "func_class" ? STREET_FUNC_CLASS_DEFAULT_LINE_WIDTH : STREET_BASE_LINE_WEIGHT),
    opacity: Number.isFinite(opacityNum)
      ? Math.max(0.2, Math.min(1, opacityNum))
      : (fieldLower === "func_class" ? STREET_FUNC_CLASS_DEFAULT_OPACITY : STREET_BASE_LINE_OPACITY)
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
  const preferredOrder = ["func_class", "highway", "surface", "lanes", "maxspeed", "oneway", "name", "ref", "id"];
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
  const fieldLower = field.toLowerCase();
  const counts = new Map();
  streetAttributeById.forEach(entry => {
    const row = entry?.row || {};
    const valueKey = normalizeStreetSymbologyValueKey(row?.[field]);
    const item = counts.get(valueKey) || { key: valueKey, label: formatStreetSymbologyValueLabel(valueKey), count: 0 };
    item.count += 1;
    counts.set(valueKey, item);
  });

  if (fieldLower === "func_class") {
    STREET_FUNC_CLASS_VALUE_ORDER.forEach(key => {
      if (counts.has(key)) return;
      counts.set(key, { key, label: key, count: 0 });
    });
  }

  const values = [...counts.values()];
  if (fieldLower === "func_class") {
    const rank = new Map(STREET_FUNC_CLASS_VALUE_ORDER.map((key, idx) => [key, idx]));
    return values.sort((a, b) => {
      const aRank = rank.has(a.key) ? rank.get(a.key) : 99;
      const bRank = rank.has(b.key) ? rank.get(b.key) : 99;
      if (aRank !== bRank) return aRank - bRank;
      if (b.count !== a.count) return b.count - a.count;
      return a.label.localeCompare(b.label, undefined, { sensitivity: "base" });
    });
  }

  return values.sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.label.localeCompare(b.label, undefined, { sensitivity: "base" });
  });
}

function resolveStreetSymbologyColorForEntry(entry) {
  const state = ensureStreetSymbologyState();
  if (!state.enabled) return STREET_BASE_LINE_COLOR;
  const field = String(state.field || "func_class");
  const valueKey = normalizeStreetSymbologyValueKey(entry?.row?.[field]);
  const explicit = state.valueColors?.[valueKey];
  if (typeof explicit === "string" && /^#[0-9a-fA-F]{6}$/.test(explicit)) {
    return explicit.toLowerCase();
  }
  if (field.toLowerCase() === "func_class") {
    const classColor = STREET_FUNC_CLASS_DEFAULT_COLORS[valueKey] || STREET_FUNC_CLASS_DEFAULT_COLORS[STREET_SYMBOLOGY_EMPTY_KEY];
    if (classColor) return classColor.toLowerCase();
  }
  return getStreetSymbologyPaletteColorByKey(valueKey);
}

function resolveStreetSymbologyWeightForEntry(entry) {
  const state = ensureStreetSymbologyState();
  if (!state.enabled) return STREET_BASE_LINE_WEIGHT;

  const width = Number.isFinite(Number(state.lineWidth))
    ? Math.max(1, Math.min(8, Number(state.lineWidth)))
    : STREET_BASE_LINE_WEIGHT;

  const field = String(state.field || "func_class");
  if (field.toLowerCase() !== "func_class") return width;

  const valueKey = normalizeStreetSymbologyValueKey(entry?.row?.[field]);
  const classWeight = Number(STREET_FUNC_CLASS_DEFAULT_WEIGHTS[valueKey]);
  if (!Number.isFinite(classWeight)) return width;

  const scaled = classWeight * (width / STREET_BASE_LINE_WEIGHT);
  return Math.max(1, Math.min(8, scaled));
}

function resolveStreetSymbologyCasingStyleForEntry(entry, centerStyle) {
  const state = ensureStreetSymbologyState();
  const field = String(state.field || "func_class");
  const valueKey = normalizeStreetSymbologyValueKey(entry?.row?.[field]);
  const centerColor = String(centerStyle?.color || resolveStreetSymbologyColorForEntry(entry)).toLowerCase();
  const centerWeight = Number(centerStyle?.weight);
  const centerOpacity = Number(centerStyle?.opacity);
  const explicit = state.valueColors?.[valueKey];

  let casingColor = STREET_CASING_FALLBACK_COLOR;
  if (field.toLowerCase() === "func_class" && !(typeof explicit === "string" && /^#[0-9a-fA-F]{6}$/.test(explicit))) {
    casingColor = String(STREET_FUNC_CLASS_CASING_COLORS[valueKey] || "").trim().toLowerCase()
      || darkenStreetHexColor(centerColor, 0.58);
  } else {
    casingColor = darkenStreetHexColor(centerColor, 0.55);
  }

  const safeCenterWeight = Number.isFinite(centerWeight)
    ? centerWeight
    : resolveStreetSymbologyWeightForEntry(entry);
  const safeCenterOpacity = Number.isFinite(centerOpacity)
    ? centerOpacity
    : (state.enabled ? state.opacity : STREET_BASE_LINE_OPACITY);
  const casingWeight = Math.max(
    safeCenterWeight + 0.9,
    Math.min(10, safeCenterWeight * 1.28)
  );

  return {
    color: casingColor,
    weight: casingWeight,
    opacity: Math.max(0.35, Math.min(1, safeCenterOpacity + 0.02))
  };
}

function getStreetSegmentBaseStyle(entry) {
  const state = ensureStreetSymbologyState();
  const center = {
    color: resolveStreetSymbologyColorForEntry(entry),
    weight: resolveStreetSymbologyWeightForEntry(entry),
    opacity: state.enabled ? state.opacity : STREET_BASE_LINE_OPACITY
  };
  const casing = resolveStreetSymbologyCasingStyleForEntry(entry, center);
  return {
    color: center.color,
    weight: center.weight,
    opacity: center.opacity,
    casingColor: casing.color,
    casingWeight: casing.weight,
    casingOpacity: casing.opacity
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
    hintNode.textContent = "Start with Street Setup Wizard: download setup package, run auto-setup, then open Street Backend Manager and click Check Backend.";
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

function updateStreetNetworkSidebarBackendPrompt() {
  const promptNode = document.getElementById("streetNetworkSidebarBackendPrompt");
  if (!promptNode) return;
  const layerToggle = document.getElementById("streetNetworkLayerToggle");
  const layerRequested = !!layerToggle?.checked;

  if (!layerRequested) {
    promptNode.hidden = true;
    return;
  }

  const backendReady = !!localStreetBackendState.available && !!localStreetBackendState.hasIndex;
  if (backendReady) {
    promptNode.hidden = true;
    return;
  }

  if (localStreetBackendState.checking) {
    promptNode.textContent = "Checking Street Network backend connection...";
    promptNode.hidden = false;
    return;
  }

  promptNode.textContent = "Street Network backend is not connected. Open Street Network Manager to complete setup and start the backend.";
  promptNode.hidden = false;
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
  updateStreetNetworkSidebarBackendPrompt();
  if (!node) return;
  if (message) {
    node.textContent = message;
    return;
  }
  if (!hasProvider) {
    node.textContent = "Street layer: Off. Click Street Setup Wizard, run setup, then use Street Backend Manager and complete steps 1-3.";
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

function normalizeLocalStreetClassValue(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (!text || text === "unknown") return "";
  return text;
}

function getLocalStreetClassFromTags(props = {}) {
  const source = props && typeof props === "object" ? props : {};
  const candidates = [
    source.highway,
    source.HIGHWAY,
    source.road_class,
    source.ROAD_CLASS,
    source.fclass,
    source.FCLASS,
    source.class,
    source.CLASS,
    source.road_type,
    source.ROAD_TYPE,
    source.route,
    source.ROUTE
  ];
  for (let i = 0; i < candidates.length; i += 1) {
    const normalized = normalizeLocalStreetClassValue(candidates[i]);
    if (normalized) return normalized;
  }
  return "";
}

function isMotorVehicleStreetClassValue(classValue) {
  const value = normalizeLocalStreetClassValue(classValue);
  if (!value) return false;
  if (LOCAL_STREET_NON_MOTOR_CLASS_DENYLIST.has(value)) return false;
  if (LOCAL_STREET_FCLASS_WHITELIST.has(value)) return true;
  if (/(motorway|trunk|primary|secondary|tertiary|residential|unclassified|living_street|service|road|track|_link|bus_guideway)/.test(value)) {
    return true;
  }
  return false;
}

function isMotorVehicleStreetTags(tags = {}, layerName = "") {
  const classValue = getLocalStreetClassFromTags(tags);
  if (classValue) return isMotorVehicleStreetClassValue(classValue);
  return LOCAL_STREET_ROAD_LAYER_RX.test(String(layerName || ""));
}

function isLocalStreetFeature(feature, layerName = "") {
  const geometryType = feature?.geometry?.type;
  if (geometryType !== "LineString" && geometryType !== "MultiLineString") return false;

  const props = feature?.properties || {};
  return isMotorVehicleStreetTags(props, layerName);
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

  const resolvedFuncClass = deriveStreetFuncClassFromTags(props);

  return {
    name: normalizeTagValue(props.name ?? props.NAME),
    highway: normalizeTagValue(props.highway ?? props.HIGHWAY ?? props.road_class ?? props.fclass ?? props.FCLASS),
    func_class: normalizeTagValue(resolvedFuncClass, "5"),
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
  ["name", "highway", "func_class", "ref", "maxspeed", "lanes", "surface", "oneway"].forEach(key => {
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
  updateLocalStreetSourceStatus("Downloading Texas source ZIP from Geofabrik... then run auto-setup and open Backend Manager.");
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
    "Downloaded setup package. Extract it, run tds-local-streets-auto-setup-launcher.cmd, then open Backend Manager. Existing users can run tds-street-backend-manager-install-launcher.cmd only."
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
      ? "\n\nBrowsers usually block direct local filesystem paths. Use Street Setup Wizard + Backend Manager instead."
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
    alert(`Could not load streets from that path.\n\n${err?.message || err}\n\nUse Street Setup Wizard and Backend Manager instead.`);
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
      : "Recommended path: Download Setup Package, run auto-setup launcher, open Backend Manager, Check Backend, then start polygon load."
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
  const entered = window.prompt("Enter password to open Street Network Manager:");
  if (entered === null) return;
  if (String(entered) !== STREET_NETWORK_MANAGER_PASSWORD) {
    alert("Incorrect password.");
    return;
  }
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
    const selectedField = String(streetSymbologyFieldSelect.value || streetSymbologyState.field || "func_class");
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
      streetSymbologyState.field = fields[0] || "func_class";
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
    streetSymbologyState.field = String(streetSymbologyFieldSelect.value || "func_class");
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
      const backendUrl = String(localStreetBackendState.baseUrl || LOCAL_STREET_BACKEND_URL_DEFAULT);
      const sourceText = String(localStreetBackendState.sourceName || "").trim();
      setStreetWizardStatus(`Backend is ready at ${backendUrl}.`);
      if (showAlert) {
        alert(
          "Backend is ready.\n\n" +
          `URL: ${backendUrl}\n` +
          (sourceText ? `Source: ${sourceText}\n` : "") +
          "If this is not your local machine backend, click Set Backend URL and switch to http://127.0.0.1:8787."
        );
      }
      return true;
    }
    if (localStreetBackendState.available && !localStreetBackendState.hasIndex) {
      setStreetWizardStatus("Backend is running but no streets index is loaded.");
      if (showAlert) {
        alert(
          "Local backend is running, but no streets index is loaded.\n\n" +
          "Run the setup package again (or open indexer from Backend Manager), then click Check Backend again."
        );
      }
      return false;
    }

    setStreetWizardStatus(`Backend not reachable at ${localStreetBackendState.baseUrl}.`);
    if (showAlert) {
      alert(
        "Local backend is not reachable.\n\n" +
        `URL: ${localStreetBackendState.baseUrl}\n` +
        "Start backend from the Street Backend Manager desktop app, then check again."
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
    setStreetWizardStatus("Installer package downloaded.");
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
    setStreetWizardStatus("Texas source ZIP download started. Next: run auto-setup launcher, then check backend.");
  });

  streetWizardAutoSetupBtn?.addEventListener("click", () => {
    downloadLocalStreetAutoSetupPackage();
    setStreetWizardStatus("Setup package downloaded. Run auto-setup launcher, then use Backend Manager and click Check Backend.");
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
  resetStreetRenderingPerformanceMode();
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

async function fetchStreetElementsFromLocalBackendBounds(bounds, limit = LOCAL_STREET_BACKEND_QUERY_LIMIT) {
  const south = bounds.getSouth();
  const west = bounds.getWest();
  const north = bounds.getNorth();
  const east = bounds.getEast();
  const query = new URLSearchParams({
    south: String(south),
    west: String(west),
    north: String(north),
    east: String(east),
    limit: String(Math.max(1000, Number(limit) || LOCAL_STREET_BACKEND_QUERY_LIMIT))
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
  const rawCount = Number(payload?.count);
  const truncated = payload?.truncated === true || String(payload?.truncated || "").toLowerCase() === "true";
  return {
    rows,
    count: Number.isFinite(rawCount) ? rawCount : rows.length,
    truncated
  };
}

function shouldSplitLocalBackendQueryBounds(bounds, depth, maxDepth) {
  if (!isValidLeafletBounds(bounds)) return false;
  if (depth >= maxDepth) return false;
  const latSpan = Math.abs(bounds.getNorth() - bounds.getSouth());
  const lonSpan = Math.abs(bounds.getEast() - bounds.getWest());
  return latSpan > LOCAL_STREET_BACKEND_QUERY_TILE_MIN_SPAN_DEG
    || lonSpan > LOCAL_STREET_BACKEND_QUERY_TILE_MIN_SPAN_DEG;
}

function splitLocalBackendQueryBounds(bounds) {
  if (!isValidLeafletBounds(bounds)) return [];
  const south = bounds.getSouth();
  const west = bounds.getWest();
  const north = bounds.getNorth();
  const east = bounds.getEast();
  const midLat = (south + north) / 2;
  const midLon = (west + east) / 2;
  const quads = [
    L.latLngBounds([south, west], [midLat, midLon]),
    L.latLngBounds([south, midLon], [midLat, east]),
    L.latLngBounds([midLat, west], [north, midLon]),
    L.latLngBounds([midLat, midLon], [north, east])
  ];
  return quads.filter(isValidLeafletBounds);
}

async function collectStreetElementsFromLocalBackend(bounds, options = {}) {
  const maxDepth = Math.max(1, Number(options.maxDepth) || LOCAL_STREET_BACKEND_QUERY_TILE_MAX_DEPTH);
  const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;
  const rowById = new Map();
  const stats = {
    requestCount: 0,
    totalRowsFromBackend: 0,
    splitQueryCount: 0,
    unresolvedTruncatedQueries: 0
  };

  const ingestRows = (rows) => {
    rows.forEach(raw => {
      const id = Number(raw?.id);
      if (!Number.isFinite(id)) return;
      if (rowById.has(id)) return;
      rowById.set(id, raw);
    });
  };

  const walk = async (targetBounds, depth) => {
    stats.requestCount += 1;
    if (onProgress) onProgress(stats.requestCount, depth);
    const response = await fetchStreetElementsFromLocalBackendBounds(targetBounds, LOCAL_STREET_BACKEND_QUERY_LIMIT);
    stats.totalRowsFromBackend += Number(response.count) || response.rows.length;

    const shouldSplit = response.truncated && shouldSplitLocalBackendQueryBounds(targetBounds, depth, maxDepth);
    if (shouldSplit) {
      stats.splitQueryCount += 1;
      const chunks = splitLocalBackendQueryBounds(targetBounds);
      for (let i = 0; i < chunks.length; i += 1) {
        // eslint-disable-next-line no-await-in-loop
        await walk(chunks[i], depth + 1);
      }
      return;
    }

    ingestRows(response.rows);
    if (response.truncated) {
      stats.unresolvedTruncatedQueries += 1;
    }
  };

  await walk(bounds, 0);
  return {
    rows: [...rowById.values()],
    requestCount: stats.requestCount,
    totalRowsFromBackend: stats.totalRowsFromBackend,
    splitQueryCount: stats.splitQueryCount,
    unresolvedTruncatedQueries: stats.unresolvedTruncatedQueries
  };
}

async function loadStreetAttributesFromLocalBackend(boundsOverride = null, polygonLayer = null) {
  const bounds = boundsOverride || map.getBounds();
  const queryResult = await collectStreetElementsFromLocalBackend(bounds, {
    maxDepth: LOCAL_STREET_BACKEND_QUERY_TILE_MAX_DEPTH,
    onProgress: (requestCount, depth) => {
      if (requestCount === 1 || requestCount % 3 === 0) {
        const splitHint = depth > 0 ? " (tiling around dense area)" : "";
        updateStreetLoadStatus(`Loading local backend streets... request ${requestCount.toLocaleString()}${splitHint}`);
      }
    }
  });
  const rows = queryResult.rows;
  if (!rows.length) {
    return {
      addedCount: 0,
      totalCount: streetAttributeById.size,
      candidateCount: 0,
      knownCount: 0,
      backendRequests: queryResult.requestCount,
      unresolvedTruncatedQueries: queryResult.unresolvedTruncatedQueries
    };
  }

  let addedCount = 0;
  let knownCount = 0;
  const batchSize = 1200;
  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    batch.forEach(raw => {
      const element = normalizeBackendStreetElement(raw);
      if (!element) return;
      if (!isMotorVehicleStreetTags(element.tags || {})) return;
      if (polygonLayer && !streetElementIntersectsPolygon(element, polygonLayer)) return;
      const normalizedTags = normalizeLocalStreetTags(element.tags || {});
      element.tags = normalizedTags;
      if (rowHasKnownStreetAttributes(normalizedTags)) knownCount += 1;
      if (upsertStreetElement(element)) addedCount += 1;
    });
    maybeEnableStreetCasingPerformanceMode();
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
    knownCount,
    backendRequests: queryResult.requestCount,
    unresolvedTruncatedQueries: queryResult.unresolvedTruncatedQueries
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
      if (!isMotorVehicleStreetTags(element.tags || {})) return;
      if (polygonLayer && !streetElementIntersectsPolygon(element, polygonLayer)) return;
      const normalizedTags = normalizeLocalStreetTags(element.tags || {});
      element.tags = normalizedTags;
      if (rowHasKnownStreetAttributes(normalizedTags)) knownCount += 1;
      if (upsertStreetElement(element)) addedCount += 1;
    });

    maybeEnableStreetCasingPerformanceMode();
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
    alert("Start backend from the Street Backend Manager app, or run the setup package.");
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

  clearRenderedStreetSegmentState();

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
    // Start with street network layer hidden on each app load.
    layerToggle.checked = false;
    storageSet(STREET_NETWORK_LAYER_VISIBLE_KEY, "off");
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
      updateLocalStreetSourceStatus("Street source is not ready. Open Street Setup Wizard, run setup + Backend Manager, then follow steps 1-3.");
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
let streetCasingPerformanceModeEnabled = false;
let streetCasingPerformanceModeNotified = false;
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
  syncAttributeSortControls();
  if (typeof window.__refreshSelectionToolsUi === "function") {
    window.__refreshSelectionToolsUi();
  }
}

function isStreetCasingEnabledForRendering() {
  return !streetCasingPerformanceModeEnabled;
}

function resetStreetRenderingPerformanceMode() {
  streetCasingPerformanceModeEnabled = false;
  streetCasingPerformanceModeNotified = false;
}

function removeStreetEntryCasingLayer(entry) {
  if (!entry?.casingLayer) return false;
  try {
    streetAttributeLayerGroup.removeLayer(entry.casingLayer);
  } catch {}
  try {
    entry.casingLayer.off?.();
  } catch {}
  entry.casingLayer = null;
  return true;
}

function maybeEnableStreetCasingPerformanceMode() {
  if (streetCasingPerformanceModeEnabled) return false;
  if (streetAttributeById.size < STREET_CASING_AUTO_DISABLE_SEGMENT_COUNT) return false;

  streetCasingPerformanceModeEnabled = true;
  let removed = 0;
  streetAttributeById.forEach(entry => {
    if (removeStreetEntryCasingLayer(entry)) removed += 1;
  });

  if (removed > 0) {
    applyStreetSelectionStyles();
    if (!streetCasingPerformanceModeNotified) {
      streetCasingPerformanceModeNotified = true;
      updateStreetLoadStatus(
        `Performance mode enabled for large street layers (${streetAttributeById.size.toLocaleString()} segments).`,
        false
      );
    }
  }
  return removed > 0;
}

function handleStreetSegmentLayerClick(id) {
  if (!isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) return;
  const attributePanel = document.getElementById("attributeTablePanel");
  if (attributeTableMode !== "streets") {
    setAttributeTableMode("streets");
  }
  if (attributePanel?.classList.contains("closed")) {
    openAttributePanel();
  }
  toggleStreetSegmentSelection(id, null, true);
}

function bindStreetSegmentLayerClick(layer, id) {
  if (!layer || typeof layer.on !== "function") return;
  layer.on("click", () => {
    handleStreetSegmentLayerClick(id);
  });
}

function setStreetSegmentStyle(entry, selected) {
  if (!entry?.layer) return;
  const baseStyle = getStreetSegmentBaseStyle(entry);
  if (selected) {
    const hasCasing = !!entry.casingLayer;
    const selectedWeight = Math.max((Number(baseStyle.weight) || STREET_BASE_LINE_WEIGHT) + (hasCasing ? 0.9 : 1.8), 5);
    const selectedCasingWeight = Math.max(
      (Number(baseStyle.casingWeight) || selectedWeight + 1),
      selectedWeight + 1
    );
    if (hasCasing && entry.casingLayer?.setStyle) {
      entry.casingLayer.setStyle({
        color: STREET_SELECTION_OUTER_COLOR,
        weight: selectedCasingWeight,
        opacity: 0.99
      });
    }
    entry.layer.setStyle({
      color: STREET_SELECTION_INNER_COLOR,
      weight: selectedWeight,
      opacity: 0.99
    });
    return;
  }

  if (entry.casingLayer?.setStyle) {
    entry.casingLayer.setStyle({
      color: baseStyle.casingColor,
      weight: baseStyle.casingWeight,
      opacity: baseStyle.casingOpacity
    });
  }
  entry.layer.setStyle({
    color: baseStyle.color,
    weight: baseStyle.weight,
    opacity: baseStyle.opacity
  });
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
    syncAttributeSortControls();
    return;
  }

  const headers = STREET_ATTRIBUTE_TABLE_HEADERS;
  const labels = STREET_ATTRIBUTE_TABLE_LABELS;
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
    status.textContent = `${streetAttributeSelectedIds.size} selected \u2022 ${rows.length} visible (${scopeLabel})`;
  }

  if (!rows.length) {
    table.innerHTML = "";
    empty.textContent = attributeState.selectedOnly
      ? "No selected street segments. Use selection tools or click street lines on the map."
      : "No street segments match the current filter.";
    empty.style.display = "block";
    syncAttributeSortControls();
    return;
  }
  empty.style.display = "none";

  const sortIndicator = key => (attributeState.sortKey === key ? (attributeState.sortDir > 0 ? " \u25B2" : " \u25BC") : "");
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
      syncAttributeSortControls();
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
  way["highway"]["highway"!~"^(footway|cycleway|path|steps|pedestrian|bridleway|corridor|platform|via_ferrata|proposed|construction)$"](${s},${w},${n},${e});
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

    if (!isMotorVehicleStreetTags(tags)) return;

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
          if (!e || e.type !== "way" || !isMotorVehicleStreetTags(e.tags || {}) || !Array.isArray(e.geom) || e.geom.length < 2) return;
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
    func_class: tags.func_class,
    ref: tags.ref,
    maxspeed: tags.maxspeed,
    lanes: tags.lanes,
    surface: tags.surface,
    oneway: tags.oneway
  };
  const existing = streetAttributeById.get(id);
  if (existing?.layer) {
    existing.row = mergeStreetAttributeRows(existing.row, incomingRow);
    if (!isStreetCasingEnabledForRendering() && existing.casingLayer) {
      removeStreetEntryCasingLayer(existing);
    }
    if (typeof existing.casingLayer?.setLatLngs === "function") existing.casingLayer.setLatLngs(latlngs);
    if (typeof existing.layer.setLatLngs === "function") existing.layer.setLatLngs(latlngs);
    setStreetSegmentStyle(existing, streetAttributeSelectedIds.has(id));
    return false;
  }
  const row = mergeStreetAttributeRows(null, incomingRow);
  const baseStyle = getStreetSegmentBaseStyle({ row });
  const useCasing = isStreetCasingEnabledForRendering();
  const casingLayer = useCasing
    ? L.polyline(latlngs, {
      color: baseStyle.casingColor,
      weight: baseStyle.casingWeight,
      opacity: baseStyle.casingOpacity,
      renderer: streetCasingCanvasRenderer,
      smoothFactor: 1.2
    })
    : null;
  const layer = L.polyline(latlngs, {
    color: baseStyle.color,
    weight: baseStyle.weight,
    opacity: baseStyle.opacity,
    renderer: streetCenterCanvasRenderer,
    smoothFactor: 1.2
  });
  bindStreetSegmentLayerClick(casingLayer, id);
  bindStreetSegmentLayerClick(layer, id);
  if (casingLayer) streetAttributeLayerGroup.addLayer(casingLayer);
  streetAttributeLayerGroup.addLayer(layer);
  streetAttributeById.set(id, { id, row, layer, casingLayer });
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
        const {
          addedCount,
          totalCount,
          candidateCount,
          knownCount,
          backendRequests,
          unresolvedTruncatedQueries
        } = backendResult;
        if (!candidateCount) {
          updateStreetLoadStatus("No backend street segments found in this area.", true);
        } else if (unresolvedTruncatedQueries > 0) {
          updateStreetLoadStatus(
            `Loaded ${addedCount} backend street segments (${totalCount} total on map) after ${backendRequests.toLocaleString()} requests, but dense areas may still be incomplete. Zoom in and reload to fill missing portions.`,
            true
          );
        } else if (!totalCount) {
          updateStreetLoadStatus("Backend streets loaded, but none intersect this polygon.", true);
        } else if (!knownCount) {
          updateStreetLoadStatus("Street geometry loaded, but local index attributes are empty. Re-run indexer with the latest converter to populate road fields.", true);
        } else {
          const requestSuffix = backendRequests > 1 ? ` in ${backendRequests.toLocaleString()} requests` : "";
          updateStreetLoadStatus(`Loaded ${addedCount} backend street segments (${totalCount} total on map)${requestSuffix}.`, false);
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
        if (!isMotorVehicleStreetTags(e?.tags || {})) return;
        if (polygonLayer && !streetElementIntersectsPolygon(e, polygonLayer)) return;
        if (upsertStreetElement(e)) chunkAdded += 1;
      });
      maybeEnableStreetCasingPerformanceMode();
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
        if (!isMotorVehicleStreetTags(e?.tags || {})) return;
        if (polygonLayer && !streetElementIntersectsPolygon(e, polygonLayer)) return;
        if (upsertStreetElement(e)) addedCount += 1;
      });
      maybeEnableStreetCasingPerformanceMode();
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
  syncAttributeSortControls();
}

const MAP_SELECTION_DRAW_MODE_KEY = "mapSelectionDrawMode";
const MAP_SELECTION_TARGET_KEY = "mapSelectionTarget";

function normalizeMapSelectionDrawMode(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "add" || normalized === "subtract" || normalized === "intersect") return normalized;
  return "replace";
}

function normalizeMapSelectionTarget(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "records" || normalized === "streets" || normalized === "both" || normalized === "active") {
    return normalized;
  }
  return "both";
}

function resolveMapSelectionTarget(targetValue = mapSelectionTarget) {
  const normalized = normalizeMapSelectionTarget(targetValue);
  if (normalized === "active") {
    return attributeTableMode === "streets" ? "streets" : "records";
  }
  return normalized;
}

function getMapSelectionModeLabel(modeValue = "replace") {
  const mode = normalizeMapSelectionDrawMode(modeValue);
  if (mode === "add") return "Add";
  if (mode === "subtract") return "Subtract";
  if (mode === "intersect") return "Intersect";
  return "Replace";
}

function getMapSelectionTargetLabel(targetValue = mapSelectionTarget) {
  const normalized = normalizeMapSelectionTarget(targetValue);
  const resolved = resolveMapSelectionTarget(normalized);
  if (normalized === "active") {
    return resolved === "streets" ? "Active Table (Street Segments)" : "Active Table (Record Stops)";
  }
  if (resolved === "streets") return "Street Segments";
  if (resolved === "records") return "Record Stops";
  return "Record Stops + Street Segments";
}

let mapSelectionDrawMode = normalizeMapSelectionDrawMode(storageGet(MAP_SELECTION_DRAW_MODE_KEY) || "replace");
let mapSelectionTarget = normalizeMapSelectionTarget(storageGet(MAP_SELECTION_TARGET_KEY) || "both");

function setMapSelectionDrawMode(modeValue, persist = true) {
  mapSelectionDrawMode = normalizeMapSelectionDrawMode(modeValue);
  if (persist) {
    storageSet(MAP_SELECTION_DRAW_MODE_KEY, mapSelectionDrawMode);
  }
  if (typeof window.__refreshSelectionToolsUi === "function") {
    window.__refreshSelectionToolsUi();
  }
}

function setMapSelectionTarget(targetValue, persist = true) {
  mapSelectionTarget = normalizeMapSelectionTarget(targetValue);
  if (persist) {
    storageSet(MAP_SELECTION_TARGET_KEY, mapSelectionTarget);
  }
  if (typeof window.__refreshSelectionToolsUi === "function") {
    window.__refreshSelectionToolsUi();
  }
}

function getCurrentAttributeModeSelectionCount() {
  return attributeTableMode === "streets"
    ? streetAttributeSelectedIds.size
    : attributeState.selectedRowIds.size;
}

function getSelectionCountForTarget(targetValue = mapSelectionTarget) {
  const resolved = resolveMapSelectionTarget(targetValue);
  if (resolved === "streets") return streetAttributeSelectedIds.size;
  if (resolved === "records") return attributeState.selectedRowIds.size;
  return attributeState.selectedRowIds.size + streetAttributeSelectedIds.size;
}

function getSelectionSummaryForTarget(targetValue = mapSelectionTarget) {
  const resolved = resolveMapSelectionTarget(targetValue);
  const recordCount = attributeState.selectedRowIds.size;
  const streetCount = streetAttributeSelectedIds.size;
  if (resolved === "streets") return `${streetCount.toLocaleString()} street segments selected`;
  if (resolved === "records") return `${recordCount.toLocaleString()} record stops selected`;
  return `${recordCount.toLocaleString()} record stops + ${streetCount.toLocaleString()} street segments selected`;
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

function splitSelectionCandidateIds(candidateIds) {
  const isScopedObject =
    candidateIds &&
    typeof candidateIds === "object" &&
    !Array.isArray(candidateIds) &&
    (
      Object.prototype.hasOwnProperty.call(candidateIds, "records") ||
      Object.prototype.hasOwnProperty.call(candidateIds, "streets")
    );

  if (!isScopedObject) {
    return { records: candidateIds, streets: candidateIds };
  }

  return {
    records: candidateIds.records,
    streets: candidateIds.streets
  };
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

function applySelectionIdsToActiveMode(candidateIds, modeValue = mapSelectionDrawMode, targetValue = mapSelectionTarget) {
  const mode = normalizeMapSelectionDrawMode(modeValue);
  const resolvedTarget = resolveMapSelectionTarget(targetValue);
  const scopedCandidates = splitSelectionCandidateIds(candidateIds);
  let recordsChanged = false;
  let streetsChanged = false;

  if (resolvedTarget === "records" || resolvedTarget === "both") {
    const selectableIds = getSelectableRecordRowIds();
    const base = new Set(
      [...attributeState.selectedRowIds].filter(id => selectableIds.has(id))
    );
    const next = combineSelectionIds(base, scopedCandidates.records, mode);
    recordsChanged = !areNumericIdSetsEqual(attributeState.selectedRowIds, next);
    attributeState.selectedRowIds = next;
    applyAttributeSelectionStyles();
  }

  if (resolvedTarget === "streets" || resolvedTarget === "both") {
    if (isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) {
      const next = combineSelectionIds(streetAttributeSelectedIds, scopedCandidates.streets, mode);
      streetsChanged = !areNumericIdSetsEqual(streetAttributeSelectedIds, next);
      if (streetsChanged) {
        replaceStreetSelection(next);
      }
      applyStreetSelectionStyles();
    }
  }

  syncSelectedStopsHeaderCount(getCurrentAttributeModeSelectionCount());
  refreshAttributeStatus();

  if (
    (attributeTableMode === "records" && recordsChanged) ||
    (attributeTableMode === "streets" && streetsChanged)
  ) {
    renderAttributeTable();
  }

  return getSelectionCountForTarget(resolvedTarget);
}

function invertVisibleSelectionInActiveMode(targetValue = mapSelectionTarget) {
  const resolvedTarget = resolveMapSelectionTarget(targetValue);
  let recordsChanged = false;
  let streetsChanged = false;

  if (resolvedTarget === "records" || resolvedTarget === "both") {
    const next = new Set(attributeState.selectedRowIds);
    const visibleIds = getVisibleSelectableRecordRowIds();
    visibleIds.forEach(id => {
      if (next.has(id)) next.delete(id);
      else next.add(id);
    });
    recordsChanged = !areNumericIdSetsEqual(attributeState.selectedRowIds, next);
    attributeState.selectedRowIds = next;
    applyAttributeSelectionStyles();
  }

  if (resolvedTarget === "streets" || resolvedTarget === "both") {
    if (isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY)) {
      const next = new Set(streetAttributeSelectedIds);
      const visibleIds = getVisibleSelectableStreetIds();
      visibleIds.forEach(id => {
        if (next.has(id)) next.delete(id);
        else next.add(id);
      });
      streetsChanged = !areNumericIdSetsEqual(streetAttributeSelectedIds, next);
      if (streetsChanged) {
        replaceStreetSelection(next);
      }
      applyStreetSelectionStyles();
    }
  }

  syncSelectedStopsHeaderCount(getCurrentAttributeModeSelectionCount());
  refreshAttributeStatus();

  if (
    (attributeTableMode === "records" && recordsChanged) ||
    (attributeTableMode === "streets" && streetsChanged)
  ) {
    renderAttributeTable();
  }

  return getSelectionCountForTarget(resolvedTarget);
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
function applyPolygonSelectionToStreets(polygonLayer) {
  const streetSelectable = isLayerManagerEntrySelectable(LAYER_MANAGER_STREET_KEY);
  if (!streetSelectable) {
    const hadSelection = streetAttributeSelectedIds.size > 0;
    if (hadSelection) {
      streetAttributeSelectedIds.clear();
      applyStreetSelectionStyles();
    }
    return hadSelection;
  }

  let nextSelectedStreetIds = new Set(streetAttributeSelectedIds);
  if (polygonLayer) {
    const candidateStreetIds = new Set();
    streetAttributeById.forEach((entry, id) => {
      if (!entry?.layer) return;
      if (!map.hasLayer(entry.layer)) return;
      if (!streetEntryIntersectsSelectionPolygon(entry, polygonLayer)) return;
      candidateStreetIds.add(id);
    });
    nextSelectedStreetIds = combineSelectionIds(nextSelectedStreetIds, candidateStreetIds, mapSelectionDrawMode);
  }

  nextSelectedStreetIds = new Set(
    [...nextSelectedStreetIds].filter(id => streetAttributeById.has(id))
  );

  const changed = !areNumericIdSetsEqual(streetAttributeSelectedIds, nextSelectedStreetIds);
  if (changed) {
    replaceStreetSelection(nextSelectedStreetIds);
    applyStreetSelectionStyles();
  }
  return changed;
}

function applyPolygonSelectionToRecords(polygonBounds) {
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

  const changed = !areNumericIdSetsEqual(attributeState.selectedRowIds, nextSelectedRowIds);
  attributeState.selectedRowIds = nextSelectedRowIds;
  return changed;
}

function updateSelectionCount() {
  const polygon = drawnLayer.getLayers()[0] || null;
  const polygonBounds = polygon?.getBounds?.() || null;
  const hasManualSelectionScope = !!selectedLayerKey || !!polygonBounds;
  const hasSelectedRecords = !!attributeState?.selectedRowIds?.size;
  const hasSelectedStreets = !!streetAttributeSelectedIds?.size;

  if (!hasManualSelectionScope && !hasSelectedRecords && !hasSelectedStreets) {
    syncSelectedStopsHeaderCount(getCurrentAttributeModeSelectionCount());
    refreshAttributeStatus();
    return;
  }

  const target = resolveMapSelectionTarget();
  let recordChanged = false;
  let streetChanged = false;

  if (target === "streets" || target === "both") {
    streetChanged = applyPolygonSelectionToStreets(polygon);
  }

  if (target === "records" || target === "both") {
    recordChanged = applyPolygonSelectionToRecords(polygonBounds);
  }

  syncSelectedStopsHeaderCount(getCurrentAttributeModeSelectionCount());
  refreshAttributeStatus();

  if (
    (attributeTableMode === "records" && recordChanged) ||
    (attributeTableMode === "streets" && streetChanged)
  ) {
    renderAttributeTable();
  }
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
    clearRenderedStreetSegmentState();
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
  updateUndoButtonState();   // ðŸ”¥ ADD THIS
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
const ROUTE_SEQUENCE_LAYER_VISIBLE_KEY = "routeSequenceLayerVisible";
let layerManagerOrderTop = [];
let layerManagerSelectableState = {};
const sequenceLayerByRouteDay = new Map();
const sequenceLayerState = {
  routeDayCount: 0,
  segmentCount: 0,
  arrowCount: 0,
  lastRunAt: 0
};
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

function resetLayerManagerOrderToDefault() {
  layerManagerOrderTop = getLayerManagerDefaultOrder();
  saveLayerManagerOrder(layerManagerOrderTop);
  return layerManagerOrderTop;
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

function applyLayerManagerPaneOrder(orderTop = ensureLayerManagerOrder()) {
  const streetOnTop = orderTop[0] === LAYER_MANAGER_STREET_KEY;
  const routeZ = streetOnTop ? 370 : 380;
  const streetZ = streetOnTop ? 380 : 370;
  const sequenceZ = routeZ - 1;

  const routePaneEl = map.getPane(ROUTE_DAY_PANE);
  if (routePaneEl) routePaneEl.style.zIndex = String(routeZ);

  const streetPaneEl = map.getPane(STREET_NETWORK_PANE);
  if (streetPaneEl) streetPaneEl.style.zIndex = String(streetZ);

  const sequencePaneEl = map.getPane(SEQUENCE_PANE);
  if (sequencePaneEl) sequencePaneEl.style.zIndex = String(sequenceZ);
}

function applyLayerManagerOrder() {
  const orderTop = ensureLayerManagerOrder();
  applyLayerManagerPaneOrder(orderTop);
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
  syncSequenceLayerVisibilityOnMap();
}

function setRouteDayLayerVisibilityFromManager(key, visible) {
  const group = routeDayGroups[key];
  if (!group || !Array.isArray(group.layers)) return;
  const nextVisible = !!visible;
  layerVisibilityState[key] = nextVisible;

  const routeDayCheckbox = [...document.querySelectorAll("#routeDayLayers input[data-key]")]
    .find(node => node.dataset.key === key);
  if (routeDayCheckbox) routeDayCheckbox.checked = nextVisible;
  applyFilters();
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

function normalizeRouteSequencerSequenceIncrement(value, fallback = 1) {
  const parsed = Math.round(Number(value));
  if (!Number.isFinite(parsed) || parsed < 1) return Math.max(1, Math.round(Number(fallback) || 1));
  return Math.min(100000, parsed);
}

function toRouteSequencerSequenceValue(sequenceOrder, sequenceIncrement) {
  const order = Math.max(1, Math.round(Number(sequenceOrder) || 1));
  const increment = normalizeRouteSequencerSequenceIncrement(sequenceIncrement, 1);
  return order * increment;
}

function getSequenceLayerLineColor(routeDayKey) {
  const symbol = symbolMap[routeDayKey] || getSymbol(routeDayKey);
  return String(symbol?.color || "#32d0ff");
}

function computeSequenceBearingDegrees(fromLatLng, toLatLng) {
  if (!fromLatLng || !toLatLng) return 0;
  const lat1 = (Number(fromLatLng.lat) * Math.PI) / 180;
  const lat2 = (Number(toLatLng.lat) * Math.PI) / 180;
  const dLon = ((Number(toLatLng.lng) - Number(fromLatLng.lng)) * Math.PI) / 180;
  const y = Math.sin(dLon) * Math.cos(lat2);
  const x = (Math.cos(lat1) * Math.sin(lat2))
    - (Math.sin(lat1) * Math.cos(lat2) * Math.cos(dLon));
  const bearing = (Math.atan2(y, x) * 180) / Math.PI;
  return (bearing + 360) % 360;
}

function buildSequenceArrowIcon(angleDeg, colorValue) {
  const color = String(colorValue || "#32d0ff");
  const safeAngle = Number.isFinite(Number(angleDeg)) ? Number(angleDeg) : 0;
  return L.divIcon({
    className: "sequence-arrow-marker",
    html: `<span class="sequence-arrow-head" style="--sequence-arrow-color:${color};transform:rotate(${safeAngle.toFixed(1)}deg);"></span>`,
    iconSize: [14, 14],
    iconAnchor: [7, 7]
  });
}

function clearSequenceLayerData(options = {}) {
  sequenceLayerByRouteDay.forEach(layerGroup => {
    try { map.removeLayer(layerGroup); } catch (_) {}
  });
  sequenceLayerByRouteDay.clear();
  sequenceLayerState.routeDayCount = 0;
  sequenceLayerState.segmentCount = 0;
  sequenceLayerState.arrowCount = 0;
  sequenceLayerState.lastRunAt = 0;
  if (!options || options.silent !== true) {
    updateSequenceLayerStatus();
  }
}

function syncSequenceLayerVisibilityOnMap() {
  const toggleNode = document.getElementById("sequenceLayerEnabled");
  const enabled = !!toggleNode?.checked;

  sequenceLayerByRouteDay.forEach((layerGroup, routeDayKey) => {
    const hasRouteDayGroup = Object.prototype.hasOwnProperty.call(routeDayGroups, routeDayKey);
    const routeVisible = hasRouteDayGroup && isRouteDayLayerVisibleOnMap(routeDayKey);
    const shouldShow = enabled && routeVisible;
    if (shouldShow) {
      if (!map.hasLayer(layerGroup)) layerGroup.addTo(map);
    } else if (map.hasLayer(layerGroup)) {
      map.removeLayer(layerGroup);
    }
  });

  updateSequenceLayerStatus();
}

function updateSequenceLayerStatus(message = "") {
  const statusNode = document.getElementById("sequenceLayerStatus");
  if (!statusNode) return;
  if (message) {
    statusNode.textContent = String(message);
    return;
  }

  const enabled = !!document.getElementById("sequenceLayerEnabled")?.checked;
  const routeCount = Number(sequenceLayerState.routeDayCount) || 0;
  const segmentCount = Number(sequenceLayerState.segmentCount) || 0;
  const arrowCount = Number(sequenceLayerState.arrowCount) || 0;

  if (!routeCount || !segmentCount) {
    statusNode.textContent = enabled
      ? "Run Sequence to generate route arrows."
      : "Layer is off.";
    return;
  }

  if (!enabled) {
    statusNode.textContent = `Layer is off. ${routeCount.toLocaleString()} route/day sequence(s) are ready.`;
    return;
  }

  statusNode.textContent = `Showing ${routeCount.toLocaleString()} route/day sequence(s), ${segmentCount.toLocaleString()} segment(s), ${arrowCount.toLocaleString()} arrow(s).`;
}

function rebuildSequenceLayerFromPlans(plans) {
  clearSequenceLayerData({ silent: true });

  const entries = Array.isArray(plans) ? plans : [];
  if (!entries.length) {
    updateSequenceLayerStatus("No sequence plans to draw.");
    return;
  }

  const maxArrowCount = 7000;
  let segmentCount = 0;
  let arrowCount = 0;

  entries.forEach(plan => {
    const routeDayKey = String(plan?.routeDayKey || "");
    if (!routeDayKey || !Object.prototype.hasOwnProperty.call(routeDayGroups, routeDayKey)) return;
    const visits = (Array.isArray(plan?.stopVisits) ? plan.stopVisits : [])
      .slice()
      .sort((a, b) => Number(a?.sequence || 0) - Number(b?.sequence || 0));
    if (visits.length < 2) return;

    const sourceLayers = Array.isArray(routeDayGroups[routeDayKey]?.layers)
      ? routeDayGroups[routeDayKey].layers
      : [];
    if (!sourceLayers.length) return;

    const latLngByRowId = new Map();
    sourceLayers.forEach(layer => {
      const rowId = Number(layer?._rowId);
      if (!Number.isFinite(rowId)) return;
      const ll = getLayerLatLng(layer);
      if (!ll) return;
      latLngByRowId.set(rowId, ll);
    });
    if (!latLngByRowId.size) return;

    const eventPoints = (Array.isArray(plan?.events) ? plan.events : [])
      .filter(event => Number.isFinite(Number(event?.lat)) && Number.isFinite(Number(event?.lon)))
      .map((event, index) => ({
        lat: Number(event.lat),
        lng: Number(event.lon),
        label: String(event?.name || event?.type || `Step ${index + 1}`),
        type: String(event?.type || ""),
        sequence: index + 1
      }));

    let pathPoints = eventPoints;
    if (pathPoints.length < 2) {
      const visits = (Array.isArray(plan?.stopVisits) ? plan.stopVisits : [])
        .slice()
        .sort((a, b) => Number(a?.sequence || 0) - Number(b?.sequence || 0))
        .map(visit => {
          const rowId = Number(visit?.rowId);
          const ll = latLngByRowId.get(rowId);
          if (!ll) return null;
          return {
            lat: Number(ll.lat),
            lng: Number(ll.lng),
            label: `Stop ${Number(visit?.sequence) || ""}`.trim(),
            type: "stop",
            sequence: Number(visit?.sequence) || 0
          };
        })
        .filter(Boolean);
      pathPoints = visits;
    }
    if (pathPoints.length < 2) return;

    const lineColor = getSequenceLayerLineColor(routeDayKey);
    const layerGroup = L.layerGroup();

    for (let i = 0; i < pathPoints.length - 1; i += 1) {
      const fromPoint = pathPoints[i];
      const toPoint = pathPoints[i + 1];
      const fromLatLng = { lat: Number(fromPoint.lat), lng: Number(fromPoint.lng) };
      const toLatLng = { lat: Number(toPoint.lat), lng: Number(toPoint.lng) };
      if (!Number.isFinite(fromLatLng.lat) || !Number.isFinite(fromLatLng.lng) || !Number.isFinite(toLatLng.lat) || !Number.isFinite(toLatLng.lng)) {
        continue;
      }
      const samePoint = Math.abs(fromLatLng.lat - toLatLng.lat) < 1e-9 && Math.abs(fromLatLng.lng - toLatLng.lng) < 1e-9;
      if (samePoint) continue;

      const line = L.polyline([fromLatLng, toLatLng], {
        pane: SEQUENCE_PANE,
        color: lineColor,
        weight: 2.5,
        opacity: 0.93,
        dashArray: "9 5",
        lineCap: "round",
        lineJoin: "round",
        interactive: false
      });
      line._sequenceRouteDayKey = routeDayKey;
      layerGroup.addLayer(line);
      segmentCount += 1;

      if (arrowCount >= maxArrowCount) continue;
      const midLat = (fromLatLng.lat + toLatLng.lat) / 2;
      const midLng = (fromLatLng.lng + toLatLng.lng) / 2;
      const bearing = computeSequenceBearingDegrees(fromLatLng, toLatLng);
      const arrowMarker = L.marker([midLat, midLng], {
        pane: SEQUENCE_PANE,
        icon: buildSequenceArrowIcon(bearing, lineColor),
        interactive: false,
        keyboard: false
      });
      arrowMarker._sequenceRouteDayKey = routeDayKey;
      layerGroup.addLayer(arrowMarker);
      arrowCount += 1;
    }

    if (layerGroup.getLayers().length) {
      sequenceLayerByRouteDay.set(routeDayKey, layerGroup);
    }
  });

  sequenceLayerState.routeDayCount = sequenceLayerByRouteDay.size;
  sequenceLayerState.segmentCount = segmentCount;
  sequenceLayerState.arrowCount = arrowCount;
  sequenceLayerState.lastRunAt = Date.now();

  syncSequenceLayerVisibilityOnMap();
}

function setSequenceLayerEnabled(enabled, options = {}) {
  const opts = options && typeof options === "object" ? options : {};
  const persist = opts.persist !== false;
  const nextEnabled = !!enabled;
  const toggleNode = document.getElementById("sequenceLayerEnabled");
  if (toggleNode && toggleNode.checked !== nextEnabled) {
    toggleNode.checked = nextEnabled;
  }
  if (persist) {
    storageSet(ROUTE_SEQUENCE_LAYER_VISIBLE_KEY, nextEnabled ? "on" : "off");
  }
  syncSequenceLayerVisibilityOnMap();
}

function initSequenceLayerControls() {
  const toggleHeader = document.getElementById("sequenceLayerToggleHeader");
  const content = document.getElementById("sequenceLayerContent");
  const masterToggle = document.getElementById("sequenceLayerEnabled");
  const statusNode = document.getElementById("sequenceLayerStatus");
  if (!toggleHeader || !content || !masterToggle || !statusNode) return;
  if (toggleHeader.dataset.sequenceLayerBound === "1") return;
  toggleHeader.dataset.sequenceLayerBound = "1";

  content.classList.add("collapsed");
  toggleHeader.classList.remove("open");

  toggleHeader.addEventListener("click", event => {
    if (event.target === masterToggle) return;
    const collapsed = content.classList.toggle("collapsed");
    toggleHeader.classList.toggle("open", !collapsed);
  });

  masterToggle.addEventListener("change", () => {
    setSequenceLayerEnabled(masterToggle.checked);
  });

  const stored = storageGet(ROUTE_SEQUENCE_LAYER_VISIBLE_KEY);
  const shouldShow = stored == null ? true : stored === "on";
  masterToggle.checked = shouldShow;
  setSequenceLayerEnabled(shouldShow, { persist: false });
  updateSequenceLayerStatus();
}

function multiDaySortRank(dayToken) {
  const normalized = normalizeDayToken(dayToken);
  const n = Number(normalized);
  if (Number.isFinite(n)) return n;
  if (normalized === "delivered") return 99;
  return 80;
}

function multiDayShortLabel(dayToken) {
  const normalized = normalizeDayToken(dayToken);
  const shortByDay = {
    "1": "Mon",
    "2": "Tue",
    "3": "Wed",
    "4": "Thu",
    "5": "Fri",
    "6": "Sat",
    "7": "Sun",
    delivered: "Delivered"
  };
  if (Object.prototype.hasOwnProperty.call(shortByDay, normalized)) return shortByDay[normalized];
  if (!normalized) return "No Day";
  return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

function multiDayLongLabel(dayToken) {
  const normalized = normalizeDayToken(dayToken);
  const longByDay = {
    "1": "Monday",
    "2": "Tuesday",
    "3": "Wednesday",
    "4": "Thursday",
    "5": "Friday",
    "6": "Saturday",
    "7": "Sunday",
    delivered: "Delivered"
  };
  if (Object.prototype.hasOwnProperty.call(longByDay, normalized)) return longByDay[normalized];
  if (!normalized) return "No Day";
  return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

function escapeMultiDayLabelText(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function normalizeMultiDayCoordinateToken(value) {
  if (value == null) return "";

  if (typeof value === "string") {
    const raw = value.trim();
    if (!raw) return "";
    const parsed = Number(raw);
    if (!Number.isFinite(parsed)) return "";
    if (!/[eE]/.test(raw)) {
      let normalized = raw.replace(/^\+/, "");
      if (normalized.includes(".")) {
        normalized = normalized.replace(/0+$/, "").replace(/\.$/, "");
      }
      if (normalized === "-0") normalized = "0";
      return normalized;
    }
    return String(parsed);
  }

  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return "";
  if (Object.is(parsed, -0)) return "0";
  return String(parsed);
}

function makeMultiDayLocationKey(lat, lon) {
  const latToken = normalizeMultiDayCoordinateToken(lat);
  const lonToken = normalizeMultiDayCoordinateToken(lon);
  if (!latToken || !lonToken) return "";
  return `${latToken}|${lonToken}`;
}

function buildRecordPopupContent(layer) {
  const routeToken = String(layer?._routeToken ?? "").trim() || "-";
  const rawDay = String(layer?._dayToken ?? "").trim();
  const dayToken = normalizeDayToken(rawDay) || rawDay;
  const dayLabel = dayName(dayToken) || dayToken || "No Day";
  const serviceDetail = String(layer?._multiDayDetailLabel || "").trim();
  const serviceLine = serviceDetail
    ? `<div><strong>Service Frequency:</strong> ${escapeMultiDayLabelText(serviceDetail)}</div>`
    : "";
  const row = layer?._rowRef || {};
  const outputFields = getRouteSequencerActiveOutputFields();
  const seqValue = parseRouteSequencerNumeric(row?.[outputFields.sequence], NaN);
  const tripValue = parseRouteSequencerNumeric(row?.[outputFields.trip], NaN);
  const loadValue = parseRouteSequencerNumeric(row?.[outputFields.cumulativeLoad], NaN);
  const sequenceLine = Number.isFinite(seqValue) && seqValue > 0
    ? `<div><strong>Sequence:</strong> ${escapeMultiDayLabelText(String(Math.round(seqValue)))}</div>`
    : "";
  const tripLine = Number.isFinite(tripValue) && tripValue > 0
    ? `<div><strong>Trip:</strong> ${escapeMultiDayLabelText(String(Math.round(tripValue)))}</div>`
    : "";
  const loadLine = Number.isFinite(loadValue) && loadValue >= 0
    ? `<div><strong>Cumulative Load:</strong> ${escapeMultiDayLabelText(loadValue.toFixed(2))} tons</div>`
    : "";
  const segmentIdValue = row?.SEGMENT_ID ?? row?.Segment_ID ?? row?.segment_id ?? "";
  const segmentSideValue = row?.SEGMENT_SIDE ?? row?.Segment_Side ?? row?.segment_side ?? "";
  const matchStatusFieldValue = row?.Segment_Match_Status ?? row?.segment_match_status ?? row?.SEGMENT_MATCH_STATUS ?? "";
  const readyStatusFieldValue = row?.Sequence_Ready ?? row?.sequence_ready ?? row?.SEQUENCE_READY ?? "";
  const segmentMethodValue = row?.Segment_Match_Method ?? row?.segment_match_method ?? row?.SEGMENT_MATCH_METHOD ?? "";
  const hasSegmentMatch = String(segmentIdValue || "").trim() !== "";
  const matchStatusValue = String(matchStatusFieldValue || (hasSegmentMatch ? "Matched" : "Unmatched")).trim();
  const readyStatusValue = String(readyStatusFieldValue || (hasSegmentMatch ? "Ready" : "Not Ready")).trim();
  const segmentIdLine = String(segmentIdValue || "").trim()
    ? `<div><strong>Segment ID:</strong> ${escapeMultiDayLabelText(String(segmentIdValue))}</div>`
    : "";
  const segmentSideLine = String(segmentSideValue || "").trim()
    ? `<div><strong>Segment Side:</strong> ${escapeMultiDayLabelText(String(segmentSideValue))}</div>`
    : "";
  const segmentMatchStatusLine = matchStatusValue
    ? `<div><strong>Segment Match:</strong> ${escapeMultiDayLabelText(matchStatusValue)}</div>`
    : "";
  const sequenceReadyLine = readyStatusValue
    ? `<div><strong>Sequence Ready:</strong> ${escapeMultiDayLabelText(readyStatusValue)}</div>`
    : "";
  const segmentMethodLine = String(segmentMethodValue || "").trim()
    ? `<div><strong>Match Method:</strong> ${escapeMultiDayLabelText(String(segmentMethodValue))}</div>`
    : "";

  return `
    <div style="font-size:14px; line-height:1.4;">
      <div><strong>Route:</strong> ${escapeMultiDayLabelText(routeToken)}</div>
      <div><strong>Day:</strong> ${escapeMultiDayLabelText(dayLabel)}</div>
      ${sequenceLine}
      ${tripLine}
      ${loadLine}
      ${segmentIdLine}
      ${segmentSideLine}
      ${segmentMatchStatusLine}
      ${sequenceReadyLine}
      ${segmentMethodLine}
      ${serviceLine}
    </div>
  `;
}

function syncRecordPopupContent(layer) {
  if (!layer || typeof layer.bindPopup !== "function") return;
  const popupContent = buildRecordPopupContent(layer);
  if (typeof layer.getPopup === "function" && layer.getPopup() && typeof layer.setPopupContent === "function") {
    layer.setPopupContent(popupContent);
    return;
  }
  layer.bindPopup(popupContent);
}

function collectSortedMultiDayDayTokens(daySet) {
  const normalizedDays = [...new Set(
    [...(daySet || [])]
      .map(v => normalizeDayToken(v) || String(v ?? "").trim())
      .filter(Boolean)
  )];

  normalizedDays.sort((a, b) => {
    const aRank = multiDaySortRank(a);
    const bRank = multiDaySortRank(b);
    if (aRank !== bRank) return aRank - bRank;
    return String(a).localeCompare(String(b), undefined, { sensitivity: "base" });
  });
  return normalizedDays;
}

function buildMultiDayServiceSummary(daySet) {
  const normalizedDays = collectSortedMultiDayDayTokens(daySet);
  if (!normalizedDays.length) {
    return { short: "No Day", detail: "No Day", dayCount: 0 };
  }

  const shortDays = normalizedDays.map(multiDayShortLabel);
  const longDays = normalizedDays.map(multiDayLongLabel);
  const dayCount = shortDays.length;

  const short = dayCount === 1
    ? `${shortDays[0]} only`
    : `${shortDays.join("/")} (${dayCount}d)`;

  const detail = dayCount === 1
    ? `${longDays[0]} only`
    : `${longDays.join(", ")} (${dayCount} days)`;

  return { short, detail, dayCount, tokens: normalizedDays };
}

function buildMultiDayFrequencyKey(dayTokens) {
  const tokens = Array.isArray(dayTokens) ? dayTokens : [];
  return JSON.stringify(tokens);
}

function hasActiveMultiDayFrequencyFilter() {
  if (!multiDayFrequencyVisibilityByKey.size) return false;
  for (const visible of multiDayFrequencyVisibilityByKey.values()) {
    if (!visible) return true;
  }
  return false;
}

function hasActiveMultiDayLabelFilter() {
  if (multiDayLabelDayVisibilityByToken.size) {
    for (const visible of multiDayLabelDayVisibilityByToken.values()) {
      if (!visible) return true;
    }
  }
  if (multiDayLabelFrequencyVisibilityByKey.size) {
    for (const visible of multiDayLabelFrequencyVisibilityByKey.values()) {
      if (!visible) return true;
    }
  }
  return false;
}

function isMultiDayFrequencyVisibleByKey(frequencyKey) {
  if (!frequencyKey) return true;
  if (!multiDayFrequencyVisibilityByKey.size) return true;
  if (!multiDayFrequencyVisibilityByKey.has(frequencyKey)) return true;
  return multiDayFrequencyVisibilityByKey.get(frequencyKey) !== false;
}

function isMultiDayLabelFrequencyVisibleByKey(frequencyKey) {
  if (!frequencyKey) return true;
  if (!multiDayLabelFrequencyVisibilityByKey.size) return true;
  if (!multiDayLabelFrequencyVisibilityByKey.has(frequencyKey)) return true;
  return multiDayLabelFrequencyVisibilityByKey.get(frequencyKey) !== false;
}

function isMultiDayLabelDayVisibleByToken(dayToken) {
  const token = String(dayToken || "").trim();
  if (!token) return true;
  if (!multiDayLabelDayVisibilityByToken.size) return true;
  if (!multiDayLabelDayVisibilityByToken.has(token)) return true;
  return multiDayLabelDayVisibilityByToken.get(token) !== false;
}

function isMultiDayProfileVisibleByLabelFilters(profile) {
  if (!profile) return true;
  if (!isMultiDayLabelFrequencyVisibleByKey(profile.frequencyKey)) return false;
  const dayTokens = Array.isArray(profile.dayTokens) ? profile.dayTokens : [];
  for (const token of dayTokens) {
    if (!isMultiDayLabelDayVisibleByToken(token)) return false;
  }
  return true;
}

function isLayerVisibleByMultiDayFrequency(layer) {
  const frequencyKey = String(layer?._multiDayFrequencyKey || "");
  return isMultiDayFrequencyVisibleByKey(frequencyKey);
}

function syncMultiDayFrequencyVisibilityState() {
  const next = new Map();
  multiDayFrequencyCatalog.forEach((entry, key) => {
    const existing = multiDayFrequencyVisibilityByKey.has(key)
      ? multiDayFrequencyVisibilityByKey.get(key)
      : true;
    next.set(key, existing !== false);
  });
  multiDayFrequencyVisibilityByKey = next;
}

function syncMultiDayLabelVisibilityState() {
  const nextDayState = new Map();
  multiDayLabelDayCatalog.forEach((entry, token) => {
    const existing = multiDayLabelDayVisibilityByToken.has(token)
      ? multiDayLabelDayVisibilityByToken.get(token)
      : true;
    nextDayState.set(token, existing !== false);
  });
  multiDayLabelDayVisibilityByToken = nextDayState;

  const nextFrequencyState = new Map();
  multiDayFrequencyCatalog.forEach((entry, key) => {
    const existing = multiDayLabelFrequencyVisibilityByKey.has(key)
      ? multiDayLabelFrequencyVisibilityByKey.get(key)
      : true;
    nextFrequencyState.set(key, existing !== false);
  });
  multiDayLabelFrequencyVisibilityByKey = nextFrequencyState;
}

function setAllMultiDayLabelDayVisibility(visible) {
  const nextVisible = !!visible;
  multiDayLabelDayCatalog.forEach((entry, token) => {
    multiDayLabelDayVisibilityByToken.set(token, nextVisible);
  });
}

function setAllMultiDayLabelFrequencyVisibility(visible) {
  const nextVisible = !!visible;
  multiDayFrequencyCatalog.forEach((entry, key) => {
    multiDayLabelFrequencyVisibilityByKey.set(key, nextVisible);
  });
}

function setMultiDayLabelDayVisibility(token, visible) {
  const key = String(token || "").trim();
  if (!key || !multiDayLabelDayCatalog.has(key)) return false;
  const nextVisible = !!visible;
  if (multiDayLabelDayVisibilityByToken.get(key) === nextVisible) return false;
  multiDayLabelDayVisibilityByToken.set(key, nextVisible);
  return true;
}

function setMultiDayLabelFrequencyVisibility(frequencyKey, visible) {
  const key = String(frequencyKey || "");
  if (!key || !multiDayFrequencyCatalog.has(key)) return false;
  const nextVisible = !!visible;
  if (multiDayLabelFrequencyVisibilityByKey.get(key) === nextVisible) return false;
  multiDayLabelFrequencyVisibilityByKey.set(key, nextVisible);
  return true;
}

function setAllMultiDayFrequencyVisibility(visible) {
  const nextVisible = !!visible;
  multiDayFrequencyCatalog.forEach((entry, key) => {
    multiDayFrequencyVisibilityByKey.set(key, nextVisible);
  });
  syncMultiDayToggleUi();
}

function setMultiDayFrequencyVisibilityByPredicate(predicate) {
  multiDayFrequencyCatalog.forEach((entry, key) => {
    const keep = !!predicate(entry);
    multiDayFrequencyVisibilityByKey.set(key, keep);
  });
  syncMultiDayToggleUi();
}

function setMultiDayFrequencyVisibility(frequencyKey, visible) {
  const key = String(frequencyKey || "");
  if (!key || !multiDayFrequencyCatalog.has(key)) return false;
  const nextVisible = !!visible;
  if (multiDayFrequencyVisibilityByKey.get(key) === nextVisible) return false;
  multiDayFrequencyVisibilityByKey.set(key, nextVisible);
  syncMultiDayToggleUi();
  return true;
}

function getMultiDayFrequencyVisibilitySummary() {
  let totalFreq = 0;
  let visibleFreq = 0;
  let totalRecords = 0;
  let visibleRecords = 0;

  multiDayFrequencyCatalog.forEach(entry => {
    totalFreq += 1;
    totalRecords += Number(entry?.recordCount) || 0;
    if (isMultiDayFrequencyVisibleByKey(entry?.key)) {
      visibleFreq += 1;
      visibleRecords += Number(entry?.recordCount) || 0;
    }
  });
  return { totalFreq, visibleFreq, totalRecords, visibleRecords };
}

function refreshMultiDayManagerUi() {
  syncMultiDayToggleUi();
  if (typeof multiDayManagerUiRenderer === "function") {
    multiDayManagerUiRenderer();
  }
  if (typeof multiDaySidebarUiRenderer === "function") {
    multiDaySidebarUiRenderer();
  }
}

function rebuildMultiDayServiceProfilesFromMapLayers() {
  multiDayServiceByLocation.clear();
  multiDayFrequencyCatalog.clear();
  multiDayLabelDayCatalog.clear();

  Object.entries(routeDayGroups).forEach(([routeDayKey, group]) => {
    const [, rawDay = ""] = String(routeDayKey || "").split("|");
    const normalizedDay = normalizeDayToken(rawDay) || String(rawDay || "").trim() || "No Day";
    const layers = Array.isArray(group?.layers) ? group.layers : [];

    layers.forEach(layer => {
      const baseLat = Number(layer?._base?.lat);
      const baseLon = Number(layer?._base?.lon);
      const latlng = Number.isFinite(baseLat) && Number.isFinite(baseLon)
        ? L.latLng(baseLat, baseLon)
        : getLayerLatLng(layer);
      if (!latlng) return;

      const key = makeMultiDayLocationKey(
        layer?._multiDayLatToken ?? latlng.lat,
        layer?._multiDayLonToken ?? latlng.lng
      );
      if (!key) return;

      let profile = multiDayServiceByLocation.get(key);
      if (!profile) {
        profile = {
          key,
          latlng,
          days: new Set(),
          layers: []
        };
        multiDayServiceByLocation.set(key, profile);
      }

      profile.days.add(normalizedDay);
      profile.layers.push(layer);
      layer._multiDayLocationKey = key;
    });
  });

  multiDayServiceByLocation.forEach(profile => {
    const summary = buildMultiDayServiceSummary(profile.days);
    const frequencyKey = buildMultiDayFrequencyKey(summary.tokens);
    const dayTokens = Array.isArray(summary.tokens) ? summary.tokens : [];
    profile.shortLabel = summary.short;
    profile.detailLabel = summary.detail;
    profile.dayCount = summary.dayCount;
    profile.isMulti = summary.dayCount > 1;
    profile.frequencyKey = frequencyKey;
    profile.dayTokens = dayTokens;

    dayTokens.forEach(token => {
      const normalizedToken = String(token || "").trim();
      if (!normalizedToken) return;
      const existingDay = multiDayLabelDayCatalog.get(normalizedToken) || {
        token: normalizedToken,
        shortLabel: multiDayShortLabel(normalizedToken),
        detailLabel: multiDayLongLabel(normalizedToken),
        locationCount: 0,
        recordCount: 0
      };
      existingDay.locationCount += 1;
      existingDay.recordCount += profile.layers.length;
      multiDayLabelDayCatalog.set(normalizedToken, existingDay);
    });

    if (frequencyKey) {
      const existing = multiDayFrequencyCatalog.get(frequencyKey) || {
        key: frequencyKey,
        shortLabel: summary.short,
        detailLabel: summary.detail,
        dayCount: summary.dayCount,
        isMulti: summary.dayCount > 1,
        tokens: dayTokens,
        locationCount: 0,
        recordCount: 0
      };
      existing.locationCount += 1;
      existing.recordCount += profile.layers.length;
      multiDayFrequencyCatalog.set(frequencyKey, existing);
    }

    profile.layers.forEach(layer => {
      layer._multiDayShortLabel = profile.shortLabel;
      layer._multiDayDetailLabel = profile.detailLabel;
      layer._multiDayFrequencyKey = frequencyKey;
      syncRecordPopupContent(layer);
    });
  });

  syncMultiDayFrequencyVisibilityState();
  syncMultiDayLabelVisibilityState();
  refreshMultiDayManagerUi();
}

function syncMultiDayToggleUi() {
  if (!multiDayToggleButtonRef) return;
  const frequencyFilterActive = hasActiveMultiDayFrequencyFilter();
  const labelFilterActive = hasActiveMultiDayLabelFilter();
  const filterActive = frequencyFilterActive || labelFilterActive;
  const buttonActive = multiDayServiceEnabled || filterActive;
  multiDayToggleButtonRef.classList.toggle("active", buttonActive);
  multiDayToggleButtonRef.setAttribute("aria-pressed", multiDayServiceEnabled ? "true" : "false");
  if (multiDayServiceEnabled && filterActive) {
    multiDayToggleButtonRef.title = "Multi-Day labels are on and filters are active.";
  } else if (multiDayServiceEnabled) {
    multiDayToggleButtonRef.title = "Multi-Day labels are on.";
  } else if (filterActive) {
    multiDayToggleButtonRef.title = "Multi-Day filters are active.";
  } else {
    multiDayToggleButtonRef.title = "Open Multi-Day manager.";
  }
}

function renderMultiDayServiceLabels() {
  multiDayLabelLayerGroup.clearLayers();
  if (!multiDayServiceEnabled) return;
  if (!multiDayServiceByLocation.size) return;

  const bounds = map.getBounds();
  let labelCount = 0;
  const orderedProfiles = [...multiDayServiceByLocation.values()].sort((a, b) => {
    if (b.dayCount !== a.dayCount) return b.dayCount - a.dayCount;
    return String(a.shortLabel || "").localeCompare(String(b.shortLabel || ""), undefined, { sensitivity: "base" });
  });

  orderedProfiles.forEach(profile => {
    if (labelCount >= MULTI_DAY_LABEL_HARD_LIMIT) return;
    if (!isMultiDayProfileVisibleByLabelFilters(profile)) return;
    const visibleLayer = profile.layers.find(layer => map.hasLayer(layer));
    if (!visibleLayer) return;

    const latlng = getLayerLatLng(visibleLayer) || profile.latlng;
    if (!latlng || !bounds.contains(latlng)) return;
    const text = profile.shortLabel || "";
    if (!text) return;

    const icon = L.divIcon({
      className: "multiday-service-label-icon",
      html: `<span class="multiday-service-label-chip ${profile.isMulti ? "multi" : "single"}">${escapeMultiDayLabelText(text)}</span>`,
      iconSize: [0, 0],
      iconAnchor: [0, 0]
    });
    const marker = L.marker(latlng, {
      icon,
      interactive: false,
      keyboard: false,
      pane: MULTI_DAY_LABEL_PANE,
      zIndexOffset: 2000
    });
    multiDayLabelLayerGroup.addLayer(marker);
    labelCount += 1;
  });
}

function setMultiDayServiceEnabled(enabled) {
  multiDayServiceEnabled = !!enabled;
  syncMultiDayToggleUi();
  renderMultiDayServiceLabels();
  refreshMultiDayManagerUi();
}

// Used by summary interactions (and visualization popup) to focus the matching route+day on the map.
window.highlightRouteDayOnMap = function(routeValue, dayValue, options = {}) {
  const opts = options && typeof options === "object" ? options : {};
  const selectRecords = !!opts.selectRecords;
  const flash = Object.prototype.hasOwnProperty.call(opts, "flash")
    ? !!opts.flash
    : !selectRecords;
  const shouldFitBounds = opts.fitBounds !== false;

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

  const selectedRowIds = new Set();

  layerVisibilityState[matchingKey] = true;
  const layerCheckbox = document.querySelector(`input[data-key="${matchingKey}"]`);
  if (layerCheckbox) layerCheckbox.checked = true;

  const bounds = L.latLngBounds();
  markers.forEach(marker => {
    map.addLayer(marker);
    const ll = getLayerLatLng(marker);
    if (ll) bounds.extend(ll);

    const rowId = Number(marker?._rowId);
    if (selectRecords && Number.isFinite(rowId)) selectedRowIds.add(rowId);

    if (flash) {
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
    }
  });

  if (selectRecords && selectedRowIds.size && typeof window.setAttributeSelectedRowIds === "function") {
    window.setAttributeSelectedRowIds([...selectedRowIds]);
  }

  if (shouldFitBounds && bounds.isValid()) {
    map.fitBounds(bounds, { padding: [40, 40], maxZoom: 15 });
  }

  return {
    ok: true,
    message: `Highlighted ${matchingKey} on the map.`,
    key: matchingKey,
    points: markers.length,
    selectedCount: selectedRowIds.size
  };
};


let symbolIndex = 0;
let globalBounds = L.latLngBounds(); // used to zoom map to all points


// Convert day number â†’ day name
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
function tokenizeRouteCode(value) {
  const text = String(value ?? "").trim();
  if (!text) return [];
  return text.match(/[A-Za-z]+|\d+|[^A-Za-z\d]+/g) || [text];
}

function compareRouteCodes(a, b) {
  const left = String(a ?? "").trim();
  const right = String(b ?? "").trim();

  if (!left && !right) return 0;
  if (!left) return 1;
  if (!right) return -1;

  const leftTokens = tokenizeRouteCode(left);
  const rightTokens = tokenizeRouteCode(right);
  const maxLen = Math.max(leftTokens.length, rightTokens.length);

  for (let i = 0; i < maxLen; i += 1) {
    const leftToken = leftTokens[i];
    const rightToken = rightTokens[i];

    if (leftToken === undefined) return -1;
    if (rightToken === undefined) return 1;

    const leftIsNumber = /^\d+$/.test(leftToken);
    const rightIsNumber = /^\d+$/.test(rightToken);

    if (leftIsNumber && rightIsNumber) {
      const numberDiff = Number(leftToken) - Number(rightToken);
      if (numberDiff !== 0) return numberDiff;

      const widthDiff = leftToken.length - rightToken.length;
      if (widthDiff !== 0) return widthDiff;
      continue;
    }

    if (leftIsNumber !== rightIsNumber) {
      return leftIsNumber ? -1 : 1;
    }

    const tokenDiff = leftToken.localeCompare(rightToken, undefined, { sensitivity: "base" });
    if (tokenDiff !== 0) return tokenDiff;
  }

  return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
}

function buildRouteCheckboxes(routes) {
  const c = document.getElementById("routeCheckboxes");
  c.innerHTML = "";
  const sortedRoutes = [...routes].sort(compareRouteCodes);

  sortedRoutes.forEach(route => {
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
  applyFilters({ chunked: true });
}

document.getElementById("routesAll").onclick  = () => setCheckboxGroup("routeCheckboxes", true);
document.getElementById("routesNone").onclick = () => setCheckboxGroup("routeCheckboxes", false);
document.getElementById("daysAll").onclick    = () => setCheckboxGroup("dayCheckboxes", true);
document.getElementById("daysNone").onclick   = () => setCheckboxGroup("dayCheckboxes", false);



// ===== Route + Day ALL / NONE =====
function setRouteDayLayerBulkState(checked) {
  document.querySelectorAll("#routeDayLayers input[type='checkbox']").forEach(cb => {
    cb.checked = checked;
    const key = String(cb.dataset.key || "").trim();
    if (key) layerVisibilityState[key] = checked;
  });
  applyFilters({ chunked: true });
}

document.getElementById("routeDayAll").onclick  = () => setRouteDayLayerBulkState(true);
document.getElementById("routeDayNone").onclick = () => setRouteDayLayerBulkState(false);


// ================= APPLY MAP FILTERS =================
let applyFiltersRunId = 0;

function applyFilters(options = {}) {
  const { chunked = false } = options || {};

  const routeCheckboxes = [...document.querySelectorAll("#routeCheckboxes input")];
  const dayCheckboxes   = [...document.querySelectorAll("#dayCheckboxes input")];

  const routeSet = new Set(routeCheckboxes.filter(i => i.checked).map(i => String(i.value)));
  const daySet = new Set(
    dayCheckboxes
      .filter(i => i.checked)
      .map(i => normalizeDayToken(i.value) || String(i.value || "").trim())
      .filter(Boolean)
  );

  const entries = Object.entries(routeDayGroups);
  const totalLayerCount = entries.reduce((sum, [, group]) => sum + ((group.layers || []).length || 0), 0);
  const shouldChunk = chunked || totalLayerCount > 3000;
  const runId = ++applyFiltersRunId;

  const finalizeFilterPass = () => {
    if (runId !== applyFiltersRunId) return;
    if (shouldChunk) {
      const orderTop = ensureLayerManagerOrder();
      applyLayerManagerPaneOrder(orderTop);
      try { drawnLayer?.bringToFront?.(); } catch (_) {}
      try { streetLoadPolygonLayerGroup?.bringToFront?.(); } catch (_) {}
      syncSequenceLayerVisibilityOnMap();
    } else {
      applyLayerManagerOrder();
    }
    updateSelectionCount();
    updateStats();
    refreshLayerManagerUiIfOpen();
    refreshMultiDayManagerUi();
    renderMultiDayServiceLabels();
  };

  if (!shouldChunk) {
    entries.forEach(([key, group]) => {
      const [routeToken = "", rawDay = ""] = String(key || "").split("|");
      const routeVisible = routeSet.has(routeToken);
      const dayToken = normalizeDayToken(rawDay) || String(rawDay || "").trim();
      const dayVisible = daySet.has(dayToken);
      const layerToggleVisible = Object.prototype.hasOwnProperty.call(layerVisibilityState, key)
        ? !!layerVisibilityState[key]
        : true;

      (group.layers || []).forEach(layer => {
        const frequencyVisible = isLayerVisibleByMultiDayFrequency(layer);
        const show = routeVisible && dayVisible && layerToggleVisible && frequencyVisible;
        const isOnMap = layer && layer._map === map;
        if (show) {
          if (!isOnMap) layer.addTo(map);
        } else if (isOnMap) {
          map.removeLayer(layer);
        }
      });
    });

    finalizeFilterPass();
    return;
  }

  let groupIndex = 0;
  let layerIndex = 0;
  const frameBudgetMs = 8;

  const runChunk = () => {
    if (runId !== applyFiltersRunId) return;

    const frameStart = performance.now();

    while (groupIndex < entries.length && (performance.now() - frameStart) < frameBudgetMs) {
      const [key, group] = entries[groupIndex];
      const [routeToken = "", rawDay = ""] = String(key || "").split("|");
      const routeVisible = routeSet.has(routeToken);
      const dayToken = normalizeDayToken(rawDay) || String(rawDay || "").trim();
      const dayVisible = daySet.has(dayToken);
      const layerToggleVisible = Object.prototype.hasOwnProperty.call(layerVisibilityState, key)
        ? !!layerVisibilityState[key]
        : true;
      const layers = group.layers || [];

      while (layerIndex < layers.length && (performance.now() - frameStart) < frameBudgetMs) {
        const layer = layers[layerIndex++];
        const frequencyVisible = isLayerVisibleByMultiDayFrequency(layer);
        const show = routeVisible && dayVisible && layerToggleVisible && frequencyVisible;
        const isOnMap = layer && layer._map === map;
        if (show) {
          if (!isOnMap) layer.addTo(map);
        } else if (isOnMap) {
          map.removeLayer(layer);
        }
      }

      if (layerIndex >= layers.length) {
        groupIndex++;
        layerIndex = 0;
      }
    }

    if (groupIndex < entries.length) {
      requestAnimationFrame(runChunk);
      return;
    }

    finalizeFilterPass();
  };

  runChunk();
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
    li.textContent = `Route ${r} \u2013 ${dayName(d)}: ${visible}`;
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
    const routeCmp = compareRouteCodes(aRoute, bRoute);
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
      applyFilters({ chunked: true });
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
  { key: "UNIQUEID", label: "Unique ID", required: false },
  { key: "LATITUDE", label: "Latitude", required: true },
  { key: "LONGITUDE", label: "Longitude", required: true },
  { key: "ROUTE", label: "Route", required: false },
  { key: "DAY", label: "Day", required: false },
  { key: "SEQNO", label: "Sequence Number", required: false },
  { key: "FREQ", label: "Service Frequency", required: false },
  { key: "NEWROUTE", label: "New Route", required: false },
  { key: "NEWDAY", label: "New Day", required: false },
  { key: "NEWSEQ", label: "New Sequence", required: false },
  { key: "NEWFREQ", label: "New Frequency", required: false },
  { key: "del_status", label: "Delivery Status", required: false },
  { key: "CSADR#", label: "Street Number", required: false },
  { key: "CSSDIR", label: "Street Direction", required: false },
  { key: "CSSTRT", label: "Street Name", required: false },
  { key: "CSSFUX", label: "Street Suffix", required: false },
  { key: "SIZE", label: "Container Size", required: false },
  { key: "QTY", label: "Quantity", required: false },
  { key: "BINNO", label: "Bin Number", required: false },
  { key: "DEMAND", label: "Demand (Tons)", required: false },
  { key: "Service_Time", label: "Service Time (MM:SS)", required: false },
  { key: "DURATION_A", label: "Service Time (Minutes)", required: false },
  { key: "EARLIEST_A", label: "Earliest Arrival", required: false },
  { key: "LATEST_ARR", label: "Latest Arrival", required: false }
];

const COLUMN_MAPPING_ALIASES = {
  UNIQUEID: ["uniqueid", "unique_id", "id", "recordid", "stopid", "customerid"],
  LATITUDE: ["latitude", "lat", "y", "ycoord", "ycoordinate"],
  LONGITUDE: ["longitude", "lon", "lng", "long", "x", "xcoord", "xcoordinate"],
  NEWROUTE: ["newroute", "route", "routecode", "routeid", "rte"],
  NEWDAY: ["newday", "day", "weekday", "serviceday"],
  NEWSEQ: ["newseq", "newsequence", "newstopsequence", "neworder"],
  NEWFREQ: ["newfreq", "newfrequency", "newservicefrequency"],
  ROUTE: ["route", "routecode", "routeid", "rte"],
  DAY: ["day", "weekday", "serviceday", "routeday"],
  SEQNO: ["seqno", "sequence", "seq", "stopsequence", "stoporder", "order"],
  FREQ: ["freq", "frequency", "servicefreq", "servicefrequency"],
  del_status: ["delstatus", "deliverystatus", "status", "delivered", "stopstatus"],
  "CSADR#": ["csadr", "addressnumber", "streetnumber", "housenumber"],
  CSSDIR: ["cssdir", "streetdir", "direction", "predir"],
  CSSTRT: ["csstrt", "street", "streetname", "road"],
  CSSFUX: ["cssfux", "streetsuffix", "suffix", "sfx"],
  SIZE: ["size", "containersize", "binsize"],
  QTY: ["qty", "quantity", "count"],
  BINNO: ["binno", "bin", "binnumber", "containerid"],
  DEMAND: ["demand", "tons", "tonnage", "load", "volume", "totaldemand"],
  Service_Time: ["servicetime", "service_time", "serviceminutes", "servicemins", "stopservicetime", "servicemmss"],
  DURATION_A: ["durationa", "duration", "servicetime", "service_time", "servicemin", "serviceminutes", "minutesperstop", "durationminutes"],
  EARLIEST_A: ["earliesta", "earliestarr", "earliestarrival", "windowstart", "timewindowstart"],
  LATEST_ARR: ["latestarr", "latestarrival", "latesta", "latest_a", "windowend", "timewindowend"]
};

function normalizeColumnName(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getMappedRowRoute(row) {
  const preferred = String(row?.ROUTE ?? "").trim();
  if (preferred) return preferred;
  return String(row?.NEWROUTE ?? "").trim();
}

function getMappedRowDayRaw(row) {
  const preferred = String(row?.DAY ?? "").trim();
  if (preferred) return preferred;
  return String(row?.NEWDAY ?? "").trim();
}

const RECORD_FIELD_MIRROR_GROUPS = Object.freeze([
  { primary: "ROUTE", mirror: "NEWROUTE" },
  { primary: "DAY", mirror: "NEWDAY" },
  { primary: "SEQNO", mirror: "NEWSEQ" },
  { primary: "FREQ", mirror: "NEWFREQ" }
]);

const EXIST_BACKUP_SOURCE_FIELDS = Object.freeze([
  "UNIQUEID",
  "DAY",
  "ROUTE",
  "SEQNO",
  "FREQ",
  "SIZE",
  "QTY",
  "BINNO",
  "NEWDAY",
  "NEWROUTE",
  "NEWSEQ",
  "NEWFREQ",
  "LONGITUDE",
  "LATITUDE",
  "DEMAND",
  "Service_Time",
  "DURATION_A",
  "EARLIEST_A",
  "LATEST_ARR"
]);

function isBlankRecordFieldValue(value) {
  return value === null || value === undefined || value === "";
}

function isRouteDayFieldName(fieldName) {
  const token = normalizeRouteSequencerFieldToken(fieldName);
  return token === "route" || token === "day";
}

function hydrateCanonicalRecordFieldsFromLegacy(row) {
  if (!row || typeof row !== "object") return;
  let changed = false;
  RECORD_FIELD_MIRROR_GROUPS.forEach(group => {
    const primaryValue = row[group.primary];
    const mirrorValue = row[group.mirror];
    if (isBlankRecordFieldValue(primaryValue) && !isBlankRecordFieldValue(mirrorValue)) {
      row[group.primary] = mirrorValue;
      changed = true;
    }
  });
  return changed;
}

function ensureCanonicalRecordFields(rows) {
  const list = Array.isArray(rows) ? rows : [];
  let changed = false;
  list.forEach(row => {
    if (hydrateCanonicalRecordFieldsFromLegacy(row)) {
      changed = true;
    }
  });
  return changed;
}

function ensureExistBackupColumns(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return false;

  const activeFields = EXIST_BACKUP_SOURCE_FIELDS.filter(field =>
    list.some(row => row && typeof row === "object" && row[field] !== undefined)
  );
  if (!activeFields.length) return false;

  let changed = false;
  list.forEach(row => {
    if (!row || typeof row !== "object") return;
    activeFields.forEach(field => {
      const backupField = `Exist_${field}`;
      if (Object.prototype.hasOwnProperty.call(row, backupField)) return;
      row[backupField] = row[field];
      changed = true;
    });
  });
  return changed;
}

function ensureServiceTimeField(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return false;
  let changed = false;
  list.forEach(row => {
    if (!row || typeof row !== "object") return;
    if (Object.prototype.hasOwnProperty.call(row, "Service_Time")) return;
    const fallback = row?.DURATION_A;
    row.Service_Time = fallback == null ? "" : fallback;
    changed = true;
  });
  return changed;
}

function refreshRouteDayGroupsFromCurrentRows() {
  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  if (!rows.length || !attributeMarkerByRowId?.size) return false;

  const previousRouteChecks = {};
  document.querySelectorAll("#routeCheckboxes input").forEach(input => {
    previousRouteChecks[String(input.value || "")] = !!input.checked;
  });

  const previousDayChecks = {};
  document.querySelectorAll("#dayCheckboxes input").forEach(input => {
    const dayToken = normalizeDayToken(input.value) || String(input.value || "").trim();
    if (dayToken) previousDayChecks[dayToken] = !!input.checked;
  });

  const nextGroups = {};
  const routeSet = new Set();
  const nextBounds = L.latLngBounds();

  attributeMarkerByRowId.forEach((marker, rowId) => {
    if (!marker) return;
    const row = rows[Number(rowId)];
    if (!row || typeof row !== "object") return;

    const route = getMappedRowRoute(row) || "Unassigned";
    const day = getMappedRowDayRaw(row) || "No Day";
    const key = `${route}|${day}`;

    marker._routeToken = route;
    marker._dayToken = day;
    marker._multiDayLatToken = row?.LATITUDE;
    marker._multiDayLonToken = row?.LONGITUDE;

    const symbol = getSymbol(key);
    if (marker._base && typeof marker._base === "object") {
      marker._base.symbol = symbol;
    }
    marker.setStyle?.({
      color: symbol.color,
      fillColor: symbol.color
    });
    syncRecordPopupContent(marker);

    if (!nextGroups[key]) nextGroups[key] = { layers: [] };
    nextGroups[key].layers.push(marker);
    routeSet.add(route);

    const ll = getLayerLatLng(marker);
    if (ll) nextBounds.extend(ll);
  });

  Object.keys(routeDayGroups).forEach(key => delete routeDayGroups[key]);
  Object.entries(nextGroups).forEach(([key, group]) => {
    routeDayGroups[key] = group;
  });
  globalBounds = nextBounds;
  clearSequenceLayerData({ silent: true });
  syncSequenceLayerVisibilityOnMap();

  Object.keys(layerVisibilityState).forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(routeDayGroups, key)) {
      delete layerVisibilityState[key];
    }
  });
  Object.keys(routeDayGroups).forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(layerVisibilityState, key)) {
      layerVisibilityState[key] = true;
    }
  });

  const sortedRoutes = [...routeSet].sort(compareRouteCodes);
  buildRouteCheckboxes(sortedRoutes);
  document.querySelectorAll("#routeCheckboxes input").forEach(input => {
    const route = String(input.value || "");
    if (Object.prototype.hasOwnProperty.call(previousRouteChecks, route)) {
      input.checked = !!previousRouteChecks[route];
    }
  });
  document.querySelectorAll("#dayCheckboxes input").forEach(input => {
    const token = normalizeDayToken(input.value) || String(input.value || "").trim();
    if (token && Object.prototype.hasOwnProperty.call(previousDayChecks, token)) {
      input.checked = !!previousDayChecks[token];
    }
  });

  buildRouteDayLayerControls();
  rebuildMultiDayServiceProfilesFromMapLayers();
  applyFilters();
  return true;
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
          : "Sample: \u2014";
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

function syncCurrentWorkbookSheetFromRows() {
  if (!window._currentWorkbook || !Array.isArray(window._currentRows)) return false;
  const firstSheetName = Array.isArray(window._currentWorkbook.SheetNames)
    ? window._currentWorkbook.SheetNames[0]
    : "";
  if (!firstSheetName) return false;
  const nextSheet = XLSX.utils.json_to_sheet(window._currentRows, { defval: "" });
  window._currentWorkbook.Sheets[firstSheetName] = nextSheet;
  return true;
}

function hasWorkbookCloudSaveContext() {
  const filePath = String(window._currentFilePath || "").trim();
  return !!filePath && !!window._currentWorkbook && Array.isArray(window._currentRows);
}

async function flushQueuedWorkbookCloudSave() {
  if (!workbookAutoSavePendingReasons.length) return false;
  if (!hasWorkbookCloudSaveContext()) {
    workbookAutoSavePendingReasons = [];
    return false;
  }
  if (workbookAutoSaveInFlight) return workbookAutoSaveInFlight;

  const reasons = [...workbookAutoSavePendingReasons];
  workbookAutoSavePendingReasons = [];
  const reasonLabel = reasons.join(", ");

  workbookAutoSaveInFlight = (async () => {
    try {
      const saved = await saveWorkbookToCloud();
      if (!saved) {
        console.warn(
          `[Workbook Auto Save] Save failed for "${window._currentFilePath}"${reasonLabel ? ` (${reasonLabel})` : ""}.`
        );
        return false;
      }
      console.info(
        `[Workbook Auto Save] Saved "${window._currentFilePath}"${reasonLabel ? ` (${reasonLabel})` : ""}.`
      );
      return true;
    } catch (error) {
      console.error(
        `[Workbook Auto Save] Error saving "${window._currentFilePath}"${reasonLabel ? ` (${reasonLabel})` : ""}:`,
        error
      );
      return false;
    } finally {
      workbookAutoSaveInFlight = null;
      if (workbookAutoSavePendingReasons.length) {
        if (workbookAutoSaveTimer) clearTimeout(workbookAutoSaveTimer);
        workbookAutoSaveTimer = setTimeout(() => {
          workbookAutoSaveTimer = null;
          flushQueuedWorkbookCloudSave();
        }, 300);
      }
    }
  })();

  return workbookAutoSaveInFlight;
}

function queueWorkbookCloudSave(reason = "update") {
  if (!hasWorkbookCloudSaveContext()) return;
  const reasonText = String(reason || "").trim();
  if (reasonText) workbookAutoSavePendingReasons.push(reasonText);
  if (workbookAutoSaveTimer) clearTimeout(workbookAutoSaveTimer);
  workbookAutoSaveTimer = setTimeout(() => {
    workbookAutoSaveTimer = null;
    flushQueuedWorkbookCloudSave();
  }, WORKBOOK_AUTO_SAVE_DEBOUNCE_MS);
}

function flushWorkbookCloudSaveOnLifecycle(force = false) {
  if (!hasWorkbookCloudSaveContext()) return;
  if (!force && !workbookAutoSavePendingReasons.length && !workbookAutoSaveTimer) return;
  if (workbookAutoSaveTimer) {
    clearTimeout(workbookAutoSaveTimer);
    workbookAutoSaveTimer = null;
  }
  if (!workbookAutoSavePendingReasons.length) {
    workbookAutoSavePendingReasons.push("lifecycle");
  }
  flushQueuedWorkbookCloudSave();
}

window.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") {
    flushWorkbookCloudSaveOnLifecycle(true);
    flushSharedFacilitiesCloudSaveOnLifecycle(true);
  }
});
window.addEventListener("pagehide", () => {
  flushWorkbookCloudSaveOnLifecycle(true);
  flushSharedFacilitiesCloudSaveOnLifecycle(true);
});
window.addEventListener("beforeunload", () => {
  flushWorkbookCloudSaveOnLifecycle(true);
  flushSharedFacilitiesCloudSaveOnLifecycle(true);
});

function getActiveAttributeSortOptions() {
  if (attributeTableMode === "streets") {
    return STREET_ATTRIBUTE_TABLE_HEADERS.map(key => ({
      key,
      label: STREET_ATTRIBUTE_TABLE_LABELS[key] || key
    }));
  }
  const headers = Array.isArray(window._attributeHeaders) ? window._attributeHeaders : [];
  return headers
    .filter(header => String(header || "").trim())
    .map(key => ({ key, label: String(key) }));
}

function syncAttributeSortControls() {
  const sortFieldSelect = document.getElementById("attributeSortFieldSelect");
  const sortAscBtn = document.getElementById("attributeSortAscBtn");
  const sortDescBtn = document.getElementById("attributeSortDescBtn");
  const clearSortBtn = document.getElementById("attributeClearSortBtn");
  const addFieldBtn = document.getElementById("attributeAddFieldBtn");
  const fieldCalcBtn = document.getElementById("attributeFieldCalcBtn");
  if (!sortFieldSelect) return;

  const options = getActiveAttributeSortOptions();
  const optionKeySet = new Set(options.map(option => option.key));
  if (attributeState.sortKey && !optionKeySet.has(attributeState.sortKey)) {
    attributeState.sortKey = null;
    attributeState.sortDir = 1;
  }

  sortFieldSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Sort field...";
  sortFieldSelect.appendChild(placeholder);
  options.forEach(option => {
    const node = document.createElement("option");
    node.value = option.key;
    node.textContent = option.label;
    sortFieldSelect.appendChild(node);
  });
  sortFieldSelect.value = attributeState.sortKey || "";

  const hasOptions = options.length > 0;
  sortFieldSelect.disabled = !hasOptions;
  if (sortAscBtn) sortAscBtn.disabled = !hasOptions;
  if (sortDescBtn) sortDescBtn.disabled = !hasOptions;
  if (clearSortBtn) clearSortBtn.disabled = !hasOptions || !attributeState.sortKey;
  if (addFieldBtn) addFieldBtn.style.display = attributeTableMode === "records" ? "" : "none";
  if (fieldCalcBtn) fieldCalcBtn.style.display = attributeTableMode === "records" ? "" : "none";
}

function applyAttributeSortFromToolbar(direction) {
  const sortFieldSelect = document.getElementById("attributeSortFieldSelect");
  const key = String(sortFieldSelect?.value || attributeState.sortKey || "").trim();
  if (!key) {
    alert("Choose a field to sort.");
    return;
  }
  attributeState.sortKey = key;
  attributeState.sortDir = direction >= 0 ? 1 : -1;
  attributeState.page = 1;
  renderAttributeTable();
}

function promptAndAddAttributeField() {
  if (attributeTableMode !== "records") {
    alert("Switch to Record Attributes to add a new field.");
    return;
  }

  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  if (!rows.length) {
    alert("Upload a route file before adding fields.");
    return;
  }

  const rawName = window.prompt("New field name:", "");
  if (rawName === null) return;
  const fieldName = String(rawName || "").trim();
  if (!fieldName) {
    alert("Field name is required.");
    return;
  }

  const existingHeaders = Array.isArray(window._attributeHeaders) ? window._attributeHeaders : getAttributeHeaders(rows);
  const duplicate = existingHeaders.find(header => String(header || "").trim().toLowerCase() === fieldName.toLowerCase());
  if (duplicate) {
    alert(`Field "${duplicate}" already exists.`);
    return;
  }

  const defaultInput = window.prompt(`Default value for "${fieldName}" (optional):`, "");
  if (defaultInput === null) return;
  const defaultValue = String(defaultInput || "");

  rows.forEach(row => {
    if (!row || typeof row !== "object") return;
    row[fieldName] = defaultValue;
  });

  window._attributeHeaders = getAttributeHeaders(rows);
  syncCurrentWorkbookSheetFromRows();
  queueWorkbookCloudSave("Add Field");
  const selectAttributesSource = document.getElementById("selectAttributesSource");
  if (selectAttributesSource && String(selectAttributesSource.value || "").toLowerCase() === "records") {
    selectAttributesSource.dispatchEvent(new Event("change"));
  }
  syncAttributeSortControls();
  renderAttributeTable();
}

const FIELD_CALCULATOR_FN_ALIAS_MAP = {
  abs: "abs",
  atn: "atan",
  atan: "atan",
  cos: "cos",
  exp: "exp",
  fix: "fix",
  int: "int",
  log: "log",
  sin: "sin",
  sqr: "sqrt",
  sqrt: "sqrt",
  tan: "tan",
  round: "round",
  min: "min",
  max: "max",
  pow: "pow",
  pi: "pi",
  e: "e"
};

const FIELD_CALCULATOR_RUNTIME = {
  abs: value => Math.abs(Number(value) || 0),
  atan: value => Math.atan(Number(value) || 0),
  cos: value => Math.cos(Number(value) || 0),
  exp: value => Math.exp(Number(value) || 0),
  fix: value => Math.trunc(Number(value) || 0),
  int: value => Math.trunc(Number(value) || 0),
  log: value => Math.log(Number(value) || 0),
  sin: value => Math.sin(Number(value) || 0),
  sqrt: value => Math.sqrt(Number(value) || 0),
  tan: value => Math.tan(Number(value) || 0),
  round: (value, decimals = 0) => {
    const val = Number(value) || 0;
    const places = Math.max(0, Math.trunc(Number(decimals) || 0));
    if (!places) return Math.round(val);
    const factor = 10 ** places;
    return Math.round(val * factor) / factor;
  },
  min: (...values) => {
    const nums = values.map(v => Number(v)).filter(Number.isFinite);
    return nums.length ? Math.min(...nums) : 0;
  },
  max: (...values) => {
    const nums = values.map(v => Number(v)).filter(Number.isFinite);
    return nums.length ? Math.max(...nums) : 0;
  },
  pow: (base, exponent) => Math.pow(Number(base) || 0, Number(exponent) || 0),
  pi: Math.PI,
  e: Math.E
};

function escapeFieldCalculatorRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function resolveFieldCalculatorFieldName(fieldToken, headers) {
  const token = String(fieldToken || "").trim().toLowerCase();
  if (!token) return "";
  const options = Array.isArray(headers) ? headers : [];
  const direct = options.find(header => String(header || "").trim().toLowerCase() === token);
  if (direct) return direct;
  return "";
}

function coerceFieldCalculatorInputValue(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "boolean") return value ? 1 : 0;
  const text = String(value).trim();
  if (!text) return 0;
  const numeric = Number(text.replace(/,/g, ""));
  if (Number.isFinite(numeric)) return numeric;
  return text;
}

function compileFieldCalculatorExpression(expressionText, headers) {
  const source = String(expressionText || "").replace(/\r?\n/g, " ").trim();
  if (!source) {
    throw new Error("Expression is required.");
  }

  const fieldRefs = [];
  let fieldTokenIndex = 0;
  const rewrittenWithFields = source.replace(/\[([^\]]+)\]/g, (_, rawFieldName) => {
    const tokenText = String(rawFieldName || "").trim();
    const resolvedField = resolveFieldCalculatorFieldName(tokenText, headers);
    if (!resolvedField) {
      throw new Error(`Field "${tokenText}" was not found.`);
    }
    const token = `f${fieldTokenIndex}`;
    fieldTokenIndex += 1;
    fieldRefs.push({ token, fieldName: resolvedField });
    return token;
  });

  let normalized = rewrittenWithFields
    .replace(/[\u00D7\u2715]/g, "*")
    .replace(/[\u00F7]/g, "/")
    .replace(/\^/g, "**")
    .trim();

  if (!normalized) {
    throw new Error("Expression is empty.");
  }

  if (/[\[\]{};"'`\\]/.test(normalized)) {
    throw new Error("Expression contains unsupported characters.");
  }

  const illegalCharMatch = normalized.match(/[^0-9A-Za-z_+\-*/%().,<>=!?:&|\s]/);
  if (illegalCharMatch) {
    throw new Error(`Unsupported character "${illegalCharMatch[0]}" in expression.`);
  }

  const identifiers = normalized.match(/\b[A-Za-z_][A-Za-z0-9_]*\b/g) || [];
  const unknownIdentifier = identifiers.find(identifier => {
    if (/^f\d+$/.test(identifier)) return false;
    return !FIELD_CALCULATOR_FN_ALIAS_MAP[String(identifier || "").toLowerCase()];
  });

  if (unknownIdentifier) {
    throw new Error(`Unknown token "${unknownIdentifier}". Use [FieldName] for fields.`);
  }

  let jsExpression = normalized;
  Object.entries(FIELD_CALCULATOR_FN_ALIAS_MAP).forEach(([alias, mapped]) => {
    const regex = new RegExp(`\\b${escapeFieldCalculatorRegex(alias)}\\b`, "gi");
    jsExpression = jsExpression.replace(regex, `__fns.${mapped}`);
  });

  const uniqueFieldTokens = [...new Set(fieldRefs.map(ref => ref.token))];
  const destructure = uniqueFieldTokens.length ? `const { ${uniqueFieldTokens.join(", ")} } = __vars;` : "";
  let evaluator;
  try {
    evaluator = new Function("__vars", "__fns", `"use strict"; ${destructure} return (${jsExpression});`);
  } catch (err) {
    throw new Error(`Expression is not valid: ${String(err?.message || err)}`);
  }

  return {
    evaluator,
    fieldRefs
  };
}

function applyFieldCalculatorToRows(targetField, expressionText, outputType = "number") {
  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  if (!rows.length) {
    throw new Error("No record rows are loaded.");
  }

  const headers = Array.isArray(window._attributeHeaders) ? window._attributeHeaders : getAttributeHeaders(rows);
  const resolvedTarget = resolveFieldCalculatorFieldName(targetField, headers) || String(targetField || "").trim();
  if (!resolvedTarget) {
    throw new Error("Choose a target field.");
  }

  const compiled = compileFieldCalculatorExpression(expressionText, headers);
  const computedValues = new Array(rows.length);

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || {};
    const vars = {};
    compiled.fieldRefs.forEach(ref => {
      vars[ref.token] = coerceFieldCalculatorInputValue(row?.[ref.fieldName]);
    });

    let rawResult;
    try {
      rawResult = compiled.evaluator(vars, FIELD_CALCULATOR_RUNTIME);
    } catch (err) {
      throw new Error(`Row ${i + 1}: ${String(err?.message || err)}`);
    }

    if (String(outputType || "number").toLowerCase() === "text") {
      computedValues[i] = rawResult === null || rawResult === undefined ? "" : String(rawResult);
      continue;
    }

    const numeric = Number(rawResult);
    if (!Number.isFinite(numeric)) {
      throw new Error(`Row ${i + 1}: result "${String(rawResult)}" is not a valid number.`);
    }
    computedValues[i] = numeric;
  }

  const routeDayTouched = isRouteDayFieldName(resolvedTarget);
  rows.forEach((row, index) => {
    if (!row || typeof row !== "object") return;
    row[resolvedTarget] = computedValues[index];
  });
  ensureExistBackupColumns(rows);
  if (routeDayTouched) refreshRouteDayGroupsFromCurrentRows();

  window._attributeHeaders = getAttributeHeaders(rows);
  syncCurrentWorkbookSheetFromRows();
  queueWorkbookCloudSave("Field Calculator");
  if (typeof refreshAllRecordPopupContent === "function") refreshAllRecordPopupContent();
  const selectAttributesSource = document.getElementById("selectAttributesSource");
  if (selectAttributesSource && String(selectAttributesSource.value || "").toLowerCase() === "records") {
    selectAttributesSource.dispatchEvent(new Event("change"));
  }
  syncAttributeSortControls();
  renderAttributeTable();

  return {
    count: rows.length,
    targetField: resolvedTarget
  };
}

function initFieldCalculatorControls() {
  const openBtn = document.getElementById("attributeFieldCalcBtn");
  const modal = document.getElementById("fieldCalculatorModal");
  const closeBtn = document.getElementById("fieldCalculatorCloseBtn");
  const clearBtn = document.getElementById("fieldCalculatorClearBtn");
  const applyBtn = document.getElementById("fieldCalculatorApplyBtn");
  const targetFieldSelect = document.getElementById("fieldCalculatorTargetField");
  const typeSelect = document.getElementById("fieldCalculatorType");
  const expressionInput = document.getElementById("fieldCalculatorExpression");
  const previewNode = document.getElementById("fieldCalculatorPreview");
  const statusNode = document.getElementById("fieldCalculatorStatus");
  const fieldsNode = document.getElementById("fieldCalculatorFieldList");
  const tokenButtons = [...document.querySelectorAll("#fieldCalculatorTokenBar .field-calculator-token-btn")];

  if (
    !openBtn || !modal || !closeBtn || !clearBtn || !applyBtn ||
    !targetFieldSelect || !typeSelect || !expressionInput || !previewNode || !statusNode || !fieldsNode
  ) {
    return;
  }
  if (openBtn.dataset.fieldCalculatorBound === "1") return;
  openBtn.dataset.fieldCalculatorBound = "1";

  const setStatus = (message, type = "") => {
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (type === "error") statusNode.classList.add("error");
    if (type === "success") statusNode.classList.add("success");
  };

  const getHeaders = () => {
    const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
    if (!rows.length) return [];
    return Array.isArray(window._attributeHeaders) && window._attributeHeaders.length
      ? window._attributeHeaders
      : getAttributeHeaders(rows);
  };

  const updatePreview = () => {
    const field = String(targetFieldSelect.value || "").trim() || "(target field)";
    const expr = String(expressionInput.value || "").trim() || "(enter expression)";
    previewNode.textContent = `${field} = ${expr}`;
  };

  const insertExpressionToken = token => {
    const text = String(token || "");
    if (!text) return;
    const input = expressionInput;
    const start = Number(input.selectionStart || 0);
    const end = Number(input.selectionEnd || start);
    const before = input.value.slice(0, start);
    const after = input.value.slice(end);
    input.value = `${before}${text}${after}`;
    const nextPos = start + text.length;
    input.setSelectionRange(nextPos, nextPos);
    input.focus();
    updatePreview();
  };

  const refreshTargetFields = (preferredField = "") => {
    const headers = getHeaders();
    targetFieldSelect.innerHTML = "";
    headers.forEach(field => {
      const option = document.createElement("option");
      option.value = field;
      option.textContent = field;
      targetFieldSelect.appendChild(option);
    });
    if (preferredField && headers.includes(preferredField)) {
      targetFieldSelect.value = preferredField;
    } else if (!targetFieldSelect.value && headers.length) {
      targetFieldSelect.value = headers[0];
    }
  };

  const renderFieldButtons = () => {
    const headers = getHeaders();
    fieldsNode.innerHTML = "";
    if (!headers.length) {
      fieldsNode.innerHTML = '<div class="select-attributes-unique-empty">No fields available.</div>';
      return;
    }
    headers.forEach(field => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "field-calculator-field-btn";
      btn.textContent = `[${field}]`;
      btn.addEventListener("click", () => insertExpressionToken(`[${field}]`));
      fieldsNode.appendChild(btn);
    });
  };

  const openModal = () => {
    if (attributeTableMode !== "records") {
      setAttributeTableMode("records");
      openAttributePanel();
    }
    const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
    if (!rows.length) {
      alert("Upload a route file before using Field Calculator.");
      return;
    }
    refreshTargetFields(targetFieldSelect.value);
    renderFieldButtons();
    updatePreview();
    setStatus("Ready.");
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  openBtn.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  clearBtn.addEventListener("click", () => {
    expressionInput.value = "";
    updatePreview();
    setStatus("Expression cleared.");
    expressionInput.focus();
  });
  applyBtn.addEventListener("click", () => {
    const targetField = String(targetFieldSelect.value || "").trim();
    const expression = String(expressionInput.value || "");
    const outputType = String(typeSelect.value || "number");

    try {
      const result = applyFieldCalculatorToRows(targetField, expression, outputType);
      setStatus(`Updated ${result.count.toLocaleString()} records in "${result.targetField}".`, "success");
      updatePreview();
    } catch (err) {
      setStatus(String(err?.message || err), "error");
    }
  });

  expressionInput.addEventListener("input", updatePreview);
  targetFieldSelect.addEventListener("change", updatePreview);
  typeSelect.addEventListener("change", updatePreview);
  tokenButtons.forEach(btn => {
    btn.addEventListener("click", () => {
      const token = String(btn.getAttribute("data-token") || "");
      if (!token) return;
      insertExpressionToken(token);
    });
  });

  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });
  window.addEventListener("keydown", event => {
    if (modal.style.display !== "flex") return;
    if (event.key === "Escape") closeModal();
  });
}

function getUpdateFieldFormatHints(targetFieldName) {
  const token = normalizeRouteSequencerFieldToken(targetFieldName);
  if (!token) {
    return [
      "Tip: Choose a target field, then update all or selected records from another field, a typed value, or clear."
    ];
  }

  const hints = [];
  if (/(demand|tons|tonnage|load|volume|qty|quantity)/.test(token)) {
    hints.push("Demand fields should be numeric tons (example: 0.5, 3, 12.75).");
  }
  if (token === "servicetime") {
    hints.push("Service_Time supports MM:SS or HH:MM:SS (example: 05:30) and numeric minutes (example: 5.5).");
  } else if (/(servicetime|servicemin|service|minperstop|minutesperstop|timeonsite|onsitemin|service)/.test(token)) {
    hints.push("Service Time should be numeric minutes per stop (example: 2, 7.5, 12).");
  }
  if (/(sequence|trip|cumtons|drivemin|servicemin|totalmin|dumpvisits)/.test(token)) {
    hints.push("Sequencer output fields should stay numeric so sorting and route popups work correctly.");
  }
  if (/(route|day)/.test(token)) {
    hints.push("DAY and ROUTE are the active planning fields used by map layers and solvers.");
  }
  if (/(seqno|sequence)/.test(token)) {
    hints.push("Auto Sequence writes stop order to SEQNO.");
  }
  if (/(newroute|newday|newseq|newfreq)/.test(token)) {
    hints.push("NEW* fields are preserved for final publishing and only change when you edit them directly.");
  }
  if (/(latitude|lat)/.test(token)) {
    hints.push("Latitude should be decimal degrees between -90 and 90.");
  }
  if (/(longitude|lon|lng)/.test(token)) {
    hints.push("Longitude should be decimal degrees between -180 and 180.");
  }
  if (!hints.length) {
    hints.push("Tip: Keep field values consistent (all numeric or all text) for reliable sorting, filtering, and charting.");
  }
  return hints;
}

function looksLikeNumericValue(value) {
  if (value == null || value === "") return false;
  if (typeof value === "number") return Number.isFinite(value);
  const text = String(value).trim();
  if (!text) return false;
  const parsed = Number(text.replace(/,/g, ""));
  return Number.isFinite(parsed);
}

function shouldTreatUpdateFieldAsNumeric(fieldName, rows) {
  const token = normalizeRouteSequencerFieldToken(fieldName);
  if (/(demand|tons|tonnage|load|volume|qty|quantity|sequence|trip|min|time|lat|lon|lng|speed|distance|count|score)/.test(token)) {
    return true;
  }
  const list = Array.isArray(rows) ? rows : [];
  let numericCount = 0;
  let nonEmptyCount = 0;
  for (let i = 0; i < list.length && nonEmptyCount < 40; i += 1) {
    const row = list[i];
    if (!row || typeof row !== "object") continue;
    const value = row[fieldName];
    if (value == null || value === "") continue;
    nonEmptyCount += 1;
    if (looksLikeNumericValue(value)) numericCount += 1;
  }
  return nonEmptyCount > 0 && (numericCount / nonEmptyCount) >= 0.7;
}

function isUpdateFieldClockMinutesCompatible(fieldName) {
  const token = normalizeRouteSequencerFieldToken(fieldName);
  if (!token) return false;
  return (
    token === "servicetime" ||
    token === "durationa" ||
    token === "duration" ||
    token === "servicemin" ||
    token === "serviceminutes" ||
    token === "timeonsite"
  );
}

function initUpdateFieldControls() {
  const openBtn = document.getElementById("updateFieldBtn");
  const openBtnMobile = document.getElementById("updateFieldBtnMobile");
  const modal = document.getElementById("updateFieldModal");
  const closeBtn = document.getElementById("updateFieldCloseBtn");
  const applyBtn = document.getElementById("updateFieldApplyBtn");
  const targetFieldSelect = document.getElementById("updateFieldTargetField");
  const sourceFieldSelect = document.getElementById("updateFieldSourceField");
  const valueInput = document.getElementById("updateFieldValueInput");
  const scopeAllRadio = document.getElementById("updateFieldScopeAll");
  const scopeSelectedRadio = document.getElementById("updateFieldScopeSelected");
  const allLabel = document.getElementById("updateFieldAllLabel");
  const selectedLabel = document.getElementById("updateFieldSelectedLabel");
  const modeFieldRadio = document.getElementById("updateFieldModeField");
  const modeValueRadio = document.getElementById("updateFieldModeValue");
  const modeClearRadio = document.getElementById("updateFieldModeClear");
  const formatHintNode = document.getElementById("updateFieldFormatHint");
  const statusNode = document.getElementById("updateFieldStatus");

  if (
    !openBtn || !modal || !closeBtn || !applyBtn || !targetFieldSelect || !sourceFieldSelect ||
    !valueInput || !scopeAllRadio || !scopeSelectedRadio || !allLabel || !selectedLabel ||
    !modeFieldRadio || !modeValueRadio || !modeClearRadio || !formatHintNode || !statusNode
  ) {
    return;
  }

  if (openBtn.dataset.updateFieldBound === "1") return;
  openBtn.dataset.updateFieldBound = "1";

  const setStatus = (message, kind = "") => {
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (kind === "error") statusNode.classList.add("error");
    if (kind === "success") statusNode.classList.add("success");
  };

  const getRows = () => (Array.isArray(window._currentRows) ? window._currentRows : []);
  const getHeaders = rows => {
    const list = Array.isArray(rows) ? rows : [];
    if (!list.length) return [];
    return Array.isArray(window._attributeHeaders) && window._attributeHeaders.length
      ? window._attributeHeaders
      : getAttributeHeaders(list);
  };

  const populateSelect = (node, headers, preferred = "") => {
    const options = Array.isArray(headers) ? headers : [];
    const previous = String(preferred || node.value || "");
    node.innerHTML = "";
    options.forEach(field => {
      const option = document.createElement("option");
      option.value = String(field);
      option.textContent = String(field);
      node.appendChild(option);
    });
    if ([...node.options].some(option => option.value === previous)) {
      node.value = previous;
    } else if (node.options.length) {
      node.value = node.options[0].value;
    } else {
      node.value = "";
    }
  };

  const getSelectedMode = () => {
    if (modeValueRadio.checked) return "value";
    if (modeClearRadio.checked) return "clear";
    return "field";
  };

  const updateModeUi = () => {
    const mode = getSelectedMode();
    sourceFieldSelect.disabled = mode !== "field";
    valueInput.disabled = mode !== "value";
    if (mode === "field") {
      valueInput.placeholder = "Disabled while 'Field' mode is selected";
    } else {
      valueInput.placeholder = "Example: 4.5";
    }
  };

  const updateScopeLabels = rows => {
    const list = Array.isArray(rows) ? rows : getRows();
    const allCount = list.length;
    const selectedCount = attributeState.selectedRowIds.size;
    allLabel.textContent = `All (${allCount.toLocaleString()} records)`;
    selectedLabel.textContent = `Selected (${selectedCount.toLocaleString()} records)`;
    scopeSelectedRadio.disabled = selectedCount < 1;
    if (scopeSelectedRadio.disabled) scopeAllRadio.checked = true;
  };

  const updateFormatHint = () => {
    const targetField = String(targetFieldSelect.value || "");
    formatHintNode.textContent = getUpdateFieldFormatHints(targetField).join(" ");
  };

  const refreshFieldSelectors = () => {
    const rows = getRows();
    const headers = getHeaders(rows);
    const prevTarget = String(targetFieldSelect.value || "");
    const prevSource = String(sourceFieldSelect.value || "");
    populateSelect(targetFieldSelect, headers, prevTarget);
    populateSelect(sourceFieldSelect, headers, prevSource);
    if (!sourceFieldSelect.value && headers.length > 1) {
      const fallback = headers.find(field => String(field) !== String(targetFieldSelect.value || ""));
      if (fallback) sourceFieldSelect.value = String(fallback);
    }
    applyBtn.disabled = !headers.length || !rows.length;
    updateScopeLabels(rows);
    updateFormatHint();
  };

  const openModal = () => {
    const rows = getRows();
    if (!rows.length) {
      alert("Upload a route file before using Update Field.");
      return;
    }
    if (attributeTableMode !== "records") {
      setAttributeTableMode("records");
      openAttributePanel();
    }
    refreshFieldSelectors();
    updateModeUi();
    setStatus("Ready.");
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  const resolveScopeRowIds = rows => {
    if (!scopeSelectedRadio.checked) {
      return rows.map((_, index) => index);
    }
    return [...attributeState.selectedRowIds]
      .map(id => Number(id))
      .filter(id => Number.isFinite(id) && id >= 0 && id < rows.length);
  };

  const applyUpdate = () => {
    const rows = getRows();
    if (!rows.length) {
      setStatus("No records are loaded.", "error");
      return;
    }

    const targetField = String(targetFieldSelect.value || "").trim();
    if (!targetField) {
      setStatus("Choose a target field.", "error");
      return;
    }

    const rowIds = resolveScopeRowIds(rows);
    if (!rowIds.length) {
      setStatus("No records match the selected scope.", "error");
      return;
    }

    const mode = getSelectedMode();
    const sourceField = String(sourceFieldSelect.value || "").trim();
    const typedRawValue = valueInput.value;

    if (mode === "field" && !sourceField) {
      setStatus("Choose a source field.", "error");
      return;
    }
    if (mode === "value" && !String(typedRawValue || "").trim()) {
      setStatus("Enter a typed value or choose Clear.", "error");
      return;
    }

    let typedValue = String(typedRawValue || "");
    if (mode === "value" && shouldTreatUpdateFieldAsNumeric(targetField, rows)) {
      const trimmed = typedValue.trim();
      if (!trimmed) typedValue = "";
      else {
        const allowClockMinutes = isUpdateFieldClockMinutesCompatible(targetField);
        if (allowClockMinutes && trimmed.includes(":")) {
          const parsedMinutes = parseRouteSequencerMinutesValue(trimmed, NaN);
          if (!Number.isFinite(parsedMinutes)) {
            setStatus(
              `"${typedRawValue}" is not a valid time value for "${targetField}". Use MM:SS or HH:MM:SS.`,
              "error"
            );
            return;
          }
          typedValue = formatRouteSequencerMinutesClock(parsedMinutes);
        } else {
          const parsed = Number(trimmed.replace(/,/g, ""));
          if (!Number.isFinite(parsed)) {
            setStatus(`"${typedRawValue}" is not a valid numeric value for "${targetField}".`, "error");
            return;
          }
          typedValue = parsed;
        }
      }
    }

    let changedCount = 0;
    const routeDayTouched = isRouteDayFieldName(targetField);
    rowIds.forEach(rowId => {
      const row = rows[rowId];
      if (!row || typeof row !== "object") return;
      if (mode === "field") {
        row[targetField] = row[sourceField];
      } else if (mode === "value") {
        row[targetField] = typedValue;
      } else {
        row[targetField] = "";
      }
      changedCount += 1;
    });
    ensureExistBackupColumns(rows);
    if (routeDayTouched) refreshRouteDayGroupsFromCurrentRows();

    window._attributeHeaders = getAttributeHeaders(rows);
    syncCurrentWorkbookSheetFromRows();
    queueWorkbookCloudSave("Update Field");
    refreshAllRecordPopupContent();
    applyAttributeSelectionStyles();
    syncAttributeSortControls();
    renderAttributeTable();
    refreshAttributeStatus();

    const selectAttributesSource = document.getElementById("selectAttributesSource");
    if (selectAttributesSource && String(selectAttributesSource.value || "").toLowerCase() === "records") {
      selectAttributesSource.dispatchEvent(new Event("change"));
    }

    const modeLabel = mode === "field" ? `field "${sourceField}"` : mode === "value" ? "typed value" : "clear";
    setStatus(
      `Updated ${changedCount.toLocaleString()} record(s) in "${targetField}" using ${modeLabel}.`,
      "success"
    );
    updateScopeLabels(rows);
    updateFormatHint();
  };

  openBtn.addEventListener("click", openModal);
  if (openBtnMobile) openBtnMobile.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  applyBtn.addEventListener("click", applyUpdate);

  [scopeAllRadio, scopeSelectedRadio, modeFieldRadio, modeValueRadio, modeClearRadio].forEach(node => {
    node.addEventListener("change", () => {
      updateModeUi();
      updateScopeLabels(getRows());
      updateFormatHint();
    });
  });
  targetFieldSelect.addEventListener("change", updateFormatHint);
  sourceFieldSelect.addEventListener("change", () => {
    if (getSelectedMode() === "field") setStatus("Ready.");
  });
  valueInput.addEventListener("input", () => {
    if (getSelectedMode() === "value") setStatus("Ready.");
  });

  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });
  window.addEventListener("keydown", event => {
    if (modal.style.display !== "flex") return;
    if (event.key === "Escape") closeModal();
  });
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
  status.textContent = `${selectedCount} selected \u2022 ${visibleCount} visible`;
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
  syncSelectedStopsHeaderCount(getCurrentAttributeModeSelectionCount());
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
  syncSelectedStopsHeaderCount(getCurrentAttributeModeSelectionCount());
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
    const headers = STREET_ATTRIBUTE_TABLE_HEADERS;
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

function initMultiDayManagerControls() {
  const openBtn = document.getElementById("multiDayToggleBtn");
  const modal = document.getElementById("multiDayManagerModal");
  const closeBtn = document.getElementById("multiDayManagerCloseBtn");
  const labelsToggle = document.getElementById("multiDayLabelsToggle");
  const showAllBtn = document.getElementById("multiDayFreqShowAllBtn");
  const hideAllBtn = document.getElementById("multiDayFreqHideAllBtn");
  const singleOnlyBtn = document.getElementById("multiDayFreqSingleOnlyBtn");
  const multiOnlyBtn = document.getElementById("multiDayFreqMultiOnlyBtn");
  const listNode = document.getElementById("multiDayFrequencyList");
  const statusNode = document.getElementById("multiDayManagerStatus");

  if (!openBtn || !modal || !closeBtn || !labelsToggle || !showAllBtn || !hideAllBtn || !singleOnlyBtn || !multiOnlyBtn || !listNode || !statusNode) {
    return;
  }
  if (openBtn.dataset.multiDayBound === "1") return;
  openBtn.dataset.multiDayBound = "1";
  multiDayToggleButtonRef = openBtn;

  const isModalOpen = () => modal.style.display === "flex";

  const setStatus = message => {
    statusNode.textContent = String(message || "");
  };

  const getSortedFrequencyEntries = () => {
    return [...multiDayFrequencyCatalog.values()].sort((a, b) => {
      if (!!b.isMulti !== !!a.isMulti) return b.isMulti ? 1 : -1;
      if ((b.dayCount || 0) !== (a.dayCount || 0)) return (b.dayCount || 0) - (a.dayCount || 0);
      return String(a.detailLabel || "").localeCompare(String(b.detailLabel || ""), undefined, { sensitivity: "base" });
    });
  };

  const renderFrequencyList = () => {
    listNode.innerHTML = "";
    const entries = getSortedFrequencyEntries();
    if (!entries.length) {
      const empty = document.createElement("div");
      empty.className = "multiday-frequency-empty";
      empty.textContent = "Load a route file to build service-frequency options.";
      listNode.appendChild(empty);
      return;
    }

    entries.forEach(entry => {
      const row = document.createElement("div");
      row.className = "multiday-frequency-row";

      const label = document.createElement("label");
      const input = document.createElement("input");
      input.type = "checkbox";
      input.checked = isMultiDayFrequencyVisibleByKey(entry.key);
      input.dataset.frequencyKey = entry.key;

      const main = document.createElement("span");
      main.className = "multiday-frequency-main";

      const chip = document.createElement("span");
      chip.className = `multiday-frequency-chip ${entry.isMulti ? "multi" : "single"}`;
      chip.textContent = entry.isMulti ? "Multi-Day" : "Single-Day";

      const text = document.createElement("span");
      text.className = "multiday-frequency-label";
      text.textContent = entry.detailLabel || "Unknown";

      const meta = document.createElement("span");
      meta.className = "multiday-frequency-meta";
      const locations = Number(entry.locationCount) || 0;
      const records = Number(entry.recordCount) || 0;
      meta.textContent = `${locations.toLocaleString()} locations | ${records.toLocaleString()} records`;

      main.append(chip, text);
      label.append(input, main, meta);
      row.appendChild(label);
      listNode.appendChild(row);
    });
  };

  const refreshStatus = () => {
    const summary = getMultiDayFrequencyVisibilitySummary();
    if (!summary.totalFreq) {
      setStatus("No service-frequency groups yet. Load route records to configure visibility.");
      return;
    }
    if (summary.visibleFreq <= 0) {
      setStatus(`All ${summary.totalFreq.toLocaleString()} service-frequency groups are hidden.`);
      return;
    }
    if (summary.visibleFreq === summary.totalFreq) {
      setStatus(
        `Showing all ${summary.totalFreq.toLocaleString()} service frequencies (${summary.visibleRecords.toLocaleString()} records).`
      );
      return;
    }
    setStatus(
      `Showing ${summary.visibleFreq.toLocaleString()} of ${summary.totalFreq.toLocaleString()} service frequencies (${summary.visibleRecords.toLocaleString()} of ${summary.totalRecords.toLocaleString()} records).`
    );
  };

  const renderModalUi = () => {
    labelsToggle.checked = !!multiDayServiceEnabled;
    renderFrequencyList();
    refreshStatus();
  };

  const openModal = () => {
    renderModalUi();
    modal.style.display = "flex";
  };

  const closeModal = () => {
    modal.style.display = "none";
  };

  labelsToggle.addEventListener("change", () => {
    setMultiDayServiceEnabled(labelsToggle.checked);
  });

  listNode.addEventListener("change", event => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    const key = String(input.dataset.frequencyKey || "");
    if (!key) return;
    if (!setMultiDayFrequencyVisibility(key, input.checked)) return;
    applyFilters();
    renderModalUi();
  });

  showAllBtn.addEventListener("click", () => {
    setAllMultiDayFrequencyVisibility(true);
    applyFilters();
    renderModalUi();
  });

  hideAllBtn.addEventListener("click", () => {
    setAllMultiDayFrequencyVisibility(false);
    applyFilters();
    renderModalUi();
  });

  singleOnlyBtn.addEventListener("click", () => {
    setMultiDayFrequencyVisibilityByPredicate(entry => !entry.isMulti);
    applyFilters();
    renderModalUi();
  });

  multiOnlyBtn.addEventListener("click", () => {
    setMultiDayFrequencyVisibilityByPredicate(entry => !!entry.isMulti);
    applyFilters();
    renderModalUi();
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

  multiDayManagerUiRenderer = () => {
    if (!isModalOpen()) return;
    renderModalUi();
  };

  setMultiDayServiceEnabled(false);
  refreshMultiDayManagerUi();
}

const ROUTE_SEQUENCER_DEPOTS_KEY = "routeSequencerDepots";
const ROUTE_SEQUENCER_CUSTOM_DUMPS_KEY = "routeSequencerCustomDumps";
const ROUTE_SEQUENCER_FACILITIES_KEY = "routeSequencerFacilities";
const ROUTE_SEQUENCER_FACILITY_ROLE_START_END = "start_end";
const ROUTE_SEQUENCER_FACILITY_ROLE_DUMP = "dump";
const ROUTE_SEQUENCER_FACILITY_ROLE_START_END_ONLY = ROUTE_SEQUENCER_FACILITY_ROLE_START_END; // Legacy alias
const ROUTE_SEQUENCER_FACILITY_ROLE_START_END_AND_INTERMEDIATE = "start_end_and_intermediate"; // Legacy value
const ROUTE_SEQUENCER_SETTINGS_KEY = "routeSequencerSettings";
const ROUTE_SEQUENCER_NODE_CELL_DEG = 0.02;
const ROUTE_SEQUENCER_MAX_NODE_SEARCH_RADIUS = 6;
const ROUTE_SEQUENCER_MAX_SNAP_MILES = 1.0;
const ROUTE_SEQUENCER_DEFAULT_CAPACITY_TONS = 13.5;
const ROUTE_SEQUENCER_DEFAULT_SPEED_MPH = 35;
const ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH = 30;
const ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD = "speed_cat";
const ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD = true;
const ROUTE_SEQUENCER_REQUIRE_NETWORK_ONLY = true;
const ROUTE_SEQUENCER_MEANDER_OVERRIDE_KEEP = "keep_meander";
const ROUTE_SEQUENCER_MEANDER_OVERRIDE_SAME_SIDE = "same_side";
const ROUTE_SEQUENCER_DEFAULT_MEANDER_OVERRIDE_DEFAULTS = Object.freeze({
  driving: ROUTE_SEQUENCER_MEANDER_OVERRIDE_KEEP,
  walking: ROUTE_SEQUENCER_MEANDER_OVERRIDE_KEEP
});
const ROUTE_SEQUENCER_SPEED_DEFAULTS_VERSION = 2;
const ROUTE_SEQUENCER_STREET_ATTR_DEFAULTS_VERSION = 2;
const ROUTE_SEQUENCER_SPEED_CAT_IDS = Object.freeze([1, 2, 3, 4, 5, 6, 7, 8]);
const ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP = Object.freeze({
  1: 65,
  2: 65,
  3: 55,
  4: 45,
  5: 35,
  6: 25,
  7: 15,
  8: 5
});
const ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP = Object.freeze({
  1: 50,
  2: 50,
  3: 40,
  4: 30,
  5: 15,
  6: 7,
  7: 7,
  8: 5
});
const ROUTE_SEQUENCER_DEFAULT_CLASS_SPEEDS = {
  ...ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
}; // Legacy alias.
const ROUTE_SEQUENCER_STREET_CLASS_IDS = Object.freeze([1, 2, 3, 4, 5, 6, 7, 8]);
const ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP = Object.freeze({
  1: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Same", crossPenaltyMinutes: 5, uTurnsRestricted: true }),
  2: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Same", crossPenaltyMinutes: 2.5, uTurnsRestricted: true }),
  3: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Same", crossPenaltyMinutes: 0.25, uTurnsRestricted: true }),
  4: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Both", crossPenaltyMinutes: 0, uTurnsRestricted: false }),
  5: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Both", crossPenaltyMinutes: 0, uTurnsRestricted: false }),
  6: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Both", crossPenaltyMinutes: 0, uTurnsRestricted: false }),
  7: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Both", crossPenaltyMinutes: 0, uTurnsRestricted: false }),
  8: Object.freeze({ travelTypesAllowed: "Any", serviceSide: "Both", crossPenaltyMinutes: 0, uTurnsRestricted: false })
});
const ROUTE_SEQUENCER_DEFAULT_STREET_ATTRIBUTE_DEFAULTS = Object.freeze({
  streetClassField: "func_class",
  allowUTurnsOnlyInCuldesacs: false,
  discourageUTurnsOutsideIntersections: true,
  meanderServiceMinutes: 0.5,
  sameSideServiceMinutes: 1,
  uTurnProrateDriving: "none",
  uTurnProrateWalking: "none",
  classAttrMap: ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP
});
const ROUTE_SEQUENCER_UTURN_RESTRICTED_PENALTY_MINUTES = 6;
const ROUTE_SEQUENCER_UTURN_CULDESAC_ONLY_PENALTY_MINUTES = 8;
const ROUTE_SEQUENCER_UTURN_PRORATE_DRIVING_FACTOR = 0.15;
const ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS = Object.freeze({
  sequence: "SEQNO",
  trip: "TRIP",
  cumulativeLoad: "CUM_TONS",
  driveMinutes: "DRIVE_MIN",
  serviceMinutes: "SERVICE_MIN",
  totalMinutes: "TOTAL_MIN",
  dumpVisits: "DUMP_VISITS"
});
let routeSequencerActiveOutputFields = { ...ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS };

function getRouteSequencerFacilityRoleLabel(roleValue) {
  return normalizeRouteSequencerFacilityRole(roleValue) === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP
    ? "Dump Location"
    : "Start/End";
}

function isProtectedNewPlanningFieldName(fieldName) {
  const token = normalizeRouteSequencerFieldToken(fieldName);
  return token === "newroute" || token === "newday" || token === "newseq" || token === "newfreq";
}

function normalizeRouteSequencerOutputFieldName(value, fallback = "") {
  const text = String(value ?? "").trim();
  if (text && isProtectedNewPlanningFieldName(text)) {
    return String(fallback || "").trim();
  }
  return text || String(fallback || "").trim();
}

function normalizeRouteSequencerOutputFields(rawFields) {
  const raw = rawFields && typeof rawFields === "object" ? rawFields : {};
  return {
    sequence: normalizeRouteSequencerOutputFieldName(raw.sequence, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.sequence),
    trip: normalizeRouteSequencerOutputFieldName(raw.trip, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.trip),
    cumulativeLoad: normalizeRouteSequencerOutputFieldName(raw.cumulativeLoad, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.cumulativeLoad),
    driveMinutes: normalizeRouteSequencerOutputFieldName(raw.driveMinutes, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.driveMinutes),
    serviceMinutes: normalizeRouteSequencerOutputFieldName(raw.serviceMinutes, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.serviceMinutes),
    totalMinutes: normalizeRouteSequencerOutputFieldName(raw.totalMinutes, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.totalMinutes),
    dumpVisits: normalizeRouteSequencerOutputFieldName(raw.dumpVisits, ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS.dumpVisits)
  };
}

function setRouteSequencerActiveOutputFields(outputFields) {
  routeSequencerActiveOutputFields = normalizeRouteSequencerOutputFields(outputFields);
  return routeSequencerActiveOutputFields;
}

function getRouteSequencerActiveOutputFields() {
  return normalizeRouteSequencerOutputFields(routeSequencerActiveOutputFields);
}

function buildRouteSequencerFormatHintText(demandField, serviceField, outputFields, sequenceIncrement = 1) {
  const demandLabel = String(demandField || "").trim() || "Demand";
  const serviceLabel = String(serviceField || "").trim() || "Service Time";
  const mapped = normalizeRouteSequencerOutputFields(outputFields);
  const increment = normalizeRouteSequencerSequenceIncrement(sequenceIncrement, 1);
  const outputList = [
    mapped.sequence,
    mapped.trip,
    mapped.cumulativeLoad,
    mapped.driveMinutes,
    mapped.serviceMinutes,
    mapped.totalMinutes,
    mapped.dumpVisits
  ].join(", ");
  return `${demandLabel}: use numeric tons (example: 0.5, 3, 12.75). ${serviceLabel}: use minutes as number (2, 7.5) or clock text (MM:SS or HH:MM:SS, example 05:30). Speed class field defaults to speed_cat. Stop-to-stop travel uses Service Speed; deadhead travel to dumps/start/end uses Driving Speed. Service Stops Before Deadhead can prioritize servicing additional stops before deadheading to dump when capacity allows. Street Attribute Defaults apply class-based U-turn restrictions, meander/same-side penalties, and cross-penalty time. Driver breaks: optional comma-separated MM:SS values (example 15:00, 30:00). Sequencing is solved per Route+Day as a single route. Time at Start/End is route-level time in MM:SS. Sequence increment: ${increment}. Evaluate Sequence keeps stop order from existing sequence values while still applying capacity, dump, depot, trip, and metric logic. Auto Sequence overwrites output fields: ${outputList}.`;
}

function normalizeRouteSequencerText(value) {
  return String(value ?? "").trim();
}

function normalizeRouteSequencerName(value, fallback = "Location") {
  const text = normalizeRouteSequencerText(value).slice(0, 120);
  return text || fallback;
}

function normalizeRouteSequencerPointEntry(raw, fallbackName = "Location") {
  if (!raw || typeof raw !== "object") return null;
  const lat = Number(raw.lat);
  const lon = Number(raw.lon);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
  const id = normalizeRouteSequencerText(raw.id) || `seq_loc_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  return {
    id,
    name: normalizeRouteSequencerName(raw.name, fallbackName),
    lat,
    lon
  };
}

function loadRouteSequencerPointList(storageKeyName, fallbackName) {
  const raw = storageGet(storageKeyName);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map(item => normalizeRouteSequencerPointEntry(item, fallbackName))
      .filter(Boolean);
  } catch {
    return [];
  }
}

function saveRouteSequencerPointList(storageKeyName, list, fallbackName) {
  const cleaned = (Array.isArray(list) ? list : [])
    .map(item => normalizeRouteSequencerPointEntry(item, fallbackName))
    .filter(Boolean);
  storageSet(storageKeyName, JSON.stringify(cleaned));
  return cleaned;
}

function createRouteSequencerFacilityMarkerIcon(roleValue) {
  const role = normalizeRouteSequencerFacilityRole(roleValue);
  const isDump = role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP;
  const pinFill = isDump ? "#84542d" : "#2d5f9b";
  const coreFill = isDump ? "#6a7f3a" : "#4fa3ff";
  const accentFill = isDump ? "#d8e2ba" : "#d8ecff";
  const html = `
    <span class="route-sequencer-facility-marker-wrap ${isDump ? "dump" : "start-end"}" aria-hidden="true">
      <svg class="route-sequencer-facility-marker-svg" viewBox="0 0 36 46">
        <path class="route-sequencer-facility-marker-pin" d="M18 1.9c-7.7 0-13.9 6.2-13.9 13.9 0 10.7 10.6 19 12.8 28.5.3 1.2 1.9 1.2 2.2 0 2.2-9.5 12.8-17.8 12.8-28.5C31.9 8.1 25.7 1.9 18 1.9z" fill="${pinFill}" />
        <circle cx="18" cy="15.1" r="9.2" fill="${coreFill}" />
        ${isDump
          ? `<rect x="12.1" y="10.7" width="11.8" height="8.6" rx="1.2" fill="${accentFill}" />
             <rect x="13.8" y="9.3" width="8.3" height="1.5" rx="0.7" fill="${accentFill}" />
             <rect x="15" y="11.8" width="1.3" height="5.6" rx="0.5" fill="${coreFill}" />
             <rect x="17.4" y="11.8" width="1.3" height="5.6" rx="0.5" fill="${coreFill}" />
             <rect x="19.8" y="11.8" width="1.3" height="5.6" rx="0.5" fill="${coreFill}" />`
          : `<path d="M18 9.4l5.1 4.2v6h-3.1v-4h-4v4H13v-6z" fill="${accentFill}" />
             <rect x="17.2" y="15.8" width="1.7" height="3.8" rx="0.5" fill="${coreFill}" />`
        }
      </svg>
    </span>
  `;
  return L.divIcon({
    className: "route-sequencer-facility-marker-icon",
    html,
    iconSize: [32, 42],
    iconAnchor: [16, 40],
    popupAnchor: [0, -34]
  });
}

function buildRouteSequencerFacilityPopupContent(facility) {
  const role = normalizeRouteSequencerFacilityRole(facility?.role);
  const roleLabel = role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP ? "Dump Location" : "Start/End Facility";
  const lat = Number(facility?.lat);
  const lon = Number(facility?.lon);
  const dumpTime = role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP
    ? formatRouteSequencerMinutesClock(facility?.facilityMinutes)
    : "";
  const rows = [
    `<div><strong>${escapeRouteSequencerHtml(facility?.name || "Facility")}</strong></div>`,
    `<div><strong>Role:</strong> ${escapeRouteSequencerHtml(roleLabel)}</div>`,
    `<div><strong>Latitude:</strong> ${Number.isFinite(lat) ? lat.toFixed(6) : "-"}</div>`,
    `<div><strong>Longitude:</strong> ${Number.isFinite(lon) ? lon.toFixed(6) : "-"}</div>`
  ];
  if (role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP) {
    rows.push(`<div><strong>Dump Time:</strong> ${escapeRouteSequencerHtml(dumpTime)}</div>`);
  } else {
    rows.push("<div><strong>Note:</strong> Uses solver Time at Start/End settings.</div>");
  }
  return `<div style="font-size:12px; line-height:1.35;">${rows.join("")}</div>`;
}

function syncRouteSequencerFacilitiesLayer() {
  routeSequencerFacilitiesLayerGroup.clearLayers();
  const facilities = getRouteSequencerFacilities();
  facilities.forEach(facility => {
    const lat = Number(facility?.lat);
    const lon = Number(facility?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    const marker = L.marker([lat, lon], {
      icon: createRouteSequencerFacilityMarkerIcon(facility?.role),
      keyboard: true,
      title: String(facility?.name || "Facility")
    });
    marker.bindPopup(buildRouteSequencerFacilityPopupContent(facility));
    routeSequencerFacilitiesLayerGroup.addLayer(marker);
  });
  if (!map.hasLayer(routeSequencerFacilitiesLayerGroup)) {
    routeSequencerFacilitiesLayerGroup.addTo(map);
  }
  try {
    routeSequencerFacilitiesLayerGroup.eachLayer(layer => layer.bringToFront?.());
  } catch {
    // Ignore marker z-order issues.
  }
}

function normalizeRouteSequencerFacilitiesList(list) {
  return (Array.isArray(list) ? list : [])
    .map(item => normalizeRouteSequencerFacilityEntry(item, "Facility"))
    .filter(Boolean);
}

function getRouteSequencerFacilityCloudSignature(item) {
  const normalized = normalizeRouteSequencerFacilityEntry(item, "Facility");
  if (!normalized) return "";
  const role = normalizeRouteSequencerFacilityRole(normalized.role);
  const lat = Number(normalized.lat);
  const lon = Number(normalized.lon);
  const latToken = Number.isFinite(lat) ? lat.toFixed(6) : "";
  const lonToken = Number.isFinite(lon) ? lon.toFixed(6) : "";
  const nameToken = String(normalized.name || "").trim().toLowerCase();
  return `${role}|${nameToken}|${latToken}|${lonToken}`;
}

function areRouteSequencerFacilityListsEquivalent(listA, listB) {
  const a = normalizeRouteSequencerFacilitiesList(listA);
  const b = normalizeRouteSequencerFacilitiesList(listB);
  if (a.length !== b.length) return false;
  const normalizeForCompare = list =>
    list
      .map(item => ({
        id: String(item.id || ""),
        signature: getRouteSequencerFacilityCloudSignature(item),
        facilityMinutes: Number(item.facilityMinutes) || 0
      }))
      .sort((left, right) => {
        const sigDelta = String(left.signature).localeCompare(String(right.signature), undefined, { sensitivity: "base", numeric: true });
        if (sigDelta !== 0) return sigDelta;
        const idDelta = String(left.id).localeCompare(String(right.id), undefined, { sensitivity: "base", numeric: true });
        if (idDelta !== 0) return idDelta;
        return left.facilityMinutes - right.facilityMinutes;
      });
  const left = normalizeForCompare(a);
  const right = normalizeForCompare(b);
  return left.every((item, index) =>
    item.id === right[index].id &&
    item.signature === right[index].signature &&
    item.facilityMinutes === right[index].facilityMinutes
  );
}

function mergeRouteSequencerFacilityLists(baseList, incomingList) {
  const merged = [];
  const indexById = new Map();
  const indexBySignature = new Map();

  const upsert = rawItem => {
    const item = normalizeRouteSequencerFacilityEntry(rawItem, "Facility");
    if (!item) return;
    const id = String(item.id || "").trim();
    const signature = getRouteSequencerFacilityCloudSignature(item);
    if (!signature) return;

    if (id && indexById.has(id)) {
      const idx = indexById.get(id);
      merged[idx] = item;
      indexBySignature.set(signature, idx);
      return;
    }

    if (indexBySignature.has(signature)) {
      const idx = indexBySignature.get(signature);
      const previous = merged[idx];
      const preservedId = String(previous?.id || id || "");
      const next = normalizeRouteSequencerFacilityEntry({
        ...previous,
        ...item,
        id: preservedId || item.id
      }, "Facility");
      if (!next) return;
      merged[idx] = next;
      if (preservedId) indexById.set(preservedId, idx);
      indexBySignature.set(signature, idx);
      return;
    }

    const idx = merged.length;
    merged.push(item);
    if (id) indexById.set(id, idx);
    indexBySignature.set(signature, idx);
  };

  normalizeRouteSequencerFacilitiesList(baseList).forEach(upsert);
  normalizeRouteSequencerFacilitiesList(incomingList).forEach(upsert);

  return merged.sort((a, b) =>
    String(a?.name || "").localeCompare(String(b?.name || ""), undefined, { sensitivity: "base", numeric: true })
  );
}

function buildSharedFacilitiesWorkbookArray(facilities) {
  const workbook = XLSX.utils.book_new();
  const rows = normalizeRouteSequencerFacilitiesList(facilities).map(item => ({
    id: String(item.id || ""),
    name: String(item.name || ""),
    role: normalizeRouteSequencerFacilityRole(item.role),
    facilityMinutes: Number(item.facilityMinutes) || 0,
    lat: Number(item.lat),
    lon: Number(item.lon)
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows, {
    header: ["id", "name", "role", "facilityMinutes", "lat", "lon"]
  });
  XLSX.utils.book_append_sheet(workbook, worksheet, SHARED_FACILITIES_SHEET_NAME);
  return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
}

function parseSharedFacilitiesWorkbookArrayBuffer(buffer) {
  const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
  const preferredSheetName = workbook.SheetNames.includes(SHARED_FACILITIES_SHEET_NAME)
    ? SHARED_FACILITIES_SHEET_NAME
    : workbook.SheetNames[0];
  if (!preferredSheetName) return [];
  const worksheet = workbook.Sheets[preferredSheetName];
  if (!worksheet) return [];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  return rows
    .map(row => normalizeRouteSequencerFacilityEntry({
      id: row?.id,
      name: row?.name,
      role: row?.role,
      facilityMinutes: row?.facilityMinutes ?? row?.serviceMinutes,
      lat: row?.lat ?? row?.latitude,
      lon: row?.lon ?? row?.longitude
    }, "Facility"))
    .filter(Boolean);
}

async function fetchSharedFacilitiesFromCloud() {
  const { data } = sb.storage.from(BUCKET).getPublicUrl(SHARED_FACILITIES_CLOUD_FILE);
  const publicUrl = String(data?.publicUrl || "").trim();
  if (!publicUrl) {
    throw new Error("Missing public URL for shared facilities file.");
  }
  const response = await fetch(`${publicUrl}?v=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) {
    const responseText = await response.text().catch(() => "");
    const looksMissing = response.status === 404 ||
      ((response.status === 400 || response.status === 403) && /not found|does not exist|no such object|not authorized/i.test(responseText));
    if (looksMissing) {
      return { exists: false, facilities: [] };
    }
    throw new Error(`Shared facilities fetch failed (${response.status}).`);
  }
  const buffer = await response.arrayBuffer();
  return {
    exists: true,
    facilities: parseSharedFacilitiesWorkbookArrayBuffer(buffer)
  };
}

async function saveSharedFacilitiesToCloud(options = {}) {
  const mergeRemote = options.mergeRemote !== false;
  const syncLocalFromSaved = options.syncLocalFromSaved !== false;
  const facilitiesOverride = options.facilitiesOverride;
  let facilitiesToSave = normalizeRouteSequencerFacilitiesList(
    facilitiesOverride == null ? getRouteSequencerFacilities() : facilitiesOverride
  );

  if (mergeRemote) {
    try {
      const remoteState = await fetchSharedFacilitiesFromCloud();
      if (remoteState.exists) {
        facilitiesToSave = mergeRouteSequencerFacilityLists(remoteState.facilities, facilitiesToSave);
      }
    } catch (error) {
      console.warn("Unable to merge remote shared facilities before save:", error);
    }
  }

  const workbookArray = buildSharedFacilitiesWorkbookArray(facilitiesToSave);
  const { error } = await sb.storage
    .from(BUCKET)
    .upload(SHARED_FACILITIES_CLOUD_FILE, workbookArray, {
      upsert: true,
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
  if (error) throw error;

  if (syncLocalFromSaved) {
    const localFacilities = getRouteSequencerFacilities();
    if (!areRouteSequencerFacilityListsEquivalent(localFacilities, facilitiesToSave)) {
      setRouteSequencerFacilities(facilitiesToSave, { queueCloudSync: false });
    }
  }

  return facilitiesToSave;
}

async function flushQueuedSharedFacilitiesCloudSave() {
  if (!sharedFacilitiesAutoSavePendingReasons.length) return false;
  if (sharedFacilitiesAutoSaveInFlight) return sharedFacilitiesAutoSaveInFlight;

  const reasons = [...sharedFacilitiesAutoSavePendingReasons];
  sharedFacilitiesAutoSavePendingReasons = [];
  const reasonLabel = reasons.join(", ");

  sharedFacilitiesAutoSaveInFlight = (async () => {
    try {
      await initSharedFacilitiesCloudSync();
      await saveSharedFacilitiesToCloud({ mergeRemote: true });
      console.info(
        `[Shared Facilities Auto Save] Saved "${SHARED_FACILITIES_CLOUD_FILE}"${reasonLabel ? ` (${reasonLabel})` : ""}.`
      );
      return true;
    } catch (error) {
      console.error(
        `[Shared Facilities Auto Save] Error saving "${SHARED_FACILITIES_CLOUD_FILE}"${reasonLabel ? ` (${reasonLabel})` : ""}:`,
        error
      );
      return false;
    } finally {
      sharedFacilitiesAutoSaveInFlight = null;
      if (sharedFacilitiesAutoSavePendingReasons.length) {
        if (sharedFacilitiesAutoSaveTimer) clearTimeout(sharedFacilitiesAutoSaveTimer);
        sharedFacilitiesAutoSaveTimer = setTimeout(() => {
          sharedFacilitiesAutoSaveTimer = null;
          flushQueuedSharedFacilitiesCloudSave();
        }, 300);
      }
    }
  })();

  return sharedFacilitiesAutoSaveInFlight;
}

function queueSharedFacilitiesCloudSave(reason = "facilities-update") {
  const reasonText = String(reason || "").trim();
  if (reasonText) sharedFacilitiesAutoSavePendingReasons.push(reasonText);
  if (sharedFacilitiesAutoSaveTimer) clearTimeout(sharedFacilitiesAutoSaveTimer);
  sharedFacilitiesAutoSaveTimer = setTimeout(() => {
    sharedFacilitiesAutoSaveTimer = null;
    flushQueuedSharedFacilitiesCloudSave();
  }, SHARED_FACILITIES_CLOUD_SAVE_DEBOUNCE_MS);
}

function flushSharedFacilitiesCloudSaveOnLifecycle(force = false) {
  if (!force && !sharedFacilitiesAutoSavePendingReasons.length && !sharedFacilitiesAutoSaveTimer) return;
  if (sharedFacilitiesAutoSaveTimer) {
    clearTimeout(sharedFacilitiesAutoSaveTimer);
    sharedFacilitiesAutoSaveTimer = null;
  }
  if (!sharedFacilitiesAutoSavePendingReasons.length) {
    sharedFacilitiesAutoSavePendingReasons.push("lifecycle");
  }
  flushQueuedSharedFacilitiesCloudSave();
}

async function initSharedFacilitiesCloudSync() {
  if (sharedFacilitiesCloudInitPromise) return sharedFacilitiesCloudInitPromise;
  sharedFacilitiesCloudInitPromise = (async () => {
    try {
      const localFacilities = getRouteSequencerFacilities();
      const remoteState = await fetchSharedFacilitiesFromCloud();

      if (!remoteState.exists) {
        await saveSharedFacilitiesToCloud({
          mergeRemote: false,
          syncLocalFromSaved: false,
          facilitiesOverride: localFacilities
        });
        return;
      }

      const remoteFacilities = normalizeRouteSequencerFacilitiesList(remoteState.facilities);
      const mergedFacilities = mergeRouteSequencerFacilityLists(remoteFacilities, localFacilities);
      if (!areRouteSequencerFacilityListsEquivalent(localFacilities, mergedFacilities)) {
        setRouteSequencerFacilities(mergedFacilities, { queueCloudSync: false });
      }
      if (!areRouteSequencerFacilityListsEquivalent(remoteFacilities, mergedFacilities)) {
        await saveSharedFacilitiesToCloud({
          mergeRemote: false,
          syncLocalFromSaved: false,
          facilitiesOverride: mergedFacilities
        });
      }
    } catch (error) {
      console.warn("Shared facilities cloud sync init failed:", error);
    }
  })();
  return sharedFacilitiesCloudInitPromise;
}

function setRouteSequencerFacilities(list, options = {}) {
  const cleaned = normalizeRouteSequencerFacilitiesList(list);
  storageSet(ROUTE_SEQUENCER_FACILITIES_KEY, JSON.stringify(cleaned));
  syncRouteSequencerFacilitiesLayer();
  if (options?.queueCloudSync !== false) {
    queueSharedFacilitiesCloudSave(String(options?.reason || "Facilities Update"));
  }
  return cleaned;
}

function loadRouteSequencerFacilities() {
  const raw = storageGet(ROUTE_SEQUENCER_FACILITIES_KEY);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return parsed
          .map(item => normalizeRouteSequencerFacilityEntry(item, "Facility"))
          .filter(Boolean);
      }
    } catch {
      // Ignore malformed storage.
    }
  }

  // Backward-compatibility migration from old depot/custom-dump lists.
  const legacyDepots = loadRouteSequencerPointList(ROUTE_SEQUENCER_DEPOTS_KEY, "Depot");
  const legacyDumps = loadRouteSequencerPointList(ROUTE_SEQUENCER_CUSTOM_DUMPS_KEY, "Dump");
  if (!legacyDepots.length && !legacyDumps.length) return [];

  const byId = new Map();
  const nextUniqueId = prefix => `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  legacyDepots.forEach(item => {
    const normalized = normalizeRouteSequencerFacilityEntry({
      ...item,
      role: ROUTE_SEQUENCER_FACILITY_ROLE_START_END,
      facilityMinutes: 0
    }, "Depot");
    if (normalized) byId.set(String(normalized.id), normalized);
  });
  legacyDumps.forEach(item => {
    const normalized = normalizeRouteSequencerFacilityEntry({
      ...item,
      id: byId.has(String(item?.id || "")) ? nextUniqueId("dump") : item?.id,
      role: ROUTE_SEQUENCER_FACILITY_ROLE_DUMP,
      facilityMinutes: item?.facilityMinutes ?? 0
    }, "Dump");
    if (normalized) byId.set(String(normalized.id), normalized);
  });

  const migrated = [...byId.values()];
  if (migrated.length) setRouteSequencerFacilities(migrated, { queueCloudSync: false });
  return migrated;
}

function getRouteSequencerFacilities() {
  return loadRouteSequencerFacilities();
}

function getRouteSequencerDepots() {
  return getRouteSequencerFacilities()
    .filter(item => normalizeRouteSequencerFacilityRole(item?.role) === ROUTE_SEQUENCER_FACILITY_ROLE_START_END)
    .map(item => ({
    id: String(item.id),
    name: String(item.name),
    lat: Number(item.lat),
    lon: Number(item.lon),
    facilityMinutes: 0,
    role: ROUTE_SEQUENCER_FACILITY_ROLE_START_END,
    allowIntermediate: false
  }));
}

function setRouteSequencerDepots(list) {
  const existing = getRouteSequencerFacilities();
  const existingById = new Map(existing.map(item => [String(item.id), item]));
  const existingDumps = existing.filter(
    item => normalizeRouteSequencerFacilityRole(item?.role) === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP
  );
  const depots = (Array.isArray(list) ? list : [])
    .map(item => normalizeRouteSequencerPointEntry(item, "Depot"))
    .filter(Boolean)
    .map(item => {
      const existing = existingById.get(String(item.id));
      return normalizeRouteSequencerFacilityEntry({
        ...item,
        role: ROUTE_SEQUENCER_FACILITY_ROLE_START_END,
        facilityMinutes: 0,
        id: existing?.id || item.id
      }, "Depot");
    })
    .filter(Boolean);
  return setRouteSequencerFacilities([...existingDumps, ...depots]);
}

function getRouteSequencerCustomDumps() {
  return getRouteSequencerFacilities()
    .filter(item => normalizeRouteSequencerFacilityRole(item?.role) === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP)
    .map(item => ({
      id: String(item.id),
      name: String(item.name),
      lat: Number(item.lat),
      lon: Number(item.lon),
      facilityMinutes: Number(item.facilityMinutes) || 0
    }));
}

function setRouteSequencerCustomDumps(list) {
  const existing = getRouteSequencerFacilities();
  const depots = existing.filter(
    item => normalizeRouteSequencerFacilityRole(item?.role) === ROUTE_SEQUENCER_FACILITY_ROLE_START_END
  );
  const existingDumpById = new Map(
    existing
      .filter(item => normalizeRouteSequencerFacilityRole(item?.role) === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP)
      .map(item => [String(item.id), item])
  );
  const incoming = (
    (Array.isArray(list) ? list : [])
      .map(item => normalizeRouteSequencerPointEntry(item, "Dump"))
      .filter(Boolean)
      .map(item => {
        const existingDump = existingDumpById.get(String(item.id));
        return normalizeRouteSequencerFacilityEntry({
          ...item,
          role: ROUTE_SEQUENCER_FACILITY_ROLE_DUMP,
          facilityMinutes: item?.facilityMinutes ?? existingDump?.facilityMinutes ?? 0
        }, "Dump");
      })
      .filter(Boolean)
  );
  return setRouteSequencerFacilities([...depots, ...incoming]);
}

function normalizeRouteSequencerFieldToken(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function getRouteSequencerFieldNames() {
  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  const fieldSet = new Set();
  rows.forEach(row => {
    if (!row || typeof row !== "object") return;
    Object.keys(row).forEach(key => {
      const text = String(key || "").trim();
      if (text) fieldSet.add(text);
    });
  });
  return [...fieldSet].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base", numeric: true }));
}

function guessRouteSequencerFieldByAliases(fieldNames, aliases = []) {
  const options = Array.isArray(fieldNames) ? fieldNames : [];
  if (!options.length) return "";
  const normalizedAliases = (Array.isArray(aliases) ? aliases : [])
    .map(alias => normalizeRouteSequencerFieldToken(alias))
    .filter(Boolean);
  if (!normalizedAliases.length) return "";

  const normalizedFieldEntries = options.map(name => ({
    name,
    token: normalizeRouteSequencerFieldToken(name)
  }));

  for (const alias of normalizedAliases) {
    const direct = normalizedFieldEntries.find(item => item.token === alias);
    if (direct) return direct.name;
  }
  for (const alias of normalizedAliases) {
    const include = normalizedFieldEntries.find(item => item.token.includes(alias) || alias.includes(item.token));
    if (include) return include.name;
  }
  return "";
}

function parseRouteSequencerNumeric(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value).trim();
  if (!text) return fallback;
  const cleaned = text.replace(/,/g, "");
  const direct = Number(cleaned);
  if (Number.isFinite(direct)) return direct;
  const match = cleaned.match(/-?\d+(?:\.\d+)?/);
  if (!match) return fallback;
  const parsed = Number(match[0]);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseRouteSequencerMinutesValue(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  if (typeof value === "number" && Number.isFinite(value)) return Math.max(0, value);

  const text = String(value).trim();
  if (!text) return fallback;
  if (text.includes(":")) {
    const parts = text.split(":").map(part => Number(String(part).trim()));
    if (parts.every(Number.isFinite)) {
      if (parts.length === 2) {
        const [minutes, seconds] = parts;
        if (minutes >= 0 && seconds >= 0) return Math.max(0, minutes + (seconds / 60));
      } else if (parts.length === 3) {
        const [hours, minutes, seconds] = parts;
        if (hours >= 0 && minutes >= 0 && seconds >= 0) {
          return Math.max(0, (hours * 60) + minutes + (seconds / 60));
        }
      }
    }
  }

  const numeric = parseRouteSequencerNumeric(text, fallback);
  return Number.isFinite(numeric) ? Math.max(0, numeric) : fallback;
}

function formatRouteSequencerMinutesClock(minutesValue) {
  const minutes = Math.max(0, Number(minutesValue) || 0);
  const totalSeconds = Math.round(minutes * 60);
  const hours = Math.floor(totalSeconds / 3600);
  const minutesRemainder = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  if (hours > 0) {
    return `${hours}:${String(minutesRemainder).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }
  return `${minutesRemainder}:${String(seconds).padStart(2, "0")}`;
}

function parseRouteSequencerBreakMinutesList(value) {
  if (Array.isArray(value)) {
    return value
      .map(item => parseRouteSequencerMinutesValue(item, NaN))
      .filter(minutes => Number.isFinite(minutes) && minutes > 0);
  }

  const text = String(value ?? "").trim();
  if (!text) return [];

  return text
    .split(/[\n,;|]+/g)
    .map(token => token.trim())
    .filter(Boolean)
    .map(token => parseRouteSequencerMinutesValue(token, NaN))
    .filter(minutes => Number.isFinite(minutes) && minutes > 0);
}

function formatRouteSequencerBreakMinutesList(value) {
  const list = parseRouteSequencerBreakMinutesList(value);
  if (!list.length) return "";
  return list.map(minutes => formatRouteSequencerMinutesClock(minutes)).join(", ");
}

function parseRouteSequencerClassId(value, fallback = null) {
  if (value == null || value === "") return fallback;
  if (typeof value === "number" && Number.isFinite(value)) {
    const rounded = Math.round(value);
    if (rounded >= 1 && rounded <= 99) return Math.max(1, Math.min(8, rounded));
    return fallback;
  }
  const text = String(value).trim();
  if (!text) return fallback;
  const match = text.match(/\d+/);
  if (!match) return fallback;
  const parsed = Number(match[0]);
  if (!Number.isFinite(parsed) || parsed < 1 || parsed > 99) return fallback;
  return Math.max(1, Math.min(8, Math.round(parsed)));
}

function normalizeRouteSequencerTravelTypesAllowed(value) {
  const token = normalizeRouteSequencerFieldToken(value);
  if (token === "driving" || token === "vehicle" || token === "car" || token === "truck") return "Driving";
  if (token === "walking" || token === "walk" || token === "pedestrian") return "Walking";
  return "Any";
}

function normalizeRouteSequencerServiceSideMode(value) {
  const token = normalizeRouteSequencerFieldToken(value);
  return token === "same" || token === "sameside" ? "Same" : "Both";
}

function normalizeRouteSequencerUTurnProrateMode(value) {
  const token = normalizeRouteSequencerFieldToken(value);
  if (token === "distance" || token === "prorate" || token === "yes" || token === "true") return "distance";
  return "none";
}

function normalizeRouteSequencerMeanderOverrideMode(value) {
  const token = normalizeRouteSequencerFieldToken(value);
  if (
    token === "same" ||
    token === "sameside" ||
    token === "same_side" ||
    token === "forcesameside" ||
    token === "override"
  ) {
    return ROUTE_SEQUENCER_MEANDER_OVERRIDE_SAME_SIDE;
  }
  return ROUTE_SEQUENCER_MEANDER_OVERRIDE_KEEP;
}

function normalizeRouteSequencerMeanderOverrideDefaults(rawValue) {
  const source = rawValue && typeof rawValue === "object" ? rawValue : {};
  const defaults = ROUTE_SEQUENCER_DEFAULT_MEANDER_OVERRIDE_DEFAULTS;
  return {
    driving: normalizeRouteSequencerMeanderOverrideMode(source.driving ?? source.vehicle ?? defaults.driving),
    walking: normalizeRouteSequencerMeanderOverrideMode(source.walking ?? source.pedestrian ?? defaults.walking)
  };
}

function normalizeRouteSequencerStreetClassAttrEntry(rawEntry, fallbackEntry = {}) {
  const fallback = fallbackEntry && typeof fallbackEntry === "object" ? fallbackEntry : {};
  const source = rawEntry && typeof rawEntry === "object" ? rawEntry : {};
  return {
    travelTypesAllowed: normalizeRouteSequencerTravelTypesAllowed(
      source.travelTypesAllowed ?? source.travelType ?? fallback.travelTypesAllowed
    ),
    serviceSide: normalizeRouteSequencerServiceSideMode(
      source.serviceSide ?? source.serviceSideMode ?? fallback.serviceSide
    ),
    crossPenaltyMinutes: Math.max(
      0,
      parseRouteSequencerMinutesValue(
        source.crossPenaltyMinutes ?? source.crossPenalty ?? fallback.crossPenaltyMinutes,
        Math.max(0, Number(fallback.crossPenaltyMinutes) || 0)
      )
    ),
    uTurnsRestricted: source.uTurnsRestricted == null
      ? !!fallback.uTurnsRestricted
      : !!source.uTurnsRestricted
  };
}

function normalizeRouteSequencerStreetClassAttrMap(rawMap) {
  const source = rawMap && typeof rawMap === "object" ? rawMap : {};
  const out = {};
  ROUTE_SEQUENCER_STREET_CLASS_IDS.forEach(classId => {
    const fallbackEntry = ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP[classId]
      || ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP[5];
    out[classId] = normalizeRouteSequencerStreetClassAttrEntry(
      source[classId] || source[String(classId)],
      fallbackEntry
    );
  });
  return out;
}

function normalizeRouteSequencerStreetAttributeDefaults(rawValue) {
  const source = rawValue && typeof rawValue === "object" ? rawValue : {};
  const classAttrSource = source.classAttrMap || source.streetClassAttributes || source.classAttributes || {};
  const defaults = ROUTE_SEQUENCER_DEFAULT_STREET_ATTRIBUTE_DEFAULTS;
  return {
    streetClassField: String(source.streetClassField || defaults.streetClassField || "func_class").trim() || "func_class",
    allowUTurnsOnlyInCuldesacs: !!source.allowUTurnsOnlyInCuldesacs,
    discourageUTurnsOutsideIntersections: source.discourageUTurnsOutsideIntersections !== false,
    meanderServiceMinutes: Math.max(
      0,
      parseRouteSequencerMinutesValue(source.meanderServiceMinutes, Number(defaults.meanderServiceMinutes) || 0)
    ),
    sameSideServiceMinutes: Math.max(
      0,
      parseRouteSequencerMinutesValue(source.sameSideServiceMinutes, Number(defaults.sameSideServiceMinutes) || 0)
    ),
    uTurnProrateDriving: normalizeRouteSequencerUTurnProrateMode(source.uTurnProrateDriving),
    uTurnProrateWalking: normalizeRouteSequencerUTurnProrateMode(source.uTurnProrateWalking),
    classAttrMap: normalizeRouteSequencerStreetClassAttrMap(classAttrSource)
  };
}

function getRouteSequencerStreetClassAttr(streetAttributeDefaults, classIdValue) {
  const normalized = normalizeRouteSequencerStreetAttributeDefaults(streetAttributeDefaults);
  const classId = parseRouteSequencerClassId(classIdValue, 5) || 5;
  return normalized.classAttrMap[classId] || normalized.classAttrMap[5] || ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP[5];
}

function normalizeRouteSequencerStopSide(value) {
  const token = normalizeRouteSequencerFieldToken(value);
  if (!token) return "";
  if (token.startsWith("left")) return "left";
  if (token.startsWith("right")) return "right";
  if (token.startsWith("center")) return "center";
  return "";
}

function resolveRouteSequencerStreetClassIdFromRow(row, streetClassField = "func_class") {
  const fieldName = String(streetClassField || "func_class").trim() || "func_class";
  const candidates = [
    getRouteSequencerRowFieldValue(row, fieldName),
    row?.func_class,
    row?.FUNC_CLASS,
    row?.speed_cat,
    row?.SPEED_CAT
  ];
  for (const candidate of candidates) {
    const parsed = parseRouteSequencerClassId(candidate, null);
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

function resolveRouteSequencerStopStreetMetadata(row, streetClassField = "func_class") {
  const directSegmentId = String(
    getSetSegmentSegmentIdString?.(row, "")
      || row?.SEGMENT_ID
      || row?.Segment_ID
      || row?.segment_id
      || ""
  ).trim();
  const numericSegmentId = parseRouteSequencerNumeric(directSegmentId, NaN);
  const streetEntry = Number.isFinite(numericSegmentId)
    ? streetAttributeById.get(Math.round(numericSegmentId))
    : null;
  const streetRow = streetEntry?.row || null;
  const streetClassId = resolveRouteSequencerStreetClassIdFromRow(streetRow || row, streetClassField);
  const sideRaw = String(
    getSetSegmentSegmentSideString?.(row, "")
      || row?.SEGMENT_SIDE
      || row?.Segment_Side
      || row?.segment_side
      || ""
  ).trim();
  return {
    segmentId: directSegmentId,
    segmentIdNumeric: Number.isFinite(numericSegmentId) ? Math.round(numericSegmentId) : NaN,
    segmentSide: sideRaw,
    streetClassId: Number.isFinite(streetClassId) ? streetClassId : null
  };
}

function normalizeRouteSequencerCandidateAdjustment(rawValue) {
  const source = rawValue && typeof rawValue === "object" ? rawValue : {};
  const notes = Array.isArray(source.notes)
    ? source.notes.map(value => String(value || "").trim()).filter(Boolean)
    : [];
  return {
    scoreMinutesDelta: Math.max(0, Number(source.scoreMinutesDelta) || 0),
    driveMinutesDelta: Math.max(0, Number(source.driveMinutesDelta) || 0),
    serviceMinutesDelta: Math.max(0, Number(source.serviceMinutesDelta) || 0),
    otherMinutesDelta: Math.max(0, Number(source.otherMinutesDelta) || 0),
    notes
  };
}

function buildRouteSequencerServiceCandidateAdjustment(fromPoint, candidate, baseMinutes, streetAttributeDefaults, meanderOverrideDefaults = null) {
  const defaults = normalizeRouteSequencerStreetAttributeDefaults(streetAttributeDefaults);
  const meanderOverrides = normalizeRouteSequencerMeanderOverrideDefaults(meanderOverrideDefaults);
  const adjustment = {
    scoreMinutesDelta: 0,
    driveMinutesDelta: 0,
    serviceMinutesDelta: 0,
    otherMinutesDelta: 0,
    notes: []
  };
  if (String(candidate?.type || "") !== "stop") return adjustment;

  const classAttr = getRouteSequencerStreetClassAttr(defaults, candidate?.streetClassId);
  const travelTypesAllowed = normalizeRouteSequencerTravelTypesAllowed(classAttr?.travelTypesAllowed);
  const activeMeanderOverride = travelTypesAllowed === "Walking"
    ? meanderOverrides.walking
    : meanderOverrides.driving;
  const effectiveServiceSide = activeMeanderOverride === ROUTE_SEQUENCER_MEANDER_OVERRIDE_SAME_SIDE
    ? "Same"
    : normalizeRouteSequencerServiceSideMode(classAttr?.serviceSide);
  if (travelTypesAllowed === "Walking") {
    adjustment.otherMinutesDelta += 30;
    adjustment.scoreMinutesDelta += 30;
    adjustment.notes.push("Travel type penalty (walking-only class)");
  }

  const fromSide = normalizeRouteSequencerStopSide(fromPoint?.segmentSide);
  const toSide = normalizeRouteSequencerStopSide(candidate?.segmentSide);
  const hasSidePair = !!fromSide && !!toSide && fromSide !== "center" && toSide !== "center";
  const oppositeServiceSide = hasSidePair && fromSide !== toSide;
  const fromSegmentId = String(fromPoint?.segmentId || "").trim();
  const toSegmentId = String(candidate?.segmentId || "").trim();
  const sameSegment = !!fromSegmentId && !!toSegmentId && fromSegmentId === toSegmentId;
  const likelyUTurn = oppositeServiceSide || sameSegment;

  if (oppositeServiceSide && effectiveServiceSide === "Same") {
    const sameSidePenalty = Math.max(0, Number(defaults.sameSideServiceMinutes) || 0);
    if (sameSidePenalty > 0) {
      adjustment.serviceMinutesDelta += sameSidePenalty;
      adjustment.scoreMinutesDelta += sameSidePenalty;
      adjustment.notes.push(
        activeMeanderOverride === ROUTE_SEQUENCER_MEANDER_OVERRIDE_SAME_SIDE
          ? "Same-side service penalty (meander override)"
          : "Same-side service penalty"
      );
    }
  }

  const crossPenaltyMinutes = Math.max(0, Number(classAttr?.crossPenaltyMinutes) || 0);
  if (oppositeServiceSide && crossPenaltyMinutes > 0) {
    adjustment.otherMinutesDelta += crossPenaltyMinutes;
    adjustment.scoreMinutesDelta += crossPenaltyMinutes;
    adjustment.notes.push("Cross penalty");
  }

  if (likelyUTurn) {
    if (classAttr?.uTurnsRestricted) {
      adjustment.otherMinutesDelta += ROUTE_SEQUENCER_UTURN_RESTRICTED_PENALTY_MINUTES;
      adjustment.scoreMinutesDelta += ROUTE_SEQUENCER_UTURN_RESTRICTED_PENALTY_MINUTES;
      adjustment.notes.push("U-turn restricted penalty");
    } else if (defaults.allowUTurnsOnlyInCuldesacs) {
      adjustment.otherMinutesDelta += ROUTE_SEQUENCER_UTURN_CULDESAC_ONLY_PENALTY_MINUTES;
      adjustment.scoreMinutesDelta += ROUTE_SEQUENCER_UTURN_CULDESAC_ONLY_PENALTY_MINUTES;
      adjustment.notes.push("U-turn cul-de-sac-only penalty");
    } else if (defaults.discourageUTurnsOutsideIntersections) {
      const meanderPenalty = Math.max(0, Number(defaults.meanderServiceMinutes) || 0);
      if (meanderPenalty > 0) {
        adjustment.serviceMinutesDelta += meanderPenalty;
        adjustment.scoreMinutesDelta += meanderPenalty;
        adjustment.notes.push("Meander service penalty");
      }
    }

    if (defaults.uTurnProrateDriving === "distance") {
      const base = Math.max(0, Number(baseMinutes) || 0);
      const prorateDelta = base * ROUTE_SEQUENCER_UTURN_PRORATE_DRIVING_FACTOR;
      if (prorateDelta > 0) {
        adjustment.driveMinutesDelta += prorateDelta;
        adjustment.scoreMinutesDelta += prorateDelta;
        adjustment.notes.push("U-turn driving prorate");
      }
    }
  }

  return adjustment;
}

function normalizeRouteSequencerFacilityRole(value) {
  const token = normalizeRouteSequencerFieldToken(value);
  if (
    token === normalizeRouteSequencerFieldToken(ROUTE_SEQUENCER_FACILITY_ROLE_START_END) ||
    token === normalizeRouteSequencerFieldToken(ROUTE_SEQUENCER_FACILITY_ROLE_START_END_ONLY) ||
    token === "startend" ||
    token === "startendonly" ||
    token === "depotonly"
  ) {
    return ROUTE_SEQUENCER_FACILITY_ROLE_START_END;
  }
  if (
    token === normalizeRouteSequencerFieldToken(ROUTE_SEQUENCER_FACILITY_ROLE_DUMP) ||
    token === normalizeRouteSequencerFieldToken(ROUTE_SEQUENCER_FACILITY_ROLE_START_END_AND_INTERMEDIATE) ||
    token === "startendandintermediate" ||
    token === "intermediate" ||
    token === "customdump" ||
    token === "dumponly"
  ) {
    return ROUTE_SEQUENCER_FACILITY_ROLE_DUMP;
  }
  return ROUTE_SEQUENCER_FACILITY_ROLE_START_END;
}

function normalizeRouteSequencerFacilityEntry(raw, fallbackName = "Facility") {
  if (!raw || typeof raw !== "object") return null;
  const lat = Number(raw.lat);
  const lon = Number(raw.lon);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
  const id = normalizeRouteSequencerText(raw.id) || `facility_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const role = normalizeRouteSequencerFacilityRole(
    raw.role != null
      ? raw.role
      : (raw.allowIntermediate ? ROUTE_SEQUENCER_FACILITY_ROLE_DUMP : ROUTE_SEQUENCER_FACILITY_ROLE_START_END)
  );
  const parsedFacilityMinutes = parseRouteSequencerMinutesValue(
    raw.facilityMinutes != null ? raw.facilityMinutes : raw.serviceMinutes,
    0
  );
  const facilityMinutes = role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP
    ? parsedFacilityMinutes
    : 0;
  return {
    id,
    name: normalizeRouteSequencerName(raw.name, fallbackName),
    lat,
    lon,
    role,
    allowIntermediate: role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP,
    facilityMinutes
  };
}

function clampRouteSequencerNumber(value, min, max, fallback) {
  const num = Number(value);
  if (!Number.isFinite(num)) return fallback;
  return Math.max(min, Math.min(max, num));
}

function haversineMiles(lat1, lon1, lat2, lon2) {
  const toRad = deg => (deg * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2
    + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  const earthMiles = 3958.7613;
  return earthMiles * c;
}

function normalizeRouteSequencerOneway(value) {
  const token = String(value ?? "").trim().toLowerCase();
  if (!token) return "both";
  if (token === "-1" || token === "reverse" || token === "backward") return "reverse";
  if (token === "yes" || token === "1" || token === "true" || token === "y" || token === "t" || token === "forward") return "forward";
  if (token === "no" || token === "0" || token === "false" || token === "n" || token === "f") return "both";
  return "both";
}

function parsePostedSpeedFromTag(maxspeedValue) {
  const text = String(maxspeedValue ?? "").trim().toLowerCase();
  if (!text) return NaN;
  if (text === "none" || text === "signals" || text === "walk") return NaN;

  const numbers = text.match(/-?\d+(?:\.\d+)?/g);
  if (!numbers || !numbers.length) return NaN;
  const first = Number(numbers[0]);
  if (!Number.isFinite(first) || first <= 0) return NaN;
  if (text.includes("km")) return first * 0.621371;
  return first;
}

function normalizeRouteSequencerSpeedClassField(value) {
  const text = String(value ?? "").trim();
  return text || ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD;
}

function normalizeRouteSequencerSpeedCatMap(rawMap, defaultMap) {
  const defaults = defaultMap && typeof defaultMap === "object"
    ? defaultMap
    : ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP;
  const out = {};
  ROUTE_SEQUENCER_SPEED_CAT_IDS.forEach(id => {
    const fallback = clampRouteSequencerNumber(defaults?.[id], 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH);
    out[id] = clampRouteSequencerNumber(rawMap?.[id], 5, 85, fallback);
  });
  return out;
}

function getRouteSequencerRowFieldValue(row, fieldName) {
  if (!row || typeof row !== "object") return "";
  const wantedKey = String(fieldName ?? "").trim();
  if (!wantedKey) return "";
  if (Object.prototype.hasOwnProperty.call(row, wantedKey)) return row[wantedKey];

  const token = normalizeRouteSequencerFieldToken(wantedKey);
  if (!token) return "";
  const matchedKey = Object.keys(row).find(key => normalizeRouteSequencerFieldToken(key) === token);
  if (!matchedKey) return "";
  return row[matchedKey];
}

function resolveRouteSequencerSpeedClassIdFromRow(row, speedClassField = ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD) {
  const fieldName = normalizeRouteSequencerSpeedClassField(speedClassField);
  const candidates = [
    getRouteSequencerRowFieldValue(row, fieldName),
    row?.speed_cat,
    row?.SPEED_CAT,
    row?.speedcat,
    row?.func_class,
    row?.FUNC_CLASS
  ];

  for (const candidate of candidates) {
    if (candidate == null || candidate === "") continue;
    if (typeof candidate === "number" && Number.isFinite(candidate)) {
      const numericId = Math.round(candidate);
      if (numericId >= 1 && numericId <= 99) {
        return Math.max(1, Math.min(8, numericId));
      }
    }
    const text = String(candidate).trim();
    if (!text) continue;
    const match = text.match(/\d+/);
    if (!match) continue;
    const parsed = Number(match[0]);
    if (Number.isFinite(parsed) && parsed >= 1 && parsed <= 99) {
      return Math.max(1, Math.min(8, Math.round(parsed)));
    }
  }
  return null;
}

function normalizeRouteSequencerTravelMode(value) {
  const token = String(value ?? "").trim().toLowerCase();
  return token === "service" ? "service" : "driving";
}

function resolveStreetSpeedsForSequencer(row, speedConfig, fallbackSpeedMph) {
  const hasConfigShape = speedConfig && typeof speedConfig === "object" && (
    Object.prototype.hasOwnProperty.call(speedConfig, "speedClassField")
    || Object.prototype.hasOwnProperty.call(speedConfig, "drivingSpeedCatMap")
    || Object.prototype.hasOwnProperty.call(speedConfig, "serviceSpeedCatMap")
    || Object.prototype.hasOwnProperty.call(speedConfig, "classSpeedMap")
    || Object.prototype.hasOwnProperty.call(speedConfig, "fallbackServiceSpeedMph")
  );
  const config = hasConfigShape
    ? speedConfig
    : { classSpeedMap: speedConfig };
  const speedClassField = normalizeRouteSequencerSpeedClassField(
    config.speedClassField || ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD
  );
  const drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
    config.drivingSpeedCatMap || config.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
    ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
  );
  const serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
    config.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
    ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
  );
  const fallbackDrivingMph = clampRouteSequencerNumber(
    fallbackSpeedMph,
    5,
    85,
    ROUTE_SEQUENCER_DEFAULT_SPEED_MPH
  );
  const fallbackServiceMph = clampRouteSequencerNumber(
    config.fallbackServiceSpeedMph,
    5,
    85,
    Math.min(fallbackDrivingMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
  );

  const posted = parsePostedSpeedFromTag(row?.maxspeed);
  let drivingMph = Number.isFinite(posted) && posted > 0
    ? clampRouteSequencerNumber(posted, 5, 85, fallbackDrivingMph)
    : NaN;

  const speedClassId = resolveRouteSequencerSpeedClassIdFromRow(row, speedClassField);
  if (!Number.isFinite(drivingMph) && Number.isFinite(speedClassId)) {
    drivingMph = clampRouteSequencerNumber(drivingSpeedCatMap?.[speedClassId], 5, 85, fallbackDrivingMph);
  }
  if (!Number.isFinite(drivingMph)) {
    drivingMph = fallbackDrivingMph;
  }

  let serviceMph = Number.isFinite(speedClassId)
    ? clampRouteSequencerNumber(serviceSpeedCatMap?.[speedClassId], 5, 85, fallbackServiceMph)
    : NaN;
  if (!Number.isFinite(serviceMph)) {
    serviceMph = fallbackServiceMph;
  }
  serviceMph = Math.min(drivingMph, Math.max(5, serviceMph));

  return {
    speedClassField,
    speedClassId,
    drivingMph,
    serviceMph
  };
}

class RouteSequencerMinHeap {
  constructor() {
    this.items = [];
  }

  get size() {
    return this.items.length;
  }

  push(item) {
    this.items.push(item);
    this._bubbleUp(this.items.length - 1);
  }

  pop() {
    if (!this.items.length) return null;
    const first = this.items[0];
    const end = this.items.pop();
    if (this.items.length && end) {
      this.items[0] = end;
      this._sinkDown(0);
    }
    return first;
  }

  _bubbleUp(index) {
    let idx = index;
    const element = this.items[idx];
    while (idx > 0) {
      const parentIdx = Math.floor((idx - 1) / 2);
      const parent = this.items[parentIdx];
      if ((parent?.cost ?? Infinity) <= (element?.cost ?? Infinity)) break;
      this.items[parentIdx] = element;
      this.items[idx] = parent;
      idx = parentIdx;
    }
  }

  _sinkDown(index) {
    let idx = index;
    const length = this.items.length;
    const element = this.items[idx];

    while (true) {
      const leftIdx = (2 * idx) + 1;
      const rightIdx = leftIdx + 1;
      let swapIdx = null;

      if (leftIdx < length) {
        const left = this.items[leftIdx];
        if ((left?.cost ?? Infinity) < (element?.cost ?? Infinity)) swapIdx = leftIdx;
      }

      if (rightIdx < length) {
        const right = this.items[rightIdx];
        const current = swapIdx === null ? element : this.items[swapIdx];
        if ((right?.cost ?? Infinity) < (current?.cost ?? Infinity)) swapIdx = rightIdx;
      }

      if (swapIdx === null) break;
      this.items[idx] = this.items[swapIdx];
      this.items[swapIdx] = element;
      idx = swapIdx;
    }
  }
}

function routeSequencerGraphNodeKey(lat, lon) {
  return `${Number(lat).toFixed(6)},${Number(lon).toFixed(6)}`;
}

function routeSequencerCellKey(lat, lon, cellSizeDeg = ROUTE_SEQUENCER_NODE_CELL_DEG) {
  const latIdx = Math.floor(Number(lat) / cellSizeDeg);
  const lonIdx = Math.floor(Number(lon) / cellSizeDeg);
  return `${latIdx}:${lonIdx}`;
}

function buildRouteSequencerStreetGraph(speedConfig, fallbackSpeedMph) {
  const hasConfigShape = speedConfig && typeof speedConfig === "object" && (
    Object.prototype.hasOwnProperty.call(speedConfig, "speedClassField")
    || Object.prototype.hasOwnProperty.call(speedConfig, "drivingSpeedCatMap")
    || Object.prototype.hasOwnProperty.call(speedConfig, "serviceSpeedCatMap")
    || Object.prototype.hasOwnProperty.call(speedConfig, "classSpeedMap")
    || Object.prototype.hasOwnProperty.call(speedConfig, "fallbackServiceSpeedMph")
  );
  const config = hasConfigShape
    ? speedConfig
    : { classSpeedMap: speedConfig };
  const speedClassField = normalizeRouteSequencerSpeedClassField(
    config.speedClassField || ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD
  );
  const drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
    config.drivingSpeedCatMap || config.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
    ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
  );
  const serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
    config.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
    ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
  );
  const fallbackDrivingMph = clampRouteSequencerNumber(
    fallbackSpeedMph,
    5,
    85,
    ROUTE_SEQUENCER_DEFAULT_SPEED_MPH
  );
  const fallbackServiceSpeedMph = clampRouteSequencerNumber(
    config.fallbackServiceSpeedMph,
    5,
    85,
    Math.min(fallbackDrivingMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
  );
  const nodes = new Map();
  let edgeCount = 0;

  const getOrCreateNode = (lat, lon) => {
    const key = routeSequencerGraphNodeKey(lat, lon);
    let node = nodes.get(key);
    if (!node) {
      node = { key, lat: Number(lat), lon: Number(lon), edges: [] };
      nodes.set(key, node);
    }
    return node;
  };

  const addDirectedEdge = (fromNode, toNode, drivingMinutes, serviceMinutes, miles) => {
    if (!fromNode || !toNode) return;
    if (!Number.isFinite(drivingMinutes) || drivingMinutes <= 0) return;
    if (!Number.isFinite(serviceMinutes) || serviceMinutes <= 0) return;
    fromNode.edges.push({
      to: toNode.key,
      minutes: drivingMinutes,
      serviceMinutes,
      miles: Number.isFinite(miles) ? miles : 0
    });
    edgeCount += 1;
  };

  streetAttributeById.forEach(entry => {
    const row = entry?.row || {};
    const latLon = flattenStreetLayerLatLngs(entry?.layer?.getLatLngs?.(), []);
    if (!Array.isArray(latLon) || latLon.length < 2) return;

    const speeds = resolveStreetSpeedsForSequencer(
      row,
      {
        speedClassField,
        drivingSpeedCatMap,
        serviceSpeedCatMap,
        fallbackServiceSpeedMph
      },
      fallbackDrivingMph
    );
    const onewayMode = normalizeRouteSequencerOneway(row?.oneway);

    for (let i = 1; i < latLon.length; i += 1) {
      const prev = latLon[i - 1];
      const next = latLon[i];
      const distMiles = haversineMiles(prev.lat, prev.lon, next.lat, next.lon);
      if (!Number.isFinite(distMiles) || distMiles <= 0) continue;
      const drivingMinutes = (distMiles / speeds.drivingMph) * 60;
      const serviceMinutes = (distMiles / speeds.serviceMph) * 60;
      if (!Number.isFinite(drivingMinutes) || drivingMinutes <= 0) continue;
      if (!Number.isFinite(serviceMinutes) || serviceMinutes <= 0) continue;

      const fromNode = getOrCreateNode(prev.lat, prev.lon);
      const toNode = getOrCreateNode(next.lat, next.lon);

      if (onewayMode !== "reverse") {
        addDirectedEdge(fromNode, toNode, drivingMinutes, serviceMinutes, distMiles);
      }
      if (onewayMode !== "forward") {
        addDirectedEdge(toNode, fromNode, drivingMinutes, serviceMinutes, distMiles);
      }
    }
  });

  const cellIndex = new Map();
  nodes.forEach(node => {
    const key = routeSequencerCellKey(node.lat, node.lon, ROUTE_SEQUENCER_NODE_CELL_DEG);
    const bucket = cellIndex.get(key) || [];
    bucket.push(node.key);
    cellIndex.set(key, bucket);
  });

  return {
    nodes,
    cellIndex,
    nodeCount: nodes.size,
    edgeCount,
    speedClassField,
    drivingSpeedCatMap,
    serviceSpeedCatMap,
    fallbackSpeedMph: fallbackDrivingMph,
    fallbackServiceSpeedMph
  };
}

function findNearestRouteSequencerGraphNode(graph, lat, lon) {
  if (!graph || !(graph.nodes instanceof Map) || !graph.nodes.size) return null;
  const latNum = Number(lat);
  const lonNum = Number(lon);
  if (!Number.isFinite(latNum) || !Number.isFinite(lonNum)) return null;

  const baseLatIdx = Math.floor(latNum / ROUTE_SEQUENCER_NODE_CELL_DEG);
  const baseLonIdx = Math.floor(lonNum / ROUTE_SEQUENCER_NODE_CELL_DEG);
  let best = null;

  const evaluateNodeKey = nodeKey => {
    const node = graph.nodes.get(nodeKey);
    if (!node) return;
    const miles = haversineMiles(latNum, lonNum, node.lat, node.lon);
    if (!Number.isFinite(miles)) return;
    if (!best || miles < best.distanceMiles) {
      best = {
        nodeKey: node.key,
        lat: node.lat,
        lon: node.lon,
        distanceMiles: miles
      };
    }
  };

  for (let radius = 0; radius <= ROUTE_SEQUENCER_MAX_NODE_SEARCH_RADIUS; radius += 1) {
    let foundCandidate = false;
    for (let dy = -radius; dy <= radius; dy += 1) {
      for (let dx = -radius; dx <= radius; dx += 1) {
        if (radius > 0 && Math.abs(dx) !== radius && Math.abs(dy) !== radius) continue;
        const key = `${baseLatIdx + dy}:${baseLonIdx + dx}`;
        const bucket = graph.cellIndex.get(key);
        if (!bucket || !bucket.length) continue;
        foundCandidate = true;
        bucket.forEach(evaluateNodeKey);
      }
    }
    if (foundCandidate && best) break;
  }

  if (!best) return null;
  if (best.distanceMiles > ROUTE_SEQUENCER_MAX_SNAP_MILES) return null;
  return best;
}

function estimateRouteSequencerTravelMinutes(fromPoint, toPoint, fallbackSpeedMph, options = {}) {
  if (!fromPoint || !toPoint) return 0;
  const distMiles = haversineMiles(fromPoint.lat, fromPoint.lon, toPoint.lat, toPoint.lon);
  if (!Number.isFinite(distMiles) || distMiles <= 0) return 0;
  const travelMode = normalizeRouteSequencerTravelMode(options?.travelMode);
  const fallbackDrivingMph = clampRouteSequencerNumber(
    fallbackSpeedMph,
    5,
    85,
    ROUTE_SEQUENCER_DEFAULT_SPEED_MPH
  );
  const fallbackServiceSpeedMph = clampRouteSequencerNumber(
    options?.fallbackServiceSpeedMph,
    5,
    85,
    Math.min(fallbackDrivingMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
  );
  const speed = travelMode === "service"
    ? fallbackServiceSpeedMph
    : fallbackDrivingMph;
  return (distMiles / speed) * 60;
}

function estimateRouteSequencerTravelMiles(fromPoint, toPoint) {
  if (!fromPoint || !toPoint) return 0;
  const distMiles = haversineMiles(fromPoint.lat, fromPoint.lon, toPoint.lat, toPoint.lon);
  if (!Number.isFinite(distMiles) || distMiles <= 0) return 0;
  return distMiles;
}

function dijkstraRouteSequencerDistances(graph, startNodeKey, targetNodeKeys, options = {}) {
  if (!graph || !graph.nodes?.size) return new Map();
  const start = String(startNodeKey || "");
  if (!start || !graph.nodes.has(start)) return new Map();
  const travelMode = normalizeRouteSequencerTravelMode(options?.travelMode);

  const targets = new Set(
    [...(targetNodeKeys || [])]
      .map(key => String(key || ""))
      .filter(key => key && graph.nodes.has(key))
  );
  if (!targets.size) return new Map();

  const result = new Map();
  if (targets.has(start)) {
    result.set(start, { minutes: 0, miles: 0 });
    targets.delete(start);
    if (!targets.size) return result;
  }

  const dist = new Map([[start, { minutes: 0, miles: 0 }]]);
  const heap = new RouteSequencerMinHeap();
  heap.push({ key: start, cost: 0 });

  while (heap.size && targets.size) {
    const current = heap.pop();
    if (!current) break;
    const currentKey = String(current.key || "");
    const currentCost = Number(current.cost);
    const knownCost = dist.get(currentKey);
    if (!Number.isFinite(currentCost) || !knownCost || !Number.isFinite(knownCost.minutes)) continue;
    if (currentCost > (knownCost.minutes + 1e-9)) continue;

    if (targets.has(currentKey)) {
      result.set(currentKey, { minutes: knownCost.minutes, miles: knownCost.miles });
      targets.delete(currentKey);
      if (!targets.size) break;
    }

    const node = graph.nodes.get(currentKey);
    if (!node || !Array.isArray(node.edges) || !node.edges.length) continue;
    node.edges.forEach(edge => {
      const toKey = String(edge?.to || "");
      if (!toKey || !graph.nodes.has(toKey)) return;
      let stepMinutes = Number(
        travelMode === "service"
          ? edge?.serviceMinutes
          : edge?.minutes
      );
      if (!Number.isFinite(stepMinutes) && travelMode === "service") {
        stepMinutes = Number(edge?.minutes);
      }
      if (!Number.isFinite(stepMinutes) || stepMinutes <= 0) return;
      const stepMiles = Math.max(0, Number(edge?.miles) || 0);
      const nextMinutes = knownCost.minutes + stepMinutes;
      const nextMiles = knownCost.miles + stepMiles;
      const prevBest = dist.get(toKey);
      const shouldUpdate = !prevBest
        || !Number.isFinite(prevBest.minutes)
        || nextMinutes < (prevBest.minutes - 1e-9)
        || (
          Math.abs(nextMinutes - prevBest.minutes) <= 1e-9
          && nextMiles < ((Number.isFinite(prevBest.miles) ? prevBest.miles : Infinity) - 1e-9)
        );
      if (!shouldUpdate) return;
      dist.set(toKey, { minutes: nextMinutes, miles: nextMiles });
      heap.push({ key: toKey, cost: nextMinutes });
    });
  }

  return result;
}

function pickNearestRouteSequencerCandidate(currentPoint, candidates, graph, fallbackSpeedMph, travelOptions = {}) {
  const candidateList = Array.isArray(candidates) ? candidates.filter(Boolean) : [];
  if (!candidateList.length) return null;
  const travelMode = normalizeRouteSequencerTravelMode(travelOptions?.travelMode);
  const requireNetworkOnly = travelOptions?.requireNetworkOnly !== false;
  const candidateAdjustmentFn = typeof travelOptions?.candidateAdjustmentFn === "function"
    ? travelOptions.candidateAdjustmentFn
    : null;
  const fallbackServiceSpeedMph = clampRouteSequencerNumber(
    travelOptions?.fallbackServiceSpeedMph,
    5,
    85,
    Math.min(
      clampRouteSequencerNumber(fallbackSpeedMph, 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH),
      ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH
    )
  );

  if (!currentPoint) {
    return {
      point: candidateList[0],
      minutes: 0,
      scoreMinutes: 0,
      miles: 0,
      usedNetwork: false,
      travelMode,
      adjustment: normalizeRouteSequencerCandidateAdjustment(null)
    };
  }

  if (requireNetworkOnly && (!graph || !currentPoint.nodeKey)) {
    return null;
  }

  let networkDistances = null;
  if (graph && currentPoint.nodeKey) {
    const targetKeys = new Set(
      candidateList
        .map(item => String(item?.nodeKey || ""))
        .filter(Boolean)
    );
    if (targetKeys.size) {
      networkDistances = dijkstraRouteSequencerDistances(
        graph,
        currentPoint.nodeKey,
        targetKeys,
        { travelMode }
      );
    }
  }

  let best = null;
  candidateList.forEach(candidate => {
    let minutes = NaN;
    let miles = NaN;
    let usedNetwork = false;
    if (networkDistances && candidate.nodeKey && networkDistances.has(candidate.nodeKey)) {
      const networkRoute = networkDistances.get(candidate.nodeKey) || {};
      minutes = Number(networkRoute.minutes);
      miles = Number(networkRoute.miles);
      usedNetwork = Number.isFinite(minutes) && Number.isFinite(miles);
    }
    if (requireNetworkOnly && !usedNetwork) return;
    if (!requireNetworkOnly && (!Number.isFinite(minutes) || !Number.isFinite(miles))) {
      minutes = estimateRouteSequencerTravelMinutes(
        currentPoint,
        candidate,
        fallbackSpeedMph,
        { travelMode, fallbackServiceSpeedMph }
      );
      miles = estimateRouteSequencerTravelMiles(currentPoint, candidate);
      usedNetwork = false;
    }
    if (!Number.isFinite(minutes) || !Number.isFinite(miles)) return;

    let rawAdjustment = null;
    if (candidateAdjustmentFn) {
      try {
        rawAdjustment = candidateAdjustmentFn({
          fromPoint: currentPoint,
          candidate,
          minutes,
          miles,
          usedNetwork,
          travelMode
        });
      } catch {
        rawAdjustment = null;
      }
    }
    const adjustment = normalizeRouteSequencerCandidateAdjustment(rawAdjustment);
    const scoreMinutes = Math.max(0, minutes + adjustment.scoreMinutesDelta);

    if (!best || scoreMinutes < best.scoreMinutes - 1e-9 || (Math.abs(scoreMinutes - best.scoreMinutes) <= 1e-9 && miles < best.miles - 1e-9)) {
      best = { point: candidate, minutes, scoreMinutes, miles, usedNetwork, travelMode, adjustment };
      return;
    }
    if (Math.abs(scoreMinutes - best.scoreMinutes) <= 1e-9 && Math.abs(miles - best.miles) <= 1e-9) {
      const candidateOrder = Number(candidate?.order ?? Number.MAX_SAFE_INTEGER);
      const bestOrder = Number(best?.point?.order ?? Number.MAX_SAFE_INTEGER);
      if (candidateOrder < bestOrder) {
        best = { point: candidate, minutes, scoreMinutes, miles, usedNetwork, travelMode, adjustment };
      }
    }
  });

  return best;
}

function routeSequencerSplitRouteDayKey(key) {
  const text = String(key || "");
  const idx = text.lastIndexOf("|");
  if (idx === -1) return { route: text || "Unassigned", day: "" };
  return {
    route: text.slice(0, idx) || "Unassigned",
    day: text.slice(idx + 1) || ""
  };
}

function buildRouteSequencerStopLabel(row) {
  if (typeof buildRecordAddressLabel === "function") {
    try {
      const label = String(buildRecordAddressLabel(row) || "").trim();
      if (label) return label;
    } catch {
      // Fall through to local fallback.
    }
  }

  const legacyAddress = [
    row?.["CSADR#"] || "",
    row?.["CSSDIR"] || "",
    row?.["CSSTRT"] || "",
    row?.["CSSFUX"] || ""
  ]
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
  if (legacyAddress) return legacyAddress;

  const fallbackFields = [
    "Address",
    "ADDRESS",
    "Street",
    "STREET",
    "Name",
    "NAME",
    "UNIQUEID",
    "BINNO"
  ];
  for (const key of fallbackFields) {
    const value = String(row?.[key] ?? "").trim();
    if (value) return value;
  }
  return "Unnamed stop";
}

function buildRouteSequencerStopPoints(routeDayKey, group, demandField, serviceTimeField, graph, streetAttributeDefaults = null) {
  const layers = Array.isArray(group?.layers) ? group.layers : [];
  const normalizedStreetDefaults = normalizeRouteSequencerStreetAttributeDefaults(streetAttributeDefaults);
  const streetClassField = normalizedStreetDefaults.streetClassField;
  const out = [];
  layers.forEach(layer => {
    const rowId = Number(layer?._rowId);
    if (!Number.isFinite(rowId)) return;
    const row = layer?._rowRef;
    if (!row || typeof row !== "object") return;
    const lat = Number(layer?._base?.lat);
    const lon = Number(layer?._base?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    const demand = Math.max(0, parseRouteSequencerNumeric(demandField ? row?.[demandField] : 0, 0));
    const serviceMinutes = Math.max(0, parseRouteSequencerMinutesValue(serviceTimeField ? row?.[serviceTimeField] : 0, 0));
    const streetMeta = resolveRouteSequencerStopStreetMetadata(row, streetClassField);
    const snap = graph ? findNearestRouteSequencerGraphNode(graph, lat, lon) : null;

    out.push({
      type: "stop",
      id: `stop_${rowId}`,
      rowId,
      routeDayKey,
      route: String(layer?._routeToken ?? ""),
      day: String(layer?._dayToken ?? ""),
      name: buildRouteSequencerStopLabel(row),
      lat,
      lon,
      nodeKey: snap?.nodeKey || "",
      snapMiles: snap?.distanceMiles || NaN,
      demand,
      serviceMinutes,
      segmentId: streetMeta.segmentId,
      segmentSide: streetMeta.segmentSide,
      streetClassId: streetMeta.streetClassId,
      order: rowId
    });
  });

  out.sort((a, b) => Number(a.order) - Number(b.order));
  return out;
}

function buildRouteSequencerStopPointsFromExistingSequence(routeDayKey, group, demandField, serviceTimeField, sequenceField, graph, streetAttributeDefaults = null) {
  const layers = Array.isArray(group?.layers) ? group.layers : [];
  const normalizedStreetDefaults = normalizeRouteSequencerStreetAttributeDefaults(streetAttributeDefaults);
  const streetClassField = normalizedStreetDefaults.streetClassField;
  const out = [];
  let validSequenceCount = 0;
  let missingSequenceCount = 0;

  layers.forEach(layer => {
    const rowId = Number(layer?._rowId);
    if (!Number.isFinite(rowId)) return;
    const row = layer?._rowRef;
    if (!row || typeof row !== "object") return;
    const lat = Number(layer?._base?.lat);
    const lon = Number(layer?._base?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    const demand = Math.max(0, parseRouteSequencerNumeric(demandField ? row?.[demandField] : 0, 0));
    const serviceMinutes = Math.max(0, parseRouteSequencerMinutesValue(serviceTimeField ? row?.[serviceTimeField] : 0, 0));
    const sequenceValue = parseRouteSequencerNumeric(sequenceField ? row?.[sequenceField] : "", NaN);
    const streetMeta = resolveRouteSequencerStopStreetMetadata(row, streetClassField);
    const hasSequence = Number.isFinite(sequenceValue) && sequenceValue > 0;
    if (hasSequence) validSequenceCount += 1;
    else missingSequenceCount += 1;

    const snap = graph ? findNearestRouteSequencerGraphNode(graph, lat, lon) : null;

    out.push({
      type: "stop",
      id: `stop_${rowId}`,
      rowId,
      routeDayKey,
      route: String(layer?._routeToken ?? ""),
      day: String(layer?._dayToken ?? ""),
      name: buildRouteSequencerStopLabel(row),
      lat,
      lon,
      nodeKey: snap?.nodeKey || "",
      snapMiles: snap?.distanceMiles || NaN,
      demand,
      serviceMinutes,
      segmentId: streetMeta.segmentId,
      segmentSide: streetMeta.segmentSide,
      streetClassId: streetMeta.streetClassId,
      sequenceOrder: hasSequence ? Number(sequenceValue) : NaN,
      hasSequence,
      order: rowId
    });
  });

  out.sort((a, b) => {
    const aSeq = Number(a?.sequenceOrder);
    const bSeq = Number(b?.sequenceOrder);
    const aHas = Number.isFinite(aSeq) && aSeq > 0;
    const bHas = Number.isFinite(bSeq) && bSeq > 0;
    if (aHas && bHas && aSeq !== bSeq) return aSeq - bSeq;
    if (aHas !== bHas) return aHas ? -1 : 1;
    return Number(a?.order) - Number(b?.order);
  });

  return { stops: out, validSequenceCount, missingSequenceCount };
}

function collectRouteSequencerLandfillDumpPoints(graph) {
  const out = [];
  texasLandfillsLayerGroup.eachLayer(layer => {
    const ll = getLayerLatLng(layer);
    if (!ll) return;
    const title = normalizeRouteSequencerName(layer?.options?.title || "Texas Landfill", "Texas Landfill");
    const snap = graph ? findNearestRouteSequencerGraphNode(graph, ll.lat, ll.lng) : null;
    out.push({
      type: "dump",
      id: `landfill_${ll.lat.toFixed(6)}_${ll.lng.toFixed(6)}`,
      source: "landfills",
      name: title,
      lat: ll.lat,
      lon: ll.lng,
      facilityMinutes: 0,
      nodeKey: snap?.nodeKey || "",
      snapMiles: snap?.distanceMiles || NaN
    });
  });
  return out;
}

function collectRouteSequencerCustomDumpPoints(graph) {
  return getRouteSequencerCustomDumps().map(item => {
    const snap = graph ? findNearestRouteSequencerGraphNode(graph, item.lat, item.lon) : null;
    return {
      type: "dump",
      id: String(item.id),
      source: "custom",
      name: normalizeRouteSequencerName(item.name, "Custom Dump"),
      lat: Number(item.lat),
      lon: Number(item.lon),
      facilityMinutes: parseRouteSequencerMinutesValue(item?.facilityMinutes, 0),
      nodeKey: snap?.nodeKey || "",
      snapMiles: snap?.distanceMiles || NaN
    };
  });
}

function getRouteSequencerDepotPointById(depotId, graph) {
  const id = normalizeRouteSequencerText(depotId);
  if (!id) return null;
  const depot = getRouteSequencerDepots().find(item => String(item.id) === id);
  if (!depot) return null;
  const snap = graph ? findNearestRouteSequencerGraphNode(graph, depot.lat, depot.lon) : null;
  return {
    type: "depot",
    id: String(depot.id),
    name: normalizeRouteSequencerName(depot.name, "Depot"),
    lat: Number(depot.lat),
    lon: Number(depot.lon),
    facilityMinutes: 0,
    allowIntermediate: false,
    nodeKey: snap?.nodeKey || "",
    snapMiles: snap?.distanceMiles || NaN
  };
}

function solveRouteDaySequencingPlan(routeDayKey, stops, options = {}) {
  const routeStops = Array.isArray(stops) ? stops.slice() : [];
  if (!routeStops.length) return null;

  routeStops.sort((a, b) => Number(a.order) - Number(b.order));
  const unserved = new Map(routeStops.map(item => [item.id, item]));
  const capacityTons = Number(options.capacityTons);
  const fallbackSpeedMph = clampRouteSequencerNumber(options.fallbackSpeedMph, 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH);
  const fallbackServiceSpeedMph = clampRouteSequencerNumber(
    options.fallbackServiceSpeedMph,
    5,
    85,
    Math.min(fallbackSpeedMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
  );
  const dumpPoints = Array.isArray(options.dumpPoints) ? options.dumpPoints.filter(Boolean) : [];
  const breakMinutesList = parseRouteSequencerBreakMinutesList(options.breakMinutesList);
  const graph = options.graph || null;
  const requireNetworkOnly = options.requireNetworkOnly !== false;
  const startDepot = options.startDepot || null;
  const endDepot = options.endDepot || null;
  const forceDumpBeforeEnd = !!options.forceDumpBeforeEnd;
  const serviceStopsBeforeDeadhead = options.serviceStopsBeforeDeadhead == null
    ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
    : options.serviceStopsBeforeDeadhead !== false;
  const timeAtStartMinutes = Math.max(0, parseRouteSequencerMinutesValue(options.timeAtStartMinutes, 0));
  const timeAtEndMinutes = Math.max(0, parseRouteSequencerMinutesValue(options.timeAtEndMinutes, 0));
  const streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(options.streetAttributeDefaults);
  const meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(options.meanderOverrideDefaults);
  if (requireNetworkOnly && (!graph || !(graph.nodes instanceof Map) || !graph.nodes.size)) {
    return null;
  }

  let fixedFirstStop = null;
  let fixedLastStop = null;
  if (!startDepot && routeStops.length) {
    fixedFirstStop = routeStops[0];
    unserved.delete(fixedFirstStop.id);
  }
  if (!endDepot && routeStops.length > 1) {
    const candidate = routeStops[routeStops.length - 1];
    if (!fixedFirstStop || candidate.id !== fixedFirstStop.id) {
      fixedLastStop = candidate;
      unserved.delete(candidate.id);
    }
  }

  const stopVisits = [];
  const events = [];
  let currentPoint = null;
  let currentLoad = 0;
  let tripNumber = 1;
  let totalDriveMinutes = 0;
  let totalDriveMiles = 0;
  let totalServiceMinutes = 0;
  let totalFacilityMinutes = 0;
  let totalDumpFacilityMinutes = 0;
  let totalStartEndFacilityMinutes = 0;
  let totalBreakMinutes = 0;
  let totalOtherMinutes = 0;
  let dumpVisits = 0;
  let networkLegs = 0;
  let fallbackLegs = 0;

  const applyDriveLeg = pickResult => {
    const adjustment = normalizeRouteSequencerCandidateAdjustment(pickResult?.adjustment);
    const baseDriveMinutes = Math.max(0, Number(pickResult?.minutes) || 0);
    const drivePenaltyMinutes = Math.max(0, Number(adjustment.driveMinutesDelta) || 0);
    const driveMinutes = baseDriveMinutes + drivePenaltyMinutes;
    const driveMiles = Math.max(0, Number(pickResult?.miles) || 0);
    totalDriveMinutes += driveMinutes;
    totalDriveMiles += driveMiles;
    const hasTravelLeg = driveMinutes > 1e-9 || driveMiles > 1e-9;
    if (hasTravelLeg) {
      if (pickResult?.usedNetwork) networkLegs += 1;
      else fallbackLegs += 1;
    }
    return { driveMinutes, driveMiles, baseDriveMinutes, drivePenaltyMinutes, adjustment };
  };

  const appendStopVisit = (stopPoint, pickResult) => {
    const { driveMinutes, driveMiles, adjustment } = applyDriveLeg(pickResult);
    const demand = Math.max(0, Number(stopPoint?.demand) || 0);
    const baseServiceMinutes = Math.max(0, Number(stopPoint?.serviceMinutes) || 0);
    const servicePenaltyMinutes = Math.max(0, Number(adjustment?.serviceMinutesDelta) || 0);
    const otherPenaltyMinutes = Math.max(0, Number(adjustment?.otherMinutesDelta) || 0);
    const serviceMinutes = baseServiceMinutes + servicePenaltyMinutes;
    if (otherPenaltyMinutes > 0) totalOtherMinutes += otherPenaltyMinutes;
    currentLoad += demand;
    totalServiceMinutes += serviceMinutes;
    const sequence = stopVisits.length + 1;
    const penaltyNotes = Array.isArray(adjustment?.notes)
      ? adjustment.notes.map(note => String(note || "").trim()).filter(Boolean).join("; ")
      : "";
    const visit = {
      sequence,
      rowId: Number(stopPoint.rowId),
      routeDayKey,
      trip: tripNumber,
      demand,
      cumulativeLoad: currentLoad,
      driveMinutes,
      driveMiles,
      serviceMinutes,
      otherMinutes: otherPenaltyMinutes,
      penaltyNotes,
      totalMinutes: driveMinutes + serviceMinutes + otherPenaltyMinutes
    };
    stopVisits.push(visit);
    events.push({
      type: "stop",
      rowId: visit.rowId,
      sequence: visit.sequence,
      trip: visit.trip,
      name: stopPoint.name,
      demand,
      driveMinutes,
      driveMiles,
      serviceMinutes,
      otherMinutes: otherPenaltyMinutes,
      penaltyNotes,
      cumulativeLoad: visit.cumulativeLoad,
      lat: Number(stopPoint?.lat),
      lon: Number(stopPoint?.lon)
    });
    currentPoint = stopPoint;
    return visit;
  };

  const appendDumpVisit = (pickResult, reason = "capacity") => {
    const point = pickResult?.point;
    if (!point) return false;
    const { driveMinutes, driveMiles } = applyDriveLeg(pickResult);
    const facilityMinutes = Math.max(0, parseRouteSequencerMinutesValue(point?.facilityMinutes, 0));
    totalDumpFacilityMinutes += facilityMinutes;
    totalFacilityMinutes += facilityMinutes;
    events.push({
      type: "dump",
      name: point.name,
      driveMinutes,
      driveMiles,
      beforeLoad: currentLoad,
      facilityMinutes,
      reason: String(reason || "capacity"),
      trip: tripNumber,
      lat: Number(point?.lat),
      lon: Number(point?.lon)
    });
    currentPoint = point;
    currentLoad = 0;
    tripNumber += 1;
    dumpVisits += 1;
    return true;
  };

  const appendBreakEvent = (minutesValue, breakNumber = 1) => {
    const breakMinutes = Math.max(0, parseRouteSequencerMinutesValue(minutesValue, 0));
    if (!(breakMinutes > 0)) return false;
    totalBreakMinutes += breakMinutes;
    totalOtherMinutes += breakMinutes;
    events.push({
      type: "break",
      name: `Driver Break ${Math.max(1, Number(breakNumber) || 1)}`,
      driveMinutes: 0,
      serviceMinutes: 0,
      facilityMinutes: 0,
      minutes: breakMinutes,
      otherMinutes: breakMinutes,
      trip: tripNumber
    });
    return true;
  };

  const pickServiceStopCandidate = candidateList => pickNearestRouteSequencerCandidate(
    currentPoint,
    candidateList,
    graph,
    fallbackSpeedMph,
    {
      travelMode: "service",
      fallbackServiceSpeedMph,
      requireNetworkOnly,
      candidateAdjustmentFn: ({ fromPoint, candidate, minutes }) =>
        buildRouteSequencerServiceCandidateAdjustment(
          fromPoint,
          candidate,
          minutes,
          streetAttributeDefaults,
          meanderOverrideDefaults
        )
    }
  );

  if (timeAtStartMinutes > 0) {
    totalStartEndFacilityMinutes += timeAtStartMinutes;
    totalFacilityMinutes += timeAtStartMinutes;
    events.push({ type: "timeAtStart", facilityMinutes: timeAtStartMinutes, trip: tripNumber });
  }

  if (breakMinutesList.length) {
    breakMinutesList.forEach((breakMinutes, index) => {
      appendBreakEvent(breakMinutes, index + 1);
    });
  }

  if (startDepot) {
    currentPoint = startDepot;
    events.push({
      type: "startDepot",
      name: startDepot.name,
      trip: tripNumber,
      lat: Number(startDepot?.lat),
      lon: Number(startDepot?.lon)
    });
  } else if (fixedFirstStop) {
    appendStopVisit(fixedFirstStop, { minutes: 0, miles: 0, usedNetwork: false });
  } else {
    const first = routeStops[0];
    unserved.delete(first.id);
    appendStopVisit(first, { minutes: 0, miles: 0, usedNetwork: false });
  }

  while (unserved.size) {
    const candidates = [...unserved.values()];
    if (!candidates.length) break;

    let nextStopPick = pickServiceStopCandidate(candidates);
    if (!nextStopPick?.point) {
      if (requireNetworkOnly) return null;
      break;
    }
    let nextStop = nextStopPick.point;
    let nextDemand = Math.max(0, Number(nextStop?.demand) || 0);
    const enforceCapacity = Number.isFinite(capacityTons) && capacityTons > 0;
    let wouldExceedCapacity = enforceCapacity
      && capacityTons > 0
      && currentLoad > 0
      && (currentLoad + nextDemand) > capacityTons;

    if (serviceStopsBeforeDeadhead && enforceCapacity && wouldExceedCapacity) {
      const remainingCapacity = Math.max(0, capacityTons - currentLoad);
      const fitCandidates = candidates.filter(candidate => {
        const demand = Math.max(0, Number(candidate?.demand) || 0);
        return demand <= (remainingCapacity + 1e-9);
      });
      if (fitCandidates.length) {
        const fitPick = pickServiceStopCandidate(fitCandidates);
        if (fitPick?.point) {
          nextStopPick = fitPick;
          nextStop = fitPick.point;
          nextDemand = Math.max(0, Number(nextStop?.demand) || 0);
          wouldExceedCapacity = enforceCapacity
            && currentLoad > 0
            && (currentLoad + nextDemand) > capacityTons;
        }
      }
    }

    if (wouldExceedCapacity && dumpPoints.length) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point && appendDumpVisit(dumpPick, "capacity-before-next-stop")) {
        continue;
      }
      if (requireNetworkOnly) return null;
    }

    unserved.delete(nextStop.id);
    appendStopVisit(nextStop, nextStopPick);

    const atOrAboveCapacity = Number.isFinite(capacityTons)
      && capacityTons > 0
      && currentLoad >= capacityTons;
    if (atOrAboveCapacity && dumpPoints.length && unserved.size) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point) {
        appendDumpVisit(dumpPick, "capacity-reached");
      } else if (requireNetworkOnly) {
        return null;
      }
    }
  }

  if (fixedLastStop) {
    const nextDemand = Math.max(0, Number(fixedLastStop?.demand) || 0);
    const wouldExceedCapacity = Number.isFinite(capacityTons)
      && capacityTons > 0
      && currentLoad > 0
      && (currentLoad + nextDemand) > capacityTons;

    if (!serviceStopsBeforeDeadhead && wouldExceedCapacity && dumpPoints.length) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point) {
        appendDumpVisit(dumpPick, "capacity-before-last-stop");
      } else if (requireNetworkOnly) {
        return null;
      }
    }

    const pick = pickNearestRouteSequencerCandidate(
      currentPoint,
      [fixedLastStop],
      graph,
      fallbackSpeedMph,
      {
        travelMode: "service",
        fallbackServiceSpeedMph,
        requireNetworkOnly,
        candidateAdjustmentFn: ({ fromPoint, candidate, minutes }) =>
          buildRouteSequencerServiceCandidateAdjustment(
            fromPoint,
            candidate,
            minutes,
            streetAttributeDefaults,
            meanderOverrideDefaults
          )
      }
    );
    if (!pick?.point) {
      return null;
    }
    appendStopVisit(fixedLastStop, pick);
  }

  if (endDepot) {
    const needsFinalDump = forceDumpBeforeEnd
      && Number.isFinite(capacityTons)
      && capacityTons > 0
      && currentLoad > 0
      && dumpPoints.length;
    if (needsFinalDump) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point) {
        appendDumpVisit(dumpPick, "final-before-end");
      } else if (requireNetworkOnly) {
        return null;
      }
    }

    const depotPick = pickNearestRouteSequencerCandidate(
      currentPoint,
      [endDepot],
      graph,
      fallbackSpeedMph,
      { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
    );
    if (!depotPick?.point) {
      return null;
    }
    const { driveMinutes, driveMiles } = applyDriveLeg(depotPick);
    events.push({
      type: "endDepot",
      name: endDepot.name,
      driveMinutes,
      driveMiles,
      trip: tripNumber,
      lat: Number(endDepot?.lat),
      lon: Number(endDepot?.lon)
    });
  }

  if (timeAtEndMinutes > 0) {
    totalStartEndFacilityMinutes += timeAtEndMinutes;
    totalFacilityMinutes += timeAtEndMinutes;
    events.push({ type: "timeAtEnd", facilityMinutes: timeAtEndMinutes, trip: tripNumber });
  }

  const split = routeSequencerSplitRouteDayKey(routeDayKey);
  return {
    routeDayKey,
    route: split.route,
    day: split.day,
    stopVisits,
    events,
    dumpVisits,
    totalDriveMinutes,
    totalDriveMiles,
    totalServiceMinutes,
    totalDumpFacilityMinutes,
    totalStartEndFacilityMinutes,
    totalBreakMinutes,
    totalOtherMinutes,
    totalFacilityMinutes,
    totalMinutes: totalDriveMinutes + totalServiceMinutes + totalFacilityMinutes + totalOtherMinutes,
    capacityTons: Number.isFinite(capacityTons) && capacityTons > 0 ? capacityTons : null,
    networkLegs,
    fallbackLegs
  };
}

function solveRouteDaySequenceEvaluationPlan(routeDayKey, stops, options = {}) {
  const orderedStops = Array.isArray(stops) ? stops.slice() : [];
  if (!orderedStops.length) return null;

  orderedStops.sort((a, b) => {
    const aSeq = Number(a?.sequenceOrder);
    const bSeq = Number(b?.sequenceOrder);
    const aHas = Number.isFinite(aSeq) && aSeq > 0;
    const bHas = Number.isFinite(bSeq) && bSeq > 0;
    if (aHas && bHas && aSeq !== bSeq) return aSeq - bSeq;
    if (aHas !== bHas) return aHas ? -1 : 1;
    return Number(a?.order) - Number(b?.order);
  });

  const graph = options.graph || null;
  const requireNetworkOnly = options.requireNetworkOnly !== false;
  const capacityTons = Number(options.capacityTons);
  const dumpPoints = Array.isArray(options.dumpPoints) ? options.dumpPoints.filter(Boolean) : [];
  const breakMinutesList = parseRouteSequencerBreakMinutesList(options.breakMinutesList);
  const fallbackSpeedMph = clampRouteSequencerNumber(options.fallbackSpeedMph, 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH);
  const fallbackServiceSpeedMph = clampRouteSequencerNumber(
    options.fallbackServiceSpeedMph,
    5,
    85,
    Math.min(fallbackSpeedMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
  );
  const startDepot = options.startDepot || null;
  const endDepot = options.endDepot || null;
  const forceDumpBeforeEnd = !!options.forceDumpBeforeEnd;
  const serviceStopsBeforeDeadhead = options.serviceStopsBeforeDeadhead == null
    ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
    : options.serviceStopsBeforeDeadhead !== false;
  const timeAtStartMinutes = Math.max(0, parseRouteSequencerMinutesValue(options.timeAtStartMinutes, 0));
  const timeAtEndMinutes = Math.max(0, parseRouteSequencerMinutesValue(options.timeAtEndMinutes, 0));
  const streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(options.streetAttributeDefaults);
  const meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(options.meanderOverrideDefaults);
  if (requireNetworkOnly && (!graph || !(graph.nodes instanceof Map) || !graph.nodes.size)) {
    return null;
  }

  const stopVisits = [];
  const events = [];
  let currentPoint = null;
  let currentLoad = 0;
  let tripNumber = 1;
  let dumpVisits = 0;
  let totalDriveMinutes = 0;
  let totalDriveMiles = 0;
  let totalServiceMinutes = 0;
  let totalFacilityMinutes = 0;
  let totalDumpFacilityMinutes = 0;
  let totalStartEndFacilityMinutes = 0;
  let totalBreakMinutes = 0;
  let totalOtherMinutes = 0;
  let networkLegs = 0;
  let fallbackLegs = 0;
  let failedNetworkOnly = false;
  const maxExistingSequence = orderedStops.reduce((maxValue, stopPoint) => {
    const seq = Number(stopPoint?.sequenceOrder);
    if (Number.isFinite(seq) && seq > maxValue) return seq;
    return maxValue;
  }, 0);
  let fallbackSequenceCursor = maxExistingSequence;

  const applyDriveLeg = pickResult => {
    const adjustment = normalizeRouteSequencerCandidateAdjustment(pickResult?.adjustment);
    const baseDriveMinutes = Math.max(0, Number(pickResult?.minutes) || 0);
    const drivePenaltyMinutes = Math.max(0, Number(adjustment?.driveMinutesDelta) || 0);
    const driveMinutes = baseDriveMinutes + drivePenaltyMinutes;
    const driveMiles = Math.max(0, Number(pickResult?.miles) || 0);
    totalDriveMinutes += driveMinutes;
    totalDriveMiles += driveMiles;
    const hasTravelLeg = driveMinutes > 1e-9 || driveMiles > 1e-9;
    if (hasTravelLeg) {
      if (pickResult?.usedNetwork) networkLegs += 1;
      else fallbackLegs += 1;
    }
    return { driveMinutes, driveMiles, baseDriveMinutes, drivePenaltyMinutes, adjustment };
  };

  const appendDumpVisit = (pickResult, reason = "capacity") => {
    const point = pickResult?.point;
    if (!point) return false;
    const { driveMinutes, driveMiles } = applyDriveLeg(pickResult);
    const facilityMinutes = Math.max(0, parseRouteSequencerMinutesValue(point?.facilityMinutes, 0));
    totalDumpFacilityMinutes += facilityMinutes;
    totalFacilityMinutes += facilityMinutes;
    events.push({
      type: "dump",
      name: point.name,
      driveMinutes,
      driveMiles,
      beforeLoad: currentLoad,
      facilityMinutes,
      reason: String(reason || "capacity"),
      trip: tripNumber,
      lat: Number(point?.lat),
      lon: Number(point?.lon)
    });
    currentPoint = point;
    currentLoad = 0;
    tripNumber += 1;
    dumpVisits += 1;
    return true;
  };

  const appendBreakEvent = (minutesValue, breakNumber = 1) => {
    const breakMinutes = Math.max(0, parseRouteSequencerMinutesValue(minutesValue, 0));
    if (!(breakMinutes > 0)) return false;
    totalBreakMinutes += breakMinutes;
    totalOtherMinutes += breakMinutes;
    events.push({
      type: "break",
      name: `Driver Break ${Math.max(1, Number(breakNumber) || 1)}`,
      driveMinutes: 0,
      serviceMinutes: 0,
      facilityMinutes: 0,
      minutes: breakMinutes,
      otherMinutes: breakMinutes,
      trip: tripNumber
    });
    return true;
  };

  const buildServiceCandidateAdjustment = (fromPoint, candidate, minutes) =>
    buildRouteSequencerServiceCandidateAdjustment(
      fromPoint,
      candidate,
      minutes,
      streetAttributeDefaults,
      meanderOverrideDefaults
    );

  if (timeAtStartMinutes > 0) {
    totalStartEndFacilityMinutes += timeAtStartMinutes;
    totalFacilityMinutes += timeAtStartMinutes;
    events.push({ type: "timeAtStart", facilityMinutes: timeAtStartMinutes, trip: tripNumber });
  }

  if (breakMinutesList.length) {
    breakMinutesList.forEach((breakMinutes, index) => {
      appendBreakEvent(breakMinutes, index + 1);
    });
  }

  if (startDepot) {
    currentPoint = startDepot;
    events.push({
      type: "startDepot",
      name: startDepot.name,
      trip: tripNumber,
      lat: Number(startDepot?.lat),
      lon: Number(startDepot?.lon)
    });
  }

  orderedStops.forEach((stopPoint, index) => {
    if (failedNetworkOnly) return;
    const nextDemand = Math.max(0, Number(stopPoint?.demand) || 0);
    const wouldExceedCapacity = Number.isFinite(capacityTons)
      && capacityTons > 0
      && currentLoad > 0
      && (currentLoad + nextDemand) > capacityTons;

    if (!serviceStopsBeforeDeadhead && wouldExceedCapacity && dumpPoints.length && currentPoint) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point) {
        appendDumpVisit(dumpPick, "capacity-before-next-stop");
      } else if (requireNetworkOnly) {
        failedNetworkOnly = true;
        return;
      }
    }

    let pick = currentPoint
      ? pickNearestRouteSequencerCandidate(
        currentPoint,
        [stopPoint],
        graph,
        fallbackSpeedMph,
        {
          travelMode: "service",
          fallbackServiceSpeedMph,
          requireNetworkOnly,
          candidateAdjustmentFn: ({ fromPoint, candidate, minutes }) =>
            buildServiceCandidateAdjustment(fromPoint, candidate, minutes)
        }
      )
      : { point: stopPoint, minutes: 0, miles: 0, usedNetwork: false };

    if (!pick?.point && currentPoint && !requireNetworkOnly) {
      const fallbackMinutes = estimateRouteSequencerTravelMinutes(currentPoint, stopPoint, fallbackSpeedMph, {
        travelMode: "service",
        fallbackServiceSpeedMph
      });
      pick = {
        point: stopPoint,
        minutes: fallbackMinutes,
        miles: estimateRouteSequencerTravelMiles(currentPoint, stopPoint),
        usedNetwork: false,
        adjustment: normalizeRouteSequencerCandidateAdjustment(
          buildServiceCandidateAdjustment(currentPoint, stopPoint, fallbackMinutes)
        )
      };
    }
    if (!pick?.point) {
      failedNetworkOnly = true;
      return;
    }

    const { driveMinutes, driveMiles, adjustment } = applyDriveLeg(pick);
    const demand = Math.max(0, Number(stopPoint?.demand) || 0);
    const baseServiceMinutes = Math.max(0, Number(stopPoint?.serviceMinutes) || 0);
    const servicePenaltyMinutes = Math.max(0, Number(adjustment?.serviceMinutesDelta) || 0);
    const otherPenaltyMinutes = Math.max(0, Number(adjustment?.otherMinutesDelta) || 0);
    const serviceMinutes = baseServiceMinutes + servicePenaltyMinutes;
    if (otherPenaltyMinutes > 0) totalOtherMinutes += otherPenaltyMinutes;
    currentLoad += demand;
    totalServiceMinutes += serviceMinutes;

    const rawSequence = Number(stopPoint?.sequenceOrder);
    const outputSequenceValue = Number.isFinite(rawSequence) && rawSequence > 0
      ? rawSequence
      : (++fallbackSequenceCursor);

    const visit = {
      sequence: outputSequenceValue,
      outputSequenceValue,
      rowId: Number(stopPoint.rowId),
      routeDayKey,
      trip: tripNumber,
      demand,
      cumulativeLoad: currentLoad,
      driveMinutes,
      driveMiles,
      serviceMinutes,
      otherMinutes: otherPenaltyMinutes,
      penaltyNotes: Array.isArray(adjustment?.notes)
        ? adjustment.notes.map(note => String(note || "").trim()).filter(Boolean).join("; ")
        : "",
      totalMinutes: driveMinutes + serviceMinutes + otherPenaltyMinutes
    };
    stopVisits.push(visit);
    events.push({
      type: "stop",
      rowId: visit.rowId,
      sequence: visit.sequence,
      trip: visit.trip,
      name: stopPoint.name,
      demand,
      driveMinutes,
      driveMiles,
      serviceMinutes,
      otherMinutes: otherPenaltyMinutes,
      penaltyNotes: visit.penaltyNotes,
      cumulativeLoad: visit.cumulativeLoad,
      lat: Number(stopPoint?.lat),
      lon: Number(stopPoint?.lon)
    });
    currentPoint = stopPoint;

    const hasRemainingStops = index < orderedStops.length - 1;
    const atOrAboveCapacity = Number.isFinite(capacityTons)
      && capacityTons > 0
      && currentLoad >= capacityTons;
    if (atOrAboveCapacity && dumpPoints.length && hasRemainingStops) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point) {
        appendDumpVisit(dumpPick, "capacity-reached");
      } else if (requireNetworkOnly) {
        failedNetworkOnly = true;
      }
    }
  });

  if (failedNetworkOnly) return null;

  if (endDepot) {
    const needsFinalDump = forceDumpBeforeEnd
      && Number.isFinite(capacityTons)
      && capacityTons > 0
      && currentLoad > 0
      && dumpPoints.length;
    if (needsFinalDump && currentPoint) {
      const dumpPick = pickNearestRouteSequencerCandidate(
        currentPoint,
        dumpPoints,
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      );
      if (dumpPick?.point) appendDumpVisit(dumpPick, "final-before-end");
      else if (requireNetworkOnly) return null;
    }

    let depotPick = currentPoint
      ? pickNearestRouteSequencerCandidate(
        currentPoint,
        [endDepot],
        graph,
        fallbackSpeedMph,
        { travelMode: "driving", fallbackServiceSpeedMph, requireNetworkOnly }
      )
      : { point: endDepot, minutes: 0, miles: 0, usedNetwork: false };
    if (!depotPick?.point && currentPoint && !requireNetworkOnly) {
      depotPick = {
        point: endDepot,
        minutes: estimateRouteSequencerTravelMinutes(currentPoint, endDepot, fallbackSpeedMph, {
          travelMode: "driving",
          fallbackServiceSpeedMph
        }),
        miles: estimateRouteSequencerTravelMiles(currentPoint, endDepot),
        usedNetwork: false
      };
    }
    if (!depotPick?.point) return null;

    const { driveMinutes, driveMiles } = applyDriveLeg(depotPick);
    events.push({
      type: "endDepot",
      name: endDepot.name,
      driveMinutes,
      driveMiles,
      trip: tripNumber,
      lat: Number(endDepot?.lat),
      lon: Number(endDepot?.lon)
    });
  }

  if (timeAtEndMinutes > 0) {
    totalStartEndFacilityMinutes += timeAtEndMinutes;
    totalFacilityMinutes += timeAtEndMinutes;
    events.push({ type: "timeAtEnd", facilityMinutes: timeAtEndMinutes, trip: tripNumber });
  }

  const split = routeSequencerSplitRouteDayKey(routeDayKey);
  return {
    routeDayKey,
    route: split.route,
    day: split.day,
    stopVisits,
    events,
    dumpVisits,
    totalDriveMinutes,
    totalDriveMiles,
    totalServiceMinutes,
    totalDumpFacilityMinutes,
    totalStartEndFacilityMinutes,
    totalBreakMinutes,
    totalOtherMinutes,
    totalFacilityMinutes,
    totalMinutes: totalDriveMinutes + totalServiceMinutes + totalFacilityMinutes + totalOtherMinutes,
    capacityTons: Number.isFinite(capacityTons) && capacityTons > 0 ? capacityTons : null,
    networkLegs,
    fallbackLegs
  };
}

function applyRouteSequencerPlansToRows(plans, outputFieldMapping, options = {}) {
  const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
  if (!rows.length) return 0;
  ensureCanonicalRecordFields(rows);
  ensureExistBackupColumns(rows);
  const outputFields = normalizeRouteSequencerOutputFields(outputFieldMapping || getRouteSequencerActiveOutputFields());
  const sequenceIncrement = normalizeRouteSequencerSequenceIncrement(options?.sequenceIncrement, 1);
  const resetFieldNames = [...new Set(Object.values(outputFields).map(name => String(name || "").trim()).filter(Boolean))];
  const idsToReset = new Set();
  Object.values(routeDayGroups).forEach(group => {
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(layer => {
      const rowId = Number(layer?._rowId);
      if (Number.isFinite(rowId)) idsToReset.add(rowId);
    });
  });

  idsToReset.forEach(rowId => {
    const row = rows[rowId];
    if (!row || typeof row !== "object") return;
    resetFieldNames.forEach(fieldName => {
      row[fieldName] = "";
    });
    row.SEQNO = "";
  });

  let updatedCount = 0;
  (Array.isArray(plans) ? plans : []).forEach(plan => {
    const dumpVisits = Number(plan?.dumpVisits) || 0;
    const visits = Array.isArray(plan?.stopVisits) ? plan.stopVisits : [];
    visits.forEach(visit => {
      const rowId = Number(visit?.rowId);
      if (!Number.isFinite(rowId)) return;
      const row = rows[rowId];
      if (!row || typeof row !== "object") return;
      const explicitSequenceValue = Number(visit?.outputSequenceValue);
      const visitSequenceOrder = Number(visit.sequence);
      const sequenceValue = Number.isFinite(explicitSequenceValue) && explicitSequenceValue > 0
        ? explicitSequenceValue
        : (
          Number.isFinite(visitSequenceOrder) && visitSequenceOrder > 0
            ? toRouteSequencerSequenceValue(visitSequenceOrder, sequenceIncrement)
            : ""
        );
      row[outputFields.sequence] = sequenceValue;
      row.SEQNO = sequenceValue;
      row[outputFields.trip] = Number(visit.trip) || "";
      row[outputFields.cumulativeLoad] = Number.isFinite(Number(visit.cumulativeLoad))
        ? Number(Number(visit.cumulativeLoad).toFixed(3))
        : "";
      row[outputFields.driveMinutes] = Number.isFinite(Number(visit.driveMinutes))
        ? Number(Number(visit.driveMinutes).toFixed(2))
        : "";
      row[outputFields.serviceMinutes] = Number.isFinite(Number(visit.serviceMinutes))
        ? Number(Number(visit.serviceMinutes).toFixed(2))
        : "";
      row[outputFields.totalMinutes] = Number.isFinite(Number(visit.totalMinutes))
        ? Number(Number(visit.totalMinutes).toFixed(2))
        : "";
      row[outputFields.dumpVisits] = dumpVisits;
      updatedCount += 1;
    });
  });

  window._attributeHeaders = getAttributeHeaders(rows);
  syncCurrentWorkbookSheetFromRows();
  queueWorkbookCloudSave("Route Sequencer");
  return updatedCount;
}

function refreshAllRecordPopupContent() {
  Object.values(routeDayGroups).forEach(group => {
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(layer => {
      syncRecordPopupContent(layer);
    });
  });
}

function escapeRouteSequencerHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderRouteSequencerResultsList(node, plans) {
  if (!node) return;
  const entries = Array.isArray(plans) ? plans : [];
  if (!entries.length) {
    node.innerHTML = '<div class="route-sequencer-results-empty">No sequencing results yet.</div>';
    return;
  }

  const formatNumber = (value, digits = 1, fallback = "-") => {
    const num = Number(value);
    return Number.isFinite(num) ? num.toFixed(digits) : fallback;
  };

  const formatDurationHms = (minutesValue, fallback = "00:00:00") => {
    const minutes = Number(minutesValue);
    if (!Number.isFinite(minutes)) return fallback;
    const totalSeconds = Math.max(0, Math.round(minutes * 60));
    const hours = Math.floor(totalSeconds / 3600);
    const minutesRemainder = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutesRemainder).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  };

  const computePlanTimeBreakdown = plan => {
    const events = Array.isArray(plan?.events) ? plan.events : [];
    const stops = Array.isArray(plan?.stopVisits) ? plan.stopVisits : [];
    const startMinutesFromEvents = events
      .filter(event => String(event?.type || "") === "timeAtStart")
      .reduce((sum, event) => sum + Math.max(0, Number(event?.facilityMinutes) || 0), 0);
    const endMinutesFromEvents = events
      .filter(event => String(event?.type || "") === "timeAtEnd")
      .reduce((sum, event) => sum + Math.max(0, Number(event?.facilityMinutes) || 0), 0);
    const dumpFacilityMinutesFromEvents = events
      .filter(event => String(event?.type || "") === "dump")
      .reduce((sum, event) => sum + Math.max(0, Number(event?.facilityMinutes) || 0), 0);
    const breakMinutesFromEvents = events
      .filter(event => String(event?.type || "") === "break")
      .reduce((sum, event) => sum + Math.max(0, Number(event?.minutes ?? event?.otherMinutes) || 0), 0);
    const otherMinutesFromEvents = events.reduce((sum, event) => {
      const type = String(event?.type || "");
      if (type === "break") return sum;
      if (type === "otherTime") {
        return sum + Math.max(0, Number(event?.minutes ?? event?.otherMinutes) || 0);
      }
      return sum + Math.max(0, Number(event?.otherMinutes) || 0);
    }, 0);

    const serviceFromStops = stops.reduce((sum, stop) => sum + Math.max(0, Number(stop?.serviceMinutes) || 0), 0);
    const demandTons = stops.reduce((sum, stop) => sum + Math.max(0, Number(stop?.demand) || 0), 0);
    const driveMinutes = Math.max(0, Number(plan?.totalDriveMinutes) || 0);
    const driveMilesFromEvents = events.reduce((sum, event) => sum + Math.max(0, Number(event?.driveMiles) || 0), 0);
    const driveMiles = Number.isFinite(Number(plan?.totalDriveMiles))
      ? Math.max(0, Number(plan.totalDriveMiles))
      : driveMilesFromEvents;
    const serviceMinutes = Math.max(serviceFromStops, Math.max(0, Number(plan?.totalServiceMinutes) || 0));
    const dumpFacilityMinutes = Number.isFinite(Number(plan?.totalDumpFacilityMinutes))
      ? Math.max(0, Number(plan.totalDumpFacilityMinutes))
      : dumpFacilityMinutesFromEvents;
    const breakMinutes = Number.isFinite(Number(plan?.totalBreakMinutes))
      ? Math.max(0, Number(plan.totalBreakMinutes))
      : breakMinutesFromEvents;
    const otherMinutesFromPlanTotal = Number.isFinite(Number(plan?.totalOtherMinutes))
      ? Math.max(0, Number(plan.totalOtherMinutes))
      : NaN;
    const otherMinutes = Number.isFinite(otherMinutesFromPlanTotal)
      ? Math.max(0, otherMinutesFromPlanTotal - breakMinutes)
      : otherMinutesFromEvents;
    const startMinutes = startMinutesFromEvents;
    const endMinutes = endMinutesFromEvents;
    const facilityMinutesFromPlan = Number.isFinite(Number(plan?.totalFacilityMinutes))
      ? Math.max(0, Number(plan.totalFacilityMinutes))
      : NaN;
    const facilityMinutesFromEvents = dumpFacilityMinutes + startMinutes + endMinutes;
    const facilityMinutesResolved = Number.isFinite(facilityMinutesFromPlan)
      ? facilityMinutesFromPlan
      : facilityMinutesFromEvents;
    const totalMinutesResolved = driveMinutes + serviceMinutes + facilityMinutesResolved + breakMinutes + otherMinutes;

    return {
      driveMinutes,
      driveMiles,
      serviceMinutes,
      dumpFacilityMinutes,
      breakMinutes,
      startMinutes,
      endMinutes,
      facilityMinutes: facilityMinutesResolved,
      otherMinutes,
      totalMinutes: totalMinutesResolved,
      demandTons
    };
  };

  const formatEventAction = (event, index) => {
    const type = String(event?.type || "");
    if (type === "stop") return `Stop ${Number(event?.sequence) || (index + 1)}`;
    if (type === "dump") return "Dump Facility";
    if (type === "startDepot") return "Start Facility";
    if (type === "endDepot") return "End Facility";
    if (type === "timeAtStart") return "Time at Start";
    if (type === "timeAtEnd") return "Time at End";
    if (type === "break") return "Driver Break";
    return type || `Step ${index + 1}`;
  };

  const formatEventNotes = event => {
    const type = String(event?.type || "");
    const name = String(event?.name || "").trim();
    if (type === "dump") {
      const beforeLoad = Number(event?.beforeLoad);
      const reason = String(event?.reason || "").trim();
      const reasonLabel = reason
        ? reason.replace(/-/g, " ").replace(/\b\w/g, ch => ch.toUpperCase())
        : "";
      const loadText = Number.isFinite(beforeLoad) ? `Load ${beforeLoad.toFixed(2)} tons` : "";
      return [name, reasonLabel, loadText].filter(Boolean).join(" | ") || "Capacity dump";
    }
    if (type === "stop") {
      const penaltyNotes = String(event?.penaltyNotes || "").trim();
      return [name || "Stop", penaltyNotes].filter(Boolean).join(" | ");
    }
    if (type === "startDepot" || type === "endDepot") return name || "Facility";
    if (type === "break") {
      const breakMinutes = Math.max(0, Number(event?.minutes ?? event?.otherMinutes) || 0);
      const breakLabel = Number.isFinite(breakMinutes) && breakMinutes > 0
        ? formatDurationHms(breakMinutes)
        : "";
      return [name || "Driver idle break", breakLabel].filter(Boolean).join(" | ");
    }
    return name || "";
  };

  const sortedEntries = entries.slice().sort((a, b) => {
    const keyA = `${a?.route || ""}|${a?.day || ""}`;
    const keyB = `${b?.route || ""}|${b?.day || ""}`;
    return keyA.localeCompare(keyB, undefined, { sensitivity: "base", numeric: true });
  });

  const summaryRowsHtml = sortedEntries.map(plan => {
    const route = plan?.route || "Unassigned";
    const day = plan?.day || "No Day";
    const stopCount = Number(plan?.stopVisits?.length) || 0;
    const dumpCount = Number(plan?.dumpVisits) || 0;
    const tripCount = new Set(
      (Array.isArray(plan?.stopVisits) ? plan.stopVisits : [])
        .map(visit => Number(visit?.trip))
        .filter(value => Number.isFinite(value) && value > 0)
    ).size || 1;
    const capacityText = Number.isFinite(Number(plan?.capacityTons))
      ? Number(plan.capacityTons).toFixed(2)
      : "-";
    const breakdown = computePlanTimeBreakdown(plan);
    return `
      <tr>
        <td>${escapeRouteSequencerHtml(route)}</td>
        <td>${escapeRouteSequencerHtml(day)}</td>
        <td>${stopCount.toLocaleString()}</td>
        <td>${tripCount.toLocaleString()}</td>
        <td>${dumpCount.toLocaleString()}</td>
        <td>${capacityText}</td>
        <td>${formatNumber(breakdown.demandTons, 2)}</td>
        <td>${formatDurationHms(breakdown.driveMinutes)}</td>
        <td>${formatNumber(breakdown.driveMiles, 2)}</td>
        <td>${formatDurationHms(breakdown.serviceMinutes)}</td>
        <td>${formatDurationHms(breakdown.dumpFacilityMinutes)}</td>
        <td>${formatDurationHms(breakdown.breakMinutes)}</td>
        <td>${formatDurationHms(breakdown.startMinutes)}</td>
        <td>${formatDurationHms(breakdown.endMinutes)}</td>
        <td>${formatDurationHms(breakdown.otherMinutes)}</td>
        <td>${formatDurationHms(breakdown.totalMinutes)}</td>
      </tr>
    `;
  }).join("");

  const summaryTableHtml = `
    <section class="route-sequencer-summary-card">
      <div class="route-sequencer-report-key">Sequencing Route Report (Excel Format)</div>
      <div class="route-sequencer-summary-table-wrap">
        <table class="route-sequencer-summary-table">
          <thead>
            <tr>
              <th>Route</th>
              <th>Day</th>
              <th>Stops</th>
              <th>Trips</th>
              <th>Dumps</th>
              <th>Cap (tons)</th>
              <th>Demand (tons)</th>
              <th>Drive (HH:MM:SS)</th>
              <th>Miles</th>
              <th>Service (HH:MM:SS)</th>
              <th>Dump Facility (HH:MM:SS)</th>
              <th>Break (HH:MM:SS)</th>
              <th>Start (HH:MM:SS)</th>
              <th>End (HH:MM:SS)</th>
              <th>Other (HH:MM:SS)</th>
              <th>Total (HH:MM:SS)</th>
            </tr>
          </thead>
          <tbody>${summaryRowsHtml}</tbody>
        </table>
      </div>
      <div class="route-sequencer-report-truncated">Time fields are shown as HH:MM:SS. Total = Drive + Service + Dump Facility + Break + Start + End + Other. Dump Facility time is counted on every dump visit (for example, 00:25:00 x 2 dumps = 00:50:00). Break time is route-level idle time. Drive includes travel between stops, to/from dump locations, from start facility to first stop, and from last stop to end facility. Other includes class-based penalties such as cross-penalty and U-turn penalties.</div>
    </section>
  `;

  const detailHtml = sortedEntries.map(plan => {
    const key = `${plan.route || "Unassigned"} | ${plan.day || "No Day"}`;
    const stopCount = Number(plan.stopVisits?.length) || 0;
    const dumpCount = Number(plan.dumpVisits) || 0;
    const capacityText = Number.isFinite(Number(plan.capacityTons))
      ? `${Number(plan.capacityTons).toFixed(2)} tons`
      : "Not set";
    const breakdown = computePlanTimeBreakdown(plan);
    const tripCount = new Set(
      (Array.isArray(plan.stopVisits) ? plan.stopVisits : [])
        .map(visit => Number(visit?.trip))
        .filter(value => Number.isFinite(value) && value > 0)
    ).size || 1;

    const events = Array.isArray(plan.events) ? plan.events : [];
    const maxReportRows = 450;
    const reportRows = events.slice(0, maxReportRows);
    let runningLoadTons = 0;
    const reportRowsHtml = reportRows.map((event, index) => {
      const trip = Number(event?.trip);
      const type = String(event?.type || "");
      const rowId = Number(event?.rowId);
      const lat = Number(event?.lat);
      const lon = Number(event?.lon);
      const isStopSelectable = type === "stop" && Number.isFinite(rowId);
      const isFacilitySelectable = (
        type === "dump"
        || type === "startDepot"
        || type === "endDepot"
      ) && Number.isFinite(lat) && Number.isFinite(lon);
      const isClickable = isStopSelectable || isFacilitySelectable;
      const rowClass = isClickable ? "route-sequencer-result-row-clickable" : "";
      const rowAttributes = [
        isClickable ? 'data-rs-clickable="1"' : "",
        isClickable ? `data-rs-event-type="${type.replace(/"/g, "&quot;")}"` : "",
        isStopSelectable ? `data-rs-row-id="${rowId}"` : "",
        isFacilitySelectable ? `data-rs-lat="${lat}"` : "",
        isFacilitySelectable ? `data-rs-lon="${lon}"` : "",
        isStopSelectable
          ? 'title="Click to select this stop on the map and in the Attribute Table."'
          : (
            isFacilitySelectable
              ? 'title="Click to focus this facility on the map."'
              : ""
          )
      ].filter(Boolean).join(" ");
      const demandTons = type === "stop"
        ? Math.max(0, Number(event?.demand) || 0)
        : 0;
      if (type === "stop") {
        const eventCumulative = Number(event?.cumulativeLoad);
        runningLoadTons = Number.isFinite(eventCumulative)
          ? Math.max(0, eventCumulative)
          : (runningLoadTons + demandTons);
      } else if (type === "dump") {
        runningLoadTons = 0;
      }
      const otherMinutes = type === "otherTime"
        ? Math.max(0, Number(event?.minutes) || 0)
        : Math.max(0, Number(event?.otherMinutes) || 0);
      return `
        <tr class="${rowClass}" ${rowAttributes}>
          <td>${index + 1}</td>
          <td>${escapeRouteSequencerHtml(formatEventAction(event, index))}</td>
          <td>${Number.isFinite(trip) && trip > 0 ? trip : "-"}</td>
          <td>${formatDurationHms(event?.driveMinutes)}</td>
          <td>${formatDurationHms(event?.serviceMinutes)}</td>
          <td>${formatDurationHms(event?.facilityMinutes)}</td>
          <td>${formatDurationHms(otherMinutes)}</td>
          <td>${formatNumber(demandTons, 2, "0.00")}</td>
          <td>${formatNumber(runningLoadTons, 2, "0.00")}</td>
          <td>${escapeRouteSequencerHtml(formatEventNotes(event))}</td>
        </tr>
      `;
    }).join("");

    const truncatedHtml = events.length > maxReportRows
      ? `<div class="route-sequencer-report-truncated">Showing first ${maxReportRows.toLocaleString()} events of ${events.length.toLocaleString()}.</div>`
      : "";

    return `
      <section class="route-sequencer-report-card">
        <div class="route-sequencer-report-head">
          <div class="route-sequencer-report-key">${escapeRouteSequencerHtml(key)}</div>
          <div class="route-sequencer-report-metrics">
            <span>Stops: ${stopCount.toLocaleString()}</span>
            <span>Dumps: ${dumpCount.toLocaleString()}</span>
            <span>Trips: ${tripCount.toLocaleString()}</span>
            <span>Capacity: ${escapeRouteSequencerHtml(capacityText)}</span>
            <span>Demand: ${formatNumber(breakdown.demandTons, 2)} tons</span>
            <span>Drive: ${formatDurationHms(breakdown.driveMinutes)}</span>
            <span>Miles: ${formatNumber(breakdown.driveMiles, 2)} mi</span>
            <span>Service: ${formatDurationHms(breakdown.serviceMinutes)}</span>
            <span>Break: ${formatDurationHms(breakdown.breakMinutes)}</span>
            <span>Start: ${formatDurationHms(breakdown.startMinutes)}</span>
            <span>End: ${formatDurationHms(breakdown.endMinutes)}</span>
            <span>Dump Facility: ${formatDurationHms(breakdown.dumpFacilityMinutes)}</span>
            <span>Other: ${formatDurationHms(breakdown.otherMinutes)}</span>
            <span>Total: ${formatDurationHms(breakdown.totalMinutes)}</span>
          </div>
        </div>
        <div class="route-sequencer-report-table-wrap">
          <table class="route-sequencer-report-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Action</th>
                <th>Trip</th>
                <th>Drive (HH:MM:SS)</th>
                <th>Service (HH:MM:SS)</th>
                <th>Facility (HH:MM:SS)</th>
                <th>Other (HH:MM:SS)</th>
                <th>Demand (tons)</th>
                <th>Load (tons)</th>
                <th>Notes</th>
              </tr>
            </thead>
            <tbody>${reportRowsHtml}</tbody>
          </table>
        </div>
        ${truncatedHtml}
      </section>
    `;
  }).join("");
  node.innerHTML = `${summaryTableHtml}${detailHtml}`;
}

function loadRouteSequencerSettings() {
  const raw = storageGet(ROUTE_SEQUENCER_SETTINGS_KEY);
  if (!raw) {
    setRouteSequencerActiveOutputFields(ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS);
    return {};
  }
  try {
    const parsed = JSON.parse(raw);
    const settings = parsed && typeof parsed === "object" ? parsed : {};
    const speedDefaultsVersion = Number(settings.speedDefaultsVersion || 0);
    if (speedDefaultsVersion < ROUTE_SEQUENCER_SPEED_DEFAULTS_VERSION) {
      settings.drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
        ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
        ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
      );
      settings.serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
        ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
        ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
      );
    } else {
      settings.drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
        settings.drivingSpeedCatMap || settings.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
        ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
      );
      settings.serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
        settings.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
        ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
      );
    }
    settings.classSpeedMap = { ...settings.drivingSpeedCatMap };
    settings.speedDefaultsVersion = ROUTE_SEQUENCER_SPEED_DEFAULTS_VERSION;

    const streetDefaultsVersion = Number(settings.streetAttributeDefaultsVersion || 0);
    if (streetDefaultsVersion < ROUTE_SEQUENCER_STREET_ATTR_DEFAULTS_VERSION) {
      settings.streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(
        ROUTE_SEQUENCER_DEFAULT_STREET_ATTRIBUTE_DEFAULTS
      );
    } else {
      settings.streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(
        settings.streetAttributeDefaults || settings
      );
    }
    settings.streetAttributeDefaultsVersion = ROUTE_SEQUENCER_STREET_ATTR_DEFAULTS_VERSION;
    settings.serviceStopsBeforeDeadhead = settings.serviceStopsBeforeDeadhead == null
      ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
      : settings.serviceStopsBeforeDeadhead !== false;
    settings.meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(
      settings.meanderOverrideDefaults
    );
    setRouteSequencerActiveOutputFields(settings.outputFields || ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS);
    return settings;
  } catch {
    setRouteSequencerActiveOutputFields(ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS);
    return {};
  }
}

function saveRouteSequencerSettings(settings) {
  const payload = settings && typeof settings === "object" ? settings : {};
  payload.drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
    payload.drivingSpeedCatMap || payload.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
    ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
  );
  payload.serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
    payload.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
    ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
  );
  payload.classSpeedMap = { ...payload.drivingSpeedCatMap };
  payload.speedDefaultsVersion = ROUTE_SEQUENCER_SPEED_DEFAULTS_VERSION;
  payload.outputFields = normalizeRouteSequencerOutputFields(payload.outputFields);
  payload.streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(payload.streetAttributeDefaults || payload);
  payload.streetAttributeDefaultsVersion = ROUTE_SEQUENCER_STREET_ATTR_DEFAULTS_VERSION;
  payload.serviceStopsBeforeDeadhead = payload.serviceStopsBeforeDeadhead == null
    ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
    : payload.serviceStopsBeforeDeadhead !== false;
  payload.meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(
    payload.meanderOverrideDefaults
  );
  storageSet(ROUTE_SEQUENCER_SETTINGS_KEY, JSON.stringify(payload));
  setRouteSequencerActiveOutputFields(payload.outputFields);
}

function initFacilitiesControls() {
  const openBtn = document.getElementById("facilitiesBtn");
  const openBtnMobile = document.getElementById("facilitiesBtnMobile");
  const modal = document.getElementById("facilitiesModal");
  const closeBtn = document.getElementById("facilitiesCloseBtn");
  const nameInput = document.getElementById("facilityNameInput");
  const roleSelect = document.getElementById("facilityRoleSelect");
  const timeInput = document.getElementById("facilityTimeInput");
  const latInput = document.getElementById("facilityLatInput");
  const lonInput = document.getElementById("facilityLonInput");
  const mapCenterBtn = document.getElementById("facilityUseMapCenterBtn");
  const saveBtn = document.getElementById("facilitySaveBtn");
  const clearBtn = document.getElementById("facilityClearBtn");
  const listSelect = document.getElementById("facilityListSelect");
  const loadSelectedBtn = document.getElementById("facilityLoadSelectedBtn");
  const removeBtn = document.getElementById("facilityRemoveBtn");
  const statusNode = document.getElementById("facilitiesStatus");

  if (
    !openBtn || !modal || !closeBtn || !nameInput || !roleSelect || !timeInput ||
    !latInput || !lonInput ||
    !mapCenterBtn || !saveBtn || !clearBtn || !listSelect || !loadSelectedBtn ||
    !removeBtn || !statusNode
  ) {
    return;
  }

  if (openBtn.dataset.facilitiesBound === "1") return;
  openBtn.dataset.facilitiesBound = "1";

  const setStatus = (message, kind = "") => {
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (kind === "error") statusNode.classList.add("error");
    if (kind === "success") statusNode.classList.add("success");
  };

  const isTexasFacilityLatLng = (lat, lon) => {
    if (typeof isTexasLatLng === "function") {
      try {
        return !!isTexasLatLng(lat, lon);
      } catch {
        // Fall through to bounds fallback.
      }
    }
    const latNum = Number(lat);
    const lonNum = Number(lon);
    if (!Number.isFinite(latNum) || !Number.isFinite(lonNum)) return false;
    // Texas bounding box fallback.
    return latNum >= 25.8372 && latNum <= 36.5007 && lonNum >= -106.6456 && lonNum <= -93.5080;
  };

  const refreshSequencerFacilityUi = () => {
    if (typeof window.__refreshRouteSequencerFacilitiesUi === "function") {
      window.__refreshRouteSequencerFacilitiesUi();
    }
  };

  const syncTimeInputForRole = () => {
    const role = normalizeRouteSequencerFacilityRole(roleSelect.value);
    const isDump = role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP;
    if (!isDump) {
      timeInput.value = "00:00";
    }
    timeInput.disabled = !isDump;
    timeInput.title = isDump
      ? "Time spent at dump location (MM:SS)."
      : "Start/End facilities use Time at Start and Time at End in the solver.";
  };

  const clearForm = () => {
    nameInput.value = "";
    roleSelect.value = ROUTE_SEQUENCER_FACILITY_ROLE_START_END;
    timeInput.value = "00:00";
    latInput.value = "";
    lonInput.value = "";
    listSelect.value = "";
    syncTimeInputForRole();
  };

  const renderFacilityList = (selectedId = "") => {
    const facilities = getRouteSequencerFacilities()
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base", numeric: true }));
    listSelect.innerHTML = "";
    if (!facilities.length) {
      listSelect.innerHTML = '<option value="">No facilities saved yet</option>';
      loadSelectedBtn.disabled = true;
      removeBtn.disabled = true;
      return;
    }
    facilities.forEach(item => {
      const option = document.createElement("option");
      option.value = String(item.id);
      const roleLabel = getRouteSequencerFacilityRoleLabel(item.role);
      const isDump = normalizeRouteSequencerFacilityRole(item.role) === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP;
      const timeLabel = isDump ? ` | Dump ${formatRouteSequencerMinutesClock(item.facilityMinutes)}` : "";
      option.textContent = `${item.name} | ${roleLabel}${timeLabel} | ${Number(item.lat).toFixed(5)}, ${Number(item.lon).toFixed(5)}`;
      listSelect.appendChild(option);
    });
    if (selectedId && [...listSelect.options].some(opt => opt.value === selectedId)) {
      listSelect.value = selectedId;
    } else if (listSelect.options.length) {
      listSelect.value = listSelect.options[0].value;
    }
    loadSelectedBtn.disabled = false;
    removeBtn.disabled = false;
  };

  const loadSelectedFacilityIntoForm = () => {
    const selectedId = String(listSelect.value || "");
    if (!selectedId) {
      setStatus("Choose a facility to load.", "error");
      return;
    }
    const facility = getRouteSequencerFacilities().find(item => String(item.id) === selectedId);
    if (!facility) {
      setStatus("Selected facility was not found.", "error");
      return;
    }
    nameInput.value = String(facility.name || "");
    roleSelect.value = normalizeRouteSequencerFacilityRole(facility.role);
    timeInput.value = formatRouteSequencerMinutesClock(facility.facilityMinutes);
    syncTimeInputForRole();
    latInput.value = String(Number(facility.lat).toFixed(6));
    lonInput.value = String(Number(facility.lon).toFixed(6));
    setStatus(`Loaded facility "${facility.name}".`);
  };

  const saveFacility = () => {
    const name = normalizeRouteSequencerName(nameInput.value, "Facility");
    const role = normalizeRouteSequencerFacilityRole(roleSelect.value);
    const facilityMinutes = role === ROUTE_SEQUENCER_FACILITY_ROLE_DUMP
      ? parseRouteSequencerMinutesValue(timeInput.value, 0)
      : 0;
    const lat = Number(latInput.value);
    const lon = Number(lonInput.value);
    const selectedId = String(listSelect.value || "");

    if (!name || name === "Facility") {
      setStatus("Enter a facility name.", "error");
      return;
    }
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
      setStatus("Enter valid latitude and longitude.", "error");
      return;
    }
    if (!isTexasFacilityLatLng(lat, lon)) {
      setStatus("Facility location must be inside Texas bounds.", "error");
      return;
    }

    const facilities = getRouteSequencerFacilities();
    const existingIndex = facilities.findIndex(item => String(item.id) === selectedId);
    const baseId = existingIndex >= 0
      ? String(facilities[existingIndex].id)
      : `facility_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

    const updatedFacility = normalizeRouteSequencerFacilityEntry({
      id: baseId,
      name,
      role,
      facilityMinutes,
      lat,
      lon
    }, "Facility");
    if (!updatedFacility) {
      setStatus("Unable to save facility. Check name and coordinates.", "error");
      return;
    }

    if (existingIndex >= 0) {
      facilities[existingIndex] = updatedFacility;
    } else {
      facilities.push(updatedFacility);
    }
    setRouteSequencerFacilities(facilities);
    renderFacilityList(String(updatedFacility.id));
    refreshSequencerFacilityUi();
    setStatus(
      existingIndex >= 0
        ? `Updated facility "${updatedFacility.name}".`
        : `Added facility "${updatedFacility.name}".`,
      "success"
    );
  };

  const removeSelectedFacility = () => {
    const selectedId = String(listSelect.value || "");
    if (!selectedId) {
      setStatus("Choose a facility to remove.", "error");
      return;
    }
    const facilities = getRouteSequencerFacilities();
    const target = facilities.find(item => String(item.id) === selectedId);
    if (!target) {
      setStatus("Selected facility was not found.", "error");
      return;
    }
    const next = facilities.filter(item => String(item.id) !== selectedId);
    setRouteSequencerFacilities(next);
    renderFacilityList("");
    refreshSequencerFacilityUi();
    clearForm();
    setStatus(`Removed facility "${target.name}".`, "success");
  };

  const useMapCenter = () => {
    const center = map?.getCenter?.();
    if (!center) {
      setStatus("Map center is unavailable.", "error");
      return;
    }
    latInput.value = String(Number(center.lat).toFixed(6));
    lonInput.value = String(Number(center.lng).toFixed(6));
    setStatus("Loaded coordinates from map center.");
  };

  const openModal = () => {
    renderFacilityList("");
    clearForm();
    setStatus("Manage reusable facilities using Texas latitude/longitude.");
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  openBtn.addEventListener("click", openModal);
  if (openBtnMobile) openBtnMobile.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });
  roleSelect.addEventListener("change", () => {
    syncTimeInputForRole();
  });
  mapCenterBtn.addEventListener("click", useMapCenter);
  saveBtn.addEventListener("click", () => {
    try {
      saveFacility();
    } catch (error) {
      console.error("Failed to save facility:", error);
      setStatus(`Unable to save facility: ${error?.message || error}`, "error");
    }
  });
  clearBtn.addEventListener("click", () => {
    clearForm();
    setStatus("Form cleared.");
  });
  loadSelectedBtn.addEventListener("click", loadSelectedFacilityIntoForm);
  removeBtn.addEventListener("click", removeSelectedFacility);
  listSelect.addEventListener("dblclick", loadSelectedFacilityIntoForm);
  window.addEventListener("keydown", event => {
    if (modal.style.display !== "flex") return;
    if (event.key === "Escape") closeModal();
  });
}

function initRouteSequencerControls() {
  const openBtn = document.getElementById("routeSequencerBtn");
  const openBtnMobile = document.getElementById("routeSequencerBtnMobile");
  const modal = document.getElementById("routeSequencerModal");
  const closeBtn = document.getElementById("routeSequencerCloseBtn");
  const evaluateBtn = document.getElementById("routeSequencerEvaluateBtn");
  const runBtn = document.getElementById("routeSequencerRunBtn");
  const demandFieldSelect = document.getElementById("routeSequencerDemandField");
  const serviceFieldSelect = document.getElementById("routeSequencerServiceTimeField");
  const sequenceOutputSelect = document.getElementById("routeSequencerSequenceOutputField");
  const tripOutputSelect = document.getElementById("routeSequencerTripOutputField");
  const cumulativeOutputSelect = document.getElementById("routeSequencerCumulativeOutputField");
  const driveOutputSelect = document.getElementById("routeSequencerDriveOutputField");
  const serviceOutputSelect = document.getElementById("routeSequencerServiceOutputField");
  const totalOutputSelect = document.getElementById("routeSequencerTotalOutputField");
  const dumpVisitsOutputSelect = document.getElementById("routeSequencerDumpVisitsOutputField");
  const scopeListSelect = document.getElementById("routeSequencerScopeList");
  const scopeAllBtn = document.getElementById("routeSequencerScopeAllBtn");
  const scopeNoneBtn = document.getElementById("routeSequencerScopeNoneBtn");
  const scopeVisibleBtn = document.getElementById("routeSequencerScopeVisibleBtn");
  const formatHintNode = document.getElementById("routeSequencerFormatHint");
  const capacityInput = document.getElementById("routeSequencerCapacityInput");
  const defaultSpeedInput = document.getElementById("routeSequencerDefaultSpeedInput");
  const sequenceIncrementInput = document.getElementById("routeSequencerSequenceIncrementInput");
  const speedClassFieldInput = document.getElementById("routeSequencerSpeedClassField");
  const streetClassFieldInput = document.getElementById("routeSequencerStreetClassField");
  const serviceStopsBeforeDeadheadToggle = document.getElementById("routeSequencerServiceStopsBeforeDeadhead");
  const allowUTurnsOnlyInCuldesacsToggle = document.getElementById("routeSequencerAllowUTurnsCuldesacs");
  const discourageUTurnsToggle = document.getElementById("routeSequencerDiscourageUTurnsOutsideIntersections");
  const meanderServiceInput = document.getElementById("routeSequencerMeanderServiceInput");
  const sameSideServiceInput = document.getElementById("routeSequencerSameSideServiceInput");
  const meanderOverrideDrivingInput = document.getElementById("routeSequencerMeanderOverrideDriving");
  const meanderOverrideWalkingInput = document.getElementById("routeSequencerMeanderOverrideWalking");
  const uTurnProrateDrivingInput = document.getElementById("routeSequencerUTurnProrateDriving");
  const uTurnProrateWalkingInput = document.getElementById("routeSequencerUTurnProrateWalking");
  const dumpSourceSelect = document.getElementById("routeSequencerDumpSource");
  const startDepotSelect = document.getElementById("routeSequencerStartDepotSelect");
  const endDepotSelect = document.getElementById("routeSequencerEndDepotSelect");
  const timeAtStartInput = document.getElementById("routeSequencerTimeAtStartInput");
  const timeAtEndInput = document.getElementById("routeSequencerTimeAtEndInput");
  const breaksInput = document.getElementById("routeSequencerBreaksInput");
  const dumpBeforeEndToggle = document.getElementById("routeSequencerDumpBeforeEnd");
  const depotNameInput = document.getElementById("routeSequencerDepotNameInput");
  const addDepotBtn = document.getElementById("routeSequencerAddDepotBtn");
  const removeDepotBtn = document.getElementById("routeSequencerRemoveDepotBtn");
  const dumpNameInput = document.getElementById("routeSequencerDumpNameInput");
  const customDumpSelect = document.getElementById("routeSequencerCustomDumpSelect");
  const addDumpBtn = document.getElementById("routeSequencerAddDumpBtn");
  const removeDumpBtn = document.getElementById("routeSequencerRemoveDumpBtn");
  const statusNode = document.getElementById("routeSequencerStatus");
  const resultsNode = document.getElementById("routeSequencerResults");
  const drivingSpeedCatInputs = {
    1: document.getElementById("routeSequencerDrivingSpeedCat1"),
    2: document.getElementById("routeSequencerDrivingSpeedCat2"),
    3: document.getElementById("routeSequencerDrivingSpeedCat3"),
    4: document.getElementById("routeSequencerDrivingSpeedCat4"),
    5: document.getElementById("routeSequencerDrivingSpeedCat5"),
    6: document.getElementById("routeSequencerDrivingSpeedCat6"),
    7: document.getElementById("routeSequencerDrivingSpeedCat7"),
    8: document.getElementById("routeSequencerDrivingSpeedCat8")
  };
  const serviceSpeedCatInputs = {
    1: document.getElementById("routeSequencerServiceSpeedCat1"),
    2: document.getElementById("routeSequencerServiceSpeedCat2"),
    3: document.getElementById("routeSequencerServiceSpeedCat3"),
    4: document.getElementById("routeSequencerServiceSpeedCat4"),
    5: document.getElementById("routeSequencerServiceSpeedCat5"),
    6: document.getElementById("routeSequencerServiceSpeedCat6"),
    7: document.getElementById("routeSequencerServiceSpeedCat7"),
    8: document.getElementById("routeSequencerServiceSpeedCat8")
  };
  const classAttrInputs = {
    1: {
      travelTypesAllowed: document.getElementById("routeSequencerClass1TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass1ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass1CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass1UTurnRestricted")
    },
    2: {
      travelTypesAllowed: document.getElementById("routeSequencerClass2TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass2ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass2CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass2UTurnRestricted")
    },
    3: {
      travelTypesAllowed: document.getElementById("routeSequencerClass3TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass3ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass3CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass3UTurnRestricted")
    },
    4: {
      travelTypesAllowed: document.getElementById("routeSequencerClass4TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass4ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass4CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass4UTurnRestricted")
    },
    5: {
      travelTypesAllowed: document.getElementById("routeSequencerClass5TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass5ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass5CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass5UTurnRestricted")
    },
    6: {
      travelTypesAllowed: document.getElementById("routeSequencerClass6TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass6ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass6CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass6UTurnRestricted")
    },
    7: {
      travelTypesAllowed: document.getElementById("routeSequencerClass7TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass7ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass7CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass7UTurnRestricted")
    },
    8: {
      travelTypesAllowed: document.getElementById("routeSequencerClass8TravelTypes"),
      serviceSide: document.getElementById("routeSequencerClass8ServiceSide"),
      crossPenalty: document.getElementById("routeSequencerClass8CrossPenalty"),
      uTurnsRestricted: document.getElementById("routeSequencerClass8UTurnRestricted")
    }
  };
  const outputFieldSelectMap = {
    sequence: sequenceOutputSelect,
    trip: tripOutputSelect,
    cumulativeLoad: cumulativeOutputSelect,
    driveMinutes: driveOutputSelect,
    serviceMinutes: serviceOutputSelect,
    totalMinutes: totalOutputSelect,
    dumpVisits: dumpVisitsOutputSelect
  };

  if (
    !openBtn || !modal || !closeBtn || !evaluateBtn || !runBtn || !demandFieldSelect || !serviceFieldSelect ||
    !sequenceOutputSelect || !tripOutputSelect || !cumulativeOutputSelect || !driveOutputSelect ||
    !serviceOutputSelect || !totalOutputSelect || !dumpVisitsOutputSelect || !formatHintNode ||
    !scopeListSelect || !scopeAllBtn || !scopeNoneBtn || !scopeVisibleBtn ||
    !capacityInput || !defaultSpeedInput || !sequenceIncrementInput || !speedClassFieldInput ||
    !streetClassFieldInput || !serviceStopsBeforeDeadheadToggle ||
    !allowUTurnsOnlyInCuldesacsToggle || !discourageUTurnsToggle ||
    !meanderServiceInput || !sameSideServiceInput ||
    !meanderOverrideDrivingInput || !meanderOverrideWalkingInput ||
    !uTurnProrateDrivingInput || !uTurnProrateWalkingInput ||
    !dumpSourceSelect || !startDepotSelect || !endDepotSelect ||
    !timeAtStartInput || !timeAtEndInput || !breaksInput ||
    !dumpBeforeEndToggle || !depotNameInput || !addDepotBtn || !removeDepotBtn ||
    !dumpNameInput || !customDumpSelect || !addDumpBtn || !removeDumpBtn ||
    !statusNode || !resultsNode ||
    Object.values(drivingSpeedCatInputs).some(node => !node) ||
    Object.values(serviceSpeedCatInputs).some(node => !node) ||
    Object.values(classAttrInputs).some(group => !group
      || !group.travelTypesAllowed
      || !group.serviceSide
      || !group.crossPenalty
      || !group.uTurnsRestricted
    )
  ) {
    return;
  }

  if (openBtn.dataset.routeSequencerBound === "1") return;
  openBtn.dataset.routeSequencerBound = "1";
  loadRouteSequencerSettings();
  let routeSequencerTaskRunning = false;
  let routeSequencerOpenCheckRunning = false;

  const setStatus = (message, kind = "") => {
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (kind === "error") statusNode.classList.add("error");
    if (kind === "success") statusNode.classList.add("success");
    if (routeSequencerTaskRunning && typeof window.showLoading === "function") {
      window.showLoading(String(message || "Running solver..."));
    }
  };

  const yieldRouteSequencerUi = async () => {
    await new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
  };

  const setRouteSequencerLoading = (active, message = "") => {
    const next = !!active;
    routeSequencerTaskRunning = next;
    evaluateBtn.disabled = next;
    runBtn.disabled = next;
    if (next) {
      if (typeof window.showLoading === "function") {
        window.showLoading(String(message || "Running solver..."));
      }
      return;
    }
    if (typeof window.hideLoading === "function") {
      window.hideLoading();
    }
  };

  const setActiveSequenceResultRow = row => {
    resultsNode
      .querySelectorAll(".route-sequencer-report-table tbody tr.route-sequencer-result-row-active")
      .forEach(node => node.classList.remove("route-sequencer-result-row-active"));
    if (row) row.classList.add("route-sequencer-result-row-active");
  };

  const focusSequencerFacilityFromResult = (lat, lon) => {
    const latNum = Number(lat);
    const lonNum = Number(lon);
    if (!Number.isFinite(latNum) || !Number.isFinite(lonNum) || !map) return false;

    let bestLayer = null;
    let bestDistanceMiles = Infinity;
    const considerLayerGroup = layerGroup => {
      if (!layerGroup || typeof layerGroup.eachLayer !== "function") return;
      layerGroup.eachLayer(layer => {
        const ll = getLayerLatLng(layer);
        if (!ll) return;
        const distanceMiles = haversineMiles(latNum, lonNum, ll.lat, ll.lng);
        if (!Number.isFinite(distanceMiles)) return;
        if (!bestLayer || distanceMiles < bestDistanceMiles) {
          bestLayer = layer;
          bestDistanceMiles = distanceMiles;
        }
      });
    };

    considerLayerGroup(routeSequencerFacilitiesLayerGroup);
    considerLayerGroup(texasLandfillsLayerGroup);

    const fallbackLatLng = L.latLng(latNum, lonNum);
    const targetZoom = Math.max(Number(map.getZoom()) || 0, 15);
    const matchedThresholdMiles = 2;
    if (bestLayer && Number.isFinite(bestDistanceMiles) && bestDistanceMiles <= matchedThresholdMiles) {
      const markerLatLng = getLayerLatLng(bestLayer) || fallbackLatLng;
      map.setView(markerLatLng, targetZoom, { animate: true });
      bestLayer.openPopup?.();
      try {
        bestLayer.bringToFront?.();
      } catch {
        // Non-path layers may not support z-order operations.
      }
      return true;
    }

    map.setView(fallbackLatLng, targetZoom, { animate: true });
    return false;
  };

  const selectSequencerStopFromResult = rowId => {
    const id = Number(rowId);
    if (!Number.isFinite(id)) return false;
    if (attributeTableMode !== "records") {
      setAttributeTableMode("records");
    }
    if (typeof window.setAttributeSelectedRowIds === "function") {
      window.setAttributeSelectedRowIds([id]);
    }
    if (typeof window.focusAttributeRowOnMap === "function") {
      window.focusAttributeRowOnMap(id);
    }
    return true;
  };

  const parseSpeedCatMapFromInputs = (inputMap, defaultMap) => {
    const map = {};
    ROUTE_SEQUENCER_SPEED_CAT_IDS.forEach(id => {
      const input = inputMap?.[id];
      const fallback = Number(defaultMap?.[id]);
      const value = clampRouteSequencerNumber(input?.value, 5, 85, fallback);
      map[id] = value;
      if (input) input.value = String(value);
    });
    return normalizeRouteSequencerSpeedCatMap(map, defaultMap);
  };

  const applySpeedCatMapToInputs = (inputMap, speedMap, defaultMap) => {
    const normalized = normalizeRouteSequencerSpeedCatMap(speedMap, defaultMap);
    ROUTE_SEQUENCER_SPEED_CAT_IDS.forEach(id => {
      const input = inputMap?.[id];
      if (!input) return;
      input.value = String(normalized[id]);
    });
  };

  const parseStreetClassAttrMapFromInputs = defaultMap => {
    const normalizedDefaultMap = normalizeRouteSequencerStreetClassAttrMap(defaultMap);
    const map = {};
    ROUTE_SEQUENCER_STREET_CLASS_IDS.forEach(id => {
      const inputGroup = classAttrInputs?.[id];
      const fallbackEntry = normalizedDefaultMap[id] || ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP[id];
      if (!inputGroup) {
        map[id] = normalizeRouteSequencerStreetClassAttrEntry(null, fallbackEntry);
        return;
      }
      const travelTypesAllowed = normalizeRouteSequencerTravelTypesAllowed(inputGroup.travelTypesAllowed?.value);
      const serviceSide = normalizeRouteSequencerServiceSideMode(inputGroup.serviceSide?.value);
      const crossPenaltyMinutes = Math.max(
        0,
        parseRouteSequencerMinutesValue(inputGroup.crossPenalty?.value, Number(fallbackEntry?.crossPenaltyMinutes) || 0)
      );
      const uTurnsRestricted = String(inputGroup.uTurnsRestricted?.value || "").trim().toLowerCase() === "yes";
      if (inputGroup.crossPenalty) {
        inputGroup.crossPenalty.value = formatRouteSequencerMinutesClock(crossPenaltyMinutes);
      }
      map[id] = normalizeRouteSequencerStreetClassAttrEntry({
        travelTypesAllowed,
        serviceSide,
        crossPenaltyMinutes,
        uTurnsRestricted
      }, fallbackEntry);
    });
    return normalizeRouteSequencerStreetClassAttrMap(map);
  };

  const applyStreetClassAttrMapToInputs = (inputMap, classAttrMap) => {
    const normalized = normalizeRouteSequencerStreetClassAttrMap(classAttrMap);
    ROUTE_SEQUENCER_STREET_CLASS_IDS.forEach(id => {
      const inputGroup = inputMap?.[id];
      if (!inputGroup) return;
      const entry = normalized[id] || ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP[id];
      if (inputGroup.travelTypesAllowed) inputGroup.travelTypesAllowed.value = normalizeRouteSequencerTravelTypesAllowed(entry?.travelTypesAllowed);
      if (inputGroup.serviceSide) inputGroup.serviceSide.value = normalizeRouteSequencerServiceSideMode(entry?.serviceSide);
      if (inputGroup.crossPenalty) {
        inputGroup.crossPenalty.value = formatRouteSequencerMinutesClock(
          Math.max(0, Number(entry?.crossPenaltyMinutes) || 0)
        );
      }
      if (inputGroup.uTurnsRestricted) inputGroup.uTurnsRestricted.value = entry?.uTurnsRestricted ? "yes" : "no";
    });
  };

  const fillSelect = (selectNode, options, includeNoneLabel = "") => {
    const select = selectNode;
    if (!select) return;
    const prev = String(select.value || "");
    select.innerHTML = "";
    if (includeNoneLabel) {
      const noneOption = document.createElement("option");
      noneOption.value = "";
      noneOption.textContent = includeNoneLabel;
      select.appendChild(noneOption);
    }
    (Array.isArray(options) ? options : []).forEach(item => {
      const option = document.createElement("option");
      option.value = String(item.value ?? "");
      option.textContent = String(item.label ?? item.value ?? "");
      select.appendChild(option);
    });
    if ([...select.options].some(option => option.value === prev)) {
      select.value = prev;
    }
  };

  const getAllRouteDayKeysForScope = () =>
    Object.keys(routeDayGroups).sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

  const buildRouteDayScopeLabel = routeDayKey => {
    const [routeToken = "", dayTokenRaw = ""] = String(routeDayKey || "").split("|");
    const dayToken = normalizeDayToken(dayTokenRaw) || String(dayTokenRaw || "").trim();
    const dayLabel = dayName(Number(dayToken)) || dayToken || "No Day";
    const group = routeDayGroups[routeDayKey];
    const stopCount = Array.isArray(group?.layers) ? group.layers.length : 0;
    return `Route ${routeToken || "Unassigned"} - ${dayLabel} (${stopCount.toLocaleString()})`;
  };

  const getSelectedRouteDayScopeKeysFromUi = () =>
    [...scopeListSelect.options]
      .filter(option => option.selected)
      .map(option => String(option.value || ""))
      .filter(key => Object.prototype.hasOwnProperty.call(routeDayGroups, key));

  const setSelectedRouteDayScopeKeys = keys => {
    const selected = new Set(
      (Array.isArray(keys) ? keys : [])
        .map(value => String(value || ""))
        .filter(key => Object.prototype.hasOwnProperty.call(routeDayGroups, key))
    );
    [...scopeListSelect.options].forEach(option => {
      option.selected = selected.has(String(option.value || ""));
    });
  };

  const populateRouteDayScopeSelector = settings => {
    const keys = getAllRouteDayKeysForScope();
    const selectedFromSettings = Array.isArray(settings?.routeDayScopeKeys)
      ? settings.routeDayScopeKeys.map(value => String(value || ""))
      : [];
    const selectedSet = new Set(selectedFromSettings.filter(key => keys.includes(key)));
    if (!selectedSet.size) keys.forEach(key => selectedSet.add(key));

    scopeListSelect.innerHTML = "";
    if (!keys.length) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "No Route+Day layers loaded.";
      option.disabled = true;
      option.selected = true;
      scopeListSelect.appendChild(option);
      return;
    }

    keys.forEach(key => {
      const option = document.createElement("option");
      option.value = key;
      option.textContent = buildRouteDayScopeLabel(key);
      option.selected = selectedSet.has(key);
      scopeListSelect.appendChild(option);
    });
  };

  const getRouteDayScopeKeysForRun = () => {
    const allKeys = getAllRouteDayKeysForScope();
    if (!allKeys.length) return [];
    const selectedKeys = getSelectedRouteDayScopeKeysFromUi();
    return selectedKeys.filter(key => allKeys.includes(key));
  };

  const readOutputFieldMappingFromUi = () => normalizeRouteSequencerOutputFields({
    sequence: sequenceOutputSelect.value,
    trip: tripOutputSelect.value,
    cumulativeLoad: cumulativeOutputSelect.value,
    driveMinutes: driveOutputSelect.value,
    serviceMinutes: serviceOutputSelect.value,
    totalMinutes: totalOutputSelect.value,
    dumpVisits: dumpVisitsOutputSelect.value
  });

  const validateOutputFieldMapping = outputFields => {
    const mapped = normalizeRouteSequencerOutputFields(outputFields);
    const values = Object.values(mapped).map(value => String(value || "").trim());
    if (values.some(value => !value)) {
      return { ok: false, message: "Choose all output fields in Output Field Mapping." };
    }
    const seen = new Set();
    for (const value of values) {
      const token = value.toLowerCase();
      if (seen.has(token)) {
        return { ok: false, message: `Each output metric must use a different field. "${value}" is duplicated.` };
      }
      seen.add(token);
    }
    return { ok: true, message: "" };
  };

  const updateFormatHint = (outputFieldsOverride = null) => {
    const outputFields = outputFieldsOverride || readOutputFieldMappingFromUi();
    formatHintNode.textContent = buildRouteSequencerFormatHintText(
      String(demandFieldSelect.value || ""),
      String(serviceFieldSelect.value || ""),
      outputFields,
      sequenceIncrementInput.value
    );
  };

  const populateOutputFieldSelectors = settings => {
    const fieldNames = getRouteSequencerFieldNames();
    const fieldSet = new Set(fieldNames);
    Object.values(ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS).forEach(name => fieldSet.add(name));
    const outputOptions = [...fieldSet]
      .filter(name => {
        const text = String(name || "").trim();
        if (!text) return false;
        return !isProtectedNewPlanningFieldName(text);
      })
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base", numeric: true }))
      .map(name => ({ value: name, label: name }));

    const mapped = normalizeRouteSequencerOutputFields(settings?.outputFields || ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS);
    Object.keys(outputFieldSelectMap).forEach(key => {
      const select = outputFieldSelectMap[key];
      if (!select) return;
      fillSelect(select, outputOptions);
      const selectedValue = String(mapped[key] || ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS[key] || "").trim();
      if ([...select.options].some(option => option.value === selectedValue)) {
        select.value = selectedValue;
      } else if ([...select.options].some(option => option.value === ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS[key])) {
        select.value = ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS[key];
      }
    });
    updateFormatHint(mapped);
  };

  const applySettingsToUi = () => {
    const settings = loadRouteSequencerSettings();
    const streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(settings.streetAttributeDefaults || settings);
    const meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(settings.meanderOverrideDefaults);
    capacityInput.value = String(
      clampRouteSequencerNumber(
        settings.capacityTons,
        0.1,
        1000,
        ROUTE_SEQUENCER_DEFAULT_CAPACITY_TONS
      )
    );
    defaultSpeedInput.value = String(
      clampRouteSequencerNumber(
        settings.fallbackSpeedMph,
        5,
        85,
        ROUTE_SEQUENCER_DEFAULT_SPEED_MPH
      )
    );
    sequenceIncrementInput.value = String(
      normalizeRouteSequencerSequenceIncrement(settings.sequenceIncrement, 1)
    );
    serviceStopsBeforeDeadheadToggle.checked = settings.serviceStopsBeforeDeadhead == null
      ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
      : settings.serviceStopsBeforeDeadhead !== false;
    dumpSourceSelect.value = ["both", "landfills", "custom", "none"].includes(String(settings.dumpSource))
      ? String(settings.dumpSource)
      : "both";
    dumpBeforeEndToggle.checked = settings.forceDumpBeforeEnd !== false;
    timeAtStartInput.value = formatRouteSequencerMinutesClock(
      parseRouteSequencerMinutesValue(settings.timeAtStartMinutes, 0)
    );
    timeAtEndInput.value = formatRouteSequencerMinutesClock(
      parseRouteSequencerMinutesValue(settings.timeAtEndMinutes, 0)
    );
    breaksInput.value = formatRouteSequencerBreakMinutesList(settings.breakMinutesList);
    speedClassFieldInput.value = normalizeRouteSequencerSpeedClassField(
      settings.speedClassField || ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD
    );
    applySpeedCatMapToInputs(
      drivingSpeedCatInputs,
      settings.drivingSpeedCatMap || settings.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
      ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
    );
    applySpeedCatMapToInputs(
      serviceSpeedCatInputs,
      settings.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
      ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
    );
    streetClassFieldInput.value = String(streetAttributeDefaults.streetClassField || "func_class");
    allowUTurnsOnlyInCuldesacsToggle.checked = !!streetAttributeDefaults.allowUTurnsOnlyInCuldesacs;
    discourageUTurnsToggle.checked = streetAttributeDefaults.discourageUTurnsOutsideIntersections !== false;
    meanderServiceInput.value = formatRouteSequencerMinutesClock(streetAttributeDefaults.meanderServiceMinutes);
    sameSideServiceInput.value = formatRouteSequencerMinutesClock(streetAttributeDefaults.sameSideServiceMinutes);
    meanderOverrideDrivingInput.value = normalizeRouteSequencerMeanderOverrideMode(meanderOverrideDefaults.driving);
    meanderOverrideWalkingInput.value = normalizeRouteSequencerMeanderOverrideMode(meanderOverrideDefaults.walking);
    uTurnProrateDrivingInput.value = normalizeRouteSequencerUTurnProrateMode(streetAttributeDefaults.uTurnProrateDriving);
    uTurnProrateWalkingInput.value = normalizeRouteSequencerUTurnProrateMode(streetAttributeDefaults.uTurnProrateWalking);
    applyStreetClassAttrMapToInputs(classAttrInputs, streetAttributeDefaults.classAttrMap);
    return settings;
  };

  const populateFieldSelectors = settings => {
    const fieldNames = getRouteSequencerFieldNames();
    const demandOptions = fieldNames.map(name => ({ value: name, label: name }));
    const serviceOptions = fieldNames.map(name => ({ value: name, label: name }));
    fillSelect(demandFieldSelect, demandOptions, "Choose demand field...");
    fillSelect(serviceFieldSelect, serviceOptions, "None (0 minutes)");
    const preferredServiceField = fieldNames.find(name => normalizeRouteSequencerFieldToken(name) === "servicetime");

    const guessedDemand = guessRouteSequencerFieldByAliases(fieldNames, [
      "demand", "demand_a", "tons", "tonnage", "load", "totaldemand", "volume"
    ]);
    const guessedService = guessRouteSequencerFieldByAliases(fieldNames, [
      "duration_a",
      "durationa",
      "duration",
      "servicetime",
      "service_time",
      "servmin",
      "servicemin",
      "minutes",
      "minutesperstop",
      "timeonsite"
    ]);

    const demandValue = String(settings?.demandField || "");
    if (demandValue && [...demandFieldSelect.options].some(option => option.value === demandValue)) {
      demandFieldSelect.value = demandValue;
    } else if (guessedDemand) {
      demandFieldSelect.value = guessedDemand;
    }

    const serviceValue = String(settings?.serviceField || "");
    if (serviceValue && [...serviceFieldSelect.options].some(option => option.value === serviceValue)) {
      serviceFieldSelect.value = serviceValue;
    } else if (preferredServiceField) {
      serviceFieldSelect.value = preferredServiceField;
    } else if (guessedService) {
      serviceFieldSelect.value = guessedService;
    } else {
      serviceFieldSelect.value = "";
    }
  };

  const populateDepotSelectors = settings => {
    const depots = getRouteSequencerDepots();
    const options = depots.map(item => ({
      value: String(item.id),
      label: `${item.name} (${Number(item.lat).toFixed(5)}, ${Number(item.lon).toFixed(5)})`
    }));
    fillSelect(startDepotSelect, options, "None (use first stop)");
    fillSelect(endDepotSelect, options, "None (use last stop)");
    if (settings?.startDepotId && [...startDepotSelect.options].some(option => option.value === String(settings.startDepotId))) {
      startDepotSelect.value = String(settings.startDepotId);
    }
    if (settings?.endDepotId && [...endDepotSelect.options].some(option => option.value === String(settings.endDepotId))) {
      endDepotSelect.value = String(settings.endDepotId);
    }
  };

  const populateCustomDumpSelector = settings => {
    const customDumps = getRouteSequencerCustomDumps();
    const options = customDumps.map(item => ({
      value: String(item.id),
      label: `${item.name} (${Number(item.lat).toFixed(5)}, ${Number(item.lon).toFixed(5)}) - time ${formatRouteSequencerMinutesClock(item.facilityMinutes)}`
    }));
    fillSelect(customDumpSelect, options, "Select custom dump...");
    if (settings?.selectedDumpId && [...customDumpSelect.options].some(option => option.value === String(settings.selectedDumpId))) {
      customDumpSelect.value = String(settings.selectedDumpId);
    }
  };

  const renderModalUi = () => {
    const settings = applySettingsToUi();
    populateFieldSelectors(settings);
    populateOutputFieldSelectors(settings);
    populateRouteDayScopeSelector(settings);
    populateDepotSelectors(settings);
    populateCustomDumpSelector(settings);
    renderRouteSequencerResultsList(resultsNode, []);
    setStatus("Ready.");
  };

  window.__refreshRouteSequencerFacilitiesUi = () => {
    if (modal.style.display !== "flex") return;
    const settings = loadRouteSequencerSettings();
    populateDepotSelectors(settings);
    populateCustomDumpSelector(settings);
  };

  const saveCurrentUiSettings = () => {
    const outputFields = readOutputFieldMappingFromUi();
    const sequenceIncrement = normalizeRouteSequencerSequenceIncrement(sequenceIncrementInput.value, 1);
    const timeAtStartMinutes = parseRouteSequencerMinutesValue(timeAtStartInput.value, 0);
    const timeAtEndMinutes = parseRouteSequencerMinutesValue(timeAtEndInput.value, 0);
    const breakMinutesList = parseRouteSequencerBreakMinutesList(breaksInput.value);
    const speedClassField = normalizeRouteSequencerSpeedClassField(speedClassFieldInput.value);
    speedClassFieldInput.value = speedClassField;
    const streetClassField = String(streetClassFieldInput.value || "").trim() || "func_class";
    streetClassFieldInput.value = streetClassField;
    const drivingSpeedCatMap = parseSpeedCatMapFromInputs(
      drivingSpeedCatInputs,
      ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
    );
    const serviceSpeedCatMap = parseSpeedCatMapFromInputs(
      serviceSpeedCatInputs,
      ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
    );
    const meanderServiceMinutes = Math.max(0, parseRouteSequencerMinutesValue(
      meanderServiceInput.value,
      ROUTE_SEQUENCER_DEFAULT_STREET_ATTRIBUTE_DEFAULTS.meanderServiceMinutes
    ));
    const sameSideServiceMinutes = Math.max(0, parseRouteSequencerMinutesValue(
      sameSideServiceInput.value,
      ROUTE_SEQUENCER_DEFAULT_STREET_ATTRIBUTE_DEFAULTS.sameSideServiceMinutes
    ));
    const meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults({
      driving: meanderOverrideDrivingInput.value,
      walking: meanderOverrideWalkingInput.value
    });
    meanderServiceInput.value = formatRouteSequencerMinutesClock(meanderServiceMinutes);
    sameSideServiceInput.value = formatRouteSequencerMinutesClock(sameSideServiceMinutes);
    meanderOverrideDrivingInput.value = meanderOverrideDefaults.driving;
    meanderOverrideWalkingInput.value = meanderOverrideDefaults.walking;
    const streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults({
      streetClassField,
      allowUTurnsOnlyInCuldesacs: !!allowUTurnsOnlyInCuldesacsToggle.checked,
      discourageUTurnsOutsideIntersections: !!discourageUTurnsToggle.checked,
      meanderServiceMinutes,
      sameSideServiceMinutes,
      uTurnProrateDriving: normalizeRouteSequencerUTurnProrateMode(uTurnProrateDrivingInput.value),
      uTurnProrateWalking: normalizeRouteSequencerUTurnProrateMode(uTurnProrateWalkingInput.value),
      classAttrMap: parseStreetClassAttrMapFromInputs(ROUTE_SEQUENCER_DEFAULT_STREET_CLASS_ATTR_MAP)
    });
    uTurnProrateDrivingInput.value = streetAttributeDefaults.uTurnProrateDriving;
    uTurnProrateWalkingInput.value = streetAttributeDefaults.uTurnProrateWalking;
    sequenceIncrementInput.value = String(sequenceIncrement);
    timeAtStartInput.value = formatRouteSequencerMinutesClock(timeAtStartMinutes);
    timeAtEndInput.value = formatRouteSequencerMinutesClock(timeAtEndMinutes);
    breaksInput.value = formatRouteSequencerBreakMinutesList(breakMinutesList);
    const payload = {
      demandField: String(demandFieldSelect.value || ""),
      serviceField: String(serviceFieldSelect.value || ""),
      capacityTons: clampRouteSequencerNumber(capacityInput.value, 0.1, 1000, ROUTE_SEQUENCER_DEFAULT_CAPACITY_TONS),
      fallbackSpeedMph: clampRouteSequencerNumber(defaultSpeedInput.value, 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH),
      sequenceIncrement,
      routeDayScopeKeys: getSelectedRouteDayScopeKeysFromUi(),
      dumpSource: ["both", "landfills", "custom", "none"].includes(String(dumpSourceSelect.value))
        ? String(dumpSourceSelect.value)
        : "both",
      startDepotId: String(startDepotSelect.value || ""),
      endDepotId: String(endDepotSelect.value || ""),
      timeAtStartMinutes,
      timeAtEndMinutes,
      breakMinutesList,
      forceDumpBeforeEnd: !!dumpBeforeEndToggle.checked,
      serviceStopsBeforeDeadhead: !!serviceStopsBeforeDeadheadToggle.checked,
      speedClassField,
      drivingSpeedCatMap,
      serviceSpeedCatMap,
      classSpeedMap: drivingSpeedCatMap,
      meanderOverrideDefaults,
      streetAttributeDefaults,
      selectedDumpId: String(customDumpSelect.value || ""),
      outputFields
    };
    saveRouteSequencerSettings(payload);
    updateFormatHint(outputFields);
    return payload;
  };

  const openModal = () => {
    renderModalUi();
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const openModalWithBackendNotice = async () => {
    if (routeSequencerOpenCheckRunning) return;
    routeSequencerOpenCheckRunning = true;
    try {
    let backendRunning = false;
    try {
      backendRunning = await checkLocalStreetBackendAvailability(true);
    } catch {
      backendRunning = false;
    }
    if (!backendRunning) {
      alert(
        "Street Network backend is not running.\n\n" +
        "Open Street Network Manager to start the backend before sequencing."
      );
    }
    openModal();
    } finally {
      routeSequencerOpenCheckRunning = false;
    }
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  const addDepotFromCenter = () => {
    const center = map?.getCenter?.();
    if (!center) {
      setStatus("Map center is unavailable.", "error");
      return;
    }
    const name = normalizeRouteSequencerName(depotNameInput.value, "Depot");
    const depots = getRouteSequencerDepots();
    depots.push({
      id: `depot_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      name,
      lat: Number(center.lat),
      lon: Number(center.lng)
    });
    setRouteSequencerDepots(depots);
    depotNameInput.value = "";
    const settings = saveCurrentUiSettings();
    populateDepotSelectors(settings);
    setStatus(`Depot "${name}" added from map center.`, "success");
  };

  const removeSelectedDepot = () => {
    const selectedId = String(startDepotSelect.value || endDepotSelect.value || "");
    if (!selectedId) {
      setStatus("Choose a start or end depot to remove.", "error");
      return;
    }
    const depots = getRouteSequencerDepots();
    const target = depots.find(item => String(item.id) === selectedId);
    if (!target) {
      setStatus("Selected depot was not found.", "error");
      return;
    }
    const next = depots.filter(item => String(item.id) !== selectedId);
    setRouteSequencerDepots(next);
    startDepotSelect.value = "";
    endDepotSelect.value = "";
    const settings = saveCurrentUiSettings();
    populateDepotSelectors(settings);
    setStatus(`Removed depot "${target.name}".`, "success");
  };

  const addDumpFromCenter = () => {
    const center = map?.getCenter?.();
    if (!center) {
      setStatus("Map center is unavailable.", "error");
      return;
    }
    const name = normalizeRouteSequencerName(dumpNameInput.value, "Custom Dump");
    const dumps = getRouteSequencerCustomDumps();
    const id = `dump_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    dumps.push({
      id,
      name,
      lat: Number(center.lat),
      lon: Number(center.lng)
    });
    setRouteSequencerCustomDumps(dumps);
    dumpNameInput.value = "";
    const settings = saveCurrentUiSettings();
    settings.selectedDumpId = id;
    populateCustomDumpSelector(settings);
    customDumpSelect.value = id;
    setStatus(`Custom dump "${name}" added from map center.`, "success");
  };

  const removeSelectedDump = () => {
    const selectedId = String(customDumpSelect.value || "");
    if (!selectedId) {
      setStatus("Select a custom dump to remove.", "error");
      return;
    }
    const dumps = getRouteSequencerCustomDumps();
    const target = dumps.find(item => String(item.id) === selectedId);
    if (!target) {
      setStatus("Selected custom dump was not found.", "error");
      return;
    }
    const next = dumps.filter(item => String(item.id) !== selectedId);
    setRouteSequencerCustomDumps(next);
    customDumpSelect.value = "";
    const settings = saveCurrentUiSettings();
    settings.selectedDumpId = "";
    populateCustomDumpSelector(settings);
    setStatus(`Removed custom dump "${target.name}".`, "success");
  };

  const runEvaluateSequence = async () => {
    const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
    if (!rows.length || !Object.keys(routeDayGroups).length) {
      setStatus("Load a route file before evaluating sequence.", "error");
      return;
    }

    const settings = saveCurrentUiSettings();
    const capacityTons = clampRouteSequencerNumber(settings.capacityTons, 0.1, 1000, ROUTE_SEQUENCER_DEFAULT_CAPACITY_TONS);
    const fallbackSpeedMph = clampRouteSequencerNumber(settings.fallbackSpeedMph, 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH);
    const speedClassField = normalizeRouteSequencerSpeedClassField(
      settings.speedClassField || ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD
    );
    const drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
      settings.drivingSpeedCatMap || settings.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
      ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
    );
    const serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
      settings.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
      ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
    );
    const fallbackServiceSpeedMph = clampRouteSequencerNumber(
      settings.fallbackServiceSpeedMph,
      5,
      85,
      Math.min(fallbackSpeedMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
    );
    const demandField = String(settings.demandField || "");
    const serviceField = String(settings.serviceField || "");
    const breakMinutesList = parseRouteSequencerBreakMinutesList(settings.breakMinutesList);
    const serviceStopsBeforeDeadhead = settings.serviceStopsBeforeDeadhead == null
      ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
      : settings.serviceStopsBeforeDeadhead !== false;
    const meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(settings.meanderOverrideDefaults);
    const streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(settings.streetAttributeDefaults);
    const dumpSource = String(settings.dumpSource || "both");
    const outputFields = normalizeRouteSequencerOutputFields(settings.outputFields || ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS);
    const sequenceField = String(outputFields.sequence || "").trim();
    const requireNetworkOnly = ROUTE_SEQUENCER_REQUIRE_NETWORK_ONLY;

    if (!sequenceField) {
      setStatus("Choose a Sequence Output field before Evaluate Sequence.", "error");
      return;
    }

    setStatus(`Building travel graph from loaded street segments (speed class field: ${speedClassField})...`);
    await yieldRouteSequencerUi();
    const graph = streetAttributeById.size
      ? buildRouteSequencerStreetGraph({
        speedClassField,
        drivingSpeedCatMap,
        serviceSpeedCatMap,
        fallbackServiceSpeedMph
      }, fallbackSpeedMph)
      : null;
    const hasGraph = !!(graph && Number(graph.nodeCount) > 0 && Number(graph.edgeCount) > 0);
    if (requireNetworkOnly && !hasGraph) {
      setStatus(
        "Street-network-only sequencing is enabled. Start the Street backend and load street segments before Evaluate Sequence.",
        "error"
      );
      renderRouteSequencerResultsList(resultsNode, []);
      clearSequenceLayerData();
      return;
    }
    const graphSummary = graph
      ? `${graph.nodeCount.toLocaleString()} nodes / ${graph.edgeCount.toLocaleString()} edges`
      : "No street graph";
    setStatus(`Travel graph ready: ${graphSummary}. Evaluating existing sequence values...`);

    let landfillPoints = [];
    if (dumpSource === "both" || dumpSource === "landfills") {
      try {
        await ensureTexasLandfillsLoaded();
      } catch (error) {
        console.warn("Unable to load Texas landfills for sequence evaluation:", error);
      }
      landfillPoints = collectRouteSequencerLandfillDumpPoints(graph);
    }
    const customDumpPoints = (dumpSource === "both" || dumpSource === "custom")
      ? collectRouteSequencerCustomDumpPoints(graph)
      : [];
    const dumpPoints = dumpSource === "none"
      ? []
      : [...customDumpPoints, ...landfillPoints];

    const startDepot = getRouteSequencerDepotPointById(settings.startDepotId, graph);
    const endDepot = getRouteSequencerDepotPointById(settings.endDepotId, graph);

    const plans = [];
    const routeDayKeys = getRouteDayScopeKeysForRun();
    if (!routeDayKeys.length) {
      setStatus("Select one or more Route+Day units in Route + Day Scope before Evaluate Sequence.", "error");
      return;
    }
    let skippedRouteDayCount = 0;
    let missingSequenceStops = 0;
    let groupsWithoutAnySequence = 0;
    let networkBlockedRouteDayCount = 0;

    for (let i = 0; i < routeDayKeys.length; i += 1) {
      const key = routeDayKeys[i];
      const group = routeDayGroups[key];
      const built = buildRouteSequencerStopPointsFromExistingSequence(
        key,
        group,
        demandField,
        serviceField,
        sequenceField,
        graph,
        streetAttributeDefaults
      );
      const stops = Array.isArray(built?.stops) ? built.stops : [];
      if (!stops.length) {
        skippedRouteDayCount += 1;
        continue;
      }

      if (!(Number(built.validSequenceCount) > 0)) {
        groupsWithoutAnySequence += 1;
        continue;
      }

      missingSequenceStops += Number(built.missingSequenceCount) || 0;
      const plan = solveRouteDaySequenceEvaluationPlan(key, stops, {
        graph,
        dumpPoints,
        capacityTons,
        fallbackSpeedMph,
        fallbackServiceSpeedMph,
        startDepot,
        endDepot,
        timeAtStartMinutes: settings.timeAtStartMinutes,
        timeAtEndMinutes: settings.timeAtEndMinutes,
        breakMinutesList,
        forceDumpBeforeEnd: !!settings.forceDumpBeforeEnd,
        serviceStopsBeforeDeadhead,
        requireNetworkOnly,
        meanderOverrideDefaults,
        streetAttributeDefaults
      });
      if (plan) plans.push(plan);
      else if (requireNetworkOnly) networkBlockedRouteDayCount += 1;

      setStatus(`Evaluating Route+Day unit ${i + 1} of ${routeDayKeys.length} from "${sequenceField}"...`);
      if (i % 4 === 0) {
        await new Promise(resolve => setTimeout(resolve, 0));
      }
    }

    if (!plans.length) {
      if (networkBlockedRouteDayCount > 0) {
        setStatus(
          `Evaluate Sequence could not run. ${networkBlockedRouteDayCount.toLocaleString()} Route+Day group(s) could not be connected on the loaded street network.`,
          "error"
        );
      } else {
        setStatus(
          `Evaluate Sequence could not run. No Route+Day group had usable numeric sequence values in "${sequenceField}".`,
          "error"
        );
      }
      renderRouteSequencerResultsList(resultsNode, []);
      clearSequenceLayerData();
      return;
    }

    const updatedRows = applyRouteSequencerPlansToRows(plans, outputFields, { sequenceIncrement: 1 });
    rebuildSequenceLayerFromPlans(plans);
    syncSequenceLayerVisibilityOnMap();
    refreshAllRecordPopupContent();
    applyAttributeSelectionStyles();
    renderAttributeTable();
    refreshAttributeStatus();

    const totalStops = plans.reduce((sum, plan) => sum + (Number(plan.stopVisits?.length) || 0), 0);
    const totalDumpVisits = plans.reduce((sum, plan) => sum + (Number(plan.dumpVisits) || 0), 0);
    const totalDriveMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalDriveMinutes) || 0), 0);
    const totalDriveMiles = plans.reduce((sum, plan) => sum + (Number(plan.totalDriveMiles) || 0), 0);
    const totalServiceMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalServiceMinutes) || 0), 0);
    const totalFacilityMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalFacilityMinutes) || 0), 0);
    const totalBreakMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalBreakMinutes) || 0), 0);
    const totalRouteMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalMinutes) || 0), 0);

    renderRouteSequencerResultsList(resultsNode, plans);
    const skipNote = skippedRouteDayCount
      ? ` ${skippedRouteDayCount.toLocaleString()} route/day groups had no mappable stops.`
      : "";
    const noSeqNote = groupsWithoutAnySequence
      ? ` ${groupsWithoutAnySequence.toLocaleString()} route/day groups had no valid sequence values in "${sequenceField}".`
      : "";
    const partialNote = missingSequenceStops
      ? ` ${missingSequenceStops.toLocaleString()} stop(s) were missing sequence values and were placed after sequenced stops.`
      : "";
    const networkBlockedNote = networkBlockedRouteDayCount
      ? ` ${networkBlockedRouteDayCount.toLocaleString()} route/day group(s) could not be connected on the street network and were skipped.`
      : "";
    setStatus(
      `Evaluate Sequence complete from "${sequenceField}": ${plans.length.toLocaleString()} of ${routeDayKeys.length.toLocaleString()} selected Route+Day groups, ${totalStops.toLocaleString()} stops, ${totalDumpVisits.toLocaleString()} dump visits, drive ${totalDriveMinutes.toFixed(1)} min, miles ${totalDriveMiles.toFixed(2)} mi, service ${totalServiceMinutes.toFixed(1)} min, facility ${totalFacilityMinutes.toFixed(1)} min, breaks ${totalBreakMinutes.toFixed(1)} min, total ${totalRouteMinutes.toFixed(1)} min.${skipNote}${noSeqNote}${partialNote}${networkBlockedNote} Updated rows: ${updatedRows.toLocaleString()}.`,
      "success"
    );
  };

  const runSequencer = async () => {
    const rows = Array.isArray(window._currentRows) ? window._currentRows : [];
    if (!rows.length || !Object.keys(routeDayGroups).length) {
      setStatus("Load a route file before running Auto Sequence.", "error");
      return;
    }

    const settings = saveCurrentUiSettings();
    const capacityTons = clampRouteSequencerNumber(settings.capacityTons, 0.1, 1000, ROUTE_SEQUENCER_DEFAULT_CAPACITY_TONS);
    const fallbackSpeedMph = clampRouteSequencerNumber(settings.fallbackSpeedMph, 5, 85, ROUTE_SEQUENCER_DEFAULT_SPEED_MPH);
    const sequenceIncrement = normalizeRouteSequencerSequenceIncrement(settings.sequenceIncrement, 1);
    const speedClassField = normalizeRouteSequencerSpeedClassField(
      settings.speedClassField || ROUTE_SEQUENCER_DEFAULT_SPEED_CLASS_FIELD
    );
    const drivingSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
      settings.drivingSpeedCatMap || settings.classSpeedMap || ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP,
      ROUTE_SEQUENCER_DEFAULT_DRIVING_SPEED_CAT_MAP
    );
    const serviceSpeedCatMap = normalizeRouteSequencerSpeedCatMap(
      settings.serviceSpeedCatMap || ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP,
      ROUTE_SEQUENCER_DEFAULT_SERVICE_SPEED_CAT_MAP
    );
    const fallbackServiceSpeedMph = clampRouteSequencerNumber(
      settings.fallbackServiceSpeedMph,
      5,
      85,
      Math.min(fallbackSpeedMph, ROUTE_SEQUENCER_DEFAULT_SERVICE_TRAVEL_SPEED_MPH)
    );
    const demandField = String(settings.demandField || "");
    const serviceField = String(settings.serviceField || "");
    const breakMinutesList = parseRouteSequencerBreakMinutesList(settings.breakMinutesList);
    const serviceStopsBeforeDeadhead = settings.serviceStopsBeforeDeadhead == null
      ? ROUTE_SEQUENCER_DEFAULT_SERVICE_STOPS_BEFORE_DEADHEAD
      : settings.serviceStopsBeforeDeadhead !== false;
    const meanderOverrideDefaults = normalizeRouteSequencerMeanderOverrideDefaults(settings.meanderOverrideDefaults);
    const streetAttributeDefaults = normalizeRouteSequencerStreetAttributeDefaults(settings.streetAttributeDefaults);
    const dumpSource = String(settings.dumpSource || "both");
    const outputFields = normalizeRouteSequencerOutputFields(settings.outputFields || ROUTE_SEQUENCER_DEFAULT_OUTPUT_FIELDS);
    const requireNetworkOnly = ROUTE_SEQUENCER_REQUIRE_NETWORK_ONLY;

    if (!demandField) {
      setStatus("Choose a Demand field (Tons) before running Auto Sequence.", "error");
      return;
    }
    const outputCheck = validateOutputFieldMapping(outputFields);
    if (!outputCheck.ok) {
      setStatus(outputCheck.message, "error");
      return;
    }

    setStatus(`Building travel graph from loaded street segments (speed class field: ${speedClassField})...`);
    await yieldRouteSequencerUi();
    const graph = streetAttributeById.size
      ? buildRouteSequencerStreetGraph({
        speedClassField,
        drivingSpeedCatMap,
        serviceSpeedCatMap,
        fallbackServiceSpeedMph
      }, fallbackSpeedMph)
      : null;
    const hasGraph = !!(graph && Number(graph.nodeCount) > 0 && Number(graph.edgeCount) > 0);
    if (requireNetworkOnly && !hasGraph) {
      setStatus(
        "Street-network-only sequencing is enabled. Start the Street backend and load street segments before Auto Sequence.",
        "error"
      );
      renderRouteSequencerResultsList(resultsNode, []);
      clearSequenceLayerData();
      return;
    }
    const graphSummary = graph
      ? `${graph.nodeCount.toLocaleString()} nodes / ${graph.edgeCount.toLocaleString()} edges`
      : "No street graph";
    setStatus(`Travel graph ready: ${graphSummary}. Gathering facilities...`);

    let landfillPoints = [];
    if (dumpSource === "both" || dumpSource === "landfills") {
      try {
        await ensureTexasLandfillsLoaded();
      } catch (error) {
        console.warn("Unable to load Texas landfills for sequencer:", error);
      }
      landfillPoints = collectRouteSequencerLandfillDumpPoints(graph);
    }
    const customDumpPoints = (dumpSource === "both" || dumpSource === "custom")
      ? collectRouteSequencerCustomDumpPoints(graph)
      : [];
    const dumpPoints = dumpSource === "none"
      ? []
      : [...customDumpPoints, ...landfillPoints];

    const startDepot = getRouteSequencerDepotPointById(settings.startDepotId, graph);
    const endDepot = getRouteSequencerDepotPointById(settings.endDepotId, graph);

    const plans = [];
    const routeDayKeys = getRouteDayScopeKeysForRun();
    if (!routeDayKeys.length) {
      setStatus("Select one or more Route+Day units in Route + Day Scope before Auto Sequence.", "error");
      return;
    }
    let skippedRouteDayCount = 0;
    let networkBlockedRouteDayCount = 0;

    for (let i = 0; i < routeDayKeys.length; i += 1) {
      const key = routeDayKeys[i];
      const group = routeDayGroups[key];
      const stops = buildRouteSequencerStopPoints(key, group, demandField, serviceField, graph, streetAttributeDefaults);
      if (!stops.length) {
        skippedRouteDayCount += 1;
        continue;
      }
      const plan = solveRouteDaySequencingPlan(key, stops, {
        graph,
        dumpPoints,
        startDepot,
        endDepot,
        timeAtStartMinutes: settings.timeAtStartMinutes,
        timeAtEndMinutes: settings.timeAtEndMinutes,
        breakMinutesList,
        capacityTons,
        fallbackSpeedMph,
        fallbackServiceSpeedMph,
        forceDumpBeforeEnd: !!settings.forceDumpBeforeEnd,
        serviceStopsBeforeDeadhead,
        requireNetworkOnly,
        meanderOverrideDefaults,
        streetAttributeDefaults
      });
      if (plan) plans.push(plan);
      else if (requireNetworkOnly) networkBlockedRouteDayCount += 1;
      setStatus(`Sequencing Route+Day unit ${i + 1} of ${routeDayKeys.length}...`);
      if (i % 4 === 0) {
        await new Promise(resolve => setTimeout(resolve, 0));
      }
    }

    if (!plans.length) {
      if (networkBlockedRouteDayCount > 0) {
        setStatus(
          `Auto Sequence could not run. ${networkBlockedRouteDayCount.toLocaleString()} Route+Day group(s) could not be connected on the loaded street network.`,
          "error"
        );
      } else {
        setStatus("No route/day stops were available to sequence.", "error");
      }
      renderRouteSequencerResultsList(resultsNode, []);
      clearSequenceLayerData();
      return;
    }

    const updatedRows = applyRouteSequencerPlansToRows(plans, outputFields, { sequenceIncrement });
    rebuildSequenceLayerFromPlans(plans);
    syncSequenceLayerVisibilityOnMap();
    refreshAllRecordPopupContent();
    applyAttributeSelectionStyles();
    renderAttributeTable();
    refreshAttributeStatus();

    const totalStops = plans.reduce((sum, plan) => sum + (Number(plan.stopVisits?.length) || 0), 0);
    const totalDumpVisits = plans.reduce((sum, plan) => sum + (Number(plan.dumpVisits) || 0), 0);
    const totalDriveMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalDriveMinutes) || 0), 0);
    const totalDriveMiles = plans.reduce((sum, plan) => sum + (Number(plan.totalDriveMiles) || 0), 0);
    const totalServiceMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalServiceMinutes) || 0), 0);
    const totalFacilityMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalFacilityMinutes) || 0), 0);
    const totalBreakMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalBreakMinutes) || 0), 0);
    const totalRouteMinutes = plans.reduce((sum, plan) => sum + (Number(plan.totalMinutes) || 0), 0);

    renderRouteSequencerResultsList(resultsNode, plans);
    const skipNote = skippedRouteDayCount
      ? ` ${skippedRouteDayCount.toLocaleString()} route/day groups had no mappable stops.`
      : "";
    const networkBlockedNote = networkBlockedRouteDayCount
      ? ` ${networkBlockedRouteDayCount.toLocaleString()} route/day group(s) could not be connected on the street network and were skipped.`
      : "";
    setStatus(
      `Auto Sequence complete: ${(plans.length).toLocaleString()} of ${routeDayKeys.length.toLocaleString()} selected Route+Day groups (each treated as one route), ${totalStops.toLocaleString()} stops, ${totalDumpVisits.toLocaleString()} dump visits, drive ${totalDriveMinutes.toFixed(1)} min, miles ${totalDriveMiles.toFixed(2)} mi, service ${totalServiceMinutes.toFixed(1)} min, facility ${totalFacilityMinutes.toFixed(1)} min, breaks ${totalBreakMinutes.toFixed(1)} min, total ${totalRouteMinutes.toFixed(1)} min, sequence increment ${sequenceIncrement}.${skipNote}${networkBlockedNote} Updated rows: ${updatedRows.toLocaleString()}.`,
      "success"
    );
  };

  openBtn.addEventListener("click", () => {
    openModalWithBackendNotice().catch(error => {
      console.warn("Unable to verify Street Network backend before opening Sequence:", error);
    });
  });
  if (openBtnMobile) {
    openBtnMobile.addEventListener("click", () => {
      openModalWithBackendNotice().catch(error => {
        console.warn("Unable to verify Street Network backend before opening Sequence (mobile):", error);
      });
    });
  }
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });

  addDepotBtn.addEventListener("click", addDepotFromCenter);
  removeDepotBtn.addEventListener("click", removeSelectedDepot);
  addDumpBtn.addEventListener("click", addDumpFromCenter);
  removeDumpBtn.addEventListener("click", removeSelectedDump);
  scopeAllBtn.addEventListener("click", () => {
    setSelectedRouteDayScopeKeys(getAllRouteDayKeysForScope());
    saveCurrentUiSettings();
    const count = getRouteDayScopeKeysForRun().length;
    setStatus(`Scope set to all Route+Day units (${count.toLocaleString()}).`);
  });
  scopeNoneBtn.addEventListener("click", () => {
    setSelectedRouteDayScopeKeys([]);
    saveCurrentUiSettings();
    setStatus("Scope cleared. Select one or more Route+Day units before run.");
  });
  scopeVisibleBtn.addEventListener("click", () => {
    const visibleKeys = getAllRouteDayKeysForScope().filter(key => isRouteDayLayerVisibleOnMap(key));
    setSelectedRouteDayScopeKeys(visibleKeys);
    saveCurrentUiSettings();
    setStatus(
      visibleKeys.length
        ? `Scope set to visible Route+Day units (${visibleKeys.length.toLocaleString()}).`
        : "No visible Route+Day units found."
    );
  });
  scopeListSelect.addEventListener("change", () => {
    saveCurrentUiSettings();
    const count = getRouteDayScopeKeysForRun().length;
    setStatus(
      count
        ? `Scope: ${count.toLocaleString()} Route+Day unit(s) selected.`
        : "Scope cleared. Select one or more Route+Day units before run."
    );
  });

  resultsNode.addEventListener("click", event => {
    const clickableRow = event.target?.closest?.('tr[data-rs-clickable="1"]');
    if (!clickableRow || !resultsNode.contains(clickableRow)) return;

    const eventType = String(clickableRow.dataset.rsEventType || "");
    if (!eventType) return;
    setActiveSequenceResultRow(clickableRow);

    if (eventType === "stop") {
      const rowId = Number(clickableRow.dataset.rsRowId);
      selectSequencerStopFromResult(rowId);
      return;
    }

    if (eventType === "dump" || eventType === "startDepot" || eventType === "endDepot") {
      const lat = Number(clickableRow.dataset.rsLat);
      const lon = Number(clickableRow.dataset.rsLon);
      focusSequencerFacilityFromResult(lat, lon);
    }
  });

  [
    demandFieldSelect,
    serviceFieldSelect,
    sequenceOutputSelect,
    tripOutputSelect,
    cumulativeOutputSelect,
    driveOutputSelect,
    serviceOutputSelect,
    totalOutputSelect,
    dumpVisitsOutputSelect,
    capacityInput,
    defaultSpeedInput,
    sequenceIncrementInput,
    streetClassFieldInput,
    serviceStopsBeforeDeadheadToggle,
    allowUTurnsOnlyInCuldesacsToggle,
    discourageUTurnsToggle,
    meanderServiceInput,
    sameSideServiceInput,
    meanderOverrideDrivingInput,
    meanderOverrideWalkingInput,
    uTurnProrateDrivingInput,
    uTurnProrateWalkingInput,
    dumpSourceSelect,
    startDepotSelect,
    endDepotSelect,
    timeAtStartInput,
    timeAtEndInput,
    breaksInput,
    dumpBeforeEndToggle,
    customDumpSelect
  ]
    .forEach(node => {
      node?.addEventListener("change", () => {
        saveCurrentUiSettings();
      });
    });
  [
    speedClassFieldInput,
    streetClassFieldInput,
    ...Object.values(drivingSpeedCatInputs),
    ...Object.values(serviceSpeedCatInputs),
    ...Object.values(classAttrInputs).flatMap(group => [
      group.travelTypesAllowed,
      group.serviceSide,
      group.crossPenalty,
      group.uTurnsRestricted
    ])
  ].forEach(input => {
    input?.addEventListener("change", () => saveCurrentUiSettings());
  });

  evaluateBtn.addEventListener("click", () => {
    if (routeSequencerTaskRunning) return;
    setRouteSequencerLoading(true, "Preparing Evaluate Sequence...");
    runEvaluateSequence().catch(error => {
      console.error("Route sequence evaluation failed:", error);
      setStatus(`Evaluate Sequence failed: ${error?.message || error}`, "error");
    }).finally(() => {
      setRouteSequencerLoading(false);
    });
  });

  runBtn.addEventListener("click", () => {
    if (routeSequencerTaskRunning) return;
    setRouteSequencerLoading(true, "Preparing Auto Sequence...");
    runSequencer().catch(error => {
      console.error("Route sequencer failed:", error);
      setStatus(`Auto Sequence failed: ${error?.message || error}`, "error");
    }).finally(() => {
      setRouteSequencerLoading(false);
    });
  });
}

const SET_SEGMENT_ID_SIDE_SETTINGS_KEY = "setSegmentIdSideSettings";
const SET_SEGMENT_MATCH_PROFILES_KEY = "setSegmentMatchProfiles";
const SET_SEGMENT_MATCH_STATUS_FIELD_FALLBACK = "Segment_Match_Status";
const SET_SEGMENT_SEQUENCE_READY_FIELD_FALLBACK = "Sequence_Ready";
const SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK = "Segment_Match_Method";
const SET_SEGMENT_PROFILE_MAX_COUNT = 48;
const SET_SEGMENT_FEET_PER_DEG_LAT = 364000;
const SET_SEGMENT_MATCH_CELL_DEG = 0.01;
const SET_SEGMENT_ZERO_SIDE_FEET = 2.5;
const SET_SEGMENT_DIRECTION_TOKENS = new Set([
  "N", "S", "E", "W", "NE", "NW", "SE", "SW",
  "NORTH", "SOUTH", "EAST", "WEST"
]);
const SET_SEGMENT_SUFFIX_TOKENS = new Set([
  "ST", "STREET", "RD", "ROAD", "AVE", "AV", "AVENUE", "BLVD", "BOULEVARD",
  "DR", "DRIVE", "LN", "LANE", "CT", "COURT", "PL", "PLACE", "PKWY", "PARKWAY",
  "HWY", "HIGHWAY", "TRL", "TRAIL", "WAY", "CIR", "CIRCLE", "TER", "TERRACE"
]);
const setSegmentRowKeyMapCache = new WeakMap();

function normalizeSetSegmentIdentityToken(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^a-z0-9]/g, "");
}

function normalizeSetSegmentCoordinateToken(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "";
  let token = num.toFixed(6);
  token = token.replace(/0+$/, "").replace(/\.$/, "");
  if (token === "-0") token = "0";
  return token;
}

function normalizeSetSegmentSideValue(value) {
  const token = String(value ?? "").trim().toLowerCase();
  if (!token) return "";
  if (token.startsWith("l")) return "Left";
  if (token.startsWith("r")) return "Right";
  if (token.startsWith("c") || token.includes("center")) return "Centerline";
  return "";
}

function normalizeSetSegmentMatchMethod(value) {
  const token = String(value ?? "").trim();
  if (!token) return "";
  return token.slice(0, 60);
}

function normalizeSetSegmentToken(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getSetSegmentRowKeyMap(row) {
  if (!row || typeof row !== "object") return new Map();
  const cached = setSegmentRowKeyMapCache.get(row);
  if (cached) return cached;
  const map = new Map();
  Object.keys(row).forEach(key => {
    map.set(normalizeSetSegmentToken(key), key);
  });
  setSegmentRowKeyMapCache.set(row, map);
  return map;
}

function getSetSegmentRowValue(row, aliases = []) {
  if (!row || typeof row !== "object") return "";
  const candidates = Array.isArray(aliases) ? aliases : [];
  for (const alias of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, alias)) return row[alias];
  }
  const keyMap = getSetSegmentRowKeyMap(row);
  for (const alias of candidates) {
    const existing = keyMap.get(normalizeSetSegmentToken(alias));
    if (existing && Object.prototype.hasOwnProperty.call(row, existing)) return row[existing];
  }
  return "";
}

function getSetSegmentSegmentIdString(row, preferredField = "") {
  if (!row || typeof row !== "object") return "";
  let raw = "";
  if (preferredField && Object.prototype.hasOwnProperty.call(row, preferredField)) {
    raw = row[preferredField];
  }
  if (String(raw ?? "").trim() === "") {
    raw = getSetSegmentRowValue(row, ["SEGMENT_ID", "Segment_ID", "segment_id", "SEGMENTID", "SegmentId"]);
  }
  const token = String(raw ?? "").trim();
  if (!token) return "";
  const lowered = token.toLowerCase();
  if (lowered === "unknown" || lowered === "null" || lowered === "undefined") return "";
  return token;
}

function getSetSegmentSegmentSideString(row, preferredField = "") {
  if (!row || typeof row !== "object") return "";
  let raw = "";
  if (preferredField && Object.prototype.hasOwnProperty.call(row, preferredField)) {
    raw = row[preferredField];
  }
  if (String(raw ?? "").trim() === "") {
    raw = getSetSegmentRowValue(row, ["SEGMENT_SIDE", "Segment_Side", "segment_side", "SEGMENTSIDE"]);
  }
  return normalizeSetSegmentSideValue(raw);
}

function getSetSegmentMatchMethodString(row, preferredField = "") {
  if (!row || typeof row !== "object") return "";
  let raw = "";
  if (preferredField && Object.prototype.hasOwnProperty.call(row, preferredField)) {
    raw = row[preferredField];
  }
  if (String(raw ?? "").trim() === "") {
    raw = getSetSegmentRowValue(row, [SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK, "segment_match_method", "SEGMENT_MATCH_METHOD"]);
  }
  return normalizeSetSegmentMatchMethod(raw);
}

function hasSetSegmentMatch(row, segmentIdField = "") {
  return !!getSetSegmentSegmentIdString(row, segmentIdField);
}

function getSetSegmentRecordIdentityKeys(row, lat, lon) {
  if (!row || typeof row !== "object") return [];
  const keys = [];
  const pushUnique = value => {
    const token = String(value || "").trim().toLowerCase();
    if (!token) return;
    if (!keys.includes(token)) keys.push(token);
  };

  const uniqueIdToken = normalizeSetSegmentIdentityToken(
    getSetSegmentRowValue(row, ["UNIQUEID", "UNIQUE_ID", "CUSTOMER_ID", "STOP_ID", "Record_ID", "ID"])
  );
  if (uniqueIdToken) pushUnique(`uid:${uniqueIdToken}`);

  const binToken = normalizeSetSegmentIdentityToken(getSetSegmentRowValue(row, ["BINNO", "BIN_NO", "CONTAINER_ID"]));
  if (binToken) pushUnique(`bin:${binToken}`);

  const addressToken = normalizeSetSegmentIdentityToken([
    getSetSegmentRowValue(row, ["CSADR#", "ADDRESS_NUMBER", "NUMBER"]),
    getSetSegmentRowValue(row, ["CSSDIR", "PREDIR", "STREET_DIR"]),
    getSetSegmentRowValue(row, ["CSSTRT", "STREET", "STREET_NAME", "ROAD"]),
    getSetSegmentRowValue(row, ["CSSFUX", "SUFFIX", "STREET_SUFFIX"])
  ].join(" "));
  if (addressToken) pushUnique(`addr:${addressToken}`);

  const latToken = normalizeSetSegmentCoordinateToken(lat ?? getSetSegmentRowValue(row, ["LATITUDE", "LAT", "Y"]));
  const lonToken = normalizeSetSegmentCoordinateToken(lon ?? getSetSegmentRowValue(row, ["LONGITUDE", "LON", "LNG", "X"]));
  if (latToken && lonToken) pushUnique(`coord:${latToken}|${lonToken}`);

  if (uniqueIdToken && latToken && lonToken) pushUnique(`uidcoord:${uniqueIdToken}|${latToken}|${lonToken}`);
  if (addressToken && latToken && lonToken) pushUnique(`addrcoord:${addressToken}|${latToken}|${lonToken}`);
  if (binToken && latToken && lonToken) pushUnique(`bincoord:${binToken}|${latToken}|${lonToken}`);
  if (uniqueIdToken && addressToken) pushUnique(`uidaddr:${uniqueIdToken}|${addressToken}`);

  return keys;
}

function normalizeSetSegmentMatchProfiles(rawProfiles) {
  const source = Array.isArray(rawProfiles) ? rawProfiles : [];
  const out = [];

  source.forEach((profile, index) => {
    const rawName = String(profile?.name ?? "").trim();
    if (!rawName) return;
    const entriesSource = (profile?.entries && typeof profile.entries === "object" && !Array.isArray(profile.entries))
      ? profile.entries
      : {};
    const entries = {};
    Object.entries(entriesSource).forEach(([rawKey, rawValue]) => {
      const key = String(rawKey || "").trim().toLowerCase();
      if (!key) return;
      let segmentId = "";
      let side = "";
      let method = "";
      if (Array.isArray(rawValue)) {
        segmentId = String(rawValue[0] ?? "").trim();
        side = normalizeSetSegmentSideValue(rawValue[1]);
        method = normalizeSetSegmentMatchMethod(rawValue[2]);
      } else if (rawValue && typeof rawValue === "object") {
        segmentId = String(rawValue.segmentId ?? "").trim();
        side = normalizeSetSegmentSideValue(rawValue.side);
        method = normalizeSetSegmentMatchMethod(rawValue.method);
      } else {
        segmentId = String(rawValue ?? "").trim();
      }
      if (!segmentId) return;
      entries[key] = { segmentId, side, method };
    });
    const entryCount = Object.keys(entries).length;
    if (!entryCount) return;
    const createdAt = String(profile?.createdAt || profile?.updatedAt || "");
    const updatedAt = String(profile?.updatedAt || profile?.createdAt || "");
    out.push({
      id: String(profile?.id || `set-segment-profile-${index + 1}-${Date.now()}`),
      name: rawName,
      createdAt: createdAt || new Date().toISOString(),
      updatedAt: updatedAt || new Date().toISOString(),
      entryCount,
      entries
    });
  });

  out.sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: "base", numeric: true }));
  return out.slice(0, SET_SEGMENT_PROFILE_MAX_COUNT);
}

function loadSetSegmentMatchProfiles() {
  const raw = storageGet(SET_SEGMENT_MATCH_PROFILES_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return normalizeSetSegmentMatchProfiles(parsed);
  } catch {
    return [];
  }
}

function saveSetSegmentMatchProfiles(profiles) {
  const normalized = normalizeSetSegmentMatchProfiles(profiles);
  storageSet(SET_SEGMENT_MATCH_PROFILES_KEY, JSON.stringify(normalized));
  return normalized;
}

function syncSetSegmentMatchStatusFields(rows, options = {}) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    return {
      segmentIdField: "SEGMENT_ID",
      segmentSideField: "SEGMENT_SIDE",
      matchStatusField: SET_SEGMENT_MATCH_STATUS_FIELD_FALLBACK,
      readyField: SET_SEGMENT_SEQUENCE_READY_FIELD_FALLBACK,
      methodField: SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK,
      total: 0,
      matchedCount: 0,
      unmatchedCount: 0,
      readyCount: 0,
      leftCount: 0,
      rightCount: 0,
      centerCount: 0
    };
  }

  const segmentIdField = options.segmentIdField || findSetSegmentOutputFieldName(list, "SEGMENT_ID");
  const segmentSideField = options.segmentSideField || findSetSegmentOutputFieldName(list, "SEGMENT_SIDE");
  const matchStatusField = options.matchStatusField || findSetSegmentOutputFieldName(list, SET_SEGMENT_MATCH_STATUS_FIELD_FALLBACK);
  const readyField = options.readyField || findSetSegmentOutputFieldName(list, SET_SEGMENT_SEQUENCE_READY_FIELD_FALLBACK);
  const methodField = options.methodField || findSetSegmentOutputFieldName(list, SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK);
  const defaultMethod = normalizeSetSegmentMatchMethod(options.defaultMethod);

  let matchedCount = 0;
  let unmatchedCount = 0;
  let readyCount = 0;
  let leftCount = 0;
  let rightCount = 0;
  let centerCount = 0;

  list.forEach(row => {
    if (!row || typeof row !== "object") return;
    const segmentId = getSetSegmentSegmentIdString(row, segmentIdField);
    const side = getSetSegmentSegmentSideString(row, segmentSideField);
    const matched = !!segmentId;
    const ready = matched;

    row[matchStatusField] = matched ? "Matched" : "Unmatched";
    row[readyField] = ready ? "Ready" : "Not Ready";
    row[segmentSideField] = side || "";
    row[methodField] = matched
      ? (getSetSegmentMatchMethodString(row, methodField) || defaultMethod || "")
      : "";

    if (matched) {
      matchedCount += 1;
      readyCount += 1;
      if (side === "Left") leftCount += 1;
      else if (side === "Right") rightCount += 1;
      else if (side === "Centerline") centerCount += 1;
    } else {
      unmatchedCount += 1;
    }
  });

  return {
    segmentIdField,
    segmentSideField,
    matchStatusField,
    readyField,
    methodField,
    total: list.length,
    matchedCount,
    unmatchedCount,
    readyCount,
    leftCount,
    rightCount,
    centerCount
  };
}

function normalizeSetSegmentStreetName(value, parseAddresses = false) {
  let text = String(value ?? "").trim().toUpperCase();
  if (!text) return "";
  text = text.replace(/&/g, " AND ");
  text = text.replace(/[^A-Z0-9\s]/g, " ");
  let tokens = text.split(/\s+/).filter(Boolean);
  if (!tokens.length) return "";

  if (parseAddresses) {
    while (tokens.length && /^\d+[A-Z]?$/.test(tokens[0])) tokens.shift();
    while (tokens.length && SET_SEGMENT_DIRECTION_TOKENS.has(tokens[0])) tokens.shift();
    while (tokens.length && SET_SEGMENT_DIRECTION_TOKENS.has(tokens[tokens.length - 1])) tokens.pop();
    while (tokens.length && SET_SEGMENT_SUFFIX_TOKENS.has(tokens[tokens.length - 1])) tokens.pop();
  }

  return tokens.join(" ").trim();
}

function buildSetSegmentNameKeys(value, parseAddresses = false) {
  const out = new Set();
  const parsed = normalizeSetSegmentStreetName(value, parseAddresses);
  if (parsed) {
    out.add(`s:${parsed}`);
    out.add(`c:${parsed.replace(/\s+/g, "")}`);
  }
  const direct = normalizeSetSegmentStreetName(value, false);
  if (direct) {
    out.add(`s:${direct}`);
    out.add(`c:${direct.replace(/\s+/g, "")}`);
  }
  return out;
}

function collectSetSegmentRecordNameKeys(row, parseAddresses = false) {
  const candidates = [];
  const streetName = getSetSegmentRowValue(row, ["CSSTRT", "STREET_NAME", "STREET", "ROAD_NAME", "ROAD"]);
  const streetDir = getSetSegmentRowValue(row, ["CSSDIR", "PREDIR", "STREET_DIR", "DIR"]);
  const streetSfx = getSetSegmentRowValue(row, ["CSSFUX", "STREET_SUFFIX", "SUFFIX", "POSTTYPE"]);
  const houseNumber = getSetSegmentRowValue(row, ["CSADR#", "ADDRESS_NUMBER", "HOUSE_NUMBER", "NUMBER"]);

  if (String(streetName || "").trim()) {
    candidates.push(`${streetDir || ""} ${streetName || ""} ${streetSfx || ""}`);
    candidates.push(streetName);
    candidates.push(`${houseNumber || ""} ${streetDir || ""} ${streetName || ""} ${streetSfx || ""}`);
  }

  if (parseAddresses) {
    [
      "ADDRESS", "FULL_ADDRESS", "SERVICE_ADDRESS", "SITE_ADDRESS", "ADDR1", "ADDR", "LOCATION", "ADDRESS1"
    ].forEach(fieldName => {
      const value = getSetSegmentRowValue(row, [fieldName]);
      if (String(value || "").trim()) candidates.push(value);
    });
  }

  const keys = new Set();
  candidates.forEach(raw => {
    buildSetSegmentNameKeys(raw, parseAddresses).forEach(key => keys.add(key));
  });
  return keys;
}

function buildSetSegmentStreetMatchIndex(options = {}) {
  const selectedStreetsOnly = !!options.selectedStreetsOnly;
  const includeAliases = !!options.includeAliases;
  const parseAddresses = !!options.parseAddresses;
  const entries = selectedStreetsOnly
    ? [...streetAttributeSelectedIds]
      .map(id => streetAttributeById.get(Number(id)))
      .filter(Boolean)
    : [...streetAttributeById.values()];

  const cellSizeDeg = SET_SEGMENT_MATCH_CELL_DEG;
  const cellMap = new Map();
  const segments = [];
  const streetById = new Map();
  const streetToSegmentIndices = new Map();
  const nameToStreetIds = new Map();

  const addCellRef = (latIdx, lonIdx, segmentIndex) => {
    const key = `${latIdx}:${lonIdx}`;
    let bucket = cellMap.get(key);
    if (!bucket) {
      bucket = [];
      cellMap.set(key, bucket);
    }
    bucket.push(segmentIndex);
  };

  const addStreetName = (streetId, rawName) => {
    const name = String(rawName ?? "").trim();
    if (!name) return;
    buildSetSegmentNameKeys(name, parseAddresses).forEach(key => {
      let ids = nameToStreetIds.get(key);
      if (!ids) {
        ids = new Set();
        nameToStreetIds.set(key, ids);
      }
      ids.add(streetId);
    });
  };

  entries.forEach(entry => {
    const streetId = Number(entry?.id ?? entry?.row?.id);
    if (!Number.isFinite(streetId)) return;
    const row = entry?.row || {};
    const latLon = flattenStreetLayerLatLngs(entry?.layer?.getLatLngs?.(), []);
    if (!Array.isArray(latLon) || latLon.length < 2) return;

    streetById.set(streetId, { id: streetId, row });
    addStreetName(streetId, row?.name);
    addStreetName(streetId, row?.ref);
    if (includeAliases) {
      [
        row?.alt_name,
        row?.official_name,
        row?.old_name,
        row?.loc_name,
        row?.short_name
      ].forEach(alias => {
        String(alias || "")
          .split(/[;|/]/)
          .map(part => part.trim())
          .filter(Boolean)
          .forEach(value => addStreetName(streetId, value));
      });
    }

    for (let i = 1; i < latLon.length; i += 1) {
      const a = latLon[i - 1];
      const b = latLon[i];
      const aLat = Number(a?.lat);
      const aLon = Number(a?.lon);
      const bLat = Number(b?.lat);
      const bLon = Number(b?.lon);
      if (!Number.isFinite(aLat) || !Number.isFinite(aLon) || !Number.isFinite(bLat) || !Number.isFinite(bLon)) continue;
      if (aLat === bLat && aLon === bLon) continue;

      const segment = {
        streetId,
        aLat,
        aLon,
        bLat,
        bLon,
        minLat: Math.min(aLat, bLat),
        maxLat: Math.max(aLat, bLat),
        minLon: Math.min(aLon, bLon),
        maxLon: Math.max(aLon, bLon)
      };
      const segmentIndex = segments.length;
      segments.push(segment);

      let segmentRefs = streetToSegmentIndices.get(streetId);
      if (!segmentRefs) {
        segmentRefs = [];
        streetToSegmentIndices.set(streetId, segmentRefs);
      }
      segmentRefs.push(segmentIndex);

      const minLatIdx = Math.floor(segment.minLat / cellSizeDeg);
      const maxLatIdx = Math.floor(segment.maxLat / cellSizeDeg);
      const minLonIdx = Math.floor(segment.minLon / cellSizeDeg);
      const maxLonIdx = Math.floor(segment.maxLon / cellSizeDeg);
      for (let latIdx = minLatIdx; latIdx <= maxLatIdx; latIdx += 1) {
        for (let lonIdx = minLonIdx; lonIdx <= maxLonIdx; lonIdx += 1) {
          addCellRef(latIdx, lonIdx, segmentIndex);
        }
      }
    }
  });

  return {
    cellSizeDeg,
    cellMap,
    segments,
    streetById,
    streetToSegmentIndices,
    nameToStreetIds,
    segmentCount: segments.length,
    streetCount: streetById.size
  };
}

function measureSetSegmentPointToSegmentFeet(lat, lon, segment) {
  const refLat = ((Number(lat) + Number(segment?.aLat) + Number(segment?.bLat)) / 3) * (Math.PI / 180);
  const feetPerDegLon = SET_SEGMENT_FEET_PER_DEG_LAT * Math.max(0.05, Math.cos(refLat));
  const ax = (Number(segment?.aLon) - Number(lon)) * feetPerDegLon;
  const ay = (Number(segment?.aLat) - Number(lat)) * SET_SEGMENT_FEET_PER_DEG_LAT;
  const bx = (Number(segment?.bLon) - Number(lon)) * feetPerDegLon;
  const by = (Number(segment?.bLat) - Number(lat)) * SET_SEGMENT_FEET_PER_DEG_LAT;

  const abx = bx - ax;
  const aby = by - ay;
  const lenSq = (abx * abx) + (aby * aby);
  let t = 0;
  if (lenSq > 1e-9) {
    t = (-(ax * abx + ay * aby)) / lenSq;
    t = Math.max(0, Math.min(1, t));
  }

  const cx = ax + (abx * t);
  const cy = ay + (aby * t);
  const distanceFeet = Math.hypot(cx, cy);
  const apx = -ax;
  const apy = -ay;
  const sideCross = (abx * apy) - (aby * apx);
  return { distanceFeet, sideCross, t };
}

function getSetSegmentCandidateIndicesByRadius(index, lat, lon, maxDistanceFeet) {
  if (!index?.segments?.length) return [];
  const latNum = Number(lat);
  const lonNum = Number(lon);
  if (!Number.isFinite(latNum) || !Number.isFinite(lonNum)) return [];
  const maxFeet = Number(maxDistanceFeet);

  if (!Number.isFinite(maxFeet) || maxFeet <= 0) {
    return index.segments.map((_, idx) => idx);
  }

  const radiusLatDeg = maxFeet / SET_SEGMENT_FEET_PER_DEG_LAT;
  const radiusCells = Math.max(1, Math.ceil(radiusLatDeg / index.cellSizeDeg) + 1);
  const baseLatIdx = Math.floor(latNum / index.cellSizeDeg);
  const baseLonIdx = Math.floor(lonNum / index.cellSizeDeg);
  const out = new Set();

  for (let dy = -radiusCells; dy <= radiusCells; dy += 1) {
    for (let dx = -radiusCells; dx <= radiusCells; dx += 1) {
      const key = `${baseLatIdx + dy}:${baseLonIdx + dx}`;
      const bucket = index.cellMap.get(key);
      if (!bucket || !bucket.length) continue;
      bucket.forEach(segmentIndex => out.add(segmentIndex));
    }
  }

  return [...out];
}

function findBestSetSegmentMatchForPoint(index, lat, lon, options = {}) {
  if (!index?.segments?.length) return null;
  const allowedStreetIds = options.allowedStreetIds instanceof Set ? options.allowedStreetIds : null;
  const maxDistanceFeet = Number(options.maxDistanceFeet);
  const hasMaxDistance = Number.isFinite(maxDistanceFeet) && maxDistanceFeet > 0;
  const indices = Array.isArray(options.candidateIndices)
    ? options.candidateIndices
    : getSetSegmentCandidateIndicesByRadius(index, lat, lon, hasMaxDistance ? maxDistanceFeet : NaN);
  if (!indices.length) return null;

  let best = null;
  const visited = new Set();
  indices.forEach(segmentIndex => {
    const idx = Number(segmentIndex);
    if (!Number.isFinite(idx) || visited.has(idx)) return;
    visited.add(idx);
    const segment = index.segments[idx];
    if (!segment) return;
    if (allowedStreetIds && !allowedStreetIds.has(segment.streetId)) return;
    const measure = measureSetSegmentPointToSegmentFeet(lat, lon, segment);
    if (!Number.isFinite(measure.distanceFeet)) return;
    if (hasMaxDistance && measure.distanceFeet > maxDistanceFeet) return;
    if (!best || measure.distanceFeet < best.distanceFeet) {
      best = {
        segmentIndex: idx,
        streetId: segment.streetId,
        distanceFeet: measure.distanceFeet,
        sideCross: measure.sideCross
      };
    }
  });

  return best;
}

function getSetSegmentRecordCandidates(scope = "all") {
  const selectedOnly = String(scope || "all") === "selected";
  const out = [];

  attributeMarkerByRowId.forEach((marker, rowId) => {
    const numericRowId = Number(rowId);
    if (!Number.isFinite(numericRowId)) return;
    if (selectedOnly && !attributeState.selectedRowIds.has(numericRowId)) return;
    const row = marker?._rowRef;
    if (!row || typeof row !== "object") return;
    const ll = getLayerLatLng(marker);
    const lat = Number(ll?.lat ?? marker?._base?.lat ?? row?.LATITUDE);
    const lon = Number(ll?.lng ?? marker?._base?.lon ?? row?.LONGITUDE);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    out.push({
      rowId: numericRowId,
      row,
      marker,
      lat,
      lon
    });
  });

  out.sort((a, b) => Number(a.rowId) - Number(b.rowId));
  return out;
}

function getSetSegmentReadinessSnapshot(scope = "all", segmentIdField = "") {
  const records = getSetSegmentRecordCandidates(scope);
  let matchedCount = 0;
  let unmatchedCount = 0;
  records.forEach(item => {
    if (hasSetSegmentMatch(item?.row, segmentIdField)) matchedCount += 1;
    else unmatchedCount += 1;
  });
  return {
    total: records.length,
    matchedCount,
    unmatchedCount,
    readyCount: matchedCount
  };
}

function findSetSegmentOutputFieldName(rows, fallback) {
  const preferred = normalizeSetSegmentToken(fallback);
  const fields = new Set();
  (Array.isArray(rows) ? rows : []).forEach(row => {
    Object.keys(row || {}).forEach(key => fields.add(String(key)));
  });
  for (const field of fields) {
    if (normalizeSetSegmentToken(field) === preferred) return field;
  }
  return fallback;
}

function matchRecordToStreetForSegmentId(recordItem, index, settings = {}) {
  if (!recordItem || !index) return null;
  const useStreetName = !!settings.matchByStreetName;
  const useProximity = !!settings.useProximity;
  const toleranceFeetRaw = Number(settings.toleranceFeet);
  const toleranceFeet = Number.isFinite(toleranceFeetRaw) && toleranceFeetRaw > 0 ? toleranceFeetRaw : NaN;
  const parseAddresses = !!settings.parseAddresses;

  let matchedStreetIds = null;
  if (useStreetName) {
    const nameKeys = collectSetSegmentRecordNameKeys(recordItem.row, parseAddresses);
    matchedStreetIds = new Set();
    nameKeys.forEach(key => {
      const ids = index.nameToStreetIds.get(key);
      if (!ids) return;
      ids.forEach(id => matchedStreetIds.add(id));
    });

    if (!matchedStreetIds.size && !useProximity) return null;
  }

  let best = null;
  let method = "";

  if (useStreetName && matchedStreetIds && matchedStreetIds.size) {
    if (useProximity) {
      best = findBestSetSegmentMatchForPoint(index, recordItem.lat, recordItem.lon, {
        maxDistanceFeet: toleranceFeet,
        allowedStreetIds: matchedStreetIds
      });
      if (best) method = "name+proximity";
    } else {
      const candidateIndices = [];
      matchedStreetIds.forEach(streetId => {
        const indices = index.streetToSegmentIndices.get(streetId);
        if (Array.isArray(indices) && indices.length) candidateIndices.push(...indices);
      });
      best = findBestSetSegmentMatchForPoint(index, recordItem.lat, recordItem.lon, {
        maxDistanceFeet: toleranceFeet,
        candidateIndices
      });
      if (best) method = "name";
    }
  }

  if (!best && useProximity) {
    best = findBestSetSegmentMatchForPoint(index, recordItem.lat, recordItem.lon, {
      maxDistanceFeet: toleranceFeet
    });
    if (best) method = useStreetName ? "proximity-fallback" : "proximity";
  }

  if (!best) return null;
  const street = index.streetById.get(best.streetId);
  const side = Math.abs(best.sideCross) <= SET_SEGMENT_ZERO_SIDE_FEET
    ? "Centerline"
    : (best.sideCross > 0 ? "Left" : "Right");

  return {
    streetId: best.streetId,
    side,
    distanceFeet: best.distanceFeet,
    method,
    streetName: String(street?.row?.name || street?.row?.ref || "").trim()
  };
}

function loadSetSegmentIdSideSettings() {
  const raw = storageGet(SET_SEGMENT_ID_SIDE_SETTINGS_KEY);
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveSetSegmentIdSideSettings(settings) {
  const payload = settings && typeof settings === "object" ? settings : {};
  storageSet(SET_SEGMENT_ID_SIDE_SETTINGS_KEY, JSON.stringify(payload));
}

function initSetSegmentIdSideControls() {
  const openBtn = document.getElementById("setSegmentIdSideBtn");
  const openBtnMobile = document.getElementById("setSegmentIdSideBtnMobile");
  const modal = document.getElementById("setSegmentIdSideModal");
  const closeBtn = document.getElementById("setSegmentIdSideCloseBtn");
  const runBtn = document.getElementById("setSegmentIdSideRunBtn");
  const inputAllRadio = document.getElementById("setSegmentInputAllRecords");
  const inputSelectedRadio = document.getElementById("setSegmentInputSelectedRecords");
  const inputAllLabel = document.getElementById("setSegmentInputAllLabel");
  const inputSelectedLabel = document.getElementById("setSegmentInputSelectedLabel");
  const selectedStreetsOnlyToggle = document.getElementById("setSegmentSelectedStreetsOnly");
  const selectedStreetsOnlyLabel = document.getElementById("setSegmentSelectedStreetsOnlyLabel");
  const parseAddressesToggle = document.getElementById("setSegmentParseAddresses");
  const matchStreetNameToggle = document.getElementById("setSegmentMatchStreetName");
  const useProximityToggle = document.getElementById("setSegmentUseProximity");
  const includeAliasesToggle = document.getElementById("setSegmentIncludeStreetAliases");
  const toleranceInput = document.getElementById("setSegmentSearchToleranceFeet");
  const statusNode = document.getElementById("setSegmentIdSideStatus");
  const layerSelect = document.getElementById("setSegmentIdSideLayer");
  const skipMatchedToggle = document.getElementById("setSegmentSkipMatchedRecords");
  const profileSelect = document.getElementById("setSegmentSavedProfileList");
  const profileNameInput = document.getElementById("setSegmentSavedProfileName");
  const applyOnlyUnmatchedToggle = document.getElementById("setSegmentApplyOnlyUnmatched");
  const saveProfileBtn = document.getElementById("setSegmentSaveProfileBtn");
  const applyProfileBtn = document.getElementById("setSegmentApplyProfileBtn");
  const deleteProfileBtn = document.getElementById("setSegmentDeleteProfileBtn");
  const readySummaryNode = document.getElementById("setSegmentReadySummary");

  if (
    !openBtn || !modal || !closeBtn || !runBtn ||
    !inputAllRadio || !inputSelectedRadio || !inputAllLabel || !inputSelectedLabel ||
    !selectedStreetsOnlyToggle || !selectedStreetsOnlyLabel || !parseAddressesToggle ||
    !matchStreetNameToggle || !useProximityToggle || !includeAliasesToggle || !toleranceInput || !statusNode || !layerSelect
  ) {
    return;
  }

  if (openBtn.dataset.segmentIdSideBound === "1") return;
  openBtn.dataset.segmentIdSideBound = "1";

  const setStatus = (message, kind = "") => {
    statusNode.textContent = String(message || "");
    statusNode.classList.remove("error", "success");
    if (kind === "error") statusNode.classList.add("error");
    if (kind === "success") statusNode.classList.add("success");
  };

  const refreshProfileList = (preferredName = "") => {
    if (!profileSelect) return [];
    const profiles = loadSetSegmentMatchProfiles();
    const desiredName = String(preferredName || profileSelect.value || "").trim();
    profileSelect.innerHTML = "";

    if (!profiles.length) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "No saved profiles";
      profileSelect.appendChild(option);
      profileSelect.disabled = true;
      if (applyProfileBtn) applyProfileBtn.disabled = true;
      if (deleteProfileBtn) deleteProfileBtn.disabled = true;
      return profiles;
    }

    profiles.forEach(profile => {
      const option = document.createElement("option");
      option.value = profile.name;
      option.textContent = `${profile.name} (${Number(profile.entryCount || 0).toLocaleString()})`;
      profileSelect.appendChild(option);
    });

    const selected = profiles.some(profile => profile.name === desiredName)
      ? desiredName
      : profiles[0].name;
    profileSelect.value = selected;
    profileSelect.disabled = false;
    if (applyProfileBtn) applyProfileBtn.disabled = false;
    if (deleteProfileBtn) deleteProfileBtn.disabled = false;
    return profiles;
  };

  const refreshReadySummary = (segmentIdField = "") => {
    if (!readySummaryNode) return;
    const allSnapshot = getSetSegmentReadinessSnapshot("all", segmentIdField);
    const selectedSnapshot = getSetSegmentReadinessSnapshot("selected", segmentIdField);
    if (!allSnapshot.total) {
      readySummaryNode.textContent = "No mappable records loaded yet. Load a route file to view match readiness.";
      return;
    }
    readySummaryNode.textContent =
      `All: ${allSnapshot.matchedCount.toLocaleString()}/${allSnapshot.total.toLocaleString()} matched and ready. ` +
      `Selected: ${selectedSnapshot.matchedCount.toLocaleString()}/${selectedSnapshot.total.toLocaleString()} matched and ready.`;
  };

  const applySettingsToUi = settings => {
    const payload = settings && typeof settings === "object" ? settings : {};
    inputSelectedRadio.checked = payload.scope === "selected";
    inputAllRadio.checked = !inputSelectedRadio.checked;
    parseAddressesToggle.checked = !!payload.parseAddresses;
    matchStreetNameToggle.checked = payload.matchByStreetName !== false;
    useProximityToggle.checked = payload.useProximity !== false;
    includeAliasesToggle.checked = !!payload.includeAliases;
    selectedStreetsOnlyToggle.checked = !!payload.selectedStreetsOnly;
    if (skipMatchedToggle) skipMatchedToggle.checked = payload.skipMatchedRecords !== false;
    if (applyOnlyUnmatchedToggle) applyOnlyUnmatchedToggle.checked = payload.applyOnlyUnmatched !== false;
    const tolerance = Number(payload.toleranceFeet);
    toleranceInput.value = Number.isFinite(tolerance) && tolerance > 0 ? String(Math.round(tolerance)) : "9999";
    layerSelect.value = "serviceLocation";
  };

  const collectUiSettings = () => {
    const payload = {
      scope: inputSelectedRadio.checked ? "selected" : "all",
      parseAddresses: !!parseAddressesToggle.checked,
      matchByStreetName: !!matchStreetNameToggle.checked,
      useProximity: !!useProximityToggle.checked,
      includeAliases: !!includeAliasesToggle.checked,
      selectedStreetsOnly: !!selectedStreetsOnlyToggle.checked,
      skipMatchedRecords: skipMatchedToggle ? !!skipMatchedToggle.checked : true,
      applyOnlyUnmatched: applyOnlyUnmatchedToggle ? !!applyOnlyUnmatchedToggle.checked : true,
      selectedProfileName: profileSelect ? String(profileSelect.value || "").trim() : "",
      toleranceFeet: Math.max(10, Math.round(clampRouteSequencerNumber(toleranceInput.value, 10, 250000, 9999)))
    };
    toleranceInput.value = String(payload.toleranceFeet);
    saveSetSegmentIdSideSettings(payload);
    return payload;
  };

  const refreshCounts = () => {
    const allRecords = getSetSegmentRecordCandidates("all").length;
    const selectedRecords = getSetSegmentRecordCandidates("selected").length;
    const selectedStreets = streetAttributeSelectedIds.size;
    inputAllLabel.textContent = `All records (${allRecords.toLocaleString()})`;
    inputSelectedLabel.textContent = `Selected records (${selectedRecords.toLocaleString()})`;
    selectedStreetsOnlyLabel.textContent = `Consider selected streets only (${selectedStreets.toLocaleString()})`;
    refreshReadySummary();
  };

  const refreshUiAfterRowChanges = (saveReason = "Set Segment ID + Side", recordsToRefresh = []) => {
    const allRows = (Array.isArray(window._currentRows) && window._currentRows.length)
      ? window._currentRows
      : recordsToRefresh.map(item => item?.row).filter(Boolean);
    window._attributeHeaders = getAttributeHeaders(allRows);
    if (typeof refreshAllRecordPopupContent === "function") refreshAllRecordPopupContent();
    else recordsToRefresh.forEach(item => syncRecordPopupContent(item.marker));
    applyAttributeSelectionStyles();
    renderAttributeTable();
    refreshAttributeStatus();
    refreshCounts();
    syncCurrentWorkbookSheetFromRows();
    queueWorkbookCloudSave(saveReason);

    const selectAttributesSource = document.getElementById("selectAttributesSource");
    if (selectAttributesSource && String(selectAttributesSource.value || "").toLowerCase() === "records") {
      selectAttributesSource.dispatchEvent(new Event("change"));
    }
  };

  const openModal = () => {
    const storedSettings = loadSetSegmentIdSideSettings();
    applySettingsToUi(storedSettings);
    refreshProfileList(storedSettings.selectedProfileName || "");
    if (profileNameInput && !String(profileNameInput.value || "").trim() && profileSelect && !profileSelect.disabled) {
      profileNameInput.value = String(profileSelect.value || "");
    }
    refreshCounts();
    setStatus("Ready.");
    modal.style.display = "flex";
    openBtn.classList.add("active");
  };

  const closeModal = () => {
    modal.style.display = "none";
    openBtn.classList.remove("active");
  };

  const runMatcher = async () => {
    if (runBtn.dataset.running === "1") return;
    runBtn.dataset.running = "1";
    runBtn.disabled = true;
    try {
      if (!streetAttributeById.size) {
        setStatus("No street segments are loaded. Load streets first, then run this tool.", "error");
        return;
      }

      const settings = collectUiSettings();
      const allRows = (Array.isArray(window._currentRows) && window._currentRows.length)
        ? window._currentRows
        : getSetSegmentRecordCandidates("all").map(item => item.row);
      const segmentIdField = findSetSegmentOutputFieldName(allRows, "SEGMENT_ID");
      const segmentSideField = findSetSegmentOutputFieldName(allRows, "SEGMENT_SIDE");
      const matchStatusField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_MATCH_STATUS_FIELD_FALLBACK);
      const readyField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_SEQUENCE_READY_FIELD_FALLBACK);
      const methodField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK);

      if (!settings.matchByStreetName && !settings.useProximity) {
        setStatus("Enable at least one match method: street name or proximity.", "error");
        return;
      }
      if (settings.selectedStreetsOnly && !streetAttributeSelectedIds.size) {
        setStatus("Selected streets only is enabled, but no street segments are selected.", "error");
        return;
      }

      const records = getSetSegmentRecordCandidates(settings.scope);
      if (!records.length) {
        setStatus(
          settings.scope === "selected"
            ? "No selected records are available to process."
            : "No mappable records are available to process.",
          "error"
        );
        return;
      }

      const skippedMatchedCount = settings.skipMatchedRecords
        ? records.filter(item => hasSetSegmentMatch(item.row, segmentIdField)).length
        : 0;
      const processRecords = settings.skipMatchedRecords
        ? records.filter(item => !hasSetSegmentMatch(item.row, segmentIdField))
        : records;

      if (!processRecords.length) {
        const syncSummary = syncSetSegmentMatchStatusFields(allRows, {
          segmentIdField,
          segmentSideField,
          matchStatusField,
          readyField,
          methodField
        });
        refreshUiAfterRowChanges("Set Segment ID + Side (No Changes)", records);
        setStatus(
          `All ${records.length.toLocaleString()} records in scope are already matched. Ready: ${syncSummary.readyCount.toLocaleString()}/${syncSummary.total.toLocaleString()}.`,
          "success"
        );
        return;
      }

      setStatus("Building street matching index...");
      const index = buildSetSegmentStreetMatchIndex({
        selectedStreetsOnly: settings.selectedStreetsOnly,
        includeAliases: settings.includeAliases,
        parseAddresses: settings.parseAddresses
      });
      if (!index.segmentCount) {
        setStatus("No street geometry is available for matching under the current street scope.", "error");
        return;
      }

      let matchedCount = 0;
      let unmatchedCount = 0;
      let leftCount = 0;
      let rightCount = 0;
      let centerCount = 0;

      for (let i = 0; i < processRecords.length; i += 1) {
        const item = processRecords[i];
        const match = matchRecordToStreetForSegmentId(item, index, settings);
        if (match) {
          item.row[segmentIdField] = Number.isFinite(Number(match.streetId)) ? Number(match.streetId) : "";
          item.row[segmentSideField] = match.side || "";
          item.row[methodField] = normalizeSetSegmentMatchMethod(match.method || "matched");
          matchedCount += 1;
          if (match.side === "Left") leftCount += 1;
          else if (match.side === "Right") rightCount += 1;
          else centerCount += 1;
        } else {
          item.row[segmentIdField] = "";
          item.row[segmentSideField] = "";
          item.row[methodField] = "";
          unmatchedCount += 1;
        }

        if (i % 180 === 0) {
          setStatus(
            `Matching records... ${Math.min(i + 1, processRecords.length).toLocaleString()}/${processRecords.length.toLocaleString()}`
          );
          await new Promise(resolve => setTimeout(resolve, 0));
        }
      }

      const syncSummary = syncSetSegmentMatchStatusFields(allRows, {
        segmentIdField,
        segmentSideField,
        matchStatusField,
        readyField,
        methodField
      });
      refreshUiAfterRowChanges("Set Segment ID + Side", records);

      setStatus(
        `Set Segment ID + Side complete. Processed ${processRecords.length.toLocaleString()} records` +
        (skippedMatchedCount ? ` (skipped ${skippedMatchedCount.toLocaleString()} already matched).` : ".") +
        ` Newly matched: ${matchedCount.toLocaleString()}, unmatched: ${unmatchedCount.toLocaleString()}.` +
        ` Side split - Left: ${leftCount.toLocaleString()}, Right: ${rightCount.toLocaleString()}, Centerline: ${centerCount.toLocaleString()}.` +
        ` Ready now: ${syncSummary.readyCount.toLocaleString()}/${syncSummary.total.toLocaleString()}.`,
        "success"
      );
    } catch (error) {
      console.error("Set Segment ID + Side failed:", error);
      setStatus(`Set Segment ID + Side failed: ${error?.message || error}`, "error");
    } finally {
      runBtn.disabled = false;
      runBtn.dataset.running = "0";
    }
  };

  const saveCurrentMatchesProfile = () => {
    const settings = collectUiSettings();
    const records = getSetSegmentRecordCandidates(settings.scope);
    if (!records.length) {
      setStatus("No records in scope to save as a profile.", "error");
      return;
    }

    const allRows = Array.isArray(window._currentRows) ? window._currentRows : records.map(item => item.row);
    const segmentIdField = findSetSegmentOutputFieldName(allRows, "SEGMENT_ID");
    const segmentSideField = findSetSegmentOutputFieldName(allRows, "SEGMENT_SIDE");
    const methodField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK);

    const profileNameRaw = String(profileNameInput?.value || profileSelect?.value || "").trim();
    if (!profileNameRaw) {
      setStatus("Type a profile name before saving.", "error");
      return;
    }
    const profileName = profileNameRaw.slice(0, 80);

    const entries = {};
    let matchedInScope = 0;
    records.forEach(item => {
      const segmentId = getSetSegmentSegmentIdString(item.row, segmentIdField);
      if (!segmentId) return;
      matchedInScope += 1;
      const side = getSetSegmentSegmentSideString(item.row, segmentSideField);
      const method = getSetSegmentMatchMethodString(item.row, methodField);
      const keys = getSetSegmentRecordIdentityKeys(item.row, item.lat, item.lon);
      keys.forEach(key => {
        if (!entries[key]) {
          entries[key] = { segmentId, side, method };
        }
      });
    });

    const entryCount = Object.keys(entries).length;
    if (!entryCount) {
      setStatus("No matched records were found in the chosen scope. Run Set Segment ID + Side first.", "error");
      return;
    }

    const profiles = loadSetSegmentMatchProfiles();
    const existingIndex = profiles.findIndex(profile =>
      String(profile?.name || "").trim().toLowerCase() === profileName.toLowerCase()
    );
    const nowIso = new Date().toISOString();
    const nextProfile = {
      id: existingIndex >= 0
        ? profiles[existingIndex].id
        : `set-segment-profile-${Date.now()}-${Math.floor(Math.random() * 1e6)}`,
      name: profileName,
      createdAt: existingIndex >= 0 ? profiles[existingIndex].createdAt : nowIso,
      updatedAt: nowIso,
      entryCount,
      entries
    };
    if (existingIndex >= 0) profiles[existingIndex] = nextProfile;
    else profiles.push(nextProfile);
    const savedProfiles = saveSetSegmentMatchProfiles(profiles);
    const selectedName = savedProfiles.find(profile =>
      String(profile?.name || "").trim().toLowerCase() === profileName.toLowerCase()
    )?.name || profileName;
    refreshProfileList(selectedName);
    if (profileNameInput) profileNameInput.value = selectedName;
    collectUiSettings();

    setStatus(
      `Saved profile "${selectedName}" with ${entryCount.toLocaleString()} mapped keys from ${matchedInScope.toLocaleString()} matched records.`,
      "success"
    );
  };

  const applySavedProfile = async () => {
    const settings = collectUiSettings();
    const selectedProfileName = String(profileSelect?.value || "").trim();
    if (!selectedProfileName) {
      setStatus("Choose a saved profile to apply.", "error");
      return;
    }

    const profiles = loadSetSegmentMatchProfiles();
    const profile = profiles.find(item => item.name === selectedProfileName);
    if (!profile) {
      setStatus(`Saved profile "${selectedProfileName}" was not found.`, "error");
      refreshProfileList("");
      return;
    }

    const records = getSetSegmentRecordCandidates(settings.scope);
    if (!records.length) {
      setStatus("No records in scope to apply profile mappings.", "error");
      return;
    }

    const allRows = Array.isArray(window._currentRows) ? window._currentRows : records.map(item => item.row);
    const segmentIdField = findSetSegmentOutputFieldName(allRows, "SEGMENT_ID");
    const segmentSideField = findSetSegmentOutputFieldName(allRows, "SEGMENT_SIDE");
    const matchStatusField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_MATCH_STATUS_FIELD_FALLBACK);
    const readyField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_SEQUENCE_READY_FIELD_FALLBACK);
    const methodField = findSetSegmentOutputFieldName(allRows, SET_SEGMENT_MATCH_METHOD_FIELD_FALLBACK);
    const applyOnlyUnmatched = applyOnlyUnmatchedToggle ? !!applyOnlyUnmatchedToggle.checked : true;

    let appliedCount = 0;
    let alreadyMatchedSkipped = 0;
    let noKeyCount = 0;
    let noMatchCount = 0;

    for (let i = 0; i < records.length; i += 1) {
      const item = records[i];
      if (applyOnlyUnmatched && hasSetSegmentMatch(item.row, segmentIdField)) {
        alreadyMatchedSkipped += 1;
        continue;
      }

      const keys = getSetSegmentRecordIdentityKeys(item.row, item.lat, item.lon);
      if (!keys.length) {
        noKeyCount += 1;
        continue;
      }

      let matchedEntry = null;
      for (const key of keys) {
        if (!Object.prototype.hasOwnProperty.call(profile.entries, key)) continue;
        matchedEntry = profile.entries[key];
        break;
      }
      if (!matchedEntry) {
        noMatchCount += 1;
        continue;
      }

      item.row[segmentIdField] = String(matchedEntry.segmentId || "").trim();
      item.row[segmentSideField] = normalizeSetSegmentSideValue(matchedEntry.side);
      item.row[methodField] = normalizeSetSegmentMatchMethod(matchedEntry.method || "saved-profile");
      appliedCount += 1;

      if (i % 220 === 0) {
        setStatus(
          `Applying "${selectedProfileName}"... ${Math.min(i + 1, records.length).toLocaleString()}/${records.length.toLocaleString()}`
        );
        await new Promise(resolve => setTimeout(resolve, 0));
      }
    }

    const syncSummary = syncSetSegmentMatchStatusFields(allRows, {
      segmentIdField,
      segmentSideField,
      matchStatusField,
      readyField,
      methodField
    });
    refreshUiAfterRowChanges("Apply Segment Match Profile", records);
    setStatus(
      `Applied profile "${selectedProfileName}". Updated ${appliedCount.toLocaleString()} records` +
      (applyOnlyUnmatched ? `, skipped ${alreadyMatchedSkipped.toLocaleString()} already matched` : "") +
      `. No key: ${noKeyCount.toLocaleString()}, no profile match: ${noMatchCount.toLocaleString()}. ` +
      `Ready now: ${syncSummary.readyCount.toLocaleString()}/${syncSummary.total.toLocaleString()}.`,
      "success"
    );
  };

  const deleteSelectedProfile = () => {
    const selectedProfileName = String(profileSelect?.value || "").trim();
    if (!selectedProfileName) {
      setStatus("Choose a profile to delete.", "error");
      return;
    }

    const confirmed = window.confirm(`Delete saved match profile "${selectedProfileName}"?`);
    if (!confirmed) return;

    const profiles = loadSetSegmentMatchProfiles();
    const nextProfiles = profiles.filter(profile => profile.name !== selectedProfileName);
    if (nextProfiles.length === profiles.length) {
      setStatus(`Profile "${selectedProfileName}" was not found.`, "error");
      refreshProfileList("");
      return;
    }
    saveSetSegmentMatchProfiles(nextProfiles);
    refreshProfileList("");
    if (profileNameInput) profileNameInput.value = "";
    collectUiSettings();
    setStatus(`Deleted profile "${selectedProfileName}".`, "success");
  };

  openBtn.addEventListener("click", openModal);
  if (openBtnMobile) openBtnMobile.addEventListener("click", openModal);
  closeBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", event => {
    if (event.target === modal) closeModal();
  });
  runBtn.addEventListener("click", () => {
    runMatcher().catch(error => {
      console.error("Set Segment ID + Side uncaught error:", error);
      setStatus(`Set Segment ID + Side failed: ${error?.message || error}`, "error");
    });
  });
  saveProfileBtn?.addEventListener("click", saveCurrentMatchesProfile);
  applyProfileBtn?.addEventListener("click", () => {
    applySavedProfile().catch(error => {
      console.error("Apply Set Segment profile failed:", error);
      setStatus(`Apply saved profile failed: ${error?.message || error}`, "error");
    });
  });
  deleteProfileBtn?.addEventListener("click", deleteSelectedProfile);
  profileSelect?.addEventListener("change", () => {
    if (profileNameInput) profileNameInput.value = String(profileSelect.value || "");
    collectUiSettings();
  });
  profileNameInput?.addEventListener("change", () => {
    profileNameInput.value = String(profileNameInput.value || "").trim().slice(0, 80);
  });

  [
    inputAllRadio,
    inputSelectedRadio,
    selectedStreetsOnlyToggle,
    parseAddressesToggle,
    matchStreetNameToggle,
    useProximityToggle,
    includeAliasesToggle,
    toleranceInput,
    skipMatchedToggle,
    applyOnlyUnmatchedToggle
  ].forEach(node => {
    node?.addEventListener("change", () => {
      collectUiSettings();
      refreshCounts();
    });
  });
}

function initServiceDayLabelLayerControls() {
  const toggleHeader = document.getElementById("serviceDayLabelLayerToggle");
  const content = document.getElementById("serviceDayLabelLayerContent");
  const allBtn = document.getElementById("serviceDayLabelLayerAll");
  const noneBtn = document.getElementById("serviceDayLabelLayerNone");
  const daysAllBtn = document.getElementById("serviceDayLabelDaysAll");
  const daysNoneBtn = document.getElementById("serviceDayLabelDaysNone");
  const freqAllBtn = document.getElementById("serviceDayLabelFreqAll");
  const freqNoneBtn = document.getElementById("serviceDayLabelFreqNone");
  const masterToggle = document.getElementById("serviceDayLabelLayerEnabled");
  const dayList = document.getElementById("serviceDayLabelDayFilters");
  const frequencyList = document.getElementById("serviceDayLabelFrequencyFilters");

  if (!toggleHeader || !content || !allBtn || !noneBtn || !daysAllBtn || !daysNoneBtn || !freqAllBtn || !freqNoneBtn || !masterToggle || !dayList || !frequencyList) return;
  if (toggleHeader.dataset.serviceDayLabelBound === "1") return;
  toggleHeader.dataset.serviceDayLabelBound = "1";

  content.classList.add("collapsed");
  toggleHeader.classList.remove("open");

  const sortDayEntries = () => {
    return [...multiDayLabelDayCatalog.values()].sort((a, b) => {
      const rankDiff = multiDaySortRank(a?.token) - multiDaySortRank(b?.token);
      if (rankDiff !== 0) return rankDiff;
      return String(a?.detailLabel || "").localeCompare(String(b?.detailLabel || ""), undefined, { sensitivity: "base" });
    });
  };

  const sortFrequencyEntries = () => {
    return [...multiDayFrequencyCatalog.values()].sort((a, b) => {
      if ((b?.dayCount || 0) !== (a?.dayCount || 0)) return (b?.dayCount || 0) - (a?.dayCount || 0);
      return String(a?.detailLabel || "").localeCompare(String(b?.detailLabel || ""), undefined, { sensitivity: "base" });
    });
  };

  const renderEmptyState = (node, text) => {
    node.innerHTML = "";
    const empty = document.createElement("div");
    empty.className = "service-day-label-empty";
    empty.textContent = text;
    node.appendChild(empty);
  };

  const renderDayList = () => {
    const entries = sortDayEntries();
    if (!entries.length) {
      renderEmptyState(dayList, "No service days found yet.");
      return;
    }

    dayList.innerHTML = "";
    entries.forEach(entry => {
      const item = document.createElement("div");
      item.className = "service-day-label-item";

      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.dayToken = entry.token;
      checkbox.checked = isMultiDayLabelDayVisibleByToken(entry.token);

      const title = document.createElement("span");
      title.className = "service-day-label-item-title";
      title.textContent = entry.detailLabel || entry.shortLabel || entry.token;

      const meta = document.createElement("span");
      meta.className = "service-day-label-item-meta";
      const locationCount = Number(entry.locationCount) || 0;
      meta.textContent = `${locationCount.toLocaleString()} labels`;

      label.append(checkbox, title);
      item.append(label, meta);
      dayList.appendChild(item);
    });
  };

  const renderFrequencyList = () => {
    const entries = sortFrequencyEntries();
    if (!entries.length) {
      renderEmptyState(frequencyList, "No service frequencies found yet.");
      return;
    }

    frequencyList.innerHTML = "";
    entries.forEach(entry => {
      const item = document.createElement("div");
      item.className = "service-day-label-item";

      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.frequencyKey = entry.key;
      checkbox.checked = isMultiDayLabelFrequencyVisibleByKey(entry.key);

      const title = document.createElement("span");
      title.className = "service-day-label-item-title";
      title.textContent = entry.detailLabel || "Unknown";

      const meta = document.createElement("span");
      meta.className = "service-day-label-item-meta";
      const locationCount = Number(entry.locationCount) || 0;
      meta.textContent = `${locationCount.toLocaleString()} labels`;

      label.append(checkbox, title);
      item.append(label, meta);
      frequencyList.appendChild(item);
    });
  };

  const renderSidebarUi = () => {
    masterToggle.checked = !!multiDayServiceEnabled;
    renderDayList();
    renderFrequencyList();
  };

  toggleHeader.addEventListener("click", event => {
    if (event.target === allBtn || event.target === noneBtn) return;
    const collapsed = content.classList.toggle("collapsed");
    toggleHeader.classList.toggle("open", !collapsed);
  });

  allBtn.addEventListener("click", () => {
    setMultiDayServiceEnabled(true);
    setAllMultiDayLabelDayVisibility(true);
    setAllMultiDayLabelFrequencyVisibility(true);
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  noneBtn.addEventListener("click", () => {
    setMultiDayServiceEnabled(false);
    setAllMultiDayLabelDayVisibility(false);
    setAllMultiDayLabelFrequencyVisibility(false);
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  daysAllBtn.addEventListener("click", () => {
    setAllMultiDayLabelDayVisibility(true);
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  daysNoneBtn.addEventListener("click", () => {
    setAllMultiDayLabelDayVisibility(false);
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  freqAllBtn.addEventListener("click", () => {
    setAllMultiDayLabelFrequencyVisibility(true);
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  freqNoneBtn.addEventListener("click", () => {
    setAllMultiDayLabelFrequencyVisibility(false);
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  masterToggle.addEventListener("change", () => {
    setMultiDayServiceEnabled(masterToggle.checked);
  });

  dayList.addEventListener("change", event => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    const token = String(input.dataset.dayToken || "").trim();
    if (!token) return;
    if (!setMultiDayLabelDayVisibility(token, input.checked)) return;
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  frequencyList.addEventListener("change", event => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    const key = String(input.dataset.frequencyKey || "");
    if (!key) return;
    if (!setMultiDayLabelFrequencyVisibility(key, input.checked)) return;
    renderMultiDayServiceLabels();
    refreshMultiDayManagerUi();
  });

  multiDaySidebarUiRenderer = renderSidebarUi;
  renderSidebarUi();
}

function escapeTexasLandfillText(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function pickTexasLandfillProperty(properties, candidates) {
  const source = properties && typeof properties === "object" ? properties : {};
  const normalizedEntries = Object.entries(source).map(([key, value]) => ({
    key,
    norm: String(key || "").trim().toLowerCase().replace(/[^a-z0-9]/g, ""),
    value
  }));

  const direct = normalizedEntries.find(entry => candidates.includes(entry.norm));
  if (direct) return String(direct.value ?? "").trim();

  const fuzzy = normalizedEntries.find(entry =>
    candidates.some(candidate => entry.norm.includes(candidate) || candidate.includes(entry.norm))
  );
  return fuzzy ? String(fuzzy.value ?? "").trim() : "";
}

function normalizeTexasLandfillToken(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function formatTexasLandfillDate(raw) {
  if (raw == null || raw === "") return "";
  const asNumber = Number(raw);
  if (Number.isFinite(asNumber) && asNumber > 0) {
    const date = new Date(asNumber);
    if (!Number.isNaN(date.getTime())) return date.toLocaleDateString();
  }
  const parsed = new Date(String(raw));
  if (!Number.isNaN(parsed.getTime())) return parsed.toLocaleDateString();
  return String(raw).trim();
}

function hasExplicitTexasLandfillAddress(properties) {
  return !!pickTexasLandfillProperty(properties, [
    "address",
    "address1",
    "address2",
    "streetaddress",
    "physicaladdress",
    "locationaddress",
    "siteaddress",
    "mailingaddress"
  ]);
}

function buildTexasLandfillAddressLabel(properties) {
  const line1 = pickTexasLandfillProperty(properties, [
    "address",
    "address1",
    "address2",
    "streetaddress",
    "physicaladdress",
    "locationaddress",
    "siteaddress",
    "mailingaddress"
  ]);
  const city = pickTexasLandfillProperty(properties, [
    "nearcity",
    "city",
    "cityname",
    "town"
  ]);
  const state = pickTexasLandfillProperty(properties, ["state", "st"]) || "TX";
  const zip = pickTexasLandfillProperty(properties, ["zip", "zipcode", "postalcode"]);
  const county = pickTexasLandfillProperty(properties, ["county", "countyname"]);

  if (line1) {
    const locality = [city, state, zip].filter(Boolean).join(" ");
    return [line1, locality].filter(Boolean).join(", ");
  }

  if (city || county) {
    const regionBits = [
      city,
      county ? `${county} County` : "",
      state || "TX"
    ].filter(Boolean);
    return regionBits.join(", ");
  }

  return "";
}

function isMajorTexasLandfillFeature(feature) {
  const properties = feature?.properties && typeof feature.properties === "object"
    ? feature.properties
    : {};

  const category = normalizeTexasLandfillToken(pickTexasLandfillProperty(properties, ["category"]));
  if (category && category !== "LANDFILL") return false;

  const typeCode = normalizeTexasLandfillToken(pickTexasLandfillProperty(properties, ["typecode"]));
  if (typeCode && !TEXAS_LANDFILL_MAJOR_TYPE_ALLOWLIST.has(typeCode)) return false;

  const siteStatus = normalizeTexasLandfillToken(pickTexasLandfillProperty(properties, [
    "sitestatu",
    "status",
    "operatingstatus",
    "currentstatus"
  ]));
  if (siteStatus && TEXAS_LANDFILL_EXCLUDED_SITE_STATUSES.has(siteStatus)) return false;

  const cecStatus = normalizeTexasLandfillToken(pickTexasLandfillProperty(properties, ["cecstat"]));
  if (cecStatus && TEXAS_LANDFILL_EXCLUDED_CEC_STATUSES.has(cecStatus)) return false;

  const legalStatus = normalizeTexasLandfillToken(pickTexasLandfillProperty(properties, ["legalstat"]));
  if (legalStatus && TEXAS_LANDFILL_EXCLUDED_LEGAL_STATUSES.has(legalStatus)) return false;

  return true;
}

function buildTexasLandfillPopupContent(properties, lat = null, lon = null) {
  const name = pickTexasLandfillProperty(properties, [
    "sitename",
    "facilityname",
    "landfillname",
    "name",
    "facility",
    "locationname"
  ]) || "Landfill";
  const category = pickTexasLandfillProperty(properties, ["category"]);
  const typeCode = pickTexasLandfillProperty(properties, ["typecode"]);
  const siteStatus = pickTexasLandfillProperty(properties, ["sitestatu", "status", "operatingstatus", "currentstatus"]);
  const cecStatus = pickTexasLandfillProperty(properties, ["cecstat"]);
  const legalStatus = pickTexasLandfillProperty(properties, ["legalstat"]);
  const statusDate = formatTexasLandfillDate(pickTexasLandfillProperty(properties, ["statusdat", "statusdate"]));
  const permit = pickTexasLandfillProperty(properties, ["permit", "permitnumber", "permitno", "permitid"]);
  const rn = pickTexasLandfillProperty(properties, ["rn", "regnum", "registrationnumber"]);
  const county = pickTexasLandfillProperty(properties, ["county", "countyname"]);
  const nearCity = pickTexasLandfillProperty(properties, ["nearcity", "city", "cityname", "town"]);
  const region = pickTexasLandfillProperty(properties, ["region"]);
  const tons2017 = pickTexasLandfillProperty(properties, ["f2017tons", "tons", "annualtons"]);
  const remainYards = pickTexasLandfillProperty(properties, ["remainyds", "remainingyards", "remainingcapacity"]);
  const remainYears = pickTexasLandfillProperty(properties, ["remainyea", "remainingyears", "remaininglife"]);
  const addressLabel = buildTexasLandfillAddressLabel(properties);
  const hasAddress = hasExplicitTexasLandfillAddress(properties);

  const latitude = Number.isFinite(Number(lat))
    ? Number(lat)
    : Number(pickTexasLandfillProperty(properties, ["latdd", "latitude"]));
  const longitude = Number.isFinite(Number(lon))
    ? Number(lon)
    : Number(pickTexasLandfillProperty(properties, ["longdd", "longitude", "lon"]));
  const coordinateText = Number.isFinite(latitude) && Number.isFinite(longitude)
    ? `${latitude.toFixed(5)}, ${longitude.toFixed(5)}`
    : "";

  const rows = [
    `<div><strong>${escapeTexasLandfillText(name)}</strong></div>`
  ];
  if (category) rows.push(`<div><strong>Category:</strong> ${escapeTexasLandfillText(category)}</div>`);
  if (typeCode) rows.push(`<div><strong>Type:</strong> ${escapeTexasLandfillText(typeCode)}</div>`);
  if (siteStatus) rows.push(`<div><strong>Site Status:</strong> ${escapeTexasLandfillText(siteStatus)}</div>`);
  if (cecStatus) rows.push(`<div><strong>CEC Status:</strong> ${escapeTexasLandfillText(cecStatus)}</div>`);
  if (legalStatus) rows.push(`<div><strong>Legal Status:</strong> ${escapeTexasLandfillText(legalStatus)}</div>`);
  if (statusDate) rows.push(`<div><strong>Status Date:</strong> ${escapeTexasLandfillText(statusDate)}</div>`);
  if (permit) rows.push(`<div><strong>Permit:</strong> ${escapeTexasLandfillText(permit)}</div>`);
  if (rn) rows.push(`<div><strong>RN:</strong> ${escapeTexasLandfillText(rn)}</div>`);
  if (addressLabel) {
    const addressLabelTitle = hasAddress ? "Address" : "Approx. Address";
    rows.push(`<div><strong>${addressLabelTitle}:</strong> ${escapeTexasLandfillText(addressLabel)}</div>`);
  }
  if (nearCity) rows.push(`<div><strong>Near City:</strong> ${escapeTexasLandfillText(nearCity)}</div>`);
  if (county) rows.push(`<div><strong>County:</strong> ${escapeTexasLandfillText(county)}</div>`);
  if (region) rows.push(`<div><strong>Region:</strong> ${escapeTexasLandfillText(region)}</div>`);
  if (tons2017) rows.push(`<div><strong>2017 Tons:</strong> ${escapeTexasLandfillText(tons2017)}</div>`);
  if (remainYards) rows.push(`<div><strong>Remaining Yards:</strong> ${escapeTexasLandfillText(remainYards)}</div>`);
  if (remainYears) rows.push(`<div><strong>Remaining Years:</strong> ${escapeTexasLandfillText(remainYears)}</div>`);
  if (coordinateText) rows.push(`<div><strong>Coordinates:</strong> ${escapeTexasLandfillText(coordinateText)}</div>`);

  return `<div style="font-size:12px; line-height:1.35;">${rows.join("")}</div>`;
}

function createTexasLandfillMarkerIcon(properties) {
  const typeCode = normalizeTexasLandfillToken(pickTexasLandfillProperty(properties, ["typecode"]));
  const isMunicipalStyle = typeCode.startsWith("1");
  const pinFill = isMunicipalStyle ? "#4f6c32" : "#6f5432";
  const groundFill = isMunicipalStyle ? "#7f9d56" : "#8f734f";
  const wasteFill = isMunicipalStyle ? "#5f7a45" : "#7a6343";
  const truckFill = "#2f3742";
  const truckWindowFill = "#d9e1ea";

  const html = `
    <span class="landfill-marker-wrap" aria-hidden="true">
      <svg class="landfill-marker-svg" viewBox="0 0 36 46">
        <path class="landfill-marker-pin" d="M18 1.8C10.4 1.8 4.2 8 4.2 15.6c0 10.9 10.6 19 12.8 28.8.3 1.1 1.7 1.1 2 0 2.2-9.8 12.8-17.9 12.8-28.8C31.8 8 25.6 1.8 18 1.8z" fill="${pinFill}" />
        <circle class="landfill-marker-core" cx="18" cy="15.2" r="9.2" />
        <path class="landfill-marker-ground" d="M8.4 20c2.4-3.2 5.8-4.8 9.6-4.8 3.8 0 7.2 1.6 9.6 4.8v3.9H8.4z" fill="${groundFill}" />
        <path class="landfill-marker-waste" d="M10.9 18.5l2.3-2.6 2.2 1.3 2.4-2.5 2.6 2 2.3-1.3 2.5 2.9V20h-14.3z" fill="${wasteFill}" />
        <rect class="landfill-marker-truck-bed" x="10.2" y="20.9" width="8.3" height="3.3" rx="0.8" fill="${truckFill}" />
        <rect class="landfill-marker-truck-cab" x="18.1" y="21.3" width="3.8" height="2.9" rx="0.7" fill="${truckFill}" />
        <rect class="landfill-marker-truck-window" x="19.1" y="21.7" width="1.9" height="1.2" rx="0.25" fill="${truckWindowFill}" />
        <circle class="landfill-marker-wheel" cx="12.5" cy="24.8" r="1.1" />
        <circle class="landfill-marker-wheel" cx="19.4" cy="24.8" r="1.1" />
      </svg>
    </span>
  `;

  return L.divIcon({
    className: "landfill-marker-icon",
    html,
    iconSize: [32, 42],
    iconAnchor: [16, 40],
    popupAnchor: [0, -36]
  });
}

function updateTexasLandfillsLayerStatus(message = "") {
  const statusNode = document.getElementById("texasLandfillsLayerStatus");
  if (!statusNode) return;

  if (message) {
    statusNode.textContent = String(message);
    return;
  }

  const visible = map.hasLayer(texasLandfillsLayerGroup);
  const shownCount = Number(texasLandfillsState.featureCount) || 0;
  const sourceCount = Number(texasLandfillsState.sourceFeatureCount) || shownCount;
  const filteredOutCount = Number(texasLandfillsState.filteredOutCount) || 0;
  if (texasLandfillsState.loading) {
    statusNode.textContent = "Loading Texas landfills...";
    return;
  }
  if (texasLandfillsState.lastError) {
    statusNode.textContent = `Unable to load Texas landfills: ${texasLandfillsState.lastError}`;
    return;
  }
  if (visible) {
    if (filteredOutCount > 0 && sourceCount > shownCount) {
      statusNode.textContent = `Layer is on (${shownCount.toLocaleString()} major landfill sites shown; ${filteredOutCount.toLocaleString()} filtered out).`;
    } else {
      statusNode.textContent = `Layer is on (${shownCount.toLocaleString()} landfill sites).`;
    }
    return;
  }
  if (texasLandfillsState.loaded) {
    if (filteredOutCount > 0 && sourceCount > shownCount) {
      statusNode.textContent = `Layer is off (${shownCount.toLocaleString()} major landfill sites ready; ${filteredOutCount.toLocaleString()} filtered out).`;
    } else {
      statusNode.textContent = `Layer is off (${shownCount.toLocaleString()} landfill sites ready).`;
    }
    return;
  }
  statusNode.textContent = "Layer is off.";
}

function texasLandfillFeatureToMarker(feature) {
  if (!feature || feature.type !== "Feature") return null;
  const geometry = feature.geometry || {};
  if (geometry.type !== "Point" || !Array.isArray(geometry.coordinates) || geometry.coordinates.length < 2) {
    return null;
  }

  const lon = Number(geometry.coordinates[0]);
  const lat = Number(geometry.coordinates[1]);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;

  const properties = feature.properties && typeof feature.properties === "object"
    ? feature.properties
    : {};
  const name = pickTexasLandfillProperty(properties, [
    "sitename",
    "facilityname",
    "landfillname",
    "name",
    "facility",
    "locationname"
  ]) || "Landfill";

  const marker = L.marker([lat, lon], {
    icon: createTexasLandfillMarkerIcon(properties),
    keyboard: true,
    title: name
  });
  marker.bindPopup(buildTexasLandfillPopupContent(properties, lat, lon));
  return marker;
}

function getTexasLandfillsArcGisErrorMessage(payload) {
  const error = payload?.error;
  if (!error || typeof error !== "object") return "";
  const code = Number(error.code);
  const message = String(error.message || "").trim();
  const details = Array.isArray(error.details)
    ? error.details.map(v => String(v || "").trim()).filter(Boolean)
    : [];

  const parts = [];
  if (Number.isFinite(code)) parts.push(`code ${code}`);
  if (message) parts.push(message);
  if (details.length) parts.push(details.join(" | "));
  return parts.join(": ").trim();
}

async function fetchTexasLandfillsRawJson(queryParams) {
  const params = queryParams instanceof URLSearchParams
    ? queryParams
    : new URLSearchParams(queryParams || {});
  if (!params.get("f")) params.set("f", "json");

  const response = await fetch(`${TEXAS_LANDFILLS_QUERY_URL}?${params.toString()}`, {
    headers: { Accept: "application/json" }
  });
  if (!response.ok) {
    let extra = "";
    try {
      const text = await response.text();
      const parsed = text ? JSON.parse(text) : null;
      const serviceError = getTexasLandfillsArcGisErrorMessage(parsed);
      if (serviceError) extra = `: ${serviceError}`;
    } catch (_) {
      // Ignore non-JSON response bodies.
    }
    throw new Error(`Landfill query failed (${response.status})${extra}`);
  }

  const payload = await response.json();
  const serviceError = getTexasLandfillsArcGisErrorMessage(payload);
  if (serviceError) {
    throw new Error(serviceError);
  }
  return payload;
}

async function fetchTexasLandfillsObjectIds() {
  const params = new URLSearchParams({
    where: "1=1",
    returnIdsOnly: "true",
    f: "json"
  });

  const payload = await fetchTexasLandfillsRawJson(params);
  const ids = Array.isArray(payload?.objectIds) ? payload.objectIds : [];
  return ids
    .map(v => Number(v))
    .filter(v => Number.isFinite(v))
    .sort((a, b) => a - b);
}

function convertArcGisJsonToGeoJson(payload) {
  const rawFeatures = Array.isArray(payload?.features) ? payload.features : [];
  const features = rawFeatures
    .map(item => {
      const attributes = item?.attributes && typeof item.attributes === "object" ? item.attributes : {};
      const geom = item?.geometry || {};
      const x = Number(geom?.x);
      const y = Number(geom?.y);
      if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
      return {
        type: "Feature",
        geometry: {
          type: "Point",
          coordinates: [x, y]
        },
        properties: attributes
      };
    })
    .filter(Boolean);

  return {
    type: "FeatureCollection",
    features
  };
}

async function fetchTexasLandfillFeaturesByIds(idChunk) {
  const ids = Array.isArray(idChunk) ? idChunk : [];
  if (!ids.length) {
    return { type: "FeatureCollection", features: [] };
  }

  const queryBase = {
    objectIds: ids.join(","),
    outFields: "*",
    returnGeometry: "true",
    outSR: "4326"
  };

  const geoParams = new URLSearchParams({ ...queryBase, f: "geojson" });
  const geoResponse = await fetch(`${TEXAS_LANDFILLS_QUERY_URL}?${geoParams.toString()}`, {
    headers: { Accept: "application/geo+json, application/json" }
  });
  if (geoResponse.ok) {
    const geoPayload = await geoResponse.json();
    if (Array.isArray(geoPayload?.features)) return geoPayload;
  }

  const jsonParams = new URLSearchParams({ ...queryBase, f: "json" });
  const jsonPayload = await fetchTexasLandfillsRawJson(jsonParams);
  return convertArcGisJsonToGeoJson(jsonPayload);
}

function chunkTexasLandfillIds(ids, size) {
  const input = Array.isArray(ids) ? ids : [];
  const chunkSize = Math.max(1, Number(size) || 1);
  const chunks = [];
  for (let i = 0; i < input.length; i += chunkSize) {
    chunks.push(input.slice(i, i + chunkSize));
  }
  return chunks;
}

async function fetchTexasLandfillsGeoJson() {
  let objectIds = [];
  try {
    objectIds = await fetchTexasLandfillsObjectIds();
  } catch (error) {
    console.warn("Texas landfill ID lookup fallback to direct query:", error);
  }

  if (objectIds.length) {
    const allFeatures = [];
    const chunks = chunkTexasLandfillIds(objectIds, TEXAS_LANDFILLS_FETCH_BATCH_SIZE);
    for (const chunk of chunks) {
      const collection = await fetchTexasLandfillFeaturesByIds(chunk);
      if (Array.isArray(collection?.features) && collection.features.length) {
        allFeatures.push(...collection.features);
      }
    }
    return { type: "FeatureCollection", features: allFeatures };
  }

  const jsonFallbackParams = new URLSearchParams({
    where: "1=1",
    outFields: "*",
    returnGeometry: "true",
    outSR: "4326",
    f: "json"
  });
  const jsonFallbackPayload = await fetchTexasLandfillsRawJson(jsonFallbackParams);
  const jsonFallbackCollection = convertArcGisJsonToGeoJson(jsonFallbackPayload);
  if (Array.isArray(jsonFallbackCollection?.features) && jsonFallbackCollection.features.length) {
    return jsonFallbackCollection;
  }

  const fallbackParams = new URLSearchParams({
    where: "1=1",
    outFields: "*",
    returnGeometry: "true",
    outSR: "4326",
    f: "geojson"
  });
  const fallbackResponse = await fetch(`${TEXAS_LANDFILLS_QUERY_URL}?${fallbackParams.toString()}`, {
    headers: { Accept: "application/geo+json, application/json" }
  });
  if (!fallbackResponse.ok) {
    throw new Error(`Landfill query failed (${fallbackResponse.status})`);
  }
  const fallbackPayload = await fallbackResponse.json();
  if (!Array.isArray(fallbackPayload?.features)) {
    throw new Error("Landfill query returned no features.");
  }
  return fallbackPayload;
}

function loadTexasLandfillsIntoLayerGroup(collection) {
  texasLandfillsLayerGroup.clearLayers();
  const features = Array.isArray(collection?.features) ? collection.features : [];
  texasLandfillsState.sourceFeatureCount = features.length;
  let added = 0;
  let filteredOut = 0;
  features.forEach(feature => {
    if (!isMajorTexasLandfillFeature(feature)) {
      filteredOut += 1;
      return;
    }
    const marker = texasLandfillFeatureToMarker(feature);
    if (!marker) return;
    texasLandfillsLayerGroup.addLayer(marker);
    added += 1;
  });
  texasLandfillsState.filteredOutCount = filteredOut;
  return added;
}

async function ensureTexasLandfillsLoaded() {
  if (texasLandfillsState.loaded && texasLandfillsState.featureCount > 0) return true;
  if (texasLandfillsState.loadPromise) return texasLandfillsState.loadPromise;

  texasLandfillsState.loading = true;
  texasLandfillsState.lastError = "";
  updateTexasLandfillsLayerStatus("Loading Texas landfills...");

  texasLandfillsState.loadPromise = (async () => {
    try {
      const collection = await fetchTexasLandfillsGeoJson();
      const count = loadTexasLandfillsIntoLayerGroup(collection);
      texasLandfillsState.featureCount = count;
      texasLandfillsState.loaded = count > 0;
      if (!count) {
        texasLandfillsState.lastError = texasLandfillsState.sourceFeatureCount > 0
          ? "No major/active landfill features matched the current filter."
          : "No landfill features were returned.";
        return false;
      }
      return true;
    } catch (error) {
      console.warn("Texas landfill layer load failed:", error);
      texasLandfillsLayerGroup.clearLayers();
      texasLandfillsState.loaded = false;
      texasLandfillsState.featureCount = 0;
      texasLandfillsState.sourceFeatureCount = 0;
      texasLandfillsState.filteredOutCount = 0;
      texasLandfillsState.lastError = error?.message || "Request failed.";
      return false;
    } finally {
      texasLandfillsState.loading = false;
      texasLandfillsState.loadPromise = null;
    }
  })();

  const ok = await texasLandfillsState.loadPromise;
  updateTexasLandfillsLayerStatus();
  return ok;
}

async function setTexasLandfillsLayerEnabled(enabled, options = {}) {
  const opts = options && typeof options === "object" ? options : {};
  const persist = opts.persist !== false;
  const shouldShow = !!enabled;
  const toggleNode = document.getElementById("texasLandfillsLayerEnabled");

  if (toggleNode && toggleNode.checked !== shouldShow) {
    toggleNode.checked = shouldShow;
  }
  if (persist) {
    storageSet(TEXAS_LANDFILLS_LAYER_VISIBLE_KEY, shouldShow ? "on" : "off");
  }

  if (!shouldShow) {
    map.removeLayer(texasLandfillsLayerGroup);
    updateTexasLandfillsLayerStatus();
    return;
  }

  const loaded = await ensureTexasLandfillsLoaded();
  const stillRequested = !!document.getElementById("texasLandfillsLayerEnabled")?.checked;
  if (!stillRequested) {
    map.removeLayer(texasLandfillsLayerGroup);
    updateTexasLandfillsLayerStatus();
    return;
  }

  if (!loaded) {
    if (toggleNode) toggleNode.checked = false;
    storageSet(TEXAS_LANDFILLS_LAYER_VISIBLE_KEY, "off");
    map.removeLayer(texasLandfillsLayerGroup);
    updateTexasLandfillsLayerStatus();
    return;
  }

  if (!map.hasLayer(texasLandfillsLayerGroup)) {
    texasLandfillsLayerGroup.addTo(map);
  }
  updateTexasLandfillsLayerStatus();
}

function initTexasLandfillsLayerControls() {
  const toggleHeader = document.getElementById("texasLandfillsLayerToggleHeader");
  const content = document.getElementById("texasLandfillsLayerContent");
  const masterToggle = document.getElementById("texasLandfillsLayerEnabled");
  const statusNode = document.getElementById("texasLandfillsLayerStatus");
  if (!toggleHeader || !content || !masterToggle || !statusNode) return;
  if (toggleHeader.dataset.texasLandfillsBound === "1") return;
  toggleHeader.dataset.texasLandfillsBound = "1";

  content.classList.add("collapsed");
  toggleHeader.classList.remove("open");

  toggleHeader.addEventListener("click", () => {
    const collapsed = content.classList.toggle("collapsed");
    toggleHeader.classList.toggle("open", !collapsed);
  });

  masterToggle.addEventListener("change", () => {
    setTexasLandfillsLayerEnabled(masterToggle.checked).catch(error => {
      console.warn("Unable to toggle Texas landfill layer:", error);
      updateTexasLandfillsLayerStatus("Unable to update layer visibility.");
    });
  });

  const stored = storageGet(TEXAS_LANDFILLS_LAYER_VISIBLE_KEY);
  const shouldShow = stored === "on";
  masterToggle.checked = shouldShow;
  updateTexasLandfillsLayerStatus();

  if (shouldShow) {
    setTexasLandfillsLayerEnabled(true, { persist: false }).catch(error => {
      console.warn("Unable to restore Texas landfill layer visibility:", error);
      updateTexasLandfillsLayerStatus("Unable to restore landfill layer.");
    });
  }
}

function getCityLimitsBoundsKey(bounds) {
  if (!bounds || typeof bounds.isValid !== "function" || !bounds.isValid()) return "";
  const southWest = bounds.getSouthWest();
  const northEast = bounds.getNorthEast();
  return [
    Number(southWest?.lat || 0).toFixed(3),
    Number(southWest?.lng || 0).toFixed(3),
    Number(northEast?.lat || 0).toFixed(3),
    Number(northEast?.lng || 0).toFixed(3)
  ].join("|");
}

function getCityLimitsZoomHintText() {
  const zoom = Number(map.getZoom?.()) || 0;
  return `Zoom in to level ${CITY_LIMITS_MIN_ZOOM}+ to load city limits (current: ${zoom.toFixed(1)}).`;
}

function updateCityLimitsLayerStatus(message = "") {
  const statusNode = document.getElementById("cityLimitsLayerStatus");
  if (!statusNode) return;

  if (message) {
    statusNode.textContent = String(message);
    return;
  }

  const requested = !!document.getElementById("cityLimitsLayerEnabled")?.checked;
  const visible = map.hasLayer(cityLimitsLayerGroup);
  const count = Number(cityLimitsState.featureCount) || 0;
  const zoom = Number(map.getZoom?.()) || 0;

  if (requested && zoom < CITY_LIMITS_MIN_ZOOM) {
    statusNode.textContent = getCityLimitsZoomHintText();
    return;
  }

  if (cityLimitsState.loading) {
    statusNode.textContent = "Loading city limits...";
    return;
  }
  if (cityLimitsState.lastError) {
    statusNode.textContent = `Unable to load city limits: ${cityLimitsState.lastError}`;
    return;
  }
  if (visible) {
    statusNode.textContent = count > 0
      ? `Layer is on (${count.toLocaleString()} city limits shown for this map view).`
      : "Layer is on (no city limits in this map view).";
    return;
  }
  if (cityLimitsState.loaded) {
    statusNode.textContent = count > 0
      ? `Layer is off (${count.toLocaleString()} city limits ready for this map view).`
      : "Layer is off (no city limits in this map view).";
    return;
  }
  statusNode.textContent = "Layer is off.";
}

function getCityLimitsQueryEnvelope(bounds) {
  if (!bounds || typeof bounds.isValid !== "function" || !bounds.isValid()) return null;
  const southWest = bounds.getSouthWest();
  const northEast = bounds.getNorthEast();
  return {
    xmin: Number(southWest.lng),
    ymin: Number(southWest.lat),
    xmax: Number(northEast.lng),
    ymax: Number(northEast.lat),
    spatialReference: { wkid: 4326 }
  };
}

function buildCityLimitsLayerStyle(feature) {
  const props = feature?.properties && typeof feature.properties === "object"
    ? feature.properties
    : {};
  const population = Number(props.POPULATION);
  const isMajorPlace = Number.isFinite(population) && population >= 50000;
  return {
    color: isMajorPlace ? "#f6b347" : "#6fb8ff",
    weight: isMajorPlace ? 2.0 : 1.4,
    opacity: 0.9,
    fillColor: isMajorPlace ? "#f6b347" : "#6fb8ff",
    fillOpacity: 0.06,
    dashArray: isMajorPlace ? "" : "5 4",
    pane: CITY_LIMITS_PANE,
    interactive: false
  };
}

function buildCityLimitsPointStyle(feature) {
  const props = feature?.properties && typeof feature.properties === "object"
    ? feature.properties
    : {};
  const population = Number(props.POPULATION);
  const isMajorPlace = Number.isFinite(population) && population >= 50000;
  return {
    pane: CITY_LIMITS_PANE,
    radius: isMajorPlace ? 4.8 : 3.8,
    color: isMajorPlace ? "#f6b347" : "#6fb8ff",
    weight: 1.2,
    opacity: 0.95,
    fillColor: isMajorPlace ? "#f6b347" : "#6fb8ff",
    fillOpacity: 0.75,
    interactive: false
  };
}

function getCityLimitsFeatureLabel(feature) {
  const props = feature?.properties && typeof feature.properties === "object"
    ? feature.properties
    : {};
  const value = String(props.NAME || props.CITY || props.PLACE || "").trim();
  return value || "";
}

function shouldShowCityLimitsLabels() {
  return (Number(map.getZoom?.()) || 0) >= CITY_LIMITS_LABEL_MIN_ZOOM;
}

function bindCityLimitsFeatureLabel(feature, layer) {
  if (!layer || typeof layer.bindTooltip !== "function") return;
  if (!shouldShowCityLimitsLabels()) return;
  const label = getCityLimitsFeatureLabel(feature);
  if (!label) return;

  const hasPointLocation = typeof layer.getLatLng === "function";
  layer.bindTooltip(label, {
    permanent: true,
    direction: hasPointLocation ? "top" : "center",
    offset: hasPointLocation ? [0, -8] : [0, 0],
    className: "city-limits-label",
    interactive: false,
    opacity: 0.95
  });
}

async function fetchCityLimitsGeoJsonForBounds(bounds) {
  const envelope = getCityLimitsQueryEnvelope(bounds);
  if (!envelope) {
    return { type: "FeatureCollection", features: [] };
  }
  const requestCityLimits = async params => {
    const response = await fetch(`${CITY_LIMITS_QUERY_URL}?${params.toString()}`, {
      cache: "no-store",
      headers: { Accept: "application/geo+json, application/json" }
    });
    if (!response.ok) {
      let extra = "";
      try {
        const text = await response.text();
        const parsed = text ? JSON.parse(text) : null;
        const serviceError = getTexasLandfillsArcGisErrorMessage(parsed);
        if (serviceError) extra = `: ${serviceError}`;
      } catch (_) {
        // Ignore non-JSON response bodies.
      }
      throw new Error(`City limits query failed (${response.status})${extra}`);
    }

    const payload = await response.json();
    const serviceError = getTexasLandfillsArcGisErrorMessage(payload);
    if (serviceError) throw new Error(serviceError);
    if (!Array.isArray(payload?.features)) {
      return { type: "FeatureCollection", features: [] };
    }
    return payload;
  };

  const params = new URLSearchParams({
    where: CITY_LIMITS_QUERY_WHERE,
    outFields: "*",
    returnGeometry: "true",
    outSR: "4326",
    resultRecordCount: String(CITY_LIMITS_FETCH_RECORD_LIMIT),
    orderByFields: "POPULATION DESC, NAME ASC",
    f: "geojson"
  });
  params.set("inSR", "4326");
  params.set("geometryType", "esriGeometryEnvelope");
  params.set("spatialRel", "esriSpatialRelIntersects");
  params.set("geometry", JSON.stringify(envelope));

  return requestCityLimits(params);
}

function loadCityLimitsIntoLayerGroup(collection) {
  cityLimitsLayerGroup.clearLayers();
  const features = Array.isArray(collection?.features) ? collection.features : [];
  if (!features.length) return 0;

  const layer = L.geoJSON(
    { type: "FeatureCollection", features },
    {
      style: buildCityLimitsLayerStyle,
      pointToLayer: (feature, latlng) => L.circleMarker(latlng, buildCityLimitsPointStyle(feature)),
      onEachFeature: bindCityLimitsFeatureLabel,
      pane: CITY_LIMITS_PANE,
      interactive: false
    }
  );
  cityLimitsLayerGroup.addLayer(layer);
  return features.length;
}

function scheduleCityLimitsLayerRefresh(options = {}) {
  if (!document.getElementById("cityLimitsLayerEnabled")?.checked) return;
  if (cityLimitsState.refreshTimer) clearTimeout(cityLimitsState.refreshTimer);
  const force = !!options.force;
  cityLimitsState.refreshTimer = setTimeout(() => {
    cityLimitsState.refreshTimer = null;
    refreshCityLimitsLayerForCurrentView({ force }).catch(error => {
      console.warn("Unable to refresh city limits layer:", error);
      updateCityLimitsLayerStatus("Unable to refresh city limits.");
    });
  }, CITY_LIMITS_MOVE_REFRESH_DEBOUNCE_MS);
}

function handleCityLimitsMapMoveEnd() {
  scheduleCityLimitsLayerRefresh();
}

function handleCityLimitsMapZoomEnd() {
  scheduleCityLimitsLayerRefresh({ force: true });
}

function setCityLimitsMoveListenerBound(enabled) {
  const shouldBind = !!enabled;
  if (shouldBind && !cityLimitsState.moveListenerBound) {
    map.on("moveend", handleCityLimitsMapMoveEnd);
    cityLimitsState.moveListenerBound = true;
  }
  if (shouldBind && !cityLimitsState.zoomListenerBound) {
    map.on("zoomend", handleCityLimitsMapZoomEnd);
    cityLimitsState.zoomListenerBound = true;
  }
  if (!shouldBind && cityLimitsState.moveListenerBound) {
    map.off("moveend", handleCityLimitsMapMoveEnd);
    cityLimitsState.moveListenerBound = false;
  }
  if (!shouldBind && cityLimitsState.zoomListenerBound) {
    map.off("zoomend", handleCityLimitsMapZoomEnd);
    cityLimitsState.zoomListenerBound = false;
  }
}

async function refreshCityLimitsLayerForCurrentView(options = {}) {
  const opts = options && typeof options === "object" ? options : {};
  const force = !!opts.force;
  const requested = !!document.getElementById("cityLimitsLayerEnabled")?.checked;

  if (!requested) {
    map.removeLayer(cityLimitsLayerGroup);
    updateCityLimitsLayerStatus();
    return false;
  }

  const bounds = map.getBounds?.();
  if (!bounds || typeof bounds.isValid !== "function" || !bounds.isValid()) {
    updateCityLimitsLayerStatus("Layer is on (waiting for map view).");
    return false;
  }

  const zoom = Number(map.getZoom?.()) || 0;
  if (zoom < CITY_LIMITS_MIN_ZOOM) {
    cityLimitsLayerGroup.clearLayers();
    cityLimitsState.featureCount = 0;
    cityLimitsState.loaded = false;
    cityLimitsState.lastBoundsKey = "";
    updateCityLimitsLayerStatus(getCityLimitsZoomHintText());
    return true;
  }

  const expandedBounds = bounds.pad(CITY_LIMITS_BOUNDS_PAD);
  const boundsKey = getCityLimitsBoundsKey(expandedBounds);
  if (!force && cityLimitsState.loaded && boundsKey && boundsKey === cityLimitsState.lastBoundsKey) {
    if (!map.hasLayer(cityLimitsLayerGroup)) cityLimitsLayerGroup.addTo(map);
    updateCityLimitsLayerStatus();
    return true;
  }

  if (cityLimitsState.loadPromise) {
    const pendingResult = await cityLimitsState.loadPromise;
    const latestBounds = map.getBounds?.();
    const latestKey = latestBounds ? getCityLimitsBoundsKey(latestBounds.pad(0.08)) : "";
    if (latestKey && latestKey !== cityLimitsState.lastBoundsKey) {
      scheduleCityLimitsLayerRefresh();
    }
    return pendingResult;
  }

  cityLimitsState.loading = true;
  cityLimitsState.lastError = "";
  updateCityLimitsLayerStatus("Loading city limits...");

  cityLimitsState.loadPromise = (async () => {
    try {
      const collection = await fetchCityLimitsGeoJsonForBounds(expandedBounds);
      const count = loadCityLimitsIntoLayerGroup(collection);
      cityLimitsState.featureCount = count;
      cityLimitsState.loaded = true;
      cityLimitsState.lastBoundsKey = boundsKey;
      return true;
    } catch (error) {
      cityLimitsLayerGroup.clearLayers();
      cityLimitsState.loaded = false;
      cityLimitsState.featureCount = 0;
      cityLimitsState.lastBoundsKey = "";
      cityLimitsState.lastError = error?.message || "Request failed.";
      return false;
    } finally {
      cityLimitsState.loading = false;
      cityLimitsState.loadPromise = null;
    }
  })();

  const ok = await cityLimitsState.loadPromise;
  const stillRequested = !!document.getElementById("cityLimitsLayerEnabled")?.checked;
  if (!stillRequested) {
    map.removeLayer(cityLimitsLayerGroup);
    updateCityLimitsLayerStatus();
    return false;
  }

  if (ok) {
    if (!map.hasLayer(cityLimitsLayerGroup)) {
      cityLimitsLayerGroup.addTo(map);
    }
  } else {
    map.removeLayer(cityLimitsLayerGroup);
  }
  updateCityLimitsLayerStatus();
  return ok;
}

async function setCityLimitsLayerEnabled(enabled, options = {}) {
  const opts = options && typeof options === "object" ? options : {};
  const persist = opts.persist !== false;
  const shouldShow = !!enabled;
  const toggleNode = document.getElementById("cityLimitsLayerEnabled");

  if (toggleNode && toggleNode.checked !== shouldShow) {
    toggleNode.checked = shouldShow;
  }
  if (persist) {
    storageSet(CITY_LIMITS_LAYER_VISIBLE_KEY, shouldShow ? "on" : "off");
  }

  if (!shouldShow) {
    if (cityLimitsState.refreshTimer) {
      clearTimeout(cityLimitsState.refreshTimer);
      cityLimitsState.refreshTimer = null;
    }
    setCityLimitsMoveListenerBound(false);
    map.removeLayer(cityLimitsLayerGroup);
    updateCityLimitsLayerStatus();
    return;
  }

  setCityLimitsMoveListenerBound(true);
  const loaded = await refreshCityLimitsLayerForCurrentView({ force: true });
  if (!loaded && cityLimitsState.lastError) {
    if (toggleNode) toggleNode.checked = false;
    storageSet(CITY_LIMITS_LAYER_VISIBLE_KEY, "off");
    setCityLimitsMoveListenerBound(false);
    map.removeLayer(cityLimitsLayerGroup);
  }
  updateCityLimitsLayerStatus();
}

function initCityLimitsLayerControls() {
  const toggleHeader = document.getElementById("cityLimitsLayerToggleHeader");
  const content = document.getElementById("cityLimitsLayerContent");
  const masterToggle = document.getElementById("cityLimitsLayerEnabled");
  const statusNode = document.getElementById("cityLimitsLayerStatus");
  if (!toggleHeader || !content || !masterToggle || !statusNode) return;
  if (toggleHeader.dataset.cityLimitsBound === "1") return;
  toggleHeader.dataset.cityLimitsBound = "1";

  content.classList.add("collapsed");
  toggleHeader.classList.remove("open");

  toggleHeader.addEventListener("click", () => {
    const collapsed = content.classList.toggle("collapsed");
    toggleHeader.classList.toggle("open", !collapsed);
  });

  masterToggle.addEventListener("change", () => {
    setCityLimitsLayerEnabled(masterToggle.checked).catch(error => {
      console.warn("Unable to toggle city limits layer:", error);
      updateCityLimitsLayerStatus("Unable to update layer visibility.");
    });
  });

  const stored = storageGet(CITY_LIMITS_LAYER_VISIBLE_KEY);
  const shouldShow = stored === "on";
  masterToggle.checked = shouldShow;
  updateCityLimitsLayerStatus();

  if (shouldShow) {
    setCityLimitsLayerEnabled(true, { persist: false }).catch(error => {
      console.warn("Unable to restore city limits layer visibility:", error);
      updateCityLimitsLayerStatus("Unable to restore city limits layer.");
    });
  }
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
    syncAttributeSortControls();
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
    syncAttributeSortControls();
    refreshAttributeStatus();
    return;
  }

  const sortIndicator = (key) => {
    if (attributeState.sortKey !== key) return "";
    return attributeState.sortDir > 0 ? " \u25B2" : " \u25BC";
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
      syncAttributeSortControls();
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

  syncAttributeSortControls();
  refreshAttributeStatus();
}

function exportAttributeVisibleRowsToCsv() {
  if (attributeTableMode === "streets") {
    const rows = getFilteredStreetAttributeRows();
    if (!rows.length) {
      alert("No street attributes to export.");
      return;
    }
    const headers = ["id", "name", "func_class", "highway", "ref", "maxspeed", "lanes", "surface", "oneway"];
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

function getAttributePopoutThemeMode() {
  return document.body.classList.contains("sun-mode") ? "sun" : "dark";
}

window.getStreetAttributePopoutSnapshot = function() {
  const headers = STREET_ATTRIBUTE_TABLE_HEADERS.slice();
  const rows = (Array.isArray(streetAttributesRows) && streetAttributesRows.length
    ? streetAttributesRows
    : [...streetAttributeById.values()].map(entry => entry?.row).filter(Boolean))
    .map(row => ({ rowId: Number(row?.id), values: { ...(row || {}) } }))
    .filter(item => Number.isFinite(item.rowId));
  return { headers, rows };
};

window.getRecordAttributePopoutSnapshot = function() {
  const recordRows = Array.isArray(window._currentRows) ? window._currentRows : [];
  const headers = Array.isArray(window._attributeHeaders) && window._attributeHeaders.length
    ? window._attributeHeaders.slice()
    : getAttributeHeaders(recordRows);
  const rows = recordRows
    .map(row => ({ rowId: getAttributeRowId(row), values: { ...(row || {}) } }))
    .filter(item => Number.isFinite(item.rowId));
  return { headers, rows };
};

window.promptAndAddAttributeFieldFromPopout = function() {
  setAttributeTableMode("records");
  openAttributePanel();
  promptAndAddAttributeField();
};

window.openFieldCalculatorFromPopout = function() {
  setAttributeTableMode("records");
  openAttributePanel();
  const calcBtn = document.getElementById("attributeFieldCalcBtn");
  if (calcBtn) {
    calcBtn.click();
    return true;
  }
  return false;
};

function openAttributeLikePopoutWindow(options = {}) {
  const title = String(options.title || "Attribute Table");
  const safeTitle = title
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
  const filterPlaceholder = String(options.filterPlaceholder || "Filter by any field...");
  const exportFileName = String(options.exportFileName || "attribute-table.csv");
  const snapshotGetterName = String(options.snapshotGetterName || "");
  const selectionGetterName = String(options.selectionGetterName || "");
  const selectionSetterName = String(options.selectionSetterName || "");
  const focusRowName = String(options.focusRowName || "");
  const zoomName = String(options.zoomName || "");
  const allowRecordTools = !!options.allowRecordTools;
  const addFieldActionName = String(options.addFieldActionName || "");
  const openFieldCalcActionName = String(options.openFieldCalcActionName || "");
  const themeMode = getAttributePopoutThemeMode();

  let snapshot = { headers: [], rows: [] };
  if (snapshotGetterName && typeof window[snapshotGetterName] === "function") {
    try {
      const data = window[snapshotGetterName]();
      if (data && typeof data === "object") snapshot = data;
    } catch {}
  } else {
    snapshot = {
      headers: Array.isArray(options.headers) ? options.headers.slice() : [],
      rows: Array.isArray(options.rows) ? options.rows.slice() : []
    };
  }
  const headers = Array.isArray(snapshot.headers) ? snapshot.headers : [];
  const rows = Array.isArray(snapshot.rows) ? snapshot.rows : [];
  if (!headers.length || !rows.length) {
    alert(`No ${title.toLowerCase()} data to open.`);
    return;
  }

  const initialSelectedIds = Array.isArray(options.initialSelectedIds)
    ? options.initialSelectedIds
    : [];
  const seed = JSON.stringify({
    title,
    headers,
    rows,
    selectedIds: initialSelectedIds,
    filterPlaceholder,
    exportFileName,
    snapshotGetterName,
    selectionGetterName,
    selectionSetterName,
    focusRowName,
    zoomName,
    allowRecordTools,
    addFieldActionName,
    openFieldCalcActionName,
    themeMode
  }).replace(/</g, "\\u003c");

  const win = window.open("", "_blank", "width=1180,height=740,resizable=yes,scrollbars=yes");
  if (!win) {
    alert(`Unable to open ${title.toLowerCase()} window.`);
    return;
  }

  win.document.write(`
    <html>
      <head>
        <title>${safeTitle}</title>
        <style>
          :root {
            --bg: #0f1822;
            --panel: #162433;
            --panel-soft: #1a2b3d;
            --line: #2d455c;
            --text: #e8f2fd;
            --muted: #b9cde2;
            --btn-bg: #2a4359;
            --btn-line: #5a7ca1;
            --btn-hover: #355671;
            --input-bg: #111d2a;
            --row-alt: #172636;
            --row-sel: rgba(84,176,255,0.24);
            --accent: #42a5f5;
            --warn: #7f97ad;
          }
          body.sun-mode {
            --bg: #f2f7fc;
            --panel: #ffffff;
            --panel-soft: #f5f9fe;
            --line: #cad9e8;
            --text: #18324a;
            --muted: #4f6a84;
            --btn-bg: #edf4fb;
            --btn-line: #a8c3dc;
            --btn-hover: #dfebf8;
            --input-bg: #ffffff;
            --row-alt: #f8fbff;
            --row-sel: rgba(66,165,245,0.22);
            --accent: #287fcb;
            --warn: #5b7894;
          }
          * { box-sizing: border-box; }
          body {
            margin: 0;
            padding: 12px;
            font-family: Roboto, "Segoe UI", Arial, sans-serif;
            background: var(--bg);
            color: var(--text);
          }
          .shell {
            border: 1px solid var(--line);
            border-radius: 14px;
            background: var(--panel);
            overflow: hidden;
            box-shadow: 0 12px 28px rgba(0, 0, 0, 0.3);
          }
          .head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            padding: 10px 12px;
            border-bottom: 1px solid var(--line);
            background: linear-gradient(180deg, var(--panel-soft) 0%, var(--panel) 100%);
          }
          .head-title {
            margin: 0;
            font-size: 16px;
            font-weight: 700;
          }
          .toolbar {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
            padding: 10px 12px;
            border-bottom: 1px solid var(--line);
            background: var(--panel-soft);
          }
          .toolbar button,
          .toolbar select,
          .toolbar input[type="text"] {
            height: 32px;
            border-radius: 8px;
            border: 1px solid var(--btn-line);
            background: var(--btn-bg);
            color: var(--text);
            padding: 0 10px;
            font-size: 12px;
          }
          .toolbar button {
            cursor: pointer;
            font-weight: 600;
          }
          .toolbar button:hover {
            background: var(--btn-hover);
          }
          .toolbar input[type="text"] {
            min-width: 240px;
            background: var(--input-bg);
          }
          .toolbar .selected-only {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            border: 1px solid var(--btn-line);
            border-radius: 8px;
            padding: 5px 8px;
            font-size: 12px;
            background: var(--input-bg);
          }
          .toolbar .spacer {
            flex: 1 1 auto;
          }
          .toolbar-status {
            min-width: 140px;
            text-align: right;
            font-size: 12px;
            color: var(--muted);
          }
          .table-wrap {
            height: calc(100vh - 174px);
            overflow: auto;
            padding: 10px 12px 12px;
          }
          table {
            border-collapse: collapse;
            min-width: 100%;
            width: max-content;
            background: var(--panel);
          }
          th, td {
            border: 1px solid var(--line);
            padding: 6px 8px;
            white-space: nowrap;
            font-size: 12px;
          }
          th {
            position: sticky;
            top: 0;
            background: var(--panel-soft);
            text-align: left;
          }
          th button {
            border: 0;
            background: transparent;
            color: inherit;
            font: inherit;
            font-weight: 700;
            cursor: pointer;
            padding: 0;
          }
          tbody tr { cursor: pointer; }
          tbody tr:nth-child(even) { background: var(--row-alt); }
          tbody tr.selected { background: var(--row-sel); }
          .note {
            font-size: 12px;
            color: var(--warn);
          }
          .record-tool {
            display: none;
          }
          .record-tool.show {
            display: inline-flex;
          }
          .refresh-btn {
            border-color: var(--accent) !important;
          }
        </style>
      </head>
      <body class="">
        <div class="shell">
          <div class="head">
            <h2 class="head-title" id="headTitle"></h2>
            <button type="button" id="closeBtn">Close</button>
          </div>
          <div class="toolbar">
            <button type="button" id="refreshBtn" class="refresh-btn">Refresh</button>
            <input id="searchInput" type="text" />
            <select id="sortFieldSelect" aria-label="Sort field"></select>
            <button id="sortAscBtn" type="button">Least->Greatest</button>
            <button id="sortDescBtn" type="button">Greatest->Least</button>
            <button id="clearSortBtn" type="button">Clear Sort</button>
            <button id="addFieldBtn" class="record-tool" type="button">Add Field</button>
            <button id="fieldCalcBtn" class="record-tool" type="button">Field Calculator</button>
            <label class="selected-only"><input id="selectedOnly" type="checkbox" /> Selected only</label>
            <button id="selectVisibleBtn" type="button">Select Visible</button>
            <button id="clearSelectedBtn" type="button">Clear Selection</button>
            <button id="zoomSelectedBtn" type="button">Zoom to Selected</button>
            <button id="exportCsvBtn" type="button">Export CSV</button>
            <button id="prevBtn" type="button">Prev</button>
            <button id="nextBtn" type="button">Next</button>
            <span id="pageInfo">Page 1/1</span>
            <span class="spacer"></span>
            <span id="status" class="toolbar-status">0 selected</span>
          </div>
          <div class="table-wrap">
            <table id="table"></table>
            <div id="note" class="note"></div>
          </div>
        </div>
        <script>
          const seed = ${seed};
          const state = { sortKey: "", sortDir: 1, filterText: "", selectedOnly: false, page: 1, pageSize: 300 };
          let headers = Array.isArray(seed.headers) ? seed.headers.slice() : [];
          let rows = Array.isArray(seed.rows) ? seed.rows.slice() : [];
          let selectedSet = new Set((seed.selectedIds || []).map(v => Number(v)).filter(Number.isFinite));

          document.body.className = seed.themeMode === "sun" ? "sun-mode" : "";
          const headTitle = document.getElementById("headTitle");
          const closeBtn = document.getElementById("closeBtn");
          const refreshBtn = document.getElementById("refreshBtn");
          const searchInput = document.getElementById("searchInput");
          const sortFieldSelect = document.getElementById("sortFieldSelect");
          const sortAscBtn = document.getElementById("sortAscBtn");
          const sortDescBtn = document.getElementById("sortDescBtn");
          const clearSortBtn = document.getElementById("clearSortBtn");
          const addFieldBtn = document.getElementById("addFieldBtn");
          const fieldCalcBtn = document.getElementById("fieldCalcBtn");
          const selectedOnly = document.getElementById("selectedOnly");
          const selectVisibleBtn = document.getElementById("selectVisibleBtn");
          const clearSelectedBtn = document.getElementById("clearSelectedBtn");
          const zoomSelectedBtn = document.getElementById("zoomSelectedBtn");
          const exportCsvBtn = document.getElementById("exportCsvBtn");
          const prevBtn = document.getElementById("prevBtn");
          const nextBtn = document.getElementById("nextBtn");
          const pageInfo = document.getElementById("pageInfo");
          const status = document.getElementById("status");
          const table = document.getElementById("table");
          const note = document.getElementById("note");

          headTitle.textContent = String(seed.title || "Attribute Table");
          searchInput.placeholder = String(seed.filterPlaceholder || "Filter by any field...");
          if (seed.allowRecordTools) {
            addFieldBtn.classList.add("show");
            fieldCalcBtn.classList.add("show");
          }

          function esc(v) {
            return String(v ?? "")
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;");
          }

          function callOpener(name, ...args) {
            if (!name) return undefined;
            if (!window.opener || window.opener.closed) return undefined;
            const fn = window.opener[name];
            if (typeof fn !== "function") return undefined;
            try {
              return fn(...args);
            } catch {
              return undefined;
            }
          }

          function setsEqual(a, b) {
            if (a.size !== b.size) return false;
            for (const value of a) {
              if (!b.has(value)) return false;
            }
            return true;
          }

          function pullSelectionFromOpener() {
            const ids = callOpener(seed.selectionGetterName) || [];
            const next = new Set((Array.isArray(ids) ? ids : []).map(v => Number(v)).filter(Number.isFinite));
            if (setsEqual(selectedSet, next)) return false;
            selectedSet = next;
            return true;
          }

          function pushSelectionToOpener() {
            callOpener(seed.selectionSetterName, [...selectedSet]);
          }

          function refreshDatasetFromOpener(message = "") {
            const snapshot = callOpener(seed.snapshotGetterName);
            if (!snapshot || typeof snapshot !== "object") return false;
            const nextHeaders = Array.isArray(snapshot.headers) ? snapshot.headers.slice() : [];
            const nextRows = Array.isArray(snapshot.rows) ? snapshot.rows.slice() : [];
            if (!nextHeaders.length || !nextRows.length) return false;
            headers = nextHeaders;
            rows = nextRows;
            if (state.sortKey && !headers.includes(state.sortKey)) {
              state.sortKey = "";
              state.sortDir = 1;
            }
            if (message) note.textContent = message;
            refreshSortFieldSelect();
            render();
            return true;
          }

          function refreshSortFieldSelect() {
            const previous = String(state.sortKey || "");
            sortFieldSelect.innerHTML = "";
            const placeholder = document.createElement("option");
            placeholder.value = "";
            placeholder.textContent = "Sort field...";
            sortFieldSelect.appendChild(placeholder);
            headers.forEach(field => {
              const opt = document.createElement("option");
              opt.value = String(field);
              opt.textContent = String(field);
              sortFieldSelect.appendChild(opt);
            });
            if ([...sortFieldSelect.options].some(option => option.value === previous)) {
              sortFieldSelect.value = previous;
            } else {
              sortFieldSelect.value = "";
            }
          }

          function getFilteredRows() {
            const needle = String(state.filterText || "").trim().toLowerCase();
            let data = rows.slice();
            if (needle) {
              data = data.filter(item =>
                headers.some(h => String(item?.values?.[h] ?? "").toLowerCase().includes(needle))
              );
            }
            if (state.selectedOnly) {
              data = data.filter(item => selectedSet.has(Number(item?.rowId)));
            }
            if (state.sortKey) {
              const key = state.sortKey;
              const dir = state.sortDir > 0 ? 1 : -1;
              data.sort((a, b) => {
                const av = a?.values?.[key];
                const bv = b?.values?.[key];
                const an = Number(av);
                const bn = Number(bv);
                if (Number.isFinite(an) && Number.isFinite(bn)) return (an - bn) * dir;
                return String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true, sensitivity: "base" }) * dir;
              });
            }
            return data;
          }

          function exportCsv(items) {
            if (!items.length) return;
            const csvEsc = value => {
              const raw = String(value ?? "");
              return /[",\\n]/.test(raw) ? ('"' + raw.replace(/"/g, '""') + '"') : raw;
            };
            const lines = [headers.map(csvEsc).join(",")];
            items.forEach(item => {
              lines.push(headers.map(h => csvEsc(item?.values?.[h] ?? "")).join(","));
            });
            const blob = new Blob([lines.join("\\n")], { type: "text/csv;charset=utf-8;" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = String(seed.exportFileName || "attribute-table.csv");
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

            const sortMarker = key => state.sortKey === key ? (state.sortDir > 0 ? " ^" : " v") : "";
            let html = "<thead><tr><th>Sel</th><th>#</th>";
            headers.forEach(h => {
              html += '<th><button type="button" data-sort="' + esc(h) + '">' + esc(h) + sortMarker(h) + "</button></th>";
            });
            html += "</tr></thead><tbody>";
            pageRows.forEach((item, index) => {
              const rowId = Number(item?.rowId);
              const checked = selectedSet.has(rowId) ? " checked" : "";
              html += '<tr data-row-id="' + rowId + '" class="' + (checked ? "selected" : "") + '">';
              html += '<td><input type="checkbox" data-row-select="' + rowId + '"' + checked + "></td>";
              html += "<td>" + (start + index + 1) + "</td>";
              headers.forEach(h => {
                html += "<td>" + esc(item?.values?.[h] ?? "") + "</td>";
              });
              html += "</tr>";
            });
            html += "</tbody>";
            table.innerHTML = html;

            pageInfo.textContent = "Page " + state.page + "/" + totalPages;
            status.textContent = selectedSet.size + " selected • " + filtered.length + " visible";
            prevBtn.disabled = state.page <= 1;
            nextBtn.disabled = state.page >= totalPages;
            sortFieldSelect.value = state.sortKey || "";

            table.querySelectorAll("button[data-sort]").forEach(btn => {
              btn.addEventListener("click", () => {
                const key = String(btn.getAttribute("data-sort") || "");
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
              tr.addEventListener("click", event => {
                if (event.target.closest("input")) return;
                const rowId = Number(tr.getAttribute("data-row-id"));
                if (!Number.isFinite(rowId)) return;
                callOpener(seed.focusRowName, rowId);
              });
            });
          }

          closeBtn.addEventListener("click", () => window.close());
          refreshBtn.addEventListener("click", () => {
            const ok = refreshDatasetFromOpener("Data refreshed from main window.");
            if (!ok) note.textContent = "Refresh did not return updated data.";
          });
          searchInput.addEventListener("input", () => {
            state.filterText = searchInput.value || "";
            state.page = 1;
            render();
          });
          sortFieldSelect.addEventListener("change", () => {
            state.sortKey = String(sortFieldSelect.value || "");
            state.sortDir = 1;
            state.page = 1;
            render();
          });
          sortAscBtn.addEventListener("click", () => {
            if (!state.sortKey) state.sortKey = String(sortFieldSelect.value || "");
            if (!state.sortKey) return;
            state.sortDir = 1;
            state.page = 1;
            render();
          });
          sortDescBtn.addEventListener("click", () => {
            if (!state.sortKey) state.sortKey = String(sortFieldSelect.value || "");
            if (!state.sortKey) return;
            state.sortDir = -1;
            state.page = 1;
            render();
          });
          clearSortBtn.addEventListener("click", () => {
            state.sortKey = "";
            state.sortDir = 1;
            state.page = 1;
            sortFieldSelect.value = "";
            render();
          });
          selectedOnly.addEventListener("change", () => {
            pullSelectionFromOpener();
            state.selectedOnly = !!selectedOnly.checked;
            state.page = 1;
            render();
          });
          selectVisibleBtn.addEventListener("click", () => {
            getFilteredRows().forEach(item => {
              const rowId = Number(item?.rowId);
              if (Number.isFinite(rowId)) selectedSet.add(rowId);
            });
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
            const ids = [...selectedSet];
            if (!ids.length) return;
            const ok = callOpener(seed.zoomName, ids);
            if (!ok && ids.length) callOpener(seed.focusRowName, ids[0]);
          });
          exportCsvBtn.addEventListener("click", () => exportCsv(getFilteredRows()));
          prevBtn.addEventListener("click", () => {
            state.page = Math.max(1, state.page - 1);
            render();
          });
          nextBtn.addEventListener("click", () => {
            state.page += 1;
            render();
          });

          addFieldBtn.addEventListener("click", () => {
            if (!seed.allowRecordTools) return;
            callOpener(seed.addFieldActionName);
            setTimeout(() => {
              refreshDatasetFromOpener("Field list refreshed after Add Field.");
            }, 150);
          });
          fieldCalcBtn.addEventListener("click", () => {
            if (!seed.allowRecordTools) return;
            const opened = callOpener(seed.openFieldCalcActionName);
            if (opened === false) note.textContent = "Unable to open Field Calculator in main window.";
            else note.textContent = "Field Calculator opened in main window.";
          });

          window.addEventListener("focus", () => {
            const changed = pullSelectionFromOpener();
            if (changed) render();
          });
          setInterval(() => {
            const changed = pullSelectionFromOpener();
            if (changed) render();
          }, 600);

          refreshSortFieldSelect();
          pullSelectionFromOpener();
          render();
        <\/script>
      </body>
    </html>
  `);
  win.document.close();
}

function openStreetAttributeTablePopout() {
  const snapshot = window.getStreetAttributePopoutSnapshot();
  if (!snapshot.rows.length) {
    alert("No street attributes to open.");
    return;
  }
  openAttributeLikePopoutWindow({
    title: "Street Attributes",
    headers: snapshot.headers,
    rows: snapshot.rows,
    initialSelectedIds: [...streetAttributeSelectedIds],
    snapshotGetterName: "getStreetAttributePopoutSnapshot",
    selectionGetterName: "getStreetSelectedSegmentIds",
    selectionSetterName: "setStreetSelectedSegmentIds",
    focusRowName: "focusStreetSegmentOnMap",
    zoomName: "zoomToStreetSegmentsOnMap",
    exportFileName: "street-attributes.csv",
    filterPlaceholder: "Filter streets by any field...",
    allowRecordTools: false
  });
}

function openAttributeTablePopout() {
  if (attributeTableMode === "streets") {
    openStreetAttributeTablePopout();
    return;
  }
  const snapshot = window.getRecordAttributePopoutSnapshot();
  if (!snapshot.headers.length || !snapshot.rows.length) {
    alert("No attribute data to open.");
    return;
  }
  openAttributeLikePopoutWindow({
    title: "Attribute Table",
    headers: snapshot.headers,
    rows: snapshot.rows,
    initialSelectedIds: [...attributeState.selectedRowIds],
    snapshotGetterName: "getRecordAttributePopoutSnapshot",
    selectionGetterName: "getAttributeSelectedRowIds",
    selectionSetterName: "setAttributeSelectedRowIds",
    focusRowName: "focusAttributeRowOnMap",
    zoomName: "zoomToAttributeRowsOnMap",
    exportFileName: "attribute-table.csv",
    filterPlaceholder: "Filter records by any field...",
    allowRecordTools: true,
    addFieldActionName: "promptAndAddAttributeFieldFromPopout",
    openFieldCalcActionName: "openFieldCalculatorFromPopout"
  });
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
  const canonicalFieldsSynced = ensureCanonicalRecordFields(rows);
  const serviceTimeFieldAdded = ensureServiceTimeField(rows);
  const backupColumnsAdded = ensureExistBackupColumns(rows);
  syncSetSegmentMatchStatusFields(rows);
  const schemaAdjustedByTdsPak = canonicalFieldsSynced || serviceTimeFieldAdded || backupColumnsAdded;

  // store globally for saving later
  window._currentRows = rows;
  window._currentWorkbook = wb;
  window._attributeHeaders = getAttributeHeaders(rows);
  attributeState.sortKey = null;
  attributeState.sortDir = 1;
  attributeState.page = 1;
  attributeState.selectedRowIds.clear();
  syncSelectedStopsHeaderCount(0);
  attributeRowToMarker = new WeakMap();
  attributeMarkerByRowId.clear();
  rows.forEach((row, rowIndex) => attributeRowToId.set(row, rowIndex));

  // Clear previous map data
  clearSequenceLayerData({ silent: true });
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

    const route = getMappedRowRoute(row) || "Unassigned";
    const day = getMappedRowDayRaw(row) || "No Day";

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

const marker = createMarker(lat, lon, symbol)
  .addTo(map);
marker._routeToken = route;
marker._dayToken = day;
marker._multiDayLatToken = row?.LATITUDE;
marker._multiDayLonToken = row?.LONGITUDE;
syncRecordPopupContent(marker);
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


    // ðŸ”¥ CRITICAL: link marker to Excel row
    marker._rowRef = row;
    marker._rowId = rowIndex;
    attributeRowToMarker.set(row, marker);
    attributeMarkerByRowId.set(rowIndex, marker);
    routeDayGroups[key].layers.push(marker);
    routeSet.add(route);
    globalBounds.extend([lat, lon]);
  });

  // Default ordering for a newly loaded route file: records (route+day) above streets.
  resetLayerManagerOrderToDefault();
  if (typeof autoExpandRouteDaySidebarIfLayersPresent === "function") {
    autoExpandRouteDaySidebarIfLayersPresent();
  } else {
    const routeDayToggleNode = document.getElementById("routeDayToggle");
    const routeDayContentNode = document.getElementById("routeDayContent");
    if (routeDayToggleNode && routeDayContentNode && Object.keys(routeDayGroups).length) {
      routeDayContentNode.classList.remove("collapsed");
      routeDayToggleNode.classList.add("open");
    }
  }

  buildRouteCheckboxes([...routeSet]);
  buildRouteDayLayerControls();
  rebuildMultiDayServiceProfilesFromMapLayers();
  applyFilters();
  renderAttributeTable();
  refreshAttributeStatus();

  if (globalBounds.isValid()) map.fitBounds(globalBounds);
  if (schemaAdjustedByTdsPak) {
    syncCurrentWorkbookSheetFromRows();
    queueWorkbookCloudSave("Schema Sync");
  }
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

async function openRouteFileFromCloud(routeName, options = {}) {
  const name = String(routeName || "").trim();
  if (!name) return false;

  const loadingMessage = String(options.loadingMessage || "Loading Excel file...");
  const successMessage = String(options.successMessage || "File loaded successfully.");
  const errorMessage = String(options.errorMessage || "Error loading file.");
  const closeManager = options.closeManager !== false;
  const loadSummary = options.loadSummary !== false;

  try {
    showLoading(loadingMessage);

    const { data } = sb.storage.from(BUCKET).getPublicUrl(name);
    const urlWithBypass = data.publicUrl + "?v=" + Date.now();
    const response = await fetch(urlWithBypass, { cache: "no-store" });
    if (!response.ok) throw new Error(`File fetch failed (${response.status}).`);

    window._currentFilePath = name;
    setCurrentFileDisplay(window._currentFilePath);
    processExcelBuffer(await response.arrayBuffer());

    if (loadSummary) {
      await loadSummaryFor(name);
    }

    hideLoading(successMessage);
    if (closeManager && fileManagerModal) {
      fileManagerModal.style.display = "none";
    }
    return true;
  } catch (error) {
    console.error(error);
    hideLoading();
    alert(errorMessage);
    return false;
  }
}

async function routeFileExistsInCloud(routeName) {
  const name = String(routeName || "").trim();
  if (!name) return false;
  const { data, error } = await sb.storage.from(BUCKET).list();
  if (error) throw error;
  return (data || []).some(file => String(file?.name || "") === name);
}

async function loadSummaryFileFromCloud(summaryName) {
  const name = String(summaryName || "").trim();
  if (!name) return false;

  const { data: urlData } = sb.storage.from(BUCKET).getPublicUrl(name);
  const summaryUrl = String(urlData?.publicUrl || "");
  if (!summaryUrl) throw new Error("Summary URL could not be resolved.");

  const separator = summaryUrl.includes("?") ? "&" : "?";
  const r = await fetch(`${summaryUrl}${separator}v=${Date.now()}`, { cache: "no-store" });
  if (!r.ok) throw new Error(`Summary fetch failed (${r.status}).`);

  const wb = XLSX.read(new Uint8Array(await r.arrayBuffer()), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];

  // Read entire sheet as grid
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

  // ===== FIND FIRST NON-EMPTY ROW =====
  const startRow = raw.findIndex(row =>
    row && row.some(cell => String(cell || "").trim() !== "")
  );

  if (startRow === -1) {
    showRouteSummary([], []);
    return true;
  }

  // ===== DETECT MULTI-ROW HEADERS (supports 1-3+) =====
  const headerRows = [raw[startRow]];
  const nextRow = raw[startRow + 1];
  const thirdRow = raw[startRow + 2];

  const looksLikeHeader = row => {
    if (!row) return false;
    const filled = row.filter(c => String(c || "").trim() !== "").length;
    const numeric = row.filter(c => !isNaN(parseFloat(c))).length;
    return filled > 0 && numeric < filled / 2;
  };

  if (looksLikeHeader(nextRow)) headerRows.push(nextRow);
  if (looksLikeHeader(thirdRow)) headerRows.push(thirdRow);

  // ===== BUILD SAFE COLUMN NAMES =====
  const columnCount = Math.max(...headerRows.map(row => row.length));
  const headers = Array.from({ length: columnCount }, (_, col) => {
    const parts = headerRows
      .map(row => String(row[col] || "").trim())
      .filter(Boolean);
    return parts.join(" ") || `Column ${col + 1}`;
  });

  // ===== DATA STARTS AFTER HEADER =====
  const dataStartIndex = startRow + headerRows.length;

  // ===== BUILD ROW OBJECTS =====
  const rows = raw.slice(dataStartIndex).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row?.[i] ?? "";
    });
    return obj;
  });

  showRouteSummary(rows, headers);
  autoCollapseSidebarsForSummary();

  // Force the panel open when a summary exists (desktop).
  const panel = document.getElementById("bottomSummary");
  const btn = document.getElementById("summaryToggleBtn");
  const isMobile = window.innerWidth <= 900;
  if (panel && btn && !isMobile) {
    panel.classList.remove("collapsed");
    btn.textContent = SUMMARY_ARROW_EXPANDED;
  }

  return true;
}

async function listFiles() {
  const { data, error } = await sb.storage.from(BUCKET).list();
  if (error) return console.error(error);

  const ul = document.getElementById("savedFiles");
  if (!ul) return;

  try {
    await ensureSummaryAttachmentsInitialized();
  } catch (error) {
    console.warn("Summary attachment sync init failed before rendering saved files:", error);
  }

  ul.innerHTML = "";
  const cloudFiles = Array.isArray(data) ? data : [];
  const routeFiles = cloudFiles.filter(file => !isRouteSummaryFileName(file.name) && !isSystemCloudFileName(file.name));
  const summaryFiles = cloudFiles.filter(file => isRouteSummaryFileName(file.name) && !isSystemCloudFileName(file.name));
  const allFileNames = cloudFiles.map(f => f.name);
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

  const summaryToRoutes = new Map();
  routeEntries.forEach(({ routeName, summaryName }) => {
    if (!summaryName) return;
    if (!summaryToRoutes.has(summaryName)) summaryToRoutes.set(summaryName, []);
    summaryToRoutes.get(summaryName).push(routeName);
  });

  const summaryEntries = summaryFiles
    .map(file => {
      const summaryName = file.name;
      const linkedRoutes = (summaryToRoutes.get(summaryName) || [])
        .slice()
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
      const addedTimestamp = getSavedFileAddedTimestamp(file);
      return { summaryName, linkedRoutes, addedTimestamp };
    })
    .sort((a, b) => {
      if (activeSortMode === "date-desc") {
        const delta = b.addedTimestamp - a.addedTimestamp;
        if (delta !== 0) return delta;
        return a.summaryName.localeCompare(b.summaryName, undefined, { numeric: true, sensitivity: "base" });
      }
      if (activeSortMode === "date-asc") {
        const delta = a.addedTimestamp - b.addedTimestamp;
        if (delta !== 0) return delta;
        return a.summaryName.localeCompare(b.summaryName, undefined, { numeric: true, sensitivity: "base" });
      }
      const aLinked = a.linkedRoutes.length > 0;
      const bLinked = b.linkedRoutes.length > 0;
      if (aLinked !== bLinked) return bLinked ? 1 : -1;
      return a.summaryName.localeCompare(b.summaryName, undefined, { numeric: true, sensitivity: "base" });
    });

  const routeFileNames = routeEntries.map(entry => entry.routeName);
  const appendSectionTitle = title => {
    const heading = document.createElement("div");
    heading.className = "saved-files-section-title";
    heading.textContent = title;
    ul.appendChild(heading);
  };

  if (!routeEntries.length && !summaryEntries.length) {
    const empty = document.createElement("div");
    empty.className = "saved-files-empty";
    empty.textContent = "No saved route or summary files found.";
    ul.appendChild(empty);
    return;
  }

  if (routeEntries.length) {
    appendSectionTitle("Route Files");
  }

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

    const routeTypeBadge = document.createElement("span");
    routeTypeBadge.className = "saved-file-type-badge route";
    routeTypeBadge.textContent = "Route File";
    metaRow.appendChild(routeTypeBadge);

    const badge = document.createElement("span");
    badge.className = `saved-summary-badge${summaryName ? "" : " missing"}`;
    badge.textContent = summaryName ? "Summary Attached" : "No Summary";
    metaRow.appendChild(badge);

    if (summaryName) {
      const summaryFileNode = document.createElement("span");
      summaryFileNode.className = "saved-summary-file";
      summaryFileNode.textContent = `Route Summary: ${summaryName}`;
      metaRow.appendChild(summaryFileNode);
    }

    infoWrap.append(nameNode, metaRow);

    const actions = document.createElement("div");
    actions.className = "saved-route-actions";

    const openBtn = document.createElement("button");
    openBtn.className = "saved-file-btn open-btn";
    openBtn.textContent = "Open Map";
    openBtn.onclick = async () => {
      await openRouteFileFromCloud(routeName, {
        loadingMessage: "Loading Excel file...",
        successMessage: "File loaded successfully.",
        errorMessage: "Error loading file.",
        closeManager: true,
        loadSummary: true
      });
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
        alert("\u274C Incorrect password. File not deleted.");
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

      alert("\u2705 File deleted successfully.");
      listFiles();
    };
    actions.appendChild(delBtn);

    li.append(infoWrap, actions);
    ul.appendChild(li);
  });

  if (summaryEntries.length) {
    appendSectionTitle("Route Summary Files");
  }

  summaryEntries.forEach(({ summaryName, linkedRoutes }) => {
    const li = document.createElement("li");
    li.className = `saved-route-item saved-summary-item${linkedRoutes.length ? " has-summary" : ""}`;

    const infoWrap = document.createElement("div");
    infoWrap.className = "saved-route-info";

    const nameNode = document.createElement("div");
    nameNode.className = "saved-route-name";
    nameNode.textContent = summaryName;

    const metaRow = document.createElement("div");
    metaRow.className = "saved-route-meta";

    const summaryTypeBadge = document.createElement("span");
    summaryTypeBadge.className = "saved-file-type-badge summary";
    summaryTypeBadge.textContent = "Route Summary";
    metaRow.appendChild(summaryTypeBadge);

    const badge = document.createElement("span");
    badge.className = `saved-summary-badge${linkedRoutes.length ? "" : " missing"}`;
    badge.textContent = linkedRoutes.length ? "Attached" : "Not Attached";
    metaRow.appendChild(badge);

    if (linkedRoutes.length) {
      const attachedRouteNode = document.createElement("span");
      attachedRouteNode.className = "saved-summary-file";
      attachedRouteNode.textContent = linkedRoutes.length === 1
        ? `Attached to Route: ${linkedRoutes[0]}`
        : `Attached to Routes: ${linkedRoutes.slice(0, 2).join(", ")}${linkedRoutes.length > 2 ? ` (+${linkedRoutes.length - 2} more)` : ""}`;
      metaRow.appendChild(attachedRouteNode);
    }

    infoWrap.append(nameNode, metaRow);

    const actions = document.createElement("div");
    actions.className = "saved-route-actions";

    const openSummaryBtn = document.createElement("button");
    openSummaryBtn.className = "saved-file-btn summary-btn";
    openSummaryBtn.textContent = "Open Summary";
    openSummaryBtn.onclick = async () => {
      try {
        await loadSummaryFileFromCloud(summaryName);
        if (fileManagerModal) fileManagerModal.style.display = "none";
      } catch (summaryError) {
        console.error("SUMMARY OPEN ERROR:", summaryError);
        alert("Error loading summary.");
      }
    };
    actions.appendChild(openSummaryBtn);

    const attachBtn = document.createElement("button");
    attachBtn.className = "saved-file-btn open-btn";
    attachBtn.textContent = "Attach";
    attachBtn.onclick = async () => {
      if (!routeFileNames.length) {
        alert("No route files are saved yet. Upload a route file first.");
        return;
      }
      const selectedRoute = await openSummaryAttachModal(summaryName, routeFileNames);
      if (!selectedRoute) return;
      setRouteSummaryAttachment(selectedRoute, summaryName);
      alert(`Attached ${summaryName} to ${selectedRoute}.`);
      listFiles();
    };
    actions.appendChild(attachBtn);

    const delBtn = document.createElement("button");
    delBtn.className = "saved-file-btn delete-btn";
    delBtn.textContent = "Delete";
    delBtn.onclick = async () => {
      const entered = prompt("Enter password to delete this file:");
      if (entered !== DELETE_PASSWORD) {
        alert("\u274C Incorrect password. File not deleted.");
        return;
      }

      const confirmed = confirm("Are you sure you want to permanently delete this file?");
      if (!confirmed) return;

      await sb.storage.from(BUCKET).remove([summaryName]);

      const map = getSummaryAttachments();
      Object.keys(map).forEach(routeKey => {
        if (map[routeKey] === summaryName) delete map[routeKey];
      });
      setSummaryAttachments(map);

      alert("\u2705 File deleted successfully.");
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
    let existingInCloud = false;
    try {
      existingInCloud = await routeFileExistsInCloud(file.name);
    } catch (listError) {
      console.warn("Unable to verify existing cloud file before upload:", listError);
    }

    if (existingInCloud) {
      const openSavedCopy = window.confirm(
        `"${file.name}" is already in Saved Files.\n\n` +
        "OK: Open saved cloud copy (keeps your previous TDS-Pak edits).\n" +
        "Cancel: Replace cloud file with this local file."
      );
      if (openSavedCopy) {
        await openRouteFileFromCloud(file.name, {
          loadingMessage: "Opening saved file...",
          successMessage: "Saved file opened.",
          errorMessage: "Could not open saved file.",
          closeManager: false,
          loadSummary: true
        });
        return;
      }
    }

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

    hideLoading("Upload Complete \u2705");

  } catch (error) {
    console.error("UPLOAD ERROR:", error);
    hideLoading();
    alert("Upload failed: " + error.message);
  }
}

async function uploadRouteSummaryAndAttach(file) {
  if (!file) return;

  try {
    await ensureSummaryAttachmentsInitialized();
    showLoading("Uploading route summary...");

    const { error } = await sb.storage
      .from(BUCKET)
      .upload(file.name, file, { upsert: true });

    if (error) throw error;

    const { data: files, error: listErr } = await sb.storage.from(BUCKET).list();
    if (listErr) throw listErr;

    const routeFileNames = files
      .filter(f => !isRouteSummaryFileName(f.name) && !isSystemCloudFileName(f.name))
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

  const normalizeSummaryHeader = value =>
    String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");

  const normalizedHeaders = headers.map(h => ({
    original: h,
    norm: normalizeSummaryHeader(h)
  }));

  const findHeaderByAliases = aliases => {
    const aliasNorms = aliases.map(normalizeSummaryHeader).filter(Boolean);
    if (!aliasNorms.length) return null;

    const direct = normalizedHeaders.find(h => aliasNorms.includes(h.norm));
    if (direct) return direct.original;

    const fuzzy = normalizedHeaders.find(h =>
      aliasNorms.some(alias => h.norm.includes(alias) || alias.includes(h.norm))
    );
    return fuzzy ? fuzzy.original : null;
  };

  const routeHeader = findHeaderByAliases([
    "route",
    "newroute",
    "routeid",
    "rte",
    "routecode"
  ]);

  const dayHeader = findHeaderByAliases([
    "day",
    "newday",
    "dispatchday",
    "serviceday",
    "weekday"
  ]);

  const routeDayHeader = findHeaderByAliases([
    "routeplusday",
    "routeandday",
    "route_day",
    "routebyday",
    "routewithday",
    "route day"
  ]);

  const parseRouteDayCombinedValue = rawValue => {
    const raw = String(rawValue ?? "").trim();
    if (!raw) return null;

    const separators = ["|", "/", "\\", " - ", "-", ":"];
    for (const separator of separators) {
      if (!raw.includes(separator)) continue;
      const parts = raw
        .split(separator)
        .map(part => String(part).trim())
        .filter(Boolean);
      if (parts.length === 2) {
        return { route: parts[0], day: parts[1] };
      }
    }

    const trailingDayMatch = raw.match(/^(.*?)\s+(mon(?:day)?|tue(?:s(?:day)?)?|wed(?:nesday)?|thu(?:r(?:s(?:day)?)?)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?|[1-7])$/i);
    if (trailingDayMatch) {
      return {
        route: String(trailingDayMatch[1] || "").trim(),
        day: String(trailingDayMatch[2] || "").trim()
      };
    }

    return null;
  };

  const getRouteDayFromSummaryRow = row => {
    let routeValue = routeHeader ? String(row?.[routeHeader] ?? "").trim() : "";
    let dayValue = dayHeader ? String(row?.[dayHeader] ?? "").trim() : "";

    if ((!routeValue || !dayValue) && routeDayHeader) {
      const parsed = parseRouteDayCombinedValue(row?.[routeDayHeader]);
      if (parsed) {
        if (!routeValue) routeValue = parsed.route;
        if (!dayValue) dayValue = parsed.day;
      }
    }

    return { routeValue, dayValue };
  };

  const clearActiveSummaryRow = () => {
    tbody.querySelectorAll("tr.summary-route-day-row.active").forEach(rowNode => {
      rowNode.classList.remove("active");
      rowNode.setAttribute("aria-selected", "false");
    });
  };

  const activateSummaryRow = rowNode => {
    clearActiveSummaryRow();
    rowNode.classList.add("active");
    rowNode.setAttribute("aria-selected", "true");
  };

  let summaryRowHoverHint = document.getElementById("summaryRowHoverHint");
  if (!summaryRowHoverHint) {
    summaryRowHoverHint = document.createElement("div");
    summaryRowHoverHint.id = "summaryRowHoverHint";
    summaryRowHoverHint.className = "summary-row-hover-hint";
    summaryRowHoverHint.textContent = "Click to highlight and select this Route + Day";
    document.body.appendChild(summaryRowHoverHint);
  }

  let summaryRowHoverHintTimer = null;
  const hideSummaryRowHoverHint = (delayMs = 0) => {
    if (summaryRowHoverHintTimer) {
      clearTimeout(summaryRowHoverHintTimer);
      summaryRowHoverHintTimer = null;
    }
    const runHide = () => summaryRowHoverHint.classList.remove("show");
    if (delayMs > 0) {
      summaryRowHoverHintTimer = setTimeout(runHide, delayMs);
      return;
    }
    runHide();
  };

  const showSummaryRowHoverHint = (event, rowNode) => {
    if (!summaryRowHoverHint || !rowNode) return;
    if (summaryRowHoverHintTimer) {
      clearTimeout(summaryRowHoverHintTimer);
      summaryRowHoverHintTimer = null;
    }

    const rowRect = rowNode.getBoundingClientRect();
    const anchorX = Number.isFinite(event?.clientX)
      ? event.clientX
      : rowRect.left + Math.min(rowRect.width * 0.6, 220);
    const anchorY = Number.isFinite(event?.clientY)
      ? event.clientY
      : rowRect.top + (rowRect.height / 2);

    let left = anchorX + 14;
    let top = anchorY - 20;

    summaryRowHoverHint.style.left = `${Math.round(left)}px`;
    summaryRowHoverHint.style.top = `${Math.round(top)}px`;
    summaryRowHoverHint.classList.add("show");

    const hintRect = summaryRowHoverHint.getBoundingClientRect();
    if (hintRect.right > window.innerWidth - 10) {
      left = Math.max(10, anchorX - hintRect.width - 14);
    }
    if (hintRect.bottom > window.innerHeight - 10) {
      top = Math.max(10, window.innerHeight - hintRect.height - 10);
    }
    if (hintRect.top < 10) {
      top = 10;
    }

    summaryRowHoverHint.style.left = `${Math.round(left)}px`;
    summaryRowHoverHint.style.top = `${Math.round(top)}px`;
    hideSummaryRowHoverHint(1200);
  };

  rows.forEach(r => {
    const tr = document.createElement("tr");
    headers.forEach((h, colIndex) => {
      const td = document.createElement("td");
      td.textContent = r[h] ?? "";
      columnCells[colIndex].push(td);
      tr.appendChild(td);
    });

    const { routeValue, dayValue } = getRouteDayFromSummaryRow(r);
    if (routeValue && dayValue) {
      tr.classList.add("summary-route-day-row");
      tr.setAttribute("tabindex", "0");
      tr.setAttribute("role", "button");
      tr.setAttribute("aria-selected", "false");
      tr.dataset.route = routeValue;
      tr.dataset.day = dayValue;
      tr.setAttribute("aria-label", `Click to highlight ${routeValue} ${dayValue} on map`);

      const onActivate = () => {
        hideSummaryRowHoverHint();
        const result = window.highlightRouteDayOnMap(routeValue, dayValue, {
          selectRecords: true,
          fitBounds: true
        });
        if (result?.ok) {
          activateSummaryRow(tr);
        }
      };

      tr.addEventListener("click", onActivate);
      tr.addEventListener("keydown", event => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          onActivate();
        }
      });
      tr.addEventListener("mouseenter", event => {
        showSummaryRowHoverHint(event, tr);
      });
      tr.addEventListener("mouseleave", () => {
        hideSummaryRowHoverHint(180);
      });
    }

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
  btn.textContent = SUMMARY_ARROW_EXPANDED;
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
  if (mobileMenuBtn) mobileMenuBtn.textContent = "\u2630";

  // Right sidebar: collapse desktop and hide mobile panel.
  if (selectionBox) {
    selectionBox.classList.add("collapsed");
    selectionBox.classList.remove("show");
  }
  if (toggleSelectionBtn) toggleSelectionBtn.textContent = RIGHT_SIDEBAR_ARROW_COLLAPSED;

  setTimeout(() => map.invalidateSize(), 180);
}



// Load matching summary file
async function loadSummaryFor(routeFileName) {
  const { data, error } = await sb.storage.from(BUCKET).list();
  if (error) {
    console.error("LIST ERROR:", error);
    return;
  }

  try {
    await ensureSummaryAttachmentsInitialized();
  } catch (syncError) {
    console.warn("Summary attachment sync init failed before loading summary:", syncError);
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

// ===== DETECT MULTI-ROW HEADERS (supports 1â€“3+) =====
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


// ðŸ”½ FORCE the panel open when a summary exists
const panel = document.getElementById("bottomSummary");
const btn = document.getElementById("summaryToggleBtn");

const isMobile = window.innerWidth <= 900;

if (panel && btn && !isMobile) {
  panel.classList.remove("collapsed");
  btn.textContent = SUMMARY_ARROW_EXPANDED;
}



}



// ================= START APP =================
// ===== TOGGLE BOTTOM SUMMARY =====
function toggleSummary() {
  const panel = document.getElementById("bottomSummary");
  const btn = document.getElementById("summaryToggleBtn");

  panel.classList.toggle("collapsed");

  // flip arrow direction
  btn.textContent = panel.classList.contains("collapsed")
    ? SUMMARY_ARROW_COLLAPSED
    : SUMMARY_ARROW_EXPANDED;
}
// ===== PLACE LOCATE BUTTON BASED ON SCREEN SIZE =====
function placeLocateButton() {
  const locateBtn = document.getElementById("locateMeBtn");
    const headerContainer = document.querySelector(".mobile-header-buttons");
  const desktopContainer = document.getElementById("desktopLocateContainer");
    const streetToggle = document.getElementById("streetLabelToggle");

  if (!locateBtn || !headerContainer || !desktopContainer) return;

  if (window.innerWidth <= 900) {
    // ðŸ“± MOBILE
    headerContainer.appendChild(locateBtn);

    if (streetToggle) {
      headerContainer.appendChild(streetToggle.parentElement);
    }

  } else {
    // ðŸ–¥ DESKTOP
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
const selectionTargetSelect = document.getElementById("selectionTargetSelect");
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
if (toggleSelectionBtn) toggleSelectionBtn.textContent = RIGHT_SIDEBAR_ARROW_COLLAPSED;

// Toggle sidebar open/closed
if (selectionBox && toggleSelectionBtn) {
  toggleSelectionBtn.onclick = () => {
    const collapsed = selectionBox.classList.toggle("collapsed");
    toggleSelectionBtn.textContent = collapsed
      ? RIGHT_SIDEBAR_ARROW_COLLAPSED
      : RIGHT_SIDEBAR_ARROW_EXPANDED;
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
  const targetLabel = getMapSelectionTargetLabel(mapSelectionTarget);
  if (selectionDrawModeSelect) selectionDrawModeSelect.value = mode;
  if (selectionTargetSelect) selectionTargetSelect.value = normalizeMapSelectionTarget(mapSelectionTarget);
  if (selectionModeBadge) selectionModeBadge.textContent = modeLabel;
  setSelectionToolsStatus(`Target: ${targetLabel}. Mode: ${modeLabel}.`);
};

window.__refreshSelectionToolsUi = refreshSelectionToolsUi;
refreshSelectionToolsUi();

selectionDrawModeSelect?.addEventListener("change", () => {
  setMapSelectionDrawMode(selectionDrawModeSelect.value, true);
  refreshSelectionToolsUi();
});

selectionTargetSelect?.addEventListener("change", () => {
  setMapSelectionTarget(selectionTargetSelect.value, true);
  refreshSelectionToolsUi();
});

function getSelectionCandidatesForScope(scope) {
  const resolvedTarget = resolveMapSelectionTarget();
  const useInView = scope === "inView";
  if (resolvedTarget === "both") {
    return useInView
      ? {
          records: getRecordRowIdsInCurrentView(),
          streets: getStreetIdsInCurrentView()
        }
      : {
          records: getVisibleSelectableRecordRowIds(),
          streets: getVisibleSelectableStreetIds()
        };
  }
  if (resolvedTarget === "streets") {
    return useInView ? getStreetIdsInCurrentView() : getVisibleSelectableStreetIds();
  }
  return useInView ? getRecordRowIdsInCurrentView() : getVisibleSelectableRecordRowIds();
}

selectionDrawPolygonBtn?.addEventListener("click", () => {
  const started = startSelectionDrawTool("polygon");
  if (!started) {
    setSelectionToolsStatus("Unable to start polygon tool.");
    return;
  }
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  const targetLabel = getMapSelectionTargetLabel(mapSelectionTarget);
  setSelectionToolsStatus(`Polygon draw started (${modeLabel}, ${targetLabel}).`);
});

selectionDrawRectangleBtn?.addEventListener("click", () => {
  const started = startSelectionDrawTool("rectangle");
  if (!started) {
    setSelectionToolsStatus("Unable to start rectangle tool.");
    return;
  }
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  const targetLabel = getMapSelectionTargetLabel(mapSelectionTarget);
  setSelectionToolsStatus(`Rectangle draw started (${modeLabel}, ${targetLabel}).`);
});

selectionSelectInViewBtn?.addEventListener("click", () => {
  const ids = getSelectionCandidatesForScope("inView");
  applySelectionIdsToActiveMode(ids, mapSelectionDrawMode, mapSelectionTarget);
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  const targetLabel = getMapSelectionTargetLabel(mapSelectionTarget);
  setSelectionToolsStatus(`Select In View applied (${modeLabel}, ${targetLabel}). ${getSelectionSummaryForTarget(mapSelectionTarget)}.`);
});

selectionSelectVisibleBtn?.addEventListener("click", () => {
  const ids = getSelectionCandidatesForScope("visible");
  applySelectionIdsToActiveMode(ids, mapSelectionDrawMode, mapSelectionTarget);
  const modeLabel = getMapSelectionModeLabel(mapSelectionDrawMode);
  const targetLabel = getMapSelectionTargetLabel(mapSelectionTarget);
  setSelectionToolsStatus(`Select All Visible applied (${modeLabel}, ${targetLabel}). ${getSelectionSummaryForTarget(mapSelectionTarget)}.`);
});

selectionInvertVisibleBtn?.addEventListener("click", () => {
  invertVisibleSelectionInActiveMode(mapSelectionTarget);
  const targetLabel = getMapSelectionTargetLabel(mapSelectionTarget);
  setSelectionToolsStatus(`Invert Visible applied (${targetLabel}). ${getSelectionSummaryForTarget(mapSelectionTarget)}.`);
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

    // ðŸ”¥ Force counter refresh everywhere (desktop + mobile)
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
const attributeSortFieldSelect = document.getElementById("attributeSortFieldSelect");
const attributeSortAscBtn = document.getElementById("attributeSortAscBtn");
const attributeSortDescBtn = document.getElementById("attributeSortDescBtn");
const attributeClearSortBtn = document.getElementById("attributeClearSortBtn");
const attributeAddFieldBtn = document.getElementById("attributeAddFieldBtn");
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
attributeSortFieldSelect?.addEventListener("change", () => {
  const key = String(attributeSortFieldSelect.value || "").trim();
  if (!key) {
    attributeState.sortKey = null;
    attributeState.sortDir = 1;
    attributeState.page = 1;
    renderAttributeTable();
    return;
  }
  attributeState.sortKey = key;
  attributeState.sortDir = 1;
  attributeState.page = 1;
  renderAttributeTable();
});
attributeSortAscBtn?.addEventListener("click", () => applyAttributeSortFromToolbar(1));
attributeSortDescBtn?.addEventListener("click", () => applyAttributeSortFromToolbar(-1));
attributeClearSortBtn?.addEventListener("click", () => {
  attributeState.sortKey = null;
  attributeState.sortDir = 1;
  attributeState.page = 1;
  renderAttributeTable();
});
attributeAddFieldBtn?.addEventListener("click", promptAndAddAttributeField);
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
initMultiDayManagerControls();
initServiceDayLabelLayerControls();
initTexasLandfillsLayerControls();
initCityLimitsLayerControls();
initSequenceLayerControls();

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
initFieldCalculatorControls();
initUpdateFieldControls();
initPrintCenterControls();
initLayerManagerControls();
initFacilitiesControls();
syncRouteSequencerFacilitiesLayer();
initRouteSequencerControls();
initSharedFacilitiesCloudSync().catch(error => {
  console.warn("Unable to initialize shared facilities cloud sync:", error);
});
initSetSegmentIdSideControls();



  
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

    mobileMenuBtn.textContent = open ? "\u2715" : "\u2630";
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
    mobileMenuBtn.textContent = "\u2630";
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
    if (toggleBtn) toggleBtn.textContent = SUMMARY_ARROW_COLLAPSED;

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

      // If user drags upward from collapsed state, immediately reveal summary content.
      if (panel.classList.contains("collapsed") && newHeight > minHeight + 2) {
        panel.classList.remove("collapsed");
        if (toggleBtn) toggleBtn.textContent = SUMMARY_ARROW_EXPANDED;
      }
    });

    document.addEventListener("pointerup", () => {
      if (!isDragging) return;
      isDragging = false;
      document.body.style.userSelect = "";
      const minHeight = 40;
      if (panel.offsetHeight <= minHeight + 2) {
        panel.classList.add("collapsed");
        panel.style.height = "40px";
        if (toggleBtn) toggleBtn.textContent = SUMMARY_ARROW_COLLAPSED;
      } else {
        panel.classList.remove("collapsed");
        if (toggleBtn) toggleBtn.textContent = SUMMARY_ARROW_EXPANDED;
        storageSet("summaryHeight", panel.offsetHeight);
      }

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
  toggleBtn.textContent = SUMMARY_ARROW_COLLAPSED;
} else {
  let restored = storageGet("summaryHeight");

  if (!restored || restored <= 60) {
    restored = window.innerWidth <= 900 ? 300 : 250;
  }

  panel.style.height = restored + "px";
  toggleBtn.textContent = SUMMARY_ARROW_EXPANDED;
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
      const summaryPopoutTheme = document.body.classList.contains("sun-mode") ? "sun" : "dark";
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
              tr.summary-route-day-row { cursor: pointer; }
              tr.summary-route-day-row:focus { outline: none; }
              tr.summary-route-day-row:focus td {
                box-shadow: inset 0 0 0 1px rgba(105, 188, 255, 0.82);
              }
              tr.summary-route-day-row.active td {
                background: rgba(255, 213, 74, 0.27);
                box-shadow: inset 0 0 0 1px rgba(236, 168, 32, 0.58);
              }
              .summary-row-hover-hint {
                position: fixed;
                z-index: 2200;
                pointer-events: none;
                max-width: min(320px, calc(100vw - 20px));
                padding: 8px 12px;
                border-radius: 11px;
                border: 1px solid rgba(96, 135, 171, 0.72);
                background: linear-gradient(160deg, rgba(25, 40, 56, 0.96) 0%, rgba(19, 31, 44, 0.96) 100%);
                color: #e9f5ff;
                font-size: 12px;
                font-weight: 700;
                letter-spacing: 0.01em;
                box-shadow: 0 12px 28px rgba(8, 16, 24, 0.48);
                backdrop-filter: blur(6px);
                opacity: 0;
                transform: translateY(6px) scale(0.98);
                transition: opacity 0.18s ease, transform 0.18s ease;
              }
              .summary-row-hover-hint.show {
                opacity: 1;
                transform: translateY(0) scale(1);
              }
              th:first-child, td:first-child { border-left: 1px solid var(--line); }
              body.dark-mode {
                --bg: #0f1822;
                --panel: #142232;
                --line: #2c4459;
                --head: #1b2d3e;
                --text: #e4effa;
                --muted: #9ab6d3;
                --accent: #42a5f5;
              }
              body.dark-mode {
                background: radial-gradient(circle at 10% 0%, #1a2c3f 0%, var(--bg) 42%);
              }
              body.dark-mode .summary-shell {
                box-shadow: 0 14px 30px rgba(0, 0, 0, 0.35);
              }
              body.dark-mode .summary-header {
                background: linear-gradient(180deg, #21374b 0%, #1a2d3f 100%);
              }
              body.dark-mode .summary-back-btn {
                border-color: #4f6f8f;
                background: #1b3044;
                color: #e4effa;
              }
              body.dark-mode .summary-back-btn:hover {
                background: #26405a;
              }
              body.dark-mode tr:nth-child(even) td {
                background: #172838;
              }
              body.dark-mode tr:hover td {
                background: #20384d;
              }
            </style>
          </head>
          <body class="${summaryPopoutTheme === "sun" ? "sun-mode" : "dark-mode"}">
            <div class="summary-shell">
              <div class="summary-header">
                <button class="summary-back-btn" onclick="returnToMap()">&#8592; Back to Map</button>
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

              (function initSummaryRowInteractions() {
                const rowCandidates = Array.from(document.querySelectorAll("tbody tr"));
                if (!rowCandidates.length) return;

                const normalizeHeader = value =>
                  String(value || "")
                    .trim()
                    .toLowerCase()
                    .replace(/[^a-z0-9]/g, "");

                const headerNodes = Array.from(document.querySelectorAll("thead th"));
                const headerNorms = headerNodes.map(th => normalizeHeader(th.textContent));

                const findIndexByAliases = aliases => {
                  const aliasNorms = aliases.map(normalizeHeader).filter(Boolean);
                  if (!aliasNorms.length) return -1;
                  const direct = headerNorms.findIndex(norm => aliasNorms.includes(norm));
                  if (direct >= 0) return direct;
                  return headerNorms.findIndex(norm =>
                    aliasNorms.some(alias => norm.includes(alias) || alias.includes(norm))
                  );
                };

                const routeIndex = findIndexByAliases(["route", "newroute", "routeid", "rte", "routecode"]);
                const dayIndex = findIndexByAliases(["day", "newday", "dispatchday", "serviceday", "weekday"]);
                const routeDayIndex = findIndexByAliases(["routeplusday", "routeandday", "route_day", "routebyday", "routewithday", "route day"]);

                const parseCombinedRouteDay = value => {
                  const raw = String(value || "").trim();
                  if (!raw) return null;
                  const separators = ["|", "/", "\\\\", " - ", "-", ":"];
                  for (const separator of separators) {
                    if (!raw.includes(separator)) continue;
                    const parts = raw.split(separator).map(part => String(part).trim()).filter(Boolean);
                    if (parts.length === 2) {
                      return { route: parts[0], day: parts[1] };
                    }
                  }
                  const trailingDay = raw.match(/^(.*?)\\s+(mon(?:day)?|tue(?:s(?:day)?)?|wed(?:nesday)?|thu(?:r(?:s(?:day)?)?)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?|[1-7])$/i);
                  if (trailingDay) {
                    return {
                      route: String(trailingDay[1] || "").trim(),
                      day: String(trailingDay[2] || "").trim()
                    };
                  }
                  return null;
                };

                const resolveRowRouteDay = row => {
                  let route = String(row.dataset.route || "").trim();
                  let day = String(row.dataset.day || "").trim();
                  const cells = row.cells || [];

                  if (!route && routeIndex >= 0 && cells[routeIndex]) {
                    route = String(cells[routeIndex].textContent || "").trim();
                  }
                  if (!day && dayIndex >= 0 && cells[dayIndex]) {
                    day = String(cells[dayIndex].textContent || "").trim();
                  }
                  if ((!route || !day) && routeDayIndex >= 0 && cells[routeDayIndex]) {
                    const parsed = parseCombinedRouteDay(cells[routeDayIndex].textContent);
                    if (parsed) {
                      if (!route) route = parsed.route;
                      if (!day) day = parsed.day;
                    }
                  }
                  return { route, day };
                };

                const rows = rowCandidates
                  .map(row => {
                    const { route, day } = resolveRowRouteDay(row);
                    if (!route || !day) return null;
                    row.classList.add("summary-route-day-row");
                    row.dataset.route = route;
                    row.dataset.day = day;
                    return row;
                  })
                  .filter(Boolean);
                if (!rows.length) return;

                const clearActiveRows = () => {
                  rows.forEach(row => {
                    row.classList.remove("active");
                    row.setAttribute("aria-selected", "false");
                  });
                };

                const setActiveRow = row => {
                  clearActiveRows();
                  row.classList.add("active");
                  row.setAttribute("aria-selected", "true");
                };

                const hint = document.createElement("div");
                hint.className = "summary-row-hover-hint";
                hint.textContent = "Click to highlight and select this Route + Day";
                document.body.appendChild(hint);

                let hintTimer = null;
                const hideHint = (delayMs = 0) => {
                  if (hintTimer) {
                    clearTimeout(hintTimer);
                    hintTimer = null;
                  }
                  const run = () => hint.classList.remove("show");
                  if (delayMs > 0) {
                    hintTimer = setTimeout(run, delayMs);
                    return;
                  }
                  run();
                };

                const showHint = (event, row) => {
                  if (hintTimer) {
                    clearTimeout(hintTimer);
                    hintTimer = null;
                  }
                  const rowRect = row.getBoundingClientRect();
                  const anchorX = Number.isFinite(event?.clientX)
                    ? event.clientX
                    : rowRect.left + Math.min(rowRect.width * 0.6, 220);
                  const anchorY = Number.isFinite(event?.clientY)
                    ? event.clientY
                    : rowRect.top + (rowRect.height / 2);

                  let left = anchorX + 14;
                  let top = anchorY - 20;
                  hint.style.left = Math.round(left) + "px";
                  hint.style.top = Math.round(top) + "px";
                  hint.classList.add("show");

                  const hintRect = hint.getBoundingClientRect();
                  if (hintRect.right > window.innerWidth - 10) {
                    left = Math.max(10, anchorX - hintRect.width - 14);
                  }
                  if (hintRect.bottom > window.innerHeight - 10) {
                    top = Math.max(10, window.innerHeight - hintRect.height - 10);
                  }
                  if (hintRect.top < 10) {
                    top = 10;
                  }

                  hint.style.left = Math.round(left) + "px";
                  hint.style.top = Math.round(top) + "px";
                  hideHint(1200);
                };

                const activateRow = row => {
                  hideHint();
                  const { route, day } = resolveRowRouteDay(row);
                  if (!route || !day) return;

                  const mapWindow = window.opener && !window.opener.closed ? window.opener : null;
                  if (!mapWindow || typeof mapWindow.highlightRouteDayOnMap !== "function") return;

                  let result = null;
                  try {
                    result = mapWindow.highlightRouteDayOnMap(route, day, {
                      selectRecords: true,
                      fitBounds: true
                    });
                  } catch (_) {
                    return;
                  }
                  if (result && result.ok) setActiveRow(row);
                };

                rows.forEach(row => {
                  row.setAttribute("tabindex", row.getAttribute("tabindex") || "0");
                  row.setAttribute("role", "button");
                  row.setAttribute("aria-selected", row.classList.contains("active") ? "true" : "false");

                  row.addEventListener("click", () => activateRow(row));
                  row.addEventListener("keydown", event => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      activateRow(row);
                    }
                  });
                  row.addEventListener("mouseenter", event => showHint(event, row));
                  row.addEventListener("mouseleave", () => hideHint(180));
                });
              })();
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
        { key: "demand", label: "Demand (Tons)", aliases: ["demand", "totaldemand", "volume", "load", "tons", "tonnage", "totaltons"] },
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
      const demandUnitLabel = "Tons";

      function makeBars(items, labelKey, valueKey, alt, kind, options = {}) {
        const config = options && typeof options === "object" ? options : {};
        const emptyMessage = String(config.emptyMessage || "No data available.");
        const requirePositive = !!config.requirePositive;
        const rows = Array.isArray(items) ? items : [];

        if (!rows.length) return `<div class="empty">${escapeHtml(emptyMessage)}</div>`;

        const values = rows.map(i => Number(i?.[valueKey]) || 0);
        if (requirePositive && !values.some(v => v > 0)) {
          return `<div class="empty">${escapeHtml(emptyMessage)}</div>`;
        }

        const max = Math.max(1, ...values);
        return rows
          .slice()
          .sort((a, b) => (Number(b?.[valueKey]) || 0) - (Number(a?.[valueKey]) || 0))
          .map(i => `
            <div class="bar-row section-route-row" data-kind="${escapeHtml(kind || "")}" data-key="${encodeURIComponent(String(i[labelKey] ?? ""))}">
              <div class="bar-label">${escapeHtml(i[labelKey])}</div>
              <div class="bar-track ${alt ? "alt" : ""}"><div class="bar-fill ${alt ? "alt" : ""}" style="width:${((Number(i?.[valueKey]) || 0) / max) * 100}%"></div></div>
              <div class="bar-value">${(Number(i?.[valueKey]) || 0).toLocaleString(undefined, { maximumFractionDigits: 2 })}</div>
            </div>
          `)
          .join("");
      }

      const demandAliases = [
        "demand",
        "totaldemand",
        "volume",
        "load",
        "tons",
        "tonnage",
        "totaltons",
        "tonsweek"
      ];

      function resolveMapDemandField(sourceRows) {
        const aliasNorms = demandAliases.map(alias => normalize(alias)).filter(Boolean);
        const scores = new Map();
        (sourceRows || []).slice(0, 120).forEach(row => {
          Object.keys(row || {}).forEach(rawKey => {
            const norm = normalize(rawKey);
            if (!norm) return;
            let score = 0;
            aliasNorms.forEach(aliasNorm => {
              if (norm === aliasNorm) score = Math.max(score, 4);
              else if (norm.includes(aliasNorm) || aliasNorm.includes(norm)) score = Math.max(score, 2);
            });
            if (!score) return;
            scores.set(rawKey, (scores.get(rawKey) || 0) + score);
          });
        });

        let bestKey = "";
        let bestScore = 0;
        scores.forEach((score, key) => {
          if (score > bestScore) {
            bestScore = score;
            bestKey = key;
          }
        });
        return bestKey || "";
      }

      const mapRows = Array.isArray(window._currentRows) ? window._currentRows : [];
      const mapDemandField = resolveMapDemandField(mapRows);

      function readDemandTonsFromRow(row) {
        if (!row || typeof row !== "object") return 0;

        if (mapDemandField && Object.prototype.hasOwnProperty.call(row, mapDemandField)) {
          const directValue = toNumber(row[mapDemandField]);
          if (Number.isFinite(directValue)) return directValue;
        }

        const fallbackEntry = Object.entries(row).find(([rawKey]) => {
          const norm = normalize(rawKey);
          if (!norm) return false;
          return demandAliases.some(alias => {
            const aliasNorm = normalize(alias);
            return norm === aliasNorm || norm.includes(aliasNorm) || aliasNorm.includes(norm);
          });
        });
        if (!fallbackEntry) return 0;
        const parsed = toNumber(fallbackEntry[1]);
        return Number.isFinite(parsed) ? parsed : 0;
      }

      const frequencyDemandMap = new Map();
      if (multiDayServiceByLocation && multiDayServiceByLocation.size) {
        multiDayServiceByLocation.forEach(profile => {
          const tokenSet = new Set(
            Array.isArray(profile?.dayTokens) && profile.dayTokens.length
              ? profile.dayTokens
              : (profile?.days || [])
          );
          const summary = buildMultiDayServiceSummary(tokenSet);
          const key = String(
            profile?.frequencyKey ||
            buildMultiDayFrequencyKey(summary.tokens || collectSortedMultiDayDayTokens(tokenSet)) ||
            JSON.stringify(collectSortedMultiDayDayTokens(tokenSet))
          );
          const label = String(profile?.detailLabel || summary.detail || summary.short || "No Day").trim() || "No Day";
          const shortLabel = String(profile?.shortLabel || summary.short || label).trim() || label;
          const layers = Array.isArray(profile?.layers) ? profile.layers : [];

          const frequencyEntry = frequencyDemandMap.get(key) || {
            key,
            label,
            shortLabel,
            demand: 0,
            recordCount: 0,
            locationCount: 0,
            routeDaySet: new Set(),
            routeNames: new Set()
          };

          let profileDemand = 0;
          layers.forEach(layer => {
            profileDemand += readDemandTonsFromRow(layer?._rowRef);
            frequencyEntry.recordCount += 1;

            const routeToken = String(layer?._routeToken ?? "").trim() || "Unassigned";
            const rawDay = String(layer?._dayToken ?? "").trim();
            const dayToken = normalizeDayToken(rawDay) || rawDay || "No Day";
            const dayLabel = dayName(Number(dayToken)) || dayToken || "No Day";
            frequencyEntry.routeNames.add(routeToken);
            frequencyEntry.routeDaySet.add(`${routeToken} | ${dayLabel}`);
          });

          frequencyEntry.demand += profileDemand;
          frequencyEntry.locationCount += 1;
          frequencyDemandMap.set(key, frequencyEntry);
        });
      }

      const frequencyDemandRows = Array.from(frequencyDemandMap.values())
        .map(entry => {
          const routeDayList = Array.from(entry.routeDaySet.values()).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
          const routeList = Array.from(entry.routeNames.values()).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
          const avgDemandPerStop = entry.recordCount > 0 ? (entry.demand / entry.recordCount) : 0;
          return {
            key: entry.key,
            label: entry.label,
            shortLabel: entry.shortLabel,
            demand: entry.demand,
            avgDemandPerStop,
            demandSharePct: 0,
            recordCount: entry.recordCount,
            locationCount: entry.locationCount,
            routeDayCount: routeDayList.length,
            routeDayList,
            routeList
          };
        })
        .sort((a, b) => b.demand - a.demand);

      const totalFrequencyDemand = frequencyDemandRows.reduce((sum, row) => sum + row.demand, 0);
      frequencyDemandRows.forEach(row => {
        row.demandSharePct = totalFrequencyDemand > 0
          ? (row.demand / totalFrequencyDemand) * 100
          : 0;
      });

      const dayDemandBars = makeBars(dayTotals, "day", "demand", false, "day", {
        requirePositive: true,
        emptyMessage: `No demand (${demandUnitLabel}) values were found for day totals.`
      });
      const demandByFrequencyBars = makeBars(frequencyDemandRows, "label", "demand", false, "frequency", {
        requirePositive: true,
        emptyMessage: `No service-frequency demand (${demandUnitLabel}) values are available.`
      });
      const demandPerStopFrequencyBars = makeBars(frequencyDemandRows, "label", "avgDemandPerStop", true, "frequency", {
        requirePositive: true,
        emptyMessage: `No service-frequency stop and demand values are available to compute average demand per stop.`
      });
      const demandShareFrequencyBars = makeBars(frequencyDemandRows, "label", "demandSharePct", true, "frequency", {
        requirePositive: true,
        emptyMessage: "No service-frequency demand share values are available."
      });
      const dayMilesBars = makeBars(dayTotals, "day", "miles", true, "day", {
        requirePositive: true,
        emptyMessage: "No miles values were found for day totals."
      });
      const dayStopsBars = makeBars(dayTotals, "day", "stops", false, "day", {
        requirePositive: true,
        emptyMessage: "No stops values were found for day totals."
      });
      const dayTripsBars = makeBars(dayTotals, "day", "trips", true, "day", {
        requirePositive: true,
        emptyMessage: "No trips values were found for day totals."
      });
      const hasFrequencyDemandRows = frequencyDemandRows.length > 0;
      const frequencySourceNote = hasFrequencyDemandRows
        ? `Service-frequency totals are derived from the Service Day Label layer frequency groups in the uploaded route file, and demand values are treated as ${demandUnitLabel}.`
        : `No service-day frequency groups are available yet. Load route records first so the Service Day Label layer can be built.`;

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
          demand: Number(r.demand) || 0,
          totalTime: Number(r.totalTime) || 0
        }))
      ).replace(/</g, "\\u003c");
      const frequencyDemandDataJson = JSON.stringify(
        frequencyDemandRows.map(r => ({
          key: r.key,
          label: r.label,
          shortLabel: r.shortLabel,
          demand: Number(r.demand) || 0,
          avgDemandPerStop: Number(r.avgDemandPerStop) || 0,
          demandSharePct: Number(r.demandSharePct) || 0,
          recordCount: Number(r.recordCount) || 0,
          locationCount: Number(r.locationCount) || 0,
          routeDayCount: Number(r.routeDayCount) || 0,
          routeDayList: Array.isArray(r.routeDayList) ? r.routeDayList : [],
          routeList: Array.isArray(r.routeList) ? r.routeList : []
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
      const summaryVizPopoutTheme = document.body.classList.contains("sun-mode") ? "sun" : "dark";
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
              .metric-bars-panel.hidden { display:none; }
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
              body.dark-mode {
                --bg:#0f1822;
                --panel:#142232;
                --line:#2c4459;
                --text:#e4effa;
                --muted:#9ab6d3;
                --a:#42a5f5;
                --b:#2bbfa9;
              }
              body.dark-mode {
                background: radial-gradient(circle at 10% 0%, #1a2c3f 0%, var(--bg) 45%);
              }
              body.dark-mode .shell {
                box-shadow:0 14px 30px rgba(0,0,0,.34);
              }
              body.dark-mode .head {
                background:linear-gradient(180deg,#21374b 0%,#1a2d3f 100%);
              }
              body.dark-mode .summary-back-btn {
                border-color:#4f6f8f;
                background:#1b3044;
                color:var(--text);
              }
              body.dark-mode .summary-back-btn:hover { background:#26405a; }
              body.dark-mode .card { background:#1a2d3f; }
              body.dark-mode .section-route-row:hover { background:#21384d; }
              body.dark-mode .bar-track { background:#1f3447; }
              body.dark-mode .bar-track.alt { background:#1a3e39; }
              body.dark-mode .chart-wrap,
              body.dark-mode .table-wrap,
              body.dark-mode .section-modal-card { background:#162a3b; }
              body.dark-mode .chart-controls select,
              body.dark-mode .chart-controls input,
              body.dark-mode .section-modal-close {
                background:#1b3044;
                color:var(--text);
                border-color:#4f6f8f;
              }
              body.dark-mode .section-modal-head { background:#1e3346; }
              body.dark-mode th { background:#1f3347; }
              body.dark-mode tr:nth-child(even) td { background:#172a3b; }
              body.dark-mode tr:hover td { background:#20384d; }
            </style>
          </head>
          <body class="${summaryVizPopoutTheme === "sun" ? "sun-mode" : "dark-mode"}">
            <div class="shell">
              <div class="head">
                <div class="head-left">
                  <button class="summary-back-btn" onclick="returnToMap()">&#8592; Back to Map</button>
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
                <h3>Distribution by Day</h3>
                <div class="chart-controls">
                  <label>Metric
                    <select id="dayDistributionMetric">
                      <option value="demand">Demand (${demandUnitLabel})</option>
                      <option value="miles">Miles</option>
                      <option value="stops">Stops</option>
                    </select>
                  </label>
                </div>
                <p id="dayDistributionMetricNote" class="note">Demand in this dashboard is interpreted as ${demandUnitLabel}.</p>
                <div id="dayDistributionDemandPanel" class="metric-bars-panel">
                  ${dayDemandBars}
                </div>
                <div id="dayDistributionMilesPanel" class="metric-bars-panel hidden">
                  ${dayMilesBars}
                </div>
                <div id="dayDistributionStopsPanel" class="metric-bars-panel hidden">
                  ${dayStopsBars}
                </div>
              </div>
              <div class="section">
                <h3>Distribution of Trips Per Day</h3>
                ${dayTripsBars}
              </div>
              <div class="section">
                <h3>Route + Day Detail</h3>
                <p class="note">Each row below is grouped by route and day to make day-level performance easier to interpret.</p>
                ${missingFields.length ? `<p class="note">Missing in this file: ${escapeHtml(missingFields.join(", "))}</p>` : ""}
                <p class="note">Totals: Stops ${totalStops.toLocaleString(undefined, { maximumFractionDigits: 2 })}, Demand (${demandUnitLabel}) ${totalDemand.toLocaleString(undefined, { maximumFractionDigits: 2 })}, Miles ${totalMiles.toLocaleString(undefined, { maximumFractionDigits: 2 })}</p>
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
                        <th>Demand (${demandUnitLabel})</th>
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
                const frequencyRows = ${frequencyDemandDataJson};
                const host = document.getElementById("stopsMilesScatterHost");
                const tooltip = document.getElementById("scatterTooltip");
                const chartToast = document.getElementById("chartToast");
                const dayFilter = document.getElementById("scatterDayFilter");
                const overOnly = document.getElementById("scatterOverOnly");
                const targetHours = ${targetHours};
                const demandUnitLabel = ${JSON.stringify(demandUnitLabel)};
                const dayDistributionMetric = document.getElementById("dayDistributionMetric");
                const dayDistributionMetricNote = document.getElementById("dayDistributionMetricNote");
                const dayDistributionPanels = {
                  demand: document.getElementById("dayDistributionDemandPanel"),
                  miles: document.getElementById("dayDistributionMilesPanel"),
                  stops: document.getElementById("dayDistributionStopsPanel")
                };
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
                      ", Demand: " + Number(r.demand || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) + " " + demandUnitLabel +
                      ", Time: " + Number(r.totalTime || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) + "h</li>"
                    ).join("");
                  }
                  modal.classList.remove("hidden");
                }

                function openSimpleListModal(title, lines) {
                  modalTitle.textContent = title;
                  const items = Array.isArray(lines) ? lines : [];
                  if (!items.length) {
                    modalList.innerHTML = "<li>No entries found.</li>";
                  } else {
                    modalList.innerHTML = items.map(line => "<li>" + escapeHtmlLocal(line) + "</li>").join("");
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

                function updateDayDistributionMetricView() {
                  const metric = (dayDistributionMetric && dayDistributionMetric.value) || "demand";
                  Object.entries(dayDistributionPanels).forEach(([key, panel]) => {
                    if (!panel) return;
                    panel.classList.toggle("hidden", key !== metric);
                  });

                  if (!dayDistributionMetricNote) return;
                  if (metric === "demand") {
                    dayDistributionMetricNote.textContent = "Demand in this dashboard is interpreted as " + demandUnitLabel + ".";
                    return;
                  }
                  if (metric === "miles") {
                    dayDistributionMetricNote.textContent = "Miles are total route miles grouped by day.";
                    return;
                  }
                  dayDistributionMetricNote.textContent = "Stops are total stop counts grouped by day.";
                }

                if (dayDistributionMetric) {
                  dayDistributionMetric.addEventListener("change", updateDayDistributionMetricView);
                }
                updateDayDistributionMetricView();

                modalClose?.addEventListener("click", closeRouteDayModal);
                modal?.addEventListener("click", e => {
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
                      return;
                    }

                    if (kind === "frequency") {
                      const matchedFrequency = frequencyRows.find(item => String(item.label || "") === key);
                      if (!matchedFrequency) {
                        openSimpleListModal("Service Frequency", ["No matching service-frequency details found."]);
                        return;
                      }

                      const detailLines = [
                        "Total demand: " + Number(matchedFrequency.demand || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) + " " + demandUnitLabel,
                        "Average demand per stop: " + Number(matchedFrequency.avgDemandPerStop || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) + " " + demandUnitLabel,
                        "Demand share: " + Number(matchedFrequency.demandSharePct || 0).toLocaleString(undefined, { maximumFractionDigits: 2 }) + "%",
                        "Records in frequency: " + Number(matchedFrequency.recordCount || 0).toLocaleString(),
                        "Locations in frequency: " + Number(matchedFrequency.locationCount || 0).toLocaleString(),
                        "Route+Day combinations: " + Number(matchedFrequency.routeDayCount || 0).toLocaleString()
                      ];

                      const routeLines = (Array.isArray(matchedFrequency.routeList) ? matchedFrequency.routeList : [])
                        .slice()
                        .sort((a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: "base" }))
                        .map(route => "Route: " + String(route || ""));

                      const routeDayLines = (Array.isArray(matchedFrequency.routeDayList) ? matchedFrequency.routeDayList : [])
                        .slice()
                        .sort((a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: "base" }))
                        .map(routeDay => "Route+Day: " + String(routeDay || ""));

                      openSimpleListModal(
                        "Service Frequency: " + (matchedFrequency.shortLabel || matchedFrequency.label || key),
                        detailLines.concat(routeLines, routeDayLines)
                      );
                    }
                  });
                });

                const days = Array.from(new Set(data.map(d => String(d.day || "").trim()).filter(Boolean))).sort();
                if (dayFilter) {
                  days.forEach(day => {
                    const opt = document.createElement("option");
                    opt.value = day;
                    opt.textContent = day;
                    dayFilter.appendChild(opt);
                  });
                }

                function render() {
                  const selectedDay = dayFilter ? dayFilter.value : "all";
                  const filtered = data.filter(d => {
                    if (selectedDay !== "all" && String(d.day) !== selectedDay) return false;
                    if (overOnly && overOnly.checked && !(d.totalTime > targetHours)) return false;
                    return true;
                  });

                  if (!filtered.length) {
                    host.innerHTML = '<div class="empty">No points match the selected filters.</div>';
                    return;
                  }

                  const hasStopsMilesSignal = filtered.some(p => {
                    const stops = Number(p?.stops) || 0;
                    const miles = Number(p?.miles) || 0;
                    return stops > 0 || miles > 0;
                  });
                  if (!hasStopsMilesSignal) {
                    host.innerHTML = '<div class="empty">No stops/miles values were found in the loaded summary, so this scatter chart cannot be drawn.</div>';
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

                  const circles = filtered.map(p => {
                    const color = p.totalTime > targetHours ? "#e25b53" : "#2f89df";
                    const info = (String(p.routeDay || "") + " | Stops: " + p.stops.toFixed(2) + " | Miles: " + p.miles.toFixed(2) + " | Demand: " + p.demand.toFixed(2) + " " + demandUnitLabel + " | Total Time: " + p.totalTime.toFixed(2) + "h")
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

                dayFilter?.addEventListener("change", render);
                overOnly?.addEventListener("change", render);
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
    clearSequenceLayerData({ silent: true });
    syncSequenceLayerVisibilityOnMap();

    // 5. Reset counters & stats
    document.getElementById("selectionCount").textContent = "0";

    // 6. Clear route/day checkbox UI
    document.getElementById("routeCheckboxes").innerHTML = "";
    buildDayCheckboxes();

    // 7. Reset bounds tracker
    globalBounds = L.latLngBounds();
    rebuildMultiDayServiceProfilesFromMapLayers();
    renderMultiDayServiceLabels();

    // 8. Clear bottom summary
    const summary = document.getElementById("routeSummaryTable");
    if (summary) summary.innerHTML = "No summary loaded";

    // 9. Collapse summary panel
    const panel = document.getElementById("bottomSummary");
    const btn = document.getElementById("summaryToggleBtn");
    if (panel && btn) {
      panel.classList.add("collapsed");
      panel.style.height = "40px";
      btn.textContent = SUMMARY_ARROW_COLLAPSED;
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
      locateBtn.textContent = "\u25A0";
      locateBtn.classList.add("tracking"); // turns button red
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


      locateBtn.textContent = "\u{1F4CD}";
      locateBtn.classList.remove("tracking"); // back to blue
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

    syncCurrentWorkbookSheetFromRows();
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

  // ðŸš« If toggle is OFF, force close and skip
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

  renderMultiDayServiceLabels();
});

map.on("moveend", () => {
  renderMultiDayServiceLabels();
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
// TEXAS LOCATOR (COORDINATES + ADDRESS + PLACE SEARCH)
////////////////////////////////////////////////////////////////////

const searchInput = document.getElementById("mapSearchInput");
const searchBtn = document.getElementById("mapSearchBtn");
const TEXAS_LOCATOR_BOUNDS = L.latLngBounds(
  [25.8372, -106.6456], // south-west
  [36.5007, -93.5080]   // north-east
);
const TEXAS_LOCATOR_VIEWBOX = "-106.6456,36.5007,-93.5080,25.8372";
const MAP_SEARCH_MAX_LOCAL_RESULTS = 90;
const MAP_SEARCH_MAX_GEOCODER_RESULTS = 10;
let mapSearchFocusMarker = null;
let mapSearchHighlightedRecord = null;

function isTexasLatLng(lat, lng) {
  const latNum = Number(lat);
  const lngNum = Number(lng);
  if (!Number.isFinite(latNum) || !Number.isFinite(lngNum)) return false;
  return TEXAS_LOCATOR_BOUNDS.contains(L.latLng(latNum, lngNum));
}

function clearMapSearchHighlight() {
  if (mapSearchHighlightedRecord?._base?.symbol) {
    const sym = mapSearchHighlightedRecord._base.symbol;
    mapSearchHighlightedRecord.setStyle?.({
      color: sym.color,
      fillColor: sym.color,
      fillOpacity: 0.95,
      opacity: 1,
      weight: 1
    });
  }
  mapSearchHighlightedRecord = null;
}

function clearMapSearchFocusMarker() {
  if (!mapSearchFocusMarker) return;
  try {
    map.removeLayer(mapSearchFocusMarker);
  } catch (_) {}
  mapSearchFocusMarker = null;
}

function clearMapSearchVisualFocus() {
  clearMapSearchHighlight();
  clearMapSearchFocusMarker();
}

function openSearchResultsPanel() {
  const panel = document.getElementById("searchResultsPanel");
  if (panel) panel.classList.remove("hidden");
}

function hideSearchResultsPanel() {
  const panel = document.getElementById("searchResultsPanel");
  if (panel) panel.classList.add("hidden");
}

function setSearchStatusMessage(message, isError = false) {
  const resultsList = document.getElementById("searchResultsList");
  if (!resultsList) return;
  resultsList.innerHTML = "";
  const node = document.createElement("div");
  node.className = `search-result-status${isError ? " error" : ""}`;
  node.textContent = String(message || "");
  resultsList.appendChild(node);
}

function markSearchResultSelected(node) {
  document
    .querySelectorAll(".search-result-item")
    .forEach(el => el.classList.remove("selected"));
  node.classList.add("selected");
}

function buildRecordAddressLabel(row) {
  const value = [
    row?.["CSADR#"] || "",
    row?.["CSSDIR"] || "",
    row?.["CSSTRT"] || "",
    row?.["CSSFUX"] || ""
  ]
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
  return value || "Unnamed stop";
}

function buildRecordSearchText(row) {
  const preferred = [];
  Object.entries(row || {}).forEach(([key, value]) => {
    if (value == null || value === "") return;
    const k = String(key || "").toLowerCase();
    if (k === "latitude" || k === "longitude") return;
    if (/(address|street|city|state|zip|route|day|name|account|customer|stop|id|cs)/i.test(k)) {
      preferred.push(String(value));
    }
  });
  if (preferred.length) return preferred.join(" ").toLowerCase();

  return Object.values(row || {})
    .filter(v => v != null && String(v).trim() !== "")
    .slice(0, 30)
    .map(v => String(v))
    .join(" ")
    .toLowerCase();
}

function parseCoordinateQuery(query) {
  const parts = String(query || "").match(/[-+]?\d+(?:\.\d+)?/g);
  if (!parts || parts.length < 2) return null;

  const first = Number(parts[0]);
  const second = Number(parts[1]);
  if (!Number.isFinite(first) || !Number.isFinite(second)) return null;

  const isValidLatLng = (lat, lng) => (
    Number.isFinite(lat) &&
    Number.isFinite(lng) &&
    lat >= -90 && lat <= 90 &&
    lng >= -180 && lng <= 180
  );

  const candidates = [
    { lat: first, lng: second, swapped: false },
    { lat: second, lng: first, swapped: true }
  ].filter(c => isValidLatLng(c.lat, c.lng));

  if (!candidates.length) return null;

  const texasCandidate = candidates.find(c => isTexasLatLng(c.lat, c.lng));
  if (texasCandidate) return { ...texasCandidate, inTexas: true };
  return { ...candidates[0], inTexas: false };
}

function focusRecordSearchResult(entry) {
  const marker = entry?.marker;
  if (!marker) return;

  clearMapSearchFocusMarker();
  if (mapSearchHighlightedRecord && mapSearchHighlightedRecord !== marker) {
    clearMapSearchHighlight();
  }

  mapSearchHighlightedRecord = marker;
  marker.setStyle?.({
    color: "#ffe066",
    fillColor: "#ffe066",
    fillOpacity: 1,
    opacity: 1,
    weight: 2
  });

  const ll = getLayerLatLng(marker);
  if (ll) map.setView(ll, Math.max(map.getZoom(), 17));
  marker.openPopup?.();
}

function focusCoordinateOrGeocodeResult(entry) {
  const lat = Number(entry?.lat);
  const lng = Number(entry?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;

  clearMapSearchHighlight();
  clearMapSearchFocusMarker();

  mapSearchFocusMarker = L.circleMarker([lat, lng], {
    radius: 8,
    color: "#ffd24c",
    fillColor: "#ff8c32",
    fillOpacity: 0.96,
    weight: 2
  }).addTo(map);

  if (entry?.title) {
    mapSearchFocusMarker.bindPopup(String(entry.title)).openPopup();
  }

  map.setView([lat, lng], Math.max(map.getZoom(), 16));
}

function appendSearchResultItem(container, { title, meta, type = "", onSelect }) {
  if (!container) return;
  const node = document.createElement("button");
  node.type = "button";
  node.className = `search-result-item${type ? ` ${type}` : ""}`;

  const titleNode = document.createElement("div");
  titleNode.className = "search-result-title";
  titleNode.textContent = String(title || "Result");

  const metaNode = document.createElement("div");
  metaNode.className = "search-result-meta";
  metaNode.textContent = String(meta || "");

  node.append(titleNode, metaNode);
  node.addEventListener("click", () => {
    markSearchResultSelected(node);
    onSelect?.();
  });
  container.appendChild(node);
}

function collectLocalTexasSearchResults(queryText) {
  const needle = String(queryText || "").trim().toLowerCase();
  if (!needle) return [];

  const seen = new Set();
  const matches = [];

  Object.entries(routeDayGroups).forEach(([key, group]) => {
    const layers = Array.isArray(group?.layers) ? group.layers : [];
    layers.forEach(marker => {
      const row = marker?._rowRef;
      const latlng = getLayerLatLng(marker);
      if (!row || !latlng || !isTexasLatLng(latlng.lat, latlng.lng)) return;

      const address = buildRecordAddressLabel(row);
      const searchable = `${address} ${buildRecordSearchText(row)}`;
      if (!searchable.includes(needle)) return;

      const rowId = Number(marker?._rowId);
      const dedupeKey = Number.isFinite(rowId)
        ? `row:${rowId}`
        : `${latlng.lat.toFixed(7)}|${latlng.lng.toFixed(7)}|${address}`;
      if (seen.has(dedupeKey)) return;
      seen.add(dedupeKey);

      const route = getMappedRowRoute(row) || "Unassigned";
      const dayRaw = getMappedRowDayRaw(row);
      const dayToken = normalizeDayToken(dayRaw) || dayRaw;
      const dayLabel = dayName(Number(dayToken)) || dayToken || "No Day";

      matches.push({
        marker,
        row,
        title: address,
        meta: `Record stop - Route ${route} | ${dayLabel}`
      });
    });
  });

  matches.sort((a, b) => String(a.title).localeCompare(String(b.title), undefined, { sensitivity: "base" }));
  return matches.slice(0, MAP_SEARCH_MAX_LOCAL_RESULTS);
}

function buildTexasGeocoderQuery(queryText) {
  const raw = String(queryText || "").trim();
  if (!raw) return "";
  if (/\btexas\b|\btx\b/i.test(raw)) return raw;
  return `${raw}, Texas`;
}

async function fetchTexasGeocoderResults(queryText) {
  const geocoderQuery = buildTexasGeocoderQuery(queryText);
  if (!geocoderQuery) return [];

  const params = new URLSearchParams({
    q: geocoderQuery,
    format: "jsonv2",
    addressdetails: "1",
    limit: String(MAP_SEARCH_MAX_GEOCODER_RESULTS),
    countrycodes: "us",
    bounded: "1",
    viewbox: TEXAS_LOCATOR_VIEWBOX
  });

  const response = await fetch(`https://nominatim.openstreetmap.org/search?${params.toString()}`, {
    headers: {
      Accept: "application/json",
      "Accept-Language": "en-US,en"
    }
  });
  if (!response.ok) {
    throw new Error(`Texas geocoder request failed (${response.status}).`);
  }

  const payload = await response.json();
  const rows = Array.isArray(payload) ? payload : [];
  const seen = new Set();
  const results = [];

  rows.forEach(item => {
    const lat = Number(item?.lat);
    const lng = Number(item?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
    if (!isTexasLatLng(lat, lng)) return;

    const key = `${lat.toFixed(6)}|${lng.toFixed(6)}`;
    if (seen.has(key)) return;
    seen.add(key);

    results.push({
      lat,
      lng,
      title: String(item?.display_name || geocoderQuery).trim(),
      meta: "Texas locator result"
    });
  });

  return results.slice(0, MAP_SEARCH_MAX_GEOCODER_RESULTS);
}

async function searchMapByAddress() {
  if (!searchInput) return;

  const rawQuery = String(searchInput.value || "").trim();
  if (!rawQuery) return;

  openSearchResultsPanel();
  setSearchStatusMessage("Searching Texas locations...");

  const resultsList = document.getElementById("searchResultsList");
  if (!resultsList) return;

  const coordinateMatch = parseCoordinateQuery(rawQuery);
  if (coordinateMatch) {
    if (!coordinateMatch.inTexas) {
      setSearchStatusMessage("Coordinates were recognized, but they are outside Texas. Texas-only search is enabled.", true);
      return;
    }
    resultsList.innerHTML = "";
    const coordText = `${coordinateMatch.lat.toFixed(6)}, ${coordinateMatch.lng.toFixed(6)}`;
    appendSearchResultItem(resultsList, {
      title: `Coordinates: ${coordText}`,
      meta: coordinateMatch.swapped ? "Texas coordinate (interpreted as lon,lat input)" : "Texas coordinate",
      type: "search-result-coordinate",
      onSelect: () => focusCoordinateOrGeocodeResult({
        lat: coordinateMatch.lat,
        lng: coordinateMatch.lng,
        title: `Texas coordinates: ${coordText}`
      })
    });
    const first = resultsList.querySelector(".search-result-item");
    first?.click();
    return;
  }

  const localMatches = collectLocalTexasSearchResults(rawQuery);
  let geocoderMatches = [];
  let geocoderError = "";

  try {
    geocoderMatches = await fetchTexasGeocoderResults(rawQuery);
  } catch (error) {
    geocoderError = error?.message || "Texas geocoder is unavailable.";
    console.warn("Texas geocoder search failed:", error);
  }

  if (!localMatches.length && !geocoderMatches.length) {
    const fallback = geocoderError
      ? `No Texas matches found in map records. ${geocoderError}`
      : "No Texas matches found. Try a full address, city, ZIP, or coordinates.";
    setSearchStatusMessage(fallback, true);
    return;
  }

  resultsList.innerHTML = "";

  localMatches.forEach(match => {
    appendSearchResultItem(resultsList, {
      title: match.title,
      meta: match.meta,
      type: "search-result-local",
      onSelect: () => focusRecordSearchResult(match)
    });
  });

  geocoderMatches.forEach(match => {
    appendSearchResultItem(resultsList, {
      title: match.title,
      meta: match.meta,
      type: "search-result-geocode",
      onSelect: () => focusCoordinateOrGeocodeResult(match)
    });
  });

  if (geocoderError && (localMatches.length || geocoderMatches.length)) {
    const note = document.createElement("div");
    note.className = "search-result-status";
    note.textContent = `Note: ${geocoderError}`;
    resultsList.appendChild(note);
  }

  const first = resultsList.querySelector(".search-result-item");
  first?.click();
}

if (searchBtn) {
  searchBtn.addEventListener("click", searchMapByAddress);
}

if (searchInput) {
  searchInput.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      searchMapByAddress();
    }
  });
}

const clearSearchBtn = document.getElementById("clearSearchResults");

if (clearSearchBtn) {
  clearSearchBtn.addEventListener("click", () => {
    if (searchInput) searchInput.value = "";
    const resultsList = document.getElementById("searchResultsList");
    if (resultsList) resultsList.innerHTML = "";
    clearMapSearchVisualFocus();
    hideSearchResultsPanel();
  });
}
////////////////central save function
async function saveWorkbookToCloud() {
  if (!window._currentWorkbook || !Array.isArray(window._currentRows)) return false;
  const filePath = String(window._currentFilePath || "").trim();
  if (!filePath) return false;
  syncCurrentWorkbookSheetFromRows();

  const bookType = filePath.toLowerCase().endsWith(".xlsm")
    ? "xlsm"
    : "xlsx";

  const wbArray = XLSX.write(window._currentWorkbook, {
    bookType,
    type: "array"
  });

  const { error } = await sb.storage
    .from(BUCKET)
    .upload(filePath, wbArray, {
      upsert: true,
      contentType: bookType === "xlsm"
        ? "application/vnd.ms-excel.sheet.macroEnabled.12"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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

function setRouteDaySidebarExpanded(expanded = true) {
  const routeDayToggle = document.getElementById("routeDayToggle");
  const routeDayContent = document.getElementById("routeDayContent");
  if (!routeDayToggle || !routeDayContent) return false;
  const open = !!expanded;
  routeDayContent.classList.toggle("collapsed", !open);
  routeDayToggle.classList.toggle("open", open);
  return true;
}

function autoExpandRouteDaySidebarIfLayersPresent() {
  if (!Object.keys(routeDayGroups).length) return;
  setRouteDaySidebarExpanded(true);
}
  
// ===== ROUTE + DAY COLLAPSIBLE =====
const routeDayToggle = document.getElementById("routeDayToggle");
const routeDayContent = document.getElementById("routeDayContent");

if (routeDayToggle && routeDayContent) {

  // Closed by default
  setRouteDaySidebarExpanded(false);

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

































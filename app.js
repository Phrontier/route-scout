const AIRPORTS_URLS = [
  "https://raw.githubusercontent.com/davidmegginson/ourairports-data/main/airports.csv",
  "https://cdn.jsdelivr.net/gh/davidmegginson/ourairports-data@main/airports.csv",
  "https://ourairports.com/data/airports.csv"
];
const RUNWAYS_URLS = [
  "https://raw.githubusercontent.com/davidmegginson/ourairports-data/main/runways.csv",
  "https://cdn.jsdelivr.net/gh/davidmegginson/ourairports-data@main/runways.csv",
  "https://ourairports.com/data/runways.csv"
];
const OPEN_METEO_FORECAST = "https://api.open-meteo.com/v1/forecast";
const OPEN_METEO_ARCHIVE = "https://archive-api.open-meteo.com/v1/archive";
const AVIATION_API_CHARTS = "https://api.aviationapi.com/v1/charts";
const FAA_DTPP_SEARCH_URL = "https://www.faa.gov/air_traffic/flight_info/aeronav/digital_products/dtpp/search/";
const FAA_DTPP_XML_MATCH = /https?:\\?\/\\?\/aeronav\.faa\.gov\\?\/upload_[^"'\s]+d-tpp_[^"'\s]+_Metafile\.xml/gi;
const FAA_IAP_CODES = new Set(["IAP", "IAPMIN", "IAPCOPTER", "IAPMIL"]);
const APP_VERSION = "0.0.26";
const VERSION_FILE_PATH = "version.json";
const LOADING_MESSAGES_PATH = "loading-messages.txt";
const LOCAL_HOSTS = new Set(["localhost", "127.0.0.1", "::1"]);
const PROFILE_BUILD_CONCURRENCY = 6;
const PROXY_PATH = "/proxy";
const MAP_STYLE_STREET = "street";
const MAP_STYLE_IFR_LOW = "ifr_low";
const MAP_STYLE_IFR_HIGH = "ifr_high";
const MAP_DEFAULT_STYLE = MAP_STYLE_IFR_LOW;
const IFR_LOW_TILE_URL = "https://tiles.arcgis.com/tiles/ssFJjBXIUyZDrSYZ/arcgis/rest/services/IFR_AreaLow/MapServer/tile/{z}/{y}/{x}";
const IFR_HIGH_TILE_URL = "https://tiles.arcgis.com/tiles/ssFJjBXIUyZDrSYZ/arcgis/rest/services/IFR_High/MapServer/tile/{z}/{y}/{x}";
const MAP_IFR_FALLBACK_WINDOW_MS = 2600;
const MAP_IFR_MIN_ERROR_THRESHOLD = 5;
const MAP_IFR_MAX_SUCCESS_FOR_FALLBACK = 1;
const MAP_STYLE_OPTIONS = [
  { value: MAP_STYLE_IFR_LOW, label: "IFR Low" },
  { value: MAP_STYLE_IFR_HIGH, label: "IFR High" },
  { value: MAP_STYLE_STREET, label: "Street" }
];
const SINGLE_PROFILE_POOL_SIZE = 45;
const DOUBLE_PROFILE_POOL_SIZE = 80;
const TRIPLE_PROFILE_POOL_SIZE = 70;
const TRIPLE_COMBINATOR_POOL_SIZE = 40;

const FLIGHT_RULES = {
  single: {
    label: "Single",
    stopCount: 1,
    stopPolicy: "No stop",
    legMin: 25,
    sliderMin: 50,
    sliderMax: 150,
    defaultMax: 100
  },
  double: {
    label: "Double",
    stopCount: 1,
    stopPolicy: "Stop at destination",
    legMin: 30,
    sliderMin: 80,
    sliderMax: 350,
    defaultMax: 220
  },
  triple: {
    label: "Triple",
    stopCount: 2,
    stopPolicy: "Two-airport route",
    legMin: 30,
    sliderMin: 100,
    sliderMax: 350,
    defaultMax: 240
  }
};

const SCORE_WEIGHTS = {
  diversityMax: 45,
  quantityMax: 20,
  windMax: 20,
  runwayMax: 10,
  trafficMax: 5
};
const INCLUDE_TACAN_FOR_T6_SCORING = false;

const FALLBACK_AIRPORTS = [
  { ident: "KSQL", name: "San Carlos", lat: 37.5119, lon: -122.2495, type: "small_airport" },
  { ident: "KOAK", name: "Metro Oakland Intl", lat: 37.7213, lon: -122.221, type: "large_airport" },
  { ident: "KSJC", name: "San Jose Intl", lat: 37.3639, lon: -121.9289, type: "large_airport" },
  { ident: "KSFO", name: "San Francisco Intl", lat: 37.6188, lon: -122.375, type: "large_airport" },
  { ident: "KSMF", name: "Sacramento Intl", lat: 38.6954, lon: -121.591, type: "large_airport" },
  { ident: "KMRY", name: "Monterey Regional", lat: 36.587, lon: -121.843, type: "medium_airport" },
  { ident: "KRNO", name: "Reno/Tahoe Intl", lat: 39.4991, lon: -119.7681, type: "large_airport" },
  { ident: "KBFL", name: "Meadows Field", lat: 35.4336, lon: -119.0577, type: "medium_airport" }
];

const FALLBACK_RUNWAYS = new Map([
  ["KSQL", [{ lengthFt: 2600, widthFt: 75, ends: [{ ident: "12", heading: 120 }, { ident: "30", heading: 300 }] }]],
  ["KOAK", [{ lengthFt: 6200, widthFt: 150, ends: [{ ident: "10L", heading: 100 }, { ident: "28R", heading: 280 }] }]],
  ["KSJC", [{ lengthFt: 11000, widthFt: 150, ends: [{ ident: "12L", heading: 120 }, { ident: "30R", heading: 300 }] }]],
  ["KSFO", [{ lengthFt: 11870, widthFt: 200, ends: [{ ident: "10L", heading: 100 }, { ident: "28R", heading: 280 }] }]],
  ["KSMF", [{ lengthFt: 8600, widthFt: 150, ends: [{ ident: "17L", heading: 170 }, { ident: "35R", heading: 350 }] }]],
  ["KMRY", [{ lengthFt: 7000, widthFt: 150, ends: [{ ident: "10R", heading: 100 }, { ident: "28L", heading: 280 }] }]],
  ["KRNO", [{ lengthFt: 11000, widthFt: 150, ends: [{ ident: "17R", heading: 170 }, { ident: "35L", heading: 350 }] }]],
  ["KBFL", [{ lengthFt: 10855, widthFt: 150, ends: [{ ident: "12L", heading: 120 }, { ident: "30R", heading: 300 }] }]]
]);

const cache = {
  airports: null,
  runwaysByAirport: null,
  weatherByKey: new Map(),
  approachesByAirport: new Map(),
  faaApproachMapPromise: null,
  aviationChartsByAirport: new Map()
};

const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");
const formEl = document.getElementById("planner-form");
const versionTickerEl = document.getElementById("version-ticker");
const dateInput = document.getElementById("flight-date");
const typeInput = document.getElementById("flight-type");
const distanceInput = document.getElementById("max-distance");
const distanceValueEl = document.getElementById("distance-value");
const distanceHelpEl = document.getElementById("distance-help");
const generateBtn = document.getElementById("generate-btn");
const allApproachesInput = document.getElementById("all-approaches");
const approachTypeInputs = [...formEl.querySelectorAll('input[name="approachType"]')];

let lastModel = null;
const expandedRouteIds = new Set();
let visibleRouteCount = 10;
const routeMapInstances = new Map();
const mapStyleById = new Map();
let loadingTickerId = null;
let loadingStartedAt = 0;
let loadingQueue = [];
let loadingLongWaitQueue = [];

const DEFAULT_LOADING_PHRASES = [
  "Comparing runway winds with likely traffic flows...",
  "Reading approach plate metadata and building training sets...",
  "Ranking route options by training value...",
  "Cross-checking approach diversity for each candidate stop...",
  "Ensuring no UFOs are blocking the route...",
  "Calibrating the sarcasm-driven autopilot...",
  "Politely arguing with the weather model...",
  "Re-checking runways so we do not invent new pavement...",
  "Computing circling-approach opportunities...",
  "Double-checking that headwinds are actually headwinds...",
  "Balancing approach variety with route practicality...",
  "Verifying published procedures against runway geometry...",
  "Estimating which approaches are most useful for training...",
  "Bribing the map tiles with virtual coffee...",
  "Asking ATC nicely for a shortcut (simulation only)...",
  "Checking if gremlins edited the CSV overnight...",
  "Looking for bad ideas so we can avoid them professionally..."
];

const DEFAULT_LONG_WAIT_PHRASES = [
  "This is taking longer than planned. We are still working on your route.",
  "Still crunching. Sorry this is taking so long to do your own work for you.",
  "Long load detected. The planner is busy overthinking every approach option.",
  "Yes, it is still loading. No, the airplane isn't going to file itself."
];
let loadingPhrases = [...DEFAULT_LOADING_PHRASES];
let longWaitPhrases = [...DEFAULT_LONG_WAIT_PHRASES];

dateInput.value = new Date().toISOString().slice(0, 10);
initVersionTicker();
void initLoadingPhrases();
updateDistanceControl(typeInput.value);
typeInput.addEventListener("change", () => updateDistanceControl(typeInput.value));
distanceInput.addEventListener("input", () => {
  distanceValueEl.textContent = `${distanceInput.value} NM`;
});
allApproachesInput?.addEventListener("change", () => syncApproachTypeUI());
for (const input of approachTypeInputs) {
  input.addEventListener("change", () => {
    if (input.checked) {
      allApproachesInput.checked = false;
    }
    if (!approachTypeInputs.some((el) => el.checked)) {
      allApproachesInput.checked = true;
    }
    syncApproachTypeUI();
  });
}
syncApproachTypeUI();

resultsEl.addEventListener("click", (event) => {
  const loadMoreBtn = event.target.closest("button[data-action='load-more']");
  if (loadMoreBtn && lastModel) {
    visibleRouteCount += 10;
    renderAll(lastModel);
    return;
  }

  const button = event.target.closest("button[data-route-id]");
  if (!button || !lastModel) return;
  const routeId = String(button.dataset.routeId || "");
  if (!routeId) return;
  if (expandedRouteIds.has(routeId)) expandedRouteIds.delete(routeId);
  else expandedRouteIds.add(routeId);
  renderAll(lastModel);
});

resultsEl.addEventListener("change", (event) => {
  const picker = event.target.closest(".rs-airfield-section-picker");
  if (picker) {
    activateAirfieldSection(picker.dataset.tabRoot, picker.value);
    return;
  }

  const styleSelect = event.target.closest(".rs-map-style-select");
  if (styleSelect) {
    const mapId = String(styleSelect.dataset.mapId || "");
    const style = String(styleSelect.value || MAP_DEFAULT_STYLE);
    if (mapId) {
      mapStyleById.set(mapId, style);
      const ctx = routeMapInstances.get(mapId);
      if (ctx) applyMapStyle(ctx, style, true);
    }
  }
});

resultsEl.addEventListener("shown.bs.tab", (event) => {
  const tabButton = event.target.closest("[data-bs-target]");
  if (!tabButton) return;
  const tabRoot = tabButton.closest(".rs-airfield-tabs");
  if (!tabRoot) return;
  const paneSelector = tabButton.getAttribute("data-bs-target") || "";
  if (!paneSelector.startsWith("#")) return;
  const picker = document.querySelector(`.rs-airfield-section-picker[data-tab-root="${tabRoot.id}"]`);
  if (picker) picker.value = paneSelector.slice(1);
});

formEl.addEventListener("submit", async (event) => {
  event.preventDefault();

  const originIdent = document.getElementById("origin").value.trim().toUpperCase();
  const flightType = typeInput.value;
  const flightDate = document.getElementById("flight-date").value;
  const timeLocal = document.getElementById("time-local").value;
  const maxLegNm = Number(distanceInput.value);
  const requiredApproachTypes = allApproachesInput?.checked
    ? []
    : [...formEl.querySelectorAll('input[name="approachType"]:checked')].map((el) => el.value);

  startLoadingTicker();
  generateBtn.disabled = true;
  generateBtn.textContent = "Generating...";
  try {
    setStatus("Loading airport/runway data...", false, true);
    const { airports, runwaysByAirport, usingFallback } = await loadCoreData();

    const origin = airports.find((a) => a.ident === originIdent);
    if (!origin) throw new Error(`Origin ${originIdent} not found in loaded airport set.`);

    setStatus("Building airport training profiles (winds + approaches)...", false, true);
    const originProfile = await buildAirportProfile({
      airport: origin,
      origin,
      runwaysByAirport,
      flightDate,
      timeLocal
    });
    const { profiles, candidateStats } = await buildAirportProfiles({
      origin,
      airports,
      runwaysByAirport,
      flightType,
      flightDate,
      timeLocal,
      maxLegNm
    });

    setStatus("Generating and scoring multiple route options...", false, true);
    const { routes, routeStats } = generateRoutes({ origin, profiles, flightType, maxLegNm, requiredApproachTypes });
    if (!routes.length) {
      throw new Error("No valid routes found for this setup. Expand distance or change origin/date.");
    }

    const model = {
      usingFallback,
      origin,
      originProfile,
      flightType,
      flightDate,
      timeLocal,
      maxLegNm,
      requiredApproachTypes,
      routes,
      candidateStats,
      routeStats
    };

    expandedRouteIds.clear();
    visibleRouteCount = 10;
    lastModel = model;
    renderAll(model);
    setStatus(`Generated ${routes.length} route options from ${origin.ident}.`);
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Failed to generate routes.", true);
    resultsEl.innerHTML = "";
  } finally {
    stopLoadingTicker();
    generateBtn.disabled = false;
    generateBtn.textContent = "Generate Routes";
  }
});

function updateDistanceControl(flightType) {
  const rules = FLIGHT_RULES[flightType];
  distanceInput.disabled = false;
  distanceInput.min = String(rules.sliderMin);
  distanceInput.max = String(rules.sliderMax);
  distanceInput.value = String(rules.defaultMax);

  if (flightType === "single") {
    distanceHelpEl.textContent = "Single routes are best kept local; set max leg up to 150 NM.";
  } else if (flightType === "double") {
    distanceHelpEl.textContent = "Double routes can go farther since you stop at destination (up to 350 NM).";
  } else {
    distanceHelpEl.textContent = "Triple routes enforce the final return leg within this distance.";
  }

  distanceValueEl.textContent = `${distanceInput.value} NM`;
}

function setStatus(text, isError = false, isLoading = false) {
  statusEl.textContent = text;
  statusEl.classList.remove("alert-info", "alert-danger", "alert-warning");
  statusEl.classList.add(isError ? "alert-danger" : "alert-info");
  if (isLoading) statusEl.classList.add("rs-status-loading");
  else statusEl.classList.remove("rs-status-loading");
}

function syncApproachTypeUI() {
  const useAll = Boolean(allApproachesInput?.checked);
  const fieldset = document.querySelector(".approach-needs");
  fieldset?.classList.toggle("all-selected", useAll);
  for (const input of approachTypeInputs) {
    input.disabled = useAll;
    if (useAll) input.checked = false;
  }
}

function startLoadingTicker() {
  stopLoadingTicker();
  loadingStartedAt = Date.now();
  loadingQueue = shuffleCopy(loadingPhrases);
  loadingLongWaitQueue = shuffleCopy(longWaitPhrases);
  const first = loadingQueue.shift() || "Working on route options...";
  setStatus(first, false, true);

  loadingTickerId = setInterval(() => {
    const elapsed = Date.now() - loadingStartedAt;
    if (loadingQueue.length) {
      setStatus(loadingQueue.shift(), false, true);
      return;
    }
    if (elapsed > 17000 && loadingLongWaitQueue.length) {
      setStatus(loadingLongWaitQueue.shift(), false, true);
      return;
    }
    stopLoadingTicker();
  }, 4200);
}

function stopLoadingTicker() {
  if (loadingTickerId) {
    clearInterval(loadingTickerId);
    loadingTickerId = null;
  }
}

async function initVersionTicker() {
  if (!versionTickerEl) return;

  let version = APP_VERSION;
  let channel = "prod";

  try {
    const bust = Date.now();
    const response = await fetch(`${VERSION_FILE_PATH}?t=${bust}`, { cache: "no-store" });
    if (response.ok) {
      const json = await response.json();
      const fileVersion = String(json?.version || "").trim();
      if (fileVersion) version = fileVersion;
      const fileChannel = String(json?.channel || "").trim();
      if (fileChannel) channel = fileChannel;
    }
  } catch (_error) {
    // Keep built-in version fallback.
  }

  versionTickerEl.textContent = `Version v${version} (${channel})`;
}

async function initLoadingPhrases() {
  try {
    const response = await fetch(`${LOADING_MESSAGES_PATH}?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) return;
    const text = await response.text();
    const parsed = parseLoadingMessages(text);
    if (parsed.normal.length) loadingPhrases = parsed.normal;
    if (parsed.longWait.length) longWaitPhrases = parsed.longWait;
  } catch (_error) {
    // Keep defaults.
  }
}

function parseLoadingMessages(text) {
  const normal = [];
  const longWait = [];
  let section = "normal";

  for (const rawLine of String(text || "").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    if (/^\[normal\]$/i.test(line)) {
      section = "normal";
      continue;
    }
    if (/^\[long_wait\]$/i.test(line)) {
      section = "long_wait";
      continue;
    }

    if (section === "long_wait") longWait.push(line);
    else normal.push(line);
  }

  return { normal, longWait };
}

async function loadCoreData() {
  if (cache.airports && cache.runwaysByAirport) {
    return { airports: cache.airports, runwaysByAirport: cache.runwaysByAirport, usingFallback: false };
  }

  let usingFallback = false;
  let airportRows;
  let runwayRows;

  try {
    const [airportsCsv, runwaysCsv] = await Promise.all([
      fetchTextFromAny(AIRPORTS_URLS, {
        validateText: (text) => looksLikeCsvWithHeaders(text, ["ident", "latitude_deg", "longitude_deg", "iso_country", "type"])
      }),
      fetchTextFromAny(RUNWAYS_URLS, {
        validateText: (text) => looksLikeCsvWithHeaders(text, ["airport_ident", "length_ft", "le_ident", "he_ident"])
      })
    ]);
    airportRows = parseCsv(airportsCsv);
    runwayRows = parseCsv(runwaysCsv);
  } catch (_error) {
    usingFallback = true;
  }

  if (usingFallback) {
    cache.airports = FALLBACK_AIRPORTS;
    cache.runwaysByAirport = FALLBACK_RUNWAYS;
    return { airports: cache.airports, runwaysByAirport: cache.runwaysByAirport, usingFallback: true };
  }

  const airportsByIdent = new Map();
  const sourceToCanonicalIdent = new Map();

  for (const row of airportRows) {
    const isUS = row.iso_country === "US";
    const typeOk = ["small_airport", "medium_airport", "large_airport"].includes(row.type);
    if (!isUS || !typeOk) continue;

    const sourceIdent = String(row.ident || "").trim().toUpperCase();
    const canonicalIdent =
      normalizeAirportCode(row.gps_code, row.iso_country) ||
      normalizeAirportCode(row.ident, row.iso_country) ||
      normalizeAirportCode(row.local_code, row.iso_country);

    if (!sourceIdent || !canonicalIdent) continue;

    const lat = Number(row.latitude_deg);
    const lon = Number(row.longitude_deg);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;

    if (!airportsByIdent.has(canonicalIdent)) {
      airportsByIdent.set(canonicalIdent, {
        ident: canonicalIdent,
        name: row.name,
        lat,
        lon,
        type: row.type,
        elevationFt: Number(row.elevation_ft) || null,
        scheduledService: String(row.scheduled_service || "").toLowerCase() === "yes",
        keywords: String(row.keywords || "")
      });
    }

    sourceToCanonicalIdent.set(sourceIdent, canonicalIdent);
    const gpsCode = String(row.gps_code || "").trim().toUpperCase();
    if (gpsCode) sourceToCanonicalIdent.set(gpsCode, canonicalIdent);
    const localCode = String(row.local_code || "").trim().toUpperCase();
    if (localCode) sourceToCanonicalIdent.set(localCode, canonicalIdent);
  }

  const airports = [...airportsByIdent.values()];
  const runwaysByAirport = new Map();

  for (const row of runwayRows) {
    const sourceIdent = String(row.airport_ident || "").trim().toUpperCase();
    const ident = sourceToCanonicalIdent.get(sourceIdent) || normalizeAirportCode(sourceIdent, "US");
    if (!ident || !airportsByIdent.has(ident)) continue;
    if (String(row.closed || "").trim() === "1") continue;

    const runway = {
      lengthFt: Number(row.length_ft) || 0,
      widthFt: Number(row.width_ft) || 0,
      lighted: String(row.lighted || "").trim() === "1",
      ends: [
        { ident: normalizeRunwayIdent(row.le_ident), heading: Number(row.le_heading_degT) },
        { ident: normalizeRunwayIdent(row.he_ident), heading: Number(row.he_heading_degT) }
      ].filter((end) => end.ident && Number.isFinite(end.heading))
    };

    if (!runway.ends.length) continue;
    if (!runwaysByAirport.has(ident)) runwaysByAirport.set(ident, []);
    runwaysByAirport.get(ident).push(runway);
  }

  cache.airports = airports;
  cache.runwaysByAirport = runwaysByAirport;

  return { airports, runwaysByAirport, usingFallback: false };
}

async function buildAirportProfiles({
  origin,
  airports,
  runwaysByAirport,
  flightType,
  flightDate,
  timeLocal,
  maxLegNm
}) {
  const candidates = airports
    .filter((a) => a.ident !== origin.ident)
    .map((a) => ({
      ...a,
      distanceFromOriginNm: haversineNm(origin.lat, origin.lon, a.lat, a.lon),
      runways: runwaysByAirport.get(a.ident) || []
    }))
    .filter((a) => a.distanceFromOriginNm <= maxLegNm)
    .filter((a) => hasRunwayAtLeast(a.runways, 4000));

  if (!candidates.length) {
    return {
      profiles: [],
      candidateStats: {
        totalFilteredCandidates: 0,
        profilePoolSizeUsed: 0
      }
    };
  }

  let pool = [];
  if (flightType === "single") {
    pool = [...candidates]
      .sort((a, b) => a.distanceFromOriginNm - b.distanceFromOriginNm)
      .slice(0, SINGLE_PROFILE_POOL_SIZE);
  } else {
    const limit = flightType === "double" ? DOUBLE_PROFILE_POOL_SIZE : TRIPLE_PROFILE_POOL_SIZE;
    pool = candidates
      .map((airport) => ({
        airport,
        preScore: scoreAirportCandidateForPool(airport, flightType, maxLegNm)
      }))
      .sort((a, b) => b.preScore - a.preScore || a.airport.ident.localeCompare(b.airport.ident))
      .slice(0, limit)
      .map((row) => row.airport);
  }

  const profiles = await mapWithConcurrency(pool, PROFILE_BUILD_CONCURRENCY, (airport) => buildAirportProfile({
    airport,
    origin,
    runwaysByAirport,
    flightDate,
    timeLocal
  }));

  return {
    profiles,
    candidateStats: {
      totalFilteredCandidates: candidates.length,
      profilePoolSizeUsed: pool.length
    }
  };
}

async function buildAirportProfile({ airport, origin, runwaysByAirport, flightDate, timeLocal }) {
  const runways = airport.runways || runwaysByAirport.get(airport.ident) || [];
  const distanceFromOriginNm = airport.ident === origin.ident
    ? 0
    : haversineNm(origin.lat, origin.lon, airport.lat, airport.lon);

  const wind = await getSurfaceWindKnots(airport, flightDate, timeLocal);
  const runwayRanking = rankRunwayEnds(runways, wind);
  const likelyRunway = runwayRanking[0] || null;

  const supplementalCharts = await fetchAviationHtmlCharts(airport.ident);
  const approaches = await fetchApproaches(airport.ident, supplementalCharts);
  const approachStats = buildApproachStats(approaches, likelyRunway?.runwayIdent, supplementalCharts.allCharts);
  const suitability = evaluateAirportSuitability(airport);

  const scoring = scoreAirportTraining({ airport, likelyRunway, approachStats, suitability });

  return {
    ...airport,
    runways,
    distanceFromOriginNm,
    wind,
    runwayRanking,
    likelyRunway,
    approaches,
    approachStats,
    suitability,
    trainingScore: scoring.total,
    scoreBreakdown: scoring.breakdown
  };
}

function generateRoutes({ origin, profiles, flightType, maxLegNm, requiredApproachTypes = [] }) {
  const rules = FLIGHT_RULES[flightType];

  if (flightType === "single" || flightType === "double") {
    const allRoutes = profiles
      .filter((p) => p.distanceFromOriginNm >= rules.legMin)
      .filter((p) => p.distanceFromOriginNm <= maxLegNm)
      .map((p) => {
        const legs = [
          { from: origin.ident, to: p.ident, distanceNm: p.distanceFromOriginNm },
          { from: p.ident, to: origin.ident, distanceNm: p.distanceFromOriginNm }
        ];
        return {
          id: `${flightType}-${p.ident}`,
          flightType,
          stops: [p],
          legs,
          totalDistanceNm: p.distanceFromOriginNm * 2,
          score: scoreRoute({
            stops: [p],
            totalDistanceNm: p.distanceFromOriginNm * 2,
            flightType,
            maxLegNm,
            requiredApproachTypes
          }),
          summary: `${origin.ident} -> ${p.ident} -> ${origin.ident}`
        };
      })
      .sort((a, b) => b.score - a.score);

    return {
      routes: allRoutes,
      routeStats: {
        preTopRouteCount: allRoutes.length
      }
    };
  }

  const top = [...profiles].sort((a, b) => b.trainingScore - a.trainingScore).slice(0, TRIPLE_COMBINATOR_POOL_SIZE);
  const routes = [];

  for (let i = 0; i < top.length; i += 1) {
    for (let j = 0; j < top.length; j += 1) {
      if (i === j) continue;

      const a = top[i];
      const b = top[j];

      const leg1 = haversineNm(origin.lat, origin.lon, a.lat, a.lon);
      const leg2 = haversineNm(a.lat, a.lon, b.lat, b.lon);
      const leg3 = haversineNm(b.lat, b.lon, origin.lat, origin.lon);

      if (leg1 < rules.legMin || leg2 < rules.legMin || leg3 < rules.legMin) continue;
      if (leg1 > maxLegNm || leg2 > maxLegNm) continue;
      if (leg3 > maxLegNm) continue;

      const totalDistanceNm = leg1 + leg2 + leg3;
      routes.push({
        id: `triple-${a.ident}-${b.ident}`,
        flightType,
        stops: [a, b],
        legs: [
          { from: origin.ident, to: a.ident, distanceNm: leg1 },
          { from: a.ident, to: b.ident, distanceNm: leg2 },
          { from: b.ident, to: origin.ident, distanceNm: leg3 }
        ],
        totalDistanceNm,
        score: scoreRoute({ stops: [a, b], totalDistanceNm, flightType, maxLegNm, requiredApproachTypes }),
        summary: `${origin.ident} -> ${a.ident} -> ${b.ident} -> ${origin.ident}`
      });
    }
  }

  routes.sort((a, b) => b.score - a.score);
  return {
    routes,
    routeStats: {
      preTopRouteCount: routes.length
    }
  };
}

function scoreRoute({ stops, totalDistanceNm, flightType, maxLegNm, requiredApproachTypes = [] }) {
  const avgTraining = stops.reduce((sum, s) => sum + s.trainingScore, 0) / stops.length;

  const approachTypeSet = new Set();
  for (const stop of stops) {
    for (const t of stop.approachStats.typesLikely) {
      if (!INCLUDE_TACAN_FOR_T6_SCORING && t === "TACAN") continue;
      approachTypeSet.add(t);
    }
  }
  const diversityBonus = Math.min(10, approachTypeSet.size * 2.5);

  const target = flightType === "single" ? 160 : flightType === "double" ? maxLegNm * 1.8 : maxLegNm * 2.3;
  const spread = flightType === "single" ? 90 : 180;
  const distanceScore = clamp(12 - Math.abs(totalDistanceNm - target) / spread * 12, 0, 12);
  const required = requiredApproachTypes.map((t) => String(t).toUpperCase());
  const availableTypes = new Set();
  for (const stop of stops) {
    const likely = stop.approachStats.typesLikely || [];
    const alternate = stop.approachStats.typesAlternate || [];
    for (const t of [...likely, ...alternate]) {
      if (!INCLUDE_TACAN_FOR_T6_SCORING && t === "TACAN") continue;
      availableTypes.add(t);
    }
  }

  let requiredScore = 0;
  if (required.length) {
    for (const type of required) {
      if (availableTypes.has(type)) requiredScore += 9;
      else requiredScore -= 12;
    }
  }

  return clamp(Math.round(avgTraining * 0.82 + diversityBonus + distanceScore + requiredScore), 1, 100);
}

function scoreAirportTraining({ airport, likelyRunway, approachStats, suitability }) {
  const likelyTypes = INCLUDE_TACAN_FOR_T6_SCORING
    ? approachStats.typesLikely
    : approachStats.typesLikely.filter((t) => t !== "TACAN");
  const alternateTypes = INCLUDE_TACAN_FOR_T6_SCORING
    ? approachStats.typesAlternate
    : approachStats.typesAlternate.filter((t) => t !== "TACAN");

  const diversityWeights = {
    ILS: 15,
    LOC: 10,
    VOR: 10,
    RNAV: 8,
    NDB: 6,
    GCA: 7,
    CIRCLING: 5,
    TACAN: 4
  };

  let diversityRaw = 0;
  for (const type of likelyTypes) diversityRaw += diversityWeights[type] || 0;
  for (const type of alternateTypes) diversityRaw += (diversityWeights[type] || 0) * 0.35;
  const diversity = clamp(diversityRaw, 0, SCORE_WEIGHTS.diversityMax);

  const quantity = clamp(
    approachStats.countForRunway * 2.8 +
    approachStats.alternateCount * 0.8 +
    Math.min(4.5, approachStats.totalCount * 0.2),
    0,
    SCORE_WEIGHTS.quantityMax
  );

  const wind = likelyRunway
    ? clamp(
        12 + Math.min(6, likelyRunway.headwindKt * 0.6) - Math.max(0, -likelyRunway.headwindKt) * 1.8 - likelyRunway.crosswindKt * 0.55,
        0,
        SCORE_WEIGHTS.windMax
      )
    : 0;

  const runwayLength = likelyRunway?.lengthFt || 0;
  const runway = runwayLength >= 9000
    ? 10
    : runwayLength >= 7000
      ? 9
      : runwayLength >= 6000
        ? 8
        : runwayLength >= 5000
          ? 7
          : runwayLength >= 4000
            ? 6
            : 2;

  const traffic = airport.type === "large_airport"
    ? 5
    : airport.type === "medium_airport"
      ? 4
      : 2;

  const preAdjusted = diversity + quantity + wind + runway + traffic - suitability.penalty;
  const total = clamp(Math.round(preAdjusted * suitability.scoreMultiplier), 1, 100);

  return {
    total,
    breakdown: {
      diversity,
      quantity,
      wind,
      runway,
      traffic,
      suitabilityPenalty: suitability.penalty,
      suitabilityMultiplier: suitability.scoreMultiplier
    }
  };
}

function buildApproachStats(approaches, likelyRunwayIdent, allCharts = []) {
  const byRunway = likelyRunwayIdent
    ? approaches.filter((a) => chartLikelyForRunway(a.name, likelyRunwayIdent))
    : [];
  const selected = byRunway.length ? byRunway : approaches;
  const alternate = byRunway.length
    ? approaches.filter((a) => !chartLikelyForRunway(a.name, likelyRunwayIdent))
    : [];
  const radarCharts = allCharts.filter((chart) => /RADMIN|RADAR[\s\-_/]*MIN(?:IMUMS?)?/i.test(chart.name || ""));
  const hasExistingGca = selected.some((approach) => classifyApproachTypes(approach.name).includes("GCA"));
  const syntheticGcaApproaches = (radarCharts.length && !hasExistingGca)
    ? [{
        name: "GCA (Radar minimums available)",
        pdf: radarCharts.find((chart) => chart.pdf)?.pdf || null,
        source: "AviationAPI HTML"
      }]
    : [];
  const selectedWithGca = dedupeApproaches([...selected, ...syntheticGcaApproaches]);

  const likelyTypeSet = new Set();
  for (const approach of selectedWithGca) {
    classifyApproachTypes(approach.name).forEach((type) => likelyTypeSet.add(type));
  }

  const alternateTypeSet = new Set();
  for (const approach of alternate) {
    classifyApproachTypes(approach.name).forEach((type) => alternateTypeSet.add(type));
  }

  const hasRadarMinimums = radarCharts.length > 0;
  if (hasRadarMinimums) likelyTypeSet.add("GCA");

  const airportDiagram = allCharts.find((chart) => /AIRPORT\s+DIAGRAM|\bAPD\b/i.test(chart.name || ""));
  const chartSupplement = allCharts.find((chart) => /CHART\s+SUPP|SUPPLEMENT|A\/FD|AIRPORT\s+FACILITY/i.test(chart.name || ""));

  return {
    countForRunway: byRunway.length,
    alternateCount: alternate.length,
    totalCount: dedupeApproaches([...approaches, ...syntheticGcaApproaches]).length,
    typesLikely: [...likelyTypeSet].sort(),
    typesAlternate: [...alternateTypeSet].sort(),
    selectedApproaches: selectedWithGca,
    alternateApproaches: alternate,
    airportDiagramPdf: airportDiagram?.pdf || null,
    chartSupplementPdf: chartSupplement?.pdf || null,
    hasRadarMinimums,
    radarChartCount: radarCharts.length,
    chartCount: allCharts.length
  };
}

function classifyApproachTypes(name) {
  const n = String(name || "").toUpperCase();
  const out = [];

  if (n.includes("ILS")) out.push("ILS");
  if (n.includes("LOC") || n.includes("LOCALIZER") || n.includes("LDA") || n.includes("SDF")) out.push("LOC");
  if (n.includes("VOR") || n.includes("VORTAC")) out.push("VOR");
  if (n.includes("RNAV") || n.includes("GPS")) out.push("RNAV");
  if (n.includes("NDB")) out.push("NDB");
  if (n.includes("GCA") || /RADMIN|RADAR[\s\-_/]*MIN(?:IMUMS?)?/.test(n)) out.push("GCA");
  if (n.includes("TACAN")) out.push("TACAN");
  if (/\b[A-Z0-9]+-[A-Z]\b/.test(n)) out.push("CIRCLING");

  return out;
}

function evaluateAirportSuitability(airport) {
  const flags = [];
  const name = String(airport.name || "").toUpperCase();
  const keywords = String(airport.keywords || "").toUpperCase();
  const text = `${name} ${keywords}`;
  let penalty = 0;
  let scoreMultiplier = 1;
  const hasNavy = /\bNAVAL\b|\bNAVY\b|\bNAVAL AIR STATION\b|\bNAS\b/.test(text);
  const hasNonNavyMilitary = /\bAFB\b|AIR FORCE BASE|\bARMY\b|ARMY AIRFIELD|JOINT BASE|\bJB\b|\bMCAS\b|\bMILITARY\b|TEST RANGE|BOMBING|\bUSAF\b|\bUSMC\b/.test(text);

  if (hasNavy) {
    flags.push("Navy operations likely");
    penalty += 8;
    scoreMultiplier *= 0.9;
  } else if (hasNonNavyMilitary) {
    flags.push("Non-Navy military operations likely (AFB-level deweighting applied)");
    penalty += 12;
    scoreMultiplier *= 0.75;
  }

  if (/PRIVATE|PVT/.test(name)) {
    flags.push("Private-use restrictions possible");
    penalty += 10;
  }

  if (/HOLLOMAN/.test(name)) {
    flags.push("Known training/ops restrictions likely");
    penalty += 12;
  }

  if (airport.type === "small_airport" && !/INTL|INTERNATIONAL/.test(name)) {
    penalty += 1.5;
  }

  return {
    flags,
    penalty: clamp(penalty, 0, 22),
    scoreMultiplier: clamp(scoreMultiplier, 0.55, 1)
  };
}

async function getSurfaceWindKnots(airport, flightDate, timeLocal) {
  const key = `${airport.ident}|${flightDate}|${timeLocal}`;
  if (cache.weatherByKey.has(key)) return cache.weatherByKey.get(key);

  const today = new Date().toISOString().slice(0, 10);
  const isArchive = flightDate < today;
  const endpoint = isArchive ? OPEN_METEO_ARCHIVE : OPEN_METEO_FORECAST;

  const params = new URLSearchParams({
    latitude: String(airport.lat),
    longitude: String(airport.lon),
    timezone: "auto",
    hourly: "wind_speed_10m,wind_direction_10m",
    wind_speed_unit: "kn",
    start_date: flightDate,
    end_date: flightDate
  });

  if (!isArchive) {
    params.delete("start_date");
    params.delete("end_date");
    params.set("forecast_days", "16");
  }

  try {
    const json = await fetchJsonFromAny([`${endpoint}?${params.toString()}`]);
    const hourly = json.hourly;
    if (!hourly?.time?.length) throw new Error("no hourly wind returned");

    const idx = pickHourIndex(hourly.time, flightDate, timeLocal);
    if (idx < 0) throw new Error("no hour match");

    const speedRaw = Number(hourly.wind_speed_10m[idx]);
    const speedKt = clamp(Number.isFinite(speedRaw) ? speedRaw : 0, 0, 80);
    const directionDeg = normalizeHeading(hourly.wind_direction_10m[idx]);

    const wind = { speedKt, directionDeg, selectedTime: hourly.time[idx], source: "Open-Meteo 10m" };
    cache.weatherByKey.set(key, wind);
    return wind;
  } catch (_error) {
    const fallback = { speedKt: 0, directionDeg: 0, selectedTime: `${flightDate}T${timeLocal}`, source: "Fallback calm" };
    cache.weatherByKey.set(key, fallback);
    return fallback;
  }
}

function rankRunwayEnds(runways, wind) {
  const rankings = [];
  for (const runway of runways) {
    for (const end of runway.ends || []) {
      const comp = windComponents(end.heading, wind.directionDeg, wind.speedKt);
      rankings.push({
        runwayIdent: end.ident,
        heading: end.heading,
        lengthFt: runway.lengthFt || 0,
        widthFt: runway.widthFt || 0,
        headwindKt: comp.headwindKt,
        crosswindKt: comp.crosswindKt,
        highlightBand: scoreBandFromComponents(comp.headwindKt, comp.crosswindKt)
      });
    }
  }

  rankings.sort((a, b) => runwaySuitabilityPoints(b) - runwaySuitabilityPoints(a));
  return rankings;
}

function scoreBandFromComponents(headwindKt, crosswindKt) {
  if (crosswindKt <= 10 && headwindKt >= -2) return "green";
  if (crosswindKt <= 18 && headwindKt >= -8) return "yellow";
  return "red";
}

function runwaySuitabilityPoints(row) {
  return 25 + row.headwindKt * 0.8 - Math.max(0, -row.headwindKt) * 2 - row.crosswindKt * 0.9 + (row.lengthFt >= 6000 ? 3 : 0);
}

function hasRunwayAtLeast(runways, minFt) {
  return runways.some((r) => Number(r.lengthFt) >= minFt);
}

function scoreAirportCandidateForPool(airport, flightType, maxLegNm) {
  const distance = Number(airport.distanceFromOriginNm) || 0;
  const safeMax = Math.max(1, Number(maxLegNm) || 1);
  const targetRatio = flightType === "double" ? 0.55 : 0.7;
  const targetNm = safeMax * targetRatio;
  const distanceScore = clamp(20 - Math.abs(distance - targetNm) / safeMax * 20, 0, 20);

  const runways = airport.runways || [];
  const runwayCount = runways.length;
  const maxLength = runways.reduce((max, r) => Math.max(max, Number(r.lengthFt) || 0), 0);
  const runwayInfraScore = clamp(runwayCount * 1.1 + clamp((maxLength - 4000) / 450, 0, 10), 0, 16);

  const cachedApproaches = cache.approachesByAirport.get(airport.ident) || [];
  const cachedChartBundle = cache.aviationChartsByAirport.get(airport.ident);
  const cachedChartApproaches = cachedChartBundle?.approaches || [];
  const cachedCombined = dedupeApproaches([...cachedApproaches, ...cachedChartApproaches]);
  const cachedTypes = new Set();
  for (const row of cachedCombined) {
    for (const type of classifyApproachTypes(row.name)) {
      if (!INCLUDE_TACAN_FOR_T6_SCORING && type === "TACAN") continue;
      cachedTypes.add(type);
    }
  }

  const approachPotentialScore = cachedCombined.length
    ? clamp(cachedCombined.length * 0.85 + cachedTypes.size * 2.2, 0, 24)
    : 12;

  return distanceScore + runwayInfraScore + approachPotentialScore;
}

async function fetchApproaches(ident, supplementalCharts = null) {
  if (cache.approachesByAirport.has(ident)) return cache.approachesByAirport.get(ident);
  const supplemental = supplementalCharts || await fetchAviationHtmlCharts(ident);

  try {
    const faaMap = await getFaaApproachMap();
    const faaRows = faaMap.get(ident) || [];
    if (faaRows.length) {
      const merged = dedupeApproaches([...faaRows, ...supplemental.approaches]);
      cache.approachesByAirport.set(ident, merged);
      return merged;
    }
  } catch (_error) {
    // Fall through to AviationAPI.
  }

  try {
    const params = new URLSearchParams({ apt: ident, group: "6" });
    const json = await fetchJsonFromAny([`${AVIATION_API_CHARTS}?${params.toString()}`]);
    const identUpper = String(ident || "").toUpperCase();
    const identLower = identUpper.toLowerCase();
    const ident3 = identUpper.startsWith("K") ? identUpper.slice(1) : identUpper;
    const rows = json?.[identUpper] || json?.[identLower] || json?.[ident3] || json?.[ident3.toLowerCase()] || (Array.isArray(json) ? json : null);
    if (!Array.isArray(rows)) {
      cache.approachesByAirport.set(ident, []);
      return [];
    }

    const mapped = dedupeApproaches(rows
      .map((r) => ({
        name: r.chart_name || r.chartName || "Approach",
        pdf: normalizeApproachPdfUrl(r.pdf_path || r.pdfPath || null, "https://api.aviationapi.com"),
        source: "AviationAPI JSON"
      }))
      .filter((r) => r.name)
      .filter((r) => isApproachChartRow({ name: r.name, section: "APPROACH" })));
    const merged = dedupeApproaches([...mapped, ...supplemental.approaches]);
    cache.approachesByAirport.set(ident, merged);
    return merged;
  } catch (_error) {
    // Fall through.
  }

  try {
    cache.approachesByAirport.set(ident, supplemental.approaches);
    return supplemental.approaches;
  } catch (_error) {
    cache.approachesByAirport.set(ident, []);
    return [];
  }
}

async function fetchAviationHtmlCharts(ident) {
  if (cache.aviationChartsByAirport.has(ident)) return cache.aviationChartsByAirport.get(ident);

  const jsonBundle = await fetchAviationChartsJson(ident);
  if (jsonBundle.allCharts.length) {
    cache.aviationChartsByAirport.set(ident, jsonBundle);
    return jsonBundle;
  }

  const airportCode = String(ident || "").trim();
  const lowerCode = airportCode.toLowerCase();
  const urlCandidates = [
    `https://www.aviationapi.com/charts?airport=${encodeURIComponent(lowerCode)}`,
    `https://www.aviationapi.com/charts/?airport=${encodeURIComponent(lowerCode)}`,
    `https://www.aviationapi.com/charts?apt=${encodeURIComponent(airportCode)}`,
    `https://www.aviationapi.com/charts/?apt=${encodeURIComponent(airportCode)}`
  ];

  let bundle = { allCharts: [], approaches: [] };
  for (const url of urlCandidates) {
    try {
      const html = await fetchTextFromAny([url]);
      const parsed = parseAviationChartsHtml(html);
      bundle = parsed;
      if (parsed.allCharts.length) break;
    } catch (_error) {
      // try next URL candidate
    }
  }

  cache.aviationChartsByAirport.set(ident, bundle);
  return bundle;
}

async function fetchAviationChartsJson(ident) {
  const airportCode = String(ident || "").trim().toUpperCase();
  if (!airportCode) return { allCharts: [], approaches: [] };

  const urlCandidates = [
    `https://api.aviationapi.com/v1/charts?apt=${encodeURIComponent(airportCode)}`,
    `https://api.aviationapi.com/v1/charts?airport=${encodeURIComponent(airportCode)}`
  ];

  for (const url of urlCandidates) {
    try {
      const json = await fetchJsonFromAny([url]);
      const mapped = mapAviationChartsJson(json, airportCode);
      if (mapped.allCharts.length) return mapped;
    } catch (_error) {
      // try next
    }
  }

  return { allCharts: [], approaches: [] };
}

function mapAviationChartsJson(json, airportCode) {
  if (!json || typeof json !== "object") return { allCharts: [], approaches: [] };
  const airportLower = airportCode.toLowerCase();
  const airport3 = airportCode.startsWith("K") ? airportCode.slice(1) : airportCode;
  const airport3Lower = airport3.toLowerCase();
  const root = json[airportCode] || json[airportLower] || json[airport3] || json[airport3Lower] || json;
  if (!root || typeof root !== "object") return { allCharts: [], approaches: [] };

  const rows = [];

  if (Array.isArray(root)) {
    for (const item of root) {
      if (!item || typeof item !== "object") continue;
      const section = String(item?.chart_code || item?.chartCode || item?.section || item?.category || "UNKNOWN").toUpperCase();
      const name = String(item?.chart_name || item?.chartName || item?.name || "").trim();
      if (!name) continue;
      const pdf = normalizeApproachPdfUrl(item?.pdf_path || item?.pdfPath || item?.pdf || null, "https://api.aviationapi.com");
      rows.push({ name, pdf, section, source: "AviationAPI JSON" });
    }
  }

  for (const [sectionRaw, list] of Object.entries(root)) {
    if (!Array.isArray(list)) continue;
    const section = String(sectionRaw || "UNKNOWN").toUpperCase();
    for (const item of list) {
      const name = String(item?.chart_name || item?.chartName || item?.name || "").trim();
      if (!name) continue;
      const pdf = normalizeApproachPdfUrl(item?.pdf_path || item?.pdfPath || item?.pdf || null, "https://api.aviationapi.com");
      rows.push({ name, pdf, section, source: "AviationAPI JSON" });
    }
  }

  const allCharts = dedupeApproaches(rows);
  const approaches = dedupeApproaches(allCharts.filter((chart) => isApproachChartRow(chart)));
  return { allCharts, approaches };
}

function parseAviationChartsHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const tables = [...doc.querySelectorAll("table")];
  const allCharts = [];

  for (const table of tables) {
    const section = tableSectionName(table);
    for (const tr of [...table.querySelectorAll("tr")]) {
      const cells = [...tr.querySelectorAll("td")];
      if (!cells.length) continue;
      const firstLink = tr.querySelector("a[href]");
      const nameText = cells.map((c) => c.textContent.trim()).filter(Boolean).join(" | ");
      const name = nameText || firstLink?.textContent?.trim() || "";
      if (!name || name.toUpperCase().includes("CHART NAME")) continue;

      let pdf = firstLink?.getAttribute("href") || null;
      pdf = normalizeApproachPdfUrl(pdf, "https://www.aviationapi.com");
      allCharts.push({ name, pdf, section, source: "AviationAPI HTML" });
    }
  }

  const dedupedAll = dedupeApproaches(allCharts);
  const approaches = dedupeApproaches(dedupedAll.filter((chart) => isApproachChartRow(chart)));
  return { allCharts: dedupedAll, approaches };
}

async function getFaaApproachMap() {
  if (cache.faaApproachMapPromise) return cache.faaApproachMapPromise;

  cache.faaApproachMapPromise = (async () => {
    const searchHtml = await fetchTextFromAny([FAA_DTPP_SEARCH_URL]);
    const xmlUrls = uniqueMatches(searchHtml, FAA_DTPP_XML_MATCH).slice(0, 2);
    if (!xmlUrls.length) throw new Error("FAA XML links not found.");

    const map = new Map();
    for (const xmlUrl of xmlUrls) {
      const xmlText = await fetchTextFromAny([xmlUrl]);
      mergeFaaXmlIntoMap(map, xmlText, xmlUrl);
    }
    return map;
  })();

  return cache.faaApproachMapPromise;
}

function mergeFaaXmlIntoMap(map, xmlText, xmlUrl) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "application/xml");

  if (doc.querySelector("parsererror")) return;
  const cycleDir = extractCycleDirectory(xmlUrl);
  const records = [...doc.getElementsByTagName("record")];

  for (const record of records) {
    const chartCode = textOf(record, "chart_code");
    if (!isInstrumentApproachCode(chartCode)) continue;

    const chartName = textOf(record, "chart_name");
    const pdfName = textOf(record, "pdf_name");
    const aptIdent = textOf(record, "apt_ident").toUpperCase();
    const aptIcao = textOf(record, "apt_icao").toUpperCase();
    const aptCity = textOf(record, "apt_city");

    const approach = {
      name: chartName || "Approach",
      pdf: pdfName && cycleDir ? `https://aeronav.faa.gov/d-tpp/${cycleDir}/${pdfName}` : null,
      source: "FAA d-TPP",
      city: aptCity
    };

    if (aptIcao) pushApproachRow(map, aptIcao, approach);
    if (aptIdent) pushApproachRow(map, aptIdent, approach);
    if (aptIdent && aptIdent.length === 3) pushApproachRow(map, `K${aptIdent}`, approach);
  }
}

function isInstrumentApproachCode(chartCode) {
  const code = (chartCode || "").toUpperCase();
  return FAA_IAP_CODES.has(code) || code.startsWith("IAP");
}

function pushApproachRow(map, key, row) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  const rows = map.get(key);
  const exists = rows.some((existing) => existing.name === row.name && existing.pdf === row.pdf);
  if (!exists) rows.push(row);
}

function uniqueMatches(text, regex) {
  const out = [];
  const seen = new Set();
  const matches = text.match(regex) || [];
  for (let value of matches) {
    value = value.replaceAll("\\/", "/");
    if (!value.startsWith("http")) continue;
    if (!seen.has(value)) {
      seen.add(value);
      out.push(value);
    }
  }
  return out;
}

function extractCycleDirectory(xmlUrl) {
  const match = String(xmlUrl).match(/d-tpp_(\d{4})_/i);
  return match ? match[1] : null;
}

function textOf(recordNode, tagName) {
  const node = recordNode.getElementsByTagName(tagName)[0];
  return (node?.textContent || "").trim();
}

function renderAll(model) {
  const totalRoutes = model.routes.length;
  const currentVisible = Math.min(visibleRouteCount, totalRoutes);
  const needed = (model.requiredApproachTypes || []).length
    ? ` | Needed: ${(model.requiredApproachTypes || []).join(", ")}`
    : "";
  const metrics = model.candidateStats && model.routeStats
    ? `Candidates ${model.candidateStats.totalFilteredCandidates} | Profile pool ${model.candidateStats.profilePoolSizeUsed} | Generated ${model.routeStats.preTopRouteCount}`
    : "";
  const summaryHtml = `
    <article class="card border-0 shadow-sm">
      <div class="card-body py-3">
        <div class="fw-semibold">${model.origin.ident} | ${FLIGHT_RULES[model.flightType].label} | ${model.flightDate} ${model.timeLocal}</div>
        <div class="small text-secondary mt-1">
          Showing ${currentVisible} of ${totalRoutes} route options ranked by training score (1-100).
          Surface winds are Open-Meteo 10m winds in knots.${needed}
        </div>
        <div class="small text-secondary mt-1">
          ${model.usingFallback ? "Airport/runway dataset fallback enabled. " : ""}
          ${metrics}
        </div>
      </div>
    </article>
  `;

  const visibleRoutes = model.routes.slice(0, currentVisible);
  const visibleRouteIdSet = new Set(visibleRoutes.map((r) => r.id));
  for (const routeId of [...expandedRouteIds]) {
    if (!visibleRouteIdSet.has(routeId)) expandedRouteIds.delete(routeId);
  }

  const listHtml = visibleRoutes
    .map((route, index) => {
      const band = scoreBand(route.score);
      const isExpanded = expandedRouteIds.has(route.id);
      const selectedClass = isExpanded ? " rs-route-selected" : "";
      const expandedHtml = isExpanded
        ? routeDetailInlineHtml(route, model.origin, model.originProfile, mapContainerId(route))
        : "";
      return `
        <article class="card border-0 shadow-sm${selectedClass}" data-route-card-index="${index}">
          <div class="card-body">
            <div class="d-flex align-items-start justify-content-between gap-3">
              <div>
                <h3 class="h5 mb-1">Option ${index + 1}: ${route.summary}</h3>
                <div class="small text-secondary">${route.totalDistanceNm.toFixed(1)} NM total | ${route.legs.length} legs</div>
              </div>
              <div class="badge rs-score-badge rs-score-${band}">${route.score}</div>
            </div>

            <div class="vstack gap-2 my-3">
            ${route.stops.map((stop) => stopOverviewHtml(stop)).join("")}
            </div>

            <div>
            ${isExpanded
              ? `<button class="btn btn-outline-secondary btn-sm" data-route-id="${route.id}">Hide Details</button>`
              : `<button class="btn btn-outline-primary btn-sm" data-route-id="${route.id}">Show More</button>`}
            </div>
            ${expandedHtml}
          </div>
        </article>
      `;
    })
    .join("\n");
  const loadMoreHtml = currentVisible < totalRoutes
    ? `<div class="d-flex justify-content-center mt-2"><button class="btn btn-outline-primary" data-action="load-more">Load More</button></div>`
    : "";
  resultsEl.innerHTML = summaryHtml + `<section class="vstack gap-3">${listHtml}${loadMoreHtml}</section>`;

  const activeMapIds = new Set();
  for (const route of visibleRoutes) {
    if (!expandedRouteIds.has(route.id)) continue;
    const mapId = mapContainerId(route);
    activeMapIds.add(mapId);
    renderRouteMap(route, model.origin, mapId);
  }
  for (const [mapId, instance] of routeMapInstances.entries()) {
    if (activeMapIds.has(mapId)) continue;
    if (instance?.map) instance.map.remove();
    if (instance?.monitor?.timerId) clearTimeout(instance.monitor.timerId);
    routeMapInstances.delete(mapId);
  }
}

function stopOverviewHtml(stop) {
  const runway = stop.likelyRunway;
  const displayTypes = stop.approachStats.typesLikely.filter((t) => INCLUDE_TACAN_FOR_T6_SCORING || t !== "TACAN");
  const types = displayTypes.length ? displayTypes.slice(0, 4).join(", ") : "None";
  const runwayBand = runway ? runway.highlightBand : "red";
  const hasWarnings = stop.suitability.flags.length > 0;

  return `
    <div class="card bg-light-subtle border">
      <div class="card-body py-2 px-3">
        <div class="d-flex flex-wrap gap-3 align-items-center small">
        <strong>${stop.ident}</strong>
          <span class="text-secondary">Wind ${stop.wind.directionDeg.toFixed(0)}° @ ${stop.wind.speedKt.toFixed(0)} kt</span>
          <span class="text-secondary"><span class="rs-dot rs-dot-${runwayBand}"></span> RWY ${runway ? runway.runwayIdent : "N/A"}</span>
          <span class="text-secondary">Types: ${types}</span>
        </div>
        ${hasWarnings ? `<div class="small text-warning-emphasis mt-1">Caution: ${escapeHtml(stop.suitability.flags[0])}. Check IFR supplement for PPR/remarks.</div>` : ""}
      </div>
    </div>
  `;
}

function routeDetailInlineHtml(route, origin, originProfile, mapId) {
  const selectedMapStyle = mapStyleById.get(mapId) || MAP_DEFAULT_STYLE;
  const legRows = route.legs
    .map((leg) => `<li class="list-group-item d-flex justify-content-between"><span>${leg.from} -> ${leg.to}</span><span>${leg.distanceNm.toFixed(1)} NM</span></li>`)
    .join("");

  const focusCards = route.stops
    .map((airport, index) => {
      return airfieldSummaryCardHtml(airport, `Destination ${index + 1}`);
    })
    .join("");

  return `
    <div class="mt-3 pt-3 border-top rs-detail-shell">
      <h3 class="h5 mb-3">Selected Route Detail</h3>
      <div class="card border rs-detail-panel">
        <div class="card-body">
          <div class="row g-3">
            <div class="col-lg-4">
              <h5 class="mb-2">Mission Plan</h5>
              <ul class="list-group list-group-flush rs-leg-list">${legRows}</ul>
            </div>
            <div class="col-lg-8">
              <h5 class="mb-2">Airfield Summary</h5>
              <div class="row g-2">
                ${airfieldSummaryCardHtml(originProfile, "Home Field")}
                ${focusCards}
              </div>
            </div>
          </div>
          <div class="mt-3 pt-3 border-top">
            <div class="d-flex justify-content-between align-items-center gap-2 flex-wrap mb-2">
              <div class="fw-semibold">Route Quick Map</div>
              <div class="d-flex align-items-center gap-2">
                <label class="small text-secondary mb-0" for="${mapId}-style">Map</label>
                <select id="${mapId}-style" class="form-select form-select-sm rs-map-style-select" data-map-id="${mapId}">
                  ${MAP_STYLE_OPTIONS.map((opt) => `<option value="${opt.value}" ${opt.value === selectedMapStyle ? "selected" : ""}>${opt.label}</option>`).join("")}
                </select>
              </div>
            </div>
            <div id="${mapId}" class="rs-route-map border rounded-3"></div>
            <div id="${mapId}-note" class="small text-secondary mt-1"></div>
          </div>
        </div>
      </div>

      <div class="vstack gap-3 mt-3">
        <section class="card border rs-airfield-panel">
          <div class="card-body">
            <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-2">
              <h4 class="h6 mb-0">Home Field Detail: ${originProfile.ident}</h4>
              <span class="small text-secondary">${originProfile.name ? escapeHtml(originProfile.name) : ""}</span>
            </div>
            ${airfieldDetailCardHtml(originProfile, "home", `${mapId}-home`)}
          </div>
        </section>
        ${route.stops.map((stop, idx) => `
          <section class="card border rs-airfield-panel">
            <div class="card-body">
              <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-2">
                <h4 class="h6 mb-0">Destination ${idx + 1}: ${stop.ident}</h4>
                <span class="small text-secondary">${stop.name ? escapeHtml(stop.name) : ""}</span>
              </div>
              ${airfieldDetailCardHtml(stop, `destination-${idx + 1}`, `${mapId}-dest-${idx}`)}
            </div>
          </section>
        `).join("")}
      </div>
    </div>
  `;
}

function airfieldSummaryCardHtml(airport, label) {
  const runway = airport.likelyRunway;
  const band = runway ? runway.highlightBand : "red";
  const approachTypes = airport.approachStats.typesLikely
    .filter((t) => INCLUDE_TACAN_FOR_T6_SCORING || t !== "TACAN")
    .slice(0, 5)
    .join(", ");
  return `
    <div class="col-md-6">
      <div class="card h-100 border">
        <div class="card-body py-2">
          <div class="small text-uppercase text-secondary">${label}</div>
          <div><strong>${airport.ident}</strong> ${airport.name ? `- ${escapeHtml(airport.name)}` : ""}</div>
          <div class="small text-secondary"><span class="rs-dot rs-dot-${band}"></span> ${airport.wind.directionDeg.toFixed(0)}° @ ${airport.wind.speedKt.toFixed(0)} kt</div>
          <div class="small text-secondary">Likely RWY ${runway ? runway.runwayIdent : "N/A"} ${runway ? `(${runway.lengthFt}x${runway.widthFt || "?"} ft)` : ""}</div>
          <div class="small text-secondary">Approach set: ${approachTypes || "None detected"}</div>
        </div>
      </div>
    </div>
  `;
}

function airfieldDetailCardHtml(stop, title, keySuffix = "") {
  const likelyApproaches = stop.approachStats.selectedApproaches;
  const alternateApproaches = stop.approachStats.alternateApproaches;
  const runwayBand = stop.likelyRunway ? stop.likelyRunway.highlightBand : "red";
  const approachSources = [...new Set((stop.approaches || []).map((a) => a.source).filter(Boolean))];
  const filteredLikely = likelyApproaches.filter((a) => INCLUDE_TACAN_FOR_T6_SCORING || !classifyApproachTypes(a.name).includes("TACAN"));
  const filteredAlternate = alternateApproaches.filter((a) => INCLUDE_TACAN_FOR_T6_SCORING || !classifyApproachTypes(a.name).includes("TACAN"));
  const runwaysList = (stop.runways || [])
    .slice(0, 8)
    .map((r) => {
      const ends = (r.ends || []).map((e) => e.ident).filter(Boolean).join("/");
      return `<li>RWY ${escapeHtml(ends || "N/A")} - ${Math.round(r.lengthFt || 0)}x${Math.round(r.widthFt || 0)} ft</li>`;
    })
    .join("");
  const towered = inferTowered(stop);
  const hasLighting = (stop.runways || []).some((r) => r.lighted);
  const tabId = `airfield-tabs-${String(stop.ident || "apt").toLowerCase()}-${String(keySuffix || title).replace(/[^a-zA-Z0-9_-]/g, "-")}`;
  const primaryTab = `${tabId}-primary`;
  const alternateTab = `${tabId}-alternate`;
  const opsTab = `${tabId}-airfield-ops`;
  const scoringTab = `${tabId}-scoring`;
  const likelyRunwayText = stop.likelyRunway ? `${stop.likelyRunway.runwayIdent}` : "N/A";
  const airportIdent = String(stop.ident || "").trim().toLowerCase();
  const chartsUrl = airportIdent
    ? `https://www.aviationapi.com/charts?airport=${encodeURIComponent(airportIdent)}`
    : "https://www.aviationapi.com/charts";

  return `
    <div>
      <div class="mb-2 d-md-none">
        <label class="form-label small text-secondary mb-1" for="${tabId}-picker">Section</label>
        <select id="${tabId}-picker" class="form-select form-select-sm rs-airfield-section-picker" data-tab-root="${tabId}">
          <option value="${primaryTab}" selected>Primary Approaches</option>
          <option value="${alternateTab}">Other Approaches</option>
          <option value="${opsTab}">Airfield &amp; Ops Notes</option>
          <option value="${scoringTab}">Scoring &amp; Source</option>
        </select>
      </div>

      <ul class="nav nav-tabs rs-airfield-tabs d-none d-md-flex" id="${tabId}" role="tablist">
        <li class="nav-item" role="presentation">
          <button class="nav-link active" id="${primaryTab}-tab" data-bs-toggle="tab" data-bs-target="#${primaryTab}" type="button" role="tab" aria-controls="${primaryTab}" aria-selected="true">Primary Approaches</button>
        </li>
        <li class="nav-item" role="presentation">
          <button class="nav-link" id="${alternateTab}-tab" data-bs-toggle="tab" data-bs-target="#${alternateTab}" type="button" role="tab" aria-controls="${alternateTab}" aria-selected="false">Other Approaches</button>
        </li>
        <li class="nav-item" role="presentation">
          <button class="nav-link" id="${opsTab}-tab" data-bs-toggle="tab" data-bs-target="#${opsTab}" type="button" role="tab" aria-controls="${opsTab}" aria-selected="false">Airfield & Ops Notes</button>
        </li>
        <li class="nav-item" role="presentation">
          <button class="nav-link" id="${scoringTab}-tab" data-bs-toggle="tab" data-bs-target="#${scoringTab}" type="button" role="tab" aria-controls="${scoringTab}" aria-selected="false">Scoring & Source</button>
        </li>
      </ul>
      <div class="tab-content border border-top-0 rounded-bottom p-3 rs-tab-content">
        <div class="tab-pane fade show active" id="${primaryTab}" role="tabpanel" aria-labelledby="${primaryTab}-tab" tabindex="0">
          ${filteredLikely.length
            ? `<ul class="list-group list-group-flush">${filteredLikely.map((a) => `<li class="list-group-item px-0">${approachLinkHtml(a)}</li>`).join("")}</ul>`
            : "<div class=\"small text-secondary\">No approach chart metadata returned.</div>"}
        </div>
        <div class="tab-pane fade" id="${alternateTab}" role="tabpanel" aria-labelledby="${alternateTab}-tab" tabindex="0">
          ${filteredAlternate.length
            ? `<ul class="list-group list-group-flush">${filteredAlternate.map((a) => `<li class="list-group-item px-0">${approachLinkHtml(a)}</li>`).join("")}</ul>`
            : "<div class=\"small text-secondary\">No additional approaches found.</div>"}
        </div>
        <div class="tab-pane fade" id="${opsTab}" role="tabpanel" aria-labelledby="${opsTab}-tab" tabindex="0">
          <div class="d-flex flex-wrap gap-2 mb-2 small text-secondary">
            <span class="badge text-bg-light border"><span class="rs-dot rs-dot-${runwayBand}"></span> Training Score ${stop.trainingScore}</span>
            <span class="badge text-bg-light border">Wind ${stop.wind.directionDeg.toFixed(0)}° @ ${stop.wind.speedKt.toFixed(0)} kt</span>
            <span class="badge text-bg-light border">Likely RWY ${likelyRunwayText}</span>
          </div>
          <div class="small text-secondary">Towered: ${towered ? "Yes" : "No"}</div>
          <div class="small text-secondary">Lighting: ${hasLighting ? "Available" : "Unknown/Unlit"} | Elev: ${stop.elevationFt ?? "N/A"} ft</div>
          <div class="small text-secondary">Service: ${stop.scheduledService ? "Scheduled" : "Non-scheduled"}</div>
          <div class="small mt-2 fw-semibold">Available Runways</div>
          <ul class="small mb-2">${runwaysList || "<li>No runway records</li>"}</ul>
          ${stop.approachStats.hasRadarMinimums ? "<div class=\"small text-secondary\">Radar minimums published: GCA training capability available.</div>" : ""}
          <div class="small mt-2">
            ${stop.approachStats.airportDiagramPdf ? `<a href="${stop.approachStats.airportDiagramPdf}" target="_blank" rel="noopener noreferrer">Airport Diagram</a>` : "Airport Diagram: Unavailable"} |
            ${stop.approachStats.chartSupplementPdf
              ? `<a href="${stop.approachStats.chartSupplementPdf}" target="_blank" rel="noopener noreferrer">Chart Supplement PDF</a>`
              : `<a href="${chartsUrl}" target="_blank" rel="noopener noreferrer">Chart Supplement / AviationAPI Charts</a>`}
          </div>
          ${stop.suitability.flags.length ? `<div class="small text-warning-emphasis mt-2">Suitability flags: ${escapeHtml(stop.suitability.flags.join("; "))}. Check IFR supplement for PPR/remarks.</div>` : ""}
        </div>
        <div class="tab-pane fade" id="${scoringTab}" role="tabpanel" aria-labelledby="${scoringTab}-tab" tabindex="0">
          <div class="small text-secondary">Score: diversity ${Math.round(stop.scoreBreakdown.diversity)}, quantity ${Math.round(stop.scoreBreakdown.quantity)}, wind ${Math.round(stop.scoreBreakdown.wind)}, runway ${Math.round(stop.scoreBreakdown.runway)}, suitability penalty ${Math.round(stop.scoreBreakdown.suitabilityPenalty)}.</div>
          <div class="small text-secondary mt-1">Approach data source: ${approachSources.length ? approachSources.join(", ") : "Unavailable"}</div>
          <div class="small text-secondary mt-1">Chart metadata: ${stop.approachStats.chartCount ?? 0} charts | Radar mins detected: ${stop.approachStats.hasRadarMinimums ? "Yes" : "No"}</div>
          ${!INCLUDE_TACAN_FOR_T6_SCORING ? "<div class=\"small text-secondary mt-1\">TACAN approaches are available in backend data but excluded from T-6 scoring/output.</div>" : ""}
        </div>
      </div>
    </div>
  `;
}

function activateAirfieldSection(tabRootId, paneId) {
  if (!tabRootId || !paneId) return;
  const tabRoot = document.getElementById(tabRootId);
  const targetPane = document.getElementById(paneId);
  if (!targetPane) return;

  const targetButton = tabRoot?.querySelector(`[data-bs-target="#${paneId}"]`);
  if (targetButton && window.bootstrap?.Tab) {
    window.bootstrap.Tab.getOrCreateInstance(targetButton).show();
  } else {
    const paneContainer = targetPane.closest(".tab-content");
    paneContainer?.querySelectorAll(".tab-pane").forEach((pane) => pane.classList.remove("show", "active"));
    targetPane.classList.add("show", "active");

    tabRoot?.querySelectorAll(".nav-link").forEach((link) => {
      link.classList.remove("active");
      link.setAttribute("aria-selected", "false");
    });
    if (targetButton) {
      targetButton.classList.add("active");
      targetButton.setAttribute("aria-selected", "true");
    }
  }
}

function scoreBand(score) {
  if (score >= 75) return "green";
  if (score >= 55) return "yellow";
  return "red";
}

function windComponents(runwayHeading, windDirection, windSpeedKt) {
  const diff = degToRad(smallestAngularDiff(runwayHeading, windDirection));
  return {
    headwindKt: windSpeedKt * Math.cos(diff),
    crosswindKt: Math.abs(windSpeedKt * Math.sin(diff))
  };
}

function smallestAngularDiff(a, b) {
  let d = Math.abs(Number(a) - Number(b)) % 360;
  if (d > 180) d = 360 - d;
  return d;
}

function pickHourIndex(hourlyTimes, flightDate, timeLocal) {
  const [targetH, targetM] = timeLocal.split(":").map(Number);
  let bestIndex = -1;
  let bestDelta = Infinity;

  for (let i = 0; i < hourlyTimes.length; i += 1) {
    const iso = String(hourlyTimes[i]);
    if (iso.slice(0, 10) !== flightDate) continue;

    const dt = new Date(iso);
    const delta = Math.abs((dt.getHours() - targetH) * 60 + (dt.getMinutes() - targetM));
    if (delta < bestDelta) {
      bestDelta = delta;
      bestIndex = i;
    }
  }

  return bestIndex;
}

function chartLikelyForRunway(chartName, runwayIdent) {
  const rw = normalizeRunwayIdent(runwayIdent);
  if (!rw) return false;

  const chartRunways = extractRunwayMentions(chartName);
  if (!chartRunways.length) return false;

  return chartRunways.includes(rw);
}

function approachLinkHtml(approach) {
  const label = escapeHtml(approach.name);
  if (!approach.pdf) return label;
  return `<a href="${approach.pdf}" target="_blank" rel="noopener noreferrer">${label}</a>`;
}

function mapContainerId(route) {
  return `route-map-${String(route.id).replace(/[^a-zA-Z0-9_-]/g, "-")}`;
}

function renderRouteMap(route, origin, mapId) {
  const mapEl = document.getElementById(mapId);
  if (!mapEl || typeof window.L === "undefined") return;

  const existing = routeMapInstances.get(mapId);
  if (existing?.map) {
    if (existing?.monitor?.timerId) clearTimeout(existing.monitor.timerId);
    existing.map.remove();
    routeMapInstances.delete(mapId);
  }

  const points = [origin, ...route.stops, origin]
    .filter((p) => Number.isFinite(p.lat) && Number.isFinite(p.lon))
    .map((p) => [p.lat, p.lon]);
  if (!points.length) return;

  const map = window.L.map(mapEl, {
    zoomControl: true,
    scrollWheelZoom: false,
    zoomAnimation: false,
    fadeAnimation: false,
    markerZoomAnimation: false
  });
  const streetLayer = window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 13,
    attribution: "&copy; OpenStreetMap contributors",
    detectRetina: true,
    updateWhenZooming: false
  });
  const ifrLowLayer = window.L.tileLayer(IFR_LOW_TILE_URL, {
    maxZoom: 13,
    minNativeZoom: 7,
    maxNativeZoom: 12,
    attribution: "FAA/Esri IFR Low",
    crossOrigin: true,
    detectRetina: true,
    updateWhenZooming: false
  });
  const ifrHighLayer = window.L.tileLayer(IFR_HIGH_TILE_URL, {
    maxZoom: 13,
    minNativeZoom: 7,
    maxNativeZoom: 12,
    attribution: "FAA/Esri IFR High",
    crossOrigin: true,
    detectRetina: true,
    updateWhenZooming: false
  });
  const noteEl = document.getElementById(`${mapId}-note`);

  const ctx = {
    mapId,
    map,
    noteEl,
    layers: {
      [MAP_STYLE_STREET]: streetLayer,
      [MAP_STYLE_IFR_LOW]: ifrLowLayer,
      [MAP_STYLE_IFR_HIGH]: ifrHighLayer
    },
    activeStyle: MAP_STYLE_STREET,
    monitor: {
      style: MAP_STYLE_STREET,
      loadCount: 0,
      errorCount: 0,
      timerId: null,
      fallbackDone: false
    }
  };
  ifrLowLayer.on("tileerror", () => onIfrTileError(ctx, MAP_STYLE_IFR_LOW));
  ifrHighLayer.on("tileerror", () => onIfrTileError(ctx, MAP_STYLE_IFR_HIGH));
  ifrLowLayer.on("tileload", () => onIfrTileLoad(ctx, MAP_STYLE_IFR_LOW));
  ifrHighLayer.on("tileload", () => onIfrTileLoad(ctx, MAP_STYLE_IFR_HIGH));

  streetLayer.addTo(map);

  const polyline = window.L.polyline(points, { color: "#0a5d8f", weight: 4 }).addTo(map);
  window.L.marker([origin.lat, origin.lon]).addTo(map).bindPopup(`Origin: ${origin.ident}`);

  for (const stop of route.stops) {
    window.L.marker([stop.lat, stop.lon]).addTo(map).bindPopup(`${stop.ident} - ${stop.name || "Stop"}`);
  }

  map.fitBounds(polyline.getBounds(), { padding: [18, 18] });
  routeMapInstances.set(mapId, ctx);
  const desiredStyle = mapStyleById.get(mapId) || MAP_DEFAULT_STYLE;
  applyMapStyle(ctx, desiredStyle, true);
  setTimeout(() => map.invalidateSize(), 0);
}

function applyMapStyle(ctx, style, persistSelection) {
  const chosenStyle = ctx.layers[style] ? style : MAP_DEFAULT_STYLE;
  if (persistSelection) mapStyleById.set(ctx.mapId, chosenStyle);

  const streetLayer = ctx.layers[MAP_STYLE_STREET];
  const ifrLowLayer = ctx.layers[MAP_STYLE_IFR_LOW];
  const ifrHighLayer = ctx.layers[MAP_STYLE_IFR_HIGH];

  if (streetLayer && !ctx.map.hasLayer(streetLayer)) {
    streetLayer.addTo(ctx.map);
  }
  if (ifrLowLayer && ctx.map.hasLayer(ifrLowLayer)) ctx.map.removeLayer(ifrLowLayer);
  if (ifrHighLayer && ctx.map.hasLayer(ifrHighLayer)) ctx.map.removeLayer(ifrHighLayer);

  if (chosenStyle !== MAP_STYLE_STREET) {
    const overlayLayer = ctx.layers[chosenStyle];
    if (overlayLayer && !ctx.map.hasLayer(overlayLayer)) overlayLayer.addTo(ctx.map);
    streetLayer?.setOpacity?.(0.28);
    ifrLowLayer?.setOpacity?.(1);
    ifrHighLayer?.setOpacity?.(1);
  } else {
    streetLayer?.setOpacity?.(1);
    ifrLowLayer?.setOpacity?.(1);
    ifrHighLayer?.setOpacity?.(1);
  }

  resetIfrMonitor(ctx, chosenStyle);
  setMapNote(ctx, "");
  ctx.activeStyle = chosenStyle;
}

function resetIfrMonitor(ctx, style) {
  if (!ctx.monitor) return;
  if (ctx.monitor.timerId) clearTimeout(ctx.monitor.timerId);
  ctx.monitor.style = style;
  ctx.monitor.loadCount = 0;
  ctx.monitor.errorCount = 0;
  ctx.monitor.timerId = null;
  ctx.monitor.fallbackDone = false;

  if (style === MAP_STYLE_STREET) return;

  ctx.monitor.timerId = setTimeout(() => {
    ctx.monitor.timerId = null;
    if (ctx.activeStyle !== style || ctx.monitor.style !== style || ctx.monitor.fallbackDone) return;

    const shouldFallback =
      ctx.monitor.errorCount >= MAP_IFR_MIN_ERROR_THRESHOLD &&
      ctx.monitor.loadCount <= MAP_IFR_MAX_SUCCESS_FOR_FALLBACK;
    if (!shouldFallback) return;

    ctx.monitor.fallbackDone = true;
    applyMapStyle(ctx, MAP_STYLE_STREET, true);
    const select = document.querySelector(`.rs-map-style-select[data-map-id="${ctx.mapId}"]`);
    if (select) select.value = MAP_STYLE_STREET;
    setMapNote(ctx, "IFR tiles unavailable for this view. Reverted to Street.");
  }, MAP_IFR_FALLBACK_WINDOW_MS);
}

function onIfrTileLoad(ctx, style) {
  if (!ctx.monitor || ctx.activeStyle !== style || ctx.monitor.style !== style) return;
  ctx.monitor.loadCount += 1;
}

function onIfrTileError(ctx, style) {
  if (!ctx.monitor || ctx.activeStyle !== style || ctx.monitor.style !== style) return;
  ctx.monitor.errorCount += 1;
}

function setMapNote(ctx, text) {
  if (!ctx?.noteEl) return;
  ctx.noteEl.textContent = text || "";
}

function dedupeApproaches(rows) {
  const out = [];
  const seen = new Set();
  for (const row of rows) {
    const key = `${row.name}|${row.pdf || ""}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(row);
  }
  return out;
}

function tableSectionName(tableEl) {
  let prev = tableEl.previousElementSibling;
  while (prev) {
    const text = (prev.textContent || "").trim();
    if (text) return text.toUpperCase();
    prev = prev.previousElementSibling;
  }
  return "UNKNOWN";
}

function isApproachChartRow(chart) {
  const section = String(chart.section || "");
  const name = String(chart.name || "").toUpperCase();
  if (isNonApproachName(name)) return false;
  if (section.includes("APPROACH")) return looksLikeApproachProcedure(name);
  return looksLikeApproachProcedure(name);
}

function isNonApproachName(nameUpper) {
  return /\bARRIVAL\b|\bDEPARTURE\b|\bSID\b|\bSTAR\b|\bTRANSITION\b|\bENROUTE\b|\bROUTE\b|\bODP\b|\bDP\b/.test(nameUpper);
}

function looksLikeApproachProcedure(nameUpper) {
  if (/RWY\s*\d{1,2}[LRC]?/.test(nameUpper)) return true;
  if (/\bILS\b|\bLOC\b|\bLOCALIZER\b|\bLDA\b|\bSDF\b|\bVOR\b|\bNDB\b|\bTACAN\b/.test(nameUpper)) return true;
  if ((/\bRNAV\b|\bGPS\b/.test(nameUpper)) && /\bRWY\b/.test(nameUpper)) return true;
  if (/\b[A-Z0-9]+-[A-Z]\b/.test(nameUpper)) return true;
  return false;
}

function extractRunwayMentions(chartName) {
  const text = String(chartName || "").toUpperCase();
  const result = [];
  const regex = /RWY\s+(\d{1,2}[LRC]?)/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    const normalized = normalizeRunwayIdent(match[1]);
    if (normalized) result.push(normalized);
  }
  return [...new Set(result)];
}

function inferTowered(airport) {
  if (airport.type === "large_airport" || airport.type === "medium_airport") return true;
  const text = `${airport.name || ""} ${airport.keywords || ""}`.toUpperCase();
  return /TOWER|ATCT/.test(text);
}

function normalizeAirportCode(value, isoCountry = "US") {
  const raw = String(value || "").trim().toUpperCase();
  if (!raw) return null;
  const lettersDigits = raw.replace(/[^A-Z0-9]/g, "");
  if (!lettersDigits) return null;

  if (/^[A-Z]{4}$/.test(lettersDigits)) return lettersDigits;
  if (isoCountry === "US" && /^[A-Z0-9]{3}$/.test(lettersDigits)) return `K${lettersDigits}`;
  return null;
}

function normalizeRunwayIdent(value) {
  if (!value) return null;
  const raw = String(value).trim().toUpperCase();
  const match = raw.match(/^(\d{1,2})([LRC])?$/);
  if (!match) return null;
  const num = Number(match[1]);
  if (!Number.isFinite(num) || num < 1 || num > 36) return null;
  const side = match[2] || "";
  return `${String(num).padStart(2, "0")}${side}`;
}

function normalizeHeading(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return (n + 360) % 360;
}

function haversineNm(lat1, lon1, lat2, lon2) {
  const Rkm = 6371;
  const dLat = degToRad(lat2 - lat1);
  const dLon = degToRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(degToRad(lat1)) * Math.cos(degToRad(lat2)) * Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Rkm * c * 0.539957;
}

function degToRad(deg) {
  return (deg * Math.PI) / 180;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

async function mapWithConcurrency(items, limit, mapper) {
  const results = new Array(items.length);
  let nextIndex = 0;

  async function worker() {
    while (true) {
      const i = nextIndex;
      nextIndex += 1;
      if (i >= items.length) return;
      results[i] = await mapper(items[i], i);
    }
  }

  const workers = [];
  const count = Math.max(1, Math.min(limit, items.length));
  for (let i = 0; i < count; i += 1) workers.push(worker());
  await Promise.all(workers);
  return results;
}

function shuffleCopy(list) {
  const arr = [...list];
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

async function fetchTextFromAny(urls, options = {}) {
  const { validateText = null } = options;
  let lastError = null;
  const expanded = urls.flatMap((url) => buildFetchCandidates(url));

  for (const url of expanded) {
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const text = await response.text();
      if (validateText && !validateText(text)) {
        throw new Error("Response failed validation");
      }
      return text;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("All fetches failed");
}

async function fetchJsonFromAny(urls) {
  let lastError = null;
  const expanded = urls.flatMap((url) => buildFetchCandidates(url));

  for (const url of expanded) {
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const text = await response.text();
      return JSON.parse(text);
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Invalid JSON response");
}

function buildFetchCandidates(url) {
  const out = [];
  if (!String(url).startsWith("http")) return [url];

  const isLocal = LOCAL_HOSTS.has(window.location.hostname);
  const encoded = encodeURIComponent(url);

  // Local static hosting (python http.server) has no /proxy route.
  // Use public proxy candidates for external hosts while developing locally.
  if (isLocal) {
    out.push(`https://corsproxy.io/?${encoded}`);
    out.push(`https://corsproxy.io/?${url}`);
    out.push(`https://api.allorigins.win/raw?url=${encoded}`);
    const forJina = url.replace(/^https?:\/\//i, "");
    out.push(`https://r.jina.ai/http://${forJina}`);
    out.push(url);
    return out;
  }

  // In deployed environments, go through same-origin proxy first to avoid browser CORS failures.
  out.push(`${PROXY_PATH}?url=${encoded}`);
  out.push(url);
  return out;
}

function normalizeApproachPdfUrl(value, baseUrl = "") {
  const raw = String(value || "").trim();
  if (!raw) return null;
  if (/^https?:\/\//i.test(raw)) return raw;

  const cleanedBase = String(baseUrl || "").replace(/\/+$/, "");
  const cleanedValue = raw.replace(/^\/+/, "");
  if (!cleanedBase) return null;
  if (!cleanedValue) return null;

  return `${cleanedBase}/${cleanedValue}`;
}

function looksLikeCsvWithHeaders(text, requiredHeaders = []) {
  const sample = String(text || "").slice(0, 1200);
  if (!sample || sample.length < 40) return false;
  if (/^\s*</.test(sample)) return false;

  const headerLine = sample.split(/\r?\n/, 1)[0] || "";
  const normalized = headerLine
    .replace(/^\uFEFF/, "")
    .split(",")
    .map((v) => String(v).trim().replace(/^"|"$/g, ""));

  if (!normalized.length) return false;
  return requiredHeaders.every((header) => normalized.includes(header));
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        value += '"';
        i += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        value += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(value);
      value = "";
    } else if (char === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else if (char !== "\r") {
      value += char;
    }
  }

  if (value.length || row.length) {
    row.push(value);
    rows.push(row);
  }

  const headers = rows.shift() || [];
  return rows.map((cells) => {
    const obj = {};
    for (let i = 0; i < headers.length; i += 1) obj[headers[i]] = cells[i] ?? "";
    return obj;
  });
}

function escapeHtml(input) {
  return String(input)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

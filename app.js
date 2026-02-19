const AIRPORTS_URLS = [
  "https://raw.githubusercontent.com/davidmegginson/ourairports-data/main/airports.csv",
  "https://cdn.jsdelivr.net/gh/davidmegginson/ourairports-data@main/airports.csv"
];
const RUNWAYS_URLS = [
  "https://raw.githubusercontent.com/davidmegginson/ourairports-data/main/runways.csv",
  "https://cdn.jsdelivr.net/gh/davidmegginson/ourairports-data@main/runways.csv"
];
const OPEN_METEO_FORECAST = "https://api.open-meteo.com/v1/forecast";
const OPEN_METEO_ARCHIVE = "https://archive-api.open-meteo.com/v1/archive";
const AVIATION_API_CHARTS = "https://api.aviationapi.com/v1/charts";
const FAA_DTPP_SEARCH_URL = "https://www.faa.gov/air_traffic/flight_info/aeronav/digital_products/dtpp/search/";
const FAA_DTPP_XML_MATCH = /https?:\\?\/\\?\/aeronav\.faa\.gov\\?\/upload_[^"'\s]+d-tpp_[^"'\s]+_Metafile\.xml/gi;
const FAA_IAP_CODES = new Set(["IAP", "IAPMIN", "IAPCOPTER", "IAPMIL"]);
const APP_VERSION = "0.0.7";
const VERSION_FILE_PATH = "version.json";
const LOCAL_HOSTS = new Set(["localhost", "127.0.0.1", "::1"]);
const PROFILE_BUILD_CONCURRENCY = 6;

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
let selectedRouteIndex = 0;
let routeMapInstance = null;
let loadingTickerId = null;
let loadingStartedAt = 0;
let loadingQueue = [];
let loadingLongWaitQueue = [];

const LOADING_PHRASES = [
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

const LONG_WAIT_PHRASES = [
  "This is taking longer than planned. We are still working on your route.",
  "Still crunching. Sorry this is taking so long to do your own work for you.",
  "Long load detected. The planner is busy overthinking every approach option.",
  "Yes, it is still loading. No, the airplane isn't going to file itself."
];

dateInput.value = new Date().toISOString().slice(0, 10);
initVersionTicker();
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
  const button = event.target.closest("button[data-route-index]");
  if (!button || !lastModel) return;
  selectedRouteIndex = Number(button.dataset.routeIndex);
  renderAll(lastModel);
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
    const profiles = await buildAirportProfiles({
      origin,
      airports,
      runwaysByAirport,
      flightType,
      flightDate,
      timeLocal,
      maxLegNm
    });

    setStatus("Generating and scoring multiple route options...", false, true);
    const routes = generateRoutes({ origin, profiles, flightType, maxLegNm, requiredApproachTypes });
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
      routes
    };

    selectedRouteIndex = 0;
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
  statusEl.classList.toggle("error", Boolean(isError));
  statusEl.classList.toggle("loading", Boolean(isLoading));
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
  loadingQueue = shuffleCopy(LOADING_PHRASES);
  loadingLongWaitQueue = shuffleCopy(LONG_WAIT_PHRASES);
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

async function loadCoreData() {
  if (cache.airports && cache.runwaysByAirport) {
    return { airports: cache.airports, runwaysByAirport: cache.runwaysByAirport, usingFallback: false };
  }

  let usingFallback = false;
  let airportRows;
  let runwayRows;

  try {
    const [airportsCsv, runwaysCsv] = await Promise.all([
      fetchTextFromAny(AIRPORTS_URLS),
      fetchTextFromAny(RUNWAYS_URLS)
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

  const airports = airportRows
    .filter((row) => {
      const isUS = row.iso_country === "US";
      const typeOk = ["small_airport", "medium_airport", "large_airport"].includes(row.type);
      const ident = row.ident || "";
      return isUS && typeOk && ident.startsWith("K");
    })
    .map((row) => ({
      ident: row.ident,
      name: row.name,
      lat: Number(row.latitude_deg),
      lon: Number(row.longitude_deg),
      type: row.type,
      elevationFt: Number(row.elevation_ft) || null,
      scheduledService: String(row.scheduled_service || "").toLowerCase() === "yes",
      keywords: String(row.keywords || "")
    }))
    .filter((a) => Number.isFinite(a.lat) && Number.isFinite(a.lon));

  const validAirportIdents = new Set(airports.map((a) => a.ident));
  const runwaysByAirport = new Map();

  for (const row of runwayRows) {
    const ident = row.airport_ident;
    if (!validAirportIdents.has(ident)) continue;
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

  if (!candidates.length) return [];

  // Limit approach fetch volume to most relevant nearby airports.
  const trimmed = candidates
    .sort((a, b) => a.distanceFromOriginNm - b.distanceFromOriginNm)
    .slice(0, flightType === "triple" ? 35 : 45);

  return mapWithConcurrency(trimmed, PROFILE_BUILD_CONCURRENCY, (airport) => buildAirportProfile({
    airport,
    origin,
    runwaysByAirport,
    flightDate,
    timeLocal
  }));
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
    const routes = profiles
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
      .sort((a, b) => b.score - a.score)
      .slice(0, 10);

    return routes;
  }

  const top = [...profiles].sort((a, b) => b.trainingScore - a.trainingScore).slice(0, 22);
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

  return routes.sort((a, b) => b.score - a.score).slice(0, 10);
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

  const likelyTypeSet = new Set();
  for (const approach of selected) {
    classifyApproachTypes(approach.name).forEach((type) => likelyTypeSet.add(type));
  }

  const alternateTypeSet = new Set();
  for (const approach of alternate) {
    classifyApproachTypes(approach.name).forEach((type) => alternateTypeSet.add(type));
  }

  const hasRadarMinimums = allCharts.some((chart) => /RADAR\s+MINIMUMS/i.test(chart.name || ""));
  if (hasRadarMinimums) alternateTypeSet.add("GCA");

  const airportDiagram = allCharts.find((chart) => /AIRPORT\s+DIAGRAM/i.test(chart.name || ""));

  return {
    countForRunway: byRunway.length,
    alternateCount: alternate.length,
    totalCount: approaches.length,
    typesLikely: [...likelyTypeSet].sort(),
    typesAlternate: [...alternateTypeSet].sort(),
    selectedApproaches: selected,
    alternateApproaches: alternate,
    airportDiagramPdf: airportDiagram?.pdf || null,
    hasRadarMinimums
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
  if (n.includes("TACAN")) out.push("TACAN");
  if (/\b[A-Z0-9]+-[A-Z]\b/.test(n)) out.push("CIRCLING");

  return out;
}

function evaluateAirportSuitability(airport) {
  const flags = [];
  const name = String(airport.name || "").toUpperCase();
  let penalty = 0;
  let scoreMultiplier = 1;

  if (/AFB|AIR FORCE BASE/.test(name)) {
    flags.push("Air Force Base operations likely");
    penalty += 12;
    scoreMultiplier *= 0.75;
  } else if (/ARMY AIRFIELD|NAVAL AIR STATION|NAS |MCAS|MILITARY|TEST RANGE|BOMBING/.test(name)) {
    flags.push("Military operations likely");
    penalty += 12;
    scoreMultiplier *= 0.85;
  }

  if (/JOINT BASE|JB /.test(name)) {
    flags.push("Joint-base operations likely");
    penalty += 8;
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
    const rows = json[ident];
    if (!Array.isArray(rows)) {
      cache.approachesByAirport.set(ident, []);
      return [];
    }

    const mapped = dedupeApproaches(rows
      .map((r) => ({
        name: r.chart_name || r.chartName || "Approach",
        pdf: r.pdf_path || r.pdfPath || null,
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

  const html = await fetchTextFromAny([
    `https://www.aviationapi.com/charts?apt=${encodeURIComponent(ident)}`,
    `https://www.aviationapi.com/charts/?apt=${encodeURIComponent(ident)}`
  ]);

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
      if (pdf && pdf.startsWith("/")) pdf = `https://www.aviationapi.com${pdf}`;
      if (pdf && !pdf.startsWith("http")) pdf = null;

      allCharts.push({ name, pdf, section, source: "AviationAPI HTML" });
    }
  }

  const dedupedAll = dedupeApproaches(allCharts);
  const approaches = dedupeApproaches(dedupedAll.filter((chart) => isApproachChartRow(chart)));
  const bundle = { allCharts: dedupedAll, approaches };
  cache.aviationChartsByAirport.set(ident, bundle);
  return bundle;
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
  const selected = model.routes[selectedRouteIndex] || model.routes[0];
  const needed = (model.requiredApproachTypes || []).length
    ? ` | Needed: ${(model.requiredApproachTypes || []).join(", ")}`
    : "";
  const summaryHtml = `
    <article class="summary">
      <div><strong>${model.origin.ident}</strong> | ${FLIGHT_RULES[model.flightType].label} | ${model.flightDate} ${model.timeLocal}</div>
      <div class="subtle">
        ${model.routes.length} route options ranked by training score (1-100).
        Surface winds are Open-Meteo 10m winds in knots.${needed}
      </div>
      <div class="small">${model.usingFallback ? "Airport/runway dataset fallback enabled." : ""}</div>
    </article>
  `;

  const listHtml = model.routes
    .map((route, index) => {
      const band = scoreBand(route.score);
      const selectedClass = index === selectedRouteIndex ? " route-card-selected" : "";
      const expandedHtml = index === selectedRouteIndex
        ? routeDetailInlineHtml(route, model.origin, model.originProfile, mapContainerId(route))
        : "";
      return `
        <article class="route-card${selectedClass}">
          <div class="route-head">
            <div>
              <h3>Option ${index + 1}: ${route.summary}</h3>
              <div class="subtle">${route.totalDistanceNm.toFixed(1)} NM total | ${route.legs.length} legs</div>
            </div>
            <div class="route-score score-${band}">${route.score}</div>
          </div>

          <div class="route-overview">
            ${route.stops.map((stop) => stopOverviewHtml(stop)).join("")}
          </div>

          <div>
            ${index === selectedRouteIndex
              ? "<span class=\"small\">Selected</span>"
              : `<button class=\"secondary\" data-route-index=\"${index}\">Show More</button>`}
          </div>
          ${expandedHtml}
        </article>
      `;
    })
    .join("\n");
  resultsEl.innerHTML = summaryHtml + `<section class="route-grid">${listHtml}</section>`;
  if (selected) renderRouteMap(selected, model.origin, mapContainerId(selected));
}

function stopOverviewHtml(stop) {
  const runway = stop.likelyRunway;
  const displayTypes = stop.approachStats.typesLikely.filter((t) => INCLUDE_TACAN_FOR_T6_SCORING || t !== "TACAN");
  const types = displayTypes.length ? displayTypes.join(", ") : "None detected";
  const runwayBand = runway ? runway.highlightBand : "red";
  const hasWarnings = stop.suitability.flags.length > 0;

  return `
    <div class="overview-row">
      <div><strong>${stop.ident}</strong> (${stop.name || ""})</div>
      <div class="subtle">Wind ${stop.wind.directionDeg.toFixed(0)}° @ ${stop.wind.speedKt.toFixed(0)} kt</div>
      <div class="subtle"><span class="dot dot-${runwayBand}"></span> Likely RWY ${runway ? runway.runwayIdent : "N/A"} ${runway ? `(${runway.lengthFt}x${runway.widthFt || "?"} ft)` : ""}</div>
      <div class="subtle">Approach types: ${types}</div>
      ${hasWarnings ? `<div class="small warning">Operational caution: ${escapeHtml(stop.suitability.flags[0])}. Check IFR supplement for PPR/remarks.</div>` : ""}
    </div>
  `;
}

function routeDetailInlineHtml(route, origin, originProfile, mapId) {
  const legRows = route.legs
    .map((leg) => `<div class="leg-row"><span>${leg.from} -> ${leg.to}</span><span>${leg.distanceNm.toFixed(1)} NM</span></div>`)
    .join("");

  const focusAirfields = [originProfile, ...route.stops];

  const focusCards = focusAirfields
    .map((airport, index) => {
      const label = index === 0
        ? "Home / Return Airfield"
        : `Training Airfield ${index + 1}`;
      return airfieldSummaryCardHtml(airport, label);
    })
    .join("");

  const detailCards = route.stops
    .map((stop, idx) => airfieldDetailCardHtml(stop, `Destination Detail ${idx + 1}`))
    .join("\n");

  const homeDetailCard = airfieldDetailCardHtml(originProfile, "Home Field Detail");

  return `
    <div class="route-inline-detail">
      <h3>Selected Route Detail</h3>
      <div class="card mb-3">
        <div class="card-body">
          <div class="row g-3">
            <div class="col-lg-5">
              <h5 class="mb-2">Mission Plan</h5>
              <div class="leg-list">${legRows}</div>
            </div>
            <div class="col-lg-7">
              <h5 class="mb-2">Primary Airfields</h5>
              <div class="row g-2">${focusCards}</div>
            </div>
          </div>
        </div>
      </div>
      <div class="map-wrap">
        <strong>Route Quick Map</strong>
        <div id="${mapId}" class="route-map"></div>
      </div>
      ${homeDetailCard}
      ${detailCards}
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
      <div class="airfield-chip">
        <div class="small text-uppercase">${label}</div>
        <div><strong>${airport.ident}</strong> ${airport.name ? `- ${escapeHtml(airport.name)}` : ""}</div>
        <div class="small"><span class="dot dot-${band}"></span> ${airport.wind.directionDeg.toFixed(0)}° @ ${airport.wind.speedKt.toFixed(0)} kt</div>
        <div class="small">Likely RWY ${runway ? runway.runwayIdent : "N/A"} ${runway ? `(${runway.lengthFt}x${runway.widthFt || "?"} ft)` : ""}</div>
        <div class="small">Approach set: ${approachTypes || "None detected"}</div>
      </div>
    </div>
  `;
}

function airfieldDetailCardHtml(stop, title) {
  const likelyApproaches = stop.approachStats.selectedApproaches;
  const alternateApproaches = stop.approachStats.alternateApproaches;
  const runwayBand = stop.likelyRunway ? stop.likelyRunway.highlightBand : "red";
  const approachSources = [...new Set((stop.approaches || []).map((a) => a.source).filter(Boolean))];
  const miniRunway = renderRunwayMiniView(stop.runways || [], stop.likelyRunway?.runwayIdent, stop.wind);
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

  return `
    <article class="detail-card card">
      <div class="card-body">
        <h4>${title}: ${stop.ident}</h4>
        <div class="detail-meta">
          <span class="dot dot-${runwayBand}"></span>
          <span>Training Score ${stop.trainingScore}</span>
          <span>Wind ${stop.wind.directionDeg.toFixed(0)}° @ ${stop.wind.speedKt.toFixed(0)} kt</span>
          <span>Likely RWY ${stop.likelyRunway ? stop.likelyRunway.runwayIdent : "N/A"}</span>
        </div>
        <div class="combined-mini row g-2">
          <div class="col-lg-5">
            <div class="airfield-mini-info">
              <strong>Airfield Mini Info</strong>
              <div class="small">Towered: ${towered ? "Yes" : "No"}</div>
              <div class="small">Lighting: ${hasLighting ? "Available" : "Unknown/Unlit"} | Elev: ${stop.elevationFt ?? "N/A"} ft</div>
              <div class="small">Service: ${stop.scheduledService ? "Scheduled" : "Non-scheduled"}</div>
              <div class="small mt-1"><strong>Available Runways</strong></div>
              <ul class="small runway-list-mini">${runwaysList || "<li>No runway records</li>"}</ul>
            </div>
          </div>
          <div class="col-lg-7">
            <div class="mini-view-wrap">${miniRunway}</div>
          </div>
        </div>
        <div class="approach-columns row g-2">
          <div class="col-md-6">
            <div class="approaches">
              <strong>Likely runway approaches</strong>
              ${filteredLikely.length
                ? `<ul>${filteredLikely.map((a) => `<li>${approachLinkHtml(a)}</li>`).join("")}</ul>`
                : "<div class=\"small\">No approach chart metadata returned.</div>"}
            </div>
          </div>
          <div class="col-md-6">
            <div class="approaches alt-approaches">
              <strong>Other available approaches</strong>
              ${filteredAlternate.length
                ? `<ul>${filteredAlternate.map((a) => `<li>${approachLinkHtml(a)}</li>`).join("")}</ul>`
                : "<div class=\"small\">No additional approaches found.</div>"}
            </div>
          </div>
        </div>
        <div class="small mt-2">
          Score: diversity ${Math.round(stop.scoreBreakdown.diversity)}, quantity ${Math.round(stop.scoreBreakdown.quantity)}, wind ${Math.round(stop.scoreBreakdown.wind)}, runway ${Math.round(stop.scoreBreakdown.runway)}, suitability penalty ${Math.round(stop.scoreBreakdown.suitabilityPenalty)}.
        </div>
        ${stop.approachStats.hasRadarMinimums ? "<div class=\"small\">Radar minimums published: GCA training capability available.</div>" : ""}
        ${stop.suitability.flags.length ? `<div class="small warning">Suitability flags: ${escapeHtml(stop.suitability.flags.join("; "))}. Check IFR supplement for PPR/remarks.</div>` : ""}
        <div class="small">Approach data source: ${approachSources.length ? approachSources.join(", ") : "Unavailable"}</div>
        ${!INCLUDE_TACAN_FOR_T6_SCORING ? "<div class=\"small\">TACAN approaches are available in backend data but excluded from T-6 scoring/output.</div>" : ""}
      </div>
    </article>
  `;
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

function renderRunwayMiniView(runways, likelyRunwayIdent, wind) {
  if (!runways.length) return "<div class=\"small\">No runway geometry available.</div>";

  const centerX = 86;
  const centerY = 96;
  const maxRadius = 72;
  const lines = [];
  const candidates = runways
    .slice(0, 14)
    .map((runway) => {
      const ends = (runway.ends || []).filter((e) => Number.isFinite(e.heading));
      if (!ends.length) return null;
      const a = ends[0];
      const b = ends[1] || null;
      const heading = normalizeHeading(a.heading) % 180;
      return { runway, a, b, headingBucket: Math.round(heading / 2) * 2 };
    })
    .filter(Boolean);

  const grouped = new Map();
  for (const item of candidates) {
    if (!grouped.has(item.headingBucket)) grouped.set(item.headingBucket, []);
    grouped.get(item.headingBucket).push(item);
  }

  for (const groupItems of grouped.values()) {
    for (let i = 0; i < groupItems.length; i += 1) {
      const item = groupItems[i];
      const angle = degToRad((item.a.heading - 90 + 360) % 360);
      const sideOffset = runwaySideOffset(item.a.ident);
      const siblingOffset = (i - (groupItems.length - 1) / 2) * 4;
      const normalX = -Math.sin(angle);
      const normalY = Math.cos(angle);
      const lateral = sideOffset + siblingOffset;

      const scale = clamp((item.runway.lengthFt || 4000) / 12000, 0.25, 1);
      const half = maxRadius * scale;
      const cx = centerX + normalX * lateral;
      const cy = centerY + normalY * lateral;

      const x1 = cx - Math.cos(angle) * half;
      const y1 = cy - Math.sin(angle) * half;
      const x2 = cx + Math.cos(angle) * half;
      const y2 = cy + Math.sin(angle) * half;

      const likely = [item.a.ident, item.b?.ident]
        .map((v) => normalizeRunwayIdent(v))
        .includes(normalizeRunwayIdent(likelyRunwayIdent));

      lines.push(`<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" class="mini-line${likely ? " likely" : ""}" />`);
      lines.push(`<text x="${(x2 + 2).toFixed(1)}" y="${(y2 - 2).toFixed(1)}" class="mini-label">${escapeHtml(item.a.ident || "")}</text>`);
      if (item.b?.ident) {
        lines.push(`<text x="${(x1 + 2).toFixed(1)}" y="${(y1 - 2).toFixed(1)}" class="mini-label">${escapeHtml(item.b.ident)}</text>`);
      }
    }
  }

  const windArrow = buildWindArrowSvg(wind, centerX, centerY);

  return `
    <svg class="mini-runway-svg" viewBox="-20 -10 280 230" role="img" aria-label="Mini runway layout">
      <defs>
        <marker id="wind-arrowhead" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse">
          <path d="M 0 0 L 10 5 L 0 10 z" fill="#174566"></path>
        </marker>
      </defs>
      <circle cx="${centerX}" cy="${centerY}" r="80" class="mini-bg"></circle>
      ${lines.join("")}
      ${windArrow}
    </svg>
  `;
}

function runwaySideOffset(ident) {
  const side = String(ident || "").toUpperCase().slice(-1);
  if (side === "L") return -6;
  if (side === "R") return 6;
  if (side === "C") return 0;
  return 0;
}

function buildWindArrowSvg(wind, centerX, centerY) {
  if (!wind) return "";
  const bearing = Number(wind.directionDeg);
  if (!Number.isFinite(bearing)) return "";

  // Keep wind indicator in a dedicated right-side lane to avoid overlapping runway labels.
  const laneX = centerX + 122;
  const laneY = centerY - 22;
  const rad = degToRad((bearing - 90 + 360) % 360);
  const length = 28;
  const x1 = laneX - Math.cos(rad) * length * 0.5;
  const y1 = laneY - Math.sin(rad) * length * 0.5;
  const x2 = laneX + Math.cos(rad) * length * 0.5;
  const y2 = laneY + Math.sin(rad) * length * 0.5;

  return `
    <text x="${(laneX - 26).toFixed(1)}" y="${(laneY - 20).toFixed(1)}" class="wind-text">Wind</text>
    <line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" class="wind-arrow"></line>
    <text x="${(laneX - 42).toFixed(1)}" y="${(laneY + 22).toFixed(1)}" class="wind-text">${Math.round(wind.directionDeg)}@${Math.round(wind.speedKt)}kt</text>
  `;
}

function mapContainerId(route) {
  return `route-map-${String(route.id).replace(/[^a-zA-Z0-9_-]/g, "-")}`;
}

function renderRouteMap(route, origin, mapId) {
  const mapEl = document.getElementById(mapId);
  if (!mapEl || typeof window.L === "undefined") return;

  if (routeMapInstance) {
    routeMapInstance.remove();
    routeMapInstance = null;
  }

  const points = [origin, ...route.stops, origin]
    .filter((p) => Number.isFinite(p.lat) && Number.isFinite(p.lon))
    .map((p) => [p.lat, p.lon]);
  if (!points.length) return;

  routeMapInstance = window.L.map(mapEl, { zoomControl: true, scrollWheelZoom: false });
  window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 13,
    attribution: "&copy; OpenStreetMap contributors"
  }).addTo(routeMapInstance);

  const polyline = window.L.polyline(points, { color: "#0a5d8f", weight: 4 }).addTo(routeMapInstance);
  window.L.marker([origin.lat, origin.lon]).addTo(routeMapInstance).bindPopup(`Origin: ${origin.ident}`);

  for (const stop of route.stops) {
    window.L.marker([stop.lat, stop.lon]).addTo(routeMapInstance).bindPopup(`${stop.ident} - ${stop.name || "Stop"}`);
  }

  routeMapInstance.fitBounds(polyline.getBounds(), { padding: [18, 18] });
  setTimeout(() => routeMapInstance?.invalidateSize(), 0);
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

async function fetchTextFromAny(urls) {
  let lastError = null;
  const expanded = urls.flatMap((url) => buildFetchCandidates(url));

  for (const url of expanded) {
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return await response.text();
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("All fetches failed");
}

async function fetchJsonFromAny(urls) {
  const text = await fetchTextFromAny(urls);
  try {
    return JSON.parse(text);
  } catch (_error) {
    throw new Error("Invalid JSON response");
  }
}

function buildFetchCandidates(url) {
  const out = [];
  if (!String(url).startsWith("http")) return [url];

  const isLocal = LOCAL_HOSTS.has(window.location.hostname);
  const encoded = encodeURIComponent(url);

  // Local static hosting (python http.server) has no /api/proxy route.
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
  out.push(`/api/proxy?url=${encoded}`);
  out.push(url);
  return out;
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

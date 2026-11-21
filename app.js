"use strict";

const COMPONENTS = ["Container", "Exterior", "Interior"];
const LEVELS = ["Started", "Advanced", "Excellent", "Elite"];
const GAME_NAMES = ["Aerial", "Distance", "Speed", "Team"];
let titles = [], scores = [], games = [];
let titlesFileLoaded = false;
let currentPdfUrl = "";

const $ = (id) => document.getElementById(id);
const pendingCheckStyle = document.createElement("style");
pendingCheckStyle.textContent = ".requirements li.pending,.requirement-list li.pending,.next-card li.pending{color:#b42318!important;font-weight:600}.achievement-check.pending{color:#b42318!important}";
document.head.appendChild(pendingCheckStyle);
const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character]));
function localDateStamp(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
function dogDisplayName() {
  const raw = String(scores[0]?.dog || "").trim();
  const number = String(scores[0]?.dogNumber || "").trim();
  return raw && !/^(?:n\/?a|na|unknown|null|none|-+)$/i.test(raw) ? raw : number ? `SDDA dog #${number}` : "SDDA dog";
}
function reportFileStem() {
  const dog = dogDisplayName().replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ");
  return `${dog} - SDDA Achievement Ledger - ${localDateStamp()}`;
}

function parseCsv(text) {
  const rows = []; let row = [], field = "", quoted = false;
  text = text.replace(/^\uFEFF/, "");
  for (let index = 0; index < text.length; index++) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') { field += '"'; index++; }
      else if (character === '"') quoted = false;
      else field += character;
    } else if (character === '"') quoted = true;
    else if (character === ",") { row.push(field); field = ""; }
    else if (character === "\n") { row.push(field.replace(/\r$/, "")); rows.push(row); row = []; field = ""; }
    else field += character;
  }
  if (field || row.length) { row.push(field.replace(/\r$/, "")); rows.push(row); }
  const headers = (rows.shift() || []).map((header) => header.trim());
  return { headers, data: rows.filter((values) => values.some((value) => value.trim())).map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]))) };
}

function tableFromRows(rows, headerTest) {
  const headerIndex = rows.findIndex((values) => {
    const headers = values.map((value) => String(value ?? "").trim());
    return headerTest(headers);
  });
  if (headerIndex < 0) return null;
  const headers = rows[headerIndex].map((header) => String(header ?? "").trim());
  const data = rows.slice(headerIndex + 1)
    .filter((values) => values.some((value) => String(value ?? "").trim()))
    .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])));
  return { headers, data };
}

async function parseAccountFile(file, headerTest, label) {
  const extension = file.name.split(".").pop().toLowerCase();
  const csvLike = extension === "csv" || /(?:text\/csv|application\/csv|text\/comma-separated-values)/i.test(file.type || "");
  if (csvLike) {
    const parsed = parseCsv(await file.text());
    if (!headerTest(parsed.headers)) throw new Error(`That is not the ${label} file.`);
    return { ...parsed, sheetName: "CSV" };
  }
  if (extension !== "xlsx") throw new Error("Please choose a CSV or XLSX file.");

  const workbook = window.XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  for (const sheetName of workbook.SheetNames) {
    const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    const parsed = tableFromRows(rows, headerTest);
    if (parsed) return { ...parsed, sheetName };
  }
  throw new Error(`No worksheet in that workbook contains the expected ${label} columns.`);
}

function levelFrom(description) {
  const text = String(description || "");
  const code = text.match(/\((S|A|E|EL|L|X)\)\s*$/i)?.[1]?.toUpperCase();
  if (code === "S") return "Started"; if (code === "A") return "Advanced"; if (code === "E") return "Excellent"; if (code === "EL" || code === "L" || code === "X") return "Elite";
  return LEVELS.find((level) => new RegExp(level, "i").test(text)) || null;
}
function componentFrom(description) { return COMPONENTS.find((component) => new RegExp(component, "i").test(description)) || null; }
function gameFrom(description) { const text = String(description || ""); const match = text.match(/\bGames?\s*[-:]\s*(Aerial|Distance|Speed|Teams?)\b/i) || text.match(/\b(Aerial|Distance|Speed|Teams?)\s+Games?\b/i); if (!match) return null; const name = match[1].toLowerCase(); return name.startsWith("team") ? "Team" : name[0].toUpperCase() + name.slice(1); }

function titleText(title) { return `${title.description} ${title.honours}`.trim(); }
function gameTitleRecord(game) { const codes = { Aerial: "GA", Distance: "GD", Speed: "GS", Team: "GT" }; return titles.find((title) => new RegExp(`(?:^|\\b)${game}(?:\s+Games)?(?:\\b|$)|SD-${codes[game]}(?:\\b|$)`, "i").test(titleText(title))); }
function gamesTitleRecorded(game) { return !!gameTitleRecord(game); }
function findTitle(pattern) { return titles.find((title) => pattern.test(titleText(title))); }
function dateLabel(value) {
  if (!value) return "Date not supplied";
  const date = new Date(`${value}T00:00:00`); return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}
function byLevel(level, list = scores) { return list.filter((score) => score.level === level); }
function componentCounts(level, list = scores) { return Object.fromEntries(COMPONENTS.map((component) => [component, list.filter((score) => score.level === level && score.component === component).length])); }
function gamePassCounts(){return Object.fromEntries(GAME_NAMES.map(g=>[g,Math.max(games.filter(r=>r.game===g&&r.pass!==false).length,gamesTitleRecorded(g)?3:0)]));}
const GAME_CHAMPIONSHIP_RULES=[{key:"gch",name:"SD-GS",type:"base",target:0,text:"3 passes in each of any 3 Games"},{key:"gch2",name:"SD-GCH2",type:"three",target:4,text:"4 additional passes across any 3 Games (13 cumulative)"},{key:"gch3",name:"SD-GCH3",type:"three",target:9,text:"5 additional passes after GCH2 (18 cumulative)"},{key:"gch4",name:"SD-GCH4",type:"three",target:16,text:"7 additional passes after GCH3 (25 cumulative)"},{key:"gch5",name:"SD-GCH5",type:"four",target:26,text:"10 additional passes after GCH4 across all 4 Games (35 cumulative)"}];
function gameChampionshipProgress(counts, rule) {
  let best = null;
  for (let a = 0; a < 4; a++) for (let b = a + 1; b < 4; b++) for (let c = b + 1; c < 4; c++) {
    const trio = [GAME_NAMES[a], GAME_NAMES[b], GAME_NAMES[c]];
    const fourth = GAME_NAMES.find((game) => !trio.includes(game));
    const qualifiedGames = trio.filter((game) => counts[game] >= 3).length;
    const baseMissing = trio.reduce((sum, game) => sum + Math.max(0, 3 - counts[game]), 0);
    const extras3 = trio.reduce((sum, game) => sum + Math.max(0, counts[game] - 3), 0);
    const extras = rule.type === "four" ? extras3 + counts[fourth] : extras3;
    const missingAdditionalGames = rule.type === "four" ? GAME_NAMES.filter((game) => counts[game] <= 0).length : 0;
    const topThreeTotal = trio.reduce((sum, game) => sum + counts[game], 0);
    const allFourTotal = GAME_NAMES.reduce((sum, game) => sum + counts[game], 0);
    const baseMet = qualifiedGames === 3;
    const remaining = rule.type === "base"
      ? baseMissing
      : baseMissing + Math.max(Math.max(0, rule.target - extras), missingAdditionalGames);
    const met = rule.type === "base"
      ? baseMet
      : baseMet && extras >= rule.target && missingAdditionalGames === 0;
    const candidate = { met, remaining, extras, baseMissing, baseMet, qualifiedGames, trio, fourth, topThreeTotal, allFourTotal };
    if (!best || candidate.remaining < best.remaining ||
      (candidate.remaining === best.remaining && candidate.qualifiedGames > best.qualifiedGames) ||
      (candidate.remaining === best.remaining && candidate.qualifiedGames === best.qualifiedGames && candidate.extras > best.extras)) {
      best = candidate;
    }
  }
  const progressText = rule.type === "base"
    ? `${best.qualifiedGames}/3 Games have at least 3 passes (${best.trio.join(", ")}) · ${best.topThreeTotal} total passes across those Games`
    : `${best.qualifiedGames}/3 Games meet the initial 3-pass minimum · ${best.extras}/${rule.target} additional passes in ${best.trio.join(", ")}${rule.type === "four" ? " + " + best.fourth : ""}`;
  return { ...best, progressText };
}


function validComponentScore(level, component, value) {
  const text = String(value ?? "").trim();
  if (!text) return false;
  const score = Number(text), started = level === "Started";
  const minimum = component === "Interior" ? (started ? 20 : 40) : (started ? 15 : 30);
  const maximum = component === "Interior" ? (started ? 40 : 80) : (started ? 30 : 60);
  return Number.isFinite(score) && score >= minimum && score <= maximum;
}

function renderGames() {
 const counts=gamePassCounts(),total=Object.values(counts).reduce((a,b)=>a+b,0),titled=GAME_NAMES.filter(g=>counts[g]>=3),progress=GAME_CHAMPIONSHIP_RULES.map(rule=>({rule,...gameChampionshipProgress(counts,rule)}));
 $("games-summary").innerHTML="<span>"+total+" banked Games passes</span><span>"+titled.length+"/4 individual Games titles</span><span>"+(progress[0].met?"SD-GS requirements met":progress[0].remaining+" more pass(es) needed for SD-GS")+"</span>";
 $("games-grid").innerHTML=GAME_NAMES.map(game=>{const passes=counts[game],entries=games.filter(r=>r.game===game&&r.pass!==false),recorded=gameTitleRecord(game),met=passes>=3,details=game==="Aerial"?"1 hide · 1 distractor · recommended 2 min; judge may adjust":game==="Distance"?"1 hide · no distractors · recommended 3 min; judge may adjust":game==="Speed"?"2–4 hides · 2–4 distractors · fixed 60 sec":"2–5 hides · 2–4 distractors · recommended 5 min; judge may adjust",rows=entries.map(r=>"<div class='game-attempt'><span>"+esc(dateLabel(r.date))+"</span><span>Trial "+esc(r.trial||"—")+"</span><strong>Q / PASS</strong></div>").join("");return "<article class='game-card'><h3>"+esc(game)+" <small>SD-G"+game[0]+"</small></h3><span class='status "+(met?"met":passes?"partial":"unknown")+"'>"+(met?"Individual title requirement met":passes?"In progress":"No pass data")+"</span><div class='game-count'><strong>"+passes+"</strong><span>/ 3 passes</span></div><p>"+details+"</p><div class='game-history'><h4>Games pass history</h4><div class='game-history-summary'><span>"+passes+" passes</span><span>"+entries.length+" listed entries</span></div><div class='game-attempt-header'><span>Date</span><span>Trial</span><span>Result</span></div>"+(rows||"<p>No Games pass entries recorded.</p>")+"</div>"+(recorded?"<p><strong>SDDA title record:</strong> "+esc(dateLabel(recorded.date))+"</p>":"")+"</article>";}).join("");
 $("games-championships").innerHTML="<h3>GCH titles</h3><div class='game-ladder'>"+progress.map(({rule,met,remaining})=>"<p><strong>"+rule.name+"</strong><span class='status "+(met?"met":"partial")+"'>"+(met?"Recorded / requirements met":remaining+" needed")+"</span></p>").join("")+"</div>";
}

function workbookDate(value) {
  const text = String(value ?? "").trim();
  const numeric = typeof value === "number" ? value : /^\d{4,5}(?:\.\d+)?$/.test(text) ? Number(text) : NaN;
  if (Number.isFinite(numeric)) {
    const parsed = window.XLSX.SSF.parse_date_code(numeric);
    if (parsed) return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.toISOString().slice(0, 10);
  if (!text) return "";
  const parsed = new Date(text); return Number.isNaN(parsed.valueOf()) ? "" : parsed.toISOString().slice(0, 10);
}
function chronologicalTrialGroups(level, list = scores) {
  const byTrial = new Map();
  const ordered = list.filter((score) => score.level === level).sort((a, b) => a.sourceRow - b.sourceRow);
  for (const row of ordered) {
    const trialNumber = Number(row.trial);
    const key = row.trial && Number.isSafeInteger(trialNumber) && trialNumber > 0
      ? `trial:${trialNumber}`
      : `row:${row.sourceRow}`;
    if (!byTrial.has(key)) byTrial.set(key, { key, trial: row.trial, rows: [] });
    byTrial.get(key).rows.push(row);
  }
  return [...byTrial.values()];
}

function firstCompletedTitleSet(level, list = scores) {
  if (level !== "Advanced" && level !== "Excellent") return null;
  const groups = chronologicalTrialGroups(level, list), seen = [], passed = new Set();
  for (const group of groups) {
    const groupRows = [...group.rows].sort((a, b) => a.sourceRow - b.sourceRow);
    const newlyPassedRows = groupRows.filter((row) => !passed.has(row.component));
    const componentsHere = new Set(groupRows.map((row) => row.component));
    seen.push(...groupRows);
    groupRows.forEach((row) => passed.add(row.component));
    if (!COMPONENTS.every((component) => passed.has(component))) continue;
    const completionRow = newlyPassedRows[newlyPassedRows.length - 1];
    if (!completionRow) return null;
    const special = COMPONENTS.every((component) => componentsHere.has(component));
    const parts = Object.fromEntries(COMPONENTS.map((component) => {
      const eligible = special || component === completionRow.component ? groupRows : seen;
      return [component, eligible.filter((row) => row.component === component)
        .sort((a, b) => b.score - a.score || a.sourceRow - b.sourceRow)[0]];
    }));
    if (!COMPONENTS.every((component) => parts[component])) return null;
    const total = COMPONENTS.reduce((sum, component) => sum + parts[component].score, 0);
    return { level, trial: group.trial, parts, total, gold: total >= 192, special,
      completionComponent: completionRow.component, completionTrial: group.trial, eligibleRows: [...seen] };
  }
  return null;
}

function initialSpecialSet(level, list = scores) {
  const completed = firstCompletedTitleSet(level, list);
  return completed?.special ? completed : null;
}

function highToLow(level, list = scores) {
  const special = list === scores ? initialSpecialSet(level, list) : null;
  const lockedRows = new Set(special ? COMPONENTS.map((component) => special.parts[component].sourceRow) : []);
  const ranked = Object.fromEntries(COMPONENTS.map((component) => [component, list.filter((score) => score.level === level && score.component === component && !lockedRows.has(score.sourceRow)).sort((a, b) => b.score - a.score)]));
  const count = Math.min(...COMPONENTS.map((component) => ranked[component].length));
  const combinations = Array.from({ length: Number.isFinite(count) ? count : 0 }, (_, index) => {
    const container = ranked.Container[index].score, interior = ranked.Interior[index].score, exterior = ranked.Exterior[index].score;
    return { level, rank: index + 1, container, interior, exterior, total: container + interior + exterior, gold: container + interior + exterior >= 192 };
  });
  if (special) combinations.unshift({ level: `${level} SP`, rank: "SP", container: special.parts.Container.score, interior: special.parts.Interior.score, exterior: special.parts.Exterior.score, total: special.total, gold: special.gold, special: true, trial: special.trial });
  return combinations;
}


function titleLockForLevel(level, list = scores) {
  const rows = list.filter((row) => row.level === level).sort((a, b) => a.sourceRow - b.sourceRow);
  if (!rows.length) return null;
  const matchingTitles = (titles || [])
    .filter((title) => {
      const description = String(title.description || "").trim().toLowerCase();
      return description === level.toLowerCase() || description.startsWith(level.toLowerCase() + " ");
    })
    .filter((title) => !String(title.description || "").toLowerCase().includes("special"))
    .sort((a, b) => String(a.date || "9999-99-99").localeCompare(String(b.date || "9999-99-99")));
  const recordedTitle = matchingTitles[0] || null;
  const eligibleHistory = recordedTitle?.date
    ? rows.filter((row) => row.date && row.date <= recordedTitle.date)
    : rows;
  const calculated = firstCompletedTitleSet(level, eligibleHistory);
  if (!calculated) return null;
  const parts = calculated.parts;
  const total = COMPONENTS.reduce((sum, component) => sum + parts[component].score, 0);
  return { ...calculated, parts, total, gold: total >= 192, special: !!calculated.special,
    recorded: !!recordedTitle, date: recordedTitle?.date || calculated.completionTrial || "" };
}

function maximizeGoldCombinations(level, list = scores, maximum = 10, includeTitleLock = true) {
  // Exact three-component matching. At most ten combinations are needed for
  // any title stage, so only the best N scores per component can matter:
  // replacing a selected score with a higher unused score cannot reduce the
  // number of qualifying (>=100) combinations.
  const titleLock = includeTitleLock ? titleLockForLevel(level, list) : null;
  const lockedRows = new Set(titleLock ? COMPONENTS.map((component) => titleLock.parts[component].sourceRow) : []);
  const locked = titleLock ? [{
    rank: titleLock.special ? "SP" : 1,
    container: titleLock.parts.Container, interior: titleLock.parts.Interior, exterior: titleLock.parts.Exterior,
    total: titleLock.total, gold: titleLock.gold, special: !!titleLock.special,
    titleLocked: !titleLock.special, titleRecorded: !!titleLock.recorded, titleDate: titleLock.date || "",
    level, trial: titleLock.trial || titleLock.parts.Container.trial || ""
  }] : [];
  const remainingLimit = Math.max(0, Math.min(10, Number(maximum) - locked.length));
  const eligible = list.filter((score) => score.level === level && !lockedRows.has(score.sourceRow));
  const ranked = Object.fromEntries(COMPONENTS.map((component) => [
    component,
    eligible.filter((score) => score.component === component)
      .sort((a, b) => b.score - a.score || a.sourceRow - b.sourceRow)
      .slice(0, remainingLimit)
  ]));
  const containers = ranked.Container, exteriors = ranked.Exterior, interiors = ranked.Interior;
  const memo = new Map();
  let searchNodes = 0;
  const bitCount = (value) => {
    let count = 0;
    while (value) { value &= value - 1; count++; }
    return count;
  };
  const better = (candidate, current) => {
    if (!current) return true;
    if (candidate.count !== current.count) return candidate.count > current.count;
    if (candidate.gold !== current.gold) return candidate.gold > current.gold;
    if (candidate.total !== current.total) return candidate.total > current.total;
    return candidate.tie < current.tie;
  };
  const solve = (index, exteriorMask, interiorMask) => {
    const key = index + "|" + exteriorMask + "|" + interiorMask;
    if (memo.has(key)) return memo.get(key);
    searchNodes++;
    if (index >= containers.length || remainingLimit === 0) {
      const empty = { count: 0, gold: 0, total: 0, tie: "", sets: [] };
      memo.set(key, empty);
      return empty;
    }
    const remainingContainers = containers.length - index;
    const maximumAdditional = Math.min(remainingContainers,
      exteriors.length - bitCount(exteriorMask),
      interiors.length - bitCount(interiorMask),
      remainingLimit);
    let best = null;
    for (let e = 0; e < exteriors.length; e++) {
      if (exteriorMask & (1 << e)) continue;
      for (let i = 0; i < interiors.length; i++) {
        if (interiorMask & (1 << i)) continue;
        const container = containers[index], exterior = exteriors[e], interior = interiors[i];
        const total = container.score + exterior.score + interior.score;
        if (total < 100) continue;
        const tail = solve(index + 1, exteriorMask | (1 << e), interiorMask | (1 << i));
        const set = { level, container, exterior, interior, total, gold: total >= 192 };
        const candidate = {
          count: tail.count + 1,
          gold: tail.gold + Number(set.gold),
          total: tail.total + total,
          tie: [container.sourceRow, exterior.sourceRow, interior.sourceRow].join(":") + "|" + tail.tie,
          sets: [set, ...tail.sets]
        };
        if (better(candidate, best)) best = candidate;
      }
    }
    // Since Container scores are descending, skipping this record cannot
    // improve Gold count or total when the best possible number is already met.
    if (!best || best.count < maximumAdditional) {
      const skipped = solve(index + 1, exteriorMask, interiorMask);
      if (better(skipped, best)) best = skipped;
    }
    if (!best) best = { count: 0, gold: 0, total: 0, tie: "", sets: [] };
    memo.set(key, best);
    return best;
  };
  const optimized = solve(0, 0, 0);
  const combinations = optimized.sets
    .sort((a, b) => b.total - a.total ||
      a.container.sourceRow - b.container.sourceRow ||
      a.exterior.sourceRow - b.exterior.sourceRow ||
      a.interior.sourceRow - b.interior.sourceRow)
    .map((set, index) => ({ ...set, rank: index + 1 }));
  const result = [...locked, ...combinations];
  result.optimizationComplete = true;
  result.searchNodes = searchNodes;
  result.pairChecks = searchNodes;
  result.optimization = "exact-three-component-matching";
  return result;
}
function eliteSets(list = scores) {
  const trials = new Map();
  list.filter((score) => score.level === "Elite").forEach((score) => {
    if (!score.trial) return;
    const key = score.trial;
    const group = trials.get(key) || { trial: score.trial, date: score.date, parts: {} };
    if (!group.parts[score.component] || group.parts[score.component].score < score.score) group.parts[score.component] = score;
    trials.set(key, group);
  });
  return [...trials.values()].filter((group) => COMPONENTS.every((component) => group.parts[component])).map((group) => {
    const total = COMPONENTS.reduce((sum, component) => sum + group.parts[component].score, 0);
    return { ...group, total, gold: total >= 192 };
  });
}

function inferredTitles() {
  const inferred = [];
  for (const level of ["Started", "Advanced", "Excellent"]) {
    const rows = byLevel(level);
    if (!COMPONENTS.every((component) => rows.some((row) => row.component === component))) continue;
    // Use the score export's top-to-bottom chronological order. Trial Number
    // values are identifiers, not sortable timestamps; group repeated rows
    // even when the same trial appears in non-adjacent rows.
    const groups = chronologicalTrialGroups(level, rows);
    const seen = [], passed = new Set();
    let completionGroup = null, special = false;
    for (const group of groups) {
      const componentsHere = new Set(group.rows.map((row) => row.component));
      const titledBeforeThisGroup = COMPONENTS.every((component) => passed.has(component));
      seen.push(...group.rows);
      group.rows.forEach((row) => passed.add(row.component));
      if (!titledBeforeThisGroup && COMPONENTS.every((component) => passed.has(component))) {
        completionGroup = group;
        special = COMPONENTS.every((component) => componentsHere.has(component));
        break;
      }
    }
    const scoringRows = special ? completionGroup.rows : seen;
    const best = Object.fromEntries(COMPONENTS.map((component) => [component, Math.max(...scoringRows.filter((row) => row.component === component).map((row) => row.score))]));
    const completed = completionGroup?.rows[completionGroup.rows.length - 1];
    inferred.push({
      description: `${level}${special ? " Special" : ""}`,
      honours: special && completed?.trial ? `Same-trial completion · Trial ${completed.trial}` : "",
      score: COMPONENTS.reduce((sum, component) => sum + best[component], 0),
      date: completed?.date || "",
      calculated: true
    });
  }
  const elite = eliteSets().sort((a, b) => String(a.date || "9999").localeCompare(String(b.date || "9999")))[0];
  if (elite) inferred.push({ description: "Elite", honours: elite.trial ? `Same-trial set · Trial ${elite.trial}` : "Same-trial set", score: elite.total, date: elite.date || "", calculated: true });
  return inferred;
}

function displayedTitles() { return titles.length ? titles.map((title) => ({ ...title, calculated: false })) : inferredTitles(); }


function reportAchievedTitles() {
  // Keep every SDDA-recorded award, then add score-derived achievements that
  // are not already represented in the Dog Titles export.
  const result = titles.map((title) => ({ ...title, calculated: false }));
  const add = (description, date = "", honours = "", score = null, note = "") => {
    if (result.some((title) => title.description.toLowerCase() === description.toLowerCase())) return;
    result.push({ description, date, honours: [honours, note].filter(Boolean).join(" · "), score, calculated: true });
  };
  for (const title of inferredTitles()) {
    if (!result.some((item) => new RegExp(title.description.replace(/ Special$/i, "") + "(?:\\s|$)", "i").test(titleText(item)))) {
      result.push(title);
    }
  }

  const advanced = maximizeGoldCombinations("Advanced");
  const excellent = maximizeGoldCombinations("Excellent");
  const chSelection = championshipTitle("ch")
    ? (historicalChProgressSelection() || { sets: [], advanced: [], excellent: [], total: 0, met: false })
    : selectProgressionCombinations(advanced, excellent, "ch");
  const historicalMachSelection = championshipTitle("mach") ? historicalMachProgressSelection() : null;
  const machSelection = historicalMachSelection || selectProgressionCombinations(advanced, excellent, "mach");
  const elite = eliteSets();
  const eliteGold = elite.filter((set) => set.gold).length;
  const chMet = chSelection.met;
  const machMet = machSelection.met;
  const elchMet = elite.length >= 5 && eliteGold >= 2;
  if (chMet && !championshipTitle("ch")) add("SD-CH", "", "5 titling combinations · " + chSelection.excellent.length + " Excellent");
  if (machMet && !championshipTitle("mach")) add("SD-MACH", "", "10 selected combinations · " + machSelection.excellent.length + " Excellent + " + machSelection.advanced.length + " Advanced · " + machSelection.sets.filter((set) => set.gold).length + " Gold");

  const counts = gamePassCounts();
  for (const game of GAME_NAMES) {
    if (counts[game] >= 3 && !gameTitleRecord(game)) add(game + " (SD-G" + game[0] + ")", "", counts[game] + " qualifying passes");
  }
  const gameStages = GAME_CHAMPIONSHIP_RULES.map((rule) => {
    const progress = gameChampionshipProgress(counts, rule);
    const start = ({ gch2: 0, gch3: 4, gch4: 9, gch5: 16 })[rule.key] ?? 0;
    const size = ({ gch2: 4, gch3: 5, gch4: 7, gch5: 10 })[rule.key] ?? 0;
    const stageProgress = Math.min(size, Math.max(0, progress.extras - start));
    const met = rule.type === "base" ? progress.met : stageProgress >= size && progress.baseMet && (rule.type !== "four" || GAME_NAMES.every((game) => counts[game] > 0));
    return { rule, progress, met };
  });
  for (const { rule, progress, met } of gameStages) {
    if (met && !championshipTitle(rule.key)) add(rule.name, "", progress.progressText);
  }

  const machRules = [
    { key: "mach2", previous: "mach", name: "SD-MACH2", excellent: 7, gold: 3, maxAdvanced: 3 },
    { key: "mach3", previous: "mach2", name: "SD-MACH3", excellent: 10, gold: 3 },
    { key: "mach4", previous: "mach3", name: "SD-MACH4", excellent: 10, gold: 4 }
  ];
  let previousMachMet = !!championshipTitle("mach") || machMet;
  for (const rule of machRules) {
    const recorded = championshipTitle(rule.key);
    const progress = machExtensionProgress(rule, rule.previous);
    const countedAdvanced = rule.maxAdvanced ? Math.min(progress.advanced, rule.maxAdvanced, Math.max(0, 10 - progress.excellent)) : 0;
    const countedTotal = Math.min(10, Math.min(10, progress.excellent) + countedAdvanced);
    const met = progress.verified && previousMachMet && countedTotal >= 10 && progress.excellent >= rule.excellent && progress.gold >= rule.gold;
    if (met && !championshipTitle(rule.key)) add(rule.name, "", "10 additional combinations · " + Math.min(10, progress.excellent) + " Excellent + " + countedAdvanced + " Advanced");
    previousMachMet = !!recorded || met;
  }

  const eliteRules = [
    { key: "elch2", previous: "elch", name: "SD-ELCH2", gold: 3 },
    { key: "elch3", previous: "elch2", name: "SD-ELCH3", gold: 4 },
    { key: "elch4", previous: "elch3", name: "SD-ELCH4", gold: 5 }
  ];
  for (const rule of eliteRules) {
    const previous = championshipTitle(rule.previous);
    const eligible = scoreRowsAfterTitle(previous);
    const sets = eligible ? eliteSets(eligible) : [];
    const goldCount = sets.filter((set) => set.gold).length;
    if (!championshipTitle(rule.key) && eligible && sets.length >= 5 && goldCount >= rule.gold) add(rule.name, "", "5 additional Elite scores · " + goldCount + " Gold");
  }

  // Normalize Special variants and Games title suffixes before ordering.
  // The source export may label these as "Started Special" or "Distance (SD-GD)".
  const normalizeReportTitle = (description) => String(description || "")
    .replace(/\s+Special$/i, "")
    .replace(/\s*\(SD-G[ADST]\)\s*$/i, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
  const order = ["Started","Advanced","Excellent","Elite","SD-CH","SD-MACH","SD-MACH2","SD-MACH3","SD-MACH4","SD-ELCH","SD-ELCH2","SD-ELCH3","SD-ELCH4","SD-GMCH","Aerial","Distance","Speed","Team","SD-GS","SD-GCH2","SD-GCH3","SD-GCH4","SD-GCH5"];
  const groupRank = (title) => {
    const value = normalizeReportTitle(title.description);
    if (/^(aerial|distance|speed|team)$/.test(value) || /^sd-gch(?:2|3|4|5)?$/.test(value)) return 2;
    if (/^sd-(?:ch|mach(?:2|3|4)?|elch(?:2|3|4)?|gmch)$/.test(value)) return 1;
    return 0;
  };
  const rank = (title) => {
    const value = normalizeReportTitle(title.description);
    const index = order.findIndex((name) => normalizeReportTitle(name) === value);
    return index < 0 ? 999 : index;
  };
  return result.sort((a,b) => groupRank(a)-groupRank(b) || rank(a)-rank(b) || String(a.description).localeCompare(String(b.description)));
}

function progressCard(title, status, requirements, note = "") {
  const met = status === "Achieved" || status === "Requirements shown as met";
  const statusClass = met ? "met" : status.includes("cannot") || status.includes("verification") ? "unknown" : "partial";
  return `<article class="progress-card${met ? " criteria-met" : ""}"><h3 class="card-title-bar">${esc(title)}</h3><span class="status ${statusClass}">${esc(status)}</span><ul class="requirements">${requirements.map((requirement) => `<li class="${requirement.pending ? "pending" : requirement.done ? "done" : ""}${requirement.info ? " info" : ""}">${requirement.pending ? "Pending — " : requirement.done ? "Met — " : requirement.info ? "Rule — " : ""}${esc(requirement.text || "Requirement details unavailable")}</li>`).join("")}</ul>${requirements.some((requirement) => requirement.pending) ? `<p class="verification-legend" style="color:#9b433d;font-size:.88em">Handler-reported judge minimum is awaiting SDDA confirmation.</p>` : ""}${note ? `<p>${esc(note)}</p>` : ""}</article>`;
}

function titlePattern(level) { return new RegExp(`^${level}(?:\s|$)`, "i"); }
function championshipTitle(kind){
 const p={
 ch:/(?:^|\b)(?:SD[- ]?)?CH(?:\b|$)|scent detection champion|^Champion$/i,
 mach:/(?:^|\b)(?:SD[- ]?)?MACH(?!\s*[2-4]\b)(?:\b|$)|master champion(?!\s*[2-4]\b)/i,
 mach2:/(?:^|\b)(?:SD[- ]?)?MACH\s*2(?:\b|$)|master champion\s*2\b/i,mach3:/(?:^|\b)(?:SD[- ]?)?MACH\s*3(?:\b|$)|master champion\s*3\b/i,mach4:/(?:^|\b)(?:SD[- ]?)?MACH\s*4(?:\b|$)|master champion\s*4\b/i,
 elch:/(?:^|\b)(?:SD[- ]?)?ELCH(?!\s*[2-4]\b)(?:\b|$)|elite champion(?!\s*[2-4]\b)/i,
 elch2:/(?:^|\b)(?:SD[- ]?)?ELCH\s*2(?:\b|$)|elite champion\s*2\b/i,elch3:/(?:^|\b)(?:SD[- ]?)?ELCH\s*3(?:\b|$)|elite champion\s*3\b/i,elch4:/(?:^|\b)(?:SD[- ]?)?ELCH\s*4(?:\b|$)|elite champion\s*4\b/i,
 gmch:/(?:^|\b)(?:SD[- ]?)?GMCH(?:\b|$)|grand master champion/i,
 gch:/(?:^|\b)(?:SD[- ]?)?GS(?:\b|$)|games champion(?!\s*[2-5]\b)|^All Games$/i,
 gch2:/(?:^|\b)(?:SD[- ]?)?GCH\s*2(?:\b|$)|games champion\s*2\b|^All Games\s*2$/i,gch3:/(?:^|\b)(?:SD[- ]?)?GCH\s*3(?:\b|$)|games champion\s*3\b|^All Games\s*3$/i,gch4:/(?:^|\b)(?:SD[- ]?)?GCH\s*4(?:\b|$)|games champion\s*4\b|^All Games\s*4$/i,gch5:/(?:^|\b)(?:SD[- ]?)?GCH\s*5(?:\b|$)|games champion\s*5\b|^All Games\s*5$/i
 };return findTitle(p[kind]);
}


function trialCutoffForTitle(title) {
  if (!title) return null;
  const honoursMatch = String(title.honours || "").match(/trial\\s*#?\\s*(\\d+)/i);
  if (honoursMatch) return Number(honoursMatch[1]);
  if (!title.date) return null;

  // Titles exports contain the award date, not the trial number. When
  // several trials share that date, use the latest trial as a conservative
  // cutoff so no score from the award date can be reused accidentally.
  const matchingTrials = [];
  if (!matchingTrials.length) {
    scores.forEach((score) => {
      if (score.date === title.date && score.trial && Number.isFinite(Number(score.trial))) {
        matchingTrials.push(Number(score.trial));
      }
    });
  }
  return matchingTrials.length ? Math.max(...matchingTrials) : null;
}

function rowsAfter(date) {
  if (!date || scores.some((score) => !score.date)) return null;
  return scores.filter((score) => score.date > date);
}

function machTitleStage(title) {
  if (!title) return 0;
  const raw = typeof title === "string" ? title : titleText(title);
  const value = raw.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  const match = value.match(/(?:master champion|mach)\s*([2-4])\b/);
  if (match) return Number(match[1]);
  return /master champion|\bmach\b/.test(value) ? 1 : 0;
}

let machCalculationCacheKey = "";
let machTrialGroupsCache = null;
let machCutoffCache = new Map();

function ensureMachCalculationCache() {
  const relevant = scores.filter((score) => score.level === "Advanced" || score.level === "Excellent")
    .sort((a, b) => a.sourceRow - b.sourceRow);
  const fingerprint = relevant.map((row) => [
    row.sourceRow, row.level, row.component, row.score, row.trial, row.date
  ].join(":")).join("|") + "||" +
    JSON.stringify((titles || []).map((title) => [title.description, title.date, title.honours]));
  if (fingerprint !== machCalculationCacheKey) {
    machCalculationCacheKey = fingerprint;
    machTrialGroupsCache = null;
    machCutoffCache = new Map();
  }
}

function machTrialGroups() {
  ensureMachCalculationCache();
  if (machTrialGroupsCache !== null) return machTrialGroupsCache;
  const relevant = scores.filter((score) => score.level === "Advanced" || score.level === "Excellent").sort((a, b) => a.sourceRow - b.sourceRow);
  if (relevant.some((score) => {
    const trial = Number(score.trial);
    return !score.trial || !Number.isSafeInteger(trial) || trial <= 0;
  })) return null;

  // The uploaded score sheet is chronological from top to bottom. Preserve
  // the first appearance of each Trial Number; never sort Trial Numbers
  // numerically. A trial may have rows separated by other trial rows, so Map
  // groups all rows for the same normalized trial while retaining first-seen order.
  const byTrial = new Map();
  for (const score of relevant) {
    const trial = Number(score.trial);
    if (!byTrial.has(trial)) byTrial.set(trial, { trial, rows: [] });
    byTrial.get(trial).rows.push(score);
  }
  machTrialGroupsCache = [...byTrial.values()];
  return machTrialGroupsCache;
}

function machStageUpperBounds(rows, level, includeTitleLock) {
  const levelRows = rows.filter((row) => row.level === level);
  const titleLock = includeTitleLock ? titleLockForLevel(level, levelRows) : null;
  const lockedRows = new Set(titleLock ? COMPONENTS.map((component) => titleLock.parts[component].sourceRow) : []);
  const componentCount = (component) => levelRows.filter((row) => row.component === component).length;
  const goldCount = (component) => levelRows.filter((row) =>
    !lockedRows.has(row.sourceRow) &&
    row.component === component &&
    ((component === "Container" && row.score > 51) ||
     (component === "Exterior" && row.score > 51) ||
     (component === "Interior" && row.score > 71))
  ).length;
  return {
    combinations: Math.min(10, ...COMPONENTS.map(componentCount)),
    gold: (titleLock?.gold ? 1 : 0) + Math.min(...COMPONENTS.map(goldCount))
  };
}


function selectProgressionCombinations(advanced, excellent, stageKey) {
  const target = stageKey === "ch" ? 5 : 10;
  const minExcellent = stageKey === "ch" ? 2 : stageKey === "mach" ? 3 : stageKey === "mach2" ? 7 : 10;
  const minGold = stageKey === "ch" ? 0 : stageKey === "mach" ? 3 : stageKey === "mach2" || stageKey === "mach3" ? 3 : 4;
  const candidates = [...advanced, ...excellent].filter((set) => set.total >= 100)
    .sort((a, b) => b.total - a.total ||
      Number(a.container?.sourceRow || 0) - Number(b.container?.sourceRow || 0) ||
      Number(a.exterior?.sourceRow || 0) - Number(b.exterior?.sourceRow || 0) ||
      Number(a.interior?.sourceRow || 0) - Number(b.interior?.sourceRow || 0));
  // A recorded title-earning combination is mandatory in its historical
  // block, even when its total is lower than the maximized combinations.
  // Keep it selected; fill the remaining places by the existing score ranking.
  const lockedCandidates = candidates.filter((set) => set.titleLocked);
  const unlockedCandidates = candidates.filter((set) => !set.titleLocked);
  const slots = Math.max(0, target - lockedCandidates.length);
  const excellentCandidates = unlockedCandidates.filter((set) => set.level === "Excellent");
  let selected;
  if (stageKey === "mach3" || stageKey === "mach4") {
    selected = excellentCandidates.slice(0, slots);
  } else if (stageKey === "mach2") {
    selected = excellentCandidates.slice(0, Math.min(7, slots));
    const remaining = unlockedCandidates.filter((set) => !selected.includes(set));
    selected.push(...remaining.slice(0, Math.max(0, slots - selected.length)));
  } else {
    selected = unlockedCandidates.slice(0, slots);
  }
  selected = [...lockedCandidates, ...selected];
  const selectedSet = new Set(selected);
  const replaceFor = (predicate, preserveExcellent = false) => {
    while (selected.filter(predicate).length < (predicate === isExcellent ? minExcellent : minGold)) {
      const replacement = unlockedCandidates.find((candidate) => !selectedSet.has(candidate) && predicate(candidate));
      if (!replacement) break;
      const removable = selected.filter((candidate) => !candidate.titleLocked && !predicate(candidate) &&
        (!preserveExcellent || candidate.level !== "Excellent" || replacement.level === "Excellent" || selected.filter(isExcellent).length > minExcellent))
        .sort((a, b) => a.total - b.total)[0];
      if (!removable) break;
      selected.splice(selected.indexOf(removable), 1, replacement);
      selectedSet.delete(removable);
      selectedSet.add(replacement);
    }
  };
  const isExcellent = (set) => set.level === "Excellent";
  const isGold = (set) => set.gold && (stageKey !== "mach2" || set.level === "Excellent");
  if (stageKey !== "mach3" && stageKey !== "mach4") replaceFor(isExcellent);
  if (minGold) replaceFor(isGold, true);
  // Keep the levels visually separate. Within each level, the title-locked
  // combination is always first, followed by the remaining combinations by score.
  const levelOrder = { Advanced: 0, Excellent: 1 };
  selected.sort((a, b) =>
    (levelOrder[a.level] ?? 2) - (levelOrder[b.level] ?? 2) ||
    Number(!!b.titleLocked) - Number(!!a.titleLocked) ||
    b.total - a.total ||
    Number(a.container?.sourceRow || 0) - Number(b.container?.sourceRow || 0) ||
    Number(a.exterior?.sourceRow || 0) - Number(b.exterior?.sourceRow || 0) ||
    Number(a.interior?.sourceRow || 0) - Number(b.interior?.sourceRow || 0));
  const excellentCount = selected.filter(isExcellent).length;
  const goldCount = selected.filter(isGold).length;
  const met = selected.length >= target && excellentCount >= minExcellent && goldCount >= minGold &&
    (stageKey !== "mach2" || selected.filter((set) => set.level === "Advanced").length <= 3);
  return { sets: selected, advanced: selected.filter((set) => set.level === "Advanced"),
    excellent: selected.filter(isExcellent), total: selected.length, excellentCount, goldCount, met };
}

function machCutoffRowForTitle(title) {
  const stage = machTitleStage(title);
  if (!stage) return null;
  ensureMachCalculationCache();
  if (machCutoffCache.has(stage)) return machCutoffCache.get(stage);
  const trialGroups = machTrialGroups();
  if (!trialGroups) return null;

  // The Dog Titles file supplies the award date. Reconstruct the actual
  // completion trial by checking each trial on that date in score-file order.
  // This matters when a dog runs multiple trials on the same calendar date.
  let cutoffIndex = -1;
  for (let currentStage = 1; currentStage <= stage; currentStage++) {
    if (machCutoffCache.has(currentStage)) {
      const cached = machCutoffCache.get(currentStage);
      if (!cached) return null;
      cutoffIndex = cached.index;
      continue;
    }

    const stageKey = currentStage === 1 ? "mach" : currentStage === 2 ? "mach2" : currentStage === 3 ? "mach3" : "mach4";
    const recorded = championshipTitle(stageKey);
    if (!titlesFileLoaded || !recorded || !recorded.date) {
      machCutoffCache.set(currentStage, null);
      return null;
    }
    const awardDate = recorded.date;
    const honourTrial = String(recorded.honours || "").match(/trial\s*#?\s*(\d+)/i);
    const prefix = [];
    let foundCutoff = false;

    for (let groupIndex = cutoffIndex + 1; groupIndex < trialGroups.length; groupIndex++) {
      const group = trialGroups[groupIndex];
      prefix.push(...group.rows);

      // Do not let a trial from an earlier date establish this title's
      // cutoff. Only trials on the recorded award date are candidates.
      const groupDates = [...new Set(group.rows.map((row) => row.date).filter(Boolean))];
      if (!groupDates.includes(awardDate)) continue;
      if (honourTrial && Number(honourTrial[1]) !== Number(group.trial)) continue;

      const includeTitleLock = currentStage === 1;
      const advancedLimit = currentStage === 1 ? 10 : currentStage === 2 ? 3 : 0;
      const advancedUpper = advancedLimit ? machStageUpperBounds(prefix, "Advanced", includeTitleLock) : { combinations: 0, gold: 0 };
      const excellentUpper = machStageUpperBounds(prefix, "Excellent", includeTitleLock);
      const couldMeet = currentStage === 1
        ? advancedUpper.combinations + excellentUpper.combinations >= 10 &&
          excellentUpper.combinations >= 3 && advancedUpper.gold + excellentUpper.gold >= 3
        : currentStage === 2
          ? excellentUpper.combinations >= 7 && excellentUpper.gold >= 3 &&
            excellentUpper.combinations + Math.min(advancedUpper.combinations, 3) >= 10
          : excellentUpper.combinations >= 10 &&
            excellentUpper.gold >= (currentStage === 3 ? 3 : 4);
      if (!couldMeet) continue;

      const advanced = advancedLimit
        ? maximizeGoldCombinations("Advanced", prefix, advancedLimit, includeTitleLock)
        : [];
      const excellent = maximizeGoldCombinations("Excellent", prefix, 10, includeTitleLock);
      const selection = selectProgressionCombinations(advanced, excellent, stageKey);
      const searchComplete = advanced.optimizationComplete !== false &&
        excellent.optimizationComplete !== false;
      if (!searchComplete && !selection.met) return null;
      if (selection.met) {
        cutoffIndex = groupIndex;
        foundCutoff = true;
        break;
      }
    }

    if (!foundCutoff) {
      machCutoffCache.set(currentStage, null);
      return null;
    }
    machCutoffCache.set(currentStage, {
      trial: trialGroups[cutoffIndex].trial,
      index: cutoffIndex,
      date: awardDate
    });
  }
  return machCutoffCache.get(stage) || null;
}




function historicalChProgressSelection() {
  const recorded = championshipTitle("ch");
  const awardDate = workbookDate(recorded?.date);
  if (!recorded || !awardDate) return null;
  const orderedScores = scores.filter((row) => row.level === "Advanced" || row.level === "Excellent")
    .sort((a, b) => a.sourceRow - b.sourceRow);
  const groups = machTrialGroups() || orderedScores.map((row) => ({ trial: row.trial, rows: [row] }));
  const honourTrial = String(recorded.honours || "").match(/trial\s*#?\s*(\d+)/i);
  const prefix = [];
  let lastSelection = null;
  for (const group of groups) {
    const groupDates = group.rows.map((row) => workbookDate(row.date)).filter(Boolean);
    const groupDate = groupDates.length ? groupDates[0] : "";
    if (groupDate && groupDate > awardDate) break;
    prefix.push(...group.rows);
    const advanced = maximizeGoldCombinations("Advanced", prefix, 10, true);
    const excellent = maximizeGoldCombinations("Excellent", prefix, 10, true);
    lastSelection = selectProgressionCombinations(advanced, excellent, "ch");
    if (honourTrial) {
      if (Number(honourTrial[1]) === Number(group.trial)) return lastSelection;
    } else if (lastSelection.met) {
      return lastSelection;
    }
  }
  return lastSelection;
}


function historicalMachProgressSelection() {
  const recorded = championshipTitle("mach");
  const awardDate = workbookDate(recorded?.date);
  if (!recorded || !awardDate) return null;

  // Rebuild the original SD-MACH qualification pool as it existed at the
  // recorded award date. Do not use the active post-MACH ledger here.
  const orderedScores = scores.filter((row) => row.level === "Advanced" || row.level === "Excellent")
    .sort((a, b) => a.sourceRow - b.sourceRow);
  const groups = machTrialGroups() || orderedScores.map((row) => ({ trial: row.trial, rows: [row] }));
  const honourTrial = String(recorded.honours || "").match(/trial\s*#?\s*(\d+)/i);
  const prefix = [];
  let lastSelection = null;

  for (const group of groups) {
    const groupDates = group.rows.map((row) => workbookDate(row.date)).filter(Boolean);
    const groupDate = groupDates.length ? groupDates[0] : "";
    if (groupDate && groupDate > awardDate) break;

    prefix.push(...group.rows);
    const advanced = maximizeGoldCombinations("Advanced", prefix, 10, true);
    const excellent = maximizeGoldCombinations("Excellent", prefix, 10, true);
    lastSelection = selectProgressionCombinations(advanced, excellent, "mach");

    // Honours may identify the exact title-earning trial when several trials
    // occurred on the same award date. Otherwise stop at earliest completion.
    if (honourTrial) {
      if (Number(honourTrial[1]) === Number(group.trial)) return lastSelection;
    } else if (lastSelection.met) {
      return lastSelection;
    }
  }
  return lastSelection;
}

function chronologicalChStory() {
  const groups = machTrialGroups();
  if (!groups) return { unavailable: "Valid trial numbers are required to establish the chronological SD-CH completion point." };
  const prefix = [], steps = [];
  for (const group of groups) {
    prefix.push(...group.rows);
    const advanced = maximizeGoldCombinations("Advanced", prefix, 10, true);
    const excellent = maximizeGoldCombinations("Excellent", prefix, 10, true);
    const selection = selectProgressionCombinations(advanced, excellent, "ch");
    steps.push({ trial: group.trial, date: group.rows.find((row) => row.date)?.date || "",
      total: selection.total, excellent: selection.excellent.length, met: selection.met });
    if (selection.met) return { completion: { trial: group.trial, date: group.rows.find((row) => row.date)?.date || "",
      sets: selection.sets, total: selection.total, excellent: selection.excellent.length }, steps, recorded: championshipTitle("ch") };
  }
  return { completion: null, steps, recorded: championshipTitle("ch") };
}

function chronologicalMachStory() {
  const groups = machTrialGroups();
  if (!groups) return { unavailable: "The score file is missing valid trial numbers. Chronological title progression cannot be established." };
  const stages = [
    { key: "mach", name: "SD-MACH", excellent: 3, gold: 3, advanced: 10 },
    { key: "mach2", name: "SD-MACH 2", excellent: 7, gold: 3, advanced: 3 },
    { key: "mach3", name: "SD-MACH 3", excellent: 10, gold: 3, advanced: 0 },
    { key: "mach4", name: "SD-MACH 4", excellent: 10, gold: 4, advanced: 0 }
  ];
  const story = [];
  let previousCutoffIndex = -1;
  for (const stage of stages) {
    const prefix = [];
    const steps = [];
    let completion = null;
    for (let index = previousCutoffIndex + 1; index < groups.length; index++) {
      prefix.push(...groups[index].rows);
      const includeTitleLock = stage.key === "mach";
      const advanced = stage.advanced
        ? maximizeGoldCombinations("Advanced", prefix, stage.advanced, includeTitleLock)
        : [];
      const excellent = maximizeGoldCombinations("Excellent", prefix, 10, includeTitleLock);
      const selection = selectProgressionCombinations(advanced, excellent, stage.key);
      const gold = selection.sets.filter((set) => set.gold).length;
      const excellentGold = selection.excellent.filter((set) => set.gold).length;
      const total = selection.total;
      const met = selection.met;
      steps.push({
        trial: groups[index].trial,
        date: groups[index].rows.find((row) => row.date)?.date || "",
        advanced: selection.advanced.length,
        excellent: selection.excellent.length,
        gold,
        excellentGold,
        met
      });
      if (met) {
        completion = {
          index,
          trial: groups[index].trial,
          date: groups[index].rows.find((row) => row.date)?.date || "",
          advanced: selection.advanced,
          excellent: selection.excellent,
          gold,
          excellentGold,
          total
        };
        break;
      }
    }
    const recorded = championshipTitle(stage.key);
    story.push({ ...stage, steps, completion, recorded });
    if (!completion) break;
    // A repeat stage may only begin after the preceding title is confirmed in
    // the uploaded Titles sheet and its recorded date agrees with the score history.
    if (!titlesFileLoaded || !recorded || !recorded.date || recorded.date !== completion.date) break;
    const honourTrial = String(recorded.honours || "").match(/trial\s*#?\s*(\d+)/i);
    if (honourTrial && Number(honourTrial[1]) !== Number(completion.trial)) break;
    previousCutoffIndex = completion.index;
  }
  return { stages: story, totalTrials: groups.length };
}

function renderChronologicalMachStory() {
  const host = $("progression-story");
  if (!host) return;
  if (!titlesFileLoaded) {
    host.innerHTML = '<p class="level-note">Upload the Dog Titles file to compare calculated completion points with SDDA-recorded title dates. Score-only calculations can still be explored, but recorded title validation is unavailable.</p>';
  }
  const result = chronologicalMachStory();
  if (result.unavailable) {
    host.innerHTML = '<p class="level-note">' + esc(result.unavailable) + '</p>';
    return;
  }
  const classification = (total) => total >= 192 ? "Gold" : total >= 172 ? "Silver" : total >= 142 ? "Bronze" : total >= 100 ? "Satisfactory" : "Insufficient";
  const stagesHtml = result.stages.map((stage) => {
    const recordedText = stage.recorded
      ? 'SDDA title sheet: recorded ' + esc(dateLabel(stage.recorded.date)) + '.'
      : (titlesFileLoaded ? 'No matching SDDA award record was found in the uploaded Titles file.' : 'No Titles file loaded; official award-date comparison is unavailable.');
    const steps = stage.steps.map((step) =>
      '<li><strong>Trial ' + esc(step.trial) + (step.date ? ' · ' + esc(dateLabel(step.date)) : '') + '</strong>: ' +
      esc(step.advanced) + ' Advanced + ' + esc(step.excellent) + ' Excellent combinations; ' +
      esc(step.gold) + ' Gold overall (' + esc(step.excellentGold) + ' Excellent Gold).' +
      (step.met ? ' <strong>First qualifying point — title requirements reached.</strong>' : '') +
      '</li>'
    ).join("");
    const selected = stage.completion
      ? '<p><strong>Cutoff established at Trial ' + esc(stage.completion.trial) + '.</strong> The next championship stage starts with the following trial.</p>' +
        '<p>At the cutoff, the high-to-low pairing contained ' + esc(stage.completion.total) + ' qualifying combinations, including ' +
        esc(stage.completion.excellent.length) + ' Excellent and ' + esc(stage.completion.gold) + ' Gold overall.</p>'
      : '<p><strong>Not reached in the uploaded score history.</strong> The available trials were checked in chronological order.</p>';
    return '<article class="next-card"><header><div><h3>' + esc(stage.name) + '</h3><p class="next-kind">' +
      (stage.completion ? 'First reached at Trial ' + esc(stage.completion.trial) : 'Still in progress') +
      '</p></div><span class="remaining-badge">' + (stage.completion ? 'Reached' : 'Not reached') + '</span></header>' +
      '<p>' + recordedText + '</p>' + selected + '<details><summary>Trial-by-trial calculation (' + stage.steps.length + ' trials)</summary><ol>' + steps + '</ol></details></article>';
  }).join("");
  host.innerHTML = '<p class="level-note">Trials are processed in the order they appear in the score export. After each complete trial is added, the eligible pool is re-paired high-to-low and checked against the title rules. The first trial that satisfies every requirement becomes the cutoff; later trials are not used to move that title earlier or improve its selected combinations.</p>' +
    (stagesHtml || '<p class="level-note">No complete championship calculation could be produced from the current score history.</p>');
}


function progressionRecordForLevel(level) {
  return titles.find((title) => {
    const value = String(title.description || "").trim().toLowerCase();
    return value === level.toLowerCase() || value === (level + " special").toLowerCase();
  }) || null;
}

function progressionIndividualSet(level) {
  if (level === "Advanced" || level === "Excellent") return firstCompletedTitleSet(level);
  const groups = chronologicalTrialGroups(level, byLevel(level));
  const seen = [], passed = new Set();
  for (const group of groups) {
    const rows = [...group.rows].sort((a, b) => a.sourceRow - b.sourceRow);
    const newlyPassed = rows.filter((row) => !passed.has(row.component));
    const here = new Set(rows.map((row) => row.component));
    seen.push(...rows);
    rows.forEach((row) => passed.add(row.component));
    if (!COMPONENTS.every((component) => passed.has(component))) continue;
    const completion = newlyPassed[newlyPassed.length - 1];
    if (!completion) return null;
    const special = COMPONENTS.every((component) => here.has(component));
    const parts = Object.fromEntries(COMPONENTS.map((component) => {
      const eligible = special || component === completion.component ? rows : seen;
      const selected = eligible.filter((row) => row.component === component)
        .sort((a, b) => b.score - a.score || a.sourceRow - b.sourceRow)[0];
      return [component, selected];
    }));
    if (!COMPONENTS.every((component) => parts[component])) return null;
    return { level, trial: group.trial, parts, total: COMPONENTS.reduce((sum, component) => sum + parts[component].score, 0), special, completionComponent: completion.component, completionDate: completion.date || "" };
  }
  return null;
}

function progressionCompletionDate(level, set) {
  if (!set) return "";
  if (set.completionDate) return set.completionDate;
  const component = set.completionComponent;
  if (!component) return "";
  const row = scores.filter((score) => score.level === level && score.component === component && String(score.trial) === String(set.completionTrial || set.trial))
    .sort((a, b) => a.sourceRow - b.sourceRow).at(-1);
  return row?.date || "";
}

function progressionScoreRecord(row) {
  return row ? esc(row.score) + ' (' + esc(row.trial || 'trial not supplied') + ')' : '—';
}

function progressionSetDetails(set, label) {
  if (!set) return '<p>No complete score set can be reconstructed from the supplied history.</p>';
  const parts = set.parts || { Container: set.container, Exterior: set.exterior, Interior: set.interior };
  const total = Number(set.total ?? COMPONENTS.reduce((sum, component) => sum + Number(parts[component]?.score || 0), 0));
  const classification = total >= 192 ? 'Gold' : total >= 172 ? 'Silver' : total >= 142 ? 'Bronze' : total >= 100 ? 'Satisfactory' : 'Insufficient';
  return '<p><strong>' + esc(label) + ':</strong> Container ' + progressionScoreRecord(parts.Container) +
    ' · Exterior ' + progressionScoreRecord(parts.Exterior) +
    ' · Interior ' + progressionScoreRecord(parts.Interior) +
    ' · Total ' + esc(total) + ' (' + classification + ').</p>';
}

function progressionTitleComparison(calculatedDate, recorded, calculatedTrial) {
  if (!titlesFileLoaded) return '<p class="level-note">No Dog Titles file loaded. This is a score-based calculation, not confirmation of an official award.</p>';
  if (!recorded) return '<p class="level-note">No matching award appears in the uploaded Dog Titles file. The score calculation is not an official SDDA award.</p>';
  if (!calculatedDate || !recorded.date) return '<p class="level-note">A calculated completion or recorded award date is missing; the two cannot be compared.</p>';
  if (calculatedDate !== recorded.date) return '<p class="level-note">Date discrepancy: the score history first meets the calculated requirements on ' + esc(dateLabel(calculatedDate)) + ', while the Titles file records ' + esc(dateLabel(recorded.date)) + '. This needs review; the checker will not force these dates to match.</p>';
    const honourTrial = String(recorded.honours || "").match(/trial\s*#?\s*(\d+)/i);
  if (honourTrial && calculatedTrial && Number(honourTrial[1]) !== Number(calculatedTrial)) {
    return '<p class="level-note">Trial discrepancy: the score calculation first completes at Trial ' + esc(calculatedTrial) + ', but the Titles file honours field points to Trial ' + esc(honourTrial[1]) + '. Review the source records.</p>';
  }
  return '<p class="level-note">The calculated completion date matches the recorded title date' +
    (calculatedTrial ? ' (calculated completion Trial ' + esc(calculatedTrial) + ')' : '') +
    '. This supports the record comparison but does not replace SDDA approval.</p>';
}


function progressionEliteTitleCutoffs(groups) {
  const records = [
    findTitle(/^Elite(?:\s|$)/i),
    championshipTitle("elch"), championshipTitle("elch2"),
    championshipTitle("elch3"), championshipTitle("elch4")
  ].filter(Boolean);
  const cutoffs = new Map();
  for (const record of records) {
    const honourTrial = String(record.honours || "").match(/trial\s*#?\s*(\d+)/i);
    let index = honourTrial ? groups.findIndex((group) => Number(group.trial) === Number(honourTrial[1])) : -1;
    if (index < 0 && record.date) groups.forEach((group, groupIndex) => {
      if (group.rows.some((row) => row.date && row.date <= record.date)) index = groupIndex;
    });
    if (index >= 0) cutoffs.set(String(groups[index].trial), record.description || "Elite title");
  }
  return cutoffs;
}
function progressionHistoryMarkup(level) {
  // Keep each level's score rows separate. MACH requirements combine selected
  // Advanced and Excellent combinations, but the shared cutoff is only a
  // marker: it must appear in both histories without importing the other level's scores.
  const groups = chronologicalTrialGroups(level, scores);
  const lock = level === "Advanced" || level === "Excellent" ? titleLockForLevel(level) : null;
  const lockedRows = new Set(lock ? COMPONENTS.map((component) => lock.parts[component].sourceRow) : []);
  const cutoffs = new Map();
  const cutoffDates = new Map();
  if (level === "Advanced" || level === "Excellent") {
    for (const stage of [{key:"mach",label:"MACH cutoff"},{key:"mach2",label:"MACH2 cutoff"},{key:"mach3",label:"MACH3 cutoff"},{key:"mach4",label:"MACH4 cutoff"}]) {
      if (!championshipTitle(stage.key)) continue;
      const cutoff = machCutoffRowForTitle(stage.key);
      if (cutoff) {
        cutoffs.set(String(cutoff.trial), stage.label);
        cutoffDates.set(String(cutoff.trial), cutoff.date);
      }
    }
  } else if (level === "Elite") progressionEliteTitleCutoffs(groups).forEach((label, trial) => cutoffs.set(trial, label));

  // Advanced and Excellent scores become historical only once a MACH
  // cutoff has actually been reconstructed and verified. Without a verified
  // cutoff, every score remains current and this history panel stays empty.
  const isMachLevel = level === "Advanced" || level === "Excellent";
  const allMachGroups = isMachLevel ? (machTrialGroups() || []) : [];
  if (isMachLevel && !cutoffs.size) {
    return '<p class="run-empty">No historical ' + esc(level) +
      ' scores yet. Scores remain current until a MACH cutoff is met and verified.</p>';
  }

  // If this level has no score in the exact MACH title trial, add an empty
  // group at the same chronological position so its shared cutoff is visible.
  const historyGroups = [...groups];
  if (isMachLevel) {
    for (const trial of cutoffs.keys()) {
      if (historyGroups.some((group) => String(group.trial) === String(trial))) continue;
      const globalIndex = allMachGroups.findIndex((group) => String(group.trial) === String(trial));
      if (globalIndex < 0) continue;
      let insertAt = historyGroups.findIndex((group) => {
        const groupIndex = allMachGroups.findIndex((allGroup) => String(allGroup.trial) === String(group.trial));
        return groupIndex > globalIndex;
      });
      if (insertAt < 0) insertAt = historyGroups.length;
      historyGroups.splice(insertAt, 0, { trial, rows: [] });
    }

    // Only retain score-file rows through the latest verified MACH cutoff.
    const latestCutoffTrial = [...cutoffs.keys()].at(-1);
    const latestCutoffIndex = allMachGroups.findIndex((group) =>
      String(group.trial) === String(latestCutoffTrial));
    if (latestCutoffIndex >= 0) {
      const cutoffRows = allMachGroups[latestCutoffIndex].rows;
      const lastCutoffSourceRow = Math.max(0, ...cutoffRows.map((row) => Number(row.sourceRow) || 0));
      const cutoffTrialSet = new Set(allMachGroups.slice(0, latestCutoffIndex + 1).map((group) => String(group.trial)));
      for (let i = historyGroups.length - 1; i >= 0; i--) {
        const group = historyGroups[i];
        if (!cutoffTrialSet.has(String(group.trial))) {
          historyGroups.splice(i, 1);
          continue;
        }
        group.rows = group.rows.filter((row) => (Number(row.sourceRow) || 0) <= lastCutoffSourceRow);
      }
    }
  }
  if (isMachLevel) {
    // Historical combinations are rebuilt only from score rows through the
    // latest verified MACH cutoff. The normal optimizer enforces component
    // non-reuse, keeps the title-lock combination first, and sorts the rest
    // from highest total to lowest.
    const historicalRows = historyGroups.flatMap((group) => group.rows);
    const historicalCombinations = maximizeGoldCombinations(level, historicalRows, 10, true);
    const cutoffSummary = [...cutoffs.entries()].map(([trial, label]) =>
      '<span class="history-cutoff"><b>' + esc(label) + '</b><small>Trial ' + esc(trial) +
      (cutoffDates.has(String(trial)) ? ' · ' + esc(dateLabel(cutoffDates.get(String(trial)))) : '') +
      '</small></span>').join('');
    return '<div class="history-panel-intro"><p>Historical combinations through the latest verified MACH cutoff. The title-lock combination is shown first; remaining combinations are sorted by total score. Components are not reused.</p>' +
      cutoffSummary + '</div>' +
      '<div class="historical-combination-rows">' + progressionRunRows(historicalCombinations) + '</div>';
  }
  if (!historyGroups.length) return '<p class="run-empty">No ' + esc(level) + ' score history is available.</p>';
  const columns = COMPONENTS.map((component) => {
    let output = "";
    for (const group of historyGroups) {
      const rows = group.rows.filter((row) => row.component === component);
      output += rows.map((row) => '<div class="history-score ' + (lockedRows.has(row.sourceRow) ? 'is-title-lock' : '') +
        '"><span>' + esc(row.score) + '</span><small>(' + esc(row.trial || "—") + ')</small>' +
        (lockedRows.has(row.sourceRow) ? '<b>Title lock</b>' : '') + '</div>').join("");
      if (cutoffs.has(String(group.trial))) output += '<div class="history-cutoff"><b>' +
        esc(cutoffs.get(String(group.trial))) + '</b><small>Trial ' + esc(group.trial) + (cutoffDates.has(String(group.trial)) ? ' · ' + esc(dateLabel(cutoffDates.get(String(group.trial)))) : '') +
        ' · scores through this title date/trial are not carried forward</small></div>';
    }
    return '<section class="history-component"><h4>' + esc(component) + '</h4>' +
      (output || '<p class="run-empty">No scores</p>') + '</section>';
  }).join("");
  const lockSummary = lock ? '<div class="history-title-lock"><b>' + esc(level) +
    ' title-lock combination</b><div class="history-lock-parts">' + COMPONENTS.map((component) => {
      const part = lock.parts[component];
      return '<span><strong>' + esc(component) + '</strong><em>' + esc(part.score) +
        ' (' + esc(part.trial || "—") + ')</em></span>';
    }).join("") + '</div><strong>Total ' + esc(lock.total) + '</strong></div>' : "";
  return '<div class="history-panel-intro">' + lockSummary +
    '<p>Scores remain in score-file order. Trial numbers in brackets are identifiers, not sort order. Red circles mark the three records permanently locked together for the regular title. Yellow lines mark championship cutoffs only.</p></div>' +
    '<div class="history-component-columns">' + columns + '</div>';
}
function progressionRunRows(sets, isElite = false) {
  const classification = (total) => total >= 192 ? "Gold" : total >= 172 ? "Silver" : total >= 142 ? "Bronze" : total >= 100 ? "Satisfactory" : "Insufficient";
  const qualifying = sets.filter((set) => set.total >= 100);
  if (!qualifying.length) return '<p class="run-empty">No qualifying combinations yet.</p>';
  return '<div class="run-grid">' + qualifying.map((set, index) => {
    const container = isElite ? set.parts?.Container : set.container;
    const exterior = isElite ? set.parts?.Exterior : set.exterior;
    const interior = isElite ? set.parts?.Interior : set.interior;
    const show = (part) => part ? esc(part.score) + ' (' + esc(part.trial || "—") + ')' : "—";
    return '<div class="run-row ' + (set.gold ? "gold" : classification(set.total).toLowerCase()) + '">' +
      '<b>' + (isElite ? "Trial " + esc(set.trial || "—") : set.special ? "SP" : set.titleLocked ? "Title lock" : "#" + (index + 1)) +
      '</b><span>Container: ' + show(container) + '</span><span>Exterior: ' + show(exterior) +
      '</span><span>Interior: ' + show(interior) + '</span><strong>Total: ' + esc(set.total) +
      ' · ' + classification(set.total) + '</strong></div>';
  }).join("") + '</div>';
}
function progressionCurrentScoreRows(rows, level) {
  const entries = rows.filter((row) => row.level === level)
    .sort((a, b) => a.sourceRow - b.sourceRow);
  // Match History: keep one component column per search area and list each
  // score with its trial identifier in brackets. Empty components stay blank.
  return '<div class="history-component-columns">' + COMPONENTS.map((component) => {
    const componentRows = entries.filter((row) => row.component === component);
    const output = componentRows.map((row) =>
      '<div class="history-score"><span>' + esc(row.score) + '</span><small>(' +
      esc(row.trial || "—") + ')</small></div>'
    ).join("");
    return '<section class="history-component"><h4>' + esc(component) + '</h4>' +
      output + '</section>';
  }).join("") + '</div>';
}

function renderProgressionScoreColumns(activeLedger = currentMachLedgerRows()) {
  const advancedRows = activeLedger.verified ? activeLedger.rows.filter((row) => row.level === "Advanced") : [];
  const excellentRows = activeLedger.verified ? activeLedger.rows.filter((row) => row.level === "Excellent") : [];
  const advanced = activeLedger.verified ? maximizeGoldCombinations("Advanced", activeLedger.rows) : [];
  const excellent = activeLedger.verified ? maximizeGoldCombinations("Excellent", activeLedger.rows) : [];
  const eliteTitle = [championshipTitle("elch4"), championshipTitle("elch3"), championshipTitle("elch2"),
    championshipTitle("elch"), findTitle(/^Elite/i)].find(Boolean);
  const eliteRows = eliteTitle ? scoreRowsAfterTitle(eliteTitle) : scores.filter((row) => row.level === "Elite");
  const elite = eliteRows ? eliteSets(eliteRows) : [];
  const cards = [
    {level:"Advanced",sets:advanced,rows:advancedRows,accent:"level-accent-silver",elite:false},
    {level:"Excellent",sets:excellent,rows:excellentRows,accent:"level-accent-gold",elite:false},
    {level:"Elite",sets:elite,rows:eliteRows || [],accent:"level-accent-elite",elite:true}
  ];
  return '<p class="level-note current-cutoff-announcement">' + esc(activeLedger.message) +
    '</p>' + 
    '<div class="progression-score-columns">' +
    cards.map((card) => {
      const completeCount = card.elite ? card.sets.length : card.sets.filter((set) => set.total >= 100).length;
      const goldCount = card.sets.filter((set) => set.gold && (card.elite || set.total >= 100)).length;
      const count = card.elite
        ? card.sets.length + ' complete same-trial · ' + goldCount + ' Gold'
        : card.rows.length + ' component scores · ' + completeCount + ' complete combinations · ' + goldCount + ' Gold';
      const current = card.elite
        ? progressionRunRows(card.sets, true)
        : !activeLedger.verified
          ? '<p class="run-empty">Current scores are withheld until the recorded MACH cutoff can be verified.</p>'
          : progressionCurrentScoreRows(card.rows, card.level);
      return '<details class="progression-score-column title-accordion ' + card.accent + '">' +
        '<summary class="progression-card-header"><span class="summary-copy"><strong>' + card.level +
        '</strong><small>' + count + '</small></span><span class="summary-action">Details</span></summary>' +
        '<div class="progression-expanded-body"><section class="progression-current-combinations">' +
        '<h4>' + (card.elite ? 'Current combinations' : 'Current scores') + '</h4>' + current + '</section>' +
        '<section class="progression-history"><h4>' + card.level + ' history</h4>' +
        progressionHistoryMarkup(card.level) + '</section></div></details>';
    }).join("") + '</div>';
}
function currentMachLedgerRows() {
  const stages = [
    { key: "mach4", title: "mach4" },
    { key: "mach3", title: "mach3" },
    { key: "mach2", title: "mach2" },
    { key: "mach", title: "mach" }
  ];
  const recordedStage = stages.find((stage) => !!championshipTitle(stage.title));
  if (!recordedStage) {
    return { verified: true, rows: scores, cutoffTrial: null, message: "Full score history · no recorded MACH title cutoff." };
  }
  const cutoff = machCutoffRowForTitle(recordedStage.title);
  const groups = machTrialGroups();
  if (!cutoff || !groups || !groups[cutoff.index]) {
    return {
      verified: false,
      rows: [],
      cutoffTrial: null,
      message: recordedStage.key.toUpperCase() + " is recorded, but its chronological cutoff could not be verified against the score export and Dog Titles sheet. The active combination list is withheld rather than showing pre-title scores."
    };
  }
  // Exclude the entire cutoff trial, not merely the three rows selected for
  // the title. sourceRow is the score export's chronology; Trial Number is
  // only an identifier and is never used to order records.
  const cutoffTrialNumber = Number(cutoff.trial);
  const lastCutoffSourceRow = scores.reduce((last, score) =>
    Number(score.trial) === cutoffTrialNumber
      ? Math.max(last, Number(score.sourceRow) || 0)
      : last, 0);
  if (!lastCutoffSourceRow) {
    return { verified: false, rows: [], cutoffTrial: cutoff.trial,
      message: "The recorded " + recordedStage.key.toUpperCase() + " cutoff trial could not be located in the full score export. The active combination list is withheld." };
  }
  return {
    verified: true,
    rows: scores.filter((score) => (Number(score.sourceRow) || 0) > lastCutoffSourceRow),
    cutoffTrial: cutoff.trial,
    lastSourceRow: lastCutoffSourceRow,
    message: "Active combinations use score-file rows after the complete " + recordedStage.key.toUpperCase() + " cutoff trial " + cutoff.trial + ". All rows from that trial are excluded."
  };
}

function machExtensionProgress(rule, priorTitleOrStage) {
  const cutoff = machCutoffRowForTitle(priorTitleOrStage);
  const trialGroups = machTrialGroups();
  if (!cutoff || !trialGroups) {
    return { verified: false, cutoffTrial: cutoff?.trial ?? null, advanced: 0, excellent: 0, gold: 0 };
  }

  // Select whole trial groups after the previous title. Trial numbers are
  // identifiers, not dates, and must never be compared numerically.
  const eligibleTrials = new Set(trialGroups.slice(cutoff.index + 1).map((group) => group.trial));
  const eligible = scores.filter((score) =>
    (score.level === "Advanced" || score.level === "Excellent") &&
    eligibleTrials.has(Number(score.trial))
  );
  const advancedLimit = rule.maxAdvanced || 0;
  const advanced = advancedLimit
    ? maximizeGoldCombinations("Advanced", eligible, advancedLimit, false)
    : [];
  const excellent = maximizeGoldCombinations("Excellent", eligible, 10, false);
  const selection = selectProgressionCombinations(advanced, excellent, rule.key);
  return {
    verified: true,
    cutoffTrial: cutoff.trial,
    advanced: selection.advanced.length,
    excellent: selection.excellent.length,
    gold: selection.excellent.filter((set) => set.gold).length,
    selected: selection.sets,
    met: selection.met
  };
}
function scoreRowsAfterTitle(priorTitle) {
  if (!priorTitle) return null;
  const eliteRows = scores.filter((score) => score.level === "Elite").sort((a, b) => a.sourceRow - b.sourceRow);
  if (eliteRows.some((score) => !score.trial || !Number.isFinite(Number(score.trial)))) return null;
  const groups = chronologicalTrialGroups("Elite", eliteRows);
  if (!groups.length) return [];

  // Prefer an explicit trial number in the title record. If only a title date
  // is available, conservatively exclude every Elite trial on that date.
  const honoursMatch = String(priorTitle.honours || "").match(/trial\s*#?\s*(\d+)/i);
  let cutoffIndex = -1;
  if (honoursMatch) {
    const trialNumber = Number(honoursMatch[1]);
    cutoffIndex = groups.findIndex((group) => Number(group.trial) === trialNumber);
  }
  if (cutoffIndex < 0 && priorTitle.date) {
    groups.forEach((group, index) => {
      if (group.rows.some((row) => row.date && row.date <= priorTitle.date)) cutoffIndex = index;
    });
  }
  if (cutoffIndex < 0) return null;
  return groups.slice(cutoffIndex + 1).flatMap((group) => group.rows);
}

function render() {
  if (!scores.length) { $("report").hidden = true; return; }
  if (!titlesFileLoaded) {
    $("report").hidden = true;
    $("message").textContent = "Dog Titles file required. Upload the Dog Titles CSV or XLSX to calculate title progress and championship cutoff dates.";
    return;
  }
  const dogNumbers = [...new Set(scores.map((score) => score.dogNumber).filter(Boolean))];
  if (dogNumbers.length > 1) { $("message").textContent = "The score file contains more than one dog. Please download and upload one dog's history at a time."; $("report").hidden = true; return; }
  $("report").hidden = false;
  const dog = dogDisplayName(), number = scores[0]?.dogNumber || "not supplied";
  document.title = reportFileStem();
  const summaryInferred = inferredTitles();
  const summaryTitle = (pattern) => titles.find((title) => pattern.test(titleText(title))) || summaryInferred.find((title) => pattern.test(titleText(title)));
  const summaryAdvanced = maximizeGoldCombinations("Advanced"), summaryExcellent = maximizeGoldCombinations("Excellent"), summaryElite = eliteSets();
  const summaryChSelection = championshipTitle("ch")
    ? (historicalChProgressSelection() || { sets: [], advanced: [], excellent: [], total: 0, met: false })
    : selectProgressionCombinations(summaryAdvanced, summaryExcellent, "ch");
  const summaryHistoricalMachSelection = championshipTitle("mach") ? historicalMachProgressSelection() : null;
  const summaryMachSelection = summaryHistoricalMachSelection || selectProgressionCombinations(summaryAdvanced, summaryExcellent, "mach");
  const scentAchieved = [
    !!summaryTitle(/^Started(?:\s|$)/i), !!summaryTitle(/^Advanced(?:\s|$)/i), !!summaryTitle(/^Excellent(?:\s|$)/i),
    !!titles.find((title) => /^Elite(?:\s|$)/i.test(titleText(title))) || ((!!summaryTitle(/^Excellent(?:\s|$)/i) || COMPONENTS.every((component) => componentCounts("Excellent")[component] > 0)) && summaryElite.length > 0),
    !!championshipTitle("ch") || summaryChSelection.met,
    !!championshipTitle("mach") || summaryMachSelection.met,
    !!championshipTitle("elch") || (summaryElite.length >= 5 && summaryElite.filter((set) => set.gold).length >= 2),
    !!championshipTitle("gmch")
  ].filter(Boolean).length;
  $("dog-name").textContent = dog; $("dog-meta").textContent = `SDDA #${number} · ${scores.length} qualifying component scores · ${scentAchieved} scent titles achieved · ${8 - scentAchieved} in progress`;
  $("message").textContent = titles.length
    ? "Scores and recorded titles loaded. Results are saved on this device."
    : "Scores loaded. Achievements are calculated from score history and are not official SDDA confirmation.";
  window.sddaGoldOptimizationWarning = [summaryAdvanced, summaryExcellent].some((sets) => sets.optimizationComplete === false)
    ? "The Gold search reached its safety limit. Displayed combinations are the best found so far, not a proven maximum."
    : "";
  if (window.sddaGoldOptimizationWarning) $("message").textContent += " " + window.sddaGoldOptimizationWarning;

  const shownTitles = displayedTitles();
  $("achieved").innerHTML = shownTitles.map((title) => `<article class="title-card"><h3>${esc(title.description)}</h3><span class="status met">${title.calculated ? "Calculated from scores" : "Recorded by SDDA"}</span><div class="score">${title.score == null || Number.isNaN(title.score) ? "—" : esc(title.score)}</div><p>${title.score == null || Number.isNaN(title.score) ? "No title score supplied" : title.calculated ? "Calculated component total" : "Recorded title score"}</p><p><strong>${esc(dateLabel(title.date))}</strong>${title.honours ? ` · ${esc(title.honours)}` : ""}</p></article>`).join("") || `<article class="title-card"><h3>No complete title found</h3><p>The score history does not yet contain a qualifying score in every required component.</p></article>`;

  const cards = [];
  const levelRules = {
    Started: ["Wintergreen; one hide in each component", "Container: 3 min · Exterior: 5 min · Interior: 5 min", "Components may be passed at different trials; no formal alert required"],
    Advanced: ["Wintergreen and Pine; two hides in each component", "Container: 3 min · Exterior: 5 min · Interior: 5 min", "Pass the corresponding Started component first; formal alert is scored"],
    Excellent: ["Wintergreen, Pine and Thyme; three hides in each component", "Container: 3 min · Exterior: 5 min · Interior: 15 min/3 rooms", "Pass the corresponding Advanced component first; call the clear Interior room"],
    Elite: ["Container, Exterior and Interior must all be passed at the Elite level in the same trial", "Call every find and clear area as required by the Elite rules"]
  };
  for (const level of LEVELS) {
    const earned = findTitle(titlePattern(level)); const counts = componentCounts(level); const missing = COMPONENTS.filter((component) => counts[component] === 0);
    let requirements = COMPONENTS.map((component) => ({ done: counts[component] > 0, text: `${component}: ${counts[component]} / 1 qualifying score${counts[component] === 1 ? "" : "s"} (have / need)` }));
    let status, note;
    if (level === "Elite") {
      const completeEliteScores = eliteSets().length;
      const excellentPrerequisite = !!findTitle(titlePattern("Excellent"))
        || !!inferredTitles().find((title) => /^Excellent\b/i.test(title.description));
      requirements = [
        { done: excellentPrerequisite, text: "Excellent title prerequisite" },
        { done: completeEliteScores > 0, text: completeEliteScores > 0 ? `${completeEliteScores} complete Elite score${completeEliteScores === 1 ? "" : "s"} (Container, Exterior and Interior all passed in the same trial)` : "One complete same-trial Elite score needed: Container, Exterior and Interior" }
      ];
      status = earned ? "Achieved" : excellentPrerequisite && completeEliteScores > 0 ? "Requirements shown as met · not yet recorded" : "In progress";
      note = earned ? `Recorded ${dateLabel(earned.date)}${earned.score == null ? "" : ` with title score ${earned.score}`}.` : "";
    } else {
      status = earned ? "Achieved" : missing.length ? "In progress" : "Requirements shown as met";
      note = earned ? `Recorded ${dateLabel(earned.date)}${earned.score == null ? "" : ` with title score ${earned.score}`}.` : missing.length ? `Needs ${missing.join(" and ")}.` : "The account shows at least one qualifying score in every component.";
    }
    requirements.push(...levelRules[level].map((text) => ({ info: true, text })));
    cards.push(progressCard(level, status, requirements, note));
  }

  const advancedSets = maximizeGoldCombinations("Advanced"), excellentSetsList = maximizeGoldCombinations("Excellent");
  const ch = championshipTitle("ch"), mach = championshipTitle("mach");
  const chSelection = ch
    ? (historicalChProgressSelection() || { sets: [], advanced: [], excellent: [], total: 0, met: false })
    : selectProgressionCombinations(advancedSets, excellentSetsList, "ch");
  const historicalMachSelection = mach ? historicalMachProgressSelection() : null;
  const machSelectionVerified = !mach || !!historicalMachSelection;
  const machSelection = historicalMachSelection || selectProgressionCombinations(advancedSets, excellentSetsList, "mach");
  const chTotal = chSelection.total, chExcellentCount = chSelection.excellent.length;
  const totalSets = machSelection.total, excellentSets = machSelection.excellent.length;
  const machAdvancedCount = machSelection.advanced.length;
  const golds = machSelection.sets.filter((set) => set.gold);
  cards.push(progressCard("SD-CH", ch ? "Achieved" : chSelection.met ? "Requirements shown as met" : "In progress", [
    { done: chTotal >= 5, text: `${chTotal}/5 selected Advanced or Excellent titling scores` },
    { done: chExcellentCount >= 2, text: `${chExcellentCount}/2 must be Excellent` }
  ], ch ? `Recorded ${dateLabel(ch.date)}.` : "Advanced components are combined only with Advanced; Excellent components only with Excellent. Initial SP titles remain locked."));
  const machRequirements = mach
    ? [
        { done: true, text: `Recorded by SDDA on ${dateLabel(mach.date)}.` },
        { done: machSelectionVerified && totalSets >= 10, text: `${machSelectionVerified ? totalSets : "—"}/10 selected Advanced or Excellent titling scores at the award trial` },
        { done: machSelectionVerified && excellentSets >= 3, text: `${machSelectionVerified ? excellentSets : "—"}/3 must be Excellent` },
        { done: machSelectionVerified && golds.length >= 3, text: `${machSelectionVerified ? golds.length : "—"}/3 Gold combinations` },
        { done: machSelectionVerified && totalSets >= 10, text: machSelectionVerified ? `${Math.min(10, totalSets)}/10 composition: ${excellentSets} Excellent + ${machAdvancedCount} Advanced` : "Historical combination counts could not be verified from the award date and trial information." }
      ]
    : [
        { done: totalSets >= 10, text: `${totalSets}/10 selected Advanced or Excellent titling scores` },
        { done: excellentSets >= 3, text: `${excellentSets}/3 must be Excellent` },
        { done: golds.length >= 3, text: `${golds.length}/3 Gold combinations (Silver combinations count toward the 10 total)` },
        { done: totalSets >= 10, text: `${Math.min(10, totalSets)}/10 composition: ${excellentSets} Excellent + ${machAdvancedCount} Advanced` }
      ];
  cards.push(progressCard("SD-MACH", mach ? "Achieved" : machSelection.met ? "Requirements shown as met" : "In progress", machRequirements,
    mach ? (machSelectionVerified ? `Historical progress reconstructed through the recorded award trial.` : "Recorded by SDDA; historical score counts could not be reconstructed.") : "The ten selected combinations preserve the Excellent and Gold requirements while favoring high totals."));

  const allElSets = eliteSets(), eliteTitleEarned = !!findTitle(titlePattern("Elite")) || allElSets.length > 0;
  const elSets = eliteTitleEarned ? allElSets.slice(1) : [];
  const elGolds = elSets.filter((set) => set.gold), elch = championshipTitle("elch");
  const elchScoresMet = eliteTitleEarned && elSets.length >= 5 && elGolds.length >= 2;
  cards.push(progressCard("SD-ELCH", elch ? "Achieved" : !eliteTitleEarned ? "Elite title required" : elchScoresMet ? "Score requirements met · title not recorded in Titles file" : "In progress", [
    { done: eliteTitleEarned, text: eliteTitleEarned ? "Elite title achieved" : "Elite title required before SD-ELCH progression begins" },
    { done: elSets.length >= 5, text: `${elSets.length}/5 additional complete same-trial Elite scores` },
    { done: elGolds.length >= 2, text: `${elGolds.length}/2 additional calculated Elite Gold scores` },
  ], elch ? `Recorded ${dateLabel(elch.date)}.` : eliteTitleEarned ? "Only Elite scores after the score that earned the Elite title count toward SD-ELCH." : "SD-ELCH progression begins after the Elite title is achieved."));
  const machScoresMet = totalSets >= 10 && excellentSets >= 3 && golds.length >= 3;
  const gmch = championshipTitle("gmch");
  const gmchScoresMet = (!!mach || machScoresMet) && (!!elch || elchScoresMet);
  cards.push(progressCard("SD-GMCH", gmch ? "Achieved" : gmchScoresMet ? "Score requirements met · title not recorded in Titles file" : "In progress", [
    { done: !!mach || machScoresMet, text: mach ? "SD-MACH recorded" : machScoresMet ? "Calculated SD-MACH score requirements met" : "SD-MACH required" }, { done: !!elch || elchScoresMet, pending: !elch && elchScoresMet, text: elch ? "SD-ELCH recorded by SDDA" : elchScoresMet ? "SD-ELCH score requirements met · title not recorded in Titles file" : "SD-ELCH required" },
  ], gmch ? `Recorded ${dateLabel(gmch.date)}.` : "The SDDA Titles file confirms whether this championship was awarded."));

  const extensions = [
    { key: "mach2", previous: "mach", prior: mach, title: "SD-MACH2", excellent: 7, gold: 3, maxAdvanced: 3 },
    { key: "mach3", previous: "mach2", prior: championshipTitle("mach2"), title: "SD-MACH3", excellent: 10, gold: 3 },
    { key: "mach4", previous: "mach3", prior: championshipTitle("mach3"), title: "SD-MACH4", excellent: 10, gold: 4 }
  ];
  const legacyMachStageMet = new Map([["mach", !!mach || machScoresMet]]);
  for (const extension of extensions) {
    const achieved = championshipTitle(extension.key), previousMet = !!extension.prior || !!legacyMachStageMet.get(extension.previous);
    if (!previousMet && !achieved) continue;
    const progress = machExtensionProgress(extension, extension.prior || extension.previous);
    const countedAdvanced = extension.maxAdvanced ? Math.min(progress.advanced, extension.maxAdvanced, Math.max(0, 10 - progress.excellent)) : 0;
    const countedExcellent = Math.min(10, progress.excellent);
    const countedTotal = Math.min(10, countedExcellent + countedAdvanced);
    const met = previousMet && progress.verified && countedTotal >= 10 && progress.excellent >= extension.excellent && progress.gold >= extension.gold;
    const requirements = [
      { done: previousMet && progress.verified, text: progress.verified ? `${extension.previous.toUpperCase()} progression requirement met${extension.prior ? " · SDDA title recorded" : " by calculated score requirements"} · cutoff Trial ${progress.cutoffTrial}` : `${extension.previous.toUpperCase()} progression is recognized, but its completion trial cannot be established from the score-file trial sequence` },
      { done: progress.verified && progress.excellent >= extension.excellent, text: `${progress.verified ? progress.excellent : "—"}/${extension.excellent} additional Excellent combinations` },
      { done: progress.verified && progress.gold >= extension.gold, text: `${progress.verified ? progress.gold : "—"}/${extension.gold} additional Excellent Gold combinations` }
    ];
    requirements.push({ done: progress.verified && countedTotal >= 10, text: progress.verified ? `${countedTotal}/10 additional combinations counted (${countedExcellent} Excellent + ${countedAdvanced} Advanced)` : "—/10 additional combinations counted until the cutoff is established" });
    if (extension.maxAdvanced) requirements.push({ done: progress.verified, text: progress.verified ? `${countedAdvanced} of up to ${extension.maxAdvanced} Advanced combinations may be included (optional)` : "Advanced allowance cannot be checked until the previous title cutoff is known" });
    requirements.push({ done: progress.verified, text: progress.verified ? `Only Qs from trials after Trial ${progress.cutoffTrial} count` : "Previous title score sequence could not be established from the uploaded chronological score list" });
    cards.push(progressCard(extension.title, achieved ? "Achieved" : !progress.verified ? "Cannot verify cutoff" : met ? "Score requirements met" : "In progress", requirements,
      `${achieved ? `Recorded ${dateLabel(achieved.date)}. ` : ""}${progress.verified ? "Repeat MACH requirements use only Qs from trials after the trial that completed the previous title." : "The cutoff is reconstructed from the trial that completed the previous title requirements."}`));
    legacyMachStageMet.set(extension.key, !!achieved || met);
  }
  $("progress").innerHTML = cards.join("");

  const classification = (total) => total >= 192 ? "Gold" : total >= 172 ? "Silver" : total >= 142 ? "Bronze" : total >= 100 ? "Satisfactory" : "Insufficient";
  const summary = (sets, label) => {
    const qualifying = sets.filter((set) => set.total >= 100);
    const counts = { Gold: 0, Silver: 0, Bronze: 0, Satisfactory: 0 };
    qualifying.forEach((set) => { counts[classification(set.total)] += 1; });
    return `<span>${qualifying.length} qualifying ${label}${qualifying.length === 1 ? "" : "s"}</span><span class="gold">${counts.Gold} Gold</span><span>${counts.Silver} Silver</span><span>${counts.Bronze} Bronze</span><span>${counts.Satisfactory} Satisfactory</span>`;
  };
  const scoreWithTrial = (component) => component ? `${component.score} (${esc(component.trial || "-")})` : "-";
  const combinationRows = (sets) => sets.filter((set) => set.total >= 100).map((set, index) => `<tr class="${set.gold ? "gold" : classification(set.total).toLowerCase()}"><td>${set.special ? `SP · Trial ${esc(set.trial || "—")}` : set.titleLocked ? "Title" : index + 1}</td><td>${scoreWithTrial(set.container)}</td><td>${scoreWithTrial(set.exterior)}</td><td>${scoreWithTrial(set.interior)}</td><td>${set.total}</td><td>${classification(set.total)}</td></tr>`).join("") || `<tr><td colspan="6">No qualifying combinations are available.</td></tr>`;
  $("advanced-gold-summary").innerHTML = summary(advancedSets, "combination");
  $("advanced-gold-rows").innerHTML = combinationRows(advancedSets);
  $("excellent-gold-summary").innerHTML = summary(excellentSetsList, "combination");
  $("excellent-gold-rows").innerHTML = combinationRows(excellentSetsList);
  $("elite-gold-summary").innerHTML = summary(elSets, "same-trial score");
  $("elite-gold-rows").innerHTML = elSets.map((set) => `<tr class="${set.gold ? "gold" : ""}"><td>Trial ${esc(set.trial || "—")}</td><td>${scoreWithTrial(set.parts.Container)}</td><td>${scoreWithTrial(set.parts.Exterior)}</td><td>${scoreWithTrial(set.parts.Interior)}</td><td>${set.total}</td><td>${classification(set.total)}</td></tr>`).join("") || `<tr><td colspan="6">No complete Elite same-trial scores are available.</td></tr>`;
  renderGames();
  renderMobileDashboard();
  document.querySelectorAll(".level-score-block table thead tr").forEach((row) => { const cells = row.querySelectorAll("th"); if (cells.length >= 4) { cells[1].textContent = "Container"; cells[2].textContent = "Exterior"; cells[3].textContent = "Interior"; } });

  $("limits").innerHTML = `<h3>Verification limits in these downloads</h3><ul><li>The Score File supplies Event Date values used for title cutoffs and chronological calculations.</li><li>The Score File does not include judge names, so two-judge and three-judge championship requirements require handler confirmation.</li><li>This report treats rows in the Dog Scores download as qualifying scores, consistent with the supplied Gold-score instructions.</li><li>SDDA remains the authority for official title approval.</li></ul>`;
}

function renderMobileDashboard() {
  const scentTarget = $("mobile-scent-titles");
  if (!scentTarget) return;
  const inferred = inferredTitles(), shown = displayedTitles();
  const titleFor = (pattern) => shown.find((title) => pattern.test(titleText(title))) || inferred.find((title) => pattern.test(titleText(title)));
  const reqList = (items) => `<ul class="requirement-list">${(items || []).map((item, index) => `<li class="${item.pending ? "pending" : item.done ? "done" : ""}">${item.pending ? "Pending — " : item.done ? "Met — " : ""}${esc(String(item.text || "").trim() || `Requirement ${index + 1} details unavailable`)}</li>`).join("") || "<li>Requirement details unavailable.</li>"}</ul>`;
  const accordion = ({ name, subtitle, achieved, pending = false, date, body, accent = "" }) => `<details class="title-accordion ${accent}${achieved ? " is-met" : ""}"><summary><span class="summary-copy"><strong>${esc(name)}</strong><small>${esc(subtitle)}</small></span>${date ? `<span class="achievement-date">${esc(dateLabel(date))}</span>` : ""}<span class="summary-action">Details</span></summary><div class="accordion-body">${body}</div></details>`;
  const excellentTitleMet = !!titleFor(/^Excellent(?:\s|$)/i)
    || COMPONENTS.every((component) => componentCounts("Excellent")[component] > 0);
  const levelTitles = [
    { name: "Started", pattern: /^Started(?:\s|$)/i }, { name: "Advanced", pattern: /^Advanced(?:\s|$)/i },
    { name: "Excellent", pattern: /^Excellent(?:\s|$)/i }, { name: "Elite", pattern: /^Elite(?:\s|$)/i, elite: true }
  ];
  const scentCards = levelTitles.map((entry) => {
    const recorded = titles.find((title) => entry.pattern.test(titleText(title))), calculated = inferred.find((title) => entry.pattern.test(titleText(title)));
    const counts = componentCounts(entry.name), sameTrial = entry.elite ? eliteSets().length : 0;
    const met = !!recorded || (entry.elite ? excellentTitleMet && sameTrial > 0 : !!calculated || COMPONENTS.every((component) => counts[component] > 0));
    const requirements = entry.elite
      ? [
          { done: excellentTitleMet, text: "Excellent title prerequisite" },
          { done: sameTrial > 0, text: `${sameTrial} complete Elite score${sameTrial === 1 ? "" : "s"} (Container, Exterior and Interior all passed in the same trial)` }
        ]
      : COMPONENTS.map((component) => ({ done: counts[component] >= 1, text: `${component}: ${counts[component]} / 1 qualifying score${counts[component] === 1 ? "" : "s"} (have / need)` }));
    return accordion({ name: entry.name, subtitle: recorded ? "Recorded title" : met ? "Requirements met · not yet recorded" : "In progress", achieved: met, date: recorded?.date, body: reqList(requirements), accent: entry.elite ? "level-accent-elite" : "" });
  });
  const activeLedger = currentMachLedgerRows();
  const advanced = activeLedger.verified ? maximizeGoldCombinations("Advanced", activeLedger.rows) : [];
  const excellent = activeLedger.verified ? maximizeGoldCombinations("Excellent", activeLedger.rows) : [];
  const elite = eliteSets();
  const recordedCh = championshipTitle("ch");
  const chSelection = recordedCh
    ? (historicalChProgressSelection() || { sets: [], advanced: [], excellent: [], total: 0, met: false })
    : selectProgressionCombinations(maximizeGoldCombinations("Advanced"), maximizeGoldCombinations("Excellent"), "ch");
  const machSelection = selectProgressionCombinations(advanced, excellent, "mach");
  const total = machSelection.total, gold = machSelection.sets.filter((set) => set.gold).length;
  const machExcellentCount = machSelection.excellent.length;
  const machAdvancedCount = machSelection.advanced.length;
  const machScoresMet = machSelection.met;
  const elchScoresMet = elite.length >= 5 && elite.filter((set) => set.gold).length >= 2;
  const championCards = [
    { name: "SD-CH", found: championshipTitle("ch"), met: chSelection.met, items: [{ done: chSelection.total >= 5, text: `${chSelection.total}/5 selected Advanced or Excellent scores` }, { done: chSelection.excellent.length >= 2, text: `${chSelection.excellent.length}/2 must be Excellent` }] },
    { name: "SD-MACH", found: championshipTitle("mach"), met: machSelection.met, items: championshipTitle("mach")
      ? [{ done: true, text: `Recorded by SDDA on ${dateLabel(championshipTitle("mach").date)}. The title requirements are satisfied.` }]
      : [{ done: total >= 10, text: `${total}/10 selected Advanced or Excellent scores` }, { done: machExcellentCount >= 3, text: `${machExcellentCount}/3 must be Excellent` }, { done: gold >= 3, text: `${gold}/3 Gold combinations (Silver combinations count toward the 10 total)` }, { done: total >= 10, text: `${Math.min(10,total)}/10 composition: ${machExcellentCount} Excellent + ${machAdvancedCount} Advanced` }] },
    { name: "SD-ELCH", found: championshipTitle("elch"), met: elite.length >= 5 && elite.filter((set) => set.gold).length >= 2, items: [{ done: elite.length >= 5, text: `${elite.length}/5 complete Elite scores` }, { done: elite.filter((set) => set.gold).length >= 2, text: `${elite.filter((set) => set.gold).length}/2 Elite Gold scores` } ] },
    { name: "SD-GMCH", found: championshipTitle("gmch"), met: (!!championshipTitle("mach") || machScoresMet) && (!!championshipTitle("elch") || elchScoresMet), items: [{ done: !!championshipTitle("mach") || machScoresMet, text: championshipTitle("mach") ? "SD-MACH recorded" : machScoresMet ? "SD-MACH score requirements met" : "SD-MACH required" }, { done: !!championshipTitle("elch") || elchScoresMet, pending: !championshipTitle("elch") && elchScoresMet, text: championshipTitle("elch") ? "SD-ELCH recorded by SDDA" : elchScoresMet ? "SD-ELCH score requirements met · title not recorded in Titles file" : "SD-ELCH required" } ] }
  ];
  const repeatMachRules = [
    { key: "mach2", previous: "mach", name: "SD-MACH2", excellent: 7, maxAdvanced: 3, gold: 3 },
    { key: "mach3", previous: "mach2", name: "SD-MACH3", excellent: 10, gold: 3 },
    { key: "mach4", previous: "mach3", name: "SD-MACH4", excellent: 10, gold: 4 }
  ];
  const machStageMet = new Map([["mach", !!championshipTitle("mach") || machScoresMet]]);
  repeatMachRules.forEach((rule) => {
    const found = championshipTitle(rule.key), previous = championshipTitle(rule.previous), progress = machExtensionProgress(rule, rule.previous);
    const previousMet = progress.verified || !!previous || !!machStageMet.get(rule.previous);
    const countedAdvanced = rule.maxAdvanced ? Math.min(progress.advanced, rule.maxAdvanced, Math.max(0, 10 - progress.excellent)) : 0;
    const countedExcellent = Math.min(10, progress.excellent);
    const countedTotal = Math.min(10, countedExcellent + countedAdvanced);
    const met = previousMet && progress.verified && countedTotal >= 10 && progress.excellent >= rule.excellent && progress.gold >= rule.gold;
    const items = [
      { done: progress.verified, text: progress.verified ? `${rule.previous.toUpperCase()} requirements reached · chronological Trial ${progress.cutoffTrial}` : `${rule.previous.toUpperCase()} score requirements could not be located in the chronological export` },
      { done: progress.verified && progress.excellent >= rule.excellent, text: `${progress.verified ? progress.excellent : "—"}/${rule.excellent} additional Excellent combinations` },
      { done: progress.verified && progress.gold >= rule.gold, text: `${progress.verified ? progress.gold : "—"}/${rule.gold} additional Excellent Gold combinations` }
    ];
    if (rule.maxAdvanced) {
      items.push({ done: progress.verified && countedTotal >= 10, text: `${progress.verified ? countedTotal : "—"}/10 additional combinations counted (${progress.verified ? countedExcellent : 0} Excellent + ${progress.verified ? countedAdvanced : 0} Advanced)` });
    } else {
      items.push({ done: progress.verified && countedExcellent >= 10, text: `${progress.verified ? countedExcellent : "—"}/10 additional Excellent combinations counted` });
    }
    if (rule.maxAdvanced) items.push({ done: progress.verified, text: progress.verified ? `${countedAdvanced} of up to ${rule.maxAdvanced} Advanced combinations may be included (optional)` : "Advanced score allowance cannot be checked until the previous title cutoff is known" });
    items.push({ done: progress.verified, text: progress.verified ? `Only Qs from trials after Trial ${progress.cutoffTrial} count` : "Previous title score sequence could not be established from the uploaded chronological score list" });
    championCards.push({ name: rule.name, found, met: !!found || met, items });
    machStageMet.set(rule.key, !!found || met);
  });
  const repeatEliteRules = [
    { key: "elch2", previous: "elch", name: "SD-ELCH2", gold: 3 },
    { key: "elch3", previous: "elch2", name: "SD-ELCH3", gold: 4 },
    { key: "elch4", previous: "elch3", name: "SD-ELCH4", gold: 5 }
  ];
  repeatEliteRules.forEach((rule) => {
    const found = championshipTitle(rule.key), previous = championshipTitle(rule.previous), eligible = scoreRowsAfterTitle(previous), sets = eligible ? eliteSets(eligible) : [], eliteGold = sets.filter((set) => set.gold).length;
    championCards.push({ name: rule.name, found, met: !!eligible && sets.length >= 5 && eliteGold >= rule.gold, items: [
      { done: !!previous, text: `${rule.previous.toUpperCase()} must be recorded first` },
      { done: !!eligible && sets.length >= 5, text: eligible ? `${sets.length}/5 new Elite scores` : "Previous-title date and complete score dates required" },
      { done: !!eligible && eliteGold >= rule.gold, text: `${eliteGold}/${rule.gold} new Elite Gold scores` },
      { done: true, text: "No additional judge-count requirement is listed for this repeat title" }
    ] });
  });
  championCards.forEach((entry) => { const achieved = !!entry.found || !!entry.met; const subtitle = entry.found ? "Recorded title" : entry.met ? "Score requirements met · not recorded in Titles file" : "In progress"; scentCards.push(accordion({ name: entry.name, subtitle, achieved, date: entry.found?.date, body: reqList(entry.items), accent: entry.name.includes("EL") || entry.name.includes("GMCH") ? "level-accent-elite" : entry.name.includes("MACH") ? "level-accent-gold" : "" })); });
  scentTarget.innerHTML = scentCards.join("");

  const runClass = (total) => total >= 192 ? "Gold" : total >= 172 ? "Silver" : total >= 142 ? "Bronze" : total >= 100 ? "Satisfactory" : "Insufficient";
  const runRows = (sets, isElite = false) => {
    const qualifying = sets.filter((set) => set.total >= 100);
    if (!qualifying.length) return `<p class="run-empty">No qualifying combinations yet.</p>`;
    return `<div class="run-grid">${qualifying.map((set, index) => { const container = isElite ? set.parts?.Container : set.container, exterior = isElite ? set.parts?.Exterior : set.exterior, interior = isElite ? set.parts?.Interior : set.interior; const show = (part) => part ? `${esc(part.score)} (${esc(part.trial || "-")})` : "—"; return `<div class="run-row ${set.gold ? "gold" : runClass(set.total).toLowerCase()}"><b>${isElite ? `Trial ${esc(set.trial || "—")}` : set.special ? "SP" : set.titleLocked ? "Title" : `#${index + 1}`}</b><span>Container: ${show(container)}</span><span>Exterior: ${show(exterior)}</span><span>Interior: ${show(interior)}</span><strong>Total: ${esc(set.total)} · ${runClass(set.total)}</strong></div>`; }).join("")}</div>`;
  };
  $("mobile-scent-runs").innerHTML = renderProgressionScoreColumns(activeLedger);
  const gameCounts = gamePassCounts();

$("mobile-game-titles").innerHTML = GAME_NAMES.map((game) => {
    const title = gameTitleRecord(game);
    const recorded = !!title;
    const passes = gameCounts[game];
    const met = recorded || passes >= 3;

    // All attempts recorded for this particular game
    const entries = games.filter(row => row.game === game);

    const fails = entries.filter(row => !row.pass).length;

    // Display individual attempts using the existing Scent ledger styling
    const history = entries.length
        ? `<div class="run-grid">
            ${entries.map(row => `
                <div class="run-row">
                    <b>Trial ${esc(row.trial || "—")}</b>
                    <span>${esc(dateLabel(row.date))}</span>
                    <strong>${row.pass ? "Q / PASS" : "FAIL"}</strong>
                </div>
            `).join("")}
           </div>`
        : `<p class="run-empty">No Games attempts recorded.</p>`;

    const body = `
        ${reqList([{
            done: met,
            text: recorded
                ? `Title awarded ${dateLabel(title.date)}`
                : `${passes}/3 qualifying passes`
        }])}

        <div class="run-summary">
            <strong>${passes} Qs / passes</strong>
            <span>${fails} failures · ${entries.length} total attempts</span>
        </div>

        ${history}
    `;

    return accordion({
        name: `${game} · SD-G${game[0]}`,
        subtitle: recorded
            ? "Recorded title"
            : `${passes}/3 passes`,
        achieved: met,
        date: title?.date,
        body,
        accent: "level-accent-silver"
    });
}).join("");
  const gameChampionships = GAME_CHAMPIONSHIP_RULES.map((rule) => {
    const found = championshipTitle(rule.key);
    const progress = gameChampionshipProgress(gameCounts, rule);
    const stageStart = ({ gch2: 0, gch3: 4, gch4: 9, gch5: 16 })[rule.key] ?? 0;
    const stageSize = ({ gch2: 4, gch3: 5, gch4: 7, gch5: 10 })[rule.key] ?? 0;
    const stageProgress = Math.min(stageSize, Math.max(0, progress.extras - stageStart));
    const stageMet = rule.type === "base" ? progress.met : stageProgress >= stageSize && progress.baseMet && (rule.type !== "four" || GAME_NAMES.every((game) => gameCounts[game] > 0));
    return { rule, found, progress, stageProgress, stageSize, stageMet };
  });
  $("mobile-game-titles").innerHTML += gameChampionships.map(({ rule, found, progress, stageProgress, stageSize, stageMet }) => {
    const base = rule.type === "base";
    const statusText = found
      ? `Recorded ${dateLabel(found.date)}`
      : base
        ? `${progress.qualifiedGames}/3 Games have reached 3 passes`
        : `${stageProgress}/${stageSize} additional passes`;
    return accordion({
      name: rule.name,
      subtitle: found ? "Recorded title" : statusText,
      achieved: !!found || stageMet,
      date: found?.date,
      body: reqList([{ done: !!found || stageMet, text: statusText }]),
      accent: "level-accent-elite"
    });
  }).join("");
    $("mobile-game-runs").innerHTML = "";

    
}

async function loadTitles(file) {
  const parsed = await parseAccountFile(file, (headers) => headers.some((header) => /^description$/i.test(header)) && headers.some((header) => /title score/i.test(header)) && headers.some((header) => /title date/i.test(header)), "Dog Titles");
  titlesFileLoaded = true;
  titles = parsed.data.map((row) => ({ description: String(row.Description || "").trim(), honours: String(row.Honours || "").trim(), score: String(row["Title Score"] || "").trim() === "" ? null : Number(row["Title Score"]), date: workbookDate(row["Title Date"]) })).filter((title) => title.description);
  $("titles-name").textContent = `${file.name} · ${titles.length} titles${parsed.sheetName === "CSV" ? "" : ` · ${parsed.sheetName}`}`; render();
  const dated = titles.filter((title) => title.date).length;
  $("message").textContent = scores.length ? `Titles loaded: ${titles.length} recorded title${titles.length === 1 ? "" : "s"}, including ${dated} title date${dated === 1 ? "" : "s"}. Progress has been recalculated.` : `Titles loaded: ${titles.length} recorded title${titles.length === 1 ? "" : "s"}, including ${dated} title date${dated === 1 ? "" : "s"}. Now upload the matching Dog Scores file.`;
}
async function loadScores(file) {
  const parsed = await parseAccountFile(file, (headers) => headers.some((header) => /^description$/i.test(header)) && headers.some((header) => /maximum score/i.test(header)) && headers.some((header) => /^score$/i.test(header)), "Dog Scores");
  const standard = [], gameRows = [];
  parsed.data.forEach((row, index) => { const description = String(row.Description || "").trim(), game = gameFrom(description); const common = { sourceRow: index + 2, dogNumber: String(row["Dog Number"] || "").trim(), dog: String(row["Registered Name"] || "").trim(), trial: String(row["Trial Number"] || "").trim(), date: workbookDate(row["Event Date"]), description }; 
if (game) {
    // Match desktop: each Games row listed in Dog Scores is a recorded pass/Q.
    gameRows.push({ ...common, game, pass: true, result: "Listed in Dog Scores" });
    return;
} const level = levelFrom(description), component = componentFrom(description), score = Number(row.Score); if (level && component && validComponentScore(level, component, row.Score)) standard.push({ ...common, level, component, score, maximum: Number(row["Maximum Score"]) }); });
  scores = standard; games = gameRows;
  $("scores-name").textContent = `${file.name} · ${scores.length} component scores${games.length ? ` · ${games.filter((row) => row.pass).length} Games passes` : ""}${parsed.sheetName === "CSV" ? "" : ` · ${parsed.sheetName}`}`; render();
}

$("titles-file").addEventListener("change", async (event) => { try { if (event.target.files[0]) await loadTitles(event.target.files[0]); } catch (error) { $("message").textContent = error.message; } });
$("scores-file").addEventListener("change", async (event) => { try { if (event.target.files[0]) await loadScores(event.target.files[0]); } catch (error) { $("message").textContent = error.message; } });

$("progression-story-button")?.addEventListener("click", () => { const panel = $("progression-story"); if (!panel) return; panel.hidden = !panel.hidden; if (!panel.hidden) renderFullProgressionExplanation(); });

document.querySelectorAll(".tracker-tab").forEach((button) => button.addEventListener("click", () => {
  document.querySelectorAll(".tracker-tab").forEach((tab) => { const active = tab === button; tab.classList.toggle("active", active); tab.setAttribute("aria-selected", String(active)); });
  document.querySelectorAll(".tracker-panel").forEach((panel) => { const active = panel.dataset.panel === button.dataset.tab; panel.hidden = !active; panel.classList.toggle("active", active); });
}));

async function downloadPdfReport() {
  if (!scores.length) return;
  const { PDFDocument, StandardFonts, rgb } = window.PDFLib;
  const pdf = await PDFDocument.create();
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const pageSize = [612, 792], margin = 46, lineHeight = 13;
  let page, y;
  function newPage() {
    page = pdf.addPage(pageSize); y = 746;
    page.drawText("SDDA Achievement Ledger", { x: margin, y: 765, size: 9, font: bold, color: rgb(.36, .28, .48) });
    page.drawText(`Page ${pdf.getPageCount()}`, { x: 520, y: 765, size: 8, font: regular, color: rgb(.4, .45, .42) });
  }
  function wrap(text, maxWidth, font, size) {
    const words = String(text || "").replace(/[^\x20-\x7E]/g, " ").replace(/\s+/g, " ").trim().split(" ");
    const lines = []; let line = "";
    for (const word of words) { const next = line ? `${line} ${word}` : word; if (font.widthOfTextAtSize(next, size) > maxWidth && line) { lines.push(line); line = word; } else line = next; }
    if (line) lines.push(line); return lines;
  }
  function text(value, options = {}) {
    const size = options.size || 10, font = options.bold ? bold : regular, color = options.color || rgb(.10, .14, .12), indent = options.indent || 0;
    const lines = wrap(value, 612 - margin * 2 - indent, font, size);
    for (const line of lines) { if (y < 52) newPage(); page.drawText(line, { x: margin + indent, y, size, font, color }); y -= options.leading || lineHeight; }
    if (options.after) y -= options.after;
  }
  function centeredText(value, options = {}) {
    const size = options.size || 10, font = options.bold ? bold : regular, color = options.color || rgb(.10, .14, .12);
    const lines = wrap(value, 612 - margin * 2, font, size);
    for (const line of lines) { if (y < 52) newPage(); const width = font.widthOfTextAtSize(line, size); page.drawText(line, { x: Math.max(margin, (612 - width) / 2), y, size, font, color }); y -= options.leading || lineHeight; }
    if (options.after) y -= options.after;
  }
  function heading(value, size = 15) { if (y < 90) newPage(); y -= 5; text(value, { size, bold: true, color: rgb(.13, .37, .27), leading: size + 4, after: 4 }); }
  function scoreTable(title, rows) {
    const headers = ["Combination / Trial", "Container", "Exterior", "Interior", "Total", "Classification"];
    const widths = [130, 70, 70, 70, 60, 120], tableWidth = widths.reduce((sum, width) => sum + width, 0), rowHeight = 19, sectionHeight = 23, headerHeight = 21;
    const drawCells = (values, height, fill, header = false) => { let x = margin; values.forEach((value, index) => { page.drawRectangle({ x, y: y - height, width: widths[index], height, color: fill, borderColor: rgb(.42, .48, .44), borderWidth: .8 }); const cellFont = header ? bold : index === 0 || index === 5 ? bold : regular, size = header ? 7.5 : 8.5, safe = String(value ?? "").replace(/[^\x20-\x7E]/g, " "), textWidth = cellFont.widthOfTextAtSize(safe, size); page.drawText(safe, { x: index > 0 ? x + Math.max(4, (widths[index] - textWidth) / 2) : x + 6, y: y - height + 6, size, font: cellFont, color: rgb(.10, .14, .12) }); x += widths[index]; }); y -= height; };
    const drawHeader = (continued = false) => { page.drawRectangle({ x: margin, y: y - sectionHeight, width: tableWidth, height: sectionHeight, color: rgb(.13, .37, .27), borderColor: rgb(.08, .24, .17), borderWidth: 1 }); page.drawText(`${title}${continued ? " (continued)" : ""}`, { x: margin + 7, y: y - 16, size: 10, font: bold, color: rgb(1, 1, 1) }); y -= sectionHeight; drawCells(headers, headerHeight, rgb(.89, .92, .90), true); };
    if (y < 52 + sectionHeight + headerHeight + rowHeight) newPage(); drawHeader();
    const tableRows = rows.length ? rows : [["No current scores", "-", "-", "-", "-", "-"]];
    for (const row of tableRows) { if (y < 52 + rowHeight) { newPage(); drawHeader(true); } drawCells(row, rowHeight, row[5] === "Gold" ? rgb(1, .96, .84) : rgb(1, 1, 1)); }
    y -= 10;
  }
  newPage();
  y -= 10;
  text(dogDisplayName(), { size: 24, bold: true, leading: 27, after: 2 });
  text(`SDDA #${scores[0]?.dogNumber || "not supplied"} | ${scores.length} qualifying component scores | ${reportAchievedTitles().length} titles achieved or score requirements met`, { size: 10, color: rgb(.35, .42, .38), after: 8 });
  text(`Generated ${new Date().toLocaleDateString()} from the member's SDDA account downloads.`, { size: 8, color: rgb(.4, .45, .42), after: 8 });

  heading("Titles achieved");
  const titleGroup = (description) => {
    const value = String(description || "").trim().toUpperCase()
      .replace(/\s+SPECIAL$/, "")
      .replace(/\s*\(SD-G[ADST]\)\s*$/, "");
    if (/^(AERIAL|DISTANCE|SPEED|TEAM)\b/.test(value) || /^SD-GS$|^SD-GCH(?:2|3|4|5)$/.test(value)) return "Games Titles and Championships";
    if (/^SD-(?:CH|MACH(?:2|3|4)?|ELCH(?:2|3|4)?|GMCH)$/.test(value)) return "Scent Championships";
    return "Scent Titles";
  };
  // Keep the source order within each category, but group every scent title
  // and championship before any Games title so the PDF never switches back.
  const titleGroupOrder = {
    "Scent Titles": 0,
    "Scent Championships": 1,
    "Games Titles and Championships": 2
  };
  const achievedTitles = reportAchievedTitles()
    .map((title, index) => ({ title, index }))
    .sort((a, b) => titleGroupOrder[titleGroup(a.title.description)] - titleGroupOrder[titleGroup(b.title.description)] || a.index - b.index)
    .map(({ title }) => title);
  let activeTitleGroup = "";
  for (const title of achievedTitles) {
    const group = titleGroup(title.description);
    if (group !== activeTitleGroup) {
      text(group, { size: 10, bold: true, color: rgb(.14, .29, .23), before: 5, after: 3 });
      activeTitleGroup = group;
    }
    const details = [dateLabel(title.date), title.score == null || Number.isNaN(title.score) ? "" : `Score ${title.score}`, title.honours || ""].filter(Boolean).join(" | ");
    text(`${title.description} — ${title.calculated ? "Calculated" : "SDDA recorded"}${details ? " | " + details : ""}`, { size: 9, bold: true, color: title.calculated ? rgb(.28, .36, .31) : rgb(.10, .14, .12), after: 2 });
  }

  heading("What is needed next");
  // Titles already listed in the achieved section do not belong in the remaining-work summary.
  const remainingProgressCards = [...document.querySelectorAll("#progress .progress-card")]
    .filter((card) => (card.querySelector(".status")?.textContent || "").trim().toLowerCase() !== "achieved");
  for (const card of remainingProgressCards) {
    const name = card.querySelector("h3")?.textContent || "Title";
    const status = card.querySelector(".status")?.textContent || "";
    const requirements = [...card.querySelectorAll("li:not(.info)")];
    text(`${name} - ${status}`, { size: 11, bold: true });
    // Keep every requirement in the PDF, including completed have / need
    // fractions, so the report shows the full component requirement at a glance.
    for (const item of card.querySelectorAll("li")) text(`${item.classList.contains("info") ? "[rule]" : item.classList.contains("done") ? "[met]" : "[needed]"} ${item.textContent}`, { indent: 12, size: 9 });
    const note = card.querySelector("p")?.textContent; if (note) text(note, { indent: 12, size: 8, color: rgb(.38, .43, .4), after: 5 });
  }

  const classification = (total) => total >= 192 ? "Gold" : total >= 172 ? "Silver" : total >= 142 ? "Bronze" : total >= 100 ? "Satisfactory" : "Insufficient";

  heading("Score classifications by level");
  text("Score columns: the number in brackets after each component score is its trial number. Advanced and Excellent components are calculated separately; Elite components are grouped by trial.", { size: 9, color: rgb(.35, .42, .38), after: 5 });
  const componentWithTrial = (component) => component ? `${component.score} (${component.trial || "-"})` : "-";
  // Keep the live post-cutoff scores at the top of each level. Historical
  // championship blocks below show the exact combinations selected at each
  // verified MACH cutoff, in chronological order.
  const activePdfLedger = currentMachLedgerRows();
  const historicalMachStages = chronologicalMachStory().stages || [];
  const pdfRowsForSets = (sets) => [...sets]
    // Keep the recorded regular-title combination visibly locked at the top
    // of its historical block; the remaining rows are the optimizer's
    // maximized combinations in their selected order.
    .sort((a, b) => Number(!!b.titleLocked) - Number(!!a.titleLocked))
    .map((set, index) => [
      set.titleLocked ? "Title locked" : set.special ? "SP" : `Combination ${set.rank || index + 1}`,
      componentWithTrial(set.container),
      componentWithTrial(set.exterior),
      componentWithTrial(set.interior),
      set.total,
      classification(set.total)
    ]);
  // The combination optimizer only returns complete three-component sets.
  // Keep every remaining current component visible as its own partial row,
  // including components that cannot yet form a qualifying combination.
  const pdfCurrentRows = (level, combinations, sourceRows) => {
    const rows = pdfRowsForSets(combinations);
    const usedSourceRows = new Set();
    combinations.forEach((set) => COMPONENTS.forEach((component) => {
      const part = set[component.toLowerCase()];
      if (part && part.sourceRow != null) usedSourceRows.add(part.sourceRow);
    }));
    const partialScores = sourceRows
      .filter((score) => score.level === level && !usedSourceRows.has(score.sourceRow))
      .sort((a, b) => (Number(a.sourceRow) || 0) - (Number(b.sourceRow) || 0));
    if (partialScores.length) {
      const cells = ["Partial", "-", "-", "-", "—", "—"];
      const componentIndexes = { Container: 1, Exterior: 2, Interior: 3 };
      partialScores.forEach((score) => {
        const componentIndex = componentIndexes[score.component];
        if (componentIndex) {
          cells[componentIndex] = cells[componentIndex] === "-"
            ? componentWithTrial(score)
            : `${cells[componentIndex]}, ${componentWithTrial(score)}`;
        }
      });
      rows.push(cells);
    }
    return rows;
  };

  for (const level of ["Advanced", "Excellent"]) {
    const currentCombinations = activePdfLedger.verified
      ? maximizeGoldCombinations(level, activePdfLedger.rows)
      : [];
    const currentRows = activePdfLedger.verified
      ? pdfCurrentRows(level, currentCombinations, activePdfLedger.rows)
      : [];
    scoreTable(`${level} — Current scores`, currentRows);

    for (const stage of historicalMachStages) {
      if (!stage.completion || !machCutoffRowForTitle(stage.key)) continue;
      const historicalSets = level === "Advanced"
        ? stage.completion.advanced
        : stage.completion.excellent;
      if (!historicalSets.length) continue;
      scoreTable(`${level} — ${stage.name} historical block`, pdfRowsForSets(historicalSets));
    }
  }
  const eliteCombinations = eliteSets();
  scoreTable("Elite - same-trial sets", eliteCombinations.map((set, index) => [`Elite ${index + 1}`, componentWithTrial(set.parts.Container), componentWithTrial(set.parts.Exterior), componentWithTrial(set.parts.Interior), set.total, classification(set.total)]));

  y -= 8;
  text("Limitations", { size: 11, bold: true, color: rgb(.61, .26, .24) });
  text("This report is a handler’s aid. SDDA is the authority on all title progression.", { size: 8, color: rgb(.35, .36, .34) });
  const bytes = await pdf.save();
  const blob = new Blob([bytes], { type: "application/pdf" });
  const safeDog = dogDisplayName().replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "");
  const filename = `${safeDog}-SDDA-title-report-${localDateStamp()}.pdf`;
  if (currentPdfUrl) URL.revokeObjectURL(currentPdfUrl);
  currentPdfUrl = URL.createObjectURL(blob);
  const ready = $("pdf-ready");
  ready.href = currentPdfUrl; ready.download = filename; ready.target = "_blank"; ready.rel = "noopener"; ready.hidden = false;
  ready.textContent = `${filename} — tap to open or download`;
  ready.click();
  return "ready";
}

$("print-button").addEventListener("click", async () => {
  const button = $("print-button"), previous = button.textContent;
  try {
    button.disabled = true; button.textContent = "Creating PDF…";
    await downloadPdfReport();
    button.textContent = "PDF created below";
  } catch (error) {
    $("message").textContent = `PDF download failed: ${error.message}`;
    button.textContent = "Try PDF download again";
  } finally {
    button.disabled = false;
    setTimeout(() => { button.textContent = previous; }, 2500);
  }
});

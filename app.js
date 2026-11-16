"use strict";

const COMPONENTS = ["Container", "Exterior", "Interior"];
const LEVELS = ["Started", "Advanced", "Excellent", "Elite"];
const GAME_NAMES = ["Aerial", "Distance", "Speed", "Team"];
let titles = [], scores = [], games = [];
let titlesFileLoaded = false;
let currentPdfUrl = "";
window.sddaJudgeVerification = window.sddaJudgeVerification || { elch: false, gmch: false };

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
const GAME_CHAMPIONSHIP_RULES=[{key:"gch",name:"SD-GCH",type:"base",target:0,text:"3 passes in each of any 3 Games"},{key:"gch2",name:"SD-GCH2",type:"three",target:4,text:"4 additional passes across any 3 Games (13 cumulative)"},{key:"gch3",name:"SD-GCH3",type:"three",target:9,text:"5 additional passes after GCH2 (18 cumulative)"},{key:"gch4",name:"SD-GCH4",type:"three",target:16,text:"7 additional passes after GCH3 (25 cumulative)"},{key:"gch5",name:"SD-GCH5",type:"four",target:26,text:"10 additional passes after GCH4 across all 4 Games (35 cumulative)"}];
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
 $("games-summary").innerHTML="<span>"+total+" banked Games passes</span><span>"+titled.length+"/4 individual Games titles</span><span>"+(progress[0].met?"SD-GCH requirements met":progress[0].remaining+" more pass(es) needed for SD-GCH")+"</span>";
 $("games-grid").innerHTML=GAME_NAMES.map(game=>{const passes=counts[game],entries=games.filter(r=>r.game===game&&r.pass!==false),recorded=gameTitleRecord(game),met=passes>=3,details=game==="Aerial"?"1 hide · 1 distractor · recommended 2 min; judge may adjust":game==="Distance"?"1 hide · no distractors · recommended 3 min; judge may adjust":game==="Speed"?"2–4 hides · 2–4 distractors · fixed 60 sec":"2–5 hides · 2–4 distractors · recommended 5 min; judge may adjust",rows=entries.map(r=>"<div class='game-attempt'><span>"+esc(dateLabel(r.date))+"</span><span>Trial "+esc(r.trial||"—")+"</span><strong>Q / PASS</strong></div>").join("");return "<article class='game-card'><h3>"+esc(game)+" <small>SD-G"+game[0]+"</small></h3><span class='status "+(met?"met":passes?"partial":"unknown")+"'>"+(met?"Individual title requirement met":passes?"In progress":"No pass data")+"</span><div class='game-count'><strong>"+passes+"</strong><span>/ 3 passes</span></div><p>"+details+"</p><div class='game-history'><h4>Games pass history</h4><div class='game-history-summary'><span>"+passes+" passes</span><span>"+entries.length+" listed entries</span></div><div class='game-attempt-header'><span>Date</span><span>Trial</span><span>Result</span></div>"+(rows||"<p>No Games pass entries recorded.</p>")+"</div>"+(recorded?"<p><strong>SDDA title record:</strong> "+esc(dateLabel(recorded.date))+"</p>":"")+"</article>";}).join("");
 $("games-championships").innerHTML="<h3>Games Championship progress · Rule Book v5.1</h3><div class='game-ladder'>"+progress.map(({rule,met,remaining,progressText})=>"<p><strong>"+rule.name+"</strong><span>"+rule.text+"</span><span class='status "+(met?"met":"partial")+"'>"+(met?"Pass requirements met":remaining+" more pass(es) needed")+"</span><small>"+progressText+"</small></p>").join("")+"</div><p class='level-note'>GCH requires 3 passes in each of any 3 Games. GCH2 adds 4 passes, GCH3 adds 5 more, GCH4 adds 7 more, and GCH5 adds 10 more across all four Games. These correspond to cumulative totals of 13, 18, 25 and 35. Surplus Games passes may be banked for repeat GCH titles, including passes earned before the previous title. Each pass counts once and cannot be counted again toward another repeat-title requirement.</p>";
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

function initialSpecialSet(level, list = scores) {
  if (level !== "Advanced" && level !== "Excellent") return null;
  // Special status depends on which trial first completed all three components.
  // Use the uploaded sheet's chronological row order and group all rows for a
  // trial together, even when that trial's rows are not adjacent.
  const groups = chronologicalTrialGroups(level, list);
  const passed = new Set();
  for (const group of groups) {
    const titledBefore = COMPONENTS.every((component) => passed.has(component));
    const componentsHere = new Set(group.rows.map((row) => row.component));
    group.rows.forEach((row) => passed.add(row.component));
    if (!titledBefore && COMPONENTS.every((component) => passed.has(component))) {
      if (!COMPONENTS.every((component) => componentsHere.has(component))) return null;
      const parts = Object.fromEntries(COMPONENTS.map((component) => [component, group.rows.filter((row) => row.component === component).sort((a, b) => b.score - a.score)[0]]));
      const total = COMPONENTS.reduce((sum, component) => sum + parts[component].score, 0);
      return { level, trial: parts.Container.trial, parts, total, gold: total >= 192, special: true };
    }
  }
  return null;
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

function maximizeGoldCombinations(level, list = scores, maximum = Infinity) {
  const originalSpecial = initialSpecialSet(level, scores);
  const includedRows = new Set(list.map((score) => score.sourceRow));
  const special = originalSpecial && COMPONENTS.every((component) => includedRows.has(originalSpecial.parts[component].sourceRow)) ? originalSpecial : null;
  const lockedRows = new Set(special ? COMPONENTS.map((component) => special.parts[component].sourceRow) : []);
  const eligible = list.filter((score) => score.level === level && !lockedRows.has(score.sourceRow));
  const ranked = Object.fromEntries(COMPONENTS.map((component) => [component, eligible.filter((score) => score.component === component).sort((a, b) => b.score - a.score || a.sourceRow - b.sourceRow)]));
  const count = Math.min(Math.max(0, maximum - (special ? 1 : 0)), ...COMPONENTS.map((component) => ranked[component].length));
  const locked = special ? [{ rank: "SP", container: special.parts.Container, interior: special.parts.Interior, exterior: special.parts.Exterior, total: special.total, gold: special.gold, special: true, level, trial: special.trial }] : [];
  if (!count) return locked;

  // Exact maximum-cardinality 3-partite matching for Gold (>=192). Each score row
  // is a distinct resource; a row can appear in at most one selected combination.
  const containers = ranked.Container.slice(0, count);
  const interiors = ranked.Interior;
  const exteriors = ranked.Exterior;
  const validPairs = containers.map((container) => {
    const pairs = [];
    for (let i = 0; i < interiors.length; i++) {
      for (let e = 0; e < exteriors.length; e++) {
        const total = container.score + interiors[i].score + exteriors[e].score;
        if (total >= 192) pairs.push({ i, e, total });
      }
    }
    pairs.sort((a, b) => a.total - b.total || a.i - b.i || a.e - b.e);
    return pairs;
  });
  const memo = new Map();
  function solve(index, usedI, usedE) {
    if (index >= containers.length) return { gold: 0, picks: [] };
    const key = index + "|" + usedI.toString(36) + "|" + usedE.toString(36);
    const cached = memo.get(key);
    if (cached) return cached;
    let best = solve(index + 1, usedI, usedE);
    best = { gold: best.gold, picks: best.picks };
    for (const pair of validPairs[index]) {
      const bitI = 1n << BigInt(pair.i), bitE = 1n << BigInt(pair.e);
      if ((usedI & bitI) !== 0n || (usedE & bitE) !== 0n) continue;
      const tail = solve(index + 1, usedI | bitI, usedE | bitE);
      if (tail.gold + 1 > best.gold) {
        best = { gold: tail.gold + 1, picks: [{ containerIndex: index, ...pair }, ...tail.picks] };
        const usedICount = usedI.toString(2).replace(/0/g, "").length;
      const usedECount = usedE.toString(2).replace(/0/g, "").length;
      const remainingUpperBound = Math.min(containers.length - index, interiors.length - usedICount, exteriors.length - usedECount);
      if (best.gold === remainingUpperBound) break;
      }
    }
    memo.set(key, best);
    return best;
  }
  const optimum = solve(0, 0n, 0n);
  const usedC = new Set(), usedI = new Set(), usedE = new Set(), goldSets = [];
  for (const pick of optimum.picks) {
    const container = containers[pick.containerIndex], interior = interiors[pick.i], exterior = exteriors[pick.e];
    usedC.add(container.sourceRow); usedI.add(interior.sourceRow); usedE.add(exterior.sourceRow);
    goldSets.push({ level, container, interior, exterior, total: pick.total, gold: true });
  }
  // Complete as many additional non-Gold titling scores as possible after the
  // maximum Gold solution; no component already used by a Gold set is reused.
  const remainingC = ranked.Container.filter((row) => !usedC.has(row.sourceRow));
  const remainingI = ranked.Interior.filter((row) => !usedI.has(row.sourceRow));
  const remainingE = ranked.Exterior.filter((row) => !usedE.has(row.sourceRow));
  const remainingCount = Math.min(count - goldSets.length, remainingC.length, remainingI.length, remainingE.length);
  const rest = Array.from({ length: remainingCount }, (_, index) => {
    const container = remainingC[index], interior = remainingI[index], exterior = remainingE[index];
    const total = container.score + interior.score + exterior.score;
    return { level, container, interior, exterior, total, gold: total >= 192 };
  });
  return [...locked, ...[...goldSets, ...rest].filter((set) => set.total >= 100).sort((a, b) => b.total - a.total).map((set, index) => ({ ...set, rank: index + 1 }))];
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

function progressCard(title, status, requirements, note = "") {
  const met = status === "Achieved" || status === "Requirements shown as met";
  const statusClass = met ? "met" : status.includes("cannot") || status.includes("verification") ? "unknown" : "partial";
  return `<article class="progress-card"><h3>${esc(title)}</h3><span class="status ${statusClass}">${esc(status)}</span><ul class="requirements">${requirements.map((requirement) => `<li class="${requirement.pending ? "pending" : requirement.done ? "done" : ""}${requirement.info ? " info" : ""}">${requirement.pending ? "Pending — " : requirement.done ? "Met — " : ""}${esc(requirement.text)}</li>`).join("")}</ul>${requirements.some((requirement) => requirement.pending) ? `<p class="verification-legend" style="color:#9b433d;font-size:.88em">Handler-reported judge minimum is awaiting SDDA confirmation.</p>` : ""}${note ? `<p>${esc(note)}</p>` : ""}</article>`;
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
 gch:/(?:^|\b)(?:SD[- ]?)?GCH(?!\s*[2-5]\b)(?:\b|$)|games champion(?!\s*[2-5]\b)|^All Games$/i,
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

function machTrialGroups() {
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
  return [...byTrial.values()];
}

function machCutoffRowForTitle(title) {
  const stage = machTitleStage(title);
  if (!stage) return null;
  const trialGroups = machTrialGroups();
  if (!trialGroups) return null;

  let cutoffIndex = -1;
  for (let currentStage = 1; currentStage <= stage; currentStage++) {
    const prefix = [];
    let foundCutoff = false;
    for (let groupIndex = cutoffIndex + 1; groupIndex < trialGroups.length; groupIndex++) {
      const group = trialGroups[groupIndex];
      prefix.push(...group.rows);
      const advanced = maximizeGoldCombinations("Advanced", prefix);
      const excellent = maximizeGoldCombinations("Excellent", prefix);
      const combined = [...advanced, ...excellent];
      const gold = combined.filter((set) => set.gold).length;
      const excellentGold = excellent.filter((set) => set.gold).length;
      let met = false;
      if (currentStage === 1) met = combined.length >= 10 && excellent.length >= 3 && gold >= 3;
      else if (currentStage === 2) met = excellent.length >= 7 && excellentGold >= 3 && excellent.length + Math.min(advanced.length, 3) >= 10;
      else if (currentStage === 3) met = excellent.length >= 10 && excellentGold >= 3;
      else if (currentStage === 4) met = excellent.length >= 10 && excellentGold >= 4;
      if (met) {
        cutoffIndex = groupIndex;
        foundCutoff = true;
        break;
      }
    }
    if (!foundCutoff) return null;
  }
  return { trial: trialGroups[cutoffIndex].trial, index: cutoffIndex };
}

function machExtensionProgress(rule, priorTitleOrStage) {
  const cutoff = machCutoffRowForTitle(priorTitleOrStage);
  const trialGroups = machTrialGroups();
  if (!cutoff || !trialGroups) {
    return { verified: false, cutoffTrial: cutoff?.trial ?? null, advanced: 0, excellent: 0, gold: 0 };
  }
  const eligibleTrials = new Set(trialGroups.slice(cutoff.index + 1).map((group) => group.trial));
  const eligible = scores.filter((score) =>
    (score.level === "Advanced" || score.level === "Excellent") &&
    eligibleTrials.has(Number(score.trial))
  );
  const advanced = maximizeGoldCombinations("Advanced", eligible);
  const excellent = maximizeGoldCombinations("Excellent", eligible);
  return {
    verified: true,
    cutoffTrial: cutoff.trial,
    advanced: advanced.length,
    excellent: excellent.length,
    gold: excellent.filter((set) => set.gold).length
  };
}
function scoreRowsAfterTitle(priorTitle) {
  const cutoffTrial = trialCutoffForTitle(priorTitle);
  const eliteScores = scores.filter((score) => score.level === "Elite");
  if (cutoffTrial == null || eliteScores.some((score) => !score.trial || !Number.isFinite(Number(score.trial)))) return null;
  return eliteScores.filter((score) => Number(score.trial) > cutoffTrial);
}

function render() {
  if (!scores.length) { $("report").hidden = true; return; }
  const dogNumbers = [...new Set(scores.map((score) => score.dogNumber).filter(Boolean))];
  if (dogNumbers.length > 1) { $("message").textContent = "The score file contains more than one dog. Please download and upload one dog's history at a time."; $("report").hidden = true; return; }
  $("report").hidden = false;
  const dog = dogDisplayName(), number = scores[0]?.dogNumber || "not supplied";
  document.title = reportFileStem();
  const summaryInferred = inferredTitles();
  const summaryTitle = (pattern) => titles.find((title) => pattern.test(titleText(title))) || summaryInferred.find((title) => pattern.test(titleText(title)));
  const summaryAdvanced = maximizeGoldCombinations("Advanced"), summaryExcellent = maximizeGoldCombinations("Excellent"), summaryElite = eliteSets();
  const summaryCombined = [...summaryAdvanced, ...summaryExcellent], summaryGolds = summaryCombined.filter((set) => set.gold).length;
  const scentAchieved = [
    !!summaryTitle(/^Started(?:\s|$)/i), !!summaryTitle(/^Advanced(?:\s|$)/i), !!summaryTitle(/^Excellent(?:\s|$)/i),
    !!titles.find((title) => /^Elite(?:\s|$)/i.test(titleText(title))) || ((!!summaryTitle(/^Excellent(?:\s|$)/i) || COMPONENTS.every((component) => componentCounts("Excellent")[component] > 0)) && summaryElite.length > 0),
    !!championshipTitle("ch") || (summaryCombined.length >= 5 && summaryExcellent.length >= 2),
    !!championshipTitle("mach") || (summaryCombined.length >= 10 && summaryExcellent.length >= 3 && summaryGolds >= 3),
    !!championshipTitle("elch") || (summaryElite.length >= 5 && summaryElite.filter((set) => set.gold).length >= 2),
    !!championshipTitle("gmch")
  ].filter(Boolean).length;
  $("dog-name").textContent = dog; $("dog-meta").textContent = `SDDA #${number} · ${scores.length} qualifying component scores · ${scentAchieved} scent titles achieved · ${8 - scentAchieved} in progress`;
  $("message").textContent = titles.length
    ? "Scores and recorded titles loaded. Results are saved on this device."
    : "Scores loaded. Achievements are calculated from score history and are not official SDDA confirmation.";

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
    let requirements = COMPONENTS.map((component) => ({ done: counts[component] > 0, text: `${component}: ${counts[component] ? `${counts[component]} qualifying score${counts[component] === 1 ? "" : "s"}` : "needed"}` }));
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
  const combined = [...advancedSets, ...excellentSetsList], golds = combined.filter((set) => set.gold);
  const ch = championshipTitle("ch"), mach = championshipTitle("mach");
  const totalSets = combined.length, excellentSets = excellentSetsList.length;
  const machExcellentCount = Math.min(excellentSets, 10);
  const machAdvancedCount = Math.min(advancedSets.length, Math.max(0, 10 - machExcellentCount));
  cards.push(progressCard("SD-CH", ch ? "Achieved" : totalSets >= 5 && excellentSets >= 2 ? "Requirements shown as met" : "In progress", [
    { done: totalSets >= 5, text: `${totalSets}/5 Advanced or Excellent titling scores` },
    { done: excellentSets >= 2, text: `${excellentSets}/2 must be Excellent` }
  ], ch ? `Recorded ${dateLabel(ch.date)}.` : "Advanced components are combined only with Advanced; Excellent components only with Excellent. Initial SP titles remain locked."));
  cards.push(progressCard("SD-MACH", mach ? "Achieved" : totalSets >= 10 && excellentSets >= 3 && golds.length >= 3 ? "Requirements shown as met" : "In progress", [
    { done: totalSets >= 10, text: `${totalSets}/10 Advanced or Excellent titling scores` },
    { done: excellentSets >= 3, text: `${excellentSets}/3 must be Excellent` },
    { done: golds.length >= 3, text: `${golds.length}/3 calculated Gold combinations (Silver combinations count toward the 10 total)` },
    { done: totalSets >= 10, text: `${Math.min(10, totalSets)}/10 composition: ${machExcellentCount} Excellent + ${machAdvancedCount} Advanced` }
  ], mach ? `Recorded ${dateLabel(mach.date)}.` : "Golds may be earned at Advanced or Excellent."));

  const elSets = eliteSets(), elGolds = elSets.filter((set) => set.gold), elch = championshipTitle("elch");
  const elchScoresMet = elSets.length >= 5 && elGolds.length >= 2;
  cards.push(progressCard("SD-ELCH", elch ? "Achieved" : elchScoresMet ? (window.sddaJudgeVerification.elch ? "Score requirements met · SDDA judge confirmation pending" : "Score requirements met · judge confirmation needed") : "In progress", [
    { done: elSets.length >= 5, text: `${elSets.length}/5 complete same-trial Elite scores` },
    { done: elGolds.length >= 2, text: `${elGolds.length}/2 calculated Elite Gold scores` },
    { done: !!window.sddaJudgeVerification.elch, pending: !!window.sddaJudgeVerification.elch, text: window.sddaJudgeVerification.elch ? "Handler confirms qualifying Elite scores were earned under at least 2 different judges" : "Handler confirmation: qualifying Elite scores under at least 2 different judges" }
  ], elch ? `Recorded ${dateLabel(elch.date)}.` : "Trial numbers are used to group Elite components."));
  const machScoresMet = totalSets >= 10 && excellentSets >= 3 && golds.length >= 3;
  const gmch = championshipTitle("gmch");
  const gmchScoresMet = (!!mach || machScoresMet) && (!!elch || elchScoresMet);
  cards.push(progressCard("SD-GMCH", gmch ? "Achieved" : gmchScoresMet ? (window.sddaJudgeVerification.gmch ? "Score requirements met · SDDA judge confirmation pending" : "Score requirements met · judge confirmation needed") : "In progress", [
    { done: !!mach || machScoresMet, text: mach ? "SD-MACH recorded" : machScoresMet ? "Calculated SD-MACH score requirements met" : "SD-MACH required" }, { done: !!elch || elchScoresMet, pending: !elch && elchScoresMet, text: elch ? "SD-ELCH recorded by SDDA" : elchScoresMet ? "SD-ELCH score requirements met · judge minimum awaiting SDDA confirmation" : "SD-ELCH required" }, { done: !!window.sddaJudgeVerification.gmch, pending: !!window.sddaJudgeVerification.gmch, text: window.sddaJudgeVerification.gmch ? "Handler confirms qualifying Scent scores were earned under at least 3 different judges" : "Handler confirmation: qualifying Scent scores under at least 3 different judges" }
  ], gmch ? `Recorded ${dateLabel(gmch.date)}.` : "The score requirements can be calculated, but judge names and official title approval require SDDA confirmation."));

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
  const combinationRows = (sets) => sets.filter((set) => set.total >= 100).map((set, index) => `<tr class="${set.gold ? "gold" : classification(set.total).toLowerCase()}"><td>${set.special ? `SP locked · Trial ${esc(set.trial || "—")}` : index + 1}</td><td>${set.container.score}</td><td>${set.exterior.score}</td><td>${set.interior.score}</td><td>${set.total}</td><td>${classification(set.total)}</td></tr>`).join("") || `<tr><td colspan="6">No qualifying combinations are available.</td></tr>`;
  $("advanced-gold-summary").innerHTML = summary(advancedSets, "combination");
  $("advanced-gold-rows").innerHTML = combinationRows(advancedSets);
  $("excellent-gold-summary").innerHTML = summary(excellentSetsList, "combination");
  $("excellent-gold-rows").innerHTML = combinationRows(excellentSetsList);
  $("elite-gold-summary").innerHTML = summary(elSets, "same-trial score");
  $("elite-gold-rows").innerHTML = elSets.map((set) => `<tr class="${set.gold ? "gold" : ""}"><td>Trial ${esc(set.trial || "—")}</td><td>${set.parts.Container.score}</td><td>${set.parts.Exterior.score}</td><td>${set.parts.Interior.score}</td><td>${set.total}</td><td>${classification(set.total)}</td></tr>`).join("") || `<tr><td colspan="6">No complete Elite same-trial scores are available.</td></tr>`;
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
  const reqList = (items) => `<ul class="requirement-list">${items.map((item) => `<li class="${item.pending ? "pending" : item.done ? "done" : ""}">${item.pending ? "Pending — " : item.done ? "Met — " : ""}${esc(item.text)}</li>`).join("")}</ul>`;
  const accordion = ({ name, subtitle, achieved, pending = false, date, body, accent = "" }) => `<details class="title-accordion ${accent}"><summary><span class="summary-copy"><strong>${esc(name)}</strong><small>${esc(subtitle)}</small></span>${date ? `<span class="achievement-date">${esc(dateLabel(date))}</span>` : ""}<span class="summary-action">Details</span></summary><div class="accordion-body">${body}</div></details>`;
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
      : COMPONENTS.map((component) => ({ done: counts[component] > 0, text: `${component}: ${counts[component]} qualifying score${counts[component] === 1 ? "" : "s"}` }));
    return accordion({ name: entry.name, subtitle: recorded ? "Recorded title" : met ? "Requirements met · not yet recorded" : "In progress", achieved: met, date: recorded?.date, body: reqList(requirements), accent: entry.elite ? "level-accent-elite" : "" });
  });
  const advanced = maximizeGoldCombinations("Advanced"), excellent = maximizeGoldCombinations("Excellent"), elite = eliteSets();
  const total = advanced.length + excellent.length, gold = [...advanced, ...excellent].filter((set) => set.gold).length;
  // Keep championship prerequisite calculations local to this renderer.
  const machExcellentCount = Math.min(excellent.length, 10);
  const machAdvancedCount = Math.min(advanced.length, Math.max(0, 10 - machExcellentCount));
  const machScoresMet = total >= 10 && excellent.length >= 3 && gold >= 3;
  const elchScoresMet = elite.length >= 5 && elite.filter((set) => set.gold).length >= 2;
  const championCards = [
    { name: "SD-CH", found: championshipTitle("ch"), met: total >= 5 && excellent.length >= 2, items: [{ done: total >= 5, text: `${total}/5 Advanced or Excellent scores` }, { done: excellent.length >= 2, text: `${excellent.length}/2 must be Excellent` }] },
    { name: "SD-MACH", found: championshipTitle("mach"), met: total >= 10 && excellent.length >= 3 && gold >= 3, items: [{ done: total >= 10, text: `${total}/10 Advanced or Excellent scores` }, { done: excellent.length >= 3, text: `${excellent.length}/3 must be Excellent` }, { done: gold >= 3, text: `${gold}/3 Gold combinations (Silver combinations count toward the 10 total)` }, { done: total >= 10, text: `${Math.min(10,total)}/10 composition: ${machExcellentCount} Excellent + ${machAdvancedCount} Advanced` }] },
    { name: "SD-ELCH", found: championshipTitle("elch"), met: elite.length >= 5 && elite.filter((set) => set.gold).length >= 2, verificationRequired: true, items: [{ done: elite.length >= 5, text: `${elite.length}/5 complete Elite scores` }, { done: elite.filter((set) => set.gold).length >= 2, text: `${elite.filter((set) => set.gold).length}/2 Elite Gold scores` }, { done: !!window.sddaJudgeVerification.elch, pending: !!window.sddaJudgeVerification.elch, text: window.sddaJudgeVerification.elch ? "Handler confirms at least 2 different judges for qualifying Elite scores" : "Handler confirmation needed: at least 2 different judges for qualifying Elite scores" }] },
    { name: "SD-GMCH", found: championshipTitle("gmch"), met: (!!championshipTitle("mach") || machScoresMet) && (!!championshipTitle("elch") || elchScoresMet), verificationRequired: true, items: [{ done: !!championshipTitle("mach") || machScoresMet, text: championshipTitle("mach") ? "SD-MACH recorded" : machScoresMet ? "SD-MACH score requirements met" : "SD-MACH required" }, { done: !!championshipTitle("elch") || elchScoresMet, pending: !championshipTitle("elch") && elchScoresMet, text: championshipTitle("elch") ? "SD-ELCH recorded by SDDA" : elchScoresMet ? "SD-ELCH score requirements met · judge minimum awaiting SDDA confirmation" : "SD-ELCH required" }, { done: !!window.sddaJudgeVerification.gmch, pending: !!window.sddaJudgeVerification.gmch, text: window.sddaJudgeVerification.gmch ? "Handler confirms at least 3 different judges for qualifying Scent scores" : "Handler confirmation needed: at least 3 different judges for qualifying Scent scores" }] }
  ];
  const repeatMachRules = [
    { key: "mach2", previous: "mach", name: "SD-MACH2", excellent: 7, maxAdvanced: 3, gold: 3 },
    { key: "mach3", previous: "mach2", name: "SD-MACH3", excellent: 10, gold: 3 },
    { key: "mach4", previous: "mach3", name: "SD-MACH4", excellent: 10, gold: 4 }
  ];
  const machStageMet = new Map([["mach", !!championshipTitle("mach") || (total >= 10 && excellent.length >= 3 && gold >= 3)]]);
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
    items.push({ done: progress.verified && countedTotal >= 10, text: `${progress.verified ? countedTotal : "—"}/10 additional combinations counted (${progress.verified ? countedExcellent : 0} Excellent + ${progress.verified ? countedAdvanced : 0} Advanced)` });
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
  championCards.forEach((entry) => { const pending = !entry.found && !!entry.met && !!entry.verificationRequired; const achieved = !!entry.found || !!entry.met; const subtitle = entry.found ? "Recorded title" : pending ? "Title requirements met · awaiting judge verification" : entry.met ? "Title requirements met · not yet recorded" : "In progress"; scentCards.push(accordion({ name: entry.name, subtitle, achieved, pending, date: entry.found?.date, body: reqList(entry.items), accent: entry.name.includes("EL") || entry.name.includes("GMCH") ? "level-accent-elite" : entry.name.includes("MACH") ? "level-accent-gold" : "" })); });
  scentTarget.innerHTML = scentCards.join("");

  const runClass = (total) => total >= 192 ? "Gold" : total >= 172 ? "Silver" : total >= 142 ? "Bronze" : total >= 100 ? "Satisfactory" : "Insufficient";
  const runRows = (sets, isElite = false) => {
    const qualifying = sets.filter((set) => set.total >= 100);
    if (!qualifying.length) return `<p class="run-empty">No qualifying combinations yet.</p>`;
    return `<div class="run-grid">${qualifying.map((set, index) => { const c = isElite ? set.parts.Container.score : set.container.score, e = isElite ? set.parts.Exterior.score : set.exterior.score, i = isElite ? set.parts.Interior.score : set.interior.score; return `<div class="run-row ${set.gold ? "gold" : runClass(set.total).toLowerCase()}"><b>${isElite ? `Trial ${esc(set.trial || "—")}` : set.special ? `SP` : `#${index + 1}`}</b><span>C ${c}</span><span>E ${e}</span><span>I ${i}</span><strong>${set.total} · ${runClass(set.total)}</strong></div>`; }).join("")}</div>`;
  };
  $("mobile-scent-runs").innerHTML = `<p class="sp-definition"><strong>SP = Special.</strong> Its three title-earning components stay together as one locked combination.</p>` + [
    accordion({ name: "Advanced combinations", subtitle: `${advanced.filter((set) => set.total >= 100).length} qualifying · ${advanced.filter((set) => set.gold && set.total >= 100).length} Gold · ${advanced.filter((set) => set.total >= 172 && set.total < 192).length} Silver · ${advanced.filter((set) => set.total >= 142 && set.total < 172).length} Bronze · ${advanced.filter((set) => set.total >= 100 && set.total < 142).length} Satisfactory`, body: `<p class="run-legend"><strong>SP = Special.</strong> Its three title-earning Advanced components stay together as one locked combination.</p>${runRows(advanced)}`, accent: "level-accent-silver" }),
    accordion({ name: "Excellent combinations", subtitle: `${excellent.filter((set) => set.total >= 100).length} qualifying · ${excellent.filter((set) => set.gold && set.total >= 100).length} Gold · ${excellent.filter((set) => set.total >= 172 && set.total < 192).length} Silver · ${excellent.filter((set) => set.total >= 142 && set.total < 172).length} Bronze · ${excellent.filter((set) => set.total >= 100 && set.total < 142).length} Satisfactory`, body: `<p class="run-legend"><strong>SP = Special.</strong> Its three title-earning Excellent components stay together as one locked combination.</p>${runRows(excellent)}`, accent: "level-accent-gold" }),
    accordion({ name: "Elite runs", subtitle: `${elite.length} same-trial scores · ${elite.filter((set) => set.gold).length} Gold`, body: runRows(elite, true), accent: "level-accent-elite" })
  ].join("");


const gameCounts = gamePassCounts(), gameTargets = [];

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

    gameTargets.push({
        name: `${game} (SD-G${game[0]})`,
        kind: "Games",
        needed: recorded ? 0 : Math.max(0, 3 - passes),
        items: [
            recorded
                ? `Recorded ${dateLabel(title.date)}`
                : `${passes}/3 passes`
        ]
    });

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
  const progressionRows = [
    { done: false, text: titlesFileLoaded ? "Titles file uploaded · official award records are shown when matched" : "Titles file not uploaded" },
    { done: !!gameChampionships[0].found || gameChampionships[0].stageMet, text: gameChampionships[0].found ? "Initial SD-GCH · Recorded by SDDA" : gameChampionships[0].progress.progressText },
    ...gameChampionships.slice(1).map(({ rule, found, stageProgress, stageSize, stageMet, progress }) => ({
      done: !!found || stageMet,
      text: found
        ? rule.name + " · Recorded by SDDA"
        : rule.name + " · " + stageProgress + "/" + stageSize + " additional passes " + (rule.type === "four" ? "across all 4 Games" : "across any 3 Games") +
          (rule.type === "four" && !GAME_NAMES.every((game) => gameCounts[game] > 0) ? " · all 4 Games must be represented" : "")
    })),
    { done: false, text: "Progression: Initial SD-GCH requires 3 passes in each of any 3 Games. GCH2 adds 4 passes, GCH3 adds 5 more, GCH4 adds 7 more, and GCH5 adds 10 more across all 4 Games. Games passes are banked and may be used toward repeat titles." }
  ];
  $("mobile-game-titles").innerHTML += accordion({
    name: "Games Championship Progression",
    subtitle: "SD-GCH through SD-GCH5",
    body: reqList(progressionRows),
    accent: "level-accent-elite"
  });
  $("mobile-game-runs").innerHTML = "";

  const targets = [];
  for (const level of ["Started", "Advanced", "Excellent"]) { if (titleFor(new RegExp(`^${level}(?:\s|$)`, "i"))) continue; const counts = componentCounts(level), missing = COMPONENTS.filter((component) => !counts[component]); targets.push({ name: level, kind: "Scent", needed: missing.length, items: missing.map((component) => `${component} qualifying score`) }); }
  if (!titleFor(/^Elite(?:\s|$)/i)) targets.push({ name: "Elite", kind: "Scent", needed: titleFor(/^Excellent(?:\s|$)/i) ? 1 : 2, items: [titleFor(/^Excellent(?:\s|$)/i) ? "One complete same-trial Elite score" : "Complete Excellent prerequisite", "One complete same-trial Elite score"] });
  const chNeeded = Math.max(Math.max(0, 5 - total), Math.max(0, 2 - excellent.length));
  const machNeeded = Math.max(Math.max(0, 10 - total), Math.max(0, 3 - excellent.length), Math.max(0, 3 - gold));
  const elchNeeded = Math.max(Math.max(0, 5 - elite.length), Math.max(0, 2 - elite.filter((set) => set.gold).length), window.sddaJudgeVerification.elch ? 0 : 1);
  if (!championshipTitle("ch") && chNeeded > 0) targets.push({ name: "SD-CH", kind: "Scent", needed: chNeeded, items: [`${Math.max(0, 5 - total)} more complete score combinations`, `${Math.max(0, 2 - excellent.length)} more Excellent combinations`] });
  if (!championshipTitle("mach") && machNeeded > 0) targets.push({ name: "SD-MACH", kind: "Scent", needed: machNeeded, items: [`${Math.max(0, 10 - total)} more complete combinations`, `${Math.max(0, 3 - excellent.length)} more Excellent combinations`, `${Math.max(0, 3 - gold)} more Gold scores`] });
  if (!championshipTitle("elch")) targets.push({ name: "SD-ELCH", kind: "Scent", verificationPending: true, needed: Math.max(1, elchNeeded), items: [`${Math.max(0, 5 - elite.length)} more Elite scores`, `${Math.max(0, 2 - elite.filter((set) => set.gold).length)} more Elite Gold scores`, { done: !!window.sddaJudgeVerification.elch, pending: !!window.sddaJudgeVerification.elch, text: window.sddaJudgeVerification.elch ? "Handler confirms at least 2 different judges for qualifying Elite scores" : "Handler confirmation needed: at least 2 different judges for qualifying Elite scores" }] });
  if (!championshipTitle("gmch")) { const machMet = !!championshipTitle("mach") || machScoresMet, elchMet = !!championshipTitle("elch") || elchScoresMet; const missing = [!machMet && "SD-MACH", !elchMet && "SD-ELCH", !window.sddaJudgeVerification.gmch && "three-judge confirmation"]; targets.push({ name: "SD-GMCH", kind: "Scent", verificationPending: true, needed: Math.max(1, missing.filter(Boolean).length), items: [...(machMet ? ["SD-MACH score requirements met"] : ["SD-MACH required"]), ...[(championshipTitle("elch") ? { done: true, text: "SD-ELCH recorded by SDDA" } : elchMet ? { done: true, pending: true, text: "SD-ELCH score requirements met · judge minimum awaiting SDDA confirmation" } : { done: false, text: "SD-ELCH required" })], { done: !!window.sddaJudgeVerification.gmch, pending: !!window.sddaJudgeVerification.gmch, text: window.sddaJudgeVerification.gmch ? "Handler confirms at least 3 different judges for qualifying Scent scores" : "Handler confirmation needed: at least 3 different judges for qualifying Scent scores" }] }); }
  repeatMachRules.forEach((rule) => {
    if (championshipTitle(rule.key) || !machStageMet.get(rule.previous)) return;
    const previous = championshipTitle(rule.previous), progress = machExtensionProgress(rule, previous);
    const countedAdvanced = rule.maxAdvanced ? Math.min(progress.advanced, rule.maxAdvanced, Math.max(0, 10 - progress.excellent)) : 0;
    const countedExcellent = Math.min(10, progress.excellent);
    const countedTotal = Math.min(10, countedExcellent + countedAdvanced);
    const items = progress.verified ? [
      previous ? `${rule.previous.toUpperCase()} recorded · cutoff Trial ${progress.cutoffTrial}` : `${rule.previous.toUpperCase()} requirements reached · inferred chronological Trial ${progress.cutoffTrial}`,
      `${progress.excellent}/${rule.excellent} additional Excellent combinations`,
      `${progress.gold}/${rule.gold} additional Excellent Gold combinations`,
      `${countedTotal}/10 additional combinations counted (${countedExcellent} Excellent + ${countedAdvanced} Advanced)`,
      ...(rule.maxAdvanced ? [`${countedAdvanced} of up to ${rule.maxAdvanced} Advanced combinations may be included (optional)`] : []),
      `Only Qs from trials after Trial ${progress.cutoffTrial} count`
    ] : ["Previous title cutoff could not be established from the chronological score rows; only later scores may count."];
    targets.push({ name: rule.name, kind: "Scent", needed: progress.verified ? Math.max(Math.max(0, 10 - Math.min(10, countedTotal)), Math.max(0,rule.excellent-progress.excellent), Math.max(0,rule.gold-progress.gold)) : 1, items });
  });
  repeatEliteRules.forEach((rule) => { if (championshipTitle(rule.key)) return; const previous = championshipTitle(rule.previous); if (!previous) return; const eligible = scoreRowsAfterTitle(previous), sets = eligible ? eliteSets(eligible) : [], eliteGold = sets.filter((set) => set.gold).length; const needed = eligible ? Math.max(Math.max(0, 5 - sets.length), Math.max(0, rule.gold - eliteGold)) : 1; targets.push({ name: rule.name, kind: "Scent", needed, items: eligible ? [`${Math.max(0, 5 - sets.length)} more new Elite scores`, `${Math.max(0, rule.gold - eliteGold)} more new Elite Gold scores`] : [`Recorded ${rule.previous.toUpperCase()} title date and complete score dates required`] }); });
  gameTargets.filter((target) => target.needed > 0).forEach((target) => targets.push(target));
  gchProgress.forEach(({ rule, found, remaining, met, progressText }) => {
    if (met || found) return;
    targets.push({ name: rule.name, kind: "Games", needed: remaining, items: [rule.text, "At least 3 passes in each of 3 different Games", `Current progress: ${progressText}`, `${remaining} more pass(es) needed from the uploaded history`] });
  });
  // Closest Titles should only list titles with actual outstanding work.
  const actionableTargets = targets.filter((target) => Number(target.needed) > 0 || target.verificationPending);
  actionableTargets.sort((a, b) => a.needed - b.needed || a.name.localeCompare(b.name));
  $("mobile-next-list").innerHTML = actionableTargets.map((target) => `<article class="next-card"><header><div><h3>${esc(target.name)}</h3><p class="next-kind">${esc(target.kind)}</p></div><span class="remaining-badge">${target.needed} needed</span></header><ol>${target.items.filter((item) => !/^0 more/i.test(item)).map((item) => typeof item === "string" ? `<li>${esc(item)}</li>` : `<li class="${item.pending ? "pending" : item.done ? "done" : ""}">${item.pending ? "Pending — " : item.done ? "Met — " : ""}${esc(item.text)}</li>`).join("")}</ol></article>`).join("") || `<article class="next-card"><h3>Everything shown is complete</h3><p>Check with SDDA for the next eligible extension.</p></article>`;
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
    const tableRows = rows.length ? rows : [["No complete combinations", "-", "-", "-", "-", "-"]];
    for (const row of tableRows) { if (y < 52 + rowHeight) { newPage(); drawHeader(true); } drawCells(row, rowHeight, row[5] === "Gold" ? rgb(1, .96, .84) : rgb(1, 1, 1)); }
    y -= 10;
  }
  newPage();
  y -= 10;
  text(dogDisplayName(), { size: 24, bold: true, leading: 27, after: 2 });
  text(`SDDA #${scores[0]?.dogNumber || "not supplied"} | ${scores.length} qualifying component scores | ${titles.length ? `${titles.length} titles recorded by SDDA` : `${inferredTitles().length} titles calculated from scores`}`, { size: 10, color: rgb(.35, .42, .38), after: 8 });
  text(`Generated ${new Date().toLocaleDateString()} from the member's SDDA account downloads.`, { size: 8, color: rgb(.4, .45, .42), after: 8 });

  heading("Titles achieved");
  for (const title of displayedTitles()) {
    text(`${title.description} - ${title.calculated ? "Calculated from scores" : "Recorded by SDDA"}`, { size: 11, bold: true });
    text(`${dateLabel(title.date)}${title.score == null || Number.isNaN(title.score) ? "" : ` | Title score ${title.score}`}${title.honours ? ` | ${title.honours}` : ""}`, { indent: 12, color: rgb(.28, .36, .31), after: 4 });
  }

  heading("What is needed next");
  for (const card of document.querySelectorAll("#progress .progress-card")) {
    const name = card.querySelector("h3")?.textContent || "Title";
    const status = card.querySelector(".status")?.textContent || "";
    const requirements = [...card.querySelectorAll("li:not(.info)")];
    const complete = status === "Achieved" || (requirements.length > 0 && requirements.every((item) => item.classList.contains("done")));
    if (complete) continue;
    text(`${name} - ${status}`, { size: 11, bold: true });
    for (const item of card.querySelectorAll("li")) text(`${item.classList.contains("info") ? "[rule]" : item.classList.contains("done") ? "[met]" : "[needed]"} ${item.textContent}`, { indent: 12, size: 9 });
    const note = card.querySelector("p")?.textContent; if (note) text(note, { indent: 12, size: 8, color: rgb(.38, .43, .4), after: 5 });
  }

  heading("Gold calculations by level");
  text("Advanced and Excellent components are calculated separately. Initial Special titles remain locked, each component is used once, and Elite components remain grouped by trial.", { size: 9, color: rgb(.35, .42, .38), after: 5 });
  for (const level of ["Advanced", "Excellent"]) {
    const combinations = maximizeGoldCombinations(level);
    scoreTable(level, combinations.map((set) => [set.special ? `SP - Trial ${set.trial || "-"}` : `Combination ${set.rank}`, set.container.score, set.exterior.score, set.interior.score, set.total, set.gold ? "Gold" : "Not Gold"]));
  }
  const eliteCombinations = eliteSets();
  scoreTable("Elite - same-trial sets", eliteCombinations.map((set) => [`Trial ${set.trial || "-"}`, set.parts.Container.score, set.parts.Exterior.score, set.parts.Interior.score, set.total, set.gold ? "Gold" : "Not Gold"]));

  y -= 8;
  text("Limitations", { size: 11, bold: true, color: rgb(.61, .26, .24) });
  text("The supplied Dog Scores download has no judge column. Judge-diversity requirements cannot be verified from this file and require handler confirmation. Event Date values are used for chronological calculations and title cutoffs. Titles calculated without the optional Dog Titles file are not official confirmation. Confirm eligibility with SDDA.", { size: 8, color: rgb(.35, .36, .34) });
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

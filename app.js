"use strict";

const COMPONENTS = ["Container", "Interior", "Exterior"];
const LEVELS = ["Started", "Advanced", "Excellent", "Elite"];
let titles = [], scores = [], trialDates = new Map(), dateMatches = 0, unmatchedTrials = [];

const $ = (id) => document.getElementById(id);
const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character]));
function localDateStamp(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
function reportFileStem() {
  const dog = (scores[0]?.dog || "SDDA-dog").trim().replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ");
  return `${dog} - SDDA Title Tracker - ${localDateStamp()}`;
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
  if (extension === "csv") {
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
function titleText(title) { return `${title.description} ${title.honours}`.trim(); }
function findTitle(pattern) { return titles.find((title) => pattern.test(titleText(title))); }
function dateLabel(value) {
  if (!value) return "Date not supplied";
  const date = new Date(`${value}T00:00:00`); return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}
function byLevel(level, list = scores) { return list.filter((score) => score.level === level); }
function componentCounts(level, list = scores) { return Object.fromEntries(COMPONENTS.map((component) => [component, list.filter((score) => score.level === level && score.component === component).length])); }

function workbookDate(value) {
  if (typeof value === "number") {
    const parsed = window.XLSX.SSF.parse_date_code(value);
    if (parsed) return `${String(parsed.y).padStart(4, "0")}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.toISOString().slice(0, 10);
  const text = String(value || "").trim();
  if (!text) return "";
  const parsed = new Date(text); return Number.isNaN(parsed.valueOf()) ? "" : parsed.toISOString().slice(0, 10);
}
function applyTrialDates() {
  dateMatches = 0;
  scores = scores.map((score) => {
    if (score.date || !score.trial) return score;
    const found = trialDates.get(String(score.trial));
    if (!found) return score;
    dateMatches++; return { ...score, date: found.date, dateSource: `Official Trial Workbook (${found.city || "trial listing"})` };
  });
  unmatchedTrials = [...new Set(scores.filter((score) => !score.date && score.trial).map((score) => String(score.trial)))].sort((a, b) => Number(a) - Number(b));
}

function initialSpecialSet(level, list = scores) {
  if (level !== "Advanced" && level !== "Excellent") return null;
  const ordered = list.filter((score) => score.level === level).sort((a, b) => {
    if (a.date && b.date && a.date !== b.date) return a.date.localeCompare(b.date);
    const aTrial = Number(a.trial), bTrial = Number(b.trial);
    if (Number.isFinite(aTrial) && Number.isFinite(bTrial) && aTrial !== bTrial) return aTrial - bTrial;
    return a.sourceRow - b.sourceRow;
  });
  const groups = [];
  for (const row of ordered) {
    const key = row.trial ? `trial:${row.trial}` : row.date ? `date:${row.date}` : `row:${row.sourceRow}`;
    const previous = groups[groups.length - 1];
    if (previous?.key === key) previous.rows.push(row); else groups.push({ key, rows: [row] });
  }
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

function maximizeGoldCombinations(list = scores, maximum = 10) {
  const specials = list === scores ? [initialSpecialSet("Advanced", list), initialSpecialSet("Excellent", list)].filter(Boolean) : [];
  const lockedRows = new Set(specials.flatMap((special) => COMPONENTS.map((component) => special.parts[component].sourceRow)));
  const eligible = list.filter((score) => (score.level === "Advanced" || score.level === "Excellent") && !lockedRows.has(score.sourceRow));
  const ranked = Object.fromEntries(COMPONENTS.map((component) => [component, eligible.filter((score) => score.component === component).sort((a, b) => b.score - a.score)]));
  const count = Math.min(Math.max(0, maximum - specials.length), ...COMPONENTS.map((component) => ranked[component].length));
  const locked = specials.map((special) => ({ rank: "SP", container: special.parts.Container, interior: special.parts.Interior, exterior: special.parts.Exterior, total: special.total, gold: special.gold, special: true, level: special.level, trial: special.trial }));
  if (!count) return locked;
  const containers = ranked.Container.slice(0, count), interiors = ranked.Interior.slice(0, count), exteriors = ranked.Exterior.slice(0, count);
  let seed = 0x5dda2026;
  const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed / 4294967296; };
  const shuffle = (values) => {
    const result = [...values];
    for (let index = result.length - 1; index > 0; index--) { const swap = Math.floor(random() * (index + 1)); [result[index], result[swap]] = [result[swap], result[index]]; }
    return result;
  };
  let best = [], bestGolds = -1;
  for (let attempt = 0; attempt < 60000; attempt++) {
    const iRows = attempt ? shuffle(interiors) : interiors, eRows = attempt ? shuffle(exteriors) : exteriors;
    const combinations = containers.map((container, index) => {
      const interior = iRows[index], exterior = eRows[index], total = container.score + interior.score + exterior.score;
      return { rank: index + 1, container, interior, exterior, total, gold: total >= 192 };
    });
    const golds = combinations.filter((combination) => combination.gold).length;
    if (golds > bestGolds) { best = combinations; bestGolds = golds; if (bestGolds === count) break; }
  }
  const arranged = best.sort((a, b) => b.total - a.total).map((combination, index) => ({ ...combination, rank: index + 1 }));
  return [...locked, ...arranged];
}

function eliteSets(list = scores) {
  const trials = new Map();
  list.filter((score) => score.level === "Elite").forEach((score) => {
    const key = score.trial || score.date || `unknown-${score.sourceRow}`;
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
    const ordered = [...rows].sort((a, b) => {
      if (a.date && b.date && a.date !== b.date) return a.date.localeCompare(b.date);
      const aTrial = Number(a.trial), bTrial = Number(b.trial);
      if (Number.isFinite(aTrial) && Number.isFinite(bTrial) && aTrial !== bTrial) return aTrial - bTrial;
      return a.sourceRow - b.sourceRow;
    });
    const groups = [];
    for (const row of ordered) {
      const key = row.trial ? `trial:${row.trial}` : row.date ? `date:${row.date}` : `row:${row.sourceRow}`;
      const previous = groups[groups.length - 1];
      if (previous?.key === key) previous.rows.push(row); else groups.push({ key, rows: [row] });
    }
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
  return `<article class="progress-card"><h3>${esc(title)}</h3><span class="status ${statusClass}">${esc(status)}</span><ul class="requirements">${requirements.map((requirement) => `<li class="${requirement.done ? "done" : ""}${requirement.info ? " info" : ""}">${esc(requirement.text)}</li>`).join("")}</ul>${note ? `<p>${esc(note)}</p>` : ""}</article>`;
}

function titlePattern(level) { return new RegExp(`^${level}(?:\\s|$)`, "i"); }
function championshipTitle(kind) {
  const patterns = {
    ch: /(^|\b)(SD[- ]?)?CH(\b|$)|scent detection champion/i,
    mach: /(^|\b)(SD[- ]?)?MACH(?!\d)|master champion(?!\s*[234])/i,
    mach2: /MACH\s*2|master champion\s*2/i, mach3: /MACH\s*3|master champion\s*3/i, mach4: /MACH\s*4|master champion\s*4/i,
    elch: /(^|\b)(SD[- ]?)?ELCH(?!\d)|elite champion(?!\s*[234])/i,
    elch2: /ELCH\s*2|elite champion\s*2/i, elch3: /ELCH\s*3|elite champion\s*3/i, elch4: /ELCH\s*4|elite champion\s*4/i,
    gmch: /GMCH|grand master champion/i
  };
  return findTitle(patterns[kind]);
}

function rowsAfter(date) {
  if (!date || scores.some((score) => !score.date)) return null;
  return scores.filter((score) => score.date > date);
}

function render() {
  if (!scores.length) { $("report").hidden = true; return; }
  const dogNumbers = [...new Set(scores.map((score) => score.dogNumber).filter(Boolean))];
  if (dogNumbers.length > 1) { $("message").textContent = "The score file contains more than one dog. Please download and upload one dog's history at a time."; $("report").hidden = true; return; }
  $("report").hidden = false;
  const dog = scores[0]?.dog || "SDDA dog", number = scores[0]?.dogNumber || "not supplied";
  document.title = reportFileStem();
  $("dog-name").textContent = dog; $("dog-meta").textContent = `SDDA #${number} · ${scores.length} qualifying component scores · ${titles.length ? `${titles.length} titles recorded by SDDA` : `${inferredTitles().length} titles calculated from scores`}`;
  $("message").textContent = titles.length
    ? `${trialDates.size ? `The Trial Workbook supplied dates for ${dateMatches} score rows. ` : ""}Scores loaded; titles are confirmed by the optional SDDA Titles file.`
    : `${trialDates.size ? `The Trial Workbook supplied dates for ${dateMatches} score rows. ` : ""}Scores loaded without a Titles file. Achievements are calculated from score history and are not official confirmation.`;

  const shownTitles = displayedTitles();
  $("achieved").innerHTML = shownTitles.map((title) => `<article class="title-card"><h3>${esc(title.description)}</h3><span class="status met">${title.calculated ? "Calculated from scores" : "Recorded by SDDA"}</span><div class="score">${title.score == null || Number.isNaN(title.score) ? "—" : esc(title.score)}</div><p>${title.score == null || Number.isNaN(title.score) ? "No title score supplied" : title.calculated ? "Calculated component total" : "Recorded title score"}</p><p><strong>${esc(dateLabel(title.date))}</strong>${title.honours ? ` · ${esc(title.honours)}` : ""}</p></article>`).join("") || `<article class="title-card"><h3>No complete title found</h3><p>The score history does not yet contain a qualifying score in every required component.</p></article>`;

  const cards = [];
  const levelRules = {
    Started: ["Wintergreen; one hide in each component", "Container: 3 min · Interior: 5 min · Exterior: 5 min", "Components may be passed at different trials; no formal alert required"],
    Advanced: ["Wintergreen and Pine; two hides in each component", "Container: 3 min · Interior: 5 min · Exterior: 5 min", "Pass the corresponding Started component first; formal alert is scored"],
    Excellent: ["Wintergreen, Pine and Thyme; three hides in each component", "Container: 3 min · Interior: 15 min/3 rooms · Exterior: 5 min", "Pass the corresponding Advanced component first; call the clear Interior room"],
    Elite: ["Excellent title required before entering Elite", "Container: 2–4 hides/3 min · Interior: 0–5 hides/10 min · Exterior/Vehicle: 0–3 hides/5 min", "All three components must qualify at the same trial; call every find and clear area", "No separate Amateur and Working streams"]
  };
  for (const level of LEVELS) {
    const earned = findTitle(titlePattern(level)); const counts = componentCounts(level); const missing = COMPONENTS.filter((component) => counts[component] === 0);
    let requirements = COMPONENTS.map((component) => ({ done: counts[component] > 0, text: `${component}: ${counts[component] ? `${counts[component]} qualifying score${counts[component] === 1 ? "" : "s"}` : "needed"}` }));
    let status, note;
    if (level === "Elite") {
      const completeEliteScores = eliteSets().length;
      requirements = [
        { done: !!findTitle(titlePattern("Excellent")) || !!inferredTitles().find((title) => /^Excellent\b/i.test(title.description)), text: "Excellent title prerequisite" },
        { done: completeEliteScores > 0, text: `${completeEliteScores} complete same-trial Elite score${completeEliteScores === 1 ? "" : "s"}` },
        ...requirements
      ];
      status = earned ? "Achieved" : completeEliteScores > 0 ? "Requirements shown as met" : "In progress";
      note = earned ? `Recorded ${dateLabel(earned.date)}${earned.score == null ? "" : ` with title score ${earned.score}`}.` : "A Container, Interior and Exterior Q must share the same trial number.";
    } else {
      status = earned ? "Achieved" : missing.length ? "In progress" : "Requirements shown as met";
      note = earned ? `Recorded ${dateLabel(earned.date)}${earned.score == null ? "" : ` with title score ${earned.score}`}.` : missing.length ? `Needs ${missing.join(" and ")}.` : "The account shows at least one qualifying score in every component.";
    }
    requirements.push(...levelRules[level].map((text) => ({ info: true, text })));
    cards.push(progressCard(level, status, requirements, note));
  }

  const excellent = highToLow("Excellent"), combined = maximizeGoldCombinations(), golds = combined.filter((set) => set.gold);
  const ch = championshipTitle("ch"), mach = championshipTitle("mach");
  const totalSets = combined.length, excellentSets = excellent.length;
  cards.push(progressCard("SD-CH", ch ? "Achieved" : totalSets >= 5 && excellentSets >= 2 ? "Requirements shown as met" : "In progress", [
    { done: totalSets >= 5, text: `${totalSets}/5 Advanced or Excellent titling scores` },
    { done: excellentSets >= 2, text: `${excellentSets}/2 must be Excellent` }
  ], ch ? `Recorded ${dateLabel(ch.date)}.` : "Calculated by mixing eligible Advanced and Excellent components while keeping an initial SP title locked."));
  cards.push(progressCard("SD-MACH", mach ? "Achieved" : totalSets >= 10 && excellentSets >= 3 && golds.length >= 3 ? "Requirements shown as met" : "In progress", [
    { done: totalSets >= 10, text: `${totalSets}/10 Advanced or Excellent titling scores` },
    { done: excellentSets >= 3, text: `${excellentSets}/3 must be Excellent` },
    { done: golds.length >= 3, text: `${golds.length}/3 calculated Gold scores` }
  ], mach ? `Recorded ${dateLabel(mach.date)}.` : "Golds may be earned at Advanced or Excellent."));

  const elSets = eliteSets(), elGolds = elSets.filter((set) => set.gold), elch = championshipTitle("elch");
  const elchScoresMet = elSets.length >= 5 && elGolds.length >= 2;
  cards.push(progressCard("SD-ELCH", elch ? "Achieved" : elchScoresMet ? "Score requirements met; judge verification needed" : "In progress", [
    { done: elSets.length >= 5, text: `${elSets.length}/5 complete same-trial Elite scores` },
    { done: elGolds.length >= 2, text: `${elGolds.length}/2 calculated Elite Gold scores` },
    { done: false, text: "Scores under at least 2 judges: cannot verify from this CSV" }
  ], elch ? `Recorded ${dateLabel(elch.date)}.` : "Trial numbers are used to group Elite components."));
  const machScoresMet = totalSets >= 10 && excellentSets >= 3 && golds.length >= 3;
  const gmch = championshipTitle("gmch");
  const gmchScoresMet = (!!mach || machScoresMet) && (!!elch || elchScoresMet);
  cards.push(progressCard("SD-GMCH", gmch ? "Achieved" : gmchScoresMet ? "Score requirements met; judge verification needed" : "In progress", [
    { done: !!mach || machScoresMet, text: mach ? "SD-MACH recorded" : machScoresMet ? "Calculated SD-MACH score requirements met" : "SD-MACH required" }, { done: !!elch || elchScoresMet, text: elch ? "SD-ELCH recorded" : elchScoresMet ? "Calculated SD-ELCH score requirements met" : "SD-ELCH required" }, { done: false, text: "At least 3 judges: cannot verify from this score download" }
  ], gmch ? `Recorded ${dateLabel(gmch.date)}.` : "The score requirements can be calculated, but judge names and official title approval require SDDA confirmation."));

  const extensions = [
    { key: "mach2", prior: mach, title: "SD-MACH2", need: 10, excellent: 7, gold: 3, goldExcellent: true },
    { key: "mach3", prior: championshipTitle("mach2"), title: "SD-MACH3", need: 10, excellent: 10, gold: 3, goldExcellent: true },
    { key: "mach4", prior: championshipTitle("mach3"), title: "SD-MACH4", need: 10, excellent: 10, gold: 4, goldExcellent: true }
  ];
  for (const extension of extensions) {
    const achieved = championshipTitle(extension.key); if (!extension.prior && !achieved) continue;
    const eligibleRows = extension.prior ? rowsAfter(extension.prior.date) : null;
    const a = eligibleRows ? highToLow("Advanced", eligibleRows) : [], e = eligibleRows ? highToLow("Excellent", eligibleRows) : [], sets = [...a, ...e], gold = e.filter((set) => set.gold).length;
    const cannot = !achieved && !eligibleRows;
    cards.push(progressCard(extension.title, achieved ? "Achieved" : cannot ? "Date verification unavailable" : "In progress", [
      { done: achieved || sets.length >= extension.need, text: achieved ? "Recorded by SDDA" : cannot ? `${extension.need} additional scores after the previous title date` : `${sets.length}/${extension.need} additional titling scores` },
      { done: achieved || e.length >= extension.excellent, text: achieved ? `Recorded ${dateLabel(achieved.date)}` : cannot ? `${extension.excellent} must be Excellent` : `${e.length}/${extension.excellent} Excellent` },
      { done: achieved || gold >= extension.gold, text: achieved ? "Gold requirement recorded as complete" : cannot ? `${extension.gold} Excellent Gold scores required` : `${gold}/${extension.gold} Excellent Gold scores` }
    ], cannot ? "Event Date is blank in one or more score rows, so scores after the prior title date cannot be isolated safely." : "Only scores after the previous championship title date are eligible."));
  }
  $("progress").innerHTML = cards.join("");

  $("gold-summary").innerHTML = `<span>${combined.length} pooled championship combinations</span><span class="gold">${golds.length} maximized Gold scores</span><span>${combined.length - golds.length} non-Gold score${combined.length - golds.length === 1 ? "" : "s"}</span>`;
  $("gold-rows").innerHTML = combined.map((set) => `<tr class="${set.gold ? "gold" : ""}"><td>${set.special ? `SP locked · Trial ${esc(set.trial || "—")}` : set.rank}</td><td>${set.container.score} <small>${set.container.level[0]}</small></td><td>${set.interior.score} <small>${set.interior.level[0]}</small></td><td>${set.exterior.score} <small>${set.exterior.level[0]}</small></td><td>${set.total}</td><td>${set.gold ? "Gold" : set.total >= 172 ? "Silver" : set.total >= 142 ? "Bronze" : "Satisfactory"}</td></tr>`).join("") || `<tr><td colspan="6">Not enough eligible components are available.</td></tr>`;

  const blanks = scores.filter((score) => !score.date).length;
  $("limits").innerHTML = `<h3>Verification limits in these downloads</h3><ul><li>The score download contains no judge column, so two-judge and three-judge championship requirements cannot be proven.</li>${dateMatches ? `<li>${dateMatches} missing Event Date value${dateMatches === 1 ? " was" : "s were"} filled by matching Trial Number in the official Trial Workbook.</li>` : ""}${blanks ? `<li>${blanks} score row${blanks === 1 ? " still has" : "s still have"} no Event Date.${unmatchedTrials.length ? ` Trial number${unmatchedTrials.length === 1 ? "" : "s"} not found in that workbook: ${unmatchedTrials.join(", ")}.` : ""} Higher championship calculations cannot reliably exclude scores earned on or before a previous title date.</li>` : ""}<li>This report treats rows in the Dog Scores download as qualifying scores, consistent with the supplied Gold-score instructions.</li><li>SDDA remains the authority for official title approval.</li></ul>`;
}

async function loadTitles(file) {
  const parsed = await parseAccountFile(file, (headers) => headers.some((header) => /^description$/i.test(header)) && headers.some((header) => /title score/i.test(header)) && headers.some((header) => /title date/i.test(header)), "Dog Titles");
  titles = parsed.data.map((row) => ({ description: String(row.Description || "").trim(), honours: String(row.Honours || "").trim(), score: String(row["Title Score"] || "").trim() === "" ? null : Number(row["Title Score"]), date: String(row["Title Date"] || "").trim() })).filter((title) => title.description);
  $("titles-name").textContent = `${file.name} · ${titles.length} titles${parsed.sheetName === "CSV" ? "" : ` · ${parsed.sheetName}`}`; render();
}
async function loadScores(file) {
  const parsed = await parseAccountFile(file, (headers) => headers.some((header) => /^description$/i.test(header)) && headers.some((header) => /maximum score/i.test(header)) && headers.some((header) => /^score$/i.test(header)), "Dog Scores");
  scores = parsed.data.map((row, index) => { const description = String(row.Description || "").trim(), level = levelFrom(description), component = componentFrom(description), score = Number(row.Score); if (!level || !component || !Number.isFinite(score)) return null; return { sourceRow: index + 2, dogNumber: String(row["Dog Number"] || "").trim(), dog: String(row["Registered Name"] || "").trim(), trial: String(row["Trial Number"] || "").trim(), date: String(row["Event Date"] || "").trim(), description, level, component, score, maximum: Number(row["Maximum Score"]) }; }).filter(Boolean);
  if (trialDates.size) applyTrialDates();
  $("scores-name").textContent = `${file.name} · ${scores.length} qualifying scores${parsed.sheetName === "CSV" ? "" : ` · ${parsed.sheetName}`}`; render();
}

async function loadTrialWorkbook(file) {
  const workbook = window.XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheet = workbook.Sheets["Processed and Future Trials"];
  if (!sheet) throw new Error("That workbook does not contain the 'Processed and Future Trials' sheet.");
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const lookup = new Map();
  for (const row of rows.slice(2)) {
    for (const start of [0, 6]) {
      const trial = String(row[start] || "").trim(), date = workbookDate(row[start + 1]), city = String(row[start + 2] || "").trim();
      if (trial && date) lookup.set(trial, { date, city });
    }
  }
  if (!lookup.size) throw new Error("No trial-number/date records were found in that workbook.");
  trialDates = lookup; applyTrialDates();
  $("workbook-name").textContent = `${file.name} · ${lookup.size} trial dates · ${dateMatches} score rows matched`;
  $("message").textContent = dateMatches ? `Matched ${dateMatches} blank score-row dates using Trial Number.` : `Loaded ${lookup.size} official trial dates, but none of this dog's blank-date trial numbers were listed.`;
  render();
}

$("titles-file").addEventListener("change", async (event) => { try { if (event.target.files[0]) await loadTitles(event.target.files[0]); } catch (error) { $("message").textContent = error.message; } });
$("scores-file").addEventListener("change", async (event) => { try { if (event.target.files[0]) await loadScores(event.target.files[0]); } catch (error) { $("message").textContent = error.message; } });
$("workbook-file").addEventListener("change", async (event) => { try { if (event.target.files[0]) await loadTrialWorkbook(event.target.files[0]); } catch (error) { $("message").textContent = `Trial Workbook import failed: ${error.message}`; } });

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
    page.drawText("SDDA Title Tracker", { x: margin, y: 765, size: 9, font: bold, color: rgb(.13, .37, .27) });
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
  function heading(value, size = 15) { if (y < 90) newPage(); y -= 5; text(value, { size, bold: true, color: rgb(.13, .37, .27), leading: size + 4, after: 4 }); }
  newPage();
  text(scores[0]?.dog || "SDDA dog", { size: 24, bold: true, leading: 27 });
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
    text(`${name} - ${status}`, { size: 11, bold: true });
    for (const item of card.querySelectorAll("li")) text(`${item.classList.contains("done") ? "[met]" : "[needed]"} ${item.textContent}`, { indent: 12, size: 9 });
    const note = card.querySelector("p")?.textContent; if (note) text(note, { indent: 12, size: 8, color: rgb(.38, .43, .4), after: 5 });
  }

  heading("Advanced and Excellent Gold calculations");
  text("Eligible components are mixed to maximize complete scores of 192 or higher. Initial Special titles remain locked and each component score is used once.", { size: 9, color: rgb(.35, .42, .38), after: 5 });
  const combinations = maximizeGoldCombinations();
  if (!combinations.length) text("No complete Advanced or Excellent combination is available.");
  for (const set of combinations) text(`${set.special ? `SP locked, trial ${set.trial}` : `Combination ${set.rank}`}: C ${set.container.score} + I ${set.interior.score} + E ${set.exterior.score} = ${set.total} (${set.gold ? "GOLD" : "not Gold"})`, { size: 9, bold: set.gold });

  y -= 8;
  text("Limitations", { size: 11, bold: true, color: rgb(.61, .26, .24) });
  text("The supplied Dog Scores download has no judge column. Judge-diversity requirements cannot be verified. Blank Event Date values prevent reliable score cutoffs for higher championship extensions. Titles calculated without the optional Dog Titles file are not official confirmation. Confirm eligibility with SDDA.", { size: 8, color: rgb(.35, .36, .34) });
  const bytes = await pdf.save();
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob), link = document.createElement("a");
  const safeDog = (scores[0]?.dog || "SDDA-dog").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "");
  link.href = url; link.download = `${safeDog}-SDDA-title-report.pdf`; document.body.appendChild(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000);
}

$("print-button").addEventListener("click", () => {
  document.title = reportFileStem();
  window.print();
});

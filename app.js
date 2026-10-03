/* Got Dirt — Cherry Creek Nursery stock location command center */
(function () {
  const $ = (s, el = document) => el.querySelector(s);
  const LS_LABELS = "gotDirt.polygonLabels";
  const LS_LABELS_CLEARED_V3 = "gotDirt.labelsClearedV3";
  const LS_LABELS_CLEARED_V4 = "gotDirt.labelsClearedV4";
  const LS_CUSTOM = "gotDirt.customOutlines";
  const LS_OUTLINES_RESET_V1 = "gotDirt.outlinesResetV1";
  const LS_DELETED_AUTO = "gotDirt.deletedAutoIndices";
  const LS_ID_NORMALIZE_V1 = "gotDirt.idNormalizeV1";
  const LS_GPS_DECISIONS = "gotDirt.gpsPinDecisions"; // legacy v22
  const LS_PHOTO_ASSIGN = "gotDirt.photoAssignments";

  const state = {
    plots: [],
    coords: {},
    polygons: [],
    polyById: new Map(),
    calibration: "unknown",
    query: "",
    hits: [],
    selected: new Set(),
    generatedAt: "",
    labelMode: false,
    labelFocusIndex: null,
    localLabels: {},
    outlineMode: false,
    hideAutoOutlines: false,
    customOutlines: [],
    drawPoints: [],
    pendingPath: null,
    selectedCustomIdx: null,
    selectedAutoIndex: null,
    deletedAutoIndices: new Set(),
    plotPhotos: {},
    photoPanelClosed: false,
    gpsConfirmMode: false, // legacy alias; use assignMode
    assignMode: false,
    gpsPins: [],
    gpsDecisions: {}, // legacy
    photoAssignments: {},
    gpsFocusIdx: null,
    assignFocusPath: null,
    gpsShowPins: true,
    assignFilter: "unassigned",
    assignQuery: "",
  };

  function clearLabelsOnceV3() {
    try {
      if (localStorage.getItem(LS_LABELS_CLEARED_V3) === "1") return;
      localStorage.removeItem(LS_LABELS);
      localStorage.setItem(LS_LABELS_CLEARED_V3, "1");
    } catch (_) {}
  }

  function clearLabelsOnceV4() {
    try {
      if (localStorage.getItem(LS_LABELS_CLEARED_V4) === "1") return;
      localStorage.removeItem(LS_LABELS);
      localStorage.setItem(LS_LABELS_CLEARED_V4, "1");
    } catch (_) {}
  }


  // One-time: format near-miss renames (board → Excel). Does not invent polygons.
  // Skips 1.35→1.35B because Excel already has plain 1.35 matched on the board.
  const ID_NORMALIZE_V1_MAP = {
    "1.1": "1.01",
    "1.2": "1.02",
    "1.3": "1.03",
    "1.4": "1.04",
    "1.5": "1.05",
    "1.6": "1.06",
    "1.7": "1.07",
    "1.8": "1.08",
    "1.9": "1.09",
    "1.26": "1.26A",
    "1.60": "1.60A",
    "1.68": "1.68B",
    "1.126": "1.126A",
    "13183": "1.183",
  };

  function normalizeIdsOnceV1() {
    try {
      if (localStorage.getItem(LS_ID_NORMALIZE_V1) === "1") return;
      const renames = [];
      let labelsChanged = false;
      let outlinesChanged = false;
      for (const [idx, id] of Object.entries(state.localLabels || {})) {
        const next = ID_NORMALIZE_V1_MAP[id];
        if (!next || next === id) continue;
        state.localLabels[idx] = next;
        labelsChanged = true;
        renames.push({ via: "auto", index: idx, from: id, to: next });
      }
      for (let i = 0; i < state.customOutlines.length; i++) {
        const o = state.customOutlines[i];
        if (!o || !o.id) continue;
        const next = ID_NORMALIZE_V1_MAP[o.id];
        if (!next || next === o.id) continue;
        const from = o.id;
        o.id = next;
        outlinesChanged = true;
        renames.push({ via: "custom", index: i, from, to: next });
      }
      if (labelsChanged) saveLocalLabels();
      if (outlinesChanged) saveCustomOutlines();
      localStorage.setItem(LS_ID_NORMALIZE_V1, "1");
      if (renames.length) {
        console.info("[got-dirt] idNormalizeV1 renames:", renames);
      } else {
        console.info("[got-dirt] idNormalizeV1: flag set, no matching old ids");
      }
    } catch (err) {
      console.warn("[got-dirt] idNormalizeV1 failed", err);
    }
  }

  function loadLocalLabels() {
    clearLabelsOnceV3();
    clearLabelsOnceV4();
    try {
      state.localLabels = JSON.parse(localStorage.getItem(LS_LABELS) || "{}") || {};
    } catch {
      state.localLabels = {};
    }
  }

  function normalizeLabelsDoc(doc) {
    if (!doc) return {};
    if (doc.labels && typeof doc.labels === "object") return doc.labels;
    return doc;
  }

  /** Pack fills gaps; local overlay wins on same index (never invent IDs). */
  function mergeLabelMaps(pack, local) {
    const out = { ...(pack || {}) };
    for (const [k, v] of Object.entries(local || {})) {
      if (v === undefined) continue;
      const key = String(k);
      // Blank local must not erase a real pack label (stale browser overlay).
      const blank = v === null || v === "";
      if (blank && out[key]) continue;
      out[key] = v;
    }
    return out;
  }

  function persistLocalLabelsOnly() {
    try {
      localStorage.setItem(LS_LABELS, JSON.stringify(state.localLabels));
    } catch (_) {}
  }

  async function loadLabelsFromPack() {
    try {
      const r = await fetch("data/polygon-labels.json?v=30");
      if (!r.ok) return;
      const doc = await r.json();
      const fromPack = normalizeLabelsDoc(doc);
      const before = Object.keys(state.localLabels || {}).length;
      state.localLabels = mergeLabelMaps(fromPack, state.localLabels);
      const after = Object.keys(state.localLabels || {}).length;
      persistLocalLabelsOnly();
      if (after > before) console.info("[got-dirt] merged polygon-labels from pack", { before, after });
    } catch (_) {}
  }

  function saveLocalLabels() {
    persistLocalLabelsOnly();
    try {
      const localMap = state.localLabels || {};
      const localCount = Object.keys(localMap).length;
      fetch("data/polygon-labels.json")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null)
        .then((doc) => {
          const packMap = normalizeLabelsDoc(doc) || {};
          const packCount = Object.keys(packMap).length;
          if (localCount === 0 && packCount > 0) {
            state.localLabels = mergeLabelMaps(packMap, localMap);
            persistLocalLabelsOnly();
            return null;
          }
          const merged = mergeLabelMaps(packMap, localMap);
          state.localLabels = merged;
          persistLocalLabelsOnly();
          const realCount = Object.values(merged).filter((v) => v && v !== "x").length;
          const body = JSON.stringify({
            version: 1,
            updatedAt: new Date().toISOString(),
            count: Object.keys(merged).length,
            realCount,
            labels: merged,
          });
          return fetch("data/polygon-labels.json", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body,
          });
        })
        .catch(() => {});
    } catch (_) {}
  }

  function clearAllLabels() {
    state.localLabels = {};
    try {
      localStorage.removeItem(LS_LABELS);
      localStorage.setItem(LS_LABELS_CLEARED_V3, "1");
    } catch (_) {}
    rebuildPolyIndex();
    renderPolygons();
    renderPins();
    renderResults();
    renderLabelQueue();
    updateMapStatus();
    updateOutlineMeta();
  }

  function clearCustomOutlinesOnceV1() {
    try {
      if (localStorage.getItem(LS_OUTLINES_RESET_V1) === "1") return;
      localStorage.removeItem(LS_CUSTOM);
      localStorage.setItem(LS_OUTLINES_RESET_V1, "1");
      console.info("[got-dirt] cleared gotDirt.customOutlines (outlinesResetV1)");
    } catch (_) {}
  }

  function loadCustomOutlines() {
    clearCustomOutlinesOnceV1();
    try {
      const raw = JSON.parse(localStorage.getItem(LS_CUSTOM) || "[]");
      state.customOutlines = Array.isArray(raw) ? raw.filter((o) => o && o.path && o.path.length >= 3) : [];
    } catch {
      state.customOutlines = [];
    }
  }

  function normalizeOutlinesDoc(doc) {
    if (!doc) return [];
    if (Array.isArray(doc.outlines)) return doc.outlines;
    if (Array.isArray(doc)) return doc;
    return [];
  }

  /** Union by id; local wins. Keep pack-only ids. Never invent paths/ids. */
  function mergeCustomOutlineLists(pack, local) {
    const byId = new Map();
    const noId = [];
    for (const o of pack || []) {
      if (!o || !o.path || o.path.length < 3) continue;
      if (o.id) byId.set(String(o.id), o);
      else noId.push(o);
    }
    for (const o of local || []) {
      if (!o || !o.path || o.path.length < 3) continue;
      if (o.id) byId.set(String(o.id), o);
      else noId.push(o);
    }
    return [...byId.values(), ...noId];
  }

  function persistCustomOutlinesOnly() {
    try {
      localStorage.setItem(LS_CUSTOM, JSON.stringify(state.customOutlines));
    } catch (_) {}
  }

  async function loadCustomOutlinesFromPack() {
    try {
      const r = await fetch("data/custom-outlines.json?v=30");
      if (!r.ok) return;
      const doc = await r.json();
      const fromPack = normalizeOutlinesDoc(doc).filter((o) => o && o.path && o.path.length >= 3);
      const before = (state.customOutlines || []).length;
      state.customOutlines = mergeCustomOutlineLists(fromPack, state.customOutlines);
      const after = state.customOutlines.length;
      persistCustomOutlinesOnly();
      if (after > before) console.info("[got-dirt] merged custom-outlines from pack", { before, after });
    } catch (_) {}
  }

  function saveCustomOutlines() {
    persistCustomOutlinesOnly();
    try {
      const localList = state.customOutlines || [];
      fetch("data/custom-outlines.json")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null)
        .then((doc) => {
          const packList = normalizeOutlinesDoc(doc);
          if (localList.length === 0 && packList.length > 0) {
            state.customOutlines = mergeCustomOutlineLists(packList, localList);
            persistCustomOutlinesOnly();
            return null;
          }
          const merged = mergeCustomOutlineLists(packList, localList);
          state.customOutlines = merged;
          persistCustomOutlinesOnly();
          const body = JSON.stringify({
            version: 1,
            updatedAt: new Date().toISOString(),
            count: merged.length,
            outlines: merged,
          });
          return fetch("data/custom-outlines.json", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body,
          });
        })
        .catch(() => {});
    } catch (_) {}
  }

  function loadDeletedAutoIndices() {
    try {
      const raw = JSON.parse(localStorage.getItem(LS_DELETED_AUTO) || "[]");
      state.deletedAutoIndices = new Set(
        Array.isArray(raw) ? raw.map((n) => Number(n)).filter((n) => Number.isFinite(n)) : []
      );
    } catch {
      state.deletedAutoIndices = new Set();
    }
  }

  function normalizeDeletedDoc(doc) {
    if (!doc) return [];
    if (Array.isArray(doc.indices)) return doc.indices;
    if (Array.isArray(doc)) return doc;
    return [];
  }

  function persistDeletedAutoOnly() {
    try {
      localStorage.setItem(LS_DELETED_AUTO, JSON.stringify([...state.deletedAutoIndices]));
    } catch (_) {}
  }

  async function loadDeletedAutoFromPack() {
    try {
      const r = await fetch("data/deleted-auto-indices.json?v=30");
      if (!r.ok) return;
      const doc = await r.json();
      const fromPack = normalizeDeletedDoc(doc)
        .map((n) => Number(n))
        .filter((n) => Number.isFinite(n));
      const before = state.deletedAutoIndices.size;
      for (const n of fromPack) state.deletedAutoIndices.add(n);
      persistDeletedAutoOnly();
      if (state.deletedAutoIndices.size > before) {
        console.info("[got-dirt] merged deleted-auto-indices from pack", {
          before,
          after: state.deletedAutoIndices.size,
        });
      }
    } catch (_) {}
  }

  function saveDeletedAutoIndices() {
    persistDeletedAutoOnly();
    try {
      const localList = [...state.deletedAutoIndices];
      fetch("data/deleted-auto-indices.json")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null)
        .then((doc) => {
          const packList = normalizeDeletedDoc(doc).map((n) => Number(n)).filter((n) => Number.isFinite(n));
          if (localList.length === 0 && packList.length > 0) {
            for (const n of packList) state.deletedAutoIndices.add(n);
            persistDeletedAutoOnly();
            return null;
          }
          const merged = [...new Set([...packList, ...localList])].sort((a, b) => a - b);
          state.deletedAutoIndices = new Set(merged);
          persistDeletedAutoOnly();
          const body = JSON.stringify({
            version: 1,
            updatedAt: new Date().toISOString(),
            count: merged.length,
            indices: merged,
          });
          return fetch("data/deleted-auto-indices.json", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body,
          });
        })
        .catch(() => {});
    } catch (_) {}
  }

  function refreshAfterOutlineChange() {
    rebuildPolyIndex();
    renderPolygons();
    renderPins();
    renderResults();
    renderLabelQueue();
    updateOutlineMeta();
    updateMapStatus();
  }

  function deleteAutoOutline(index, { confirmLabeled = false } = {}) {
    if (index == null || !Number.isFinite(Number(index))) return false;
    const idx = Number(index);
    const feat = state.polygons.find((f) => f.index === idx);
    if (!feat) return false;
    if (confirmLabeled) {
      const id = effectiveId(feat);
      if (id) {
        const ok = window.confirm(`Delete auto outline #${idx} (labeled ${id})? This hides it in your browser until you clear deleted-auto storage.`);
        if (!ok) return false;
      }
    }
    state.deletedAutoIndices.add(idx);
    saveDeletedAutoIndices();
    state.polygons = state.polygons.filter((f) => f.index !== idx);
    if (state.selectedAutoIndex === idx) state.selectedAutoIndex = null;
    if (state.labelFocusIndex === idx) state.labelFocusIndex = null;
    // Drop local label for deleted index (optional cleanup)
    if (Object.prototype.hasOwnProperty.call(state.localLabels, String(idx))) {
      delete state.localLabels[String(idx)];
      saveLocalLabels();
    }
    refreshAfterOutlineChange();
    return true;
  }

  function clearAllAutoOutlines() {
    const n = state.polygons.length;
    if (!n) {
      updateOutlineMeta();
      return;
    }
    const ok = window.confirm(`Delete ALL ${n} auto outlines? Custom outlines stay. This is saved in your browser (plot-polygons.json is unchanged).`);
    if (!ok) return;
    for (const f of state.polygons) state.deletedAutoIndices.add(f.index);
    saveDeletedAutoIndices();
    state.polygons = [];
    state.selectedAutoIndex = null;
    state.labelFocusIndex = null;
    refreshAfterOutlineChange();
  }

  function deleteSelectedOutline() {
    if (state.selectedCustomIdx != null && state.customOutlines[state.selectedCustomIdx]) {
      deleteSelectedCustom();
      return;
    }
    if (state.selectedAutoIndex != null) {
      deleteAutoOutline(state.selectedAutoIndex);
    }
  }

  /** Effective plot id for a polygon feature (localStorage overrides file). */
  function effectiveId(feat) {
    if (!feat) return null;
    let raw = null;
    if (feat.source === "aj") raw = feat.id || null;
    else {
      const key = String(feat.index);
      if (Object.prototype.hasOwnProperty.call(state.localLabels, key)) {
        const loc = state.localLabels[key];
        raw = loc ? loc : null;
      } else raw = feat.id || null;
    }
    return canonicalPlotId(raw);
  }

  /** Inventory id for a label. Near-miss board ids (1.4 -> 1.04) resolve on every lookup, not once. "x" is not an id. */
  function canonicalPlotId(raw) {
    if (raw == null) return null;
    const t = String(raw).trim();
    if (!t || t.toLowerCase() === "x") return null;
    const mapped = Object.prototype.hasOwnProperty.call(ID_NORMALIZE_V1_MAP, t) ? ID_NORMALIZE_V1_MAP[t] : t;
    if (state.plots && state.plots.length) {
      const hit = state.plots.find((p) => p.id.toLowerCase() === mapped.toLowerCase());
      if (hit) return hit.id;
      if (mapped !== t) {
        const hitRaw = state.plots.find((p) => p.id.toLowerCase() === t.toLowerCase());
        if (hitRaw) return hitRaw.id;
      }
    }
    return mapped;
  }

  // Inventory B rows share one existing labeled block. Lookup only.
  // Do not draw a new outline and do not invent a plot id.
  // 1.23 is intentionally absent (stays unlabeled/pin until Aj says otherwise).
  // Future empty Farm 2/3 ids are not aliases.
  const SAME_BLOCK_CANDIDATES = {
    "1.26B": ["1.26A", "1.26"],
    "1.35B": ["1.35", "1.35A"],
    "1.60B": ["1.60A", "1.60"],
    "1.126B": ["1.126A", "1.126"],
  };

  function sameBlockCandidates(id) {
    if (!id) return null;
    const key = String(id);
    return SAME_BLOCK_CANDIDATES[key] || SAME_BLOCK_CANDIDATES[key.toUpperCase()] || null;
  }

  /** Outline id a B inventory row should highlight, or null if that sibling is not labeled. */
  function siblingOutlineKey(id) {
    const cands = sameBlockCandidates(id);
    if (!cands || !state.polyById) return null;
    for (const c of cands) {
      const sk = canonicalPlotId(c) || c;
      if (state.polyById.has(sk)) return sk;
      if (state.polyById.has(c)) return c;
    }
    return null;
  }

  function outlineFor(id) {
    if (!id) return null;
    const key = canonicalPlotId(id) || String(id);
    const direct = state.polyById.get(key) || state.polyById.get(String(id));
    if (direct) return direct;
    const sib = siblingOutlineKey(key) || siblingOutlineKey(id);
    return sib ? state.polyById.get(sib) : null;
  }

  function outlineSelected(outlineId) {
    if (!outlineId || !state.selected || !state.selected.size) return false;
    if (state.selected.has(outlineId)) return true;
    for (const sel of state.selected) {
      const sib = siblingOutlineKey(sel);
      if (sib && sib === outlineId) return true;
    }
    return false;
  }

  function rebuildPolyIndex() {
    state.polyById = new Map();
    // Custom outlines win for search/highlight
    for (const o of state.customOutlines) {
      if (!o || !o.id) continue;
      const id = canonicalPlotId(o.id);
      if (!id || state.polyById.has(id)) continue;
      state.polyById.set(id, {
        id,
        path: o.path,
        source: "aj",
        centroid: centroidFromPath(o.path),
      });
    }
    // Always index autos by effectiveId (localStorage labels + file ids).
    // hideAuto only affects idle paint — search must still resolve labeled boxes.
    for (const f of state.polygons) {
      const id = effectiveId(f);
      if (id && !state.polyById.has(id)) state.polyById.set(id, f);
    }
  }

  function centroidFromPath(path) {
    if (!path || !path.length) return [50, 50];
    let sx = 0, sy = 0;
    for (const pt of path) {
      sx += pt[0];
      sy += pt[1];
    }
    return [sx / path.length, sy / path.length];
  }

  function norm(s) {
    return String(s || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function parseQuerySizes(q) {
    const sizes = [];
    const re = /(\d+(?:\.\d+)?)\s*(?:\"|''|”|″|in(?:ches?)?(?![a-z]))(?![0-9.])/gi;
    let m;
    while ((m = re.exec(q))) {
      const n = parseFloat(m[1]);
      if (!Number.isNaN(n) && n > 0 && n < 100) sizes.push(n);
    }
    return sizes;
  }

  function stripSizePhrases(q) {
    return q
      .replace(/(\d+(?:\.\d+)?)\s*(?:\"|''|”|″|in(?:ches?)?(?![a-z]))(?![0-9.])/gi, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function extractInchRanges(text) {
    const t = String(text || "");
    const ranges = [];
    const rangeRe = /(\d+(?:\.\d+)?)\s*[-–—]\s*(\d+(?:\.\d+)?)\s*(?:\"|''|”|″|in(?:ches?)?(?![a-z]))(?![0-9.])/gi;
    let m;
    while ((m = rangeRe.exec(t))) {
      const a = parseFloat(m[1]);
      const b = parseFloat(m[2]);
      if (Number.isNaN(a) || Number.isNaN(b)) continue;
      ranges.push({ lo: Math.min(a, b), hi: Math.max(a, b), raw: m[0], index: m.index });
    }
    const singleRe = /(?:\(\d+\))?\s*(\d+(?:\.\d+)?)\s*(?:\"|''|”|″|in(?:ches?)?(?![a-z]))(?![0-9.])/gi;
    while ((m = singleRe.exec(t))) {
      const start = m.index;
      const overlapsRange = ranges.some((r) => start >= r.index && start < r.index + r.raw.length);
      if (overlapsRange) continue;
      const n = parseFloat(m[1]);
      if (Number.isNaN(n) || n <= 0 || n >= 100) continue;
      ranges.push({ lo: n, hi: n, raw: m[0], index: start });
    }
    return ranges;
  }

  function sizeCoveredByPlants(plants, size) {
    const ranges = extractInchRanges(plants);
    return ranges.some((r) => size >= r.lo - 1e-6 && size <= r.hi + 1e-6);
  }

  function sizeNearCultivar(plants, textToks, size) {
    if (!textToks.length) return sizeCoveredByPlants(plants, size);
    const lower = String(plants || "").toLowerCase();
    const clauses = lower.split(/[,;]+/).map((c) => c.trim()).filter(Boolean);
    let sawCultivarClause = false;
    let cultivarClauseHadSize = false;
    for (const clause of clauses) {
      const hits = textToks.filter((t) => clause.includes(t)).length;
      if (hits === 0) continue;
      const strong = hits >= Math.ceil(textToks.length * 0.6) || hits >= 2 || textToks.length === 1;
      if (!strong) continue;
      sawCultivarClause = true;
      const ranges = extractInchRanges(clause);
      if (ranges.length) {
        cultivarClauseHadSize = true;
        if (ranges.some((r) => size >= r.lo - 1e-6 && size <= r.hi + 1e-6)) return true;
      }
    }
    if (sawCultivarClause && cultivarClauseHadSize) return false;
    if (sawCultivarClause) return sizeCoveredByPlants(plants, size);
    return false;
  }

  function parseQuery(qRaw) {
    const q = norm(qRaw);
    const sizes = parseQuerySizes(q);
    const textQ = stripSizePhrases(q);
    const toks = textQ.split(/[^a-z0-9.]+/).filter((t) => t.length > 1);
    return { q, sizes, textQ, toks };
  }

  function scorePlot(p, parsed) {
    const { q, sizes, textQ, toks } = parsed;
    if (!q) return 0;
    const id = p.id.toLowerCase();
    const idNodot = id.replace(/\./g, "");
    const plants = (p.plants || "").toLowerCase();
    const search = (p.search || `${id} ${plants}`).toLowerCase();
    let score = 0;
    const qForId = textQ || q;
    const qNodot = qForId.replace(/\./g, "");

    if (textQ) {
      if (id === textQ || idNodot === qNodot) score += 1000;
      else if (id.startsWith(textQ) || idNodot.startsWith(qNodot)) score += 400;
      else if (id.includes(textQ) || idNodot.includes(qNodot)) score += 120;
      if (plants.includes(textQ)) score += 200;
    }

    if (toks.length) {
      let all = true;
      let partial = 0;
      for (const t of toks) {
        const inId = id.includes(t) || idNodot.includes(t.replace(/\./g, ""));
        const inPlant = plants.includes(t);
        const inTok = (p.tokens || []).some((x) => x === t || x.startsWith(t) || t.startsWith(x));
        if (inId || inPlant || inTok) partial += 1;
        else all = false;
      }
      if (all) score += 80 + toks.length * 15;
      else score += partial * 10;
    } else if (!sizes.length && search.includes(q)) {
      score += 40;
    }

    if (sizes.length) {
      let ok = true;
      let nearBonus = 0;
      for (const sz of sizes) {
        if (!sizeNearCultivar(plants, toks, sz)) {
          ok = false;
          break;
        }
        if (toks.length && sizeCoveredByPlants(plants, sz)) nearBonus += 50;
        const ranges = extractInchRanges(plants).filter(
          (r) => sz >= r.lo - 1e-6 && sz <= r.hi + 1e-6
        );
        if (ranges.length) {
          const tight = Math.min(...ranges.map((r) => r.hi - r.lo));
          nearBonus += Math.max(0, 40 - tight * 8);
        }
      }
      if (!ok) return 0;
      score += 100 + nearBonus;
    }

    if (score === 0 && !sizes.length && search.includes(q)) score += 40;
    return score;
  }

  function search(qRaw) {
    const parsed = parseQuery(qRaw);
    const q = parsed.q;
    state.query = q;
    state.parsed = parsed;
    if (!q) {
      state.hits = [];
      state.selected = new Set();
      renderResults();
      clearHighlights();
      return;
    }
    const scored = [];
    for (const p of state.plots) {
      const s = scorePlot(p, parsed);
      if (s > 0) scored.push({ p, s });
    }
    scored.sort((a, b) => b.s - a.s || a.p.id.localeCompare(b.p.id, undefined, { numeric: true }));
    state.hits = scored.slice(0, 80).map((x) => x.p);
    state.selected = new Set();
    if (state.hits.length) {
      const top = scorePlot(state.hits[0], parsed);
      for (const h of state.hits) {
        if (
          scorePlot(h, parsed) >= top * 0.85 ||
          h.id.toLowerCase() === q ||
          h.id.replace(/\./g, "") === q.replace(/\./g, "")
        ) {
          state.selected.add(h.id);
        }
      }
      const idProbe = parsed.textQ || q;
      if (/^\d+(\.\d+[a-z]*)?$/i.test(idProbe) || /^\d+\.\d+$/.test(idProbe)) {
        state.selected = new Set(
          state.hits
            .filter((h) => {
              const id = h.id.toLowerCase();
              const idn = id.replace(/\./g, "");
              const qn = idProbe.replace(/\./g, "");
              return id === idProbe || idn === qn || id.startsWith(idProbe);
            })
            .map((h) => h.id)
        );
        const exact = state.hits.find(
          (h) =>
            h.id.toLowerCase() === idProbe ||
            h.id.replace(/\./g, "").toLowerCase() === idProbe.replace(/\./g, "")
        );
        if (exact) state.selected = new Set([exact.id]);
      }
      if (!state.selected.size && state.hits.length) {
        state.hits.slice(0, 12).forEach((h) => state.selected.add(h.id));
      }
    }
    // Every selected hit: labeled auto or custom polygon (canonical id). Pin only if that plot id has no outline.
    rebuildPolyIndex();
    renderResults();
    renderPolygons();
    renderPins();
  }

  function highlightText(text, q) {
    if (!q || !text) return escapeHtml(text || "");
    const esc = escapeHtml(text);
    const parts = q.split(/[^a-z0-9.]+/).filter((t) => t.length > 1);
    const sizes = (state.parsed && state.parsed.sizes) || parseQuerySizes(q);
    let out = esc;
    if (parts.length) {
      const re = new RegExp("(" + parts.map(escapeReg).join("|") + ")", "ig");
      out = out.replace(re, "<mark>$1</mark>");
    }
    if (sizes.length) {
      for (const r of extractInchRanges(text)) {
        if (!sizes.some((sz) => sz >= r.lo - 1e-6 && sz <= r.hi + 1e-6)) continue;
        const rawEsc = escapeHtml(r.raw);
        out = out.replace(rawEsc, "<mark>" + rawEsc + "</mark>");
      }
    }
    return out;
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
  function escapeReg(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function pathToSvg(path) {
    if (!path || path.length < 2) return "";
    let d = "";
    for (let i = 0; i < path.length; i++) {
      const [x, y] = path[i];
      d += (i === 0 ? "M" : "L") + x + " " + y + " ";
    }
    return d + "Z";
  }

  function pathToPolyline(path) {
    if (!path || path.length < 1) return "";
    let d = "";
    for (let i = 0; i < path.length; i++) {
      const [x, y] = path[i];
      d += (i === 0 ? "M" : "L") + x + " " + y + " ";
    }
    return d;
  }

  function centroidOf(feat) {
    const path = feat && feat.path;
    if (!path || !path.length) return { x: 50, y: 50 };
    let sx = 0, sy = 0;
    for (const pt of path) {
      sx += pt[0];
      sy += pt[1];
    }
    return { x: sx / path.length, y: sy / path.length };
  }

  function polygonSortKey(a, b) {
    const fa = a.farm || 0,
      fb = b.farm || 0;
    if (fa !== fb) return fa - fb;
    const ca = centroidOf(a),
      cb = centroidOf(b);
    if (ca.y !== cb.y) return ca.y - cb.y;
    return ca.x - cb.x;
  }

  function unlabeledQueue() {
    return state.polygons.filter((f) => !effectiveId(f)).slice().sort(polygonSortKey);
  }

  function labeledQueue() {
    return state.polygons.filter((f) => !!effectiveId(f)).slice().sort(polygonSortKey);
  }

  function unusedPlotIdsForFarm(farm) {
    const used = new Set();
    for (const f of state.polygons) {
      const id = effectiveId(f);
      if (id) used.add(id);
    }
    for (const o of state.customOutlines) {
      if (o && o.id) used.add(o.id);
    }
    return state.plots
      .filter((p) => (!farm || p.farm === farm) && !used.has(p.id))
      .map((p) => p.id);
  }

  function focusQueueItem(feat, { scroll = true } = {}) {
    if (!feat) return;
    state.labelFocusIndex = feat.index;
    const curId = effectiveId(feat) || "";
    const input = $("#labelInput");
    if (input) {
      // Prefill current id (OCR / prior label) so Aj can verify & correct
      input.value = curId;
      const suggestions = unusedPlotIdsForFarm(feat.farm).slice(0, 3);
      input.placeholder = suggestions.length ? `e.g. ${suggestions[0]}` : "e.g. 1.12";
      input.focus();
      if (curId) input.select();
    }
    renderPolygons();
    renderLabelQueue();
    if (scroll) scrollFeatureIntoView(feat);
    const activeQ = $(`.lq-item.active`);
    if (activeQ) activeQ.scrollIntoView({ block: "nearest", behavior: "smooth" });
    const meta = $("#labelQueueMeta");
    const q = unlabeledQueue();
    const pos = q.findIndex((f) => f.index === feat.index);
    if (meta) {
      if (curId) {
        meta.textContent = `Editing ${curId} · Farm ${feat.farm || "?"} · outline #${feat.index} · ${q.length} still unlabeled`;
      } else if (q.length) {
        meta.textContent = `Needs label: ${q.length} · working ${pos + 1}/${q.length} · Farm ${feat.farm || "?"} · outline #${feat.index}`;
      } else {
        meta.textContent = "All outlines labeled — click any block to correct";
      }
    }
  }

  function focusNextUnlabeled(dir = 1) {
    const q = unlabeledQueue();
    if (!q.length) {
      state.labelFocusIndex = null;
      renderPolygons();
      renderLabelQueue();
      const meta = $("#labelQueueMeta");
      if (meta) meta.textContent = "All outlines labeled — Copy labels JSON when ready";
      return;
    }
    let pos = q.findIndex((f) => f.index === state.labelFocusIndex);
    if (pos < 0) pos = dir > 0 ? -1 : 0;
    pos = (pos + dir + q.length) % q.length;
    focusQueueItem(q[pos]);
  }

  function renderLabelQueue() {
    const box = $("#labelQueue");
    if (!box) return;
    if (!state.labelMode) {
      box.hidden = true;
      box.setAttribute("hidden", "");
      return;
    }
    box.hidden = false;
    box.removeAttribute("hidden");
    const q = unlabeledQueue();
    const labeled = labeledQueue();
    const parts = [];
    if (q.length) {
      parts.push(
        `<div class="lq-head">Needs label — ${q.length} (purple on map; yellow = current)</div>`,
        `<div class="lq-hint">Pick a row or click ANY block on the map (labeled or not). Type printed plot id, Enter / Apply + next.</div>`
      );
      for (let i = 0; i < q.length; i++) {
        const f = q[i];
        const c = centroidOf(f);
        const active = state.labelFocusIndex === f.index ? " active" : "";
        const sugg = unusedPlotIdsForFarm(f.farm).slice(0, 2).join(", ");
        parts.push(
          `<button type="button" class="lq-item${active}" data-qindex="${f.index}">` +
            `<span class="lq-num">#${i + 1}</span>` +
            `<span><div>Farm ${f.farm || "?"} · outline ${f.index}</div>` +
            `<div class="lq-meta">map ~${c.x.toFixed(0)}%,${c.y.toFixed(0)}%` +
            (sugg ? ` · unused nearby: ${escapeHtml(sugg)}` : "") +
            `</div></span></button>`
        );
      }
    } else {
      parts.push(
        `<div class="lq-head">Label queue</div>`,
        `<div class="lq-hint">All ${state.polygons.length} outlines have plot ids. Click any teal block on the map (or a row below) to verify/correct OCR. Copy labels JSON when ready.</div>`
      );
    }
    if (labeled.length) {
      parts.push(
        `<div class="lq-head">Labeled — ${labeled.length} (click to correct)</div>`,
        `<div class="lq-hint">Teal on map. Click a row or the outline to edit the id.</div>`
      );
      for (const f of labeled) {
        const id = effectiveId(f) || "";
        const c = centroidOf(f);
        const active = state.labelFocusIndex === f.index ? " active" : "";
        parts.push(
          `<button type="button" class="lq-item lq-labeled${active}" data-qindex="${f.index}">` +
            `<span class="lq-num">F${f.farm || "?"}</span>` +
            `<span><div>outline ${f.index}</div>` +
            `<div class="lq-meta">map ~${c.x.toFixed(0)}%,${c.y.toFixed(0)}%</div></span>` +
            `<span class="lq-id">${escapeHtml(id)}</span></button>`
        );
      }
    }
    box.innerHTML = parts.join("");
    box.querySelectorAll(".lq-item").forEach((el) => {
      el.addEventListener("click", () => {
        const idx = parseInt(el.getAttribute("data-qindex"), 10);
        const feat = state.polygons.find((f) => f.index === idx);
        focusQueueItem(feat);
      });
    });
  }

  function renderPolygons() {
    const svg = $("#polyLayer");
    if (!svg) return;
    const multi = state.selected.size > 1;
    const parts = [];
    const idle = !state.labelMode && !state.outlineMode;

    for (const f of state.polygons) {
      const id = effectiveId(f);
      const selected = outlineSelected(id);
      // hideAuto: skip non-selected autos (Outline help checkbox). Search-selected still paint.
      if (state.hideAutoOutlines && !state.labelMode && !selected) continue;
      const unlabeled = !id;
      let cls = "plot-poly";
      if (state.labelMode) {
        // Label mode: never paint search .selected flood — only needs-label / label-focus
        if (unlabeled) cls += " unlabeled needs-label";
        else cls += " labeled-idle";
        if (state.labelFocusIndex === f.index) cls += " label-focus";
      } else if (state.outlineMode) {
        if (state.selectedAutoIndex === f.index) cls += " outline-selected";
        else if (selected) cls += " selected" + (multi ? " multi" : "");
        else cls += " dim";
        if (unlabeled) cls += " unlabeled";
      } else {
        // Idle: invisible unless search-selected (amber block highlight)
        if (selected) cls += " selected" + (multi ? " multi" : "");
      }
      const title = id || `(unlabeled #${f.index})`;
      parts.push(
        `<path class="${cls}" data-index="${f.index}" data-id="${escapeHtml(id || "")}" d="${pathToSvg(f.path)}"><title>${escapeHtml(title)}</title></path>`
      );
    }

    for (let i = 0; i < state.customOutlines.length; i++) {
      const o = state.customOutlines[i];
      if (!o || !o.path) continue;
      const id = canonicalPlotId(o.id) || "";
      const selected = outlineSelected(id);
      const customSel = state.selectedCustomIdx === i;
      let cls = "custom-poly";
      if (idle) {
        // Idle: invisible until search (or rare custom selection)
        if (selected) cls += " hit-selected";
        else if (customSel) cls += " selected";
        else cls += " idle-hidden";
      } else {
        if (customSel) cls += " selected";
        else if (selected) cls += " hit-selected";
      }
      const title = id ? `${id} (custom)` : `(custom #${i})`;
      parts.push(
        `<path class="${cls}" data-custom="${i}" data-id="${escapeHtml(id)}" d="${pathToSvg(o.path)}"><title>${escapeHtml(title)}</title></path>`
      );
    }

    svg.innerHTML = parts.join("");

    svg.querySelectorAll("path.plot-poly").forEach((el) => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        const idx = parseInt(el.getAttribute("data-index"), 10);
        const feat = state.polygons.find((f) => f.index === idx) || state.polygons[idx];
        if (!feat) return;
        if (state.outlineMode) {
          state.selectedAutoIndex = feat.index;
          state.selectedCustomIdx = null;
          state.pendingPath = null;
          state.drawPoints = [];
          renderDrawLayer();
          renderPolygons();
          updateOutlineMeta();
          return;
        }
        if (state.labelMode) {
          // Every outline (labeled + unlabeled) is editable in Label mode
          focusQueueItem(feat, { scroll: false });
          return;
        }
        const id = effectiveId(feat);
        if (!id) return;
        state.selected = new Set([id]);
        state.photoPanelClosed = false;
        if (!state.hits.some((h) => h.id === id)) {
          const p = state.plots.find((x) => x.id === id);
          if (p) state.hits = [p, ...state.hits].slice(0, 80);
        }
        renderResults();
        renderPolygons();
        renderPins();
        updatePhotoPanel();
        scrollFeatureIntoView(feat);
        const hit = $(`.hit[data-id="${CSS.escape(id)}"]`);
        if (hit) hit.scrollIntoView({ block: "nearest", behavior: "smooth" });
      });
    });

    svg.querySelectorAll("path.custom-poly").forEach((el) => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        const ci = parseInt(el.getAttribute("data-custom"), 10);
        const o = state.customOutlines[ci];
        if (!o) return;
        if (state.outlineMode) {
          state.selectedCustomIdx = ci;
          state.selectedAutoIndex = null;
          state.pendingPath = null;
          state.drawPoints = [];
          renderDrawLayer();
          const input = $("#outlineIdInput");
          if (input) {
            input.value = o.id || "";
            input.focus();
            input.select();
          }
          renderPolygons();
          updateOutlineMeta();
          return;
        }
        if (o.id) {
          state.selected = new Set([o.id]);
          state.photoPanelClosed = false;
          if (!state.hits.some((h) => h.id === o.id)) {
            const p = state.plots.find((x) => x.id === o.id);
            if (p) state.hits = [p, ...state.hits].slice(0, 80);
          }
          renderResults();
          renderPolygons();
          renderPins();
          updatePhotoPanel();
          scrollFeatureIntoView({ path: o.path, centroid: centroidFromPath(o.path) });
        }
      });
    });

    updateMapStatus();
    renderDrawLayer();
  }

  function renderDrawLayer() {
    const svg = $("#drawLayer");
    if (!svg) return;
    const pts = state.drawPoints;
    if (!pts.length) {
      svg.innerHTML = "";
      return;
    }
    const bits = [];
    if (pts.length >= 2) {
      bits.push(`<path class="draw-line" d="${pathToPolyline(pts)}" />`);
    }
    if (pts.length >= 3) {
      bits.push(`<path class="draw-preview" d="${pathToSvg(pts)}" />`);
    }
    for (const [x, y] of pts) {
      bits.push(`<circle class="draw-dot" cx="${x}" cy="${y}" r="0.45" />`);
    }
    svg.innerHTML = bits.join("");
  }

  function updateOutlineMeta() {
    const meta = $("#outlineMeta");
    if (!meta) return;
    const n = state.customOutlines.length;
    const autoN = state.polygons.length;
    const drawing = state.drawPoints.length;
    const pending = state.pendingPath ? state.pendingPath.length : 0;
    const selCustom =
      state.selectedCustomIdx != null && state.customOutlines[state.selectedCustomIdx]
        ? state.customOutlines[state.selectedCustomIdx].id || `#${state.selectedCustomIdx}`
        : null;
    const parts = [`${autoN} auto`, `${n} custom`];
    if (drawing) parts.push(`drawing ${drawing} pts`);
    if (pending) parts.push(`closed — type id + Save`);
    if (selCustom) parts.push(`custom ${selCustom} selected · Delete to remove`);
    else if (state.selectedAutoIndex != null) {
      parts.push(`auto #${state.selectedAutoIndex} selected · Delete to remove`);
    }
    meta.textContent = parts.join(" · ");
  }

  function updateMapStatus() {
    const ids = [...state.selected];
    let polyN = 0;
    let pinN = 0;
    let missing = 0;
    for (const id of ids) {
      if (outlineFor(id)) polyN++;
      else if (state.coords[id]) pinN++;
      else missing++;
    }
    const labeled = state.polygons.filter((f) => effectiveId(f)).length;
    const unlabeled = state.polygons.length - labeled;
    const customN = state.customOutlines.filter((o) => o && o.id).length;
    const partial = state.hideAutoOutlines
      ? ` · outlines off · custom ${customN} · Label mode to assign ids`
      : ` · ${state.polygons.length} outlines · unlabeled ${unlabeled} · labeled ${labeled} · custom ${customN}`;
    if (!ids.length) {
      if (state.outlineMode) {
        $("#mapStatus").textContent = `Outline help on · hide auto ${state.hideAutoOutlines ? "ON" : "OFF"}${partial}`;
      } else if (state.labelMode) {
        $("#mapStatus").textContent = `Label mode on · click outline → type plot id${partial}`;
      } else {
        $("#mapStatus").textContent = `No highlight${partial}`;
      }
      return;
    }
    const bits = [];
    if (polyN) bits.push(`${polyN} outline${polyN === 1 ? "" : "s"}`);
    if (pinN) bits.push(`${pinN} pin fallback${pinN === 1 ? "" : "s"}`);
    if (missing) bits.push(`${missing} missing`);
    $("#mapStatus").textContent = bits.join(" · ") + partial;
  }

  function clearHighlights() {
    const stage = $("#pins");
    if (stage) stage.innerHTML = "";
    renderPolygons();
    updateMapStatus();
  }

  function renderPins() {
    const stage = $("#pins");
    if (!stage) return;
    stage.innerHTML = "";
    const ids = [...state.selected];
    const multi = ids.length > 1;
    for (const id of ids) {
      if (outlineFor(id)) continue;
      const c = state.coords[id];
      if (!c) continue;
      const pin = document.createElement("div");
      pin.className = "pin" + (multi ? " multi" : "");
      pin.style.left = c.x + "%";
      pin.style.top = c.y + "%";
      pin.title = id + " (pin fallback)";
      pin.innerHTML = `<div class="dot"></div><div class="label">${escapeHtml(id)}</div>`;
      pin.addEventListener("click", (e) => {
        e.stopPropagation();
        state.selected = new Set([id]);
        state.photoPanelClosed = false;
        renderResults();
        renderPolygons();
        renderPins();
        updatePhotoPanel();
        const hit = $(`.hit[data-id="${CSS.escape(id)}"]`);
        if (hit) hit.scrollIntoView({ block: "nearest", behavior: "smooth" });
      });
      stage.appendChild(pin);
    }
    updateMapStatus();
  }

  function renderResults() {
    const box = $("#results");
    const meta = $("#resultMeta");
    if (!state.query) {
      meta.textContent =
        'Type cultivar, plot, or size (e.g. October Glory 3", boxwood, 1.12) — sizes match ranges too';
      box.innerHTML = `<div class="empty">Search the live <b>25-26</b> inventory.<br>Matching plots highlight on the nursery map.</div>`;
      updatePhotoPanel();
      return;
    }
    const sizeNote =
      state.parsed && state.parsed.sizes && state.parsed.sizes.length
        ? `<span class="chip">size ${state.parsed.sizes.map((n) => n + '"').join(", ")} in range</span>`
        : "";
    meta.innerHTML = `<span class="chip"><b>${state.hits.length}</b> match${state.hits.length === 1 ? "" : "es"}</span>
      <span class="chip">${state.selected.size} highlighted</span>${sizeNote}`;
    if (!state.hits.length) {
      box.innerHTML = `<div class="empty">No plots match <b>${escapeHtml(state.query)}</b>.</div>`;
      updatePhotoPanel();
      return;
    }
    box.innerHTML = state.hits
      .map((p) => {
        const active = state.selected.has(p.id) ? " active" : "";
        const empty = p.empty ? `<span class="empty-tag">OPEN/EMPTY</span>` : "";
        const hasPoly = !!outlineFor(p.id);
        const mapTag = hasPoly
          ? `<span class="empty-tag" style="color:var(--green);background:#123524;border-color:#1f6b52">outline</span>`
          : state.coords[p.id]
            ? `<span class="empty-tag" style="color:var(--accent)">pin</span>`
            : "";
        const plants =
          p.empty && !p.plants
            ? `<span style="color:var(--muted)">(no plant text — open/blank)</span>`
            : highlightText(p.plants || "(empty)", state.query);
        return `<div class="hit${active}" data-id="${escapeHtml(p.id)}" role="button" tabindex="0">
          <div class="top">
            <span class="plot">${escapeHtml(p.id)}</span>
            <span class="farm">Farm ${p.farm}</span>
            ${empty}${mapTag}
          </div>
          <div class="plants">${plants}</div>
        </div>`;
      })
      .join("");
    box.querySelectorAll(".hit").forEach((el) => {
      el.addEventListener("click", () => {
        const id = el.getAttribute("data-id");
        state.selected = new Set([id]);
        state.photoPanelClosed = false;
        renderResults();
        renderPolygons();
        renderPins();
        updatePhotoPanel();
        const feat = outlineFor(id);
        if (feat) scrollFeatureIntoView(feat);
        else scrollPinIntoView(id);
      });
    });
    updatePhotoPanel();
  }

  function scrollPinIntoView(id) {
    const c = state.coords[id];
    if (!c) return;
    scrollPct(c.x, c.y);
  }

  function scrollFeatureIntoView(feat) {
    if (!feat) return;
    let x = 50,
      y = 50;
    if (feat.centroid && feat.centroid.length >= 2) {
      x = feat.centroid[0];
      y = feat.centroid[1];
    } else {
      const c = centroidOf(feat);
      x = c.x;
      y = c.y;
    }
    scrollPct(x, y);
  }

  function scrollPct(x, y) {
    const wrap = $(".map-wrap");
    const stage = $(".map-stage");
    if (!wrap || !stage) return;
    const rect = stage.getBoundingClientRect();
    const px = (x / 100) * rect.width;
    const py = (y / 100) * rect.height;
    wrap.scrollTo({
      left: Math.max(0, stage.offsetLeft + px - wrap.clientWidth / 2),
      top: Math.max(0, stage.offsetTop + py - wrap.clientHeight / 2),
      behavior: "smooth",
    });
  }

  function applyLabel() {
    if (state.labelFocusIndex == null) return;
    const raw = ($("#labelInput").value || "").trim();
    const key = String(state.labelFocusIndex);
    if (!raw) {
      state.localLabels[key] = "";
    } else {
      const hit = state.plots.find(
        (p) =>
          p.id.toLowerCase() === raw.toLowerCase() ||
          p.id.replace(/\./g, "").toLowerCase() === raw.replace(/\./g, "").toLowerCase()
      );
      state.localLabels[key] = hit ? hit.id : raw;
    }
    saveLocalLabels();
    rebuildPolyIndex();
    if (raw) focusNextUnlabeled(1);
    else {
      renderPolygons();
      renderPins();
      renderResults();
      renderLabelQueue();
    }
  }

  function resolvePlotId(raw) {
    const t = (raw || "").trim();
    if (!t) return null;
    const hit = state.plots.find(
      (p) =>
        p.id.toLowerCase() === t.toLowerCase() ||
        p.id.replace(/\./g, "").toLowerCase() === t.replace(/\./g, "").toLowerCase()
    );
    return hit ? hit.id : t;
  }

  function copyLabelsJson() {
    const labels = {};
    for (const [k, v] of Object.entries(state.localLabels)) {
      labels[k] = v === "" || v == null ? null : v;
    }
    const payload = {
      source: "got-dirt label mode",
      generatedAt: new Date().toISOString(),
      labels,
    };
    const text = JSON.stringify(payload, null, 2);
    const done = () => {
      const btn = $("#btnCopyLabels");
      if (!btn) return;
      const prev = btn.textContent;
      btn.textContent = "Copied!";
      setTimeout(() => (btn.textContent = prev), 1200);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopy(text, done));
    } else {
      fallbackCopy(text, done);
    }
  }

  function copyOutlinesJson() {
    const payload = {
      source: "got-dirt outline help",
      generatedAt: new Date().toISOString(),
      outlines: state.customOutlines.map((o) => ({
        id: o.id || null,
        path: o.path,
        source: o.source || "aj",
      })),
    };
    const text = JSON.stringify(payload, null, 2);
    const done = () => {
      const btn = $("#btnCopyOutlines");
      if (!btn) return;
      const prev = btn.textContent;
      btn.textContent = "Copied!";
      setTimeout(() => (btn.textContent = prev), 1200);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopy(text, done));
    } else {
      fallbackCopy(text, done);
    }
  }

  function fallbackCopy(text, done) {
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
    } catch (_) {}
    document.body.removeChild(ta);
    done();
  }

  function setLabelMode(on) {
    state.labelMode = !!on;
    if (state.labelMode && state.outlineMode) setOutlineMode(false);
    if (state.labelMode) {
      // Label mode needs autos visible — full-block outlines to click/assign
      state.hideAutoOutlines = false;
      const chk = $("#chkHideAuto");
      if (chk) chk.checked = false;
      // Drop search selection so map is not flood-highlighted amber
      state.selected = new Set();
      state.selectedCustomIdx = null;
      const pins = $("#pins");
      if (pins) pins.innerHTML = "";
      rebuildPolyIndex();
    }
    document.body.classList.toggle("label-mode", state.labelMode);
    const btn = $("#btnLabelMode");
    if (btn) {
      btn.classList.toggle("active-toggle", state.labelMode);
      btn.setAttribute("aria-pressed", state.labelMode ? "true" : "false");
      btn.textContent = state.labelMode ? "Label mode ON" : "Label mode";
    }
    const bar = $("#labelBar");
    if (bar) {
      if (state.labelMode) bar.removeAttribute("hidden");
      else bar.setAttribute("hidden", "");
      bar.hidden = !state.labelMode;
    }
    const copyBtn = $("#btnCopyLabels");
    if (copyBtn) copyBtn.hidden = !state.labelMode;
    const results = $("#results");
    if (results) results.style.display = state.labelMode ? "none" : "";
    if (!state.labelMode) {
      state.labelFocusIndex = null;
      renderLabelQueue();
      try {
        renderPolygons();
      } catch (err) {
        console.error(err);
      }
      try {
        updateMapStatus();
      } catch (err) {
        console.error(err);
      }
      const status = $("#mapStatus");
      if (status && !state.outlineMode) status.style.color = "";
      return;
    }
    try {
      renderPolygons();
    } catch (err) {
      console.error("renderPolygons failed", err);
    }
    renderLabelQueue();
    focusNextUnlabeled(1);
    const status = $("#mapStatus");
    if (status) {
      status.textContent = "LABEL QUEUE — all purple need ids; yellow is current. Use list or Next.";
      status.style.color = "#c4b5fd";
    }
  }

  function setOutlineMode(on) {
    state.outlineMode = !!on;
    if (state.outlineMode && state.labelMode) setLabelMode(false);
    document.body.classList.toggle("outline-mode", state.outlineMode);
    const btn = $("#btnOutlineMode");
    if (btn) {
      btn.classList.toggle("active-toggle", state.outlineMode);
      btn.classList.toggle("outline-toggle", state.outlineMode);
      btn.setAttribute("aria-pressed", state.outlineMode ? "true" : "false");
      btn.textContent = state.outlineMode ? "Outline help ON" : "Outline help";
    }
    const bar = $("#outlineBar");
    if (bar) {
      if (state.outlineMode) bar.removeAttribute("hidden");
      else bar.setAttribute("hidden", "");
      bar.hidden = !state.outlineMode;
    }
    if (state.outlineMode) {
      const chk = $("#chkHideAuto");
      if (chk) state.hideAutoOutlines = !!chk.checked;
      rebuildPolyIndex();
      renderPolygons();
      renderPins();
      updateOutlineMeta();
      const status = $("#mapStatus");
      if (status) {
        status.textContent = "OUTLINE HELP — click auto to select/Delete, or draw corners; uncheck Hide auto to see autos";
        status.style.color = "#5eead4";
      }
    } else {
      state.drawPoints = [];
      state.pendingPath = null;
      state.selectedCustomIdx = null;
      state.selectedAutoIndex = null;
      renderDrawLayer();
      // Keep auto hidden (number-split junk) unless checkbox says otherwise
      const chk = $("#chkHideAuto");
      if (chk) state.hideAutoOutlines = !!chk.checked;
      else state.hideAutoOutlines = true;
      rebuildPolyIndex();
      renderPolygons();
      renderPins();
      updateMapStatus();
      const status = $("#mapStatus");
      if (status && !state.labelMode) status.style.color = "";
    }
  }

  function stagePctFromEvent(e) {
    const stage = $("#mapStage") || $(".map-stage");
    if (!stage) return null;
    const rect = stage.getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    const x = ((e.clientX - rect.left) / rect.width) * 100;
    const y = ((e.clientY - rect.top) / rect.height) * 100;
    return [
      Math.max(0, Math.min(100, Math.round(x * 1000) / 1000)),
      Math.max(0, Math.min(100, Math.round(y * 1000) / 1000)),
    ];
  }

  function closeDrawPolygon() {
    if (state.drawPoints.length < 3) {
      updateOutlineMeta();
      return false;
    }
    state.pendingPath = state.drawPoints.slice();
    state.drawPoints = [];
    state.selectedCustomIdx = null;
    state.selectedAutoIndex = null;
    renderDrawLayer();
    // show closed preview as temporary custom-looking path via draw layer
    const svg = $("#drawLayer");
    if (svg && state.pendingPath) {
      svg.innerHTML = `<path class="draw-preview" d="${pathToSvg(state.pendingPath)}" />`;
    }
    const input = $("#outlineIdInput");
    if (input) {
      input.focus();
      input.select();
    }
    updateOutlineMeta();
    return true;
  }

  function saveOutline() {
    let path = state.pendingPath;
    const input = $("#outlineIdInput");
    const raw = (input && input.value) || "";
    const id = resolvePlotId(raw);

    if (state.selectedCustomIdx != null && state.customOutlines[state.selectedCustomIdx]) {
      // Re-id selected custom outline
      if (!id) return;
      state.customOutlines[state.selectedCustomIdx].id = id;
      saveCustomOutlines();
      rebuildPolyIndex();
      renderPolygons();
      renderPins();
      renderResults();
      if (input) input.value = "";
      updateOutlineMeta();
      return;
    }

    if (!path || path.length < 3) {
      if (state.drawPoints.length >= 3) {
        closeDrawPolygon();
        path = state.pendingPath;
      }
    }
    if (!path || path.length < 3 || !id) {
      updateOutlineMeta();
      return;
    }
    // Replace existing custom with same id
    const existing = state.customOutlines.findIndex((o) => o.id === id);
    const entry = { id, path: path.map((p) => [p[0], p[1]]), source: "aj" };
    if (existing >= 0) state.customOutlines[existing] = entry;
    else state.customOutlines.push(entry);
    state.pendingPath = null;
    state.drawPoints = [];
    state.selectedCustomIdx = null;
    state.selectedAutoIndex = null;
    saveCustomOutlines();
    rebuildPolyIndex();
    renderPolygons();
    renderPins();
    renderResults();
    if (input) input.value = "";
    updateOutlineMeta();
  }

  function cancelDraw() {
    state.drawPoints = [];
    state.pendingPath = null;
    renderDrawLayer();
    updateOutlineMeta();
  }

  function deleteSelectedCustom() {
    if (state.selectedCustomIdx == null) return;
    const i = state.selectedCustomIdx;
    if (i < 0 || i >= state.customOutlines.length) return;
    state.customOutlines.splice(i, 1);
    state.selectedCustomIdx = null;
    state.selectedAutoIndex = null;
    saveCustomOutlines();
    rebuildPolyIndex();
    renderPolygons();
    renderPins();
    renderResults();
    renderLabelQueue();
    updateOutlineMeta();
    updateMapStatus();
  }

  function clearMyOutlines() {
    state.customOutlines = [];
    state.selectedCustomIdx = null;
    state.selectedAutoIndex = null;
    state.pendingPath = null;
    state.drawPoints = [];
    saveCustomOutlines();
    rebuildPolyIndex();
    renderPolygons();
    renderPins();
    renderResults();
    renderLabelQueue();
    updateOutlineMeta();
    updateMapStatus();
  }


  function photosFromAssignments(id) {
    const out = [];
    const seen = new Set();
    for (const a of Object.values(state.photoAssignments || {})) {
      if (!a || typeof a !== "object") continue;
      if (a.status && a.status !== "assigned") continue;
      if (String(a.plotId || "") !== String(id)) continue;
      const rel = a.thumb || null;
      if (!rel || seen.has(rel)) continue;
      seen.add(rel);
      out.push(rel);
    }
    return out;
  }

  function cultivarFromAssignments(id) {
    for (const a of Object.values(state.photoAssignments || {})) {
      if (!a || typeof a !== "object") continue;
      if (a.status && a.status !== "assigned") continue;
      if (String(a.plotId || "") !== String(id)) continue;
      if (a.cultivar) return String(a.cultivar);
    }
    return "";
  }

  /** Pack plot-photos.json + assigned thumbs (localStorage + pack photo-assignments). Never invent IDs. */
  function photoEntryFor(id) {
    if (!id) return null;
    const raw = state.plotPhotos[id];
    let photos = [];
    let cultivar = "";
    if (Array.isArray(raw)) {
      photos = raw.slice();
    } else if (raw && Array.isArray(raw.photos)) {
      photos = raw.photos.slice();
      cultivar = raw.cultivar || "";
    }
    const seen = new Set(photos);
    for (const rel of photosFromAssignments(id)) {
      if (seen.has(rel)) continue;
      seen.add(rel);
      photos.push(rel);
    }
    if (!cultivar) cultivar = cultivarFromAssignments(id);
    if (!photos.length && !cultivar) return null;
    return { photos, cultivar };
  }

  function openLightbox(src, alt) {
    const lb = $("#lightbox");
    const img = $("#lightboxImg");
    if (!lb || !img) return;
    img.src = src;
    img.alt = alt || "Plot photo";
    lb.hidden = false;
  }

  function closeLightbox() {
    const lb = $("#lightbox");
    const img = $("#lightboxImg");
    if (lb) lb.hidden = true;
    if (img) img.removeAttribute("src");
  }

  function updatePhotoPanel() {
    const panel = $("#photoPanel");
    const body = $("#photoPanelBody");
    const plotEl = $("#photoPanelPlot");
    const subEl = $("#photoPanelSub");
    const layout = document.querySelector(".layout");
    if (!panel || !body || !plotEl) return;

    // Assign-photos mode owns the right panel
    if (state.assignMode) {
      const pin = state.gpsPins.find((p) => p.path === state.assignFocusPath);
      updateAssignPreview(pin || null);
      if (layout) layout.classList.add("assign-open", "photo-open");
      return;
    }

    const ids = [...state.selected];
    if (!ids.length || state.photoPanelClosed) {
      plotEl.textContent = "—";
      if (subEl) subEl.textContent = "";
      body.innerHTML =
        `<div class="photo-empty">Select a plot (search hit or map outline) to view photos.</div>`;
      if (layout) layout.classList.remove("photo-open");
      return;
    }

    // Prefer single selection; if multi, use first for panel title but note count
    const id = ids.length === 1 ? ids[0] : ids[0];
    const entry = photoEntryFor(id);
    const plot = state.plots.find((p) => p.id === id);
    plotEl.textContent = id;
    if (subEl) {
      const cult = (entry && entry.cultivar) || "";
      const plants = plot && plot.plants ? String(plot.plants).slice(0, 80) : "";
      subEl.textContent =
        cult || plants || (ids.length > 1 ? `${ids.length} plots highlighted — showing ${id}` : "");
    }
    if (layout) layout.classList.add("photo-open");

    const photos = entry && entry.photos ? entry.photos : [];
    if (!photos.length) {
      body.innerHTML = `<div class="photo-empty">No photos for plot <b>${escapeHtml(id)}</b> yet.</div>`;
      return;
    }

    body.innerHTML =
      `<div class="photo-grid">` +
      photos
        .map((rel) => {
          const src = `data/${rel}?v=30`;
          const name = String(rel).split("/").pop() || rel;
          return `<button type="button" class="photo-thumb" data-src="${escapeHtml(src)}" title="Enlarge">
            <img src="${escapeHtml(src)}" alt="${escapeHtml(name)}" loading="lazy" decoding="async">
            <div class="photo-cap">${escapeHtml(name)}</div>
          </button>`;
        })
        .join("") +
      `</div>`;

    body.querySelectorAll(".photo-thumb").forEach((btn) => {
      btn.addEventListener("click", () => {
        openLightbox(btn.getAttribute("data-src"), btn.querySelector("img")?.alt || "");
      });
    });
  }

  function formatCt(iso) {
    if (!iso) return "—";
    try {
      const d = new Date(iso);
      return (
        d.toLocaleString("en-US", {
          timeZone: "America/Chicago",
          month: "short",
          day: "numeric",
          year: "numeric",
          hour: "numeric",
          minute: "2-digit",
        }) + " CT"
      );
    } catch {
      return iso;
    }
  }


  function loadGpsDecisions() {
    try {
      state.gpsDecisions = JSON.parse(localStorage.getItem(LS_GPS_DECISIONS) || "{}") || {};
    } catch (_) {
      state.gpsDecisions = {};
    }
  }

  function normalizeAssignmentsDoc(doc) {
    if (!doc) return {};
    if (doc.assignments && typeof doc.assignments === "object") return doc.assignments;
    // bare map (localStorage shape)
    return doc;
  }

  function assignmentTimestamp(a) {
    if (!a || !a.at) return 0;
    const t = Date.parse(a.at);
    return Number.isFinite(t) ? t : 0;
  }

  /** Merge maps keyed by photo path; keep newer `at` (tie → keep existing). Never invent plot IDs. */
  function mergePhotoAssignmentMaps(base, incoming) {
    const out = { ...(base || {}) };
    for (const [path, a] of Object.entries(incoming || {})) {
      if (!a || typeof a !== "object") continue;
      const prev = out[path];
      if (!prev) {
        out[path] = a;
        continue;
      }
      if (assignmentTimestamp(a) > assignmentTimestamp(prev)) out[path] = a;
    }
    return out;
  }

  function loadPhotoAssignments() {
    try {
      state.photoAssignments = JSON.parse(localStorage.getItem(LS_PHOTO_ASSIGN) || "{}") || {};
    } catch (_) {
      state.photoAssignments = {};
    }
    // One-way migrate accepted/moved legacy GPS decisions into photoAssignments
    try {
      for (const [path, d] of Object.entries(state.gpsDecisions || {})) {
        if (!d || state.photoAssignments[path]) continue;
        if ((d.status === "accepted" || d.status === "moved") && d.plotId) {
          state.photoAssignments[path] = {
            plotId: String(d.plotId),
            status: "assigned",
            at: d.at || new Date().toISOString(),
            source: "migrated-gpsPinDecisions",
          };
        } else if (d.status === "rejected") {
          state.photoAssignments[path] = {
            plotId: null,
            status: "rejected",
            at: d.at || new Date().toISOString(),
            source: "migrated-gpsPinDecisions",
          };
        }
      }
      persistPhotoAssignmentsLocalOnly();
    } catch (_) {}
  }

  /** Fetch pack file and merge so localhost ↔ Tailscale origin switches keep work. */
  async function loadPhotoAssignmentsFromPack() {
    try {
      const r = await fetch("data/photo-assignments.json?v=30");
      if (!r.ok) return;
      const doc = await r.json();
      const fromPack = normalizeAssignmentsDoc(doc);
      const before = Object.keys(state.photoAssignments || {}).length;
      state.photoAssignments = mergePhotoAssignmentMaps(state.photoAssignments, fromPack);
      const after = Object.keys(state.photoAssignments || {}).length;
      // Persist merged view back to this origin's localStorage (do not PUT yet — avoids empty wipe races)
      persistPhotoAssignmentsLocalOnly();
      // If this browser origin had nothing but pack had data, also push pack file refresh timestamp locally only
      if (after > before) {
        try {
          renderAssignList();
          renderGpsPins();
        } catch (_) {}
      }
    } catch (_) {}
  }

  function persistPhotoAssignmentsLocalOnly() {
    try {
      localStorage.setItem(LS_PHOTO_ASSIGN, JSON.stringify(state.photoAssignments));
    } catch (_) {}
  }

  function savePhotoAssignments() {
    persistPhotoAssignmentsLocalOnly();
    // Persist to pack file (survives rebuilds + origin switches). Merge-safe; never clobber pack with empty.
    try {
      const localMap = state.photoAssignments || {};
      const localCount = Object.keys(localMap).length;
      fetch("data/photo-assignments.json")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null)
        .then((doc) => {
          const packMap = normalizeAssignmentsDoc(doc) || {};
          const packCount = Object.keys(packMap).length;
          if (localCount === 0 && packCount > 0) {
            // Empty origin must not wipe pack (e.g. fresh browser / other origin).
            state.photoAssignments = mergePhotoAssignmentMaps(packMap, localMap);
            persistPhotoAssignmentsLocalOnly();
            return null;
          }
          const merged = mergePhotoAssignmentMaps(packMap, localMap);
          state.photoAssignments = merged;
          persistPhotoAssignmentsLocalOnly();
          const body = JSON.stringify({
            version: 1,
            updatedAt: new Date().toISOString(),
            count: Object.keys(merged).length,
            assignments: merged,
          });
          return fetch("data/photo-assignments.json", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body,
          });
        })
        .catch(() => {});
    } catch (_) {}
  }

  function assignmentFor(path) {
    return (state.photoAssignments && state.photoAssignments[path]) || null;
  }

  function assignedPlotId(pin) {
    const a = assignmentFor(pin.path);
    if (a && a.status === "assigned" && a.plotId) return a.plotId;
    return null;
  }

  function isPinAssigned(pin) {
    const a = assignmentFor(pin.path);
    return !!(a && a.status === "assigned" && a.plotId);
  }

  function isPinRejected(pin) {
    const a = assignmentFor(pin.path);
    return !!(a && a.status === "rejected");
  }

  function liveLabeledPlotIds() {
    const ids = new Set();
    for (const id of state.polyById.keys()) if (id) ids.add(String(id));
    for (const o of state.customOutlines || []) if (o && o.id) ids.add(String(o.id));
    for (const v of Object.values(state.localLabels || {})) if (v) ids.add(String(v));
    for (const p of state.plots || []) if (p && p.id) ids.add(String(p.id));
    return [...ids].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  function knownPlotId(plotId) {
    if (!plotId) return false;
    return liveLabeledPlotIds().includes(plotId);
  }

  function filteredAssignPins() {
    const q = (state.assignQuery || "").trim().toLowerCase();
    let list = state.gpsPins.slice();
    const filt = state.assignFilter || "unassigned";
    if (filt === "unassigned") {
      list = list.filter((p) => !isPinAssigned(p) && !isPinRejected(p));
    } else if (filt === "assigned") {
      list = list.filter((p) => isPinAssigned(p));
    }
    if (q) {
      list = list.filter((p) => {
        const hay = `${p.cultivar || ""} ${p.path || ""}`.toLowerCase();
        return hay.includes(q);
      });
    }
    return list;
  }

  function formatPhotoDto(dto) {
    if (!dto) return "";
    // EXIF often "YYYY:MM:DD HH:MM:SS"
    const m = String(dto).match(/^(\d{4}):(\d{2}):(\d{2})(?:\s+(\d{2}):(\d{2}))?/);
    if (!m) return String(dto);
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const mon = months[Math.max(0, parseInt(m[2], 10) - 1)] || m[2];
    let s = `${mon} ${parseInt(m[3], 10)}, ${m[1]}`;
    if (m[4] != null) s += ` · ${parseInt(m[4], 10)}:${m[5]}`;
    return s;
  }

  function thumbUrlForPin(pin) {
    if (pin.thumb) return `data/${pin.thumb}?v=30`;
    // derive from path hash+stem convention matching manifest
    return null;
  }

  function renderGpsPins() {
    const stage = $("#gpsPins");
    if (!stage) return;
    stage.innerHTML = "";
    if (!state.assignMode || !state.gpsShowPins) return;
    const focusPath = state.assignFocusPath;
    // Show focused pin always; also show assigned pins lightly; optionally nearby unassigned skipped to reduce clutter
    const show = [];
    for (const pin of state.gpsPins) {
      if (isPinRejected(pin)) continue;
      const assigned = isPinAssigned(pin);
      const focus = pin.path === focusPath;
      if (!focus && !assigned) continue; // only focus + saved assignments on map
      show.push({ pin, assigned, focus });
    }
    // If nothing focused yet, still allow clicking won't apply — ensure focus pin rendered
    if (focusPath) {
      const pin = state.gpsPins.find((p) => p.path === focusPath);
      if (pin && !show.some((s) => s.pin.path === pin.path)) {
        show.push({ pin, assigned: isPinAssigned(pin), focus: true });
      }
    }
    for (const { pin, assigned, focus } of show) {
      const [x, y] = pin.mapPct || [];
      if (x == null || y == null) continue;
      const el = document.createElement("div");
      el.className =
        "gps-pin" +
        (assigned ? " done" : " pending") +
        (focus ? " focus" : "");
      el.style.left = x + "%";
      el.style.top = y + "%";
      const pid = assignedPlotId(pin) || "?";
      el.title = `${pid} · ${pin.path}`;
      el.innerHTML = `<div class="gps-dot"></div><div class="gps-label">${escapeHtml(String(pid))}</div>`;
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        selectAssignPin(pin.path);
      });
      stage.appendChild(el);
    }
  }

  function scrollMapToPct(xPct, yPct) {
    const wrap = $(".map-wrap");
    const stage = $("#mapStage");
    if (!wrap || !stage) return;
    const img = stage.querySelector("img");
    const w = (img && img.clientWidth) || stage.clientWidth;
    const h = (img && img.clientHeight) || stage.clientHeight;
    const px = (xPct / 100) * w;
    const py = (yPct / 100) * h;
    wrap.scrollTo({
      left: Math.max(0, stage.offsetLeft + px - wrap.clientWidth / 2),
      top: Math.max(0, stage.offsetTop + py - wrap.clientHeight / 2),
      behavior: "smooth",
    });
  }

  function renderAssignList() {
    const box = $("#assignList");
    const meta = $("#assignMeta");
    if (!box) return;
    if (!state.assignMode) {
      box.hidden = true;
      box.innerHTML = "";
      return;
    }
    box.hidden = false;
    const all = state.gpsPins.length;
    const assigned = state.gpsPins.filter((p) => isPinAssigned(p)).length;
    const list = filteredAssignPins();
    if (meta) {
      meta.textContent = `Assign photos: ${list.length} shown · ${assigned} assigned / ${all} GPS`;
    }
    if (!list.length) {
      box.innerHTML = `<div class="photo-empty">No photos in this filter.</div>`;
      return;
    }
    const frag = document.createDocumentFragment();
    for (const pin of list) {
      const btn = document.createElement("button");
      btn.type = "button";
      const pid = assignedPlotId(pin);
      btn.className =
        "assign-row" +
        (pin.path === state.assignFocusPath ? " active" : "") +
        (pid ? " assigned" : "");
      btn.dataset.path = pin.path;
      const file = (pin.path || "").split("/").pop() || pin.path;
      const cult = pin.cultivar || (pin.path || "").split("/")[0] || "—";
      const date = formatPhotoDto(pin.dto);
      const thumb = pin.thumb ? `data/${pin.thumb}?v=30` : "";
      btn.innerHTML =
        (thumb
          ? `<img src="${escapeHtml(thumb)}" alt="" loading="lazy" decoding="async" onerror="this.style.opacity=.25">`
          : `<img alt="" style="opacity:.25">`) +
        `<div><div class="ar-cult">${escapeHtml(cult)}</div>` +
        `<div class="ar-file">${escapeHtml(file)}</div>` +
        (date ? `<div class="ar-date">${escapeHtml(date)}</div>` : "") +
        `</div>` +
        `<span class="ar-badge${pid ? "" : " empty"}">${pid ? escapeHtml(pid) : "—"}</span>`;
      btn.addEventListener("click", () => selectAssignPin(pin.path));
      frag.appendChild(btn);
    }
    box.innerHTML = "";
    box.appendChild(frag);
    // scroll active into view
    const active = box.querySelector(".assign-row.active");
    if (active) active.scrollIntoView({ block: "nearest" });
  }

  function selectAssignPin(path) {
    state.assignFocusPath = path;
    const pin = state.gpsPins.find((p) => p.path === path);
    state.photoPanelClosed = false;
    renderAssignList();
    renderGpsPins();
    updateAssignPreview(pin);
    if (pin && pin.mapPct) scrollMapToPct(pin.mapPct[0], pin.mapPct[1]);
    const layout = $(".layout");
    if (layout) layout.classList.add("assign-open", "photo-open");
  }

  function updateAssignPreview(pin) {
    const body = $("#photoPanelBody");
    const title = $("#photoPanelPlot");
    const sub = $("#photoPanelSub");
    const kicker = document.querySelector(".photo-panel-kicker");
    if (!body) return;
    if (!pin) {
      if (title) title.textContent = "—";
      if (sub) sub.textContent = "";
      body.innerHTML = `<div class="photo-empty">Pick a photo from the list to preview and assign a plot ID.</div>`;
      return;
    }
    if (kicker) kicker.textContent = "Assign photo";
    const file = (pin.path || "").split("/").pop() || pin.path;
    const cult = pin.cultivar || (pin.path || "").split("/")[0] || "—";
    if (title) title.textContent = cult;
    if (sub) sub.textContent = file + (pin.dto ? ` · ${formatPhotoDto(pin.dto)}` : "");

    const thumb = pin.thumb ? `data/${pin.thumb}?v=30` : "";
    const excel = (pin.excelPlotIds || []).filter(Boolean);
    const excelHint = excel.length
      ? `Excel inventory hint (optional): ${excel.join(", ")} — you still choose.`
      : "No Excel inventory hint for this cultivar.";
    const existing = assignedPlotId(pin) || "";
    const liveIds = liveLabeledPlotIds();
    const options = liveIds.map((id) => `<option value="${escapeHtml(id)}"></option>`).join("");

    body.innerHTML =
      `<div class="assign-preview">` +
      (thumb
        ? `<img class="ap-main" id="assignPreviewImg" src="${escapeHtml(thumb)}" alt="${escapeHtml(file)}" loading="eager" decoding="async">`
        : `<div class="photo-empty">No preview thumb packed for this file yet.<br><span style="font-size:11px">${escapeHtml(pin.path)}</span></div>`) +
      `<div class="assign-form">` +
      `<div class="af-hint">${escapeHtml(excelHint)}</div>` +
      `<label for="assignPlotInput">Plot ID</label>` +
      `<input id="assignPlotInput" type="text" list="assignPlotIds" placeholder="Type or pick a labeled plot id" maxlength="16" autocomplete="off" value="${escapeHtml(existing)}">` +
      `<datalist id="assignPlotIds">${options}</datalist>` +
      `<div class="af-warn" id="assignWarn" hidden></div>` +
      `<div class="af-actions">` +
      `<button type="button" id="btnAssignSave">Save assignment</button>` +
      `<button type="button" id="btnAssignClear" title="Remove saved assignment">Clear</button>` +
      `<button type="button" id="btnAssignNext">Next</button>` +
      `</div>` +
      `<div class="af-hint">Saved to this browser + pack file data/photo-assignments.json (survives localhost↔Tailscale). Does not auto-file Synology photos until you export.</div>` +
      `</div></div>`;

    const img = $("#assignPreviewImg");
    if (img) {
      img.addEventListener("click", () => openLightbox(img.src, file));
    }
    const saveBtn = $("#btnAssignSave");
    if (saveBtn) saveBtn.addEventListener("click", () => saveAssignFocus());
    const clearBtn = $("#btnAssignClear");
    if (clearBtn) clearBtn.addEventListener("click", () => clearAssignFocus());
    const nextBtn = $("#btnAssignNext");
    if (nextBtn) nextBtn.addEventListener("click", () => advanceAssign(1));
    const input = $("#assignPlotInput");
    if (input) {
      input.addEventListener("input", () => validateAssignInput());
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          saveAssignFocus();
        }
      });
      validateAssignInput();
      setTimeout(() => input.focus(), 30);
    }
  }

  function validateAssignInput() {
    const input = $("#assignPlotInput");
    const warn = $("#assignWarn");
    if (!input || !warn) return true;
    const v = input.value.trim();
    if (!v) {
      warn.hidden = true;
      warn.textContent = "";
      return false;
    }
    if (!knownPlotId(v)) {
      warn.hidden = false;
      warn.textContent = `Plot id “${v}” is not in current live labels / inventory. You can still save after confirm.`;
      return false;
    }
    warn.hidden = true;
    warn.textContent = "";
    return true;
  }

  function saveAssignFocus() {
    const pin = state.gpsPins.find((p) => p.path === state.assignFocusPath);
    if (!pin) return;
    const input = $("#assignPlotInput");
    const plotId = (input && input.value.trim()) || "";
    if (!plotId) {
      window.alert("Type or pick a plot ID before saving.");
      return;
    }
    if (!knownPlotId(plotId)) {
      const ok = window.confirm(
        `Plot id ${plotId} is not in current labeled outlines / inventory. Save anyway? (Still will NOT auto-file into plot-photos.json.)`
      );
      if (!ok) return;
    }
    state.photoAssignments[pin.path] = {
      plotId,
      status: "assigned",
      at: new Date().toISOString(),
      cultivar: pin.cultivar || null,
      thumb: pin.thumb || null,
      originalPath: pin.path,
    };
    savePhotoAssignments();
    // Optional in-session: surface under plot photos panel for that id (memory only)
    try {
      if (!state.plotPhotos[plotId]) state.plotPhotos[plotId] = { cultivar: pin.cultivar || "", photos: [] };
      const rel = pin.thumb || null;
      if (rel && Array.isArray(state.plotPhotos[plotId].photos)) {
        if (!state.plotPhotos[plotId].photos.includes(rel)) {
          state.plotPhotos[plotId].photos = [rel, ...state.plotPhotos[plotId].photos];
        }
      }
    } catch (_) {}
    renderAssignList();
    renderGpsPins();
    advanceAssign(1);
  }

  function clearAssignFocus() {
    const path = state.assignFocusPath;
    if (!path) return;
    if (state.photoAssignments[path]) {
      delete state.photoAssignments[path];
      savePhotoAssignments();
    }
    renderAssignList();
    renderGpsPins();
    const pin = state.gpsPins.find((p) => p.path === path);
    updateAssignPreview(pin);
  }

  function advanceAssign(dir) {
    const list = filteredAssignPins();
    if (!list.length) {
      updateAssignPreview(null);
      return;
    }
    let idx = list.findIndex((p) => p.path === state.assignFocusPath);
    if (idx < 0) idx = 0;
    else idx = (idx + dir + list.length) % list.length;
    selectAssignPin(list[idx].path);
  }

  function setAssignMode(on) {
    state.assignMode = !!on;
    state.gpsConfirmMode = state.assignMode; // legacy
    const bar = $("#assignBar");
    const btn = $("#btnAssignPhotos");
    const list = $("#assignList");
    if (bar) bar.hidden = !on;
    if (btn) btn.setAttribute("aria-pressed", on ? "true" : "false");
    document.body.classList.toggle("assign-mode", on);
    document.body.classList.toggle("gps-confirm-mode", false);
    const layout = $(".layout");
    if (on) {
      if (list) list.hidden = false;
      if (layout) layout.classList.add("assign-open", "photo-open");
      state.photoPanelClosed = false;
      // turn off other modes that fight the sidebar
      if (state.labelMode) {
        state.labelMode = false;
        const lb = $("#btnLabelMode");
        if (lb) lb.setAttribute("aria-pressed", "false");
        const lbar = $("#labelBar");
        if (lbar) lbar.hidden = true;
      }
      renderAssignList();
      if (!state.assignFocusPath && state.gpsPins.length) {
        const first = filteredAssignPins()[0] || state.gpsPins[0];
        if (first) selectAssignPin(first.path);
      } else if (state.assignFocusPath) {
        selectAssignPin(state.assignFocusPath);
      } else {
        updateAssignPreview(null);
      }
    } else {
      if (list) {
        list.hidden = true;
        list.innerHTML = "";
      }
      if (layout) layout.classList.remove("assign-open");
      const kicker = document.querySelector(".photo-panel-kicker");
      if (kicker) kicker.textContent = "Photos";
      // restore normal photo panel for selected plot
      updatePhotoPanel();
    }
    renderGpsPins();
  }

  async function load() {
    loadLocalLabels();
    loadCustomOutlines();
    loadDeletedAutoIndices();
    loadGpsDecisions();
    loadPhotoAssignments();
    const [flat, coords, polys, photos, gpsPinsDoc, reviewDoc] = await Promise.all([
      fetch("data/stock-flat.json?v=30").then((r) => r.json()),
      fetch("data/plot-coords.json?v=30").then((r) => r.json()),
      fetch("data/plot-polygons.json?v=30")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null),
      fetch("data/plot-photos.json?v=30")
        .then((r) => (r.ok ? r.json() : {}))
        .catch(() => ({})),
      fetch("data/photo-gps-pins.json?v=30")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null),
      fetch("data/gps-review-manifest.json?v=30")
        .then((r) => (r.ok ? r.json() : null))
        .catch(() => null),
    ]);
    await loadPhotoAssignmentsFromPack();
    // Pack merge BEFORE normalize + polygon filter so Tailscale/LAN get localhost labels.
    await loadLabelsFromPack();
    await loadCustomOutlinesFromPack();
    await loadDeletedAutoFromPack();
    normalizeIdsOnceV1();
    if (reviewDoc) window.__gotDirtReviewManifest = reviewDoc;
    state.plotPhotos = photos || {};
    state.gpsPins = (gpsPinsDoc && gpsPinsDoc.pins) || [];
    // Attach packed review thumbs when present on pin or via companion manifest
    for (const pin of state.gpsPins) {
      if (!pin.thumb && pin.thumbFile) {
        pin.thumb = `plot-photos/_gps-review/${pin.thumbFile}`;
      }
    }
    // Prefer dedicated review list if shipped
    if (window.__gotDirtReviewManifest && Array.isArray(window.__gotDirtReviewManifest.items)) {
      const byPath = new Map(window.__gotDirtReviewManifest.items.map((it) => [it.path, it]));
      for (const pin of state.gpsPins) {
        const it = byPath.get(pin.path);
        if (it && it.thumb) pin.thumb = it.thumb;
      }
    }
    state.plots = (flat.plots || []).map((p) => ({
      ...p,
      search: `${p.id} ${p.plants || ""}`.toLowerCase(),
    }));
    state.coords = coords.plots || {};
    state.calibration = (polys && polys.calibration) || coords.calibration || "unknown";
    state.generatedAt = flat.generatedAt || "";
    // Map stable index first, then filter deleted — do NOT renumber remaining features
    state.polygons = ((polys && polys.features) || [])
      .map((f, index) => ({ ...f, index }))
      .filter((f) => !state.deletedAutoIndices.has(f.index));
    // Full-block autos: show as subtle dim strokes; Label mode paints unlabeled.
    // Hide checkbox still available if Aj wants plain photo only.
    state.hideAutoOutlines = false;
    state.selected = new Set();
    state.selectedCustomIdx = null;
    state.selectedAutoIndex = null;
    state.labelFocusIndex = null;
    const chkHide = $("#chkHideAuto");
    if (chkHide) chkHide.checked = false;
    rebuildPolyIndex();
    $("#statPlots").textContent = String(state.plots.length);
    $("#statSheet").textContent = flat.sourceSheet || "25-26";
    $("#statGen").textContent = formatCt(state.generatedAt);
    const calLabel =
      state.calibration === "partial"
        ? "PARTIAL (outlines)"
        : state.calibration === "needs-redraw"
          ? "NEEDS REDRAW"
          : state.calibration === "approximate"
            ? "APPROXIMATE"
            : String(state.calibration).toUpperCase();
    $("#statCal").textContent = calLabel;
    $("#statCal").style.color =
      state.calibration === "approximate" ||
      state.calibration === "partial" ||
      state.calibration === "needs-redraw"
        ? "var(--accent)"
        : "var(--green)";
    renderResults();
    renderPolygons();
    updateMapStatus();
    updatePhotoPanel();
    renderGpsPins();
  }


  // Show my location — browser GPS only, Farm 1 georef, nothing stored or sent.
  let myLocWatchId = null;
  let farm1Fit = null;
  let farm1FitPromise = null;
  let myLocPending = null;
  let myLocScrolled = false;

  function smoothstep01(edge0, edge1, x) {
    if (edge1 === edge0) return x < edge0 ? 0 : 1;
    let t = (x - edge0) / (edge1 - edge0);
    if (t < 0) t = 0;
    else if (t > 1) t = 1;
    return t * t * (3 - 2 * t);
  }

  function projectFarm1(lat, lon, fit) {
    const enu = fit.enuOrigin;
    const inv = fit.inverse_EN_to_map_percent;
    const Bw = inv.B2inv_2x2;
    const Be = inv.east_B2inv_2x2;
    const tw = inv.west_t;
    const te = inv.east_t;
    const split = Number(fit.piecewise.splitXPct);
    const blend = Number(fit.piecewise.blendPct);
    const E = (lon - enu.lon) * enu.mPerDegLon;
    const N = (lat - enu.lat) * enu.mPerDegLat;
    const mv = (M, v0, v1) => [v0 * M[0][0] + v1 * M[1][0], v0 * M[0][1] + v1 * M[1][1]];
    const xyOf = (B, t) => mv(B, E - t[0], N - t[1]);
    let xy = xyOf(Bw, tw);
    const lo = split - blend / 2;
    const hi = split + blend / 2;
    for (let i = 0; i < 8; i++) {
      const w = 1 - smoothstep01(lo, hi, xy[0]);
      const a = xyOf(Bw, tw);
      const b = xyOf(Be, te);
      xy = [w * a[0] + (1 - w) * b[0], w * a[1] + (1 - w) * b[1]];
    }
    const rel = fit.reliableRegionMapPct || {};
    const xr = rel.xPct || [0, 100];
    const yr = rel.yPct || [0, 100];
    // 2% pad: in-sample edge residual is ~9 m (~1%). Keeps a phone on the fit edge.
    // Points outside the Farm 1 fit still project far outside this box.
    const pad = 2;
    const inside =
      xy[0] >= xr[0] - pad && xy[0] <= xr[1] + pad && xy[1] >= yr[0] - pad && xy[1] <= yr[1] + pad;
    return { x: xy[0], y: xy[1], inside };
  }

  function setMyLocStatus(text, kind) {
    const el = $("#myLocStatus");
    const btn = $("#btnMyLocation");
    if (btn) btn.title = text || "Show this phone on the Farm 1 map. GPS stays in the browser and is not saved.";
    if (!el) return;
    if (!text) {
      el.hidden = true;
      el.textContent = "";
      el.removeAttribute("data-kind");
      return;
    }
    el.hidden = false;
    el.dataset.kind = kind || "";
    el.textContent = text;
  }

  function hideMyLocDot() {
    const dot = $("#myLocDot");
    if (!dot) return;
    dot.hidden = true;
  }

  function showMyLocDot(x, y) {
    const dot = $("#myLocDot");
    if (!dot) return;
    dot.hidden = false;
    dot.style.left = x + "%";
    dot.style.top = y + "%";
    if (!myLocScrolled) {
      myLocScrolled = true;
      try {
        dot.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
      } catch (_) {}
    }
  }

  function ensureFarm1Fit() {
    if (farm1Fit) return Promise.resolve(farm1Fit);
    if (farm1FitPromise) return farm1FitPromise;
    farm1FitPromise = fetch("data/farm1-georef.json?v=30")
      .then((r) => {
        if (!r.ok) throw new Error("missing");
        return r.json();
      })
      .then((doc) => {
        if (!doc || doc.farm !== 1 || !doc.inverse_EN_to_map_percent || !doc.enuOrigin) {
          throw new Error("bad fit");
        }
        farm1Fit = doc;
        return doc;
      })
      .catch((err) => {
        farm1FitPromise = null;
        throw err;
      });
    return farm1FitPromise;
  }

  function applyMyLocation(pos) {
    if (!farm1Fit || !pos || !pos.coords) return;
    const lat = pos.coords.latitude;
    const lon = pos.coords.longitude;
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    const proj = projectFarm1(lat, lon, farm1Fit);
    if (!proj.inside) {
      hideMyLocDot();
      setMyLocStatus("Outside Farm 1 map", "out");
      return;
    }
    showMyLocDot(proj.x, proj.y);
    const acc = pos.coords.accuracy;
    const accTxt = Number.isFinite(acc) ? " · \u00b1" + Math.round(acc) + " m" : "";
    setMyLocStatus("On Farm 1" + accTxt, "on");
  }

  function stopMyLocation() {
    if (myLocWatchId != null && navigator.geolocation) {
      navigator.geolocation.clearWatch(myLocWatchId);
    }
    myLocWatchId = null;
    myLocPending = null;
    myLocScrolled = false;
    const btn = $("#btnMyLocation");
    if (btn) {
      btn.setAttribute("aria-pressed", "false");
      btn.textContent = "Show my location";
    }
    hideMyLocDot();
    setMyLocStatus("");
  }

  function onMyLocPosition(pos) {
    myLocPending = pos;
    if (!farm1Fit) return;
    applyMyLocation(pos);
  }

  function onMyLocError(err) {
    hideMyLocDot();
    let msg = "Location unavailable";
    const code = err && err.code;
    if (code === 1) msg = "Location blocked";
    else if (code === 3) msg = "Location timed out";
    if (code === 1) {
      stopMyLocation();
      setMyLocStatus(msg, "err");
      return;
    }
    setMyLocStatus(msg, "err");
  }

  function startMyLocation() {
    const btn = $("#btnMyLocation");
    if (!navigator.geolocation) {
      setMyLocStatus("GPS not available", "err");
      return;
    }
    if (btn) {
      btn.setAttribute("aria-pressed", "true");
      btn.textContent = "Show my location";
    }
    setMyLocStatus("Locating\u2026", "wait");
    myLocWatchId = navigator.geolocation.watchPosition(onMyLocPosition, onMyLocError, {
      enableHighAccuracy: true,
      maximumAge: 1000,
      timeout: 20000,
    });
    ensureFarm1Fit()
      .then(() => {
        if (myLocWatchId == null) return;
        if (myLocPending) applyMyLocation(myLocPending);
      })
      .catch(() => {
        if (myLocWatchId == null) return;
        stopMyLocation();
        setMyLocStatus("Farm 1 map fit missing", "err");
      });
  }

  function toggleMyLocation() {
    const btn = $("#btnMyLocation");
    const on = btn && btn.getAttribute("aria-pressed") === "true";
    if (on || myLocWatchId != null) stopMyLocation();
    else startMyLocation();
  }

  function bind() {
    const btnPhotoClose = $("#btnPhotoClose");
    if (btnPhotoClose) {
      btnPhotoClose.addEventListener("click", () => {
        state.photoPanelClosed = true;
        updatePhotoPanel();
        closeLightbox();
      });
    }

    const btnAssign = $("#btnAssignPhotos");
    if (btnAssign) {
      btnAssign.addEventListener("click", () => setAssignMode(!state.assignMode));
    }
    const chkGps = $("#chkShowGpsPins");
    if (chkGps) {
      chkGps.addEventListener("change", () => {
        state.gpsShowPins = !!chkGps.checked;
        renderGpsPins();
      });
    }
    const assignFilter = $("#assignFilter");
    if (assignFilter) {
      assignFilter.addEventListener("change", () => {
        state.assignFilter = assignFilter.value || "unassigned";
        renderAssignList();
      });
    }
    const assignSearch = $("#assignSearch");
    if (assignSearch) {
      let t = null;
      assignSearch.addEventListener("input", () => {
        clearTimeout(t);
        t = setTimeout(() => {
          state.assignQuery = assignSearch.value || "";
          renderAssignList();
        }, 80);
      });
    }
    const btnLb = $("#btnLightboxClose");
    if (btnLb) btnLb.addEventListener("click", closeLightbox);
    const lb = $("#lightbox");
    if (lb) {
      lb.addEventListener("click", (e) => {
        if (e.target === lb) closeLightbox();
      });
    }
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") closeLightbox();
    });

    const input = $("#q");
    let t = null;
    input.addEventListener("input", () => {
      clearTimeout(t);
      t = setTimeout(() => search(input.value), 80);
    });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        input.value = "";
        search("");
      }
    });
    $("#btnClear").addEventListener("click", () => {
      input.value = "";
      search("");
      input.focus();
    });
    $("#btnFit").addEventListener("click", () => {
      $(".map-wrap").scrollTo({ left: 0, top: 0, behavior: "smooth" });
    });

    const btnMyLocation = $("#btnMyLocation");
    if (btnMyLocation) btnMyLocation.addEventListener("click", toggleMyLocation);

    const clearLabelsBtn = $("#btnClearLabels");
    if (clearLabelsBtn) clearLabelsBtn.addEventListener("click", clearAllLabels);

    const outlineBtn = $("#btnOutlineMode");
    if (outlineBtn) {
      outlineBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        setOutlineMode(!state.outlineMode);
      });
    }

    const labelBtn = $("#btnLabelMode");
    if (labelBtn) {
      labelBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        setLabelMode(!state.labelMode);
      });
    }
    window.__gotDirtToggleLabelMode = () => setLabelMode(!state.labelMode);
    window.__gotDirtToggleOutlineMode = () => setOutlineMode(!state.outlineMode);

    const copyBtn = $("#btnCopyLabels");
    if (copyBtn) copyBtn.addEventListener("click", copyLabelsJson);
    const applyBtn = $("#btnLabelApply");
    if (applyBtn) applyBtn.addEventListener("click", applyLabel);
    const prevBtn = $("#btnLabelPrev");
    if (prevBtn) prevBtn.addEventListener("click", () => focusNextUnlabeled(-1));
    const nextBtn = $("#btnLabelNext");
    if (nextBtn) nextBtn.addEventListener("click", () => focusNextUnlabeled(1));
    const skipBtn = $("#btnLabelSkip");
    if (skipBtn) skipBtn.addEventListener("click", () => focusNextUnlabeled(1));
    const clearPolyBtn = $("#btnLabelClearPoly");
    if (clearPolyBtn)
      clearPolyBtn.addEventListener("click", () => {
        if (state.labelFocusIndex == null) return;
        state.localLabels[String(state.labelFocusIndex)] = "";
        saveLocalLabels();
        rebuildPolyIndex();
        $("#labelInput").value = "";
        renderPolygons();
        renderPins();
        renderResults();
        renderLabelQueue();
        focusNextUnlabeled(1);
      });
    const labelInputEl = $("#labelInput");
    if (labelInputEl)
      labelInputEl.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          applyLabel();
        }
      });

    // Outline help controls
    const btnDone = $("#btnOutlineDone");
    if (btnDone) btnDone.addEventListener("click", () => closeDrawPolygon());
    const btnCancel = $("#btnOutlineCancel");
    if (btnCancel) btnCancel.addEventListener("click", cancelDraw);
    const btnDel = $("#btnOutlineDelete");
    if (btnDel) btnDel.addEventListener("click", deleteSelectedOutline);
    const btnSave = $("#btnOutlineSave");
    if (btnSave) btnSave.addEventListener("click", saveOutline);
    const btnCopyOut = $("#btnCopyOutlines");
    if (btnCopyOut) btnCopyOut.addEventListener("click", copyOutlinesJson);
    const btnClearOut = $("#btnClearOutlines");
    if (btnClearOut) btnClearOut.addEventListener("click", clearMyOutlines);
    const btnClearAuto = $("#btnClearAutoOutlines");
    if (btnClearAuto) btnClearAuto.addEventListener("click", clearAllAutoOutlines);
    const btnLabelDelPoly = $("#btnLabelDeletePoly");
    if (btnLabelDelPoly)
      btnLabelDelPoly.addEventListener("click", () => {
        if (state.labelFocusIndex == null) return;
        const deleted = deleteAutoOutline(state.labelFocusIndex, { confirmLabeled: true });
        if (deleted) {
          const input = $("#labelInput");
          if (input) input.value = "";
          focusNextUnlabeled(1);
        }
      });
    const chkHide = $("#chkHideAuto");
    if (chkHide)
      chkHide.addEventListener("change", () => {
        state.hideAutoOutlines = !!chkHide.checked;
        rebuildPolyIndex();
        renderPolygons();
        renderPins();
        updateMapStatus();
        updateOutlineMeta();
      });
    const outlineIdEl = $("#outlineIdInput");
    if (outlineIdEl)
      outlineIdEl.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          saveOutline();
        }
      });

    // Map drawing clicks
    const stage = $("#mapStage") || $(".map-stage");
    if (stage) {
      stage.addEventListener("click", (e) => {
        if (!state.outlineMode) return;
        // Ignore clicks on polys (handled separately for select)
        if (e.target && e.target.closest && e.target.closest("path.custom-poly, path.plot-poly")) return;
        e.preventDefault();
        const pct = stagePctFromEvent(e);
        if (!pct) return;
        // Starting a new draw clears pending/selection
        if (state.pendingPath) {
          state.pendingPath = null;
        }
        state.selectedCustomIdx = null;
        state.selectedAutoIndex = null;
        state.drawPoints.push(pct);
        renderDrawLayer();
        updateOutlineMeta();
      });
      stage.addEventListener("dblclick", (e) => {
        if (!state.outlineMode) return;
        e.preventDefault();
        e.stopPropagation();
        // Remove accidental duplicate point from second click of dblclick
        if (state.drawPoints.length >= 2) {
          const a = state.drawPoints[state.drawPoints.length - 1];
          const b = state.drawPoints[state.drawPoints.length - 2];
          if (a && b && Math.abs(a[0] - b[0]) < 0.15 && Math.abs(a[1] - b[1]) < 0.15) {
            state.drawPoints.pop();
          }
        }
        closeDrawPolygon();
      });
    }

    document.querySelectorAll("[data-demo]").forEach((btn) => {
      btn.addEventListener("click", () => {
        input.value = btn.getAttribute("data-demo");
        search(input.value);
        input.focus();
      });
    });
  }

  bind();
  load().catch((err) => {
    $("#results").innerHTML = `<div class="empty" style="color:var(--red)">Failed to load data: ${escapeHtml(
      String(err)
    )}<br>Serve from got-dirt root or /app/ with data/ present.</div>`;
  });
})();

/* bvr-plan.js — общая логика проекта блока для панели БВР и формы буровика.
   Без обращений к DOM: только расчёты и строка SVG. Подключается через <script src="bvr-plan.js">.
   Источник данных: CSV проекта на бурение (колонки HOLE, BLOCK, EAST_START … HOLE_TYPE) и ответ
   бэкенда getDrillPlan (план блока + факт из «БВР-Смены»). */
var BP = (function () {
  "use strict";

  function num(v) { var n = Number(String(v == null ? "" : v).replace(",", ".").trim()); return isFinite(n) ? n : 0; }
  function r2(x) { return Math.round(x * 100) / 100; }
  function isNum(s) { return s !== "" && s != null && isFinite(Number(String(s).replace(",", "."))); }

  // ── CSV ──────────────────────────────────────────────────────────
  // Разбор с учётом кавычек; разделитель определяется по первой строке (запятая или точка с запятой).
  function splitCsv(text) {
    text = String(text || "").replace(/^﻿/, "");
    var first = text.split(/\r?\n/)[0] || "";
    var sep = (first.split(";").length > first.split(",").length) ? ";" : ",";
    var rows = [], row = [], cur = "", q = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c;
      } else if (c === '"') q = true;
      else if (c === sep) { row.push(cur); cur = ""; }
      else if (c === "\n" || c === "\r") { if (c === "\r" && text[i + 1] === "\n") i++; row.push(cur); rows.push(row); row = []; cur = ""; }
      else cur += c;
    }
    if (cur !== "" || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }

  var COLS = ["HOLE", "BLOCK", "EAST_START", "NORTH_START", "RL_START", "EAST_END", "NORTH_END", "RL_END", "DIP", "AZIMUTH", "DEPTH", "HOLE DIAM", "SUBDRILL", "HOLE_TYPE"];

  // Возвращает { ok, error, block, holes[], summary{}, warnings[] }.
  // HOLE_TYPE в проекте — отклонение скважины от вертикали, градусы (подтверждено мастером: при угле -60 стоит 30).
  function parsePlanCsv(text) {
    var rows = splitCsv(text);
    if (!rows.length) return { ok: false, error: "Файл пустой" };
    var head = rows[0].map(function (h) { return String(h).trim().toUpperCase(); });
    var ix = {}, missing = [];
    COLS.forEach(function (c) { var i = head.indexOf(c); if (i < 0) missing.push(c); else ix[c] = i; });
    if (missing.length) return { ok: false, error: "В файле нет колонок: " + missing.join(", ") + ". Ожидается CSV проекта на бурение." };
    var holes = [], blocks = {}, tail = [], seen = {}, warnings = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (r.every(function (x) { return String(x).trim() === ""; })) continue;
      var hn = String(r[ix.HOLE] || "").trim();
      if (!hn) {                                   // итоговые строки внизу файла (без номера скважины)
        if (isNum(r[ix.DEPTH])) tail.push({ depth: num(r[ix.DEPTH]), diam: num(r[ix["HOLE DIAM"]]) });
        continue;
      }
      if (!/^\d+$/.test(hn)) return { ok: false, error: "Строка " + (i + 1) + ": номер скважины «" + hn + "» не число" };
      var n = parseInt(hn, 10);
      if (seen[n]) return { ok: false, error: "Номер скважины " + n + " повторяется (строка " + (i + 1) + ")" };
      seen[n] = 1;
      var blk = String(r[ix.BLOCK] || "").trim(); blocks[blk] = 1;
      var h = {
        hole: n, eastStart: num(r[ix.EAST_START]), northStart: num(r[ix.NORTH_START]), rlStart: num(r[ix.RL_START]),
        eastEnd: num(r[ix.EAST_END]), northEnd: num(r[ix.NORTH_END]), rlEnd: num(r[ix.RL_END]),
        dip: num(r[ix.DIP]), azimuth: num(r[ix.AZIMUTH]), depth: num(r[ix.DEPTH]), diameter: num(r[ix["HOLE DIAM"]]),
        subdrill: num(r[ix.SUBDRILL]), deviation: num(r[ix.HOLE_TYPE]), row: 0
      };
      if (!(h.depth > 0 && h.depth <= 100)) return { ok: false, error: "Скважина №" + n + ": глубина " + h.depth + " вне диапазона 0–100 м" };
      if (!(h.diameter > 0)) return { ok: false, error: "Скважина №" + n + ": не указан диаметр" };
      holes.push(h);
    }
    if (!holes.length) return { ok: false, error: "В файле нет скважин" };
    var bl = Object.keys(blocks);
    if (bl.length !== 1 || !bl[0]) return { ok: false, error: "В файле должен быть один блок, найдено: " + (bl.join(", ") || "пусто") };
    holes.sort(function (a, b) { return a.hole - b.hole; });
    var total = r2(holes.reduce(function (s, h) { return s + h.depth; }, 0));
    var summary = { holes: holes.length, planMeters: total, avgDepth: r2(total / holes.length), area: 0, volume: 0 };
    // Итоговые строки внизу CSV: средняя глубина, сумма п.м., объём (+ выход м3/м), площадь.
    // Берём объём и площадь, только если первые две строки сходятся с нашим расчётом.
    if (tail.length >= 4) {
      var okAvg = Math.abs(tail[0].depth - summary.avgDepth) <= 0.01, okTot = Math.abs(tail[1].depth - total) <= 0.01;
      if (okAvg && okTot) { summary.volume = tail[2].depth; summary.area = tail[3].depth; }
      else warnings.push("Итоговые строки в конце CSV не сходятся с суммой по скважинам (в файле: средняя " + tail[0].depth + ", итого " + tail[1].depth + "; по скважинам: " + summary.avgDepth + ", " + total + "). Объём и площадь из файла не взяты.");
    }
    var asg = assignRows(holes);
    return { ok: true, block: bl[0], holes: holes, summary: summary, rowsMethod: asg.method, rowsCount: asg.count, warnings: warnings };
  }

  // ── номер ряда ───────────────────────────────────────────────────
  // В CSV колонки «ряд» нет. Правило по нумерации проекта: ряд открывает серия наклонных скважин,
  // за ней идут вертикальные того же ряда (так на блоке 3780316 получается 7 рядов — как в параметрах PDF).
  // Если наклонных серий меньше двух (блок целиком вертикальный) — ряды по геометрии:
  // скважины группируются по расстоянию поперёк ряда.
  // ВЫВОД, не данные проекта: распределение показывается в предпросмотре для сверки со схемой.
  function isIncl(h) { return Math.abs(Math.abs(h.dip) - 90) > 0.5; }
  function assignRows(holes) {
    var row = 0, prev = false, runs = 0;
    holes.forEach(function (h) {
      var inc = isIncl(h);
      if (inc && !prev) { row++; runs++; } else if (row === 0) row = 1;
      h.row = row; prev = inc;
    });
    if (runs >= 2) return { method: "нумерация", count: row };
    // геометрия
    var v = holes.filter(function (h) { return !isIncl(h); });
    if (v.length < 4) { holes.forEach(function (h) { h.row = 1; }); return { method: "один ряд", count: 1 }; }
    var dx = v[1].eastStart - v[0].eastStart, dy = v[1].northStart - v[0].northStart, L = Math.sqrt(dx * dx + dy * dy) || 1;
    var ux = dx / L, uy = dy / L;
    var items = holes.map(function (h) { return { h: h, p: -uy * h.eastStart + ux * h.northStart }; }).sort(function (a, b) { return a.p - b.p; });
    var cnt = 1, startP = items[0].p; items[0].h.row = 1;
    for (var i = 1; i < items.length; i++) { if (items[i].p - items[i - 1].p > 1.5) cnt++; items[i].h.row = cnt; }
    return { method: "геометрия", count: cnt };
  }

  // Распределение по рядам для предпросмотра: [{row, count, meters, from, to}]
  function rowSummary(holes) {
    var m = {};
    holes.forEach(function (h) {
      var o = m[h.row] || (m[h.row] = { row: h.row, count: 0, meters: 0, nums: [] });
      o.count++; o.meters += num(h.depth); o.nums.push(Number(h.hole));
    });
    return Object.keys(m).map(Number).sort(function (a, b) { return a - b; }).map(function (k) {
      var o = m[k]; o.meters = r2(o.meters); o.from = Math.min.apply(null, o.nums); o.to = Math.max.apply(null, o.nums); delete o.nums; return o;
    });
  }

  // ── статус скважины по ответу getDrillPlan ───────────────────────
  // todo — не бурилась; ok — пробурена; replaced — разрушилась и заменена (перебур)
  function statusOf(h) {
    if (h.replacedBy) return "replaced";
    return h.done ? "ok" : "todo";
  }

  // ── статистика блока ─────────────────────────────────────────────
  // res — ответ getDrillPlan. today — "ГГГГ-ММ-ДД" для прогноза.
  // Процент выполнения считается по ПРОЕКТНЫМ метрам пробуренных скважин, чтобы перебур и отклонение глубины
  // не искажали его; фактические метры показываются отдельно.
  function planStats(res, today) {
    var holes = res.holes || [], redrills = res.redrills || [], extra = res.extra || [];
    var S = { planHoles: holes.length, planMeters: 0, doneHoles: 0, donePlanMeters: 0, factMeters: 0, remainHoles: 0, remainMeters: 0,
      replacedHoles: 0, redrillCount: redrills.length, redrillMeters: 0, extraCount: extra.length, extraMeters: 0,
      pctHoles: 0, pctMeters: 0, byDiam: [], byRow: [], byRig: [], byEmp: [], deviations: [], repeated: [], inclined: { holes: 0, done: 0, meters: 0 },
      firstDate: "", lastDate: "", activeDays: 0, pace: 0, daysLeft: 0, forecast: "" };
    var diam = {}, rows = {}, rigs = {}, emps = {}, days = {};
    function bump(map, key, plan, fact, isDone) { var o = map[key] || (map[key] = { key: key, holes: 0, done: 0, planMeters: 0, donePlanMeters: 0, factMeters: 0 }); o.holes += plan ? 1 : 0; o.planMeters += plan; o.done += isDone ? 1 : 0; o.donePlanMeters += isDone ? plan : 0; o.factMeters += fact; return o; }
    holes.forEach(function (h) {
      var pd = num(h.depth), done = !!h.done, fd = done ? num(h.factDepth) : 0;
      S.planMeters += pd;
      if (done) { S.doneHoles++; S.donePlanMeters += pd; S.factMeters += fd; } else { S.remainHoles++; S.remainMeters += pd; }
      if (h.replacedBy) S.replacedHoles++;
      bump(diam, h.diameter, pd, fd, done); bump(rows, h.row, pd, fd, done);
      if (isIncl(h)) { S.inclined.holes++; if (done) S.inclined.done++; S.inclined.meters += pd; }
      if (done) {
        var d = r2(fd - pd);
        if (Math.abs(d) >= 0.05) S.deviations.push({ hole: h.hole, row: h.row, plan: pd, fact: fd, diff: d, date: h.date, rig: h.rig });
        if (Number(h.count) > 1) S.repeated.push({ hole: h.hole, count: h.count });
        if (h.date) days[h.date] = 1;
        var rg = rigs[h.rig || "—"] || (rigs[h.rig || "—"] = { key: h.rig || "—", holes: 0, meters: 0 }); rg.holes++; rg.meters += fd;
        var em = emps[h.empName || "—"] || (emps[h.empName || "—"] = { key: h.empName || "—", holes: 0, meters: 0 }); em.holes++; em.meters += fd;
      }
    });
    redrills.forEach(function (x) {
      S.redrillMeters += num(x.depth); if (x.date) days[x.date] = 1;
      var rg = rigs[x.rig || "—"] || (rigs[x.rig || "—"] = { key: x.rig || "—", holes: 0, meters: 0 }); rg.meters += num(x.depth);
      var em = emps[x.empName || "—"] || (emps[x.empName || "—"] = { key: x.empName || "—", holes: 0, meters: 0 }); em.meters += num(x.depth);
    });
    extra.forEach(function (x) { S.extraMeters += num(x.depth); if (x.date) days[x.date] = 1; });
    S.planMeters = r2(S.planMeters); S.donePlanMeters = r2(S.donePlanMeters); S.factMeters = r2(S.factMeters);
    S.remainMeters = r2(S.remainMeters); S.redrillMeters = r2(S.redrillMeters); S.extraMeters = r2(S.extraMeters);
    S.inclined.meters = r2(S.inclined.meters);
    S.pctHoles = S.planHoles ? Math.round(S.doneHoles / S.planHoles * 1000) / 10 : 0;
    S.pctMeters = S.planMeters ? Math.round(S.donePlanMeters / S.planMeters * 1000) / 10 : 0;
    function fin(o) { o.planMeters = r2(o.planMeters); o.donePlanMeters = r2(o.donePlanMeters); o.factMeters = r2(o.factMeters); return o; }
    S.byDiam = Object.keys(diam).map(Number).sort(function (a, b) { return a - b; }).map(function (k) { return fin(diam[k]); });
    S.byRow = Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).map(function (k) { return fin(rows[k]); });
    S.byRig = Object.keys(rigs).sort().map(function (k) { rigs[k].meters = r2(rigs[k].meters); return rigs[k]; });
    S.byEmp = Object.keys(emps).sort().map(function (k) { emps[k].meters = r2(emps[k].meters); return emps[k]; });
    var ds = Object.keys(days).sort();
    S.activeDays = ds.length; S.firstDate = ds[0] || ""; S.lastDate = ds[ds.length - 1] || "";
    // Темп — средние метры за сутки, в которые на блоке бурили (фактические + перебур).
    var totalFact = S.factMeters + S.redrillMeters + S.extraMeters;
    if (S.activeDays > 0 && totalFact > 0) {
      S.pace = r2(totalFact / S.activeDays);
      if (S.remainMeters > 0 && S.pace > 0) {
        S.daysLeft = Math.ceil(S.remainMeters / S.pace);
        if (today) { var d0 = new Date(today + "T00:00:00Z"); d0.setUTCDate(d0.getUTCDate() + S.daysLeft); S.forecast = d0.toISOString().slice(0, 10); }
      }
    }
    return S;
  }

  // ── схема блока (SVG-строка) ─────────────────────────────────────
  // holes — план блока (координаты в метрах). opts: status — функция h -> "todo|ok|replaced",
  // current — номер выбранной скважины, labels — подписывать номера, sel — делать скважины нажимаемыми (data-h).
  var COLORS = { todo: "#94A3B8", ok: "#16A34A", replaced: "#DC2626" };
  function planSvg(holes, opts) {
    opts = opts || {};
    if (!holes || !holes.length) return "";
    var minE = 1e18, maxE = -1e18, minN = 1e18, maxN = -1e18;
    holes.forEach(function (h) {
      [h.eastStart, h.eastEnd].forEach(function (e) { if (e < minE) minE = e; if (e > maxE) maxE = e; });
      [h.northStart, h.northEnd].forEach(function (n) { if (n < minN) minN = n; if (n > maxN) maxN = n; });
    });
    var pad = 4, W = maxE - minE + 2 * pad, H = maxN - minN + 2 * pad;
    function X(e) { return r2(e - minE + pad); }
    function Y(n) { return r2(maxN - n + pad); }
    var st = opts.status || statusOf, cur = opts.current == null ? "" : String(opts.current);
    var out = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + r2(W) + " " + r2(H) + '" style="width:100%;height:auto;display:block;background:#F8FAFC;border-radius:8px;' + (opts.sel ? "touch-action:manipulation;" : "") + '">';
    holes.forEach(function (h) {
      var s = st(h), col = COLORS[s] || COLORS.todo, hasLine = (Math.abs(h.eastEnd - h.eastStart) + Math.abs(h.northEnd - h.northStart)) > 0.05;
      var isCur = String(h.hole) === cur;
      out += '<g' + (opts.sel ? ' data-h="' + h.hole + '" style="cursor:pointer"' : "") + ">";
      if (hasLine) out += '<line x1="' + X(h.eastStart) + '" y1="' + Y(h.northStart) + '" x2="' + X(h.eastEnd) + '" y2="' + Y(h.northEnd) + '" stroke="' + col + '" stroke-width="0.35" opacity="0.7"/>';
      if (opts.sel) out += '<circle cx="' + X(h.eastStart) + '" cy="' + Y(h.northStart) + '" r="2.1" fill="transparent"/>';   // увеличенная область нажатия
      if (isCur) out += '<circle cx="' + X(h.eastStart) + '" cy="' + Y(h.northStart) + '" r="2.3" fill="none" stroke="#F4A52A" stroke-width="0.6"/>';
      out += '<circle cx="' + X(h.eastStart) + '" cy="' + Y(h.northStart) + '" r="1.0" fill="' + col + '"/>';
      if (opts.labels) out += '<text x="' + (X(h.eastStart) + 1.3) + '" y="' + (Y(h.northStart) - 1.0) + '" font-size="1.7" fill="#0B1A35" font-family="sans-serif">' + h.hole + "</text>";
      out += "</g>";
    });
    return out + "</svg>";
  }

  return { parsePlanCsv: parsePlanCsv, assignRows: assignRows, rowSummary: rowSummary, statusOf: statusOf, planStats: planStats, planSvg: planSvg, isIncl: isIncl, COLORS: COLORS };
})();
if (typeof module !== "undefined") module.exports = BP;

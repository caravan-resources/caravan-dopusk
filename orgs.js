// ═══════════════════════════════════════════════════════════════
// СПРАВОЧНИК ОРГАНИЗАЦИЙ — общий для anketa.html, instructor.html, training.html
//
// Одно место, где задаётся, какие организации есть и как узнавать
// подрядчика по разным написаниям. В базе подрядчик хранится строкой
// «Подрядчик: <название>» в поле ТОО (см. anketa.html), название вводили
// вручную, поэтому одна организация встречается в нескольких написаниях.
// orgKey() приводит любое написание к одному ключу.
//
// Новый подрядчик: добавить строку в CONTRACTORS — name (как показывать
// и как записывать в анкету) и match (куски названия в нижнем регистре,
// по которым его узнаём; кириллица и латиница, без кавычек).
// ═══════════════════════════════════════════════════════════════
(function () {
  var CONTRACTOR_PREFIX = "Подрядчик: ";

  // Свои ТОО: код в поле «Основное ТОО» → название
  var OWN = [
    { code: "SRC",  name: "Saryarka Resources Capital" },
    { code: "КУ",   name: "Караван Улытау" },
    { code: "CL",   name: "Караван Логистик" },
    { code: "SACP", name: "Sary-Arka Copper Processing" },
    { code: "KAC",  name: "Kyzyl Aray Copper" }
  ];

  var CONTRACTORS = [
    { name: "Орика Казахстан",                 match: ["орика", "orica"] },
    { name: "Eurasian Machinery",              match: ["eurasian", "евразиан"] },
    { name: "Epiroc",                          match: ["epiroc", "эпирок"] },
    { name: "Азиябурмаш",                      match: ["азиябурмаш", "азия бур", "asiaburmash"] },
    { name: "Azot Mining Services Kazakhstan", match: ["azot", "азот"] }
  ];

  // Нижний регистр, одинаковые буквы (турецкая ı → i, ё → е), без кавычек и знаков
  function norm(s) {
    return String(s == null ? "" : s)
      .toLowerCase()
      .replace(/ı/g, "i").replace(/ё/g, "е")
      .replace(/[«»"“”„'`.,;:()\-–—_]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function isContractorValue(too) {
    return String(too || "").indexOf(CONTRACTOR_PREFIX) === 0;
  }

  // Ключ организации по значению поля ТОО:
  //   свои ТОО        → код ("CL", "SRC", ...)
  //   подрядчик       → "Подрядчик: <каноническое название>"
  //   подрядчик, которого нет в справочнике → "Подрядчик: <как введено>"
  //   пусто           → ""
  function orgKey(too) {
    var v = String(too == null ? "" : too).trim();
    if (!v) return "";
    if (isContractorValue(v)) {
      var raw = v.slice(CONTRACTOR_PREFIX.length).trim();
      var n = norm(raw);
      for (var i = 0; i < CONTRACTORS.length; i++) {
        for (var j = 0; j < CONTRACTORS[i].match.length; j++) {
          if (n.indexOf(CONTRACTORS[i].match[j]) !== -1) return CONTRACTOR_PREFIX + CONTRACTORS[i].name;
        }
      }
      return CONTRACTOR_PREFIX + raw;
    }
    return v;
  }

  // Подпись для людей
  function orgLabel(key) {
    if (!key) return "(не указано)";
    if (key.indexOf(CONTRACTOR_PREFIX) === 0) return key.slice(CONTRACTOR_PREFIX.length) + " (подрядчик)";
    for (var i = 0; i < OWN.length; i++) if (OWN[i].code === key) return OWN[i].code + " — " + OWN[i].name;
    return key;
  }

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  // <option>-ы для фильтра. values — массив значений поля ТОО из данных:
  // подрядчики, которых нет в справочнике, всё равно попадут в список,
  // чтобы ни один сотрудник не пропал из фильтра.
  function filterOptionsHtml(values) {
    var present = {};
    (values || []).forEach(function (v) { present[orgKey(v)] = true; });
    var out = ['<option value="">— все организации —</option>'];
    out.push('<optgroup label="Свои ТОО">');
    OWN.forEach(function (o) { out.push('<option value="' + esc(o.code) + '">' + esc(orgLabel(o.code)) + "</option>"); });
    out.push("</optgroup>");
    out.push('<optgroup label="Подрядчики">');
    var known = {};
    CONTRACTORS.forEach(function (c) {
      var k = CONTRACTOR_PREFIX + c.name;
      known[k] = true;
      out.push('<option value="' + esc(k) + '">' + esc(orgLabel(k)) + "</option>");
    });
    Object.keys(present).forEach(function (k) {
      if (k.indexOf(CONTRACTOR_PREFIX) === 0 && !known[k]) {
        out.push('<option value="' + esc(k) + '">' + esc(orgLabel(k)) + " — нет в справочнике</option>");
      }
    });
    out.push("</optgroup>");
    out.push('<option value="__none__">(не указано)</option>');
    return out.join("");
  }

  // Подходит ли значение поля ТОО под выбранный ключ фильтра ("" = все)
  function matches(too, filterKey) {
    if (!filterKey) return true;
    var k = orgKey(too);
    if (filterKey === "__none__") return k === "";
    return k === filterKey;
  }

  window.ORGS = {
    PREFIX: CONTRACTOR_PREFIX,
    OWN: OWN,
    CONTRACTORS: CONTRACTORS,
    contractorNames: function () { return CONTRACTORS.map(function (c) { return c.name; }); },
    orgKey: orgKey,
    orgLabel: orgLabel,
    filterOptionsHtml: filterOptionsHtml,
    matches: matches
  };
})();

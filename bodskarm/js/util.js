/* Gemensamma hjälpfunktioner: datum, text, Excel-värden. */
(function () {
  "use strict";
  var Bod = (window.Bod = window.Bod || {});
  var DAG = 86400000;

  var VECKODAGAR = ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"];
  var VECKODAGAR_KORT = ["sön", "mån", "tis", "ons", "tor", "fre", "lör"];
  var MANADER = ["januari", "februari", "mars", "april", "maj", "juni", "juli",
    "augusti", "september", "oktober", "november", "december"];
  var MANADER_KORT = ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec"];

  var u = (Bod.util = {});

  // Register över vyer. Varje fil i js/vyer/ anropar Bod.vyer.registrera({...}).
  Bod.vyer = Bod.vyer || {
    lista: {},
    registrera: function (vy) { this.lista[vy.id] = vy; }
  };

  /* ---------------------------------------------------------- datum -- */

  // Alla datum hanteras som lokal midnatt så att jämförelser blir enkla.
  u.dag = function (d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  };
  u.plusDagar = function (d, n) {
    var r = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
    return r;
  };
  u.dagarMellan = function (a, b) {
    return Math.round((u.dag(b) - u.dag(a)) / DAG);
  };
  u.mandag = function (d) {
    var wd = (d.getDay() + 6) % 7;
    return u.plusDagar(u.dag(d), -wd);
  };
  u.helg = function (d) {
    var wd = d.getDay();
    return wd === 0 || wd === 6;
  };
  u.isoVecka = function (d) {
    var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    var wd = t.getUTCDay() || 7;
    t.setUTCDate(t.getUTCDate() + 4 - wd);
    var start = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
    return Math.ceil(((t - start) / DAG + 1) / 7);
  };
  u.isoDatum = function (d) {
    return d.getFullYear() + "-" + u.tva(d.getMonth() + 1) + "-" + u.tva(d.getDate());
  };
  u.tva = function (n) {
    return (n < 10 ? "0" : "") + n;
  };
  u.veckodag = function (d) { return VECKODAGAR[d.getDay()]; };
  u.veckodagKort = function (d) { return VECKODAGAR_KORT[d.getDay()]; };
  u.manad = function (d) { return MANADER[d.getMonth()]; };

  // "2 okt"
  u.kortDatum = function (d) {
    return d ? d.getDate() + " " + MANADER_KORT[d.getMonth()] : "";
  };
  // "fredag 2 oktober"
  u.langtDatum = function (d) {
    return VECKODAGAR[d.getDay()] + " " + d.getDate() + " " + MANADER[d.getMonth()];
  };
  u.klocka = function (d) {
    return u.tva(d.getHours()) + ":" + u.tva(d.getMinutes());
  };
  // "Idag", "Imorgon", "måndag 5 okt"
  u.relativDag = function (d, idag) {
    var n = u.dagarMellan(idag, d);
    if (n === 0) return "Idag";
    if (n === 1) return "Imorgon";
    if (n === -1) return "Igår";
    return u.veckodag(d) + " " + u.kortDatum(d);
  };
  u.datumIntervall = function (a, b) {
    if (!b || u.dagarMellan(a, b) === 0) return u.kortDatum(a);
    if (a.getMonth() === b.getMonth()) return a.getDate() + "–" + u.kortDatum(b);
    return u.kortDatum(a) + " – " + u.kortDatum(b);
  };

  /* --------------------------------------------------- Excel-värden -- */

  // Excel-serienummer -> lokalt datum (tid ignoreras).
  u.franSerie = function (n) {
    var ms = Math.round((n - 25569) * DAG);
    var t = new Date(ms);
    return new Date(t.getUTCFullYear(), t.getUTCMonth(), t.getUTCDate());
  };

  // Tolka ett cellvärde som datum: serienummer, Date eller text (2026-10-05 / 5/10 2026).
  u.tillDatum = function (v) {
    if (v === null || v === undefined || v === "") return null;
    if (v instanceof Date) return isNaN(v) ? null : u.dag(v);
    if (typeof v === "number") return v > 1000 ? u.franSerie(v) : null;
    var s = String(v).trim();
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})[./](\d{1,2})[./ ]+(\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    return null;
  };

  // Tolka klockslag: Excel-bråktal (0.2917), Date eller text "07:00" / "7.30".
  u.tillTid = function (v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") {
      var min = Math.round((v % 1) * 1440);
      return u.tva(Math.floor(min / 60) % 24) + ":" + u.tva(min % 60);
    }
    if (v instanceof Date) return u.klocka(v);
    var m = String(v).trim().match(/^(\d{1,2})[:.](\d{2})/);
    return m ? u.tva(+m[1]) + ":" + m[2] : String(v).trim();
  };

  u.tillTal = function (v, standard) {
    if (typeof v === "number" && isFinite(v)) return v;
    if (typeof v === "string") {
      var n = parseFloat(v.replace(",", ".").replace("%", ""));
      if (isFinite(n)) return /%/.test(v) ? n / 100 : n;
    }
    return standard;
  };

  u.text = function (v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  };

  u.jaNej = function (v, standard) {
    var s = u.text(v).toLowerCase();
    if (!s) return standard;
    return s === "ja" || s === "j" || s === "x" || s === "true" || s === "1" || s === "yes";
  };

  /* -------------------------------------------------------------- html -- */

  u.esc = function (s) {
    return u.text(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };

  // Text med radbrytningar -> html.
  u.escRader = function (s) {
    return u.esc(s).replace(/\r?\n/g, "<br>");
  };

  u.procent = function (x) {
    return Math.round((x || 0) * 100) + " %";
  };

  u.dela = function (lista, n) {
    var ut = [];
    for (var i = 0; i < lista.length; i += n) ut.push(lista.slice(i, i + n));
    return ut;
  };

  u.qrSvg = function (text, storlek) {
    if (!text || typeof window.qrcode !== "function") return "";
    try {
      var qr = window.qrcode(0, "M");
      qr.addData(unescape(encodeURIComponent(text)));
      qr.make();
      var antal = qr.getModuleCount();
      var cell = Math.max(2, Math.floor((storlek || 220) / (antal + 8)));
      return qr.createSvgTag({ cellSize: cell, margin: cell * 4, scalable: true });
    } catch (e) {
      return "";
    }
  };

  // Sökväg från config -> URL som fungerar både på disk och på webbserver.
  u.url = function (sokvag) {
    var s = u.text(sokvag).replace(/\\/g, "/");
    if (/^[A-Za-z]:\//.test(s)) return "file:///" + encodeURI(s);
    if (/^\/\//.test(s)) return "file:" + encodeURI(s);
    if (/^[a-z]+:/i.test(s)) return s;
    return encodeURI(s);
  };
})();

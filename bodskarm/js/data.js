/*
 * Läser Excel-filerna och gör om dem till en enkel datamodell.
 *
 *   Bod.data.lasFil(url)             -> Promise<ArrayBuffer>
 *   Bod.data.tolkaPlan(wb, config)   -> { aktiviteter, omraden, sparad }
 *   Bod.data.tolkaInnehall(wb)       -> { installningar, vyer, meddelanden, ... }
 */
(function () {
  "use strict";
  var Bod = (window.Bod = window.Bod || {});
  var u = Bod.util;
  var d = (Bod.data = {});

  /* ------------------------------------------------------- filläsning -- */

  // XMLHttpRequest (inte fetch) eftersom fetch inte kan läsa file://-adresser.
  d.lasFil = function (url) {
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      var sep = /^file:/.test(url) || !/^https?:/.test(location.protocol) ? "" : (url.indexOf("?") < 0 ? "?" : "&") + "t=" + Date.now();
      xhr.open("GET", url + sep, true);
      xhr.responseType = "arraybuffer";
      xhr.onload = function () {
        var ok = xhr.status === 200 || (xhr.status === 0 && xhr.response && xhr.response.byteLength > 0);
        if (ok) resolve(xhr.response);
        else reject(new Error("Kunde inte läsa " + url + " (status " + xhr.status + ")"));
      };
      xhr.onerror = function () {
        reject(new Error("Kunde inte läsa " + url));
      };
      try {
        xhr.send();
      } catch (e) {
        reject(e);
      }
    });
  };

  d.oppnaArbetsbok = function (buffer) {
    // cellDates:false -> datum kommer som serienummer, vilket undviker tidszonsfel.
    return XLSX.read(new Uint8Array(buffer), {
      type: "array", cellDates: false, cellFormula: false, cellHTML: false,
      cellStyles: false, bookVBA: false
    });
  };

  d.sparadTid = function (wb) {
    var p = wb && wb.Props;
    var t = p && (p.ModifiedDate || p.CreatedDate);
    return t instanceof Date && !isNaN(t) ? t : null;
  };

  /* ------------------------------------------------ tabellhjälp -- */

  function rader(ws) {
    if (!ws || !ws["!ref"]) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  }

  function norm(s) {
    return u.text(s).toLowerCase().replace(/\s+/g, " ");
  }

  // Hittar rubrikraden (bland de första 15 raderna) som innehåller alla
  // angivna kolumnnamn. Returnerar { rad, kol: {namn: index} } eller null.
  function hittaRubriker(alla, kravda) {
    for (var r = 0; r < Math.min(alla.length, 15); r++) {
      var rad = alla[r] || [];
      var kol = {};
      for (var c = 0; c < rad.length; c++) {
        var n = norm(rad[c]);
        if (n && !(n in kol)) kol[n] = c;
      }
      var ok = kravda.every(function (k) { return norm(k) in kol; });
      if (ok) return { rad: r, kol: kol };
    }
    return null;
  }

  // Läser ett blad som en tabell med rubrikrad -> lista av objekt med
  // nycklar enligt `falt` ({ nyckel: ["Rubrik", "Alternativ rubrik"] }).
  function lasTabell(wb, bladNamn, falt, kravda) {
    var ws = hittaBlad(wb, bladNamn);
    if (!ws) return [];
    var alla = rader(ws);
    var rub = hittaRubriker(alla, kravda);
    if (!rub) return [];
    var index = {};
    Object.keys(falt).forEach(function (nyckel) {
      var namn = [].concat(falt[nyckel]);
      for (var i = 0; i < namn.length; i++) {
        if (norm(namn[i]) in rub.kol) { index[nyckel] = rub.kol[norm(namn[i])]; break; }
      }
    });
    var ut = [];
    for (var r = rub.rad + 1; r < alla.length; r++) {
      var rad = alla[r] || [];
      var obj = {};
      var tom = true;
      Object.keys(index).forEach(function (k) {
        var v = rad[index[k]];
        if (v !== null && v !== undefined && u.text(v) !== "") tom = false;
        obj[k] = v === undefined ? null : v;
      });
      if (!tom) ut.push(obj);
    }
    return ut;
  }

  function hittaBlad(wb, namn) {
    if (wb.Sheets[namn]) return wb.Sheets[namn];
    var n = norm(namn);
    for (var i = 0; i < wb.SheetNames.length; i++) {
      if (norm(wb.SheetNames[i]) === n) return wb.Sheets[wb.SheetNames[i]];
    }
    return null;
  }

  /* ---------------------------------------------- 4-veckorsplanering -- */

  var PLAN_KOLUMNER = {
    aktivitet: ["Aktivitet"],
    del: ["-", "Del", "Linje", "Del/Linje"],
    typ: ["Typ"],
    start: ["Datum start", "Start"],
    dagar: ["Antal dagar"],
    slut: ["Datum slut", "Slut"],
    dp: ["DP"],
    block: ["BLOCK", "Disciplin"],
    ata: ["ÄTA"],
    framdrift: ["Framdrift %", "Framdrift"],
    helg: ["Helgarbete"],
    ansvarig: ["Ansvarig", "Ansvarig/UE", "UE"]
  };

  // "742 - SIKTHALL" -> { kod: "742", namn: "Sikthall" }
  function delaOmrade(blad) {
    var m = blad.match(/^\s*(\d+)\s*[-–]\s*(.+)$/);
    var namn = m ? m[2] : blad;
    namn = namn.trim().toLowerCase().replace(/(^|[\s(\/-])([a-zåäö])/g, function (x, a, b) { return a + b.toUpperCase(); });
    return { kod: m ? m[1] : "", namn: namn };
  }

  d.tolkaPlan = function (wb, config) {
    var blad = (config.planBlad && config.planBlad.length) ? config.planBlad : wb.SheetNames.filter(function (n) {
      var undanta = (config.planBladUndanta || []).some(function (x) {
        return n.toUpperCase().indexOf(String(x).toUpperCase()) >= 0;
      });
      return !undanta;
    });

    var aktiviteter = [];
    var omraden = [];

    blad.forEach(function (bladNamn) {
      var ws = hittaBlad(wb, bladNamn);
      if (!ws) return;
      // Bara kolumnerna A:Z behövs – resten är Gantt-rutnät.
      var ref = ws["!ref"];
      if (ref) {
        var r = XLSX.utils.decode_range(ref);
        r.e.c = Math.min(r.e.c, 25);
        ws = Object.assign({}, ws, { "!ref": XLSX.utils.encode_range(r) });
      }
      var alla = rader(ws);
      var rub = hittaRubriker(alla, ["Aktivitet", "Datum start"]);
      if (!rub) return;
      var ix = {};
      Object.keys(PLAN_KOLUMNER).forEach(function (k) {
        var namn = PLAN_KOLUMNER[k];
        for (var i = 0; i < namn.length; i++) {
          if (norm(namn[i]) in rub.kol) { ix[k] = rub.kol[norm(namn[i])]; break; }
        }
      });

      var om = delaOmrade(bladNamn);
      var omrade = { id: bladNamn, kod: om.kod, namn: om.namn, index: omraden.length };
      var antal = 0;
      var aktuellDel = "";

      for (var r2 = rub.rad + 1; r2 < alla.length; r2++) {
        var rad = alla[r2] || [];
        var cell = function (k) { return ix[k] === undefined ? null : rad[ix[k]]; };
        var namn = u.text(cell("aktivitet"));
        var start = u.tillDatum(cell("start"));
        var delText = u.text(cell("del"));
        if (!start) {
          // Rubrikrad ("Linje M") – kom ihåg som grupp.
          if (namn || delText) aktuellDel = delText || namn;
          continue;
        }
        if (!namn) continue;
        var slut = u.tillDatum(cell("slut")) || start;
        if (slut < start) slut = start;
        var framdrift = u.tillTal(cell("framdrift"), 0);
        if (framdrift > 1.0001) framdrift = framdrift / 100;
        antal++;
        aktiviteter.push({
          id: (om.kod || bladNamn) + "-" + (r2 + 1),
          rad: r2 + 1,
          omrade: omrade,
          del: delText || aktuellDel,
          aktivitet: namn,
          typ: u.text(cell("typ")),
          start: start,
          slut: slut,
          dagar: u.tillTal(cell("dagar"), u.dagarMellan(start, slut) + 1),
          dp: u.text(cell("dp")),
          block: u.text(cell("block")),
          ata: u.jaNej(cell("ata"), false),
          helg: u.jaNej(cell("helg"), false),
          ansvarig: u.text(cell("ansvarig")),
          framdrift: Math.max(0, Math.min(1, framdrift)),
          preliminar: /prelimin/i.test(u.text(cell("typ"))),
          milstolpe: u.tillTal(cell("dagar"), 1) === 0
        });
      }
      if (antal) omraden.push(omrade);
    });

    aktiviteter.sort(function (a, b) {
      return a.omrade.index - b.omrade.index || a.start - b.start || a.slut - b.slut || a.rad - b.rad;
    });
    return { aktiviteter: aktiviteter, omraden: omraden, sparad: d.sparadTid(wb) };
  };

  // Status för en aktivitet en viss dag.
  d.status = function (a, idag, config) {
    if (a.framdrift >= 0.999) return "klar";
    if (a.slut < idag) return config.visaForsenade === false ? "klar" : "forsenad";
    if (a.start <= idag) return "pagar";
    return "kommande";
  };

  d.STATUS_TEXT = { klar: "Klar", forsenad: "Försenad", pagar: "Pågår", kommande: "Kommande" };

  /* ------------------------------------------------ innehållsfilen -- */

  var STANDARD_VYER = [
    { id: "laget", namn: "Läget just nu", visa: true, sekunder: 20 },
    { id: "vecka", namn: "Denna vecka", visa: true, sekunder: 25 },
    { id: "fyraveckor", namn: "4 veckor framåt", visa: true, sekunder: 25 },
    { id: "kommande", namn: "Startar & ska bli klart", visa: true, sekunder: 20 },
    { id: "bilder", namn: "Modell & ritningar", visa: true, sekunder: 20 },
    { id: "leveranser", namn: "Leveranser & lyft", visa: true, sekunder: 20 },
    { id: "hinder", namn: "Hinder & risker", visa: true, sekunder: 20 },
    { id: "vader", namn: "Väder", visa: true, sekunder: 15 },
    { id: "meddelanden", namn: "Meddelanden & säkerhet", visa: true, sekunder: 20 },
    { id: "kontakter", namn: "Kontakter & nödläge", visa: true, sekunder: 15 }
  ];
  d.STANDARD_VYER = STANDARD_VYER;

  d.tolkaInnehall = function (wb) {
    var inst = {};
    lasTabell(wb, "Inställningar", { namn: "Inställning", varde: "Värde" }, ["Inställning", "Värde"])
      .forEach(function (r) { inst[norm(r.namn)] = r.varde; });

    var vyer = lasTabell(wb, "Vyer",
      { id: "Vy-id", namn: "Namn", visa: "Visa", sekunder: "Sekunder", ordning: "Ordning" },
      ["Vy-id", "Visa"]).map(function (r, i) {
        return {
          id: u.text(r.id).toLowerCase(),
          namn: u.text(r.namn),
          visa: u.jaNej(r.visa, true),
          sekunder: u.tillTal(r.sekunder, 0),
          ordning: u.tillTal(r.ordning, i + 1)
        };
      });

    var meddelanden = lasTabell(wb, "Meddelanden", {
      rubrik: "Rubrik", text: "Text", typ: "Typ", fran: ["Visa från", "Från"],
      till: ["Visa till", "Till"], rullande: ["Rullande text", "Rullande"]
    }, ["Rubrik"]).map(function (r) {
      return {
        rubrik: u.text(r.rubrik), text: u.text(r.text), typ: u.text(r.typ) || "Info",
        fran: u.tillDatum(r.fran), till: u.tillDatum(r.till), rullande: u.jaNej(r.rullande, false)
      };
    });

    var hinder = lasTabell(wb, "Hinder & risker", {
      omrade: "Område", beskrivning: "Beskrivning", konsekvens: "Konsekvens", atgard: "Åtgärd",
      ansvarig: "Ansvarig", klart: ["Klart senast", "Datum"], status: "Status", niva: ["Allvarlighet", "Nivå"]
    }, ["Beskrivning"]).map(function (r) {
      return {
        omrade: u.text(r.omrade), beskrivning: u.text(r.beskrivning), konsekvens: u.text(r.konsekvens),
        atgard: u.text(r.atgard), ansvarig: u.text(r.ansvarig), klart: u.tillDatum(r.klart),
        status: u.text(r.status) || "Öppen", niva: u.text(r.niva) || "Medel"
      };
    });

    var leveranser = lasTabell(wb, "Leveranser", {
      datum: "Datum", tid: "Tid", vad: ["Vad", "Leverans"], leverantor: "Leverantör",
      plats: ["Plats", "Område/plats", "Område"], lyft: ["Kranlyft", "Lyft"], mottagare: ["Mottagare", "Kontakt"],
      kommentar: "Kommentar"
    }, ["Datum", "Vad"]).map(function (r) {
      return {
        datum: u.tillDatum(r.datum), tid: u.tillTid(r.tid), vad: u.text(r.vad), leverantor: u.text(r.leverantor),
        plats: u.text(r.plats), lyft: u.jaNej(r.lyft, false), mottagare: u.text(r.mottagare), kommentar: u.text(r.kommentar)
      };
    }).filter(function (r) { return r.datum; });

    var bilder = lasTabell(wb, "Bilder", {
      ordning: "Ordning", rubrik: "Rubrik", beskrivning: "Beskrivning", fil: ["Bildfil", "Fil"],
      lank: ["Länk", "Länk (QR-kod)"], kategori: "Kategori", visa: "Visa", fran: ["Visa från", "Från"],
      till: ["Visa till", "Till"]
    }, ["Rubrik", "Bildfil"]).map(function (r, i) {
      return {
        ordning: u.tillTal(r.ordning, i + 1), rubrik: u.text(r.rubrik), beskrivning: u.text(r.beskrivning),
        fil: u.text(r.fil), lank: u.text(r.lank), kategori: u.text(r.kategori), visa: u.jaNej(r.visa, true),
        fran: u.tillDatum(r.fran), till: u.tillDatum(r.till)
      };
    }).filter(function (b) { return b.fil; });

    var kontakter = lasTabell(wb, "Kontakter", {
      roll: "Roll", namn: "Namn", telefon: "Telefon", foretag: "Företag", visa: "Visa"
    }, ["Roll", "Namn"]).map(function (r) {
      return {
        roll: u.text(r.roll), namn: u.text(r.namn), telefon: u.text(r.telefon),
        foretag: u.text(r.foretag), visa: u.jaNej(r.visa, true)
      };
    }).filter(function (k) { return k.visa && (k.namn || k.telefon); });

    return {
      installningar: inst, vyer: vyer, meddelanden: meddelanden, hinder: hinder,
      leveranser: leveranser, bilder: bilder, kontakter: kontakter, sparad: d.sparadTid(wb)
    };
  };

  // Hämtar en inställning ur innehållsfilen (skiftlägesokänsligt namn).
  d.inst = function (innehall, namn, standard) {
    var v = innehall && innehall.installningar ? innehall.installningar[norm(namn)] : undefined;
    return v === undefined || v === null || u.text(v) === "" ? standard : v;
  };

  // Aktiva poster i ett datumintervall (fran/till kan saknas).
  d.aktiv = function (post, idag) {
    if (post.fran && post.fran > idag) return false;
    if (post.till && post.till < idag) return false;
    return true;
  };
})();

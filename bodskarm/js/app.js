/*
 * Bodskärmens motor: läser in data, roterar vyerna, ritar sidhuvud/sidfot.
 *
 * Tangentbord (för test): ← → byt vy · mellanslag pausa · R läs om data
 * URL-parametrar (för test): ?datum=2026-10-05  ?vy=vecka  ?sekunder=5
 */
(function () {
  "use strict";
  var Bod = window.Bod;
  var u = Bod.util;
  var D = Bod.data;
  var config = window.BODSKARM_CONFIG || {};
  var param = new URLSearchParams(location.search);

  var state = {
    plan: null, innehall: null, vader: null,
    fel: { plan: null, innehall: null, vader: null },
    lastData: null,
    vyIndex: -1, sidIndex: 0, sidor: [], vy: null,
    paus: param.has("paus"), timer: null, slutTid: 0, sekunder: 20
  };

  var el = {};

  /* ------------------------------------------------------ tid & ctx -- */

  function nu() {
    var d = new Date();
    var p = param.get("datum");
    var f = p && u.tillDatum(p);
    if (f) return new Date(f.getFullYear(), f.getMonth(), f.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
    return d;
  }

  function inst(namn, standard) {
    return D.inst(state.innehall, namn, standard);
  }

  function ctx() {
    var n = nu();
    var idag = u.dag(n);
    return {
      nu: n, idag: idag, config: config, plan: state.plan, innehall: state.innehall, vader: state.vader,
      inst: inst,
      status: function (a) { return D.status(a, idag, config); },
      bildUrl: function (fil) {
        if (/^([a-z]+:|[A-Za-z]:[\\/]|\/)/.test(fil)) return u.url(fil);
        return u.url((config.bildMapp || "data/bilder/") + fil) + "?v=" + (state.lastData || 0);
      }
    };
  }

  /* --------------------------------------------------------- data -- */

  function lasPlan() {
    return D.lasFil(u.url(config.planFil)).then(function (buf) {
      state.plan = D.tolkaPlan(D.oppnaArbetsbok(buf), config);
      state.fel.plan = null;
    }).catch(function (e) {
      state.fel.plan = e.message || String(e);
      console.warn(e);
    });
  }

  function lasInnehall() {
    return D.lasFil(u.url(config.innehallFil)).then(function (buf) {
      state.innehall = D.tolkaInnehall(D.oppnaArbetsbok(buf));
      state.fel.innehall = null;
    }).catch(function (e) {
      state.fel.innehall = e.message || String(e);
      console.warn(e);
    });
  }

  function lasData() {
    return Promise.all([lasPlan(), lasInnehall()]).then(function () {
      state.lastData = Date.now();
      sidfot();
      if (!state.vader) hamtaVader();
    });
  }

  function koordinater() {
    var lat = u.tillTal(inst("Latitud", ""), NaN);
    var lon = u.tillTal(inst("Longitud", ""), NaN);
    return isFinite(lat) && isFinite(lon) ? { lat: lat, lon: lon } : null;
  }

  function hamtaVader() {
    var k = koordinater();
    if (!k) { state.vader = null; return Promise.resolve(); }
    if (!state.vader) state.vader = Bod.vader.senaste(k.lat, k.lon);
    return Bod.vader.hamta(k.lat, k.lon).then(function (w) {
      state.vader = w;
      state.fel.vader = null;
      sidhuvud();
    }).catch(function (e) {
      state.fel.vader = e.message;
      console.warn(e);
    });
  }

  /* --------------------------------------------------- vyrotation -- */

  function vyLista() {
    var fran = (state.innehall && state.innehall.vyer.length) ? state.innehall.vyer : [];
    var std = D.STANDARD_VYER;
    var lista = std.map(function (s, i) {
      var r = fran.filter(function (v) { return v.id === s.id; })[0];
      return {
        id: s.id, namn: (r && r.namn) || s.namn, visa: r ? r.visa : s.visa,
        sekunder: (r && r.sekunder) || 0, ordning: r ? r.ordning : 100 + i
      };
    });
    // Egna vyer som lagts till i js/vyer/ och i Excel-bladet "Vyer".
    fran.forEach(function (r) {
      if (!std.some(function (s) { return s.id === r.id; }) && Bod.vyer.lista[r.id]) lista.push(r);
    });
    var bara = param.get("vy");
    return lista.filter(function (v) {
      return Bod.vyer.lista[v.id] && (bara ? v.id === bara : v.visa);
    }).sort(function (a, b) { return a.ordning - b.ordning; });
  }

  function standardSekunder() {
    return u.tillTal(inst("Standard visningstid (sek)", 20), 20);
  }

  function byggSidor(vy) {
    try {
      return Bod.vyer.lista[vy.id].sidor(ctx()) || [];
    } catch (e) {
      console.error("Fel i vyn " + vy.id, e);
      return [];
    }
  }

  function nasta(steg) {
    steg = steg || 1;
    var vyer = vyLista();
    if (!vyer.length) return visaTomt();

    // Fler sidor kvar i samma vy?
    var ny = state.sidIndex + steg;
    if (state.vy && ny >= 0 && ny < state.sidor.length) {
      state.sidIndex = ny;
      return visa();
    }

    var start = vyer.map(function (v) { return v.id; }).indexOf(state.vy && state.vy.id);
    for (var i = 1; i <= vyer.length; i++) {
      var ix = ((start + steg * i) % vyer.length + vyer.length) % vyer.length;
      var sidor = byggSidor(vyer[ix]);
      if (sidor.length) {
        state.vy = vyer[ix];
        state.sidor = sidor;
        state.sidIndex = steg < 0 ? sidor.length - 1 : 0;
        return visa();
      }
    }
    visaTomt();
  }

  function visa() {
    var vy = state.vy;
    var sida = state.sidor[state.sidIndex];
    var vyer = vyLista();
    state.sekunder = u.tillTal(param.get("sekunder"), 0) || vy.sekunder || standardSekunder();

    el.titel.textContent = sida.titel || vy.namn;
    el.undertitel.textContent = sida.undertitel || "";
    el.prickar.innerHTML = vyer.map(function (v) {
      return '<i class="' + (v.id === vy.id ? "aktiv" : "") + '"></i>';
    }).join("") + (state.sidor.length > 1 ? '<span class="sidnr">' + (state.sidIndex + 1) + "/" + state.sidor.length + "</span>" : "");

    var gammal = el.innehall.firstElementChild;
    var div = document.createElement("div");
    div.className = "vy vy-" + vy.id;
    div.innerHTML = sida.html;
    el.innehall.appendChild(div);
    requestAnimationFrame(function () { div.classList.add("synlig"); });
    if (gammal) {
      gammal.classList.remove("synlig");
      setTimeout(function () { if (gammal.parentNode) gammal.parentNode.removeChild(gammal); }, 700);
    }
    startaTimer();
    ticker();
  }

  // Bygger om den vy som visas med färsk data (efter manuell omläsning).
  function ritaOm() {
    if (!state.vy) return nasta(1);
    state.sidor = byggSidor(state.vy);
    if (!state.sidor.length) return nasta(1);
    state.sidIndex = Math.min(state.sidIndex, state.sidor.length - 1);
    visa();
  }

  function visaTomt() {
    state.vy = null;
    el.titel.textContent = "Bodskärm";
    el.undertitel.textContent = "";
    el.innehall.innerHTML = '<div class="vy synlig">' + felruta() + "</div>";
    startaTimer();
  }

  function startaTimer() {
    clearTimeout(state.timer);
    var ms = state.sekunder * 1000;
    state.slutTid = Date.now() + ms;
    el.progress.style.transition = "none";
    el.progress.style.width = "0%";
    if (state.paus) return;
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        el.progress.style.transition = "width " + ms + "ms linear";
        el.progress.style.width = "100%";
      });
    });
    state.timer = setTimeout(function () { nasta(1); }, ms);
  }

  /* ------------------------------------------- sidhuvud & sidfot -- */

  function sidhuvud() {
    var n = nu();
    el.projekt.textContent = u.text(inst("Projektnamn", "Bodskärm"));
    el.underrubrik.textContent = u.text(inst("Underrubrik", ""));
    el.vecka.textContent = "Vecka " + u.isoVecka(n);
    el.datum.textContent = u.langtDatum(n);
    el.klocka.textContent = u.klocka(n);
    var w = state.vader;
    if (w && w.nu && w.nu.temp != null) {
      el.vader.innerHTML = '<span class="vh-ikon">' + Bod.vader.kod(w.nu.kod)[1] + "</span>" +
        '<span class="vh-temp">' + Bod.vader.grader(w.nu.temp) + "</span>" +
        '<span class="vh-vind">' + Math.round(w.nu.vind) + "<small> (" + Math.round(w.nu.byar) + ") m/s</small></span>";
      el.vader.hidden = false;
    } else {
      el.vader.hidden = true;
    }
  }

  function sidfot() {
    var delar = [];
    if (state.plan && state.plan.sparad) delar.push("Planering sparad " + u.kortDatum(state.plan.sparad) + " " + u.klocka(state.plan.sparad));
    if (state.lastData) delar.push("inläst " + u.klocka(new Date(state.lastData)));
    var vad = state.fel.plan ? "planeringen" : state.fel.innehall ? "innehållsfilen" : "";
    var harGammal = state.fel.plan ? state.plan : state.innehall;
    el.status.innerHTML = '<i class="' + (vad ? "fel" : "ok") + '"></i>' + u.esc(delar.join(" · ")) +
      (vad ? ' <span class="rod">· kunde inte läsa ' + vad + (harGammal ? ", visar senast inlästa" : "") + "</span>" : "");
  }

  function ticker() {
    var c = ctx();
    var texter = (state.innehall ? state.innehall.meddelanden : [])
      .filter(function (m) { return m.rullande && D.aktiv(m, c.idag); })
      .map(function (m) { return "<b>" + u.esc(m.rubrik) + "</b>" + (m.text ? " – " + u.esc(m.text.replace(/\s+/g, " ")) : ""); });

    var grans = u.tillTal(inst("Vindgräns lyft (m/s)", 12), 12);
    var byar = Bod.vader.maxByarArbetstid(state.vader, c.idag);
    if (byar != null && byar >= grans) texter.unshift('<b class="rod">⚠ Vindbyar upp till ' + Math.round(byar) + " m/s idag – kontrollera lyft</b>");

    var lev = (state.innehall ? state.innehall.leveranser : []).filter(function (l) { return u.dagarMellan(c.idag, l.datum) === 0; });
    if (lev.length) texter.push("<b>Leveranser idag:</b> " + lev.map(function (l) { return u.esc((l.tid ? l.tid + " " : "") + l.vad); }).join(" · "));

    var html = texter.join('<span class="sep">◆</span>');
    if (html === el.ticker.dataset.html) return;
    el.ticker.dataset.html = html;
    el.ticker.innerHTML = html ? '<div class="ticker-spar"><span>' + html + '</span><span aria-hidden="true">' + html + "</span></div>" : "";
    var spar = el.ticker.firstElementChild;
    if (spar) {
      var bredd = spar.firstElementChild.offsetWidth;
      if (bredd < el.ticker.offsetWidth * 0.9) {
        spar.classList.add("still");
        spar.lastElementChild.remove();
      } else {
        spar.style.animationDuration = Math.max(20, bredd / 90) + "s";
      }
    }
  }

  function felruta() {
    var f = state.fel;
    var rader = [];
    if (f.plan) rader.push("<li><b>Planeringen:</b> " + u.esc(f.plan) + "</li>");
    if (f.innehall) rader.push("<li><b>Innehållsfilen:</b> " + u.esc(f.innehall) + "</li>");
    var fil = location.protocol === "file:";
    return '<div class="felruta"><h2>' + (rader.length ? "Kunde inte läsa Excel-filerna" : "Inget att visa just nu") + "</h2>" +
      (rader.length ? "<ul>" + rader.join("") + "</ul>" : "<p>Alla vyer är avstängda eller tomma. Kontrollera bladet <b>Vyer</b> i innehållsfilen.</p>") +
      (rader.length && fil ? "<p>Starta skärmen med <code>kiosk\\Starta bodskarm.bat</code> (eller <code>kiosk\\Forhandsgranska.bat</code>) – " +
        "webbläsaren behöver startas med tillåtelse att läsa lokala filer. Kontrollera också sökvägarna i <code>config.js</code>.</p>" : "") +
      (rader.length ? '<p>Vill du bara titta? Välj filerna manuellt:</p><label class="knapp">Välj Excel-filer…<input type="file" id="valjFiler" multiple accept=".xlsx,.xlsm"></label>' : "") +
      "</div>";
  }

  // Manuellt filval – bra för att förhandsgranska utan startskriptet.
  document.addEventListener("change", function (e) {
    if (e.target.id !== "valjFiler") return;
    var filer = Array.prototype.slice.call(e.target.files);
    Promise.all(filer.map(function (f) {
      return f.arrayBuffer().then(function (buf) {
        var wb = D.oppnaArbetsbok(buf);
        if (wb.SheetNames.indexOf("Inställningar") >= 0 || wb.SheetNames.indexOf("Vyer") >= 0) {
          state.innehall = D.tolkaInnehall(wb); state.fel.innehall = null;
        } else {
          state.plan = D.tolkaPlan(wb, config); state.fel.plan = null;
        }
      });
    })).then(function () {
      state.lastData = Date.now();
      config.uppdateraMinuter = 0; // läs inte om från disk – filerna valdes manuellt
      sidhuvud(); sidfot(); hamtaVader(); nasta(1);
    });
  });

  /* ------------------------------------------------- skalning -- */

  function skala() {
    var s = Math.min(window.innerWidth / 1920, window.innerHeight / 1080);
    el.scen.style.transform = "translate(-50%, -50%) scale(" + s + ")";
  }

  /* -------------------------------------------------- start -- */

  function start() {
    ["scen", "projekt", "underrubrik", "titel", "undertitel", "prickar", "vecka", "datum", "klocka", "vader",
      "innehall", "progress", "status", "ticker"].forEach(function (id) { el[id] = document.getElementById(id); });
    skala();
    window.addEventListener("resize", skala);

    document.addEventListener("keydown", function (e) {
      if (e.key === "ArrowRight") nasta(1);
      else if (e.key === "ArrowLeft") nasta(-1);
      else if (e.key === " ") { state.paus = !state.paus; document.body.classList.toggle("pausad", state.paus); startaTimer(); }
      else if (e.key === "r" || e.key === "R") lasData().then(ritaOm);
    });
    document.body.classList.toggle("pausad", state.paus);

    lasData().then(function () {
      sidhuvud();
      nasta(1);
      document.body.classList.add("redo");
    });

    setInterval(sidhuvud, 1000 * 15);

    var minuter = config.uppdateraMinuter || 2;
    setInterval(function () {
      if (config.uppdateraMinuter) lasData();
    }, minuter * 60000);

    setInterval(hamtaVader, (config.vaderMinuter || 30) * 60000);

    // Ladda om hela sidan en gång per natt.
    var omstart = String(config.omstartKlockan || "04:00").split(":");
    setInterval(function () {
      var d = new Date();
      if (d.getHours() === +omstart[0] && d.getMinutes() === +(omstart[1] || 0) && performance.now() > 120000) location.reload();
    }, 30000);
  }

  window.Bod.app = { state: state, nasta: nasta, lasData: lasData };
  document.addEventListener("DOMContentLoaded", start);
})();

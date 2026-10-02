/* Vyerna som bygger på 4-veckorsplaneringen: Läget, Denna vecka, 4 veckor framåt, Kommande. */
(function () {
  "use strict";
  var Bod = window.Bod;
  var u = Bod.util;

  function harPlan(ctx) {
    return ctx.plan && ctx.plan.aktiviteter.length;
  }

  // Aktiviteter som berör intervallet [fran, till].
  function iIntervall(ctx, fran, till) {
    return ctx.plan.aktiviteter.filter(function (a) {
      return a.start <= till && a.slut >= fran;
    });
  }

  function planeradAndel(a, idag) {
    if (idag >= a.slut) return 1;
    if (idag < a.start) return 0;
    return (u.dagarMellan(a.start, idag) + 1) / (u.dagarMellan(a.start, a.slut) + 1);
  }

  function framdrift(lista, idag) {
    var vikt = 0, verklig = 0, plan = 0;
    lista.forEach(function (a) {
      var w = Math.max(a.dagar || 0, a.milstolpe ? 0 : 1);
      vikt += w;
      verklig += w * a.framdrift;
      plan += w * planeradAndel(a, idag);
    });
    return { verklig: vikt ? verklig / vikt : 0, plan: vikt ? plan / vikt : 0 };
  }

  function tidplanSidor(ctx, fran, dagar, rubrik) {
    if (!harPlan(ctx)) return [];
    var till = u.plusDagar(fran, dagar - 1);
    var lista = iIntervall(ctx, fran, till);
    if (!lista.length) {
      return [{ undertitel: rubrik, html: '<div class="tomt">Inga aktiviteter planerade ' + u.esc(rubrik.toLowerCase()) + ".</div>" }];
    }
    var sidor = Bod.gantt.sidor(lista, ctx.config.raderPerSida || 13);
    return sidor.map(function (rader) {
      return {
        undertitel: rubrik + " · " + lista.length + " aktiviteter",
        html: Bod.gantt.rita({ rader: rader, fran: fran, dagar: dagar, idag: ctx.idag, status: ctx.status }) + legend()
      };
    });
  }

  function legend() {
    return '<div class="legend">' +
      ["pagar", "kommande", "klar", "forsenad"].map(function (s) {
        return '<span><i class="st-' + s + '"></i>' + Bod.data.STATUS_TEXT[s] + "</span>";
      }).join("") +
      '<span><i class="prelim-ikon"></i>Preliminär</span><span><i class="milstolpe-ikon"></i>Syn / milstolpe</span>' +
      '<span class="dampad">Linjen = idag</span></div>';
  }

  /* -------------------------------------------------------- Läget -- */
  Bod.vyer.registrera({
    id: "laget",
    titel: "Läget just nu",
    sidor: function (ctx) {
      if (!harPlan(ctx)) return [];
      var idag = ctx.idag;
      var man = u.mandag(idag);
      var son = u.plusDagar(man, 6);
      var alla = ctx.plan.aktiviteter;

      var pagar = alla.filter(function (a) { return ctx.status(a) === "pagar"; });
      var startar = alla.filter(function (a) { return a.start > idag && a.start <= son; });
      var klartVecka = alla.filter(function (a) { return a.slut >= idag && a.slut <= son && ctx.status(a) !== "klar"; });
      var forsenade = alla.filter(function (a) { return ctx.status(a) === "forsenad"; });

      function kpi(tal, text, kl) {
        return '<div class="kpi ' + (kl || "") + '"><div class="kpi-tal">' + tal + '</div><div class="kpi-text">' + text + "</div></div>";
      }
      var kpier = '<div class="kpi-rad">' +
        kpi(pagar.length, "pågår idag", "bla") +
        kpi(startar.length, "startar resten av veckan") +
        kpi(klartVecka.length, "ska bli klara denna vecka") +
        (ctx.config.visaForsenade === false ? "" : kpi(forsenade.length, "försenade / ej avrapporterade", forsenade.length ? "rod" : "gron")) +
        "</div>";

      var kort = ctx.plan.omraden.map(function (o) {
        var egna = alla.filter(function (a) { return a.omrade === o; });
        var f = framdrift(egna, idag);
        var nu = egna.filter(function (a) { return ctx.status(a) === "pagar"; });
        var snart = egna.filter(function (a) { return a.start > idag && a.start <= u.plusDagar(idag, 7); });
        var sena = egna.filter(function (a) { return ctx.status(a) === "forsenad"; });
        var klara = egna.filter(function (a) { return ctx.status(a) === "klar"; });
        var nasta = egna.filter(function (a) { return a.start > idag; })[0];
        var just = nu.slice(0, 3).map(function (a) { return "<li>" + u.esc(a.aktivitet) + "</li>"; }).join("");
        var diff = Math.round((f.verklig - f.plan) * 100);
        return '<div class="omr-kort" style="--omr:' + Bod.gantt.farg(o) + '">' +
          '<div class="omr-huvud"><span class="omr-kod">' + u.esc(o.kod) + '</span><span class="omr-namn">' + u.esc(o.namn) + "</span></div>" +
          '<div class="omr-framdrift"><div class="omr-procent">' + u.procent(f.verklig) + "</div>" +
          '<div class="omr-plan">plan idag ' + u.procent(f.plan) +
          (Math.abs(diff) >= 3 ? ' <b class="' + (diff < 0 ? "rod" : "gron") + '">' + (diff > 0 ? "+" : "") + diff + "</b>" : "") + "</div></div>" +
          '<div class="stapel-yttre"><div class="stapel-plan" style="width:' + Math.round(f.plan * 100) + '%"></div>' +
          '<div class="stapel-inre" style="width:' + Math.round(f.verklig * 100) + '%"></div></div>' +
          '<div class="omr-tal"><span><b>' + nu.length + "</b> pågår</span><span><b>" + snart.length + "</b> startar 7 d</span>" +
          "<span><b>" + klara.length + "</b>/" + egna.length + " klara</span>" +
          (sena.length && ctx.config.visaForsenade !== false ? '<span class="rod"><b>' + sena.length + "</b> försenade</span>" : "") + "</div>" +
          (just ? '<div class="omr-nu"><div class="liten-rubrik">Pågår nu</div><ul>' + just + (nu.length > 3 ? '<li class="dampad">+ ' + (nu.length - 3) + " till</li>" : "") + "</ul></div>"
            : nasta ? '<div class="omr-nu"><div class="liten-rubrik">Nästa start</div><ul><li>' + u.esc(nasta.aktivitet) + ' <span class="dampad">' + u.relativDag(nasta.start, idag) + "</span></li></ul></div>" : "") +
          "</div>";
      }).join("");

      return [{
        undertitel: "Framdrift enligt planeringen per " + u.kortDatum(idag),
        html: kpier + '<div class="omr-rutnat n' + ctx.plan.omraden.length + '">' + kort + "</div>"
      }];
    }
  });

  /* -------------------------------------------------- Denna vecka -- */
  Bod.vyer.registrera({
    id: "vecka",
    titel: "Denna vecka",
    sidor: function (ctx) {
      var man = u.mandag(ctx.idag);
      return tidplanSidor(ctx, man, 7, "Vecka " + u.isoVecka(ctx.idag) + " (" + u.datumIntervall(man, u.plusDagar(man, 6)) + ")");
    }
  });

  /* --------------------------------------------- 4 veckor framåt -- */
  Bod.vyer.registrera({
    id: "fyraveckor",
    titel: "4 veckor framåt",
    sidor: function (ctx) {
      var man = u.mandag(ctx.idag);
      var sista = u.plusDagar(man, 27);
      return tidplanSidor(ctx, man, 28, "Vecka " + u.isoVecka(man) + "–" + u.isoVecka(sista));
    }
  });

  /* ---------------------------------------------------- Kommande -- */
  Bod.vyer.registrera({
    id: "kommande",
    titel: "Startar & ska bli klart",
    sidor: function (ctx) {
      if (!harPlan(ctx)) return [];
      var idag = ctx.idag;
      var dagar = ctx.config.dagarFramat || 14;
      var till = u.plusDagar(idag, dagar);
      var max = 11;
      var startar = ctx.plan.aktiviteter.filter(function (a) { return a.start >= idag && a.start <= till; })
        .sort(function (a, b) { return a.start - b.start; });
      var klart = ctx.plan.aktiviteter.filter(function (a) { return a.slut >= idag && a.slut <= till && ctx.status(a) !== "klar"; })
        .sort(function (a, b) { return a.slut - b.slut; });

      function lista(akt, falt) {
        if (!akt.length) return '<div class="tomt liten">Inget de närmaste ' + dagar + " dagarna.</div>";
        return '<ul class="handelser">' + akt.slice(0, max).map(function (a) {
          var d = a[falt];
          var n = u.dagarMellan(idag, d);
          return '<li style="--omr:' + Bod.gantt.farg(a.omrade) + '"><div class="h-datum' + (n <= 1 ? " snart" : "") + '"><b>' +
            (n === 0 ? "Idag" : n === 1 ? "Imorgon" : u.veckodagKort(d)) + "</b><span>" + u.kortDatum(d) + "</span></div>" +
            '<div class="h-text"><div class="h-namn">' + u.esc(a.aktivitet) + '</div><div class="h-detalj"><span class="omr-kod">' +
            u.esc(a.omrade.kod) + "</span> " + u.esc(a.omrade.namn) + (Bod.gantt.del(a) ? " · " + u.esc(Bod.gantt.del(a)) : "") + "</div></div></li>";
        }).join("") + (akt.length > max ? '<li class="fler">+ ' + (akt.length - max) + " till</li>" : "") + "</ul>";
      }

      return [{
        undertitel: "De närmaste " + dagar + " dagarna",
        html: '<div class="tva-kol"><section><h2>▶ Startar</h2>' + lista(startar, "start") + "</section>" +
          "<section><h2>✔ Ska bli klart</h2>" + lista(klart, "slut") + "</section></div>"
      }];
    }
  });
})();

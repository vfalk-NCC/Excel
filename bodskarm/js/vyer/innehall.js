/* Vyerna som bygger på Bodskarm_Innehall.xlsx: bilder, leveranser, hinder, meddelanden, kontakter. */
(function () {
  "use strict";
  var Bod = window.Bod;
  var u = Bod.util;
  var D = Bod.data;

  /* ------------------------------------------------------ Bilder -- */
  Bod.vyer.registrera({
    id: "bilder",
    titel: "Modell & ritningar",
    sidor: function (ctx) {
      var bilder = (ctx.innehall ? ctx.innehall.bilder : [])
        .filter(function (b) { return b.visa && D.aktiv(b, ctx.idag); })
        .sort(function (a, b) { return a.ordning - b.ordning; });
      return bilder.map(function (b, i) {
        var qr = b.lank ? u.qrSvg(b.lank, 260) : "";
        return {
          titel: b.kategori || null,
          undertitel: b.rubrik + (bilder.length > 1 ? " · bild " + (i + 1) + " av " + bilder.length : ""),
          html: '<div class="bild-vy"><div class="bild-ram"><img src="' + u.esc(ctx.bildUrl(b.fil)) + '" alt="" ' +
            'onerror="this.parentNode.classList.add(\'saknas\')"><div class="bild-fel">Hittar inte bilden<br><code>' +
            u.esc(b.fil) + "</code></div></div>" +
            '<aside class="bild-info">' + (b.kategori ? '<div class="etikett">' + u.esc(b.kategori) + "</div>" : "") +
            "<h2>" + u.esc(b.rubrik) + "</h2>" + (b.beskrivning ? "<p>" + u.escRader(b.beskrivning) + "</p>" : "") +
            (qr ? '<div class="qr"><div class="qr-bild">' + qr + '</div><div class="qr-text">Skanna med mobilen<br>för att öppna i Trimble Connect</div></div>' : "") +
            "</aside></div>"
        };
      });
    }
  });

  /* -------------------------------------------------- Leveranser -- */
  Bod.vyer.registrera({
    id: "leveranser",
    titel: "Leveranser & lyft",
    sidor: function (ctx) {
      if (!ctx.innehall) return [];
      var idag = ctx.idag;
      var dagar = ctx.config.dagarFramat || 14;
      var till = u.plusDagar(idag, dagar);
      var lista = (ctx.innehall ? ctx.innehall.leveranser : [])
        .filter(function (l) { return l.datum >= idag && l.datum <= till; })
        .sort(function (a, b) { return a.datum - b.datum || (a.tid || "99").localeCompare(b.tid || "99"); });
      if (!lista.length) {
        return [{ undertitel: "De närmaste " + dagar + " dagarna", html: '<div class="tomt">Inga leveranser inlagda de närmaste ' + dagar + " dagarna.</div>" }];
      }
      var grans = u.tillTal(ctx.inst("Vindgräns lyft (m/s)", 12), 12);

      // Gruppera per dag och dela upp i sidor om ca 9 rader.
      var grupper = [];
      lista.forEach(function (l) {
        var g = grupper[grupper.length - 1];
        if (!g || u.dagarMellan(g.datum, l.datum) !== 0) grupper.push(g = { datum: l.datum, rader: [] });
        g.rader.push(l);
      });
      var sidor = [[]];
      var antal = 0;
      grupper.forEach(function (g) {
        if (antal + g.rader.length + 1 > 10 && antal > 0) { sidor.push([]); antal = 0; }
        sidor[sidor.length - 1].push(g);
        antal += g.rader.length + 1;
      });

      return sidor.map(function (sida) {
        return {
          undertitel: "De närmaste " + dagar + " dagarna · " + lista.length + " st",
          html: '<div class="leveranser">' + sida.map(function (g) {
            var byar = Bod.vader.maxByarArbetstid(ctx.vader, g.datum);
            var harLyft = g.rader.some(function (l) { return l.lyft; });
            var varning = harLyft && byar != null && byar >= grans
              ? '<span class="varning">⚠ Vindbyar upp till ' + Math.round(byar) + " m/s – kontrollera lyft</span>" : "";
            return '<div class="lev-dag"><h3>' + u.relativDag(g.datum, idag) +
              (u.dagarMellan(idag, g.datum) < 2 ? ' <span class="dampad">' + u.veckodag(g.datum) + " " + u.kortDatum(g.datum) + "</span>" : "") +
              varning + "</h3>" + g.rader.map(function (l) {
                return '<div class="lev-rad' + (l.lyft ? " lyft" : "") + '"><div class="lev-tid">' + u.esc(l.tid || "–") + "</div>" +
                  '<div class="lev-vad"><b>' + u.esc(l.vad) + "</b>" + (l.leverantor ? ' <span class="dampad">' + u.esc(l.leverantor) + "</span>" : "") +
                  (l.kommentar ? '<div class="lev-kommentar">' + u.esc(l.kommentar) + "</div>" : "") + "</div>" +
                  '<div class="lev-plats">' + u.esc(l.plats) + "</div>" +
                  '<div class="lev-mottagare">' + u.esc(l.mottagare) + "</div>" +
                  '<div class="lev-lyft">' + (l.lyft ? '<span class="chip lyft-chip">🏗 Kranlyft</span>' : "") + "</div></div>";
              }).join("") + "</div>";
          }).join("") + "</div>"
        };
      });
    }
  });

  /* ------------------------------------------------------ Hinder -- */
  var NIVA = { "hög": 0, "hog": 0, "medel": 1, "låg": 2, "lag": 2 };
  Bod.vyer.registrera({
    id: "hinder",
    titel: "Hinder & risker",
    sidor: function (ctx) {
      var oppna = (ctx.innehall ? ctx.innehall.hinder : [])
        .filter(function (h) { return !/^(löst|klar|stängd)/i.test(h.status); })
        .sort(function (a, b) {
          var na = NIVA[a.niva.toLowerCase()], nb = NIVA[b.niva.toLowerCase()];
          return (na === undefined ? 1 : na) - (nb === undefined ? 1 : nb) || (a.klart || 9e15) - (b.klart || 9e15);
        });
      if (!oppna.length) return [];
      return u.dela(oppna, 6).map(function (sida) {
        return {
          undertitel: oppna.length + " öppna",
          html: '<div class="hinder-rutnat">' + sida.map(function (h) {
            var niva = (h.niva || "Medel").toLowerCase().replace("ö", "o").replace("å", "a");
            var sen = h.klart && h.klart < ctx.idag;
            return '<div class="hinder niva-' + niva + '"><div class="hinder-topp"><span class="chip niva-chip">' + u.esc(h.niva) + "</span>" +
              (h.omrade ? '<span class="dampad">' + u.esc(h.omrade) + "</span>" : "") + '<span class="hinder-status">' + u.esc(h.status) + "</span></div>" +
              "<h3>" + u.esc(h.beskrivning) + "</h3>" +
              (h.konsekvens ? '<p><span class="liten-rubrik">Konsekvens</span>' + u.esc(h.konsekvens) + "</p>" : "") +
              (h.atgard ? '<p><span class="liten-rubrik">Åtgärd</span>' + u.esc(h.atgard) + "</p>" : "") +
              '<div class="hinder-fot">' + (h.ansvarig ? "👤 " + u.esc(h.ansvarig) : "") +
              (h.klart ? '<span class="' + (sen ? "rod" : "") + '">📅 ' + (sen ? "Skulle varit klart " : "Klart senast ") + u.kortDatum(h.klart) + "</span>" : "") +
              "</div></div>";
          }).join("") + "</div>"
        };
      });
    }
  });

  /* ------------------------------------------------ Meddelanden -- */
  Bod.vyer.registrera({
    id: "meddelanden",
    titel: "Meddelanden & säkerhet",
    sidor: function (ctx) {
      var aktiva = (ctx.innehall ? ctx.innehall.meddelanden : [])
        .filter(function (m) { return D.aktiv(m, ctx.idag); });
      var ordning = { "säkerhet": 0, "viktigt": 1, "info": 2 };
      aktiva.sort(function (a, b) {
        var oa = ordning[a.typ.toLowerCase()], ob = ordning[b.typ.toLowerCase()];
        return (oa === undefined ? 3 : oa) - (ob === undefined ? 3 : ob);
      });
      var olycka = u.tillDatum(ctx.inst("Senaste olycka med frånvaro (datum)", ""));
      var malText = u.text(ctx.inst("Säkerhetsbudskap", ""));
      if (!aktiva.length && !olycka && !malText) return [];

      var sido = "";
      if (olycka || malText) {
        var dagar = olycka ? Math.max(0, u.dagarMellan(olycka, ctx.idag)) : null;
        sido = '<aside class="sakerhet">' +
          (dagar !== null ? '<div class="olycka-tal">' + dagar + '</div><div class="olycka-text">dagar utan olycka med frånvaro</div>' : "") +
          (malText ? '<div class="budskap">' + u.escRader(malText) + "</div>" : "") + "</aside>";
      }
      var sidor = aktiva.length ? u.dela(aktiva, sido ? 4 : 6) : [[]];
      return sidor.map(function (sida) {
        return {
          undertitel: aktiva.length ? aktiva.length + " aktuella" : "",
          html: '<div class="meddelande-vy' + (sido ? " med-sido" : "") + '"><div class="meddelanden">' +
            sida.map(function (m) {
              var typ = m.typ.toLowerCase().replace("ä", "a");
              return '<div class="meddelande typ-' + u.esc(typ) + '"><div class="etikett">' + u.esc(m.typ) + "</div>" +
                "<h3>" + u.esc(m.rubrik) + "</h3>" + (m.text ? "<p>" + u.escRader(m.text) + "</p>" : "") + "</div>";
            }).join("") + "</div>" + sido + "</div>"
        };
      });
    }
  });

  /* --------------------------------------------------- Kontakter -- */
  Bod.vyer.registrera({
    id: "kontakter",
    titel: "Kontakter & nödläge",
    sidor: function (ctx) {
      var kontakter = ctx.innehall ? ctx.innehall.kontakter : [];
      var nod = [
        ["Återsamlingsplats", "📍"], ["Hjärtstartare", "❤️"], ["Första hjälpen", "⛑️"],
        ["Ögondusch", "👁️"], ["Brandsläckare", "🧯"], ["Adress för utryckning", "🚑"]
      ].map(function (n) {
        var v = u.text(ctx.inst(n[0], ""));
        return v ? '<div class="nod-rad"><span class="nod-ikon">' + n[1] + '</span><div><div class="liten-rubrik">' + n[0] + "</div><div>" + u.escRader(v) + "</div></div></div>" : "";
      }).join("");
      if (!kontakter.length && !nod) return [];
      return [{
        undertitel: "",
        html: '<div class="kontakt-vy"><div class="kontakter">' + kontakter.slice(0, 12).map(function (k) {
          return '<div class="kontakt"><div class="liten-rubrik">' + u.esc(k.roll) + '</div><div class="kontakt-namn">' + u.esc(k.namn) + "</div>" +
            (k.telefon ? '<div class="kontakt-tel">📞 ' + u.esc(k.telefon) + "</div>" : "") +
            (k.foretag ? '<div class="dampad">' + u.esc(k.foretag) + "</div>" : "") + "</div>";
        }).join("") + "</div>" +
          '<aside class="nodlage"><div class="nod-112"><span>Vid olycka ring</span><b>112</b></div>' + nod + "</aside></div>"
      }];
    }
  });
})();

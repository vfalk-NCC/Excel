/* Tidplansdiagram (Gantt) som används av vyerna "Denna vecka" och "4 veckor framåt". */
(function () {
  "use strict";
  var Bod = (window.Bod = window.Bod || {});
  var u = Bod.util;
  var g = (Bod.gantt = {});

  g.OMRADESFARGER = ["#3da9fc", "#f5a623", "#b48cf2", "#2ec4b6", "#ff7aa2", "#9ccc65", "#ffd166", "#7aa2ff"];
  g.farg = function (omrade) {
    return g.OMRADESFARGER[omrade.index % g.OMRADESFARGER.length];
  };

  // Del/linje, men inte när den bara upprepar områdets namn ("741 - SEKTIONSFICKOR (Hus del 2)").
  g.del = function (a) {
    var d = a.del.toLowerCase();
    if (!d) return "";
    if (a.omrade.kod && d.indexOf(a.omrade.kod) === 0) return "";
    if (d.indexOf(a.omrade.namn.toLowerCase()) >= 0) return "";
    return a.del;
  };

  // Delar upp aktiviteter i sidor med områdesrubriker. En rubrik upprepas
  // ("forts.") när ett område fortsätter på nästa sida.
  g.sidor = function (aktiviteter, raderPerSida) {
    var sidor = [];
    var sida = [];
    var forra = null;
    aktiviteter.forEach(function (a, i) {
      var behov = a.omrade !== forra ? 2 : 1;
      if (sida.length + behov > raderPerSida && sida.length) {
        sidor.push(sida);
        sida = [];
        forra = null;
        if (i > 0 && a.omrade === aktiviteter[i - 1].omrade) {
          sida.push({ grupp: a.omrade, forts: true });
          forra = a.omrade;
        }
      }
      if (a.omrade !== forra) {
        sida.push({ grupp: a.omrade });
        forra = a.omrade;
      }
      sida.push(a);
    });
    if (sida.length) sidor.push(sida);
    return sidor;
  };

  /*
   * opt: { rader, fran, dagar, idag, status(a) }
   */
  g.rita = function (opt) {
    var fran = opt.fran;
    var antal = opt.dagar;
    var till = u.plusDagar(fran, antal - 1);
    var bred = 100 / antal;
    var veckovy = antal <= 7;
    var idagIx = u.dagarMellan(fran, opt.idag);

    // Rubrikrad: veckor + dagar.
    var dagar = [];
    for (var i = 0; i < antal; i++) dagar.push(u.plusDagar(fran, i));

    var veckor = "";
    if (!veckovy) {
      var start = 0;
      for (var j = 1; j <= antal; j++) {
        if (j === antal || dagar[j].getDay() === 1) {
          veckor += '<div class="g-vecka" style="left:' + (start * bred) + "%;width:" + ((j - start) * bred) + '%">' +
            "Vecka " + u.isoVecka(dagar[start]) + "</div>";
          start = j;
        }
      }
    }

    var dagHuvud = dagar.map(function (d, ix) {
      var kl = "g-dag" + (u.helg(d) ? " helg" : "") + (ix === idagIx ? " idag" : "");
      var text = veckovy
        ? '<span class="g-dagnamn">' + u.veckodag(d) + "</span><span>" + u.kortDatum(d) + "</span>"
        : "<span>" + u.veckodagKort(d).charAt(0).toUpperCase() + "</span><span>" + d.getDate() + "</span>";
      return '<div class="' + kl + '">' + text + "</div>";
    }).join("");

    var bakgrund = dagar.map(function (d, ix) {
      return '<div class="g-kol' + (u.helg(d) ? " helg" : "") + (d.getDay() === 1 ? " mandag" : "") +
        (ix === idagIx ? " idag" : "") + '"></div>';
    }).join("");

    var rader = opt.rader.map(function (r) {
      if (r.grupp) {
        return '<div class="g-rad g-grupp" style="--omr:' + g.farg(r.grupp) + '">' +
          '<div class="g-etikett"><span class="omr-kod">' + u.esc(r.grupp.kod) + "</span> " +
          u.esc(r.grupp.namn) + (r.forts ? ' <span class="dampad">(forts.)</span>' : "") + "</div><div></div></div>";
      }
      var a = r;
      var st = opt.status(a);
      var s = Math.max(0, u.dagarMellan(fran, a.start));
      var e = Math.min(antal - 1, u.dagarMellan(fran, a.slut));
      var klippV = a.start < fran;
      var klippH = a.slut > till;
      var stapel = "";
      if (e >= s && e >= 0 && s <= antal - 1) {
        var kl = "g-stapel st-" + st + (a.preliminar ? " prelim" : "") + (klippV ? " klipp-v" : "") + (klippH ? " klipp-h" : "");
        var inre = "";
        if (a.milstolpe) {
          kl += " milstolpe";
        } else {
          if (a.framdrift > 0 && a.framdrift < 1) inre += '<div class="g-framdrift" style="width:' + Math.round(a.framdrift * 100) + '%"></div>';
          var bredd = (e - s + 1) * bred;
          if (veckovy || bredd > 12) {
            inre += '<span class="g-stapeltext">' + (klippV ? "◂ " : "") + u.datumIntervall(a.start, a.slut) + (klippH ? " ▸" : "") + "</span>";
          }
        }
        stapel = '<div class="' + kl + '" style="left:' + (s * bred) + "%;width:" + ((e - s + 1) * bred) + '%">' + inre + "</div>";
      }
      var detaljer = [g.del(a), a.typ, a.ansvarig].filter(Boolean).map(u.esc).join(" · ");
      return '<div class="g-rad" style="--omr:' + g.farg(a.omrade) + '">' +
        '<div class="g-etikett"><div class="g-namn">' + u.esc(a.aktivitet) + "</div>" +
        (detaljer ? '<div class="g-detalj">' + detaljer + "</div>" : "") + "</div>" +
        '<div class="g-tid">' + stapel + "</div>" +
        '<div class="g-status"><span class="chip st-' + st + '">' + Bod.data.STATUS_TEXT[st] +
        (st === "pagar" && a.framdrift > 0 ? " " + u.procent(a.framdrift) : "") + "</span></div></div>";
    }).join("");

    return '<div class="gantt' + (veckovy ? " veckovy" : "") + '">' +
      '<div class="g-huvud"><div class="g-etikett"></div><div class="g-tidhuvud">' +
      (veckor ? '<div class="g-veckor">' + veckor + "</div>" : "") +
      '<div class="g-dagar">' + dagHuvud + '</div></div><div class="g-status"></div></div>' +
      '<div class="g-kropp"><div class="g-bakgrund">' + bakgrund + "</div>" + rader + "</div></div>";
  };
})();

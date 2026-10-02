/* Vädervyn: läget nu, kommande timmar och 7-dygnsprognos med vindvarning för lyft. */
(function () {
  "use strict";
  var Bod = window.Bod;
  var u = Bod.util;
  var V = Bod.vader;

  Bod.vyer.registrera({
    id: "vader",
    titel: "Väder",
    sidor: function (ctx) {
      var w = ctx.vader;
      if (!w || !w.timmar.length) return [];
      var ort = u.text(ctx.inst("Ort", ""));
      var grans = u.tillTal(ctx.inst("Vindgräns lyft (m/s)", 12), 12);
      var nu = w.nu;
      var kod = V.kod(nu.kod);
      var nuTid = ctx.nu.getTime();

      var timmar = w.timmar.filter(function (t) { return t.tid >= nuTid - 3600000; }).slice(0, 12);
      var timHtml = timmar.map(function (t) {
        var d = new Date(t.tid);
        var hog = t.byar != null && t.byar >= grans;
        return '<div class="timme' + (hog ? " blasigt" : "") + '"><div class="t-tid">' + u.tva(d.getHours()) + "</div>" +
          '<div class="t-ikon">' + V.kod(t.kod)[1] + '</div><div class="t-temp">' + V.grader(t.temp) + "</div>" +
          '<div class="t-regn">' + (t.nederbord > 0 ? t.nederbord.toFixed(1) + " mm" : "&nbsp;") + "</div>" +
          '<div class="t-vind">' + Math.round(t.vind) + " <small>(" + Math.round(t.byar) + ")</small></div></div>";
      }).join("");

      var dagHtml = w.dagar.slice(0, 7).map(function (d) {
        var dag = new Date(d.datum);
        var byar = V.maxByarArbetstid(w, dag);
        var hog = byar != null && byar >= grans;
        return '<div class="vdag' + (hog ? " blasigt" : "") + '"><div class="v-dagnamn">' + (u.dagarMellan(ctx.idag, dag) === 0 ? "Idag" : u.veckodag(dag)) + "</div>" +
          '<div class="v-ikon">' + V.kod(d.kod)[1] + '</div><div class="v-temp"><b>' + V.grader(d.max) + "</b> " + V.grader(d.min) + "</div>" +
          '<div class="v-regn">💧 ' + (d.nederbord || 0).toFixed(1) + " mm</div>" +
          '<div class="v-vind">💨 ' + Math.round(d.vind) + " m/s<br><small>byar " + (byar != null ? Math.round(byar) : Math.round(d.byar)) + " m/s</small></div>" +
          (hog ? '<div class="v-varning">⚠ Lyft?</div>' : "") + "</div>";
      }).join("");

      var byarIdag = V.maxByarArbetstid(w, ctx.idag);
      var varning = byarIdag != null && byarIdag >= grans
        ? '<div class="vind-varning">⚠ Vindbyar upp till ' + Math.round(byarIdag) + " m/s under arbetstid idag. Gränsen för lyft är satt till " + grans + " m/s – kontrollera mot lyftplanen.</div>" : "";

      return [{
        undertitel: (ort ? ort + " · " : "") + "uppdaterad " + u.klocka(new Date(w.hamtad)),
        html: '<div class="vader-vy"><div class="vader-nu"><div class="nu-ikon">' + kod[1] + '</div><div><div class="nu-temp">' + V.grader(nu.temp) + "</div>" +
          '<div class="nu-text">' + u.esc(kod[0]) + " · känns som " + V.grader(nu.kanns) + "</div>" +
          '<div class="nu-vind">💨 ' + Math.round(nu.vind) + " m/s " + V.vindriktning(nu.riktning) + " · byar " + Math.round(nu.byar) + " m/s</div></div></div>" +
          varning +
          '<h2>Kommande timmar <span class="dampad">(vind m/s, byar inom parentes)</span></h2><div class="timmar">' + timHtml + "</div>" +
          '<h2>Kommande dagar</h2><div class="vdagar">' + dagHtml + "</div></div>"
      }];
    }
  });
})();

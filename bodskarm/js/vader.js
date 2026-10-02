/*
 * Väder från Open-Meteo (gratis, ingen API-nyckel). Senaste prognosen sparas
 * i webbläsaren så att något visas även om nätet tillfälligt ligger nere.
 */
(function () {
  "use strict";
  var Bod = (window.Bod = window.Bod || {});
  var u = Bod.util;
  var v = (Bod.vader = {});
  var MINNE = "bodskarm.vader";

  // WMO-väderkoder -> svensk text + symbol.
  var KODER = {
    0: ["Klart", "☀️"], 1: ["Mestadels klart", "🌤️"], 2: ["Halvklart", "⛅"], 3: ["Mulet", "☁️"],
    45: ["Dimma", "🌫️"], 48: ["Underkyld dimma", "🌫️"],
    51: ["Lätt duggregn", "🌦️"], 53: ["Duggregn", "🌦️"], 55: ["Kraftigt duggregn", "🌧️"],
    56: ["Underkylt duggregn", "🌧️"], 57: ["Underkylt duggregn", "🌧️"],
    61: ["Lätt regn", "🌦️"], 63: ["Regn", "🌧️"], 65: ["Kraftigt regn", "🌧️"],
    66: ["Underkylt regn", "🌧️"], 67: ["Underkylt regn", "🌧️"],
    71: ["Lätt snöfall", "🌨️"], 73: ["Snöfall", "🌨️"], 75: ["Kraftigt snöfall", "❄️"], 77: ["Kornsnö", "🌨️"],
    80: ["Regnskurar", "🌦️"], 81: ["Regnskurar", "🌧️"], 82: ["Kraftiga skurar", "⛈️"],
    85: ["Snöbyar", "🌨️"], 86: ["Kraftiga snöbyar", "❄️"],
    95: ["Åska", "⛈️"], 96: ["Åska med hagel", "⛈️"], 99: ["Åska med hagel", "⛈️"]
  };
  v.kod = function (k) {
    return KODER[k] || ["", "🌡️"];
  };

  v.hamta = function (lat, lon) {
    var url = "https://api.open-meteo.com/v1/forecast?latitude=" + encodeURIComponent(lat) +
      "&longitude=" + encodeURIComponent(lon) +
      "&current=temperature_2m,apparent_temperature,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m,precipitation" +
      "&hourly=temperature_2m,precipitation,precipitation_probability,weather_code,wind_speed_10m,wind_gusts_10m" +
      "&daily=weather_code,temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max,wind_gusts_10m_max,sunrise,sunset" +
      "&wind_speed_unit=ms&timezone=Europe%2FStockholm&forecast_days=7";
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open("GET", url, true);
      xhr.timeout = 20000;
      xhr.onload = function () {
        if (xhr.status !== 200) return reject(new Error("Väder: status " + xhr.status));
        try {
          var data = tolka(JSON.parse(xhr.responseText));
          try { localStorage.setItem(MINNE, JSON.stringify({ lat: lat, lon: lon, data: data })); } catch (e) { /* ingen lagring */ }
          resolve(data);
        } catch (e) {
          reject(e);
        }
      };
      xhr.onerror = xhr.ontimeout = function () { reject(new Error("Väder: inget svar")); };
      xhr.send();
    });
  };

  v.senaste = function (lat, lon) {
    try {
      var s = JSON.parse(localStorage.getItem(MINNE) || "null");
      if (s && s.lat == lat && s.lon == lon) return s.data;
    } catch (e) { /* ignorera */ }
    return null;
  };

  function tid(s) {
    // "2026-10-02T14:00" tolkas som lokal tid.
    var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : null;
  }

  function tolka(j) {
    var c = j.current || {};
    var h = j.hourly || {};
    var dd = j.daily || {};
    return {
      hamtad: Date.now(),
      nu: {
        temp: c.temperature_2m, kanns: c.apparent_temperature, kod: c.weather_code,
        vind: c.wind_speed_10m, byar: c.wind_gusts_10m, riktning: c.wind_direction_10m, nederbord: c.precipitation
      },
      timmar: (h.time || []).map(function (t, i) {
        return {
          tid: tid(t).getTime(), temp: h.temperature_2m[i], nederbord: h.precipitation[i],
          risk: h.precipitation_probability ? h.precipitation_probability[i] : null,
          kod: h.weather_code[i], vind: h.wind_speed_10m[i], byar: h.wind_gusts_10m[i]
        };
      }),
      dagar: (dd.time || []).map(function (t, i) {
        return {
          datum: tid(t).getTime(), kod: dd.weather_code[i], max: dd.temperature_2m_max[i], min: dd.temperature_2m_min[i],
          nederbord: dd.precipitation_sum[i], vind: dd.wind_speed_10m_max[i], byar: dd.wind_gusts_10m_max[i],
          upp: dd.sunrise ? dd.sunrise[i] : null, ner: dd.sunset ? dd.sunset[i] : null
        };
      })
    };
  }

  // Högsta vindbyar under arbetstid (06–18) en viss dag.
  v.maxByarArbetstid = function (data, dag) {
    if (!data) return null;
    var max = null;
    data.timmar.forEach(function (t) {
      var d = new Date(t.tid);
      if (u.dagarMellan(dag, d) === 0 && d.getHours() >= 6 && d.getHours() <= 18 && t.byar != null) {
        max = max === null ? t.byar : Math.max(max, t.byar);
      }
    });
    return max;
  };

  v.vindriktning = function (grader) {
    if (grader == null) return "";
    var r = ["N", "NO", "O", "SO", "S", "SV", "V", "NV"];
    return r[Math.round(grader / 45) % 8];
  };

  v.grader = function (t) {
    return t == null ? "–" : Math.round(t) + "°";
  };
})();

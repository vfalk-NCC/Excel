/*
 * BODSKÄRM – tekniska inställningar
 * ---------------------------------
 * Det mesta du vill ändra till vardags (meddelanden, bilder, vilka vyer som
 * visas, visningstider, väderort ...) ändrar du i Excel-filen
 * data/Bodskarm_Innehall.xlsx. Den här filen behöver du bara röra när du
 * flyttar filer eller vill finjustera tekniken.
 *
 * Sökvägar kan vara relativa till den här mappen (t.ex. "data/Plan.xlsm")
 * eller absoluta (t.ex. "C:/Users/bod/OneDrive - Företag/NSV/Plan.xlsm").
 */
window.BODSKARM_CONFIG = {
  // Din 4-veckorsplanering (samma fil som du jobbar i till vardags).
  planFil: "../source/NSV_Bygg_4veckorsplanering_ORIGINAL.xlsm",

  // Excel-filen med meddelanden, hinder, leveranser, bilder, kontakter m.m.
  innehallFil: "data/Bodskarm_Innehall.xlsx",

  // Mapp där bilderna (Trimble-vyer, APD-plan ...) ligger.
  bildMapp: "data/bilder/",

  // Vilka blad i planeringsfilen som är områdesblad.
  // Tom lista = hitta automatiskt alla blad med kolumnerna "Aktivitet" och
  // "Datum start" (blad vars namn innehåller något i planBladUndanta hoppas över).
  planBlad: [],
  planBladUndanta: ["RESURS", "DP1 - Kvarstående", "AKT", "LISTOR"],

  // Hur ofta Excel-filerna läses in på nytt (minuter).
  uppdateraMinuter: 2,

  // Hur ofta vädret hämtas (minuter).
  vaderMinuter: 30,

  // Klockslag då hela sidan laddas om (rensar minnet, tar in kodändringar).
  omstartKlockan: "04:00",

  // Antal rader per sida i tidplansvyerna innan nästa sida tar vid.
  raderPerSida: 13,

  // Antal dagar framåt i "Kommande"- och "Leveranser"-vyerna.
  dagarFramat: 14,

  // Visa aktiviteter vars slutdatum passerat utan 100 % framdrift som "Försenad".
  // Sätt till false om ni inte rapporterar framdrift i planeringen.
  visaForsenade: true
};

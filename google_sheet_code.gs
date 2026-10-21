// Poznatky:
// - ak sa nasadzuje upravena verzia kodu, vlozit novy kod, ulozit (ctrl+s), staci kliknut na manage deployments, 
//   tam kliknut na ceruzku, v prvom dropdowne kliknut na new version a nasledne pridat popis (volitelne) a dat save/deploy

// TODO (resp. napady na rozsirenie):
// - moznost zmenit termin (zakaznik) (FUH, MAXIMALNE SA MI ZATIAL NECHCE)

const SHEET_NAME = 'Rezervácie';
const SHEET_CUSTOM = 'Otv. hodiny mimo bežné';
const SHEET_DEFAULT = 'Bežné otv. hodiny';
const EMAIL_BARBAR = 'matusjacko1@gmail.com';
const MIESTO = 'B02/321, Purkyňovy koleje'; // kde sa strihá: len v mailoch a na stránke potvrdenia (odkaz z mailu), zámerne NIE na verejnom webe
const MAIL_NAZOV = 'Barbar Shop'; // meno odosielateľa, ktoré uvidia príjemcovia vo svojej schránke

// --- VALIDÁCIA VSTUPU ---
// doPost je verejné API (URL je vidieť v script.js), takže klientská validácia v prehliadači sa dá
// jednoducho obísť priamym volaním endpointu. Server si preto musí všetko overiť sám.
const ALLOWED_SLUZBY = ['Pánsky strih', 'Úprava brady', 'Vlasy a brada'];
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const DATE_REGEX = /^\d{4}-\d{2}-\d{2}$/;
const TIME_REGEX = /^([01]\d|2[0-3]):([0-5]\d)$/;
const MIN_MINUT_VOPRED = 30;
const RATE_LIMIT_SEKUND = 30;
const DNI_MAPA = {'Pondelok':1,'Utorok':2,'Streda':3,'Štvrtok':4,'Piatok':5,'Sobota':6,'Nedeľa':7};

// --- PÍSMO PRE STRÁNKY Z APPS SCRIPTU ---
// Stránky zrušenia/detailu sa servírujú z domény Googlu, takže nevidia súbory z projektu - písmo
// Roboto si preto berú z Vercelu (fonts/ v repozitári, CORS povolený vo vercel.json). Ak sa nenačíta,
// ostane záložné systémové písmo, nič sa nerozbije.
const FONT_BASE_URL = 'https://barbarshop-mu.vercel.app/fonts/';
function fontFaceCss() {
  const rozsahy = {
    'latin': 'U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD',
    'latin-ext': 'U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF'
  };
  let css = '';
  [400, 700].forEach(function(w) {
    Object.keys(rozsahy).forEach(function(subset) {
      css += "@font-face{font-family:'Roboto';font-style:normal;font-weight:" + w + ";font-display:swap;" +
             "src:url('" + FONT_BASE_URL + 'roboto-' + subset + '-' + w + "-normal.woff2') format('woff2');" +
             'unicode-range:' + rozsahy[subset] + ';}';
    });
  });
  return css;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml(str) {
  return String(str == null ? '' : str).replace(/[&<>"']/g, function(c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

// Bezpečné vloženie JSON do <script> bloku v HTML šablóne (ochrana pred "</script>" injekciou).
function jsonForScript(obj) {
  return JSON.stringify(obj)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
}

// Deň v týždni (1=Pondelok...7=Nedeľa) vypočítaný čisto z čísel dátumu, bez new Date().getDay() -
// ten je závislý od časovej zóny nastavenej v projekte Apps Scriptu a pri zlom nastavení by tesne
// okolo polnoci mohol vrátiť nesprávny deň (a teda zlý rozvrh otváracích hodín).
function getIsoDayOfWeek(ymd) {
  const parts = String(ymd).split('-').map(Number);
  let y = parts[0], m = parts[1], d = parts[2];
  const t = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
  if (m < 3) y -= 1;
  const dow = (y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) + t[m - 1] + d) % 7;
  return dow === 0 ? 7 : dow;
}

// Načíta pravidlá otváracích hodín (vlastné + bežné) do formátu, ktorý používa getValidSlotsForDay.
// Zdieľané medzi doPost a handleEdit, aby obe miesta používali presne rovnakú logiku.
function buildScheduleRules(sheetCustom, sheetDefault) {
  let customRules = [];
  if (sheetCustom) {
    const custData = sheetCustom.getDataRange().getDisplayValues();
    for (let i = 1; i < custData.length; i++) {
      let rDatum = custData[i][0];
      let rOd = custData[i][1];
      let rDo = custData[i][2];
      let rStav = custData[i][3] ? custData[i][3].trim().toUpperCase() : "";
      if (rDatum && rOd && rDo && rStav) {
        if (rStav !== 'ZAVRETÉ') {
          rOd = formatToGridStart(rOd);
          rDo = formatToGridEnd(rDo);
        }
        let expanded = expandDates(rDatum);
        expanded.forEach(d => {
          customRules.push({ datum: d, odMin: parseTime(rOd), doMin: parseTime(rDo), stav: rStav });
        });
      }
    }
  }

  let defaultHoursMap = {};
  if (sheetDefault) {
    const defaultData = sheetDefault.getDataRange().getDisplayValues();
    for (let i = 1; i < defaultData.length; i++) {
      let denCislo = DNI_MAPA[defaultData[i][0].trim()];
      if (denCislo) {
        let bloky = [];
        for (let j = 1; j < defaultData[i].length; j += 2) {
          if (defaultData[i][j] && defaultData[i][j + 1]) {
            bloky.push({ odMin: parseTime(formatToGridStart(defaultData[i][j])), doMin: parseTime(formatToGridEnd(defaultData[i][j + 1])) });
          }
        }
        defaultHoursMap[denCislo] = bloky;
      }
    }
  }

  return { customRules: customRules, defaultHoursMap: defaultHoursMap };
}

function isSlotAvailable(datumYMD, casHM) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rules = buildScheduleRules(ss.getSheetByName(SHEET_CUSTOM), ss.getSheetByName(SHEET_DEFAULT));
  const validSlots = getValidSlotsForDay(datumYMD, rules.customRules, rules.defaultHoursMap);
  return validSlots.indexOf(parseTime(casHM)) !== -1;
}

// Jednoduchá ochrana pred spamom/duplicitným odoslaním: z toho istého e-mailu neprejde
// druhá rezervácia skôr, než o RATE_LIMIT_SEKUND sekúnd.
function checkRateLimit(email) {
  const cache = CacheService.getScriptCache();
  const key = 'rl_' + email.toLowerCase();
  if (cache.get(key)) return false;
  cache.put(key, '1', RATE_LIMIT_SEKUND);
  return true;
}

// Overí všetky dáta z formulára. Klientská validácia v script.js je len pre pohodlie používateľa,
// dá sa jednoducho obísť priamym volaním API - túto validáciu preto nesmieme vynechať.
function validateBookingData(data) {
  if (!data || typeof data !== 'object') return 'Neplatné dáta.';

  const meno = String(data.meno || '').trim();
  if (meno.length < 2 || meno.length > 100) return 'Meno musí mať 2 až 100 znakov.';

  const email = String(data.email || '').trim();
  if (email.length > 200 || !EMAIL_REGEX.test(email)) return 'Zadaj platný e-mail.';

  if (ALLOWED_SLUZBY.indexOf(data.sluzba) === -1) return 'Neplatná služba.';

  if (!DATE_REGEX.test(data.datum) || !TIME_REGEX.test(data.cas)) return 'Neplatný formát termínu.';

  const dateParts = data.datum.split('-').map(Number);
  const y = dateParts[0], m = dateParts[1], d = dateParts[2];
  if (m < 1 || m > 12 || d < 1 || d > 31) return 'Neplatný dátum.';

  const timeParts = data.cas.split(':').map(Number);
  const slotStart = new Date(y, m - 1, d, timeParts[0], timeParts[1], 0);
  if (isNaN(slotStart.getTime())) return 'Neplatný dátum.';

  // "now" prevedené na bratislavský čas a spätne sparsované rovnakým spôsobom (z lokálnych čísel)
  // ako slotStart vyššie - vďaka tomu vyjde rozdiel v minútach správne bez ohľadu na to,
  // aká časová zóna je nastavená v samotnom projekte Apps Scriptu.
  const nowBaStr = Utilities.formatDate(new Date(), 'Europe/Bratislava', 'yyyy/MM/dd HH:mm:ss');
  const now = new Date(nowBaStr);
  const diffMin = (slotStart.getTime() - now.getTime()) / 60000;
  if (diffMin < MIN_MINUT_VOPRED) return `Termín je možné rezervovať minimálne ${MIN_MINUT_VOPRED} minút vopred.`;

  if (!isSlotAvailable(data.datum, data.cas)) return 'Tento termín už nie je dostupný. Obnov si stránku a vyber iný.';

  return null;
}

function skontrolujKvotu() {
  var zostavajucaKvota = MailApp.getRemainingDailyQuota();
  
  // Vypíše to do logu v editore
  console.log("Zostávajúci počet e-mailov na dnes: " + zostavajucaKvota);
  
  // Zobrazí vyskakovaciu bublinu priamo v tvojej Google Tabuľke
  SpreadsheetApp.getActiveSpreadsheet().toast("Na dnes ti zostáva " + zostavajucaKvota + " e-mailov.", "Dostupná kvóta", 10);
}

// --- ŠETRENIE DENNOU KVÓTOU MAILOV ---
// Keď ostáva menej ako KVOTA_REZERVA mailov, posielajú sa už len maily, bez ktorých to nejde (potvrdenie termínu,
// zrušenie zákazníkovi, upozornenie na konflikt v rozvrhu). Maily holičovi (nová rezervácia, zrušenie, nová
// recenzia) sa odložia do hárku "Odložené pre holiča" a o SUHRN_HODINA mu príde jeden súhrnný mail.
// Prosby o recenziu počkajú na ďalší deň (sendReviewRequests). Keď je kvóty dosť, všetko chodí ako doteraz.
const KVOTA_REZERVA = 20;
const SUHRN_HODINA = 20;
const SHEET_SUHRN = 'Odložené pre holiča';
const SUHRN_HLAVICKA = ['Čas', 'Typ', 'Text', 'Odkazy (JSON)'];

function mailQuotaLow_() {
  try {
    return MailApp.getRemainingDailyQuota() < KVOTA_REZERVA;
  } catch (err) {
    return false; // keď sa kvóta nedá zistiť, správame sa ako doteraz
  }
}

// Mail holičovi, ktorý môže počkať. polozka = {typ, text, odkazy: [{t, u}]} pre večerný súhrn. Nikdy nehádže chybu.
function sendBarberMailOrDefer_(mail, polozka) {
  if (!mailQuotaLow_()) {
    try {
      MailApp.sendEmail(Object.assign({ name: MAIL_NAZOV, to: EMAIL_BARBAR }, mail));
      return 'odoslane';
    } catch (err) {
      console.error('Mail holičovi neodišiel, ide do večerného súhrnu:', err);
    }
  }
  try {
    getDigestSheet_().appendRow([
      Utilities.formatDate(new Date(), 'Europe/Bratislava', 'd.M.yyyy HH:mm'),
      polozka.typ,
      safeCell(polozka.text),
      JSON.stringify(polozka.odkazy || [])
    ]);
  } catch (err) {
    console.error('Položku do súhrnu sa nepodarilo uložiť:', err);
  }
  return 'odlozene';
}

function getDigestSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SUHRN);
  if (!sheet) {
    const povodny = ss.getActiveSheet();
    sheet = ss.insertSheet(SHEET_SUHRN);
    sheet.getRange(1, 1, 1, SUHRN_HLAVICKA.length).setValues([SUHRN_HLAVICKA]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    try { ss.setActiveSheet(povodny); } catch (err) {}
  }
  return sheet;
}

// Denný trigger o SUHRN_HODINA (installSpeedTriggers). Ak nič odložené nie je, nepošle nič.
function sendBarberDigest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUHRN);
  if (!sheet || sheet.getLastRow() < 2) return;
  let kvota = 1;
  try { kvota = MailApp.getRemainingDailyQuota(); } catch (err) {}
  if (kvota < 1) {
    console.log('Súhrn pre holiča počká na zajtra, kvóta je vyčerpaná.');
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const n = sheet.getLastRow() - 1;
    const rows = sheet.getRange(2, 1, n, SUHRN_HLAVICKA.length).getValues();
    MailApp.sendEmail({ name: MAIL_NAZOV, to: EMAIL_BARBAR,
      subject: 'Súhrn dňa - Barbar Shop (' + rows.length + ')',
      htmlBody: generateDigestEmailHtml_(rows, kvota)
    });
    sheet.deleteRows(2, n); // nové položky pribudnuté medzitým sú až za týmito riadkami a ostanú na zajtra
  } catch (err) {
    console.error('Súhrn pre holiča sa nepodarilo odoslať (položky ostávajú):', err);
  } finally {
    lock.releaseLock();
  }
}

function generateDigestEmailHtml_(rows, kvota) {
  const skupiny = {};
  rows.forEach(function(r) {
    const typ = String(r[1] || 'Ostatné');
    (skupiny[typ] = skupiny[typ] || []).push(r);
  });
  let telo = '';
  Object.keys(skupiny).forEach(function(typ) {
    telo += '<h2 style="color:#1a1a1a;font-size:18px;margin:26px 0 10px 0;border-bottom:2px solid #f0c419;padding-bottom:6px;">' + escapeHtml(typ) + ' (' + skupiny[typ].length + ')</h2>';
    skupiny[typ].forEach(function(r) {
      let odkazy = [];
      try { odkazy = JSON.parse(r[3] || '[]'); } catch (err) {}
      const linky = odkazy.filter(function(o) { return o && /^https:\/\//.test(o.u); }).map(function(o) {
        return '<a href="' + escapeHtml(o.u) + '" target="_blank" style="color:#d32f2f;font-weight:bold;text-decoration:none;margin-right:14px;">' + escapeHtml(o.t) + '</a>';
      }).join('');
      telo += '<div style="padding:10px 0;border-bottom:1px solid #eeeeee;">' +
        '<div style="font-size:12px;color:#999999;">' + escapeHtml(String(r[0])) + '</div>' +
        '<div style="font-size:15px;color:#1a1a1a;margin:2px 0 4px 0;">' + escapeHtml(String(r[2]).replace(/^'/, '')) + '</div>' +
        (linky ? '<div style="font-size:14px;">' + linky + '</div>' : '') + '</div>';
    });
  });
  return '<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background-color:#f4f4f5;font-family:Arial,sans-serif;">' +
    '<table border="0" cellpadding="0" cellspacing="0" width="100%" style="padding:20px 0;"><tr><td align="center">' +
    '<table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color:#ffffff;border-radius:12px;overflow:hidden;border-collapse:separate;max-width:100%;">' +
    '<tr><td align="center" style="background-color:#1a1a1a;padding:24px;border-bottom:4px solid #f0c419;"><img src="https://raw.githubusercontent.com/Neonka-Svk/barbarshop/refs/heads/main/barbar_logo_small.png" alt="Barbar Shop" width="130" style="display:block;max-width:130px;"></td></tr>' +
    '<tr><td style="padding:30px 36px;color:#333333;line-height:1.5;font-size:15px;">' +
    '<h1 style="color:#1a1a1a;font-size:22px;margin:0 0 10px 0;text-align:center;">Súhrn dňa</h1>' +
    '<p style="margin:0;padding:12px 14px;background-color:#fff8e1;border-left:4px solid #f0c419;border-radius:6px;font-size:14px;">Dnes ostávalo málo mailov (denná kvóta Gmailu), preto ti tieto upozornenia chodia naraz. Zákazníkom maily odchádzali normálne. Zostáva ešte <strong>' + kvota + '</strong> mailov.</p>' +
    telo +
    '</td></tr>' +
    '<tr><td align="center" style="background-color:#f9f9f9;padding:20px;border-top:1px solid #eaeaea;font-size:12px;color:#999999;">Tento e-mail bol vygenerovaný automaticky systémom Barbar Shop. &copy; 2026 Barbar Shop. Sila a česť.</td></tr>' +
    '</table></td></tr></table></body></html>';
}

function updateWebAppUrl() {
  // Volá sa na začiatku doPost/doGet MIMO ich try/catch - keby tu niečo hodilo chybu (napr. dočasný
  // výpadok PropertiesService), celý request by spadol a namiesto JSON-u by prišla HTML chybová stránka.
  // Preto si chybu ticho zalogujeme a ideme ďalej - toto je len pomocná "housekeeping" funkcia,
  // nesmie zablokovať skutočnú odpoveď pre zákazníka.
  try {
    const url = ScriptApp.getService().getUrl();
    // Uložíme to len ak adresa vyzerá ako ostré nasadenie (obsahuje /exec)
    if (url && url.indexOf('/exec') !== -1) {
      // Zapisuje sa len pri zmene: čítanie je rýchle, zápis sa pri každej návšteve stránky platiť nemusí.
      const props = PropertiesService.getScriptProperties();
      if (props.getProperty('WEB_APP_URL') !== url) props.setProperty('WEB_APP_URL', url);
    }
  } catch (err) {
    console.error('updateWebAppUrl zlyhalo (ignorujeme, nie je to kritické):', err);
  }
}

// --- POMOCNÉ FUNKCIE PRE DÁTUMY A ZAOKRÚHĽOVANIE ---
// Bezpečnostný strop na počet dní, ktoré vieme "rozbaliť" z jedného rozsahu (cca 3 roky).
// Bez neho by preklep v dátume (napr. zlý rok) mohol nechať skript behať deň po dni cez
// tisícky dní a presne to spôsobilo "Exceeded maximum execution time" pri checkUpcomingConflicts.
const MAX_EXPAND_DNI = 1100;

function expandDates(dateInput) {
  let results = [];
  if (!dateInput) return results;
  let parts = String(dateInput).split(',');

  parts.forEach(part => {
    part = part.trim();
    if (part.includes('-')) {
      let rangeParts = part.split('-').map(p => p.trim());
      if (rangeParts.length === 2) {
        let start = parseSlovakDate(rangeParts[0]);
        let end = parseSlovakDate(rangeParts[1]);
        if (start && end) {
          let current = new Date(start);
          let pocet = 0;
          while (current <= end && pocet < MAX_EXPAND_DNI) {
            results.push(`${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, '0')}-${String(current.getDate()).padStart(2, '0')}`);
            current.setDate(current.getDate() + 1);
            pocet++;
          }
          if (pocet >= MAX_EXPAND_DNI) {
            console.error('expandDates: rozsah "' + part + '" je podozrivo široký (' + MAX_EXPAND_DNI + '+ dní) - skontroluj, či v hárku "Otv. hodiny mimo bežné" nie je preklep v dátume.');
          }
        }
      }
    } else {
      let d = parseSlovakDate(part);
      if (d) {
        results.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
      }
    }
  });
  return results;
}

// Rýchla kontrola "je targetYMD súčasťou tohto zadania dátumov?" BEZ toho, aby sa musel celý
// rozsah rozbaliť deň po dni - stačia 2 porovnania na riadok, nezávisle od šírky rozsahu.
// Používa checkUpcomingConflicts, kde nás aj tak zaujímajú len 2 konkrétne dni (dnes/zajtra).
function dateInInput(targetYMD, dateInput) {
  if (!dateInput) return false;
  const targetParts = String(targetYMD).split('-').map(Number);
  const target = new Date(targetParts[0], targetParts[1] - 1, targetParts[2]).getTime();

  let parts = String(dateInput).split(',');
  for (let i = 0; i < parts.length; i++) {
    let part = parts[i].trim();
    if (part.includes('-')) {
      let rangeParts = part.split('-').map(p => p.trim());
      if (rangeParts.length === 2) {
        let start = parseSlovakDate(rangeParts[0]);
        let end = parseSlovakDate(rangeParts[1]);
        if (start && end && target >= start.getTime() && target <= end.getTime()) return true;
      }
    } else {
      let d = parseSlovakDate(part);
      if (d && d.getTime() === target) return true;
    }
  }
  return false;
}

function parseSlovakDate(str) {
  let p = str.replace(/\s/g, '').split('.');
  if (p.length === 3) {
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return null;
}

function parseTime(timeStr) {
  if (!timeStr) return 0;
  let parts = String(timeStr).split(':');
  if (parts.length >= 2) {
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  }
  return 0;
}

function formatToGridStart(timeStr) {
  if (!timeStr) return "";
  let parts = String(timeStr).split(':');
  if (parts.length >= 2) {
    let h = parseInt(parts[0], 10);
    let m = parseInt(parts[1], 10);
    if (m === 0) {
      m = 0;
    } else if (m > 0 && m <= 30) {
      m = 30;
    } else {
      m = 0;
      h = (h + 1) % 24;
    }
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }
  return String(timeStr).trim();
}

function formatToGridEnd(timeStr) {
  if (!timeStr) return "";
  let parts = String(timeStr).split(':');
  if (parts.length >= 2) {
    let h = parseInt(parts[0], 10);
    let m = parseInt(parts[1], 10);
    if (m >= 30) {
      m = 30;
    } else {
      m = 0;
    }
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }
  return String(timeStr).trim();
}

// NOVÁ FUNKCIA: Presný výpočet mriežky na pozadí (Zjednotené s webom)
// Odpočíta z rozsahu `range` ({start,end} v minútach) všetky rozsahy z poľa `occupied`
// a vráti zvyšné (neprekrývajúce sa) kusy. Používa sa na to, aby ROZŠÍRENÉ generovalo sloty
// len z toho, čo bežné/OTVORENÉ hodiny ešte nepokrývajú (inak by pri prekryve s iným rastrom
// vznikli poprehadzované/preskočené okná).
function subtractRanges(range, occupied) {
  let pieces = [range];
  occupied.forEach(o => {
    let next = [];
    pieces.forEach(p => {
      if (o.end <= p.start || o.start >= p.end) { next.push(p); return; }
      if (o.start > p.start) next.push({ start: p.start, end: Math.min(o.start, p.end) });
      if (o.end < p.end) next.push({ start: Math.max(o.end, p.start), end: p.end });
    });
    pieces = next;
  });
  return pieces.filter(p => p.end > p.start);
}

// Poskladá 90-minútové sloty z poľa rozsahov {start,end}, každý rozsah od svojho vlastného začiatku.
function slotsFromRanges(ranges) {
  let out = [];
  ranges.forEach(r => {
    let start = r.start;
    while (start + 90 <= r.end) { out.push(start); start += 90; }
  });
  return out;
}

function toRange(r) {
  return { start: r.odMin, end: r.doMin < r.odMin ? r.doMin + 1440 : r.doMin };
}

function getValidSlotsForDay(targetYMD, customRules, defaultHoursMap) {
  let slots = [];
  let pravidlaPreDen = customRules.filter(r => r.datum === targetYMD);
  let otvorene = pravidlaPreDen.filter(r => r.stav === 'OTVORENÉ');
  let rozsirene = pravidlaPreDen.filter(r => r.stav === 'ROZŠÍRENÉ');
  let zavrete = pravidlaPreDen.filter(r => r.stav === 'ZAVRETÉ');

  if (otvorene.length > 0) {
    let otvoreneRanges = otvorene.map(toRange);
    slots = slots.concat(slotsFromRanges(otvoreneRanges));

    // ROZŠÍRENÉ pridáva čas NAVYŠE k OTVORENÉ - časť, ktorá sa s ním prekrýva, je už zarátaná
    // vyššie, takže tu spracujeme len tie kusy, čo ležia mimo OTVORENÉ rozsahu.
    rozsirene.forEach(r => {
      let zvysne = subtractRanges(toRange(r), otvoreneRanges);
      slots = slots.concat(slotsFromRanges(zvysne));
    });
  } else {
    let dayOfWeek = getIsoDayOfWeek(targetYMD);
    let defBloky = defaultHoursMap[dayOfWeek] || [];
    let defRanges = defBloky.map(toRange);

    slots = slots.concat(slotsFromRanges(defRanges));

    // ROZŠÍRENÉ pridáva čas NAVYŠE k bežným hodinám - časť, ktorá sa s nimi prekrýva, je už
    // zarátaná vyššie, takže tu spracujeme len tie kusy rozsahu, čo ležia mimo bežných hodín.
    rozsirene.forEach(r => {
      let zvysne = subtractRanges(toRange(r), defRanges);
      slots = slots.concat(slotsFromRanges(zvysne));
    });
  }

  zavrete.forEach(r => {
    let odMin = r.odMin;
    let doMin = r.doMin < r.odMin ? r.doMin + 1440 : r.doMin;
    if (odMin === 0 && doMin === 1440) {
      slots = [];
    } else {
      slots = slots.filter(sMin => {
        let eMin = sMin + 90;
        return (eMin <= odMin || sMin >= doMin);
      });
    }
  });

  return [...new Set(slots)];
}

// --- HLAVNÉ API PRE WEB ---
function doPost(e) {
  updateWebAppUrl();

  try {
    const data = JSON.parse(e.postData.contents);

    // Akcie nad čakajúcou rezerváciou z panela po odoslaní objednávky (odoslať mail, zmeniť e-mail, zrušiť, stav).
    if (data && data.action) {
      return jsonOut(handlePendingAction_(data));
    }

    // Honeypot - skryté pole vo formulári, ktoré vidia iba boti. Ak je vyplnené,
    // len predstierame úspech a nič nerobíme (nezapisujeme, neposielame maily).
    if (data.website) {
      return jsonOut({status: 'success'});
    }

    const validationError = validateBookingData(data);
    if (validationError) {
      return jsonOut({error: validationError});
    }

    const meno = String(data.meno).trim();
    const email = String(data.email).trim();
    const sluzba = data.sluzba;

    if (!checkRateLimit(email)) {
      return jsonOut({error: 'Prosím počkaj pár sekúnd pred ďalším odoslaním.'});
    }

    const [y, m, d] = data.datum.split('-');
    const slovakDate = `${d}.${m}.${y}`;
    const id = Utilities.getUuid();
    // Tajný kľúč pre prehliadač, ktorý rezerváciu vytvoril (zmena e-mailu, zrušenie, stav). Do tabuľky
    // ide len jeho odtlačok a nikam sa neposiela mailom.
    const kluc = Utilities.getUuid() + Utilities.getUuid();
    const teraz = Date.now();

    // Zámok proti súbežnému zápisu - bez neho by mohli dvaja zákazníci naraz
    // zarezervovať presne ten istý termín (obaja by prešli kontrolou obsadenosti skôr,
    // než by ktorýkoľvek z nich stihol zapísať svoj riadok).
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      return jsonOut({error: 'Server je momentálne vyťažený, skús to prosím o chvíľu znova.'});
    }

    try {
      if (isSlotTaken_(slovakDate, data.cas, teraz)) {
        return jsonOut({error: 'Tento termín si medzitým niekto zarezervoval. Vyber si prosím iný.'});
      }

      // Rezervácia najprv čaká v hárku "Nepotvrdené" (termín je pre ostatných obsadený). Mail s odkazom
      // na potvrdenie odíde až po krátkom odpočte na stránke, počas ktorého si zákazník môže opraviť e-mail.
      getPendingSheet_().appendRow([
        slovakDate, data.cas, new Date(teraz), safeCell(meno), sluzba, safeCell(email), id,
        CAKAJUCE_STAV_NEODOSLANE, teraz + POTVRDENIE_ODKLAD_SEKUND * 1000, '', Utilities.getUuid(), hashKey_(kluc), 0, false
      ]);
      const cakajuce = getPendingSheet_();
      cakajuce.getRange(cakajuce.getLastRow(), CAKAJUCE_STLPEC_POTVRDIT).insertCheckboxes();
      PropertiesService.getScriptProperties().setProperty('PENDING_ANY', '1');
      invalidatePublicCache_(); // kalendár na webe má hneď ukázať držaný termín
    } finally {
      lock.releaseLock();
    }

    return jsonOut({
      status: 'pending',
      id: id,
      key: kluc,
      email: email,
      sendInMs: POTVRDENIE_ODKLAD_SEKUND * 1000,
      holdMinutes: POTVRDENIE_LEHOTA_MINUT,
      changesLeft: POTVRDENIE_MAX_ZMIEN
    });
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

function doGet(e) {
  const action = (e && e.parameter) ? e.parameter.action : null;
  const id = (e && e.parameter) ? e.parameter.id : null;
  const role = (e && e.parameter) ? e.parameter.role : 'klient';

  // Najčastejšia požiadavka (načítanie stránky) ide rýchlou cestou: hotová odpoveď z cache bez otvárania tabuľky.
  if (!action) return publicDataResponse_();

  updateWebAppUrl();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRez = ss.getSheetByName(SHEET_NAME);
  
  const htmlStart = "<style>" + fontFaceCss() + "</style><div style='font-family:Roboto, Helvetica, Arial, sans-serif; background:#1a1a1a; color:#f0c419; height:100vh; display:flex; justify-content:center; align-items:center; text-align:center; margin:0;'><h2>";
  const htmlEnd = "</h2></div>";

  // --- AKCIA 1: ZOBRAZENIE STRÁNKY NA ZADANIE DÔVODU ---
  if (action === 'cancelPage' && id) {
    const data = sheetRez.getDataRange().getDisplayValues(); 
    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === id) { 
        const template = HtmlService.createTemplateFromFile('CancelPage');
        template.id = id;
        template.role = role;
        template.meno = data[i][3];
        template.datum = data[i][0]; 
        template.cas = data[i][1].replace(/^'/, '');
        template.scriptUrl = ScriptApp.getService().getUrl();
        template.fontFaceCss = fontFaceCss();
        template.busyCss = BUSY_BUTTON_CSS;
        template.busyJs = BUSY_BUTTON_JS;
        return template.evaluate().setTitle("Zrušenie rezervácie").addMetaTag('viewport', 'width=device-width, initial-scale=1');
      }
    }
    // Ešte nepotvrdená rezervácia (odkaz "Zrušiť termín" je už v maile s potvrdením).
    const pend = findPending_(id);
    if (pend.p && isPendingActive_(pend.p.r, Date.now())) {
      const template = HtmlService.createTemplateFromFile('CancelPage');
      template.id = id;
      template.role = 'klient';
      template.meno = plain_(pend.p.r[3]);
      template.datum = plain_(pend.p.r[0]);
      template.cas = plain_(pend.p.r[1], 'HH:mm');
      template.scriptUrl = ScriptApp.getService().getUrl();
      template.fontFaceCss = fontFaceCss();
      template.busyCss = BUSY_BUTTON_CSS;
      template.busyJs = BUSY_BUTTON_JS;
      return template.evaluate().setTitle("Zrušenie rezervácie").addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    return HtmlService.createHtmlOutput(htmlStart + "Rezervácia nebola nájdená!" + htmlEnd);
  }

  // --- AKCIA 2: FYZICKÉ ZMAZANIE A ODOSLANIE HTML E-MAILU ---
  if (action === 'executeCancel' && id) {
    const dovodInput = (e && e.parameter.dovod) ? e.parameter.dovod.trim() : "";
    let termiZruseny = false;
    let nepotvrdena = false;

    // Pod zámkom: riadky sa mažú podľa čísla a potvrdenie/nová rezervácia ich medzitým môže preusporiadať.
    const lockZrus = LockService.getScriptLock();
    if (!lockZrus.tryLock(15000)) {
      return HtmlService.createHtmlOutput(htmlStart + "Server je práve vyťažený, skús to o chvíľu znova." + htmlEnd);
    }
    try {
      const data = sheetRez.getDataRange().getDisplayValues();

      for (let i = 1; i < data.length; i++) {
        if (data[i][7] === id) { 
          let email = data[i][5];
          let meno = data[i][3];
          let dStr = data[i][0];
          let cas = data[i][1].replace(/^'/, '');
          
          archiveReservation(data[i], role === 'holic' ? 'Zrušil holič' : 'Zrušil zákazník', dovodInput);
          sheetRez.deleteRow(i + 1);
          termiZruseny = true;
          invalidatePublicCache_(); // termín sa na webe má hneď uvoľniť
          
          try {
            if (role === 'holic') {
              const subject = "⚠️ Zrušenie rezervácie - Barbar Shop";
              const bodyHtml = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, true);
              MailApp.sendEmail({ name: MAIL_NAZOV, to: email, subject: subject, htmlBody: bodyHtml });
            } else {
              // Najprv zákazník (potvrdenie musí odísť), holičovi to pri nízkej kvóte príde vo večernom súhrne.
              const subjectKlient = "Potvrdenie zrušenia termínu";
              const bodyHtmlKlient = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, true, true);
              MailApp.sendEmail({ name: MAIL_NAZOV, to: email, subject: subjectKlient, htmlBody: bodyHtmlKlient });

              sendBarberMailOrDefer_(
                { subject: `❌ Zrušený termín: ${meno}`, htmlBody: generateCancellationEmailHtml(meno, dStr, cas, dovodInput, false) },
                { typ: 'Zrušené termíny', text: `${meno}, ${dStr} o ${cas}` + (dovodInput ? ` (dôvod: ${dovodInput})` : '') }
              );
            }
          } catch(err) {
            console.error('Zlyhalo odoslanie e-mailu pri zrušení (executeCancel):', err);
          }
          break;
        }
      }

      // Nepotvrdená rezervácia: len uvoľniť termín, bez mailov (holič o nej ešte nevie).
      if (!termiZruseny) {
        const pend = findPending_(id);
        if (pend.p) {
          releasePending_(pend.sheet, pend.p, 'Zrušená cez odkaz v maile pred potvrdením' + (dovodInput ? ': ' + dovodInput : ''));
          invalidatePublicCache_();
          termiZruseny = true;
          nepotvrdena = true;
        }
      }
    } finally {
      lockZrus.releaseLock();
    }

    const cardTitle = termiZruseny ? "Termín bol zrušený" : "Termín nenájdený";
    const cardText = nepotvrdena ? "Nepotvrdená rezervácia bola zrušená a termín je znova voľný. Môžeš zatvoriť túto stránku."
      : termiZruseny ? "O zrušení termínu bol odoslaný potvrdzujúci e-mail. Môžeš zatvoriť túto stránku." : "Tento termín už neexistuje alebo bol zrušený v minulosti.";
    const cardIcon = termiZruseny ? "<div class='success-icon'>✓</div>" : "<div class='error-icon'>❌</div>";
    
    const successHtml = `
      <!DOCTYPE html>
      <html lang="sk">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Výsledok | Barbar Shop</title>
        <style>
          ${fontFaceCss()}
          body { font-family: 'Roboto', 'Segoe UI', Helvetica, Arial, sans-serif; background-color: #1a1a1a; color: #e0e0e0; margin: 0; padding: 20px; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
          .card { background-color: #2b1d16; border: 1px solid #f0c419; border-radius: 12px; padding: 40px; max-width: 450px; width: 100%; box-shadow: 0 10px 30px rgba(0,0,0,0.5); text-align: center; }
          h2 { color: #f0c419; margin-top: 0; margin-bottom: 20px; }
          .success-icon { font-size: 60px; color: #4caf50; margin-bottom: 20px; }
          .error-icon { font-size: 60px; color: #d32f2f; margin-bottom: 20px; }
          p { margin-bottom: 0; line-height: 1.5; }
        </style>
      </head>
      <body>
        <div class="card">${cardIcon}<h2>${cardTitle}</h2><p>${cardText}</p></div>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(successHtml)
      .setTitle("Zrušenie rezervácie")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // --- AKCIA 3: ZOBRAZENIE E-MAILU V PREHLIADAČI ---
  if (action === 'viewEmail' && id) {
    const data = sheetRez.getDataRange().getDisplayValues();
    let found = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === id) { 
        found = { datum: data[i][0], cas: data[i][1].replace(/^'/, ''), meno: data[i][3], sluzba: data[i][4], email: data[i][5] }; 
        break; 
      }
    }
    
    if (found) {
      const template = HtmlService.createTemplateFromFile('EmailSablona');
      template.titulok = "Detail rezervácie"; 
      template.meno = found.meno; 
      template.uvodnyText = "Tu sú detaily tvojej potvrdenej rezervácie v Barbar Shope."; 
      template.sluzba = found.sluzba; 
      template.datum = found.datum; 
      template.cas = found.cas; 
      template.jeHolic = false; 
      template.emailKlienta = found.email; 
      template.cancelUrl = ScriptApp.getService().getUrl() + "?action=cancelPage&id=" + id + "&role=klient"; 
      template.viewUrl = "#"; 
      template.vyzvaText = "Ak chceš tento termín zrušiť, klikni na tlačidlo nižšie:";
      
      return template.evaluate()
        .setTitle("Rezervácia Barbar Shop")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      // Nepotvrdenú rezerváciu tu neukazujeme celú: odkaz na potvrdenie smie mať len ten, kto dostal mail.
      const pend = findPending_(id);
      if (pend.p && isPendingActive_(pend.p.r, Date.now())) {
        return HtmlService.createHtmlOutput(htmlStart + "Rezervácia čaká na potvrdenie. Potvrď ju tlačidlom POTVRDIŤ TERMÍN v maile." + htmlEnd);
      }
      return HtmlService.createHtmlOutput(htmlStart + "Ľutujeme, ale táto rezervácia už neexistuje!" + htmlEnd);
    }
  }

  // --- AKCIA 3b: STRÁNKA NA NAPÍSANIE RECENZIE ---
  if (action === 'reviewPage' && id) {
    const rez = findReservationForReview_(id);
    const chyba = !rez ? 'nenajdene' : (rez.chyba || (reviewExists_(id) ? 'existuje' : null));
    if (chyba) {
      return HtmlService.createHtmlOutput(htmlStart + escapeHtml(REVIEW_CHYBY[chyba]) + htmlEnd);
    }
    const template = HtmlService.createTemplateFromFile('RecenziaPage');
    template.id = id;
    template.meno = suggestDisplayName_(rez.meno);
    template.datum = rez.datum;
    template.sluzba = rez.sluzba;
    template.privacyUrl = SITE_URL + '/privacy.html';
    template.fontFaceCss = fontFaceCss();
    return template.evaluate().setTitle("Recenzia | Barbar Shop").addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // --- AKCIA 3e: POTVRDENIE TERMÍNU ZÁKAZNÍKOM (odkaz z mailu, 2 kroky) ---
  if (action === 'confirmPage') {
    return showConfirmPage_(id, e.parameter.kod);
  }
  if (action === 'doConfirm') {
    return doConfirm_(id, e.parameter.kod);
  }

  // --- AKCIA 3d: SCHVÁLENIE RECENZIE HOLIČOM (odkaz z mailu, 2 kroky) ---
  if (action === 'approveReview') {
    return showApproveReviewPage_(id, e.parameter.t);
  }
  if (action === 'doApproveReview') {
    return doApproveReview_(id, e.parameter.t, e.parameter.odpoved);
  }
  if (action === 'deleteReview') {
    return showDeleteReviewPage_(id, e.parameter.t);
  }
  if (action === 'doDeleteReview') {
    return doDeleteReview_(id, e.parameter.t);
  }

  // --- AKCIA 3c: SCHVÁLENÉ RECENZIE PRE WEB ---
  if (action === 'reviews') {
    return jsonOut(getPublicReviews_());
  }

  // --- AKCIA 4: JSON DÁTA PRE KALENDÁR (aj pre neznáme akcie) ---
  return publicDataResponse_();
}

// ===================== DÁTA PRE WEB (kalendár, recenzie, galéria) + CACHE =====================
// Odpoveď pre stránku sa skladá z 4 hárkov a galérie z Disku, čo je pomalé. Hotový JSON sa preto drží v cache
// (PUBLIC_CACHE_SEKUND) a maže sa vždy, keď sa niečo zmení: nová/zrušená rezervácia, schválená/zmazaná recenzia,
// ručná úprava tabuľky (handleEdit, onChange). Ďalšia poistka je časový trigger (installSpeedTriggers), ktorý
// cache každých 10 minút obnoví, takže návštevník väčšinou čaká na hotovú odpoveď. Rezerváciu to neohrozuje:
// doPost si obsadenosť vždy overuje priamo v tabuľke.
const PUBLIC_CACHE_KLUC = 'public_data_v1';
const PUBLIC_CACHE_VERZIA_KLUC = 'public_data_ver';
const PUBLIC_CACHE_SEKUND = 900;

// Zmaže hotovú odpoveď a zvýši "verziu", aby sa nezapísala odpoveď, ktorá sa práve zostavovala zo starých dát.
// Nikdy nehádže chybu: volá sa z miest, ktoré už dáta úspešne zmenili.
function invalidatePublicCache_() {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(PUBLIC_CACHE_VERZIA_KLUC, Utilities.getUuid(), 21600);
    cache.remove(PUBLIC_CACHE_KLUC);
  } catch (err) {
    console.error('Zmazanie cache zlyhalo (zmena sa na webe prejaví najneskôr po vypršaní cache):', err);
  }
}

// Vráti { json, zdroj }. preskocitCache = zostaviť odznova (trigger); vynutitGaleriu = znovu prečítať aj Disk.
function getPublicData_(preskocitCache, vynutitGaleriu) {
  const cache = CacheService.getScriptCache();
  if (!preskocitCache) {
    const hit = cache.get(PUBLIC_CACHE_KLUC);
    if (hit) return { json: hit, zdroj: 'cache' };
  }
  const verzia = cache.get(PUBLIC_CACHE_VERZIA_KLUC);
  updateWebAppUrl();
  const built = buildPublicData_(vynutitGaleriu);
  const json = JSON.stringify(built.data);
  // Zapíše sa len ak sa medzitým nič nezmenilo a ak sa všetko načítalo (výpadok galérie sa nemá držať v cache).
  if (built.uplne && cache.get(PUBLIC_CACHE_VERZIA_KLUC) === verzia) {
    try { cache.put(PUBLIC_CACHE_KLUC, json, PUBLIC_CACHE_SEKUND); } catch (err) { /* nad 100 kB sa cache preskočí */ }
  }
  return { json: json, zdroj: 'tabuľka' };
}

function publicDataResponse_() {
  const t0 = Date.now();
  const r = getPublicData_(false, false);
  console.log('Dáta pre web: ' + r.zdroj + ', ' + (Date.now() - t0) + ' ms, ' + r.json.length + ' znakov');
  return ContentService.createTextOutput(r.json).setMimeType(ContentService.MimeType.JSON);
}

// Časový trigger (každých 10 min): znovu prečíta tabuľky aj Disk, aby cache bola stále čerstvá.
function refreshPublicCache() {
  const t0 = Date.now();
  const r = getPublicData_(true, true);
  console.log('Cache obnovená za ' + (Date.now() - t0) + ' ms (' + r.json.length + ' znakov)');
}

// Trigger "pri zmene tabuľky" (vloženie/zmazanie riadku rukou, čo handleEdit nezachytí): obnoví len tabuľkovú časť.
function onSheetChange() {
  getPublicData_(true, false);
}

// Spusti RAZ ručne z editora (po nasadení novej verzie; opakované spustenie nič nezdvojí): založí triggery
// pre cache aj pre nepotvrdené rezervácie a hneď naplní cache.
function installSpeedTriggers() {
  const moje = ['refreshPublicCache', 'onSheetChange', 'processPendingReservations', 'sendBarberDigest'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (moje.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshPublicCache').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('onSheetChange').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onChange().create();
  ScriptApp.newTrigger('processPendingReservations').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('sendBarberDigest').timeBased().everyDays(1).atHour(SUHRN_HODINA).inTimezone('Europe/Bratislava').create();
  refreshPublicCache();
  console.log('Triggery založené: refreshPublicCache (každých 10 minút), onSheetChange (pri zmene tabuľky), processPendingReservations (každú minútu), sendBarberDigest (denne okolo ' + SUHRN_HODINA + ':00).');
}

// Zostaví dáta pre web. uplne = false, ak sa niečo (recenzie, galéria) nepodarilo načítať; takú odpoveď necachujeme.
function buildPublicData_(vynutitGaleriu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRez = ss.getSheetByName(SHEET_NAME);
  const sheetCustom = ss.getSheetByName(SHEET_CUSTOM);
  const sheetDefault = ss.getSheetByName(SHEET_DEFAULT); 
  
  let bookings = [];
  if (sheetRez) {
    const resData = sheetRez.getDataRange().getDisplayValues();
    for (let i = 1; i < resData.length; i++) {
      let parts = resData[i][0].split('.');
      if (parts.length === 3) {
        bookings.push({ 
          datum: `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`, 
          cas: resData[i][1].replace(/^'/, '')
        });
      }
    }
  }

  // Termíny, ktoré čakajú na potvrdenie mailom, sú pre ostatných obsadené (na webe označené ako držané).
  let uplne = true;
  try {
    activePendingBookings_(Date.now()).forEach(function(b) {
      bookings.push({ datum: b.datum, cas: b.cas, pending: true });
    });
  } catch (err) {
    uplne = false;
    console.error('Načítanie nepotvrdených rezervácií zlyhalo (kalendár ide ďalej):', err);
  }

  let customRanges = [];
  if (sheetCustom) {
    const custData = sheetCustom.getDataRange().getDisplayValues();
    for (let i = 1; i < custData.length; i++) {
      let stav = custData[i][3] ? custData[i][3].trim().toUpperCase() : "";
      if (!stav) continue; 
      
      let odFormatted = custData[i][1]; 
      let doFormatted = custData[i][2];
      
      if (stav !== 'ZAVRETÉ') { 
        odFormatted = formatToGridStart(custData[i][1]); 
        doFormatted = formatToGridEnd(custData[i][2]); 
      }
      
      let expanded = expandDates(custData[i][0]);
      expanded.forEach(d => { 
        customRanges.push({ datum: d, od: odFormatted, do: doFormatted, stav: stav }); 
      });
    }
  }

  let defaultHours = [];
  if (sheetDefault) {
    const defData = sheetDefault.getDataRange().getDisplayValues();
    const dniMapa = {'Pondelok':1,'Utorok':2,'Streda':3,'Štvrtok':4,'Piatok':5,'Sobota':6,'Nedeľa':7};
    
    for (let i = 1; i < defData.length; i++) {
      let denCislo = dniMapa[defData[i][0].trim()];
      if (!denCislo) continue;
      
      let rozpis = { den: denCislo, bloky: [] };
      for (let j = 1; j < defData[i].length; j += 2) {
        if (defData[i][j] && defData[i][j+1]) {
          rozpis.bloky.push({ 
            od: formatToGridStart(defData[i][j]), 
            do: formatToGridEnd(defData[i][j+1]) 
          });
        }
      }
      if (rozpis.bloky.length > 0) {
        defaultHours.push(rozpis);
      }
    }
  }
  
  // Schválené recenzie idú v tej istej odpovedi ako kalendár: web tak vie hneď, či nejaké sú, a nemusí
  // hádať (kostra, ktorá zmizne) ani používať uloženú kópiu (zmazaná recenzia by sa ešte ukázala).
  let reviews = { count: 0, average: 0, reviews: [] };
  try {
    reviews = getPublicReviews_();
  } catch (err) {
    uplne = false;
    console.error('Načítanie recenzií zlyhalo (kalendár ide ďalej):', err);
  }

  let gallery = { items: [] };
  try {
    gallery = getGalleryItems_(vynutitGaleriu);
  } catch (err) {
    uplne = false;
    console.error('Načítanie galérie zlyhalo (kalendár ide ďalej):', err);
  }

  return { data: { bookings, custom: customRanges, defaultHours, reviews, gallery }, uplne: uplne };
}

// --- LOGIKA PRE ZRUŠENIE A UPOZORNENIA PRIAMO V TABUĽKE ---
// Obal okolo skutočnej obsluhy: cache pre web sa zmaže AŽ po dokončení (dialógy v tabuľke môžu trvať aj desiatky
// sekúnd a medzitým by si ju návštevník stihol zostaviť zo starých dát).
function handleEdit(e) {
  try {
    handleEditInner_(e);
  } finally {
    if (e && e.range) invalidatePublicCache_();
  }
}

function handleEditInner_(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // 1. Zrušenie cez CHECKBOX priamo v tabuľke
  if (sheetName === SHEET_NAME && e.range.getColumn() === 7 && e.range.getValue() === true) {
    let ui = SpreadsheetApp.getUi();
    let prompt = ui.prompt('Zrušenie rezervácie', 'Zadaj dôvod zrušenia, ktorý odíde klientovi na e-mail:', ui.ButtonSet.OK_CANCEL);
    
    if (prompt.getSelectedButton() === ui.Button.OK) {
       let dovod = prompt.getResponseText() || "Neočakávané prevádzkové dôvody.";
       let row = e.range.getRow();
       let data = sheet.getRange(row, 1, 1, 8).getDisplayValues()[0];
       
       try { 
         const bodyHtml = generateCancellationEmailHtml(data[3], data[0], data[1].replace(/^'/, ''), dovod, true);
         MailApp.sendEmail({ name: MAIL_NAZOV, 
           to: data[5], 
           subject: "⚠️ Zrušenie rezervácie - Barbar Shop", 
           htmlBody: bodyHtml
         });
       } catch(err) {
         console.error('Zlyhalo odoslanie e-mailu pri zrušení (checkbox):', err);
       }

       archiveReservation(data, 'Zrušil holič', dovod);
       sheet.deleteRow(row);
    } else {
       e.range.setValue(false); 
    }
    return;
  }

  // 1a. Ručné potvrdenie čakajúcej rezervácie (checkbox v hárku Nepotvrdené), napr. kamarátovi, ktorý nemá prístup k mailu
  if (sheetName === SHEET_CAKAJUCE && e.range.getColumn() === CAKAJUCE_STLPEC_POTVRDIT && e.range.getRow() > 1 && e.range.getValue() === true) {
    const ui = SpreadsheetApp.getUi();
    const r = sheet.getRange(e.range.getRow(), 1, 1, CAKAJUCE_HLAVICKA.length).getValues()[0];
    const odpoved = ui.alert('Potvrdiť rezerváciu?',
      'Rezervácia "' + plain_(r[3]) + '" na ' + plain_(r[0]) + ' o ' + plain_(r[1], 'HH:mm') + ' sa potvrdí bez kliknutia v maile a presunie do hárku Rezervácie. Pokračovať?',
      ui.ButtonSet.YES_NO);
    if (odpoved !== ui.Button.YES) {
      e.range.setValue(false);
      return;
    }
    const v = confirmBooking_(String(r[6]), null, true);
    if (v.title === 'Termín je potvrdený') {
      SpreadsheetApp.getActiveSpreadsheet().toast('Rezervácia je potvrdená a presunutá do hárku Rezervácie.', 'Hotovo', 6);
    } else {
      // checkbox vrátime len ak je na tom mieste stále ten istý riadok (pri uvoľnení sa riadky posunú)
      try { if (String(sheet.getRange(e.range.getRow(), 7).getValue()) === String(r[6])) e.range.setValue(false); } catch (err) {}
      ui.alert(v.title, String(v.inner).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(), ui.ButtonSet.OK);
    }
    return;
  }

  // 1b. Zmazanie recenzie cez CHECKBOX v hárku "Recenzie" (posledný stĺpec)
  if (sheetName === SHEET_RECENZIE && e.range.getColumn() === RECENZIE_STLPEC_SCHVALENE && e.range.getRow() > 1) {
    // Ručné (od)schválenie v tabuľke: fotky recenzie sa podľa toho sprístupnia / znova zneprístupnia.
    const rs = sheet.getRange(e.range.getRow(), 1, 1, RECENZIE_HLAVICKA.length).getValues()[0];
    syncReviewPhotoSharing_(rs, e.range.getValue() === true);
    return;
  }

  if (sheetName === SHEET_RECENZIE && e.range.getColumn() === RECENZIE_STLPEC_ZMAZAT && e.range.getRow() > 1 && e.range.getValue() === true) {
    const ui = SpreadsheetApp.getUi();
    const r = sheet.getRange(e.range.getRow(), 1, 1, RECENZIE_HLAVICKA.length).getValues()[0];
    const odpoved = ui.alert(
      'Zmazať recenziu?',
      'Recenzia od "' + r[4] + '" (' + r[2] + '/5) sa nenávratne odstráni z tabuľky aj z webu. Pokračovať?',
      ui.ButtonSet.YES_NO
    );
    if (odpoved === ui.Button.YES) {
      deleteReviewPhotos_(r);
      sheet.deleteRow(e.range.getRow());
    } else {
      e.range.setValue(false);
    }
    return;
  }

  // 2. Kontrola konfliktov (Aktivuje sa zmenou v Otváracích hodinách)
  if (sheetName === SHEET_CUSTOM && e.range.getColumn() <= 4) {
    let row = e.range.getRow();
    if (row === 1) return;
    
    let data = sheet.getRange(row, 1, 1, 4).getDisplayValues()[0];
    if (!data[0] || !data[1] || !data[2] || !data[3]) return;
    
    let dateInput = data[0];
    let zadaneDatumy = expandDates(dateInput); 
    if (zadaneDatumy.length === 0) return;
    
    let defaultSheet = e.source.getSheetByName(SHEET_DEFAULT);
    let rules = buildScheduleRules(sheet, defaultSheet);
    let customRules = rules.customRules;
    let defaultHoursMap = rules.defaultHoursMap;

    let rezSheet = e.source.getSheetByName(SHEET_NAME);
    if (!rezSheet) return;
    let rezData = rezSheet.getDataRange().getDisplayValues();
    let konfliktneDatumyObj = [];
    
    const dnesPreKontrolu = new Date(); 
    dnesPreKontrolu.setHours(0,0,0,0);

    for (let i = 1; i < rezData.length; i++) {
      let p = String(rezData[i][0]).replace(/\s/g, '').split('.');
      if (p.length === 3) {
        let rezDatumYMD = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
        let rezObjekt = new Date(p[2], p[1]-1, p[0]);
        
        if (rezObjekt >= dnesPreKontrolu && zadaneDatumy.includes(rezDatumYMD)) {
          let bStart = parseTime(rezData[i][1].replace(/^'/, ''));
          
          // ZJEDNOTENÁ LOGIKA: Overíme presnú mriežku pre daný deň
          let platneSloty = getValidSlotsForDay(rezDatumYMD, customRules, defaultHoursMap);
          
          if (!platneSloty.includes(bStart)) {
             konfliktneDatumyObj.push({ 
               row: i + 1, 
               datum: rezData[i][0], 
               cas: rezData[i][1], 
               meno: rezData[i][3], 
               email: rezData[i][5] 
             });
          }
        }
      }
    }

    if (konfliktneDatumyObj.length > 0) {
      let template = HtmlService.createTemplateFromFile('DialogZrusenie');
      template.konflikty = jsonForScript(konfliktneDatumyObj);
      // Jednorazový kľúč: processCancellations je verejne volateľná cez google.script.run (aj z cudzej
      // stránky tejto aplikácie), preto ju povolíme len z dialógu, ktorý sme práve otvorili my.
      const nonce = Utilities.getUuid();
      CacheService.getScriptCache().put('cancelNonce_' + nonce, '1', 21600);
      template.nonce = nonce;
      let html = template.evaluate().setWidth(550).setHeight(450);
      SpreadsheetApp.getUi().showModalDialog(html, 'Konflikt pri rezerváciách a časoch!');
    }
  }
}

// Hromadné rušenie z Popup Okna
function processCancellations(cancellations, nonce) {
  const cache = CacheService.getScriptCache();
  if (!nonce || !cache.get('cancelNonce_' + nonce) || !Array.isArray(cancellations)) {
    throw new Error('Neplatná relácia, otvor dialóg s konfliktmi znova.');
  }
  cache.remove('cancelNonce_' + nonce);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  cancellations.sort((a,b) => b.row - a.row);

  cancellations.forEach(c => {
    try {
      // Údaje (meno, e-mail, termín) berieme priamo z tabuľky, nie z toho, čo prišlo z dialógu.
      const r = sheet.getRange(c.row, 1, 1, 8).getDisplayValues()[0];
      const bodyHtml = generateCancellationEmailHtml(r[3], r[0], String(r[1]).replace(/^'/, ''), c.dovod, true);
      MailApp.sendEmail({ name: MAIL_NAZOV,
        to: r[5],
        subject: "⚠️ Zrušenie rezervácie - Barbar Shop",
        htmlBody: bodyHtml
      });
      archiveReservation(r, 'Zrušil holič', c.dovod);
      sheet.deleteRow(c.row);
    } catch(err) {
      console.error("Nepodarilo sa zrušiť riadok:", c.row);
    }
  });
  invalidatePublicCache_();
}

// ===================== RECENZIE =====================
// Tok: deň po strihu príde zákazníkovi mail s odkazom (?action=reviewPage&id=<ID rezervácie>) ->
// stránka RecenziaPage.html -> submitReview() zapíše riadok do hárku "Recenzie" -> holič ho
// odklikne (stĺpec "Schválené") -> riadok sa objaví na webe (?action=reviews).
// Jedna rezervácia = najviac jedna recenzia; rezervácia sa hľadá v "Rezervácie" aj v "Archív".
const SHEET_RECENZIE = 'Recenzie';
const RECENZIE_HLAVICKA = ['Identifikátor rezervácie', 'Dátum strihania', 'Hodnotenie (1-5)', 'Recenzia', 'Zobrazené meno', 'Odoslané (súhlas so zverejnením)', 'Schválené', 'Odpoveď holiča', 'Zmazať recenziu?', 'Fotky (ID súborov)', 'Priečinok s fotkami'];
const RECENZIE_STLPEC_SCHVALENE = 7;
const RECENZIE_STLPEC_ZMAZAT = 9;
const RECENZIE_STLPEC_FOTKY = 10;
const RECENZIE_STLPEC_PRIECINOK = 11;
const REVIEW_PLATNOST_DNI = 30;
const REVIEW_MAX_ZOBRAZENYCH = 12;
const REVIEW_STLPEC_MAILU = 10; // J v hárku Rezervácie: označí, že mail s recenziou už odišiel
const SITE_URL = 'https://barbarshop-mu.vercel.app';
const REVIEW_CHYBY = {
  nenajdene: 'Rezervácia nebola nájdená.',
  zrusena: 'Tento termín bol zrušený, takže sa k nemu nedá napísať recenzia.',
  este_nie: 'Na recenziu si ešte počkaj, kým strih prebehne.',
  expirovane: 'Platnosť odkazu na recenziu už vypršala.',
  existuje: 'K tejto návšteve už recenzia existuje, ďakujeme!'
};

function getReviewSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RECENZIE);
  if (!sheet) {
    const povodny = ss.getActiveSheet();
    sheet = ss.insertSheet(SHEET_RECENZIE);
    sheet.getRange(1, 1, 1, RECENZIE_HLAVICKA.length).setValues([RECENZIE_HLAVICKA]).setFontWeight('bold');
    sheet.getRange(1, 1, sheet.getMaxRows(), 2).setNumberFormat('@');
    sheet.getRange(1, 6, sheet.getMaxRows(), 1).setNumberFormat('@');
    sheet.setFrozenRows(1);
    try { ss.setActiveSheet(povodny); } catch (err) {}
  }
  upgradeReviewSheet_(sheet);
  return sheet;
}

// Starší hárok "Recenzie" sa sám doplní o nové stĺpce (hlavička, a pri "Zmazať recenziu?" aj checkboxy).
function upgradeReviewSheet_(sheet) {
  if (!sheet.getRange(1, RECENZIE_STLPEC_ZMAZAT).getValue()) {
    sheet.getRange(1, RECENZIE_STLPEC_ZMAZAT).setValue(RECENZIE_HLAVICKA[RECENZIE_STLPEC_ZMAZAT - 1]);
    sheet.getRange(1, RECENZIE_STLPEC_ZMAZAT - 1).copyFormatToRange(sheet, RECENZIE_STLPEC_ZMAZAT, RECENZIE_STLPEC_ZMAZAT, 1, 1);
    if (sheet.getLastRow() > 1) sheet.getRange(2, RECENZIE_STLPEC_ZMAZAT, sheet.getLastRow() - 1, 1).insertCheckboxes();
  }
  for (let c = RECENZIE_STLPEC_ZMAZAT + 1; c <= RECENZIE_HLAVICKA.length; c++) {
    if (!sheet.getRange(1, c).getValue()) {
      sheet.getRange(1, c).setValue(RECENZIE_HLAVICKA[c - 1]);
      sheet.getRange(1, RECENZIE_STLPEC_ZMAZAT - 1).copyFormatToRange(sheet, c, c, 1, 1);
    }
  }
}

function slovakDateToYMD_(str) {
  const p = String(str).replace(/\s/g, '').split('.');
  if (p.length !== 3) return null;
  return p[2] + '-' + p[1].padStart(2, '0') + '-' + p[0].padStart(2, '0');
}

// Aktuálny bratislavský čas prevedený na Date z lokálnych čísel - rovnaký trik ako vo validateBookingData,
// takže rozdiely oproti new Date(y, m, d, h, min) vychádzajú správne bez ohľadu na časovú zónu projektu.
function bratislavaNow_() {
  return new Date(Utilities.formatDate(new Date(), 'Europe/Bratislava', 'yyyy/MM/dd HH:mm:ss'));
}

// Vráti {datum, cas, meno, sluzba}, alebo {chyba: kód z REVIEW_CHYBY}, alebo null (nenájdené).
function findReservationForReview_(id) {
  if (!id || !/^[0-9a-f-]{36}$/i.test(id)) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rez = null;

  const sheetRez = ss.getSheetByName(SHEET_NAME);
  if (sheetRez) {
    const data = sheetRez.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === id) {
        rez = { datum: data[i][0], cas: String(data[i][1]).replace(/^'/, ''), meno: data[i][3], sluzba: data[i][4] };
        break;
      }
    }
  }
  if (!rez) {
    const sheetArch = ss.getSheetByName(SHEET_ARCHIV);
    if (sheetArch && sheetArch.getLastRow() > 1) {
      const data = sheetArch.getRange(2, 1, sheetArch.getLastRow() - 1, ARCHIV_HLAVICKA.length).getDisplayValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][6] === id) {
          if (data[i][7] !== 'Prebehla') return { chyba: 'zrusena' };
          rez = { datum: data[i][0], cas: String(data[i][1]).replace(/^'/, ''), meno: data[i][3], sluzba: data[i][4] };
          break;
        }
      }
    }
  }
  if (!rez) return null;

  const ymd = slovakDateToYMD_(rez.datum);
  if (!ymd) return null;
  const d = ymd.split('-').map(Number);
  const t = String(rez.cas).split(':').map(Number);
  const start = new Date(d[0], d[1] - 1, d[2], t[0] || 0, t[1] || 0, 0);
  const odStartu = (bratislavaNow_().getTime() - start.getTime()) / 60000;
  if (odStartu < 90) return { chyba: 'este_nie' };
  if (odStartu > REVIEW_PLATNOST_DNI * 1440) return { chyba: 'expirovane' };
  return rez;
}

function reviewExists_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECENZIE);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  return ids.some(function(r) { return r[0] === id; });
}

// "Jakub Kramár" -> "Jakub K." (zákazník si to na stránke môže upraviť)
function suggestDisplayName_(full) {
  const parts = String(full || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '';
  if (parts.length === 1) return parts[0];
  return parts[0] + ' ' + parts[parts.length - 1].charAt(0).toUpperCase() + '.';
}

// ---------- FOTKY K RECENZIÁM ----------
// Bezpečnostný model (súbory od cudzích ľudí sú riziko):
//  - prijímame LEN JPEG a overujeme ho na serveri (podpis súboru, ukončenie, rozmery z hlavičky); stránka fotku
//    pred odoslaním vždy prekóduje do JPEG v canvase, čím zmizne aj poloha (EXIF) a čokoľvek skryté v origináli,
//  - meno súboru, typ ani priečinok nikdy nebereme od klienta: súbory sa volajú foto-1.jpg ... a ukladajú sa
//    do priečinka tejto recenzie v súkromnom priečinku "reviews",
//  - najviac REVIEW_MAX_FOTIEK fotiek, limity na veľkosť a rozmery, jedna recenzia na rezerváciu,
//  - súbory sú SÚKROMNÉ, kým holič recenziu neschváli; po schválení sa zdieľajú len tie, pri zmazaní recenzie
//    sa priečinok s fotkami presunie do koša.
const REVIEWS_PRIECINOK_NAZOV = 'reviews';
const REVIEW_MAX_FOTIEK = 5;
const REVIEW_FOTKA_MAX_BAJTOV = 1500000;
const REVIEW_FOTKY_SPOLU_MAX_BAJTOV = 6000000;
const REVIEW_FOTKA_MAX_ROZMER = 4096;

function splitIds_(str) {
  return String(str || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return /^[\w-]{10,}$/.test(s); });
}

// Priečinok daného názvu, ktorého vlastníkom je účet, pod ktorým skript beží (nie niekto, kto nám podobný zdieľal).
function getOwnedFolder_(name, vytvorit) {
  const me = Session.getEffectiveUser().getEmail();
  const it = DriveApp.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    try {
      if (!f.isTrashed() && f.getOwner() && f.getOwner().getEmail() === me) return f;
    } catch (err) {}
  }
  return vytvorit ? DriveApp.getRootFolder().createFolder(name) : null;
}

// Overí jednu fotku (base64 text) ako JPEG. Vráti {ok:true, bytes} alebo {ok:false, error}.
function validateJpegBase64_(b64) {
  if (typeof b64 !== 'string' || b64.length < 100 || !/^[A-Za-z0-9+\/]+={0,2}$/.test(b64)) return { ok: false, error: 'má neplatný formát dát' };
  if (b64.length > Math.ceil(REVIEW_FOTKA_MAX_BAJTOV * 4 / 3) + 8) return { ok: false, error: 'je príliš veľká' };
  let raw;
  try { raw = Utilities.base64Decode(b64); } catch (err) { return { ok: false, error: 'má neplatný formát dát' }; }
  const n = raw.length;
  if (n < 1000 || n > REVIEW_FOTKA_MAX_BAJTOV) return { ok: false, error: 'má neprijateľnú veľkosť' };
  const b = function(i) { return raw[i] & 255; };

  if (b(0) !== 0xFF || b(1) !== 0xD8 || b(2) !== 0xFF) return { ok: false, error: 'nie je obrázok JPEG' };
  if (b(n - 2) !== 0xFF || b(n - 1) !== 0xD9) return { ok: false, error: 'je poškodená' };

  // Prejdeme segmenty hlavičky až po SOF (tam sú rozmery). SOS/EOI pred SOF = poškodený súbor.
  let i = 2, w = 0, h = 0;
  while (i + 8 < n) {
    if (b(i) !== 0xFF) return { ok: false, error: 'je poškodená' };
    const m = b(i + 1);
    if (m === 0xFF) { i++; continue; }
    if (m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
    if (m === 0xD8 || m === 0xD9 || m === 0xDA) return { ok: false, error: 'je poškodená' };
    const len = (b(i + 2) << 8) | b(i + 3);
    if (len < 2) return { ok: false, error: 'je poškodená' };
    const jeSof = m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC;
    if (jeSof) {
      h = (b(i + 5) << 8) | b(i + 6);
      w = (b(i + 7) << 8) | b(i + 8);
      break;
    }
    i += 2 + len;
  }
  if (!w || !h) return { ok: false, error: 'je poškodená' };
  if (w < 16 || h < 16 || w > REVIEW_FOTKA_MAX_ROZMER || h > REVIEW_FOTKA_MAX_ROZMER || w * h > 12000000) return { ok: false, error: 'má neprijateľné rozmery' };
  return { ok: true, bytes: raw };
}

// Overí všetky fotky naraz (nič sa nezapíše, kým nie sú v poriadku všetky). Vráti {ok, files|error}.
function validateReviewPhotos_(photos, suhlas) {
  if (photos === undefined || photos === null) return { ok: true, files: [] };
  if (!Array.isArray(photos)) return { ok: false, error: 'Neplatné fotky.' };
  if (photos.length === 0) return { ok: true, files: [] };
  if (photos.length > REVIEW_MAX_FOTIEK) return { ok: false, error: 'K recenzii môžeš pridať najviac ' + REVIEW_MAX_FOTIEK + ' fotiek.' };
  if (suhlas !== true) return { ok: false, error: 'K fotkám je potrebné potvrdiť súhlas so zverejnením.' };
  const files = [];
  let spolu = 0;
  for (let i = 0; i < photos.length; i++) {
    const v = validateJpegBase64_(photos[i]);
    if (!v.ok) return { ok: false, error: 'Fotka č. ' + (i + 1) + ' ' + v.error + '.' };
    spolu += v.bytes.length;
    if (spolu > REVIEW_FOTKY_SPOLU_MAX_BAJTOV) return { ok: false, error: 'Fotky sú dokopy príliš veľké.' };
    files.push(v.bytes);
  }
  return { ok: true, files: files };
}

// Uloží fotky do nového súkromného priečinka tejto recenzie: reviews/<dátum> - <meno> - <začiatok ID>/foto-N.jpg
function saveReviewPhotos_(id, rez, zobrazeneMeno, files) {
  const root = getOwnedFolder_(REVIEWS_PRIECINOK_NAZOV, true);
  const datum = slovakDateToYMD_(rez.datum) || 'bez-datumu';
  const meno = String(zobrazeneMeno).replace(/[^\p{L}\p{N} ._-]/gu, '_').slice(0, 30);
  const folder = root.createFolder(datum + ' - ' + meno + ' - ' + id.slice(0, 8));
  const fileIds = [];
  try {
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    files.forEach(function(bytes, i) {
      const file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', 'foto-' + (i + 1) + '.jpg'));
      fileIds.push(file.getId());
    });
  } catch (err) {
    try { folder.setTrashed(true); } catch (e2) {}
    throw err;
  }
  return { folderId: folder.getId(), fileIds: fileIds };
}

// Po schválení sa fotky sprístupnia cez odkaz (aby ich vedel načítať web), po zrušení schválenia sú zasa súkromné.
function syncReviewPhotoSharing_(row, schvalene) {
  splitIds_(row[RECENZIE_STLPEC_FOTKY - 1]).forEach(function(fid) {
    try {
      const f = DriveApp.getFileById(fid);
      if (schvalene) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      else f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    } catch (err) {
      console.error('Nastavenie zdieľania fotky zlyhalo:', fid, err);
    }
  });
}

// Pri zmazaní recenzie ide do koša celý priečinok s jej fotkami (a pre istotu aj jednotlivé súbory).
function deleteReviewPhotos_(row) {
  const m = String(row[RECENZIE_STLPEC_PRIECINOK - 1] || '').match(/folders\/([\w-]{10,})/);
  if (m) {
    try { DriveApp.getFolderById(m[1]).setTrashed(true); } catch (err) { console.error('Zmazanie priečinka s fotkami zlyhalo:', err); }
  }
  splitIds_(row[RECENZIE_STLPEC_FOTKY - 1]).forEach(function(fid) {
    try { DriveApp.getFileById(fid).setTrashed(true); } catch (err) {}
  });
}

// Náhľad fotiek na stránke pre holiča: fotky sú súkromné, preto idú priamo v stránke (base64), nie cez odkaz.
function reviewPhotoPreviewHtml_(row) {
  let out = '';
  splitIds_(row[RECENZIE_STLPEC_FOTKY - 1]).forEach(function(fid) {
    try {
      const bytes = DriveApp.getFileById(fid).getBlob().getBytes();
      if (bytes.length <= REVIEW_FOTKA_MAX_BAJTOV + 100000) {
        out += '<img src="data:image/jpeg;base64,' + Utilities.base64Encode(bytes) + '" alt="Fotka k recenzii">';
      }
    } catch (err) {
      console.error('Náhľad fotky zlyhal:', fid, err);
    }
  });
  return out ? '<div class="photos">' + out + '</div>' : '';
}

// Volá sa zo stránky RecenziaPage.html cez google.script.run. Všetko sa overuje tu na serveri.
function submitReview(payload) {
  try {
    if (!payload || typeof payload !== 'object') return { ok: false, error: 'Neplatné dáta.' };
    const id = String(payload.id || '');
    const stars = Number(payload.stars);
    const text = String(payload.text || '').trim();
    const name = String(payload.name || '').trim();

    if (payload.consent !== true) return { ok: false, error: 'Bez súhlasu so zverejnením recenziu nevieme uložiť.' };
    if (!(stars >= 1 && stars <= 5 && Math.floor(stars) === stars)) return { ok: false, error: 'Vyber počet hviezdičiek.' };
    if (text.length < 5 || text.length > 1000) return { ok: false, error: 'Recenzia musí mať 5 až 1000 znakov.' };
    if (name.length < 1 || name.length > 40) return { ok: false, error: 'Zobrazené meno musí mať 1 až 40 znakov.' };

    // Fotky sa overia pred čímkoľvek iným: ak je čo i len jedna zlá, neuloží sa nič.
    const overene = validateReviewPhotos_(payload.photos, payload.photosConsent);
    if (!overene.ok) return { ok: false, error: overene.error };

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return { ok: false, error: 'Server je vyťažený, skús to prosím o chvíľu znova.' };

    let rezInfo = null;
    try {
      const rez = findReservationForReview_(id);
      if (!rez) return { ok: false, error: REVIEW_CHYBY.nenajdene };
      if (rez.chyba) return { ok: false, error: REVIEW_CHYBY[rez.chyba] };
      if (reviewExists_(id)) return { ok: false, error: REVIEW_CHYBY.existuje };
      rezInfo = rez;

      const sheet = getReviewSheet_();
      const novyRiadok = [
        id,
        rez.datum,
        stars,
        safeCell(text),
        safeCell(name),
        Utilities.formatDate(new Date(), 'Europe/Bratislava', 'd.M.yyyy HH:mm:ss'),
        false,
        '',
        false,
        '',
        ''
      ];
      // Najnovšia recenzia hore (hneď pod hlavičkou), nech si na nové a neodpovedané nezabudneš.
      let cielovyRiadok;
      if (sheet.getLastRow() < 2) {
        sheet.appendRow(novyRiadok);
        cielovyRiadok = sheet.getLastRow();
      } else {
        sheet.insertRowAfter(1);
        sheet.getRange(3, 1, 1, RECENZIE_HLAVICKA.length).copyFormatToRange(sheet, 1, RECENZIE_HLAVICKA.length, 2, 2);
        sheet.getRange(2, 1, 1, RECENZIE_HLAVICKA.length).setValues([novyRiadok]);
        cielovyRiadok = 2;
      }
      sheet.getRange(cielovyRiadok, RECENZIE_STLPEC_SCHVALENE).insertCheckboxes();
      sheet.getRange(cielovyRiadok, RECENZIE_STLPEC_ZMAZAT).insertCheckboxes();
    } finally {
      lock.releaseLock();
    }

    // Fotky sa nahrávajú až po uložení recenzie a mimo zámku (Disk je pomalý). Ak by zlyhali, text recenzie ostáva.
    let pocetFotiek = 0;
    let fotkyZlyhali = false;
    if (overene.files.length > 0) {
      let ulozene = null;
      try {
        ulozene = saveReviewPhotos_(id, rezInfo, name, overene.files);
        const lock2 = LockService.getScriptLock();
        if (!lock2.tryLock(15000)) throw new Error('Zámok sa nepodarilo získať.');
        try {
          const r = findReview_(id);
          if (!r) throw new Error('Riadok recenzie sa nenašiel.');
          const sheet2 = getReviewSheet_();
          sheet2.getRange(r.row, RECENZIE_STLPEC_FOTKY).setValue(ulozene.fileIds.join(','));
          sheet2.getRange(r.row, RECENZIE_STLPEC_PRIECINOK).setValue('https://drive.google.com/drive/folders/' + ulozene.folderId);
        } finally {
          lock2.releaseLock();
        }
        pocetFotiek = ulozene.fileIds.length;
      } catch (err) {
        console.error('Uloženie fotiek k recenzii zlyhalo:', err);
        fotkyZlyhali = true;
        if (ulozene) { try { DriveApp.getFolderById(ulozene.folderId).setTrashed(true); } catch (e2) {} }
      }
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const baseUrl = ScriptApp.getService().getUrl();
      const approveUrl = baseUrl + '?action=approveReview&id=' + id + '&t=' + reviewToken_(id, 'approve');
      const deleteUrl = baseUrl + '?action=deleteReview&id=' + id + '&t=' + reviewToken_(id, 'delete');
      const sheetUrl = ss.getUrl() + '#gid=' + getReviewSheet_().getSheetId();
      sendBarberMailOrDefer_(
        { subject: 'Nová recenzia čaká na schválenie (' + stars + '/5)', htmlBody: generateNewReviewEmailHtml_(name, stars, text, approveUrl, deleteUrl, sheetUrl, pocetFotiek) },
        { typ: 'Nové recenzie', text: name + ' (' + stars + '/5): ' + (text.length > 140 ? text.slice(0, 140) + '...' : text) + (pocetFotiek ? ' [' + pocetFotiek + ' foto]' : ''),
          odkazy: [{ t: 'Schváliť', u: approveUrl }, { t: 'Zmazať', u: deleteUrl }] }
      );
    } catch (err) {
      console.error('Upozornenie na novú recenziu sa neodoslalo:', err);
    }
    return { ok: true, fotkyZlyhali: fotkyZlyhali };
  } catch (err) {
    console.error('submitReview zlyhalo:', err);
    return { ok: false, error: 'Niečo sa pokazilo, skús to prosím znova.' };
  }
}

// Podpis odkazu na schválenie. Samotné ID recenzie nestačí (pozná ho aj zákazník zo svojho odkazu na
// recenziu), preto sa odkaz podpisuje tajným kľúčom, ktorý zná len skript.
// Každá akcia (approve / delete) má vlastný podpis, takže odkaz na schválenie nemožno použiť na zmazanie.
function reviewToken_(id, purpose) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('REVIEW_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('REVIEW_SECRET', secret);
  }
  const sig = Utilities.computeHmacSha256Signature(id + '|' + purpose, secret);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '').slice(0, 32);
}

function reviewApprovalToken_(id) {
  return reviewToken_(id, 'approve');
}

// Vráti {row, id, stars, text, name, approved, reply, raw} alebo null (raw = celý riadok z hárku).
function findReview_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECENZIE);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECENZIE_HLAVICKA.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      return { row: i + 2, id: id, stars: Number(rows[i][2]), text: String(rows[i][3]), name: String(rows[i][4]), approved: rows[i][6] === true, reply: String(rows[i][7] || ''), raw: rows[i] };
    }
  }
  return null;
}

// Odozva tlačidiel na stránkach s potvrdením (schváliť, zmazať, zrušiť termín): po odoslaní formulára sa na
// tlačidle ukáže spinner s textom z atribútu data-busy a druhé odoslanie sa ignoruje (server odpovedá niekoľko
// sekúnd a bez odozvy by sa dalo kliknúť znova). Pri návrate tlačidlom "späť" sa tlačidlo vráti do pôvodného stavu.
const BUSY_BUTTON_CSS = `
.btn.busy{opacity:.85;cursor:wait}
.btn .spin{display:inline-block;width:14px;height:14px;margin-right:9px;vertical-align:-2px;border:2px solid rgba(255,255,255,.35);border-top-color:#fff;border-radius:50%;animation:btnspin .8s linear infinite}
@keyframes btnspin{to{transform:rotate(360deg)}}`;
const BUSY_BUTTON_JS = `
(function () {
  document.querySelectorAll('form').forEach(function (form) {
    var btn = form.querySelector('button[type="submit"]');
    if (!btn) return;
    var label = btn.textContent;
    form.addEventListener('submit', function (ev) {
      if (form.getAttribute('data-sent')) { ev.preventDefault(); return; }
      form.setAttribute('data-sent', '1');
      form.querySelectorAll('textarea').forEach(function (t) { t.readOnly = true; });
      btn.classList.add('busy');
      btn.setAttribute('aria-busy', 'true');
      btn.textContent = '';
      var spin = document.createElement('span');
      spin.className = 'spin';
      btn.appendChild(spin);
      btn.appendChild(document.createTextNode(btn.getAttribute('data-busy') || 'Pracujem...'));
    });
    window.addEventListener('pageshow', function (ev) {
      if (!ev.persisted) return;
      form.removeAttribute('data-sent');
      form.querySelectorAll('textarea').forEach(function (t) { t.readOnly = false; });
      btn.classList.remove('busy');
      btn.removeAttribute('aria-busy');
      btn.textContent = label;
    });
  });
})();`;

// Jednoduchá stránka pre holiča (schválenie recenzie). Všetko dynamické sa escapuje.
function reviewAdminPage_(title, innerHtml) {
  return cardPage_(title, innerHtml, 'Recenzia | Barbar Shop');
}

// Tmavá karta v štýle webu (stránky recenzií aj potvrdenia termínu). title a tabTitle sa escapujú, innerHtml nie.
function cardPage_(title, innerHtml, tabTitle) {
  const html = '<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>' + escapeHtml(tabTitle) + '</title>' +
    '<style>' + fontFaceCss() +
    'body{font-family:Roboto,"Segoe UI",Helvetica,Arial,sans-serif;background:#1a1a1a;color:#e0e0e0;margin:0;padding:20px;display:flex;justify-content:center;align-items:center;min-height:100vh;box-sizing:border-box}' +
    '.card{background:#2b1d16;border:1px solid #f0c419;border-radius:12px;padding:32px;max-width:480px;width:100%;box-shadow:0 10px 30px rgba(0,0,0,.5);box-sizing:border-box}' +
    'h2{color:#f0c419;margin:0 0 18px;text-align:center}.stars{color:#f0c419;font-size:28px;letter-spacing:3px;text-align:center}.stars .off{color:#555}' +
    '.quote{background:#1a1a1a;border-left:4px solid #f0c419;border-radius:6px;padding:12px 15px;margin:16px 0;white-space:pre-line;overflow-wrap:anywhere}' +
    '.who{color:#aaa;font-size:14px;text-align:center}label{display:block;color:#f0c419;font-weight:bold;font-size:14px;margin:18px 0 6px}' +
    'textarea{width:100%;box-sizing:border-box;min-height:80px;background:#1a1a1a;border:1px solid #444;color:#fff;padding:12px;border-radius:6px;font-family:inherit;font-size:16px}' +
    '.btn{display:block;width:100%;box-sizing:border-box;text-align:center;text-decoration:none;background:#2e7d32;color:#fff;border:none;padding:15px;font-size:16px;font-weight:bold;border-radius:6px;cursor:pointer;margin-top:18px}.btn:hover{background:#256528}' +
    '.link{display:block;text-align:center;color:#f0c419;margin-top:16px;font-size:14px}.ok{font-size:60px;color:#4caf50;text-align:center}p{line-height:1.5}' +
    '.photos{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:6px;margin:14px 0}.photos img{width:100%;aspect-ratio:1/1;object-fit:cover;border-radius:6px;display:block}' +
    '.details{background:#1a1a1a;border-left:4px solid #f0c419;border-radius:6px;padding:12px 15px;margin:16px 0;line-height:1.7;overflow-wrap:anywhere}.details b{color:#f0c419}' +
    '.btn.gold{background:#f0c419;color:#1a1a1a}.btn.gold:hover{background:#d9b016}.center{text-align:center}' +
    BUSY_BUTTON_CSS +
    '</style></head><body><div class="card"><h2>' + escapeHtml(title) + '</h2>' + innerHtml + '</div><script>' + BUSY_BUTTON_JS + '</script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(tabTitle).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function starsHtml_(n) {
  let out = '';
  for (let i = 1; i <= 5; i++) out += i <= n ? '<span>&#9733;</span>' : '<span class="off">&#9733;</span>';
  return '<div class="stars">' + out + '</div>';
}

// Krok 1: stránka s náhľadom recenzie (zatiaľ nič nemení, preto ju bezpečne otvorí aj antivírus/skener mailu).
function showApproveReviewPage_(id, token) {
  if (!id || token !== reviewApprovalToken_(id)) return reviewAdminPage_('Neplatný odkaz', '<p>Tento odkaz na schválenie nie je platný.</p>');
  const r = findReview_(id);
  if (!r) return reviewAdminPage_('Recenzia sa nenašla', '<p>Recenzia už neexistuje (možno bola odstránená z tabuľky).</p>');
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + getReviewSheet_().getSheetId();
  if (r.approved) {
    return reviewAdminPage_('Už schválené', '<p>Táto recenzia už je schválená a zobrazuje sa na webe.</p><a class="link" href="' + sheetUrl + '" target="_top">Otvoriť tabuľku</a>');
  }
  return reviewAdminPage_('Schváliť recenziu?',
    starsHtml_(r.stars) +
    '<div class="quote">' + escapeHtml(r.text) + '</div>' +
    reviewPhotoPreviewHtml_(r.raw) +
    '<div class="who">Zobrazí sa ako: <strong>' + escapeHtml(r.name) + '</strong></div>' +
    '<form method="GET" action="' + escapeHtml(ScriptApp.getService().getUrl()) + '" target="_top">' +
    '<input type="hidden" name="action" value="doApproveReview"><input type="hidden" name="id" value="' + escapeHtml(id) + '"><input type="hidden" name="t" value="' + escapeHtml(token) + '">' +
    '<label for="odpoved">Tvoja odpoveď pod recenziou (nepovinné)</label>' +
    '<textarea id="odpoved" name="odpoved" maxlength="500" placeholder="Napr. Ďakujeme, tešíme sa na ďalšiu návštevu!"></textarea>' +
    '<button type="submit" class="btn" data-busy="ZVEREJŇUJEM...">SCHVÁLIŤ A ZVEREJNIŤ</button></form>' +
    '<a class="link" href="' + sheetUrl + '" target="_top">Radšej otvoriť tabuľku</a>');
}

// Krok 2: skutočné schválenie.
function doApproveReview_(id, token, odpoved) {
  if (!id || token !== reviewApprovalToken_(id)) return reviewAdminPage_('Neplatný odkaz', '<p>Tento odkaz na schválenie nie je platný.</p>');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return reviewAdminPage_('Skús znova', '<p>Server je práve vyťažený, skús to o chvíľu.</p>');
  try {
    const r = findReview_(id);
    if (!r) return reviewAdminPage_('Recenzia sa nenašla', '<p>Recenzia už neexistuje.</p>');
    if (r.approved) return reviewAdminPage_('Už schválené', '<p>Táto recenzia už bola schválená.</p>');
    const sheet = getReviewSheet_();
    sheet.getRange(r.row, RECENZIE_STLPEC_SCHVALENE).setValue(true);
    const reply = String(odpoved || '').trim().slice(0, 500);
    if (reply) sheet.getRange(r.row, 8).setValue(safeCell(reply));
    syncReviewPhotoSharing_(r.raw, true);
    invalidatePublicCache_();
  } finally {
    lock.releaseLock();
  }
  return reviewAdminPage_('Hotovo', '<div class="ok">&#10003;</div><p style="text-align:center">Recenzia je schválená a na webe sa zobrazí do pár minút. Môžeš zatvoriť túto stránku.</p>');
}

// Zmazanie recenzie, krok 1: náhľad a potvrdenie (nič sa nemaže, kým sa nepotvrdí).
function showDeleteReviewPage_(id, token) {
  if (!id || token !== reviewToken_(id, 'delete')) return reviewAdminPage_('Neplatný odkaz', '<p>Tento odkaz na zmazanie nie je platný.</p>');
  const r = findReview_(id);
  if (!r) return reviewAdminPage_('Recenzia sa nenašla', '<p>Recenzia už neexistuje (možno už bola zmazaná).</p>');
  return reviewAdminPage_('Zmazať recenziu?',
    starsHtml_(r.stars) +
    '<div class="quote">' + escapeHtml(r.text) + '</div>' +
    reviewPhotoPreviewHtml_(r.raw) +
    '<div class="who">Od: <strong>' + escapeHtml(r.name) + '</strong>' + (r.approved ? ' (zatiaľ zverejnená na webe)' : '') + '</div>' +
    '<p style="text-align:center;color:#ff8a80">Recenzia (aj jej fotky) sa nenávratne odstráni z tabuľky aj z webu.</p>' +
    '<form method="GET" action="' + escapeHtml(ScriptApp.getService().getUrl()) + '" target="_top">' +
    '<input type="hidden" name="action" value="doDeleteReview"><input type="hidden" name="id" value="' + escapeHtml(id) + '"><input type="hidden" name="t" value="' + escapeHtml(token) + '">' +
    '<button type="submit" class="btn" style="background:#d32f2f" data-busy="MAŽEM...">ZMAZAŤ NAVŽDY</button></form>');
}

// Krok 2: skutočné zmazanie riadku.
function doDeleteReview_(id, token) {
  if (!id || token !== reviewToken_(id, 'delete')) return reviewAdminPage_('Neplatný odkaz', '<p>Tento odkaz na zmazanie nie je platný.</p>');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return reviewAdminPage_('Skús znova', '<p>Server je práve vyťažený, skús to o chvíľu.</p>');
  try {
    const r = findReview_(id);
    if (!r) return reviewAdminPage_('Recenzia sa nenašla', '<p>Recenzia už neexistuje.</p>');
    deleteReviewPhotos_(r.raw);
    getReviewSheet_().deleteRow(r.row);
    invalidatePublicCache_();
  } finally {
    lock.releaseLock();
  }
  return reviewAdminPage_('Zmazané', '<div class="ok">&#10003;</div><p style="text-align:center">Recenzia bola odstránená. Môžeš zatvoriť túto stránku.</p>');
}

function generateNewReviewEmailHtml_(name, stars, text, approveUrl, deleteUrl, sheetUrl, pocetFotiek) {
  let hviezdy = '';
  for (let i = 1; i <= 5; i++) {
    hviezdy += '<span style="color:' + (i <= stars ? '#f0c419' : '#d8d8d8') + ';">&#9733;</span>';
  }
  const btn = function(href, label, bg, color) {
    return '<a href="' + href + '" target="_blank" style="display:inline-block;margin:6px;padding:14px 22px;background-color:' + bg + ';color:' + color + ';font-family:Arial,sans-serif;font-size:14px;font-weight:bold;text-decoration:none;border-radius:6px;">' + label + '</a>';
  };
  return '<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background-color:#f4f4f5;font-family:Arial,sans-serif;">' +
    '<table border="0" cellpadding="0" cellspacing="0" width="100%" style="padding:20px 0;"><tr><td align="center">' +
    '<table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color:#ffffff;border-radius:12px;overflow:hidden;border-collapse:separate;max-width:100%;">' +
    '<tr><td align="center" style="background-color:#1a1a1a;padding:30px;border-bottom:4px solid #f0c419;"><img src="https://raw.githubusercontent.com/Neonka-Svk/barbarshop/refs/heads/main/barbar_logo_small.png" alt="Barbar Shop" width="150" style="display:block;max-width:150px;"></td></tr>' +
    '<tr><td style="padding:36px 40px;color:#333333;line-height:1.6;font-size:16px;">' +
    '<h1 style="color:#1a1a1a;font-size:24px;margin:0 0 6px 0;text-align:center;">Nová recenzia</h1>' +
    '<p style="margin:0 0 20px 0;text-align:center;color:#888888;font-size:14px;">Čaká na tvoje schválenie</p>' +
    '<div style="text-align:center;font-size:34px;letter-spacing:4px;line-height:1.2;margin:0 0 6px 0;">' + hviezdy + '</div>' +
    '<p style="margin:0 0 20px 0;text-align:center;color:#555555;font-size:14px;">od <strong style="color:#1a1a1a;">' + escapeHtml(name) + '</strong> &middot; ' + stars + ' z 5</p>' +
    '<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#fafafa;border:1px solid #eaeaea;border-left:4px solid #f0c419;border-radius:8px;border-collapse:separate;"><tr><td style="padding:18px 20px;color:#1a1a1a;font-size:16px;line-height:1.6;">' +
    escapeHtml(text).replace(/\n/g, '<br>') + '</td></tr></table>' +
    (pocetFotiek > 0 ? '<p style="margin:14px 0 0 0;text-align:center;color:#555555;font-size:14px;">&#128247; K recenzii je priložených <strong>' + pocetFotiek + '</strong> ' + (pocetFotiek === 1 ? 'fotka' : (pocetFotiek < 5 ? 'fotky' : 'fotiek')) + ' (uvidíš ich v náhľade po kliknutí na Schváliť).</p>' : '') +
    '<div style="text-align:center;margin-top:28px;">' +
    btn(approveUrl, 'SCHVÁLIŤ', '#2e7d32', '#ffffff') +
    btn(deleteUrl, 'ZMAZAŤ', '#d32f2f', '#ffffff') +
    btn(sheetUrl, 'OTVORIŤ TABUĽKU', '#1a1a1a', '#f0c419') +
    '</div>' +
    '<p style="margin:22px 0 0 0;text-align:center;color:#888888;font-size:13px;">Po kliknutí na "Schváliť" alebo "Zmazať" uvidíš najprv náhľad a až potvrdením sa niečo vykoná.</p>' +
    '</td></tr>' +
    '<tr><td align="center" style="background-color:#f9f9f9;padding:22px;border-top:1px solid #eaeaea;font-size:12px;color:#999999;line-height:1.5;"><p style="margin:0;">Tento e-mail bol vygenerovaný automaticky systémom Barbar Shop. &copy; 2026 Barbar Shop. Sila a česť.</p></td></tr>' +
    '</table></td></tr></table></body></html>';
}

function formatReviewDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'Europe/Bratislava', 'd.M.yyyy') : String(v);
}

// Len schválené recenzie, bez ID a bez akýchkoľvek kontaktných údajov.
function getPublicReviews_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECENZIE);
  if (!sheet || sheet.getLastRow() < 2) return { count: 0, average: 0, reviews: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECENZIE_HLAVICKA.length).getValues();
  const approved = rows.filter(function(r) { return r[6] === true && Number(r[2]) >= 1; });
  if (approved.length === 0) return { count: 0, average: 0, reviews: [] };
  const sum = approved.reduce(function(s, r) { return s + Number(r[2]); }, 0);
  // Najnovšie prvé podľa času odoslania (stĺpec F, "d.M.yyyy HH:mm:ss"); pri zhode ostáva poradie z tabuľky.
  const casOdoslania = function(r) {
    const m = String(r[5]).match(/^(\d+)\.(\d+)\.(\d+)\s+(\d+):(\d+):(\d+)/);
    return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +m[6]).getTime() : 0;
  };
  const zoradene = approved.slice().sort(function(a, b) { return casOdoslania(b) - casOdoslania(a); });
  const reviews = zoradene.slice(0, REVIEW_MAX_ZOBRAZENYCH).map(function(r) {
    return { stars: Number(r[2]), text: String(r[3]), name: String(r[4]), date: formatReviewDate_(r[1]), reply: String(r[7] || ''), photos: splitIds_(r[RECENZIE_STLPEC_FOTKY - 1]) };
  });
  return { count: approved.length, average: Math.round(sum / approved.length * 10) / 10, reviews: reviews };
}

function generateReviewEmailHtml_(meno, datum, reviewUrl) {
  meno = escapeHtml(meno);
  datum = escapeHtml(datum);
  return '<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background-color:#f4f4f5;font-family:Arial,sans-serif;">' +
    '<table border="0" cellpadding="0" cellspacing="0" width="100%" style="padding:20px 0;"><tr><td align="center">' +
    '<table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color:#ffffff;border-radius:12px;overflow:hidden;border-collapse:separate;max-width:100%;">' +
    '<tr><td align="center" style="background-color:#1a1a1a;padding:30px;border-bottom:4px solid #f0c419;"><img src="https://raw.githubusercontent.com/Neonka-Svk/barbarshop/refs/heads/main/barbar_logo_small.png" alt="Barbar Shop" width="150" style="display:block;max-width:150px;"></td></tr>' +
    '<tr><td style="padding:40px;color:#333333;line-height:1.6;font-size:16px;">' +
    '<h1 style="color:#1a1a1a;font-size:24px;margin:0 0 20px 0;text-align:center;">Ako sa ti páčil strih?</h1>' +
    '<p style="margin:0 0 15px 0;">Zdravím ťa, <strong>' + meno + '</strong>,</p>' +
    '<p style="margin:0 0 25px 0;">ďakujeme za návštevu (' + datum + '). Ak ti chvíľku zostane, napíš nám krátku recenziu. Pomôže to ďalším bojovníkom a nám ukáže, čo robíme dobre. Zaberie to minútu.</p>' +
    '<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td align="center"><table border="0" cellpadding="0" cellspacing="0"><tr><td align="center" bgcolor="#f0c419" style="border-radius:6px;">' +
    '<a href="' + reviewUrl + '" target="_blank" style="font-size:15px;font-family:Arial,sans-serif;color:#1a1a1a;text-decoration:none;border-radius:6px;padding:14px 28px;display:inline-block;font-weight:bold;">NAPÍSAŤ RECENZIU</a>' +
    '</td></tr></table></td></tr></table>' +
    '<p style="margin:25px 0 0 0;font-size:13px;color:#888888;">Odkaz platí ' + REVIEW_PLATNOST_DNI + ' dní. Recenzia sa na webe zobrazí až po schválení a len s menom, ktoré si zvolíš.</p>' +
    '</td></tr>' +
    '<tr><td align="center" style="background-color:#f9f9f9;padding:25px;border-top:1px solid #eaeaea;font-size:12px;color:#999999;line-height:1.5;"><p style="margin:0 0 10px 0;">Tento e-mail bol vygenerovaný automaticky systémom Barbar Shop.<br>Prosíme, neodpovedajte naň.</p><p style="margin:0;">&copy; 2026 Barbar Shop. Sila a česť.</p></td></tr>' +
    '</table></td></tr></table></body></html>';
}

// Spúšťa denný trigger (installReviewTrigger). Pošle jeden mail s odkazom na recenziu ku každej
// rezervácii z uplynulých dní, ktorá ešte mail nedostala.
function sendReviewRequests() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return;
    if (!sheet.getRange(1, REVIEW_STLPEC_MAILU).getValue()) {
      sheet.getRange(1, REVIEW_STLPEC_MAILU).setValue('Mail s recenziou');
      sheet.getRange(1, 9).copyFormatToRange(sheet, REVIEW_STLPEC_MAILU, REVIEW_STLPEC_MAILU, 1, 1);
    }

    const data = sheet.getDataRange().getDisplayValues();
    const dnes = bratislavaNow_();
    dnes.setHours(0, 0, 0, 0);
    const baseUrl = ScriptApp.getService().getUrl();

    for (let i = 1; i < data.length; i++) {
      if (data[i][REVIEW_STLPEC_MAILU - 1]) continue;
      const ymd = slovakDateToYMD_(data[i][0]);
      const id = data[i][7];
      const email = data[i][5];
      if (!ymd || !id || !email) continue;

      const d = ymd.split('-').map(Number);
      const dniOdTerminu = Math.round((dnes.getTime() - new Date(d[0], d[1] - 1, d[2]).getTime()) / 86400000);
      if (dniOdTerminu < 1 || dniOdTerminu > 5) continue;

      // Rezervujeme si časť dennej kvóty pre potvrdenia rezervácií.
      // (neodoslané ostanú neoznačené a skúsia sa pri ďalšom dennom behu)
      if (mailQuotaLow_()) break;

      try {
        MailApp.sendEmail({ name: MAIL_NAZOV,
          to: email,
          subject: 'Ako sa ti páčil strih? - Barbar Shop',
          htmlBody: generateReviewEmailHtml_(data[i][3], data[i][0], baseUrl + '?action=reviewPage&id=' + id)
        });
        sheet.getRange(i + 1, REVIEW_STLPEC_MAILU).setValue('POSLANÉ');
      } catch (err) {
        console.error('Mail s recenziou sa neodoslal (riadok ' + (i + 1) + '):', err);
      }
    }
  } catch (err) {
    console.error('sendReviewRequests zlyhalo:', err);
  }
}

// Spusti RAZ ručne z editora (po tom, čo nasadíš tento kód): založí denný trigger okolo 13:00 bratislavského času.
function installReviewTrigger() {
  const uz = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'sendReviewRequests'; });
  if (uz) { console.log('Trigger sendReviewRequests už existuje.'); return; }
  ScriptApp.newTrigger('sendReviewRequests').timeBased().everyDays(1).atHour(13).inTimezone('Europe/Bratislava').create();
  console.log('Trigger založený: sendReviewRequests, denne okolo 13:00.');
}

// Spusti RAZ ručne z editora pod účtom, z ktorého majú chodiť maily (roman.barbarshop): založí VŠETKY triggery
// projektu pod týmto účtom. Triggery bežia (a maily posielajú) pod tým, kto ich založil. Staré triggery tohto
// účtu sa najprv zmažú, takže opakované spustenie nič nezdvojí. Triggery INÉHO účtu odtiaľto zmazať nejde:
// treba ich zmazať v jeho profile (Apps Script -> Spúšťače).
function installAllTriggers() {
  const moje = ['handleEdit', 'checkUpcomingConflicts', 'cleanupOldReservations', 'sendReviewRequests'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (moje.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('checkUpcomingConflicts').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('cleanupOldReservations').timeBased().everyDays(1).atHour(0).inTimezone('Europe/Bratislava').create();
  ScriptApp.newTrigger('sendReviewRequests').timeBased().everyDays(1).atHour(13).inTimezone('Europe/Bratislava').create();
  installSpeedTriggers(); // cache, nepotvrdené rezervácie, večerný súhrn
  console.log('Všetky triggery založené pod účtom ' + Session.getEffectiveUser().getEmail() + ': handleEdit (pri úprave), checkUpcomingConflicts (každých 10 min), cleanupOldReservations (denne ~0:00), sendReviewRequests (denne ~13:00) + triggery z installSpeedTriggers.');
}

// ===================== GALÉRIA PRÁC (Google Disk) =====================
// Holič nahrá fotky/videá do priečinka "web_photos" na Disku účtu, pod ktorým skript beží. Priečinok treba
// zdieľať ako "Ktokoľvek s odkazom: Prezeranie", súbory v ňom to zdedia (setupDriveFolders() to nastaví samo).
// Všetky fotky a videá z priečinka (aj z jeho podpriečinkov) idú do jednej všeobecnej galérie.
// Zoznam sa na GALERIA_CACHE_SEKUND ukladá do cache, aby sa Disk nevolal pri každom načítaní stránky
// (menu "Obnoviť galériu na webe" cache zmaže hneď).
const GALERIA_PRIECINOK_NAZOV = 'web_photos';
const GALERIA_PRIECINOK_ID = ''; // voliteľné: ID priečinka z jeho URL, ak by existovalo viac priečinkov s rovnakým názvom
const GALERIA_MAX = 24;
const GALERIA_CACHE_SEKUND = 900;
const GALERIA_CACHE_KLUC = 'gallery_v2';

// Berie sa len priečinok, ktorého vlastníkom je tento účet (priečinok s rovnakým názvom, ktorý by nám niekto
// cudzí zdieľal, sa ignoruje, aby nešiel na web cudzí obsah). Výnimka: ak je nastavené GALERIA_PRIECINOK_ID.
function getGalleryFolder_() {
  if (GALERIA_PRIECINOK_ID) return DriveApp.getFolderById(GALERIA_PRIECINOK_ID);
  return getOwnedFolder_(GALERIA_PRIECINOK_NAZOV, false);
}

// Spusti RAZ ručne z editora: vytvorí priečinky na Disku (ak ešte neexistujú). "web_photos" sa nastaví ako
// verejne prezeraný cez odkaz (ide z neho galéria na webe), "reviews" ostáva súkromný (fotky z recenzií).
function setupDriveFolders() {
  const galeria = getOwnedFolder_(GALERIA_PRIECINOK_NAZOV, true);
  galeria.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const recenzie = getOwnedFolder_(REVIEWS_PRIECINOK_NAZOV, true);
  recenzie.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  console.log('Priečinok web_photos: ' + galeria.getUrl());
  console.log('Priečinok reviews: ' + recenzie.getUrl());
}

function collectGalleryFiles_(folder, out) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.isTrashed()) continue;
    const mime = String(f.getMimeType());
    const type = /^image\//.test(mime) ? 'image' : (/^video\//.test(mime) ? 'video' : null);
    if (!type) continue;
    out.push({ id: f.getId(), type: type, name: f.getName().replace(/\.[^.]+$/, ''), updated: f.getLastUpdated().getTime() });
  }
}

// Vráti {items: [{id, type: 'image'|'video', name}]}, najnovšie prvé. Berú sa súbory priamo v priečinku
// aj v jeho podpriečinkoch (jedna úroveň), všetky idú do jednej všeobecnej galérie.
function getGalleryItems_(vynutit) {
  const cache = CacheService.getScriptCache();
  if (!vynutit) {
    const cached = cache.get(GALERIA_CACHE_KLUC);
    if (cached) return JSON.parse(cached);
  }

  let items = [];
  const folder = getGalleryFolder_();
  if (folder) {
    collectGalleryFiles_(folder, items);
    const subs = folder.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      if (!sub.isTrashed()) collectGalleryFiles_(sub, items);
    }
    items.sort(function(a, b) { return b.updated - a.updated; });
    items = items.slice(0, GALERIA_MAX);
  }
  const result = { items: items.map(function(i) { return { id: i.id, type: i.type, name: i.name }; }) };
  try { cache.put(GALERIA_CACHE_KLUC, JSON.stringify(result), GALERIA_CACHE_SEKUND); } catch (err) {}
  return result;
}

function clearGalleryCache() {
  CacheService.getScriptCache().remove(GALERIA_CACHE_KLUC);
  invalidatePublicCache_();
  let hotovo = true;
  try { refreshPublicCache(); } catch (err) { hotovo = false; console.error('Obnova galérie zlyhala:', err); }
  try { SpreadsheetApp.getActiveSpreadsheet().toast(hotovo ? 'Galéria na webe je obnovená.' : 'Galéria sa na webe obnoví pri najbližšom načítaní stránky.', 'Galéria', 5); } catch (err) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ Správa rezervácií')
    .addItem('Vymazať staré rezervácie (7+ dni)', 'cleanupOldReservations')
    .addItem('Obnoviť galériu na webe', 'clearGalleryCache')
    .addToUi();

  // Staršiemu hárku "Recenzie" doplní stĺpec "Zmazať recenziu?" (hlavička + checkboxy).
  try {
    const recenzie = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECENZIE);
    if (recenzie) upgradeReviewSheet_(recenzie);
  } catch (err) {
    console.error('Úprava hárku Recenzie zlyhala:', err);
  }
}

// ===================== POTVRDENIE REZERVÁCIE E-MAILOM =====================
// Tok: objednávka -> riadok v hárku "Nepotvrdené" (termín je pre ostatných hneď obsadený) -> na stránke beží
// krátky odpočet, počas ktorého sa dá opraviť e-mail alebo rezervácia zrušiť (ako "Vrátiť späť" v Gmaile) ->
// odíde mail s tlačidlom POTVRDIŤ TERMÍN -> po potvrdení sa riadok presunie do "Rezervácie" a až vtedy príde
// mail holičovi. Bez potvrdenia do POTVRDENIE_LEHOTA_MINUT sa termín uvoľní a riadok ide do archívu ako
// "Nepotvrdená" (vymyslený alebo zle zadaný e-mail tak nič neblokuje).
// Zavretá stránka počas odpočtu nevadí: trigger processPendingReservations (každú minútu) mail pošle sám.
// Všetky zmeny v hárku "Nepotvrdené" idú pod zámkom a riadky sa čítajú až v ňom (čísla riadkov sa pri mazaní posúvajú).
const SHEET_CAKAJUCE = 'Nepotvrdené';
const CAKAJUCE_HLAVICKA = ['Dátum strihania', 'Objednaný čas', 'Čas objednávky', 'Meno', 'Služba', 'E-mail', 'Identifikátor rezervácie', 'Stav', 'Odoslať mail po (ms)', 'Mail odoslaný (ms)', 'Kód potvrdenia', 'Kľúč prehliadača (odtlačok)', 'Zmeny e-mailu', 'Potvrdiť ručne'];
const CAKAJUCE_STLPEC_POTVRDIT = 14; // checkbox: holič potvrdí rezerváciu sám (napr. kamarátovi bez mailu)
const CAKAJUCE_STAV_NEODOSLANE = 'ČAKÁ NA ODOSLANIE MAILU';
const CAKAJUCE_STAV_ODOSLANE = 'ČAKÁ NA POTVRDENIE';
const POTVRDENIE_ODKLAD_SEKUND = 15;      // odpočet na stránke pred odoslaním mailu (čas na opravu e-mailu)
const POTVRDENIE_LEHOTA_MINUT = 15;       // koľko minút od odoslania mailu má zákazník na potvrdenie
const POTVRDENIE_MAX_ZMIEN = 2;           // koľkokrát sa dá zmeniť e-mail
const POTVRDENIE_NEODOSLANE_MAX_MINUT = 5; // poistka: mail, ktorý z nejakého dôvodu neodišiel, termín dlhšie nedrží
// Trigger pošle mail za stránku až po tejto rezerve od konca odpočtu: kým zákazník píše opravený e-mail, odpočet
// na stránke stojí a mail nesmie odísť na pôvodnú adresu. Pri zavretej stránke tak mail príde do ~3 minút.
const POTVRDENIE_TRIGGER_REZERVA_SEKUND = 120;
const ARCHIV_STAV_NEPOTVRDENA = 'Nepotvrdená';
const ARCHIV_NEPOTVRDENE_ANONYMIZOVAT_PO_DNOCH = 30;
const UUID_REGEX = /^[0-9a-f-]{36}$/i;

function getPendingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CAKAJUCE);
  if (!sheet) {
    const povodny = ss.getActiveSheet();
    sheet = ss.insertSheet(SHEET_CAKAJUCE);
    sheet.getRange(1, 1, 1, CAKAJUCE_HLAVICKA.length).setValues([CAKAJUCE_HLAVICKA]).setFontWeight('bold');
    sheet.getRange(1, 1, sheet.getMaxRows(), 2).setNumberFormat('@'); // dátum a čas ako text, nech ich Sheets nepreklopí
    sheet.setFrozenRows(1);
    try { ss.setActiveSheet(povodny); } catch (err) {}
  } else if (sheet.getLastColumn() < CAKAJUCE_HLAVICKA.length) {
    // starší hárok bez stĺpca Potvrdiť ručne: doplní sa hlavička
    sheet.getRange(1, CAKAJUCE_STLPEC_POTVRDIT).setValue(CAKAJUCE_HLAVICKA[CAKAJUCE_STLPEC_POTVRDIT - 1]).setFontWeight('bold');
  }
  return sheet;
}

function hashKey_(key) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(key), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}

// Hodnota bunky ako text bez apostrofu, ktorý pridáva safeCell (a dátum/čas aj vtedy, keby ho Sheets preklopili na Date).
function plain_(v, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Bratislava', fmt || 'dd.MM.yyyy');
  return String(v == null ? '' : v).replace(/^'/, '').trim();
}

// {sheet, rows: [{row: číslo riadku v hárku, r: hodnoty}]}
function readPending_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CAKAJUCE);
  if (!sheet || sheet.getLastRow() < 2) return { sheet: sheet, rows: [] };
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, CAKAJUCE_HLAVICKA.length).getValues();
  return { sheet: sheet, rows: vals.map(function(r, i) { return { row: i + 2, r: r }; }) };
}

function findPending_(id) {
  const data = readPending_();
  const p = data.rows.filter(function(x) { return String(x.r[6]) === String(id); })[0] || null;
  return { sheet: data.sheet, p: p };
}

function pendingTimes_(r) {
  return { sendAfter: Number(r[8]) || 0, sentAt: Number(r[9]) || 0 };
}

// Drží ešte termín? Po odoslaní mailu POTVRDENIE_LEHOTA_MINUT, pred odoslaním najviac POTVRDENIE_NEODOSLANE_MAX_MINUT.
function isPendingActive_(r, now) {
  const t = pendingTimes_(r);
  if (t.sentAt) return now < t.sentAt + POTVRDENIE_LEHOTA_MINUT * 60000;
  return now < t.sendAfter + POTVRDENIE_NEODOSLANE_MAX_MINUT * 60000;
}

// Aktívne čakajúce rezervácie ako [{datum: 'yyyy-mm-dd', cas: 'HH:mm'}] (pre kalendár a kontrolu obsadenosti).
// Vypršané sa ignorujú hneď, aj keby ich trigger ešte neodpratal.
function activePendingBookings_(now) {
  const out = [];
  readPending_().rows.forEach(function(p) {
    if (!isPendingActive_(p.r, now)) return;
    const ymd = slovakDateToYMD_(plain_(p.r[0]));
    if (ymd) out.push({ datum: ymd, cas: plain_(p.r[1], 'HH:mm') });
  });
  return out;
}

// Čo sa stalo s rezerváciou, ktorá už nečaká: potvrdená / zrušená (panel, odkaz v maile) / vypršala.
function finishedState_(id) {
  if (isConfirmed_(id)) return 'confirmed';
  const arch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ARCHIV);
  if (arch && arch.getLastRow() > 1) {
    const data = arch.getRange(2, 1, arch.getLastRow() - 1, ARCHIV_HLAVICKA.length).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][6]) === String(id)) return /^Zrušená/.test(String(data[i][8])) ? 'cancelled' : 'expired';
    }
  }
  return 'gone';
}

function isConfirmed_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return false;
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === id) return true;
  }
  return false;
}

// Je termín obsadený potvrdenou alebo ešte platnou čakajúcou rezerváciou? Volať pod zámkom.
function isSlotTaken_(slovakDate, cas, now) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const existujuce = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < existujuce.length; i++) {
    const existCas = String(existujuce[i][1]).replace(/^'/, '').trim();
    if (existujuce[i][0] === slovakDate && existCas === cas) return true;
  }
  const ymd = slovakDateToYMD_(slovakDate);
  return activePendingBookings_(now).some(function(b) { return b.datum === ymd && b.cas === cas; });
}

// Riadok z "Nepotvrdené" do archívu (rovnaký tvar ako rezervácia A-H) a preč z hárku. Volať pod zámkom.
function releasePending_(sheet, p, dovod) {
  const r = p.r;
  archiveReservation([plain_(r[0]), plain_(r[1], 'HH:mm'), r[2], plain_(r[3]), plain_(r[4]), plain_(r[5]), false, r[6]], ARCHIV_STAV_NEPOTVRDENA, dovod);
  sheet.deleteRow(p.row);
}

function sendPendingMail_(r) {
  const id = r[6];
  const baseUrl = ScriptApp.getService().getUrl();
  const datum = plain_(r[0]);
  const cas = plain_(r[1], 'HH:mm');
  const sablona = HtmlService.createTemplateFromFile('EmailSablona');
  sablona.titulok = 'Potvrď svoj termín';
  sablona.meno = plain_(r[3]);
  sablona.uvodnyText = 'Ešte jeden klik a termín je tvoj. Potvrď ho tlačidlom nižšie do ' + POTVRDENIE_LEHOTA_MINUT + ' minút, inak sa uvoľní pre ďalších bojovníkov.';
  sablona.sluzba = plain_(r[4]);
  sablona.datum = datum;
  sablona.cas = cas;
  // Pozor: parametre 'c' a 'sid' má Apps Script vyhradené (Google URL s nimi ani nepustí ku skriptu), preto 'kod'.
  sablona.confirmUrl = baseUrl + '?action=confirmPage&id=' + id + '&kod=' + r[10];
  sablona.cancelUrl = baseUrl + '?action=cancelPage&id=' + id + '&role=klient';
  sablona.viewUrl = baseUrl + '?action=viewEmail&id=' + id;
  sablona.vyzvaText = 'Ak si o tento termín nežiadal alebo ho chceš zrušiť, klikni na tlačidlo nižšie:';
  sablona.jeHolic = false;
  MailApp.sendEmail({ name: MAIL_NAZOV,
    to: plain_(r[5]),
    subject: 'Barbar Shop - Potvrď termín (' + datum + ' o ' + cas + ')',
    htmlBody: sablona.evaluate().getContent()
  });
}

// Pošle mail s odkazom na potvrdenie a zapíše čas odoslania (od neho beží lehota). Vráti chybový text alebo null.
function deliverPendingMail_(sheet, p) {
  try {
    sendPendingMail_(p.r);
  } catch (err) {
    console.error('Mail s potvrdením termínu sa nepodarilo odoslať:', err);
    return 'Mail sa nepodarilo odoslať. Skús to o chvíľu znova.';
  }
  const odoslane = Date.now();
  sheet.getRange(p.row, 8).setValue(CAKAJUCE_STAV_ODOSLANE);
  sheet.getRange(p.row, 10).setValue(odoslane);
  p.r[7] = CAKAJUCE_STAV_ODOSLANE;
  p.r[9] = odoslane;
  return null;
}

function sendBarberNewBookingMail_(meno, sluzba, datum, cas, email, id) {
  const baseUrl = ScriptApp.getService().getUrl();
  const sablona = HtmlService.createTemplateFromFile('EmailSablona');
  sablona.titulok = "Nová rezervácia!";
  sablona.meno = "Roman";
  sablona.uvodnyText = "Máš nového bojovníka v kresle! Tu sú detaily:";
  sablona.sluzba = sluzba;
  sablona.datum = datum;
  sablona.cas = cas;
  sablona.emailKlienta = email;
  sablona.cancelUrl = `${baseUrl}?action=cancelPage&id=${id}&role=holic`;
  sablona.viewUrl = `${baseUrl}?action=viewEmail&id=${id}`;
  sablona.vyzvaText = "Ak potrebuješ tento termín zrušiť z prevádzkových alebo iných dôvodov, klikni nižšie na tlačidlo:";
  sablona.jeHolic = true;
  sendBarberMailOrDefer_(
    { subject: `Nová rezervácia: ${meno} (${datum})`, htmlBody: sablona.evaluate().getContent() },
    { typ: 'Nové rezervácie', text: `${meno}, ${datum} o ${cas}, ${sluzba} (${email})`, odkazy: [{ t: 'Zrušiť', u: sablona.cancelUrl }] }
  );
}

// Stav pre stránku. Časy sú relatívne (ms odteraz), aby nezáležalo na tom, ako presne idú hodiny u zákazníka.
function pendingState_(p, now) {
  const t = pendingTimes_(p.r);
  return {
    state: t.sentAt ? 'sent' : 'waiting',
    email: plain_(p.r[5]),
    sendInMs: t.sentAt ? 0 : Math.max(0, t.sendAfter - now),
    expiresInMs: t.sentAt ? Math.max(0, t.sentAt + POTVRDENIE_LEHOTA_MINUT * 60000 - now) : null,
    changesLeft: Math.max(0, POTVRDENIE_MAX_ZMIEN - (Number(p.r[12]) || 0))
  };
}

// Akcie z panela na stránke. Overuje sa kľúč, ktorý dostal len prehliadač, ktorý rezerváciu vytvoril.
// Potvrdiť sa odtiaľto nedá: to ide len odkazom z mailu (inak by sa neoverilo, že schránka existuje).
function handlePendingAction_(data) {
  const id = String(data.id || '');
  const key = String(data.key || '');
  const action = String(data.action || '');
  if (!UUID_REGEX.test(id) || key.length < 20 || key.length > 200) return { error: 'Neplatná požiadavka.' };
  if (['status', 'send', 'changeEmail', 'cancel'].indexOf(action) === -1) return { error: 'Neznáma akcia.' };

  // Stav sa pýta často (stránka sa každých pár sekúnd pozrie, či už zákazník klikol v maile), preto bez zámku.
  if (action === 'status') {
    const found = findPending_(id);
    if (!found.p) return { state: finishedState_(id) };
    if (String(found.p.r[11]) !== hashKey_(key)) return { error: 'Neplatná relácia.' };
    const now = Date.now();
    if (!isPendingActive_(found.p.r, now)) return { state: 'expired' };
    return pendingState_(found.p, now);
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { error: 'Server je momentálne vyťažený, skús to prosím o chvíľu znova.' };
  let zmena = false;
  try {
    const now = Date.now();
    const found = findPending_(id);
    if (!found.p) return { state: finishedState_(id) };
    if (String(found.p.r[11]) !== hashKey_(key)) return { error: 'Neplatná relácia.' };

    if (!isPendingActive_(found.p.r, now)) {
      releasePending_(found.sheet, found.p, 'Nepotvrdená včas');
      zmena = true;
      return { state: 'expired' };
    }

    if (action === 'cancel') {
      releasePending_(found.sheet, found.p, 'Zrušená zákazníkom pred potvrdením');
      zmena = true;
      return { state: 'cancelled' };
    }

    if (action === 'changeEmail') {
      const email = String(data.email || '').trim();
      if (email.length > 200 || !EMAIL_REGEX.test(email)) return { error: 'Zadaj platný e-mail.' };
      const zmeny = Number(found.p.r[12]) || 0;
      if (zmeny >= POTVRDENIE_MAX_ZMIEN) return { error: 'E-mail sa už nedá znova zmeniť. Zruš rezerváciu a vytvor novú.' };
      // Nový kód: odkaz z mailu na pôvodnú (zlú) adresu prestane platiť. Odpočet beží odznova.
      const novyKod = Utilities.getUuid();
      found.sheet.getRange(found.p.row, 6).setValue(safeCell(email));
      found.sheet.getRange(found.p.row, 8, 1, 6).setValues([[CAKAJUCE_STAV_NEODOSLANE, now + POTVRDENIE_ODKLAD_SEKUND * 1000, '', novyKod, found.p.r[11], zmeny + 1]]);
      found.p.r[5] = email;
      found.p.r[7] = CAKAJUCE_STAV_NEODOSLANE;
      found.p.r[8] = now + POTVRDENIE_ODKLAD_SEKUND * 1000;
      found.p.r[9] = '';
      found.p.r[10] = novyKod;
      found.p.r[12] = zmeny + 1;
      return pendingState_(found.p, now);
    }

    // action === 'send': odpočet na stránke skončil (alebo zákazník klikol "odoslať hneď")
    if (!pendingTimes_(found.p.r).sentAt) {
      const chyba = deliverPendingMail_(found.sheet, found.p);
      if (chyba) return { error: chyba };
    }
    return pendingState_(found.p, Date.now());
  } finally {
    lock.releaseLock();
    if (zmena) invalidatePublicCache_();
  }
}

// Trigger každú minútu (installSpeedTriggers): pošle maily, ktoré stránka neposlala (zavretá počas odpočtu),
// a uvoľní termíny, ktoré nikto nepotvrdil. Keď nič nečaká, skončí hneď po jednom prečítaní vlastnosti.
function processPendingReservations() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PENDING_ANY')) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  let zmena = false;
  try {
    const now = Date.now();
    const data = readPending_();
    let zostava = 0;
    for (let i = data.rows.length - 1; i >= 0; i--) { // odspodu, nech mazanie neposúva riadky, ktoré ešte prídu na rad
      const p = data.rows[i];
      if (!p.r[6]) continue; // prázdny riadok
      if (!isPendingActive_(p.r, now)) {
        releasePending_(data.sheet, p, pendingTimes_(p.r).sentAt ? 'Nepotvrdená do ' + POTVRDENIE_LEHOTA_MINUT + ' minút' : 'Mail s potvrdením neodišiel');
        zmena = true;
        continue;
      }
      zostava++;
      const t = pendingTimes_(p.r);
      if (!t.sentAt && now >= t.sendAfter + POTVRDENIE_TRIGGER_REZERVA_SEKUND * 1000) deliverPendingMail_(data.sheet, p);
    }
    if (zostava === 0) props.deleteProperty('PENDING_ANY');
  } finally {
    lock.releaseLock();
  }
  if (zmena) invalidatePublicCache_();
}

// --- Potvrdenie odkazom z mailu (2 kroky, aby rezerváciu nepotvrdil skener odkazov v pošte) ---
function bookingDetailsHtml_(r) {
  return '<div class="details"><b>Bojovník:</b> ' + escapeHtml(plain_(r[3])) + '<br><b>Služba:</b> ' + escapeHtml(plain_(r[4])) +
    '<br><b>Termín:</b> ' + escapeHtml(plain_(r[0])) + ' o ' + escapeHtml(plain_(r[1], 'HH:mm')) + '<br><b>Kde:</b> ' + escapeHtml(MIESTO) + '</div>';
}

function bookingPage_(title, innerHtml) {
  return cardPage_(title, innerHtml, 'Potvrdenie termínu | Barbar Shop');
}

// Výsledky potvrdenia ako {title, inner}: rovnaký obsah ide do celej stránky aj do odpovede pre google.script.run.
function confirmResult_(title, inner) {
  return { title: title, inner: inner };
}

function expiredConfirmResult_() {
  return confirmResult_('Čas na potvrdenie vypršal',
    '<p class="center">Termín sa uvoľnil, lebo nebol potvrdený do ' + POTVRDENIE_LEHOTA_MINUT + ' minút (alebo bola rezervácia zrušená). Ak ho stále chceš, vyber si ho na webe znova.</p>' +
    '<a class="btn gold" href="' + SITE_URL + '" target="_top">OBJEDNAŤ ZNOVA</a>');
}

// Spoločná kontrola pre oba kroky. Vráti {vysledok: {title, inner}} (niečo nesedí) alebo {sheet, p}.
function findConfirmable_(id, code, bezKodu) {
  id = String(id || '');
  if (!UUID_REGEX.test(id)) return { vysledok: confirmResult_('Neplatný odkaz', '<p class="center">Tento odkaz nie je platný.</p>') };
  if (isConfirmed_(id)) {
    return { vysledok: confirmResult_('Termín je už potvrdený', '<div class="ok">&#10003;</div><p class="center">Tento termín si už potvrdil. Tešíme sa na teba!</p>') };
  }
  const found = findPending_(id);
  if (!found.p || !isPendingActive_(found.p.r, Date.now())) return { vysledok: expiredConfirmResult_() };
  if (!bezKodu && String(found.p.r[10]) !== String(code || '')) {
    return { vysledok: confirmResult_('Tento odkaz už neplatí', '<p class="center">E-mail pri tejto rezervácii sa medzitým zmenil. Použi odkaz z najnovšieho mailu.</p>') };
  }
  return found;
}

// Krok 1 (odkaz z mailu). Krok 2 ide cez google.script.run priamo z tejto stránky, nie formulárom: pri prehliadači
// prihlásenom do viacerých Google účtov Google presmerovanie formulára na adresu skriptu niekedy zle nasmeruje
// ("Súbor momentálne nie je možné otvoriť"). Kód ide do stránky cez jsonForScript (bezpečne v <script>).
function showConfirmPage_(id, code) {
  const f = findConfirmable_(id, code);
  if (f.vysledok) return bookingPage_(f.vysledok.title, f.vysledok.inner);
  return bookingPage_('Potvrdiť termín?',
    bookingDetailsHtml_(f.p.r) +
    '<button type="button" class="btn" id="potvrdit">POTVRDIŤ TERMÍN</button>' +
    '<p class="center" id="chyba" style="color:#ff8a80;min-height:1em"></p>' +
    '<script>(function () {' +
    'var ID = ' + jsonForScript(String(id)) + ', KOD = ' + jsonForScript(String(code)) + ';' +
    'var b = document.getElementById("potvrdit"), chyba = document.getElementById("chyba");' +
    'b.addEventListener("click", function () {' +
    '  if (b.disabled) return;' +
    '  b.disabled = true; b.classList.add("busy"); b.textContent = "";' +
    '  var s = document.createElement("span"); s.className = "spin"; b.appendChild(s); b.appendChild(document.createTextNode("POTVRDZUJEM..."));' +
    '  chyba.textContent = "";' +
    '  google.script.run.withSuccessHandler(function (v) {' +
    '    var card = document.querySelector(".card"); card.innerHTML = "";' +
    '    var h = document.createElement("h2"); h.textContent = v.title; card.appendChild(h);' +
    '    card.insertAdjacentHTML("beforeend", v.inner);' +
    '  }).withFailureHandler(function () {' +
    '    b.disabled = false; b.classList.remove("busy"); b.textContent = "POTVRDIŤ TERMÍN";' +
    '    chyba.textContent = "Spojenie zlyhalo, skús to prosím znova.";' +
    '  }).confirmBookingFromPage(ID, KOD);' +
    '});' +
    '})();</script>');
}

// Volá stránka z kroku 1 (google.script.run). Bez správneho kódu z mailu nič neurobí.
function confirmBookingFromPage(id, code) {
  return confirmBooking_(id, code);
}

// Starší tvar kroku 2 (formulár / priamy odkaz), keby niekto mal otvorenú stránku spred zmeny.
function doConfirm_(id, code) {
  const v = confirmBooking_(id, code);
  return bookingPage_(v.title, v.inner);
}

// bezKodu = ručné potvrdenie holičom v tabuľke (checkbox), inak sa vyžaduje kód z mailu.
function confirmBooking_(id, code, bezKodu) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return confirmResult_('Skús znova', '<p class="center">Server je práve vyťažený, skús to o chvíľu.</p>');
  let r;
  try {
    const f = findConfirmable_(id, code, bezKodu);
    if (f.vysledok) return f.vysledok;
    r = f.p.r;
    const datum = plain_(r[0]);
    const cas = plain_(r[1], 'HH:mm');
    const sheetRez = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // Poistka: potvrdená rezervácia na ten istý termín by tu byť nemala (čakajúca ho drží), ale keby predsa...
    const existujuce = sheetRez.getDataRange().getDisplayValues();
    for (let i = 1; i < existujuce.length; i++) {
      if (existujuce[i][0] === datum && String(existujuce[i][1]).replace(/^'/, '').trim() === cas) {
        releasePending_(f.sheet, f.p, 'Termín medzičasom obsadený');
        invalidatePublicCache_();
        return confirmResult_('Termín už nie je voľný', '<p class="center">Tento termín medzičasom obsadil niekto iný. Prepáč, vyber si prosím iný.</p><a class="btn gold" href="' + SITE_URL + '" target="_top">VYBRAŤ INÝ TERMÍN</a>');
      }
    }

    sheetRez.appendRow([datum, "'" + cas, r[2], safeCell(plain_(r[3])), plain_(r[4]), safeCell(plain_(r[5])), false, String(id)]);
    const lastRow = sheetRez.getLastRow();
    sheetRez.getRange(lastRow, 7).insertCheckboxes();
    if (lastRow > 1) {
      sheetRez.getRange(2, 1, lastRow - 1, sheetRez.getLastColumn()).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
    }
    f.sheet.deleteRow(f.p.row);
    invalidatePublicCache_();
  } finally {
    lock.releaseLock();
  }

  try {
    sendBarberNewBookingMail_(plain_(r[3]), plain_(r[4]), plain_(r[0]), plain_(r[1], 'HH:mm'), plain_(r[5]), String(id));
  } catch (err) {
    console.error('Mail holičovi o novej rezervácii sa nepodarilo odoslať:', err);
  }
  return confirmResult_('Termín je potvrdený', '<div class="ok">&#10003;</div>' + bookingDetailsHtml_(r) +
    '<p class="center">Sila a česť! Tešíme sa na teba. Ak budeš potrebovať termín zrušiť, použi tlačidlo v maile.</p>');
}

// --- ARCHÍV REZERVÁCIÍ ---
// Každá rezervácia, ktorá zmizne z hárku "Rezervácie" (prebehla / zrušená), sa najprv skopíruje sem.
// ID rezervácie ostáva v archíve navždy (budú sa podľa neho overovať aj recenzie), osobné údaje
// (meno, e-mail) sa po ARCHIV_ANONYMIZOVAT_PO_MESIACOCH mesiacoch automaticky nahradia textom
// "(anonymizované)". Číslo treba držať v súlade s textom na stránke privacy.html.
const SHEET_ARCHIV = 'Archív';
const ARCHIV_HLAVICKA = ['Dátum strihania', 'Objednaný čas', 'Čas objednávky', 'Meno', 'Služba', 'E-mail', 'Identifikátor rezervácie', 'Stav', 'Dôvod', 'Archivované'];
const ARCHIV_ANONYMIZOVAT_PO_MESIACOCH = 24;
const ARCHIV_ANONYMIZOVANE = '(anonymizované)';

// Text začínajúci =, +, - alebo @ by Google Sheets pri zápise vyhodnotil ako vzorec (napr. meno
// "=IMPORTDATA(...)" by vedelo poslať obsah tabuľky von). Predpísaný apostrof z toho spraví obyčajný text.
function safeCell(value) {
  const s = String(value == null ? '' : value);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function getArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ARCHIV);
  if (!sheet) {
    const povodny = ss.getActiveSheet();
    sheet = ss.insertSheet(SHEET_ARCHIV);
    sheet.getRange(1, 1, 1, ARCHIV_HLAVICKA.length).setValues([ARCHIV_HLAVICKA]).setFontWeight('bold');
    sheet.getRange(1, 1, sheet.getMaxRows(), 3).setNumberFormat('@'); // dátum a časy ako text, nech ich Sheets nepreklopí podľa lokality
    sheet.setFrozenRows(1);
    try { ss.setActiveSheet(povodny); } catch (err) {}
  }
  return sheet;
}

// row = riadok z hárku Rezervácie (A-H: dátum, čas, čas objednávky, meno, služba, e-mail, checkbox, ID).
// Vráti true/false, aby sa dalo rozhodnúť, či riadok z Rezervácií zmazať.
function archiveReservation(row, stav, dovod) {
  try {
    getArchiveSheet().appendRow([
      row[0],
      String(row[1]).replace(/^'/, ''),
      row[2],
      safeCell(row[3]),
      safeCell(row[4]),
      safeCell(row[5]),
      row[7] || '',
      stav,
      safeCell(dovod),
      Utilities.formatDate(new Date(), 'Europe/Bratislava', 'd.M.yyyy HH:mm:ss')
    ]);
    return true;
  } catch (err) {
    console.error('Archivácia rezervácie zlyhala:', err);
    return false;
  }
}

function anonymizeOldArchive() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ARCHIV);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ARCHIV_HLAVICKA.length).getDisplayValues();
  const hranica = new Date();
  hranica.setHours(0, 0, 0, 0);
  hranica.setMonth(hranica.getMonth() - ARCHIV_ANONYMIZOVAT_PO_MESIACOCH);

  // Nepotvrdené rezervácie (často s vymysleným alebo cudzím e-mailom) sa anonymizujú už po pár dňoch od archivácie.
  const hranicaNepotvrdene = new Date();
  hranicaNepotvrdene.setHours(0, 0, 0, 0);
  hranicaNepotvrdene.setDate(hranicaNepotvrdene.getDate() - ARCHIV_NEPOTVRDENE_ANONYMIZOVAT_PO_DNOCH);

  data.forEach((r, i) => {
    if (r[5] === ARCHIV_ANONYMIZOVANE) return;
    const nepotvrdena = r[7] === ARCHIV_STAV_NEPOTVRDENA;
    const p = String(nepotvrdena ? String(r[9]).split(' ')[0] : r[0]).replace(/\s/g, '').split('.');
    if (p.length !== 3) return;
    if (new Date(p[2], p[1] - 1, p[0]) < (nepotvrdena ? hranicaNepotvrdene : hranica)) {
      sheet.getRange(i + 2, 4).setValue(ARCHIV_ANONYMIZOVANE); // Meno
      sheet.getRange(i + 2, 6).setValue(ARCHIV_ANONYMIZOVANE); // E-mail
    }
  });
}

function cleanupOldReservations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();

  const dnes = new Date();
  dnes.setHours(0,0,0,0);
  const hranica = new Date(dnes.setDate(dnes.getDate() - 7));

  for (let i = data.length - 1; i >= 1; i--) {
    let p = data[i][0].split('.');
    if (p.length === 3) {
      let dTab = new Date(p[2], p[1]-1, p[0]);
      if (dTab < hranica) {
        // Ak sa archivácia nepodarí, riadok nemažeme - skúsi sa to znova pri ďalšom behu.
        if (archiveReservation(data[i], 'Prebehla', '')) {
          sheet.deleteRow(i + 1);
        }
      }
    }
  }

  invalidatePublicCache_();

  try {
    anonymizeOldArchive();
  } catch (err) {
    console.error('Anonymizácia archívu zlyhala:', err);
  }
}

// --- DIGITÁLNY ASISTENT: PRIPOMIENKY ---
// --- DIGITÁLNY ASISTENT: PRIPOMIENKY ---
function checkUpcomingConflicts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRez = ss.getSheetByName(SHEET_NAME);
    const sheetCustom = ss.getSheetByName(SHEET_CUSTOM);
    const sheetDefault = ss.getSheetByName(SHEET_DEFAULT); 
    if (!sheetRez || !sheetCustom || !sheetDefault) return;

    const now = new Date();
    const currentTimeMs = now.getTime();
    const currentHour = now.getHours();
    const currentMinOfHour = now.getMinutes();

    // --- LOGIKA VEČERNEJ HLIADKY ---
    const isEveningCheckWindow = (currentHour >= 20 && (currentHour < 22 || (currentHour === 22 && currentMinOfHour <= 30)));
    
    const todayYMD = Utilities.formatDate(now, "Europe/Bratislava", "yyyy-MM-dd");
    const dTomorrow = new Date(now);
    dTomorrow.setDate(dTomorrow.getDate() + 1);
    const tomorrowYMD = Utilities.formatDate(dTomorrow, "Europe/Bratislava", "yyyy-MM-dd");

    // Načítanie pravidiel (OPTIMALIZOVANÉ NA RÝCHLOSŤ)
    let customRules = [];
    const custData = sheetCustom.getDataRange().getDisplayValues();
    for (let i = 1; i < custData.length; i++) {
      let rDatum = custData[i][0]; 
      let rOd = custData[i][1]; 
      let rDo = custData[i][2]; 
      let rStav = custData[i][3] ? custData[i][3].trim().toUpperCase() : "";
      
      if (rDatum && rOd && rDo && rStav) {
        if (rStav !== 'ZAVRETÉ') { rOd = formatToGridStart(rOd); rDo = formatToGridEnd(rDo); }

        // Namiesto rozbaľovania celého rozsahu deň po dni (čo pri širokom/preklepnutom rozsahu
        // spôsobilo "Exceeded maximum execution time") len rovno overíme, či je v ňom dnešok/zajtrajšok.
        if (dateInInput(todayYMD, rDatum)) {
          customRules.push({ datum: todayYMD, odMin: parseTime(rOd), doMin: parseTime(rDo), stav: rStav });
        }
        if (dateInInput(tomorrowYMD, rDatum)) {
          customRules.push({ datum: tomorrowYMD, odMin: parseTime(rOd), doMin: parseTime(rDo), stav: rStav });
        }
      }
    }

    let defaultData = sheetDefault.getDataRange().getDisplayValues();
    let dniMapa = {'Pondelok':1,'Utorok':2,'Streda':3,'Štvrtok':4,'Piatok':5,'Sobota':6,'Nedeľa':7};
    let defaultHoursMap = {};
    for (let i = 1; i < defaultData.length; i++) {
      let denCislo = dniMapa[defaultData[i][0].trim()];
      if (denCislo) {
        let bloky = [];
        for (let j = 1; j < defaultData[i].length; j += 2) {
          if (defaultData[i][j] && defaultData[i][j+1]) bloky.push({ odMin: parseTime(formatToGridStart(defaultData[i][j])), doMin: parseTime(formatToGridEnd(defaultData[i][j+1])) });
        }
        defaultHoursMap[denCislo] = bloky;
      }
    }

    const rezData = sheetRez.getDataRange().getDisplayValues();
    for (let i = 1; i < rezData.length; i++) {
      let p = String(rezData[i][0]).replace(/\s/g, '').split('.');
      if (p.length !== 3) continue;
      
      let rezYMD = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
      
      // Preskočíme všetky rezervácie, ktoré nie sú dnes ani zajtra
      if (rezYMD !== todayYMD && rezYMD !== tomorrowYMD) continue;

      let casString = rezData[i][1].replace(/^'/, '');
      let [hod, min] = casString.split(':').map(Number);
      let rezDatumObj = new Date(p[2], p[1]-1, p[0], hod, min, 0);
      let rezTimeMs = rezDatumObj.getTime();
      let bStart = parseTime(casString);

      let timeDiff = (rezTimeMs - currentTimeMs) / (1000 * 60);
      let uzPoslane = rezData[i][8]; 

      let isWithinStandardWindow = (timeDiff > 30 && timeDiff <= 250);
      let isEarlyMorningTomorrow = (isEveningCheckWindow && rezYMD === tomorrowYMD && bStart < 480); 

      if ((isWithinStandardWindow || isEarlyMorningTomorrow) && !uzPoslane) {
        let meno = rezData[i][3];
        let platneSloty = getValidSlotsForDay(rezYMD, customRules, defaultHoursMap);

        if (!platneSloty.includes(bStart)) {
          let resId = rezData[i][7]; 
          let baseUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
          if (!baseUrl) baseUrl = ScriptApp.getService().getUrl();

          let dovodText = isEarlyMorningTomorrow ? 
              "je naplánovaný na zajtrajšie ráno, ale tvoj harmonogram vyzerá ináč" : 
              "sa už nenachádza v tvojom harmonograme";

          const bodyHtml = generateReminderEmailHtml(meno, casString, Math.round(timeDiff), dovodText, resId, baseUrl);
          
          MailApp.sendEmail({ name: MAIL_NAZOV, 
            to: EMAIL_BARBAR, 
            subject: isEarlyMorningTomorrow ? `🌙 VEČERNÁ HLIADKA: Ranný konflikt!` : `⏰ BUDÍČEK: Zblúdilý bojovník na ceste!`, 
            htmlBody: bodyHtml 
          });
          
          sheetRez.getRange(i + 1, 9).setValue("POSLANÉ");
        }
      }
    }
  } catch (error) {
    // Ak Google server padne, potichu to zalogujeme a skript nezlyhá
    console.error("Zachytený dočasný výpadok servera: " + error.toString());
  }
}

// --- ŠABLÓNY PRE HTML E-MAILY ---
function generateCancellationEmailHtml(meno, datum, cas, dovod, preKlienta, klientZrusilSam = false) {
    // Táto šablóna je poskladaná ako obyčajný JS reťazec (nie cez HtmlService <?= ?>), takže
    // na rozdiel od EmailSablona.html sa tu nič neescapuje automaticky - musíme to spraviť ručne,
    // inak by meno/dôvod zadaný zákazníkom mohol vložiť vlastné HTML do mailu (aj do toho pre majiteľa).
    meno = escapeHtml(meno);
    datum = escapeHtml(datum);
    cas = escapeHtml(cas);
    dovod = escapeHtml(dovod);

    let uvodnyText = "";
    if (preKlienta && !klientZrusilSam) { 
      uvodnyText = `Bohužiaľ, z prevádzkových dôvodov musím tvoj termín zrušiť. Prosím, vyber si nový termín na webe!`; 
    } else if (preKlienta && klientZrusilSam) { 
      uvodnyText = `Tvoj termín bol úspešne zrušený.`; 
    } else { 
      uvodnyText = `Bojovník <strong>${meno}</strong> práve zrušil svoj termín.`; 
    }

    let dovodHtml = dovod && dovod.trim() !== "" ? `<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eaeaea; margin-top: 15px;"><tr><td style="font-size: 12px; text-transform: uppercase; color: #888888; font-weight: bold; padding-top: 15px; padding-bottom: 4px;">Dôvod zrušenia</td></tr><tr><td style="font-size: 16px; color: #d32f2f; font-weight: bold;">${dovod}</td></tr></table>` : "";

    return `<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"><style>body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; } table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; } img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; } table { border-collapse: collapse !important; } body { height: 100% !important; margin: 0 !important; padding: 0 !important; width: 100% !important; background-color: #f4f4f5; font-family: Arial, sans-serif; }</style></head><body><table border="0" cellpadding="0" cellspacing="0" width="100%" style="padding: 20px 0;"><tr><td align="center"><table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-collapse: separate;"><tr><td align="center" style="background-color: #1a1a1a; padding: 30px; border-bottom: 4px solid #f0c419;"><img src="https://raw.githubusercontent.com/Neonka-Svk/barbarshop/refs/heads/main/barbar_logo_small.png" alt="Barbar Shop" width="150" style="display: block; max-width: 150px;"></td></tr><tr><td style="padding: 40px; color: #333333; line-height: 1.6; font-size: 16px;"><h1 style="color: #d32f2f; font-size: 24px; font-weight: bold; margin: 0 0 20px 0; text-align: center; font-family: Arial, sans-serif;">Zrušenie termínu</h1><p style="margin: 0 0 15px 0;">Zdravím ťa, <strong style="color: #1a1a1a;">${preKlienta ? meno : 'Roman'}</strong>,</p><p style="margin: 0 0 25px 0;">${uvodnyText}</p><table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fafafa; border: 1px solid #eaeaea; border-radius: 8px; border-collapse: separate;"><tr><td style="padding: 20px;"><table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 0;"><tr><td style="font-size: 12px; text-transform: uppercase; color: #888888; font-weight: bold; padding-bottom: 4px;">Dátum a čas zrušeného termínu</td></tr><tr><td style="font-size: 18px; color: #1a1a1a; font-weight: bold;">${datum} o ${cas}</td></tr></table>${dovodHtml}</td></tr></table></td></tr><tr><td align="center" style="background-color: #f9f9f9; padding: 25px; border-top: 1px solid #eaeaea; font-size: 12px; color: #999999; line-height: 1.5;"><p style="margin: 0 0 10px 0;">Tento e-mail bol vygenerovaný automaticky systémom Barbar Shop.<br>Prosíme, neodpovedajte naň.</p><p style="margin: 0;">&copy; 2026 Barbar Shop. Sila a česť.</p></td></tr></table></td></tr></table></body></html>`;
}

function generateReminderEmailHtml(meno, cas, timeDiff, dovodText, id, baseUrl) {
    // Rovnako ako pri generateCancellationEmailHtml - toto ide priamo do HTML reťazca bez auto-escapovania.
    meno = escapeHtml(meno);
    cas = escapeHtml(cas);

    // Vytvoríme odkaz na zrušovaciu stránku pre holiča
    const cancelUrl = baseUrl + "?action=cancelPage&id=" + id + "&role=holic";

    // --- LOGIKA PREFORMÁTOVANIA ČASU ---
    let casovyUdaj = "";
    if (timeDiff < 60) {
        casovyUdaj = timeDiff + " minút";
    } else {
        let hodiny = Math.floor(timeDiff / 60);
        let minuty = Math.round(timeDiff % 60);
        
        if (minuty === 0) {
            casovyUdaj = hodiny + " hod.";
        } else {
            casovyUdaj = hodiny + " hod. a " + minuty + " min.";
        }
    }
    // ------------------------------------

    return `<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"><style>body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; } table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; } img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; } table { border-collapse: collapse !important; } body { height: 100% !important; margin: 0 !important; padding: 0 !important; width: 100% !important; background-color: #f4f4f5; font-family: Arial, sans-serif; }</style></head><body><table border="0" cellpadding="0" cellspacing="0" width="100%" style="padding: 20px 0;"><tr><td align="center"><table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-collapse: separate;"><tr><td align="center" style="background-color: #1a1a1a; padding: 30px; border-bottom: 4px solid #f0c419;"><img src="https://raw.githubusercontent.com/Neonka-Svk/barbarshop/refs/heads/main/barbar_logo_small.png" alt="Barbar Shop" width="150" style="display: block; max-width: 150px;"></td></tr><tr><td style="padding: 40px; color: #333333; line-height: 1.6; font-size: 16px;"><h1 style="color: #f0c419; font-size: 24px; font-weight: bold; margin: 0 0 20px 0; text-align: center; font-family: Arial, sans-serif;">⏰ BUDÍČEK</h1><p style="margin: 0 0 15px 0;">Zdar,</p><p style="margin: 0 0 25px 0;">Len ti pripomínam, že o cca <strong>${casovyUdaj}</strong> (čas: ${cas}) ťa čaká bojovník: <strong style="color: #1a1a1a;">${meno}</strong>.</p><table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fafafa; border: 1px solid #eaeaea; border-radius: 8px; border-collapse: separate;"><tr><td style="padding: 20px;"><table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 0;"><tr><td style="font-size: 12px; text-transform: uppercase; color: #888888; font-weight: bold; padding-bottom: 4px;">Hlásenie systému</td></tr><tr><td style="font-size: 16px; color: #d32f2f; font-weight: bold;">Tento termín ${dovodText}.</td></tr></table><p style="margin: 15px 0 0 0; font-size: 14px; color: #666;">Ak chceš túto rezerváciu vykonať aj napriek upozorneniu, tento e-mail jednoducho ignoruj. Ináč možeš túto rezerváciu zrušiť kliknutím na tlačidlo nižšie.</p></td></tr></table>
    
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top: 30px;">
        <tr>
            <td align="center">
                <table border="0" cellpadding="0" cellspacing="0" style="background-color: #d32f2f; border-radius: 6px;">
                    <tr>
                        <td align="center">
                            <a href="${cancelUrl}" target="_blank" style="font-size: 15px; font-family: Arial, sans-serif; color: #ffffff; text-decoration: none; border-radius: 6px; padding: 14px 28px; display: inline-block; font-weight: bold; text-transform: uppercase;">Zrušiť termín</a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    </td></tr><tr><td align="center" style="background-color: #f9f9f9; padding: 25px; border-top: 1px solid #eaeaea; font-size: 12px; color: #999999; line-height: 1.5;"><p style="margin: 0 0 10px 0;">Tento e-mail bol vygenerovaný automaticky tvojím Digitálnym Asistentom.</p><p style="margin: 0;">&copy; 2026 Barbar Shop. Sila a česť.</p></td></tr></table></td></tr></table></body></html>`;
}
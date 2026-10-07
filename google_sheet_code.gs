// Poznatky:
// - ak sa nasadzuje upravena verzia kodu, vlozit novy kod, ulozit (ctrl+s), staci kliknut na manage deployments, 
//   tam kliknut na ceruzku, v prvom dropdowne kliknut na new version a nasledne pridat popis (volitelne) a dat save/deploy

// TODO (resp. napady na rozsirenie):
// - moznost zmenit termin (zakaznik) (FUH, MAXIMALNE SA MI ZATIAL NECHCE)

const SHEET_NAME = 'Rezervácie';
const SHEET_CUSTOM = 'Otv. hodiny mimo bežné';
const SHEET_DEFAULT = 'Bežné otv. hodiny';
const MOJ_EMAIL = 'matusjacko1@gmail.com';
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

function updateWebAppUrl() {
  // Volá sa na začiatku doPost/doGet MIMO ich try/catch - keby tu niečo hodilo chybu (napr. dočasný
  // výpadok PropertiesService), celý request by spadol a namiesto JSON-u by prišla HTML chybová stránka.
  // Preto si chybu ticho zalogujeme a ideme ďalej - toto je len pomocná "housekeeping" funkcia,
  // nesmie zablokovať skutočnú odpoveď pre zákazníka.
  try {
    const url = ScriptApp.getService().getUrl();
    // Uložíme to len ak adresa vyzerá ako ostré nasadenie (obsahuje /exec)
    if (url && url.indexOf('/exec') !== -1) {
      PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', url);
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

    // Zámok proti súbežnému zápisu - bez neho by mohli dvaja zákazníci naraz
    // zarezervovať presne ten istý termín (obaja by prešli kontrolou obsadenosti skôr,
    // než by ktorýkoľvek z nich stihol zapísať svoj riadok).
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      return jsonOut({error: 'Server je momentálne vyťažený, skús to prosím o chvíľu znova.'});
    }

    let zapisany = false;
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      const existujuce = sheet.getDataRange().getDisplayValues();

      for (let i = 1; i < existujuce.length; i++) {
        const existCas = String(existujuce[i][1]).replace(/^'/, '').trim();
        if (existujuce[i][0] === slovakDate && existCas === data.cas) {
          return jsonOut({error: 'Tento termín si medzitým niekto zarezervoval. Vyber si prosím iný.'});
        }
      }

      sheet.appendRow([slovakDate, "'" + data.cas, new Date(), safeCell(meno), sluzba, safeCell(email), false, id]);

      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 7).insertCheckboxes();

      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
      }
      zapisany = true;
    } finally {
      lock.releaseLock();
    }

    if (!zapisany) {
      return jsonOut({error: 'Rezerváciu sa nepodarilo uložiť.'});
    }

    // --- E-MAILY (mimo zámku, nech ho nedržíme počas pomalého MailApp volania) ---
    const baseUrl = ScriptApp.getService().getUrl();
    const viewUrl = `${baseUrl}?action=viewEmail&id=${id}`;
    const cancelUrlKlient = `${baseUrl}?action=cancelPage&id=${id}&role=klient`;
    const cancelUrlHolic = `${baseUrl}?action=cancelPage&id=${id}&role=holic`;

    // --- E-MAIL PRE KLIENTA ---
    const sablonaKlient = HtmlService.createTemplateFromFile('EmailSablona');
    sablonaKlient.titulok = "Rezervácia potvrdená";
    sablonaKlient.meno = meno;
    sablonaKlient.uvodnyText = "Tvoja rezervácia v Barbar Shope bola úspešne prijatá. Tešíme sa na teba, bojovník!";
    sablonaKlient.sluzba = sluzba;
    sablonaKlient.datum = slovakDate;
    sablonaKlient.cas = data.cas;
    sablonaKlient.cancelUrl = cancelUrlKlient;
    sablonaKlient.viewUrl = viewUrl;
    sablonaKlient.vyzvaText = "V prípade, že si o tento termín nežiadal alebo ho chceš zrušiť, klikni na tlačidlo nižšie:";
    sablonaKlient.jeHolic = false;

    MailApp.sendEmail({ name: MAIL_NAZOV,
      to: email,
      subject: `Barbar Shop - Potvrdenie rezervácie (${slovakDate} o ${data.cas})`,
      htmlBody: sablonaKlient.evaluate().getContent()
    });

    // --- E-MAIL PRE HOLIČA ---
    const sablonaHolic = HtmlService.createTemplateFromFile('EmailSablona');
    sablonaHolic.titulok = "Nová rezervácia!";
    sablonaHolic.meno = "Roman";
    sablonaHolic.uvodnyText = "Máš nového bojovníka v kresle! Tu sú detaily:";
    sablonaHolic.sluzba = sluzba;
    sablonaHolic.datum = slovakDate;
    sablonaHolic.cas = data.cas;
    sablonaHolic.emailKlienta = email;
    sablonaHolic.cancelUrl = cancelUrlHolic;
    sablonaHolic.viewUrl = viewUrl;
    sablonaHolic.vyzvaText = "Ak potrebuješ tento termín zrušiť z prevádzkových alebo iných dôvodov, klikni nižšie na tlačidlo:";
    sablonaHolic.jeHolic = true;

    MailApp.sendEmail({ name: MAIL_NAZOV,
      to: MOJ_EMAIL,
      subject: `Nová rezervácia: ${meno} (${slovakDate})`,
      htmlBody: sablonaHolic.evaluate().getContent()
    });

    return jsonOut({status: 'success'});
  } catch (err) {
    return jsonOut({error: err.toString()});
  }
}

function doGet(e) {
  updateWebAppUrl();
  const action = (e && e.parameter) ? e.parameter.action : null;
  const id = (e && e.parameter) ? e.parameter.id : null;
  const role = (e && e.parameter) ? e.parameter.role : 'klient';
  
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
        return template.evaluate().setTitle("Zrušenie rezervácie").addMetaTag('viewport', 'width=device-width, initial-scale=1');
      }
    }
    return HtmlService.createHtmlOutput(htmlStart + "Rezervácia nebola nájdená!" + htmlEnd);
  }

  // --- AKCIA 2: FYZICKÉ ZMAZANIE A ODOSLANIE HTML E-MAILU ---
  if (action === 'executeCancel' && id) {
    const dovodInput = (e && e.parameter.dovod) ? e.parameter.dovod.trim() : "";
    const data = sheetRez.getDataRange().getDisplayValues();
    let termiZruseny = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === id) { 
        let email = data[i][5];
        let meno = data[i][3];
        let dStr = data[i][0];
        let cas = data[i][1].replace(/^'/, '');
        
        archiveReservation(data[i], role === 'holic' ? 'Zrušil holič' : 'Zrušil zákazník', dovodInput);
        sheetRez.deleteRow(i + 1);
        termiZruseny = true;
        
        try {
          if (role === 'holic') {
            const subject = "⚠️ Zrušenie rezervácie - Barbar Shop";
            const bodyHtml = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, true);
            MailApp.sendEmail({ name: MAIL_NAZOV, to: email, subject: subject, htmlBody: bodyHtml });
          } else {
            const subjectHolic = `❌ Zrušený termín: ${meno}`;
            const bodyHtmlHolic = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, false);
            MailApp.sendEmail({ name: MAIL_NAZOV, to: MOJ_EMAIL, subject: subjectHolic, htmlBody: bodyHtmlHolic });

            const subjectKlient = "Potvrdenie zrušenia termínu";
            const bodyHtmlKlient = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, true, true);
            MailApp.sendEmail({ name: MAIL_NAZOV, to: email, subject: subjectKlient, htmlBody: bodyHtmlKlient });
          }
        } catch(err) {
          console.error('Zlyhalo odoslanie e-mailu pri zrušení (executeCancel):', err);
        }
        break;
      }
    }

    const cardTitle = termiZruseny ? "Termín bol zrušený" : "Termín nenájdený";
    const cardText = termiZruseny ? "O zrušení termínu bol odoslaný potvrdzujúci e-mail. Môžeš zatvoriť túto stránku." : "Tento termín už neexistuje alebo bol zrušený v minulosti.";
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

  // --- AKCIA 4: JSON DÁTA PRE KALENDÁR ---
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
    console.error('Načítanie recenzií zlyhalo (kalendár ide ďalej):', err);
  }

  return ContentService.createTextOutput(JSON.stringify({ bookings, custom: customRanges, defaultHours, reviews }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- LOGIKA PRE ZRUŠENIE A UPOZORNENIA PRIAMO V TABUĽKE ---
function handleEdit(e) {
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

  // 1b. Zmazanie recenzie cez CHECKBOX v hárku "Recenzie" (posledný stĺpec)
  if (sheetName === SHEET_RECENZIE && e.range.getColumn() === RECENZIE_HLAVICKA.length && e.range.getRow() > 1 && e.range.getValue() === true) {
    const ui = SpreadsheetApp.getUi();
    const r = sheet.getRange(e.range.getRow(), 1, 1, RECENZIE_HLAVICKA.length).getValues()[0];
    const odpoved = ui.alert(
      'Zmazať recenziu?',
      'Recenzia od "' + r[4] + '" (' + r[2] + '/5) sa nenávratne odstráni z tabuľky aj z webu. Pokračovať?',
      ui.ButtonSet.YES_NO
    );
    if (odpoved === ui.Button.YES) {
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
}

// ===================== RECENZIE =====================
// Tok: deň po strihu príde zákazníkovi mail s odkazom (?action=reviewPage&id=<ID rezervácie>) ->
// stránka RecenziaPage.html -> submitReview() zapíše riadok do hárku "Recenzie" -> holič ho
// odklikne (stĺpec "Schválené") -> riadok sa objaví na webe (?action=reviews).
// Jedna rezervácia = najviac jedna recenzia; rezervácia sa hľadá v "Rezervácie" aj v "Archív".
const SHEET_RECENZIE = 'Recenzie';
const RECENZIE_HLAVICKA = ['Identifikátor rezervácie', 'Dátum strihania', 'Hodnotenie (1-5)', 'Recenzia', 'Zobrazené meno', 'Odoslané (súhlas so zverejnením)', 'Schválené', 'Odpoveď holiča', 'Zmazať recenziu?'];
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

// Starší hárok "Recenzie" (ešte bez stĺpca "Zmazať recenziu?") sa sám doplní o hlavičku a checkboxy.
function upgradeReviewSheet_(sheet) {
  const stlpec = RECENZIE_HLAVICKA.length;
  if (sheet.getRange(1, stlpec).getValue()) return;
  sheet.getRange(1, stlpec).setValue(RECENZIE_HLAVICKA[stlpec - 1]);
  sheet.getRange(1, stlpec - 1).copyFormatToRange(sheet, stlpec, stlpec, 1, 1);
  if (sheet.getLastRow() > 1) sheet.getRange(2, stlpec, sheet.getLastRow() - 1, 1).insertCheckboxes();
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

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return { ok: false, error: 'Server je vyťažený, skús to prosím o chvíľu znova.' };

    try {
      const rez = findReservationForReview_(id);
      if (!rez) return { ok: false, error: REVIEW_CHYBY.nenajdene };
      if (rez.chyba) return { ok: false, error: REVIEW_CHYBY[rez.chyba] };
      if (reviewExists_(id)) return { ok: false, error: REVIEW_CHYBY.existuje };

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
        false
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
      sheet.getRange(cielovyRiadok, 7).insertCheckboxes();
      sheet.getRange(cielovyRiadok, 9).insertCheckboxes();
    } finally {
      lock.releaseLock();
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const baseUrl = ScriptApp.getService().getUrl();
      const approveUrl = baseUrl + '?action=approveReview&id=' + id + '&t=' + reviewToken_(id, 'approve');
      const deleteUrl = baseUrl + '?action=deleteReview&id=' + id + '&t=' + reviewToken_(id, 'delete');
      const sheetUrl = ss.getUrl() + '#gid=' + getReviewSheet_().getSheetId();
      MailApp.sendEmail({ name: MAIL_NAZOV,
        to: MOJ_EMAIL,
        subject: 'Nová recenzia čaká na schválenie (' + stars + '/5)',
        htmlBody: generateNewReviewEmailHtml_(name, stars, text, approveUrl, deleteUrl, sheetUrl)
      });
    } catch (err) {
      console.error('Upozornenie na novú recenziu sa neodoslalo:', err);
    }
    return { ok: true };
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

// Vráti {row, id, stars, text, name, approved, reply} alebo null.
function findReview_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECENZIE);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECENZIE_HLAVICKA.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      return { row: i + 2, id: id, stars: Number(rows[i][2]), text: String(rows[i][3]), name: String(rows[i][4]), approved: rows[i][6] === true, reply: String(rows[i][7] || '') };
    }
  }
  return null;
}

// Jednoduchá stránka pre holiča (schválenie recenzie). Všetko dynamické sa escapuje.
function reviewAdminPage_(title, innerHtml) {
  const html = '<!DOCTYPE html><html lang="sk"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Recenzia | Barbar Shop</title>' +
    '<style>' + fontFaceCss() +
    'body{font-family:Roboto,"Segoe UI",Helvetica,Arial,sans-serif;background:#1a1a1a;color:#e0e0e0;margin:0;padding:20px;display:flex;justify-content:center;align-items:center;min-height:100vh;box-sizing:border-box}' +
    '.card{background:#2b1d16;border:1px solid #f0c419;border-radius:12px;padding:32px;max-width:480px;width:100%;box-shadow:0 10px 30px rgba(0,0,0,.5);box-sizing:border-box}' +
    'h2{color:#f0c419;margin:0 0 18px;text-align:center}.stars{color:#f0c419;font-size:28px;letter-spacing:3px;text-align:center}.stars .off{color:#555}' +
    '.quote{background:#1a1a1a;border-left:4px solid #f0c419;border-radius:6px;padding:12px 15px;margin:16px 0;white-space:pre-line;overflow-wrap:anywhere}' +
    '.who{color:#aaa;font-size:14px;text-align:center}label{display:block;color:#f0c419;font-weight:bold;font-size:14px;margin:18px 0 6px}' +
    'textarea{width:100%;box-sizing:border-box;min-height:80px;background:#1a1a1a;border:1px solid #444;color:#fff;padding:12px;border-radius:6px;font-family:inherit;font-size:16px}' +
    '.btn{display:block;width:100%;box-sizing:border-box;text-align:center;text-decoration:none;background:#2e7d32;color:#fff;border:none;padding:15px;font-size:16px;font-weight:bold;border-radius:6px;cursor:pointer;margin-top:18px}.btn:hover{background:#256528}' +
    '.link{display:block;text-align:center;color:#f0c419;margin-top:16px;font-size:14px}.ok{font-size:60px;color:#4caf50;text-align:center}p{line-height:1.5}' +
    '</style></head><body><div class="card"><h2>' + escapeHtml(title) + '</h2>' + innerHtml + '</div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Recenzia | Barbar Shop').addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
    '<div class="who">Zobrazí sa ako: <strong>' + escapeHtml(r.name) + '</strong></div>' +
    '<form method="GET" action="' + escapeHtml(ScriptApp.getService().getUrl()) + '" target="_top">' +
    '<input type="hidden" name="action" value="doApproveReview"><input type="hidden" name="id" value="' + escapeHtml(id) + '"><input type="hidden" name="t" value="' + escapeHtml(token) + '">' +
    '<label for="odpoved">Tvoja odpoveď pod recenziou (nepovinné)</label>' +
    '<textarea id="odpoved" name="odpoved" maxlength="500" placeholder="Napr. Ďakujeme, tešíme sa na ďalšiu návštevu!"></textarea>' +
    '<button type="submit" class="btn">SCHVÁLIŤ A ZVEREJNIŤ</button></form>' +
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
    sheet.getRange(r.row, 7).setValue(true);
    const reply = String(odpoved || '').trim().slice(0, 500);
    if (reply) sheet.getRange(r.row, 8).setValue(safeCell(reply));
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
    '<div class="who">Od: <strong>' + escapeHtml(r.name) + '</strong>' + (r.approved ? ' (zatiaľ zverejnená na webe)' : '') + '</div>' +
    '<p style="text-align:center;color:#ff8a80">Recenzia sa nenávratne odstráni z tabuľky aj z webu.</p>' +
    '<form method="GET" action="' + escapeHtml(ScriptApp.getService().getUrl()) + '" target="_top">' +
    '<input type="hidden" name="action" value="doDeleteReview"><input type="hidden" name="id" value="' + escapeHtml(id) + '"><input type="hidden" name="t" value="' + escapeHtml(token) + '">' +
    '<button type="submit" class="btn" style="background:#d32f2f">ZMAZAŤ NAVŽDY</button></form>');
}

// Krok 2: skutočné zmazanie riadku.
function doDeleteReview_(id, token) {
  if (!id || token !== reviewToken_(id, 'delete')) return reviewAdminPage_('Neplatný odkaz', '<p>Tento odkaz na zmazanie nie je platný.</p>');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return reviewAdminPage_('Skús znova', '<p>Server je práve vyťažený, skús to o chvíľu.</p>');
  try {
    const r = findReview_(id);
    if (!r) return reviewAdminPage_('Recenzia sa nenašla', '<p>Recenzia už neexistuje.</p>');
    getReviewSheet_().deleteRow(r.row);
  } finally {
    lock.releaseLock();
  }
  return reviewAdminPage_('Zmazané', '<div class="ok">&#10003;</div><p style="text-align:center">Recenzia bola odstránená. Môžeš zatvoriť túto stránku.</p>');
}

function generateNewReviewEmailHtml_(name, stars, text, approveUrl, deleteUrl, sheetUrl) {
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
    return { stars: Number(r[2]), text: String(r[3]), name: String(r[4]), date: formatReviewDate_(r[1]), reply: String(r[7] || '') };
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
      if (MailApp.getRemainingDailyQuota() < 20) break;

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

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ Správa rezervácií')
    .addItem('Vymazať staré rezervácie (7+ dni)', 'cleanupOldReservations')
    .addToUi();

  // Staršiemu hárku "Recenzie" doplní stĺpec "Zmazať recenziu?" (hlavička + checkboxy).
  try {
    const recenzie = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECENZIE);
    if (recenzie) upgradeReviewSheet_(recenzie);
  } catch (err) {
    console.error('Úprava hárku Recenzie zlyhala:', err);
  }
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

  data.forEach((r, i) => {
    if (r[5] === ARCHIV_ANONYMIZOVANE) return;
    const p = String(r[0]).replace(/\s/g, '').split('.');
    if (p.length !== 3) return;
    if (new Date(p[2], p[1] - 1, p[0]) < hranica) {
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
            to: MOJ_EMAIL, 
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
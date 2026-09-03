// Poznatky:
// - ak sa nasadzuje upravena verzia kodu, vlozit novy kod, ulozit (ctrl+s), staci kliknut na manage deployments, 
//   tam kliknut na ceruzku, v prvom dropdowne kliknut na new version a nasledne pridat popis (volitelne) a dat save/deploy

// TODO (resp. napady na rozsirenie):
// - moznost zmenit termin (zakaznik) (FUH, MAXIMALNE SA MI ZATIAL NECHCE)

const SHEET_NAME = 'Rezervácie';
const SHEET_CUSTOM = 'Otv. hodiny mimo bežné';
const SHEET_DEFAULT = 'Bežné otv. hodiny';
const MOJ_EMAIL = 'matusjacko1@gmail.com';

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
          while (current <= end) {
            results.push(`${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, '0')}-${String(current.getDate()).padStart(2, '0')}`);
            current.setDate(current.getDate() + 1);
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

      sheet.appendRow([slovakDate, "'" + data.cas, new Date(), meno, sluzba, email, false, id]);

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

    MailApp.sendEmail({
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

    MailApp.sendEmail({
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
  
  const htmlStart = "<div style='font-family:Montserrat, sans-serif; background:#1a1a1a; color:#f0c419; height:100vh; display:flex; justify-content:center; align-items:center; text-align:center; margin:0;'><h2>";
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
        
        sheetRez.deleteRow(i + 1);
        termiZruseny = true;
        
        try {
          if (role === 'holic') {
            const subject = "⚠️ Zrušenie rezervácie - Barbar Shop";
            const bodyHtml = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, true);
            MailApp.sendEmail({ to: email, subject: subject, htmlBody: bodyHtml });
          } else {
            const subjectHolic = `❌ Zrušený termín: ${meno}`;
            const bodyHtmlHolic = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, false);
            MailApp.sendEmail({ to: MOJ_EMAIL, subject: subjectHolic, htmlBody: bodyHtmlHolic });

            const subjectKlient = "Potvrdenie zrušenia termínu";
            const bodyHtmlKlient = generateCancellationEmailHtml(meno, dStr, cas, dovodInput, true, true);
            MailApp.sendEmail({ to: email, subject: subjectKlient, htmlBody: bodyHtmlKlient });
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
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Montserrat', sans-serif; background-color: #1a1a1a; color: #e0e0e0; margin: 0; padding: 20px; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
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
  
  return ContentService.createTextOutput(JSON.stringify({ bookings, custom: customRanges, defaultHours }))
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
       let data = sheet.getRange(row, 1, 1, 6).getDisplayValues()[0]; 
       
       try { 
         const bodyHtml = generateCancellationEmailHtml(data[3], data[0], data[1].replace(/^'/, ''), dovod, true);
         MailApp.sendEmail({ 
           to: data[5], 
           subject: "⚠️ Zrušenie rezervácie - Barbar Shop", 
           htmlBody: bodyHtml
         });
       } catch(err) {
         console.error('Zlyhalo odoslanie e-mailu pri zrušení (checkbox):', err);
       }

       sheet.deleteRow(row);
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
      let html = template.evaluate().setWidth(550).setHeight(450);
      SpreadsheetApp.getUi().showModalDialog(html, 'Konflikt pri rezerváciách a časoch!');
    }
  }
}

// Hromadné rušenie z Popup Okna
function processCancellations(cancellations) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  cancellations.sort((a,b) => b.row - a.row);
  
  cancellations.forEach(c => {
    try {
      const bodyHtml = generateCancellationEmailHtml(c.meno, c.datum, c.cas.replace(/^'/, ''), c.dovod, true);
      MailApp.sendEmail({ 
        to: c.email, 
        subject: "⚠️ Zrušenie rezervácie - Barbar Shop", 
        htmlBody: bodyHtml 
      });
      sheet.deleteRow(c.row);
    } catch(err) { 
      console.error("Nepodarilo sa zrušiť riadok:", c.row); 
    }
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ Správa rezervácií')
    .addItem('Vymazať staré rezervácie (7+ dni)', 'cleanupOldReservations')
    .addToUi();
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
        sheet.deleteRow(i + 1);
      }
    }
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
        
        let expanded = expandDates(rDatum);
        expanded.forEach(d => { 
          // Ukladáme IBA pravidlá pre dnešok a zajtrajšok! Zvyšok ignorujeme.
          if (d === todayYMD || d === tomorrowYMD) {
            customRules.push({ datum: d, odMin: parseTime(rOd), doMin: parseTime(rDo), stav: rStav }); 
          }
        });
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
          
          MailApp.sendEmail({ 
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
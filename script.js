// TODO: dalo by sa to refaktornut ngl xdd

// Nasadenie pod účtom roman.barbarshop (web aplikácia beží a maily posiela pod tým, kto ju nasadil).
const PROD_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw65KS-ksdoc0ep2pt3YdrTCZY39zaN0EuXW3rzIHx2gAEyvvy4kyRd8SFHSESTISJZmw/exec';

// Lokálne testovanie proti testovacej kópii tabuľky: http://localhost:8124/?api=<URL testovacieho nasadenia>
// (adresa sa zapamätá, ?api=prod vráti ostrú). Funguje LEN na localhoste, na ostrom webe sa dá zmeniť jedine tu v kóde,
// inak by stačil podstrčený odkaz s ?api= a objednávky by odchádzali na cudzí server.
const SCRIPT_URL = (() => {
    if (!['localhost', '127.0.0.1'].includes(location.hostname)) return PROD_SCRIPT_URL;
    try {
        const api = new URLSearchParams(location.search).get('api');
        if (api === 'prod') localStorage.removeItem('barbarTestApi');
        else if (api && /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(api)) localStorage.setItem('barbarTestApi', api);
        const test = localStorage.getItem('barbarTestApi');
        if (test) {
            document.title = '[TEST] ' + document.title;
            console.warn('Web beží proti TESTOVACIEMU skriptu:', test);
            return test;
        }
    } catch (e) { /* bez localStorage ostáva ostrá adresa */ }
    return PROD_SCRIPT_URL;
})();
let currentWeekOffset = 0;
let globalData = { bookings: [], custom: [], defaultHours: [] };

function getLocalDateString(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// Deň v týždni (1=Pondelok...7=Nedeľa) čisto z čísel dátumu, bez new Date(ymd).getDay() -
// ten je citlivý na časové pásmo prehliadača a vie tesne okolo polnoci vrátiť iný deň, než bol zadaný.
// Rovnaká logika beží aj na serveri (google_sheet_code.gs), nech obe strany počítajú identicky.
function getIsoDayOfWeek(ymd) {
    const parts = String(ymd).split('-').map(Number);
    let y = parts[0], m = parts[1], d = parts[2];
    const t = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
    if (m < 3) y -= 1;
    const dow = (y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) + t[m - 1] + d) % 7;
    return dow === 0 ? 7 : dow;
}

function pridaj90Minut(cas) {
    let [h, m] = cas.split(':').map(Number);
    let d = new Date();
    d.setHours(h, m + 90, 0);
    return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}

function casNaMinuty(cas) {
    const [h, m] = cas.split(':').map(Number);
    return h * 60 + m;
}

function minutyNaCas(min) {
    const h = Math.floor(min / 60) % 24; // Zabezpečí prechod 24 -> 00, 25 -> 01 atď.
    const m = min % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

// Prevedie {od, do} (časy v tvare "HH:MM") na rozsah v minútach {start, end}, s ošetrením prechodu cez polnoc.
function toRange(r) {
    const start = casNaMinuty(r.od);
    let end = casNaMinuty(r.do);
    if (end < start) end += 24 * 60;
    return { start, end };
}

// Odpočíta z rozsahu `range` ({start,end} v minútach) všetky rozsahy z poľa `occupied` a vráti
// zvyšné (neprekrývajúce sa) kusy. Vďaka tomu ROZŠÍRENÉ generuje sloty len z toho, čo bežné/OTVORENÉ
// hodiny ešte nepokrývajú (inak by pri prekryve s iným 90-min rastrom vznikli poprehadzované okná).
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

function getRawSlotsForDate(dateStr) {
    let rawSlots = [];
    const denneCustom = globalData.custom.filter(c => c.datum === dateStr);
    
    // Rozdelíme si stavy
    const otvoreneCustom = denneCustom.filter(c => c.stav === 'OTVORENÉ');
    const rozsireneCustom = denneCustom.filter(c => c.stav === 'ROZŠÍRENÉ');
    const zavreteCustom = denneCustom.filter(c => c.stav === 'ZAVRETÉ');

    if (otvoreneCustom.length > 0) {
        // Ak má vlastné OTVORENÉ, úplne ignorujeme bežné hodiny
        const otvoreneRanges = otvoreneCustom.map(toRange);
        rawSlots = rawSlots.concat(slotsFromRanges(otvoreneRanges));

        // ROZŠÍRENÉ pridáva čas NAVYŠE k OTVORENÉ - časť, ktorá sa s ním prekrýva, je už zarátaná
        // vyššie, takže tu spracujeme len tie kusy, čo ležia mimo OTVORENÉ rozsahu.
        rozsireneCustom.forEach(r => {
            const zvysne = subtractRanges(toRange(r), otvoreneRanges);
            rawSlots = rawSlots.concat(slotsFromRanges(zvysne));
        });
    } else {
        // Ak NEMA vlastne otvorene, pouzijeme BEŽNÉ hodiny
        let dayOfWeek = getIsoDayOfWeek(dateStr);

        const rozpisNaDnes = globalData.defaultHours.find(h => h.den === dayOfWeek);
        const defRanges = (rozpisNaDnes && rozpisNaDnes.bloky) ? rozpisNaDnes.bloky.map(toRange) : [];

        rawSlots = rawSlots.concat(slotsFromRanges(defRanges));

        // ROZŠÍRENÉ pridáva čas NAVYŠE k bežným hodinám - časť, ktorá sa s nimi prekrýva, je už
        // zarátaná vyššie, takže tu spracujeme len tie kusy rozsahu, čo ležia mimo bežných hodín.
        rozsireneCustom.forEach(r => {
            const zvysne = subtractRanges(toRange(r), defRanges);
            rawSlots = rawSlots.concat(slotsFromRanges(zvysne));
        });
    }

    // --- APLIKÁCIA "ZAVRETÉ" (Tvoja pôvodná a fungujúca logika) ---
    zavreteCustom.forEach(r => {
        const odMin = casNaMinuty(r.od);
        let doMin = casNaMinuty(r.do);
        if (doMin < odMin) doMin += 24 * 60; // Ak by pauza išla cez polnoc
        
        // Ak sa celá pauza rovná 00:00 - 24:00 (čo je celodenná pauza, 0 - 1440 min)
        if (odMin === 0 && doMin === 1440) {
           rawSlots = []; // Uplne vyprazdnime pole, nic sa nevygeneruje
        } else {
            // Ak slot začína alebo končí počas "ZAVRETÉHO" intervalu, vyhodíme ho z rawSlots
            rawSlots = rawSlots.filter(sMin => {
                const eMin = sMin + 90;
                // Je to v poriadku, iba ak končí pred prestávkou alebo začína po prestávke
                return (eMin <= odMin || sMin >= doMin); 
            });
        }
    });

    // POZOR: odstránenie duplicít (ak by sa bežné a rozšírené prekrývali)
    // a usporiadanie podľa času (aby termíny vo finále išli pekne poporade)
    let uniqueSlots = [...new Set(rawSlots)];
    return uniqueSlots.sort((a, b) => a - b);
}

function getProcessedSlotsForDate(targetDateStr) {
    let processedSlots = [];
    const dnesneBookings = globalData.bookings.filter(b => b.datum === targetDateStr);
    
    let yesterday = new Date(targetDateStr);
    yesterday.setDate(yesterday.getDate() - 1);
    let yesterdayStr = getLocalDateString(yesterday);
    const vcerajsieBookings = globalData.bookings.filter(b => b.datum === yesterdayStr);

    let rawPotentialSlots = getRawSlotsForDate(targetDateStr);
    let yesterdayRaw = getRawSlotsForDate(yesterdayStr);

    // 1. Zohľadníme OBSADENÉ včerajšie rezervácie, ktoré pretiekli cez polnoc (OPRAVA)
    vcerajsieBookings.forEach(b => {
        let bStart = casNaMinuty(b.cas);
        if (bStart + 90 > 1440) {
            processedSlots.push({
                id: `${yesterdayStr}_${b.cas}_cont`, 
                startMin: 0,
                durationMins: (bStart + 90) - 1440,
                originalDate: yesterdayStr,
                originalTime: b.cas,
                isContinuation: true,
                isForceBooked: true 
            });
        }
    });

    // 2. Dnešné OBSADENÉ rezervácie (s orezaním dĺžky na polnoc) (OPRAVA)
    dnesneBookings.forEach(b => {
        let bStart = casNaMinuty(b.cas);
        processedSlots.push({
            id: `${targetDateStr}_${b.cas}`,
            startMin: bStart,
            durationMins: Math.min(90, 1440 - bStart), // poistka, aby nepretieklo cez 24:00
            originalDate: targetDateStr,
            originalTime: b.cas,
            isContinuation: false,
            isForceBooked: true 
        });
    });

    // 3. Doplníme voľné sloty z DNEŠNEJ šmeny
    rawPotentialSlots.forEach(m => {
        if (m < 1440) {
            const endM = m + 90;
            const koliziaDnes = dnesneBookings.some(b => {
                const bStart = casNaMinuty(b.cas);
                return (m < (bStart + 90) && endM > bStart);
            });
            
            const koliziaVcera = vcerajsieBookings.some(b => {
                const bStart = casNaMinuty(b.cas);
                const bEndDnes = (bStart + 90) - 1440;
                return (bStart + 90 > 1440) && (m < bEndDnes);
            });

            if (!koliziaDnes && !koliziaVcera) {
                processedSlots.push({
                    id: `${targetDateStr}_${minutyNaCas(m)}`,
                    startMin: m,
                    durationMins: Math.min(90, 1440 - m),
                    originalDate: targetDateStr,
                    originalTime: minutyNaCas(m),
                    isContinuation: false
                });
            }
        }
    });

    // 4. Spill-over z včerajšej nočnej voľnej šmeny
    yesterdayRaw.forEach(m => {
        if (m < 1440 && (m + 90) > 1440) {
            const kolizia = vcerajsieBookings.some(b => {
                const bStart = casNaMinuty(b.cas);
                return (m < (bStart + 90) && (m + 90) > bStart);
            });
            
            if (!kolizia) {
                processedSlots.push({
                    id: `${yesterdayStr}_${minutyNaCas(m)}`,
                    startMin: 0,
                    durationMins: (m + 90) - 1440,
                    originalDate: yesterdayStr,
                    originalTime: minutyNaCas(m),
                    isContinuation: true
                });
            }
        } 
        else if (m >= 1440) {
            let startMinDnes = m - 1440;
            let endMinDnes = startMinDnes + 90;
            
            const koliziaDnes = dnesneBookings.some(b => {
                const bStart = casNaMinuty(b.cas);
                return (startMinDnes < (bStart + 90) && endMinDnes > bStart);
            });
            
            const koliziaVcera = vcerajsieBookings.some(b => {
                const bStart = casNaMinuty(b.cas);
                const bEndDnes = (bStart + 90) - 1440;
                return (bStart + 90 > 1440) && (startMinDnes < bEndDnes);
            });

            if (!koliziaDnes && !koliziaVcera && startMinDnes < 1440) {
                processedSlots.push({
                    id: `${targetDateStr}_${minutyNaCas(m)}`,
                    startMin: startMinDnes,
                    durationMins: Math.min(90, 1440 - startMinDnes),
                    originalDate: targetDateStr,
                    originalTime: minutyNaCas(m),
                    isContinuation: false
                });
            }
        }
    });

    return processedSlots.sort((a, b) => a.startMin - b.startMin);
}

// --- RECENZIE ---
// Schválené recenzie prichádzajú v tej istej odpovedi ako kalendár (globalData.reviews), takže sekcia
// sa ukáže až s hotovými dátami: žiadna kostra, ktorá by po zistení "nič tu nie je" zmizla, a žiadna
// uložená kópia, ktorá by ukazovala už zmazanú recenziu. Kým dáta nie sú, sekcia je skrytá.
try { localStorage.removeItem('barbarshop_reviews_v1'); } catch (e) { /* staré uloženie z predošlej verzie, už ho nepoužívame */ }

function hideReviews() {
    const section = document.getElementById('reviews-section');
    const list = document.getElementById('reviews-list');
    if (list) list.innerHTML = '';
    if (section) section.hidden = true;
}

function renderReviews(data) {
    const section = document.getElementById('reviews-section');
    const list = document.getElementById('reviews-list');
    const summary = document.getElementById('reviews-summary');
    if (!section) return;
    if (!data || !Array.isArray(data.reviews) || data.reviews.length === 0) {
        hideReviews();
        return;
    }

    const starsText = (n) => '★'.repeat(n) + '☆'.repeat(5 - n);
    const makeEl = (tag, className, text) => {
        const el = document.createElement(tag);
        if (className) el.className = className;
        if (text !== undefined) el.textContent = text; // textContent: text recenzie píšu zákazníci, nikdy ho nevkladáme ako HTML
        return el;
    };

    summary.textContent = `${starsText(Math.round(data.average))} ${String(data.average).replace('.', ',')} / 5 (${data.count})`;

    list.innerHTML = '';
    data.reviews.forEach(r => {
        const stars = Math.min(5, Math.max(0, Math.round(Number(r.stars) || 0)));
        const card = makeEl('div', 'review-card');
        card.appendChild(makeEl('div', 'review-stars', starsText(stars)));
        card.appendChild(makeEl('p', 'review-text', r.text));
        card.appendChild(makeEl('div', 'review-meta', `${r.name} · ${r.date}`));
        if (r.reply) {
            const reply = makeEl('div', 'review-reply');
            reply.appendChild(makeEl('strong', null, 'Odpoveď holiča: '));
            reply.appendChild(document.createTextNode(r.reply));
            card.appendChild(reply);
        }
        const photos = (Array.isArray(r.photos) ? r.photos : []).filter(isSafeDriveId).map(id => ({ id, type: 'image', name: 'Fotka k recenzii od ' + r.name }));
        if (photos.length > 0) {
            const strip = makeEl('div', 'review-photos');
            photos.forEach((p, i) => strip.appendChild(buildMediaButton(p, photos, i, 'review-photo', 300)));
            card.insertBefore(strip, card.querySelector('.review-meta'));
        }
        list.appendChild(card);
    });

    section.hidden = false;
}

// --- GALÉRIA (fotky a videá z priečinka web_photos na Google Disku) + fotky v recenziách ---
// Rovnako ako recenzie prichádza v odpovedi s kalendárom (globalData.gallery) a sekcia je skrytá, kým nie sú dáta.
// Vo vizuáli stránky nie je nikde odkaz na Disk: obrázky sa len načítavajú z Googlu a video sa prehráva
// v lightboxe, kde sú odkazy na Disk (vyskakovacie okno prehrávača) zakryté.
const galleryThumbUrl = (id, width) => `https://drive.google.com/thumbnail?id=${id}&sz=w${width}`;
const galleryVideoUrl = (id) => `https://drive.google.com/file/d/${id}/preview`;
const isSafeDriveId = (id) => /^[\w-]{10,}$/.test(String(id));

let lightboxItems = [];
let lightboxIndex = -1;
let lightboxOpener = null;

function hideGallery() {
    const section = document.getElementById('gallery-section');
    const grid = document.getElementById('gallery-grid');
    if (grid) grid.innerHTML = '';
    if (section) section.hidden = true;
}

// Tlačidlo s miniatúrou, ktoré po kliknutí otvorí lightbox na danej položke zoznamu.
function buildMediaButton(item, list, index, extraClass, thumbWidth) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = extraClass + (item.type === 'video' ? ' is-video' : '');
    btn.setAttribute('aria-label', (item.type === 'video' ? 'Prehrať video: ' : 'Zväčšiť fotku: ') + (item.name || ''));
    const img = document.createElement('img');
    img.src = galleryThumbUrl(item.id, thumbWidth);
    img.alt = item.name || '';
    img.loading = 'lazy';
    img.referrerPolicy = 'no-referrer';
    btn.appendChild(img);
    btn.addEventListener('click', () => openLightbox(list, index, btn));
    return btn;
}

function renderGallery(data) {
    const section = document.getElementById('gallery-section');
    if (!section) return;

    // ID ide do URL, preto prijmeme len bezpečné znaky; názvy sa vkladajú výlučne cez textContent / alt
    const items = (data && Array.isArray(data.items) ? data.items : [])
        .filter(i => i && isSafeDriveId(i.id) && (i.type === 'image' || i.type === 'video'));
    if (items.length === 0) {
        hideGallery();
        return;
    }

    const grid = document.getElementById('gallery-grid');
    grid.innerHTML = '';
    items.forEach((item, index) => grid.appendChild(buildMediaButton(item, items, index, 'gallery-item', 600)));
    section.hidden = false;
}

function openLightbox(list, index, opener) {
    const lb = document.getElementById('lightbox');
    const stage = document.getElementById('lightbox-stage');
    const item = list && list[index];
    if (!lb || !item) return;

    lightboxItems = list;
    lightboxIndex = index;
    if (opener) lightboxOpener = opener;
    stage.innerHTML = '';
    if (item.type === 'video') {
        const wrap = document.createElement('div');
        wrap.className = 'video-wrap';
        const frame = document.createElement('iframe');
        frame.src = galleryVideoUrl(item.id);
        frame.allow = 'autoplay; fullscreen';
        frame.allowFullscreen = true;
        frame.title = item.name || 'Video';
        const mask = document.createElement('div'); // zakryje tlačidlo prehrávača, ktoré by otvorilo Disk
        mask.className = 'video-mask';
        wrap.appendChild(frame);
        wrap.appendChild(mask);
        stage.appendChild(wrap);
    } else {
        const img = document.createElement('img');
        img.src = galleryThumbUrl(item.id, 1600);
        img.alt = item.name || '';
        img.referrerPolicy = 'no-referrer';
        stage.appendChild(img);
    }
    document.getElementById('lightbox-caption').textContent = item.name || '';
    const multiple = list.length > 1;
    document.getElementById('lightbox-prev').hidden = !multiple;
    document.getElementById('lightbox-next').hidden = !multiple;

    lb.hidden = false;
    document.body.style.overflow = 'hidden';
    document.getElementById('lightbox-close').focus();
}

function closeLightbox() {
    const lb = document.getElementById('lightbox');
    if (!lb || lb.hidden) return;
    document.getElementById('lightbox-stage').innerHTML = ''; // zastaví aj prehrávané video
    lb.hidden = true;
    document.body.style.overflow = '';
    lightboxIndex = -1;
    if (lightboxOpener && document.contains(lightboxOpener)) lightboxOpener.focus();
}

function stepLightbox(delta) {
    if (lightboxIndex < 0 || lightboxItems.length < 2) return;
    openLightbox(lightboxItems, (lightboxIndex + delta + lightboxItems.length) % lightboxItems.length);
}

function setupLightbox() {
    const lb = document.getElementById('lightbox');
    if (!lb) return;
    document.getElementById('lightbox-close').addEventListener('click', closeLightbox);
    document.getElementById('lightbox-prev').addEventListener('click', () => stepLightbox(-1));
    document.getElementById('lightbox-next').addEventListener('click', () => stepLightbox(1));
    lb.addEventListener('click', (e) => {
        if (e.target === lb || e.target.id === 'lightbox-stage') closeLightbox();
    });
    document.addEventListener('keydown', (e) => {
        if (lb.hidden) return;
        if (e.key === 'Escape') closeLightbox();
        else if (e.key === 'ArrowLeft') stepLightbox(-1);
        else if (e.key === 'ArrowRight') stepLightbox(1);
    });
    let touchX = null;
    lb.addEventListener('touchstart', (e) => { touchX = e.changedTouches[0].clientX; }, { passive: true });
    lb.addEventListener('touchend', (e) => {
        if (touchX === null) return;
        const dx = e.changedTouches[0].clientX - touchX;
        touchX = null;
        if (Math.abs(dx) > 50) stepLightbox(dx > 0 ? -1 : 1);
    }, { passive: true });
}

// Na veľkej obrazovke je formulár a kalendár vedľa seba, takže automatické posúvanie stránky netreba.
const isDesktopLayout = () => window.matchMedia('(min-width: 1000px)').matches;

document.addEventListener('DOMContentLoaded', async () => {
    setupLightbox();
    setupPending();
    resumePending(); // rozpracovaná rezervácia z tejto karty (napr. po obnovení stránky pri prepnutí do mailu)
    await refreshData();

    // Navigácia týždňov
    document.getElementById('nextWeek').addEventListener('click', () => {
        if (currentWeekOffset < 4) { currentWeekOffset++; updateUI(); }
    });
    document.getElementById('prevWeek').addEventListener('click', () => {
        if (currentWeekOffset > 0) { currentWeekOffset--; updateUI(); }
    });

    document.getElementById('btn-go-to-calendar').addEventListener('click', () => {
        if (!isDesktopLayout()) document.querySelector('.timetable-section').scrollIntoView({ behavior: 'smooth' });
    });

    document.getElementById('btn-clear-termin').addEventListener('click', () => {
        document.getElementById('date').value = '';
        document.getElementById('time').value = '';
        updateSelectedTerminUI(); // Zmení vizuál na "nevybraté"
        renderTimetable(); // Prekreslí tabuľku, aby zmizol zlatý blok
    });
    
    // Prvotné nastavenie UI pri štarte stránky
    updateSelectedTerminUI();
    // --- PRIDANÉ: Spustenie "živých hodín" každú 1 minútu (60000 ms) ---
    let lastMinute = -1;
    setInterval(() => {
        const aktualnyCas = getCETime();
        const aktualnaMinuta = aktualnyCas.getMinutes();
        
        // Ak sa minúta zmenila (presne na prelome), aktualizuj čiaru a zamkni termíny
        if (aktualnaMinuta !== lastMinute) {
            lastMinute = aktualnaMinuta;
            updateTimeLine();
        }
    }, 1000);
});

async function refreshData() {
    const container = document.querySelector('.timetable-container');
    const btnGoToCalendar = document.getElementById('btn-go-to-calendar');
    const calendarBtns = document.querySelector(".week-controls");
    let overlay = document.getElementById('loading-overlay');

    if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'loading-overlay';
        overlay.innerHTML = '<span>Aktualizujem bojové plány...</span>';
        container.appendChild(overlay);
        container.style.overflow = "hidden";
    }
    overlay.style.display = 'flex';

    if (btnGoToCalendar) {
        btnGoToCalendar.disabled = true;
        btnGoToCalendar.innerText = "NAČÍTAVAM KALENDÁR...";
        btnGoToCalendar.style.opacity = "0.5";
        btnGoToCalendar.style.cursor = "not-allowed";
    }

    if (calendarBtns) {
        calendarBtns.style.opacity = "0.5";
        calendarBtns.style.pointerEvents = "none";
    }

    try {
        const response = await fetch(SCRIPT_URL);
        if (response.ok) {
            globalData = await response.json();
            renderReviews(globalData.reviews);
            renderGallery(globalData.gallery);
        }
    } catch (e) {
        console.error('Chyba dát', e);
    } finally {
        overlay.style.display = 'none';
        container.style.overflow = "auto";

        if (btnGoToCalendar) {
            btnGoToCalendar.disabled = false;
            btnGoToCalendar.innerText = "VYBRAŤ TERMÍN";
            btnGoToCalendar.style.opacity = "1";
            btnGoToCalendar.style.cursor = "pointer";
        }

        if (calendarBtns) {
        calendarBtns.style.opacity = "1";
        calendarBtns.style.pointerEvents = "all";
    }

        updateUI();
    }
}

function updateUI() {
    const labels = ["Tento týždeň", "Budúci týždeň", "O 2 týždne", "O 3 týždne", "O 4 týždne"];
    document.getElementById('weekLabel').innerText = labels[currentWeekOffset] || "Vzdialený termín";
    document.getElementById('prevWeek').disabled = currentWeekOffset === 0;
    document.getElementById('nextWeek').disabled = currentWeekOffset === 4;
    renderTimetable();
}

function renderTimetable() {
    const timetable = document.getElementById('timetable');
    const selDate = document.getElementById('date').value;
    const selTime = document.getElementById('time').value;
    timetable.innerHTML = '';

    const todayObj = getCETime();
    const todayStr = getLocalDateString(todayObj);

    const today = new Date();
    const distanceToMonday = (today.getDay() === 0 ? -6 : 1 - today.getDay()) + (currentWeekOffset * 7);
    const targetMonday = new Date(today.setDate(today.getDate() + distanceToMonday));

    let minMin = 24 * 60; 
    let maxMin = 0;
    let dniDatumy = [];

    // Získame všetky sloty pre aktuálny týždeň
    for (let i = 0; i < 7; i++) {
        const d = new Date(targetMonday); d.setDate(targetMonday.getDate() + i);
        const dStr = getLocalDateString(d);
        
        // --- NOVINKA: Získame inteligentne rozrezané sloty ---
        const pSlots = getProcessedSlotsForDate(dStr);
        
        const formatovanyDen = String(d.getDate()).padStart(2, '0');
        const formatovanyMesiac = String(d.getMonth() + 1).padStart(2, '0');
        
        dniDatumy.push({ dStr, pSlots, label: `${formatovanyDen}.${formatovanyMesiac}.` });
        
        pSlots.forEach(s => {
            minMin = Math.min(minMin, s.startMin);
            maxMin = Math.max(maxMin, s.startMin + s.durationMins);
        });
    }

    // --- OPRAVA: Zohľadníme ZAVRETÉ bloky do rozmerov kalendára ---
    for (let i = 0; i < 7; i++) {
        const d = new Date(targetMonday); d.setDate(targetMonday.getDate() + i);
        const dStr = getLocalDateString(d);
        let zavrete = globalData.custom.filter(c => c.datum === dStr && c.stav === 'ZAVRETÉ');
        
        zavrete.forEach(z => {
            let start = casNaMinuty(z.od);
            let end = casNaMinuty(z.do);
            if (end < start) end += 1440; // Ošetrenie prechodu cez polnoc
            minMin = Math.min(minMin, start);
            maxMin = Math.max(maxMin, end);
        });
    }

    // // Ak po všetkom nemáme žiadne sloty (napr. len celotýždňová dovolenka)
    // if (minMin >= maxMin || minMin === 24 * 60) { 
    //     minMin = 9 * 60; 
    //     maxMin = 17 * 60; 
    // } else {
    //     // Poistka proti pretečeniu pod tabuľku (polnoc je hranica)
    //     if (maxMin > 1440) maxTime = 1440; 
    // }

    // vyuzivane pre vizualne zobrazenie aj zatvorenych casti 
    // (aby to neboli len pri jednotnom nastaveni casov napr 2 riadky - zvysok tabulky bol prazdny)
    let kvaziZaciatok = 10;
    let kvaziKoniec = 17;

    // Ak po všetkom nemáme žiadne sloty (napr. len celotýždňová dovolenka)
    if (minMin >= maxMin || minMin === 24 * 60) { 
        minMin = kvaziZaciatok * 60; 
        maxMin = kvaziKoniec * 60; 
    } else {
        // Zafixujeme, aby kalendár z dizajnových dôvodov VŽDY začínal aspoň o 9:00.
        // Ak však holič príde do práce o 6:00, kalendár začne od 6:00 (vyberie menšiu hodnotu).
        minMin = Math.min(minMin, kvaziZaciatok * 60);
        
        // Zafixujeme, aby kalendár VŽDY končil aspoň o 17:00.
        // Ak má otvorené do polnoci, roztiahne sa až tam.
        maxMin = Math.max(maxMin, kvaziKoniec * 60);

        // Opravená poistka proti pretečeniu pod tabuľku (polnoc je tvrdá hranica)
        if (maxMin > 1440) maxMin = 1440; 
    }

    minMin = Math.floor(minMin / 60) * 60;
    maxMin = Math.ceil(maxMin / 60) * 60;

    const pocetRiadkov = (maxMin - minMin) / 30;
    const dynamickyGrid = `50px repeat(${pocetRiadkov}, minmax(25px, 1fr))`;

    // ČASOVÝ STĹPEC
    const timeCol = document.createElement('div');
    timeCol.className = 'time-col';
    timeCol.style.gridTemplateRows = dynamickyGrid; 
    timeCol.innerHTML = `
        <div class="cell header diagonal">
            <span class="diagonal-den">Deň</span>
            <span class="diagonal-hod">Hod.</span>
        </div>`;
    for (let m = minMin; m < maxMin; m += 30) {
        timeCol.innerHTML += `<div class="cell">${minutyNaCas(m)}</div>`;
    }
    timetable.appendChild(timeCol);

    // STĹPCE PRE DNI
    const skratky = ['Pon', 'Uto', 'Str', 'Štv', 'Pia', 'Sob', 'Ned'];
    dniDatumy.forEach((den, i) => {
        const dayCol = document.createElement('div');
        dayCol.className = 'day-col';
        dayCol.style.gridTemplateRows = dynamickyGrid; 
        dayCol.setAttribute('data-date', den.dStr); 
        dayCol.innerHTML = `<div class="cell header">${skratky[i]}<br>${den.label}</div>`;

        let m = minMin;
        while (m < maxMin) {
            // Nájdeme, či pre túto minútu existuje slot
            const activeSlot = den.pSlots.find(s => s.startMin === m);

            if (activeSlot) {
                const cell = document.createElement('div');
                const booking = globalData.bookings.find(b => b.datum === activeSlot.originalDate && b.cas === activeSlot.originalTime);
                const isBooked = !!booking;

                cell.className = `cell ${isBooked ? 'booked' : 'open'}`;
                if (booking && booking.pending) {
                    // Niekto si termín práve objednal a ešte ho nepotvrdil v maile.
                    cell.classList.add('pending');
                    cell.title = 'Čaká na potvrdenie. Ak ho objednávateľ nepotvrdí do 15 minút, termín sa uvoľní.';
                }
                cell.setAttribute('data-min', m); 
                cell.setAttribute('data-orig-date', activeSlot.originalDate); 
                cell.setAttribute('data-orig-min', casNaMinuty(activeSlot.originalTime)); 
                cell.setAttribute('data-duration', activeSlot.durationMins);
                cell.setAttribute('data-slot-id', activeSlot.id); // Spoločné ID pre obe rozrezané časti
                
                // Dynamický span podľa toho, aký dlhý je kus slotu (napr. 30min = span 1, 60min = span 2, 90min = span 3)
                let spanRows = activeSlot.durationMins / 30;
                cell.style.gridRow = `span ${spanRows}`;

                const origCasZaciatku = activeSlot.originalTime;
                const endCasStr = minutyNaCas(casNaMinuty(origCasZaciatku) + 90);

                // Vizuálne prepojenie textov
                if (activeSlot.isContinuation) {
                    cell.innerHTML = `...<br>${endCasStr}`;
                    cell.classList.add('slot-tail');
                } else if (spanRows < 3) {
                    cell.innerHTML = `${origCasZaciatku}<br>...`;
                    cell.classList.add('slot-head');
                } else {
                    cell.innerHTML = `${origCasZaciatku}<br>-<br>${endCasStr}`;
                }

                // --- MAGICKÝ HOVER EFEKT: Rozsvieti všetky časti s rovnakým ID ---
                cell.addEventListener('mouseenter', () => {
                    document.querySelectorAll(`[data-slot-id="${activeSlot.id}"]`).forEach(el => el.classList.add('hover-active'));
                });
                cell.addEventListener('mouseleave', () => {
                    document.querySelectorAll(`[data-slot-id="${activeSlot.id}"]`).forEach(el => el.classList.remove('hover-active'));
                });

                if (!isBooked) {
                    // Magický hover efekt len pre voľné termíny
                    cell.addEventListener('mouseenter', () => {
                        document.querySelectorAll(`[data-slot-id="${activeSlot.id}"]`).forEach(el => el.classList.add('hover-active'));
                    });
                    cell.addEventListener('mouseleave', () => {
                        document.querySelectorAll(`[data-slot-id="${activeSlot.id}"]`).forEach(el => el.classList.remove('hover-active'));
                    });

                    cell.onclick = () => {
                        if (cell.classList.contains('past') || cell.classList.contains('locked-buffer')) return; 
                        
                        // Zákazník klikol na akúkoľvek časť, ale reálne bookujeme originálny čas
                        document.getElementById('date').value = activeSlot.originalDate;
                        document.getElementById('time').value = activeSlot.originalTime;

                        updateSelectedTerminUI();
                        renderTimetable();
                        if (!isDesktopLayout()) document.querySelector('.booking-container').scrollIntoView({ behavior: 'smooth', block: 'center' });
                    };
                }
                
                // Zvýraznenie už zvoleného termínu
                if (activeSlot.originalDate === selDate && activeSlot.originalTime === selTime) {
                    cell.classList.add('selected-preview');
                }
                
                dayCol.appendChild(cell);
                m += activeSlot.durationMins; // Preskočíme správny počet minút podľa dĺžky kusu
            } else {
                const empty = document.createElement('div');
                empty.className = 'cell closed';
                empty.setAttribute('data-min', m); 
                dayCol.appendChild(empty);
                m += 30;
            }
        }
        timetable.appendChild(dayCol);
    });

    updateTimeLine(); 
}

// --- NOVÁ FUNKCIA PRE ŽIVÝ ČAS A ZAMYKANIE ---
function updateTimeLine() {
    const todayObj = getCETime();
    const todayStr = getLocalDateString(todayObj);
    const currentMin = todayObj.getHours() * 60 + todayObj.getMinutes();

    document.querySelectorAll('.current-time-line').forEach(line => line.remove());
    document.querySelectorAll('.cell').forEach(c => c.style.zIndex = '');

    document.querySelectorAll('.day-col').forEach(dayCol => {
        const dateStr = dayCol.getAttribute('data-date');
        if (!dateStr) return;

        const isToday = dateStr === todayStr;
        const isPastDay = dateStr < todayStr;

        if (isToday) dayCol.classList.add('current-day');
        else dayCol.classList.remove('current-day');

        dayCol.querySelectorAll('.cell:not(.header)').forEach(cell => {
            const m = parseInt(cell.getAttribute('data-min'));
            const isBookable = cell.classList.contains('open') || cell.classList.contains('booked') || cell.classList.contains('selected-preview');
            const span = isBookable ? 90 : 30;

            // --- NOVÁ LOGIKA ROZLIŠOVANIA MINULOSTI A BUFFRA ---
            const buffer = isBookable ? 30 : 0; 
            
            // 1. Tvrdá minulosť (čas naozaj ubehol)
            const isPastReal = isPastDay || (isToday && currentMin >= m);
            
            // 2. Uzamknutý buffer (čas ešte neubehol, ale je to menej ako 30 minút do štartu)
            const isLockedBuffer = !isPastReal && isToday && (currentMin >= m - buffer) && isBookable && !cell.classList.contains('booked');

            // Najprv očistíme bunku od starých stavov
            cell.classList.remove('past', 'locked-buffer');
            cell.removeAttribute('data-tooltip');

            // Priradíme správny stav
            if (isPastReal) {
                cell.classList.add('past'); // Šedé, úplne neaktívne
            } else if (isLockedBuffer) {
                cell.classList.add('locked-buffer'); // Tmavé zlaté, nedá sa kliknúť, ale funguje hover
                cell.setAttribute('data-tooltip', 'Tento termín bolo možné rezervovať max. 30 minút vopred.');
            }

            if (isToday && currentMin >= m && currentMin < m + span) {
                const topPercent = ((currentMin - m) / span) * 100;
                const line = document.createElement('div');
                line.className = 'current-time-line';
                line.style.top = `${topPercent}%`;
                cell.appendChild(line);
                cell.style.zIndex = '10'; 
            }
        });
    });
}

// Úprava submitu - kontrola e-mailu a toast notifikácie
// Úprava submitu - kontrola e-mailu, toast notifikácie a spinner
document.getElementById("booking-form").addEventListener("submit", async function(event) {
    event.preventDefault();

    const datum = document.getElementById("date").value;
    const cas = document.getElementById("time").value;
    const meno = document.getElementById("name").value;
    const email = document.getElementById("email").value;
    const sluzba = document.getElementById("service").value;
    const websiteEl = document.getElementById("website");
    const website = websiteEl ? websiteEl.value : ""; // honeypot proti botom - človek toto pole nevidí a nevyplní

    // --- REGEX KONTROLA E-MAILU ---
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        showToast("Bojovník, tvoj e-mail nevyzerá správne. Skontroluj ho!", "error");
        return;
    }

    if (!sluzba) {
        showToast("Vyber si službu, bojovník!", "error");
        document.getElementById("service").focus();
        return;
    }

    if (!datum || !cas) {
        showToast("Najprv si musíš zvoliť termín v tabuľke obsadenosti!", "error");
        document.querySelector('.timetable-section').scrollIntoView({ behavior: 'smooth' });
        return;
    }

    const btn = document.querySelector('.btn-submit');
    const btnText = btn.querySelector('.btn-text'); // Chytíme textovú časť tlačidla
    const kalendar = document.querySelector('.timetable-container');
    
    // --- ZAPNUTIE SPINNERA ---
    btn.classList.add('loading');
    btnText.innerText = "OBJEDNÁVAM...";
    btn.disabled = true;
    kalendar.style.pointerEvents = "none";

    try {
        // Bez mode:'no-cors' - vďaka tomu vieme prečítať skutočnú odpoveď servera (napr. "termín je už
        // obsadený" alebo "zadaj platný e-mail") namiesto toho, aby sme vždy len predpokladali úspech.
        const response = await fetch(SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ meno, email, sluzba, datum, cas, website })
        });

        let result = null;
        try {
            result = await response.json();
        } catch (parseErr) {
            result = null; // odpoveď sa nepodarilo prečítať - správame sa ako predtým (predpokladáme úspech)
        }

        if (result && result.error) {
            showToast(result.error, "error");
        } else if (result && result.status === 'pending') {
            // Termín je držaný, mail s potvrdením odíde po odpočte v paneli (dovtedy sa dá opraviť e-mail).
            startPending({ id: result.id, key: result.key, meno, termin: formatTermin(datum, cas, sluzba) }, {
                state: 'waiting', email: result.email, sendInMs: result.sendInMs, changesLeft: result.changesLeft, holdMinutes: result.holdMinutes
            });

            document.getElementById("booking-form").reset();
            document.getElementById("date").value = "";
            document.getElementById("time").value = "";
            updateSelectedTerminUI();
            refreshData();
        } else {
            // Úspešný toast namiesto alertu
            showToast(`Sila a česť, ${meno}! Objednávka bola odoslaná.`, "success");

            document.getElementById("booking-form").reset();
            document.getElementById("date").value = "";
            document.getElementById("time").value = "";
            updateSelectedTerminUI();

            // TRIK PRE TLAČIDLO: Odstránili sme slovo "await", aby sa tlačidlo
            // odblokovalo okamžite a tabuľka sa načítala na pozadí
            refreshData();
        }

    } catch (e) {
        console.error('Chyba!', e);
        showToast("Chyba v boji. Skús to znova!", "error");
    } finally {
        // --- VYPNUTIE SPINNERA ---
        btn.classList.remove('loading');
        btnText.innerText = "ODOSLAŤ OBJEDNÁVKU";
        kalendar.style.pointerEvents = "all";
        btn.disabled = false;
    }
});

// Aktualizuje vizuál poľa "Zvolený termín"
function updateSelectedTerminUI() {
    const dateVal = document.getElementById('date').value;
    const timeVal = document.getElementById('time').value;
    const emptyBox = document.getElementById('termin-selector-empty');
    const filledBox = document.getElementById('termin-selector-filled');
    const textSpan = document.getElementById('selected-termin-text');

    if (dateVal && timeVal) {
        // Dátum z inputu je vždy YYYY-MM-DD (takže už obsahuje nuly, napr. 2026-05-04)
        const [y, m, d] = dateVal.split('-');
        const endCasVal = minutyNaCas(casNaMinuty(timeVal) + 90);
        
        // Vypíše: 04.05.2026 | 08:00 - 09:30
        textSpan.innerText = `${d}.${m}.${y} | ${timeVal} - ${endCasVal}`;
        
        emptyBox.style.display = 'none';
        filledBox.style.display = 'flex';
    } else {
        emptyBox.style.display = 'block';
        filledBox.style.display = 'none';
    }
}

// Bezpečná funkcia na získanie aktuálneho času v Bratislave
function getCETime() {
    const now = new Date();
    // Prevedie aktuálny čas do bratislavskej zóny a vráti ho ako nový Date objekt
    const baTimeStr = now.toLocaleString('en-US', { timeZone: 'Europe/Bratislava' });
    return new Date(baTimeStr);
}

// --- FUNKCIA PRE VYSKAKOVACIE SPRÁVY (TOAST) ---
function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    // textContent (nie innerHTML) - message môže obsahovať meno, ktoré zákazník sám zadal do formulára,
    // takže ho nesmieme vykresľovať ako HTML (inak by šlo vložiť napr. <img onerror=...> a spustiť si JS).
    toast.textContent = message;

    container.appendChild(toast);

    // Správa zmizne po 4 sekundách
    setTimeout(() => {
        toast.style.animation = 'fadeOutToast 0.5s ease forwards';
        setTimeout(() => toast.remove(), 500); // Fyzické vymazanie po animácii
    }, 4000);
}
// --- POTVRDENIE REZERVÁCIE E-MAILOM ---
// Po objednávke je termín len držaný. Panel ukáže, na akú adresu pôjde mail, a počas krátkeho odpočtu sa dá
// e-mail opraviť alebo rezervácia zrušiť (ako "Vrátiť späť" v Gmaile). Potom odíde mail s tlačidlom
// POTVRDIŤ TERMÍN a panel sa každých pár sekúnd pozrie, či už zákazník klikol. Kľúč od servera (ním sa overuje,
// že zmeny robí ten, kto rezerváciu vytvoril) je v sessionStorage tejto karty, aby panel prežil aj obnovenie stránky.
const PENDING_STORAGE_KEY = 'barbarPending';
const PENDING_POLL_MS = 5000;
const PENDING_FINAL = ['confirmed', 'expired', 'gone', 'cancelled'];
const pendingEmailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

let pending = null; // { id, key, meno, termin }
const ps = {
    state: null, email: '', changesLeft: 0,
    sendAt: 0, sendTotal: 10000, expiresAt: 0, holdMs: 15 * 60000,
    busy: false, polling: false, editing: false, pausedMs: 0, lastPoll: 0,
    minimized: false, sendFailures: 0, cancelArmedUntil: 0,
};
let pendingTimer = null;
let pendingSeq = 0;      // poradové číslo požiadavky
let pendingMinSeq = 0;   // odpovede starších požiadaviek (napr. kontrola stavu spred zmeny e-mailu) sa zahodia

const pEl = (id) => document.getElementById(id);
const mmss = (ms) => { const s = Math.ceil(ms / 1000); return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`; };

function formatTermin(datum, cas, sluzba) {
    const [y, m, d] = datum.split('-');
    return `${d}.${m}.${y} o ${cas}` + (sluzba ? ` · ${sluzba}` : '');
}

function savePending() {
    try {
        if (pending && !PENDING_FINAL.includes(ps.state)) sessionStorage.setItem(PENDING_STORAGE_KEY, JSON.stringify(pending));
        else sessionStorage.removeItem(PENDING_STORAGE_KEY);
    } catch (e) { /* súkromné okno a pod.: panel funguje, len neprežije obnovenie stránky */ }
}

function loadPending() {
    try { return JSON.parse(sessionStorage.getItem(PENDING_STORAGE_KEY) || 'null'); } catch (e) { return null; }
}

// Vráti [poradové číslo, odpoveď]. userAction = odpovede skorších kontrol stavu sa po nej už neuplatnia.
async function pendingApi(action, extra, userAction) {
    const seq = ++pendingSeq;
    if (userAction) pendingMinSeq = seq;
    const response = await fetch(SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(Object.assign({ action, id: pending.id, key: pending.key }, extra || {}))
    });
    return [seq, await response.json()];
}

function resetPendingState() {
    Object.assign(ps, { state: null, email: '', changesLeft: 0, busy: false, polling: false, editing: false, minimized: false, sendFailures: 0, cancelArmedUntil: 0, lastPoll: 0 });
    pEl('pending-error').textContent = '';
}

function startPending(info, state) {
    pending = info;
    resetPendingState();
    applyPendingState(state);
    showPendingPanel();
}

// Prevezme stav zo servera. Časy prichádzajú relatívne (ms odteraz), takže nezáleží na hodinách v zariadení.
function applyPendingState(s, seq) {
    if (!s || !pending) return;
    if (seq && seq < pendingMinSeq) return;
    if (s.error) { pEl('pending-error').textContent = s.error; renderPending(); return; }
    if (PENDING_FINAL.includes(ps.state)) return; // výsledok sa už nemení

    if (s.state !== ps.state) pEl('pending-error').textContent = '';
    if (s.email) ps.email = s.email;
    if (typeof s.changesLeft === 'number') ps.changesLeft = s.changesLeft;
    if (s.holdMinutes) ps.holdMs = s.holdMinutes * 60000;
    if (s.state === 'waiting') {
        ps.sendTotal = Math.max(1000, s.sendInMs || 0);
        ps.sendAt = Date.now() + (s.sendInMs || 0);
    }
    if (s.state === 'sent') ps.expiresAt = Date.now() + (typeof s.expiresInMs === 'number' ? s.expiresInMs : ps.holdMs);

    const changed = s.state !== ps.state;
    ps.state = s.state;
    savePending();

    if (PENDING_FINAL.includes(ps.state)) {
        ps.editing = false;
        stopPendingTimer();
        if (changed) refreshData(); // termín sa v kalendári natrvalo obsadil alebo uvoľnil
        if (ps.minimized) showPendingPanel(); // výsledok ukážeme, aj keď bol panel zmenšený
    } else {
        startPendingTimer();
    }
    renderPending();
}

function renderPending() {
    if (!pending) return;
    const st = ps.state;
    const active = st === 'waiting' || st === 'sent';
    const set = (id, text) => { pEl(id).textContent = text || ''; };
    const show = (id, visible) => { pEl(id).hidden = !visible; };

    let title = '', text = '', hint = '';
    if (st === 'waiting') {
        title = 'Skontroluj svoj e-mail';
        text = 'Potvrdenie termínu pošleme na:';
        hint = 'Preklep? Teraz ho ešte opravíš, mail zatiaľ neodišiel.';
    } else if (st === 'sent') {
        title = 'Potvrď termín v maile';
        text = 'Poslali sme ti mail na:';
        hint = `Klikni v ňom na POTVRDIŤ TERMÍN. Bez potvrdenia sa termín o ${Math.round(ps.holdMs / 60000)} minút uvoľní. Nevidíš ho? Pozri aj priečinok Spam alebo Promo akcie.`;
    } else if (st === 'confirmed') {
        title = 'Termín je potvrdený ✓';
        text = `Sila a česť, ${pending.meno}! Tešíme sa na teba na B02/321, Purkyňovy koleje.`;
        hint = 'Ak budeš potrebovať termín zrušiť, použi tlačidlo v maile.';
    } else if (st === 'cancelled') {
        title = 'Rezervácia je zrušená';
        text = 'Termín je znova voľný. Ak chceš iný, vyber si ho v kalendári.';
    } else if (st === 'gone') {
        title = 'Rezervácia už nie je aktívna';
        text = 'Bola zrušená alebo nepotvrdená včas, termín je znova voľný.';
    } else {
        title = 'Čas na potvrdenie vypršal';
        text = 'Termín sa uvoľnil, lebo nebol potvrdený včas. Ak ho stále chceš, vyber si ho znova.';
    }

    set('pending-title', title);
    set('pending-termin', pending.termin);
    set('pending-text', ps.editing ? 'Napíš správnu adresu, mail pošleme tam:' : text);
    set('pending-email', active && !ps.editing ? ps.email : '');
    set('pending-hint', ps.editing ? '' : hint);

    show('pending-progress', active && !ps.editing);
    show('pending-change', ps.editing);
    show('pending-actions', !ps.editing);
    show('pending-send-now', st === 'waiting');
    show('pending-change-btn', active && ps.changesLeft > 0);
    show('pending-cancel-btn', active);
    show('pending-close-btn', st !== 'waiting');
    pEl('pending-close-btn').textContent = st === 'sent' ? 'ZAVRIEŤ, POTVRDÍM V MAILI' : 'ZAVRIEŤ';

    const armed = Date.now() < ps.cancelArmedUntil;
    pEl('pending-cancel-btn').textContent = armed ? 'NAOZAJ ZRUŠIŤ?' : 'ZRUŠIŤ REZERVÁCIU';
    pEl('pending-cancel-btn').classList.toggle('armed', armed);
    ['pending-send-now', 'pending-change-btn', 'pending-cancel-btn', 'pending-change-save', 'pending-change-back']
        .forEach(id => { pEl(id).disabled = ps.busy; });

    updatePendingClock();
}

function updatePendingClock() {
    const now = Date.now();
    let timer = '';
    let ratio = 0;
    if (ps.state === 'waiting') {
        const left = Math.max(0, ps.sendAt - now);
        if (ps.editing) timer = 'Odpočet stojí, kým upravuješ e-mail.';
        else if (ps.sendFailures >= 3 && !ps.busy) timer = 'Mail sa zatiaľ nepodarilo odoslať, skúsime to ešte automaticky.';
        else if (ps.busy || left === 0) timer = 'Posielam mail...';
        else timer = `Mail odošleme o ${Math.ceil(left / 1000)} s`;
        ratio = ps.editing ? 0 : left / ps.sendTotal;
    } else if (ps.state === 'sent') {
        const left = Math.max(0, ps.expiresAt - now);
        timer = `Na potvrdenie zostáva ${mmss(left)}`;
        ratio = left / ps.holdMs;
        pEl('pending-mini-text').textContent = `Termín čaká na potvrdenie v maile · ${mmss(left)}`;
    }
    pEl('pending-timer').textContent = timer;
    pEl('pending-bar').style.width = `${Math.max(0, Math.min(1, ratio)) * 100}%`;
}

function startPendingTimer() {
    if (!pendingTimer) pendingTimer = setInterval(pendingTick, 250);
}

function stopPendingTimer() {
    clearInterval(pendingTimer);
    pendingTimer = null;
}

function pendingTick() {
    if (!pending || PENDING_FINAL.includes(ps.state)) { stopPendingTimer(); return; }
    const now = Date.now();
    if (ps.cancelArmedUntil && now >= ps.cancelArmedUntil) { ps.cancelArmedUntil = 0; renderPending(); }

    if (!ps.busy && !ps.editing && ps.state === 'waiting' && now >= ps.sendAt && ps.sendFailures < 3) {
        sendPendingNow();
    } else if (!ps.busy && !ps.polling && document.visibilityState === 'visible' && now - ps.lastPoll >= PENDING_POLL_MS &&
               (ps.state === 'sent' || (ps.state === 'waiting' && ps.sendFailures >= 3))) {
        // čakáme, či zákazník klikol v maile (alebo či mail za nás poslal server, ak sa nedalo spojiť)
        pollPending();
    }
    updatePendingClock();
}

async function sendPendingNow() {
    ps.busy = true;
    renderPending();
    try {
        const [seq, s] = await pendingApi('send', null, true);
        ps.busy = false;
        if (s.error) { ps.sendFailures++; ps.sendAt = Date.now() + 3000; }
        applyPendingState(s, seq);
    } catch (e) {
        ps.busy = false;
        ps.sendFailures++;
        ps.sendAt = Date.now() + 3000;
        pEl('pending-error').textContent = ps.sendFailures >= 3
            ? 'Nepodarilo sa spojiť so serverom. Mail ti odošleme automaticky do pár minút.'
            : 'Spojenie zlyhalo, skúšam znova...';
        renderPending();
    }
}

async function pollPending() {
    ps.lastPoll = Date.now();
    ps.polling = true;
    try {
        const [seq, s] = await pendingApi('status');
        applyPendingState(s, seq);
    } catch (e) {
        // bez spojenia: skúsi sa pri ďalšom kole
    } finally {
        ps.polling = false;
    }
}

async function pendingUserAction(action, extra) {
    ps.busy = true;
    pEl('pending-error').textContent = '';
    renderPending();
    try {
        const [seq, s] = await pendingApi(action, extra, true);
        ps.busy = false;
        return [seq, s];
    } catch (e) {
        ps.busy = false;
        pEl('pending-error').textContent = 'Spojenie zlyhalo, skús to prosím znova.';
        renderPending();
        return [0, null];
    }
}

function showPendingPanel() {
    ps.minimized = false;
    pEl('pending-mini').hidden = true;
    pEl('pending-overlay').hidden = false;
    document.body.style.overflow = 'hidden';
    renderPending();
    const first = ['pending-send-now', 'pending-close-btn'].map(pEl).find(b => !b.hidden);
    if (first) first.focus();
}

function closePendingPanel() {
    pEl('pending-overlay').hidden = true;
    document.body.style.overflow = '';
    if (ps.state === 'sent') {
        // zákazník ide do mailu: panel sa len zmenší a ďalej sleduje, či termín potvrdil
        ps.minimized = true;
        pEl('pending-mini').hidden = false;
    } else {
        pEl('pending-mini').hidden = true;
        stopPendingTimer();
        pending = null;
        savePending();
    }
}

function setupPending() {
    if (!pEl('pending-overlay')) return;

    pEl('pending-send-now').addEventListener('click', () => {
        if (ps.busy || ps.state !== 'waiting') return;
        ps.sendFailures = 0;
        sendPendingNow();
    });

    pEl('pending-change-btn').addEventListener('click', () => {
        if (ps.busy) return;
        ps.editing = true;
        ps.pausedMs = Math.max(0, ps.sendAt - Date.now());
        pEl('pending-new-email').value = ps.email;
        pEl('pending-error').textContent = '';
        renderPending();
        pEl('pending-new-email').focus();
        pEl('pending-new-email').select();
    });

    pEl('pending-change-back').addEventListener('click', () => {
        ps.editing = false;
        if (ps.state === 'waiting') ps.sendAt = Date.now() + Math.max(ps.pausedMs, 3000);
        pEl('pending-error').textContent = '';
        renderPending();
    });

    pEl('pending-change').addEventListener('submit', async (e) => {
        e.preventDefault();
        if (ps.busy) return;
        const email = pEl('pending-new-email').value.trim();
        if (!pendingEmailRegex.test(email)) { pEl('pending-error').textContent = 'Tento e-mail nevyzerá správne. Skontroluj ho.'; return; }
        if (email.toLowerCase() === ps.email.toLowerCase()) { pEl('pending-error').textContent = 'To je ten istý e-mail ako doteraz.'; return; }
        const [seq, s] = await pendingUserAction('changeEmail', { email });
        if (!s) return;
        if (!s.error) ps.editing = false;
        ps.sendFailures = 0;
        applyPendingState(s, seq);
    });

    pEl('pending-cancel-btn').addEventListener('click', async () => {
        if (ps.busy) return;
        if (Date.now() >= ps.cancelArmedUntil) {
            ps.cancelArmedUntil = Date.now() + 4000; // druhé kliknutie do 4 s naozaj zruší
            renderPending();
            return;
        }
        ps.cancelArmedUntil = 0;
        const [seq, s] = await pendingUserAction('cancel');
        if (s) applyPendingState(s, seq);
    });

    pEl('pending-close-btn').addEventListener('click', closePendingPanel);
    pEl('pending-mini-open').addEventListener('click', showPendingPanel);

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && !pEl('pending-overlay').hidden && ps.state !== 'waiting' && !ps.editing) closePendingPanel();
    });

    // Po návrate z mailovej aplikácie sa hneď pozrieme, či už je termín potvrdený.
    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible' && pending && ps.state === 'sent') ps.lastPoll = 0;
    });
}

async function resumePending() {
    const saved = loadPending();
    if (!saved || !saved.id || !saved.key) return;
    pending = saved;
    resetPendingState();
    try {
        const [seq, s] = await pendingApi('status');
        if (!s || s.error) { pending = null; savePending(); return; }
        applyPendingState(s, seq);
        showPendingPanel();
    } catch (e) {
        pending = null; // bez spojenia: uložený stav ostane, skúsi sa pri ďalšom načítaní
    }
}

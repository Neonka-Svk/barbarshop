// TODO: dalo by sa to refaktornut ngl xdd

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzQyJbBAKbE6doPZOgNCcIDiUnsU70AZ_WhMULilehSN4VEN-5i8fHM_t39KUTvMCoq/exec'; 
let currentWeekOffset = 0; 
let globalObsadeneTerminy = [];
let globalData = { bookings: [], custom: [], defaultHours: [] };

function getLocalDateString(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
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

function getRawSlotsForDate(dateStr) {
    let rawSlots = [];
    const denneCustom = globalData.custom.filter(c => c.datum === dateStr);
    
    // Rozdelíme si stavy
    const otvoreneCustom = denneCustom.filter(c => c.stav === 'OTVORENÉ');
    const rozsireneCustom = denneCustom.filter(c => c.stav === 'ROZŠÍRENÉ');
    const zavreteCustom = denneCustom.filter(c => c.stav === 'ZAVRETÉ');

    if (otvoreneCustom.length > 0) {
        // Ak má vlastné OTVORENÉ, úplne ignorujeme bežné hodiny
        // Pre istotu k nim ale prilepíme aj ROZŠÍRENÉ (ak by si náhodou zadal oba stavy v ten istý deň)
        const vsetkyOtvorene = [...otvoreneCustom, ...rozsireneCustom];
        
        vsetkyOtvorene.forEach(r => {
            let start = casNaMinuty(r.od);
            let koniec = casNaMinuty(r.do);
            if (koniec < start) koniec += 24 * 60; 
            while (start + 90 <= koniec) {
                rawSlots.push(start);
                start += 90;
            }
        });
    } else {
        // Ak NEMA vlastne otvorene, pouzijeme BEŽNÉ hodiny
        const d = new Date(dateStr);
        let dayOfWeek = d.getDay(); 
        if (dayOfWeek === 0) dayOfWeek = 7;

        const rozpisNaDnes = globalData.defaultHours.find(h => h.den === dayOfWeek);

        if (rozpisNaDnes && rozpisNaDnes.bloky && rozpisNaDnes.bloky.length > 0) {
            rozpisNaDnes.bloky.forEach(blok => {
                let start = casNaMinuty(blok.od);
                let koniec = casNaMinuty(blok.do);
                if (koniec < start) koniec += 24 * 60; 
                while (start + 90 <= koniec) {
                    rawSlots.push(start);
                    start += 90;
                }
            });
        }
        
        // ... a k bežným hodinám PRIDÁME "ROZŠÍRENÉ" bloky
        if (rozsireneCustom.length > 0) {
            rozsireneCustom.forEach(r => {
                let start = casNaMinuty(r.od);
                let koniec = casNaMinuty(r.do);
                if (koniec < start) koniec += 24 * 60; 
                while (start + 90 <= koniec) {
                    rawSlots.push(start);
                    start += 90;
                }
            });
        }
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

document.addEventListener('DOMContentLoaded', async () => {
    await refreshData();

    // Navigácia týždňov
    document.getElementById('nextWeek').addEventListener('click', () => {
        if (currentWeekOffset < 4) { currentWeekOffset++; updateUI(); }
    });
    document.getElementById('prevWeek').addEventListener('click', () => {
        if (currentWeekOffset > 0) { currentWeekOffset--; updateUI(); }
    });

    document.getElementById('btn-go-to-calendar').addEventListener('click', () => {
        document.querySelector('.timetable-section').scrollIntoView({ behavior: 'smooth' });
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

function updateAvailableTimes() {
    const dateInput = document.getElementById('date');
    const timeSelect = document.getElementById('time');
    const dateVal = dateInput.value;
    const povodnyVyber = timeSelect.value;
    
    if (!dateVal) return;

    const selectedDate = new Date(dateVal);
    const dayIndex = selectedDate.getDay(); 
    const isWeekend = (dayIndex === 0 || dayIndex === 6);
    const defaultCasy = isWeekend ? vikendoveCasy : pracovneCasy;
    
    timeSelect.innerHTML = '<option value="">Vyber si čas</option>';

    // Zistíme, či pre tento deň existujú adminom pridané extra časy
    const extraCasy = globalObsadeneTerminy
        .filter(t => t.datum === dateVal && t.meno === "ADMIN_OPEN")
        .map(t => t.cas);
    
    // Spojíme bežné časy s extra časmi a odstránime duplicity
    const vsetkyMozneCasy = [...new Set([...defaultCasy, ...extraCasy])].sort();

    vsetkyMozneCasy.forEach(timeStr => {
        const zaznam = globalObsadeneTerminy.find(t => t.datum === dateVal && t.cas === timeStr);
        const isBooked = zaznam && zaznam.meno !== "ADMIN_OPEN";
        const isAdminBlocked = zaznam && zaznam.meno === "ADMIN_BLOCK";

        if (isAdminBlocked) return; // Admin tento čas úplne skryl

        const opt = document.createElement('option');
        opt.value = timeStr;
        opt.innerText = timeStr;
        if (isBooked) { opt.disabled = true; opt.innerText += ' (Obsadené)'; }
        timeSelect.appendChild(opt);
    });

    // KONTROLA: Ak bol čas vybratý a po zmene dňa už nie je v ponuke, vynuluj ho
    const staleDostupny = Array.from(timeSelect.options).some(o => o.value === povodnyVyber && !o.disabled);
    timeSelect.value = staleDostupny ? povodnyVyber : "";

    updateUI(); 
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

    // Ak po všetkom nemáme žiadne sloty (napr. len celotýždňová dovolenka)
    if (minMin >= maxMin || minMin === 24 * 60) { 
        minMin = 9 * 60; 
        maxMin = 17 * 60; 
    } else {
        // Poistka proti pretečeniu pod tabuľku (polnoc je hranica)
        if (maxMin > 1440) maxTime = 1440; 
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
                const isBooked = globalData.bookings.some(b => b.datum === activeSlot.originalDate && b.cas === activeSlot.originalTime);
                
                cell.className = `cell ${isBooked ? 'booked' : 'open'}`;
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
                        document.querySelector('.booking-container').scrollIntoView({ behavior: 'smooth', block: 'center' });
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

    // --- REGEX KONTROLA E-MAILU ---
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        showToast("Bojovník, tvoj e-mail nevyzerá správne. Skontroluj ho!", "error");
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
        await fetch(SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors', 
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ meno, email, sluzba, datum, cas })
        });

        // Úspešný toast namiesto alertu
        showToast(`Sila a česť, ${meno}! Objednávka bola odoslaná.`, "success");
        
        document.getElementById("booking-form").reset();
        document.getElementById("date").value = "";
        document.getElementById("time").value = "";
        updateSelectedTerminUI();

        // TRIK PRE TLAČIDLO: Odstránili sme slovo "await", aby sa tlačidlo 
        // odblokovalo okamžite a tabuľka sa načítala na pozadí
        refreshData(); 

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
    toast.innerHTML = message;

    container.appendChild(toast);

    // Správa zmizne po 4 sekundách
    setTimeout(() => {
        toast.style.animation = 'fadeOutToast 0.5s ease forwards';
        setTimeout(() => toast.remove(), 500); // Fyzické vymazanie po animácii
    }, 4000);
}
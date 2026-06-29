// ============================================================
//  MOEVA PILATES — Google Apps Script Backend
//  Kopiraj to kodo v Google Apps Script (script.google.com)
//  in jo poveži s tvojim Google Sheets dokumentom.
// ============================================================

const SHEET_ID   = '1Tp-gLGgM-n2W1LolyXjtjTNQaXPMuuIcX4-6AmDOXmU';      // tvoj Sheet ID ✓
const ADMIN_EMAIL = 'moeva.pilates@gmail.com';   // tvoj email za obvestila ✓
const STUDIO_IME  = 'MOEVA PILATES';

// Koliko ur pred treningom pošlji opomnik strankam
const OPOMNIK_URE = 12; // pošlji opomnik 12h pred treningom

// Pri koliko preostalih urah paketa pošlji opozorilo
const PAKET_OPOZORILO_URE = 2;

const SHEETS = {
  clients:   'Stranke',
  slots:     'Termini',
  bookings:  'Rezervacije',
  payments:  'Placila',
  waitlist:  'CakalnaLista'
};

// ════════════════════════════════════════════════════════════
//  EMAIL HELPERS
// ════════════════════════════════════════════════════════════

function sendEmail(to, subject, htmlBody) {
  try {
    // Odstrani emojije iz zadeve (subject) — v nekaterih programih se pokvarijo
    const cleanSubject = subject.replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}]/gu, '').trim();
    GmailApp.sendEmail(to, cleanSubject, '', {
      htmlBody: '<meta charset="UTF-8">' + htmlBody,
      name: STUDIO_IME
    });
  } catch(e) {
    console.log('Email napaka (' + to + '): ' + e.message);
  }
}

function sendAdminEmail(subject, htmlBody) {
  sendEmail(ADMIN_EMAIL, subject, htmlBody);
}

function emailTemplate(title, emoji, rows, footer) {
  // Pretvori emoji v barvno piko (emojiji se v nekaterih email programih pokvarijo)
  const dotColors = {
    '🟢':'#6A9E7A', '🔴':'#C17A6A', '👤':'#8A9E8C', '✏️':'#7A6E5F',
    '🗑️':'#999', '📅':'#8A9E8C', '💶':'#6A9E7A', '✅':'#6A9E7A',
    '📋':'#7A6E5F', '⏰':'#E6A23C', '⚠️':'#C17A6A', '🎉':'#B8860B',
    '📨':'#8A9E8C'
  };
  const dotColor = dotColors[emoji] || '#8A9E8C';
  const dot = `<span style="display:inline-block;width:12px;height:12px;border-radius:50%;background:${dotColor};margin-right:8px;vertical-align:middle"></span>`;

  const rowsHtml = rows.map(([label, value]) =>
    `<tr>
      <td style="padding:7px 14px;color:#7A6E5F;font-size:13px;width:130px;vertical-align:top">${label}</td>
      <td style="padding:7px 14px;font-size:13px;font-weight:500;color:#2C2C2C">${value||'—'}</td>
    </tr>`
  ).join('');
  const footerHtml = footer
    ? `<div style="padding:14px 24px;font-size:13px;color:#2C2C2C;background:#EEF3EE;border-top:1px solid #C5D4C6">${footer}</div>`
    : '';
  return `
  <div style="font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;max-width:480px;margin:0 auto;background:#FAF9F7;border-radius:10px;overflow:hidden;border:1px solid #DDD8D0">
    <div style="background:#2C2C2C;padding:22px 24px">
      <div style="font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#C5D4C6;margin-bottom:6px">${STUDIO_IME}</div>
      <div style="font-size:22px;color:#FAF9F7;font-weight:300">${dot}${title}</div>
    </div>
    <table style="width:100%;border-collapse:collapse;background:white;margin:0">${rowsHtml}</table>
    ${footerHtml}
    <div style="padding:12px 24px;font-size:11px;color:#aaa;border-top:1px solid #DDD8D0">
      ${new Date().toLocaleString('sl-SI')} · Avtomatsko obvestilo
    </div>
  </div>`;
}

function fmtSlot(slot) {
  if (!slot) return '—';
  const d = new Date(slot.datum + 'T00:00:00');
  const dni = ['ned','pon','tor','sre','čet','pet','sob'];
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  return `${dni[d.getDay()]}, ${dd}/${mm}/${d.getFullYear()} ob ${slot.cas}`;
}

// ════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      // Google Sheets shrani datume/čase kot Date objekte — pretvorimo v tekst
      if (v instanceof Date) {
        if (h === 'datum' || h === 'rok' || h === 'datum_vpisa' ||
            h === 'datum_rezervacije' || h === 'datum_placila' || h === 'datum_prijave') {
          // Datum → "yyyy-MM-dd"
          v = Utilities.formatDate(v, Session.getScriptTimeZone() || 'Europe/Ljubljana', 'yyyy-MM-dd');
        } else if (h === 'cas') {
          // Čas → "HH:mm"
          v = Utilities.formatDate(v, Session.getScriptTimeZone() || 'Europe/Ljubljana', 'HH:mm');
        } else {
          v = Utilities.formatDate(v, Session.getScriptTimeZone() || 'Europe/Ljubljana', 'yyyy-MM-dd');
        }
      }
      obj[h] = v;
    });
    return obj;
  });
}

function uid() { return Utilities.getUuid().substring(0, 8); }

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) return i + 1;
  }
  return -1;
}

function getMonday(d) {
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}

// ════════════════════════════════════════════════════════════
//  SETUP
// ════════════════════════════════════════════════════════════

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const defs = {
    Stranke:      ['id','ime','telefon','email','paket','ure_skupaj','ure_porabljene','opomba','datum_vpisa','st_odpovedi','st_noshow'],
    Termini:      ['id','datum','cas','tip','max_mest','naziv','aktiven'],
    Rezervacije:  ['id','slot_id','client_id','ime','email','telefon','datum_rezervacije','status','prisotnost'],
    Placila:      ['id','client_id','opis','znesek','rok','status','datum_placila'],
    CakalnaLista: ['id','slot_id','ime','email','telefon','datum_prijave']
  };
  Object.entries(defs).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#8A9E8C').setFontColor('white');
    }
  });
  return { ok: true, message: 'Listi pripravljeni!' };
}

// ════════════════════════════════════════════════════════════
//  POSODOBI STRUKTURO (za obstoječe baze)
//  Zaženi ENKRAT če si že imel staro verzijo brez novih stolpcev
// ════════════════════════════════════════════════════════════
function posodobiStrukturo() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Dodaj manjkajoče stolpce v Stranke
  const strankeSheet = ss.getSheetByName(SHEETS.clients);
  const strankeHead = strankeSheet.getRange(1,1,1,strankeSheet.getLastColumn()).getValues()[0];
  if (strankeHead.indexOf('st_odpovedi') < 0) {
    strankeSheet.getRange(1, strankeHead.length+1).setValue('st_odpovedi').setFontWeight('bold').setBackground('#8A9E8C').setFontColor('white');
  }
  if (strankeHead.indexOf('st_noshow') < 0) {
    const col = strankeSheet.getLastColumn()+1;
    strankeSheet.getRange(1, col).setValue('st_noshow').setFontWeight('bold').setBackground('#8A9E8C').setFontColor('white');
  }

  // Dodaj prisotnost v Rezervacije
  const rezSheet = ss.getSheetByName(SHEETS.bookings);
  const rezHead = rezSheet.getRange(1,1,1,rezSheet.getLastColumn()).getValues()[0];
  if (rezHead.indexOf('prisotnost') < 0) {
    rezSheet.getRange(1, rezHead.length+1).setValue('prisotnost').setFontWeight('bold').setBackground('#8A9E8C').setFontColor('white');
  }

  return { ok: true, message: 'Struktura posodobljena!' };
}

// ════════════════════════════════════════════════════════════
//  HTTP HANDLERS
// ════════════════════════════════════════════════════════════

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let result;
  try {
    if      (action === 'getSlots')      result = getSlots(ss);
    else if (action === 'getClients')    result = getClients(ss);
    else if (action === 'getBookings')   result = getBookings(ss);
    else if (action === 'getPayments')   result = getPayments(ss);
    else if (action === 'getDashboard')  result = getDashboard(ss);
    else if (action === 'getWaitlist')   result = getWaitlist(ss);
    else result = { error: 'Neznan ukaz' };
  } catch(err) { result = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let result;
  try {
    if      (action === 'addClient')       result = addClient(ss, data);
    else if (action === 'updateClient')    result = updateClient(ss, data);
    else if (action === 'deleteClient')    result = deleteClient(ss, data);
    else if (action === 'addSlot')         result = addSlot(ss, data);
    else if (action === 'updateSlot')      result = updateSlot(ss, data);
    else if (action === 'deleteSlot')      result = deleteSlot(ss, data);
    else if (action === 'addBooking')      result = addBooking(ss, data);
    else if (action === 'cancelBooking')   result = cancelBooking(ss, data);
    else if (action === 'addPayment')      result = addPayment(ss, data);
    else if (action === 'updatePayment')   result = updatePayment(ss, data);
    else if (action === 'deletePayment')   result = deletePayment(ss, data);
    else if (action === 'addWaitlist')     result = addWaitlist(ss, data);
    else if (action === 'removeWaitlist')  result = removeWaitlist(ss, data);
    else if (action === 'oznaciPrisotnost') result = oznaciPrisotnost(ss, data);
    else if (action === 'ponavljajocaRezervacija') result = ponavljajocaRezervacija(ss, data);
    else result = { error: 'Neznan ukaz' };
  } catch(err) { result = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  CLIENTS
// ════════════════════════════════════════════════════════════

function getClients(ss) {
  return sheetToObjects(ss.getSheetByName(SHEETS.clients));
}

function addClient(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.clients);
  const id = uid();
  const novaVrstica = sheet.getLastRow() + 1;
  // Stolpec telefon (C = 3) naj bo TEKST, da se ohrani začetna nič
  sheet.getRange(novaVrstica, 3).setNumberFormat('@');
  sheet.appendRow([
    id, data.ime, data.telefon||'', data.email||'',
    data.paket||'', data.ure_skupaj||0, 0,
    data.opomba||'', new Date().toISOString().slice(0,10)
  ]);
  sendAdminEmail(`👤 Nova stranka: ${data.ime}`,
    emailTemplate('Nova stranka dodana','👤',[
      ['Ime', data.ime], ['Telefon', data.telefon], ['Email', data.email],
      ['Paket', data.paket], ['Ur v paketu', data.ure_skupaj||'—'], ['Opomba', data.opomba]
    ]));
  return { ok: true, id };
}

function updateClient(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.clients);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Stranka ni najdena' };
  // Ohrani obstoječe porabljene ure, če niso poslane
  const obstojeca = sheetToObjects(sheet).find(x => String(x.id) === String(data.id));
  const porabljene = data.ure_porabljene!==undefined ? data.ure_porabljene : (obstojeca?.ure_porabljene||0);
  // Stolpec telefon (C = 3) naj bo TEKST, da se ohrani začetna nič
  sheet.getRange(row, 3).setNumberFormat('@');
  sheet.getRange(row, 2, 1, 7).setValues([[
    data.ime, data.telefon||'', data.email||'',
    data.paket||'', data.ure_skupaj||0, porabljene, data.opomba||''
  ]]);
  sendAdminEmail(`✏️ Stranka posodobljena: ${data.ime}`,
    emailTemplate('Podatki stranke posodobljeni','✏️',[
      ['Ime', data.ime], ['Paket', data.paket],
      ['Obiski skupaj', data.ure_skupaj||'—'], ['Porabljeni', porabljene]
    ]));
  return { ok: true };
}

function deleteClient(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.clients);
  const c = sheetToObjects(sheet).find(x => x.id === data.id);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Stranka ni najdena' };
  sheet.deleteRow(row);
  sendAdminEmail(`🗑️ Stranka izbrisana: ${c?.ime||'—'}`,
    emailTemplate('Stranka izbrisana','🗑️',[['Ime', c?.ime], ['Email', c?.email], ['Paket', c?.paket]]));
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  SLOTS
// ════════════════════════════════════════════════════════════

function getSlots(ss) {
  const sheet = ss.getSheetByName(SHEETS.slots);
  const bookSheet = ss.getSheetByName(SHEETS.bookings);
  const waitSheet = ss.getSheetByName(SHEETS.waitlist);
  const slots = sheetToObjects(sheet).filter(s => s.aktiven != false && s.aktiven !== 'false');
  const bookings = sheetToObjects(bookSheet).filter(b => b.status === 'potrjena');
  const waitlist = sheetToObjects(waitSheet);

  const counts = {};
  bookings.forEach(b => { counts[b.slot_id] = (counts[b.slot_id]||0) + 1; });

  const waitCounts = {};
  waitlist.forEach(w => { waitCounts[w.slot_id] = (waitCounts[w.slot_id]||0) + 1; });

  return slots.map(s => ({
    ...s,
    rezervirano: counts[s.id] || 0,
    prostih: Math.max(0, (parseInt(s.max_mest)||1) - (counts[s.id]||0)),
    cakalna: waitCounts[s.id] || 0
  }));
}

function addSlot(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.slots);
  const id = uid();
  sheet.appendRow([id, data.datum, data.cas, data.tip||'Individualni', data.max_mest||1, data.naziv||'', true]);
  sendAdminEmail(`📅 Nov termin: ${data.datum} ob ${data.cas}`,
    emailTemplate('Nov termin dodan','📅',[
      ['Datum', data.datum], ['Čas', data.cas],
      ['Vrsta', data.tip], ['Naziv', data.naziv], ['Mesta', data.max_mest||1]
    ]));
  return { ok: true, id };
}

function updateSlot(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.slots);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Termin ni najden' };
  sheet.getRange(row, 2, 1, 6).setValues([[
    data.datum, data.cas, data.tip||'Individualni',
    data.max_mest||1, data.naziv||'', data.aktiven !== false
  ]]);
  sendAdminEmail(`✏️ Termin spremenjen: ${data.datum} ob ${data.cas}`,
    emailTemplate('Termin posodobljen','✏️',[
      ['Datum', data.datum], ['Čas', data.cas], ['Vrsta', data.tip], ['Mesta', data.max_mest]
    ]));
  return { ok: true };
}

function deleteSlot(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.slots);
  const s = sheetToObjects(sheet).find(x => x.id === data.id);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Termin ni najden' };
  sheet.deleteRow(row);
  sendAdminEmail(`🗑️ Termin izbrisan: ${s?.datum||'—'} ob ${s?.cas||'—'}`,
    emailTemplate('Termin izbrisan','🗑️',[['Datum', s?.datum], ['Čas', s?.cas], ['Vrsta', s?.tip]]));
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  BOOKINGS
// ════════════════════════════════════════════════════════════

function getBookings(ss) {
  return sheetToObjects(ss.getSheetByName(SHEETS.bookings));
}

function addBooking(ss, data) {
  const slotSheet = ss.getSheetByName(SHEETS.slots);
  const bookSheet = ss.getSheetByName(SHEETS.bookings);
  const clientSheet = ss.getSheetByName(SHEETS.clients);

  const slots = sheetToObjects(slotSheet);
  const slot = slots.find(s => s.id === data.slot_id);
  if (!slot) return { error: 'Termin ne obstaja' };

  // ── PREVERI ali je stranka v bazi (po emailu ALI telefonu) ──
  const clients = sheetToObjects(clientSheet);
  const vnesEmail = (data.email||'').trim().toLowerCase();
  const vnesTel = (data.telefon||'').replace(/\s|\/|-/g,''); // odstrani presledke
  const najdena = clients.find(c => {
    const cEmail = (c.email||'').trim().toLowerCase();
    const cTel = (c.telefon||'').toString().replace(/\s|\/|-/g,'');
    return (vnesEmail && cEmail === vnesEmail) || (vnesTel && cTel && cTel === vnesTel);
  });
  if (!najdena) {
    return { error: 'Ta email/telefon ni v naši bazi strank. Za prijavo se obrnite na Evo.' };
  }

  const existing = sheetToObjects(bookSheet).filter(b => b.slot_id === data.slot_id && b.status === 'potrjena');
  if (existing.length >= parseInt(slot.max_mest||1)) return { error: 'Termin je zaseden' };

  // Preveri duplikat (po client_id, emailu ali telefonu)
  const dup = existing.find(b =>
    b.client_id === najdena.id ||
    (vnesEmail && (b.email||'').trim().toLowerCase() === vnesEmail)
  );
  if (dup) return { error: 'Že imate rezervacijo za ta termin' };

  const id = uid();
  bookSheet.appendRow([
    id, data.slot_id, najdena.id,
    najdena.ime, najdena.email||'', najdena.telefon||'',
    new Date().toISOString().slice(0,10), 'potrjena', ''
  ]);

  // Uporabi podatke iz baze (ne kar je vnesel uporabnik)
  data.ime = najdena.ime;
  data.email = najdena.email;
  data.telefon = najdena.telefon;
  data.client_id = najdena.id;

  // ── Paket sledenje: odštej 1 uro ──────────────────────────
  let paketOpozorilo = null;
  if (data.client_id) {
    const clients = sheetToObjects(clientSheet);
    const client = clients.find(c => c.id === data.client_id);
    if (client && parseInt(client.ure_skupaj||0) > 0) {
      const clientRow = findRowById(clientSheet, data.client_id);
      const porabljene = parseInt(client.ure_porabljene||0) + 1;
      clientSheet.getRange(clientRow, 7).setValue(porabljene); // stolpec 7 = ure_porabljene
      const ostalo = parseInt(client.ure_skupaj||0) - porabljene;
      if (ostalo <= PAKET_OPOZORILO_URE && ostalo >= 0) {
        paketOpozorilo = ostalo;
      }
    }
  }

  const prostihPo = parseInt(slot.max_mest||1) - existing.length - 1;
  sendAdminEmail(`🟢 Nova rezervacija: ${data.ime}`,
    emailTemplate('Nova rezervacija','🟢',[
      ['Stranka', data.ime], ['Email', data.email], ['Telefon', data.telefon],
      ['Termin', fmtSlot(slot)], ['Vrsta', slot.tip], ['Prosto še', `${prostihPo} mest`]
    ], paketOpozorilo !== null
      ? `⚠️ <strong>${data.ime}</strong> ima v paketu le še <strong>${paketOpozorilo} ur</strong>.`
      : null
  ));

  // ── Potrdilni email stranki ────────────────────────────────
  if (data.email) {
    sendEmail(data.email, `Potrjena rezervacija — ${STUDIO_IME}`,
      emailTemplate('Rezervacija potrjena 🌿','✅',[
        ['Ime', data.ime],
        ['Termin', fmtSlot(slot)],
        ['Vrsta', slot.tip||'—'],
        ['Naziv', slot.naziv||'—'],
      ], 'Odpoved je možna do 4 ure pred treningom. Se vidimo! 🧘')
    );
  }

  return { ok: true, id };
}

function cancelBooking(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.bookings);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Rezervacija ni najdena' };

  const booking = sheetToObjects(sheet).find(b => b.id === data.id);
  let slot = null;

  if (booking) {
    const slotSheet = ss.getSheetByName(SHEETS.slots);
    slot = sheetToObjects(slotSheet).find(s => s.id === booking.slot_id);

    // 4h preverjanje
    if (slot && slot.datum && slot.cas) {
      const slotDT = new Date(slot.datum + 'T' + slot.cas + ':00');
      const diffH = (slotDT - new Date()) / (1000*60*60);
      if (diffH < 4) return { error: 'Odpoved ni mogoča — do termina je manj kot 4 ure.' };
    }

    // ── Paket sledenje: vrni 1 uro ─────────────────────────
    if (booking.client_id) {
      const clientSheet = ss.getSheetByName(SHEETS.clients);
      const client = sheetToObjects(clientSheet).find(c => c.id === booking.client_id);
      if (client) {
        const clientRow = findRowById(clientSheet, booking.client_id);
        // Vrni uro v paket
        if (parseInt(client.ure_skupaj||0) > 0) {
          const porabljene = Math.max(0, parseInt(client.ure_porabljene||0) - 1);
          clientSheet.getRange(clientRow, 7).setValue(porabljene);
        }
        // Preštej odpoved (st_odpovedi je stolpec 10)
        const head = clientSheet.getRange(1,1,1,clientSheet.getLastColumn()).getValues()[0];
        const odpCol = head.indexOf('st_odpovedi') + 1;
        if (odpCol > 0) {
          const trenutno = parseInt(client.st_odpovedi||0) + 1;
          clientSheet.getRange(clientRow, odpCol).setValue(trenutno);
        }
      }
    }

    sendAdminEmail(`🔴 Odpoved: ${booking.ime}`,
      emailTemplate('Rezervacija odpovedana','🔴',[
        ['Stranka', booking.ime], ['Email', booking.email], ['Telefon', booking.telefon],
        ['Termin', fmtSlot(slot)], ['Čas odpovedi', new Date().toLocaleString('sl-SI')]
      ]));

    // Obvesti prvega na čakalni listi
    notifyWaitlist(ss, booking.slot_id, slot);
  }

  sheet.getRange(row, 8).setValue('odpovedana');
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  WAITLIST — čakalna lista
// ════════════════════════════════════════════════════════════

function getWaitlist(ss) {
  return sheetToObjects(ss.getSheetByName(SHEETS.waitlist));
}

function addWaitlist(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.waitlist);

  // Preveri duplikat
  const existing = sheetToObjects(sheet).find(w => w.slot_id === data.slot_id && w.email === data.email);
  if (existing) return { error: 'Že ste na čakalni listi za ta termin.' };

  const id = uid();
  sheet.appendRow([id, data.slot_id, data.ime, data.email||'', data.telefon||'', new Date().toISOString().slice(0,10)]);

  // Poišči info o terminu
  const slotSheet = ss.getSheetByName(SHEETS.slots);
  const slot = sheetToObjects(slotSheet).find(s => s.id === data.slot_id);

  sendAdminEmail(`📋 Čakalna lista: ${data.ime}`,
    emailTemplate('Nova prijava na čakalno listo','📋',[
      ['Stranka', data.ime], ['Email', data.email], ['Telefon', data.telefon],
      ['Termin', fmtSlot(slot)]
    ]));

  // Potrdilni email osebi na čakalni listi
  if (data.email) {
    sendEmail(data.email, `Čakalna lista — ${STUDIO_IME}`,
      emailTemplate('Na čakalni listi ste','📋',[
        ['Ime', data.ime],
        ['Termin', fmtSlot(slot)],
      ], 'Takoj ko se sprosti mesto, vas obvestimo. Hvala za potrpežljivost! 🌿')
    );
  }

  return { ok: true, id };
}

function removeWaitlist(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.waitlist);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Vnos ni najden' };
  sheet.deleteRow(row);
  return { ok: true };
}

function notifyWaitlist(ss, slotId, slot) {
  const sheet = ss.getSheetByName(SHEETS.waitlist);
  const waiting = sheetToObjects(sheet).filter(w => w.slot_id === slotId);
  if (!waiting.length) return;

  // Obvesti prvega na listi
  const first = waiting[0];
  if (first.email) {
    sendEmail(first.email, `🎉 Sprostilo se je mesto — ${STUDIO_IME}`,
      emailTemplate('Sprostilo se je mesto!','🎉',[
        ['Termin', fmtSlot(slot)],
        ['Vrsta', slot?.tip||'—'],
      ], `Hitro rezervirajte — mesto je odprto na prvi pridejo, prvi serve! 🧘`)
    );
  }

  sendAdminEmail(`📋 Obveščen čakalnik: ${first.ime}`,
    emailTemplate('Čakalnik obveščen o prostem mestu','📋',[
      ['Stranka', first.ime], ['Email', first.email], ['Termin', fmtSlot(slot)]
    ]));
}

// ════════════════════════════════════════════════════════════
//  PAYMENTS
// ════════════════════════════════════════════════════════════

function getPayments(ss) {
  return sheetToObjects(ss.getSheetByName(SHEETS.payments));
}

function addPayment(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.payments);
  const id = uid();
  const clients = sheetToObjects(ss.getSheetByName(SHEETS.clients));
  const client = clients.find(c => c.id === data.client_id);
  sheet.appendRow([id, data.client_id, data.opis||'', data.znesek, data.rok||'', data.status||'caka', '']);
  const statusLabel = {caka:'Čaka', placano:'Plačano', zamuda:'Zamuda'}[data.status||'caka'];
  sendAdminEmail(`💶 Novo plačilo: ${client?.ime||'—'}`,
    emailTemplate('Novo plačilo dodano','💶',[
      ['Stranka', client?.ime], ['Opis', data.opis],
      ['Znesek', `${parseFloat(data.znesek||0).toFixed(2)} €`],
      ['Rok', data.rok], ['Status', statusLabel]
    ]));
  return { ok: true, id };
}

function updatePayment(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.payments);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Plačilo ni najdeno' };
  const clients = sheetToObjects(ss.getSheetByName(SHEETS.clients));
  const client = clients.find(c => c.id === data.client_id);
  const wasPaid = data.status === 'placano';
  sheet.getRange(row, 3, 1, 5).setValues([[
    data.opis||'', data.znesek, data.rok||'', data.status,
    wasPaid ? new Date().toISOString().slice(0,10) : ''
  ]]);
  const statusLabel = {caka:'Čaka', placano:'✅ Plačano', zamuda:'⚠️ Zamuda'}[data.status]||data.status;
  sendAdminEmail(`${wasPaid?'✅':'💶'} Plačilo ${wasPaid?'potrjeno':'posodobljeno'}: ${client?.ime||'—'}`,
    emailTemplate(wasPaid?'Plačilo potrjeno':'Plačilo posodobljeno', wasPaid?'✅':'💶',[
      ['Stranka', client?.ime], ['Opis', data.opis],
      ['Znesek', `${parseFloat(data.znesek||0).toFixed(2)} €`], ['Status', statusLabel]
    ]));
  return { ok: true };
}

function deletePayment(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.payments);
  const p = sheetToObjects(sheet).find(x => x.id === data.id);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Plačilo ni najdeno' };
  sheet.deleteRow(row);
  const clients = sheetToObjects(ss.getSheetByName(SHEETS.clients));
  const client = clients.find(c => c.id === p?.client_id);
  sendAdminEmail(`🗑️ Plačilo izbrisano: ${client?.ime||'—'}`,
    emailTemplate('Plačilo izbrisano','🗑️',[
      ['Stranka', client?.ime], ['Opis', p?.opis],
      ['Znesek', `${parseFloat(p?.znesek||0).toFixed(2)} €`]
    ]));
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  OPOMNIK ~12H PRED TRENINGOM — zaženi kot časovni sprožilec
//     Apps Script → Ure → Dodaj sprožilec → sendDayBeforeReminders
//     Nastavi: VSAKO URO (Hour timer → Every hour)
//     Pošlje opomnik strankam, katerih termin je čez ~12 ur.
// ════════════════════════════════════════════════════════════

function sendDayBeforeReminders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const slots = sheetToObjects(ss.getSheetByName(SHEETS.slots));
  const bookings = sheetToObjects(ss.getSheetByName(SHEETS.bookings));

  const zdaj = new Date();
  let skupajPoslano = 0;

  slots.forEach(slot => {
    if (!slot.datum || !slot.cas) return;
    if (slot.aktiven === false || slot.aktiven === 'false') return;

    // Sestavi datum+čas termina
    const [y,m,d] = slot.datum.split('-').map(Number);
    const [hh,mm] = slot.cas.split(':').map(Number);
    const terminDT = new Date(y, m-1, d, hh, mm, 0);

    // Koliko ur do termina
    const urDoTermina = (terminDT - zdaj) / (1000*60*60);

    // Pošlji če je med (OPOMNIK_URE - 0.5) in (OPOMNIK_URE + 0.5) ur
    // → ujame 12h okno, če sprožilec teče vsako uro
    if (urDoTermina < (OPOMNIK_URE - 0.5) || urDoTermina > (OPOMNIK_URE + 0.5)) return;

    const slotBookings = bookings.filter(b => b.slot_id === slot.id && b.status === 'potrjena');
    slotBookings.forEach(b => {
      if (!b.email) return;
      sendEmail(b.email, `Opomnik: trening ob ${slot.cas} — ${STUDIO_IME}`,
        emailTemplate('Opomnik za trening 🌿','⏰',[
          ['Ime', b.ime],
          ['Datum', fmtSlot(slot)],
          ['Čas', slot.cas],
          ['Naziv', slot.naziv||'Pilates'],
        ], 'Vaš trening je čez približno 12 ur. Odpoved je možna do 4 ure pred treningom. Se vidimo! 🧘')
      );
      skupajPoslano++;
    });
  });

  // Obvestilo tebi koliko opomnikov je šlo ven
  if (skupajPoslano > 0) {
    sendAdminEmail(`📨 Poslano ${skupajPoslano} opomnikov`,
      emailTemplate('Opomniki poslani','📨',[
        ['Datum', tomorrowStr],
        ['Terminov', tomorrowSlots.length],
        ['Opomnikov poslanih', skupajPoslano]
      ]));
  }
}

// ════════════════════════════════════════════════════════════
//  DASHBOARD
// ════════════════════════════════════════════════════════════

function getDashboard(ss) {
  const clients  = sheetToObjects(ss.getSheetByName(SHEETS.clients));
  const slots    = sheetToObjects(ss.getSheetByName(SHEETS.slots));
  const bookings = sheetToObjects(ss.getSheetByName(SHEETS.bookings));
  const payments = sheetToObjects(ss.getSheetByName(SHEETS.payments));

  const today = new Date().toISOString().slice(0,10);
  const weekStart = getMonday(new Date());
  const weekEnd   = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 6);

  const weekSlots = slots.filter(s =>
    s.datum >= weekStart.toISOString().slice(0,10) &&
    s.datum <= weekEnd.toISOString().slice(0,10));

  const paidMonth = payments
    .filter(p => p.status === 'placano' && p.rok && p.rok.slice(0,7) === today.slice(0,7))
    .reduce((s,p) => s + parseFloat(p.znesek||0), 0);

  const unpaid = payments
    .filter(p => p.status !== 'placano')
    .reduce((s,p) => s + parseFloat(p.znesek||0), 0);

  // Stranke z malo urami v paketu
  const paketOpozorila = clients.filter(c => {
    const skupaj = parseInt(c.ure_skupaj||0);
    const porabljene = parseInt(c.ure_porabljene||0);
    return skupaj > 0 && (skupaj - porabljene) <= PAKET_OPOZORILO_URE;
  }).map(c => ({
    ime: c.ime,
    ostalo: parseInt(c.ure_skupaj||0) - parseInt(c.ure_porabljene||0)
  }));

  return {
    stevilo_strank: clients.length,
    termini_teden: weekSlots.length,
    placano_mesec: paidMonth,
    neplacano: unpaid,
    paket_opozorila: paketOpozorila
  };
}

// ════════════════════════════════════════════════════════════
//  1. TEDENSKI POVZETEK — vsak ponedeljek ob 7:00
//     Apps Script → ⏰ → Dodaj sprožilec → sendWeeklySummary
//     Nastavi: vsak teden, ponedeljek, 7:00–8:00
// ════════════════════════════════════════════════════════════

function sendWeeklySummary() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const clients  = sheetToObjects(ss.getSheetByName(SHEETS.clients));
  const slots    = sheetToObjects(ss.getSheetByName(SHEETS.slots));
  const bookings = sheetToObjects(ss.getSheetByName(SHEETS.bookings));
  const payments = sheetToObjects(ss.getSheetByName(SHEETS.payments));

  const weekStart = getMonday(new Date());
  const weekEnd   = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 6);
  const wsStr = weekStart.toISOString().slice(0,10);
  const weStr = weekEnd.toISOString().slice(0,10);

  // Termini ta teden
  const weekSlots = slots
    .filter(s => s.datum >= wsStr && s.datum <= weStr && s.aktiven !== 'false')
    .sort((a,b) => (a.datum+a.cas).localeCompare(b.datum+b.cas));

  // Rezervacije za ta teden
  const weekBookings = bookings.filter(b =>
    b.status === 'potrjena' &&
    weekSlots.some(s => s.id === b.slot_id)
  );

  // Pričakovani prihodki (neplačani z rokom ta teden)
  const weekPayments = payments.filter(p =>
    p.status !== 'placano' && p.rok >= wsStr && p.rok <= weStr
  );
  const pricakovano = weekPayments.reduce((s,p) => s+parseFloat(p.znesek||0), 0);

  // Neplačano skupaj
  const neplacano = payments
    .filter(p => p.status !== 'placano')
    .reduce((s,p) => s+parseFloat(p.znesek||0), 0);

  // Zgradi HTML urnik po dnevih
  const dni = ['ned','pon','tor','sre','čet','pet','sob'];
  let urnikHtml = '';
  weekSlots.forEach(s => {
    const prijavljeni = weekBookings.filter(b => b.slot_id === s.id);
    const d = new Date(s.datum);
    const full = prijavljeni.length >= parseInt(s.max_mest||1);
    urnikHtml += `
      <tr style="border-top:1px solid #EEF3EE">
        <td style="padding:8px 14px;font-size:13px;color:#7A6E5F;white-space:nowrap">${dni[d.getDay()]} ${d.getDate()}.${d.getMonth()+1}.</td>
        <td style="padding:8px 14px;font-size:13px;font-weight:500">${s.cas} · ${s.tip}</td>
        <td style="padding:8px 14px;font-size:13px;text-align:center">
          <span style="background:${full?'#fce4ec':'#EEF3EE'};color:${full?'#880e4f':'#2e7d32'};padding:2px 8px;border-radius:20px;font-size:11px;font-weight:500">
            ${prijavljeni.length}/${s.max_mest||1}
          </span>
        </td>
        <td style="padding:8px 14px;font-size:12px;color:#aaa">${prijavljeni.map(b=>b.ime).join(', ')||'—'}</td>
      </tr>`;
  });

  const html = `
  <div style="font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;max-width:560px;margin:0 auto;background:#FAF9F7;border-radius:10px;overflow:hidden;border:1px solid #DDD8D0">
    <div style="background:#2C2C2C;padding:22px 24px">
      <div style="font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#C5D4C6;margin-bottom:6px">${STUDIO_IME}</div>
      <div style="font-size:22px;color:#FAF9F7;font-weight:300">📅 Tedenski povzetek</div>
      <div style="font-size:13px;color:#aaa;margin-top:4px">${wsStr} – ${weStr}</div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;background:white;border-bottom:1px solid #EEF3EE">
      <div style="padding:16px 18px;text-align:center;border-right:1px solid #EEF3EE">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#7A6E5F;margin-bottom:4px">Terminov</div>
        <div style="font-size:28px;font-weight:300;color:#2C2C2C">${weekSlots.length}</div>
      </div>
      <div style="padding:16px 18px;text-align:center;border-right:1px solid #EEF3EE">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#7A6E5F;margin-bottom:4px">Prijavljenih</div>
        <div style="font-size:28px;font-weight:300;color:#2C2C2C">${weekBookings.length}</div>
      </div>
      <div style="padding:16px 18px;text-align:center">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#7A6E5F;margin-bottom:4px">Neplačano</div>
        <div style="font-size:28px;font-weight:300;color:#C17A6A">${neplacano.toFixed(0)} €</div>
      </div>
    </div>
    ${weekSlots.length > 0 ? `
    <table style="width:100%;border-collapse:collapse;background:white">
      <thead><tr style="background:#EEF3EE">
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Dan</th>
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Termin</th>
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:center;font-weight:500">Mesta</th>
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Stranke</th>
      </tr></thead>
      <tbody>${urnikHtml}</tbody>
    </table>` : '<div style="padding:20px;text-align:center;color:#aaa;font-size:13px">Ta teden ni terminov.</div>'}
    <div style="padding:12px 24px;font-size:11px;color:#aaa;border-top:1px solid #DDD8D0">
      Samodejni tedenski povzetek · ${STUDIO_IME}
    </div>
  </div>`;

  sendAdminEmail(`📅 Tedenski povzetek ${wsStr}`, html);
}

// ════════════════════════════════════════════════════════════
//  2. NEPLAČANI OPOMNIK — vsak dan ob 9:00
//     Apps Script → ⏰ → Dodaj sprožilec → sendOverdueReminders
//     Nastavi: vsak dan, 9:00–10:00
// ════════════════════════════════════════════════════════════

function sendOverdueReminders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const payments = sheetToObjects(ss.getSheetByName(SHEETS.payments));
  const clients  = sheetToObjects(ss.getSheetByName(SHEETS.clients));

  const today = new Date().toISOString().slice(0,10);

  // Najdi zamude (rok pretečen, še ni plačano)
  const overdue = payments.filter(p =>
    p.status !== 'placano' && p.rok && p.rok < today
  );

  if (!overdue.length) return;

  // Posodobi status na 'zamuda'
  const sheet = ss.getSheetByName(SHEETS.payments);
  overdue.forEach(p => {
    if (p.status !== 'zamuda') {
      const row = findRowById(sheet, p.id);
      if (row > 0) sheet.getRange(row, 6).setValue('zamuda');
    }
  });

  // Pošlji tebi povzetek vseh zamud
  const rows = overdue.map(p => {
    const client = clients.find(c => c.id === p.client_id);
    return [client?.ime||'—', `${parseFloat(p.znesek||0).toFixed(2)} €`, p.rok||'—', p.opis||'—'];
  });

  const rowsHtml = rows.map(([ime, znesek, rok, opis]) => `
    <tr style="border-top:1px solid #EEF3EE">
      <td style="padding:8px 14px;font-size:13px;font-weight:500">${ime}</td>
      <td style="padding:8px 14px;font-size:13px;color:#C17A6A;font-weight:500">${znesek}</td>
      <td style="padding:8px 14px;font-size:13px;color:#aaa">${rok}</td>
      <td style="padding:8px 14px;font-size:13px;color:#aaa">${opis}</td>
    </tr>`).join('');

  const skupaj = overdue.reduce((s,p) => s+parseFloat(p.znesek||0), 0);

  const html = `
  <div style="font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;max-width:520px;margin:0 auto;background:#FAF9F7;border-radius:10px;overflow:hidden;border:1px solid #DDD8D0">
    <div style="background:#2C2C2C;padding:22px 24px">
      <div style="font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#C5D4C6;margin-bottom:6px">${STUDIO_IME}</div>
      <div style="font-size:22px;color:#FAF9F7;font-weight:300">⚠️ Neplačani računi</div>
    </div>
    <div style="background:white;padding:14px 18px;border-bottom:1px solid #EEF3EE">
      <span style="font-size:13px;color:#7A6E5F">Skupaj neplačano: </span>
      <strong style="font-size:16px;color:#C17A6A">${skupaj.toFixed(2)} €</strong>
    </div>
    <table style="width:100%;border-collapse:collapse;background:white">
      <thead><tr style="background:#EEF3EE">
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Stranka</th>
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Znesek</th>
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Rok</th>
        <th style="padding:8px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7A6E5F;text-align:left;font-weight:500">Opis</th>
      </tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>
    <div style="padding:12px 24px;font-size:11px;color:#aaa;border-top:1px solid #DDD8D0">
      ${today} · ${STUDIO_IME}
    </div>
  </div>`;

  sendAdminEmail(`⚠️ ${overdue.length} neplačanih računov — ${skupaj.toFixed(2)} €`, html);

  // Pošlji tudi email vsaki stranki (če ima email)
  overdue.forEach(p => {
    const client = clients.find(c => c.id === p.client_id);
    if (!client?.email) return;
    sendEmail(client.email, `Opomnik za plačilo — ${STUDIO_IME}`,
      emailTemplate('Opomnik za plačilo','⚠️',[
        ['Opis',    p.opis||'—'],
        ['Znesek',  `${parseFloat(p.znesek||0).toFixed(2)} €`],
        ['Rok',     p.rok||'—'],
      ], 'Prosimo, uredite plačilo čim prej. Hvala! 🌿')
    );
  });
}

// ════════════════════════════════════════════════════════════
//  USTVARI FIKSNE TEDENSKE TERMINE
//  MOEVA urnik:
//    Ponedeljek 20:00–21:00
//    Torek      20:00–21:00
//    Sreda      18:00–18:45
//    Četrtek    20:00–21:00
//
//  ZAŽENI: izberi 'generirajTermine' → Zaženi
//  Ustvari termine za 6 mesecev vnaprej (preskoči obstoječe in praznike)
// ════════════════════════════════════════════════════════════

function generirajTermine() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.slots);
  const obstojeci = sheetToObjects(sheet);

  // Urnik: dan v tednu (1=pon, 2=tor, 3=sre, 4=čet) → {ura, mest}
  const URNIK = {
    1: { cas: '20:00', mest: 8 },  // ponedeljek
    2: { cas: '20:00', mest: 8 },  // torek
    3: { cas: '18:00', mest: 8 },  // sreda
    4: { cas: '20:00', mest: 8 }   // četrtek
  };

  // Slovenski prazniki (preskočimo jih)
  const PRAZNIKI = ['01-01','01-02','02-08','04-27','05-01','05-02',
    '06-25','08-15','10-31','11-01','12-25','12-26'];
  const PRAZNIKI_DATUM = ['2025-04-21','2026-04-06','2027-03-29']; // velikonočni ponedeljek

  const danes = new Date();
  danes.setHours(0,0,0,0);
  const konec = new Date(danes);
  konec.setMonth(konec.getMonth() + 6); // 6 mesecev vnaprej

  let ustvarjenih = 0;
  let preskocenih = 0;
  const novVrstice = [];

  for (let d = new Date(danes); d <= konec; d.setDate(d.getDate() + 1)) {
    const dan = d.getDay(); // 0=ned, 1=pon, ...
    if (!URNIK[dan]) continue; // samo pon-čet

    // Sestavi datum ročno (brez časovnega pasu, da se dan ne premakne)
    const datumStr = d.getFullYear() + '-' +
      String(d.getMonth()+1).padStart(2,'0') + '-' +
      String(d.getDate()).padStart(2,'0');
    const md = String(d.getMonth()+1).padStart(2,'0') + '-' +
      String(d.getDate()).padStart(2,'0');

    // Preskoči praznike
    if (PRAZNIKI.indexOf(md) >= 0 || PRAZNIKI_DATUM.indexOf(datumStr) >= 0) {
      preskocenih++;
      continue;
    }

    const cas = URNIK[dan].cas;

    // Preskoči če termin že obstaja (isti datum + ura)
    const zeObstaja = obstojeci.some(s => s.datum === datumStr && s.cas === cas) ||
                      novVrstice.some(r => r[1] === datumStr && r[2] === cas);
    if (zeObstaja) { preskocenih++; continue; }

    // Dodaj: [id, datum, cas, tip, max_mest, naziv, aktiven]
    novVrstice.push([uid(), datumStr, cas, 'Skupinski', URNIK[dan].mest, 'Pilates', true]);
    ustvarjenih++;
  }

  // Zapiši vse naenkrat (hitreje)
  if (novVrstice.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    // Stolpca datum (B) in cas (C) naj bosta TEKST, da ju Sheets ne pretvori v datum
    sheet.getRange(startRow, 2, novVrstice.length, 2).setNumberFormat('@');
    sheet.getRange(startRow, 1, novVrstice.length, 7).setValues(novVrstice);
  }

  Logger.log(`✅ Ustvarjenih ${ustvarjenih} terminov, preskočenih ${preskocenih} (obstoječi/prazniki).`);
  return { ok: true, ustvarjenih, preskocenih };
}

// ════════════════════════════════════════════════════════════
//  POBRIŠI STARE/PRETEKLE TERMINE (neobvezno čiščenje)
//  Zaženi občasno da Sheet ostane pregleden
// ════════════════════════════════════════════════════════════

function pobrisiPretekleTermine() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.slots);
  const danes = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  const data = sheet.getDataRange().getValues();

  let pobrisanih = 0;
  // Od zadaj naprej (da se vrstice ne premaknejo)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] && data[i][1] < danes) {
      sheet.deleteRow(i + 1);
      pobrisanih++;
    }
  }
  Logger.log(`🗑️ Pobrisanih ${pobrisanih} preteklih terminov.`);
  return { ok: true, pobrisanih };
}

// ════════════════════════════════════════════════════════════
//  POPRAVI & PONOVNO USTVARI TERMINE
//  Zaženi ENKRAT če so termini v napačnem formatu (datum/čas).
//  Pobriše vse termine in jih ustvari na novo pravilno.
// ════════════════════════════════════════════════════════════

function popraviTermine() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.slots);

  // Pobriši vse razen glave
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  Logger.log('🗑️ Stari termini pobrisani. Ustvarjam nove...');

  // Ustvari na novo (z pravilnim formatom)
  return generirajTermine();
}

// ════════════════════════════════════════════════════════════
//  PRISOTNOST — označi kdo je prišel / ni prišel (no-show)
//  data: { id (booking id), prisoten: true/false }
// ════════════════════════════════════════════════════════════

function oznaciPrisotnost(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.bookings);
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Rezervacija ni najdena' };

  const head = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const prisCol = head.indexOf('prisotnost') + 1;
  if (prisCol < 1) return { error: 'Stolpec prisotnost ne obstaja — zaženi posodobiStrukturo()' };

  const vrednost = data.prisoten ? 'prisoten' : 'ni-prisel';
  sheet.getRange(row, prisCol).setValue(vrednost);

  // Če ni prišel (no-show), preštej pri stranki
  const booking = sheetToObjects(sheet).find(b => b.id === data.id);
  if (booking && booking.client_id && !data.prisoten) {
    const clientSheet = ss.getSheetByName(SHEETS.clients);
    const client = sheetToObjects(clientSheet).find(c => c.id === booking.client_id);
    if (client) {
      const clientRow = findRowById(clientSheet, booking.client_id);
      const cHead = clientSheet.getRange(1,1,1,clientSheet.getLastColumn()).getValues()[0];
      const nsCol = cHead.indexOf('st_noshow') + 1;
      if (nsCol > 0) {
        const trenutno = parseInt(client.st_noshow||0) + 1;
        clientSheet.getRange(clientRow, nsCol).setValue(trenutno);
      }
    }
  }

  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  PONAVLJAJOČA REZERVACIJA
//  Rezervira isti dan+uro za več tednov vnaprej
//  data: { client_id, dan_v_tednu (1-4), cas, tednov (npr. 4) }
// ════════════════════════════════════════════════════════════

function ponavljajocaRezervacija(ss, data) {
  const slotSheet = ss.getSheetByName(SHEETS.slots);
  const bookSheet = ss.getSheetByName(SHEETS.bookings);
  const clientSheet = ss.getSheetByName(SHEETS.clients);

  const client = sheetToObjects(clientSheet).find(c => c.id === data.client_id);
  if (!client) return { error: 'Stranka ni najdena' };

  const slots = sheetToObjects(slotSheet);
  const tednov = parseInt(data.tednov) || 4;
  let rezerviranih = 0;
  let preskocenih = 0;
  const rezultati = [];

  // Najdi vse prihodnje termine z ujemajočim dnem in uro
  const danes = new Date(); danes.setHours(0,0,0,0);

  const ustreznilSlots = slots.filter(s => {
    if (!s.datum || !s.cas) return false;
    if (s.cas !== data.cas) return false;
    const d = new Date(s.datum + 'T00:00:00');
    if (d < danes) return false;
    return d.getDay() === parseInt(data.dan_v_tednu);
  }).sort((a,b) => a.datum.localeCompare(b.datum)).slice(0, tednov);

  ustreznilSlots.forEach(slot => {
    const existing = sheetToObjects(bookSheet).filter(b => b.slot_id === slot.id && b.status === 'potrjena');
    // Preskoči če zasedeno ali že rezervirano
    if (existing.length >= parseInt(slot.max_mest||1)) { preskocenih++; return; }
    if (existing.some(b => b.client_id === client.id)) { preskocenih++; return; }

    bookSheet.appendRow([
      uid(), slot.id, client.id, client.ime, client.email||'', client.telefon||'',
      new Date().toISOString().slice(0,10), 'potrjena', ''
    ]);
    rezerviranih++;
    rezultati.push(fmtSlot(slot));
  });

  if (rezerviranih > 0) {
    sendAdminEmail(`🟢 Ponavljajoča rezervacija: ${client.ime} (${rezerviranih}x)`,
      emailTemplate('Ponavljajoča rezervacija','🟢',[
        ['Stranka', client.ime],
        ['Rezerviranih', `${rezerviranih} terminov`],
        ['Termini', rezultati.join('<br>')]
      ]));
  }

  return { ok: true, rezerviranih, preskocenih };
}

// ════════════════════════════════════════════════════════════
//  MESEČNI IZVOZ — ustvari nov list z vsemi podatki za mesec
//  Zaženi ročno: izberi 'mesecniIzvoz' → Zaženi
//  Ustvari list "Izvoz-YYYY-MM" ki ga lahko preneseš kot Excel
// ════════════════════════════════════════════════════════════

function mesecniIzvoz() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const bookings = sheetToObjects(ss.getSheetByName(SHEETS.bookings));
  const slots = sheetToObjects(ss.getSheetByName(SHEETS.slots));
  const payments = sheetToObjects(ss.getSheetByName(SHEETS.payments));
  const clients = sheetToObjects(ss.getSheetByName(SHEETS.clients));

  const now = new Date();
  const mesec = now.toISOString().slice(0,7); // YYYY-MM
  const listIme = 'Izvoz-' + mesec;

  // Pobriši star list če obstaja
  const obstojeci = ss.getSheetByName(listIme);
  if (obstojeci) ss.deleteSheet(obstojeci);

  const sheet = ss.insertSheet(listIme);
  const slotMap = {}; slots.forEach(s => slotMap[s.id] = s);
  const clientMap = {}; clients.forEach(c => clientMap[c.id] = c);

  // ── REZERVACIJE ta mesec ──
  sheet.appendRow(['REZERVACIJE ZA ' + mesec]);
  sheet.appendRow(['Datum','Čas','Stranka','Email','Telefon','Status','Prisotnost']);
  const monthBookings = bookings.filter(b => {
    const slot = slotMap[b.slot_id];
    return slot && slot.datum && slot.datum.slice(0,7) === mesec;
  }).sort((a,b) => {
    const sa = slotMap[a.slot_id], sb = slotMap[b.slot_id];
    return (sa.datum+sa.cas).localeCompare(sb.datum+sb.cas);
  });
  monthBookings.forEach(b => {
    const slot = slotMap[b.slot_id];
    sheet.appendRow([slot.datum, slot.cas, b.ime, b.email, b.telefon, b.status, b.prisotnost||'']);
  });

  sheet.appendRow([]);
  sheet.appendRow(['PLAČILA ZA ' + mesec]);
  sheet.appendRow(['Stranka','Opis','Znesek (€)','Rok','Status']);
  const monthPayments = payments.filter(p => p.rok && p.rok.slice(0,7) === mesec);
  let skupajPlacano = 0;
  monthPayments.forEach(p => {
    const c = clientMap[p.client_id];
    sheet.appendRow([c?c.ime:'—', p.opis, parseFloat(p.znesek||0), p.rok, p.status]);
    if (p.status === 'placano') skupajPlacano += parseFloat(p.znesek||0);
  });
  sheet.appendRow([]);
  sheet.appendRow(['SKUPAJ PLAČANO:', skupajPlacano + ' €']);

  // Oblikovanje glave
  sheet.getRange(1,1).setFontWeight('bold').setFontSize(14);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(3, 150);

  Logger.log('✅ Izvoz ustvarjen: list "' + listIme + '". Prenesi ga preko Datoteka → Prenesi → Excel.');
  return { ok: true, list: listIme, rezervacij: monthBookings.length, placano: skupajPlacano };
}

// ════════════════════════════════════════════════════════════
//  NASTAVI STOLPEC TELEFON KOT BESEDILO
//  Zaženi ENKRAT — prepreči da Google briše začetne ničle
//  Po tem ročno popravi obstoječe številke (dodaj manjkajoče 0)
// ════════════════════════════════════════════════════════════

function telefonKotBesedilo() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.clients);
  // Stolpec C (telefon) = 3, vse vrstice
  const lastRow = Math.max(sheet.getLastRow(), 100);
  sheet.getRange(2, 3, lastRow, 1).setNumberFormat('@');
  Logger.log('✅ Stolpec telefon nastavljen kot besedilo. Zdaj ročno dodaj manjkajoče začetne ničle.');
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  MESEČNI RESET PAKETOV
//  Zaženi avtomatsko 1. v mesecu (sprožilec).
//  Resetira porabljene obiske na 0 (paket šteje od začetka).
//  Neporabljeni obiski propadejo. Pošlje opomnik za plačilo.
// ════════════════════════════════════════════════════════════

function mesecniResetPaketov() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.clients);
  const clients = sheetToObjects(sheet);

  const head = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const porabCol = head.indexOf('ure_porabljene') + 1; // stolpec porabljenih

  let resetiranih = 0;
  const seznam = [];

  clients.forEach(c => {
    // Samo stranke s paketom z omejenim številom obiskov
    if (c.paket && parseInt(c.ure_skupaj||0) > 0) {
      const row = findRowById(sheet, c.id);
      if (row > 0 && porabCol > 0) {
        sheet.getRange(row, porabCol).setValue(0); // reset porabljenih
        resetiranih++;
        seznam.push(`${c.ime} — ${c.paket}`);
      }
    }
  });

  // Obvesti Evo
  if (resetiranih > 0) {
    sendAdminEmail(`🔄 Nov mesec — ${resetiranih} paketov resetiranih`,
      emailTemplate('Mesečni reset paketov','🔄',[
        ['Resetiranih', `${resetiranih} strank`],
        ['Stranke', seznam.join('<br>')],
        ['Opomba', 'Obiski so resetirani na polno. Stranke morajo plačati nov paket za ta mesec.']
      ], 'Ne pozabi pobrati plačil za nov mesec! 🌿'));
  }

  Logger.log(`✅ Resetiranih ${resetiranih} paketov za nov mesec.`);
  return { ok: true, resetiranih };
}

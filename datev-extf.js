// datev-extf.js: camt.053 (parsed by camt053.js) -> DATEV-Format
// "Buchungsstapel" (booking batch) CSV, the EXTF import file DATEV
// Kanzlei-Rechnungswesen and most other DATEV-compatible accounting
// software read (Extras > Kanzlei-Rechnungswesen > ASCII-Schnittstelle,
// or Beleg > Import, depending on the DATEV product): a header record,
// a column-header row, then one row per booking.
//
// Pure, dependency-free, identical in Node and the browser (same
// convention as camt053.js/xlsx-writer.js/licence.js/mt940.js).
//
// Sources for the header/column structure (fetched 2026-09-06):
//  - https://developer.datev.de/de/file-format/details/datev-format/format-description/header
//    and .../format-description/booking-batch -- DATEV's own canonical
//    reference (a JS-rendered page: the two explainer pages below quote
//    its content in accessible plain HTML and are what this module's
//    field choices are checked against).
//  - https://dokuwandel.de/ratgeber/datev-buchungsstapel-erklaerung and
//    https://smartkontoauszug.de/blog/datev-extf-format-erklaert --
//    confirm the header carries the identifier "EXTF", format version
//    700, category 21 ("Buchungsstapel"), and, at fixed positions,
//    Beraternummer, Mandantennummer, WJ-Beginn (fiscal year start),
//    Sachkontenlänge (account-number length), Datum vom/bis (period) and
//    Bezeichnung (free-text label); and the booking-row column order
//    Umsatz (ohne Soll/Haben-Kz), Soll/Haben-Kennzeichen, WKZ Umsatz,
//    Kurs, Basis-Umsatz, WKZ Basis-Umsatz, Konto, Gegenkonto (ohne
//    BU-Schlüssel), BU-Schlüssel, Belegdatum (DDMM), Belegfeld 1,
//    Belegfeld 2, Skonto, Buchungstext.
//
// Honest limits (see README.md): DATEV's real header carries several
// more optional trailing fields (Diktatkürzel, Buchungstyp, Kontenrahmen,
// ...) that this module leaves off, per the format's own documented
// convention that a header line may end once every field after it would
// be empty; the exact values of a couple of metadata-only header fields
// this module does set (Formatversion of the Buchungsstapel sub-format,
// the 2-letter Herkunft origin code) are a defensible generic choice, not
// a value independently confirmed against DATEV's own source above (its
// JS-rendered page could not be fetched as plain text). None of that
// affects the booking rows themselves (amounts, dates, Konto/Gegenkonto,
// Buchungstext), which follow the confirmed column structure exactly.
// Always test-import a small file before relying on this in production.
//
// Soll/Haben-Kennzeichen: "Konto" here is always the bank account itself
// (SKR03/SKR04 1200/1800 by default). For an asset account, money coming
// IN increases the balance, which is booked to the Soll (debit) side;
// money going OUT decreases it, booked to Haben (credit). So a camt.053
// CRDT entry (credit to the bank account) is Soll-Buchung "S", and a
// DBIT entry (debit from the bank account) is Haben-Buchung "H" -- the
// same convention every DATEV bank-statement import guide describes.

// ─────────────────────────── Windows-1252 / Latin-1 text ────────────────────
// DATEV's ASCII/EXTF format is conventionally Windows-1252 (a.k.a. ANSI):
// German umlauts and ß are KEPT (unlike mt940.js's SWIFT X output, which
// transliterates them away), but anything outside the Latin-1 range
// (Windows-1252 and Unicode agree on 0x00-0xFF except the rarely-typed
// 0x80-0x9F block, which this function does not need since it maps those
// few typographic characters to a plain-ASCII equivalent instead) has no
// representation and is replaced or dropped.

const CP1252_LOOKALIKES = {
  '–': '-', '—': '-', // en dash, em dash
  '‘': "'", '’': "'", '‚': "'",
  '“': '"', '”': '"', '„': '"',
  '…': '...',
};

/** Keeps German umlauts/ß and everything else already in Latin-1
 * (code points 0x00-0xFF); maps a handful of common "smart" typography
 * to a Latin-1 equivalent; drops anything else outside that range. Never
 * throws on non-string input. */
export function toCp1252SafeText(value) {
  let s = String(value === null || value === undefined ? '' : value);
  s = s.replace(/[–—‘’‚“”„…]/g, (c) => CP1252_LOOKALIKES[c] || '');
  let out = '';
  for (const ch of s) {
    const code = ch.codePointAt(0);
    if (code <= 0xff) out += ch;
  }
  return out;
}

/** The final download step: a toCp1252SafeText()-cleaned string (every
 * character already <= 0xFF) turned into the actual single-byte
 * Windows-1252/Latin-1 file bytes. Do not feed this arbitrary Unicode:
 * run toCp1252SafeText() (toDatevBuchungsstapel()'s output already is)
 * first. */
export function toCp1252Bytes(safeString) {
  const s = String(safeString || '');
  const bytes = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i) & 0xff;
  return bytes;
}

// ─────────────────────────────── small helpers ──────────────────────────────

function clip(value, maxLen) {
  const s = toCp1252SafeText(value).trim();
  return maxLen > 0 ? s.slice(0, maxLen) : '';
}

/** DATEV EXTF quoting: text/alphanumeric fields are wrapped in double
 * quotes (embedded quotes doubled, same convention as ordinary CSV);
 * numeric/coded fields (amounts, account numbers, dates, S/H) are not. */
function q(value) {
  return '"' + String(value === null || value === undefined ? '' : value).replace(/"/g, '""') + '"';
}

function digitsOnly(value, fallback) {
  const d = String(value || '').replace(/\D/g, '');
  return d || fallback || '';
}

function yyyymmdd(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? m[1] + m[2] + m[3] : '';
}
function ddmm(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? m[3] + m[2] : '0000';
}

function amountComma(n) {
  const abs = Math.abs(Number(n) || 0);
  return abs.toFixed(2).replace('.', ',');
}

function timestamp17(date) {
  const d = date instanceof Date ? date : new Date();
  const p2 = (n) => String(n).padStart(2, '0');
  const p3 = (n) => String(n).padStart(3, '0');
  return `${d.getFullYear()}${p2(d.getMonth() + 1)}${p2(d.getDate())}${p2(d.getHours())}${p2(d.getMinutes())}${p2(d.getSeconds())}${p3(d.getMilliseconds())}`;
}

function allEntries(parsed) {
  if (!parsed || !Array.isArray(parsed.statements)) return [];
  const rows = [];
  parsed.statements.forEach((stmt) => (stmt.entries || []).forEach((e) => rows.push(e)));
  return rows;
}

// ──────────────────────────────── header record ─────────────────────────────

export const FORMAT_VERSION = 700;
export const FORMAT_CATEGORY = 21; // Buchungsstapel

function buildHeader(entries, opts) {
  const bookingDates = entries.map((e) => e.bookingDate).filter(Boolean).sort();
  const periodFrom = opts.periodFrom || bookingDates[0] || opts.now && opts.now.toISOString().slice(0, 10) || new Date().toISOString().slice(0, 10);
  const periodTo = opts.periodTo || bookingDates[bookingDates.length - 1] || periodFrom;
  const fyStart = opts.fiscalYearStart || `${periodFrom.slice(0, 4)}-01-01`;

  const fields = [
    q('EXTF'),                                        // 1  Kennzeichen
    String(FORMAT_VERSION),                           // 2  Versionsnummer
    String(FORMAT_CATEGORY),                          // 3  Formatkategorie (21 = Buchungsstapel)
    q('Buchungsstapel'),                              // 4  Formatname
    '12',                                             // 5  Formatversion (Buchungsstapel sub-format)
    timestamp17(opts.now),                            // 6  Erzeugt am
    q(''),                                            // 7  Importiert (blank on export)
    q('SV'),                                          // 8  Herkunft (generic third-party origin code)
    q(clip(opts.exportedBy || 'ARLing camt.053', 25)), // 9  Exportiert von
    q(''),                                            // 10 Importiert von (blank on export)
    digitsOnly(opts.advisorNumber, '1001'),           // 11 Beraternummer
    digitsOnly(opts.clientNumber, '1'),               // 12 Mandantennummer
    yyyymmdd(fyStart) || '20260101',                  // 13 WJ-Beginn
    digitsOnly(opts.accountLength, '4'),              // 14 Sachkontenlänge
    yyyymmdd(periodFrom),                             // 15 Datum vom
    yyyymmdd(periodTo),                               // 16 Datum bis
    q(clip(opts.designation || 'camt.053 Import', 30)), // 17 Bezeichnung
  ];
  return fields.join(';');
}

// ──────────────────────────────── column header ──────────────────────────────

const COLUMN_HEADERS = [
  'Umsatz (ohne Soll/Haben-Kz)', 'Soll/Haben-Kennzeichen', 'WKZ Umsatz', 'Kurs',
  'Basis-Umsatz', 'WKZ Basis-Umsatz', 'Konto', 'Gegenkonto (ohne BU-Schlüssel)',
  'BU-Schlüssel', 'Belegdatum', 'Belegfeld 1', 'Belegfeld 2', 'Skonto', 'Buchungstext',
];

function buildColumnHeaderRow() {
  return COLUMN_HEADERS.map((h) => q(toCp1252SafeText(h))).join(';');
}

// ──────────────────────────────── one booking row ────────────────────────────

function buildBookingText(row) {
  const parts = [row.counterpartyName, row.message].filter((s) => s && String(s).trim());
  return clip(parts.join(' ') || (row.direction === 'DBIT' ? 'Belastung' : 'Gutschrift'), 60);
}

function buildRow(row, opts) {
  const soll = row.direction === 'DBIT' ? 'H' : 'S';
  const beleg1 = clip(row.endToEndId && row.endToEndId.toUpperCase() !== 'NOTPROVIDED' ? row.endToEndId : row.bankRef, 36);
  return [
    amountComma(row.amount),                          // Umsatz
    soll,                                              // Soll/Haben-Kennzeichen
    q(clip(row.currency || 'EUR', 3)),                // WKZ Umsatz
    '',                                                // Kurs
    '',                                                // Basis-Umsatz
    q(''),                                             // WKZ Basis-Umsatz
    digitsOnly(opts.bankAccount, '1200'),             // Konto (the bank account itself)
    opts.counterAccount ? digitsOnly(opts.counterAccount, '') : '', // Gegenkonto
    '',                                                // BU-Schlüssel
    ddmm(row.bookingDate),                            // Belegdatum
    q(beleg1),                                         // Belegfeld 1
    q(''),                                             // Belegfeld 2
    '',                                                // Skonto
    q(buildBookingText(row)),                          // Buchungstext
  ].join(';');
}

// ──────────────────────────────────── main ───────────────────────────────────

/**
 * Converts a parse() result (or any {statements:[...]} object, same
 * convention as camt053.js's summarize()/toRows() and mt940.js's
 * toMt940()) into a DATEV-Format Buchungsstapel CSV: header record,
 * column-header row, one row per entry across every statement. Returns a
 * plain string already restricted to Windows-1252/Latin-1 code points
 * (see toCp1252SafeText()/toCp1252Bytes() above for turning it into the
 * actual downloadable file bytes) with CRLF line endings.
 *
 * @param {{statements: Array}} parsed
 * @param {{bankAccount?:string|number, counterAccount?:string|number,
 *   accountLength?:string|number, advisorNumber?:string|number,
 *   clientNumber?:string|number, fiscalYearStart?:string,
 *   periodFrom?:string, periodTo?:string, designation?:string,
 *   exportedBy?:string, now?:Date}} [options]
 */
export function toDatevBuchungsstapel(parsed, options) {
  const opts = options && typeof options === 'object' ? options : {};
  const entries = allEntries(parsed).filter((e) => e.amount !== null && e.amount !== undefined);
  const lines = [buildHeader(entries, opts), buildColumnHeaderRow()];
  entries.forEach((row) => lines.push(buildRow(row, opts)));
  return lines.join('\r\n') + '\r\n';
}

const DatevExtfConverter = { toDatevBuchungsstapel, toCp1252SafeText, toCp1252Bytes, FORMAT_VERSION, FORMAT_CATEGORY };

if (typeof window !== 'undefined') {
  window.DatevExtfConverter = DatevExtfConverter;
}

// mt940.js: camt.053 (parsed by camt053.js) -> SWIFT MT940 bank statement
// message, so a camt.053 XML file can still feed a German MT940-only
// import (DATEV Kanzlei-Rechnungswesen's own manual file import still
// expects MT940, not camt.053, even after German banks stopped issuing
// MT940 in November 2025: see README.md and the accompanying research
// notes for sources) or any other legacy accounting software that never
// added camt.053 support.
//
// Pure, dependency-free, works identically in Node and the browser (same
// approach as camt053.js/xlsx-writer.js/licence.js: no DOMParser, no
// Buffer-only API). Reads the {statements:[...]} shape parse() returns
// (or any hand-built object with the same shape, e.g. a multi-file merge:
// same convention as camt053.js's own summarize()/toRows()) and returns
// one SWIFT MT940 message per statement, concatenated into one file
// (multiple MT940 messages in one .sta file is standard practice: each
// message is self-contained, terminated by a line with just "-", and the
// next message's :20: follows immediately).
//
// Sources for the MT940 field structure and the German :86: subfield
// convention (fetched 2026-09-06, both describe the same Deutsche
// Kreditwirtschaft / ZKA "Datenformate für den Kontoauszug" convention
// that every German MT940-issuing bank, e.g. Sparkasse/Volksbank, used
// before the November 2025 camt.053 switch):
//  - https://www.kontopruef.de/mt940s.shtml -- field-by-field breakdown
//    of :20:, :25:, :28C:, :60F/M:, :61: (subfields: Valuta YYMMDD,
//    Buchungsdatum MMDD, Soll/Haben-Kennung C/D/RC/RD, Betrag with comma,
//    Buchungsschlüssel 4 chars starting with N, Kundenreferenz or
//    NONREF, "//" + Bankreferenz), :86: (GVC 3-digit code, then ?00
//    Buchungstext, ?10 Primanota, ?20-?29 Verwendungszweck (10 x 27
//    chars), ?30 Bankkennung, ?31 Kontonummer, ?32-?33 Name, ?34
//    Textschlüsselergänzung; max 390 characters over 6 lines of 65),
//    :62F/M:, :64:/:65:.
//  - https://www.hettwer-beratung.de/sepa-spezialwissen/sepa-technische-anforderungen/sepa-gesch%C3%A4ftsvorfallcodes-gvc-mt-940/
//    (and the DK "SEPA-Begleittext zu Feld 86" it documents) -- the SEPA
//    reference tags placed inside the ?20-?29 Verwendungszweck lines:
//    EREF+ (EndToEndId), KREF+, MREF+ (mandate reference), CRED+
//    (creditor scheme id), DEBT+, then SVWZ+ (the actual Verwendungszweck
//    free text), in that order, each tag starting a fresh ?2x subfield
//    and continuing into the next one (without repeating the tag) when
//    its content runs past 27 characters.
// The SWIFT transaction-type codes this file maps to (NTRF = Transfer,
// NDDT = Direct Debit item, NMSC = Miscellaneous, NCHG = Charges) are the
// generic SWIFT MT940 Buchungsschlüssel values; German banks additionally
// append the 3-digit DK GVC after a "+" inside the camt.053 proprietary
// code (Issr "DK"), e.g. "NTRF+166": this project's own
// SAMPLE_CAMT053_XML_DE (camt053.js) already carries exactly that
// convention (NTRF+166, NDDT+105, NMSC+106, NCHG+808), which this file's
// mapTransactionCode() reads directly when present, falling back to a
// mapping from the ISO 20022 BkTxCd family/sub-family code otherwise.
//
// Honest limits (see README.md): the exact wording of ?00 Buchungstext,
// which SEPA tags are populated (only EREF+/SVWZ+ here: camt053.js's row
// shape does not carry a mandate id or creditor scheme id, so MREF+/
// CRED+/DEBT+ are only ever emitted if a future engine change adds those
// fields) and the GVC-to-Buchungstext wording all vary bank to bank in
// real MT940 exports; this module produces a well-formed, honestly
// generic MT940 message from the fields camt053.js actually extracts,
// not a byte-for-byte replica of any one bank's own export.

// ─────────────────────────── SWIFT "X" character set ───────────────────────
// Allowed: A-Z a-z 0-9 space / - ? : ( ) . , ' + and CR LF. Everything else
// (umlauts, other diacritics, other punctuation) must be transliterated or
// stripped before it goes into an MT940 field.

const UMLAUT_MAP = {
  'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'Ä': 'Ae', 'Ö': 'Oe', 'Ü': 'Ue', 'ß': 'ss',
};

/** Transliterates German umlauts/ß and folds other accented Latin letters
 * to their base form (NFD + strip combining marks). Never throws on
 * non-string input. Split out from transliterateSwiftX() below so callers
 * that need umlaut folding but not the full SWIFT X charset restriction
 * (e.g. index.html's download-filename sanitizer, which wants to keep
 * "_" and other characters SWIFT's X charset does not allow) can use just
 * this step. */
export function foldDiacritics(value) {
  let s = String(value === null || value === undefined ? '' : value);
  s = s.replace(/[äöüÄÖÜß]/g, (c) => UMLAUT_MAP[c] || c);
  try { s = s.normalize('NFD').replace(/[̀-ͯ]/g, ''); } catch (e) { /* normalize always exists in Node20+/modern browsers */ }
  return s;
}

/** Transliterates umlauts/ß, folds other accented Latin letters to their
 * base form, then drops anything still outside the SWIFT X character set.
 * Never throws on non-string input. */
export function transliterateSwiftX(value) {
  let s = foldDiacritics(value);
  // eslint-disable-next-line no-control-regex
  s = s.replace(/[^A-Za-z0-9 \/\-\?:\(\)\.,'+\r\n]/g, '');
  // A stripped character (e.g. "&", an em dash) often leaves a double
  // space behind ("Müller & Söhne" -> "Mueller  Soehne"): collapse runs
  // of plain spaces/tabs, but never touch the CRLF this function also
  // has to let through untouched (field 86 wraps across lines).
  s = s.replace(/[ \t]{2,}/g, ' ');
  return s;
}

function clip(value, maxLen) {
  const s = transliterateSwiftX(value).trim();
  return maxLen > 0 ? s.slice(0, maxLen) : '';
}

// ─────────────────────────────── date / amount ──────────────────────────────

function yymmdd(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? m[1].slice(2) + m[2] + m[3] : '000000';
}
function mmdd(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? m[2] + m[3] : '0000';
}

/** Absolute amount, 2 decimals, comma as the decimal mark (SWIFT amount
 * subfield convention: "1234,56", never a "." and no thousands grouping). */
function amountComma(n) {
  const abs = Math.abs(Number(n) || 0);
  return abs.toFixed(2).replace('.', ',');
}

// ─────────────────────────────── balances ───────────────────────────────────
// Same OPBD/PRCD (opening) and CLBD/CLAV (closing) type codes camt053.js's
// own summarizeStatement()/pickBalance() use, reimplemented locally so this
// file has no dependency beyond the plain {statements:[...]} data shape.

function pickBalance(balances, codes) {
  const list = Array.isArray(balances) ? balances : [];
  for (const code of codes) {
    const found = list.find((b) => b && b.type === code);
    if (found) return found;
  }
  return null;
}

function balanceField(tag, bal, fallbackDate, fallbackCcy) {
  const amount = bal && typeof bal.amount === 'number' ? bal.amount : 0;
  const mark = amount < 0 ? 'D' : 'C';
  const date = yymmdd((bal && bal.date) || fallbackDate);
  const ccy = clip((bal && bal.currency) || fallbackCcy || 'EUR', 3).toUpperCase() || 'EUR';
  return `${tag}${mark}${date}${ccy}${amountComma(amount)}`;
}

// ───────────────────────── transaction type mapping ─────────────────────────
// See the header comment: prefer the DK-issued "<SWIFT code>+<3-digit GVC>"
// proprietary code camt053.js already folds into a row's txType (e.g.
// "PMNT-RCDT-ESCT / NTRF+166 (DK)"); otherwise fall back to a mapping from
// the ISO 20022 structured family/sub-family code that formatTxType() also
// puts into txType (e.g. "PMNT-RDDT-ESDD").

function mapTransactionCode(row) {
  const t = String((row && row.txType) || '');
  const dk = /\b(N[A-Z]{3})\+(\d{3})\b/.exec(t);
  if (dk) return { code: dk[1], gvc: dk[2] };
  if (/RDDT|ESDD|DDT/.test(t)) return { code: 'NDDT', gvc: '105' };
  if (/CCRD/.test(t)) return { code: 'NMSC', gvc: '106' };
  if (/CHRG|MDOP/.test(t)) return { code: 'NCHG', gvc: '808' };
  if (/RCDT|ICDT/.test(t)) return { code: 'NTRF', gvc: '166' };
  return { code: 'NMSC', gvc: '999' };
}

function isReversal(row) {
  return /RVSL/.test(String((row && row.txType) || ''));
}

// ──────────────────────────────── field :61: ────────────────────────────────
// Concatenated per the SWIFT spec, no separators between the fixed-length
// subfields (the receiving parser knows the boundaries): Valuta(6) +
// Buchungsdatum(4) + Soll/Haben-Kennung(1-2) + Betrag(var, digits+comma) +
// Buchungsschlüssel(4) + Kundenreferenz(var, <=16) + "//" + Bankreferenz
// (var, <=16, only when present). Hard-clipped so the whole ":61:..." line
// never exceeds 65 characters even for pathological input.

function field61(row) {
  const dir = row.direction === 'DBIT' ? 'D' : 'C';
  const mark = isReversal(row) ? 'R' + dir : dir;
  const { code } = mapTransactionCode(row);
  const prefix = yymmdd(row.valueDate) + mmdd(row.bookingDate) + mark + amountComma(row.amount) + code;

  const rawRef = String(row.endToEndId || '').trim();
  const hasRef = rawRef && rawRef.toUpperCase() !== 'NOTPROVIDED' && rawRef.toUpperCase() !== 'NONREF';
  let customerRef = hasRef ? clip(rawRef, 16) : 'NONREF';
  let bankRef = clip(row.bankRef, 16);

  // Budget check against the 65-char line limit (tag ":61:" = 4 chars);
  // trim bankRef first (less valuable for reconciliation than the payment
  // reference), then customerRef, rather than ever emitting a line >65.
  const fixedLen = 4 + prefix.length;
  let tail = customerRef + (bankRef ? '//' + bankRef : '');
  while (fixedLen + tail.length > 65 && bankRef) {
    bankRef = bankRef.slice(0, -1);
    tail = customerRef + (bankRef ? '//' + bankRef : '');
  }
  while (fixedLen + tail.length > 65 && customerRef.length > 1) {
    customerRef = customerRef.slice(0, -1);
    tail = customerRef + (bankRef ? '//' + bankRef : '');
  }
  return `:61:${prefix}${tail}`;
}

// ──────────────────────────────── field :86: ────────────────────────────────

function purposeLines(row) {
  const lines = [];
  const rawRef = String(row.endToEndId || '').trim();
  const hasRef = rawRef && rawRef.toUpperCase() !== 'NOTPROVIDED' && rawRef.toUpperCase() !== 'NONREF';
  if (hasRef) pushTagged(lines, 'EREF+', rawRef);
  if (row.message) pushTagged(lines, 'SVWZ+', row.message);
  return lines;
}

// One tag (e.g. "EREF+") starts a fresh ?2x subfield; if its content runs
// past 27 characters it continues into the following ?2x subfield(s)
// without repeating the tag, exactly as the DK SEPA-Begleittext convention
// (cited in the header comment) documents.
function pushTagged(lines, tag, value) {
  const text = clip(tag + value, 100000); // transliterate once, no length cap yet
  let rest = text;
  let first = true;
  while (rest.length > 0) {
    const budget = 27;
    lines.push(rest.slice(0, budget));
    rest = rest.slice(budget);
    first = false;
  }
  if (first) lines.push(tag.slice(0, 27));
}

function buildField86(row) {
  const { gvc } = mapTransactionCode(row);
  let content = gvc;

  const buchungstext = clip(row.counterpartyName || (row.direction === 'DBIT' ? 'Belastung' : 'Gutschrift'), 27);
  content += '?00' + buchungstext;

  const primanota = clip(row.bankRef, 10);
  if (primanota) content += '?10' + primanota;

  purposeLines(row).slice(0, 10).forEach((line, i) => {
    content += '?' + String(20 + i).padStart(2, '0') + line;
  });

  if (row.counterpartyBic) content += '?30' + clip(row.counterpartyBic, 12);
  if (row.counterpartyIban) content += '?31' + clip(row.counterpartyIban, 24);
  if (row.counterpartyName) {
    const name = clip(row.counterpartyName, 54); // 2 x 27
    content += '?32' + name.slice(0, 27);
    if (name.length > 27) content += '?33' + name.slice(27, 54);
  }

  // Wrap into up to 6 lines of 65 chars; the first line carries the ":86:"
  // tag (4 chars), so its content budget is 61, continuation lines get the
  // full 65 (no tag repeated, per SWIFT multi-line field convention).
  // Overflow beyond 6 lines is dropped (documented in README.md): the same
  // truncation a real bank applies when a Verwendungszweck runs too long.
  const lines = [];
  let rest = content;
  let budget = 65 - 4;
  while (rest.length > 0 && lines.length < 6) {
    lines.push(rest.slice(0, budget));
    rest = rest.slice(budget);
    budget = 65;
  }
  return ':86:' + lines.join('\r\n');
}

// ────────────────────────────── one message ─────────────────────────────────

/** One full MT940 message (block of SWIFT fields) for a single statement,
 * ending with a line containing only "-". `opts.reference` overrides the
 * default ":20:" reference (real Sparkasse/Volksbank MT940 deliveries
 * commonly use a fixed generic value here, not a per-statement one; the
 * cited kontopruef.de source's own worked example uses exactly this kind
 * of constant marker, "STARTUMSE"). */
export function toMt940Statement(stmt, opts) {
  const o = opts && typeof opts === 'object' ? opts : {};
  const reference = clip(o.reference || 'STARTUMSE', 16) || 'STARTUMSE';
  const account = clip((stmt.account && (stmt.account.iban || stmt.account.otherId)) || '', 35);
  const stmtNo = clip((stmt.legalSeqNb || stmt.elctrncSeqNb || '1').replace(/\D/g, '') || '1', 5) || '1';
  const ccy = (stmt.account && stmt.account.currency) || 'EUR';

  const opening = pickBalance(stmt.balances, ['OPBD', 'PRCD']);
  const closing = pickBalance(stmt.balances, ['CLBD', 'CLAV']);
  const available = pickBalance(stmt.balances, ['CLAV']);

  const lines = [];
  lines.push(':20:' + reference);
  lines.push(':25:' + account);
  lines.push(':28C:' + stmtNo);
  lines.push(balanceField(':60F:', opening, stmt.fromDateTime, ccy));

  (stmt.entries || []).forEach((row) => {
    if (row.amount === null || row.amount === undefined) return;
    lines.push(field61(row));
    lines.push(buildField86(row));
  });

  lines.push(balanceField(':62F:', closing, stmt.toDateTime, ccy));
  // Optional :64: (current valuta balance): only emitted when the
  // statement carries a distinct CLAV figure (differs from CLBD); most
  // camt.053 exports (including both this project's own sample files)
  // report only OPBD/CLBD, so this line is typically absent.
  if (available && closing && Math.abs((available.amount || 0) - (closing.amount || 0)) > 0.004) {
    lines.push(balanceField(':64:', available, stmt.toDateTime, ccy));
  }
  lines.push('-');
  return lines.join('\r\n');
}

// ──────────────────────────────────── file ───────────────────────────────────

/**
 * Converts a parse() result (or any {statements:[...]} object, e.g. a
 * hand-merged multi-file upload: same convention as camt053.js's own
 * summarize()/toRows()) into MT940 text: one SWIFT message per statement,
 * concatenated (a multi-statement .sta file is ordinary multiple MT940
 * messages back to back). CRLF line endings throughout, terminated with a
 * trailing CRLF. Returns '' for anything without at least one statement.
 */
export function toMt940(parsed, opts) {
  if (!parsed || !Array.isArray(parsed.statements) || parsed.statements.length === 0) return '';
  return parsed.statements.map((stmt) => toMt940Statement(stmt, opts)).join('\r\n') + '\r\n';
}

const Mt940Converter = { toMt940, toMt940Statement, transliterateSwiftX, foldDiacritics };

if (typeof window !== 'undefined') {
  window.Mt940Converter = Mt940Converter;
}

export { pickBalance as _pickBalanceForTest, mapTransactionCode as _mapTransactionCodeForTest };

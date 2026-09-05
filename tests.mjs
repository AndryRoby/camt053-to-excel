// tests.mjs — plain Node test runner for camt053.js and xlsx-writer.js
// (no external dependencies). Run with: node tests.mjs

import { parse, toRows, toCsv, summarize, COLUMNS, SAMPLE_CAMT053_XML, SAMPLE_CAMT053_XML_DE, bankFromBic, parseXml } from './camt053.js';
import { buildXlsx, buildZip, crc32, colLetter } from './xlsx-writer.js';
import { parse as parseLicence, verify as verifyLicence, isValid as isValidLicence, load as loadLicence, save as saveLicence, clear as clearLicence, todayIso as licenceTodayIso, STORAGE_KEY as LICENCE_STORAGE_KEY, DEFAULT_PLAN } from './licence.js';
import {
  LANGS, DEFAULT_LANG, DICT, COLUMN_LABELS, t, tf, columnLabel, columnLabelsMap,
  formatAmountForLang, formatDateForLang, defaultCsvOptsForLang, localeTagForLang,
  ogLocaleForLang, langFromLocale, langFromQueryString, findIncompleteEntries,
  TEMPLATE_PRESETS, templateLabel, templateOrderForLang,
  DEFAULT_COLUMNS_BY_LANG, defaultColumnsForLang,
} from './i18n.js';

// Minimal in-memory localStorage polyfill: Node has no Web Storage API by
// default, and licence.js is meant to degrade to a no-op when it's
// absent — so the load/save/clear round-trip test below needs one
// installed, exactly like a real browser tab would provide.
if (typeof globalThis.localStorage === 'undefined') {
  const store = new Map();
  globalThis.localStorage = {
    getItem: (k) => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => { store.set(String(k), String(v)); },
    removeItem: (k) => { store.delete(String(k)); },
    clear: () => { store.clear(); },
  };
}

let pass = 0;
let fail = 0;
const failures = [];

function ok(name, cond, detail) {
  if (cond) { pass++; } else { fail++; failures.push(`${name}${detail ? ' — ' + detail : ''}`); }
}
function eq(name, actual, expected) {
  const cond = actual === expected;
  ok(name, cond, cond ? '' : `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}
function deepEq(name, actual, expected) {
  const cond = JSON.stringify(actual) === JSON.stringify(expected);
  ok(name, cond, cond ? '' : `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}
function includes(name, haystack, needle) {
  const cond = typeof haystack === 'string' && haystack.includes(needle);
  ok(name, cond, cond ? '' : `expected string to include ${JSON.stringify(needle)}`);
}
function close(name, actual, expected, eps) {
  const cond = typeof actual === 'number' && Math.abs(actual - expected) <= (eps || 0.005);
  ok(name, cond, cond ? '' : `expected ~${expected}, got ${actual}`);
}

// ═══════════════════════════ fixtures ═══════════════════════════════════

// camt.053.001.08 fixture, two statements, built to exercise the paths the
// bundled .02 SAMPLE_CAMT053_XML does not: a bare Ntry with no NtryDtls at
// all, RmtInf/Strd/CdtrRefInf/Ref-sourced VS, a batched Ntry whose two
// TxDtls each carry their own Amt/CdtDbtInd, an Amt with no Ccy attribute
// (falls back to the account currency), an entity in free text, and a
// second statement whose stated closing balance does not reconcile.
const SAMPLE_08_XML = `<?xml version="1.0" encoding="UTF-8"?>
<Document xmlns="urn:iso:std:iso:20022:tech:xsd:camt.053.001.08">
  <BkToCstmrStmt>
    <GrpHdr><MsgId>ARL-TEST-08-1</MsgId><CreDtTm>2026-09-06T09:00:00+02:00</CreDtTm></GrpHdr>
    <Stmt>
      <Id>SK1111000000000099999999-001-260906</Id>
      <LglSeqNb>1</LglSeqNb>
      <Acct><Id><IBAN>SK1111000000000099999999</IBAN></Id><Ccy>EUR</Ccy></Acct>
      <Bal><Tp><CdOrPrtry><Cd>OPBD</Cd></CdOrPrtry></Tp><Amt Ccy="EUR">500.00</Amt><CdtDbtInd>CRDT</CdtDbtInd><Dt><Dt>2026-09-01</Dt></Dt></Bal>
      <Bal><Tp><CdOrPrtry><Cd>CLBD</Cd></CdOrPrtry></Tp><Amt Ccy="EUR">574.50</Amt><CdtDbtInd>CRDT</CdtDbtInd><Dt><Dt>2026-09-05</Dt></Dt></Bal>
      <Ntry>
        <NtryRef>N1</NtryRef>
        <Amt>200.00</Amt>
        <CdtDbtInd>CRDT</CdtDbtInd>
        <Sts>BOOK</Sts>
        <BookgDt><Dt>2026-09-02</Dt></BookgDt>
        <ValDt><Dt>2026-09-02</Dt></ValDt>
      </Ntry>
      <Ntry>
        <NtryRef>N2</NtryRef>
        <Amt Ccy="EUR">50.00</Amt>
        <CdtDbtInd>DBIT</CdtDbtInd>
        <Sts>BOOK</Sts>
        <BookgDt><Dt>2026-09-03</Dt></BookgDt>
        <ValDt><Dt>2026-09-03</Dt></ValDt>
        <NtryDtls>
          <TxDtls>
            <Refs><AcctSvcrRef>N2-ref</AcctSvcrRef></Refs>
            <RltdPties><Cdtr><Nm>Dodavatel Structured s. r. o.</Nm></Cdtr><CdtrAcct><Id><IBAN>SK2222000000000088888888</IBAN></Id></CdtrAcct></RltdPties>
            <RmtInf><Strd><CdtrRefInf><Ref>445566</Ref></CdtrRefInf></Strd></RmtInf>
          </TxDtls>
        </NtryDtls>
      </Ntry>
      <Ntry>
        <NtryRef>N3</NtryRef>
        <Amt Ccy="EUR">75.50</Amt>
        <CdtDbtInd>DBIT</CdtDbtInd>
        <Sts>BOOK</Sts>
        <BookgDt><Dt>2026-09-04</Dt></BookgDt>
        <ValDt><Dt>2026-09-04</Dt></ValDt>
        <NtryDtls>
          <TxDtls>
            <Amt>30.00</Amt>
            <CdtDbtInd>DBIT</CdtDbtInd>
            <Refs><EndToEndId>/VS1001/SS02/KS0308</EndToEndId></Refs>
            <RltdPties><Cdtr><Nm>Prvy dodavatel</Nm></Cdtr></RltdPties>
            <RmtInf><Ustrd>Faktura A &amp; B</Ustrd></RmtInf>
          </TxDtls>
          <TxDtls>
            <Amt>45.50</Amt>
            <CdtDbtInd>DBIT</CdtDbtInd>
            <Refs><EndToEndId>/VS1002/SS03/KS0308</EndToEndId></Refs>
            <RltdPties><Cdtr><Nm>Druhy dodavatel</Nm></Cdtr></RltdPties>
            <RmtInf><Ustrd>Faktura C</Ustrd></RmtInf>
          </TxDtls>
        </NtryDtls>
      </Ntry>
    </Stmt>
    <Stmt>
      <Id>SK3333000000000077777777-001-260906</Id>
      <LglSeqNb>1</LglSeqNb>
      <Acct><Id><IBAN>SK3333000000000077777777</IBAN></Id><Ccy>EUR</Ccy></Acct>
      <Bal><Tp><CdOrPrtry><Cd>OPBD</Cd></CdOrPrtry></Tp><Amt Ccy="EUR">100.00</Amt><CdtDbtInd>CRDT</CdtDbtInd><Dt><Dt>2026-09-01</Dt></Dt></Bal>
      <Bal><Tp><CdOrPrtry><Cd>CLBD</Cd></CdOrPrtry></Tp><Amt Ccy="EUR">999.00</Amt><CdtDbtInd>CRDT</CdtDbtInd><Dt><Dt>2026-09-05</Dt></Dt></Bal>
      <Ntry>
        <NtryRef>N4</NtryRef>
        <Amt Ccy="EUR">10.00</Amt>
        <CdtDbtInd>CRDT</CdtDbtInd>
        <Sts>BOOK</Sts>
        <BookgDt><Dt>2026-09-02</Dt></BookgDt>
        <ValDt><Dt>2026-09-02</Dt></ValDt>
      </Ntry>
    </Stmt>
  </BkToCstmrStmt>
</Document>
`;

const NOT_CAMT_XML = `<?xml version="1.0"?><Document xmlns="urn:pain.001"><CstmrCdtTrfInitn><GrpHdr><MsgId>X</MsgId></GrpHdr></CstmrCdtTrfInitn></Document>`;

const MALFORMED_XML = `<Document xmlns="urn:iso:std:iso:20022:tech:xsd:camt.053.001.02"><BkToCstmrStmt><GrpHdr><MsgId>X</MsgId></GrpHdr><Stmt><Id>A</Id>`;

const GARBAGE = 'toto vobec nie je XML subor, len obycajny text';

// ═══════════════════════════ parse(): .02 sample ═════════════════════════

const p02 = parse(SAMPLE_CAMT053_XML);
ok('parse .02: ok is true', p02.ok === true);
eq('parse .02: not malformed', p02.malformed, false);
eq('parse .02: version detected as 001.02', p02.version, '001.02');
includes('parse .02: namespace captured', p02.namespace, 'camt.053.001.02');
eq('parse .02: one statement', p02.statements.length, 1);
eq('parse .02: MsgId read', p02.groupHeader.msgId, 'ARL-20260906-000001');
eq('parse .02: bank detected from Svcr BIC', p02.bank, 'tatrabanka');

const stmt02 = p02.statements[0];
eq('parse .02: statement account IBAN', stmt02.account.iban, 'SK6809000000000012345678');
eq('parse .02: statement account currency', stmt02.account.currency, 'EUR');
eq('parse .02: statement id', stmt02.id, 'SK6809000000000012345678-001-260906');
eq('parse .02: from/to datetime read', stmt02.fromDateTime, '2026-09-01T00:00:00+02:00');
eq('parse .02: 3 entries parsed', stmt02.entries.length, 3);
eq('parse .02: 2 balances parsed', stmt02.balances.length, 2);

const rows02 = toRows(p02);
eq('toRows .02: 3 rows total', rows02.length, 3);
eq('toRows .02: first row signed amount is positive (CRDT)', rows02[0].amount, 450);
eq('toRows .02: second row signed amount is negative (DBIT)', rows02[1].amount, -89.9);
eq('toRows .02: third row signed amount is negative (DBIT)', rows02[2].amount, -120.8);
eq('toRows .02: currency read from Amt Ccy attr', rows02[0].currency, 'EUR');
eq('toRows .02: txType combines Domn/Fmly/SubFmlyCd', rows02[0].txType, 'PMNT-RCDT-ESCT');
eq('toRows .02: counterparty is Dbtr for a CRDT entry', rows02[0].counterpartyName, 'Jozef Odberatel');
eq('toRows .02: counterparty is Cdtr for a DBIT entry', rows02[1].counterpartyName, 'Dodavatel Novak s. r. o.');
eq('toRows .02: bankRef prefers AcctSvcrRef', rows02[0].bankRef, '2026090100001');

// ── VS/SS/KS: three extraction sources ──────────────────────────────────
eq('VS source 1/3: from EndToEndId, VS', rows02[0].vs, '2026001');
eq('VS source 1/3: from EndToEndId, SS keeps leading zeros', rows02[0].ss, '0000');
eq('VS source 1/3: from EndToEndId, KS keeps leading zeros', rows02[0].ks, '0308');
eq('VS source 1/3: vsSource tagged "endtoend"', rows02[0].vsSource, 'endtoend');
eq('VS source 1/3: entry 2 VS with no SS present', rows02[1].vs, '789');
eq('VS source 1/3: entry 2 SS is empty when absent', rows02[1].ss, '');
eq('VS source 3/3: from RmtInf/Ustrd free text "VS: 445566"', rows02[2].vs, '445566');
eq('VS source 3/3: vsSource tagged "ustrd"', rows02[2].vsSource, 'ustrd');

// ═══════════════════════════ parse(): German .02 sample (de/en button) ═══
// SAMPLE_CAMT053_XML_DE is what the sample button loads for German and
// English visitors: a Sparkasse/Volksbank-style statement for a fictional
// Muster GmbH, four entries. These assertions pin down everything the
// table shows for it, that it reconciles, that the fictional IBANs are
// nonetheless mod-97 valid, and that umlauts survive parse -> CSV -> XLSX.

// ISO 13616 mod-97 check, written out here on purpose (the engine has no
// IBAN validator of its own; this guards the sample, not the engine).
function ibanMod97Ok(iban) {
  const s = String(iban || '').replace(/\s+/g, '').toUpperCase();
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/.test(s)) return false;
  const rearranged = s.slice(4) + s.slice(0, 4);
  const digits = rearranged.replace(/[A-Z]/g, (c) => String(c.charCodeAt(0) - 55));
  let r = 0;
  for (const ch of digits) r = (r * 10 + Number(ch)) % 97;
  return r === 1;
}
eq('ibanMod97Ok: known-good IBAN passes (sanity check of the checker itself)', ibanMod97Ok('DE89370400440532013000'), true);
eq('ibanMod97Ok: one digit changed fails', ibanMod97Ok('DE89370400440532013001'), false);

const pDe = parse(SAMPLE_CAMT053_XML_DE);
ok('DE sample: parses without error (ok true)', pDe.ok === true, pDe.errors.join('; '));
eq('DE sample: not malformed', pDe.malformed, false);
deepEq('DE sample: no parser errors at all', pDe.errors, []);
eq('DE sample: version detected as 001.02', pDe.version, '001.02');
eq('DE sample: exactly one statement', pDe.statements.length, 1);
eq('DE sample: MsgId read', pDe.groupHeader.msgId, 'camt053-20260912-000037');
eq('DE sample: fictional BIC MUSTDEFFXXX is not a Slovak bank -> "iná" (rendered as "andere"/"other")', pDe.bank, 'iná');

const stmtDe = pDe.statements[0];
eq('DE sample: statement id', stmtDe.id, 'KA-2026-037');
eq('DE sample: electronic sequence number', stmtDe.elctrncSeqNb, '37');
eq('DE sample: account owner is Muster GmbH', stmtDe.account.ownerName, 'Muster GmbH');
eq('DE sample: account currency EUR', stmtDe.account.currency, 'EUR');
eq('DE sample: account IBAN', stmtDe.account.iban, 'DE40123456780000123456');
eq('DE sample: servicer BIC read from Acct/Svcr (camt.053.001.02 schema placement, not only Stmt/Svcr)', stmtDe.servicerBic, 'MUSTDEFFXXX');
eq('DE sample: statement period starts in 2026', stmtDe.fromDateTime.slice(0, 4), '2026');
eq('DE sample: statement period ends in 2026', stmtDe.toDateTime.slice(0, 4), '2026');
eq('DE sample: 4 entries', stmtDe.entries.length, 4);
eq('DE sample: 2 balances (OPBD, CLBD)', stmtDe.balances.length, 2);

const rowsDe = toRows(pDe);
eq('DE rows: 4 rows total', rowsDe.length, 4);
deepEq('DE rows: directions CRDT, DBIT, DBIT, DBIT', rowsDe.map((r) => r.direction), ['CRDT', 'DBIT', 'DBIT', 'DBIT']);
deepEq('DE rows: signed amounts', rowsDe.map((r) => r.amount), [1785, -214, -86.37, -12.9]);
deepEq('DE rows: German counterparty names (Dbtr on the credit, Cdtr on the debits)', rowsDe.map((r) => r.counterpartyName), ['Beispiel Handels GmbH', 'Stadtwerke Musterstadt GmbH', 'Bürobedarf Musterstadt', 'Musterbank']);
ok('DE rows: every row has a booking date in 2026', rowsDe.every((r) => /^2026-\d{2}-\d{2}$/.test(r.bookingDate)));
ok('DE rows: every row has a value date in 2026', rowsDe.every((r) => /^2026-\d{2}-\d{2}$/.test(r.valueDate)));
ok('DE rows: every row has currency EUR', rowsDe.every((r) => r.currency === 'EUR'));
ok('DE rows: every row has status BOOK', rowsDe.every((r) => r.status === 'BOOK'));
ok('DE rows: every row has a bank reference', rowsDe.every((r) => r.bankRef.length > 0));
ok('DE rows: every row has a counterparty IBAN', rowsDe.every((r) => r.counterpartyIban.startsWith('DE')));
ok('DE rows: every row has a counterparty BIC', rowsDe.every((r) => /^[A-Z]{6}[A-Z0-9]{2}(XXX)?$/.test(r.counterpartyBic)));
ok('DE rows: every row has an EndToEndId', rowsDe.every((r) => r.endToEndId.length > 0));
ok('DE rows: every row has a Verwendungszweck (RmtInf/Ustrd)', rowsDe.every((r) => r.message.length > 0));
ok('DE rows: every row has a bank transaction code', rowsDe.every((r) => r.txType.length > 0));
eq('DE rows: received SEPA-Überweisung is PMNT-RCDT-ESCT with the DK proprietary code', rowsDe[0].txType, 'PMNT-RCDT-ESCT / NTRF+166 (DK)');
eq('DE rows: paid SEPA-Lastschrift is PMNT-RDDT-ESDD (received direct debit, payer side)', rowsDe[1].txType, 'PMNT-RDDT-ESDD / NDDT+105 (DK)');
eq('DE rows: card payment is PMNT-CCRD-POSD', rowsDe[2].txType, 'PMNT-CCRD-POSD / NMSC+106 (DK)');
eq('DE rows: bank fee is ACMT-MDOP-CHRG', rowsDe[3].txType, 'ACMT-MDOP-CHRG / NCHG+808 (DK)');
eq('DE rows: invoice EndToEndId', rowsDe[0].endToEndId, 'RE-2026-0417');
includes('DE rows: Verwendungszweck carries the Rechnungsnummer', rowsDe[0].message, 'Rechnung RE-2026-0417');
includes('DE rows: Verwendungszweck carries the Kundenreferenz', rowsDe[0].message, 'Kundenreferenz KD-10245');
includes('DE rows: direct debit Verwendungszweck is the Stromabschlag', rowsDe[1].message, 'Stromabschlag');
includes('DE rows: direct debit Verwendungszweck names the mandate', rowsDe[1].message, 'Mandat M-2024-0042');
includes('DE rows: umlaut in counterparty name survives parse ("Bürobedarf")', rowsDe[2].counterpartyName, 'ü');
includes('DE rows: umlaut in Verwendungszweck survives parse ("Kontoführungsentgelt")', rowsDe[3].message, 'Kontoführungsentgelt');
includes('DE sample XML: direct debit carries a MndtId', SAMPLE_CAMT053_XML_DE, '<MndtId>M-2024-0042</MndtId>');
includes('DE sample XML: direct debit carries a creditor identifier in the DE..ZZZ.. pattern', SAMPLE_CAMT053_XML_DE, '<Id>DE98ZZZ09999999999</Id>');
ok('DE sample XML: creditor identifier matches the German Gläubiger-ID shape', /DE\d{2}ZZZ\d{11}/.test(SAMPLE_CAMT053_XML_DE));
ok('DE rows: no Slovak VS/SS/KS false positives from the German free text', rowsDe.every((r) => r.vs === '' && r.ss === '' && r.ks === '' && r.vsSource === ''));

// IBANs: account, every counterparty, and every IBAN literally present in
// the XML must pass mod-97.
eq('DE IBAN: account IBAN passes mod-97', ibanMod97Ok(stmtDe.account.iban), true);
ok('DE IBAN: every counterparty IBAN passes mod-97', rowsDe.every((r) => ibanMod97Ok(r.counterpartyIban)), rowsDe.map((r) => r.counterpartyIban).join(','));
{
  const allIbans = Array.from(new Set((SAMPLE_CAMT053_XML_DE.match(/<IBAN>([^<]+)<\/IBAN>/g) || []).map((m) => m.replace(/<\/?IBAN>/g, ''))));
  ok('DE IBAN: at least 5 distinct IBANs in the sample (account + 4 counterparties)', allIbans.length >= 5, String(allIbans.length));
  ok('DE IBAN: every <IBAN> in the XML passes mod-97', allIbans.every(ibanMod97Ok), allIbans.filter((i) => !ibanMod97Ok(i)).join(','));
  ok('DE IBAN: every <IBAN> in the XML is a 22-character German IBAN', allIbans.every((i) => /^DE\d{20}$/.test(i)));
}

const sDe = summarize(pDe);
eq('DE summarize: balance check passes', sDe.balanceCheckOk, true);
close('DE summarize: balanceDiff is 0', sDe.balanceDiff, 0);
eq('DE summarize: entryCount 4', sDe.entryCount, 4);
eq('DE summarize: 1 credit, 3 debits', sDe.creditCount + '/' + sDe.debitCount, '1/3');
close('DE summarize: openingBalance 12480.55', sDe.openingBalance, 12480.55);
close('DE summarize: closingBalance 13952.28', sDe.closingBalance, 13952.28);
close('DE summarize: creditSum 1785.00', sDe.creditSum, 1785);
close('DE summarize: debitSum 313.27', sDe.debitSum, 313.27);
close('DE summarize: opening + credits - debits = closing, computed here independently', sDe.openingBalance + sDe.creditSum - sDe.debitSum, sDe.closingBalance, 0.001);
deepEq('DE summarize: single currency EUR', sDe.currencies, ['EUR']);

// umlauts through the export pipeline
{
  const csvDe = toCsv(rowsDe, { bom: false, decimalComma: true, delimiter: ';', labels: columnLabelsMap('de') });
  includes('DE CSV: umlaut counterparty survives toCsv', csvDe, 'Bürobedarf Musterstadt');
  includes('DE CSV: umlaut Verwendungszweck survives toCsv', csvDe, 'Kontoführungsentgelt');
  includes('DE CSV: German decimal comma on the amount', csvDe, '-86,37');
  eq('DE CSV: header + 4 data lines', csvDe.split('\r\n').length, 5);
  const xlsxDe = buildXlsx('camt053', COLUMNS.map((c) => columnLabel(c.key, 'de')), rowsDe.map((r) => COLUMNS.map((c) => r[c.key])));
  const xlsxText = new TextDecoder('utf-8').decode(xlsxDe);
  includes('DE XLSX: umlaut counterparty is UTF-8 inside the sheet (STORED zip, so readable as text)', xlsxText, 'Bürobedarf Musterstadt');
  includes('DE XLSX: umlaut Verwendungszweck is UTF-8 inside the sheet', xlsxText, 'Kontoführungsentgelt');
  includes('DE XLSX: German column header written', xlsxText, 'Buchungsdatum');
}

// The two samples are distinct and each still parses on its own.
ok('samples: Slovak and German samples are different documents', SAMPLE_CAMT053_XML !== SAMPLE_CAMT053_XML_DE);
eq('samples: Slovak sample still detects Tatra banka via Stmt/Svcr (unchanged by the Acct/Svcr fallback)', parse(SAMPLE_CAMT053_XML).bank, 'tatrabanka');
ok('DE sample: no Slovak counterparty leaked into the German sample', !SAMPLE_CAMT053_XML_DE.includes('Odberatel') && !SAMPLE_CAMT053_XML_DE.includes('SK68'));

// ═══════════════════════════ default column set per language ═════════════
// What the table and the download show before the visitor touches a
// toggle (index.html checkedColumns starts as defaultColumnsForLang). Every
// default column must be filled in every row of the sample that language
// loads: de/en load the German sample, where VS/ŠS/KS are always empty, so
// they must not be on by default there; sk loads the Tatra banka sample,
// where VS/ŠS/KS are filled in at least one row each (a Slovak bank does
// not set all three symbols on every payment).
{
  const filled = (r, k) => (k === 'amount' || k === 'charges') ? Number.isFinite(r[k]) : (typeof r[k] === 'string' && r[k].length > 0);
  const colKeys = COLUMNS.map((c) => c.key);
  deepEq('DEFAULT_COLUMNS_BY_LANG: one entry per supported language', Object.keys(DEFAULT_COLUMNS_BY_LANG).sort(), [...LANGS].sort());
  for (const lang of LANGS) {
    const d = defaultColumnsForLang(lang);
    ok(`defaultColumnsForLang(${lang}): every key is a real COLUMNS key`, d.every((k) => colKeys.includes(k)), d.filter((k) => !colKeys.includes(k)).join(','));
    ok(`defaultColumnsForLang(${lang}): no duplicate keys`, new Set(d).size === d.length);
    ok(`defaultColumnsForLang(${lang}): booking date, amount, currency, counterparty and message are on`, ['bookingDate', 'amount', 'currency', 'counterpartyName', 'message'].every((k) => d.includes(k)));
    ok(`defaultColumnsForLang(${lang}): returns a fresh copy each call`, (() => { const a = defaultColumnsForLang(lang); a.push('charges'); return !defaultColumnsForLang(lang).includes('charges'); })());
  }
  deepEq('defaultColumnsForLang(sk): Slovak keeps VS/ŠS/KS on, the historical default', defaultColumnsForLang('sk'), ['bookingDate', 'valueDate', 'amount', 'currency', 'counterpartyName', 'vs', 'ss', 'ks', 'message']);
  for (const lang of ['de', 'en']) {
    const d = defaultColumnsForLang(lang);
    ok(`defaultColumnsForLang(${lang}): VS/ŠS/KS are off (a German statement never fills them)`, !d.includes('vs') && !d.includes('ss') && !d.includes('ks'));
    ok(`defaultColumnsForLang(${lang}): counterparty IBAN is on instead`, d.includes('counterpartyIban'));
    const gaps = d.filter((k) => !rowsDe.every((r) => filled(r, k)));
    ok(`defaultColumnsForLang(${lang}): every default column is filled in every row of the German sample`, gaps.length === 0, 'empty somewhere: ' + gaps.join(','));
  }
  {
    const d = defaultColumnsForLang('sk');
    const never = d.filter((k) => !rows02.some((r) => filled(r, k)));
    ok('defaultColumnsForLang(sk): every default column is filled in at least one row of the Slovak sample', never.length === 0, 'never filled: ' + never.join(','));
    const core = d.filter((k) => !['vs', 'ss', 'ks'].includes(k));
    const gaps = core.filter((k) => !rows02.every((r) => filled(r, k)));
    ok('defaultColumnsForLang(sk): every non-symbol default column is filled in every row of the Slovak sample', gaps.length === 0, 'empty somewhere: ' + gaps.join(','));
  }
  deepEq('defaultColumnsForLang: unknown language falls back to the active language (en in Node)', defaultColumnsForLang('xx'), defaultColumnsForLang('en'));
}

// ═══════════════════════════ parse(): .08 sample, 2 statements ═══════════

const p08 = parse(SAMPLE_08_XML);
ok('parse .08: ok is true', p08.ok === true);
eq('parse .08: version detected as 001.08', p08.version, '001.08');
eq('parse .08: two statements', p08.statements.length, 2);

const rows08 = toRows(p08);
eq('toRows .08: 5 rows total (3 + 1 bare + 2 batched - wait: 1+1+2 in stmt1, 1 in stmt2 = 5)', rows08.length, 5);

const bareRow = rows08[0];
eq('bare Ntry (no NtryDtls): still produces exactly one row', bareRow.amount, 200);
eq('bare Ntry: counterparty name empty (nothing to read)', bareRow.counterpartyName, '');
eq('bare Ntry: currency falls back to account Ccy (Amt had none)', bareRow.currency, 'EUR');

const structuredRow = rows08[1];
eq('VS source 2/3: from RmtInf/Strd/CdtrRefInf/Ref', structuredRow.vs, '445566');
eq('VS source 2/3: vsSource tagged "structured"', structuredRow.vsSource, 'structured');
eq('structured-ref row: signed amount negative (DBIT)', structuredRow.amount, -50);

const batchRow1 = rows08[2];
const batchRow2 = rows08[3];
eq('batched Ntry with 2 TxDtls: produces 2 rows, first amount from its own TxDtls/Amt', batchRow1.amount, -30);
eq('batched Ntry: second row amount from its own TxDtls/Amt', batchRow2.amount, -45.5);
eq('batched Ntry: rows keep distinct counterparties', batchRow1.counterpartyName, 'Prvy dodavatel');
eq('batched Ntry: rows keep distinct counterparties (2nd)', batchRow2.counterpartyName, 'Druhy dodavatel');
eq('batched Ntry: entity in Ustrd decoded ("&amp;" -> "&")', batchRow1.message, 'Faktura A & B');
eq('batched Ntry: both rows share the same parent entryIndex', batchRow1.entryIndex, batchRow2.entryIndex);

const statement2Row = rows08[4];
eq('second statement: its row is present too', statement2Row.statementId, 'SK3333000000000077777777-001-260906');

// ═══════════════════════════ summarize() ═══════════════════════════════

const s08 = summarize(p08);
eq('summarize .08: aggregates 5 entries across 2 statements', s08.entryCount, 5);
close('summarize .08: statement 1 balances (OPBD 500 + net 74.50 = CLBD 574.50)', s08.perStatement[0].balanceDiff, 0);
eq('summarize .08: statement 1 balanceCheckOk true', s08.perStatement[0].balanceCheckOk, true);
eq('summarize .08: statement 2 balanceCheckOk false (CLBD does not reconcile)', s08.perStatement[1].balanceCheckOk, false);
close('summarize .08: statement 2 balanceDiff is -889.00', s08.perStatement[1].balanceDiff, -889, 0.01);
eq('summarize .08: overall balanceCheckOk false when any statement fails', s08.balanceCheckOk, false);
close('summarize .08: aggregated creditSum (200 + 10)', s08.creditSum, 210);
close('summarize .08: aggregated debitSum (50 + 30 + 45.50)', s08.debitSum, 125.5);

const s02 = summarize(p02);
eq('summarize .02: balanceCheckOk true for the bundled sample', s02.balanceCheckOk, true);
close('summarize .02: creditSum is 450.00', s02.creditSum, 450);
close('summarize .02: debitSum is 210.70', s02.debitSum, 210.7);
close('summarize .02: netSum is creditSum - debitSum', s02.netSum, 239.3);
eq('summarize .02: openingBalance read from OPBD', s02.openingBalance, 1000);
eq('summarize .02: closingBalance read from CLBD', s02.closingBalance, 1239.3);

// multi-file merge: caller concatenates statements from two parse() calls
const merged = { statements: [...p02.statements, ...p08.statements] };
const sMerged = summarize(merged);
eq('summarize: accepts a hand-merged {statements} object (multi-file upload)', sMerged.statementCount, 3);
eq('summarize: merged entryCount is the sum of all three statements', sMerged.entryCount, 3 + 5);
const rowsMerged = toRows(merged);
eq('toRows: accepts the same hand-merged object', rowsMerged.length, 8);

// defensive: missing/invalid input never throws
deepEq('toRows(null) returns an empty array', toRows(null), []);
deepEq('toRows(undefined) returns an empty array', toRows(undefined), []);
eq('summarize(null) returns ok:false with zeroed totals', summarize(null).ok, false);
eq('summarize(null): entryCount is 0', summarize(null).entryCount, 0);

// ═══════════════════════════ tolerant parsing / bad input ═══════════════

const pNotCamt = parse(NOT_CAMT_XML);
eq('parse: well-formed XML but not camt.053 -> ok false', pNotCamt.ok, false);
eq('parse: BkToCstmrStmt missing is reported in errors', pNotCamt.errors.some((e) => e.includes('BkToCstmrStmt')), true);
eq('parse: unrecognized namespace -> version "other"', pNotCamt.version, 'other');

const pMalformed = parse(MALFORMED_XML);
eq('parse: unclosed tags -> malformed true', pMalformed.malformed, true);
ok('parse: unclosed tags -> at least one parser error reported', pMalformed.errors.length > 0);
eq('parse: malformed input -> ok false', pMalformed.ok, false);

const pGarbage = parse(GARBAGE);
eq('parse: plain text with no XML -> ok false', pGarbage.ok, false);
ok('parse: plain text -> errors reported', pGarbage.errors.length > 0);

const pEmpty = parse('');
eq('parse: empty string -> ok false, no throw', pEmpty.ok, false);

// ═══════════════════════════ bankFromBic() ═══════════════════════════════

eq('bankFromBic: TATRSKBX -> tatrabanka', bankFromBic('TATRSKBX'), 'tatrabanka');
eq('bankFromBic: GIBASKBX -> slsp', bankFromBic('GIBASKBX'), 'slsp');
eq('bankFromBic: SUBASKBX -> vub', bankFromBic('SUBASKBX'), 'vub');
eq('bankFromBic: CEKOSKBX -> csob', bankFromBic('CEKOSKBX'), 'csob');
eq('bankFromBic: unknown BIC -> "iná"', bankFromBic('ZZZZZZZZ'), 'iná');
eq('bankFromBic: empty -> empty string', bankFromBic(''), '');
eq('bankFromBic: lowercase input still matches (case-insensitive)', bankFromBic('tatrskbx'), 'tatrabanka');

// ═══════════════════════════ toCsv() ═════════════════════════════════════

const csvRows = [
  { bookingDate: '2026-09-02', amount: 450, currency: 'EUR', vs: '123', message: 'Faktura "2026"; polozka\nriadok2' },
  { bookingDate: '2026-09-03', amount: -1234.5, currency: 'EUR', vs: '456', message: 'bez specialnych znakov' },
];

const csvDefault = toCsv(csvRows, { columns: ['bookingDate', 'amount', 'vs', 'message'] });
eq('toCsv: has UTF-8 BOM by default', csvDefault.charCodeAt(0), 0xfeff);
const csvNoBom = toCsv(csvRows, { columns: ['bookingDate', 'amount', 'vs', 'message'], bom: false });
ok('toCsv: bom:false omits the BOM', csvNoBom.charCodeAt(0) !== 0xfeff);
includes('toCsv: header row uses column labels', csvNoBom, 'Dátum zaúčtovania;Suma;VS;Správa pre príjemcu');
includes('toCsv: field with ";", quote and newline is quoted and doubled', csvNoBom, '"Faktura ""2026""; polozka\nriadok2"');
includes('toCsv: plain field is left unquoted', csvNoBom, 'bez specialnych znakov');
includes('toCsv: amount formatted to 2 decimals with "." by default', csvNoBom, '450.00');
includes('toCsv: negative amount formatted with "." by default', csvNoBom, '-1234.50');
const csvComma = toCsv(csvRows, { columns: ['amount'], bom: false, decimalComma: true });
includes('toCsv: decimalComma:true swaps "." for ","', csvComma, '450,00');
ok('toCsv: decimalComma:true never leaves a bare "." in an amount', !csvComma.includes('450.00'));
const csvTab = toCsv(csvRows, { columns: ['bookingDate', 'amount'], bom: false, delimiter: '\t' });
includes('toCsv: custom delimiter (tab) used between columns', csvTab, '2026-09-02\t450.00');
const csvSubset = toCsv(rows02, { columns: ['vs'], bom: false });
eq('toCsv: columns option restricts to just the requested column', csvSubset.split('\r\n')[0], 'VS');
const csvNullAmount = toCsv([{ amount: null }], { columns: ['amount'], bom: false });
eq('toCsv: null amount renders as an empty cell, not "null" or "NaN"', csvNullAmount.split('\r\n')[1], '');
eq('toCsv: line endings are CRLF', csvNoBom.includes('\r\n'), true);
eq('toCsv: full-run CSV on the .02 sample has header + 3 data lines', toCsv(rows02, { bom: false }).split('\r\n').length, 4);
eq('COLUMNS: 18 columns defined', COLUMNS.length, 18);

// ═══════════════════════════ tolerant XML parser (internals) ════════════

const treeOk = parseXml('<a x="1"><b>hi &amp; bye</b><c/></a>');
eq('parseXml: well-formed input is not malformed', treeOk.malformed, false);
const aNode = treeOk.root.children[0].node;
eq('parseXml: attribute value read', aNode.attrs.x, '1');
eq('parseXml: entity decoded in text content', aNode.children[0].node.children[0].text, 'hi & bye');
eq('parseXml: self-closing element has no children', aNode.children[1].node.children.length, 0);

// ═══════════════════════════ xlsx-writer.js ══════════════════════════════

eq('crc32: matches the standard "123456789" check value (0xCBF43926)', crc32(new TextEncoder().encode('123456789')), 0xcbf43926);
eq('crc32: empty input is 0', crc32(new Uint8Array(0)), 0);

eq('colLetter: 0 -> A', colLetter(0), 'A');
eq('colLetter: 25 -> Z', colLetter(25), 'Z');
eq('colLetter: 26 -> AA', colLetter(26), 'AA');
eq('colLetter: 27 -> AB', colLetter(27), 'AB');
eq('colLetter: 51 -> AZ', colLetter(51), 'AZ');
eq('colLetter: 701 -> ZZ', colLetter(701), 'ZZ');
eq('colLetter: 702 -> AAA', colLetter(702), 'AAA');

const zipBytes = buildZip([
  { name: 'a.xml', data: new TextEncoder().encode('<x/>') },
  { name: 'b.xml', data: new TextEncoder().encode('<y/>') },
]);
eq('buildZip: starts with the local file header signature (PK\\x03\\x04)', zipBytes[0] === 0x50 && zipBytes[1] === 0x4b && zipBytes[2] === 0x03 && zipBytes[3] === 0x04, true);
function countLocalHeaders(bytes) {
  let n = 0;
  for (let i = 0; i < bytes.length - 3; i++) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) n++;
  }
  return n;
}
eq('buildZip: exactly one local file header per input file', countLocalHeaders(zipBytes), 2);
function hasEocd(bytes) {
  for (let i = 0; i < bytes.length - 3; i++) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) return true;
  }
  return false;
}
ok('buildZip: end-of-central-directory record present', hasEocd(zipBytes));

const xlsxBytes = buildXlsx('camt053', COLUMNS.map((c) => c.label), [
  ['2026-09-02', 450], ['2026-09-03', -89.9],
]);
ok('buildXlsx: produces a non-trivial byte array', xlsxBytes.length > 200);
eq('buildXlsx: starts with a ZIP signature', xlsxBytes[0] === 0x50 && xlsxBytes[1] === 0x4b, true);
eq('buildXlsx: exactly 5 parts (5 local file headers)', countLocalHeaders(xlsxBytes), 5);
ok('buildXlsx: end-of-central-directory record present', hasEocd(xlsxBytes));

// central-directory entry count matches file count (2 bytes, little-endian,
// 10 bytes before the very end of the EOCD record's fixed 22-byte tail).
function eocdEntryCount(bytes) {
  const eocdStart = bytes.length - 22; // no zip comment written, so it's fixed-size
  return bytes[eocdStart + 10] | (bytes[eocdStart + 11] << 8);
}
eq('buildXlsx: EOCD reports 5 entries in the central directory', eocdEntryCount(xlsxBytes), 5);

// ═══════════════════════════ licence.js (Pro) ════════════════════════════
// licence.js's real verify()/isValid() check every licence against the
// ARLing service's actual public key baked into that file, and this
// repo, correctly, does not hold the matching private key. So every test
// below signs its own fixture licences with a throwaway Ed25519 keypair
// generated right here, and passes that test key in as verify()/
// isValid()'s documented test-only override, so the mechanism under test
// is licence.js's real code, not a reimplementation of it. Same approach
// as sepa-pain001-generator/tests.mjs.

function licB64u(bytes) {
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return Buffer.from(bin, 'binary').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function licStableJson(obj) {
  return '{' + Object.keys(obj).sort().map((k) => JSON.stringify(k) + ':' + JSON.stringify(obj[k])).join(',') + '}';
}
async function licSign(payloadObj, privateKey) {
  const payloadBytes = new TextEncoder().encode(licStableJson(payloadObj));
  const sig = new Uint8Array(await crypto.subtle.sign('Ed25519', privateKey, payloadBytes));
  return licB64u(payloadBytes) + '.' + licB64u(sig);
}
function licAddDaysIso(iso, days) {
  const [y, mo, d] = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(y, mo - 1, d));
  dt.setUTCDate(dt.getUTCDate() + days);
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, '0')}-${String(dt.getUTCDate()).padStart(2, '0')}`;
}

await (async () => {
  eq('licence DEFAULT_PLAN: is exactly "sepa-pro" (the shared bundle plan)', DEFAULT_PLAN, 'sepa-pro');

  const testKeyPair = await crypto.subtle.generateKey('Ed25519', true, ['sign', 'verify']);
  const testPubRaw = new Uint8Array(await crypto.subtle.exportKey('raw', testKeyPair.publicKey));

  const TODAY = licenceTodayIso();
  const TOMORROW = licAddDaysIso(TODAY, 1);
  const YESTERDAY = licAddDaysIso(TODAY, -1);
  const basePayload = { p: DEFAULT_PLAN, e: TOMORROW, s: 'abcd1234', m: '0123456789abcdef' };
  const validKey = await licSign(basePayload, testKeyPair.privateKey);

  eq('licence parse: malformed key returns null', parseLicence('garbage'), null);
  {
    const parsed = parseLicence(validKey);
    ok('licence parse: well-formed key parses', parsed !== null);
    eq('licence parse: plan field round-trips to "sepa-pro"', parsed.payload.p, 'sepa-pro');
  }

  eq('licence verify: valid signature against the matching (test) pubkey', await verifyLicence(validKey, testPubRaw), true);
  eq('licence verify: signature by a foreign keypair rejected by the real embedded ARLing pubkey', await verifyLicence(validKey), false);

  {
    const r = await isValidLicence(validKey, { pubKey: testPubRaw });
    eq('isValid: valid "sepa-pro" licence -> valid true', r.valid, true);
    eq('isValid: valid "sepa-pro" licence -> reason "ok"', r.reason, 'ok');
  }
  {
    const key = await licSign({ ...basePayload, e: YESTERDAY }, testKeyPair.privateKey);
    const r = await isValidLicence(key, { pubKey: testPubRaw });
    eq('isValid: expired licence -> valid false', r.valid, false);
    eq('isValid: expired licence -> reason "expired"', r.reason, 'expired');
  }
  {
    const key = await licSign({ ...basePayload, p: 'sepa-generator-pro' }, testKeyPair.privateKey);
    const r = await isValidLicence(key, { pubKey: testPubRaw });
    eq('isValid: a different ARLing tool\'s own plan is rejected here (camt.053 only accepts "sepa-pro")', r.valid, false);
    eq('isValid: wrong plan -> reason "plan"', r.reason, 'plan');
  }

  clearLicence();
  eq('licence load: nothing stored returns null', loadLicence(), null);
  eq('licence save: reports success', saveLicence(validKey), true);
  eq('licence load: round-trips the exact stored string', loadLicence(), validKey);
  eq('licence clear: reports success', clearLicence(), true);
  eq('licence load: returns null again after clear', loadLicence(), null);
  ok('licence STORAGE_KEY: is exactly "arling_licence_sepa-pro" (shared with the other 3 tools)', LICENCE_STORAGE_KEY === 'arling_licence_sepa-pro');
})();

// ═══════════════════════════ i18n.js: SK/EN/DE dictionary ═══════════════
// The tool is one page for accountants in Slovakia, Germany, Austria and
// Switzerland (SK/EN/DE), driven by a single dictionary object and pure
// helpers in i18n.js. These assertions check the dictionary itself, not
// the DOM wiring (applyI18n/setLang), which needs a real browser: every
// key has all three languages filled in, the specific header/number/date
// rules from the brief hold per language, and no Slovak string leaked
// into the English or German copy.

eq('i18n LANGS: exactly [sk, en, de]', LANGS.join(','), 'sk,en,de');
eq('i18n DEFAULT_LANG: "en" (fallback when navigator.language is neither de nor sk/cs)', DEFAULT_LANG, 'en');

{
  const incomplete = findIncompleteEntries();
  deepEq('i18n dictionary: every DICT/COLUMN_LABELS entry has a non-empty sk/en/de', incomplete, []);
}
ok('i18n dictionary: has a substantial number of keys (every visible string on the page)', Object.keys(DICT).length >= 100);
deepEq('i18n COLUMN_LABELS: has exactly the same 18 column keys as camt053.js COLUMNS', Object.keys(COLUMN_LABELS).slice().sort(), COLUMNS.map((c) => c.key).slice().sort());

// ── header rows per language (the VS/SS/KS carve-out from the brief) ─────
eq('columnLabel: vs header stays "VS" in Slovak', columnLabel('vs', 'sk'), 'VS');
eq('columnLabel: vs header becomes "Reference (VS/SS/KS)" in English', columnLabel('vs', 'en'), 'Reference (VS/SS/KS)');
eq('columnLabel: vs header becomes "Verwendungszweck / Referenz" in German', columnLabel('vs', 'de'), 'Verwendungszweck / Referenz');
eq('columnLabel: bookingDate translated to English', columnLabel('bookingDate', 'en'), 'Booking date');
eq('columnLabel: bookingDate translated to German', columnLabel('bookingDate', 'de'), 'Buchungsdatum');
eq('columnLabel: unknown column key falls back to the key itself', columnLabel('doesNotExist', 'en'), 'doesNotExist');
{
  const enMap = columnLabelsMap('en');
  eq('columnLabelsMap: covers every COLUMNS key', Object.keys(enMap).length, COLUMNS.length);
  eq('columnLabelsMap: message column keeps a translated RmtInf free-text label in English', enMap.message, 'Message (RmtInf free text)');
}

// toCsv()'s opts.labels override (the one permitted change to camt053.js):
// still optional (existing sk-default behaviour untouched) and, when
// supplied, produces a translated header row without touching row data.
{
  const rows = [{ vs: '123', message: 'hello' }];
  const csvDefault = toCsv(rows, { columns: ['vs', 'message'], bom: false });
  includes('toCsv: opts.labels is optional, header stays Slovak by default', csvDefault, 'VS;Správa pre príjemcu');
  const csvEn = toCsv(rows, { columns: ['vs', 'message'], bom: false, labels: columnLabelsMap('en') });
  includes('toCsv: opts.labels overrides the header row with English column labels', csvEn, 'Reference (VS/SS/KS);Message (RmtInf free text)');
  includes('toCsv: opts.labels does not touch row data, only the header', csvEn, '123;hello');
  const csvDe = toCsv(rows, { columns: ['vs', 'message'], bom: false, labels: columnLabelsMap('de') });
  includes('toCsv: opts.labels overrides the header row with German column labels', csvDe, 'Verwendungszweck / Referenz;Verwendungszweck-Text (RmtInf)');
}

// ── t()/tf() lookup ────────────────────────────────────────────────────
eq('t: unknown key returns the key itself (missing translation stays visible, not blank)', t('no.such.key', 'en'), 'no.such.key');
eq('t: falls back to DEFAULT_LANG for an unsupported language code', t('js.status.ok', 'fr'), t('js.status.ok', 'en'));
includes('tf: fills a single {placeholder}', tf('js.download.all', { n: 5 }, 'en'), '5');
includes('tf: fills a {placeholder} used inside a longer German string', tf('js.sizeWarn', { size: '25 MB', limit: '20 MB' }, 'de'), '25 MB');

// ── number formatting per language ──────────────────────────────────────
eq('formatAmountForLang: English keeps a decimal point', formatAmountForLang(1234.5, 'en'), '1234.50');
eq('formatAmountForLang: Slovak uses a decimal comma', formatAmountForLang(1234.5, 'sk'), '1234,50');
eq('formatAmountForLang: German uses a decimal comma', formatAmountForLang(1234.5, 'de'), '1234,50');
eq('formatAmountForLang: negative amount, German decimal comma', formatAmountForLang(-89.9, 'de'), '-89,90');
eq('formatAmountForLang: null amount renders as an empty string, not "null"', formatAmountForLang(null, 'en'), '');
eq('formatAmountForLang: NaN renders as an empty string', formatAmountForLang(NaN, 'sk'), '');

// ── date formatting per language ────────────────────────────────────────
eq('formatDateForLang: English keeps ISO yyyy-mm-dd', formatDateForLang('2026-09-02', 'en'), '2026-09-02');
eq('formatDateForLang: Slovak reformats to dd.mm.yyyy', formatDateForLang('2026-09-02', 'sk'), '02.09.2026');
eq('formatDateForLang: German reformats to dd.mm.yyyy', formatDateForLang('2026-09-02', 'de'), '02.09.2026');
eq('formatDateForLang: empty input passes through as an empty string', formatDateForLang('', 'en'), '');
eq('formatDateForLang: non-ISO input passes through unchanged', formatDateForLang('n/a', 'de'), 'n/a');

// ── CSV export defaults per language: comma+semicolon for sk/de, point+
// comma for en, exactly as specified in the brief ────────────────────────
deepEq('defaultCsvOptsForLang: Slovak defaults to decimal comma + semicolon', defaultCsvOptsForLang('sk'), { decimalComma: true, delimiter: ';' });
deepEq('defaultCsvOptsForLang: German defaults to decimal comma + semicolon', defaultCsvOptsForLang('de'), { decimalComma: true, delimiter: ';' });
deepEq('defaultCsvOptsForLang: English defaults to decimal point + comma', defaultCsvOptsForLang('en'), { decimalComma: false, delimiter: ',' });

// ── locale detection (pure logic; the DOM-facing detectLang() wraps this
// with location.search / localStorage / navigator.language, untestable
// under Node without a browser) ──────────────────────────────────────────
eq('langFromLocale: "de-DE" -> de', langFromLocale('de-DE'), 'de');
eq('langFromLocale: "de-AT" -> de (Austrian German)', langFromLocale('de-AT'), 'de');
eq('langFromLocale: "de-CH" -> de (Swiss German)', langFromLocale('de-CH'), 'de');
eq('langFromLocale: "sk-SK" -> sk', langFromLocale('sk-SK'), 'sk');
eq('langFromLocale: "cs-CZ" -> sk (Czech maps to Slovak, per the brief)', langFromLocale('cs-CZ'), 'sk');
eq('langFromLocale: "fr-FR" -> en (anything else defaults to English)', langFromLocale('fr-FR'), 'en');
eq('langFromLocale: empty/undefined -> en', langFromLocale(''), 'en');
eq('langFromQueryString: "?lang=de" -> de', langFromQueryString('?lang=de'), 'de');
eq('langFromQueryString: "?lang=SK" is case-insensitive -> sk', langFromQueryString('?lang=SK'), 'sk');
eq('langFromQueryString: unsupported ?lang= value -> null (caller falls through)', langFromQueryString('?lang=fr'), null);
eq('langFromQueryString: no ?lang= param -> null', langFromQueryString('?other=1'), null);

// ── misc per-language lookups used in the page ───────────────────────────
eq('localeTagForLang: sk -> sk-SK (history timestamp locale)', localeTagForLang('sk'), 'sk-SK');
eq('localeTagForLang: de -> de-DE', localeTagForLang('de'), 'de-DE');
eq('localeTagForLang: en -> en-GB', localeTagForLang('en'), 'en-GB');
eq('ogLocaleForLang: sk -> sk_SK', ogLocaleForLang('sk'), 'sk_SK');
eq('ogLocaleForLang: de -> de_DE', ogLocaleForLang('de'), 'de_DE');
eq('ogLocaleForLang: en -> en_US', ogLocaleForLang('en'), 'en_US');

// ── accounting-software templates: SK list stays (generic only), DE/EN
// additionally offer DATEV/Lexware/sevDesk generic-CSV presets ───────────
deepEq('templateOrderForLang: Slovak only offers the generic preset', templateOrderForLang('sk'), ['generic']);
deepEq('templateOrderForLang: German offers generic + DATEV/Lexware/sevDesk', templateOrderForLang('de'), ['generic', 'datev', 'lexware', 'sevdesk']);
deepEq('templateOrderForLang: English offers the same list as German', templateOrderForLang('en'), templateOrderForLang('de'));
eq('TEMPLATE_PRESETS.generic: null (signal value, keep current columns)', TEMPLATE_PRESETS.generic, null);
ok('TEMPLATE_PRESETS.datev: a real column subset, includes bookingDate', Array.isArray(TEMPLATE_PRESETS.datev) && TEMPLATE_PRESETS.datev.includes('bookingDate'));
eq('templateLabel: DATEV preset labelled as a generic CSV layout in German, not vendor-certified', templateLabel('datev', 'de'), 'DATEV (generisches CSV-Layout)');

// ── no leftover Slovak in the English/German copy ────────────────────────
{
  const leftoverWords = ['Máte', 'Dostanete'];
  let anyLeftover = false;
  for (const [key, entry] of Object.entries(DICT)) {
    for (const lang of ['en', 'de']) {
      for (const w of leftoverWords) {
        if (String(entry[lang] || '').includes(w)) anyLeftover = true;
      }
    }
  }
  eq('i18n dictionary: no leftover Slovak "Máte"/"Dostanete" in any English or German value', anyLeftover, false);
}
eq('rendered English: playground heading has no leftover "Dostanete"', t('s2.h2', 'en').includes('Dostanete'), false);
eq('rendered German: playground heading has no leftover "Dostanete"', t('s2.h2', 'de').includes('Dostanete'), false);
eq('rendered English: hero lead has no leftover "Máte"', t('hero.lead', 'en').includes('Máte'), false);
eq('rendered German: hero lead has no leftover "Máte"', t('hero.lead', 'de').includes('Máte'), false);

// ── sample-button copy matches the sample each language actually loads:
// sk -> Tatra banka sample (3 entries), en/de -> German sample (4) ───────
includes('i18n sample title: Slovak names Tatra banka', t('s2.sample.btn.title', 'sk'), 'Tatra banky');
includes('i18n sample title: Slovak says 3 entries', t('s2.sample.btn.title', 'sk'), '3 polo');
includes('i18n sample title: English says a German bank', t('s2.sample.btn.title', 'en'), 'German bank');
includes('i18n sample title: English says 4 entries', t('s2.sample.btn.title', 'en'), '4 entries');
includes('i18n sample title: German says a German bank', t('s2.sample.btn.title', 'de'), 'deutschen Bank');
includes('i18n sample title: German says 4 Positionen', t('s2.sample.btn.title', 'de'), '4 Positionen');
eq('i18n sample title: English no longer claims 3 entries', t('s2.sample.btn.title', 'en').includes('3 entries'), false);
eq('i18n sample title: German no longer claims 3 Positionen', t('s2.sample.btn.title', 'de').includes('3 Positionen'), false);
includes('i18n sample loaded: Slovak names the Tatra banka sample with 3 entries', tf('js.sample.loaded', { kbd1: 'A', kbd2: 'B' }, 'sk'), 'Tatra banky, 3 polo');
includes('i18n sample loaded: English names the German statement with 4 entries', tf('js.sample.loaded', { kbd1: 'A', kbd2: 'B' }, 'en'), 'German bank statement, 4 entries');
includes('i18n sample loaded: German names the German statement with 4 Positionen', tf('js.sample.loaded', { kbd1: 'A', kbd2: 'B' }, 'de'), 'deutscher Kontoauszug, 4 Positionen');
{
  const entryCountSk = toRows(parse(SAMPLE_CAMT053_XML)).length;
  const entryCountDe = toRows(parse(SAMPLE_CAMT053_XML_DE)).length;
  includes('i18n sample title: the Slovak number matches the Slovak sample row count', t('s2.sample.btn.title', 'sk'), String(entryCountSk) + ' polo');
  includes('i18n sample title: the German number matches the German sample row count', t('s2.sample.btn.title', 'de'), String(entryCountDe) + ' Positionen');
  includes('i18n sample title: the English number matches the German sample row count', t('s2.sample.btn.title', 'en'), String(entryCountDe) + ' entries');
}

// ═══════════════════════════ summary ═══════════════════════════════════

console.log(`\n${pass} passed, ${fail} failed (${pass + fail} total assertions)`);
if (fail > 0) {
  console.log('\nFailures:');
  for (const f of failures) console.log(' - ' + f);
  process.exit(1);
}

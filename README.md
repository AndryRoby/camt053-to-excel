# camt.053 to Excel: bank statement XML to a CSV or Excel table

camt.053 to Excel converts a camt.053 XML bank statement (ISO 20022) into a CSV or Excel table in your browser, for accountants and companies that get statements from Tatra banka, SLSP, VÚB, ČSOB or a German, Austrian or Swiss bank and need the entries in a spreadsheet or in DATEV. Converting the whole statement, the totals and the balance check are free; without a licence the downloaded file holds the first 20 rows as a sample, and the full download plus MT940 and DATEV Buchungsstapel exports come with the Pro licence for €9 a month or €79 a year (VAT included) at https://arling.sk/bankove-nastroje/, which also covers SEPA pain.001 Generator and Payment matcher.

Live: https://arling.sk/camt053-to-excel/ (Slovak) · https://arling.sk/camt053-to-excel/en/ (English) · https://arling.sk/camt053-to-excel/de/ (German)

camt.053 in MT940 umwandeln, für DATEV (German page for the Pro MT940 and DATEV exports, €9 a month or €79 a year, VAT included): https://arling.sk/camt053-to-excel/mt940/?utm_source=github · English: https://arling.sk/camt053-to-excel/mt940/en/?utm_source=github

Nothing you upload is sent anywhere. The page has a language switcher
and is fully available in Slovak, English and German, with column
headers, error messages, the FAQ and the Pro section translated in
each language, and per-language defaults for number/date formatting
(decimal comma and dd.mm.yyyy for SK/DE, decimal point and yyyy-mm-dd
for EN) and CSV export presets (DATEV/Lexware/sevDesk generic layouts
for DE/EN). You don't have to eyeball raw XML, or pay for accounting
software's built-in import, just to see what's on a statement, pair it
against invoices, or check it by eye.

## What it's for

Every one of the four major Slovak banks issues its account statement
for accounting (výpis do účtovníctva) as camt.053 XML by default, the
national "SEPA XML SK štandard" the Slovak Bank Association (SBA)
published in 2013: an accountant or company that needs those
transactions in a spreadsheet (matching payments against invoices by
variabilný symbol, a manual reconciliation pass, or feeding a
bookkeeping process that isn't wired up to a bank's own accounting
software connector) has no direct route from "an XML file in the
Downloads folder" to a table, short of opening the XML by hand or
buying software that imports it. This tool is that route: upload the
file (or several), and download a table. The page's built-in sample
("ukážka"/"sample"/"Beispiel" button) is a German bank statement
(fictional Muster GmbH, four entries) when the page is in German or
English, and a Slovak Tatra banka statement (three entries) in Slovak.

## Input format

1. **The statement file(s).** Upload one or more `.xml` files with the
   file picker, or paste the XML text directly into the page. An
   uploaded file is read with the browser's own `FileReader` API and
   never leaves the page. The free version converts one file at a
   time; converting several files in one go (several months, or
   several accounts), each kept as its own table and its own download,
   is part of Pro.

2. **Supported camt.053 versions:**

   | Version | Namespace | Where it's used |
   |---|---|---|
   | camt.053.001.02 | `urn:iso:std:iso:20022:tech:xsd:camt.053.001.02` | The version behind the SBA's "SEPA XML SK štandard": what Tatra banka, SLSP, VÚB, and ČSOB all issue as "výpis do účtovníctva" / "XML - camt.053 SK štandard" from internet banking. |
   | camt.053.001.08 | `urn:iso:std:iso:20022:tech:xsd:camt.053.001.08` | A newer ISO 20022 revision, for a statement from a non-Slovak bank or a source that has moved to it. The entry-level fields this tool reads (`BookgDt`, `ValDt`, `Amt`, `CdtDbtInd`, remittance/reference text) are unchanged between the two versions. |

## What comes out

One row per underlying payment: a plain `Ntry` (statement entry) becomes
one row, and a batched entry with several `NtryDtls/TxDtls` blocks
becomes one row per `TxDtls`, so VS/ŠS/KS pairing against invoices
works per payment, not per bank-side batch line.

| Column | Source in the XML |
|---|---|
| Číslo výpisu | `Stmt/Id` |
| Účet (IBAN) | `Stmt/Acct/Id/IBAN` |
| Dátum zaúčtovania | `BookgDt` |
| Dátum valuty | `ValDt` |
| Suma | `Amt` (entry- or transaction-level), signed by `CdtDbtInd` (`DBIT` negative, `CRDT` positive) |
| Mena | `Amt/@Ccy`, falling back to the account currency |
| Status | `Sts` (e.g. `BOOK`, `PDNG`) |
| Referencia banky | `AcctSvcrRef`, falling back to `NtryRef` |
| Typ transakcie | `BkTxCd`: the structured `Domn/Fmly/SubFmlyCd` code, the bank's own `Prtry` code, or both |
| Protistrana | counterparty name from `RltdPties` (creditor on a debit entry, debtor on a credit entry) |
| IBAN protistrany | counterparty IBAN from `RltdPties`, where the statement includes it |
| BIC protistrany | counterparty BIC from `RltdAgts`, where present |
| EndToEndId | `Refs/EndToEndId` |
| VS / ŠS / KS | parsed out of `EndToEndId`'s `/VS.../SS.../KS...` convention, the structured remittance reference, or the free-text message, whichever carries them |
| Správa pre príjemcu | `RmtInf/Ustrd` |
| Poplatok | `Chrgs/TtlChrgsAndTaxAmt` or `Chrgs/Amt`, where the entry carries a fee |

A short summary above the table carries the account IBAN, statement
period, entry/credit/debit counts and totals, and the opening/closing
balance from the statement's own header (`Bal` entries): opening
balance plus the net of all entries is checked against the reported
closing balance, so a mismatch is visible before the row-level data is
trusted. The statement's servicing bank (`Svcr/FinInstnId/BIC`) is also
matched against a table of Slovak bank BICs to label which bank issued
each file.

## How it works (client-side only)

Everything runs in your browser. There is no backend and no account.
You upload or paste one or more camt.053 files, and the page parses
them, shows the totals, the balance check and the table on screen (the
first 200 rows of a longer statement), and offers the result as a CSV
or Excel download. Without a licence the download holds the first 20
rows with a visible sample header; with a licence you get the whole
file (see Pro below).

Nothing about your statement is sent anywhere: no IBANs, no
transaction amounts, no counterparty names. The only network activity
this site generates is:

- loading its own static assets (HTML/CSS/JS) from GitHub Pages,
- anonymous product-analytics events (page view, "convert" clicked,
  etc.) sent to a self-hosted Umami instance: event names and counts
  only, never the content of your statement,
- and, only if you type an email into the optional mailing-list form,
  a request to the subscribe endpoint carrying that email address and
  nothing else.

You can verify this yourself: open your browser's network tab while
using the tool, or just read `index.html` and its engine script
`camt053.js`; it's static files with no build step. The engine has no
dependency on `DOMParser`, so it also runs unmodified in Node (used by
this project's own test suite, `tests.mjs`), the same approach this
project's sibling tool, SEPA pain.001 Doctor, uses for parsing XML.

## Pro: full download, MT940 and DATEV Buchungsstapel exports

Pro is the Banking tools licence sold at
https://arling.sk/bankove-nastroje/ for €9 a month or €79 a year, VAT
included. One licence activates Pro in three tools: camt.053 to Excel,
SEPA pain.001 Generator and Payment matcher (SEPA pain.001 Doctor is
free and needs no licence). In this tool Pro adds the whole downloaded
file as CSV and Excel, the MT940 and DATEV exports described below,
several files at once, a history of past conversions stored in your
browser, and priority email support.

The licence is sold through Stripe Managed Payments. The merchant of
record is Link (Sold through Link, LLC, which provides that service for
Stripe): Link sends the receipt and the invoice as a PDF, and Stripe
calculates and remits the VAT; ARLing s. r. o. delivers the tool and the
licence key. Cancel or change the subscription at any time in the
Stripe customer portal
(https://billing.stripe.com/p/login/3cIaER9M63hNeFcg8B4ko00); it stays
active until the end of the paid period. For a monthly or yearly
subscription, ARLing refunds the payment on request within 14 days of
purchase, without you giving a reason: write to support@arling.sk.
Full terms: https://arling.sk/podmienky/en/ (sections 4 to 6).

German banks stopped issuing MT940 in November 2025: the German banking
industry's own rulebook now specifies camt.053 only for the electronic
account statement. DATEV Kanzlei-Rechnungswesen, however, still has no
direct camt.053 file import; its manual file import (Stapelverarbeitung)
still expects MT940 (the paid DATEV Bankdatenservice covers camt.053 a
different way). That gap is what these two Pro exports are for, both
generated client-side exactly like the free CSV/Excel export, just
gated on the Banking tools licence for the actual download; without a
licence, clicking either button shows a preview of the first eight
lines instead. A dedicated landing page for this pair of exports
lives at `mt940/index.html` (German, live at
https://arling.sk/camt053-to-excel/mt940/) and `mt940/en/index.html`
(English, https://arling.sk/camt053-to-excel/mt940/en/); its CTA links to the converter with `?export=mt940`, which
`index.html` reads to show a hint and highlight the MT940 button once a
statement is converted.

- **MT940 (`.sta`)**, built by `mt940.js`: one SWIFT MT940 message per
  camt.053 `<Stmt>`, with the German `:86:` structured subfield
  convention (GVC code, `?00`-`?34` subfields, SEPA `EREF+`/`SVWZ+` tags
  inside the `?20`-`?29` Verwendungszweck lines). For an import into
  DATEV or any other MT940-only software.
- **DATEV Buchungsstapel (EXTF CSV)**, built by `datev-extf.js`: the
  DATEV-Format "Buchungsstapel" (booking batch) header, its column
  header row, then one row per entry, with a small options block
  (Kontonummer der Bank, Sachkontenlänge, Beraternummer,
  Mandantennummer, Wirtschaftsjahr-Beginn) so the file matches your own
  DATEV client setup.

**Known limits, honestly:**

- The exact wording of a bank's own `:86:` `?00` Buchungstext, and which
  of `EREF+`/`MREF+`/`CRED+`/`DEBT+`/`SVWZ+` a bank fills in, differs
  bank to bank; camt.053 (and this tool's parser) does not carry a
  mandate id or creditor scheme id today, so this export only ever
  writes `EREF+` (from `EndToEndId`) and `SVWZ+` (from the remittance
  text). It is a well-formed, generic MT940 message built from the
  fields camt.053 actually exposes, not a byte-for-byte replica of any
  one bank's own historical MT940 export.
- The DATEV Buchungsstapel export always books against the bank account
  (Konto) you configure, leaves BU-Schlüssel/Gegenkonto to your own
  chart of accounts, and its header uses a defensible generic value for
  a couple of DATEV-internal metadata fields (the Buchungsstapel
  sub-format version, the origin code) that this project could not
  independently verify against DATEV's own (JavaScript-rendered)
  developer documentation.
- Both exports are built and tested against this project's own sample
  statements, not against DATEV Kanzlei-Rechnungswesen itself (which
  this project has no access to). **Test-import a small file first**
  before relying on either export for real bookkeeping, and check your
  Konto/Gegenkonto setup in DATEV before the first real import.

## Privacy

- No account, no login, no cookies for the tool itself.
- No server-side processing of your statement; the "backend" is your
  own browser's JavaScript engine.
- Analytics (Umami) records that a file was converted, not what was
  in it.
- If you're paranoid (understandable, given the subject matter),
  download the repo and open `index.html` locally with your network
  disconnected; it still works.

## Running it locally

Static files, no bundler. The only generated files are the English and
German pages: `en/index.html` and `de/index.html` are prerendered from
`index.html` plus the dictionary in `i18n.js` by `build-i18n.mjs`
(plain Node, no dependencies), so each language has its own URL for
search engines while the tool's runtime stays one code base. After
editing `index.html` or `i18n.js`, run `node build-i18n.mjs` and
commit the two folders; `node tests.mjs` fails while they are stale.

```bash
git clone https://github.com/AndryRoby/camt053-to-excel.git
cd camt053-to-excel
# any static file server works, e.g.:
npx serve .
# or just open index.html directly in a browser
```

## Reporting a missing case / wrong output

Found a bank whose camt.053 export this tool doesn't parse correctly,
a namespace/version it doesn't recognize, or a field it gets wrong?
Please open an issue on the GitHub repo with:

1. Which bank issued the statement, and which camt.053 version
   (check the `xmlns` on the `<Document>` element).
2. What the tool produced, and what it should have produced instead.

Redact anything sensitive (real IBANs, names, amounts) before posting;
issues are public.

## Disclaimer

This tool is provided **as is**, with no warranty of any kind. It
parses the `camt.053.001.02` and `camt.053.001.08` entry fields
documented above and lays them out as a table; it does not verify
that a statement is complete, that balances reconcile, or that a
particular bank's export will always match this format exactly. Tatra
banka, Slovenská sporiteľňa, VÚB, and ČSOB are not affiliated with
this tool, and their export format may change at any time. Always
check a converted table against the original statement before relying
on it for accounting or reconciliation.

## About

Built by ARLing s. r. o. (Bratislava, Slovakia).
Contact: support@arling.sk

Sibling tools from the same "Doctor" family:
- SEPA pain.001 Generátor (builds a pain.001 XML payment file from
  Excel/CSV): https://arling.sk/sepa-pain001-generator/
- SEPA pain.001 Doctor (checks a finished pain.001 file against
  bank-specific rules): https://arling.sk/sepa-pain001-doctor/
- More ARLing tools: https://arling.sk/

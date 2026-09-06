// i18n.js: SK/EN/DE dictionary and tiny rendering engine for camt.053 do
// Excelu, so the same page works for accountants in Slovakia, Germany,
// Austria and Switzerland. No framework: every visible string lives in one
// DICT object below, keyed by {sk, en, de}; index.html marks translatable
// elements with data-i18n* attributes, and applyI18n() below fills them in.
//
// Split in two halves on purpose:
//  - pure helpers (t, columnLabel, formatAmountForLang, formatDateForLang,
//    langFromLocale, detectLang's query/storage logic) never touch the DOM,
//    so tests.mjs can import and assert on them directly under Node, the
//    same way it already does for camt053.js and xlsx-writer.js.
//  - DOM-touching code (applyI18n, setLang, the bootstrap at the bottom)
//    is guarded behind `typeof document !== 'undefined'` so importing this
//    file under Node never throws.

export const LANGS = ['sk', 'en', 'de'];
export const DEFAULT_LANG = 'en';
export const STORAGE_KEY = 'arling_lang';

// ─────────────────────────────── dictionary ────────────────────────────────
// Every value has all three languages. tests.mjs asserts this exhaustively.

export const DICT = {
  // ── header / nav / language switch ────────────────────────────────────
  'skip': { sk: 'Skočiť na konverziu', en: 'Skip to conversion', de: 'Zur Umwandlung springen' },
  'wordmark': { sk: 'camt.053 do Excelu', en: 'camt.053 to Excel', de: 'camt.053 nach Excel' },
  'brand.sub': { sk: 'nástroj ARLing', en: 'an ARLing tool', de: 'ein ARLing-Tool' },
  'nav.how': { sk: 'Ako to funguje', en: 'How it works', de: 'So funktioniert es' },
  'nav.convert': { sk: 'Previesť', en: 'Convert', de: 'Umwandeln' },
  'nav.api': { sk: 'API', en: 'API', de: 'API' },
  'nav.pro': { sk: 'Pro', en: 'Pro', de: 'Pro' },
  'nav.faq': { sk: 'Otázky', en: 'FAQ', de: 'FAQ' },
  'lang.switch.aria': { sk: 'Jazyk stránky', en: 'Page language', de: 'Sprache der Seite' },
  'lang.sk.aria': { sk: 'Slovenčina', en: 'Slovak', de: 'Slowakisch' },
  'lang.en.aria': { sk: 'English', en: 'English', de: 'Englisch' },
  'lang.de.aria': { sk: 'Deutsch', en: 'German', de: 'Deutsch' },

  // ── hero ─────────────────────────────────────────────────────────────
  'hero.h1': {
    sk: 'Výpis z banky v XML (camt.053) na tabuľku za sekundu.',
    en: 'Bank statement XML (camt.053) to a spreadsheet in seconds.',
    de: 'Kontoauszug im XML-Format (camt.053) in Sekunden zur Tabelle.',
  },
  'hero.lead': {
    sk: 'Nahrajte alebo vložte výpis, ktorý ste stiahli z internet bankingu Tatra banky, SLSP, VÚB alebo ČSOB, a stiahnite tabuľku CSV alebo Excel s dátumami, sumami, protistranou a variabilným symbolom. Bez odosielania, bez účtu.',
    en: 'Upload or paste the camt.053 statement you downloaded from your bank’s internet banking, and download a CSV or Excel table with dates, amounts, the counterparty and payment reference. Nothing is uploaded, no account needed.',
    de: 'Laden Sie den camt.053-Kontoauszug aus Ihrem Online-Banking hoch oder fügen Sie ihn ein, und laden Sie eine CSV- oder Excel-Tabelle mit Datum, Betrag, Gegenpartei und Verwendungszweck herunter. Nichts wird hochgeladen, kein Konto nötig.',
  },
  'hero.cta': { sk: 'Previesť výpis', en: 'Convert a statement', de: 'Kontoauszug umwandeln' },
  'hero.source': { sk: 'Zdrojový kód na GitHube', en: 'Source code on GitHub', de: 'Quellcode auf GitHub' },
  'hero.video': { sk: 'Video, 2 minúty (po nemecky): výpis z banky do Excelu', en: 'Video, 2 minutes (in German): bank statement to Excel', de: 'Video, 2 Minuten: Kontoauszug nach Excel' },
  'hero.fact.banks': {
    sk: '4 banky: Tatra banka, SLSP, VÚB, ČSOB',
    en: 'any bank exporting camt.053 (ISO 20022)',
    de: 'jede Bank mit camt.053-Export (ISO 20022)',
  },
  'hero.fact.versions': { sk: '2 verzie camt.053: .001.02 a .001.08', en: '2 camt.053 versions: .001.02 and .001.08', de: '2 camt.053-Versionen: .001.02 und .001.08' },
  'hero.fact.tests': { sk: '546 automatizovaných testov', en: '546 automated tests', de: '546 automatisierte Tests' },
  'hero.fact.maxsize': { sk: 'max. 20 MB', en: 'max. 20 MB', de: 'max. 20 MB' },
  'hero.fact.free': { sk: 'Bez účtu, beží vo vašom prehliadači', en: 'No account, runs in your browser', de: 'Ohne Konto, läuft im Browser' },
  'hero.fact.freeFirst': { sk: 'Celá tabuľka na obrazovke zadarmo, bez limitu.', en: 'The full table on screen, free, no limit.', de: 'Die vollständige Tabelle am Bildschirm, kostenlos, ohne Limit.' },

  // ── section 01: three reasons ───────────────────────────────────────
  's1.h2': {
    sk: 'Tri dôvody, prečo účtovník potrebuje tabuľku z výpisu, nie len PDF.',
    en: 'Three reasons an accountant needs a table from the statement, not just a PDF.',
    de: 'Drei Gründe, warum eine Buchhaltung eine Tabelle aus dem Auszug braucht, nicht nur ein PDF.',
  },
  's1.sub': {
    sk: 'camt.053 XML má všetky dáta, ktoré PDF výpis len vypíše na papier. Problém je, že XML sa nedá otvoriť v Exceli dvojklikom.',
    en: 'camt.053 XML holds all the data a PDF statement only prints on paper. The problem is you can’t double-click an XML file open in Excel.',
    de: 'Die camt.053-XML enthält alle Daten, die ein PDF-Auszug nur auf Papier ausdruckt. Das Problem: Eine XML-Datei lässt sich nicht per Doppelklick in Excel öffnen.',
  },
  's1.r1.title': { sk: 'VS, ŠS a KS nie sú v jednom poli.', en: 'The payment reference isn’t in one field.', de: 'Der Verwendungszweck steht nicht in einem einzigen Feld.' },
  's1.r1.body': {
    sk: 'Banka ich zapisuje do <code>EndToEndId</code>, do štruktúrovanej referencie alebo len do textu správy, podľa toho, ako platbu zadal platiteľ. Nástroj skúša všetky tri miesta v tomto poradí, takže párovanie s faktúrou funguje aj vtedy, keď PDF ukáže len jedno dlhé referenčné číslo.',
    en: 'A bank writes it into <code>EndToEndId</code>, into a structured reference, or only into the free-text message, depending on how the payer entered it. The tool tries all three places in that order, so matching against an invoice works even when the PDF shows only one long reference number.',
    de: 'Die Bank schreibt ihn je nach Eingabe des Zahlers in <code>EndToEndId</code>, in eine strukturierte Referenz oder nur in den Freitext der Nachricht. Das Tool prüft alle drei Stellen in dieser Reihenfolge, sodass der Abgleich mit einer Rechnung auch funktioniert, wenn das PDF nur eine lange Referenznummer zeigt.',
  },
  's1.r2.title': { sk: 'Znamienko rozhoduje o súčte, nie farba v PDF.', en: 'The sign decides the total, not a colour in a PDF.', de: 'Das Vorzeichen entscheidet über die Summe, nicht eine Farbe im PDF.' },
  's1.r2.body': {
    sk: 'Suma sa do tabuľky zapíše so znamienkom podľa <code>CdtDbtInd</code> (DBIT záporná, CRDT kladná), takže sa dá priamo sčítať a skontrolovať proti začiatočnému a konečnému zostatku, ktoré výpis sám uvádza.',
    en: 'The amount is written into the table signed by <code>CdtDbtInd</code> (DBIT negative, CRDT positive), so the column can be summed directly and checked against the opening and closing balance the statement itself reports.',
    de: 'Der Betrag wird in der Tabelle mit Vorzeichen nach <code>CdtDbtInd</code> eingetragen (DBIT negativ, CRDT positiv), sodass die Spalte direkt summiert und mit dem im Auszug angegebenen Anfangs- und Endsaldo verglichen werden kann.',
  },
  's1.r3.title': { sk: 'Jeden bankový zápis môže byť viac platieb.', en: 'One bank entry can hold several payments.', de: 'Ein Sammelposten kann mehrere Zahlungen enthalten.' },
  's1.r3.body': {
    sk: 'Dávkový záznam s viacerými <code>NtryDtls/TxDtls</code> sa rozpíše na samostatný riadok pre každú platbu, takže párovanie podľa VS funguje na úrovni platby, nie na úrovni bankového zápisu.',
    en: 'A batched entry with several <code>NtryDtls/TxDtls</code> blocks is split into a separate row for each payment, so matching by reference works at the payment level, not the bank-entry level.',
    de: 'Ein Sammelposten mit mehreren <code>NtryDtls/TxDtls</code>-Blöcken wird in eine eigene Zeile je Zahlung aufgeteilt, sodass der Abgleich per Verwendungszweck auf Zahlungsebene funktioniert, nicht auf Ebene des Sammelpostens.',
  },

  // ── section 02: playground ──────────────────────────────────────────
  's2.h2': {
    sk: 'Nahrajte výpis. Dostanete tabuľku, súhrn aj kontrolu zostatku.',
    en: 'Upload a statement. You get a table, a summary, and a balance check.',
    de: 'Auszug hochladen. Sie erhalten Tabelle, Zusammenfassung und Saldoprüfung.',
  },
  's2.sub': {
    sk: 'Nič z toho, čo nahráte alebo vložíte, sa neodosiela. Parsovanie aj generovanie CSV/Excelu beží vo vašom prehliadači.',
    en: 'Nothing you upload or paste is sent anywhere. Parsing and building the CSV/Excel both run in your browser.',
    de: 'Nichts von dem, was Sie hochladen oder einfügen, wird versendet. Parsen und der Aufbau von CSV/Excel laufen im Browser.',
  },
  's2.export.hint': {
    sk: 'Prišli ste kvôli MT940/DATEV? Nahrajte výpis alebo kliknite na „ukážka“, potom na „Previesť“ a napokon na tlačidlo „MT940“ alebo „DATEV“ pri výsledku nižšie.',
    en: 'Here for MT940/DATEV? Upload a statement or click “sample”, then “Convert”, then use the “MT940” or “DATEV” button next to the result below.',
    de: 'Wegen MT940/DATEV hier? Laden Sie einen Kontoauszug hoch oder klicken Sie auf „Beispiel“, dann auf „Umwandeln“ und danach unten beim Ergebnis auf „MT940“ oder „DATEV“.',
  },
  's2.input.label': { sk: 'Vstup', en: 'Input', de: 'Eingabe' },
  's2.output.label': { sk: 'Výstup', en: 'Output', de: 'Ausgabe' },
  's2.sample.btn': { sk: 'ukážka', en: 'sample', de: 'Beispiel' },
  's2.sample.btn.title': { sk: 'Načítať malý ukážkový výpis Tatra banky s 3 položkami', en: 'Load a small sample statement from a German bank with 4 entries', de: 'Kleinen Beispiel-Kontoauszug einer deutschen Bank mit 4 Positionen laden' },
  's2.convert.btn': { sk: 'Previesť ↵', en: 'Convert ↵', de: 'Umwandeln ↵' },
  's2.files.title': { sk: 'Súbor(y) camt.053', en: 'camt.053 file(s)', de: 'camt.053-Datei(en)' },
  's2.files.hint': { sk: '.xml, max. 20 MB; viac súborov naraz', en: '.xml, max. 20 MB; several files at once', de: '.xml, max. 20 MB; mehrere Dateien gleichzeitig' },
  's2.paste.title': { sk: 'alebo vložte XML text', en: 'or paste the XML text', de: 'oder XML-Text einfügen' },
  's2.paste.placeholder': { sk: 'Sem vložte celý obsah camt.053 XML súboru.', en: 'Paste the full contents of a camt.053 XML file here.', de: 'Fügen Sie hier den vollständigen Inhalt einer camt.053-XML-Datei ein.' },
  's2.columns.title': { sk: 'Stĺpce v tabuľke aj v stiahnutí', en: 'Columns in the table and the download', de: 'Spalten in Tabelle und Download' },
  's2.csv.title': { sk: 'Formát CSV', en: 'CSV format', de: 'CSV-Format' },
  's2.csv.comma.label': { sk: 'desatinná čiarka (450,00 namiesto 450.00)', en: 'decimal comma (450,00 instead of 450.00)', de: 'Dezimalkomma (450,00 statt 450.00)' },
  's2.csv.delim.label': { sk: 'oddeľovač', en: 'delimiter', de: 'Trennzeichen' },
  's2.csv.delim.semi': { sk: 'bodkočiarka ; (Excel SK)', en: 'semicolon ; (Excel DE/SK)', de: 'Semikolon ; (Excel DE)' },
  's2.csv.delim.comma': { sk: 'čiarka ,', en: 'comma , (Excel US/UK)', de: 'Komma ,' },
  's2.history.btn': { sk: 'história', en: 'history', de: 'Verlauf' },
  's2.pro.badge': { sk: 'PRO', en: 'PRO', de: 'PRO' },
  's2.placeholder.pre': { sk: 'nahrajte alebo vložte výpis, stlačte', en: 'upload or paste a statement, press', de: 'Auszug hochladen oder einfügen, dann' },
  's2.placeholder.or': { sk: 'alebo', en: 'or', de: 'oder' },
  's2.kbd.convert': { sk: 'Previesť', en: 'Convert', de: 'Umwandeln' },
  's2.output.placeholder.full': {
    sk: 'nahrajte alebo vložte výpis, stlačte <span class="kbd">Previesť</span> alebo <span class="kbd">⌘↵</span><span class="blink"></span>',
    en: 'upload or paste a statement, press <span class="kbd">Convert</span> or <span class="kbd">⌘↵</span><span class="blink"></span>',
    de: 'Auszug hochladen oder einfügen, dann <span class="kbd">Umwandeln</span> oder <span class="kbd">⌘↵</span><span class="blink"></span>',
  },

  // ── section 03: endpoint / API ───────────────────────────────────────
  's3.h2': {
    sk: 'Dve funkcie. XML dnu, riadky von, hneď skontrolované na zostatok.',
    en: 'Two functions. XML in, rows out, checked against the balance right away.',
    de: 'Zwei Funktionen. XML hinein, Zeilen heraus, sofort gegen den Saldo geprüft.',
  },
  's3.sub': {
    sk: 'Žiadny server, žiadny API kľúč, žiadna registrácia. <code>camt053.js</code> je čistý JavaScript: prečítajte si ho, forknite, alebo ho spustite vo vlastnom kóde či CI.',
    en: 'No server, no API key, no sign-up. <code>camt053.js</code> is plain JavaScript: read it, fork it, or run it in your own code or CI.',
    de: 'Kein Server, kein API-Schlüssel, keine Anmeldung. <code>camt053.js</code> ist reines JavaScript: lesen, forken oder im eigenen Code bzw. in der CI ausführen.',
  },
  's3.codeblock.label': { sk: '0 závislostí', en: '0 dependencies', de: '0 Abhängigkeiten' },
  's3.code.comment1': { sk: 'alebo globálne: const { parse, toRows, toCsv, summarize } = window.CamtConverter;', en: 'or globally: const { parse, toRows, toCsv, summarize } = window.CamtConverter;', de: 'oder global: const { parse, toRows, toCsv, summarize } = window.CamtConverter;' },
  's3.code.comment2': { sk: 'camt.053.001.02 alebo .001.08', en: 'camt.053.001.02 or .001.08', de: 'camt.053.001.02 oder .001.08' },
  's3.code.comment3': { sk: '1 riadok = 1 platba (VS, suma, protistrana...)', en: '1 row = 1 payment (reference, amount, counterparty...)', de: '1 Zeile = 1 Zahlung (Referenz, Betrag, Gegenpartei...)' },
  's3.code.comment4': { sk: 'OPBD + pohyby =? CLBD', en: 'OPBD + movements =? CLBD', de: 'OPBD + Bewegungen =? CLBD' },
  's3.code.comment5': { sk: 'Excel bez knižnice: minimálny .xlsx ako ZIP (STORED), pozri xlsx-writer.js', en: 'Excel with no library: a minimal .xlsx as a ZIP (STORED), see xlsx-writer.js', de: 'Excel ohne Bibliothek: minimales .xlsx als ZIP (STORED), siehe xlsx-writer.js' },
  's3.copy.p1': {
    sk: 'Parsovanie XML, výpočet súhrnu aj skladanie CSV/Excelu beží presne tak, ako ho volá formulár vyššie, vo vašom prehliadači. Neexistuje backend ani API kľúč, ktorý by obsah vášho výpisu odniesol inam.',
    en: 'Parsing the XML, computing the summary, and building the CSV/Excel run exactly as the form above calls them, in your browser. There is no backend or API key that could carry your statement’s content anywhere else.',
    de: 'Das Parsen der XML, die Berechnung der Zusammenfassung und der Aufbau von CSV/Excel laufen genau so, wie das Formular oben sie aufruft, in Ihrem Browser. Es gibt kein Backend und keinen API-Schlüssel, der den Inhalt Ihres Auszugs woandershin übertragen könnte.',
  },
  's3.copy.p2': {
    sk: 'Engine používa vlastný tolerantný XML parser (rovnaký prístup ako sesterský nástroj SEPA pain.001 Doctor), nie <code>DOMParser</code>, takže beží nezmenený aj v Node: presne to spúšťa aj testovacia sada tohto projektu, <code>tests.mjs</code>.',
    en: 'The engine uses its own tolerant XML parser (the same approach as the sibling tool SEPA pain.001 Doctor), not <code>DOMParser</code>, so it also runs unmodified in Node: exactly what this project’s own test suite, <code>tests.mjs</code>, runs.',
    de: 'Die Engine verwendet einen eigenen toleranten XML-Parser (der gleiche Ansatz wie beim Schwester-Tool SEPA pain.001 Doctor), nicht <code>DOMParser</code>, und läuft daher unverändert auch in Node: genau das führt die Testsuite dieses Projekts, <code>tests.mjs</code>, aus.',
  },
  's3.copy.p3': {
    sk: 'Zadarmo, otvorené, bez reklamy a bez sledovania nad rámec anonymných počtov použitia. Prípad, ktorý nástroj spracuje zle, nahláste ako <a href="https://github.com/AndryRoby/camt053-to-excel/issues" target="_blank" rel="noopener">issue na GitHube</a>.',
    en: 'Free, open, no ads, and no tracking beyond anonymous usage counts. If the tool gets a case wrong, please report it as an <a href="https://github.com/AndryRoby/camt053-to-excel/issues" target="_blank" rel="noopener">issue on GitHub</a>.',
    de: 'Kostenlos, offen, ohne Werbung und ohne Tracking über anonyme Nutzungszahlen hinaus. Verarbeitet das Tool einen Fall falsch, melden Sie ihn bitte als <a href="https://github.com/AndryRoby/camt053-to-excel/issues" target="_blank" rel="noopener">Issue auf GitHub</a>.',
  },
  's3.copy.p4.sk': {
    sk: 'Nástroj nie je prepojený s Tatra bankou, Slovenskou sporiteľňou, VÚB ani ČSOB a nie je banka. Kontrola zostatku upozorní na nezrovnalosť, nezaručuje ale úplnosť výpisu.',
    en: 'This tool is not affiliated with any bank and is not a bank itself. The balance check flags a mismatch, but it does not guarantee a statement is complete.',
    de: 'Dieses Tool ist mit keiner Bank verbunden und ist selbst keine Bank. Die Saldoprüfung weist auf eine Abweichung hin, garantiert aber nicht, dass ein Auszug vollständig ist.',
  },

  // ── section 04: Pro ──────────────────────────────────────────────────
  's4.h2': { sk: 'Pro: MT940 a DATEV Buchungsstapel, viac súborov naraz.', en: 'Pro: MT940 and DATEV Buchungsstapel, several files at once.', de: 'Pro: MT940 und DATEV-Buchungsstapel, mehrere Dateien gleichzeitig.' },
  's4.sub': {
    sk: 'Prevod aj celá tabuľka na obrazovke sú zadarmo, bez limitu a bez konta: vidíte každý riadok, sumy aj kontrolu zostatku. Bez licencie sa stiahne prvých 20 riadkov s viditeľnou hlavičkou o ukážke. S licenciou stiahnete celý súbor v CSV aj Exceli, k tomu export do MT940 a do DATEV Buchungsstapel (EXTF CSV), spracovanie viacerých súborov naraz a históriu konverzií, jednou licenciou spoločnou pre štyri bankové nástroje ARLing.',
    en: 'The conversion and the full table on screen are free, no limit and no account: you see every row, the totals and the balance check. Without a licence the download gives you the first 20 rows with a visible sample header. With a licence you download the whole file as CSV and Excel, plus an MT940 export and a DATEV Buchungsstapel (EXTF CSV) export, processing several files at once, and a conversion history, with one licence shared across the four ARLing banking tools.',
    de: 'Die Umwandlung und die vollständige Tabelle am Bildschirm sind kostenlos, ohne Limit und ohne Konto: Sie sehen jede Zeile, die Summen und die Saldenprüfung. Ohne Lizenz enthält der Download die ersten 20 Zeilen mit sichtbarer Muster-Kopfzeile. Mit Lizenz laden Sie die vollständige Datei als CSV und Excel herunter, dazu einen MT940-Export und einen DATEV-Buchungsstapel-Export (EXTF-CSV), die gleichzeitige Verarbeitung mehrerer Dateien sowie einen Umwandlungsverlauf, mit einer Lizenz für alle vier ARLing-Banktools.',
  },
  's4.freeFirst': { sk: 'Celá tabuľka na obrazovke zadarmo, bez limitu.', en: 'The full table on screen, free, no limit.', de: 'Die vollständige Tabelle am Bildschirm, kostenlos, ohne Limit.' },
  's4.r0.title': { sk: 'MT940 a DATEV Buchungsstapel export.', en: 'MT940 and DATEV Buchungsstapel export.', de: 'MT940- und DATEV-Buchungsstapel-Export.' },
  's4.r0.body': {
    sk: 'MT940 vypadlo z pravidiel nemeckého bankového styku v novembri 2025, mnohé banky už výpis posielajú iba ako camt.053. DATEV Kanzlei-Rechnungswesen ale priamy import camt.053 nemá, ručný import súborov stále čaká MT940. Pro k tomu pridá aj export priamo do formátu DATEV Buchungsstapel (EXTF CSV).',
    en: 'MT940 was removed from German banks\' rulebook in November 2025; many banks already deliver only camt.053. DATEV Kanzlei-Rechnungswesen still has no direct camt.053 file import though: its manual file import still expects MT940. Pro also adds an export straight into the DATEV Buchungsstapel (EXTF CSV) format.',
    de: 'MT940 wurde im November 2025 aus dem Regelwerk der deutschen Kreditwirtschaft gestrichen, viele Banken liefern bereits nur noch camt.053. DATEV Kanzlei-Rechnungswesen hat aber weiterhin keinen direkten camt.053-Dateiimport, der manuelle Dateiimport erwartet nach wie vor MT940. Pro fügt zusätzlich einen Export direkt in das Format DATEV-Buchungsstapel (EXTF-CSV) hinzu.',
  },
  's4.r1.title': { sk: 'Viac súborov naraz.', en: 'Several files at once.', de: 'Mehrere Dateien gleichzeitig.' },
  's4.r1.body': {
    sk: 'Nahrajte viac výpisov naraz a spracujte ich v jednom kroku namiesto po jednom. Bez Pro sa spracuje vždy len prvý súbor.',
    en: 'Upload several statements at once and process them in one step instead of one by one. Without Pro, only the first file is processed.',
    de: 'Laden Sie mehrere Auszüge gleichzeitig hoch und verarbeiten Sie sie in einem Schritt statt einzeln. Ohne Pro wird immer nur die erste Datei verarbeitet.',
  },
  's4.r2.title': { sk: 'História konverzií.', en: 'Conversion history.', de: 'Umwandlungsverlauf.' },
  's4.r2.body': {
    sk: 'Posledné spracované výpisy (názov, dátum, banka, počet položiek) uložené vo vašom prehliadači, dostupné cez tlačidlo „história“ pri výstupe.',
    en: 'The most recently processed statements (name, date, bank, entry count) stored in your browser, available via the “history” button by the output.',
    de: 'Die zuletzt verarbeiteten Auszüge (Name, Datum, Bank, Anzahl der Positionen) werden in Ihrem Browser gespeichert und sind über die Schaltfläche „Verlauf“ bei der Ausgabe abrufbar.',
  },
  's4.r3.title': { sk: 'Prednostná podpora e-mailom.', en: 'Priority email support.', de: 'Bevorzugter E-Mail-Support.' },
  's4.r3.body': {
    sk: 'Otázka alebo prípad, ktorý si nástroj pomýlil? Odpoveď prednostne, priamo od autora nástroja.',
    en: 'A question, or a case the tool got wrong? A priority reply, directly from the tool’s author.',
    de: 'Eine Frage oder ein Fall, den das Tool falsch verarbeitet hat? Bevorzugte Antwort, direkt vom Autor des Tools.',
  },
  's4.cta.p': {
    sk: '<b>Jedna licencia pre štyri nástroje.</b> Pro pre camt.053 do Excelu sa aktivuje rovnakou licenciou ako SEPA pain.001 Doctor, SEPA pain.001 Generátor a Párovač platieb: 9&nbsp;€ mesačne alebo 79&nbsp;€ ročne pre všetky štyri nástroje, DPH v cene, faktúru pošle Stripe.',
    en: '<b>One licence for four tools.</b> Pro for camt.053 to Excel is activated by the same licence as SEPA pain.001 Doctor, SEPA pain.001 Generator and Payment matcher: €9/month or €79/year for all four tools, VAT included, Stripe sends the invoice.',
    de: '<b>Eine Lizenz für vier Tools.</b> Pro für camt.053 nach Excel wird mit derselben Lizenz aktiviert wie SEPA pain.001 Doctor, SEPA-pain.001-Generator und Zahlungsabgleich: 9&nbsp;€/Monat oder 79&nbsp;€/Jahr für alle vier Tools, inkl. MwSt., die Rechnung stellt Stripe.',
  },
  's4.buy.year.btn': { sk: 'Kúpiť Pro, 79 €/rok', en: 'Buy Pro, €79/year', de: 'Pro kaufen, 79 €/Jahr' },
  's4.buy.month.btn': { sk: 'alebo 9 €/mesiac', en: 'or €9/month', de: 'oder 9 €/Monat' },
  's4.included.list': {
    sk: '<li>Export do MT940 (.sta) a DATEV Buchungsstapel (EXTF CSV)</li><li>Viac súborov naraz, história konverzií v prehliadači</li><li>Jeden licenčný kľúč pre štyri nástroje: camt.053 do Excelu, SEPA pain.001 Doctor, SEPA pain.001 Generátor, Párovač platieb</li><li>Prednostná podpora e-mailom</li>',
    en: '<li>MT940 (.sta) and DATEV Buchungsstapel (EXTF CSV) export</li><li>Several files at once, conversion history in the browser</li><li>One licence key for four tools: camt.053 to Excel, SEPA pain.001 Doctor, SEPA pain.001 Generator, Payment matcher</li><li>Priority email support</li>',
    de: '<li>MT940-Export (.sta) und DATEV-Buchungsstapel-Export (EXTF-CSV)</li><li>Mehrere Dateien gleichzeitig, Umwandlungsverlauf im Browser</li><li>Ein Lizenzschlüssel für vier Tools: camt.053 nach Excel, SEPA pain.001 Doctor, SEPA-pain.001-Generator, Zahlungsabgleich</li><li>Bevorzugter E-Mail-Support</li>',
  },
  's4.buy.fineprint': {
    sk: 'Platba cez Stripe, DPH v cene, mesačne zrušiteľné, žiadna viazanosť: odkaz na zrušenie nájdete priamo v potvrdení platby od Stripe. Licenčný kľúč dostanete hneď po zaplatení na potvrdzovacej stránke.',
    en: 'Payment via Stripe, VAT included, cancel monthly, no minimum term: the cancellation link is right in Stripe’s payment confirmation email. You get the licence key on the confirmation page right after payment.',
    de: 'Zahlung über Stripe, inkl. MwSt., monatlich kündbar, keine Mindestlaufzeit: den Kündigungslink schickt Stripe direkt in der Zahlungsbestätigung. Den Lizenzschlüssel erhalten Sie sofort nach der Zahlung auf der Bestätigungsseite.',
  },
  's4.trust': {
    sk: 'ARLing s.&nbsp;r.&nbsp;o., Bratislava, IČ DPH SK2122352100. Prevod beží vo vašom prehliadači, výpis sa nikam neposiela. Ak sa vám MT940 alebo DATEV súbor neimportuje, napíšte s chybovou hláškou na <a href="mailto:andrej@arling.sk">andrej@arling.sk</a>, budeme to riešiť prednostne. <a href="https://arling.sk/podmienky/">Podmienky</a> · <a href="https://arling.sk/gdpr/">GDPR</a>',
    en: 'ARLing s.&nbsp;r.&nbsp;o., Bratislava, Slovakia, VAT ID SK2122352100. The conversion runs in your browser, your statement is never uploaded. If the MT940 or DATEV file does not import for you, write to <a href="mailto:andrej@arling.sk">andrej@arling.sk</a> with the error message, we prioritise fixing it. <a href="https://arling.sk/podmienky/">Terms</a> · <a href="https://arling.sk/gdpr/">Privacy (GDPR)</a>',
    de: 'ARLing s.&nbsp;r.&nbsp;o., Bratislava, Slowakei, USt-IdNr. SK2122352100. Die Umwandlung läuft in Ihrem Browser, Ihr Kontoauszug wird nicht hochgeladen. Importiert die MT940- oder DATEV-Datei bei Ihnen nicht, schreiben Sie mit der Fehlermeldung an <a href="mailto:andrej@arling.sk">andrej@arling.sk</a>, wir kümmern uns vorrangig darum. <a href="https://arling.sk/podmienky/">AGB</a> · <a href="https://arling.sk/gdpr/">Datenschutz</a>',
  },
  's4.bundle.link': { sk: 'Čo všetko je v balíku', en: 'What is in the bundle', de: 'Was im Paket enthalten ist' },
  's4.licence.manual.label': {
    sk: 'Licenčný kľúč nájdete na potvrdzovacej stránke hneď po zaplatení. Kúpili ste ho na inom počítači alebo v inom nástroji? Vložte ho sem.',
    en: 'The licence key is on the confirmation page right after payment. Bought it on another computer or in another tool? Paste it here.',
    de: 'Den Lizenzschlüssel finden Sie direkt nach der Zahlung auf der Bestätigungsseite. Auf einem anderen Computer oder in einem anderen Tool gekauft? Hier einfügen.',
  },
  's4.licence.input.placeholder': { sk: 'Licenčný kľúč (dlhý reťazec s bodkou uprostred)', en: 'Licence key (a long string with a dot in the middle)', de: 'Lizenzschlüssel (langer Text mit Punkt in der Mitte)' },
  's4.licence.activate.btn': { sk: 'Aktivovať', en: 'Activate', de: 'Aktivieren' },
  's4.licence.remove.btn': { sk: 'Odstrániť licenciu', en: 'Remove licence', de: 'Lizenz entfernen' },
  's4.sticky.text': { sk: 'Licencia Pro pre všetky štyri nástroje, mesačne zrušiteľná, žiadna viazanosť.', en: 'Pro licence for all four tools, cancel monthly, no minimum term.', de: 'Pro-Lizenz für alle vier Tools, monatlich kündbar, keine Mindestlaufzeit.' },

  // ── section 05: pricing / ask ────────────────────────────────────────
  's5.h2': { sk: 'Zadarmo. Bez limitov, natrvalo.', en: 'Free. No limits, for good.', de: 'Kostenlos. Ohne Limits, dauerhaft.' },
  's5.sub': {
    sk: 'Vznikol z vlastnej potreby: previesť výpis z banky na tabuľku bez ručného prepisovania XML. Bez účtu, bez platby, bez limitu na počet výpisov ani stiahnutí.',
    en: 'Built out of a real need: turn a bank statement into a table without retyping XML by hand. No account, no payment, no limit on statements or downloads.',
    de: 'Entstanden aus echtem Bedarf: einen Kontoauszug in eine Tabelle umwandeln, ohne XML von Hand abzutippen. Kein Konto, keine Zahlung, kein Limit bei Auszügen oder Downloads.',
  },
  's5.ask.p': {
    sk: 'Ak vám ušetrí popoludnie, napíšte, čo nástroj spracoval zle. <a href="https://github.com/AndryRoby/camt053-to-excel/issues" target="_blank" rel="noopener">Otvorte issue na GitHube</a>.',
    en: 'If it saves you an afternoon, let us know what the tool got wrong. <a href="https://github.com/AndryRoby/camt053-to-excel/issues" target="_blank" rel="noopener">Open an issue on GitHub</a>.',
    de: 'Wenn es Ihnen einen Nachmittag erspart, sagen Sie uns, was das Tool falsch gemacht hat. <a href="https://github.com/AndryRoby/camt053-to-excel/issues" target="_blank" rel="noopener">Issue auf GitHub öffnen</a>.',
  },
  's5.subscribe.p': {
    sk: '<b>Dajte mi vedieť o novom nástroji.</b> Len nové nástroje. Žiadny newsletter, žiadne zdieľanie. Odhlásenie odpoveďou na mail.',
    en: '<b>Let me know about a new tool.</b> New tools only. No newsletter, no sharing. Unsubscribe by replying to the email.',
    de: '<b>Informieren Sie mich über ein neues Tool.</b> Nur neue Tools. Kein Newsletter, keine Weitergabe. Abmeldung per Antwort auf die E-Mail.',
  },
  's5.subscribe.email.placeholder': { sk: 'vas@email.sk', en: 'you@email.com', de: 'ihre@email.de' },
  's5.subscribe.email.aria': { sk: 'E-mailová adresa', en: 'Email address', de: 'E-Mail-Adresse' },
  's5.subscribe.btn': { sk: 'Dať vedieť', en: 'Notify me', de: 'Benachrichtigen' },
  's5.subscribe.thanks': { sk: 'Ďakujeme. Ozveme sa len vtedy, keď bude niečo nové.', en: 'Thanks. We’ll only write when there’s something new.', de: 'Danke. Wir melden uns nur, wenn es etwas Neues gibt.' },
  's5.subscribe.error': {
    sk: 'Nepodarilo sa uložiť. Napíšte na <a href="mailto:andrej@arling.sk">andrej@arling.sk</a>.',
    en: 'Could not save it. Please write to <a href="mailto:andrej@arling.sk">andrej@arling.sk</a>.',
    de: 'Speichern fehlgeschlagen. Bitte schreiben Sie an <a href="mailto:andrej@arling.sk">andrej@arling.sk</a>.',
  },
  's5.subscribe.privacy': { sk: 'Súkromie', en: 'Privacy', de: 'Datenschutz' },
  's5.business.p': {
    sk: '<b>Potrebujete to pravidelne?</b> Hromadné spracovanie, párovanie s faktúrami, napojenie na účtovný systém.',
    en: '<b>Need this regularly?</b> Bulk processing, matching against invoices, a connection into your accounting system.',
    de: '<b>Brauchen Sie das regelmäßig?</b> Massenverarbeitung, Abgleich mit Rechnungen, Anbindung an Ihre Buchhaltungssoftware.',
  },
  's5.business.btn': { sk: 'Napísať, čo potrebujem', en: 'Tell us what you need', de: 'Schreiben Sie uns Ihren Bedarf' },
  's5.business.subject': { sk: 'camt.053 pre firmu', en: 'camt.053 for a business', de: 'camt.053 für ein Unternehmen' },

  // ── FAQ ──────────────────────────────────────────────────────────────
  's6.h2': { sk: 'Otázky, ktoré ľudia naozaj hľadajú.', en: 'Questions people actually search for.', de: 'Fragen, die wirklich gestellt werden.' },
  's6.sub': {
    sk: 'Priame odpovede na otázky o prevode camt.053 výpisu na tabuľku.',
    en: 'Direct answers about converting a camt.053 statement into a table.',
    de: 'Direkte Antworten zur Umwandlung eines camt.053-Auszugs in eine Tabelle.',
  },
  'faq.q1': { sk: 'Kde v internet bankingu nájdem camt.053?', en: 'Where in internet banking do I find camt.053?', de: 'Wo finde ich camt.053 im Online-Banking?' },
  'faq.a1': {
    sk: 'V Tatra banke je to položka „Výpis pre účtovníctvo“: zvoľte účet, obdobie a číslo výpisu, formát XML camt.053 je od roku 2016 jediný ponúkaný. V SLSP, VÚB aj ČSOB hľadajte v sekcii Výpisy/Dokumenty položku s názvom podobným „elektronický výpis“, „výpis pre účtovníctvo“ alebo „XML výpis“: presné pomenovanie sa medzi bankami líši, ale všetky štyri banky ponúkajú camt.053 XML ako štandardný formát výpisu pre účtovníctvo.',
    en: 'In German, Austrian and Swiss internet banking, camt.053 (ISO 20022) is the standard electronic statement export, usually next to an older MT940 option, offered by any bank, for example Sparkasse, Volksbank, Deutsche Bank, Commerzbank, Raiffeisen, Erste, UBS or PostFinance. Look under statement export, account statement download, or “camt.053 / ISO 20022”. In Slovak internet banking (Tatra banka, SLSP, VÚB, ČSOB) look for “výpis pre účtovníctvo” or “XML výpis”.',
    de: 'Im deutschen, österreichischen und Schweizer Online-Banking ist camt.053 (ISO 20022) das Standardformat für den elektronischen Kontoauszug, meist neben einer älteren MT940-Option, angeboten von jeder Bank, zum Beispiel Sparkasse, Volksbank, Deutsche Bank, Commerzbank, Raiffeisen, Erste, UBS oder PostFinance. Suchen Sie unter Kontoauszug-Export, Auszugsdownload oder „camt.053 / ISO 20022“. Im slowakischen Online-Banking (Tatra banka, SLSP, VÚB, ČSOB) heißt es „výpis pre účtovníctvo“ bzw. „XML výpis“.',
  },
  'faq.q2': { sk: 'Prečo je suma záporná?', en: 'Why is the amount negative?', de: 'Warum ist der Betrag negativ?' },
  'faq.a2': {
    sk: 'Stĺpec Suma má znamienko podľa <code>CdtDbtInd</code> v XML: DBIT (peniaze odišli z účtu) sa zapíše ako záporné číslo, CRDT (peniaze prišli) ako kladné. Vďaka tomu sa dá stĺpec priamo sčítať a porovnať so začiatočným a konečným zostatkom vo výpise, presne to robí súhrn nad tabuľkou.',
    en: 'The amount column carries a sign from <code>CdtDbtInd</code> in the XML: DBIT (money left the account) is written as a negative number, CRDT (money arrived) as a positive one. That lets the column be summed directly and compared against the statement’s own opening and closing balance, which is exactly what the summary above the table does.',
    de: 'Die Betragsspalte trägt ein Vorzeichen nach <code>CdtDbtInd</code> in der XML: DBIT (Geld hat das Konto verlassen) wird als negative Zahl geschrieben, CRDT (Geld ist eingegangen) als positive. Dadurch lässt sich die Spalte direkt summieren und mit dem im Auszug angegebenen Anfangs- und Endsaldo vergleichen, genau das macht die Zusammenfassung über der Tabelle.',
  },
  'faq.q3': { sk: 'Kde je variabilný symbol?', en: 'Where is the payment reference (VS/SS/KS)?', de: 'Wo ist der Verwendungszweck (VS/SS/KS)?' },
  'faq.a3': {
    sk: 'camt.053 nemá pre VS, ŠS ani KS samostatné pole. Banky ich zapisujú buď do <code>EndToEndId</code> v poradí <code>/VS.../SS.../KS...</code>, alebo do štruktúrovanej referencie (<code>RmtInf/Strd/CdtrRefInf/Ref</code>), alebo len do voľného textu správy (<code>RmtInf/Ustrd</code>). Nástroj skúša všetky tri miesta v tomto poradí, takže párovanie s faktúrou funguje aj vtedy, keď je vo výpise vidieť len jedno dlhé referenčné číslo. Zdroj, z ktorého sa VS/ŠS/KS zobral, uvidíte ako popis (title) nad bunkou po prejdení myšou.',
    en: 'VS/SS/KS (variabilný/špecifický/konštantný symbol) are Slovak domestic-payment reference numbers; camt.053 has no dedicated field for them. A bank writes them either into <code>EndToEndId</code> as <code>/VS.../SS.../KS...</code>, into a structured reference (<code>RmtInf/Strd/CdtrRefInf/Ref</code>), or only into the free-text message (<code>RmtInf/Ustrd</code>). The tool tries all three places in that order. For a German, Austrian or Swiss statement these columns are usually empty; use the raw message column (<code>RmtInf</code>) instead, which carries the Verwendungszweck text.',
    de: 'VS/SS/KS (variabilný/špecifický/konštantný symbol) sind slowakische Referenznummern für Inlandszahlungen; camt.053 hat dafür kein eigenes Feld. Eine Bank schreibt sie entweder in <code>EndToEndId</code> als <code>/VS.../SS.../KS...</code>, in eine strukturierte Referenz (<code>RmtInf/Strd/CdtrRefInf/Ref</code>) oder nur in den Freitext der Nachricht (<code>RmtInf/Ustrd</code>). Das Tool prüft alle drei Stellen in dieser Reihenfolge. Bei einem deutschen, österreichischen oder Schweizer Auszug bleiben diese Spalten meist leer; nutzen Sie stattdessen die Spalte mit dem rohen Nachrichtentext (<code>RmtInf</code>), die den Verwendungszweck trägt.',
  },
  'faq.q4': { sk: 'Otvorí to Excel s diakritikou?', en: 'Will Excel open it with accented characters intact?', de: 'Öffnet Excel es mit Umlauten korrekt?' },
  'faq.a4': {
    sk: 'Áno. Stiahnutý CSV má na začiatku UTF-8 BOM, ktorý aj slovenská verzia Excelu rozpozná automaticky a diakritiku zobrazí správne bez ručného výberu kódovania. Súbor .xlsx je rovno vo formáte, ktorý Excel číta natívne v UTF-8, takže diakritika je v poriadku aj tam.',
    en: 'Yes. The downloaded CSV starts with a UTF-8 BOM, which Excel (including German/Austrian/Swiss locales) recognises automatically and displays accented characters correctly without picking an encoding by hand. The .xlsx file is already in the format Excel reads natively as UTF-8, so accented characters are fine there too.',
    de: 'Ja. Die heruntergeladene CSV beginnt mit einem UTF-8-BOM, das Excel (auch in deutscher/österreichischer/Schweizer Version) automatisch erkennt und Umlaute korrekt anzeigt, ohne die Kodierung von Hand wählen zu müssen. Die .xlsx-Datei liegt ohnehin in dem Format vor, das Excel nativ als UTF-8 liest, Umlaute sind also auch dort in Ordnung.',
  },
  'faq.q5': { sk: 'Odosielajú sa moje dáta niekam?', en: 'Is my data sent anywhere?', de: 'Werden meine Daten irgendwohin gesendet?' },
  'faq.a5': {
    sk: 'Nie. Parsovanie XML aj generovanie CSV a Excelu beží celé vo vašom prehliadači, nástroj nemá backend. Jedinou sieťovou aktivitou je načítanie statických súborov stránky a anonymné počítadlo použitia (Umami), ktoré zaznamená len to, že prebehla konverzia a koľko mala riadkov, nikdy obsah výpisu (IBAN, sumy, mená).',
    en: 'No. Parsing the XML and building the CSV/Excel both run entirely in your browser; the tool has no backend. The only network activity is loading the page’s static files and an anonymous usage counter (Umami), which records only that a conversion happened and how many rows it had, never the statement’s content (IBANs, amounts, names).',
    de: 'Nein. Das Parsen der XML und der Aufbau von CSV/Excel laufen vollständig im Browser; das Tool hat kein Backend. Die einzige Netzwerkaktivität ist das Laden der statischen Seitendateien und ein anonymer Nutzungszähler (Umami), der nur erfasst, dass eine Umwandlung stattfand und wie viele Zeilen sie hatte, nie den Inhalt des Auszugs (IBAN, Beträge, Namen).',
  },
  'faq.q6': { sk: 'Aké verzie camt.053 podporujete?', en: 'Which camt.053 versions are supported?', de: 'Welche camt.053-Versionen werden unterstützt?' },
  'faq.a6': {
    sk: '<code>camt.053.001.02</code> (menný priestor <code>urn:iso:std:iso:20022:tech:xsd:camt.053.001.02</code>), ktorý ako jediný formát výpisu pre účtovníctvo posielajú Tatra banka, SLSP, VÚB aj ČSOB, a novšiu verziu <code>camt.053.001.08</code>. Nástroj menný priestor v súbore rozpozná automaticky; polia, ktoré číta (dátumy, sumy, protistrana, referencie), sú v oboch verziách rovnaké.',
    en: '<code>camt.053.001.02</code> (namespace <code>urn:iso:std:iso:20022:tech:xsd:camt.053.001.02</code>), what Tatra banka, SLSP, VÚB and ČSOB issue, and the newer <code>camt.053.001.08</code>, more common at banks in Germany, Austria and Switzerland. The tool detects the namespace in the file automatically; the fields it reads (dates, amounts, counterparty, references) are the same in both versions.',
    de: '<code>camt.053.001.02</code> (Namespace <code>urn:iso:std:iso:20022:tech:xsd:camt.053.001.02</code>), das Format von Tatra banka, SLSP, VÚB und ČSOB, sowie das neuere <code>camt.053.001.08</code>, verbreiteter bei Banken in Deutschland, Österreich und der Schweiz. Das Tool erkennt den Namespace der Datei automatisch; die gelesenen Felder (Datum, Beträge, Gegenpartei, Referenzen) sind in beiden Versionen gleich.',
  },
  'faq.q7': { sk: 'Čo dostanem v Pro?', en: 'What do I get with Pro?', de: 'Was bekomme ich mit Pro?' },
  'faq.a7': {
    sk: 'Stiahnutie celého súboru v CSV aj Exceli, export do MT940 (.sta) a do DATEV Buchungsstapel (EXTF CSV), spracovanie viacerých súborov naraz a históriu doterajších konverzií uloženú vo vašom prehliadači. Samotný prevod a celá tabuľka na obrazovke ostávajú zadarmo, bez limitu; bez licencie sa stiahne prvých 20 riadkov ako ukážka. Pro sa aktivuje jednou licenciou zo stránky <a href="https://arling.sk/bankove-nastroje/">Bankové nástroje pre účtovníkov</a>, ktorá funguje aj v SEPA pain.001 Doctor, SEPA pain.001 Generátor a Párovač platieb. Pozrite si <a href="#pro">sekciu Pro</a>.',
    en: 'The full-file download as CSV and Excel, an MT940 (.sta) export and a DATEV Buchungsstapel (EXTF CSV) export, processing several files at once, and a history of past conversions stored in your browser. The conversion itself and the full table on screen stay free, no limit; without a licence the download gives you the first 20 rows as a sample. Pro is activated by one licence from the <a href="https://arling.sk/bankove-nastroje/">Banking tools for accountants</a> page, which also works in SEPA pain.001 Doctor, SEPA pain.001 Generator and Payment matcher. See the <a href="#pro">Pro section</a>.',
    de: 'Den Download der vollständigen Datei als CSV und Excel, einen Export nach MT940 (.sta) und in das Format DATEV-Buchungsstapel (EXTF-CSV), die gleichzeitige Verarbeitung mehrerer Dateien sowie einen im Browser gespeicherten Verlauf bisheriger Umwandlungen. Die Umwandlung selbst und die vollständige Tabelle am Bildschirm bleiben kostenlos, ohne Limit; ohne Lizenz enthält der Download die ersten 20 Zeilen als Muster. Pro wird mit einer Lizenz von der Seite <a href="https://arling.sk/bankove-nastroje/">Banktools für Buchhalter</a> aktiviert, die auch in SEPA pain.001 Doctor, SEPA-pain.001-Generator und Zahlungsabgleich funktioniert. Siehe den <a href="#pro">Pro-Abschnitt</a>.',
  },
  'faq.q8': { sk: 'Prevediete camt.053 aj na MT940 pre DATEV?', en: 'Can you also convert camt.053 to MT940 for DATEV?', de: 'Wandeln Sie camt.053 auch für DATEV in MT940 um?' },
  'faq.a8': {
    sk: 'Áno, ako Pro funkciu. Formát MT940 vypadol z pravidiel nemeckého bankového styku v novembri 2025, mnohé banky už výpis pre účtovníctvo posielajú iba ako camt.053. DATEV Kanzlei-Rechnungswesen ale priamy import súboru camt.053 nemá, jeho ručný import súborov stále čaká MT940 (platený DATEV Bankdatenservice import camt.053 rieši inou cestou). Preto tento nástroj vie výpis previesť na MT940 (.sta) pre ručný import do DATEV alebo do iného staršieho softvéru, a tiež priamo do formátu DATEV Buchungsstapel (EXTF CSV). Oba exporty bežia rovnako v prehliadači, nič sa nikam neodosiela. Skontrolujte v DATEV nastavenie účtu banky (Konto) a protiúčtu (Gegenkonto) pred prvým importom a najprv vyskúšajte na malom súbore: konvencia poľa <code>:86:</code> pri MT940 sa medzi bankami mierne líši a Buchungstext/Gegenkonto v DATEV importe závisí od vášho účtovného rozvrhu.',
    en: 'Yes, as a Pro feature. MT940 was removed from the German banking industry\'s rulebook in November 2025; many banks already deliver the account statement for accounting only as camt.053. DATEV Kanzlei-Rechnungswesen still has no direct camt.053 file import though: its manual file import still expects MT940 (a paid DATEV Bankdatenservice add-on handles camt.053 a different way). So this tool can convert a statement to MT940 (.sta) for a manual DATEV import or any other legacy software, and also directly to the DATEV Buchungsstapel (EXTF CSV) format. Both exports run in the browser just like the rest of the tool: nothing is uploaded. Check your bank account (Konto) and contra account (Gegenkonto) setup in DATEV before the first import, and test with a small file first: the <code>:86:</code> field convention in MT940 varies slightly bank to bank, and the Buchungstext/Gegenkonto in a DATEV import depend on your own chart of accounts.',
    de: 'Ja, als Pro-Funktion. MT940 wurde im November 2025 aus dem Regelwerk der deutschen Kreditwirtschaft gestrichen, viele Banken liefern den Kontoauszug für die Buchhaltung bereits nur noch als camt.053. DATEV Kanzlei-Rechnungswesen hat aber weiterhin keinen direkten camt.053-Dateiimport, der manuelle Dateiimport erwartet nach wie vor MT940 (der kostenpflichtige DATEV-Bankdatenservice deckt camt.053 auf einem anderen Weg ab). Dieses Tool kann einen Auszug daher nach MT940 (.sta) für den manuellen DATEV-Import oder andere ältere Software umwandeln, und ebenso direkt in das Format DATEV-Buchungsstapel (EXTF-CSV). Beide Exporte laufen wie der Rest des Tools im Browser, nichts wird hochgeladen. Prüfen Sie vor dem ersten Import die Einstellung von Bankkonto (Konto) und Gegenkonto in DATEV, und testen Sie zuerst mit einer kleinen Datei: Die <code>:86:</code>-Feldkonvention bei MT940 unterscheidet sich leicht von Bank zu Bank, und Buchungstext/Gegenkonto beim DATEV-Import hängen von Ihrem eigenen Kontenrahmen ab.',
  },

  // ── footer ───────────────────────────────────────────────────────────
  'footer.sisters.label': { sk: 'Sesterské nástroje:', en: 'Sibling tools:', de: 'Schwester-Tools:' },
  'footer.bundle.label': { sk: 'Pro pre všetky štyri:', en: 'Pro for all four:', de: 'Pro für alle vier:' },
  'footer.all.tools': { sk: 'Všetky nástroje ARLing', en: 'All ARLing tools', de: 'Alle ARLing-Tools' },
  'footer.privacy': { sk: 'Súkromie', en: 'Privacy', de: 'Datenschutz' },
  'footer.tool.doctor': { sk: 'SEPA pain.001 Doctor', en: 'SEPA pain.001 Doctor', de: 'SEPA pain.001 Doctor' },
  'footer.tool.generator': { sk: 'SEPA pain.001 Generátor', en: 'SEPA pain.001 Generator', de: 'SEPA-pain.001-Generator' },
  'footer.tool.matcher': { sk: 'Párovač platieb', en: 'Payment matcher', de: 'Zahlungsabgleich' },
  'footer.bundle.name': { sk: 'Bankové nástroje pre účtovníkov', en: 'Banking tools for accountants', de: 'Banktools für Buchhalter' },
  'footer.country': { sk: 'Slovensko', en: 'Slovakia', de: 'Slowakei' },
  'footer.note': {
    sk: 'Nič neopúšťa váš prehliadač okrem anonymných počtov použitia cez self-hosted Umami (a e-mailu, ak sa prihlásite na odber nižšie).',
    en: 'Nothing leaves your browser except anonymous usage counts via self-hosted Umami (and an email address, if you sign up for updates below).',
    de: 'Nichts verlässt Ihren Browser außer anonymen Nutzungszahlen über das selbst gehostete Umami (und einer E-Mail-Adresse, falls Sie sich unten anmelden).',
  },

  // ── meta / SEO ───────────────────────────────────────────────────────
  'meta.title': {
    sk: 'camt.053 výpis z banky do Excelu',
    en: 'camt.053 bank statement to Excel',
    de: 'camt.053-Kontoauszug nach Excel',
  },
  'meta.description': {
    sk: 'Nahrajte alebo vložte výpis z banky vo formáte camt.053 XML (Tatra banka, SLSP, VÚB, ČSOB) a stiahnite tabuľku CSV alebo Excel. Zadarmo, priamo v prehliadači, nič sa neodosiela.',
    en: 'Upload or paste a camt.053 XML bank statement (any bank in Slovakia, Germany, Austria or Switzerland) and download a CSV or Excel table. Free, runs in your browser, nothing is uploaded.',
    de: 'Kontoauszug im camt.053-XML-Format hochladen oder einfügen (jede Bank in der Slowakei, Deutschland, Österreich oder der Schweiz) und als CSV- oder Excel-Tabelle herunterladen. Kostenlos, läuft im Browser, nichts wird hochgeladen.',
  },

  // ── dynamic JS strings (status pills, errors, dynamic labels) ──────────
  'js.status.waiting': { sk: 'čaká sa', en: 'waiting', de: 'wartet' },
  'js.status.ok': { sk: 'ok', en: 'ok', de: 'ok' },
  'js.status.balanceMismatch': { sk: 'zostatok nesedí', en: 'balance mismatch', de: 'Saldo stimmt nicht' },
  'js.status.failed': { sk: 'zlyhalo', en: 'failed', de: 'fehlgeschlagen' },
  'js.status.error': { sk: 'chyba', en: 'error', de: 'Fehler' },
  'js.status.activating': { sk: 'aktivujem…', en: 'activating…', de: 'aktiviere…' },
  'js.status.proActive': { sk: 'Pro aktívne', en: 'Pro active', de: 'Pro aktiv' },
  'js.status.noLicence': { sk: 'bez licencie', en: 'no licence', de: 'keine Lizenz' },

  'js.error.noInput': { sk: 'Nahrajte aspoň jeden .xml súbor, alebo vložte obsah camt.053 XML do textového poľa.', en: 'Upload at least one .xml file, or paste camt.053 XML content into the text field.', de: 'Laden Sie mindestens eine .xml-Datei hoch oder fügen Sie camt.053-XML-Inhalt in das Textfeld ein.' },
  'js.error.sizeExceeded': { sk: 'Súbory spolu presahujú limit {limit}. Odstráňte niektoré a skúste znova.', en: 'The files together exceed the {limit} limit. Remove some and try again.', de: 'Die Dateien überschreiten zusammen das Limit von {limit}. Entfernen Sie einige und versuchen Sie es erneut.' },
  'js.error.readFile': { sk: 'Súbor sa nepodarilo prečítať.', en: 'The file could not be read.', de: 'Die Datei konnte nicht gelesen werden.' },
  'js.error.parseFallback': { sk: 'Súbor sa nepodarilo spracovať ako camt.053 XML.', en: 'The file could not be processed as camt.053 XML.', de: 'Die Datei konnte nicht als camt.053-XML verarbeitet werden.' },
  'js.error.activationFailed': { sk: 'Aktivácia zlyhala. Skúste vložiť kľúč ručne nižšie, alebo napíšte na andrej@arling.sk.', en: 'Activation failed. Try pasting the key manually below, or write to andrej@arling.sk.', de: 'Aktivierung fehlgeschlagen. Versuchen Sie, den Schlüssel unten manuell einzufügen, oder schreiben Sie an andrej@arling.sk.' },
  'js.error.licenceKeyMissing': { sk: 'Vložte licenčný kľúč.', en: 'Paste a licence key.', de: 'Lizenzschlüssel einfügen.' },

  'js.sizeWarn': { sk: 'Súbory spolu majú {size}, limit je {limit}. Odstráňte niektoré, alebo ich prevádzajte po menších skupinách.', en: 'The files together are {size}, the limit is {limit}. Remove some, or convert them in smaller batches.', de: 'Die Dateien sind zusammen {size} groß, das Limit liegt bei {limit}. Entfernen Sie einige oder wandeln Sie in kleineren Gruppen um.' },
  'js.file.remove': { sk: 'Odstrániť', en: 'Remove', de: 'Entfernen' },
  'js.file.removeAria': { sk: 'Odstrániť {name}', en: 'Remove {name}', de: '{name} entfernen' },

  'js.notePartial': { sk: 'Spracoval sa len prvý súbor ({name}). Spracovanie viacerých súborov naraz je súčasť <a href="#pro">Pro</a>.', en: 'Only the first file ({name}) was processed. Processing several files at once is part of <a href="#pro">Pro</a>.', de: 'Es wurde nur die erste Datei verarbeitet ({name}). Die gleichzeitige Verarbeitung mehrerer Dateien ist Teil von <a href="#pro">Pro</a>.' },

  'js.copy.tsv': { sk: 'kopírovať (TSV)', en: 'copy (TSV)', de: 'kopieren (TSV)' },
  'js.copy.done': { sk: 'skopírované ✓', en: 'copied ✓', de: 'kopiert ✓' },
  'js.copy.failed': { sk: 'zlyhalo', en: 'failed', de: 'fehlgeschlagen' },
  'js.download.csv': { sk: 'Stiahnuť CSV', en: 'Download CSV', de: 'CSV herunterladen' },
  'js.download.xlsx': { sk: 'Stiahnuť Excel', en: 'Download Excel', de: 'Excel herunterladen' },
  'js.download.all': { sk: 'Stiahnuť všetky ako CSV ({n})', en: 'Download all as CSV ({n})', de: 'Alle als CSV herunterladen ({n})' },
  'js.download.mt940': { sk: 'Stiahnuť MT940 (.sta)', en: 'Download MT940 (.sta)', de: 'MT940 (.sta) herunterladen' },
  'js.download.datev': { sk: 'Stiahnuť DATEV Buchungsstapel (EXTF CSV)', en: 'Download DATEV Buchungsstapel (EXTF CSV)', de: 'DATEV Buchungsstapel (EXTF CSV) herunterladen' },
  'js.pro.export.badge': { sk: 'PRO export', en: 'PRO export', de: 'PRO-Export' },

  'js.pro.preview.title.mt940': { sk: 'Ukážka MT940 (Pro funkcia)', en: 'MT940 preview (Pro feature)', de: 'MT940-Vorschau (Pro-Funktion)' },
  'js.pro.preview.title.datev': { sk: 'Ukážka DATEV Buchungsstapel (Pro funkcia)', en: 'DATEV Buchungsstapel preview (Pro feature)', de: 'Vorschau DATEV-Buchungsstapel (Pro-Funktion)' },
  'js.pro.preview.note': {
    sk: 'Prvých {n} riadkov súboru, ktorý by sa stiahol. Celý súbor a jeho stiahnutie sú súčasť Pro (balík Bankové nástroje).',
    en: 'The first {n} lines of the file that would download. The full file, and downloading it, are part of Pro (the Banking tools bundle).',
    de: 'Die ersten {n} Zeilen der Datei, die heruntergeladen würde. Die vollständige Datei und der Download sind Teil von Pro (Paket Banktools).',
  },
  'js.pro.preview.buyPrompt': { sk: 'Odomknúť Pro:', en: 'Unlock Pro:', de: 'Pro freischalten:' },
  'js.pro.preview.close': { sk: 'Zavrieť ukážku', en: 'Close preview', de: 'Vorschau schließen' },

  'js.ukazka.preco': {
    sk: 'MT940 a DATEV Buchungsstapel sú súčasťou licencie. Tu vidíte, ako súbor vyzerá.',
    en: 'MT940 and DATEV Buchungsstapel are part of the licence. Here is what the file looks like.',
    de: 'MT940 und DATEV-Buchungsstapel gehören zur Lizenz. Hier sehen Sie, wie die Datei aussieht.',
  },
  'js.ukazka.head': {
    sk: 'Stiahli ste ukážku: {n} z {spolu} riadkov',
    en: 'You downloaded a sample: {n} of {spolu} rows',
    de: 'Sie haben ein Muster geladen: {n} von {spolu} Zeilen',
  },
  'js.ukazka.body': {
    sk: 'V súbore chýba {chyba} riadkov. Celá tabuľka je vyššie na obrazovke a ostáva zadarmo. S licenciou sa stiahne celý súbor bez hlavičky o ukážke.',
    en: 'The file is missing {chyba} rows. The full table is on screen above and stays free. With a licence the download contains the whole file, without the sample header.',
    de: 'In der Datei fehlen {chyba} Zeilen. Die vollständige Tabelle steht oben am Bildschirm und bleibt kostenlos. Mit Lizenz enthält der Download die ganze Datei, ohne Muster-Kopfzeile.',
  },
  'js.ukazka.price': {
    sk: '9 € mesačne alebo 79 € ročne za všetky štyri bankové nástroje ARLing, DPH v cene, zrušiť môžete kedykoľvek.',
    en: '9 € a month or 79 € a year for all four ARLing banking tools, VAT included, cancel any time.',
    de: '9 € im Monat oder 79 € im Jahr für alle vier ARLing-Banktools, inklusive Umsatzsteuer, jederzeit kündbar.',
  },

  's2.datev.title': { sk: 'Nastavenia DATEV Buchungsstapel (Pro)', en: 'DATEV Buchungsstapel settings (Pro)', de: 'Einstellungen DATEV-Buchungsstapel (Pro)' },
  's2.datev.hint': {
    sk: 'Platia pre export DATEV Buchungsstapel (EXTF CSV) nižšie pri výsledku. Nastavenie je zadarmo, samotné stiahnutie súboru je Pro.',
    en: 'Used by the DATEV Buchungsstapel (EXTF CSV) export below, next to the result. Setting these is free; downloading the file itself is Pro.',
    de: 'Gelten für den DATEV-Buchungsstapel-Export (EXTF-CSV) unten beim Ergebnis. Die Einstellung ist kostenlos, der Download der Datei selbst ist Pro.',
  },
  's2.datev.bankAccount.label': { sk: 'Číslo bankového účtu (Konto)', en: 'Bank account number (Konto)', de: 'Kontonummer der Bank' },
  's2.datev.accountLength.label': { sk: 'Dĺžka čísla účtu (Sachkontenlänge)', en: 'Account number length (Sachkontenlänge)', de: 'Sachkontenlänge' },
  's2.datev.advisor.label': { sk: 'Číslo poradcu (Beraternummer)', en: 'Advisor number (Beraternummer)', de: 'Beraternummer' },
  's2.datev.client.label': { sk: 'Číslo klienta (Mandantennummer)', en: 'Client number (Mandantennummer)', de: 'Mandantennummer' },
  's2.datev.fyStart.label': { sk: 'Začiatok hospodárskeho roka', en: 'Fiscal year start', de: 'Wirtschaftsjahr-Beginn' },

  'js.sample.loaded': { sk: 'ukážka vložená (výpis Tatra banky, 3 položky), stlačte {kbd1} alebo {kbd2}', en: 'sample loaded (German bank statement, 4 entries), press {kbd1} or {kbd2}', de: 'Beispiel geladen (deutscher Kontoauszug, 4 Positionen), drücken Sie {kbd1} oder {kbd2}' },

  'js.tableNote': { sk: 'Zobrazených prvých {shown} z {total} riadkov. Stiahnutie CSV/Excel obsahuje všetkých {total}.', en: 'Showing the first {shown} of {total} rows. The CSV/Excel download contains all {total}.', de: 'Angezeigt werden die ersten {shown} von {total} Zeilen. Der CSV/Excel-Download enthält alle {total}.' },

  'js.summary.statementCount': { sk: 'Počet výpisov (statements)', en: 'Number of statements', de: 'Anzahl der Auszüge' },
  'js.summary.entryCount': { sk: 'Počet položiek', en: 'Number of entries', de: 'Anzahl der Positionen' },
  'js.summary.credits': { sk: 'Príjmy (CRDT)', en: 'Credits (CRDT)', de: 'Eingänge (CRDT)' },
  'js.summary.debits': { sk: 'Výdavky (DBIT)', en: 'Debits (DBIT)', de: 'Ausgänge (DBIT)' },
  'js.summary.net': { sk: 'Čistý pohyb', en: 'Net change', de: 'Nettoveränderung' },
  'js.summary.account': { sk: 'Účet', en: 'Account', de: 'Konto' },
  'js.summary.period': { sk: 'Obdobie', en: 'Period', de: 'Zeitraum' },
  'js.summary.opening': { sk: 'Začiatočný zostatok', en: 'Opening balance', de: 'Anfangssaldo' },
  'js.summary.closing': { sk: 'Konečný zostatok', en: 'Closing balance', de: 'Endsaldo' },
  'js.summary.check': { sk: 'Kontrola zostatku', en: 'Balance check', de: 'Saldoprüfung' },
  'js.summary.checkOk': { sk: 'sedí', en: 'matches', de: 'stimmt' },
  'js.summary.checkFail': { sk: 'nesedí', en: 'does not match', de: 'stimmt nicht' },
  'js.summary.checkUnknown': { sk: 'nedá sa overiť', en: 'cannot be checked', de: 'nicht prüfbar' },
  'js.summary.entries': { sk: 'pol.', en: 'entries', de: 'Pos.' },

  'js.banner.mismatch': { sk: 'Kontrola zostatku nesedí (začiatočný zostatok + pohyby ≠ konečný zostatok podľa výpisu): {diffs}. Skontrolujte, či ide o kompletný výpis, alebo či ho výpisy nedopĺňajú.', en: 'The balance check does not match (opening balance + movements ≠ the closing balance the statement reports): {diffs}. Check whether this is the complete statement, or whether other statements complete it.', de: 'Die Saldoprüfung stimmt nicht (Anfangssaldo + Bewegungen ≠ der im Auszug angegebene Endsaldo): {diffs}. Prüfen Sie, ob es sich um den vollständigen Auszug handelt oder ob weitere Auszüge ihn ergänzen.' },
  'js.banner.missing': { sk: 'Výpis neobsahuje začiatočný aj konečný zostatok (OPBD/CLBD), kontrola sa nedala vykonať.', en: 'The statement does not include both an opening and closing balance (OPBD/CLBD), so the check could not run.', de: 'Der Auszug enthält nicht sowohl Anfangs- als auch Endsaldo (OPBD/CLBD), die Prüfung konnte nicht durchgeführt werden.' },
  'js.banner.statementFallback': { sk: 'výpis', en: 'statement', de: 'Auszug' },

  'js.licence.reason.expired': { sk: 'licencia vypršala', en: 'licence expired', de: 'Lizenz abgelaufen' },
  'js.licence.reason.signature': { sk: 'neplatný kľúč', en: 'invalid key', de: 'ungültiger Schlüssel' },
  'js.licence.reason.plan': { sk: 'kľúč pre iný produkt', en: 'key for a different product', de: 'Schlüssel für ein anderes Produkt' },
  'js.licence.reason.malformed': { sk: 'neplatný kľúč', en: 'invalid key', de: 'ungültiger Schlüssel' },
  'js.licence.reason.unsupported': { sk: 'prehliadač nepodporovaný', en: 'browser not supported', de: 'Browser nicht unterstützt' },
  'js.licence.reason.default': { sk: 'neplatná licencia', en: 'invalid licence', de: 'ungültige Lizenz' },

  'js.licence.detail.unsupported': { sk: 'Pro vyžaduje aktuálny prehliadač s podporou WebCrypto Ed25519 (Chrome, Firefox alebo Safari 17+). Aktualizujte prehliadač a skúste znova.', en: 'Pro needs a current browser with WebCrypto Ed25519 support (Chrome, Firefox or Safari 17+). Update your browser and try again.', de: 'Pro benötigt einen aktuellen Browser mit WebCrypto-Ed25519-Unterstützung (Chrome, Firefox oder Safari 17+). Aktualisieren Sie Ihren Browser und versuchen Sie es erneut.' },
  'js.licence.detail.expired': { sk: 'Táto licencia už vypršala. Kúpou novej licencie na arling.sk/bankove-nastroje ju obnovíte.', en: 'This licence has already expired. Buying a new licence at arling.sk/bankove-nastroje renews it.', de: 'Diese Lizenz ist bereits abgelaufen. Mit dem Kauf einer neuen Lizenz auf arling.sk/bankove-nastroje wird sie erneuert.' },
  'js.licence.detail.plan': { sk: 'Tento kľúč platí pre iný produkt ARLing, nie pre balík Bankové nástroje.', en: 'This key is valid for a different ARLing product, not the Banking tools bundle.', de: 'Dieser Schlüssel gilt für ein anderes ARLing-Produkt, nicht für das Banktools-Paket.' },
  'js.licence.detail.malformed': { sk: 'Kľúč sa nepodarilo prečítať, skontrolujte, či ste ho skopírovali celý.', en: 'The key could not be read. Check that you copied it in full.', de: 'Der Schlüssel konnte nicht gelesen werden. Prüfen Sie, ob Sie ihn vollständig kopiert haben.' },
  'js.licence.detail.default': { sk: 'Kľúč sa nepodarilo overiť.', en: 'The key could not be verified.', de: 'Der Schlüssel konnte nicht überprüft werden.' },
  'js.licence.validUntil': { sk: 'Licencia platná do {date}.', en: 'Licence valid until {date}.', de: 'Lizenz gültig bis {date}.' },
  'js.licence.removeConfirm': { sk: 'Odstrániť licenciu z tohto prehliadača?', en: 'Remove the licence from this browser?', de: 'Lizenz aus diesem Browser entfernen?' },

  'js.history.empty': { sk: 'Zatiaľ žiadna história. Prekonvertujte prvý výpis.', en: 'No history yet. Convert your first statement.', de: 'Noch kein Verlauf. Wandeln Sie Ihren ersten Auszug um.' },
  'js.history.rows': { sk: 'riadkov', en: 'rows', de: 'Zeilen' },
  'js.history.noName': { sk: '(bez názvu)', en: '(no name)', de: '(kein Name)' },

  'js.vs.endtoend': { sk: 'z EndToEndId (/VS.../SS.../KS...)', en: 'from EndToEndId (/VS.../SS.../KS...)', de: 'aus EndToEndId (/VS.../SS.../KS...)' },
  'js.vs.structured': { sk: 'zo štruktúrovanej referencie (RmtInf/Strd)', en: 'from the structured reference (RmtInf/Strd)', de: 'aus der strukturierten Referenz (RmtInf/Strd)' },
  'js.vs.ustrd': { sk: 'z textu správy (RmtInf/Ustrd)', en: 'from the message text (RmtInf/Ustrd)', de: 'aus dem Nachrichtentext (RmtInf/Ustrd)' },
  'js.vs.none': { sk: 'nenájdené', en: 'not found', de: 'nicht gefunden' },
  'js.bank.unknown': { sk: 'iná', en: 'other', de: 'andere' },
  'js.bank.tatrabanka': { sk: 'Tatra banka', en: 'Tatra banka', de: 'Tatra banka' },
  'js.bank.slsp': { sk: 'Slovenská sporiteľňa (SLSP)', en: 'Slovenská sporiteľňa (SLSP)', de: 'Slovenská sporiteľňa (SLSP)' },
  'js.bank.vub': { sk: 'VÚB', en: 'VÚB', de: 'VÚB' },
  'js.bank.csob': { sk: 'ČSOB', en: 'ČSOB', de: 'ČSOB' },
  'js.pastedText': { sk: 'vložený text', en: 'pasted text', de: 'eingefügter Text' },
  's2.tpl.label': { sk: 'Predloha pre softvér', en: 'Accounting software preset', de: 'Vorlage für Software' },
};

// ── column headers (VS/SS/KS carve-out per the brief: camt053.js's own
// COLUMNS array keeps its Slovak c.label as the wire default; toCsv() now
// accepts an opts.labels override so a translated header can be supplied
// without changing anything else in the engine). ─────────────────────────
export const COLUMN_LABELS = {
  statementId: { sk: 'Číslo výpisu', en: 'Statement number', de: 'Kontoauszugsnummer' },
  account: { sk: 'Účet (IBAN)', en: 'Account (IBAN)', de: 'Konto (IBAN)' },
  bookingDate: { sk: 'Dátum zaúčtovania', en: 'Booking date', de: 'Buchungsdatum' },
  valueDate: { sk: 'Dátum valuty', en: 'Value date', de: 'Valutadatum' },
  amount: { sk: 'Suma', en: 'Amount', de: 'Betrag' },
  currency: { sk: 'Mena', en: 'Currency', de: 'Währung' },
  status: { sk: 'Status', en: 'Status', de: 'Status' },
  bankRef: { sk: 'Referencia banky', en: 'Bank reference', de: 'Bankreferenz' },
  txType: { sk: 'Typ transakcie', en: 'Transaction type', de: 'Transaktionsart' },
  counterpartyName: { sk: 'Protistrana', en: 'Counterparty', de: 'Gegenpartei' },
  counterpartyIban: { sk: 'IBAN protistrany', en: 'Counterparty IBAN', de: 'IBAN der Gegenpartei' },
  counterpartyBic: { sk: 'BIC protistrany', en: 'Counterparty BIC', de: 'BIC der Gegenpartei' },
  endToEndId: { sk: 'EndToEndId', en: 'EndToEndId', de: 'EndToEndId' },
  vs: { sk: 'VS', en: 'Reference (VS/SS/KS)', de: 'Verwendungszweck / Referenz' },
  ss: { sk: 'ŠS', en: 'Reference (SS)', de: 'Verwendungszweck / Referenz (SS)' },
  ks: { sk: 'KS', en: 'Reference (KS)', de: 'Verwendungszweck / Referenz (KS)' },
  message: { sk: 'Správa pre príjemcu', en: 'Message (RmtInf free text)', de: 'Verwendungszweck-Text (RmtInf)' },
  charges: { sk: 'Poplatok', en: 'Charges', de: 'Gebühren' },
};

// Accounting-software column presets offered under the CSV export: plain
// CSV column subsets, labelled honestly as generic, not vendor-certified
// imports. The SK list stays as it was (just "generic", i.e. the current
// column selection); the DE list additionally offers DATEV/Lexware/sevDesk
// presets, since those are the DE/AT/CH tools accountants using this page
// are most likely to feed the download into. EN gets the same list as DE.
export const TEMPLATE_LABELS = {
  generic: { sk: 'Všeobecné (aktuálne stĺpce)', en: 'Generic (current columns)', de: 'Allgemein (aktuelle Spalten)' },
  datev: { en: 'DATEV (generic CSV layout)', de: 'DATEV (generisches CSV-Layout)' },
  lexware: { en: 'Lexware (generic CSV layout)', de: 'Lexware (generisches CSV-Layout)' },
  sevdesk: { en: 'sevDesk (generic CSV layout)', de: 'sevDesk (generisches CSV-Layout)' },
};

export const TEMPLATE_ORDER_BY_LANG = {
  sk: ['generic'],
  en: ['generic', 'datev', 'lexware', 'sevdesk'],
  de: ['generic', 'datev', 'lexware', 'sevdesk'],
};

// Column-key subsets for each preset (all keys must exist in COLUMN_LABELS
// above / camt053.js COLUMNS). "generic" is a signal value: callers keep
// whatever columns are currently checked rather than replacing them.
export const TEMPLATE_PRESETS = {
  generic: null,
  datev: ['bookingDate', 'valueDate', 'amount', 'currency', 'counterpartyName', 'counterpartyIban', 'message'],
  lexware: ['bookingDate', 'amount', 'currency', 'counterpartyName', 'message', 'bankRef'],
  sevdesk: ['bookingDate', 'amount', 'currency', 'counterpartyName', 'endToEndId', 'message'],
};

export function templateLabel(id, lang) {
  const l = resolveLang(lang);
  const entry = TEMPLATE_LABELS[id];
  if (!entry) return id;
  return entry[l] || entry.en || entry.sk || id;
}

export function templateOrderForLang(lang) {
  const l = resolveLang(lang);
  return TEMPLATE_ORDER_BY_LANG[l] || TEMPLATE_ORDER_BY_LANG.en;
}

// Columns ticked by default (table and download) per language. Slovak keeps
// VS/ŠS/KS on: Slovak banks fill them and Slovak bookkeeping matches
// payments by VS. German and English visitors get the counterparty IBAN
// instead: a German, Austrian or Swiss statement has no VS/ŠS/KS (they
// would be three empty columns next to the real Verwendungszweck, which
// lives in `message`), and every default column is filled in every row of
// SAMPLE_CAMT053_XML_DE, the sample those languages load (tests.mjs pins
// that). All keys must exist in COLUMN_LABELS above / camt053.js COLUMNS;
// VS/ŠS/KS stay available as toggles in every language.
export const DEFAULT_COLUMNS_BY_LANG = {
  sk: ['bookingDate', 'valueDate', 'amount', 'currency', 'counterpartyName', 'vs', 'ss', 'ks', 'message'],
  en: ['bookingDate', 'valueDate', 'amount', 'currency', 'counterpartyName', 'counterpartyIban', 'message'],
  de: ['bookingDate', 'valueDate', 'amount', 'currency', 'counterpartyName', 'counterpartyIban', 'message'],
};

export function defaultColumnsForLang(lang) {
  const l = resolveLang(lang);
  return (DEFAULT_COLUMNS_BY_LANG[l] || DEFAULT_COLUMNS_BY_LANG.en).slice();
}

// ─────────────────────────────── pure helpers ───────────────────────────────

// The page's "active" language. Every helper below that takes an optional
// `lang` argument falls back to this, NOT to DEFAULT_LANG, when `lang` is
// omitted or unrecognized: that is what lets index.html's script call
// t('some.key') (no second argument) everywhere and get the language the
// visitor actually has selected, rather than silently always getting
// English. Declared here (used by resolveLang() right below) and exported
// as read-only via getLang(); only setLang() (DOM engine section, further
// down) is allowed to change it. Stays 'en' (DEFAULT_LANG) for the whole
// process under Node (tests.mjs never calls setLang()), so every existing
// explicit-lang assertion is unaffected.
let currentLang = DEFAULT_LANG;

/** Current active language (see currentLang above). */
export function getLang() {
  return currentLang;
}

function resolveLang(lang) {
  return LANGS.includes(lang) ? lang : currentLang;
}

/** True/false without throwing on a non-string. */
function isNonEmptyString(v) {
  return typeof v === 'string' && v.trim().length > 0;
}

/** Resolves a locale tag (e.g. "de-DE", "cs-CZ", "fr-FR") to one of LANGS. */
export function langFromLocale(tag) {
  const s = String(tag || '').toLowerCase();
  if (s.startsWith('de')) return 'de';
  if (s.startsWith('sk') || s.startsWith('cs')) return 'sk';
  return DEFAULT_LANG;
}

/** Translates one dictionary key. Unknown key returns the key itself so a
 * missing translation is visible instead of silently blank. Omitting
 * `lang` uses the page's current active language (see resolveLang above). */
export function t(key, lang) {
  const l = resolveLang(lang);
  const entry = DICT[key];
  if (!entry) return key;
  return entry[l] || entry.en || entry.sk || key;
}

/** Same lookup, but with {placeholders} filled in from `vars`. */
export function tf(key, vars, lang) {
  let s = t(key, lang);
  if (vars) {
    Object.keys(vars).forEach((k) => {
      s = s.split('{' + k + '}').join(String(vars[k]));
    });
  }
  return s;
}

export function columnLabel(key, lang) {
  const l = resolveLang(lang);
  const entry = COLUMN_LABELS[key];
  if (!entry) return key;
  return entry[l] || entry.en || entry.sk || key;
}

/** {key: translatedLabel} for every column, for toCsv()'s opts.labels and
 * for the XLSX/table header row. */
export function columnLabelsMap(lang) {
  const out = {};
  Object.keys(COLUMN_LABELS).forEach((k) => { out[k] = columnLabel(k, lang); });
  return out;
}

/** "450.00" -> "450,00" for sk/de, unchanged for en. Mirrors
 * camt053.js's own formatAmountForCsv() decimalComma behaviour, just keyed
 * off the active language instead of a checkbox. */
export function formatAmountForLang(amount, lang) {
  if (amount === null || amount === undefined || !Number.isFinite(amount)) return '';
  const fixed = amount.toFixed(2);
  const l = resolveLang(lang);
  return l === 'en' ? fixed : fixed.replace('.', ',');
}

/** "2026-09-02" -> "02.09.2026" for sk/de, unchanged (already ISO) for en. */
export function formatDateForLang(iso, lang) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  if (!m) return iso || '';
  const l = resolveLang(lang);
  return l === 'en' ? `${m[1]}-${m[2]}-${m[3]}` : `${m[3]}.${m[2]}.${m[1]}`;
}

/** Default CSV export options per language, per the brief: decimal comma +
 * semicolon for sk/de, decimal point + comma for en. */
export function defaultCsvOptsForLang(lang) {
  const l = resolveLang(lang);
  return l === 'en' ? { decimalComma: false, delimiter: ',' } : { decimalComma: true, delimiter: ';' };
}

/** BCP-47 locale tag to pass to toLocaleString() for history timestamps. */
export function localeTagForLang(lang) {
  const l = resolveLang(lang);
  return l === 'sk' ? 'sk-SK' : l === 'de' ? 'de-DE' : 'en-GB';
}

export function ogLocaleForLang(lang) {
  const l = resolveLang(lang);
  return l === 'sk' ? 'sk_SK' : l === 'de' ? 'de_DE' : 'en_US';
}

/** Every DICT/COLUMN_LABELS entry has a non-empty string for every LANGS
 * member. Used both by tests.mjs and by the verify-i18n check script. */
export function findIncompleteEntries() {
  const bad = [];
  const check = (obj, prefix) => {
    Object.keys(obj).forEach((key) => {
      const entry = obj[key];
      LANGS.forEach((l) => {
        if (!isNonEmptyString(entry[l])) bad.push(`${prefix}${key}.${l}`);
      });
    });
  };
  check(DICT, '');
  check(COLUMN_LABELS, 'COLUMN_LABELS.');
  return bad;
}

/** Reads ?lang= from a query string (no DOM/location dependency), for
 * both the browser bootstrap below and tests.mjs. */
export function langFromQueryString(search) {
  try {
    const params = new URLSearchParams(search || '');
    const q = (params.get('lang') || '').toLowerCase();
    return LANGS.includes(q) ? q : null;
  } catch (e) {
    return null;
  }
}

// ─────────────────────────────── DOM engine ────────────────────────────────
// Everything below touches document/window/localStorage/navigator and only
// ever runs in a browser; every access is guarded so importing this module
// under Node (tests.mjs) is side-effect-free beyond the pure helpers above.
// currentLang/getLang() live up in the pure-helpers section (resolveLang()
// needs them); only setLang() below is allowed to change currentLang.

function readStoredLang() {
  try {
    if (typeof localStorage === 'undefined' || !localStorage) return null;
    const v = localStorage.getItem(STORAGE_KEY);
    return LANGS.includes(v) ? v : null;
  } catch (e) {
    return null;
  }
}

/** Query param wins, then localStorage, then navigator.language, then
 * DEFAULT_LANG ("en", per the brief: de -> DE, sk/cs -> SK, else EN). */
export function detectLang() {
  try {
    if (typeof location !== 'undefined') {
      const fromQuery = langFromQueryString(location.search);
      if (fromQuery) return fromQuery;
    }
  } catch (e) {}
  const stored = readStoredLang();
  if (stored) return stored;
  try {
    if (typeof navigator !== 'undefined' && navigator.language) return langFromLocale(navigator.language);
  } catch (e) {}
  return DEFAULT_LANG;
}

function setMetaByName(name, value) {
  const el = document.querySelector(`meta[name="${name}"]`);
  if (el) el.setAttribute('content', value);
}
function setMetaByProperty(prop, value) {
  const el = document.querySelector(`meta[property="${prop}"]`);
  if (el) el.setAttribute('content', value);
}

function updateUrlLang(lang) {
  try {
    if (typeof history === 'undefined' || typeof location === 'undefined') return;
    // The prerendered en/ and de/ folders (build-i18n.mjs) carry the language
    // in their path already; keep those URLs clean.
    if (document.documentElement.hasAttribute('data-lang-static')) return;
    const url = new URL(location.href);
    url.searchParams.set('lang', lang);
    history.replaceState(null, '', url.pathname + '?' + url.searchParams.toString() + url.hash);
  } catch (e) {}
}

/** Fills in every data-i18n* element and the document-level bits (title,
 * meta description/OG, <html lang>, language-switch button state) for the
 * given (already-resolved) language. Pure DOM sync, no persistence. */
export function applyI18n(lang) {
  if (typeof document === 'undefined') return;
  const l = LANGS.includes(lang) ? lang : currentLang;
  currentLang = l;

  document.documentElement.setAttribute('lang', l);

  document.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.getAttribute('data-i18n'), l); });
  document.querySelectorAll('[data-i18n-html]').forEach((el) => { el.innerHTML = t(el.getAttribute('data-i18n-html'), l); });
  document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => { el.setAttribute('placeholder', t(el.getAttribute('data-i18n-placeholder'), l)); });
  document.querySelectorAll('[data-i18n-aria-label]').forEach((el) => { el.setAttribute('aria-label', t(el.getAttribute('data-i18n-aria-label'), l)); });
  document.querySelectorAll('[data-i18n-title]').forEach((el) => { el.setAttribute('title', t(el.getAttribute('data-i18n-title'), l)); });

  document.title = t('meta.title', l);
  setMetaByName('description', t('meta.description', l));
  setMetaByProperty('og:title', t('meta.title', l));
  setMetaByProperty('og:description', t('meta.description', l));
  setMetaByProperty('og:locale', ogLocaleForLang(l));

  document.querySelectorAll('[data-set-lang]').forEach((btn) => {
    const active = btn.getAttribute('data-set-lang') === l;
    if (btn.tagName === 'A') {
      // Links to the static language folders (./, en/, de/): aria-current
      // marks the one the visitor is reading.
      if (active) btn.setAttribute('aria-current', 'true'); else btn.removeAttribute('aria-current');
    } else {
      btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    }
    btn.classList.toggle('lang-active', active);
  });

  document.querySelectorAll('form[data-subscribe]').forEach((f) => f.setAttribute('data-lang', l));

  const businessLink = document.getElementById('business-link');
  if (businessLink) {
    businessLink.href = 'mailto:andrej@arling.sk?subject=' + encodeURIComponent(t('s5.business.subject', l));
  }

  // The Pro-section "what is in the bundle" link sends visitors to the
  // bankove-nastroje bundle page in the language they are already reading.
  const bundleLink = document.getElementById('pro-bundle-link');
  if (bundleLink) bundleLink.href = 'https://arling.sk/bankove-nastroje/?lang=' + l;

  try { document.dispatchEvent(new CustomEvent('arling:langchange', { detail: { lang: l } })); } catch (e) {}
}

/** Sets the active language, persists it, syncs the URL and re-renders. */
export function setLang(lang) {
  if (!LANGS.includes(lang)) return;
  currentLang = lang;
  try { if (typeof localStorage !== 'undefined' && localStorage) localStorage.setItem(STORAGE_KEY, lang); } catch (e) {}
  applyI18n(lang);
  updateUrlLang(lang);
}

function wireLangSwitch() {
  document.querySelectorAll('[data-set-lang]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const lang = btn.getAttribute('data-set-lang');
      if (btn.tagName === 'A' && btn.getAttribute('href')) {
        // The switcher is a link to the language's own URL (./ for Slovak,
        // en/ and de/ for the prerendered folders): remember the choice so
        // the page the browser is about to load agrees, then let it navigate.
        if (!LANGS.includes(lang)) return;
        try { if (typeof localStorage !== 'undefined' && localStorage) localStorage.setItem(STORAGE_KEY, lang); } catch (e) {}
        return;
      }
      setLang(lang);
    });
  });
}

if (typeof document !== 'undefined') {
  const boot = () => {
    wireLangSwitch();
    setLang(detectLang());
  };
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
}

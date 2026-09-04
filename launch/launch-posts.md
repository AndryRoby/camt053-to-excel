# Launch: camt.053 výpis banky do Excelu

Tool: https://arling.sk/camt053-to-excel/
Repo: https://github.com/AndryRoby/camt053-to-excel
Researched: 2026-09-06 (WebSearch + WebFetch + GitHub REST/search API, unauthenticated; `gh` CLI not
available in this environment).

## 1) Live-thread research: what was actually found

Same rule as the sibling SEPA pain.001 launches: **closed (or, for an issue tracker, effectively
abandoned by its maintainer) + last human activity older than 12 months -> skip.** Beyond the
literal rule, a thread only counts as **post**-worthy if replying with this specific tool would be
real, on-topic help (someone with a Slovak-bank camt.053 statement who actually needs it as a
table), not a keyword match.

**Bottom line up front: nothing found qualifies for "post" today.** Same finding as both earlier
SEPA pain.001 launches in this niche.

| Query / source | What came back | URL | Status | Date | Verdict |
|---|---|---|---|---|---|
| GitHub issue: `raphaelm/python-sepaxml` #58, "Support camt.053" | The single closest-sounding hit in the whole search. But reading it: the requester (andrewshadura) actually wanted to build SEPA XML **from** CSV (payment generation), and only mentioned camt.053 as a tangent; the maintainer (raphaelm) replied "I don't think this library can be easily transformed from 'generating' to 'parsing'... I don't personally plan on adding that." 3 comments total, all same day. | https://github.com/raphaelm/python-sepaxml/issues/58 | Open (in name), dead in practice | Opened 2023-06-14, last comment 2023-06-15 | **skip**: 3+ years stale, maintainer declined, and the actual need expressed (CSV -> XML) is this tool's *sibling* generator's job, not this converter's |
| GitHub search API, `camt.053 is:issue is:open`, sorted by last update (13 results checked) | A mix of internal backlog/roadmap tickets inside various private-feeling accounting/ERP side projects (`resilio-tech/dolibarr-camt053-reader-and-link`, `thorstenhornung1/mietfuchs`, `rodoHasArrived/Meridian-main`, `NonoHM/budgetpilot`, `guycorbaz/kesh`, `ContextualWisdomLab/accounting-information-platform`, `splattner/openzev`). All are the maintainer's own planning/dev-task issues (0-1 comments, no public question), not someone asking for help. None names Slovakia or a Slovak bank. | n/a (13 individual issues) | Open | 2026-03 to 2026-09 | **skip (all)**: dev backlog tickets, not live questions; no audience match either |
| Issues on the closest market-context repos: `JoeggiCH/camt.053-to-CSV`, `fjacquet/camt-csv`, `erfannariman/CAMT053_Import` | Each repository shows **0 issues**, open or closed. | (3 repos) | n/a | n/a | **skip**: nothing to reply to |
| `site:porada.sk camt.053 OR výpis xml excel banka` | Real, on-topic **historical** threads exist: "Tatra banka i-výpisy" (`t-222365`), "Pohoda import výpisov SLSP do Pohody" (`t-207827`), "Humanet - import výpisov" (`t-335324`), "import z XML..." (`t-111758`). This confirms porada.sk is a genuine, recurring venue for exactly this question. Direct fetch of every one of these URLs returns **HTTP 403** to automated requests (same block hit by both earlier SEPA pain.001 launches), so freshness and whether a reply today would still land can't be verified without a human, logged-in browser session. | https://www.porada.sk/ | Unknown (403, unverifiable) | Unknown | **skip for now, flagged below**: real venue, but needs Andrej's own eyes, not an automated post |
| `site:bizforum.sk` výpis/xml/excel/účtovníctvo | Only generic accounting explainer articles (double-entry bookkeeping, VAT, free accounting software); nothing about bank-statement XML import. | https://www.bizforum.sk/ | n/a | n/a | **skip**: no matching thread |
| `site:stackoverflow.com camt.053 parse`, `site:reddit.com camt.053 bank statement excel` | Zero threads on either site; results are entirely library/product pages (RubyGems `camt`, PHP `Genkgo.CAMT`, `sepa_file_parser`, PyPI `okane`). | n/a | n/a | n/a | **skip**: no discussion thread found |

### Market context (not a thread, but relevant to positioning)

A crowded field of **generic** camt.053-to-CSV/Excel converters already exists: `JoeggiCH/camt.053-to-CSV`
and `fjacquet/camt-csv` (Python, open source), `erfannariman/CAMT053_Import` (Python/pandas),
`dpocock/camt053-xsl` (XSLT), `cstuder/camt2csv` (PHP, built for Firefly III), plus hosted converters
TreasuryHost, ZZP Pulse, BankXLSX, easybankconvert.com, and kibervarnost.si. None of them frame
themselves around a Slovak bank, mention the SBA's 2013 "SEPA XML SK štandard" by name, or pair with
a Slovak payment-generation tool the way this one pairs with SEPA pain.001 Generátor/Doctor. That's
the actual differentiation (see FACTS below), not a fact to lead a forum reply with.

### Why nothing scored "post", and what to do instead

Same shape of finding as both SEPA pain.001 launches: this is a Slovak-specific niche question that
mostly gets solved by asking an accountant, a bank's business-banking hotline, or inside a
closed/non-indexed community, not on a public, search-indexed thread with an open problem waiting
for an answer today.

**Standing watch (2 minutes, zero cost):**
- Google/Bing Alerts: `camt.053 excel Slovensko`, `"výpis z banky" XML excel prevod`, `"camt.053"
  Tatra banka OR SLSP OR VÚB OR ČSOB`
- GitHub: watch `raphaelm/python-sepaxml` issues in case a real CSV/statement-reading request
  reappears with more traction than #58 got.
- porada.sk: Andrej to open the four threads above manually (logged in) and judge whether any is
  live enough to answer; none of this research's automated checks could get past the 403.

## 2) Fakty pre majiteľa (8 bodov)

1. Nástroj číta výpisy vo formáte camt.053.001.02 aj camt.053.001.08 a každú položku (`Ntry`)
   prevedie na riadok tabuľky (dátum zaúčtovania, dátum valuty, sumu so znamienkom podľa
   `CdtDbtInd`, menu, protistranu, VS/referenciu), celé v prehliadači bez uploadu.
2. Formát vstupu nie je okrajový prípad: všetky štyri veľké slovenské banky (Tatra banka, SLSP,
   VÚB, ČSOB) generujú presne tento XML formát ako predvolený "výpis do účtovníctva". Je to
   národný štandard "SEPA XML SK štandard", ktorý vydala Slovenská banková asociácia (SBA) v
   roku 2013 na základe ISO 20022 camt.053 schémy.
3. Existuje viacero bezplatných generických camt.053 to CSV/Excel prevodníkov (napr.
   JoeggiCH/camt.053-to-CSV, erfannariman/CAMT053_Import, fjacquet/camt-csv, TreasuryHost, ZZP
   Pulse, BankXLSX, easybankconvert.com), no žiaden sa neprezentuje ako slovenský nástroj ani sa
   nespája s existujúcimi nástrojmi SEPA pain.001 Generátor/Doctor do jedného workflow ("výpis
   dnu, príkaz von").
4. Výskum dnes nenašiel žiadne živé fórum ani GitHub vlákno, kde by odpoveď s odkazom na tento
   nástroj bola skutočná, k téme sa viažuca pomoc (detaily vyššie); rovnaký záver ako pri
   predchádzajúcich dvoch launchoch SEPA pain.001 nástrojov v tejto téme.
5. Jediné na prvý pohľad blízke GitHub vlákno (`raphaelm/python-sepaxml` #58, "Support camt.053")
   je v skutočnosti opačná potreba: autor chcel generovať XML platobný príkaz z CSV, nie čítať
   camt.053 výpis do tabuľky, a správca knižnice to v roku 2023 odmietol pridať. Nehodí sa.
6. porada.sk má reálne, roky staré vlákna presne na túto tému (import bankového výpisu do
   účtovného softvéru: "Tatra banka i-výpisy", "Pohoda import výpisov SLSP do Pohody", "Humanet -
   import výpisov"), teda otázka sa tam reálne rieši. Priamy automatizovaný fetch stránky ale
   vracia HTTP 403, takže aktuálnosť a zmysel odpovede vie posúdiť len Andrej ručne, prihlásený.
7. Nástroj neoveruje úplnosť výpisu ani nezosúlaďuje zostatky s bankou: zobrazí počiatočný a
   konečný zostatok z hlavičky výpisu vedľa súčtu prevedených riadkov, aby bol prípadný nesúlad
   vidieť, ale nezisťuje jeho príčinu.
8. Reálne najlacnejší a najškálovateľnejší kanál je SEO (rovnaký záver ako pri oboch predošlých
   launchoch): kľúčové frázy nižšie, umiestnené doslovne v title/h1/FAQ, sú pasívny kanál bez
   cold e-mailu, ktorý ale zaberie týždne, nie dni, kým ho Google zaindexuje.

## 3) SK kľúčové slová do title / h1 / FAQ

Frázy, ktoré sa oplatí mať doslovne (nie len rozhádzané v texte), zoradené podľa toho, ako presne
sedia na to, čo si účtovník s týmto problémom reálne vygoogli:

- „camt.053 do Excelu“
- „camt.053 excel“
- „výpis z banky do Excelu“
- „výpis xml do excelu“
- „bankový výpis XML prevod na tabuľku“
- „výpis do účtovníctva XML“
- „camt.053 prevodník“ / „camt.053 converter“

Frázy s konkrétnou bankou (jedna z nich pravdepodobne presne sedí na to, čo hľadá čitateľ podľa
toho, kde má účet, warto ich mať aspoň vo FAQ otázkach, nie nutne v h1):

- „Tatra banka výpis XML excel“
- „VÚB výpis XML excel“
- „SLSP výpis XML excel“ / „Slovenská sporiteľňa výpis XML excel“
- „ČSOB výpis XML excel“

## Files referenced

- Bank-format facts (SBA 2013 SEPA XML SK štandard, camt.053.001.02, per-bank "výpis do
  účtovníctva" naming) are sourced fresh in this research round via WebSearch, not reused from the
  sibling SEPA pain.001 projects.
- README's "reporting a missing case" section: `../README.md`

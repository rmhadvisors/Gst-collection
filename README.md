# Collection Sheet Automation

Generates a builder **collection sheet** automatically from:

1. A scanned **Agreement for Sale** PDF (one per flat), and
2. The **Tally receipts** workbook (all receipts for the project).

It reproduces the two sections the CA maintains — the **`==` summary** (one row
per flat) and the **`Wing B` detail** (a block per flat listing every payment) —
and computes GST using the affordable-housing **45-lakh rule**.

---

## What it extracts and computes

| Item | Source | How |
| --- | --- | --- |
| Agreement value (मोबदला / consideration) | Agreement PDF, **page 1** (registration receipt) | OCR — the 2nd `रु.<amount>/-` figure (market value, **consideration**, stamp duty) |
| Agreement date | Agreement PDF, page 1 | OCR — the `dd/mm/yyyy` receipt date |
| Purchaser name(s) | Agreement PDF, **pages 6 & 7** (parties block) | OCR — tries both pages; supports `(PAN ...), aged about` format |
| Flat number | PDF **filename** (e.g. `1804-CHHAYA KADAM.pdf`, `FLAT NO.904.pdf`) | falls back to fuzzy name match against Tally |
| Payments (date + amount) | **Tally receipts** workbook | rows matched by `Flat No.<n>` |
| GST on agreement value | computed | `> 45,00,000 → 5%`, otherwise `1%` |
| GST on each advance | computed | same rate × each receipt amount |
| Balance outstanding | computed | `(agreement value + GST) − total received` |

> The agreement PDFs are fully scanned images (no text layer). The tool extracts
> the embedded full-page JPEGs directly from the PDF and runs OCR
> (`tesseract.js`, English) on **page 1** (agreement value) and **pages 6 & 7**
> (purchaser names — whichever page has the parties block).

### All flats, not just the ones with an agreement

The collection sheet now lists **every flat that appears in the Tally
receipts**, not only the flats you upload an agreement for:

- **Flats with an agreement** → full computation (agreement value, GST on the
  45-lakh rule, balance outstanding).
- **Flats without an agreement** (`tally-only`) → their receipts are listed and
  GST is computed at the **default 1%** on each receipt. Agreement value and
  balance are left blank (no agreement to compare against).

**Future-proofing the GST rate:** if a later Tally export adds a column whose
header contains "GST" (a rate or percentage), the tool reads it per receipt and
uses it instead of the 1% default. Values are interpreted as `1`/`1%` → 1%,
`5`/`5%` → 5%, `0.01` → 1%. With no such column today, it falls back to 1%.

### Trained samples

| Project | Flat | Agreement value | GST | Notes |
| --- | --- | --- | --- | --- |
| Space Homes (Wing B) | 1804 | 44,90,000 | 1% | Names on page 6; bounced cheque dedup |
| Sunray Realty | 904 | 1,23,03,000 | 5% | Names on page 7; GST payment row added via override |

See `config/samples/sunray-904-reference.json` and `config/overrides.json`.

---

## Install

```bash
npm install
```

Requires Node.js 18+ (developed/tested on Node 26).

## Usage

```bash
node src/index.js --tally "tally receipts.xlsx" --out collection-output.xlsx "1804-CHHAYA KADAM.pdf"
```

Process a whole folder of agreements at once:

```bash
node src/index.js --tally "tally receipts.xlsx" --out collection-output.xlsx ".\agreements"
```

### Options

| Flag | Description |
| --- | --- |
| `--tally <file>` | Tally receipts workbook (**required**) |
| `--sheet <name>` | Tally sheet name (default: first sheet) |
| `--out <file>` | Output workbook (default: `collection-output.xlsx`) |
| `--overrides <file>` | JSON of per-flat corrections (default: `config/overrides.json`) |
| `--keep-pages` | Keep the extracted page JPEGs for inspection |
| `-h`, `--help` | Show help |

Name each agreement PDF starting with the flat number (`1804-...pdf`) so it maps
to the correct Tally flat.

## Output

The generated workbook has three sheets:

- **`==`** — one summary row per flat (matches the CA's summary section).
- **`Wing B`** — detailed block per flat: header row + one row per receipt.
- **`Audit`** — what OCR read, the value source, the GST rate, the Tally match
  reason/score, and the candidate amounts found on page 1. Use this to spot-check.

## Fixing OCR mistakes / adding missing data

OCR is reliable for printed digits and English names but can occasionally
misread. Some columns (flat type, carpet area, registration no.) are not present
in the inputs at all. Supply corrections in `config/overrides.json`, keyed by
flat number:

```json
{
  "1804": {
    "agreementValue": 4490000,
    "name": "Mrs. Chhaya Tukaram Kadam",
    "flatType": "2BHK",
    "carpetArea": 51.17,
    "regNo": "2/91"
  }
}
```

Any field you set there overrides the value derived from the PDF/Tally.

For rows that appear in the CA collection sheet but not in Tally (e.g. a separate GST payment), use `extraPayments`. To match a CA sheet snapshot that excludes newer Tally entries, use `paymentCutoff` (e.g. `"1/31/26"`).

## Notes on reconciliation

Payment dates and amounts come **only from the Tally receipts**, as specified.
If the CA's manual sheet differs (e.g. a bounced cheque re-received in a later
month), that is a manual adjustment outside the source data and will not be
reproduced automatically — the `Audit` sheet makes the automated figures
transparent so such cases are easy to spot.

## Tests

```bash
npm test
```

Covers the GST brackets, OCR text parsing, name matching, the date timezone fix,
and an end-to-end check against the real Tally workbook (both a `1%` and a `5%`
flat).

## Project layout

```
src/
  index.js            CLI entry point
  pdf/extractPages.js extract embedded JPEG pages from a scanned PDF
  pdf/ocr.js          tesseract.js worker wrapper
  pdf/agreement.js    OCR + parse agreement value, date, names
  tally/parse.js      parse Tally receipts -> per-flat payments
  engine/collection.js GST + balance + tally matching + record builder
  output/workbook.js  write the == / Wing B / Audit sheets
  util/               text + date helpers
config/overrides.json per-flat manual corrections
test/run.js           test suite
```

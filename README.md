# Excel Web — a single-file spreadsheet that looks and behaves like Microsoft Excel

An offline Excel clone: one `index.html`, no server, no CDN, no install. Double-click
it in Chrome or Edge and it opens.

The UI is English and the ribbon follows the real Excel layout: File (Backstage), Home,
Insert, Page Layout, Formulas, Data, Review, View, Help — each tab split into named
groups with dialog launchers, a formula bar, sheet tabs, and the Excel status bar with
view buttons and a zoom slider.

## What's in it

| Tab | Groups |
|---|---|
| File | Home, New, Open, Info, Save, Save As, Print, Export, Account, Options, Close (full-window Backstage) |
| Home | Clipboard (Paste, Cut, Copy, Format Painter) · Font (name, size, grow/shrink, B/I/U, borders, fill & font colour) · Alignment (vertical ×3, horizontal ×3, orientation, indent, Wrap Text, Merge & Center) · Number (format list, currency, percent, comma, decimal ±) · Styles (Conditional Formatting, Format as Table, Cell Styles) · Cells (Insert, Delete, Format) · Editing (AutoSum, Fill, Clear, Sort & Filter, Find & Select) |
| Insert | Tables, Charts (column/line/pie), Sparklines, Links, Text, Symbols, Function |
| Page Layout | Themes, Page Setup (margins, orientation, size, print area, breaks, scale), Sheet Options (gridlines, headings, print gridlines), Arrange |
| Formulas | Function Library (financial, logical, text, date, lookup, math), Defined Names, Formula Auditing, Calculation (recalculate) |
| Data | Sort A→Z / Z→A, Filter, Text to Columns, Remove Duplicates, Data Validation, Subtotal |
| Review | Spelling, Comments, Protect Sheet, Workbook Statistics |
| View | Normal / Page Layout / Page Break Preview, Gridlines, Formula Bar, Freeze Panes, Zoom |
| Help | Help, About |

**Grid and editing**
- 200 rows × 30 columns (A–AD) per sheet, many sheets (add, rename, duplicate, delete, tab colour).
- Type straight into a cell, edit in the cell (F2 / double-click) or in the formula bar.
- Arrow keys, Tab, Enter, Home/End, PageUp/PageDown, Ctrl+A, Ctrl+G.
- Range selection with Shift+click / Shift+arrows; whole columns and rows from the headers.
- **Fill handle**: drag to copy formulas with relative references, or to extend a series (10, 20 → 30, 40, …).
- **Drag header borders** to resize columns and rows.
- Undo/redo (Ctrl+Z / Ctrl+Y, also the title-bar buttons), copy/paste with relative reference
  adjustment, cut, clear contents / formats / comments / all.
- Cell formatting: bold, italic, underline (single and double), strikethrough, font family and
  size, font and fill colour, horizontal and vertical alignment, indent, wrap text, merge,
  borders (all, outline, per side), number formats (General, Number, Currency, Accounting,
  Percent, Scientific, Short Date, Long Date, Time, Text) plus increase/decrease decimals.
- Freeze panes (rows, columns, or both), hide gridlines, hide headings, zoom 50–200 %.
- Conditional formatting: highlight greater/less/between, duplicates, top/bottom N, data bars,
  colour scales.
- Charts (column, line, pie) drawn on canvas from a range, removable.
- Print the used range only, portrait or landscape.
- Autosaves to the browser's `localStorage` ("Saved to this browser" in the status bar).

**Formulas** — own engine, ~70 functions:
`SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK MEDIAN STDEV VAR PRODUCT ABS SQRT POWER MOD INT
ROUND ROUNDUP ROUNDDOWN CEILING FLOOR EXP LN LOG10 SIN COS TAN ASIN ACOS ATAN
IF IFERROR AND OR NOT TRUE FALSE CONCAT CONCATENATE LEFT RIGHT MID LEN UPPER LOWER PROPER TRIM
SUBSTITUTE FIND SEARCH REPT TEXT VALUE TODAY NOW YEAR MONTH DAY DATE WEEKDAY
RAND RANDBETWEEN VLOOKUP HLOOKUP INDEX MATCH SUMIF COUNTIF AVERAGEIF PMT FV ISNUMBER ISTEXT ISBLANK ISERROR NA`

Supports `+ - * / ^ &`, parentheses, comparisons, `A1` references, absolute `$A$1`, ranges
`A1:B9`, and either `,` or `;` as the argument separator.

**Files**
- Open `.xlsx`, `.xls`, `.csv`, `.ods` (multi-sheet, formulas and values).
- Save `.xlsx` (formulas and values are written; see the limits below) or `.csv`.
- Dragging a file onto the window also works.

## Rebuilding

```
python tools/build.py
```

Concatenates `src/` in this order —
`head.html → 01-state → 02-ribbon → 03-core → 04-formula → 05-features → 06-tools → 08-extras → 07-io`
(`08-extras` must come before `07-io` because `boot()` reads its module-level flags) — then
inlines `xlsx.full.min.js` from `tracking-klaim-ui/vendor` and closes `</script></body></html>`.

## Verified (Chrome, 1440×900)

- 6,000 cells render on init; every control on all nine tabs was clicked in sequence → 0 errors.
- Formulas: `=B2*C2`→7500, `=SUM(D2:D3)`→57500, `=IF(D4>20000,"High","Low")`→High,
  `=ROUND(AVERAGE(C2:C3),0)`→3750, `=COUNTIF(C2:C3,">3000")`→1, `=VLOOKUP("Pen",A2:D3,3)`→2500.
- Number formats render US-style: `1,200.00`, `$3,600.00`, `12.50%`, `1.50E+03`.
- Fill handle: series 10, 20 → 30, 40, 50; `=A1*2` copied to `=A2*2`, `=A3*2`.
- Copy/paste shifts references relative to the target; undo restores the previous state.
- Conditional formatting: data bars scale per value, greater-than rules recolour matching cells.
- Ribbon fits one row at every width from 1147 px up: 1235 px of content at 1147 px (the Editing
  group clips at the window edge, exactly as Excel does), 1218 px at 1366 px, 1314 px at 1440 px.
  Labels are never dropped — narrow windows scroll the strip instead of shrinking buttons to icons.
- `.xlsx` export validated with `zipfile` + an XML parser: `D2` carries `f=B2*C2` with `v=15000`,
  `D4` carries `f=SUM(D2:D3)` with `v=40000`; re-importing restores the formulas.
- Print CSS hides the ribbon, tabs, formula bar and status bar; only the grid prints.

## Multi-workbook (berbilang buku kerja)

Lebih daripada satu buku kerja boleh dibuka serentak, masing-masing dengan nama dan data sendiri.

- Bar buku kerja di bawah tab ribbon: senarai semua buku terbuka, titik merah menandakan belum
  disimpan, `✕` menutup, `＋` membuka buku baru. Klik dua kali untuk menamakan semula.
- Menukar buku melalui `switchWB()`; sejarah undo/redo setiap buku berasingan.
- Semua buku disimpan berasingan dalam `localStorage` (`excel-web-workbooks-v1`) dan dipulihkan
  selepas dimuat semula.

### Mod samping (side by side)

Butang `▥` membahagi tetingkap kepada dua panel, setiap panel memaparkan buku kerja yang berlainan
serentak — pada tablet ini benar-benar dua fail kelihatan bersamaan. Garis pemisah boleh diseret
(seret dengan jari pun boleh), dan setiap panel memaparkan nilai yang telah dikira (bukan teks
rumus).

### Mengenai "dua tetingkap Android"

APK yang dibungkus sebagai TWA mempunyai **satu** activity, jadi JavaScript di dalamnya tidak boleh
membuka tetingkap Android kedua — itu memerlukan kod Java/Kotlin asli. Yang disediakan:

1. **Mod samping** di atas (dua buku dalam satu aplikasi).
2. Butang `↗` "Open in new window" — membuka tetingkap/ tab pelayar sebenar di mana platform
   membenarkan; jika disekat, ia beralih kepada mod samping secara automatik.
3. **Split-screen Android asli** masih boleh digunakan: buka aplikasi, kemudian guna gesture
   split-screen sistem (Samsung/OnePlus/Xiaomi ada) — itu dua contoh aplikasi, masing-masing
   dengan storannya sendiri.

## Measured against real Excel 2016

A screenshot of Excel 2016 (Office16, 150 % DPI) was captured and pixel-measured; the clone was
rendered at the same window size and device scale factor for a like-for-like comparison.

| Band | Real Excel | Clone |
| --- | --- | --- |
| Title bar | 48 px | 48 px |
| Tab strip | 29 px | 29 px |
| Ribbon | 107 px | 105 px |
| Formula bar | 22 px | 21 px |
| Column header | 20 px | 19 px |
| Title bar colour | `#107c41` | `#107c41` |
| Gridline colour | `#e0e0e0` | `#e0e0e0` |
| Header fill | `#f0f0f0` | `#f0f0f0` |

Group order, row structure inside each group (Clipboard = big Paste + icon column, Font = 2 rows,
Alignment = 3 rows ending in Wrap Text / Merge & Center, Number = format box over 5 icon buttons,
Styles / Cells / Editing = big labelled buttons with trailing chevrons) and the formula-bar order
(Name Box → `fx` → ✕ → ✓ → input) all match.

## Known differences from Excel

- **Icons are line art, not Microsoft's multi-tone MDL2 set.** They are tinted to Excel's palette
  (blue structure, red delete, orange painter) but are not the original bitmaps.
- **No Draw or Automate tab.** Every tab that exists has working commands; a Draw tab without real
  ink tools would be decoration, so it is left out rather than faked.
- **No Clipboard task pane, no Quick Analysis, no Tell Me command list beyond the built-in search.**
- Dialog launchers appear on Font, Alignment and Number only; Excel also shows one on Clipboard,
  which opens a pane this app does not have.

## Android (PWA + APK)

Fail sedia PWA ada di repo: `manifest.webmanifest`, `sw.js` (cache-first, berfungsi luar talian),
dan `icons/` (192, 512, maskable 512). `.well-known/assetlinks.json` + `.nojekyll` disediakan supaya
TWA dapat disahkan dan berjalan tanpa bar alamat.

Sentuhan Android: sasaran ketik diperbesar (`@media (pointer:coarse)`), ribbon menggelongsor
mendatar pada telefon dan tidak berpecah menjadi dua baris, dan padding `env(safe-area-inset-*)` serta
`100dvh` digunakan supaya sesuai dengan skrin bertakuk dan bar navigasi.

URL Pages: https://vanjamkrindo.github.io/excel-web/

### Membina APK

```bash
python tools/retry_apk.py      # enqueue -> poll -> muat turun ke apk-out/package.zip
```

ZIP mengandungi `.apk`, `.aab`, `signing.keystore` dan `assetlinks.json`. Salin cap jari SHA-256
daripada `signing-key-info.txt` ke dalam `.well-known/assetlinks.json`, hantar, tunggu ia menjawab
200, kemudian pasang APK.

## Honest limits

- **Cell styling is not written into `.xlsx`.** The inlined SheetJS Community Edition does not
  emit `styles.xml` on export — bold, colours, borders and number formats survive in the app and
  in `localStorage`, but are lost when the file is opened in Excel. Values and formulas are fine.
- Sheets are fixed at 200 × 30 (not 1,048,576 rows).
- Functions the engine doesn't know are kept as their cached value from the imported file.
- No PivotTables, structured Tables, macros or add-ins. "Format as Table" applies the visual
  style, not a real ListObject.
- Charts are drawn on the app's canvas; they are not exported to `.xlsx`.
- Filter hides rows rather than querying data; sort acts on the active column.
- Autosave is browser-only (`localStorage`), never back to the original file on disk.
- Sheet protection prevents editing in the app; it is not encryption.

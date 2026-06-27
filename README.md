# AutoDryLedger

> A Python desktop tool that automates bulk dry goods ledger generation by parsing supplier bills into pre-built Excel templates. Built with **Tkinter** and **openpyxl**.

## Overview

Managing bulk dry goods (grain, oil, dry staples) ledgers by hand is slow and error-prone. This tool turns a supplier's bill file into fully formatted daily ledgers in seconds:

1. Select a **ledger template** workbook with a master sheet layout.
2. Select a **supplier bill** workbook (`.xls` or `.xlsx`).
3. Click **Run** — it creates one sheet per date, fills items, adjusts rows, copies formatting, and cleans up unused sheets.

## Features

- **One-click workflow** — pick template, bill, and output path; the rest is automatic.
- **`.xls` → `.xlsx` conversion** — handles legacy Excel files via `win32com` (Windows) or `xlrd` fallback.
- **Master sheet cloning** — duplicates the template's master sheet for each date, preserving all formatting (fonts, borders, fills, number formats, alignment, row heights).
- **Dynamic row adjustment** — expands or shrinks the data area to match the number of items per date (default template has 20 data rows).
- **Auto-fill totals** — computes the subtotal amount per sheet and writes it to the total row.
- **Clean output** — removes the master template sheet and any non-date sheets, leaving only the daily ledgers.
- **Live log** — real-time processing feedback in the GUI.

## Installation

```bash
git clone https://github.com/luckyfishes/auto-dry-ledger.git
cd auto-dry-ledger
pip install -r requirements.txt
```

### Requirements

- Python 3.8+
- `openpyxl`
- `xlrd` (for `.xls` parsing and conversion fallback)
- `pywin32` (optional — recommended on Windows for better `.xls` → `.xlsx` conversion via Excel COM)

> **Note:** `pywin32` produces the most faithful `.xls` to `.xlsx` conversion. Without it, the tool falls back to `xlrd`, which preserves cell values but may lose some formatting.

## Usage

```bash
python fill_grain_ledger_gui.py
```

1. **Ledger Template** — select an `.xls` or `.xlsx` template with a master sheet layout.
2. **Bill File** — select the supplier's bill (`.xls` or `.xlsx`).
3. **Output Path** — choose where to save the generated ledger. Auto-derives from the bill file name if left empty.
4. Click **Start Processing**.

## How It Works

1. **Convert Template** — if the template is `.xls`, converts it to a temporary `.xlsx` first.
2. **Parse Bill** — reads the supplier bill sheet-by-sheet, grouping items by date.
3. **Copy Template** — copies the converted template to the output path.
4. **Clone Master Sheets** — for each date found in the bill, clones the master sheet and renames it to the date (`YYYY-MM-DD`).
5. **Adjust Rows** —
   - If more than 20 items: inserts rows and copies the formatting from the reference row.
   - If fewer than 20 items: deletes extra rows and copies the total-row formatting.
6. **Fill Data** — writes item number, name, unit, quantity, price, and amount into each row.
7. **Write Total** — writes the subtotal amount to the total row and clears other cells in that row.
8. **Clean Up** — deletes the master template sheet and any sheets that don't match a `YYYY-MM-DD` pattern.
9. **Save** — writes the final workbook to the chosen output path.

## Template Structure

Your ledger template should be an `.xls` or `.xlsx` file with a **single master sheet** (e.g., named `Sheet1`). The tool will clone this sheet for each date.

### Layout Rules

| Row | Content | Columns |
|-----|---------|---------|
| 2 | Header date field | Column F: `时间：YYYY.MM.DD` (auto-filled by tool) |
| 4 | First data row | A: `序号`, B: `品名`, C: `单位`, D: `数量`, E: `单价`, F: `金额`, G: `备注` |
| 4–23 | Data area (20 rows) | Filled by tool; expands/shrinks dynamically |
| 24 | Total row | B: `小计金额`, F: subtotal amount (auto-computed) |

> The tool assumes **20 data rows** (rows 4–23) as the default. If your template uses a different number, adjust the constants in `process()` accordingly.

## Bill File Structure

The supplier's bill must be an `.xls` or `.xlsx` file with a single sheet containing:

- **Date rows** — Column A contains a date in `YYYY-MM-DD` format. This triggers a new daily group.
- **Item rows** — After a date row, each row represents one item with:
  - Column A: `序号` (sequence number, integer > 0)
  - Column B: `品名` (item name)
  - Column C: `单位` (unit)
  - Column D: `数量` (quantity)
  - Column E: (optional / unused)
  - Column F: `单价` (unit price)
  - Column G: `金额` (total amount)

Rows with values like `金额`, `总计`, `小计`, or empty rows are automatically skipped.

## Example

**Input Bill:**

| | A | B | C | D | F | G |
|---|---|---|---|---|---|---|
| ... | | | | | | |
| 10 | 2026-01-05 | | | | | |
| 11 | 1 | 大米 | 袋 | 50 | 120.00 | 6000.00 |
| 12 | 2 | 食用油 | 桶 | 20 | 85.00 | 1700.00 |
| 13 | 3 | 生抽 | 瓶 | 100 | 8.50 | 850.00 |
| 14 | | | | | | |
| 15 | 2026-01-08 | | | | | |
| 16 | 1 | 面粉 | 袋 | 30 | 65.00 | 1950.00 |
| ... | | | | | | |

**Output:**

```
输出结果.xlsx
├── 2026-01-05  (sheet with 3 items + total)
├── 2026-01-08  (sheet with 1 item + total)
└── (master sheet removed)
```

## Screenshot

```
┌──────────────────────────────────────────────┐
│      粮油干货台账自动填充                      │
│                                              │
│  台账模板: [____________________] [选择文件]  │
│  对账单:   [____________________] [选择文件]  │
│  输出路径: [____________________] [选择路径]  │
│                                              │
│        [      开始处理      ]                │
│  ------------------------------------------- │
│  [INFO] 正在转换模板...                      │
│  [INFO] 正在解析对账单...                    │
│  共解析到 2 个日期: ['2026-01-05', '2026-01-08'] │
│    2026-01-05: 3 条                          │
│    2026-01-08: 1 条                          │
│  [INFO] 已复制模板 → 输出结果.xlsx            │
│    Sheet 2026-01-05: 3 条，小计 8550.00       │
│    Sheet 2026-01-08: 1 条，小计 1950.00       │
│  完成！                                      │
│  共生成 2 个日期 sheet                         │
│  输出文件: 输出结果.xlsx                       │
└──────────────────────────────────────────────┘
```

## License

MIT License — see [LICENSE](LICENSE) for details.

## Contributing

Issues and PRs are welcome. Please open an issue first to discuss major changes.

## Author

- **liyatnok** — created for automating bulk dry goods ledger workflows at a school/institution cafeteria.

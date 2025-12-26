# Excel VBA — Filter / Aggregate / Export / Email Automation

This VBA module automates a simple reporting workflow inside Excel:

1. **Clear** previous data from staging sheets  
2. **Filter + copy** rows for a selected category (e.g., `XXXX`)  
3. **Derive duration metrics** (seconds → hours) and compute totals  
4. **Export** the result sheet to a standalone `.xlsx` file  
5. **Prepare an Outlook email** with the exported file attached

> Note: The sample code uses placeholder names such as `XXXX`. Replace them with your real sheet names / categories.

---

## ✨ What it does (high level)

### `Clear`
- Clears content in:
  - `Data` sheet range `A1:Z60000`
  - `XXXX` sheet range `A1:Z60000`

### `Sort`
- Enables AutoFilter on the `Data` sheet
- Filters by **Field 1** (column A) where value equals `XXXX`
- Copies the filtered block into the `XXXX` sheet
- Adds calculated columns:
  - `Duration_h` = `(D / 60) / 60`
  - `Duration_cut_off_h` = `(E / 60) / 60`
- Converts formulas to values and pastes them into columns **D** and **E**
- Computes totals:
  - `TOTAL Duration_h` = `SUM(D:D)`
  - `TOTAL Duration_cut_off_h` = `SUM(E:E)`

### `SaveMacro` / `XXXX_Save_Macro`
- Creates a new workbook
- Copies the `XXXX` sheet into it
- Saves it as: `XXXX_data.xlsx` in the same folder as the macro workbook

### `sendTo`
- Opens Outlook and creates an email draft:
  - **To**: taken from `Mailing_lists` sheet at a given cell (e.g., `"C2"`)
  - **Subject**: `Daily Data_<date> <varFC>`
  - **Attachment**: `<workbook path>\<varFC>_.xlsx`
  - **Body**: includes a short message + FAQ

### `Send_FR`
- Example “wrapper macro” calling `sendTo "XXXX", "C2"`

---

## 🧩 Requirements

- Microsoft Excel (Desktop) with **macros enabled**
- Outlook (Desktop) configured on the same machine (because `CreateObject("Outlook.Application")`)
- Worksheets expected by the macro:
  - `Data`
  - `XXXX` (replace with your real target sheet)
  - `General` (expects a date in cell `A1`)
  - `Mailing_lists` (expects recipient list in a chosen cell, e.g. `C2`)

---

## 🚀 How to run

1. Open the `.xlsm` workbook
2. Enable macros
3. Run in this order:
   1. `Clear`
   2. `Sort`
   3. `SaveMacro`
   4. `Send_FR` (or your own send wrapper)

---

## 🔧 Customization checklist

- Replace `XXXX` everywhere with your real:
  - category value in the filter
  - target sheet name
  - output file prefix
- Adjust range limits:
  - `A1:Z60000` (clearing)
  - `$A$1:$Z$50000` (filtering)
  - `Z2:Z1000` / `AA2:AA1000` (formulas)
- Confirm which columns contain the raw duration fields:
  - current code assumes raw durations are in **D** and **E** (seconds)

---

## ⚠️ Notes / known issues in the sample snippet

Your pasted snippet contains a couple of typical copy/paste issues to watch for:

- `ThisWorkbook.Sheets"XXXX").Copy ...`  
  should be `ThisWorkbook.Sheets("XXXX").Copy ...`
- Attachment name uses `varFC & "_.xlsx"` but the exporter saves `XXXX_data.xlsx`.  
  Make sure naming is consistent (either export to `<varFC>_.xlsx` or attach `XXXX_data.xlsx`).

These are easy fixes and worth polishing before you show this repo publicly.

---

## 📄 License

MIT (or your preferred license).

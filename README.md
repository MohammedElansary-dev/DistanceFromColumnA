# 📏 Column Distance Status Tracker

Displays the distance (column number) from Column A in the Excel status bar whenever a user selects a cell.

---

## 📌 Overview

This Excel VBA script enhances your worksheet navigation by showing how far the currently selected cell is from **Column A** — directly in the Excel **status bar** (the bar at the bottom of the Excel window).

Useful for:

* 🧾 Quickly identifying how deep into the sheet a value lies
* 🔍 Debugging templates or formulas involving column offsets
* 📐 Understanding horizontal structure in wide sheets

---

## ⚙️ How It Works

Every time the user changes selection in any sheet:

1. The macro calculates the column number of the selected cell.
2. It shows this distance as: `"Distance from Column A: X"` in the status bar.
3. If something goes wrong, it clears the status bar to prevent clutter.


![image](https://github.com/user-attachments/assets/f887ec9d-2811-4d48-91c0-42e095b39cc5)

---

## 📂 Setup Instructions

1. Open your Excel file.
2. Press **Alt + F11** to launch the **VBA Editor**.
3. In the **Project Explorer**, double-click `ThisWorkbook` under `Microsoft Excel Objects`.

   * 🟡 *This script must go into `ThisWorkbook` and **not** a sheet or standard module because it listens for changes across all sheets.*
4. Paste the VBA code into the `ThisWorkbook` module.
5. Save your file as **.xlsm**.

---

## 🔧 Customization

| Variable        | Description                                        | Default         |
| --------------- | -------------------------------------------------- | --------------- |
| `Target.Column` | Measures distance from column A using column index | `Target.Column` |

➡️ Want to measure **row** distance instead?
Replace this line:

```vb
Dim dist As Long: dist = Target.Column
```

with:

```vb
Dim dist As Long: dist = Target.Row
```

---

## ❗ Notes

* Only the **active cell** in a selection is used.
* Works across **all worksheets** in the workbook.
* Clears the status bar if something fails or if selection is invalid.

---

## 🧠 Use Case Ideas

* Spreadsheet QA and debugging
* Column mapping for imports
* Orientation in large datasets

---

## 📄 License

MIT License — use freely, contribute back if helpful 💙

---

## 👏 Author

Created by Mohamed El-ansary. This tool was built to simplify Excel workflows and boost daily productivity.

---


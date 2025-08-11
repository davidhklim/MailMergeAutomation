# MailMergeAutomation

# 📄 MailMergeAutomation

This guide explains how to set up a **Microsoft Word Mail Merge** and use **VBA automation** to automatically split the merged document into **individual files**, saving them with custom names in specified folders.

> ⚠ **Important:**  
> The VBA code must be run as the final step. Do **not** manually complete the mail merge — the VBA code is designed to handle both merging and splitting in one go.

---

## ✨ Features

- Automates the mail merge process in Word.
- Splits merged output into **individual documents**.
- Saves each file using a custom naming convention from your Excel data source.
- Allows specifying the destination folder for each document.

---

## 🛠 Part 1: Setting Up the Mail Merge

### **Step 1: Prepare Your Data Source (Excel Workbook)**

1. Open **Microsoft Excel** and create a workbook with the data you want to use.  
   - The **first row** should contain **column headers** (e.g., `Entity`, `Authorized Signatory`, `Address`, etc.).
   - Add the data for each recipient in rows below the headers (“Variables”).
2. Add **two additional columns**:
   - `DocFolder` → Full local path where the split files will be saved.
   - `FileName` → Naming convention for each output file.  
     Example formula:  
     ```excel
     ="ThisContract (" & A2 & ")"
     ```
3. Save the workbook locally or in iManage.

---

### **Step 2: Create the Mail Merge Template (Word Document)**

1. Open your Word template.
2. Go to **Mailings > Select Recipients**:
   - If using iManage: `Select from iManage`.
   - If using Excel: `Use an Existing List` and select your workbook.
3. When prompted, select the sheet that contains your variables.
4. Insert merge fields where needed: **Mailings > Insert Merge Field**.
5. (Optional) Format merge fields:  
   - Press `Alt + F9` to toggle field codes.  
   - Add formatting switches, e.g.:  
     ```
     \# #,##0
     ```
   - Press `Alt + F9` again to toggle back.
6. Preview results via **Mailings > Preview Results**.
7. Save the document.

---

## 🤖 Part 2: Automating with VBA

### **Step 1: Save as Macro-Enabled Document**
1. Save your mail merge template as a **.docm** file:  
   `File > Save As > Word Macro-Enabled Document (*.docm)`

---

### **Step 2: Insert the VBA Code**
1. Press `Alt + F11` in Word to open the VBA editor.
2. Go to **Insert > Module**.
3. Paste in the VBA macro code (below).
4. Save your macro-enabled file.

---

## 💻 VBA Macro Code

```vba
Sub DocAndPdfMailMergeDoLoop()

    Dim MasterDoc As Document, SingleMergeDoc As Document, LastRecordNum As Integer
    Set MasterDoc = ActiveDocument

    MasterDoc.MailMerge.DataSource.ActiveRecord = wdLastRecord
    LastRecordNum = MasterDoc.MailMerge.DataSource.ActiveRecord
    MasterDoc.MailMerge.DataSource.ActiveRecord = wdFirstRecord

    Do While LastRecordNum > 0

        MasterDoc.MailMerge.Destination = wdSendToNewDocument
        MasterDoc.MailMerge.DataSource.FirstRecord = MasterDoc.MailMerge.DataSource.ActiveRecord
        MasterDoc.MailMerge.DataSource.LastRecord = MasterDoc.MailMerge.DataSource.ActiveRecord
        MasterDoc.MailMerge.Execute False

        Set SingleMergeDoc = ActiveDocument

        SingleMergeDoc.SaveAs2 _
            FileName:=MasterDoc.MailMerge.DataSource.DataFields("DocFolder").Value & Application.PathSeparator & _
                MasterDoc.MailMerge.DataSource.DataFields("FileName").Value & ".docx", _
            FileFormat:=wdFormatXMLDocument

        SingleMergeDoc.ExportAsFixedFormat _
            OutputFileName:=MasterDoc.MailMerge.DataSource.DataFields("PdfFolder").Value & Application.PathSeparator & _
                MasterDoc.MailMerge.DataSource.DataFields("FileName").Value & ".pdf", _
            ExportFormat:=wdExportFormatPDF

        SingleMergeDoc.Close False

        If MasterDoc.MailMerge.DataSource.ActiveRecord >= LastRecordNum Then
            LastRecordNum = 0
        Else
            MasterDoc.MailMerge.DataSource.ActiveRecord = wdNextRecord
        End If

    Loop

End Sub

---

### **Step 3: Run the Macro**
- Running the macro will:
  1. Perform the **mail merge**.
  2. Split the merged document into **separate files** saved in the `DocFolder` path from your Excel sheet.
  3. Name the files according to the `FileName` column.

---

## 📌 Notes & Tips
- If using iManage:
  - The macro will copy the **DocID** from the main document to all split files.
  - To avoid duplicates, either:
    - Delete the DocID field before running the macro, **or**
    - Update the DocID after generating the files.
- When re-running the mail merge:
  - **Relink** your Excel data source when prompted.
  - Do **not** click "Yes" to the default prompt without reviewing.

---

## 📂 Example Excel Setup

| Entity        | Authorized Signatory | Address         | DocFolder                      | FileName                   |
|---------------|----------------------|-----------------|--------------------------------|----------------------------|
| ABC Corp      | John Doe             | 123 Main St     | `C:\Contracts\ABC`             | `ThisContract (ABC Corp)`  |
| XYZ Ltd       | Jane Smith           | 456 Oak Ave     | `C:\Contracts\XYZ`             | `ThisContract (XYZ Ltd)`   |

---

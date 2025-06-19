# 📧 Excel VBA Email Automation

This project automates email sending directly from Excel using VBA and Microsoft Outlook. A clickable shape in the workbook triggers the macro, which sends personalized emails to recipients listed in the dataset.

---

## 🚀 Features

- One-click email sending via shape-triggered macro
- Personalized emails for each recipient
- Seamless Microsoft Outlook integration
- Simple and clear dataset format

---

## ✅ Requirements

- Microsoft Excel (macro-enabled)
- Microsoft Outlook installed and configured
- Macro settings enabled in Excel

---

## 🧠 How It Works

1. The Excel sheet contains a table with the following columns:
   - Name
   - Email
   - Subject
   - Message

2. A shape labeled **"Send Emails"** is placed on the sheet. When clicked, it runs the macro.

3. The macro reads each row and sends an email via Outlook using the provided details.

---

## 🔄 Testing Before Sending Emails

If you want to **test the email content before actually sending**, change the `.Send` line in the code to `.Display`. This will open the email drafts in Outlook for review instead of sending them directly.

---

## 💻 VBA Code

```vba
Sub SendEmails()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust if your sheet has a different name
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Initialize Outlook
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Loop through each row
    For i = 2 To lastRow
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = ws.Cells(i, 2).Value
            .Subject = ws.Cells(i, 3).Value
            .Body = ws.Cells(i, 4).Value
            .Send ' Change to .Display to preview emails before sending
        End With
    Next i

    MsgBox "Emails Sent!"
End Sub
```

## 👩‍💻 Author

**Chamudi**  
Excel | VBA Automation | Data Solutions

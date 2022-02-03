Sub TableToEmail()

'' CREATE CHART OBJECTS
Dim chrt As ChartObject, chrtSht As Worksheet, ws As Worksheet, Path As String, copyrng As Range
'' CREATE EMAIL OBJECTS
Dim olApp As Outlook.Application
Dim olEmail As Outlook.MailItem


'' SETUP VARIABLES
Set ws = ActiveSheet
Set chrtSht = ActiveWorkbook.Sheets.Add

'' Skip renaming if there is sheet with that name already
On Error Resume Next
chrtSht.Name = "Chart Sheet"
On Error GoTo 0
Set copyrng = ws.UsedRange
Path = ActiveWorkbook.Path + Application.PathSeparator

'' CREATE IMAGE OF USED RANGE IN ACTIVE SHEET
copyrng.CopyPicture xlScreen, xlPicture

With chrtSht.ChartObjects.Add(copyrng.Left, copyrng.Top, copyrng.Width, copyrng.Height)
    .Activate
    .Chart.Paste
    .Chart.Export Filename:=Path + "table1.png", filtername:="png"
End With

'' SETUP EMAIL OBJECTS

Set olApp = Outlook.Application
Set olEmail = olApp.CreateItem(0)

'' CREATE EMAIL FEATURES
With olEmail
    .Display
    .Attachments.Add Source:=Path + "table1.PNG"
    
    .HTMLbody = "<p>Dear Executive Department," & _
    "<p> Please see sales per rep below:</p>" & _
    "<img src='" & Path & "table1.PNG'>" & _
    "<p> Please send bonus as soon as possible. Kindly remit payment at your earliest convenience.<p>Thank you,</p>"

    .Subject = "test snip tool"
    .To = "johnsmith@funny.com"

End With

End Sub
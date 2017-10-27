Sub senmail()
Dim OutApp As Object
    Dim OutMail As Object
'    Dim MailBody As Range
    Dim MailSubject As String
    Dim MailTo
    Dim MailCC
    Dim t As Long, c As Long
    t = Range("M100").End(xlUp).Row
    c = Range("N100").End(xlUp).Row
    Application.ScreenUpdating = False
   For i = 3 To t
   MailTo = MailTo & ";" & Cells(i, 13).Value 'TO list的范围
   Next
   For j = 3 To c
   MailCC = MailCC & ";" & Cells(j, 14).Value 'CC list的范围
   Next
'   Set MailBody = Range("B1:T" & Range("B53434").End(xlUp).Row)
     MailSubject = Range("L3")
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(olMailItem)
        On Error Resume Next
        With OutMail
            .To = MailTo
            .cc = MailCC
            .BCC = ""
            .Subject = MailSubject
'            .BodyFormat = Outlook.OlBodyFormat.olFormatHTML
'            .HTMLBody = RangetoHTML(MailBody)
            .Display
            
        End With
        On Error GoTo 0
        Set OutMail = Nothing
        Set OutApp = Nothing
        MsgBox "请完善邮件内容"
       Application.ScreenUpdating = True
End Sub

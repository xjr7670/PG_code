Sub Get_SAP_Object()
    ' 这个是总过程
    ' 用于启动A6P。并获取会话对象
    ' 然后再把这个会话对象传递其所有需要执行查询的过程
    
    Dim conn As Object
    Dim conn_count As Integer
    
    If Not IsObject(sap_pplication) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set sap_application = SapGuiAuto.GetScriptingEngine
    End If
    
    ' 统计当前连接对象数量。如果为0则创建连接。即启动A6P
    ' 如果不为0表示已经打开了A6P。则使用第1个连接对象
    conn_count = sap_application.Connections.Count
    If conn_count = 0 Then
        Set conn = sap_application.OpenConnection("A6P SC Prod(EN) - SSO")
    Else
        Set conn = sap_application.Connections(0)
    End If
    
    If Not IsObject(session) Then
       Set session = conn.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.connectObject session, "on"
       WScript.connectObject Application, "on"
    End If
    
    
    Dim sht As Worksheet
    Dim var_sht As Worksheet
    Dim last_row As Integer
    
    Set sht = Worksheets(1)
    Set var_sht = ActiveWorkbook.Worksheets(2)
    
    ' 先去除筛选，并清空原有信息。
    If sht.Range("A2") <> "" Then
        last_row = sht.Range("A2").End(xlDown).Row
        With sht.Range("A2:K" & last_row)
            .AutoFilter
            .ClearContents
        End With
    End If
    
    ' 执行查询
    'On Error GoTo exi
    Call run_zc527(session, var_sht)
    Call read_download_file
    Call filt_and_copy
    
    ' 发邮件
    'Call SendActiveWorkbook(var_sht)
'exi:
    'Exit Sub
End Sub

Sub run_zc527(ByVal sess As Object, ByVal sht As Worksheet)
    Dim session As Object
    Dim var_sht As Worksheet
    Dim executor As String
    Dim sap_variant As String
    Dim v_creator As String
    Dim add_day As Integer
    
    Set session = sess
    Set var_sht = sht
    executor = sht.Cells(2, 2).Text
    sap_variant = sht.Cells(4, 2).Text
    v_creator = sht.Cells(5, 2).Text
    add_day = sht.Cells(6, 2).Text
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzc527"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = sap_variant
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = v_creator
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").Text = Date + add_day
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlW_REF_ALV_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlW_REF_ALV_CONTAINER/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\" & executor & "\Documents\SAP\SAP GUI\"     ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zc527.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub

Sub read_download_file()
    Dim wbk As Workbook
    Dim sht As Worksheet
    Dim rng As Range
    Dim last_row As Integer
    
    Set wbk = GetObject("C:\Users\xian.jr\Documents\SAP\SAP GUI\zc527.xls")
    Set sht = wbk.Worksheets(1)
    last_row = sht.Range("B6").End(xlDown).Row
    Set rng = sht.Range("B6:L" & last_row)
    
    rng.Copy Workbooks(1).ActiveSheet.Range("A2")
    
    Set wbk = Nothing
    Set sht = Nothing
    ' 在模板表中做筛选并把符合条件的内容加边框，并复制
End Sub

Sub filt_and_copy()
    Dim sht As Worksheet
    Dim rng As Range
    Dim last_row As Integer
    
    Set sht = Worksheets(1)
    last_row = sht.Range("A2").End(xlDown).Row
    Set rng = sht.Range("A1:L" & last_row)
    
    rng.AutoFilter Field:=12, Criteria1:=0
    With sht.Range("A1:G" & sht.Range("A1").End(xlDown).Row)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Select
        .Copy
    End With
End Sub

Sub SendActiveWorkbook(ByVal sht As Worksheet)
    Dim var_sht As Worksheet
    Dim email_rng As Range
    Dim email_last_row As Integer
    Dim email_arr() As String
    
    Set var_sht = sht
    
    If var_sht.Cells(9, 2) <> "" Then
        email_last_row = var_sht.Range("B8").End(xlDown).Row
    Else
        email_last_row = 8
    End If
    
    Set email_rng = var_sht.Range("B8:B" & email_last_row)
    ReDim email_arr(email_rng.Rows.Count) As String

    For r = 8 To email_last_row
        email_arr(r - 7) = var_sht.Cells(r, 2)
    Next
    
    ActiveWorkbook.SendMail Recipients:=email_arr, Subject:="CM Gap Auto Report - " & Format(Date, "dd/mm/yy")

End Sub

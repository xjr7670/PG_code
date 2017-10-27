Sub summary_output()
    Dim wbk As Workbook
    Dim actual_output_sht As Worksheet
    Dim cm_output_sht As Worksheet
    Dim degree_sht As Worksheet
    Dim gap_sht As Worksheet
    Dim degree_rng As Range
    Dim temp_rng As Range
    Dim cm_output_sht_last_row As Integer
    Dim degree_sht_last_row As Integer
    Dim xqfw_td_plan As Single
    Dim xqfw_ld_actual As Single
    Dim line_fp_td_plan As Long
    Dim line_fp_ld_actual As Long
    Dim line_wip_td_plan As Long
    Dim line_wip_ld_actual As Long
    Dim plancodever As String
    Dim last_day As Date
    Dim td As Date
    
    Set wbk = Application.ThisWorkbook
    Set actual_output_sht = wbk.Worksheets("Actual Output Status")
    Set cm_output_sht = wbk.Worksheets("CM Output")
    Set degree_sht = wbk.Worksheets("Degree")
    Set gap_sht = wbk.Worksheets("CM output gap reason")
    cm_output_sht_last_row = cm_output_sht.Range("E65535").End(xlUp).Row
    degree_sht_last_row = degree_sht.Range("C65535").End(xlUp).Row
    Set degree_rng = degree_sht.Range("A2:A" & degree_sht_last_row)
    
    xqfw_td_plan = 0
    xqfw_ld_actual = 0
    line_fp_td_plan = 0
    line_fp_ld_actual = 0
    line_wip_td_plan = 0
    line_wip_ld_actual = 0
    td = actual_output_sht.Cells(24, 1)
    last_day = actual_output_sht.Cells(22, 1)
    
    
    For r1 = 2 To cm_output_sht_last_row
        
        If cm_output_sht.Cells(r1, 2) = "XQFW" Then
            ' 计算XQFW今天的plan和昨天actual
            
            plancodever = cm_output_sht.Cells(r1, 1) & cm_output_sht.Cells(r1, 5) & cm_output_sht.Cells(r1, 3)
            Set temp_rng = degree_rng.Find(What:=plancodever, LookIn:=xlValues)
            If temp_rng Is Nothing Then
                MsgBox plancodever & "  No degree"
                Exit Sub
            End If
            temp_r = temp_rng.Row
            If cm_output_sht.Cells(r1, 8) = last_day Then
                ' 计算昨天的actual
                xqfw_ld_actual = xqfw_ld_actual + (cm_output_sht.Cells(r1, 13) * degree_sht.Cells(temp_r, 6))
                
                ' 把gap写到output sheet
                Call get_gap_record(cm_output_sht, gap_sht, r1)
            ElseIf cm_output_sht.Cells(r1, 8) = td Then
                ' 计算今天的plan
                xqfw_td_plan = xqfw_td_plan + (cm_output_sht.Cells(r1, 12) * degree_sht.Cells(temp_r, 6))
            End If
        End If
        
        If cm_output_sht.Cells(r1, 2) <> "XQFW" Then
            ' 计算其它3条line今天的plan和昨天actual
            
            If cm_output_sht.Cells(r1, 8) = last_day Then
           
                ' 计算昨天的半品和成品actual
                If Left(cm_output_sht.Cells(r1, 5), 1) = 8 Then
                    line_fp_ld_actual = line_fp_ld_actual + cm_output_sht.Cells(r1, 13)
                ElseIf Left(cm_output_sht.Cells(r1, 5), 1) = 9 Then
                    line_wip_ld_actual = line_wip_ld_actual + cm_output_sht.Cells(r1, 13)
                End If
            
            ElseIf cm_output_sht.Cells(r1, 8) = td Then
            
                ' 计算今天的半品和成品plan
                If Left(cm_output_sht.Cells(r1, 5), 1) = 8 Then
                    line_fp_td_plan = line_fp_td_plan + cm_output_sht.Cells(r1, 12)
                ElseIf Left(cm_output_sht.Cells(r1, 5), 1) = 9 Then
                    line_wip_td_plan = line_wip_td_plan + cm_output_sht.Cells(r1, 12)
                End If
            End If
        End If
    Next
    
    For c = 2 To 40
        If actual_output_sht.Cells(1, c) = last_day Then
            ' 把昨天的actual写到表里
            actual_output_sht.Cells(5, c) = xqfw_ld_actual
            actual_output_sht.Cells(9, c) = line_fp_ld_actual
            actual_output_sht.Cells(11, c) = line_wip_ld_actual
        End If
        If actual_output_sht.Cells(1, c) = td Then
            ' 把今天的plan写到表里
            actual_output_sht.Cells(4, c) = xqfw_td_plan
            actual_output_sht.Cells(8, c) = line_fp_td_plan
            actual_output_sht.Cells(10, c) = line_wip_td_plan
        End If
    Next

End Sub

Sub run_zprs()
    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim conn As Object
    Dim session As Object
    Dim conn_count As Integer
    
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine

    conn_count = SapApp.Connections.Count
    If conn_count = 0 Then
        Set conn = SapApp.OpenConnection("A6P SC Prod(EN) - SSO")
    Else
        Set conn = SapApp.Children(0)
    End If
    
    Set session = conn.Children(0)

    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject Application, "on"
    End If
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzprs"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = ThisWorkbook.ActiveSheet.Cells(30, 1)
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ThisWorkbook.ActiveSheet.Cells(33, 1)
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    'session.findById("wnd[0]/usr/ctxtP_PLANT").Text = "1864"
    session.findById("wnd[0]/usr/ctxtP_DATEL").Text = ActiveSheet.Cells(22, 1)
    session.findById("wnd[0]/usr/ctxtP_DATEH").Text = ActiveSheet.Cells(24, 1)
    session.findById("wnd[0]/usr/ctxtP_DATEH").SetFocus
    session.findById("wnd[0]/usr/ctxtP_DATEH").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Dim temp_wbk As Workbook
    Set temp_wbk = ActiveWorkbook
    Dim sht As Worksheet
    Set sht = temp_wbk.Worksheets(1)
    Dim last_row As Integer
    last_row = sht.Range("E65536").End(xlUp).Row
    sht.Range("A2:R" & last_row).Copy
    ' Set sht = Nothing
    ' temp_wbk.Close SaveChanges:=False
End Sub

Sub get_gap_record(cm_output_sht, gap_sht, r)
    ' 把昨天Gap大于40的生产记录放到gap reason表里面
    gap_sht_start_row = gap_sht.Range("D2").End(xlDown).Row + 1
    If cm_output_sht.Cells(r, 12) - cm_output_sht.Cells(r, 13) > 40 Then
        gap_sht.Cells(gap_sht_start_row, 3) = cm_output_sht.Cells(r, 8)
        gap_sht.Cells(gap_sht_start_row, 4) = cm_output_sht.Cells(r, 5)
        gap_sht.Cells(gap_sht_start_row, 5) = cm_output_sht.Cells(r, 6)
        gap_sht.Cells(gap_sht_start_row, 6) = cm_output_sht.Cells(r, 12)
        gap_sht.Cells(gap_sht_start_row, 7) = cm_output_sht.Cells(r, 13)
        gap_sht.Cells(gap_sht_start_row, 8) = "=G" & gap_sht_start_row & "-" & "F" & gap_sht_start_row
    End If
End Sub


Sub copy_output()
    Dim last_col As Integer
    last_col = Range("AK5").End(xlToLeft).Column
    Range(Cells(1, 1), Cells(15, last_col)).Copy
End Sub

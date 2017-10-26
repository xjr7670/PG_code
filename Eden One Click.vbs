Sub Main_Sub()
    ' 这个是总过程
    ' 用于启动A6P。并获取会话对象
    ' 然后再把这个会话对象传递其所有需要执行查询的过程
    
    'Dim SapGuiAuto As Object
    'Dim sap_application As Object
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
        Set conn = sap_application.openConnection("A6P SC Prod(EN) - SSO")
    Else
        Set conn = sap_application.Connections(0)
    End If
    
    If Not IsObject(session) Then
       Set session = conn.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If

    Dim variant_sht As Worksheet
    Dim mb51_variant As String
    Dim mb51_variant_creator As String
    Dim mb51_start_date As String
    Dim mb51_end_date As String
    Dim bma_variant As String
    Dim bma_variant_creator As String
    Dim mb52_variant As String
    Dim mb52_variant_creator As String
    Dim zder_variant As String
    Dim zder_variant_creator As String
    Dim zder_start_date As String
    Dim zder_end_date As String
    Dim zccr_variant As String
    Dim zccr_variant_creator As String
    Dim zccr_change_start_date As String
    Dim zccr_change_end_date As String
    Dim zccr_order_start_date As String
    Dim zccr_order_end_date As String
    Dim ao_variant As String
    Dim ao_variant_creator As String
    Dim zlir_variant As String
    Dim zlir_variant_creator As String
    Dim username As String
    Dim folder_path As String

    
    Set variant_sht = ActiveWorkbook.Worksheets("Variant")
    mb51_variant = variant_sht.Cells(2, 2).Text
    mb51_variant_creator = variant_sht.Cells(2, 3).Text
    mb51_start_date = variant_sht.Cells(2, 4).Text
    mb51_end_date = variant_sht.Cells(2, 5).Text
    bma_variant = variant_sht.Cells(3, 2).Text
    bma_variant_creator = variant_sht.Cells(3, 3).Text
    mb52_variant = variant_sht.Cells(4, 2).Text
    mb52_variant_creator = variant_sht.Cells(4, 3).Text
    zder_variant = variant_sht.Cells(5, 2).Text
    zder_variant_creator = variant_sht.Cells(5, 3).Text
    zder_start_date = variant_sht.Cells(5, 4).Text
    zder_end_date = variant_sht.Cells(5, 5).Text
    zccr_variant = variant_sht.Cells(6, 2).Text
    zccr_variant_creator = variant_sht.Cells(6, 3).Text
    zccr_change_start_date = variant_sht.Cells(6, 4).Text
    zccr_change_end_date = variant_sht.Cells(6, 5).Text
    zccr_order_start_date = variant_sht.Cells(6, 6).Text
    zccr_order_end_date = variant_sht.Cells(6, 7).Text
    ao_variant = variant_sht.Cells(7, 2)
    ao_variant_creator = variant_sht.Cells(7, 3)
    zlir_variant = variant_sht.Cells(8, 2)
    zlir_variant_creator = variant_sht.Cells(8, 3)
    username = variant_sht.Cells(9, 2).Text
    folder_path = "C:\Users\" & username & "\Documents\SAP\SAP GUI\"
    
    ' 判断文件夹是否存在
    If Dir(folder_path, vbDirectory) = vbNullString Then
        MsgBox "Please change Variant!B9 to you email name! Without ""pg.com""", vbOKOnly, "Path Error"
        ActiveWorkbook.Worksheets("Variant").Activate
        ActiveWorkbook.Worksheets("Variant").Range("B9").Select
        Exit Sub
    End If
    
    'On Error Resume Next
    'Call run_zder(session, zder_variant, zder_variant_creator, zder_start_date, zder_end_date, folder_path, username)
    Call run_bma(session, bma_variant, bma_variant_creator, folder_path, username)
    'On Error Resume Next
    'Call run_oos(session, zccr_variant, zccr_variant_creator, zccr_change_start_date, zccr_change_end_date, zccr_order_start_date, zccr_order_end_date, folder_path, username)
    Call run_mb52(session, mb52_variant, mb52_variant_creator, folder_path, username)
    Call run_zlir(session, zlir_variant, zlir_variant_creator, folder_path, username)
    Call run_mb51(session, mb51_variant, mb51_variant_creator, mb51_start_date, mb51_end_date, folder_path, username)
    Call run_ao(session, ao_variant, ao_variant_creator, folder_path, username)
    Call read_data(username)
    
    ' 刷新Sheet3中左边AO及右边Channel的数据透视表
    ActiveWorkbook.Worksheets("Sheet3").PivotTables("AO_PT").RefreshTable
    ActiveWorkbook.Worksheets("Sheet3").PivotTables("Channel_PT").RefreshTable
    
    ActiveWorkbook.Activate
    ActiveWorkbook.Save
    MsgBox "Update Finished and saved", vbOKOnly, "Done!"
End Sub
Sub run_zder(ByVal sess As Object, ByVal var As String, ByVal creator As String, ByVal date_start As String, ByVal date_end As String, ByVal folder_path As String, ByVal usr As String)
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzder"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var                  ' SAP Variant
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator          ' Variant Creator
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_AUDAT-LOW").Text = date_start    ' 订单开始日期
    session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").Text = date_end ' 订单结束日期
    session.findById("wnd[0]/usr/ctxtS_AUDAT-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_AUDAT-LOW").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'If session.findById("wnd[1]/tbar[0]/btn[0]") = True Then        ' 如果弹出了没有数据的对话框
        'Debug.Print "yes"
        'Exit Sub
    'End If
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path    ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zder.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub

Sub run_bma(ByVal sess As Object, ByVal var As String, ByVal creator As String, ByVal folder_path As String, ByVal usr As String)
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False    ' 这句是用来设置窗口大小的
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzcxxbma"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator          ' Variant Creator
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press                 ' 这个是执行按钮

    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path     ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "bma.xls"                                 ' 文件名称
    session.findById("wnd[1]/tbar[0]/btn[11]").press        ' Replace按钮
End Sub

Sub run_zlir(ByVal sess As Object, ByVal var As String, ByVal creator As String, ByVal folder_path As String, ByVal usr As String)
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False    ' 这句是用来设置窗口大小的
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzlir"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator
    session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zlir.xls"
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub


Sub run_oos(ByVal sess As Object, var As String, creator As String, change_start As String, change_end As String, order_start As String, order_end As String, ByVal folder_path As String, usr As String)
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False
    'session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZCCR"
    session.findById("wnd[0]").sendVKey 0
    ActiveSheet.Range("A21:A22").Copy
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    'session.findById("wnd[1]/usr/btn%_ENAME_%_APP_%-VALU_PUSH").press
    'session.findById("wnd[2]/tbar[0]/btn[16]").press
    'session.findById("wnd[2]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_DATEOC-LOW").Text = change_start
    session.findById("wnd[0]/usr/ctxtS_DATEOC-HIGH").Text = change_end
    session.findById("wnd[0]/usr/ctxtS_AUDAT-LOW").Text = order_start
    session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").Text = order_end
    session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").SetFocus
    
    ' 不需要改材料了，直接用variant里面的code，即固定的Eden
    'session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press     ' 这个是更改Material按钮
    'session.findById("wnd[1]/tbar[0]/btn[16]").press    ' 这个是清空Material的按钮
    'session.findById("wnd[1]/tbar[0]/btn[24]").press    ' 这个是从粘贴板上传数据的按钮
    'session.findById("wnd[1]/tbar[0]/btn[8]").press     ' 这个是Copy按钮，即确认按钮
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press     ' 这个是执行按钮
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path     ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "OOS.xls"                                 ' 文件名称
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub


Sub run_mb52(ByVal sess As Object, var As String, creator As String, ByVal folder_path As String, usr As String)
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr").horizontalScrollbar.Position = 28
    
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path     ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "mb52.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press    ' 这个是Replace按钮

End Sub

Sub run_mb51(ByVal sess As Object, var As String, creator As String, date_start As String, date_end As String, ByVal folder_path As String, usr As String)
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb51"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = date_start
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = date_end
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    'session.findById("wnd[0]/tbar[1]/btn[48]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path     ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "mb51.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub

Sub run_ao(ByVal sess As Object, var As String, creator As String, ByVal folder_path As String, usr As String)
    ' 2017/9/20前，把从ZDER得到的订单当AO用
    
    Dim session As Object
    Set session = sess
    
    session.findById("wnd[0]").resizeWorkingPane 116, 19, False
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzder"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = var                  ' SAP Variant
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = creator          ' Variant Creator
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_AUDAT-LOW").Text = "2017/09/20"      ' 订单开始日期
    session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").Text = "2017/09/30"     ' 订单结束日期
    session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'If session.findById("wnd[1]/tbar[0]/btn[0]") = True Then        ' 如果弹出了没有数据的对话框
        'Debug.Print "yes"
        'Exit Sub
    'End If
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = folder_path     ' 文件保存路径
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ao.xls"
    'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub

Sub read_data(ByVal usr As String)
    Dim mb51_sht As Worksheet
    Dim bma_sht As Worksheet
    Dim zlir_sht As Worksheet
    Dim mb52_sht As Worksheet
    Dim zder_sht As Worksheet
    Dim oos_sht As Worksheet
    Dim ao_sht As Worksheet
    Dim arr
    Dim temp_wbk As Workbook
    Dim temp_sht As Worksheet
    
    Set mb51_sht = ActiveWorkbook.Worksheets("MB51")
    Set bma_sht = ActiveWorkbook.Worksheets("BMA")
    Set zlir_sht = ActiveWorkbook.Worksheets("Intransit")
    Set mb52_sht = ActiveWorkbook.Worksheets("MB52")
    Set ao_sht = ActiveWorkbook.Worksheets("AO")
    Set zder_sht = ActiveWorkbook.Worksheets("Shipment")
    Set oos_sht = ActiveWorkbook.Worksheets("OOS")
    
    ' 先不run zder 和 oos，因为现在还没有数
    ' arr = Array("mb51", "bma", "mb52", "zder", "oos")
    arr = Array("mb51", "bma", "zlir", "mb52", "ao")
    
    For Each f In arr
        fname = "C:\Users\" & usr & "\Documents\SAP\SAP GUI\" & f & ".xls"
        Set temp_wbk = GetObject(fname)
        Set temp_sht = temp_wbk.Worksheets(1)
        Select Case f
            Case "mb51"
                ' 每次都先清空上一次数据
                If mb51_sht.Range("C6") <> "" Then
                    mb51_sht.Range("B6:U" & mb51_sht.Range("C6").End(xlDown).Row).ClearContents
                End If
                temp_sht.Range("B6:U" & temp_sht.Range("C6").End(xlDown).Row).Copy Destination:=mb51_sht.Range("B6")
            Case "bma"
                ' 每次都先清空上一次数据
                If bma_sht.Range("B21") <> "" Then
                    bma_sht.Range("B21:Z" & bma_sht.Range("B21").End(xlDown).Row).ClearContents
                End If
                temp_sht.Range("B21:Z" & temp_sht.Range("B21").End(xlDown).Row).Copy Destination:=bma_sht.Range("B21")
            Case "zlir"
                ' 每次都先清空上一次数据
                zlir_sht.Range("A1:O" & zlir_sht.Range("D65536").End(xlUp).Row).ClearContents
                temp_sht.Range("A1:O" & temp_sht.Range("D65536").End(xlUp).Row).Copy Destination:=zlir_sht.Range("A1")
            Case "mb52"
                ' 每次都先清空上一次数据
                If mb52_sht.Range("B4") <> "" Then
                    mb52_sht.Range("B4:K" & mb52_sht.Range("B4").End(xlDown).Row).ClearContents
                End If
                temp_sht.Range("B4:K" & temp_sht.Range("B4").End(xlDown).Row).Copy Destination:=mb52_sht.Range("B4")
            Case "zder"
                ' 把每次run出来的shipment数据放在末尾
                Dim zder_last_row As Integer
                If zder_sht.Range("C3") <> "" Then
                    zder_last_row = zder_sht.Range("C2").End(xlDown).Row + 1
                Else
                    zder_last_row = 3
                End If
                temp_sht.Range("C25:U" & temp_sht.Range("C25").End(xlDown).Row).Copy Destination:=zder_sht.Range("C" & zder_last_row)
            Case "oos"
                ' 把每次run出来的oos数据放在末尾
                Dim oos_last_row As Integer
                If oos_sht.Range("C6") <> "" Then
                    oos_last_row = oos_sht.Range("C5").End(xlDown).Row + 1
                Else
                    oos_last_row = 6
                End If
                temp_sht.Range("C8:V" & temp_sht.Range("C8").End(xlDown).Row).Copy Destination:=oos_sht.Range("C" & oos_last_row)
            Case "ao"
                ' 先清空上一次数据，再把本次数据复制过来
                
                If ao_sht.Range("C2") <> "" Then
                    ao_sht.Range("C2:U" & ao_sht.Range("C2").End(xlDown).Row).ClearContents
                End If
                temp_sht.Range("C25:U" & temp_sht.Range("C25").End(xlDown).Row).Copy Destination:=ao_sht.Range("C2")
                
                ' 再从CSR list中获取Channel
                Dim ao_last_row As Integer
                Dim csr_sht As Worksheet
                
                Set csr_sht = ActiveWorkbook.Worksheets("CSR list")
                ao_last_row = ao_sht.Range("C2").End(xlDown).Row
                For ao_r = 2 To ao_last_row
                    For csr_r = 2 To 2219
                        If ao_sht.Cells(ao_r, 10) = csr_sht.Cells(csr_r, 1) Then
                            ao_sht.Cells(ao_r, 22) = csr_sht.Cells(csr_r, 3)
                        End If
                    Next
                Next
        End Select
        
        Set temp_sht = Nothing
        temp_wbk.Close
    Next f
End Sub


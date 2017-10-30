Attribute VB_Name = "Module1"
Sub InventoryCheckMain()

    Rem �������ǿ�����������
    Rem ���XQ��HP������Ķ����븳ֵ
    Rem Ȼ�����MarkOther������ɿ�������
    
    Dim xq_sht As Worksheet
    Dim hp_sht As Worksheet
    Dim xq_max_row As Integer
    Dim hp_max_row As Integer
    
    Set xq_sht = Worksheets("XQ (1864 & 9216)")
    Set hp_sht = Worksheets("HP(0386 & 5578 & 0538)")
    xq_max_row = xq_sht.Range("B2").End(xlDown).Row
    hp_max_row = hp_sht.Range("B2").End(xlDown).Row


    Dim xq_cs_cell As Range        ' ����xq cs��Ԫ��
    Dim xq_cs_column As Integer    ' ����xq cs����
    Dim xq_ss_cell As Range        ' ����xq ss��Ԫ��
    Dim xq_ss_column As Integer    ' ����xq ss����
    Dim xq_mpq_cell As Range       ' ����xq mpq��Ԫ��
    Dim xq_mpq_column As Integer   ' ����xq mpq��Ԫ��
    
    Set xq_cs_cell = xq_sht.Range("1:1").Find("CS")
    xq_cs_column = xq_cs_cell.Column + 1
    Set xq_ss_cell = xq_sht.Range("1:1").Find("SS")
    xq_ss_column = xq_ss_cell.Column
    Set xq_mpq_cell = xq_sht.Range("1:1").Find("MPQ")
    xq_mpq_column = xq_mpq_cell.Column
    
    Dim hp_cs_cell As Range        ' ����hp cs��Ԫ��
    Dim hp_cs_column As Integer    ' ����hp cs����
    Dim hp_ss_cell As Range        ' ����hp ss��Ԫ��
    Dim hp_ss_column As Integer    ' ����hp ss����
    Dim hp_mpq_cell As Range       ' ����hp mpq��Ԫ��
    Dim hp_mpq_column As Integer   ' ����hp mpq��Ԫ��
    
    Set hp_cs_cell = hp_sht.Range("1:1").Find("CS")
    hp_cs_column = hp_cs_cell.Column + 1
    Set hp_ss_cell = hp_sht.Range("1:1").Find("SS")
    hp_ss_column = hp_ss_cell.Column
    Set hp_mpq_cell = hp_sht.Range("1:1").Find("MPQ")
    hp_mpq_column = hp_mpq_cell.Column

    '����֮ǰ���Ȱ���һ�α����ɫȥ��
    Call ClearColor(xq_sht, xq_max_row)
    Call ClearColor(hp_sht, hp_max_row)
    
    'Call XQ��
    Call MarkWIP(xq_sht, xq_max_row, xq_cs_column, xq_ss_column, xq_mpq_column)
    'Call MarkOtherStatues(xq_sht, xq_max_row)
    Call MarkDesc(xq_sht, xq_max_row, xq_cs_column)
    Call MarkNoMark(xq_sht, xq_max_row, xq_cs_column, xq_ss_column, xq_mpq_column)
    Call Align_To_Right(xq_sht)
    
    'Call HP��
    Call MarkWIP(hp_sht, hp_max_row, hp_cs_column, hp_ss_column, hp_mpq_column)
    Call MarkDesc(hp_sht, hp_max_row, hp_cs_column)
    Call MarkNoMark(hp_sht, hp_max_row, hp_cs_column, hp_ss_column, hp_mpq_column)
    Call Align_To_Right(hp_sht)
End Sub

Private Function MarkDict(sht As Worksheet, max_row As Integer) As Object
    Rem �����������ռ���Ҫ���İ�Ʒ�ı��
    Rem ����Щ��ǵ�����Ϊ���ֵ�ļ������ֶ�Ӧ���ֵĴ�����Ϊ�ֵ��ֵ
    Rem ���շ�������ֵ�
    
    Set MarkDict = CreateObject("Scripting.Dictionary")
    
    Dim dic As Object
    Dim n As Variant
    Dim nc As Integer
    Set dic = CreateObject("Scripting.Dictionary")
    
    For r = 2 To max_row
        n = sht.Cells(r, 1)
        If n <> "" Then
            nc = Application.WorksheetFunction.CountIf(sht.Range("A2:A" & max_row), n)
            If Not dic.Exists(n) Then
                dic.Add n, nc
            End If
        End If
    Next
    
    Rem ����ֵ����ʹ��Set�����ֱ��ʹ�õȺ��ǻ�����
    Set MarkDict = dic
End Function

Sub MarkWIP(sht As Worksheet, max_row As Integer, cs_column As Integer, ss_column As Integer, mpq_column As Integer)
    Rem
    Rem ��������WIP�����ĵ�2���汾
    Rem ������ʵ���ж����>=2����ƷʱҲ������ȷ���
    Rem ���������Ʒͬʱ�������춼�п�棬���ҵڶ�������������Ŀ��ֵ
    Rem ��ѵڶ�����
    
    Dim rng As Range
    Dim mark1 As Variant        '��һ�����ҵ��ı�־���ڵ�Ԫ��
    Dim mark2 As Variant        '�ڶ������ҵ��ı�־���ڵ�Ԫ��
    Dim r1 As Integer           '��һ����־������
    Dim r2 As Integer           '�ڶ�����־������
    Dim dic As Object
    Dim arr()                   '���ձ�־
    Dim nc                      '���ձ�־����
    Dim target_inv1 As Single   '���յ�һ��code��Ŀ����
    Dim target_inv2 As Single   '���յڶ���code��Ŀ����
    
    Set rng = sht.Range("A:A")
    Set dic = CreateObject("Scripting.Dictionary")
    
    Rem ���ú����������ֵ�
    Set dic = MarkDict(sht, max_row)
    
    Rem ���ֵ�ļ���������arr
    arr = dic.keys
    For Each num In arr
        '���ж���������Ʒ����һ��
        nc = dic(num)
        
        If nc = 1 Then
            '�����1����Ʒ
            
            Set mark1 = rng.Find(num)
            r1 = mark1.Row
            target_inv1 = sht.Cells(r1, mpq_column) + sht.Cells(r1, ss_column)
            For h = cs_column + 1 To cs_column + 31
                If sht.Cells(r1, h) > target_inv1 And sht.Cells(r1, h + 1) > target_inv1 Then
                    sht.Cells(r1, h + 1).Interior.Color = 255
                    sht.Cells(r1, 2).Interior.Color = 255
                    Exit For
                End If
            Next
        Else
            '����ж����Ʒ
            ReDim row_arr(nc) As Integer
            Set mark1 = rng.Find(num)
            row_arr(1) = mark1.Row

            For n = 2 To nc
                Set mark1 = rng.FindNext(mark1)
                row_arr(n) = mark1.Row
            Next
            
            '�Ѿ������а�Ʒ���ڵ���������������arr��
            
            ' ������������ֱ����ڴ��ͬ����Ʒǰ������Ŀ��
            ReDim inv_arr1(nc) As Single
            ReDim inv_arr2(nc) As Single
            
            ' ����target�������ڴ��ÿ����Ʒ��Ӧ��targetֵ
            ReDim target_arr(nc) As Single
            For n = 1 To nc
                target_arr(n) = sht.Cells(row_arr(n), mpq_column) + sht.Cells(row_arr(n), ss_column) + sht.Cells(row_arr(n), cs_column - 1)
            Next
            
            For c = cs_column + 1 To cs_column + 31
                For n = 1 To nc
                    inv_arr1(n) = sht.Cells(row_arr(n), c)
                    inv_arr2(n) = sht.Cells(row_arr(n), c + 1)
                Next
                
                flag = True
                For n = 1 To nc
                    If inv_arr1(n) > target_arr(n) And inv_arr2(n) > target_arr(n) Then
                        '
                    Else
                        flag = False
                    End If
                Next
                
                If flag Then
                    For n = 1 To nc
                        sht.Cells(row_arr(n), c + 1).Interior.Color = 255
                        sht.Cells(row_arr(n), 2).Interior.Color = 255
                    Next
                    Exit For
                End If
            Next
        End If
    Next num

End Sub


Sub Get_Current_Inv(base_sht As Worksheet, base_last_row As Integer, target_sht As Worksheet, target_last_row As Integer)
    
    Rem �������������Current Inv
    Rem �����ˣ�ֱ���ù�ʽ
End Sub

Sub MarkDesc(sht As Worksheet, max_row As Integer, cs_column As Integer)
    
    Rem ���������ڰѿ�涼��0��code������������ɫ
    
    For r = 2 To max_row
        If Application.WorksheetFunction.Sum(sht.Range(sht.Cells(r, cs_column), sht.Cells(r, cs_column + 30))) = 0 Then
            sht.Cells(r, 3).Interior.Color = RGB(160, 160, 160)
        End If
    Next
End Sub

Sub MarkNoMark(sht As Worksheet, max_row As Integer, cs_column As Integer, ss_column As Integer, mpq_column As Integer)
    
    Rem ���������ڴ���û�б�־��code
    Rem �������������Ŀ�����target����ѵ�2����
    
    Dim target_inv As Single
    
    For r = 2 To max_row
        If sht.Cells(r, 1) = "" Then
            target_inv = sht.Cells(r, ss_column) + sht.Cells(r, mpq_column) + sht.Cells(r, cs_column - 1)
            For h = cs_column To cs_column + 30
                If sht.Cells(r, h) > target_inv And sht.Cells(r, h + 1) > target_inv Then
                    sht.Cells(r, h + 1).Interior.Color = 255
                    sht.Cells(r, 2).Interior.Color = 255
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Sub ClearColor(sht As Worksheet, max_row As Integer)

    Rem ��������������һ�ν���Mark֮ǰ���Ȱ���һ�α����ɫȥ��
    sht.Range(sht.Cells(2, 1), sht.Cells(max_row, 30)).ClearFormats
End Sub

Sub Align_To_Right(sht As Worksheet)
    Rem
    Rem ���������ڰ�HP��XQ����Sheet��A���Ҷ��루�Ǳ�Ҫ����)
    
    sht.Range("A:A").HorizontalAlignment = xlRight
End Sub


Rem ====================================================================================================================================
Rem                                                                                                                                   ||
Rem ********************************************������������code�����̵ķָ�******************************************************||
Rem                                                                                                                                   ||
Rem ====================================================================================================================================

Sub CheckNewCodeMain()

    Rem �������Ǽ����code��������
    Rem ��Base data���е���code��ӵ�XQ��HP���ĩβ
    Rem ͨ������CheckNewCode()����ʵ��
    Rem
    
    Dim bd_sht As Worksheet
    Dim xq_sht As Worksheet
    Dim hp_sht As Worksheet
    Dim bd_maxrow As Integer
    Dim k As Long
    Dim v As String
    Dim hp_dic As Object
    Dim xq_dic As Object
    Dim date_col As Integer
    Dim date_rng As Range
    
    Set bd_sht = Worksheets("Base data")
    Set xq_sht = Worksheets("XQ (1864 & 9216)")
    Set hp_sht = Worksheets("HP(0386 & 5578 & 0538)")
    Set hp_dic = CreateObject("Scripting.Dictionary")
    Set xq_dic = CreateObject("Scripting.Dictionary")
    
    '�����ڸ��Ƶ�XQ��HP����
    date_col = bd_sht.Range("H1").End(xlToRight).Column
    ' ���ҳ�CS������
    Dim xq_cs_column As Integer
    Dim hp_cs_column As Integer
    Dim xq_cs_cell As Range
    Dim hp_cs_cell As Range
    Set xq_cs_cell = xq_sht.Range("1:1").Find(What:="CS")
    Set hp_cs_cell = hp_sht.Range("1:1").Find(What:="CS")
    xq_cs_column = xq_cs_cell.Column + 2
    hp_cs_column = hp_cs_cell.Column + 2
    Set date_rng = bd_sht.Range(bd_sht.Cells(1, 8), bd_sht.Cells(1, date_col))
    date_rng.Copy xq_sht.Cells(1, xq_cs_column)
    date_rng.Copy hp_sht.Cells(1, hp_cs_column)
    
    bd_maxrow = bd_sht.Range("B2").End(xlDown).Row
    Set bd_rng = bd_sht.Range("B2:B" & bd_maxrow)
    
    For i = 2 To bd_maxrow
        If bd_sht.Cells(i, 1) = 1864 Or bd_sht.Cells(i, 1) = 9216 Then
            k = bd_sht.Cells(i, 2)
            v = bd_sht.Cells(i, 3)
            If Not xq_dic.Exists(k) Then
                xq_dic.Add k, v
            End If
        ElseIf bd_sht.Cells(i, 1) = 386 Or bd_sht.Cells(i, 1) = 538 Or bd_sht.Cells(i, 1) = 5578 Then
            k = bd_sht.Cells(i, 2)
            v = bd_sht.Cells(i, 3)
            If Not hp_dic.Exists(k) Then
                hp_dic.Add k, v
            End If
        End If
    Next
    
    Call CheckNewCode(xq_sht, xq_dic)
    Call CheckNewCode(hp_sht, hp_dic)
    
    'ȡ����code������乫ʽ
    Call FillDownIt(xq_sht, xq_cs_column)
    Call FillDownIt(hp_sht, hp_cs_column)
End Sub

Sub CheckNewCode(sht As Worksheet, dic As Object)
    Rem
    Rem ������ʵ�������code�����
    Rem
    
    Dim rng As Range
    Dim maxrow As Integer
    Dim arr()
    
    maxrow = sht.Range("D2").End(xlDown).Row
    Set rng = sht.Range("B2:B" & maxrow)
    arr = dic.keys
    
    For Each code In arr
        If rng.Find(What:=code, LookIn:=xlValues) Is Nothing Then
            maxrow = maxrow + 1
            sht.Cells(maxrow, 2) = code
            sht.Cells(maxrow, 3) = dic.Item(code)
        End If
    Next code
    
End Sub

Sub FillDownIt(sht As Worksheet, cs_column)
    '���ҵ�Ŀǰ��ʽ�Ѿ����õ����һ��
    Dim last_row As Integer
    Dim last_column As Integer
    Dim new_last_row As Integer

    last_row = sht.Cells(1, cs_column).End(xlDown).Row
    new_last_row = sht.Range("C2").End(xlDown).Row
    last_column = sht.Range("H1").End(xlToRight).Column

    If last_row < new_last_row Then
        sht.Range(sht.Cells(last_row, cs_column), sht.Cells(new_last_row, last_column)).FillDown
    Else
        Exit Sub
    End If
End Sub

Attribute VB_Name = "Check_Input"
Sub Check_Input()
    Dim sht As Worksheet
    Dim mdt_rng As Range
    Dim opt_rng As Range
    Dim mmii_rng As Range
    Dim value_rng As Range
    Dim last_row As Integer
    
    Set sht = Worksheets(4)
    
    last_row = sht.Range("B2").CurrentRegion.Rows.Count
    Set value_rng = sht.Range("M3:AF" & last_row)
    value_rng.Interior.ColorIndex = xlNone
    'Debug.Print last_row
    Call Check_Mandatory_Area(sht, last_row)
    Call Check_Optional_Area(sht, last_row)
    Call Check_Master_Input(sht, last_row)
End Sub

Sub Check_Mandatory_Area(ByVal sht As Worksheet, ByVal last_row As Integer)
    Dim flag As Boolean
    flag = True
    
    For r = 3 To last_row
        For c = 2 To 5
            If sht.Cells(r, c) = "" Then
                sht.Cells(r, c).Interior.ColorIndex = 3
                flag = False
            End If
            If c = 5 And sht.Cells(r, c) <> "Change" And (sht.Cells(r, 6) = "" Or sht.Cells(r, 7) = "") Then
                sht.Range(Cells(r, 6), Cells(r, 7)).Interior.ColorIndex = 3
                flag = False
            End If
        Next
        ' Format Plant column
        If sht.Cells(r, 8) = "" Then
            sht.Cells(r, 8).Interior.ColorIndex = 3
        End If
    Next
    
    'Check_Mandatory_Area = flag
End Sub

Sub Check_Optional_Area(ByVal sht As Worksheet, ByVal last_row As Integer)
    Dim flag As Boolean
    flag = True
    
    For r = 3 To last_row
        If sht.Cells(r, 9) <> "" Then
            If Application.WorksheetFunction.CountA(sht.Range(Cells(r, 10), Cells(r, 12))) > 0 Then
                sht.Range(Cells(r, 10), Cells(r, 12)).Interior.ColorIndex = 3
            Else
                sht.Range(Cells(r, 10), Cells(r, 12)).Interior.ColorIndex = 16
            End If
        End If
        
        If Application.WorksheetFunction.CountA(sht.Range(Cells(r, 9), Cells(r, 12))) = 0 Then
            sht.Range(Cells(r, 9), Cells(r, 12)).Interior.ColorIndex = 3
        End If
    Next
    'Check_Optional_Area = flag
End Sub

Sub Check_Master_Input(ByVal sht As Worksheet, ByVal last_row As Integer)
    Dim flag As Boolean
    Dim m As Integer
    Dim not_blank_count As Integer
    Dim dic As Object
    
    Set dic = CreateObject("Scripting.Dictionary")
    flag = True
    
    
    For r = 3 To last_row
        dic.RemoveAll
        ' 这一块先检查当Request Type=Change时Material Master Information Input是否为空的情况
        not_blank_count = Application.WorksheetFunction.CountA(sht.Range(sht.Cells(r, 13), sht.Cells(r, 32)))
        
        If sht.Cells(r, 5) = "Change" Then
            If not_blank_count < 2 Then
                sht.Range(sht.Cells(r, 13), sht.Cells(r, 32)).Interior.ColorIndex = 3
                flag = False
            Else
                For c = 13 To 32
                    m = c Mod 2
                    If sht.Cells(r, c) <> "" Then
                        If m = 1 And sht.Cells(r, c + 1) = "" Then
                            ' 余数为1表示当前为Field列
                            sht.Cells(r, c + 1).Interior.ColorIndex = 3
                            flag = False
                        ElseIf m = 0 And sht.Cells(r, c - 1) = "" Then
                            ' 余数为0表示当前为Setting列
                            sht.Cells(r, c - 1).Interior.ColorIndex = 3
                            flag = False
                        Else
                            sht.Cells(r, c).Interior.ColorIndex = xlNone
                        End If
                    End If
                Next
                
                ' 当这个区域的非空列大于等于2时，还要检查Field1到Field10是否有重复项
                For c2 = 13 To 32
                    k = sht.Cells(r, c2)
                    v = c2
                    'Debug.Print k
                    If k <> "" Then
                        If Not dic.Exists(k) Then
                            dic.Add k, v
                        Else
                            sht.Cells(r, c2).Interior.ColorIndex = 3
                            sht.Cells(r, dic.Item(k)).Interior.ColorIndex = 3
                        End If
                    End If
                    c2 = c2 + 1
                Next
            End If
        End If
    Next
    
    'Check_Master_Input = flag
End Sub

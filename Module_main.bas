Dim UDCount As Long
Dim NewDate As Date
Dim Client As String
Dim Contents As String
Dim Classification As String
Dim Amount As Long
Sub Enter()
    
    On Error GoTo Exception
    
    NewDate = InputBox("日付を入力", "新規取引", "")
    Client = InputBox("支払先を入力", "新規取引", "")
    Contents = InputBox("内容を入力", "新規取引", "")
    Classification = InputBox("分類を入力", "新規取引", "")
    Amount = InputBox("金額を入力", "新規取引", "")
    
    For UDCount = 1 To 100000 Step 1
    
        If Cells(2 + UDCount, 2) = "" Then
            
            Cells(2 + UDCount, 2) = NewDate
            Cells(2 + UDCount, 4) = Client
            Cells(2 + UDCount, 5) = Contents
            Cells(2 + UDCount, 6) = Classification
            
            Cells(2 + UDCount, 7).Validation.Delete
            
            Cells(2 + UDCount, 7).Validation.Add _
                                  Type:=xlValidateList, _
                                  Formula1:= _
                                  "A,B,C"
                                  
            Cells(2 + UDCount, 8) = Amount
            
            Exit Sub
            
        Else
            
        End If
        
    Next UDCount
        
Exception:
        MsgBox "処理がキャンセルされました", vbOKOnly + vbInformation, "お知らせ"
        Exit Sub
        
End Sub
Sub ReRecord()
    Dim Counter As Long
    
    Counter = Application.WorksheetFunction.CountA(Range("G3:G100000"))
    
    If Cells(Counter + 2, 7) = "A" Then
        Worksheets("手段1").Select
        Exit Sub
        
    ElseIf Cells(Counter + 2, 7) = "B" Then
        Worksheets("手段2").Select
        'Call CashReRecord
        Exit Sub
        
    ElseIf Cells(Counter + 2, 7) = "C" Then
        Worksheets("クレジット1").Select
        'Call CashReRecord
        Exit Sub
            
    Else
        MsgBox "MeasnERROR", vbOKOnly + vbCritical, "ERROR"
        Exit Sub
        
    End If
            
End Sub

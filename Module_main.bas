Sub Enter()
    Dim NewDate As Date
    Dim Client As String
    Dim Contents As String
    Dim Classification As String
    Dim Amount As Long
    Dim UDCounter As Long
    Dim CheckUDCounter As Long
    
    UDCounter = Range("メインテーブル[#ALL]").Rows.Count - 2
    
    For CheckUDCounter = 0 To UDCounter Step 1
    
        If Cells(CheckUDCounter + 3, 2) = "" Then
            MsgBox "日付の空白を検知したため処理を続行できません。", vbOKOnly + vbCritical, "Error"
            Cells(CheckUDCounter + 3, 2).Select
            Exit Sub
        Else
        
        End If
        
    Next CheckUDCounter
    
    On Error GoTo Exception
    
        NewDate = InputBox("日付を入力", "新規取引", "")
        Client = InputBox("支払先を入力", "新規取引", "")
        Contents = InputBox("内容を入力", "新規取引", "")
        Classification = InputBox("分類を入力", "新規取引", "")
        Amount = InputBox("金額を入力", "新規取引", "")

    MsgBox UDCounter
            
        Cells(UDCounter + 4, 2) = NewDate
        Cells(UDCounter + 4, 4) = Client
        Cells(UDCounter + 4, 5) = Contents
        Cells(UDCounter + 4, 6) = Classification
            
        Cells(UDCounter + 4, 7).Validation.Delete
            
        Cells(UDCounter + 4, 7).Validation.Add _
                              Type:=xlValidateList, _
                              Formula1:= _
                                  "現金,ICカード,クレジットカード"
                                  
        Cells(UDCounter + 4, 8) = Amount
              
Exception:
        MsgBox "処理がキャンセルされました", vbOKOnly + vbInformation, "お知らせ"
        Exit Sub
        
End Sub
Sub ReRecord()
    Dim NewDate As Date
    Dim Client As String
    Dim Contents As String
    Dim Classification As String
    Dim Amount As Long
    Dim UDCounter As Long
    Dim CheckUDCounter As Long
    Dim InstUDCounter As Long
    
    Worksheets("出費明細").Select
    
    UDCounter = Range("メインテーブル[#ALL]").Rows.Count
    
    NewDate = Cells(1 + UDCounter, 2)
    Client = Cells(1 + UDCounter, 4)
    Classification = Cells(1 + UDCounter, 6)
    Amount = Cells(1 + UDCounter, 8)
    
    MsgBox UDCounter
    
    If Cells(UDCounter + 1, 7) = "現金" Then
        Worksheets("現金").Select
        
        InstUDCounter = Range("現金テーブル[#ALL]").Rows.Count
        
        Cells(3 + InstUDCounter, 2) = NewDate
        Cells(3 + InstUDCounter, 4) = "出金"
        Cells(3 + InstUDCounter, 5) = InputBox(ActiveSheet.Name & "へ記録する内容の入力", "転記")
        Cells(3 + InstUDCounter, 6) = Amount
        
        Exit Sub
        
    ElseIf Cells(UDCounter + 1, 7) = "ICカード" Then
        Worksheets("ICカード").Select
        For ExUDCount = 1 To 1000000
        
            If Cells(3 + ExUDCount, 2) = "" Then
                Cells(3 + ExUDCount, 2) = NewDate
                Cells(3 + ExUDCount, 4) = "出金"
                Cells(3 + ExUDCount, 6) = Amount
                
                Exit Sub
            Else
            
            End If
            
        Next ExUDCount
        
    ElseIf Cells(UDCounter + 1, 7) = "クレジットカード" Then
    
        Worksheets("クレジットカード").Select
        
        InstUDCounter = Range("クレジットテーブル[#ALL]").Rows.Count
        
        Cells(3 + InstUDCounter, 2) = NewDate
        Cells(3 + InstUDCounter, 4) = Client
        Cells(3 + InstUDCounter, 5) = InputBox(ActiveSheet.Name & "に記録する内容を入力", ActiveSheet.Name & "への記録", "")
        Cells(3 + InstUDCounter, 6) = Classification
        Cells(3 + InstUDCounter, 9) = Amount
            
    Else
        MsgBox "項目[決済手段]に不備が存在する可能性があります。", vbOKOnly + vbCritical, "ERROR"
        Exit Sub
        
    End If
            
End Sub

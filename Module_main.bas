Attribute VB_Name = "Module_main"
Sub Enter()

    Dim UDCount As Long
    Dim NewDate As Date
    Dim Client As String
    Dim Contents As String
    Dim Classification As String
    Dim Measn As String
    Dim Amount As Integer
    
    NewDate = InputBox("���t�����", "�V�K���", "")
    Client = InputBox("�x��������", "�V�K���", "")
    Contents = InputBox("���e�����", "�V�K���", "")
    Classification = InputBox("���ނ����", "�V�K���", "")
    Amount = InputBox("���z�����", "�V�K���", "")
    
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
                                  "��i1,��i2,��i3"
                                  
            Cells(2 + UDCount, 8) = Amount
            
            
            Exit Sub
            
        Else
            
        End If
        
    Next UDCount
    
End Sub

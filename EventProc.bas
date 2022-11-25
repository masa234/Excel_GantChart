
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err
    
    'パブリック変数に格納
    If SetToPublic = False Then
        MsgBox "出力に失敗しました。", vbInformation
        GoTo 正方形長方形1_Click_Exit
    End If
    
    If CreateExcelFile = False Then
        MsgBox "出力に失敗しました。", vbInformation
        GoTo 正方形長方形1_Click_Exit
    End If
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
End Sub



'「概要」：パブリック変数に格納
Public Function InputCheck() As Boolean
On Error GoTo InputCheck_Err
    
    InputCheck = False
    
    '「工数」がLong型でなければ、終了
    If IsLong(gloHours) = False Then
        GoTo InputCheck_Exit
    End If
    
    '「月」がInt型でなければ、終了
    If IsInt(gloMonth) = False Then
        GoTo InputCheck_Exit
    End If
    
    '「月」が1～12の範囲内でない場合、終了
    If CInt(gloMonth) < 1 Or _
        CInt(gloMonth) > 12 Then
        GoTo InputCheck_Exit
    End If

    InputCheck = True
    
InputCheck_Err:

InputCheck_Exit:
End Function


'「概要」：パブリック変数に格納
Public Function SetToPublic() As Boolean
On Error GoTo SetToPublic_Err
    
    SetToPublic = False
    
    '人数
    gloPersonCount = ThisWorkbook.Worksheets("勤務表出力").Cells(1, 2).Value

    '工数
    gloHours = ThisWorkbook.Worksheets("勤務表出力").Cells(2, 2).Value
    
    '年
    gloYear = "2022"
    
    '工数
    gloMonth = ThisWorkbook.Worksheets("勤務表出力").Cells(3, 2).Value
    
    '祝日配列に格納
    gloArrHoliday = GetRowDatasToArray("祝日", 2, 2)

    SetToPublic = True
    
SetToPublic_Err:

SetToPublic_Exit:
End Function


'「概要」：パブリック変数に格納
Public Function CreateExcelFile() As Boolean
On Error GoTo CreateExcelFile_Err
    
    CreateExcelFile = False

    'ブック作成
    Workbooks.Add
    
    'シート名を変更
    ActiveSheet.Name = CStr(gloYear) & "_" & CStr(gloMonth) & "月勤務表"
    
    'ヘッダ描画
    If CreateHeader = False Then
        GoTo CreateExcelFile_Exit
    End If
    
    '人数枠
    If CreateBorder = False Then
        GoTo CreateExcelFile_Exit
    End If
    
    '線を引く
    If DrawLine = False Then
        GoTo CreateExcelFile_Exit
    End If

    CreateExcelFile = True
    
CreateExcelFile_Err:

CreateExcelFile_Exit:
End Function


'「概要」：ヘッダ（1日分）
Public Function CreateHeader() As Boolean
On Error GoTo CreateHeader_Err
    
    CreateHeader = False
    
    Dim lngHeaderRow As Long
    Dim lngStartColumn As Long
    Dim lngCount As Long
    
    lngHeaderRow = 1
    lngStartColumn = 4

    '30回繰り返す
    For lngCount = 1 To 30
        'ヘッダ（1日分）描画
        If CreateHeaderPerDay(lngHeaderRow, lngStartColumn, gloHours, lngCount) = False Then
            GoTo CreateHeader_Exit
        End If
        '開始列を工数分進める
        lngStartColumn = lngStartColumn + gloHours
    Next lngCount
    
    CreateHeader = True
    
CreateHeader_Err:

CreateHeader_Exit:
End Function


'「概要」：ヘッダ（1日分）
Public Function CreateHeaderPerDay(ByVal lngHeaderRow As Long, _
                                ByVal lngStartColumn As Long, _
                                ByVal lngHours As Long, _
                                ByVal lngDate As Long) As Boolean
On Error GoTo CreateHeaderPerDay_Err
    
    CreateHeaderPerDay = False
    
    'セル「行：lngHeaderRow, 列：lngStartColumn」～セル「行：lngHeaderRow, 列：lngStartColumn + lngHours- 1」の範囲で囲む
    With Range(Cells(lngHeaderRow, lngStartColumn), Cells(lngHeaderRow, lngStartColumn + lngHours - 1))
        .BorderAround ColorIndex:=vbBlack, Weight:=xlThick
        .Interior.Color = GetHeaderColor(CStr(gloYear) & "/" & Format(CStr(gloMonth), "00") & "/" & Format(CStr(lngDate), "00"))
    End With
    
    'セル「行：lngHeaderRow, 列：lngStartColumn」に日付を格納
    Cells(lngHeaderRow, lngStartColumn).Value = CStr(lngDate)
    
    CreateHeaderPerDay = True
    
CreateHeaderPerDay_Err:

CreateHeaderPerDay_Exit:
End Function

'「概要」：ヘッダ塗りつぶし色取得
Public Function GetHeaderColor(ByVal strDate As String) As Long
On Error GoTo GetHeaderColor_Err
    
    '曜日によって色を取得
    
    '祝日の場合
    If IsHoliday(strDate) = True Then
        GetHeaderColor = RGB(250, 219, 218)
        '土曜日の場合
        If Weekday(strDate) = vbSaturday Then
            GetHeaderColor = RGB(157, 204, 224)
        End If
        GoTo GetHeaderColor_Exit
    End If
    
    '平日の場合
    GetHeaderColor = RGB(255, 166, 0)
    
GetHeaderColor_Err:

GetHeaderColor_Exit:
End Function


'「概要」：人数枠入力
Public Function CreateBorder() As Boolean
On Error GoTo CreateBorder_Err
    
    CreateBorder = False
    
    '開始行から繰り返す
    Dim lngStartRow As Long
    
    lngStartRow = 2
    lngStartColumn = 1
    
    '人数分繰り返す
    For lngCount = 1 To gloPersonCount
        '人数1人分描画
        If CreateBorderOfPerson(lngStartRow, lngStartColumn) = False Then
            GoTo CreateBorder_Exit
        End If
        '開始行＋２
        lngStartRow = lngStartRow + 2
    Next lngCount
    
    CreateBorder = True
    
CreateBorder_Err:

CreateBorder_Exit:
End Function


'「概要」：人数（1人分）
Public Function CreateBorderOfPerson(ByVal lngStartRow As Long, ByVal lngStartColumn As Long) As Boolean
On Error GoTo CreateBorderOfPerson_Err
    
    CreateBorderOfPerson = False
    
    'セル「行：lngStartRow, 列：lngStartColumn」～セル「行：lngStartRow +1, 列：lngStartColumn + 4」の範囲で囲む
    Range(Cells(lngStartRow, lngStartColumn), Cells(lngStartRow + 1, lngStartColumn + 2)).BorderAround ColorIndex:=vbBlack, Weight:=xlThick
    
    CreateBorderOfPerson = True
    
CreateBorderOfPerson_Err:

CreateBorderOfPerson_Exit:
End Function


'「概要」：工数入力
Public Function DrawLine() As Boolean
On Error GoTo DrawLine_Err
    
    DrawLine = False
    
    Dim lngStartColumn As Long
    Dim lngDrawRow As Long
    Dim lngCount As Long
    Dim lngDrawCount As Long
    
    '開始列
    lngDrawRow = 3
    lngStartColumn = 4
    
    For lngDrawCount = 1 To gloPersonCount
        '1～30まで繰り返す
        For lngCount = 1 To 30
            '平日の場合
            If IsHoliday(CStr(gloYear) & "/" & Format(CStr(gloMonth), "00") & "/" & Format(CStr(lngCount), "00")) = False Then
                '工数描画
                If DrawLineOneDay(lngDrawRow, lngStartColumn, gloHours) = False Then
                    GoTo DrawLine_Exit
                End If
            Else
            '祝日の場合
                'グレーアウト
                If GlayoutOneDay(lngDrawRow, lngStartColumn, gloHours) = False Then
                    GoTo DrawLine_Exit
                End If
            End If
            '開始列＝開始列＋工数
            lngStartColumn = lngStartColumn + gloHours
        Next lngCount
        lngStartColumn = 4
        lngDrawRow = lngDrawRow + 2
    Next lngDrawCount

    DrawLine = True
    
DrawLine_Err:

DrawLine_Exit:
End Function


'「概要」：工数（1日分）
Public Function DrawLineOneDay(ByVal lngDrawRow As Long, ByVal lngStartColumn As Long, ByVal lngHours As Long) As Boolean
On Error GoTo DrawLineOneDay_Err
    
    DrawLineOneDay = False
    
    '色を塗る
    'セル「行：lngDrawRow , 列: lngStartColumn」～「行：lngDrawRow , 列: lngStartColumn + lngHours -1」
    Range(Cells(lngDrawRow, lngStartColumn), Cells(lngDrawRow, lngStartColumn + lngHours - 1)).Interior.Color = RGB(157, 204, 224)

    DrawLineOneDay = True
    
DrawLineOneDay_Err:

DrawLineOneDay_Exit:
End Function


'「概要」：グレーアウト工数（1日分）
Public Function GlayoutOneDay(ByVal lngDrawRow As Long, ByVal lngStartColumn As Long, ByVal lngHours As Long) As Boolean
On Error GoTo GlayoutOneDay_Err
    
    GlayoutOneDay = False
    
    '色を塗る
    'セル「行：lngDrawRow , 列: lngStartColumn」～「行：lngDrawRow , 列: lngStartColumn + lngHours -1」
    Range(Cells(lngDrawRow - 1, lngStartColumn), Cells(lngDrawRow, lngStartColumn + lngHours - 1)).Interior.Color = RGB(128, 128, 128)

    GlayoutOneDay = True
    
GlayoutOneDay_Err:

GlayoutOneDay_Exit:
End Function

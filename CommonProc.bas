'工数
Public gloHours As Long
'年
Public gloYear As Long
'月
Public gloMonth As Long
'人数
Public gloPersonCount As Long
'祝日配列
Public gloArrHoliday() As Variant


'「概要」：休日かどうか？
Public Function IsHoliday(ByVal strDate As String) As Boolean
On Error GoTo IsHoliday_Err

    IsHoliday = False
    
    'Weekday関数の返り値が土、日になっている場合、trueを返す
    If Weekday(strDate) = vbSunday Or _
        Weekday(strDate) = vbSaturday Then
        IsHoliday = True
        GoTo IsHoliday_Exit
    End If
    
    
    '祝日の場合、trueを返す
    If InArray(gloArrHoliday, strDate) = True Then
        IsHoliday = True
        GoTo IsHoliday_Exit
    End If
    
IsHoliday_Err:

IsHoliday_Exit:
End Function


'「概要」：配列内に値が存在するか
Public Function InArray(ByVal arrSearch As Variant, ByVal strSearch As String) As Boolean
On Error GoTo InArray_Err

    InArray = False
    
    Dim lngCount As Long
    Dim lngArrIndex As Long
    
    lngArrIndex = 0
    
    '配列の終端まで繰り返す
    For lngCount = 0 To UBound(arrSearch)
        '配列の要素が検索値の場合、trueを返す
        If arrSearch(lngArrIndex) = strSearch Then
            InArray = True
            GoTo InArray_Exit
        End If
        lngArrIndex = lngArrIndex + 1
    Next lngCount
    
InArray_Err:

InArray_Exit:
End Function


'「概要」：シートの列データを配列に格納
Public Function GetRowDatasToArray(ByVal strSheetName As String, _
                                ByVal lngHeaderRow As Long, _
                                ByVal lngDataColumn As Long) As Variant
On Error GoTo GetRowDatasToArray_Err

    GetRowDatasToArray = False
    
    Dim arrData(16) As Variant
    Dim lngLastRow As Long
    Dim lngCount As Long
    Dim lngArrIndex As Long
    
    lngArrIndex = 0
    
    'シートの最終行を取得
    lngLastRow = ThisWorkbook.Worksheets(strSheetName).Cells(lngHeaderRow, lngDataColumn).End(xlDown).Row
    
    'ヘッダ行から最終行まで繰り返す
    For lngCount = lngHeaderRow To lngLastRow
        'シートのセル（行：lngCount, 列：lngDataColumn）の値を配列に格納
        arrData(lngArrIndex) = ThisWorkbook.Worksheets(strSheetName).Cells(lngCount, lngDataColumn).Value
        lngArrIndex = lngArrIndex + 1
    Next lngCount
    
    GetRowDatasToArray = arrData()
    
GetRowDatasToArray_Err:

GetRowDatasToArray_Exit:
End Function


'「概要」：数値かどうか？
Public Function IsInt(ByVal strCheck As String) As Boolean
On Error GoTo IsInt_Err

    IsInt = False
    
    Dim i As Integer
    
    '空の場合、falseを返す
    If strCheck = vbNullString Then
        GoTo IsInt_Exit
    End If
    
    'int型に変換できない場合、エラーが発生する
    i = CInt(strCheck)
    
    IsInt = True
    
IsInt_Err:

IsInt_Exit:
End Function


'「概要」：数値かどうか？
Public Function IsLong(ByVal strCheck As String) As Boolean
On Error GoTo IsLong_Err

    IsLong = False
    
    Dim lng As Long
    
    '空の場合、falseを返す
    If strCheck = vbNullString Then
        GoTo IsLong_Exit
    End If
    
    'long型に変換できない場合、エラーが発生する
    lng = CInt(strCheck)
    
    IsLong = True
    
IsLong_Err:

IsLong_Exit:
End Function

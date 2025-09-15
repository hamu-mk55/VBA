Attribute VB_Name = "CleanSheet"
Option Explicit


Sub CleanSheet()
    Dim mainSheet As Worksheet
    Dim dataSheet As Worksheet
    
    Dim keepCols As Variant
    
    ' パラメータ
    Set mainSheet = ThisWorkbook.Worksheets("main")
    Set dataSheet = ThisWorkbook.Worksheets("data")
    
    ' 非表示行を消す
    Call DeleteHiddenRows(dataSheet)
    
    ' 保持する列をシートから取得し、それ以外の列を削除
    keepCols = GetStrList(mainSheet)
    Call DeleteColsExceptArr(dataSheet, keepCols, 0)

End Sub


' 非表示の行を削除
Sub DeleteHiddenRows(ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    
    If ws.AutoFilterMode Then
        lastRow = ws.AutoFilter.Range(ws.AutoFilter.Range.Count).Row
    Else
        lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    End If
        
    ' ヘッダー以外を残して、行を下から探索。非表示行を消す
    For r = lastRow To 2 Step -1
        If ws.rows(r).Hidden Then
            ws.rows(r).Delete
        End If
    Next r

End Sub


' ws のシートで keepCols に含まれない列を全削除
' ws: ワークシート
' keepCols: 保持したい列記号の配列(A,B,C)
' offset: 配列で参照した列番号と、ワークシートの列番号のオフセット
' 例：配列中でB列に対応する列が、ワークシートだとA列の場合、offset=1
Private Sub DeleteColsExceptArr(ws As Worksheet, keepCols As Variant, offset As Long)
    Dim maxCol As Long
    
    Dim i As Long
    Dim col_no As Long
    Dim keep_dict As Object

    Set keep_dict = CreateObject("Scripting.Dictionary")
    keep_dict.CompareMode = 1 ' 大文字小文字を区別しない

    ' 列番号に変換したうえで辞書登録
    ' 辞書　key: 列番号、val: True/False
    For i = LBound(keepCols) To UBound(keepCols)
        col_no = ColumnLetterToNumber(CStr(keepCols(i)))
        keep_dict(col_no - offset) = True
    Next

    maxCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                           SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    ' 削除による列番号ずれの影響を受けないように、右→左に削除
    For i = maxCol To 1 Step -1
        If Not keep_dict.Exists(i) Then
            ws.Columns(i).Delete
        End If
    Next

End Sub

' 列記号(A,B,C)を列番号(1,2,3)に変換
' s: 列記号
Private Function ColumnLetterToNumber(s As String) As Long
    Dim i As Long
    Dim ch As Integer
    Dim n As Long
    
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1)) - 64
        n = n * 26 + ch
    Next
    ColumnLetterToNumber = n
End Function

' シートのセル内容を読み取り、配列に格納
' ws: ワークシート
Private Function GetStrList(ws As Worksheet) As Variant
    Dim startRow As Long
    Dim startCol As Long
    
    Dim i As Long
    Dim tmp As Collection
    Dim arr() As Variant

    Set tmp = New Collection

    ' 開始行/列の設定
    startRow = 1
    startCol = 2

    ' 指定行/列から開始し、セル値を配列に格納。
    ' 空白セルまで下側へ探索。
    For i = startRow To ws.rows.Count
        If LenB(Trim(ws.Cells(i, startCol).Value)) > 0 Then
            tmp.Add Trim(ws.Cells(i, startCol).Value)
        End If
    Next i

    ' Collection → 配列に変換
    ReDim arr(0 To tmp.Count - 1)
    For i = 1 To tmp.Count
        arr(i - 1) = tmp(i)
    Next i

    GetStrList = arr
End Function


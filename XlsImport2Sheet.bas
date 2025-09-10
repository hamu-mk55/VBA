Attribute VB_Name = "Module2"
Option Explicit

Sub Run_Import2Sheet()
    Dim xls_file As String
    Dim xls_sheetname As String
    Dim dst_sheetname As String
    Dim row_cnt As Double
    
    xls_file = "C:\Users\データ.xlsx"
    xls_sheetname = "データ"
    dst_sheetname = "data"

    row_cnt = 1
    Call CopyToDataSheet(xls_file, xls_sheetname, dst_sheetname, True, row_cnt, True)

    Debug.Print (row_cnt)

End Sub


' 指定ブック・指定シートを指定シートへ転記
' 必要に応じて書式コピー（keepFormats:=True）
Public Sub CopyToDataSheet( _
    xlsPath As String, _
    xlsSheetName As String, _
    Optional destSheetName As String = "data", _
    Optional clearDest As Boolean = True, _
    Optional ByRef startRow As Double = 1, _
    Optional keepFormats As Boolean = False)

    Dim wbSrc As Workbook
    Dim wbOpen As Boolean
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim rngSrc As Range
    Dim rngDst As Range
    Dim arr As Variant
    Dim appCalc As XlCalculation
    Dim scrUpdt As Boolean
    Dim evtEnabled As Boolean
    Dim f As Workbook
    Dim r As Long, c As Long

    On Error GoTo EH

    ' 高速化
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 対象ファイルを既に開いているか検索
    For Each f In Application.Workbooks
        If LCase$(f.FullName) = LCase$(xlsPath) Then
            Set wbSrc = f
            wbOpen = True
            Exit For
        End If
    Next f
    If wbSrc Is Nothing Then
        If Len(dir$(xlsPath)) = 0 Then
            Call Err.Raise(53, , "指定ファイルが見つかりません: " & xlsPath)
        End If
        Set wbSrc = Application.Workbooks.Open(FileName:=xlsPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    ' 読込シート取得
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(xlsSheetName)
    On Error GoTo EH
    If wsSrc Is Nothing Then
        Call Err.Raise(9, , "指定シートが見つかりません: " & xlsSheetName)
    End If

    ' 出力シート取得
    Set wsDst = SetSheet(ThisWorkbook, destSheetName)

    ' 出力シート初期化
    If clearDest Then
        wsDst.Cells.Clear
    End If

    ' 出力開始行
    Set rngDst = wsDst.Range("A" & startRow)

    ' ソース範囲
    Dim srcStart As Range
    Dim srcRow As Long
    Dim srcCol As Long
    Dim dstRowOff As Long
    Dim lastCol As Long
    Dim cCount As Long
    Dim MAX_GUARD As Long
    
    ' 起点セルを指定
    Set srcStart = wsSrc.Range("B3")
    
    srcRow = srcStart.Row
    srcCol = srcStart.Column
    dstRowOff = 0
    
    ' 念のため無限ループ防止（大きめに）
    MAX_GUARD = wsSrc.rows.Count
    
    Do While LenB(CStr(wsSrc.Cells(srcRow, srcCol).Value)) > 0 And MAX_GUARD > 0
        ' この行の最終使用列（左端から右へ）
        lastCol = wsSrc.Cells(srcRow, wsSrc.Columns.Count).End(xlToLeft).Column
        If lastCol < srcCol Then
            ' 起点列より左で止まっている＝実質データ無しとみなして終了
            Exit Do
        End If
    
        cCount = lastCol - srcCol + 1
    
        ' 書式も必要なら、先に1行コピー（keepFormats=Trueのとき）
        If keepFormats Then
            wsSrc.Cells(srcRow, srcCol).Resize(1, cCount).Copy _
                Destination:=rngDst.Offset(dstRowOff, 0)
        End If
    
        ' 値だけ上書き（1行分）
        rngDst.Offset(dstRowOff, 0).Resize(1, cCount).Value = _
            wsSrc.Cells(srcRow, srcCol).Resize(1, cCount).Value
    
        ' 次の行へ
        srcRow = srcRow + 1
        dstRowOff = dstRowOff + 1
        MAX_GUARD = MAX_GUARD - 1
    Loop


CLEANUP:
    If Not wbOpen Then wbSrc.Close SaveChanges:=False

    Application.ScreenUpdating = scrUpdt
    Application.EnableEvents = evtEnabled
    Application.Calculation = appCalc
    Exit Sub

EH:
    On Error Resume Next
    If Not wbSrc Is Nothing Then If Not wbOpen Then wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = scrUpdt
    Application.EnableEvents = evtEnabled
    Application.Calculation = appCalc
    MsgBox "転記に失敗しました。" & vbCrLf & Err.Description, vbExclamation, "CopyToDataSheet"
End Sub

' シートをセット。既存で無ければ末尾に追加
Private Function SetSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set SetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If SetSheet Is Nothing Then
        Set SetSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        SetSheet.name = sheetName
    End If
End Function



Attribute VB_Name = "Module2"
Option Explicit

Sub Run_CopyToData()
    Dim xls As String
    xls = "C:\work\report.xlsx"

    ' 1) 値のみで上書き転記（dataシートを事前に全消去）
    CopyToDataSheet xls, "ExportMe", "data", True, "A1", False

    ' 2) 書式もコピーしたい場合
    'CopyToDataSheet xls, "ExportMe", "data", True, "A1", True
End Sub


'=========================================
' 指定ブック・指定シート → ThisWorkbook("data") へ転記
' ・値のみを一括転記（高速）
' ・必要に応じて書式コピーも可（keepFormats:=True）
' ・ファイル名だけのsrcPathにも対応
'=========================================
Public Sub CopyToDataSheet( _
    ByVal xlsPath As String, _
    ByVal xlsSheetName As String, _
    Optional destSheetName As String = "data", _
    Optional clearDest As Boolean = True, _
    Optional startCell As String = "A1", _
    Optional keepFormats As Boolean = False)

    Dim wbSrc As Workbook, wasOpen As Boolean
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim rngSrc As Range, dst As Range
    Dim arr As Variant
    Dim appCalc As XlCalculation
    Dim scrUpdt As Boolean, evtEnabled As Boolean
    Dim f As Workbook
    Dim r As Long, c As Long

    On Error GoTo EH

    '=== 環境を一時的に軽量化 ===
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '=== srcPath を解決（相対/ファイル名のみ対応）===
    srcPath = ResolvePath(srcPath, ThisWorkbook.path)

    '=== ソースブックを取得（既に開いていればそれを使用）===
    For Each f In Application.Workbooks
        If LCase$(f.FullName) = LCase$(srcPath) Then
            Set wbSrc = f
            wasOpen = True
            Exit For
        End If
    Next f
    If wbSrc Is Nothing Then
        If Len(dir$(srcPath)) = 0 Then Err.Raise 53, , "指定ファイルが見つかりません: " & srcPath
        Set wbSrc = Application.Workbooks.Open(FileName:=srcPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    '=== ソース/デスティネーション シート取得（無ければ作成）===
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(srcSheetName)
    On Error GoTo EH
    If wsSrc Is Nothing Then Err.Raise 9, , "指定シートが見つかりません: " & srcSheetName

    Set wsDst = GetOrCreateSheet(ThisWorkbook, destSheetName)

    '=== 転記先開始セル ===
    Set dst = wsDst.Range(startCell)

    '=== 転記前クリア ===
    If clearDest Then
        wsDst.Cells.Clear ' 書式も消す。書式を残したい場合は ClearContents に変更
    End If

    '=== ソース範囲 ===
    Set rngSrc = wsSrc.UsedRange
    If Application.WorksheetFunction.CountA(rngSrc) = 0 Then
        ' 空なら何もせず終了
        GoTo CLEANUP
    End If

    '=== 書式も欲しい場合は先にコピー（後で値で上書き）===
    If keepFormats Then
        rngSrc.Copy Destination:=dst
    End If

    '=== 値のみ一括転記（高速：配列→一括代入）===
    arr = rngSrc.Value
    If IsArray(arr) Then
        r = UBound(arr, 1) - LBound(arr, 1) + 1
        c = UBound(arr, 2) - LBound(arr, 2) + 1
        dst.Resize(r, c).Value = arr
    Else
        ' 単一セルの場合
        dst.Value = arr
    End If


CLEANUP:
    If Not wasOpen Then wbSrc.Close SaveChanges:=False

    Application.ScreenUpdating = scrUpdt
    Application.EnableEvents = evtEnabled
    Application.Calculation = appCalc
    Exit Sub

EH:
    On Error Resume Next
    If Not wbSrc Is Nothing Then If Not wasOpen Then wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = scrUpdt
    Application.EnableEvents = evtEnabled
    Application.Calculation = appCalc
    MsgBox "転記に失敗しました。" & vbCrLf & Err.Description, vbExclamation, "CopyToDataSheet"
End Sub

' シート取得。無ければ末尾に追加して返す
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.name = name
    End If
End Function



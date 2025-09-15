Attribute VB_Name = "XlsImport"
Option Explicit

Sub ImportXlsFiles()
    Dim xls_file As Variant
    Dim xls_path As String
    Dim xls_sheetname As String
    Dim xls_startcell As String
    Dim xls_import_cols As Long
    Dim dst_sheetname As String
    Dim row_cnt As Long
    
    Dim mainSheet As Worksheet
    Dim dataSheet As Worksheet
    Dim xlsDirPath As String
    Dim FSO As Object
    Dim file_ext As String
    
    ' パラメータ
    Set mainSheet = ThisWorkbook.Worksheets("main")
    
    xls_sheetname = "101"
    xls_startcell = "B3"
    dst_sheetname = "data"
    
    xls_import_cols = 7
    
    
    ' ホルダ指定
    xlsDirPath = SelectDir()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not Len(xlsDirPath) > 0 Then
        Exit Sub
    End If
    
    ' 初期化
    Set dataSheet = GetWorkSheet(ThisWorkbook, dst_sheetname, True)
    dataSheet.Cells.Clear
    
    
    'フォルダ内のファイル処理
    row_cnt = 1
    For Each xls_file In FSO.getFolder(xlsDirPath).Files
        Debug.Print (xls_file & "," & row_cnt)
    
        file_ext = LCase$(FSO.GetExtensionName(xls_file.name))
        
        Select Case file_ext
            Case "xls", "xlsx", "xlsm"
                xls_path = xlsDirPath & "\" & xls_file.name
                Call ImportSingleXls(xls_path, xls_sheetname, xls_startcell, xls_import_cols, dst_sheetname, row_cnt)
            Case Else
                ' SKIP
        End Select
    Next
        
    Set FSO = Nothing
    
End Sub



' 指定ブック・指定シートを指定シートへ転記
' 必要に応じて書式コピー（keepFormats:=True）
' srcPath: 読み込むエクセルファイルのフルパス
' srcSheetName: 読み込むエクセルファイルのシート名
' srcStartCell: 読み込むエクセルファイルの読込開始セル(C3など)
' srcImportColNum: 読み込むエクセルファイルの読込列数(C~E列を読み込む場合は3)
' destSheetName: 出力先のシート名
' destStartRow: 出力先の出力開始行(>1のとき、2ファイル目と認識)
' keepFormats: 書式コピーするかどうか
Private Sub ImportSingleXls( _
    srcPath As String, _
    srcSheetName As String, _
    srcStartCell As String, _
    srcImportColNum As Long, _
    destSheetName As String, _
    ByRef destStartRow As Long, _
    Optional keepFormats As Boolean = False)

    Dim wbSrc As Workbook
    Dim wbOpen As Boolean
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    
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
        If LCase$(f.FullName) = LCase$(srcPath) Then
            Set wbSrc = f
            wbOpen = True
            Exit For
        End If
    Next f
    
    If wbSrc Is Nothing Then
        Set wbSrc = Application.Workbooks.Open(FileName:=srcPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    ' 読込シート取得。無ければエラーメッセージ出した上で処理終了
    Set wsSrc = GetWorkSheet(wbSrc, srcSheetName, False)
    If wsSrc Is Nothing Then
        MsgBox ("No sheetname for import: " & vbLf & wbSrc.name & vbLf & srcSheetName)
        GoTo CLEANUP
    End If

    ' 出力シート取得。無ければ追加
    Set wsDst = GetWorkSheet(ThisWorkbook, destSheetName, True)
    
    ' データ転記処理
    Dim rngSrc As Range
    Dim srcRow As Long
    Dim srcCol As Long
    Dim lastCol As Long
    
    Dim rngDst As Range
    
    Dim row_cnt As Long
    Dim MAX_LOOP As Long
    
    ' 出力開始行(列はA列固定)
    Set rngDst = wsDst.Range("A" & destStartRow)
    
    ' 読込開始セルを指定
    Set rngSrc = wsSrc.Range(srcStartCell)
    srcRow = rngSrc.Row
    srcCol = rngSrc.Column
    
    lastCol = srcCol + srcImportColNum - 1
    
    ' startRowが1で無い(=2回目以降)場合は、ヘッダSKIP
    If destStartRow > 1 Then
        srcRow = srcRow + 1
    End If
    
    ' 無限ループ防止
    MAX_LOOP = wsSrc.rows.Count
    
    row_cnt = 0
    Do While LenB(CStr(wsSrc.Cells(srcRow, srcCol).Value)) > 0
    
        ' 書式も必要なら、先に1行コピー
        If keepFormats Then
            wsSrc.Cells(srcRow, srcCol).Resize(1, srcImportColNum).Copy _
                Destination:=rngDst.offset(row_cnt, 0)
        End If
    
        ' 値を転記(１行)
        rngDst.offset(row_cnt, 0).Resize(1, srcImportColNum).Value = _
            wsSrc.Cells(srcRow, srcCol).Resize(1, srcImportColNum).Value
    
        ' 次の行へ
        srcRow = srcRow + 1
        row_cnt = row_cnt + 1
        
        ' MAX_LOOP以上でループ抜ける
        If row_cnt > MAX_LOOP Then
            Exit Do
        End If
        
        ' フリーズ対策（一定間隔でDoEvents）
        If srcRow Mod 1000 = 0 Then
            DoEvents
        End If
    Loop
    
    destStartRow = destStartRow + row_cnt


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

' 指定シート名のワークシートを探索（半角全角・大文字小文字は同一）
' wb: ワークブック
' sheetName: シート名
Private Function GetWorkSheet(wb As Workbook, sheetName As String, Optional make_sheet As Boolean = False) As Worksheet
    Dim ws As Variant
    
    For Each ws In wb.Worksheets
        If StrFormat(ws.name) = StrFormat(sheetName) Then
            Set GetWorkSheet = wb.Worksheets(ws.name)
            Exit Function
        End If
    Next
     
    If make_sheet Then
        Set GetWorkSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetWorkSheet.name = StrFormat(sheetName)
    Else
        Set GetWorkSheet = Nothing
    End If
    
End Function

' 文字列の正規化(小文字＋半角化)
' s: 文字列
Private Function StrFormat(ByVal s As String) As String
    
    s = StrConv(s, vbLowerCase)
    s = StrConv(s, vbNarrow)
    
    StrFormat = s
End Function


' ホルダ指定
Private Function SelectDir() As String
    Dim CurrentFilePath As String
    
    CurrentFilePath = ThisWorkbook.path & "\"
    
    'ホルダ開く
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CurrentFilePath
        .AllowMultiSelect = False
        .Show
        
        If .SelectedItems.Count = 0 Then
            SelectDir = ""
            Exit Function
        End If
        
        SelectDir = .SelectedItems(1)
        
    End With
        
End Function


Attribute VB_Name = "main"
Option Explicit



Sub Run_Export()
    Dim xls_file As String
    Dim csv_file As String
    
    xls_file = "input.xlsx"
    csv_file = "output.csv"
    
    Call ExportWorksheetToCSV(xls_file, "sheetname", csv_file)

End Sub


'=========================================
' 指定シートを CSV で書き出し
'=========================================
Public Sub ExportWorksheetToCSV(xlsPath As String, sheetName As String, csvPath As String)

    Dim wbSrc As Workbook
    Dim wbTemp As Workbook
    Dim ws As Worksheet
    Dim appCalc As XlCalculation
    Dim scrUpdt As Boolean
    Dim evtEnabled As Boolean
    Dim wbAlreadyOpen As Boolean
    Dim f As Workbook
    
    On Error GoTo EH

    ' 高速化
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 既に開いているか検索
    For Each f In Application.Workbooks
        If LCase$(f.FullName) = LCase$(xlsPath) Then
            Set wbSrc = f
            wbAlreadyOpen = True
            Exit For
        End If
    Next

    ' 開いてなければ読取専用で開く
    If wbSrc Is Nothing Then
        If Len(dir$(xlsPath)) = 0 Then
            Call Err.Raise(53, , "指定ファイルが見つかりません: " & xlsPath)
        End If
        Set wbSrc = Application.Workbooks.Open(FileName:=xlsPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    ' 出力先パスをフルパスへ変換
    csvPath = Convert2FullPath(csvPath)
    
    ' 対象シート
    On Error Resume Next
    Set ws = wbSrc.Worksheets(sheetName)
    On Error GoTo EH
    If ws Is Nothing Then
        Call Err.Raise(9, , "指定シートが見つかりません: " & sheetName)
    End If

    ' 単一シートの新規ブックを作る
    ws.Copy
    Set wbTemp = ActiveWorkbook

    ' CSV保存（OSの区切り記号・既定コードページで保存）
    Application.DisplayAlerts = False   ' 上書き確認を抑止
    wbTemp.SaveAs FileName:=csvPath, FileFormat:=xlCSV, CreateBackup:=False
    Application.DisplayAlerts = True

CLEANUP:
    If Not wbTemp Is Nothing Then wbTemp.Close SaveChanges:=False
    If Not wbAlreadyOpen Then wbSrc.Close SaveChanges:=False

    ' 復元
    Application.ScreenUpdating = scrUpdt
    Application.EnableEvents = evtEnabled
    Application.Calculation = appCalc
    Application.DisplayAlerts = True
    Exit Sub

EH:
    On Error Resume Next
    If Not wbTemp Is Nothing Then wbTemp.Close SaveChanges:=False
    If Not wbSrc Is Nothing Then If Not wbAlreadyOpen Then wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = scrUpdt
    Application.EnableEvents = evtEnabled
    Application.Calculation = appCalc
    Application.DisplayAlerts = True
    MsgBox "CSV出力に失敗しました。" & vbCrLf & Err.Description, vbExclamation, "ExportWorksheetToCSV_SJIS"
End Sub

'=== ユーティリティ ==========================================

' outPath をフルパスへ変換
' 例:
'  - "out.csv"            → ThisWorkbook.path\out.csv
'  - "sub\out.csv"        → ThisWorkbook.path\sub\out.csv
'  - "C:\work\out.csv"    → そのまま
Private Function Convert2FullPath(outPath As String) As String
    Dim path As String
    Dim dir As String
    Dim hasDrive As Boolean
    
    If Len(outPath) = 0 Then
        Call Err.Raise(5, , "出力ファイル名が空です。")
    End If

    path = Replace(outPath, "/", "\")
    hasDrive = (InStr(path, ":") > 0)

    If Not hasDrive Then
        dir = ThisWorkbook.path
        path = dir & "\" & path
    End If

    Convert2FullPath = path
End Function

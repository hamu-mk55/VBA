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
' �w��V�[�g�� CSV �ŏ����o��
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

    ' ������
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ���ɊJ���Ă��邩����
    For Each f In Application.Workbooks
        If LCase$(f.FullName) = LCase$(xlsPath) Then
            Set wbSrc = f
            wbAlreadyOpen = True
            Exit For
        End If
    Next

    ' �J���ĂȂ���Γǎ��p�ŊJ��
    If wbSrc Is Nothing Then
        If Len(dir$(xlsPath)) = 0 Then
            Call Err.Raise(53, , "�w��t�@�C����������܂���: " & xlsPath)
        End If
        Set wbSrc = Application.Workbooks.Open(FileName:=xlsPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    ' �o�͐�p�X���t���p�X�֕ϊ�
    csvPath = Convert2FullPath(csvPath)
    
    ' �ΏۃV�[�g
    On Error Resume Next
    Set ws = wbSrc.Worksheets(sheetName)
    On Error GoTo EH
    If ws Is Nothing Then
        Call Err.Raise(9, , "�w��V�[�g��������܂���: " & sheetName)
    End If

    ' �P��V�[�g�̐V�K�u�b�N�����
    ws.Copy
    Set wbTemp = ActiveWorkbook

    ' CSV�ۑ��iOS�̋�؂�L���E����R�[�h�y�[�W�ŕۑ��j
    Application.DisplayAlerts = False   ' �㏑���m�F��}�~
    wbTemp.SaveAs FileName:=csvPath, FileFormat:=xlCSV, CreateBackup:=False
    Application.DisplayAlerts = True

CLEANUP:
    If Not wbTemp Is Nothing Then wbTemp.Close SaveChanges:=False
    If Not wbAlreadyOpen Then wbSrc.Close SaveChanges:=False

    ' ����
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
    MsgBox "CSV�o�͂Ɏ��s���܂����B" & vbCrLf & Err.Description, vbExclamation, "ExportWorksheetToCSV_SJIS"
End Sub

'=== ���[�e�B���e�B ==========================================

' outPath ���t���p�X�֕ϊ�
' ��:
'  - "out.csv"            �� ThisWorkbook.path\out.csv
'  - "sub\out.csv"        �� ThisWorkbook.path\sub\out.csv
'  - "C:\work\out.csv"    �� ���̂܂�
Private Function Convert2FullPath(outPath As String) As String
    Dim path As String
    Dim dir As String
    Dim hasDrive As Boolean
    
    If Len(outPath) = 0 Then
        Call Err.Raise(5, , "�o�̓t�@�C��������ł��B")
    End If

    path = Replace(outPath, "/", "\")
    hasDrive = (InStr(path, ":") > 0)

    If Not hasDrive Then
        dir = ThisWorkbook.path
        path = dir & "\" & path
    End If

    Convert2FullPath = path
End Function

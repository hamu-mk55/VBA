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
    
    ' �p�����[�^
    Set mainSheet = ThisWorkbook.Worksheets("main")
    
    xls_sheetname = "101"
    xls_startcell = "B3"
    dst_sheetname = "data"
    
    xls_import_cols = 7
    
    
    ' �z���_�w��
    xlsDirPath = SelectDir()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not Len(xlsDirPath) > 0 Then
        Exit Sub
    End If
    
    ' ������
    Set dataSheet = GetWorkSheet(ThisWorkbook, dst_sheetname, True)
    dataSheet.Cells.Clear
    
    
    '�t�H���_���̃t�@�C������
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



' �w��u�b�N�E�w��V�[�g���w��V�[�g�֓]�L
' �K�v�ɉ����ď����R�s�[�ikeepFormats:=True�j
' srcPath: �ǂݍ��ރG�N�Z���t�@�C���̃t���p�X
' srcSheetName: �ǂݍ��ރG�N�Z���t�@�C���̃V�[�g��
' srcStartCell: �ǂݍ��ރG�N�Z���t�@�C���̓Ǎ��J�n�Z��(C3�Ȃ�)
' srcImportColNum: �ǂݍ��ރG�N�Z���t�@�C���̓Ǎ���(C~E���ǂݍ��ޏꍇ��3)
' destSheetName: �o�͐�̃V�[�g��
' destStartRow: �o�͐�̏o�͊J�n�s(>1�̂Ƃ��A2�t�@�C���ڂƔF��)
' keepFormats: �����R�s�[���邩�ǂ���
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

    ' ������
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' �Ώۃt�@�C�������ɊJ���Ă��邩����
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

    ' �Ǎ��V�[�g�擾�B������΃G���[���b�Z�[�W�o������ŏ����I��
    Set wsSrc = GetWorkSheet(wbSrc, srcSheetName, False)
    If wsSrc Is Nothing Then
        MsgBox ("No sheetname for import: " & vbLf & wbSrc.name & vbLf & srcSheetName)
        GoTo CLEANUP
    End If

    ' �o�̓V�[�g�擾�B������Βǉ�
    Set wsDst = GetWorkSheet(ThisWorkbook, destSheetName, True)
    
    ' �f�[�^�]�L����
    Dim rngSrc As Range
    Dim srcRow As Long
    Dim srcCol As Long
    Dim lastCol As Long
    
    Dim rngDst As Range
    
    Dim row_cnt As Long
    Dim MAX_LOOP As Long
    
    ' �o�͊J�n�s(���A��Œ�)
    Set rngDst = wsDst.Range("A" & destStartRow)
    
    ' �Ǎ��J�n�Z�����w��
    Set rngSrc = wsSrc.Range(srcStartCell)
    srcRow = rngSrc.Row
    srcCol = rngSrc.Column
    
    lastCol = srcCol + srcImportColNum - 1
    
    ' startRow��1�Ŗ���(=2��ڈȍ~)�ꍇ�́A�w�b�_SKIP
    If destStartRow > 1 Then
        srcRow = srcRow + 1
    End If
    
    ' �������[�v�h�~
    MAX_LOOP = wsSrc.rows.Count
    
    row_cnt = 0
    Do While LenB(CStr(wsSrc.Cells(srcRow, srcCol).Value)) > 0
    
        ' �������K�v�Ȃ�A���1�s�R�s�[
        If keepFormats Then
            wsSrc.Cells(srcRow, srcCol).Resize(1, srcImportColNum).Copy _
                Destination:=rngDst.offset(row_cnt, 0)
        End If
    
        ' �l��]�L(�P�s)
        rngDst.offset(row_cnt, 0).Resize(1, srcImportColNum).Value = _
            wsSrc.Cells(srcRow, srcCol).Resize(1, srcImportColNum).Value
    
        ' ���̍s��
        srcRow = srcRow + 1
        row_cnt = row_cnt + 1
        
        ' MAX_LOOP�ȏ�Ń��[�v������
        If row_cnt > MAX_LOOP Then
            Exit Do
        End If
        
        ' �t���[�Y�΍�i���Ԋu��DoEvents�j
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
    MsgBox "�]�L�Ɏ��s���܂����B" & vbCrLf & Err.Description, vbExclamation, "CopyToDataSheet"
End Sub

' �w��V�[�g���̃��[�N�V�[�g��T���i���p�S�p�E�啶���������͓���j
' wb: ���[�N�u�b�N
' sheetName: �V�[�g��
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

' ������̐��K��(�������{���p��)
' s: ������
Private Function StrFormat(ByVal s As String) As String
    
    s = StrConv(s, vbLowerCase)
    s = StrConv(s, vbNarrow)
    
    StrFormat = s
End Function


' �z���_�w��
Private Function SelectDir() As String
    Dim CurrentFilePath As String
    
    CurrentFilePath = ThisWorkbook.path & "\"
    
    '�z���_�J��
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


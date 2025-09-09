Attribute VB_Name = "Module2"
Option Explicit

Sub Run_CopyToData()
    Dim xls As String
    xls = "C:\work\report.xlsx"

    ' 1) �l�݂̂ŏ㏑���]�L�idata�V�[�g�����O�ɑS�����j
    CopyToDataSheet xls, "ExportMe", "data", True, "A1", False

    ' 2) �������R�s�[�������ꍇ
    'CopyToDataSheet xls, "ExportMe", "data", True, "A1", True
End Sub


'=========================================
' �w��u�b�N�E�w��V�[�g �� ThisWorkbook("data") �֓]�L
' �E�l�݂̂��ꊇ�]�L�i�����j
' �E�K�v�ɉ����ď����R�s�[���ikeepFormats:=True�j
' �E�t�@�C����������srcPath�ɂ��Ή�
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

    '=== �����ꎞ�I�Ɍy�ʉ� ===
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '=== srcPath �������i����/�t�@�C�����̂ݑΉ��j===
    srcPath = ResolvePath(srcPath, ThisWorkbook.path)

    '=== �\�[�X�u�b�N���擾�i���ɊJ���Ă���΂�����g�p�j===
    For Each f In Application.Workbooks
        If LCase$(f.FullName) = LCase$(srcPath) Then
            Set wbSrc = f
            wasOpen = True
            Exit For
        End If
    Next f
    If wbSrc Is Nothing Then
        If Len(dir$(srcPath)) = 0 Then Err.Raise 53, , "�w��t�@�C����������܂���: " & srcPath
        Set wbSrc = Application.Workbooks.Open(FileName:=srcPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    '=== �\�[�X/�f�X�e�B�l�[�V���� �V�[�g�擾�i������΍쐬�j===
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(srcSheetName)
    On Error GoTo EH
    If wsSrc Is Nothing Then Err.Raise 9, , "�w��V�[�g��������܂���: " & srcSheetName

    Set wsDst = GetOrCreateSheet(ThisWorkbook, destSheetName)

    '=== �]�L��J�n�Z�� ===
    Set dst = wsDst.Range(startCell)

    '=== �]�L�O�N���A ===
    If clearDest Then
        wsDst.Cells.Clear ' �����������B�������c�������ꍇ�� ClearContents �ɕύX
    End If

    '=== �\�[�X�͈� ===
    Set rngSrc = wsSrc.UsedRange
    If Application.WorksheetFunction.CountA(rngSrc) = 0 Then
        ' ��Ȃ牽�������I��
        GoTo CLEANUP
    End If

    '=== �������~�����ꍇ�͐�ɃR�s�[�i��Œl�ŏ㏑���j===
    If keepFormats Then
        rngSrc.Copy Destination:=dst
    End If

    '=== �l�݈̂ꊇ�]�L�i�����F�z�񁨈ꊇ����j===
    arr = rngSrc.Value
    If IsArray(arr) Then
        r = UBound(arr, 1) - LBound(arr, 1) + 1
        c = UBound(arr, 2) - LBound(arr, 2) + 1
        dst.Resize(r, c).Value = arr
    Else
        ' �P��Z���̏ꍇ
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
    MsgBox "�]�L�Ɏ��s���܂����B" & vbCrLf & Err.Description, vbExclamation, "CopyToDataSheet"
End Sub

' �V�[�g�擾�B������Ζ����ɒǉ����ĕԂ�
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.name = name
    End If
End Function



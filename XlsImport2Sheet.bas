Attribute VB_Name = "Module2"
Option Explicit

Sub Run_Import2Sheet()
    Dim xls_file As String
    Dim xls_sheetname As String
    Dim dst_sheetname As String
    Dim row_cnt As Double
    
    xls_file = "C:\Users\�f�[�^.xlsx"
    xls_sheetname = "�f�[�^"
    dst_sheetname = "data"

    row_cnt = 1
    Call CopyToDataSheet(xls_file, xls_sheetname, dst_sheetname, True, row_cnt, True)

    Debug.Print (row_cnt)

End Sub


' �w��u�b�N�E�w��V�[�g���w��V�[�g�֓]�L
' �K�v�ɉ����ď����R�s�[�ikeepFormats:=True�j
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

    ' ������
    scrUpdt = Application.ScreenUpdating
    evtEnabled = Application.EnableEvents
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' �Ώۃt�@�C�������ɊJ���Ă��邩����
    For Each f In Application.Workbooks
        If LCase$(f.FullName) = LCase$(xlsPath) Then
            Set wbSrc = f
            wbOpen = True
            Exit For
        End If
    Next f
    If wbSrc Is Nothing Then
        If Len(dir$(xlsPath)) = 0 Then
            Call Err.Raise(53, , "�w��t�@�C����������܂���: " & xlsPath)
        End If
        Set wbSrc = Application.Workbooks.Open(FileName:=xlsPath, ReadOnly:=True, UpdateLinks:=False)
    End If

    ' �Ǎ��V�[�g�擾
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(xlsSheetName)
    On Error GoTo EH
    If wsSrc Is Nothing Then
        Call Err.Raise(9, , "�w��V�[�g��������܂���: " & xlsSheetName)
    End If

    ' �o�̓V�[�g�擾
    Set wsDst = SetSheet(ThisWorkbook, destSheetName)

    ' �o�̓V�[�g������
    If clearDest Then
        wsDst.Cells.Clear
    End If

    ' �o�͊J�n�s
    Set rngDst = wsDst.Range("A" & startRow)

    ' �\�[�X�͈�
    Dim srcStart As Range
    Dim srcRow As Long
    Dim srcCol As Long
    Dim dstRowOff As Long
    Dim lastCol As Long
    Dim cCount As Long
    Dim MAX_GUARD As Long
    
    ' �N�_�Z�����w��
    Set srcStart = wsSrc.Range("B3")
    
    srcRow = srcStart.Row
    srcCol = srcStart.Column
    dstRowOff = 0
    
    ' �O�̂��ߖ������[�v�h�~�i�傫�߂Ɂj
    MAX_GUARD = wsSrc.rows.Count
    
    Do While LenB(CStr(wsSrc.Cells(srcRow, srcCol).Value)) > 0 And MAX_GUARD > 0
        ' ���̍s�̍ŏI�g�p��i���[����E�ցj
        lastCol = wsSrc.Cells(srcRow, wsSrc.Columns.Count).End(xlToLeft).Column
        If lastCol < srcCol Then
            ' �N�_���荶�Ŏ~�܂��Ă��遁�����f�[�^�����Ƃ݂Ȃ��ďI��
            Exit Do
        End If
    
        cCount = lastCol - srcCol + 1
    
        ' �������K�v�Ȃ�A���1�s�R�s�[�ikeepFormats=True�̂Ƃ��j
        If keepFormats Then
            wsSrc.Cells(srcRow, srcCol).Resize(1, cCount).Copy _
                Destination:=rngDst.Offset(dstRowOff, 0)
        End If
    
        ' �l�����㏑���i1�s���j
        rngDst.Offset(dstRowOff, 0).Resize(1, cCount).Value = _
            wsSrc.Cells(srcRow, srcCol).Resize(1, cCount).Value
    
        ' ���̍s��
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
    MsgBox "�]�L�Ɏ��s���܂����B" & vbCrLf & Err.Description, vbExclamation, "CopyToDataSheet"
End Sub

' �V�[�g���Z�b�g�B�����Ŗ�����Ζ����ɒǉ�
Private Function SetSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set SetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If SetSheet Is Nothing Then
        Set SetSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        SetSheet.name = sheetName
    End If
End Function



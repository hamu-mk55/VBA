Attribute VB_Name = "CleanSheet"
Option Explicit


Sub CleanSheet()
    Dim mainSheet As Worksheet
    Dim dataSheet As Worksheet
    
    Dim keepCols As Variant
    
    ' �p�����[�^
    Set mainSheet = ThisWorkbook.Worksheets("main")
    Set dataSheet = ThisWorkbook.Worksheets("data")
    
    ' ��\���s������
    Call DeleteHiddenRows(dataSheet)
    
    ' �ێ��������V�[�g����擾���A����ȊO�̗���폜
    keepCols = GetStrList(mainSheet)
    Call DeleteColsExceptArr(dataSheet, keepCols, 0)

End Sub


' ��\���̍s���폜
Sub DeleteHiddenRows(ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    
    If ws.AutoFilterMode Then
        lastRow = ws.AutoFilter.Range(ws.AutoFilter.Range.Count).Row
    Else
        lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    End If
        
    ' �w�b�_�[�ȊO���c���āA�s��������T���B��\���s������
    For r = lastRow To 2 Step -1
        If ws.rows(r).Hidden Then
            ws.rows(r).Delete
        End If
    Next r

End Sub


' ws �̃V�[�g�� keepCols �Ɋ܂܂�Ȃ����S�폜
' ws: ���[�N�V�[�g
' keepCols: �ێ���������L���̔z��(A,B,C)
' offset: �z��ŎQ�Ƃ�����ԍ��ƁA���[�N�V�[�g�̗�ԍ��̃I�t�Z�b�g
' ��F�z�񒆂�B��ɑΉ�����񂪁A���[�N�V�[�g����A��̏ꍇ�Aoffset=1
Private Sub DeleteColsExceptArr(ws As Worksheet, keepCols As Variant, offset As Long)
    Dim maxCol As Long
    
    Dim i As Long
    Dim col_no As Long
    Dim keep_dict As Object

    Set keep_dict = CreateObject("Scripting.Dictionary")
    keep_dict.CompareMode = 1 ' �啶������������ʂ��Ȃ�

    ' ��ԍ��ɕϊ����������Ŏ����o�^
    ' �����@key: ��ԍ��Aval: True/False
    For i = LBound(keepCols) To UBound(keepCols)
        col_no = ColumnLetterToNumber(CStr(keepCols(i)))
        keep_dict(col_no - offset) = True
    Next

    maxCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                           SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    ' �폜�ɂ���ԍ�����̉e�����󂯂Ȃ��悤�ɁA�E�����ɍ폜
    For i = maxCol To 1 Step -1
        If Not keep_dict.Exists(i) Then
            ws.Columns(i).Delete
        End If
    Next

End Sub

' ��L��(A,B,C)���ԍ�(1,2,3)�ɕϊ�
' s: ��L��
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

' �V�[�g�̃Z�����e��ǂݎ��A�z��Ɋi�[
' ws: ���[�N�V�[�g
Private Function GetStrList(ws As Worksheet) As Variant
    Dim startRow As Long
    Dim startCol As Long
    
    Dim i As Long
    Dim tmp As Collection
    Dim arr() As Variant

    Set tmp = New Collection

    ' �J�n�s/��̐ݒ�
    startRow = 1
    startCol = 2

    ' �w��s/�񂩂�J�n���A�Z���l��z��Ɋi�[�B
    ' �󔒃Z���܂ŉ����֒T���B
    For i = startRow To ws.rows.Count
        If LenB(Trim(ws.Cells(i, startCol).Value)) > 0 Then
            tmp.Add Trim(ws.Cells(i, startCol).Value)
        End If
    Next i

    ' Collection �� �z��ɕϊ�
    ReDim arr(0 To tmp.Count - 1)
    For i = 1 To tmp.Count
        arr(i - 1) = tmp(i)
    Next i

    GetStrList = arr
End Function


Attribute VB_Name = "csv"
Option Explicit

Sub SaveFile()
    Dim DataSheetName As String
    Dim CsvFileName As String
    
    DataSheetName = "preview"
    CsvFileName = "test.csv"
    
    Call save2csvfile(DataSheetName, CsvFileName, 1, 1, 19)


End Sub


'�V�[�g�̓��e��CSV�t�@�C���֏o��
Private Sub save2csvfile(DataSheetName As String, _
                            CsvFileName As String, _
                            start_row As Long, _
                            output_start_col As Long, _
                            output_end_col As Long)
    
    Dim DataSheet As Worksheet
    Dim CsvFile As Variant
    Dim FSO As Object
    Dim TS As TextStream
    
    Dim R As Range
    Dim row_num As Long
    
    Dim i As Long
    
    '�V�[�g�m�F
    If CheckSheetExist(DataSheetName) Then
        Set DataSheet = Worksheets(DataSheetName)
    Else
        MsgBox ("�f�[�^�V�[�g������܂���" & vbLf & DataSheetName)
        Exit Sub
    End If
  
    'CSV�o��
    CsvFile = Application.GetSaveAsFilename(InitialFileName:=CsvFileName, _
                                            FileFilter:="CSV�t�@�C��(*.csv),*.csv", _
                                            FilterIndex:=1, _
                                            Title:="�ۑ��t�@�C���̎w��")
    If CsvFile = False Then
        MsgBox ("�t�@�C�����I������܂���ł���")
        Exit Sub
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TS = FSO.OpenTextFile(FileName:=CsvFile, IOMode:=ForWriting, Create:=True)
    
    For i = 2 To 10
        Debug.Print DataSheet.Cells(i, 2), DataSheet.rows(i).Hidden
    Next i
    
    
    'With wsheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible)
    '    row_num = .Count - 1
    '    Debug.Print "NUM: "
    '    Debug.Print row_num
    '    For Each R In .rows
    '        Debug.Print R.Cells(1, 6)
    '    Next
    'End With
        
    With DataSheet.Cells(start_row, 1).CurrentRegion.Offset(1, 0)
        row_num = .Resize(.rows.count - 1).Columns(1).SpecialCells(xlCellTypeVisible).count
        Debug.Print "NUM: " & row_num
        
        For Each R In .Resize(.rows.count - 1).SpecialCells(xlCellTypeVisible).rows
            TS.WriteLine output_rows(R, output_start_col, output_end_col)
        Next R
    End With
  
    TS.Close
    Set TS = Nothing
    Set FSO = Nothing

    MsgBox ("CSV�o�͊���")
    
End Sub

'Range�̓��e���ACSV�`���ň�s�o��
Private Function output_rows(ByVal row As Range, _
                                start_col As Long, _
                                end_col As Long) As String
    Dim output As String
    Dim tmp As String
    Dim i As Long

    output = ""
    For i = start_col To end_col
        Select Case True
            '��
            Case IsEmpty(row.Cells(1, i))
                tmp = ""
        
            '�����̏ꍇ�͕�����ɕϊ�
            Case IsNumeric(row.Cells(1, i))
                tmp = CStr(CDbl(row.Cells(1, i)))
            
            '���̑�
            Case Else
                tmp = CStr(row.Cells(1, i))
                tmp = Replace(tmp, vbLf, "")
                tmp = Replace(tmp, ",", "")
        End Select
        
        If output = "" Then
            output = tmp
        Else
            output = output & "," & tmp
        End If
        
    Next
    
    output_rows = output

End Function



Attribute VB_Name = "xml"
Option Explicit

Sub CheckSingleFile()
    Dim FilePath As String
    Dim row_cnt As Long
    
    FilePath = SelectXmlFile()
    
    row_cnt = 2
    If Len(FilePath) > 0 Then
        Call ReadXmlFile(FilePath, row_cnt)
    End If
    
End Sub

Sub CheckFiles()
    Dim DirPath As String
    Dim FSO As Object
    Dim row_cnt As Long
    
    DirPath = SelectDir()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Len(DirPath) > 0 Then
        row_cnt = 2
        Call ReadAllFiles(DirPath, FSO, row_cnt)
        
        Set FSO = Nothing
    End If

End Sub

Private Function SelectXmlFile() As String
    Dim CurrentFilePath As String
    Dim SelectFilePath As String
    
    
    CurrentFilePath = ThisWorkbook.Path & "\"
    
    '�t�@�C���J��
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "xml", "*.xml"
        .InitialFileName = CurrentFilePath
        .AllowMultiSelect = False
        .Show
        
        If .SelectedItems.count = 0 Then
            SelectXmlFile = ""
            Exit Function
        End If
        
        SelectFilePath = .SelectedItems(1)
        
    End With
    
    SelectXmlFile = SelectFilePath
        
End Function

Private Function SelectDir() As String

    Dim CurrentFilePath As String
    
    CurrentFilePath = ThisWorkbook.Path & "\"
    
    '�t�H���_�J��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CurrentFilePath
        .AllowMultiSelect = False
        .Show
        
        If .SelectedItems.count = 0 Then
            SelectDir = ""
            Exit Function
        End If
        
        SelectDir = .SelectedItems(1)
        
    End With
        
End Function

Private Sub ReadAllFiles(ByVal inputDir As String, ByVal FSO As Object, ByRef row_cnt As Long)
    Dim fileType As String
    Dim xmldir As Object
    Dim xmlfile As Object
    Dim xmlpath As String
    
    Dim msg_res As Integer
    
    fileType = "xml"
    
    ' row_cnt�������������m�F����
    If (row_cnt - 1) > 10000 Then
        msg_res = MsgBox("�f�[�^���F" & (row_cnt - 1) & vbCrLf & _
                         "�t�H���_��" & inputDir & vbCrLf & _
                          "�����܂����H")
        If msg_res = vbNo Then
            Exit Sub
        End If
    End If
    
    '�t�H���_���̃T�u�t�H���_���ċA����
    For Each xmldir In FSO.getFolder(inputDir).SubFolders
        Call ReadAllFiles(xmldir.Path, FSO, row_cnt)
    Next
    
    '�t�H���_���̃t�@�C������
    For Each xmlfile In FSO.getFolder(inputDir).Files
        If LCase(FSO.GetExtensionName(xmlfile.Name)) = fileType Then
            xmlpath = inputDir & "\" & xmlfile.Name
        
            Call ReadXmlFile(xmlpath, row_cnt)
        End If
    Next
        
End Sub

'�P��t�@�C�����J��
'param: FilePath    XML�t�@�C���̃t���p�X
'param: row_cnt
Private Sub ReadXmlFile(ByVal FilePath As String, ByRef row_cnt As Long)
    
    Dim CurrentWorkSheet As Worksheet
    Dim TableWorkSheet As Worksheet
    Dim TableName As String
    
    Dim StartDate, EndDate As Date
    Dim SearchKey As String
    
    Dim xml As DOMDocument60
    Dim xml_items As IXMLDOMNodeList
    Dim xml_item As IXMLDOMNode
    
    Dim isFirst As Boolean
    
    '�p�����[�^�ݒ�
    TableName = "preview"
    StartDate = Range("C4").Text
    EndDate = Range("C5").Text
    SearchKey = Range("C6").Text

    If row_cnt <= 2 Then
        isFirst = True
    Else
        isFirst = False
    End If
    
    '�V�[�g������
    Set CurrentWorkSheet = ActiveSheet
    
    If CheckSheetExist(TableName) Then
        Set TableWorkSheet = Worksheets(TableName)
        If isFirst Then
            TableWorkSheet.Cells.Clear
        End If
    Else
        If isFirst Then
            Set TableWorkSheet = Worksheets.Add(after:=CurrentWorkSheet)
            TableWorkSheet.Name = TableName
        Else
            MsgBox ("ERROR: No Sheet")
            Exit Sub
        End If
    End If
    
    CurrentWorkSheet.Activate
        
    '�ǂݍ���
    Dim xr As XmlReader
    Set xr = New XmlReader
    
    Call xr.SetFirstMode(isFirst)
    Call xr.LoadXmlFile(FilePath)
    Call xr.SetWorkSheet(TableWorkSheet)
    Call xr.SetSearchKey(SearchKey)
    Call xr.SetDate(StartDate, EndDate)
    
    Call xr.GetDefectList(row_cnt)
    
    With TableWorkSheet.rows
        .WrapText = False
        .AutoFit
    End With
    
End Sub





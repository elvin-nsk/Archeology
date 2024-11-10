Attribute VB_Name = "Archeology"
'===============================================================================
'   ������          : Archeology
'   ������          : 2024.11.10
'   �����           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Archeology"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_FILEBASENAME As String = "elvin_" & APP_NAME
Public Const APP_VERSION As String = "2024.11.10"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const TEXT_PREFIX As String = "���. "
Private Const CDR_FOLDER_NAME As String = "CDR"
Private Const PDF_FOLDER_NAME As String = "PDF"

'===============================================================================
' # Entry points

Sub Rename()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
    End With
    
    Dim StartingNumber As Long
    If Not TryGetStartingNumberFromFileName(StartingNumber) Then
        Warn "�� ������ ��������� ����� � �������� �����.", APP_DISPLAYNAME
        Exit Sub
    End If
        
    If Not IsThereAnyValidShape Then
        Warn "�� ������� ���������� ��������� ��������.", APP_DISPLAYNAME
        Exit Sub
    End If
    
    BoostStart "��������� � ������"
        
    RenameInActiveDoc StartingNumber
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub RenameInFolder()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
    End With
    
    If ActiveDocument.Dirty Then
        Notify "��������� �������� ����� ��������.", APP_DISPLAYNAME
        Exit Sub
    End If
    
    Dim StartingNumber As Long: StartingNumber = 1
    If Not AskForLong("��������� �����:", StartingNumber, APP_DISPLAYNAME) Then
        Exit Sub
    End If
    
    Dim RootPath As String: RootPath = ActiveDocument.FilePath
    ActiveDocument.Close
    
    #If DebugMode = 0 Then
    Optimization = True
    #End If
    
    OpenRenameSaveAsAndExportForPath RootPath, StartingNumber
    
Finally:
    Optimization = False
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Function TryGetStartingNumberFromFileName( _
                     ByRef StartingNumber As Long _
                 ) As Boolean
    Dim Result As Variant: Result = FindFirstInteger(ActiveDocument.FileName)
    If Not VBA.IsEmpty(Result) Then
        StartingNumber = Result
        TryGetStartingNumberFromFileName = True
    End If
End Function

Private Property Get IsThereAnyValidShape() As Boolean
    Dim Page As Page
    Dim Shape As Shape
    For Each Page In ActiveDocument.Pages
        For Each Shape In Page.Shapes
            If IsValid(Shape) Then
                IsThereAnyValidShape = True
                Exit Property
            End If
        Next Shape
    Next Page
End Property

Private Property Get IsValid(ByVal Shape As Shape) As Boolean
    If Shape.Type <> cdrTextShape Then Exit Property
    Dim Name As String: Name = Shape.Text.Story.Text
    If Not Name Like TEXT_PREFIX & "#*" Then Exit Property
    IsValid = True
End Property

Private Sub RenameInActiveDoc( _
                ByVal StartingNumber As Long, _
                Optional ByRef LastNumber As Long _
            )
    Dim Counter As Long: Counter = StartingNumber
    Dim Page As Page
    Dim Shape As Shape
    Dim ShapesOnPage As ShapeRange
    
    For Each Page In ActiveDocument.Pages
        
        Set ShapesOnPage = CreateShapeRange
        For Each Shape In Page.Shapes
            If IsValid(Shape) Then ShapesOnPage.Add Shape
        Next Shape
        
        ShapesOnPage.Sort "@Shape1.Top > @Shape2.Top"
        
        For Each Shape In ShapesOnPage
            ReplaceNumber Shape, Counter
            Counter = Counter + 1
        Next Shape
    
    Next Page
    LastNumber = Counter - 1
End Sub

Private Sub ReplaceNumber(ByVal Shape As Shape, ByVal Number As Long)
    Dim Story As TextRange: Set Story = Shape.Text.Story
    Dim Text As String: Text = Story.Text
    Dim LastDigitPosition As Long
    For LastDigitPosition = Len(TEXT_PREFIX) + 1 To Story.Characters.Count
        If Not Mid(Text, LastDigitPosition + 1, 1) Like "#" Then Exit For
    Next LastDigitPosition
    Story.Range(0, LastDigitPosition).Replace TEXT_PREFIX & CStr(Number)
End Sub

Private Sub OpenRenameSaveAsAndExportForPath( _
                ByVal RootPath As String, _
                ByVal StartingNumber As Long _
            )
    Dim CdrPath As String: CdrPath = MakePath(RootPath & CDR_FOLDER_NAME)
    Dim PdfPath As String: PdfPath = MakePath(RootPath & PDF_FOLDER_NAME)
    Dim Files As Collection: Set Files = GetFilesFromFolder(RootPath)
    SortFiles Files
    
    Dim NextNumber As Long: NextNumber = StartingNumber
    Dim File As File
    For Each File In Files
        TryOpenRenameSaveAsAndExportFile _
            File.Path, CdrPath, PdfPath, NextNumber
    Next File
End Sub

Private Sub TryOpenRenameSaveAsAndExportFile( _
                ByVal FileCandidate As String, _
                ByVal CdrPath As String, _
                ByVal PdfPath As String, _
                ByRef NextNumber As Long _
            )
    Dim File As FileSpec: Set File = FileSpec.New_(FileCandidate)
    If Not File.Ext = "cdr" Then Exit Sub
    OpenDocument File
    
    Dim LastNumber As Long
    RenameInActiveDoc NextNumber, LastNumber
    File.Name = GetFileNameWithReplacedNumbers(File.Name, NextNumber, LastNumber)
    File.Path = CdrPath
    
    SaveAs File
    File.Path = PdfPath
    File.Ext = "pdf"
    ExportPdf File
    ActiveDocument.Close
    
    NextNumber = LastNumber + 1
End Sub

Private Sub SortFiles(ByVal Files As Collection)
    If Files.Count < 2 Then Exit Sub
    Dim i As Long, j As Long
    Dim Temp As File
    'Two loops to bubble sort
    For i = 1 To Files.Count - 1
        For j = i + 1 To Files.Count
            If ToComparable(Files(i).Name) _
             > ToComparable(Files(j).Name) Then
                'store the lesser item
                Set Temp = Files(j)
                'remove the lesser item
                Files.Remove j
                're-add the lesser item before the
                'greater Item
                Files.Add Item:=Temp, Before:=i
            End If
        Next j
    Next i
End Sub

Private Function SaveAs(ByVal File As String)
    Dim Options As New StructSaveAsOptions
    With Options
        .Version = cdrVersion15
    End With
    ActiveDocument.SaveAs File, Options
End Function

Private Function ExportPdf(ByVal File As String)
    With ActiveDocument.PDFSettings
        .PublishRange = pdfWholeDocument
    End With
    ActiveDocument.PublishToPDF File
End Function

Private Property Get ToComparable(ByVal s As String) As String
    Dim Pos As Long
    Pos = VBA.InStr(1, s, "=")
    If Pos = 0 Then
        ToComparable = s
    Else
        ToComparable = Mid(s, Pos)
    End If
End Property

Private Property Get GetFileNameWithReplacedNumbers( _
                ByVal FileName As String, _
                ByVal FirstNumber As Long, _
                ByVal LastNumber As Long _
            )
    Dim Pos As Long
    Pos = VBA.InStr(1, FileName, "=")
    If Pos = 0 Then Exit Property
    GetFileNameWithReplacedNumbers = _
        "���_" _
      & VBA.Format(FirstNumber, "0000") _
      & "-" _
      & VBA.Format(LastNumber, "0000") _
      & Mid(FileName, Pos)
End Property

'===============================================================================
' # Tests

Private Sub TestSomething()
    Dim cFruit As New Collection
    'fill the collection
    cFruit.Add "���_0013-0016=������2_����������5.cdr"
    cFruit.Add "���_0005-0008=������2_����������2.cdr"
    cFruit.Add "���_0009-0012=������2_����������4.cdr"
    
    SortCollection cFruit
    Show cFruit
End Sub

Private Sub Test2()
    Show ToComparable("���_0013-0016=������2_����������5.cdr")
End Sub

Private Sub Test3()
    Dim Files As Collection
    Set Files = GetFilesFromFolder("d:\temp\images\")
    Show Files
    SortFiles Files
    Show Files
End Sub

Private Sub Test4()
    Dim Name As String
    Name = "���_0013-0016=������2_����������5.cdr"
    Show GetFileNameWithReplacedNumbers(Name, 6, 10)
End Sub

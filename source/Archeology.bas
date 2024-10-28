Attribute VB_Name = "Archeology"
'===============================================================================
'   Макрос          : Archeology
'   Версия          : 2024.10.28
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Archeology"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_FILEBASENAME As String = "elvin_" & APP_NAME
Public Const APP_VERSION As String = "2024.10.28"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const TEXT_PREFIX As String = "Илл. "
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
        Warn "Не найден стартовый номер в названии файла.", APP_DISPLAYNAME
        Exit Sub
    End If
        
    If Not IsThereAnyValidShape Then
        Warn "Не найдено подходящих текстовых объектов.", APP_DISPLAYNAME
        Exit Sub
    End If
    
    BoostStart "Нумерация в тексте"
        
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
        Notify "Сохраните документ перед запуском.", APP_DISPLAYNAME
        Exit Sub
    End If
    
    Dim RootPath As String: RootPath = ActiveDocument.FilePath
    ActiveDocument.Close
    
    #If DebugMode = 0 Then
    Optimization = True
    #End If
    
    OpenRenameSaveAsAndExportForPath RootPath
    
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

Private Sub RenameInActiveDoc(ByVal StartingNumber As Long)
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

Private Sub OpenRenameSaveAsAndExportForPath(ByVal RootPath As String)
    Dim CdrPath As String: CdrPath = MakePath(RootPath & CDR_FOLDER_NAME)
    Dim PdfPath As String: PdfPath = MakePath(RootPath & PDF_FOLDER_NAME)
    Dim File As File
    For Each File In FSO.GetFolder(RootPath).Files
        TryOpenRenameSaveAsAndExportFile File.Path, CdrPath, PdfPath
    Next File
End Sub

Private Sub TryOpenRenameSaveAsAndExportFile( _
                ByVal FileCandidate As String, _
                ByVal CdrPath As String, _
                ByVal PdfPath As String _
            )
    Dim File As FileSpec: Set File = FileSpec.New_(FileCandidate)
    If Not File.Ext = "cdr" Then Exit Sub
    Dim StartingNumber As Long
    OpenDocument File
    If TryGetStartingNumberFromFileName(StartingNumber) Then
        RenameInActiveDoc StartingNumber
        File.Path = CdrPath
        SaveAs File
        File.Path = PdfPath
        File.Ext = "pdf"
        ExportPdf File
    End If
    ActiveDocument.Close
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

'===============================================================================
' # Tests

Private Sub testSomething()
    OpenDocument "e:\WORK\макросы Corel\на заказ\НПЦ Денис\Archeology\материалы\тестовая папка\илл_0001-0000=Р06-погребение_002 КАК ОБРАЗЕЦ для скрипта.cdr"
End Sub

Attribute VB_Name = "Archeology"
'===============================================================================
'   Макрос          : Archeology
'   Версия          : 2024.10.25
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
Public Const APP_VERSION As String = "2024.10.25"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const TEXT_PREFIX As String = "Илл. "

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
        Warn "Не найден стартовый номер в названии файла."
        Exit Sub
    End If
        
    If Not IsThereAnyValidShape Then
        Warn "Не найдено подходящих текстовых объектов."
        Exit Sub
    End If
    
    BoostStart "Нумерация в тексте"
        
    RenameInActiveDoc StartingNumber
    
    '??? PROFIT!
    
Finally:
    BoostFinish
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

'===============================================================================
' # Tests

Private Sub testSomething()
    ReplaceNumber ActiveShape, 555
End Sub

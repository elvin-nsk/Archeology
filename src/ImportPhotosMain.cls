VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportPhotosMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    Cfg As Dictionary
    Progress As ProgressBar
End Type
Private This As This

'===============================================================================
' # Constructors

Private Sub Class_Initialize()
    With This
    End With
End Sub

Private Sub Class_Terminate()
    This = BlankThis
End Sub

Private Property Get BlankThis() As This
End Property

'===============================================================================
' # Main

Public Sub Run()
    If ShowImportPhotosView(This.Cfg) = Fail Then Exit Sub
    
    If Not FSO.FolderExists(This.Cfg!WorkPath) Then
        Warn "Путь к рабочей папке не найден."
        Exit Sub
    End If
    If Not FSO.FileExists(This.Cfg!TemplateFile) Then
        Warn "Не найден файл шаблона."
        Exit Sub
    End If
    
    Dim ValidFiles As Collection: GetValidFiles This.Cfg!WorkPath, ValidFiles
    'Show ValidFiles.Count
    
    If ValidFiles.Count = 0 Then
        Warn "В рабочей папке не найдено подходящих файлов."
        Exit Sub
    End If
    
    Dim Doc As Document: Set Doc = Application.OpenDocument(This.Cfg!TemplateFile)
    If ValidateTemplate = Fail Then
        Warn "На шаблоне отсутствуют места размещения."
        Exit Sub
    End If
    
    #If DebugMode = 0 Then
    Optimization = True
    #End If
        
    AddPages ValidFiles.Count
    Set This.Progress = ProgressBar.New_(ProgressBarNumeric, ValidFiles.Count)
    This.Progress.Caption = "Импорт..."
    Dim Images As ShapeRange: Set Images = ImportImages(ValidFiles)
    Images.SetPositionEx _
        cdrTopLeft, ActivePage.RightX + ActivePage.SizeWidth / 20, ActivePage.TopY
    PlaceImages Images
    SaveResult This.Cfg!WorkPath
    ActiveDocument.Close
End Sub

'===============================================================================
' # Helpers

Private Function ValidateTemplate() As BooleanResult
    Dim Page As Page
    For Each Page In ActiveDocument.Pages
        If IsNone(FindShape(Page, IMPORTER_UPPER_FRAME_NAME)) Then Exit Function
        If IsNone(FindShape(Page, IMPORTER_LOWER_FRAME_NAME)) Then Exit Function
    Next Page
    ValidateTemplate = Ok
End Function

Private Function ImportImages(ByVal Files As Collection) As ShapeRange
    Set ImportImages = CreateShapeRange
    ActivateImportLayer
    Dim File As Scripting.File
    For Each File In Files
        ActiveLayer.Import File
        ImportImages.Add ActiveShape
    Next File
End Function

Private Sub AddPages(ByVal ImagesCount As Long)
    Dim PagesCount As Long
    PagesCount = Clamp(Ceiling(ImagesCount / 2), , IMPORTER_MAX_PAGES)
    If ActiveDocument.Pages.Count >= PagesCount Then Exit Sub
    ActiveDocument.Pages.Last.Activate
    DuplicateActivePage PagesCount - ActiveDocument.Pages.Count
End Sub

Private Sub PlaceImages(ByVal Images As ShapeRange)
    Dim i As Long, Page As Page
    For i = 1 To Clamp(Images.Count, , IMPORTER_MAX_PAGES * 2) Step 2
        Set Page = _
            ActiveDocument.Pages( _
                Clamp(Ceiling(i / 2), , ActiveDocument.Pages.Count) _
            )
        PlaceImage Images(i), Page, IMPORTER_UPPER_FRAME_NAME
        If i + 1 <= Images.Count Then _
            PlaceImage Images(i + 1), Page, IMPORTER_LOWER_FRAME_NAME
    Next i
End Sub

Private Sub PlaceImage( _
                ByVal Image As Shape, _
                ByVal Page As Page, _
                ByVal FrameName As String _
            )
    This.Progress.Update
    Dim Placeholder As Shape: Set Placeholder = FindShape(Page, FrameName)
    Dim Layer As Layer: Set Layer = Placeholder.Layer
    Image.MoveToLayer Layer
    FitInside Image, Placeholder.BoundingBox
    Image.Bitmap.Resample _
        ResolutionX:=IMPORTER_RESOLUTION, ResolutionY:=IMPORTER_RESOLUTION
    Image.LeftX = Placeholder.LeftX
    Image.TopY = Placeholder.TopY
    Placeholder.Delete
End Sub

Private Sub SaveResult(ByVal Path As String)
    Dim File As FileSpec: Set File = FileSpec.New_
    File.Path = Path
    File.Name = ActiveDocument.FileName
    This.Progress.Caption = "Сохранение..."
    ActiveDocument.SaveAsCopy File
End Sub

Private Property Get FindShape(ByVal Page As Page, ByVal Name As String) As Shape
    GetShapeByName Page.Shapes.All, Name, FindShape
End Property

Private Sub ActivateImportLayer()
    FindShape(ActivePage, IMPORTER_UPPER_FRAME_NAME).Layer.Activate
End Sub

Private Sub GetValidFiles( _
                         ByVal Path As String, _
                         ByRef Files As Collection _
                     )
    Set Files = _
        FindFiles(This.Cfg!WorkPath, Split(IMPORTER_VALID_EXT_PATTERNS, ","), True)
End Sub

'===============================================================================
' # View and config

Private Function ShowImportPhotosView(ByRef Cfg As Dictionary) As BooleanResult
    Dim FileBinder As JsonFileBinder: Set FileBinder = BindConfig
    Set Cfg = FileBinder.GetOrMakeSubDictionary("ImportPhotos")
    Dim View As New ImportPhotosView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg, _
            View:=View, _
            ControlNames:=Pack("TemplateFile", "WorkPath") _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowImportPhotosView = View.IsOk
End Function

Private Function BindConfig() As JsonFileBinder
    Set BindConfig = JsonFileBinder.New_(APP_FILEBASENAME)
End Function

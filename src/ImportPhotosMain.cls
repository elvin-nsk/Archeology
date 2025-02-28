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
    
    Dim Doc As Document: Set Doc = Application.OpenDocument(This.Cfg!TemplateFile)
    If ValidateTemplate = Fail Then
        Warn "На шаблоне отсутствуют места размещения."
        Exit Sub
    End If
    
    #If DebugMode = 0 Then
    Optimization = True
    #End If
    
    Dim Subfolders As Folders: Set Subfolders = _
        FSO.GetFolder(This.Cfg!WorkPath).Subfolders
    Set This.Progress = ProgressBar.New_(ProgressBarNumeric, Subfolders.Count)
    This.Progress.Caption = "Обработка папок"
    Dim Folder As Folder
    For Each Folder In Subfolders
        ProcessFolder Doc, Folder
    Next Folder
    
    Doc.Close
End Sub

Public Sub ProcessFolder(ByVal Source As Document, ByVal Folder As Folder)
    Dim ValidFiles As Collection: GetValidFiles Folder, ValidFiles
    'Show ValidFiles.Count
    If ValidFiles.Count = 0 Then Exit Sub
    
    Source.Duplicate.Activate
    
    This.Progress.Update
    AddPages ValidFiles.Count
    ActiveDocument.Pages.First.Activate
    
    Dim Images As ShapeRange: Set Images = ImportImages(ValidFiles)
    ResizeImages Images
    PlaceImages Images
    SaveResult Folder.Name
    
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

Private Sub ResizeImages(ByVal Images As ShapeRange)
    Dim Size As Rect: Set Size = _
        FindShape(ActivePage, IMPORTER_UPPER_FRAME_NAME).BoundingBox
    Dim Image As Shape
    For Each Image In Images
        ResizeImage Image, Size
        Image.LeftX = ActivePage.RightX
        Image.TopY = ActivePage.TopY
    Next Image
End Sub

Private Sub PlaceImage( _
                ByVal Image As Shape, _
                ByVal Page As Page, _
                ByVal FrameName As String _
            )
    Dim Placeholder As Shape: Set Placeholder = FindShape(Page, FrameName)
    Dim Layer As Layer: Set Layer = Placeholder.Layer
    Image.MoveToLayer Layer
    ResizeImage Image, Placeholder.BoundingBox
    Image.Bitmap.Resample _
        ResolutionX:=This.Cfg!Resolution, ResolutionY:=This.Cfg!Resolution
    Image.LeftX = Placeholder.LeftX
    Image.TopY = Placeholder.TopY
    Placeholder.Delete
End Sub

Private Sub SaveResult(ByVal Suffix As String)
    Dim File As FileSpec: Set File = FileSpec.New_(This.Cfg!TemplateFile)
    File.Path = This.Cfg!WorkPath
    File.BaseName = File.BaseName & Suffix
    File.Ext = "cdr"
    
    'Show File.ToString
    ActiveDocument.SaveAs File
End Sub

Private Property Get FindShape(ByVal Page As Page, ByVal Name As String) As Shape
    GetShapeByName Page.Shapes.All, Name, FindShape
End Property

Private Sub ActivateImportLayer()
    FindShape(ActivePage, IMPORTER_UPPER_FRAME_NAME).Layer.Activate
End Sub

Public Sub ResizeImage( _
               ByVal Image As Shape, _
               ByVal TargetRect As Rect _
           )
    FitInside Image, TargetRect
End Sub

Private Sub GetValidFiles( _
                         ByVal Folder As Scripting.Folder, _
                         ByRef Files As Collection _
                     )
    Set Files = New Collection
    AppendFilesFromFolder _
        Folder, Files, Split(IMPORTER_VALID_EXT_PATTERNS, ","), True
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
            ControlNames:=Pack("TemplateFile", "WorkPath", "Resolution") _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowImportPhotosView = View.IsOk
End Function

Private Function BindConfig() As JsonFileBinder
    Set BindConfig = JsonFileBinder.New_(APP_FILEBASENAME)
End Function

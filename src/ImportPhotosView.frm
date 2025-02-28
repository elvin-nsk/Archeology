VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportPhotosView 
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   OleObjectBlob   =   "ImportPhotosView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportPhotosView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Public IsOk As Boolean
Public IsCancel As Boolean

Public WorkPathHandler As FolderBrowserHandler
Public ResolutionHandler As TextBoxHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = APP_DISPLAYNAME & " - ResizePhotos (v" & APP_VERSION & ")"
    btnOk.Default = True
        
    Set WorkPathHandler = _
        FolderBrowserHandler.New_(WorkPath, WorkPathBrowse)
    Set ResolutionHandler = _
        TextBoxHandler.New_(Resolution, TextBoxTypeLong, 1)
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    If Resolution = vbNullString Then Resolution = 300
End Sub

Private Sub TemplateFileBrowse_Click()
    Dim File As String
    File = _
        CorelScriptTools.GetFileBox( _
            Filter:="cdr|*.cdr", _
            Title:="Открыть шаблон", _
            Type:=0, _
            Extension:="cdr" _
        )
    If Not File = vbNullString Then TemplateFile = File
End Sub

Private Sub btnOk_Click()
    FormОК
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormОК()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers


'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub


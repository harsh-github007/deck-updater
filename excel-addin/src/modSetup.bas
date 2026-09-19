Attribute VB_Name = "modSetup"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync. Setup sheet layout and constants.
' The Setup sheet holds two small tables; their cell positions live only here.
' ---------------------------------------------------------------------------

Public Const APP_NAME As String = "SlideSync"
Public Const APP_VERSION As String = "0.1.0"

Public Const SHEET_HOME As String = "Home"
Public Const SHEET_SETUP As String = "Setup"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_HELP As String = "Instructions"

' Named ranges on the Setup sheet (created by BuildWorkbookUI / shipped in SlideSync.xlsx)
Public Const RNG_SOURCES As String = "tblSources"      ' Alias | Folder | File name
Public Const RNG_DECKS As String = "tblDecks"          ' Template folder | Template file | Save as
Public Const RNG_OPT_CLOSE As String = "optCloseSources"
Public Const RNG_OPT_CLEARLOG As String = "optClearLog"

' Column offsets inside tblSources
Public Const SRC_ALIAS As Long = 1
Public Const SRC_FOLDER As Long = 2
Public Const SRC_FILE As Long = 3

' Column offsets inside tblDecks
Public Const DECK_FOLDER As Long = 1
Public Const DECK_FILE As Long = 2
Public Const DECK_SAVEAS As Long = 3

' PowerPoint enum values (late bound, so we define what we use)
Public Const MSO_TRUE As Long = -1
Public Const MSO_GROUP As Long = 6
Public Const PP_PASTE_PNG As Long = 6
Public Const PP_PASTE_EMF As Long = 2
Public Const PP_WINDOW_NORMAL As Long = 1

Public Function SetupRange(ByVal nameOrAddress As String) As Range
    Set SetupRange = ThisWorkbook.Names(nameOrAddress).RefersToRange
End Function

Public Function OptionIsYes(ByVal rangeName As String) As Boolean
    On Error Resume Next
    OptionIsYes = (UCase$(Trim$(CStr(SetupRange(rangeName).Value))) = "YES")
End Function

' Join a folder and file name, tolerating a missing trailing separator.
Public Function JoinPath(ByVal folder As String, ByVal fileName As String) As String
    folder = Trim$(folder)
    fileName = Trim$(fileName)
    If Len(folder) = 0 Then
        JoinPath = fileName
    ElseIf Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        JoinPath = folder & fileName
    Else
        JoinPath = folder & "\" & fileName
    End If
End Function

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(path) > 0) And (Len(Dir$(path)) > 0)
End Function

Public Function StripExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then StripExtension = Left$(fileName, p - 1) Else StripExtension = fileName
End Function

Public Function FolderOf(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p = 0 Then p = InStrRev(fullPath, "/")
    If p > 0 Then FolderOf = Left$(fullPath, p) Else FolderOf = ""
End Function

Attribute VB_Name = "modMain"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync, entry point.  Assign RunSlideSync to the button on the Home sheet
' (or run it from Alt+F8).  Uses late binding so no PowerPoint reference is
' required in Tools > References.
' ---------------------------------------------------------------------------

Public Sub RunSlideSync()
    Dim pptApp As Object, decks As Range, r As Long
    Dim templatePath As String, savePath As String
    Dim deckCount As Long, screenState As Boolean
    Dim weStartedPpt As Boolean

    screenState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo Failed

    StartLog OptionIsYes(RNG_OPT_CLEARLOG)
    If LoadSourcesFromSetup() = 0 Then
        MsgBox "Add at least one source workbook on the Setup sheet.", vbExclamation, APP_NAME
        GoTo CleanUp
    End If

    Set decks = SetupRange(RNG_DECKS)
    Set pptApp = GetPowerPoint(weStartedPpt)

    For r = 1 To decks.Rows.Count
        If Len(Trim$(CStr(decks.Cells(r, DECK_FILE).Value))) > 0 Then
            templatePath = JoinPath(CStr(decks.Cells(r, DECK_FOLDER).Value), CStr(decks.Cells(r, DECK_FILE).Value))
            savePath = ResolveSavePath(templatePath, CStr(decks.Cells(r, DECK_SAVEAS).Value))
            ProcessDeck pptApp, templatePath, savePath
            deckCount = deckCount + 1
        End If
    Next r

    If deckCount = 0 Then
        MsgBox "Add at least one template deck on the Setup sheet.", vbExclamation, APP_NAME
    Else
        MsgBox deckCount & " deck(s) processed." & vbCrLf & _
               "OK: " & LogOkCount() & "   Errors: " & LogErrorCount() & vbCrLf & vbCrLf & _
               "See the Log sheet for details.", IIf(LogErrorCount() > 0, vbExclamation, vbInformation), APP_NAME
    End If

CleanUp:
    On Error Resume Next
    If OptionIsYes(RNG_OPT_CLOSE) Then CloseSourcesOpenedByUs
    If weStartedPpt And Not pptApp Is Nothing Then
        If pptApp.Presentations.Count = 0 Then pptApp.Quit
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = screenState
    Exit Sub

Failed:
    MsgBox "SlideSync stopped: " & Err.Description, vbCritical, APP_NAME
    Resume CleanUp
End Sub

' ---------------------------------------------------------------------------
Private Function GetPowerPoint(ByRef weStartedIt As Boolean) As Object
    On Error Resume Next
    Set GetPowerPoint = GetObject(, "PowerPoint.Application")
    If GetPowerPoint Is Nothing Then
        Set GetPowerPoint = CreateObject("PowerPoint.Application")
        weStartedIt = True
    End If
    On Error GoTo 0
    If GetPowerPoint Is Nothing Then
        Err.Raise vbObjectError + 300, APP_NAME, "PowerPoint could not be started."
    End If
    GetPowerPoint.Visible = MSO_TRUE
End Function

' "Save as" may be blank (-> <template>_updated.pptx), a bare file name
' (-> same folder as the template) or a full path.
Private Function ResolveSavePath(ByVal templatePath As String, ByVal saveAs As String) As String
    saveAs = Trim$(saveAs)
    If Len(saveAs) = 0 Then
        ResolveSavePath = FolderOf(templatePath) & StripExtension(Mid$(templatePath, Len(FolderOf(templatePath)) + 1)) & "_updated.pptx"
    ElseIf InStr(1, saveAs, "\") = 0 And InStr(1, saveAs, "/") = 0 Then
        ResolveSavePath = FolderOf(templatePath) & saveAs
    Else
        ResolveSavePath = saveAs
    End If
    If LCase$(Right$(ResolveSavePath, 5)) <> ".pptx" Then ResolveSavePath = ResolveSavePath & ".pptx"
End Function

' ---------------------------------------------------------------------------
Private Sub ProcessDeck(ByVal pptApp As Object, ByVal templatePath As String, ByVal savePath As String)
    Dim pres As Object, slide As Object, shp As Object
    Dim deckName As String, slideNo As Long

    deckName = Mid$(templatePath, Len(FolderOf(templatePath)) + 1)

    If Not FileExists(templatePath) Then
        LogEntry deckName, 0, "", "", templatePath, "deck", "ERROR", "Template not found"
        Exit Sub
    End If

    ' Open a copy so the template is never modified.
    Set pres = pptApp.Presentations.Open(fileName:=templatePath, ReadOnly:=MSO_TRUE, Untitled:=MSO_TRUE, WithWindow:=MSO_TRUE)

    slideNo = 0
    For Each slide In pres.Slides
        slideNo = slideNo + 1
        For Each shp In slide.Shapes
            ProcessShape deckName, slide, slideNo, shp
        Next shp
    Next slide

    pres.SaveAs savePath
    LogEntry deckName, 0, "", "", savePath, "deck", "OK", "saved"
    pres.Close
End Sub

Private Sub ProcessShape(ByVal deckName As String, ByVal slide As Object, ByVal slideNo As Long, ByVal shp As Object)
    Dim d As Directive, child As Object, r As Long, c As Long, msg As String

    If shp.Type = MSO_GROUP Then
        For Each child In shp.GroupItems
            ProcessShape deckName, slide, slideNo, child
        Next child
        Exit Sub
    End If

    d = ParseDirective(shp.Name)
    If Len(d.Kind) > 0 Then
        On Error Resume Next
        Select Case d.Kind
            Case "table":   msg = FillTable(shp, ResolveReference(d.Reference), d.Header)
            Case "picture": msg = ReplaceWithRangePicture(slide, shp, ResolveReference(d.Reference))
            Case "chart":   msg = UpdateChart(shp, ResolveReference(d.Reference), d.Orient)
        End Select
        If Err.Number <> 0 Then
            LogEntry deckName, slideNo, shp.Name, shp.Name, d.Reference, d.Kind, "ERROR", Err.Description
            Err.Clear
        Else
            LogEntry deckName, slideNo, shp.Name, shp.Name, d.Reference, d.Kind, "OK", msg
        End If
        On Error GoTo 0
        If d.Kind = "picture" Or d.Kind = "chart" Then Exit Sub   ' shape replaced / nothing textual
    End If

    ' Tokens in ordinary text
    If shp.HasTextFrame = MSO_TRUE Then
        If shp.TextFrame.HasText = MSO_TRUE Then
            ReplaceTokens shp.TextFrame.TextRange, deckName, slideNo, shp.Name
        End If
    End If

    ' Tokens inside table cells
    If shp.HasTable = MSO_TRUE Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                ReplaceTokens shp.Table.Cell(r, c).Shape.TextFrame.TextRange, deckName, slideNo, shp.Name
            Next c
        Next r
    End If
End Sub

' ---------------------------------------------------------------------------
Public Sub ShowAbout()
    MsgBox APP_NAME & " " & APP_VERSION & vbCrLf & vbCrLf & _
           "Refresh PowerPoint decks from Excel data." & vbCrLf & _
           "Tokens: {{Sheet!A1}}   Directives: table: / picture: / chart:" & vbCrLf & vbCrLf & _
           "MIT licensed, see the project README.", vbInformation, APP_NAME
End Sub

Public Sub GoToSheet(ByVal sheetName As String)
    ThisWorkbook.Worksheets(sheetName).Activate
End Sub
Public Sub GoHome():    GoToSheet SHEET_HOME:  End Sub
Public Sub GoSetup():   GoToSheet SHEET_SETUP: End Sub
Public Sub GoLog():     GoToSheet SHEET_LOG:   End Sub
Public Sub GoHelp():    GoToSheet SHEET_HELP:  End Sub

Attribute VB_Name = "modUI"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync, run InstallButtons once after importing the modules into
' SlideSync.xlsx (then save as .xlsm).  It draws the buttons on the Home sheet
' and wires them to the macros in modMain.
' ---------------------------------------------------------------------------

Public Sub InstallButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_HOME)

    RemoveButtons ws
    AddButton ws, "btnRun", "Update decks", "RunSlideSync", ws.Range("C12"), RGB(31, 78, 121)
    AddButton ws, "btnSetup", "Open Setup", "GoSetup", ws.Range("E12"), RGB(89, 89, 89)
    AddButton ws, "btnLog", "View Log", "GoLog", ws.Range("G12"), RGB(89, 89, 89)
    AddButton ws, "btnHelp", "Instructions", "GoHelp", ws.Range("I12"), RGB(89, 89, 89)
    AddButton ws, "btnAbout", "About", "ShowAbout", ws.Range("K12"), RGB(89, 89, 89)

    AddButton ThisWorkbook.Worksheets(SHEET_SETUP), "btnRunFromSetup", "Update decks", "RunSlideSync", _
              ThisWorkbook.Worksheets(SHEET_SETUP).Range("H2"), RGB(31, 78, 121)
    AddButton ThisWorkbook.Worksheets(SHEET_SETUP), "btnHomeFromSetup", "Home", "GoHome", _
              ThisWorkbook.Worksheets(SHEET_SETUP).Range("J2"), RGB(89, 89, 89)
    AddButton ThisWorkbook.Worksheets(SHEET_LOG), "btnHomeFromLog", "Home", "GoHome", _
              ThisWorkbook.Worksheets(SHEET_LOG).Range("H2"), RGB(89, 89, 89)
    AddButton ThisWorkbook.Worksheets(SHEET_HELP), "btnHomeFromHelp", "Home", "GoHome", _
              ThisWorkbook.Worksheets(SHEET_HELP).Range("H2"), RGB(89, 89, 89)

    ws.Activate
    MsgBox "Buttons installed. Save the workbook as .xlsm to keep the macros.", vbInformation, APP_NAME
End Sub

Private Sub RemoveButtons(ByVal ws As Worksheet)
    Dim shp As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, 3) = "btn" Then shp.Delete
    Next i
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal name As String, ByVal caption As String, _
                      ByVal macro As String, ByVal anchor As Range, ByVal fillColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 110, 32)
    With shp
        .Name = name
        .OnAction = macro
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Adjustments(1) = 0.2
        With .TextFrame2
            .TextRange.Text = caption
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.Font.Name = "Arial"
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 4: .MarginRight = 4
        End With
    End With
End Sub

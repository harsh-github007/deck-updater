Attribute VB_Name = "modDirectives"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync, shape directives stored in the shape name (SPEC.md §4):
'   table:Sheet!A1:D10[;header=0]
'   picture:Sheet!A1:D10
'   chart:Sheet!A1:D10[;orient=rows]
' ---------------------------------------------------------------------------

Public Type Directive
    Kind As String        ' table | picture | chart | "" when not a directive
    Reference As String
    Header As Boolean
    Orient As String      ' cols | rows
End Type

Public Function ParseDirective(ByVal shapeName As String) As Directive
    Dim d As Directive, p As Long, kind As String, rest As String
    Dim parts() As String, i As Long, kv() As String

    d.Header = True
    d.Orient = "cols"
    p = InStr(1, shapeName, ":")
    If p = 0 Then
        ParseDirective = d
        Exit Function
    End If
    kind = LCase$(Trim$(Left$(shapeName, p - 1)))
    If kind <> "table" And kind <> "picture" And kind <> "chart" Then
        ParseDirective = d
        Exit Function
    End If

    rest = Mid$(shapeName, p + 1)
    parts = Split(rest, ";")
    d.Kind = kind
    d.Reference = Trim$(parts(0))
    For i = 1 To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            Select Case LCase$(Trim$(kv(0)))
                Case "header": d.Header = Not (Trim$(kv(1)) = "0" Or LCase$(Trim$(kv(1))) = "false" Or LCase$(Trim$(kv(1))) = "no")
                Case "orient": d.Orient = LCase$(Trim$(kv(1)))
            End Select
        End If
    Next i
    ParseDirective = d
End Function

' ---------------------------------------------------------------------------
' table:  fill a native PowerPoint table, resizing rows/columns to the range
' ---------------------------------------------------------------------------
Public Function FillTable(ByVal shp As Object, ByVal rng As Range, ByVal hasHeader As Boolean) As String
    Dim tbl As Object, needRows As Long, needCols As Long, r As Long, c As Long
    Dim origWidth As Single

    If shp.HasTable <> MSO_TRUE Then Err.Raise vbObjectError + 200, APP_NAME, "Shape is not a table"
    Set tbl = shp.Table
    needRows = rng.Rows.Count
    needCols = rng.Columns.Count
    origWidth = shp.Width

    ' Columns: add/remove at the end so template styling carries over.
    Do While tbl.Columns.Count < needCols
        tbl.Columns.Add
    Loop
    Do While tbl.Columns.Count > needCols
        tbl.Columns(tbl.Columns.Count).Delete
    Loop
    For c = 1 To needCols
        tbl.Columns(c).Width = origWidth / needCols
    Next c

    ' Rows, keep at least the header + one body row while shrinking.
    Do While tbl.Rows.Count < needRows
        tbl.Rows.Add
    Loop
    Do While tbl.Rows.Count > needRows
        tbl.Rows(tbl.Rows.Count).Delete
    Loop

    For r = 1 To needRows
        For c = 1 To needCols
            tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = FormatCell(rng.Cells(r, c))
        Next c
    Next r
    FillTable = needRows & "x" & needCols
End Function

' ---------------------------------------------------------------------------
' picture:  copy the range as a picture and swap it in at the same position
' ---------------------------------------------------------------------------
Public Function ReplaceWithRangePicture(ByVal slide As Object, ByVal shp As Object, ByVal rng As Range) As String
    Dim newShape As Object, l As Single, t As Single, w As Single, h As Single
    Dim oldName As String, z As Long

    l = shp.Left: t = shp.Top: w = shp.Width: h = shp.Height
    oldName = shp.Name
    z = shp.ZOrderPosition

    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Set newShape = slide.Shapes.PasteSpecial(PP_PASTE_EMF)(1)

    ' Scale to fit inside the old box, preserving aspect ratio.
    newShape.LockAspectRatio = MSO_TRUE
    If newShape.Width / newShape.Height > w / h Then
        newShape.Width = w
    Else
        newShape.Height = h
    End If
    newShape.Left = l
    newShape.Top = t

    shp.Delete
    newShape.Name = oldName
    On Error Resume Next
    Do While newShape.ZOrderPosition > z
        newShape.ZOrder 3        ' msoSendBackward
    Loop
    On Error GoTo 0
    Application.CutCopyMode = False
    ReplaceWithRangePicture = rng.Rows.Count & "x" & rng.Columns.Count & " rendered"
End Function

' ---------------------------------------------------------------------------
' chart:  write the range into the chart's embedded data sheet and re-point it
' ---------------------------------------------------------------------------
Public Function UpdateChart(ByVal shp As Object, ByVal rng As Range, ByVal orient As String) As String
    Dim cht As Object, dataWb As Object, dataWs As Object
    Dim values As Variant, r As Long, c As Long, addr As String

    If shp.HasChart <> MSO_TRUE Then Err.Raise vbObjectError + 201, APP_NAME, "Shape is not a chart"
    If rng.Rows.Count < 2 Or rng.Columns.Count < 2 Then
        Err.Raise vbObjectError + 202, APP_NAME, "Chart range needs a header row/column plus data"
    End If

    Set cht = shp.Chart
    cht.ChartData.Activate
    Set dataWb = cht.ChartData.Workbook
    Set dataWs = dataWb.Worksheets(1)

    dataWs.Cells.Clear
    values = rng.Value
    dataWs.Range(dataWs.Cells(1, 1), dataWs.Cells(rng.Rows.Count, rng.Columns.Count)).Value = values
    ' carry number formats so axis/data labels look like the source
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            dataWs.Cells(r, c).NumberFormat = rng.Cells(r, c).NumberFormat
        Next c
    Next r

    addr = "'" & dataWs.Name & "'!" & dataWs.Range(dataWs.Cells(1, 1), dataWs.Cells(rng.Rows.Count, rng.Columns.Count)).Address(True, True)
    If orient = "rows" Then
        cht.SetSourceData Source:=addr, PlotBy:=1   ' xlRows
    Else
        cht.SetSourceData Source:=addr, PlotBy:=2   ' xlColumns
    End If
    dataWb.Close
    UpdateChart = (IIf(orient = "rows", rng.Rows.Count, rng.Columns.Count) - 1) & " series"
End Function

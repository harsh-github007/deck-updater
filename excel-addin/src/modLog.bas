Attribute VB_Name = "modLog"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync, run log on the Log sheet (SPEC.md §6).
' Columns: Deck | Slide | Shape | Directive/Token | Reference | Action | Status | Message
' ---------------------------------------------------------------------------

Private Const LOG_HEADER_ROW As Long = 5
Private Const LOG_FIRST_COL As Long = 2   ' column B

Private mNextRow As Long
Private mErrors As Long
Private mOk As Long

Public Sub StartLog(ByVal clearExisting As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    If clearExisting Then
        ws.Range(ws.Cells(LOG_HEADER_ROW + 1, LOG_FIRST_COL), ws.Cells(ws.Rows.Count, LOG_FIRST_COL + 7)).ClearContents
        mNextRow = LOG_HEADER_ROW + 1
    Else
        mNextRow = ws.Cells(ws.Rows.Count, LOG_FIRST_COL).End(xlUp).Row + 1
        If mNextRow <= LOG_HEADER_ROW Then mNextRow = LOG_HEADER_ROW + 1
    End If
    mErrors = 0
    mOk = 0
End Sub

Public Sub LogEntry(ByVal deckName As String, ByVal slideNo As Long, ByVal shapeName As String, _
                    ByVal directive As String, ByVal reference As String, ByVal action As String, _
                    ByVal status As String, ByVal message As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    With ws
        .Cells(mNextRow, LOG_FIRST_COL + 0).Value = deckName
        .Cells(mNextRow, LOG_FIRST_COL + 1).Value = slideNo
        .Cells(mNextRow, LOG_FIRST_COL + 2).Value = shapeName
        .Cells(mNextRow, LOG_FIRST_COL + 3).Value = "'" & directive
        .Cells(mNextRow, LOG_FIRST_COL + 4).Value = "'" & reference
        .Cells(mNextRow, LOG_FIRST_COL + 5).Value = action
        .Cells(mNextRow, LOG_FIRST_COL + 6).Value = status
        .Cells(mNextRow, LOG_FIRST_COL + 7).Value = message
        If status = "ERROR" Then
            .Cells(mNextRow, LOG_FIRST_COL + 6).Font.Color = RGB(192, 0, 0)
            mErrors = mErrors + 1
        Else
            .Cells(mNextRow, LOG_FIRST_COL + 6).Font.Color = RGB(0, 112, 60)
            mOk = mOk + 1
        End If
    End With
    mNextRow = mNextRow + 1
End Sub

Public Function LogErrorCount() As Long
    LogErrorCount = mErrors
End Function

Public Function LogOkCount() As Long
    LogOkCount = mOk
End Function

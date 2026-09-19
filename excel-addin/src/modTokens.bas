Attribute VB_Name = "modTokens"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync, {{reference|format}} tokens inside PowerPoint text.
'
' PowerPoint's TextRange.Characters() spans runs transparently, so a token that
' was split across runs by editing is still replaced as one unit and the new
' text inherits the formatting of the token's first character.
' ---------------------------------------------------------------------------

Private Const TOKEN_OPEN As String = "{{"
Private Const TOKEN_CLOSE As String = "}}"

' Replace all tokens in a PowerPoint TextRange (late bound).
' Returns the number of tokens replaced; logs each via modLog.
Public Function ReplaceTokens(ByVal textRange As Object, ByVal deckName As String, _
                              ByVal slideNo As Long, ByVal shapeName As String) As Long
    Dim fullText As String, startPos As Long, endPos As Long, searchFrom As Long
    Dim token As String, inner As String, refText As String, fmtText As String
    Dim replacement As String, ok As Boolean, msg As String

    fullText = textRange.Text
    If InStr(1, fullText, TOKEN_OPEN) = 0 Then Exit Function

    searchFrom = 1
    Do
        startPos = InStr(searchFrom, fullText, TOKEN_OPEN)
        If startPos = 0 Then Exit Do
        endPos = InStr(startPos + 2, fullText, TOKEN_CLOSE)
        If endPos = 0 Then Exit Do

        token = Mid$(fullText, startPos, endPos - startPos + 2)
        inner = Mid$(token, 3, Len(token) - 4)
        SplitToken inner, refText, fmtText

        ok = True
        On Error Resume Next
        replacement = ValueForToken(refText, fmtText)
        If Err.Number <> 0 Then
            ok = False
            msg = Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        If ok Then
            ' Characters() positions are 1-based like InStr, so this maps directly.
            textRange.Characters(startPos, Len(token)).Text = replacement
            LogEntry deckName, slideNo, shapeName, token, refText, "text", "OK", replacement
            ReplaceTokens = ReplaceTokens + 1
            fullText = textRange.Text
            searchFrom = startPos + Len(replacement)
        Else
            LogEntry deckName, slideNo, shapeName, token, refText, "text", "ERROR", msg
            searchFrom = endPos + 2
        End If
    Loop
End Function

Private Sub SplitToken(ByVal inner As String, ByRef refText As String, ByRef fmtText As String)
    Dim p As Long
    p = InStr(1, inner, "|")
    If p > 0 Then
        refText = Trim$(Left$(inner, p - 1))
        fmtText = Trim$(Mid$(inner, p + 1))
    Else
        refText = Trim$(inner)
        fmtText = ""
    End If
End Sub

Private Function ValueForToken(ByVal refText As String, ByVal fmtText As String) As String
    Dim rng As Range
    Set rng = ResolveReference(refText)
    ValueForToken = FormatCell(rng.Cells(1, 1), fmtText)
End Function

' Render a cell as text. With no explicit format, Excel's own displayed text is used.
Public Function FormatCell(ByVal cell As Range, Optional ByVal fmtText As String = "") As String
    Dim v As Variant
    v = cell.Value
    If IsEmpty(v) Or IsNull(v) Then
        FormatCell = ""
        Exit Function
    End If
    If IsError(v) Then
        FormatCell = CStr(cell.Text)
        Exit Function
    End If

    If Len(fmtText) = 0 Then
        FormatCell = CStr(cell.Text)
    ElseIf LCase$(Left$(fmtText, 5)) = "date:" Then
        If IsDate(v) Then
            FormatCell = Format$(CDate(v), Mid$(fmtText, 6))
        Else
            FormatCell = CStr(v)
        End If
    ElseIf IsNumeric(v) Then
        FormatCell = Format$(v, fmtText)
    ElseIf IsDate(v) Then
        FormatCell = Format$(CDate(v), fmtText)
    Else
        FormatCell = CStr(v)
    End If
End Function

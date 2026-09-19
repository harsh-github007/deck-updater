Attribute VB_Name = "modSources"
Option Explicit

' ---------------------------------------------------------------------------
' SlideSync: source workbooks and reference resolution.
'
' A reference (see SPEC.md §2) is one of:
'   Sheet1!B3           'Q3 Data'!A1:D12      TotalRevenue
'   budget:Sheet1!B3    budget:TotalRevenue
' The optional "alias:" prefix picks a source workbook; otherwise the first
' source listed on the Setup sheet is used.
' ---------------------------------------------------------------------------

Private mAliases As Collection      ' ordered list of alias strings
Private mPaths As Object            ' Scripting.Dictionary  alias -> full path
Private mBooks As Object            ' Scripting.Dictionary  alias -> Workbook
Private mOpenedByUs As Object       ' Scripting.Dictionary  alias -> True if we opened it

Public Sub ClearSources()
    Set mAliases = New Collection
    Set mPaths = CreateObject("Scripting.Dictionary")
    Set mBooks = CreateObject("Scripting.Dictionary")
    Set mOpenedByUs = CreateObject("Scripting.Dictionary")
    mPaths.CompareMode = vbTextCompare
    mBooks.CompareMode = vbTextCompare
    mOpenedByUs.CompareMode = vbTextCompare
End Sub

' Read the Sources table from the Setup sheet. Returns the number of sources.
Public Function LoadSourcesFromSetup() As Long
    Dim tbl As Range, r As Long
    Dim alias As String, fullPath As String

    ClearSources
    Set tbl = SetupRange(RNG_SOURCES)
    For r = 1 To tbl.Rows.Count
        fullPath = JoinPath(CStr(tbl.Cells(r, SRC_FOLDER).Value), CStr(tbl.Cells(r, SRC_FILE).Value))
        If Len(Trim$(CStr(tbl.Cells(r, SRC_FILE).Value))) > 0 Then
            alias = Trim$(CStr(tbl.Cells(r, SRC_ALIAS).Value))
            If Len(alias) = 0 Then alias = "src" & r
            If mPaths.Exists(alias) Then
                Err.Raise vbObjectError + 100, APP_NAME, "Duplicate source alias '" & alias & "' on Setup sheet."
            End If
            mAliases.Add alias
            mPaths.Add alias, fullPath
        End If
    Next r
    LoadSourcesFromSetup = mAliases.Count
End Function

Public Function DefaultAlias() As String
    If mAliases Is Nothing Or mAliases.Count = 0 Then
        Err.Raise vbObjectError + 101, APP_NAME, "No source workbooks listed on the Setup sheet."
    End If
    DefaultAlias = mAliases(1)
End Function

' Returns the workbook for an alias, opening it read-only if needed.
'
' Two things this is careful about:
'   * An already-open workbook is only reused when its FullName matches the
'     configured path. Matching on Name alone means a stale copy of
'     "sales.xlsx" left open from another folder silently supplies last
'     period's numbers to every deck.
'   * A source is data, not a program. AutomationSecurity is forced to disable
'     macros around the Open call so a macro-enabled source cannot run code
'     just because the tool opened it.
Public Function SourceBook(ByVal alias As String) As Workbook
    Dim wb As Workbook, fullPath As String, shortName As String
    Dim prevSecurity As MsoAutomationSecurity, securityChanged As Boolean

    If Len(alias) = 0 Then alias = DefaultAlias()
    If Not mPaths.Exists(alias) Then
        Err.Raise vbObjectError + 102, APP_NAME, "Unknown source alias '" & alias & "'."
    End If
    If mBooks.Exists(alias) Then
        Set SourceBook = mBooks(alias)
        Exit Function
    End If

    fullPath = mPaths(alias)
    shortName = Mid$(fullPath, Len(FolderOf(fullPath)) + 1)

    ' Reuse an already-open copy only when it is genuinely the same file.
    For Each wb In Application.Workbooks
        If StrComp(NormalizePath(wb.FullName), NormalizePath(fullPath), vbTextCompare) = 0 Then
            Set SourceBook = wb
            mBooks.Add alias, wb
            mOpenedByUs.Add alias, False
            Exit Function
        End If
    Next wb

    ' Same file name, different folder: Excel cannot hold both open at once,
    ' so say so rather than quietly using the wrong one.
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, shortName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 104, APP_NAME, _
                "A different '" & shortName & "' is already open (" & wb.FullName & ")." & vbCrLf & _
                "Close it, then run again so the configured file can be used: " & fullPath
        End If
    Next wb

    If Not FileExists(fullPath) Then
        Err.Raise vbObjectError + 103, APP_NAME, "Source file not found: " & fullPath
    End If

    On Error GoTo Cleanup
    prevSecurity = Application.AutomationSecurity
    securityChanged = True
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wb = Application.Workbooks.Open(fileName:=fullPath, UpdateLinks:=0, ReadOnly:=True)
    Application.AutomationSecurity = prevSecurity
    securityChanged = False

    mBooks.Add alias, wb
    mOpenedByUs.Add alias, True
    Set SourceBook = wb
    Exit Function

Cleanup:
    ' Restore the user's setting whatever happened, then re-raise.
    If securityChanged Then Application.AutomationSecurity = prevSecurity
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Trims a path for comparison. Excel reports UNC and mapped-drive paths as
' given, so this only normalises what can be normalised safely.
Private Function NormalizePath(ByVal path As String) As String
    NormalizePath = Trim$(path)
    Do While Right$(NormalizePath, 1) = "\"
        NormalizePath = Left$(NormalizePath, Len(NormalizePath) - 1)
    Loop
End Function

Public Sub CloseSourcesOpenedByUs()
    Dim key As Variant
    If mBooks Is Nothing Then Exit Sub
    For Each key In mBooks.Keys
        If mOpenedByUs(key) Then
            On Error Resume Next
            mBooks(key).Close SaveChanges:=False
            On Error GoTo 0
        End If
    Next key
    ClearSources
End Sub

' ---------------------------------------------------------------------------
' Reference parsing
' ---------------------------------------------------------------------------

' Resolve a reference string to a Range. Raises an error when it cannot.
Public Function ResolveReference(ByVal refText As String) As Range
    Dim alias As String, body As String, sheetName As String, address As String
    Dim wb As Workbook, ws As Worksheet, p As Long

    body = Trim$(refText)

    ' alias: prefix, a colon before any "!" and not inside quotes
    p = InStr(1, body, ":")
    If p > 0 Then
        If InStr(1, body, "!") = 0 Or p < InStr(1, body, "!") Then
            If Left$(body, 1) <> "'" And Not IsRangeAddress(body) Then
                alias = Trim$(Left$(body, p - 1))
                body = Trim$(Mid$(body, p + 1))
            End If
        End If
    End If

    Set wb = SourceBook(alias)

    p = InStr(1, body, "!")
    If p = 0 Then
        ' defined name
        On Error Resume Next
        Set ResolveReference = wb.Names(body).RefersToRange
        On Error GoTo 0
        If ResolveReference Is Nothing Then
            Err.Raise vbObjectError + 104, APP_NAME, "Defined name '" & body & "' not found in " & wb.Name
        End If
        Exit Function
    End If

    sheetName = Trim$(Left$(body, p - 1))
    address = Trim$(Mid$(body, p + 1))
    If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then
        sheetName = Mid$(sheetName, 2, Len(sheetName) - 2)
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 105, APP_NAME, "Sheet '" & sheetName & "' not found in " & wb.Name
    End If

    On Error Resume Next
    Set ResolveReference = ws.Range(address)
    On Error GoTo 0
    If ResolveReference Is Nothing Then
        Err.Raise vbObjectError + 106, APP_NAME, "Invalid address '" & address & "' on sheet '" & sheetName & "'"
    End If
End Function

' True for things like "A1:B2" so a bare range with a colon is not read as an alias.
Private Function IsRangeAddress(ByVal text As String) As Boolean
    Dim i As Long, ch As String, sawLetter As Boolean, sawDigit As Boolean
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "[A-Za-z]" Then
            sawLetter = True
        ElseIf ch Like "#" Then
            sawDigit = True
        ElseIf ch <> ":" And ch <> "$" Then
            Exit Function
        End If
    Next i
    IsRangeAddress = sawLetter And sawDigit
End Function

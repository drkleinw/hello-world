Option Explicit

' =============================================================================
' ImportCSV_Events_To_NewTab_AttendeesOneColumn
' =============================================================================
' Main routine: imports a CSV file of calendar events from the user's Desktop,
' splits each attendee (required and optional) onto its own row, and produces
' three output sheets:
'   1. "Events (From CSV)"  - all rows, one attendee per row
'   2. "NoDuplicates"       - same data with duplicate attendees removed
'   3. "only External"      - only rows where the attendee contains "@"
' =============================================================================
Sub ImportCSV_Events_To_NewTab_AttendeesOneColumn()
    Dim csvPath As String
    Dim srcWB As Workbook, srcWS As Worksheet
    Dim outWS As Worksheet, dedupWS As Worksheet, extWS As Worksheet
    Dim hdrRow As Long, lastRow As Long
    Dim cSubj As Long, cStart As Long, cStartDate As Long, cStartTime As Long
    Dim cReq As Long, cOpt As Long
    Dim r As Long, outR As Long
    Dim subj As String, startVal As Variant, dt As Date, d As Variant, t As Variant
    Dim reqRaw As String, optRaw As String
    Dim reqArr As Variant, optArr As Variant
    Dim i As Long

    '----- Pick CSV from Desktop -----
    csvPath = PickCSVFromDesktop()

    ' Exit if user cancelled the file dialog
    If Len(csvPath) = 0 Then Exit Sub

    ' Suppress screen flicker, events, alerts, and auto-calc while processing
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Open the selected CSV workbook (read-only source)
    Set srcWB = Workbooks.Open(Filename:=csvPath)
    Set srcWS = srcWB.Sheets(1)

    ' Detect which row contains column headers (scan first 10 rows)
    hdrRow = DetectHeaderRow(srcWS, 10)
    If hdrRow = 0 Then
        MsgBox "Could not find a header row in the CSV.", vbExclamation
        GoTo Cleanup
    End If

    ' Find the last row with data in column A
    lastRow = srcWS.Cells(srcWS.Rows.Count, 1).End(xlUp).Row

    ' Map CSV columns to known header names (case-insensitive matching)
    cSubj = FindColByNames(srcWS, hdrRow, Array("Subject", "Title", "Event", "Event Subject"))
    cStart = FindColByNames(srcWS, hdrRow, Array("Start", "StartDateTime", "Start Date/Time", "Begin"))
    cStartDate = FindColByNames(srcWS, hdrRow, Array("Start Date", "StartDate"))
    cStartTime = FindColByNames(srcWS, hdrRow, Array("Start Time", "StartTime"))
    cReq = FindColByNames(srcWS, hdrRow, Array("Required Attendees", "RequiredAttendees", "Required"))
    cOpt = FindColByNames(srcWS, hdrRow, Array("Optional Attendees", "OptionalAttendees", "Optional"))

    ' Subject column is mandatory
    If cSubj = 0 Then
        MsgBox "Couldn't find a Subject/Title column.", vbExclamation
        GoTo Cleanup
    End If

    ' Delete any previous output sheet with the same name, then create a fresh one
    On Error Resume Next
    ThisWorkbook.Worksheets("Events (From CSV)").Delete
    On Error GoTo 0

    Set outWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    outWS.Name = "Events (From CSV)"

    ' Write column headers: A=Subject, B=Start Date, C=Start Time, D=Attendee
    outWS.Range("A1:D1").Value = Array("Subject", "Start Date", "Start Time", "Attendee")
    outR = 2  ' first data row

    ' ---- Process each data row in the CSV ----
    For r = hdrRow + 1 To lastRow
        subj = NzStr(srcWS.Cells(r, cSubj).Value)

        ' Parse date and time -- supports both combined and separate columns
        d = Empty: t = Empty
        If cStartDate > 0 Then d = srcWS.Cells(r, cStartDate).Value
        If cStartTime > 0 Then t = srcWS.Cells(r, cStartTime).Value

        ' Fall back to the combined Start column if separate date/time are empty
        If IsEmpty(d) And IsEmpty(t) Then
            If cStart > 0 Then
                startVal = srcWS.Cells(r, cStart).Value
                If TryParseDateTime(startVal, dt) Then
                    d = Int(dt)          ' date portion
                    t = dt - Int(dt)     ' time portion
                Else
                    SplitStartFallback CStr(startVal), d, t
                End If
            End If
        End If

        ' Split attendee strings strictly by semicolon
        reqRaw = IIf(cReq > 0, NzStr(srcWS.Cells(r, cReq).Value), "")
        optRaw = IIf(cOpt > 0, NzStr(srcWS.Cells(r, cOpt).Value), "")

        reqArr = SplitBySemicolon(reqRaw)
        optArr = SplitBySemicolon(optRaw)

        ' One row per REQUIRED attendee (col D)
        If IsArrayAllocated(reqArr) Then
            For i = LBound(reqArr) To UBound(reqArr)
                WriteEventRow_OneAttendee outWS, outR, subj, d, t, reqArr(i)
                outR = outR + 1
            Next i
        End If

        ' One row per OPTIONAL attendee (col D)
        If IsArrayAllocated(optArr) Then
            For i = LBound(optArr) To UBound(optArr)
                WriteEventRow_OneAttendee outWS, outR, subj, d, t, optArr(i)
                outR = outR + 1
            Next i
        End If

        ' If no attendees at all, still output one row with a blank attendee
        If Not IsArrayAllocated(reqArr) And Not IsArrayAllocated(optArr) Then
            WriteEventRow_OneAttendee outWS, outR, subj, d, t, ""
            outR = outR + 1
        End If
    Next r

    ' ---- Format the output sheet ----
    With outWS
        .Columns("A:D").EntireColumn.AutoFit
        If outR > 2 Then
            .Range("B2:B" & outR - 1).NumberFormat = "yyyy-mm-dd"
            .Range("C2:C" & outR - 1).NumberFormat = "hh:mm"
        End If
        .Range("A1:D1").Font.Bold = True
        .Range("D:D").WrapText = True
    End With

    ' ===== Create "NoDuplicates" sheet: copy of Events, deduped on Col D =====
    On Error Resume Next
    ThisWorkbook.Worksheets("NoDuplicates").Delete
    On Error GoTo 0

    ' Copy the entire output sheet, then remove duplicate attendee rows
    outWS.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set dedupWS = ActiveSheet
    dedupWS.Name = "NoDuplicates"

    Dim lastDedupRow As Long
    lastDedupRow = dedupWS.Cells(dedupWS.Rows.Count, "A").End(xlUp).Row
    If lastDedupRow >= 2 Then
        dedupWS.Range("A1:D" & lastDedupRow).RemoveDuplicates Columns:=4, Header:=xlYes
        dedupWS.Columns("A:D").EntireColumn.AutoFit
    End If

    ' ===== Create "only External" sheet: keep only attendees containing "@" =====
    On Error Resume Next
    ThisWorkbook.Worksheets("only External").Delete
    On Error GoTo 0

    ' Start from a copy of NoDuplicates
    dedupWS.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set extWS = ActiveSheet
    extWS.Name = "only External"

    ' Filter to rows where Attendee (Col D) contains "@", then copy visible
    ' rows to a temp sheet so we can discard internal (non-email) entries
    With extWS
        Dim lastExtRow As Long
        lastExtRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastExtRow >= 2 Then
            If .AutoFilterMode Then .AutoFilterMode = False
            .Range("A1:D" & lastExtRow).AutoFilter Field:=4, Criteria1:="=*@*"

            ' Copy visible (filtered) rows to a temporary sheet
            On Error Resume Next
            .Range("A2:D" & lastExtRow).SpecialCells(xlCellTypeVisible).Copy
            Dim tmpWS As Worksheet
            Set tmpWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            tmpWS.Name = "___tmp_onlyext"
            tmpWS.Range("A1").Resize(1, 4).Value = Array("Subject", "Start Date", "Start Time", "Attendee")
            tmpWS.Range("A2").PasteSpecial xlPasteValues
            Application.CutCopyMode = False

            ' Replace the filtered sheet with the clean temp sheet
            .Delete
            Set extWS = tmpWS
            extWS.Name = "only External"

            ' Format the external-only sheet
            With extWS
                If .AutoFilterMode Then .AutoFilterMode = False
                .Columns("A:D").AutoFit
                .Range("B2:B" & .Cells(.Rows.Count, "A").End(xlUp).Row).NumberFormat = "yyyy-mm-dd"
                .Range("C2:C" & .Cells(.Rows.Count, "A").End(xlUp).Row).NumberFormat = "hh:mm"
                .Range("A1:D1").Font.Bold = True
            End With
        End If
    End With

Cleanup:
    ' Restore application state regardless of success or failure
    On Error Resume Next
    If Not srcWB Is Nothing Then srcWB.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

'==================== Helpers ====================

' Prompt the user to pick a CSV file from their Desktop folder
Private Function PickCSVFromDesktop() As String
    Dim fd As FileDialog, initPath As String
    initPath = Environ$("USERPROFILE") & "\Desktop\"
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the CSV with calendar events"
        .InitialFileName = initPath
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then PickCSVFromDesktop = .SelectedItems(1)
    End With
End Function

' Scan the first maxScan rows to find one that looks like a header row
Private Function DetectHeaderRow(ws As Worksheet, Optional maxScan As Long = 10) As Long
    Dim r As Long
    For r = 1 To Application.Max(1, maxScan)
        If RowLooksLikeHeader(ws, r) Then DetectHeaderRow = r: Exit Function
    Next r
End Function

' A row "looks like a header" if it has at least 3 non-empty cells
Private Function RowLooksLikeHeader(ws As Worksheet, r As Long) As Boolean
    Dim c As Long, nonEmpty As Long
    For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then nonEmpty = nonEmpty + 1
    Next c
    RowLooksLikeHeader = (nonEmpty >= 3)
End Function

' Search a header row for a column matching any of the given names (case-insensitive)
' Returns the column number, or 0 if not found
Private Function FindColByNames(ws As Worksheet, hdrRow As Long, names As Variant) As Long
    Dim lastC As Long, c As Long, i As Long, v As String
    lastC = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        v = LCase$(Trim$(CStr(ws.Cells(hdrRow, c).Value)))
        If Len(v) > 0 Then
            For i = LBound(names) To UBound(names)
                If v = LCase$(names(i)) Then FindColByNames = c: Exit Function
            Next i
        End If
    Next c
End Function

' Attempt to parse a variant value as a Date, handling ISO 8601 "T" and "Z" formats
' Returns True on success and sets outDT to the parsed date
Private Function TryParseDateTime(ByVal v As Variant, ByRef outDT As Date) As Boolean
    On Error GoTo Fail
    If IsDate(v) Then outDT = CDate(v): TryParseDateTime = True: Exit Function
    Dim s As String: s = CStr(v)
    s = Replace$(s, "T", " ")
    s = Replace$(s, "Z", "")
    If IsDate(s) Then outDT = CDate(s): TryParseDateTime = True: Exit Function
Fail:
End Function

' Fallback parser for a combined start string: split on space into date and time parts
Private Sub SplitStartFallback(ByVal s As String, ByRef d As Variant, ByRef t As Variant)
    Dim tmp As String: tmp = Trim$(s)
    ' Normalise ISO separators
    If InStr(1, tmp, "T") > 0 Then tmp = Replace$(tmp, "T", " ")
    If InStr(1, tmp, "Z") > 0 Then tmp = Replace$(tmp, "Z", "")
    Dim parts() As String
    parts = Split(tmp, " ")
    If UBound(parts) >= 1 Then
        d = parts(0): t = parts(1)
    Else
        d = "": t = ""
    End If
End Sub

' Split a raw string by semicolons, trimming each token and ignoring empty entries.
' Returns a String array, or an unallocated Variant if the input is empty.
Private Function SplitBySemicolon(ByVal raw As String) As Variant
    Dim s As String: s = Trim$(raw)
    Dim tokens() As String, cleaned() As String
    Dim i As Long, tok As String, n As Long
    If Len(s) = 0 Then Exit Function

    tokens = Split(s, ";")
    n = -1
    For i = LBound(tokens) To UBound(tokens)
        tok = Trim$(tokens(i))
        If Len(tok) > 0 Then
            n = n + 1
            ReDim Preserve cleaned(0 To n)
            cleaned(n) = tok
        End If
    Next i

    If n >= 0 Then SplitBySemicolon = cleaned
End Function

' Safely check whether a Variant holds an allocated, non-empty array
Private Function IsArrayAllocated(arr As Variant) As Boolean
    On Error GoTo ErrH
    If IsArray(arr) Then
        If Not IsError(LBound(arr)) And Not IsError(UBound(arr)) Then
            IsArrayAllocated = (UBound(arr) - LBound(arr) + 1) > 0
        End If
    End If
    Exit Function
ErrH:
End Function

' Write a single event row to the output sheet at row r
' Columns: A=Subject, B=Start Date, C=Start Time, D=Attendee
Private Sub WriteEventRow_OneAttendee(ByRef ws As Worksheet, ByVal r As Long, _
                                      ByVal subj As String, ByVal d As Variant, ByVal t As Variant, _
                                      ByVal attendee As String)
    ws.Cells(r, 1).Value = subj

    ' Write date, attempting to parse string dates into proper Date values
    If Not IsEmpty(d) Then
        If IsDate(d) Then
            ws.Cells(r, 2).Value = CDate(d)
        Else
            Dim dd As Date
            If TryParseDateTime(d, dd) Then ws.Cells(r, 2).Value = Int(dd) Else ws.Cells(r, 2).Value = d
        End If
    End If

    ' Write time, formatting as hh:mm
    If Not IsEmpty(t) Then
        If IsDate(t) Then
            ws.Cells(r, 3).Value = Format$(CDate(t), "hh:mm")
        Else
            Dim tt As Date
            If TryParseDateTime(t, tt) Then ws.Cells(r, 3).Value = Format$(tt, "hh:mm") Else ws.Cells(r, 3).Value = t
        End If
    End If

    ws.Cells(r, 4).Value = attendee
End Sub

' Return an empty string for Null, Empty, or Error variant values; otherwise CStr
Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

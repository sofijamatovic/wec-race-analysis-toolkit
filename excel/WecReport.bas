Attribute VB_Name = "WecReport"
'==============================================================================
' WEC Race Analysis Toolkit - one-click post-race report
'
' Data flow:
'   data/wec.db --(python src/export.py)--> data/export/*.csv
'   --[1 Load data]--> hidden Data_* sheets --[2 Generate report]--> report sheets
'   --[3 Export PDF]--> excel/reports/<race>_<car>.pdf
'
' Macros (Alt+F8):
'   SetupWorkbook   builds the Control sheet, buttons and report sheets (run once)
'   LoadData        imports the CSV files written by src/export.py
'   GenerateReport  fills the report for the race and car chosen on Control
'   ExportPdf       saves the report sheets as one PDF
'
' Columns are always looked up by header name, so the CSV column order can change.
'==============================================================================
Option Explicit

Private Const SH_CONTROL As String = "Control"
Private Const SH_SUMMARY As String = "Summary"
Private Const SH_STINTS As String = "Stints"
Private Const SH_PITS As String = "Pit Stops"
Private Const SH_LAPS As String = "Lap Times"
Private Const SH_COMP As String = "Class Comparison"

Private Const ADDR_RACE As String = "C6"
Private Const ADDR_CAR As String = "C8"
Private Const ADDR_DIR As String = "C12"
Private Const ADDR_STATUS As String = "C14"
Private Const ADDR_LAST_RACE As String = "Z1"
Private Const ADDR_LAST_CAR As String = "Z2"

Private Const DATA_FILES As String = "races,car_summary,stints,pit_stops,driver_summary,laps"
Private Const SEC_PER_DAY As Double = 86400#
Private Const FMT_LAP As String = "m:ss.000"
Private Const FMT_DELTA As String = "+0.000;-0.000;0.000"

' Column positions in the array returned by CarCols()
Private Const CC_LABEL As Long = 1
Private Const CC_TEAM As Long = 2
Private Const CC_CLASS As Long = 3
Private Const CC_MANUF As Long = 4
Private Const CC_POSALL As Long = 5
Private Const CC_POSCLASS As Long = 6
Private Const CC_LAPS As Long = 7
Private Const CC_BEST As Long = 8
Private Const CC_AVG As Long = 9
Private Const CC_CLEAN As Long = 10
Private Const CC_GAP As Long = 11
Private Const CC_STINTS As Long = 12
Private Const CC_CONS As Long = 13
Private Const CC_DEG As Long = 14
Private Const CC_PITS As Long = 15
Private Const CC_TOTPIT As Long = 16
Private Const CC_AVGPIT As Long = 17
Private Const CC_PITDC As Long = 18
Private Const CC_PITNC As Long = 19
Private Const CC_CAUTION As Long = 20

'------------------------------------------------------------------------------
' Colours
'------------------------------------------------------------------------------
Private Function ClrAccent() As Long
    ClrAccent = RGB(204, 0, 0)
End Function

Private Function ClrDark() As Long
    ClrDark = RGB(38, 38, 38)
End Function

Private Function ClrMuted() As Long
    ClrMuted = RGB(118, 118, 118)
End Function

Private Function ClrPanel() As Long
    ClrPanel = RGB(242, 242, 242)
End Function

Private Function ClrHighlight() As Long
    ClrHighlight = RGB(255, 226, 226)
End Function

Private Function ClrTeammate() As Long
    ClrTeammate = RGB(237, 237, 237)
End Function

'==============================================================================
' 0. Setup (run once)
'==============================================================================
Public Sub SetupWorkbook()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_CONTROL, Nothing)

    Dim prev As Worksheet, sh As Variant
    Set prev = ws
    For Each sh In Array(SH_SUMMARY, SH_STINTS, SH_PITS, SH_LAPS, SH_COMP)
        Set prev = EnsureSheet(CStr(sh), prev)
        prev.Activate
        ActiveWindow.DisplayGridlines = False
    Next sh

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    With ws
        .Cells.Clear
        On Error Resume Next
        .Buttons.Delete
        On Error GoTo Fail

        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 16
        .Columns("C").ColumnWidth = 46
        .Columns("D").ColumnWidth = 4
        .Columns("E").ColumnWidth = 26
        .Columns("Z").Hidden = True
        .Rows(6).RowHeight = 24
        .Rows(8).RowHeight = 24
        .Rows(10).RowHeight = 24

        .Range("B2").Value = "WEC Race Analysis Toolkit"
        .Range("B2").Font.Size = 20
        .Range("B2").Font.Bold = True
        .Range("B3").Value = "One-click post-race report on official FIA WEC timing data (Al Kamel Systems)"
        .Range("B3").Font.Color = ClrMuted()

        .Range("B6").Value = "Race"
        .Range("B8").Value = "Car"
        .Range("B12").Value = "Data folder"
        .Range("B14").Value = "Status"
        .Range("B6,B8,B12,B14").Font.Bold = True
        .Range("B6,B8").VerticalAlignment = xlCenter

        With .Range(ADDR_RACE & "," & ADDR_CAR)
            .Interior.Color = RGB(255, 242, 204)
            .Font.Size = 12
            .VerticalAlignment = xlCenter
            .Borders.Color = RGB(191, 191, 191)
        End With

        .Range(ADDR_DIR).Value = DefaultDataDir()
        .Range(ADDR_DIR).Font.Color = ClrMuted()
        .Range(ADDR_STATUS).Value = "Click '1  Load data' to start."

        .Range("B17").Value = "How to use"
        .Range("B17").Font.Bold = True
        .Range("B18").Value = "1. Run the Python pipeline once: python src/run_all.py  (downloads timing data, builds the database, writes data/export/*.csv)"
        .Range("B19").Value = "2. Click 'Load data' to import the CSV files into this workbook."
        .Range("B20").Value = "3. Pick a race, then a car (yellow cells). Click 'Generate report'."
        .Range("B21").Value = "4. Click 'Export PDF' to save the report in excel\reports\."
        .Range("B23").Value = "Metric definitions (clean lap, consistency, degradation, pit time): see README.md and sql/views.sql."
        .Range("B23").Font.Color = ClrMuted()
    End With

    AddButton ws, "E6", "1  Load data", "LoadData"
    AddButton ws, "E8", "2  Generate report", "GenerateReport"
    AddButton ws, "E10", "3  Export PDF", "ExportPdf"

    DeleteEmptyDefaultSheets
    ws.Activate
    ws.Range(ADDR_RACE).Select

Done:
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    MsgBox "Setup failed: " & Err.Description, vbCritical
    Resume Done
End Sub

'==============================================================================
' 1. Load data
'==============================================================================
Public Sub LoadData()
    On Error GoTo Fail
    Dim dataDir As String
    dataDir = DataFolder()

    Dim files() As String, i As Long, f As String
    files = Split(DATA_FILES, ",")
    For i = LBound(files) To UBound(files)
        f = dataDir & "\" & files(i) & ".csv"
        If Dir(f) = "" Then
            MsgBox "Missing file:" & vbLf & f & vbLf & vbLf & _
                   "Run the Python pipeline first:  python src/run_all.py", vbExclamation
            Exit Sub
        End If
    Next i

    Application.ScreenUpdating = False
    For i = LBound(files) To UBound(files)
        SetStatus "Loading " & files(i) & ".csv ..."
        ImportCsv dataDir & "\" & files(i) & ".csv", "Data_" & files(i)
    Next i

    SetDefaults
    DefineLists
    ApplyValidation
    SetStatus "Data loaded from " & dataDir & " (" & Format$(Now, "yyyy-mm-dd hh:nn") & _
              "). Pick a race and a car, then click 'Generate report'."
    ThisWorkbook.Worksheets(SH_CONTROL).Activate

Done:
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    MsgBox "Loading data failed: " & Err.Description, vbCritical
    Resume Done
End Sub

Private Sub ImportCsv(ByVal path As String, ByVal sheetName As String)
    Dim ws As Worksheet, last As Worksheet
    Set last = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set ws = EnsureSheet(sheetName, last)
    ws.Cells.Clear

    Dim qt As QueryTable
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 65001           ' UTF-8 (driver names like Sebastien with accents)
        .TextFileDecimalSeparator = "."     ' CSV always uses "." whatever the Windows locale
        .TextFileThousandsSeparator = ","
        .AdjustColumnWidth = False
        .RefreshStyle = xlOverwriteCells
        .Refresh BackgroundQuery:=False
    End With
    qt.Delete

    ' clean up the connection and the name Excel creates for each import
    Dim k As Long
    For k = ws.Names.Count To 1 Step -1
        ws.Names(k).Delete
    Next k
    For k = ThisWorkbook.Connections.Count To 1 Step -1
        ThisWorkbook.Connections(k).Delete
    Next k

    ws.Visible = xlSheetHidden
End Sub

' Pick Le Mans and the first Toyota by default (or the first race / car).
Private Sub SetDefaults()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_CONTROL)

    Dim races As Variant, race As String
    races = SelectRows("races", Array(), Array("event"))
    race = CStr(ws.Range(ADDR_RACE).Value)
    If Not InList(races, race) Then
        If InList(races, "LE MANS") Then race = "LE MANS" Else race = CStr(races(1, 1))
        ws.Range(ADDR_RACE).Value = race
    End If

    Dim cars As Variant, car As String, i As Long
    cars = SelectRows("car_summary", Array("event", race), Array("car_label"))
    car = CStr(ws.Range(ADDR_CAR).Value)
    If Not InList(cars, car) Then
        car = CStr(cars(1, 1))
        For i = 1 To UBound(cars, 1)
            If InStr(1, CStr(cars(i, 1)), "Toyota", vbTextCompare) > 0 Then
                car = CStr(cars(i, 1))
                Exit For
            End If
        Next i
        ws.Range(ADDR_CAR).Value = car
    End If
End Sub

' RaceList = all races. CarList = only the cars of the race selected in C6
' (works because car_summary.csv is sorted by race).
Private Sub DefineLists()
    Dim wsR As Worksheet, wsC As Worksheet
    Set wsR = ThisWorkbook.Worksheets("Data_races")
    Set wsC = ThisWorkbook.Worksheets("Data_car_summary")

    Dim nR As Long, evR As String
    nR = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    evR = ColLetter(ColIndex(wsR, "event"))
    AddName "RaceList", "='" & wsR.Name & "'!$" & evR & "$2:$" & evR & "$" & nR

    Dim ev As String, lb As String, src As String, sel As String
    ev = ColLetter(ColIndex(wsC, "event"))
    lb = ColLetter(ColIndex(wsC, "car_label"))
    src = "'" & wsC.Name & "'!"
    sel = SH_CONTROL & "!$C$6"
    AddName "CarList", "=OFFSET(" & src & "$" & lb & "$1," & _
        "MATCH(" & sel & "," & src & "$" & ev & ":$" & ev & ",0)-1,0," & _
        "COUNTIF(" & src & "$" & ev & ":$" & ev & "," & sel & "),1)"
End Sub

Private Sub ApplyValidation()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_CONTROL)
    With ws.Range(ADDR_RACE).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=RaceList"
        .InCellDropdown = True
    End With
    With ws.Range(ADDR_CAR).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CarList"
        .InCellDropdown = True
    End With
End Sub

'==============================================================================
' 2. Generate report
'==============================================================================
Public Sub GenerateReport()
    On Error GoTo Fail
    If Not SheetExists("Data_car_summary") Then
        MsgBox "Load the data first (button 1).", vbExclamation
        Exit Sub
    End If

    Dim race As String, car As String
    race = Trim$(CStr(ThisWorkbook.Worksheets(SH_CONTROL).Range(ADDR_RACE).Value))
    car = Trim$(CStr(ThisWorkbook.Worksheets(SH_CONTROL).Range(ADDR_CAR).Value))

    Dim info As Variant
    info = SelectRows("car_summary", Array("event", race, "car_label", car), CarCols())
    If IsEmpty(info) Then
        MsgBox "'" & car & "' did not race at " & race & "." & vbLf & _
               "Pick a car from the drop-down list for this race.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    SetStatus "Building report for " & car & " at " & race & " ..."

    BuildSummary race, car, info
    BuildStints race, car
    BuildPitStops race, car
    BuildLapTimes race, car
    BuildComparison race, car, CStr(info(1, CC_CLASS))

    Dim sh As Variant
    Application.PrintCommunication = False
    For Each sh In Array(SH_SUMMARY, SH_STINTS, SH_PITS, SH_LAPS, SH_COMP)
        SetupPrint ThisWorkbook.Worksheets(CStr(sh))
    Next sh
    Application.PrintCommunication = True

    With ThisWorkbook.Worksheets(SH_CONTROL)
        .Range(ADDR_LAST_RACE).Value = race
        .Range(ADDR_LAST_CAR).Value = car
    End With
    SetStatus "Report generated for " & car & " at " & race & " (" & Format$(Now, "yyyy-mm-dd hh:nn") & ")."
    ThisWorkbook.Worksheets(SH_SUMMARY).Activate
    ThisWorkbook.Worksheets(SH_SUMMARY).Range("A1").Select

Done:
    Application.PrintCommunication = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    MsgBox "Report failed: " & Err.Description, vbCritical
    Resume Done
End Sub

Private Function CarCols() As Variant
    CarCols = Array("car_label", "team", "class", "manufacturer", "pos_overall", "pos_class", _
                    "laps", "best_lap_s", "avg_clean_lap_s", "clean_laps", "pace_gap_s", _
                    "stints", "consistency_s", "avg_deg_s_per_lap", "pit_stops", "total_pit_s", _
                    "avg_pit_s", "avg_pit_driver_change_s", "avg_pit_no_change_s", "stops_under_caution")
End Function

'------------------------------------------------------------------------------
' Summary: KPI cards + driver table
'------------------------------------------------------------------------------
Private Sub BuildSummary(ByVal race As String, ByVal car As String, info As Variant)
    Dim ws As Worksheet
    Set ws = ResetSheet(SH_SUMMARY)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:F").ColumnWidth = 26
    ws.Columns("G:I").ColumnWidth = 18

    Dim cls As String, classSize As Long
    cls = CStr(info(1, CC_CLASS))
    classSize = CountRows("car_summary", Array("event", race, "class", cls))

    WriteTitle ws, "Post-race report  |  " & race & "  |  " & car, _
        cls & "  |  " & CStr(info(1, CC_MANUF)) & _
        "  |  Source: FIA WEC official timing (Al Kamel Systems)  |  Generated " & Format$(Now, "dd mmm yyyy")

    WriteSection ws.Range("B5"), "Result"
    Card ws.Range("B6"), "Class position", "P" & info(1, CC_POSCLASS) & " of " & classSize, ""
    Card ws.Range("C6"), "Overall position", "P" & info(1, CC_POSALL), ""
    Card ws.Range("D6"), "Laps completed", info(1, CC_LAPS), "0"
    Card ws.Range("E6"), "Stints", info(1, CC_STINTS), "0"
    Card ws.Range("F6"), "Pit stops", info(1, CC_PITS), "0"

    WriteSection ws.Range("B9"), "Pace"
    Card ws.Range("B10"), "Best lap", Secs(info(1, CC_BEST)), FMT_LAP
    Card ws.Range("C10"), "Average clean lap", Secs(info(1, CC_AVG)), FMT_LAP
    Card ws.Range("D10"), "Gap to class best average (s)", info(1, CC_GAP), FMT_DELTA
    Card ws.Range("E10"), "Consistency: avg stint SD (s)", info(1, CC_CONS), "0.000"
    Card ws.Range("F10"), "Avg degradation (s/lap)", info(1, CC_DEG), FMT_DELTA

    WriteSection ws.Range("B13"), "Pit lane"
    Card ws.Range("B14"), "Total time in pit lane", Secs(info(1, CC_TOTPIT)), "[h]:mm:ss"
    Card ws.Range("C14"), "Average stop (s)", info(1, CC_AVGPIT), "0.0"
    Card ws.Range("D14"), "Avg stop with driver change (s)", info(1, CC_PITDC), "0.0"
    Card ws.Range("E14"), "Avg stop, same driver (s)", info(1, CC_PITNC), "0.0"
    Card ws.Range("F14"), "Stops under FCY / safety car", info(1, CC_CAUTION), "0"

    WriteSection ws.Range("B17"), "Drivers"
    Dim d As Variant, hdr As Range
    d = SelectRows("driver_summary", Array("event", race, "car_label", car), _
        Array("driver_name", "laps", "clean_laps", "stints", "best_lap_s", "avg_clean_lap_s", _
              "consistency_s", "avg_deg_s_per_lap"))
    ToExcelTime d, Array(5, 6)
    Set hdr = ws.Range("B18")
    WriteTable hdr, Array("Driver", "Laps", "Clean laps", "Stints", "Best lap", "Avg clean lap", _
                          "Consistency SD (s)", "Avg deg (s/lap)"), d
    FormatCols hdr, d, Array(5, 6), FMT_LAP
    FormatCols hdr, d, Array(7), "0.000"
    FormatCols hdr, d, Array(8), FMT_DELTA

    Dim r As Long
    r = hdr.Row + RowCount(d) + 3
    With ws.Cells(r, 2)
        .Value = "Notes"
        .Font.Bold = True
    End With
    ws.Cells(r + 1, 2).Value = "Clean lap: green flag, not an in/out lap, not lap 1, and within 107% of the stint median."
    ws.Cells(r + 2, 2).Value = "Consistency: standard deviation of clean lap times within a stint, averaged over stints (lower = more consistent)."
    ws.Cells(r + 3, 2).Value = "Degradation: slope of clean lap time vs lap in stint. Net of fuel burn, so tyre wear is partly hidden."
    ws.Cells(r + 4, 2).Value = "Positions are approximated from laps completed and race time; the official classification can differ."
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 4, 2)).Font.Color = ClrMuted()
End Sub

'------------------------------------------------------------------------------
' Stints: table + two charts
'------------------------------------------------------------------------------
Private Sub BuildStints(ByVal race As String, ByVal car As String)
    Dim ws As Worksheet
    Set ws = ResetSheet(SH_STINTS)
    WriteTitle ws, "Stints  |  " & race & "  |  " & car, _
        "One row per stint (out-lap to next in-lap). Colour scale: green = better, red = worse."

    Dim d As Variant
    d = SelectRows("stints", Array("event", race, "car_label", car), _
        Array("stint_no", "driver_name", "first_lap", "last_lap", "laps", "clean_laps", _
              "best_lap_s", "avg_clean_lap_s", "std_clean_lap_s", "deg_s_per_lap"))
    If IsEmpty(d) Then
        ws.Range("B5").Value = "No stints found for this car."
        Exit Sub
    End If

    ToExcelTime d, Array(7, 8)
    Dim hdr As Range, n As Long
    Set hdr = ws.Range("B5")
    n = RowCount(d)
    WriteTable hdr, Array("Stint", "Driver", "First lap", "Last lap", "Laps", "Clean laps", _
                          "Best lap", "Avg clean lap", "Consistency SD (s)", "Deg (s/lap)"), d
    FormatCols hdr, d, Array(7, 8), FMT_LAP
    FormatCols hdr, d, Array(9), "0.000"
    FormatCols hdr, d, Array(10), FMT_DELTA
    AddHeatmap hdr.Offset(1, 7).Resize(n, 1)
    AddHeatmap hdr.Offset(1, 8).Resize(n, 1)
    AddHeatmap hdr.Offset(1, 9).Resize(n, 1)
    ws.Columns("B:K").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 22

    Dim ch As Chart, xs As Range, vals As Range
    Set xs = hdr.Offset(1, 0).Resize(n, 1)

    Set vals = hdr.Offset(1, 7).Resize(n, 1)
    Set ch = NewChart(ws, ws.Range("M5"), 560, 280)
    AddSeries ch, "Avg clean lap", vals, xs
    ch.ChartType = xlColumnClustered
    StyleChart ch, "Average clean lap by stint", False
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = ClrAccent()
    ch.Axes(xlValue).TickLabels.NumberFormat = "m:ss.0"
    SetValueAxis ch, vals, 1# / SEC_PER_DAY
    AxisTitles ch, "Stint", ""

    Set vals = hdr.Offset(1, 9).Resize(n, 1)
    Set ch = NewChart(ws, ws.Range("M25"), 560, 280)
    AddSeries ch, "Degradation", vals, xs
    ch.ChartType = xlColumnClustered
    StyleChart ch, "Degradation by stint (s/lap, positive = getting slower)", False
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = ClrDark()
    ch.Axes(xlValue).TickLabels.NumberFormat = "0.00"
    AxisTitles ch, "Stint", "s/lap"
End Sub

'------------------------------------------------------------------------------
' Pit stops: table + chart (red = driver change)
'------------------------------------------------------------------------------
Private Sub BuildPitStops(ByVal race As String, ByVal car As String)
    Dim ws As Worksheet
    Set ws = ResetSheet(SH_PITS)
    WriteTitle ws, "Pit stops  |  " & race & "  |  " & car, _
        "Pit lane time = time from pit entry to pit exit, as reported by official timing."

    Dim d As Variant
    d = SelectRows("pit_stops", Array("event", race, "car_label", car), _
        Array("stop_no", "in_lap", "driver_in", "driver_out", "driver_change", "under_flag", "pit_time_s"))
    If IsEmpty(d) Then
        ws.Range("B5").Value = "No pit stops for this car."
        Exit Sub
    End If

    Dim hdr As Range, n As Long, i As Long
    Set hdr = ws.Range("B5")
    n = RowCount(d)
    WriteTable hdr, Array("Stop", "In lap", "Driver in", "Driver out", "Driver change", _
                          "Flag on in-lap", "Pit lane time (s)"), d
    FormatCols hdr, d, Array(7), "0.0"
    AddHeatmap hdr.Offset(1, 6).Resize(n, 1)
    ws.Columns("B:H").ColumnWidth = 14
    ws.Columns("D:E").ColumnWidth = 22

    Dim ch As Chart
    Set ch = NewChart(ws, ws.Range("J5"), 600, 300)
    AddSeries ch, "Pit lane time", hdr.Offset(1, 6).Resize(n, 1), hdr.Offset(1, 0).Resize(n, 1)
    ch.ChartType = xlColumnClustered
    StyleChart ch, "Pit lane time per stop (red = driver change)", False
    AxisTitles ch, "Stop", "seconds"
    For i = 1 To n
        If CStr(d(i, 5)) = "Yes" Then
            ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = ClrAccent()
        Else
            ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = RGB(166, 166, 166)
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Lap times: clean laps grouped by driver + scatter chart (one series per driver)
'------------------------------------------------------------------------------
Private Sub BuildLapTimes(ByVal race As String, ByVal car As String)
    Dim ws As Worksheet
    Set ws = ResetSheet(SH_LAPS)
    WriteTitle ws, "Clean lap times  |  " & race & "  |  " & car, _
        "Only clean laps (green flag, no in/out laps, within 107% of the stint median)."

    Dim drivers As Variant, laps As Variant
    drivers = SelectRows("driver_summary", Array("event", race, "car_label", car), Array("driver_name"))
    laps = SelectRows("laps", Array("event", race, "car_label", car, "is_clean", "1"), _
                      Array("lap_number", "driver_name", "stint_no", "lap_time_s"))
    If IsEmpty(laps) Or IsEmpty(drivers) Then
        ws.Range("B5").Value = "No clean laps for this car."
        Exit Sub
    End If

    ' regroup rows driver by driver, so each driver is one block (one chart series)
    Dim n As Long, nd As Long, i As Long, j As Long, k As Long, p As Long
    n = RowCount(laps)
    nd = RowCount(drivers)
    Dim out() As Variant, startRow() As Long, cnt() As Long
    ReDim out(1 To n, 1 To 4)
    ReDim startRow(1 To nd)
    ReDim cnt(1 To nd)
    For i = 1 To nd
        startRow(i) = p + 1
        For j = 1 To n
            If CStr(laps(j, 2)) = CStr(drivers(i, 1)) Then
                p = p + 1
                For k = 1 To 4
                    out(p, k) = laps(j, k)
                Next k
            End If
        Next j
        cnt(i) = p - startRow(i) + 1
    Next i

    Dim hdr As Range
    Set hdr = ws.Range("B5")
    WriteTable hdr, Array("Lap", "Driver", "Stint", "Lap time (s)"), out
    FormatCols hdr, out, Array(4), "0.000"
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D:E").ColumnWidth = 12

    Dim ch As Chart, s As Series, colours As Variant
    colours = Array(ClrAccent(), ClrDark(), RGB(68, 114, 196), RGB(237, 125, 49))
    Set ch = NewChart(ws, ws.Range("G5"), 680, 380)
    For i = 1 To nd
        If cnt(i) > 0 Then
            Set s = AddSeries(ch, CStr(drivers(i, 1)), _
                              hdr.Offset(startRow(i), 3).Resize(cnt(i), 1), _
                              hdr.Offset(startRow(i), 0).Resize(cnt(i), 1))
        End If
    Next i
    ch.ChartType = xlXYScatter
    StyleChart ch, "Clean lap times by driver", True
    ch.Legend.Position = xlLegendPositionBottom
    For i = 1 To ch.SeriesCollection.Count
        With ch.SeriesCollection(i)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 4
            .MarkerBackgroundColor = colours((i - 1) Mod 4)
            .MarkerForegroundColor = colours((i - 1) Mod 4)
        End With
    Next i
    If p > 0 Then SetValueAxis ch, hdr.Offset(1, 3).Resize(p, 1), 0.5
    ch.Axes(xlCategory).MinimumScale = 0
    AxisTitles ch, "Lap", "Lap time (s)"
End Sub

'------------------------------------------------------------------------------
' Class comparison: every car in the same class + gap chart
'------------------------------------------------------------------------------
Private Sub BuildComparison(ByVal race As String, ByVal car As String, ByVal cls As String)
    Dim ws As Worksheet
    Set ws = ResetSheet(SH_COMP)
    WriteTitle ws, "Class comparison  |  " & race & "  |  " & cls, _
        "Selected car highlighted in red, team-mates in grey. Gap = average clean lap minus the best average in the class."

    Dim d As Variant
    d = SelectRows("car_summary", Array("event", race, "class", cls), _
        Array("pos_class", "car_label", "laps", "best_lap_s", "avg_clean_lap_s", "pace_gap_s", _
              "consistency_s", "avg_deg_s_per_lap", "pit_stops", "avg_pit_s", "total_pit_s"))
    If IsEmpty(d) Then Exit Sub

    ToExcelTime d, Array(4, 5, 11)
    Dim hdr As Range, n As Long, i As Long
    Set hdr = ws.Range("B5")
    n = RowCount(d)
    WriteTable hdr, Array("Pos", "Car", "Laps", "Best lap", "Avg clean lap", "Gap to best (s)", _
                          "Consistency SD (s)", "Avg deg (s/lap)", "Pit stops", "Avg pit (s)", "Total pit time"), d
    FormatCols hdr, d, Array(4, 5), FMT_LAP
    FormatCols hdr, d, Array(6), FMT_DELTA
    FormatCols hdr, d, Array(7), "0.000"
    FormatCols hdr, d, Array(8), FMT_DELTA
    FormatCols hdr, d, Array(10), "0.0"
    FormatCols hdr, d, Array(11), "[h]:mm:ss"
    ws.Columns("B").ColumnWidth = 6
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D:L").ColumnWidth = 13

    Dim team As String, rowRng As Range
    team = TeamOf(car)
    For i = 1 To n
        Set rowRng = hdr.Offset(i, 0).Resize(1, 11)
        If CStr(d(i, 2)) = car Then
            rowRng.Interior.Color = ClrHighlight()
            rowRng.Font.Bold = True
        ElseIf TeamOf(CStr(d(i, 2))) = team Then
            rowRng.Interior.Color = ClrTeammate()
        End If
    Next i

    Dim ch As Chart
    Set ch = NewChart(ws, hdr.Offset(n + 3, 0), 720, 60 + n * 18)
    AddSeries ch, "Gap to class best", hdr.Offset(1, 5).Resize(n, 1), hdr.Offset(1, 1).Resize(n, 1)
    ch.ChartType = xlBarClustered
    StyleChart ch, "Average clean-lap gap to class best (s)", False
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.Axes(xlValue).TickLabels.NumberFormat = "0.0"
    ch.ChartGroups(1).GapWidth = 40
    For i = 1 To n
        With ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor
            If CStr(d(i, 2)) = car Then
                .RGB = ClrAccent()
            ElseIf TeamOf(CStr(d(i, 2))) = team Then
                .RGB = ClrDark()
            Else
                .RGB = RGB(191, 191, 191)
            End If
        End With
    Next i
End Sub

'==============================================================================
' 3. Export PDF
'==============================================================================
Public Sub ExportPdf()
    On Error GoTo Fail
    Dim race As String, car As String
    race = CStr(ThisWorkbook.Worksheets(SH_CONTROL).Range(ADDR_LAST_RACE).Value)
    car = CStr(ThisWorkbook.Worksheets(SH_CONTROL).Range(ADDR_LAST_CAR).Value)
    If race = "" Then
        MsgBox "Generate a report first (button 2).", vbExclamation
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Save the workbook first.", vbExclamation
        Exit Sub
    End If

    Dim outDir As String, f As String
    outDir = ThisWorkbook.Path & "\reports"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    f = outDir & "\" & SafeFileName(race & "_" & car) & ".pdf"

    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(Array(SH_SUMMARY, SH_STINTS, SH_PITS, SH_LAPS, SH_COMP)).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, OpenAfterPublish:=True
    ThisWorkbook.Worksheets(SH_SUMMARY).Select
    SetStatus "PDF saved: " & f
    Exit Sub
Fail:
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SUMMARY).Select
    MsgBox "PDF export failed: " & Err.Description, vbCritical
End Sub

'==============================================================================
' Data helpers
'==============================================================================

' Rows of Data_<key> where every filter matches, as a 1-based 2D array with the
' requested columns. filters = Array("col1", value1, "col2", value2, ...).
' Returns Empty when nothing matches.
Private Function SelectRows(ByVal key As String, filters As Variant, cols As Variant) As Variant
    Dim ws As Worksheet, data As Variant
    Set ws = ThisWorkbook.Worksheets("Data_" & key)
    data = ws.UsedRange.Value

    Dim nf As Long, nc As Long, k As Long
    nf = (UBound(filters) - LBound(filters) + 1) \ 2
    nc = UBound(cols) - LBound(cols) + 1

    Dim fIdx() As Long, fVal() As String, cIdx() As Long
    If nf > 0 Then
        ReDim fIdx(1 To nf)
        ReDim fVal(1 To nf)
        For k = 1 To nf
            fIdx(k) = ColIndex(ws, CStr(filters(LBound(filters) + 2 * (k - 1))))
            fVal(k) = CStr(filters(LBound(filters) + 2 * (k - 1) + 1))
        Next k
    End If
    ReDim cIdx(1 To nc)
    For k = 1 To nc
        cIdx(k) = ColIndex(ws, CStr(cols(LBound(cols) + k - 1)))
    Next k

    Dim r As Long, n As Long, ok As Boolean
    Dim hits() As Long
    ReDim hits(1 To UBound(data, 1))
    For r = 2 To UBound(data, 1)
        ok = True
        For k = 1 To nf
            If CStr(data(r, fIdx(k))) <> fVal(k) Then
                ok = False
                Exit For
            End If
        Next k
        If ok Then
            If Not IsEmpty(data(r, 1)) Then
                n = n + 1
                hits(n) = r
            End If
        End If
    Next r

    If n = 0 Then
        SelectRows = Empty
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(1 To n, 1 To nc)
    For r = 1 To n
        For k = 1 To nc
            out(r, k) = data(hits(r), cIdx(k))
        Next k
    Next r
    SelectRows = out
End Function

Private Function CountRows(ByVal key As String, filters As Variant) As Long
    Dim v As Variant
    v = SelectRows(key, filters, Array(CStr(filters(LBound(filters)))))
    CountRows = RowCount(v)
End Function

Private Function RowCount(v As Variant) As Long
    If IsEmpty(v) Then
        RowCount = 0
    Else
        RowCount = UBound(v, 1)
    End If
End Function

Private Function ColIndex(ws As Worksheet, ByVal header As String) As Long
    Dim m As Variant
    m = Application.Match(header, ws.Rows(1), 0)
    If IsError(m) Then
        Err.Raise vbObjectError + 513, "WecReport", _
            "Column '" & header & "' not found in sheet " & ws.Name & ". Re-run src/export.py and Load data."
    End If
    ColIndex = CLng(m)
End Function

Private Function InList(arr As Variant, ByVal value As String) As Boolean
    Dim i As Long
    If IsEmpty(arr) Then Exit Function
    For i = 1 To UBound(arr, 1)
        If CStr(arr(i, 1)) = value Then
            InList = True
            Exit Function
        End If
    Next i
End Function

' seconds -> Excel time (fraction of a day) so cells can use m:ss.000
Private Function Secs(ByVal v As Variant) As Variant
    If IsEmpty(v) Then
        Secs = Empty
    ElseIf IsNumeric(v) Then
        Secs = CDbl(v) / SEC_PER_DAY
    Else
        Secs = Empty
    End If
End Function

Private Sub ToExcelTime(ByRef arr As Variant, cols As Variant)
    Dim r As Long, c As Variant
    If IsEmpty(arr) Then Exit Sub
    For Each c In cols
        For r = 1 To UBound(arr, 1)
            arr(r, c) = Secs(arr(r, c))
        Next r
    Next c
End Sub

' '#7 Toyota Racing' -> 'Toyota Racing'
Private Function TeamOf(ByVal carLabel As String) As String
    Dim p As Long
    p = InStr(1, carLabel, " ")
    If p > 0 Then TeamOf = Mid$(carLabel, p + 1) Else TeamOf = carLabel
End Function

'==============================================================================
' Formatting helpers
'==============================================================================
Private Function ResetSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    ws.Cells.FormatConditions.Delete
    ws.Cells.Clear
    ws.Cells.Font.Size = 10
    Set ResetSheet = ws
End Function

Private Sub WriteTitle(ws As Worksheet, ByVal title As String, ByVal subtitle As String)
    With ws.Range("B2")
        .Value = title
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = ClrDark()
    End With
    With ws.Range("B3")
        .Value = subtitle
        .Font.Size = 9
        .Font.Color = ClrMuted()
    End With
    ws.Columns("A").ColumnWidth = 3
End Sub

Private Sub WriteSection(target As Range, ByVal text As String)
    With target
        .Value = UCase$(text)
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = ClrAccent()
    End With
End Sub

Private Sub Card(target As Range, ByVal label As String, ByVal v As Variant, ByVal fmt As String)
    With target
        .Value = label
        .Font.Size = 9
        .Font.Color = ClrMuted()
        .Interior.Color = ClrPanel()
        .IndentLevel = 1
    End With
    With target.Offset(1, 0)
        If IsEmpty(v) Then
            .Value = "-"
        ElseIf CStr(v) = "" Then
            .Value = "-"
        Else
            .Value = v
        End If
        If fmt <> "" Then .NumberFormat = fmt
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = ClrDark()
        .Interior.Color = ClrPanel()
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    With target.Resize(2, 1).Borders(xlEdgeLeft)
        .Color = ClrAccent()
        .Weight = xlThick
    End With
End Sub

Private Sub WriteTable(topLeft As Range, headers As Variant, rows As Variant)
    Dim nc As Long, n As Long
    nc = UBound(headers) - LBound(headers) + 1
    With topLeft.Resize(1, nc)
        .Value = headers
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = ClrDark()
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    topLeft.EntireRow.RowHeight = 30

    n = RowCount(rows)
    If n = 0 Then Exit Sub
    With topLeft.Offset(1, 0).Resize(n, nc)
        .Value = rows
        .Borders(xlEdgeBottom).Color = RGB(217, 217, 217)
        If n > 1 Then .Borders(xlInsideHorizontal).Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub FormatCols(topLeft As Range, rows As Variant, cols As Variant, ByVal fmt As String)
    Dim c As Variant, n As Long
    n = RowCount(rows)
    If n = 0 Then Exit Sub
    For Each c In cols
        topLeft.Offset(1, CLng(c) - 1).Resize(n, 1).NumberFormat = fmt
    Next c
End Sub

' green = low (good), red = high (bad)
Private Sub AddHeatmap(target As Range)
    Dim cs As Object
    Set cs = target.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
End Sub

'==============================================================================
' Chart helpers
'==============================================================================
Private Function NewChart(ws As Worksheet, anchor As Range, ByVal w As Double, ByVal h As Double) As Chart
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left, anchor.Top, w, h)
    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop
    Set NewChart = co.Chart
End Function

Private Function AddSeries(ch As Chart, ByVal seriesName As String, vals As Range, xs As Range) As Series
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = seriesName
    s.Values = vals
    s.XValues = xs
    Set AddSeries = s
End Function

Private Sub StyleChart(ch As Chart, ByVal title As String, ByVal showLegend As Boolean)
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.ChartTitle.Font.Size = 12
    ch.ChartTitle.Font.Bold = True
    ch.HasLegend = showLegend
    On Error Resume Next
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(225, 225, 225)
    ch.ChartArea.Format.Line.Visible = msoFalse
End Sub

Private Sub AxisTitles(ch As Chart, ByVal xTitle As String, ByVal yTitle As String)
    If xTitle <> "" Then
        ch.Axes(xlCategory).HasTitle = True
        ch.Axes(xlCategory).AxisTitle.Text = xTitle
    End If
    If yTitle <> "" Then
        ch.Axes(xlValue).HasTitle = True
        ch.Axes(xlValue).AxisTitle.Text = yTitle
    End If
End Sub

Private Sub SetValueAxis(ch As Chart, vals As Range, ByVal pad As Double)
    If Application.WorksheetFunction.Count(vals) = 0 Then Exit Sub
    Dim lo As Double, hi As Double
    lo = Application.WorksheetFunction.Min(vals)
    hi = Application.WorksheetFunction.Max(vals)
    With ch.Axes(xlValue)
        .MaximumScale = hi + pad
        .MinimumScale = lo - pad
    End With
End Sub

'==============================================================================
' Workbook helpers
'==============================================================================
Private Function EnsureSheet(ByVal sheetName As String, ByVal after As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        If after Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=after)
        End If
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Private Function IsToolkitSheet(ByVal sheetName As String) As Boolean
    Select Case sheetName
        Case SH_CONTROL, SH_SUMMARY, SH_STINTS, SH_PITS, SH_LAPS, SH_COMP
            IsToolkitSheet = True
        Case Else
            IsToolkitSheet = (Left$(sheetName, 5) = "Data_")
    End Select
End Function

' remove the empty "Sheet1" a new workbook starts with
Private Sub DeleteEmptyDefaultSheets()
    Dim i As Long, ws As Worksheet
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not IsToolkitSheet(ws.Name) Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True
End Sub

Private Sub AddButton(ws As Worksheet, ByVal addr As String, ByVal caption As String, ByVal macroName As String)
    Dim r As Range, b As Object
    Set r = ws.Range(addr)
    Set b = ws.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
    b.Caption = caption
    b.OnAction = macroName
End Sub

Private Sub AddName(ByVal nameText As String, ByVal refersTo As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersTo
End Sub

Private Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, colNum).Address(True, False), "$")(0)
End Function

Private Function DefaultDataDir() As String
    If ThisWorkbook.Path = "" Then
        DefaultDataDir = "(save this workbook in the excel folder of the repo, then run SetupWorkbook again)"
    Else
        DefaultDataDir = CreateObject("Scripting.FileSystemObject") _
            .GetAbsolutePathName(ThisWorkbook.Path & "\..\data\export")
    End If
End Function

Private Function DataFolder() As String
    Dim p As String
    p = Trim$(CStr(ThisWorkbook.Worksheets(SH_CONTROL).Range(ADDR_DIR).Value))
    If p = "" Or Left$(p, 1) = "(" Then p = DefaultDataDir()
    DataFolder = p
End Function

Private Sub SetStatus(ByVal text As String)
    ThisWorkbook.Worksheets(SH_CONTROL).Range(ADDR_STATUS).Value = text
    Application.StatusBar = False
End Sub

Private Sub SetupPrint(ws As Worksheet)
    On Error Resume Next   ' PageSetup can fail when no printer is installed
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterFooter = "WEC Race Analysis Toolkit  |  &A  |  page &P of &N"
    End With
End Sub

Private Function SafeFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", "#", " ")
        s = Replace(s, CStr(ch), "_")
    Next ch
    Do While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Loop
    SafeFileName = s
End Function

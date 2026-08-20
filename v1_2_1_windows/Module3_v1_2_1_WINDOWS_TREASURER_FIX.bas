Attribute VB_Name = "Module3"
Option Explicit

Private gRunStage As String
Private gAdapterOutcome As String
Private gAdapterNote As String

' ============================================================================
' COOK COUNTY PROPERTY DUE DILIGENCE v1.2.1a - MAC + WINDOWS EXCEL
'
' Design:
'   * Enter a Cook County PIN in the yellow Start-sheet PIN field (row 4)
'   * GenerateCookPropertyReport calls each public source independently
'   * A failed/slow source is logged and the run continues
'   * Results populate workbook sheets and export to one PDF
'
' Production target: current Excel for Mac and Windows Excel.
' Mac HTTP uses Microsoft AppleScriptTask + the included CookPropertyHTTP helper.
' Windows HTTP uses late-bound WinHTTP. Parsing uses pure VBA for cross-platform compatibility.
'
' IMPORTANT:
'   Parcel Universe exposes econ_enterprise_zone_num and its data year.
'   v1.0.4 uses that field as a Cook County Assessor spatial SIGNAL only.
'   The official Illinois DCEO map remains the boundary-verification source.
' ============================================================================

Private Const SOCRATA_BASE As String = "https://datacatalog.cookcountyil.gov/resource/"
Private Const ADDRESS_DATASET As String = "3723-97qp"
Private Const UNIVERSE_DATASET As String = "pabr-t5kh"
Private Const ASSESSMENT_DATASET As String = "uzyt-m557"
Private Const SALES_DATASET As String = "wvhk-k5uv"
Private Const APPEALS_DATASET As String = "y282-6ig3"
Private Const BOR_DATASET As String = "7pny-nedm"
Private Const PERMITS_DATASET As String = "6yjf-dfxs"


Public Sub SetupCookDueDiligence_v121()
    On Error GoTo SetupError

    EnsureV121WorkbookStructure
    InstallRunButton_v121
    ApplyV121WorkbookSettings
    InstallTaxDetailHyperlinks
    InstallTreasurerLookupButton_v121
    UpdatePlatformReadyStatus

    ThisWorkbook.Worksheets("Start").Range("B8").value = "v1.2.1 installed - ready"
    MsgBox "v1.2.1 is installed." & vbCrLf & vbCrLf & _
           "Treasurer is automatic. Safari now validates the Overview first, captures it, then captures the 20-Year History in the same property session and returns both for parsing." & vbCrLf & _
           "PDFs save to Downloads and stay closed.", _
           vbInformation, "Cook County Property Due Diligence v1.2.1"
    Exit Sub

SetupError:
    MsgBox "v1.2.1 setup error " & CStr(Err.Number) & ": " & Err.Description, _
           vbCritical, "Cook County Property Due Diligence v1.2.1"
End Sub

Private Sub ApplyV121WorkbookSettings()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Config")
    ws.Range("B17").value = ""
    ws.Range("C17").value = "Blank = Downloads folder automatically"
    ws.Range("D17").value = "v1.2.1 defaults to the user's Downloads folder."
    ws.Range("B18").value = "NO"
    ws.Range("D18").value = "NO saves the PDF without launching Adobe/Preview."

    ThisWorkbook.Worksheets("Start").Range("B4:D4").NumberFormat = "@"
    ThisWorkbook.Worksheets("Property").Range("B5").NumberFormat = "@"
    ThisWorkbook.Worksheets("Report").Range("B5").NumberFormat = "@"
    ThisWorkbook.Worksheets("Report").Range("E8").NumberFormat = "@"
    ThisWorkbook.Worksheets("Sales-Deeds").Range("C5:C200").NumberFormat = "@"
    ThisWorkbook.Worksheets("Permits").Range("C5:D200").NumberFormat = "@"
    ThisWorkbook.Worksheets("Documents").Range("A5:A200").NumberFormat = "@"
    ThisWorkbook.Worksheets("Tax History").Range("E5:E50").NumberFormat = "@"

    With ThisWorkbook.Worksheets("Report")
        .Range("A14:H21").WrapText = True
        .Range("A14:H21").VerticalAlignment = xlTop
        .Range("A24:H24").WrapText = True
        .Range("A24:H24").VerticalAlignment = xlTop
        .rows(24).RowHeight = 72
    End With

    Set ws = ThisWorkbook.Worksheets("Tax Detail")
    With ws
        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 10
        .Range("A1:H70").WrapText = True
        .Range("A1:H70").VerticalAlignment = xlTop

        .Columns("A").ColumnWidth = 26
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 22
        .Columns("F").ColumnWidth = 24
        .Columns("G").ColumnWidth = 25
        .Columns("H").ColumnWidth = 40

        .Range("A1:H1").Interior.Color = RGB(23, 50, 77)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
        .Range("A1:H1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .rows(1).RowHeight = 30

        .Range("A2:H2").Font.Color = RGB(100, 116, 139)
        .Range("A2:H2").Font.Italic = True
        .rows(2).RowHeight = 24

        .Range("A4:H4").Interior.Color = RGB(241, 245, 249)
        .Range("A4:H4").Font.Bold = True
        .rows(4).RowHeight = 36
        .Range("B4").NumberFormat = "0.000"
        .Range("D4").NumberFormat = "@"

        .Range("A7:G7").Interior.Color = RGB(217, 226, 243)
        .Range("A7:G7").Font.Bold = True
        .Range("A8:G8").Interior.Color = RGB(226, 232, 240)
        .Range("A8:G8").Font.Bold = True

        .Range("A17:H17").Interior.Color = RGB(217, 226, 243)
        .Range("A17:H17").Font.Bold = True
        .Range("A18:H22").Interior.Color = RGB(248, 250, 252)
        .Range("A18:H22").Font.Bold = False
        .Range("A18,C18,E18,G18,A20,C20,E20,G20,A22,C22,E22,G22").Font.Bold = True

        .Range("A25:E25").Interior.Color = RGB(234, 240, 246)
        .Range("A25:E25").Font.Bold = True
        .Range("A26:E26").Interior.Color = RGB(226, 232, 240)
        .Range("A26:E26").Font.Bold = True

        .Range("A35:G35").Interior.Color = RGB(217, 226, 243)
        .Range("A35:G35").Font.Bold = True
        .Range("A36:G36").Interior.Color = RGB(226, 232, 240)
        .Range("A36:G36").Font.Bold = True

        .Range("A45:E45").Interior.Color = RGB(217, 226, 243)
        .Range("A45:E45").Font.Bold = True
        .Range("A46:E46").Interior.Color = RGB(226, 232, 240)
        .Range("A46:E46").Font.Bold = True

        .Range("G25:H25").Interior.Color = RGB(234, 240, 246)
        .Range("G25:H25").Font.Bold = True
        .Range("G70:H70").Interior.Color = RGB(255, 242, 204)
        .Range("G70:H70").Font.Bold = True
        .Range("G72:H72").Interior.Color = RGB(226, 232, 240)
        .Range("G72:H72").Font.Bold = True

        .Range("G74:H74").Interior.Color = RGB(220, 252, 231)
        .Range("G74:H74").Font.Bold = True

        .Range("H26:H33").Font.Color = RGB(31, 78, 121)
        .Range("H26:H33").Font.Underline = True

        .rows("17:22").RowHeight = 30
        .rows("25:33").RowHeight = 26
        .rows(35).RowHeight = 28
        .rows(45).RowHeight = 28
    End With

    With ThisWorkbook.Worksheets("Appeals")
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 24
        .Columns("F:H").ColumnWidth = 14
        .Columns("I").ColumnWidth = 40
        .Range("A1:I200").WrapText = True
    End With
End Sub

Private Sub EnsureV121WorkbookStructure()
    Dim ws As Worksheet, r As Long, found As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Tax Detail")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Tax History"))
        ws.Name = "Tax Detail"
    End If

    With ws
        .Cells.Clear

        .Range("A1").value = "COOK COUNTY TAX INTELLIGENCE"
        .Range("A2").value = "Property Tax Portal history + Cook County Treasurer overview and 20-year research"

        ' Snapshot cards.
        .Range("A4").value = "CURRENT RATE"
        .Range("B4").value = ""
        .Range("C4").value = "TAX CODE"
        .Range("D4").value = ""
        .Range("E4").value = "LATEST BILL"
        .Range("F4").value = ""
        .Range("G4").value = "PORTAL REFUND"
        .Range("H4").value = ""

        ' Portal table.
        .Range("A7").value = "PROPERTY TAX PORTAL - 5-YEAR RATE / BILL HISTORY"
        .Range("A8").value = "Tax Year"
        .Range("B8").value = "Composite Rate"
        .Range("C8").value = "Tax Code"
        .Range("D8").value = "Billed Amount"
        .Range("E8").value = "Payment / Status"
        .Range("F8").value = "Exemptions"
        .Range("G8").value = "Tax Sale / Delinquency"

        ' Treasurer overview cards.
        .Range("A17").value = "COOK COUNTY TREASURER - PROPERTY OVERVIEW"
        .Range("A18").value = "Lookup Status"
        .Range("C18").value = "Property Location"
        .Range("E18").value = "Refund / Overpayment"
        .Range("G18").value = "Current Amount Due"

        .Range("A20").value = "20-Year Start"
        .Range("C20").value = "20-Year Latest"
        .Range("E20").value = "20-Year Change"
        .Range("G20").value = "Percent Change"

        .Range("A22").value = "Debt Attributed"
        .Range("C22").value = "Treasurer Property Value"
        .Range("E22").value = "Debt / Property Value"
        .Range("G22").value = "Mailing Information"

        ' Treasurer exemptions.
        .Range("A25").value = "TREASURER - RECENT EXEMPTION CHECK"
        .Range("A26").value = "Exemption"
        .Range("B26").value = "2024"
        .Range("C26").value = "2023"
        .Range("D26").value = "2022"
        .Range("E26").value = "2021"

        .Range("A27").value = "Homeowner"
        .Range("A28").value = "Senior Citizen"
        .Range("A29").value = "Senior Freeze"
        .Range("A30").value = "Returning Veteran"
        .Range("A31").value = "Disabled Person"
        .Range("A32").value = "Disabled Veteran"

        ' Current installments.
        .Range("A35").value = "TREASURER - CURRENT / RECENT INSTALLMENTS"
        .Range("A36").value = "Tax Year"
        .Range("B36").value = "Installment"
        .Range("C36").value = "Original Billed"
        .Range("D36").value = "Due Date"
        .Range("E36").value = "Current Amount Due"
        .Range("F36").value = "Tax"
        .Range("G36").value = "Interest"

        ' 20-year history.
        .Range("A45").value = "TREASURER - 20-YEAR TAX BILL HISTORY"
        .Range("A46").value = "Tax Year"
        .Range("B46").value = "Total Tax Bill"
        .Range("C46").value = "$ Change vs Prior"
        .Range("D46").value = "% Change vs Prior"
        .Range("E46").value = "Source"

        ' Side-tab quick links / extra information.
        .Range("G25").value = "TREASURER SIDE TABS / ADDITIONAL CHECKS"
        .Range("G26").value = "Overview / Payments"
        .Range("H26").value = "https://www.cookcountytreasurer.com/setsearchparameters.aspx"
        .Range("G27").value = "20-Year Tax Bill History"
        .Range("H27").value = "https://www.cookcountytreasurer.com/taxbillhistoryresults.aspx"
        .Range("G28").value = "Taxing Districts' Financials"
        .Range("H28").value = "https://www.cookcountytreasurer.com/taxingdistrictsresults.aspx"
        .Range("G29").value = "Overpayment Refunds"
        .Range("H29").value = "https://www.cookcountytreasurer.com/"
        .Range("G30").value = "Tax Exemptions"
        .Range("H30").value = "https://www.cookcountytreasurer.com/"
        .Range("G31").value = "Sold Taxes"
        .Range("H31").value = "https://www.cookcountytreasurer.com/"
        .Range("G32").value = "PTAB Decisions"
        .Range("H32").value = "https://www.cookcountytreasurer.com/"
        .Range("G33").value = "Debt to Property Value"
        .Range("H33").value = "https://www.cookcountytreasurer.com/"

        .Range("G70").value = "AUTOMATION NOTE"
        .Range("H70").value = "Treasurer property pages use an older session-sensitive ASP.NET workflow. v1.2.1 tries both the overview and 20-year result pages automatically. If the site requires an interactive property session, Treasurer remains PARTIAL while all other report sources continue."

        .Range("G72").value = "SOURCE STATUS"
        .Range("H72").value = "Property Tax Portal = primary automatic tax-detail source. Treasurer = independent cross-check / expanded historical source."

        .Range("G74").value = "BROWSER ASSIST"
        .Range("H74").value = "Mac: AUTO in v1.2.1 when direct Treasurer HTTP is incomplete. Safari opens only for Treasurer, after the other fast lookup attempt. The green button below remains available to rerun Treasurer manually."
    End With

    ' Source-status row.
    Set ws = ThisWorkbook.Worksheets("Start")
    found = False
    For r = 12 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Cook County Treasurer", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        r = 23
        ws.Cells(r, 1).value = "Cook County Treasurer"
        ws.Cells(r, 2).value = "Not run"
        ws.Cells(r, 5).value = "Overview + side-tab 20-year history"
        ws.Cells(r, 6).value = "https://www.cookcountytreasurer.com/setsearchparameters.aspx"
        ws.Cells(r, 7).value = "No"
        ws.Cells(r, 8).value = "Direct HTTP + optional Safari browser assist (Mac)"
    Else
        ws.Cells(r, 5).value = "Overview + side-tab 20-year history"
        ws.Cells(r, 8).value = "Direct HTTP + optional Safari browser assist (Mac)"
    End If

    ' Config entries.
    Set ws = ThisWorkbook.Worksheets("Config")
    found = False
    For r = 5 To 45
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Treasurer Overview Results URL", vbTextCompare) = 0 Then
            found = True
            ws.Cells(r, 2).value = "https://www.cookcountytreasurer.com/yourpropertytaxoverviewresults.aspx"
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "Treasurer Overview Results URL"
        ws.Cells(r, 2).value = "https://www.cookcountytreasurer.com/yourpropertytaxoverviewresults.aspx"
        ws.Cells(r, 3).value = "Cook County Treasurer"
        ws.Cells(r, 4).value = "Session-sensitive property overview page."
    End If

    found = False
    For r = 5 To 45
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Treasurer 20-Year History URL", vbTextCompare) = 0 Then
            found = True
            ws.Cells(r, 2).value = "https://www.cookcountytreasurer.com/taxbillhistoryresults.aspx"
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "Treasurer 20-Year History URL"
        ws.Cells(r, 2).value = "https://www.cookcountytreasurer.com/taxbillhistoryresults.aspx"
        ws.Cells(r, 3).value = "Cook County Treasurer"
        ws.Cells(r, 4).value = "Side-tab result page; may require an active Treasurer property session."
    End If

    found = False
    For r = 5 To 45
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Automatic Treasurer Browser Assist", vbTextCompare) = 0 Then
            found = True
            ws.Cells(r, 2).value = "YES"
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "Automatic Treasurer Browser Assist"
        ws.Cells(r, 2).value = "YES"
        ws.Cells(r, 3).value = "YES / NO"
        ws.Cells(r, 4).value = "Mac only. If YES, Safari runs automatically only when direct Treasurer HTTP does not return both Overview and 20-Year History."
    End If
    ' Optional official appeal-website cross-check.
    Set ws = ThisWorkbook.Worksheets("Start")
    found = False
    For r = 12 To 45
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "PTAB Status Check", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        r = 24
        ws.Cells(r, 1).value = "PTAB Status Check"
        ws.Cells(r, 2).value = "Not run"
        ws.Cells(r, 5).value = "Official Illinois PTAB status check"
        ws.Cells(r, 6).value = "https://www.ptab.illinois.gov/asi/"
        ws.Cells(r, 7).value = "No"
        ws.Cells(r, 8).value = "SMART PTAB browser check (Mac)"
    End If

    Set ws = ThisWorkbook.Worksheets("Config")

    found = False
    For r = 5 To 50
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "PTAB Website Check", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "PTAB Website Check"
        ws.Cells(r, 2).value = "SMART"
        ws.Cells(r, 3).value = "SMART / ALWAYS / OFF"
        ws.Cells(r, 4).value = "SMART only opens PTAB when Assessor/BOR data already shows an appeal signal."
    End If

    found = False
    For r = 5 To 50
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Include Sources Sheet in PDF", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "Include Sources Sheet in PDF"
        ws.Cells(r, 2).value = "NO"
        ws.Cells(r, 3).value = "YES / NO"
        ws.Cells(r, 4).value = "NO speeds PDF creation and keeps the client-facing packet shorter; the Sources sheet remains in Excel."
    End If

    found = False
    For r = 5 To 50
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Skip Empty Optional PDF Sheets", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "Skip Empty Optional PDF Sheets"
        ws.Cells(r, 2).value = "YES"
        ws.Cells(r, 3).value = "YES / NO"
        ws.Cells(r, 4).value = "YES omits empty Issues, Permits, Documents, Sales and Appeals pages from the PDF."
    End If

    found = False
    For r = 5 To 50
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), "Mac Treasurer Strategy", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        r = LastUsedRowAnyColumn(ws) + 1
        ws.Cells(r, 1).value = "Mac Treasurer Strategy"
        ws.Cells(r, 2).value = "SAFARI-FIRST"
        ws.Cells(r, 3).value = "SAFARI-FIRST / HTTP-FIRST"
        ws.Cells(r, 4).value = "SAFARI-FIRST skips the usually-unsuccessful Treasurer HTTP attempts on Mac and uses the proven property-session browser helper directly."
    End If

End Sub

Private Sub InstallTaxDetailHyperlinks()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Tax Detail")

    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0

    For r = 26 To 33
        If Len(Trim$(CStr(ws.Cells(r, 8).value))) > 0 Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 8), _
                              Address:=CStr(ws.Cells(r, 8).value), _
                              TextToDisplay:="Open official Treasurer page"
        End If
    Next r
End Sub

Private Sub InstallTreasurerLookupButton_v121()
    Dim ws As Worksheet, shp As Shape, target As Range

    Set ws = ThisWorkbook.Worksheets("Tax Detail")
    Set target = ws.Range("G76:H77")

    On Error Resume Next
    ws.Shapes("btnTreasurerBrowserLookup").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 target.Left, target.Top, _
                                 target.Width, target.Height)

    With shp
        .Name = "btnTreasurerBrowserLookup"
        .OnAction = "'" & ThisWorkbook.Name & "'!RunTreasurerBrowserLookup_v121"
        .TextFrame.Characters.Text = "RE-RUN TREASURER LOOKUP"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(22, 101, 52)
        .Line.ForeColor.RGB = RGB(22, 101, 52)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 11
        .Placement = xlMoveAndSize
    End With

    ws.Range("G78").value = "Automatic fallback is ON by default. Use this button to rerun Treasurer only."
    ws.Range("G78:H78").WrapText = True
End Sub

Public Sub InstallRunButton_v121()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim target As Range

    Set ws = ThisWorkbook.Worksheets("Start")
    Set target = ws.Range("B6:D7")

    On Error Resume Next
    ws.Shapes("btnRunPropertyResearch").Delete
    ws.Range("E6:H7").UnMerge
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 target.Left, target.Top, _
                                 target.Width, target.Height)

    With shp
        .Name = "btnRunPropertyResearch"
        .OnAction = "'" & ThisWorkbook.Name & "'!GenerateCookPropertyReport_v121"
        .TextFrame.Characters.Text = "RUN PROPERTY RESEARCH"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(23, 50, 77)
        .Line.ForeColor.RGB = RGB(23, 50, 77)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 11
        .Placement = xlMoveAndSize
    End With

    ws.Range("B6:D7").ClearContents
    ws.Range("E6:H7").Merge
#If Mac Then
    ws.Range("E6").value = "Mac: click the blue button. If Status says the helper is missing, run Install Mac Helper.command once."
#Else
    ws.Range("E6").value = "Windows: click the blue button to run. No external helper is required."
#End If
    ws.Range("E6").WrapText = True
End Sub

Private Sub UpdatePlatformReadyStatus()
#If Mac Then
    If MacHelperReady() Then
        ThisWorkbook.Worksheets("Start").Range("B8").value = "Ready - Mac helper installed"
    Else
        ThisWorkbook.Worksheets("Start").Range("B8").value = "HTTP helper v1.1.3 missing - run split-helper installer once"
    End If
#Else
    ThisWorkbook.Worksheets("Start").Range("B8").value = "Ready - Windows"
#End If
End Sub

Private Function PlatformReady() As Boolean
#If Mac Then
    If MacHelperReady() Then
        PlatformReady = True
    Else
        MsgBox "The Mac HTTP helper is not installed yet." & vbCrLf & vbCrLf & _
               "Run 'Install Mac Helper v1.2.1.command', then reopen this workbook.", _
               vbExclamation, "Cook County Property Due Diligence"
        ThisWorkbook.Worksheets("Start").Range("B8").value = "HTTP helper v1.1.3 missing - run split-helper installer once"
        PlatformReady = False
    End If
#Else
    PlatformReady = True
#End If
End Function

Private Function MacHelperReady() As Boolean
#If Mac Then
    Dim result As String
    On Error GoTo Missing
    result = AppleScriptTask("CookPropertyHTTP.applescript", "helperVersion", "")
    MacHelperReady = (Trim$(result) = "1.1.3")
    Exit Function
Missing:
    Err.Clear
    MacHelperReady = False
#Else
    MacHelperReady = True
#End If
End Function

Public Sub GenerateCookPropertyReport_v121()
    Dim pin As String, pinInputAddress As String
    Dim t0 As Double
    Dim fatalNumber As Long, fatalDescription As String
    Dim priorCalc As XlCalculation
    t0 = Timer

    On Error GoTo FatalError

    ThisWorkbook.Worksheets("Start").Range("B8").value = "v1.2.1 started - checking platform"
    DoEvents

    gRunStage = "Platform readiness"
    If Not PlatformReady() Then Exit Sub

    priorCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    gRunStage = "Reading PIN"
    pin = ReadPinFromStartSheet(pinInputAddress)
    If Len(pin) <> 14 Then
        MsgBox "Enter a valid 14-digit Cook County PIN in the yellow PIN field.", vbExclamation
        GoTo SafeExit
    End If

    If Len(pinInputAddress) = 0 Then pinInputAddress = "B4"
    gRunStage = "Writing PIN"
    ThisWorkbook.Worksheets("Start").Range(pinInputAddress).value = FormatPin(pin)

    gRunStage = "Resetting workbook"
    ResetRunSheets

    gRunStage = "Starting research"
    SetRunStatus "Starting research for " & FormatPin(pin)

    ' PIN comes from user input and is valid even if a public-record adapter fails.
    SetProperty "PIN", FormatPin(pin), "User Input", Format(Date, "yyyy-mm-dd"), "OK", _
                "Normalized from the Start-sheet PIN field."

    gRunStage = "Assessor Addresses"
    SafeAdapter "Assessor Addresses", "FetchAddress", pin
    gRunStage = "Parcel Universe"
    SafeAdapter "Parcel Universe", "FetchUniverse", pin
    gRunStage = "Assessed Values"
    SafeAdapter "Assessed Values", "FetchAssessments", pin
    gRunStage = "Parcel Sales"
    SafeAdapter "Parcel Sales", "FetchSales", pin
    gRunStage = "Assessor Appeals"
    SafeAdapter "Assessor Appeals", "FetchAssessorAppeals", pin
    gRunStage = "Board of Review"
    SafeAdapter "Board of Review", "FetchBORAppeals", pin
    gRunStage = "PTAB Status Check"
    SafeAdapter "PTAB Status Check", "FetchPTABStatus", pin
    gRunStage = "Permits"
    SafeAdapter "Permits", "FetchPermits", pin
    gRunStage = "Cook County GIS"
    SafeAdapter "Cook County GIS", "FetchGISParcel", pin
    gRunStage = "Property Tax Portal"
    SafeAdapter "Property Tax Portal", "FetchTaxPortal", pin
    gRunStage = "Cook County Treasurer"
    SafeAdapter "Cook County Treasurer", "FetchTreasurerOverview", pin
    gRunStage = "TIF GIS"
    SafeAdapter "TIF GIS", "FetchTIF", pin
    gRunStage = "Enterprise Zone"
    SafeAdapter "Enterprise Zone", "FetchEnterpriseZone", pin

    gRunStage = "Building Issues"
    BuildIssues
    gRunStage = "Building Report"
    BuildReport
    Application.Calculate
    gRunStage = "Exporting PDF"
    ExportDueDiligencePDF pin

    SetRunStatus "Complete in " & Format(Timer - t0, "0.0") & " seconds"

SafeExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    On Error Resume Next
    Application.Calculation = priorCalc
    On Error GoTo 0
    Exit Sub

FatalError:
    fatalNumber = Err.Number
    fatalDescription = Err.Description

    On Error Resume Next
    SetRunStatus "Fatal error at " & gRunStage & ": " & fatalDescription
    On Error GoTo 0

    MsgBox "The run stopped at: " & gRunStage & vbCrLf & vbCrLf & _
           "Error " & CStr(fatalNumber) & ": " & fatalDescription, _
           vbCritical, "Cook County Property Due Diligence v1.2.1"
    Resume SafeExit
End Sub

Private Sub SafeAdapter(ByVal sourceName As String, ByVal procName As String, ByVal pin As String)
    On Error GoTo EH

    gAdapterOutcome = ""
    gAdapterNote = ""

    UpdateSourceStatus sourceName, "Running", "Starting", ""

    Select Case procName
        Case "FetchAddress": FetchAddress pin
        Case "FetchUniverse": FetchUniverse pin
        Case "FetchAssessments": FetchAssessments pin
        Case "FetchSales": FetchSales pin
        Case "FetchAssessorAppeals": FetchAssessorAppeals pin
        Case "FetchBORAppeals": FetchBORAppeals pin
        Case "FetchPTABStatus": FetchPTABStatus pin
        Case "FetchPermits": FetchPermits pin
        Case "FetchGISParcel": FetchGISParcel pin
        Case "FetchTaxPortal": FetchTaxPortal pin
        Case "FetchTreasurerOverview": FetchTreasurerOverview pin
        Case "FetchTIF": FetchTIF pin
        Case "FetchEnterpriseZone": FetchEnterpriseZone pin
    End Select

    If StrComp(sourceName, "Enterprise Zone", vbTextCompare) = 0 Then
        UpdateSourceStatus sourceName, "PARTIAL", _
                           "Finished - boundary verification remains manual", _
                           "Cook County Assessor signal collected automatically; final DCEO boundary verification remains manual."
    ElseIf StrComp(gAdapterOutcome, "PARTIAL", vbTextCompare) = 0 Then
        UpdateSourceStatus sourceName, "PARTIAL", "Finished - manual follow-up retained", gAdapterNote
    ElseIf StrComp(gAdapterOutcome, "SKIPPED", vbTextCompare) = 0 Then
        UpdateSourceStatus sourceName, "SKIPPED", "Skipped by SMART mode", gAdapterNote
    ElseIf StrComp(gAdapterOutcome, "NO RECORDS", vbTextCompare) = 0 Then
        UpdateSourceStatus sourceName, "NO RECORDS", "Finished - no rows returned", gAdapterNote
    Else
        UpdateSourceStatus sourceName, "OK", "Finished", gAdapterNote
    End If
    Exit Sub

EH:
    UpdateSourceStatus sourceName, "FAILED", "Stopped", Err.Description
    AppendIssue "HIGH", sourceName & " could not be verified automatically.", sourceName, _
                "Open the official source link on the Start/Sources sheet and verify manually. Error: " & Err.Description
    Err.Clear
End Sub

Private Sub MarkAdapterNoRecords(ByVal noteText As String)
    gAdapterOutcome = "NO RECORDS"
    gAdapterNote = noteText
End Sub

Private Sub MarkAdapterPartial(ByVal noteText As String)
    gAdapterOutcome = "PARTIAL"
    gAdapterNote = noteText
End Sub

Private Sub MarkAdapterSkipped(ByVal noteText As String)
    gAdapterOutcome = "SKIPPED"
    gAdapterNote = noteText
End Sub

Private Sub FetchAddress(ByVal pin As String)
    Dim url As String, csv As String, rows As Collection, d As Collection
    url = SocrataCsv(ADDRESS_DATASET, pin, "$order=year DESC&$limit=1")
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        DoEvents
        csv = HttpGet(url)
        Set rows = ParseCsv(csv)
    End If

    If rows.count < 2 Then Err.Raise vbObjectError + 101, , _
        "No address record returned after two live attempts."
    Set d = CsvRowDict(rows(1), rows(2))

    SetProperty "PIN", FormatPin(pin), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Property Address", DictGet(d, "prop_address_full"), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Property City / ZIP", Trim$(Nz(DictGet(d, "prop_address_city_name")) & " " & Nz(DictGet(d, "prop_address_zipcode_1"))), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Owner Name", DictGet(d, "owner_address_name"), "Assessor Addresses", Nz(DictGet(d, "year")), "CAUTION", "Owner/mailing data may be intermittently updated."
    SetProperty "Taxpayer / Mailing Name", DictGet(d, "mail_address_name"), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Mailing Address", Trim$(Nz(DictGet(d, "mail_address_full")) & ", " & Nz(DictGet(d, "mail_address_city_name")) & ", " & Nz(DictGet(d, "mail_address_state")) & " " & Nz(DictGet(d, "mail_address_zipcode_1"))), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
End Sub

Private Sub FetchUniverse(ByVal pin As String)
    Dim url As String, csv As String, rows As Collection, d As Collection
    Dim ezNum As String, ezYear As String
    Dim universeSource As String

    url = SocrataCsv(UNIVERSE_DATASET, pin, "$limit=1")
    universeSource = "Parcel Universe Current"
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        DoEvents
        csv = HttpGet(url)
        Set rows = ParseCsv(csv)
    End If

    If rows.count < 2 Then
        universeSource = "Parcel Universe Historical fallback"
        url = SocrataCsv("nj4t-kc8j", pin, "$order=year DESC&$limit=1")
        csv = HttpGet(url)
        Set rows = ParseCsv(csv)
    End If

    If rows.count < 2 Then Err.Raise vbObjectError + 102, , _
        "No Parcel Universe row returned from current view or historical fallback."

    Set d = CsvRowDict(rows(1), rows(2))

    SetProperty "Township", DictGet(d, "township_name"), universeSource, Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Assessor Class", DictGet(d, "class"), universeSource, Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Assessor Neighborhood", DictGet(d, "nbhd_code"), universeSource, Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Tax Code", DictGet(d, "tax_code"), universeSource, Nz(DictGet(d, "year")), "CAUTION", "County metadata states this tax-code field is not currently up-to-date; Property Tax Portal takes precedence."
    SetProperty "Longitude", DictGet(d, "lon"), universeSource, Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Latitude", DictGet(d, "lat"), universeSource, Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Centroid X (3435)", DictGet(d, "x_3435"), universeSource, Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Centroid Y (3435)", DictGet(d, "y_3435"), universeSource, Nz(DictGet(d, "year")), "OK", ""

    ezNum = Trim$(DictGet(d, "econ_enterprise_zone_num"))
    ezYear = Trim$(DictGet(d, "econ_enterprise_zone_data_year"))

    If Len(ezNum) > 0 And ezNum <> "0" Then
        SetProperty "Enterprise Zone (Assessor Signal)", _
                    "Signal present: Zone " & ezNum, _
                    universeSource, FirstText(ezYear, Nz(DictGet(d, "year"))), "CHECK", _
                    "Cook County Assessor spatial signal only; verify the parcel in the official Illinois DCEO Enterprise Zone map before reliance."
    Else
        SetProperty "Enterprise Zone (Assessor Signal)", _
                    "No county signal returned", _
                    universeSource, FirstText(ezYear, Nz(DictGet(d, "year"))), "CHECK", _
                    "Absence of the county signal is not treated as a conclusive outside-zone determination; verify in the official Illinois DCEO map."
    End If

End Sub

Private Sub FetchAssessments(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection

    n = CLng(GetConfigValue("Assessment Years", 8))
    url = SocrataCsv(ASSESSMENT_DATASET, pin, "$order=year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        Err.Raise vbObjectError + 103, , _
            "No assessed-value rows returned for a valid parcel."
    End If

    WriteSelectedCsv "Assessment", csv, 5, _
        Array("year", "class", "mailed_land", "mailed_bldg", "mailed_tot", "certified_land", "certified_bldg", "certified_tot", "board_land", "board_bldg", "board_tot")
End Sub

Private Sub FetchSales(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection

    n = CLng(GetConfigValue("Sales Limit", 20))
    url = SocrataCsv(SALES_DATASET, pin, "$order=sale_date DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        MarkAdapterNoRecords "No Assessor parcel-sale rows returned for this PIN."
        Exit Sub
    End If

    WriteMappedTable "Sales-Deeds", csv, 5, _
        Array("sale_date", "sale_price", "doc_no", "deed_type", "mydec_deed_type", "buyer_name", "seller_name", "is_multisale", "num_parcels_sale"), _
        Array("Sale Date", "Sale Price", "Document No.", "Deed Type", "MyDec Deed Type", "Buyer", "Seller", "Multi-PIN?", "# Parcels")
End Sub

Private Sub FetchAssessorAppeals(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection, i As Long, d As Collection

    n = CLng(GetConfigValue("Appeals Limit", 30))
    url = SocrataCsv(APPEALS_DATASET, pin, "$order=year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        MarkAdapterNoRecords "No Assessor appeal rows returned for this PIN."
        Exit Sub
    End If

    For i = 2 To rows.count
        Set d = CsvRowDict(rows(1), rows(i))
        AppendAppeal "Assessor", DictGet(d, "year"), DictGet(d, "case_no"), _
                     FirstNonBlank(d, Array("appeal_type", "hearing_type")), DictGet(d, "status"), _
                     DictGet(d, "mailed_tot"), DictGet(d, "certified_tot"), "Assessor Appeals"
    Next i
End Sub

Private Sub FetchBORAppeals(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection, i As Long, d As Collection

    n = CLng(GetConfigValue("BOR Limit", 30))
    url = SocrataCsv(BOR_DATASET, pin, "$order=tax_year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        MarkAdapterNoRecords "No Board of Review decision rows returned for this PIN."
        Exit Sub
    End If

    For i = 2 To rows.count
        Set d = CsvRowDict(rows(1), rows(i))
        AppendAppeal "Board of Review", DictGet(d, "tax_year"), _
                     Nz(DictGet(d, "appealtrk")) & "-" & Nz(DictGet(d, "appealseq")), _
                     FirstNonBlank(d, Array("appealtypedescription", "appealtype")), DictGet(d, "result"), _
                     DictGet(d, "assessor_totalvalue"), DictGet(d, "bor_totalvalue"), "BOR Decision History"
    Next i
End Sub

Private Sub FetchPTABStatus(ByVal pin As String)
    Dim mode As String
#If Mac Then
    Dim payload As String
#End If

    mode = UCase$(Trim$(CStr(GetConfigValue("PTAB Website Check", "SMART"))))

    If mode = "OFF" Then
        MarkAdapterSkipped "PTAB website check is OFF in Config."
        Exit Sub
    End If

    If mode = "SMART" And Not AppealsHaveSignal() Then
        MarkAdapterSkipped "No Assessor/BOR appeal signal was found, so the PTAB lookup was skipped to keep the cold run fast."
        Exit Sub
    End If

#If Mac Then
    On Error GoTo BrowserFailed

    payload = AppleScriptTask("CookPropertyPTABBrowser.applescript", "lookupPTAB", FormatPin(pin))

    If Left$(payload, 10) = "__ERROR__|" Then
        MarkAdapterPartial "Official PTAB lookup could not complete: " & Mid$(payload, 11)
        Exit Sub
    End If

    ParsePTABDocketLinks pin, payload

    If Len(Trim$(gAdapterNote)) = 0 Then
        gAdapterNote = "Illinois PTAB Appeal Status Inquiry checked in SMART browser mode."
    End If
    Exit Sub

BrowserFailed:
    MarkAdapterPartial "Official PTAB lookup failed: " & Err.Description
    Err.Clear
#Else
    MarkAdapterPartial "PTAB live website cross-check is browser-assisted on Mac in v1.2.1. Existing Cook County Assessor/BOR history remains available on Windows."
#End If
End Sub

Private Function AppealsHaveSignal() As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Appeals")
    lastRow = LastUsedRow(ws, 1)

    For r = 5 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, 1).value))) > 0 Then
            AppealsHaveSignal = True
            Exit Function
        End If
    Next r
End Function

Private Sub ParsePTABDocketLinks(ByVal pin As String, ByVal linksText As String)
    Dim lines() As String, i As Long, lineText As String
    Dim sepPos As Long, docketNo As String, docketUrl As String
    Dim foundAny As Boolean, detailHtml As String, detailText As String
    Dim statusText As String, taxYear As String, count As Long

    linksText = Replace(linksText, vbCrLf, vbLf)
    linksText = Replace(linksText, vbCr, vbLf)
    lines = Split(linksText, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) > 0 Then
            sepPos = InStr(1, lineText, vbTab, vbBinaryCompare)
            If sepPos > 0 Then
                docketNo = Trim$(Left$(lineText, sepPos - 1))
                docketUrl = Trim$(Mid$(lineText, sepPos + 1))
            Else
                docketNo = Trim$(lineText)
                docketUrl = "https://www.ptab.illinois.gov/asi/property.asp?DocketNo=" & UrlEncode(docketNo)
            End If

            If LooksLikePTABDocket(docketNo) Then
                foundAny = True
                count = count + 1
                If count > 5 Then Exit For

                detailHtml = HttpGet(docketUrl)
                detailText = CollapseWhitespace(HtmlToText(detailHtml))
                statusText = PTABStatusSummary(detailText)
                taxYear = PTABDocketTaxYear(docketNo, detailText)

                AppendAppeal "PTAB", taxYear, docketNo, "Appeal Status Inquiry", _
                             statusText, "", "", _
                             "Illinois PTAB ASI | " & docketUrl
            End If
        End If
    Next i

    If Not foundAny Then
        AppendAppeal "PTAB", "", "", "Appeal Status Inquiry", _
                     "No PTAB docket found by PIN", "", "", _
                     "Illinois PTAB Appeal Status Inquiry checked by PIN."
    End If
End Sub

Private Function LooksLikePTABDocket(ByVal docketNo As String) As Boolean
    docketNo = Trim$(docketNo)
    If Len(docketNo) < 7 Then Exit Function

    If Len(docketNo) >= 8 And IsNumeric(Left$(docketNo, 2)) And Mid$(docketNo, 3, 1) = "-" Then
        LooksLikePTABDocket = True
    ElseIf Len(docketNo) >= 10 And IsNumeric(Left$(docketNo, 4)) And Mid$(docketNo, 5, 1) = "-" Then
        LooksLikePTABDocket = True
    End If
End Function

Private Function PTABDocketTaxYear(ByVal docketNo As String, ByVal detailText As String) As String
    Dim p As Long, tailText As String

    p = InStr(1, detailText, "Information for Docket No:", vbTextCompare)
    If p > 0 Then
        tailText = Trim$(Mid$(detailText, p + Len("Information for Docket No:")))
        If Len(tailText) >= 4 And IsNumeric(Left$(tailText, 4)) Then
            PTABDocketTaxYear = Left$(tailText, 4)
            Exit Function
        End If
    End If

    docketNo = Trim$(docketNo)
    If Len(docketNo) >= 10 And IsNumeric(Left$(docketNo, 4)) And Mid$(docketNo, 5, 1) = "-" Then
        PTABDocketTaxYear = Left$(docketNo, 4)
    ElseIf Len(docketNo) >= 8 And IsNumeric(Left$(docketNo, 2)) And Mid$(docketNo, 3, 1) = "-" Then
        PTABDocketTaxYear = "20" & Left$(docketNo, 2)
    End If
End Function

Private Function PTABStatusSummary(ByVal txt As String) As String
    If InStr(1, txt, "Case closed on", vbTextCompare) > 0 Then
        PTABStatusSummary = "Closed"
    ElseIf InStr(1, txt, "Hearing Scheduled", vbTextCompare) > 0 Or _
           InStr(1, txt, "Hearing Set", vbTextCompare) > 0 Then
        PTABStatusSummary = "Hearing Scheduled"
    ElseIf InStr(1, txt, "All evidence has been submitted", vbTextCompare) > 0 Then
        PTABStatusSummary = "Evidence Complete / Decision Pending"
    ElseIf InStr(1, txt, "Evidence from County received", vbTextCompare) > 0 Then
        PTABStatusSummary = "County Evidence Received"
    ElseIf InStr(1, txt, "Appellant filing is complete", vbTextCompare) > 0 Then
        PTABStatusSummary = "Open - Filing Complete"
    ElseIf InStr(1, txt, "Appellant new appeal received", vbTextCompare) > 0 Then
        PTABStatusSummary = "Open - Appeal Received"
    Else
        PTABStatusSummary = "Open / See ASI Detail"
    End If
End Function

Private Sub FetchPermits(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection

    n = CLng(GetConfigValue("Permits Limit", 50))
    url = SocrataCsv(PERMITS_DATASET, pin, "$order=date_issued DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)

    If rows.count < 2 Then
        MarkAdapterNoRecords "No Assessor permit rows returned for this PIN."
        Exit Sub
    End If

    WriteMappedTable "Permits", csv, 5, _
        Array("date_issued", "year", "local_permit_number", "permit_number", "status", "assessable", "amount", "municipality", "applicant_name", "work_description"), _
        Array("Issued Date", "Year", "Local Permit No.", "CCAO Permit No.", "Status", "Assessable?", "Amount", "Municipality", "Applicant", "Work Description")
End Sub

Private Sub FetchGISParcel(ByVal pin As String)
    Dim baseUrl As String, whereClause As String, url As String, js As String
    Dim pinHit As String, taxYear As String, taxCode As String, x As String, y As String
    Dim muniTax As String, muniSpatial As String, muniUrl As String, muniJs As String

    baseUrl = CStr(GetConfigValue("GIS Current Parcel URL", _
        "https://gis.cookcountyil.gov/traditional/rest/services/CookViewer3Parcels/MapServer/0"))
    whereClause = "PIN14='" & pin & "' OR PIN14_dash='" & FormatPin(pin) & "'"

    url = baseUrl & "/query?where=" & UrlEncode(whereClause) & _
          "&outFields=PIN14,PIN14_dash,TAXYR,TAXDIST,XCOORD,YCOORD,street_address,township_name,latitude,longitude," & _
          "LANDSF,CURRENTVALUE_TOTAL,CURRENTVALUE_LAND,CURRENTVALUE_BLDG,BLDGSQFT,bldg_const_desc,BCLASS," & _
          "major_class_description,class_description,NBHD,BLDGAGE,tax_municipality_name" & _
          "&returnGeometry=false&f=pjson"
    js = HttpGet(url)
    EnsureArcGisResponse js, "CookViewer current parcel layer"

    If ArcGisFeatureCount(js) = 0 Then
        Err.Raise vbObjectError + 151, , "CookViewer current parcel layer did not return the PIN."
    End If

    pinHit = JsonScalar(js, "PIN14")
    taxYear = JsonScalar(js, "TAXYR")
    taxCode = JsonScalar(js, "TAXDIST")
    x = JsonScalar(js, "XCOORD")
    y = JsonScalar(js, "YCOORD")
    muniTax = JsonScalar(js, "tax_municipality_name")

    SetProperty "GIS PIN Match", pinHit, "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "GIS Building Class", JsonScalar(js, "BCLASS"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "GIS Neighborhood", JsonScalar(js, "NBHD"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "GIS Class Description", JsonScalar(js, "class_description"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "GIS Major Class", JsonScalar(js, "major_class_description"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "Lot Size (SqFt)", JsonScalar(js, "LANDSF"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "Building (SqFt)", JsonScalar(js, "BLDGSQFT"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "Building Construction", JsonScalar(js, "bldg_const_desc"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "Building Age", JsonScalar(js, "BLDGAGE"), "CookViewer Parcels Current", taxYear, "OK", ""
    SetProperty "GIS Current Land Value", JsonScalar(js, "CURRENTVALUE_LAND"), "CookViewer Parcels Current", taxYear, "OK", "CookViewer CURRENTVALUE_LAND field."
    SetProperty "GIS Current Building Value", JsonScalar(js, "CURRENTVALUE_BLDG"), "CookViewer Parcels Current", taxYear, "OK", "CookViewer CURRENTVALUE_BLDG field."
    SetProperty "GIS Current Total Value", JsonScalar(js, "CURRENTVALUE_TOTAL"), "CookViewer Parcels Current", taxYear, "OK", "CookViewer CURRENTVALUE_TOTAL field."
    SetProperty "Municipality - Tax Record", muniTax, "CookViewer Parcels Current", taxYear, "CHECK", "Tax-record municipality; compare with current spatial boundary."

    If Len(GetPropertyValue("Property Address")) = 0 Then
        SetProperty "Property Address", JsonScalar(js, "street_address"), _
                    "CookViewer Parcels Current", taxYear, "CHECK", _
                    "GIS fallback because Assessor Parcel Addresses did not populate."
    End If

    If Len(GetPropertyValue("Township")) = 0 Then
        SetProperty "Township", JsonScalar(js, "township_name"), _
                    "CookViewer Parcels Current", taxYear, "CHECK", _
                    "GIS fallback because Parcel Universe did not populate."
    End If

    If Len(GetPropertyValue("Longitude")) = 0 Then SetProperty "Longitude", JsonScalar(js, "longitude"), "CookViewer Parcels Current", taxYear, "OK", ""
    If Len(GetPropertyValue("Latitude")) = 0 Then SetProperty "Latitude", JsonScalar(js, "latitude"), "CookViewer Parcels Current", taxYear, "OK", ""
    If Len(GetPropertyValue("Centroid X (3435)")) = 0 And Len(x) > 0 Then SetProperty "Centroid X (3435)", x, "CookViewer Parcels Current", taxYear, "OK", ""
    If Len(GetPropertyValue("Centroid Y (3435)")) = 0 And Len(y) > 0 Then SetProperty "Centroid Y (3435)", y, "CookViewer Parcels Current", taxYear, "OK", ""
    If Len(taxCode) > 0 Then SetProperty "Tax Code", taxCode, "CookViewer Parcels Current", taxYear, "CAUTION", "Property Tax Portal takes precedence for current tax-rate context."

    x = GetPropertyValue("Centroid X (3435)")
    y = GetPropertyValue("Centroid Y (3435)")
    If Len(x) > 0 And Len(y) > 0 Then
        On Error Resume Next
        muniUrl = CStr(GetConfigValue("Spatial Municipality URL", _
            "https://gis.cookcountyil.gov/traditional/rest/services/politicalBoundary/MapServer/2"))
        url = muniUrl & "/query?geometry=" & UrlEncode(x & "," & y) & _
              "&geometryType=esriGeometryPoint&inSR=3435&spatialRel=esriSpatialRelIntersects" & _
              "&outFields=MUNICIPALITY,AGENCY_DESC&returnGeometry=false&f=pjson"
        muniJs = HttpGet(url)
        EnsureArcGisResponse muniJs, "Cook County municipality boundary"
        If Err.Number = 0 Then
            If ArcGisFeatureCount(muniJs) > 0 Then
                muniSpatial = JsonScalar(muniJs, "MUNICIPALITY")
                If Len(muniSpatial) = 0 Then muniSpatial = JsonScalar(muniJs, "AGENCY_DESC")
                SetProperty "Municipality - Spatial", muniSpatial, "Cook County Municipality Boundary", "Current", "OK", "GIS-maintained municipal boundary."
            Else
                AppendIssue "MEDIUM", "No municipality polygon intersected the parcel centroid.", _
                            "Cook County Municipality Boundary", "Verify whether the parcel is in unincorporated Cook County or near a municipal boundary."
            End If
        Else
            AppendIssue "MEDIUM", "Current spatial municipality could not be verified automatically.", _
                        "Cook County Municipality Boundary", "Verify municipality in CookViewer / Cook County GIS. Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub FetchTaxPortal(ByVal pin As String)
    Dim url As String, html As String, txt As String
    Dim rate As String, taxCode As String, rateBlock As String, billBlock As String, exBlock As String
    Dim taxSaleBlock As String, refundBlock As String
    Dim portalClass As String, portalClassDesc As String, portalLot As String, portalBldg As String
    Dim r As Long, yr As String, exText As String
    Dim portalBase As String, portalFallback As String
    Dim formattedPin As String

    formattedPin = FormatPin(pin)
    portalBase = CStr(GetConfigValue("Property Tax Portal Results URL", "https://www.cookcountypropertyinfo.com/PINResults.aspx"))
    portalFallback = "https://www.cookcountypropertyinfo.com/cookviewerpinresults.aspx"

    ' The Portal can redirect a plain direct request back to its search page.
    ' Try the human-formatted PIN first, then the alternate official results route,
    ' then the unformatted legacy variants.
    url = portalBase & "?pin=" & UrlEncode(formattedPin)
    html = HttpGetPortal(url)
    txt = CollapseWhitespace(HtmlToText(html))

    If Not PortalPageLooksValid(txt, pin) Then
        url = portalFallback & "?pin=" & UrlEncode(formattedPin)
        html = HttpGetPortal(url)
        txt = CollapseWhitespace(HtmlToText(html))
    End If

    If Not PortalPageLooksValid(txt, pin) Then
        url = portalBase & "?pin=" & pin
        html = HttpGetPortal(url)
        txt = CollapseWhitespace(HtmlToText(html))
    End If

    If Not PortalPageLooksValid(txt, pin) Then
        url = portalFallback & "?pin=" & pin
        html = HttpGetPortal(url)
        txt = CollapseWhitespace(HtmlToText(html))
    End If

    If Not PortalPageLooksValid(txt, pin) Then
        Err.Raise vbObjectError + 160, , _
            "Property Tax Portal returned its search/generic page instead of PIN-specific results for " & formattedPin & "."
    End If

    rate = ValueAfterLabel(txt, "Tax Rate", "0123456789.")
    taxCode = ValueAfterLabel(txt, "Tax Code", "0123456789")
    portalClass = ValueAfterLabel(txt, "Property Class", "0123456789-")
    portalLot = ValueAfterLabel(txt, "Lot Size (SqFt)", "0123456789,.")
    portalBldg = ValueAfterLabel(txt, "Building (SqFt)", "0123456789,.")
    portalClassDesc = CleanPortalText(TextBetween(txt, "Property Class Description", "Tax Rate"))

    If Len(rate) > 0 Then
        ThisWorkbook.Worksheets("Report").Range("E7").value = rate
        ThisWorkbook.Worksheets("Tax Detail").Range("B4").value = rate
    End If

    If Len(taxCode) > 0 Then
        SetProperty "Tax Code", taxCode, "Property Tax Portal", "Live", "OK", "Portal tax code used for current tax-rate context."
        ThisWorkbook.Worksheets("Tax Detail").Range("D4").NumberFormat = "@"
        ThisWorkbook.Worksheets("Tax Detail").Range("D4").value = taxCode
    End If

    If Len(portalClass) > 0 And Len(GetPropertyValue("Assessor Class")) > 0 Then
        If NormalizeCompare(portalClass) <> NormalizeCompare(GetPropertyValue("Assessor Class")) Then
            AppendIssue "MEDIUM", "Property Tax Portal class differs from Parcel Universe class.", "Property Tax Portal / Parcel Universe", _
                        "Portal: " & portalClass & "; Parcel Universe: " & GetPropertyValue("Assessor Class") & ". Verify current classification."
        End If
    End If
    If Len(portalClassDesc) > 0 And Len(GetPropertyValue("GIS Class Description")) = 0 Then SetProperty "GIS Class Description", portalClassDesc, "Property Tax Portal", "Live", "CHECK", "Portal class description used because GIS description was blank."
    CrossCheckPortalSize "Lot Size (SqFt)", portalLot
    CrossCheckPortalSize "Building (SqFt)", portalBldg

    rateBlock = PortalRateHistoryBlock(txt)
    ParsePortalRateHistory rateBlock, taxCode

    billBlock = TextBetween(txt, "TAX BILLED AMOUNTS & TAX HISTORY", "EXEMPTIONS")
    If Len(billBlock) = 0 Then billBlock = TextBetween(txt, "TAX BILLED AMOUNTS", "EXEMPTIONS")
    If Len(billBlock) = 0 Then billBlock = TextBetween(txt, "TAX BILLED AMOUNTS", "APPEALS")
    ParsePortalBilledHistory billBlock, taxCode

    exBlock = TextBetweenAfter(txt, "TAX BILLED AMOUNTS", "EXEMPTIONS", "APPEALS")
    If Len(exBlock) > 0 Then
        For r = 5 To LastUsedRow(ThisWorkbook.Worksheets("Tax History"), 1)
            yr = Trim$(CStr(ThisWorkbook.Worksheets("Tax History").Cells(r, 1).value))
            If Len(yr) = 4 Then
                exText = PortalExemptionForYear(exBlock, yr)
                If Len(exText) > 0 Then
                    UpsertTaxHistory yr, "", "", "", taxCode, exText
                    UpsertTaxDetailRow yr, "", taxCode, "", "", exText, ""
                End If
            End If
        Next r
    End If

    refundBlock = TextBetween(txt, "REFUNDS AVAILABLE", "TAX SALE (DELINQUENCIES)")
    If Len(refundBlock) > 0 Then
        If InStr(1, refundBlock, "No Refund Available", vbTextCompare) > 0 Then
            ThisWorkbook.Worksheets("Tax Detail").Range("H4").value = "No Refund Available"
        Else
            ThisWorkbook.Worksheets("Tax Detail").Range("H4").value = CleanPortalText(refundBlock)
            AppendIssue "LOW", "Property Tax Portal does not show the standard 'No Refund Available' message.", _
                        "Cook County Property Tax Portal", _
                        "Review the Portal refund section; an overpayment/refund may exist or the Portal wording may have changed."
        End If
    End If
    taxSaleBlock = TextBetween(txt, "TAX SALE (DELINQUENCIES)", "DOCUMENTS, DEEDS & LIENS")
    If Len(taxSaleBlock) > 0 Then ParsePortalTaxSaleStatuses taxSaleBlock
    ParsePortalDocuments txt
End Sub

Private Sub ParsePortalRateHistory(ByVal rateBlock As String, ByVal taxCode As String)
    Dim y As Long, rateText As String

    If Len(rateBlock) = 0 Then Exit Sub

    For y = Year(Date) + 1 To 1999 Step -1
        rateText = LooseNumberAfterYear(rateBlock, CStr(y))

        If Len(rateText) > 0 Then
            UpsertTaxHistory CStr(y), "", "", rateText, taxCode, ""
            UpsertTaxDetailRow CStr(y), rateText, taxCode, "", "", "", ""
        End If
    Next y
End Sub

Private Sub ParsePortalBilledHistory(ByVal billBlock As String, ByVal taxCode As String)
    Dim y As Long, seg As String, billed As String, statusText As String, payOnline As String
    Dim firstBill As Boolean, isInstallment As Boolean
    If Len(billBlock) = 0 Then Exit Sub
    firstBill = True
    For y = Year(Date) + 1 To 1999 Step -1
        If InStr(1, billBlock, CStr(y) & ":", vbTextCompare) > 0 Then
            seg = PortalYearSegment(billBlock, CStr(y))
            billed = FirstNumericToken(seg)
            If Len(billed) > 0 Then
                isInstallment = (InStr(1, Left$(seg, 80), "*", vbBinaryCompare) > 0)
                statusText = ""
                If InStr(1, seg, "Paid in Full", vbTextCompare) > 0 Then
                    statusText = "Paid in Full"
                ElseIf InStr(1, seg, "Pay Online:", vbTextCompare) > 0 Then
                    payOnline = ValueAfterLabel(seg, "Pay Online", "$0123456789,.")
                    If Len(payOnline) > 0 Then statusText = "Pay Online: " & payOnline
                ElseIf InStr(1, seg, "Payment History", vbTextCompare) > 0 Then
                    statusText = "Payment History"
                ElseIf isInstallment Then
                    statusText = "Partial-year / installment amount shown"
                End If
                UpsertTaxHistory CStr(y), billed, statusText, "", taxCode, ""
                UpsertTaxDetailRow CStr(y), "", taxCode, billed, statusText, "", ""
                If firstBill Then
                    If isInstallment Then
                        ThisWorkbook.Worksheets("Report").Range("E9").value = "$" & billed & " (1st installment)"
                        ThisWorkbook.Worksheets("Tax Detail").Range("F4").value = "$" & billed & " (1st installment)"
                    Else
                        ThisWorkbook.Worksheets("Report").Range("E9").value = "$" & billed
                        ThisWorkbook.Worksheets("Tax Detail").Range("F4").value = "$" & billed
                    End If
                    firstBill = False
                End If
            End If
        End If
    Next y
End Sub

Private Sub FetchTreasurerOverview(ByVal pin As String)
    Dim overviewUrl As String, historyUrl As String
    Dim html As String, txt As String, histHtml As String, histTxt As String
    Dim formattedPin As String
    Dim overviewOK As Boolean, historyOK As Boolean
    Dim refundBlock As String, historyBlock As String, debtBlock As String
    Dim propertyLoc As String, mailingInfo As String
    Dim totalDue As String, debtAmt As String, propValue As String, debtPct As String
    Dim firstYear As String, firstAmount As String, lastYear As String, lastAmount As String
    Dim diffAmount As String, pctChange As String

    formattedPin = FormatPin(pin)
    overviewUrl = CStr(GetConfigValue("Treasurer Overview Results URL", _
        "https://www.cookcountytreasurer.com/yourpropertytaxoverviewresults.aspx"))
    historyUrl = CStr(GetConfigValue("Treasurer 20-Year History URL", _
        "https://www.cookcountytreasurer.com/taxbillhistoryresults.aspx"))

#If Mac Then
    If UCase$(Trim$(CStr(GetConfigValue("Mac Treasurer Strategy", "SAFARI-FIRST")))) = "SAFARI-FIRST" And _
       UCase$(Trim$(CStr(GetConfigValue("Automatic Treasurer Browser Assist", "YES")))) <> "NO" Then

        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "SAFARI-FIRST: using the proven Treasurer property-session helper directly to avoid failed HTTP probes."
        TreasurerBrowserAssist pin, txt, histTxt, overviewOK, historyOK
    Else
        txt = TreasurerTryPinPage(overviewUrl, pin)
        overviewOK = TreasurerOverviewLooksValid(txt, pin)
        histTxt = TreasurerTryPinPage(historyUrl, pin)
        historyOK = TreasurerHistoryPageLooksValid(histTxt, pin)
    End If
#Else
    ' Real-world testing on Mac found direct HTTP probes to the Treasurer site
    ' essentially never succeed on their own (see the SAFARI-FIRST comment
    ' above) - the site appears to require an interactive browser session.
    ' Windows applies the same lesson: go straight to the IE-based browser
    ' assist rather than spend time on a probe that is unlikely to work,
    ' unless the user has explicitly turned automatic browser assist off.
    If UCase$(Trim$(CStr(GetConfigValue("Windows Treasurer Strategy", "IE-FIRST")))) = "IE-FIRST" And _
       UCase$(Trim$(CStr(GetConfigValue("Automatic Treasurer Browser Assist", "YES")))) <> "NO" Then

        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "IE-FIRST: using the Treasurer property-session helper directly to avoid failed HTTP probes."
        TreasurerBrowserAssist pin, txt, histTxt, overviewOK, historyOK
    Else
        txt = TreasurerTryPinPage(overviewUrl, pin)
        overviewOK = TreasurerOverviewLooksValid(txt, pin)
        histTxt = TreasurerTryPinPage(historyUrl, pin)
        historyOK = TreasurerHistoryPageLooksValid(histTxt, pin)
    End If
#End If

#If Mac Then
    ' Direct HTTP is fastest. If it is incomplete, make one isolated Safari
    ' attempt. TreasurerBrowserAssist catches helper/browser errors internally,
    ' so this automatic path never interrupts the full report with a dialog.
    If (Not overviewOK Or Not historyOK) And _
       UCase$(Trim$(CStr(GetConfigValue("Automatic Treasurer Browser Assist", "YES")))) <> "NO" And _
       UCase$(Trim$(CStr(GetConfigValue("Mac Treasurer Strategy", "SAFARI-FIRST")))) <> "SAFARI-FIRST" Then

        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Direct Treasurer HTTP was incomplete. Trying the isolated Safari Treasurer helper automatically."

        TreasurerBrowserAssist pin, txt, histTxt, overviewOK, historyOK

        If Not overviewOK And Not historyOK Then
            ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
                "Automatic Safari Treasurer lookup did not return property-specific data. Treasurer will remain PARTIAL; the report will continue."
        End If
    End If
#Else
    If (Not overviewOK Or Not historyOK) And _
       UCase$(Trim$(CStr(GetConfigValue("Automatic Treasurer Browser Assist", "YES")))) <> "NO" And _
       UCase$(Trim$(CStr(GetConfigValue("Windows Treasurer Strategy", "IE-FIRST")))) <> "IE-FIRST" Then

        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Direct Treasurer HTTP was incomplete. Trying the isolated Internet Explorer Treasurer helper automatically."

        TreasurerBrowserAssist pin, txt, histTxt, overviewOK, historyOK

        If Not overviewOK And Not historyOK Then
            ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
                "Automatic Internet Explorer Treasurer lookup did not return property-specific data. Treasurer will remain PARTIAL; the report will continue."
        End If
    End If
#End If

    If overviewOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("B18").value = "Overview automatic lookup succeeded"

        propertyLoc = CleanTreasurerText(TextBetween(txt, "Property Location:", "Volume:"))
        If Len(propertyLoc) = 0 Then propertyLoc = CleanTreasurerText(TextBetween(txt, "Property Location:", "Mailing Information:"))
        mailingInfo = CleanTreasurerText(TextBetween(txt, "Mailing Information:", "Tax Year"))

        If Len(propertyLoc) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D18").value = propertyLoc
        If Len(mailingInfo) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("H22").value = mailingInfo

        refundBlock = TextBetween(txt, "Are There Any Overpayments on Your PIN?", "Have You Received Your Exemptions")
        If Len(refundBlock) > 0 Then
            ThisWorkbook.Worksheets("Tax Detail").Range("F18").value = CleanTreasurerText(refundBlock)
            If InStr(1, refundBlock, "refund available", vbTextCompare) > 0 And _
               InStr(1, refundBlock, "do not indicate a refund available", vbTextCompare) = 0 Then
                AppendIssue "MEDIUM", "Cook County Treasurer indicates a possible property-tax refund or overpayment.", _
                            "Cook County Treasurer", _
                            "Review the Treasurer refund side tab and confirm the entitled payor before relying on the amount."
            End If
        End If

        totalDue = LastValueAfterLabel(txt, "Total Amount Due:", "$0123456789,.")
        If Len(totalDue) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("H18").value = totalDue

        ParseTreasurerInstallments txt
        ParseTreasurerExemptionGrid txt

        historyBlock = TextBetween(txt, "20-Year Property Tax Bill History", "Taxing District Debt Attributed to Your Property")
        If Len(historyBlock) > 0 Then
            ParseTreasurerTwentyYearHistory historyBlock, firstYear, firstAmount, lastYear, lastAmount

            diffAmount = ValueAfterLabel(historyBlock, "Difference", "+-$0123456789,.")
            pctChange = ValueAfterLabel(historyBlock, "Percent Change", "+-0123456789.%")

            If Len(firstYear) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("B20").value = firstYear & ": $" & firstAmount
            If Len(lastYear) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D20").value = lastYear & ": $" & lastAmount
            If Len(diffAmount) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("F20").value = diffAmount
            If Len(pctChange) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("H20").value = pctChange
        End If

        debtBlock = TextBetween(txt, "Taxing District Debt Attributed to Your Property", "Highlights of Your Taxing Districts")
        If Len(debtBlock) = 0 Then debtBlock = TextBetween(txt, "Taxing District Debt Attributed to Your Property", "Reports and Data")

        debtAmt = ValueAfterLabel(debtBlock, "Total Taxing District Debt Attributed to Your Property", "$0123456789,.")
        propValue = ValueAfterLabel(debtBlock, "Property Value", "$0123456789,.")
        debtPct = ValueAfterLabel(debtBlock, "Total Debt % Attributed to Your Property Value", "0123456789.%")

        If Len(debtAmt) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("B22").value = debtAmt
        If Len(propValue) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D22").value = propValue
        If Len(debtPct) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("F22").value = debtPct
    End If

    If historyOK Then
        ParseTreasurerTwentyYearHistory histTxt, firstYear, firstAmount, lastYear, lastAmount
        If Len(ThisWorkbook.Worksheets("Tax Detail").Range("B20").value) = 0 And Len(firstYear) > 0 Then _
            ThisWorkbook.Worksheets("Tax Detail").Range("B20").value = firstYear & ": $" & firstAmount
        If Len(ThisWorkbook.Worksheets("Tax Detail").Range("D20").value) = 0 And Len(lastYear) > 0 Then _
            ThisWorkbook.Worksheets("Tax Detail").Range("D20").value = lastYear & ": $" & lastAmount
    End If

    If overviewOK And historyOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("B18").value = "Overview + 20-year history lookup succeeded"
    ElseIf overviewOK Then
        MarkAdapterPartial "Treasurer overview loaded, but the dedicated 20-year side-tab page still required an interactive property session."
    ElseIf historyOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("B18").value = "20-year history loaded; overview requires property session"
        MarkAdapterPartial "The dedicated 20-year page loaded, but the Treasurer overview page still required an interactive property session."
    Else
        ThisWorkbook.Worksheets("Tax Detail").Range("B18").value = "PARTIAL - Treasurer requires interactive property session"
        MarkAdapterPartial "Treasurer's property pages did not accept direct PIN routes. Use the side-tab links on Tax Detail; Portal tax detail remains automatic."
    End If
End Sub

Private Function TreasurerBrowserHelperReady_v121() As Boolean
#If Mac Then
    Dim result As String
    On Error GoTo Missing

    result = AppleScriptTask("CookPropertyTreasurerBrowser.applescript", "helperVersion", "")
    TreasurerBrowserHelperReady_v121 = (Trim$(result) = "1.1.5")
    Exit Function

Missing:
    Err.Clear
    TreasurerBrowserHelperReady_v121 = False
#Else
    TreasurerBrowserHelperReady_v121 = False
#End If
End Function

Public Sub TestMacHelpers_v121()
#If Mac Then
    Dim httpResult As String, browserResult As String, appealResult As String

    On Error Resume Next
    httpResult = AppleScriptTask("CookPropertyHTTP.applescript", "helperVersion", "")
    If Err.Number <> 0 Then
        httpResult = "ERROR " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If

    browserResult = AppleScriptTask("CookPropertyTreasurerBrowser.applescript", "helperVersion", "")
    If Err.Number <> 0 Then
        browserResult = "ERROR " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    appealResult = AppleScriptTask("CookPropertyPTABBrowser.applescript", "helperVersion", "")
    If Err.Number <> 0 Then
        appealResult = "ERROR " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    MsgBox "Mac helper diagnostic" & vbCrLf & vbCrLf & _
           "HTTP helper: " & httpResult & vbCrLf & _
           "Treasurer Safari helper: " & browserResult & vbCrLf & _
           "PTAB browser helper: " & appealResult & vbCrLf & vbCrLf & _
           "Expected version: 1.1.3", _
           vbInformation, "Cook County Due Diligence v1.2.1"
#Else
    MsgBox "Mac helper diagnostics are only needed on Excel for Mac.", vbInformation
#End If
End Sub

Private Sub TreasurerBrowserAssist(ByVal pin As String, ByRef overviewText As String, _
                                   ByRef historyText As String, ByRef overviewOK As Boolean, _
                                   ByRef historyOK As Boolean)
#If Mac Then
    Dim payload As String, pOverview As Long, pHistory As Long
    Dim overviewPart As String, historyPart As String

    On Error GoTo BrowserAssistFailed

    ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
        "Safari browser assist starting for Treasurer only. Other county sources are not being rerun."

    payload = AppleScriptTask("CookPropertyTreasurerBrowser.applescript", "treasurerBrowserLookup", FormatPin(pin))

    If Left$(payload, 10) = "__ERROR__|" Then
        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Browser assist unavailable: " & Mid$(payload, 11)
        Exit Sub
    End If

    pOverview = InStr(1, payload, "__OVERVIEW__", vbBinaryCompare)
    pHistory = InStr(1, payload, "__HISTORY__", vbBinaryCompare)

    If pOverview > 0 And pHistory > pOverview Then
        overviewPart = Mid$(payload, pOverview + Len("__OVERVIEW__"), _
                           pHistory - (pOverview + Len("__OVERVIEW__")))
        historyPart = Mid$(payload, pHistory + Len("__HISTORY__"))

        overviewPart = CollapseWhitespace(overviewPart)
        historyPart = CollapseWhitespace(historyPart)

        ' Step 1: the Overview is the property identity gate.
        If Not overviewOK And TreasurerOverviewLooksValid(overviewPart, pin) Then
            overviewText = overviewPart
            overviewOK = True
        End If

        ' Step 2: once the Overview has been validated for this PIN, the 20-year
        ' page is trusted as part of the same Safari property session if it
        ' contains genuine multi-year tax-history content. It does not need to
        ' repeat the PIN in visible page text.
        If Not historyOK Then
            If overviewOK And TreasurerHistoryContentLooksValid(historyPart) Then
                historyText = historyPart
                historyOK = True
            ElseIf TreasurerHistoryPageLooksValid(historyPart, pin) Then
                historyText = historyPart
                historyOK = True
            End If
        End If

        If overviewOK And historyOK Then
            ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
                "Safari browser assist succeeded: Overview + 20-Year History captured."
        ElseIf overviewOK Then
            ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
                "Safari browser assist captured the Overview; 20-Year History still requires manual review."
        ElseIf historyOK Then
            ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
                "Safari browser assist captured 20-Year History; Overview still requires manual review."
        Else
            ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
                "Safari captured pages, but Overview did not validate for this PIN and the 20-year page could not be trusted as the same property session."
        End If
    End If

    Exit Sub

BrowserAssistFailed:
    ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
        "Safari helper unavailable/failed: " & CStr(Err.Number) & " - " & Err.Description & ". Full report is continuing."
    Err.Clear
#Else
    ' Windows equivalent of the Mac Safari helper: drive a real, hidden Internet
    ' Explorer session via COM automation so the Treasurer site's JS-rendered,
    ' session-dependent pages actually load (a plain WinHTTP GET cannot execute
    ' the page's script or carry a browser session the way IE's own engine can).
    ' Requires no extra install - the IE COM automation object ships with
    ' Windows - but some newer/locked-down Windows builds have it disabled.
    ' Any failure here degrades gracefully: Treasurer stays PARTIAL and the
    ' rest of the report is unaffected, exactly like the Mac path.
    Dim ie As Object
    Dim overviewUrl As String, historyUrl As String
    Dim overviewPart As String, historyPart As String

    On Error GoTo BrowserAssistFailedWin

    ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
        "Internet Explorer browser assist starting for Treasurer only. Other county sources are not being rerun."

    overviewUrl = CStr(GetConfigValue("Treasurer Overview Results URL", _
        "https://www.cookcountytreasurer.com/yourpropertytaxoverviewresults.aspx")) & "?pin=" & UrlEncode(FormatPin(pin))
    historyUrl = CStr(GetConfigValue("Treasurer 20-Year History URL", _
        "https://www.cookcountytreasurer.com/taxbillhistoryresults.aspx")) & "?pin=" & UrlEncode(FormatPin(pin))

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Silent = True
    ie.Visible = (UCase$(Trim$(CStr(GetConfigValue("Show Treasurer Browser Window", "NO")))) = "YES")

    overviewPart = CollapseWhitespace(IeNavigateAndGetText(ie, overviewUrl))
    historyPart = CollapseWhitespace(IeNavigateAndGetText(ie, historyUrl))

    ie.Quit
    Set ie = Nothing

    ' Step 1: the Overview is the property identity gate.
    If Not overviewOK And TreasurerOverviewLooksValid(overviewPart, pin) Then
        overviewText = overviewPart
        overviewOK = True
    End If

    ' Step 2: once the Overview has been validated for this PIN, the 20-year
    ' page is trusted as part of the same IE property session if it contains
    ' genuine multi-year tax-history content. It does not need to repeat the
    ' PIN in visible page text.
    If Not historyOK Then
        If overviewOK And TreasurerHistoryContentLooksValid(historyPart) Then
            historyText = historyPart
            historyOK = True
        ElseIf TreasurerHistoryPageLooksValid(historyPart, pin) Then
            historyText = historyPart
            historyOK = True
        End If
    End If

    If overviewOK And historyOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Internet Explorer browser assist succeeded: Overview + 20-Year History captured."
    ElseIf overviewOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Internet Explorer browser assist captured the Overview; 20-Year History still requires manual review."
    ElseIf historyOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Internet Explorer browser assist captured 20-Year History; Overview still requires manual review."
    Else
        ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
            "Internet Explorer captured pages, but Overview did not validate for this PIN and the 20-year page could not be trusted as the same property session."
    End If

    Exit Sub

BrowserAssistFailedWin:
    On Error Resume Next
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
    On Error GoTo 0
    ThisWorkbook.Worksheets("Tax Detail").Range("H74").value = _
        "Internet Explorer browser assist unavailable/failed: " & CStr(Err.Number) & " - " & Err.Description & _
        ". IE automation is disabled on some newer Windows builds - the report is continuing without it. " & _
        "Use the Tax Detail side-tab links to check the Treasurer site manually."
    Err.Clear
#End If
End Sub

Private Function IeNavigateAndGetText(ByVal ie As Object, ByVal url As String) As String
    Dim startTime As Double, timeoutSeconds As Double, settleStart As Double

    timeoutSeconds = CDbl(GetConfigValue("Treasurer Browser Timeout (sec)", 25))

    ie.Navigate url
    startTime = Timer
    Do While (ie.Busy Or ie.readyState <> 4) And (Timer - startTime) < timeoutSeconds
        DoEvents
    Loop

    ' A little extra settle time: some pages report "done loading" before
    ' their script-rendered content has actually landed in the DOM.
    settleStart = Timer
    Do While (Timer - settleStart) < 1.5
        DoEvents
    Loop

    On Error Resume Next
    IeNavigateAndGetText = ie.document.body.innerText
    On Error GoTo 0
End Function

Public Sub RunTreasurerBrowserLookup_v121()
#If Mac Then
    Dim pin As String, pinAddress As String
    Dim overviewText As String, historyText As String
    Dim overviewOK As Boolean, historyOK As Boolean

    On Error GoTo LookupError

    pin = ReadPinFromStartSheet(pinAddress)
    If Len(pin) <> 14 Then
        MsgBox "Enter a valid Cook County PIN on Start first.", vbExclamation
        Exit Sub
    End If

    If Not TreasurerBrowserHelperReady_v121() Then
        MsgBox "The separate Treasurer Safari helper could not be loaded for this manual lookup." & vbCrLf & vbCrLf & _
               "Run 'Install Mac Helpers v1.1.3.command' once, then try the green Treasurer button again." & vbCrLf & vbCrLf & _
               "The automatic full report never stops for this condition; Treasurer simply remains PARTIAL.", _
               vbExclamation, "Treasurer Browser Helper"
        Exit Sub
    End If

    EnsureV121WorkbookStructure
    ApplyV121WorkbookSettings
    InstallTreasurerLookupButton_v121

    UpdateSourceStatus "Cook County Treasurer", "Running", "Safari browser assist", _
                       "Entering the PIN on the official Treasurer site."

    TreasurerBrowserAssist pin, overviewText, historyText, overviewOK, historyOK

    If overviewOK Or historyOK Then
        PopulateTreasurerCapturedData pin, overviewText, historyText, overviewOK, historyOK

        If overviewOK And historyOK Then
            UpdateSourceStatus "Cook County Treasurer", "OK", "Safari browser assist finished", _
                               "Overview + 20-Year History captured from the live Treasurer property session."
            MsgBox "Treasurer browser lookup succeeded." & vbCrLf & vbCrLf & _
                   "Overview: OK" & vbCrLf & "20-Year History: OK", _
                   vbInformation, "Treasurer Lookup v1.2.1"
        Else
            UpdateSourceStatus "Cook County Treasurer", "PARTIAL", "Safari browser assist finished", _
                               "Some Treasurer property data was captured; one page still requires manual review."
            MsgBox "Treasurer browser lookup returned partial property data." & vbCrLf & vbCrLf & _
                   "Overview: " & IIf(overviewOK, "OK", "Not captured") & vbCrLf & _
                   "20-Year History: " & IIf(historyOK, "OK", "Not captured"), _
                   vbInformation, "Treasurer Lookup v1.2.1"
        End If
    Else
        UpdateSourceStatus "Cook County Treasurer", "PARTIAL", "Safari browser assist stopped", _
                           "Safari did not return property-specific Treasurer data."
        MsgBox "Safari did not return property-specific Treasurer data." & vbCrLf & vbCrLf & _
               "Check Tax Detail > Browser Assist for the exact reason.", _
               vbExclamation, "Treasurer Lookup v1.2.1"
    End If
    Exit Sub

LookupError:
    UpdateSourceStatus "Cook County Treasurer", "PARTIAL", "Safari browser assist error", Err.Description
    MsgBox "Treasurer browser lookup error " & CStr(Err.Number) & ": " & Err.Description, _
           vbExclamation, "Treasurer Lookup v1.2.1"
    Err.Clear
#Else
    MsgBox "The Safari Treasurer lookup button is Mac-only.", vbInformation
#End If
End Sub

Public Sub TestTreasurerBrowserAssist_v121()
    RunTreasurerBrowserLookup_v121
End Sub

Private Sub PopulateTreasurerCapturedData(ByVal pin As String, ByVal overviewText As String, _
                                          ByVal historyText As String, ByVal overviewOK As Boolean, _
                                          ByVal historyOK As Boolean)
    Dim refundBlock As String, historyBlock As String, debtBlock As String
    Dim propertyLoc As String, mailingInfo As String
    Dim totalDue As String, debtAmt As String, propValue As String, debtPct As String
    Dim firstYear As String, firstAmount As String, lastYear As String, lastAmount As String
    Dim diffAmount As String, pctChange As String

    If overviewOK Then
        ThisWorkbook.Worksheets("Tax Detail").Range("B18").value = "Safari property-session lookup succeeded"

        propertyLoc = CleanTreasurerText(TextBetween(overviewText, "Property Location:", "Volume:"))
        If Len(propertyLoc) = 0 Then propertyLoc = CleanTreasurerText(TextBetween(overviewText, "Property Location:", "Mailing Information:"))
        mailingInfo = CleanTreasurerText(TextBetween(overviewText, "Mailing Information:", "Tax Year"))

        If Len(propertyLoc) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D18").value = propertyLoc
        If Len(mailingInfo) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("H22").value = mailingInfo

        refundBlock = TextBetween(overviewText, "Are There Any Overpayments on Your PIN?", "Have You Received Your Exemptions")
        If Len(refundBlock) > 0 Then
            ThisWorkbook.Worksheets("Tax Detail").Range("F18").value = CleanTreasurerText(refundBlock)
        End If

        totalDue = LastValueAfterLabel(overviewText, "Total Amount Due:", "$0123456789,.")
        If Len(totalDue) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("H18").value = totalDue

        ParseTreasurerInstallments overviewText
        ParseTreasurerExemptionGrid overviewText

        historyBlock = TextBetween(overviewText, "20-Year Property Tax Bill History", "Taxing District Debt Attributed to Your Property")
        If Len(historyBlock) > 0 Then
            ParseTreasurerTwentyYearHistory historyBlock, firstYear, firstAmount, lastYear, lastAmount
            diffAmount = ValueAfterLabel(historyBlock, "Difference", "+-$0123456789,.")
            pctChange = ValueAfterLabel(historyBlock, "Percent Change", "+-0123456789.%")

            If Len(firstYear) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("B20").value = firstYear & ": $" & firstAmount
            If Len(lastYear) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D20").value = lastYear & ": $" & lastAmount
            If Len(diffAmount) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("F20").value = diffAmount
            If Len(pctChange) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("H20").value = pctChange
        End If

        debtBlock = TextBetween(overviewText, "Taxing District Debt Attributed to Your Property", "Highlights of Your Taxing Districts")
        If Len(debtBlock) = 0 Then debtBlock = TextBetween(overviewText, "Taxing District Debt Attributed to Your Property", "Reports and Data")

        debtAmt = ValueAfterLabel(debtBlock, "Total Taxing District Debt Attributed to Your Property", "$0123456789,.")
        propValue = ValueAfterLabel(debtBlock, "Property Value", "$0123456789,.")
        debtPct = ValueAfterLabel(debtBlock, "Total Debt % Attributed to Your Property Value", "0123456789.%")

        If Len(debtAmt) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("B22").value = debtAmt
        If Len(propValue) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D22").value = propValue
        If Len(debtPct) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("F22").value = debtPct
    End If

    If historyOK Then
        ParseTreasurerTwentyYearHistory historyText, firstYear, firstAmount, lastYear, lastAmount
        If Len(firstYear) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("B20").value = firstYear & ": $" & firstAmount
        If Len(lastYear) > 0 Then ThisWorkbook.Worksheets("Tax Detail").Range("D20").value = lastYear & ": $" & lastAmount
    End If
End Sub

Private Function TreasurerTryPinPage(ByVal baseUrl As String, ByVal pin As String) As String
    Dim html As String, txt As String, formattedPin As String, url As String

    formattedPin = FormatPin(pin)

    url = baseUrl & "?pin=" & UrlEncode(formattedPin)
    On Error Resume Next
    html = HttpGet(url)
    On Error GoTo 0
    txt = CollapseWhitespace(HtmlToText(html))
    If Len(txt) > 0 Then TreasurerTryPinPage = txt

    If InStr(1, txt, formattedPin, vbTextCompare) = 0 And InStr(1, txt, pin, vbTextCompare) = 0 Then
        url = baseUrl & "?pin=" & pin
        On Error Resume Next
        html = HttpGet(url)
        On Error GoTo 0
        txt = CollapseWhitespace(HtmlToText(html))
        If Len(txt) > 0 Then TreasurerTryPinPage = txt
    End If

    If InStr(1, txt, formattedPin, vbTextCompare) = 0 And InStr(1, txt, pin, vbTextCompare) = 0 Then
        url = baseUrl & "?PIN=" & pin
        On Error Resume Next
        html = HttpGet(url)
        On Error GoTo 0
        txt = CollapseWhitespace(HtmlToText(html))
        If Len(txt) > 0 Then TreasurerTryPinPage = txt
    End If
End Function

Private Function TreasurerOverviewLooksValid(ByVal txt As String, ByVal pin As String) As Boolean
    Dim pinMatch As Boolean, markerCount As Long

    pinMatch = (InStr(1, txt, FormatPin(pin), vbTextCompare) > 0) Or _
               (InStr(1, txt, pin, vbTextCompare) > 0)

    If InStr(1, txt, "Overview - Payments", vbTextCompare) > 0 Then markerCount = markerCount + 1
    If InStr(1, txt, "Property Location:", vbTextCompare) > 0 Then markerCount = markerCount + 1
    If InStr(1, txt, "Are There Any Overpayments on Your PIN?", vbTextCompare) > 0 Then markerCount = markerCount + 1
    If InStr(1, txt, "Have You Received Your Exemptions", vbTextCompare) > 0 Then markerCount = markerCount + 1
    If InStr(1, txt, "Original Billed Amount:", vbTextCompare) > 0 Then markerCount = markerCount + 1
    If InStr(1, txt, "Total Amount Due:", vbTextCompare) > 0 Then markerCount = markerCount + 1

    TreasurerOverviewLooksValid = pinMatch And (markerCount >= 2)
End Function

Private Function TreasurerHistoryPageLooksValid(ByVal txt As String, ByVal pin As String) As Boolean
    Dim pinMatch As Boolean

    pinMatch = (InStr(1, txt, FormatPin(pin), vbTextCompare) > 0) Or _
               (InStr(1, txt, pin, vbTextCompare) > 0)

    TreasurerHistoryPageLooksValid = pinMatch And TreasurerHistoryContentLooksValid(txt)
End Function

Private Function TreasurerHistoryContentLooksValid(ByVal txt As String) As Boolean
    Dim y As Long, yearHits As Long, dollarHits As Long, p As Long, searchPos As Long

    If InStr(1, txt, "20-Year Tax Bill History", vbTextCompare) = 0 And _
       InStr(1, txt, "20-Year Property Tax Bill History", vbTextCompare) = 0 And _
       InStr(1, txt, "20-Year History", vbTextCompare) = 0 Then
        Exit Function
    End If

    For y = 1999 To Year(Date) + 1
        If InStr(1, txt, CStr(y), vbTextCompare) > 0 Then
            yearHits = yearHits + 1
            If yearHits >= 5 Then Exit For
        End If
    Next y

    searchPos = 1
    Do
        p = InStr(searchPos, txt, "$", vbBinaryCompare)
        If p = 0 Then Exit Do
        dollarHits = dollarHits + 1
        If dollarHits >= 5 Then Exit Do
        searchPos = p + 1
    Loop

    TreasurerHistoryContentLooksValid = (yearHits >= 3 And dollarHits >= 3)
End Function

Private Function CleanTreasurerText(ByVal s As String) As String
    s = CollapseWhitespace(s)
    If Len(s) > 500 Then s = Left$(s, 500) & "..."
    CleanTreasurerText = Trim$(s)
End Function

Private Function LastValueAfterLabel(ByVal s As String, ByVal labelText As String, ByVal allowedChars As String) As String
    Dim searchPos As Long, p As Long, valueText As String, lastValue As String

    searchPos = 1

    Do
        p = InStr(searchPos, s, labelText, vbTextCompare)
        If p = 0 Then Exit Do

        valueText = ValueAfterLabel(Mid$(s, p), labelText, allowedChars)
        If Len(valueText) > 0 Then lastValue = valueText

        searchPos = p + Len(labelText)
    Loop

    LastValueAfterLabel = lastValue
End Function

Private Sub ParseTreasurerInstallments(ByVal txt As String)
    Dim ws As Worksheet
    Dim y As Long, yearPos As Long, yearEnd As Long, p As Long, testYear As Long
    Dim seg As String, firstPos As Long, secondPos As Long, endPos As Long
    Dim instSeg As String, outRow As Long
    Dim billed As String, dueDate As String, currentDue As String
    Dim taxAmt As String, interestAmt As String

    Set ws = ThisWorkbook.Worksheets("Tax Detail")
    outRow = 37

    For y = Year(Date) + 1 To 2000 Step -1
        yearPos = InStr(1, txt, "Tax Year " & CStr(y) & " (billed in ", vbTextCompare)

        If yearPos > 0 Then
            yearEnd = Len(txt) + 1

            For testYear = 2000 To Year(Date) + 2
                If testYear <> y Then
                    p = InStr(yearPos + 1, txt, "Tax Year " & CStr(testYear) & " (billed in ", vbTextCompare)
                    If p > 0 And p < yearEnd Then yearEnd = p
                End If
            Next testYear

            seg = Mid$(txt, yearPos, yearEnd - yearPos)

            firstPos = InStr(1, seg, "1st INSTALLMENT - Tax Year", vbTextCompare)
            secondPos = InStr(1, seg, "2nd INSTALLMENT - Tax Year", vbTextCompare)

            If firstPos > 0 And outRow <= 43 Then
                If secondPos > firstPos Then
                    endPos = secondPos
                Else
                    endPos = Len(seg) + 1
                End If

                instSeg = Mid$(seg, firstPos, endPos - firstPos)
                billed = ValueAfterLabel(instSeg, "Original Billed Amount", "$0123456789,.")
                dueDate = ValueAfterLabel(instSeg, "Due Date", "0123456789/")
                currentDue = ValueAfterLabel(instSeg, "Current Amount Due", "$0123456789,.")
                taxAmt = ValueAfterLabel(instSeg, "Tax", "$0123456789,.")
                interestAmt = ValueAfterLabel(instSeg, "Interest", "$0123456789,.")

                ws.Cells(outRow, 1).value = y
                ws.Cells(outRow, 2).value = "1st"
                ws.Cells(outRow, 3).value = billed
                ws.Cells(outRow, 4).value = dueDate
                ws.Cells(outRow, 5).value = currentDue
                ws.Cells(outRow, 6).value = taxAmt
                ws.Cells(outRow, 7).value = interestAmt
                outRow = outRow + 1
            End If

            If secondPos > 0 And outRow <= 43 Then
                p = InStr(secondPos + Len("2nd INSTALLMENT - Tax Year"), seg, "1st INSTALLMENT - Tax Year", vbTextCompare)
                If p > secondPos Then
                    endPos = p
                Else
                    endPos = Len(seg) + 1
                End If

                instSeg = Mid$(seg, secondPos, endPos - secondPos)
                billed = ValueAfterLabel(instSeg, "Original Billed Amount", "$0123456789,.")
                dueDate = ValueAfterLabel(instSeg, "Due Date", "0123456789/")
                currentDue = ValueAfterLabel(instSeg, "Current Amount Due", "$0123456789,.")
                taxAmt = ValueAfterLabel(instSeg, "Tax", "$0123456789,.")
                interestAmt = ValueAfterLabel(instSeg, "Interest", "$0123456789,.")

                ws.Cells(outRow, 1).value = y
                ws.Cells(outRow, 2).value = "2nd"
                ws.Cells(outRow, 3).value = billed
                ws.Cells(outRow, 4).value = dueDate
                ws.Cells(outRow, 5).value = currentDue
                ws.Cells(outRow, 6).value = taxAmt
                ws.Cells(outRow, 7).value = interestAmt
                outRow = outRow + 1
            End If
        End If
    Next y
End Sub

Private Sub ParseTreasurerExemptionGrid(ByVal txt As String)
    Dim blockText As String
    Dim names As Variant, rowNums As Variant
    Dim years As Variant, colNums As Variant
    Dim i As Long, j As Long, valueText As String

    blockText = TextBetween(txt, "Have You Received Your Exemptions in These Tax Years?", _
                           "20-Year Property Tax Bill History")
    If Len(blockText) = 0 Then Exit Sub

    names = Array("Homeowner Exemption", "Senior Citizen Exemption", "Senior Freeze Exemption", _
                  "Returning Veteran Exemption", "Disabled Person Exemption", "Disabled Veteran Exemption")
    rowNums = Array(27, 28, 29, 30, 31, 32)
    years = Array("2024", "2023", "2022", "2021")
    colNums = Array(2, 3, 4, 5)

    For i = LBound(names) To UBound(names)
        For j = LBound(years) To UBound(years)
            valueText = TreasurerExemptionValue(blockText, CStr(names(i)), CStr(years(j)))
            If Len(valueText) > 0 Then
                ThisWorkbook.Worksheets("Tax Detail").Cells(CLng(rowNums(i)), CLng(colNums(j))).value = valueText
            End If
        Next j
    Next i
End Sub

Private Function TreasurerExemptionValue(ByVal blockText As String, ByVal exemptionName As String, ByVal taxYear As String) As String
    Dim yearPos As Long, nextYearPos As Long, sectionText As String
    Dim p As Long, valueText As String

    yearPos = InStr(1, blockText, "Tax Year " & taxYear, vbTextCompare)

    If yearPos > 0 Then
        nextYearPos = Len(blockText) + 1
        p = InStr(yearPos + 1, blockText, "Tax Year ", vbTextCompare)
        If p > 0 Then nextYearPos = p
        sectionText = Mid$(blockText, yearPos, nextYearPos - yearPos)
    Else
        sectionText = blockText
    End If

    p = InStr(1, sectionText, exemptionName & ":", vbTextCompare)
    If p = 0 Then Exit Function

    valueText = UCase$(Trim$(Mid$(sectionText, p + Len(exemptionName) + 1, 8)))
    If Left$(valueText, 3) = "YES" Then
        TreasurerExemptionValue = "YES"
    ElseIf Left$(valueText, 2) = "NO" Then
        TreasurerExemptionValue = "NO"
    End If
End Function

Private Sub ParseTreasurerTwentyYearHistory(ByVal blockText As String, _
                                            ByRef firstYear As String, ByRef firstAmount As String, _
                                            ByRef lastYear As String, ByRef lastAmount As String)
    Dim y As Long, amountText As String, outRow As Long
    Dim minYear As Long, maxYear As Long
    Dim prevAmount As Double, currentAmount As Double, hasPrev As Boolean
    Dim amountClean As String

    minYear = 9999
    maxYear = 0
    outRow = 47

    ' Rebuild the 20-year section on every Treasurer capture so stale rows from
    ' a prior PIN cannot survive a partial browser run.
    ThisWorkbook.Worksheets("Tax Detail").Range("A47:E68").ClearContents
    hasPrev = False

    For y = 1999 To Year(Date) + 1
        If InStr(1, blockText, "Tax Year " & CStr(y), vbTextCompare) > 0 Or _
           InStr(1, blockText, CStr(y) & ":", vbTextCompare) > 0 Then

            amountText = LooseNumberAfterYear(blockText, CStr(y))

            If Len(amountText) > 0 Then
                If y < minYear Then
                    minYear = y
                    firstYear = CStr(y)
                    firstAmount = amountText
                End If

                If y > maxYear Then
                    maxYear = y
                    lastYear = CStr(y)
                    lastAmount = amountText
                End If

                If outRow <= 68 Then
                    With ThisWorkbook.Worksheets("Tax Detail")
                        .Cells(outRow, 1).value = y
                        .Cells(outRow, 2).value = "$" & amountText

                        amountClean = Replace(Replace(amountText, ",", ""), "$", "")
                        If IsNumeric(amountClean) Then
                            currentAmount = CDbl(amountClean)

                            If hasPrev Then
                                .Cells(outRow, 3).value = currentAmount - prevAmount
                                .Cells(outRow, 3).NumberFormat = "$#,##0.00;[Red]-$#,##0.00"

                                If prevAmount <> 0 Then
                                    .Cells(outRow, 4).value = (currentAmount - prevAmount) / prevAmount
                                    .Cells(outRow, 4).NumberFormat = "0.00%;[Red]-0.00%"
                                End If
                            End If

                            prevAmount = currentAmount
                            hasPrev = True
                        End If

                        .Cells(outRow, 5).value = "Cook County Treasurer"
                    End With
                    outRow = outRow + 1
                End If
            End If
        End If
    Next y
End Sub

Private Sub FetchTIF(ByVal pin As String)
    Dim x As String, y As String, currentUrl As String, detailUrl As String
    Dim url As String, js As String, detailJs As String
    Dim currentHit As Boolean, detailHit As Boolean
    Dim currentName As String, currentId As String, detailName As String, detailMuni As String, firstYear As String

    x = GetPropertyValue("Centroid X (3435)")
    y = GetPropertyValue("Centroid Y (3435)")
    If Len(x) = 0 Or Len(y) = 0 Then
        Err.Raise vbObjectError + 171, , "Parcel centroid CRS 3435 coordinates are unavailable."
    End If

    currentUrl = CStr(GetConfigValue("TIF Current Boundary URL", _
        "https://gis.cookcountyil.gov/traditional/rest/services/politicalBoundary/MapServer/24"))
    detailUrl = CStr(GetConfigValue("TIF Detail URL", _
        "https://gis.cookcountyil.gov/traditional/rest/services/tifSrvc/MapServer/3"))

    url = currentUrl & "/query?geometry=" & UrlEncode(x & "," & y) & _
          "&geometryType=esriGeometryPoint&inSR=3435&spatialRel=esriSpatialRelIntersects" & _
          "&outFields=*&returnGeometry=false&f=pjson"
    js = HttpGet(url)
    EnsureArcGisResponse js, "Cook County current TIF boundary"
    currentHit = (ArcGisFeatureCount(js) > 0)

    If currentHit Then
        currentName = JsonFirstNonBlank(js, Array("AGENCY_DESCRIPTION", "AGENCY_DESC", "AGENCY_DES", "TIF_NAME"))
        currentId = JsonFirstNonBlank(js, Array("AGENCY", "AGENCYNUM"))
        AppendIncentive "TIF - Current Boundary", "Intersection found", _
                        FirstText(currentName, currentId), "Cook County current political boundary", _
                        "Primary current-boundary signal; confirm district status for transaction-specific reliance.", currentUrl
    Else
        AppendIncentive "TIF - Current Boundary", "No intersection found", "", _
                        "Cook County current political boundary", _
                        "No current boundary intersection at parcel centroid; verify if the parcel lies on/near a boundary.", currentUrl
    End If

    url = detailUrl & "/query?geometry=" & UrlEncode(x & "," & y) & _
          "&geometryType=esriGeometryPoint&inSR=3435&spatialRel=esriSpatialRelIntersects" & _
          "&outFields=*&returnGeometry=false&f=pjson"
    detailJs = HttpGet(url)
    EnsureArcGisResponse detailJs, "Cook County TIF detail layer"
    detailHit = (ArcGisFeatureCount(detailJs) > 0)

    If detailHit Then
        detailName = JsonFirstNonBlank(detailJs, Array("TIF_NAME", "TIF_Name", "AGENCY_DES"))
        detailMuni = JsonFirstNonBlank(detailJs, Array("Municipality", "MUNICIPALITY"))
        firstYear = JsonScalar(detailJs, "First_Year")
        AppendIncentive "TIF - 2024 Detail", "Intersection found", detailName, _
                        "Cook County TIF revenue/detail layer (2024)", _
                        "Municipality: " & detailMuni & IIf(Len(firstYear) > 0, "; First Year: " & firstYear, "") & _
                        ". Secondary cross-check; layer is labeled 2024.", detailUrl
    Else
        AppendIncentive "TIF - 2024 Detail", "No intersection found", "", _
                        "Cook County TIF revenue/detail layer (2024)", _
                        "Secondary 2024 layer; not treated as the sole current-boundary source.", detailUrl
    End If

    If currentHit <> detailHit Then
        AppendIssue "MEDIUM", "Current TIF boundary and 2024 TIF detail layer disagree.", _
                    "Cook County GIS TIF layers", _
                    "Review the parcel in Cook County GIS and confirm whether the TIF is current, recently created, terminated, or changed."
    End If
End Sub

Private Sub FetchEnterpriseZone(ByVal pin As String)
    Dim appUrl As String, countySignal As String, signalResult As String, signalName As String

    appUrl = CStr(GetConfigValue("DCEO Enterprise Zone App", _
        "https://idor.maps.arcgis.com/apps/webappviewer/index.html?id=f82fc6b62fde435abb41f5f72db2db48"))
    countySignal = Trim$(GetPropertyValue("Enterprise Zone (Assessor Signal)"))

    If InStr(1, countySignal, "Signal present:", vbTextCompare) = 1 Then
        signalResult = "County signal present"
        signalName = Trim$(Replace(countySignal, "Signal present:", "", 1, 1, vbTextCompare))
        AppendIssue "MEDIUM", "Cook County Assessor data returns an Enterprise Zone signal for this parcel.", _
                    "Assessor Parcel Universe / Illinois DCEO", _
                    "Treat the county signal as a research lead. Confirm the parcel boundary in the official Illinois DCEO Enterprise Zone map before relying on eligibility."
    Else
        signalResult = "No county signal returned"
        signalName = ""
    End If

    AppendIncentive "Illinois Enterprise Zone - Assessor Signal", _
                    signalResult, signalName, _
                    "Assessor Parcel Universe (" & GetPropertyAsOf("Enterprise Zone (Assessor Signal)") & ")", _
                    "County spatial signal only; not a substitute for DCEO boundary verification.", _
                    "https://datacatalog.cookcountyil.gov/d/pabr-t5kh"

    SetProperty "Enterprise Zone (DCEO Verification)", "Manual boundary verification required", _
                "Illinois DCEO", "Current official map", "PARTIAL", _
                "Use the official DCEO map for the final boundary check."

    AppendIncentive "Illinois Enterprise Zone - DCEO Verification", _
                    "Manual boundary verification required", "", _
                    "Official Illinois DCEO Enterprise Zone map", _
                    "Final Enterprise Zone boundary verification remains manual in v1.0.4.", appUrl
End Sub

Private Sub BuildIssues()
    Dim aClass As String, gClass As String, mTax As String, mSpatial As String
    Dim taxWs As Worksheet, tr As Long, taxLast As Long, payStatus As String
    aClass = NormalizeCompare(GetPropertyValue("Assessor Class"))
    gClass = NormalizeCompare(GetPropertyValue("GIS Building Class"))
    mTax = NormalizeMunicipality(GetPropertyValue("Municipality - Tax Record"))
    mSpatial = NormalizeMunicipality(GetPropertyValue("Municipality - Spatial"))

    If Len(aClass) > 0 And Len(gClass) > 0 And aClass <> gClass Then
        AppendIssue "HIGH", "Assessor class and current CookViewer building class do not match.", _
                    "Parcel Universe / CookViewer", _
                    "Review both current sources; do not silently overwrite either classification."
    End If

    If Len(mTax) > 0 And Len(mSpatial) > 0 And mTax <> mSpatial Then
        AppendIssue "HIGH", "Tax-record municipality and spatial municipality differ.", _
                    "CookViewer tax municipality / Cook County Municipality Boundary", _
                    "Confirm the municipal boundary for zoning/incentive work. Cook County itself warns tax and spatial municipality concepts can disagree."
    End If

    If Len(GetPropertyValue("Owner Name")) = 0 Then
        AppendIssue "MEDIUM", "Owner name was not returned.", "Assessor Addresses", _
                    "Check recent sales/recorded documents; Assessor owner/mailing data can lag."
    End If

    If Len(GetPropertyValue("Tax Code")) = 0 Then
        AppendIssue "HIGH", "Current tax code was not verified.", "Property Tax Portal / Parcel Universe / CookViewer", _
                    "Verify on the Cook County Property Tax Portal."
    End If

    If Len(GetPropertyValue("Building (SqFt)")) = 0 Then
        AppendIssue "MEDIUM", "Building square footage was not returned by current CookViewer/Portal sources.", _
                    "CookViewer / Property Tax Portal", _
                    "Verify building area from Assessor detail, municipal records, appraisal, survey, or reliable property documents."
    End If

    Set taxWs = ThisWorkbook.Worksheets("Tax History")
    taxLast = LastUsedRow(taxWs, 1)
    For tr = 5 To taxLast
        payStatus = Trim$(CStr(taxWs.Cells(tr, 3).value))
        If InStr(1, payStatus, "Pay Online:", vbTextCompare) > 0 Then
            AppendIssue "HIGH", "Property Tax Portal shows a current balance/payment amount.", _
                        "Cook County Property Tax Portal", _
                        "Review current Treasurer payment status before closing or relying on tax-payment assumptions. Portal status: " & payStatus
            Exit For
        End If
    Next tr
End Sub

Private Sub BuildReport()
    Dim ws As Worksheet, issuesText As String, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Report")

    ws.Range("B5").NumberFormat = "@"
    ws.Range("B5").value = FormatPin(GetPropertyValue("PIN"))
    ws.Range("B6").value = GetPropertyValue("Property Address")
    ws.Range("B7").value = FirstText(GetPropertyValue("Municipality - Spatial"), GetPropertyValue("Municipality - Tax Record"))
    ws.Range("B8").value = GetPropertyValue("Township")
    ws.Range("B9").value = GetPropertyValue("Assessor Class")
    ws.Range("B10").value = GetPropertyValue("GIS Building Class")
    ws.Range("B11").value = GetPropertyValue("Building (SqFt)")
    ws.Range("B12").value = GetPropertyValue("Lot Size (SqFt)")

    ws.Range("E5").value = GetPropertyValue("Owner Name")
    ws.Range("E6").value = GetPropertyValue("Taxpayer / Mailing Name")
    ws.Range("E8").NumberFormat = "@"
    ws.Range("E8").value = GetPropertyValue("Tax Code")
    ws.Range("E12").value = Format(Date, "mmmm d, yyyy")

    ws.Range("E10").value = FindIncentiveResult("TIF - Current Boundary")
    ws.Range("E11").value = FindIncentiveResult("Illinois Enterprise Zone - Assessor Signal")
    If Len(ws.Range("E11").value) > 0 Then
        ws.Range("E11").value = ws.Range("E11").value & " | DCEO verify"
    Else
        ws.Range("E11").value = "DCEO verification required"
    End If

    lastRow = LastUsedRow(ThisWorkbook.Worksheets("Issues"), 1)
    For r = 5 To lastRow
        If Len(CStr(ThisWorkbook.Worksheets("Issues").Cells(r, 2).value)) > 0 Then
            issuesText = issuesText & ChrW(&H2022) & " " & ThisWorkbook.Worksheets("Issues").Cells(r, 1).value & ": " & _
                         ThisWorkbook.Worksheets("Issues").Cells(r, 2).value & vbCrLf
        End If
    Next r
    If Len(issuesText) = 0 Then issuesText = "No material automated warnings were generated. Public records should still be independently verified."
    ws.Range("A14").value = issuesText
End Sub

Private Sub ExportDueDiligencePDF(ByVal pin As String)
    Dim folder As String, fn As String, fullPath As String
    Dim arr As Variant
    Dim pdfBook As Workbook
    Dim pdfErr As Long, pdfDesc As String

    On Error GoTo PdfError

    gRunStage = "PDF: reading output folder setting"
    folder = CStr(GetConfigValue("PDF Output Folder", ""))

    If Len(Trim$(folder)) = 0 Then
        gRunStage = "PDF: deriving Downloads folder"
        folder = DefaultDownloadsFolder()
    End If

    gRunStage = "PDF: normalizing output path"
    folder = EnsureTrailingPathSeparator(folder)

    gRunStage = "PDF: building output filename"
    fn = "Cook_Property_Due_Diligence_" & pin & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    fullPath = folder & fn

    gRunStage = "PDF: building lean sheet list"
    arr = BuildPdfSheetList()

    gRunStage = "PDF: configuring included print areas"
    ConfigurePrintAreas arr

    ' Copy only the sheets that belong in this packet. Empty optional sheets
    ' and the Sources registry can be omitted to reduce PDF time and clutter.

    gRunStage = "PDF: copying finished sheets"
    ThisWorkbook.Worksheets(arr).Copy
    Set pdfBook = ActiveWorkbook

    gRunStage = "PDF: putting executive report first"
    pdfBook.Worksheets("Report").Move Before:=pdfBook.Worksheets(1)

    gRunStage = "PDF: exporting temporary workbook"
    pdfBook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    gRunStage = "PDF: closing temporary workbook"
    pdfBook.Close SaveChanges:=False
    Set pdfBook = Nothing

    ThisWorkbook.Worksheets("Start").Activate
    SetRunStatus "Complete - PDF saved to " & fullPath
    Exit Sub

PdfError:
    pdfErr = Err.Number
    pdfDesc = Err.Description

    On Error Resume Next
    If Not pdfBook Is Nothing Then pdfBook.Close SaveChanges:=False
    ThisWorkbook.Worksheets("Start").Activate
    On Error GoTo 0

    Err.Raise pdfErr, , gRunStage & " - " & pdfDesc
End Sub

Private Function EnsureTrailingPathSeparator(ByVal folder As String) As String
#If Mac Then
    If Right$(folder, 1) <> "/" Then folder = folder & "/"
#Else
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
#End If
    EnsureTrailingPathSeparator = folder
End Function

Private Function DefaultDownloadsFolder() As String
#If Mac Then
    Dim p As String, homePath As String, pos1 As Long, pos2 As Long
    Dim parts() As String

    p = ThisWorkbook.Path

    ' Normal POSIX path, e.g. /Users/markrogers/Downloads.
    pos1 = InStr(1, p, "/Users/", vbTextCompare)
    If pos1 > 0 Then
        pos2 = InStr(pos1 + Len("/Users/"), p, "/")
        If pos2 > 0 Then
            homePath = Left$(p, pos2 - 1)
            DefaultDownloadsFolder = homePath & "/Downloads"
            Exit Function
        End If
    End If

    ' Older HFS-style path, e.g. Macintosh HD:Users:markrogers:Downloads.
    pos1 = InStr(1, p, ":Users:", vbTextCompare)
    If pos1 > 0 Then
        parts = Split(Mid$(p, pos1 + Len(":Users:")), ":")
        If UBound(parts) >= 0 And Len(parts(0)) > 0 Then
            DefaultDownloadsFolder = "/Users/" & parts(0) & "/Downloads"
            Exit Function
        End If
    End If

    ' Safe fallback: save beside the workbook.
    If Len(p) > 0 Then
        DefaultDownloadsFolder = p
    Else
        Err.Raise vbObjectError + 392, , _
            "Could not derive the Mac Downloads folder from the workbook path."
    End If
#Else
    DefaultDownloadsFolder = Environ$("USERPROFILE") & "\Downloads"
#End If
End Function

Private Function DefaultDocumentsFolder() As String
#If Mac Then
    Dim result As String
    result = AppleScriptTask("CookPropertyHTTP.applescript", "documentsFolder", "")
    If Left$(result, 10) = "__ERROR__|" Then Err.Raise vbObjectError + 390, , "Could not resolve the Mac Documents folder: " & result
    DefaultDocumentsFolder = result
#Else
    DefaultDocumentsFolder = Environ$("USERPROFILE") & "\\Documents"
#End If
End Function

Private Sub OpenGeneratedFile(ByVal fullPath As String)
#If Mac Then
    ' v1.2.1 intentionally leaves the PDF closed on Mac.
#Else
    ThisWorkbook.FollowHyperlink fullPath
#End If
End Sub

' -----------------------
' Workbook writing helpers
' -----------------------
Private Sub ResetRunSheets()
    Dim s As Variant, ws As Worksheet, r As Long
    Dim lastRow As Long, lastCol As Long

    For Each s In Array("Assessment", "Tax History", "Sales-Deeds", "Documents", "Appeals", "Incentives", "Permits", "Issues")
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        lastRow = LastUsedRowAnyColumn(ws)
        lastCol = LastUsedColumn(ws)

        If lastRow >= 5 And lastCol >= 1 Then
            ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol)).ClearContents
        End If
    Next s

    Set ws = ThisWorkbook.Worksheets("Property")
    ws.Range("B5:F45").ClearContents

    Set ws = ThisWorkbook.Worksheets("Tax Detail")
    ws.Range("B4,D4,F4,H4").ClearContents
    ws.Range("A9:G15").ClearContents
    ws.Range("B18,D18,F18,H18,B20,D20,F20,H20,B22,D22,F22,H22").ClearContents
    ws.Range("B27:E32").ClearContents
    ws.Range("A37:G43").ClearContents
    ws.Range("A47:E68").ClearContents

    Set ws = ThisWorkbook.Worksheets("Report")
    ws.Range("B5:B12").ClearContents
    ws.Range("E5:E12").ClearContents
    ws.Range("A14:H21").ClearContents

    Set ws = ThisWorkbook.Worksheets("Start")
    For r = 12 To 24
        ws.Cells(r, 2).value = "Not run"
        ws.Cells(r, 3).ClearContents
        ws.Cells(r, 4).ClearContents
    Next r
End Sub

Private Function LastUsedRowAnyColumn(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedRowAnyColumn = 1
    Else
        LastUsedRowAnyColumn = f.row
    End If
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = f.Column
    End If
End Function

Private Sub SetRunStatus(ByVal msg As String)
    ThisWorkbook.Worksheets("Start").Range("B8").value = msg
    DoEvents
End Sub

Private Sub UpdateSourceStatus(ByVal sourceName As String, ByVal status As String, ByVal progress As String, ByVal notes As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Start")
    For r = 12 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), sourceName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).value = status
            ws.Cells(r, 3).value = progress & " " & Format(Now, "hh:mm:ss")
            ws.Cells(r, 4).value = notes
            Exit For
        End If
    Next r
    SetRunStatus sourceName & ": " & status
End Sub

Private Sub SetProperty(ByVal fieldName As String, ByVal value As Variant, ByVal sourceName As String, _
                        ByVal asOfText As String, ByVal status As String, ByVal notes As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Property")
    For r = 5 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), fieldName, vbTextCompare) = 0 Then
            If StrComp(fieldName, "PIN", vbTextCompare) = 0 Or _
               StrComp(fieldName, "Tax Code", vbTextCompare) = 0 Or _
               StrComp(fieldName, "GIS PIN Match", vbTextCompare) = 0 Then
                ws.Cells(r, 2).NumberFormat = "@"
            End If
            ws.Cells(r, 2).value = CStr(value)
            ws.Cells(r, 3).value = sourceName
            ws.Cells(r, 4).value = asOfText
            ws.Cells(r, 5).value = status
            ws.Cells(r, 6).value = notes
            Exit Sub
        End If
    Next r
End Sub

Private Function GetPropertyValue(ByVal fieldName As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Property")
    For r = 5 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), fieldName, vbTextCompare) = 0 Then
            GetPropertyValue = Trim$(CStr(ws.Cells(r, 2).value))
            Exit Function
        End If
    Next r
End Function

Private Function GetPropertyAsOf(ByVal fieldName As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Property")
    For r = 5 To 45
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), fieldName, vbTextCompare) = 0 Then
            GetPropertyAsOf = Trim$(CStr(ws.Cells(r, 4).value))
            Exit Function
        End If
    Next r
End Function


Private Sub AppendAppeal(ByVal level As String, ByVal taxYear As Variant, ByVal caseNo As String, _
                         ByVal appealType As String, ByVal status As String, ByVal beforeAV As Variant, _
                         ByVal afterAV As Variant, ByVal sourceNote As String)
    Dim ws As Worksheet, r As Long, chg As Variant
    Set ws = ThisWorkbook.Worksheets("Appeals")
    r = LastUsedRow(ws, 1) + 1
    If r < 5 Then r = 5
    If IsNumeric(beforeAV) And IsNumeric(afterAV) Then chg = CDbl(afterAV) - CDbl(beforeAV) Else chg = ""
    ws.Cells(r, 1).value = level
    ws.Cells(r, 2).value = taxYear
    ws.Cells(r, 3).value = caseNo
    ws.Cells(r, 4).value = appealType
    ws.Cells(r, 5).value = status
    ws.Cells(r, 6).value = beforeAV
    ws.Cells(r, 7).value = afterAV
    ws.Cells(r, 8).value = chg
    ws.Cells(r, 9).value = sourceNote
End Sub

Private Sub AppendIncentive(ByVal programName As String, ByVal result As String, ByVal nm As String, _
                            ByVal dataYear As String, ByVal verification As String, ByVal sourceUrl As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Incentives")
    r = LastUsedRow(ws, 1) + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).value = programName
    ws.Cells(r, 2).value = result
    ws.Cells(r, 3).value = nm
    ws.Cells(r, 4).value = dataYear
    ws.Cells(r, 5).value = verification
    ws.Cells(r, 6).value = sourceUrl
End Sub


Private Sub AppendDocument(ByVal docNo As String, ByVal docType As String, ByVal recordedDate As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Documents")
    r = LastUsedRow(ws, 1) + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).value = docNo
    ws.Cells(r, 2).value = docType
    ws.Cells(r, 3).value = recordedDate
    ws.Cells(r, 4).value = "Cook County Property Tax Portal document index; obtain the recorded instrument from the Clerk/Recorder for substantive review."
End Sub

Private Function FindIncentiveResult(ByVal programName As String) As String
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Incentives")
    lastRow = LastUsedRow(ws, 1)
    For r = 5 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).value), programName, vbTextCompare) = 0 Then
            FindIncentiveResult = CStr(ws.Cells(r, 2).value)
            If Len(CStr(ws.Cells(r, 3).value)) > 0 Then FindIncentiveResult = FindIncentiveResult & " - " & CStr(ws.Cells(r, 3).value)
            Exit Function
        End If
    Next r
End Function

Private Sub AppendIssue(ByVal priority As String, ByVal issueText As String, ByVal sourceText As String, ByVal verifyText As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Issues")
    r = LastUsedRow(ws, 1) + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).value = priority
    ws.Cells(r, 2).value = issueText
    ws.Cells(r, 3).value = sourceText
    ws.Cells(r, 4).value = verifyText
End Sub

Private Sub UpsertTaxHistory(ByVal taxYear As String, ByVal billed As String, ByVal statusText As String, _
                             ByVal taxRate As String, ByVal taxCode As String, ByVal exemptionsText As String)
    Dim ws As Worksheet, r As Long, lastRow As Long, foundRow As Long
    Set ws = ThisWorkbook.Worksheets("Tax History")
    lastRow = LastUsedRow(ws, 1)
    foundRow = 0

    For r = 5 To IIf(lastRow < 5, 5, lastRow)
        If Trim$(CStr(ws.Cells(r, 1).value)) = taxYear Then
            foundRow = r
            Exit For
        End If
    Next r

    If foundRow = 0 Then
        foundRow = lastRow + 1
        If foundRow < 5 Then foundRow = 5
        ws.Cells(foundRow, 1).value = taxYear
    End If

    If Len(billed) > 0 Then ws.Cells(foundRow, 2).value = billed
    If Len(statusText) > 0 Then ws.Cells(foundRow, 3).value = statusText
    If Len(taxRate) > 0 Then ws.Cells(foundRow, 4).value = taxRate
    If Len(taxCode) > 0 Then
        ws.Cells(foundRow, 5).NumberFormat = "@"
        ws.Cells(foundRow, 5).value = taxCode
    End If
    If Len(exemptionsText) > 0 Then ws.Cells(foundRow, 6).value = exemptionsText
    ws.Cells(foundRow, 7).value = "Cook County Property Tax Portal"
End Sub

Private Function IsIdentifierField(ByVal fieldName As String) As Boolean
    Select Case LCase$(Trim$(fieldName))
        Case "pin", "doc_no", "permit_number", "local_permit_number", "tax_code", "case_no", "appealtrk", "appealseq"
            IsIdentifierField = True
    End Select
End Function

Private Function CsvDisplayValue(ByVal fieldName As String, ByVal rawValue As String) As String
    Dim f As String
    f = LCase$(Trim$(fieldName))

    If (f = "sale_date" Or f = "date_issued") And Len(rawValue) >= 10 Then
        If Mid$(rawValue, 5, 1) = "-" And Mid$(rawValue, 8, 1) = "-" Then
            CsvDisplayValue = Left$(rawValue, 10)
            Exit Function
        End If
    End If

    CsvDisplayValue = rawValue
End Function

Private Sub WriteSelectedCsv(ByVal sheetName As String, ByVal csv As String, ByVal startRow As Long, ByVal fields As Variant)
    Dim headers() As String, i As Long
    ReDim headers(LBound(fields) To UBound(fields))
    For i = LBound(fields) To UBound(fields)
        headers(i) = CStr(fields(i))
    Next i
    WriteMappedTable sheetName, csv, startRow, fields, headers
End Sub

Private Sub WriteMappedTable(ByVal sheetName As String, ByVal csv As String, ByVal startRow As Long, _
                             ByVal fields As Variant, ByVal displayHeaders As Variant)
    Dim rows As Collection, d As Collection, ws As Worksheet
    Dim i As Long, j As Long, outRow As Long
    Dim fieldName As String, displayValue As String, targetCol As Long
    Set rows = ParseCsv(csv)
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If rows.count < 2 Then Exit Sub

    outRow = startRow
    For i = 2 To rows.count
        Set d = CsvRowDict(rows(1), rows(i))
        For j = LBound(fields) To UBound(fields)
            fieldName = CStr(fields(j))
            displayValue = CsvDisplayValue(fieldName, DictGet(d, fieldName))
            targetCol = j - LBound(fields) + 1

            If IsIdentifierField(fieldName) Then ws.Cells(outRow, targetCol).NumberFormat = "@"
            ws.Cells(outRow, targetCol).value = displayValue
        Next j
        outRow = outRow + 1
    Next i
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.rows.count, col).End(xlUp).row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Private Sub ConfigurePrintAreas(ByVal sheetList As Variant)
    Dim s As Variant, ws As Worksheet, lastRow As Long, lastCol As Long
    Dim lastRowCell As Range, lastColCell As Range

    For Each s In sheetList
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        Set lastRowCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        Set lastColCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

        If Not lastRowCell Is Nothing And Not lastColCell Is Nothing Then
            lastRow = lastRowCell.row
            lastCol = lastColCell.Column

            With ws.PageSetup
                If StrComp(CStr(s), "Tax Detail", vbTextCompare) = 0 Then
                    .PrintArea = "$A$1:$H$68"
                Else
                    .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
                End If

                .Orientation = xlLandscape
                .Zoom = False
                .FitToPagesWide = 1
                If StrComp(CStr(s), "Report", vbTextCompare) = 0 Then
                    .FitToPagesTall = 1
                Else
                    .FitToPagesTall = False
                End If
                .LeftMargin = Application.InchesToPoints(0.25)
                .RightMargin = Application.InchesToPoints(0.25)
                .TopMargin = Application.InchesToPoints(0.45)
                .BottomMargin = Application.InchesToPoints(0.45)
                .HeaderMargin = Application.InchesToPoints(0.15)
                .FooterMargin = Application.InchesToPoints(0.15)
                .CenterHorizontally = True
                .CenterHeader = "&B" & EscapeExcelHeaderText(GetPropertyValue("Property Address"))
                .RightHeader = EscapeExcelHeaderText(FormatPin(GetPropertyValue("PIN")))
                .CenterFooter = "Public-record information should be independently verified before reliance."
                .RightFooter = "Page &P of &N"
            End With

            If StrComp(CStr(s), "Tax Detail", vbTextCompare) = 0 Then
                On Error Resume Next
                ws.ResetAllPageBreaks
                ws.HPageBreaks.Add Before:=ws.rows(35)
                On Error GoTo 0
            End If
        Else
            ws.PageSetup.PrintArea = ""
        End If

        Set lastRowCell = Nothing
        Set lastColCell = Nothing
    Next s
End Sub

Private Function BuildPdfSheetList() As Variant
    Dim c As New Collection, arr() As Variant, i As Long
    Dim skipEmpty As Boolean, includeSources As Boolean

    skipEmpty = (UCase$(Trim$(CStr(GetConfigValue("Skip Empty Optional PDF Sheets", "YES")))) <> "NO")
    includeSources = (UCase$(Trim$(CStr(GetConfigValue("Include Sources Sheet in PDF", "NO")))) = "YES")

    c.Add "Report"
    c.Add "Property"
    c.Add "Assessment"
    c.Add "Tax History"
    c.Add "Tax Detail"

    AddPdfOptionalSheet c, "Sales-Deeds", skipEmpty
    AddPdfOptionalSheet c, "Documents", skipEmpty
    AddPdfOptionalSheet c, "Appeals", skipEmpty

    ' Incentive/TIF findings are meaningful even when the answer is "No".
    c.Add "Incentives"

    AddPdfOptionalSheet c, "Permits", skipEmpty
    AddPdfOptionalSheet c, "Issues", skipEmpty

    If includeSources Then c.Add "Sources"

    ReDim arr(0 To c.count - 1)
    For i = 1 To c.count
        arr(i - 1) = CStr(c(i))
    Next i

    BuildPdfSheetList = arr
End Function

Private Sub AddPdfOptionalSheet(ByRef c As Collection, ByVal sheetName As String, ByVal skipEmpty As Boolean)
    If Not skipEmpty Then
        c.Add sheetName
    ElseIf SheetHasRunData(sheetName) Then
        c.Add sheetName
    End If
End Sub

Private Function SheetHasRunData(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = LastUsedRowAnyColumn(ws)

    If lastRow < 5 Then
        SheetHasRunData = False
    Else
        SheetHasRunData = (Application.WorksheetFunction.CountA(ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, LastUsedColumn(ws)))) > 0)
    End If
End Function

Private Function EscapeExcelHeaderText(ByVal s As String) As String
    EscapeExcelHeaderText = Replace(s, "&", "&&")
End Function

Private Sub EnsureArcGisResponse(ByVal js As String, ByVal sourceName As String)
    If InStr(1, js, """error""", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 320, , sourceName & " returned an ArcGIS error response."
    End If
    If InStr(1, js, """features""", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 321, , sourceName & " did not return an ArcGIS feature response."
    End If
End Sub

Private Function ArcGisFeatureCount(ByVal js As String) As Long
    Dim compact As String
    compact = Replace(js, " ", "")
    compact = Replace(compact, vbCr, "")
    compact = Replace(compact, vbLf, "")
    compact = Replace(compact, vbTab, "")
    If InStr(1, compact, """features"":[]", vbTextCompare) > 0 Then
        ArcGisFeatureCount = 0
    ElseIf InStr(1, compact, """features"":[{", vbTextCompare) > 0 Then
        ArcGisFeatureCount = 1
    Else
        Err.Raise vbObjectError + 322, , "ArcGIS response contained a features key but its array could not be interpreted."
    End If
End Function

' -----------------------
' HTTP / source helpers
' -----------------------
Private Function SocrataCsv(ByVal datasetId As String, ByVal pin As String, ByVal extraQuery As String) As String
    Dim q As String
    q = "$where=" & UrlEncode("pin='" & pin & "'")
    If Len(extraQuery) > 0 Then q = q & "&" & Replace(extraQuery, " ", "%20")
    SocrataCsv = SOCRATA_BASE & datasetId & ".csv?" & q
End Function

Private Function HttpGetNoThrow(ByVal url As String) As String
    On Error Resume Next
    Err.Clear
    HttpGetNoThrow = HttpGet(url)
    Err.Clear
    On Error GoTo 0
End Function

Private Function HttpGetPortal(ByVal url As String) As String
#If Mac Then
    Dim result As String
    On Error GoTo MacPortalError

    result = AppleScriptTask("CookPropertyHTTP.applescript", "httpGetPortal", url)
    If Left$(result, 10) = "__ERROR__|" Then
        Err.Raise vbObjectError + 303, , Mid$(result, 11)
    End If

    HttpGetPortal = result
    Exit Function

MacPortalError:
    Err.Raise vbObjectError + 304, , "Mac Property Tax Portal request failed: " & Err.Description
#Else
    HttpGetPortal = HttpGet(url)
#End If
End Function

Private Function HttpGet(ByVal url As String) As String
#If Mac Then
    Dim result As String
    On Error GoTo MacHttpError
    result = AppleScriptTask("CookPropertyHTTP.applescript", "httpGet", url)
    If Left$(result, 10) = "__ERROR__|" Then Err.Raise vbObjectError + 301, , Mid$(result, 11)
    HttpGet = result
    Exit Function
MacHttpError:
    Err.Raise vbObjectError + 302, , "Mac HTTP request failed. If the helper is missing, run Install Mac Helper.command once. " & Err.Description
#Else
    Dim http As Object, timeoutMs As Long
    timeoutMs = CLng(GetConfigValue("HTTP Timeout (ms)", 30000))
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 Excel-CookPropertyDueDiligence"
    http.SetRequestHeader "Accept", "*/*"
    http.Send
    If http.status < 200 Or http.status >= 300 Then Err.Raise vbObjectError + 301, , "HTTP " & http.status & " from " & url
    HttpGet = CStr(http.ResponseText)
#End If
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, ch As String, code As Integer, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122
                out = out & ch
            Case 45, 46, 95, 126
                out = out & ch
            Case 32
                out = out & "%20"
            Case Else
                out = out & "%" & Right$("0" & Hex(code And &HFF), 2)
        End Select
    Next i
    UrlEncode = out
End Function

' -----------------------
' CSV parser
' -----------------------
Private Function ParseCsv(ByVal csv As String) As Collection
    Dim result As New Collection, row As Collection
    Dim field As String, inQuotes As Boolean, i As Long, ch As String, nxt As String

    ' AppleScript do shell script converts Unix LF to Mac CR by default unless
    ' "without altering line endings" is specified. Normalize either form so
    ' Socrata CSV remains parseable regardless of helper/runtime behavior.
    csv = Replace(csv, vbCrLf, vbLf)
    csv = Replace(csv, vbCr, vbLf)

    Set row = New Collection
    For i = 1 To Len(csv)
        ch = Mid$(csv, i, 1)
        If ch = """" Then
            If inQuotes And i < Len(csv) Then
                nxt = Mid$(csv, i + 1, 1)
                If nxt = """" Then
                    field = field & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            row.Add field
            field = ""
        ElseIf ch = vbLf And Not inQuotes Then
            row.Add field
            field = ""
            result.Add CollectionToArray(row)
            Set row = New Collection
        ElseIf ch = vbCr And Not inQuotes Then
            ' ignore
        Else
            field = field & ch
        End If
    Next i

    If Len(field) > 0 Or row.count > 0 Then
        row.Add field
        result.Add CollectionToArray(row)
    End If

    Set ParseCsv = result
End Function

Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    CollectionToArray = a
End Function

Private Function CsvRowDict(ByVal headerRow As Variant, ByVal dataRow As Variant) As Collection
    Dim d As New Collection, i As Long, maxI As Long, k As String
    maxI = UBound(headerRow)
    If UBound(dataRow) < maxI Then maxI = UBound(dataRow)
    For i = 0 To maxI
        k = LCase$(Trim$(CStr(headerRow(i))))
        If Len(k) > 0 Then
            On Error Resume Next
            d.Add dataRow(i), k
            On Error GoTo 0
        End If
    Next i
    Set CsvRowDict = d
End Function

Private Function DictGet(ByVal d As Collection, ByVal key As String) As String
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = d(LCase$(Trim$(key)))
    If Err.Number = 0 Then DictGet = Trim$(CStr(v)) Else DictGet = ""
    Err.Clear
    On Error GoTo 0
End Function

Private Function FirstNonBlank(ByVal d As Collection, ByVal keys As Variant) As String
    Dim k As Variant, v As String
    For Each k In keys
        v = DictGet(d, CStr(k))
        If Len(v) > 0 Then
            FirstNonBlank = v
            Exit Function
        End If
    Next k
End Function



Private Sub ParsePortalTaxSaleStatuses(ByVal blockText As String)
    Dim yr As String, seg As String, statusText As String, y As Long
    For y = Year(Date) + 1 To 1999 Step -1
        yr = CStr(y)
        If InStr(1, blockText, yr & ":", vbTextCompare) > 0 Then
            seg = PortalYearSegment(blockText, yr)
            statusText = ""
            If InStr(1, seg, "Tax Sale Has Not Occurred", vbTextCompare) > 0 Or InStr(1, seg, "Tax Sale Has Not Occureed", vbTextCompare) > 0 Then
                statusText = "Tax Sale Has Not Occurred"
            ElseIf InStr(1, seg, "No Tax Sale", vbTextCompare) > 0 Then
                statusText = "No Tax Sale"
            ElseIf InStr(1, seg, "Not Available", vbTextCompare) > 0 Then
                statusText = "Not Available"
            Else
                statusText = CollapseWhitespace(seg)
                If Len(statusText) > 120 Then statusText = Left$(statusText, 120) & "..."
                If Len(statusText) > 0 Then
                    AppendIssue "HIGH", "Tax-sale/delinquency section returned a nonstandard status for " & yr & ".", "Cook County Property Tax Portal", "Review the Portal and Clerk/Treasurer records before relying on tax status. Portal text: " & statusText
                    statusText = "REVIEW - " & statusText
                End If
            End If
            If Len(statusText) > 0 Then UpsertTaxSaleStatus yr, statusText
        End If
    Next y
End Sub

Private Function PortalYearSegment(ByVal blockText As String, ByVal taxYear As String) As String
    Dim p1 As Long, p2 As Long, p As Long, y As Long
    p1 = InStr(1, blockText, taxYear & ":", vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(taxYear) + 1
    p2 = Len(blockText) + 1

    For y = 1999 To 2100
        If CStr(y) <> taxYear Then
            p = InStr(p1, blockText, CStr(y) & ":", vbTextCompare)
            If p > 0 And p < p2 Then p2 = p
        End If
    Next y

    PortalYearSegment = Trim$(Mid$(blockText, p1, p2 - p1))
End Function

Private Sub UpsertTaxSaleStatus(ByVal taxYear As String, ByVal statusText As String)
    Dim ws As Worksheet, r As Long, lastRow As Long, foundRow As Long
    Set ws = ThisWorkbook.Worksheets("Tax History")
    lastRow = LastUsedRow(ws, 1)

    For r = 5 To IIf(lastRow < 5, 5, lastRow)
        If Trim$(CStr(ws.Cells(r, 1).value)) = taxYear Then
            foundRow = r
            Exit For
        End If
    Next r

    If foundRow = 0 Then
        foundRow = lastRow + 1
        If foundRow < 5 Then foundRow = 5
        ws.Cells(foundRow, 1).value = taxYear
    End If

    ws.Cells(foundRow, 8).value = statusText
    ws.Cells(foundRow, 7).value = "Cook County Property Tax Portal"
    UpsertTaxDetailRow taxYear, "", "", "", "", "", statusText
End Sub

Private Sub ParsePortalDocuments(ByVal txt As String)
    Dim blockText As String, seen As New Collection
    Dim pos As Long, docNo As String, docType As String, recDate As String
    blockText = TextBetween(txt, "DOCUMENTS, DEEDS & LIENS", "All years referenced herein")
    If Len(blockText) = 0 Then Exit Sub
    pos = 1
    Do
        If Not NextPortalDocument(blockText, pos, docNo, docType, recDate) Then Exit Do
        If Len(docNo) > 0 And Not CollectionHasKey(seen, docNo) Then
            CollectionAddKey seen, docNo
            AppendDocument docNo, docType, recDate
            If InStr(1, docType, "LIS PENDENS", vbTextCompare) > 0 Or InStr(1, docType, "FORECLOSURE", vbTextCompare) > 0 Then
                AppendIssue "HIGH", "Recent recorded-document index includes " & docType & ".", "Cook County Property Tax Portal Documents / Deeds / Liens", "Obtain and review document " & docNo & " recorded " & recDate & " from the Cook County Clerk/Recorder."
            ElseIf InStr(1, docType, "LIEN", vbTextCompare) > 0 Then
                AppendIssue "MEDIUM", "Recent recorded-document index includes " & docType & ".", "Cook County Property Tax Portal Documents / Deeds / Liens", "Obtain and review document " & docNo & " recorded " & recDate & "."
            End If
        End If
    Loop
End Sub

Private Function NextPortalDocument(ByVal blockText As String, ByRef pos As Long, ByRef docNo As String, ByRef docType As String, ByRef recDate As String) As Boolean
    Dim n As Long, i As Long, j As Long, digitStart As Long, digitEnd As Long
    Dim pDash1 As Long, pDash2 As Long, d As String
    n = Len(blockText): i = pos
    Do While i <= n - 7
        If IsDigitChar(Mid$(blockText, i, 1)) Then
            digitStart = i: j = i
            Do While j <= n And IsDigitChar(Mid$(blockText, j, 1)): j = j + 1: Loop
            digitEnd = j - 1
            If digitEnd - digitStart + 1 >= 8 And digitEnd - digitStart + 1 <= 12 Then
                pDash1 = InStr(digitEnd + 1, blockText, " - ", vbBinaryCompare)
                If pDash1 > 0 And pDash1 - digitEnd <= 5 Then
                    pDash2 = InStr(pDash1 + 3, blockText, " - ", vbBinaryCompare)
                    If pDash2 > 0 Then
                        d = Trim$(Mid$(blockText, pDash2 + 3, 10))
                        If LooksLikeUsDate(d) Then
                            docNo = Mid$(blockText, digitStart, digitEnd - digitStart + 1)
                            docType = Trim$(Mid$(blockText, pDash1 + 3, pDash2 - (pDash1 + 3)))
                            recDate = d
                            pos = pDash2 + 13
                            NextPortalDocument = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop
    pos = n + 1
End Function

Private Function CollectionHasKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = c(LCase$(key))
    CollectionHasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub CollectionAddKey(ByVal c As Collection, ByVal key As String)
    On Error Resume Next
    c.Add True, LCase$(key)
    On Error GoTo 0
End Sub

Private Function LooksLikeUsDate(ByVal s As String) As Boolean
    LooksLikeUsDate = (Len(s) = 10) And (Mid$(s, 3, 1) = "/") And (Mid$(s, 6, 1) = "/") And IsAllDigits(Left$(s, 2)) And IsAllDigits(Mid$(s, 4, 2)) And IsAllDigits(Right$(s, 4))
End Function

Private Function TextBetweenAfter(ByVal s As String, ByVal anchorText As String, ByVal startText As String, ByVal endText As String) As String
    Dim pAnchor As Long, p1 As Long, p2 As Long
    pAnchor = InStr(1, s, anchorText, vbTextCompare)
    If pAnchor = 0 Then Exit Function

    p1 = InStr(pAnchor + Len(anchorText), s, startText, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(startText)

    p2 = InStr(p1, s, endText, vbTextCompare)
    If p2 = 0 Then p2 = Len(s) + 1

    TextBetweenAfter = Mid$(s, p1, p2 - p1)
End Function


Private Function PortalRateHistoryBlock(ByVal txt As String) As String
    Dim p As Long, lastP As Long, searchPos As Long, pEnd As Long

    searchPos = 1
    Do
        p = InStr(searchPos, txt, "Tax Rate History", vbTextCompare)
        If p = 0 Then Exit Do
        lastP = p
        searchPos = p + Len("Tax Rate History")
    Loop

    If lastP = 0 Then Exit Function

    pEnd = InStr(lastP + Len("Tax Rate History"), txt, "Tax Code :", vbTextCompare)
    If pEnd = 0 Then pEnd = InStr(lastP + Len("Tax Rate History"), txt, "Tax Code", vbTextCompare)
    If pEnd = 0 Then pEnd = Len(txt) + 1

    PortalRateHistoryBlock = Mid$(txt, lastP + Len("Tax Rate History"), pEnd - (lastP + Len("Tax Rate History")))
End Function

Private Function PortalPageLooksValid(ByVal txt As String, ByVal pin As String) As Boolean
    Dim hasPin As Boolean, hasCharacteristics As Boolean, hasTaxHistory As Boolean

    hasPin = (InStr(1, txt, FormatPin(pin), vbTextCompare) > 0)
    hasCharacteristics = _
        (InStr(1, txt, "PROPERTY CHARACTERISTICS", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Property Characteristics for PIN", vbTextCompare) > 0)
    hasTaxHistory = (InStr(1, txt, "TAX BILLED AMOUNTS", vbTextCompare) > 0)

    PortalPageLooksValid = hasPin And hasCharacteristics And hasTaxHistory
End Function

Private Function LooseNumberAfterYear(ByVal blockText As String, ByVal taxYear As String) As String
    Dim searchPos As Long, p As Long, i As Long, ch As String, out As String

    searchPos = 1

    Do
        p = InStr(searchPos, blockText, taxYear, vbTextCompare)
        If p = 0 Then Exit Function

        i = p + Len(taxYear)

        Do While i <= Len(blockText)
            ch = Mid$(blockText, i, 1)

            If IsDigitChar(ch) Then Exit Do

            If ch = "$" Or ch = "+" Or ch = "-" Or ch = "." Or ch = "," Or _
               ch = ":" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or AscW(ch) = 160 Then
                i = i + 1
            Else
                searchPos = p + Len(taxYear)
                GoTo TryNextYearOccurrence
            End If
        Loop

        Do While i <= Len(blockText)
            ch = Mid$(blockText, i, 1)
            If IsDigitChar(ch) Or ch = "." Or ch = "," Then
                out = out & ch
                i = i + 1
            Else
                Exit Do
            End If
        Loop

        If Len(out) > 0 Then
            LooseNumberAfterYear = Trim$(out)
            Exit Function
        End If

        searchPos = p + Len(taxYear)
TryNextYearOccurrence:
    Loop
End Function

Private Sub UpsertTaxDetailRow(ByVal taxYear As String, ByVal rateText As String, ByVal taxCode As String, _
                               ByVal billed As String, ByVal statusText As String, ByVal exemptions As String, _
                               ByVal taxSaleStatus As String)
    Dim ws As Worksheet, r As Long, foundRow As Long

    Set ws = ThisWorkbook.Worksheets("Tax Detail")

    For r = 9 To 15
        If Trim$(CStr(ws.Cells(r, 1).value)) = taxYear Then
            foundRow = r
            Exit For
        End If
    Next r

    If foundRow = 0 Then
        For r = 9 To 15
            If Len(Trim$(CStr(ws.Cells(r, 1).value))) = 0 Then
                foundRow = r
                ws.Cells(r, 1).value = taxYear
                Exit For
            End If
        Next r
    End If

    If foundRow = 0 Then Exit Sub

    If Len(rateText) > 0 Then ws.Cells(foundRow, 2).value = rateText

    If Len(taxCode) > 0 Then
        ws.Cells(foundRow, 3).NumberFormat = "@"
        ws.Cells(foundRow, 3).value = taxCode
    End If

    If Len(billed) > 0 Then ws.Cells(foundRow, 4).value = "$" & billed
    If Len(statusText) > 0 Then ws.Cells(foundRow, 5).value = statusText
    If Len(exemptions) > 0 Then ws.Cells(foundRow, 6).value = exemptions
    If Len(taxSaleStatus) > 0 Then ws.Cells(foundRow, 7).value = taxSaleStatus
End Sub

Private Sub CrossCheckPortalSize(ByVal fieldName As String, ByVal portalValue As String)
    Dim existingValue As String

    If Len(Trim$(portalValue)) = 0 Then Exit Sub

    existingValue = GetPropertyValue(fieldName)

    If Len(existingValue) = 0 Then
        SetProperty fieldName, portalValue, "Property Tax Portal", "Live", "CHECK", _
                    "Portal value used because CookViewer value was blank."
    ElseIf MaterialNumericDifference(existingValue, portalValue) Then
        AppendIssue "MEDIUM", fieldName & " differs between CookViewer and the Property Tax Portal.", _
                    "CookViewer / Property Tax Portal", _
                    "CookViewer: " & existingValue & "; Portal: " & portalValue & ". Verify before relying on the measurement."
    End If
End Sub

Private Function MaterialNumericDifference(ByVal leftText As String, ByVal rightText As String) As Boolean
    Dim a As Double, b As Double, diff As Double, larger As Double
    Dim aText As String, bText As String

    aText = Replace(Replace(Trim$(leftText), ",", ""), "$", "")
    bText = Replace(Replace(Trim$(rightText), ",", ""), "$", "")

    If Not IsNumeric(aText) Or Not IsNumeric(bText) Then
        MaterialNumericDifference = (NormalizeNumberText(leftText) <> NormalizeNumberText(rightText))
        Exit Function
    End If

    a = CDbl(aText)
    b = CDbl(bText)
    diff = Abs(a - b)

    ' Ignore trivial measurement/rounding differences.
    If diff <= 2 Then
        MaterialNumericDifference = False
        Exit Function
    End If

    If Abs(a) > Abs(b) Then
        larger = Abs(a)
    Else
        larger = Abs(b)
    End If

    If larger > 0 Then
        If diff / larger <= 0.0025 Then
            MaterialNumericDifference = False
            Exit Function
        End If
    End If

    MaterialNumericDifference = True
End Function

Private Function PortalExemptionForYear(ByVal blockText As String, ByVal taxYear As String) As String
    Dim p1 As Long, p2 As Long, p As Long, y As Long
    Dim seg As String, out As String, countText As String

    p1 = InStr(1, blockText, taxYear & ":", vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(taxYear) + 1
    p2 = Len(blockText) + 1

    For y = 1999 To 2100
        If CStr(y) <> taxYear Then
            p = InStr(p1, blockText, CStr(y) & ":", vbTextCompare)
            If p > 0 And p < p2 Then p2 = p
        End If
    Next y

    seg = Trim$(Mid$(blockText, p1, p2 - p1))
    If InStr(1, seg, "Not Available", vbTextCompare) > 0 Then
        PortalExemptionForYear = "Not Available"
        Exit Function
    End If

    countText = NumberBeforePhrase(seg, "Exemptions Received")
    If Len(countText) > 0 Then out = countText & " exemption(s) received"

    AppendIfContains out, seg, "Homeowner Exemption"
    AppendIfContains out, seg, "Senior Citizen Exemption"
    AppendIfContains out, seg, "Senior Exemption"
    AppendIfContains out, seg, "Senior Assessment Freeze Exemption"
    AppendIfContains out, seg, "Senior Freeze Exemption"
    AppendIfContains out, seg, "Disabled Persons Exemption"
    AppendIfContains out, seg, "Disabled Veterans"
    AppendIfContains out, seg, "Returning Veterans"
    AppendIfContains out, seg, "Long-Time Occupant"
    AppendIfContains out, seg, "Home Improvement Exemption"
    AppendIfContains out, seg, "Certificate of Error Applied"

    PortalExemptionForYear = out
End Function

Private Function NumberBeforePhrase(ByVal s As String, ByVal phrase As String) As String
    Dim p As Long, leftText As String, i As Long, ch As String, out As String
    p = InStr(1, s, phrase, vbTextCompare)
    If p = 0 Then Exit Function
    leftText = RTrim$(Left$(s, p - 1))
    For i = Len(leftText) To 1 Step -1
        ch = Mid$(leftText, i, 1)
        If IsDigitChar(ch) Then
            out = ch & out
        ElseIf Len(out) > 0 Then
            Exit For
        End If
    Next i
    NumberBeforePhrase = out
End Function

Private Sub AppendIfContains(ByRef outText As String, ByVal haystack As String, ByVal needle As String)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
        If Len(outText) > 0 Then outText = outText & "; "
        outText = outText & needle
    End If
End Sub

Private Function NormalizeNumberText(ByVal s As String) As String
    Dim t As String
    t = Replace(Trim$(s), ",", "")
    t = Replace(t, "$", "")
    If IsNumeric(t) Then
        NormalizeNumberText = Format$(CDbl(t), "0.########")
    Else
        NormalizeNumberText = t
    End If
End Function

Private Function CleanPortalText(ByVal s As String) As String
    Dim t As String, p As Long
    t = Trim$(s)

    p = InStrRev(t, "Property Class Description", -1, vbTextCompare)
    If p > 0 Then t = Trim$(Mid$(t, p + Len("Property Class Description")))

    t = Replace(t, ChrW(215), "")
    CleanPortalText = CollapseWhitespace(t)
End Function

Private Function JsonFirstNonBlank(ByVal js As String, ByVal keys As Variant) As String
    Dim k As Variant, v As String
    For Each k In keys
        v = JsonScalar(js, CStr(k))
        If Len(v) > 0 Then
            JsonFirstNonBlank = v
            Exit Function
        End If
    Next k
End Function

' -----------------------
' HTML / JSON lightweight parsing
' -----------------------
Private Function HtmlToText(ByVal html As String) As String
    Dim s As String, parts As Variant, i As Long, p As Long, out As String
    s = RemoveHtmlBlock(html, "<script", "</script>")
    s = RemoveHtmlBlock(s, "<style", "</style>")

    parts = Split(s, "<")
    out = CStr(parts(0))
    For i = 1 To UBound(parts)
        p = InStr(1, CStr(parts(i)), ">", vbBinaryCompare)
        If p > 0 Then out = out & " " & Mid$(CStr(parts(i)), p + 1)
    Next i

    out = Replace(out, "&nbsp;", " ")
    out = Replace(out, "&#160;", " ")
    out = Replace(out, "&amp;", "&")
    out = Replace(out, "&#39;", "'")
    out = Replace(out, "&quot;", """")
    out = Replace(out, "&lt;", "<")
    out = Replace(out, "&gt;", ">")
    HtmlToText = out
End Function

Private Function RemoveHtmlBlock(ByVal s As String, ByVal startMarker As String, ByVal endMarker As String) As String
    Dim p1 As Long, p2 As Long
    Do
        p1 = InStr(1, s, startMarker, vbTextCompare)
        If p1 = 0 Then Exit Do
        p2 = InStr(p1, s, endMarker, vbTextCompare)
        If p2 = 0 Then s = Left$(s, p1 - 1): Exit Do
        s = Left$(s, p1 - 1) & " " & Mid$(s, p2 + Len(endMarker))
    Loop
    RemoveHtmlBlock = s
End Function

Private Function CollapseWhitespace(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    CollapseWhitespace = Trim$(s)
End Function

Private Function ArcGisFirstAttributesJson(ByVal js As String) As String
    Dim p As Long, openPos As Long, i As Long, depth As Long, ch As String
    Dim insideString As Boolean, escapedChar As Boolean

    p = VBA.InStr(1, js, """attributes""", vbTextCompare)
    If p = 0 Then
        ArcGisFirstAttributesJson = js
        Exit Function
    End If

    openPos = VBA.InStr(p, js, "{", vbBinaryCompare)
    If openPos = 0 Then
        ArcGisFirstAttributesJson = js
        Exit Function
    End If

    depth = 0
    insideString = False
    escapedChar = False

    For i = openPos To Len(js)
        ch = Mid$(js, i, 1)

        If insideString Then
            If escapedChar Then
                escapedChar = False
            ElseIf AscW(ch) = 92 Then
                escapedChar = True
            ElseIf AscW(ch) = 34 Then
                insideString = False
            End If
        Else
            If AscW(ch) = 34 Then
                insideString = True
            ElseIf ch = "{" Then
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1

                If depth = 0 Then
                    ArcGisFirstAttributesJson = Mid$(js, openPos, i - openPos + 1)
                    Exit Function
                End If
            End If
        End If
    Next i

    ArcGisFirstAttributesJson = js
End Function

Private Function JsonScalar(ByVal js As String, ByVal key As String) As String
    Dim p As Long, colonPos As Long, i As Long, ch As String, out As String
    js = ArcGisFirstAttributesJson(js)
    p = InStr(1, js, """" & key & """", vbTextCompare)
    If p = 0 Then Exit Function
    colonPos = InStr(p + Len(key) + 2, js, ":", vbBinaryCompare)
    If colonPos = 0 Then Exit Function
    i = colonPos + 1
    Do While i <= Len(js) And IsWhitespaceChar(Mid$(js, i, 1)): i = i + 1: Loop
    If i > Len(js) Then Exit Function
    If Mid$(js, i, 1) = """" Then
        i = i + 1
        Do While i <= Len(js)
            ch = Mid$(js, i, 1)
            If ch = "\" And i < Len(js) Then
                i = i + 1: ch = Mid$(js, i, 1)
                Select Case ch
                    Case """": out = out & """"
                    Case "\": out = out & "\"
                    Case "n", "r", "t": out = out & " "
                    Case Else: out = out & ch
                End Select
            ElseIf ch = """" Then
                Exit Do
            Else
                out = out & ch
            End If
            i = i + 1
        Loop
        JsonScalar = out
    Else
        Do While i <= Len(js)
            ch = Mid$(js, i, 1)
            If ch = "," Or ch = "}" Or ch = "]" Or IsWhitespaceChar(ch) Then Exit Do
            out = out & ch: i = i + 1
        Loop
        If LCase$(Trim$(out)) <> "null" Then JsonScalar = Trim$(out)
    End If
End Function

Private Function ValueAfterLabel(ByVal s As String, ByVal labelText As String, ByVal allowedChars As String) As String
    Dim searchPos As Long, p As Long, i As Long, ch As String, out As String
    searchPos = 1

    Do
        p = InStr(searchPos, s, labelText, vbTextCompare)
        If p = 0 Then Exit Function

        i = p + Len(labelText)
        Do While i <= Len(s) And (Mid$(s, i, 1) = " " Or Mid$(s, i, 1) = vbTab Or Mid$(s, i, 1) = vbCr Or Mid$(s, i, 1) = vbLf)
            i = i + 1
        Loop

        If i <= Len(s) And Mid$(s, i, 1) = ":" Then
            i = i + 1
            Do While i <= Len(s) And (Mid$(s, i, 1) = " " Or Mid$(s, i, 1) = vbTab Or Mid$(s, i, 1) = vbCr Or Mid$(s, i, 1) = vbLf)
                i = i + 1
            Loop
        ElseIf i > Len(s) Or InStr(1, allowedChars, Mid$(s, i, 1), vbBinaryCompare) = 0 Then
            searchPos = p + Len(labelText)
            GoTo TryNextOccurrence
        End If

        Do While i <= Len(s)
            ch = Mid$(s, i, 1)
            If InStr(1, allowedChars, ch, vbBinaryCompare) = 0 Then Exit Do
            out = out & ch
            i = i + 1
        Loop

        If Len(out) > 0 Then
            ValueAfterLabel = Trim$(out)
            Exit Function
        End If

        searchPos = p + Len(labelText)
TryNextOccurrence:
    Loop
End Function

Private Function FirstNumericToken(ByVal s As String) As String
    Dim i As Long, ch As String, started As Boolean, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsDigitChar(ch) Then
            started = True: out = out & ch
        ElseIf started And (ch = "," Or ch = ".") Then
            out = out & ch
        ElseIf started Then
            Exit For
        End If
    Next i
    FirstNumericToken = out
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    IsDigitChar = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not IsDigitChar(Mid$(s, i, 1)) Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function IsWhitespaceChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then
        IsWhitespaceChar = False
    Else
        IsWhitespaceChar = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or AscW(ch) = 160)
    End If
End Function

Private Function TextBetween(ByVal s As String, ByVal startText As String, ByVal endText As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, s, startText, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(startText)
    p2 = InStr(p1, s, endText, vbTextCompare)
    If p2 = 0 Then p2 = Len(s) + 1
    TextBetween = Mid$(s, p1, p2 - p1)
End Function

' -----------------------
' General helpers
' -----------------------
Private Function NormalizePin(ByVal rawPin As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(rawPin)
        ch = Mid$(rawPin, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    NormalizePin = out
End Function

Private Function ReadPinFromStartSheet(ByRef inputAddress As String) As String
    Dim ws As Worksheet, c As Range, candidate As String
    Dim bestCandidate As String, bestAddress As String

    Set ws = ThisWorkbook.Worksheets("Start")

    ' Scan the visible PIN-entry row instead of assuming one physical cell.
    ' This is deliberately tolerant of Mac Excel merge/active-cell behavior.
    For Each c In ws.Range("A4:H4").Cells
        candidate = NormalizePin(CStr(c.Value2))

        If Len(candidate) = 14 Then
            On Error Resume Next
            bestAddress = c.MergeArea.Cells(1, 1).Address(False, False)
            If Err.Number <> 0 Or Len(bestAddress) = 0 Then
                Err.Clear
                bestAddress = c.Address(False, False)
            End If
            On Error GoTo 0

            inputAddress = bestAddress
            ReadPinFromStartSheet = candidate
            Exit Function
        End If

        If Len(candidate) > Len(bestCandidate) Then
            bestCandidate = candidate
            bestAddress = c.Address(False, False)
        End If
    Next c

    ' Fallback keeps the best numeric candidate so the validation message
    ' still behaves predictably if the user entered an incomplete PIN.
    inputAddress = bestAddress
    ReadPinFromStartSheet = bestCandidate
End Function

Private Function FormatPin(ByVal pin As String) As String
    If Len(pin) = 14 Then
        FormatPin = Left$(pin, 2) & "-" & Mid$(pin, 3, 2) & "-" & Mid$(pin, 5, 3) & "-" & Mid$(pin, 8, 3) & "-" & Right$(pin, 4)
    Else
        FormatPin = pin
    End If
End Function

Private Function GetConfigValue(ByVal settingName As String, ByVal defaultValue As Variant) As Variant
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Config")
    For r = 5 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), settingName, vbTextCompare) = 0 Then
            If Len(Trim$(CStr(ws.Cells(r, 2).value))) > 0 Then
                GetConfigValue = ws.Cells(r, 2).value
            Else
                GetConfigValue = defaultValue
            End If
            Exit Function
        End If
    Next r
    GetConfigValue = defaultValue
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

Private Function NormalizeMunicipality(ByVal s As String) As String
    Dim t As String

    t = UCase$(Trim$(s))

    If Left$(t, 8) = "CITY OF " Then t = Mid$(t, 9)
    If Left$(t, 11) = "VILLAGE OF " Then t = Mid$(t, 12)
    If Left$(t, 8) = "TOWN OF " Then t = Mid$(t, 9)

    t = Replace(t, ".", "")
    t = Replace(t, ",", "")
    t = Replace(t, "-", "")
    t = Replace(t, " ", "")

    NormalizeMunicipality = t
End Function

Private Function NormalizeCompare(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, "-", "")
    t = Replace(t, " ", "")
    NormalizeCompare = t
End Function

Private Function FirstText(ByVal a As String, ByVal b As String) As String
    If Len(Trim$(a)) > 0 Then FirstText = a Else FirstText = b
End Function

' -----------------------
' Developer / regression tests
' -----------------------
Public Sub TestAdapters()
    Dim pin As String, x As String, y As String
    pin = "16302040200000"

    Debug.Print "Testing PIN " & FormatPin(pin)
    Debug.Print "Normalize PIN test: " & NormalizePin("16-30-204-020-0000")
    Debug.Print "Address URL: " & SocrataCsv(ADDRESS_DATASET, pin, "$order=year DESC&$limit=1")
    Debug.Print "Universe URL: " & SocrataCsv(UNIVERSE_DATASET, pin, "$limit=1")
    Debug.Print "Universe EZ fields expected: econ_enterprise_zone_num, econ_enterprise_zone_data_year"
    Debug.Print "Assessment URL: " & SocrataCsv(ASSESSMENT_DATASET, pin, "$order=year DESC&$limit=3")
    Debug.Print "Sales URL: " & SocrataCsv(SALES_DATASET, pin, "$order=sale_date DESC&$limit=3")
    Debug.Print "Assessor Appeals URL: " & SocrataCsv(APPEALS_DATASET, pin, "$order=year DESC&$limit=3")
    Debug.Print "BOR URL: " & SocrataCsv(BOR_DATASET, pin, "$order=tax_year DESC&$limit=3")
    Debug.Print "Permits URL: " & SocrataCsv(PERMITS_DATASET, pin, "$order=date_issued DESC&$limit=3")
    Debug.Print "CookViewer current PIN URL: " & CStr(GetConfigValue("GIS Current Parcel URL", "")) & _
                "/query?where=" & UrlEncode("PIN14='" & pin & "'") & "&outFields=*&returnGeometry=false&f=pjson"
    Debug.Print "Tax Portal: " & CStr(GetConfigValue("Property Tax Portal Results URL", "https://www.cookcountypropertyinfo.com/PINResults.aspx")) & "?pin=" & pin

    MsgBox "Adapter smoke-test URLs were printed to the Immediate window. Run GenerateCookPropertyReport for the end-to-end workbook/PDF test.", vbInformation
End Sub


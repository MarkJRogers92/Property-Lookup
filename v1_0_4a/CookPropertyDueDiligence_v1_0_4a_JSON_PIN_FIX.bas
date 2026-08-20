Attribute VB_Name = "Module1"
Option Explicit

' ============================================================================
' COOK COUNTY PROPERTY DUE DILIGENCE v1.0.4a - MAC + WINDOWS EXCEL
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


Public Sub Auto_Open()
    On Error Resume Next
    InstallRunButton
    UpdatePlatformReadyStatus
    On Error GoTo 0
End Sub

Public Sub InstallRunButton()
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
        .OnAction = "'" & ThisWorkbook.Name & "'!GenerateCookPropertyReport"
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
        ThisWorkbook.Worksheets("Start").Range("B8").value = "Mac helper missing - run Install Mac Helper.command once"
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
               "Double-click 'Install Mac Helper.command' from the v1.0.3 package, then reopen this workbook.", _
               vbExclamation, "Cook County Property Due Diligence"
        ThisWorkbook.Worksheets("Start").Range("B8").value = "Mac helper missing - run installer once"
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
    result = AppleScriptTask("CookPropertyHTTP.applescript", "selfTest", "")
    MacHelperReady = (UCase$(Trim$(result)) = "OK")
    Exit Function
Missing:
    Err.Clear
    MacHelperReady = False
#Else
    MacHelperReady = True
#End If
End Function

Public Sub GenerateCookPropertyReport()
    Dim pin As String, pinInputAddress As String
    Dim t0 As Double
    t0 = Timer

    On Error GoTo FatalError

    If Not PlatformReady() Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    pin = ReadPinFromStartSheet(pinInputAddress)
    If Len(pin) <> 14 Then
        MsgBox "Enter a valid 14-digit Cook County PIN in the yellow PIN field.", vbExclamation
        GoTo SafeExit
    End If

    If Len(pinInputAddress) = 0 Then pinInputAddress = "B4"
    ThisWorkbook.Worksheets("Start").Range(pinInputAddress).value = FormatPin(pin)
    ResetRunSheets
    SetRunStatus "Starting research for " & FormatPin(pin)

    SafeAdapter "Assessor Addresses", "FetchAddress", pin
    SafeAdapter "Parcel Universe", "FetchUniverse", pin
    SafeAdapter "Assessed Values", "FetchAssessments", pin
    SafeAdapter "Parcel Sales", "FetchSales", pin
    SafeAdapter "Assessor Appeals", "FetchAssessorAppeals", pin
    SafeAdapter "Board of Review", "FetchBORAppeals", pin
    SafeAdapter "Permits", "FetchPermits", pin
    SafeAdapter "Cook County GIS", "FetchGISParcel", pin
    SafeAdapter "Property Tax Portal", "FetchTaxPortal", pin
    SafeAdapter "TIF GIS", "FetchTIF", pin
    SafeAdapter "Enterprise Zone", "FetchEnterpriseZone", pin

    BuildIssues
    BuildReport
    ExportDueDiligencePDF pin

    SetRunStatus "Complete in " & Format(Timer - t0, "0.0") & " seconds"

SafeExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub

FatalError:
    SetRunStatus "Fatal error: " & Err.Description
    MsgBox "The run stopped unexpectedly: " & Err.Description, vbCritical
    Resume SafeExit
End Sub

Private Sub SafeAdapter(ByVal sourceName As String, ByVal procName As String, ByVal pin As String)
    On Error GoTo EH

    UpdateSourceStatus sourceName, "Running", "Starting", ""
    Select Case procName
        Case "FetchAddress": FetchAddress pin
        Case "FetchUniverse": FetchUniverse pin
        Case "FetchAssessments": FetchAssessments pin
        Case "FetchSales": FetchSales pin
        Case "FetchAssessorAppeals": FetchAssessorAppeals pin
        Case "FetchBORAppeals": FetchBORAppeals pin
        Case "FetchPermits": FetchPermits pin
        Case "FetchGISParcel": FetchGISParcel pin
        Case "FetchTaxPortal": FetchTaxPortal pin
        Case "FetchTIF": FetchTIF pin
        Case "FetchEnterpriseZone": FetchEnterpriseZone pin
    End Select

    If StrComp(sourceName, "Enterprise Zone", vbTextCompare) = 0 Then
        UpdateSourceStatus sourceName, "PARTIAL", "Finished - boundary verification remains manual", _
                           "Cook County Assessor signal collected automatically; final DCEO boundary verification remains manual."
    Else
        UpdateSourceStatus sourceName, "OK", "Finished", ""
    End If
    Exit Sub

EH:
    UpdateSourceStatus sourceName, "FAILED", "Stopped", Err.Description
    AppendIssue "HIGH", sourceName & " could not be verified automatically.", sourceName, _
                "Open the official source link on the Start/Sources sheet and verify manually. Error: " & Err.Description
    Err.Clear
End Sub

Private Sub FetchAddress(ByVal pin As String)
    Dim url As String, csv As String, rows As Collection, d As Collection
    url = SocrataCsv(ADDRESS_DATASET, pin, "$order=year DESC&$limit=1")
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)
    If rows.Count < 2 Then Err.Raise vbObjectError + 101, , "No address record returned."
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

    url = SocrataCsv(UNIVERSE_DATASET, pin, "$limit=1")
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)
    If rows.Count < 2 Then Err.Raise vbObjectError + 102, , "No current Parcel Universe row returned."
    Set d = CsvRowDict(rows(1), rows(2))

    SetProperty "Township", DictGet(d, "township_name"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Assessor Class", DictGet(d, "class"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Assessor Neighborhood", DictGet(d, "nbhd_code"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Tax Code", DictGet(d, "tax_code"), "Parcel Universe", Nz(DictGet(d, "year")), "CAUTION", "County metadata states this tax-code field is not currently up-to-date; Property Tax Portal takes precedence."
    SetProperty "Longitude", DictGet(d, "lon"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Latitude", DictGet(d, "lat"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Centroid X (3435)", DictGet(d, "x_3435"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Centroid Y (3435)", DictGet(d, "y_3435"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""

    ezNum = Trim$(DictGet(d, "econ_enterprise_zone_num"))
    ezYear = Trim$(DictGet(d, "econ_enterprise_zone_data_year"))

    If Len(ezNum) > 0 And ezNum <> "0" Then
        SetProperty "Enterprise Zone (Assessor Signal)", _
                    "Signal present: Zone " & ezNum, _
                    "Assessor Parcel Universe", FirstText(ezYear, Nz(DictGet(d, "year"))), "CHECK", _
                    "Cook County Assessor spatial signal only; verify the parcel in the official Illinois DCEO Enterprise Zone map before reliance."
    Else
        SetProperty "Enterprise Zone (Assessor Signal)", _
                    "No county signal returned", _
                    "Assessor Parcel Universe", FirstText(ezYear, Nz(DictGet(d, "year"))), "CHECK", _
                    "Absence of the county signal is not treated as a conclusive outside-zone determination; verify in the official Illinois DCEO map."
    End If

End Sub

Private Sub FetchAssessments(ByVal pin As String)
    Dim n As Long, url As String, csv As String
    n = CLng(GetConfigValue("Assessment Years", 8))
    url = SocrataCsv(ASSESSMENT_DATASET, pin, "$order=year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    WriteSelectedCsv "Assessment", csv, 5, _
        Array("year", "class", "mailed_land", "mailed_bldg", "mailed_tot", "certified_land", "certified_bldg", "certified_tot", "board_land", "board_bldg", "board_tot")
End Sub

Private Sub FetchSales(ByVal pin As String)
    Dim n As Long, url As String, csv As String
    n = CLng(GetConfigValue("Sales Limit", 20))
    url = SocrataCsv(SALES_DATASET, pin, "$order=sale_date DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
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
    If rows.Count < 2 Then Exit Sub

    For i = 2 To rows.Count
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
    If rows.Count < 2 Then Exit Sub

    For i = 2 To rows.Count
        Set d = CsvRowDict(rows(1), rows(i))
        AppendAppeal "Board of Review", DictGet(d, "tax_year"), _
                     Nz(DictGet(d, "appealtrk")) & "-" & Nz(DictGet(d, "appealseq")), _
                     FirstNonBlank(d, Array("appealtypedescription", "appealtype")), DictGet(d, "result"), _
                     DictGet(d, "assessor_totalvalue"), DictGet(d, "bor_totalvalue"), "BOR Decision History"
    Next i
End Sub

Private Sub FetchPermits(ByVal pin As String)
    Dim n As Long, url As String, csv As String
    n = CLng(GetConfigValue("Permits Limit", 50))
    url = SocrataCsv(PERMITS_DATASET, pin, "$order=date_issued DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
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

    portalBase = CStr(GetConfigValue("Property Tax Portal Results URL", "https://www.cookcountypropertyinfo.com/PINResults.aspx"))
    portalFallback = "https://www.cookcountypropertyinfo.com/cookviewerpinresults.aspx"

    url = portalBase & "?pin=" & pin
    On Error Resume Next
    html = HttpGet(url)
    Err.Clear
    On Error GoTo 0
    txt = CollapseWhitespace(HtmlToText(html))

    If Not PortalPageLooksValid(txt, pin) Then
        url = portalFallback & "?pin=" & pin
        html = HttpGet(url)
        txt = CollapseWhitespace(HtmlToText(html))
    End If
    If Not PortalPageLooksValid(txt, pin) Then Err.Raise vbObjectError + 160, , "Property Tax Portal did not return a PIN-specific results page for " & FormatPin(pin) & "."

    rate = ValueAfterLabel(txt, "Tax Rate", "0123456789.")
    taxCode = ValueAfterLabel(txt, "Tax Code", "0123456789")
    portalClass = ValueAfterLabel(txt, "Property Class", "0123456789-")
    portalLot = ValueAfterLabel(txt, "Lot Size (SqFt)", "0123456789,.")
    portalBldg = ValueAfterLabel(txt, "Building (SqFt)", "0123456789,.")
    portalClassDesc = CleanPortalText(TextBetween(txt, "Property Class Description", "Tax Rate"))

    If Len(rate) > 0 Then ThisWorkbook.Worksheets("Report").Range("E7").value = rate
    If Len(taxCode) > 0 Then SetProperty "Tax Code", taxCode, "Property Tax Portal", "Live", "OK", "Portal tax code used for current tax-rate context."

    If Len(portalClass) > 0 And Len(GetPropertyValue("Assessor Class")) > 0 Then
        If NormalizeCompare(portalClass) <> NormalizeCompare(GetPropertyValue("Assessor Class")) Then
            AppendIssue "MEDIUM", "Property Tax Portal class differs from Parcel Universe class.", "Property Tax Portal / Parcel Universe", _
                        "Portal: " & portalClass & "; Parcel Universe: " & GetPropertyValue("Assessor Class") & ". Verify current classification."
        End If
    End If
    If Len(portalClassDesc) > 0 And Len(GetPropertyValue("GIS Class Description")) = 0 Then SetProperty "GIS Class Description", portalClassDesc, "Property Tax Portal", "Live", "CHECK", "Portal class description used because GIS description was blank."
    CrossCheckPortalSize "Lot Size (SqFt)", portalLot
    CrossCheckPortalSize "Building (SqFt)", portalBldg

    rateBlock = TextBetween(txt, "Tax Rate History", "Tax Code")
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
                If Len(exText) > 0 Then UpsertTaxHistory yr, "", "", "", taxCode, exText
            End If
        Next r
    End If

    refundBlock = TextBetween(txt, "REFUNDS AVAILABLE", "TAX SALE (DELINQUENCIES)")
    If Len(refundBlock) > 0 And InStr(1, refundBlock, "No Refund Available", vbTextCompare) = 0 Then AppendIssue "LOW", "Property Tax Portal does not show the standard 'No Refund Available' message.", "Cook County Property Tax Portal", "Review the Portal refund section; an overpayment/refund may exist or the Portal wording may have changed."
    taxSaleBlock = TextBetween(txt, "TAX SALE (DELINQUENCIES)", "DOCUMENTS, DEEDS & LIENS")
    If Len(taxSaleBlock) > 0 Then ParsePortalTaxSaleStatuses taxSaleBlock
    ParsePortalDocuments txt
End Sub

Private Sub ParsePortalRateHistory(ByVal rateBlock As String, ByVal taxCode As String)
    Dim y As Long, rateText As String
    If Len(rateBlock) = 0 Then Exit Sub
    For y = Year(Date) + 1 To 1999 Step -1
        If InStr(1, rateBlock, CStr(y), vbTextCompare) > 0 Then
            rateText = ValueAfterLabel(rateBlock, CStr(y), "0123456789.")
            If Len(rateText) > 0 Then UpsertTaxHistory CStr(y), "", "", rateText, taxCode, ""
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
                If firstBill Then
                    If isInstallment Then
                        ThisWorkbook.Worksheets("Report").Range("E9").value = "$" & billed & " (1st installment)"
                    Else
                        ThisWorkbook.Worksheets("Report").Range("E9").value = "$" & billed
                    End If
                    firstBill = False
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
    mTax = NormalizeCompare(GetPropertyValue("Municipality - Tax Record"))
    mSpatial = NormalizeCompare(GetPropertyValue("Municipality - Spatial"))

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
    Dim folder As String, fn As String, fullPath As String, openAfter As Boolean
    Dim arr As Variant

    folder = CStr(GetConfigValue("PDF Output Folder", ""))
    If Len(Trim$(folder)) = 0 Then folder = ThisWorkbook.Path
    If Len(Trim$(folder)) = 0 Then folder = DefaultDocumentsFolder()
    folder = EnsureTrailingPathSeparator(folder)

    fn = "Cook_Property_Due_Diligence_" & pin & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    fullPath = folder & fn

    ConfigurePrintAreas
    arr = Array("Report", "Property", "Assessment", "Tax History", "Sales-Deeds", "Documents", "Appeals", "Incentives", "Permits", "Issues", "Sources")
    ThisWorkbook.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ThisWorkbook.Worksheets("Start").Select

    openAfter = (UCase$(CStr(GetConfigValue("Open PDF After Export", "YES"))) = "YES")
    If openAfter Then OpenGeneratedFile fullPath
End Sub

Private Function EnsureTrailingPathSeparator(ByVal folder As String) As String
#If Mac Then
    folder = AppleScriptTask("CookPropertyHTTP.applescript", "normalizeFolder", folder)
    If Right$(folder, 1) <> "/" Then folder = folder & "/"
#Else
    If Right$(folder, 1) <> "\\" Then folder = folder & "\\"
#End If
    EnsureTrailingPathSeparator = folder
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
    Dim result As String
    result = AppleScriptTask("CookPropertyHTTP.applescript", "openFile", fullPath)
    If Left$(result, 10) = "__ERROR__|" Then Err.Raise vbObjectError + 391, , "PDF was created but could not be opened automatically: " & result
#Else
    ThisWorkbook.FollowHyperlink fullPath
#End If
End Sub

' -----------------------
' Workbook writing helpers
' -----------------------
Private Sub ResetRunSheets()
    Dim s As Variant, ws As Worksheet, r As Long
    For Each s In Array("Assessment", "Tax History", "Sales-Deeds", "Documents", "Appeals", "Incentives", "Permits", "Issues")
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        ws.rows("5:" & ws.rows.Count).ClearContents
    Next s

    Set ws = ThisWorkbook.Worksheets("Property")
    ws.Range("B5:F40").ClearContents
    ThisWorkbook.Worksheets("Report").Range("B5:B12,E5:E12,A14").ClearContents

    For r = 12 To 22
        ThisWorkbook.Worksheets("Start").Cells(r, 2).value = "Not run"
        ThisWorkbook.Worksheets("Start").Cells(r, 3).ClearContents
        ThisWorkbook.Worksheets("Start").Cells(r, 4).ClearContents
    Next r
End Sub

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
            ws.Cells(r, 2).value = value
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
    If Len(taxCode) > 0 Then ws.Cells(foundRow, 5).value = taxCode
    If Len(exemptionsText) > 0 Then ws.Cells(foundRow, 6).value = exemptionsText
    ws.Cells(foundRow, 7).value = "Cook County Property Tax Portal"
End Sub

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
    Set rows = ParseCsv(csv)
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If rows.Count < 2 Then Exit Sub

    outRow = startRow
    For i = 2 To rows.Count
        Set d = CsvRowDict(rows(1), rows(i))
        For j = LBound(fields) To UBound(fields)
            ws.Cells(outRow, j - LBound(fields) + 1).value = DictGet(d, CStr(fields(j)))
        Next j
        outRow = outRow + 1
    Next i
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.rows.Count, col).End(xlUp).row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Private Sub ConfigurePrintAreas()
    Dim s As Variant, ws As Worksheet, lastRow As Long, lastCol As Long
    Dim lastRowCell As Range, lastColCell As Range

    For Each s In Array("Report", "Property", "Assessment", "Tax History", "Sales-Deeds", "Documents", "Appeals", "Incentives", "Permits", "Issues", "Sources")
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        Set lastRowCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        Set lastColCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

        If Not lastRowCell Is Nothing And Not lastColCell Is Nothing Then
            lastRow = lastRowCell.row
            lastCol = lastColCell.Column

            With ws.PageSetup
                .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
                .Orientation = xlLandscape
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .CenterHeader = "&B" & EscapeExcelHeaderText(GetPropertyValue("Property Address"))
                .RightHeader = EscapeExcelHeaderText(FormatPin(GetPropertyValue("PIN")))
                .CenterFooter = "Public-record information should be independently verified before reliance."
                .RightFooter = "Page &P of &N"
            End With
        Else
            ws.PageSetup.PrintArea = ""
        End If

        Set lastRowCell = Nothing
        Set lastColCell = Nothing
    Next s
End Sub

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

    If Len(field) > 0 Or row.Count > 0 Then
        row.Add field
        result.Add CollectionToArray(row)
    End If

    Set ParseCsv = result
End Function

Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim a() As Variant, i As Long
    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count
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


Private Function PortalPageLooksValid(ByVal txt As String, ByVal pin As String) As Boolean
    PortalPageLooksValid = _
        (InStr(1, txt, FormatPin(pin), vbTextCompare) > 0) And _
        (InStr(1, txt, "PROPERTY CHARACTERISTICS", vbTextCompare) > 0) And _
        (InStr(1, txt, "TAX BILLED AMOUNTS", vbTextCompare) > 0)
End Function

Private Sub CrossCheckPortalSize(ByVal fieldName As String, ByVal portalValue As String)
    Dim existingValue As String
    If Len(Trim$(portalValue)) = 0 Then Exit Sub

    existingValue = GetPropertyValue(fieldName)
    If Len(existingValue) = 0 Then
        SetProperty fieldName, portalValue, "Property Tax Portal", "Live", "CHECK", "Portal value used because CookViewer value was blank."
    ElseIf NormalizeNumberText(existingValue) <> NormalizeNumberText(portalValue) Then
        AppendIssue "MEDIUM", fieldName & " differs between CookViewer and the Property Tax Portal.", _
                    "CookViewer / Property Tax Portal", _
                    "CookViewer: " & existingValue & "; Portal: " & portalValue & ". Verify before relying on the measurement."
    End If
End Sub

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

' ArcGIS REST query responses place a "fieldAliases" map (field name -> alias, which
' defaults to the field name itself when no custom alias is configured) before the
' actual "attributes" of each returned feature. A "find this key anywhere in the JSON"
' scan can match that alias entry instead of the real value, which silently returns the
' field's own name as if it were the data (e.g. BCLASS -> "BCLASS" instead of "2-02").
' Scoping every lookup to the first feature's "attributes" object avoids that entirely.
' Falls back to the whole document if no "attributes" object is found, so non-feature
' JSON (or a malformed/truncated response) still gets a best-effort lookup.
Private Function ArcGisFirstAttributesJson(ByVal js As String) As String
    Dim p As Long, openPos As Long, i As Long, depth As Long, ch As String
    Dim inStr As Boolean, esc As Boolean

    p = InStr(1, js, """attributes""", vbTextCompare)
    If p = 0 Then
        ArcGisFirstAttributesJson = js
        Exit Function
    End If
    openPos = InStr(p, js, "{", vbBinaryCompare)
    If openPos = 0 Then
        ArcGisFirstAttributesJson = js
        Exit Function
    End If

    depth = 0
    inStr = False
    esc = False
    For i = openPos To Len(js)
        ch = Mid$(js, i, 1)
        If inStr Then
            If esc Then
                esc = False
            ElseIf ch = "\" Then
                esc = True
            ElseIf ch = """" Then
                inStr = False
            End If
        Else
            Select Case ch
                Case """": inStr = True
                Case "{": depth = depth + 1
                Case "}"
                    depth = depth - 1
                    If depth = 0 Then
                        ArcGisFirstAttributesJson = Mid$(js, openPos, i - openPos + 1)
                        Exit Function
                    End If
            End Select
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


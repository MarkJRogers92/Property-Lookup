Attribute VB_Name = "CookPropertyDueDiligence_Standalone"
Option Explicit

' ============================================================================
' COOK COUNTY PROPERTY DUE DILIGENCE - STANDALONE WINDOWS EXCEL MVP
'
' Design:
'   * Enter a Cook County PIN in Start!B4
'   * GenerateCookPropertyReport calls each public source independently
'   * A failed/slow source is logged and the run continues
'   * Results populate workbook sheets and export to one PDF
'
' Production target: Windows Excel.
' Uses late-bound WinHTTP / VBScript.RegExp / Scripting.Dictionary.
' No external references are required.
'
' IMPORTANT:
'   Enterprise Zone status is reported from the Cook County Assessor's own
'   Parcel Universe spatial field (econ_enterprise_zone_num), not by guessing
'   at the DCEO ArcGIS web app's underlying FeatureServer. That county field
'   is a useful automatic signal but is not the live state boundary, so the
'   official DCEO map link is always preserved for transaction-sensitive use.
' ============================================================================

Private Const SOCRATA_BASE As String = "https://datacatalog.cookcountyil.gov/resource/"
Private Const ADDRESS_DATASET As String = "3723-97qp"
Private Const UNIVERSE_DATASET As String = "pabr-t5kh"
Private Const ASSESSMENT_DATASET As String = "uzyt-m557"
Private Const SALES_DATASET As String = "wvhk-k5uv"
Private Const APPEALS_DATASET As String = "y282-6ig3"
Private Const BOR_DATASET As String = "7pny-nedm"
Private Const PERMITS_DATASET As String = "6yjf-dfxs"

' Cross-adapter state populated by FetchUniverse and consumed later by
' FetchTIF / FetchEnterpriseZone within the same run. Reset in ResetRunSheets.
Private mUniverseOk As Boolean
Private mEnterpriseZoneNum As String
Private mEnterpriseZoneYear As String
Private mTaxTifDistrictName As String

Public Sub GenerateCookPropertyReport()
    Dim pin As String
    Dim t0 As Double
    t0 = Timer

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    pin = NormalizePin(CStr(ThisWorkbook.Worksheets("Start").Range("B4").Value2))
    If Len(pin) <> 14 Then
        MsgBox "Enter a valid 14-digit Cook County PIN.", vbExclamation
        GoTo SafeExit
    End If

    ThisWorkbook.Worksheets("Start").Range("B4").Value = FormatPin(pin)
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
    UpdateSourceStatus sourceName, "OK", "Finished", ""
    Exit Sub

EH:
    UpdateSourceStatus sourceName, "FAILED", "Stopped", Err.Description
    AppendIssue "HIGH", sourceName & " could not be verified automatically.", sourceName, _
                "Open the official source link on the Start/Sources sheet and verify manually. Error: " & Err.Description
    Err.Clear
End Sub

Private Sub FetchAddress(ByVal pin As String)
    Dim url As String, csv As String, rows As Collection, d As Object
    url = SocrataCsv(ADDRESS_DATASET, pin, "$order=year DESC&$limit=1")
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)
    If rows.Count < 2 Then Err.Raise vbObjectError + 101, , "No address record returned."
    Set d = CsvRowDict(rows(1), rows(2))

    SetProperty "PIN", pin, "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Property Address", DictGet(d, "prop_address_full"), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Property City / ZIP", Trim$(Nz(DictGet(d, "prop_address_city_name")) & " " & Nz(DictGet(d, "prop_address_zipcode_1"))), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Owner Name", DictGet(d, "owner_address_name"), "Assessor Addresses", Nz(DictGet(d, "year")), "CAUTION", "Owner/mailing data may be intermittently updated."
    SetProperty "Taxpayer / Mailing Name", DictGet(d, "mail_address_name"), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Mailing Address", Trim$(Nz(DictGet(d, "mail_address_full")) & ", " & Nz(DictGet(d, "mail_address_city_name")) & ", " & Nz(DictGet(d, "mail_address_state")) & " " & Nz(DictGet(d, "mail_address_zipcode_1"))), "Assessor Addresses", Nz(DictGet(d, "year")), "OK", ""
End Sub

Private Sub FetchUniverse(ByVal pin As String)
    Dim url As String, csv As String, rows As Collection, d As Object
    url = SocrataCsv(UNIVERSE_DATASET, pin, "$limit=1")
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)
    If rows.Count < 2 Then Err.Raise vbObjectError + 102, , "No current Parcel Universe row returned."
    Set d = CsvRowDict(rows(1), rows(2))

    SetProperty "Township", DictGet(d, "township_name"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Assessor Class", DictGet(d, "class"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Assessor Neighborhood", FirstNonBlank(d, Array("nbhd_code", "nbhd")), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Tax Code", DictGet(d, "tax_code"), "Parcel Universe", Nz(DictGet(d, "year")), "CAUTION", "County metadata warns this field may not be current."
    SetProperty "Longitude", DictGet(d, "lon"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Latitude", DictGet(d, "lat"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Centroid X (3435)", DictGet(d, "x_3435"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""
    SetProperty "Centroid Y (3435)", DictGet(d, "y_3435"), "Parcel Universe", Nz(DictGet(d, "year")), "OK", ""

    ' Confirmed against CCAO's own open-data view definition (open_data.vw_parcel_universe_historical,
    ' ccao-data/data-architecture on GitHub): tax_municipality_name is assigned via the parcel's tax
    ' code (tax.agency_info, minor_type='MUNI'); cook_municipality_name is a spatial point-in-polygon
    ' join against the county's municipal boundary layer. The two can legitimately disagree.
    SetProperty "Municipality - Tax Record", DictGet(d, "tax_municipality_name"), "Parcel Universe", Nz(DictGet(d, "year")), "CHECK", "Tax-record municipality (assigned via tax code); spatial municipality may disagree."
    SetProperty "Municipality - Spatial", DictGet(d, "cook_municipality_name"), "Parcel Universe", Nz(DictGet(d, "year")), "CHECK", "Spatial municipality (point-in-polygon); preferred for geographic inclusion. Verify if blank."

    ' Captured here (not shown on the Property sheet) for use by FetchTIF and FetchEnterpriseZone,
    ' which run as later, independent adapters and cannot re-query this row themselves.
    mEnterpriseZoneNum = DictGet(d, "econ_enterprise_zone_num")
    mEnterpriseZoneYear = DictGet(d, "econ_enterprise_zone_data_year")
    mTaxTifDistrictName = DictGet(d, "tax_tif_district_name")
    mUniverseOk = True
End Sub

Private Sub FetchAssessments(ByVal pin As String)
    Dim n As Long, url As String, csv As String
    n = CLng(GetConfigValue("Assessment Years", 8))
    url = SocrataCsv(ASSESSMENT_DATASET, pin, "$order=year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    WriteSelectedCsv "Assessment", csv, 5, _
        Array("year","class","mailed_land","mailed_bldg","mailed_tot","certified_land","certified_bldg","certified_tot","board_land","board_bldg","board_tot")
End Sub

Private Sub FetchSales(ByVal pin As String)
    Dim n As Long, url As String, csv As String
    n = CLng(GetConfigValue("Sales Limit", 20))
    url = SocrataCsv(SALES_DATASET, pin, "$order=sale_date DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    WriteMappedTable "Sales-Deeds", csv, 5, _
        Array("sale_date","sale_price","doc_no","deed_type","mydec_deed_type","buyer_name","seller_name","is_multisale","num_parcels_sale"), _
        Array("Sale Date","Sale Price","Document No.","Deed Type","MyDec Deed Type","Buyer","Seller","Multi-PIN?","# Parcels")
End Sub

Private Sub FetchAssessorAppeals(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection, i As Long, d As Object
    n = CLng(GetConfigValue("Appeals Limit", 30))
    url = SocrataCsv(APPEALS_DATASET, pin, "$order=year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)
    If rows.Count < 2 Then Exit Sub

    For i = 2 To rows.Count
        Set d = CsvRowDict(rows(1), rows(i))
        AppendAppeal "Assessor", DictGet(d,"year"), DictGet(d,"case_no"), _
                     FirstNonBlank(d, Array("appeal_type","hearing_type")), DictGet(d,"status"), _
                     DictGet(d,"mailed_tot"), DictGet(d,"certified_tot"), "Assessor Appeals"
    Next i
End Sub

Private Sub FetchBORAppeals(ByVal pin As String)
    Dim n As Long, url As String, csv As String, rows As Collection, i As Long, d As Object
    n = CLng(GetConfigValue("BOR Limit", 30))
    url = SocrataCsv(BOR_DATASET, pin, "$order=tax_year DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    Set rows = ParseCsv(csv)
    If rows.Count < 2 Then Exit Sub

    For i = 2 To rows.Count
        Set d = CsvRowDict(rows(1), rows(i))
        AppendAppeal "Board of Review", DictGet(d,"tax_year"), _
                     Nz(DictGet(d,"appealtrk")) & "-" & Nz(DictGet(d,"appealseq")), _
                     FirstNonBlank(d, Array("appealtype","appealtypedescription")), "", _
                     DictGet(d,"assessor_totalvalue"), DictGet(d,"bor_totalvalue"), "BOR Decision History"
    Next i
End Sub

Private Sub FetchPermits(ByVal pin As String)
    Dim n As Long, url As String, csv As String
    n = CLng(GetConfigValue("Permits Limit", 50))
    url = SocrataCsv(PERMITS_DATASET, pin, "$order=date_issued DESC&$limit=" & CStr(n))
    csv = HttpGet(url)
    WriteMappedTable "Permits", csv, 5, _
        Array("date_issued","year","local_permit_number","permit_number","status","assessable","amount","municipality","applicant_name","work_description"), _
        Array("Issued Date","Year","Local Permit No.","CCAO Permit No.","Status","Assessable?","Amount","Municipality","Applicant","Work Description")
End Sub

Private Sub FetchGISParcel(ByVal pin As String)
    Dim yr As String, baseUrl As String, whereClause As String, url As String, js As String
    Dim cls As String, nbhd As String, muni As String, taxCode As String, pinHit As String

    yr = CStr(GetConfigValue("GIS Parcel Year", 2025))
    baseUrl = "https://gis.cookcountyil.gov/traditional/rest/services/parcelHistorical/MapServer/" & yr & "/query?"
    whereClause = "Name='" & pin & "' OR Name='" & FormatPin(pin) & "'"
    url = baseUrl & "where=" & UrlEncode(whereClause) & _
          "&outFields=Name,PIN10,TAXCODE,Latitude,Longitude,AssessorBLDGclass,AssessorNBHD,MUNICIPALITY&returnGeometry=false&f=json"
    js = HttpGet(url)

    If ArcGisFeatureCount(js) = 0 Then
        Err.Raise vbObjectError + 151, , "GIS did not return the PIN in the configured parcel layer (year " & yr & ")."
    End If

    pinHit = JsonScalar(js, "Name")
    cls = JsonScalar(js, "AssessorBLDGclass")
    nbhd = JsonScalar(js, "AssessorNBHD")
    muni = JsonScalar(js, "MUNICIPALITY")
    taxCode = JsonScalar(js, "TAXCODE")

    SetProperty "GIS PIN Match", pinHit, "Cook County GIS", yr, "OK", ""
    SetProperty "GIS Building Class", cls, "Cook County GIS", yr, "OK", ""
    SetProperty "GIS Neighborhood", nbhd, "Cook County GIS", yr, "OK", ""
    If Len(muni) > 0 Then SetProperty "Municipality - Spatial", muni, "Cook County GIS", yr, "OK", "GIS parcel layer value."
    If Len(taxCode) > 0 And Len(GetPropertyValue("Tax Code")) = 0 Then
        SetProperty "Tax Code", taxCode, "Cook County GIS", yr, "OK", ""
    End If
End Sub

Private Sub FetchTaxPortal(ByVal pin As String)
    Dim url As String, html As String, txt As String
    Dim rate As String, taxCode As String, rateBlock As String, billBlock As String
    Dim re As Object, ms As Object, m As Object, r As Long

    url = "https://www.cookcountypropertyinfo.com/Pages/Pin-Results.aspx?pin=" & pin
    html = HttpGet(url)
    txt = CollapseWhitespace(HtmlToText(html))

    If InStr(1, txt, "PROPERTY CHARACTERISTICS", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 160, , "Property Tax Portal page retrieved but expected property content was not found."
    End If

    rate = RegexFirst(txt, "Tax Rate\s*:\s*([0-9\.]+)", 1)
    taxCode = RegexFirst(txt, "Tax Code\s*:\s*([0-9]+)", 1)

    If Len(rate) > 0 Then
        ThisWorkbook.Worksheets("Report").Range("E7").Value = rate
    End If
    If Len(taxCode) > 0 Then
        SetProperty "Tax Code", taxCode, "Property Tax Portal", "Live", "OK", "Portal tax code takes precedence for current tax-rate context."
    End If

    ' Tax-rate history block.
    rateBlock = TextBetween(txt, "Tax Rate History", "Tax Code")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(20[0-9]{2})\s+([0-9\.]+)"
    Set ms = re.Execute(rateBlock)

    r = 5
    For Each m In ms
        ThisWorkbook.Worksheets("Tax History").Cells(r, 1).Value = m.SubMatches(0)
        ThisWorkbook.Worksheets("Tax History").Cells(r, 4).Value = m.SubMatches(1)
        ThisWorkbook.Worksheets("Tax History").Cells(r, 5).Value = taxCode
        ThisWorkbook.Worksheets("Tax History").Cells(r, 6).Value = "Cook County Property Tax Portal"
        r = r + 1
    Next m

    ' Billed amounts. We deliberately parse only obvious year/$ pairs.
    billBlock = TextBetween(txt, "TAX BILLED AMOUNTS", "TAX EXEMPTIONS")
    If Len(billBlock) = 0 Then billBlock = TextBetween(txt, "TAX BILLED AMOUNTS", "DOCUMENTS")
    re.Pattern = "(20[0-9]{2})\s*:\s*\$?([0-9,\.]+)"
    Set ms = re.Execute(billBlock)
    Dim firstBill As Boolean
    firstBill = True
    For Each m In ms
        WriteTaxHistoryValue CStr(m.SubMatches(0)), CStr(m.SubMatches(1)), taxCode
        If firstBill Then
            ThisWorkbook.Worksheets("Report").Range("E8").Value = "$" & CStr(m.SubMatches(1))
            firstBill = False
        End If
    Next m
End Sub

Private Sub FetchTIF(ByVal pin As String)
    Dim x As String, y As String, tifUrl As String, url As String, js As String, tifName As String
    Dim note As String

    x = GetPropertyValue("Centroid X (3435)")
    y = GetPropertyValue("Centroid Y (3435)")
    If Len(x) = 0 Or Len(y) = 0 Then
        Err.Raise vbObjectError + 171, , "Parcel centroid CRS 3435 coordinates are unavailable."
    End If

    tifUrl = ResolveTifLayerUrl()
    url = tifUrl & "/query?geometry=" & UrlEncode(x & "," & y) & _
          "&geometryType=esriGeometryPoint&inSR=3435&spatialRel=esriSpatialRelIntersects" & _
          "&outFields=*&returnGeometry=false&f=json"
    js = HttpGet(url)

    If ArcGisFeatureCount(js) = 0 Then
        If mUniverseOk And Len(mTaxTifDistrictName) > 0 Then
            AppendIncentive "TIF", "No spatial intersection, but tax code indicates TIF", mTaxTifDistrictName, _
                            "Cook County TIF layer (" & tifUrl & ")", _
                            "Spatial layer and the parcel's tax-code TIF designation disagree. Verify before reliance.", tifUrl
            AppendIssue "HIGH", "Spatial TIF layer and the parcel's tax-code TIF designation disagree.", _
                        "Cook County TIF GIS / Parcel Universe (tax_tif_district_name)", _
                        "Confirm current TIF district status; the spatial layer (point-in-polygon) and the tax-code assignment can lag each other."
        Else
            AppendIncentive "TIF", "No intersection found", "", "Cook County TIF layer (" & tifUrl & ")", _
                            "Verify if transaction is sensitive to TIF status.", tifUrl
        End If
    Else
        tifName = JsonScalar(js, "TIF_NAME")
        If Len(tifName) = 0 Then tifName = JsonScalar(js, "AGENCY_DES")
        note = "Verify district is active/current before reliance."
        If mUniverseOk And Len(mTaxTifDistrictName) = 0 Then
            note = note & " Note: parcel's tax-code TIF designation is blank, which disagrees with this spatial result."
        End If
        AppendIncentive "TIF", "Intersection found", tifName, "Cook County TIF layer (" & tifUrl & ")", note, tifUrl
    End If
End Sub

' Attempts to resolve the current/latest TIF polygon layer under the configured TIF MapServer by
' inspecting the service root for a layer named like "Tax Increment ... (YYYY)" and picking the
' highest year. Falls back to the statically configured layer URL on any failure, so a MapServer
' change never breaks the adapter outright (Config!TIF Layer URL remains the source of truth if
' discovery fails).
Private Function ResolveTifLayerUrl() As String
    Dim configured As String, root As String, js As String, slashPos As Long
    Dim re As Object, ms As Object, m As Object
    Dim bestId As Long, bestYear As Long, thisId As Long, thisYear As Long

    configured = CStr(GetConfigValue("TIF Layer URL", "https://gis.cookcountyil.gov/traditional/rest/services/tifSrvc/MapServer/3"))
    ResolveTifLayerUrl = configured

    On Error GoTo Fallback
    slashPos = InStrRev(configured, "/MapServer")
    If slashPos = 0 Then Exit Function
    root = Left$(configured, slashPos + Len("/MapServer") - 1)
    js = HttpGet(root & "?f=json")

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "\{[^{}]*""id""\s*:\s*(\d+)[^{}]*""name""\s*:\s*""([^""]*(?:Tax Increment|TIF)[^""]*)""[^{}]*\}"
    Set ms = re.Execute(js)

    bestId = -1
    bestYear = -1
    For Each m In ms
        thisId = CLng(m.SubMatches(0))
        thisYear = 0
        On Error Resume Next
        thisYear = CLng(RegexFirst(CStr(m.SubMatches(1)), "(20[0-9]{2})", 1))
        On Error GoTo Fallback
        If thisYear = 0 Then thisYear = thisId
        If thisYear > bestYear Then
            bestYear = thisYear
            bestId = thisId
        End If
    Next m

    If bestId >= 0 Then ResolveTifLayerUrl = root & "/" & CStr(bestId)
    Exit Function

Fallback:
    ResolveTifLayerUrl = configured
End Function

Private Sub FetchEnterpriseZone(ByVal pin As String)
    Dim appUrl As String
    appUrl = CStr(GetConfigValue("DCEO Enterprise Zone App", _
        "https://idor.maps.arcgis.com/apps/webappviewer/index.html?id=f82fc6b62fde435abb41f5f72db2db48"))

    ' The DCEO ArcGIS web app's underlying statewide FeatureServer is intentionally NOT
    ' reverse-engineered here (per review guidance: do not substitute a guessed layer).
    ' Instead, the Parcel Universe dataset already carries the Assessor's own spatial join
    ' against an Enterprise Zone boundary layer (econ_enterprise_zone_num / _data_year,
    ' confirmed via CCAO's open_data.vw_parcel_universe_historical view). That value is
    ' captured for free during FetchUniverse and used as an automatic, but non-authoritative,
    ' signal here. It is county-maintained, not the live DCEO/IDOR boundary, so legal or
    ' transaction-sensitive reliance should still be confirmed on the official map.
    If mUniverseOk Then
        If Len(mEnterpriseZoneNum) > 0 Then
            AppendIncentive "Illinois Enterprise Zone", "Zone indicated (Assessor spatial layer)", _
                            "Zone " & mEnterpriseZoneNum, "Assessor spatial layer, data year " & mEnterpriseZoneYear, _
                            "County-maintained spatial join, not the live DCEO boundary. Confirm on the official DCEO interactive map before relying on this for a transaction.", appUrl
            AppendIssue "MEDIUM", "Enterprise Zone was indicated by the Assessor's spatial layer, not the live state map.", _
                        "Parcel Universe (econ_enterprise_zone_num) / Illinois DCEO Enterprise Zone Map", _
                        "Confirm current zone boundary and expiration on the official DCEO interactive map before relying on this for a transaction."
        Else
            AppendIncentive "Illinois Enterprise Zone", "No zone indicated (Assessor spatial layer)", "", _
                            "Assessor spatial layer, data year " & mEnterpriseZoneYear, _
                            "County-maintained spatial join, not the live DCEO boundary. Confirm on the official map if Enterprise Zone status is transaction-relevant.", appUrl
        End If
    Else
        AppendIncentive "Illinois Enterprise Zone", "Manual verification required", "", _
                        "Official DCEO interactive map", "Parcel Universe centroid/zone fields were unavailable this run.", appUrl
        AppendIssue "MEDIUM", "Enterprise Zone status was not automatically resolved.", _
                    "Illinois DCEO Enterprise Zone Map", _
                    "Use the official DCEO interactive map. The Parcel Universe adapter (source of the county's own zone field) did not complete this run."
    End If
End Sub

Private Sub BuildIssues()
    Dim aClass As String, gClass As String, mTax As String, mSpatial As String
    aClass = NormalizeCompare(GetPropertyValue("Assessor Class"))
    gClass = NormalizeCompare(GetPropertyValue("GIS Building Class"))
    mTax = NormalizeCompare(GetPropertyValue("Municipality - Tax Record"))
    mSpatial = NormalizeCompare(GetPropertyValue("Municipality - Spatial"))

    If Len(aClass) > 0 And Len(gClass) > 0 And aClass <> gClass Then
        AppendIssue "HIGH", "Assessor class and GIS building class do not match.", _
                    "Parcel Universe / Cook County GIS", _
                    "Review both current Assessor records and the configured GIS layer; do not silently overwrite either value."
    End If

    If Len(mTax) > 0 And Len(mSpatial) > 0 And mTax <> mSpatial Then
        AppendIssue "HIGH", "Tax-record municipality and spatial municipality differ.", _
                    "Parcel Universe / Cook County GIS", _
                    "Confirm municipal boundary for zoning/incentive work. The Assessor dataset itself warns the two municipality concepts can disagree."
    End If

    If Len(GetPropertyValue("Owner Name")) = 0 Then
        AppendIssue "MEDIUM", "Owner name was not returned.", "Assessor Addresses", _
                    "Check recent sales/recorded documents; Assessor owner/mailing data can lag."
    End If

    If Len(GetPropertyValue("Tax Code")) = 0 Then
        AppendIssue "HIGH", "Current tax code was not verified.", "Property Tax Portal / Parcel Universe / GIS", _
                    "Verify on the Cook County Property Tax Portal."
    End If
End Sub

Private Sub BuildReport()
    Dim ws As Worksheet, issuesText As String, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Report")

    ws.Range("B5").Value = FormatPin(GetPropertyValue("PIN"))
    ws.Range("B6").Value = GetPropertyValue("Property Address")
    ws.Range("B7").Value = FirstText(GetPropertyValue("Municipality - Spatial"), GetPropertyValue("Municipality - Tax Record"))
    ws.Range("B8").Value = GetPropertyValue("Township")
    ws.Range("B9").Value = GetPropertyValue("Assessor Class")
    ws.Range("B10").Value = GetPropertyValue("GIS Building Class")
    ws.Range("B11").Value = GetPropertyValue("Tax Code")

    ws.Range("E5").Value = GetPropertyValue("Owner Name")
    ws.Range("E6").Value = GetPropertyValue("Taxpayer / Mailing Name")
    ws.Range("E11").Value = Format(Date, "mmmm d, yyyy")

    ws.Range("E9").Value = FindIncentiveResult("TIF")
    ws.Range("E10").Value = FindIncentiveResult("Illinois Enterprise Zone")

    lastRow = LastUsedRow(ThisWorkbook.Worksheets("Issues"), 1)
    For r = 5 To lastRow
        If Len(CStr(ThisWorkbook.Worksheets("Issues").Cells(r, 2).Value)) > 0 Then
            issuesText = issuesText & "• " & ThisWorkbook.Worksheets("Issues").Cells(r, 1).Value & ": " & _
                         ThisWorkbook.Worksheets("Issues").Cells(r, 2).Value & vbCrLf
        End If
    Next r
    If Len(issuesText) = 0 Then issuesText = "No material automated warnings were generated. Public records should still be independently verified."
    ws.Range("A14").Value = issuesText
End Sub

Private Sub ExportDueDiligencePDF(ByVal pin As String)
    Dim folder As String, fn As String, fullPath As String, openAfter As Boolean
    Dim arr As Variant

    folder = CStr(GetConfigValue("PDF Output Folder", ""))
    If Len(Trim$(folder)) = 0 Then folder = ThisWorkbook.Path
    If Len(Trim$(folder)) = 0 Then folder = Environ$("USERPROFILE") & "\Documents"
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    fn = "Cook_Property_Due_Diligence_" & pin & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    fullPath = folder & fn

    ConfigurePrintAreas
    arr = Array("Report","Property","Assessment","Tax History","Sales-Deeds","Appeals","Incentives","Permits","Issues")
    ThisWorkbook.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ThisWorkbook.Worksheets("Start").Select

    openAfter = (UCase$(CStr(GetConfigValue("Open PDF After Export", "YES"))) = "YES")
    If openAfter Then ThisWorkbook.FollowHyperlink fullPath
End Sub

' -----------------------
' Workbook writing helpers
' -----------------------
Private Sub ResetRunSheets()
    Dim s As Variant, ws As Worksheet
    For Each s In Array("Assessment","Tax History","Sales-Deeds","Appeals","Incentives","Permits","Issues")
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        ws.Rows("5:" & ws.Rows.Count).ClearContents
    Next s

    Set ws = ThisWorkbook.Worksheets("Property")
    ws.Range("B5:F24").ClearContents
    ThisWorkbook.Worksheets("Report").Range("B5:B11,E5:E11,A14").ClearContents

    Dim r As Long
    For r = 12 To 22
        ThisWorkbook.Worksheets("Start").Cells(r, 2).Value = "Not run"
        ThisWorkbook.Worksheets("Start").Cells(r, 3).ClearContents
        ThisWorkbook.Worksheets("Start").Cells(r, 4).ClearContents
    Next r

    mUniverseOk = False
    mEnterpriseZoneNum = ""
    mEnterpriseZoneYear = ""
    mTaxTifDistrictName = ""
End Sub

Private Sub SetRunStatus(ByVal msg As String)
    ThisWorkbook.Worksheets("Start").Range("B8").Value = msg
    DoEvents
End Sub

Private Sub UpdateSourceStatus(ByVal sourceName As String, ByVal status As String, ByVal progress As String, ByVal notes As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Start")
    For r = 12 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), sourceName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value = status
            ws.Cells(r, 3).Value = progress & " " & Format(Now, "hh:mm:ss")
            ws.Cells(r, 4).Value = notes
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
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), fieldName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value = value
            ws.Cells(r, 3).Value = sourceName
            ws.Cells(r, 4).Value = asOfText
            ws.Cells(r, 5).Value = status
            ws.Cells(r, 6).Value = notes
            Exit Sub
        End If
    Next r
End Sub

Private Function GetPropertyValue(ByVal fieldName As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Property")
    For r = 5 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), fieldName, vbTextCompare) = 0 Then
            GetPropertyValue = Trim$(CStr(ws.Cells(r, 2).Value))
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
    ws.Cells(r, 1).Value = level
    ws.Cells(r, 2).Value = taxYear
    ws.Cells(r, 3).Value = caseNo
    ws.Cells(r, 4).Value = appealType
    ws.Cells(r, 5).Value = status
    ws.Cells(r, 6).Value = beforeAV
    ws.Cells(r, 7).Value = afterAV
    ws.Cells(r, 8).Value = chg
    ws.Cells(r, 9).Value = sourceNote
End Sub

Private Sub AppendIncentive(ByVal programName As String, ByVal result As String, ByVal nm As String, _
                            ByVal dataYear As String, ByVal verification As String, ByVal sourceUrl As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Incentives")
    r = LastUsedRow(ws, 1) + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).Value = programName
    ws.Cells(r, 2).Value = result
    ws.Cells(r, 3).Value = nm
    ws.Cells(r, 4).Value = dataYear
    ws.Cells(r, 5).Value = verification
    ws.Cells(r, 6).Value = sourceUrl
End Sub

Private Function FindIncentiveResult(ByVal programName As String) As String
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Incentives")
    lastRow = LastUsedRow(ws, 1)
    For r = 5 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).Value), programName, vbTextCompare) = 0 Then
            FindIncentiveResult = CStr(ws.Cells(r, 2).Value)
            If Len(CStr(ws.Cells(r, 3).Value)) > 0 Then FindIncentiveResult = FindIncentiveResult & " - " & CStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r
End Function

Private Sub AppendIssue(ByVal priority As String, ByVal issueText As String, ByVal sourceText As String, ByVal verifyText As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Issues")
    r = LastUsedRow(ws, 1) + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).Value = priority
    ws.Cells(r, 2).Value = issueText
    ws.Cells(r, 3).Value = sourceText
    ws.Cells(r, 4).Value = verifyText
End Sub

Private Sub WriteTaxHistoryValue(ByVal taxYear As String, ByVal billed As String, ByVal taxCode As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Tax History")
    lastRow = LastUsedRow(ws, 1)
    For r = 5 To IIf(lastRow < 5, 5, lastRow)
        If CStr(ws.Cells(r, 1).Value) = taxYear Then
            ws.Cells(r, 2).Value = billed
            ws.Cells(r, 5).Value = taxCode
            ws.Cells(r, 6).Value = "Cook County Property Tax Portal"
            Exit Sub
        End If
    Next r
    r = lastRow + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).Value = taxYear
    ws.Cells(r, 2).Value = billed
    ws.Cells(r, 5).Value = taxCode
    ws.Cells(r, 6).Value = "Cook County Property Tax Portal"
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
    Dim rows As Collection, d As Object, ws As Worksheet
    Dim i As Long, j As Long, outRow As Long
    Set rows = ParseCsv(csv)
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If rows.Count < 2 Then Exit Sub

    outRow = startRow
    For i = 2 To rows.Count
        Set d = CsvRowDict(rows(1), rows(i))
        For j = LBound(fields) To UBound(fields)
            ws.Cells(outRow, j - LBound(fields) + 1).Value = DictGet(d, CStr(fields(j)))
        Next j
        outRow = outRow + 1
    Next i
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Private Sub ConfigurePrintAreas()
    Dim s As Variant, ws As Worksheet, lastRow As Long, lastCol As Long
    Dim foundRow As Range, foundCol As Range
    For Each s In Array("Report","Property","Assessment","Tax History","Sales-Deeds","Appeals","Incentives","Permits","Issues")
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        Set foundRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If foundRow Is Nothing Then
            lastRow = 1
            lastCol = 1
        Else
            lastRow = foundRow.Row
            Set foundCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            lastCol = foundCol.Column
        End If
        With ws.PageSetup
            .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHeader = "&B" & EscapeHeaderFooterAmp(GetPropertyValue("Property Address"))
            .RightHeader = FormatPin(GetPropertyValue("PIN"))
            .CenterFooter = "Public-record information should be independently verified before reliance."
            .RightFooter = "Page &P of &N"
        End With
    Next s
End Sub

' Excel header/footer codes use "&" to introduce formatting tokens (e.g. "&B" for bold), so a
' literal "&" in the printed text (a corner-lot address like "MAIN ST & 5TH AVE") must be doubled
' or it is silently swallowed / misinterpreted.
Private Function EscapeHeaderFooterAmp(ByVal s As String) As String
    EscapeHeaderFooterAmp = Replace(s, "&", "&&")
End Function

' -----------------------
' HTTP / source helpers
' -----------------------
Private Function SocrataCsv(ByVal datasetId As String, ByVal pin As String, ByVal extraQuery As String) As String
    Dim q As String
    q = "$where=" & UrlEncode("pin='" & pin & "'")
    If Len(extraQuery) > 0 Then q = q & "&" & extraQuery
    ' extraQuery contains SoQL like "$order=year DESC" with a literal space, which is not a
    ' valid character inside an HTTP request line/URL. Encode it after assembly so callers can
    ' keep writing plain SoQL.
    SocrataCsv = SOCRATA_BASE & datasetId & ".csv?" & Replace(q, " ", "%20")
End Function

' Counts ArcGIS REST "attributes" objects in a query response. Robust to both f=json (compact)
' and f=pjson (pretty-printed, which Esri servers space as `"features" : []`) since it counts
' the per-feature "attributes" key rather than pattern-matching an empty features array literal.
' Raises if the payload looks like an ArcGIS error response instead of a feature set, so a
' service error is never silently reported as "no intersection" / "not found".
Private Function ArcGisFeatureCount(ByVal js As String) As Long
    If InStr(1, js, """features""", vbTextCompare) = 0 Then
        If InStr(1, js, """error""", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 190, , "ArcGIS REST service returned an error: " & js
        End If
        ArcGisFeatureCount = 0
        Exit Function
    End If

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = """attributes""\s*:"
    ArcGisFeatureCount = re.Execute(js).Count
End Function

Private Function HttpGet(ByVal url As String) As String
    Dim http As Object, timeoutMs As Long
    timeoutMs = CLng(GetConfigValue("HTTP Timeout (ms)", 30000))

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 Excel-CookPropertyDueDiligence"
    http.SetRequestHeader "Accept", "*/*"
    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 301, , "HTTP " & http.Status & " from " & url
    End If
    HttpGet = CStr(http.ResponseText)
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

Private Function CsvRowDict(ByVal headerRow As Variant, ByVal dataRow As Variant) As Object
    Dim d As Object, i As Long, maxI As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    maxI = UBound(headerRow)
    If UBound(dataRow) < maxI Then maxI = UBound(dataRow)
    For i = 0 To maxI
        d(Trim$(CStr(headerRow(i)))) = dataRow(i)
    Next i
    Set CsvRowDict = d
End Function

Private Function DictGet(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then DictGet = Trim$(CStr(d(key))) Else DictGet = ""
End Function

Private Function FirstNonBlank(ByVal d As Object, ByVal keys As Variant) As String
    Dim k As Variant, v As String
    For Each k In keys
        v = DictGet(d, CStr(k))
        If Len(v) > 0 Then
            FirstNonBlank = v
            Exit Function
        End If
    Next k
End Function

' -----------------------
' HTML / JSON lightweight parsing
' -----------------------
Private Function HtmlToText(ByVal html As String) As String
    Dim re As Object, s As String
    s = html
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    re.Pattern = "<script[\s\S]*?</script>"
    s = re.Replace(s, " ")
    re.Pattern = "<style[\s\S]*?</style>"
    s = re.Replace(s, " ")
    re.Pattern = "<[^>]+>"
    s = re.Replace(s, " ")

    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&quot;", """")
    HtmlToText = s
End Function

Private Function CollapseWhitespace(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\s+"
    CollapseWhitespace = Trim$(re.Replace(s, " "))
End Function

Private Function RegexFirst(ByVal s As String, ByVal pattern As String, ByVal subIndex As Long) As String
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = pattern
    Set ms = re.Execute(s)
    If ms.Count > 0 Then
        If subIndex <= ms(0).SubMatches.Count Then
            RegexFirst = CStr(ms(0).SubMatches(subIndex - 1))
        Else
            RegexFirst = CStr(ms(0).Value)
        End If
    End If
End Function

Private Function JsonScalar(ByVal js As String, ByVal key As String) As String
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    re.Pattern = """" & key & """\s*:\s*""([^""]*)"""
    Set ms = re.Execute(js)
    If ms.Count > 0 Then
        JsonScalar = CStr(ms(0).SubMatches(0))
        Exit Function
    End If

    re.Pattern = """" & key & """\s*:\s*([\-0-9\.]+)"
    Set ms = re.Execute(js)
    If ms.Count > 0 Then JsonScalar = CStr(ms(0).SubMatches(0))
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
        If ch Like "#" Then out = out & ch
    Next i
    NormalizePin = out
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
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), settingName, vbTextCompare) = 0 Then
            If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then
                GetConfigValue = ws.Cells(r, 2).Value
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
    ' Safe public PIN used only as a smoke-test example.
    ' Replace with known residential, commercial, TIF, and EZ PINs for regression.
    Dim pin As String
    pin = "16302040200000"

    Debug.Print "Testing PIN " & FormatPin(pin)
    Debug.Print "Address URL: " & SocrataCsv(ADDRESS_DATASET, pin, "$order=year DESC&$limit=1")
    Debug.Print "Universe URL: " & SocrataCsv(UNIVERSE_DATASET, pin, "$limit=1")
    Debug.Print "Assessment URL: " & SocrataCsv(ASSESSMENT_DATASET, pin, "$order=year DESC&$limit=3")
    Debug.Print "Sales URL: " & SocrataCsv(SALES_DATASET, pin, "$order=sale_date DESC&$limit=3")

    MsgBox "Adapter URL smoke test completed. Run GenerateCookPropertyReport for end-to-end testing.", vbInformation
End Sub


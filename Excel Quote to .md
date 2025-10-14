Option Explicit

'========================
'====== SETTINGS ========
'========================
Private Const VAULT_ROOT       As String = "C:\Users\CraigMoravec\Obsidian Sync Cloud"
Private Const QUOTES_SUB       As String = "001-Work\Quotes"
Private Const PROPS_SUB        As String = "001-Work\Properties"
Private Const CONTACTS_SUB     As String = "001-Work\Contacts"

Private Const TABLE_NAME       As String = "ShareTable"   ' found even if sheet is hidden

' ShareTable column headers
Private Const COL_QUOTE_NUM    As String = "Quote Number"
Private Const COL_DATE         As String = "Quote Date"   ' uses your actual header
Private Const COL_PROPERTY     As String = "Job Site"
Private Const COL_ACCOUNT      As String = "Account Number"
Private Const COL_CONTACT      As String = "Contact Name"
Private Const COL_PHONE        As String = "Contact Phone" ' NEW
Private Const COL_EMAIL        As String = "Contact Email"        ' NEW
Private Const COL_ADDRESS1     As String = "Address"
Private Const COL_CITYSTATEZIP As String = "City, State, Zip"
Private Const COL_BUILDINGID   As String = "Building ID"
Private Const COL_INSPECTION   As String = "Inspection ID"
Private Const COL_TOTAL        As String = "Total"

' Nominatim for geocoding
Private Const NOMINATIM_BASE   As String = "https://nominatim.openstreetmap.org/search"
Private Const NOMINATIM_UA     As String = "Obsidian-Geocoder (craig.pyebarker@gmail.com)"

'===============================
'====== PUBLIC ENTRYPOINT ======
'===============================
Public Sub Export_All_From_ShareTable()
    Dim lo As ListObject, i As Long

    Set lo = GetTableAnywhere(TABLE_NAME)
    If lo Is Nothing Then
        MsgBox "Table '" & TABLE_NAME & "' not found in the ACTIVE workbook.", vbCritical
        Exit Sub
    End If

    If lo.DataBodyRange Is Nothing Or lo.DataBodyRange.Rows.Count = 0 Then
        MsgBox "Table '" & TABLE_NAME & "' has no data.", vbExclamation
        Exit Sub
    End If

    EnsureFolder CombinePath3(VAULT_ROOT, QUOTES_SUB, "")
    EnsureFolder CombinePath3(VAULT_ROOT, PROPS_SUB, "")
    EnsureFolder CombinePath3(VAULT_ROOT, CONTACTS_SUB, "")

    Application.ScreenUpdating = False
    For i = 1 To lo.DataBodyRange.Rows.Count
        ExportOne lo, i
    Next i
    Application.ScreenUpdating = True

    MsgBox "Export complete for all rows.", vbInformation
End Sub

'==========================
'====== CORE LOGIC  =======
'==========================
Private Sub ExportOne(lo As ListObject, idx As Long)
    Dim rec As Object: Set rec = RowToDict(lo, idx)

    Dim qNum$, qDate$, prop$, acct$, contact$, phone$, email$, addr1$, citystzip$, bldg$, insp$, total$
    qNum = Nz(rec(COL_QUOTE_NUM))
    qDate = DateForYaml(rec(COL_DATE))               ' yyyy-mm-dd or ""
    prop = Nz(rec(COL_PROPERTY))
    acct = Nz(rec(COL_ACCOUNT))
    contact = Nz(rec(COL_CONTACT))
    ' NEW: pick up phone/email from the table
    On Error Resume Next
    phone = Nz(rec(COL_PHONE))
    email = Nz(rec(COL_EMAIL))
    On Error GoTo 0
    addr1 = Nz(rec(COL_ADDRESS1))
    citystzip = Nz(rec(COL_CITYSTATEZIP))
    bldg = Nz(rec(COL_BUILDINGID))
    insp = Nz(rec(COL_INSPECTION))
    total = Nz(rec(COL_TOTAL))

    Dim propPath$, contactPath$, quotePath$
    propPath = CombinePath3(VAULT_ROOT, PROPS_SUB, SanitizeFileName(IIf(Len(prop) > 0, prop, "Property")) & ".md")
    contactPath = CombinePath3(VAULT_ROOT, CONTACTS_SUB, SanitizeFileName(IIf(Len(contact) > 0, contact, "Contact")) & ".md")
    quotePath = CombinePath3(VAULT_ROOT, QUOTES_SUB, SanitizeFileName(IIf(Len(prop) > 0, prop, "Property") & " - " & IIf(Len(qNum) > 0, qNum, "Quote")) & ".md")

    UpsertPropertyNote propPath, prop, addr1, citystzip, acct, bldg, contact
    UpsertContactNote contactPath, contact, acct, prop, phone, email       ' << pass phone/email
    UpsertQuoteNote quotePath, qNum, qDate, prop, acct, contact, addr1, citystzip, bldg, insp, total
End Sub

'==============================================
'======  NOTE BUILDERS / UPSERT HELPERS   =====
'==============================================
Private Sub UpsertPropertyNote(ByVal path$, ByVal prop$, ByVal addr1$, ByVal citystzip$, ByVal acct$, ByVal bldg$, ByVal contact$)
    Dim d As Object: Set d = LoadYamlAsDict(path)

    d("Class") = "Property"

    ' address fields
    If Not d.Exists("Address") Then d("Address") = addr1 Else If Len(d("Address")) = 0 Then d("Address") = addr1
    If Not d.Exists("City, State, Zip") Then d("City, State, Zip") = citystzip _
        Else If Len(d("City, State, Zip")) = 0 Then d("City, State, Zip") = citystzip

    If Not d.Exists("Building ID") Then d("Building ID") = bldg
    If Not d.Exists("Account Number") Then d("Account Number") = acct
    If Not d.Exists("tags") Then d("tags") = "Property"

    ' attach Contact (as a quoted wikilink)
    If Len(contact$) > 0 Then
        If Not d.Exists("Contact") Then
            d("Contact") = AsWikiNormalized(contact$)
        ElseIf Len(Trim$(CStr(d("Contact")))) = 0 Then
            d("Contact") = AsWikiNormalized(contact$)
        End If
    End If

    ' fileclass fields
    If Not d.Exists("icon") Then d("icon") = "building"
    If Not d.Exists("color") Then d("color") = "Blue"
    If Not d.Exists("location") Then d("location") = ""

    WriteTextSafe path, BuildPropertyYaml(d) & vbCrLf

    ' Geocode using Address + City, State, Zip
    If Len(Trim$(d("location"))) = 0 And (Len(Trim$(addr1 & " " & citystzip)) > 0) Then
        Dim lat$, lon$
        If NominatimGeocode(addr1, citystzip, lat, lon) Then
            UpdateYamlKey_YamlOnly path, "location", lat & "," & lon
        End If
    End If
End Sub

' REMOVE Property & location from Contact note; populate phone/email
Private Sub UpsertContactNote(ByVal path$, ByVal contact$, ByVal acct$, ByVal prop$, ByVal phone$, ByVal email$)
    Dim d As Object: Set d = LoadYamlAsDict(path)

    d("Class") = "Contact"
    d("Account Number") = acct$

    ' Ensure Property and location keys are NOT present in Contact notes
    If d.Exists("Property") Then d.Remove "Property"
    If d.Exists("location") Then d.Remove "location"

    ' Phone/Email: write values if present; otherwise ensure keys exist as empty
    If Len(Trim$(phone$)) > 0 Then
        d("Phone Number") = Trim$(phone$)
    ElseIf Not d.Exists("Phone Number") Then
        d("Phone Number") = ""
    End If

    If Len(Trim$(email$)) > 0 Then
        d("Email") = Trim$(email$)
    ElseIf Not d.Exists("Email") Then
        d("Email") = ""
    End If

    If Not d.Exists("tags") Then d("tags") = "Contact"

    WriteTextSafe path, BuildContactYaml(d) & vbCrLf
End Sub

Private Sub UpsertQuoteNote(ByVal path$, ByVal qNum$, ByVal qDate$, ByVal prop$, ByVal acct$, ByVal contact$, ByVal addr1$, ByVal citystzip$, ByVal bldg$, ByVal insp$, ByVal total$)
    Dim d As Object: Set d = LoadYamlAsDict(path)

    d("Class") = "Quote"
    d("Quote Number") = qNum$
    d("Date") = qDate$
    d("Account Number") = acct$
    d("Address") = addr1$
    d("City, State, Zip") = citystzip$
    d("Building ID") = bldg$
    d("Inspection ID") = insp$
    d("Total") = total$

    If Len(prop$) > 0 Then d("Property") = AsWikiNormalized(prop$)
    If Len(contact$) > 0 Then d("Contact") = AsWikiNormalized(contact$)
    If Not d.Exists("tags") Then d("tags") = "Quote"

    WriteTextSafe path, BuildQuoteYaml(d) & vbCrLf
End Sub

'=============================
'====== YAML BUILDERS ========
'=============================
Private Function BuildPropertyYaml(d As Object) As String
    Dim order
    order = Array("Class", _
                  "Address", "City, State, Zip", "location", "icon", "color", _
                  "Building ID", "Account Number", "Contact", "tags")
    BuildPropertyYaml = DictToYaml(d, order)
End Function

Private Function BuildContactYaml(d As Object) As String
    Dim order
    ' Removed "Property" from Contact note
    order = Array("Class", "Phone Number", "Email", "Account Number", "tags")
    BuildContactYaml = DictToYaml(d, order)
End Function

Private Function BuildQuoteYaml(d As Object) As String
    Dim order
    order = Array("Class", "Quote Number", "Date", "Property", "Account Number", _
                  "Contact", "Address", "City, State, Zip", "Building ID", "Inspection ID", "Total", "tags")
    BuildQuoteYaml = DictToYaml(d, order)
End Function

' Convert dictionary to ordered YAML (handles list fields)
Private Function DictToYaml(d As Object, orderArr) As String
    Dim sb As String, k As Variant, i As Long, v As String, line As String
    sb = "---" & vbCrLf

    ' ordered keys first
    For i = LBound(orderArr) To UBound(orderArr)
        k = orderArr(i)
        If d.Exists(k) Then
            v = CStr(d(k))
            line = WriteYamlLine(CStr(k), v)
            If Len(line) > 0 Then sb = sb & line
        End If
    Next i

    ' extras not in the ordered list
    For Each k In d.Keys
        If Not InArray(k, orderArr) Then
            ' Skip Property/location if they sneak in
            If LCase$(CStr(k)) <> "property" And LCase$(CStr(k)) <> "location" Then
                v = CStr(d(k))
                line = WriteYamlLine(CStr(k), v)
                If Len(line) > 0 Then sb = sb & line
            End If
        End If
    Next

    sb = sb & "---"
    DictToYaml = sb
End Function

' ---- YAML emit helpers ----
Private Function WriteYamlLine(ByVal key$, ByVal val$) As String
    Dim raw$
    raw = Dequote(Trim$(val$))

    Select Case LCase$(key)
        Case "date"
            Dim d$: d = ISODateOnly(ToISODateSmart(raw))
            If Len(d) = 0 Then
                WriteYamlLine = ""
            Else
                WriteYamlLine = "Date: " & d & vbCrLf
            End If

        Case "property", "contact"
            If Len(raw) = 0 Then
                WriteYamlLine = ""
            Else
                WriteYamlLine = key & ": " & """" & AsWikiNormalized(raw) & """" & vbCrLf
            End If

        Case "tags"
            If Len(raw) = 0 Then
                WriteYamlLine = ""
            Else
                WriteYamlLine = key & ":" & vbCrLf & "  - " & QuoteIfNeeded(raw) & vbCrLf
            End If

        Case Else
            WriteYamlLine = key & ": " & QuoteIfNeeded(raw) & vbCrLf
    End Select
End Function

' Remove one pair of matching quotes if present
Private Function Dequote(ByVal s$) As String
    If Len(s) >= 2 Then
        If (Left$(s, 1) = """" And Right$(s, 1) = """") _
        Or (Left$(s, 1) = "'" And Right$(s, 1) = "'") Then
            Dequote = Mid$(s, 2, Len(s) - 2)
            Exit Function
        End If
    End If
    Dequote = s
End Function

' Normalize to "[[Name]]" with no inner quotes
Private Function AsWikiNormalized(ByVal s As String) As String
    Dim t$: t = Trim$(s)
    t = Dequote(t)
    If Len(t) >= 4 And Left$(t, 2) = "[[" And Right$(t, 2) = "]]" Then
        t = Mid$(t, 3, Len(t) - 4)
        t = Dequote(Trim$(t))
        AsWikiNormalized = "[[" & t & "]]"
    Else
        t = Dequote(t)
        AsWikiNormalized = IIf(Len(t) = 0, "", "[[" & t & "]]")
    End If
End Function

' Always produce a YAML-safe scalar. Empty -> ""
Private Function QuoteIfNeeded(ByVal s$) As String
    Dim t$: t = Trim$(s)
    If t = "" Then
        QuoteIfNeeded = """" & """"
        Exit Function
    End If
    If Left$(t, 1) = " " Or Right$(t, 1) = " " _
       Or InStr(1, t, ":", vbBinaryCompare) > 0 _
       Or t Like "*[#{}\[\],&*?]|*" Then
        QuoteIfNeeded = """" & Replace$(t, """", """""") & """"
    Else
        QuoteIfNeeded = t
    End If
End Function

'=============================
'====== YAML LOAD/PARSE ======
'=============================
Private Function LoadYamlAsDict(ByVal path$) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    If FileExists(path) Then
        Dim fm$: fm = ExtractFrontMatter(ReadAllText(path))
        If Len(fm) > 0 Then Set d = ParseYamlToDict(fm): d.CompareMode = 1
    End If
    Set LoadYamlAsDict = d
End Function

Private Sub UpdateYamlKey_YamlOnly(ByVal fullPath$, ByVal key$, ByVal value$)
    Dim fm$, d As Object
    fm = ExtractFrontMatter(ReadAllText(fullPath))
    If Len(fm) > 0 Then
        Set d = ParseYamlToDict(fm): d.CompareMode = 1
    Else
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    End If
    d(key) = value$

    Dim newYaml$
    If d.Exists("Class") Then
        Select Case CStr(d("Class"))
            Case "Property": newYaml = BuildPropertyYaml(d)
            Case "Quote":    newYaml = BuildQuoteYaml(d)
            Case "Contact":  newYaml = BuildContactYaml(d)
            Case Else:       newYaml = DictToYaml(d, Array())
        End Select
    Else
        newYaml = DictToYaml(d, Array())
    End If

    WriteTextSafe fullPath, newYaml & vbCrLf
End Sub

Private Function ExtractFrontMatter(ByVal fullText$) As String
    Dim p1&, p2&
    p1 = InStr(1, fullText, "---" & vbCrLf, vbBinaryCompare)
    If p1 = 1 Then
        p2 = InStr(4, fullText, vbCrLf & "---", vbBinaryCompare)
        If p2 > 0 Then ExtractFrontMatter = Mid$(fullText, 1, p2 + 3): Exit Function
    End If
    ExtractFrontMatter = ""
End Function

Private Function ParseYamlToDict(ByVal fm$) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Dim lines() As String, i As Long, line$, k$, v$
    If Len(fm) = 0 Then Set ParseYamlToDict = d: Exit Function
    fm = Replace(fm, vbCr, "")
    If Left$(fm, 4) = "---" & vbLf Then fm = Mid$(fm, 5)
    If Right$(fm, 4) = vbLf & "---" Then fm = Left$(fm, Len(fm) - 4)
    lines = Split(fm, vbLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Or Left$(line, 1) = "#" Then GoTo ContinueLoop
        If InStr(line, ":") > 0 Then
            k = Trim$(Left$(line, InStr(line, ":") - 1))
            v = Trim$(Mid$(line, InStr(line, ":") + 1))
            v = StripYamlQuotes(v)
            d(k) = v
        End If
ContinueLoop:
    Next i
    Set ParseYamlToDict = d
End Function

Private Function StripYamlQuotes(ByVal s$) As String
    If Len(s) >= 2 Then
        If (Left$(s, 1) = """" And Right$(s, 1) = """") Or (Left$(s, 1) = "'" And Right$(s, 1) = "'") Then
            StripYamlQuotes = Mid$(s, 2, Len(s) - 2): Exit Function
        End If
    End If
    StripYamlQuotes = s
End Function

'=============================
'====== GEOCODING ============
'=============================
Private Function NominatimGeocode(ByVal addr1$, ByVal citystzip$, ByRef lat$, ByRef lon$) As Boolean
    Dim ok As Boolean
    ok = TryCensusGeocode(addr1$, citystzip$, lat$, lon$)
    If ok Then
        NominatimGeocode = True
        Exit Function
    End If
    On Error GoTo Fail

    Dim city$, state$, zip$
    ParseCityStateZip citystzip$, city, state, zip

    Dim base$, url$, json$
    base = NOMINATIM_BASE & "?format=json&limit=1&addressdetails=0&countrycodes=us"

    Dim xhr As Object: Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")

    url = base & "&street=" & UrlEncode(addr1$) _
               & "&city=" & UrlEncode(city$) _
               & "&state=" & UrlEncode(state$) _
               & IIf(Len(zip$) > 0, "&postalcode=" & UrlEncode(zip$), "") _
               & "&country=USA"

    xhr.Open "GET", url, False
    xhr.SetRequestHeader "Accept-Language", "en"
    xhr.SetRequestHeader "User-Agent", NOMINATIM_UA
    xhr.Send
    json = xhr.ResponseText

    lat = RegexGet(json, """lat""\s*:\s*""([^""]+)""")
    lon = RegexGet(json, """lon""\s*:\s*""([^""]+)""")
    If Len(lat) > 0 And Len(lon) > 0 Then NominatimGeocode = True: Exit Function

    url = NOMINATIM_BASE & "?format=json&limit=1&countrycodes=us&q=" & UrlEncode(Trim$(addr1$ & ", " & citystzip$ & ", USA"))
    xhr.Open "GET", url, False
    xhr.SetRequestHeader "Accept-Language", "en"
    xhr.SetRequestHeader "User-Agent", NOMINATIM_UA
    xhr.Send
    json = xhr.ResponseText

    lat = RegexGet(json, """lat""\s*:\s*""([^""]+)""")
    lon = RegexGet(json, """lon""\s*:\s*""([^""]+)""")
    NominatimGeocode = (Len(lat) > 0 And Len(lon) > 0)
    Exit Function
Fail:
    NominatimGeocode = False
End Function

Private Function TryCensusGeocode(ByVal addr1$, ByVal citystzip$, ByRef lat$, ByRef lon$) As Boolean
    On Error GoTo Fail
    Dim city$, state$, zip$
    ParseCityStateZip citystzip$, city, state, zip

    If Len(addr1$) = 0 Or Len(city$) = 0 Or Len(state$) = 0 Then
        TryCensusGeocode = False
        Exit Function
    End If

    Dim base$, url$, json$
    base = "https://geocoding.geo.census.gov/geocoder/locations/address"
    url = base & "?format=json&benchmark=Public_AR_Current" _
               & "&street=" & UrlEncode(addr1$) _
               & "&city=" & UrlEncode(city$) _
               & "&state=" & UrlEncode(state$) _
               & IIf(Len(zip$) > 0, "&zip=" & UrlEncode(zip$), "")

    Dim xhr As Object: Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    xhr.Open "GET", url, False
    xhr.SetRequestHeader "Accept-Language", "en"
    xhr.SetRequestHeader "User-Agent", "Excel-VBA (US-Census-Geocoder)"
    xhr.Send
    json = xhr.ResponseText

    Dim y$, x$
    y = RegexGet(json, """y""\s*:\s*([-\d.]+)")
    x = RegexGet(json, """x""\s*:\s*([-\d.]+)")

    If Len(y) > 0 And Len(x) > 0 Then
        lat$ = y
        lon$ = x
        TryCensusGeocode = True
    Else
        TryCensusGeocode = False
    End If
    Exit Function
Fail:
    TryCensusGeocode = False
End Function

Private Sub ParseCityStateZip(ByVal csz$, ByRef city$, ByRef state$, ByRef zip$)
    Dim parts() As String
    Dim s As String: s = Trim$(csz$)
    city = "": state = "": zip = ""

    If Len(s) = 0 Then Exit Sub
    parts = Split(s, ",")

    If UBound(parts) >= 0 Then city = Trim$(parts(0))
    If UBound(parts) >= 1 Then
        Dim rest As String: rest = Trim$(parts(1))
        Dim tokens() As String: tokens = Split(rest, " ")
        If UBound(tokens) >= 0 Then state = UCase$(Trim$(tokens(0)))
        If UBound(tokens) >= 1 Then zip = Trim$(tokens(1))
    End If
End Sub

'=============================
'====== UTILITIES ============
'=============================
Private Function GetTableAnywhere(ByVal tableName As String) As ListObject
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set GetTableAnywhere = lo
            Exit Function
        End If
    Next ws
End Function

Private Function RowToDict(lo As ListObject, idx As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Dim c As ListColumn, v
    For Each c In lo.ListColumns
        v = lo.DataBodyRange.Cells(idx, c.Index).value  ' raw variant (Date/Double/String)
        d(c.Name) = v
    Next c
    Set RowToDict = d
End Function

Private Function Nz(v As Variant) As String
    If IsError(v) Then
        Nz = ""
    ElseIf IsNull(v) Then
        Nz = ""
    ElseIf Len(CStr(v)) = 0 Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

Private Function ToISODate(ByVal s$) As String
    On Error GoTo Fallback
    If Len(s) = 0 Then ToISODate = "": Exit Function
    Dim d As Date: d = CDate(s)
    ToISODate = Format$(d, "yyyy-mm-dd")
    Exit Function
Fallback:
    ToISODate = s
End Function

Private Function CombinePath3(ByVal root As String, ByVal subp As String, ByVal leaf As String) As String
    Dim p As String: p = root
    If Len(subp) > 0 Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        p = p & subp
    End If
    If Len(leaf) > 0 Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        p = p & leaf
    End If
    CombinePath3 = p
End Function

Private Sub EnsureFolder(ByVal path$)
    If Len(path) = 0 Then Exit Sub
    Dim parts, i&, cur$
    parts = Split(path, "\")
    If UBound(parts) < 0 Then Exit Sub
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

Private Function FileExists(ByVal p$) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(p, vbNormal)) > 0)
    On Error GoTo 0
End Function

Private Function ReadAllText(ByVal fullPath$) As String
    Dim f As Integer: f = FreeFile
    Open fullPath For Input As #f
    ReadAllText = Input$(LOF(f), f)
    Close #f
End Function

Private Sub WriteTextSafe(ByVal fullPath$, ByVal content$)
    On Error GoTo Fail
    EnsureFolder Left$(fullPath, InStrRev(fullPath, "\") - 1)
    Dim f As Integer: f = FreeFile
    Open fullPath For Output As #f
    Print #f, content;
    Close #f
    Exit Sub
Fail:
    MsgBox "Failed to write file:" & vbCrLf & fullPath & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant: bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    Dim i As Long
    For i = LBound(bad) To UBound(bad): s = Replace$(s, bad(i), " "): Next i
    s = Trim$(s)
    Do While InStr(s, "  ") > 0: s = Replace$(s, "  ", " "): Loop
    If Len(s) = 0 Then s = "untitled"
    SanitizeFileName = s
End Function

Private Function UrlEncode(ByVal s$) As String
    Dim i&, ch$, o$
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126, 47: o = o & ch
            Case 32: o = o & "%20"
            Case Else: o = o & "%" & Right$("0" & Hex(AscW(ch)), 2)
        End Select
    Next i
    UrlEncode = o
End Function

Private Function RegexGet(ByVal text$, ByVal pattern$) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.Global = False
    re.IgnoreCase = True
    If re.Test(text) Then
        Set m = re.Execute(text)(0)
        RegexGet = m.SubMatches(0)
    Else
        RegexGet = ""
    End If
End Function

Private Function InArray(val As Variant, arr As Variant) As Boolean
    Dim i As Long
    If Not IsArray(arr) Then
        InArray = False
        Exit Function
    End If
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            InArray = True
            Exit Function
        End If
    Next i
    InArray = False
End Function

'-----------------------
' DATE HELPERS
'-----------------------
Private Function DateForYaml(ByVal v As Variant) As String
    On Error GoTo Fallback

    If VarType(v) = vbDate Then
        DateForYaml = Format$(CDate(v), "yyyy-mm-dd")
        Exit Function
    End If

    If IsNumeric(v) Then
        Dim n As Double: n = CDbl(v)
        If n >= 29221 And n < 60000 Then
            DateForYaml = Format$(CDate(n), "yyyy-mm-dd")
            Exit Function
        Else
            DateForYaml = ""
            Exit Function
        End If
    End If

    Dim s$: s = Trim$(CStr(v))
    If LCase$(s) = "mm/dd/yyyy" Or s = "" Then
        DateForYaml = ""
        Exit Function
    End If

    DateForYaml = ToISODateSmart(s)
    Exit Function
Fallback:
    DateForYaml = ""
End Function

Private Function ToISODateSmart(ByVal v As Variant) As String
    On Error GoTo Fallback

    Dim s$: s = Trim$(CStr(v))
    If s = "" Then
        ToISODateSmart = ""
        Exit Function
    End If

    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        ToISODateSmart = s
        Exit Function
    End If

    s = Replace$(s, "T", " ")

    If IsNumeric(s) Then
        Dim n As Double: n = CDbl(s)
        If n >= 29221 And n < 60000 Then
            ToISODateSmart = Format$(CDate(n), "yyyy-mm-dd")
            Exit Function
        Else
            ToISODateSmart = ""
            Exit Function
        End If
    End If

    ToISODateSmart = Format$(CDate(s), "yyyy-mm-dd")
    Exit Function

Fallback:
    ToISODateSmart = ""
End Function

Private Function ISODateOnly(ByVal s As String) As String
    Dim t$: t = Trim$(s)
    If Len(t) >= 10 And Mid$(t, 5, 1) = "-" And Mid$(t, 8, 1) = "-" Then
        ISODateOnly = Left$(t, 10)
    Else
        On Error Resume Next
        ISODateOnly = Format$(CDate(t), "yyyy-mm-dd")
        If Err.Number <> 0 Then ISODateOnly = ""
        On Error GoTo 0
    End If
End Function



Option Explicit

Type Point
    X As Double
    Y As Double
End Type

Type PointUTM
    E As Double
    N As Double
End Type

Public NextRunTime As Date

Private Sub Workbook_Open()
    Call StartTimer
End Sub

Private Sub Workbook_BeforeClose()
    Call StopTimer
End Sub

'This conversion script was written by Anthony Imberi (anthony.imberi@noaa.gov)
'Written August 2025
'This script can be used to convert Hypack LNW files into route files that can be imported to a Sperry ECDIS
Sub readLnw(Optional inputFile As String)
    Dim name As String, prefix As String, suffix As String, speed As Double, radius As Double, dpm As Double, xtd As Double, reverse As Boolean, utmZone As Integer
    Dim lineName As String, wptName As String, lat As Double, lon As Double
    'Dim fso As Object, ts As Object
    Dim textLine As String, text As String
    Dim foundRoutes() As Integer, foundWpts() As Integer, routeCount As Integer, wptCount As Integer, strPosition As Integer, rteEndPos As Integer
    Dim routeData() As Variant
    
    'Ask for the source file
    If inputFile = "" Then
        name = Application.GetOpenFilename("Hypack Line Files (*.lnw),*.lnw,All Files (*.*),*.*")
    Else
        name = inputFile
    End If
    
    
    If name = "False" Then GoTo EndEarly
    
    'Load in the settings from the spreadsheet
    prefix = CStr(Range("C2").Value)
    suffix = CStr(Range("C3").Value)
    speed = CDbl(Range("C4").Value)
    radius = CDbl(Range("C5").Value)
    xtd = CDbl(Range("C6").Value)
    reverse = Range("C7").Value = "Yes"

    dpm = RotCalc(speed, radius) ' Calculate the degrees per minute at the given speed and turn radius
    
    'Open the source file and read its contents into a string
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'Set ts = fso.OpenTextFile(name, ForReading)
    
    'Do While Not ts.AtEndOfStream
    '    textLine = ts.ReadLine
    '    text = text + textLine + vbCrLf
    'Loop
    
    'ts.Close
    'Set ts = Nothing
    'Set fso = Nothing
    
    'Open the source file and read its contents into a string
    Open name For Input As #1
        Do Until EOF(1)
            Line Input #1, textLine
            text = text + textLine + vbCrLf
        Loop
    Close #1
    
    If InStr(1, text, "ZON ", vbTextCompare) <= 0 Then
        utmZone = CInt(Range("C8").Value)
    Else
        utmZone = CInt(ExtractNumbers(Mid(text, InStr(1, text, "ZON ", vbTextCompare) + 4, 2)))
    End If
    
    strPosition = 1
    routeCount = 0
    
    'Search for all of the routes in the source file
    Do While InStr(strPosition, text, "LIN ", vbTextCompare) > 0
        strPosition = InStr(strPosition, text, "LIN ", vbTextCompare)
        routeCount = routeCount + 1
        ReDim Preserve foundRoutes(1 To routeCount)
        foundRoutes(routeCount) = strPosition
        strPosition = strPosition + Len("LIN ")
    Loop
    
    ReDim routeDetails(1 To routeCount + 1, 1 To 1)
    Dim maxWpt As Integer
    maxWpt = 0
    
    Dim namePos As Integer, nameClosePos As Integer, searchStrA As String, searchStrB As String
    Dim i As Integer
    For i = 1 To routeCount
        searchStrA = "LNN "
        searchStrB = vbCrLf
        namePos = InStr(foundRoutes(i), text, searchStrA, vbTextCompare) + Len(searchStrA)
        nameClosePos = InStr(namePos, text, searchStrB)
        lineName = Mid(text, namePos, nameClosePos - namePos)
        
        'Set the route names
        If (IsNumeric(lineName)) Then
            If CInt(lineName) < 10 And CInt(lineName) >= 0 Then
                lineName = "0" & lineName
            ElseIf CInt(lineName) > -10 And CInt(lineName) < 0 Then
                lineName = "-0" & Abs(CInt(lineName))
            End If
        End If
        
        If prefix <> "" Then lineName = prefix & lineName
        If suffix <> "" Then lineName = lineName & suffix
        
        'Find all of the waypoints in all of the routes
        wptCount = 0
        strPosition = foundRoutes(i)
        rteEndPos = InStr(foundRoutes(i), text, "EOL", vbTextCompare)
        Do While InStr(strPosition, text, "PTS ", vbTextCompare) > 0 And InStr(strPosition, text, "PTS ", vbTextCompare) < rteEndPos
            strPosition = InStr(strPosition, text, "PTS ", vbTextCompare) + Len("PTS ")
            wptCount = wptCount + 1
            ReDim Preserve foundWpts(1 To wptCount)
            foundWpts(wptCount) = strPosition
            strPosition = strPosition + Len("PTS ")
        Loop
        
        If (maxWpt < wptCount) Then
            ReDim Preserve routeDetails(1 To routeCount + 1, 1 To (wptCount * 2) + 1)
            maxWpt = wptCount
        End If
        
        routeDetails(1 + i, 1) = lineName
        
        'Format the waypoints and then convert the easting and northing coordinates to Lat/Lon
        Dim j As Integer, eastPos As Integer, northPos As Integer, eastDigits As Integer, northDigits As Integer, easting As Double, northing As Double, eastingPrev As Double, northingPrev As Double, latlon() As Double, pt1 As Point, pt2 As Point
        Dim jCount As Integer
        jCount = 1 ' This will count the waypoints that are saved and not the ones that are skipped
        For j = 1 To wptCount
            If (j < 10) Then
                wptName = "WPT 0" & j ' Add a zero to single digit waypoint names
            Else
                wptName = "WPT " & j
            End If

            routeDetails(1, j * 2) = wptName & " Lat"
            routeDetails(1, 1 + (j * 2)) = wptName & " Lon"
            
            eastPos = foundWpts(j)
            eastDigits = InStr(eastPos, text, " ", vbTextCompare) - eastPos
            northPos = InStr(eastPos, text, " ", vbTextCompare) + 1
            northDigits = InStr(northPos, text, vbCrLf, vbTextCompare) - northPos
            
            easting = CDbl(Mid(text, eastPos, eastDigits))
            northing = CDbl(Mid(text, northPos, northDigits))
            
            latlon = UTMtoLatLon(easting, northing, utmZone)
            
            lat = latlon(0)
            lon = latlon(1)
            
            If jCount > 1 Then
                pt1.Y = northingPrev ' Northing of previous waypoint
                pt1.X = eastingPrev ' Easting of previous waypoint
                pt2.Y = northing
                pt2.X = easting
                
                If MetersToNM(Distance(pt1, pt2)) > radius Then
                    routeDetails(1 + i, jCount * 2) = lat ' Save Lat if the distance between points exceeds the turn radius
                    routeDetails(1 + i, 1 + (jCount * 2)) = lon ' Save Lon if the distance between points exceeds the turn radius
                    jCount = jCount + 1
                Else
                    Debug.Print "WPT " & j & " skipped because the turn radius is greater than its distance to the previous waypoint"
                    Debug.Print "distance:", MetersToNM(Distance(pt1, pt2)), "radius:", radius
                End If
            
            Else
                routeDetails(1 + i, jCount * 2) = lat ' Save the Lat of the first waypoint
                routeDetails(1 + i, 1 + (jCount * 2)) = lon ' Save the Lon of the first waypoint
                jCount = jCount + 1
            End If
            
            northingPrev = northing
            eastingPrev = easting
        Next j
    Next i
    routeDetails(1, 1) = "Line Name"
    
    ' If there are any line names that are identical, add a sequential number to them to differentiate
    Dim duplicates As Integer
    For i = LBound(routeDetails, 1) To UBound(routeDetails, 1)
        duplicates = 0
        For j = LBound(routeDetails, 1) To UBound(routeDetails, 1)
            If routeDetails(i, 1) = routeDetails(j, 1) Then
                duplicates = duplicates + 1
                If duplicates > 1 Then routeDetails(j, 1) = routeDetails(j, 1) & "." & duplicates
            End If
        Next j
        If duplicates > 1 Then routeDetails(i, 1) = routeDetails(i, 1) & ".1"
    Next i
    
    ' Now work on setting up the output files and write to them
    Dim path As String, outfileName As String, outfileNameRev As String
    If (InStrRev(name, "/") > 0) Then 'This enables compatibility with either Mac/Linux or Windows
        path = Left(name, InStrRev(name, "/"))
    Else
        path = Left(name, InStrRev(name, "\"))
    End If
    
    If inputFile = "" And Range("C9").Value = "Yes" Then
        DeleteFile (path & "*.rtz")
    End If
    
    Dim r As Integer, C As Integer, routeLaz() As Integer
    ReDim routeLaz(1 To routeCount + 1)
    For r = LBound(routeDetails, 1) To UBound(routeDetails, 1)
        If r = 1 Then
            GoTo NextRowIteration1
        End If
        
        If routeDetails(r, 2) = "" Then
            GoTo NextRowIteration1
        End If
        
        ' Calculate a distance-weighted average azimuth for route naming.
        routeLaz(r) = (CInt(averageLaz(routeDetails, r, maxWpt)) + 360) Mod 360
        
NextRowIteration1:
    Next r
    
    ' Disable the line combination algorythm because it's pretty buggy
    'Dim n As Integer, wpt0 As Point, wpt1 As Point, wpt2 As Point
    'n = 3
    'Do While n <= routeCount + 1
    '    wpt0.Y = CDbl(routeDetails(n - 1, (lineWptNum(n - 1) - 1) * 2))
    '    wpt0.X = CDbl(routeDetails(n - 1, ((lineWptNum(n - 1) - 1) * 2) + 1))
    '    wpt1.Y = CDbl(routeDetails(n - 1, lineWptNum(n - 1) * 2))
    '    wpt1.X = CDbl(routeDetails(n - 1, (lineWptNum(n - 1) * 2) + 1))
    '    wpt2.Y = CDbl(routeDetails(n, 2))
    '    wpt2.X = CDbl(routeDetails(n, 3))
    '    If Abs(startLaz(n) - endLaz(n - 1)) <= 20 Then
    '        If Abs(calcLaz(wpt2.Y, wpt2.X, wpt1.Y, wpt1.X) - startLaz(n)) <= 20 _
    '            Or CrossTrackDistance(wpt0, wpt1, wpt2) <= 20 Then
    '            routeDetails(n - 1, 1) = routeDetails(n - 1, 1) & "-" & ExtractNumbers(routeDetails(n, 1))
    '            routeDetails(n, 1) = "$$Skip$$"
    '            If lineWptNum(n) + lineWptNum(n - 1) > maxWpt Then
    '                maxWpt = lineWptNum(n) + lineWptNum(n - 1)
    '                ReDim Preserve routeDetails(1 To routeCount + 1, 1 To (maxWpt * 2) + 1)
    '            End If
    '
    '            i = 1
    '            Do While i <= lineWptNum(n)
    '                If i <> 1 Or Distance(wpt1, wpt2) >= 50 Then
    '                    routeDetails(n - 1, (lineWptNum(n - 1) * 2) + (i * 2)) = routeDetails(n, i * 2)
    '                    routeDetails(n - 1, (lineWptNum(n - 1) * 2) + (i * 2) + 1) = routeDetails(n, (i * 2) + 1)
    '                End If
    '                i = i + 1
    '            Loop
    '        End If
    '    End If
    '    n = n + 1
    'Loop
    
    
    For r = LBound(routeDetails, 1) To UBound(routeDetails, 1)
        If r = 1 Then
            GoTo NextRowIteration2
        End If
        
        If routeDetails(r, 2) = "" Then
            GoTo NextRowIteration2
        End If
    
        If routeDetails(r, 1) = "$$Skip$$" Then
            GoTo NextRowIteration2
        End If
        
        Dim lazBaseStr As String, lazRevStr As String
        lazBaseStr = CStr(routeLaz(r))
        lazRevStr = CStr(((routeLaz(r) + 180) + 360) Mod 360)
        
        Do While Len(lazBaseStr) < 3
            lazBaseStr = "0" & lazBaseStr
        Loop
        Do While Len(lazRevStr) < 3
            lazRevStr = "0" & lazRevStr
        Loop
        
        If routeDetails(r, 1) = "" Then
            If r < 10 Then
                routeDetails(r, 1) = "0" & r
            Else
                routeDetails(r, 1) = r
            End If
            If prefix <> "" Then routeDetails(r, 1) = prefix & routeDetails(r, 1)
            If suffix <> "" Then routeDetails(r, 1) = routeDetails(r, 1) & suffix
        End If
        
        outfileName = path & routeDetails(r, 1) & " LAZ " & lazBaseStr & ".rtz"
        outfileNameRev = path & routeDetails(r, 1) & " LAZ " & lazRevStr & ".rtz"
        
        'Open the output file and the reverse output file, if requested
        Dim fileNum1 As Integer, fileNum2 As Integer
        
        fileNum1 = FreeFile()
        If IsFileOpen(outfileName) Then outfileName = outfileName & ".rtz"
        
        Open outfileName For Output As #fileNum1
        
        fileNum2 = FreeFile()
        
        If reverse Then
            If IsFileOpen(outfileNameRev) Then outfileNameRev = outfileNameRev & ".rtz"
            Open outfileNameRev For Output As #fileNum2
        End If
        
        'Print the XML header
        Print #fileNum1, "<?xml version=""1.0"" encoding=""utf-8""?>"
        Print #fileNum1, "<route version=""1.0"" xmlns=""http://www.cirm.org/RTZ/1/0"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.cirm.org/RTZ/1/0/rtz.xsd"">"
        
        If reverse Then
            Print #fileNum2, "<?xml version=""1.0"" encoding=""utf-8""?>"
            Print #fileNum2, "<route version=""1.0"" xmlns=""http://www.cirm.org/RTZ/1/0"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.cirm.org/RTZ/1/0/rtz.xsd"">"
        End If
  
        'Print the Route Name
        Print #fileNum1, vbTab & "<routeInfo routeName=""" & routeDetails(r, 1) & " LAZ " & lazBaseStr & """ />"
        
        If reverse Then Print #fileNum2, vbTab & "<routeInfo routeName=""" & routeDetails(r, 1) & " LAZ " & lazRevStr & """ />"
            
        'Print the section header and the default waypoint info
        Print #fileNum1, vbTab & "<waypoints>"
        Print #fileNum1, vbTab & vbTab & "<defaultWaypoint radius=""1"">"
        Print #fileNum1, vbTab & vbTab & vbTab & "<leg starboardXTD=""0.0539956803455724"" portsideXTD=""0.0539956803455724"" geometryType=""Loxodrome"" speedMin=""6"" speedMax=""13"" />"
        Print #fileNum1, vbTab & vbTab & "</defaultWaypoint>"
        
        If reverse Then
            Print #fileNum2, vbTab & "<waypoints>"
            Print #fileNum2, vbTab & vbTab & "<defaultWaypoint radius=""1"">"
            Print #fileNum2, vbTab & vbTab & vbTab & "<leg starboardXTD=""0.0539956803455724"" portsideXTD=""0.0539956803455724"" geometryType=""Loxodrome"" speedMin=""6"" speedMax=""13"" />"
            Print #fileNum2, vbTab & vbTab & "</defaultWaypoint>"
        End If
            
        Dim colLat As Double, colLon As Double, wptNum As Integer
        
        For C = LBound(routeDetails, 2) To UBound(routeDetails, 2)
            If C = 1 Then
                wptNum = 1
            ElseIf C Mod 2 = 0 Then
                If routeDetails(r, C) = "" Then GoTo NextColIteration
                colLat = routeDetails(r, C)
            ElseIf C Mod 2 = 1 Then
                If routeDetails(r, C) = "" Then GoTo NextColIteration
                colLon = routeDetails(r, C)
                PrintWaypoint wptNum, "WPT " & wptNum, colLat, colLon, speed, xtd, radius, fileNum1
                wptNum = wptNum + 1
            End If
            
NextColIteration:
        Next C
        
        If reverse Then
            For C = UBound(routeDetails, 2) To LBound(routeDetails) Step -1
                If C = UBound(routeDetails, 2) Then wptNum = 1
                If C = LBound(routeDetails, 2) Then
                    GoTo NextColRevIteration
                ElseIf C Mod 2 = 0 Then
                    If routeDetails(r, C) = "" Then GoTo NextColRevIteration
                    colLat = routeDetails(r, C)
                    
                    Dim wptNm As String
                    If True Then
                        If C < 10 Then
                            wptNm = "WPT 0" & wptNum
                        Else
                            wptNm = "WPT " & wptNum
                        End If
                    Else
                        wptNm = Left(routeDetails(1, C), Len(routeDetails(1, C)) - 4)
                    End If
                    
                    PrintWaypoint wptNum, wptNm, colLat, colLon, speed, xtd, radius, fileNum2
                    wptNum = wptNum + 1
                ElseIf C Mod 2 = 1 Then
                    If routeDetails(r, C) = "" Then GoTo NextColRevIteration
                    colLon = routeDetails(r, C)
                End If
                
NextColRevIteration:
            Next C
        End If
        
        'Close the XML section
        Print #fileNum1, vbTab & "</waypoints>"
        
        If reverse Then Print #fileNum2, vbTab & "</waypoints>"
        
        'Print the Schedule section
        Dim mytab As String
        mytab = vbTab & vbTab  'Double tab
        Print #fileNum1, vbTab & "<schedules>"
        Print #fileNum1, mytab & "<schedule id=""1"" name=""Default Schedule"">"
        Print #fileNum1, mytab & vbTab & "<manual>"
        
        If reverse Then
            Print #fileNum2, vbTab & "<schedules>"
            Print #fileNum2, mytab & "<schedule id=""1"" name=""Default Schedule"">"
            Print #fileNum2, mytab & vbTab & "<manual>"
        End If
        
        Dim s As Integer
        For s = 1 To wptNum - 1
            If s = 1 Then
                Print #fileNum1, mytab & mytab & "<sheduleElement waypointId=""1"" etd=""" & Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & ".0000000+00:00"" />"
                If reverse Then Print #fileNum2, mytab & mytab & "<sheduleElement waypointId=""1"" etd=""" & Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & ".0000000+00:00"" />"
            Else
                Print #fileNum1, mytab & mytab & "<sheduleElement waypointId=""" & s & """ speed=""" & speed & """ />"
                If reverse Then Print #fileNum2, mytab & mytab & "<sheduleElement waypointId=""" & s & """ speed=""" & speed & """ />"
            End If
        Next s
        
          'Close out the section in the XML
        Print #fileNum1, mytab & vbTab & "</manual>"
        Print #fileNum1, mytab & "</schedule>"
        Print #fileNum1, vbTab & "</schedules>"
        
        If reverse Then
            Print #fileNum2, mytab & vbTab & "</manual>"
            Print #fileNum2, mytab & "</schedule>"
            Print #fileNum2, vbTab & "</schedules>"
        End If
        
        'Print the Last part of the file including the Description of the Route
        Print #fileNum1, vbTab & "<extensions>"
        Print #fileNum1, mytab & "<extension manufacturer=""Sperry"" name=""AdditionalRouteData"">"
        Print #fileNum1, mytab & vbTab & "<AdditionalRouteData LastModified=""" & Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & ".0000000+00:00"">"
        Print #fileNum1, mytab & mytab & "<Description>" & routeDetails(r, 1) & " LAZ " & lazBaseStr & "</Description>"
        Print #fileNum1, mytab & mytab & "<Notes />"
        Print #fileNum1, mytab & vbTab & "</AdditionalRouteData>"
        Print #fileNum1, mytab & "</extension>"
        Print #fileNum1, vbTab & "</extensions>"
        Print #fileNum1, "</route>"
        
        Close #fileNum1
        
        If reverse Then
            Print #fileNum2, vbTab & "<extensions>"
            Print #fileNum2, mytab & "<extension manufacturer=""Sperry"" name=""AdditionalRouteData"">"
            Print #fileNum2, mytab & vbTab & "<AdditionalRouteData LastModified=""" & Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & ".0000000+00:00"">"
            Print #fileNum2, mytab & mytab & "<Description>" & routeDetails(r, 1) & " LAZ " & lazRevStr & "</Description>"
            Print #fileNum2, mytab & mytab & "<Notes />"
            Print #fileNum2, mytab & vbTab & "</AdditionalRouteData>"
            Print #fileNum2, mytab & "</extension>"
            Print #fileNum2, vbTab & "</extensions>"
            Print #fileNum2, "</route>"
        
            Close #fileNum2
        End If
        
NextRowIteration2:
    Next r
    
    Dim msg As String
    If reverse Then
        msg = "Successfully saved " & routeCount & " routes and " & routeCount & " reversed routes."
    Else
        msg = "Successfully saved " & routeCount & " routes."
    End If

    msg = msg & vbCrLf & vbCrLf & "Found routes:" & vbCrLf

    Dim ct As Integer
    ct = 1
    while ct < routeCount + 1
        If routeDetails(ct, 2) <> "" Then
            msg = msg & vbCrLf & routeDetails(ct, 1)
        End If
        ct = ct + 1
    Wend
    
    MsgBox (msg)
    
EndEarly:
    
End Sub

Sub StartTimer()
    NextRunTime = Now + TimeValue("00:01:00")
    Application.OnTime NextRunTime, "SaveAutolineData"
End Sub

Sub SaveAutolineData()
    On Error GoTo AutolineError
    Dim projectFolder As String, outputFolder As String, recentFile As String
    
    Debug.Print "Start autoline backup"
    
    projectFolder = Sheets("Autoline Extractor").Range("F2").Value & "\Raw\"
    projectFolder = Replace(projectFolder, "\\", "\") ' Fix for if user gave folder with \ at the end
    outputFolder = Environ("USERPROFILE") & "\Desktop\Sperry\Hypack-Sperry-Tools\Autolines\"
    recentFile = GetMostRecentRawFile(projectFolder)
    
    Debug.Print "Reading file: " & recentFile
    
    Dim fso As Object, file As textStream, keepSearching As Boolean, startWriting As Boolean, count As Integer
    Dim textLine As String, utmZone As Variant, centMer As Integer, parts As Variant, fileTime As Date, timeStr As String, text As String, tndLine As String, numLines As Integer
    keepSearching = True
    startWriting = False
    utmZone = False
    text = ""
    count = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(recentFile, ForReading)
    
    Do While keepSearching
        textLine = file.ReadLine
        If textLine Like "INI ZoneName=*" Then
            utmZone = CInt(ExtractNumbers(Mid(textLine, 19, 2)))
        End If
        If textLine Like "INI CentralMeridian=*" Then
            centMer = CInt(ExtractNumbers(Mid(textLine, Len("INI CentralMeridian="), 4)))
        End If
        If textLine Like "TND*" Then
            tndLine = textLine
            parts = Split(textLine)
            If UBound(parts) >= 2 Then
                fileTime = CDate(parts(2) & " " & parts(1))
                timeStr = Format(fileTime, "HH.nn yyyy-mm-dd")
            End If
        End If
        If textLine Like "LIN #*" Then
            startWriting = True
            numLines = CInt(ExtractNumbers(Mid(textLine, Len("LIN "))))
        End If
        If startWriting Then
            If textLine Like "LBP*" Then
                ' Do Nothing
            ElseIf textLine Like "LNN*" Then
                text = text & textLine & " " & timeStr & vbCrLf
            Else
                text = text & textLine & vbCrLf
                count = count + 1
            End If
        End If
        If textLine Like "EOL*" Then keepSearching = False
    Loop
    
    file.Close
    
    parts = Split(recentFile, "\")
    recentFile = parts(UBound(parts))
    parts = Split(recentFile, ".")
    recentFile = parts(LBound(parts))
    
    If fso.FileExists(outputFolder & recentFile & " " & timeStr & ".lnw") Then
        Set file = fso.OpenTextFile(outputFolder & recentFile & " " & timeStr & ".lnw", ForReading)
    Else
        GoTo SkipCompare
    End If
    
    Dim prevLines As Integer
    keepSearching = True
    startWriting = False
    prevLines = 0
    
    Do While keepSearching
        textLine = file.ReadLine
        If textLine Like "LIN*" Then
            prevLines = CInt(ExtractNumbers(Mid(textLine, Len("LIN "))))
            keepSearching = False
        ElseIf textLine Like "EOL*" Then keepSearching = False
        End If
    Loop
    
    file.Close
    
    Dim completeTime As String
    completeTime = Format(Now, "HH:nn yyyy-mm-dd")
    If numLines <= prevLines Then
        Debug.Print "No updates found"
        Sheets("Autoline Extractor").Range("K7").Value = completeTime
        GoTo AutolineError
    End If
    
SkipCompare:
    
    Set file = fso.CreateTextFile(outputFolder & recentFile & " " & timeStr & ".lnw", True)
    
    If utmZone = False Then
        utmZone = CInt(WorksheetFunction.Floor((centMer + 180) / 6, 1) + 1)
    End If
    
    file.WriteLine "ZON " & utmZone
    file.WriteLine tndLine
    file.WriteLine ("LNS 1")
    file.Write text
    
    
    completeTime = Format(Now, "HH:nn yyyy-mm-dd")
    Sheets("Autoline Extractor").Range("K6").Value = recentFile
    Sheets("Autoline Extractor").Range("K7").Value = completeTime
    Sheets("Autoline Extractor").Range("K8").Value = completeTime
    
    Debug.Print "Autoline backup complete at " & completeTime
    
AutolineError:
    Set file = Nothing
    Set fso = Nothing
    
    Call StartTimer
    completeTime = Format(Now, "HH:nn yyyy-mm-dd")
    Debug.Print "Exiting autoline backup at " & completeTime
End Sub

Sub StopTimer()
    On Error Resume Next
    Application.OnTime NextRunTime, "SaveAutolineData", , False
End Sub

Sub RestartTimer()
    Call StopTimer
    Call StartTimer
End Sub

Function GetMostRecentRawFile(fldPath As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim latestDate As Date
    Dim latestFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(fldPath) Then
        GetMostRecentRawFile = ""
        Exit Function
    End If
    
    Set folder = fso.GetFolder(fldPath)
    
    ' Loop through files
    For Each file In folder.Files
        ' Check if extension is "RAW" (UCase handles .raw or .RAW)
        If UCase(fso.GetExtensionName(file.name)) = "RAW" Then
            If file.DateLastModified > latestDate Then
                latestDate = file.DateLastModified
                latestFileName = file.path
            End If
        End If
    Next file
    
    GetMostRecentRawFile = latestFileName
    
    Set fso = Nothing
    Set folder = Nothing
End Function

Function IsFileOpen(fileName As String) As Boolean
    On Error Resume Next
    Dim ff As Long
    ff = FreeFile
    Open fileName For Binary Access Read Write Lock Read Write As #ff
    Close #ff
    If Err.Number <> 0 Then IsFileOpen = True
    On Error GoTo 0
End Function

Private Function CheckName(nm) As String
   'need to do a regex replace for double quotes
   GoTo Skip
   On Error GoTo ErrorHandler
   Dim rx As New RegExp
   strPattern = """"""
   rx.Pattern = strPattern
   If rx.test(name) Then
     CheckName = rx.Replace(nm, "")
   Else
     CheckName = nm
   End If
ErrorHander:
   Err.Clear
Skip:
End Function

'Converts meters to nautical miles
Private Function MetersToNM(meters)
  MetersToNM = Format(meters / 1852, "0.000000000000000")
End Function

'Prints each waypoint
Private Sub PrintWaypoint(waypointID, waypointName, lat, lon, speed, xtd, radius, FileNum As Integer)
  Dim mytab As String, strName As String
  mytab = vbTab & vbTab 'double tab
  'strName = CheckName(waypointName)
  strName = waypointName
  If strName = "" Then
    If waypointID < 10 Then
        strName = "WPT 0" & waypointID
    Else
        strName = "WPT " & waypointID
    End If
  End If
  
  'If it's the first waypoint, there's no turn so no radius
  If waypointID = 1 Then
    Print #FileNum, mytab & "<waypoint id=""1"" name=""" & strName & """>"
  Else
    Print #FileNum, mytab & "<waypoint id=""" & waypointID & """ name=""" & strName & """ radius=""" & radius & """>"
  End If
  
  'Print the position
  Print #FileNum, mytab & vbTab & "<position lat=""" & lat & """ lon=""" & lon & """ />"
  
  'Print XTD in NM
  Dim xtdNM As String
  If waypointID > 1 Then
    xtdNM = MetersToNM(xtd)
    Print #FileNum, mytab & vbTab & "<leg starboardXTD=""" & xtdNM & """ portsideXTD=""" & xtdNM & """ />"
  End If
  
  'Determine the turn speeds based on the speed
  Dim turnSpeed As String, minSpeed As String, maxSpeed As String
  turnSpeed = Format(0.514444444444441 * speed, "0.000000000000000") 'Converts speed to what Sperry needs for some reason
  minSpeed = Format(0.514444444444441 * (speed - 2), "0.000000000000000")
  maxSpeed = Format(0.514444444444441 * (speed + 2), "0.000000000000000")
  
  'Print the turn info
  Print #FileNum, mytab & vbTab & "<extensions>"
  Print #FileNum, mytab & mytab & "<extension manufacturer=""Sperry"" name=""AdditionalWaypointData"">"
  Print #FileNum, mytab & mytab & vbTab & "<AdditionalWaypointData TurnSpeed=""" & turnSpeed & """ MinTurnSpeed=""" & minSpeed & """ MaxTurnSpeed=""" & maxSpeed & """ LeftOffTrackAlarmLimitForTurn=""" & xtd & """ RightOffTrackAlarmLimitForTurn=""" & xtd & """ />"
  Print #FileNum, mytab & mytab & "</extension>"
  Print #FileNum, mytab & mytab & "<extension manufacturer=""Sperry"" name=""AdditionalTrackData"">"
  Print #FileNum, mytab & mytab & vbTab & "<AdditionalTrackData OffTrackAlarmEnabledForTurn=""true"" OffTrackAlarmEnabledForDepartingLeg=""true"" />"
  Print #FileNum, mytab & mytab & "</extension>"
  Print #FileNum, mytab & vbTab & "</extensions>"
  
  'Close out the waypoint
  Print #FileNum, mytab & "</waypoint>"
End Sub

'Used to calculate the radius which will set the rate of turn at 100 degrees per minute
Private Function RadiusCalc(speed As Double, Optional dpm As Double)
    If dpm = 0 Then dpm = 100 ' dpm specifies the number of degrees per minute to turn
    RadiusCalc = Format(speed / 60 * (360 / dpm) / (2 * Application.WorksheetFunction.Pi()), "0.000000000000000")
End Function

Private Function RotCalc(speed As Double, radius As Double)
    RotCalc = Format(60 * (2 * Application.WorksheetFunction.Pi() * radius) / speed, "0.000000000000000")
End Function

' Distance in meters between two lat/lon points
Private Function Distance(p1 As Point, p2 As Point) As Double
    Dim utm1() As Double, utm2() As Double
    
    If p1.Y < 90 Then
        utm1 = LatLonToUTM(p1.Y, p1.X)
        utm2 = LatLonToUTM(p2.Y, p2.X)
    
        p1.Y = utm1(0)
        p1.X = utm1(1)
        p2.Y = utm2(0)
        p2.X = utm2(1)
    End If
    
    Distance = Math.Sqr((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
End Function

' Calculate crosstrack distance in meters of a lat/lon point on a line
Private Function CrossTrackDistance(wpt0 As Point, wpt1 As Point, wpt2 As Point)
    Dim pt0 As Point, pt1 As Point, pt2 As Point, utm0() As Double, utm1() As Double, utm2() As Double
    
    If wpt0.Y > 90 Then
        utm0 = LatLonToUTM(wpt0.Y, wpt0.X)
        utm1 = LatLonToUTM(wpt1.Y, wpt1.X)
        utm2 = LatLonToUTM(wpt2.Y, wpt2.X)
    
        pt0.Y = utm0(0)
        pt0.X = utm0(1)
        pt1.Y = utm1(0)
        pt1.X = utm1(1)
        pt2.Y = utm2(0)
        pt2.X = utm2(1)
    Else
        pt0.X = wpt0.X
        pt0.Y = wpt0.Y
        pt1.X = wpt1.X
        pt1.Y = wpt1.Y
        pt2.X = wpt2.X
        pt2.Y = wpt2.Y
    End If
    
    Dim dx As Double, dy As Double, mag2 As Double, t As Double, projX As Double, projY As Double
    
    dx = pt1.X - pt0.X
    dy = pt1.Y - pt0.Y
    mag2 = dx * dx + dy * dy
    
    If mag2 = 0 Then
        CrossTrackDistance = Math.Sqr((pt2.X - pt0.X) ^ 2 + (pt2.Y - pt0.Y) ^ 2)
    Else
        t = ((pt2.X - pt0.X) * dx + (pt2.Y - pt0.Y) * dy) / mag2
        projX = pt0.X + t * dx
        projY = pt0.Y + t * dy
        CrossTrackDistance = Math.Sqr((pt2.X - projX) ^ 2 + (pt2.Y - projY) ^ 2)
    End If
    
End Function

'Calculate the line azimuth between two coordinates (degrees)
Private Function calcLaz(latOne As Double, lonOne As Double, latTwo As Double, lonTwo As Double) As Double
    Dim phi1 As Double, phi2 As Double, lam1 As Double, lam2 As Double, dlam As Double
    Dim Y As Double, X As Double, az As Double

    ' Convert degrees to radians
    phi1 = latOne * WorksheetFunction.Pi / 180#
    phi2 = latTwo * WorksheetFunction.Pi / 180#
    lam1 = lonOne * WorksheetFunction.Pi / 180#
    lam2 = lonTwo * WorksheetFunction.Pi / 180#
    dlam = lam2 - lam1

    Y = Sin(dlam) * Cos(phi2)
    X = Cos(phi1) * Sin(phi2) - Sin(phi1) * Cos(phi2) * Cos(dlam)

    ' Excel ATAN2 takes (x,y)?  WorksheetFunction.Atan2(y, x) returns angle in radians
    az = WorksheetFunction.Atan2(X, Y) * 180# / WorksheetFunction.Pi
    If az < 0 Then az = az + 360#

    calcLaz = az
End Function

' Calculate a distance-weighted average azimuth for a full route.
' Uses leg length as the weight so longer legs contribute proportionally more.
Private Function averageLaz(routeDetails As Variant, routeRow As Integer, maxWpt As Integer) As Double
    Dim i As Integer
    Dim lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double
    Dim legAz As Double, legDist As Double
    Dim xComp As Double, yComp As Double
    Dim avgAz As Double
    Dim p1 As Point, p2 As Point

    xComp = 0
    yComp = 0

    For i = 2 To (maxWpt * 2) - 2 Step 2
        If routeDetails(routeRow, i) <> "" And routeDetails(routeRow, i + 1) <> "" And routeDetails(routeRow, i + 2) <> "" And routeDetails(routeRow, i + 3) <> "" Then
            lat1 = CDbl(routeDetails(routeRow, i))
            lon1 = CDbl(routeDetails(routeRow, i + 1))
            lat2 = CDbl(routeDetails(routeRow, i + 2))
            lon2 = CDbl(routeDetails(routeRow, i + 3))

            p1.Y = lat1
            p1.X = lon1
            p2.Y = lat2
            p2.X = lon2
            legDist = Distance(p1, p2)

            If legDist > 0 Then
                legAz = calcLaz(lat1, lon1, lat2, lon2) * WorksheetFunction.Pi / 180#
                xComp = xComp + Sin(legAz) * legDist
                yComp = yComp + Cos(legAz) * legDist
            End If
        End If
    Next i

    If xComp = 0 And yComp = 0 Then
        averageLaz = 0
        Exit Function
    End If

    avgAz = WorksheetFunction.Atan2(xComp, yComp) * 180# / WorksheetFunction.Pi
    If avgAz < 0 Then avgAz = avgAz + 360#
    averageLaz = avgAz
End Function

'Convert UTM to Lat/Lon Coordinates (WGS84)
Private Function UTMtoLatLon(E As Double, N As Double, utmZone As Integer) As Double()
    Dim zone As Integer
    Dim a As Double, f As Double, e2 As Double, ep2 As Double, kZero As Double
    Dim FE As Double, FN As Double, lambdaZero As Double
    Dim X As Double, Y As Double, M As Double, eOne As Double, mu As Double
    Dim phiOne As Double
    Dim sinPhi1 As Double, cosPhi1 As Double
    Dim TOne As Double, COne As Double, NOne As Double, ROne As Double, D As Double
    Dim latRad As Double, lonRad As Double
    Dim latlon(1) As Double

    zone = utmZone

    a = 6378137#
    f = 1 / 298.257223563
    e2 = f * (2 - f)                 ' eccentricity^2
    ep2 = e2 / (1 - e2)              ' second eccentricity^2
    kZero = 0.9996
    FE = 500000#
    FN = 0#                          ' Northern hemisphere only in your case; set to 10000000 for southern
    lambdaZero = (-183 + 6 * zone) * WorksheetFunction.Pi / 180

    X = E - FE
    Y = N - FN

    M = Y / kZero

    eOne = (1 - Sqr(1 - e2)) / (1 + Sqr(1 - e2))
    mu = M / (a * (1 - e2 / 4 - 3 * e2 ^ 2 / 64 - 5 * e2 ^ 3 / 256))

    phiOne = mu _
        + (3 * eOne / 2 - 27 * eOne ^ 3 / 32) * Sin(2 * mu) _
        + (21 * eOne ^ 2 / 16 - 55 * eOne ^ 4 / 32) * Sin(4 * mu) _
        + (151 * eOne ^ 3 / 96) * Sin(6 * mu) _
        + (1097 * eOne ^ 4 / 512) * Sin(8 * mu)

    sinPhi1 = Sin(phiOne)
    cosPhi1 = Cos(phiOne)

    TOne = Tan(phiOne) ^ 2
    COne = ep2 * (cosPhi1 ^ 2)                       ' <- fix: cos(phi)^2
    NOne = a / Sqr(1 - e2 * (sinPhi1 ^ 2))           ' <- fix: sin(phi)^2
    ROne = NOne * (1 - e2) / (1 - e2 * (sinPhi1 ^ 2))
    D = X / (NOne * kZero)

    latRad = phiOne - (NOne * Tan(phiOne) / ROne) * _
        (D ^ 2 / 2 _
         - (5 + 3 * TOne + 10 * COne - 4 * COne ^ 2 - 9 * ep2) * D ^ 4 / 24 _
         + (61 + 90 * TOne + 298 * COne + 45 * TOne ^ 2 - 252 * ep2 - 3 * COne ^ 2) * D ^ 6 / 720)

    lonRad = lambdaZero + _
        (D - (1 + 2 * TOne + COne) * D ^ 3 / 6 + (5 - 2 * COne + 28 * TOne - 3 * COne ^ 2 + 8 * ep2 + 24 * TOne ^ 2) * D ^ 5 / 120) / Cos(phiOne)

    latlon(0) = latRad * 180# / WorksheetFunction.Pi
    latlon(1) = lonRad * 180# / WorksheetFunction.Pi

    UTMtoLatLon = latlon
End Function

'Convert Lat/Lon to UTM Coordinates (WGS84)
Private Function LatLonToUTM(latDeg As Double, lonDeg As Double) As Double()
    Dim zone As Integer
    Dim a As Double, f As Double, e2 As Double, ep2 As Double, kZero As Double
    Dim FE As Double, FN As Double, lambdaZero As Double
    Dim latRad As Double, lonRad As Double
    Dim N As Double, t As Double, C As Double, bigA As Double
    Dim M As Double
    Dim E As Double, Nnorth As Double
    Dim utm(1) As Double
    
    ' Ellipsoid parameters (WGS84)
    a = 6378137#
    f = 1 / 298.257223563
    e2 = f * (2 - f)                  ' eccentricity^2
    ep2 = e2 / (1 - e2)               ' second eccentricity^2
    kZero = 0.9996
    FE = 500000#
    
    ' Determine UTM zone from longitude
    zone = WorksheetFunction.Floor((lonDeg + 180) / 6, 1) + 1
    lambdaZero = (-183 + 6 * zone) * WorksheetFunction.Pi / 180#
    
    ' Convert input to radians
    latRad = latDeg * WorksheetFunction.Pi / 180#
    lonRad = lonDeg * WorksheetFunction.Pi / 180#
    
    ' Precompute terms
    N = a / Sqr(1 - e2 * Sin(latRad) ^ 2)
    t = Tan(latRad) ^ 2
    C = ep2 * Cos(latRad) ^ 2
    bigA = (lonRad - lambdaZero) * Cos(latRad)
    
    ' Meridional arc
    M = a * ((1 - e2 / 4 - 3 * e2 ^ 2 / 64 - 5 * e2 ^ 3 / 256) * latRad _
        - (3 * e2 / 8 + 3 * e2 ^ 2 / 32 + 45 * e2 ^ 3 / 1024) * Sin(2 * latRad) _
        + (15 * e2 ^ 2 / 256 + 45 * e2 ^ 3 / 1024) * Sin(4 * latRad) _
        - (35 * e2 ^ 3 / 3072) * Sin(6 * latRad))
    
    ' Easting
    E = FE + kZero * N * (bigA + (1 - t + C) * bigA ^ 3 / 6 _
        + (5 - 18 * t + t ^ 2 + 72 * C - 58 * ep2) * bigA ^ 5 / 120)
    
    ' Northing (assuming northern hemisphere here)
    Nnorth = kZero * (M + N * Tan(latRad) * (bigA ^ 2 / 2 _
        + (5 - t + 9 * C + 4 * C ^ 2) * bigA ^ 4 / 24 _
        + (61 - 58 * t + t ^ 2 + 600 * C - 330 * ep2) * bigA ^ 6 / 720))
    
    ' Return [Easting, Northing]
    utm(0) = E
    utm(1) = Nnorth
    LatLonToUTM = utm
End Function

' Offset a polyline (UTM coordinates) left or right by given distance.
' side = "Left" for left, "Right" for right
' Inputs: arrays E() and N() of same length (polyline vertices)
' Returns: 2D array (2 x nPoints): result(0,i)=Easting, result(1,i)=Northing
Public Function OffsetPolyline(E() As Double, N() As Double, dist As Double, side As String) As Double()
    Dim smN As Long
    smN = UBound(E) - LBound(E) + 1
    If smN < 2 Then Err.Raise vbObjectError + 520, , "Polyline must have at least two points."
    
    Dim result() As Double
    ReDim result(1, LBound(E) To UBound(E))
    
    Dim i As Long
    Dim dx As Double, dy As Double, leng As Double
    Dim px As Double, py As Double
    Dim prevDx As Double, prevDy As Double
    Dim nextDx As Double, nextDy As Double
    Dim sideFactor As Double
    
    If side = "Left" Then
        sideFactor = -1
    ElseIf side = "Right" Then
        sideFactor = 1
    Else
        Err.Raise vbObjectError + 521, , "Side must be 'Left' or 'Right'."
    End If
    
    ' Loop through vertices
    For i = LBound(E) To UBound(E)
        If i = LBound(E) Then
            ' First point: use segment i->i+1
            dx = E(i + 1) - E(i)
            dy = N(i + 1) - N(i)
        ElseIf i = UBound(E) Then
            ' Last point: use segment i-1->i
            dx = E(i) - E(i - 1)
            dy = N(i) - N(i - 1)
        Else
            ' Interior point: average segment before and after
            prevDx = E(i) - E(i - 1)
            prevDy = N(i) - N(i - 1)
            nextDx = E(i + 1) - E(i)
            nextDy = N(i + 1) - N(i)
            dx = prevDx + nextDx
            dy = prevDy + nextDy
        End If
        
        leng = Sqr(dx * dx + dy * dy)
        If leng = 0 Then
            px = 0: py = 0
        Else
            ' Unit perpendicular vector
            px = sideFactor * dy / leng
            py = sideFactor * -dx / leng
        End If
        
        ' Scale and offset
        result(0, i) = E(i) + px * dist
        result(1, i) = N(i) + py * dist
    Next i
    
    OffsetPolyline = result
End Function


Function ShiftLine()
    Dim fileName As String, path As String, meters As Double, rl As String, newName As String
    Dim text As String, textLine As String, strPosition As Integer, wptCount As Integer, wptPositions() As Integer, wptDetails() As Double
    
    fileName = Application.GetOpenFilename("Sperry Route Files (*.rtz),*.rtz,All Files (*.*),*.*")
    
    If fileName = "False" Then GoTo EndEarly
    
    meters = CDbl(Range("C3").Value)
    rl = Range("C4").Value
    
    If Range("C3").Value = "" Or rl = "" Then
        MsgBox ("You must specify a number of meters to shift and a direction")
        GoTo EndEarly
    End If
    
    ' Separate the filename from the path and configure the new file name
    If (InStrRev(fileName, "/") > 0) Then 'This enables compatibility with either Mac/Linux or Windows
        path = Left(fileName, InStrRev(fileName, "/"))
        fileName = Right(fileName, Len(fileName) - InStrRev(fileName, "/"))
        newName = Left(fileName, Len(fileName) - 4) & " Shift " & CInt(meters) & "m " & rl & ".rtz"
    Else
        path = Left(fileName, InStrRev(fileName, "\"))
        fileName = Right(fileName, Len(fileName) - InStrRev(fileName, "\"))
        newName = "(" & Left(rl, 1) & meters & "m) " & Left(fileName, Len(fileName) - 4) & ".rtz"
    End If

    ' Read the route file into memory
    Open path & fileName For Input As #1
        Do Until EOF(1)
            Line Input #1, textLine
            text = text + textLine + vbCrLf
        Loop
    Close #1
    
    strPosition = InStr(1, text, "<waypoints", vbTextCompare)
    wptCount = 0
    
    ' Search for all waypoints in the route file
    Do While InStr(strPosition, text, "<position", vbTextCompare) > 0
        strPosition = InStr(strPosition, text, "<position", vbTextCompare)
        wptCount = wptCount + 1
        ReDim Preserve wptPositions(1 To wptCount)
        wptPositions(wptCount) = strPosition
        strPosition = strPosition + Len("<position")
    Loop
    
    ReDim wptDetails(0 To 1, 0 To wptCount - 1)
    
    Dim E() As Double, N() As Double, shifted() As Double, conversion() As Double
    Dim count As Integer, latPos As Integer, lat As Double, lonPos As Integer, lon As Double
    ReDim E(0 To wptCount - 1)
    ReDim N(0 To wptCount - 1)
    
    ' Save all of the waypoints as Lat/Lon and also as UTM Easting/Northing
    count = 0
    Do While count < wptCount
        count = count + 1
        latPos = InStr(wptPositions(count), text, "lat=""", vbTextCompare) + 5
        lat = Mid(text, latPos, InStr(latPos, text, """", vbTextCompare) - latPos)
        lonPos = InStr(wptPositions(count), text, "lon=""", vbTextCompare) + 5
        lon = Mid(text, lonPos, InStr(lonPos, text, """", vbTextCompare) - lonPos)
        
        wptDetails(0, count - 1) = lat
        wptDetails(1, count - 1) = lon
        conversion = LatLonToUTM(lat, lon)
        E(count - 1) = conversion(0)
        N(count - 1) = conversion(1)
    Loop
    
    ' Shift the UTM coordinates right or left by the number of meters specified
    shifted = OffsetPolyline(E, N, meters, rl)
    
    Dim newLatLon() As Double, latlon() As Double, zone As Integer
    ReDim newLatLon(0 To 1, 0 To wptCount - 1)
    ReDim latlon(0 To 1)
    
    ' Convert the UTM coordinates back to Lat/Lon
    count = 0
    Do While count < wptCount
        zone = WorksheetFunction.Floor((wptDetails(1, count) + 180) / 6, 1) + 1
        latlon = UTMtoLatLon(shifted(0, count), shifted(1, count), zone)
        newLatLon(0, count) = latlon(0)
        newLatLon(1, count) = latlon(1)
        count = count + 1
    Loop
    
    ' Update the route description
    text = Replace(text, Left(fileName, Len(fileName) - 4), Left(newName, Len(newName) - 4))
    
    ' Replace the original coordinates with the new ones
    strPosition = InStr(1, text, "<waypoints", vbTextCompare)
    count = 0
    Do While InStr(strPosition, text, "<position", vbTextCompare) > 0
        strPosition = InStr(strPosition, text, "<position", vbTextCompare)
        text = Left(text, strPosition - 1) & Replace(text, "lat=""" & CStr(wptDetails(0, count)), "lat=""" & CStr(newLatLon(0, count)), strPosition, 1, vbTextCompare)
        text = Left(text, strPosition - 1) & Replace(text, "lon=""" & CStr(wptDetails(1, count)), "lon=""" & CStr(newLatLon(1, count)), strPosition, 1, vbTextCompare)
        count = count + 1
        strPosition = strPosition + 5
    Loop
    
    ' Output the new file
    Open path & newName For Output As #1
        Print #1, text
    Close #1
    
EndEarly:
    
End Function

Function ExtractNumbers(str)
    Dim regex As RegExp, matches As MatchCollection, mtch As Match, result As Double
    
    Set regex = New RegExp
    regex.Pattern = "-?\d+(\.\d+)?"
    regex.Global = False
    
    If regex.test(str) = True Then
        Set matches = regex.Execute(str)
        For Each mtch In matches
            result = mtch
        Next
    End If
    
    ExtractNumbers = result
    
End Function

Function ReadRawFiles()
    Dim fso As Object, dt As Object, folder As Object, file As Object, fileNames() As String, folderPath As String, outputPath As String
    Dim count As Integer, hoursBack As Double, utmZone As Variant, centMer As Integer
    Dim text As String, textLine As String, textFiles() As String, f As textStream
    
    utmZone = False

    ' Load configuration data
    folderPath = CStr(Range("F2").Value)
    outputPath = CStr(Range("F3").Value)
    hoursBack = CDbl(Range("F4").Value)
    ' Read optional extension distances (meters) for start/end of track
    Dim extendStartMeters As Double, extendEndMeters As Double, combine As Double, lazDiff As Double, backup As String, fileExt As String
    extendStartMeters = 0
    extendEndMeters = 0
    If IsNumeric(Range("C10").Value) Then extendStartMeters = CDbl(Range("C10").Value) ' Reverse input order (C9/C10) to account for reversed route
    If IsNumeric(Range("C9").Value) Then extendEndMeters = CDbl(Range("C9").Value)
    combine = CDbl(Range("C11").Value)
    lazDiff = 100
    backup = CStr(Range("C12").Value)
    
    If backup = "Yes" Then
        folderPath = Environ("USERPROFILE") & "\Desktop\Sperry\Hypack-Sperry-Tools\Autolines\"
        fileExt = ".lnw"
    Else
        fileExt = ".raw"
        ' Check formatting of the path
        If Right(folderPath, 1) <> "\" Then
            folderPath = folderPath & "\"
        End If
        If Right(folderPath, 4) <> "Raw\" Then
            folderPath = folderPath & "Raw\"
        End If
    End If
    If Right(outputPath, 1) <> "\" Then
        outputPath = outputPath & "\"
    End If

    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get the folder object
    Set folder = fso.GetFolder(folderPath)

    count = 0
    ' Loop through each file in the folder
    For Each file In folder.Files
        ' Perform actions with each file, e.g., print its name
        ' Match any file with a .raw extension (case-insensitive). This avoids skipping
        ' filenames that contain spaces or other characters that the previous Like
        ' pattern didn't capture.
        If LCase(Right(file.name, 4)) = fileExt Then
            If isBetween(file.path, hoursBack) Then
                count = count + 1
                ReDim Preserve fileNames(1 To count)
                fileNames(count) = file.name
            End If
        End If
    Next file
    
    If count = 0 Then
        MsgBox ("No routes found in project folder.")
        GoTo EndReadEarly
    End If
    
    ReDim textFiles(1 To UBound(fileNames))
    Dim keepSearching As Boolean, startWriting As Boolean, parts As Variant, fileTime As Date, timeStr As String
    Dim i As Integer
    For i = LBound(fileNames) To UBound(fileNames)
        
        ' Open each matching file and read their contents
        count = 0
        text = ""
        keepSearching = True
        startWriting = False
        utmZone = False
        Set f = fso.OpenTextFile(folderPath & fileNames(i), ForReading)
            Do While keepSearching
                textLine = f.ReadLine
                If textLine Like "INI ZoneName=*" Then
                    utmZone = CInt(ExtractNumbers(Mid(textLine, Len("INI ZoneName="))))
                End If
                If textLine Like "ZON*" Then
                    utmZone = CInt(ExtractNumbers(Mid(textLine, Len("ZON "))))
                End If
                If textLine Like "INI CentralMeridian=*" Then
                    centMer = CInt(ExtractNumbers(Mid(textLine, Len("INI CentralMeridian="), 4)))
                End If
                If textLine Like "TND*" Then
                    parts = Split(textLine)
                    If UBound(parts) >= 2 Then
                        fileTime = CDate(parts(2) & " " & parts(1))
                        timeStr = Format(fileTime, "HH:nn mm-dd")
                        timeStr = Replace(timeStr, ":", ".")
                    End If
                End If
                If textLine Like "LIN #*" Then startWriting = True
                If startWriting Then
                    If textLine Like "LBP*" Then
                        ' Do Nothing
                    ElseIf textLine Like "LNN*" Then
                        text = text & textLine & " " & timeStr & vbCrLf
                    Else
                        text = text & textLine & vbCrLf
                        count = count + 1
                    End If
                End If
                If textLine Like "EOL*" Then keepSearching = False
            Loop

            ' Read the last line of the file and output it to a variable.
            Dim lastLine As String
            lastLine = GetLastLineFast(folderPath & fileNames(i))

            ' If the file ends with FIX, reverse only PTS lines in the captured text.
            If True Then ' Disabled reversing of PTS lines for crashed files
            ' If UCase$(Left$(Trim$(lastLine), 3)) <> "FIX" Then
                Dim textLines() As String, ptsLines() As String
                Dim lineIdx As Long, ptsCount As Long, replaceIdx As Long

                textLines = Split(text, vbCrLf)
                ptsCount = 0

                For lineIdx = LBound(textLines) To UBound(textLines)
                    If UCase$(Left$(Trim$(textLines(lineIdx)), 3)) = "PTS" Then
                        ptsCount = ptsCount + 1
                        ReDim Preserve ptsLines(1 To ptsCount)
                        ptsLines(ptsCount) = textLines(lineIdx)
                    End If
                Next lineIdx

                If ptsCount > 1 Then
                    replaceIdx = ptsCount
                    For lineIdx = LBound(textLines) To UBound(textLines)
                        If UCase$(Left$(Trim$(textLines(lineIdx)), 3)) = "PTS" Then
                            textLines(lineIdx) = ptsLines(replaceIdx)
                            replaceIdx = replaceIdx - 1
                        End If
                    Next lineIdx
                    text = Join(textLines, vbCrLf)
                End If
            End If

            f.Close

            Dim pt1Str As String, ptNStr As String

            ' Extend first leg
            If (extendStartMeters > 0) Then
                ' Extract first two PTS points
                Dim pt1 As Point, pt2 As Point, strPos1 As Integer, strPos2 As Integer
                
                strPos1 = InStr(1, text, "PTS ", vbTextCompare)
                strPos2 = InStr(strPos1 + 4, text, " ", vbTextCompare)
                pt1.X = Mid(text, strPos1 + 4, strPos2 - (strPos1 + 4))
                
                strPos1 = strPos2
                strPos2 = InStr(strPos1 + 1, text, vbCrLf, vbTextCompare)
                pt1.Y = Mid(text, strPos1 + 1, strPos2 - (strPos1 + 1))
                
                strPos1 = InStr(strPos2, text, "PTS ", vbTextCompare)
                strPos2 = InStr(strPos1 + 4, text, " ", vbTextCompare)
                pt2.X = Mid(text, strPos1 + 4, strPos2 - (strPos1 + 4))
                
                strPos1 = strPos2
                strPos2 = InStr(strPos1 + 4, text, vbCrLf, vbTextCompare)
                pt2.Y = Mid(text, strPos1 + 1, strPos2 - (strPos1 + 1))

                Dim laz As Double, dx As Double, dy As Double
                
                laz = Sqr((pt2.X - pt1.X) ^ 2 + (pt2.Y - pt1.Y) ^ 2)
                dx = (pt1.X - pt2.X) / laz
                dy = (pt1.Y - pt2.Y) / laz
                
                Debug.Print ("Start ext length: " & (Sqr((dx * extendStartMeters) ^ 2 + (dy * extendStartMeters) ^ 2)))

                pt1Str = "PTS " & Format(dx * extendStartMeters + pt1.X, "0.00") & " " & Format(dy * extendStartMeters + pt1.Y, "0.00")
            End If

            If (extendEndMeters > 0) Then
                ' Extract last two PTS points
                Dim ptN1 As Point, ptN2 As Point, strPosN1 As Integer, strPosN2 As Integer, lastPtPos As Integer
                
                strPosN1 = InStrRev(text, "PTS ", -1, vbTextCompare) ' InStrRev(text, "PTS ", -1, vbTextCompare)
                lastPtPos = strPosN1
                strPosN2 = InStr(strPosN1 + 4, text, " ", vbTextCompare)
                ptN1.X = Mid(text, strPosN1 + 4, strPosN2 - (strPosN1 + 4))
                
                strPosN1 = strPosN2
                strPosN2 = InStr(strPosN1 + 1, text, vbCrLf, vbTextCompare)
                ptN1.Y = Mid(text, strPosN1 + 1, strPosN2 - (strPosN1 + 1))
                
                strPosN1 = InStrRev(text, "PTS ", lastPtPos - 1, vbTextCompare)
                strPosN2 = InStr(strPosN1 + 4, text, " ", vbTextCompare)
                ptN2.X = Mid(text, strPosN1 + 4, strPosN2 - (strPosN1 + 4))
                
                strPosN1 = strPosN2
                strPosN2 = InStr(strPosN1 + 1, text, vbCrLf, vbTextCompare)
                ptN2.Y = Mid(text, strPosN1 + 1, strPosN2 - (strPosN1 + 1))

                Dim lazN As Double, dxN As Double, dyN As Double
                
                lazN = Sqr((ptN1.X - ptN2.X) ^ 2 + (ptN1.Y - ptN2.Y) ^ 2)
                dxN = (ptN1.X - ptN2.X) / lazN
                dyN = (ptN1.Y - ptN2.Y) / lazN
                
                Debug.Print ("End ext length: " & (Sqr((dxN * extendEndMeters) ^ 2 + (dyN * extendEndMeters) ^ 2)))

                ptNStr = "PTS " & Format(dxN * extendEndMeters + ptN1.X, "0.00") & " " & Format(dyN * extendEndMeters + ptN1.Y, "0.00")
            End If

            If extendStartMeters > 0 Or extendEndMeters > 0 Then
                Dim linesArr() As String, ind As Integer, modText As String, ptsFound As Boolean
                linesArr = Split(text, vbCrLf)
                modText = ""
                ptsFound = True

                For ind = LBound(linesArr) To UBound(linesArr)
                    If ptsFound And UCase$(Left$(Trim(linesArr(ind)), 3)) = "PTS" Then
                        ptsFound = False
                        If extendStartMeters > 0 Then
                            modText = modText & pt1Str & vbCrLf
                        End If
                        modText = modText & linesArr(ind) & vbCrLf
                    ElseIf UCase$(Left$(Trim(linesArr(ind)), 3)) = "LNN" Then
                        If extendEndMeters > 0 Then
                            modText = modText & ptNStr & vbCrLf
                        End If
                        modText = modText & linesArr(ind) & vbCrLf
                    Else
                        modText = modText & linesArr(ind) & vbCrLf
                    End If
                Next ind
                text = modText
            End If

            textFiles(i) = text
    Next i
    
    If Range("C8").Value = "Yes" Then
        DeleteFile (outputPath & "*.rtz")
        ' DeleteFile (outputPath & "*.lnw")
    End If
    
    Set f = fso.CreateTextFile(outputPath & "autolines.lnw", True)
    If utmZone = False Then
        utmZone = CInt(WorksheetFunction.Floor((centMer + 180) / 6, 1) + 1)
    End If
    f.WriteLine "ZON " & utmZone
    f.WriteLine ("LNS " & UBound(textFiles))
    Dim strPos As Integer
    For i = LBound(textFiles) To UBound(textFiles)
        f.Write (textFiles(i))
    Next i
    
    outputPath = outputPath & "autolines.lnw"
    readLnw outputPath
    
EndReadEarly:
Debug.Print "Execution finished"

    ' Clean up objects
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Function

Private Function GetLastLineFast(filePath As String) As String
    Dim FileNum As Integer
    Dim fileLen As Long, pos As Long, endPos As Long, lineStart As Long
    Dim lineLen As Long
    Dim ch As String * 1
    Dim result As String

    On Error GoTo Cleanup

    FileNum = FreeFile
    Open filePath For Binary Access Read As #FileNum

    fileLen = LOF(FileNum)
    If fileLen <= 0 Then GoTo Cleanup

    pos = fileLen

    ' Skip trailing CR/LF so we get the last content line even if file ends with a newline.
    Do While pos > 0
        Get #FileNum, pos, ch
        If ch = vbCr Or ch = vbLf Then
            pos = pos - 1
        Else
            Exit Do
        End If
    Loop

    If pos <= 0 Then GoTo Cleanup

    endPos = pos

    ' Find the previous line break scanning backward.
    Do While pos > 0
        Get #FileNum, pos, ch
        If ch = vbCr Or ch = vbLf Then Exit Do
        pos = pos - 1
    Loop

    lineStart = pos + 1
    lineLen = endPos - lineStart + 1

    If lineLen > 0 Then
        result = Space$(lineLen)
        Get #FileNum, lineStart, result
    End If

Cleanup:
    On Error Resume Next
    If FileNum > 0 Then Close #FileNum
    GetLastLineFast = result
End Function

Function ProcessRouteMerging(xte As Double, lazDiff As Double)
    Dim allRoutes As New Collection
    Dim i As Long
    
    If allRoutes.count < 2 Then Exit Function
    
    Dim currentRoute() As PointUTM, nextRoute() As PointUTM
    ' currentRoute = allRoutes(i)
    
    ' For i = 2 To allRoutes.count
        ' nextRoute = allRoutes(i)
        
        ' If CanCombine(
End Function

Function isBetween(filePath As String, hoursBack As Double) As Boolean
    Dim fso As Object, ts As Object, line As String
    Dim parts As Variant, fileTime As Date, timeStr As String
    Dim dt As Object, dtFile As Object
    Dim utcNow As Date, fileTimeUtc As Date, threshold As Date
    Dim fname As String

    On Error GoTo Cleanup

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then GoTo Cleanup

    fname = fso.GetFileName(filePath)

    Set ts = fso.OpenTextFile(filePath, 1)
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) >= 3 And UCase(Left(line, 3)) = "TND" Then
            parts = Split(line)
            If UBound(parts) >= 2 Then
                ' Expect parts: TND time date ... e.g. TND 04:29:58 03/26/2026 0
                On Error GoTo Cleanup
                fileTime = CDate(parts(2) & " " & parts(1)) ' parsed by CDate (local interpretation)
                timeStr = Format(fileTime, "HH:nn yyyy-mm-dd")

                ' Get current UTC time
                Set dt = CreateObject("WbemScripting.SWbemDateTime")
                dt.SetVarDate Now
                utcNow = dt.GetVarDate(False)

                ' Convert the parsed file time (which CDate interpreted as local) into UTC
                Set dtFile = CreateObject("WbemScripting.SWbemDateTime")
                dtFile.SetVarDate fileTime
                fileTimeUtc = dtFile.GetVarDate(False)

                ' Compare file time (UTC) to threshold (utcNow - hoursBack)
                threshold = DateAdd("h", -hoursBack, utcNow)
                If fileTimeUtc >= threshold Then
                    isBetween = True
                    Debug.Print fname & ": match"
                Else
                    isBetween = False
                    Debug.Print fname & ": no match"
                End If

                ts.Close
                Set fso = Nothing
                Exit Function
            End If
        End If
    Loop
    ts.Close

Cleanup:
    isBetween = False
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(filePath) Then Debug.Print fso.GetFileName(filePath) & ": no match"
    End If
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Function

Sub DeleteFile(fileName As String)
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    fso.DeleteFile fileName, True
    
    Set fso = Nothing
    
End Sub
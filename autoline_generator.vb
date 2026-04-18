Option Explicit

' Incremental autoline pipeline for very large Hypack files:
' 1) RAW/HSX -> compact intermediate file (append-only incremental updates)
' 2) Intermediate -> LNW output when requested
'
' Reuses existing project helpers where possible (OffsetPolyline, ExtractNumbers).

Private Type AutolinePoint
    E As Double
    N As Double
End Type

Private Type AutolineMonitorState
    SourcePath As String
    LastOffset As Long
    Carry As String
    UTMZone As Long
    LineName As String
End Type

Public NextMonitorRun As Date

Private gMonitorEnabled As Boolean
Private gMonitorSourcePath As String
Private gMonitorIntermediatePath As String
Private gMonitorIncludePTS As Boolean
Private gMonitorIntervalMinutes As Double

Private Const INTERMEDIATE_SIGNATURE As String = "#ALX1"
Private Const STATE_SIGNATURE As String = "ALXSTATE1"
Private Const READ_CHUNK_SIZE As Long = 1024 * 1024

Dim projectFolder As String, outputFolder As String, recentFile As String, recentFileName As String, parts As Variant
    
projectFolder = Sheets("Autoline Extractor").Range("F2").Value & "\Raw\"
projectFolder = Replace(projectFolder, "\\", "\") ' Fix for if user gave folder with \ at the end
outputFolder = Environ("USERPROFILE") & "\Desktop\Sperry\Hypack-Sperry-Tools\Autolines\"
recentFile = GetMostRecentRawFile(projectFolder)
parts = Split(recentFileName, "\")
recentFileName = parts(UBound(parts))
parts = Split(recentFileName, ".")
recentFileName = parts(LBound(parts))

' ---------------- USER INPUT PLACEHOLDERS ----------------
Dim sourceFilePath As String
sourceFilePath = recentFile ' Leave blank to browse for .RAW/.HSX

Dim intermediatePath As String
intermediatePath = outputFolder ' Leave blank for default *_autoline.alx next to source

Dim includePTS As Boolean
includePTS = False ' Save PTS too (fallback geometry if POS unavailable)

Dim resetIntermediate As Boolean
resetIntermediate = False ' True to rebuild intermediate from byte 0
' ----------------------------------------------------------

' Backward-compatible entry point: now updates intermediate only (no LNW output).
Public Sub GenerateAutolineFromHypack()

    If sourceFilePath = "" Then
        sourceFilePath = Application.GetOpenFilename("Hypack Survey Files (*.raw;*.hsx),*.raw;*.hsx,All Files (*.*),*.*")
        If sourceFilePath = "False" Then Exit Sub
    End If

    If intermediatePath = "" Then
        intermediatePath = BuildDefaultIntermediatePath(sourceFilePath)
    End If

    UpdateIntermediateCore sourceFilePath, intermediatePath, includePTS, resetIntermediate, True
End Sub

Public Sub StartHypackIntermediateMonitor()
    ' ---------------- USER INPUT PLACEHOLDERS ----------------
    'Dim sourceFilePath As String
    'sourceFilePath = "" ' Leave blank to browse for .RAW/.HSX

    'Dim intermediatePath As String
    'intermediatePath = "" ' Leave blank for default *_autoline.alx

    Dim intervalMinutes As Double
    intervalMinutes = 1# ' Poll every 1-2 minutes

    'Dim includePTS As Boolean
    'includePTS = True
    ' ----------------------------------------------------------

    If sourceFilePath = "" Then
        sourceFilePath = Application.GetOpenFilename("Hypack Survey Files (*.raw;*.hsx),*.raw;*.hsx,All Files (*.*),*.*")
        If sourceFilePath = "False" Then Exit Sub
    End If

    If intermediatePath = "" Then
        intermediatePath = BuildDefaultIntermediatePath(sourceFilePath)
    End If

    If intervalMinutes <= 0 Then intervalMinutes = 1#

    gMonitorSourcePath = sourceFilePath
    gMonitorIntermediatePath = intermediatePath
    gMonitorIncludePTS = includePTS
    gMonitorIntervalMinutes = intervalMinutes
    gMonitorEnabled = True

    ' Run immediately, then schedule periodic updates.
    RunHypackMonitorTick

    MsgBox "Hypack monitor started." & vbCrLf & _
           "Source: " & gMonitorSourcePath & vbCrLf & _
           "Intermediate: " & gMonitorIntermediatePath & vbCrLf & _
           "Interval (min): " & Format(gMonitorIntervalMinutes, "0.##"), vbInformation
End Sub

Public Sub RunHypackMonitorTick()
    On Error GoTo ScheduleNext

    If Not gMonitorEnabled Then Exit Sub

    UpdateIntermediateCore gMonitorSourcePath, gMonitorIntermediatePath, gMonitorIncludePTS, False, False

ScheduleNext:
    If gMonitorEnabled Then
        NextMonitorRun = Now + (gMonitorIntervalMinutes / (24# * 60#))
        Application.OnTime NextMonitorRun, "RunHypackMonitorTick"
    End If
End Sub

Public Sub StopHypackIntermediateMonitor()
    On Error Resume Next

    gMonitorEnabled = False

    If NextMonitorRun > 0 Then
        Application.OnTime NextMonitorRun, "RunHypackMonitorTick", , False
    End If

    NextMonitorRun = 0

    MsgBox "Hypack monitor stopped.", vbInformation
End Sub

Public Sub GenerateAutolineLNWFromIntermediate()
    ' ---------------- USER INPUT PLACEHOLDERS ----------------
    'Dim intermediatePath As String
    'intermediatePath = "" ' Leave blank to browse for .alx

    Dim outputLnwPath As String
    outputLnwPath = outputFolder & "\" & recentFileName & "_autoline.lnw" ' Leave blank for default *_autoline.lnw

    Dim overlapPercent As Double
    overlapPercent = 20# ' Target overlap percent (default 20)

    Dim speedKnots As Double
    speedKnots = 8# ' Vessel speed used for turn-radius constraint

    Dim turnDegreesPerMinute As Double
    turnDegreesPerMinute = 100# ' dpm in turn-radius formula

    Dim beamAngleDeg As Double
    beamAngleDeg = 120# ' Placeholder MBES total beam angle

    Dim swathScaleFactor As Double
    swathScaleFactor = 1# ' Effective swath multiplier, e.g. 0.9 conservative

    Dim defaultDepthMeters As Double
    defaultDepthMeters = 30# ' Fallback if no EC1 depth available

    Dim depthPercentileForCoverage As Double
    depthPercentileForCoverage = 20# ' Lower percentile depth for conservative spacing

    Dim offsetSide As String
    offsetSide = "Both" ' "Left" or "Right" relative to traveled direction

    Dim simplifyToleranceMeters As Double
    simplifyToleranceMeters = 0# ' 0 = auto (15% of computed spacing)
    ' ----------------------------------------------------------

    If intermediatePath = "" Then
        intermediatePath = Application.GetOpenFilename("Autoline Intermediate (*.alx),*.alx,All Files (*.*),*.*")
        If intermediatePath = "False" Then Exit Sub
    End If

    If outputLnwPath = "" Then
        outputLnwPath = BuildDefaultOutputFromIntermediate(intermediatePath)
    End If

    Dim pointsPos() As AutolinePoint, pointsPts() As AutolinePoint, depthValues() As Double
    Dim posCount As Long, ptsCount As Long, depthCount As Long
    Dim utmZone As Long, lineName As String

    If Not ParseIntermediateFile(intermediatePath, utmZone, lineName, pointsPos, posCount, pointsPts, ptsCount, depthValues, depthCount) Then
        MsgBox "Unable to parse intermediate file: " & intermediatePath, vbExclamation
        Exit Sub
    End If

    Dim basePoints() As AutolinePoint, baseCount As Long
    If posCount >= 2 Then
        basePoints = CopyPoints(pointsPos, posCount)
        baseCount = posCount
    ElseIf ptsCount >= 2 Then
        basePoints = CopyPoints(pointsPts, ptsCount)
        baseCount = ptsCount
    Else
        MsgBox "Intermediate file needs at least 2 POS or 2 PTS records.", vbExclamation
        Exit Sub
    End If

    Dim depthRef As Double
    depthRef = ResolveDepth(depthValues, depthCount, defaultDepthMeters, depthPercentileForCoverage)

    Dim spacingMeters As Double
    spacingMeters = ComputeSpacingMeters(depthRef, beamAngleDeg, overlapPercent, swathScaleFactor)
    If spacingMeters <= 0 Then
        MsgBox "Computed spacing is invalid. Check overlap/depth/beam settings.", vbExclamation
        Exit Sub
    End If

    If simplifyToleranceMeters <= 0 Then simplifyToleranceMeters = spacingMeters * 0.15

    Dim e() As Double, n() As Double
    ReDim e(0 To baseCount - 1)
    ReDim n(0 To baseCount - 1)

    Dim i As Long
    For i = 0 To baseCount - 1
        e(i) = basePoints(i).E
        n(i) = basePoints(i).N
    Next i

    Dim shifted() As Double
    If offsetSide = "Both" Then
        Dim shiftedLeft() As Double, shiftedRight() As Double
        shiftedLeft = OffsetPolyline(e, n, spacingMeters, "Left")
        shiftedRight = OffsetPolyline(e, n, spacingMeters, "Right")
        ReDim shifted(0 To 1, 0 To baseCount - 1)
        For i = 0 To baseCount - 1
            shifted(0, i) = (shiftedLeft(0, i) + shiftedRight(0, i)) / 2
            shifted(1, i) = (shiftedLeft(1, i) + shiftedRight(1, i)) / 2
        Next i
    Else
        shifted = OffsetPolyline(e, n, spacingMeters, offsetSide)
    End If

    Dim shiftedPts() As AutolinePoint
    ReDim shiftedPts(0 To baseCount - 1)
    For i = 0 To baseCount - 1
        shiftedPts(i).E = shifted(0, i)
        shiftedPts(i).N = shifted(1, i)
    Next i

    Dim simplified() As AutolinePoint
    simplified = SimplifyRDP(shiftedPts, simplifyToleranceMeters)

    Dim minTurnRadiusMeters As Double
    minTurnRadiusMeters = CalcTurnRadiusMeters(speedKnots, turnDegreesPerMinute)

    Dim finalPts() As AutolinePoint
    finalPts = EnforceMinSegmentLength(simplified, minTurnRadiusMeters)

    If UBound(finalPts) < 1 Then
        MsgBox "Autoline generation produced fewer than two waypoints.", vbExclamation
        Exit Sub
    End If

    If lineName = "" Then lineName = "Autoline Generated"
    If utmZone = 0 Then utmZone = 17

    WriteLNW outputLnwPath, utmZone, lineName, finalPts

    MsgBox "LNW autoline created:" & vbCrLf & outputLnwPath & vbCrLf & _
           "Intermediate: " & intermediatePath & vbCrLf & _
           "Base points: " & CStr(baseCount) & vbCrLf & _
           "Final points: " & CStr(UBound(finalPts) + 1) & vbCrLf & _
           "Depth used (m): " & Format(depthRef, "0.00") & vbCrLf & _
           "Spacing (m): " & Format(spacingMeters, "0.00") & vbCrLf & _
           "Min turn radius (m): " & Format(minTurnRadiusMeters, "0.00"), vbInformation
End Sub

Private Sub UpdateIntermediateCore(ByVal sourceFilePath As String, _
                                   ByVal intermediatePath As String, _
                                   ByVal includePTS As Boolean, _
                                   ByVal resetIntermediate As Boolean, _
                                   ByVal showMessage As Boolean)
    On Error GoTo Fail

    If Len(Dir$(sourceFilePath)) = 0 Then
        MsgBox "Source file not found: " & sourceFilePath, vbExclamation
        Exit Sub
    End If

    Dim statePath As String
    statePath = intermediatePath & ".state"

    If resetIntermediate Then
        SafeDeleteFile intermediatePath
        SafeDeleteFile statePath
    End If

    If Len(Dir$(intermediatePath)) = 0 Then
        CreateIntermediateHeader intermediatePath, sourceFilePath
    End If

    Dim st As AutolineMonitorState
    If Not LoadMonitorState(statePath, st) Then
        st.SourcePath = sourceFilePath
        st.LastOffset = 0
        st.Carry = ""
        st.UTMZone = 0
        st.LineName = ""
    End If

    If LCase$(st.SourcePath) <> LCase$(sourceFilePath) Then
        st.SourcePath = sourceFilePath
        st.LastOffset = 0
        st.Carry = ""
        st.UTMZone = 0
        st.LineName = ""
    End If

    Dim recordsAppended As Long
    AppendCriticalDataFromSource sourceFilePath, intermediatePath, includePTS, st, recordsAppended

    SaveMonitorState statePath, st

    If showMessage Then
        MsgBox "Intermediate update complete." & vbCrLf & _
               "Source: " & sourceFilePath & vbCrLf & _
               "Intermediate: " & intermediatePath & vbCrLf & _
               "Records appended: " & CStr(recordsAppended) & vbCrLf & _
               "Last byte offset: " & CStr(st.LastOffset), vbInformation
    End If

    Exit Sub

Fail:
    MsgBox "Intermediate update failed: " & Err.Description, vbExclamation
End Sub

Private Sub AppendCriticalDataFromSource(ByVal sourcePath As String, _
                                         ByVal intermediatePath As String, _
                                         ByVal includePTS As Boolean, _
                                         ByRef st As AutolineMonitorState, _
                                         ByRef recordsAppended As Long)
    Dim srcFF As Integer, outFF As Integer
    Dim fileLen As Long

    recordsAppended = 0

    srcFF = FreeFile
    Open sourcePath For Binary Access Read Shared As #srcFF

    fileLen = LOF(srcFF)

    If fileLen < st.LastOffset Then
        st.LastOffset = 0
        st.Carry = ""
    End If

    If fileLen <= st.LastOffset Then
        Close #srcFF
        Exit Sub
    End If

    outFF = FreeFile
    Open intermediatePath For Append As #outFF

    Dim readPos As Long
    readPos = st.LastOffset + 1

    Do While readPos <= fileLen
        Dim readLen As Long
        readLen = READ_CHUNK_SIZE
        If readPos + readLen - 1 > fileLen Then readLen = fileLen - readPos + 1

        Dim bytesArr() As Byte
        ReDim bytesArr(1 To readLen)
        Get #srcFF, readPos, bytesArr

        Dim chunkText As String
        chunkText = StrConv(bytesArr, vbUnicode)

        ProcessChunkText chunkText, st, outFF, includePTS, recordsAppended

        readPos = readPos + readLen
    Loop

    st.LastOffset = fileLen

    Close #outFF
    Close #srcFF
End Sub

Private Sub ProcessChunkText(ByVal chunkText As String, _
                             ByRef st As AutolineMonitorState, _
                             ByVal outFF As Integer, _
                             ByVal includePTS As Boolean, _
                             ByRef recordsAppended As Long)
    Dim buffer As String
    buffer = st.Carry & chunkText

    buffer = Replace$(buffer, vbCrLf, vbLf)
    buffer = Replace$(buffer, vbCr, vbLf)

    Dim lines() As String
    lines = Split(buffer, vbLf)

    Dim hasTail As Boolean
    hasTail = (Right$(buffer, 1) <> vbLf)

    Dim lastComplete As Long
    If hasTail Then
        st.Carry = lines(UBound(lines))
        lastComplete = UBound(lines) - 1
    Else
        st.Carry = ""
        lastComplete = UBound(lines)
    End If

    Dim i As Long
    For i = LBound(lines) To lastComplete
        ParseCriticalLineAndWrite Trim$(lines(i)), st, outFF, includePTS, recordsAppended
    Next i
End Sub

Private Sub ParseCriticalLineAndWrite(ByVal t As String, _
                                      ByRef st As AutolineMonitorState, _
                                      ByVal outFF As Integer, _
                                      ByVal includePTS As Boolean, _
                                      ByRef recordsAppended As Long)
    If Len(t) = 0 Then Exit Sub

    If Left$(t, 11) = "INI ZoneId=" Then
        st.UTMZone = ZoneFromZoneId(CLng(ExtractNumbers(Mid$(t, 12))))
        If st.UTMZone > 0 Then
            Print #outFF, "ZONE," & CStr(st.UTMZone)
            recordsAppended = recordsAppended + 1
        End If

    ElseIf Left$(t, 13) = "INI ZoneName=" Then
        If st.UTMZone = 0 Then
            st.UTMZone = CLng(ExtractNumbers(t))
            If st.UTMZone > 0 Then
                Print #outFF, "ZONE," & CStr(st.UTMZone)
                recordsAppended = recordsAppended + 1
            End If
        End If

    ElseIf Left$(t, 4) = "LNN " Then
        st.LineName = Mid$(t, 5)
        Print #outFF, "LINE," & CsvEscape(st.LineName)
        recordsAppended = recordsAppended + 1

    ElseIf Left$(t, 4) = "TND " Then
        Dim tndParts() As String
        tndParts = SplitCompact(t)
        If UBound(tndParts) >= 2 Then
            Print #outFF, "TND," & CsvEscape(tndParts(1)) & "," & CsvEscape(tndParts(2))
            recordsAppended = recordsAppended + 1
        End If

    ElseIf Left$(t, 4) = "POS " Then
        Dim posParts() As String
        posParts = SplitCompact(t)
        ' POS <dev> <time> <E> <N>
        If UBound(posParts) >= 4 Then
            If IsNumeric(posParts(2)) And IsNumeric(posParts(3)) And IsNumeric(posParts(4)) Then
                Print #outFF, "POS," & InvariantNumberString(CDbl(posParts(2))) & "," & _
                              InvariantNumberString(CDbl(posParts(3))) & "," & _
                              InvariantNumberString(CDbl(posParts(4)))
                recordsAppended = recordsAppended + 1
            End If
        End If

    ElseIf Left$(t, 4) = "EC1 " Then
        Dim ecParts() As String
        ecParts = SplitCompact(t)
        ' EC1 <dev> <time> <depth>
        If UBound(ecParts) >= 3 Then
            If IsNumeric(ecParts(2)) And IsNumeric(ecParts(3)) Then
                Print #outFF, "EC1," & InvariantNumberString(CDbl(ecParts(2))) & "," & _
                              InvariantNumberString(CDbl(ecParts(3)))
                recordsAppended = recordsAppended + 1
            End If
        End If

    ElseIf includePTS And Left$(t, 4) = "PTS " Then
        Dim ptsParts() As String
        ptsParts = SplitCompact(t)
        If UBound(ptsParts) >= 2 Then
            If IsNumeric(ptsParts(1)) And IsNumeric(ptsParts(2)) Then
                Print #outFF, "PTS," & InvariantNumberString(CDbl(ptsParts(1))) & "," & _
                              InvariantNumberString(CDbl(ptsParts(2)))
                recordsAppended = recordsAppended + 1
            End If
        End If
    End If
End Sub

Private Sub CreateIntermediateHeader(ByVal intermediatePath As String, ByVal sourcePath As String)
    Dim ff As Integer
    ff = FreeFile

    Open intermediatePath For Output As #ff
    Print #ff, INTERMEDIATE_SIGNATURE
    Print #ff, "#SOURCE," & CsvEscape(sourcePath)
    Print #ff, "#CREATED," & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #ff
End Sub

Private Function ParseIntermediateFile(ByVal intermediatePath As String, _
                                       ByRef utmZone As Long, _
                                       ByRef lineName As String, _
                                       ByRef pointsPos() As AutolinePoint, _
                                       ByRef posCount As Long, _
                                       ByRef pointsPts() As AutolinePoint, _
                                       ByRef ptsCount As Long, _
                                       ByRef depthValues() As Double, _
                                       ByRef depthCount As Long) As Boolean
    On Error GoTo Fail

    Dim ff As Integer
    ff = FreeFile

    Open intermediatePath For Input As #ff

    Dim lineText As String
    Do Until EOF(ff)
        Line Input #ff, lineText
        lineText = Trim$(lineText)
        If Len(lineText) = 0 Then GoTo NextLine
        If Left$(lineText, 1) = "#" Then GoTo NextLine

        Dim p() As String
        p = Split(lineText, ",")
        If UBound(p) < 1 Then GoTo NextLine

        Select Case UCase$(p(0))
            Case "ZONE"
                If IsNumeric(p(1)) Then utmZone = CLng(p(1))

            Case "LINE"
                lineName = CsvUnescape(p(1))

            Case "POS"
                If UBound(p) >= 3 Then
                    If IsNumeric(p(2)) And IsNumeric(p(3)) Then
                        AppendPoint pointsPos, posCount, CDbl(p(2)), CDbl(p(3))
                    End If
                End If

            Case "PTS"
                If UBound(p) >= 2 Then
                    If IsNumeric(p(1)) And IsNumeric(p(2)) Then
                        AppendPoint pointsPts, ptsCount, CDbl(p(1)), CDbl(p(2))
                    End If
                End If

            Case "EC1"
                If UBound(p) >= 2 Then
                    If IsNumeric(p(2)) Then AppendDouble depthValues, depthCount, CDbl(p(2))
                End If
        End Select

NextLine:
    Loop

    Close #ff
    ParseIntermediateFile = True
    Exit Function

Fail:
    On Error Resume Next
    If ff > 0 Then Close #ff
    ParseIntermediateFile = False
End Function

Private Function BuildDefaultIntermediatePath(ByVal sourcePath As String) As String
    Dim slashPos As Long
    slashPos = InStrRev(sourcePath, "/")
    If slashPos = 0 Then slashPos = InStrRev(sourcePath, "\")

    If slashPos = 0 Then
        BuildDefaultIntermediatePath = sourcePath & "_autoline.alx"
        Exit Function
    End If

    Dim folderPath As String, baseName As String
    folderPath = Left$(sourcePath, slashPos)
    baseName = Mid$(sourcePath, slashPos + 1)

    Dim dotPos As Long
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)

    BuildDefaultIntermediatePath = folderPath & baseName & "_autoline.alx"
End Function

Private Function BuildDefaultOutputFromIntermediate(ByVal intermediatePath As String) As String
    Dim slashPos As Long
    slashPos = InStrRev(intermediatePath, "/")
    If slashPos = 0 Then slashPos = InStrRev(intermediatePath, "\")

    If slashPos = 0 Then
        BuildDefaultOutputFromIntermediate = intermediatePath & ".lnw"
        Exit Function
    End If

    Dim folderPath As String, baseName As String
    folderPath = Left$(intermediatePath, slashPos)
    baseName = Mid$(intermediatePath, slashPos + 1)

    Dim dotPos As Long
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)

    BuildDefaultOutputFromIntermediate = folderPath & baseName & ".lnw"
End Function

Private Function ZoneFromZoneId(ByVal zoneId As Long) As Long
    If zoneId >= 32601 And zoneId <= 32660 Then
        ZoneFromZoneId = zoneId - 32600
    ElseIf zoneId >= 32701 And zoneId <= 32760 Then
        ZoneFromZoneId = zoneId - 32700
    Else
        ZoneFromZoneId = zoneId
    End If
End Function

Private Function ComputeSpacingMeters(ByVal depthMeters As Double, _
                                      ByVal beamAngleDeg As Double, _
                                      ByVal overlapPercent As Double, _
                                      ByVal swathScaleFactor As Double) As Double
    If depthMeters <= 0 Then Exit Function
    If beamAngleDeg <= 0 Or beamAngleDeg >= 179 Then Exit Function
    If overlapPercent < 0 Or overlapPercent >= 100 Then Exit Function
    If swathScaleFactor <= 0 Then Exit Function

    Dim halfSwath As Double, fullSwath As Double
    halfSwath = depthMeters * Tan((beamAngleDeg / 2#) * Application.WorksheetFunction.Pi() / 180#)
    fullSwath = 2# * halfSwath * swathScaleFactor

    ComputeSpacingMeters = fullSwath * (1# - overlapPercent / 100#)
End Function

Private Function ResolveDepth(ByRef depthValues() As Double, _
                              ByVal depthCount As Long, _
                              ByVal fallbackDepth As Double, _
                              ByVal percentile As Double) As Double
    If depthCount = 0 Then
        ResolveDepth = fallbackDepth
        Exit Function
    End If

    If percentile < 0 Then percentile = 0
    If percentile > 100 Then percentile = 100

    Dim sorted() As Double
    sorted = CopyDoubles(depthValues, depthCount)
    QuickSortDoubles sorted, LBound(sorted), UBound(sorted)

    Dim idx As Long
    idx = CLng((percentile / 100#) * (depthCount - 1))
    If idx < 0 Then idx = 0
    If idx > depthCount - 1 Then idx = depthCount - 1

    ResolveDepth = sorted(idx)
End Function

Private Function CalcTurnRadiusMeters(ByVal speedKnots As Double, ByVal dpm As Double) As Double
    If dpm <= 0 Then dpm = 100#
    CalcTurnRadiusMeters = ((speedKnots / 60#) * (360# / dpm) / (2# * Application.WorksheetFunction.Pi())) * 1852#
End Function

Private Function SimplifyRDP(ByRef points() As AutolinePoint, ByVal epsilon As Double) As AutolinePoint()
    Dim n As Long
    n = UBound(points) - LBound(points) + 1

    If n <= 2 Then
        SimplifyRDP = points
        Exit Function
    End If

    Dim keep() As Boolean
    ReDim keep(LBound(points) To UBound(points))
    keep(LBound(points)) = True
    keep(UBound(points)) = True

    RDPMark points, LBound(points), UBound(points), epsilon, keep

    Dim outCount As Long
    outCount = 0

    Dim i As Long
    For i = LBound(points) To UBound(points)
        If keep(i) Then outCount = outCount + 1
    Next i

    Dim result() As AutolinePoint
    ReDim result(0 To outCount - 1)

    Dim k As Long
    k = 0
    For i = LBound(points) To UBound(points)
        If keep(i) Then
            result(k) = points(i)
            k = k + 1
        End If
    Next i

    SimplifyRDP = result
End Function

Private Sub RDPMark(ByRef points() As AutolinePoint, _
                    ByVal firstIdx As Long, _
                    ByVal lastIdx As Long, _
                    ByVal epsilon As Double, _
                    ByRef keep() As Boolean)
    If lastIdx <= firstIdx + 1 Then Exit Sub

    Dim i As Long
    Dim maxDist As Double, idx As Long
    maxDist = -1#
    idx = -1

    For i = firstIdx + 1 To lastIdx - 1
        Dim d As Double
        d = PerpDistance(points(i), points(firstIdx), points(lastIdx))
        If d > maxDist Then
            maxDist = d
            idx = i
        End If
    Next i

    If maxDist > epsilon And idx >= 0 Then
        keep(idx) = True
        RDPMark points, firstIdx, idx, epsilon, keep
        RDPMark points, idx, lastIdx, epsilon, keep
    End If
End Sub

Private Function PerpDistance(ByVal p As AutolinePoint, ByVal a As AutolinePoint, ByVal b As AutolinePoint) As Double
    Dim dx As Double, dy As Double
    dx = b.E - a.E
    dy = b.N - a.N

    Dim mag2 As Double
    mag2 = dx * dx + dy * dy

    If mag2 = 0 Then
        PerpDistance = DistanceMeters(p, a)
        Exit Function
    End If

    Dim t As Double
    t = ((p.E - a.E) * dx + (p.N - a.N) * dy) / mag2

    Dim proj As AutolinePoint
    proj.E = a.E + t * dx
    proj.N = a.N + t * dy

    PerpDistance = DistanceMeters(p, proj)
End Function

Private Function EnforceMinSegmentLength(ByRef points() As AutolinePoint, ByVal minSegLength As Double) As AutolinePoint()
    If minSegLength <= 0 Then
        EnforceMinSegmentLength = points
        Exit Function
    End If

    Dim n As Long
    n = UBound(points) - LBound(points) + 1
    If n <= 2 Then
        EnforceMinSegmentLength = points
        Exit Function
    End If

    Dim tmp() As AutolinePoint
    ReDim tmp(0 To n - 1)

    Dim keepCount As Long
    keepCount = 1
    tmp(0) = points(LBound(points))

    Dim i As Long
    For i = LBound(points) + 1 To UBound(points) - 1
        If DistanceMeters(tmp(keepCount - 1), points(i)) >= minSegLength Then
            tmp(keepCount) = points(i)
            keepCount = keepCount + 1
        End If
    Next i

    If DistanceMeters(tmp(keepCount - 1), points(UBound(points))) < minSegLength Then
        tmp(keepCount - 1) = points(UBound(points))
    Else
        tmp(keepCount) = points(UBound(points))
        keepCount = keepCount + 1
    End If

    If keepCount < 2 Then
        ReDim tmp(0 To 1)
        tmp(0) = points(LBound(points))
        tmp(1) = points(UBound(points))
        keepCount = 2
    End If

    Dim result() As AutolinePoint
    ReDim result(0 To keepCount - 1)

    For i = 0 To keepCount - 1
        result(i) = tmp(i)
    Next i

    EnforceMinSegmentLength = result
End Function

Private Function DistanceMeters(ByVal p1 As AutolinePoint, ByVal p2 As AutolinePoint) As Double
    DistanceMeters = Sqr((p1.E - p2.E) ^ 2 + (p1.N - p2.N) ^ 2)
End Function

Private Sub WriteLNW(ByVal outputPath As String, _
                     ByVal utmZone As Long, _
                     ByVal lineName As String, _
                     ByRef points() As AutolinePoint)
    Dim ff As Integer
    ff = FreeFile

    Open outputPath For Output As #ff

    Print #ff, "ZON " & utmZone
    Print #ff, "LNS 1"
    Print #ff, "LIN 2"

    Dim i As Long
    For i = LBound(points) To UBound(points)
        Print #ff, "PTS " & Format(points(i).E, "0.00") & " " & Format(points(i).N, "0.00")
    Next i

    Print #ff, "LBP " & Format(points(UBound(points)).E, "0.00") & " " & Format(points(UBound(points)).N, "0.00")
    Print #ff, "LNN " & lineName
    Print #ff, "EOL"

    Close #ff
End Sub

Private Function LoadMonitorState(ByVal statePath As String, ByRef st As AutolineMonitorState) As Boolean
    On Error GoTo Fail

    If Len(Dir$(statePath)) = 0 Then
        LoadMonitorState = False
        Exit Function
    End If

    Dim ff As Integer
    ff = FreeFile

    Open statePath For Input As #ff

    Dim firstLine As String
    If EOF(ff) Then GoTo Fail
    Line Input #ff, firstLine
    If Trim$(firstLine) <> STATE_SIGNATURE Then GoTo Fail

    Dim lineText As String
    Do Until EOF(ff)
        Line Input #ff, lineText
        Dim eqPos As Long
        eqPos = InStr(1, lineText, "=", vbBinaryCompare)
        If eqPos <= 0 Then GoTo NextLine

        Dim k As String, v As String
        k = Left$(lineText, eqPos - 1)
        v = Mid$(lineText, eqPos + 1)

        Select Case UCase$(k)
            Case "SOURCE"
                st.SourcePath = StateUnescape(v)
            Case "LAST_OFFSET"
                st.LastOffset = CLng(Val(v))
            Case "CARRY"
                st.Carry = StateUnescape(v)
            Case "ZONE"
                st.UTMZone = CLng(Val(v))
            Case "LINE"
                st.LineName = StateUnescape(v)
        End Select

NextLine:
    Loop

    Close #ff
    LoadMonitorState = True
    Exit Function

Fail:
    On Error Resume Next
    If ff > 0 Then Close #ff
    LoadMonitorState = False
End Function

Private Sub SaveMonitorState(ByVal statePath As String, ByRef st As AutolineMonitorState)
    Dim ff As Integer
    ff = FreeFile

    Open statePath For Output As #ff
    Print #ff, STATE_SIGNATURE
    Print #ff, "SOURCE=" & StateEscape(st.SourcePath)
    Print #ff, "LAST_OFFSET=" & CStr(st.LastOffset)
    Print #ff, "CARRY=" & StateEscape(st.Carry)
    Print #ff, "ZONE=" & CStr(st.UTMZone)
    Print #ff, "LINE=" & StateEscape(st.LineName)
    Close #ff
End Sub

Private Sub SafeDeleteFile(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath
End Sub

Private Function InvariantNumberString(ByVal value As Double) As String
    Dim s As String
    s = Format$(value, "0.###############")
    InvariantNumberString = Replace$(s, ",", ".")
End Function

Private Function StateEscape(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, "\", "\\")
    t = Replace$(t, vbCr, "\r")
    t = Replace$(t, vbLf, "\n")
    StateEscape = t
End Function

Private Function StateUnescape(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, "\r", vbCr)
    t = Replace$(t, "\n", vbLf)
    t = Replace$(t, "\\", "\")
    StateUnescape = t
End Function

Private Function CsvEscape(ByVal s As String) As String
    CsvEscape = Replace$(s, ",", " ")
End Function

Private Function CsvUnescape(ByVal s As String) As String
    CsvUnescape = s
End Function

Private Sub AppendPoint(ByRef points() As AutolinePoint, ByRef count As Long, ByVal e As Double, ByVal n As Double)
    If count = 0 Then
        ReDim points(0 To 0)
    Else
        ReDim Preserve points(0 To count)
    End If

    points(count).E = e
    points(count).N = n
    count = count + 1
End Sub

Private Sub AppendDouble(ByRef values() As Double, ByRef count As Long, ByVal value As Double)
    If count = 0 Then
        ReDim values(0 To 0)
    Else
        ReDim Preserve values(0 To count)
    End If

    values(count) = value
    count = count + 1
End Sub

Private Function CopyPoints(ByRef points() As AutolinePoint, ByVal count As Long) As AutolinePoint()
    Dim result() As AutolinePoint
    ReDim result(0 To count - 1)

    Dim i As Long
    For i = 0 To count - 1
        result(i) = points(i)
    Next i

    CopyPoints = result
End Function

Private Function CopyDoubles(ByRef values() As Double, ByVal count As Long) As Double()
    Dim result() As Double
    ReDim result(0 To count - 1)

    Dim i As Long
    For i = 0 To count - 1
        result(i) = values(i)
    Next i

    CopyDoubles = result
End Function

Private Sub QuickSortDoubles(ByRef arr() As Double, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, temp As Double

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortDoubles arr, first, j
    If i < last Then QuickSortDoubles arr, i, last
End Sub

Private Function SplitCompact(ByVal s As String) As String()
    Dim t As String
    t = Trim$(s)

    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop

    SplitCompact = Split(t, " ")
End Function
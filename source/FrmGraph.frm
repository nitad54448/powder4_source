VERSION 5.00
Object = "{60C2F168-424A-101C-AA69-0040052BC4EA}#1.0#0"; "pesgo32.ocx"
Begin VB.Form FrmGraph 
   Appearance      =   0  'Flat
   Caption         =   "Graph"
   ClientHeight    =   3860
   ClientLeft      =   450
   ClientTop       =   510
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3860
   ScaleWidth      =   5970
   Begin PesgoLib.Pesgo Pesgo1 
      Height          =   3610
      Left            =   0
      OleObjectBlob   =   "FrmGraph.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   5890
   End
   Begin VB.Menu mnuGraphFile 
      Caption         =   "File"
      Begin VB.Menu mnuGraphRefresh 
         Caption         =   "Refresh graphic"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGraphGeneralOptions 
         Caption         =   "General Options "
      End
      Begin VB.Menu mnuCustomizeGraph 
         Caption         =   "Customize graph"
      End
      Begin VB.Menu mnuGraphExport 
         Caption         =   "Export "
      End
      Begin VB.Menu mColorSetup 
         Caption         =   "Setup color"
         Visible         =   0   'False
         Begin VB.Menu mColorRaw 
            Caption         =   "Raw data"
         End
         Begin VB.Menu mColorSmoothed 
            Caption         =   "Smoothed data"
         End
         Begin VB.Menu mColorRemSmoothing 
            Caption         =   "Removed by Smoothing"
         End
         Begin VB.Menu mColorBackRemoved 
            Caption         =   "Background removed"
         End
         Begin VB.Menu mColorBackContribution 
            Caption         =   "Background contribution"
         End
         Begin VB.Menu mColorK2strip 
            Caption         =   "Ka2 stripped data"
         End
         Begin VB.Menu mColorK2contribution 
            Caption         =   "Ka2 contribution"
         End
      End
      Begin VB.Menu gl2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGraphClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuGraphData 
      Caption         =   "Data treatment"
      Begin VB.Menu mnuGraphSmooth 
         Caption         =   "Smooth"
         Begin VB.Menu mnuSmoothSavitzkyParabola 
            Caption         =   "Savitzky-Golay (parabola)"
         End
         Begin VB.Menu ln 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSmoothAdjacent 
            Caption         =   "Adjacent averaging"
         End
         Begin VB.Menu mnuGraphSmoothSG 
            Caption         =   "Moving-Window"
         End
      End
      Begin VB.Menu mnuGraphBackground 
         Caption         =   "Remove background"
         Begin VB.Menu mnuBackgroundAutomatic 
            Caption         =   "Automatic mode "
         End
         Begin VB.Menu mnuBackgroundManual 
            Caption         =   "Manual mode "
         End
         Begin VB.Menu mnuForcePositive 
            Caption         =   "Force positive value"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove K alpha 2"
         Begin VB.Menu mnuKa2Dong 
            Caption         =   "Ladell/Dong method"
         End
         Begin VB.Menu mnuGraphK2strip 
            Caption         =   "Rachinger method"
         End
      End
      Begin VB.Menu l12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPeak 
         Caption         =   "Find Peaks"
         Begin VB.Menu mnuPeakSonneveld 
            Caption         =   "Sonneveld-Visser"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuGraphPeaks 
            Caption         =   "Simple search (second derivative)"
         End
      End
      Begin VB.Menu mnuGraphExportPeaks 
         Caption         =   "Export peaks to "
         Enabled         =   0   'False
         Begin VB.Menu mnuExportUnitCell 
            Caption         =   "UnitCell window (this program)"
         End
         Begin VB.Menu g3 
            Caption         =   "-"
         End
         Begin VB.Menu exportDI 
            Caption         =   "Ascii file <2theta values, weight>"
         End
      End
      Begin VB.Menu mnuMainPeaks 
         Caption         =   "Show peaks ..."
         Begin VB.Menu mnuShowPeaks 
            Caption         =   "experimental peaks"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuShowPeakNumbers 
            Caption         =   "add peak number"
         End
         Begin VB.Menu mnL1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSimulatePeaks 
            Caption         =   "simulated peaks (for a given cell)"
         End
      End
      Begin VB.Menu mnuInsertPeak 
         Caption         =   "Add peaks (manually)"
      End
      Begin VB.Menu l15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataApplyIntensityCorrection 
         Caption         =   "Apply Intensity correction factor"
      End
      Begin VB.Menu mnuAddtoIntensity 
         Caption         =   "Add a constant to Intensity "
      End
      Begin VB.Menu mnuApplyZero 
         Caption         =   "Apply zero correction (2 theta)"
      End
      Begin VB.Menu mnuProfilefit 
         Caption         =   "Profile fit"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuProfiledefine 
            Caption         =   "Define profile function"
         End
         Begin VB.Menu mnuGraphProfilerefine 
            Caption         =   "Refine profile"
         End
         Begin VB.Menu mnuProfileLeBail 
            Caption         =   "Extract LeBail intensities"
         End
         Begin VB.Menu mnuGraphExportShelx 
            Caption         =   "Export F2 to a Shelx97 file"
         End
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu mScaleMain 
         Caption         =   "Scale"
         Begin VB.Menu mScale 
            Caption         =   "Auto"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mScale 
            Caption         =   "Linear"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mScale 
            Caption         =   "Log"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuGraphActiveSubset 
      Caption         =   "Active subset:"
      Begin VB.Menu workingSubset 
         Caption         =   "Raw data"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu workingSubset 
         Caption         =   "Smoothed data"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu workingSubset 
         Caption         =   "Background stripped"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu workingSubset 
         Caption         =   "K alpha 2 stripped"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu mnuGraphDetails 
      Caption         =   "Options"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSetPoint 
      Caption         =   "&Set Point: 1<ALT+S>"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSelectDone 
      Caption         =   "&Done <ALT+D>"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FrmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GraphTempData() As Single
Dim NumberOfPeaks As Integer
Dim xgraphData() As Single
Dim ygraphdata() As Single

Private Type peaks
Xposition As Single
Yintensity As Single
DerivIntensity As Single
NegativeArea As Single
End Type


Dim amrawdata As Boolean
Dim amsmoothdata As Boolean
Dim amdiffdata As Boolean
Dim ambackstrip As Boolean
Dim ambackonly As Boolean
Dim amk2strip As Boolean
Dim amk2only As Boolean
Dim ampeaks As Boolean
Dim amderivative As Boolean
Dim donecode As Boolean, setpoints As Boolean
Dim changesinraw As Boolean
Sub AdjustMenus()
'run time adjustments of the menus upon the history of work...

End Sub











Private Sub exportDI_Click()
On Error GoTo errortrap
If NumberOfPeaks < 2 Then Err.Raise 1101, , "No peaks to send to the UnitCell..."
Dim returncode As Boolean, i As Integer, t As String
On Error GoTo errortrap
outfil = FreeFile
Dim intensity As Boolean
t = MsgBox("Do you want to export the peak intensity also ?", vbYesNo, prog_name)
If t = vbYes Then intensity = True
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
Open outputfile For Output As outfil
Print #outfil, prog_name & " peaks " & Now
For i = 1 To NumberOfPeaks
t = CStr(Format$((Format$(xgraphData(2, i), "##0.0000")), "@@@@@@@@"))
If intensity Then t = t & CStr(Format$((Format$(ygraphdata(9, i), "#######0.##")), "@@@@@@@@@@@@"))
Print #outfil, t
t = ""
Next i
Close
Exit Sub
errortrap:
Err.Clear

Close
Exit Sub
End Sub

Private Sub Form_Load()
Dim outfil As Integer, result As Integer, i As Integer, test As Long
'setez proprietatile graficului
On Error GoTo errortrap

Pesgo1.FontSizeGlobalCntl = 1
Pesgo1.FontSizeLegendCntl = 0.9

If numarvalori > 32700 Then
MsgBox "The graph can show only 32700 data points; data in memory will be reduced to this amount."
numarvalori = 32700
End If
  
ReDim Preserve GraphTempData(numarvalori)
ReDim xgraphData(3, numarvalori)
ReDim ygraphdata(10, numarvalori)
Pesgo1.Subsets = 1
Pesgo1.points = numarvalori
Pesgo1.PrepareImages = True
amsmoothdata = False: amdiffdata = False: ambackstrip = False: ambackonly = False: amk2strip = False: amk2only = False: ampeaks = False: amderivative = False
amrawdata = True
changesinraw = False
With Pesgo1

.AllowAnnotationControl = False
.AllowArea = False
.AllowBar = False
.AllowBestFitCurve = False
.AllowBestFitLine = True
.AllowBubble = False
.AllowCoordPrompting = True
.AllowCustomization = True
.AllowDataHotSpots = False
.AllowDataLabels = NoDataLabels
.AllowDebugOutput = False
.AllowExporting = True
.AllowGraphAnnotHotSpots = True
.AllowGraphHotSpots = False
.AllowHorzLineAnnotHotSpots = False
.AllowLine = True
.AllowMaximization = True
.AllowOleExport = False
.AllowPlotCustomization = True
.AllowPoint = True
.AllowPointsPlusLine = True
.AllowPointsPlusSpline = False
.AllowPopup = True
.AllowSpline = True
.AllowStick = False
.AllowSubsetHotSpots = False
.AllowUserInterface = True
.AllowVertLineAnnotHotSpots = False
.AllowXAxisAnnotHotSpots = True
.AllowYAxisAnnotHotSpots = True
.AllowZooming = HorzPlusVertZooming
.AnnotationsInFront = True
.ArrowCursor = 0
.AutoScaleData = True
'
.CursorMode = NoCursor
.CursorPageAmount = 50
.CursorPromptStyle = XandYvalues
.CursorPromptTracking = True
.CurveGranularity = FineLines
.DataPrecision = ThreeDecimals
.DataShadows = False
.GridLineControl = NoGrid
.MarkDataPoints = False
.MouseCursorControl = True
''.MultiAxesSubsets(8) = 1
''.MultiAxesSubsets(9) = 1

.NegativeFromXAxis = False
.NoRandomPointsToExport = True
.NoScrollingSubsetControl = False
.NullDataValueY = -99999
'randomsubsets to graph
.PlottingMethod = 0 ' line
.RYAxisLongTicks = False
''.RYAxisScaleControl = Linear

.PrepareImages = True
.RandomSubsetsToGraph(0) = 0
.ScaleForYData = 0
.ScrollingHorzZoom = True
.ShadowColor = RGB(255, 255, 255)

.ShowAnnotations = True
.ShowGraphAnnotations = True
.ShowRYAxis = Empty
.SubsetByPoint = True
.XAxisLongTicks = False
.XAxisScaleControl = Linear
.YAxisScaleControl = Linear
.YAxisLongTicks = False
.XAxisLabel = "2 theta /deg."
.YAxisLabel = "counts /a.u."
''.Subsets = 9
.SubsetLabels(0) = "raw data"
.SubsetLabels(1) = "smooth data"
.SubsetLabels(2) = "raw-smooth"
.SubsetLabels(3) = "backgr. stripped"
.SubsetLabels(4) = "background"
.SubsetLabels(5) = "Ka2 stripped"
.SubsetLabels(6) = "Ka2 contribution"
.SubsetLabels(7) = "peaks"
.SubsetLabels(8) = "derivative"
'colors and linestyles set here, the user may change them if he wish to do so
For i = 1 To 8
Pesgo1.SubsetLineTypes(i) = 0
Next i
.SubTitle = ""
End With

        Pesgo1.SubsetColors(0) = QBColor(2)
        Pesgo1.SubsetColors(1) = QBColor(12)
        Pesgo1.SubsetColors(2) = QBColor(7)
        Pesgo1.SubsetColors(3) = QBColor(1)
        Pesgo1.SubsetColors(4) = QBColor(0)
        Pesgo1.SubsetColors(5) = QBColor(9)
        Pesgo1.SubsetColors(6) = QBColor(5)


FrmGraph.Show

'pun doar subsetul de raw data..
DoEvents

''For j = 1 To 9
For i = 1 To numarvalori
GraphTempData(i - 1) = CSng(Y(i))
''use always OPtionBase 1
ygraphdata(1, i) = CSng(Y(i))
Next i
''Next j
test = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori)
ReDim Preserve GraphTempData(numarvalori)
For i = 1 To numarvalori
GraphTempData(i - 1) = CSng(X(i))
xgraphData(1, i) = CSng(X(i))
Next i
test = PEvset(Pesgo1, PEP_faXDATA, GraphTempData(0), numarvalori)
''For i = 1 To numarvalori
''Pesgo1.XData(0, i - 1) = x(i)
''Pesgo1.YData(0, i - 1) = y(i)

''Next i

Pesgo1.SubsetsToLegend(0) = 0
FrmGraph.Enabled = True
Pesgo1.PEactions = 0
Convert3Main.Visible = False
Exit Sub
errortrap:
Err.Clear
raport "An error has occured."
End Sub

Private Sub Form_Resize()
On Error GoTo errtrap
If Me.Width < 3000 Then Me.Width = 3000
If Me.Height < 3000 Then Me.Height = 3000
'Pesgo1.Width = Me.Width - 50
'Pesgo1.Height = Me.Height - 50
''mnuGraphRefresh_Click
Pesgo1.Width = Me.Width - 240 ''- 1200
Pesgo1.Height = Me.Height - 750
Pesgo1.top = (Me.Height - Pesgo1.Height) / 10
Pesgo1.left = (Me.Width - Pesgo1.Width) / 3
errtrap:
Exit Sub
End Sub

Sub Form_Unload(Cancel As Integer)
 On Error GoTo errortrap
Dim i As Integer, t As Integer, j As Integer, s As Single
For i = 1 To 3
If workingSubset(i).Checked Then Exit For
Next i
If i > 0 And i < 4 Then
t = MsgBox("Replace raw data with: " & workingSubset(i).Caption & " ?", vbYesNo, prog_name)
If t = vbNo Then Err.Raise 1102, , ""
'change the values
    'smooth is 2, back is 4, ka2 strip must be 6 ---check here
    
    For j = 1 To numarvalori
    Y(j) = CDbl(ygraphdata(2 * i, j))
    Next j
   
Else
If changesinraw Then
t = MsgBox("Keep the changes in the raw data ? ", vbYesNo, prog_name)
        If t = vbYes Then
        For s = 1 To numarvalori: Y(s) = CDbl(ygraphdata(1, s)): X(s) = CDbl(xgraphData(1, s)): Next s
        startx = X(1)
        stepx = Val(100000 * X(2) - 100000 * X(1)) / 100000
        'stepx = Val(CInt(1000 * stepx)) / 1000 'correction for 0.0199999 instead of 0.02
        endx = stepx * numarvalori + startx
        amfullxdata = True
        End If
End If
End If
changesinraw = False
Convert3Main.Visible = True
Unload Me
Exit Sub
errortrap:
Convert3Main.Visible = True
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub



Private Sub mnuAddtoIntensity_Click()
On Error GoTo errtrap
Dim t As Single, i As Single
t = InputBox("Add to Y the value: ", prog_name)
If t < -1500 Or t > 1500 Then Err.Raise 1101, , "Invalid value (accepted: -1500 to 1500). Try again..."
For i = 1 To numarvalori Step 1
ygraphdata(1, i) = ygraphdata(1, i) + t
Next i
changesinraw = True
mnuGraphRefresh_Click
Exit Sub
errtrap:
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub mnuApplyZero_Click()


On Error GoTo errtrap

Dim t As Single, i As Single
t = InputBox("Add to X the value: ", prog_name)
If t < -15 Or t > 15 Then Err.Raise 1101, , "Invalid value (accepted: -15 to 15). Try again..."
For i = 1 To numarvalori Step 1
xgraphData(1, i) = xgraphData(1, i) + t
Next i
changesinraw = True
mnuGraphRefresh_Click

Exit Sub
errtrap:
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub

End Sub

Private Sub mnuBackgroundAutomatic_Click()
Dim nc1 As Single, nc2 As Single, curve As Single, ibroad As Single
Dim isampl As Integer, iter As Integer, dremp() As Single, N As Integer
On Error GoTo errortrap
Dim tt As Integer, t As Single
Static vazut As Boolean
'
For tt = 0 To 3
If workingSubset(tt).Checked = True Then Exit For
Next tt
If tt = 3 Then tt = 6
If tt = 2 Then tt = 3
If tt = 1 Then tt = 2
If tt = 0 Then tt = 1

If Not (vazut) Then
t = MsgBox("This procedure is based on the method of Sonneveld and Visser - TNO Delft." & vbCrLf & " You have to input the sampling interval (usually 20) and a broadening coefficient (usually 0.3).  These values should be changed to fit your data", vbOKCancel, prog_name & " - background")
''If t = vbCancel Then Err.Raise 1102, , ""
vazut = True
End If

t = InputBox("enter the sampling interval <5, 50>", prog_name, 20)
If t < 5 Or t > 50 Then Err.Raise 1102, , ""
isampl = t

t = InputBox("enter the curvature constant <0.1, 3>", prog_name, 0.3)
If t < 0.1 Or t > 3 Then Err.Raise 1102, , ""
curv = t

t = InputBox("enter the number of iterations <1, 100>", prog_name, 25)
If t < 1 Or t > 100 Then Err.Raise 1102, , ""
iter = t


'nc1 is 1, nc2 is numarvalori,...
'curv is a single value depending on the background curvature, 0.3 initially
'ibroad is a single, depending somehow as curv
nc1 = 1
nc2 = numarvalori
''curv = 0.3
ibroad = 1
''isampl = 20
''iter = 30 'number of iterations
ReDim dremp(CLng(numarvalori / isampl) + 2)
N = ibroad
isampl = isampl * N
Dim nd1 As Single, nd2 As Single, ipts1 As Integer, rpts1 As Single, iend As Single
nd1 = nc1 / isampl + 1
nd2 = nc2 / isampl
 ipts1 = 3
 rpts1 = 1 / ipts1
 iend = nd1 - 1
For i = CInt(nd1) To CInt(nd2)
dremp(i) = curv
Next i
Dim max2 As Integer, itel As Integer, func4() As Single
max2 = numarvalori
ReDim func4(numarvalori / isampl + 1)
     iend = max2 - 1
itel = 0
For i = 2 To iend Step isampl
itel = itel + 1
func4(itel) = ygraphdata(tt, i)
Next i
l1as = (m1 - 3) / isampl + 1
L2AS = (m2 - 2) / isampl + 1
MAXAS = itel
INKOF = (ipts1 - 1) / 2
     n1 = L2AS + INKOF + 1
     n2 = MAXAS - INKOF
For j = 1 To iter
For m = CInt(n1) To n2
For i = 1 To INKOF

FM = FM + func4(m + i) + func4(m - i)


Next i

FM = func4(m)
     FM = func4(m) + func4(m + 1) + func4(m - 1)
     FM = FM * rpts1
''''  RPTS1=1./IPTS1  RPTS1=1/3
    If (func4(m) < (FM + dremp(m))) Then GoTo 820
   func4(m) = FM
   ''---------------------///////////
     dremp(m) = 0# ''

820:

Next m
Next j

If N = 1 Then GoTo hereisbackgr
Dim recn As Single
recn = 1 / N
For i = 1 To MAXAS
dremp(i) = func4(i)
Next i

For i = 2 To MAXAS
         k = N * (i - 2) + 1
         func4(k + N) = dremp(i)
         SLOPE = (dremp(i) - dremp(i - 1)) * recn
For j = 1 To N - 1
 func4(k + j) = func4(k) + j * SLOPE
Next j
Next i
    MAXAS = (MAXAS - 1) * N + 1
     l1as = (l1as - 2) * N + 2
     L2AS = L2AS * N
hereisbackgr:
'
      RSAMPL = 1# / (isampl)
      ipts1 = 3
''c
Dim k1 As Single, k2 As Single
      INKOF = (ipts1 - 1) / 2
      k1 = L2AS + INKOF + 1
      k2 = MAXAS - INKOF - 1

For i = 1 To isampl
ygraphdata(5, i) = ygraphdata(tt, i)
ygraphdata(5, numarvalori - i + 1) = ygraphdata(tt, numarvalori - i + 1)
Next i
 For k = 1 To k2 + 1
 ''modified by me, 1 to k2 +1
 delta = (func4(k + 1) - func4(k)) * RSAMPL
      KK = (k - 1) * isampl + 2

 For i = 1 To isampl
 
      FI = i - 1

''      ARLINT = func4(k) + FI * delta
      j = KK + i - 1
       ygraphdata(5, j) = func4(k) + FI * delta


ygraphdata(4, j) = ygraphdata(tt, j) - ygraphdata(5, j)
 
 Next i
 Next k
      
      NREC = (k2 * isampl) / 1000 + 1
      j1 = k2 * isampl + 2
      j2 = NREC * 1000
      
If j2 < Max Then j2 = Max
      l1mb = m1
      l2mb = (k1 - 1) * isampl + 1
      MAXMB = j1 - 1
'transfer here the values for the background ...

ambackstrip = True
ambackonly = True
FrmGraph.workingSubset(2).Enabled = True 'may apply correction to background
mnuGraphRefresh_Click
FrmGraph.workingSubset(2).Enabled = True 'if succesful, may apply corr to back
workingSubset_Click (2)

Exit Sub
errortrap:
''If Err.Number <> 1102 Then MsgBox Err.Description
Err.Clear
Exit Sub

End Sub

Private Sub mnuBackgroundManual_Click()
Dim t As Integer, i As Integer, xb(100) As Double, yb(100) As Double
On Error GoTo errtrap
Dim tt As Integer
'
For tt = 0 To 3
If workingSubset(tt).Checked = True Then Exit For
Next tt
If tt = 3 Then tt = 6
If tt = 2 Then tt = 3
If tt = 1 Then tt = 2
If tt = 0 Then tt = 1
'i apply background correction to the data requested,...in working subset, on the screen should be the same
'this routine determines the background of the signal after the user chooses
'the polynomial degree to be used
'the number of points is let to the user
'if the user chooses less than poldegree+1 points an error will occur with err.raise method
mnuSelectDone.Visible = True
mnuSetPoint.Caption = "Set point: 1" & " <ALT+S>"
mnuSetPoint.Visible = True

'make all the other menus inactive
mnuGraphFile.Enabled = False: mnuGraphData.Enabled = False: mnuGraphActiveSubset.Enabled = False: mnuGraphDetails.Enabled = False
t = InputBox("Input the degree of the interpolation polynomial to be used (1-9)" & vbCrLf & "Later, you have to select the points on the graph (max. 100). At least n+1 points are required, where n is the polynomial degree." & vbCrLf & "Click on <Done> when finished.", prog_name & " - Background", 3)
If t < 1 Or t > 9 Then i = 1 / 0
'we have now to read the coordinates of the points read by the user
'100 points seems appropriate
'go and find out the points, all the other operations are when returning from the "Done menus, ce intoarce un boolean de "done""
'here find out the points
'set here the cursor in pesgo
Pesgo1.CursorMode = InvertedBlock
Pesgo1.MouseCursorControl = True
Pesgo1.AllowDataHotSpots = True
i = 1
Do Until donecode
DoEvents
'let the user play around with the mouse, give some processor time
If setpoints Then
DoEvents
setpoints = False
xb(i) = CDbl(X(Pesgo1.CursorPoint + 1))
yb(i) = CDbl(Y(Pesgo1.CursorPoint + 1))
DoEvents
i = i + 1
mnuSetPoint.Caption = "Set point: " & CStr(i) & " <ALT+S>"

If i = 101 Then donecode = True
DoEvents
End If
DoEvents
Loop
i = i - 1

donecode = False
setpoints = False
mnuSelectDone.Visible = False: mnuSetPoint.Visible = False
mnuGraphFile.Enabled = True: mnuGraphData.Enabled = True: mnuGraphActiveSubset.Enabled = True: mnuGraphDetails.Enabled = True
Pesgo1.CursorMode = NoCursor
Pesgo1.MouseCursorControl = False
Pesgo1.AllowDataHotSpots = False
'the computation of the polynomial starts here
'the polynomial degree is t, I have i points, I have to make a polynomial interpolation
'with the pseudoinverse matrix method
If i < (t + 1) Then
MsgBox "not enough data points...try again"
i = 1 / 0 'raise an error
End If

Dim eroare As Boolean, z() As Double, solution() As Double, ecuatii As Integer, necunoscute As Integer
ReDim z(i, t + 1), ii(i), solution(t + 1)
'I have to buid the matrix z
For ecuatii = 1 To i
For necunoscute = 1 To t + 1
z(ecuatii, necunoscute) = xb(ecuatii) ^ (necunoscute - 1)
Next necunoscute
Next ecuatii
Call pseudoinv(i, t + 1, z(), yb(), solution(), 0.0000000001, eroare)
'generate the background now with the polynomial found
If eroare Then
MsgBox "Error in the Inverse routine..."
i = 1 / 0
End If
Dim j As Single, k As Integer, suma As Single
For j = 1 To numarvalori
suma = solution(1)
For k = 2 To t + 1
suma = suma + solution(k) * (xgraphData(1, j)) ^ (k - 1)
Next k
ygraphdata(5, j) = suma
ygraphdata(4, j) = ygraphdata(tt, j) - ygraphdata(5, j)
Next j
ambackstrip = True
ambackonly = True
FrmGraph.workingSubset(2).Enabled = True 'may apply correction to background
mnuGraphRefresh_Click
workingSubset_Click (2)
Exit Sub
errtrap:
donecode = False: setpoints = False
mnuSelectDone.Visible = False: mnuSetPoint.Visible = False
mnuGraphFile.Enabled = True: mnuGraphData.Enabled = True: mnuGraphActiveSubset.Enabled = True: mnuGraphDetails.Enabled = True

Err.Clear
Exit Sub

End Sub

Private Sub mnuCustomizeGraph_Click()
Pesgo1.PElaunchcustomize
End Sub

Private Sub mnuDataApplyIntensityCorrection_Click()
On Error GoTo errtrap
Dim t As Single, i As Single
t = InputBox("Multiply Y data with the value: ", prog_name)
If t < 0 Or t > 1000 Then Err.Raise 1101, , "Invalid value. Try again..."
For i = 1 To numarvalori Step 1
ygraphdata(1, i) = ygraphdata(1, i) * t
Next i
changesinraw = True
mnuGraphRefresh_Click

Exit Sub
errtrap:
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub mnuDataMarker_Click()

End Sub

Private Sub mnuExportUnitCell_Click()
'I will send the marked peaks in the Unit Cell data grid.
'the former positions there will be erased...
'I need to open from here the UnitCell form...
On Error GoTo errortrap
If NumberOfPeaks < 2 Then Err.Raise 1101, , "No peaks to send to the UnitCell..."
frmRefine.Show
'we delete all the data there
frmRefine.mnuErase_Click
frmRefine.mnuSetDataType_Click (0)
frmRefine.grid.Rows = NumberOfPeaks + 1
For i = 1 To NumberOfPeaks
DoEvents
frmRefine.grid.Row = i
frmRefine.grid.Col = 4
frmRefine.grid.Text = CStr(Format$(xgraphData(2, i), "##0.000##"))
frmRefine.grid.Col = 5
frmRefine.grid.Text = CStr(Format$(ygraphdata(9, i), "#####0.0#"))
Next i
Exit Sub
errortrap:
If Err.Number = 1101 Then MsgBox Err.Description
Exit Sub

End Sub

Private Sub mnuGraphClose_Click()
Form_Unload (-1)
Exit Sub
End Sub

Private Sub mnuGraphDetails_Click()
Dim i As Integer
Load frmOptions
frmOptions.Show
End Sub

Private Sub mnuGraphExport_Click()
Pesgo1.PElaunchexport
End Sub

Private Sub mnuGraphGeneralOptions_Click()
Dim a As Boolean

a = Pesgo1.PElaunchpopupmenu(1, 1)
End Sub

Private Sub mnuGraphK2strip_Click()
''use computeshift2theta
On Error GoTo errortrap
Dim CorrectionDone() As Boolean, RachConst As Integer, nIntervals As Integer, minim As Double, jcount As Integer
''rachconts is a shift constant equal to 2 or 3...
Dim eroare As Boolean, k As Single, j As Single, t As String, wl As Double, wl1 As Double, wl2 As Double, wlratio As Double, xout As Double
Static vazut As Boolean
If Not (vazut) Then
vazut = True
MsgBox "This procedure uses the Rachinger method for Ka2 stripping." & vbCrLf & "It is <normal> if you see meaningless fluctuations in the resulting intensity curve, at the end of the peak. These are due to the errors in the measurement, see Klug and Alexaner, pag. 628. You can later apply brute force, i.e. smoothing."
End If
RachConst = 0
t = InputBox("Input the tube you used (Cu, Cr, Mo, Fe, Co, Ag )", prog_name, "Cu")
    Select Case UCase$(left$(t, 2))
    Case "CU"
        wl1 = 1.54056: wl2 = 1.54439: wl = 1.54178
    Case "CR"
        wl1 = 2.2897: wl2 = 2.2936: wl = 2.29089
    Case "MO"
        wl1 = 0.7093: wl2 = 0.71359: wl = 0.71073
    Case "AG"
        wl1 = 0.5594: wl2 = 0.56379: wl = 0.5608
    Case "CO"
        wl1 = 1.78896: wl2 = 1.79285: wl = 1.7915
    Case "FE"
        wl1 = 1.93604: wl2 = 1.93998: wl = 1.93734
    Case Else
        Err.Raise 1102, , ""
    End Select
nIntervals = 0
t = InputBox("Input the Ka2/Ka1 ratio :", prog_name, "0.49")
wlratio = Val(t)
If wlratio < 0.01 Or wlratio > 0.99 Then Err.Raise 1101, , "Wrong value..."
    t = InputBox("Input a shift constant (0-4):", prog_name, "0")
    RachConst = CInt(t)
If RachConst < 0 Or RachConst > 50 Then Err.Raise 1101, , "Incorrect value..."
    nIntervals = nIntervals + 1
    IamBusy True
        For iii = 0 To 3
            If workingSubset(iii).Checked Then Exit For
        Next iii
            Select Case iii
    Case 0 'raw
        iii = 1
    Case 1 'smoothed
        iii = 2
    Case 2 'background stripped
        iii = 4
    Case 3 'ka2 stripped
        iii = 6
        End Select
Dim tempydata() As Single, tempxdata() As Single, maxval As Single, uncorrected(5000) As Single
ReDim tempydata(numarvalori * nIntervals), tempxdata(numarvalori * nIntervals)
ReDim CorrectionDone(numarvalori * nIntervals)
maxval = 0
DoEvents
stepx = CLng(100000 * xgraphData(1, 2) - 100000 * xgraphData(1, 1)) / 100000
For i = 1 To numarvalori - 1
    tempydata(i) = ygraphdata(iii, i)
''ygraphdata(7, i) = -100
    If tempydata(i) > maxval Then maxval = tempydata(i)
Next i
'the number of corrections is jcount
    For i = 1 + RachConst To numarvalori - RachConst - 1
    Call ComputeShift2Theta(wl1, wl2, CDbl(xgraphData(1, i)), xout, eroare)
'this is ka2 contribution
    If eroare Then Err.Raise 1101, , "Untrappeable error..."
    If xout > xgraphData(1, numarvalori - 1) Then Exit For
''If xout > 90 Then MsgBox "90,,,,,,,"
    k = Fix(((xout - xgraphData(1, i)) / stepx) + 0.5)
If Not (CorrectionDone(i + k)) Then
    If k < 1 Then
tempydata(i) = tempydata(i) * (1 / (1 + wlratio))
Else
''-------------
If (tempydata(i + k) - 0.5 * tempydata(i)) > 0 Then
tempydata(i + k) = tempydata(i + k) - wlratio * tempydata(i - RachConst)
Else
jcount = jcount + 1
uncorrected(jcount) = i + k
End If
End If
CorrectionDone(i + k) = True
End If
Next i
For i = 1 To jcount
tempydata(uncorrected(i)) = (tempydata(uncorrected(i) - 1) + tempydata(uncorrected(i) + 1)) / 2
Next i
For i = 3 To numarvalori - 3
If (Not (CorrectionDone(i)) Or tempydata(i) < 0) Then tempydata(i) = (tempydata(i - 1) + tempydata(i + 1)) / 2
Next i
For i = 1 To numarvalori
ygraphdata(6, i) = tempydata(i)
ygraphdata(7, i) = ygraphdata(iii, i) - ygraphdata(6, i)
Next i
amk2strip = True
FrmGraph.workingSubset(3).Enabled = True 'may apply correction to ka2 stripped
IamBusy False
mnuGraphRefresh_Click
workingSubset_Click (3)
Exit Sub
errortrap:
IamBusy False
If Err.Number = 1101 Then MsgBox Err.Description
Exit Sub
End Sub
























Private Sub mnuGraphPeaks_Click()
'this routine tries to find the peaks automaticaly by defining, minim value of the peak in percent
'minimum width..
'I will use the second derivative over 7 adjacent points (3 points on each side)
''Dim ipoints As Integer
Dim cofnul As Single, cof(3) As Single, i As Single, iii As Integer, Pic(2000) As peaks, MinWidth As Single
On Error GoTo errortrap
Static mesaj As Boolean
ipoints = 3
cofnul = -4
cof(1) = -3
cof(2) = 0
cof(3) = 4
Dim eroare As Boolean, threshold As Single, maxpeaks As Integer
ampeaks = False
NumberOfPeaks = 0
Pesgo1.ShowGraphAnnotations = False
''maxpeaks = 1000 'no more than 1000 peaks are allowed
If Not (mesaj) Then
MsgBox "If you use this function on noisy data or with high curvature background you will get too many peaks.  Smooth and/or remove the background before attempting a peak hunt." & vbCrLf & "This routine uses the second derivative to determine the peak position. You need to input only one parameter: minimum intensity in %" & vbCrLf & "Warning: if the intensity values are very high or very low this routine may not work. If the peaks are not displayed try to adjust the intensity by using <Data/Apply Ontensity Correction Factor> menu. "

'the peaks are stored in the 9th position,''the flag is ampeaks
mesaj = True
End If
threshold = InputBox(" Input the the minimum height for a peak to be considered (given in percent)", prog_name & "-peak hunt", 2)
If threshold < 0.0001 Or threshold > 90 Then Err.Raise 1101, , "The threshold should be higher than 0.01 percent and smaller than 90%"
threshold = 1.41 * threshold
''MinWidth = InputBox(" Input the the minimum distance between peaks (2theta)", prog_name & "-peak hunt", 0.2)
''If MinWidth < 0.01 Or MinWidth > 5 Then Err.Raise 1101, , "The minWidth should be higher than 0.01 and smaller than 5"
MinWidth = 3 * stepx

IamBusy True
DoEvents
'first see to which curve this applies...
'threshold should be input by the user
threshold = threshold / 100
For iii = 0 To 3
If workingSubset(iii).Checked Then Exit For
Next iii
Select Case iii
Case 0
'raw
iii = 1
Case 1
'smoothed
iii = 2
Case 2
'background stripped
iii = 4
Case 3
'ka2 stripped
iii = 6
End Select

'for now i use only 7 points,..2nd degree polynomial
Dim xdata() As Double, ydata() As Double, switch As Integer, localmin As Single, som1 As Single, biggestSum As Single, lastval As Double, minval As Double, drd2 As Single, maxval As Double, ddydata() As Double, suma As Single, minvald2 As Single
ReDim xdata(numarvalori)
ReDim ydata(numarvalori)
ReDim ddydata(numarvalori)

For i = 1 To numarvalori
xdata(i) = xgraphData(1, i)
ydata(i) = ygraphdata(iii, i)
Next i
DoEvents 'give the processor some space
For i = 4 To numarvalori - 4
ddydata(i) = cofnul * ydata(i)
For j = 1 To 3
ddydata(i) = ddydata(i) + cof(j) * (ydata(i + j) + ydata(i - j))
Next j
ddydata(i) = ddydata(i) / 42#
Next i
'minvald2 is the minimum value of the derivative
'I will eliminate all values of the peaks smaller than threshold*minvald2
'find the max of the graph and the minimum value of the derivative
maxval = 0
minvald2 = 0
'smooth the derivative
For i = 4 To numarvalori - 4
suma = 7 * ddydata(i)
suma = suma + 6 * (ddydata(i + 1) + ddydata(i - 1)) + 3 * (ddydata(i + 2) + ddydata(i - 2)) - 2 * (ddydata(i + 3) + ddydata(i - 3))
ddydata(i) = suma / 21
Next i
For i = 1 To numarvalori
If ygraphdata(iii, i) > maxval Then maxval = ygraphdata(iii, i)
If ddydata(i) < minvald2 Then minvald2 = ddydata(i)
Next i
switch = 0 'how many peaks are there
biggestSum = 0
i = 3
drd2 = threshold * minvald2 ''the limit in the second derivative
Dim xd(3) As Double, yd(3) As Double, solution(3) As Double, xmintest As Double, realmind2 As Double, xtest As Double, ytest As Double
Do

    i = i + 1
    localmin = 0
    If i > numarvalori - 4 Then Exit Do
    If ddydata(i) < drd2 Then 'the derivative is smaller than threshold
        'go until the derivative change
        Do
        'determin the localminimum
        If ddydata(i) < localmin Then localmin = i 'the index of x where the derivative is even smaller
        If i > numarvalori - 3 Then Exit Do
        i = i + 1
        'deteremines the negative area in som1
        som1 = som1 - ddydata(i) 'get a positive value for the negative area
        Loop Until ddydata(i) > ddydata(i - 1)
        'here I have passed over the minimum of the derivative
        i = i - 1
        switch = switch + 1
        If som1 > biggestSum Then biggestSum = som1
        If switch > 1998 Then Err.Raise 1101, , "two thousand peaks detected ???...Are you sure you have smoothed and removed the background ?"
       
        ''MsgBox CStr(som1) & " " & CStr(i)
      ''  tempYgraph(switch) = ydata(i) + 0.025 * maxval
      ''  tempXgraph(switch) = xdata(i) ''+ ddydata(i) * (xdata(i + 1) - xdata(i)) / (ddydata(i + 1) - ddydata(i))
        'now go until the derivative change again in slope
'find the real minimum
        xd(1) = xdata(i - 1): yd(1) = ddydata(i - 1)
        xd(2) = xdata(i): yd(2) = ddydata(i)
        xd(3) = xdata(i + 1): yd(3) = ddydata(i + 1)

Call CoefOfaParabola(xd(), yd(), solution(), eroare)
''MsgBox CStr(solution(1))
realmind2 = 0
'I will remove the maxval shift
maxval = 0
For j = 10 To 25
xtest = xdata(i - 1) + j * (xdata(i) - xdata(i - 1)) / 20
Call IntPolValue(2, solution, xtest, ytest, eroare)
If ytest < realmind2 Then realmind2 = ytest: xmintest = xtest
Next j

        Pic(switch).DerivIntensity = ddydata(i)
        Pic(switch).Xposition = xmintest
        Pic(switch).Yintensity = CSng(ydata(i)) + 0.025 * maxval
        Pic(switch).NegativeArea = som1
            
        Do
        If i > numarvalori - 4 Then Exit Do
       
        i = i + 1
              
        Loop Until ddydata(i) < ddydata(i - 1)
    
    
    End If
    
Loop Until i = numarvalori - 4

j = 0
For i = 1 To 1998
If Pic(i).NegativeArea > threshold * biggestSum Then
If Pic(i).Yintensity > threshold * maxval Then
j = j + 1
ygraphdata(9, j) = Pic(i).Yintensity
xgraphData(2, j) = Pic(i).Xposition
End If
End If
Next i
NumberOfPeaks = j ''- 1 last peak is always missing

'erase those veryclose
Do
j = 0
For i = 1 To NumberOfPeaks - 1
If Abs(xgraphData(2, i + 1) - xgraphData(2, i)) < MinWidth Then
j = 1
    If ygraphdata(9, i) > ygraphdata(9, i + 1) Then
    ''erase i+1
    ygraphdata(9, i + 1) = -9999
    xgraphData(2, i + 1) = 1000 * xgraphData(2, i + 1)
    Else
    ''erase i
    ygraphdata(9, i) = -9999
    xgraphData(2, i) = 1000 * xgraphData(2, i)
    End If
    
   

Exit For
End If

Next i
Loop Until j = 0

'now erase all the pics with intensity=-9999
j = 0
For i = 1 To NumberOfPeaks
If ygraphdata(9, i) > 0 Then
j = j + 1

ygraphdata(9, j) = ygraphdata(9, i)
xgraphData(2, j) = xgraphData(2, i)
End If
Next i
NumberOfPeaks = j ''- 1


If NumberOfPeaks < 1 Then Err.Raise 1101, , "No peaks detected under these conditions..."

''----------
ampeaks = True
mnuGraphRefresh_Click

mnuGraphExportPeaks.Enabled = True 'if I found peaks I may export them
IamBusy False
Exit Sub
errortrap:
If Err.Number = 1101 Then MsgBox Err.Description
raport Now & "  - here is the error trap routine in mnuGraphPeaks."
IamBusy False
Exit Sub
Close
End Sub

Sub mnuGraphRefresh_Click()
Dim i As Single, j As Single
''Pesgo1.ShowYAxis = PESA_ALL
On Error GoTo errortrap
''Pesgo1.ScrollingSubsets = 1
IamBusy True
''Pesgo1.PEactions = 0
''Pesgo1.ShowAnnotations = False
'graphic style settings
'Dim amrawdata, amsmoothdata, amdiffdata, ambackstrip, ambackonly, amk2strip, amk2only, ampeaks, amderivative As Boolean
'transfer again all data by pevset, depending upon the data actually present
'if I am here I sure have amrawdata
''set colors here
        Pesgo1.SubsetColors(0) = QBColor(2)
        Pesgo1.SubsetColors(1) = QBColor(12)
        Pesgo1.SubsetColors(2) = QBColor(7)
        Pesgo1.SubsetColors(3) = QBColor(1)
        Pesgo1.SubsetColors(4) = QBColor(0)
        Pesgo1.SubsetColors(5) = QBColor(9)
        Pesgo1.SubsetColors(6) = QBColor(5)
If amsmoothdata Then
    If ambackstrip Then
        If amk2strip Then
        ''SY, BY, KY
        Pesgo1.Subsets = 7 'raw, smooth and difference
        Pesgo1.SubsetLabels(1) = "smoothed data"
        Pesgo1.SubsetLabels(2) = "removed by smoothing"
        Pesgo1.SubsetLabels(3) = "Background stripped data"
        Pesgo1.SubsetLabels(4) = "Background contribution"
        Pesgo1.SubsetLabels(5) = "Ka2 stripped data"
        Pesgo1.SubsetLabels(6) = "Ka2 contribution"

        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        For i = 0 To numarvalori - 1
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = ygraphdata(2, i + 1)
        GraphTempData(2 * numarvalori + i) = ygraphdata(3, i + 1)
        GraphTempData(3 * numarvalori + i) = ygraphdata(4, i + 1)
        GraphTempData(4 * numarvalori + i) = ygraphdata(5, i + 1)
        GraphTempData(5 * numarvalori + i) = ygraphdata(6, i + 1)
        GraphTempData(6 * numarvalori + i) = ygraphdata(7, i + 1)
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        
        Else
        ''SY, BY, KN
        
                
        Pesgo1.Subsets = 5 'raw, smooth and difference
        Pesgo1.SubsetLabels(1) = "smoothed data"
        Pesgo1.SubsetLabels(2) = "removed by smoothing"
        Pesgo1.SubsetLabels(3) = "Background stripped data"
        Pesgo1.SubsetLabels(4) = "Background contribution"
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        For i = 0 To numarvalori - 1
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = ygraphdata(2, i + 1)
        GraphTempData(2 * numarvalori + i) = ygraphdata(3, i + 1)
        GraphTempData(3 * numarvalori + i) = ygraphdata(4, i + 1)
        GraphTempData(4 * numarvalori + i) = ygraphdata(5, i + 1)
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        End If
        Else
    ''i don't have backdtrip, i have smoothdata
        If amk2strip Then
        ''SY BN KY
        
        Pesgo1.Subsets = 5 'raw, smooth and difference
        Pesgo1.SubsetLabels(1) = "smoothed data"
        Pesgo1.SubsetLabels(2) = "removed by smoothing"
        Pesgo1.SubsetLabels(3) = "Ka2 stripped data"
        Pesgo1.SubsetLabels(4) = "Ka2 contribution"
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        For i = 0 To numarvalori - 1
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = ygraphdata(2, i + 1)
        GraphTempData(2 * numarvalori + i) = ygraphdata(3, i + 1)
        GraphTempData(3 * numarvalori + i) = ygraphdata(6, i + 1)
        GraphTempData(4 * numarvalori + i) = ygraphdata(7, i + 1)
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
          
        Else
        ''SY BN KN
        
        Pesgo1.Subsets = 3 'raw, smooth and difference
        Pesgo1.SubsetLabels(1) = "smoothed data"
        Pesgo1.SubsetLabels(2) = "removed by smoothing"
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        For i = 0 To numarvalori - 1
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = ygraphdata(2, i + 1)
        GraphTempData(2 * numarvalori + i) = ygraphdata(3, i + 1)
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        End If
      End If
Else
'I don't have smooth data
    If ambackstrip Then
        If amk2strip Then
        ''SN BY KY
        Pesgo1.Subsets = 5
        DoEvents
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        'when computing backstrip we can have background contribution by the difference
        Pesgo1.SubsetLabels(1) = "Background stripped data"
        Pesgo1.SubsetLabels(2) = "Background contribution"
        Pesgo1.SubsetLabels(3) = "Ka2 stripped data"
        Pesgo1.SubsetLabels(4) = "Ka2 contribution"
        
        For i = 0 To numarvalori - 1 ''add background strip data
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = ygraphdata(4, i + 1)
        GraphTempData(2 * numarvalori + i) = ygraphdata(5, i + 1)
        GraphTempData(3 * numarvalori + i) = ygraphdata(6, i + 1)
        GraphTempData(4 * numarvalori + i) = ygraphdata(7, i + 1)
        Next i

        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        
        
        Else
        'SN BY KN
        Pesgo1.Subsets = 3
        DoEvents
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        'when computing backstrip we can have background contribution by the difference
        Pesgo1.SubsetLabels(1) = "Background stripped data"
        Pesgo1.SubsetLabels(2) = "Background contribution"
        For i = 0 To numarvalori - 1 ''add background strip data
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = CSng(ygraphdata(4, i + 1))
        GraphTempData(2 * numarvalori + i) = CSng(ygraphdata(5, i + 1))
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        End If
     Else
    ''SN, BN ''nothing to do here
        If amk2strip Then ''--------------------------------
        ''KY, BN, SN
        Pesgo1.Subsets = 3
        Pesgo1.SubsetLabels(1) = "Ka2 stripped data"
        Pesgo1.SubsetLabels(2) = "Ka2 contribution"
        DoEvents
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        For i = 0 To numarvalori - 1
        GraphTempData(i) = CSng(ygraphdata(1, i + 1))
        GraphTempData(numarvalori + i) = ygraphdata(6, i + 1)
        GraphTempData(2 * numarvalori + i) = ygraphdata(7, i + 1)
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        Else ''-------------------------------------------------
        'KN, BN, SN ,only raw data here
        Pesgo1.Subsets = 1
        DoEvents
        ReDim GraphTempData(numarvalori * CInt(Pesgo1.Subsets))
        ''put ydata for raw here
        For i = 1 To numarvalori
        GraphTempData(i - 1) = CSng(ygraphdata(1, i))
        Next i
        result = PEvset(Pesgo1, PEP_faYDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."
   
        End If
End If
End If

''transfer here all
ReDim GraphTempData(numarvalori * Pesgo1.Subsets)

For j = 0 To Pesgo1.Subsets - 1
        For i = 1 To numarvalori
        GraphTempData(j * numarvalori + i - 1) = CSng(xgraphData(1, i))
        Next i
Next j
        'send x data by pevset
        result = PEvset(Pesgo1, PEP_faXDATA, GraphTempData(0), numarvalori * Pesgo1.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Xdata array)..."


Pesgo1.GraphAnnotationX(-1) = 1

If mnuShowPeaks.Checked And ampeaks Then
For i = 1 To NumberOfPeaks
Pesgo1.GraphAnnotationText(i) = ""
If mnuShowPeakNumbers.Checked Then Pesgo1.GraphAnnotationText(i) = CStr(CInt(i))
Pesgo1.GraphAnnotationType(i) = 23 '23 it was before
Pesgo1.GraphAnnotationX(i) = xgraphData(2, i)
Pesgo1.GraphAnnotationY(i) = ygraphdata(9, i)
Pesgo1.GraphAnnotationColor(i) = QBColor(12)
Next i
Else
For i = 1 To NumberOfPeaks
Pesgo1.GraphAnnotationText(i) = ""
Pesgo1.GraphAnnotationType(i) = 23
Pesgo1.GraphAnnotationX(i) = -10000
Pesgo1.GraphAnnotationY(i) = -10000
Pesgo1.GraphAnnotationColor(i) = QBColor(12)
Next i
End If

If mnuSimulatePeaks.Checked And amsentpeaks = True Then
For i = 1 To NumberOfSimulatedPeaks
j = NumberOfPeaks + i
Pesgo1.GraphAnnotationText(j) = CStr(valori(i).h) & CStr(valori(i).k) & CStr(valori(i).l)
Pesgo1.GraphAnnotationType(j) = 13
Pesgo1.GraphAnnotationX(j) = CSng(valori(i).doitheta)
Pesgo1.GraphAnnotationY(j) = CSng(valori(i).ygrec)
Pesgo1.GraphAnnotationColor(j) = RGB(0, 0, 255)
Next i
Else
For i = 1 To NumberOfSimulatedPeaks
j = NumberOfPeaks + i
Pesgo1.GraphAnnotationText(j) = ""
Pesgo1.GraphAnnotationType(j) = 13
Pesgo1.GraphAnnotationX(j) = -10000
Pesgo1.GraphAnnotationY(j) = -10000
Pesgo1.GraphAnnotationColor(j) = QBColor(12)
Next i
End If

        For j = 0 To Pesgo1.Subsets - 1
        Pesgo1.SubsetsToLegend(j) = j 'and the legends
        Next j

If Pesgo1.Subsets > 1 Then
mnuDataApplyIntensityCorrection.Enabled = False 'valid only for raw data
mnuApplyZero.Enabled = False
mnuAddtoIntensity.Enabled = False
End If
Pesgo1.ScaleForYData = 0
Pesgo1.PEactions = 0
IamBusy False
DoEvents
Exit Sub
errortrap:
Pesgo1.PEactions = 0
IamBusy False
Err.Clear
Exit Sub
End Sub





Private Sub mnuGraphSmoothSG_Click()
Dim t As Integer, iii As Integer, i As Integer, j As Integer, suma As Double, points As Integer, degree As Integer
Dim xdata() As Double, ydata() As Double, solution() As Double, eroare As Boolean
Dim yg As Double
Static FirstResponse As Boolean
On Error GoTo errtrap
If Not (FirstResponse) Then
t = MsgBox("Warning :This routine does not work well on some datasets (it depends on the peaks width)." & vbCrLf & "Smooth the data applying a p-degree polynomial interpolation (p=1 to 6) over 2*n+1 points. N points will be considered to both the left and the right of the point to be corrected. You have to choose N and P; take care in choosing these values to avoid 'massaging' the data. " & vbCrLf & "Depending on how many data you have this procedure may be very slow.", vbOKCancel, prog_name & " - smooth")
If t = vbCancel Then Err.Raise 1102, , ""
FirstResponse = True
End If
t = InputBox("How many points to be used for smooth (n= 1 to 15; on each side) ?", "Powder-Smooth ", 2)
If t < 1 Or t > 15 Then Err.Raise 1101, , "Accepted values: 1 to 15. Try again..."
points = 2 * t + 1
t = InputBox("Input the polynomial degree to be used over " & CStr(points) & " points (1 to 6)", "Powder-Smooth ", 2)
If t < 1 Or t > 6 Then Err.Raise 1101, , "Accepted values 1 to 6. Try again..."
If t > (points + 1) Then Err.Raise 1101, , "Not enough points for this polynomial. Try again..."
degree = t
ReDim xdata(points), ydata(points), solution(degree + 1)
'apply smooth to active subset

For iii = 0 To 3
If workingSubset(iii).Checked Then Exit For
Next iii
Select Case iii
Case 0
'raw
iii = 1
Case 1
'smoothed
iii = 2
Case 2
'background stripped
iii = 4
Case 3
'ka2 stripped
iii = 6
End Select
'i here has the value of the checked menu, smooth will be applied to the activesubset...
    'raw data '' the normal case
        'the first and last t points remain the same...
        t = CInt((points - 1) / 2)
        For j = 1 To t + 1
        ygraphdata(2, j) = ygraphdata(iii, j)
        ygraphdata(2, numarvalori - j) = ygraphdata(iii, numarvalori - j)
        Next j
'i need to determin the interpolation polynomial, of degree degree, using points
IamBusy True
For i = t + 1 To numarvalori - t - 1
DoEvents 'give the processor some space
'build xdata and ydata
jj = 0
For j = i - t To i + t
jj = jj + 1
xdata(jj) = xgraphData(1, j)
ydata(jj) = ygraphdata(iii, j)
Next j
Call InterPolynomial(degree, points, xdata(), ydata(), 1E-17, solution(), eroare)
If eroare Then Err.Raise 1101, , "error in polynomial calculation, abort..."
'make the corrections
Call IntPolValue(degree, solution, CDbl(xgraphData(1, i)), yg, eroare)
If eroare Then Err.Raise 1101, , "error in polynomial calculation, abort..."
ygraphdata(2, i) = CSng(yg)
ygraphdata(3, i) = ygraphdata(1, i) - ygraphdata(2, i)
Next i
''Dim amrawdata, amsmoothdata, amdiffdata, ambackstrip, ambackonly, amk2strip, amk2only, ampeaks, amderivative As Boolean
IamBusy False
amsmoothdata = True
amdiffdata = True
'store all the data in the 2nd subset, compute the third by the difference
'go to refresh and use pevset for the data being in memory
'if smooth is OK then go to a adjustmenus routine
FrmGraph.workingSubset(1).Enabled = True 'may apply corrections on smoothed data
mnuGraphRefresh_Click
workingSubset_Click (1)

Exit Sub
errtrap:
IamBusy False
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub mnuInsertPeak_Click()
Static t As Boolean

mnuInsertPeak.Checked = Not (mnuInsertPeak.Checked)
If Not (t) Then
MsgBox ("You can manually add peaks on the graph (if you already searched them automatically). When this menu is Checked, a double click on the graph adds a new peak. However, if you automatically search the peaks these values will be lost.")
t = True
End If
Exit Sub
End Sub

Private Sub mnuKa2Dong_Click()
On Error GoTo errortrap
'this routine seems quite slow,...I have to improve it's speed
'at least I can make some operations before asking wlratio, etc
Dim minim As Double, jcount As Integer, eroare2 As Boolean
Dim eroare As Boolean, k As Single, j As Single, t As String, wl As Double, wl1 As Double, wl2 As Double, wlratio As Double, xout As Double
Dim Klevers() As Double, Kweights() As Double
Dim tempydata() As Single, tempxdata() As Single, maxval As Single, uncorrected(5000) As Single
Dim gradpolinom As Integer, nrpuncte As Integer
Dim points As Integer, degree As Integer
Dim solution() As Double, sinval As Double
Dim yg As Double
Dim xdata() As Double, ydata() As Double

ReDim tempydata(numarvalori), tempxdata(numarvalori)

        
        For iii = 0 To 3
            If workingSubset(iii).Checked Then Exit For
        Next iii
            Select Case iii
    Case 0 'raw
        iii = 1
    Case 1 'smoothed
        iii = 2
    Case 2 'background stripped
        iii = 4
    Case 3 'ka2 stripped
        iii = 6
        End Select



For i = 1 To numarvalori
tempxdata(i) = xgraphData(1, i)
tempydata(i) = ygraphdata(iii, i)
Next i






t = InputBox("Input the number of points of the Ka2 histogram calculation (3, 5, 7, 9, 15 or 25).  This procedure may take some time,...please wait ", prog_name, "7")
histo = CInt(t)
If histo < 3 Or histo > 25 Then Err.Raise 1101, , "Your input was " & CStr(histo) & " Accepted values 3, 5, 7, 9, 15 or 25. Aborting operation...."
ReDim Klevers(histo), Kweights(histo)
Select Case histo
Case 3
Klevers(1) = 0.998506815: Kweights(1) = 0.005134296
Klevers(2) = 0.997503913: Kweights(2) = 0.491686047
Klevers(3) = 0.99669946: Kweights(3) = 0.003179657

Case 5
Klevers(1) = 0.998471576: Kweights(1) = 0.00261441
Klevers(2) = 0.997935524: Kweights(2) = 0.011928014
Klevers(3) = 0.99750353: Kweights(3) = 0.480406807
Klevers(4) = 0.997163494: Kweights(4) = 0.002121807
Klevers(5) = 0.996606519: Kweights(5) = 0.002928802

Case 7
Klevers(1) = 0.998563433: Kweights(1) = 0.001580069
Klevers(2) = 0.998204025: Kweights(2) = 0.003463773
Klevers(3) = 0.997825027: Kweights(3) = 0.015566472
Klevers(4) = 0.997522195: Kweights(4) = 0.422601977
Klevers(5) = 0.997297615: Kweights(5) = 0.053632977
Klevers(6) = 0.996844235: Kweights(6) = 0.001572467
Klevers(7) = 0.996516288: Kweights(7) = 0.001615265

Case 9
Klevers(1) = 0.998609749: Kweights(1) = 0.001138001
Klevers(2) = 0.998334027: Kweights(2) = 0.00195272
Klevers(3) = 0.998054914: Kweights(3) = 0.004324464
Klevers(4) = 0.99776062: Kweights(4) = 0.019246541
Klevers(5) = 0.997527844: Kweights(5) = 0.394175823
Klevers(6) = 0.997327154: Kweights(6) = 0.079159001
Klevers(7) = 0.997028978: Kweights(7) = -0.003591547
Klevers(8) = 0.996734639: Kweights(8) = 0.002505604
Klevers(9) = 0.99646335: Kweights(9) = 0.001089392

Case 15
Klevers(1) = 0.998671599: Kweights(1) = 0.000614225
Klevers(2) = 0.998509111: Kweights(2) = 0.000810836
Klevers(3) = 0.998346447: Kweights(3) = 0.001134775
Klevers(4) = 0.998183442: Kweights(4) = 0.001723265
Klevers(5) = 0.998019704: Kweights(5) = 0.002968405
Klevers(6) = 0.997854063: Kweights(6) = 0.006433676
Klevers(7) = 0.997680649: Kweights(7) = -0.025753847
Klevers(8) = 0.997533314: Kweights(8) = 0.345872599
Klevers(9) = 0.997377391: Kweights(9) = 0.100578092
Klevers(10) = 0.997266106: Kweights(10) = 0.014493969
Klevers(11) = 0.997060614: Kweights(11) = -0.004176171
Klevers(12) = 0.996888005: Kweights(12) = 0.000678688
Klevers(13) = 0.996741151: Kweights(13) = 0.001610333
Klevers(14) = 0.996583672: Kweights(14) = 0.000918077
Klevers(15) = 0.996418168: Kweights(15) = 0.000585391

Case 25
Klevers(1) = 0.998706192: Kweights(1) = 0.000349669
Klevers(2) = 0.998608958: Kweights(2) = 0.000408044
Klevers(3) = 0.998511721: Kweights(3) = 0.000484578
Klevers(4) = 0.998414475: Kweights(4) = 0.000587457
Klevers(5) = 0.998317209: Kweights(5) = 0.000730087
Klevers(6) = 0.998219906: Kweights(6) = 0.000935685
Klevers(7) = 0.998122538: Kweights(7) = 0.001247401
Klevers(8) = 0.998025057: Kweights(8) = 0.001753233
Klevers(9) = 0.997927367: Kweights(9) = 0.002657209
Klevers(10) = 0.997829244: Kweights(10) = 0.004531817
Klevers(11) = 0.997730044: Kweights(11) = 0.009591103
Klevers(12) = 0.997626987: Kweights(12) = 0.034998436
Klevers(13) = 0.997535705: Kweights(13) = 0.2876498
Klevers(14) = 0.997458223: Kweights(14) = 0.074954321
Klevers(15) = 0.997346989: Kweights(15) = 0.065000871
Klevers(16) = 0.997277763: Kweights(16) = 0.016762729
Klevers(17) = 0.997161452: Kweights(17) = -0.00306221
Klevers(18) = 0.997057942: Kweights(18) = -0.002717412
Klevers(19) = 0.996982688: Kweights(19) = -0.000902322
Klevers(20) = 0.99686108: Kweights(20) = 0.000915701
Klevers(21) = 0.996769728: Kweights(21) = 0.001036484
Klevers(22) = 0.996675255: Kweights(22) = 0.000808199
Klevers(23) = 0.996578407: Kweights(23) = 0.000539899
Klevers(24) = 0.996480641: Kweights(24) = 0.000398896
Klevers(25) = 0.99638324: Kweights(25) = 0.000330325
Case Else
Err.Raise 1101, , "Your input was " & CStr(histo) & " Accepted values 3, 5, 7, 9, 15 or 25. Aborting operation...."
End Select

t = InputBox("Input the Ka2/Ka1 ratio :", prog_name, "0.49")
If Len(CStr(t)) = 0 Then Exit Sub
wlratio = Val(t)
If wlratio < 0.01 Or wlratio > 0.99 Then Err.Raise 1101, , "Wrong value..."


For i = 1 To histo
tempydata(i) = tempydata(i) * (1 / (1 + wlratio))
Next i
For i = numarvalori - histo To numarvalori
tempydata(i) = tempydata(i) * (1 / (1 + wlratio))
Next i

Screen.MousePointer = 11
DoEvents

'tempydata is the a2 contribution
jcount = 1

For i = histo To numarvalori - histo
    ksuma = 0
        sinval = Sin((xgraphData(1, i) / 2 / rd))
    For j = 1 To histo
        ''xx = rd * 2 * asin((Klevers(j) * sinval))
        ''eveluate the arcsin here, is faster than call
        xx = rd * 2 * (Atn((Klevers(j) * sinval) / Sqr(-(Klevers(j) * sinval) * (Klevers(j) * sinval) + 1)))
        DoEvents 'give the processor some space
        jcount = Fix(((xx - startx) / stepx + 0.5))
        
        yg = tempydata(jcount) + (tempydata(jcount + 1) - tempydata(jcount)) * ((xx - xgraphData(1, jcount)) / (xgraphData(1, jcount + 1) - xgraphData(1, jcount)))
    'jcount seems to be always smaller than i
        ksuma = ksuma + Kweights(j) * yg
    Next j
    tempydata(i) = tempydata(i) - ksuma
'ygraphdata(6, i) = tempydata(i) - ksuma
Next i

'make a smooth and asign
For i = 8 To numarvalori - 8
DoEvents 'give the processor some space
'build xdata and ydata
suma = 7 * tempydata(i)
suma = suma + 6 * (tempydata(i + 1) + tempydata(i - 1)) + 3 * (tempydata(i + 2) + tempydata(i - 2)) - 2 * (tempydata(i + 3) + tempydata(i - 3))
tempydata(i) = suma / 21
ygraphdata(6, i) = suma / 21
ygraphdata(7, i) = ygraphdata(iii, i) - ygraphdata(6, i)
Next i

'For i = 1 To numarvalori
'ygraphdata(6, i) = tempydata(i)
'ygraphdata(7, i) = ygraphdata(iii, i) - ygraphdata(6, i)
'Next i

amk2strip = True
FrmGraph.workingSubset(3).Enabled = True 'may apply correction to ka2 stripped
Screen.MousePointer = 0
mnuGraphRefresh_Click
workingSubset_Click (3)

Exit Sub
errortrap:
Screen.MousePointer = 0
IamBusy False
If Err.Number = 1101 Then
MsgBox Err.Description
Else
MsgBox "Error encountered:  " & Err.Description
End If
Exit Sub
End Sub

Private Sub mnuPeakSonneveld_Click()
On Error GoTo errortrap
Dim t As Double, i As Integer, iii As Integer, Pic(500) As peaks, MinWidth As Single, scalefactor As Single
Dim eroare As Boolean, threshold As Single, maxpeaks As Integer
ChDir (App.Path & "\peak\")
'ChDrive "g:\"
'temporaire
ChDrive (App.Path)
'MsgBox CurDir
'use pk_transd program for converting a data file
'the file to be converted should be written in the format:
'name (a80)
'the data file must be named pw_pkdta, the program pk_transd will
'save a peakin.dat which can be used further on
'the peak search should be done in maximum four steps of 8000 points each
Static mesaj As Boolean
ampeaks = False
NumberOfPeaks = 0
Pesgo1.ShowGraphAnnotations = False
If Not (mesaj) Then
MsgBox "This use the method of Sonneveld and Visser and most of their routine. You can adjust the parameters listed in the table of the graph for your data."
'the peaks are stored in the 9th position,''the flag is ampeaks
mesaj = True
End If
DoEvents
'first see to which curve this applies...
'threshold should be input by the user
'threshold = threshold / 100
For iii = 0 To 3
If workingSubset(iii).Checked Then Exit For
Next iii
Select Case iii
Case 0
'raw
iii = 1
Case 1
'smoothed
iii = 2
Case 2
'background stripped
iii = 4
Case 3
'ka2 stripped
iii = 6
End Select
Dim xdata() As Double, ydata() As Double, switch As Integer, localmin As Single, som1 As Single, biggestSum As Single, lastval As Double, minval As Double, drd2 As Single, maxval As Double, ddydata() As Double, suma As Single, minvald2 As Single
ReDim xdata(numarvalori)
ReDim ydata(numarvalori)
ReDim ddydata(numarvalori)
'make four searches each of 8000 points, then add all the peaks together
Dim numvaloriLimit As Integer
If numarvalori > 8980 Then
numarvaloriLimit = 8800
MsgBox "The routine PeakSearch of Sonneveld and Visser is limited to 9000 points. The search will be made only for the first 9000 points "
Else
numarvaloriLimit = numarvalori

End If

    'j shows the pass
    For i = 1 To numarvaloriLimit
        xdata(i) = xgraphData(1, i)
        ydata(i) = ygraphdata(iii, i)
    Next i
        outfil = FreeFile
        
 
        
        On Error GoTo thereisnoFile
        Kill App.Path & "\peak\pw_pkdta"
        Kill App.Path & "\peak\peakin.dat"
        Kill App.Path & "\peak\list.lst"
        Kill App.Path & "\peak\lstpk.lst"
        On Error GoTo errortrap
        Open App.Path & "\peak\pw_pkdta" For Output As #outfil
        
        
            Print #outfil, prog_name & title
            Print #outfil, Format$(Format$(startx, "#0.000##"), "@@@@@@@@"); Format$(Format$(stepx, "#0.000##"), "@@@@@@@@")
            For i = 1 To Fix(numarvaloriLimit / 10)
            Print #outfil, Format$(Format$(Val(ydata((i - 1) * 10 + 1)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 2)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 3)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 4)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 5)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 6)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 7)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 8)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 9)), "#######0.00#"), "@@@@@@@@@@@@") + Format$(Format$(Val(ydata((i - 1) * 10 + 10)), "#######0.00#"), "@@@@@@@@@@@@")
            Next i
        Print #outfil, "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 " & "        0.0 "
        Close #outfil
    DoEvents 'give the processor some space
    t = ShellAndWait(App.Path & "\peak\pk_transd.exe", vbMinimizedNoFocus)
    'added "app.path\peak on 20 06 2002
    DoEvents
    'read the value of the scalefactor
    Dim a1 As Single, a2 As Single, a3 As Single
    '''Close
    'inutile ?
    Open App.Path & "\peak\peakin.dat" For Input As #outfil
    Line Input #outfil, title
    Input #outfil, a1, a2, a3, scalefactor
    Close #outfil
    'the intensity read in pklst.lst should be updated
    t = ShellAndWait(App.Path & "\peak\peak.exe", vbMinimizedNoFocus)
    ''added "app.path\peak on 20 06 2002
    If t = 0 Then raport "seems to be OK..."
    DoEvents
    
    Open App.Path & "\peak\lstpk.lst" For Input As #outfil
    'added "app.path\peak on 20 06 2002
    Line Input #outfil, title
    Input #outfil, NumberOfPeaks
    For j = 1 To NumberOfPeaks
    Input #outfil, Pic(j).Xposition, Pic(j).Yintensity, a1
    Pic(j).Yintensity = Pic(j).Yintensity * scalefactor + Pic(j).Yintensity * scalefactor * 0.02
    ygraphdata(9, j) = Pic(j).Yintensity
    xgraphData(2, j) = Pic(j).Xposition
    Next j
    Close #outfil
        
        
        If NumberOfPeaks < 1 Then Err.Raise 1101, , "No peaks detected under these conditions..."

t = MsgBox(CStr(NumberOfPeaks) & " peaks have been detected. The details of the peaks are saved in the file List.lst, subdirectory \PEAK. Do you want to see the list now ? ", vbYesNo + vbDefaultButton2, prog_name)
      If t = vbYes Then Shell "notepad.exe list.lst", vbNormalFocus
    DoEvents
    

''----------

ampeaks = True

mnuGraphExportPeaks.Enabled = True 'if I found peaks I may export them
mnuGraphRefresh_Click
IamBusy (False)
Exit Sub

thereisnoFile:
Resume Next

errortrap:
MsgBox " Error encountered in mnuPeakSonneveld, abort. " & Err.Description
IamBusy False
Exit Sub
Close
End Sub

Private Sub mnuSelectDone_Click()
donecode = True
DoEvents
End Sub

Private Sub mnuSetPoint_Click()
setpoints = True
End Sub

Private Sub mnuShowPeakNumbers_Click()
If mnuShowPeaks.Checked Then
mnuShowPeakNumbers.Checked = Not (mnuShowPeakNumbers.Checked)
mnuGraphRefresh_Click
End If
End Sub

Private Sub mnuShowPeaks_Click()
mnuShowPeaks.Checked = Not (mnuShowPeaks.Checked)
If mnuShowPeaks.Checked = False Then mnuShowPeakNumbers.Checked = False
mnuGraphRefresh_Click
Exit Sub
End Sub

Private Sub mnuSimulatePeaks_Click()
mnuSimulatePeaks.Checked = Not (mnuSimulatePeaks.Checked)
CallFromGraphForCell = True
''Me.Visible = False

frmRefine.Show
mnuGraphRefresh_Click
Exit Sub
End Sub

Private Sub mnuSmoothAdjacent_Click()
Dim t As Integer, i As Integer, j As Integer, suma As Double, iii As Integer
On Error GoTo errtrap
t = InputBox("How many points <on each side> to be used for averaging (1-9)? " & vbCrLf & "<warning: too many will flat out the peaks>", "Powder-Smooth adjacent", 1)
If t < 1 Or t > 9 Then Err.Raise 1101, , "Accepted values: 1 to 9. Try again..."
For iii = 0 To 3
If workingSubset(iii).Checked Then Exit For
Next iii
Select Case iii
Case 0
'raw
iii = 1
Case 1
'smoothed
iii = 2
Case 2
'background stripped
iii = 4
Case 3
'ka2 stripped
iii = 6
End Select
'smooth will be applied to the activesubset...
'if the data is not available then warn and get out of here
'raw data '' the normal case
'the first and last t points remain the same...
For j = 1 To t + 1
ygraphdata(2, j) = ygraphdata(iii, j)
ygraphdata(2, numarvalori - j + 1) = ygraphdata(iii, numarvalori - j + 1)
Next j

For j = t To numarvalori - t
suma = 0
For i = -t To t Step 1
suma = suma + ygraphdata(iii, i + j)
Next i
ygraphdata(2, j) = CSng(suma / (2 * t + 1))
Next j
'XGraphData is in all 8 cases the same, I will need only 3 dimension for this
For i = 1 To numarvalori
ygraphdata(3, i) = ygraphdata(1, i) - ygraphdata(2, i)
Next i
amsmoothdata = True
amdiffdata = True
'store all the data in the 2nd subset, compute the third by the difference
'go to refresh and use pevset for the data being in memory
'if smooth is OK then go to a adjustmenus routine
FrmGraph.workingSubset(1).Enabled = True 'may apply corrections on smoothed data
mnuGraphRefresh_Click
workingSubset_Click (1)

Exit Sub
errtrap:
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub


Private Sub mnuSmoothSavitzkyParabola_Click()
Dim iii As Integer, i As Integer, j As Integer, suma As Single
On Error GoTo errortrap
For iii = 0 To 3
If workingSubset(iii).Checked Then Exit For
Next iii
Select Case iii
Case 0
'raw
iii = 1
Case 1
'smoothed
iii = 2
Case 2
'background stripped
iii = 4
Case 3
'ka2 stripped
iii = 6
End Select
        For j = 1 To 9
        ygraphdata(2, j) = ygraphdata(iii, j)
        ygraphdata(2, numarvalori - j + 1) = ygraphdata(iii, numarvalori - j + 1)
        Next j
'i need to determin the interpolation polynomial, of degree degree, using points
Me.MousePointer = 11
For i = 8 To numarvalori - 8
DoEvents 'give the processor some space
'build xdata and ydata
suma = 7 * ygraphdata(iii, i)
suma = suma + 6 * (ygraphdata(iii, i + 1) + ygraphdata(iii, i - 1)) + 3 * (ygraphdata(iii, i + 2) + ygraphdata(iii, i - 2)) - 2 * (ygraphdata(iii, i + 3) + ygraphdata(iii, i - 3))
ygraphdata(2, i) = suma / 21
ygraphdata(3, i) = ygraphdata(1, i) - ygraphdata(2, i)
Next i
''Dim amrawdata, amsmoothdata, amdiffdata, ambackstrip, ambackonly, amk2strip, amk2only, ampeaks, amderivative As Boolean
Me.MousePointer = 0
amsmoothdata = True
amdiffdata = True
'store all the data in the 2nd subset, compute the third by the difference
'go to refresh and use pevset for the data being in memory
'if smooth is OK then go to a adjustmenus routine
FrmGraph.workingSubset(1).Enabled = True 'may apply corrections on smoothed data
'put the Smooth on the graph

mnuGraphRefresh_Click
workingSubset_Click (1)

Exit Sub
errortrap:
Me.MousePointer = 0
amsmoothdata = False
amdiffdata = fasle
 MsgBox Err.Description
Err.Clear
Exit Sub


End Sub


Private Sub mScale_Click(Index As Integer)
Dim jj As Integer
For jj = 1 To 3
mScale(jj - 1).Checked = False
Next jj
mScale(Index).Checked = True
Pesgo1.YAxisScaleControl = Index
DoEvents

End Sub

Private Sub Pesgo1_DblClick()
DoEvents
On Error GoTo errortrap
Dim t As Integer, i As Single, zzx As Single, zzy As Single, yy As Single, j As Single, tt As Single

If mnuInsertPeak.Checked Then
t = MsgBox("Do you want to add a peak ?", vbYesNo, prog_name)
If t = vbNo Then Exit Sub
''
tt = InputBox("Enter the 2theta value at which you want to insert the peak:", prog_name & " - Insert peak")
If tt < xgraphData(1, 1) Or t > xgraphData(1, numarvalori) Then Exit Sub
''ReDim ygraphdata(9, NumberOfPeaks + 1)
''ReDim xgraphData(2, NumberOfPeaks + 1)
yy = InputBox("Enter the Y position where you'd like to see the marker ", prog_name & " - Insert peak")
DoEvents



For i = 1 To NumberOfPeaks
If xgraphData(2, i) < tt And xgraphData(2, i + 1) > tt Then Exit For
Next i
''i here is the position where we should insert the peak
For j = NumberOfPeaks To i Step -1
xgraphData(2, j + 1) = xgraphData(2, j)
ygraphdata(9, j + 1) = ygraphdata(9, j)
Next j
xgraphData(2, i) = CSng(tt)
ygraphdata(9, i) = yy
NumberOfPeaks = NumberOfPeaks + 1
mnuGraphRefresh_Click
End If
Exit Sub
errortrap:
Exit Sub
End Sub

Private Sub Pesgo1_GraphAnnotHotSpot(DblClk As Integer, nIndex As Integer)
Dim t As Integer, i As Single
If NumberOfPeaks < 1 Then Exit Sub
If nIndex > NumberOfPeaks Then Exit Sub
t = MsgBox("Do you want to remove this peak ?", vbYesNo, prog_name)
If t = vbNo Then Exit Sub
If nIndex = NumberOfPeaks Then NumberOfPeaks = NumberOfPeaks - 1: Exit Sub
For i = nIndex To NumberOfPeaks - 1
ygraphdata(9, i) = ygraphdata(9, i + 1)
xgraphData(2, i) = xgraphData(2, i + 1)
Next i
NumberOfPeaks = NumberOfPeaks - 1
mnuGraphRefresh_Click

Exit Sub
End Sub

Sub workingSubset_Click(Index As Integer)
Dim i As Integer
For i = 0 To 3
workingSubset(i).Checked = False
Next i
workingSubset(Index).Checked = True
''Pesgo1.RandomSubsetsToGraph(-1) = 0
'Pesgo1.ScrollingSubsets = 1
Select Case Index
Case 0
i = 0
Case 1
i = 1
Case 2
i = 3
Case 3
i = 5
Case Else
'impossible
End Select
Pesgo1.RandomSubsetsToGraph(0) = i
Pesgo1.PEactions = 0
Exit Sub
'change here the subset shown and upgrade the graph
End Sub

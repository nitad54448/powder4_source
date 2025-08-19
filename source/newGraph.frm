VERSION 5.00
Object = "{60C2F168-424A-101C-AA69-0040052BC4EA}#1.0#0"; "pesgo32.ocx"
Begin VB.Form newGraph 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "gDBWS_Graph"
   ClientHeight    =   4750
   ClientLeft      =   110
   ClientTop       =   590
   ClientWidth     =   7610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4750
   ScaleWidth      =   7610
   StartUpPosition =   3  'Windows Default
   Begin PesgoLib.Pesgo PE 
      Height          =   4088
      Left            =   120
      OleObjectBlob   =   "newGraph.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   7208
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mOpenPlotinfo 
         Caption         =   "Open plotinfo type file"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveData 
         Caption         =   "&Save data"
      End
      Begin VB.Menu mFileOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mCustomize 
         Caption         =   "Customize &Graph"
      End
      Begin VB.Menu mExport 
         Caption         =   "E&xport"
      End
   End
   Begin VB.Menu mSet 
      Caption         =   "Set"
      Begin VB.Menu mXScale 
         Caption         =   "X Scale"
         Begin VB.Menu mSetMinX 
            Caption         =   "min X value"
         End
         Begin VB.Menu mSetMaxX 
            Caption         =   "max X value"
         End
      End
      Begin VB.Menu mIobs 
         Caption         =   "Iobs - Y Scale"
         Begin VB.Menu mSetMinIobs 
            Caption         =   "min Y value"
         End
         Begin VB.Menu mSetMaxIobs 
            Caption         =   "max Y value"
         End
      End
      Begin VB.Menu mIcalc 
         Caption         =   "I calc - Y scale"
         Begin VB.Menu mSetMinIcal 
            Caption         =   "min Y value"
         End
         Begin VB.Menu mSetMaxIcal 
            Caption         =   "max Y value"
         End
      End
      Begin VB.Menu mDiff 
         Caption         =   "Difference - Y Scale"
         Begin VB.Menu mSetDiffMinY 
            Caption         =   "min Y value"
         End
         Begin VB.Menu mSetDiffMaxY 
            Caption         =   "max Y value "
         End
      End
      Begin VB.Menu m 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mHKLmarkers 
         Caption         =   "HKL markers"
         Begin VB.Menu mTickMarkerSize 
            Caption         =   "Size (Ka1)"
         End
         Begin VB.Menu mSizek2 
            Caption         =   "Size (Ka2)"
         End
         Begin VB.Menu mPosition 
            Caption         =   "Position"
         End
         Begin VB.Menu ml 
            Caption         =   "-"
            HelpContextID   =   2
            Index           =   0
         End
         Begin VB.Menu mTickShowHKL 
            Caption         =   "Show HKL values"
         End
         Begin VB.Menu mTickShowIntensity 
            Caption         =   "show Intensity values"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mVerticalShift 
            Caption         =   "Multiphase Shift "
         End
      End
      Begin VB.Menu ml2 
         Caption         =   "-"
      End
      Begin VB.Menu mPointSize 
         Caption         =   "Point size"
         Begin VB.Menu mpSize 
            Caption         =   "Micro"
            Index           =   0
         End
         Begin VB.Menu mpSize 
            Caption         =   "Small"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mpSize 
            Caption         =   "Medium"
            Index           =   2
         End
         Begin VB.Menu mpSize 
            Caption         =   "Large"
            Index           =   3
         End
      End
      Begin VB.Menu mFontSizeAll 
         Caption         =   "Font Size (all)"
      End
      Begin VB.Menu mFontLegend 
         Caption         =   "Font Size (legends)"
      End
   End
End
Attribute VB_Name = "newGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub fileopenplotinfo(inputfile As String)

On Error GoTo errorTRAP
minXval = 0: maxXval = 0: minYIcalc = 0: maxYIcalc = 0: minYIobs = 0: maxYIobs = 0: minYIdiff = 0: maxYIdiff = 0
TickPhaseShift = 100 'in % for each phase
CustomFontsizeglobalcntl = 1
CustomFontsizelegendcntl = 0.9
'Warning: data is scaled to 1000


inpfil = FreeFile
Open inputfile For Input As #inpfil
Line Input #inpfil, sTitle
Line Input #inpfil, sLine
nPhas = CInt(Val(Mid$(sLine, 18)))
ReDim DBW(nPhas)
Line Input #inpfil, sLine
For i = 1 To nPhas
DBW(i).nreflex = CInt(Mid$(sLine, 38 + 4 * (i - 1), 4))
Next i
Line Input #inpfil, sLine ' contains bragg pos string
For i = 1 To nPhas
Line Input #inpfil, sLine ' this line is empty
Line Input #inpfil, sLine
DBW(i).sTitle = sLine

For j = 1 To DBW(i).nreflex
Line Input #inpfil, sLine
If UCase$(Mid$(sLine, 1, 6)) = " NPTSZ" Then nval = CInt(Val(Mid$(sLine, 7))): DBW(i).nreflex = j - 1: Exit For

DBW(i).sDoiTeta(j) = CSng(Val(sLine))
    If InStr(sLine, "K1") Then
        DBW(i).bK1(j) = True
    Else
        DBW(i).bK2(j) = True
    End If
DBW(i).nH(j) = CInt(Val(Mid$(sLine, 14, 3)))
DBW(i).nK(j) = CInt(Val(Mid$(sLine, 18, 3)))
DBW(i).nL(j) = CInt(Val(Mid$(sLine, 22, 3)))
DBW(i).sIntens(j) = CSng(Val(Mid$(sLine, 26, 9)))
Next j
Next i
'another empty line
'Line Input #inpfil, sLine
maxval = 0

Do Until Mid$(sLine, 1, 6) = " NPTSZ"
Line Input #inpfil, sLine
If Mid$(sLine, 1, 6) = " NPTSZ" Then nval = CInt(Val(Mid$(sLine, 7)))
Loop
ReDim dbX(nval), dbYraw(nval), dbYsum(nval), dbYdiff(nval)

Line Input #inpfil, sLine
startx = CSng(Val(Mid$(sLine, 8)))
Line Input #inpfil, sLine
stepx = CSng(Val(Mid$(sLine, 8)))
Line Input #inpfil, sLine 'another empty line

minXval = startx
maxXval = startx + nval * stepx


For i = 0 To nval - 1
Input #inpfil, dbYraw(i), dbYsum(i)
If dbYraw(i) > maxYIobs Then maxYIobs = dbYraw(i)
If dbYraw(i) < minYIobs Then minYIobs = dbYraw(i)
If dbYsum(i) > maxYIcalc Then maxYIcalc = dbYsum(i)
If dbYsum(i) < minYIcalc Then minYIcalc = dbYsum(i)

dbX(i) = CSng(i * stepx + startx)
dbYdiff(i) = dbYraw(i) - dbYsum(i)

If dbYdiff(i) > maxYIdiff Then maxYIdiff = dbYdiff(i)
If dbYdiff(i) < minYIdiff Then minYIdiff = dbYdiff(i)
Next i


TickMarkShift = 0 '0 means on top
TickMarkSize = 0.03 * (maxYIobs - minYIobs) '3 percent
TickMarkAlpha2Size = 0.5 * TickMarkSize 'half of that of alpha1

If maxYIcalc > maxYIobs Then maxYIobs = maxYIcalc
If minYIcalc < minYIobs Then minYIobs = minYIcalc
maxYIcalc = maxYIobs
minYcalc = minYIobs

For i = 1 To nPhas
For j = 1 To DBW(i).nreflex
DBW(i).sIntens(j) = maxval / 1000 * DBW(i).sIntens(j)
Next j
Next i

Close #inpfil


colorSubset2 = RGB(255, 0, 0)
colorSubset1 = RGB(0, 255, 0)
colorSubset3 = RGB(0, 0, 255)

PE.Subsets = 3
PE.points = numarvalori
PE.PrepareImages = True

PE.Subsets = 3
PE.points = nval

With PE
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
    .CursorMode = NoCursor
    .CursorPageAmount = 50
    .CursorPromptStyle = XandYvalues
    .CursorMode = NoCursor
    .CursorPromptTracking = True
    .CurveGranularity = FineLines
    .DataPrecision = ThreeDecimals
    .DataShadows = False
    .GridLineControl = NoGrid
    .MainTitle = ""
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
    '.RYAxisLongTicks = False
    .RYAxisColor = RGB(255, 25, 25)
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
    .YAxisColor = RGB(25, 25, 255)
    .AxesAnnotationTextSize = 50
    
    .XAxisLabel = "2 theta /deg."
    .YAxisLabel = "I obs /counts"
    .SubsetColors(0) = RGB(25, 25, 255)
    .SubsetColors(1) = RGB(255, 25, 25)
    .SubsetColors(2) = RGB(25, 255, 25)
    .SubsetPointTypes(0) = PEPT_SQUARE
    .SubsetPointTypes(1) = PEPT_DIAMONDSOLID
    .SubsetPointTypes(2) = PEPT_PLUS
    .SubTitle = ""
End With

Call refreshRietveldGraph

Exit Sub
errorTRAP:

MsgBox "An error has occured, location: gDBWS Graph. " & Err.Description
Err.Clear
Exit Sub






End Sub



Sub refreshRietveldGraph()
On Error GoTo errorTRAP
    Dim dk As Integer, GraphTempData() As Single
    ReDim GraphTempData(nval * CInt(PE.Subsets))
        For i = 0 To nval - 1
            GraphTempData(i) = dbYraw(i)
            GraphTempData(nval + i) = dbYsum(i)
            GraphTempData(2 * nval + i) = dbYdiff(i)
        Next i
        result = PEvset(PE, PEP_faYDATA, GraphTempData(0), nval * PE.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Ydata array)..."

For j = 0 To PE.Subsets - 1
        For i = 1 To nval
            GraphTempData(j * nval + i - 1) = dbX(i)
        Next i
Next j
        'send x data by pevset
        result = PEvset(PE, PEP_faXDATA, GraphTempData(0), nval * PE.Subsets)
        If result = 0 Then MsgBox "An error occured in the DLL transfer routine (Xdata array)..."


PE.FontSizeGlobalCntl = CustomFontsizeglobalcntl
PE.FontSizeLegendCntl = CustomFontsizelegendcntl


PE.SubsetLineTypes(0) = 0
PE.SubsetLineTypes(1) = 0 'thin solid line
PE.SubsetLineTypes(2) = 0
PE.RandomSubsetsToGraph(0) = -1
PE.SubsetLabels(0) = "I obs"
PE.SubsetLabels(1) = "I calc"
PE.SubsetLabels(2) = "Difference"
PE.ManualScaleControlY = ManualMinAndMax
PE.ManualMinY = minYIobs
PE.ManualMaxY = maxYIobs


PE.AllowPlotCustomization = True
PE.AllowUserInterface = True

PE.MultiAxesSubsets(0) = 2
PE.MultiAxesSubsets(1) = 1

PE.MultiAxesSubsets(1) = 1
PE.RYAxisComparisonSubsets = 1
PE.RYAxisScaleControl = Linear
PE.ManualScaleControlRY = ManualMinAndMax
PE.ManualMinRY = minYIcalc
PE.ManualMaxRY = maxYIcalc
PE.ForceRightYAxis = True

PE.ManualScaleControlX = ManualMinAndMax
PE.ManualMinX = minXval
PE.ManualMaxX = maxXval

PE.RYAxisLabel = "I calc /counts"
PE.RYAxisLongTicks = True
'PE.ShowRYAxis = TicksPlusAxisLabels

PE.MultiAxesSeparators = 0
PE.MultiAxesProportions(0) = 0.8
PE.MultiAxesProportions(1) = 0.2
PE.WorkingAxis = 1
PE.YAxisLabel = "Iobs -Icalc"
PE.ManualScaleControlY = ManualMinAndMax
PE.ManualMinY = minYIdiff
PE.ManualMaxY = maxYIdiff

PE.WorkingAxis = 0
DoEvents
PE.AllowAnnotationControl = False

For j = 1 To nPhas
k = 0
If j > 1 Then k = DBW(j - 1).nreflex
For i = 1 To DBW(j).nreflex
dk = 2 * (i + k) - 1

PE.GraphAnnotationType(dk) = PEGAT_THINSOLIDLINE
PE.GraphAnnotationText(dk) = "" ' CStr(DBW(j).nH(i)) & " " & CStr(DBW(j).nK(i)) & " " & CStr(DBW(j).nL(i))
PE.GraphAnnotationX(dk) = DBW(j).sDoiTeta(i)

PE.GraphAnnotationY(dk) = TickMarkShift - (j - 1) * TickPhaseShift / 100 * TickMarkSize
PE.GraphAnnotationColor(dk) = QBColor(1)
PE.GraphAnnotationType(dk + 1) = PEGAT_LINECONTINUE
If mTickShowHKL.Checked Then
If DBW(j).bK1(i) Then PE.GraphAnnotationText(dk + 1) = CStr(DBW(j).nH(i)) & " " & CStr(DBW(j).nK(i)) & " " & CStr(DBW(j).nL(i))
Else
PE.GraphAnnotationText(dk + 1) = ""
End If
PE.GraphAnnotationX(dk + 1) = DBW(j).sDoiTeta(i)
If DBW(j).bK1(i) Then
PE.GraphAnnotationY(dk + 1) = TickMarkShift - (j - 1) * TickPhaseShift / 100 * TickMarkSize - TickMarkSize
    Else
PE.GraphAnnotationY(dk + 1) = TickMarkShift - (j - 1) * TickPhaseShift / 100 * TickMarkSize - TickMarkAlpha2Size
    End If
PE.GraphAnnotationColor(dk + 1) = QBColor(1)
Next i
Next j
PE.ShowGraphAnnotations = True
PE.ZoomMinY = 0
PE.PointSize = nPointSize
PE.PEactions = 0
IamBusy False
DoEvents
Exit Sub
errorTRAP:
Exit Sub
End Sub

Sub Form_Load()
Call fileopenplotinfo(inputfile)

End Sub

Private Sub Form_Resize()
On Error GoTo errtrap
Dim ratio As Single
If Me.Width < 4000 Then Me.Width = 4000
If Me.Height < 4000 Then Me.Height = 4000
PE.Width = Me.Width - 20
PE.Height = Me.Height - 600
'''mnuGraphRefresh_Click


PE.top = 0 '(Me.Height - PE.Height) / 12
PE.left = 0 ' (Me.Width - PE.Width) / 3
Exit Sub
errtrap:
Exit Sub
End Sub

Sub Form_Unload(Cancel As Integer)
Dim t As Integer
t = MsgBox("Are you sure you want to close this window ? ", vbDefaultButton1 + vbOKCancel, prog_name & " - gDBWS graphic")
If t = vbOK Then
newGraph.Visible = False
Unload Me
Else
Cancel = 1
Exit Sub
End If
End Sub

Private Sub mCustomize_Click()
PE.PElaunchcustomize

End Sub

Private Sub mExport_Click()
PE.PElaunchexport

End Sub

Private Sub mFileOptions_Click()
Dim a As Boolean
a = PE.PElaunchpopupmenu(1, 1)
End Sub

Private Sub mFontLegend_Click()

On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input a number denoting the size of fonts for legends: 0.5 = small, 1.5 = large ", prog_name & " - all fonts", CStr(CustomFontsizelegendcntl))
If t >= 0.5 And t <= 1.5 Then CustomFontsizelegendcntl = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub


End Sub

Private Sub mFontSizeAll_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input a number denoting the size of font : 0.5 = small, 1.5 = large ", prog_name & " - all fonts", CStr(CustomFontsizeglobalcntl))
If t >= 0.5 And t <= 1.5 Then CustomFontsizeglobalcntl = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub
End Sub

Sub mOpenPlotinfo_Click()
On Error GoTo errorTRAP
'here set up a control file with a generic name, let s say
Dim return_code As Boolean
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
'dimension parts here
return_code = False
Call fileopenplotinfo(inputfile)

Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mPosition_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the position for the marker (in counts)", prog_name & " - Y scale", CStr(CInt(TickMarkShift)))
If t > -1000 Then TickMarkShift = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub


End Sub

Private Sub mpSize_Click(Index As Integer)
Dim i As Integer
For i = 0 To 3
mpSize(i).Checked = False
Next i
mpSize(Index).Checked = True

Select Case Index
Case 0
nPointSize = Micro
Case 1
nPointSize = 0 'Small
Case 2
nPointSize = 1 'Medium
Case 3
nPointSize = 2 'Large
End Select

Call refreshRietveldGraph


End Sub


Private Sub mSaveData_Click()
On Error GoTo errorTRAP
Dim returncode As Boolean
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
'MsgBox "A file with Iobs, I calc, and hkl data was saved as gDBWS_pw.out"
Open outputfile For Output As outfil
Print #outfil, "gDBWS - Graphic data export "; prog_name & "  -  " & Now
Print #outfil, "2 theta,   I obs,    I calc,  I obs - I calc "
For i = 0 To nval - 1
Print #outfil, sForFormat(dbX(i), "F8.4") & " " & sForFormat(dbYraw(i), "F10.3") & " " & sForFormat(dbYsum(i), "F10.3") & " " & sForFormat(dbYdiff(i), "F8.4")
Next i
Print #outfil, "number of phases: " & CStr(nPhas)
For i = 1 To nPhas
Print #outfil, DBW(i).sTitle
Print #outfil, "no., K1 or K2, 2 theta, intens., H, K, L"
For j = 1 To DBW(i).nreflex
sLine = "    2   "
If DBW(i).bK1(j) Then sLine = "    1   "
Print #outfil, sForFormat(j, "i5") & sLine & sForFormat(DBW(i).sDoiTeta(j), "F9.4") & "  " & sForFormat(DBW(i).sIntens(j), "F10.3") & "  " & sForFormat(DBW(i).nH(j), "I4") & " " & sForFormat(DBW(i).nK(j), "I4") & " " & sForFormat(DBW(i).nL(j), "I4")
DBW(i).sDoiTeta(j) = CSng(Val(sLine))
Next j
Next i
Close #outfil

Exit Sub
errorTRAP:
MsgBox Err.Description
Close
Exit Sub
End Sub

Private Sub mSetDiffMaxY_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the maximum value for the difference Y scale ", prog_name & " - Y scale", CStr(maxYIdiff))
If t > -100 Then maxYIdiff = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mSetDiffMinY_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the minimum value for the Difference scale ", prog_name & " - diff Y scale", CStr(minYIdiff))
 minYIdiff = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mSetMaxIcal_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the maximum value for the Y scale (I calc)", prog_name & " - Y scale", CStr(maxYIcalc))
If t > -10 Then maxYIcalc = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mSetMaxIobs_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the maximum value for the Y scale (I obs)", prog_name & " - Y scale", CStr(maxYIobs))
If t > -10 Then maxYIobs = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mSetMaxX_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the maximum value for the X scale", prog_name & " - X scale", CStr(maxXval))
If t > -10 And t < 178 Then maxXval = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub
End Sub

Private Sub mSetMinIcal_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the minimum value for the Y scale (I calc)", prog_name & " - Y scale", CStr(minYIcalc))
 minYIcalc = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mSetMinIobs_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the minimum value for the Y scale (I obs)", prog_name & " - Y scale", CStr(minYIobs))
minYIobs = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub

End Sub

Private Sub mSetMinX_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the minimum value for the X scale", prog_name & " - X scale", CStr(minXval))
If t > -10 And t < 178 Then minXval = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:
Exit Sub
End Sub

Private Sub mSizek2_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the size of the reflexions markers (in counts)", prog_name & " - tick scale", sForFormat(TickMarkAlpha2Size, "F4.1"))
If t > -10 And t < 178 Then TickMarkAlpha2Size = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:

Exit Sub

End Sub

Private Sub mTickMarkerSize_Click()
On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the size of the reflexions markers (in counts)", prog_name & " - tick scale", CStr(CInt(TickMarkSize)))
If t > 0 Then TickMarkSize = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:

Exit Sub
End Sub

Private Sub mTickShowHKL_Click()
mTickShowHKL.Checked = Not (mTickShowHKL.Checked)
If mTickShowHKL.Checked Then Call refreshRietveldGraph
Exit Sub
End Sub

Private Sub mTickShowIntensity_Click()
mTickShowIntensity.Checked = Not (mTickShowIntensity.Checked)
End Sub

Private Sub mVerticalShift_Click()

On Error GoTo errorTRAP
Dim t As Single
t = InputBox("Input the vertical shift for each phase  (in percent of the Ka1 size)", prog_name & " - tick scale", 100)
If t >= 0 And t < 500 Then TickPhaseShift = t
Call refreshRietveldGraph
Exit Sub
errorTRAP:

Exit Sub


End Sub

Private Sub PE_GraphAnnotHotSpot(DblClk As Integer, nIndex As Integer)
Dim t As Integer, i As Single
'If (DblClk = 0) Then Exit Sub

For j = 1 To nPhas
k = 0
If j > 1 Then k = DBW(j - 1).nreflex
For i = 1 To DBW(j).nreflex
dk = 2 * (i + k) - 1
If dk = nIndex Or dk + 1 = nIndex Then MsgBox " phase " & CStr(j) & ", hkl: " & CStr(DBW(j).nH(i)) & " " & CStr(DBW(j).nK(i)) & " " & CStr(DBW(j).nL(i)): Exit Sub
Next i
Next j

Exit Sub

End Sub

Private Sub PE_ZoomIn()
Call refreshRietveldGraph
Exit Sub
End Sub

Private Sub PE_ZoomOut()
Call refreshRietveldGraph
Exit Sub
End Sub

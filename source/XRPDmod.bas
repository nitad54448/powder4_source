Attribute VB_Name = "Module1"

'///////////////////////////////////////////////////////////////////
'// 32 BIT PEGRAP32 API FUNCTIONS AND DEFINES //
'// For VB4/32 bit                            //
'// Copyright (c) 1996 Gigasoft, Inc.         //
'///////////////////////////////////////////////
Global Const PESTA_CENTER = 0
Global Const PESTA_LEFT = 1
Global Const PESTA_RIGHT = 2
Global Const PEDO_DRIVERDEFAULT = 0
Global Const PEDO_LANDSCAPE = 1
Global Const PEDO_PORTRAIT = 2
Global Const PEVS_COLOR = 0
Global Const PEVS_MONO = 1
Global Const PEVS_MONOWITHSYMBOLS = 2
Global Const PEFS_LARGE = 0
Global Const PEFS_MEDIUM = 1
Global Const PEFS_SMALL = 2

Global Const PEVB_NONE = 0
Global Const PEVB_TOP = 1
Global Const PEVB_BOTTOM = 2
Global Const PEVB_TOPANDBOTTOM = 3

Global Const PEAC_AUTO = 0
Global Const PEAC_NORMAL = 1
Global Const PEAC_LOG = 2

Global Const PEGLC_BOTH = 0
Global Const PEGLC_YAXIS = 1
Global Const PEGLC_XAXIS = 2
Global Const PEGLC_NONE = 3

Global Const PEAS_SUMPP = 51
Global Const PEAS_MINAP = 1
Global Const PEAS_MINPP = 52
Global Const PEAS_MAXAP = 2
Global Const PEAS_MAXPP = 53
Global Const PEAS_AVGAP = 3
Global Const PEAS_AVGPP = 54
Global Const PEAS_P1SDAP = 4
Global Const PEAS_P1SDPP = 55
Global Const PEAS_P2SDAP = 5
Global Const PEAS_P2SDPP = 56
Global Const PEAS_P3SDAP = 6
Global Const PEAS_P3SDPP = 57
Global Const PEAS_M1SDAP = 7
Global Const PEAS_M1SDPP = 58
Global Const PEAS_M2SDAP = 8
Global Const PEAS_M2SDPP = 59
Global Const PEAS_M3SDAP = 9
Global Const PEAS_M3SDPP = 60
Global Const PEAS_PARETO_ASC = 90
Global Const PEAS_PARETO_DEC = 91

Global Const PEPTGI_FIRSTPOINTS = 0
Global Const PEPTGI_LASTPOINTS = 1

Global Const PEPTGV_SEQUENTIAL = 0
Global Const PEPTGV_RANDOM = 1

Global Const PEGPT_GRAPH = 0
Global Const PEGPT_TABLE = 1
Global Const PEGPT_BOTH = 2

Global Const PETW_GRAPHED = 0
Global Const PETW_ALLSUBSETS = 1

Global Const PEDLT_PERCENTAGE = 0
Global Const PEDLT_VALUE = 1

Global Const PEMSC_NONE = 0
Global Const PEMSC_MIN = 1
Global Const PEMSC_MAX = 2
Global Const PEMSC_MINMAX = 3

Global Const PEHS_NONE = 0
Global Const PEHS_SUBSET = 1
Global Const PEHS_POINT = 2
Global Const PEHS_GRAPH = 3
Global Const PEHS_TABLE = 4
Global Const PEHS_DATAPOINT = 5
Global Const PEHS_ANNOTATION = 6
Global Const PEHS_XAXISANNOTATION = 7
Global Const PEHS_YAXISANNOTATION = 8
Global Const PEHS_HORZLINEANNOTATION = 9
Global Const PEHS_VERTLINEANNOTATION = 10

Global Const PESPM_NONE = 0
Global Const PESPM_HIGHLOWBAR = 1
Global Const PESPM_HIGHLOWLINE = 2
Global Const PESPM_HIGHLOWCLOSE = 3
Global Const PESPM_OPENHIGHLOWCLOSE = 4
Global Const PESPM_BOXPLOT = 5

Global Const PELT_THINSOLID = 0
Global Const PELT_DASH = 1
Global Const PELT_DOT = 2
Global Const PELT_DASHDOT = 3
Global Const PELT_DASHDOTDOT = 4
Global Const PELT_MEDIUMSOLID = 5
Global Const PELT_THICKSOLID = 6

Global Const PEPS_SMALL = 0
Global Const PEPS_MEDIUM = 1
Global Const PEPS_LARGE = 2
Global Const PEPS_MICRO = 3

Global Const PEPT_PLUS = 0
Global Const PEPT_CROSS = 1
Global Const PEPT_DOT = 2
Global Const PEPT_DOTSOLID = 3
Global Const PEPT_SQUARE = 4
Global Const PEPT_SQUARESOLID = 5
Global Const PEPT_DIAMOND = 6
Global Const PEPT_DIAMONDSOLID = 7
Global Const PEPT_UPTRIANGLE = 8
Global Const PEPT_UPTRIANGLESOLID = 9
Global Const PEPT_DOWNTRIANGLE = 10
Global Const PEPT_DOWNTRIANGLESOLID = 11

Global Const PEADL_NONE = 0
Global Const PEADL_DATAVALUES = 1
Global Const PEADL_POINTLABELS = 2
Global Const PEADL_DATAPOINTLABELS = 3

Global Const PEAZ_NONE = 0
Global Const PEAZ_HORIZONTAL = 1
Global Const PEAZ_VERTICAL = 2
Global Const PEAZ_HORZANDVERT = 3

Global Const PEBFD_2ND = 0
Global Const PEBFD_3RD = 1
Global Const PEBFD_4TH = 2

Global Const PEBS_SMALL = 0
Global Const PEBS_MEDIUM = 1
Global Const PEBS_LARGE = 2

Global Const PECG_COARSE = 0
Global Const PECG_MEDIUM = 1
Global Const PECG_FINE = 2

Global Const PEAE_NONE = 0
Global Const PEAE_ALLSUBSETS = 1
Global Const PEAE_INDSUBSETS = 2

Global Const PECM_NOCURSOR = 0
Global Const PECM_POINT = 1
Global Const PECM_DATACROSS = 2
Global Const PECM_DATASQUARE = 3

Global Const PEGAT_NOSYMBOL = 0
Global Const PEGAT_PLUS = 1
Global Const PEGAT_CROSS = 2
Global Const PEGAT_DOT = 3
Global Const PEGAT_DOTSOLID = 4
Global Const PEGAT_SQUARE = 5
Global Const PEGAT_SQUARESOLID = 6
Global Const PEGAT_DIAMOND = 7
Global Const PEGAT_DIAMONDSOLID = 8
Global Const PEGAT_UPTRIANGLE = 9
Global Const PEGAT_UPTRIANGLESOLID = 10
Global Const PEGAT_DOWNTRIANGLE = 11
Global Const PEGAT_DOWNTRIANGLESOLID = 12
Global Const PEGAT_SMALLPLUS = 13
Global Const PEGAT_SMALLCROSS = 14
Global Const PEGAT_SMALLDOT = 15
Global Const PEGAT_SMALLDOTSOLID = 16
Global Const PEGAT_SMALLSQUARE = 17
Global Const PEGAT_SMALLSQUARESOLID = 18
Global Const PEGAT_SMALLDIAMOND = 19
Global Const PEGAT_SMALLDIAMONDSOLID = 20
Global Const PEGAT_SMALLUPTRIANGLE = 21
Global Const PEGAT_SMALLUPTRIANGLESOLID = 22
Global Const PEGAT_SMALLDOWNTRIANGLE = 23
Global Const PEGAT_SMALLDOWNTRIANGLESOLID = 24
Global Const PEGAT_LARGEPLUS = 25
Global Const PEGAT_LARGECROSS = 26
Global Const PEGAT_LARGEDOT = 27
Global Const PEGAT_LARGEDOTSOLID = 28
Global Const PEGAT_LARGESQUARE = 29
Global Const PEGAT_LARGESQUARESOLID = 30
Global Const PEGAT_LARGEDIAMOND = 31
Global Const PEGAT_LARGEDIAMONDSOLID = 32
Global Const PEGAT_LARGEUPTRIANGLE = 33
Global Const PEGAT_LARGEUPTRIANGLESOLID = 34
Global Const PEGAT_LARGEDOWNTRIANGLE = 35
Global Const PEGAT_LARGEDOWNTRIANGLESOLID = 36

Global Const PEGAT_POINTER = 37

Global Const PEGAT_THINSOLIDLINE = 38
Global Const PEGAT_DASHLINE = 39
Global Const PEGAT_DOTLINE = 40
Global Const PEGAT_DASHDOTLINE = 41
Global Const PEGAT_DASHDOTDOTLINE = 42
Global Const PEGAT_MEDIUMSOLIDLINE = 43
Global Const PEGAT_THICKSOLIDLINE = 44
Global Const PEGAT_LINECONTINUE = 45
                                     
Global Const PEGAT_TOPLEFT = 46
Global Const PEGAT_BOTTOMRIGHT = 47

Global Const PEGAT_RECT_THIN = 48
Global Const PEGAT_RECT_DASH = 49
Global Const PEGAT_RECT_DOT = 50
Global Const PEGAT_RECT_DASHDOT = 51
Global Const PEGAT_RECT_DASHDOTDOT = 52
Global Const PEGAT_RECT_MEDIUM = 53
Global Const PEGAT_RECT_THICK = 54
Global Const PEGAT_RECT_FILL = 55

Global Const PEGAT_ROUNDRECT_THIN = 56
Global Const PEGAT_ROUNDRECT_DASH = 57
Global Const PEGAT_ROUNDRECT_DOT = 58
Global Const PEGAT_ROUNDRECT_DASHDOT = 59
Global Const PEGAT_ROUNDRECT_DASHDOTDOT = 60
Global Const PEGAT_ROUNDRECT_MEDIUM = 61
Global Const PEGAT_ROUNDRECT_THICK = 62
Global Const PEGAT_ROUNDRECT_FILL = 63

Global Const PEGAT_ELLIPSE_THIN = 64
Global Const PEGAT_ELLIPSE_DASH = 65
Global Const PEGAT_ELLIPSE_DOT = 66
Global Const PEGAT_ELLIPSE_DASHDOT = 67
Global Const PEGAT_ELLIPSE_DASHDOTDOT = 68
Global Const PEGAT_ELLIPSE_MEDIUM = 69
Global Const PEGAT_ELLIPSE_THICK = 70
Global Const PEGAT_ELLIPSE_FILL = 71

Global Const PEDTM_NONE = 0
Global Const PEDTM_VB = 1
Global Const PEDTM_DELPHI = 2

Global Const PESA_ALL = 0
Global Const PESA_AXISLABELS = 1
Global Const PESA_GRIDNUMBERS = 2
Global Const PESA_NONE = 3

Global Const PEGS_THIN = 0
Global Const PEGS_THICK = 1
Global Const PEGS_DOT = 2
Global Const PEGS_DASH = 3

Global Const PEFVP_AUTO = 0
Global Const PEFVP_VERT = 1
Global Const PEFVP_HORZ = 2

Global Const PEMAS_NONE = 0
Global Const PEMAS_THIN = 1
Global Const PEMAS_MEDIUM = 2
Global Const PEMAS_THICK = 3
Global Const PEMAS_THICKPLUSTICK = 4

Type Rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Type POINTSTRUCT
    X As Long
    Y As Long
End Type

Type HOTSPOTDATA
    HotSpotL As Long
    HotSpotT As Long
    HotSpotR As Long
    HotSpotB As Long
    nHotSpotType As Long
    n1 As Long
    n2 As Long
End Type

Type GLOBALPROPERTIES
    nObjectType As Long
    szMainTitle As String * 48
    szSubTitle As String * 48
    nSubsets As Long
    npoints As Long
    bMonoWithSymbols As Long
    nDefOrientation As Long
    nPrepareImages As Long
    b3dDialogs As Long
    bDataShadows As Long
    bAllowCustomization As Long
    bAllowExporting As Long
    bAllowMaximization As Long
    bAllowPopup As Long
    nPageWidth As Long
    nPageHeight As Long
    rectLogicalLoc As Rect
    bCustom As Long
    nViewingStyle As Long
    nCViewingStyle As Long
    dwMonoDeskColor As Long
    dwMonoTextColor As Long
    dwMonoShadowColor As Long
    dwMonoGraphForeColor As Long
    dwMonoGraphBackColor As Long
    dwMonoTableForeColor As Long
    dwMonoTableBackColor As Long
    dwCMonoDeskColor As Long
    dwCMonoTextColor As Long
    dwCMonoShadowColor As Long
    dwCMonoGraphForeColor As Long
    dwCMonoGraphBackColor As Long
    dwCMonoTableForeColor As Long
    dwCMonoTableBackColor As Long
    dwDeskColor As Long
    dwTextColor As Long
    dwShadowColor As Long
    dwGraphForeColor As Long
    dwGraphBackColor As Long
    dwTableForeColor As Long
    dwTableBackColor As Long
    dwCDeskColor As Long
    dwCTextColor As Long
    dwCShadowColor As Long
    dwCGraphForeColor As Long
    dwCGraphBackColor As Long
    dwCTableForeColor As Long
    dwCTableBackColor As Long
    nDataPrecision As Long
    nCDataPrecision As Long
    nFontSize As Long
    nCFontSize As Long
    szMainTitleFont As String * 48
    bMainTitleBold As Long
    bMainTitleItalic As Long
    bMainTitleUnderline As Long
    szCMainTitleFont As String * 48
    bCMainTitleBold As Long
    bCMainTitleItalic As Long
    bCMainTitleUnderline As Long
    szSubTitleFont As String * 48
    bSubTitleBold As Long
    bSubTitleItalic As Long
    bSubTitleUnderline As Long
    szCSubTitleFont As String * 48
    bCSubTitleBold As Long
    bCSubTitleItalic As Long
    bCSubTitleUnderline As Long
    szLabelFont As String * 48
    bLabelBold As Long
    bLabelItalic As Long
    bLabelUnderline As Long
    szCLabelFont As String * 48
    bCLabelBold As Long
    bCLabelItalic As Long
    bCLabelUnderline As Long
    szTableFont As String * 48
    szCTableFont As String * 48
    bAllowSubsetHotSpots As Long
    bAllowPointHotSpots As Long
End Type


'Scientific Graph Plotting Methods for PESGO32.OCX
Global Const SGPM_LINE = 0
Global Const SGPM_POINT = 1
Global Const SGPM_STICK = 2
Global Const SGPM_POINTSPLUSBFL = 3
Global Const SGPM_POINTSPLUSBFC = 4
Global Const SGPM_POINTSPLUSSPLINE = 5
Global Const SGPM_SPLINE = 6
Global Const SGPM_BUBBLE = 7
Global Const SGPM_POINTSPLUSLINE = 8
Global Const SGPM_AREA = 9
Global Const SGPM_BAR = 10
Global Const SGPM_SPECIFICPLOTMODE = 11


'// MOST COMMON CONSTANTS USED WITH VB //'
'// IF OTHERS ARE NEEDED, FIND THEM IN PEGRPAPI.TXT //'
Global Const PEP_nSUBSETS = 2115
Global Const PEP_nPOINTS = 2120
Global Const PEP_szaSUBSETLABELS = 2125
Global Const PEP_szaPOINTLABELS = 2130
Global Const PEP_faXDATA = 2135
Global Const PEP_faYDATA = 2140
Global Const PEP_bCUSTOM = 2225
Global Const PEP_faAPPENDYDATA = 3276
Global Const PEP_szaAPPENDPOINTLABELDATA = 3277

'////// API FUNCTIONS //////'
Declare Function PEsetglobal Lib "PEGRAP32.DLL" (ByVal hObject&, lpData As GLOBALPROPERTIES) As Long
Declare Function PEgetglobal Lib "PEGRAP32.DLL" (ByVal hObject&, lpData As GLOBALPROPERTIES) As Long
Declare Function PEvset Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, lpvData As Any, ByVal nItems&) As Long
Declare Function PEvget Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, lpvDest As Any) As Long
Declare Function PEvsetcell Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nCell&, lpvData As Any) As Long
Declare Function PEvgetcell Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nCell&, lpvDest As Any) As Long
Declare Function PEszset Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, ByVal szData$) As Long
Declare Function PEszget Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, ByVal szData$) As Long
Declare Function PEnset Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nData&) As Long
Declare Function PEnget Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&) As Long
Declare Function PElset Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nData&) As Long
Declare Function PElget Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nProperty&) As Long
Declare Function PEcreate Lib "PEGRAP32.DLL" (ByVal nObjectType&, ByVal dwStyle&, lpRect As Rect, ByVal hParent&, ByVal nID&) As Long
Declare Function PEdestroy Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PEload Lib "PEGRAP32.DLL" (ByVal hObject&, lphGlbl As Any) As Long
Declare Function PEstore Lib "PEGRAP32.DLL" (ByVal hObject&, lphGlbl As Any, lpdwSize As Any) As Long
Declare Function PEloadpartial Lib "PEGRAP32.DLL" (ByVal hObject&, lphGlbl As Any) As Long
Declare Function PEstorepartial Lib "PEGRAP32.DLL" (ByVal hObject&, lphGlbl As Any, lpdwSize As Any) As Long
Declare Function PEgetmeta Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PEresetimage Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal nLength&, ByVal nHeight&) As Long
Declare Function PElaunchcustomize Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PElaunchexport Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PElaunchmaximize Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PElaunchtextexport Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal bToFile&, ByVal lpszFilename$) As Long
Declare Function PElaunchprintdialog Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal bFullPage&, lpPoint As POINTSTRUCT) As Long
Declare Function PElaunchcolordialog Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PElaunchfontdialog Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PElaunchpopupmenu Lib "PEGRAP32.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEreinitialize Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PEreinitializecustoms Lib "PEGRAP32.DLL" (ByVal hObject&) As Long
Declare Function PEgethelpcontext Lib "PEGRAP32.DLL" (ByVal hwnd&) As Long
Declare Function PEcopymetatoclipboard Lib "PEGRAP32.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEcopymetatofile Lib "PEGRAP32.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long
Declare Function PEcopybitmaptoclipboard Lib "PEGRAP32.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEcopybitmaptofile Lib "PEGRAP32.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long
Declare Function PEcopyoletoclipboard Lib "PEGRAP32.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
Declare Function PEprintgraph Lib "PEGRAP32.DLL" (ByVal hObject&, ByVal hDC&, ByVal nWidth&, ByVal nHeight&, ByVal nOrient&) As Long
Declare Function PEconvpixeltograph Lib "PEGRAP32.DLL" (ByVal hObject&, ByRef nAxis&, ByRef nX&, ByRef nY&, ByRef fX#, ByRef fY#, ByVal bRight&, ByVal bTop&, ByVal bVV&) As Long
Declare Function PEreset Lib "PEGRAP32.DLL" (ByVal hObject&) As Long

Global AtomEdited As Integer
Global phaseEdited As Integer

''----------
Option Explicit
Option Base 1
Global CallFromGraphForCell As Boolean

Global AuthorWebPage As String
'Global DicvolLocationDirectory As String
'Global ItoLocationDirectory As String
'Global TreorLocationDirectory As String
Global WorkingDir As String
Global Const prog_name = "Powder v4.0"
Global Const version_name = "version 0.7d - 05 mar 03 "
'v 07c length of the stepx in gsas esd format adjustet to fit for Xcell step
'v 07b correction of the ESD file format
'v 07 - the X and step format modified to allow for Xcellerator strange steps
'v 06d - 3500 reflexions permitted

Global Const strLinie = "-------------------------------------------------"
Global Const rd As Double = 180# / 3.14159265359
Global ignoralinii As Integer
Global X() As Double, Y() As Double, z() As Double
Global title As String, startx As Double, endx As Double, stepx As Double
Global inputfile As String, outputfile As String, numarvalori As Long ''Integer
Global amfullxdata As Boolean, amxdata As Boolean, amydata As Boolean, amzdata As Boolean
Global expName As String
Global dbwsDataFile As String
Global dbwsControlFile As String
Global dbwsOutputFile As String
Global nPointSize As Integer, MaxXScale As Single, minXScale As Single
Global maxYScale As Single, maxRYScale As Single, minYScale As Single, minRYScale As Single
Global colorSubset1 As OLE_COLOR, colorSubset2 As OLE_COLOR, colorSubset3 As OLE_COLOR
Global nPointSubset1Color As OLE_COLOR, nPointSubset2Color As OLE_COLOR, npointSubset3Color As OLE_COLOR
Global MinYBottomScale As Single, MaxYBottomScale As Single
Global dbX() As Single, dbYraw() As Single, dbYsum() As Single, dbYdiff() As Single
Global nPhas As Integer, sTitle As String, maxval As Single, nval As Integer
Global minXval As Single, maxXval As Single, minYIcalc As Single, maxYIcalc As Single
Global minYIobs As Single, maxYIobs As Single, minYIdiff As Single, maxYIdiff As Single
Global TickMarkShift As Single, TickMarkSize As Single, TickMarkAlpha2Size As Single
Global TickPhaseShift As Single
Global CustomFontsizeglobalcntl As Single
Global CustomFontsizelegendcntl As Single
'a user defined datatype, called PhaseInfo
Type PhaseInfo
    Name As String * 72
    Atomi As Integer
    FormulaUnits As Integer
    ParticleAbsorptionFactor As Single
    PrefOrientation(3) As Single
    SpaceGroupSymbol As String
    scalefactor As Single
    OverallThermal As Single
    scaleFactorCode As Single
    OverallThermalCode As Single
    U As Single
    v As Single
    W As Single
    CT As Single
    z As Single
    X As Single
    Y As Single
    codeU As Single
    codeV As Single
    codeW As Single
    codeCT As Single
    codeZ As Single
    codeX As Single
    codeY As Single
    a As Single
    b As Single
    c As Single
    Alpha As Single
    Beta As Single
    gamma As Single
    codeA As Single
    codeB As Single
    codeC As Single
    codeAlpha As Single
    codeBeta As Single
    codegamma As Single
    G1 As Single
    G2 As Single
    P As Single
    codeG1 As Single
    codeG2 As Single
    codeP As Single
    NA As Single
    NB As Single
    NC As Single
    SP7A  As Single
    hNA As Single
    hNB As Single
    hNC As Single
    codeNA As Single
    codeNB As Single
    codeNC As Single
    codeSP7A  As Single
    codehNA As Single
    codehNB As Single
    codehNC As Single
End Type


Global Const nreflex = 3500
Global minDiff As Single, maxDiff As Single
Type DBWSPlotinfo
    sTitle As String * 70
    nreflex As Integer
    sDoiTeta(nreflex) As Single
    bK1(nreflex) As Boolean
    bK2(nreflex) As Boolean
    nH(nreflex) As Integer
    nK(nreflex) As Integer
    nL(nreflex) As Integer
    sIntens(nreflex) As Single
End Type

Global DBW() As DBWSPlotinfo

Type AtomInfo
    Label As String * 4
    Multiplicity As Single
    Ntyp As String * 4
    X As Single
    Y As Single
    z As Single
    IsotropicThermal As Single
    SiteOccupancy As Single
    codeX As Single
    codeY As Single
    codeZ As Single
    codeIsotropicThermal As Single
    codeSiteOccupancy As Single
    Beta11 As Single
    Beta22 As Single
    Beta33 As Single
    Beta12 As Single
    Beta13 As Single
    Beta23 As Single
    codeBeta11 As Single
    codeBeta22 As Single
    codeBeta33 As Single
    codeBeta12 As Single
    codeBeta13 As Single
    codeBeta23 As Single
 End Type
Global Atoms(200, 15) As AtomInfo 'arbitrary used limit for 15 phases
Global Phases(15) As PhaseInfo ' 15 is restricted by dbws

Type ScatteringFactors
    IntTable As Boolean
    Name As String
    RePart As Single
    ImPart As Single
    AtWeight As Single
    NineCoeff(9) As Single
    PosScatt(2, 100) As Single
End Type

Global Scattering(20) As ScatteringFactors
Global totalScatt As Integer

Public Type crystallo
    h As Integer
    k As Integer
    l As Integer
    doitheta As Single
    d As Single
End Type

Public Type BaseLine
Hmax As Integer
Kmax As Integer
Lmax As Integer
HKLmax As Integer
End Type


Public Type SimulatedPeaks
    h As Integer
    k As Integer
    l As Integer
    doitheta As Single
    ygrec As Single
End Type

Global valori(1000) As SimulatedPeaks
Global amsentpeaks As Boolean
Global NumberOfSimulatedPeaks As Integer


Sub Main()
Dim i As Integer
    'frmSplash.Show
    'frmSplash.Refresh
    DoEvents
    Load Convert3Main
    Convert3Main.Show
    DoEvents

   Convert3Main.Enabled = True
    '
    
    ''Unload frmSplash
End Sub




Sub RietveldBoardMessage(mesaj As String)
Dim outfil As Integer
DoEvents
On Error GoTo errorTRAP
frmgDBWSmain.txtMainMessage.Text = frmgDBWSmain.txtMainMessage.Text & mesaj & vbCrLf
Exit Sub
errorTRAP:
'more than 32k message ??
outfil = FreeFile
Open "Rietveld_Board_message.dat" For Append As outfil '
Print #outfil, frmgDBWSmain.txtMainMessage.Text
Close outfil
frmgDBWSmain.txtMainMessage.Text = Now & vbCrLf & "An error has occured. Report appended to Rietveld_board_message.dat"
Err.Clear
Exit Sub

End Sub




Sub IamBusy(ByRef t As Boolean)
Screen.MousePointer = 0
If t Then Screen.MousePointer = 11
DoEvents
Exit Sub
End Sub




Sub fct(N As Integer, nrpuncte As Integer, ind() As Integer, h() As Integer, k() As Integer, l() As Integer, theta() As Double, pondere() As Double, f As Double, r() As Double, X() As Double, g() As Double)
''SHARED nef, npar, ndet, dmdet(), fit(), par(), w(), e(), par$(), dmax, blockG
On Error GoTo handleit
Dim ae As Double, be As Double, ce As Double, cae As Double, cbe As Double, cce As Double
Dim i As Integer, ix As Double, j As Integer, q As Integer
f = 0
ix = 0
'se defineste f=suma de patrate
' se pun: x(i)=parametri
'si se definesc derivatele:  g(i)=d f/ d x(i)
    
    q = 0
    For i = 1 To 8
    If ind(i) = 1 Then q = q + 1: r(i) = X(q)
     Next i
   
    
      ae = r(1)
      be = r(2)
      ce = r(3)
      cae = Cos(r(4) / rd)
      cbe = Cos(r(5) / rd)
      cce = Cos(r(6) / rd)
ReDim g(8)
q = 0
For j = 1 To 8

If ind(j) = 1 Then

q = q + 1
For i = 1 To nrpuncte
g(q) = g(q) + f_derivata(0.00001, r, h(i), k(i), l(i), theta(i), j, pondere(i))
Next i
End If

Next j



For i = 1 To nrpuncte
f = f + pondere(i) * ((ae * h(i)) ^ 2 + (be * k(i)) ^ 2 + (ce * l(i)) ^ 2 + 2# * (h(i) * k(i) * ae * be * cce + k(i) * l(i) * be * ce * cae + l(i) * h(i) * ce * ae * cbe) - (Sin((theta(i) - r(8) / 2) / rd) * 2 / r(7)) ^ 2)
Next i
f = f * f

For j = 1 To N
g(j) = 2 * f * g(j)
Next j
'ReDim q(nrval, 8) As Double
     '' If (indic = 1 Or indic = 2) Then b(4) = b(3)
'ReDim Preserve x(n)
''ReDim Preserve g(n)
Exit Sub
handleit:
raport "error in the function evaluation-conjugated gradients"
Err.Clear
Exit Sub
End Sub

Sub deriv(N As Single, dx() As Double, dy() As Double, ddy() As Double, eroare As Boolean)
'calculeaza derivata numerica prin derivata polinomului de interpolare
'este posibil ca eroarea sa fie mare
On Error GoTo localhandle
Dim i As Single
''n is single to permit more than 32000 points
'ddy is the derivative
'dx and dy are the points; these names are used to avoid the conflict with the global
'variables x and y (theoretically it shouldn t be any conflict, but...)
ddy(1) = 0
ddy(N) = 0
For i = 2 To N - 1
ddy(i) = dy(i - 1) * (dx(i) - dx(i + 1)) / (dx(i - 1) - dx(i)) / (dx(i - 1) - dx(i + 1)) + Y(i) * (2 * dx(i) - dx(i - 1) - dx(i + 1)) / (dx(i) - dx(i - 1)) / (dx(i) - dx(i + 1)) + Y(i + 1) * (dx(i) - dx(i - 1)) / (dx(i + 1) - dx(i - 1)) / (dx(i + 1) - dx(i))
Next i
'i = n - 1
'ddy(i) = y(i - 1) * (dx(i) - dx(i + 1)) / (dx(i - 1) - dx(i)) / (dx(i - 1) - dx(i + 1)) + y(i) * (2 * dx(i) - dx(i - 1) - dx(i + 1)) / (dx(i) - dx(i - 1)) / (dx(i) - dx(i + 1)) + y(i + 1) * (dx(i) - dx(i - 1)) / (dx(i + 1) - dx(i - 1)) / (dx(i + 1) - dx(i))
Exit Sub
localhandle:
eroare = True
Exit Sub
End Sub










Sub minfp(N As Integer, nrpuncte As Integer, f As Double, r() As Double, X() As Double, eps As Double, est As Double, limit As Integer, ier As Integer, nef As Integer, ind() As Integer, hh() As Integer, KK() As Integer, ll() As Integer, theta() As Double, pondere() As Double)
 Rem subrutina Flechter -Powell
 On Error GoTo handleit
 Dim m As Integer, h() As Double, g() As Double, nr As Integer, n2 As Integer, n3 As Integer, n31 As Integer
 Dim xoptim() As Double, goptim() As Double, foptim As Double, t As Double, dy As Double, hn As Double, gn As Double
 Dim k As Integer, j As Integer, nj As Integer, l As Integer, k1 As Integer, fv As Double, fY As Double, pas As Double, alfa As Double
 Dim fX As Double, dx As Double, i As Integer, z As Double, q As Double, dpas As Double, dxl As Double
 f = 0
 m = N * (N + 7) / 2
 ReDim h(m), g(N)
 foptim = 1E+20
 ''ReDim xoptim(n), goptim(n)
 nef = 0
 Call fct(N, nrpuncte, ind, hh, KK, ll, theta, pondere, f, r, X, g)
 nef = nef + 1
 nr = 0
 ier = 0
 n2 = 2 * N
 n3 = 3 * N
 n31 = n3 + 1
fp20:
 k = n31
 For j = 1 To N
 h(k) = 1
 nj = N - j
 If nj <= 0 Then GoTo fp1:
 For l = 1 To nj
 k1 = k + l
 h(k1) = 0
 Next l
 k = k1 + 1
 Next j
fp1:
 nr = nr + 1
 fv = f
 For j = 1 To N
 k = N + j
 h(k) = g(j)
 k = k + N
 h(k) = X(j)
 k = j + n3
 t = 0
 For l = 1 To N
 t = t - g(l) * h(k)
 If l >= j Then GoTo fp2:
 k = k + N - l
 GoTo fp4
fp2:
 k = k + 1
fp4:
 Next l
 h(j) = t
 Next j
 dy = 0: hn = 0: gn = 0
 For j = 1 To N
 hn = hn + Abs(h(j))
 gn = gn + Abs(g(j))
 dy = dy + h(j) * g(j)
 Next j
 If dy >= 0 Then GoTo fp3:
 If hn / gn <= eps Then GoTo fp3:
 fY = f
 pas = 2 * (est - f) / dy
 alfa = 1
 If pas <= 0 Then GoTo fp5:
 If pas >= alfa Then GoTo fp5:
 alfa = pas
fp5:
 pas = 0
fp10:
 fX = fY
 dx = dy
 For i = 1 To N
 X(i) = X(i) + alfa * h(i)
 Next i
 Call fct(N, nrpuncte, ind, hh, KK, ll, theta, pondere, f, r, X, g)
 nef = nef + 1
 If nef >= limit Then GoTo fp6:
 fY = f: dy = 0
 For i = 1 To N
 dy = dy + g(i) * h(i)
 Next i
 If dy > 0 Then GoTo fp7:
 If dy = 0 Then GoTo fp8:
 ''If dy < 0 Then GoTo fp9:
fp9:
 If fY >= fX Then GoTo fp7:
 alfa = alfa + pas
 pas = alfa
 If hn * alfa <= 1E+20 Then GoTo fp10:
 ier = 2
 GoTo fp17
fp7:
 t = 0
fp25:
 If alfa = 0 Then GoTo fp8:
 z = 3 * (fX - fY) / alfa + dx + dy
 If z * z < dx * dy Then GoTo fp3:
 q = Sqr(z * z - dx * dy)
 pas = (dy + q - z) * alfa / (dy - dx + 2 * q)
 For i = 1 To N
 X(i) = X(i) + (t - pas) * h(i)
 Next i
 ''Call fct(n, f, x(), g())
  Call fct(N, nrpuncte, ind, hh, KK, ll, theta, pondere, f, r, X, g)

 nef = nef + 1
 If nef >= limit Then GoTo fp6:
 If f > fX Then GoTo fp11:
 If f <= fY Then GoTo fp8:
fp11:
 dpas = 0
 For i = 1 To N
 dpas = dpas + g(i) * h(i)
 Next i
 If dpas >= 0 Then GoTo fp12:
 If f > fX Then GoTo fp12:
 If f = fX Then GoTo fp13:
 If f < fX Then GoTo fp14:
fp13:
 If dxl = dpas Then GoTo fp8:
fp14:
 fX = f
 dx = dpas
 t = pas
 alfa = pas
  GoTo fp25
fp12:
 If fY <> f Then GoTo fp15:
 If dy = dpas Then GoTo fp8:
fp15:
 fY = f
 dy = dpas
 alfa = alfa - pas
 GoTo fp7:
fp8:
 For j = 1 To N
 k = N + j
 h(k) = g(j) - h(k)
 k = k + N
 h(k) = X(j) - h(k)
 Next j
 If (fv + eps) < f Then GoTo fp3:
 ier = 0
 If nr < N Then GoTo fp16:
 t = 0
 z = 0
 For j = 1 To N
 k = N + j: q = h(k): k = k + N: t = t + Abs(h(k))
 z = z + q * h(k)
 Next j
 If hn > eps Then GoTo fp16:
 If (t <= eps) Then GoTo fp17:
fp16:
 pas = 0
 For j = 1 To N
 k = n3 + j
 q = 0
 For l = 1 To N
 k1 = N + l
 q = q + h(k1) * h(k)
 If l >= j Then GoTo fp18:
 k = k + N - l
 GoTo fp19:
fp18:
 k = k + 1
fp19:
 Next l
 k = N + j
 pas = pas + q * h(k)
 h(j) = q
 Next j
 If z * pas = 0 Then GoTo fp20:
 k = n31
 For l = 1 To N
 k1 = n2 + l
 For j = l To N
 nj = n2 + j
 h(k) = h(k) + h(k1) * h(nj) / z - h(l) * h(j) / pas
 k = k + 1
 Next j
 Next l
 GoTo fp1:
fp6:
 ier = 1
 GoTo fp17
fp3:
 For j = 1 To N
 k = n2 + j
 X(j) = h(k)
 Next j
  Call fct(N, nrpuncte, ind, hh, KK, ll, theta, pondere, f, r, X, g)
 nef = nef + 1
 If nef >= limit Then GoTo fp6:
 If gn <= eps Then GoTo fp21:
 If ier = 3 Then GoTo fp17:
 ier = 3
 GoTo fp20:
fp21:
 ier = 0
fp17:
''f = foptim
''For i = 1 To n
''x(i) = xoptim(i)
''g(i) = goptim(i)
''Next i
''Erase h, g
''Erase xoptim, goptim
Exit Sub
handleit:
raport "Error in Davidon-Fletcher routine..."
Err.Clear
Exit Sub
End Sub


Function doitheta_deg(res() As Double, lam As Double, h As Integer, k As Integer, l As Integer) As Double
On Error GoTo handleit
doitheta_deg = lam / 2 * (Sqr(h * h * res(1) * res(1) + k * k * res(2) * res(2) + l * l * res(3) * res(3) + 2 * l * h * res(3) * res(1) * Cos(res(5) / rd) + 2 * l * k * res(3) * res(2) * Cos(res(4) / rd) + 2 * h * k * res(2) * res(1) * Cos(res(6) / rd)))
doitheta_deg = asin(doitheta_deg) * rd * 2#
Exit Function
handleit:
raport "Error in 2 theta calculation...?"
Err.Clear
Exit Function
End Function

Function asin(X As Double) As Double
On Error GoTo handleit
If X > 0.9999 And X < 1.0001 Then
asin = 90 / rd
Exit Function
End If

asin = Atn(X / Sqr(-X * X + 1))
Exit Function
handleit:
raport strLinie & vbCrLf & Now & "  <--- Fatal error in ASN function."
Err.Clear
Exit Function
End Function



Function f_derivata(delta As Double, r() As Double, h As Integer, k As Integer, l As Integer, theta As Double, j As Integer, pondere As Double)
'derivata functiei in raport cu termenul j
On Error GoTo handleit
Dim F1 As Double, r2(8) As Double, m As Integer
F1 = functie(r, pondere, h, k, l, theta)
For m = 1 To 8: r2(m) = r(m): Next m
r2(j) = r2(j) + delta
f_derivata = (functie(r2, pondere, h, k, l, theta) - F1) / delta
Exit Function
handleit:
raport "error in the derivative calculation..."
Exit Function
End Function





Sub fonc(theta() As Double, nr As Integer, npaf As Integer, b() As Double, h() As Integer, k() As Integer, l() As Integer, r As Double, rr As Double)
Dim r1 As Double, r2 As Double, ae As Double, be As Double, ce As Double, cae As Double, cbe As Double, cce As Double
Dim dd As Double, yc As Double, d As Double, npaf2 As Integer, i As Integer
''Dim qq() As Double
npaf2 = npaf + 2
ReDim qq(nr, npaf2)
      r1 = 0
      r2 = 0
      ae = b(3)
      be = b(4)
      ce = b(5)
      cae = Cos(b(6))
      cbe = Cos(b(7))
      cce = Cos(b(8))
   For i = 1 To nr
      ''DO 2 I=1,NR
      dd = (ae * h(i)) ^ 2 + (be * k(i)) ^ 2 + (ce * l(i)) ^ 2 + 2# * (h(i) * k(i) * ae * be * cce + k(i) * l(i) * be * ce * cae + l(i) * h(i) * ce * ae * cbe)
      d = 1# / Sqr(dd)
      yc = b(1) + Atn((b(2) / d) / Sqr(-(b(2) / d) * (b(2) / d) + 1))
      r1 = r1 + (yc - theta(i)) ^ 2
      r2 = r2 + theta(i) ^ 2
     qq(i, npaf2) = yc
Next i
      r = r1 / (nr - npaf)
      rr = Sqr(r1 / r2)
      Exit Sub
End Sub

Sub eracel_based(nr As Integer, indic As Integer, ifin As Integer, b() As Double, bb() As Double, sig() As Double, npaf As Integer, afi() As Integer, h() As Integer, k() As Integer, l() As Integer, theta() As Double, pds() As Double)
Dim dum(3) As Double
Dim volum As Double, i As Integer, j As Integer, npaf2 As Integer
Dim iffi As Integer, ik As Integer, jj As Integer, qq() As Double, r As Double, rr As Double
Dim a(60) As Double, imm As Single, mm As Integer, m1 As Integer, m2 As Integer
Dim no As Integer, IQ As Integer, il As Integer, ier As Integer, dmn(8) As Double
Dim suma As Double, imax As Integer, kli As Integer, kmi As Integer, m As Integer
Dim term As Double, denom As Double, ii As Integer, db(8) As Double
Dim ad(6), sInp(3) As Double, cosp(3) As Double, sp(3), ss(3), cc(3) As Double, dqd(3), am(1) As Double
Dim cabc2 As Double, q2 As Double, qmn As Double, yc As Double
Dim ae As Double, be As Double, ce As Double, cae As Double, cbe As Double, cce As Double
Dim dd As Double, d As Double, rad As Double, f As Double, r1 As Double, r2 As Double
Dim q() As Double, ir As Integer, kmn As Integer, lmn As Integer, nfail As Integer, kdm As Integer
Dim ndmax As Integer
Dim bout(8) As Double
'indic este constrainte de rafinare
'ifin este numarul de pasi
ReDim q(nr, 8) As Double, qq(nr, 10) As Double
On Error GoTo handleit
ndmax = nr + 1
If indic = 0 Or indic > 3 Then indic = 3
iffi = Fix(afi(3) + afi(4) + afi(5) + 0.1)
   
      If (iffi = 0) Or (indic = 3) Then GoTo 1230
      ik = 3 - indic
     For i = 1 To ik
 ''atentie aici de rescris!!!!!!!?????
If (indic - 2) < 0 Then afi(2 + indic + 1) = 0
If (indic - 2) >= 0 Then afi(2 + indic - 1 + i) = 0
1200
1210
1220
Next i
      afi(3) = 1#

1230
1240
1250   ''PDS(nr) = 1#  'weighting ???
      For i = 1 To nr
      theta(i) = theta(i) / rd ''!!!!!!!!!! THETA in radiani
      Next i
1260
      b(1) = b(1) / rd
      b(2) = b(2) * 0.5
      For i = 6 To 8
      b(i) = b(i) / rd
      Next i
1270


Call inver(b, dum, volum, 0)

      For i = 1 To 3
      dum(i) = b(5 + i) * rd
      Next i
1280
''    WRITE(IWR,120)(B(I),I=3,5),DUM,VOLUM
''C....."NPAF" : NOMBRE DE PARAMETRES A AFFINER
''C....."BB()" : TABLEAU DES PARAMETRES A AFFINER
      j = 0
      For i = 1 To 8
    ''  bout(i) = b(i)
      If (afi(i) = 1) Then j = j + 1: bb(j) = b(i)
1285
      Next i
1290
      npaf = j
      If (npaf = 8) Then raport "So, you want to refine all parameters in the same time ?!"
      npaf2 = npaf + 2
''      ReDim qq(npaf, nr)
    
''---------mcrl
Call mcrnl(qq, ndmax, theta, b, bb, afi, npaf, nr, pds, ifin, indic, h(), k(), l())
 ''-------------
''C.....NOUVELLES VALEURS DES PARAMETRES


      j = 0
  For i = 1 To 8
      If (afi(i) = 1) Then j = j + 1: b(i) = bb(j)
      Next i
''For i = 1 To 8
''MsgBox CStr(b(i))

''Next i



''C   ......VALEURS DES ANGLES THETA CALCULES
Call fonc(theta, nr, npaf, b, h, k, l, r, rr)
      jj = 0
      For i = 1 To 8
      If (afi(i)) > 0 Then
      jj = jj + 1
      sig(i) = Sqr(qq(jj, jj) * r)
      Else
      sig(i) = 0#

      End If
      Next i
7330
      
      If (indic = 1 Or indic = 2) Then sig(4) = sig(3)
      If (indic = 1) Then sig(5) = sig(3)
      For i = 1 To 3
      bb(i) = b(i + 2)
      bb(i + 3) = b(i + 5) * rd
      Next i
7340

Exit Sub
''ies cu valorile reciproce...si pentru sigma....



Call inver(b, sig, volum, 1)  ''am pus bb in loc de b...

''----------
      sig(1) = sig(1) * rd
      sig(2) = sig(2) * 2#
''c
''C.....SORTIE DES RESULTATS
      volum = 1# / volum
      b(1) = b(1) * rd
      b(2) = b(2) * 2#
      For i = 6 To 8
      sig(i) = sig(i) * rd
   b(i) = b(i) * rd
''raport "param " & CStr(i) & CStr(b(i))
Next i
Exit Sub
handleit:
raport "Error in least squares general routine..."
Err.Clear
Exit Sub
End Sub


Function functie(res() As Double, pondere As Double, h As Integer, k As Integer, l As Integer, theta As Double) As Double
'functie intoarce valoarea functiei de parametrii results, lambda, etc...
On Error GoTo handleit
functie = pondere * ((h * h * res(1) * res(1) + k * k * res(2) * res(2) + l * l * res(3) * res(3) + 2 * l * h * res(3) * res(1) * Cos(res(5) / rd) + 2 * l * k * res(3) * res(2) * Cos(res(4) / rd) + 2 * h * k * res(2) * res(1) * Cos(res(6) / rd) - 4 * Sin((theta - res(8) / 2) / rd) * Sin((theta - res(8) / 2) / rd) / res(7) / res(7)))
Exit Function
handleit:
raport "error in computing the derivative..."
Exit Function
End Function

Function corectie(X, polcoeff() As Double) As Double
On Error GoTo handleit
Dim i As Integer
corectie = polcoeff(1)
For i = 2 To 9
corectie = corectie + polcoeff(i) * X ^ (i - 1)
Next i
Exit Function
handleit:
corectie = 0
Err.Clear
Exit Function
End Function

Sub InterPolynomial(poldeg As Integer, npoints As Integer, xdata() As Double, ydata() As Double, eps As Double, solution() As Double, errorcode As Boolean)
'this routine determines the coeeficients for the interpolation polynomial
'ydata=c0+c1xdata+c2xdata^2..etc of order norder
'as input requires:
'polynomial degree <poldeg>
'number of data points <npoints>, (npoints) must be >= (poldeg+1)
'xdata, array of length (npoints) (option base 1), double
'ydata, array of length (npoints), double
'precision limit eps, double , used further in the matrix inversion, usually 10-12
'as output gives
'errorcode a boolean of succes, true is trouble, false is OK
'solution(poldeg+1), an array of double storing the coefficients of the polynomial, the first is the free term
'external : use PseudoInverse matrix routine
'---made in january 99 from other old pieces, ND
Dim z() As Double, i As Integer, j As Integer
ReDim z(npoints, poldeg + 1)
'I have to buid the matrix z
For i = 1 To npoints
    For j = 1 To poldeg + 1
    z(i, j) = xdata(i) ^ (j - 1)
    Next j
Next i
'call the inverse ''it was poldeg+1
Call pseudoinv(npoints, poldeg + 1, z(), ydata(), solution(), eps, errorcode)
'use exit sub to clear err. event
Exit Sub
End Sub

Sub ComputeShift2Theta(w1 As Double, w2 As Double, doit1 As Double, doit2 As Double, eroare As Boolean)
'this routine is used for alpha2 strip
'w1 is the 1st wavelength, a1
'w2 is a2
'doit1 is where is appeared - consider only a1 contribution
'doit2 is the theta for the a2 contribution
'eroare is boolean
On Error GoTo errorTRAP
eroare = False
'output is eroare and doit2
Dim d As Double
d = w1 / (2 * Sin(doit1 / 2 / rd))
doit2 = rd * 2 * asin(w2 / 2 / d)
Exit Sub
errorTRAP:
eroare = True
Exit Sub








End Sub






Sub CoefOfaParabola(X() As Double, Y() As Double, solution() As Double, eroare As Boolean)
'determines the coefficients of a parabola, unique solution,...
'solution 1 is a, 2 is b and 3 is c where y=a+bx+cx2
On Error GoTo errorTRAP
eroare = False
solution(3) = (Y(1) * (X(2) - X(1)) - Y(3) * (X(2) - X(1)) - X(1) * Y(2) - X(1) * Y(1)) / (X(1) * X(1) + X(1) * (X(2) * X(2) - X(1) * X(1)))
solution(2) = (Y(2) - Y(1) - solution(3) * (X(2) * X(2) - X(1) * X(1))) / (X(2) - X(1))
solution(1) = Y(1) - solution(2) * X(1) - solution(3) * X(1) * X(1)
Exit Sub
errorTRAP:
eroare = True
Exit Sub

End Sub





Sub IntPolValue(poldeg As Integer, coeff() As Double, X As Double, yout As Double, errorcode As Boolean)
'this routine determines the value of the interpolated polynomial
'as input
'poldeg, an integer >0, the degree of the interpolated polynomial
'coeff(poldeg+1), a double array, the first one is the free term, c0, base 1
'x, the value of x where the polynomial should be computed, double
'yout, the actual value of the polynomial, double
'errorcode, a test of the succes or not of the routine; false is good
'external: none
On Error GoTo errtrap
Dim i As Integer, j As Integer
errorcode = False
yout = 0
For i = 1 To poldeg + 1
yout = yout + (X ^ (i - 1)) * coeff(i)
Next i
Exit Sub
errtrap:
errorcode = True
Exit Sub
End Sub



Sub calculdtheta(ar As Double, br As Double, cr As Double, alr As Double, ber As Double, gar As Double, lambda As Double, zero As Double, h As Integer, k As Integer, l As Integer, dexp As Double, dcalc As Double, ddif As Double, thetaexp As Double, thetacalc As Double, thetacor As Double, thetadif As Double, coderoare)
On Error GoTo handleit:
Dim sinsqteta As Double
sinsqteta = lambda * lambda / 4 * (h * h * ar * ar + k * k * br * br + l * l * cr * cr + 2 * l * h * cr * ar * Cos(ber / rd) + 2 * l * k * cr * br * Cos(alr / rd) + 2 * h * k * br * ar * Cos(gar / rd))
sinsqteta = Sqr(sinsqteta)
thetacalc = rd * Atn(sinsqteta / Sqr(-sinsqteta * sinsqteta + 1))
thetacor = thetaexp + zero / 2
thetadif = thetacalc - thetacor
dexp = lambda / 2 / (Sin(thetacor / rd))
dcalc = lambda / 2 / (Sin(thetacalc / rd))
ddif = dcalc - dexp
coderoare = False
Exit Sub
handleit:
coderoare = True
Exit Sub
End Sub



Function calcul_theta(lambda, h As Integer, k As Integer, l As Integer, ar As Double, br As Double, cr As Double, alr As Double, ber As Double, gar As Double) As Double
Dim sinsqteta As Double
sinsqteta = lambda * lambda / 4 * (h * h * ar * ar + k * k * br * br + l * l * cr * cr + 2 * l * h * cr * ar * Cos(ber / rd) + 2 * l * k * cr * br * Cos(alr / rd) + 2 * h * k * br * ar * Cos(gar / rd))
sinsqteta = Sqr(sinsqteta)
calcul_theta = rd * Atn(sinsqteta / Sqr(-sinsqteta * sinsqteta + 1))
End Function

Sub mcrnl(q() As Double, id As Integer, Y() As Double, b() As Double, bb() As Double, afi() As Integer, m As Integer, N As Integer, P() As Double, ifin As Integer, indic As Integer, h() As Integer, k() As Integer, l() As Integer)
On Error GoTo handleit
''Call mcrnl(qq, ndmax, theta, bb, npaf, nr, pds, IFIN)
  Dim mm As Integer, m1 As Integer, m2 As Integer, j As Integer, i As Integer, IQ As Integer, il As Integer
  Dim ae As Double, be As Double, ce As Double, cae As Double, cbe As Double, cce As Double
Dim dd As Double, rad As Double, d As Double, f As Double, ir As Integer
Dim r As Double, no As Integer, a(120) As Double, ier As Integer, imm As Integer, qq() As Double
   ReDim qq(N, 10)
      mm = 2 * m
      m1 = m + 1
      m2 = m + 2
      
510
      ifin = ifin - 1
      j = 0
''calc----
''Call calc(q, qq, b, bb, h, k, l, theta)
''For i = 1 To m: raport CStr(bb(i)): Next i
    For i = 1 To 8
      If (afi(i) = 0) Then GoTo 1
      j = j + 1
      b(i) = bb(j)
1
     Next i
      If (indic = 1 Or indic = 2) Then b(4) = b(3)
      If (indic = 1) Then b(5) = b(3)
      ae = b(3)
      be = b(4)
      ce = b(5)
      cae = Cos(b(6))
      cbe = Cos(b(7))
      cce = Cos(b(8))
          ''DO 2 I=1,NR
      For i = 1 To N
      dd = (ae * h(i)) ^ 2 + (be * k(i)) ^ 2 + (ce * l(i)) ^ 2 + 2# * (h(i) * k(i) * ae * be * cce + k(i) * l(i) * be * ce * cae + l(i) * h(i) * ce * ae * cbe)
           
      d = 1# / Sqr(dd)
      rad = Sqr(Abs((1# - (b(2) ^ 2) * dd)))  ''????????
      f = b(2) * d / rad
      q(i, 1) = 1#
      q(i, 2) = 1# / (d * rad)
      q(i, 3) = f * h(i) * (h(i) * ae + k(i) * be * cce + l(i) * ce * cbe)
      q(i, 4) = f * k(i) * (k(i) * be + l(i) * ce * cae + h(i) * ae * cce)
      q(i, 5) = f * l(i) * (l(i) * ce + h(i) * ae * cbe + k(i) * be * cae)
      q(i, 6) = -f * k(i) * l(i) * be * ce * Sin(b(6))
      q(i, 7) = -f * l(i) * h(i) * ce * ae * Sin(b(7))
      q(i, 8) = -f * h(i) * k(i) * ae * be * Sin(b(8))
2     qq(i, m + 2) = b(1) + Atn((b(2) / d) / Sqr(Abs((-(b(2) / d) * (b(2) / d) + 1)))) ''a fost qq
      Next i
      For ir = 1 To N
      ''DO 3 IR=1,NR
      j = 0
      For i = 1 To 8
      ''DO 3 I=1,8
      If (afi(i) = 0#) Then GoTo 3
      j = j + 1
      qq(ir, j) = q(ir, i) ''a fost q
3
Next i
Next ir  ''bb sunt rezultatele  ''  WRITE(IWR,5)(BB(I),I=1,NPAF)

''C-----LES Y CALCULES SONT DANS LA COLONNE M+2
     For j = 1 To m
''      DO 30 J=1,m
      r = 0#
  For i = 1 To N
      ''DO 20 I=1,N
    r = r + (Y(i) - qq(i, m2)) * qq(i, j) * P(i)
    Next i
    qq(j, m1) = r
30 Next j
''C-----CONSTRUCTION DE LA MATRICE SYMETRIQUE A=Q*QT
      no = 0
  For IQ = 1 To m
  For il = IQ To m
  ''    DO 50 IQ=1,m
   ''   DO 50 IL=IQ,M
      no = no + 1
      r = 0#
      For i = 1 To N
''      DO 40 I=1,N
    r = r + qq(i, IQ) * qq(i, il) * P(i)
Next i
    a(no) = r
Next il
Next IQ


Dim lmn As Integer

Call matinv(a, m, ier)
  If ier > 0 Then raport "Error flag in matrix inversion..."
  If ifin <= 0 Then GoTo 5110
''C-----CALCUL DES NOUVEAUX PARAMETRES
560
   For i = 1 To m
''   DO 100 I=1,M
      r = 0#
      imm = (i - 1) * (mm - i) / 2
  For j = 1 To m
      ''DO 90 J=1,M
      If (j - i) < 0 Then
      GoTo 580
      Else
      GoTo 570
      End If
570    lmn = imm + j
      GoTo 590
580    lmn = (j - 1) * (mm - j) / 2 + i
590    r = r + a(lmn) * qq(j, m1)
Next j
5100   bb(i) = bb(i) + r ''a fost bb
Next i
      GoTo 510
''C-----REMISE DES ELEMENTS DIAGONAUX DANS Q(I,I),APRES LE DERNIER CYCLE
5110
   ''DO 120 I=1,M
     For i = 1 To m
      lmn = (i - 1) * (mm - i) / 2 + i
5120   q(i, i) = a(lmn) ''??????????
    Next i


Exit Sub
handleit:
raport "Unexpected error in least squares routine..."
Err.Clear
Exit Sub
End Sub



Sub matinv(a, N, nfail)
Dim kmn As Integer, suma As Double, m As Integer, imax As Integer, lmn As Integer
Dim kli As Integer, kmi As Integer, i As Integer, j As Integer, term As Double, denom As Double
Dim kdm As Integer, ii As Integer
''C-----INVERSION DE LA MATRICE A
      kmn = 1
      If N - 1 < 0 Then GoTo 480
       If N - 1 = 0 Then GoTo 410
      If N - 1 > 0 Then GoTo 420
     
''      IF(N-1)80,10,20
410    a(1) = 1# / a(1)
      GoTo 4200
420
For m = 1 To N
   ''20 DO 110 M=1,N
      imax = m - 1
      For lmn = m To N      ''??????/
      ''DO 100 L=M,N
      suma = 0#
      kli = lmn
      kmi = m
      If (imax <= 0) Then
      GoTo 450
      Else
      GoTo 430
      End If
''C     *****SUM OVER I=1,M-1 A(L,I)*A(M,I) *****
430
For i = 1 To imax
   ''DO 40 I=1,IMAX
      suma = suma + a(kli) * a(kmi)
      j = N - i
      kli = kli + j
440    kmi = kmi + j
Next i
''C     *****TERM=C(L,M)-SUM *****
450    term = a(kmn) - suma
      If (lmn - m) <= 0 Then
      GoTo 460
      Else
      GoTo 490
      End If
      
460
      If (term <= 0) Then
      GoTo 480
      Else
      GoTo 470
      End If
''C     ***** A(M,M)=SQRT(TERM) *****
470    denom = Sqr(term)
      a(kmn) = denom
      GoTo 4100
480    nfail = kmn
      GoTo 4210
490    a(kmn) = term / denom
4100   kmn = kmn + 1
Next lmn
4110 Next m
4120   a(1) = 1# / a(1)
      kdm = 1
      For lmn = 2 To N
      ''DO 150 L=2,N
      kdm = kdm + N - lmn + 2
      term = 1# / a(kdm)
      a(kdm) = term
      kmi = 0
      kli = lmn
      imax = lmn - 1
''C     ***** STEP M OF B(L,M) *****
      For m = 1 To imax
      ''DO 140 M=1,IMAX
      kmn = kli
''C     ***** SUM TERMS *****
      suma = 0#
      For i = m To imax
      ''DO 130 I=M,IMAX
      ii = kmi + i
      suma = suma - a(kli) * a(ii)
4130   kli = kli + N - i
      Next i
''C     ***** MULT SUM * RECIP OF DIAGONAL *****
      a(kmn) = suma * term
      j = N - m
      kli = kmn + j
4140   kmi = kmi + j
Next m
4150  Next lmn
4160   kmn = 1
      For m = 1 To N
      kli = kmn
      For lmn = m To N
      kmi = kmn
      imax = N - lmn + 1
      suma = 0#
      For i = 1 To imax
      suma = suma + a(kli) * a(kmi)
      kli = kli + 1
4170   kmi = kmi + 1
      Next i
      a(kmn) = suma
4180   kmn = kmn + 1
      Next lmn
      Next m
4190
4200   nfail = 0
4210

End Sub


Sub inver(b() As Double, db() As Double, volum As Double, iv As Integer)
Dim cabc2 As Integer, i As Integer, ad(3) As Double, cosp(3) As Double, sInp(3) As Double, j As Integer, jj As Integer, kmn As Integer, dqd(3) As Double, ss(3) As Double
Dim cc(3) As Double, q2 As Double, qmn As Double, dmn(6) As Double, r As Double, sig(8) As Double
Dim sp(3) As Double

cabc2 = 0
''      DO 10 I=1,3
For i = 1 To 3
   ad(i) = b(i + 2)
      cosp(i) = Cos(b(i + 5))
      sInp(i) = Sin(b(i + 5))
''10    CONTINUE
Next i
''      DO 15 I=1,3
For i = 1 To 3
     j = (i Mod 3) + 1
     kmn = ((i + 1) Mod 3) + 1
      dqd(i) = cosp(i) - cosp(j) * cosp(kmn)
      ss(i) = sInp(j) * sInp(kmn)
      cc(i) = -dqd(i) / ss(i)
      cabc2 = cabc2 + cosp(i) * cosp(i)
''15    CONTINUE
Next i
''c
      q2 = 1# - cabc2 + 2# * cosp(1) * cosp(2) * cosp(3)
      qmn = Sqr(q2)
     volum = ad(1) * ad(2) * ad(3) * qmn
''c
 ''     DO 20 I=1,3
For i = 1 To 3
    b(i + 2) = sInp(i) / (ad(i) * qmn)
    b(i + 5) = acos(cc(i))
    sp(i) = Sin(b(i + 5))
''20    CONTINUE
Next i
If iv = 0 Then Exit Sub
''c
''C.....DERIVEES DES PARAMETRES A , B , C
''      DO 30 I=1,3
''ReDim dmn(8)
For i = 1 To 3
   j = (i Mod 3) + 1
    kmn = ((i + 1) Mod 3) + 1
    dmn(i) = -sInp(i) / ad(i)
    dmn(j) = 0#
    dmn(kmn) = 0#
    dmn(i + 3) = cosp(i) - sInp(i) * sInp(i) * dqd(i) / (qmn * qmn)
    dmn(j + 3) = -ss(kmn) * dqd(j) / q2
    dmn(kmn + 3) = -ss(j) * dqd(kmn) / q2
 Call sigma(dmn, db, r)
 sig(i + 3) = Sqr(r) / sp(i)
Next i
''c
''C.....DERIVEES DES ANGLES DE LA MAILLE
For i = 1 To 3
'dO 50 I=1,3
For jj = 1 To 3
 
 ''     DO 40 JJ=1,3
 dmn(jj) = 0#
Next jj
     j = (i Mod 3) + 1
    kmn = ((i + 1) Mod 3) + 1
      dmn(i + 3) = sInp(i) / ss(i)
      dmn(j + 3) = cosp(kmn) / sInp(kmn) + cosp(j) * cc(i) / sInp(j)
      dmn(kmn + 3) = cosp(j) / sInp(j) + cosp(kmn) * cc(i) / sInp(kmn)

''---------
Call sigma(dmn, db, r)
sig(i + 3) = Sqr(r) / sp(i)
Next i
''c
 ''DO 60 I=1,6
For i = 1 To 6
    sig(i + 2) = sig(i)
Next i
'

End Sub






Sub sigma(dmn() As Double, sig() As Double, r As Double)
On Error GoTo handleit
Dim ii As Integer
       r = 0#
For ii = 1 To 6
r = r + (dmn(ii) * sig(ii + 2)) ^ 2
Next ii
Exit Sub
handleit:
raport "error in Sigma routine.."
Err.Clear
Exit Sub
End Sub



Function CloseWindow(cwMesaj As String, cwTitle As String) As Boolean
Dim t As Integer
CloseWindow = False
t = MsgBox(cwMesaj, vbOKCancel + vbDefaultButton1, cwTitle)
If t = vbOK Then CloseWindow = True
Exit Function
End Function



















Sub generate_reflections(ByRef r() As Double, b As String, tlimit As Single, coderoare As Boolean)
On Error GoTo handleit
coderoare = False
Dim b1 As String * 1, b2 As String * 4, h As Integer, k As Integer, l As Integer, i As Integer, outfil As Integer
Dim returncode As Boolean, lasthkl As Integer
Dim date_c(2000) As crystallo



b1 = LCase(left$(b, 1))
b2 = LCase(right$(b, 4))
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The file is " & outputfile

i = 0
Select Case b2
Case "cubi"
'cazul cel mai general este primitiv; exclud ceva daca apar semne de extinctie
For l = 0 To 30
For k = 0 To 30
For h = 0 To 30


Next h
Next k
Next l







Case "tetr"

Case "hexa"

Case "romb"

Case "orth"
Case "mono"

Case "tric"
Case Else
raport "unknown identifier in reflection generator - " & b
Err.Raise 1101
End Select

Exit Sub
handleit:
coderoare = True
Err.Clear
End Sub







Sub searchlinear(celula As Integer, results() As Double, ind() As Integer, steps As Integer, widthsearch As Double, nrval As Integer, t() As Double, h() As Integer, k() As Integer, l() As Integer, pondere() As Double, intoarce() As Double, sumamin As Double, coderoare As Boolean)
On Error GoTo handleit
Dim s As Double, i As Integer, dv1 As Double, dv2 As Double, dv3 As Double, dv4 As Double, dv5 As Double, dv6 As Double, dv7 As Double, dv8 As Double
''Dim smin As Double
Dim resultint(8) As Double
''smin = 100000
sumamin = 1.4E+34
DoEvents

Select Case celula
Case 0
    For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To results(1) + ind(1) * widthsearch * results(1) Step widthsearch * results(1) / steps
    For dv7 = Abs(results(7) - ind(7) * widthsearch * results(7)) To results(7) + ind(7) * widthsearch * results(7) Step widthsearch * results(7) / steps
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step 0.1 / steps
    For i = 1 To nrval
        s = s + ((pondere(i) * (h(i) * h(i) + k(i) * k(i) + l(i) * l(i)) * dv1 * dv1) - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7) ^ 2
    DoEvents
    Next i
    If s < sumamin Then
    sumamin = s
    resultint(1) = dv1
    resultint(2) = dv1
    resultint(3) = dv1
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(1)
    results(3) = resultint(1)
    results(7) = resultint(7)
    results(8) = resultint(8)
Case 1 'tetragonal
    For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To (results(1) + ind(1) * widthsearch * results(1)) Step ((widthsearch * results(1) / steps))
    For dv3 = Abs(results(3) - ind(3) * widthsearch * results(3)) To results(3) + ind(3) * widthsearch * results(3) Step widthsearch * results(3) / steps
    For dv7 = Abs(results(7) - ind(7) * widthsearch * results(7)) To (results(7) + ind(7) * widthsearch * results(7)) Step (widthsearch * results(7) / steps)
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step (0.1 / steps)
    For i = 1 To nrval
    s = s + (((pondere(i) * ((h(i) * h(i) + k(i) * k(i)) * dv1 * dv1 + l(i) * l(i) * dv3 * dv3)) - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7) ^ 2)
    DoEvents
    Next i
    If s < sumamin Then
    sumamin = s
    resultint(1) = dv1
    resultint(3) = dv3
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv3
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(1)
    results(3) = resultint(3)
    results(7) = resultint(7)
    results(8) = resultint(8)

    
Case 2 'orto
    For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To (results(1) + ind(1) * widthsearch * results(1)) Step ((widthsearch * results(1) / steps))
    For dv2 = Abs(results(2) - ind(2) * widthsearch * results(2)) To (results(2) + ind(2) * widthsearch * results(2)) Step ((widthsearch * results(2) / steps))
    For dv3 = Abs(results(3) - ind(3) * widthsearch * results(3)) To results(3) + ind(3) * widthsearch * results(3) Step widthsearch * results(3) / steps
    For dv7 = (results(7) - ind(7) * widthsearch * results(7)) To (results(7) + ind(7) * widthsearch * results(7)) Step (widthsearch * results(7) / steps)
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step (0.1 / steps)
    For i = 1 To nrval
    s = s + (pondere(i) * (h(i) * h(i) * dv1 * dv1 + k(i) * k(i) * dv2 * dv2 + l(i) * l(i) * dv3 * dv3) - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7) ^ 2
    DoEvents
    Next i
    If s < sumamin Then
    sumamin = s
    resultint(1) = dv1
    resultint(2) = dv2
    resultint(3) = dv3
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv3
    Next dv2
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(2)
    results(3) = resultint(3)
    results(7) = resultint(7)
    results(8) = resultint(8)
Case 3 ''rom"
    For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To (results(1) + ind(1) * widthsearch * results(1)) Step ((widthsearch * results(1) / steps))
    For dv4 = Abs(results(4) - ind(4) * widthsearch * results(4)) To (results(4) + ind(4) * widthsearch * results(4)) Step ((widthsearch * results(4) / steps))
    For dv7 = (results(7) - ind(7) * widthsearch * results(7)) To (results(7) + ind(7) * widthsearch * results(7)) Step (widthsearch * results(7) / steps)
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step (0.1 / steps)
    For i = 1 To nrval
    s = s + ((pondere(i) * ((h(i) * h(i) + k(i) * k(i) + l(i) * l(i) + 2 * (k(i) * l(i) + h(i) * l(i) * h(i) * k(i)) * Cos(dv4))) * dv1 * dv1 - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7)) ^ 2
    DoEvents
    Next i
    If s < sumamin Then
    sumamin = s
    resultint(1) = dv1
    resultint(4) = dv4
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv4
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(1)
    results(3) = resultint(1)
    results(4) = resultint(4)
    results(5) = resultint(4)
    results(6) = resultint(4)
    results(7) = resultint(7)
    results(8) = resultint(8)

Case 4 ''"hex"
    For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To (results(1) + ind(1) * widthsearch * results(1)) Step ((widthsearch * results(1) / steps))
    For dv3 = Abs(results(3) - ind(3) * widthsearch * results(3)) To (results(3) + ind(3) * widthsearch * results(3)) Step ((widthsearch * results(3) / steps))
    For dv7 = (results(7) - ind(7) * widthsearch * results(7)) To (results(7) + ind(7) * widthsearch * results(7)) Step (widthsearch * results(7) / steps)
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step (0.1 / steps)
    For i = 1 To nrval
    s = s + ((pondere(i) * ((h(i) * h(i) + k(i) * k(i) + h(i) * k(i)) * dv1 * dv1 + l(i) * l(i) * dv3 - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7)) ^ 2)
    Next i
    DoEvents
    If s < sumamin Then
    sumamin = s
    raport "better data..." & "  --> " & CStr(sumamin)
    resultint(1) = dv1
    resultint(3) = dv3
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv3
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(1)
    results(3) = resultint(3)
    results(7) = resultint(7)
    results(8) = resultint(8)

Case 5 ''"mon"
   For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To (results(1) + ind(1) * widthsearch * results(1)) Step ((widthsearch * results(1) / steps))
    For dv2 = Abs(results(2) - ind(2) * widthsearch * results(2)) To (results(2) + ind(2) * widthsearch * results(2)) Step ((widthsearch * results(2) / steps))
    For dv3 = Abs(results(3) - ind(3) * widthsearch * results(3)) To (results(4) + ind(3) * widthsearch * results(3)) Step ((widthsearch * results(3) / steps))
    For dv5 = Abs(results(5) - ind(5) * widthsearch * results(5)) To (results(4) + ind(5) * widthsearch * results(5)) Step ((widthsearch * results(5) / steps))
    For dv7 = (results(7) - ind(7) * widthsearch * results(7)) To (results(7) + ind(7) * widthsearch * results(7)) Step (widthsearch * results(7) / steps)
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step (0.1 / steps)
    For i = 1 To nrval
        s = s + (pondere(i) * (h(i) * h(i) * dv1 * dv1 + k(i) * k(i) * dv2 * dv2 + l(i) * l(i) * dv3 * dv3 + 2 * l(i) * h(i) * dv3 * dv1 * Cos(dv5 / rd) - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7)) ^ 2
    DoEvents
    Next i
    If s < sumamin Then
    sumamin = s
    resultint(1) = dv1
    resultint(2) = dv2
    resultint(3) = dv3
    resultint(5) = dv5
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv5
    Next dv3
    Next dv2
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(2)
    results(3) = resultint(3)
    results(4) = resultint(4)
    results(5) = resultint(5)
    results(6) = resultint(6)
    results(7) = resultint(7)
    results(8) = resultint(8)

Case 6 ''"tri"
    For dv1 = Abs(results(1) - ind(1) * widthsearch * results(1)) To (results(1) + ind(1) * widthsearch * results(1)) Step ((widthsearch * results(1) / steps))
    For dv2 = Abs(results(2) - ind(2) * widthsearch * results(2)) To (results(2) + ind(2) * widthsearch * results(2)) Step ((widthsearch * results(2) / steps))
    For dv3 = Abs(results(3) - ind(3) * widthsearch * results(3)) To (results(3) + ind(3) * widthsearch * results(3)) Step ((widthsearch * results(3) / steps))
    For dv4 = Abs(results(4) - ind(4) * widthsearch * results(4)) To (results(4) + ind(4) * widthsearch * results(4)) Step ((widthsearch * results(4) / steps))
    For dv5 = Abs(results(5) - ind(5) * widthsearch * results(5)) To (results(5) + ind(5) * widthsearch * results(5)) Step ((widthsearch * results(5) / steps))
    For dv6 = Abs(results(6) - ind(6) * widthsearch * results(6)) To (results(6) + ind(6) * widthsearch * results(6)) Step ((widthsearch * results(6) / steps))
    For dv7 = (results(7) - ind(7) * widthsearch * results(7)) To (results(7) + ind(7) * widthsearch * results(7)) Step (widthsearch * results(7) / steps)
    For dv8 = results(8) - ind(8) * widthsearch To results(8) + ind(8) * widthsearch Step (0.1 / steps)
    For i = 1 To nrval
        s = s + (pondere(i) * (h(i) * h(i) * dv1 * dv1 + k(i) * k(i) * dv2 * dv2 + l(i) * l(i) * dv3 * dv3 + 2 * l(i) * h(i) * dv3 * dv1 * Cos(dv5 / rd) + 2 * l(i) * k(i) * dv3 * dv2 * Cos(dv4 / rd) + 2 * h(i) * k(i) * dv2 * dv1 * Cos(dv6 / rd) - pondere(i) * 4 * Sin((t(i) + dv8 / 2) / rd) * Sin((t(i) + dv8 / 2) / rd) / dv7 / dv7)) ^ 2
    DoEvents
    Next i
    If s < sumamin Then
    sumamin = s
    resultint(1) = dv1
    resultint(2) = dv2
    resultint(3) = dv3
    resultint(4) = dv4
    resultint(5) = dv5
    resultint(6) = dv6
    resultint(7) = dv7
    resultint(8) = dv8
    End If
    s = 0
    Next dv8
    Next dv7
    Next dv6
    Next dv5
    Next dv4
    Next dv3
    Next dv2
    Next dv1
    results(1) = resultint(1)
    results(2) = resultint(2)
    results(3) = resultint(3)
    results(4) = resultint(4)
    results(5) = resultint(5)
    results(6) = resultint(6)
    results(7) = resultint(7)
    results(8) = resultint(8)

End Select
''sumamin = smin

coderoare = False
Exit Sub
handleit:
''sumamin = smin

coderoare = True
Exit Sub
End Sub

Sub verificadate(cerex As Boolean, cerefullx As Boolean, cerey As Boolean, returnok As Boolean)
'face verificare generala a datelor pentru salvare
Dim i As Single, j As Double
On Error GoTo errorTRAP
returnok = False

If cerey Then
    If Not (amydata) Then
        raport "I don't have any data..."
        i = MsgBox("I don't have any data. Read a data file and then try again.", , prog_name): returnok = False: Exit Sub
     End If
End If


If cerex Then
    If (Not (amxdata)) Then
     raport "I don't have the X data range."
        j = MsgBox("I don't have the x data range. You have to enter the start 2 theta value and step.", vbDefaultButton1 + vbOKCancel, prog_name)
        If j = vbCancel Then returnok = False: Exit Sub
        j = InputBox("Enter initial 2 theta", prog_name, CStr(X(1)))
        startx = Val(j)
        j = InputBox("Enter step 2 theta", prog_name, CStr(Fix(10000 * (X(2) - X(1)))) / 10000)
        stepx = Val(j)
        endx = startx + (numarvalori - 1) * stepx
        If startx < 0 Or stepx = 0# Then MsgBox "Incorrect values": returnok = False: Exit Sub
        amxdata = True
End If
End If

    If cerefullx Then
    If Not (amfullxdata) Then
        If (amxdata) Then
            ReDim X(numarvalori)
            For i = 1 To numarvalori
            X(i) = startx + (i - 1) * stepx
            Next i
            raport "Computing X values..."
        Else
            raport "I don't have the X data range."
            i = MsgBox("I don't have the x data range. You have to enter the start 2 theta value and step.", vbDefaultButton1 + vbOKCancel, prog_name)
            If i = vbCancel Then returnok = False: Exit Sub
            j = InputBox("Enter initial 2 theta", prog_name, 0)
            startx = Val(j)
            j = InputBox("Enter step 2 theta", prog_name, 0.02)
            stepx = Val(j)
            endx = startx + (numarvalori - 1) * stepx
            If startx < 0 Or stepx = 0# Then MsgBox "Incorrect values": returnok = False: Exit Sub
            
            ReDim X(numarvalori)
            For i = 1 To numarvalori
            X(i) = startx + (i - 1) * stepx
            Next i
            amfullxdata = True
        End If
    End If
    End If

If Not (amzdata) Then
ReDim z(numarvalori)  'it was preserve, and two lines up

For i = 1 To numarvalori
z(i) = 0.1
If Y(i) > 0 Then z(i) = Sqr(Y(i))
Next i
End If

returnok = True
Exit Sub
errorTRAP:
raport Err.Description
raport "An error occured in CheckData routine. Have you inserted good values ?"
returnok = False
Err.Clear
Exit Sub
End Sub






Function CopyFile(Src As String, Dst As String) As Single
Dim response As Integer
Static Buf$
Dim BTest!, FSize! 'declare the needed variables
Dim Chunk%, F1%, F2%

Const BUFSIZE = 1024 'set the buffer size

If Len(Dir(Dst)) Then 'check to see if the destination file already exists
   'response = MsgBox(Dst + Chr(10) + Chr(10) + "File already exists. Do you want to overwrite it?", vbYesNo + vbQuestion) 'prompt the user with a message box
    response = vbYes
   If response = vbNo Then 'if the "No" button was clicked
      Exit Function 'exit the procedure
   Else             'otherwise
      Kill Dst      'delete the already found file, and carryon with the code
   End If
End If
 
'On Error GoTo FileCopyError 'incase of error goto this label
F1 = FreeFile 'returns file number available
Open Src For Binary As F1 'open the source file
F2 = FreeFile 'returns file number available
Open Dst For Binary As F2 'open the destination file
 
FSize = LOF(F1)
BTest = FSize - LOF(F2)

Do
If BTest < BUFSIZE Then
   Chunk = BTest
Else
   Chunk = BUFSIZE
End If
      
Buf = String(Chunk, " ")
Get F1, , Buf
Put F2, , Buf
BTest = FSize - LOF(F2)

'ProgressBar.Value = (100 - Int(100 * BTest / FSize)) 'advance the progress bar as the file is copied

Loop Until BTest = 0
Close F1 'closes the source file
Close F2 'closes the destination file
CopyFile = FSize
'ProgressBar.Value = 0 'returns the progress bar to zero
Exit Function 'exit the procedure

FileCopyError: 'file copy error label
MsgBox " Error trying to write a data file..." 'display message box with error
Close F1 'closes the source file
Close F2 'closes the destination file
Return
Exit Function 'exit the procedure

End Function








Sub allow_line(dd As crystallo, doitetalimit As Single, cod As String, permis As Boolean)
On Error GoTo handleit
If dd.d <= 0 Or dd.doitheta > doitetalimit Then Err.Raise 1101
permis = True
Select Case cod
Case "pcubi"

Case "icubi"
Case "fcubi"
Case "ptetr"
Case "itetr"
Case "porth"
Case "corth"
Case "iorth"
Case "forth"
Case "rromb"
Case "phexa"
Case "pmono"
Case "cmono"
Case "ptric"

End Select

Exit Sub
handleit:
permis = False


End Sub



Sub raport(mesaj As String)
Dim outfil As Integer
DoEvents
On Error GoTo errorTRAP
Convert3Main.txtraport.Text = Convert3Main.txtraport.Text & vbCrLf & mesaj
Exit Sub
errorTRAP:
'ma mult de 32 de k ??
outfil = FreeFile
Open App.Path & "\Pw3_ReportPad.bak" For Append As outfil '
Print #outfil, Convert3Main.txtraport.Text
Close outfil
Convert3Main.txtraport.Text = Now & vbCrLf & "An error has occured. Report appended to pw3_ReportPad.bakt"
Err.Clear
Exit Sub
End Sub




 Sub averageint(catevaloriam As Long, valori() As Double, valoaremedie As Double, valoaremaxima As Double, coderoare As Boolean)
On Error GoTo errorTRAP
Dim i As Long, suma As Double
suma = 0
valoaremaxima = 0
For i = 1 To catevaloriam
If valori(i) > valoaremaxima Then valoaremaxima = valori(i)
suma = suma + valori(i)
Next i
valoaremedie = suma / catevaloriam
coderoare = False
Exit Sub
errorTRAP:
coderoare = True
Exit Sub
End Sub





Function acos(X As Double) As Double
On Error GoTo handleit
Dim dif As Double
If X = 0# Then GoTo 20
If X >= 1 And X < 1.0001 Then GoTo 30
If X <= -1 And X > -1.0001 Then GoTo 40
    dif = Sqr(1# - X * X) / X
 If dif >= 0 Then GoTo 10
    dif = -dif
    acos = 3.141592 - Atn(dif)
Exit Function
10  acos = Atn(dif)
Exit Function
20  acos = 1.570796
Exit Function
30  acos = 0#
Exit Function
40  acos = 3.1415926
Exit Function
handleit:
raport strLinie
raport "Fatal error in ACOS function." & vbCrLf & strLinie
Err.Clear
Exit Function
End Function
'amfull inseamna ca am citit x pentru fiecare y, amx inseamna ca am doar startx, stepx..

Function number_of_lines_for_indexing() As Integer
On Error GoTo errorTRAP
Dim N As Integer
'return an integer n, which is the number of lines in the grid
'if n<0 then this is an error indication
frmRefine.grid.Row = 1
frmRefine.grid.Col = 4
'out from here with error or when the value of the text is 0
N = 1
Do
If Val(CStr(frmRefine.grid.Text)) <= 0 Then Exit Do
N = N + 1
frmRefine.grid.Row = N
Loop
number_of_lines_for_indexing = N - 1

Exit Function
errorTRAP:
If N = 0 Then N = -2
number_of_lines_for_indexing = N - 2
Err.Clear
Exit Function
End Function



Sub reciproc(inp() As Double, out() As Double, eroare As Boolean)
'input si output sunt vectori de 7, a,b,c...si volum
'daca eroare este true atunci a fost ceva nasol...
On Error GoTo errorTRAP
eroare = False
out(1) = inp(2) * inp(3) * Sin(inp(4) / rd) / inp(7)
out(2) = inp(1) * inp(3) * Sin(inp(5) / rd) / inp(7)
out(3) = inp(1) * inp(2) * Sin(inp(6) / rd) / inp(7)
out(4) = (Cos(inp(5) / rd) * Cos(inp(6) / rd) - Cos(inp(4) / rd)) / (Sin(inp(5) / rd) * Sin(inp(6) / rd))
out(5) = (Cos(inp(4) / rd) * Cos(inp(6) / rd) - Cos(inp(5) / rd)) / (Sin(inp(4) / rd) * Sin(inp(6) / rd))
out(6) = (Cos(inp(4) / rd) * Cos(inp(5) / rd) - Cos(inp(6) / rd)) / (Sin(inp(5) / rd) * Sin(inp(4) / rd))
out(7) = 1 / inp(7)
'back to deg or 1/deg; out is cos now, transform it ...
out(4) = acos(out(4)) * rd
out(5) = acos(out(5)) * rd
out(6) = acos(out(6)) * rd
'adjust to the fifth decimal point...
out(4) = Int(100000 * out(4) + 1) / 100000
out(5) = Int(100000 * out(5) + 1) / 100000
out(6) = Int(100000 * out(6) + 1) / 100000

Exit Sub
errorTRAP:
eroare = True
Exit Sub
End Sub




Sub pseudoinv(ne As Integer, N As Integer, z() As Double, ii() As Double, X() As Double, lowsize As Double, eroare As Boolean)
'ne este numarul de ecuatii
'n este numarul de necunoscute, n<=ne
'lowsize este valoarea minima a pivotului, de ordin a 10^-10
'z este matricea coeficientilor sistemului de determinat
'i este termenul liber
'x este solutia
eroare = False
On Error GoTo handleit
Dim i As Integer, j As Integer, k As Integer, semn As Integer, lc As Integer
Dim maxx As Double, d As Double, aaA As Double, m As Integer, bb() As Double
ReDim a(N, N) As Double, X(N)
    For i = 1 To N: For j = 1 To N: a(i, j) = 0: For k = 1 To ne
        a(i, j) = a(i, j) + z(k, i) * z(k, j)
    Next k: Next j: Next i
 ReDim c(N, N) As Double
    For i = 1 To N
         c(i, i) = 1
     Next i
 semn = 1
     For m = 1 To N - 1: maxx = Abs(a(m, m)): lc = m
         For i = m + 1 To N
         If Abs(a(i, m)) > maxx Then maxx = Abs(a(i, m)): lc = i
         Next i
         If maxx < lowsize Then Err.Raise 1101
         For i = 1 To N
             d = a(lc, i): a(lc, i) = a(m, i): a(m, i) = d
             d = c(lc, i): c(lc, i) = c(m, i): c(m, i) = d
         Next i
 If lc <> m Then semn = -semn
 For i = m + 1 To N: aaA = a(i, m)
     For j = 1 To N
         a(i, j) = a(i, j) - a(m, j) * aaA / a(m, m)
         c(i, j) = c(i, j) - c(m, j) * aaA / a(m, m)
     Next j: Next i: Next m
     For m = N To 2 Step -1
         For i = m - 1 To 1 Step -1: aaA = a(i, m)
             For j = N To 1 Step -1
             a(i, j) = a(i, j) - a(m, j) * aaA / a(m, m)
             c(i, j) = c(i, j) - c(m, j) * aaA / a(m, m)
             Next j
         Next i
     Next m
 For i = 1 To N: For j = 1 To N: c(i, j) = c(i, j) / a(i, i): Next j: Next i
 ReDim bb(N, ne)
 For i = 1 To N
 For j = 1 To ne
 For k = 1 To N
 bb(i, j) = bb(i, j) + c(i, k) * z(j, k)
 Next k
 Next j
 Next i
 
 For i = 1 To N

 X(i) = 0
 For k = 1 To ne
 X(i) = X(i) + bb(i, k) * ii(k)
 Next k
 Next i
 Erase a, c, bb 's ar putea sa fie inutile, nu sunt globale
Exit Sub
handleit:
eroare = True
Exit Sub
End Sub



Sub open_file(ByVal nume_fisier As String, ByVal intrare_iesire As Integer, return_code As Boolean)
return_code = False 'something is wrong
'nume_fisier must be with the all path name/DOS
'scriere si citire din directorul curent
'if intrare_iesire=1 then input, the file must exist
'if intrare_iesire=2 then output, the file should not exist,
'eventually ask for overwrite
'check for error to input and output
'return_code is true if everthing is OK, false if cancel or error
'--------------------------------------------------
On Error GoTo error_open
Convert3Main.MousePointer = 11
Select Case intrare_iesire
    Case 1 'citirea path_input_text
Convert3Main.Dialog.Filter = "text file (*.txt) |*.txt|data file (*.dat) |*.dat|GSAS lst file (*.lst) |*.lst|show all (*.*) |*.*"
Convert3Main.Dialog.FilterIndex = 4
Convert3Main.Dialog.Flags = &H1000& Or &H4& Or &H800&
'ofn_filemustexist 'ofn_readonly 'ofn_pathmustexist
Convert3Main.Dialog.DialogTitle = prog_name & " - Input file"
Convert3Main.Dialog.Action = 1
inputfile = Convert3Main.Dialog.FileName
return_code = True

    Case 2 'citire path_output_text
Convert3Main.Dialog.Filter = "text file (*.txt) |*.txt|data file (*.dat) |*.dat|show all (*.*) |*.*"
Convert3Main.Dialog.FilterIndex = 3
Convert3Main.Dialog.FileName = ""
Convert3Main.Dialog.Flags = &H2& Or &H1& Or &H800& Or &H4&
'ofn_overwriteprompt'ofn_readonly'ofn_pathmustexist
Convert3Main.Dialog.DialogTitle = prog_name & " - Output file"
Convert3Main.Dialog.Action = 2
outputfile = Convert3Main.Dialog.FileName
return_code = True

End Select
Convert3Main.MousePointer = 0
Close
Exit Sub
error_open:
Err.Clear
return_code = False
Convert3Main.MousePointer = 0
Close
Exit Sub
End Sub



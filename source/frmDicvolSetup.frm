VERSION 5.00
Begin VB.Form frmDicvolSetup 
   Caption         =   "Dicvol-Setup"
   ClientHeight    =   5480
   ClientLeft      =   50
   ClientTop       =   280
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5480
   ScaleWidth      =   7980
   Begin VB.CommandButton Command1 
      Caption         =   "&Editor"
      Height          =   372
      Left            =   4320
      TabIndex        =   34
      Top             =   4800
      Width           =   1332
   End
   Begin VB.CommandButton btnSee 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Run Dicvol91"
      Height          =   375
      Left            =   480
      TabIndex        =   33
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton btnCancelDicvol 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   32
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Make file"
      Height          =   375
      Left            =   2400
      TabIndex        =   31
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "DICVOL - Properties"
      Height          =   4452
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7692
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   30
         Text            =   "0.01"
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optError 
         Caption         =   "Const. error, 2 theta"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   29
         Top             =   3600
         Width           =   2055
      End
      Begin VB.OptionButton optError 
         Caption         =   "Error for each line"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   3240
         Width           =   2415
      End
      Begin VB.OptionButton optError 
         Caption         =   "Error: 0.03 deg. 2 theta"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   26
         Text            =   "5.0"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   2
         Left            =   6120
         TabIndex        =   25
         Text            =   "0.0"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   24
         Text            =   "0.0"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   23
         Text            =   "0.0"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtmax 
         Height          =   285
         Index           =   4
         Left            =   6120
         TabIndex        =   18
         Text            =   "1500.0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtmax 
         Height          =   285
         Index           =   3
         Left            =   6120
         TabIndex        =   17
         Text            =   "125.0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtmax 
         Height          =   285
         Index           =   2
         Left            =   6120
         TabIndex        =   16
         Text            =   "20.0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtmax 
         Height          =   285
         Index           =   1
         Left            =   6120
         TabIndex        =   15
         Text            =   "20.0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtmax 
         Height          =   285
         Index           =   0
         Left            =   6120
         TabIndex        =   14
         Text            =   "20.0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtmin 
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   13
         Text            =   "100.0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtmin 
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   12
         Text            =   "90.0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Caption         =   "test triclinic"
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Caption         =   "test monoclinic"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Caption         =   "test orthorombic"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Caption         =   "test hexagonal"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Caption         =   "test tetragonal"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Alignment       =   1  'Right Justify
         Caption         =   "test cubic"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   7320
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   240
         X2              =   7320
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Right Justify
         Caption         =   "min F Merit"
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   22
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         Caption         =   "+/-"
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   21
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Right Justify
         Caption         =   "Density"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   20
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mol. weight"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblmin 
         Alignment       =   1  'Right Justify
         Caption         =   "volume range"
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblmin 
         Alignment       =   1  'Right Justify
         Caption         =   "beta range"
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblmin 
         Alignment       =   1  'Right Justify
         Caption         =   "C max"
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblmin 
         Alignment       =   1  'Right Justify
         Caption         =   "B max"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblmin 
         Alignment       =   1  'Right Justify
         Caption         =   "A max"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDicvolSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnCancelDicvol_Click()
frmRefine.mnuQuit.Enabled = True
Form_Unload (-1)
End Sub
Private Sub btnSee_Click()
Dim t As Double
Call cmdOK_Click
On Error GoTo errortrap
t = ShellAndLoop("pwd_dicvol91.exe pwd_dicvol91.in", vbMaximizedFocus)
If t = 0 Then
raport "The program pwd_Dicvol91 was called."
Else
raport "pwd_dicvol91" & " has been started." & vbCrLf & " Main window handle: " & Hex(t)
raport strLinie
End If
hWndShell "write.exe " & "pwd_dicvol91.out", vbNormalFocus
't = Shell("pwd_dicvol91.exe pwd_dicvol91.in", vbNormalFocus)
'If t = 0 Then Err.Raise 111111, , "can not call pwd_dicvol91.exe"
Exit Sub
errortrap:
raport strLinie
raport "error in btnSee Routine...dicvol91 call failed..."
raport "you can manually run the program dicvol by using the file pwd_dicvol91.in saved in the current working directory (i.e. " & CurDir & ") "
raport Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub cmdOK_Click()
Dim i As Integer, s As Integer, N As Integer, outfil As Integer, itype As Integer, jc As Integer, jt As Integer, jh As Integer, jo As Integer, jm As Integer, jtr As Integer
Dim amax As Single, bmax As Single, cmax As Single, volmin As Single, volmax As Single, bemin As Single, bemax As Single
Dim WAVE As Single, poimol As Single, DENS As Single, delden As Single
Dim eps As Single, fom As Single, test As String
On Error GoTo errortrap
ChDrive App.Path
ChDir App.Path & "\dicvol"
IamBusy True
'first check that the data are consistent, no negative values,...
'save then a temporary file in the app.path and then call dicvol in shell
'exe location should be at least 3 chars length
''If Len(txtExeLocation.Text) < 3 Then Err.Raise 1101, , "Input the correct location of the Dicvol program, then try again."
'the output file must be at least 1 char length
'at least one of the symmetries must be checked
s = 0
For i = 0 To 5
If chkTest(i).Value Then s = s + 1
Next i
If s = 0 Then Err.Raise 1101, , "You want indexing ? Select what symmetry to check..."
'amax, bmax,etc...should be higher than 3
For i = 0 To 2
If Abs(Val(txtmax(i))) > 0 And (Val(txtmax(i))) < 2 Then Err.Raise 1101, , "Check the maximum value for the unit cell...."
Next i
If Val(txtmax(3)) < 0 Then Err.Raise 1101, , "Unreasonable values for beta max..."
If Val(txtmax(4)) < 0 Then Err.Raise 1101, , "Unreasonable values for vol max..."
If Val(txtmin(0)) < 0 Or Val(txtmin(0)) > Val(txtmax(3)) Then Err.Raise 1101, , "Check the values of beta min/beta max."
If Val(txtmin(1)) < 0 Or Val(txtmin(1)) > Val(txtmax(4)) Then Err.Raise 1101, , "Check the values of vol min/vol max."
If Val(Text2.Text) < 0# Or Val(Text2.Text) > 0.25 Then Err.Raise 1101, , "Don't play around..." & vbCrLf & "Check the EPS value. If you really have that 2theta error maybe it is worth to find another job..."
If Val(txtDetails(0)) < 0 Then Err.Raise 1101, , "Great! You found the phlogistic ! "
If Val(txtDetails(1)) < 0 Then Err.Raise 1101, , "oops: you used a picnometer to get that value ?? "
If Val(txtDetails(2)) < 0 Or Val(txtDetails(2)) > Val(txtDetails(1).Text) Then Err.Raise 1101, , "The density error is bigger than the density ? "
If Val(txtDetails(3)) < 5 Then txtDetails(3).Text = "5.0"
'if eps is required then I should read them in column 5 of unitcell form
'make a file compatible with dicvol, named
'_PwdDicvol.inp, located in the directory of dicvol, if that file exists it will be deleted
outfil = FreeFile
'itype is the kind of data available
'1 is theta deg
'2 is 2 theta
'3 is d-dpace in angst
'4 is q; not available in powder

'1 and 2 must be in increasing order, 3 is d in decresing order
'make order and define all other parameters before opening and saving the file
'n is the number of lines used...
'call numbering routine to find out this
N = number_of_lines_for_indexing
If N = 0 Then
Err.Raise 1101, , "I need some data though...Insert the data in the 4th column of the grid in PowderCell window."
Else
If N < 1 Then Err.Raise 1101, , "Error encountered...Check the data."
End If
'establish itype based on powdercell
frmRefine.grid.Col = 4
frmRefine.grid.Row = 0
N = N - 1
Select Case left$(frmRefine.grid.Text, 1)
Case "d"
itype = 3
Case "t"
itype = 1
Case "2"
itype = 2
Case Else
'not available
Err.Raise 1101, , "Check the data type in PowderCell (allowed: theta, 2 theta or d)."
End Select

'establish the values for jc,..etc
If chkTest(0).Value Then jc = 1
If chkTest(1).Value Then jt = 1
If chkTest(2).Value Then jh = 1
If chkTest(3).Value Then jo = 1
If chkTest(4).Value Then jm = 1
If chkTest(5).Value Then jtr = 1
'see amax, bmax, etc without checking here
amax = Val(txtmax(0).Text)
bmax = Val(txtmax(1).Text)
cmax = Val(txtmax(2).Text)
volmin = Val(txtmin(1).Text)
bemin = Val(txtmin(0).Text)
volmax = Val(txtmax(4).Text)
bemax = Val(txtmax(3).Text)

'see wave, poimol, dens, deldens
WAVE = Val(frmRefine.txt.Text)
poimol = Val(txtDetails(0).Text)
DENS = Val(txtDetails(1).Text)
delden = Val(txtDetails(2).Text)
'eps and fom
Dim epscode As Boolean 'if true read the fifth column as well
fom = Val(txtDetails(3).Text)
If optError(0).Value Then eps = 0
If optError(1).Value Then epscode = True
If optError(2).Value Then eps = Val(Text2.Text)
Open "_Pwddicv.bad" For Output As outfil
Print #outfil, title
'here I have to print itype
Dim nn As Integer
nn = N
If N > 40 Then nn = 40
Print #outfil, CStr(nn) & "   " & CStr(itype) & "   " & CStr(jc) & "   " & CStr(jt) & "   " & CStr(jh) & "   " & CStr(jo) & "   " & CStr(jm) & "   " & CStr(jtr)
Print #outfil, Format$(Format$((amax), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((bmax), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((cmax), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((volmin), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((volmax), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((bemin), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((bemax), "##0.0000"), "@@@@@@@@")
Print #outfil, Format$(Format$((WAVE), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((poimol), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((DENS), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((delden), "##0.0000"), "@@@@@@@@")
Print #outfil, Format$(Format$((eps), "##0.0000"), "@@@@@@@@") & "   " & Format$(Format$((fom), "##0.0000"), "@@@@@@@@")
'to print data in itself

frmRefine.grid.Col = 4

For i = 1 To N
test = ""
frmRefine.grid.Row = i
frmRefine.grid.Col = 4
test = Format$(CStr(frmRefine.grid.Text), "##0.0000")
If epscode Then
frmRefine.grid.Col = 5
test = test + "  " + Format$(CStr(frmRefine.grid.Text), "##0.0000")
frmRefine.grid.Col = 4
End If
Print #outfil, test
DoEvents
Next i
'seems to be ok by now
Close

DoEvents
Dim inpfil As Integer
inpfil = FreeFile
Dim newlinie As String, linie As String
Dim carrier As String
'clean the file
Open "_PwdDicv.bad" For Input As inpfil
outfil = FreeFile
Open "pwd_dicvol91.in" For Output As outfil
Do While Not (EOF(inpfil))
DoEvents
Line Input #inpfil, linie
newlinie = ""
For i = 1 To Len(linie)
''MsgBox "am citit:" & Mid$(linie, i, 1)
''de lucrat pe aici
carrier = Mid$(linie, i, 1)
If ((Asc(carrier) > 126) Or (Asc(carrier) < 32)) Then carrier = " "
newlinie = newlinie + carrier
Next i
If Len(linie) > 0 Then Print #outfil, newlinie
Loop
Close #inpfil
Close #outfil
IamBusy False
Kill "_PwdDicv.bad"
'give it a try
''Shell CStr((txtExeLocation.Text) & "dicvol91.exe" & " _PwdDicv.inp" & " " & Text1.Text)
''raport "Dicvol started...check back later when finished..."
raport "A Datafile for DICVOL has been created; the name of this file is pwd_dicvol91.in"
Exit Sub
errortrap:
Close
IamBusy False
'Kill "_PwdDicv.bad"
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub Command1_Click()
ChDir App.Path
Shell "write.exe pwd_Dicvol91.out", vbMaximizedFocus
Exit Sub
End Sub

Private Sub Form_Load()
'disable quit cell
''txtExeLocation.Text = DicvolLocationDirectory
title = InputBox("Set a title for this run: ", prog_name, "")
frmRefine.mnuQuit.Enabled = False
'sort data in the grid
frmRefine.mnuSortData_Click
title = prog_name & " - " & CStr(Now) & " - " & left$(title, 30)
''C  CARD 1  TITLE                                       FORMAT(16A4)
''c
''c
''C  CARD 2  N,ITYPE,JC,JT,JH,JO,JM,JTR                  FREE FORMAT
''c
''C          N               NUMBER OF LINES USED.
''C          ITYPE           SPACING DATA TYPE.
''C                      =1  THETA BRAGG IN DEGREES.
''C                      =2  2-THETA ANGLE IN DEGREES.
''C                      =3  D-SPACING IN ANGSTROMS.
''C                      =4  Q SPECIFIED IN Q-UNITS AS E+04/D**2.
''C          JC          =0  CUBIC SYSTEM IS NOT TESTED.
''C                      =1  CUBIC SYSTEM IS TESTED.
''C          JT          =0  TETRAGONAL SYSTEM IS NOT TESTED.
''C                      =1  TETRAGONAL SYSTEM IS TESTED.
''C          JH          =0  HEXAGONAL SYSTEM IS NOT TESTED.
''C                      =1  HEXAGONAL SYSTEM IS TESTED.
''C          JO          =0  ORTHORHOMBIC SYSTEM IS NOT TESTED.
''C                      =1  ORTHORHOMBIC SYSTEM IS TESTED.
''C          JM          =0  MONOCLINIC SYSTEM IS NOT TESTED.
''C                      =1  MONOCLINIC SYSTEM IS TESTED.
''C          JTR         =0  TRICLINIC SYSTEM IS NOT TESTED.
''C                      =1  TRICLINIC SYSTEM IS TESTED.
''c
''c
''C  CARD 3  AMAX,BMAX,CMAX,VOLMIN,VOLMAX,BEMIN,BEMAX    FREE FORMAT
''c
''C          AMAX    MAXIMUM VALUE OF UNIT CELL DIMENSION A IN ANGSTROMS.
''C                  (IF AMAX= 0.0 DEFAULT= 20. ANGSTROMS)
''C          BMAX    MAXIMUM VALUE OF UNIT CELL DIMENSION B IN ANGSTROMS.
''C                  (IF BMAX= 0.0 DEFAULT= 20. ANGSTROMS)
''C          CMAX    MAXIMUM VALUE OF UNIT CELL DIMENSION C IN ANGSTROMS.
''C                  (IF CMAX= 0.0 DEFAULT= 20. ANGSTROMS)
''C          VOLMIN  MINIMUM VOLUME FOR TRIAL UNIT CELLS IN ANGSTROMS**3.
''C          VOLMAX  MAXIMUM VOLUME FOR TRIAL UNIT CELLS IN ANGSTROMS**3.
''C                  (IF VOLMAX= 0.0 DEFAULT= 1500. ANGSTROMS**3)
''C          BEMIN   MINIMUM ANGLE FOR UNIT CELL IN DEGREES
''C                  (IF BEMIN= 0.0 DEFAULT= 90. DEGREES).
''C          BEMAX   MAXIMUM ANGLE FOR UNIT CELL IN DEGREES
''C                  (IF BEMAX= 0.0 DEFAULT= 125. DEGREES).
''c
''c
''C  CARD 4  WAVE,POIMOL,DENS,DELDEN                     FREE FORMAT
''c
''C          WAVE    WAVELENGTH IN ANGSTROMS (DEFAULT=0.0 IF CU K ALPHA1).
''C          POIMOL  MOLECULAR WEIGHT OF ONE FORMULA UNIT IN A.M.U.
''C                  (DEFAULT =0.0 IF FORMULA WEIGHT NOT KNOWN).
''C          DENS    MEASURED DENSITY IN G.CM(-3)
''C                  (DEFAULT =0.0 IF DENSITY NOT KNOWN).
''C          DELDEN  ABSOLUTE ERROR IN MEASURED DENSITY.
''c
''c
''C  CARD 5  EPS,FOM                                     FREE FORMAT
''c
''C          EPS       =0.0  THE ABSOLUTE ERROR ON EACH OBSERVED LINE
''C                          IS TAKEN TO .03 DEG. 2THETA, WHATEVER THE
''C                          SPACING DATA TYPE (ITYPE IN CARD 2).
''C                    =1.0  THE ABSOLUTE ERROR ON EACH OBSERVED LINE IS
''C                          INPUT INDIVIDUALLY IN THE FOLLOWING CARDS,
''C                          TOGETHER WITH THE OBSERVED 'D(I)', ACCORDING
''C                          WITH THE SPACING DATA UNIT.
''C                    EPS NE 0.0 AND 1.0
''C                          THE ABSOLUTE ERROR IS TAKEN AS A CONSTANT
''C                          (=EPS),IN DEG. 2THETA, WHATEVER THE SPACING
''C                          DATA TYPE (ITYPE IN CARD 2).
''C          FOM             LOWER FIGURE OF MERIT M(N) REQUIRED FOR PRINTED
''C                          SOLUTION(S) (DEFAULT=0.0 IF LOWER M(N)=5.0).
''c
''c
''C  CARD 6 TO 6+N  D(I),EPSIL(I)                        FREE FORMAT
''c
''C          (ONE FOR EACH OBSERVED LINE, UP TO N)
''C          D(I)    VALUE DESCRIBING THE OBSERVED POSITION
''C                  OF THIS LINE ACCORDING TO 'ITYPE'.
''C          EPSIL   ABSOLUTE ERROR IN 'D(I)', ACCORDING TO 'ITYPE',
''C                  ONLY IF EPS=1.0 (CARD 5).
''c NOTE:
''C          IF ITYPE=1,2,4 THE VALUES OF 'D(I)' AND 'EPSIL(I)' MUST
''C          BE LISTED IN INCREASING ORDER.
''C          IF ITYPE=3 THEY MUST BE IN DECREASING ORDER.
''c


End Sub

Sub Form_Unload(Cancel As Integer)
'here to make all checkings before closing
frmRefine.mnuQuit.Enabled = True
Unload Me
End Sub



Private Sub optError_Click(Index As Integer)
Text2.Visible = False
If Index = 2 Then Text2.Visible = True
End Sub

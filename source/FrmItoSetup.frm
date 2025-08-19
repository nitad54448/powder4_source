VERSION 5.00
Begin VB.Form FrmItoSetup 
   Caption         =   "ITO-Setup"
   ClientHeight    =   4990
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4990
   ScaleWidth      =   8290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Editor"
      Height          =   372
      Left            =   4440
      TabIndex        =   26
      Top             =   4200
      Width           =   1452
   End
   Begin VB.CommandButton btnSee 
      Caption         =   "Run ITO"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "&Make file"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "ITO - Properties "
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.TextBox txtMinNrSol 
         Height          =   285
         Left            =   6720
         TabIndex        =   20
         Text            =   "14"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtTol 
         Height          =   285
         Index           =   2
         Left            =   6720
         TabIndex        =   19
         Text            =   "6"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox FOMmin 
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Text            =   "4.0"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   17
         Text            =   "0.0"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   16
         Text            =   "0.0"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CheckBox chkTestOutput 
         Caption         =   "test output"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkZeroError 
         Caption         =   "Zero error check"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtTol 
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   11
         Text            =   "4.5"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtTol 
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   10
         Text            =   "3.0"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkSymmetry 
         Caption         =   "check triclinic"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkSymmetry 
         Caption         =   "check monoclinic"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkSymmetry 
         Caption         =   "check orthorombic"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtSolPrint 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Text            =   "4"
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox chkreadInt 
         Caption         =   "Read intensities (col. 5 in PowderCell)"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   3120
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Measured density"
         Height          =   375
         Index           =   4
         Left            =   4680
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mol. weight"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   24
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Minim nr. of indexed lines"
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   23
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tolerance on obs/calc. lines"
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   22
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Minim F of M to print"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   21
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblTol 
         Alignment       =   1  'Right Justify
         Caption         =   "3 dimensional range search"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   13
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblTol 
         Alignment       =   1  'Right Justify
         Caption         =   "2dimensional range search"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblSolPrint 
         Caption         =   "Number of solutions to print"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmItoSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub btnCancel_Click()
frmRefine.mnuQuit.Enabled = True
Form_Unload (-1)
End Sub

Sub btnRun_Click()
On Error GoTo errortrap
ChDir App.Path & "\ito"
Me.MousePointer = 11
Dim man As String * 1, instr As String * 1, intens As String * 1, nsolmx As Integer, nsyst(3) As String
Dim tol2 As Single, tol3 As Single, wavel As Single, linco As Integer
Dim lzerck As Integer, wmol As Single, dobs As Single, tolg As Single, prntmr As Single, prntln As Single
'determines number of lines,,,
Dim N As Integer, i As Integer, outfil As Integer
frmRefine.grid.Col = 4
frmRefine.grid.Row = 0
If left$(frmRefine.grid.Text, 1) = "t" Then Err.Raise 1101, , "The data are as THETA. ITO does not accept theta values as input. Go in the PowderCell routine and convert them to d or 2theta. Use Data/Change data To: command."
'sort routine
''''grid_sort
outfil = FreeFile
'the output file must be at least 1 char length
N = number_of_lines_for_indexing
If N = 0 Then
Err.Raise 1101, , "data not found...Insert data in the 4th column of the grid in PowderCell window."
Else
If N < 20 Then Err.Raise 1101, , "You need at least 20 lines for ITO...."
End If
'establish itype based on powdercell
'see wave, poimol, dens, deldens
wavel = Val(frmRefine.txt.Text)
wmol = Val(txtDetails(0).Text)
dobs = Val(txtDetails(1).Text)
Dim intenscode As Boolean 'if true read the fifth column as well
intens = 0
If chkreadInt.Value Then intens = 1: intenscode = True
prntmr = Val(FOMmin.Text)
For i = 0 To 2
nsyst(i + 1) = "-1"
If chkSymmetry(i).Value Then nsyst(i + 1) = "+1"
Next i
lzerck = 0: If chkZeroError.Value Then lzerck = 1
Dim testout As Integer
testout = 0: If chkTestOutput.Value Then testout = 1
Open "itoinp.bad" For Output As outfil
Print #outfil, title
'here I have to print a lot of stuff
Dim test As String
'convert to 80 chars
If intens = 0 Then intens = " "
Dim wm As String, dobss As String
wm = Format$(Format$(Val(wmol), "####.00000"), "@@@@@@@@@@")
dobss = Format$(Format$(Val(dobs), "####.00000"), "@@@@@@@@@@")
If dobs <= 0 Then dobss = "          "
If wmol <= 0 Then wm = "          "
test = " 1" & CStr(intens) & Format$(Format$(txtSolPrint, "#"), "@") & nsyst(1) & nsyst(2) & nsyst(3) & Format$(Format$(Val(txtTol(0).Text), "##.00"), "@@@@@") & Format$(Format$(Val(txtTol(1).Text), "##.00"), "@@@@@") & Format$(Format$(Val(wavel), "###0.00000"), "@@@@@@@@@@") & "  " & CStr(lzerck) & "              " & "  " & CStr(testout) & wm & dobss & Format$(Format$(Val(txtTol(2).Text), "###.0000"), "@@@@@@@@")
Print #outfil, test
test = "    0.0000" & Format$(Format$(Val(FOMmin.Text), "########.0"), "@@@@@@@@@@") & Format$(Format$(Val(txtMinNrSol.Text), "#########."), "@@@@@@@@@@")
Print #outfil, test
'to print data in itself
frmRefine.grid.Col = 4

For i = 1 To N
test = ""
frmRefine.grid.Row = i
frmRefine.grid.Col = 4
test = Format$((frmRefine.grid.Text), "##0.0000")
If intenscode Then
frmRefine.grid.Col = 5
test = test + " " + Format$(Format$((frmRefine.grid.Text), "##0.00"), "@@@@@@")
frmRefine.grid.Col = 4
End If
Print #outfil, test
Next i
Print #outfil, "0.0"
Print #outfil, "END"
'seems to be ok by now
Close #outfil

Dim inpfil As Integer
inpfil = FreeFile
Dim newlinie As String, linie As String
Dim carrier As String
'clean the file
Open "itoinp.bad" For Input As inpfil
outfil = FreeFile
Open "itoinp.dat" For Output As outfil
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
Kill "itoinp.bad"
'give it a try
DoEvents
''
raport "A Datafile for ITO has been created; the name of this file is itoinp.dat"
raport "You might need to adjust more parameters for ITO. "
Me.MousePointer = 0

Exit Sub
errortrap:
Me.MousePointer = 0
MsgBox Err.Description
Err.Clear
Close
Exit Sub
End Sub

Private Sub btnSee_Click()
On Error GoTo errortrap
ChDir (App.Path & "\ito")
ChDrive App.Path
Dim j As Integer
Dim t As Double
    Call btnRun_Click
t = ShellAndLoop("pwd_ito.exe", vbMaximizedFocus)
DoEvents
    If t = 0 Then
        raport "the program pwd_ito was called."
    Else
        raport "pwd_Ito" & " has been started." & vbCrLf & " Main window handle: " & Hex(t)
        raport strLinie
    End If
DoEvents
    hWndShell "write.exe itout.lst", vbNormalFocus
't = Shell("pwd_dicvol91.exe pwd_dicvol91.in", vbNormalFocus)
't = Shell("pwd_ito.exe", vbNormalFocus)
    If t = 0 Then Err.Raise 111111, , "can not call pwd_ito.exe"
Exit Sub
errortrap:
    raport strLinie
    raport "error in btnSee Routine...ito call failed..."
    raport "you can manually run the program ITO by using the file itoinp.dat saved in the current working directory (i.e. " & CurDir & ") "
    raport Err.Description
    Err.Clear
Exit Sub
End Sub

Private Sub Command1_Click()
Shell "write.exe itout.lst", vbMaximizedFocus
Exit Sub

End Sub

Private Sub Form_Load()
'disable quit cell
title = InputBox("Set a title for this run: ", prog_name, "")
frmRefine.mnuQuit.Enabled = False
title = prog_name & " - " & CStr(Now) & " - " & left$(title, 30)
'sort the data first
frmRefine.mnuSortData_Click
End Sub

Sub Form_Unload(Cancel As Integer)
'here to make all checkings before closing
frmRefine.mnuQuit.Enabled = True
Unload Me

End Sub

Private Sub txtExeLocation_Change()
End Sub

VERSION 5.00
Begin VB.Form TreorSetup 
   Caption         =   "Treor 90 - Setup"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   10090
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Editor"
      Height          =   372
      Left            =   5760
      TabIndex        =   75
      Top             =   6840
      Width           =   1332
   End
   Begin VB.Frame Frame2 
      Caption         =   "Treor90 - Properties "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6492
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9972
      Begin VB.TextBox TXTNIX 
         Height          =   285
         Left            =   360
         TabIndex        =   73
         Text            =   "2"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtTolSSQ 
         Height          =   285
         Left            =   9000
         TabIndex        =   71
         Text            =   "0.05"
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtTolD2 
         Height          =   285
         Left            =   9000
         TabIndex        =   70
         Text            =   "0.0004"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtTolD1 
         Height          =   285
         Left            =   9000
         TabIndex        =   69
         Text            =   "0.0002"
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox ChkShort 
         Caption         =   "short axis test (SHORT)"
         Height          =   375
         Left            =   360
         TabIndex        =   68
         Top             =   5640
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.CheckBox chkIDIV 
         Caption         =   "adjust by higher order lines (IDIV)"
         Height          =   375
         Left            =   360
         TabIndex        =   67
         Top             =   6000
         Value           =   1  'Checked
         Width           =   4575
      End
      Begin VB.TextBox txtSelect 
         Height          =   285
         Left            =   9000
         TabIndex        =   66
         Text            =   "0"
         Top             =   6000
         Width           =   735
      End
      Begin VB.CheckBox chkList 
         Caption         =   "print cells for refinement (LIST)"
         Height          =   375
         Left            =   360
         TabIndex        =   65
         Top             =   5280
         Width           =   4095
      End
      Begin VB.TextBox txtIQ 
         Height          =   285
         Left            =   9000
         TabIndex        =   64
         Text            =   "16"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtUse 
         Height          =   285
         Left            =   9000
         TabIndex        =   63
         Text            =   "19"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtbeta 
         Height          =   285
         Left            =   9000
         TabIndex        =   62
         Text            =   "0"
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkmonogam 
         Caption         =   "Refine trial sets (monogam)"
         Height          =   375
         Left            =   360
         TabIndex        =   61
         Top             =   4920
         Width           =   5295
      End
      Begin VB.TextBox txtMonoset 
         Height          =   285
         Left            =   9000
         TabIndex        =   60
         Text            =   "0"
         Top             =   5640
         Width           =   735
      End
      Begin VB.Frame frm 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   4320
         TabIndex        =   56
         Top             =   2520
         Width           =   2172
         Begin VB.OptionButton optMon 
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   72
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton optMon 
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton optMon 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optMon 
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   58
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox txtHKL 
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   42
         Text            =   "3"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtHKL 
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   41
         Text            =   "3"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtHKL 
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   40
         Text            =   "4"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtHKL 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   39
         Text            =   "4"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtHKL 
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   38
         Text            =   "6"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Index           =   4
         Left            =   3120
         TabIndex        =   37
         Text            =   "2"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   36
         Text            =   "2"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   35
         Text            =   "4"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   34
         Text            =   "4"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   33
         Text            =   "4"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   32
         Text            =   "2"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   31
         Text            =   "2"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   30
         Text            =   "4"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   29
         Text            =   "4"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   28
         Text            =   "4"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   27
         Text            =   "2"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   26
         Text            =   "2"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   25
         Text            =   "4"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   24
         Text            =   "4"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   23
         Text            =   "4"
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox chkSymmetry 
         Alignment       =   1  'Right Justify
         Caption         =   "triclinic"
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   21
         Top             =   3120
         Width           =   1332
      End
      Begin VB.CheckBox chkSymmetry 
         Alignment       =   1  'Right Justify
         Caption         =   "orthorhombic"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1452
      End
      Begin VB.CheckBox chkSymmetry 
         Alignment       =   1  'Right Justify
         Caption         =   "monoclinic"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1452
      End
      Begin VB.CheckBox chkSymmetry 
         Alignment       =   1  'Right Justify
         Caption         =   "cubic"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   1452
      End
      Begin VB.CheckBox chkSymmetry 
         Alignment       =   1  'Right Justify
         Caption         =   "tetragonal"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1452
      End
      Begin VB.CheckBox chkSymmetry 
         Alignment       =   1  'Right Justify
         Caption         =   "hexagonal"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.TextBox txtVOL 
         Height          =   285
         Left            =   9000
         TabIndex        =   9
         Text            =   "750"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtCEM 
         Height          =   285
         Left            =   9000
         TabIndex        =   8
         Text            =   "25"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtMolW 
         Height          =   285
         Left            =   9000
         TabIndex        =   7
         Text            =   "0"
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtDens 
         Height          =   285
         Left            =   9000
         TabIndex        =   6
         Text            =   "0"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox FOMmin 
         Height          =   285
         Left            =   9000
         TabIndex        =   5
         Text            =   "10"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtTolDens 
         Height          =   285
         Left            =   9000
         TabIndex        =   4
         Text            =   "0"
         Top             =   5160
         Width           =   735
      End
      Begin VB.Frame frm 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   4320
         TabIndex        =   52
         Top             =   2040
         Width           =   1695
         Begin VB.OptionButton optOrt 
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   55
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton optOrt 
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   54
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton optOrt 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Label Label4 
         Caption         =   "max. number of not indexable lines"
         Height          =   375
         Left            =   1320
         TabIndex        =   74
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "base line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4440
         TabIndex        =   43
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "use lines"
         Height          =   252
         Index           =   13
         Left            =   7320
         TabIndex        =   51
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "lines for lst. sq. (IQ)"
         Height          =   252
         Index           =   12
         Left            =   5760
         TabIndex        =   50
         Top             =   1080
         Width           =   3132
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "tolerance in SSQTL"
         Height          =   252
         Index           =   11
         Left            =   5640
         TabIndex        =   49
         Top             =   3960
         Width           =   3252
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "tolerance in D2"
         Height          =   372
         Index           =   10
         Left            =   6600
         TabIndex        =   48
         Top             =   3600
         Width           =   2292
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "tolerance in D1"
         Height          =   372
         Index           =   9
         Left            =   5640
         TabIndex        =   47
         Top             =   3240
         Width           =   3252
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "select base lines (SELECT)"
         Height          =   372
         Index           =   6
         Left            =   4800
         TabIndex        =   46
         Top             =   6000
         Width           =   4092
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "max beta cell (MONO)"
         Height          =   372
         Index           =   2
         Left            =   6360
         TabIndex        =   45
         Top             =   2280
         Width           =   2532
      End
      Begin VB.Label lbldetalii 
         Alignment       =   1  'Right Justify
         Caption         =   "select base lines for monoclinic (Monoset)"
         Height          =   372
         Index           =   0
         Left            =   4440
         TabIndex        =   44
         Top             =   5640
         Width           =   4452
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "max H, K, L and max H+K+L base lines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblTol 
         Alignment       =   1  'Right Justify
         Caption         =   "maximum volume"
         Height          =   372
         Index           =   0
         Left            =   6120
         TabIndex        =   18
         Top             =   1920
         Width           =   2772
      End
      Begin VB.Label lblTol 
         Alignment       =   1  'Right Justify
         Caption         =   "maximum cell edge"
         Height          =   372
         Index           =   1
         Left            =   6000
         TabIndex        =   17
         Top             =   1560
         Width           =   2892
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "minimum Figure of Merit"
         Height          =   372
         Index           =   0
         Left            =   6480
         TabIndex        =   16
         Top             =   2760
         Width           =   2412
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "density tolerance (EDENS)"
         Height          =   372
         Index           =   1
         Left            =   5040
         TabIndex        =   15
         Top             =   5160
         Width           =   3852
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "molecular weight"
         Height          =   372
         Index           =   3
         Left            =   6000
         TabIndex        =   14
         Top             =   4440
         Width           =   2892
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "measured density"
         Height          =   372
         Index           =   4
         Left            =   5760
         TabIndex        =   13
         Top             =   4800
         Width           =   3132
      End
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "&Make file"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton btnSee 
      Caption         =   "&Run TREOR"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   6840
      Width           =   1455
   End
End
Attribute VB_Name = "TreorSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cub As BaseLine, tet As BaseLine, ort(3) As BaseLine, mon(4) As BaseLine, tri As BaseLine

Option Explicit



Private Sub btnBrowse1_Click()

End Sub

Private Sub btnCancel_Click()
frmRefine.mnuQuit.Enabled = True
Form_Unload (-1)

End Sub

Private Sub btnRun_Click()
On Error GoTo errortrap
ChDir App.Path & "\treor"
ChDrive App.Path
IamBusy True
Dim N As Integer, i As Integer, outfil As Integer
Dim CHOICE As Integer, KH As Integer, KK As Integer, KL As Integer, KS As Integer, THH As Integer, THK As Integer, THL As Integer, THS As Integer
Dim OH(3) As Integer, OK(3) As Integer, OL(3) As Integer, OS(3) As Integer
Dim MH(4) As Integer, MK(4) As Integer, ml(4) As Integer, MS(4) As Integer
Dim MONOSET As Integer, MONOGAM As Integer, MONO As Integer, SHORT As Integer, USE As Integer, IQ As Integer, tLIST As Integer, tSELECT As Integer, MERIT As Integer, NIX As Integer
Dim IDIV As Integer, WAVE As Single, VOL As Single, CEM As Single, D1 As Single, D2 As Single, SSQTL As Single, DENS As Single, EDENS As Single, MOLW As Single, TRIC As Single
frmRefine.grid.Col = 4
frmRefine.grid.Row = 0
If left$(frmRefine.grid.Text, 1) = "t" Then CHOICE = 2
If left$(frmRefine.grid.Text, 1) = "2" Then CHOICE = 3
If left$(frmRefine.grid.Text, 1) = "d" Then CHOICE = 4

outfil = FreeFile

'the output file must be at least 1 char length
N = number_of_lines_for_indexing
If N = 0 Then Err.Raise 1101, , "data not found...Insert data in the 4th column of the grid in PowderCell window."
'see wave, poimol, dens, deldens
WAVE = Val(frmRefine.txt.Text)
If chkmonogam.Value Then MONOGAM = 1
If chkList.Value Then tLIST = 1
If ChkShort.Value Then SHORT = 1
If chkIDIV.Value Then IDIV = 1
USE = CInt(txtUse.Text)
IQ = CInt(txtIQ.Text)
CEM = Val(txtCEM.Text)
VOL = Val(txtVOL.Text)
MONO = Val(txtbeta.Text)
MERIT = Val(FOMmin.Text)
D1 = Val(txtTolD1.Text)
D2 = Val(txtTolD2.Text)
SSQTL = Val(txtTolSSQ.Text)
MOLW = Val(txtMolW.Text)
DENS = Val(txtDens.Text)
EDENS = Val(txtTolDens.Text)
MONOSET = CInt(txtMonoset.Text)
tSELECT = CInt(txtSelect.Text)

Open "pw_treor.in" For Output As outfil
Print #outfil, title
'here I have to print a lot of stuff
frmRefine.grid.Col = 4
For i = 1 To N
frmRefine.grid.Row = i
frmRefine.grid.Col = 4
Print #outfil, Format$((frmRefine.grid.Text), "##0.0000")
Next i
Print #outfil, ""
'here is the all other stuff to print
Print #outfil, "CHOICE=" & CStr(CHOICE) & ","
If chkSymmetry(0).Value Then
Print #outfil, "KH=" & cub.Hmax & ", KK=" & cub.Kmax & ", KL=" & cub.Lmax & ", KS=" & cub.HKLmax & ", "
Else
Print #outfil, "KS=0,"
End If

If chkSymmetry(1).Value Then
Print #outfil, "THH=" & tet.Hmax & ", THK=" & tet.Kmax & ", THL=" & tet.Lmax & ", THS=" & tet.HKLmax & ", "
Else
Print #outfil, "THS=0,"
End If
If chkSymmetry(3).Value Then
Print #outfil, "OH1=" & ort(1).Hmax & ", OK1=" & ort(1).Kmax & ", OL1=" & ort(1).Lmax & ", OS1=" & ort(1).HKLmax & ", "
Print #outfil, "OH2=" & ort(2).Hmax & ", OK2=" & ort(2).Kmax & ", OL2=" & ort(2).Lmax & ", OS2=" & ort(2).HKLmax & ", "
Print #outfil, "OH3=" & ort(3).Hmax & ", OK3=" & ort(3).Kmax & ", OL3=" & ort(3).Lmax & ", OS3=" & ort(3).HKLmax & ", "
Else
Print #outfil, "OS1=0,"
End If

Print #outfil, "MH1=" & mon(1).Hmax & ", MK1=" & mon(1).Kmax & ", ML1=" & mon(1).Lmax & ", MS1=" & mon(1).HKLmax & ", "
Print #outfil, "MH2=" & mon(2).Hmax & ", MK2=" & mon(2).Kmax & ", ML2=" & mon(2).Lmax & ", MS2=" & mon(2).HKLmax & ", "
Print #outfil, "MH3=" & mon(3).Hmax & ", MK3=" & mon(3).Kmax & ", ML3=" & mon(3).Lmax & ", MS3=" & mon(3).HKLmax & ", "
Print #outfil, "MH4=" & mon(4).Hmax & ", MK4=" & mon(4).Kmax & ", ML4=" & mon(4).Lmax & ", MS4=" & mon(4).HKLmax & ", "

If Not (chkSymmetry(5).Value) Then Print #outfil, "TRIC=0,"

Print #outfil, "MONOSET=" & Val(txtMonoset.Text) & ", "
If chkmonogam.Value Then
Print #outfil, "MONOGAM=1, "
Else
Print #outfil, "MONOGAM=0, "
End If

Print #outfil, "MONO=" & CStr(txtbeta.Text) & ", "

If ChkShort.Value Then
Print #outfil, "SHORT=1, "
Else
Print #outfil, "SHORT=0, "
End If

Print #outfil, "USE=" & CStr(CInt(txtUse.Text)) & ", "
Print #outfil, "IQ=" & CStr(CInt(txtIQ.Text)) & ", "


If chkList.Value Then
Print #outfil, "LIST=1, "
Else
Print #outfil, "LIST=0, "
End If

Print #outfil, "SELECT=" & CStr(CInt(txtSelect.Text)) & ", MERIT=" & CStr(Val(FOMmin.Text)) & ", NIX=" & CStr(CInt(TXTNIX.Text)) & ", "
If chkIDIV.Value Then
Print #outfil, "IDIV=1, "
Else
Print #outfil, "IDIV=0, "
End If

If Abs(Val(frmRefine.txt.Text) - 1.5405) > 0.01 Then
Print #outfil, "WAVE=" & Val(frmRefine.txt.Text) & " ,"
End If

Print #outfil, "VOL=" & CStr(Val(txtVOL.Text)) & ",CEM=" & CStr(Val(txtCEM.Text)) & ", "
Print #outfil, "D1=" & CStr(Val(txtTolD1.Text)) & ",SSQTL=" & CStr(Val(txtTolSSQ.Text)) & ",D2=" & CStr(Val(txtTolD2.Text)) & ", "
Print #outfil, "DENS=" & CStr(Val(txtDens.Text)) & ",EDENS=" & CStr(Val(txtTolDens.Text)) & ",MOLW=" & CStr(Val(txtMolW.Text)) & ", "

Print #outfil, "END*"
'seems to be ok by now
Close #outfil
IamBusy False
raport "A Datafile for TREOR has been created; the name of this file is pw_treor.in"
raport "You might need to adjust more parameters for TREOR. "
Exit Sub

errortrap:

IamBusy False
MsgBox Err.Description
Err.Clear
Close
Exit Sub
End Sub

Private Sub btnSee_Click()
On Error GoTo errortrap
Dim t As Double, ts As String
Call btnRun_Click
On Error GoTo errortrap
MsgBox "the data file is pw_treor.in"
raport strLinie
t = ShellAndClose("pwd_treor.exe", vbMaximizedFocus)
If t = 0 Then
raport "calling the program pwd_TREOR."
Else
raport "pwd_treor " & " has been started." & vbCrLf & _
                     "Main window handle: " & Hex(t)
raport strLinie
End If
ts = InputBox("The program pw_treor finished. Please enter the name of the output file you want to see:")
If Len(ts) > 0 Then
hWndShell "write.exe " & ts, vbMaximizedFocus
End If
Exit Sub
errortrap:
raport strLinie
raport "error in btnSee Routine...TREOR call failed..."
raport "you can manually run the program TREOR by using the file pw_treor.in saved in the current working directory (i.e. " & CurDir & ") "
raport Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub Command1_Click()
ChDir App.Path
Shell "write.exe", vbMaximizedFocus
Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Integer, N As Integer
'''!!txtExeLocation.Text = TreorLocationDirectory
title = InputBox("Set a title for this run: ", prog_name, "")
frmRefine.mnuQuit.Enabled = False
title = prog_name & " - " & CStr(Now) & " - " & left$(title, 30)
'sort the data first
frmRefine.mnuSortData_Click
N = number_of_lines_for_indexing
If N > 6 Then
txtUse.Text = CInt(N)
txtIQ.Text = CInt(N - 3)
End If
cub.Hmax = 4
cub.Kmax = 4
cub.Lmax = 4
cub.HKLmax = 6
tet.Hmax = 4
tet.Kmax = 4
tet.Lmax = 4
tet.HKLmax = 4

For i = 1 To 3
ort(i).Hmax = 2
ort(i).Kmax = 2
ort(i).Lmax = 2
ort(i).HKLmax = 4
Next i
ort(1).HKLmax = 3
For i = 1 To 4
mon(i).Hmax = 2
mon(i).Kmax = 2
mon(i).Lmax = 2
mon(i).HKLmax = 3
Next i
mon(4).HKLmax = 4
Exit Sub
End Sub

Sub Form_Unload(Cancel As Integer)
'here to make all checkings before closing
frmRefine.mnuQuit.Enabled = True
Unload Me


End Sub

Private Sub optMon_Click(Index As Integer)
txtH(4).Text = mon(Index + 1).Hmax
txtK(4).Text = mon(Index + 1).Kmax
txtL(4).Text = mon(Index + 1).Lmax
txtHKL(4).Text = mon(Index + 1).HKLmax

End Sub

Private Sub optOrt_Click(Index As Integer)
txtH(3).Text = ort(Index + 1).Hmax
txtK(3).Text = ort(Index + 1).Kmax
txtL(3).Text = ort(Index + 1).Lmax
txtHKL(3).Text = ort(Index + 1).HKLmax
Exit Sub

End Sub

Private Sub txtExeLocation_Change()

End Sub

Private Sub txtH_Change(Index As Integer)
'h
Select Case Index
Case 0
cub.Hmax = Val(txtH(Index).Text)
Case 1
'never available,
Case 2
tet.Hmax = Val(txtH(Index).Text)
Case 3
If optOrt(0).Value Then ort(1).Hmax = Val(txtH(Index).Text)
If optOrt(1).Value Then ort(2).Hmax = Val(txtH(Index).Text)
If optOrt(2).Value Then ort(3).Hmax = Val(txtH(Index).Text)


Case 4
If optMon(0).Value Then mon(1).Hmax = Val(txtH(Index).Text)
If optMon(1).Value Then mon(2).Hmax = Val(txtH(Index).Text)
If optMon(2).Value Then mon(3).Hmax = Val(txtH(Index).Text)
If optMon(3).Value Then mon(4).Hmax = Val(txtH(Index).Text)



Case 5
tri.Hmax = Val(txtH(Index).Text)
End Select

Exit Sub
End Sub

Private Sub txtHKL_Change(Index As Integer)
Select Case Index
Case 0
cub.HKLmax = Val(txtHKL(Index).Text)
Case 1
'never available,
Case 2
tet.HKLmax = Val(txtHKL(Index).Text)
Case 3
If optOrt(0).Value Then ort(1).HKLmax = Val(txtHKL(Index).Text)
If optOrt(1).Value Then ort(2).HKLmax = Val(txtHKL(Index).Text)
If optOrt(2).Value Then ort(3).HKLmax = Val(txtHKL(Index).Text)
Case 4
If optMon(0).Value Then mon(1).HKLmax = Val(txtHKL(Index).Text)
If optMon(1).Value Then mon(2).HKLmax = Val(txtHKL(Index).Text)
If optMon(2).Value Then mon(3).HKLmax = Val(txtHKL(Index).Text)
If optMon(3).Value Then mon(4).HKLmax = Val(txtHKL(Index).Text)

Case 5
tri.HKLmax = Val(txtHKL(Index).Text)
End Select

End Sub

Private Sub txtK_Change(Index As Integer)
Select Case Index
Case 0
cub.Kmax = Val(txtK(Index).Text)
Case 1
'never available,
Case 2
tet.Kmax = Val(txtK(Index).Text)
Case 3
If optOrt(0).Value Then ort(1).Kmax = Val(txtK(Index).Text)
If optOrt(1).Value Then ort(2).Kmax = Val(txtK(Index).Text)
If optOrt(2).Value Then ort(3).Kmax = Val(txtK(Index).Text)

Case 4
If optMon(0).Value Then mon(1).Kmax = Val(txtK(Index).Text)
If optMon(1).Value Then mon(2).Kmax = Val(txtK(Index).Text)
If optMon(2).Value Then mon(3).Kmax = Val(txtK(Index).Text)
If optMon(3).Value Then mon(4).Kmax = Val(txtK(Index).Text)


Case 5
tri.Kmax = Val(txtK(Index).Text)
End Select

End Sub

Private Sub txtL_Change(Index As Integer)
Select Case Index
Case 0
cub.Lmax = Val(txtL(Index).Text)
Case 1
'never available,
Case 2
tet.Lmax = Val(txtL(Index).Text)
Case 3
If optOrt(0).Value Then ort(1).Lmax = Val(txtL(Index).Text)
If optOrt(1).Value Then ort(2).Lmax = Val(txtL(Index).Text)
If optOrt(2).Value Then ort(3).Lmax = Val(txtL(Index).Text)

Case 4
If optMon(0).Value Then mon(1).Lmax = Val(txtL(Index).Text)
If optMon(1).Value Then mon(2).Lmax = Val(txtL(Index).Text)
If optMon(2).Value Then mon(3).Lmax = Val(txtL(Index).Text)
If optMon(3).Value Then mon(4).Lmax = Val(txtL(Index).Text)

Case 5
tri.Lmax = Val(txtL(Index).Text)
End Select

End Sub

VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3010
   ClientLeft      =   410
   ClientTop       =   690
   ClientWidth     =   4520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3010
   ScaleWidth      =   4520
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblCompanyProduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Powder4  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   220
      Left            =   240
      TabIndex        =   2
      Tag             =   "CompanyProduct"
      Top             =   240
      Width           =   4080
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1568
      Index           =   2
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "click to close"
      Top             =   1200
      Width           =   4088
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   248
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3968
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAboutOK_Click()
End Sub

Private Sub Form_Activate()
If (Len(prog_name) < 2) Then Load Convert3Main
Exit Sub
End Sub

Private Sub Form_Click()
Unload Me
Convert3Main.Show
Exit Sub
End Sub

Private Sub Form_Load()
lblAbout(1).Caption = version_name
lblAbout(2).Caption = "The author of this program is:" & vbCrLf & vbCrLf & " N. Dragoe" & vbCrLf & "Universite Paris Sud, " & vbCrLf & " LPCES - UMR 8648 CNRS, Bat. 414 " & vbCrLf & " Orsay 91405, France"
Me.Refresh
DoEvents
End Sub


Private Sub lblAbout_Click(Index As Integer)
Unload Me
Convert3Main.Show
DoEvents
Exit Sub
End Sub

Private Sub Picture1_Click()
Unload Me
Convert3Main.Show
DoEvents
Exit Sub
End Sub


VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2232
   ClientLeft      =   2968
   ClientTop       =   1512
   ClientWidth     =   2816
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2232
   ScaleWidth      =   2816
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label lblCompanyProduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Powder v4 beta 01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   208
      Left            =   240
      TabIndex        =   0
      Tag             =   "CompanyProduct"
      Top             =   720
      Width           =   2312
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   '' lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   '' lblProductName.Caption = App.title
'lblAuthor.Caption = "The author of this program is: " & vbCrLf & vbCrLf & "Dr. N. Dragoe " & vbCrLf & "University of Bucharest, Faculty of Chemistry" & vbCrLf & "Department of Physical-Chemistry" & vbCrLf & "Bucharest, Romania"
lblCompanyProduct = "Powder v3 beta 01" & vbCrLf & "X-rays tools"
DoEvents
End Sub



Private Sub timerSplash_Timer()

MsgBox "5 sec"
End Sub

Private Sub lblCompanyProduct_Click()

End Sub

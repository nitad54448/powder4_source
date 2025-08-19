VERSION 5.00
Begin VB.Form ReadCustom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom"
   ClientHeight    =   2950
   ClientLeft      =   30
   ClientTop       =   260
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2950
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Annuler 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   370
      Left            =   2520
      TabIndex        =   6
      Top             =   2160
      Width           =   970
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   370
      Left            =   3960
      TabIndex        =   5
      Top             =   2160
      Width           =   1090
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Text            =   "1"
      Top             =   1200
      Width           =   610
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Text            =   "1"
      Top             =   480
      Width           =   610
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "point(s)"
      Height          =   250
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   1450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "and skip every "
      Height          =   370
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Read point no."
      Height          =   370
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1570
   End
End
Attribute VB_Name = "ReadCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Annuler_Click()
Unload Me
End Sub

Private Sub OK_Click()
Dim return_code As Boolean, nr_linii As Long, i As Long
Dim linie As String, valCustom As Double, primPoint As Integer, skipPoint As Integer
Dim j As Long, k As Long, dlinie As Double

primPoint = CInt(Text1(0).Text)
skipPoint = CInt(Text1(1).Text)
If (primPoint < 0) Or (CInt(Text1(0).Text) < 0) Then
MsgBox ("Negative values ? try again..."): Exit Sub
End If


amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errortrap
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie

raport inputfile & " open; custom file read.  Read point " & CStr(CInt(Text1(0).Text)) & " and every " & CStr(CInt(Text1(1).Text)) & " points thereafter."
'Call numar_linii(inputfile, nr_linii)
'raport "This file has " & CStr(nr_linii) & " lines."

Open inputfile For Input As inpfil
numarvalori = 0
raport CStr(ignoralinii) & " line(s) will be ignored."
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
Do While Not (EOF(inpfil))
Input #inpfil, dlinie
numarvalori = numarvalori + 1
Loop
Close
raport "this file has " & CStr(numarvalori) & " points."
numarvalori = numarvalori / (1 + skipPoint)
ReDim Y(numarvalori)
amfullxdata = False: amydata = True
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If

k = 0
For i = 1 To primPoint
Input #inpfil, Y(1)
Next i
k = k + 1

For i = 1 To numarvalori
For j = 1 To skipPoint + 1
Input #inpfil, dlinie
Next j
k = k + 1
Y(k) = dlinie
Next i
numarvalori = k

Close
raport "Read Custom done..." & vbCrLf & Now
Unload Me
Exit Sub
errortrap:
Err.Clear
numarvalori = k
raport CStr(numarvalori) & " points read."

Close
Unload Me
Exit Sub



























End Sub

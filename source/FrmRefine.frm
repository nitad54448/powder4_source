VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRefine 
   Caption         =   "PowderCell"
   ClientHeight    =   4760
   ClientLeft      =   830
   ClientTop       =   1480
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4760
   ScaleWidth      =   8280
   Begin VB.Frame Frame3 
      Caption         =   "Data"
      Height          =   4452
      Left            =   2520
      TabIndex        =   25
      Top             =   120
      Width           =   3252
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3130
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   3010
         _ExtentX        =   5309
         _ExtentY        =   5521
         _Version        =   393216
         BackColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         ScrollBars      =   2
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.ComboBox cmbWave 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Select radiation"
         Top             =   360
         Width           =   1572
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "Refine zero "
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   28
         ToolTipText     =   "2theta real-2theta measured"
         Top             =   720
         Width           =   1210
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   27
         ToolTipText     =   "Toggle refinement flag (not reccomended)"
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Text            =   "1.54178"
         ToolTipText     =   "Put here the wavelength for CW or the beam angle, 2 theta, for the energy dispersive data"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Standard"
      Height          =   4452
      Left            =   5880
      TabIndex        =   20
      Top             =   120
      Width           =   2292
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3130
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   2050
         _ExtentX        =   3616
         _ExtentY        =   5521
         _Version        =   393216
         BackColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         ScrollBars      =   2
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.CheckBox ChkRef 
         Alignment       =   1  'Right Justify
         Caption         =   "Use standard"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtPolynom 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Text            =   "3"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbldetails 
         Alignment       =   1  'Right Justify
         Caption         =   "Polynomial"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1092
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cell"
      Height          =   3852
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2292
      Begin VB.ComboBox CmbCellType 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   288
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Select unit cell"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtRefine 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Text            =   "5.0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtRefine 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Text            =   "5.0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtRefine 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   16
         Text            =   "5.0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtRefine 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   15
         Text            =   "90."
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtRefine 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   14
         Text            =   "90."
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtRefine 
         BackColor       =   &H8000000E&
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   13
         Text            =   "90."
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "a"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   12
         ToolTipText     =   "Toggle refinement flag"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "b"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   11
         ToolTipText     =   "Toggle refinement flag"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "c"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   10
         ToolTipText     =   "Toggle refinement flag"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "alpha"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "Toggle refinement flag"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "beta"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "Toggle refinement flag"
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkRefine 
         Alignment       =   1  'Right Justify
         Caption         =   "gamma"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Toggle refinement flag"
         Top             =   3240
         Width           =   975
      End
   End
   Begin VB.TextBox txtDetails 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Text            =   "3"
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDetails 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Text            =   "10"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDetails 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Text            =   "3"
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbldetails 
      Alignment       =   1  'Right Justify
      Caption         =   "refine"
      Height          =   250
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   730
   End
   Begin VB.Label lbldetails 
      Alignment       =   1  'Right Justify
      Caption         =   "Ref. cycles"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   1212
   End
   Begin VB.Label lbldetails 
      Alignment       =   1  'Right Justify
      Caption         =   "width (%)"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuIgnore 
         Caption         =   "&Ignore: none"
      End
      Begin VB.Menu mnuImportdata 
         Caption         =   "Import data (Ascii)"
         Begin VB.Menu mnuImportascii 
            Caption         =   "4 values on line"
         End
         Begin VB.Menu mnuImport5 
            Caption         =   "5 values on line"
         End
      End
      Begin VB.Menu mnuImportStd 
         Caption         =   "Import calibration data"
      End
      Begin VB.Menu m5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportDataToGraphic 
         Caption         =   "Export data grid to Graphic"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export data grid to ascii file"
      End
      Begin VB.Menu mnuExportStdGridToAscii 
         Caption         =   "Export standard grid to ascii file"
      End
      Begin VB.Menu m1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Data"
      Begin VB.Menu mnuSortData 
         Caption         =   "Sort"
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetData 
         Caption         =   "Set data type to:"
         Begin VB.Menu mnuSetDataType 
            Caption         =   "2 theta"
            Index           =   0
         End
         Begin VB.Menu mnuSetDataType 
            Caption         =   "theta"
            Index           =   1
         End
         Begin VB.Menu mnuSetDataType 
            Caption         =   "d"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuSetDataType 
            Caption         =   "energy"
            Index           =   3
         End
      End
      Begin VB.Menu m20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeDataTo 
         Caption         =   "Change Data to:"
         Begin VB.Menu mnuChangeDataTo2theta 
            Caption         =   "2 theta "
         End
         Begin VB.Menu mnuChangeDataToTheta 
            Caption         =   "theta "
         End
         Begin VB.Menu mnuChangeDataToD 
            Caption         =   "d "
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuChangeDataToEnergy 
            Caption         =   "energy dispersive"
         End
      End
      Begin VB.Menu m17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertRow 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDeleteRow 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuDelColumn 
         Caption         =   "Delete column"
      End
      Begin VB.Menu m4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add to column"
         Index           =   0
         Begin VB.Menu mnuAddconst 
            Caption         =   "a constant value"
         End
         Begin VB.Menu mnuAddcol 
            Caption         =   "column "
         End
         Begin VB.Menu mnuAddreccol 
            Caption         =   "1/column "
         End
         Begin VB.Menu m1_ 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddweight1 
            Caption         =   "to weight: 1"
         End
      End
      Begin VB.Menu mnuMultiply 
         Caption         =   "Multiply column"
         Index           =   0
         Begin VB.Menu mnuMultiplyconst 
            Caption         =   "with a constant value"
         End
         Begin VB.Menu mnuMultiplycol 
            Caption         =   "with column"
         End
         Begin VB.Menu mnuMultiplyreccol 
            Caption         =   "with 1/column"
         End
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReadColumn 
         Caption         =   "Input data in column"
      End
      Begin VB.Menu m14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErase 
         Caption         =   "Clear data grid"
      End
      Begin VB.Menu mnuSetDataRows 
         Caption         =   "Set the number of rows"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasteDicvol 
         Caption         =   "Paste DICVOL results in data grid"
      End
   End
   Begin VB.Menu mnuEditStdGrid 
      Caption         =   "Standard"
      Begin VB.Menu mnuChangeStdFrom 
         Caption         =   "Change standard data FROM:"
         Begin VB.Menu mnuChangeStdDataFrom2theta 
            Caption         =   "2 theta,  TO"
            Begin VB.Menu mnuChangeStdData2thetaToTheta 
               Caption         =   "theta"
            End
            Begin VB.Menu mnuChangeStdData2thetaToD 
               Caption         =   "d"
            End
         End
         Begin VB.Menu mnuChangeStdFromtheta 
            Caption         =   "theta, TO"
            Begin VB.Menu mnuChangeStdDataFromthetaTo2theta 
               Caption         =   "2 theta"
            End
            Begin VB.Menu mnuChangeStdDataFromthetaTod 
               Caption         =   "d"
            End
         End
         Begin VB.Menu mnuChangeStdFromD 
            Caption         =   "d, TO"
            Begin VB.Menu mnuChangeStdDataFromdTo2theta 
               Caption         =   "2 theta"
            End
            Begin VB.Menu mnuChangeStdDataFromdTotheta 
               Caption         =   "theta"
            End
            Begin VB.Menu mnuChangeStdDataFromdToenergy 
               Caption         =   "energy dispersive"
            End
         End
         Begin VB.Menu mnuChangeStdDataFromEnergyTod 
            Caption         =   "energy dispersive  TO d"
         End
      End
      Begin VB.Menu m15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInputStdData 
         Caption         =   "Input std. data in column"
      End
      Begin VB.Menu m16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertRowStd 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuDeleteStdRow 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuDeleteStdColumn 
         Caption         =   "Delete Column"
      End
      Begin VB.Menu m13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearStd 
         Caption         =   "Clear std. grid"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Cell"
      Begin VB.Menu mnuAssignHKL 
         Caption         =   "Assign hkl's to peaks"
      End
      Begin VB.Menu m12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUnitcellparam 
         Caption         =   "Refinement of unit cell parameters"
         Begin VB.Menu mnuCELLLSTSQ 
            Caption         =   "least squares"
         End
         Begin VB.Menu m10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCellDispEnBatch 
            Caption         =   "batch run - dispersive energy data"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDavidonFletcher 
            Caption         =   "conjugated gradients"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLeastSquares 
            Caption         =   "overdetermined system, Newton"
         End
         Begin VB.Menu mnuSearchparam 
            Caption         =   "multidimensional search"
         End
         Begin VB.Menu m11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCellDispEnergy 
            Caption         =   "least squares, energy dispersive data"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuindex 
         Caption         =   "Automatic indexing (external)"
         Begin VB.Menu mnuIndexDicvol 
            Caption         =   "Use DICVOL "
         End
         Begin VB.Menu mnuUseIto 
            Caption         =   "Use ITO "
         End
         Begin VB.Menu mnuUseTreor 
            Caption         =   "Use TREOR90"
         End
      End
      Begin VB.Menu m7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComputeHKL 
         Caption         =   "Compute HKL, d / theta"
         Begin VB.Menu mnuPutinGrid 
            Caption         =   "send to data grid"
         End
         Begin VB.Menu mnuGenerate 
            Caption         =   "send to report pad "
         End
         Begin VB.Menu m19 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSendHKLToStd 
            Caption         =   "send to standard grid (theor. values)"
         End
      End
      Begin VB.Menu mnuSpaceGroup 
         Caption         =   "Compute HKL - space group check (HKLGEN)"
      End
   End
   Begin VB.Menu mnuOverlap 
      Caption         =   "Decimate"
      Begin VB.Menu mnuOverlapFullprof 
         Caption         =   "FULLPROF .fou file"
      End
      Begin VB.Menu mnuOverlapGsas 
         Caption         =   "GSAS .rfl file"
      End
      Begin VB.Menu mnuOverlapShelx 
         Caption         =   "SHELX .hkl file"
      End
   End
End
Attribute VB_Name = "frmRefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Function par(X As Integer) As Boolean
par = False
If X - 2 * (X \ 2) = 0 Then par = True
Exit Function
End Function


Sub putdata(dest As Integer)
'destinatie este  1,2,3 adica raport, grid sau grid 1
Dim tlimit As Single, i As Integer, results(8) As Double, cell(7) As Double, coderoare As Boolean, cellr(7) As Double, returncode As Boolean, lambda As Double, sT As String, bravais As String * 5
Dim hlimit As Integer, klimit As Integer, llimit As Integer
Dim d() As Integer, jdate As Integer, ih As Integer, ik As Integer, il As Integer, j As Integer
Dim dd() As crystallo, jdatebune As Integer, indicatorschimba As Boolean, finalsort() As Integer
Dim ar As Double, br As Double, cr As Double, alr As Double, ber As Double, gar As Double
'dd.h(1) = 1
On Error GoTo errortrap
'verific parametrii
For i = 0 To 5
If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect cell parameters. Try again."
Next i
lambda = Val(txt)
If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."
tlimit = InputBox("Input the maximum 2 theta : ", prog_name, 60)
If tlimit < 2 Or tlimit > 160 Then Err.Raise 1101, , "Choose another value for 2 theta limit."
raport "The maximum value for h,k,l is related to the available memory... "
hlimit = InputBox("Input the maximum h value : ", prog_name, 10)
klimit = InputBox("Input the maximum k value : ", prog_name, 10)
llimit = InputBox("Input the maximum l value :", prog_name, 10)
If (llimit < 1 Or klimit < 1 Or hlimit < 1) Then Err.Raise 1101, , "Choose another values for max H, K, L (higher than 1)."
Screen.MousePointer = 11
'aflu numarul de valori nrval - din grid si redimensionez
'citesc cell si apoi calculez cellr
    For i = 1 To 6
    cell(i) = Val(txtRefine(i - 1).Text)
    Next i
'trimit in grade
cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))
Call reciproc(cell, cellr, coderoare)
If (coderoare) Then Err.Raise 1101, , "Error in computing the reciproc cell."
For i = 1 To 6
results(i) = cellr(i)

Next i

ar = results(1)

br = results(2)

cr = results(3)

alr = results(4)

ber = results(5)

gar = results(6)



results(7) = lambda

results(8) = 0

raport strLinie

raport "Direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i

raport strLinie

raport "Computing d, theta. This may take a while, please wait."

DoEvents

raport "Wavelength used : " & CStr(lambda)

raport "2 theta limit : " & CStr(tlimit)

raport "HKL limits :" & CStr(hlimit) & "  " & CStr(klimit) & "  " & CStr(llimit)



Dim tempdata As Double, lasthkl As Integer, lasthk As Integer, lastll As Integer, lasthh As Integer, lastkk As Integer

'generez toate reflectiile permise pentru celula...cubic, etc...

'le ordonez, apoi le separ...

'ulterior le separ in functie de simetrie

''aici e doar generarea tuturor posibile

Select Case CmbCellType.ListIndex

Case 0 'cubic

ReDim dd((hlimit + 1) * (1 + klimit) * (llimit + 1))

ReDim d((hlimit + 1) * (1 + klimit) * (llimit + 1))

j = 0

For il = 0 To llimit

For ik = 0 To klimit

For ih = 0 To hlimit

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * (ih * ih + ik * ik + il * il) * ar * ar)

If Abs(tempdata) < 1 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next ih

Next ik

Next il

jdate = j - 1





Case 1 'tetra

ReDim dd((hlimit + 1) * (1 + klimit) * (llimit + 1))

ReDim d((hlimit + 1) * (1 + klimit) * (llimit + 1))

j = 0

For il = 0 To llimit

For ik = 0 To klimit

For ih = 0 To hlimit

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * ((ih * ih + ik * ik) * ar * ar + il * il * cr * cr))

If Abs(tempdata) < 1 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next ih

Next ik

Next il

jdate = j - 1





Case 2 'ortho

ReDim dd((hlimit + 1) * (1 + klimit) * (llimit + 1))

ReDim d((hlimit + 1) * (1 + klimit) * (llimit + 1))





j = 0

For il = 0 To llimit

For ik = 0 To klimit

For ih = 0 To hlimit

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * ((ih * ih * ar * ar + ik * ik * br * br + il * il * cr * cr)))

If Abs(tempdata) < 1 And Abs(tempdata) > 0.01 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next ih

Next ik

Next il

jdate = j - 1













Case 3 'rhombo

ReDim dd((2 * hlimit + 1) * (1 + 2 * klimit) * (2 * llimit + 1))

ReDim d((2 * hlimit + 1) * (1 + 2 * klimit) * (2 * llimit + 1))

j = 0

For ih = llimit To -hlimit Step -1

For ik = klimit To -klimit Step -1

For il = hlimit To -llimit Step -1

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * ((ih * ih + ik * ik + il * il + 2 * (ik * il + il * ih + ik * ih) * Cos(alr / rd)) * ar * ar))

If Abs(tempdata) < 1 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next il

Next ik

Next ih

jdate = j - 1







Case 4 'hex

ReDim dd((2 * hlimit + 1) * (1 + 2 * klimit) * (llimit + 1))

ReDim d((2 * hlimit + 1) * (1 + 2 * klimit) * (llimit + 1))

j = 0

For il = 0 To llimit

For ik = klimit To -klimit Step -1

For ih = hlimit To -hlimit Step -1

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * ((ih * ih + ik * ik + ih * ik) * ar * ar + il * il * cr * cr))

If Abs(tempdata) < 1 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next ih

Next ik

Next il

jdate = j - 1















Case 5 'mono

ReDim dd((2 * hlimit + 1) * (1 + klimit) * (2 * llimit + 1))

ReDim d((2 * hlimit + 1) * (1 + klimit) * (2 * llimit + 1))



j = 0

For il = llimit To -llimit Step -1

For ik = 0 To klimit

For ih = hlimit To -hlimit Step -1

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * ((ih * ih * ar * ar + ik * ik * br * br + il * il * cr * cr + 2 * il * ih * cr * ar * Cos(ber / rd))))



If tempdata < 1 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next ih

Next ik

Next il

jdate = j - 1





Case 6 'tric

ReDim dd((2 * hlimit + 1) * (1 + 2 * klimit) * (2 * llimit + 1))

ReDim d((2 * hlimit + 1) * (1 + 2 * klimit) * (2 * llimit + 1))



j = 0

For il = llimit To -llimit Step -1

For ik = klimit To -klimit Step -1

For ih = hlimit To -hlimit Step -1

j = j + 1

tempdata = Sqr(lambda * lambda / 4 * ((ih * ih * ar * ar + ik * ik * br * br + il * il * cr * cr + 2 * il * ih * cr * ar * Cos(ber / rd) + 2 * il * ik * cr * br * Cos(alr / rd) + 2 * ih * ik * br * ar * Cos(gar / rd))))



If tempdata < 1 Then

dd(j).doitheta = CSng(2 * rd * asin(tempdata))

Else

dd(j).doitheta = 0

End If

If dd(j).doitheta <= 0 Or dd(j).doitheta > 160 Then

dd(j).d = 0

Else

dd(j).d = CSng(lambda / 2 / (Sin(dd(j).doitheta / 2 / rd)))

End If

dd(j).h = ih

dd(j).k = ik

dd(j).l = il

Next ih

Next ik

Next il

jdate = j - 1

End Select

''le arunc pe cele care sunt aiurea

''numar cate sunt, maxime pentru p1



jdatebune = 0

For j = 1 To jdate

''Call allow_line(dd(j), tlimit, bravais, allowed)

If dd(j).d > 0.001 And dd(j).doitheta < tlimit Then

'le pastrez

jdatebune = jdatebune + 1

d(jdatebune) = j

End If

Next j

''le ordonez dupa cresterea lui theta

''astfel pot sa separ dupa reflectii echivalente ulterior

' tin minte nr de ordine j, pointer......



Dim swap As Variant

indicatorschimba = True

''ReDim Preserve d(jdatebune) 'pastrez doar ce este bun

'sort

Do Until indicatorschimba = False

indicatorschimba = False

For i = 1 To jdatebune - 1

If dd(d(i)).d < dd(d(i + 1)).d Then

swap = d(i)

d(i) = d(i + 1)

d(i + 1) = swap

     indicatorschimba = True

     End If

Next i

Loop



'aici sunt in ordine



'aici arunc pe toate cele care au acelasi d; si in care hh, kk si ll sunt identice







For i = 1 To jdatebune - 1

If dd(d(i)).d = dd(d(i + 1)).d Then

'daca au acelasi d, pe ultimul il arunc daca au acelasi hh, kk si ll

If (dd(d(i)).h) ^ 2 = (dd(d(i + 1)).h) ^ 2 And (dd(d(i)).k) ^ 2 = (dd(d(i + 1)).k) ^ 2 And (dd(d(i)).l) ^ 2 = (dd(d(i + 1)).l) ^ 2 Then dd(d(i + 1)).d = 0

End If

Next i





Select Case CmbCellType.ListIndex

Case 0 'cubic

'le dau afara pe alea care sunt echivalente

lasthkl = 0

For i = 1 To jdatebune

If (dd(d(i)).h) ^ 2 + (dd(d(i)).k) ^ 2 + (dd(d(i)).l) ^ 2 = lasthkl Then

'il fac pe d=0, ulterior mai arunc un set de date...

dd(d(i)).d = 0

Else

lasthkl = (dd(d(i)).h) ^ 2 + (dd(d(i)).k) ^ 2 + (dd(d(i)).l) ^ 2

End If

Next i



sT = InputBox("Input the Bravais symbol (standard setting only: P, I, F) :", prog_name, "P")

sT = UCase$(sT)

raport LCase(CStr(CmbCellType.List(CmbCellType.ListIndex))) & " cell, " & sT

If Not ((sT = "P") Or (sT = "I") Or (sT = "F")) Then

raport "Allowed symbols: Triclinic P; Monoclinic P, C; Orthorombic P, C, I, F; Tetragonal P, I; Hexagonal P; Trigonal R; Cubic P, I, F." & vbCrLf

Err.Raise 1102

End If



Case 1 'tetragonal

lasthkl = -1

lastll = -1 ''nu trebuie sa l arunc pe 00 1 etc...

For i = 1 To jdatebune

If (((dd(d(i)).h) ^ 2 + (dd(d(i)).k) ^ 2 = lasthkl) And ((dd(d(i)).l) ^ 2 = lastll)) Then

'il fac pe d=0, ulterior mai arunc un set de date...

dd(d(i)).d = 0

Else

lasthkl = (dd(d(i)).h) ^ 2 + (dd(d(i)).k) ^ 2

lastll = (dd(d(i)).l) ^ 2

End If

Next i



sT = InputBox("Input the Bravais symbol (standard setting only: P, I) :", prog_name, "P")

sT = UCase$(sT)

raport LCase(CStr(CmbCellType.List(CmbCellType.ListIndex))) & " cell, " & sT



If Not ((sT = "P") Or (sT = "I")) Then

raport "Allowed symbols: Triclinic P; Monoclinic P, C; Orthorombic P, C, I, F; Tetragonal P, I; Hexagonal P; Trigonal R; Cubic P, I, F." & vbCrLf

Err.Raise 1102

End If





Case 2 'ortho



sT = InputBox("Input the Bravais symbol (standard setting only: P, C, I, F) :", prog_name, "P")

sT = UCase$(sT)

raport LCase(CStr(CmbCellType.List(CmbCellType.ListIndex))) & " cell, " & sT



If Not ((sT = "P") Or (sT = "C") Or (sT = "I") Or (sT = "F")) Then

raport "Allowed symbols: Triclinic P; Monoclinic P, C; Orthorombic P, C, I, F; Tetragonal P, I; Hexagonal P; Trigonal R; Cubic P, I, F." & vbCrLf

Err.Raise 1102

End If





Case 3 'rombo

raport "assuming R cell."

lasthkl = 0

For i = 1 To jdatebune

If (dd(d(i)).h) ^ 2 + (dd(d(i)).k) ^ 2 + (dd(d(i)).l) ^ 2 + dd(d(i)).l * dd(d(i)).h + dd(d(i)).l * dd(d(i)).k + dd(d(i)).k * dd(d(i)).h = lasthkl Then

'il fac pe d=0, ulterior mai arunc un set de date...

dd(d(i)).d = 0

Else

lasthkl = (dd(d(i)).h) ^ 2 + (dd(d(i)).k) ^ 2 + (dd(d(i)).l) ^ 2 + dd(d(i)).l * dd(d(i)).h + dd(d(i)).l * dd(d(i)).k + dd(d(i)).k * dd(d(i)).h

End If

Next i











Case 4  'hexa

raport "assuming P cell"



Case 5 'mono



sT = InputBox("Input the Bravais symbol (standard setting only : P, C) :", prog_name, "P")

sT = UCase$(sT)

raport LCase(CStr(CmbCellType.List(CmbCellType.ListIndex))) & " cell, " & sT

If Not ((sT = "P") Or (sT = "C")) Then

raport "Allowed symbols: Triclinic P; Monoclinic P, C; Orthorombic P, C, I, F; Tetragonal P, I; Hexagonal P; Trigonal R; Cubic P, I, F." & vbCrLf

Err.Raise 1102

End If



Case 6

raport "assuming P cell"

End Select



Select Case sT



Case "P"

'nimic de facut

Case "R"

'tot nimic de facut

Case "F"

'toti trebuie sa fie fie pari fie impari

For i = 1 To jdatebune

If par(dd(d(i)).h) Then

'h este par

    If Not (par(dd(d(i)).k) And par(dd(d(i)).l)) Then dd(d(i)).d = 0

Else

'h este impar

    If Not ((Not par(dd(d(i)).k) And Not (par(dd(d(i)).l)))) Then dd(d(i)).d = 0

End If

Next i



Case "I"

'suma lor trebuie sa fie para

For i = 1 To jdatebune

If Not (par(dd(d(i)).h + dd(d(i)).k + dd(d(i)).l)) Then dd(d(i)).d = 0

Next i



Case "C"

'h +k trebuie sa fie pare...

For i = 1 To jdatebune

If Not (par(dd(d(i)).h + dd(d(i)).k)) Then dd(d(i)).d = 0

Next i



End Select



'o ultima sortare, mai arunc pe alea care sunt echivalente cum ar fi 010, 100 la cubic......

''zerourile out, sortare finala



jdatebune = 0

For j = 1 To jdate

''Call allow_line(dd(j), tlimit, bravais, allowed)

If dd(j).d > lambda / 2 And dd(j).doitheta < tlimit Then

'le pastrez

jdatebune = jdatebune + 1

d(jdatebune) = j

End If

Next j





raport CStr(jdatebune) & " planes..."

'sortare inca odata

indicatorschimba = True

ReDim Preserve d(jdatebune) 'pastrez doar ce este bun

'sort

Do Until indicatorschimba = False

indicatorschimba = False

For i = 1 To jdatebune - 1

If dd(d(i)).d < dd(d(i + 1)).d Then

swap = d(i)

d(i) = d(i + 1)

d(i + 1) = swap

     indicatorschimba = True

     End If

Next i

Loop



Select Case dest

Case 1 ''Convert3Main.txtraport

raport strLinie

raport "           h    k    l       d /A   2 theta /deg"

raport strLinie

For i = 1 To jdatebune

raport Format$(i, "@@@@      ") & Format$(Format$(CStr(dd(d(i)).h), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).k), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).l), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).d), " ##0.00000 "), "@@@@@@@@@@@") & Format$(Format$(CStr(dd(d(i)).doitheta), " ##0.00000 "), "@@@@@@@@@@@")

''raport CStr(dd(d(i)).d) & " H= " & CStr(dd(d(i)).h) & " K= " & CStr(dd(d(i)).k) & " L= " & CStr(dd(d(i)).l)

Next i

Case 2 ''frmRefine.grid

raport "data sent to grid, as h,k,l and d; weigth is 1"

'clear data grid

mnuErase_Click

DoEvents

grid.Rows = jdatebune + 3

For i = 1 To jdatebune

grid.Row = i

grid.Col = 1: grid.Text = CStr(dd(d(i)).h)

grid.Col = 2: grid.Text = CStr(dd(d(i)).k)

grid.Col = 3: grid.Text = CStr(dd(d(i)).l)

grid.Col = 4: grid.Text = CStr(dd(d(i)).d)

grid.Col = 5: grid.Text = "1"

''raport Format$(i, "@@@@      ") & Format$(Format$(CStr(dd(d(i)).h), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).k), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).l), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).d), " ##0.00000 "), "@@@@@@@@@@@") & Format$(Format$(CStr(dd(d(i)).doitheta), " ##0.00000 "), "@@@@@@@@@@@")

''raport CStr(dd(d(i)).d) & " H= " & CStr(dd(d(i)).h) & " K= " & CStr(dd(d(i)).k) & " L= " & CStr(dd(d(i)).l)

Next i



Case 3 ''frmRefine.grid1



raport strLinie

raport "Data sent to standard grid, theroretical values, as d values"

raport strLinie

raport "           h    k    l       d /A   2 theta /deg"

raport strLinie

Grid1.Rows = jdatebune + 2

DoEvents

Grid1.Row = 1

Grid1.Col = 0

For i = 1 To jdatebune

raport Format$(i, "@@@@      ") & Format$(Format$(CStr(dd(d(i)).h), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).k), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).l), "#0   "), "@@@@@") & Format$(Format$(CStr(dd(d(i)).d), " ##0.00000 "), "@@@@@@@@@@@") & Format$(Format$(CStr(dd(d(i)).doitheta), " ##0.00000 "), "@@@@@@@@@@@")

Grid1.Text = Format$(dd(d(i)).d, "##0.00###")

Grid1.Row = Grid1.Row + 1

Next i







End Select











Screen.MousePointer = 0



Exit Sub

errortrap:

Screen.MousePointer = 0

If Not (Err.Number = 1102) Then raport Err.Description

Err.Clear

Exit Sub

















End Sub



















Sub numargrid(jcount As Integer)
On Error GoTo handleit
Dim i As Integer, strA As String, strB As String, strC As String, strd As String
grid.Row = 1
grid.Col = 1
jcount = 1
Do Until grid.Row = grid.Rows - 1
strA = grid.Text
grid.Col = 2: strB = grid.Text
grid.Col = 3: strC = grid.Text
grid.Col = 4: strd = grid.Text
If Val(strA) + Val(strB) + Val(strC) + Val(strd) = 0 Then Exit Do
jcount = jcount + 1
grid.Row = grid.Row + 1
grid.Col = 1
Loop
Exit Sub
handleit:
raport "Error in numarGrid routine"
Exit Sub
End Sub





Sub usestandard(polcoeff() As Double)

'fac aici verificarile pentru corectii;9 coeff

On Error GoTo handleit

Dim i As Integer, j As Integer, np As Integer, pgrad As Integer 'np este nr de puncte, pgrad este gradul polyn

Dim txt1 As String, ex() As Double, ob() As Double, dif() As Double

''ex sunt expected values, ob sunt observate, dif este diferenta

Dim z() As Double, ii() As Double, solutie() As Double, coderoare As Boolean



txtPolynom = CStr(Abs(Val(txtPolynom.Text)))

pgrad = CInt((txtPolynom))

If pgrad < 0 Then Err.Raise 1101, , "Try again, nothing to do......"

If pgrad > 8 Then Err.Raise 1101, , "The polynomial degree is too high, nothing to do..."

'numar cate date am

Grid1.Row = 1

Do

If Grid1.Row = Grid1.Rows - 3 Then Grid1.Rows = Grid1.Rows + 1

Grid1.Col = 0

txt1 = Grid1.Text

Grid1.Col = 1

If Not (Val(Grid1.Text) <> 0 And Val(txt1) <> 0) Then Exit Do

Grid1.Row = Grid1.Row + 1

Loop

np = CInt(Grid1.Row) - 1

If np < pgrad + 1 Then Err.Raise 1101, "Not enough data points for this polynom degree..."

raport "I don't check for the consistency of your data. Beware..."

raport "If the data type for standard is different from the experimental data you'll get a mess."
ReDim ex(np), ob(np), dif(np)
For i = 1 To np
Grid1.Row = i
Grid1.Col = 0: ex(i) = Val(Grid1.Text)
Grid1.Col = 1: ob(i) = Val(Grid1.Text)
dif(i) = ob(i) - ex(i)
Next i
'ca x pentru polinom calculez pe ob; y va fi diferenta
'trebuie sa intorc coeficientii unui polinom care imi da diferenta dintre ob si ex pentru un anumit ob
ReDim z(np, pgrad + 1) As Double, ii(np) As Double, solutie(pgrad + 1) As Double
For j = 1 To np
For i = 2 To pgrad + 1
z(j, i) = ob(j) ^ (i - 1)
Next i
Next j

For i = 1 To np
ii(i) = dif(i)
z(i, 1) = 1
Next i

Call pseudoinv(np, pgrad + 1, z, ii, solutie, 0.0000001, coderoare)
If coderoare Then Err.Raise 1101, , "Unexpected error in PseudoInverse routine."
For i = 1 To 9
polcoeff(i) = 0
Next i
For i = 1 To pgrad + 1
polcoeff(i) = solutie(i)
Next i

Exit Sub
handleit:
For i = 1 To 9
polcoeff(i) = 0
Next i
raport "Routine: usesStandard"
raport "An error has occured...check the polynomial degree or the standard data...."
Err.Clear
Exit Sub

End Sub















Private Sub chkRefine_Click(Index As Integer)

Select Case Index

Case 0

Case 1

Case 2

Case 3

Case 4

Case 5

Case 7

'zero error

Case 6

'lambda

cmbWave.ListIndex = 24

DoEvents



End Select



End Sub



Private Sub CmbCellType_Click()

CmbCellType_GotFocus

End Sub



Private Sub CmbCellType_GotFocus()

On Error GoTo handleit

Dim i As Integer

For i = 0 To 5

txtRefine(i).Enabled = False: chkRefine(i).Enabled = False: chkRefine(i).Value = False

Next i

Select Case CmbCellType.ListIndex

Case 0 'cubic

txtRefine(0).Enabled = True: chkRefine(0).Enabled = True

txtRefine(1).Text = txtRefine(0).Text: txtRefine(2).Text = txtRefine(0).Text

txtRefine(3).Text = "90": txtRefine(4).Text = "90": txtRefine(5).Text = "90":



Case 1 'tetragonal

txtRefine(0).Enabled = True: chkRefine(0).Enabled = True

txtRefine(2).Enabled = True: chkRefine(2).Enabled = True

txtRefine(1).Text = txtRefine(0).Text

txtRefine(3).Text = "90": txtRefine(4).Text = "90": txtRefine(5).Text = "90"



Case 2 'ortho

txtRefine(0).Enabled = True: chkRefine(0).Enabled = True

txtRefine(1).Enabled = True: chkRefine(1).Enabled = True

txtRefine(2).Enabled = True: chkRefine(2).Enabled = True

txtRefine(3).Text = "90": txtRefine(4).Text = "90": txtRefine(5).Text = "90"



Case 3 'rombo

txtRefine(0).Enabled = True: chkRefine(0).Enabled = True

txtRefine(3).Enabled = True: chkRefine(3).Enabled = True

txtRefine(1).Text = txtRefine(0).Text: txtRefine(2).Text = txtRefine(0).Text

txtRefine(4).Text = txtRefine(3).Text: txtRefine(5).Text = txtRefine(3).Text



Case 4 'hexa

txtRefine(0).Enabled = True: chkRefine(0).Enabled = True

txtRefine(2).Enabled = True: chkRefine(2).Enabled = True

txtRefine(1).Text = txtRefine(0).Text

txtRefine(3).Text = "90": txtRefine(5).Text = "120": txtRefine(4).Text = "90"





Case 5 'mono

txtRefine(0).Enabled = True: chkRefine(0).Enabled = True

txtRefine(1).Enabled = True: chkRefine(1).Enabled = True

txtRefine(2).Enabled = True: chkRefine(2).Enabled = True

txtRefine(4).Enabled = True: chkRefine(4).Enabled = True

txtRefine(3).Text = "90": txtRefine(5).Text = "90"



Case 6 'triclinic

'no constraints applied...

For i = 0 To 5 'le fac pe toate enabled la intrare

txtRefine(i).Enabled = True: chkRefine(i).Enabled = True

Next i

End Select

Exit Sub

handleit:

Err.Clear

raport "Error in CmbCellType ..."

Exit Sub

End Sub



Private Sub CmbCellType_KeyDown(KeyCode As Integer, Shift As Integer)

CmbCellType_GotFocus

End Sub



Private Sub cmbWave_Click()

cmbWave_GotFocus

DoEvents

End Sub



Private Sub cmbWave_GotFocus()

On Error GoTo errortrap

mnuCellDispEnergy.Enabled = False

DoEvents

Select Case cmbWave.ListIndex

Case 0

txt.Text = 1.54178

Case 1

txt.Text = 1.540562

Case 2

txt.Text = 1.54439

Case 3

txt.Text = 1.3921

Case 4

txt.Text = 2.29089

Case 5

txt.Text = 2.2897

Case 6

txt.Text = 2.2936

Case 7

txt.Text = 2.0848

Case 8

txt.Text = 1.93734

Case 9

txt.Text = 1.93604

Case 10

txt.Text = 1.93998

Case 11

txt.Text = 1.75653

Case 12

txt.Text = 1.7915

Case 13

txt.Text = 1.78896

Case 14

txt.Text = 1.79285

Case 15

txt.Text = 1.62075

Case 16

txt.Text = 0.71073

Case 17

txt.Text = 0.7093

Case 18

txt.Text = 0.71359

Case 19

txt.Text = 0.63225

Case 20

txt.Text = 0.5608

Case 21

txt.Text = 0.5594

Case 22

txt.Text = 0.56379

Case 23

txt.Text = 0.49701

Case 24

'other

mnuCellDispEnergy.Enabled = True

End Select

Exit Sub

DoEvents

errortrap:

Exit Sub

End Sub













Private Sub Form_Load()
Dim largime As Single, i As Integer
DoEvents
largime = grid.Width / 8
grid.Cols = 6
grid.ColWidth(0) = 0.75 * largime
grid.ColWidth(1) = largime
grid.ColWidth(2) = largime
grid.ColWidth(3) = largime
grid.ColWidth(4) = 2 * largime
grid.ColWidth(5) = 1.6 * largime

'numar liniile
Refresh
grid.Col = 0
For i = 1 To grid.Rows - 1
''Grid.ColAlignment = 3
grid.Row = i
grid.Text = CStr(i)
''''''''Grid.CellAlignment = 3
Next i
'pun titlu
grid.Row = 0
grid.Col = 1
''''''''Grid.CellAlignment = 3
grid.Text = "H"
grid.Col = 2
''''''''Grid.CellAlignment = 3
grid.Text = "K"
grid.Col = 3
''''''''Grid.CellAlignment = 3
grid.Text = "L"
grid.Col = 4
''''''''Grid.CellAlignment = 3
grid.Text = "d /A"
grid.Col = 5
''''''''Grid.CellAlignment = 3
grid.Text = "weigth"
'pun focus in alta parte
Grid1.ColWidth(0) = Grid1.Width / 2.2
Grid1.ColWidth(1) = Grid1.Width / 2.2
Grid1.Col = 0
Grid1.Row = 0
Grid1.Text = "Theor. "
Grid1.Col = 1
Grid1.Text = "Obs."
grid.Col = 1
grid.Row = 1
frmRefine.grid_Click




CmbCellType.AddItem "Cubic"

CmbCellType.AddItem "Tetragonal"

CmbCellType.AddItem "Orthorhombic"

CmbCellType.AddItem "Rhombohedral"
CmbCellType.AddItem "Hexagonal"
CmbCellType.AddItem "Monoclinic"
CmbCellType.AddItem "Triclinic"
CmbCellType.ListIndex = 0

cmbWave.AddItem "Cu -K alpha"
cmbWave.AddItem "Cu -K alpha 1"
cmbWave.AddItem "Cu -K alpha 2"
cmbWave.AddItem "Cu -K beta"
cmbWave.AddItem "Cr -K alpha"
cmbWave.AddItem "Cr -K alpha 1"
cmbWave.AddItem "Cr -K alpha 2"
cmbWave.AddItem "Cr -K beta"
cmbWave.AddItem "Fe -K alpha"
cmbWave.AddItem "Fe -K alpha 1"
cmbWave.AddItem "Fe -K alpha 2"
cmbWave.AddItem "Fe -K beta"
cmbWave.AddItem "Co -K alpha"
cmbWave.AddItem "Co -K alpha 1"
cmbWave.AddItem "Co -K alpha 2"
cmbWave.AddItem "Co -K beta"
cmbWave.AddItem "Mo -K alpha"
cmbWave.AddItem "Mo -K alpha 1"
cmbWave.AddItem "Mo -K alpha 2"
cmbWave.AddItem "Mo -K beta"
cmbWave.AddItem "Ag -K alpha"
cmbWave.AddItem "Ag -K alpha 1"
cmbWave.AddItem "Ag -K alpha 2"
cmbWave.AddItem "Ag -K beta"
''cmbWave.AddItem "Energy /keV"

cmbWave.AddItem "other/2theta detector..."
cmbWave.ListIndex = 1





End Sub



Private Sub Form_Unload(Cancel As Integer)
If mnuQuit.Enabled = True Then
If Not (CloseWindow("Are you sure you want to close this window ?", prog_name & " - UnitCell")) Then Cancel = -1
Else
Cancel = -1
Exit Sub
End If
End Sub

Sub grid_Click()
On Error GoTo handleit
If grid.Row > grid.Rows - 2 Then
grid.Rows = grid.Rows + 1
grid.Col = 0
grid.Row = grid.Rows - 1
grid.Text = CStr(grid.Row)
''''''Grid.CellAlignment = 3
End If
Exit Sub
handleit:
raport "Error in grid_Click routine..."
Err.Clear
Exit Sub
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
If grid.Row = 0 Or grid.Col = 0 Then Exit Sub
Select Case KeyCode
Case 8 'backspace
If Len(grid.Text) > 0 Then grid.Text = left$(grid.Text, Len(grid.Text) - 1)
Case 13 'enter
If grid.Col = 5 Then
grid.Col = 1
If grid.Row > grid.Rows - 2 Then frmRefine.grid_Click
grid.Row = grid.Row + 1
Else
grid.Col = grid.Col + 1
frmRefine.grid_Click
End If

Case 46 'delete
If Len(grid.Text) > 0 Then grid.Text = right$(grid.Text, Len(grid.Text) - 1)
Case 110, 188, 190 ', .
grid.Text = grid.Text + "."
Case 189, 109 '-
grid.Text = "-" & grid.Text
Case 48, 96 '0
grid.Text = grid.Text + "0"
Case 49, 97 '1
grid.Text = grid.Text + "1"
Case 50, 98 '2
grid.Text = grid.Text + "2"
Case 51, 99 '3
grid.Text = grid.Text + "3"
Case 52, 100 '4
grid.Text = grid.Text + "4"
Case 53, 101 '5
grid.Text = grid.Text + "5"
Case 54, 102 '6
grid.Text = grid.Text + "6"
Case 55, 103 '7
grid.Text = grid.Text + "7"
Case 56, 104 '8
grid.Text = grid.Text + "8"
Case 57, 105 '9
grid.Text = grid.Text + "9"
End Select
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuEdit
End Sub


Private Sub grid1_Click()
If Grid1.Row > Grid1.Rows - 2 Then Grid1.Rows = Grid1.Rows + 1
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Grid1.Row = 0 Then Exit Sub
Select Case KeyCode
Case 8 'backspace
If Len(Grid1.Text) > 0 Then Grid1.Text = left$(Grid1.Text, Len(Grid1.Text) - 1)
Case 13 'enter
If grid.Col = 1 Then
grid.Col = 0
If grid.Row > grid.Rows - 2 Then frmRefine.grid_Click
Grid1.Row = Grid1.Row + 1

Else

Grid1.Col = Grid1.Col + 1
''frmRefine.grid1_Click
End If
Case 46 'delete
If Len(Grid1.Text) > 0 Then Grid1.Text = right$(Grid1.Text, Len(Grid1.Text) - 1)
Case 110, 188, 190 ', .
Grid1.Text = Grid1.Text + "."
Case 109, 189 '-
''grid1.Text = "-" & grid.Text
Case 48, 96 '0
Grid1.Text = Grid1.Text + "0"
Case 49, 97 '1
Grid1.Text = Grid1.Text + "1"
Case 50, 98 '2
Grid1.Text = Grid1.Text + "2"
Case 51, 99 '3
Grid1.Text = Grid1.Text + "3"
Case 52, 100 '4
Grid1.Text = Grid1.Text + "4"
Case 53, 101 '5
Grid1.Text = Grid1.Text + "5"
Case 54, 102 '6
Grid1.Text = Grid1.Text + "6"
Case 55, 103 '7
Grid1.Text = Grid1.Text + "7"
Case 56, 104 '8
Grid1.Text = Grid1.Text + "8"
Case 57, 105 '9
Grid1.Text = Grid1.Text + "9"
End Select







End Sub









Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then PopupMenu mnuEditStdGrid



End Sub



Private Sub mnuAddcol_Click()

Dim t As Integer, i As Integer, j As Integer, txt As String
On Error GoTo errortrap

t = 4

t = InputBox("Insert the data column you want to modify (H is in 1) :", prog_name, 4)

j = InputBox("Insert the number of column you want to add (attention, data may be unusable after this operation) :", prog_name, 1)

If t < 6 And t > 0 And j < 6 And j > 0 Then

For i = 1 To grid.Rows - 1

grid.Col = j



grid.Row = i

txt = grid.Text

If grid.Text = "" Then Exit For

grid.Col = t

grid.Text = Format$(Val(grid.Text) + txt, "####0.0000#")

Next i

End If

Exit Sub

errortrap:

Exit Sub



End Sub



Private Sub mnuAddconst_Click()

AddToColumn grid

End Sub



Sub AddToColumn(grid As Object)

Dim t As Integer, i As Integer, j As Single

On Error GoTo errortrap

t = 4

t = InputBox("Insert the data column you want to modify (H is in 1) :", prog_name, 4)

If t < 6 And t > 0 Then

grid.Col = t

j = InputBox("Insert the value to add (attention, data may be unusable after this operation) :", prog_name, 1)

For i = 1 To grid.Rows - 1

grid.Row = i

grid.Text = Format$(Val(grid.Text) + j, "####0.00###")

Next i

End If

Exit Sub

errortrap:

Exit Sub



End Sub

Private Sub mnuAddreccol_Click()

Dim t As Integer, i As Integer, j As Integer, txt As String

On Error GoTo errortrap

t = 4

t = InputBox("Insert the data column you want to modify (H is in 1) :", prog_name, 4)

j = InputBox("Insert the number of column you want to add as 1/column (attention, data may be unusable after this operation) :", prog_name, 1)

If t < 6 And t > 0 And j < 6 And j > 0 Then

For i = 1 To grid.Rows - 1

grid.Col = j



grid.Row = i

txt = CStr(1 / Val(grid.Text))

If grid.Text = "" Then Exit For

grid.Col = t

grid.Text = Format$(Val(grid.Text) + txt, "####0.0000#")

Next i

End If

Exit Sub

errortrap:

raport "Error in mnuAddr _click routine."

Exit Sub

End Sub



Private Sub mnuAddweight1_Click()

Dim valori As Integer, i As Integer, j As Single, t1 As Integer, t2 As Integer

On Error GoTo errortrap

'numar cate coloane am, numar si reduc numarul de valori din rows

t1 = InputBox("Input the starting row", prog_name, 1)

t2 = InputBox("Input the ending row", prog_name, 99)

If t2 < t1 Or t1 < 1 Then Err.Raise 1101

grid.Col = 5

For i = t1 To t2

grid.Row = i

grid.Text = CStr(Val(grid.Text) + 1)

Next i

Exit Sub

errortrap:

Exit Sub

End Sub



Private Sub mnuAssignHKL_Click()

  
Dim tt As Double, t As Integer, v As String, limit As Single, out1 As Integer, out2 As Integer, in1 As Integer
Static message As Boolean
Dim nhkls As Integer
Dim nrlines As Integer, TwoThetaLimit As Single
Dim AssignHKL() As crystallo
raport "No consistency check made (between the unit cell parameters and the space group...)"
On Error GoTo errtrap
nrlines = number_of_lines_for_indexing
If nrlines < 1 Then
MsgBox "Could not find any data"
raport "No data in the grid..."
Exit Sub
End If

'change the data type to d
''mnuSetDataType_Click (2)
grid.Col = 4
grid.Row = 0
If Not (left$(frmRefine.grid.Text, 1) = "d") Then
mnuChangeDataToD_Click
''MsgBox "The data in the 4th column must be in d/A. You can compute this by the command: Data/Change data to d. Try again..."
End If
'see how many lines are there

If Not (message) Then
message = True
MsgBox "This routine calculates hkl and d's for a given space group (it uses an external program, HKLGEN, made by Armel Le Bail). You must have already the data in the column 4."
End If
v = InputBox("Input the space group symbol (standard settings, put some spaces...)", prog_name & "- hklgen", "P 1")
limit = InputBox("Input the tolerance in d (Angst.) ", prog_name, 0.02)
If limit < 0 Or limit > 1 Then Err.Raise 1101, , ""
out1 = FreeFile

''sort the data in d
mnuSortData_Click
'take the last value of d
grid.Col = 4
grid.Row = nrlines - 1
TwoThetaLimit = 1 + 2 * rd * asin(Val(txt.Text) / (2 * Val(grid.Text)))
ChDrive App.Path
ChDir (App.Path & "\hklgen\")
raport "App.Path -> " & CStr(App.Path)
raport strLinie

raport "the starting file is " & App.Path & "\hklgen\_pwdHKL.in" & vbCrLf & "the output file is " & App.Path & "\hklgen\_pwdHKL.out"
Open App.Path & "\hklgen\_pwdHKL.in" For Output As out1
Print #out1, "title : " & title
'raport "Powder-- " & title
raport "wavelength " & Val(txt.Text)
Print #out1, Val(txt.Text)
Print #out1, v
raport "space group " & v
Print #out1, txtRefine(0).Text & "  " & txtRefine(1).Text & "  " & txtRefine(2).Text & "  " & txtRefine(3).Text & "  " & txtRefine(4).Text & "  " & txtRefine(5).Text
raport "unit cell: " & txtRefine(0).Text & "  " & txtRefine(1).Text & "  " & txtRefine(2).Text & "  " & txtRefine(3).Text & "  " & txtRefine(4).Text & "  " & txtRefine(5).Text
raport "2 theta limit " & Format$(limit, "##0.0##")
'raport strLinie
Print #out1, Format$(TwoThetaLimit, "##0.0##")
Close #out1


'tt = Shell("hklgen.exe _pwdHKL ", 1)
DoEvents
''tt = ShellAndLoop("hklgen.exe _pwdHKL ", vbMaximizedFocus)
'old version
tt = ShellAndWait("hklgen.exe _pwdHKL ", vbMaximizedFocus)
''hWndShell ("hklgen _pwdHKL")
If tt = 0 Then
raport "The program pwd_Dicvol91 was called."
Else
raport "pwd_dicvol91" & " has been started." & vbCrLf & _
                     "Main window handle: " & Hex(tt) & vbCrLf & strLinie

End If

DoEvents
IamBusy False
'try to open this file,..if it can t find the file or the file is open , it gives an error ??
On Error GoTo erroropen
Close
in1 = FreeFile
j = 0
Do Until j >= 5000
Open "_pwdHKL.out" For Input As in1
'if reaches this point then the program is finished
Exit Do
Loop

''frmRefine.grid.Rows = 1
''frmRefine.grid.Rows = 50
On Error GoTo errtrap
 Line Input #in1, v
 Line Input #in1, v
 Line Input #in1, v
 Line Input #in1, v
 Line Input #in1, v
 Line Input #in1, v
nhkls = CInt(right$(v, 6))
raport CStr(nhkls) & " values computed. searching..."
Line Input #in1, v
frmRefine.grid.Col = 4
frmRefine.grid.Rows = nrlines + 2
ReDim AssignHKL(nhkls)


For i = 1 To nhkls
''frmRefine.grid.Row = i
'here reads the values
Line Input #in1, v
AssignHKL(i).h = CInt(Val(left$(v, 5)))
AssignHKL(i).k = CInt(Val(Mid$(v, 6, 4)))
AssignHKL(i).l = CInt(Val(Mid$(v, 10, 4)))
AssignHKL(i).d = CStr(Val(Mid$(v, 20, 8)))
''frmRefine.grid.Col = 1
''frmRefine.grid.Col = 2
''frmRefine.grid.Col = 3
''frmRefine.grid.Col = 4
''frmRefine.grid.Col = 5
Next i
Close
'clear the hkl grid
For i = 1 To nrlines
grid.Row = i
frmRefine.grid.Col = 1
grid.Text = ""
frmRefine.grid.Col = 2
grid.Text = ""
frmRefine.grid.Col = 3
grid.Text = ""
Next i


For i = 1 To nrlines
grid.Row = i
grid.Col = 4
For t = 1 To nhkls
If Abs(AssignHKL(t).d - Val(grid.Text)) < limit Then
grid.Col = 1
grid.Text = CStr(AssignHKL(t).h)
grid.Col = 2
grid.Text = CStr(AssignHKL(t).k)
grid.Col = 3
grid.Text = CStr(AssignHKL(t).l)
Exit For
End If
Next t
Next i

DoEvents
IamBusy False
ChDir (App.Path)
Exit Sub
erroropen:
j = j + 1
Resume Next

Close
ChDir App.Path
Exit Sub
errtrap:
raport "Error trap routine: something is wrong (wrong parameters, space group,.. or the HKLGEN is not finished). Check  _pwdHKL.out in the application directory."
raport Err.Description
Err.Clear
Close
ChDir App.Path
Exit Sub
End Sub

Private Sub mnuCellDispEnergy_Click()

'calculez celula elementara cu energy dispersive

'transform E din keV intr-un fictional theta cu lambda=1.54

''

'calculez ce trebuie si apoi revin in E; lambda nu poate fi rafinat aici insa zerro error da..

On Error GoTo handleit

Dim dist As Double, unghi As Double, D1 As Double, D2 As Double, int1 As Integer, int2 As Integer

Dim i As Integer, coderoare As Boolean, cell(7) As Double, cellr(7) As Double

''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1

Dim nrval As Integer, test As Integer, results(8) As Double, datatype As String, valoriminime() As Double

Dim pondere() As Double, intoarce(8) As Double, ind(8) As Integer, steps As Integer, refineagain As Integer, widthsearch As Double   '0,1,2 weighting scheme

Dim h() As Integer, k() As Integer, l() As Integer, teta() As Double, zero As Double, lambda As Double

Dim sumamin As Double, polcoeff(9) As Double, bout(10) As Double  '',tcalc() As Double

Dim npaf As Integer, bb() As Double, z() As Double, ii() As Double, iint As Integer

Dim F1 As Double, F2 As Double, dcalc As Double, dexp As Double, ddif As Double, thetaexp As Double, thetacalc As Double, thetacor As Double, thetadif As Double

Dim indic As Integer, unghidetector As Double

Screen.MousePointer = 11

'verific parametrii

Dim sig(8) As Double

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect starting cell parameters. Try again."

Next i

lambda = 1.54 '' dummy lambda

unghidetector = Val(txt)

steps = CInt(txtDetails(1).Text)

If unghidetector < 0 Or unghidetector > 180 Then Err.Raise 1101, , "I expect here an angle of the detector 2 theta, in degrees..."

If steps < 1 Or steps > 1000 Then Err.Raise 1101, , "Inconsistent value in the Steps field..."

'prima valoare e pentru zero, 2 e pentru lambda

'aflu numarul de valori nrval - din grid si redimensionez

unghidetector = unghidetector / 2 'transform in theta

grid.Col = 1

    For i = 1 To grid.Rows - 1

    grid.Row = i

    If grid.Text = "" Then nrval = i - 1: Exit For

    Next i

If nrval < 5 Then Err.Raise 1101, , "I don't have enough data. Check if you have at least one empty row at the end of the data. If not, correct this with Set data rows function..."



ReDim h(nrval), k(nrval), l(nrval), teta(nrval), pondere(nrval)

    For i = 1 To nrval

    grid.Row = i

    grid.Col = 1: h(i) = CInt(grid.Text)

    grid.Col = 2: k(i) = CInt(grid.Text)

    grid.Col = 3: l(i) = CInt(grid.Text)

    grid.Col = 4: teta(i) = Val(grid.Text)

    grid.Col = 5: If Val(grid.Text) < 0 Then grid.Text = "0"

    pondere(i) = Val(grid.Text)

    Next i



If ChkRef.Value Then

Call usestandard(polcoeff)

'coef polimomului, 8 este gradul maxim...

raport strLinie

raport "Correction applied to data (correction, new value):"

For i = 1 To nrval

teta(i) = teta(i) - corectie(teta(i), polcoeff)

raport Format$(Format$(Val(-corectie(teta(i), polcoeff)), "##0.000000"), "@@@@@@@@@@") & Format$(Format$(Val(teta(i)), "##0.000000"), "@@@@@@@@@@")

Next i

End If



'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

''trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

Call reciproc(cell, cellr, coderoare)

If (coderoare) Then Err.Raise 1101, , "Error in computing the reciproc cell."

raport strLinie

raport "Starting values: direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i

    

ReDim bval(8) As Double, afi(8) As Integer

For i = 1 To 6

bval(i + 2) = cell(i)

Next i

bval(1) = 0

bval(2) = lambda



chkRefine(7).Value = False

chkRefine(6).Value = False



npaf = 0

For i = 1 To 8

If chkRefine(i - 1).Value Then npaf = npaf + 1

Next i

If npaf < 1 Then raport "Nothing to do...No refinement selected..."

For i = 1 To 6

If chkRefine(i - 1).Value Then afi(i + 2) = 1

Next i

raport "Wavelength and zero are not refined.."



grid.Col = 4

grid.Row = 0



Select Case grid.Text

Case "d /A"

datatype = "d /A"

For i = 1 To nrval

teta(i) = Atn(lambda / 2 / (teta(i)) / Sqr(-lambda / 2 / teta(i) * lambda / 2 / teta(i) + 1))

teta(i) = teta(i) * rd

Next i

Case "2 theta", "theta"

Err.Raise 1101, , "Only d/A or E /keV are accepted here..."



Case "E /keV"

'calculez un teta fictiv cunosc unghidetector

For i = 1 To nrval

teta(i) = 6.199 / (teta(i) * Sin(unghidetector / rd)) 'este acum d, A

teta(i) = Atn(lambda / 2 / (teta(i)) / Sqr(-lambda / 2 / teta(i) * lambda / 2 / teta(i) + 1))

teta(i) = teta(i) * rd

Next i



End Select

DoEvents

'asta e un teta fictiv, va trebui sa revin la d si la energie



Select Case CmbCellType.ListIndex

Case 0, 3

indic = 1

Case 1, 4

indic = 2

Case Else

indic = 0

End Select



raport strLinie

raport "Summary of work..."

raport "energy dispersive data, " & "detector angle, theta " & CStr(unghidetector)

raport "Input data " & datatype

raport "number of points " & CStr(nrval)

raport "Constraints applied in refinement :" & CmbCellType.List(CmbCellType.ListIndex)

''trimit toate datele la celref.....dupa ce introduc 2theta

''least squares part!

''Call celref_based(nrval, 2, 20, bval, sig, npaf, afi, h, k, l, teta, pondere)

'un singur step, apelez de mai ulte ori ??



Call eracel_based(nrval, indic, steps, bval, bout, sig, npaf, afi, h, k, l, teta, pondere)

Dim sigma(8) As Double

results(8) = -2 * rd * bval(1)

sigma(8) = 2 * rd * sig(1)

results(7) = 2 * bval(2)

sigma(7) = 2 * sig(2)

For j = 1 To 6

results(j) = bval(j + 2)

sigma(j) = sig(j + 2)

Next j



For i = 4 To 6

results(i) = rd * results(i)

sigma(i) = rd * sigma(i)

Next i



For i = 1 To 6

If (afi(i + 2) = 0) Then sigma(i) = 0

Next i

If afi(2) = 0 Then sigma(7) = 0

If afi(1) = 0 Then sigma(8) = 0







Select Case CmbCellType.ListIndex

Case 0 'cubic

results(2) = results(1)

sigma(2) = sigma(1)

results(3) = results(1)

sigma(3) = sigma(1)

For i = 4 To 6: results(i) = 90: Next i



Case 1 'tetra

results(2) = results(1)

sigma(2) = sigma(1)



For i = 4 To 6: results(i) = 90:  Next i



Case 2 'ortho

For i = 4 To 6: results(i) = 90:  Next i



Case 3 'rhombo

sigma(3) = sigma(1)

sigma(2) = sigma(1)

sigma(5) = sigma(4)

sigma(6) = sigma(4)



results(2) = results(1)

results(3) = results(1)

results(5) = results(4)

results(6) = results(4)

Case 4 'hexa

results(2) = results(1)

sigma(2) = sigma(1)

results(4) = 90

results(5) = 90

results(6) = 60

Case 5 'mono

results(4) = 90: results(6) = 90

Case 6 'tric





End Select



Dim cell2(7) As Double

For i = 1 To 6

cellr(i) = results(i)

Next i

cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie

Call reciproc(cellr, cell, coderoare)

    

For i = 1 To 6

cellr(i) = results(i) + sigma(i)

If cell(i) < 0 Then Err.Raise 1101, , "Fatal error; negative cell parameters ?!"

Next i

cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie

Call reciproc(cellr, cell2, coderoare)

    

    For i = 1 To 6

    sigma(i) = Abs(cell(i) - cell2(i))

    Next i

    If coderoare Then raport "Error, can not compute the unit cell..."



raport "Energy dispersive data, no zero correction available"

raport "Results :"

    For i = 1 To 6

    raport Format$(Format$(Val(cell(i)), "##0.00000"), "@@@@@@@@@") & "(" & Format$(Format$(Val(sigma(i)), "0.00000"), "@@@@@@@") & ")       " & Format$(Format$(Val(cellr(i)), "##0.00000  "), "@@@@@@@@@@@")

    Next i

''results(8) = -results(8) ''????????????

raport "cell volume : " & Format$(Format$(Val(cell(7)), "##0.00000  "), "@@@@@@@@@@@")



''raport "lambda = " & Format$(Format$(Val(results(7)), "##0.00000"), "@@@@@@@@@") & "(" & Format$(Format$(sigma(7), "0.00000"), "@@@@@@@") & ")"

''raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000") & "(" & Format$(Format$(sigma(8), "0.0000"), "@@@@@@") & ")"

raport strLinie

''For i = 1 To 6

''results(i) = cellr(i)

''Next i

'tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d

raport "  H   K   L  Eexp/keV  Ecalc/keV  diff.E/keV   d calc.   d exp.    diff.d    weight"

For i = 1 To nrval

Call calculdtheta(results(1), results(2), results(3), results(4), results(5), results(6), results(7), results(8), h(i), k(i), l(i), dexp, dcalc, ddif, rd * teta(i), thetacalc, thetacor, thetadif, coderoare)

If coderoare Then raport "Error - can not evaluate parameters."

raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & " " & Format$(Format$(Val(6.199 / dexp / Sin(unghidetector / rd)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(6.199 / dcalc / Sin(unghidetector / rd)), "##0.0000  "), "@@@@@@@@@@") & "  " & Format$(Format$(Val(6.199 / dexp / Sin(unghidetector / rd) - 6.199 / dcalc / Sin(unghidetector / rd)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dcalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dexp), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(ddif), "##0.0000 "), "@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.000"), "@@@@@@@@@")

Next i

raport strLinie



Screen.MousePointer = 0

Exit Sub

handleit:

Screen.MousePointer = 0



raport Err.Description

Exit Sub

































End Sub



Private Sub mnuCELLLSTSQ_Click()
On Error GoTo handleit
'add one announcement about weight
Dim dist As Double, unghi As Double, D1 As Double, D2 As Double, int1 As Integer, int2 As Integer
Dim i As Integer, coderoare As Boolean, cell(7) As Double, cellr(7) As Double
''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1
Dim nrval As Integer, test As Integer, results(8) As Double, datatype As String, valoriminime() As Double
Dim pondere() As Double, intoarce(8) As Double, ind(8) As Integer, steps As Integer, refineagain As Integer, widthsearch As Double   '0,1,2 weighting scheme
Dim h() As Integer, k() As Integer, l() As Integer, teta() As Double, zero As Double, lambda As Double
Dim sumamin As Double, polcoeff(9) As Double, bout(10) As Double  '',tcalc() As Double
Dim npaf As Integer, bb() As Double, z() As Double, ii() As Double, iint As Integer
Dim F1 As Double, F2 As Double, dcalc As Double, dexp As Double, ddif As Double, thetaexp As Double, thetacalc As Double, thetacor As Double, thetadif As Double
Dim indic As Integer
Static weight_announce As Boolean
If Not (weight_announce) Then
MsgBox "The Weight column must contain positive values...A zero in the weigth column will remove the corresponding line from the cell refinement..." & "You can automatically add values in this or other column by the command <Data/Input data in column>"
weight_announce = True
End If

Screen.MousePointer = 11

'verific parametrii

Dim sig(8) As Double

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect starting cell parameters. Try again."

Next i

lambda = Val(txt)

steps = CInt(txtDetails(1).Text)

If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."

If steps < 1 Or steps > 1000 Then Err.Raise 1101, , "Inconsistent value in the Steps field..."

'prima valoare e pentru zero, 2 e pentru lambda

'aflu numarul de valori nrval - din grid si redimensionez

nrval = grid.Rows - 1

grid.Col = 1

    For i = 1 To grid.Rows - 1

    grid.Row = i

    If grid.Text = "" Then nrval = i - 1: Exit For

    Next i

If nrval < 5 Then Err.Raise 1101, , "I don't have enough data. Check if you have at least one empty row at the end of the data. If not, correct this with Set data rows function..."





ReDim h(nrval), k(nrval), l(nrval), teta(nrval), pondere(nrval)

    For i = 1 To nrval

    grid.Row = i

    grid.Col = 1: h(i) = CInt(grid.Text)

    grid.Col = 2: k(i) = CInt(grid.Text)

    grid.Col = 3: l(i) = CInt(grid.Text)

    grid.Col = 4: teta(i) = Val(grid.Text)

    grid.Col = 5: If Val(grid.Text) < 0 Then grid.Text = "0"

    pondere(i) = Val(grid.Text)

    Next i



If ChkRef.Value Then

Call usestandard(polcoeff)

'coef polimomului, 8 este gradul maxim...

raport strLinie

raport "Correction applied to data (correction, new value):"

For i = 1 To nrval

teta(i) = teta(i) - corectie(teta(i), polcoeff)

raport Format$(Format$(Val(-corectie(teta(i), polcoeff)), "##0.000000"), "@@@@@@@@@@") & Format$(Format$(Val(teta(i)), "##0.000000"), "@@@@@@@@@@")

Next i

End If



'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

''trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

Call reciproc(cell, cellr, coderoare)

If (coderoare) Then Err.Raise 1101, , "Error in computing the reciproc cell."

raport strLinie

raport "Starting values: direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i

    

ReDim bval(8) As Double, afi(8) As Integer

For i = 1 To 6

bval(i + 2) = cell(i)

Next i

bval(1) = 0

bval(2) = lambda



npaf = 0

For i = 1 To 8

If chkRefine(i - 1).Value Then npaf = npaf + 1

Next i

If npaf < 1 Then raport "Nothing to do...No refinement selected..."

For i = 1 To 6

If chkRefine(i - 1).Value Then afi(i + 2) = 1

Next i

If chkRefine(7).Value Then afi(1) = 1

If chkRefine(6).Value Then afi(2) = 1



''Call open_file(outputfile, 2, coderoare)

''If Not (coderoare) Then outputfile = "": Screen.MousePointer = 0: Exit Sub

''Open outputfile For Output As outfil'

''For i = 1 To 6

''results(i) = cellr(i)

''Next i

''results(8) = 0

''results(7) = lambda

grid.Col = 4

grid.Row = 0



Select Case grid.Text

Case "d /A"

datatype = "d /A"

For i = 1 To nrval

teta(i) = Atn(lambda / 2 / (teta(i)) / Sqr(-lambda / 2 / teta(i) * lambda / 2 / teta(i) + 1))

teta(i) = teta(i) * rd

Next i

Case "2 theta"

datatype = "2 theta; degrees"

For i = 1 To nrval

teta(i) = teta(i) / 2

Next i

Case "theta"

datatype = "theta; degrees"

For i = 1 To nrval

teta(i) = teta(i)

Next i

End Select

DoEvents







Select Case CmbCellType.ListIndex

Case 0, 3

indic = 1

Case 1, 4

indic = 2

Case Else

indic = 0

End Select



raport strLinie

raport "Summary of work..."

raport "Input data " & datatype

raport "number of points " & CStr(nrval)

raport "Constraints applied in refinement :" & CmbCellType.List(CmbCellType.ListIndex)

''trimit toate datele la celref.....dupa ce introduc 2theta

''least squares part!

''Call celref_based(nrval, 2, 20, bval, sig, npaf, afi, h, k, l, teta, pondere)

'un singur step, apelez de mai ulte ori ??



Call eracel_based(nrval, indic, steps, bval, bout, sig, npaf, afi, h, k, l, teta, pondere)

Dim sigma(8) As Double

results(8) = -2 * rd * bval(1)

sigma(8) = 2 * rd * sig(1)

results(7) = 2 * bval(2)

sigma(7) = 2 * sig(2)

For j = 1 To 6

results(j) = bval(j + 2)

sigma(j) = sig(j + 2)

Next j



For i = 4 To 6

results(i) = rd * results(i)

sigma(i) = rd * sigma(i)

Next i



For i = 1 To 6

If (afi(i + 2) = 0) Then sigma(i) = 0

Next i

If afi(2) = 0 Then sigma(7) = 0

If afi(1) = 0 Then sigma(8) = 0







Select Case CmbCellType.ListIndex

Case 0 'cubic

results(2) = results(1)

sigma(2) = sigma(1)

results(3) = results(1)

sigma(3) = sigma(1)

For i = 4 To 6: results(i) = 90: Next i



Case 1 'tetra

results(2) = results(1)

sigma(2) = sigma(1)



For i = 4 To 6: results(i) = 90:  Next i



Case 2 'ortho

For i = 4 To 6: results(i) = 90:  Next i



Case 3 'rhombo

sigma(3) = sigma(1)

sigma(2) = sigma(1)

sigma(5) = sigma(4)

sigma(6) = sigma(4)



results(2) = results(1)

results(3) = results(1)

results(5) = results(4)

results(6) = results(4)

Case 4 'hexa

results(2) = results(1)

sigma(2) = sigma(1)

results(4) = 90

results(5) = 90

results(6) = 60 'valori reciproce, corectat in aprilie 26

Case 5 'mono

results(4) = 90: results(6) = 90

Case 6 'tric





End Select



Dim cell2(7) As Double

For i = 1 To 6

cellr(i) = results(i)

Next i

cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie

Call reciproc(cellr, cell, coderoare)

    

For i = 1 To 6

cellr(i) = results(i) + sigma(i)

If cell(i) < 0 Then Err.Raise 1101, , "Fatal error; negative cell parameters ?!"

Next i

cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie

Call reciproc(cellr, cell2, coderoare)

    

    For i = 1 To 6

    sigma(i) = Abs(cell(i) - cell2(i))

    Next i

    

    

    

    If coderoare Then raport "Error, can not compute the unit cell..."





raport "Results :"

    

    

    

    For i = 1 To 6

    raport Format$(Format$(Val(cell(i)), "##0.00000"), "@@@@@@@@@") & "(" & Format$(Format$(Val(sigma(i)), "0.00000"), "@@@@@@@") & ")       " & Format$(Format$(Val(cellr(i)), "##0.00000  "), "@@@@@@@@@@@")

    Next i

''results(8) = -results(8) ''????????????

raport "cell volume : " & Format$(Format$(Val(cell(7)), "##0.00000  "), "@@@@@@@@@@@")

raport "lambda = " & Format$(Format$(Val(results(7)), "##0.00000"), "@@@@@@@@@") & "(" & Format$(Format$(sigma(7), "0.00000"), "@@@@@@@") & ")"

raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000") & "(" & Format$(Format$(sigma(8), "0.0000"), "@@@@@@") & ")"

raport strLinie

''For i = 1 To 6

''results(i) = cellr(i)

''Next i

'tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d

raport "  H   K   L  2theta  zero corr. 2th.calc. diff.2th.  d calc.   d exp.    diff.d      weight"

For i = 1 To nrval

Call calculdtheta(results(1), results(2), results(3), results(4), results(5), results(6), results(7), results(8), h(i), k(i), l(i), dexp, dcalc, ddif, rd * teta(i), thetacalc, thetacor, thetadif, coderoare)

If coderoare Then raport "Error - can not evaluate parameters."

raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & Format$(Format$(Val(2 * rd * teta(i)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacor), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetadif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dcalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dexp), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(ddif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.00000"), "@@@@@@@@@@@")

Next i

raport strLinie



Screen.MousePointer = 0

Exit Sub

handleit:

Screen.MousePointer = 0



raport Err.Description

Exit Sub

End Sub



 Sub mnuChangeDataTo2theta_Click()
grid.Col = 4
grid.Row = 0
ChangeData grid.Text, "2 theta", grid, 4
mnuSetDataType_Click 0

Exit Sub

End Sub



Sub ChangeData(fromData As String, toData As String, dest As Object, coloana As Integer)

On Error GoTo handleit

Dim nrdate As Integer, lambdaORtheta As Double, dist() As Double
dest.Rows = dest.Rows + 2
ReDim dist(dest.Rows)
Dim i As Integer
'din ce este transform in d

dest.Col = coloana

dest.Row = 1
DoEvents
nrdate = dest.Rows - 1
For i = 1 To dest.Rows - 1
dist(i) = Val(dest.Text)
If dist(i) <= 0 Then nrdate = i - 1: Exit For
dest.Row = dest.Row + 1
Next i

ReDim Preserve dist(nrdate)

grid.Col = 4

grid.Row = 1



Select Case fromData

Case "2 theta"

For i = 1 To nrdate

dist(i) = Val(txt) / 2 / Sin(dist(i) / 2 / rd)

Next i

Case "theta"

For i = 1 To nrdate

dist(i) = Val(txt) / 2 / Sin(dist(i) / rd)

Next i

Case "d /A"

'nimic de facut, sunt d

Case "E /keV"

For i = 1 To nrdate

dist(i) = 6.199 / (dist(i) * Sin(Val(txt) / 2 / rd))

Next i

End Select

'din d transform in toData si scriu rezultatele



Select Case toData

Case "2 theta"

For i = 1 To nrdate

dist(i) = 2 * rd * asin(Val(txt) / 2 / dist(i))

Next i

Case "theta"

For i = 1 To nrdate

dist(i) = rd * asin(Val(txt) / 2 / dist(i))

Next i

Case "d /A"

'nimic de facut, sunt d

Case "E /keV"

For i = 1 To nrdate

dist(i) = 6.199 / (dist(i) * Sin(Val(txt) / 2 / rd))

Next i

End Select



''scriu rezultatele

''dest.Col = dest.Cols - 1

dest.Row = 1

For i = 1 To nrdate

dest.Text = Format$(dist(i), "##0.00###")

dest.Row = dest.Row + 1

Next i

raport "Conversion done..."

Exit Sub

handleit:

raport "An error has occured in the conversion process..."

Exit Sub

End Sub





Private Sub mnuChangeDataToD_Click()
grid.Col = 4
grid.Row = 0

ChangeData grid.Text, "d /A", grid, 4

mnuSetDataType_Click 2



Exit Sub

End Sub



Private Sub mnuChangeDataToEnergy_Click()

grid.Col = 4

grid.Row = 0

ChangeData grid.Text, "E /keV", grid, 4

mnuSetDataType_Click 3



Exit Sub

End Sub



Private Sub mnuChangeDataToTheta_Click()



grid.Col = 4

grid.Row = 0

ChangeData grid.Text, "theta", grid, 4

mnuSetDataType_Click 1



Exit Sub

End Sub



Private Sub mnuChangeStdData2thetaToD_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "2 theta", "d /A", Grid1, t - 1



Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuChangeStdData2thetaToTheta_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "2 theta", "theta", Grid1, t - 1

Exit Sub

handleit:

Exit Sub

End Sub



Private Sub mnuChangeStdDataFromdTo2theta_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "d /A", "2 theta", Grid1, t - 1

Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuChangeStdDataFromdToenergy_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "d /A", "E /keV", Grid1, t - 1

Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuChangeStdDataFromdTotheta_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "d /A", "theta", Grid1, t - 1

Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuChangeStdDataFromEnergyTod_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "E /keV", "d /A", Grid1, t - 1

Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuChangeStdDataFromthetaTo2theta_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "theta", "2 theta", Grid1, t - 1

Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuChangeStdDataFromthetaTod_Click()

On Error GoTo handleit

Dim t As Integer

t = InputBox("Which column you want to change: 1 for Theoretical and 2 for Observed values:", prog_name, 2)

ChangeData "theta", "d /A", Grid1, t - 1

Exit Sub

handleit:

Exit Sub



End Sub



Private Sub mnuClearStd_Click()

'curat gridul standard

Dim t As Integer, i As Integer

On Error GoTo errortrap

For i = 1 To Grid1.Rows - 1

Grid1.Row = i

Grid1.Col = 0

Grid1.Text = ""

Grid1.Col = 1

Grid1.Text = ""

Next i

Exit Sub

errortrap:

raport "error in mnuClearStd: " & Err.Description

Exit Sub

End Sub



Private Sub mnuDavidonFletcher_Click()

On Error GoTo handleit

Dim i As Integer, coderoare As Boolean, cell(7) As Double, cellr(7) As Double, damp As Double

''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1

Dim nrval As Integer, test As Integer, results(8) As Double, datatype As String, valoriminime() As Double

Dim pondere() As Double, intoarce(8) As Double, ind(8) As Integer, steps As Integer, refineagain As Integer, widthsearch As Double   '0,1,2 weighting scheme

Dim h() As Integer, k() As Integer, l() As Integer, teta() As Double, zero As Double, lambda As Double

Dim sumamin As Double, polcoeff(9) As Double '',tcalc() As Double

Dim npaf As Integer, bb() As Double, z() As Double, ii() As Double, iint As Integer

Dim F1 As Double, F2 As Double, dcalc As Double, dexp As Double, ddif As Double, thetaexp As Double, thetacalc As Double, thetacor As Double, thetadif As Double

Screen.MousePointer = 11

'verific parametrii

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect starting cell parameters. Try again."

Next i

lambda = Val(txt)

''widthsearch = Val(txtDetails(0).Text)

''refineagain = CInt(txtDetails(2).Text)

steps = CInt(txtDetails(1).Text)

If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."

''If refineagain < 0 Or refineagain > 100 Then Err.Raise 1101, , "Refine value is invalid, a positive integer less than 100 is suitable."

''If widthsearch < 0 Or widthsearch > 500 Then Err.Raise 1101, , "Width search domain is invalid. "

If steps < 1 Then Err.Raise 1101, , "Inconsistent value in the Steps field..."

'prima valoare e pentru zero, 2 e pentru lambda

'aflu numarul de valori nrval - din grid si redimensionez

grid.Col = 1

    For i = 1 To grid.Rows - 2

    grid.Row = i

    If grid.Text = "" Then nrval = i - 1: Exit For

    Next i

If nrval < 5 Then Err.Raise 1101, , "I don't have enough data... ..."





ReDim h(nrval), k(nrval), l(nrval), teta(nrval), pondere(nrval)

    For i = 1 To nrval

    grid.Row = i

    grid.Col = 1: h(i) = CInt(grid.Text)

    grid.Col = 2: k(i) = CInt(grid.Text)

    grid.Col = 3: l(i) = CInt(grid.Text)

    grid.Col = 4: teta(i) = Val(grid.Text)

    grid.Col = 5: If Val(grid.Text) < 0 Then grid.Text = "0"

    pondere(i) = Val(grid.Text)

    Next i



If ChkRef.Value Then

Call usestandard(polcoeff)

'coef polimomului, 8 este gradul maxim...

raport strLinie

raport "Correction applied to data (correction, new value):"

For i = 1 To nrval

teta(i) = teta(i) - corectie(teta(i), polcoeff)

raport Format$(Format$(Val(-corectie(teta(i), polcoeff)), "##0.000000"), "@@@@@@@@@@") & Format$(Format$(Val(teta(i)), "##0.000000"), "@@@@@@@@@@")

Next i

End If

raport strLinie

'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

''trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

Call reciproc(cell, cellr, coderoare)

If (coderoare) Then Err.Raise 1101, , "Error in computing the reciproc cell."

raport strLinie

raport "Direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i



npaf = 0

For i = 1 To 8

If chkRefine(i - 1).Value Then npaf = npaf + 1: ind(i) = 1

Next i

If npaf < 1 Then Err.Raise 1101, , "Nothing to do...No refinement selected..."

''Call open_file(outputfile, 2, coderoare)

''If Not (coderoare) Then outputfile = "": Screen.MousePointer = 0: Exit Sub

''Open outputfile For Output As outfil'

For i = 1 To 6

results(i) = cellr(i)

Next i

results(8) = 0

results(7) = lambda

grid.Col = 4

grid.Row = 0



Select Case grid.Text

Case "d /A"

datatype = "d"

For i = 1 To nrval

teta(i) = Atn(lambda / 2 / (teta(i)) / Sqr(-lambda / 2 / teta(i) * lambda / 2 / teta(i) + 1))

teta(i) = teta(i) * rd

Next i

Case "2 theta"

datatype = "2"

For i = 1 To nrval

teta(i) = teta(i) / 2

Next i

Case "theta"

datatype = "t"

For i = 1 To nrval

teta(i) = teta(i)

Next i

End Select

DoEvents

''trimit toate datele la celref.....dupa ce introduc 2theta

''least squares part!

ReDim z(nrval, 8), ii(nrval), solutie(8) As Double

damp = 1

ReDim valoriminime(steps + 1, 9)

   

     

      ae = results(1)

      be = results(2)

      ce = results(3)

      cae = Cos(results(4) / rd)

      cbe = Cos(results(5) / rd)

      cce = Cos(results(6) / rd)

ReDim q(nrval, 8) As Double

Dim f As Double, ier As Integer, nef As Integer



''f este valoarea functiei...

''ier si nef sunt niste indicatori intregi

ReDim solutie(npaf)

Call minfp(npaf, nrval, f, results, solutie, 0.0001, 0.000000001, steps, ier, nef, ind, h, k, l, teta, pondere)



raport "la iesire din fp, functia este " & CStr(f)

For i = 1 To npaf

raport "solutie" & CStr(solutie(i))

Next i





Select Case CmbCellType.ListIndex

Case 0 'cubic

results(2) = results(1)

results(3) = results(1)

For i = 4 To 6: results(i) = 90: Next i



Case 1 'tetra

results(2) = results(1)

For i = 4 To 6: results(i) = 90: Next i



Case 2 'ortho

For i = 4 To 6: results(i) = 90: Next i



Case 3 'rhombo

results(2) = results(1)

results(3) = results(1)

results(5) = results(4)

results(6) = results(4)

Case 4 'hexa

results(2) = results(1)



results(4) = 90

results(5) = 90

results(6) = 120

Case 5 'mono

results(4) = 90: results(6) = 90

Case 6 'tric





End Select









sumamin = 0

For i = 1 To nrval

    sumamin = sumamin + (pondere(i) * (h(i) * h(i) * results(1) * results(1) + k(i) * k(i) * results(2) * results(2) + l(i) * l(i) * results(3) * results(3) + 2 * l(i) * h(i) * results(3) * results(1) * Cos(results(5) / rd) + 2 * l(i) * k(i) * results(3) * results(2) * Cos(results(4) / rd) + 2 * h(i) * k(i) * results(2) * results(1) * Cos(results(6) / rd) - pondere(i) * 4 * Sin((teta(i) + results(8) / 2) / rd) * Sin((teta(i) + results(8) / 2) / rd) / results(7) / results(7))) ^ 2

Next i

raport "sq. dev. at cycle " & CStr(jj) & "  " & CStr(sumamin)

''results(8) = -results(8)

valoriminime(jj, 1) = sumamin

For i = 1 To 8: valoriminime(jj, i + 1) = results(i): Next i

''Next iint



''For i = 1 To 6

''cellr(i) = results(i)

''Next i

''cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

''raport strLinie

''raport "Results obtained after " & CStr(steps) & " cycle(s) :"

''Call reciproc(cellr, cell, coderoare)

''    If coderoare Then raport "Error, can not compute the unit cell..."

''    For i = 1 To 7

''    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

''    Next i



''raport "lambda = " & CStr(results(7))

''raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000")

''raport strLinie

''tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d

''raport "  H   K   L  2theta  zero corr. 2th.calc. diff.2th.  d calc.   d exp.    diff.d      weight"

''For i = 1 To nrval

''Call calculdtheta(results(1), results(2), results(3), results(4), results(5), results(6), results(7), results(8), h(i), k(i), l(i), dexp, dcalc, ddif, teta(i), thetacalc, thetacor, thetadif, coderoare)

''If coderoare Then raport "Error - can not evaluate parameters."

''raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & Format$(Format$(Val(2 * teta(i)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacor), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetadif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dcalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dexp), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(ddif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.00000"), "@@@@@@@@@@@")

''Next i

raport strLinie

''raport strLinie



'selectez cele mai mici reziduals





For i = 1 To steps - 1

If valoriminime(i, 1) < valoriminime(i + 1, 1) Then

For j = 1 To 8

results(j) = valoriminime(i, j + 1)

Next j

End If

Next i



For i = 1 To 6

cellr(i) = results(i)

Next i

cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie

raport "Results with minimum  sq. deviation:"

Call reciproc(cellr, cell, coderoare)

    If coderoare Then raport "Error, can not compute the unit cell..."

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i

''results(8) = -results(8) ''????????????

raport "lambda = " & CStr(results(7))

raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000")

raport strLinie



'tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d

raport "  H   K   L  2theta  zero corr. 2th.calc. diff.2th.  d calc.   d exp.    diff.d      weight"

For i = 1 To nrval

Call calculdtheta(results(1), results(2), results(3), results(4), results(5), results(6), results(7), results(8), h(i), k(i), l(i), dexp, dcalc, ddif, teta(i), thetacalc, thetacor, thetadif, coderoare)

If coderoare Then raport "Error - can not evaluate parameters."





''unghi = -results(8) / 2 + calcul_theta(results(7), h(i), k(i), l(i), results(1), results(2), results(3), results(4), results(5), results(6))

'unghi este theta calculat cu parametrii astia,correctat pentru zero

''d1 = (results(7) / 2 / (Sin((teta(i) + results(8) / 2) / rd))) 'd1 este d introdus

''d2 = (results(7) / 2 / (Sin((unghi) / rd))) 'd2 este d calculat

raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & Format$(Format$(Val(2 * teta(i)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacor), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetadif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dcalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dexp), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(ddif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.00000"), "@@@@@@@@@@@")

Next i

raport strLinie

'calculez abaterea standard pentru fiecare parametru obtinu



Dim paramsig(100) As Double

''Call sigma(results, 1, nrval, teta(), h(), k(), l(), paramsig)

    

    'calculez amediu-ai totul la patrat, suma,..etc

Screen.MousePointer = 0

Exit Sub

handleit:

Screen.MousePointer = 0

If Err = 1101 Then

raport Err.Description

Exit Sub

Else

raport "Unexpected error in mnuLeastSquares routine."

End If

Exit Sub

End Sub



Private Sub mnuDelColumn_Click()

delete_column grid

End Sub



Private Sub mnuDeleteRow_Click()

delete_linie grid



End Sub



Sub delete_column(grid As Object)

Dim t As Integer, i As Integer, j As Integer, carr(5) As String

On Error GoTo errortrap

t = 4

t = InputBox("Which column you want to delete ? (1 is H). Warning: no UNDO available ", prog_name, 1)

If t < 6 Or t > 0 Then

grid.Col = t

jcount = 0

''grid.Rows = grid.Rows + 1

For i = 1 To grid.Rows - 2

grid.Row = i

grid.Text = ""

Next i

End If

Exit Sub

errortrap:

ignoralinii = 0

raport "Error, mnuDeleteColumn routine ..."

Err.Clear

Exit Sub





End Sub



Sub delete_linie(grid As Object)

Dim t As Integer, i As Integer, j As Integer, carr(5) As String

On Error GoTo errortrap

t = 1

t = InputBox("Position in grid ?", prog_name, 1)

If t < grid.Rows - 2 Then



grid.Rows = grid.Rows + 1

For i = t + 1 To grid.Rows - 2

grid.Row = i

For j = 1 To 5

grid.Col = j

carr(j) = grid.Text

Next j

grid.Row = grid.Row - 1

For j = 1 To 5

grid.Col = j

grid.Text = carr(j)

Next j

Next i

End If

Exit Sub



errortrap:

ignoralinii = 0

raport "Error, mnuDeleteRow routine ..."

Err.Clear

Exit Sub



End Sub



 Sub mnuErase_Click()
'curat gridul si numerotez din nou
Dim t As Integer, i As Integer
On Error GoTo errortrap
grid.Rows = 1
DoEvents
grid.Rows = 100

For i = 1 To grid.Rows - 1
grid.Row = i
grid.Col = 0
grid.Text = CStr(i)
For t = 1 To 5
grid.Col = t
grid.Text = ""
Next t
Next i
DoEvents
Exit Sub

errortrap:

Exit Sub

End Sub



Private Sub mnuExport_Click()

Dim returncode As Boolean, i As Integer, linie As String

On Error GoTo errortrap

raport strLinie

raport "Export data from grid...."

outfil = FreeFile

Call open_file(outputfile, 2, returncode)

If Not (returncode) Then Exit Sub

raport "The file is " & outputfile

Open outputfile For Output As outfil

linie = "H"

grid.Row = 0

For ii% = 2 To 5

grid.Col = ii%

linie = linie + " " + Format$(grid.Text, "@@@@@@@@")

Next ii%

Print #outfil, linie



For i = 1 To grid.Rows - 1

grid.Row = i

grid.Col = 1

If grid.Text = "" Then Exit For

linie = 0

linie = grid.Text

For ii% = 2 To 5

grid.Col = ii%

linie = linie + " " + Format$(grid.Text, "@@@@@@@@")

Next ii%

Print #outfil, linie

Next i

raport "Done..."

Close

Exit Sub

errortrap:

Close

Exit Sub

End Sub



Private Sub mnuExportDataToGraphic_Click()
'see how many peaks are here
Dim i As Integer
On Error GoTo errortrap
Call numargrid(i)
If i < 2 Then Err.Raise 1101, , "No peaks to send..."
''order as 2theta
NumberOfSimulatedPeaks = i - 1
If mnuChangeDataTo2theta.Enabled Then mnuChangeDataTo2theta_Click
FrmGraph.Show
DoEvents

amsentpeaks = True
For i = 1 To NumberOfSimulatedPeaks
grid.Row = i
grid.Col = 1
valori(i).h = CStr(Val(grid.Text))
grid.Col = 2
valori(i).k = CStr(Val(grid.Text))
grid.Col = 3
valori(i).l = CStr(Val(grid.Text))
grid.Col = 4
valori(i).doitheta = CStr(Val(grid.Text))
grid.Col = 5
valori(i).ygrec = CStr(Val(grid.Text))
Next i


FrmGraph.mnuGraphRefresh_Click
DoEvents
Exit Sub
errortrap:
If Err.Number = 1101 Then MsgBox Err.Description
Exit Sub

End Sub

Private Sub mnuExportStdGridToAscii_Click()

Dim returncode As Boolean, i As Integer, linie As String

On Error GoTo errortrap

raport strLinie

raport "Export standard data from grid...."

outfil = FreeFile

Call open_file(outputfile, 2, returncode)

If Not (returncode) Then Exit Sub

raport "The file is " & outputfile

Open outputfile For Output As outfil

Print #outfil, "theoretical    observed"



For i = 1 To Grid1.Rows - 1

Grid1.Row = i

Grid1.Col = 0

If grid.Text = "" Then Exit For

linie = grid.Text

grid.Col = 1

linie = linie + " " + Format$(grid.Text, "@@@@@@@@")

Print #outfil, linie

Next i

raport "Done..."

Close

Exit Sub

errortrap:

Close

Exit Sub



End Sub



Private Sub mnuGenerate_Click()

putdata 1

End Sub







Private Sub mnuHKLGENDataGrid_Click()
End Sub

Private Sub mnuIgnore_Click()
Dim t As Integer
On Error GoTo errortrap
t = 0
t = InputBox("How many lines want to ignore ?", prog_name, 1)
Select Case CInt(t)
Case 0
mnuIgnore.Caption = "&Ignore : none"
Case 1
mnuIgnore.Caption = "&Ignore first line"
Case Else
mnuIgnore.Caption = "&Ignore the first " & CInt(t) & " lines"
End Select
Exit Sub


errortrap:
ignoralinii = 0

mnuIgnore.Caption = "&Ignore : none"

Err.Clear

Exit Sub



End Sub



Private Sub mnuImport5_Click()

'verific cate linii trebuie sa ignor

On Error GoTo handleit

Dim return_code As Boolean, linie As String, ignor As Integer

Dim a1 As Integer, a2 As Integer, a3 As Integer, a4 As Double, a5 As Double

inpfil = FreeFile

Call open_file(inputfile, 1, return_code)

If Not (return_code) Then inputfile = "": Exit Sub

raport strLinie

raport inputfile & " open; this file must have h,k,l and d, theta or 2theta, weigth"

ignor = 0

If Len(mnuIgnore.Caption) > 14 Then

ignor = 1

If Len(mnuIgnore.Caption) > 18 Then ignor = Val(right$(mnuIgnore.Caption, Len(mnuIgnore.Caption) - 17))

End If

raport CStr(ignor) & " line(s) will be ignored."

raport "One data set per line is expected, without the intensity."

raport "The data will be appended to grid."

grid.Col = 1

For i% = 1 To grid.Rows - 2

grid.Row = i%

If grid.Text = "" Then grid.Row = grid.Row: Exit For

Next i%



Open inputfile For Input As inpfil

If ignor > 0 Then

For i% = 1 To ignor

Line Input #inpfil, linie

Next i%

End If



Do While Not (EOF(inpfil))

Input #inpfil, a1, a2, a3, a4, a5

grid.Rows = grid.Rows + 1

grid.Col = 1: grid.Text = a1

grid.Col = 2: grid.Text = a2

grid.Col = 3: grid.Text = a3

grid.Col = 4: grid.Text = a4

grid.Col = 5: grid.Text = a5

grid.Row = grid.Row + 1

Loop

Close

Exit Sub

handleit:

Close

Exit Sub



End Sub



Private Sub mnuImportAscii_Click()

'verific cate linii trebuie sa ignor

On Error GoTo handleit

Dim return_code As Boolean, linie As String, ignor As Integer

Dim a1 As Integer, a2 As Integer, a3 As Integer, a4 As Double

inpfil = FreeFile

Call open_file(inputfile, 1, return_code)

If Not (return_code) Then inputfile = "": Exit Sub

raport strLinie

raport inputfile & " open; this file must have h,k,l and d, theta or 2theta"

ignor = 0

If Len(mnuIgnore.Caption) > 14 Then

ignor = 1

If Len(mnuIgnore.Caption) > 18 Then ignor = Val(right$(mnuIgnore.Caption, Len(mnuIgnore.Caption) - 17))

End If

raport CStr(ignor) & " line(s) will be ignored."

raport "One data set per line is expected, without the intensity."

raport "The data will be appended to grid."

grid.Col = 1

For i% = 1 To grid.Rows - 2

grid.Row = i%

If grid.Text = "" Then grid.Row = grid.Row: Exit For

Next i%



Open inputfile For Input As inpfil

If ignor > 0 Then

For i% = 1 To ignor

Line Input #inpfil, linie

Next i%

End If



Do While Not (EOF(inpfil))

Input #inpfil, a1, a2, a3, a4

grid.Rows = grid.Rows + 1

grid.Col = 1: grid.Text = a1

grid.Col = 2: grid.Text = a2

grid.Col = 3: grid.Text = a3

grid.Col = 4: grid.Text = a4

grid.Col = 5: grid.Text = "1"

grid.Row = grid.Row + 1

Loop

Close

Exit Sub

handleit:

Close

Exit Sub



End Sub



Private Sub mnuImportStd_Click()

'verific cate linii trebuie sa ignor

On Error GoTo handleit

raport "2 columns of data are expected"

Dim return_code As Boolean, linie As String, ignor As Integer

Dim a1 As Double, a2 As Double

inpfil = FreeFile

Call open_file(inputfile, 1, return_code)

If Not (return_code) Then inputfile = "": Exit Sub

raport strLinie

raport inputfile & " open; this file must have internal standard data."

ignor = 0

If Len(mnuIgnore.Caption) > 14 Then

ignor = 1

If Len(mnuIgnore.Caption) > 18 Then ignor = Val(right$(mnuIgnore.Caption, Len(mnuIgnore.Caption) - 17))

End If

raport CStr(ignor) & " line(s) will be ignored."

raport "The data will be added to standard grid."

Grid1.Col = 1

For i% = 1 To Grid1.Rows - 1

Grid1.Row = i%

If Grid1.Text = "" Then Grid1.Row = Grid1.Row: Exit For

Next i%



Open inputfile For Input As inpfil

If ignor > 0 Then

For i% = 1 To ignor

Line Input #inpfil, linie

Next i%

End If



Do While Not (EOF(inpfil))

Input #inpfil, a1, a2

Grid1.Rows = Grid1.Rows + 1

Grid1.Col = 0: Grid1.Text = a1

Grid1.Col = 1: Grid1.Text = a2

Grid1.Row = Grid1.Row + 1

Loop

Close

Exit Sub

handleit:

raport "Unexpected error...."

Close

Exit Sub



End Sub



Private Sub mnuIndexDicvol_Click()
frmDicvolSetup.Show
End Sub

Private Sub mnuInputStdData_Click()
On Error GoTo handleit
Dim return_code As Boolean, linie As String, ignor As Integer
Dim a1 As Single
Dim k As Integer, j As Integer, t As Integer
k = 0
t = InputBox("In which column you want to input data (1 or 2) :", prog_name, 1)
If Not (t < 3 And t > 0) Then Err.Raise 1101, , "Try again later..."
t = t - 1
k = InputBox("Read one data, skip n... Insert n (0 if you want all data):", prog_name, 0)
If k < 0 Or k > 10 Then Err.Raise 1101, , "Try again later, incorrect value <0 to 100>..."
inpfil = FreeFile
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie
raport inputfile & " open; read one data, skip " & CStr(k)
ignor = 0
If Len(mnuIgnore.Caption) > 14 Then
ignor = 1
If Len(mnuIgnore.Caption) > 18 Then ignor = Val(right$(mnuIgnore.Caption, Len(mnuIgnore.Caption) - 17))
End If
raport CStr(ignor) & " line(s) will be ignored."
Grid1.Col = t
Grid1.Row = 1
Open inputfile For Input As inpfil
If ignor > 0 Then
For i% = 1 To ignor
Line Input #inpfil, linie
Next i%
End If

Do While Not (EOF(inpfil))
Input #inpfil, a1
j = MsgBox(CStr(a1) & vbCrLf & "This is the first data you want ?", vbYesNoCancel + vbDefaultButton2, prog_name)
If j = vbCancel Then Err.Raise 1101, , "Cancel..."
If j = vbYes Then Exit Do
Loop

DoEvents
Grid1.Text = CStr(a1)
Grid1.Row = Grid1.Row + 1
DoEvents
For i = 1 To k
Input #inpfil, a1
Next i
Do While Not (EOF(inpfil))
Input #inpfil, a1
Grid1.Text = CStr(a1)
Grid1.Rows = Grid1.Rows + 1
Grid1.Row = Grid1.Row + 1
For i = 1 To k
Input #inpfil, a1
Next i
Loop
Close
Exit Sub

handleit:
Close
Exit Sub
End Sub

Private Sub mnuInsertRow_Click()
insert_linie grid
End Sub

Sub insert_linie(grid As Object)
Dim t As Integer, i As Integer, j As Integer, carr(5) As String
On Error GoTo errortrap
t = 1
t = InputBox("Position in grid ?", prog_name, 1)
If t < grid.Rows - 2 Then
grid.Rows = grid.Rows + 1
For i = grid.Rows - 2 To t Step -1
grid.Row = i
For j = 1 To 5
grid.Col = j
carr(j) = grid.Text
Next j
grid.Row = grid.Row + 1

For j = 1 To 5

grid.Col = j

grid.Text = carr(j)

Next j

Next i

grid.Row = t

For i = 1 To 5

grid.Col = i



grid.Text = ""

Next i

End If



Exit Sub



errortrap:

ignoralinii = 0

mnuIgnore.Caption = "&Ignore : none"

Err.Clear

Exit Sub





End Sub



Private Sub mnuLeastSquares_Click()
On Error GoTo handleit
Dim dist As Double, unghi As Double, D1 As Double, D2 As Double, int1 As Integer, int2 As Integer
Dim i As Integer, coderoare As Boolean, cell(7) As Double, cellr(7) As Double, damp As Double
''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1
Dim nrval As Integer, test As Integer, results(8) As Double, datatype As String, valoriminime() As Double
Dim pondere() As Double, intoarce(8) As Double, ind(8) As Integer, steps As Integer, refineagain As Integer, widthsearch As Double   '0,1,2 weighting scheme
Dim h() As Integer, k() As Integer, l() As Integer, teta() As Double, zero As Double, lambda As Double
Dim sumamin As Double, polcoeff(9) As Double '',tcalc() As Double
Dim npaf As Integer, bb() As Double, z() As Double, ii() As Double, iint As Integer
Dim F1 As Double, F2 As Double, dcalc As Double, dexp As Double, ddif As Double, thetaexp As Double, thetacalc As Double, thetacor As Double, thetadif As Double

Static weight_announce As Boolean
If Not (weight_announce) Then
MsgBox "The Weight column must contain positive values...A zero in the weigth column will remove the corresponding line from the cell refinement..." & "You can automatically add values in this or other column by the command <Data/Input data in column>"
weight_announce = True
End If



Screen.MousePointer = 11

'verific parametrii

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect starting cell parameters. Try again."

Next i

lambda = Val(txt)

''widthsearch = Val(txtDetails(0).Text)

''refineagain = CInt(txtDetails(2).Text)

steps = CInt(txtDetails(1).Text)

If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."

''If refineagain < 0 Or refineagain > 100 Then Err.Raise 1101, , "Refine value is invalid, a positive integer less than 100 is suitable."

''If widthsearch < 0 Or widthsearch > 500 Then Err.Raise 1101, , "Width search domain is invalid. "

If steps < 1 Then Err.Raise 1101, , "Inconsistent value in the Steps field..."

'prima valoare e pentru zero, 2 e pentru lambda

'aflu numarul de valori nrval - din grid si redimensionez

grid.Col = 1

nrval = grid.Rows - 1



    For i = 1 To grid.Rows - 1

    grid.Row = i

    If grid.Text = "" Then nrval = i - 1: Exit For

    Next i

If nrval < 5 Then Err.Raise 1101, , "I don't have enough data... ..."





ReDim h(nrval), k(nrval), l(nrval), teta(nrval), pondere(nrval)

    For i = 1 To nrval

    grid.Row = i

    grid.Col = 1: h(i) = CInt(grid.Text)

    grid.Col = 2: k(i) = CInt(grid.Text)

    grid.Col = 3: l(i) = CInt(grid.Text)

    grid.Col = 4: teta(i) = Val(grid.Text)

    grid.Col = 5: If Val(grid.Text) < 0 Then grid.Text = "0"

    pondere(i) = Val(grid.Text)

    Next i



If ChkRef.Value Then

Call usestandard(polcoeff)

'coef polimomului, 8 este gradul maxim...

raport strLinie

raport "Correction applied to data (correction, new value):"

For i = 1 To nrval

teta(i) = teta(i) - corectie(teta(i), polcoeff)

raport Format$(Format$(Val(-corectie(teta(i), polcoeff)), "##0.000000"), "@@@@@@@@@@") & Format$(Format$(Val(teta(i)), "##0.000000"), "@@@@@@@@@@")

Next i

End If

raport strLinie

'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

''trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

reciproc cell, cellr, coderoare

If (coderoare) Then Err.Raise 1101, , "Error in computing the reciproc cell."

raport strLinie

raport "Direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i



npaf = 0

For i = 1 To 8

If chkRefine(i - 1).Value Then npaf = npaf + 1: ind(i) = 1

Next i

If npaf < 1 Then Err.Raise 1101, , "Nothing to do...No refinement selected..."

''Call open_file(outputfile, 2, coderoare)

''If Not (coderoare) Then outputfile = "": Screen.MousePointer = 0: Exit Sub

''Open outputfile For Output As outfil'

For i = 1 To 6

results(i) = cellr(i)

Next i

results(8) = 0

results(7) = lambda

grid.Col = 4

grid.Row = 0



Select Case grid.Text

Case "d /A"

datatype = "d"

For i = 1 To nrval

teta(i) = Atn(lambda / 2 / (teta(i)) / Sqr(-lambda / 2 / teta(i) * lambda / 2 / teta(i) + 1))

teta(i) = teta(i) * rd

Next i

Case "2 theta"

datatype = "2"

For i = 1 To nrval

teta(i) = teta(i) / 2

Next i

Case "theta"

datatype = "t"

For i = 1 To nrval

teta(i) = teta(i)

Next i

End Select

DoEvents

''trimit toate datele la celref.....dupa ce introduc 2theta

''least squares part!

ReDim z(nrval, 8), ii(nrval), solutie(8) As Double

damp = 1

ReDim valoriminime(steps + 1, 9)

   

 For jj = 1 To steps

     

      ae = results(1)

      be = results(2)

      ce = results(3)

      cae = Cos(results(4) / rd)

      cbe = Cos(results(5) / rd)

      cce = Cos(results(6) / rd)

ReDim q(nrval, 8) As Double

Dim f As Double

     '' If (indic = 1 Or indic = 2) Then b(4) = b(3)

      ''If (indic = 1) Then b(5) = b(3)

  

   

          ''DO 2 I=1,NR

      For i = 1 To nrval

    ''  dd = (ae * h(i)) ^ 2 + (be * k(i)) ^ 2 + (ce * l(i)) ^ 2 + 2# * (h(i) * k(i) * ae * be * cce + k(i) * l(i) * be * ce * cae + l(i) * h(i) * ce * ae * cbe)

    ''

    ''  d = 1# / Sqr(dd)

   ''   rad = Sqr(Abs((1# - (0.25 * results(7) ^ 2) * dd))) ''????????

   ''   f = 0.5 * results(7) * d / rad

 ''     q(i, 8) = -8 / (results(7) ^ 2) * Sin(((teta(i) - results(8)) / rd)) * Cos((teta(i) - results(8)) / rd)

q(i, 8) = f_derivata(0.0001, results, h(i), k(i), l(i), teta(i), 8, 1)

 q(i, 7) = f_derivata(0.0001, results, h(i), k(i), l(i), teta(i), 7, 1)

    

''      q(i, 7) = 8 * (Sin(teta(i) - results(8) / 2) / rd) ^ 2 / (results(7) ^ 3)



      q(i, 1) = 2 * h(i) * (h(i) * ae + k(i) * be * cce + l(i) * ce * cbe)

      q(i, 2) = 2 * k(i) * (k(i) * be + l(i) * ce * cae + h(i) * ae * cce)

      q(i, 3) = 2 * l(i) * (l(i) * ce + h(i) * ae * cbe + k(i) * be * cae)

      q(i, 4) = -2 * k(i) * l(i) * be * ce * Sin(results(4) / rd)

      q(i, 5) = -2 * l(i) * h(i) * ce * ae * Sin(results(5) / rd)

      q(i, 6) = -2 * h(i) * k(i) * ae * be * Sin(results(6) / rd)

''2     qq(i, npaf2) = b(1) + Atn((b(2) / d) / Sqr(-(b(2) / d) * (b(2) / d) + 1))

 ii(i) = -functie(results, pondere(i), h(i), k(i), l(i), teta(i))

     

      Next i

      For ir = 1 To nrval

      ''DO 3 IR=1,NR

      j = 0

      For i = 1 To 8

      ''DO 3 I=1,8

      If (ind(i) = 1) Then j = j + 1:         z(ir, j) = pondere(ir) * q(ir, i)





 Next i

Next ir

Call pseudoinv(nrval, npaf, z, ii, solutie, 0.000000000001, coderoare)

j = 0

For i = 1 To 8

      If (ind(i) = 1) Then j = j + 1:  results(i) = results(i) + ind(i) * solutie(j) / damp



Next i



results(8) = -results(8)

Select Case CmbCellType.ListIndex

Case 0 'cubic

results(2) = results(1)

results(3) = results(1)

For i = 4 To 6: results(i) = 90: Next i



Case 1 'tetra

results(2) = results(1)

For i = 4 To 6: results(i) = 90: Next i



Case 2 'ortho

For i = 4 To 6: results(i) = 90: Next i



Case 3 'rhombo

results(2) = results(1)

results(3) = results(1)

results(5) = results(4)

results(6) = results(4)

Case 4 'hexa

results(2) = results(1)



results(4) = 90

results(5) = 90

results(6) = 60 ''corectat, e valoarea reciproca !!

Case 5 'mono

results(4) = 90: results(6) = 90

Case 6 'tric





End Select









sumamin = 0

For i = 1 To nrval

    sumamin = sumamin + (pondere(i) * (h(i) * h(i) * results(1) * results(1) + k(i) * k(i) * results(2) * results(2) + l(i) * l(i) * results(3) * results(3) + 2 * l(i) * h(i) * results(3) * results(1) * Cos(results(5) / rd) + 2 * l(i) * k(i) * results(3) * results(2) * Cos(results(4) / rd) + 2 * h(i) * k(i) * results(2) * results(1) * Cos(results(6) / rd) - pondere(i) * 4 * Sin((teta(i) + results(8) / 2) / rd) * Sin((teta(i) + results(8) / 2) / rd) / results(7) / results(7))) ^ 2

Next i

raport "sq. dev. at cycle " & CStr(jj) & "  " & CStr(sumamin)

''results(8) = -results(8)

valoriminime(jj, 1) = sumamin

For i = 1 To 8: valoriminime(jj, i + 1) = results(i): Next i

If coderoare Then raport "Unexpected error in PseudoInverse routine; still trying..."

''Next iint



Next jj







''For i = 1 To 6

''cellr(i) = results(i)

''Next i

''cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

''raport strLinie

''raport "Results obtained after " & CStr(steps) & " cycle(s) :"

''Call reciproc(cellr, cell, coderoare)

''    If coderoare Then raport "Error, can not compute the unit cell..."

''    For i = 1 To 7

''    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

''    Next i



''raport "lambda = " & CStr(results(7))

''raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000")

''raport strLinie

''tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d

''raport "  H   K   L  2theta  zero corr. 2th.calc. diff.2th.  d calc.   d exp.    diff.d      weight"

''For i = 1 To nrval

''Call calculdtheta(results(1), results(2), results(3), results(4), results(5), results(6), results(7), results(8), h(i), k(i), l(i), dexp, dcalc, ddif, teta(i), thetacalc, thetacor, thetadif, coderoare)

''If coderoare Then raport "Error - can not evaluate parameters."

''raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & Format$(Format$(Val(2 * teta(i)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacor), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetadif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dcalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dexp), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(ddif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.00000"), "@@@@@@@@@@@")

''Next i

raport strLinie

''raport strLinie



'selectez cele mai mici reziduals





For i = 1 To steps - 1

If valoriminime(i, 1) < valoriminime(i + 1, 1) Then

For j = 1 To 8

results(j) = valoriminime(i, j + 1)

Next j

End If

Next i



For i = 1 To 6

cellr(i) = results(i)

Next i

cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie

raport "Results with minimum  sq. deviation:"

Call reciproc(cellr, cell, coderoare)

    If coderoare Then raport "Error, can not compute the unit cell..."

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i

''results(8) = -results(8) ''????????????

raport "lambda = " & CStr(results(7))

raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000")

raport strLinie



'tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d

raport "  H   K   L  2theta  zero corr. 2th.calc. diff.2th.  d calc.   d exp.    diff.d      weight"

For i = 1 To nrval

Call calculdtheta(results(1), results(2), results(3), results(4), results(5), results(6), results(7), results(8), h(i), k(i), l(i), dexp, dcalc, ddif, teta(i), thetacalc, thetacor, thetadif, coderoare)

If coderoare Then raport "Error - can not evaluate parameters."





''unghi = -results(8) / 2 + calcul_theta(results(7), h(i), k(i), l(i), results(1), results(2), results(3), results(4), results(5), results(6))

'unghi este theta calculat cu parametrii astia,correctat pentru zero

''d1 = (results(7) / 2 / (Sin((teta(i) + results(8) / 2) / rd))) 'd1 este d introdus

''d2 = (results(7) / 2 / (Sin((unghi) / rd))) 'd2 este d calculat

raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & Format$(Format$(Val(2 * teta(i)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacor), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetacalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * thetadif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dcalc), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(dexp), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(ddif), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.00000"), "@@@@@@@@@@@")

Next i

raport strLinie

'calculez abaterea standard pentru fiecare parametru obtinu



Dim paramsig(100) As Double

''Call sigma(results, 1, nrval, teta(), h(), k(), l(), paramsig)

    

    'calculez amediu-ai totul la patrat, suma,..etc

Screen.MousePointer = 0

Exit Sub

handleit:

Screen.MousePointer = 0

If Err = 1101 Then

raport Err.Description

Exit Sub

Else

raport "Unexpected error in mnuLeastSquares routine."

End If

Exit Sub







End Sub







Private Sub mnuMultiplycol_Click()

Dim t As Integer, i As Integer, j As Integer, txt As String

On Error GoTo errortrap

t = 4

t = InputBox("Insert the data column you want to modify (H is in 1) :", prog_name, 4)

j = InputBox("Insert the number of column you want to multiply with (attention, data may be unusable after this operation) :", prog_name, 1)

If t < 6 And t > 0 And j < 6 And j > 0 Then

For i = 1 To grid.Rows - 1

grid.Col = j



grid.Row = i

txt = (Val(grid.Text))

If grid.Text = "" Then Exit For

grid.Col = t

grid.Text = Format$(Val(grid.Text) * Val(txt), "####0.0000#")

Next i

End If

Exit Sub

errortrap:

raport "Error in mnuMultiply _click routine."

Exit Sub



End Sub



Private Sub mnuMultiplyconst_Click()

Dim t As Integer, i As Integer, j As Single

On Error GoTo errortrap

t = 4

t = InputBox("Insert the data column you want to modify (H is in 1) :", prog_name, 4)

j = InputBox("Insert the value you want to multiply with (attention, data may be unusable after this operation) :", prog_name, 1)

If t < 6 And t > 0 Then

grid.Col = t

For i = 1 To grid.Rows - 1

grid.Row = i

If grid.Text = "" Then Exit For

grid.Text = Format$(Val(grid.Text) * Val(j), "####0.0000#")

Next i

End If

Exit Sub

errortrap:

raport "Error in mnuMultiplyconst _click routine."

Exit Sub



End Sub



Private Sub mnuMultiplyreccol_Click()

Dim t As Integer, i As Integer, j As Integer, txt As String

On Error GoTo errortrap

t = 4

t = InputBox("Insert the data column you want to modify (H is in 1) :", prog_name, 4)

j = InputBox("Insert the number of column you want to add as 1/column (attention, data may be unusable after this operation) :", prog_name, 1)

If t < 6 And t > 0 And j < 6 And j > 0 Then

For i = 1 To grid.Rows - 1

grid.Col = j



grid.Row = i

txt = CStr(1 / Val(grid.Text))



If Val(txt) = 0 Or grid.Text = "" Then Exit For



grid.Col = t

grid.Text = Format$(Val(grid.Text) * Val(txt), "####0.0000#")

Next i

End If

Exit Sub

errortrap:

raport "Error in mnuAddr _click routine."

Exit Sub



End Sub



Private Sub mnuOverlapFullprof_Click()

On Error GoTo errortrap
Dim return_code As Boolean, t As Integer, sT(6) As Single, tlimit As Single, linie As String
Dim olddoitheta As Double, newdoitheta As Double, jcount1 As Integer, jcount2 As Integer
Dim i As Integer, cell(7) As Double, cellr(7) As Double
''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1

Dim h As Integer, k As Integer, l As Integer, lambda As Double

Screen.MousePointer = 11

'verific parametrii

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect cell parameters. Try again."

Next i

lambda = Val(txt)

If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."

'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

'trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

Call reciproc(cell, cellr, return_code)

If (return_code) Then Err.Raise 1101, , "Error in computing the reciprocal cell."

raport strLinie

raport "Direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i



raport strLinie

Screen.MousePointer = 11

raport "Trying to read and ""decimate"" a Fullprof fou file." & vbCrLf & "Please wait..."

raport "Intended output: h, k, l, Fosq, esd(Fosq) "

raport "The output format is for Shelx .hkl file"

inpfil = FreeFile

outfil = FreeFile + 1

inputfile = ""

Call open_file(inputfile, 1, return_code)

If Not (return_code) Then Err.Raise 1101, , "Input file: operation cancelled"

raport "The file is " & inputfile

Call open_file(outputfile, 2, return_code)

If Not (return_code) Then Err.Raise 1101, , "Output file: operation cancelled"

raport "The output file is " & outputfile

raport "If the file exists the data will be merged."

tlimit = InputBox("Input 2 theta limit (degrees) : ", prog_name, 0.01)

If tlimit < 0.005 Or tlimit > 1 Then Err.Raise 1101, , "Incorrect value, try again..."

t = InputBox("Batch number for Shelx", prog_name, 1)

t = CInt(Val(t))



Open inputfile For Input As inpfil

Open outputfile For Append As outfil

jcount1 = 0

jcount2 = 0

On Error GoTo eroarepast

Line Input #inpfil, linie

raport "the first line in .fou file ignored..."

Do While Not (EOF(inpfil))

Line Input #inpfil, linie

sT(1) = CInt(left$(linie, 4))

sT(2) = CInt(Mid$(linie, 5, 4))

sT(3) = CInt(Mid$(linie, 9, 4))

sT(4) = Val(Mid$(linie, 13, 8))

sT(5) = Val(Mid$(linie, 21, 8))

''st(6) = CInt(Right$(linie, 4))

If ((sT(1) = 0) And (sT(2) = 0) And (sT(3) = 0)) Then

Exit Do

Else

jcount1 = jcount1 + 1

newdoitheta = doitheta_deg(cellr, lambda, CInt(sT(1)), CInt(sT(2)), CInt(sT(3)))

If olddoitheta + tlimit < newdoitheta Then Print #outfil, Format$(Format$(CInt(sT(1)), "###0"), "@@@@") & Format$(Format$(CInt(sT(2)), "###0"), "@@@@") & Format$(Format$(CInt(sT(3)), "###0"), "@@@@") & Format$(Format$(Val(left$(sT(4), 8)), "######0."), "@@@@@@@@") & Format$(Format$(Val(left$(sT(5), 8)), "######0."), "@@@@@@@@") & Format$(Format$(t, "###0"), "@@@@"): jcount2 = jcount2 + 1: olddoitheta = newdoitheta

End If

Loop



outofhere:

On Error GoTo errortrap

Print #outfil, "   0   0   0      0.      0.   0"

Close

Err.Clear

Screen.MousePointer = 0

raport CStr(jcount1) & " lines found..."

raport CStr(jcount1 - jcount2) & " lines discarded..."

Exit Sub



eroarepast:

Err.Clear

Resume outofhere



errortrap:

Close

Screen.MousePointer = 0

raport Err.Description

Exit Sub

End Sub



Private Sub mnuOverlapGsas_Click()

On Error GoTo errortrap

'rutina trebuie sa dea cu flitu unor reflectii care sunt prea aproape in theta

'sau in d; ca overlapu lui lebail

Dim return_code As Boolean, sT(11) As Single, tlimit As Single

Dim olddoitheta As Double, newdoitheta As Double, jcount1 As Integer, jcount2 As Integer

Dim i As Integer, coderoare As Boolean, cell(7) As Double, cellr(7) As Double

''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1

Dim h As Integer, k As Integer, l As Integer, lambda As Double

Screen.MousePointer = 11

'verific parametrii

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect cell parameters. Try again."

Next i

lambda = Val(txt)

If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."

'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

'trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

Call reciproc(cell, cellr, coderoare)

If (coderoare) Then Err.Raise 1101, , "Error in computing the reciprocal cell."

raport strLinie

raport "Direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i



raport strLinie

Screen.MousePointer = 11

raport "Trying to read a GSAS Reflection file." & vbCrLf & "Please wait..."

raport "Warning: the file must be saved with the option R/ascii in Reflist program"

raport "Intended output: h, k, l, Fosq, esd(Fosq) "

raport "The output format is for Shelx .hkl file"

inpfil = FreeFile

outfil = FreeFile + 1

inputfile = ""

Call open_file(inputfile, 1, return_code)

If Not (return_code) Then Err.Raise 1101, , "Input file: operation cancelled"

raport "The file is " & inputfile

Call open_file(outputfile, 2, return_code)

If Not (return_code) Then Err.Raise 1101, , "Output file: operation cancelled"

raport "The output file is " & outputfile

raport "If the file exists the data will be merged."

tlimit = InputBox("Input 2 theta limit (degrees) : ", prog_name, 0.01)

tlimit = (Val(tlimit))

If tlimit < 0.0005 Or tlimit > 0.5 Then Err.Raise 1101, , "Incorrect value, try again..."

t = InputBox("Batch number for Shelx", prog_name, 1)

t = CInt(Val(t))

'citesc cate linii are fisierul, fac dimensionarea si verific formatul

Open inputfile For Input As inpfil

Open outputfile For Append As outfil

jcount1 = 0

jcount2 = 0

On Error GoTo eroarepast

Do While Not (EOF(inpfil))

Input #inpfil, sT(1), sT(2), sT(3)

If ((sT(1) = 0) And (sT(2) = 0) And (sT(3) = 0)) Then

Exit Do

Else

jcount1 = jcount1 + 1

Input #inpfil, sT(4), sT(5), sT(6), sT(7), sT(8), sT(9), sT(10), sT(11)

newdoitheta = doitheta_deg(cellr, lambda, CInt(sT(1)), CInt(sT(2)), CInt(sT(3)))

If olddoitheta + tlimit < newdoitheta Then Print #outfil, Format$(Format$(CInt(sT(1)), "###0"), "@@@@") & Format$(Format$(CInt(sT(2)), "###0"), "@@@@") & Format$(Format$(CInt(sT(3)), "###0"), "@@@@") & Format$(Format$(Val(left$(sT(8), 8)), "######0."), "@@@@@@@@") & Format$(Format$(Val(left$(sT(9), 8)), "######0."), "@@@@@@@@") & Format$(Format$(CInt(t), "###0"), "@@@@"): jcount2 = jcount2 + 1: olddoitheta = newdoitheta

End If

Loop

outofhere:

On Error GoTo errortrap

Print #outfil, "   0   0   0      0.      0.   0"

Close

Screen.MousePointer = 0

raport CStr(jcount1) & " lines found..."

raport CStr(jcount1 - jcount2) & " lines discarded..."

Exit Sub

eroarepast:

Err.Clear

Resume outofhere



errortrap:

Close

Screen.MousePointer = 0

raport Err.Description

Exit Sub

End Sub



Private Sub mnuOverlapShelx_Click()

On Error GoTo errortrap

'rutina trebuie sa dea cu flitu unor reflectii care sunt prea aproape in theta

'sau in d; ca overlapu lui lebail

Dim return_code As Boolean, sT(6) As Single, tlimit As Single, linie As String

Dim olddoitheta As Double, newdoitheta As Double, jcount1 As Integer, jcount2 As Integer

Dim i As Integer, cell(7) As Double, cellr(7) As Double

''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1

Dim h As Integer, k As Integer, l As Integer, lambda As Double

Screen.MousePointer = 11

'verific parametrii

For i = 0 To 5

If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect cell parameters. Try again."

Next i

lambda = Val(txt)

If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."

'citesc cell si apoi calculez cellr

    For i = 1 To 6

    cell(i) = Val(txtRefine(i - 1).Text)

    Next i

'trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))

Call reciproc(cell, cellr, return_code)

If (return_code) Then Err.Raise 1101, , "Error in computing the reciprocal cell."

raport strLinie

raport "Direct and reciprocal values of the parameters:"

    For i = 1 To 7

    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")

    Next i



raport strLinie

Screen.MousePointer = 11

raport "Trying to read and ""decimate"" a shelx hkl file." & vbCrLf & "Please wait..."

raport "Intended output: h, k, l, Fosq, esd(Fosq) "

raport "The output format is for Shelx .hkl file"

inpfil = FreeFile

outfil = FreeFile + 1

inputfile = ""

Call open_file(inputfile, 1, return_code)

If Not (return_code) Then Err.Raise 1101, , "Input file: operation cancelled"

raport "The file is " & inputfile

Call open_file(outputfile, 2, return_code)

If Not (return_code) Then Err.Raise 1101, , "Output file: operation cancelled"

raport "The output file is " & outputfile

raport "If the file exists the data will be merged."

tlimit = InputBox("Input 2 theta limit (degrees) : ", prog_name, 0.01)

If tlimit < 0.005 Or tlimit > 1 Then Err.Raise 1101, , "Incorrect value, try again..."

Open inputfile For Input As inpfil

Open outputfile For Append As outfil

jcount1 = 0

jcount2 = 0

On Error GoTo eroarepast

Do While Not (EOF(inpfil))

Line Input #inpfil, linie

sT(1) = CInt(left$(linie, 4))

sT(2) = CInt(Mid$(linie, 5, 4))

sT(3) = CInt(Mid$(linie, 9, 4))

sT(4) = Val(Mid$(linie, 13, 8))

sT(5) = Val(Mid$(linie, 21, 8))

sT(6) = CInt(right$(linie, 4))

If ((sT(1) = 0) And (sT(2) = 0) And (sT(3) = 0)) Then

Exit Do

Else

jcount1 = jcount1 + 1

newdoitheta = doitheta_deg(cellr, lambda, CInt(sT(1)), CInt(sT(2)), CInt(sT(3)))

If olddoitheta + tlimit < newdoitheta Then Print #outfil, Format$(Format$(CInt(sT(1)), "###0"), "@@@@") & Format$(Format$(CInt(sT(2)), "###0"), "@@@@") & Format$(Format$(CInt(sT(3)), "###0"), "@@@@") & Format$(Format$(Val(left$(sT(4), 8)), "######0."), "@@@@@@@@") & Format$(Format$(Val(left$(sT(5), 8)), "######0."), "@@@@@@@@") & Format$(Format$(sT(6), "###0"), "@@@@"): jcount2 = jcount2 + 1: olddoitheta = newdoitheta

End If

Loop



outofhere:

On Error GoTo errortrap

Print #outfil, "   0   0   0      0.      0.   0"

Close

Screen.MousePointer = 0

raport CStr(jcount1) & " lines found..."

raport CStr(jcount1 - jcount2) & " lines discarded..."

Exit Sub



eroarepast:

Err.Clear

Resume outofhere



errortrap:

Close

Screen.MousePointer = 0

raport Err.Description
Exit Sub
End Sub



Private Sub mnuPasteDicvol_Click()
ChDir App.Path & "\dicvol"
On Error GoTo errortrap
Dim t As Integer, test As String, outfil As Integer, i As String * 4, j As String * 4, k As String * 4, dobs As String * 9
t = MsgBox("First, you must copy the results in the Clipboard (mark text, Ctrl+C). " & vbCrLf & "You will loose any data you have now in the grid." & vbCrLf & "Proceed ?", vbYesNo + vbDefaultButton2, prog_name)
If t = vbNo Then Exit Sub
'paste all the data I have in the clipboard in a file
outfil = FreeFile
Open "_pwPaDic.txt" For Output As outfil
Print #outfil, Clipboard.GetText
Close #outfil
'clear the data grid
mnuErase_Click
grid.Row = 1
grid.Rows = 300
Open "_pwPaDic.txt" For Input As outfil
Do Until EOF(outfil)
DoEvents
grid.Col = 1

Line Input #outfil, test
If Len(test) < 1 Then Exit Do

h = left$(test, 8)
k = Mid$(test, 9, 4)
l = Mid$(test, 13, 4)
dobs = Val(Mid$(test, 21, 9))
If dobs > 0 Then
grid.Text = CStr(CInt(h))
grid.Col = 2: grid.Text = CStr(CInt(k))
grid.Col = 3: grid.Text = CStr(CInt(l))
grid.Col = 4: grid.Text = Val(dobs)
grid.Col = 5: grid.Text = "1"
grid.Row = grid.Row + 1
End If
Loop
Close #outfil
Kill "_pwPaDic.txt"
Call mnuSetDataType_Click(2)

Exit Sub
errortrap:
Err.Clear
Close
MsgBox " Error: Clipboard empty, something else ?"
Exit Sub
End Sub

Private Sub mnuPutinGrid_Click()

mnuSetDataType_Click 2

putdata 2

End Sub



Private Sub mnuQuit_Click()
    Close
    Unload Me
Exit Sub
End Sub



Private Sub mnuReadColumn_Click()
On Error GoTo handleit
Dim return_code As Boolean, linie As String, ignor As Integer
Dim a1 As Single
Dim k As Integer, j As Integer, t As Integer
k = 0
t = InputBox("Which column in grid you want to modify (H is in 1) :", prog_name, 4)
If Not (t < 6 And t > 0) Then Err.Raise 1101, , "Try again later..."
k = InputBox("Read one data, skip n... Insert n (0 if you want all data):", prog_name, 0)
If k < 0 Or k > 10 Then Err.Raise 1101, , "Try again later, incorrect value <0 to 100>..."
inpfil = FreeFile
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie
raport inputfile & " open; read one data, skip " & CStr(k)
ignor = 0
If Len(mnuIgnore.Caption) > 14 Then
ignor = 1
If Len(mnuIgnore.Caption) > 18 Then ignor = Val(right$(mnuIgnore.Caption, Len(mnuIgnore.Caption) - 17))
End If
raport CStr(ignor) & " line(s) will be ignored."
grid.Col = t
grid.Row = 1

Open inputfile For Input As inpfil
If ignor > 0 Then
For i% = 1 To ignor
Line Input #inpfil, linie
Next i%
End If


Do While Not (EOF(inpfil))
Input #inpfil, a1
j = MsgBox(CStr(a1) & vbCrLf & "This is the first data you want ?", vbYesNoCancel + vbDefaultButton2, prog_name)
If j = vbCancel Then Err.Raise 1101, , "Cancel..."
If j = vbYes Then Exit Do
Loop
DoEvents
grid.Text = CStr(a1)
grid.Row = grid.Row + 1

DoEvents

For i = 1 To k

Input #inpfil, a1

Next i

Do While Not (EOF(inpfil))

Input #inpfil, a1

grid.Text = CStr(a1)

grid.Rows = grid.Rows + 1

grid.Row = grid.Row + 1

For i = 1 To k

Input #inpfil, a1

Next i

Loop

Close

Exit Sub

handleit:

Close

Exit Sub





End Sub



Private Sub mnuSearchparam_Click()
''this routine was not performing well, I disabled it on november 17, 2000
''i will keep this in the source for a future graphical indexation
On Error GoTo handleit
Dim dist As Double, unghi As Double, D1 As Double, D2 As Double
Static ianswer As Integer
Dim i As Integer, coderoare As Boolean, tcalc() As Double, cell(7) As Double, cellr(7) As Double
''param directi si reciproci sunt cell si cellr, al 7lea e volumul, incep de la 1
Dim nrval As Integer, test As Integer, results(8) As Double, datatype As String, valoriminime() As Double
Dim pondere() As Double, intoarce(8) As Double, ind(8) As Integer, steps As Integer, refineagain As Integer, widthsearch As Double   '0,1,2 weighting scheme
Dim h() As Integer, k() As Integer, l() As Integer, teta() As Double, zero As Double, lambda As Double
Dim sumamin As Double, polcoeff(9) As Double

If Not (ianswer = vbYes) Then
ianswer = MsgBox("This routine is intended for educational purposes (manual indexations, simple search for h,k,l correspondence, etc...). It may take a lot of time especially for low symmetry. Are you sure you want to do this ?", vbYesNo + vbDefaultButton2, "Linear search")
If ianswer = vbNo Then Exit Sub
End If

Screen.MousePointer = 11
'verific parametrii

For i = 0 To 5
If (Val(txtRefine(i).Text) <= 0 Or Val(txtRefine(i).Text) > 180) Then Err.Raise 1101, , "Incorrect cell parameters. Try again."
Next i
lambda = Val(txt)
widthsearch = Val(txtDetails(0).Text)
refineagain = CInt(txtDetails(2).Text)
steps = CInt(txtDetails(1).Text)
If lambda < 0 Or lambda > 5 Then Err.Raise 1101, , "Wrong wavelength..."
If refineagain < 0 Or refineagain > 100 Then Err.Raise 1101, , "Refine value is invalid, a positive integer less than 100 is suitable."
If widthsearch < 0 Or widthsearch > 500 Then Err.Raise 1101, , "Width search domain is invalid. "
If steps < 1 Then Err.Raise 1101, , "Inconsistent value in the Steps field..."
'ultima valoare e pentru zero, 7 e pentru lambda
For i = 1 To 8
ind(i) = 0
If chkRefine(i - 1).Value Then ind(i) = 1
Next i

'aflu numarul de valori nrval - din grid si redimensionez
grid.Col = 1
nrval = grid.Rows - 1
    For i = 1 To grid.Rows - 1
    grid.Row = i
    If grid.Text = "" Then nrval = i - 1: Exit For
    Next i
If nrval < 2 Then raport "I don't have any data... I will compute only the reciprocal values ..."

ReDim h(nrval), k(nrval), l(nrval), teta(nrval), pondere(nrval)
    For i = 1 To nrval
    grid.Row = i
    grid.Col = 1: h(i) = CInt(grid.Text)
    grid.Col = 2: k(i) = CInt(grid.Text)
    grid.Col = 3: l(i) = CInt(grid.Text)
    grid.Col = 4: teta(i) = Val(grid.Text)
    grid.Col = 5:

    If grid.Text = "" Then grid.Text = "1"
    If Val(grid.Text) <= 0 Then grid.Text = "0"
    pondere(i) = Val(grid.Text)
    Next i


If ChkRef.Value Then
Call usestandard(polcoeff)
'coef polimomului, 8 este gradul maxim...
raport strLinie
raport "Correction applied to data (correction, new value):"


For i = 1 To nrval
teta(i) = teta(i) - corectie(teta(i), polcoeff)
raport Format$(Format$(Val(-corectie(teta(i), polcoeff)), "##0.000000"), "@@@@@@@@@@") & Format$(Format$(Val(teta(i)), "##0.000000"), "@@@@@@@@@@")
Next i
End If

'citesc cell si apoi calculez cellr

    For i = 1 To 6
    cell(i) = Val(txtRefine(i - 1).Text)
    Next i

'trimit in grade

cell(7) = cell(1) * cell(2) * cell(3) * Sqr(1 - Cos(cell(4) / rd) * Cos(cell(4) / rd) - Cos(cell(5) / rd) * Cos(cell(5) / rd) - Cos(cell(6) / rd) * Cos(cell(6) / rd) + 2 * Cos(cell(4) / rd) * Cos(cell(5) / rd) * Cos(cell(6) / rd))
Call reciproc(cell, cellr, coderoare)
If (coderoare) Then Err.Raise 1101, , "Error in computing the reciproc cell."
raport strLinie

raport "Direct and reciprocal values of the parameters:"
    For i = 1 To 7
    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")
    Next i


For i = 1 To 8
If chkRefine(i - 1).Value Then test = test + 1
Next i

If nrval < 2 Then Err.Raise 1101, , "I can't compute more than this..."
If test < 1 Then Err.Raise 1101, , "Nothing to do...No refinement selected..."

''Call open_file(outputfile, 2, coderoare)
''If Not (coderoare) Then outputfile = "": Screen.MousePointer = 0: Exit Sub
''Open outputfile For Output As outfil'

For i = 1 To 6
results(i) = cellr(i)
Next i
results(7) = lambda
results(8) = 0
grid.Col = 4
grid.Row = 0


Select Case grid.Text
Case "d /A"
datatype = "d"
For i = 1 To nrval
teta(i) = Atn(lambda / 2 / (teta(i)) / Sqr(-lambda / 2 / teta(i) * lambda / 2 / teta(i) + 1))
teta(i) = teta(i) * rd
Next i

Case "2 theta"
datatype = "2"
For i = 1 To nrval
teta(i) = teta(i) / 2
Next i

Case "theta"
datatype = "t"
For i = 1 To nrval
teta(i) = teta(i)
Next i

End Select

't(i) sunt teta de lucru
'results(7) este de lucru -  este lambda de lucru
''results(8) este zero de lucru

widthsearch = widthsearch / 100 ''procente
sumamin = 100000000000000# ' it was smin, changed to sumamin on nov 17, 2000
ReDim valoriminime(refineagain + 1, 9)

'in cele 8 valori pastrez cele mai bune rezultate
For ii = 1 To refineagain  ' + 1  shouldnt be refineagain +1?  modif. 17 nov 2000
DoEvents
widthsearch = widthsearch / ii
Call searchlinear(CmbCellType.ListIndex, results, ind, steps, widthsearch, nrval, teta, h, k, l, pondere, intoarce, sumamin, coderoare)
DoEvents
For i = 1 To 6
cellr(i) = results(i)
Next i

''a problem was detected here when weigth is too high
cellr(7) = cellr(1) * cellr(2) * cellr(3) * Sqr(1 - Cos(cellr(4) / rd) * Cos(cellr(4) / rd) - Cos(cellr(5) / rd) * Cos(cellr(5) / rd) - Cos(cellr(6) / rd) * Cos(cellr(6) / rd) + 2 * Cos(cellr(4) / rd) * Cos(cellr(5) / rd) * Cos(cellr(6) / rd))

raport strLinie
raport "Results obtained after " & CStr(ii) & " cycle(s) :"
Call reciproc(cellr, cell, coderoare)
    If coderoare Then raport "Error, can not compute the unit cell...": eroare = 1 / 0
    For i = 1 To 7
    raport Format$(Format$(Val(cell(i)), "##0.0000  "), "@@@@@@@@@@") & "       " & Format$(Format$(Val(cellr(i)), "##0.0000  "), "@@@@@@@@@@")
    Next i
raport "lambda = " & CStr(results(7))
raport "zero error (2theta real - 2theta measured) = " & Format$(results(8), "##0.0000")
raport "Overall squared deviation :" & CStr(sumamin)
valoriminime(ii, 1) = sumamin
raport strLinie
For i = 1 To 8
valoriminime(ii, i + 1) = results(i)
Next i
Next ii
'selectez cele mai mici reziduals
For i = 1 To refineagain - 1 'numarvalori, it was an error here, corrected nov 17
If valoriminime(i, 1) < valoriminime(i + 1, i) Then
For ii = 1 To 8
results(ii) = valoriminime(i, ii + 1)
Next ii
End If
Next i


'tiparesc rezultate finale: h,k,l, theta observat, 2theta observat, 2theta-zero, 2theta calculat, dif. between theta, d observat, dif. between d
raport "  H   K   L  2theta  zero corr. 2th.calc. diff.2th.  d calc.   d exp.    diff.d      weight"
For i = 1 To nrval
''Call calculdtheta(results(7), results(8), h(i), k(i), l(i), results(1), results(2), results(3), results(4), results(5), results(6), dist, unghi, coderoare)
''If coderoare Then
''raport "Error - can not evaluate parameters."
''Exit For
''End If
unghi = -results(8) / 2 + calcul_theta(results(7), h(i), k(i), l(i), results(1), results(2), results(3), results(4), results(5), results(6))
'unghi este theta calculat cu parametrii astia,correctat pentru zero
D1 = (results(7) / 2 / (Sin((teta(i) + results(8) / 2) / rd))) 'd1 este d introdus
D2 = (results(7) / 2 / (Sin((unghi) / rd))) 'd2 este d calculat
raport Format$(Format$(CInt(h(i)), "##0 "), "@@@@") & Format$(Format$(CInt(k(i)), "##0 "), "@@@@") & Format$(Format$(CInt(l(i)), "##0 "), "@@@@") & Format$(Format$(Val(2 * teta(i)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * teta(i) + results(8)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * unghi), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(2 * unghi - 2 * teta(i) - results(8)), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(D2), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(D1), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(D1 - D2), "##0.0000  "), "@@@@@@@@@@") & Format$(Format$(Val(pondere(i)), "####0.00000"), "@@@@@@@@@@@")
Next i
raport strLinie
raport "The step at the end of refinement was " & Format$((100 * widthsearch), "##0.00000") & " percent"
'calculez abaterea standard pentru fiecare parametru obtinu


    'calculez amediu-ai totul la patrat, suma,..etc
raport strLinie
Screen.MousePointer = 0
Exit Sub

handleit:
Screen.MousePointer = 0
If Err = 1101 Then
raport Err.Description
Exit Sub
Else
raport "Unexpected error in mnuSearchparam routine."
End If
Exit Sub
End Sub





Private Sub mnuSendHKLToStd_Click()

putdata 3

End Sub



Private Sub mnuSetDataRows_Click()

'curat gridul si numerotez din nou

Dim t As Integer, i As Integer

On Error GoTo errortrap

t = InputBox("How many data rows you want ?", prog_name, 100)

If t < 5 Then Err.Raise 1101, , "Accepted value, at least 5..."

grid.Rows = 1 + CInt(t)



For i = 1 To grid.Rows - 1

''''''Grid.CellAlignment = 3

DoEvents

grid.Row = i

grid.Col = 0

grid.Text = CStr(i)

Next i

Exit Sub

errortrap:

raport Err.Description

Exit Sub



End Sub



Private Sub mnuSetRows_Click()



End Sub



Sub mnuSetDataType_Click(Index As Integer)

On Error GoTo handleit
Dim i As Integer
For i = 0 To 3
mnuSetDataType(i).Checked = False
Next i
mnuSetDataType(Index).Checked = True
mnuChangeDataTo2theta.Enabled = True
mnuChangeDataToTheta.Enabled = True
mnuChangeDataToD.Enabled = True
mnuChangeDataToEnergy.Enabled = True
grid.Col = 4
grid.Row = 0

Select Case Index
Case 0
''''''Grid.CellAlignment = 3
grid.Text = "2 theta"
mnuChangeDataToEnergy.Enabled = False
mnuChangeDataTo2theta.Enabled = False

Case 1

'''''''Grid.CellAlignment = 3
grid.Text = "theta"
mnuChangeDataToEnergy.Enabled = False
mnuChangeDataToTheta.Enabled = False

Case 2

'''''''Grid.CellAlignment = 3

grid.Text = "d /A"

mnuChangeDataToD.Enabled = False



Case 3

'energy

'''''''''Grid.CellAlignment = 3

grid.Text = "E /keV"

mnuChangeDataToEnergy.Enabled = False

mnuChangeDataTo2theta.Enabled = False

mnuChangeDataToTheta.Enabled = False

cmbWave.ListIndex = 24
DoEvents
End Select
Exit Sub

handleit:
raport "error in mnuSetDataType routine..."
Err.Clear
Exit Sub
End Sub



Sub mnuSortData_Click()
'sorting
On Error GoTo errortrap
'count how many data
Dim nData As Integer, i As Integer, ch() As Variant, th() As Integer, tk() As Integer, tl() As Integer, tdata() As Single, tw() As Single
Me.MousePointer = 11
Call numargrid(nData)
nData = nData - 1
If nData < 2 Then Err.Raise 1101, , "Trying to order the lines: there is nothing to do..."
grid.Rows = nData + 1
If nData < 25 Then grid.Rows = 25
ReDim ch(nData), th(nData), tk(nData), tl(nData), tdata(nData), tw(nData)
''read the grid
grid.Col = 1

For i = 1 To nData
grid.Row = i

th(i) = CInt(Val(grid.Text))
grid.Col = 2
tk(i) = CInt(Val(grid.Text))
grid.Col = 3
tl(i) = CInt(Val(grid.Text))
grid.Col = 4
tdata(i) = (Val(grid.Text))
grid.Col = 5
tw(i) = Val(grid.Text)
grid.Col = 1
Next i
'sort here in increasing order...
Dim changed As Boolean 'if i change something make it true
changed = True
Do
changed = False
For i = 2 To nData
If tdata(i) < tdata(i - 1) Then
'change them
changed = True
ch(1) = th(i): ch(2) = tk(i): ch(3) = tl(i): ch(4) = tdata(i): ch(5) = tw(i)
th(i) = th(i - 1): tk(i) = tk(i - 1): tl(i) = tl(i - 1): tdata(i) = tdata(i - 1): tw(i) = tw(i - 1)
th(i - 1) = ch(1): tk(i - 1) = ch(2): tl(i - 1) = ch(3): tdata(i - 1) = ch(4): tw(i - 1) = ch(5)
End If
Next i
If Not (changed) Then Exit Do
Loop


grid.Row = 0
grid.Col = 4
Select Case left$(grid.Text, 1)
Case "2", "t"
'increasing order of theta, normal
'put them here
grid.Col = 1
For i = 1 To nData
grid.Row = i
grid.Text = CStr(th(i))
grid.Col = 2
grid.Text = CStr(tk(i))
grid.Col = 3
grid.Text = CStr(tl(i))
grid.Col = 4
grid.Text = CStr(tdata(i))
grid.Col = 5
grid.Text = CStr(tw(i))
grid.Col = 1
Next i

Case "d"
'inverse order
grid.Col = 1
For i = 1 To nData
grid.Row = nData - i + 1
grid.Text = CStr(th(i))
grid.Col = 2
grid.Text = CStr(tk(i))
grid.Col = 3
grid.Text = CStr(tl(i))
grid.Col = 4
grid.Text = CStr(tdata(i))
grid.Col = 5
grid.Text = CStr(tw(i))
grid.Col = 1
Next i

Case Else
Err.Raise 1101, , "Sorry, not available here..."
End Select
Me.MousePointer = 0
Exit Sub
errortrap:
Me.MousePointer = 0
If Err.Number = 1101 Then MsgBox Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub mnuSpaceGroup_Click()
Dim t As Integer, v As String, limit As Single, out1 As Integer, out2 As Integer, in1 As Integer
Static message As Boolean
Dim nhkls As Integer
raport "No consistency check made (between the unit cell parameters and the space group...)"
On Error GoTo errtrap
If Not (message) Then
message = True
MsgBox "This routine calculates hkl and 2 theta for a given space group (it uses an external program, HKLGEN, made by Armel Le Bail)"
End If
v = InputBox("Input the space group symbol (standard settings, put some spaces...)", prog_name & "- hklgen", "P 1")
limit = InputBox("Input the 2 theta limit ", prog_name, 50)
If limit < 0 Or limit > 160 Then Err.Raise 1101, , ""

ChDir App.Path & "\hklgen"

out1 = FreeFile
'change the data type to d
mnuSetDataType_Click (2)
raport strLinie
raport "the starting file is " & App.Path & "\_pwdHKL.in"
raport "the output file is " & App.Path & "\_pwdHKL.out"
Open "_pwdHKL.in" For Output As out1
Print #out1, sForFormat(title, "A80")
raport prog_name & " " & title
raport "wavelength " & Val(txt.Text)
Print #out1, sForFormat(Val(txt.Text), "F8.6")
Print #out1, sForFormat(v, "A20")
raport "space group " & v
Print #out1, sForFormat(txtRefine(0).Text, "F8.5") & " " & sForFormat(txtRefine(1).Text, "F8.5") & " " & sForFormat(txtRefine(2).Text, "F8.5") & " " & sForFormat(txtRefine(3).Text, "F8.4") & " " & sForFormat(txtRefine(4).Text, "f8.4") & " " & sForFormat(txtRefine(5).Text, "f8.4")
raport "unit cell: " & txtRefine(0).Text & "  " & txtRefine(1).Text & "  " & txtRefine(2).Text & "  " & txtRefine(3).Text & "  " & txtRefine(4).Text & "  " & txtRefine(5).Text
raport "2 theta limit " & Format$(limit, "##0.0##")
raport strLinie
Print #out1, sForFormat(limit, "F8.3")
Print #out1, " "
Close #out1
DoEvents
Call ShellAndLoop("hklgen.exe _pwdHKL > debug", vbMinimizedNoFocus)

DoEvents
Open "_pwdHKL.out" For Input As out1
frmRefine.grid.Rows = 1
frmRefine.grid.Rows = 50

 Line Input #out1, v
 Line Input #out1, v
 Line Input #out1, v
 Line Input #out1, v
 Line Input #out1, v
 Line Input #out1, v 'this v has the number of hkls
nhkls = CInt(right$(v, 6))
Line Input #out1, v
frmRefine.grid.Col = 4
frmRefine.grid.Rows = nhkls + 2

For i = 1 To nhkls
frmRefine.grid.Row = i
'here reads the values
Line Input #out1, v
frmRefine.grid.Col = 1
frmRefine.grid.Text = CStr(Val(left$(v, 5)))
frmRefine.grid.Col = 2
frmRefine.grid.Text = CStr(Val(Mid$(v, 6, 4)))
frmRefine.grid.Col = 3
frmRefine.grid.Text = CStr(Val(Mid$(v, 10, 4)))

frmRefine.grid.Col = 4
frmRefine.grid.Text = CStr(Val(Mid$(v, 20, 8)))

frmRefine.grid.Col = 5

frmRefine.grid.Text = "1"

Next i

Close
Exit Sub


errtrap:
Err.Clear
raport "Error trap routine: something is wrong (wrong parameters, space group ?). Check for the file _pwdHKL.out and/or <debug> in the application directory. " & Err.Description
Close
Exit Sub

End Sub

Private Sub mnuUseIto_Click()
FrmItoSetup.Show
End Sub

Private Sub mnuUseTreor_Click()
TreorSetup.Show
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)

cmbWave.ListIndex = 24

DoEvents

End Sub



Private Sub txtRefine_Change(Index As Integer)

  Select Case CmbCellType.ListIndex

  Case 0 'cubic

  txtRefine(1).Text = txtRefine(0).Text

  txtRefine(2).Text = txtRefine(0).Text

  txtRefine(4).Text = txtRefine(3).Text

  txtRefine(5).Text = txtRefine(3).Text

  Case 1 'tetragonal

  txtRefine(1).Text = txtRefine(0).Text

  txtRefine(4).Text = txtRefine(3).Text

  txtRefine(5).Text = txtRefine(3).Text

  Case 2 'ortho

  txtRefine(4).Text = txtRefine(3).Text

  txtRefine(5).Text = txtRefine(3).Text

  Case 3 'rombo

  txtRefine(1).Text = txtRefine(0).Text

  txtRefine(2).Text = txtRefine(0).Text

  txtRefine(4).Text = txtRefine(3).Text

  txtRefine(5).Text = txtRefine(3).Text

  Case 4 'hexa

  txtRefine(1).Text = txtRefine(0).Text

  End Select

End Sub


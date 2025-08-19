VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Convert3Main 
   Caption         =   "Powder4"
   ClientHeight    =   3170
   ClientLeft      =   180
   ClientTop       =   650
   ClientWidth     =   7010
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PowderMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3170
   ScaleWidth      =   7010
   Begin VB.TextBox txtraport 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   3012
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   0
      Width           =   6972
   End
   Begin VB.Frame FrameCompute 
      Caption         =   "Compute d/angle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Choose wavelength and input either d or angle."
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select radiation"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "2theta /deg."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   " d /Angst."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3960
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Begin VB.Menu mnuDbws 
            Caption         =   "&Dbws file"
         End
         Begin VB.Menu mnuGsasESD 
            Caption         =   "Gsas - CW &ESD"
         End
         Begin VB.Menu mnuGsasSTD 
            Caption         =   "&Gsas - CW STD"
         End
         Begin VB.Menu mnuOpenLHPM 
            Caption         =   "LHPM file"
         End
         Begin VB.Menu mnuPhilips 
            Caption         =   "&Philips PC-UDF"
         End
         Begin VB.Menu mnuRiet7 
            Caption         =   "&Riet7 data file"
         End
         Begin VB.Menu mnuScintag 
            Caption         =   "Scintag data"
         End
         Begin VB.Menu mnuSiemens 
            Caption         =   "&Siemens, ascii"
         End
         Begin VB.Menu mnuOpenSietronics 
            Caption         =   "Sietronics - CPI"
         End
         Begin VB.Menu mnuWppf 
            Caption         =   "&Wppf/Profit"
         End
         Begin VB.Menu m_11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLine 
            Caption         =   "&Line"
            Begin VB.Menu mnuAscii1 
               Caption         =   "the first value, as Y"
            End
            Begin VB.Menu mnuAscii2 
               Caption         =   "the first 2 values, as XY"
            End
            Begin VB.Menu mnuAscii3 
               Caption         =   "the first 3 values, as XYZ"
            End
         End
         Begin VB.Menu mnuY 
            Caption         =   "&Y - free ascii"
         End
         Begin VB.Menu mnuXY 
            Caption         =   "&X, Y - free ascii"
         End
         Begin VB.Menu mnuXYZ 
            Caption         =   "X, Y, &Z - free ascii"
         End
         Begin VB.Menu m_10 
            Caption         =   "-"
         End
         Begin VB.Menu mnumxp18 
            Caption         =   "&MXP18 (unix), binary"
         End
         Begin VB.Menu mnuMacScience 
            Caption         =   "MAC Science (Win NT), binary"
         End
         Begin VB.Menu mnuPhilipsBinary 
            Caption         =   "Philips RD/SD, binary"
         End
         Begin VB.Menu s1 
            Caption         =   "-"
         End
         Begin VB.Menu mOpenCustom 
            Caption         =   "Custom"
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save as"
         Begin VB.Menu mnuSaveDbws 
            Caption         =   "&Dbws file"
         End
         Begin VB.Menu mnuSaveGsasESD 
            Caption         =   "Gsas - CW &ESD"
         End
         Begin VB.Menu mnuSaveGsasSTD 
            Caption         =   "&Gsas - CW STD"
         End
         Begin VB.Menu mnuSaveLHPM 
            Caption         =   "LHPM file"
         End
         Begin VB.Menu mnuSavePhilips 
            Caption         =   "&Philips UDF"
         End
         Begin VB.Menu mnuSaveRiet 
            Caption         =   "&Riet7 file"
         End
         Begin VB.Menu mnuSaveScintag 
            Caption         =   "Scintag"
         End
         Begin VB.Menu mnuSaveSiemens 
            Caption         =   "&Siemens"
         End
         Begin VB.Menu mnuSaveSietronics 
            Caption         =   "Sietronics - &CPI"
         End
         Begin VB.Menu mnuSaveWppf1 
            Caption         =   "Wppf/Profit &1"
         End
         Begin VB.Menu mnuSaveWppf2 
            Caption         =   "Wppf/Profit &2"
         End
         Begin VB.Menu mnuSaveY 
            Caption         =   "&Y - ascii"
         End
         Begin VB.Menu mnuSaveXY 
            Caption         =   "&X, Y - ascii"
         End
         Begin VB.Menu mnuSaveXYZ 
            Caption         =   "X, Y, &Z - ascii"
         End
         Begin VB.Menu m11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportDplot 
            Caption         =   "a DPLOT file"
         End
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu menumergeXY 
         Caption         =   "Merge X&Y files"
      End
      Begin VB.Menu mergeXYZ 
         Caption         =   "Merge XY&Z files"
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMDItoGSAS 
         Caption         =   "batch MDI --> GSAS"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveReportPad 
         Caption         =   "Save report pad"
      End
      Begin VB.Menu ml2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShellNotepad 
         Caption         =   "Shell to &Notepad"
      End
      Begin VB.Menu ln 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHeader 
      Caption         =   "Header"
      Begin VB.Menu mnuWithout 
         Caption         =   "I &don't have a header"
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "&Ignore first: none"
      End
   End
   Begin VB.Menu mnuGraphic 
      Caption         =   "&Graphic "
   End
   Begin VB.Menu mnuRefine 
      Caption         =   "&UnitCell"
   End
   Begin VB.Menu mnuRietveld 
      Caption         =   "&Rietveld"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuHstdmpold 
         Caption         =   "Read &HSTDMP from GSAS lst file"
      End
      Begin VB.Menu mnu_reflist 
         Caption         =   "Read REFLIST from &GSAS/rfl file "
      End
      Begin VB.Menu m_9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompute 
         Caption         =   "&Compute &d/angle"
      End
      Begin VB.Menu mnuRemovechars 
         Caption         =   "Remove strange characters"
      End
      Begin VB.Menu mnuTruncate 
         Caption         =   "Truncate data"
      End
      Begin VB.Menu m8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSwap 
         Caption         =   "&Swap X, Y"
      End
      Begin VB.Menu m5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Show &report "
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuClearReport 
         Caption         =   "&Clear report"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu m9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents - pdf file"
      End
      Begin VB.Menu mnuWWW 
         Caption         =   "Author &WWW page"
      End
      Begin VB.Menu mWWWCCp14 
         Caption         =   "CCP14 homepage"
      End
   End
End
Attribute VB_Name = "Convert3Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub remove_spaces(intrare, iesire, cod_eroare)
'inlocuieste spatiile adiacente cu un singur spatiu in stringul iesire
On Error GoTo errorTRAP
Dim schimbare As Integer
iesire = intrare
Do
schimbare = 0
For i = 1 To Len(intrare)
If Mid$(intrare, i, 2) = "  " Then schimbare = 1: iesire = left$(intrare, i) + right$(intrare, Len(intrare) - i - 1): intrare = iesire
Next i
Loop Until schimbare = 0
cod_eroare = False
Exit Sub
errorTRAP:
cod_eroare = True
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

ReDim Preserve z(numarvalori)
If Not (amzdata) Then
For i = 1 To numarvalori
z(i) = 0.1
If Y(i) > 0 Then z(i) = Sqr(Y(i))
Next i
End If

returnok = True
Exit Sub
errorTRAP:
raport "An error occured in CheckData routine. Have you inserted good values ?"
returnok = False
Err.Clear
Exit Sub
End Sub

Private Sub Combo_Click()
Combo_GotFocus
DoEvents
End Sub

Private Sub Combo_GotFocus()
Dim xx As Double
On Error GoTo errorTRAP
DoEvents
Select Case Combo.ListIndex
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

End Select
DoEvents
xx = Val(txt.Text) / 2 / Val(Text1.Text)
Text2.Text = Format$(180 / 3.14157 * (2 * Atn(xx / Sqr(1 - xx * xx))), "##0.00000")
Exit Sub
errorTRAP:
Err.Clear
Exit Sub
End Sub

Private Sub Form_Load()
Dim inpfil As Integer
inpfil = FreeFile
On Error GoTo errorTRAP
raport prog_name & " loaded." & vbCrLf & version_name & vbCrLf & strLinie & vbCrLf & Now
Combo.AddItem "Cu -K alpha"
Combo.AddItem "Cu -K alpha 1"
Combo.AddItem "Cu -K alpha 2"
Combo.AddItem "Cu -K beta"
Combo.AddItem "Cr -K alpha"
Combo.AddItem "Cr -K alpha 1"
Combo.AddItem "Cr -K alpha 2"
Combo.AddItem "Cr -K beta"
Combo.AddItem "Fe -K alpha"
Combo.AddItem "Fe -K alpha 1"
Combo.AddItem "Fe -K alpha 2"
Combo.AddItem "Fe -K beta"
Combo.AddItem "Co -K alpha"
Combo.AddItem "Co -K alpha 1"
Combo.AddItem "Co -K alpha 2"
Combo.AddItem "Co -K beta"
Combo.AddItem "Mo -K alpha"
Combo.AddItem "Mo -K alpha 1"
Combo.AddItem "Mo -K alpha 2"
Combo.AddItem "Mo -K beta"
Combo.AddItem "Ag -K alpha"
Combo.AddItem "Ag -K alpha 1"
Combo.AddItem "Ag -K alpha 2"
Combo.AddItem "Ag -K beta"
Combo.AddItem "Other..."
Combo.ListIndex = 0
ignoralinii = 0
raport "If the ASCII file is from Unix/linux systems or you have received the file by e-mail as an attachment you may have an extra character for the EOL position."
raport "If you get an error when trying to read a file, check the data file in question with a standard DOS editor like PFE or Edit. You must not have empty lines in the data file. If this is the case use the function Tools/Remove strange characters in Powder. "
raport "You may try it before actually reading the data. The use of this function is harmless..."
raport strLinie
raport "This program uses dinamic memory allocations; it will first search how many data points are listed in the file.  All errors will be listed here. Most of the errors listed in this report pad are benign trapped errors (for instance reaching the EOF)."
raport "It is advisable to carefully check all the output of this program."
raport strLinie & vbCrLf & strLinie
raport "Note that the treatment of the data (smooth, peak find, etc.) might be dependent of the values of the intensity.  You may need to scale the intensity if the values are either very small or very large (see the Graphic/Edit menu)."
'Open App.Path & "\" & "_PwdInternalData.txt" For Input As inpfil
AuthorWebPage = "http://www.u-psud.fr/lpces"
Exit Sub
errorTRAP:
raport "error encountered, " & Err.Description
Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo errorTRAP
txtraport.Width = Convert3Main.Width - 100
txtraport.Height = Convert3Main.Height - 700
txtraport.left = 0
txtraport.top = 0
Exit Sub
errorTRAP:
txtraport.Width = Convert3Main.Width - Convert3Main.Width / 10
txtraport.Height = Convert3Main.Height - Convert3Main.Width / 7
txtraport.left = 0
txtraport.top = 0
Err.Clear
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim t As Integer
t = MsgBox("Are you sure you want to quit ? ", vbDefaultButton2 + vbOKCancel, prog_name)
If t = vbOK Then End
Cancel = 1
Exit Sub
End Sub


Private Sub menumergeXY_Click()
raport strLinie & vbCrLf & "The MERGE routine will mix/average up to 10 files.  You need to have the same step for all the files (and of course that you need to have them measured with the same wavelength).  The ESD is the SQR of the average ESD^2, divided by the sum of the normalization factors."
raport Now & vbCrLf & strLinie
Load mergeXY
mergeXY.Show

End Sub

Private Sub mergeXYZ_Click()
raport strLinie & vbCrLf & "The MERGE routine will mix/average up to 10 files.  You need to have the same step for all the files (and of course that you need to have them measured with the same wavelength).  The ESD is the SQR of the average ESD^2, divided by the sum of the normalization factors."
raport Now & vbCrLf & strLinie
Load merge
merge.Show

End Sub

Private Sub mnu_reflist_Click()
'salveaza un fisier cu rez de la gsas, option R in gsas REFLIST
On Error GoTo errorTRAP
raport strLinie
Screen.MousePointer = 11
raport "Trying to read a GSAS Reflection file." & vbCrLf & "Please wait..."
raport "Warning: the file must be saved with the option R/ascii in Reflist program"
raport "Intended output: h, k, l, Fosq, esd(Fosq) and a batch number."
raport "The output format is for Shelx .hkl file"
Dim return_code As Boolean, sT(11) As Single
inpfil = FreeFile
outfil = FreeFile + 1
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Screen.MousePointer = 0: Exit Sub
raport "The input filename was " & inputfile
outputfile = ""
On Error GoTo errorcancel
Convert3Main.Dialog.Filter = "show all (*.*) |*.*"
Convert3Main.Dialog.FileName = ""
Convert3Main.Dialog.DialogTitle = prog_name & " - Output file"
Convert3Main.Dialog.Action = 2
outputfile = Convert3Main.Dialog.FileName
raport "The output filename was " & outputfile & vbCrLf & " Warning: If the file exists the data will be merged."
On Error GoTo errorTRAP
t = InputBox("Batch number for Shelx", prog_name)
If Len(CStr(t)) = 0 Then raport "Reflist reading aborted at" & Now & vbCrLf & strLinie: Exit Sub
t = CInt(Val(t))
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Open inputfile For Input As inpfil
Open outputfile For Append As outfil
Do While Not (EOF(inpfil))
Input #inpfil, sT(1), sT(2), sT(3)
If ((sT(1) = 0) And (sT(2) = 0) And (sT(3) = 0)) Then
sT(8) = 0: sT(9) = 0: t = 0
Else
Input #inpfil, sT(4), sT(5), sT(6), sT(7), sT(8), sT(9), sT(10), sT(11)
End If
Print #outfil, Format$(Format$(CInt(sT(1)), "###0"), "@@@@") & Format$(Format$(CInt(sT(2)), "###0"), "@@@@") & Format$(Format$(CInt(sT(3)), "###0"), "@@@@") & Format$(Format$(Val(left$(sT(8), 8)), "######0."), "@@@@@@@@") & Format$(Format$(Val(left$(sT(9), 8)), "######0."), "@@@@@@@@") & Format$(Format$(CInt(t), "###0"), "@@@@")
Loop
Close #inpfil
Close #outfil
raport inputfile & " was succesfully converted." & vbCrLf & Now & vbCrLf & strLinie
Screen.MousePointer = 0
Exit Sub
errorcancel:
Err.Clear
Screen.MousePointer = 0
Exit Sub
errorTRAP:
raport "Here is an error trap routine (location: mnu_reflist). " & vbCrLf & "Probably finished the job..."
Close
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub mnuAbout_Click()
About.Show
End Sub

Private Sub mnuAscii1_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, linie As String, i As Long
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie
raport inputfile & " open; data read as 1 record per line, ascii "
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines."
raport CStr(ignoralinii) & " line(s) will be ignored."
numarvalori = nr_linii - ignoralinii
raport "Maximum allowed data number is " & CStr(numarvalori)
ReDim Y(numarvalori)
amydata = True
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
For i = 1 To numarvalori
Line Input #inpfil, linie
Y(i) = CDbl(Val(linie))
Next i
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
raport "An error has occured (location: mnuAscii1)."
Err.Clear
Close
Exit Sub
End Sub

Private Sub mnuAscii2_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, i As Long
Dim linie As String
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie

raport inputfile & " open; data read as 2 records per line, ascii "
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines."
raport CStr(ignoralinii) & " line(s) will be ignored."
numarvalori = nr_linii - ignoralinii
raport CStr(numarvalori) & " points allowed"
ReDim X(numarvalori), Y(numarvalori)
amfullxdata = True: amydata = True
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
For i = 1 To numarvalori
Input #inpfil, X(i), Y(i), linie
Next i
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Err.Clear
raport "An error has occured (location: mnuAscii2)."
Close
Exit Sub
End Sub

Private Sub mnuAscii3_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, i As Long
Dim linie As String
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie & vbCrLf & inputfile & " open; data read as X, Y, Z values per line, ascii "
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines." & vbCrLf & CStr(ignoralinii) & " line(s) will be ignored."
numarvalori = nr_linii - ignoralinii
raport "Maximum allowed data number is " & CStr(numarvalori)
ReDim X(numarvalori), Y(numarvalori), z(numarvalori)

Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Input #inpfil, linie
Next i
End If
For i = 1 To numarvalori
Input #inpfil, X(i), Y(i), z(i), linie
Next i
Close

startx = X(1)
endx = X(numarvalori)
stepx = CLng(CInt(100000 * X(2) - 100000 * X(1))) / 100000
amxdata = True: amfullxdata = True: amydata = True: amzdata = True


raport "Done..." & vbCrLf & Now
raport strLinie
Exit Sub
errorTRAP:
Err.Clear
raport "An error has occured (location: mnuAscii3)."
Close
Exit Sub
End Sub

Private Sub mnuClearReport_Click()
Convert3Main.txtraport.Text = "report deleted " & Now & vbCrLf
End Sub

Private Sub mnuCompute_Click()
    mnuCompute.Checked = Not (mnuCompute.Checked)
    DoEvents
    FrameCompute.Visible = mnuCompute.Checked
    If mnuReport.Checked Then mnuReport_Click
End Sub


Private Sub mnuContents_Click()
On Error GoTo errorTRAP
''cale = "C:\windows\desktop\utile\web page\"
Dim dRet As Double
ChDir App.Path
    raport "reading the PWD2HLP.html file, it should be located in :" & CStr(App.Path)
    dRet = ShellExecute(Me.hwnd, "Open", "pwd2hlp.html", "", App.Path, 1)
    If dRet < 100 Then Err.Raise 1101, , " "
    DoEvents
Exit Sub
errorTRAP:
    raport "An error has occured : no registered browser or pwd2hlp.html is missing ? " & Err.Description
    raport strLinie
Exit Sub
End Sub

Private Sub mnuDbws_Click()
    title = " " 'prog_name
    amxdata = False: amfullxdata = False: amydata = False: amzdata = False
    On Error GoTo errorTRAP
    Dim return_code As Boolean, nr_linii As Long, linie As String
    inpfil = FreeFile
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub

IamBusy True
raport strLinie & vbCrLf & inputfile & " open; this file is supposed to be a DBWS file"
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines." & vbCrLf & CStr(ignoralinii) & " line(s) will be ignored."
raport "Maximum allowed data sets : " & CStr(8 * (nr_linii - ignoralinii))
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
    For i = 1 To ignoralinii
        Line Input #inpfil, linie
    Next i
End If

If Not (mnuWithout.Checked) Then
    amxdata = True
        Line Input #inpfil, linie
        startx = Val(left$(linie, 8))
        stepx = Val(Mid$(linie, 9, 8))
        endx = Val(Mid$(linie, 17, 8))
            If Len(linie) > 25 Then title = title + right$(linie, Len(linie) - 25)
        numarvalori = 1 + CLng((endx - startx) / stepx)
    'numarvalori = CInt(nr_linii - ignoralinii - 1) * 8
        raport "Domain in header ; start, step, end :" & CStr(startx) & ", " & CStr(stepx) & ", " & CStr(endx)
        raport "The number of points is " & CStr(numarvalori)
    Else
        numarvalori = CLng(nr_linii - ignoralinii) * 8
        raport "The X range is not defined; the number of points is set to maximum possible for this file."
    End If
ReDim Y(numarvalori)
    amydata = True
    For i = 1 To CLng(0.5 + numarvalori / 8)
'citesc doar primele 64 de caractere, in vechiul fortran ultimile opt caractere
'erau de comentariu, numarul cartelei de regula (inca mai exista ...)
    Line Input #inpfil, linie
    Y((i - 1) * 8 + 1) = Val(Mid$(linie, 1, 8))
    Y((i - 1) * 8 + 2) = Val(Mid$(linie, 9, 8))
    Y((i - 1) * 8 + 3) = Val(Mid$(linie, 17, 8))
    Y((i - 1) * 8 + 4) = Val(Mid$(linie, 25, 8))
    Y((i - 1) * 8 + 5) = Val(Mid$(linie, 33, 8))
    Y((i - 1) * 8 + 6) = Val(Mid$(linie, 41, 8))
    Y((i - 1) * 8 + 7) = Val(Mid$(linie, 49, 8))
    Y((i - 1) * 8 + 8) = Val(Mid$(linie, 57, 8))
    Next i

Close
raport "Done, " & Now & vbCrLf & strLinie
IamBusy False
Exit Sub
errorTRAP:
Close
IamBusy False
raport "An error has occured (EOF related ?)..." & vbCrLf & "This error is trapped, you get this message if you have an incomplete line at the end of file (i.e. the nr. data is not a multiple of 8)"
Err.Clear
Exit Sub
End Sub


Sub numar_linii(ByVal intrare As String, nr_linii As Long)
'modif on march 14 , 2001, add the dimension for the string test, faster
Dim test As String
On Error GoTo errorTRAP
f = FreeFile
nr_linii = 0
Open intrare For Input As f
Do While Not (EOF(f))
Line Input #f, test
nr_linii = nr_linii + 1
'pot fi si linii goale aici, acesta este numarul total de linii
Loop
Close
Exit Sub
errorTRAP:
Err.Clear
Close
Exit Sub
End Sub


Private Sub mnuExportDplot_Click()
Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "DPLOT/Windows data file."
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The filename is " & outputfile
raport "Attention: the default extension for Dplot file is grf"
''Call averageint(numarvalori, Y, valmedie, valmax, returncode)
''If Not (returncode) Then raport "An error has occured; still trying..."
''raport "The average value of intensity is " & CStr(valmedie)
''raport "The maximum value of intensity is " & CStr(valmax)

Open outputfile For Output As outfil
Print #outfil, "DPLOT/W v1.2"
Print #outfil, "data"
Print #outfil, "1"
Print #outfil, numarvalori
For i = 1 To (numarvalori)
Print #outfil, Format$(Val(X(i)), "###0.000##") & "," & Format$(Val(Y(i)), "##########0.000##")
Next i
Print #outfil, " 1    0"
Print #outfil, ""
Print #outfil, ""
Print #outfil, left$(title, 35)
Print #outfil, ""
Print #outfil, "2 theta"
Print #outfil, "int. /a.u."
Print #outfil, "1"
Print #outfil, "0.7560322,0.06053269"
Print #outfil, "Grid Type"
Print #outfil, "2"
Print #outfil, "PointSizes"
Print #outfil, "7"
Print #outfil, " 10 12 12 11 11 10 10"
Print #outfil, "Stop"

''Print #outfil, " 1   0"
''Print #outfil, ""
Close
raport "DPlot data saved..." & vbCrLf & Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured, location: mnuExportDplot. " & Err.Description & vbCrLf & strLinie
Exit Sub
End Sub

Private Sub mnuGraphic_Click()
On Error GoTo errorTRAP
Dim returncode As Boolean
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Err.Raise 1102, , " Error in VerificaDate routine: calling by mnuGraphic_Click "
DoEvents
amxdata = True
Load FrmGraph
FrmGraph.Show
FrmGraph.Enabled = True
Convert3Main.Visible = False
Exit Sub
errorTRAP:
Err.Clear
raport "An error has occurred, location: mnuGraphic_click " & Err.Description
Exit Sub
End Sub

Private Sub mnuGsasESD_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = " "
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, linieout As String
Dim linie As String, nr_spatii As Integer, spatiu(80) As Integer, i As Long
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie
IamBusy True
raport inputfile & " open; this file is supposed to be a GSAS ESD file"
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines."
Open inputfile For Input As inpfil
raport CStr(ignoralinii) & " line(s) will be ignored."
raport "Maximum allowed data number is " & CStr(5 * (nr_linii - ignoralinii))

If ignoralinii > 0 Then
    For i = 1 To ignoralinii
        Line Input #inpfil, linie
    Next i
End If
    If Not (mnuWithout.Checked) Then
        Line Input #inpfil, title
        raport "The title for this line is" & vbCrLf & title
        Line Input #inpfil, linie
        raport "INSTRUMENT PARAMETER statement will be ignored, if any."
        If UCase$(left$(linie, 20)) = "INSTRUMENT PARAMETER" Then Line Input #inpfil, linie
        If (Not (left$(linie, 4)) = "BANK") Then raport "Something may be wrong, here...Where is the statement BANK ?"
nr_spatii = 1
'inlocuiesc toate spatiile adiacente din linie cu un singur spatiu
Call remove_spaces(linie, linieout, return_code)
If return_code Then raport "An error has occured in remove_spaces routine.": Exit Sub
linie = linieout
For i = 6 To 80
If InStr(i, linie, " ") = i Then
spatiu(nr_spatii) = i
nr_spatii = nr_spatii + 1
End If
Next i
    numarvalori = Val(Mid$(linie, spatiu(1), spatiu(2) - spatiu(1)))
    raport "The number of data points is " & CStr(numarvalori)
    startx = Val(Mid$(linie, spatiu(4), spatiu(5) - spatiu(4))) / 100
    stepx = Val(Mid$(linie, spatiu(5), spatiu(6) - spatiu(5))) / 100
    endx = startx + stepx * (numarvalori - 1)
    raport "Domain in header ; start, step, end :" & CStr(startx) & ", " & CStr(stepx) & ", " & CStr(endx)
    amxdata = True
Else
numarvalori = nr_linii * 5
amxdata = False
End If
ReDim Y(numarvalori), z(numarvalori)
amydata = True: amzdata = True
'vechea versiune - merge cand liniile sunt complete
'in cazul unei linii incomplete ultimele valori sunt zero /Lachlan 1februarie98
'For i = 1 To CInt(numarvalori / 5)
'Input #inpfil, Y((i - 1) * 5 + 1), z((i - 1) * 5 + 1), Y((i - 1) * 5 + 2), z((i - 1) * 5 + 2), Y((i - 1) * 5 + 3), z((i - 1) * 5 + 3), Y((i - 1) * 5 + 4), z((i - 1) * 5 + 4), Y((i - 1) * 5 + 5), z((i - 1) * 5 + 5)
'Next i
'noua versiune citeste pana la numarvalori
    For i = 1 To numarvalori
        Input #inpfil, Y(i), z(i)
    Next i

Close
IamBusy False
raport "Done, " & Now & vbCrLf & strLinie
Exit Sub

errorTRAP:
IamBusy False
raport "passing an ErrorTrap routine, location: mnuGsasESD, " & Err.Description
Close
Exit Sub
End Sub

Private Sub mnuGsasSTD_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
'nu trateaza instrument_parameter
On Error GoTo errorTRAP
raport strLinie

raport "Reading a GSAS STD type file."
Dim return_code As Boolean, linieout As String
Dim linie As String, nr_spatii As Integer, spatiu(80) As Integer, i As Long, nr_linii As Long
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
IamBusy True
raport "The input filename is " & inputfile
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines, " & CStr(ignoralinii) & " will be ignored."
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
If Not (mnuWithout.Checked) Then
   Line Input #inpfil, title
   raport "The title in this file is " & title
    Line Input #inpfil, linie
    If UCase$(left$(linie, 20)) = "INSTRUMENT PARAMETER" Then raport "Instrument parameter not used.": Line Input #inpfil, linie
        If Not (UCase$(left$(linie, 4)) = "BANK") Then raport "Something may be wrong..."
Call remove_spaces(linie, linieout, return_code)
If return_code Then Err.Raise 1101, , "An error has occured in remove_spaces routine."
linie = linieout

nr_spatii = 1
For i = 6 To 80
If InStr(i, linie, " ") = i Then
spatiu(nr_spatii) = i
nr_spatii = nr_spatii + 1
End If
Next i
'    nr_bank = Val(Mid$(linie, 5, spatiu(1) - 5))
    numarvalori = Val(Mid$(linie, spatiu(1), spatiu(2) - spatiu(1)))
 ''   nr_linii_gsas = Val(Mid$(linie, spatiu(2), spatiu(3) - spatiu(2)))
     startx = Val(Mid$(linie, spatiu(4), spatiu(5) - spatiu(4))) / 100
    stepx = Val(Mid$(linie, spatiu(5), spatiu(6) - spatiu(5))) / 100
    endx = startx + stepx * (numarvalori - 1)
Else
numarvalori = nr_linii * 10
End If
ReDim Y(numarvalori)
amydata = True: amxdata = True
For i = 1 To CLng(numarvalori / 10) + 1
Line Input #inpfil, linie_caracter
count1 = Val(Mid$(linie_caracter, 1, 2))
If count1 <= 0 Then count1 = 1
Y(10 * (i - 1) + 1) = Val(Mid$(linie_caracter, 3, 6)) / count1
count2 = Val(Mid$(linie_caracter, 9, 2))
If count2 <= 0 Then count2 = 1
Y(10 * (i - 1) + 2) = Val(Mid$(linie_caracter, 11, 6)) / count2
count3 = Val(Mid$(linie_caracter, 17, 2))
If count3 <= 0 Then count3 = 1
Y(10 * (i - 1) + 3) = Val(Mid$(linie_caracter, 19, 6)) / count3
count4 = Val(Mid$(linie_caracter, 25, 2))
If count4 <= 0 Then count4 = 1
Y(10 * (i - 1) + 4) = Val(Mid$(linie_caracter, 27, 6)) / count4
count5 = Val(Mid$(linie_caracter, 33, 2))
If count5 <= 0 Then count5 = 1
Y(10 * (i - 1) + 5) = Val(Mid$(linie_caracter, 35, 6)) / count5
count6 = Val(Mid$(linie_caracter, 41, 2))
If count6 <= 0 Then count6 = 1
Y(10 * (i - 1) + 6) = Val(Mid$(linie_caracter, 43, 6)) / count6
count7 = Val(Mid$(linie_caracter, 49, 2))
If count7 <= 0 Then count7 = 1
Y(10 * (i - 1) + 7) = Val(Mid$(linie_caracter, 51, 6)) / count7
count8 = Val(Mid$(linie_caracter, 57, 2))
If count8 <= 0 Then count8 = 1
Y(10 * (i - 1) + 8) = Val(Mid$(linie_caracter, 59, 6)) / count8
count9 = Val(Mid$(linie_caracter, 65, 2))
If count9 <= 0 Then count9 = 1
Y(10 * (i - 1) + 9) = Val(Mid$(linie_caracter, 67, 6)) / count9
count10 = Val(Mid$(linie_caracter, 73, 2))
If count10 <= 0 Then count10 = 1
Y(10 * (i - 1) + 10) = Val(Mid$(linie_caracter, 75, 6)) / count10
Next i
Close
IamBusy False
Exit Sub
errorTRAP:
IamBusy False
raport "Something is wrong here...This error may appear in case of an incomplete line. "
raport "location: mnuGsasSTD, " & Err.Description
Close
Exit Sub
End Sub

Private Sub mnuHstdmpold_Click()
Screen.MousePointer = 11
'salveaza un fisier cu rez de la hstdmp
On Error GoTo errorTRAP
raport strLinie

raport "Trying to read a GSAS lst file." & vbCrLf & "Please wait..."
Dim return_code As Boolean, linie As String, numarvalori As Long, i As Long, hstdetect As Boolean
inpfil = FreeFile
outfil = FreeFile + 1
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Screen.MousePointer = 0: Exit Sub
raport "The file is " & inputfile
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Open inputfile For Input As inpfil
Open "_hstdmp.dat" For Output As outfil
hstdetect = False
raport "Searching the lst file for ""Program HSTDMP"" statement. Please wait..." & vbCrLf & Now
Do While Not (EOF(inpfil))
Line Input #inpfil, linie
    If InStr(linie, "Program HSTDMP") Then
    For i = 1 To 10: Line Input #inpfil, linie: Next i 'skip vreo zece linii
    'ies de aici cand gasesc primul Allen C. Larson ori o eroare
    Print #outfil, "Program HSTDMP detected in the file: " & inputfile: hstdetect = True
   raport "Program HSTDMP detected in the file: " & inputfile
   raport "Data saved in the file _hstdmp.dat as: " & vbCrLf & "Theta  I(obs)    I(cal)   Incdnt      BkGnd    Weight    Chwdt  Rf1  Rf2   "
    Print #outfil, "Theta  I(obs)    I(cal)   Incdnt      BkGnd    Weight    Chwdt  Rf1  Rf2   "
        Do
        Line Input #inpfil, linie
        If (InStr(linie, " RecNo Code Theta") Or (InStr(linie, " HSTDMP ")) Or (left$(linie, 7) = "       ") Or linie = "" Or linie = " ") Then
        Else
        Print #outfil, right$(linie, Len(linie) - 9)
        numarlinii = numarlinii + 1
        End If
        Loop
        
    End If
Loop
Close
If hstdetect Then
raport inputfile & " was succesfully converted to _hstdmp.dat" & vbCrLf & Now & strLinie
Else
raport "HSTDMP data not found."
MsgBox "Can not find HSTDMP data. Check the file"
End If
Screen.MousePointer = 0
Close
Exit Sub
errorTRAP:
Close
Screen.MousePointer = 0
If hstdetect Then
raport inputfile & " was succesfully converted to _hstdmp.dat"
Else
MsgBox "Can not find HSTDMP data", vbOKOnly, prog_name
raport "Can not find HSTDMP data, this is an error trapping routine. Check the file" & vbCrLf & strLinie
End If
Err.Clear
Exit Sub
End Sub

Private Sub mnuIgnore_Click()
Dim t As Integer
On Error GoTo errorTRAP
t = 0
t = InputBox("How many lines want to ignore (less than 32700) ?", prog_name, 1)
If Len(CStr(t)) = 0 Then Exit Sub
Select Case CInt(t)
Case 0
mnuIgnore.Caption = "&Ignore : none"
raport "All file to read."
ignoralinii = 0
Case 1
mnuIgnore.Caption = "&Ignore the first line"
raport "The first line will be ignored."
ignoralinii = 1
Case Else
mnuIgnore.Caption = "&Ignore the first " & CInt(t) & " lines"
ignoralinii = CInt(t)
raport "The first " & CInt(t) & " lines will be ignored."
End Select
Exit Sub
errorTRAP:
ignoralinii = 0
mnuIgnore.Caption = "&Ignore first: none"
Err.Clear
Exit Sub
End Sub

Private Sub mnuMacScience_Click()
On Error GoTo errorTRAP
inpfil = FreeFile
title = ""
raport strLinie
raport "MAC Science - binary file - Windows NT file system."

startx = 0: stepx = 0: endx = 0
amxdata = False: amfullxdata = False: amydata = False: amzdata = False
Dim return_code As Boolean, i As Long, linie As String * 1024
Dim pos As Integer, newlinie As String
Dim carrier As String, val1 As Byte, val2 As Byte, val3 As Byte, val4 As Byte
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport inputfile & " open; this file is supposed to be a binary file"
raport "Read header file for some info..."
Open inputfile For Binary Access Read As inpfil
Get #inpfil, 1, linie
'curat linie de mizeriile de caractere
newlinie = ""
For i = 1 To Len(linie)
carrier = Mid$(linie, i, 1)
If ((Asc(carrier) > 126) Or (Asc(carrier) < 32)) Then carrier = " "
newlinie = newlinie + carrier
Next i
linie = newlinie
pos = InStr(LCase$(linie), "start") + 6
startx = Val(right$(linie, Len(linie) - pos)) / 10000
pos = InStr(LCase$(linie), "stopa") + 6
endx = Val(right$(linie, (Len(linie) - pos))) / 10000
pos = InStr(LCase$(linie), "stepw") + 6
stepx = Val(right$(linie, Len(linie) - pos)) / 10000
raport "Please wait, it may be slow..."
If (Not (startx = stepx) And (stepx > 0)) Then amxdata = True
amfullxdata = True
numarvalori = CLng((endx - startx) / stepx)
raport "Number of points: " & CStr(numarvalori)
ReDim X(numarvalori), Y(numarvalori)
Get #inpfil, 1024, val1
amydata = True
For i = 1 To CLng((endx - startx) / stepx)
Get #inpfil, , val1
Get #inpfil, , val2
Get #inpfil, , val3
Get #inpfil, , val4
vals1 = right$("00" & CStr(Hex(val1)), 2)
vals2 = right$("00" & CStr(Hex(val2)), 2)
vals3 = right$("00" & CStr(Hex(val3)), 2)
vals4 = right$("00" & CStr(Hex(val4)), 2)
X(i) = Val("&h" & vals4 & vals3 & vals2 & vals1 & "&")
X(i) = X(i) / 10000
Get #inpfil, , val1
Get #inpfil, , val2
Get #inpfil, , val3
Get #inpfil, , val4
vals1 = right$("00" & CStr(Hex(val1)), 2)
vals2 = right$("00" & CStr(Hex(val2)), 2)
vals3 = right$("00" & CStr(Hex(val3)), 2)
vals4 = right$("00" & CStr(Hex(val4)), 2)
Y(i) = Val("&h" & vals4 & vals3 & vals2 & vals1 & "&")
Next i
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Close
raport "An error has occured..."
Err.Clear
Exit Sub

End Sub

Private Sub mnuMDItoGSAS_Click()


amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, dlinie As Double, i As Long, inpfil As Integer, outfil As Integer
Dim linie As String, banknumber As Integer, MDIWave As String * 11, MDItime As Single, MDIspace As Integer
inpfil = FreeFile
DoEvents
outfil = FreeFile + 1
inputfile = ""
raport strLinie & vbCrLf & "MDI to GSAS conversion, multiple data sets..."
raport "Warning, only the last data set will be kept in memory...."
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The input filename is " & inputfile
Call open_file(outputfile, 2, return_code)
If Not (return_code) Then Exit Sub
raport "The output file is " & outputfile


Open inputfile For Input As inpfil



numarvalori = 0
'raport "no lines will be ignored..."
'pass here when start to read MDI file
Line Input #inpfil, linie
title = left$(linie, 72)
'Line Input #inpfil, linie 'here is another line, consider it useless
Input #inpfil, startx, stepx, linie
'MDItime, MDIWave, endx, numarvalori 'here is the detail for the experiment
'line=mditime " "
MDIspace = 0
For j = 1 To Len(linie)
If Mid$(linie, j, 1) = " " Then MDIspace = MDIspace + 1
If MDIspace = 4 Then numarvalori = CLng(CStr(Val(Mid$(linie, j + 1)))): Exit For
Next j

'''numarvalori = CLng(right$(linie, 5))
'here open the file for GSAS output
Open outputfile For Output As outfil

'this is the first data set

banknumber = 1

ReDim Y(numarvalori)
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i

'here write them
completare = ""
    For i = 1 To (80 - Len(title))
    completare = completare + " "
    Next i
    Print #outfil, title + completare
    linie_de_80 = "BANK " & CStr(banknumber) & " " & 10 * CLng(numarvalori / 10) & " " & CLng(numarvalori / 10) & " CONST " & (startx * 100) & " " & (stepx * 100) & " 0 0 STD"
    completare = ""
    For i = 1 To (80 - Len(linie_de_80))
    completare = completare + " "
    Next i
    linie_de_80 = linie_de_80 + completare
    Print #outfil, linie_de_80
For i = 1 To Fix(numarvalori / 10) '+ 1
Print #outfil, "  " + Format$(Format$(Val(Y((i - 1) * 10 + 1)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 2)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 3)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 4)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 5)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 6)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 7)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 8)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 9)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 10)), "#####0"), "@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 10 * Fix(numarvalori / 10) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 10 * Fix(numarvalori / 10) + 1 To numarvalori
ultimalinie = ultimalinie + "  " + Format$(Format$(Val(Y(i)), "#####0"), "@@@@@@")
Next i

    completare = ""
    For i = 1 To (80 - Len(ultimalinie))
    completare = completare + " "
    Next i
    ultimalinie = ultimalinie + completare
Print #outfil, ultimalinie
End If


'other data sets, if there is anything else left

Do While (Not (EOF(inpfil)))
banknumber = banknumber + 1
'read again the startx, etc
Input #inpfil, startx, stepx, linie

MDIspace = 0
For j = 1 To Len(linie)
If Mid$(linie, j, 1) = " " Then MDIspace = MDIspace + 1
If MDIspace = 4 Then numarvalori = CLng(CStr(Val(Mid$(linie, j + 1)))): Exit For
Next j

ReDim Y(numarvalori)
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i

'here write them
    linie_de_80 = "BANK " & CStr(banknumber) & " " & 10 * CLng(numarvalori / 10) & " " & CLng(numarvalori / 10) & " CONST " & (startx * 100) & " " & (stepx * 100) & " 0 0 STD"
    completare = ""
    For i = 1 To (80 - Len(linie_de_80))
    completare = completare + " "
    Next i
    linie_de_80 = linie_de_80 + completare
    Print #outfil, linie_de_80
For i = 1 To Fix(numarvalori / 10) '+ 1
Print #outfil, "  " + Format$(Format$(Val(Y((i - 1) * 10 + 1)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 2)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 3)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 4)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 5)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 6)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 7)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 8)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 9)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 10)), "#####0"), "@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 10 * Fix(numarvalori / 10) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 10 * Fix(numarvalori / 10) + 1 To numarvalori
ultimalinie = ultimalinie + "  " + Format$(Format$(Val(Y(i)), "#####0"), "@@@@@@")
Next i

    completare = ""
    For i = 1 To (80 - Len(ultimalinie))
    completare = completare + " "
    Next i
    ultimalinie = ultimalinie + completare
Print #outfil, ultimalinie
End If



Loop

Close
'raport "The number of points is: " & CStr(numarvalori)
amydata = True
amxdata = True
'ReDim y(numarvalori)
'Open inputfile For Input As inpfil
'If ignoralinii > 0 Then
'For i = 1 To ignoralinii
'Line Input #inpfil, linie
'Next i
'End If
'If Not (mnuWithout.Checked) Then raport "There are no header for this file type..."
'For i = 1 To numarvalori
'Input #inpfil, y(i)
'Next i
Close
raport "done --> " & Now

Exit Sub
errorTRAP:
raport "An error has occured..--> " & Err.Description
Err.Clear
Close

Exit Sub
















End Sub

Private Sub mnumxp18_Click()
On Error GoTo errorTRAP
inpfil = FreeFile
title = ""
raport strLinie & "MAC Science - MXP 18, rotating anode instrument - UNIX "
startx = 0: stepx = 0: endx = 0
amxdata = False: amfullxdata = False: amydata = False: amzdata = False
Dim return_code As Boolean, i As Long, linie As String * 1024
Dim pos As Integer, newlinie As String
Dim carrier As String, val1 As Byte, val2 As Byte, val3 As Byte, val4 As Byte
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport inputfile & " open; this file is supposed to be a - MXP unix, binary file"
raport "Read header file for some info..."
Open inputfile For Binary Access Read As inpfil
Get #inpfil, 1, linie
'curat linie de mizeriile de caractere
newlinie = ""
For i = 1 To Len(linie)
carrier = Mid$(linie, i, 1)
If ((Asc(carrier) > 126) Or (Asc(carrier) < 32)) Then carrier = " "
newlinie = newlinie + carrier
Next i
linie = newlinie
pos = InStr(LCase$(linie), "start") + 6
startx = Val(right$(linie, Len(linie) - pos)) / 10000
pos = InStr(LCase$(linie), "stopa") + 6
endx = Val(right$(linie, (Len(linie) - pos))) / 10000
pos = InStr(LCase$(linie), "stepw") + 6
stepx = Val(right$(linie, Len(linie) - pos)) / 10000
raport "Please wait, it may be slow..."
If (Not (startx = stepx) And (stepx > 0)) Then amxdata = True
amfullxdata = True
numarvalori = CLng((endx - startx) / stepx)
raport "Number of points: " & CStr(numarvalori)
ReDim X(numarvalori), Y(numarvalori)
Get #inpfil, 1024, val1
amydata = True
For i = 1 To CLng((endx - startx) / stepx)
Get #inpfil, , val1
Get #inpfil, , val2
Get #inpfil, , val3
Get #inpfil, , val4
vals1 = right$("00" & CStr(Hex(val1)), 2)
vals2 = right$("00" & CStr(Hex(val2)), 2)
vals3 = right$("00" & CStr(Hex(val3)), 2)
vals4 = right$("00" & CStr(Hex(val4)), 2)
X(i) = Val("&h" & vals1 & vals2 & vals3 & vals4 & "&")
X(i) = X(i) / 10000
Get #inpfil, , val1
Get #inpfil, , val2
Get #inpfil, , val3
Get #inpfil, , val4
vals1 = right$("00" & CStr(Hex(val1)), 2)
vals2 = right$("00" & CStr(Hex(val2)), 2)
vals3 = right$("00" & CStr(Hex(val3)), 2)
vals4 = right$("00" & CStr(Hex(val4)), 2)
Y(i) = Val("&h" & vals1 & vals2 & vals3 & vals4 & "&")
Next i
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Close
raport "An error has occured, location: mnuMxp18, " & Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub mnuOpenLHPM_Click()
title = ""
amxdata = False: amfullxdata = False: amydata = False: amzdata = False
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, linie As String
inpfil = FreeFile
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport strLinie
raport inputfile & " open; this file is supposed to be a LHPM file, " & "ten values on each line are to be read."
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines."
raport CStr(ignoralinii) & " line(s) will be ignored."
raport "Maximum allowed data number is " & CStr(10 * (nr_linii - ignoralinii))
Open inputfile For Input As inpfil
Do While Not (EOF(inpfil))
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
If Not (mnuWithout.Checked) Then
amxdata = True
    Line Input #inpfil, linie
    startx = Val(left$(linie, 8))
    stepx = Val(Mid$(linie, 9, 8))
    endx = Val(Mid$(linie, 17, 8))
    title = " no title "
    If Len(linie) > 25 Then title = right$(linie, Len(linie) - 25)
    numarvalori = 1 + CLng((endx - startx) / stepx)
    'numarvalori = CInt(nr_linii - ignoralinii - 1) * 8
    raport "Domain in header ; start, step, end :" & CStr(startx) & ", " & CStr(stepx) & ", " & CStr(endx)
    raport "The number of points is " & CStr(numarvalori)
    Else
    numarvalori = CLng(nr_linii - ignoralinii) * 10
    raport "The X range is not defined; the number of points is set to maximum possible for this file."
End If
ReDim Y(numarvalori)
amydata = True
For i = 1 To CLng(0.5 + numarvalori / 10)
'citesc doar primele 64 de caractere, in vechiul fortran ultimile opt caractere
'erau de comentariu, numarul cartelei de regula (inca mai exista ...)
Line Input #inpfil, linie
Y((i - 1) * 10 + 1) = Val(Mid$(linie, 1, 8))
Y((i - 1) * 10 + 2) = Val(Mid$(linie, 9, 8))
Y((i - 1) * 10 + 3) = Val(Mid$(linie, 17, 8))
Y((i - 1) * 10 + 4) = Val(Mid$(linie, 25, 8))
Y((i - 1) * 10 + 5) = Val(Mid$(linie, 33, 8))
Y((i - 1) * 10 + 6) = Val(Mid$(linie, 41, 8))
Y((i - 1) * 10 + 7) = Val(Mid$(linie, 49, 8))
Y((i - 1) * 10 + 8) = Val(Mid$(linie, 57, 8))
Y((i - 1) * 10 + 9) = Val(Mid$(linie, 65, 8))
Y((i - 1) * 10 + 10) = Val(Mid$(linie, 74, 8))

Next i
Loop
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Close
raport "An error has occured..." & vbCrLf & "This error is trapped, you get this message if you have an incomplete line at the end of file (i.e. the nr. data is not a multiple of 10)"
Err.Clear
Exit Sub
End Sub

Private Sub mnuOpenSietronics_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean
Dim linie As String, i As Long, nr_linii As Long, t As String, tt As String
inpfil = FreeFile
inputfile = ""
raport strLinie
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Call numar_linii(inputfile, nr_linii)
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
If Not (mnuWithout.Checked) Then
Line Input #inpfil, title
Line Input #inpfil, t
startx = Val(t)
Line Input #inpfil, t
endx = Val(t)
Line Input #inpfil, t
stepx = Val(t)
i = 1
Do
i = i + 1
Line Input #inpfil, t
If i > 50 Then Err.Raise 1101, , "Could not find the SCANDATA statement...Aborting.." & vbCrLf & " You can read a CPI file ,ignoring the header, as free Y file."
Loop Until UCase$(left$(t, 8)) = "SCANDATA"

numarvalori = 1 + ((endx - startx) / stepx)
End If ' este sfarsitul de la header_from
'de aici incepe partea de citire standard
ReDim Y(numarvalori)
amydata = True
amxdata = True
'noua versiune - citeste si liniile incomplete-poate adauga niste zerouri atunci cand ceva e gresit...
raport CStr(numarvalori) & " datapoints to read."
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i
raport "Done..."
Exit Sub
errorTRAP:
If Err.Number = 1101 Then MsgBox Err.Description
raport "An error has occured. This might be due to unrecognized characters in the data file.  Try Tools/Remove strange characters and then try again... "
Err.Clear
Close
Exit Sub
End Sub

Private Sub mnuPhilips_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean
Dim linie As String, i As Long, nr_linii As Long, t As String, tt As String
inpfil = FreeFile
inputfile = ""
raport strLinie

Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Call numar_linii(inputfile, nr_linii)
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If

If Not (mnuWithout.Checked) Then
Do
Line Input #inpfil, tt
cod_eroare = cod_eroare + 1
t = LCase$(left$(tt, 6))
Select Case t
Case "sample"
title = (right$(tt, Len(tt) - 12))
If Len(title) < 3 Then title = inputfile
'Case "title1"
'titlu = titlu + (Right$(tt, Len(tt) - 6))
''Case "anode,"
'citesc doar numele anodului, daca nu-l cunosc il iau drept cupru
''wlnume = Left$(Right$(tt, 4), 2)
    Case "dataan"
    DataAn = (right$(tt, Len(tt) - 15)) ' le separ apoi in startx si endx
    startx = Val(left$(DataAn, 10))
    endx = Val(right$(DataAn, 10))
    Case "scanst"
    If LCase$(left$(tt, 12)) = "scanstepsize" Then stepx = Val(right$(tt, Len(tt) - 13))
End Select
If cod_eroare = 50 Then MsgBox "Fatal error. I can not find the RAW DATA statement. ": Close: Exit Sub
linii_titlu = linii_titlu + 1
Loop Until t = "rawsca"
'cod_eroare este nedefinit
numarvalori = (nr_linii - linii_titlu - ignoralinii - 1) * 8 'are o linie goala la sfarsit
amxdata = True
Else
numarvalori = (nr_linii - ignoralinii) * 8
End If ' este sfarsitul de la header_from
'de aici incepe partea de citire standard
ReDim Y(numarvalori)
amydata = True
'-------------------------------------------
'SampleIdent,2223_iod_step scan,/
'Title1,Chimie des Solides Orsay,/
'Title2,PC-APD, Diffraction software,/
'DiffrType,PW3710,/
'DiffrNumber,1,/
'Anode,Cu,/
'LabdaAlpha1, 1.54060,/ 'nu le folosesc decat la transferul in wppf/profit
'LabdaAlpha2, 1.54439,/
'RatioAlpha21, 0.50000,/
'DivergenceSlit,Fixed,1,/
'ReceivingSlit,0.1,/
'MonochromatorUsed,NO  ,/    'defaulturi ca la Orsay
'GeneratorVoltage,  40,/
'TubeCurrent,  20,/
'FileDateTime,19-feb-1995 11:48,/
'DataAngleRange,  15.0000,  77.2100,/
'ScanStepSize,   0.010,/
'ScanType,STEP,/
'ScanStepTime,    5.00,/
'RawScan
'vechea versiune: vezi nota de la gsas esd
''For i = 1 To CInt(numarvalori / 8)
''Input #inpfil, Y((i - 1) * 8 + 1), Y((i - 1) * 8 + 2), Y((i - 1) * 8 + 3), Y((i - 1) * 8 + 4), Y((i - 1) * 8 + 5), Y((i - 1) * 8 + 6), Y((i - 1) * 8 + 7), Y((i - 1) * 8 + 8)
''Next i
'noua versiune - citeste si liniile incomplete-poate adauga niste zerouri atunci cand ceva e gresit...
raport CStr(numarvalori) & " datapoints to read."
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i
raport "Done..."
Exit Sub
errorTRAP:
Err.Clear
Close
Exit Sub
End Sub

Private Sub mnuPhilipsBinary_Click()
On Error GoTo errorTRAP
inpfil = FreeFile
title = ""
raport strLinie
raport "binary file, RD/SD Philips APD "
raport "please be patient; it may take some time..."
startx = 0: stepx = 0: endx = 0
amxdata = False: amfullxdata = False: amydata = False: amzdata = False
Dim return_code As Boolean, filetype As String, i As Single, j As Single
Dim val1 As String * 2, val2 As Byte, vl2 As Byte, jcount As Integer, val3 As Double
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
IamBusy True
raport inputfile & " open; this file is supposed to be a Philips RD or SD binary file"
raport "Warning: only the data will be read (kV, mA, settings, ..ignored)."
Close
Open inputfile For Binary Access Read As inpfil
Get #inpfil, 1, val1
filetype = "unknown "
If UCase$(CStr(val1)) = "V3" Then filetype = "RD file"
If UCase$(CStr(val1)) = "V5" Then filetype = "SD file"
Get #inpfil, 215, val3
stepx = Val(val3)
Get #inpfil, 223, val3
startx = Val(val3)
Get #inpfil, 231, val3
endx = Val(val3)
numarvalori = (endx - startx) / stepx
ReDim Y(numarvalori)
raport "File type: " & filetype
If filetype = "SD file" Then jcount = 560 'adjust here the shift
j = 0
Dim vals1 As String, xxx As Single
For i = 251 + jcount To 251 + jcount + 2 * numarvalori - 1 Step 2
j = j + 1
Get #inpfil, i, val2
Get #inpfil, i + 1, vl2
vals1 = right$("00" & CStr(Hex(vl2)), 2) & right$("00" & CStr(Hex(val2)), 2)
''xxx = CSng("&h" & CStr(vl2) & right$("00" & CStr(Hex(vals1)), 2) & "&")
Y(j) = Val((0.01 * (Val("&H" & vals1)) ^ 2))
Next i
Close
raport CStr(numarvalori) & " points read.."
raport "Done..." & Now
amxdata = True
amydata = True
IamBusy False
Exit Sub
errorTRAP:
IamBusy False
Close
raport "An error has occured: " & Err.Description
Err.Clear
Exit Sub
End Sub



Private Sub mnuQuit_Click()
Dim t As Integer
t = MsgBox("Are you sure you want to quit ? ", vbDefaultButton2 + vbOKCancel, prog_name)
If t = vbOK Then End
Cancel = 1
Exit Sub
End Sub

Private Sub mnuRefine_Click()
raport "Refine cell routine..."
frmRefine.Show
End Sub

Private Sub mnuRemovechars_Click()
title = ""
Screen.MousePointer = 11
raport "Trying to remove strange characters..(other than standard DOS eol)."
raport "This is useful only for some Unix files."
raport "This may be long. Please wait."
On Error GoTo errorTRAP
Dim return_code As Boolean, linie As String, newlinie As String, carrier As String * 1, nr_linii As Long
inpfil = FreeFile
outfil = FreeFile + 1
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Screen.MousePointer = 0: Exit Sub
raport inputfile & " open; this file will be preserved as it is. Choose the outputfile."
Call open_file(outputfile, 2, return_code)
If Not (return_code) Then outputfile = "": Screen.MousePointer = 0: Exit Sub
raport outputfile & " open."
Open inputfile For Input As inpfil
Open outputfile For Output As outfil
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
Close
Screen.MousePointer = 0
Call numar_linii(outputfile, nr_linii)
raport "This new file has " & CStr(nr_linii) & " line(s)."
raport "All characters within the extended ascii code (and those with code smaller than 32) were replaced by space."
Exit Sub
Exit Sub
errorTRAP:
Close
Screen.MousePointer = 0
raport "An error has occured."

Exit Sub
End Sub

Private Sub mnuReport_Click()
mnuReport.Checked = Not (mnuReport.Checked)
DoEvents
txtraport.Visible = mnuReport.Checked
End Sub

Private Sub mnuRiet7_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean
Dim linie As String, i As Long, nr_linii As Long, t As String, tt As String
inpfil = FreeFile
inputfile = ""
raport strLinie

Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Call numar_linii(inputfile, nr_linii)
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If


If (mnuWithout.Checked) Then Err.Raise 1101, , "If you ignore the header, read a Riet7 file as Ascii Y file"
Line Input #inpfil, title
Line Input #inpfil, linie
Line Input #inpfil, linie
Line Input #inpfil, linie
raport "Details ignored (such as monochromator, date, etc,..."
    Line Input #inpfil, linie
    startx = Val(left$(linie, 8))
    stepx = Val(Mid$(linie, 9, 8))
    endx = Val(Mid$(linie, 17, 8))
numarvalori = 1 + ((endx - startx) / stepx)
'de aici incepe partea de citire standard
ReDim Y(numarvalori)
amydata = True
amxdata = True
raport CStr(numarvalori) & " datapoints to read."
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i
raport "Done..."
Exit Sub
errorTRAP:
MsgBox Err.Description
raport "Please check that you don't have extra EOL characters. Look at the file with PFE or a standard DOS editor"
Err.Clear
Close
Exit Sub

End Sub

Private Sub mnuRietveld_Click()
Load frmgDBWSmain
frmgDBWSmain.Show
End Sub

Private Sub mnuSaveDbws_Click()
Dim returncode As Boolean, i As Long, ultimalinie As String
On Error GoTo errorTRAP
raport strLinie
raport "Output in DBWS format."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The file is " & outputfile
Open outputfile For Output As outfil
Print #outfil, Format$(Format$(startx, "#0.000##"), "@@@@@@@@"); Format$(Format$(stepx, "#0.000##"), "@@@@@@@@"); Format$(Format$(endx, "#0.000##"), "@@@@@@@@"), left$(title, 30)
For i = 1 To Fix(numarvalori / 8)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 8 + 1)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 2)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 3)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 4)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 5)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 6)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 7)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 8)), "#####0. "), "@@@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 8 * Fix(numarvalori / 8) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 8 * Fix(numarvalori / 8) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "#####0. "), "@@@@@@@@")
Next i
Print #outfil, ultimalinie
End If
Close
raport "DBWS file written, it seems to be ok...Check and adjust the output data to your needs."
raport Now
Exit Sub
errorTRAP:
Err.Clear
raport "An error has occured."
Close
Exit Sub
End Sub

Private Sub mnuSaveGsasESD_Click()
Dim returncode As Boolean, i As Long, completare As String, linie_de_80 As String, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "Trying to save GSAS ESD file. Checking data."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The requested file is " & outputfile
'clean stepx
stepx = (CLng(10000000 * stepx)) / 10000000
Open outputfile For Output As outfil
completare = ""
    For i = 1 To (80 - Len(title))
    completare = completare + " "
    Next i
    Print #outfil, title + completare
    linie_de_80 = "BANK 1 " & 5 * CLng(numarvalori / 5) & " " & CLng(numarvalori / 5) & " CONST " & (startx * 100) & " " & (stepx * 100) & " 0 0 ESD"
    completare = ""
    For i = 1 To (80 - Len(linie_de_80))
    completare = completare + " "
    Next i
    linie_de_80 = linie_de_80 + completare
    Print #outfil, linie_de_80
''If Not (amzdata) Then
''ReDim z(numarvalori)
'' For i = 1 To numarvalori
''  z(i) = 1
''  Next i
''End If
For i = 1 To Fix(numarvalori / 5) '+ 1
Print #outfil, Format$(Format$(Val(Y((i - 1) * 5 + 1)), "####0.0#"), "@@@@@@@@") + Format$(Format$(z((i - 1) * 5 + 1), "##0.0###"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 5 + 2)), "####0.0#"), "@@@@@@@@") + Format$(Format$(z((i - 1) * 5 + 2), "##0.0###"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 5 + 3)), "####0.0#"), "@@@@@@@@") + Format$(Format$(z((i - 1) * 5 + 3), "##0.0###"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 5 + 4)), "####0.0#"), "@@@@@@@@") + Format$(Format$(z((i - 1) * 5 + 4), "##0.0###"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 5 + 5)), "####0.0#"), "@@@@@@@@") + Format$(Format$(z((i - 1) * 5 + 5), "##0.0###"), "@@@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 5 * Fix(numarvalori / 5) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = (5 * Fix(numarvalori / 5) + 1) To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "####0.0#"), "@@@@@@@@") + Format$(Format$(z(i), "##0.0###"), "@@@@@@@@")
Next i


    completare = ""
   For i = 1 To (80 - Len(ultimalinie))
   completare = completare + " "
   Next i
   ultimalinie = ultimalinie + completare
Print #outfil, ultimalinie
End If


Close
raport "GSAS ESD file written. Attention: you may have to adjust the datafile."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
raport Now
Exit Sub
End Sub

Private Sub mnuSaveGsasSTD_Click()
Dim returncode As Boolean, i As Long, completare As String, linie_de_80 As String, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie

raport "Trying to save a GSAS STD file. Checking..."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The requested file is " & outputfile
Open outputfile For Output As outfil
completare = ""
    For i = 1 To (80 - Len(title))
    completare = completare + " "
    Next i
    Print #outfil, title + completare
    linie_de_80 = "BANK 1 " & 10 * CLng(numarvalori / 10) & " " & CLng(numarvalori / 10) & " CONST " & (startx * 100) & " " & (stepx * 100) & " 0 0 STD"
    completare = ""
    For i = 1 To (80 - Len(linie_de_80))
    completare = completare + " "
    Next i
    linie_de_80 = linie_de_80 + completare
    Print #outfil, linie_de_80
For i = 1 To Fix(numarvalori / 10) '+ 1
Print #outfil, "  " + Format$(Format$(Val(Y((i - 1) * 10 + 1)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 2)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 3)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 4)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 5)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 6)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 7)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 8)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 9)), "#####0"), "@@@@@@") + "  " + Format$(Format$(Val(Y((i - 1) * 10 + 10)), "#####0"), "@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 10 * Fix(numarvalori / 10) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 10 * Fix(numarvalori / 10) + 1 To numarvalori
ultimalinie = ultimalinie + "  " + Format$(Format$(Val(Y(i)), "#####0"), "@@@@@@")
Next i

    completare = ""
    For i = 1 To (80 - Len(ultimalinie))
    completare = completare + " "
    Next i
    ultimalinie = ultimalinie + completare
Print #outfil, ultimalinie
End If

Close
raport "GSAS STD file written. Attention: you may have to adjust the datafile."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveLHPM_Click()
Dim returncode As Boolean, i As Long, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "Output in LHPM format."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
IamBusy True
raport "The file is " & outputfile
Open outputfile For Output As outfil
Print #outfil, Format$(Format$(startx, "#0.000##"), "@@@@@@@@"); Format$(Format$(stepx, "#0.000##"), "@@@@@@@@"); Format$(Format$(endx, "###0.0##"), "@@@@@@@@"), left$(title, 30)
For i = 1 To Fix(numarvalori / 10)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 10 + 1)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 2)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 3)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 4)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 5)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 6)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 7)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 8)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 9)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 10)), "#######0"), "@@@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 10 * Fix(numarvalori / 10) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 10 * Fix(numarvalori / 10) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "#######0"), "@@@@@@@@")
Next i
Print #outfil, ultimalinie
End If
Close
IamBusy False
raport "LHPM file saved, it seems to be ok...Check and adjust the output data to your needs."
raport Now
Exit Sub
errorTRAP:
Err.Clear
IamBusy False
raport "An error has occured."
Close
Exit Sub
End Sub

Private Sub mnuSavePhilips_Click()
Dim returncode As Boolean, i As Long, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie

raport "Philips PC-UDF format - APD 3.51  "
raport "Warning: there are various UDF header formats. Check the output file."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The outputfile is " & outputfile
Open outputfile For Output As outfil
Print #outfil, "SampleIdent," & left$(title, 30) & ",/"
Print #outfil, "Title1, " + prog_name + ",/"
Print #outfil, "Title2," & left$(title, 25) & ",/"
Print #outfil, "DiffrType,PW3710,/"
Print #outfil, "DiffrNumber,1,/"
Print #outfil, "Anode, Cu,/"
Print #outfil, "LabdaAlpha1, " + Format$(Format$(1.5406, "0.00000"), "@@@@@@@") & ",/"
Print #outfil, "LabdaAlpha2, " + Format$(Format$(1.54439, "0.00000"), "@@@@@@@") & ",/"
Print #outfil, "RatioAlpha21, " + Format$(Format$(0.5, "0.00000"), "@@@@@@@") & ",/"
Print #outfil, "DivergenceSlit,Fixed,1,/"
Print #outfil, "ReceivingSlit,0.1,/"
Print #outfil, "MonochromatorUsed,NO  ,/"
Print #outfil, "GeneratorVoltage,  40,/"
Print #outfil, "TubeCurrent,  20,/"
Print #outfil, "FileDateTime," & Format(Now, "dd-mmm-yyyy") & " " & Format(Now, "ttttt") & ",/"
Print #outfil, "DataAngleRange,  " & Format$(Format$((startx), "#0.0000"), "@@@@@@@") & ",  " & Format$(Format$((endx), "#0.0000"), "@@@@@@@") & ",/"
Print #outfil, "ScanStepSize,   " & Format$(Format$((stepx), "0.000"), "@@@@@") & ",/"
Print #outfil, "ScanType,STEP,/"
Print #outfil, "ScanStepTime,    5.00,/"
Print #outfil, "RawScan"
For i = 1 To Fix(numarvalori / 8)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 8 + 1)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 2)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 3)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 4)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 5)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 6)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 7)), "#######0"), "@@@@@@@@") + "," + Format$(Format$(Val(Y((i - 1) * 8 + 8)), "#######0"), "@@@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 8 * Fix(numarvalori / 8) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 8 * Fix(numarvalori / 8) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "#######0"), "@@@@@@@@") + ","
Next i
Print #outfil, ultimalinie
End If
Print #outfil, "/"
Close
raport "Philips UDF file written; some of the fields are default - date, wavelength..."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveReportPad_Click()
Dim returncode As Boolean, outfil As Integer
On Error GoTo errorTRAP
raport strLinie & vbCrLf & "Saving report"
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The file is " & outputfile
Open outputfile For Output As outfil
Print #outfil, Convert3Main.txtraport.Text
Close outfil
Err.Clear
Exit Sub
errorTRAP:
'cancel or something
Exit Sub
End Sub

Private Sub mnuSaveRiet_Click()
Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "Riet7 ascii file."
raport "Warning: Details ignored (wavelength, date, etc). You will need to edit the header."
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
IamBusy True
raport "The filename is " & outputfile
Open outputfile For Output As outfil
title = title & "                                        "
Print #outfil, left$(title, 37) & "DataFileName " & left$(outputfile, 8)
Print #outfil, "DiffrType PW3710   GeneratorVoltage 40   TubeCurrent 40"
Print #outfil, "Anode Cu    Alpha1  1.54056    Alpha2  1.54439    Ratio  0.50000"
Print #outfil, "MonochromatorUsed YES   DivergenceSlit 1      ReceivingSlit 0.3"
Print #outfil, Format$(Format$(startx, "#0.000##"), "@@@@@@@@"); Format$(Format$(stepx, "#0.000##"), "@@@@@@@@"); Format$(Format$(endx, "###0.0##"), "@@@@@@@@") & "    MeasureDateTime 19/02/1999   15:29   StepTime   3.00"
For i = 1 To Fix(numarvalori / 10)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 10 + 1)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 2)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 3)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 4)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 5)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 6)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 7)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 8)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 9)), "#######0"), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 10 + 10)), "#######0"), "@@@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 10 * Fix(numarvalori / 10) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 10 * Fix(numarvalori / 10) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "#######0"), "@@@@@@@@")
Next i
Print #outfil, ultimalinie
End If
Close
raport "Riet7 file saved, it seems to be ok...Check and adjust the header."
raport Now
IamBusy False
Exit Sub
errorTRAP:
Err.Clear
IamBusy False
Close
raport "An error has occured."
Exit Sub







End Sub

Private Sub mnuSaveScintag_Click()
Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "Scintag ascii file."
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The filename is " & outputfile
Call averageint(numarvalori, Y, valmedie, valmax, returncode)
If Not (returncode) Then raport "An error has occured; still trying..."
raport "The average value of intensity is " & CStr(valmedie)
raport "The maximum value of intensity is " & CStr(valmax)

Open outputfile For Output As outfil
raport "Most of the fields are fixed, check the outputfile."
''If Not (amzdata) Then
''ReDim z(numarvalori)
''For i = 1 To numarvalori: z(i) = 1: Next i
''raport "I don't have Z data, put 1.0 insted."
''End If
Print #outfil, title
Print #outfil, Format$(Format$(startx, "###0.0000000"), "@@@@@@@@@@@@") & Format$(Format$(endx, "###0.0000000"), "@@@@@@@@@@@@") & Format$(Format$(stepx, "###0.0000000"), "@@@@@@@@@@@@") & " 10.00000000  0.00000000  1.54059994"
Print #outfil, Format$(Format$(numarvalori, "#######0"), "@@@@@@@@") & "       0       0"
Print #outfil, Format(Now, "  mm  dd  yy")
Print #outfil, Format(Now, "  hh  mm  ss")
Print #outfil, "    4451" & Format$(Format$(CInt(valmedie), "########"), "@@@@@@@@") & Format$(Format$(valmax, "#######."), "@@@@@@@@")
Print #outfil, "   0   0   0   0   0   0   0   0   0   0   0   0"
For i = 1 To numarvalori
Print #outfil, Format$(Format$(Val(X(i)), "###0.00000"), "@@@@@@@@@@") & Format$(Format$(Val(Y(i)), "######0."), "@@@@@@@@") & Format$(Format$(Val(z(i)), "######0."), "@@@@@@@@")
Next i

Close
raport "A Scintag file type written; some of the fields are default - date, wavelength..."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveSiemens_Click()
Dim returncode As Boolean, i As Long, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie

raport "Siemens ascii file."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
Call averageint(numarvalori, Y, valmedie, valmax, returncode)
If Not (returncode) Then raport "An error has occured; still trying..."
raport "The average value of intensity is " & CStr(valmedie)
raport "The maximum value of intensity is " & CStr(valmax)

raport "The filename is " & outputfile
Open outputfile For Output As outfil
raport "Most of the fields are fixed, check the datafile."
Print #outfil, ";AW      Converted by " & prog_name
Print #outfil, "_FILEVERSION=' '"
Print #outfil, "_SAMPLE='" + title + "'"
Print #outfil, "_+SAMPLE=' Old file structure " + prog_name + "'"
Print #outfil, "_DATAMEASURED='" + Format(Now, "dd-mmm-yyyy") & " " & Format(Now, "ttttt") & "'"
Print #outfil, "_WL1=1.54056"
Print #outfil, "_WL2=1.54439"
Print #outfil, "_WL3=1.3921"
Print #outfil, "_WLRATIO=0.500"
Print #outfil, "_ANODE=Cu"
Print #outfil, ";Range  1"
Print #outfil, "_DRIVE='COUPLED'"
Print #outfil, "_STEPTIME= 10.0"
Print #outfil, "_STEPSIZE= " + Format$(Format$((stepx), "0.0000000"), "@@@@@@@@@")
Print #outfil, "_STEPMODE=S"
Print #outfil, "_START=" + Format$(Format$((startx), "#0.000"), "@@@@@@")
Print #outfil, "_2THETA=" + Format$(Format$((startx), "#0.000"), "@@@@@@")
Print #outfil, "_THETA=" + Format$(Format$((startx / 2), "#0.000"), "@@@@@@")
Print #outfil, "_KHI=   0.000"
Print #outfil, "_PHI=   0.000"
Print #outfil, "_STEPCOUNT= " + Format$(Format$(((endx - startx) / stepx), "#0000"), "@@@@@")
Print #outfil, "_COUNTS"
For i = 1 To Fix(numarvalori / 8)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 8 + 1)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 2)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 3)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 4)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 5)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 6)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 7)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 8)), "######0."), "@@@@@@@@")
Next i

'scriu eventualele puncte ramase...
If 8 * Fix(numarvalori / 8) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 8 * Fix(numarvalori / 8) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "######0."), "@@@@@@@@")
Next i
Print #outfil, ultimalinie
End If


Close
raport "Siemens ascii file written; some of the fields are default - date, wavelength..."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveSietronics_Click()

Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "Sietronics CPI datafile"
Call verificadate(False, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
IamBusy True
raport "The filename is " & outputfile
Open outputfile For Output As outfil
Print #outfil, "SIETRONICS XRD SCAN"
Print #outfil, CStr(startx)
Print #outfil, CStr(endx)
Print #outfil, CStr(stepx)
Print #outfil, "Cu"
Print #outfil, "1.54056"
Print #outfil, "16-1-1998"
Print #outfil, "1"
Print #outfil, "SampleIdent" & left$(title, 25)
Print #outfil, "SCANDATA"
For i = 1 To (numarvalori)
Print #outfil, CStr(Y(i))
Next i
Close
IamBusy False
raport "Sietronics CPI ascii file written; you might need to adjust the header..." & vbCrLf & Now
Exit Sub
errorTRAP:
IamBusy False
Err.Clear
Close
raport "An error has occured."
Exit Sub


End Sub

Private Sub mnuSaveWppf1_Click()
Dim returncode As Boolean, i As Long, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie

raport "WPPF/Profit file format 1."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The output file is " & outputfile
IamBusy True
Open outputfile For Output As outfil
Print #outfil, left$(title, 60)
Print #outfil, "  0  " + Format$(0, "0") + Format$(Format$(1.54056, "0.0000##"), "@@@@@@@@") & Format$(Format$(1.54439, "0.0000##"), "@@@@@@@@") & Format$(Format$(0.5, "0.0000##"), "@@@@@@@@") & Format$(Format$(startx, "##0.0000"), "@@@@@@@@") & Format$(Format$(endx, "##0.0000"), "@@@@@@@@") & Format$(Format$(stepx, "#0.0000#"), "@@@@@@@@") & "  5.0000"
For i = 1 To Fix(numarvalori / 10)
Print #outfil, Format$(Format$(Y((i - 1) * 10 + 1), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 2), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 3), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 4), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 5), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 6), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 7), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 8), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 9), "######0"), "@@@@@@@") + Format$(Format$(Y((i - 1) * 10 + 10), "######0"), "@@@@@@@")
Next i

'scriu eventualele puncte ramase...
If 10 * Fix(numarvalori / 10) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 10 * Fix(numarvalori / 10) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "######0"), "@@@@@@@")
Next i
Print #outfil, ultimalinie
End If

IamBusy False
Close
raport "Wppf file written; some of the fields are default - , wavelength..."
raport Now
Exit Sub
errorTRAP:
IamBusy False
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveWppf2_Click()
Dim returncode As Boolean, i As Long, ultimalinie As String, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie

raport "WPPF/Profit file format 2."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The output file is " & outputfile

Open outputfile For Output As outfil
Print #outfil, left$(title, 60)
Print #outfil, "  0  " + Format$(0, "0") + Format$(Format$(1.54056, "0.0000##"), "@@@@@@@@") & Format$(Format$(1.54439, "0.0000##"), "@@@@@@@@") & Format$(Format$(0.5, "0.0000##"), "@@@@@@@@") & Format$(Format$(startx, "##0.0000"), "@@@@@@@@") & Format$(Format$(endx, "##0.0000"), "@@@@@@@@") & Format$(Format$(stepx, "#0.0000#"), "@@@@@@@@") & "  5.0000"
For i = 1 To Fix(numarvalori / 8)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 8 + 1)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 2)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 3)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 4)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 5)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 6)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 7)), "######0."), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 8)), "######0."), "@@@@@@@@")
Next i

'scriu eventualele puncte ramase...
If 8 * Fix(numarvalori / 8) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 8 * Fix(numarvalori / 8) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "######0."), "@@@@@@@@")
Next i
Print #outfil, ultimalinie
End If


Close
raport "Wppf file written; some of the fields are default - , wavelength..."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveXY_Click()

Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "XY data file, ascii."
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The filename is " & outputfile
Open outputfile For Output As outfil
Print #outfil, left$(title, 30) & " modified by "; prog_name & "  -  " & Now

For i = 1 To (numarvalori)
Print #outfil, Format$(Format$(Val(X(i)), "######0.00000##"), "@@@@@@@@@@@@@@@") & ",   " & Format$(Format$(Val(Y(i)), "######0.000##"), "@@@@@@@@@@@@@")
Next i
Close
raport "X, Y ascii file written; one record per line..."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveXYZ_Click()
Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "XYZ data file."
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The filename is " & outputfile
Open outputfile For Output As outfil
Print #outfil, left$(title, 30) & " modified by "; prog_name & " " & Now
'If Not (amzdata) Then
'ReDim z(numarvalori)
'For i = 1 To numarvalori: z(i) = 1: Next i
'raport "I don't have Z data, put 1.0 insted."
'End If

For i = 1 To numarvalori
Print #outfil, Format$(Format$(Val(X(i)), "######0.00000##"), "@@@@@@@@@@@@@@@") & ",   " & Format$(Format$(Val(Y(i)), "######0.000##"), "@@@@@@@@@@@@@") & ",   " & Format$(Format$(Val(z(i)), "######0.00#"), "@@@@@@@@@@@")
Next i
Close
raport "X ,Y ,Z ascii file; one record per line, done..."
raport Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuSaveY_Click()
Dim returncode As Boolean, i As Long, valmedie As Double, valmax As Double
On Error GoTo errorTRAP
raport strLinie
raport "Y datafile"
Call verificadate(False, False, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The filename is " & outputfile
Open outputfile For Output As outfil
Print #outfil, left$(title, 30) & " modified by "; prog_name & " " & Now
For i = 1 To (numarvalori)
Print #outfil, Format$(Format$(Val(Y(i)), "######0.000##"), "@@@@@@@@@@@@@")
Next i
Close
raport "Y ascii file written; one record per line..." & vbCrLf & Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuScintag_Click()
'format descris de Linda Mansker/New Mexico
'neglijez o buna parte din definitiile de header
raport strLinie

amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
raport "Reading data file in Scintag format."
raport "Some of the fields in the header will be neglected (date, time, WL ratio...)."
Dim return_code As Boolean
Dim linie As String, junkstring As String, nr_linii As Long, junkdata As Double, cod_eroare As Integer
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The filename is " & inputfile
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines. You should have " & CStr(nr_linii - 7) & " data points."
raport CStr(ignoralinii) & " line(s) will be ignored."
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
linii_titlu = 0
cod_eroare = 0
If Not (mnuWithout.Checked) Then
raport "Reading some info from header."
'citeste intreg headerul intai, dupa fisierul trimis de mansker:
'-------------------------------------------
' INC 200 WEEKEND, VAC. DRY
''   5.0000000  70.0000000   0.0200000  1.19999993  0.00000000  1.54059994
''    3200       0       0
''   1  29  98
''  19  28   0
''    6551    2609  14350.
''   0   0   0   0   0   0   0   0   0   0   0   0
''   5.00000   2500.    353.
''   5.02000   2500.    353.
''   5.04000   2400.    346.
Line Input #inpfil, title
Input #inpfil, startx, endx, stepx
Input #inpfil, junkdata
Input #inpfil, junkdata
Input #inpfil, junkdata
Input #inpfil, numarvalori
''4 linii sarite

For i = 1 To 5
Line Input #inpfil, linie
Next i

raport "Domain in header ; start, step, end :" & CStr(startx) & ", " & CStr(stepx) & ", " & CStr(endx)

Else
    numarvalori = (nr_linii - 1 - ignoralinii)
raport "No header definition: the maximum data points is " & CStr(numarvalori)
End If
ReDim X(numarvalori), Y(numarvalori), z(numarvalori)
For i = 1 To numarvalori
Input #inpfil, X(i), Y(i), z(i)
Next i
amxdata = True: amydata = True: amzdata = True: amfullxdata = True
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuShellNotepad_Click()
On Error GoTo errorTRAP
Shell "notepad.exe", vbMaximizedFocus
SendKeys "%"
SendKeys "{enter}"
Exit Sub
errorTRAP:
MsgBox "Error: maybe Notepad is missing ?"
Exit Sub
End Sub

Private Sub mnuSiemens_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
raport strLinie & vbCrLf & "Reading data file in Siemens format."
Dim return_code As Boolean
Dim linie As String, t As String, tt As String, nr_linii As Long, linii_titlu As Integer, cod_eroare As Integer
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The filename is " & inputfile
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines" & vbCrLf & CStr(ignoralinii) & " line(s) will be ignored."
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
linii_titlu = 0
cod_eroare = 0
If Not (mnuWithout.Checked) Then
raport "Reading some info from header."
'citeste intreg headerul intai, dupa fisierul trimis de ddragoi ar arata cam asa:
'-------------------------------------------
'AW     Converted to UXD format by ...prima linie e un titlu
'iau din acest header doar y si start, step, end
'linia cu _sample care o consider titlu
'liniile cu _wl1, _wl2 _wl3 si _wlratio
'linia cu _anode
'linia cu stepsize
'linia cu _start
'linia cu _stepcounts
'linia cu _counts
Do
Line Input #inpfil, tt
cod_eroare = cod_eroare + 1

t = LCase$(left$(tt, 4))
Select Case t
Case "_sam"
title = Val(right$(tt, Len(tt) - 9))


Case "_sta"
startx = Val(right$(tt, Len(tt) - 7))

Case "_ste"
If LCase$(left$(tt, 10)) = "_stepcount" Then
numarvalori = Val(right$(tt, Len(tt) - 11))
raport "The number of data points is " & CStr(numarvalori)
Else
If LCase$(left$(tt, 9)) = "_stepsize" Then
stepx = Val(right$(tt, Len(tt) - 11))
End If
End If

Case "_cou"
'nimic special, trece la wend
End Select

If cod_eroare = 50 Then raport "Fatal error. I can not find the _COUNTS statement. ": Close: Exit Sub
linii_titlu = linii_titlu + 1
Loop Until t = "_cou"
'cod_eroare este nedefinit
If numarvalori = 0 Then raport "The label _stepcount could not be found.  The number of points is estimated based on the number of lines of the file.":     numarvalori = (nr_linii - linii_titlu - ignoralinii) * 8
    endx = numarvalori * stepx + startx
    raport "Domain in header ; start, step, end :" & CStr(startx) & ", " & CStr(stepx) & ", " & CStr(endx)

Else
    numarvalori = (nr_linii - linii_titlu - ignoralinii) * 8
raport "No header definition: the maximum data points is " & CStr(numarvalori)
End If ' este sfarsitul de la header_from
'de aici incepe partea de citire standard
ReDim Y(numarvalori + 8)
''For i = 1 To CInt(numarvalori / 8)
''Input #inpfil, Y((i - 1) * 8 + 1), Y((i - 1) * 8 + 2), Y((i - 1) * 8 + 3), Y((i - 1) * 8 + 4), Y((i - 1) * 8 + 5), Y((i - 1) * 8 + 6), Y((i - 1) * 8 + 7), Y((i - 1) * 8 + 8)
''Next i
'noua versiune - citeste si liniile incomplete-poate adauga niste zerouri atunci cand ceva e gresit...
i = 0
Do While Not (EOF(inpfil))
i = i + 1
Input #inpfil, Y(i)
Loop
Close #inpfil
raport "Done..." & vbCrLf & Now & vbCrLf & strLinie
amxdata = True: amydata = True
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured. " & vbCrLf
Exit Sub
End Sub

Private Sub mnuSwap_Click()
Dim returncode As Boolean, i As Long, zz As Double
On Error GoTo errorTRAP
raport strLinie

raport "Swap X, Y data. This is a strange option for X ray diffraction data. "
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
For i = 1 To numarvalori
zz = X(i)
X(i) = Y(i)
Y(i) = zz
Next i
raport "Done..." & vbCrLf & Now
amfullxdata = True
amydata = True
Exit Sub
errorTRAP:
Err.Clear
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuTruncate_Click()
On Error GoTo errorTRAP
Dim t As Integer, sT As Single, en As Single, i As Single, j As Single
'check if I have data

Dim returncode As Boolean
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Err.Raise 1102, , ""
DoEvents
amxdata = True

Dim tempx() As Double, tempy() As Double
ReDim tempx(numarvalori), tempy(numarvalori)
On Error GoTo errorTRAP
t = MsgBox("This function truncate the data (you don't normally need this..). The data in memory will be irreversibly lost. Continue ?", vbOKCancel + vbDefaultButton2, prog_name)
If t = vbCancel Then Err.Raise 1101, , ""
sT = InputBox("Input initial 2theta you want to keep", prog_name, Val(startx))
en = InputBox("Input final 2theta you want to keep", prog_name, Val(endx))

If sT >= en Then Err.Raise 1101, , ""
If sT < X(1) Then sT = X(1)
If en > X(numarvalori) Then en = X(numarvalori)
j = 0
For i = 1 To numarvalori
If sT <= X(i) And en >= X(i) Then
j = j + 1
tempx(j) = X(i)
tempy(j) = Y(i)
End If
Next i
numarvalori = j
startx = sT
endx = en
For i = 1 To j
X(i) = tempx(i)
Y(i) = tempy(i)
Next i

raport strLinie
raport "Data truncated; kept only between " & CStr(sT) & " and " & CStr(en)
raport strLinie
Exit Sub
errorTRAP:
Err.Clear
Exit Sub
End Sub

Private Sub mnuWithout_Click()
mnuWithout.Checked = Not (mnuWithout.Checked)
DoEvents
If mnuWithout.Checked Then raport "No header in the inputfile."
End Sub



Private Sub mnuWppf_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
raport "Reading WPPF/Profit file."
Dim return_code As Boolean, i As Long, nfor As Integer, nr_linii As Long
Dim linie As String, wave_number As Integer
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The inputfile is " & inputfile
'citesc cate linii are fisierul, fac dimensionarea si verific formatul
Call numar_linii(inputfile, nr_linii)
raport "The number of lines in this file is " & CStr(nr_linii)
raport CStr(ignoralinii) & "line(s) will be ignored."
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
If Not (mnuWithout.Checked) Then
    Line Input #inpfil, title
    raport "The title in this file is: " & title
Input #inpfil, nfor, wave_number, startx, endx, stepx
    numarvalori = (nr_linii - 2 - ignoralinii) * 10
    If nfor Then numarvalori = (nr_linii - 2 - ignoralinii) * 8
    Else
    'este 0 10i7 sau 1 8f8.0
    Open inputfile For Input As inpfil
    Line Input #inpfil, linie
    nfor = 0
    If (InStr(linie, ".") > 0) Then nfor = 1
    numarvalori = (nr_linii - ignoralinii) * 10
    If nfor Then numarvalori = (nr_linii - ignoralinii) * 8
    End If
ReDim Y(numarvalori)
raport "The number of datapoints is " & CStr(numarvalori)

Select Case nfor
Case 0, 1
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i
Case Else
raport "Error in WPPF/PROFIT header. Look at the NFOR descriptor."
End Select
Close
raport "Done..." & vbCrLf & Now
Exit Sub

errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuWWW_Click()
On Error GoTo errorTRAP
ret& = ShellExecute(Me.hwnd, "Open", AuthorWebPage, "", App.Path, 1)
Exit Sub
errorTRAP:
Err.Clear
On Error GoTo errortrap2
ret& = ShellExecute(Me.hwnd, "Open", "http://www.upsud.fr/lpces", "", App.Path, 1)
Err.Clear
Exit Sub
errortrap2:
raport "An error occured. You don't have a browser ?.." & vbCrLf & Now
Err.Clear
Exit Sub
End Sub

Private Sub mnuXY_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
raport strLinie
raport "X, Y : ascii file."
Dim return_code As Boolean, i As Long, nr_linii As Long, a As Double, b As Double
Dim linie As String
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The filename is " & inputfile
raport CStr(ignoralinii) & " line(s) will be ignored."
Call numar_linii(inputfile, nr_linii)
raport "This file has " & CStr(nr_linii) & " lines"
numarvalori = 0
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If

'Do Until (EOF(inpfil))
'Input #inpfil, a, b
'numarvalori = numarvalori + 1
'Loop

Do While Not (EOF(inpfil))
Input #inpfil, a, b
numarvalori = numarvalori + 1
Loop

Close #inpfil
numarvalori = numarvalori - ignoralinii
amfullxdata = True: amydata = True
ReDim X(numarvalori), Y(numarvalori)
raport "The number of points is " & CStr(numarvalori)
If Not (mnuWithout.Checked) Then raport "There is no header for this file type..."

Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If

For i = 1 To numarvalori
Input #inpfil, X(i), Y(i)
Next i

Close #inpfil
startx = X(1): stepx = X(2) - X(1)
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:

'Err.Clear
'numarvalori = numarvalori - 1
'Resume OutLoop
'End If
raport "An error has occured; number " & CStr(Err.Number) & ". " & Err.Description
Err.Clear
Close
Exit Sub
End Sub

Private Sub mnuXYZ_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
raport strLinie & vbCrLf & "X, Y, Z : ascii file."
Dim return_code As Boolean, nr_linii As Long, a As Double, b As Double, c As Double, i As Long
Dim linie As String
inpfil = FreeFile
inputfile = ""
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The filename is " & inputfile
Open inputfile For Input As inpfil
numarvalori = 0
raport CStr(ignoralinii) & " line(s) will be ignored."
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
Do While Not (EOF(inpfil))
Input #inpfil, a, b, c
numarvalori = numarvalori + 1
Loop
Close
ReDim X(numarvalori), Y(numarvalori), z(numarvalori)
raport "The number of points is " & CStr(numarvalori)
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
'If Not (mnuWithout.Checked) Then raport "There are no header for this file type..."
For i = 1 To numarvalori
Input #inpfil, X(i), Y(i), z(i)
Next i
Close
startx = X(1)
endx = X(numarvalori)
stepx = CLng(CInt(100000 * X(2) - 100000 * X(1))) / 100000
amxdata = True: amfullxdata = True: amydata = True: amzdata = True

raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mnuY_Click()
amxdata = False: amfullxdata = False: amydata = False: amzdata = False: title = ""
On Error GoTo errorTRAP
Dim return_code As Boolean, nr_linii As Long, dlinie As Double, i As Long, inpfil As Integer
Dim linie As String
inpfil = FreeFile
inputfile = ""
raport strLinie

raport "Y data file, ascii"
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
raport "The filename is " & inputfile
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
raport "The number of points is: " & CStr(numarvalori)
amydata = True
ReDim Y(numarvalori)
Open inputfile For Input As inpfil
If ignoralinii > 0 Then
For i = 1 To ignoralinii
Line Input #inpfil, linie
Next i
End If
'If Not (mnuWithout.Checked) Then raport "There are no header for this file type..."
For i = 1 To numarvalori
Input #inpfil, Y(i)
Next i
Close
raport "Done..." & vbCrLf & Now
Exit Sub
errorTRAP:
Err.Clear
Close
raport "An error has occured."
Exit Sub
End Sub

Private Sub mOpenCustom_Click()

ReadCustom.Show
End Sub

Private Sub mWWWCCp14_Click()


On Error GoTo errorTRAP
ret& = ShellExecute(Me.hwnd, "Open", "http://www.ccp14.ac.uk/index.html", "", App.Path, 1)
Exit Sub
errorTRAP:
Err.Clear
On Error GoTo errortrap2
raport "An error has occured." & vbCrLf & CStr(Now)
Err.Clear
Exit Sub
errortrap2:
raport "An error occured. You don't have a browser ?..Try inserting the address http:\\www.ccp14.ac.uk in your browser" & vbCrLf & Now
Err.Clear
Exit Sub







End Sub


Private Sub Popup1_Click()

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'se introduce un d
Dim X As Double
DoEvents
On Error GoTo errorTRAP
X = Val(txt.Text) / 2 / Val(Text1.Text)
Text2.Text = Format$(180 / 3.14157 * (2 * Atn(X / Sqr(1 - X * X))), "##0.00000")
Exit Sub
errorTRAP:
Text2.Text = ""
Err.Clear
Exit Sub
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
'se introduce un 2theta
Dim X As Double
DoEvents
On Error GoTo errorTRAP

Text1.Text = Format$(Val(txt.Text) / 2 / Sin(Val(Text2.Text / 2) * 3.1415 / 180), "##0.00000")
Exit Sub
errorTRAP:
Text1.Text = ""
Err.Clear
Exit Sub
End Sub

Private Sub timerSplash_Timer()
Unload frmSplash
timerSplash.Interval = 0
timerSplash.Enabled = False
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errorTRAP
Dim X As Double
Combo.ListIndex = 24
DoEvents
X = Val(txt.Text) / 2 / Val(Text1.Text)
Text2.Text = Format$(180 / 3.14157 * (2 * Atn(X / Sqr(1 - X * X))), "##0.00000")
Exit Sub
errorTRAP:
Err.Clear
Exit Sub
End Sub


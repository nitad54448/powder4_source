VERSION 5.00
Begin VB.Form frmgDBWSmain 
   Caption         =   "gDBWS"
   ClientHeight    =   3444
   ClientLeft      =   132
   ClientTop       =   504
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3444
   ScaleWidth      =   6540
   Begin VB.TextBox txtMainMessage 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   3252
      HideSelection   =   0   'False
      Left            =   0
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6492
   End
   Begin VB.Menu mnuMainFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveReportPad 
         Caption         =   "Save Report pad"
      End
      Begin VB.Menu mnuMainNotepadExport 
         Caption         =   "Export Report pad"
         Visible         =   0   'False
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMainEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuMainEditControlFile 
         Caption         =   "Control File"
      End
      Begin VB.Menu mnuMiainEditPreferences 
         Caption         =   "Preferences"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMainRun 
      Caption         =   "Run"
      Begin VB.Menu mnuMainRunRietveld 
         Caption         =   "DBWS 98"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mSelectData 
         Caption         =   "Select DBWS Data File"
      End
      Begin VB.Menu mSelectControl 
         Caption         =   "Select DBWS Control File"
      End
      Begin VB.Menu mSelectOutput 
         Caption         =   "Select DBWS Output File"
      End
      Begin VB.Menu mnuMainRunSimulate 
         Caption         =   "Simulate pattern"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainRunFourier 
         Caption         =   "Fourier analysis"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuMainGraphic 
      Caption         =   "Graphic"
      Begin VB.Menu mnuMainGraphicUpdatePlot 
         Caption         =   "Show Refinement Plots"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainGraphicDBWSPlot 
         Caption         =   "Open DBWS Plot file"
      End
      Begin VB.Menu mnuMainGraphicFourierMap 
         Caption         =   "Fourier Map"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMainHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmgDBWSmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'set global parameters
phaseEdited = 1
AtomEdited = 1

'resize the txtmainMessage object at startup
frmgDBWSmain.txtMainMessage.Width = Me.Width - 140
frmgDBWSmain.txtMainMessage.Height = Me.Height - 720

RietveldBoardMessage "Rietveld utilities started at " & Now
RietveldBoardMessage "This program is based on DBWS-9807; proper citation is required:" & vbCrLf & "Young, R. A.; Sakthivel, A.; Moss, T. S.; Paiva-Santos, C. O., J. Appl. Cryst., 1995, 28, 366-367."
RietveldBoardMessage strLinie & vbCrLf & "The default DBWS Control File is saved in the directory : " & App.Path & "\dbws\pw_ControlFile.cfg. You can edit this file for your default parameters."
End Sub



Private Sub Form_Resize()
'resize the txtmainMessage object when the form is resized
'windowstate=1 means minimized
If frmgDBWSmain.WindowState = 0 Or frmgDBWSmain.WindowState = 2 Then
frmgDBWSmain.txtMainMessage.Width = Me.Width - 140
frmgDBWSmain.txtMainMessage.Height = Me.Height - 720
End If

'give some slack
DoEvents
End Sub

Sub Form_Unload(cancel As Integer)
'If cancel = 2 Then cancel = 0: Exit Sub

If Not (CloseWindow("Are you sure to close this window ?", prog_name & " gDBWS")) Then cancel = -1
End Sub

Private Sub mnuMainEditControlFile_Click()
frmControlFile.Show
End Sub


Private Sub mnuMainExit_Click()
Unload Me
Exit Sub
End Sub

Private Sub mnuMainGraphicDBWSPlot_Click()
On Error GoTo errorTRAP
Dim return_code As Boolean
inpfil = FreeFile
inputfile = "plotinfo"

Call open_file(inputfile, 1, return_code)
If Not (return_code) Then
inputfile = ""
Exit Sub
End If

newGraph.Show
Exit Sub
errorTRAP:
Exit Sub
End Sub

Private Sub mnuMainRunRietveld_Click()
'put a shell here
On Error GoTo errorTRAP
Dim t As Double, sT As String
'ChDir App.Path & "\dbws"
If dbwsDataFile = "" Then dbwsDataFile = App.Path & "\dbws\pw_data.txt"
If dbwsControlFile = "" Then dbwsControlFile = App.Path & "\dbws\pw_icf.txt"
If dbwsOutputFile = "" Then dbwsOutputFile = App.Path & "\dbws\pw_dbws.out"

If Not (left$(dbwsDataFile, 1) = """") Then dbwsDataFile = """" & dbwsDataFile & """"
If Not (left$(dbwsOutputFile, 1) = """") Then dbwsOutputFile = """" & dbwsOutputFile & """"
If Not (left$(dbwsControlFile, 1) = """") Then dbwsControlFile = """" & dbwsControlFile & """"



sT = InputBox("This command will run the program pw_DBWS3.  You can change the filenames either manually or by selecting them with the Run/Select menu commands.  The order in this list is (i.e. datafile filename, control filename, output filename)", prog_name & " - shell DBWS98", dbwsDataFile & " " & dbwsControlFile & " " & dbwsOutputFile)
If Len(sT) < 2 Then Exit Sub
t = ShellAndLoop(App.Path & "\dbws\pw_DBWS3.exe " & sT, vbMaximizedFocus)
RietveldBoardMessage "pw_DBWS3 called at " & Now
RietveldBoardMessage strLinie

hWndShell "write.exe " & dbwsOutputFile, vbNormalFocus
Exit Sub

errorTRAP:
RietveldBoardMessage Err.Description
RietveldBoardMessage Now
RietveldBoardMessage strLinie
Err.Clear
Exit Sub

End Sub



Private Sub mnuSaveReportPad_Click()
Dim returncode As Boolean, outfil As Integer
On Error GoTo errorTRAP
RietveldBoardMessage strLinie & vbCrLf & "Saving report"
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
RietveldBoardMessage "The file is " & outputfile
Open outputfile For Output As outfil
Print #outfil, frmgDBWSmain.txtMainMessage.Text
Close outfil
Err.Clear
Exit Sub
errorTRAP:
'cancel or something
Exit Sub
End Sub

Private Sub mSelectControl_Click()
On Error GoTo error_open
Me.MousePointer = 11
'Select Case intrare_iesire
'    Case 1 'citirea path_input_text
Convert3Main.Dialog.Filter = "text file (*.txt) |*.txt|data file (*.dat) |*.dat|show all (*.*) |*.*"
Convert3Main.Dialog.FilterIndex = 3
Convert3Main.Dialog.Flags = &H1000& Or &H4& Or &H800&
'ofn_filemustexist 'ofn_readonly 'ofn_pathmustexist
Convert3Main.Dialog.DialogTitle = prog_name & " - select DBWS Control file"
Convert3Main.Dialog.Action = 1
dbwsControlFile = Convert3Main.Dialog.FileName
Me.MousePointer = 0

Exit Sub
error_open:
Err.Clear
dbwsControlFile = ""
Me.MousePointer = 0
Exit Sub

End Sub

Private Sub mSelectData_Click()

On Error GoTo error_open
Me.MousePointer = 11
'Select Case intrare_iesire
'    Case 1 'citirea path_input_text
Convert3Main.Dialog.Filter = "text file (*.txt) |*.txt|data file (*.dat) |*.dat|show all (*.*) |*.*"
Convert3Main.Dialog.FilterIndex = 3
Convert3Main.Dialog.Flags = &H1000& Or &H4& Or &H800&
'ofn_filemustexist 'ofn_readonly 'ofn_pathmustexist
Convert3Main.Dialog.DialogTitle = prog_name & " - select DBWS data file"
Convert3Main.Dialog.Action = 1
dbwsDataFile = Convert3Main.Dialog.FileName

Me.MousePointer = 0

Exit Sub
error_open:
Err.Clear
dbwsDataFile = ""
Me.MousePointer = 0
Exit Sub


End Sub

Private Sub mSelectOutput_Click()
On Error GoTo error_open
Me.MousePointer = 11

Convert3Main.Dialog.Filter = "text file (*.txt) |*.txt|data file (*.dat) |*.dat|show all (*.*) |*.*"
Convert3Main.Dialog.FilterIndex = 3
Convert3Main.Dialog.FileName = ""
Convert3Main.Dialog.Flags = &H2& Or &H1& Or &H800& Or &H4&
'ofn_overwriteprompt'ofn_readonly'ofn_pathmustexist
Convert3Main.Dialog.DialogTitle = prog_name & " - Select DBWS Output file"
Convert3Main.Dialog.Action = 2
dbwsOutputFile = Convert3Main.Dialog.FileName
return_code = True

Me.MousePointer = 0
Close
Exit Sub
error_open:
Err.Clear
dbwsOutputFile = ""
return_code = False
Me.MousePointer = 0
Close
Exit Sub
End Sub

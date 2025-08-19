VERSION 5.00
Begin VB.Form merge 
   Caption         =   "Merge Files"
   ClientHeight    =   5440
   ClientLeft      =   880
   ClientTop       =   840
   ClientWidth     =   6320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5440
   ScaleWidth      =   6320
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   9
      Left            =   3600
      TabIndex        =   33
      Text            =   "1.0"
      Top             =   2880
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   8
      Left            =   3600
      TabIndex        =   32
      Text            =   "1.0"
      Top             =   3360
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   7
      Left            =   3600
      TabIndex        =   31
      Text            =   "1.0"
      Top             =   3840
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   6
      Left            =   3600
      TabIndex        =   30
      Text            =   "1.0"
      Top             =   4320
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   5
      Left            =   3600
      TabIndex        =   29
      Text            =   "1.0"
      Top             =   4800
      Width           =   610
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   9
      Left            =   840
      TabIndex        =   28
      Top             =   2880
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   8
      Left            =   840
      TabIndex        =   27
      Top             =   3360
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   7
      Left            =   840
      TabIndex        =   26
      Top             =   3840
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   6
      Left            =   840
      TabIndex        =   25
      Top             =   4320
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   5
      Left            =   840
      TabIndex        =   24
      Top             =   4800
      Width           =   2410
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   9
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   8
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   7
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   6
      Left            =   240
      TabIndex        =   20
      Top             =   4320
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   370
   End
   Begin VB.CommandButton btnMergeCancel 
      Caption         =   "&Cancel"
      Height          =   370
      Left            =   4920
      TabIndex        =   16
      Top             =   1080
      Width           =   1090
   End
   Begin VB.CommandButton btnMergeOK 
      Caption         =   "O&K"
      Height          =   370
      Left            =   4920
      TabIndex        =   15
      Top             =   480
      Width           =   1090
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   370
   End
   Begin VB.CommandButton btnFile 
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   370
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   4
      Left            =   3600
      TabIndex        =   9
      Text            =   "1.0"
      Top             =   2400
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   3
      Left            =   3600
      TabIndex        =   8
      Text            =   "1.0"
      Top             =   1920
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      Text            =   "1.0"
      Top             =   1440
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Text            =   "1.0"
      Top             =   960
      Width           =   610
   End
   Begin VB.TextBox txtNormFactor 
      Height          =   300
      Index           =   0
      Left            =   3600
      TabIndex        =   5
      Text            =   "1.0"
      Top             =   480
      Width           =   610
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2410
   End
   Begin VB.TextBox txtMerge 
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2410
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   5160
      Y2              =   480
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Normalization"
      Height          =   370
      Index           =   1
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   1570
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Filenames"
      Height          =   370
      Index           =   0
      Left            =   840
      TabIndex        =   17
      Top             =   120
      Width           =   2290
   End
End
Attribute VB_Name = "merge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFile_Click(Index As Integer)
    'title = prog_name
    
    Dim return_code As Boolean, nr_linii As Long, linie As String
    inpfil = FreeFile
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
txtMerge(Index).Text = inputfile

Exit Sub

End Sub

Private Sub btnMergeCancel_Click()
Unload Me
End Sub

Private Sub btnMergeOK_Click()
'merge xyz
Dim mergeXmin As Double, mergeXmax As Double
Dim mergeStep(10) As Double, pointerX(10) As Long
Dim j As Integer, ign As Integer
Dim return_code As Boolean, i As Long, nr_linii As Long, a As Double, b As Double, c As Double
Dim linie As String
Dim numvalori(10) As Long, normFactor(10) As Double, mergeFile(10) As Boolean
Dim returncode As Boolean
Dim Xt() As Double, Yt() As Double, Zt2() As Double, Ft() As Double, f() As Double
Dim jj As Integer
On Error GoTo errorTRAP
inpfil = FreeFile
mergeXmin = 180
mergeXmax = 0
For j = 0 To 9
mergeFile(j + 1) = False
inputfile = txtMerge(j).Text
normFactor(j + 1) = CDbl(txtNormFactor(j).Text)
If Not (txtMerge(j).Text = "") Then
    numarvalori = 0
    mergeFile(j + 1) = True
    Open inputfile For Input As inpfil
        For ign = 1 To ignoralinii: Line Input #inpfil, linie: Next ign
        Do While Not (EOF(inpfil))
        Input #inpfil, a, b, c
        numarvalori = numarvalori + 1
    Loop
    Close #inpfil
    jj = jj + 1
numvalori(j + 1) = numarvalori
End If

Next j
'jj is the number of files to be merged
'0 means there is no file
    If jj = 0 Then
        MsgBox "No data files have been selected. Please try again. "
        Exit Sub
    End If
'keep temporary at 32700
ReDim Xt(10, 32700)
ReDim Yt(10, 32700)
ReDim Zt2(10, 32700)
ReDim Ft(10, 32700)
For j = 0 To 9
If mergeFile(j + 1) Then
    Open txtMerge(j).Text For Input As inpfil
    For ign = 1 To ignoralinii: Line Input #inpfil, linie: Next ign

    Do While Not (EOF(inpfil))
        For i = 1 To numvalori(j + 1)
            Input #inpfil, Xt(j + 1, i), Yt(j + 1, i), Zt2(j + 1, i)
            Zt2(j + 1, i) = Zt2(j + 1, i) ^ 2
        Next i
    Loop
    Close #inpfil
'startx = X(1): stepx = X(2) - X(1)
    mergeStep(j + 1) = CDbl(CLng(100000000 * (Xt(j + 1, 2) - Xt(j + 1, 1))) / 100000000)
    stepx = mergeStep(j + 1)
If Xt(j + 1, 1) < mergeXmin Then mergeXmin = Xt(j + 1, 1)
If Xt(j + 1, numvalori(j + 1)) > mergeXmax Then mergeXmax = Xt(j + 1, numvalori(j + 1))
End If
Next j
'pointerX(5) is a long, show where the common point is
For j = 1 To 10
If mergeFile(j) Then
pointerX(j) = CLng((Xt(j, 1) - mergeXmin) / stepx) + 1
If normFactor(j) = 0 Then Err.Raise 1101, , "Error: check the normalization factor. "
End If
Next j

numarvalori = CLng((mergeXmax - mergeXmin) / stepx) + 1
ReDim X(numarvalori)
ReDim Y(numarvalori)
ReDim z(numarvalori)
ReDim f(numarvalori)


ReDim Xt(10, numarvalori)
ReDim Yt(10, numarvalori)
ReDim Zt2(10, numarvalori)
ReDim Ft(10, numarvalori)

For i = 1 To numarvalori
X(i) = mergeXmin + (i - 1) * stepx
For j = 1 To 10
Xt(j, i) = mergeXmin + (i - 1) * stepx
Yt(j, i) = 0
Ft(j, i) = 0
Zt2(j, i) = 0
Next j
Next i

For j = 1 To 10
If mergeFile(j) Then
    Open txtMerge(j - 1).Text For Input As inpfil
        For ign = 1 To ignoralinii: Line Input #inpfil, linie: Next ign
    
    Do While Not (EOF(inpfil))
        For i = 0 To numvalori(j) - 1
            Input #inpfil, Xt(j, i + pointerX(j)), Yt(j, i + pointerX(j)), Zt2(j, i + pointerX(j))
            Ft(j, i + pointerX(j)) = normFactor(j)
            'Yt(j, i + pointerX(j)) = Yt(j, i + pointerX(j)) / normFactor(j)
            Zt2(j, i + pointerX(j)) = Zt2(j, i + pointerX(j)) ^ 2
            
        Next i
    Loop
    Close #inpfil
End If
Next j
For j = 1 To 10
    If mergeFile(j) Then
       For i = 1 To numarvalori
        Y(i) = Y(i) + Yt(j, i)
        f(i) = f(i) + Ft(j, i)
        z(i) = z(i) + Zt2(j, i)
        Next i
    End If
Next j
For i = 1 To numarvalori
If f(i) > 0 Then
Y(i) = Y(i) / f(i)
z(i) = Sqr(z(i)) / f(i)
End If
Next i
'write here the merged xyz file, ask for a file first
amfullxdata = True: amydata = True: amzdata = True
startx = mergeXmin: endx = mergeXmax
'checked on 15 of july
Call verificadate(False, True, True, returncode)
If Not (returncode) Then Exit Sub
outfil = FreeFile
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
raport "The filename is " & outputfile
Open outputfile For Output As outfil
Print #outfil, left$(title, 30) & " -> merged  " & Now
For i = 1 To (numarvalori)
Print #outfil, Format$(Format$(Val(X(i)), "##0.00000###"), "@@@@@@@@@@@@") & " , " & Format$(Format$(Val(Y(i)), "#####0.00###"), "@@@@@@@@@@@@") & " , " & Format$(Format$(Val(z(i)), "#####0.00###"), "@@@@@@@@@@@@")
Next i
Close
MsgBox CStr(jj) & " files merged. Please check carefully the output data file."
raport "X, Y, ESD ascii file written; one record per line..."
raport Now
Exit Sub

Exit Sub
errorTRAP:
raport Err.Description
Err.Clear
Close
raport "An error has occured, routine : merge XYZ files " & CStr(Now)
raport "Common cause : not an XYZ input data file.  Otherwise check the data files for EOF character (use Tools/Remove strange characters ! ), have the files the same step ? "
Exit Sub
End Sub


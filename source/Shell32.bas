Attribute VB_Name = "Shell32"

' ****************************************************************
'  Shell32.Bas, Copyright ©1996-97 Karl E. Peterson
' ****************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' ****************************************************************
'  Three methods to "Shell and Wait" under Win32.
'  One deals with the infamous "Finished" behavior of Win95.
'  Fourth method that simply shells and returns top-level hWnd.
' ****************************************************************
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const SYNCHRONIZE = &H100000

Public Const WAIT_FAILED = -1&        'Error on call
Public Const WAIT_OBJECT_0 = 0        'Normal completion
Public Const WAIT_ABANDONED = &H80&   '
Public Const WAIT_TIMEOUT = &H102&    'Timeout period elapsed
Public Const IGNORE = 0               'Ignore signal
Public Const INFINITE = -1&           'Infinite timeout

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9

Private Const WM_CLOSE = &H10
Private Const GW_HWNDNEXT = 2
Private Const GW_OWNER = 4

Public Function ShellAndWait(ByVal JobToDo As String, Optional ExecMode, Optional TimeOut) As Long
   '
   ' Shells a new process and waits for it to complete.
   ' Calling application is totally non-responsive while
   ' new process executes.
   '
   Dim ProcessID As Long
   Dim hProcess As Long
   Dim nRet As Long
   Const fdwAccess = SYNCHRONIZE

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         ShellAndWait = vbObjectError + Err.Number
         Exit Function
      End If
   On Error GoTo 0

   If IsMissing(TimeOut) Then
      TimeOut = INFINITE
   End If

   hProcess = OpenProcess(fdwAccess, False, ProcessID)
   nRet = WaitForSingleObject(hProcess, CLng(TimeOut))
   Call CloseHandle(hProcess)

   Select Case nRet
      Case WAIT_TIMEOUT: Debug.Print "Timed out!"
      Case WAIT_OBJECT_0: Debug.Print "Normal completion."
      Case WAIT_ABANDONED: Debug.Print "Wait Abandoned!"
      Case WAIT_FAILED: Debug.Print "Wait Error:"; Err.LastDllError
   End Select
   ShellAndWait = nRet
End Function

Public Function ShellAndLoop(ByVal JobToDo As String, Optional ExecMode) As Long
   '
   ' Shells a new process and waits for it to complete.
   ' Calling application is responsive while new process
   ' executes. It will react to new events, though execution
   ' of the current thread will not continue.
   '
   Dim ProcessID As Long
   Dim hProcess As Long
   Dim nRet As Long
   Const fdwAccess = PROCESS_QUERY_INFORMATION

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         ShellAndLoop = vbObjectError + Err.Number
         Exit Function
      End If
   On Error GoTo 0

   hProcess = OpenProcess(fdwAccess, False, ProcessID)
   Do
      GetExitCodeProcess hProcess, nRet
      DoEvents
      Sleep 100
   Loop While nRet = STILL_ACTIVE
   Call CloseHandle(hProcess)

   ShellAndLoop = nRet
End Function

Public Function ShellAndClose(ByVal JobToDo As String, Optional ExecMode) As Long
   '
   ' Shells a new process and waits for it to complete.
   ' Calling application is responsive while new process
   ' executes. It will react to new events, though execution
   ' of the current thread will not continue.
   '
   ' Will close a DOS box when Win95 doesn't. More overhead
   ' than ShellAndLoop but useful when needed.
   '
   Dim ProcessID As Long
   Dim PID As Long
   Dim hProcess As Long
   Dim hWndJob As Long
   Dim nRet As Long
   Dim TitleTmp As String
   Const fdwAccess = PROCESS_QUERY_INFORMATION

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         ShellAndClose = vbObjectError + Err.Number
         Exit Function
      End If
   On Error GoTo 0

   hWndJob = FindWindow(vbNullString, vbNullString)
   Do Until hWndJob = 0
      If GetParent(hWndJob) = 0 Then
         Call GetWindowThreadProcessId(hWndJob, PID)
         If PID = ProcessID Then Exit Do
      End If
      hWndJob = GetWindow(hWndJob, GW_HWNDNEXT)
   Loop

   hProcess = OpenProcess(fdwAccess, False, ProcessID)
   Do
      TitleTmp = Space(256)
      nRet = GetWindowText(hWndJob, TitleTmp, Len(TitleTmp))
      If nRet Then
         TitleTmp = UCase(left(TitleTmp, nRet))
         If InStr(TitleTmp, "FINISHED") = 1 Then
            Call SendMessage(hWndJob, WM_CLOSE, 0, 0)
         End If
      End If

      GetExitCodeProcess hProcess, nRet
      DoEvents
      Sleep 100
   Loop While nRet = STILL_ACTIVE
   Call CloseHandle(hProcess)

   ShellAndClose = nRet
End Function

Public Function hWndShell(ByVal JobToDo As String, Optional ExecMode) As Long
   '
   ' Shells a new process and returns the hWnd
   ' of its main window.
   '
   Dim ProcessID As Long
   Dim PID As Long
   Dim hProcess As Long
   Dim hWndJob As Long

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         hWndShell = 0
         Exit Function
      End If
   On Error GoTo 0

   hWndJob = FindWindow(vbNullString, vbNullString)
   Do While hWndJob <> 0
      If GetParent(hWndJob) = 0 Then
         Call GetWindowThreadProcessId(hWndJob, PID)
         If PID = ProcessID Then
            hWndShell = hWndJob
            Exit Do
         End If
      End If
      hWndJob = GetWindow(hWndJob, GW_HWNDNEXT)
   Loop
End Function

Function sForFormat(vTrans As Variant, sFormat As String) As String
'this function gets a variant Trans and returns a string sForFormat using a FORTRAN77-like
'definition. This function is necessary because of many incompatibilities of Visual Basic Format$ function
'with the Fortran one (the Integer alignment, etc...).  Works only for writting; the reading in VB is flexible !!
'it will accept three types of SIMPLE Fortran-like formats Fa.b, Ax, Iy (no multiple or embedded formats)
'Warning: I don't keep the compatibility of "****" return of this function in Fortran, just truncate (this can be adapted easily)
'the sign is included in the length of the fortran definition,..
'so, if it's + ignore it (or better delete it), for "-" keep in mind that the total length is smaller with one unit
'made by N.D. on december 11th, 2000
'tested on december 11th, 2000; seems OK but not tested extensively
'adjusted on 12th of december, minor errors found

Dim i As Integer, j As Integer, point As Integer, bNegValue As Boolean, trans As String
Dim nrInChar() As String * 1, nrOutChar() As String * 1, lungime As Integer, floats As Integer, nskip As Integer
On Error GoTo handleit

trans = CStr(vTrans) 'make the variant as string
Select Case UCase$(Mid$(sFormat, 1, 1)) ' main select ------------------------------

Case "F"
    'the format is F lungime.floats
    'for floats 0 there is a special case
    'determin how long will be the string at exit, it should include the sign
'bnegvalue is a logical showing if the value is negative
'bNegValue = False
'there is a problem with the variant conversion
'If Val(trans) < 0 Then bNegValue = True: trans = Mid$(trans, 2, Len(trans) - 1)
    lungime = CInt(Val(Mid$(sFormat, 2)))
    floats = CInt(Val(Mid$(sFormat, 1 + (InStr(sFormat, ".")))))
'point shows the position in the initial string where the decimal . is located
point = 0
ReDim nrOutChar(lungime)
'check if there is a decimal point, otherwise add one . at the end
If InStr(trans, ".") = 0 Then trans = trans & "."
point = InStr(trans, ".")
ReDim nrInChar(Len(trans))
    For j = 1 To CInt(Len(trans))
        nrInChar(j) = Mid$(trans, j, 1)
    Next j
''determine  a correct float value if it uses the Fx.0 stupid format expression
'consider first the most significant digits
If floats = 0 Then
    j = 0
    For i = 1 To Len(trans)
        If ((Asc(nrInChar(i)) = 46)) Then Exit For
    
    
    Next i
    floats = lungime - i '- 1
           
'    floats = lungime - 1 'the case with decimals only
    '''
'    j = 0
'    For i = lungime To 1 Step -1
'    j = j + 1
'        If i <= Len(trans) Then
'        'get out of here at the first significant digit
'        If ((Asc(nrInChar(i)) = 46) Or (Asc(nrInChar(i)) - 48) > 0) Then Exit For
'        End If
'    Next i
'the j here is the possible float
'    floats = floats - j - 1 'this will drop a decimal ?/
End If


'first put spaces in the outstring
    For i = 1 To lungime
        nrOutChar(i) = "_"
    Next i
    'normal fortran format
    'put the . in the corect place
    nrOutChar(lungime - floats) = "."
    'now put the rest values starting from the point down and up, whichever order
    'mind that some of the zero might be insignificant while the other not
    j = 0
    For i = lungime - floats + 1 To lungime
        j = j + 1
        If point + j > Len(trans) Then Exit For
        nrOutChar(i) = nrInChar(point + j)
    Next i
    'now I have to write the first digits..
    j = 0
    For i = (lungime - floats - 1) To 1 Step -1
    j = j + 1
    If (point - j) = 0 Then Exit For
        nrOutChar(i) = nrInChar(point - j)
    Next i

'remove the nonsignificant first zeros
    For i = 1 To lungime - floats - 1
        'get out of here at the first significant digit
        If (((Asc(nrOutChar(i)) = 46) Or (nrOutChar(i) = "-") Or (Asc(nrOutChar(i)) - 48) > 0)) And Not (nrOutChar(i) = "_") Then Exit For
        nrOutChar(i) = "_"
    Next i
If floats > 0 Then
If (nrOutChar(lungime - floats + 1) = "_") Then nrOutChar(lungime - floats + 1) = "0"
Else
If (nrOutChar(lungime - 1) = "_") Then nrOutChar(lungime - 1) = "0"
End If

nskip = 0
    
    For i = 1 To lungime
        If nrOutChar(i) = "_" Then
        'skip transfer
        nskip = nskip + 1
        Else
        sForFormat = sForFormat & nrOutChar(i)
        End If
    Next i
'pad with spaces up to lungime both in the left and in the right side, if possible
'estetic reason only
Select Case nskip
Case 0
'nothing
Case Else
'add a space to the left, the rest (if any) to the right
sForFormat = " " & sForFormat
For i = 2 To nskip
sForFormat = sForFormat & " "
Next i
End Select


Case "I" '----------------integer definition--------------------------------------------
'this part returns right-alignement integer
'bnegvalue is a logical showin if the value is negative
bNegValue = False
If Val(trans) < 0 Then bNegValue = True: trans = Mid$(trans, 2, Len(trans) - 1)
lungime = CInt(Val(Mid$(sFormat, 2)))
'len(trans) must be smaller than lungime,..no check is made here
trans = CStr(CInt(Val(trans))) ' attention, here I use CLong to have more digits possible
'this makes the equivalent of a Kind_4 definition in Fortran 95
sForFormat = trans
If bNegValue Then sForFormat = "-" & sForFormat
'pad with spaces to the left
Do
    If Len(sForFormat) = lungime Then Exit Do
    sForFormat = " " & sForFormat
    'not so fast but nicer than making selfreference into a for loop
    Loop


Case "A" '-----------------------------Strings---------------------------
'piece of cake; neglect the BZ or Holerith symbols
'alignement here is to the left, pad with spaces at the end
lungime = CInt(Val(Mid$(sFormat, 2)))
ReDim nrOutChar(lungime)
ReDim nrInChar(Len(trans))
    For j = 1 To CInt(Len(trans))
        nrInChar(j) = Mid$(trans, j, 1)
    Next j
    If lungime > Len(trans) Then
        For i = 1 To Len(trans)
            nrOutChar(i) = nrInChar(i)
        Next i
        'now add the spaces
        For i = Len(trans) + 1 To lungime
            nrOutChar(i) = " "
        Next i
    Else
    'truncate it
        For i = 1 To lungime
            nrOutChar(i) = nrInChar(i)
        Next i
    End If
'final assignement
    For i = 1 To lungime
        sForFormat = sForFormat & nrOutChar(i)
    Next i

Case Else
'nothing yet, to do someday the E
End Select 'end of the main select------------------------------------------

Exit Function
handleit:
    MsgBox Err.Description
    Err.Clear
    Exit Function
    
End Function


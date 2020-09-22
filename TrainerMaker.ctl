VERSION 5.00
Begin VB.UserControl TrainerMaker 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   MaskPicture     =   "TrainerMaker.ctx":0000
   Picture         =   "TrainerMaker.ctx":4042
   ScaleHeight     =   945
   ScaleWidth      =   960
   ToolboxBitmap   =   "TrainerMaker.ctx":8084
End
Attribute VB_Name = "TrainerMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Paint()
Width = 64 * Screen.TwipsPerPixelX
Height = 64 * Screen.TwipsPerPixelY
End Sub

Public Function WriteToMemory(MemoryAddress As Long, Length As Integer, Value As Long, ProgramClass As String) As Boolean
Dim ProcesshWnd As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
ProcesshWnd = FindWindow(ProgramClass, vbNullString)
If ProcesshWnd = False Then WriteToMemory = False: Exit Function
GetWindowThreadProcessId ProcesshWnd, ProcessID
ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
Call WriteProcessMemory(ProcessHandle, MemoryAddress, Value, Length, 0&)
CloseHandle ProcessHandle
WriteToMemory = True
End Function

Public Function ReadFromMemory(ProgramClass As String, MemoryAddress As Long, Length As Long) As Long
Dim ProcesshWnd As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
Dim buffer As Long
ProcesshWnd = FindWindow(ProgramClass, vbNullString)
If ProcesshWnd = False Then ReadFromMemory = False: Exit Function
GetWindowThreadProcessId ProcesshWnd, ProcessID
ProcessHandle = OpenProcess(&H1F0FFF, False, ProcessID)
Call ReadProcessMemory(ProcessHandle, MemoryAddress, buffer, Length, 0&)
CloseHandle ProcessHandle
ReadFromMemory = buffer
End Function

Public Function ReadFloatFromMemory(ProgramClass As String, MemoryAddress As Long, Length As Long) As Single
Dim ProcesshWnd As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
Dim buffer As Long
ProcesshWnd = FindWindow(ProgramClass, vbNullString)
If ProcesshWnd = False Then ReadFloatFromMemory = False: Exit Function
GetWindowThreadProcessId ProcesshWnd, ProcessID
ProcessHandle = OpenProcess(&H1F0FFF, False, ProcessID)
Call ReadProcessMemory(ProcessHandle, MemoryAddress, buffer, Length, 0&)
CloseHandle ProcessHandle
ReadFloatFromMemory = buffer
End Function

Public Function IsProgramRunning(ProgramCaption As String, ProgramClass As String, FindWay As FindOptions)
Dim ProgCaption As String
Dim ProgClass As String
Select Case FindWay:
Case 1:
ProgCaption = vbNullString
ProgClass = ProgramClass
Case 2:
ProgCaption = ProgramCaption
ProgClass = vbNullString
Case 3:
ProgCaption = ProgramCaption
ProgClass = ProgramClass
End Select
IsProgramRunning = IIf(FindWindow(ProgClass, ProgCaption) <> 0, True, False)
End Function

Public Function TerminateWindow(WindowCaption As String, WindowClass As String) As Boolean
Dim XTemp As Long
XTemp = FindWindow(WindowClass, WindowCaption)
TerminateWindow = False
If XTemp <> 0 Then
SendMessage XTemp, WM_CLOSE, 0, 0
TerminateWindow = True
End If
End Function

Public Function CheckForTS(WichVersion As TrainerSpy_Versions) As Boolean
Select Case WichVersion:
Case 1:
CheckForTS = IIf(FindWindow(vbNullString, "TRAINER SPY") <> 0, True, False)
Case 2:
CheckForTS = IIf(FindWindow(vbNullString, "TrainerSpy XP + NT / 2000 / XP + Coded By BofeN") <> 0, True, False)
End Select
End Function

Public Function TerminateTS()
On Error Resume Next
' Now we will terminate TrainerSpy, I know the caption (title) of game spy window
' So I will terminate the process by caption
Call TerminateWindow("TRAINER SPY", vbNullString)
Call TerminateWindow("TrainerSpy XP + NT / 2000 / XP + Coded By BofeN", vbNullString)
' Kill TrainerSpy logging file wich is:
'           C:\logwmemory.bin
'--------------------------------------
' First, reset the attributes to normal
SetAttr "C:\logwmemory.bin", vbNormal
' Second, KILL THE DAMN FILE!!!
Kill "C:\logwmemory.bin"
' Now make an empty file
Dim TempSA As SECURITY_ATTRIBUTES
Call CreateFile("C:\logwmemory.bin", &H1F0FFF, 1, TempSA, 2, 1, 0)
' Now attribute it as SYSTEM+READ ONLY+HIDDIN so that TrainerSpy cannot overwrite it
SetAttr "C:\logwmemory.bin", vbReadOnly + vbHidden + vbSystem
End Function

Public Function ActivateWindow(WindowCaption As String, WindowClass As String, Options As ShowOptions) As Boolean
Dim ProgramHWND As Long
ProgramHWND = FindWindow(WindowClass, WindowCaption)
If ProgramHWND = 0 Then ActivateWindow = False: Exit Function
Call ShowWindow(ProgramHWND, Options)
ActivateWindow = True
End Function

Public Function SendKeysToGame(ByVal KeysToSend As String)
 Pause 500     ' Pause for sometime, while windows switch us to the game
 SendKeys KeysToSend
End Function

Public Function WaitHotKey(HotKeyToWait As Long) As Boolean
WaitHotKey = False
If GetAsyncKeyState(HotKeyToWait) <> 0 Then WaitHotKey = True
End Function

Public Function ExecuteProgram(ProgramToExecute As String, ShowCommand As ShowOptions, Optional RunDirectory = "C:\")
Call ShellExecute(UserControl.hwnd, "open", ProgramToExecute, 0, RunDirectory, ShowCommand)
End Function

Public Function About()
On Error Resume Next
frmAbout.Show vbModal
End Function
Public Function Pause(MilliSeconds As Long)
Dim Something%, CurrentTicket As Long
CurrentTicket = GetTickCount()
Do
Something% = DoEvents
Loop Until CurrentTicket + MilliSeconds < GetTickCount
End Function

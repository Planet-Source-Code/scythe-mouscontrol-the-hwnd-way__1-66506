VERSION 5.00
Begin VB.UserControl MouseControl 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   495
   ToolboxBitmap   =   "UserControl1.ctx":27A2
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   360
   End
End
Attribute VB_Name = "MouseControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Easy MouseOver
'© Scythe
'scythe@cablenet.de

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointApi) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type PointApi
 x As Long
 Y As Long
End Type

Private Type ControlNames
 Handle As Long
 CtrlName As String
 CtrlIdx As Long
End Type

Dim Contro() As ControlNames
Dim Ctr As Long
Dim LastControl As Long
Dim ButtonDown As Boolean
Dim BtnCtrl As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Const VK_MBUTTON = 4

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Type Msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As PointApi
End Type
Private Const WM_MOUSEWHEEL = 522


Public Event MouseOver(ControlName As String, Index As Long)
Public Event MouseOut(ControlName As String, Index As Long)
Public Event MouseMiddle(ControlName As String, Index As Long, ButtonDown As Boolean)
Public Event MouseWheel(ControlName As String, Index As Long, ScrollUp As Boolean)

Public Sub Init(Frm As Form)
ReDim Contro(1000)
Dim x As Long
Dim Ctrl As Control
Dim Tmp As Long

'Search all controls on the Frame we use
For Each Ctrl In Frm.Controls
 Tmp = CheckForHwnd(Ctrl)
 If Tmp <> -1 Then
  Contro(x).CtrlName = Ctrl.Name
  Contro(x).Handle = Tmp
  Contro(x).CtrlIdx = CheckForIndex(Ctrl)
 x = x + 1
 End If
Next

Ctr = x - 1
ReDim Preserve Contro(Ctr)
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()

 Dim CursorPos As PointApi
 Dim hWndOver As Long
 Dim Whl As Msg
 Dim WhUp As Boolean
 
 'Get cursor position
 GetCursorPos CursorPos
 'Get window/control cursor is over
 hWndOver = WindowFromPoint(CursorPos.x, CursorPos.Y)
 
 'If the control under the mouse changed
 If hWndOver <> LastControl Then
  
  'Fire MouseOut Event
  If LastControl <> 0 Then
   For i = 0 To Ctr
    If Contro(i).Handle = LastControl Then
     RaiseEvent MouseOut(Contro(i).CtrlName, Contro(i).CtrlIdx)
     Exit For
    End If
   Next i
  End If
  
  'Fire MouseOver Event
  For i = 0 To Ctr
   If Contro(i).Handle = hWndOver Then
    RaiseEvent MouseOver(Contro(i).CtrlName, Contro(i).CtrlIdx)
    LastControl = hWndOver
    Exit Sub
   End If
  Next i
  
  LastControl = 0
 End If
 
 
 'Fire MiddleButton Event
 If ButtonDown <> CBool(GetAsyncKeyState(VK_MBUTTON)) Then
  ButtonDown = Not ButtonDown
  If ButtonDown Then BtnCtrl = hWndOver
  For i = 0 To Ctr
   If Contro(i).Handle = BtnCtrl Then
    RaiseEvent MouseMiddle(Contro(i).CtrlName, Contro(i).CtrlIdx, ButtonDown)
   End If
  Next i
 End If
 
 'Fire Wheel Event
 GetMessage Whl, 0, 0, 0
 DispatchMessage Whl
 If Whl.Message = WM_MOUSEWHEEL Then
  For i = 0 To Ctr
   If Contro(i).Handle = Whl.hwnd Then
   If Whl.wParam > 0 Then WhUp = True
    If Whl.Message = WM_MOUSEWHEEL Then RaiseEvent MouseWheel(Contro(i).CtrlName, Contro(i).CtrlIdx, WhUp)
   End If
  Next i
 End If
 
End Sub

'See if the Control has an Index
Private Function CheckForIndex(Ctrl As Control) As Long
 On Error GoTo ErrOut
 CheckForIndex = Ctrl.Index
 Exit Function
ErrOut:
 CheckForIndex = -1
End Function
Private Function CheckForHwnd(Ctrl As Control) As Long
On Error GoTo ErrOut
 CheckForHwnd = Ctrl.hwnd
 Exit Function
ErrOut:
 CheckForHwnd = -1
End Function
Private Sub UserControl_Resize()
  UserControl.Height = 540
  UserControl.Width = 540
End Sub

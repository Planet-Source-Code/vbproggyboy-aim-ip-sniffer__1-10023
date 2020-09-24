<div align="center">

## AIM IP Sniffer


</div>

### Description

This code gets an AIM users IP address through an IM window.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[vbproggyboy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vbproggyboy.md)
**Level**          |Intermediate
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vbproggyboy-aim-ip-sniffer__1-10023/archive/master.zip)

### API Declarations

(none)


### Source Code

```
'Add two textboxes, one for the persons screen 'name and the other for what the link should say
'Add a command button to send the IM
'Add a ListBox, so the IPs can be stored in it
'Add a winsock control
Private Sub Command1_Click()
Call SendIM(Text1, "<a XXXX=" & """" & Winsock1.LocalIP & """" & ">" & Text2 & "<\a>)
End Sub
'XXXX = href
Private Sub Form_Load()
winsock1.localport = 80
winsock1.listen
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
list1.additem winsock1.remotehostip ' Adds the remote IP address to the list box
End Sub
' Add the following code to a module
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' Global & Public Const
Const EM_UNDO = &HC7
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const ENTA = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const EM_LINESCROLL = &HB6
Private Const SPI_SCREENSAVERRUNNING = 97
Type RECT
  Left As Long
  Top As Long
  Right As Long
  bottom As Long
End Type
Type POINTAPI
  X As Long
  y As Long
End Type
Sub IM_Send(SendName As String, SayWhat As String, CloseIM As Boolean)
' My send IM comes with a little thing where you can eather close
' it or not close it....
' Ex: Call IM_Send("ThereSn","Sup man",True) <-- that closes the IM
' Put False to not close the IM, All the IM sends have the TRUE FALSE thing
  Dim BuddyList As Long
  BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
  If BuddyList& <> 0& Then
    GoTo Start
  Else
   Exit Sub
  End If
Start:
  Dim TabWin As Long, IMButtin As Long, IMWin As Long
  Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
  Dim EditThing As Long, TextSet2 As Long, SendButtin As Long, Click As Long
  BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
  TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
  IMButtin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
  Click& = SendMessage(IMButtin&, WM_LBUTTONDOWN, 0, 0&)
  Click& = SendMessage(IMButtin&, WM_LBUTTONUP, 0, 0&)
  IMWin& = FindWindow("AIM_IMessage", vbNullString)
  ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
  TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
  TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
  EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
  EditThing& = GetWindow(EditThing&, 2)
  TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, SayWhat$)
  SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
  Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
  Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
  If CloseIM = True Then
    Win_Killwin (IMWin&)
  Else
    Exit Sub
  End If
End Sub
Sub Win_Killwin(TheWind&)
  Call PostMessage(TheWind&, WM_CLOSE, 0&, 0&)
End Sub
'If you have any questions or problems please leave feedback, or email me at vbproggy_boy@hotmail.com, thanks
```


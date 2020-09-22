Attribute VB_Name = "modMenuHint"
Option Explicit

Public lpPrevWndProc As Long
Public gHWnd As Long

'-------------------------------------------------------------------

' Win API: WinProcHandler
'-------------------------------------------------------------------
Public Const GWL_WNDPROC = (-4)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'-------------------------------------------------------------------

' Win API: MenuHandling
'-------------------------------------------------------------------
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const MF_STRING = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_BITMAP = &H4&
Public Const MF_CHECKED = &H8&
Public Const MF_POPUP = &H10&
Public Const MF_HILITE = &H80&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_SEPARATOR = &H800&
Public Const MF_SYSMENU = &H2000&
Public Const MF_MOUSESELECT = &H8000&

Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Public Sub pHook(ByVal lHwnd As Long)

  ' Sub class the form to trap for Windows messages.
  lpPrevWndProc = SetWindowLong(lHwnd, GWL_WNDPROC, AddressOf fWindowProc)

End Sub

Public Sub pUnhook(ByVal lHwnd As Long)

  ' Remove the subclassing.
  Call SetWindowLong(lHwnd, GWL_WNDPROC, lpPrevWndProc)

End Sub

Function fWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim lhMenu As Long
  Dim lMenuItem As Long
  Dim lFlags As Long

  If (hw = gHWnd) And (uMsg = WM_MENUSELECT Or uMsg = WM_MENUCHAR) Then
    lMenuItem = LoWord(wParam)
    lFlags = HiWord(wParam)
    lhMenu = lParam
    frmMain.StatusBar1.SimpleText = GetMenuHint(lhMenu, lMenuItem, lFlags)
  End If
  
  ' Call the original window procedure associated with this form.
  fWindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)

End Function

Function GetMenuHint(ByVal lhMenu As Long, ByVal lMenuItem As Long, ByVal lFlags As Long) As String

  Dim sMenuString As String
  Dim lResult As Long
  Dim lcmdFlag As Long
  
  GetMenuHint = ""
  
' Flags which indicates, that the item is not a valid selected menu-entry.
  If (lFlags And MF_SEPARATOR) = MF_SEPARATOR Then Exit Function
  If (lFlags And MF_HILITE) = 0 Then Exit Function
  
  lcmdFlag = MF_BYCOMMAND
  If (lFlags And MF_POPUP) = MF_POPUP Then lcmdFlag = MF_BYPOSITION

' Get Item-Caption
  sMenuString = Space(100)
  lResult = GetMenuString(lhMenu, lMenuItem, sMenuString, 100, lcmdFlag)
  If lResult > 0 Then
    sMenuString = Trim(Left(sMenuString, lResult))
  Else
    Exit Function
  End If

' List of Items, where a Hint should be displayed
  If sMenuString = "&New Archive..." Then
  GetMenuHint = "Create New Archive"
  ElseIf sMenuString = "&Open Archive... " Then: GetMenuHint = "Open An Exsiting Archive"
  ElseIf sMenuString = "&Favorite Zip Folders..." Then: GetMenuHint = "Add Files from your Favorite Zip Folder"
  ElseIf sMenuString = "&Close Archive..." Then: GetMenuHint = "Close the Archive"
  ElseIf sMenuString = "Properties..." Then: GetMenuHint = "Open Properties menu"
  ElseIf sMenuString = "Create Shortcut" Then: GetMenuHint = "Create Archive Shortcut"
  End If

End Function

Function HiWord(ByVal lDWord As Long) As Long

  Dim i As Long
  Dim dblTemp As Double

' Generate unsigned 32-bit value, if param is negative
' To prevent getting the VB "Overflow"-Error, dont add more than &H7FFFFFFF at a time.
  dblTemp = lDWord
  If dblTemp < 0 Then
    dblTemp = &H7FFFFFFF
    dblTemp = dblTemp + &H7FFFFFFF
    dblTemp = (dblTemp + 2) - Abs(lDWord)
  End If
  
' No "Shift"-operator in VB. Must be divided by two, 16 times.
  For i = 0 To 15
    dblTemp = Fix(dblTemp / 2)
  Next i

  lDWord = dblTemp
  HiWord = lDWord

End Function

Function LoWord(ByVal lDWord As Long) As Long

  Dim dblTemp As Double
  
' Generate unsigned 32-bit value, if param is negative
' To prevent getting the VB "Overflow"-Error, dont add more than &H7FFFFFFF at a time.
  dblTemp = lDWord
  If dblTemp < 0 Then
    dblTemp = &H7FFFFFFF
    dblTemp = dblTemp + &H7FFFFFFF
    dblTemp = (dblTemp + 2) - Abs(lDWord)
  End If
  
' To prevent getting the VB "Overflow"-Error with the "AND"-operation, delete the signed bit first.
  If dblTemp > &H7FFFFFFF Then
    dblTemp = dblTemp - &H7FFFFFFF
    dblTemp = dblTemp - 1
  End If
  
  lDWord = dblTemp
  lDWord = lDWord And 65535

  LoWord = lDWord

End Function


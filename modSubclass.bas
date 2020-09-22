Attribute VB_Name = "modSubclass"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4&)
Public Const WM_SYSCOMMAND = &H112
Public Const WM_ACTIVATE As Long = &H6
Public Const WM_ACTIVATEAPP As Long = &H1C
Private Const HWND_TOP As Long = 0
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim PrevWndProc As Long
Dim mnuHandle As Long
Dim nmenuHandle As Long
Public Sub Init(hwnd As Long)
    mnuHandle = GetSystemMenu(hwnd, False)
    ' Add menu
    Call FormOnTop(frmMain)
    Call AppendMenu(mnuHandle, MF_SEPARATOR, 0, "")
    Call AppendMenu(mnuHandle, MF_STRING, &H200, "Here Is An Example Of System Menu")
    Call AppendMenu(mnuHandle, MF_SEPARATOR, 0, "")
    Call AppendMenu(mnuHandle, MF_STRING, &H201, "Enable Or Disable Menu ")
    Call AppendMenu(mnuHandle, MF_STRING, &H202, "Click Me")
    Call AppendMenu(mnuHandle, MF_STRING, &H203, "About")
    Call AppendMenu(mnuHandle, MF_STRING, &H204, "EXIT")
    
    PrevWndProc = SetWindowLong(mnuHandle, GWL_WNDPROC, AddressOf SubWndProc)
    
    PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub

Public Sub Terminate(hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)
End Sub

Public Function SubWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Result As Long
    
    If Msg = WM_SYSCOMMAND Then
        Select Case wParam
            
            Case &H200
                Result = GetMenuState(mnuHandle, &H200, MF_BYCOMMAND)
                If Result And MF_CHECKED Then ' Checking Checked Menu
                    Call CheckMenuItem(mnuHandle, &H200, MF_BYCOMMAND Or MF_UNCHECKED)
                    Call FormOnTop(frmMain)
                    MsgBox "Unchecked Menu", vbExclamation
                Else
                    Call CheckMenuItem(mnuHandle, &H200, MF_BYCOMMAND Or MF_CHECKED)
                    
                    Call FormOnTop(frmMain)
                    MsgBox "Checked Menu", vbCritical
                End If
            
            Case &H201
                EnableMenuItem mnuHandle, &H201, MF_DISABLED Or MF_GRAYED 'Disabled Menu
                MsgBox "Disabled Menu", vbCritical
            Case &H202 ' Delete Menu
                DeleteMenu mnuHandle, &H202, MF_BYCOMMAND Or MF_DELETE
                MsgBox "Again Click On Caption Bar You'll see that the menu is deleted", vbInformation
                
            Case &H203 ' Show frmAbout Form
                SetWindowPos frmAbout.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
                
            Case &H204 ' EXIT
                Unload frmMain
                
        End Select
    End If
    
    SubWndProc = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Sub FormOnTop(frm As Form)
    SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub



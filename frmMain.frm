VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Click On Caption Bar (Please Click On EXIT To Unload Me)"
   ClientHeight    =   5715
   ClientLeft      =   2145
   ClientTop       =   1590
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.Label Label4 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   300
      TabIndex        =   3
      Top             =   2130
      Width           =   5985
   End
   Begin VB.Label Label3 
      Caption         =   "Debasis Ghosh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4650
      TabIndex        =   2
      Top             =   4740
      Width           =   1605
   End
   Begin VB.Label Label2 
      Caption         =   "You can Reach Me At : -  debughosh@vsnl.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   675
      TabIndex        =   1
      Top             =   3675
      Width           =   5010
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":00C6
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Call RemSysMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Terminate(Me.hwnd)
End Sub
Sub RemSysMenu()
    Dim hsMenu As Long
    Dim Cnt As Long
    hsMenu = GetSystemMenu(Me.hwnd, False)

    If hsMenu Then
        
        Cnt = GetMenuItemCount(hsMenu)
        If Cnt Then
        
            RemoveMenu hsMenu, Cnt - 1, MF_BYPOSITION Or MF_REMOVE 'Remove Close Menu
            RemoveMenu hsMenu, Cnt - 2, MF_BYPOSITION Or MF_REMOVE 'Remove Separator
            RemoveMenu hsMenu, Cnt - 6, MF_BYPOSITION Or MF_REMOVE 'Remove Move Window
            RemoveMenu hsMenu, Cnt - 7, MF_BYPOSITION Or MF_REMOVE 'Remove Restore
            DrawMenuBar Me.hwnd

        End If
    End If
    Call Init(Me.hwnd)
End Sub



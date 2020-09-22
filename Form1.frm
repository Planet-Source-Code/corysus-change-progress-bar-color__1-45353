VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CHANGE PROGRESS BAR COLOR"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Progress Bar Scrolling Smooth"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Progress Bar Scrolling Standard"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin MSComctlLib.ProgressBar cProgres 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   37
         Scrolling       =   1
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   4080
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *************************************************************
' * PROJECT : CHANGE PROGRESS BAR COLOR
' * DESCRIPTION :
' * If you need cool progress bar without any OCX or User Control then this is code
' * JUST FOR YOU ! This code using two constants and one API Function
' * AUTOR : CORYSUS
' * LOACATION : BOSNIA & HERCEGOVINA
' *************************************************************





' API DECLARATION "USER 32"
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

' CONSTANTS
Const PBM_SETBARCOLOR = &H409
Const PBM_SETBKCOLOR = &H2001

Private iPos As Integer ' THIS IS FOR TIMER

Private Sub Option1_Click()
If Option1.Value = True Then
cProgres.Scrolling = ccScrollingStandard
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
cProgres.Scrolling = ccScrollingSmooth
End If
End Sub

Private Sub Timer1_Timer()

' MAIN FUNCTION
'------------------------------------------------------------------------------
    '// PLAYING WITH BAR COLOR [ RGB(xxx,xxx,xxx) ]
    PostMessage cProgres.hwnd, PBM_SETBARCOLOR, 0, RGB(255, 204, 51)
    '// PLAYING WITH BACK COLOR [ RGB(xxx,xxx,xxx) ]
    PostMessage cProgres.hwnd, PBM_SETBKCOLOR, 0, RGB(51, 102, 153)
'------------------------------------------------------------------------------

' PROGRESS FUNCTION
    cProgres.Value = cProgres.Value + iPos
    If cProgres.Value = cProgres.Max Then iPos = -1
    If cProgres.Value = cProgres.Min Then iPos = 1
    
End Sub

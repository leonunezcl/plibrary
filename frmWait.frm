VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   810
   ClientLeft      =   2940
   ClientTop       =   3315
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblGlosa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "FRMWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Sub Form_Activate()
    Refresh
End Sub

Private Sub Form_Load()

    CenterWindow hWnd
    
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = &HFF&
        .Color2 = 0
        .Draw Me
    End With
        
    DoEvents
    
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set FRMWait = Nothing
End Sub



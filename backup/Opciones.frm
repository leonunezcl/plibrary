VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   3090
   ClientLeft      =   2955
   ClientTop       =   3045
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Opciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOpci 
      Caption         =   "Configuración de internet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Index           =   2
      Left            =   1320
      TabIndex        =   7
      Top             =   285
      Width           =   4380
      Begin VB.ListBox lstPag 
         Height          =   1185
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   1095
         Width           =   4170
      End
      Begin VB.TextBox txtPagIni 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "http:/www.vbsoftware.cl"
         Top             =   510
         Width           =   4170
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         Caption         =   "Páginas visitadas (para eliminar marcar casilla)"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   840
         Width           =   3315
      End
      Begin VB.Label lblPag 
         AutoSize        =   -1  'True
         Caption         =   "Página de inicio:"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   285
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4965
      TabIndex        =   4
      Top             =   810
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   4965
      TabIndex        =   3
      Top             =   330
      Width           =   1215
   End
   Begin VB.Frame fraOpci 
      Caption         =   "Opciones miscelaneas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Index           =   1
      Left            =   435
      TabIndex        =   2
      Top             =   375
      Width           =   4380
      Begin VB.CheckBox chkOpci 
         Caption         =   "Respaldar libreria al salir"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   525
         Width           =   2265
      End
      Begin VB.CheckBox chkOpci 
         Caption         =   "Colorizar código"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   1515
      End
   End
   Begin MSComctlLib.TabStrip tabOpc 
      Height          =   3045
      Left            =   375
      TabIndex        =   1
      Top             =   0
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   5371
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Libreria"
            Object.ToolTipText     =   "Opciones de libreria"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Internet"
            Object.ToolTipText     =   "Configuración de internet"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3045
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
'graba las opciones de comparacion
Private Sub GrabarOpciones()

    Dim k As Integer
    Dim n As Integer
    
    glbColorizarCodigo = chkOpci(0).Value
    glbRespaldarLibreria = chkOpci(1).Value
    glbPaginaInicio = txtPagIni.Text
    
    'eliminar las seleccionadas
    n = 0
    For k = lstPag.SelCount To 0 Step -1
        If lstPag.Selected(k) Then
            frmMain.cboAddress.RemoveItem k
        Else
            n = n + 1
        End If
    Next k
    
    Call GrabaIni(C_INI, "opciones", "colorizar", chkOpci(0).Value)
    Call GrabaIni(C_INI, "opciones", "respaldar", chkOpci(1).Value)
    Call GrabaIni(C_INI, "web", "numero", n)
    
    'grabar las paginas marcadas
    For k = 0 To lstPag.ListCount - 1
        Call GrabaIni(C_INI, "web", "www" & k + 1, lstPag.List(k))
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call GrabarOpciones
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, Me.Caption)
    
    picDraw.Refresh
                    
    If glbColorizarCodigo Then
        chkOpci(0).Value = 1
    Else
        chkOpci(0).Value = 0
    End If
    
    If glbRespaldarLibreria Then
        chkOpci(1).Value = 1
    Else
        chkOpci(1).Value = 0
    End If
    
    fraOpci(1).ZOrder 0
    fraOpci(2).Left = fraOpci(1).Left
    fraOpci(2).Top = fraOpci(1).Top
    fraOpci(2).Height = fraOpci(1).Height
    fraOpci(2).Width = fraOpci(1).Width
    
    txtPagIni.Text = glbPaginaInicio
    
    For k = 0 To frmMain.cboAddress.ListCount - 1
        lstPag.AddItem frmMain.cboAddress.List(k)
    Next k
    
    Call Hourglass(hWnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmOpciones = Nothing
End Sub


Private Sub tabOpc_Click()

    fraOpci(tabOpc.SelectedItem.Index).ZOrder 0
    
End Sub



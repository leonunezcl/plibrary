VERSION 5.00
Begin VB.Form frmItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Categoría"
   ClientHeight    =   5100
   ClientLeft      =   3645
   ClientTop       =   3975
   ClientWidth     =   5760
   Icon            =   "Frmnuevo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   705
      TabIndex        =   2
      Top             =   75
      Width           =   4950
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   4665
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Crear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   1395
      TabIndex        =   0
      Top             =   4665
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ID_CATEGORIA As Integer
Public Accion As Integer
Private Sub cmd_Click(Index As Integer)

    Dim Carpeta As String
        
    If Index = 0 Then
        Carpeta = Trim$(txt.Text)
        If Carpeta <> "" Then
            If CreaCarpeta(Carpeta, ID_CATEGORIA) Then
                MsgBox "Carpeta creada con éxito.", vbInformation
                frmMain.tvTreeView.Nodes.Add "root", tvwChild, "k" & ID_CATEGORIA, Carpeta, C_ICONO_CLOSE, C_ICONO_CLOSE
                frmMain.tvTreeView.Nodes("k" & ID_CATEGORIA).EnsureVisible
            Else
                MsgBox "Categoria ya existe.", vbCritical
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call CenterWindow(hWnd)
        
    If Accion = 1 Then
        Me.Caption = "Nuevo item"
    Else
        Me.Caption = "Modificar item"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCategoria = Nothing
    
End Sub



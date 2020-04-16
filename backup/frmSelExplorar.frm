VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelExplorar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione archivos a comparar"
   ClientHeight    =   4875
   ClientLeft      =   1935
   ClientTop       =   2115
   ClientWidth     =   6090
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelExplorar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4875
      Left            =   0
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   5
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4815
      TabIndex        =   4
      Top             =   765
      Width           =   1200
   End
   Begin VB.OptionButton opt 
      Caption         =   "No Seleccionar Todos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2460
      TabIndex        =   3
      Top             =   60
      Width           =   2220
   End
   Begin VB.OptionButton opt 
      Caption         =   "Seleccionar Todos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   2
      Top             =   60
      Width           =   1905
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   330
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   4485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":04E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":06C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":089E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":0BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":0DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":10F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelExplorar.frx":12D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treeProyecto 
      Height          =   4425
      Left            =   420
      TabIndex        =   0
      Top             =   330
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   7805
      _Version        =   393217
      Indentation     =   882
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
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
End
Attribute VB_Name = "frmSelExplorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Pasar As Boolean
Private mGradient As New clsGradient
Public Origen As Boolean

'carga archivos del proyecto para poder seleccionarlos
Private Sub CargaArchivosDelProyecto()

    Dim k As Integer
    Dim bForm As Boolean
    Dim bModule As Boolean
    Dim bControl As Boolean
    Dim bClase As Boolean
    Dim bPags As Boolean
    Dim bDocRel As Boolean
    Dim bDesigner As Boolean
    
    treeProyecto.Nodes.Add(, , "PRO", Proyecto.Nombre & " (" & Proyecto.Archivo & ")", C_ICONO_PROYECTO).EnsureVisible
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            If Not bForm Then
                Call treeProyecto.Nodes.Add("PRO", tvwChild, "FRM", "Formularios", C_ICONO_OPEN).EnsureVisible
                bForm = True
            End If
            Call treeProyecto.Nodes.Add("FRM", tvwChild, Proyecto.aArchivos(k).KeyNodeFrm, Proyecto.aArchivos(k).Nombre, C_ICONO_FORM, C_ICONO_FORM)
            
            treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeFrm).Checked = True
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            If Not bModule Then
                Call treeProyecto.Nodes.Add("PRO", tvwChild, "BAS", "Módulos", C_ICONO_OPEN).EnsureVisible
                bModule = True
            End If
            Call treeProyecto.Nodes.Add("BAS", tvwChild, Proyecto.aArchivos(k).KeyNodeBas, Proyecto.aArchivos(k).Nombre, C_ICONO_BAS, C_ICONO_BAS)
            treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeBas).Checked = True
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            If Not bControl Then
                Call treeProyecto.Nodes.Add("PRO", tvwChild, "CTL", "Controles de Usuario", C_ICONO_OPEN).EnsureVisible
                bControl = True
            End If
            Call treeProyecto.Nodes.Add("CTL", tvwChild, Proyecto.aArchivos(k).KeyNodeKtl, Proyecto.aArchivos(k).Nombre, C_ICONO_CONTROL, C_ICONO_CONTROL)
            treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeKtl).Checked = True
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            If Not bClase Then
                Call treeProyecto.Nodes.Add("PRO", tvwChild, "CLS", "Módulos de Clase", C_ICONO_OPEN).EnsureVisible
                bClase = True
            End If
            Call treeProyecto.Nodes.Add("CLS", tvwChild, Proyecto.aArchivos(k).KeyNodeCls, Proyecto.aArchivos(k).Nombre, C_ICONO_CLS, C_ICONO_CLS)
            treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeCls).Checked = True
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
            If Not bPags Then
                Call treeProyecto.Nodes.Add("PRO", tvwChild, "PAG", "Páginas de Propiedades", C_ICONO_OPEN).EnsureVisible
                bPags = True
            End If
            Call treeProyecto.Nodes.Add("PAG", tvwChild, Proyecto.aArchivos(k).KeyNodePag, Proyecto.aArchivos(k).Nombre, C_ICONO_PAGINA, C_ICONO_PAGINA)
            treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodePag).Checked = True
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
            If Not bDesigner Then
                Call treeProyecto.Nodes.Add("PRO", tvwChild, "DSR", "Diseñadores", C_ICONO_OPEN).EnsureVisible
                bDesigner = True
            End If
            Call treeProyecto.Nodes.Add("DSR", tvwChild, Proyecto.aArchivos(k).KeyNodeDsr, Proyecto.aArchivos(k).Nombre, 10, 10)
            treeProyecto.Nodes(Proyecto.aArchivos(k).KeyNodeDsr).Checked = True
        End If
    Next k
    
End Sub

'marca los nodos del arbol segun seleccion
Private Sub MarcaNodos(ByVal Estado As Boolean)

    Dim k As Integer
    
    For k = 1 To treeProyecto.Nodes.Count
        treeProyecto.Nodes(k).Checked = Estado
    Next k
    
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 1 Then
        glbSelArchivos = False
        Call MarcaNodos(False)
    End If
    
    Unload Me
        
End Sub

Private Sub Form_Activate()
    opt(1).Value = True
    DoEvents
    opt(0).Value = True
End Sub

Private Sub Form_Load()
    
    CenterWindow hWnd
    
    If Origen Then
        LSet Proyecto = ProyectoO
    Else
        LSet Proyecto = ProyectoD
    End If
    
    Call CargaArchivosDelProyecto
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, Me.Caption)
    
    picDraw.Refresh
    
    Pasar = False
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim k As Integer
    Dim j As Integer
    Dim Found As Boolean
    Dim total As Integer
    
    glbSelArchivos = False
    
    total = UBound(Proyecto.aArchivos)
    
    For j = 1 To treeProyecto.Nodes.Count
        Found = False
            
        If Len(treeProyecto.Nodes(j).Key) > 3 Then
            'recorrer archivos del proyecto
            For k = 1 To total
                If Proyecto.aArchivos(k).KeyNodeFrm = treeProyecto.Nodes(j).Key Then
                    Found = True
                ElseIf Proyecto.aArchivos(k).KeyNodeBas = treeProyecto.Nodes(j).Key Then
                    Found = True
                ElseIf Proyecto.aArchivos(k).KeyNodeCls = treeProyecto.Nodes(j).Key Then
                    Found = True
                ElseIf Proyecto.aArchivos(k).KeyNodeKtl = treeProyecto.Nodes(j).Key Then
                    Found = True
                ElseIf Proyecto.aArchivos(k).KeyNodePag = treeProyecto.Nodes(j).Key Then
                    Found = True
                End If
                
                'archivo marcado ?
                If Found Then
                    Proyecto.aArchivos(k).Explorar = treeProyecto.Nodes(j).Checked
                    If Proyecto.aArchivos(k).Explorar Then
                        glbSelArchivos = True
                    End If
                    Exit For
                End If
            Next k
        End If
    Next j
    
    If Origen Then
        LSet ProyectoO = Proyecto
    Else
        LSet ProyectoD = Proyecto
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmSelExplorar = Nothing
    
End Sub


Private Sub opt_Click(Index As Integer)

    If Pasar Then Exit Sub
    
    If Index = 0 Then
        glbSelArchivos = True
    Else
        glbSelArchivos = False
    End If
    
    Call MarcaNodos(glbSelArchivos)
    
End Sub

Private Sub treeProyecto_Collapse(ByVal Node As MSComctlLib.Node)

    Select Case Node.Text
        Case "Referencias", "Componentes", "Formularios", "Módulos", "Módulos de Clase"
            Node.SelectedImage = 8
            Node.Image = 8
        Case "Controles de Usuario", "Páginas de Propiedades", "Documentos Relacionados"
            Node.SelectedImage = 8
            Node.Image = 8
    End Select
    
End Sub

Private Sub treeProyecto_Expand(ByVal Node As MSComctlLib.Node)

    Select Case Node.Text
        Case "Referencias", "Componentes", "Formularios", "Módulos", "Módulos de Clase"
            Node.Image = 9
            Node.SelectedImage = 9
        Case "Controles de Usuario", "Páginas de Propiedades", "Documentos Relacionados"
            Node.Image = 9
            Node.SelectedImage = 9
    End Select
    
End Sub


Private Sub treeProyecto_NodeCheck(ByVal Node As MSComctlLib.Node)

    Dim k As Integer
    Dim j As Integer
    Dim Found As Boolean
    Dim total As Integer
    
    total = UBound(Proyecto.aArchivos)
    
    'todo el proyecto
    If Node.Key = "PRO" Then
        For j = 1 To treeProyecto.Nodes.Count
            treeProyecto.Nodes(j).Checked = Node.Checked
        Next j
        
        glbSelArchivos = Node.Checked
        
        Pasar = True
        If glbSelArchivos Then
            opt(0).Value = 1
        Else
            opt(1).Value = 1
        End If
        Pasar = False
        Exit Sub
    End If
    
    Found = False
    For k = 1 To total
        If Proyecto.aArchivos(k).KeyNodeFrm = Node.Key Then
            Found = True
        ElseIf Proyecto.aArchivos(k).KeyNodeBas = Node.Key Then
            Found = True
        ElseIf Proyecto.aArchivos(k).KeyNodeCls = Node.Key Then
            Found = True
        ElseIf Proyecto.aArchivos(k).KeyNodeKtl = Node.Key Then
            Found = True
        ElseIf Proyecto.aArchivos(k).KeyNodePag = Node.Key Then
            Found = True
        End If
        If Found Then Exit For
    Next k
    
    'si no es un archivo entonces es un conjunto de archivos
    If Not Found Then
        For j = 1 To treeProyecto.Nodes.Count
            If Left$(treeProyecto.Nodes(j).Key, 3) = Left$(Node.Key, 3) Then
                If Len(treeProyecto.Nodes(j).Key) <> 3 Then
                    treeProyecto.Nodes(j).Checked = Node.Checked
                End If
            End If
        Next j
    End If
    
End Sub


VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Project Library"
   ClientHeight    =   6150
   ClientLeft      =   1350
   ClientTop       =   3405
   ClientWidth     =   12060
   HelpContextID   =   10
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgCatItem 
      Left            =   9915
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0952
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1906
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2272
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2596
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTb 
      Left            =   1260
      Top             =   3495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   50
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F02
            Key             =   ""
            Object.Tag             =   "&Eliminar item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37DE
            Key             =   ""
            Object.Tag             =   "&Agregar item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40BA
            Key             =   ""
            Object.Tag             =   "&Módulo .bas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4216
            Key             =   ""
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":466A
            Key             =   ""
            Object.Tag             =   "Mó&dulo .cls"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47C6
            Key             =   ""
            Object.Tag             =   "&Indice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AE2
            Key             =   ""
            Object.Tag             =   "&Ir a VBSoftware"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DFE
            Key             =   ""
            Object.Tag             =   "&Imprimir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F5A
            Key             =   ""
            Object.Tag             =   "&Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C36
            Key             =   ""
            Object.Tag             =   "&Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D4A
            Key             =   ""
            Object.Tag             =   "C&ortar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F72
            Key             =   ""
            Object.Tag             =   "&Pegar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7086
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":719A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7726
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7882
            Key             =   ""
            Object.Tag             =   "Exportar a rt&f"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E5A
            Key             =   ""
            Object.Tag             =   "Ac&tualizar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FB6
            Key             =   ""
            Object.Tag             =   "Borr&ar"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80CE
            Key             =   ""
            Object.Tag             =   "&Respaldar libreria a .zip"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":89AA
            Key             =   ""
            Object.Tag             =   "&Importar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E02
            Key             =   ""
            Object.Tag             =   "&Exportar"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":925A
            Key             =   ""
            Object.Tag             =   "&Formulario"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93B6
            Key             =   ""
            Object.Tag             =   "&Control"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9512
            Key             =   ""
            Object.Tag             =   "&Página de propiedades"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":966E
            Key             =   ""
            Object.Tag             =   "Pro&yecto"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":97CA
            Key             =   ""
            Object.Tag             =   "Exportar a &texto "
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AE6
            Key             =   ""
            Object.Tag             =   "Exportar a &html"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E02
            Key             =   ""
            Object.Tag             =   "B&uscar Siguiente"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F5E
            Key             =   ""
            Object.Tag             =   "&Reemplazar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A0BA
            Key             =   ""
            Object.Tag             =   "&Buscar código ..."
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3DE
            Key             =   ""
            Object.Tag             =   "&Ver código fuente"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A5C2
            Key             =   ""
            Object.Tag             =   "&Modificar item"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A71E
            Key             =   ""
            Object.Tag             =   "&Seleccionar todos los itemes"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A87E
            Key             =   ""
            Object.Tag             =   "&Quitar selección de itemes"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9DE
            Key             =   ""
            Object.Tag             =   "&Invertir selección"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AC9E
            Key             =   ""
            Object.Tag             =   "&Tip del dia ..."
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AFC2
            Key             =   ""
            Object.Tag             =   "&Opciones del libreria"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B416
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B732
            Key             =   ""
            Object.Tag             =   "&Email a VBSoftware"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B88E
            Key             =   ""
            Object.Tag             =   "A&gregar a bookmark"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BBB2
            Key             =   ""
            Object.Tag             =   "&Quitar de bookmark"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C516
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Código"
      Height          =   2775
      Index           =   4
      Left            =   660
      TabIndex        =   21
      Top             =   3060
      Visible         =   0   'False
      Width           =   4035
      Begin MSComctlLib.ListView lvwBookmark 
         Height          =   1380
         Left            =   105
         TabIndex        =   22
         Top             =   510
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   2434
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgTb"
         SmallIcons      =   "imgTb"
         ColHdrIcons     =   "imgTb"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción de código"
            Object.Width           =   12347
         EndProperty
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4065
      Top             =   1260
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Itemes"
      Height          =   2955
      Index           =   3
      Left            =   2250
      TabIndex        =   15
      Top             =   2460
      Visible         =   0   'False
      Width           =   4410
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   855
         TabIndex        =   17
         Top             =   480
         Width           =   3795
      End
      Begin MSComctlLib.Toolbar tbToolBar 
         Height          =   450
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         Style           =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Object.ToolTipText     =   "Atrás"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Forward"
               Object.ToolTipText     =   "Adelante"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stop"
               Object.ToolTipText     =   "Detener"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Actualizar"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Home"
               Object.ToolTipText     =   "Inicio"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Search"
               Object.ToolTipText     =   "Búsqueda"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin SHDocVwCtl.WebBrowser brwWebBrowser 
         Height          =   2025
         Left            =   45
         TabIndex        =   19
         Top             =   855
         Width           =   2880
         ExtentX         =   5080
         ExtentY         =   3572
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   -1  'True
         NoClientEdge    =   -1  'True
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Dirección:"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Tag             =   "&Dirección:"
         Top             =   510
         Width           =   795
      End
   End
   Begin VB.Timer tmrCodigo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3270
      Top             =   1065
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3585
      Top             =   2010
   End
   Begin VB.FileListBox filCode 
      Height          =   480
      Left            =   900
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4095
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picCat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   480
      ScaleHeight     =   510
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   735
      Width           =   2385
      Begin VB.Label lblSeccion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Categorias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   525
         TabIndex        =   12
         Top             =   150
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   15
         Picture         =   "frmMain.frx":C83A
         Top             =   15
         Width           =   480
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Itemes"
      Height          =   2955
      Index           =   1
      Left            =   7365
      TabIndex        =   7
      Top             =   2745
      Visible         =   0   'False
      Width           =   4410
      Begin VB.CommandButton cmdNav 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   450
         TabIndex        =   29
         ToolTipText     =   "Ir a la primera página"
         Top             =   345
         Width           =   435
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   915
         TabIndex        =   28
         ToolTipText     =   "Retroceder una página"
         Top             =   345
         Width           =   435
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2175
         TabIndex        =   27
         ToolTipText     =   "Avanzar una página"
         Top             =   345
         Width           =   435
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   26
         ToolTipText     =   "Ir a la última página"
         Top             =   345
         Width           =   435
      End
      Begin VB.TextBox txtNav 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   25
         ToolTipText     =   "Ir a página "
         Top             =   345
         Width           =   750
      End
      Begin MSComctlLib.ListView lvwItemes 
         Height          =   1380
         Left            =   450
         TabIndex        =   8
         Top             =   645
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   2434
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgTb"
         SmallIcons      =   "imgTb"
         ColHdrIcons     =   "imgTb"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción de código"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.Image imgCat 
         Height          =   300
         Left            =   0
         Picture         =   "frmMain.frx":CC7C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblDesCat 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Categorias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   405
         TabIndex        =   23
         Top             =   0
         Width           =   1110
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4530
      Left            =   3180
      TabIndex        =   5
      Top             =   825
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7990
      HotTracking     =   -1  'True
      ImageList       =   "imgTb"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Itemes"
            Object.ToolTipText     =   "Codigo fuente asociado a la seccion"
            ImageVarType    =   2
            ImageIndex      =   17
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Código &fuente"
            Object.ToolTipText     =   "Código fuente Visual Basic del item seleccionado"
            ImageVarType    =   2
            ImageIndex      =   16
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar en internet"
            Object.ToolTipText     =   "Buscar código en internet"
            ImageVarType    =   2
            ImageIndex      =   41
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&BookMark"
            Object.ToolTipText     =   "Ficha de almacén de selecciones"
            ImageVarType    =   2
            ImageIndex      =   44
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbStatus 
      Height          =   285
      Left            =   1215
      TabIndex        =   4
      Top             =   5490
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Splitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   4455
      MouseIcon       =   "frmMain.frx":D0BE
      MousePointer    =   99  'Custom
      ScaleHeight     =   363
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   3
      Tag             =   "0"
      Top             =   315
      Width           =   45
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTb"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEliItem"
            Object.ToolTipText     =   "Eliminar item"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAddItem"
            Object.ToolTipText     =   "Agregar item"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSave"
            Object.ToolTipText     =   "Respaldar libreria de código"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdActualizar"
            Object.ToolTipText     =   "Actualizar información de libreria"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCortar"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCopiar"
            Object.ToolTipText     =   "Copiar codigo del item"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPegar"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdBuscar"
            Object.ToolTipText     =   "Buscar codigo en libreria"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdImprimir"
            Object.ToolTipText     =   "Imprimir codigo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTexto"
            Object.ToolTipText     =   "Exportar codigo a texto"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRtf"
            Object.ToolTipText     =   "Exportar codigo a rtf"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdHtml"
            Object.ToolTipText     =   "Exportar codigo a html"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdBas"
            Object.ToolTipText     =   "Importar módulo .bas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCls"
            Object.ToolTipText     =   "Importar módulo .cls"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdForm"
            Object.ToolTipText     =   "Importar formulario"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdCtl"
            Object.ToolTipText     =   "Importar control de usuario"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPag"
            Object.ToolTipText     =   "Importar página de propiedades"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdVbp"
            Object.ToolTipText     =   "Importar proyecto"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdWeb"
            Object.ToolTipText     =   "Ir al sitio web de vbsoftware"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAyuda"
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Salir de la aplicacion"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4725
      Left            =   0
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   705
      Width           =   360
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   5835
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCat 
      Height          =   2025
      Left            =   495
      TabIndex        =   6
      Top             =   1350
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgCatItem"
      SmallIcons      =   "imgCatItem"
      ColHdrIcons     =   "imgCatItem"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Categorias"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2190
      Top             =   4485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D210
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D4F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DAB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E07A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Código"
      Height          =   2775
      Index           =   2
      Left            =   9135
      TabIndex        =   9
      Top             =   2025
      Visible         =   0   'False
      Width           =   4035
      Begin MSComctlLib.Toolbar tbCodigo 
         Height          =   330
         Left            =   600
         TabIndex        =   24
         Top             =   330
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "imgTb"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmdFecha"
               Object.ToolTipText     =   "Insertar fecha al código"
               ImageIndex      =   48
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmdHora"
               Object.ToolTipText     =   "Insertar hora al código"
               ImageIndex      =   49
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmdComentario"
               Object.ToolTipText     =   "Insertar comentarios"
               ImageIndex      =   50
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "cmdCabezera"
                     Text            =   "Cabezera"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "cmdFuncion"
                     Text            =   "Función"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "cmdSub"
                     Text            =   "Sub"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "cmdGet"
                     Text            =   "Property Get"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "cmdLet"
                     Text            =   "Property Let"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "cmdSet"
                     Text            =   "Property Set"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbCodigo 
         Height          =   1515
         Left            =   750
         TabIndex        =   10
         Top             =   690
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   2672
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":E35C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox PicLines 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   120
         ScaleHeight     =   1860
         ScaleWidth      =   405
         TabIndex        =   14
         Top             =   180
         Width           =   400
      End
      Begin VB.Label lblDescripItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   585
         TabIndex        =   20
         Top             =   0
         Width           =   1110
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      HelpContextID   =   20
      Begin VB.Menu mnuArchivo_Importar 
         Caption         =   "|Importar archivos visual basic|&Importar"
         Begin VB.Menu mnuArchivo_Importar_Bas 
            Caption         =   "|Importar módulo .bas|&Módulo .bas"
         End
         Begin VB.Menu mnuArchivo_Importar_Cls 
            Caption         =   "|Importar módulo .cls|Mó&dulo .cls"
         End
         Begin VB.Menu mnuArchivo_Importar_Formulario 
            Caption         =   "|Importar formulario|&Formulario"
         End
         Begin VB.Menu mnuArchivo_Importar_Control 
            Caption         =   "|Importar control de usuario|&Control de Usuario"
         End
         Begin VB.Menu mnuArchivo_Importar_Pagina 
            Caption         =   "|Importar página de propiedades|&Página de Propiedades"
         End
         Begin VB.Menu mnuArchivo_Importar_Proyecto 
            Caption         =   "|Importar proyecto visual basic|Pro&yecto Visual Basic"
         End
      End
      Begin VB.Menu mnuArchivo_Exportar 
         Caption         =   "|Exportar archivos|&Exportar"
         Begin VB.Menu mnuArchivo_Exportar_Bas 
            Caption         =   "|Exportar a módulo .bas|&Módulo .bas"
         End
         Begin VB.Menu mnuArchivo_Exportar_Cls 
            Caption         =   "|Exportar módulo .cls|Mó&dulo .cls"
         End
      End
      Begin VB.Menu mnuArchivo_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_ConfImpresora 
         Caption         =   "|Configurar impresora|&Configurar impresora"
      End
      Begin VB.Menu mnuArchivo_Impresora 
         Caption         =   "|Imprimir diferencias en impresora|&Imprimir"
      End
      Begin VB.Menu mnuArchivo_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Salir 
         Caption         =   "|Salir de la aplicacion|&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      HelpContextID   =   30
      Begin VB.Menu mnuEdicion_Copiar 
         Caption         =   "|Copiar contenido al portapapeles|&Copiar"
      End
      Begin VB.Menu mnuEdicion_Cortar 
         Caption         =   "|Cortar selección al portapapeles|C&ortar"
      End
      Begin VB.Menu mnuEdicion_Pegar 
         Caption         =   "|Pegar contenido del portapapeles|&Pegar"
      End
      Begin VB.Menu mnuEdicion_Borrar 
         Caption         =   "|Borrar texto seleccionado o todo |Borr&ar"
      End
      Begin VB.Menu mnuEdicion_SelTodo 
         Caption         =   "|Seleccionar todo el código|&Seleccionar todo ..."
      End
      Begin VB.Menu mnuEdicion_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_Buscar 
         Caption         =   "|Buscar en código|&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdicion_BusSiguiente 
         Caption         =   "|Buscar siguiente ocurrencia de busqueda en texto|B&uscar Siguiente"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdicion_Reemplazar 
         Caption         =   "|Reemplazar texto en código|&Reemplazar"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEdicion_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_ExpTexto 
         Caption         =   "|Exportar código a archivo de texto|Exportar a &texto "
      End
      Begin VB.Menu mnuEdicion_ExpRtf 
         Caption         =   "|Exportar código a formato enriquecido|Exportar a rt&f"
      End
      Begin VB.Menu mnuEdicion_ExpHtml 
         Caption         =   "|Exportar código a formato hypertexto|Exportar a &html"
      End
   End
   Begin VB.Menu mnuLibreria 
      Caption         =   "&Libreria"
      HelpContextID   =   40
      Begin VB.Menu mnuLibreria_AgregarItem 
         Caption         =   "|Agregar código a libreria|&Agregar item"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLibreria_ModificarItem 
         Caption         =   "|Modificar código en libreria|&Modificar item"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuRepositorio_EliminarItem 
         Caption         =   "|Eliminar código de libreria|&Eliminar item"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLibreria_Buscar 
         Caption         =   "|Buscar código en libreria|&Buscar código ..."
      End
      Begin VB.Menu mnuLibreria_Actualizar 
         Caption         =   "|Actualizar itemes de categoria|Ac&tualizar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRepositorio_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibreria_VerCodigo 
         Caption         =   "|Ver el código del item seleccionado|&Ver código fuente"
      End
      Begin VB.Menu mnuRepositorio_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibreria_SelTodos 
         Caption         =   "|Seleccionar todo el código|&Seleccionar todos los itemes"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuLibreria_Invertir 
         Caption         =   "|Invertir seleccion de itemes de código|&Invertir selección"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuLibreria_Quitar 
         Caption         =   "|Eliminar seleccion de itemes|&Quitar selección de itemes"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuRepositorio_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibreria_AgregarBookmark 
         Caption         =   "|Agregar a selección de código a bookmark|A&gregar a bookmark"
      End
      Begin VB.Menu mnuLibreria_QuitarBookmark 
         Caption         =   "|Eliminar código seleccionado de bookmark|&Quitar de bookmark"
      End
      Begin VB.Menu mnuRepositorio_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Respaldo 
         Caption         =   "|Respaldar la libreria a archivo .zip|&Respaldar libreria a .zip"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      HelpContextID   =   60
      Begin VB.Menu mnuOpciones_OpcEditor 
         Caption         =   "|Configurar opciones de Project Library|&Opciones del libreria"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyuda_Indice 
         Caption         =   "|Indice de la ayuda de Project Library|&Indice"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAyuda_Buscar 
         Caption         =   "|Buscar en archivo de ayuda|B&usqueda ..."
      End
      Begin VB.Menu mnuAyuda_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_WebSite 
         Caption         =   "|Ir al sitio WWW de VBSoftware|&Ir a VBSoftware"
      End
      Begin VB.Menu mnuAyuda_Email 
         Caption         =   "|Escribe un email a VBSoftware|&Email a VBSoftware"
      End
      Begin VB.Menu mnuAyuda_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_Tips 
         Caption         =   "|Mostrar tips del dia|&Tip del dia ..."
      End
      Begin VB.Menu mnuAyuda_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_AcercaDe 
         Caption         =   "|Información de Copyright y del Autor|Acerca &de ..."
      End
   End
   Begin VB.Menu mnuBookmark 
      Caption         =   "Bookmark"
      Visible         =   0   'False
      Begin VB.Menu mnuBookmark_SelTodos 
         Caption         =   "|Seleccionar todo el código|&Seleccionar todos los itemes"
      End
      Begin VB.Menu mnuBookmark_Invertir 
         Caption         =   "|Invertir seleccion de itemes de código|&Invertir selección"
      End
      Begin VB.Menu mnuBookmark_Quitar 
         Caption         =   "|Eliminar seleccion de itemes|&Quitar selección de itemes"
      End
      Begin VB.Menu mnuBookmark_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookmark_ExportarBas 
         Caption         =   "|Exportar a módulo .bas|&Módulo .bas"
      End
      Begin VB.Menu mnuBookmark_ExportarCls 
         Caption         =   "|Exportar módulo .cls|Mó&dulo .cls"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private cc As New GCommonDialog
Private clsXmenu As New CXtremeMenu
Private WithEvents MyHelpCallBack As HelpCallBack
Attribute MyHelpCallBack.VB_VarHelpID = -1
Private WithEvents m_cZ As cZip
Attribute m_cZ.VB_VarHelpID = -1
Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1
Private Itmx As ListItem
Private pos As Long
Private contador As Integer
Private LineCountChange As Integer           '// This is used to determin if we need _
                                             to redraw the numbers
Private FirstLine As Long                    '// Dim the First visible line
Private FirstLineNow As Long
Private fTip As Boolean
Private Const MIN_VERT_BUFFER As Integer = 20
Private Const MIN_HORZ_BUFFER As Integer = 13
Private Const CURSOR_DEDUCT As Integer = 10
Private Const SPLT_WDTH As Integer = 4
Private Const SPLT_HEIGHT As Integer = 4
Private Const CTRL_OFFSET As Integer = 28
Private fInitiateDrag As Boolean
Private Cargando As Boolean
Private UltimaCat As Integer
Private UltimoIte As Integer
Private StartingAddress As String
Private mbDontNavigateNow As Boolean
Private intPageCount As Integer
Private intCount As Integer
Private intRecord As Integer
Private intPage As Integer

'abre archivo .bas o .cls
Private Sub AbrirArchivo(ByVal tipo As String)

    Dim Glosa As String
    Dim Archivo As String
    
    If tipo = "bas" Then
        Glosa = "Módulos (*.bas)|*.bas|"
    ElseIf tipo = "cls" Then
        Glosa = "Módulos de Clase (*.cls)|*.cls|"
    ElseIf tipo = "frm" Then
        Glosa = "Formularios (*.frm)|*.frm|"
    ElseIf tipo = "ctl" Then
        Glosa = "Control de Usuario (*.ctl)|*.ctl|"
    ElseIf tipo = "pag" Then
        Glosa = "Páginas de Propiedades (*.pag)|*.pag|"
    ElseIf tipo = "vbp" Then
        Glosa = "Proyectos Visual Basic (*.vbp)|*.vbp|"
    End If
            
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If cc.VBGetOpenFileName(Archivo, , , , , , Glosa, , App.Path, "Abrir archivo ...", tipo) Then
        Timer1.Enabled = False
        Call GetProjectDetails(Archivo)
        Timer1.Enabled = True
    End If
    
End Sub

'regresa el archivo de codigo
Private Function ArchivoDeCodigo(ByVal Seccion As Integer, ByVal item As Integer) As String

    Dim ret As String
        
    glbSQL = "select linea from codigo where "
    glbSQL = glbSQL & "     id = " & Seccion
    glbSQL = glbSQL & " and item = " & item
    glbSQL = glbSQL & " and correlativo = 0"
    
    glbRecordset.Open glbSQL, glbConnection
    
    If Not glbRecordset.EOF Then
        ret = App.Path & Trim$(glbRecordset!Linea)
    End If
    
    glbRecordset.Close
    
    ArchivoDeCodigo = ret
    
End Function

'cuenta el codigo en las categorias
'contar codigo
Public Sub ContarCodigo()

    Dim ret As Long
    Dim k As Integer
    Dim total As Integer
    
    'ciclar x categorias
    For k = 1 To UBound(Arr_Categorias)
        glbSQL = "select count(*) as cuenta "
        glbSQL = glbSQL & "from "
        glbSQL = glbSQL & "itemes "
        glbSQL = glbSQL & "where "
        glbSQL = glbSQL & "id = " & k
        
        glbRecordset.Open glbSQL, glbConnection
        
        If Not glbRecordset.EOF Then
            If Not IsNull(glbRecordset!cuenta) Then
                total = glbRecordset!cuenta
            Else
                total = 1
            End If
        Else
            total = 1
        End If
        
        ret = ret + total
        
        lvwCat.ListItems(k).Text = Arr_Categorias(k).Descripcion & " (" & total & ")"
        
        glbRecordset.Close
    Next k
    
    lblSeccion.Caption = "Categorias (" & ret & ")"
        
End Sub

'contar itemes
Public Function ContarItemes(ByVal Seccion As Integer) As Long

    Dim ret As Long
    
    glbSQL = "select count(*) as cuenta from itemes "
    glbSQL = glbSQL & "where "
    glbSQL = glbSQL & "id = " & Seccion
    
    glbRecordset.Open glbSQL, glbConnection
    
    If Not glbRecordset.EOF Then
        If Not IsNull(glbRecordset!cuenta) Then
            ret = glbRecordset!cuenta
        Else
            ret = 0
        End If
    Else
        ret = 0
    End If
    
    glbRecordset.Close
    
    ContarItemes = ret
    
End Function

'exportar código
Private Function Exportar(ByVal Msg As String, ByVal Glosa As String, ByVal Ext As String) As Boolean

    On Local Error GoTo ErrorExportar
        
    Dim ret As Boolean
    Dim Archivo As String
    Dim ArchivoZip As String
    Dim ArchivoCode As String
    Dim k As Integer
    Dim Seccion As Integer
    Dim item As Integer
    Dim Itmx As ListItem
    Dim nFreeFile As Integer
    Dim nFreeFile2 As Integer
    Dim Linea As String
    Dim HayCodigo As Boolean
    
    ret = False
    
    'hay itemes seleccionados ?
    If lvwBookmark.ListItems.Count > 0 Then
        'verificar si hay rutinas seleccionadas
        For k = 1 To lvwBookmark.ListItems.Count
            'exportar solo lo seleccionado
            If lvwBookmark.ListItems(k).Checked Then
                HayCodigo = True
                Exit For
            End If
        Next k
        
        'hay codigo seleccionado ?
        If HayCodigo Then
            If Confirma(Msg) = vbYes Then
                If cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar como ...", Ext) Then
                
                    Call InhabilitaToolbar(False)
                    Call Hourglass(hWnd, True)
                    Call MyMsg("Generando archivo. Espere ...")
                    nFreeFile = FreeFile
                    
                    'abrir archivo nuevo ...
                    Open Archivo For Output As #nFreeFile
                        'para abrir archivo de codigo
                        nFreeFile2 = FreeFile
                        'ciclar x los itemes seleccionados
                        For k = 1 To lvwBookmark.ListItems.Count
                            'exportar solo lo seleccionado
                            If lvwBookmark.ListItems(k).Checked Then
                                Set Itmx = lvwBookmark.ListItems(k)
                                
                                Seccion = Left$(Itmx.Key, InStr(1, Itmx.Key, "-") - 1)
                                item = Val(Mid$(Itmx.Key, InStr(1, Itmx.Key, "-") + 1))
                                
                                'seleccionar codigo
                                glbSQL = "select correlativo , linea from codigo where id = " & Seccion
                                glbSQL = glbSQL & " and item = " & item
                                
                                glbRecordset.Open glbSQL, glbConnection, adOpenForwardOnly
                            
                                'ciclar x el código
                                Do While Not glbRecordset.EOF
                                    'el correlativo cero es el archivo
                                    If glbRecordset!correlativo = 0 Then
                                        'verificar si existe zip
                                        ArchivoZip = App.Path & "\" & Arr_Categorias(Seccion).Descripcion
                                        ArchivoZip = ArchivoZip & "\" & Arr_Categorias(Seccion).Descripcion & ".zip"
                                        If VBOpenFile(ArchivoZip) Then
                                            AbrirZip App.Path & "\" & Arr_Categorias(Seccion).Descripcion, ArchivoZip
                                        
                                            'verificar si archivo existe
                                            ArchivoCode = App.Path & glbRecordset!Linea
                                            If Not VBOpenFile(ArchivoCode) Then
                                                'alguien borro el archivo de código
                                                Exit Do
                                            End If
                                        
                                            'abrir el archivo de código
                                            Open ArchivoCode For Input As #nFreeFile2
                                                Do While Not EOF(nFreeFile2)
                                                    Line Input #nFreeFile2, Linea
                                                    Print #nFreeFile, Linea
                                                Loop
                                            Close #nFreeFile2
                                        End If
                                    Else
                                        Exit Do
                                    End If
                                    glbRecordset.MoveNext
                                Loop
                                
                                glbRecordset.Close
                            End If
                        Next k
                    Close #nFreeFile
                    
                    ret = True
                End If
            End If
        Else
            MsgBox "No hay itemes seleccionados en bookmark.", vbCritical
        End If
    Else
        MsgBox "No hay itemes en bookmark.", vbCritical
    End If
    
    GoTo SalirExportar
    
ErrorExportar:
    ret = False
    MsgBox "Exportar : " & Err & " " & Error$
    Resume SalirExportar
    
SalirExportar:
    Call MyMsg("Listo")
    Call InhabilitaToolbar(True)
    Call Hourglass(hWnd, False)
    Exportar = ret
    Err = 0
    
End Function
'obtener detalle del proyecto
Private Sub GetProjectDetails(ByVal Archivo As String)
    
    On Error Resume Next
   
    If Not VBOpenFile(Archivo) Then
        MsgBox "Archivo no encontrado. Seleccione otro archivo", vbInformation
        Exit Sub
    End If
   
    Call Hourglass(hWnd, True)
    Call InhabilitaToolbar(False)
   
    Call MyMsg("Analizando " & ExtractFileName(Archivo) & " ...")
        
    ' obtener la extension del archivo
    Select Case UCase$(ExtractFileExt(Archivo))
        Case "VBP"  'abrir archivo y extraer
            Dim nMark As Integer, nHandle As Integer
            Dim sString As String, sFile As String, sPath As String
            sPath = ExtractPath(Archivo)

            nHandle = FreeFile
            Open Archivo For Input Access Read Shared As #nHandle
      
            Do While Not EOF(nHandle)  ' Loop until end of file.
                Line Input #nHandle, sString
      
                If UCase$(Left(sString, 4)) = "FORM" Then
                    nMark = InStr(sString, "=")
                    If nMark > 0 Then
                        sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
                    End If
                    SetOutline sFile, MT_FORM, True
                ElseIf UCase$(Left(sString, 6)) = "MODULE" Then
                    nMark = InStr(sString, ";")
                    If nMark > 0 Then
                        sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
                    End If
                    SetOutline sFile, MT_MODULE, True
                ElseIf UCase$(Left(sString, 5)) = "CLASS" Then
                    nMark = InStr(sString, ";")
                    If nMark > 0 Then
                        sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
                    End If
                    SetOutline sFile, MT_CLASS, True
                ElseIf UCase$(Left(sString, 11)) = "USERCONTROL" Then
                    nMark = InStr(sString, "=")
                    If nMark > 0 Then
                        sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
                    End If
                    SetOutline sFile, MT_CONTROL, True
                ElseIf UCase$(Left(sString, 12)) = "PROPERTYPAGE" Then
                    nMark = InStr(sString, "=")
                    If nMark > 0 Then
                        sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
                    End If
                    SetOutline sFile, MT_PROPERTY, True
                ElseIf UCase$(Left(sString, 12)) = "USERDOCUMENT" Then
                    nMark = InStr(sString, "=")
                    If nMark > 0 Then
                        sFile = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
                    End If
                End If
            Loop
            Close #nHandle
            
            Unload FRMWait
            frmSelCodigo.Show vbModal
        Case "FRM"
           SetOutline Archivo, MT_FORM
        
        Case "BAS"
           SetOutline Archivo, MT_MODULE
        
        Case "CLS"
           SetOutline Archivo, MT_CLASS
        
        Case "CTL"
           SetOutline Archivo, MT_CONTROL
        
        Case "PAG"
           SetOutline Archivo, MT_PROPERTY
        
        Case "DOB"
           SetOutline Archivo, MT_DOCUMENT

    End Select

    Unload FRMWait
    
    MyMsg ("Listo")
    
    Call Hourglass(hWnd, False)
    Call InhabilitaToolbar(True)
   
End Sub

'carga/actualiza libreria
Private Sub Actualizar(ByVal Inicio As Integer)

    Dim k As Integer
    Dim total As Long
    
    If Inicio = 0 Then  'inicio
        If Not AbrirBaseDatos() Then
            MsgBox "La base de datos de código no ha podido ser abierta.", vbCritical
            End
        End If
        
        If Not CargarLibreria() Then
            MsgBox "No se puedo cargar la libreria de código.", vbCritical
            End
        End If
    Else        'actualización
        Call MyMsg("Cargando información ...")
        Call InhabilitaToolbar(False)
        Call Hourglass(hWnd, True)
        Call CompruebaCambios
        
        'limpiar
        lvwCat.ListItems.Clear
        lvwItemes.ListItems.Clear
        rtbCodigo.Text = ""
        tabMain.Tabs(2).Caption = "Código &fuente item"
        
        If Not CargarLibreria() Then
            MsgBox "No se puedo cargar la libreria de código.", vbCritical
            End
        End If
        
        Call MyMsg("Listo")
        Call Hourglass(hWnd, False)
        Call InhabilitaToolbar(True)
    End If
            
    'contar codigo
    Call ContarCodigo
    
End Sub

'graba item
Public Function GrabaItem(ByVal Seccion As Integer, ByVal item As Integer, _
                           ByVal DescripItem As String) As Boolean

    Dim ret As Boolean
    
    ret = True
    
    glbSQL = "SELECT id from itemes where "
    glbSQL = glbSQL & "id = " & Seccion
    glbSQL = glbSQL & " and item = " & item
    
    glbRecordset.Open glbSQL, glbConnection
    
    If glbRecordset.EOF Then
        glbSQL = "insert into itemes (id, item, descripción) values ("
        glbSQL = glbSQL & Seccion & " , " & item & " , '" & DescripItem & "')"
    Else
        glbSQL = "update itemes set descripción = '" & DescripItem & "'"
        glbSQL = glbSQL & " where id = " & Seccion
        glbSQL = glbSQL & " and item = " & item
    End If
    
    glbRecordset.Close
    
    glbConnection.Execute glbSQL

    GrabaItem = ret
    
End Function

'imprimir código fuente
Private Function ImprimirCódigo() As Boolean
    
    On Local Error GoTo ErrorImprimirCódigo
    
    Dim ret As Boolean
    
    ret = True
    
    rtbCodigo.SelPrint Printer.hdc
    
    GoTo SalirImprimirCódigo
    
ErrorImprimirCódigo:
    ret = False
    MsgBox "ImprimirCódigo : " & Err & " " & Error$, vbCritical
    Resume SalirImprimirCódigo
    
SalirImprimirCódigo:
    ImprimirCódigo = ret
    Err = 0
    
End Function

'mensaje en statusbar
Private Sub MyMsg(ByVal Msg As String)
    stbMain.Panels(1).Text = Msg
    DoEvents
End Sub

'respaldar codigo
Private Sub RespaldarCodigo()

    Dim k As Integer
    Dim j As Integer
    Dim Path As String
    Dim ArchivoZip As String
    Dim ArchivoBkp As String
    Dim ArchivoCode As String
    Dim First As Boolean
    Dim total_itemes As Integer
    
    Call Hourglass(hWnd, True)
    
    InhabilitaToolbar False
    
    'ciclar x las categorias
    For k = 1 To UBound(Arr_Categorias)
        'verificar si hay codigo en cada una de estas
        total_itemes = ContarItemes(k)
        
        If total_itemes > 0 Then
            First = True
            
            Path = App.Path & "\" & Arr_Categorias(k).Descripcion
            ArchivoZip = Path & "\" & Arr_Categorias(k).Descripcion & ".zip"
            ArchivoBkp = Path & "\" & Arr_Categorias(k).Descripcion & ".bkp"
            
            'ver si se respalda categoria
            filCode.Path = App.Path
            filCode.Refresh
            filCode.Path = Path
            filCode.Refresh
            filCode.Pattern = "*.dat"
        
            'se hizo algo ?
            If filCode.ListCount - 1 > 0 Then
                Call MyMsg("Guardando : " & Arr_Categorias(k).Descripcion)
                
                'verificar si existe
                If VBOpenFile(ArchivoZip) Then
                    First = False
                Else
                    First = True
                End If
                
                'renombrar zip antiguo
                'MoveFile ArchivoZip, ArchivoBkp
          
                'ciclar x el codigo
                glbSQL = "select linea from codigo"
                glbSQL = glbSQL & " where "
                glbSQL = glbSQL & "     id = " & k
                glbSQL = glbSQL & " and correlativo = 0"
                glbSQL = glbSQL & " group by"
                glbSQL = glbSQL & " linea"
                
                glbRecordset.Open glbSQL, glbConnection
                
                Do While Not glbRecordset.EOF
                    'archivo de codigo
                    ArchivoCode = App.Path & Trim$(glbRecordset!Linea)
                    
                    'verificar si archivo de código existe
                    If VBOpenFile(ArchivoCode) Then
                        'zipear el archivo de codigo
                        Call Zipear(ArchivoZip, ArchivoCode, First)
                    End If
                    
                    'borrar el archivo de código
                    DeleteFile ArchivoCode
                    
                    glbRecordset.MoveNext
                Loop
                
                'DeleteFile ArchivoBkp
                
                glbRecordset.Close
            End If
        End If
    Next k
    
    InhabilitaToolbar True
    
    Call Hourglass(hWnd, False)
    
End Sub
'respaldar la libreria
Private Sub RespaldarLibreria(ByVal Archivo As String)

    Dim k As Integer
    Dim First As Boolean
    Dim Path As String
    Dim ArchivoCat As String
    Dim ArchivoBkp As String
    Dim total_itemes As Long
    
    'renombrar respaldo antiguo
    'MoveFile Archivo, App.Path & "\libreria.bkp"
    
    First = True
    
    Call Wait("Respaldando libreria. Espere ...", 1, 14)
    
    'respaldar archivo .zip
    Call Zipear(Archivo, App.Path & "\plibrary.mdb", First)
                
    'respaldado los archivos ciclar x todas las categorias
    For k = 1 To UBound(Arr_Categorias)
        FRMWait.lblGlosa.Caption = Arr_Categorias(k).Descripcion
        FRMWait.pgb.Value = k
        
        'verificar si hay codigo en cada una de estas
        total_itemes = ContarItemes(k)
        
        If total_itemes > 0 Then
            'archivo
            Path = App.Path & "\" & Arr_Categorias(k).Descripcion
            ArchivoCat = Path & "\" & Arr_Categorias(k).Descripcion & ".zip"
            ArchivoBkp = Path & "\" & Arr_Categorias(k).Descripcion & ".bkp"
        
            'ver si se respalda categoria
            filCode.Path = App.Path
            filCode.Refresh
            filCode.Path = Path
            filCode.Refresh
            filCode.Pattern = "*.zip"
    
            'se hizo algo ?
            If filCode.ListCount > 0 Then
                Call MyMsg("Respaldando : " & ArchivoCat)
                
                'respaldar archivo .zip
                Call Zipear(Archivo, ArchivoCat, First)
                
            End If
        End If
    Next k
    
    'eliminar respaldo antiguo
    If First = False Then
        DeleteFile App.Path & "\libreria.bkp"
    End If
    
    InhabilitaToolbar True

    Unload FRMWait
    
    Call Hourglass(hWnd, False)
            
End Sub

'zipea los archivos
Private Sub Zipear(ByVal ArchivoZip As String, ByVal ArchivoCode As String, _
                   ByRef First As Boolean)

    With m_cZ
       .ZipFile = ArchivoZip
       .StoreFolderNames = False
       .RecurseSubDirs = False
       .ClearFileSpecs
       
       If First Then
            m_cZ.AllowAppend = False
            First = False
        Else
            m_cZ.AllowAppend = True
        End If
            
        .AddFileSpec ArchivoCode
       .Zip
    End With
            
End Sub

Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    Dim i As Integer
    Dim bFound As Boolean
    
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
    
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)

    Select Case Index
        Case 0  'Ir Primero
            txtNav.Text = 1
            Call Navegar(txtNav.Text)
        Case 1  'Retroceder
            If Validar() Then
                If CLng(txtNav.Text) > 1 Then
                    txtNav.Text = txtNav.Text - 1
                    Call Navegar(txtNav.Text)
                Else
                    txtNav.Text = 1
                    Call Navegar(txtNav.Text)
                End If
            End If
        Case 2  'Avanzar
            If CLng(txtNav.Text) < intPageCount Then
                txtNav.Text = txtNav.Text + 1
                Call Navegar(txtNav.Text)
            Else
                txtNav.Text = intPageCount
                Call Navegar(txtNav.Text)
            End If
        Case 3  'Ir al ultimo
            txtNav.Text = intPageCount
            Call Navegar(intPageCount)
    End Select
    
End Sub

Private Sub Form_Activate()
    Timer1.Enabled = True
End Sub

Private Sub Navegar(ByVal Seccion As Integer, ByVal Pagina As Long)

    'numero de registro de pagina
    Dim Rec As Long
    Dim Key As String
    
    If Pagina = 1 Then
        Rec = 1
    Else
        Rec = (15 * (Pagina - 1)) + 1
    End If
    
    lvwItemes.ListItems.Clear
    
    If glbRecordset.EOF Then Exit Sub
    
    glbRecordset.AbsolutePage = Pagina

    intCount = 1

    'navegar por los registros de la pagina
    For intRecord = 1 To glbRecordset.PageSize
        
        Key = "k" & glbRecordset!item
        
        ValidateRect lvwItemes.hWnd, 0&
        
        If (intRecord Mod 5) = 0 Then InvalidateRect lvwItemes.hWnd, 0&, 0&
        
        lvwItemes.ListItems.Add intCount, Key, Format(Rec, "0000"), 21, 21
        lvwItemes.ListItems(Key).Tag = Seccion & "-" & glbRecordset!item
        lvwItemes.ListItems(Key).SubItems(1) = Trim$(glbRecordset!descripción)
        
        intCount = intCount + 1
                      
        glbRecordset.MoveNext
        
        Rec = Rec + 1
        If glbRecordset.EOF Then Exit For
    Next

End Sub

Private Sub Form_Load()

    Dim k As Integer
    Dim Valor As Variant
    
    'carga/actualiza libreria
    Call Actualizar(0)
    
    'configuracion para colorizar código
    Call InitColorize
    
    'configurar menu
    Set MyHelpCallBack = New HelpCallBack
    Call clsXmenu.Install(hWnd, MyHelpCallBack, Me.imgTb)
    Call clsXmenu.FontName(hWnd, "Tahoma")
    
    'configurar zip
    Set m_cZ = New cZip
    Set m_cUnzip = New cUnzip
    
    mbDontNavigateNow = True
    glbPaginaInicio = LeeIni("web", "inicio", C_INI)
    If glbPaginaInicio = "" Then
        glbPaginaInicio = "http://www.vbsoftware.cl"
    End If
    
    cboAddress.AddItem glbPaginaInicio
    cboAddress.ListIndex = 0
    
    'cargar historial
    Valor = LeeIni("web", "numero", C_INI)
        
    If Valor = "" Then
        Call GrabaIni(C_INI, "web", "numero", 5)
        cboAddress.AddItem "http://www.vbsoftware.cl"
        cboAddress.AddItem "http://www.vbcode.com"
        cboAddress.AddItem "http://www.vbdiamond.com"
        cboAddress.AddItem "http://www.vbaccelerator.com"
        cboAddress.AddItem "http://www.lawebdelprogramador.com"
    Else
        'cargar sitios visitados
        For k = 1 To Valor
            Valor = LeeIni("web", "www" & k, C_INI)
            cboAddress.AddItem Valor
        Next k
    End If
    
    Splitter.Move ScaleWidth \ 3, CTRL_OFFSET + 2, SPLT_WDTH, (ScaleHeight - (CTRL_OFFSET * 2)) - 4
        
    Call Form_Resize
            
    Call CargaOpciones
    
    'eliminar x
    RemoveMenus Me, False, False, _
        False, False, False, True, True
        
    fraMain(1).Visible = True
    
    mbDontNavigateNow = False
    
    SetAppHelp hWnd
    
    Call MyMsg("Listo")
    
End Sub

Private Sub Form_Paint()
    DrawNumbers
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim Msg As String
    
    Call CompruebaCambios
    
    Msg = "Confirma salir de " & App.Title
    
    If Confirma(Msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
            
    Call MyMsg("Cerrando ...")
    
    tmrCodigo.Enabled = False
    Timer1.Enabled = False
    
    Call RespaldarCodigo
        
    If glbRespaldarLibreria Then
        Call RespaldarLibreria(App.Path & "\libreria.zip")
    End If
    
    Call GrabarHistorialWeb
    
End Sub
'graba el historial de las paginas web
Private Sub GrabarHistorialWeb()

    Dim k As Integer
    
    Call GrabaIni(C_INI, "web", "numero", cboAddress.ListCount - 1)
    For k = 0 To cboAddress.ListCount - 1
        Call GrabaIni(C_INI, "web", "www" & k + 1, cboAddress.List(k))
    Next k
    
End Sub

Sub DrawNumbers()
    
    Dim LineCount As Long '// How many lines in total
    Dim i As Integer      '// Just an integer
    
    Timer1.Enabled = False
    Call MyMsg("Contando lineas de código ...")
    
    '// Get number of lines in Rtftext
    LineCount = SendMessage(rtbCodigo.hWnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1
        
    '// Same lines ?
    LineCountChange = LineCount
    
    
    '// Get first visible line in rtfText
    FirstLine = SendMessage(rtbCodigo.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary
    
    PicLines.Cls '// Clear the PicLines
    PicLines.CurrentY = 40  '// Move the .top text by 40 twips
    
    '// Print the number of each line on a picture
    For i = 0 To LineCount - FirstLine
       PicLines.CurrentY = PicLines.CurrentY + 7.49 '// Where on Y
       PicLines.CurrentX = 20 '-2                   '// Where on X
       PicLines.Print i + FirstLine + 1             '// print the number
    Next
    'LineCountChange = LineCount '// Remember the last line count
    FirstLineNow = FirstLine     '// Is the first visible line still the same ?
    
    Call MyMsg("Listo")
    
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Resize()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim FrameWidth As Integer
    
    If WindowState <> vbMinimized Then
        Timer1.Enabled = True
        
        DoEvents
        ' maximized, lock update to avoid nasty window flashing
        If WindowState = vbMaximized Then Call LockWindowUpdate(hWnd)
        
        Call Hourglass(hWnd, True)

        ' handle minimum height. if you were to remove the
        ' controlbox you would need to handle minimum width also
        If Height < 3500 Then Height = 3500
        If Width < 3500 Then Width = 3500
        
        picMain.Left = 0
        picMain.Top = tlbMain.Height + 1
        picMain.Height = ScaleHeight - tlbMain.Height - stbMain.Height
                
        ' the width of the window frame
        FrameWidth = ((Width \ Screen.TwipsPerPixelX) - ScaleWidth) \ 2
    
        ' handle a form resize that hides the vertical splitter
        If ((ScaleWidth - CTRL_OFFSET) - (Splitter.Left + Splitter.Width)) < 12 Then
            Splitter.Left = ScaleWidth - ((CTRL_OFFSET * 4) + (FrameWidth * 2))
        End If
                
        'height y width del picture que contiene el treeview proyecto origen
        Dim height_picOri As Integer
        Dim Width_picOri As Integer
                
        picCat.Move 24, picMain.Top, Splitter.Left - Splitter.Width - picMain.Width + 3
        
        height_picOri = ScaleHeight - tlbMain.Height - stbMain.Height
        Width_picOri = Splitter.Left - Splitter.Width - picMain.Width + 3
                
        lvwCat.Move 24, picMain.Top + picCat.Height, Width_picOri, height_picOri - picCat.Height
                
        'cambiar el tamaño del splitter
        Splitter.Top = 25
        Splitter.Height = lvwCat.Height + picCat.Height
                        
        'height y width del picture que contiene el treeview proyecto destino
        Dim height_picDes As Integer
        Dim Width_picDes As Integer
        Dim left_picDes As Integer
                
        height_picDes = ScaleHeight - tlbMain.Height - stbMain.Height
        Width_picDes = ScaleWidth - Splitter.Width - picMain.Width - lvwCat.Width
                
        tabMain.Move Splitter.Left + 5, picMain.Top, Width_picDes - 2, height_picDes
        
        'frames de datos
        fraMain(1).Left = tabMain.Left + 5
        fraMain(1).Top = tabMain.Top + 22
        fraMain(1).Height = tabMain.Height - 30
        fraMain(1).Width = tabMain.Width - 10
                
        lblDesCat.Width = fraMain(1).Width * Screen.TwipsPerPixelY - 450
                
        'contador lineas de codigo
        PicLines.Top = 650
        PicLines.Left = 30
        PicLines.Height = fraMain(1).Height * Screen.TwipsPerPixelX - 660
        
        lvwItemes.Left = 30
        lvwItemes.Top = 450
        lvwItemes.Height = fraMain(1).Height * Screen.TwipsPerPixelX - 430
        lvwItemes.Width = fraMain(1).Width * Screen.TwipsPerPixelY - 30
        
        fraMain(2).Move fraMain(1).Left, fraMain(1).Top, fraMain(1).Width, fraMain(1).Height
                
        lblDescripItem.Left = 20
        lblDescripItem.Width = fraMain(2).Width * Screen.TwipsPerPixelY - 10 '- 450
        
        tbCodigo.Left = 30
        tbCodigo.Top = 300
        tbCodigo.Width = fraMain(2).Width * Screen.TwipsPerPixelY - 30
        
        rtbCodigo.Left = 450
        rtbCodigo.Top = 630
        rtbCodigo.Height = fraMain(2).Height * Screen.TwipsPerPixelX - 650
        rtbCodigo.Width = fraMain(2).Width * Screen.TwipsPerPixelY - 450
        
        'PicLines.Top = 300
        'PicLines.Left = 450
        'PicLines.Height = rtbCodigo.Height - 300
        
        'internet
        fraMain(3).Move fraMain(1).Left, fraMain(1).Top, fraMain(1).Width, fraMain(1).Height
        brwWebBrowser.Left = 30
        brwWebBrowser.Height = fraMain(3).Height * Screen.TwipsPerPixelX - 900
        brwWebBrowser.Width = fraMain(3).Width * Screen.TwipsPerPixelY - 30
        
        cboAddress.Width = fraMain(3).Width * Screen.TwipsPerPixelY - 850
        
        'bookmark
        fraMain(4).Move fraMain(1).Left, fraMain(1).Top, fraMain(1).Width, fraMain(1).Height
        
        lvwBookmark.Left = 30
        lvwBookmark.Top = 100
        lvwBookmark.Height = fraMain(4).Height * Screen.TwipsPerPixelX - 100
        lvwBookmark.Width = fraMain(4).Width * Screen.TwipsPerPixelY - 30
        
        pgbStatus.Top = ScaleHeight - 15
        pgbStatus.Left = stbMain.Panels(2).Left + 4
        pgbStatus.Height = stbMain.Height - 10
        pgbStatus.Width = stbMain.Panels(2).Width - 7
        pgbStatus.ZOrder 0

        With mGradient
            .Angle = 90 '.Angle
            .Color1 = 16744448
            .Color2 = 0
            .Draw picMain
        End With
            
        Call FontStuff(picMain, App.Title & " Beta Versión : " & App.major & "." & App.minor & "." & App.Revision)
                        
        picMain.Refresh
        
        If Splitter.Left < 24 Then
            Splitter.Left = 200
            Call Form_Resize
        End If
                
        Splitter.ZOrder 0
                
        ' if it's locked unlock the window
        If WindowState = vbMaximized Then Call LockWindowUpdate(0&)
        Call Hourglass(hWnd, False)
        DoEvents
    Else
        Timer1.Enabled = False
    End If
    
    Err = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Call QuitHelp
    Set m_cZ = Nothing
    Set m_cUnzip = Nothing
    
End Sub

Private Sub lvwBookmark_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuBookmark
    End If
    
End Sub


Private Sub lvwCat_ItemClick(ByVal item As MSComctlLib.ListItem)
    
    Call CargaItemes
        
End Sub
'carga los itemes
Public Sub CargaItemes()

    Dim Seccion As Integer
    Dim k As Integer
    Dim Key As String
    Dim total As Integer
    Dim item As ListItem
    Dim i As Integer
    Dim e As Long
    
    'hay una categoria seleccionada
    If lvwCat.SelectedItem Is Nothing Then
        MsgBox "Debes seleccionar una categoria.", vbCritical
        Exit Sub
    End If
    
    Call Hourglass(hWnd, True)
    
    'hay cambios
    Call CompruebaCambios
    
    Cargando = True
    
    Call MyMsg("Cargando itemes ...")
    Call InhabilitaToolbar(False)
    Call ShowProgress(True)
                
    Timer1.Enabled = False
    
    Set item = lvwCat.SelectedItem
        
    Seccion = item.Index
    UltimaCat = Seccion
    
    imgCat.Picture = imgCatItem.ListImages(Seccion).Picture
    lblDesCat.Caption = Arr_Categorias(Seccion).Descripcion
        
    lblDescripItem.Caption = Arr_Categorias(Seccion).Descripcion
    
    lvwItemes.ListItems.Clear
    rtbCodigo.Text = ""
        
    tabMain.Tabs(2).Caption = "Código &fuente item"
    
    'cargar itemes asociados
    pgbStatus.Min = 1
    total = ContarItemes(Seccion)
    If total > 0 Then
        pgbStatus.Max = total + 1
        stbMain.Panels(2).Text = "1 de " & total
    Else
        stbMain.Panels(2).Text = ""
    End If
    
    stbMain.Panels(4).Text = ""
    stbMain.Panels(5).Text = ""
    
    'cargar itemes
    If glbRecordset.State > 0 Then glbRecordset.Close
    glbRecordset.CursorLocation = adUseClient
    glbSQL = "select item, descripción from itemes where id = " & Seccion
    glbRecordset.Open glbSQL, glbConnection ', , , adCmdTable
    glbRecordset.PageSize = 15
    intPageCount = glbRecordset.PageCount
    
    txtNav.Text = 1
    Call Navegar(Seccion, txtNav.Text)
        
    InvalidateRect lvwItemes.hWnd, 0&, 0&
    
    If lvwItemes.ListItems.Count > 0 Then
        lvwItemes.ListItems(1).Selected = True
    End If
    
    tabMain.Tabs(1).Selected = True
    tabMain.Tabs(1).Caption = "Código de sección : (" & lvwItemes.ListItems.Count & ")"
    
    Cargando = False
    
    stbMain.Panels(5).Text = ""
        
    Timer1.Enabled = True
    
    Set Itmx = Nothing
    
    Call ShowProgress(False)
    Call InhabilitaToolbar(True)
    Call Hourglass(hWnd, False)
    Call MyMsg("Listo")
    
End Sub

Private Sub lvwItemes_DblClick()
    
    Dim Itmx As ListItem
    
    If Cargando Then Exit Sub
    
    If lvwItemes.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    tabMain.Tabs(2).Selected = True
    
    Set Itmx = lvwItemes.SelectedItem
    
    lblDescripItem.Caption = Itmx.SubItems(1)
        
    On Local Error Resume Next
    rtbCodigo.SetFocus
    Err = 0
    
End Sub

Private Sub lvwItemes_ItemClick(ByVal item As MSComctlLib.ListItem)

    Dim Seccion As Integer
    Dim nitem As Integer
    Dim k As Integer
    Dim Path As String
    Dim Archivo As String
    
    Timer1.Enabled = False
        
    'categoria seleccionada ?
    If lvwCat.SelectedItem Is Nothing Then
        Timer1.Enabled = True
        Exit Sub
    End If
    
    'item seleccionado ?
    If lvwItemes.SelectedItem Is Nothing Then
        Timer1.Enabled = True
        Exit Sub
    End If
    
    'hay cambios anteriores
    
    Call Hourglass(hWnd, True)
    Call InhabilitaToolbar(False)
    
    Call MyMsg("Cargando itemes ...")
    
    Call CompruebaCambios
        
    Seccion = lvwCat.SelectedItem.Index
    nitem = Val(Mid$(lvwItemes.SelectedItem.Tag, InStr(1, lvwItemes.SelectedItem.Tag, "-") + 1))
    
    Cargando = True
    
    Path = App.Path & "\" & Arr_Categorias(Seccion).Descripcion
    
    'buscar el archivo asociado
    Archivo = ArchivoDeCodigo(Seccion, nitem)
                
    'cargar codigo asociado a categoria/item
    rtbCodigo.Text = ""
    rtbCodigo.SelColor = RGB(0, 0, 0)
    
    If Len(Archivo) > 0 Then
        'verificar si los archivos se cargaron
        filCode.Path = App.Path
        filCode.Refresh
        filCode.Path = Path
        filCode.Refresh
        filCode.Pattern = "*.dat"
        
        InhabilitaToolbar False
        
        'hay archivos de codigo ?
        If filCode.ListCount <= 0 Then
            Call MyMsg("Cargando : " & Arr_Categorias(Seccion).Descripcion)
            AbrirZip Path, Path & "\" & Arr_Categorias(Seccion).Descripcion & ".zip"
        End If
        
        'existe archivo de codigo ?
        If Not VBOpenFile(Archivo) Then
            AbrirZip Path, Path & "\" & Arr_Categorias(Seccion).Descripcion & ".zip"
        End If
        
        'existe archivo de codigo ?
        If VBOpenFile(Archivo) Then
            rtbCodigo.LoadFile Archivo, rtfText
            Call ColorizeVB(Me.rtbCodigo)
        End If
    End If
        
    glbLinea = ""
    glbCambio = False
    Cargando = False
        
    Timer1.Enabled = True
    
    Call InhabilitaToolbar(True)
    Call Hourglass(hWnd, False)
    Call MyMsg("Listo")
    
End Sub

'abrir archivo .zip de codigo
Private Function AbrirZip(ByVal Path As String, ByVal Archivo As String) As Boolean
          
    m_cUnzip.ZipFile = Archivo
    m_cUnzip.Directory
    m_cUnzip.UnzipFolder = Path
    m_cUnzip.Unzip
        
End Function


Private Sub lvwItemes_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        lvwItemes_DblClick
    End If
    
End Sub

Private Sub lvwItemes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuLibreria
    End If
    
End Sub

Private Sub mnuArchivo_ConfImpresora_Click()

    Call cc.VBPageSetupDlg(hWnd)

End Sub

Private Sub mnuArchivo_Exportar_Bas_Click()

    Dim Msg As String
    Dim Glosa As String
            
    Msg = "Confirma generar archivo Módulo .Bas"
    
    Glosa = "Archivos de módulos .bas (*.BAS)|*.BAS|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Exportar(Msg, Glosa, "BAS") Then
        MsgBox "Código exportado con éxito!", vbInformation
    End If
    
End Sub

Private Sub mnuArchivo_Exportar_Cls_Click()

    Dim Msg As String
    Dim Glosa As String
            
    Msg = "Confirma generar archivo Módulo .Cls"
    
    Glosa = "Archivos de módulos .bas (*.CLS)|*.CLS|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Exportar(Msg, Glosa, "CLS") Then
        MsgBox "Código exportado con éxito!", vbInformation
    End If
    
End Sub
Private Sub mnuArchivo_Importar_Bas_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdBas"))
End Sub

Private Sub mnuArchivo_Importar_Cls_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdCls"))
End Sub


Private Sub mnuArchivo_Importar_Control_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdCtl"))
End Sub

Private Sub mnuArchivo_Importar_Formulario_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdForm"))
End Sub


Private Sub mnuArchivo_Importar_Pagina_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdPag"))
End Sub

Private Sub mnuArchivo_Importar_Proyecto_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdVbp"))
End Sub


Private Sub mnuArchivo_Impresora_Click()
    
    Dim Msg As String
    
    If Len(rtbCodigo.Text) > 0 Then
        Msg = "Confirma imprimir código."
        If Confirma(Msg) = vbYes Then
            If ImprimirCódigo() Then
                MsgBox "Código impreso con éxito", vbInformation
            End If
        End If
    End If
    
End Sub

Private Sub mnuArchivo_Salir_Click()
    Unload Me
End Sub

Private Sub mnuAyuda_AcercaDe_Click()
    frmAcerca.Show vbModal
End Sub

Private Sub mnuAyuda_Buscar_Click()
    Call SearchHelp
End Sub

Private Sub mnuAyuda_Email_Click()
    Shell_Email
End Sub

Private Sub mnuAyuda_Indice_Click()
    Call ShowHelpContents
End Sub

Private Sub mnuAyuda_Tips_Click()
    frmTip.Show vbModal
End Sub

Private Sub mnuAyuda_WebSite_Click()
    Shell_PaginaWeb
End Sub


Private Sub mnuBookmark_ExportarBas_Click()
    mnuArchivo_Exportar_Bas_Click
End Sub

Private Sub mnuBookmark_ExportarCls_Click()
    mnuArchivo_Exportar_Cls_Click
End Sub

Private Sub mnuBookmark_Invertir_Click()

    Dim k As Integer
    Dim e As Long
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwBookmark.ListItems.Count
        e = DoEvents
        ValidateRect lvwBookmark.hWnd, 0&
        If (k Mod 10) = 0 Then InvalidateRect lvwBookmark.hWnd, 0&, 0&
        lvwBookmark.ListItems(k).Checked = Not lvwBookmark.ListItems(k).Checked
    Next k
    
    InvalidateRect lvwBookmark.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuBookmark_Quitar_Click()
    mnuLibreria_QuitarBookmark_Click
End Sub

Private Sub mnuBookmark_SelTodos_Click()

    Dim k As Integer
    Dim e As Long
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwBookmark.ListItems.Count
        e = DoEvents
        ValidateRect lvwBookmark.hWnd, 0&
        If (k Mod 10) = 0 Then InvalidateRect lvwBookmark.hWnd, 0&, 0&
        lvwBookmark.ListItems(k).Checked = True
    Next k
    
    InvalidateRect lvwBookmark.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuEdicion_Borrar_Click()
    
    Dim Msg As String
    
    If Len(rtbCodigo.SelText) > 0 Then
        rtbCodigo.SelText = ""
    Else
        If Len(rtbCodigo.Text) > 0 Then
            Msg = "Confirma borrar contenido"
            If Confirma(Msg) = vbYes Then
                rtbCodigo.Text = ""
            End If
        End If
    End If
End Sub

Private Sub mnuEdicion_Buscar_Click()
    If Len(rtbCodigo.Text) > 0 Then
        frmMain.tabMain.Tabs(2).Selected = True
        frmFind.Show
    End If
End Sub

Private Sub mnuEdicion_BusSiguiente_Click()
    Call FindText
End Sub

Private Sub mnuEdicion_Copiar_Click()
    Clipboard.Clear
    Clipboard.SetText rtbCodigo.SelText
End Sub

Private Sub mnuEdicion_Cortar_Click()
    Clipboard.SetText rtbCodigo.SelText
    rtbCodigo.SelText = ""
End Sub


Private Sub mnuEdicion_ExpHtml_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdHtml"))
End Sub

Private Sub mnuEdicion_ExpRtf_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdRtf"))
End Sub

Private Sub mnuEdicion_ExpTexto_Click()
    Call tlbMain_ButtonClick(tlbMain.Buttons("cmdTexto"))
End Sub

Private Sub mnuEdicion_Pegar_Click()
    rtbCodigo.SelText = Clipboard.GetText(rtfText)
    Call DrawNumbers
End Sub


Private Sub mnuEdicion_Reemplazar_Click()
    If Len(rtbCodigo.Text) > 0 Then
        frmMain.tabMain.Tabs(2).Selected = True
        frmReemplazar.Show
    End If
End Sub

Private Sub mnuEdicion_SelTodo_Click()
    
    If Len(rtbCodigo.Text) > 0 Then
        tabMain.Tabs(2).Selected = True
        rtbCodigo.SelStart = 0
        rtbCodigo.SelLength = Len(rtbCodigo.Text)
        rtbCodigo.SetFocus
    End If
    
End Sub

'actualiza cambios a la libreria
Private Sub ActualizaCambios()

    On Local Error GoTo ErrorActualizaCambios
    
    Dim Msg As String
    Dim Seccion As Integer
    Dim item As Integer
    Dim DescripItem As String
    Dim lineas As Integer
    Dim Linea As String
    Dim Archivo As String
    
    Dim k As Long
    
    If Cargando Then Exit Sub
    
    If lvwItemes.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Call Hourglass(hWnd, True)
    Call InhabilitaToolbar(False)
    
    'categoria
    Seccion = UltimaCat
        
    'item seleccionado
    item = Val(Mid$(lvwItemes.SelectedItem.Tag, InStr(1, lvwItemes.SelectedItem.Tag, "-") + 1))
    
    DescripItem = lvwItemes.SelectedItem.SubItems(1)
            
    Archivo = "\" & Arr_Categorias(Seccion).Descripcion & "\"
    Archivo = Archivo & Arr_Categorias(Seccion).Descripcion & "_" & item & ".dat"
        
    'iniciar trx
    glbConnection.IsolationLevel = adXactReadCommitted
    glbConnection.BeginTrans
    
    'actualizar info en tabla itemes
    If GrabaItem(Seccion, item, DescripItem) Then
        'eliminar el codigo anterior para reemplazarlo por el nuevo
        glbSQL = "delete from codigo "
        glbSQL = glbSQL & " where id = " & Seccion
        glbSQL = glbSQL & " and item = " & item
            
        glbConnection.Execute glbSQL
    
        'insertar la linea en el archivo
        glbSQL = "insert into codigo (id, item, correlativo , linea) values ("
        glbSQL = glbSQL & Seccion & " , " & item & " , 0 , '" & Archivo & "')"
        
        glbConnection.Execute glbSQL
        
        'guardar archivo
        rtbCodigo.SaveFile App.Path & Archivo, rtfText
    End If
    
    'fin trx
    glbConnection.CommitTrans
    glbConnection.IsolationLevel = adXactUnspecified
    
    GoTo SalirActualizaCambios
    
ErrorActualizaCambios:
    'rollback
    If glbConnection.IsolationLevel <> adXactUnspecified Then
        glbConnection.RollbackTrans
    End If
    MsgBox "ErrorActualizaCambios : " & Err & " " & Error$, vbCritical
    Resume SalirActualizaCambios
    
SalirActualizaCambios:
    glbCambio = False
    Call Hourglass(hWnd, False)
    Call InhabilitaToolbar(True)
    Err = 0
    
End Sub

Private Sub mnuLibreria_Actualizar_Click()
    Call CargaItemes
End Sub

Private Sub mnuLibreria_AgregarBookmark_Click()

    Dim k As Integer
    Dim total As Integer
    Dim e As Long
    
    Call Hourglass(hWnd, True)
    
    total = lvwBookmark.ListItems.Count + 1
    
    For k = 1 To lvwItemes.ListItems.Count
        e = DoEvents
        If lvwItemes.ListItems(k).Checked Then
            On Error Resume Next
            ValidateRect lvwBookmark.hWnd, 0&
            If (k Mod 10) = 0 Then InvalidateRect lvwBookmark.hWnd, 0&, 0&
            lvwBookmark.ListItems.Add , lvwItemes.ListItems(k).Tag, Format(total, "0000"), 44, 44
            lvwBookmark.ListItems(lvwItemes.ListItems(k).Tag).SubItems(1) = lvwItemes.ListItems(k).SubItems(1)
            total = total + 1
            Err = 0
        End If
    Next k
    
    InvalidateRect lvwBookmark.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuLibreria_AgregarItem_Click()

    Dim Seccion As Integer
        
    If lvwCat.SelectedItem Is Nothing Then
        MsgBox "Seleccione una libreria.", vbCritical
        Exit Sub
    End If
    
    'ver si hay cambios
    glbCambio = True
    Call CompruebaCambios
    
    'agregar item
    Seccion = lvwCat.SelectedItem.Index
        
    frmNew.tipo = 0
    frmNew.Seccion = Seccion
    frmNew.Show vbModal
        
End Sub

Private Sub mnuLibreria_Buscar_Click()
    Timer1.Enabled = False
    Load frmBuscar
    frmBuscar.Show
End Sub

Private Sub mnuLibreria_Invertir_Click()

    Dim k As Integer
    Dim e As Long
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwItemes.ListItems.Count
        e = DoEvents
        ValidateRect lvwItemes.hWnd, 0&
        If (k Mod 10) = 0 Then InvalidateRect lvwItemes.hWnd, 0&, 0&
        lvwItemes.ListItems(k).Checked = Not lvwItemes.ListItems(k).Checked
    Next k
    
    InvalidateRect lvwItemes.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuLibreria_ModificarItem_Click()

    On Local Error GoTo ErrormnuLibreria_ModificarItem_Click
    
    Dim Seccion As Integer
    Dim item As Integer
    Dim Descripcion As String
    Dim Itmx As ListItem
    
    'categoria seleccionada
    If lvwCat.SelectedItem Is Nothing Then
        MsgBox "Seleccione una libreria.", vbCritical
        Exit Sub
    End If
    
    'item seleccionado
    If lvwItemes.SelectedItem Is Nothing Then
        MsgBox "Seleccione un item.", vbCritical
        Exit Sub
    End If
    
    'ver si hay cambios
    glbCambio = True
    Call CompruebaCambios
    
    Set Itmx = lvwItemes.SelectedItem
    'agregar item
    Seccion = lvwCat.SelectedItem.Index
    item = Val(Mid$(Itmx.Tag, InStr(1, Itmx.Tag, "-") + 1))
        
    'ingreso/modificacion
    frmNew.tipo = 1
    frmNew.Seccion = Seccion
    frmNew.item = item
    frmNew.Key = Itmx.Key
    frmNew.txtDescrip.Text = Itmx.SubItems(1)
    frmNew.Show vbModal
        
    GoTo SalirmnuLibreria_ModificarItem_Click
    
ErrormnuLibreria_ModificarItem_Click:
    MsgBox "mnuLibreria_ModificarItem_Click : " & Err & " " & Error$, vbCritical
    Resume SalirmnuLibreria_ModificarItem_Click
    
SalirmnuLibreria_ModificarItem_Click:
    Err = 0
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuLibreria_Quitar_Click()
    
    Dim k As Integer
    Dim e As Long
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwItemes.ListItems.Count
        e = DoEvents
        ValidateRect lvwItemes.hWnd, 0&
        If (k Mod 10) = 0 Then InvalidateRect lvwItemes.hWnd, 0&, 0&
        lvwItemes.ListItems(k).Checked = False
    Next k
    
    InvalidateRect lvwItemes.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
End Sub

Private Sub mnuLibreria_QuitarBookmark_Click()

    Dim k As Integer
    Dim e As Long
    
    Call Hourglass(hWnd, True)
    
    For k = lvwBookmark.ListItems.Count To 1 Step -1
        e = DoEvents
        If lvwBookmark.ListItems(k).Checked Then
            ValidateRect lvwBookmark.hWnd, 0&
            If (k Mod 10) = 0 Then InvalidateRect lvwBookmark.hWnd, 0&, 0&
            lvwBookmark.ListItems.Remove lvwBookmark.ListItems(k).Key
        End If
    Next k
    
    InvalidateRect lvwBookmark.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuLibreria_SelTodos_Click()

    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwItemes.ListItems.Count
        ValidateRect lvwItemes.hWnd, 0&
        If (k Mod 10) = 0 Then InvalidateRect lvwItemes.hWnd, 0&, 0&
        lvwItemes.ListItems(k).Checked = True
    Next k
    
    InvalidateRect lvwItemes.hWnd, 0&, 0&
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuLibreria_VerCodigo_Click()

    If Len(rtbCodigo.Text) > 0 Then
        tabMain.Tabs(2).Selected = True
    End If
    
End Sub

Private Sub mnuOpciones_OpcEditor_Click()
    frmOpciones.Show vbModal
End Sub

Private Sub mnuOpciones_Respaldo_Click()

    Dim Msg As String
    Dim k As Integer
    Dim First As Boolean
    Dim Path As String
    Dim ArchivoCat As String
    Dim ArchivoZip As String
    Dim ArchivoBkp As String
    Dim Glosa As String
    
    Msg = "Confirma respaldar libreria."
    
    Glosa = Glosa & "Archivos .zip (*.ZIP)|*.ZIP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
    
    If Confirma(Msg) = vbYes Then
        'respaldar todos los archivos primero
        Call RespaldarCodigo
        
        'archivo libreria
        If cc.VBGetSaveFileName(ArchivoZip, , , Glosa, , App.Path, "Guardar como ...", "ZIP") Then
            Call RespaldarLibreria(ArchivoZip)
            
            MsgBox "Libreria respaldada con éxito!", vbInformation
        End If
    End If
    
    Call MyMsg("Listo")
    
End Sub

Private Sub mnuRepositorio_EliminarItem_Click()

    On Local Error GoTo ErrormnuRepositorio_EliminarItem_Click
    
    Dim Seccion As Integer
    Dim item As Integer
    Dim Descripcion As String
    Dim Itmx As ListItem
    Dim k As Integer
    Dim j As Integer
    Dim fSel As Boolean
    Dim Msg As String
    Dim i As Integer
    
    'categoria seleccionada
    If lvwCat.SelectedItem Is Nothing Then
        MsgBox "Seleccione una libreria.", vbCritical
        Exit Sub
    End If
    
    'item seleccionado
    If lvwItemes.SelectedItem Is Nothing Then
        MsgBox "Seleccione un item.", vbCritical
        Exit Sub
    End If
    
    'ver si hay cambios
    glbCambio = True
    Call CompruebaCambios
    
    'verificar si hay algo seleccionado
    For k = lvwItemes.ListItems.Count To 1 Step -1
        If lvwItemes.ListItems(k).Checked Then
            fSel = True
            Exit For
        End If
    Next k
    
    'hay seleccionado ?
    If Not fSel Then
        MsgBox "Debe seleccionar un item a eliminar.", vbCritical
        Exit Sub
    End If
    
    Set Itmx = lvwItemes.SelectedItem
        
    Msg = "Confirma eliminar código seleccionado."
    If Confirma(Msg) = vbNo Then
        Exit Sub
    End If
    
    Call Hourglass(hWnd, True)
    Call InhabilitaToolbar(False)
    
    Timer1.Enabled = False
        
    'agregar item
    Seccion = lvwCat.SelectedItem.Index
    
    i = 1
    
    'ciclar x los itemes seleccionados
    For k = lvwItemes.ListItems.Count To 1 Step -1
        'esta seleccionado ?
        If lvwItemes.ListItems(k).Checked Then
            Set Itmx = lvwItemes.ListItems(k)
            
            item = Val(Mid$(Itmx.Tag, InStr(1, Itmx.Tag, "-") + 1))
        
            'eliminar código
            glbSQL = "delete from codigo where "
            glbSQL = glbSQL & "     id = " & Seccion
            glbSQL = glbSQL & " and item = " & item
                                
            glbConnection.Execute glbSQL
            
            'actualizar info en tabla itemes
            glbSQL = "delete from itemes where "
            glbSQL = glbSQL & "id = " & Seccion
            glbSQL = glbSQL & " and item = " & item
                                
            glbConnection.Execute glbSQL
                                    
            ValidateRect lvwItemes.hWnd, 0&
            If (i Mod 10) = 0 Then InvalidateRect lvwItemes.hWnd, 0&, 0&
        
            'eliminar item
            lvwItemes.ListItems.Remove Itmx.Key
            
            i = i + 1
        End If
    Next k
    
    InvalidateRect lvwItemes.hWnd, 0&, 0&
    
    'contar codigo
    Call ContarCodigo
    
    tabMain.Tabs(1).Caption = "Código de sección : (" & ContarItemes(Seccion) & ")"
    
    GoTo SalirmnuRepositorio_EliminarItem_Click
    
ErrormnuRepositorio_EliminarItem_Click:
    MsgBox "mnuRepositorio_EliminarItem_Click : " & Err & " " & Error$, vbCritical
    Resume SalirmnuRepositorio_EliminarItem_Click
    
SalirmnuRepositorio_EliminarItem_Click:
    Err = 0
    Timer1.Enabled = True
    Call InhabilitaToolbar(True)
    Call Hourglass(hWnd, False)

End Sub
Private Sub MyHelpCallBack_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    stbMain.Panels(1).Text = MenuHelp
End Sub

Private Sub rtbCodigo_Change()

    If Not Cargando Then
        Call DrawNumbers
    End If
End Sub

Private Sub rtbCodigo_KeyPress(KeyAscii As Integer)

    If Not Cargando Then
        If Not tmrCodigo.Enabled Then
            contador = 0
            tmrCodigo.Enabled = True
        Else
            contador = 0
        End If
    End If
    
End Sub


Private Sub rtbCodigo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuEdicion
    End If
    
End Sub

Private Sub rtbCodigo_SelChange()
    
    Dim estado As Boolean
    
    If Not Cargando Then
        glbCambio = True
    End If
    
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' if the left button is down set the flag
    If Button = 1 Then fInitiateDrag = True
End Sub


Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the flag isn't set then the left button wasn't
    ' pressed while the mouse was over one of the splitters
    If fInitiateDrag <> True Then Exit Sub

    ' if the left button is down then we want to move the splitter
    If Button = 1 Then ' if the Tag is false then we need to set
        If Splitter.Tag = False Then ' the color and clip the cursor.
    
            Splitter.BackColor = &H808080 '<- set the "dragging" color here
            Splitter.Tag = True
        End If
    
        Splitter.Left = (Splitter.Left + x) - (SPLT_WDTH \ 3)
    End If
    
End Sub


Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' if the left button is the one being released we need to reset
    ' the color, Tag, flag, cancel ClipCursor and call form_resize
  
    If Button = 1 Then           ' to move the list and text boxes
        Splitter.Tag = False
        fInitiateDrag = False
        'ClipCursor ByVal 0&
        Splitter.BackColor = &H8000000F  '<- set to original color
        Form_Resize
    End If
    
End Sub


Private Sub tabMain_Click()

    Call CompruebaCambios
    
    If tabMain.SelectedItem.Index = 1 Then      'itemes
        fraMain(1).Visible = True
        fraMain(4).Visible = False
        fraMain(3).Visible = False
        fraMain(2).Visible = False
        fraMain(2).ZOrder 0
    ElseIf tabMain.SelectedItem.Index = 2 Then
        fraMain(2).Visible = True
        fraMain(4).Visible = False
        fraMain(3).Visible = False
        fraMain(1).Visible = False
        fraMain(2).ZOrder 0
        
        If Not lvwItemes.SelectedItem Is Nothing Then
            If Len(lvwItemes.SelectedItem.SubItems(1)) <= 50 Then
                lblDescripItem.Caption = lvwItemes.SelectedItem.SubItems(1)
            Else
                lblDescripItem.Caption = Left$(lvwItemes.SelectedItem.SubItems(1), 50) & " ..."
            End If
        End If
    ElseIf tabMain.SelectedItem.Index = 3 Then 'internet
        fraMain(3).Visible = True
        fraMain(4).Visible = False
        fraMain(1).Visible = False
        fraMain(2).Visible = False
        fraMain(3).ZOrder 0
    ElseIf tabMain.SelectedItem.Index = 4 Then 'bookmark
        fraMain(4).Visible = True
        fraMain(3).Visible = False
        fraMain(2).Visible = False
        fraMain(1).Visible = False
        fraMain(4).ZOrder 0
    End If
    
    If Len(rtbCodigo.Text) > 0 Then
        DrawNumbers
    End If
    
End Sub

Private Sub CompruebaCambios()

    If glbCambio Then
        Call ActualizaCambios
        glbCambio = False
    End If
    
End Sub

Private Sub tbCodigo_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "cmdFecha" Then
        rtbCodigo.SelText = date$
    ElseIf Button.Key = "cmdHora" Then
        rtbCodigo.SelText = Time$
    End If
    
End Sub

Private Sub tbCodigo_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    Dim Buffer As String
    
    Select Case ButtonMenu.Key
        Case "cmdCabezera"
            Buffer = "'*******************************************************************************" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Módulo         : " & vbNewLine
            Buffer = Buffer & "' Archivo        : " & vbNewLine
            Buffer = Buffer & "' Autor          : " & GetUser() & vbNewLine
            Buffer = Buffer & "' Fecha Creación : " & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Copyright      : " & Year(Now) & " - " & GetComputer() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Descripción    : " & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Historial      : " & vbNewLine
            Buffer = Buffer & "' 1.0 " & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'Versión Inicial" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'*******************************************************************************" & vbNewLine
        Case "cmdFuncion"
            Buffer = "'*******************************************************************************" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' (Nombre de la Función)" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Autor          : " & GetUser() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Descripción    :" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Copyright      : " & Year(Now) & " - " & GetComputer() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'*******************************************************************************"
        Case "cmdSub"
            Buffer = "'*******************************************************************************" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' (Nombre del Procedimiento)" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Autor          : " & GetUser() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Descripción    :" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Copyright      : " & Year(Now) & " - " & GetComputer() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'*******************************************************************************"
        Case "cmdGet"
            Buffer = "'*******************************************************************************" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' (Property Get)" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Autor          : " & GetUser() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Descripción    :" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Copyright      : " & Year(Now) & " - " & GetComputer() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'*******************************************************************************"
        Case "cmdLet"
            Buffer = "'*******************************************************************************" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' (Property Let)" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Autor          : " & GetUser() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Descripción    :" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Copyright      : " & Year(Now) & " - " & GetComputer() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'*******************************************************************************"
        Case "cmdSet"
            Buffer = "'*******************************************************************************" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' (Property Set)" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Autor          : " & GetUser() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Descripción    :" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "' Copyright      : " & Year(Now) & " - " & GetComputer() & vbNewLine
            Buffer = Buffer & "'" & vbNewLine
            Buffer = Buffer & "'*******************************************************************************"
    End Select
    
    rtbCodigo.SelText = Buffer
    Call ColorizeVB(Me.rtbCodigo)
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
    End Select
    
End Sub

Private Sub Timer1_Timer()

    If Not fTip Then
        Timer1.Enabled = False
        Dim ShowAtStartup As Variant
        fTip = True
        ShowAtStartup = GetSetting(App.EXEName, "Opciones", "Mostrar sugerencias al iniciar", 1)
        If ShowAtStartup <> 0 Then
            frmTip.Show vbModal
        End If
        Timer1.Enabled = True
    End If
    
    '// Get first visible line in rtfText
    FirstLine = SendMessage(rtbCodigo.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary
    If Not FirstLineNow = FirstLine Then DrawNumbers '// I can't hook to a scrollbar so I used a sucker-timer
    
End Sub

Private Sub timTimer_Timer()

    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Call MyMsg(brwWebBrowser.LocationName)
    Else
        Call MyMsg("Trabajando...")
    End If
    
End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "cmdEliItem"
            mnuRepositorio_EliminarItem_Click
        Case "cmdAddItem"
            mnuLibreria_AgregarItem_Click
        Case "cmdSave"
            mnuOpciones_Respaldo_Click
        Case "cmdActualizar"
            Call Actualizar(1)
        Case "cmdCortar"
            Call mnuEdicion_Cortar_Click
        Case "cmdCopiar"
            Call mnuEdicion_Copiar_Click
        Case "cmdPegar"
            Call mnuEdicion_Pegar_Click
        Case "cmdBuscar"
            Call mnuEdicion_Buscar_Click
        Case "cmdImprimir"
            Call mnuArchivo_Impresora_Click
        Case "cmdTexto"  'texto
            If GrabarReporte(1) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        Case "cmdRtf"  'rtf
            If GrabarReporte(2) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        Case "cmdHtml"  'html
            If GrabarReporte(3) Then
                MsgBox "Archivo exportado con éxito!", vbInformation
            End If
        Case "cmdBas"
            Call AbrirArchivo("bas")
        Case "cmdCls"
            Call AbrirArchivo("cls")
        Case "cmdForm"
            Call AbrirArchivo("frm")
        Case "cmdCtl"
            Call AbrirArchivo("ctl")
        Case "cmdPag"
            Call AbrirArchivo("pag")
        Case "cmdVbp"
            Call AbrirArchivo("vbp")
        Case "cmdWeb"   'web site
            Call mnuAyuda_WebSite_Click
        Case "cmdAyuda"
            mnuAyuda_Indice_Click
        Case "cmdSalir"
            Unload Me
    End Select
    
End Sub
'exporta el texto a un formato de archivo
Private Function GrabarReporte(ByVal ModoG As Integer) As Boolean

    On Local Error GoTo ErrorGrabarReporte
    
    Dim Archivo As String
    Dim Glosa As String
    Dim Ext As String
    Dim ret As Boolean
    
    ret = False
    
    If Len(Trim$(rtbCodigo.Text)) = 0 Then
        Exit Function
    End If
    
    If ModoG = 1 Then
        Glosa = "Archivos de texto (*.TXT)|*.TXT|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        Ext = "TXT"
    ElseIf ModoG = 2 Then
        Glosa = "Archivos de texto enriquecido (*.RTF)|*.RTF|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        Ext = "RTF"
    ElseIf ModoG = 3 Then
        Glosa = "Archivos de hypertexto (*.HTM)|*.HTM|"
        Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        Ext = "HTM"
    End If
    
    If cc.VBGetSaveFileName(Archivo, , , Glosa, , App.Path, "Guardar reporte como ...", Ext, Me.hWnd) Then
        If Archivo <> "" Then
            If InStr(Archivo, ".") = 0 Then
                Archivo = Archivo & ".rtf"
                Call rtbCodigo.SaveFile(Archivo, rtfRTF)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "TXT" Then
                Call rtbCodigo.SaveFile(Archivo, rtfText)
                ret = True
            ElseIf UCase$(Right$(Archivo, 3)) = "RTF" Then
                Call rtbCodigo.SaveFile(Archivo, rtfRTF)
                ret = True
            Else
                'gsHtml = RichToHTML(Me.rtbcodigo, 0&, Len(rtbcodigo.Text))
                gsHtml = RTF2HTML(rtbCodigo.TextRTF)
                ret = GuardarArchivoHtml(Archivo, Me.Caption)
            End If
        End If
    End If
            
    GoTo SalirGrabarReporte
    
ErrorGrabarReporte:
    ret = False
    MsgBox ("GrabarReporte : " & Err & " " & Error$), vbCritical
    Resume SalirGrabarReporte
    
SalirGrabarReporte:
    GrabarReporte = ret
    Err = 0
        
End Function


'habilitar/deshabilitar tb
Private Sub InhabilitaToolbar(ByVal estado As Boolean)

    Dim k As Integer
            
    mnuArchivo.Enabled = estado
    mnuEdicion.Enabled = estado
    mnuLibreria.Enabled = estado
    mnuOpciones.Enabled = estado
    mnuAyuda.Enabled = estado
    Timer1.Enabled = estado
    
    'esperar a que llegue algun registro
    For k = 1 To tlbMain.Buttons.Count
        tlbMain.Buttons(k).Enabled = estado
    Next k
    
End Sub


Private Sub tmrCodigo_Timer()
    
    If contador > 3 Then    '3 segundos
        Call FormateaCodigo
    Else
        contador = contador + 1
    End If
    
End Sub


'formatear el codigo del rich
Private Sub FormateaCodigo()

    Timer1.Enabled = False
    timTimer.Enabled = False
    tmrCodigo.Enabled = False
    
    contador = 0
    pos = rtbCodigo.SelStart
    Call Hourglass(hWnd, True)
    Call InhabilitaToolbar(False)
    Call MyMsg("Formateando código. Por favor espere ...")
    Call ColorizeVB(Me.rtbCodigo)
    Call MyMsg("Listo")
    Call InhabilitaToolbar(True)
    Call Hourglass(hWnd, False)
    Timer1.Enabled = True
    rtbCodigo.SelStart = pos
    rtbCodigo.SelLength = 0
    On Local Error Resume Next
    rtbCodigo.SetFocus
    Err = 0
    
End Sub


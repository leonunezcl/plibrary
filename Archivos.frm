VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchivos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos"
   ClientHeight    =   6165
   ClientLeft      =   90
   ClientTop       =   2325
   ClientWidth     =   11820
   Icon            =   "Archivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwArchivo 
      Height          =   6120
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10795
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iml"
      SmallIcons      =   "iml"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Archivo"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre Logico"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Origen"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destino"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Diferencias"
         Object.Width           =   7937
      EndProperty
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   7845
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":071E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":09D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":0B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":0C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":0F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Archivos.frx":10A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   6120
      Left            =   0
      ScaleHeight     =   406
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Public tipo_comparacion As Integer
Public glosa_caption As String
Public glosa_comparacion As String
Private Sub Comparar()

    If tipo_comparacion = 2 Then
        Call CompararArchivos
    ElseIf tipo_comparacion = 3 Then
        Call CompararReferencias
    ElseIf tipo_comparacion = 4 Then
        Call CompararComponentes
    End If
    
End Sub

Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call Hourglass(hWnd, True)
    
    Call CenterWindow(hWnd)
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picMain
    End With
        
    Me.Caption = glosa_caption
    
    Call FontStuff(Me.picMain, glosa_comparacion)
        
    Call Comparar
    
    Call Hourglass(hWnd, False)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mGradient = Nothing
    Set frmArchivos = Nothing
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar código en libreria"
   ClientHeight    =   4575
   ClientLeft      =   2475
   ClientTop       =   4005
   ClientWidth     =   7560
   Icon            =   "frmBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwResul 
      Height          =   2490
      Left            =   405
      TabIndex        =   3
      Top             =   2040
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   4392
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgCatItem"
      SmallIcons      =   "imgCatItem"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Categoría"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   10583
      EndProperty
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
      Left            =   6225
      TabIndex        =   5
      Top             =   705
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Buscar"
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
      Left            =   6225
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtBuscar 
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   1485
      Width           =   3855
   End
   Begin VB.ComboBox cboForBus 
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   885
      Width           =   3870
   End
   Begin VB.ComboBox cboCat 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   255
      Width           =   3885
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4545
      Left            =   15
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   6
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.ImageList imgCatItem 
      Left            =   4725
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscar.frx":1CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&Itemes encontrados (doble clic para seleccionar en pantalla principal)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   435
      TabIndex        =   10
      Top             =   1830
      Width           =   5940
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&Digite texto a buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   9
      Top             =   1275
      Width           =   1785
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar &forma de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   8
      Top             =   660
      Width           =   2700
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&Seleccionar categoría"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   7
      Top             =   45
      Width           =   1905
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

'buscar codigo en las categorias
Private Function BuscaCodigo() As Boolean
    
    On Local Error GoTo ErrorBuscaCodigo
    
    Dim ret As Boolean
    Dim k As Integer
    Dim j As Integer
    Dim Texto As String
    Dim Indice As Integer
    Dim TipoBusqueda As Integer
    
    Call Hourglass(hWnd, True)
        
    cmd(0).Enabled = False
    cmd(1).Enabled = False
    
    Indice = cboCat.ListIndex
    TipoBusqueda = cboForBus.ListIndex
    Texto = Trim$(txtBuscar.Text)
    j = 1
    lvwResul.ListItems.Clear
    
    'buscar en todas las categorias ?
    If Indice = 0 Then
        'ciclar x todas las categorias
        For k = 1 To UBound(Arr_Categorias)
            If TipoBusqueda = 0 Then        'exacta
                glbSQL = "SELECT item , descripción from itemes where id = " & k & " and descripción like '%" & Texto & "%'"
            ElseIf TipoBusqueda = 1 Then    'izquierda
                glbSQL = "SELECT item , descripción from itemes where id = " & k & " and descripción like '" & Texto & "%'"
            ElseIf TipoBusqueda = 2 Then    'centro
                glbSQL = "SELECT item , descripción from itemes where id = " & k & " and descripción like '%" & Texto & "%'"
            ElseIf TipoBusqueda = 3 Then    'derecha
                glbSQL = "SELECT item , descripción from itemes where id = " & k & " and descripción like '%" & Texto & "'"
            ElseIf TipoBusqueda = 4 Then    'no esta en ...
                glbSQL = "SELECT item , descripción from itemes where id = " & k & " and descripción <> '" & Texto & "'"
            End If
            
            glbRecordset.Open glbSQL, glbConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            'ciclar x los registros devueltos
            Do While Not glbRecordset.EOF
                lvwResul.ListItems.Add , "k" & j, j, 1, 1
                lvwResul.ListItems("k" & j).SubItems(1) = Trim$(Arr_Categorias(k).Descripcion)
                lvwResul.ListItems("k" & j).SubItems(2) = Trim$(glbRecordset!descripción)
                lvwResul.ListItems("k" & j).Tag = k & "-" & glbRecordset!item
                j = j + 1
                glbRecordset.MoveNext
            Loop
            
            glbRecordset.Close
        Next k
    Else
        'buscar x la categoria seleccionada
        If TipoBusqueda = 0 Then        'exacta
            glbSQL = "SELECT descripción from itemes where id = " & Indice & " and descripción like '%" & Texto & "%'"
        ElseIf TipoBusqueda = 1 Then    'izquierda
            glbSQL = "SELECT descripción from itemes where id = " & Indice & " and descripción like '" & Texto & "%'"
        ElseIf TipoBusqueda = 2 Then    'centro
            glbSQL = "SELECT descripción from itemes where id = " & Indice & " and descripción like '%" & Texto & "%'"
        ElseIf TipoBusqueda = 3 Then    'derecha
            glbSQL = "SELECT descripción from itemes where id = " & Indice & " and descripción like '%" & Texto & "'"
        ElseIf TipoBusqueda = 4 Then    'no esta en ...
            glbSQL = "SELECT descripción from itemes where id = " & Indice & " and descripción <> '" & Texto & "'"
        End If
        
        glbRecordset.Open glbSQL, glbConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        'ciclar x los registros devueltos
        Do While Not glbRecordset.EOF
            lvwResul.ListItems.Add , "k" & j, j, 1, 1
            lvwResul.ListItems("k" & j).SubItems(1) = Trim$(Arr_Categorias(Indice).Descripcion)
            lvwResul.ListItems("k" & j).SubItems(2) = Trim$(glbRecordset!descripción)
            lvwResul.ListItems("k" & j).Tag = Indice & "-" & glbRecordset!item
            j = j + 1
            glbRecordset.MoveNext
        Loop
        
        glbRecordset.Close
    End If
    
    'hay itemes ?
    If lvwResul.ListItems.Count > 0 Then
        ret = True
    Else
        ret = False
        MsgBox "No se encontro código para el texto de búsqueda.", vbCritical
    End If
    
    GoTo SalirBuscaCodigo
    
ErrorBuscaCodigo:
    ret = False
    MsgBox "BuscaCodigo : " & Err & " " & Error$, vbCritical
    Resume SalirBuscaCodigo
    
SalirBuscaCodigo:
    cmd(0).Enabled = True
    cmd(1).Enabled = True
    Call Hourglass(hWnd, False)
    BuscaCodigo = ret
    Err = 0
    
End Function

Private Sub cmd_Click(Index As Integer)

    Dim Msg As String
    
    If Index = 0 Then
        If cboCat.ListIndex <> -1 Then
            If cboForBus.ListIndex <> -1 Then
                If txtBuscar.Text <> "" Then
                    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
                    Msg = "Confirma realizar búsqueda."
                    If Confirma(Msg) = vbYes Then
                        If BuscaCodigo() Then
                            Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
                            MsgBox "Búsqueda realizada con éxito!", vbInformation
                            Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
                        End If
                    End If
                Else
                    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
                    MsgBox "Debe digitar el texto a buscar.", vbCritical
                    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
                    txtBuscar.SetFocus
                End If
            Else
                Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
                MsgBox "Debe seleccionar una forma de búsqueda.", vbCritical
                cboForBus.SetFocus
                Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
            End If
        Else
            Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
            MsgBox "Debe seleccionar una categoría.", vbCritical
            cboCat.SetFocus
            Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    Call CenterWindow(hWnd)
        
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
    
    cboCat.AddItem "Todas las categorias"
    For k = 1 To UBound(Arr_Categorias)
        cboCat.AddItem Arr_Categorias(k).Descripcion
    Next k
    
    cboForBus.AddItem "Exacta"
    cboForBus.AddItem "Esta Izquierda de ..."
    cboForBus.AddItem "Esta Centro de ..."
    cboForBus.AddItem "Esta Derecha de ..."
    cboForBus.AddItem "No esta en ..."
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, Me.Caption)
    
    picDraw.Refresh
    
    Call Hourglass(hWnd, False)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, Me.Left, Me.Top, Me.Width, Me.Height, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmBuscar = Nothing
    
End Sub


Private Sub lvwResul_DblClick()

    Dim Cat As Integer
    Dim item As Integer
    
    If lvwResul.ListItems.Count > 0 Then
        If Not lvwResul.SelectedItem Is Nothing Then
            Cat = Left$(lvwResul.SelectedItem.Tag, InStr(1, lvwResul.SelectedItem.Tag, "-") - 1)
            item = Mid$(lvwResul.SelectedItem.Tag, InStr(1, lvwResul.SelectedItem.Tag, "-") + 1)
                
            frmMain.SetFocus
            frmMain.lvwCat.ListItems(Cat).Selected = True
            frmMain.CargaItemes
            frmMain.lvwItemes.ListItems("k" & item).Selected = True
        End If
    End If
    
End Sub


Private Sub txtBuscar_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmd_Click(0)
    End If
    
End Sub



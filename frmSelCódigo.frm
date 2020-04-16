VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelCodigo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar procedimientos"
   ClientHeight    =   5130
   ClientLeft      =   2085
   ClientTop       =   1770
   ClientWidth     =   7740
   Icon            =   "frmSelCódigo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5130
      Left            =   0
      ScaleHeight     =   340
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   14
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   3735
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCódigo.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCódigo.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCódigo.frx":05C2
            Key             =   ""
         EndProperty
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
      Index           =   5
      Left            =   4290
      TabIndex        =   8
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "A&ceptar"
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
      Index           =   4
      Left            =   2340
      TabIndex        =   7
      Top             =   4680
      Width           =   1200
   End
   Begin VB.ComboBox cboCat 
      Height          =   315
      Left            =   420
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   3030
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Q&uitar"
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
      Index           =   3
      Left            =   3465
      TabIndex        =   5
      Top             =   2175
      Width           =   1200
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Quitar Todas"
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
      Index           =   2
      Left            =   3450
      TabIndex        =   4
      Top             =   1755
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "A&gregar"
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
      Left            =   3450
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwRutinas 
      Height          =   3690
      Left            =   420
      TabIndex        =   1
      Top             =   885
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   6509
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlMain"
      SmallIcons      =   "imlMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rutina"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSelRutinas 
      Height          =   3690
      Left            =   4680
      TabIndex        =   6
      Top             =   885
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   6509
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlMain"
      SmallIcons      =   "imlMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rutina"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Agregar Todas"
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
      Left            =   3450
      TabIndex        =   2
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label lbltotsel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7170
      TabIndex        =   13
      Top             =   690
      Width           =   90
   End
   Begin VB.Label lbltotrut 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   3270
      TabIndex        =   12
      Top             =   705
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar categoria"
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
      Left            =   420
      TabIndex        =   11
      Top             =   105
      Width           =   1875
   End
   Begin VB.Label lblSeleccionadas 
      AutoSize        =   -1  'True
      Caption         =   "Rutinas seleccionadas"
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
      Left            =   4695
      TabIndex        =   10
      Top             =   675
      Width           =   1935
   End
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar rutinas"
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
      Left            =   405
      TabIndex        =   9
      Top             =   660
      Width           =   1650
   End
End
Attribute VB_Name = "frmSelCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private nkey As Integer
'graba rutinas para categoria
Private Function ActualizaCambios() As Boolean

    On Local Error GoTo ErrorActualizaCambios
    
    Dim Msg As String
    Dim Seccion As Integer
    Dim item As Integer
    Dim DescripItem As String
    Dim lineas As Integer
    Dim Linea As String
    Dim Archivo As String
    Dim ret As Boolean
    Dim k As Integer
    Dim c As Integer
    Dim j As Integer
    Dim l As Integer
    Dim total As Integer
    Dim total_itemes As Integer
    
    Dim nFreeFile As Integer
    
    Call Hourglass(hWnd, True)
    Call Habilita(False)
    
    ret = True
    total = lvwSelRutinas.ListItems.Count
    
    Seccion = cboCat.ListIndex + 1
    
    glbConnection.IsolationLevel = adXactReadCommitted
    glbConnection.BeginTrans
    
    'actualizar info en tabla itemes
    total_itemes = frmMain.ContarItemes(Seccion)
    item = total_itemes + 1
            
    'ciclar x todas las rutinas
    For k = 1 To total
        DescripItem = Replace(lvwSelRutinas.ListItems(k).Text, "'", "")
                            
        Archivo = "\" & Arr_Categorias(Seccion).Descripcion & "\"
        Archivo = Archivo & Arr_Categorias(Seccion).Descripcion & "_" & item & ".dat"
    
        'actualizar info en tabla itemes
        glbSQL = "SELECT id , descripción from itemes where "
        glbSQL = glbSQL & "id = " & Seccion
        glbSQL = glbSQL & " and item = " & item
    
        glbRecordset.Open glbSQL, glbConnection
    
        If glbRecordset.EOF Then
            glbSQL = "insert into itemes (id, item, descripción) values ("
            glbSQL = glbSQL & Seccion & " , " & item & " , '" & Left$(DescripItem, 50) & "')"
        Else
            glbSQL = "update itemes set descripción = '" & Left$(DescripItem, 50) & "'"
            glbSQL = glbSQL & " where id = " & Seccion
            glbSQL = glbSQL & " and item = " & item
        End If
    
        glbRecordset.Close
    
        glbConnection.Execute glbSQL
    
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
        nFreeFile = FreeFile
                                        
        Open App.Path & Archivo For Output As #nFreeFile
            For c = 1 To UBound(Mdl)
                For j = 1 To Mdl(c).ProcCount
                    If Mdl(c).Proc(j).Name = lvwSelRutinas.ListItems(k).Text Then
                        For l = 1 To Mdl(c).Proc(j).Lines
                            Print #nFreeFile, Mdl(c).Proc(j).Code(l)
                        Next l
                        Exit For
                    End If
                Next j
            Next c
        Close #nFreeFile
        
        item = item + 1
    Next k
            
    glbConnection.CommitTrans
    
    'contar codigo
    Call frmMain.ContarCodigo
    
    GoTo SalirActualizaCambios
    
ErrorActualizaCambios:
    glbConnection.RollbackTrans
    ret = False
    MsgBox "ActualizaCambios : " & Err & " " & Error$, vbCritical
    Resume SalirActualizaCambios
    
SalirActualizaCambios:
    glbConnection.IsolationLevel = adXactUnspecified
    Call Hourglass(hWnd, False)
    Call Habilita(True)
    ActualizaCambios = ret
    Err = 0
    
End Function
Private Sub Agregar()
    
    Dim k As Integer
    Dim i As Integer
    
    If lvwSelRutinas.ListItems.Count > 0 Then
        If lvwRutinas.ListItems.Count = 0 Then
            Exit Sub
        End If
    End If
    
    Call Hourglass(hWnd, True)
    Call Habilita(False)
    
volver:
    For k = 1 To lvwRutinas.ListItems.Count
        If lvwRutinas.ListItems(k).Selected Then
            If Left$(lvwRutinas.ListItems(k).key, 3) = "sub" Then
                lvwSelRutinas.ListItems.Add , "sub_" & nkey, lvwRutinas.ListItems(k).Text, 1, 1
            ElseIf Left$(lvwRutinas.ListItems(k).key, 3) = "fun" Then
                lvwSelRutinas.ListItems.Add , "fun_" & nkey, lvwRutinas.ListItems(k).Text, 2, 2
            ElseIf Left$(lvwRutinas.ListItems(k).key, 3) = "pro" Then
                lvwSelRutinas.ListItems.Add , "pro_" & nkey, lvwRutinas.ListItems(k).Text, 3, 3
            End If
            
            ValidateRect lvwSelRutinas.hWnd, 0&
            If (i Mod 10) = 0 Then InvalidateRect lvwSelRutinas.hWnd, 0&, 0&
            
            lvwRutinas.ListItems.Remove lvwRutinas.ListItems(k).key
            nkey = nkey + 1
            i = i + 1
            GoTo volver
        End If
    Next k
    
    InvalidateRect lvwSelRutinas.hWnd, 0&, 0&
    
    lbltotrut.Caption = lvwRutinas.ListItems.Count
    lbltotsel.Caption = lvwSelRutinas.ListItems.Count
    
    Call Hourglass(hWnd, False)
    Call Habilita(True)
    
End Sub

'agregar todos los procedimientos
Private Sub AgregarTodas()

    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    
    If lvwSelRutinas.ListItems.Count > 0 Then
        If lvwRutinas.ListItems.Count = 0 Then
            Exit Sub
        End If
    End If
    
    Call Hourglass(hWnd, True)
    Call Habilita(False)
    
    j = lvwSelRutinas.ListItems.Count + 1
    
volver:
    For k = 1 To lvwRutinas.ListItems.Count
        If Left$(lvwRutinas.ListItems(k).key, 3) = "sub" Then
            lvwSelRutinas.ListItems.Add , "sub_" & nkey, lvwRutinas.ListItems(k).Text, 1, 1
        ElseIf Left$(lvwRutinas.ListItems(k).key, 3) = "fun" Then
            lvwSelRutinas.ListItems.Add , "fun_" & nkey, lvwRutinas.ListItems(k).Text, 2, 2
        ElseIf Left$(lvwRutinas.ListItems(k).key, 3) = "pro" Then
            lvwSelRutinas.ListItems.Add , "pro_" & nkey, lvwRutinas.ListItems(k).Text, 3, 3
        End If
            
        ValidateRect lvwSelRutinas.hWnd, 0&
        If (i Mod 10) = 0 Then InvalidateRect lvwSelRutinas.hWnd, 0&, 0&
            
        lvwRutinas.ListItems.Remove lvwRutinas.ListItems(k).key
        i = i + 1
        nkey = nkey + 1
        GoTo volver
    Next k
    
    InvalidateRect lvwSelRutinas.hWnd, 0&, 0&
    
    lvwRutinas.ListItems.Clear
    
    lbltotrut.Caption = lvwRutinas.ListItems.Count
    lbltotsel.Caption = lvwSelRutinas.ListItems.Count
    
    Call Habilita(True)
    Call Hourglass(hWnd, False)
    
End Sub

'habilita botones
Private Sub Habilita(ByVal estado As Boolean)

    Dim k As Integer
    
    For k = 0 To 5
        cmd(k).Enabled = estado
    Next k
    
End Sub


Private Sub Quitar()

    Dim k As Integer
    Dim i As Integer
    
    If lvwRutinas.ListItems.Count > 0 Then
        If lvwSelRutinas.ListItems.Count = 0 Then
            Exit Sub
        End If
    End If
    
    Call Hourglass(hWnd, True)
    Call Habilita(False)
    
volver:
    For k = 1 To lvwSelRutinas.ListItems.Count
        If lvwSelRutinas.ListItems(k).Selected Then
            If Left$(lvwSelRutinas.ListItems(k).key, 3) = "sub" Then
                lvwRutinas.ListItems.Add , "sub_" & nkey, lvwSelRutinas.ListItems(k).Text, 1, 1
            ElseIf Left$(lvwSelRutinas.ListItems(k).key, 3) = "fun" Then
                lvwRutinas.ListItems.Add , "fun_" & nkey, lvwSelRutinas.ListItems(k).Text, 2, 2
            ElseIf Left$(lvwSelRutinas.ListItems(k).key, 3) = "pro" Then
                lvwRutinas.ListItems.Add , "pro_" & nkey, lvwSelRutinas.ListItems(k).Text, 3, 3
            End If
            
            ValidateRect lvwRutinas.hWnd, 0&
            If (i Mod 10) = 0 Then InvalidateRect lvwRutinas.hWnd, 0&, 0&
        
            lvwSelRutinas.ListItems.Remove lvwSelRutinas.ListItems(k).key
            nkey = nkey + 1
            i = i + 1
            GoTo volver
        End If
    Next k
    
    InvalidateRect lvwRutinas.hWnd, 0&, 0&
    
    lbltotrut.Caption = lvwRutinas.ListItems.Count
    lbltotsel.Caption = lvwSelRutinas.ListItems.Count
    
    Call Hourglass(hWnd, False)
    Call Habilita(True)
    
End Sub
'quitar todas
Private Sub QuitarTodas()

    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    
    If lvwRutinas.ListItems.Count > 0 Then
        If lvwSelRutinas.ListItems.Count = 0 Then
            Exit Sub
        End If
    End If
    
    Call Hourglass(hWnd, True)
    Call Habilita(False)
    
volver:
    For k = 1 To lvwSelRutinas.ListItems.Count
        If Left$(lvwSelRutinas.ListItems(k).key, 3) = "sub" Then
            lvwRutinas.ListItems.Add , "sub_" & nkey, lvwSelRutinas.ListItems(k).Text, 1, 1
        ElseIf Left$(lvwSelRutinas.ListItems(k).key, 3) = "fun" Then
            lvwRutinas.ListItems.Add , "fun_" & nkey, lvwSelRutinas.ListItems(k).Text, 2, 2
        ElseIf Left$(lvwSelRutinas.ListItems(k).key, 3) = "pro" Then
            lvwRutinas.ListItems.Add , "pro_" & nkey, lvwSelRutinas.ListItems(k).Text, 3, 3
        End If
        
        ValidateRect lvwRutinas.hWnd, 0&
        If (i Mod 10) = 0 Then InvalidateRect lvwRutinas.hWnd, 0&, 0&
            
        nkey = nkey + 1
        i = i + 1
        lvwSelRutinas.ListItems.Remove lvwSelRutinas.ListItems(k).key
        GoTo volver
    Next k
    
    InvalidateRect lvwRutinas.hWnd, 0&, 0&
    
    lvwSelRutinas.ListItems.Clear
    
    lbltotrut.Caption = lvwRutinas.ListItems.Count
    lbltotsel.Caption = lvwSelRutinas.ListItems.Count
    
    Call Hourglass(hWnd, False)
    Call Habilita(True)
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim Msg As String
    
    Select Case Index
        Case 0  'a. todas
            Call AgregarTodas
        Case 1  'agregar
            Call Agregar
        Case 2  'quitar todas
            Call QuitarTodas
        Case 3  'quitar
            Call Quitar
        Case 4  'aceptar
            If lvwSelRutinas.ListItems.Count > 0 Then
                If cboCat.ListIndex <> -1 Then
                    Msg = "Confirma agregar rutinas."
                    If Confirma(Msg) = vbYes Then
                        If ActualizaCambios() Then
                            MsgBox "Rutinas cargadas con éxito!", vbInformation
                        End If
                    End If
                Else
                    MsgBox "Debes seleccionar una categoria.", vbCritical
                End If
            Else
                MsgBox "No hay rutinas seleccionadas.", vbCritical
            End If
        Case 5  'salir
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    Dim j As Integer
    
    Call Hourglass(hWnd, True)
    
    CenterWindow hWnd
    
    'cargar categorias
    For k = 1 To UBound(Arr_Categorias)
        cboCat.AddItem Arr_Categorias(k).Descripcion
    Next k
    
    'cargar los procedimientos
    nkey = 1
    For k = 1 To UBound(Mdl)
        For j = 1 To Mdl(k).ProcCount
            If Mdl(k).Proc(j).Type = PT_SUB Then
                lvwRutinas.ListItems.Add , "sub_" & nkey, Mdl(k).Proc(j).Name, 1, 1
            ElseIf Mdl(k).Proc(j).Type = PT_FUNCTION Then
                lvwRutinas.ListItems.Add , "fun_" & nkey, Mdl(k).Proc(j).Name, 2, 2
            ElseIf Mdl(k).Proc(j).Type = PT_PROPERTY Then
                lvwRutinas.ListItems.Add , "pro_" & nkey, Mdl(k).Proc(j).Name, 3, 3
            End If
            nkey = nkey + 1
        Next j
    Next k
    
    lbltotrut.Caption = lvwRutinas.ListItems.Count
    lbltotsel.Caption = lvwSelRutinas.ListItems.Count
    
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


Private Sub Form_Unload(Cancel As Integer)
    Set mGradient = Nothing
    Set frmSelCodigo = Nothing
End Sub



VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo item"
   ClientHeight    =   3255
   ClientLeft      =   4320
   ClientTop       =   4200
   ClientWidth     =   5895
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Top             =   0
      Width           =   360
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
      Left            =   3465
      TabIndex        =   3
      Top             =   2745
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
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
      Left            =   1575
      TabIndex        =   2
      Top             =   2745
      Width           =   1215
   End
   Begin VB.TextBox txtDescrip 
      Height          =   2355
      Left            =   420
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   5370
   End
   Begin VB.Label lblGlosa 
      BackColor       =   &H00808080&
      Caption         =   "Descripción"
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   420
      TabIndex        =   1
      Top             =   45
      Width           =   5370
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Public Seccion As Integer
Public Item As Integer
Public key As String
Public tipo As Integer
'Actualiza descripción
Private Function Actualiza() As Boolean

    On Local Error GoTo ErrorActualiza
    
    Dim Descripcion As String
    Dim ret As Boolean
    
    Call Hourglass(hWnd, True)
        
    ret = True
    
    Descripcion = Trim$(txtDescrip.Text)
    
    If Len(Descripcion) > 0 Then
    
        Descripcion = Left$(Replace(Descripcion, "'", ""), 255)
        Descripcion = Left$(Replace(Descripcion, Chr$(10), ""), 255)
        Descripcion = Left$(Replace(Descripcion, Chr$(13), ""), 255)
        Descripcion = Left$(Replace(Descripcion, Chr$(0), ""), 255)
        
        'actualizar info en tabla itemes
        glbSQL = "SELECT id from itemes where "
        glbSQL = glbSQL & "     id = " & Seccion
        glbSQL = glbSQL & " and item = " & Item
        
        glbRecordset.Open glbSQL, glbConnection
        
        If glbRecordset.EOF Then
            glbSQL = "insert into itemes (id, item, descripción) values ("
            glbSQL = glbSQL & Seccion & " , " & Item & " , '" & Descripcion & "')"
        Else
            glbSQL = "update itemes set descripción = '" & Descripcion & "'"
            glbSQL = glbSQL & " where id = " & Seccion
            glbSQL = glbSQL & " and item = " & Item
        End If
        
        glbRecordset.Close
        
        glbConnection.Execute glbSQL
        
        frmMain.lvwItemes.ListItems(key).SubItems(1) = Descripcion
        frmMain.lvwItemes.ListItems(key).Selected = True
    End If
    
    GoTo SalirActualiza
    
ErrorActualiza:
    ret = False
    MsgBox "Actualiza : " & Err & " " & Error$, vbCritical
    Resume SalirActualiza
    
SalirActualiza:
    Err = 0
    Actualiza = ret
    Call Hourglass(hWnd, False)
    
End Function

'ingresa un nuevo registro
Private Function Ingresa() As Boolean

    On Local Error GoTo ErrorIngresa
        
    Dim ret As Boolean
    Dim total As Integer
    Dim total_itemes As Long
    Dim Descripcion As String
    Dim key As String
    Dim LastKey As String
    Dim Item As Integer
    
    ret = True
    
    Call Hourglass(hWnd, True)
    
    Descripcion = Trim$(txtDescrip.Text)
    
    If Len(Descripcion) > 0 Then
        'contar itemes de seccion
        total = frmMain.ContarItemes(Seccion)
        
        'hay codigo ?
        If total > 0 Then
            glbSQL = "select item as cuenta from itemes "
            glbSQL = glbSQL & "where "
            glbSQL = glbSQL & "id = " & Seccion
            glbSQL = glbSQL & " order by item desc"
            
            glbRecordset.Open glbSQL, glbConnection
        
            If Not glbRecordset.EOF Then
                total = glbRecordset!cuenta
            Else
                total = 0
            End If
            
            glbRecordset.Close
        Else
            total = 0
        End If
        total = total + 1
        
        Descripcion = Left$(Replace(Descripcion, "'", ""), 255)
        Descripcion = Left$(Replace(Descripcion, Chr$(10), ""), 255)
        Descripcion = Left$(Replace(Descripcion, Chr$(13), ""), 255)
        Descripcion = Left$(Replace(Descripcion, Chr$(0), ""), 255)
        
        'grabar item nuevo
        If frmMain.GrabaItem(Seccion, total, Descripcion) Then
            'obtener ultimo item de la lista
            
            total_itemes = frmMain.lvwItemes.ListItems.Count
            If total_itemes > 0 Then
                Item = Val(frmMain.lvwItemes.ListItems(frmMain.lvwItemes.ListItems.Count).Text) + 1
                key = "k" & total
            Else
                key = "k" & total
                Item = total
            End If
                                    
            frmMain.lvwItemes.ListItems.Add , key, Format(Item, "0000"), 21, 21
            
            frmMain.lvwItemes.ListItems(key).Tag = Seccion & "-" & total
            frmMain.lvwItemes.ListItems(key).SubItems(1) = Descripcion
            frmMain.lvwItemes.ListItems(key).Selected = True
        End If
        
        frmMain.tabMain.Tabs(1).Caption = "Código de sección : (" & frmMain.lvwItemes.ListItems.Count & ")"
    End If
    
    'contar codigo
    Call frmMain.ContarCodigo
    
    frmMain.rtbCodigo.Text = ""
    frmMain.rtbCodigo.SelColor = RGB(0, 0, 0)
    
    GoTo SalirIngresa
    
ErrorIngresa:
    ret = False
    MsgBox "Ingresa : " & Err & " " & Error$, vbCritical
    Resume SalirIngresa
    
SalirIngresa:
    Err = 0
    Ingresa = ret
    Call Hourglass(hWnd, False)
    
End Function
Private Sub cmd_Click(Index As Integer)

    Dim Msg As String
    
    If Index = 0 Then
        Msg = "Confirma ingresar descripción."
        If Len(Trim$(txtDescrip.Text)) > 0 Then
            If Confirma(Msg) = vbYes Then
                If tipo = 0 Then    'ingresar
                    If Ingresa() Then
                        MsgBox "Información grabada con éxito!", vbInformation
                        Unload Me
                    End If
                Else
                    If Actualiza() Then
                        MsgBox "Información grabada con éxito!", vbInformation
                        Unload Me
                    End If
                End If
            End If
        Else
            MsgBox "Debes ingresar la descripción.", vbCritical
            txtDescrip.SetFocus
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    txtDescrip.SetFocus
End Sub

Private Sub Form_Load()

    CenterWindow hWnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(picDraw, Me.Caption)
    
    picDraw.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mGradient = Nothing
    Set frmNew = Nothing
End Sub

Private Sub txtDescrip_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
    
End Sub


Attribute VB_Name = "mMain"
Option Explicit

Public glbConnection  As New ADODB.Connection
Public glbRecordset As New ADODB.Recordset
Public glbSQL As String
Public gbInicio As Boolean
Public glbCambio As Boolean
Public glbPaginaInicio As String
Private gsBlackKeywords As String
Private gsBlueKeyWords As String
Public glbLinea As String

Public Type eLibreria
    Id As Integer
    Descripcion As String
End Type
Public Arr_Categorias() As eLibreria
'abrir base de datos
Public Function AbrirBaseDatos() As Boolean

    On Local Error GoTo ErrorAbrirBaseDatos
    
    Dim ret As Boolean
    Dim Conexion As String
    
    ret = True
    
    Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                    & "Persist Security Info=False;Data Source=" & App.Path & "\plibrary.mdb"
                    
    glbConnection.Open Conexion
    
    GoTo SalirAbrirBaseDatos
    
ErrorAbrirBaseDatos:
    ret = False
    MsgBox "AbrirBaseDatos : " & Err & " " & Error$, vbCritical
    Resume SalirAbrirBaseDatos
    
SalirAbrirBaseDatos:
    AbrirBaseDatos = ret
    Err = 0
    
End Function

'cargar las secciones definidas en archivo
Public Function CargarLibreria() As Boolean

    On Local Error GoTo ErrorCargarSecciones
    
    Dim ret As Boolean
    Dim k As Integer
    Dim i As Integer
    Dim c As Integer
    
    ret = True
            
    'cargar categorias
    glbSQL = "SELECT id, descripcion FROM categorias order by id asc"
    
    glbRecordset.Open glbSQL, glbConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ReDim Arr_Categorias(0)
    
    'espera
    Call Wait("Cargando libreria. Espere ...", 1, 14)
        
    k = 1
    With glbRecordset
        Do While Not .EOF
            ReDim Preserve Arr_Categorias(k)
            Arr_Categorias(k).Id = !Id
            Arr_Categorias(k).Descripcion = Trim$(!Descripcion)
            
            FRMWait.lblGlosa.Caption = Arr_Categorias(k).Descripcion
            FRMWait.pgb.Value = k
            frmMain.lvwCat.ListItems.Add , "k" & Arr_Categorias(k).Id, Arr_Categorias(k).Descripcion, k, k
                        
            CreaDirectorio App.Path & "\" & Arr_Categorias(k).Descripcion

            k = k + 1
            .MoveNext
        Loop
        .Close
    End With
    
    GoTo SalirCargarSecciones
    
ErrorCargarSecciones:
    ret = False
    MsgBox "CargarLibreria : " & Err & " " & Error$, vbCritical
    Resume SalirCargarSecciones
    
SalirCargarSecciones:
    Unload FRMWait
    CargarLibreria = ret
    Err = 0
    
End Function


'muestra/oculta progress bar principal
Public Sub ShowProgress(ByVal Mode As Boolean)

    frmMain.stbMain.Panels(3).Visible = Mode
    
    If Mode Then
        With frmMain.pgbStatus
            .Left = frmMain.stbMain.Panels(3).Left
            .Top = frmMain.stbMain.Top + 2
            .Width = frmMain.stbMain.Panels(3).Width
            .Height = frmMain.stbMain.Height - 2
            .Visible = True
            .Max = 100
            .Value = 1
            .ZOrder 0
        End With
    Else
        frmMain.pgbStatus.Visible = False
    End If
    
End Sub
Public Sub InitColorize()
' **********************************************************************
' * Comments : Initialize the VB keywords
' *
' *
' **********************************************************************

    gsBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
    gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*Friend*"

End Sub
Public Sub ColorizeVB2(RTF As RichTextBox)
    ' #VBIDEUtils#************************************************************
    ' * Programmer Name : Waty Thierry
    ' * Web Site : http://www.vbdiamond.com
    ' * E-Mail :
    ' * Date : 30/10/98
    ' * Time : 14:47
    ' * Module Name : Colorize_Module
    ' * Module Filename : Colorize.bas
    ' * Procedure Name : ColorizeVB
    ' * Parameters :
    ' * rtf As RichTextBox
    ' **********************************************************************
    ' * Comments : Colorize in black, blue, green the VB keywords
    ' *
    ' *
    ' **********************************************************************
    
    On Local Error Resume Next
    
    Dim sBuffer As String
    Dim nI As Long
    Dim nJ As Long
    Dim sTmpWord As String
    Dim nStartPos As Long
    Dim nSelLen As Long
    Dim nWordPos As Long
            
    If Not glbColorizarCodigo Then
        Exit Sub
    End If
    
    sBuffer = RTF.Text
    sTmpWord = ""
    With RTF
        .Visible = False
        For nI = 1 To Len(sBuffer)
            'DoEvents
            Select Case Mid$(sBuffer, nI, 1)
                Case "A" To "Z", "a" To "z" ', "_"
                    If sTmpWord = "" Then nStartPos = nI
                    sTmpWord = sTmpWord & Mid$(sBuffer, nI, 1)
                
                Case Chr$(34)
                    nSelLen = 1
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI + 1, 1) = Chr$(34) Then
                            nI = nI + 2
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                
                Case Chr$(39)
                    .SelStart = nI - 1
                    nSelLen = 0
                    For nJ = 1 To 9999999
                    'Do
                        If Mid$(sBuffer, nI, 2) = "" Then
                            Exit For
                        ElseIf Mid$(sBuffer, nJ, 2) = vbCrLf Then
                            Exit For
                        'ElseIf Mid$(sBuffer, nJ, 2) = vbNewLine Then
                        '    Exit For
                        '    Exit Do
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    'Loop
                    Next nJ
                    
                    .SelLength = nSelLen
                    .SelColor = RGB(0, 127, 0)
                
                Case Else
                    If Not (Len(sTmpWord) = 0) Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, gsBlackKeywords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 0)
                            .SelText = Mid$(gsBlackKeywords, nWordPos + 1, Len(sTmpWord))
                        End If
                        nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 127)
                            .SelText = Mid$(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                        End If
                        
                        If nWordPos = 0 Then
                            .SelColor = RGB(0, 0, 0)
                        End If
                        
                        If UCase$(sTmpWord) = "REM" Then
                            .SelStart = nI - 4
                            .SelLength = 3
                            For nJ = 1 To 9999999
                                If Mid$(sBuffer, nI, 2) = vbCrLf Then
                                    Exit For
                                Else
                                    .SelLength = .SelLength + 1
                                    nI = nI + 1
                                End If
                            Next
                            .SelColor = RGB(0, 127, 0)
                            .SelText = LCase$(.SelText)
                        End If
                    End If
                    sTmpWord = ""
            End Select
        Next
        .SelStart = 0
        .Visible = True
        .SetFocus
    End With

    Err = 0
    
End Sub

Public Sub ColorizeVB(RTF As RichTextBox)
    ' #VBIDEUtils#************************************************************
    ' * Programmer Name : Waty Thierry
    ' * Web Site : http://www.vbdiamond.com
    ' * E-Mail :
    ' * Date : 30/10/98
    ' * Time : 14:47
    ' * Module Name : Colorize_Module
    ' * Module Filename : Colorize.bas
    ' * Procedure Name : ColorizeVB
    ' * Parameters :
    ' * rtf As RichTextBox
    ' **********************************************************************
    ' * Comments : Colorize in black, blue, green the VB keywords
    ' *
    ' *
    ' **********************************************************************
    
    Dim sBuffer As String
    Dim nI As Long
    Dim nJ As Long
    Dim sTmpWord As String
    Dim nStartPos As Long
    Dim nSelLen As Long
    Dim nWordPos As Long
            
    sBuffer = RTF.Text
    sTmpWord = ""
    With RTF
        For nI = 1 To Len(sBuffer)
            Select Case Mid$(sBuffer, nI, 1)
                Case "A" To "Z", "a" To "z", "_"
                    If sTmpWord = "" Then nStartPos = nI
                    sTmpWord = sTmpWord & Mid$(sBuffer, nI, 1)
                
                Case Chr$(34)
                    nSelLen = 1
                    For nJ = 1 To 10000 '9999999
                        If Mid$(sBuffer, nI + 1, 1) = Chr$(34) Then
                            nI = nI + 2
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                
                Case Chr$(39)
                    .SelStart = nI - 1
                    nSelLen = 0
                    For nJ = 1 To 10000
                        If Mid$(sBuffer, nI, 2) = vbCrLf Then
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                    .SelLength = nSelLen
                    .SelColor = RGB(0, 127, 0)
                
                Case Else
                    If Not (Len(sTmpWord) = 0) Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, gsBlackKeywords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 0)
                            .SelText = Mid$(gsBlackKeywords, nWordPos + 1, Len(sTmpWord))
                        End If
                        nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 127)
                            .SelText = Mid$(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                        End If
                        If UCase$(sTmpWord) = "REM" Then
                            .SelStart = nI - 4
                            .SelLength = 3
                            For nJ = 1 To 10000 '9999999
                                If Mid$(sBuffer, nI, 2) = vbCrLf Then
                                    Exit For
                                Else
                                    .SelLength = .SelLength + 1
                                    nI = nI + 1
                                End If
                            Next
                            .SelColor = RGB(0, 127, 0)
                            .SelText = LCase$(.SelText)
                        End If
                    End If
                    sTmpWord = ""
            End Select
        Next
        .SelStart = 0
    End With

End Sub

'espera ...
Public Sub Wait(ByVal Caption As String, ByVal Minimo As Integer, ByVal Maximo As Integer)

    Load FRMWait
    FRMWait.lblGlosa.Caption = Caption
    FRMWait.pgb.Min = Minimo
    FRMWait.pgb.Max = Maximo
    FRMWait.Show
    DoEvents
    
End Sub



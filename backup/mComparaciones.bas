Attribute VB_Name = "mComparaciones"
Option Explicit

Private Itmx As ListItem

Public arr_ComProyecto() As String
Public arr_ComArchivos() As String
Public arr_ComCodigo() As String
Public glbOpcionComparar As Boolean
Public glbSeComparo As Boolean

Public Type eDif
    Icono As Integer
    Archivo As String
    Ubicacion As String
    Linea As Integer
    DifOrigen As String
    DifDestino As String
    DecOrigen As String
    DecDestino As String
End Type

Public arr_diferencias() As eDif

Private Sub AgregaListaDeDiferencias(ByVal Icono As Integer, _
                                     ByVal Archivo As String, _
                                     ByVal Ubicacion As String, _
                                     ByVal Linea As Long, _
                                     ByVal DifOrigen As String, _
                                     ByVal DifDestino As String, _
                                     Optional ByVal DecOrigen As String = "", _
                                     Optional ByVal DecDestino As String = "")

    Dim c As Integer
    Dim sKey As String
    
    DoEvents
    
    ValidateRect frmMain.lvwProblemas.hWnd, 0&
    
    c = frmMain.lvwProblemas.ListItems.Count + 1
    sKey = "k" & c
        
    frmMain.lvwProblemas.ListItems.Add , sKey, CStr(c), Icono, Icono
    frmMain.lvwProblemas.ListItems(sKey).SubItems(1) = Archivo
    frmMain.lvwProblemas.ListItems(sKey).SubItems(2) = Ubicacion
    frmMain.lvwProblemas.ListItems(sKey).SubItems(3) = Linea
    frmMain.lvwProblemas.ListItems(sKey).SubItems(4) = DifOrigen
    frmMain.lvwProblemas.ListItems(sKey).SubItems(5) = DifDestino
    frmMain.lvwProblemas.ListItems(sKey).SubItems(6) = DecOrigen
    frmMain.lvwProblemas.ListItems(sKey).SubItems(7) = DecDestino
    
    ReDim Preserve arr_diferencias(c)
    arr_diferencias(c).Archivo = Archivo
    arr_diferencias(c).Ubicacion = Ubicacion
    arr_diferencias(c).Linea = Linea
    arr_diferencias(c).DifOrigen = DifOrigen
    arr_diferencias(c).DifDestino = DifDestino
    arr_diferencias(c).DecOrigen = DecOrigen
    arr_diferencias(c).DecDestino = DecDestino
    
    If c Mod 100 = 0 Then
        InvalidateRect frmMain.lvwProblemas.hWnd, 0&, 0&
    End If
    
    'DoEvents
    
End Sub
'carga las opciones de comparacion
Public Sub CargaOpcionesDeComparacion()

    Dim k As Integer
    Dim Valor As String
    
    ReDim arr_ComProyecto(4)
    ReDim arr_ComArchivos(6)
    ReDim arr_ComCodigo(11)
    
    glbSeComparo = False
    
    For k = 1 To 4
        Valor = LeeIni("Proyecto", "Valor" & k, C_INI)
        If Valor = "" Then Valor = "1"
        arr_ComProyecto(k) = Valor
    Next k
    
    For k = 1 To 6
        Valor = LeeIni("Archivos", "Valor" & k, C_INI)
        If Valor = "" Then Valor = "1"
        arr_ComArchivos(k) = Valor
    Next k
    
    For k = 1 To 11
        Valor = LeeIni("Codigo", "Valor" & k, C_INI)
        If Valor = "" Then Valor = "1"
        arr_ComCodigo(k) = Valor
    Next k
    
    Valor = LeeIni("General", "Tipo", C_INI)
    If Valor = "" Then Valor = 0
    glbOpcionComparar = Valor
    
End Sub


'comparar las apis
Private Sub CompararApis(ByVal k As Integer, ByVal UbiOrigen As String, _
                         ByVal j As Integer, ByVal UbiDestino As String, _
                         ByVal Icono As Integer)

    Dim a1 As Integer
    Dim a2 As Integer
    Dim NombreO As String
    Dim NombreApiO As String
    Dim NombreD As String
    Dim NombreApiD As String
    Dim Found As Boolean
    
    Call HelpCarga("Apis ...")
    
    'ciclar x las apis del archivo origen
    For a1 = 1 To UBound(ProyectoO.aArchivos(k).aApis)
        If ProyectoO.aArchivos(k).Explorar Then
            NombreO = ProyectoO.aArchivos(k).aApis(a1).NombreVariable
            NombreApiO = ProyectoO.aArchivos(k).aApis(a1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For a2 = 1 To UBound(ProyectoD.aArchivos(j).aApis)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aApis(a2).NombreVariable
                    NombreApiD = ProyectoD.aArchivos(j).aApis(a2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        
                        If NombreApiO <> NombreApiD Then
                            Call AgregaListaDeDiferencias(C_ICONO_API, UbiOrigen, "Apis", 0, "Declaración Api", "", "<" & NombreApiO & ">", "<" & NombreApiD & ">")
                        End If
                        
                        Exit For
                    End If
                End If
            Next a2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_API, UbiOrigen, "Apis", 0, "Api No Existe en Destino", "", "<" & NombreO & ">", NombreApiO)
            End If
        End If
    Next a1
    
    'ciclar desde el destino al origen
    For a1 = 1 To UBound(ProyectoD.aArchivos(j).aApis)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aApis(a1).NombreVariable
            NombreApiO = ProyectoD.aArchivos(j).aApis(a1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For a2 = 1 To UBound(ProyectoO.aArchivos(k).aApis)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aApis(a2).NombreVariable
                    NombreApiD = ProyectoO.aArchivos(k).aApis(a2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next a2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_API, UbiOrigen, "Apis", 0, "Api No Existe en Origen", "", "", "<" & NombreO & ">")
            End If
        End If
    Next a1
    
    
End Sub

'comparar arreglos
Private Sub CompararArreglos(ByVal k As Integer, ByVal UbiOrigen As String, _
                               ByVal j As Integer, ByVal UbiDestino As String, _
                               ByVal Icono As Integer)

    Dim a1 As Integer
    Dim a2 As Integer
    Dim NombreO As String
    Dim NombreArrO As String
    Dim NombreD As String
    Dim NombreArrD As String
    Dim Found As Boolean
    
    Call HelpCarga("Arrays ...")
    
    'ciclar x los arreglos del archivo origen
    For a1 = 1 To UBound(ProyectoO.aArchivos(k).aArray)
        If ProyectoO.aArchivos(k).Explorar Then
            NombreO = ProyectoO.aArchivos(k).aArray(a1).NombreVariable
            NombreArrO = ProyectoO.aArchivos(k).aArray(a1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For a2 = 1 To UBound(ProyectoD.aArchivos(j).aArray)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aArray(a2).NombreVariable
                    NombreArrD = ProyectoD.aArchivos(j).aArray(a2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        
                        If NombreArrO <> NombreArrD Then
                            Call AgregaListaDeDiferencias(C_ICONO_ARRAY, UbiOrigen, "Arrays", 0, "Declaración Array", "", "<" & NombreArrO & ">", "<" & NombreArrD & ">")
                        End If
                        
                        Exit For
                    End If
                End If
            Next a2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_ARRAY, UbiOrigen, "Arrays", 0, "Array No Existe en Destino", "", "<" & NombreO & ">", "<" & NombreArrO & ">")
            End If
        End If
    Next a1
    
    'ciclar desde el destino al origen
    For a1 = 1 To UBound(ProyectoD.aArchivos(j).aArray)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aArray(a1).NombreVariable
            NombreArrO = ProyectoD.aArchivos(j).aArray(a1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For a2 = 1 To UBound(ProyectoO.aArchivos(k).aArray)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aArray(a2).NombreVariable
                    NombreArrD = ProyectoO.aArchivos(k).aArray(a2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next a2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_ARRAY, UbiOrigen, "Arrays", 0, "Array No Existe en Origen", "", "", "<" & NombreO & ">")
            End If
        End If
    Next a1
        
End Sub

'comparar constantes
Private Sub CompararConstantes(ByVal k As Integer, ByVal UbiOrigen As String, _
                               ByVal j As Integer, ByVal UbiDestino As String, _
                               ByVal Icono As Integer)
    
    Dim c1 As Integer
    Dim c2 As Integer
    Dim NombreO As String
    Dim NombreVariableO As String
    Dim NombreD As String
    Dim NombreVariableD As String
    Dim Found As Boolean
    
    Call HelpCarga("Constantes ...")
    
    'ciclar x las constantes del archivo origen
    For c1 = 1 To UBound(ProyectoO.aArchivos(k).aConstantes)
        If ProyectoO.aArchivos(k).Explorar Then
            NombreO = ProyectoO.aArchivos(k).aConstantes(c1).NombreVariable
            NombreVariableO = ProyectoO.aArchivos(k).aConstantes(c1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For c2 = 1 To UBound(ProyectoD.aArchivos(j).aConstantes)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aConstantes(c2).NombreVariable
                    NombreVariableD = ProyectoD.aArchivos(j).aConstantes(c2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        
                        If NombreVariableO <> NombreVariableD Then
                            Call AgregaListaDeDiferencias(C_ICONO_CONSTANTE, UbiOrigen, "Constantes", 0, "Declaración Constante", "", "<" & NombreVariableO & ">", "<" & NombreVariableD & ">")
                        End If
                        
                        Exit For
                    End If
                End If
            Next c2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_CONSTANTE, UbiOrigen, "Constantes", 0, "Constante No Existe en Destino", "", "<" & NombreO & ">", NombreVariableO)
            End If
        End If
    Next c1
    
    'ciclar desde el destino al origen
    For c1 = 1 To UBound(ProyectoD.aArchivos(j).aConstantes)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aConstantes(c1).NombreVariable
            NombreVariableO = ProyectoD.aArchivos(j).aConstantes(c1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For c2 = 1 To UBound(ProyectoO.aArchivos(k).aConstantes)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aConstantes(c2).NombreVariable
                    NombreVariableD = ProyectoO.aArchivos(k).aConstantes(c2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next c2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_CONSTANTE, UbiOrigen, "Constantes", 0, "Constante No Existe en Origen", "", "", "<" & NombreO & ">")
            End If
        End If
    Next c1
    
End Sub

'comparar las enumeraciones
Private Sub CompararEnumeraciones(ByVal k As Integer, ByVal UbiOrigen As String, _
                                  ByVal j As Integer, ByVal UbiDestino As String, _
                                  ByVal Icono As Integer)

    Dim e1 As Integer
    Dim ee1 As Integer
    Dim e2 As Integer
    Dim ee2 As Integer
    Dim NombreO As String
    Dim NombreEO As String
    Dim NombreEnumO As String
    Dim NombreD As String
    Dim NombreED As String
    Dim NombreEnumD As String
    Dim Found As Boolean
    Dim FoundE As Boolean
    
    Call HelpCarga("Enumeraciones ...")
    
    'ciclar x las enumeraciones del archivo origen
    For e1 = 1 To UBound(ProyectoO.aArchivos(k).aEnumeraciones)
        If ProyectoO.aArchivos(k).Explorar Then
            NombreO = ProyectoO.aArchivos(k).aEnumeraciones(e1).NombreVariable
            NombreEnumO = ProyectoO.aArchivos(k).aEnumeraciones(e1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For e2 = 1 To UBound(ProyectoD.aArchivos(j).aEnumeraciones)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aEnumeraciones(e2).NombreVariable
                    NombreEnumD = ProyectoD.aArchivos(j).aEnumeraciones(e2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        
                        If NombreEnumO <> NombreEnumD Then
                            Call AgregaListaDeDiferencias(C_ICONO_ENUMERACION, UbiOrigen, "Enumeraciones", 0, "Declaración Enumeración", "", "<" & NombreEnumO & ">", "<" & NombreEnumD & ">")
                        End If
                        
                        'comparar el total de elementos de la enumeracion
                        ee1 = UBound(ProyectoO.aArchivos(k).aEnumeraciones(e1).aElementos)
                        ee2 = UBound(ProyectoD.aArchivos(j).aEnumeraciones(e2).aElementos)
                        
                        If ee1 <> ee2 Then
                            Call AgregaListaDeDiferencias(C_ICONO_ENUMERACION, UbiOrigen & "<" & NombreO & ">", "Enumeraciones", 0, "Total elementos", "", "<" & ee1 & ">", "<" & ee2 & ">")
                        End If
                        
                        'comparar los elementos de la enumeraciom
                        For ee1 = 1 To UBound(ProyectoO.aArchivos(k).aEnumeraciones(e1).aElementos)
                            NombreEO = ProyectoO.aArchivos(k).aEnumeraciones(e1).aElementos(ee1).Nombre
                            FoundE = False
                            
                            'ciclar x los elementos de la enumeracion
                            For ee2 = 1 To UBound(ProyectoD.aArchivos(j).aEnumeraciones(e2).aElementos)
                                NombreED = ProyectoD.aArchivos(j).aEnumeraciones(e2).aElementos(ee2).Nombre
                                
                                If LCase$(NombreEO) = LCase$(NombreED) Then
                                    FoundE = True
                                    Exit For
                                End If
                            Next ee2
                            
                            If Not FoundE Then
                                Call AgregaListaDeDiferencias(C_ICONO_ENUMERACION, UbiOrigen & "<" & NombreO & ">", "Enumeraciones", 0, "Elemento Enumeración No Existe en Destino", "", "<" & NombreEO & ">", "")
                            End If
                        Next ee1
                        
                        'comparar elementos desde destino->origen
                        For ee1 = 1 To UBound(ProyectoD.aArchivos(j).aEnumeraciones(e1).aElementos)
                            NombreEO = ProyectoD.aArchivos(j).aEnumeraciones(e1).aElementos(ee1).Nombre
                            FoundE = False
                            
                            'ciclar x los elementos de la enumeracion
                            For ee2 = 1 To UBound(ProyectoO.aArchivos(k).aEnumeraciones(e2).aElementos)
                                NombreED = ProyectoO.aArchivos(k).aEnumeraciones(e2).aElementos(ee2).Nombre
                                
                                If LCase$(NombreEO) = LCase$(NombreED) Then
                                    FoundE = True
                                    Exit For
                                End If
                            Next ee2
                            
                            If Not FoundE Then
                                Call AgregaListaDeDiferencias(C_ICONO_ENUMERACION, UbiOrigen & "<" & NombreO & ">", "Enumeraciones", 0, "Elemento Enumeración No Existe en Origen", "", "", "<" & NombreEO & ">")
                            End If
                        Next ee1
                        
                        Exit For
                    End If
                End If
            Next e2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_ENUMERACION, UbiOrigen, "Enumeraciones", 0, "Enumeración No Existe en Destino", "", "<" & NombreO & ">", NombreEnumO)
            End If
        End If
    Next e1
    
    'ciclar desde el destino al origen
    For e1 = 1 To UBound(ProyectoD.aArchivos(j).aEnumeraciones)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aEnumeraciones(e1).NombreVariable
            NombreEnumO = ProyectoD.aArchivos(j).aEnumeraciones(e1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For e2 = 1 To UBound(ProyectoO.aArchivos(k).aEnumeraciones)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aEnumeraciones(e2).NombreVariable
                    NombreEnumD = ProyectoO.aArchivos(k).aEnumeraciones(e2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next e2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_ENUMERACION, UbiOrigen, "Enumeraciones", 0, "Enumeración No Existe en Origen", "", "", "<" & NombreEnumO & ">")
            End If
        End If
    Next e1

End Sub

'comparar eventos
Private Sub CompararEventos(ByVal k As Integer, ByVal UbiOrigen As String, _
                            ByVal j As Integer, ByVal UbiDestino As String, _
                            ByVal Icono As Integer)

    Dim e1 As Integer
    Dim e2 As Integer
    Dim NombreO As String
    Dim NombreEveO As String
    Dim NombreD As String
    Dim NombreEveD As String
    Dim Found As Boolean
    
    Call HelpCarga("Eventos ...")
    
    'ciclar x los eventos del archivo origen
    For e1 = 1 To UBound(ProyectoO.aArchivos(k).aEventos)
        If ProyectoO.aArchivos(k).Explorar Then
            NombreO = ProyectoO.aArchivos(k).aEventos(e1).NombreVariable
            NombreEveO = ProyectoO.aArchivos(k).aEventos(e1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For e2 = 1 To UBound(ProyectoD.aArchivos(j).aEventos)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aEventos(e2).NombreVariable
                    NombreEveD = ProyectoD.aArchivos(j).aEventos(e2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        
                        If NombreEveO <> NombreEveD Then
                            Call AgregaListaDeDiferencias(C_ICONO_EVENTO, UbiOrigen, "Eventos", 0, "Declaración Evento", "", "<" & NombreEveO & ">", "<" & NombreEveD & ">")
                        End If
                        
                        Exit For
                    End If
                End If
            Next e2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_EVENTO, UbiOrigen, "Eventos", 0, "Evento No Existe en Destino", "", "<" & NombreO & ">", NombreEveO)
            End If
        End If
    Next e1
    
    'ciclar desde el destino al origen
    For e1 = 1 To UBound(ProyectoD.aArchivos(j).aEventos)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aEventos(e1).NombreVariable
            NombreEveO = ProyectoD.aArchivos(j).aEventos(e1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For e2 = 1 To UBound(ProyectoO.aArchivos(k).aEventos)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aEventos(e2).NombreVariable
                    NombreEveD = ProyectoO.aArchivos(k).aEventos(e2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next e2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_EVENTO, UbiOrigen, "Eventos", 0, "Evento No Existe en Origen", "", "", "<" & NombreO & ">")
            End If
        End If
    Next e1
    
End Sub

'comparar subs/funciones/propiedades
Private Sub CompararProcedimientos(ByVal k As Integer, ByVal UbiOrigen As String, _
                                   ByVal j As Integer, ByVal UbiDestino As String, _
                                   ByVal Icono As Integer, ByVal TipoProc As eTipoRutinas)

    Dim r1 As Integer
    Dim rp1 As Integer
    Dim r2 As Integer
    Dim rp2 As Integer
    Dim NombreO As String
    Dim NombrePO As String
    Dim NombreProcO As String
    Dim NombreD As String
    Dim NombrePD As String
    Dim NombreProcD As String
    Dim Found As Boolean
    Dim FoundP As Boolean
    Dim FoundCode As Boolean
    
    Dim r1tld As Integer
    Dim r1tlb As Integer
    Dim r1tlc As Integer
    Dim r1tva As Integer
    
    Dim r2tld As Integer
    Dim r2tlb As Integer
    Dim r2tlc As Integer
    Dim r2tva As Integer
    
    Dim LineaO As String
    Dim LineaD As String
    
    If TipoProc = TIPO_FUN Then
        Call HelpCarga("Funciones ...")
    ElseIf TipoProc = TIPO_PROPIEDAD Then
        Call HelpCarga("Propiedades ...")
    Else
        Call HelpCarga("Procedimientos ...")
    End If
    
    'ciclar x las rutinas del archivo origen
    For r1 = 1 To UBound(ProyectoO.aArchivos(k).aRutinas)
        If ProyectoO.aArchivos(k).Explorar Then
            'comparar las que se piden
            If ProyectoO.aArchivos(k).aRutinas(r1).Tipo = TipoProc Then
                NombreO = ProyectoO.aArchivos(k).aRutinas(r1).NombreRutina
                NombreProcO = ProyectoO.aArchivos(k).aRutinas(r1).Nombre
                r1tld = ProyectoO.aArchivos(k).aRutinas(r1).NumeroDeLineas
                r1tlb = ProyectoO.aArchivos(k).aRutinas(r1).NumeroDeBlancos
                r1tlc = ProyectoO.aArchivos(k).aRutinas(r1).NumeroDeComentarios
                r1tva = UBound(ProyectoO.aArchivos(k).aRutinas(r1).aVariables)
                
                Found = False
                'ciclar x las rutinas del proyecto destino
                For r2 = 1 To UBound(ProyectoD.aArchivos(j).aRutinas)
                    If ProyectoD.aArchivos(j).Explorar Then
                        If ProyectoD.aArchivos(j).aRutinas(r2).Tipo = TipoProc Then
                            NombreD = ProyectoD.aArchivos(j).aRutinas(r2).NombreRutina
                            NombreProcD = ProyectoD.aArchivos(j).aRutinas(r2).Nombre
                            r2tld = ProyectoD.aArchivos(j).aRutinas(r2).NumeroDeLineas
                            r2tlb = ProyectoD.aArchivos(j).aRutinas(r2).NumeroDeBlancos
                            r2tlc = ProyectoD.aArchivos(j).aRutinas(r2).NumeroDeComentarios
                            r2tva = UBound(ProyectoD.aArchivos(j).aRutinas(r2).aVariables)
                        
                            'pille la rutina
                            If LCase$(NombreO) = LCase$(NombreD) Then
                                Found = True
                                
                                If NombreProcO <> NombreProcD Then
                                    If TipoProc = TIPO_FUN Then
                                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Declaración Función", "", "<" & NombreProcO & ">", "<" & NombreProcD & ">")
                                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Declaración Propiedad", "", "<" & NombreProcO & ">", "<" & NombreProcD & ">")
                                    Else
                                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Declaración Procedimiento", "", "<" & NombreProcO & ">", "<" & NombreProcD & ">")
                                    End If
                                End If
                                
                                'comparar el total de lineas de codigo
                                If r1tld <> r2tld Then
                                    If TipoProc = TIPO_FUN Then
                                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Total lineas de codigo", "", "<" & r1tld & ">", "<" & r2tld & ">")
                                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Total lineas de codigo", "", "<" & r1tld & ">", "<" & r2tld & ">")
                                    Else
                                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Total lineas de codigo", "", "<" & r1tld & ">", "<" & r2tld & ">")
                                    End If
                                End If
                                
                                'comparar el total de lineas en blanco
                                If r1tlb <> r2tlb Then
                                    If TipoProc = TIPO_FUN Then
                                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Total lineas en blanco", "", "<" & r1tlb & ">", "<" & r2tlb & ">")
                                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Total lineas en blanco", "", "<" & r1tlb & ">", "<" & r2tlb & ">")
                                    Else
                                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Total lineas en blanco", "", "<" & r1tlb & ">", "<" & r2tlb & ">")
                                    End If
                                End If
                                
                                'comparar el total de lineas de comentarios
                                If r1tlc <> r2tlc Then
                                    If TipoProc = TIPO_FUN Then
                                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Total lineas comentarios", "", "<" & r1tlc & ">", "<" & r2tlc & ">")
                                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Total lineas comentarios", "", "<" & r1tlc & ">", "<" & r2tlc & ">")
                                    Else
                                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Total lineas comentarios", "", "<" & r1tlc & ">", "<" & r2tlc & ">")
                                    End If
                                End If
                                
                                'comparar el total de variables
                                If r1tva <> r2tva Then
                                    If TipoProc = TIPO_FUN Then
                                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Total Variables Fun.", "", "<" & r1tva & ">", "<" & r2tva & ">")
                                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Total Variables Prop.", "", "<" & r1tva & ">", "<" & r2tva & ">")
                                    Else
                                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Total Variables Sub.", "", "<" & r1tva & ">", "<" & r2tva & ">")
                                    End If
                                End If
                                
                                'comparar el total de parametros de
                                rp1 = UBound(ProyectoO.aArchivos(k).aRutinas(r1).Aparams)
                                rp2 = UBound(ProyectoD.aArchivos(j).aRutinas(r2).Aparams)
                                
                                If rp1 <> rp2 Then
                                    If TipoProc = TIPO_FUN Then
                                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen & "<" & NombreO & ">", "Funciones", 0, "Total Parámetros", "", "<" & rp1 & ">", "<" & rp2 & ">")
                                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen & "<" & NombreO & ">", "Propiedades", 0, "Total Parámetros", "", "<" & rp1 & ">", "<" & rp2 & ">")
                                    Else
                                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen & "<" & NombreO & ">", "Procedimientos", 0, "Total Parámetros", "", "<" & rp1 & ">", "<" & rp2 & ">")
                                    End If
                                End If
                                
                                'comparar los parametros
                                For rp1 = 1 To UBound(ProyectoO.aArchivos(k).aRutinas(r1).Aparams)
                                    NombrePO = ProyectoO.aArchivos(k).aRutinas(r1).Aparams(rp1).Nombre
                                    FoundP = False
                                    
                                    'ciclar x los parametros de la rutina
                                    For rp2 = 1 To UBound(ProyectoD.aArchivos(j).aRutinas(r2).Aparams)
                                        NombrePD = ProyectoD.aArchivos(j).aRutinas(r2).Aparams(rp2).Nombre
                                        
                                        If LCase$(NombrePO) = LCase$(NombrePD) Then
                                            FoundP = True
                                            Exit For
                                        End If
                                    Next rp2
                                    
                                    If Not FoundP Then
                                        If TipoProc = TIPO_FUN Then
                                            Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen & "<" & NombreO & ">", "Funciones", 0, "Parametro Función No Existe en Destino", "", "<" & NombrePO & ">", "")
                                        ElseIf TipoProc = TIPO_PROPIEDAD Then
                                            Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen & "<" & NombreO & ">", "Propiedades", 0, "Parametro Propiedad No Existe en Destino", "", "<" & NombrePO & ">", "")
                                        Else
                                            Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen & "<" & NombreO & ">", "Procedimientos", 0, "Parametro Procedimiento No Existe en Destino", "", "<" & NombrePO & ">", "")
                                        End If
                                    End If
                                Next rp1
                                
                                'comparar elementos desde destino->origen
                                For rp2 = 1 To UBound(ProyectoD.aArchivos(j).aRutinas(r2).Aparams)
                                    NombrePO = ProyectoD.aArchivos(j).aRutinas(r2).Aparams(rp2).Nombre
                                    FoundP = False
                                    
                                    'ciclar x los elementos del origen
                                    For rp1 = 1 To UBound(ProyectoO.aArchivos(k).aRutinas(r1).Aparams)
                                        NombrePD = ProyectoO.aArchivos(k).aRutinas(r1).Aparams(rp1).Nombre
                                        
                                        If LCase$(NombrePO) = LCase$(NombrePD) Then
                                            FoundP = True
                                            Exit For
                                        End If
                                    Next rp1
                                    
                                    If Not FoundP Then
                                        If TipoProc = TIPO_FUN Then
                                            Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen & "<" & NombreO & ">", "Funciones", 0, "Parametro Función No Existe en Origen", "", "", "<" & NombrePO & ">")
                                        ElseIf TipoProc = TIPO_PROPIEDAD Then
                                            Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen & "<" & NombreO & ">", "Propiedades", 0, "Parametro Propiedad No Existe en Origen", "", "", "<" & NombrePO & ">")
                                        Else
                                            Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen & "<" & NombreO & ">", "Procedimientos", 0, "Parametro Procedimiento No Existe en Origen", "", "", "<" & NombrePO & ">")
                                        End If
                                    End If
                                Next rp2
                                
                                'comparar el codigo
                                For rp1 = 1 To UBound(ProyectoO.aArchivos(k).aRutinas(r1).aCodigoRutina)
                                    LineaO = Trim$(ProyectoO.aArchivos(k).aRutinas(r1).aCodigoRutina(rp1).Codigo)
                                    FoundP = False
                                    
                                    'comparar contra el destino
                                    If glbOpcionComparar Then   'comparar linea x linea
                                        For rp2 = 1 To UBound(ProyectoD.aArchivos(j).aRutinas(r2).aCodigoRutina)
                                            LineaD = Trim$(ProyectoD.aArchivos(j).aRutinas(r2).aCodigoRutina(rp2).Codigo)
                                            If LCase$(LineaO) <> LCase$(LineaD) Then
                                                If TipoProc = TIPO_FUN Then
                                                    Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen & "<" & NombreO & ">", "Funciones", 0, "Código Función", "", "<" & LineaO & ">", "<" & LineaD & ">")
                                                ElseIf TipoProc = TIPO_PROPIEDAD Then
                                                    Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen & "<" & NombreO & ">", "Propiedades", 0, "Código Propiedad", "", "<" & LineaO & ">", "<" & LineaD & ">")
                                                Else
                                                    Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen & "<" & NombreO & ">", "Procedimientos", 0, "Código Procedimiento", "", "<" & LineaO & ">", "<" & LineaD & ">")
                                                End If
                                            End If
                                        Next rp2
                                    Else
                                        'buscar la linea en el codigo
                                        FoundCode = False
                                        For rp2 = 1 To UBound(ProyectoD.aArchivos(j).aRutinas(r2).aCodigoRutina)
                                            LineaD = Trim$(ProyectoD.aArchivos(j).aRutinas(r2).aCodigoRutina(rp2).Codigo)
                                            If LCase$(LineaO) = LCase$(LineaD) Then
                                                FoundCode = True
                                                Exit For
                                            End If
                                        Next rp2
                                        
                                        If Not FoundCode Then
                                            If TipoProc = TIPO_FUN Then
                                                Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen & "<" & NombreO & ">", "Funciones", 0, "Código Función", "", "<" & LineaO & ">", "")
                                            ElseIf TipoProc = TIPO_PROPIEDAD Then
                                                Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen & "<" & NombreO & ">", "Propiedades", 0, "Código Propiedad", "", "<" & LineaO & ">", "")
                                            Else
                                                Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen & "<" & NombreO & ">", "Procedimientos", 0, "Código Procedimiento", "", "<" & LineaO & ">", "")
                                            End If
                                        End If
                                    End If
                                Next rp1
                                Exit For
                            End If
                        End If
                    End If
                Next r2
                
                If Not Found Then
                    If TipoProc = TIPO_FUN Then
                        Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Función No Existe en Destino", "", "<" & NombreO & ">", NombreProcO)
                    ElseIf TipoProc = TIPO_PROPIEDAD Then
                        Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Propiedad No Existe en Destino", "", "<" & NombreO & ">", NombreProcO)
                    Else
                        Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Procedimiento No Existe en Destino", "", "<" & NombreO & ">", NombreProcO)
                    End If
                End If
            End If
        End If
    Next r1
    
    'ciclar desde el destino al origen
    For r1 = 1 To UBound(ProyectoD.aArchivos(j).aRutinas)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aRutinas(r1).NombreRutina
            NombreProcO = ProyectoD.aArchivos(j).aRutinas(r1).Nombre
            
            Found = False
            'ciclar x el proyecto origen
            For r2 = 1 To UBound(ProyectoO.aArchivos(k).aRutinas)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aRutinas(r2).NombreRutina
                    NombreProcD = ProyectoO.aArchivos(k).aRutinas(r2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next r2
            
            If Not Found Then
                If TipoProc = TIPO_FUN Then
                    Call AgregaListaDeDiferencias(C_ICONO_FUNCION, UbiOrigen, "Funciones", 0, "Función No Existe en Origen", "", "", "<" & NombreProcO & ">")
                ElseIf TipoProc = TIPO_PROPIEDAD Then
                    Call AgregaListaDeDiferencias(C_ICONO_PROPIEDAD_PUBLICA, UbiOrigen, "Propiedades", 0, "Propiedad No Existe en Origen", "", "", "<" & NombreProcO & ">")
                Else
                    Call AgregaListaDeDiferencias(C_ICONO_SUB, UbiOrigen, "Procedimientos", 0, "Procedimiento No Existe en Origen", "", "", "<" & NombreProcO & ">")
                End If
            End If
        End If
    Next r1
    
End Sub

'realiza las comparaciones entre proyectos
Public Function CompararProyectosSeleccionados() As Boolean

    On Local Error GoTo ErrorCompararProyectosSeleccionados
    
    Dim ret As Boolean
    Dim k As Integer
    Dim j As Integer
    
    ret = True
    
    Call Hourglass(frmMain.hWnd, True)
    
    frmMain.lvwProblemas.ListItems.Clear
    frmMain.lblDif.Caption = frmMain.lblDif.Tag
    
    Call HelpCarga("Comparando ...")
    
    'comparaciones a nivel de proyecto
    ReDim arr_diferencias(0)
    
    'informacion de proyecto
    If arr_ComProyecto(1) = 1 Then
        Call DiferenciaInformacionProyecto
    End If
    
    'informacion de componentes
    If arr_ComProyecto(2) = 1 Then
        Call DiferenciaInformacionComponentes
    End If
    
    'informacion de referencias
    If arr_ComProyecto(3) = 1 Then
        Call DiferenciaInformacionReferencias
    End If
    
    'informacion de archivos
    If arr_ComProyecto(4) = 1 Then
        Call DiferenciaInformacionArchivos
    End If
    
    'comparar formularios
    If arr_ComArchivos(1) = 1 Then
        Call DiferenciaInformacionArchivosProyecto(TIPO_ARCHIVO_FRM, C_ICONO_FORM)
    End If
    
    'comparar modulos .bas
    If arr_ComArchivos(2) = 1 Then
        Call DiferenciaInformacionArchivosProyecto(TIPO_ARCHIVO_BAS, C_ICONO_BAS)
    End If
    
    'comparar modulos .cls
    If arr_ComArchivos(3) = 1 Then
        Call DiferenciaInformacionArchivosProyecto(TIPO_ARCHIVO_CLS, 3)
    End If
    
    'comparar controles
    If arr_ComArchivos(4) = 1 Then
        Call DiferenciaInformacionArchivosProyecto(TIPO_ARCHIVO_OCX, 4)
    End If
    
    'comparar paginas
    If arr_ComArchivos(5) = 1 Then
        Call DiferenciaInformacionArchivosProyecto(TIPO_ARCHIVO_PAG, 5)
    End If
    
    'comparar diseñadores
    If arr_ComArchivos(6) = 1 Then
        Call DiferenciaInformacionArchivosProyecto(TIPO_ARCHIVO_DSR, 6)
    End If
    
    GoTo SalirCompararProyectosSeleccionados
    
ErrorCompararProyectosSeleccionados:
    ret = False
    MsgBox "CompararProyectosSeleccionados : " & Err & " " & Error$, vbCritical
    Resume SalirCompararProyectosSeleccionados
    
SalirCompararProyectosSeleccionados:
    Call Hourglass(frmMain.hWnd, False)
    CompararProyectosSeleccionados = ret
    If ret Then
        frmMain.lblDif.Caption = frmMain.lvwProblemas.ListItems.Count & " " & frmMain.lblDif.Caption
    End If
    Call HelpCarga("Listo")
    glbSeComparo = ret
    Err = 0
    
End Function

'compara la seccion general del archivo origen->destino
Private Sub CompararSeccionGeneral(ByVal k As Integer, ByVal UbiOrigen As String, _
                                   ByVal j As Integer, ByVal UbiDestino As String, _
                                   ByVal Icono As Integer)

    Dim g1 As Integer
    Dim g2 As Integer
    Dim Linea1 As String
    Dim Linea2 As String
    Dim Found As Boolean
    
    Call HelpCarga("Sección generales ...")
    
    'desde origen->destino
    For g1 = 1 To UBound(ProyectoO.aArchivos(k).aGeneral)
        If ProyectoO.aArchivos(k).Explorar Then
            Linea1 = Trim$(ProyectoO.aArchivos(k).aGeneral(g1).Codigo)
            
            'comparar linea a linea ?
            If glbOpcionComparar Then
                If g1 <= UBound(ProyectoD.aArchivos(j).aGeneral) Then
                    Linea2 = Trim$(ProyectoD.aArchivos(j).aGeneral(g1).Codigo)
                    
                    'comparar linea x linea
                    If Linea1 <> Linea2 Then
                        Call AgregaListaDeDiferencias(Icono, UbiOrigen, "Generales", g1, "General linea", "", Linea1, Linea2)
                    End If
                Else
                    Call AgregaListaDeDiferencias(Icono, UbiOrigen, "Generales", g1, "Linea General No Existe en Destino", "", Linea1, "")
                End If
            Else
                Found = False
                'buscar la cadena en la seccion general destino
                For g2 = 1 To UBound(ProyectoD.aArchivos(j).aGeneral)
                    Linea2 = Trim$(ProyectoD.aArchivos(j).aGeneral(g2).Codigo)
                    
                    If Linea1 = Linea2 Then
                        Found = True
                        Exit For
                    End If
                Next g2
                
                If Not Found Then
                    Call AgregaListaDeDiferencias(Icono, UbiOrigen, "Generales", g1, "Linea General No Existe en Destino", "", Linea1, "")
                End If
            End If
        End If
    Next g1
    
End Sub
'comparar tipos
Private Sub CompararTipos(ByVal k As Integer, ByVal UbiOrigen As String, _
                          ByVal j As Integer, ByVal UbiDestino As String, _
                          ByVal Icono As Integer)

    Dim t1 As Integer
    Dim et1 As Integer
    Dim t2 As Integer
    Dim et2 As Integer
    Dim NombreO As String
    Dim NombreEO As String
    Dim NombreTipoO As String
    Dim NombreD As String
    Dim NombreED As String
    Dim NombreTipoD As String
    Dim Found As Boolean
    Dim FoundE As Boolean
    
    Call HelpCarga("Tipos ...")
    
    'ciclar x los tipos del archivo origen
    For t1 = 1 To UBound(ProyectoO.aArchivos(k).aTipos)
        If ProyectoO.aArchivos(k).Explorar Then
            NombreO = ProyectoO.aArchivos(k).aTipos(t1).NombreVariable
            NombreTipoO = ProyectoO.aArchivos(k).aTipos(t1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For t2 = 1 To UBound(ProyectoD.aArchivos(j).aTipos)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aTipos(t2).NombreVariable
                    NombreTipoD = ProyectoD.aArchivos(j).aTipos(t2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        
                        If NombreTipoO <> NombreTipoD Then
                            Call AgregaListaDeDiferencias(C_ICONO_TIPOS, UbiOrigen, "Tipos", 0, "Declaración Tipo", "", "<" & NombreTipoO & ">", "<" & NombreTipoD & ">")
                        End If
                        
                        'comparar el total de elementos del tipo
                        et1 = UBound(ProyectoO.aArchivos(k).aTipos(t1).aElementos)
                        et2 = UBound(ProyectoD.aArchivos(j).aTipos(t2).aElementos)
                        
                        If et1 <> et2 Then
                            Call AgregaListaDeDiferencias(C_ICONO_TIPOS, UbiOrigen & "<" & NombreO & ">", "Tipos", 0, "Total elementos", "", "<" & et1 & ">", "<" & et2 & ">")
                        End If
                        
                        'comparar los elementos de los tipos
                        For et1 = 1 To UBound(ProyectoO.aArchivos(k).aTipos(t1).aElementos)
                            NombreEO = ProyectoO.aArchivos(k).aTipos(t1).aElementos(et1).Nombre
                            FoundE = False
                            
                            'ciclar x los elementos del tipo destino
                            For et2 = 1 To UBound(ProyectoD.aArchivos(j).aTipos(t2).aElementos)
                                NombreED = ProyectoD.aArchivos(j).aTipos(t2).aElementos(et2).Nombre
                                
                                If LCase$(NombreEO) = LCase$(NombreED) Then
                                    FoundE = True
                                    Exit For
                                End If
                            Next et2
                            
                            If Not FoundE Then
                                Call AgregaListaDeDiferencias(C_ICONO_TIPOS, UbiOrigen & "<" & NombreO & ">", "Tipos", 0, "Elemento Tipo No Existe en Destino", "", "<" & NombreEO & ">", "")
                            End If
                        Next et1
                        
                        'comparar los elementos de tipos desde destino->origen
                        For et1 = 1 To UBound(ProyectoD.aArchivos(j).aTipos(t1).aElementos)
                            NombreEO = ProyectoD.aArchivos(j).aTipos(t1).aElementos(et1).Nombre
                            FoundE = False
                            
                            'ciclar x los elementos del tipo destino
                            For et2 = 1 To UBound(ProyectoO.aArchivos(k).aTipos(t2).aElementos)
                                NombreED = ProyectoO.aArchivos(k).aTipos(t2).aElementos(et2).Nombre
                                
                                If LCase$(NombreEO) = LCase$(NombreED) Then
                                    FoundE = True
                                    Exit For
                                End If
                            Next et2
                            
                            If Not FoundE Then
                                Call AgregaListaDeDiferencias(C_ICONO_TIPOS, UbiOrigen & "<" & NombreO & ">", "Tipos", 0, "Elemento Tipo No Existe en Destino", "", "", "<" & NombreEO & ">")
                            End If
                        Next et1
                        Exit For
                    End If
                End If
            Next t2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_TIPOS, UbiOrigen, "Tipos", 0, "Tipos No Existe en Destino", "", "<" & NombreO & ">", NombreTipoO)
            End If
        End If
    Next t1
    
    'ciclar desde el destino al origen
    For t1 = 1 To UBound(ProyectoD.aArchivos(j).aTipos)
        If ProyectoD.aArchivos(j).Explorar Then
            NombreO = ProyectoD.aArchivos(j).aTipos(t1).NombreVariable
            NombreTipoO = ProyectoD.aArchivos(j).aTipos(t1).Nombre
            
            Found = False
            'ciclar x el proyecto destino
            For t2 = 1 To UBound(ProyectoO.aArchivos(k).aTipos)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aTipos(t2).NombreVariable
                    NombreTipoD = ProyectoO.aArchivos(k).aTipos(t2).Nombre
                    
                    If LCase$(NombreO) = LCase$(NombreD) Then
                        Found = True
                        Exit For
                    End If
                End If
            Next t2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_TIPOS, UbiOrigen, "Tipos", 0, "Tipo No Existe en Origen", "", "", "<" & NombreTipoO & ">")
            End If
        End If
    Next t1

End Sub

'compara las variables declaradas en la seccion general
Private Sub CompararVariables(ByVal k As Integer, ByVal UbiOrigen As String, _
                              ByVal j As Integer, ByVal UbiDestino As String, _
                              ByVal Icono As Integer)

    Dim v1 As Integer
    Dim v2 As Integer
    Dim NombreO As String
    Dim NombreVariableO As String
    Dim TipoO As String
    Dim AlcanceO As Boolean
    Dim BasicStyleO As Boolean
    Dim VariantO As Boolean
    Dim UsaDimO As Boolean
    Dim UsaPrivateO As Boolean
    Dim UsaGlobalO As Boolean
    
    Dim NombreD As String
    Dim NombreVariableD As String
    Dim TipoD As String
    Dim AlcanceD As Boolean
    Dim BasicStyleD As Boolean
    Dim VariantD As Boolean
    Dim UsaDimD As Boolean
    Dim UsaPrivateD As Boolean
    Dim UsaGlobalD As Boolean
    
    Dim Found As Boolean
    Dim Origen As String
    Dim Destino As String
    
    Call HelpCarga("Variables ...")
    
    'ciclar desde origen->destino
    For v1 = 1 To UBound(ProyectoO.aArchivos(k).aVariables)
        If ProyectoO.aArchivos(k).Explorar Then
            'informacion variables origen
            NombreO = ProyectoO.aArchivos(k).aVariables(v1).NombreVariable
            TipoO = ProyectoO.aArchivos(k).aVariables(v1).Tipo
            AlcanceO = ProyectoO.aArchivos(k).aVariables(v1).Publica
            BasicStyleO = ProyectoO.aArchivos(k).aVariables(v1).BasicOldStyle
            VariantO = ProyectoO.aArchivos(k).aVariables(v1).Predefinido
            UsaDimO = ProyectoO.aArchivos(k).aVariables(v1).UsaDim
            UsaPrivateO = ProyectoO.aArchivos(k).aVariables(v1).UsaPrivate
            UsaGlobalO = ProyectoO.aArchivos(k).aVariables(v1).UsaGlobal
            
            Found = False
            
            'informacion variables destino
            For v2 = 1 To UBound(ProyectoD.aArchivos(j).aVariables)
                If ProyectoD.aArchivos(j).Explorar Then
                    NombreD = ProyectoD.aArchivos(j).aVariables(v2).NombreVariable
                    TipoD = ProyectoD.aArchivos(j).aVariables(v2).Tipo
                    AlcanceD = ProyectoD.aArchivos(j).aVariables(v2).Publica
                    BasicStyleD = ProyectoD.aArchivos(j).aVariables(v2).BasicOldStyle
                    VariantD = ProyectoD.aArchivos(j).aVariables(v2).Predefinido
                    UsaDimD = ProyectoD.aArchivos(j).aVariables(v2).UsaDim
                    UsaPrivateD = ProyectoD.aArchivos(j).aVariables(v2).UsaPrivate
                    UsaGlobalD = ProyectoD.aArchivos(j).aVariables(v2).UsaGlobal
                    
                    'son iguales
                    If NombreO = NombreD Then
                        Found = True
                        
                        'comparar propiedades de variables
                        If TipoO <> TipoD Then
                            Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Tipo de Variable", "", "<" & TipoO & ">", "<" & TipoD & ">")
                        End If
                        
                        'If AlcanceO <> AlcanceD Then
                        '    Origen = "Alcance Variable Origen : <" & AlcanceO & "> en destino : <" & AlcanceD & ">"
                        '    Call AgregaListaDeDiferencias(UbiOrigen & "." & NombreO, Origen, C_ICONO_DIM)
                        'End If
                        
                        If BasicStyleO <> BasicStyleD Then
                            Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Estilo Basic", "", "<" & BasicStyleO & ">", "<" & BasicStyleD & ">")
                        End If
                                                        
                        If VariantO <> VariantD Then
                            Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Variant Predefinido", "", "<" & VariantO & ">", "<" & VariantD & ">")
                        End If
                        
                        If UsaDimO <> UsaDimD Then
                            Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Declaración con Dim", "", "<" & UsaDimO & ">", "<" & UsaDimD & ">")
                        End If
                        
                        If UsaPrivateO <> UsaPrivateD Then
                            Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Declaración con Private", "", "<" & UsaPrivateO & ">", "<" & UsaPrivateD & ">")
                        End If
                        
                        If UsaGlobalO <> UsaGlobalD Then
                            Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Declaración con Global", "", "<" & UsaGlobalO & ">", "<" & UsaGlobalD & ">")
                        End If
                        Exit For
                    End If
                End If
            Next v2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Variable No Existe en Destino", "", "<" & NombreO & ">", "")
            End If
        End If
    Next v1
    
    'ciclar desde destino->origen
    For v1 = 1 To UBound(ProyectoD.aArchivos(j).aVariables)
        If ProyectoD.aArchivos(j).Explorar Then
            'informacion variables origen
            NombreO = ProyectoD.aArchivos(j).aVariables(v1).NombreVariable
            NombreVariableO = ProyectoD.aArchivos(j).aVariables(v1).Nombre
            Found = False
            
            'informacion variables destino
            For v2 = 1 To UBound(ProyectoO.aArchivos(k).aVariables)
                If ProyectoO.aArchivos(k).Explorar Then
                    NombreD = ProyectoO.aArchivos(k).aVariables(v2).NombreVariable
                    NombreVariableD = ProyectoO.aArchivos(k).aVariables(v2).Nombre
                    
                    'son iguales
                    If NombreO = NombreD Then
                        Found = True
                        Exit For
                    End If
                End If
            Next v2
            
            If Not Found Then
                Call AgregaListaDeDiferencias(C_ICONO_DIM, UbiOrigen & "<" & NombreO & ">", "Variables", 0, "Variable No Existe en Destino", "", "", "<" & NombreO & ">")
            End If
        End If
    Next v1
    
End Sub

'comprueba las diferencias entre archivos
Private Sub DiferenciaInformacionArchivos()

    Dim k As Integer
    Dim j As Integer
    Dim c As Integer
    Dim Found As Boolean
    Dim fSize As Boolean
    Dim fTime As Boolean
    Dim Icono As Integer
    Dim ArchivoO As String
    Dim ArchivoD As String
    Dim FileSizeO As Double
    Dim FileSizeD As Double
    Dim FileTimeO As String
    Dim FileTimeD As String
    Dim Nombre As String
    
    Call HelpCarga("Archivos ...")
    
    'ciclar x los archivos del proyecto origen - destino
    'comprobar si existen
    c = 1
    For k = 1 To UBound(ProyectoO.aArchivos)
        If ProyectoO.aArchivos(k).Explorar Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoO.aArchivos(k).PathFisico))
            Nombre = ProyectoO.aArchivos(k).ObjectName
            FileSizeO = ProyectoO.aArchivos(k).FileSize
            FileTimeO = ProyectoO.aArchivos(k).FILETIME
        
            'verificar el tipo de archivo
            If ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Icono = C_ICONO_FORM
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Icono = C_ICONO_BAS
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Icono = C_ICONO_CLS
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Icono = C_ICONO_CONTROL
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Icono = C_ICONO_PAGINA
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                Icono = C_ICONO_DESIGNER
            End If
        
            Found = False
            fSize = False
            fTime = False
            
            'ciclar x el proyecto destino
            For j = 1 To UBound(ProyectoD.aArchivos)
                If ProyectoD.aArchivos(j).Explorar Then
                    ArchivoD = LCase$(VBArchivoSinPath(ProyectoD.aArchivos(j).PathFisico))
                    FileSizeD = ProyectoD.aArchivos(j).FileSize
                    FileTimeD = ProyectoD.aArchivos(j).FILETIME
                    
                    If ArchivoO = ArchivoD Then
                        'comparar hora y fecha de modificacion
                        If FileSizeO <> FileSizeD Then fSize = True
                        If FileTimeO <> FileTimeD Then fTime = True
                        
                        Found = True
                        Exit For
                    End If
                End If
            Next j
        
            'archivo no existe en proyecto destino ?
            If Not Found Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Archivos", 0, "Archivo no existe en destino", "", "<" & ArchivoO & ">", "")
            End If
        
            'tamaño
            If fSize Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Archivos", 0, "Tamaño Archivo", "<" & ArchivoO & ">", "<" & FileSizeO & ">", "<" & FileSizeD & ">")
            End If
            
            'tamaño
            If fTime Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Archivos", 0, "Fecha Ultima Modificación", "<" & ArchivoO & ">", "<" & FileTimeO & ">", "<" & FileTimeD & ">")
            End If
        End If
    Next k
    
    'ciclar de destino a origen
    For k = 1 To UBound(ProyectoD.aArchivos)
        If ProyectoD.aArchivos(k).Explorar Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoD.aArchivos(k).PathFisico))
            Nombre = ProyectoD.aArchivos(k).ObjectName
            FileSizeO = ProyectoD.aArchivos(k).FileSize
            FileTimeO = ProyectoD.aArchivos(k).FILETIME
        
            'verificar el tipo de archivo
            If ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
                Icono = C_ICONO_FORM
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
                Icono = C_ICONO_BAS
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
                Icono = C_ICONO_CLS
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
                Icono = C_ICONO_CONTROL
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_PAG Then
                Icono = C_ICONO_PAGINA
            ElseIf ProyectoO.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_DSR Then
                Icono = C_ICONO_DESIGNER
            End If
        
            Found = False
            fSize = False
            fTime = False
        
            'ciclar x el proyecto origen
            For j = 1 To UBound(ProyectoO.aArchivos)
                If ProyectoO.aArchivos(k).Explorar Then
                    ArchivoD = LCase$(VBArchivoSinPath(ProyectoO.aArchivos(j).PathFisico))
                    FileSizeD = ProyectoO.aArchivos(j).FileSize
                    FileTimeD = ProyectoO.aArchivos(j).FILETIME
                    
                    If ArchivoO = ArchivoD Then
                        'comparar hora y fecha de modificacion
                        If FileSizeO <> FileSizeD Then fSize = True
                        If FileTimeO <> FileTimeD Then fTime = True
                        
                        Found = True
                        Exit For
                    End If
                End If
            Next j
        
            'archivo no existe en proyecto destino ?
            If Not Found Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Archivos", 0, "", "Archivo No Existe en Origen.", "", ArchivoO)
            End If
        End If
    Next k
    
End Sub

'diferencias a nivel de componentes .ocx
Private Sub DiferenciaInformacionComponentes()

    Dim k As Integer
    Dim j As Integer
    Dim c As Integer
    Dim Icono As Integer
    
    Dim Found As Boolean
    Dim fSize As Boolean
    Dim fTime As Boolean
    Dim fVersion As Boolean
    Dim fDescrip As Boolean
    Dim fGuid As Boolean
    Dim fNombre As Boolean
                
    Dim ArchivoO As String
    Dim ArchivoD As String
    Dim FileSizeO As Double
    Dim FileSizeD As Double
    Dim FileTimeO As String
    Dim FileTimeD As String
    Dim VersionO As String
    Dim VersionD As String
    Dim DescripO As String
    Dim DescripD As String
    Dim GuidO As String
    Dim GuidD As String
    Dim NombreO As String
    Dim NombreD As String
        
    'ciclar x los archivos del proyecto origen
    'comprobar si existen
    Call HelpCarga("Componentes ...")
    
    c = 1
    For k = 1 To UBound(ProyectoO.aDepencias)
        'seleccionar solo aquellas referencia de tipo dll
        If ProyectoO.aDepencias(k).Tipo = TIPO_OCX Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoO.aDepencias(k).ContainingFile))
            NombreO = ProyectoO.aDepencias(k).Name
            FileSizeO = ProyectoO.aDepencias(k).FileSize
            FileTimeO = ProyectoO.aDepencias(k).FILETIME
            GuidO = ProyectoO.aDepencias(k).GUID
            DescripO = ProyectoO.aDepencias(k).HelpString
            VersionO = ProyectoO.aDepencias(k).MajorVersion & "." & ProyectoO.aDepencias(k).MinorVersion
                
            Icono = C_ICONO_OCX
                
            'flags para comparar propiedades
            Found = False
            fSize = False
            fTime = False
            fVersion = False
            fDescrip = False
            fGuid = False
            fNombre = False
        
            'ciclar x el proyecto destino
            For j = 1 To UBound(ProyectoD.aDepencias)
                If ProyectoD.aDepencias(j).Tipo = TIPO_OCX Then
                    ArchivoD = LCase$(VBArchivoSinPath(ProyectoD.aDepencias(j).ContainingFile))
                    FileSizeD = ProyectoD.aDepencias(j).FileSize
                    FileTimeD = ProyectoD.aDepencias(j).FILETIME
                    NombreD = ProyectoD.aDepencias(j).Name
                    GuidD = ProyectoD.aDepencias(j).GUID
                    DescripD = ProyectoD.aDepencias(j).HelpString
                    VersionD = ProyectoD.aDepencias(j).MajorVersion & "." & ProyectoD.aDepencias(j).MinorVersion
                
                    If ArchivoO = ArchivoD Then
                        'comparar hora y fecha de modificacion
                        If FileSizeO <> FileSizeD Then fSize = True
                        If FileTimeO <> FileTimeD Then fTime = True
                        If NombreO <> NombreD Then fNombre = True
                        If GuidO <> GuidD Then fGuid = True
                        If DescripO <> DescripD Then fDescrip = True
                        If VersionO <> VersionD Then fVersion = True
                        
                        Found = True
                        Exit For
                    End If
                End If
            Next j
            
            'archivo no existe en proyecto destino ?
            If Not Found Then
                Call AgregaListaDeDiferencias(Icono, ProyectoO.Archivo, "Componentes", 0, "Componente no existe en destino", "", "<" & ArchivoO & ">")
            End If
            
            'tamaño
            If fSize Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "Tamaño en KBytes", "", "<" & FileSizeO & ">", "<" & FileSizeD & ">")
            End If
            
            'fecha
            If fTime Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "Fecha de Ultima Modificación", "", "<" & FileTimeO & ">", "<" & FileTimeD & ">")
            End If
            
            'nombre
            If fNombre Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "Nombre Componente", "", "<" & NombreO & ">", "<" & NombreD & ">")
            End If
            
            'guid
            If fGuid Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "GUID Componente", "", "<" & GuidO & ">", "<" & GuidD & ">")
            End If
            
            'descripcion
            If fDescrip Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "Descripcion Componente", "", "<" & DescripO & ">", "<" & DescripD & ">")
            End If
            
            'version
            If fVersion Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "Versión Componente", "", "<" & VersionO & ">", "<" & VersionD & ">")
            End If
        End If
    Next k
    
    'comparar desde destino->origen
    For k = 1 To UBound(ProyectoD.aDepencias)
        'seleccionar solo aquellas referencia de tipo dll
        If ProyectoD.aDepencias(k).Tipo = TIPO_OCX Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoD.aDepencias(k).ContainingFile))
            NombreO = ProyectoD.aDepencias(k).Name
            FileSizeO = ProyectoD.aDepencias(k).FileSize
            FileTimeO = ProyectoD.aDepencias(k).FILETIME
            GuidO = ProyectoD.aDepencias(k).GUID
            DescripO = ProyectoD.aDepencias(k).HelpString
            VersionO = ProyectoD.aDepencias(k).MajorVersion & "." & ProyectoD.aDepencias(k).MinorVersion
                
            Icono = C_ICONO_OCX
                
            'flags para comparar propiedades
            Found = False
            fSize = False
            fTime = False
            fVersion = False
            fDescrip = False
            fGuid = False
            fNombre = False
        
            'ciclar x el proyecto origen
            For j = 1 To UBound(ProyectoO.aDepencias)
                If ProyectoO.aDepencias(j).Tipo = TIPO_OCX Then
                    ArchivoD = LCase$(VBArchivoSinPath(ProyectoO.aDepencias(j).ContainingFile))
                    FileSizeD = ProyectoO.aDepencias(j).FileSize
                    FileTimeD = ProyectoO.aDepencias(j).FILETIME
                    NombreD = ProyectoO.aDepencias(j).Name
                    GuidD = ProyectoO.aDepencias(j).GUID
                    DescripD = ProyectoO.aDepencias(j).HelpString
                    VersionD = ProyectoO.aDepencias(j).MajorVersion & "." & ProyectoO.aDepencias(j).MinorVersion
                
                    If ArchivoO = ArchivoD Then
                        'comparar hora y fecha de modificacion
                        If FileSizeO <> FileSizeD Then fSize = True
                        If FileTimeO <> FileTimeD Then fTime = True
                        If NombreO <> NombreD Then fNombre = True
                        If GuidO <> GuidD Then fGuid = True
                        If DescripO <> DescripD Then fDescrip = True
                        If VersionO <> VersionD Then fVersion = True
                        
                        Found = True
                        Exit For
                    End If
                End If
            Next j
            
            'archivo no existe en proyecto destino ?
            If Not Found Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Componentes", 0, "Componente No Existe en destino.", "")
            End If
        End If
    Next k
    
End Sub

'compara las diferencias entre formularios
Private Sub DiferenciaInformacionArchivosProyecto(ByVal TipoArchivo As eTipoArchivo, ByVal Icono As Integer)

    Dim k As Integer
    Dim j As Integer
    Dim ArchivoO As String
    Dim ArchivoD As String
    Dim NombreO As String
    Dim NombreD As String
    
    Call HelpCarga("Archivos del proyecto ...")
    
    'ciclar x el proyecto origen
    For k = 1 To UBound(ProyectoO.aArchivos)
        If ProyectoO.aArchivos(k).Explorar Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoO.aArchivos(k).PathFisico))
            NombreO = ProyectoO.aArchivos(k).ObjectName
                    
            'verificar el tipo de archivo
            If ProyectoO.aArchivos(k).TipoDeArchivo = TipoArchivo Then
                'ciclar x el proyecto destino
                For j = 1 To UBound(ProyectoD.aArchivos)
                    If ProyectoD.aArchivos(j).Explorar Then
                        ArchivoD = LCase$(VBArchivoSinPath(ProyectoD.aArchivos(j).PathFisico))
                        NombreD = ProyectoD.aArchivos(j).ObjectName
                        
                        'comparar solo los archivos que existen en destino
                        If LCase$(ArchivoO) = LCase$(ArchivoD) Then
                            'comparar la seccion general
                            If arr_ComCodigo(1) = 1 Then
                                Call CompararSeccionGeneral(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar variables
                            If arr_ComCodigo(2) = 1 Then
                                Call CompararVariables(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar constantes
                            If arr_ComCodigo(3) = 1 Then
                                Call CompararConstantes(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar enumeraciones
                            If arr_ComCodigo(4) = 1 Then
                                Call CompararEnumeraciones(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar tipos
                            If arr_ComCodigo(5) = 1 Then
                                Call CompararTipos(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar arreglos
                            If arr_ComCodigo(6) = 1 Then
                                Call CompararArreglos(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar apis
                            If arr_ComCodigo(7) = 1 Then
                                Call CompararApis(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                            
                            'comparar propiedades
                            If arr_ComCodigo(8) = 1 Then
                                Call CompararProcedimientos(k, ArchivoO, j, ArchivoD, Icono, TIPO_PROPIEDAD)
                            End If
                            
                            'comparar funciones
                            If arr_ComCodigo(9) = 1 Then
                                Call CompararProcedimientos(k, ArchivoO, j, ArchivoD, Icono, TIPO_FUN)
                            End If
                            
                            'comparar subs
                            If arr_ComCodigo(10) = 1 Then
                                Call CompararProcedimientos(k, ArchivoO, j, ArchivoD, Icono, TIPO_SUB)
                            End If
                            
                            'comparar eventos
                            If arr_ComCodigo(11) = 1 Then
                                Call CompararEventos(k, ArchivoO, j, ArchivoD, Icono)
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next k
    
End Sub

'diferencias a nivel de informacion de proyecto
Private Sub DiferenciaInformacionProyecto()
    
    Dim total_ref_o As Integer
    Dim total_ref_d As Integer
    Dim total_com_o As Integer
    Dim total_com_d As Integer
    Dim total_frm_o As Integer
    Dim total_frm_d As Integer
    Dim total_bas_o As Integer
    Dim total_bas_d As Integer
    Dim total_cls_o As Integer
    Dim total_cls_d As Integer
    Dim total_ctl_o As Integer
    Dim total_ctl_d As Integer
    Dim total_pag_o As Integer
    Dim total_pag_d As Integer
    Dim total_rel_o As Integer
    Dim total_rel_d As Integer
    Dim total_dsr_o As Integer
    Dim total_dsr_d As Integer
        
    Call HelpCarga("Proyecto ...")
    
    'diferencias entre definicion de proyectos
    If ProyectoO.Nombre <> ProyectoD.Nombre Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Nombre Proyecto", "", "<" & ProyectoO.Nombre & ">", "Nombre Destino : <" & ProyectoO.Nombre & ">")
    End If
    
    If UBound(ProyectoO.aArchivos) <> UBound(ProyectoD.aArchivos) Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Archivos", "", "<" & UBound(ProyectoO.aArchivos) & ">", "<" & UBound(ProyectoD.aArchivos) & ">")
    End If
        
    If UBound(ProyectoO.aDepencias) <> UBound(ProyectoD.aDepencias) Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Dependencias", "", "<" & UBound(ProyectoO.aDepencias) & ">", "<" & UBound(ProyectoD.aDepencias) & ">")
    End If
    
    If ProyectoO.ExeName <> ProyectoD.ExeName Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Nombre Exe ", "", "<" & ProyectoO.ExeName & ">", "<" & ProyectoD.ExeName & ">")
    End If
    
    If ProyectoO.FileSize <> ProyectoD.FileSize Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Tamaño KB", "", "<" & ProyectoO.FileSize & ">", "<" & ProyectoD.FileSize & ">")
    End If
    
    If ProyectoO.FILETIME <> ProyectoD.FILETIME Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Fecha Ultima Modificación", "", "<" & ProyectoO.FILETIME & ">", "<" & ProyectoD.FILETIME & ">")
    End If
    
    If ProyectoO.Version <> ProyectoD.Version Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Versión Proyecto", "", "<" & ProyectoO.Version & ">", "<" & ProyectoD.Version & ">")
    End If
    
    'diferencias generales
    If TotalesProyectoO.TotalLineasDeCodigo <> TotalesProyectoD.TotalLineasDeCodigo Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Lineas de Código", "", "<" & TotalesProyectoO.TotalLineasDeCodigo & ">", "<" & TotalesProyectoD.TotalLineasDeCodigo & ">")
    End If
        
    If TotalesProyectoO.TotalLineasDeComentarios <> TotalesProyectoD.TotalLineasDeComentarios Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Lineas Comentarios", "", "<" & TotalesProyectoO.TotalLineasDeComentarios & ">", "<" & TotalesProyectoD.TotalLineasDeComentarios & ">")
    End If
    
    If TotalesProyectoO.TotalLineasEnBlancos <> TotalesProyectoD.TotalLineasEnBlancos Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Lineas Blanco", "", "<" & TotalesProyectoO.TotalLineasEnBlancos & ">", "<" & TotalesProyectoD.TotalLineasEnBlancos & ">")
    End If
    
    'total referencias
    total_ref_o = ContarTipoDependencias(TIPO_DLL, ProyectoO)
    total_ref_d = ContarTipoDependencias(TIPO_DLL, ProyectoD)
    If total_ref_o <> total_ref_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Referencias", "", "<" & total_ref_o & ">", "<" & total_ref_d & ">")
    End If
        
    'total componentes
    total_com_o = ContarTipoDependencias(TIPO_OCX, ProyectoO)
    total_com_d = ContarTipoDependencias(TIPO_OCX, ProyectoD)
    
    If total_com_o <> total_com_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Componentes", "", "<" & total_com_o & ">", "<" & total_com_d & ">")
    End If
    
    'total formularios
    total_frm_o = ContarTiposDeArchivos(TIPO_ARCHIVO_FRM, ProyectoO)
    total_frm_d = ContarTiposDeArchivos(TIPO_ARCHIVO_FRM, ProyectoD)
    
    If total_frm_o <> total_frm_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Formularios", "", "<" & total_frm_o & ">", "<" & total_frm_d & ">")
    End If
    
    'total modulos .bas
    total_bas_o = ContarTiposDeArchivos(TIPO_ARCHIVO_BAS, ProyectoO)
    total_bas_d = ContarTiposDeArchivos(TIPO_ARCHIVO_BAS, ProyectoD)
    
    If total_bas_o <> total_bas_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Módulos .Bas", "", "<" & total_bas_o & ">", "<" & total_bas_d & ">")
    End If
    
    'total modulos .cls
    total_cls_o = ContarTiposDeArchivos(TIPO_ARCHIVO_CLS, ProyectoO)
    total_cls_d = ContarTiposDeArchivos(TIPO_ARCHIVO_CLS, ProyectoD)
    
    If total_cls_o <> total_cls_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Módulos .Cls", "", "<" & total_cls_o & ">", "<" & total_cls_d & ">")
    End If
    
    total_ctl_o = ContarTiposDeArchivos(TIPO_ARCHIVO_OCX, ProyectoO)
    total_ctl_d = ContarTiposDeArchivos(TIPO_ARCHIVO_OCX, ProyectoD)
    
    If total_ctl_o <> total_ctl_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Controles de Usuario", "", "<" & total_ctl_o & ">", "<" & total_ctl_d & ">")
    End If
    
    total_pag_o = ContarTiposDeArchivos(TIPO_ARCHIVO_PAG, ProyectoO)
    total_pag_d = ContarTiposDeArchivos(TIPO_ARCHIVO_PAG, ProyectoD)
    
    If total_ctl_o <> total_ctl_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Páginas de Propiedades", "", "<" & total_pag_o & ">", "<" & total_pag_d & ">")
    End If
    
    total_rel_o = ContarTiposDeArchivos(TIPO_ARCHIVO_REL, ProyectoO)
    total_rel_d = ContarTiposDeArchivos(TIPO_ARCHIVO_REL, ProyectoD)
    
    If total_rel_o <> total_rel_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Archivos Relacionados", "", "<" & total_rel_o & ">", "<" & total_rel_d & ">")
    End If
    
    total_dsr_o = ContarTiposDeArchivos(TIPO_ARCHIVO_DSR, ProyectoO)
    total_dsr_d = ContarTiposDeArchivos(TIPO_ARCHIVO_DSR, ProyectoD)
    
    If total_dsr_o <> total_dsr_d Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Archivos Diseñadores", "", "<" & total_dsr_o & ">", "<" & total_dsr_d & ">")
    End If
        
    If TotalesProyectoO.TotalSubs <> TotalesProyectoD.TotalSubs Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Subs", "", "<" & TotalesProyectoO.TotalSubs & ">", "<" & TotalesProyectoD.TotalSubs & ">")
    End If

    If TotalesProyectoO.TotalFunciones <> TotalesProyectoD.TotalFunciones Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Funciones", "", "<" & TotalesProyectoO.TotalFunciones & ">", "<" & TotalesProyectoD.TotalFunciones & ">")
    End If
    
    If TotalesProyectoO.TotalPropertyLets <> TotalesProyectoD.TotalPropertyLets Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Property Let", "", "<" & TotalesProyectoO.TotalPropertyLets & ">", "<" & TotalesProyectoD.TotalPropertyLets & ">")
    End If
    
    If TotalesProyectoO.TotalPropertySets <> TotalesProyectoD.TotalPropertySets Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Property Set", "", "<" & TotalesProyectoO.TotalPropertySets & ">", "<" & TotalesProyectoD.TotalPropertySets & ">")
    End If
    
    If TotalesProyectoO.TotalPropertyGets <> TotalesProyectoD.TotalPropertyGets Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Property Get", "", "<" & TotalesProyectoO.TotalPropertyGets & ">", "<" & TotalesProyectoD.TotalPropertyGets & ">")
    End If
    
    If TotalesProyectoO.TotalVariables <> TotalesProyectoD.TotalVariables Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Variables", "", "<" & TotalesProyectoO.TotalVariables & ">", "<" & TotalesProyectoD.TotalVariables & ">")
    End If
    
    If TotalesProyectoO.TotalConstantes <> TotalesProyectoD.TotalConstantes Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Constantes", "", "<" & TotalesProyectoO.TotalConstantes & ">", "<" & TotalesProyectoD.TotalConstantes & ">")
    End If
    
    If TotalesProyectoO.TotalTipos <> TotalesProyectoD.TotalTipos Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Tipos", "", "<" & TotalesProyectoO.TotalTipos & ">", "<" & TotalesProyectoD.TotalTipos & ">")
    End If
    
    If TotalesProyectoO.TotalEnumeraciones <> TotalesProyectoD.TotalEnumeraciones Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Enumeraciones", "", "<" & TotalesProyectoO.TotalEnumeraciones & ">", "<" & TotalesProyectoD.TotalEnumeraciones & ">")
    End If
    
    If TotalesProyectoO.TotalApi <> TotalesProyectoD.TotalApi Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Apis", "", "<" & TotalesProyectoO.TotalApi & ">", "<" & TotalesProyectoD.TotalApi & ">")
    End If
    
    If TotalesProyectoO.TotalArray <> TotalesProyectoD.TotalArray Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Arrays", "", "<" & TotalesProyectoO.TotalArray & ">", "<" & TotalesProyectoD.TotalArray & ">")
    End If
    
    If TotalesProyectoO.TotalControles <> TotalesProyectoD.TotalControles Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Controles", "", "<" & TotalesProyectoO.TotalControles & ">", "<" & TotalesProyectoD.TotalControles & ">")
    End If
    
    If TotalesProyectoO.TotalEventos <> TotalesProyectoD.TotalEventos Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Eventos", "", "<" & TotalesProyectoO.TotalEventos & ">", "<" & TotalesProyectoD.TotalEventos & ">")
    End If
    
    If TotalesProyectoO.TotalMiembrosPublicos <> TotalesProyectoD.TotalMiembrosPublicos Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Miembros Públicos", "", "<" & TotalesProyectoO.TotalMiembrosPublicos & ">", "<" & TotalesProyectoD.TotalMiembrosPublicos & ">")
    End If
    
    If TotalesProyectoO.TotalMiembrosPrivados <> TotalesProyectoD.TotalMiembrosPrivados Then
        Call AgregaListaDeDiferencias(C_ICONO_PROYECTO, ProyectoO.Archivo, "Proyecto", 0, "Total Miembros Privados", "", "<" & TotalesProyectoO.TotalMiembrosPrivados & ">", "<" & TotalesProyectoD.TotalMiembrosPrivados & ">")
    End If
    
End Sub
'diferencias a nivel de referencias
Private Sub DiferenciaInformacionReferencias()

    Dim k As Integer
    Dim j As Integer
    Dim c As Integer
    Dim Icono As Integer
    
    Dim Found As Boolean
    Dim fSize As Boolean
    Dim fTime As Boolean
    Dim fVersion As Boolean
    Dim fDescrip As Boolean
    Dim fGuid As Boolean
    Dim fNombre As Boolean
                
    Dim ArchivoO As String
    Dim ArchivoD As String
    Dim FileSizeO As Double
    Dim FileSizeD As Double
    Dim FileTimeO As String
    Dim FileTimeD As String
    Dim VersionO As String
    Dim VersionD As String
    Dim DescripO As String
    Dim DescripD As String
    Dim GuidO As String
    Dim GuidD As String
    Dim NombreO As String
    Dim NombreD As String
        
    'ciclar x los archivos del proyecto origen
    'comprobar si existen
    
    Call HelpCarga("Referencias ...")
    
    c = 1
    For k = 1 To UBound(ProyectoO.aDepencias)
        'seleccionar solo aquellas referencia de tipo dll
        If ProyectoO.aDepencias(k).Tipo = TIPO_DLL Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoO.aDepencias(k).ContainingFile))
            NombreO = ProyectoO.aDepencias(k).Name
            FileSizeO = ProyectoO.aDepencias(k).FileSize
            FileTimeO = ProyectoO.aDepencias(k).FILETIME
            GuidO = ProyectoO.aDepencias(k).GUID
            DescripO = ProyectoO.aDepencias(k).HelpString
            VersionO = ProyectoO.aDepencias(k).MajorVersion & "." & ProyectoO.aDepencias(k).MinorVersion
                
            Icono = C_ICONO_DLL
                
            'flags para comparar propiedades
            Found = False
            fSize = False
            fTime = False
            fVersion = False
            fDescrip = False
            fGuid = False
            fNombre = False
        
            'ciclar x el proyecto destino
            For j = 1 To UBound(ProyectoD.aDepencias)
                If ProyectoD.aDepencias(j).Tipo = TIPO_DLL Then
                    ArchivoD = LCase$(VBArchivoSinPath(ProyectoD.aDepencias(j).ContainingFile))
                    FileSizeD = ProyectoD.aDepencias(j).FileSize
                    FileTimeD = ProyectoD.aDepencias(j).FILETIME
                    NombreD = ProyectoD.aDepencias(k).Name
                    GuidD = ProyectoD.aDepencias(k).GUID
                    DescripD = ProyectoD.aDepencias(k).HelpString
                    VersionD = ProyectoD.aDepencias(k).MajorVersion & "." & ProyectoD.aDepencias(k).MinorVersion
            
                    If ArchivoO = ArchivoD Then
                        'comparar hora y fecha de modificacion
                        If FileSizeO <> FileSizeD Then fSize = True
                        If FileTimeO <> FileTimeD Then fTime = True
                        If NombreO <> NombreD Then fNombre = True
                        If GuidO <> GuidD Then fGuid = True
                        If DescripO <> DescripD Then fDescrip = True
                        If VersionO <> VersionD Then fVersion = True
                        
                        Found = True
                        Exit For
                    End If
                End If
            Next j
            
            'archivo no existe en proyecto destino ?
            If Not Found Then
                Call AgregaListaDeDiferencias(Icono, ProyectoO.Archivo, "Referencias", 0, "Referencia no existe", "", "", "<" & ArchivoO & ">")
            End If
            
            'tamaño
            If fSize Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "Tamaño KBytes", "", "<" & FileSizeO & ">", "<" & FileSizeD & ">")
            End If
            
            'fecha
            If fTime Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "Fecha Ultima Modificación", "", "<" & FileTimeO & ">", "<" & FileTimeD & ">")
            End If
            
            'nombre
            If fNombre Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "Nombre Referencia", "", "<" & NombreO & ">", "<" & NombreD & ">")
            End If
            
            'guid
            If fGuid Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "GUID Referencia", "", "<" & GuidO & ">", "<" & GuidD & ">")
            End If
            
            'descripcion
            If fDescrip Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "Descripcion Referencia", "", "<" & DescripO & ">", "<" & DescripD & ">")
            End If
            
            'version
            If fVersion Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "Versión Referencia", "", "<" & VersionO & ">", "<" & VersionD & ">")
            End If
        End If
    Next k
    
    'comparar desde destino->origen
    c = 1
    For k = 1 To UBound(ProyectoD.aDepencias)
        'seleccionar solo aquellas referencia de tipo dll
        If ProyectoD.aDepencias(k).Tipo = TIPO_DLL Then
            ArchivoO = LCase$(VBArchivoSinPath(ProyectoD.aDepencias(k).ContainingFile))
            NombreO = ProyectoD.aDepencias(k).Name
            FileSizeO = ProyectoD.aDepencias(k).FileSize
            FileTimeO = ProyectoD.aDepencias(k).FILETIME
            GuidO = ProyectoD.aDepencias(k).GUID
            DescripO = ProyectoD.aDepencias(k).HelpString
            VersionO = ProyectoD.aDepencias(k).MajorVersion & "." & ProyectoD.aDepencias(k).MinorVersion
                
            Icono = C_ICONO_DLL
                
            'flags para comparar propiedades
            Found = False
            fSize = False
            fTime = False
            fVersion = False
            fDescrip = False
            fGuid = False
            fNombre = False
        
            'ciclar x el proyecto destino
            For j = 1 To UBound(ProyectoO.aDepencias)
                If ProyectoD.aDepencias(j).Tipo = TIPO_DLL Then
                    ArchivoD = LCase$(VBArchivoSinPath(ProyectoO.aDepencias(j).ContainingFile))
                    FileSizeD = ProyectoO.aDepencias(j).FileSize
                    FileTimeD = ProyectoO.aDepencias(j).FILETIME
                    NombreD = ProyectoO.aDepencias(j).Name
                    GuidD = ProyectoO.aDepencias(j).GUID
                    DescripD = ProyectoO.aDepencias(j).HelpString
                    VersionD = ProyectoO.aDepencias(j).MajorVersion & "." & ProyectoO.aDepencias(j).MinorVersion
            
                    If ArchivoO = ArchivoD Then
                        'comparar hora y fecha de modificacion
                        If FileSizeO <> FileSizeD Then fSize = True
                        If FileTimeO <> FileTimeD Then fTime = True
                        If NombreO <> NombreD Then fNombre = True
                        If GuidO <> GuidD Then fGuid = True
                        If DescripO <> DescripD Then fDescrip = True
                        If VersionO <> VersionD Then fVersion = True
                        
                        Found = True
                        Exit For
                    End If
                End If
            Next j
            
            'archivo no existe en proyecto destino ?
            If Not Found Then
                Call AgregaListaDeDiferencias(Icono, ArchivoO, "Referencias", 0, "Referencia No Existe en destino", "", "<" & ArchivoO & ">", "")
            End If
        End If
    Next k
    
End Sub


Public Sub FiltraComparaciones(ByVal Filtro As String)

    Dim k As Integer
    Dim Glosa As String
    Dim c As Integer
    Dim sKey As String
    
    Call Hourglass(frmMain.hWnd, True)
    
    frmMain.lvwProblemas.ListItems.Clear
    
    c = 1
    
    Call HelpCarga("Cargando " & Filtro & " ...")
    'ciclar x las diferencias
    For k = 1 To UBound(arr_diferencias)
        If arr_diferencias(k).Ubicacion = Filtro Then
            c = frmMain.lvwProblemas.ListItems.Count + 1
            sKey = "k" & c
            
            frmMain.lvwProblemas.ListItems.Add , sKey, CStr(c), arr_diferencias(k).Icono, arr_diferencias(k).Icono
            frmMain.lvwProblemas.ListItems(sKey).SubItems(1) = arr_diferencias(k).Archivo
            frmMain.lvwProblemas.ListItems(sKey).SubItems(2) = arr_diferencias(k).Ubicacion
            frmMain.lvwProblemas.ListItems(sKey).SubItems(3) = arr_diferencias(k).Linea
            frmMain.lvwProblemas.ListItems(sKey).SubItems(4) = arr_diferencias(k).DifOrigen
            frmMain.lvwProblemas.ListItems(sKey).SubItems(5) = arr_diferencias(k).DifDestino
            frmMain.lvwProblemas.ListItems(sKey).SubItems(6) = arr_diferencias(k).DecOrigen
            frmMain.lvwProblemas.ListItems(sKey).SubItems(7) = arr_diferencias(k).DecDestino
        End If
    Next k
    
    Call HelpCarga("Listo")
    Call Hourglass(frmMain.hWnd, False)
    
End Sub



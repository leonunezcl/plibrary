Attribute VB_Name = "mEstProyecto"
Option Explicit

Enum eTipoRutinas
    TIPO_SUB = 1
    TIPO_FUN = 2
    TIPO_PROPIEDAD = 6
End Enum

Enum eTipoArchivo
    TIPO_ARCHIVO_FRM = 1
    TIPO_ARCHIVO_BAS = 2
    TIPO_ARCHIVO_CLS = 3
    TIPO_ARCHIVO_OCX = 4
    TIPO_ARCHIVO_PAG = 5
    TIPO_ARCHIVO_REL = 6
    TIPO_ARCHIVO_DSR = 7
End Enum

Enum eEstado
    NOCHEQUEADO = 0
    LIVE = 1
    DEAD = 2
End Enum

Enum eTipoPropiedad
    TIPO_GET = 1
    TIPO_LET = 2
    TIPO_SET = 3
End Enum

Type eDatosVariables
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    Operador As String
    KeyNode As String
    Estado As eEstado
    Tipo As String
    TipoVb As String
    Predefinido As Boolean 'tipo variant x defecto ?
    UsaDim As Boolean 'para versiones > 4 debiera usar private
    UsaGlobal As Boolean 'para versiones > 4 debiersa usar public
    UsaPrivate As Boolean 'para constantes con const x =
    BasicOldStyle As Boolean 'definida al viejo estilo basic $,%,&
    Linea As Integer        'linea de la rutina
End Type

Type eDatosControl
    Nombre As String        'nombre del control
    Clase As String         'clase del control
    Eventos As String       'eventos programados
    Numero As Integer       'cantidad de controles
    Descripcion As String   'descripcion
End Type

Type eTipoDeVariable
    TipoDefinido As String
    Cantidad As Integer
End Type

Type eDatosParametros
    PorValor As Boolean
    Nombre As String
    Glosa As String
    TipoParametro As String
    Estado As eEstado
    BasicStyle As Boolean
End Type

Type eInfoAnalisis
    Icono As Integer
    Problema As String
    Linea As Integer
End Type

Type eCodigo
    Codigo As String
    Linea As Integer
End Type

Type eRutinas
    Nombre As String
    NombreRutina As String
    Aparams() As eDatosParametros       'informacion de los parametros
    nVariables As Long
    aVariables() As eDatosVariables     'variables de las rutinas
    aRVariables() As eTipoDeVariable    'resumen de las variables
    aAnalisis() As eInfoAnalisis
    nAnalisis As Integer
    Tipo As eTipoRutinas                'funcion/sub/propiedad
    
    TipoProp As eTipoPropiedad          'get/let/set
    Publica As Boolean
    KeyNode As String
    TempFileName As String
    TempCodigoRutina As String
    aCodigoRutina() As eCodigo          'guardar el codigo de la rutina
    NumeroDeLineas As Integer
    NumeroDeComentarios As Integer
    NumeroDeBlancos As Integer
    TotalLineas As Integer
    Estado As eEstado                   'usada/no usada
    RegresaValor As Boolean             'usado para las funciones
    Mensaje As String
    IsObjectSub As Boolean              'es sub de control ?
    IsMenu As Boolean
    IsSeparador As Boolean
    Linea As Integer                    'linea del archivo
    TipoRetorno As String
    BasicStyle As Boolean
End Type

Type eElementosTipos
    Nombre As String
    Tipo As String
    Estado As eEstado
    KeyNode As String
    Linea As Integer
End Type

Type eTipos
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    KeyNode As String
    Estado As eEstado
    Linea As Integer
    aElementos() As eElementosTipos
End Type

Type eElementosEnum
    Nombre As String
    Valor As String
    Estado As eEstado
    KeyNode As String
    Linea As Integer
End Type

Type eEnum
    Nombre As String
    NombreVariable As String
    Publica As Boolean
    KeyNode As String
    Estado As eEstado
    Linea As Integer
    aElementos() As eElementosEnum
End Type

Type eDatos
    OptionExplicit As Boolean       'usa option explicit
    Explorar As Boolean             'analizar archivo
    Nombre As String                'nombre
    PathFisico As String            'path fisico
    FileSize As Long                'tamaño
    FILETIME As String              'fecha/hora
    ObjectName As String            'nombre logico
    Descripcion As String
    Usado As Boolean                'se hace referencia a alguna variable/sub/propiedad
    TipoDeArchivo As eTipoArchivo   'frm,bas,cls,pag,ocx
    
    aGeneral() As eCodigo           'guardar codigo de general
    aAnalisis() As eInfoAnalisis    'arreglo donde se guarda los problemas de analisis
    nAnalisis As Integer            'contador del arreglo de analisis
    Linea As Integer                'linea de codigo de la seccion general
    KeyNodeFrm As String            'llaves de los nodos
    KeyNodeBas As String            'llaves de los nodos
    KeyNodeCls As String            'llaves de los nodos
    KeyNodeKtl As String            'llaves de los nodos
    KeyNodePag As String            'llaves de los nodos
    KeyNodeRel As String            'llaves de los nodos
    KeyNodeDsr As String            'llaves de los nodos
    
    nControles As Integer           'total de controles de archivo
    aControles() As eDatosControl   'guardar controles
        
    nVariables As Integer
    nVariablesPrivadas As Integer
    nVariablesPublicas As Integer
    aVariables() As eDatosVariables     'guardar variables
    aTipoVariable() As eTipoDeVariable  'acumulador de tipos de variables
    KeyNodeVar As String
    
    nConstantes As Integer
    nConstantesPrivadas As Integer
    nConstantesPublicas As Integer
    aConstantes() As eDatosVariables    'guardar constantes
    KeyNodeCte As String
    
    nEnumeraciones As Integer
    nEnumeracionesPrivadas As Integer
    nEnumeracionesPublicas As Integer
    aEnumeraciones() As eEnum 'guardar enumeraciones
    KeyNodeEnum As String
    
    nArray As Integer
    nArrayPrivadas As Integer
    nArrayPublicas As Integer
    aArray() As eDatosVariables         'guardar arrays
    KeyNodeArr As String
    
    nRutinas As Integer
    nTipoSub As Integer
    nTipoSubPublicas As Integer
    nTipoSubPrivadas As Integer
    KeyNodeSub As String
    
    NumeroDeLineas As Integer
    NumeroDeLineasEnBlanco As Integer
    NumeroDeLineasComentario As Integer
    TotalLineas As Integer
    
    aRutinas() As eRutinas              'guardar rutinas
    
    nTipoFun As Integer
    nTipoFunPublica As Integer
    nTipoFunPrivada As Integer
    KeyNodeFun As String
    
    nTipoApi As Integer
    KeyNodeApi As String
    aApis() As eDatosVariables          'guardar apis
    
    nTipos As Integer
    nTiposPrivadas As Integer
    nTiposPublicas As Integer
    aTipos() As eTipos         'guardar tipos
    KeyNodeTipo As String
    
    nPropiedades As Integer
    nPropertyLet As Integer
    nPropertySet As Integer
    nPropertyGet As Integer
        
    KeyNodeProp As String
    
    nEventos As Integer
    nEventosPrivadas As Integer
    nEventosPublicas As Integer
    aEventos() As eDatosVariables       'guardar eventos
    KeyNodeEvento As String
    
    MiembrosPrivados As Integer
    MiembrosPublicos As Integer
End Type

Public Enum eTipoDepencia
    TIPO_DLL = 1
    TIPO_OCX = 2
    TIPO_RES = 3
    TIPO_PAGE = 4
End Enum

Public Type eDependencias
    Tipo As eTipoDepencia
    Archivo As String
    GUID As String
    KeyNode As String
    Name As String
    ContainingFile As String
    HelpString As String
    HelpFile As String
    MajorVersion As Long
    MinorVersion As Long
    FileSize As Long
    FILETIME As String
End Type

Public Enum eTipoProyecto
    PRO_TIPO_NONE = 0
    PRO_TIPO_EXE = 1
    PRO_TIPO_DLL = 2
    PRO_TIPO_OCX = 3
    PRO_TIPO_EXE_X = 4
End Enum

Public Type eProyecto
    Nombre As String
    Archivo As String
    Icono As Integer
    Version As Integer
    PathFisico As String
    ExeName As String
    TipoProyecto As eTipoProyecto
    FileSize As Long
    FILETIME As String
    Startup As String
    Analizado As Boolean
    aArchivos() As eDatos
    aDepencias() As eDependencias
End Type
Public ProyectoO As eProyecto
Public ProyectoD As eProyecto

Public Type eTotalesProyecto
    TotalVariables As Long
    TotalVariablesPrivadas As Long
    TotalVariablesPublicas As Long
    
    TotalConstantes As Long
    TotalConstantesPrivadas As Long
    TotalConstantesPublicas As Long
    
    TotalEnumeraciones As Long
    TotalEnumeracionesPrivadas As Long
    TotalEnumeracionesPublicas As Long
    
    TotalApi As Long
    
    TotalArray As Long
    TotalArrayPrivadas As Long
    TotalArrayPublicas As Long
    
    TotalTipos As Long
    TotalTiposPrivadas As Long
    TotalTiposPublicas As Long
    
    TotalSubs As Long
    TotalSubsPrivadas As Long
    TotalSubsPublicas As Long
    
    TotalFunciones As Long
    TotalFuncionesPrivadas As Long
    TotalFuncionesPublicas As Long
        
    TotalLineasDeCodigo As Long
    TotalLineasEnBlancos As Long
    TotalLineasDeComentarios As Long
    
    TotalPropiedades As Long
    TotalPropertyLets As Integer
    TotalPropertySets As Integer
    TotalPropertyGets As Integer
    
    TotalControles As Long
    TotalEventos As Long
    
    TotalMiembrosPrivados As Long
    TotalMiembrosPublicos As Long
End Type
Public TotalesProyectoO As eTotalesProyecto
Public TotalesProyectoD As eTotalesProyecto

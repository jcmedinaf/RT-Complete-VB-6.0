Attribute VB_Name = "Gerenal"

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal a As String, ByVal b As Long, ByVal C As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, ByVal g As Long, ByVal h As Integer) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Pn
Public Declare Sub PortOut Lib "io.dll" (ByVal Port As Integer, ByVal Data As Byte)
Public Query As String
Public NombreEquipo As String
Public RutaInformes, RutaFotos, Foto, FotoEmp, FotoSimul, FotoSimul2, Dir_ZETA As String
Public IpRemota, PortRemoto As String
Public User_Priv As String
Public Const WM_CAP_START = &H400
Public Const WM_cap_driver_connect = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11

Public Const WM_CAP_EDIT_COPY = WM_CAP_START + 30
Public Const WM_cap_set_preview = WM_CAP_START + 50
Public Const wm_cap_set_overlay = WM_CAP_START + 51
Public Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
Public Const WM_CAP_SEQUENCE = WM_CAP_START + 62
Public Const WM_CAP_SINGLE_FRAME_OPEN = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME = WM_CAP_START + 72
Public IdReg
Public Especia As String
Public Editar
Enum EACCION
     AGREGAR_REGISTRO = 0
     EDITAR_REGISTRO = 1
End Enum
Public Gram As TextoPos
Public Type TextoPos
    Texto As String
    Poscur As String
End Type
Public op
Public Consulta
Public Consultaa
Public Reg_Actual(0 To 70) As String 'Alamacena datos para comprar en la bitacora
Public PyFs(0 To 23, 0 To 2) As String
Public ListFunc(0 To 23, 2) As String
Public ListCond(0 To 23, 2) As String
Public Bi
Public OpcionReporte
Public Tipo
Public Tipo1
Public opcion
Public Ayuda
Public T_U
Public Usuario
Public Direc
Public StrText
Public IDCLI, IdPaci, NoFact As String
Public IdEmpl ' id que identifica al empleado seleccionado
Public IdPac1 ' el id paciente
Public IdUser ' el id usuario
Public IdMedT ' el id de medicos tratantes
Public IdCliente As Integer
Public N_fac
Public Detalle
Public Observa
Public DNI
Public Recom
Public rs As New ADODB.Recordset
Public RsReporte As New ADODB.Recordset
Public RsWeb As New ADODB.Recordset
Public BD57 As New ADODB.Recordset
Public BD75 As New ADODB.Recordset
Public Cnn As New ADODB.Connection
Public WebCnn As New ADODB.Connection
Public dRegistro As Long
Public ACCION As EACCION
Public CSql As String
Public CSql1 As String
Public Fila
Public CodProd
Public DescPro
Public IvaProd
Public PreProd
Public IVA As Integer
Public IO As Integer
Public N_Factur
Public Carac
Public Ban
Public ModulO 'variable para determinar el modulo desde el que se llama a la lista de pacientes en espera
Public Cedul  'Variable con la cedula del paciente de la lista de espera
'<<<<<<Declaracion de variables para los conceptos de nomina>>>>>>
Public FormulA As String
Public IdsConcepto(1 To 50) As Integer
Public IdsCampo(1 To 50) As Integer
Public IdsConstante(1 To 50) As Integer
Public IDcon
Public FotoP As String
Public BdConstante As New ADODB.Recordset
Dim Cnt As Boolean
Public Const MAX_COMPUTERNAME_LENGTH = 255
Public BeamDescripcion
Public IdEmprs As Integer
Public Intentos As Integer

Public IdL As String
Public IdLIdPac As String
Public IdLDefault As String
Public NuevoIdL As String
Public IdLIdInf As String

Public Function ComputerName() As String
'Devuelve el nombre del equipo actual
Dim sComputerName As String
Dim ComputerNameLength As Long

sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
ComputerNameLength = MAX_COMPUTERNAME_LENGTH
Call GetComputerName(sComputerName, ComputerNameLength)
ComputerName = Mid(sComputerName, 1, ComputerNameLength)
NombreEquipo = ComputerName
End Function

Sub Main()
On Error GoTo Resolv
If App.PrevInstance = False Then
IdL = "I"
IdLDefault = "I"
ComputerName
    'Conexion Local con Access
    'StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB OncoAmerica\OncoAmerica.mdb" & ";Jet OLEDB:System Database=" & Direc & "\OA.mdw;"
    'StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB OncoAmerica\OncoAmerica.mdb;"
    
    'Conexion Remota con Access
    'StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Direc & "\OA.mdb & ";Jet OLEDB:System Database=" & Direc & "\OA.mdw;"
    'StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Direc & "\OA.mdb;"
    
    'Conexion Local con Sql Server
    
    'Ing03
    StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=Ing03"
    ' StrCn = "Driver={SQL Server}; Server=ING03; Database=OACLINICA; UID=sa; PWD=458921957JAr;"
    
    'Ing04
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=Ing04"
  
    'Ing04Lapto
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=Ing04Lapto"
    
    'Server Indio Mara
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=Server"
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=192.168.1.190"
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OATest; Data Source=192.168.1.253"
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OATest; Data Source=oaindiomara.no-ip.org"
    'StrCn = "Driver={SQL Server}; Server=oaindiomara.no-ip.org; Database=OACLINICA; UID=sa; PWD=458921957JAr;"
    
    'Server OA
    ' StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=server"
    ' StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=192.168.1.104"
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OATest; Data Source=192.168.1.253"
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OATest; Data Source=oaindiomara.no-ip.org"
    'StrCn = "Driver={SQL Server}; Server=oa.no-ip.org; Database=OACLINICA; UID=sa; PWD=458921957JAr;"
    
    'HP
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=HP"
    
    'ComputerName
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source='" & NombreEquipo & "'"
    
    'VAIO
    'StrCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False; Initial Catalog=OACLINICA; Data Source=VAIOC190G"
    
    'Sql Server 2008
    'StrCn = "Provider=SQLNCLI10;Server=Ing04\SqlExpress;Database=OACLINICA;Uid=sa; Pwd=458921957JAr;"
    'StrCn = "Provider=SQLNCLI10;Server=Server\SqlExpress;Database=OACLINICA; Trusted_Connection=yes;"
   
   FrmPrincipal.Tag = "0"
    If Cnn.State = 0 Then
        Cnn.ConnectionString = StrCn
        Cnn.Open
    End If
Volrr:
    If Not Cnn.State Then
        Load FrmSplash
        FrmSplash.Show
        FrmPrincipal.Tag = "1"
    End If

Else
    MsgBox "El programa ya se encuentra en ejecución !!!", vbInformation, "Mensaje de la Aplicación"
End If
Exit Sub

Resolv:
    MsgBox "Oppss! hubo un problema en la conexión al servidor, Contacte al administrador!", vbCritical + vbOKOnly, "Error de Conexión!"
    'MsgBox Cnn.Errors
    FrmReconexion.Show
'    If Cnn.State = adStateOpen Then
'        Load FrmSplash
'        FrmSplash.Show
'    End If

End Sub
Public Function CrearRS(strTabla As String) As ADODB.Recordset
On Error GoTo Resolv:
Dim RsTemp As New ADODB.Recordset

Volverr:

RsTemp.ActiveConnection = Cnn
RsTemp.CursorLocation = adUseClient
RsTemp.CursorType = adOpenDynamic
RsTemp.LockType = adLockOptimistic
RsTemp.Source = strTabla
RsTemp.Open
Set CrearRS = RsTemp
Set RsTemp = Nothing

Exit Function
Resolv:

    If Err.Number = -2147467259 Then
    
        Cnn.Close
    
        FrmReconexion.Show
        Verificar_CNN
        GoTo Volverr
    Else
        'MsgBox Err.Source & "MMMMMM" & Err.Description
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    
End Function

Sub Verificar_CNN()
While Cnn.State = 0
    DoEvents
Wend
End Sub


Sub ConectarHosting()

StrCn = "Provider=SQLOLEDB.1;Data Source=64.78.59.214;Initial Catalog=oasql;User ID=wendy;Password=wendy123456"
WebCnn.ConnectionString = StrCn
Intentos = Intentos + 1
WebCnn.Open
        
End Sub

Sub ConectarIVSSHosting()

'IVSS
StrCn2 = "Provider=sqloledb; Data Source=ivss.db.6280798.hostedresource.com; Initial Catalog=ivss; User ID=ivss; Password=Seguro120389"


WebCnn.ConnectionString = StrCn2
Intentos = Intentos + 1
WebCnn.Open
        
End Sub

Public Function CrearIVSSRsWeb(strTabla As String) As ADODB.Recordset
On Error GoTo Resolv:
Dim RsTemp As New ADODB.Recordset

Volverr:

RsTemp.ActiveConnection = WebCnn
RsTemp.CursorLocation = adUseClient
RsTemp.CursorType = adOpenDynamic
RsTemp.LockType = adLockOptimistic
RsTemp.Source = strTabla
RsTemp.Open
Set CrearIVSSRsWeb = RsTemp
Set RsTemp = Nothing

Exit Function
Resolv:

   If Err.Number = -2147467259 Then
    
        Cnn.Close
'        Cnn.Open
        
        FrmReconexion.Show
        Verificar_CNN
        GoTo Volverr
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function


Public Function CrearRsWeb(strTabla As String) As ADODB.Recordset
On Error GoTo Resolv:
Dim RsTemp As New ADODB.Recordset

Volverr:

RsTemp.ActiveConnection = WebCnn
RsTemp.CursorLocation = adUseClient
RsTemp.CursorType = adOpenDynamic
RsTemp.LockType = adLockOptimistic
RsTemp.Source = strTabla
RsTemp.Open
Set CrearRsWeb = RsTemp
Set RsTemp = Nothing

Exit Function
Resolv:

   If Err.Number = -2147467259 Then
    
        Cnn.Close
'        Cnn.Open
        
        FrmReconexion.Show
        Verificar_CNN
        GoTo Volverr
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

Public Function VerificarNulo(StrText)
If Not IsNull(StrText) Then
    VerificarNulo = StrText
Else
    VerificarNulo = Empty
End If
End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<fin>>>>>>>>>>>>>>>>>>>>>>>>
'Public CNn As New ADODB.Recordset
Sub Verifica_campo()
sCaracter = "Campo("
Z = 1
dcon = ""
For h = 1 To Len(FormulA)
    cars = Mid$(FormulA, h, 6)
    T = 0
    If UCase(sCaracter) = UCase(cars) Then
    
        Formula1 = Mid(FormulA, h + 6, Len(FormulA))
        For w = 1 To Len(Formula1)
            d = Mid(Formula1, w, 1)
            If d = ";" Then T = 1
        
            If T = 1 And Not d = ";" And Not d = ")" Then dcon = dcon + d
            If d = ")" Then Exit For
        Next w
        If Trim(dcon) <> "" And Trim(dcon) <> "0" Then IdsConcepto(Z) = dcon: GoSub consulta8: Formula2 = Formula2 & Valor: Z = Z + 1: dcon = ""
        
        h = h + 6 + w
    End If
    Formula2 = Formula2 & Mid(FormulA, h, 1)
Next h
FormulA = Formula2
'For df = 1 To Z - 1
'MsgBox idsconcepto(df)
'Next df
Exit Sub
consulta8:
CSql = "Select * From CamposDelTrabajador Where IdCampoNomina = " & IdsConcepto(Z) & " And IdEmpleado = " & IdEmpl
Set BdConstante = CrearRS(CSql)
If Not BdConstante.EOF Then
With BdConstante
If Not IsNull(.Fields("valorn")) Or Not (Trim(.Fields("valorn")) = "") Then Valor = .Fields("valorn")
'If Not IsNull(.Fields("valort")) Or Not (Trim(.Fields("valorT")) = "") Then valor = .Fields("valort")
'If Not IsNull(.Fields("valorf")) Or Not (Trim(.Fields("valorf")) = "") Then valor = .Fields("valorf")
End With
End If
BdConstante.Close
Return
End Sub

Sub verifica_constante()
sCaracter = "Constante("
Z = 1
dcon = ""
For h = 1 To Len(FormulA)
    cars = Mid$(FormulA, h, 10)
    T = 0
    If UCase(sCaracter) = UCase(cars) Then
       
        Formula1 = Mid(FormulA, h + 10, Len(FormulA))
        For w = 1 To Len(Formula1)
            d = Mid(Formula1, w, 1)
            If d = ";" Then T = 1
        
            If T = 1 And Not d = ";" And Not d = ")" Then dcon = dcon + d
            If d = ")" Then Exit For
        Next w
        If Trim(dcon) <> "" And Trim(dcon) <> "0" Then IdsConstante(Z) = dcon: GoSub Consulta: Formula2 = Formula2 & Valor: Z = Z + 1: dcon = ""
        
        h = h + 10 + w
                
    End If
    Formula2 = Formula2 & Mid(FormulA, h, 1)
Next h
'For df = 1 To Z - 1
FormulA = Formula2
'MsgBox idsconstante(df)
'Next df
Exit Sub
Consulta:
CSql = "Select * From ConstantesDeNomina Where IdConstante = " & IdsConstante(Z)
Set BdConstante = CrearRS(CSql)
If Not BdConstante.EOF Then
With BdConstante
If Not IsNull(.Fields("valorn")) Or Not (Trim(.Fields("valorn")) = "") Then Valor = .Fields("valorn")
'If Not IsNull(.Fields("valort")) Or Not (Trim(.Fields("valorT")) = "") Then valor = .Fields("valort")
'If Not IsNull(.Fields("valorf")) Or Not (Trim(.Fields("valorf")) = "") Then valor = .Fields("valorf")
End With
End If
BdConstante.Close
Return
End Sub
Sub verifica_concepto()
sCaracter = "Concepto("
Z = 1
dcon = ""
For h = 1 To Len(FormulA)
    cars = Mid$(FormulA, h, 9)
    T = 0
    If UCase(sCaracter) = UCase(cars) Then
    
    Formula1 = Mid(FormulA, h + 9, Len(FormulA))
    For w = 1 To Len(Formula1)
        d = Mid(Formula1, w, 1)
        If d = ";" Then T = 1
        
        If T = 1 And Not d = ";" And Not d = ")" Then dcon = dcon + d
        If d = ")" Then Exit For
    Next w
    If Trim(dcon) <> "" And Trim(dcon) <> "0" Then IdsConcepto(Z) = dcon: GoSub consulta5: Formula2 = Formula2 & Valor: Z = Z + 1: dcon = ""
        
    h = h + 9 + w
        
    End If
    Formula2 = Formula2 & Mid(FormulA, h, 1)
Next h
 'For df = 1 To Z - 1
 
FormulA = Formula2
 'MsgBox idsconstante(df)
 'Next df
Exit Sub
consulta5:
CSql = "Select * From Concepto Where IdConcepto = " & IdsConcepto(Z)
Set BdConstante = CrearRS(CSql)
If Not BdConstante.EOF Then
BdConstante.MoveFirst
Valor = BdConstante.Fields("formula")
End If
BdConstante.Close
Return
End Sub

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMM devuelve la FORMULA de un concepto  MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Function Validar_Concepto(ByRef Cadena As String) As String
sCaracter = "Concepto("
Z = 1
dcon = ""
For h = 1 To Len(Cadena)
    cars = Mid$(Cadena, h, 9)
    T = 0
    
    If UCase(sCaracter) = UCase(cars) Then
        Formula1 = Mid(Cadena, h + 9, Len(Cadena))
    
        For w = 1 To Len(Formula1)
            d = Mid(Formula1, w, 1)
            If d = ";" Then T = 1
            
            If T = 1 And Not d = ";" And Not d = ")" Then dcon = dcon + d
            If d = ")" Then Exit For
        Next w
    
        If Trim(dcon) <> "" And Trim(dcon) <> "0" Then
            IdsConcepto(Z) = dcon
            GoSub consulta5
            Formula2 = Formula2 & Valor: Z = Z + 1
            dcon = ""
            h = h + 9 + w
        End If
    End If
        Formula2 = Formula2 & Mid(Cadena, h, 1)
Next h
 'For df = 1 To Z - 1
 
Validar_Concepto = Formula2
 'MsgBox idsconstante(df)
 'Next df
Exit Function
consulta5:
CSql = "Select * From Concepto Where IdConcepto = " & IdsConcepto(Z)
Set BdConstante = CrearRS(CSql)
If BdConstante.RecordCount <> 0 Then
    BdConstante.MoveFirst
    Valor = BdConstante.Fields("formula")
Else
    Valor = 0
End If
BdConstante.Close
Return
End Function
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMM devuelve el valor de un CAMPO MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Function Validar_Campo(ByRef Cadena As String) As String
sCaracter = "Campo("
Z = 1
dcon = ""
For h = 1 To Len(Cadena)
    cars = Mid$(Cadena, h, 6)
    T = 0
    If UCase(sCaracter) = UCase(cars) Then
    
        Formula1 = Mid(Cadena, h + 6, Len(Cadena))
        For w = 1 To Len(Formula1)
            d = Mid(Formula1, w, 1)
            If d = ";" Then T = 1
        
            If T = 1 And Not d = ";" And Not d = ")" Then dcon = dcon + d
            If d = ")" Then Exit For
        Next w
        If Trim(dcon) <> "" And Trim(dcon) <> "0" Then IdsConcepto(Z) = dcon: GoSub consulta8: Formula2 = Formula2 & Valor: Z = Z + 1: dcon = ""
        
        h = h + 6 + w
    End If
    Formula2 = Formula2 & Mid(Cadena, h, 1)
Next h
Validar_Campo = Formula2
'For df = 1 To Z - 1
'MsgBox idsconcepto(df)
'Next df
Exit Function
consulta8:
CSql = "Select * From CamposDelTrabajador Where IdCampoNomina = " & IdsConcepto(Z) & _
        " And IdEmpleado = " & IdEmpl & " AND Tipo='CA'"
Set BdConstante = CrearRS(CSql)
If Not BdConstante.EOF Then
    With BdConstante
        If Not IsNull(.Fields("valorn")) Or Not (Trim(.Fields("valorn")) = "") Then
            Valor = .Fields("valorn").Value
        End If
        'If Not IsNull(.Fields("valort")) Or Not (Trim(.Fields("valorT")) = "") Then valor = .Fields("valort")
        'If Not IsNull(.Fields("valorf")) Or Not (Trim(.Fields("valorf")) = "") Then valor = .Fields("valorf")
    End With
Else
    Valor = 0
End If
BdConstante.Close
Return
End Function

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMM devuelve el VALOR de una constante MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Function Validar_Constante(ByVal Cadena As String) As String
sCaracter = "Constante("
Z = 1
dcon = ""
For h = 1 To Len(Cadena)
    cars = Mid$(Cadena, h, 10)
    T = 0
    If UCase(sCaracter) = UCase(cars) Then
       
        Formula1 = Mid(Cadena, h + 10, Len(Cadena))
        For w = 1 To Len(Formula1)
            d = Mid(Formula1, w, 1)
            If d = ";" Then T = 1
        
            If T = 1 And Not d = ";" And Not d = ")" Then dcon = dcon + d
            If d = ")" Then Exit For
        Next w
        If Trim(dcon) <> "" And Trim(dcon) <> "0" Then IdsConstante(Z) = dcon: GoSub Consulta: Formula2 = Formula2 & Valor: Z = Z + 1: dcon = ""
        
        h = h + 10 + w
                
    End If
    Formula2 = Formula2 & Mid(Cadena, h, 1)
Next h
'For df = 1 To Z - 1
Validar_Constante = Formula2
'MsgBox idsconstante(df)
'Next df
Exit Function
Consulta:
CSql = "Select * From ConstantesDeNomina Where IdConstante = " & IdsConstante(Z)
Set BdConstante = CrearRS(CSql)
If Not BdConstante.EOF Then
With BdConstante
If Not IsNull(.Fields("valorn")) Or Not (Trim(.Fields("valorn")) = "") Then
    Valor = Replace(.Fields("valorn"), ",", ".")
Else
    Valor = 0
End If
'If Not IsNull(.Fields("valort")) Or Not (Trim(.Fields("valorT")) = "") Then valor = .Fields("valort")
'If Not IsNull(.Fields("valorf")) Or Not (Trim(.Fields("valorf")) = "") Then valor = .Fields("valorf")
End With
End If
BdConstante.Close
Return
End Function

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMM devuelve el valor de una Funcion MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Function Validar_Funcion(Cadena As String, ByVal Period As Integer, Anio As String) As String
Dim Cad As String
Dim FncCad As String
Dim PosCad As Integer
Dim ICont As Integer
Dim NTemp As Integer

Cad = Cadena
ICont = 0
PosCad = 1

Cargar_Lista_De_Funciones

Call Calcular_Periodos(Format(CDate(Anio), "yyyy"))

For i = 0 To 23
    
    If IsNull(ListFunc(i, 0)) Then Exit For
    If Trim(ListFunc(i, 0)) = "" Then Exit For
    PosCad = 1
    While PosCad <> 0
        ICont = ICont + 1
        b = UCase(Cad)
        a = UCase(ListFunc(i, 0))
        PosCad = InStr(ICont, b, a, vbTextCompare)
        
        If PosCad <> 0 Then
            FncCad = Mid(Cad, PosCad, Len(a))
            
            If UCase(FncCad) = UCase(ListFunc(0, 0)) Then NTemp = LunesDelPeriodo(PyFs(Period - 1, 1), PyFs(Period - 1, 2))
            If UCase(FncCad) = UCase(ListFunc(1, 0)) Then NTemp = LunesDelMes(PyFs(Period - 1, 1))
            
            Cad = Mid(Cad, 1, PosCad - 1) & CStr(NTemp) & Mid(Cad, PosCad + Len(a))
        End If
        ICont = 0
    Wend
Next
Validar_Funcion = Cad

End Function

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Sub Cargar_Lista_De_Funciones()
    ListFunc(0, 0) = "(FuncLP)":    ListFunc(0, 1) = "Calcula la cantidad de LUNES del Periodo"
    ListFunc(1, 0) = "(FuncLM)":    ListFunc(1, 1) = "Calcula la cantidad de LUNES del Mes"
End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Function Validar_SSO(Cadena As String) As String
Dim FncCad As String
Dim PosCad As Integer
Dim ICont As Integer
Dim NTemp As Integer

ICont = 0
PosCad = 1

Cargar_Lista_De_Condicionales

For i = 0 To 23
    If IsNull(ListCond(i, 0)) Then Exit For
    If Trim(ListCond(i, 0)) = "" Then Exit For
    
    PosCad = 1
    While PosCad <> 0
        ICont = ICont + 1
        b = UCase(Cadena)
        a = UCase(ListCond(i, 0))
        PosCad = InStr(ICont, b, a, vbTextCompare)
        
        If PosCad <> 0 Then
            FncCad = Mid(Cadena, PosCad, Len(a))
            Cadena = Replace(Cadena, FncCad, "")
        End If
        ICont = 0
    Wend
Next

Validar_SSO = Cadena

End Function

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Funcion que devuelve 0 si el empleado gana menos de N salarios minimos
' para que no halla errores, el campo SUELDO MENSUAL debe tener el ID uno. IdCampoNomina=1
Public Function Calcular_SSO(ByVal Cadena As String, ByVal resultado As Double, ByVal IdEmpla As Integer) As Double
Dim SueldoMin As Double
Dim SueldoEmp As Double
Dim Resul As Double
Dim ValorMult As Byte
Dim PosB As Integer
Dim i As Integer
Dim ConVal As Integer

CSql = "SELECT DATEDIFF(day, '" & Format(Now, "dd/mm/yy") & "', Anio) AS DiffDate, SueldoM, Valor From Sueldo_Minimo ORDER BY Anio"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Calcular_SSO = CDbl(resultado): Exit Function


RsTemp.MoveFirst
ConVal = -1
While Not ConVal > 0
    ConVal = Val(RsTemp.Fields("DiffDate").Value)
    
    
    If ConVal <= 0 Then
        If Not IsNull(RsTemp.Fields("SueldoM").Value) Then
            SueldoMin = CDbl(RsTemp.Fields("SueldoM").Value)
            ValorMult = Val(RsTemp.Fields("Valor").Value)
        Else
            Calcular_SSO = CDbl(resultado)
            Exit Function
        End If
        RsTemp.MoveNext
        
        If RsTemp.EOF Then
            ConVal = 1
        Else
            ConVal = Val(RsTemp.Fields("DiffDate").Value)
            If ConVal > 0 Then
                ConVal = 1
            End If
        End If
    Else
        ConVal = 1
    End If
Wend

' consulta para saber el sueldo del empleado
CSql = "SELECT ValorN FROM CamposDelTrabajador WHERE IdCampoNomina=1 AND Tipo='CA' AND IdEmpleado=" & IdEmpla
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("ValorN").Value) Then
    SueldoEmp = CDbl(RsTemp.Fields("ValorN").Value)
Else
    Calcular_SSO = resultado
    Exit Function
End If

Resul = (SueldoMin * ValorMult)

PosB = 0

If (SueldoEmp >= Resul) Then
    
    For i = 1 To Len(Cadena)
        PosB = InStr(i, Cadena, SueldoEmp)
        If PosB Then
            Cadena = Mid(Cadena, 1, PosB - 1) & Resul & Mid(Cadena, PosB + Len(CStr(SueldoEmp)))
            'Cadena = Mid(Cadena, PosB, Len(CStr(SueldoEmp))) & Resul & Mid(Cadena, PosB + Len(CStr(SueldoEmp)))
            Exit For
        End If
    Next i
    resultado = FrmPrincipal.ScriptControl1.Eval(Cadena)
    Calcular_SSO = CDbl(resultado)
Else
    Calcular_SSO = resultado
End If

End Function

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Sub Cargar_Lista_De_Condicionales()
    ListCond(0, 0) = "(CondSSO)":    ListCond(0, 1) = "Condicion del SSO mayor a N Salarios minimos"
    'ListCond(1, 0) = "(FuncLM)":    ListCond(1, 1) = "Calcula la cantidad de LUNES del Mes"
End Sub

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Sub Calcular_Periodos(ByRef Anio As String)
Dim FechaT As String

FechaT = "01/01/" & Anio 'Format(DTPicker1.Value, "yyyy")
i = 0
While i < 24
    
    i = i + 1
    PyFs(i - 1, 0) = i
    PyFs(i - 1, 1) = "01/" & Format(CDate(FechaT), "MM/yyyy")
    PyFs(i - 1, 2) = "15/" & Format(CDate(FechaT), "MM/yyyy")
    
    i = i + 1
    PyFs(i - 1, 0) = i
    PyFs(i - 1, 1) = "16/" & Format(CDate(FechaT), "MM/yyyy")
    PyFs(i - 1, 2) = Format(DateSerial(Year(CDate(FechaT)), Month(CDate(FechaT)) + 1, 0), "dd/MM/yyyy")
    
    FechaT = "01/" & Month(CDate(FechaT)) + 1 & "/" & Year(CDate(FechaT))
Wend

End Sub
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


Sub Quitar(CArac2)
sCaracter = "., "
stmp = ""
If IsNull(CArac2) Then CArac2 = 0
For h = 1 To Len(CArac2)
      If InStr(sCaracter, Mid$(CArac2, h, 1)) = 0 Then
          stmp = stmp & Mid$(CArac2, h, 1)
      Else
          stmp = stmp & "."
      End If
Next h

If stmp = "" Then stmp = 0
Carac = stmp
End Sub
Sub QuitarCaracter(CArac1)
sCaracter = ". "
stmp = ""
If IsNull(CArac1) Then CArac1 = 0
For h = 1 To Len(CArac1)
      If InStr(sCaracter, Mid$(CArac1, h, 1)) = 0 Then
          stmp = stmp & Mid$(CArac1, h, 1)
      End If
Next h

If stmp = "" Then stmp = 0
Carac = stmp
End Sub
Function Num_texto(Numero)
Dim Texto
Dim Millones
Dim Miles
Dim Cientos
Dim Decimales
Dim StrText
Dim CadMillones
Dim CadMiles
Dim CadCientos
Texto = Numero
Texto = FormatNumber(Texto, 2)
Texto = Right(Space(14) & Texto, 14)
Millones = Mid(Texto, 1, 3)
Miles = Mid(Texto, 5, 3)
Cientos = Mid(Texto, 9, 3)
Decimales = Mid(Texto, 13, 2)
CadMillones = ConvierteCifra(Millones, 1)
CadMiles = ConvierteCifra(Miles, 1)
CadCientos = ConvierteCifra(Cientos, 0)
If Trim(CadMillones) > "" Then
    If Trim(CadMillones) = "UN" Then
        StrText = CadMillones & " MILLON"
    Else
        StrText = CadMillones & " MILLONES"
    End If
End If
If Trim(CadMiles) > "" Then
    StrText = StrText & " " & CadMiles & " MIL"
End If
If Trim(CadMiles & CadCientos) = "UN" Then
    StrText = StrText & "UNO CON " & Decimales & "/100"
Else
    If Miles & Cientos = "000000" Then
        StrText = StrText & " " & Trim(CadCientos) & " " & Decimales & "/100"
    Else
        StrText = StrText & " " & Trim(CadCientos) & " " & Decimales & "/100"
    End If
End If
Num_texto = Trim(StrText)
End Function
Function ConvierteCifra(Texto, SW)
Dim Centena
Dim Decena
Dim Unidad
Dim txtCentena
Dim txtDecena
Dim txtUnidad
Centena = Mid(Texto, 1, 1)
Decena = Mid(Texto, 2, 1)
Unidad = Mid(Texto, 3, 1)
Select Case Centena
    Case "1"
        txtCentena = "CIEN"
        If Decena & Unidad <> "00" Then
            txtCentena = "CIENTO"
        End If
    Case "2"
        txtCentena = "DOSCIENTOS"
    Case "3"
        txtCentena = "TRESCIENTOS"
    Case "4"
        txtCentena = "CUATROCIENTOS"
    Case "5"
        txtCentena = "QUINIENTOS"
    Case "6"
        txtCentena = "SEISCIENTOS"
    Case "7"
        txtCentena = "SETECIENTOS"
    Case "8"
        txtCentena = "OCHOCIENTOS"
    Case "9"
        txtCentena = "NOVECIENTOS"
End Select

Select Case Decena
    Case "1"
        txtDecena = "DIEZ"
        Select Case Unidad
            Case "1"
                txtDecena = "ONCE"
            Case "2"
                txtDecena = "DOCE"
            Case "3"
                txtDecena = "TRECE"
            Case "4"
                txtDecena = "CATORCE"
            Case "5"
                txtDecena = "QUINCE"
            Case "6"
                txtDecena = "DIECISEIS"
            Case "7"
                txtDecena = "DIECISIETE"
            Case "8"
                txtDecena = "DIECIOCHO"
            Case "9"
                txtDecena = "DIECINUEVE"
        End Select
    Case "2"
        txtDecena = "VEINTE"
        If Unidad <> "0" Then
            txtDecena = "VEINTI"
        End If
    Case "3"
        txtDecena = "TREINTA"
        If Unidad <> "0" Then
            txtDecena = "TREINTA Y "
        End If
    Case "4"
        txtDecena = "CUARENTA"
        If Unidad <> "0" Then
            txtDecena = "CUARENTA Y "
        End If
    Case "5"
        txtDecena = "CINCUENTA"
        If Unidad <> "0" Then
            txtDecena = "CINCUENTA Y "
        End If
    Case "6"
        txtDecena = "SESENTA"
        If Unidad <> "0" Then
            txtDecena = "SESENTA Y "
        End If
    Case "7"
        txtDecena = "SETENTA"
        If Unidad <> "0" Then
            txtDecena = "SETENTA Y "
        End If
    Case "8"
        txtDecena = "OCHENTA"
        If Unidad <> "0" Then
            txtDecena = "OCHENTA Y "
        End If
    Case "9"
        txtDecena = "NOVENTA"
        If Unidad <> "0" Then
            txtDecena = "NOVENTA Y "
        End If
End Select

If Decena <> "1" Then
    Select Case Unidad
        Case "1"
            If SW Then
                txtUnidad = "UN"
            Else
                txtUnidad = "UNO"
            End If
        Case "2"
            txtUnidad = "DOS"
        Case "3"
            txtUnidad = "TRES"
        Case "4"
            txtUnidad = "CUATRO"
        Case "5"
            txtUnidad = "CINCO"
        Case "6"
            txtUnidad = "SEIS"
        Case "7"
            txtUnidad = "SIETE"
        Case "8"
            txtUnidad = "OCHO"
        Case "9"
            txtUnidad = "NUEVE"
    End Select
End If
ConvierteCifra = txtCentena & " " & txtDecena & txtUnidad
End Function

Sub Sonar_Timbre()
PortOut &H378, 1
End Sub

Sub Apagar_Timbre()
PortOut &H378, 0
End Sub

Sub VerificaPaciente()
If IdPac1 = "" Then Exit Sub

Dim BdLista8 As New ADODB.Recordset
CSql = "Select * From Ubi_Paciente Where IdPaciente = " & IdPac1
Set BdLista8 = CrearRS(CSql)
If Not BdLista8.EOF Then
    Select Case BdLista8.Fields("modul")
    Case Is = 0
    m = "Nutrición"
    Case Is = 1
    m = "Psicología"
    Case Is = 2
    m = "Tratamiento de Radioterapia"
    Case Is = 3
    m = "Dirección Médica"
    Case Is = 4
    m = "Oncología Radioterapeuta"
    Case Is = 5
    m = "Administración"
    End Select
    If BdLista8.Fields("modul") <> ModulO Then
        Msg = "El paciente está siendo atendido en " & m
        MsgBox Msg, vbOKOnly + vbCritical, "Paciente Ocupado"
        Cnt = True
        BdLista8.Close
        Exit Sub
    End If
End If

End Sub
Sub Llamar()
Dim RsTll As New ADODB.Recordset
Dim RsTl2 As New ADODB.Recordset
Dim RsTl3 As New ADODB.Recordset

CSql = "Select * From llamado1"
Set RsTll = CrearRS(CSql)
C = RsTll.RecordCount + 1
RsTll.Close

CSql = "Select * From llamado2"
Set RsTl2 = CrearRS(CSql)
d = RsTl2.RecordCount + 1
RsTl2.Close


CSql = "Select * From llamado45"
Set RsTl3 = CrearRS(CSql)
e = RsTl3.RecordCount + 1
RsTl3.Close

'<<<<se verifica si el paciente esta siendo atendido por otro usuario>>>>>>>>>
Cnt = False
Call VerificaPaciente

If Cnt = True Then Exit Sub

'<<<<si no esta siendo atendido se asientan registros necesarios para el llamado por pantalla en el FrmLlamador
If IdPac1 <> "" Then

    '<<< se blanque las tablas de llamado1 y 2 para luego registrar el llamdo
    
    Dim BdLlamado As New ADODB.Recordset
    CSql = "Delete From Llamado1 Where Modulo = " & ModulO
    Set BdLlamado = CrearRS(CSql)
    
    Dim BdLlamado1 As New ADODB.Recordset
    CSql = "Delete From Llamado2 Where Modulo = " & ModulO
    Set BdLlamado1 = CrearRS(CSql)
    DoEvents
    
    '<<<se registra el llamado en cada una de las pantallas en blanco
    CSql = "Insert Into Llamado1(IdLlamado, Modulo, IdPaciente, Pantalla, MiniLlamador) Values(" & C & "," & ModulO & ",0,0,0)"
    Dim BdLlamado2 As New ADODB.Recordset
    Set BdLlamado2 = CrearRS(CSql)
    Call Espera(2)
    
    CSql = "Insert Into Llamado2(IdLlamado, Modulo, IdPaciente, Pantalla, MiniLlamador) Values(" & d & "," & ModulO & ",0,0,0)"
    Set BdLlamado2 = CrearRS(CSql)
    Call Espera(2)
    
    CSql = "Delete From Llamado1 Where Modulo = " & ModulO
    Set BdLlamado = CrearRS(CSql)
    
    CSql = "Delete From Llamado2 Where Modulo = " & ModulO
    Set BdLlamado = CrearRS(CSql)
    
    CSql = "Insert Into Llamado1(IdLlamado,IdPaciente, Modulo,Pantalla, MiniLlamador) Values(" & C & "," & IdPac1 & "," & ModulO & ",0,0)"
    Set BdLlamado2 = CrearRS(CSql)
    
    CSql = "Insert Into Llamado2(IdLlamado,IdPaciente, Modulo,Pantalla, MiniLlamador) Values(" & d & "," & IdPac1 & "," & ModulO & ",0,0)"
    Set BdLlamado = CrearRS(CSql)
    
    CSql = "Delete From Ubi_Paciente Where Modul = " & ModulO
    Dim BdLista8 As New ADODB.Recordset
    Set BdLista8 = CrearRS(CSql)
    
    CSql = "Insert Into Ubi_Paciente(Modul, IdPaciente) Values(" & ModulO & "," & IdPac1 & ")"
    Dim bdlista9 As New ADODB.Recordset
    Set bdlista9 = CrearRS(CSql)

Else
    Msg = "No Ha seleccionado algun paciente, para hacer el llamado." & Chr(13) & " Selecciones alguno e intente nuevamente"
    MsgBox Msg, vbOKOnly + vbCritical, "Sin paciente"
End If
End Sub
Sub Espera(Segundos As Single)
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = Timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > Timer
      DoEvents
      If ComienzoSeg > Timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop
End Sub

Function Centrar(frm As Form)
  'Centra el formulario en la pantalla
  frm.Left = (FrmPrincipal.Width - frm.Width) / 2
  frm.Top = (FrmPrincipal.Height - frm.Height) / 2 - 670
End Function

Public Function FechaSQL(ByVal vFecha As String) As String
 'La fecha la convierte al formato: #yyyy/mm/dd#
 On Local Error GoTo SQLDateValErr
 If IsDate(vFecha) Then
   'si es una fecha válida, convertirla
   FechaSQL = "#" & Format$(vFecha, "yyyy/mm/dd") & "#"
 Else
   'si no es una fecha válida, devolverlo sin modificar
   FechaSQL = vFecha
 End If
 Exit Function
    
SQLDateValErr:
 'Si hay error, la fecha por defecto 1-Ene-1980
 Err = 0
 FechaSQL = "#1980/01/01#"
End Function


Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
 'Permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
 If InStr("0123456789/-", Chr(KeyAscii)) = 0 Then
   SoloNumeros = 0
 Else
   SoloNumeros = KeyAscii
 End If
 'Teclas especiales permitidas
 If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
 If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Function IsEmptyRecordset(rs As Recordset) As Boolean
  IsEmptyRecorSet = ((rs.BOF = True) And (rs.EOF = True))
End Function

Sub Enviar_Bitacora(ByVal P_IdUser As String, ByVal P_Modulo As String, ByVal P_Accion As String, ByVal P_Nota As String)
Dim RsBitacora As New ADODB.Recordset
Dim Instruc As String
Dim P_Codig As String

Instruc = "Select MAX(Codigo)+1 as NuevoCod FROM Arocatib"
Set RsBitacora = CrearRS(Instruc)

If RsBitacora.RecordCount <> 0 Then
    If Not RsBitacora.Fields("NuevoCod").Value = Null Then
            P_Codig = Val(RsBitacora.Fields("NuevoCod").Value)
        Else
            P_Codig = "1"
    End If
Else
    P_Codig = "1"
End If

Instruc = "INSERT INTO Arocatib VALUES ('" & P_Codig & "','" & P_IdUser & "','" & Format(Now, "DD/MM/YYYY") & _
            "','" & DateTime.Time & "','" & P_Modulo & "','" & P_Accion & "','" & P_Nota & "')"
Set RsBitacora = CrearRS(Instruc)
End Sub


Function Validar_Camara(ByVal a As String, ByVal b As Long, ByVal C As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, ByVal g As Long, ByVal h As Integer) As Boolean

hwndc = capCreateCaptureWindow(a, b, C, d, e, f, g, h)

If (hwndc <> 0) Then
    temp = SendMessage(hwndc, WM_cap_driver_connect, 0, 0)
    temp2 = SendMessage(hwndc, WM_CAP_DRIVER_DISCONNECT, 0, 0)
    
    If temp = 0 Then
        Validar_Camara = False
        Else
        Validar_Camara = True
    End If
End If

End Function

Function GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
Dim sBuffer As String, lRet As Long
sBuffer = String$(255, 0)

lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
       
If lRet = 0 Then
GetFromINI = sDefault
Else
GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
End If
End Function

Public Function Gramatica(Cad As String, Pos As String) As TextoPos
Dim TamCad As Integer
Dim i, J As Integer

TamCad = Len(Cad)

Cad = UCase(Mid(Cad, 1, 1)) & LCase(Mid(Cad, 2, TamCad - 1))

For i = 1 To TamCad
    
    If i > TamCad Then Exit For
    
    ' Coloca un punto "." antes del caracter ENTER
    If Asc(Mid(Cad, i, 1)) = 13 Then
        For J = i - 1 To 1 Step -1
            If Mid(Cad, J, 1) = "." Then Exit For
            If Not Mid(Cad, J, 1) = " " And Not Asc(Mid(Cad, J, 1)) = 10 And Not Asc(Mid(Cad, J, 1)) = 13 Then
                Cad = Mid(Cad, 1, J) & "." & Mid(Cad, J + 1, TamCad)
                TamCad = Len(Cad)
                Pos = Pos + 1
                Exit For
            End If
        Next J
    End If
    
    ' Eliminar los espacios en blancos que se encuentran ANTES de un punto "." y crea un espacio
    ' en blanco delante del mismo en el caso de que no lo tenga
    If Mid(Cad, i, 1) = "." Then
        For J = i - 1 To 1 Step -1
            If Mid(Cad, J, 1) = " " Then
                Cad = Mid(Cad, 1, J - 1) & Replace(Cad, " ", "", J, 1)
                TamCad = Len(Cad)
                i = i - 1
                Pos = Pos - 1
                Else
                Exit For
            End If
        Next J
        
        ' Coloca el espacio en Blanco despues de un punto
        'If Not Mid(Cad, i + 1, 1) = " " Then
        '    Cad = Mid(Cad, 1, i) & " " & Mid(Cad, i + 1, TamCad)
        '    TamCad = Len(Cad)
        '    Pos = Pos + 1
        'End If
        
    End If
    
    ' Cambia a Mayuscula la letra despues de un punto o un ENTER,
    If (Mid(Cad, i, 1) = ".") Or (Asc(Mid(Cad, i, 1)) = 10) Then
        For J = i + 1 To TamCad
            If Not Mid(Cad, J, 1) = " " And Not Asc(Mid(Cad, J, 1)) = 10 Then
                Cad = Mid(Cad, 1, J - 1) & UCase(Mid(Cad, J, 1)) & LCase(Mid(Cad, J + 1, TamCad))
                Exit For
            End If
        Next J
    End If
Next i

Gramatica.Texto = Cad
Gramatica.Poscur = Pos

End Function

Public Function Verificar_Internet() As Boolean
On Error GoTo ErrorL
Dim StrCn2  As String
Dim Cnn2 As New ADODB.Connection

StrCn2 = "Driver={SQL Server}; Server=oa.no-ip.org; Database=OACLINICA; UID=sa; PWD=458921957JAr;"

If Cnn2.State = 0 Then
    Cnn2.ConnectionString = StrCn2
    Cnn2.Open
    Verificar_Internet = True
End If

Exit Function

ErrorL:
    Verificar_Internet = False

End Function


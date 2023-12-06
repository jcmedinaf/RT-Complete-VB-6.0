VERSION 5.00
Begin VB.Form FrmIVSSAdmEquipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipos"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "FrmIVSSAdmEquipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton BtnBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton BtnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton BtnAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox CboCategorias 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox TxtAnio 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox TxtIVSSID 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox TxtNumeroSerie 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox TxtModelo 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox TxtMarca 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.ComboBox CboHospital 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Categoria:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2820
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Año:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2445
      Width           =   330
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "IVSS I.D:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2085
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Número de Serie:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1725
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hospital:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "FrmIVSSAdmEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IdEquipo As String
Private Sub BtnAgregar_Click()
Editar = 0
Limpiar_Campos

End Sub

Sub Limpiar_Campos()
CboCategorias.ListIndex = -1
CboHospital.ListIndex = -1
TxtDescripcion.Text = ""
TxtMarca.Text = ""
TxtModelo.Text = ""
TxtNumeroSerie.Text = ""
TxtAnio.Text = ""
TxtIVSSID.Text = ""
End Sub

Private Sub BtnBuscar_Click()
Dim Buscar As String
Dim RsBuscar As New ADODB.Recordset

Buscar = InputBox("Ingrese el codigo del equipo", "Busqueda")

If Buscar = "" Then
    MsgBox "tiene que ingresar un codigo para realizar la buqueda!!!", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

ConectarIVSSHosting

CSql = "Select * From Equipo Where IdEquipo=" & Buscar & ""
RsBuscar.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic


    Editar = 1
    IdEquipo = RsBuscar.Fields("IdEquipo").Value
    TxtDescripcion.Text = RsBuscar.Fields("descripcion").Value
    TxtModelo.Text = RsBuscar.Fields("Modelo").Value
    TxtMarca.Text = RsBuscar.Fields("Marca").Value
    TxtAnio.Text = RsBuscar.Fields("Ano").Value
    TxtNumeroSerie.Text = RsBuscar.Fields("numero_serial").Value
    TxtIVSSID.Text = RsBuscar.Fields("reg_ivss").Value
    

WebCnn.Close
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
Dim RsGuardar As New ADODB.Recordset

If CboHospital.ListIndex = -1 Then
    Msg = "seleccione el hospital donde se encuentra el equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtDescripcion.Text = "" Then
    Msg = "Ingrese la descripcion del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtMarca.Text = "" Then
    Msg = "Ingrese la marca del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtModelo.Text = "" Then
    Msg = "Ingrese el modelo del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtNumeroSerie.Text = "" Then
    Msg = "Ingrese el numero de serie del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtIVSSID.Text = "" Then
    Msg = "Ingrese el IVSS I.D del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtAnio.Text = "" Then
    Msg = "Ingrese el año del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If CboCategorias.ListIndex = -1 Then
    Msg = "seleccione la categoria del equipo"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

Select Case Editar

Case Is = 0


'    ConectarIVSSHosting
'    CSql = "Select * From Equipo"
'    RsGuardar.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
'
'    RsGuardar.AddNew
'
'    RsGuardar.Fields("idsubcategoria").Value = CboCategorias.ItemData(CboCategorias.ListIndex)
'    RsGuardar.Fields("idhospital").Value = CboHospital.ItemData(CboHospital.ListIndex)
'    RsGuardar.Fields("descripcion").Value = TxtDescripcion.Text
'    RsGuardar.Fields("marca").Value = TxtMarca.Text
'    RsGuardar.Fields("modelo").Value = TxtModelo.Text
'    RsGuardar.Fields("numero_serial").Value = TxtNumeroSerie.Text
'    RsGuardar.Fields("ano").Value = TxtAnio.Text
'    RsGuardar.Fields("reg_ivss").Value = TxtIVSSID.Text
'    RsGuardar.Update
'
'    WebCnn.Close
    
    MsgBox "Registro Guardado Correctamente!!!", vbOKOnly + vbInformation, "Guardado Exitoso"
Case Is = 1

    
'    ConectarIVSSHosting
'    CSql = "Select * From Equipo Where IdEquipo='" & IdEquipo & "'"
'    RsGuardar.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
'
'
'    RsGuardar.Fields("idsubcategoria").Value = CboCategorias.ItemData(CboCategorias.ListIndex)
'    RsGuardar.Fields("idhospital").Value = CboHospital.ItemData(CboHospital.ListIndex)
'    RsGuardar.Fields("descripcion").Value = TxtDescripcion.Text
'    RsGuardar.Fields("marca").Value = TxtMarca.Text
'    RsGuardar.Fields("modelo").Value = TxtModelo.Text
'    RsGuardar.Fields("numero_serial").Value = TxtNumeroSerie.Text
'    RsGuardar.Fields("ano").Value = TxtAnio.Text
'    RsGuardar.Fields("reg_ivss").Value = TxtIVSSID.Text
'    RsGuardar.Update
'
'    WebCnn.Close
    MsgBox "Registro Actualizado Correctamente!!!", vbOKOnly + vbInformation, "Actualido Exitoso"
End Select
Limpiar_Campos
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim RsListado As New ADODB.Recordset

ConectarIVSSHosting
    
CSql = "Select * From Hospitales"
'CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1 FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE (subcategoria.idsubcategoria =" & Pn & ")"
RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
Do While Not RsListado.EOF
    With CboHospital
        .AddItem RsListado.Fields("Hospital").Value
        .ItemData(.NewIndex) = RsListado.Fields("IdHospital").Value
    End With
    RsListado.MoveNext
Loop
    WebCnn.Close

ConectarIVSSHosting
    
CSql = "Select * From SubCategoria"
'CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1 FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE (subcategoria.idsubcategoria =" & Pn & ")"
RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
Do While Not RsListado.EOF
    With CboCategorias
        .AddItem RsListado.Fields("SubCategoriac").Value
        .ItemData(.NewIndex) = RsListado.Fields("IdSubCategoria").Value
    End With
    RsListado.MoveNext
Loop
    WebCnn.Close
End Sub

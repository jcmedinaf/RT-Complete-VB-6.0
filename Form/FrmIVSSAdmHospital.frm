VERSION 5.00
Begin VB.Form FrmIVSSAdmHospital 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospital"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   Icon            =   "FrmIVSSAdmHospital.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboEstados 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox TxtHospital 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox TxtDireccion 
      Height          =   645
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox TxtTelefono 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
   Begin VB.ComboBox CboCiudad 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton BtnAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton BtnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton BtnBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton BtnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hospital:"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefonos:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   750
   End
End
Attribute VB_Name = "FrmIVSSAdmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
Editar = 0
Limpiar_Campos
End Sub

Private Sub BtnBuscar_Click()
Editar = 1
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
Dim RsGuardar As New ADODB.Recordset

If CboEstado.ListIndex = -1 Then
    Msg = "seleccione el estado donde se encuentra el hospital"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If CboCiudad.ListIndex = -1 Then
    Msg = "seleccione la ciudad donde se encuentra el hospital"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtHospital.Text = "" Then
    Msg = "ingrese el nombre del Hospital"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtDireccion.Text = "" Then
    Msg = "ingrese la direccion del hospital"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtTelefono.Text = "" Then
    Msg = "ingrese el numero de telefono del hospital"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

Select Case Editar

Case Is = 0


'    ConectarIVSSHosting
'    CSql = "Select * From Hospitales"
'    RsGuardar.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
'
'    RsGuardar.AddNew
'
'    RsGuardar.Fields("idsubcategoria").Value = CboEstados.ItemData(CboEstados.ListIndex)
'    RsGuardar.Fields("idhospital").Value = CboCiudad.ItemData(CboCiudad.ListIndex)
'    RsGuardar.Fields("descripcion").Value = TxtHospital.Text
'    RsGuardar.Fields("marca").Value = TxtDireccion.Text
'    RsGuardar.Fields("modelo").Value = TxtTelefono.Text
'    RsGuardar.Update
'
'    WebCnn.Close
'
    MsgBox "Registro Guardado Correctamente!!!", vbOKOnly + vbInformation, "Guardado Exitoso"
Case Is = 1


'    ConectarIVSSHosting
'    CSql = "Select * From Equipo Where IdEquipo='" & IdEquipo & "'"
'    RsGuardar.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
'
'
'    RsGuardar.Fields("idsubcategoria").Value = CboEstados.ItemData(CboEstados.ListIndex)
'    RsGuardar.Fields("idhospital").Value = CboCiudad.ItemData(CboCiudad.ListIndex)
'    RsGuardar.Fields("descripcion").Value = TxtHospital.Text
'    RsGuardar.Fields("marca").Value = TxtDireccion.Text
'    RsGuardar.Fields("modelo").Value = TxtTelefono.Text
'    RsGuardar.Update
'
'    WebCnn.Close
    MsgBox "Registro Actualizado Correctamente!!!", vbOKOnly + vbInformation, "Actualido Exitoso"
End Select

Limpiar_Campos
End Sub

Private Sub CboEstados_Click()
ConectarIVSSHosting
Dim RsCiudad As New ADODB.Recordset

CSql = "Select * From Ciudad Where IdEstado ='" & CboEstados.ItemData(CboEstados.ListIndex) & "' Order By Ciudad"
RsCiudad.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic


Do While Not RsCiudad.EOF
    With CboCiudad
        .AddItem RsCiudad.Fields("Ciudad").Value
        .ItemData(.NewIndex) = RsCiudad.Fields("IdCiudad").Value
    End With
    RsCiudad.MoveNext
Loop
WebCnn.Close
End Sub

Private Sub Form_Load()
ConectarIVSSHosting
Dim RsEstados As New ADODB.Recordset

CSql = "Select * From Estado Order By Estado"
RsEstados.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic


Do While Not RsEstados.EOF
    With CboEstados
        .AddItem RsEstados.Fields("Estado").Value
        .ItemData(.NewIndex) = RsEstados.Fields("IdEstado").Value
    End With
    RsEstados.MoveNext
Loop
WebCnn.Close
End Sub
Sub Limpiar_Campos()
TxtHospital.Text = ""
TxtDireccion.Text = ""
TxtTelefono.Text = ""
CboCiudad.ListIndex = -1
CboEstados.ListIndex = -1
End Sub

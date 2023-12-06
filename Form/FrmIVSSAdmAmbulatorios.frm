VERSION 5.00
Begin VB.Form FrmIVSSAdmAmbulatorios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ambulatorios"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "FrmIVSSAdmAmbulatorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCentroAmbulatorio 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox TxtDireccion 
      Height          =   645
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox TxtTelefono 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   4095
   End
   Begin VB.ComboBox CboCiudad 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton BtnAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton BtnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton BtnBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton BtnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Centro Ambulatorio:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefonos:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   750
   End
End
Attribute VB_Name = "FrmIVSSAdmAmbulatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
Editar = 0
Limpiar_Campos
End Sub
Sub Limpiar_Campos()
TxtCentroAmbulatorio.Text = ""
TxtDireccion.Text = ""
TxtTelefono.Text = ""
CboCiudadListIndex = -1

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
Dim RsGuardar As New ADODB.Recordset


If CboCiudad.ListIndex = -1 Then
    Msg = "Seleccione una ciudad"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtCentroAmbulatorio.Text = "" Then
    Msg = "ingrese el nombre del ambulatorio"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtDireccion.Text = "" Then
    Msg = "ingrese la direccion del ambulatorio"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtTelefono.Text = "" Then
    Msg = "ingrese el numero de telefono del ambulatorio"
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
'    RsGuardar.Fields("idsubcategoria").Value = CboCiudad.ItemData(CboCiudad.ListIndex)
'    RsGuardar.Fields("reg_ivss").Value = TxtCentroAmbulatorio.Text
'    RsGuardar.Fields("reg_ivss").Value = TxtDireccion.Text
'    RsGuardar.Fields("reg_ivss").Value = TxtTelefono.Text
'
'    RsGuardar.Update
'
'    WebCnn.Close
    
    MsgBox "Registro Guardado Correctamente!!!", vbOKOnly + vbInformation, "Guardado Exitoso"
Case Is = 1

'
'    ConectarIVSSHosting
'    CSql = "Select * From Equipo Where IdEquipo='" & IdEquipo & "'"
'    RsGuardar.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
'
'
'    RsGuardar.Fields("idsubcategoria").Value = CboCiudad.ItemData(CboCiudad.ListIndex)
'    RsGuardar.Fields("reg_ivss").Value = TxtCentroAmbulatorio.Text
'    RsGuardar.Fields("reg_ivss").Value = TxtDireccion.Text
'    RsGuardar.Fields("reg_ivss").Value = TxtTelefono.Text
'
'    RsGuardar.Update
'
'    WebCnn.Close
    MsgBox "Registro Actualizado Correctamente!!!", vbOKOnly + vbInformation, "Actualido Exitoso"
End Select
    Limpiar_Campos
End Sub


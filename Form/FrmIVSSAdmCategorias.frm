VERSION 5.00
Begin VB.Form FrmIVSSAdmCategorias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categorias"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   Icon            =   "FrmIVSSAdmCategorias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSubCategoria 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   4095
   End
   Begin VB.ComboBox CboCategoria 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton BtnAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton BtnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton BtnBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton BtnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SubCategoria:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Categoria:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmIVSSAdmCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
Editar = 0
Limpiar_Campos
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub
 
Sub Limpiar_Campos()
TxtSubCategoria.Text = ""
CboCategoriaListIndex = -1
End Sub

Private Sub BtnGuardar_Click()
Dim RsGuardar As New ADODB.Recordset

If CboCategoria.ListIndex = -1 Then
    Msg = "seleccione la categoria"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtSubCategoria.Text = "" Then
    Msg = "Ingrese el nombre o descripcion para la subcategoria"
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
'    RsGuardar.Fields("reg_ivss").Value = TxtSubCategoria.Text
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
'    RsGuardar.Fields("idsubcategoria").Value = CboCategorias.ItemData(CboCategorias.ListIndex)
'    RsGuardar.Fields("reg_ivss").Value = TxtSubCategoria.Text
'
'    RsGuardar.Update
'
'    WebCnn.Close
    MsgBox "Registro Actualizado Correctamente!!!", vbOKOnly + vbInformation, "Actualido Exitoso"
End Select
Limpiar_Campos

End Sub

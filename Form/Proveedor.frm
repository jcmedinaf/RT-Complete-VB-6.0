VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmProveedores 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   5490
   ClientLeft      =   6180
   ClientTop       =   2490
   ClientWidth     =   8250
   Icon            =   "Proveedor.frx":0000
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8250
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   4560
         Width           =   7815
         Begin ChamaleonButton.ChameleonBtn BtnAgregar 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":1002
            PICN            =   "Proveedor.frx":101E
            PICH            =   "Proveedor.frx":11AB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBorrar 
            Height          =   375
            Left            =   2520
            TabIndex        =   20
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Borrar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":13E0
            PICN            =   "Proveedor.frx":13FC
            PICH            =   "Proveedor.frx":15A0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6720
            TabIndex        =   16
            ToolTipText     =   "Cerrar"
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Cerrar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":173F
            PICN            =   "Proveedor.frx":175B
            PICH            =   "Proveedor.frx":1924
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardar 
            Height          =   375
            Left            =   1320
            TabIndex        =   17
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Guardar"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":1B59
            PICN            =   "Proveedor.frx":1B75
            PICH            =   "Proveedor.frx":1E04
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnDesHacer 
            Height          =   375
            Left            =   5520
            TabIndex        =   19
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Deshacer"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":2245
            PICN            =   "Proveedor.frx":2261
            PICH            =   "Proveedor.frx":2543
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
            Height          =   375
            Left            =   4560
            TabIndex        =   21
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":2794
            PICN            =   "Proveedor.frx":27B0
            PICH            =   "Proveedor.frx":2A46
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAnterior 
            Height          =   375
            Left            =   3960
            TabIndex        =   22
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":2CA5
            PICN            =   "Proveedor.frx":2CC1
            PICH            =   "Proveedor.frx":2F56
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Proveedor"
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7815
         Begin VB.ComboBox CboStatus 
            Height          =   315
            ItemData        =   "Proveedor.frx":31B2
            Left            =   5880
            List            =   "Proveedor.frx":31B4
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Agente de Retención"
            Height          =   255
            Left            =   1200
            TabIndex        =   30
            Top             =   3960
            Width           =   2535
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   4560
            Top             =   3000
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Filtro de Busqueda"
            Height          =   735
            Left            =   4560
            TabIndex        =   25
            Top             =   3480
            Width           =   3135
            Begin VB.TextBox TxtBuscar 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000A&
               Height          =   375
               Left            =   120
               TabIndex        =   26
               Text            =   "Busqueda"
               ToolTipText     =   "Ingrese la busqueda por Nombre, Razon Social, Cédula o Rif"
               Top             =   240
               Width           =   1455
            End
            Begin ChamaleonButton.ChameleonBtn BtnBuscar 
               Height          =   375
               Left            =   1680
               TabIndex        =   27
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Busqueda"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   14215660
               BCOLO           =   14215660
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   16777215
               MPTR            =   1
               MICON           =   "Proveedor.frx":31B6
               PICN            =   "Proveedor.frx":31D2
               PICH            =   "Proveedor.frx":3437
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin VB.ComboBox CboTipo 
            Height          =   315
            ItemData        =   "Proveedor.frx":36C9
            Left            =   1200
            List            =   "Proveedor.frx":36D9
            TabIndex        =   23
            Top             =   3480
            Width           =   3255
         End
         Begin VB.TextBox TxtFax 
            Height          =   375
            Left            =   1200
            TabIndex        =   7
            Top             =   3000
            Width           =   3255
         End
         Begin VB.TextBox TxtTelefono2 
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox TxtDireccion 
            Height          =   555
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1440
            Width           =   6375
         End
         Begin VB.TextBox TxtRazonSocial 
            Height          =   375
            Left            =   1200
            TabIndex        =   4
            Top             =   960
            Width           =   6375
         End
         Begin VB.TextBox TxtRif 
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox TxtTelefono1 
            Height          =   375
            Left            =   1200
            TabIndex        =   2
            Top             =   2040
            Width           =   3255
         End
         Begin ChamaleonButton.ChameleonBtn BtnListadoProveedores 
            Height          =   375
            Left            =   3360
            TabIndex        =   24
            ToolTipText     =   "Eliminar"
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Listado de Proveedores"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":3722
            PICN            =   "Proveedor.frx":373E
            PICH            =   "Proveedor.frx":39D6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnModificar 
            Height          =   375
            Left            =   5760
            TabIndex        =   29
            ToolTipText     =   "Modificar"
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Modificar Reg."
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Proveedor.frx":3DEF
            PICN            =   "Proveedor.frx":3E0B
            PICH            =   "Proveedor.frx":40AF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatus:"
            Height          =   195
            Left            =   5280
            TabIndex        =   32
            Top             =   2220
            Width           =   570
         End
         Begin VB.Label NroReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro de Registro: 0 / 0"
            Height          =   195
            Left            =   6000
            TabIndex        =   28
            Top             =   3120
            Width           =   1545
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   3090
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 2:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2610
            Width           =   810
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "RIF.:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   540
            Width           =   735
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1620
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   2130
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   3540
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "FrmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BD74 As New ADODB.Recordset
Public BD75 As New ADODB.Recordset
Dim BD76 As New ADODB.Recordset
Dim Nuevo
Dim Cambio
Dim IdProv As Integer
Dim RsTemp As New ADODB.Recordset

Private Sub BtnAgregar_Click()
'command1
NroReg.Caption = "Nro de Registro: Nuevo Registro"
Blanqueo
Nuevo = 1
Frame2.BackColor = &HE0E0E0
Frame2.Enabled = True

Habilitar_Campos True

BtnGuardar.Enabled = True
BtnModificar.Enabled = False
BtnAgregar.Enabled = False
BtnBorrar.Enabled = False
BtnSiguiente.Enabled = False
BtnAnterior.Enabled = False
Frame4.Enabled = False
BtnListadoProveedores.Enabled = False
Me.Caption = "Proveedores"
End Sub

Sub Habilitar_Campos(Condicion As Boolean)

If Condicion Then
    TxtRif.Locked = False
    TxtRazonSocial.Locked = False
    TxtDireccion.Locked = False
    TxtTelefono1.Locked = False
    TxtTelefono2.Locked = False
    TxtFax.Locked = False
    CboTipo.Locked = False
Else
    TxtRif.Locked = True
    TxtRazonSocial.Locked = True
    TxtDireccion.Locked = True
    TxtTelefono1.Locked = True
    TxtTelefono2.Locked = True
    TxtFax.Locked = True
    CboTipo.Locked = True
End If
End Sub

Private Sub BtnBorrar_Click()
On Error Resume Next
Dim Rsp

If IdProv < 1 Then
    MsgBox "Debe elegir un proveedor para realizar esta operación!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

Rsp = MsgBox("Se procedera a elimnar el registro actual cuyo Rif es " & Trim(TxtRif.Text) & Chr(13) & " Desea Continuar?", vbQuestion + vbYesNo, "confirmación")
If Rsp = vbNo Then Exit Sub

CSql = "DELETE FROM Proveedores WHERE IdProveedor=" & IdProv
Set RsTemp = CrearRS(CSql)

MsgBox "El registro ha sido Eliminado Exitosamente!", vbInformation + vbOKOnly, "Operación Exitosa!"

BtnDesHacer_Click

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Sub Blanqueo()

TxtRif.Text = ""
TxtRazonSocial.Text = ""
TxtDireccion.Text = ""
TxtTelefono1.Text = ""
TxtTelefono2.Text = ""
TxtFax.Text = ""
CboTipo.Text = ""
CboStatus.ListIndex = -1

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub BtnAnterior_Click()
If BD74.RecordCount <> 0 Then
    BD74.MovePrevious
    If BD74.BOF Then BD74.MoveLast
    Call CargaProve
End If
End Sub

Private Sub BtnBuscar_Click()
Blanqueo
If Trim(TxtBuscar.Text) = "" Or UCase(TxtBuscar.Text) = UCase("Busqueda") Then
    'f = "Buscar"
    CSql = "select * from Proveedores"
Else
    CSql = "select * from Proveedores where Nombre='" & Trim(TxtBuscar.Text) & "' or RifProveedor ='" & Trim(TxtBuscar.Text) & "'"
End If

Set BD74 = CrearRS(CSql)
CargaProve

BtnGuardar.Enabled = True
BtnModificar.Enabled = False
BtnAgregar.Enabled = True
BtnBorrar.Enabled = True

End Sub

Private Sub BtnDesHacer_Click()
Frame2.BackColor = &HEAEFEF
Check1.BackColor = &HEAEFEF
BtnAgregar.Enabled = True
BtnBorrar.Enabled = True
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
Frame4.Enabled = True
BtnListadoProveedores.Enabled = True
Nuevo = 0

Call Ref
Call CargaProve

Habilitar_Campos False

End Sub

Private Sub BtnGuardar_Click()
'command2

If Trim(TxtRif.Text) = "" Then
    MsgBox "Ingrese el RIF del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf Trim(TxtRazonSocial.Text) = "" Then
    MsgBox "Ingrese la Razón Social del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf Trim(TxtDireccion.Text) = "" Then
    MsgBox "Ingrese la Dirección del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf Trim(TxtTelefono1.Text) = "" Then
    MsgBox "Ingrese el Teléfono 1 del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf Trim(TxtTelefono2.Text) = "" Then
    MsgBox "Ingrese el Teléfono 2 del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf Trim(TxtFax.Text) = "" Then
    MsgBox "Ingrese el Número FAX del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf CboTipo.ListIndex < 0 Then
    MsgBox "Ingrese el Tipo del proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
ElseIf CboTipo.ListIndex = -1 Then
    MsgBox "Seleccione el Status del Proveedor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
End If

Select Case Nuevo
    Case Is = 1
    
        CSql = "Select MAX(IdProveedor)+1 As NuevoId From Proveedores"
        Set BD75 = CrearRS(CSql)
        
        If BD75.RecordCount <> 0 Then
            If Not IsNull(BD75.Fields("NuevoId").Value) Then
                C = BD75.Fields("NuevoId").Value
            Else
                C = "1"
            End If
        Else
            C = "1"
        End If
        
        CSql = "Select * From Proveedores"
        Set BD75 = CrearRS(CSql)
        
        BD75.AddNew
        BD75.Fields("IdProveedor").Value = C
        BD75.Fields("Nombre").Value = TxtRazonSocial.Text
        BD75.Fields("RifProveedor").Value = TxtRif.Text
        BD75.Fields("Direccion").Value = TxtDireccion.Text
        BD75.Fields("TelefProveedor1").Value = TxtTelefono1.Text
        BD75.Fields("TelefProveedor2").Value = Val(TxtTelefono2.Text)
        BD75.Fields("FaxProveedor").Value = Val(TxtFax.Text)
        BD75.Fields("TipoProveedor").Value = CboTipo.Text
        BD75.Fields("IdUsuario").Value = IdUser
        BD75.Fields("FechaProveedor").Value = DateTime.Date
        BD75.Fields("Retencion").Value = Check1.Value
               
        BD75.Fields("Status").Value = CboStatus.ItemData(CboStatus.ListIndex)
        
        BD75.Update
        
        Msg = "Registro Agregado Satisfactoriamente!!!"
        MsgBox Msg, vbOKOnly + vbInformation, "Guardado Satisfactorio"
        Call Blanqueo
        Nuevo = 0
        Call Form_Load
    Case Is = 0
    
    If Cambio = 1 Then
                  
'        CSql = "Select MAX(IdProveedor)+1 As NuevoId From Proveedores where Idproveedor = '" & IdProv & "'"
'        Set BD75 = CrearRS(CSql)
'
'        If BD75.RecordCount <> 0 Then
'            If Not IsNull(BD75.Fields("NuevoId")) Then
'                C = BD75.Fields(NuevoId)
'            Else
'                C = "1"
'            End If
'        Else
'            C = "1"
'        End If
        
        CSql = "Select * From Proveedores where IdProveedor = '" & IdProv & "'"
        Set BD75 = CrearRS(CSql)
        
       ' BD75.Fields("IdProveedor").Value = IdProv
        BD75.Fields("Nombre").Value = TxtRazonSocial.Text
        BD75.Fields("RifProveedor").Value = TxtRif.Text
        BD75.Fields("Direccion").Value = TxtDireccion.Text
        BD75.Fields("TelefProveedor1").Value = TxtTelefono1.Text
        BD75.Fields("TelefProveedor2").Value = TxtTelefono2.Text
        BD75.Fields("FaxProveedor").Value = TxtFax.Text
        BD75.Fields("TipoProveedor").Value = CboTipo.Text
        BD75.Fields("IdUsuario").Value = IdUser
        BD75.Fields("FechaProveedor").Value = Format(Now, "dd/MM/yyyy")
        BD75.Fields("Retencion").Value = Check1.Value
        BD75.Fields("Status").Value = CboStatus.ItemData(CboStatus.ListIndex)
        BD75.Update
        
        Msg = "Registro Actuliazado Satisfactoriamente"
        MsgBox Msg, vbOKOnly + vbInformation, "Guardado Satisfactorio"
        Nuevo = 0
        Call Form_Load
        End If

End Select

Call Ref
Call CargaProve

BtnDesHacer_Click

Exit Sub

End Sub

Private Sub BtnListadoProveedores_Click()
Tipo = "LstProveedor"
FrmListadoProveedor.Show
End Sub

Private Sub BtnModificar_Click()
NroReg.Caption = "Nro de Registro: MODIFICANDO"

Nuevo = 0
Frame2.BackColor = &HE0E0E0
Check1.BackColor = &HE0E0E0
Frame2.Enabled = True

Habilitar_Campos True

BtnGuardar.Enabled = True
BtnModificar.Enabled = False
BtnAgregar.Enabled = False
BtnBorrar.Enabled = False
BtnSiguiente.Enabled = False
BtnAnterior.Enabled = False
Frame4.Enabled = False
BtnListadoProveedores.Enabled = False
Me.Caption = "Proveedores"
End Sub

Private Sub BtnSiguiente_Click()
If BD74.RecordCount <> 0 Then
    BD74.MoveNext
    If BD74.EOF Then BD74.MoveFirst
    Call CargaProve
End If
End Sub

Sub CargaProve()
BtnGuardar.Enabled = False
If BD74.RecordCount > 0 Then
    Nuevo = 0
    If Trim(BD74.Fields("Nombre").Value) <> "" Then TxtRazonSocial.Text = BD74.Fields("Nombre").Value
    If Trim(BD74.Fields("Direccion").Value) <> "" Then TxtDireccion.Text = BD74.Fields("Direccion").Value
    If Trim(BD74.Fields("RifProveedor").Value) <> "" Then TxtRif.Text = BD74.Fields("RifProveedor").Value
    If Trim(BD74.Fields("TelefProveedor1").Value) <> "" Then TxtTelefono1.Text = BD74.Fields("TelefProveedor1").Value
    If Trim(BD74.Fields("TelefProveedor2").Value) <> "" Then TxtTelefono2.Text = BD74.Fields("TelefProveedor2").Value
    If Trim(BD74.Fields("FaxProveedor").Value) <> "" Then TxtFax.Text = BD74.Fields("FaxProveedor").Value
    If Trim(BD74.Fields("TipoProveedor").Value) <> "" Then CboTipo.Text = BD74.Fields("TipoProveedor").Value
    
    If BD74.Fields("Retencion").Value = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    
    If BD74.Fields("Status").Value = 1 Then
        CboStatus.ListIndex = 1
        'CboStatus.Text = "Activo"
    Else
        CboStatus.ListIndex = 0
       ' CboStatus.Text = "Inactivo"
    End If
    
'    For i = 1 To CboStatus.ListCount - 1
'        If IsNull(BD74.Fields("Status").Value) Then
'            CboStatus.ListIndex = -1
'            Exit For
'        ElseIf CboStatus.ItemData(i) = Trim(BD74.Fields("Status").Value) Then
'            CboStatus.ListIndex = i
'            Exit For
'        End If
'
'    Next i
    
   
    IdProv = BD74.Fields("idproveedor").Value
              
    Me.Caption = "Proveedores - Id: " & IdProv
    NroReg.Caption = "Nro de Registro: " & BD74.AbsolutePosition & " / " & BD74.RecordCount
    BtnModificar.Enabled = True
    BtnBorrar.Enabled = True
Else
    IdProv = 0
    Nuevo = 1
    NroReg.Caption = "Nro de Registro: 0 / 0"
    BtnModificar.Enabled = False
    BtnBorrar.Enabled = False
End If


End Sub

Private Sub Form_Load()
Centrar Me

CboStatus.AddItem "Suspendido"
CboStatus.ItemData(CboStatus.NewIndex) = 0
CboStatus.AddItem "Activo"
CboStatus.ItemData(CboStatus.NewIndex) = 1

Cambio = 0
Nuevo = 0
Call Ref
Call CargaProve
Frame2.BackColor = &HEAEFEF
End Sub
Sub Ref()
CSql = "SELECT * FROM Proveedores ORDER BY IdProveedor"
If BD74.State = adStateOpen Then BD74.Close
Set BD74 = CrearRS(CSql)
End Sub



Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub


Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890-", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtDireccion_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(TxtDireccion.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(TxtDireccion.Text)
    pru = LCase(Mid(TxtDireccion.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

TxtDireccion.Text = StrText
TxtDireccion.SelStart = Len(TxtDireccion.Text)

End Sub

Private Sub TxtFax_Change()
Cambio = 1
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii <> 8 Then
              
    If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtRif_Click()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(TxtRif.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(TxtRif.Text)
    pru = LCase(Mid(TxtRif.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

TxtRif.Text = StrText
TxtRif.SelStart = Len(TxtRif.Text)

End Sub

Private Sub CboTipo_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(CboTipo.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(CboTipo.Text)
    pru = LCase(Mid(CboTipo.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

CboTipo.Text = StrText
CboTipo.SelStart = Len(CboTipo.Text)

End Sub

Private Sub TxtRazonSocial_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(TxtRazonSocial.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(TxtRazonSocial.Text)
    pru = LCase(Mid(TxtRazonSocial.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

TxtRazonSocial.Text = StrText
TxtRazonSocial.SelStart = Len(TxtRazonSocial.Text)
End Sub

Private Sub TxtTelefono1_Change()
Cambio = 1
End Sub

Private Sub TxtTelefono1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii <> 8 Then
              
    If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtTelefono2_Change()
Cambio = 1
End Sub

Private Sub TxtTelefono2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii <> 8 Then
              
    If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

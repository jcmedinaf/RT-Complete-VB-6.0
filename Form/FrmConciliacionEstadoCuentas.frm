VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConciliacionEstadoCuentas 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de cuentas"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "FrmConciliacionEstadoCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   8655
      Begin VB.OptionButton OptEgreso 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Egreso"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   2820
         Width           =   1095
      End
      Begin VB.OptionButton OptIngreso 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Ingreso"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   2820
         Width           =   1095
      End
      Begin VB.ComboBox CboTipo 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   2310
         Width           =   2535
      End
      Begin VB.TextBox TxtBancos 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox TxtDocumento 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   735
         Left            =   1080
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1440
         Width           =   7455
      End
      Begin VB.TextBox TxtIdBanco 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPickerFecha 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52559873
         CurrentDate     =   40232
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         ToolTipText     =   "Buscar Bancos"
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Buscar"
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
         MICON           =   "FrmConciliacionEstadoCuentas.frx":1002
         PICN            =   "FrmConciliacionEstadoCuentas.frx":101E
         PICH            =   "FrmConciliacionEstadoCuentas.frx":1283
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2850
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Movimiento:"
         Height          =   195
         Left            =   3720
         TabIndex        =   20
         Top             =   2370
         Width           =   1440
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   570
         Width           =   870
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1050
         Width           =   510
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2370
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   570
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   8655
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   7560
         TabIndex        =   11
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
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
         MICON           =   "FrmConciliacionEstadoCuentas.frx":1515
         PICN            =   "FrmConciliacionEstadoCuentas.frx":1531
         PICH            =   "FrmConciliacionEstadoCuentas.frx":16FA
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
         Height          =   495
         Left            =   6360
         TabIndex        =   10
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         MICON           =   "FrmConciliacionEstadoCuentas.frx":192F
         PICN            =   "FrmConciliacionEstadoCuentas.frx":194B
         PICH            =   "FrmConciliacionEstadoCuentas.frx":1C2D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Guardar"
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
         MICON           =   "FrmConciliacionEstadoCuentas.frx":1E7E
         PICN            =   "FrmConciliacionEstadoCuentas.frx":1E9A
         PICH            =   "FrmConciliacionEstadoCuentas.frx":2129
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
End
Attribute VB_Name = "FrmConciliacionEstadoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMaxId As New ADODB.Recordset
Dim RsGuardar As New ADODB.Recordset

Sub Blanqueo()
TxtDocumento.Text = ""
TxtIdBanco.Text = ""
TxtDetalle.Text = ""
TxtMonto.Text = ""
CboTipo.ListIndex = -1
OptIngreso.Value = False
OptEgreso.Value = False
DTPickerFecha.Value = DateTime.Date
End Sub

Private Sub BtnBuscar_Click()
Ban = 7
FrmListadoBancos.Show vbModal
End Sub

Private Sub BtnBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtDetalle.SetFocus
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
End Sub

Private Sub BtnGuardarActualizar_Click()
'***************************************
'Validacion de campos

If TxtDocumento.Text = "" Then
    MsgBox "Debe de ingresar el número del Movimiento!", vbCritical + vbOKOnly, "Error"
    TxtDocumento.SetFocus
    Exit Sub
End If

If TxtIdBanco.Text = "" Then
    MsgBox "Debe de Seleccionar el banco!", vbCritical + vbOKOnly, "Error"
    BtnBuscar.SetFocus
    Exit Sub
End If

If TxtDetalle.Text = "" Then
    MsgBox "Debe de ingresar el detalle del Movimiento!", vbCritical + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
End If

If TxtMonto.Text = "" Then
    MsgBox "Debe de ingresar el Monto del movimiento!", vbCritical + vbOKOnly, "Error"
    TxtMonto.SetFocus
    Exit Sub
End If

If CboTipo.ListIndex = -1 Then
    MsgBox "Debe de Selecionar el Tipo de Movimiento!", vbCritical + vbOKOnly, "Error"
    CboTipo.SetFocus
    Exit Sub
End If

If OptIngreso.Value = False And OptEgreso.Value = False Then
    MsgBox "Debe de Selecionar el Tipo de Transacción!", vbCritical + vbOKOnly, "Error"
    OptIngreso.SetFocus
    Exit Sub
End If

'***************************************
'Guardados de datos

CSql = "Select max(IdMovEstado) +1 as MaxId From EstadoDeCuenta"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    NuevoId = RsMaxId.Fields("MaxId").Value
Else
    NuevoId = "1"
End If


CSql = "Select * From EstadoDeCuenta"
Set RsGuardar = CrearRS(CSql)


RsGuardar.AddNew
RsGuardar.Fields("IdMovEstado").Value = NuevoId
RsGuardar.Fields("IdCajaBanco").Value = TxtIdBanco.Text
If OptIngreso.Value = True Then
    RsGuardar.Fields("Ingr_Egr").Value = 1
ElseIf OptEgreso.Value = True Then
    RsGuardar.Fields("Ingr_Egr").Value = 2
End If
RsGuardar.Fields("n_comprobante").Value = TxtDocumento.Text
RsGuardar.Fields("Monto_Mov").Value = TxtMonto.Text
RsGuardar.Fields("Tipo_Mov").Value = CboTipo.ListIndex
RsGuardar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
RsGuardar.Fields("Conciliado").Value = 0
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Fields("Detalle").Value = TxtDetalle.Text

RsGuardar.Update

FrmConcilicacionBancaria.Movi
'FrmConcilicacionBancaria.calcular
FrmConcilicacionBancaria.Conciliacion
FrmConcilicacionBancaria.TotalConciliado
FrmConcilicacionBancaria.TotalNoConciliado
Unload Me

End Sub

Private Sub CboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OptIngreso.SetFocus
End If
End Sub

Private Sub DTPickerFecha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    BtnBuscar.SetFocus
End If

End Sub

Private Sub Form_Load()
DTPickerFecha.Value = DateTime.Date


CboTipo.AddItem "Efectivo"
CboTipo.ItemData(CboTipo.NewIndex) = 1
CboTipo.AddItem "Cheque"
CboTipo.ItemData(CboTipo.NewIndex) = 2
CboTipo.AddItem "Depósito"
CboTipo.ItemData(CboTipo.NewIndex) = 3
CboTipo.AddItem "Transferencia"
CboTipo.ItemData(CboTipo.NewIndex) = 4
CboTipo.AddItem "Tarjeta de Crédito"
CboTipo.ItemData(CboTipo.NewIndex) = 5
CboTipo.AddItem "Tarjeta de Débito"
CboTipo.ItemData(CboTipo.NewIndex) = 6
CboTipo.AddItem "Comisión Bancaria"
CboTipo.ItemData(CboTipo.NewIndex) = 7

End Sub

Private Sub OptEgreso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnGuardarActualizar.SetFocus
End If
End Sub

Private Sub OptIngreso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OptEgreso.SetFocus
End If
End Sub

Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMonto.SetFocus
End If
End Sub

'///////////////////////////////////Valido TextBox: TxtDocumento//////////////////////////////
Private Sub TxtDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPickerFecha.SetFocus
    Else
        If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

'///////////////////////////////////Valido TextBox: TxtMonto//////////////////////////////
Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboTipo.SetFocus
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

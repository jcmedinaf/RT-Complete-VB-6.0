VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTransferenciaFondos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Fondos"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "FrmTransferenciaFondos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   8655
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7560
         TabIndex        =   14
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
         MICON           =   "FrmTransferenciaFondos.frx":1002
         PICN            =   "FrmTransferenciaFondos.frx":101E
         PICH            =   "FrmTransferenciaFondos.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "FrmTransferenciaFondos.frx":141C
         PICN            =   "FrmTransferenciaFondos.frx":1438
         PICH            =   "FrmTransferenciaFondos.frx":15C5
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
         Left            =   6360
         TabIndex        =   16
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
         MICON           =   "FrmTransferenciaFondos.frx":17FA
         PICN            =   "FrmTransferenciaFondos.frx":1816
         PICH            =   "FrmTransferenciaFondos.frx":1AF8
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
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         MICON           =   "FrmTransferenciaFondos.frx":1D49
         PICN            =   "FrmTransferenciaFondos.frx":1D65
         PICH            =   "FrmTransferenciaFondos.frx":1FF4
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.TextBox TxtBancosDestino 
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox TxtIdBancoDestino 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   1200
      End
      Begin VB.TextBox TxtBancos 
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtNoTransferencia 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   735
         Left            =   1080
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1680
         Width           =   7455
      End
      Begin VB.TextBox TxtIdBanco 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPickerFecha 
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54919169
         CurrentDate     =   40232
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   7200
         TabIndex        =   7
         ToolTipText     =   "Buscar"
         Top             =   720
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
         MICON           =   "FrmTransferenciaFondos.frx":2435
         PICN            =   "FrmTransferenciaFondos.frx":2451
         PICH            =   "FrmTransferenciaFondos.frx":26B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar2 
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         ToolTipText     =   "Buscar"
         Top             =   1200
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
         MICON           =   "FrmTransferenciaFondos.frx":2948
         PICN            =   "FrmTransferenciaFondos.frx":2964
         PICH            =   "FrmTransferenciaFondos.frx":2BC9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Destino:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Transferencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Origen:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   5760
         TabIndex        =   10
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3240
         TabIndex        =   8
         Top             =   330
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmTransferenciaFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
Blanqueo
Frame2.Enabled = True
BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
TxtNoTransferencia.SetFocus
End Sub

Private Sub BtnBuscar_Click()
Ban = 8
FrmListadoBancos.Show vbModal
End Sub

Private Sub BtnBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar2.SetFocus
End If
End Sub

Private Sub BtnBuscar2_Click()
Ban = 9
FrmListadoBancos.Show vbModal
End Sub

Private Sub BtnBuscar2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtDetalle.SetFocus
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub
Sub Blanqueo()
TxtDetalle.Text = ""
TxtNoTransferencia.Text = ""
TxtIdBanco.Text = ""
TxtBancos.Text = ""
TxtIdBancoDestino.Text = ""
TxtBancosDestino.Text = ""
TxtMonto.Text = ""
DTPickerFecha.Value = DateTime.Date
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo

Frame2.Enabled = fasle
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False

End Sub

Private Sub BtnGuardarActualizar_Click()
'--------------------------------------
'Validacion de campos
'--------------------------------------

If TxtNoTransferencia.Text = "" Then
    MsgBox "Debe de seleccionar el numero de la transferencia!", vbCritical + vbOKOnly, "Error"
    TxtNoTransferencia.SetFocus
    Exit Sub
End If

If TxtIdBanco.Text = "" Then
    MsgBox "Debe de seleccionar el Banco de origen de la transferecia!", vbCritical + vbOKOnly, "Error"
    BtnBuscar.SetFocus
    Exit Sub
End If

If TxtMonto.Text = "" Then
    MsgBox "Debe de ingresar el valor del monto!", vbCritical + vbOKOnly, "Error"
    TxtMonto.SetFocus
    Exit Sub
End If

If TxtIdBancoDestino.Text = "" Then
    MsgBox "Debe de seleccionar el Banco destino de la transferencia!", vbCritical + vbOKOnly, "Error"
    BtnBuscar2.SetFocus
    Exit Sub
End If

If TxtDetalle.Text = "" Then
    MsgBox "Debe de Ingresar el detalle de la transferencia", vbCritical + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
End If


'###################################################################
'Verifica si hay saldo en la cuenta para poder realizar el cheque

CSql = "Select * From Movi_BanCaja Where IdCajaBanco='" & TxtIdBanco.Text & "' And Ingr_Egr='2'"
Set RsVerificarSaldo = CrearRS(CSql)

If RsVerificarSaldo.RecordCount = 0 Then
    MsgBox "Esta Cuenta No Posee Saldo para realizar la transferencia de fondos!!!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If


'###################################################################
'Agregar deposito

CSql = "Select Max(IdMovCajaBanco)+1 as MaxId From Movi_BanCaja"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdMax = RsMaxId.Fields("MaxId").Value
Else
    IdMax = 1
End If


CSql = "Select * From Movi_BanCaja"
Set RsGuardar = CrearRS(CSql)

RsGuardar.AddNew
RsGuardar.Fields("IdMovCajaBanco").Value = IdMax
RsGuardar.Fields("IdCajaBanco").Value = TxtIdBanco.Text
RsGuardar.Fields("Ingr_Egr").Value = 2
RsGuardar.Fields("n_comprobante").Value = TxtNoTransferencia.Text
RsGuardar.Fields("Monto_Mov").Value = TxtMonto.Text
RsGuardar.Fields("Detalle").Value = TxtDetalle.Text
RsGuardar.Fields("Tipo_Mov").Value = 6
RsGuardar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
RsGuardar.Fields("Conciliado").Value = 0
RsGuardar.Fields("Anulado").Value = 0
RsGuardar.Fields("NoEndosable").Value = 0
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Update

FrmLibroBancos.Movi

CSql = "Select Max(IdMovCajaBanco)+1 as MaxId From Movi_BanCaja"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdMax = RsMaxId.Fields("MaxId").Value
Else
    IdMax = 1
End If


CSql = "Select * From Movi_BanCaja"
Set RsGuardar = CrearRS(CSql)

RsGuardar.AddNew
RsGuardar.Fields("IdMovCajaBanco").Value = IdMax
RsGuardar.Fields("IdCajaBanco").Value = TxtIdBancoDestino.Text
RsGuardar.Fields("Ingr_Egr").Value = 1
RsGuardar.Fields("n_comprobante").Value = TxtNoTransferencia.Text
RsGuardar.Fields("Monto_Mov").Value = TxtMonto.Text
RsGuardar.Fields("Detalle").Value = TxtDetalle.Text
RsGuardar.Fields("Tipo_Mov").Value = 6
RsGuardar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
RsGuardar.Fields("Conciliado").Value = 0
RsGuardar.Fields("Anulado").Value = 0
RsGuardar.Fields("NoEndosable").Value = 0
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Update

MsgBox "Registro Almacenado con Exito!", vbInformation + vbOKOnly, "Movimiento Guardado"

FrmLibroBancos.Movi

Frame2.Enabled = False
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub DTPickerFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMonto.SetFocus
End If

End Sub

Private Sub Form_Load()
DTPickerFecha.Value = DateTime.Time
Frame2.Enabled = False
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
End Sub

'///////////////////////////////////Valido TextBox: txtmonto//////////////////////////////
Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BtnBuscar.SetFocus
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub



'///////////////////////////////////Valido TextBox: TxtNoTransferencia//////////////////////////////
Private Sub TxtNoTransferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPickerFecha.SetFocus
    Else
        If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTransaccionDepositos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depósito"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "FrmTransaccionDeposito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   8655
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7560
         TabIndex        =   9
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
         MICON           =   "FrmTransaccionDeposito.frx":1002
         PICN            =   "FrmTransaccionDeposito.frx":101E
         PICH            =   "FrmTransaccionDeposito.frx":11E7
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
         TabIndex        =   6
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
         MICON           =   "FrmTransaccionDeposito.frx":141C
         PICN            =   "FrmTransaccionDeposito.frx":1438
         PICH            =   "FrmTransaccionDeposito.frx":15C5
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
         TabIndex        =   8
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
         MICON           =   "FrmTransaccionDeposito.frx":17FA
         PICN            =   "FrmTransaccionDeposito.frx":1816
         PICH            =   "FrmTransaccionDeposito.frx":1AF8
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
         TabIndex        =   7
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
         MICON           =   "FrmTransaccionDeposito.frx":1D49
         PICN            =   "FrmTransaccionDeposito.frx":1D65
         PICH            =   "FrmTransaccionDeposito.frx":1FF4
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox TxtFactura 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6840
         TabIndex        =   26
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ComboBox CboPaciente 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   2670
         Width           =   4695
      End
      Begin VB.ComboBox CboCliente 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   2190
         Width           =   4695
      End
      Begin VB.TextBox TxtBeneficiario 
         Height          =   375
         Left            =   1080
         MaxLength       =   1000
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1680
         Width           =   7455
      End
      Begin VB.TextBox TxtDepositante 
         Height          =   375
         Left            =   1080
         MaxLength       =   1000
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   1200
         Width           =   7455
      End
      Begin VB.TextBox TxtIdBanco 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPickerFecha 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50987009
         CurrentDate     =   40232
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   735
         Left            =   1080
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3120
         Width           =   7455
      End
      Begin VB.TextBox TxtDocumento 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6840
         TabIndex        =   3
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox TxtBancos 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   4815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         ToolTipText     =   "Buscar Bancos"
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
         MICON           =   "FrmTransaccionDeposito.frx":2435
         PICN            =   "FrmTransaccionDeposito.frx":2451
         PICH            =   "FrmTransaccionDeposito.frx":26B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura"
         Height          =   195
         Left            =   5880
         TabIndex        =   27
         Top             =   2250
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paciente:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2730
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2250
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficiario:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depositante:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1290
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3390
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   6240
         TabIndex        =   13
         Top             =   2730
         Width           =   495
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   870
      End
   End
End
Attribute VB_Name = "FrmTransaccionDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGuardar As New ADODB.Recordset
Dim RsActualizar As New ADODB.Recordset
Dim RsMaxId As New ADODB.Recordset
Private Sub BtnAgregar_Click()
Editar = 0
Blanqueo
Frame2.Enabled = True
TxtDocumento.SetFocus
BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
End Sub

Private Sub BtnBuscar_Click()
Ban = 2
FrmListadoBancos.Show vbModal
End Sub

Private Sub BtnCerrar_Click()
Unload Me
Editar = 0
End Sub

Sub Blanqueo()
TxtDetalle.Text = ""
TxtDocumento.Text = ""
TxtBancos.Text = ""
TxtIdBanco.Text = ""
TxtDepositante.Text = ""
TxtBeneficiario.Text = ""
CboPaciente.ListIndex = -1
CboCliente.ListIndex = -1
TxtMonto.Text = ""
DTPickerFecha.Value = DateTime.Date
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error GoTo WrtError
'###################################################################
'Validacion de campos

If TxtDocumento.Text = "" Then
    MsgBox "Debe de ingresar el numero de documento", vbCritical + vbOKOnly, "Error"
    TxtDocumento.SetFocus
    Exit Sub

ElseIf TxtMonto.Text = "" Then
    MsgBox "Debe de ingresar el monto", vbCritical + vbOKOnly, "Error"
    TxtMonto.SetFocus
    Exit Sub

ElseIf TxtIdBanco.Text = "" Then
    MsgBox "Falta seleccionar el Banco", vbCritical + vbOKOnly, "Error"
    TxtIdBanco.SetFocus
    Exit Sub
    
ElseIf TxtBeneficiario.Text = "" Then
    MsgBox "Debe de ingresar el Beneficiario", vbCritical + vbOKOnly, "Error"
    TxtBeneficioario.SetFocus
    Exit Sub
    
ElseIf TxtDepositante.Text = "" Then
    MsgBox "Debe de ingresar el depositante", vbCritical + vbOKOnly, "Error"
    TxtDepositante.SetFocus
    Exit Sub
    
ElseIf TxtDetalle.Text = "" Then
    MsgBox "Debe de ingresar un breve detalle del movimiento", vbCritical + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
    
'ElseIf CboCliente.ListIndex = -1 Then
'    MsgBox "Debe de seleccionar un cliente", vbCritical + vbOKOnly, "Error"
'    CboCliente.SetFocus
'    Exit Sub
'
'ElseIf CboPaciente.ListIndex = -1 Then
'    MsgBox "Debe de seleccionar a un paciente", vbCritical + vbOKOnly, "Error"
'    CboPaciente.SetFocus
'    Exit Sub
    
End If

Select Case Editar

Case Is = 0
    '###################################################################
    'Agregar deposito
    CSql = "Select max(IdMovCajaBanco) + 1 as MaxId From Movi_BanCaja"
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
    RsGuardar.Fields("Ingr_Egr").Value = 1
    RsGuardar.Fields("n_comprobante").Value = TxtDocumento.Text
    RsGuardar.Fields("Monto_Mov").Value = CDbl(TxtMonto.Text)
    RsGuardar.Fields("Beneficiario").Value = TxtBeneficiario.Text
    RsGuardar.Fields("Depositante").Value = TxtDepositante.Text
    RsGuardar.Fields("Detalle").Value = TxtDetalle.Text
    RsGuardar.Fields("Tipo_Mov").Value = 1
    RsGuardar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
    RsGuardar.Fields("Anulado").Value = 0
    RsGuardar.Fields("Conciliado").Value = 0
    RsGuardar.Fields("IdUsuario").Value = IdUser
    RsGuardar.Fields("NoEndosable").Value = 0
    
    If TxtFactura.Text <> "" Then
        RsGuardar.Fields("N_Factura").Value = TxtFactura.Text
    Else
        RsGuardar.Fields("N_Factura").Value = 0
    End If
    
'    RsGuardar.Fields("IdPaciente").Value = CboPaciente.ItemData(CboPaciente.ListIndex)
'    RsGuardar.Fields("IdCliente").Value = CboCliente.ItemData(CboCliente.ListIndex)
    
    If CboPaciente.ItemData(CboPaciente.ListIndex) = -1 Then RsGuardar.Fields("IdPaciente").Value = -1 Else RsGuardar.Fields("IdPaciente").Value = CboPaciente.ItemData(CboPaciente.ListIndex)
    If CboCliente.ItemData(CboCliente.ListIndex) = -1 Then RsGuardar.Fields("IdCliente").Value = -1 Else RsGuardar.Fields("IdCliente").Value = CboCliente.ItemData(CboCliente.ListIndex)
    
    RsGuardar.Update
    
    MsgBox "Registro Almacenado con Exito!", vbInformation + vbOKOnly, "Movimiento Guardado"
    FrmLibroBancos.Grid2
    FrmLibroBancos.Movi

Case Is = 1
    '###################################################################
    'Actualiza deposito
    CSql = "Select * From Movi_BanCaja Where IdMovCajaBanco = '" & IdReg & "'"
    Set RsActualizar = CrearRS(CSql)
    
    RsActualizar.Fields("IdCajaBanco").Value = TxtIdBanco.Text
    RsActualizar.Fields("Ingr_Egr").Value = 1
    RsActualizar.Fields("n_comprobante").Value = TxtDocumento.Text
    RsActualizar.Fields("Monto_Mov").Value = CDbl(TxtMonto.Text)
    RsActualizar.Fields("Beneficiario").Value = TxtBeneficiario.Text
    RsActualizar.Fields("Depositante").Value = TxtDepositante.Text
    RsActualizar.Fields("Detalle").Value = TxtDetalle.Text
    RsActualizar.Fields("Tipo_Mov").Value = 1
    RsActualizar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
    RsActualizar.Fields("Conciliado").Value = 0
    RsActualizar.Fields("Anulado").Value = 0
    RsActualizar.Fields("Conciliado").Value = 0
    RsActualizar.Fields("IdUsuario").Value = IdUser
    RsActualizar.Fields("NoEndosable").Value = 0
    
    If TxtFactura.Text <> "" Then
        RsActualizar.Fields("N_Factura").Value = TxtFactura.Text
    Else
        RsActualizar.Fields("N_Factura").Value = 0
    End If
    
    If CboPaciente.ItemData(CboPaciente.ListIndex) = -1 Then RsActualizar.Fields("IdPaciente").Value = -1 Else RsActualizar.Fields("IdPaciente").Value = CboPaciente.ItemData(CboPaciente.ListIndex)
    If CboCliente.ItemData(CboCliente.ListIndex) = -1 Then RsActualizar.Fields("IdCliente").Value = -1 Else RsActualizar.Fields("IdCliente").Value = CboCliente.ItemData(CboCliente.ListIndex)
    
    RsActualizar.Update
    
    MsgBox "Registro Actualizado con Exito!", vbInformation + vbOKOnly, "Movimiento Actualizado"
    FrmLibroBancos.Grid2
    FrmLibroBancos.Movi
    
End Select

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
Frame2.Enabled = False
Editar = 0
Unload Me

Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
'MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
Editar = 0
Unload Me
End Sub

Private Sub DTPickerFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        TxtMonto.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim RsCboPaciente As New ADODB.Recordset
Dim RsCboCliente As New ADODB.Recordset

'*********************

CSql = "Select * From Paciente Order By IdPaciente"
Set RsCboPaciente = CrearRS(CSql)

Do While Not RsCboPaciente.EOF
    CboPaciente.AddItem Trim(RsCboPaciente.Fields("NombreP").Value) & ", " & Trim(RsCboPaciente.Fields("ApellidoP").Value)
    CboPaciente.ItemData(CboPaciente.NewIndex) = RsCboPaciente.Fields("IdPaciente").Value
    RsCboPaciente.MoveNext
Loop

'*********************

CSql = "Select * From Cliente Order By IdCliente"
Set RsCboCliente = CrearRS(CSql)

Do While Not RsCboCliente.EOF
    CboCliente.AddItem Trim(RsCboCliente.Fields("Razon").Value)
    CboCliente.ItemData(CboCliente.NewIndex) = RsCboCliente.Fields("IdCliente").Value
    RsCboCliente.MoveNext
Loop

'*********************

If Editar = 1 Then
    Frame2.Enabled = True
    BtnGuardarActualizar.Enabled = True


    CSql = "Select * From Movi_BanCaja Where IdMovCajaBanco = '" & IdReg & "'"
    Set RsCargarMovimientos = CrearRS(CSql)
    
    TxtDocumento.Text = RsCargarMovimientos.Fields("n_comprobante").Value
    TxtMonto.Text = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
    TxtIdBanco.Text = RsCargarMovimientos.Fields("IdCajaBanco").Value
    If Not IsNull(RsCargarMovimientos.Fields("Depositante").Value) Then TxtDepositante.Text = RsCargarMovimientos.Fields("Depositante").Value Else TxtDepositante.Text = ""
    If Not IsNull(RsCargarMovimientos.Fields("Beneficiario").Value) Then TxtBeneficiario.Text = RsCargarMovimientos.Fields("Beneficiario").Value Else TxtBeneficiario.Text = ""
    If Not IsNull(RsCargarMovimientos.Fields("Detalle").Value) Then TxtDetalle.Text = RsCargarMovimientos.Fields("Detalle").Value Else TxtDetalle.Text = ""
    DTPickerFecha.Value = RsCargarMovimientos.Fields("Fecha_Transa").Value
          
    If Not IsNull(RsCargarMovimientos.Fields("N_Factura").Value) Then
        If RsCargarMovimientos.Fields("N_Factura").Value <> 0 Then
           TxtFactura.Text = RsCargarMovimientos.Fields("N_Factura").Value
        Else
            TxtFactura.Text = ""
        End If
    Else
        TxtFactura.Text = ""
    End If
        
    For i = 1 To CboPaciente.ListCount - 1
        If IsNull(RsCargarMovimientos.Fields("IdPaciente").Value) Then
            CboPaciente.ListIndex = -1
            Exit For
        ElseIf CboPaciente.ItemData(i) = RsCargarMovimientos.Fields("IdPaciente").Value Then
            CboPaciente.ListIndex = i
            Exit For
        End If
    Next i
    
    For i = 1 To CboCliente.ListCount - 1
        If IsNull(RsCargarMovimientos.Fields("IdCliente").Value) Then
            CboCliente.ListIndex = -1
            Exit For
        ElseIf CboCliente.ItemData(i) = RsCargarMovimientos.Fields("IdCliente").Value Then
            CboCliente.ListIndex = i
            Exit For
        End If
    Next i
       
    CSql = "Select * From CajasBancos Where IdCajaBanco = '" & TxtIdBanco.Text & "'"
    Set RsCargarMovimientos = CrearRS(CSql)
    TxtBancos.Text = RsCargarMovimientos.Fields("Descripcion").Value
    

Else
    DTPickerFecha.Value = DateTime.Date
    BtnGuardarActualizar.Enabled = False
    Frame2.Enabled = False
    CboPaciente.ListIndex = -1
    CboCliente.ListIndex = -1
End If

End Sub

'///////////////////////////////////Valido TextBox: TxtMonto//////////////////////////////
Private Sub TxtDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPickerFecha.SetFocus
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

'///////////////////////////////////Valido TextBox: TxtMonto//////////////////////////////
Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BtnBuscar.SetFocus
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

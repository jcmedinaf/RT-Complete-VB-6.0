VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEditorFacturasCxC 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Facturas Cuentas por Cobrar"
   ClientHeight    =   5940
   ClientLeft      =   5085
   ClientTop       =   4470
   ClientWidth     =   9090
   Icon            =   "FrmEditorFacturasCxC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8895
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Height          =   1815
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   8655
         Begin VB.TextBox TxtOtros 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1080
            TabIndex        =   42
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox TxtRetenciones 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6840
            TabIndex        =   37
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox TxtTimbresFiscales 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1080
            TabIndex        =   36
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtDescuentos 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3960
            TabIndex        =   35
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox TxtTasaImpuesto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6840
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtPorCobrar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtTotalGeneral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtImpuesto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1080
            TabIndex        =   25
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox TxtSubTotal 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3960
            TabIndex        =   24
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtDiasCredito 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Otros:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7680
            TabIndex        =   41
            Top             =   300
            Width           =   150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retenciones:"
            Height          =   195
            Left            =   5640
            TabIndex        =   40
            Top             =   660
            Width           =   945
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Timbres Fisc."
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1020
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descuentos:"
            Height          =   195
            Left            =   2880
            TabIndex        =   38
            Top             =   660
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa Impuesto:"
            Height          =   195
            Left            =   5640
            TabIndex        =   34
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Por Cobrar:"
            Height          =   195
            Left            =   5640
            TabIndex        =   32
            Top             =   1013
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total General:"
            Height          =   195
            Left            =   2880
            TabIndex        =   31
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   660
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal:"
            Height          =   195
            Left            =   2880
            TabIndex        =   29
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Días Crédito:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   300
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   8655
         Begin VB.TextBox TxtCedulaPaciente 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtNombre 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox TxtApellido 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cedula:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   420
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   4440
            TabIndex        =   19
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   780
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Cliente"
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   8655
         Begin VB.TextBox TxtRif 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtRazonSocial 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   600
            Width           =   7455
         End
         Begin VB.TextBox TxtCodigo 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif:"
            Height          =   195
            Left            =   2880
            TabIndex        =   14
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   660
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.TextBox TxtNoFactura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DtpFechaEmision 
         Height          =   330
         Left            =   7200
         TabIndex        =   5
         Top             =   225
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   51314689
         CurrentDate     =   40141
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Factura:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emisión:"
         Height          =   195
         Left            =   5760
         TabIndex        =   6
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8895
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7800
         TabIndex        =   1
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
         MICON           =   "FrmEditorFacturasCxC.frx":1002
         PICN            =   "FrmEditorFacturasCxC.frx":101E
         PICH            =   "FrmEditorFacturasCxC.frx":11E7
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
         Left            =   120
         TabIndex        =   2
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
         MICON           =   "FrmEditorFacturasCxC.frx":141C
         PICN            =   "FrmEditorFacturasCxC.frx":1438
         PICH            =   "FrmEditorFacturasCxC.frx":16C7
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
Attribute VB_Name = "FrmEditorFacturasCxC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim RsUpdate As New ADODB.Recordset

' Actualiza datos de la tabla C_COBRAR
CSql = "Select * From C_Cobrar Where (Forma_Pago<>0 or Forma_Pago<> -1) And IdCliente='" & FrmCuentasPorCobrar.IdCliente & "' And IdPaciente='" & FrmCuentasPorCobrar.IdPaciente & "' And N_Factura='" & FrmCuentasPorCobrar.NoFact & "'"
Set RsUpdate = CrearRS(CSql)
RsUpdate.Fields("N_Factura").Value = TxtNoFactura.Text
RsUpdate.Fields("Fecha").Value = Format(DtpFechaEmision.Value, "DD/MM/YYYY")
RsUpdate.Fields("PorCobrar").Value = CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text))
RsUpdate.Fields("TasaImpuesto").Value = CDbl(TxtTasaImpuesto.Text)
RsUpdate.Fields("Impuesto").Value = CDbl(TxtImpuesto.Text)
RsUpdate.Fields("SubTotal").Value = CDbl(TxtSubTotal.Text)
RsUpdate.Fields("Monto").Value = CDbl(TxtTotalGeneral.Text)

RsUpdate.Fields("Descuentos").Value = CDbl(TxtDescuentos.Text)
RsUpdate.Fields("Retenciones").Value = CDbl(TxtRetenciones.Text)
RsUpdate.Fields("TimbresFiscales").Value = CDbl(TxtTimbresFiscales.Text)
RsUpdate.Fields("Otros").Value = CDbl(TxtOtros.Text)

FrmCuentasPorCobrar.NoFact = TxtNoFactura.Text
FrmCuentasPorCobrar.FechaEmi = Format(DtpFechaEmision.Value, "DD/MM/YYYY")
FrmCuentasPorCobrar.PorCobrar = CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text))
FrmCuentasPorCobrar.TasaImpuesto = CDbl(TxtTasaImpuesto.Text)
FrmCuentasPorCobrar.Impuesto = CDbl(TxtImpuesto.Text)
FrmCuentasPorCobrar.SubTotal = CDbl(TxtSubTotal.Text)
FrmCuentasPorCobrar.Monto = CDbl(TxtTotalGeneral.Text)

'FrmCuentasPorCobrar.Impuesto = CDbl(TxtImpuesto.Text)
'FrmCuentasPorCobrar.SubTotal = CDbl(TxtSubTotal.Text)
'FrmCuentasPorCobrar.Monto = CDbl(TxtTotalGeneral.Text)

RsUpdate.Update

MsgBox "Los Cambios han sido registrados!", vbInformation + vbOKOnly, "Operacion Exitosa!"
Unload Me

' Actualiza datos de la tabla CLIENTE
'CSql = "Select * From Cliente Where IdCliente ='" & FrmCuentasPorCobrar.IdCliente & "'"
'Set RsUpdate = CrearRS(CSql)
'RsUpdate.Fields("IdCliente").Value = TxtCodigo.Text
'RsUpdate.Fields("Rif").Value = TxtRif.Text
'RsUpdate.Fields("Razon").Value = TxtRazonSocial.Text
'RsUpdate.Update

' Actualiza datos de la tabla PACIENTE
'CSql = "Select * From Paciente Where IdPaciente ='" & FrmCuentasPorCobrar.IdPaciente & "'"
'Set RsBuscarPaciente = CrearRS(CSql)
'RsBuscarPaciente.Fields("ApellidoP").Value = TxtApellido.Text
'RsBuscarPaciente.Fields("NombreP").Value = TxtNombre.Text
'RsBuscarPaciente.Fields("CedulaP").Value = TxtCedulaPaciente.Text

    FrmCuentasPorCobrar.CargarAbonos
    FrmCuentasPorCobrar.CargarCuentasPorCobrar
    FrmCuentasPorCobrar.Cancelar_Pago

    For i = 1 To FrmCuentasPorCobrar.LstCuentasCobrar.ListItems.Count
        If FrmCuentasPorCobrar.LstCuentasCobrar.ListItems(i).ListSubItems(3).Text = FrmCuentasPorCobrar.NoFact Then
            FrmCuentasPorCobrar.LstCuentasCobrar.ListItems(i).Selected = True
            FrmCuentasPorCobrar.LstCuentasCobrar.ListItems(i).EnsureVisible
        End If
    Next i
End Sub

Private Sub Form_Load()
Centrar Me
CargarFactura
End Sub

Sub CargarFactura()

Dim RsCuentaCobrar As New ADODB.Recordset

CSql = "Select * From C_Cobrar Where (Forma_Pago<>0 or Forma_Pago<> -1) And IdCliente='" & FrmCuentasPorCobrar.IdCliente & "' And IdPaciente='" & FrmCuentasPorCobrar.IdPaciente & "' And N_Factura='" & FrmCuentasPorCobrar.NoFact & "'"
Set RsCuentaCobrar = CrearRS(CSql)
    
CSql = "Select * From Cliente Where IdCliente ='" & FrmCuentasPorCobrar.IdCliente & "'"
Set RsBuscarCliente = CrearRS(CSql)

CSql = "Select * From Paciente Where IdPaciente ='" & FrmCuentasPorCobrar.IdPaciente & "'"
Set RsBuscarPaciente = CrearRS(CSql)

TxtCodigo.Text = RsBuscarCliente.Fields("IdCliente").Value
TxtRif.Text = RsBuscarCliente.Fields("Rif").Value
TxtRazonSocial.Text = RsBuscarCliente.Fields("Razon").Value

If RsBuscarPaciente.RecordCount > 0 Then
    TxtApellido.Text = RsBuscarPaciente.Fields("ApellidoP").Value
    TxtNombre.Text = RsBuscarPaciente.Fields("NombreP").Value
    TxtCedulaPaciente.Text = RsBuscarPaciente.Fields("CedulaP").Value
End If

TxtNoFactura.Text = RsCuentaCobrar.Fields("N_Factura").Value
DtpFechaEmision.Value = RsCuentaCobrar.Fields("Fecha").Value
TxtSubTotal.Text = Format(RsCuentaCobrar.Fields("SubTotal").Value, "#,##0.00")
TxtTotalGeneral.Text = Format(RsCuentaCobrar.Fields("Monto").Value, "#,##0.00")
TxtTasaImpuesto.Text = Format(RsCuentaCobrar.Fields("TasaImpuesto").Value, "#,##0.00")
TxtImpuesto.Text = Format(RsCuentaCobrar.Fields("Impuesto").Value, "#,##0.00")

TxtDescuentos.Text = Format(RsCuentaCobrar.Fields("Descuentos").Value, "#,##0.00")
TxtRetenciones.Text = Format(RsCuentaCobrar.Fields("Retenciones").Value, "#,##0.00")
TxtTimbresFiscales.Text = Format(RsCuentaCobrar.Fields("TimbresFiscales").Value, "#,##0.00")
TxtOtros.Text = Format(RsCuentaCobrar.Fields("Otros").Value, "#,##0.00")
TxtPorCobrar.Text = Format(RsCuentaCobrar.Fields("PorCobrar").Value, "#,##0.00")
End Sub
Private Sub DtpFechaEmision_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtCodigo.SetFocus
        Case vbKeyLeft
            TxtNoFactura.SetFocus
        Case vbKeyDown
            TxtCodigo.SetFocus
    End Select
End If
End Sub

Private Sub TxtApellido_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNombre.SetFocus
        Case vbKeyUp
            TxtCedulaPaciente.SetFocus
        Case vbKeyRight
            TxtNombre.SetFocus
        Case vbKeyDown
            TxtDiasCredito.SetFocus
    End Select
End If
End Sub

Private Sub TxtCedulaPaciente_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtApellido.SetFocus
        Case vbKeyUp
            TxtRazonSocial.SetFocus
        Case vbKeyDown
            TxtApellido.SetFocus
    End Select
End If
End Sub

Private Sub TxtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtRif.SetFocus
        Case vbKeyUp
            TxtNoFactura.SetFocus
        Case vbKeyRight
            TxtRif.SetFocus
        Case vbKeyDown
            TxtRazonSocial.SetFocus
    End Select
End If
End Sub

Private Sub TxtDescuentos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    a = Replace(TxtDescuentos.Text, ".", ",")
    TxtDescuentos.Text = Format(a, "#,##0.00")
    TxtPorCobrar.Text = Format(CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text) + CDbl(TxtOtros.Text)), "#,##0.00")
End If
End Sub

Private Sub TxtDiasCredito_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtPorCobrar.SetFocus
        Case vbKeyUp
            TxtApellido.SetFocus
        Case vbKeyRight
            TxtTasaImpuesto.SetFocus
        Case vbKeyDown
            TxtPorCobrar.SetFocus
    End Select
End If
End Sub

Private Sub TxtImpuesto_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtSubTotal.SetFocus
        Case vbKeyUp
            TxtTasaImpuesto.SetFocus
        Case vbKeyRight
            TxtTotalGeneral.SetFocus
        Case vbKeyLeft
            TxtPorCobrar.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoFactura_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaEmision.SetFocus
        Case vbKeyRight
            DtpFechaEmision.SetFocus
        Case vbKeyDown
            TxtCodigo.SetFocus
    End Select
End If
End Sub

Private Sub TxtNombre_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDiasCredito.SetFocus
        Case vbKeyUp
            TxtCedulaPaciente.SetFocus
        Case vbKeyLeft
            TxtApellido.SetFocus
        Case vbKeyDown
            TxtDiasCredito.SetFocus
    End Select
End If
End Sub

Private Sub TxtOtros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    C = Replace(TxtOtros.Text, ".", ",")
    TxtOtros.Text = Format(C, "#,##0.00")
    TxtPorCobrar.Text = Format(CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text) + CDbl(TxtOtros.Text)), "#,##0.00")
End If
End Sub

Private Sub TxtPorCobrar_Change()
TxtPorCobrar.Text = Format(CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text) + CDbl(TxtOtros.Text)), "#,##0.00")
End Sub

Private Sub TxtPorCobrar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtTasaImpuesto.SetFocus
        Case vbKeyUp
            TxtDiasCredito.SetFocus
        Case vbKeyRight
            TxtImpuesto.SetFocus
        Case vbKeyDown
            BtnGuardarActualizar.SetFocus
    End Select
End If
End Sub

Private Sub TxtRazonSocial_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtCedulaPaciente.SetFocus
        Case vbKeyUp
            TxtCodigo.SetFocus
        Case vbKeyDown
            TxtCedulaPaciente.SetFocus
    End Select
End If
End Sub

Private Sub TxtRetenciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    b = Replace(TxtRetenciones.Text, ".", ",")
    TxtRetenciones.Text = Format(b, "#,##0.00")
    TxtPorCobrar.Text = Format(CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text) + CDbl(TxtOtros.Text)), "#,##0.00")
End If
End Sub

Private Sub TxtRif_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtRazonSocial.SetFocus
        Case vbKeyUp
            TxtNoFactura.SetFocus
        Case vbKeyLeft
            TxtCodigo.SetFocus
        Case vbKeyDown
            TxtRazonSocial.SetFocus
    End Select
End If
End Sub

Private Sub TxtSubTotal_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtTotalGeneral.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyLeft
            TxtTasaImpuesto.SetFocus
        Case vbKeyDown
            TxtTotalGeneral.SetFocus
    End Select
End If
End Sub

Private Sub TxtTasaImpuesto_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtImpuesto.SetFocus
        Case vbKeyUp
            TxtApellido.SetFocus
        Case vbKeyLeft
            TxtDiasCredito.SetFocus
        Case vbKeyRight
            TxtSubTotal.SetFocus
        Case vbKeyDown
            TxtImpuesto.SetFocus
    End Select
End If
End Sub

Private Sub TxtTimbresFiscales_Change()
'TxtPorCobrar.Text = CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text))
End Sub

Private Sub TxtTimbresFiscales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    C = Replace(TxtTimbresFiscales.Text, ".", ",")
    TxtTimbresFiscales.Text = Format(C, "#,##0.00")
    TxtPorCobrar.Text = Format(CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text) + CDbl(TxtOtros.Text)), "#,##0.00")
End If
End Sub

Private Sub TxtTimbresFiscales_LostFocus()

'TxtPorCobrar.Text = Format(CDbl(TxtTotalGeneral.Text) - (CDbl(TxtRetenciones.Text) + CDbl(TxtTimbresFiscales.Text) + CDbl(TxtDescuentos.Text)), "#,##0.00")
End Sub

Private Sub TxtTotalGeneral_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGuardarActualizar.SetFocus
        Case vbKeyUp
            TxtSubTotal.SetFocus
        Case vbKeyLeft
            TxtImpuesto.SetFocus
        Case vbKeyDown
            BtnGuardarActualizar.SetFocus
    End Select
End If
End Sub


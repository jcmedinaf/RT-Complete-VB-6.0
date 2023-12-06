VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEditorFacturas 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Facturas"
   ClientHeight    =   3945
   ClientLeft      =   6450
   ClientTop       =   3705
   ClientWidth     =   9090
   Icon            =   "FrmEditorFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   3000
      Width           =   8895
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   7800
         TabIndex        =   36
         ToolTipText     =   "Cerrar "
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
         MICON           =   "FrmEditorFacturas.frx":1002
         PICN            =   "FrmEditorFacturas.frx":101E
         PICH            =   "FrmEditorFacturas.frx":11E7
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
         TabIndex        =   37
         ToolTipText     =   "Guardar / Actualizar "
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
         MICON           =   "FrmEditorFacturas.frx":141C
         PICN            =   "FrmEditorFacturas.frx":1438
         PICH            =   "FrmEditorFacturas.frx":16C7
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin MSComCtl2.DTPicker DtpFechaRecepcion 
         Height          =   330
         Left            =   7200
         TabIndex        =   33
         Top             =   712
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   60620801
         CurrentDate     =   40141
      End
      Begin VB.TextBox TxtDescuentos 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   32
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox TxtPorPagar 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1200
         TabIndex        =   30
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalGeneral 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7200
         TabIndex        =   28
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TxtExentos 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TxtImpuesto 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1200
         TabIndex        =   24
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TxtSubTotal 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7200
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalBase 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TxtDiasCredito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TxtRif 
         Height          =   300
         Left            =   3960
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtNombre 
         Height          =   300
         Left            =   1200
         TabIndex        =   14
         Top             =   1440
         Width           =   7575
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtNoFactura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtNoControl 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtNoOrden 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtNoCompra 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DtpFechaEmision 
         Height          =   330
         Left            =   7200
         TabIndex        =   34
         Top             =   352
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60620801
         CurrentDate     =   40141
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   3000
         TabIndex        =   31
         Top             =   2580
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Pagar:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2580
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total General:"
         Height          =   195
         Left            =   6120
         TabIndex        =   27
         Top             =   2220
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exentos:"
         Height          =   195
         Left            =   3000
         TabIndex        =   25
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impuesto:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   2220
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal:"
         Height          =   195
         Left            =   6120
         TabIndex        =   21
         Top             =   1860
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Base:"
         Height          =   195
         Left            =   3000
         TabIndex        =   19
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Días Crédito:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1860
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rif:"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1500
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Recepción:"
         Height          =   195
         Left            =   5760
         TabIndex        =   10
         Top             =   780
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Emisión:"
         Height          =   195
         Left            =   5760
         TabIndex        =   9
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Factura:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Control:"
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Orden:"
         Height          =   195
         Left            =   3000
         TabIndex        =   3
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Compra:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmEditorFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarCuentasPorPagar As New ADODB.Recordset
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim RsUpdate As New ADODB.Recordset

'    If (DateValue(Format(DtpFechaEmision.Value, "DD/MM/YY")) - DateValue(Now)) <= 0 Then
'        TxtDiasCredito.Text = "0"
'        Else
'        TxtDiasCredito.Text = Format((DateValue(Format(DtpFechaEmision.Value, "DD/MM/YY")) - DateValue(Now)), "DD/MM/YYYY")
'    End If

    TxtDiasCredito.Text = FrmCuentasPorPagar.DiasCredit
    
    CSql = "Select * From CtaPorPagar Where IdProveedor='" & FrmCuentasPorPagar.IdProv & "' And NoFactura='" & FrmCuentasPorPagar.NoFact & "'"
    Set RsUpdate = CrearRS(CSql)

    RsUpdate.Fields("IdProveedor").Value = TxtCodigo.Text
    RsUpdate.Fields("Nombre").Value = TxtNombre.Text
    RsUpdate.Fields("Rif").Value = TxtRif.Text
    RsUpdate.Fields("FechaEmision").Value = Format(DtpFechaEmision.Value, "DD/MM/YYYY")
    RsUpdate.Fields("FechaRecepcion").Value = Format(DtpFechaRecepcion.Value, "DD/MM/YYYY")
    RsUpdate.Fields("NoFactura").Value = TxtNoFactura.Text
    RsUpdate.Fields("NoControl").Value = TxtNoControl.Text
    RsUpdate.Fields("NumeroOrden").Value = TxtNoOrden.Text
    RsUpdate.Fields("NumeroCompra").Value = TxtNoCompra.Text
    RsUpdate.Fields("DiasCredito").Value = TxtDiasCredito.Text
    RsUpdate.Fields("Impuesto").Value = CDbl(TxtImpuesto.Text)
    RsUpdate.Fields("TotalExentos").Value = CDbl(TxtExentos.Text)
    RsUpdate.Fields("SubTotal").Value = CDbl(TxtSubTotal.Text)
    RsUpdate.Fields("TotalGeneral").Value = CDbl(TxtTotalGeneral.Text)
    RsUpdate.Fields("TotalBase").Value = CDbl(TxtTotalBase.Text)
    RsUpdate.Fields("Descuento").Value = CDbl(TxtDescuentos.Text)
    RsUpdate.Fields("PorPagar").Value = CDbl(TxtPorPagar.Text)
    RsUpdate.Update
    
    
    FrmCuentasPorPagar.NoComp = TxtNoCompra.Text
    FrmCuentasPorPagar.NoOrd = TxtNoOrden.Text
    FrmCuentasPorPagar.NoFact = TxtNoFactura.Text
    FrmCuentasPorPagar.NoCont = TxtNoControl.Text
    FrmCuentasPorPagar.FechaEmi = Format(DtpFechaEmision.Value, "DD/MM/YYYY")
    FrmCuentasPorPagar.IdProv = TxtCodigo.Text
    FrmCuentasPorPagar.TotalG = CDbl(TxtTotalGeneral.Text)
    FrmCuentasPorPagar.PorPag = CDbl(TxtPorPagar.Text)

    
    MsgBox "Los Cambios han sido registrados!", vbInformation + vbOKOnly, "Operacion Exitosa!"
    Unload Me

    FrmCuentasPorPagar.CargarAbonos
    FrmCuentasPorPagar.CargarCuentasPorPagar
    FrmCuentasPorPagar.TotalesCuentasPorPagar
    FrmCuentasPorPagar.Cancelar_Pago

    For i = 1 To FrmCuentasPorPagar.LstCuentasPagar.ListItems.Count
        If FrmCuentasPorPagar.LstCuentasPagar.ListItems(i).ListSubItems(3).Text = FrmCuentasPorPagar.NoFact Then
            FrmCuentasPorPagar.LstCuentasPagar.ListItems(i).Selected = True
            FrmCuentasPorPagar.LstCuentasPagar.ListItems(i).EnsureVisible
        End If
    Next i
End Sub

Private Sub Form_Load()
Centrar Me
CargarFactura
End Sub

Sub CargarFactura()
CSql = "Select * From CtaPorPagar Where IdProveedor='" & FrmCuentasPorPagar.IdProv & "' And NoFactura='" & FrmCuentasPorPagar.NoFact & "'"
Set RsCargarCuentasPorPagar = CrearRS(CSql)
   
TxtCodigo.Text = RsCargarCuentasPorPagar.Fields("IdProveedor").Value
TxtNombre.Text = RsCargarCuentasPorPagar.Fields("Nombre").Value
TxtRif.Text = RsCargarCuentasPorPagar.Fields("Rif").Value
DtpFechaEmision = RsCargarCuentasPorPagar.Fields("FechaEmision").Value
DtpFechaRecepcion = RsCargarCuentasPorPagar.Fields("FechaRecepcion").Value
TxtNoFactura.Text = RsCargarCuentasPorPagar.Fields("NoFactura").Value
TxtNoControl.Text = RsCargarCuentasPorPagar.Fields("NoControl").Value
TxtNoOrden.Text = RsCargarCuentasPorPagar.Fields("NumeroOrden").Value
TxtNoCompra.Text = RsCargarCuentasPorPagar.Fields("NumeroCompra").Value
TxtDiasCredito.Text = RsCargarCuentasPorPagar.Fields("DiasCredito").Value
TxtImpuesto.Text = Format(RsCargarCuentasPorPagar.Fields("Impuesto").Value, "#,##0.00")
TxtExentos.Text = Format(RsCargarCuentasPorPagar.Fields("TotalExentos").Value, "#,##0.00")
TxtSubTotal.Text = Format(RsCargarCuentasPorPagar.Fields("SubTotal").Value, "#,##0.00")
TxtTotalGeneral.Text = Format(RsCargarCuentasPorPagar.Fields("TotalGeneral").Value, "#,##0.00")
TxtTotalBase.Text = Format(RsCargarCuentasPorPagar.Fields("TotalBase").Value, "#,##0.00")
TxtDescuentos.Text = Format(RsCargarCuentasPorPagar.Fields("Descuento").Value, "#,##0.00")
TxtPorPagar.Text = Format(RsCargarCuentasPorPagar.Fields("PorPagar").Value, "#,##0.00")
    
RsCargarCuentasPorPagar.Close
End Sub

Private Sub DtpFechaEmision_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaRecepcion.SetFocus
        Case vbKeyLeft
            TxtNoOrden.SetFocus
        Case vbKeyDown
            DtpFechaRecepcion.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaRecepcion_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNombre.SetFocus
        Case vbKeyUp
            DtpFechaEmision.SetFocus
        Case vbKeyLeft
            TxtNoControl.SetFocus
        Case vbKeyDown
            TxtNombre.SetFocus
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
            TxtNombre.SetFocus
    End Select
End If
End Sub

Private Sub TxtDescuentos_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGuardarActualizar.SetFocus
        Case vbKeyUp
            TxtExentos.SetFocus
        Case vbKeyLeft
            TxtPorPagar.SetFocus
        Case vbKeyRight
            TxtTotalGeneral.SetFocus
        Case vbKeyDown
            BtnGuardarActualizar.SetFocus
    End Select
End If
End Sub

Private Sub TxtDiasCredito_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtImpuesto.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyRight
            TxtTotalBase.SetFocus
        Case vbKeyDown
            TxtImpuesto.SetFocus
    End Select
End If
End Sub

Private Sub TxtExentos_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDescuentos.SetFocus
        Case vbKeyUp
            TxtTotalBase.SetFocus
        Case vbKeyLeft
            TxtImpuesto.SetFocus
        Case vbKeyRight
            TxtTotalGeneral.SetFocus
        Case vbKeyDown
            TxtDescuentos.SetFocus
    End Select
End If
End Sub

Private Sub TxtImpuesto_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtPorPagar.SetFocus
        Case vbKeyUp
            TxtDiasCredito.SetFocus
        Case vbKeyRight
            TxtExentos.SetFocus
        Case vbKeyDown
            TxtPorPagar.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoCompra_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoOrden.SetFocus
        Case vbKeyRight
            TxtNoOrden.SetFocus
        Case vbKeyDown
            TxtNoFactura.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoControl_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtCodigo.SetFocus
        Case vbKeyUp
            TxtNoOrden.SetFocus
        Case vbKeyLeft
            TxtNoFactura.SetFocus
        Case vbKeyRight
            DtpFechaRecepcion.SetFocus
        Case vbKeyDown
            TxtRif.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoFactura_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoControl.SetFocus
        Case vbKeyUp
            TxtNoCompra.SetFocus
        Case vbKeyRight
            TxtNoControl.SetFocus
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
            TxtCodigo.SetFocus
        Case vbKeyDown
            TxtDiasCredito.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoOrden_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoFactura.SetFocus
        Case vbKeyLeft
            TxtNoCompra.SetFocus
        Case vbKeyRight
            DtpFechaEmision.SetFocus
        Case vbKeyDown
            TxtNoControl.SetFocus
    End Select
End If
End Sub

Private Sub TxtPorPagar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtTotalBase.SetFocus
        Case vbKeyUp
            TxtImpuesto.SetFocus
        Case vbKeyRight
            TxtDescuentos.SetFocus
        Case vbKeyDown
            BtnGuardarActualizar.SetFocus
    End Select
End If
End Sub

Private Sub TxtRif_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaEmision.SetFocus
        Case vbKeyUp
            TxtNoControl.SetFocus
        Case vbKeyLeft
            TxtCodigo.SetFocus
        Case vbKeyRight
            DtpFechaRecepcion.SetFocus
        Case vbKeyDown
            TxtNombre.SetFocus
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
            TxtTotalBase.SetFocus
        Case vbKeyDown
            TxtTotalGeneral.SetFocus
    End Select
End If
End Sub

Private Sub TxtTotalBase_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtExentos.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyLeft
            TxtDiasCredito.SetFocus
        Case vbKeyRight
            TxtSubTotal.SetFocus
        Case vbKeyDown
            TxtExentos.SetFocus
    End Select
End If
End Sub

Private Sub TxtTotalGeneral_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGuardarActualizar.SetFocus
        Case vbKeyUp
            TxtSubTotal.SetFocus
        Case vbKeyLeft
            TxtExentos.SetFocus
        Case vbKeyDown
            BtnGuardarActualizar.SetFocus
    End Select
End If
End Sub


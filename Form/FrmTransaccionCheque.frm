VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTransaccionCheques 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Cheques"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "FrmTransaccionCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   8655
      Begin VB.TextBox TxtPaguese 
         Height          =   375
         Left            =   1920
         MaxLength       =   300
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1200
         Width           =   5160
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Imprimir Comprobante"
         Height          =   255
         Left            =   6600
         TabIndex        =   9
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox ChkAnular 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Anular Cheque"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox ChkNoEndosable 
         BackColor       =   &H00EAEFEF&
         Caption         =   "No Endosable"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TxtBancos 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtNoCheque 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   735
         Left            =   720
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1680
         Width           =   7815
      End
      Begin VB.TextBox TxtIdBanco 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPickerFecha 
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51118081
         CurrentDate     =   40232
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   7200
         TabIndex        =   3
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
         MICON           =   "FrmTransaccionCheque.frx":1002
         PICN            =   "FrmTransaccionCheque.frx":101E
         PICH            =   "FrmTransaccionCheque.frx":1283
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscarBeneficiario 
         Height          =   375
         Left            =   7200
         TabIndex        =   5
         ToolTipText     =   "Buscar Benefiario"
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
         MICON           =   "FrmTransaccionCheque.frx":1515
         PICN            =   "FrmTransaccionCheque.frx":1531
         PICH            =   "FrmTransaccionCheque.frx":1796
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblNoEndosable 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Endosable"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paguese a la Orden de:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1290
         Width           =   1680
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cheque:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   5760
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3120
         TabIndex        =   16
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   8655
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2520
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7560
         TabIndex        =   12
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
         MICON           =   "FrmTransaccionCheque.frx":1A28
         PICN            =   "FrmTransaccionCheque.frx":1A44
         PICH            =   "FrmTransaccionCheque.frx":1C0D
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
         TabIndex        =   14
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
         MICON           =   "FrmTransaccionCheque.frx":1E42
         PICN            =   "FrmTransaccionCheque.frx":1E5E
         PICH            =   "FrmTransaccionCheque.frx":1FEB
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
         TabIndex        =   11
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
         MICON           =   "FrmTransaccionCheque.frx":2220
         PICN            =   "FrmTransaccionCheque.frx":223C
         PICH            =   "FrmTransaccionCheque.frx":251E
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
         TabIndex        =   10
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
         MICON           =   "FrmTransaccionCheque.frx":276F
         PICN            =   "FrmTransaccionCheque.frx":278B
         PICH            =   "FrmTransaccionCheque.frx":2A1A
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
Attribute VB_Name = "FrmTransaccionCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGuardar As New ADODB.Recordset
Dim RsMaxId As New ADODB.Recordset
Dim RsVerificarSaldo As New ADODB.Recordset
Private Sub BtnAgregar_Click()
Blanqueo
Editar = 0
Frame2.Enabled = True

BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
TxtNoCheque.Enabled = True
TxtNoCheque.SetFocus
End Sub

Private Sub BtnBuscar_Click()
Ban = 3
FrmListadoBancos.Show vbModal
End Sub

Private Sub BtnBuscarBeneficiario_Click()

FrmListadoBeneficiarios.Show vbModal
End Sub

Private Sub BtnCerrar_Click()
Editar = 0
Unload Me
End Sub

Sub Blanqueo()
TxtNoCheque.Text = ""
TxtNoCheque.Text = ""
TxtBancos.Text = ""
TxtIdBanco.Text = ""
TxtMonto.Text = ""
DTPickerFecha.Value = DateTime.Date
TxtPaguese.Text = ""
TxtDetalle.Text = ""
ChkAnular.Value = 0
ChkNoEndosable.Value = 0
Check3.Value = 0
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub BtnGuardarActualizar_Click()

'###################################################################
'Validacion de campos

If TxtNoCheque.Text = "" Then
    MsgBox "Debe de ingresar el numero del cheque", vbCritical + vbOKOnly, "Error"
    TxtNoCheque.SetFocus
    Exit Sub

ElseIf TxtMonto.Text = "" Then
    MsgBox "Debe de ingresar el monto", vbCritical + vbOKOnly, "Error"
    TxtMonto.SetFocus
    Exit Sub

ElseIf TxtIdBanco.Text = "" Then
    MsgBox "Falta seleccionar el Banco", vbCritical + vbOKOnly, "Error"
    TxtIdBanco.SetFocus
    Exit Sub

ElseIf TxtPaguese.Text = "" Then
    MsgBox "Debe de ingresar el nombre a quien va dirigido el cheque", vbCritical + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
  
ElseIf TxtDetalle.Text = "" Then
    MsgBox "Debe de ingresar un breve detalle del movimiento", vbCritical + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
End If

'###################################################################
'Verifica si hay saldo en la cuenta para poder realizar el cheque

CSql = "Select * From Movi_BanCaja Where IdCajaBanco='" & TxtIdBanco.Text & "' And Ingr_Egr='1'"
Set RsVerificarSaldo = CrearRS(CSql)

If RsVerificarSaldo.RecordCount = 0 Then
    MsgBox "Esta Cuenta No Posee Saldo para realizar algun cheque!!!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If


Select Case Editar

Case Is = 0
'###################################################################
'Agregar Cheque


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
RsGuardar.Fields("n_comprobante").Value = TxtNoCheque.Text
RsGuardar.Fields("Monto_Mov").Value = CDbl(TxtMonto.Text)
RsGuardar.Fields("Detalle").Value = TxtDetalle.Text
RsGuardar.Fields("Beneficiario").Value = TxtPaguese.Text
RsGuardar.Fields("Tipo_Mov").Value = 2
RsGuardar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
RsGuardar.Fields("Conciliado").Value = 0
RsGuardar.Fields("Anulado").Value = ChkAnular.Value
RsGuardar.Fields("NoEndosable").Value = ChkNoEndosable.Value
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Update

MsgBox "Registro Almacenado con Exito!", vbInformation + vbOKOnly, "Movimiento Guardado"

FrmLibroBancos.Movi


'------------------------
'imprime la forma del cheque

If Check3.Value = 1 Then
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\Recibo_Cheque.rpt"
        .Connect = "DSN=CrReporte;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{MoviBancoCaja.IdMovCajaBanco} ='" & IdMax
        .WindowTitle = "Recibo de cheque No. " & TxtNoCheque.Text
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If


Case Is = 1

CSql = "Select * From Movi_BanCaja Where IdMovCajaBanco = '" & IdReg & "'"
Set RsGuardar = CrearRS(CSql)


RsGuardar.Fields("IdCajaBanco").Value = TxtIdBanco.Text
RsGuardar.Fields("Ingr_Egr").Value = 2
RsGuardar.Fields("n_comprobante").Value = TxtNoCheque.Text
RsGuardar.Fields("Monto_Mov").Value = CDbl(TxtMonto.Text)
RsGuardar.Fields("Detalle").Value = TxtDetalle.Text
RsGuardar.Fields("Beneficiario").Value = TxtPaguese.Text
RsGuardar.Fields("Tipo_Mov").Value = 2
RsGuardar.Fields("Fecha_Transa").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
RsGuardar.Fields("Conciliado").Value = 0
RsGuardar.Fields("Anulado").Value = ChkAnular.Value
RsGuardar.Fields("NoEndosable").Value = ChkNoEndosable.Value
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Update

MsgBox "Registro Actualizado con Exito!", vbInformation + vbOKOnly, "Movimiento Actualizado"

FrmLibroBancos.Movi

'------------------------
'imprime la forma del cheque

If Check3.Value = 1 Then
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\Recibo_Cheque.rpt"
        .Connect = "DSN=CrReporte;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{MoviBancoCaja.IdMovCajaBanco} =" & IdReg
        .WindowTitle = "Recibo de cheque No. " & TxtNoCheque.Text
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If


End Select


BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
Frame2.Enabled = False

BtnDesHacer_Click
Unload Me

Exit Sub


WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
Editar = 0
Unload Me

End Sub


Private Sub ChkAnular_Click()
If ChkAnular.Value = 1 Then
    TxtNoCheque.Enabled = False
    TxtDetalle.Enabled = False
    TxtPaguese.Enabled = False
    TxtDetalle.Text = "ANULADO"
    TxtPaguese.Text = "ANULADO"
Else
    TxtDetalle.Text = ""
    TxtPaguese.Text = ""
    TxtNoCheque.Enabled = True
    TxtDetalle.Enabled = True
    TxtPaguese.Enabled = True
End If
End Sub

Private Sub ChkNoEndosable_Click()
If ChkNoEndosable.Value = 1 Then
    LblNoEndosable.Visible = True
Else
    LblNoEndosable.Visible = False
End If
End Sub

'///////////////////////////////////Valido TextBox: DTPickerFecha//////////////////////////////
Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar.SetFocus
Else
    If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_Load()

If Editar = 1 Then
    Frame2.Enabled = True
    'DTPickerFecha.Value = DateTime.Date
    BtnGuardarActualizar.Enabled = True


    CSql = "Select * From Movi_BanCaja Where IdMovCajaBanco = '" & IdReg & "'" 'IdCajaBanco = '" & txtCodigo.Text & "' And Fecha_Transa >='" & Trim(TxtFechaDesde.Text) & "' and Fecha_Transa <='" & Trim(TxtFechaHasta.Text) & "' Order by Fecha_Transa asc"
    Set RsCargarMovimientos = CrearRS(CSql)
    
    TxtNoCheque.Text = RsCargarMovimientos.Fields("n_comprobante").Value
    TxtMonto.Text = RsCargarMovimientos.Fields("Monto_Mov").Value
    TxtIdBanco.Text = RsCargarMovimientos.Fields("IdCajaBanco").Value
    TxtDetalle.Text = RsCargarMovimientos.Fields("Detalle").Value
    TxtPaguese.Text = RsCargarMovimientos.Fields("Beneficiario").Value
    DTPickerFecha.Value = RsCargarMovimientos.Fields("Fecha_Transa").Value
    ChkAnular.Value = RsCargarMovimientos.Fields("Anulado").Value
    ChkNoEndosable.Value = RsCargarMovimientos.Fields("NoEndosable").Value
    
    CSql = "Select * From CajasBancos Where IdCajaBanco = '" & TxtIdBanco.Text & "'"  'IdCajaBanco = '" & txtCodigo.Text & "' And Fecha_Transa >='" & Trim(TxtFechaDesde.Text) & "' and Fecha_Transa <='" & Trim(TxtFechaHasta.Text) & "' Order by Fecha_Transa asc"
    Set RsCargarMovimientos = CrearRS(CSql)
    TxtBancos.Text = RsCargarMovimientos.Fields("Descripcion").Value
    
Else
    Frame2.Enabled = False
    DTPickerFecha.Value = DateTime.Date
    BtnGuardarActualizar.Enabled = False
End If
End Sub

'///////////////////////////////////Valido TextBox: DTPickerFecha//////////////////////////////
Private Sub DTPickerFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMonto.SetFocus
End If
End Sub

'///////////////////////////////////Valido TextBox: TxtNoCheque//////////////////////////////
Private Sub TxtNoCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DTPickerFecha.SetFocus
Else
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

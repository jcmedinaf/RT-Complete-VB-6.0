VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoBancos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Bancos"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   Icon            =   "FrmListadoBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   7080
         Width           =   3975
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
            TabIndex        =   5
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Número de Cuenta o Descripción"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "FrmListadoBancos.frx":1002
            PICN            =   "FrmListadoBancos.frx":101E
            PICH            =   "FrmListadoBancos.frx":1283
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
         Height          =   735
         Left            =   4200
         TabIndex        =   1
         Top             =   7080
         Width           =   3135
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   720
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   2040
            TabIndex        =   2
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
            MICON           =   "FrmListadoBancos.frx":1515
            PICN            =   "FrmListadoBancos.frx":1531
            PICH            =   "FrmListadoBancos.frx":16FA
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
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   6735
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   11880
         Object.Width           =   7185
         Object.Height          =   6705
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarBancos As New ADODB.Recordset
Dim RsBuscarBancos As New ADODB.Recordset
Dim RsSeleccionarBancos As New ADODB.Recordset

Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From CajasBancos Where (Descripcion like '%" & Trim(TxtBuscar.Text) & "%' or N_Cuenta like '%" & Trim(TxtBuscar.Text) & "%') AND activo=1"
Else
    CSql = "Select * From CajasBancos where activo=1"
End If

Set RsBuscarBancos = CrearRS(CSql)

If RsBuscarBancos.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarBancos.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarBancos.Fields("IdCajaBanco").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarBancos.Fields("Descripcion").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarBancos.Fields("N_Cuenta").Value
        RsBuscarBancos.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If
RsBuscarBancos.Close

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    
    CSql = "Select * From CajasBancos Where IdCajaBanco='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsSeleccionarBancos = CrearRS(CSql)
    If RsSeleccionarBancos.RecordCount > 0 Then
                
        Select Case Ban
            Case Is = 1
                'IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
                FrmLibroBancos.TxtCodigo.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmLibroBancos.TxtNombreBanco.Text = RsSeleccionarBancos.Fields("Descripcion").Value
                FrmLibroBancos.TxtNoCuenta.Text = RsSeleccionarBancos.Fields("N_Cuenta").Value
            
            Case Is = 2
                FrmTransaccionDepositos.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmTransaccionDepositos.TxtBancos.Text = RsSeleccionarBancos.Fields("Descripcion").Value
            
            Case Is = 3
                FrmTransaccionCheques.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmTransaccionCheques.TxtBancos.Text = RsSeleccionarBancos.Fields("Descripcion").Value
            
            Case Is = 4
                FrmTransaccionNotaCredito.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmTransaccionNotaCredito.TxtBancos.Text = RsSeleccionarBancos.Fields("Descripcion").Value
            
            Case Is = 5
                FrmTransaccionNotasDebitos.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmTransaccionNotasDebitos.TxtBancos.Text = RsSeleccionarBancos.Fields("Descripcion").Value
            
            Case Is = 6
                'IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
                FrmConcilicacionBancaria.TxtCodigo.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmConcilicacionBancaria.TxtNombreBanco.Text = RsSeleccionarBancos.Fields("Descripcion").Value
                FrmConcilicacionBancaria.TxtNoCuenta.Text = RsSeleccionarBancos.Fields("N_Cuenta").Value
            
            Case Is = 7
                'IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
                FrmConciliacionEstadoCuentas.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmConciliacionEstadoCuentas.TxtBancos.Text = RsSeleccionarBancos.Fields("Descripcion").Value
                'FrmConciliacionEstadoCuentas.TxtNoCuenta.Text = RsSeleccionarBancos.Fields("N_Cuenta").Value
      
            Case Is = 8
                FrmTransferenciaFondos.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmTransferenciaFondos.TxtBancos.Text = RsSeleccionarBancos.Fields("Descripcion").Value
                       
            Case Is = 9
                FrmTransferenciaFondos.TxtIdBancoDestino.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmTransferenciaFondos.TxtBancosDestino.Text = RsSeleccionarBancos.Fields("Descripcion").Value
                
            Case Is = 10
                FrmComprobanteRetencion.TxtIdBanco.Text = RsSeleccionarBancos.Fields("IdCajaBanco").Value
                FrmComprobanteRetencion.TxtBanco.Text = RsSeleccionarBancos.Fields("Descripcion").Value
                
                        
        End Select
    End If
    RsSeleccionarBancos.Close
    Unload Me
End If
End Sub

Private Sub Form_Load()
Grid1


CSql = "Select * From CajasBancos"
Set RsCargarBancos = CrearRS(CSql)

Do While Not RsCargarBancos.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarBancos.Fields("IdCajaBanco").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarBancos.Fields("Descripcion").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarBancos.Fields("N_Cuenta").Value
        RsCargarBancos.MoveNext
Loop

DMGrid1.PaintMGrid

End Sub

Sub Grid1()
DMGrid1.Rows = 1
DMGrid1.Cols = 3
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(1).Width = 1000
DMGrid1.DColumnas(2).Width = 3000
DMGrid1.DColumnas(3).Width = 3000
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre o Razón Social del Banco"
DMGrid1.DColumnas(3).Caption = "Cuenta Banco"

End Sub


Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConcilicacionBancaria 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación Bancaria"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15330
   Icon            =   "FrmConciliacionb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   15330
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   37
         Top             =   7080
         Width           =   6375
         Begin VB.TextBox TxtFechaHasta 
            Height          =   375
            Left            =   3840
            TabIndex        =   42
            ToolTipText     =   "Ingrese la Fecha de conciliación hasta"
            Top             =   270
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DtpFechaDesde 
            Height          =   375
            Left            =   2400
            TabIndex        =   40
            Top             =   270
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   52887553
            CurrentDate     =   40242
         End
         Begin VB.TextBox TxtFechaDesde 
            Height          =   375
            Left            =   1200
            TabIndex        =   39
            ToolTipText     =   "Ingrese la Fecha de conciliación desde"
            Top             =   270
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DtpFechaHasta 
            Height          =   375
            Left            =   5040
            TabIndex        =   43
            Top             =   270
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   52887553
            CurrentDate     =   40242
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarFiltro 
            Height          =   375
            Left            =   5400
            TabIndex        =   44
            ToolTipText     =   "Buscar"
            Top             =   270
            Width           =   855
            _ExtentX        =   1508
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
            MICON           =   "FrmConciliacionb.frx":1002
            PICN            =   "FrmConciliacionb.frx":101E
            PICH            =   "FrmConciliacionb.frx":1283
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   2850
            TabIndex        =   41
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.TextBox TxtChequesTransitos 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   6600
         Width           =   2175
      End
      Begin VB.TextBox TxtSaldoDisponible 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   12840
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   6600
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Height          =   975
         Left            =   10920
         TabIndex        =   22
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton OptDiferencias 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Diferencias"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton OptTodos 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   855
         End
         Begin VB.OptionButton OptConciliados 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Conciliados"
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   240
            Width           =   1170
         End
         Begin VB.OptionButton OptNoConciliados 
            BackColor       =   &H00EAEFEF&
            Caption         =   "No Conciliados"
            Height          =   255
            Left            =   2280
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtNoConciliado 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   6600
         Width           =   2175
      End
      Begin VB.TextBox TxtConciliado 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   6600
         Width           =   2175
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   6600
         TabIndex        =   6
         Top             =   7080
         Width           =   8415
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   1440
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
            Left            =   7320
            TabIndex        =   7
            ToolTipText     =   "Cerrar Tablas de Pacientes"
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
            MICON           =   "FrmConciliacionb.frx":1515
            PICN            =   "FrmConciliacionb.frx":1531
            PICH            =   "FrmConciliacionb.frx":16FA
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
            Left            =   6120
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
            MICON           =   "FrmConciliacionb.frx":192F
            PICN            =   "FrmConciliacionb.frx":194B
            PICH            =   "FrmConciliacionb.frx":1C2D
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
            Left            =   3720
            TabIndex        =   9
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
            MICON           =   "FrmConciliacionb.frx":1E7E
            PICN            =   "FrmConciliacionb.frx":1E9A
            PICH            =   "FrmConciliacionb.frx":2130
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
            Left            =   3120
            TabIndex        =   10
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
            MICON           =   "FrmConciliacionb.frx":238F
            PICN            =   "FrmConciliacionb.frx":23AB
            PICH            =   "FrmConciliacionb.frx":2640
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnImprimir 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Imprimir"
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
            MICON           =   "FrmConciliacionb.frx":289C
            PICN            =   "FrmConciliacionb.frx":28B8
            PICH            =   "FrmConciliacionb.frx":29DD
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Movimientos Bancarios Disponibles"
         Height          =   5175
         Left            =   3240
         TabIndex        =   5
         Top             =   1320
         Width           =   11775
         Begin VB.TextBox TxtTotalMovimientos 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "0"
            Top             =   4800
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Cuenta Conciliada"
            Height          =   195
            Left            =   9960
            TabIndex        =   32
            Top             =   4800
            Width           =   1695
         End
         Begin SystemOncoAmerica.DMGrid DMGrid2 
            Height          =   4455
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   7858
            Object.Width           =   11505
            Object.Height          =   4425
            Cols            =   8
            Rows            =   1
            MarqueeStyle    =   2
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Movimientos:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   4860
            Width           =   1560
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Conciliaciones Disponibles"
         Height          =   5175
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
         Begin VB.TextBox TxtTotalConciliaciones 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "0"
            Top             =   4800
            Width           =   855
         End
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   4455
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   7858
            Object.Width           =   2745
            Object.Height          =   4425
            ScrollBar       =   1
            MarqueeStyle    =   2
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Conciliaciones:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   4860
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información Bancaria:"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10695
         Begin VB.TextBox TxtNoCuenta 
            Height          =   375
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox TxtNombreBanco 
            Height          =   375
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   9360
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmConciliacionb.frx":2C6D
            PICN            =   "FrmConciliacionb.frx":2C89
            PICH            =   "FrmConciliacionb.frx":2EEE
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
            Caption         =   "No. de Cuenta Bancaria"
            Height          =   195
            Left            =   6480
            TabIndex        =   17
            Top             =   240
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre o Razón Social del Banco"
            Height          =   195
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   2445
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque en transitos:"
         Height          =   195
         Left            =   7320
         TabIndex        =   29
         Top             =   6690
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Disponible:"
         Height          =   195
         Left            =   11520
         TabIndex        =   24
         Top             =   6690
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Conciliado:"
         Height          =   195
         Left            =   3600
         TabIndex        =   21
         Top             =   6690
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conciliado:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   6690
         Width           =   780
      End
   End
   Begin VB.Menu MnuTransaccion 
      Caption         =   "Transacción"
      Begin VB.Menu MnuEstadoCuenta 
         Caption         =   "Estado de Cuenta"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConciliar 
         Caption         =   "Conciliar"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCerrar 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "FrmConcilicacionBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarMovimientos As New ADODB.Recordset
Dim RsTotalConci As New ADODB.Recordset
Dim RsTotalNoConci As New ADODB.Recordset
Dim RsChequesTransitos As New ADODB.Recordset
Private Sub BtnBuscar_Click()
Ban = 6
FrmListadoBancos.Show vbModal, FrmPrincipal
Movi
Conciliacion
TotalConciliado
TotalNoConciliado
ChequesTransitos
End Sub

Private Sub BtnBuscarFiltro_Click()
    
If TxtCodigo.Text = "" Then
    MsgBox "Seleccione un banco para poder realizar la busqueda!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If
    
If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
    Grid2
    CSql = "Select * From Movi_BanCaja Where Fecha_Transa>='" & TxtFechaDesde.Text & "' And Fecha_Transa<='" & TxtFechaHasta.Text & "' And IdCajaBanco='" & TxtCodigo.Text & "'"
    Set RsCargarMovimientos = CrearRS(CSql)
    DMGrid2.Rows = 0
    DMGrid2.Clear
    If RsCargarMovimientos.RecordCount = 0 Then Exit Sub
    If RsCargarMovimientos.RecordCount > 0 Then
        Do While Not RsCargarMovimientos.EOF
            
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsCargarMovimientos.Fields("Fecha_Transa").Value
            
            If RsCargarMovimientos.Fields("Tipo_Mov").Value = 1 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Efe."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 2 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Che."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 3 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Dep."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 4 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Tran."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 5 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Cre."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 6 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Deb."
            End If
             
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsCargarMovimientos.Fields("n_comprobante").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsCargarMovimientos.Fields("Detalle").Value
            
            If RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
            End If
            
            If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
            End If
             
            If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
            Else
                DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
            End If
            
            If RsCargarMovimientos.Fields("Conciliado").Value = True Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "SI"
            Else
                DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO"
            End If
            RsCargarMovimientos.MoveNext
        Loop
        
        DMGrid2.PaintMGrid
        TxtTotalMovimientos.Text = DMGrid2.Rows
        TxtSaldoDisponible.Text = Format(DMGrid2.ValorCelda(DMGrid2.Rows, 7), "#,##0.00")
    Else
        TxtTotalMovimientos.Text = 0
        TxtSaldoDisponible.Text = Format(0, "#,##0.00")
    End If
Else
    Movi
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Movi
Conciliacion
TxtFechaDesde.Text = ""
TxtFechaHasta.Text = ""
End Sub

Private Sub BtnImprimir_Click()
If TxtCodigo.Text = "" Then
    MsgBox "Debe de seleccionar un banco para poder imprimir su estado de cuenta!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

If TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
    ''========= ESTE ES EL CODIGO NUEVO ==========
    With CrystalReport1
        .ReportFileName = RutaInformes & "\EstadoConciliacion.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{EstadoDeCuentas.IdCajaBanco} = " & TxtCodigo.Text
        '.ReportTitle = "Reporte Orden de Compras No. " & LblNoOrden.Caption
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
Else
 ''========= ESTE ES EL CODIGO NUEVO ==========
    With CrystalReport1
        .ReportFileName = RutaInformes & "\EstadoConciliacion.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{EstadoDeCuentas.IdCajaBanco} = " & TxtCodigo.Text & " And {EstadoDeCuentas.Fecha_Transa}>=" & FechaSQL(TxtFechaDesde.Text) & " And {EstadoDeCuentas.Fecha_Transa}<=" & FechaSQL(TxtFechaHasta.Text) & ""
        '.ReportTitle = "Reporte Orden de Compras No. " & LblNoOrden.Caption
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With

End If
End Sub

Private Sub Check1_Click()
If DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO" Then
    DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO"
    DMGrid2.PaintMGrid
Else
    DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "SI"
    DMGrid2.PaintMGrid
End If

End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton Then
    Grid2
    CSql = "Select * From Movi_BanCaja Where Conciliado=1 And FechaConciliacion='" & DMGrid1.ValorCelda(lRow, 2) & "'"
    Set RsCargarMovimientos = CrearRS(CSql)
    DMGrid2.Rows = 0
    DMGrid2.Clear
    If RsCargarMovimientos.RecordCount = 0 Then Exit Sub
    If RsCargarMovimientos.RecordCount > 0 Then
        Do While Not RsCargarMovimientos.EOF
            
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsCargarMovimientos.Fields("Fecha_Transa").Value
            
            If RsCargarMovimientos.Fields("Tipo_Mov").Value = 1 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Efe."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 2 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Che."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 3 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Dep."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 4 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Tran."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 5 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Cre."
            ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 6 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Deb."
            End If
             
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsCargarMovimientos.Fields("n_comprobante").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsCargarMovimientos.Fields("Detalle").Value
            
            If RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
            End If
            
            If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
            End If
             
            If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
            Else
                DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
            End If
            
            If RsCargarMovimientos.Fields("Conciliado").Value = True Then
                DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "SI"
            Else
                DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO"
            End If
            RsCargarMovimientos.MoveNext
        Loop
        
        DMGrid2.PaintMGrid
        TxtTotalMovimientos.Text = DMGrid2.Rows
        TxtSaldoDisponible.Text = Format(DMGrid2.ValorCelda(DMGrid2.Rows, 7), "#,##0.00")
    Else
        TxtTotalMovimientos.Text = 0
        TxtSaldoDisponible.Text = Format(0, "#,##0.00")
    End If
    
End If
    
End Sub

Private Sub DMGrid2_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)

If Button = vbLeftButton Then
    If DMGrid2.ValorCelda(lRow, 8) = "SI" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End If

End Sub


Private Sub DtpFechaDesde_Change()
TxtFechaDesde.Text = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaHasta_Change()
TxtFechaHasta.Text = Format(DtpFechaHasta.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Centrar Me
Grid1
Grid2
End Sub

Sub ChequesTransitos()
CSql = "SELECT Sum(Monto_Mov) as TotalMonto From Movi_BanCaja " & _
       "WHERE (Conciliado = '0' AND Tipo_Mov = '1') And IdCajaBanco = '" & TxtCodigo.Text & "'"
Set RsChequesTransitos = CrearRS(CSql)

If RsChequesTransitos.RecordCount > 0 Then
    TxtChequesTransitos.Text = Format(RsChequesTransitos.Fields("TotalMonto").Value, "#,##0.00")
Else
    TxtChequesTransitos.Text = Format(0, "#,##0.00")
End If

End Sub

Sub TotalConciliado()

CSql = "Select Sum(Monto_Mov) as TotalConci From Movi_BanCaja Where Conciliado='1' And IdCajaBanco = '" & TxtCodigo.Text & "'"
Set RsTotalConci = CrearRS(CSql)

If RsTotalConci.RecordCount > 0 Then
    TxtConciliado.Text = Format(RsTotalConci.Fields("TotalConci").Value, "#,##0.00")
Else
    TxtConciliado.Text = Format(0, "#,##0.00")
End If
RsTotalConci.Close

End Sub

Sub TotalNoConciliado()

CSql = "Select Sum(Monto_Mov) as TotalNoConci From Movi_BanCaja Where Conciliado='0' And IdCajaBanco = '" & TxtCodigo.Text & "'"
Set RsTotalNoConci = CrearRS(CSql)

If RsTotalNoConci.RecordCount > 0 Then
    TxtNoConciliado.Text = Format(RsTotalNoConci.Fields("TotalNoConci").Value, "#,##0.00")
Else
    TxtNoConciliado.Text = Format(0, "#,##0.00")
End If

RsTotalNoConci.Close

End Sub

Sub Grid1()
DMGrid1.Rows = 1
DMGrid1.Cols = 2
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 1
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(1).Width = 1200
DMGrid1.DColumnas(2).Width = 1200
DMGrid1.DColumnas(1).Caption = "Cuenta Banco"
DMGrid1.DColumnas(2).Caption = "Fecha"

End Sub
Sub Grid2()
DMGrid2.Rows = 0
DMGrid2.Cols = 8
DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 0
DMGrid2.DColumnas(4).Alignment = 0
DMGrid2.DColumnas(5).Alignment = 1
DMGrid2.DColumnas(6).Alignment = 1
DMGrid2.DColumnas(7).Alignment = 1
DMGrid2.DColumnas(8).Alignment = 0

DMGrid2.DColumnas(1).Locked = True
DMGrid2.DColumnas(2).Locked = True
DMGrid2.DColumnas(3).Locked = True
DMGrid2.DColumnas(4).Locked = True
DMGrid2.DColumnas(5).Locked = True
DMGrid2.DColumnas(6).Locked = True
DMGrid2.DColumnas(7).Locked = True
DMGrid2.DColumnas(8).Locked = True

DMGrid2.DColumnas(1).Width = 1200
DMGrid2.DColumnas(2).Width = 1000
DMGrid2.DColumnas(3).Width = 1750
DMGrid2.DColumnas(4).Width = 3000
DMGrid2.DColumnas(5).Width = 1750
DMGrid2.DColumnas(6).Width = 1750
DMGrid2.DColumnas(7).Width = 1750
DMGrid2.DColumnas(8).Width = 1000

DMGrid2.DColumnas(1).Caption = "Fecha"
DMGrid2.DColumnas(2).Caption = "Tipo Doc."
DMGrid2.DColumnas(3).Caption = "Número Transacción"
DMGrid2.DColumnas(4).Caption = "Detalle"
DMGrid2.DColumnas(5).Caption = "Cargos"
DMGrid2.DColumnas(6).Caption = "Abonos"
DMGrid2.DColumnas(7).Caption = "Saldo"
DMGrid2.DColumnas(8).Caption = "Conciliada"

End Sub

Sub Movi()
Grid2
'CSql = "Select * From EstadoDeCuenta Where IdCajaBanco = '" & TxtCodigo.Text & "' Order by Fecha_Transa, Tipo_Mov"

CSql = "Select * From Movi_BanCaja Where IdCajaBanco = '" & TxtCodigo.Text & "' Order by Fecha_Transa, Tipo_Mov"
Set RsCargarMovimientos = CrearRS(CSql)


'DMGrid2.Cols = 8
DMGrid2.Rows = 0
DMGrid2.Clear
If RsCargarMovimientos.RecordCount = 0 Then Exit Sub
If RsCargarMovimientos.RecordCount > 0 Then
    Do While Not RsCargarMovimientos.EOF
        
        DMGrid2.Rows = DMGrid2.Rows + 1
        DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsCargarMovimientos.Fields("Fecha_Transa").Value
        
        If RsCargarMovimientos.Fields("Tipo_Mov").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Efe."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Che."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 3 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Dep."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 4 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Tran."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 5 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Cre."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 6 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Deb."
        End If
         
        DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsCargarMovimientos.Fields("n_comprobante").Value
        DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsCargarMovimientos.Fields("Detalle").Value
        
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
        
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
         
        If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        End If
        
        If RsCargarMovimientos.Fields("Conciliado").Value = True Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "SI"
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO"
        End If
        RsCargarMovimientos.MoveNext
    Loop
    
    DMGrid2.PaintMGrid
    TxtTotalMovimientos.Text = DMGrid2.Rows
    TxtSaldoDisponible.Text = Format(DMGrid2.ValorCelda(DMGrid2.Rows, 7), "#,##0.00")
Else
    TxtTotalMovimientos.Text = 0
    TxtSaldoDisponible.Text = Format(0, "#,##0.00")
End If


End Sub

Sub Movi1()
Grid2
'CSql = "Select * From EstadoDeCuenta Where Conciliado=1 And IdCajaBanco='" & Val(TxtCodigo.Text) & "'"
CSql = "Select * From Movi_BanCaja Where Conciliado=1 And IdCajaBanco='" & Val(TxtCodigo.Text) & "'"

Set RsCargarMovimientos = CrearRS(CSql)
'DMGrid2.Cols = 8
DMGrid2.Rows = 0
DMGrid2.Clear
If RsCargarMovimientos.RecordCount = 0 Then Exit Sub
If RsCargarMovimientos.RecordCount > 0 Then
    Do While Not RsCargarMovimientos.EOF
        
        DMGrid2.Rows = DMGrid2.Rows + 1
        DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsCargarMovimientos.Fields("Fecha_Transa").Value
        
        If RsCargarMovimientos.Fields("Tipo_Mov").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Efe."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Che."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 3 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Dep."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 4 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Tran."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 5 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Cre."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 6 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Deb."
        End If
         
        DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsCargarMovimientos.Fields("n_comprobante").Value
        DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsCargarMovimientos.Fields("Detalle").Value
        
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
        
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
         
        If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        End If
        
        If RsCargarMovimientos.Fields("Conciliado").Value = True Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "SI"
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO"
        End If
        RsCargarMovimientos.MoveNext
    Loop
    
    DMGrid2.PaintMGrid
    TxtTotalMovimientos.Text = DMGrid2.Rows
    TxtSaldoDisponible.Text = Format(DMGrid2.ValorCelda(DMGrid2.Rows, 7), "#,##0.00")
Else
    TxtTotalMovimientos.Text = 0
    TxtSaldoDisponible.Text = Format(0, "#,##0.00")
End If
End Sub

Sub Movi2()
Grid2
CSql = "Select * From Movi_BanCaja Where Conciliado=0 And IdCajaBanco='" & TxtCodigo.Text & "'"
Set RsCargarMovimientos = CrearRS(CSql)
DMGrid2.Clear
DMGrid2.Rows = 0
'DMGrid2.Cols = 8
If RsCargarMovimientos.RecordCount = 0 Then Exit Sub
If RsCargarMovimientos.RecordCount > 0 Then
    Do While Not RsCargarMovimientos.EOF
        
        DMGrid2.Rows = DMGrid2.Rows + 1
        DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsCargarMovimientos.Fields("Fecha_Transa").Value
        
        If RsCargarMovimientos.Fields("Tipo_Mov").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Efe."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Che."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 3 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Dep."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 4 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Tran."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 5 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Cre."
        ElseIf RsCargarMovimientos.Fields("Tipo_Mov").Value = 6 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Deb."
        End If
         
        DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsCargarMovimientos.Fields("n_comprobante").Value
        DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsCargarMovimientos.Fields("Detalle").Value
        
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
        
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
         
        If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        End If
        
        If RsCargarMovimientos.Fields("Conciliado").Value = True Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "SI"
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = "NO"
        End If
        RsCargarMovimientos.MoveNext
    Loop
    
    DMGrid2.PaintMGrid
    TxtTotalMovimientos.Text = DMGrid2.Rows
    TxtSaldoDisponible.Text = Format(DMGrid2.ValorCelda(DMGrid2.Rows, 7), "#,##0.00")
Else
    TxtTotalMovimientos.Text = 0
    TxtSaldoDisponible.Text = Format(0, "#,##0.00")
End If
End Sub
Sub Conciliacion()
Grid1
CSql = "SELECT DISTINCT IdCajaBanco, FechaConciliacion From Movi_BanCaja Where Conciliado=1 And IdCajaBanco='" & Val(TxtCodigo.Text) & "'"
Set RsCargarMovimientos = CrearRS(CSql)
DMGrid1.Clear
DMGrid1.Rows = 0
'DMGrid2.Cols = 8
If RsCargarMovimientos.RecordCount > 0 Then
    Do While Not RsCargarMovimientos.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarMovimientos.Fields("IdCajaBanco").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarMovimientos.Fields("FechaConciliacion").Value
        RsCargarMovimientos.MoveNext
    Loop

    DMGrid1.PaintMGrid
    TxtTotalConciliaciones = DMGrid1.Rows
Else
    DMGrid1.Rows = 0
    DMGrid1.Clear
    DMGrid1.PaintMGrid
    TxtTotalConciliaciones = 0
End If
End Sub

Private Sub MnuConciliar_Click()
If FrmConcilicacionBancaria.TxtCodigo.Text <> "" Then
    FrmConciliacion.Show vbModal
Else
    MsgBox "Dede de Seleccionar un Banco para poder realizar la conciliación!", vbCritical + vbOKOnly, "Error"
End If
End Sub

Private Sub MnuEstadoCuenta_Click()
FrmConciliacionEstadoCuentas.Show vbModal
End Sub

Private Sub OptDiferencias_Click()
If OptDiferencias.Value = True Then
    If TxtCodigo.Text <> "" Then
        Grid2
        'Movi3
        Conciliacion
    End If
End If
End Sub

Private Sub OptTodos_Click()
If OptTodos.Value = True Then
    If TxtCodigo.Text <> "" Then
        Grid2
        Movi
        Conciliacion
        TxtFechaDesde.Text = ""
        TxtFechaHasta.Text = ""
    End If
End If
End Sub

Private Sub OptConciliados_Click()
If OptConciliados.Value = True Then
    If TxtCodigo.Text <> "" Then
        Grid2
        Movi1
        Conciliacion
    End If
End If
End Sub

Private Sub OptNoConciliados_Click()
If OptNoConciliados.Value = True Then
    If TxtCodigo.Text <> "" Then
        Grid2
        Movi2
        Conciliacion
    End If
End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmLibroBancos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de  Bancos"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15315
   Icon            =   "FrmLibroBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   15315
   Begin VB.Frame FrameBusquedaAvanzada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Busqueda Avanzada"
      Height          =   1935
      Left            =   4200
      TabIndex        =   41
      Top             =   3480
      Visible         =   0   'False
      Width           =   7095
      Begin VB.ComboBox CboTipo 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TxtFechaOperacion 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtN_Comprobante 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TxtMonto 
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DtpFechaOperacion 
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   40464
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar2 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Realiza una Busqueda Avanzada"
         Top             =   1440
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
         MICON           =   "FrmLibroBancos.frx":1002
         PICN            =   "FrmLibroBancos.frx":101E
         PICH            =   "FrmLibroBancos.frx":1283
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar2 
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         ToolTipText     =   "Cerrar"
         Top             =   1440
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
         MICON           =   "FrmLibroBancos.frx":1515
         PICN            =   "FrmLibroBancos.frx":1531
         PICH            =   "FrmLibroBancos.frx":16FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Operación:"
         Height          =   195
         Left            =   3600
         TabIndex        =   45
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Comprobante:"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   450
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   3600
         TabIndex        =   43
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Operación:"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   930
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   8760
         Width           =   8055
         Begin VB.TextBox TxtFechaDesde 
            Height          =   375
            Left            =   1200
            TabIndex        =   32
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox TxtFechaHasta 
            Height          =   375
            Left            =   3840
            TabIndex        =   30
            Top             =   270
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DtpFechaDesde 
            Height          =   375
            Left            =   2400
            TabIndex        =   31
            Top             =   270
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   40242
         End
         Begin MSComCtl2.DTPicker DtpFechaHasta 
            Height          =   375
            Left            =   5040
            TabIndex        =   33
            Top             =   270
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   40242
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarFiltro 
            Height          =   375
            Left            =   5400
            TabIndex        =   34
            ToolTipText     =   "Buscar"
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmLibroBancos.frx":192F
            PICN            =   "FrmLibroBancos.frx":194B
            PICH            =   "FrmLibroBancos.frx":1BB0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBusquedaAvanzada 
            Height          =   375
            Left            =   6720
            TabIndex        =   40
            ToolTipText     =   "Realiza una Busqueda Avanzada"
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Avanzada"
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
            MICON           =   "FrmLibroBancos.frx":1E42
            PICN            =   "FrmLibroBancos.frx":1E5E
            PICH            =   "FrmLibroBancos.frx":20C3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   2760
            TabIndex        =   35
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.TextBox TxtTotalMovimientos 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0"
         Top             =   8280
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información Bancaria:"
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   9975
         Begin VB.TextBox txtCodigo 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox TxtNombreBanco 
            Height          =   375
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox TxtNoCuenta 
            Height          =   375
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   2775
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   8640
            TabIndex        =   19
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
            MICON           =   "FrmLibroBancos.frx":2355
            PICN            =   "FrmLibroBancos.frx":2371
            PICH            =   "FrmLibroBancos.frx":25D6
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
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            Height          =   195
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Cuenta Bancaria"
            Height          =   195
            Left            =   5760
            TabIndex        =   20
            Top             =   240
            Width           =   1710
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Movimientos Bancarios Disponibles"
         Height          =   6975
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   14895
         Begin SystemOncoAmerica.DMGrid DMGrid2 
            Height          =   6615
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   11668
            Object.Width           =   14625
            Object.Height          =   6585
            Cols            =   7
            Rows            =   0
            ScrollBar       =   1
            MarqueeStyle    =   2
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   8280
         TabIndex        =   9
         Top             =   8760
         Width           =   6735
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   5640
            TabIndex        =   10
            ToolTipText     =   "Cerrar "
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
            MICON           =   "FrmLibroBancos.frx":2868
            PICN            =   "FrmLibroBancos.frx":2884
            PICH            =   "FrmLibroBancos.frx":2A4D
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
            Left            =   4440
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
            MICON           =   "FrmLibroBancos.frx":2C82
            PICN            =   "FrmLibroBancos.frx":2C9E
            PICH            =   "FrmLibroBancos.frx":2F80
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
            Left            =   1200
            TabIndex        =   12
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
            MICON           =   "FrmLibroBancos.frx":31D1
            PICN            =   "FrmLibroBancos.frx":31ED
            PICH            =   "FrmLibroBancos.frx":3312
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnEliminar 
            Height          =   375
            Left            =   120
            TabIndex        =   37
            ToolTipText     =   "Eliminar Usuario"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "FrmLibroBancos.frx":35A2
            PICN            =   "FrmLibroBancos.frx":35BE
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Height          =   975
         Left            =   10200
         TabIndex        =   8
         Top             =   240
         Width           =   4815
         Begin VB.TextBox TxtSaldoDisponible 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0,00"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox TxtSaldoLibro 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Disponible:"
            Height          =   195
            Left            =   1515
            TabIndex        =   26
            Top             =   653
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo en Libro:"
            Height          =   195
            Left            =   1680
            TabIndex        =   24
            Top             =   300
            Width           =   1065
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnEditorFacturas 
         Height          =   375
         Left            =   13080
         TabIndex        =   38
         Top             =   8280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Editor Facturas"
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
         MICON           =   "FrmLibroBancos.frx":3762
         PICN            =   "FrmLibroBancos.frx":377E
         PICH            =   "FrmLibroBancos.frx":3A1D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnReporteOperacion 
         Height          =   375
         Left            =   10920
         TabIndex        =   39
         ToolTipText     =   "Reporte"
         Top             =   8280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Reporte Operanción"
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
         MICON           =   "FrmLibroBancos.frx":3E52
         PICN            =   "FrmLibroBancos.frx":3E6E
         PICH            =   "FrmLibroBancos.frx":3F93
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   10440
         Top             =   8280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Libro de Bancos"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Movimientos:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   8370
         Width           =   1560
      End
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuBanco 
         Caption         =   "Bancos"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBeneficiario 
         Caption         =   "Beneficiarios"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConceptos 
         Caption         =   "Conceptos"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu MnuMovimientos 
      Caption         =   "Movimientos"
      Begin VB.Menu SubMnuCheques 
         Caption         =   "Cheques"
      End
      Begin VB.Menu SubMnuDepositos 
         Caption         =   "Depósitos"
      End
      Begin VB.Menu se1 
         Caption         =   "-"
      End
      Begin VB.Menu SubMnuNotasCreditos 
         Caption         =   "Notas de Créditos"
      End
      Begin VB.Menu SubMnuNotasDebito 
         Caption         =   "Notas de Débitos"
      End
      Begin VB.Menu se2 
         Caption         =   "-"
      End
      Begin VB.Menu SubMnuTransferenciaFondos 
         Caption         =   "Transferencia de Fondos"
      End
   End
End
Attribute VB_Name = "FrmLibroBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarMovimientos As New ADODB.Recordset
Dim RsTotalConci As New ADODB.Recordset

Private Sub BtnBuscar_Click()
Ban = 1
FrmListadoBancos.Show vbModal, FrmPrincipal
Movi
End Sub

Private Sub BtnBuscar2_Click()
On Error GoTo Most
Dim RsBuscarCliente As New ADODB.Recordset
Dim wer As String

wher = ""

If TxtMonto.Text = "" And TxtN_Comprobante.Text = "" And TxtFechaOperacion.Text = "" And CboTipo.ListIndex = -1 Then
    Msg = "Por favor ingrese el Número de Comprobante o el monto o seleccione la fecha de Operación " & Chr(13) & "o tipo de operación para realizar la busqueda!!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
Else

    If TxtN_Comprobante.Text <> "" Then
        wer = wer & " N_Comprobante = '" & TxtN_Comprobante.Text & "'"
    End If

    If CboTipo.ListIndex <> -1 Then
        If wer = "" Then
            wer = wer & " Tipo_Mov = '" & CboTipo.ItemData(CboTipo.ListIndex) & "'"
        Else
            wer = wer & " And Tipo_Mov = '" & CboTipo.ItemData(CboTipo.ListIndex) & "'"
        End If
    End If

    If TxtFechaOperacion.Text <> "" Then
        DtpFechaOperacion.Value = TxtFechaOperacion.Text
        If wer = "" Then
            wer = wer & " Fecha_Transa= '" & TxtFechaOperacion.Text & "'"
        Else
            wer = wer & " And Fecha_Transa = '" & TxtFechaOperacion.Text & "'"
        End If
    End If

    If TxtMonto.Text <> "" Then
        If wer = "" Then
            wer = wer & " Monto_Mov like '%" & TxtMonto.Text & "%'"
        Else
            wer = wer & " And Monto_Mov like '%" & TxtMonto.Text & "%'"
        End If
    End If
   
End If

Grid2
    CSql = "Select * From Movi_BanCaja Where " & wer & " And IdCajaBanco = '" & txtCodigo.Text & "' Order by Fecha_Transa asc"
    Set RsCargarMovimientos = CrearRS(CSql)
    DMGrid2.Rows = 0
    DMGrid2.Clear
    If RsCargarMovimientos.RecordCount = 0 Then
        TxtTotalMovimientos.Text = 0
        TxtSaldoLibro.Text = Format(0, "#,##0.00")
        TxtSaldoDisponible.Text = Format(0, "#,##0.00")
        Exit Sub
    End If
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
         
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        ElseIf RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
         
        If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        End If
         
        RsCargarMovimientos.MoveNext
    Loop
    
    DMGrid2.PaintMGrid
    
    TxtTotalMovimientos.Text = DMGrid2.Rows
    TxtSaldoLibro.Text = DMGrid2.ValorCelda(DMGrid2.Rows, 7)
    TxtSaldoDisponible.Text = DMGrid2.ValorCelda(DMGrid2.Rows, 7)



FrameBusquedaAvanzada.Visible = False
Frame1.Enabled = True
Limpiar

Exit Sub

Most:
    MsgBox "La fecha que ha ingresado no es valida!", vbInformation + vbOKOnly, "Informacion"
    TxtFechaOperacion.Text = ""

End Sub

Sub Limpiar()
TxtN_Comprobante.Text = ""
CboTipo.ListIndex = -1
TxtMonto.Text = ""
TxtFechaOperacion.Text = ""
End Sub

Private Sub BtnBuscarFiltro_Click()

If txtCodigo.Text = "" Then
    MsgBox "Seleccione un banco para poder realizar la busqueda!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If
If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
    Grid2
    CSql = "Select * From Movi_BanCaja Where IdCajaBanco = '" & txtCodigo.Text & "' And Fecha_Transa >='" & Trim(TxtFechaDesde.Text) & "' and Fecha_Transa <='" & Trim(TxtFechaHasta.Text) & "' Order by Fecha_Transa asc"
    Set RsCargarMovimientos = CrearRS(CSql)
    DMGrid2.Rows = 0
    DMGrid2.Clear
    If RsCargarMovimientos.RecordCount = 0 Then Exit Sub
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
         
        If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        ElseIf RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
        End If
         
        If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        Else
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
        End If
         
        RsCargarMovimientos.MoveNext
    Loop
    
    DMGrid2.PaintMGrid
    
    TxtTotalMovimientos.Text = DMGrid2.Rows
    TxtSaldoLibro.Text = DMGrid2.ValorCelda(DMGrid2.Rows, 7)
    TxtSaldoDisponible.Text = DMGrid2.ValorCelda(DMGrid2.Rows, 7)
Else
    Grid2
    Movi
End If
End Sub

Private Sub BtnBusquedaAvanzada_Click()
If txtCodigo.Text = "" Then
    MsgBox "Seleccione un banco para poder realizar la busqueda!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

FrameBusquedaAvanzada.Visible = True
Frame1.Enabled = False
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnCerrar2_Click()
FrameBusquedaAvanzada.Visible = False
Frame1.Enabled = True
End Sub

Private Sub BtnDesHacer_Click()
Movi
TxtFechaDesde.Text = ""
TxtFechaHasta.Text = ""
End Sub

Private Sub BtnEditorFacturas_Click()
Editar = 1
    IdReg = DMGrid2.ValorCelda(DMGrid2.Row, 8)
    If DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Dep." Then
            FrmTransaccionDepositos.Show vbModal, FrmPrincipal
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
            FrmTransaccionCheques.Show vbModal, FrmPrincipal
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Dep." Then
            FrmTransaccionDepositos.Show vbModal, FrmPrincipal
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
    
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
    
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
    
    End If
End Sub

Private Sub BtnEliminar_Click()

If DMGrid2.Rows = 0 Then Exit Sub


mensaje = MsgBox("Estas seguro de eliminar el asiento registrado?", vbYesNo + vbInformation, "Mensaje")

If mensaje = vbYes Then
    CSql = "Delete From Movi_BanCaja Where IdMovCajaBanco='" & IdReg & "'"
    Set RsTemp = CrearRS(CSql)
End If

Grid2
Movi

MsgBox "Registro eliminado satisfactoriamente", vbOKOnly + vbInformation, "Operación Exitosa"

End Sub

Private Sub BtnImprimir_Click()
If txtCodigo.Text <> "" Then
    
    If TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
        If DMGrid2.Rows > 0 Then
            With CrystalReport1
                .ReportFileName = RutaInformes & "\LibroBancos.rpt"
                .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
                .DiscardSavedData = True
                .RetrieveDataFiles
                .ReportSource = 0
                .SelectionFormula = "{MoviBancoCaja.IdCajaBanco} = " & txtCodigo.Text
                '.WindowTitle = "Reporte Historia Medica No. " & Label12.Caption
                .Destination = crptToWindow
                .PrintFileType = crptCrystal
                .WindowState = crptMaximized
                .WindowMaxButton = False
                .WindowMinButton = False
                .Action = 1
            End With
        Else
            MsgBox "El Banco seleccionado NO posee movimientos bancarios para imprimir!", vbOKOnly + vbCritical, "Error"
        End If
    Else
        If DMGrid2.Rows > 0 Then
            With CrystalReport1
                .ReportFileName = RutaInformes & "\LibroBancos.rpt"
                .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
                .DiscardSavedData = True
                .RetrieveDataFiles
                .ReportSource = 0
                .SelectionFormula = "{MoviBancoCaja.IdCajaBanco} = " & txtCodigo.Text & " And {MoviBancoCaja.Fecha_Transa} = " & FechaSQL(TxtFechaDesde.Text) & " And {MoviBancoCaja.Fecha_Transa} = " & FechaSQL(TxtFechaHasta.Text) & ""
                '.ReportTitle = "Reporte Historia Medica No. " & Label12.Caption
                .Destination = crptToWindow
                .PrintFileType = crptCrystal
                .WindowState = crptMaximized
                .WindowMaxButton = False
                .WindowMinButton = False
                .Action = 1
            End With
        Else
            MsgBox "El Banco seleccionado NO posee movimientos bancarios para imprimir entre las fechas " & TxtFechaDesde.Text & " Al " & TxtFechaHasta.Text & "!", vbOKOnly + vbCritical, "Error"
        End If
    End If
Else
    MsgBox "Seleccione un Banco para poder ver los movimientos bancarios!", vbOKOnly + vbCritical, "Error"
End If
End Sub

Private Sub BtnSiguiente_Click()

End Sub

Private Sub BtnReporteOperacion_Click()
FrmOperacionesLibroBancos.Show vbModal, FrmPrincipal
End Sub

Private Sub DMGrid2_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton Then
   IdReg = DMGrid2.ValorCelda(DMGrid2.Row, 8)
End If

If Button = vbRightButton Then
    Editar = 1
    IdReg = DMGrid2.ValorCelda(DMGrid2.Row, 8)
    If DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Dep." Then
            FrmTransaccionDepositos.Show vbModal, FrmPrincipal
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
            FrmTransaccionCheques.Show vbModal, FrmPrincipal
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Dep." Then
            FrmTransaccionDepositos.Show vbModal, FrmPrincipal
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
    
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
    
    ElseIf DMGrid2.ValorCelda(DMGrid2.Row, 2) = "Che." Then
    
    End If
End If
End Sub

Private Sub DtpFechaDesde_Change()
TxtFechaDesde.Text = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaHasta_Change()
TxtFechaHasta.Text = Format(DtpFechaHasta.Value, "dd/mm/yyyy")
End Sub


Private Sub DtpFechaOperacion_Change()
TxtFechaOperacion.Text = Format(DtpFechaOperacion.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Centrar Me
Grid2
Movi
DtpFechaOperacion.Value = DateTime.Date

CboTipo.AddItem "Depósito"
CboTipo.ItemData(CboTipo.NewIndex) = 1

CboTipo.AddItem "Cheque"
CboTipo.ItemData(CboTipo.NewIndex) = 2

CboTipo.AddItem "Notas de Crédito"
CboTipo.ItemData(CboTipo.NewIndex) = 4

CboTipo.AddItem "Notas de Dédito"
CboTipo.ItemData(CboTipo.NewIndex) = 5

CboTipo.AddItem "Transferencia de Fondos"
CboTipo.ItemData(CboTipo.NewIndex) = 6

End Sub

Sub Grid2()
DMGrid2.Rows = 1
DMGrid2.Cols = 8
DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 0
DMGrid2.DColumnas(4).Alignment = 0
DMGrid2.DColumnas(5).Alignment = 1
DMGrid2.DColumnas(6).Alignment = 1
DMGrid2.DColumnas(7).Alignment = 1
DMGrid2.DColumnas(8).Alignment = 1

DMGrid2.DColumnas(1).Locked = True
DMGrid2.DColumnas(2).Locked = True

DMGrid2.DColumnas(1).Width = 1200
DMGrid2.DColumnas(2).Width = 750
DMGrid2.DColumnas(3).Width = 1250
DMGrid2.DColumnas(4).Width = 6500
DMGrid2.DColumnas(5).Width = 1500
DMGrid2.DColumnas(6).Width = 1500
DMGrid2.DColumnas(7).Width = 1500
DMGrid2.DColumnas(8).Width = 5

DMGrid2.DColumnas(1).Caption = "Fecha"
DMGrid2.DColumnas(2).Caption = "Tipo"
DMGrid2.DColumnas(3).Caption = "Documento"
DMGrid2.DColumnas(4).Caption = "Descripción"
DMGrid2.DColumnas(5).Caption = "Debe"
DMGrid2.DColumnas(6).Caption = "Haber"
DMGrid2.DColumnas(7).Caption = "Saldo"
DMGrid2.DColumnas(8).Caption = ""

End Sub

Sub Movi()
Grid2
CSql = "Select * From Movi_BanCaja Where IdCajaBanco = '" & txtCodigo.Text & "' Order by Fecha_Transa asc"
Set RsCargarMovimientos = CrearRS(CSql)
DMGrid2.Rows = 0
DMGrid2.Clear
If RsCargarMovimientos.RecordCount = 0 Then
    TxtTotalMovimientos.Text = 0
    TxtSaldoLibro.Text = "0,00"
    TxtSaldoDisponible.Text = "0,00"
    Exit Sub
End If
Do While Not RsCargarMovimientos.EOF
    
    DMGrid2.Rows = DMGrid2.Rows + 1
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsCargarMovimientos.Fields("Fecha_Transa").Value
    
    If RsCargarMovimientos.Fields("Tipo_Mov").Value = 1 Then
        DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "Dep."
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
     
    If RsCargarMovimientos.Fields("Ingr_Egr").Value = 2 Then
        DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
    ElseIf RsCargarMovimientos.Fields("Ingr_Egr").Value = 1 Then
        DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(RsCargarMovimientos.Fields("Monto_Mov").Value, "#,##0.00")
    End If
     
    If IsEmpty(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) Then
        DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
    Else
        DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows - 1, 7)) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)) - CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00")
    End If
    DMGrid2.ValorCelda(DMGrid2.Rows, 8) = RsCargarMovimientos.Fields("IdMovCajaBanco").Value
    RsCargarMovimientos.MoveNext
Loop

DMGrid2.PaintMGrid

TxtTotalMovimientos.Text = DMGrid2.Rows
TxtSaldoLibro.Text = DMGrid2.ValorCelda(DMGrid2.Rows, 7)
TxtSaldoDisponible.Text = DMGrid2.ValorCelda(DMGrid2.Rows, 7)
End Sub

Private Sub MnuBanco_Click()
FrmCajasBancos.Show vbModal, FrmPrincipal
End Sub

Private Sub MnuBeneficiario_Click()
FrmBeneficiario.Show vbModal, FrmPrincipal
End Sub

Private Sub MnuCerrar_Click()
Unload Me
End Sub

Private Sub MnuConceptos_Click()
FrmConceptosBancos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCheques_Click()
FrmTransaccionCheques.Show vbModal
End Sub

Private Sub SubMnuDepositos_Click()
FrmTransaccionDepositos.Show vbModal
End Sub

Private Sub SubMnuNotasCreditos_Click()
FrmTransaccionNotaCredito.Show vbModal
End Sub

Private Sub SubMnuNotasDebito_Click()
FrmTransaccionNotasDebitos.Show vbModal
End Sub

Private Sub SubMnuTransferenciaFondos_Click()
FrmTransferenciaFondos.Show vbModal
End Sub

Private Sub TxtFechaOperacion_LostFocus()
On Error GoTo Most

DtpFechaOperacion.Value = TxtFechaOperacion.Text

Exit Sub

Most:
    MsgBox "La fecha que ha ingresado no es valida!", vbInformation + vbOKOnly, "Informacion"
    TxtFechaOperacion.Text = ""
    
End Sub

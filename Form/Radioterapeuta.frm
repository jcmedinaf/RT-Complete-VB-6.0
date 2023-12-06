VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmRadioTerapeuta 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oncología"
   ClientHeight    =   10875
   ClientLeft      =   3930
   ClientTop       =   1050
   ClientWidth     =   12210
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "Radioterapeuta.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   12210
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Paciente"
      Height          =   2655
      Left            =   120
      TabIndex        =   72
      Top             =   120
      Width           =   12015
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3360
         Top             =   240
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   79
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   78
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1680
         TabIndex        =   77
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5760
         TabIndex        =   76
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1680
         TabIndex        =   75
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1680
         TabIndex        =   74
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   5760
         TabIndex        =   73
         Top             =   720
         Width           =   1095
      End
      Begin ChamaleonButton.ChameleonBtn BtnLlamar 
         Height          =   375
         Left            =   5880
         TabIndex        =   80
         ToolTipText     =   "Llamar"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Llamar"
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
         MICON           =   "Radioterapeuta.frx":1002
         PICN            =   "Radioterapeuta.frx":101E
         PICH            =   "Radioterapeuta.frx":12BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnListaEspera 
         Height          =   375
         Left            =   4080
         TabIndex        =   81
         ToolTipText     =   "Lista de Espera"
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Lista de Espera"
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
         MICON           =   "Radioterapeuta.frx":14EF
         PICN            =   "Radioterapeuta.frx":150B
         PICH            =   "Radioterapeuta.frx":1794
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFechaRegistro 
         Height          =   375
         Left            =   8400
         TabIndex        =   82
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48037889
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaNac 
         Height          =   375
         Left            =   8400
         TabIndex        =   83
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48037889
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   375
         Left            =   8400
         TabIndex        =   84
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48037889
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   375
         Left            =   8400
         TabIndex        =   85
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48037889
         CurrentDate     =   40121
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
         Height          =   375
         Left            =   11040
         TabIndex        =   86
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   2160
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
         MICON           =   "Radioterapeuta.frx":1A2C
         PICN            =   "Radioterapeuta.frx":1A48
         PICH            =   "Radioterapeuta.frx":1CDE
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
         Left            =   10200
         TabIndex        =   87
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   2160
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
         MICON           =   "Radioterapeuta.frx":1F3D
         PICN            =   "Radioterapeuta.frx":1F59
         PICH            =   "Radioterapeuta.frx":21EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnDesocuparAlPacienteAtendido 
         Height          =   375
         Left            =   5880
         TabIndex        =   88
         ToolTipText     =   "Desocupar al Paciente Atendido"
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Desocupar al Paciente"
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
         MICON           =   "Radioterapeuta.frx":244A
         PICN            =   "Radioterapeuta.frx":2466
         PICH            =   "Radioterapeuta.frx":260A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnTipoCancer 
         Height          =   375
         Left            =   4080
         TabIndex        =   89
         ToolTipText     =   "Lista de Espera"
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Tipo de Cancer"
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
         MICON           =   "Radioterapeuta.frx":283F
         PICN            =   "Radioterapeuta.frx":285B
         PICH            =   "Radioterapeuta.frx":2AE4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label NoReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro "
         Height          =   195
         Left            =   8280
         TabIndex        =   103
         Top             =   2250
         Width           =   630
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Historia:"
         Height          =   195
         Left            =   4200
         TabIndex        =   102
         Top             =   330
         Width           =   870
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   9960
         Picture         =   "Radioterapeuta.frx":2D7C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   5160
         TabIndex        =   101
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula:"
         Height          =   195
         Left            =   945
         TabIndex        =   100
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro:"
         Height          =   255
         Left            =   6840
         TabIndex        =   99
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre(s):"
         Height          =   195
         Left            =   720
         TabIndex        =   98
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellido(s):"
         Height          =   195
         Left            =   720
         TabIndex        =   97
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Tratante:"
         Height          =   195
         Left            =   270
         TabIndex        =   96
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Inicio:"
         Height          =   195
         Left            =   7200
         TabIndex        =   95
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de C&ulminación:"
         Height          =   375
         Left            =   7320
         TabIndex        =   94
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Remitente:"
         Height          =   195
         Left            =   150
         TabIndex        =   93
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nac.:"
         Height          =   195
         Left            =   7200
         TabIndex        =   92
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad:"
         Height          =   195
         Left            =   5280
         TabIndex        =   91
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         Height          =   195
         Left            =   5280
         TabIndex        =   90
         Top             =   810
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   10080
      Width           =   3615
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
         TabIndex        =   8
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad o Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Buscar"
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
         MICON           =   "Radioterapeuta.frx":3D79
         PICN            =   "Radioterapeuta.frx":3D95
         PICH            =   "Radioterapeuta.frx":3FFA
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   10080
      Width           =   8295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7200
         TabIndex        =   1
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
         MICON           =   "Radioterapeuta.frx":428C
         PICN            =   "Radioterapeuta.frx":42A8
         PICH            =   "Radioterapeuta.frx":4471
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
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Guardar / Actualizar "
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
         MICON           =   "Radioterapeuta.frx":46A6
         PICN            =   "Radioterapeuta.frx":46C2
         PICH            =   "Radioterapeuta.frx":4951
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
         TabIndex        =   3
         ToolTipText     =   "Agregar "
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
         MICON           =   "Radioterapeuta.frx":4D92
         PICN            =   "Radioterapeuta.frx":4DAE
         PICH            =   "Radioterapeuta.frx":4F3B
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
         Left            =   6000
         TabIndex        =   4
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
         MICON           =   "Radioterapeuta.frx":5170
         PICN            =   "Radioterapeuta.frx":518C
         PICH            =   "Radioterapeuta.frx":546E
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
         Left            =   2760
         TabIndex        =   5
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
         MICON           =   "Radioterapeuta.frx":56BF
         PICN            =   "Radioterapeuta.frx":56DB
         PICH            =   "Radioterapeuta.frx":587F
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
         Left            =   6600
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
      End
      Begin Crystal.CrystalReport CrystalReport2 
         Left            =   5280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Historia clinica"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   48037889
      CurrentDate     =   39801
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Complicaciones"
      Height          =   495
      Index           =   4
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   306
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Informe Final"
      Height          =   495
      Index           =   3
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "InmunoHistoquimica"
      Height          =   495
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Estadiaje / Seguimiento"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Informe Medico"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2880
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   12015
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Informe Médico"
         Height          =   5895
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   11775
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento:"
            Height          =   2175
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   3600
            Width           =   11535
            Begin VB.ComboBox CboMetas 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":5A1E
               Left            =   9600
               List            =   "Radioterapeuta.frx":5A2E
               TabIndex        =   45
               Top             =   1680
               Width           =   1815
            End
            Begin VB.ComboBox CboModificarMedicoTratante 
               Height          =   315
               Left            =   7080
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   1680
               Width           =   2415
            End
            Begin VB.TextBox Text17 
               Alignment       =   2  'Center
               Height          =   435
               Left            =   10680
               TabIndex        =   43
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":5A63
               Left            =   8880
               List            =   "Radioterapeuta.frx":5A6D
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   300
               Width           =   855
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   9000
               TabIndex        =   41
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   7800
               TabIndex        =   40
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Text22 
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Top             =   480
               Width           =   6855
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   10200
               TabIndex        =   38
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Metas:"
               Height          =   195
               Left            =   9600
               TabIndex        =   54
               Top             =   1440
               Width           =   480
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Médico Tratante:"
               Height          =   195
               Left            =   7080
               TabIndex        =   53
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "¿Cuantas?"
               Height          =   195
               Left            =   9840
               TabIndex        =   52
               Top             =   360
               Width           =   765
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Simulación Tomográfica:"
               Height          =   195
               Left            =   7080
               TabIndex        =   51
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Duración:"
               Height          =   195
               Left            =   7080
               TabIndex        =   50
               Top             =   1050
               Width           =   690
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis Diarias"
               Height          =   195
               Left            =   9000
               TabIndex        =   49
               Top             =   720
               Width           =   915
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total de Dosis"
               Height          =   195
               Left            =   7800
               TabIndex        =   48
               Top             =   720
               Width           =   1020
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción:"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   885
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sesiones"
               Height          =   195
               Left            =   10200
               TabIndex        =   46
               Top             =   720
               Width           =   645
            End
         End
         Begin VB.TextBox Text18 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Top             =   2760
            Width           =   5655
         End
         Begin VB.TextBox Text20 
            Height          =   735
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   66
            Top             =   1440
            Width           =   5775
         End
         Begin VB.TextBox Text16 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   65
            Top             =   1680
            Width           =   5655
         End
         Begin VB.TextBox Text19 
            Height          =   615
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   64
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox Text21 
            Height          =   765
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            Top             =   2730
            Width           =   5775
         End
         Begin VB.TextBox Text15 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   480
            Width           =   5655
         End
         Begin VB.ComboBox CboTCancers 
            Height          =   315
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   2340
            Width           =   4695
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Antecedentes Familiares:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1770
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Anatomía &Patológica:"
            Height          =   195
            Left            =   5880
            TabIndex        =   70
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dia&gnóstico:"
            Height          =   195
            Left            =   5880
            TabIndex        =   69
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nuevo Informe"
            Height          =   195
            Left            =   1920
            TabIndex        =   68
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "En&fermedad Actual:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   1395
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Motivo de la Consulta:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico:"
            Height          =   195
            Left            =   5880
            TabIndex        =   58
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Complicaciones"
         Height          =   5895
         Index           =   4
         Left            =   120
         TabIndex        =   307
         Top             =   240
         Width           =   11775
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   0
            Left            =   120
            TabIndex        =   340
            ToolTipText     =   "Complicaciones Agudas"
            Top             =   840
            Width           =   11535
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               ItemData        =   "Radioterapeuta.frx":5A79
               Left            =   3000
               List            =   "Radioterapeuta.frx":5A90
               Style           =   2  'Dropdown List
               TabIndex        =   380
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   6
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   365
               Top             =   3000
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   364
               Top             =   2640
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               ItemData        =   "Radioterapeuta.frx":5AC3
               Left            =   3000
               List            =   "Radioterapeuta.frx":5ADA
               Style           =   2  'Dropdown List
               TabIndex        =   363
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   362
               Top             =   2280
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               ItemData        =   "Radioterapeuta.frx":5B0D
               Left            =   3000
               List            =   "Radioterapeuta.frx":5B24
               Style           =   2  'Dropdown List
               TabIndex        =   361
               Top             =   2280
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   360
               Top             =   1920
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               ItemData        =   "Radioterapeuta.frx":5B57
               Left            =   3000
               List            =   "Radioterapeuta.frx":5B6E
               Style           =   2  'Dropdown List
               TabIndex        =   359
               Top             =   1920
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   358
               Top             =   1560
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               ItemData        =   "Radioterapeuta.frx":5BA1
               Left            =   3000
               List            =   "Radioterapeuta.frx":5BB8
               Style           =   2  'Dropdown List
               TabIndex        =   357
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   356
               Top             =   1200
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               ItemData        =   "Radioterapeuta.frx":5BEB
               Left            =   3000
               List            =   "Radioterapeuta.frx":5C02
               Style           =   2  'Dropdown List
               TabIndex        =   355
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   354
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   353
               Top             =   3435
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               ItemData        =   "Radioterapeuta.frx":5C35
               Left            =   3000
               List            =   "Radioterapeuta.frx":5C4C
               Style           =   2  'Dropdown List
               TabIndex        =   352
               Top             =   3420
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               ItemData        =   "Radioterapeuta.frx":5C7F
               Left            =   3000
               List            =   "Radioterapeuta.frx":5C96
               Style           =   2  'Dropdown List
               TabIndex        =   351
               Top             =   3000
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   8
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   350
               Top             =   3915
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               ItemData        =   "Radioterapeuta.frx":5CC9
               Left            =   3000
               List            =   "Radioterapeuta.frx":5CE0
               Style           =   2  'Dropdown List
               TabIndex        =   349
               Top             =   3900
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   9
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   348
               Top             =   855
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   9
               ItemData        =   "Radioterapeuta.frx":5D13
               Left            =   8280
               List            =   "Radioterapeuta.frx":5D2A
               Style           =   2  'Dropdown List
               TabIndex        =   347
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   10
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   346
               Top             =   1215
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   10
               ItemData        =   "Radioterapeuta.frx":5D5D
               Left            =   8280
               List            =   "Radioterapeuta.frx":5D74
               Style           =   2  'Dropdown List
               TabIndex        =   345
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   11
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   344
               Top             =   1575
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   11
               ItemData        =   "Radioterapeuta.frx":5DA7
               Left            =   8280
               List            =   "Radioterapeuta.frx":5DBE
               Style           =   2  'Dropdown List
               TabIndex        =   343
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   12
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   342
               Top             =   1995
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   12
               ItemData        =   "Radioterapeuta.frx":5DF1
               Left            =   8280
               List            =   "Radioterapeuta.frx":5E08
               Style           =   2  'Dropdown List
               TabIndex        =   341
               Top             =   1980
               Width           =   1335
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulmón"
               Height          =   195
               Left            =   7605
               TabIndex        =   385
               Top             =   900
               Width           =   525
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Genitourinario"
               Height          =   195
               Left            =   6360
               TabIndex        =   384
               Top             =   1260
               Width           =   1770
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corazón"
               Height          =   195
               Left            =   7545
               TabIndex        =   383
               Top             =   1620
               Width           =   585
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sistema Nervioso Central"
               Height          =   195
               Index           =   0
               Left            =   6360
               TabIndex        =   382
               Top             =   2040
               Width           =   1770
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piel"
               Height          =   195
               Left            =   2640
               TabIndex        =   381
               Top             =   900
               Width           =   255
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Laringe"
               Height          =   195
               Left            =   2370
               TabIndex        =   379
               Top             =   3060
               Width           =   525
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Faringe Esófago"
               Height          =   195
               Left            =   1740
               TabIndex        =   378
               Top             =   2700
               Width           =   1155
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glándulas Salivales"
               Height          =   195
               Left            =   1515
               TabIndex        =   377
               Top             =   2340
               Width           =   1380
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Oido"
               Height          =   195
               Left            =   2565
               TabIndex        =   376
               Top             =   1980
               Width           =   330
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ojo"
               Height          =   195
               Left            =   2655
               TabIndex        =   375
               Top             =   1620
               Width           =   240
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Membranas Mucosas"
               Height          =   315
               Left            =   120
               TabIndex        =   374
               Top             =   1200
               Width           =   2775
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tracto Gastro Intestinal Superior"
               Height          =   195
               Left            =   615
               TabIndex        =   373
               Top             =   3480
               Width           =   2280
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tracto Gastro Intestinal Inferior Incluyendo Pelvis"
               Height          =   435
               Left            =   120
               TabIndex        =   372
               Top             =   3840
               Width           =   2775
            End
            Begin VB.Line Line9 
               X1              =   120
               X2              =   11400
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line8 
               X1              =   6120
               X2              =   6120
               Y1              =   360
               Y2              =   4680
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Órgano o Tejido"
               Height          =   195
               Index           =   0
               Left            =   1440
               TabIndex        =   371
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   0
               Left            =   3360
               TabIndex        =   370
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   0
               Left            =   4440
               TabIndex        =   369
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Órgano o Tejido"
               Height          =   195
               Index           =   1
               Left            =   6660
               TabIndex        =   368
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   1
               Left            =   8640
               TabIndex        =   367
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   1
               Left            =   9720
               TabIndex        =   366
               Top             =   360
               Width           =   1320
            End
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   1
            Left            =   120
            TabIndex        =   310
            ToolTipText     =   "Toxicidad Hematológica Aguda"
            Top             =   840
            Width           =   11535
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   17
               ItemData        =   "Radioterapeuta.frx":5E3B
               Left            =   2880
               List            =   "Radioterapeuta.frx":5E52
               Style           =   2  'Dropdown List
               TabIndex        =   397
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   17
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   396
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   16
               ItemData        =   "Radioterapeuta.frx":5E85
               Left            =   2880
               List            =   "Radioterapeuta.frx":5E9C
               Style           =   2  'Dropdown List
               TabIndex        =   395
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   16
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   394
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   15
               ItemData        =   "Radioterapeuta.frx":5ECF
               Left            =   2880
               List            =   "Radioterapeuta.frx":5EE6
               Style           =   2  'Dropdown List
               TabIndex        =   391
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   15
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   390
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   14
               ItemData        =   "Radioterapeuta.frx":5F19
               Left            =   2880
               List            =   "Radioterapeuta.frx":5F30
               Style           =   2  'Dropdown List
               TabIndex        =   389
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   14
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   388
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   13
               ItemData        =   "Radioterapeuta.frx":5F63
               Left            =   2880
               List            =   "Radioterapeuta.frx":5F7A
               Style           =   2  'Dropdown List
               TabIndex        =   387
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   13
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   386
               Top             =   735
               Width           =   1335
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   393
               Top             =   360
               Width           =   840
            End
            Begin VB.Label Label51 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   2
               Left            =   2880
               TabIndex        =   392
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label94 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glóbulos Blancos (x10 ³/ml)"
               Height          =   195
               Left            =   435
               TabIndex        =   339
               Top             =   720
               Width           =   1950
            End
            Begin VB.Label Label95 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plaquetas (x10 ³/ml)"
               Height          =   195
               Left            =   705
               TabIndex        =   338
               Top             =   1080
               Width           =   1410
            End
            Begin VB.Label Label96 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Neutrófilos (x10 ³/ml)"
               Height          =   195
               Left            =   675
               TabIndex        =   337
               Top             =   1440
               Width           =   1470
            End
            Begin VB.Label Label97 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hemoglobina (g/dl)"
               Height          =   195
               Left            =   735
               TabIndex        =   336
               Top             =   1800
               Width           =   1350
            End
            Begin VB.Label Label98 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hematocrito (%)"
               Height          =   195
               Left            =   855
               TabIndex        =   335
               Top             =   2160
               Width           =   1110
            End
            Begin VB.Line Line10 
               X1              =   120
               X2              =   5880
               Y1              =   600
               Y2              =   600
            End
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Opciones"
            Height          =   495
            Index           =   3
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   443
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Complicaciones Crónicas"
            Height          =   495
            Index           =   2
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   311
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Toxicidad Hematológica Aguda"
            Height          =   495
            Index           =   1
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   308
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Complicaciones Agudas"
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   309
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   2
            Left            =   120
            TabIndex        =   312
            ToolTipText     =   "Complicaciones Crónicas"
            Top             =   840
            Width           =   11535
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   34
               ItemData        =   "Radioterapeuta.frx":5FAD
               Left            =   8040
               List            =   "Radioterapeuta.frx":5FC4
               Style           =   2  'Dropdown List
               TabIndex        =   431
               Top             =   2880
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   34
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   430
               Top             =   2895
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   33
               ItemData        =   "Radioterapeuta.frx":5FF7
               Left            =   8040
               List            =   "Radioterapeuta.frx":600E
               Style           =   2  'Dropdown List
               TabIndex        =   429
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   33
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   428
               Top             =   2535
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   32
               ItemData        =   "Radioterapeuta.frx":6041
               Left            =   8040
               List            =   "Radioterapeuta.frx":6058
               Style           =   2  'Dropdown List
               TabIndex        =   427
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   32
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   426
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   31
               ItemData        =   "Radioterapeuta.frx":608B
               Left            =   8040
               List            =   "Radioterapeuta.frx":60A2
               Style           =   2  'Dropdown List
               TabIndex        =   425
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   31
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   424
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   30
               ItemData        =   "Radioterapeuta.frx":60D5
               Left            =   8040
               List            =   "Radioterapeuta.frx":60EC
               Style           =   2  'Dropdown List
               TabIndex        =   423
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   30
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   422
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   29
               ItemData        =   "Radioterapeuta.frx":611F
               Left            =   8040
               List            =   "Radioterapeuta.frx":6136
               Style           =   2  'Dropdown List
               TabIndex        =   421
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   29
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   420
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   28
               ItemData        =   "Radioterapeuta.frx":6169
               Left            =   8040
               List            =   "Radioterapeuta.frx":6180
               Style           =   2  'Dropdown List
               TabIndex        =   419
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   28
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   418
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   27
               ItemData        =   "Radioterapeuta.frx":61B3
               Left            =   2520
               List            =   "Radioterapeuta.frx":61CA
               Style           =   2  'Dropdown List
               TabIndex        =   417
               Top             =   3960
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   27
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   416
               Top             =   3975
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   26
               ItemData        =   "Radioterapeuta.frx":61FD
               Left            =   2520
               List            =   "Radioterapeuta.frx":6214
               Style           =   2  'Dropdown List
               TabIndex        =   415
               Top             =   3600
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   26
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   414
               Top             =   3615
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   25
               ItemData        =   "Radioterapeuta.frx":6247
               Left            =   2520
               List            =   "Radioterapeuta.frx":625E
               Style           =   2  'Dropdown List
               TabIndex        =   413
               Top             =   3240
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   25
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   412
               Top             =   3255
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   24
               ItemData        =   "Radioterapeuta.frx":6291
               Left            =   2520
               List            =   "Radioterapeuta.frx":62A8
               Style           =   2  'Dropdown List
               TabIndex        =   411
               Top             =   2880
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   24
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   410
               Top             =   2895
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   23
               ItemData        =   "Radioterapeuta.frx":62DB
               Left            =   2520
               List            =   "Radioterapeuta.frx":62F2
               Style           =   2  'Dropdown List
               TabIndex        =   409
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   23
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   408
               Top             =   2535
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   22
               ItemData        =   "Radioterapeuta.frx":6325
               Left            =   2520
               List            =   "Radioterapeuta.frx":633C
               Style           =   2  'Dropdown List
               TabIndex        =   407
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   22
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   406
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   21
               ItemData        =   "Radioterapeuta.frx":636F
               Left            =   2520
               List            =   "Radioterapeuta.frx":6386
               Style           =   2  'Dropdown List
               TabIndex        =   405
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   21
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   404
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   20
               ItemData        =   "Radioterapeuta.frx":63B9
               Left            =   2520
               List            =   "Radioterapeuta.frx":63D0
               Style           =   2  'Dropdown List
               TabIndex        =   403
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   20
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   402
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   19
               ItemData        =   "Radioterapeuta.frx":6403
               Left            =   2520
               List            =   "Radioterapeuta.frx":641A
               Style           =   2  'Dropdown List
               TabIndex        =   401
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   19
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   400
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   18
               ItemData        =   "Radioterapeuta.frx":644D
               Left            =   2520
               List            =   "Radioterapeuta.frx":6464
               Style           =   2  'Dropdown List
               TabIndex        =   399
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   18
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   398
               Top             =   735
               Width           =   1335
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo / Meses  "
               Height          =   195
               Index           =   4
               Left            =   9510
               TabIndex        =   437
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   4
               Left            =   8400
               TabIndex        =   436
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organo o Tejido"
               Height          =   315
               Index           =   3
               Left            =   6480
               TabIndex        =   435
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo / Meses  "
               Height          =   195
               Index           =   3
               Left            =   3990
               TabIndex        =   434
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   3
               Left            =   2880
               TabIndex        =   433
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organo o Tejido"
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   432
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label114 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Articulación"
               Height          =   195
               Left            =   5400
               TabIndex        =   329
               Top             =   2940
               Width           =   2340
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hueso"
               Height          =   195
               Left            =   5400
               TabIndex        =   328
               Top             =   2580
               Width           =   2340
            End
            Begin VB.Label Label112 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Vejiga"
               Height          =   195
               Left            =   5400
               TabIndex        =   327
               Top             =   2220
               Width           =   2340
            End
            Begin VB.Label Label111 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Riñon"
               Height          =   195
               Left            =   5400
               TabIndex        =   326
               Top             =   1860
               Width           =   2340
            End
            Begin VB.Label Label107 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hígado"
               Height          =   195
               Left            =   5400
               TabIndex        =   325
               Top             =   1500
               Width           =   2340
            End
            Begin VB.Label Label110 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cerebro"
               Height          =   195
               Left            =   1785
               TabIndex        =   324
               Top             =   2580
               Width           =   555
            End
            Begin VB.Label Label109 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Médula Espinal"
               Height          =   195
               Left            =   120
               TabIndex        =   323
               Top             =   2220
               Width           =   2220
            End
            Begin VB.Label Label108 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Glándulas Salivales"
               Height          =   195
               Left            =   120
               TabIndex        =   322
               Top             =   1860
               Width           =   2220
            End
            Begin VB.Label Label106 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Membranas Mucosas"
               Height          =   195
               Left            =   120
               TabIndex        =   321
               Top             =   1500
               Width           =   2220
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tejido Subcutáneo"
               Height          =   195
               Left            =   120
               TabIndex        =   320
               Top             =   1140
               Width           =   2220
            End
            Begin VB.Label Label104 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piel"
               Height          =   195
               Left            =   2085
               TabIndex        =   319
               Top             =   780
               Width           =   255
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ojo"
               Height          =   195
               Left            =   120
               TabIndex        =   318
               Top             =   2940
               Width           =   2220
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laringe"
               Height          =   195
               Left            =   120
               TabIndex        =   317
               Top             =   3300
               Width           =   2220
            End
            Begin VB.Line Line12 
               X1              =   5760
               X2              =   5760
               Y1              =   360
               Y2              =   4680
            End
            Begin VB.Line Line11 
               X1              =   120
               X2              =   11400
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulmón"
               Height          =   195
               Left            =   1815
               TabIndex        =   316
               Top             =   3660
               Width           =   525
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Esófago"
               Height          =   195
               Left            =   5400
               TabIndex        =   315
               Top             =   780
               Width           =   2340
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corazón"
               Height          =   195
               Left            =   1755
               TabIndex        =   314
               Top             =   4020
               Width           =   585
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Intestino Grueso Delgado"
               Height          =   195
               Left            =   5400
               TabIndex        =   313
               Top             =   1140
               Width           =   2340
            End
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   3
            Left            =   120
            TabIndex        =   438
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtOtrasObs 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   2655
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   445
               Top             =   1320
               Width           =   5175
            End
            Begin SystemOncoAmerica.DMGrid DMGrid1 
               Height          =   3255
               Left            =   5880
               TabIndex        =   442
               Top             =   720
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   5741
               Object.Width           =   5505
               Object.Height          =   3225
               BackColor       =   15396847
               ScrollBar       =   1
            End
            Begin ChamaleonButton.ChameleonBtn BtnAgregarComplica 
               Height          =   375
               Left            =   360
               TabIndex        =   439
               ToolTipText     =   "Agregar "
               Top             =   4320
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
               MICON           =   "Radioterapeuta.frx":6497
               PICN            =   "Radioterapeuta.frx":64B3
               PICH            =   "Radioterapeuta.frx":6640
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnGuardarComplica 
               Height          =   375
               Left            =   1680
               TabIndex        =   440
               ToolTipText     =   "Guardar / Actualizar "
               Top             =   4320
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
               MICON           =   "Radioterapeuta.frx":6875
               PICN            =   "Radioterapeuta.frx":6891
               PICH            =   "Radioterapeuta.frx":6B20
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   2880
               TabIndex        =   441
               Top             =   390
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   48037889
               CurrentDate     =   40371
            End
            Begin ChamaleonButton.ChameleonBtn BtnEliminarComplica 
               Height          =   375
               Left            =   9600
               TabIndex        =   447
               ToolTipText     =   "Eliminar"
               Top             =   4320
               Width           =   1815
               _ExtentX        =   3201
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
               MICON           =   "Radioterapeuta.frx":6F61
               PICN            =   "Radioterapeuta.frx":6F7D
               PICH            =   "Radioterapeuta.frx":7121
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnCancelar 
               Height          =   375
               Left            =   3000
               TabIndex        =   450
               ToolTipText     =   "Eliminar"
               Top             =   4320
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Cancelar"
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
               MICON           =   "Radioterapeuta.frx":72C0
               PICN            =   "Radioterapeuta.frx":72DC
               PICH            =   "Radioterapeuta.frx":7480
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de creación del informe:"
               Height          =   195
               Left            =   480
               TabIndex        =   448
               Top             =   480
               Width           =   2190
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otras observaciones:"
               Height          =   195
               Left            =   240
               TabIndex        =   446
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Listado general"
               Height          =   195
               Left            =   6000
               TabIndex        =   444
               Top             =   360
               Width           =   1080
            End
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "InmunoHistoquimica"
         Height          =   5895
         Index           =   2
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   11775
         Begin VB.Frame Frame10 
            BackColor       =   &H00EAEFEF&
            Caption         =   "InmunoHistoquímica"
            Height          =   5415
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   11535
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD99/MIC-2"
               Height          =   375
               Index           =   47
               Left            =   7080
               TabIndex        =   223
               Top             =   2040
               Width           =   1215
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   56
               ItemData        =   "Radioterapeuta.frx":761F
               Left            =   10560
               List            =   "Radioterapeuta.frx":762F
               TabIndex        =   222
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   55
               ItemData        =   "Radioterapeuta.frx":7643
               Left            =   8400
               List            =   "Radioterapeuta.frx":7653
               TabIndex        =   221
               Top             =   4920
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   54
               ItemData        =   "Radioterapeuta.frx":7667
               Left            =   8400
               List            =   "Radioterapeuta.frx":7677
               TabIndex        =   220
               Top             =   4560
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   53
               ItemData        =   "Radioterapeuta.frx":768B
               Left            =   8400
               List            =   "Radioterapeuta.frx":769B
               TabIndex        =   219
               Top             =   4200
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   52
               ItemData        =   "Radioterapeuta.frx":76AF
               Left            =   8400
               List            =   "Radioterapeuta.frx":76BF
               TabIndex        =   218
               Top             =   3840
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   51
               ItemData        =   "Radioterapeuta.frx":76D3
               Left            =   8400
               List            =   "Radioterapeuta.frx":76E3
               TabIndex        =   217
               Top             =   3480
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   50
               ItemData        =   "Radioterapeuta.frx":76F7
               Left            =   8400
               List            =   "Radioterapeuta.frx":7707
               TabIndex        =   216
               Top             =   3120
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   49
               ItemData        =   "Radioterapeuta.frx":771B
               Left            =   8400
               List            =   "Radioterapeuta.frx":772B
               TabIndex        =   215
               Top             =   2760
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   48
               ItemData        =   "Radioterapeuta.frx":773F
               Left            =   8400
               List            =   "Radioterapeuta.frx":774F
               TabIndex        =   214
               Top             =   2400
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "OTROS"
               Height          =   375
               Index           =   57
               Left            =   9360
               TabIndex        =   213
               Top             =   600
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "WT"
               Height          =   375
               Index           =   56
               Left            =   9360
               TabIndex        =   212
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD15/LEUM1"
               Height          =   375
               Index           =   55
               Left            =   7080
               TabIndex        =   211
               Top             =   4920
               Width           =   1335
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD30/KL-1/BERH-2"
               Height          =   375
               Index           =   54
               Left            =   7080
               TabIndex        =   210
               Top             =   4560
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD3"
               Height          =   375
               Index           =   53
               Left            =   7080
               TabIndex        =   209
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD45-RO/UCHL-1"
               Height          =   375
               Index           =   52
               Left            =   7080
               TabIndex        =   208
               Top             =   3840
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD79A"
               Height          =   375
               Index           =   51
               Left            =   7080
               TabIndex        =   207
               Top             =   3480
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD20/L26"
               Height          =   375
               Index           =   50
               Left            =   7080
               TabIndex        =   206
               Top             =   3120
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "LCA/CD45"
               Height          =   375
               Index           =   49
               Left            =   7080
               TabIndex        =   205
               Top             =   2760
               Width           =   1215
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   47
               ItemData        =   "Radioterapeuta.frx":7763
               Left            =   8400
               List            =   "Radioterapeuta.frx":7773
               TabIndex        =   204
               Top             =   2040
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "NSD"
               Height          =   375
               Index           =   48
               Left            =   7080
               TabIndex        =   203
               Top             =   2400
               Width           =   1215
            End
            Begin VB.TextBox TxtOtros 
               Enabled         =   0   'False
               Height          =   375
               Left            =   9360
               TabIndex        =   202
               Top             =   960
               Width           =   2055
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   46
               ItemData        =   "Radioterapeuta.frx":7787
               Left            =   8400
               List            =   "Radioterapeuta.frx":7797
               TabIndex        =   201
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RA"
               Height          =   375
               Index           =   46
               Left            =   7080
               TabIndex        =   200
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   45
               ItemData        =   "Radioterapeuta.frx":77AB
               Left            =   8400
               List            =   "Radioterapeuta.frx":77BB
               TabIndex        =   199
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "ALK-1"
               Height          =   375
               Index           =   45
               Left            =   7080
               TabIndex        =   198
               Top             =   1320
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   44
               ItemData        =   "Radioterapeuta.frx":77CF
               Left            =   8400
               List            =   "Radioterapeuta.frx":77DF
               TabIndex        =   197
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "PRB"
               Height          =   375
               Index           =   44
               Left            =   7080
               TabIndex        =   196
               Top             =   960
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   43
               ItemData        =   "Radioterapeuta.frx":77F3
               Left            =   8400
               List            =   "Radioterapeuta.frx":7803
               TabIndex        =   195
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "BEL-2"
               Height          =   375
               Index           =   43
               Left            =   7080
               TabIndex        =   194
               Top             =   600
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   42
               ItemData        =   "Radioterapeuta.frx":7817
               Left            =   8400
               List            =   "Radioterapeuta.frx":7827
               TabIndex        =   193
               Top             =   270
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "BEL-1"
               Height          =   375
               Index           =   42
               Left            =   7080
               TabIndex        =   192
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   41
               ItemData        =   "Radioterapeuta.frx":783B
               Left            =   6120
               List            =   "Radioterapeuta.frx":784B
               TabIndex        =   191
               Top             =   4950
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "WT1"
               Height          =   375
               Index           =   41
               Left            =   5040
               TabIndex        =   190
               Top             =   4920
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   40
               ItemData        =   "Radioterapeuta.frx":785F
               Left            =   6120
               List            =   "Radioterapeuta.frx":786F
               TabIndex        =   189
               Top             =   4560
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HPAP"
               Height          =   375
               Index           =   40
               Left            =   5040
               TabIndex        =   188
               Top             =   4560
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   39
               ItemData        =   "Radioterapeuta.frx":7883
               Left            =   6120
               List            =   "Radioterapeuta.frx":7893
               TabIndex        =   187
               Top             =   4230
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HMB-45"
               Height          =   375
               Index           =   39
               Left            =   5040
               TabIndex        =   186
               Top             =   4200
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   38
               ItemData        =   "Radioterapeuta.frx":78A7
               Left            =   6120
               List            =   "Radioterapeuta.frx":78B7
               TabIndex        =   185
               Top             =   3870
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HCG"
               Height          =   375
               Index           =   38
               Left            =   5040
               TabIndex        =   184
               Top             =   3840
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   37
               ItemData        =   "Radioterapeuta.frx":78CB
               Left            =   6120
               List            =   "Radioterapeuta.frx":78DB
               TabIndex        =   183
               Top             =   3510
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "E-CAD"
               Height          =   375
               Index           =   37
               Left            =   5040
               TabIndex        =   182
               Top             =   3480
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   36
               ItemData        =   "Radioterapeuta.frx":78EF
               Left            =   6120
               List            =   "Radioterapeuta.frx":78FF
               TabIndex        =   181
               Top             =   3150
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CEA-D14"
               Height          =   375
               Index           =   36
               Left            =   5040
               TabIndex        =   180
               Top             =   3120
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   35
               ItemData        =   "Radioterapeuta.frx":7913
               Left            =   6120
               List            =   "Radioterapeuta.frx":7923
               TabIndex        =   179
               Top             =   2790
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CEA"
               Height          =   375
               Index           =   35
               Left            =   5040
               TabIndex        =   178
               Top             =   2760
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   34
               ItemData        =   "Radioterapeuta.frx":7937
               Left            =   6120
               List            =   "Radioterapeuta.frx":7947
               TabIndex        =   177
               Top             =   2400
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CA125"
               Height          =   375
               Index           =   34
               Left            =   5040
               TabIndex        =   176
               Top             =   2400
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   33
               ItemData        =   "Radioterapeuta.frx":795B
               Left            =   6120
               List            =   "Radioterapeuta.frx":796B
               TabIndex        =   175
               Top             =   2070
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CA199"
               Height          =   375
               Index           =   33
               Left            =   5040
               TabIndex        =   174
               Top             =   2040
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   32
               ItemData        =   "Radioterapeuta.frx":797F
               Left            =   6120
               List            =   "Radioterapeuta.frx":798F
               TabIndex        =   173
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "SMA"
               Height          =   375
               Index           =   32
               Left            =   5040
               TabIndex        =   172
               Top             =   1680
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   31
               ItemData        =   "Radioterapeuta.frx":79A3
               Left            =   6120
               List            =   "Radioterapeuta.frx":79B3
               TabIndex        =   171
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "GFAP"
               Height          =   375
               Index           =   31
               Left            =   5040
               TabIndex        =   170
               Top             =   1320
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   30
               ItemData        =   "Radioterapeuta.frx":79C7
               Left            =   6120
               List            =   "Radioterapeuta.frx":79D7
               TabIndex        =   169
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK903"
               Height          =   375
               Index           =   30
               Left            =   5040
               TabIndex        =   168
               Top             =   960
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   29
               ItemData        =   "Radioterapeuta.frx":79EB
               Left            =   6120
               List            =   "Radioterapeuta.frx":79FB
               TabIndex        =   167
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "AE3"
               Height          =   375
               Index           =   29
               Left            =   5040
               TabIndex        =   166
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "AE1"
               Height          =   375
               Index           =   28
               Left            =   5040
               TabIndex        =   165
               Top             =   240
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   27
               ItemData        =   "Radioterapeuta.frx":7A0F
               Left            =   4080
               List            =   "Radioterapeuta.frx":7A1F
               TabIndex        =   164
               Top             =   4950
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "KIT"
               Height          =   375
               Index           =   27
               Left            =   2400
               TabIndex        =   163
               Top             =   4920
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   26
               ItemData        =   "Radioterapeuta.frx":7A33
               Left            =   4080
               List            =   "Radioterapeuta.frx":7A43
               TabIndex        =   162
               Top             =   4590
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "EGFR"
               Height          =   375
               Index           =   26
               Left            =   2400
               TabIndex        =   161
               Top             =   4560
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   25
               ItemData        =   "Radioterapeuta.frx":7A57
               Left            =   4080
               List            =   "Radioterapeuta.frx":7A67
               TabIndex        =   160
               Top             =   4230
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD57"
               Height          =   375
               Index           =   25
               Left            =   2400
               TabIndex        =   159
               Top             =   4200
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   24
               ItemData        =   "Radioterapeuta.frx":7A7B
               Left            =   4080
               List            =   "Radioterapeuta.frx":7A8B
               TabIndex        =   158
               Top             =   3870
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD56"
               Height          =   375
               Index           =   24
               Left            =   2400
               TabIndex        =   157
               Top             =   3840
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   23
               ItemData        =   "Radioterapeuta.frx":7A9F
               Left            =   4080
               List            =   "Radioterapeuta.frx":7AAF
               TabIndex        =   156
               Top             =   3510
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "SINAPTOFISINA"
               Height          =   375
               Index           =   23
               Left            =   2400
               TabIndex        =   155
               Top             =   3480
               Width           =   1575
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   22
               ItemData        =   "Radioterapeuta.frx":7AC3
               Left            =   4080
               List            =   "Radioterapeuta.frx":7AD3
               TabIndex        =   154
               Top             =   3150
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CROMOGRANINA"
               Height          =   375
               Index           =   22
               Left            =   2400
               TabIndex        =   153
               Top             =   3120
               Width           =   1695
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   21
               ItemData        =   "Radioterapeuta.frx":7AE7
               Left            =   4080
               List            =   "Radioterapeuta.frx":7AF7
               TabIndex        =   152
               Top             =   2790
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "TTF-1"
               Height          =   375
               Index           =   21
               Left            =   2400
               TabIndex        =   151
               Top             =   2760
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   20
               ItemData        =   "Radioterapeuta.frx":7B0B
               Left            =   4080
               List            =   "Radioterapeuta.frx":7B1B
               TabIndex        =   150
               Top             =   2430
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CAM 5,2"
               Height          =   375
               Index           =   20
               Left            =   2400
               TabIndex        =   149
               Top             =   2400
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   19
               ItemData        =   "Radioterapeuta.frx":7B2F
               Left            =   4080
               List            =   "Radioterapeuta.frx":7B3F
               TabIndex        =   148
               Top             =   2070
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK20"
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   147
               Top             =   2040
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   18
               ItemData        =   "Radioterapeuta.frx":7B53
               Left            =   4080
               List            =   "Radioterapeuta.frx":7B63
               TabIndex        =   146
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK7"
               Height          =   375
               Index           =   18
               Left            =   2400
               TabIndex        =   145
               Top             =   1680
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   17
               ItemData        =   "Radioterapeuta.frx":7B77
               Left            =   4080
               List            =   "Radioterapeuta.frx":7B87
               TabIndex        =   144
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK6"
               Height          =   375
               Index           =   17
               Left            =   2400
               TabIndex        =   143
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   16
               ItemData        =   "Radioterapeuta.frx":7B9B
               Left            =   4080
               List            =   "Radioterapeuta.frx":7BAB
               TabIndex        =   142
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK5"
               Height          =   375
               Index           =   16
               Left            =   2400
               TabIndex        =   141
               Top             =   960
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   15
               ItemData        =   "Radioterapeuta.frx":7BBF
               Left            =   4080
               List            =   "Radioterapeuta.frx":7BCF
               TabIndex        =   140
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD117"
               Height          =   375
               Index           =   15
               Left            =   2400
               TabIndex        =   139
               Top             =   600
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   14
               ItemData        =   "Radioterapeuta.frx":7BE3
               Left            =   4080
               List            =   "Radioterapeuta.frx":7BF3
               TabIndex        =   138
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD34"
               Height          =   375
               Index           =   14
               Left            =   2400
               TabIndex        =   137
               Top             =   240
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   13
               ItemData        =   "Radioterapeuta.frx":7C07
               Left            =   1440
               List            =   "Radioterapeuta.frx":7C17
               TabIndex        =   136
               Top             =   4950
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD31"
               Height          =   375
               Index           =   13
               Left            =   120
               TabIndex        =   135
               Top             =   4920
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   12
               ItemData        =   "Radioterapeuta.frx":7C2B
               Left            =   1440
               List            =   "Radioterapeuta.frx":7C3B
               TabIndex        =   134
               Top             =   4590
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "PGP"
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   133
               Top             =   4560
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   11
               ItemData        =   "Radioterapeuta.frx":7C4F
               Left            =   1440
               List            =   "Radioterapeuta.frx":7C5F
               TabIndex        =   132
               Top             =   4230
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "PROT-S-100"
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   131
               Top             =   4200
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   10
               ItemData        =   "Radioterapeuta.frx":7C73
               Left            =   1440
               List            =   "Radioterapeuta.frx":7C83
               TabIndex        =   130
               Top             =   3870
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "AFP"
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   129
               Top             =   3840
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   9
               ItemData        =   "Radioterapeuta.frx":7C97
               Left            =   1440
               List            =   "Radioterapeuta.frx":7CA7
               TabIndex        =   128
               Top             =   3510
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "ACE"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   127
               Top             =   3480
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               ItemData        =   "Radioterapeuta.frx":7CBB
               Left            =   1440
               List            =   "Radioterapeuta.frx":7CCB
               TabIndex        =   126
               Top             =   3150
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "DESMINA"
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   125
               Top             =   3120
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               ItemData        =   "Radioterapeuta.frx":7CDF
               Left            =   1440
               List            =   "Radioterapeuta.frx":7CEF
               TabIndex        =   124
               Top             =   2790
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "P53"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   123
               Top             =   2760
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               ItemData        =   "Radioterapeuta.frx":7D03
               Left            =   1440
               List            =   "Radioterapeuta.frx":7D13
               TabIndex        =   122
               Top             =   2430
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CERB-2"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   121
               Top             =   2400
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               ItemData        =   "Radioterapeuta.frx":7D27
               Left            =   1440
               List            =   "Radioterapeuta.frx":7D37
               TabIndex        =   120
               Top             =   2070
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CAE"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   119
               Top             =   2040
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               ItemData        =   "Radioterapeuta.frx":7D4B
               Left            =   1440
               List            =   "Radioterapeuta.frx":7D5B
               TabIndex        =   118
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "VIM"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   117
               Top             =   1680
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               ItemData        =   "Radioterapeuta.frx":7D6F
               Left            =   1440
               List            =   "Radioterapeuta.frx":7D7F
               TabIndex        =   116
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "EMA"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   115
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               ItemData        =   "Radioterapeuta.frx":7D93
               Left            =   1440
               List            =   "Radioterapeuta.frx":7DA3
               TabIndex        =   114
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HER/2-NEU"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   113
               Top             =   960
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               ItemData        =   "Radioterapeuta.frx":7DB7
               Left            =   1440
               List            =   "Radioterapeuta.frx":7DC7
               TabIndex        =   112
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RP"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   111
               Top             =   600
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               ItemData        =   "Radioterapeuta.frx":7DDB
               Left            =   1440
               List            =   "Radioterapeuta.frx":7DEB
               TabIndex        =   110
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RE"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   109
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   28
               ItemData        =   "Radioterapeuta.frx":7DFE
               Left            =   6120
               List            =   "Radioterapeuta.frx":7E0E
               TabIndex        =   108
               Top             =   240
               Width           =   735
            End
            Begin VB.Line Line7 
               X1              =   9240
               X2              =   9240
               Y1              =   240
               Y2              =   5280
            End
            Begin VB.Line Line6 
               X1              =   6960
               X2              =   6960
               Y1              =   240
               Y2              =   5280
            End
            Begin VB.Line Line5 
               X1              =   4920
               X2              =   4920
               Y1              =   240
               Y2              =   5280
            End
            Begin VB.Line Line4 
               X1              =   2280
               X2              =   2280
               Y1              =   240
               Y2              =   5280
            End
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Estadiaje / Seguimiento"
         Height          =   5895
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   11775
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Estadiaje:"
            Height          =   3015
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2055
            Begin VB.ComboBox Combo7 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":7E22
               Left            =   840
               List            =   "Radioterapeuta.frx":7E35
               TabIndex        =   29
               Text            =   "G."
               Top             =   2040
               Width           =   855
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":7E4D
               Left            =   840
               List            =   "Radioterapeuta.frx":7ECC
               TabIndex        =   28
               Text            =   "A."
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":7F97
               Left            =   840
               List            =   "Radioterapeuta.frx":7FA7
               TabIndex        =   27
               Text            =   "C."
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":7FBC
               Left            =   840
               List            =   "Radioterapeuta.frx":7FFF
               TabIndex        =   26
               Text            =   "T."
               Top             =   600
               Width           =   855
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":8067
               Left            =   840
               List            =   "Radioterapeuta.frx":8098
               TabIndex        =   25
               Text            =   "N."
               Top             =   960
               Width           =   855
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":80E2
               Left            =   840
               List            =   "Radioterapeuta.frx":80FE
               TabIndex        =   24
               Text            =   "M."
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox ChkGleason 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Gleason"
               Height          =   375
               Left            =   120
               TabIndex        =   23
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox TxtGleason 
               Height          =   315
               Left            =   1080
               TabIndex        =   22
               Top             =   2430
               Width           =   855
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "G:"
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   2100
               Width           =   165
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estadio:"
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   1740
               Width           =   570
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C / P:"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   300
               Width           =   420
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "T:"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   660
               Width           =   150
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N:"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   1020
               Width           =   165
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "M:"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   1380
               Width           =   180
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Bordes de Resección"
            Height          =   2295
            Left            =   120
            TabIndex        =   15
            Top             =   3360
            Width           =   2055
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Rx"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RO"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "R1"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   18
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "R2"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   17
               Top             =   1320
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "R3"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   16
               Top             =   1680
               Width           =   975
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Seguimiento"
            Height          =   5415
            Left            =   2280
            TabIndex        =   224
            Top             =   240
            Width           =   9375
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   334
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   6
               Left            =   240
               MaxLength       =   5
               TabIndex        =   333
               Top             =   3840
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   4
               Left            =   240
               MaxLength       =   5
               TabIndex        =   332
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   2
               Left            =   240
               MaxLength       =   5
               TabIndex        =   331
               Top             =   1680
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   0
               Left            =   240
               MaxLength       =   5
               TabIndex        =   330
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   7
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   279
               Top             =   4170
               Width           =   975
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   278
               Top             =   4200
               Width           =   1215
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   277
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   276
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   275
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   274
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   273
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   272
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   271
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   270
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   269
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   268
               Top             =   3840
               Width           =   975
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   267
               Top             =   3840
               Width           =   1095
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   5
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   265
               Top             =   3090
               Width           =   975
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   264
               Top             =   3120
               Width           =   1215
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   263
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   262
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   261
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   260
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   259
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   258
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   257
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   256
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   255
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   254
               Top             =   2760
               Width           =   975
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   253
               Top             =   2760
               Width           =   1095
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   3
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   251
               Top             =   2010
               Width           =   975
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   250
               Top             =   2040
               Width           =   1215
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   249
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   248
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   247
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   246
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   245
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   244
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   243
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   242
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   241
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   240
               Top             =   1680
               Width           =   975
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   239
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   1
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   237
               Top             =   937
               Width           =   975
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   236
               Top             =   960
               Width           =   1215
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   235
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   234
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   233
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   232
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   231
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   230
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   229
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   228
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   227
               Top             =   600
               Width           =   975
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   226
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Muerte:"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   266
               Top             =   3600
               Width           =   540
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recaida:"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   252
               Top             =   2520
               Width           =   645
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Progresión:"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   238
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Libre de Enfermedad:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   225
               Top             =   360
               Width           =   1515
            End
            Begin VB.Line Line3 
               X1              =   120
               X2              =   9120
               Y1              =   3480
               Y2              =   3480
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   9120
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   9120
               Y1              =   1320
               Y2              =   1320
            End
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnVerInforme 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ver Informe"
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
         MICON           =   "Radioterapeuta.frx":8125
         PICN            =   "Radioterapeuta.frx":8141
         PICH            =   "Radioterapeuta.frx":83DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnVerHistoria 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ver Historia"
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
         MICON           =   "Radioterapeuta.frx":881D
         PICN            =   "Radioterapeuta.frx":8839
         PICH            =   "Radioterapeuta.frx":8AC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnEvolucionOncologica 
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Evolución Clinica"
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
         MICON           =   "Radioterapeuta.frx":8F08
         PICN            =   "Radioterapeuta.frx":8F24
         PICH            =   "Radioterapeuta.frx":91BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAntecedentes 
         Height          =   375
         Left            =   4800
         TabIndex        =   449
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Antecedentes"
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
         MICON           =   "Radioterapeuta.frx":9443
         PICN            =   "Radioterapeuta.frx":945F
         PICH            =   "Radioterapeuta.frx":96F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnExamenes 
         Height          =   375
         Left            =   6360
         TabIndex        =   451
         Top             =   6240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Examenes Laboratorio"
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
         MICON           =   "Radioterapeuta.frx":997B
         PICN            =   "Radioterapeuta.frx":9997
         PICH            =   "Radioterapeuta.frx":9DC2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Informe Médico Final"
         Height          =   5895
         Index           =   3
         Left            =   120
         TabIndex        =   280
         Top             =   240
         Width           =   11775
         Begin VB.TextBox TxtInidiceFin 
            Height          =   855
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   304
            Top             =   2640
            Width           =   5775
         End
         Begin VB.TextBox TxtCompliFin 
            Height          =   735
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   302
            Top             =   1560
            Width           =   5775
         End
         Begin VB.TextBox TxtExamFin 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   298
            Top             =   1560
            Width           =   5655
         End
         Begin VB.TextBox TxtDiagFin 
            Height          =   885
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   294
            Top             =   2640
            Width           =   5655
         End
         Begin VB.TextBox TxtAnatFin 
            Height          =   735
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   293
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox TxtExamFIni 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   292
            Top             =   480
            Width           =   5655
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento:"
            Height          =   2175
            Index           =   1
            Left            =   120
            TabIndex        =   281
            Top             =   3600
            Width           =   11535
            Begin VB.ComboBox Combo9 
               Height          =   315
               Left            =   7080
               Style           =   2  'Dropdown List
               TabIndex        =   300
               Top             =   1200
               Width           =   2415
            End
            Begin VB.TextBox TxtSesionesFin 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   10200
               TabIndex        =   286
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox TxtTratamientoFin 
               Height          =   1575
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   285
               Top             =   360
               Width           =   6855
            End
            Begin VB.TextBox TxtDosisT 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   7800
               TabIndex        =   284
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox TxtDosisD 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   9000
               TabIndex        =   283
               Top             =   480
               Width           =   1095
            End
            Begin VB.ComboBox Combo8 
               Height          =   315
               ItemData        =   "Radioterapeuta.frx":A05A
               Left            =   9600
               List            =   "Radioterapeuta.frx":A06A
               TabIndex        =   282
               Top             =   1200
               Width           =   1815
            End
            Begin ChamaleonButton.ChameleonBtn BtnInformeFinal 
               Height          =   375
               Left            =   9720
               TabIndex        =   452
               Top             =   1680
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Informe Final"
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
               MICON           =   "Radioterapeuta.frx":A09F
               PICN            =   "Radioterapeuta.frx":A0BB
               PICH            =   "Radioterapeuta.frx":A344
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Medico Tratante:"
               Height          =   195
               Left            =   7080
               TabIndex        =   301
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sesiones"
               Height          =   195
               Left            =   10200
               TabIndex        =   291
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total de Dosis"
               Height          =   195
               Left            =   7800
               TabIndex        =   290
               Top             =   240
               Width           =   1020
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis Diarias"
               Height          =   195
               Left            =   9000
               TabIndex        =   289
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Duración:"
               Height          =   195
               Left            =   7080
               TabIndex        =   288
               Top             =   570
               Width           =   690
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Metas:"
               Height          =   195
               Left            =   9600
               TabIndex        =   287
               Top             =   960
               Width           =   480
            End
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Respuesta Clínica:"
            Height          =   195
            Left            =   5880
            TabIndex        =   305
            Top             =   2400
            Width           =   1350
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Complicaciones:"
            Height          =   195
            Left            =   5880
            TabIndex        =   303
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico Final:"
            Height          =   195
            Left            =   120
            TabIndex        =   299
            Top             =   1320
            Width           =   1470
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   297
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dia&gnóstico:"
            Height          =   195
            Left            =   120
            TabIndex        =   296
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Anatomía &Patológica:"
            Height          =   195
            Left            =   5880
            TabIndex        =   295
            Top             =   240
            Width           =   1530
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente2 
         Height          =   375
         Left            =   11280
         TabIndex        =   453
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   6240
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
         MICON           =   "Radioterapeuta.frx":A5E8
         PICN            =   "Radioterapeuta.frx":A604
         PICH            =   "Radioterapeuta.frx":A89A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior1 
         Height          =   375
         Left            =   10560
         TabIndex        =   454
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   6240
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
         MICON           =   "Radioterapeuta.frx":AAF9
         PICN            =   "Radioterapeuta.frx":AB15
         PICH            =   "Radioterapeuta.frx":ADAA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnExamenIngreso 
         Height          =   375
         Left            =   8520
         TabIndex        =   455
         Top             =   6240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Examen de Ingreso"
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
         MICON           =   "Radioterapeuta.frx":B006
         PICN            =   "Radioterapeuta.frx":B022
         PICH            =   "Radioterapeuta.frx":B44D
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00EAEFEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   10080
      TabIndex        =   57
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "FrmRadioTerapeuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsInformeMed As New ADODB.Recordset         'Para desplazamientos
Dim RsCargarPacientes As New ADODB.Recordset
Dim CSql As String
Dim Cambio
Dim actualiza
Dim IdReg
Dim MError
Dim NuevoId2 As String
Dim NuevoId As Integer

Dim MaxReg As String
Dim d1, d2, d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, dm, ds, de
Dim Gram As TextoPos
Dim Sesiones As Double
Dim tomog, Estadiaje
Dim MaxRegP As String
Dim RsIdMax As New ADODB.Recordset
Dim RsRegPendiente As New ADODB.Recordset
Dim IdInf
Dim OptBordesSel As String
Dim RsTemp As New ADODB.Recordset
Dim i As Integer
Dim IdRegFinal

Sub IniDMGrid()

DMGrid1.Cols = 3

DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True

DMGrid1.DColumnas(3).Visible = False

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300

DMGrid1.DColumnas(1).Caption = "Fecha / Semana"
DMGrid1.DColumnas(2).Caption = "realizado por:"

' MMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMM Numero de semanas MMMMM

' MsgBox (DatePart("ww", DateValue(Now)) + 1) - DatePart("ww", "01/01/" & Year(Now))

' MMMMMMMMMMMMMMMMMMMMMMMMMMMM


End Sub

Sub CARGAR_COMPLICA()

DMGrid1.Clear
DMGrid1.Rows = 0

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
TxtOtrasObs.Enabled = False
TxtOtrasObs.BackColor = &HE0E0E0

DTPicker1.Enabled = False
BtnGuardarComplica.Enabled = False
For i = 0 To TxtDosisGrdo.Count - 1
    TxtDosisGrdo(i).Enabled = False
    TxtDosisGrdo(i).Text = ""
    TxtDosisGrdo(i).BackColor = &HE0E0E0
Next
For i = 0 To CboGrado.Count - 1
    CboGrado(i).Enabled = False
    CboGrado(i).ListIndex = 0
Next
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "SELECT caP,  caPTXT,  caMM,  caMMTXT,  caO,  caOTXT,  caOI,  caOITXT,  caGS,  caGSTXT,  caFE,  caFETXT, " & _
  " caL,  caLTXT,  caTGIS,  caTGISTXT,  caTGIIIP,  caGIIIPTXT,  caPUL,  caPULTXT,  caG,  caGTXT, " & _
  " caC,  caCTXT,  caS,  caSTXT, " & _
  " thaG,  thaGTXT,  thaP,  thaPTXT,  [thaN],  thaNTXT,  thaHGBN,  thaHGBNTXT,  thaHTCT,  thaHTCTTXT, " & _
  " ccP,  ccPTXT,  ccTS,  ccTSTXT,  ccMM,  ccMMTXT,  ccGS,  ccGSTXT,  ccME,  ccMETXT,  ccC,  ccCTXT, " & _
  " ccO,  ccOTXT,  ccL,  ccLTXT,  ccPUL,  ccPULTXT,  ccCORAZON,  ccCORAZONTXT,  ccE,  ccETXT, " & _
  " ccIGD , ccIGDTXT, ccH, ccHTXT, ccR, ccRTXT, ccV, ccVTXT, ccHUESO, ccHUESOTXT, ccA, ccATXT, Id, " & _
  " Apellidos, Informe_Medico5.Fecha, Nombre " & _
  " FROM Informe_Medico5 INNER JOIN Usuarios ON (Informe_Medico5.IdUsuario = Usuarios.IdUsuario) " & _
  " WHERE Informe_Medico5.IdInforme=" & IdInf & " ORDER BY ID"

Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

DMGrid1.Clear
DMGrid1.Rows = 0

DMGrid1.RowBackColor 1, vbWhite

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Fecha").Value & " / " & Val(DatePart("ww", DateValue(RsTemp.Fields("Fecha").Value)) + 1) - DatePart("ww", "01/01/" & Year(RsTemp.Fields("Fecha").Value))
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Nombre").Value & ", " & RsTemp.Fields("Apellidos").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Id").Value
    RsTemp.MoveNext
Wend

DMGrid1.PaintMGrid
BtnAgregarComplica.Enabled = True

End Sub

Sub CARGAR_INFORME2()

CSql = "SELECT Re,Rp,Her2Neu,EMA,VIM,CAE,[CERB-2],P53,DESMINA,ACE,AFP,[PROT-S-100],PGP,CD31,CD34, " & _
    " CD117,CK5,CK6,CK7,CK20,[CAM 5,2],[TTF-1],CROMOGRANINA,SINAPTOFISINA,CD56,CD57,EGFR,KIT,AE1,AE3, " & _
    " CK903,GFAP,SMA,CA199,CA125,CEA,[CEA-D14],[E-CAD],HCG,[HMB-45],HPAP,WT1,[BEL-1],[BEL-2],PRB,[ALK-1]," & _
    " RA,CD99MID2,NSD,LCACD45,CD20L26,CD79A,CD45ROUCHL1,CD3,CD30KL1BERH2,CD15LEUM1,WT,OTROS, " & _
    " IdTipoCancer,T,N,M,Estadio,CP,G,Gleason,Reseccion " & _
    " FROM Informe_Medico2 Where IdPaciente=" & IdPac1 & " AND IdInforme=" & IdInf
    
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then

    If Trim(RsTemp.Fields("CP").Value) <> "" Then
        For i = 0 To Combo3.ListCount - 1
            If Trim(RsTemp.Fields("CP").Value) = Combo3.List(i) Then
                Combo3.ListIndex = i
                Exit For
            End If
            Combo3.ListIndex = -1
        Next i
    End If
    
    If Trim(RsTemp.Fields("T").Value) <> "" Then
        For i = 0 To Combo4.ListCount - 1
            If Trim(RsTemp.Fields("T").Value) = Combo4.List(i) Then
                Combo4.ListIndex = i
                Exit For
            End If
            Combo4.ListIndex = -1
        Next i
    End If
    
    If Trim(RsTemp.Fields("N").Value) <> "" Then
        For i = 0 To Combo5.ListCount - 1
            If Trim(RsTemp.Fields("N").Value) = Combo5.List(i) Then
                Combo5.ListIndex = i
                Exit For
            End If
            Combo5.ListIndex = -1
        Next i
    End If
    
    If Trim(RsTemp.Fields("M").Value) <> "" Then
        For i = 0 To Combo6.ListCount - 1
            If Trim(RsTemp.Fields("M").Value) = Combo6.List(i) Then
                Combo6.ListIndex = i
                Exit For
            End If
            Combo6.ListIndex = -1
        Next i
    End If
    
    If Trim(RsTemp.Fields("Estadio").Value) <> "" Then
        For i = 0 To Combo2.ListCount - 1
            If Trim(RsTemp.Fields("Estadio").Value) = Combo2.List(i) Then
                Combo2.ListIndex = i
                Exit For
            End If
            Combo2.ListIndex = -1
        Next i
    End If
    
    If Trim(RsTemp.Fields("G").Value) <> "" Then
        For i = 0 To Combo7.ListCount - 1
            If Trim(RsTemp.Fields("G").Value) = Combo7.List(i) Then
                Combo7.ListIndex = i
                Exit For
            End If
            Combo7.ListIndex = -1
        Next i
    End If
    
    If Trim(RsTemp.Fields("Gleason").Value) <> "" Then
        TxtGleason.Text = Trim(RsTemp.Fields("Gleason").Value)
        ChkGleason.Value = 1
    Else
        TxtGleason.Text = ""
        ChkGleason.Value = 0
    End If
    
    OptBordes(0).Value = True
    OptBordes(0).Value = False
    For i = 0 To OptBordes.Count - 1
        If OptBordes(i).Caption = Trim(RsTemp.Fields("Reseccion").Value) Then
            OptBordes(i).Value = True
            Exit For
        End If
    Next i
    
    For i = 0 To CboTCancers.ListCount - 1
        If CboTCancers.ItemData(i) = RsTemp.Fields("IdTipoCancer").Value Then
            CboTCancers.ListIndex = i
            Exit For
        End If
        CboTCancers.ListIndex = -1
    Next i
    
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    For i = 0 To ChkGlobal.Count - 1
        ChkGlobal(i).BackColor = &HEAEFEF
        ChkGlobal(i).Value = 0
    Next i
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    ' Ciclo que carga TODOS los valores para la "Gran Matriz" para la inmonohistoquimica
    For i = 0 To ChkGlobal.Count - 1
        If Trim(RsTemp.Fields(i).Value) <> "" Then
        
            If i <= 56 Then
                For J = 0 To CboGeneral(i).ListCount - 1
                    If Trim(CboGeneral(i).List(J)) = Trim(RsTemp.Fields(i).Value) Then
                        CboGeneral(i).ListIndex = J
                        ChkGlobal(i).Value = 1
                        Exit For
                    End If
                Next J
            Else
                If Trim(RsTemp.Fields("OTROS").Value) <> "" Then
                    TxtOtros.Text = Trim(RsTemp.Fields("OTROS").Value)
                    ChkGlobal(57).Value = 1
                End If
            End If
        Else
            If i <= 56 Then
                ChkGlobal(i).Value = 0
                CboGeneral(i).ListIndex = -1
            Else
                ChkGlobal(57).Value = 0
                TxtOtros.Text = ""
            End If
        End If
    Next i
Else
    For i = 0 To ChkGlobal.Count - 1
        If i <= 56 Then
            ChkGlobal(i).Value = 0
            CboGeneral(i).ListIndex = -1
        Else
            ChkGlobal(57).Value = 0
            TxtOtros.Text = ""
        End If
    Next i
End If

For i = 0 To ChkMuert.Count - 1
    ChkMuert(i).BackColor = &HEAEFEF
    ChkMuert(i).Value = 0
Next i
For i = 0 To ChkRecaida.Count - 1
    ChkRecaida(i).BackColor = &HEAEFEF
    ChkRecaida(i).Value = 0
Next i
For i = 0 To LstProg.Count - 1
    LstProg(i).BackColor = &HEAEFEF
    LstProg(i).Value = 0
Next i
For i = 0 To LstEnfer.Count - 1
    LstEnfer(i).BackColor = &HEAEFEF
    LstEnfer(i).Value = 0
Next i

For i = 0 To Text5.Count - 1
    Text5(i).Text = ""
Next i

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMM consulta para mostrar los valores del seguimiento!
CSql = "SELECT LE_MinM,  LE_6M,  LE_12M,  LE_18M,  LE_24M,  LE_30M,  LE_36M,  LE_42M,  LE_48M,  LE_54M,  LE_60M,  LE_MaxM, " & _
  " P_MinM,  P_6M,  P_12M,  P_18M,  P_24M,  P_30M,  P_36M,  P_42M,  P_48M,  P_54M,  P_60M,  P_MaxM, " & _
  " R_MinM,  R_6M,  R_12M,  R_18M,  R_24M,  R_30M,  R_36M,  R_42M,  R_48M,  R_54M,  R_60M,  R_MaxM, " & _
  " M_MinM , M_6M, M_12M, M_18M, M_24M, M_30M, M_36M, M_42M, M_48M, M_54M, M_60M, M_MaxM " & _
  " FROM Informe_Medico3 Where IdPaciente=" & IdPac1 & " AND IdInforme=" & IdInf

' 0 - 11 campos / 12 - 23 campos / 24 - 35 campos / 36 - 47 campos

Set RsTemp = CrearRS(CSql)

' Carga TODOS los valores para al "SEGUIMIENTO"

If RsTemp.RecordCount <> 0 Then
    
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    For i = 1 To LstEnfer.Count - 2
        If RsTemp.Fields(i).Value Then
            LstEnfer(i).Value = 1
        End If
    Next i
    
    If Trim(RsTemp.Fields("LE_MaxM").Value) <> "" Then
        LstEnfer(LstEnfer.Count - 1).Value = 1
        Text5(1).Text = Trim(RsTemp.Fields("LE_MaxM").Value)
    End If
    If Trim(RsTemp.Fields("LE_MinM").Value) <> "" Then
        LstEnfer(0).Value = 1
        Text5(0).Text = Trim(RsTemp.Fields("LE_MinM").Value)
    End If
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  ' 13-22
    For i = 1 To LstProg.Count - 2
        If RsTemp.Fields(i + 12).Value Then
            LstProg(i).Value = 1
        End If
    Next i
    
    If Trim(RsTemp.Fields("P_MaxM").Value) <> "" Then
        LstProg(LstProg.Count - 1).Value = 1
        Text5(3).Text = Trim(RsTemp.Fields("P_MaxM").Value)
    End If
    If Trim(RsTemp.Fields("P_MinM").Value) <> "" Then
        LstProg(0).Value = 1
        Text5(2).Text = Trim(RsTemp.Fields("P_MinM").Value)
    End If
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  ' 25-34
    For i = 1 To ChkRecaida.Count - 2
        If RsTemp.Fields(i + 24).Value Then
            ChkRecaida(i).Value = 1
        End If
    Next i
    
    If Trim(RsTemp.Fields("R_MaxM").Value) <> "" Then
        ChkRecaida(ChkRecaida.Count - 1).Value = 1
        Text5(5).Text = Trim(RsTemp.Fields("R_MaxM").Value)
    End If
    If Trim(RsTemp.Fields("R_MinM").Value) <> "" Then
        ChkRecaida(0).Value = 1
        Text5(4).Text = Trim(RsTemp.Fields("R_MinM").Value)
    End If
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  ' 37-46
    For i = 1 To ChkMuert.Count - 2
        If RsTemp.Fields(i + 36).Value Then
            ChkMuert(i).Value = 1
        End If
    Next i
    
    If Trim(RsTemp.Fields("M_MaxM").Value) <> "" Then
        ChkMuert(ChkMuert.Count - 1).Value = 1
        Text5(7).Text = Trim(RsTemp.Fields("M_MaxM").Value)
    End If
    If Trim(RsTemp.Fields("M_MinM").Value) <> "" Then
        ChkMuert(0).Value = 1
        Text5(6).Text = Trim(RsTemp.Fields("M_MinM").Value)
    End If
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Else
    For i = 0 To LstEnfer.Count - 1
        LstEnfer(i).Value = 0
    Next i
    For i = 0 To LstProg.Count - 1
        LstProg(i).Value = 0
    Next i
    For i = 0 To ChkRecaida.Count - 1
        ChkRecaida(i).Value = 0
    Next i
    For i = 0 To ChkMuert.Count - 1
        ChkMuert(i).Value = 0
    Next i
    For i = 0 To Text5.Count - 1
        Text5(i).Text = ""
    Next i
End If

CARGAR_INFORME3
Cambio = 0
End Sub

Sub CARGAR_INFORME3()
' Carga el informe Final

CSql = "SELECT * FROM Informe_Medico4 Where IdPaciente=" & IdPac1 & " AND IdInforme=" & IdInf
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    
    TxtExamFIni.Text = RsTemp.Fields("ExamFisicoInicial").Value
    TxtExamFin.Text = RsTemp.Fields("ExamFisicoFinal").Value
    TxtDiagFin.Text = RsTemp.Fields("Diagnostico").Value
    TxtAnatFin.Text = RsTemp.Fields("Anatomia").Value
    TxtCompliFin.Text = RsTemp.Fields("Complicaciones").Value
    TxtInidiceFin.Text = RsTemp.Fields("Indice").Value
    TxtTratamientoFin.Text = RsTemp.Fields("Comentarios").Value
    TxtDosisT.Text = RsTemp.Fields("DosisT").Value
    TxtDosisD.Text = RsTemp.Fields("DosisD").Value
    TxtSesionesFin.Text = RsTemp.Fields("Sesiones").Value
    IdRegFinal = RsTemp.Fields("Id").Value
    If RsTemp.Fields("IdMedicoT").Value <> 0 Then
        For i = 0 To Combo9.ListCount - 1
            If Val(RsTemp.Fields("IdMedicoT").Value) = Combo9.ItemData(i) Then
                Combo9.ListIndex = i
                Exit For
            End If
        Next i
    Else
        Combo9.ListIndex = -1
    End If
        
    If Trim(RsTemp.Fields("Metas").Value) <> "" Then
        For i = 0 To Combo8.ListCount - 1
            If RsTemp.Fields("Metas").Value = Combo8.List(i) Then
                Combo8.ListIndex = i
                Exit For
            End If
        Next i
        Combo8.Text = RsTemp.Fields("Metas").Value
    Else
        Combo8.ListIndex = -1
    End If
Else
    TxtExamFIni.Text = ""
    TxtExamFin.Text = ""
    TxtDiagFin.Text = ""
    TxtAnatFin.Text = ""
    TxtCompliFin.Text = ""
    TxtInidiceFin.Text = ""
    TxtTratamientoFin.Text = ""
    TxtDosisT.Text = ""
    TxtDosisD.Text = ""
    TxtSesionesFin.Text = ""
    Combo9.ListIndex = -1
    Combo8.Text = ""
    Combo8.ListIndex = -1
End If
End Sub

Sub Limpiar_MGral()
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
    Combo6.ListIndex = -1
    Combo7.ListIndex = -1
    TxtGleason.Text = ""
    ChkGleason.Value = 0
    OptBordes(0).Value = True
    OptBordes(0).Value = False
    
    For ii = 0 To ChkGlobal.Count - 1
        ChkGlobal(ii).Value = 0
        If ii <= 56 Then CboGeneral(ii).ListIndex = -1
    Next ii
    ChkGlobal(57).Value = 0
    TxtOtros.Text = ""
    
    
 ' Informe Final
 
    TxtExamFIni.Text = ""
    TxtExamFin.Text = ""
    TxtDiagFin.Text = ""
    TxtAnatFin.Text = ""
    TxtCompliFin.Text = ""
    TxtInidiceFin.Text = ""
    TxtTratamientoFin.Text = ""
    TxtDosisT.Text = ""
    TxtDosisD.Text = ""
    TxtSesionesFin.Text = ""
    Combo9.ListIndex = -1
    Combo8.ListIndex = -1
        
End Sub

Sub Leer_Tipos_Ca()
On Error Resume Next

CSql = "SELECT * FROM Tipos_Ca"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then MsgBox "No existen tipos de cancer registrados en la base de datos!", vbInformation + vbOKOnly, "Información": Exit Sub

CboTCancers.Clear

While Not RsTemp.EOF
    CboTCancers.AddItem RsTemp.Fields("Descrip_Tipos_ca").Value
    CboTCancers.ItemData(CboTCancers.NewIndex) = RsTemp.Fields("Id_Tipos_ca").Value
    RsTemp.MoveNext
Wend

End Sub

Sub carga_datos_radio()
On Error GoTo WrtError
'IdLIdInf = IdLDefault
IdInf = ""
Limpiar_MGral

If RsInformeMed.RecordCount <> 0 Then

    If Not IsNull(RsInformeMed.Fields("Dosis").Value) Then Text8.Text = Val(RsInformeMed.Fields("Dosis").Value) Else Text8.Text = ""
    d1 = Val(RsInformeMed.Fields("Dosis").Value)
    If Not IsNull(RsInformeMed.Fields("Dosisd").Value) Then Text9.Text = Val(RsInformeMed.Fields("Dosisd").Value) Else Text9.Text = ""
    d2 = Val(RsInformeMed.Fields("Dosisd").Value)
    If Not IsNull(RsInformeMed.Fields("antecedente_flia").Value) Then Text15.Text = RsInformeMed.Fields("antecedente_flia").Value Else Text15.Text = ""
    d3 = Trim(RsInformeMed.Fields("antecedente_flia").Value)
    If Not IsNull(RsInformeMed.Fields("enfermedad_act").Value) Then Text16.Text = RsInformeMed.Fields("enfermedad_act").Value Else Text16.Text = ""
    d4 = Trim(RsInformeMed.Fields("enfermedad_act").Value)
    If Not IsNull(RsInformeMed.Fields("Cuantas").Value) Then Text17.Text = RsInformeMed.Fields("Cuantas").Value Else Text17.Text = ""
    d5 = Trim(RsInformeMed.Fields("Cuantas").Value)
    If Not IsNull(RsInformeMed.Fields("Motivo_Con").Value) Then Text18.Text = RsInformeMed.Fields("Motivo_Con").Value Else Text18.Text = ""
    d6 = Trim(RsInformeMed.Fields("Motivo_Con").Value)
    If Not IsNull(RsInformeMed.Fields("anatomia_patol").Value) Then Text19.Text = RsInformeMed.Fields("anatomia_patol").Value Else Text19.Text = ""
    d7 = Trim(RsInformeMed.Fields("anatomia_patol").Value)
    If Not IsNull(RsInformeMed.Fields("Examen_Fis").Value) Then Text20.Text = RsInformeMed.Fields("Examen_Fis").Value Else Text20.Text = ""
    d8 = Trim(RsInformeMed.Fields("Examen_Fis").Value)
    If Not IsNull(RsInformeMed.Fields("Diagnotico").Value) Then Text21.Text = RsInformeMed.Fields("Diagnotico").Value Else Text21.Text = ""
    d9 = Trim(RsInformeMed.Fields("Diagnotico").Value)
    If Not IsNull(RsInformeMed.Fields("Tratamiento").Value) Then Text22.Text = RsInformeMed.Fields("Tratamiento").Value Else Text22.Text = ""
    d10 = Trim(RsInformeMed.Fields("Tratamiento").Value)
    
    If Not IsNull(RsInformeMed.Fields("Metas").Value) Then CboMetas.Text = RsInformeMed.Fields("Metas").Value Else CboMetas.Text = ""
    dm = Trim(RsInformeMed.Fields("Metas").Value)
    
    If Not IsNull(RsInformeMed.Fields("Sesiones").Value) Then Text2.Text = RsInformeMed.Fields("Sesiones").Value Else Text2.Text = ""
    ds = RsInformeMed.Fields("Sesiones").Value
    
    Label13.Caption = "Registro: " & RsInformeMed.AbsolutePosition & " / " & RsInformeMed.RecordCount
    
    DtpFecha.Value = Trim(RsInformeMed.Fields("fecha").Value)
    d11 = RsInformeMed.Fields("fecha")
    
        IdReg = Trim(RsInformeMed.Fields("IdInforme").Value)
    IdInf = Trim(RsInformeMed.Fields("IdInforme").Value)
    IdLIdInf = RsInformeMed.Fields("IdL").Value
    
    Text15.ToolTipText = IdInf
    
    d14 = IdReg
    
    
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' Carga los algunos datos al informe final
    
    If Not IsNull(RsInformeMed.Fields("Dosis").Value) Then TxtDosisT.Text = Val(RsInformeMed.Fields("Dosis").Value) Else TxtDosisT.Text = ""
      
    If Not IsNull(RsInformeMed.Fields("Dosisd").Value) Then TxtDosisD.Text = Val(RsInformeMed.Fields("Dosisd").Value) Else TxtDosisD.Text = ""
       
    'If Not IsNull(RsInformeMed.Fields("Cuantas").Value) Then Text17.Text = RsInformeMed.Fields("Cuantas").Value Else Text17.Text = ""
   
    If Not IsNull(RsInformeMed.Fields("Examen_Fis").Value) Then TxtExamFIni.Text = RsInformeMed.Fields("Examen_Fis").Value Else TxtExamFIni.Text = ""
        
    If Not IsNull(RsInformeMed.Fields("Diagnotico").Value) Then TxtDiagFin.Text = RsInformeMed.Fields("Diagnotico").Value Else TxtDiagFin.Text = ""
    
    If Not IsNull(RsInformeMed.Fields("Tratamiento").Value) Then TxtTratamientoFin.Text = RsInformeMed.Fields("Tratamiento").Value Else TxtTratamientoFin.Text = ""
    
    If Not IsNull(RsInformeMed.Fields("Metas").Value) Then Combo8.Text = RsInformeMed.Fields("Metas").Value Else Combo8.Text = ""
       
    If Not IsNull(RsInformeMed.Fields("Sesiones").Value) Then TxtSesionesFin.Text = RsInformeMed.Fields("Sesiones").Value Else TxtSesionesFin.Text = ""
   
       
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    CSql = "SELECT IdTipoCancer FROM Informe_Medico2 WHERE IdInforme=" & IdInf
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then
        If Not IsNull(RsTemp.Fields("IdTipoCancer").Value) Then
            For ii = 0 To CboTCancers.ListCount - 1
                If CboTCancers.ItemData(ii) = Val(RsTemp.Fields("IdTipoCancer").Value) Then
                    CboTCancers.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    If Not IsNull(RsInformeMed.Fields("Idmedicot").Value) Then
        d15 = Val(RsInformeMed.Fields("Idmedicot").Value)
        
        For i = 0 To CboModificarMedicoTratante.ListCount - 1
            If CboModificarMedicoTratante.ItemData(i) = d15 Then
                CboModificarMedicoTratante.ListIndex = i
                Exit For
            End If
        Next i
    End If
            'roberto
    If IsNull(RsInformeMed.Fields("Tomografia").Value) Then
        Combo1.ListIndex = 0
        Text17.Text = "0"
        d12 = 0
        d13 = 0
    Else
        If Val(RsInformeMed.Fields("Tomografia").Value) = 0 Then
            Combo1.ListIndex = 0
            Text17.Text = "0"
            d12 = 0
            d13 = 0
        Else
            Combo1.ListIndex = 1
            Text17.Text = Val(RsInformeMed.Fields("Cuantas").Value)
            d12 = 1
            d13 = Val(RsInformeMed.Fields("Cuantas").Value)
        End If
    End If
    
       If Not IsNull(RsInformeMed.Fields("Estadiaje").Value) Then
        ss = InStr(1, RsInformeMed.Fields("Estadiaje").Value, "(", vbTextCompare)
        If ss < 1 Then ss = 1
        stt = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss - 1)
        Combo2.Text = Trim(Mid(stt, 3))
        
        ss1 = InStr(1, RsInformeMed.Fields("Estadiaje").Value, ".", vbTextCompare)
        If ss1 < 1 Then ss1 = 1
        stt1 = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss1 - 1)
        Combo3.Text = Trim(Mid(stt1, ss + 1))
        
        ss2 = InStr(ss1 + 1, RsInformeMed.Fields("Estadiaje").Value, ".", vbTextCompare)
        If ss2 < 1 Then ss2 = 1
        stt2 = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss2 - 1)
        Combo4.Text = Trim(Mid(stt2, ss1 + 1))
        
        ss3 = InStr(ss2 + 1, RsInformeMed.Fields("Estadiaje").Value, ".", vbTextCompare)
        If ss3 < 1 Then ss3 = 1
        stt3 = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss3 - 1)
        Combo5.Text = Trim(Mid(stt3, ss2 + 1))
        
        ss4 = InStr(ss3 + 1, RsInformeMed.Fields("Estadiaje").Value, ".", vbTextCompare)
        If ss4 < 1 Then ss4 = 1
        stt4 = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss4 - 1)
        Combo6.Text = Trim(Mid(stt4, ss3 + 1))
        
        ss5 = InStr(ss4 + 1, RsInformeMed.Fields("Estadiaje").Value, ".", vbTextCompare)
        If ss5 < 1 Then ss5 = 1
        stt5 = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss5 - 1)
        Combo7.Text = Trim(Mid(stt5, ss4 + 1))
        
        ss6 = InStr(ss6 + 1, RsInformeMed.Fields("Estadiaje").Value, ")", vbTextCompare)
        If ss6 < 1 Then ss6 = 1
        stt6 = Mid(RsInformeMed.Fields("Estadiaje").Value, 1, ss6 - 1)
        TxtGleason.Text = Trim(Mid(stt6, ss5 + 1))
        
        de = RsInformeMed.Fields("Estadiaje").Value
    Else
        Combo2.Text = "I."
        Combo3.Text = "C."
        Combo4.Text = "T."
        Combo5.Text = "N."
        Combo6.Text = "M."
        Combo6.Text = "G."
        Combo6.Text = "Gleason"
    End If
    
    Cambio = 0
    actualiza = 1
    CboModificarMedicoTratante.Enabled = True
    Call Habilita_Btns("Todos")
    DesactivarTextos
    
'    Frame2.Enabled = True
'    Frame4(0).BackColor = &HE0E0E0
'    Frame5(0).BackColor = &HE0E0E0
'    Frame5(1).BackColor = &HE0E0E0
    CARGAR_INFORME2
Else
    IdReg = ""
    Cambio = 0
    actualiza = 0
    BtnEliminar.Enabled = False
    CboModificarMedicoTratante.Enabled = False
    'BtnGuardarActualizar.Enabled = False
    ActivarTextos
End If
Cambio = 0

Exit Sub
WrtError:

MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Sub Blanqueo_General()
On Error GoTo WrtError
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Label12.Caption = ""
CboTCancers.ListIndex = -1

Text15.Text = ""
Text19.Text = ""
Text16.Text = ""
Text17.Text = ""
Text20.Text = ""
Text18.Text = ""
Text21.Text = ""
Text8.Text = ""
Text9.Text = ""
Text22.Text = ""
Label13.Caption = "Registro 0 / 0  (Sin Informe Médico)"
NoReg.Caption = "Registro 0 / 0"

Exit Sub

WrtError:

MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Sub Blanqueo()
On Error GoTo WrtError
'BtnAnterior1.Enabled = False
'BtnSiguiente2.Enabled = False
'BtnGuardarActualizar.Enabled = False
'BtnEliminar.Enabled = False
CboTCancers.ListIndex = -1
Text15.Text = ""
Text19.Text = ""
Text16.Text = ""
Text17.Text = ""
Text20.Text = ""
Text18.Text = ""
Text21.Text = ""
Text8.Text = ""
Text9.Text = ""
Text22.Text = ""
Combo1.ListIndex = -1
CboMetas.Text = ""
DtpFecha.Value = Date
Label13.Caption = "Registro 0 / 0  (Sin Informe Médico)"
Limpiar_MGral
Exit Sub

WrtError:

MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Sub blanqueo1()
On Error Resume Next
    DtpFecha.Value = Date
    Label13.Caption = "Registro NUEVO"
End Sub

Private Sub BtnAgregar_Click()
On Error Resume Next
'Unload Me
If Trim(IdPac1) = "" Then MsgBox "Debe seleccionar un Paciente antes de agregar un Informe Medico!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub
IO = 1
Call Deshabilita_Btns
Call blanqueo1
Cambio = 0
actualiza = 0
Frame1.Enabled = False
Frame2.BackColor = &HE0E0E0
Frame4(0).BackColor = &HE0E0E0
Frame5(0).BackColor = &HE0E0E0
Frame5(1).BackColor = &HE0E0E0
CboModificarMedicoTratante.Enabled = True
DesactivarTextos
OptInformeMedico(0).Value = False
OptInformeMedico(0).Value = True
DtpFecha.Value = Now

For i = 0 To FrameDato.Count - 1
    FrameDato(i).BackColor = &HE0E0E0
Next i

For i = 0 To OptBordes.Count - 1
    OptBordes(i).BackColor = &HE0E0E0
Next i

For i = 0 To ChkMuert.Count - 1
    ChkMuert(i).BackColor = &HE0E0E0
Next i
For i = 0 To ChkRecaida.Count - 1
    ChkRecaida(i).BackColor = &HE0E0E0
Next i
For i = 0 To LstProg.Count - 1
    LstProg(i).BackColor = &HE0E0E0
Next i
For i = 0 To LstEnfer.Count - 1
    LstEnfer(i).BackColor = &HE0E0E0
Next i

For i = 0 To ChkGlobal.Count - 1
    ChkGlobal(i).BackColor = &HE0E0E0
Next i

Frame10.BackColor = &HE0E0E0
Frame8.BackColor = &HE0E0E0
ChkGleason.BackColor = &HE0E0E0

End Sub

Sub ActivarTextos()
Text8.Locked = True
Text9.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
Text19.Locked = True
Text20.Locked = True
Text21.Locked = True
Text22.Locked = True

CboModificarMedicoTratante.Locked = True
CboMetas.Locked = True
Combo1.Locked = True
Combo2.Locked = True
Combo3.Locked = True
Combo4.Locked = True
Combo5.Locked = True
Combo6.Locked = True
DtpFecha.Enabled = True
End Sub

Sub DesactivarTextos()
Text8.Locked = False
Text9.Locked = False
Text15.Locked = False
Text16.Locked = False
Text17.Locked = False
Text18.Locked = False
Text19.Locked = False
Text20.Locked = False
Text21.Locked = False
Text22.Locked = False
CboModificarMedicoTratante.Locked = False
CboMetas.Locked = False
Combo1.Locked = False
Combo2.Locked = False
Combo3.Locked = False
Combo4.Locked = False
Combo5.Locked = False
Combo6.Locked = False
DtpFecha.Enabled = False
End Sub
'Sub Deshabilita_Btns(cad As String)
Sub Deshabilita_Btns()
On Error Resume Next
'If UCase(Cad) = UCase("Todos") Then
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = False
    BtnVerInforme.Enabled = False
    BtnVerHistoria.Enabled = False
    BtnEvolucionOncologica.Enabled = False
    BtnAnterior1.Enabled = False
    BtnSiguiente2.Enabled = False
    BtnExamenes.Enabled = False
    BtnAntecedentes.Enabled = False
'End If
End Sub

Sub Habilita_Btns(Cad As String)
On Error Resume Next
If UCase(Cad) = UCase("Todos") Then
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = True
    BtnVerInforme.Enabled = True
    BtnVerHistoria.Enabled = True
    BtnEvolucionOncologica.Enabled = True
    BtnAnterior1.Enabled = True
    BtnSiguiente2.Enabled = True
    BtnExamenes.Enabled = True
    BtnAntecedentes.Enabled = True
ElseIf UCase(Cad) = UCase("sin informe") Then
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnVerInforme.Enabled = False
    BtnEvolucionOncologica.Enabled = False
    BtnVerHistoria.Enabled = True
    BtnAnterior1.Enabled = False
    BtnSiguiente2.Enabled = False
    BtnExamenes.Enabled = True
    BtnAntecedentes.Enabled = True
End If
Frame2.Enabled = True
End Sub

Private Sub BtnAgregarComplica_Click()
DTPicker1.Value = Now

BtnAgregarComplica.Enabled = False
BtnCancelar.Enabled = True

TxtOtrasObs.Enabled = True
TxtOtrasObs.BackColor = vbWhite
TxtOtrasObs.Text = ""

DTPicker1.Enabled = True
BtnGuardarComplica.Enabled = True
For i = 0 To TxtDosisGrdo.Count - 1
    TxtDosisGrdo(i).Enabled = True
    TxtDosisGrdo(i).BackColor = vbWhite
    TxtDosisGrdo(i).Text = ""
Next

For i = 0 To CboGrado.Count - 1
    CboGrado(i).Enabled = True
    CboGrado(i).ListIndex = 0
Next

End Sub

Private Sub BtnAntecedentes_Click()
On Error Resume Next
FrmAntecedentes.IdPacA = IdPac1
FrmAntecedentes.IdLIdPacA = IdLIdPac
FrmAntecedentes.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAnterior_Click()
On Error Resume Next

If RsCargarPacientes.RecordCount <> 0 Then
    Call Blanqueo
    If Not RsCargarPacientes.BOF Then
        RsCargarPacientes.MovePrevious
        If RsCargarPacientes.BOF Then RsCargarPacientes.MoveLast
        Call Carga_De_Datos
        Call CONSULTA_INFORME
        If RsInformeMed.RecordCount <> 0 Then
            Call carga_datos_radio
            Cambio = 0
            actualiza = 1
            Else
            Call Habilita_Btns("sin informe")
            CboModificarMedicoTratante.Enabled = False
            Cambio = 0
            actualiza = 0
        End If
    End If
    CARGAR_INFORME2
Else
    MsgBox "No se encontraron datos, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros cargados"
End If
If OptInformeMedico(0).Value Then BtnGuardarActualizar.Enabled = Fals
End Sub

Private Sub BtnAnterior1_Click()
On Error Resume Next

If Trim(IdPac1) = "" Then Exit Sub

If Cambio = 1 Then
    Msg = "HA REALIZADO CAMBIOS EN ESTE REGISTRO!!!!!!!" & Chr(13) & "DESEA GUARDAR ESTOS CAMBIOS????"
    f = MsgBox(Msg, vbYesNo, "GUARDAR CAMBIOS??")
    If f = 6 Then Call BtnGuardarActualizar_Click
End If

Call Blanqueo
If RsInformeMed.RecordCount <> 0 Then
    If Not RsInformeMed.BOF Then
        RsInformeMed.MovePrevious
        If RsInformeMed.BOF Then RsInformeMed.MoveLast
    End If
    Call carga_datos_radio
    Cambio = 0
    actualiza = 1
    Else
    MsgBox "No existen informes medicos para este paciente!", vbExclamation + vbOKOnly, "No existen los datos"
End If
End Sub

Public Sub BtnBuscar_Click()
On Error GoTo WrtError
If Cambio = 1 And IdPac1 <> "" Then
    Msg = "HA REALIZADO CAMBIOS EN ESTE REGISTRO!!!!!!!" & Chr(13) & "DESEA GUARDAR ESTOS CAMBIOS????"
    f = MsgBox(Msg, vbYesNo, "GUARDAR CAMBIOS??")
    If f = 6 Then
        Call BtnGuardarActualizar_Click
    End If
End If

'BtnDesHacer_Click
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Limpiar Campos

ActivarTextos
CboModificarMedicoTratante.Enabled = False

Call Blanqueo
Call Habilita_Btns("sin informe")

Cambio = 0
actualiza = 0
Frame1.Enabled = True
Frame2.BackColor = &HEAEFEF
Frame4(0).BackColor = &HE0E0E0
Frame5(0).BackColor = &HE0E0E0
Frame5(1).BackColor = &HE0E0E0

Frame2.BackColor = &HEAEFEF
Frame4(0).BackColor = &HEAEFEF
Frame5(0).BackColor = &HEAEFEF
Frame5(1).BackColor = &HEAEFEF
For i = 0 To FrameDato.Count - 1
    FrameDato(i).BackColor = &HEAEFEF
Next i

For i = 0 To OptBordes.Count - 1
    OptBordes(i).BackColor = &HEAEFEF
Next i

Frame10.BackColor = &HEAEFEF
Frame8.BackColor = &HEAEFEF
ChkGleason.BackColor = &HEAEFEF

For i = 0 To ChkMuert.Count - 1
    ChkMuert(i).BackColor = &HEAEFEF
    ChkMuert(i).Value = 0
Next i
For i = 0 To ChkRecaida.Count - 1
    ChkRecaida(i).BackColor = &HEAEFEF
    ChkRecaida(i).Value = 0
Next i
For i = 0 To LstProg.Count - 1
    LstProg(i).BackColor = &HEAEFEF
    LstProg(i).Value = 0
Next i
For i = 0 To LstEnfer.Count - 1
    LstEnfer(i).BackColor = &HEAEFEF
    LstEnfer(i).Value = 0
Next i

For i = 0 To Text5.Count - 1
    Text5(i).Text = ""
Next i

For i = 0 To ChkGlobal.Count - 1
    ChkGlobal(i).BackColor = &HEAEFEF
    ChkGlobal(i).Value = 0
Next i

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

If Trim(TxtBuscar.Text) = "" Or UCase(TxtBuscar.Text) = UCase("Busqueda") Then
    f = "Buscar"
    CSql = "select * from Paciente Order by IdPaciente"
Else
    CSql = "select * from Paciente where Historia='" & TxtBuscar.Text & "' or cedulaP = " & Val(TxtBuscar.Text) & " or nombreP like '%" & TxtBuscar.Text & "%' or ApellidoP like '%" & TxtBuscar.Text & "%'"
End If

Set RsCargarPacientes = CrearRS(CSql)

If RsCargarPacientes.RecordCount = 0 Then
    MsgBox "No Existe el registro", vbExclamation + vbOKOnly, "No hay datos"
    NoReg = "Registro 0 / 0"
    IdPac1 = ""
    IdLIdPac = IdLDefault
    IdLIdInf = IdLDefault
    IdInf = ""
    Blanqueo_General
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnVerHistoria.Enabled = False
    BtnEvolucionOncologica.Enabled = False
    BtnVerInforme.Enabled = False
    BtnAnterior1.Enabled = False
    BtnSiguiente2.Enabled = False
    CboModificarMedicoTratante.Enabled = False
    
    
    Exit Sub
    'CSql = "select * from Paciente "
    'Set RsCargarPacientes = CrearRS(CSql)
    'RsCargarPacientes.MoveFirst
End If

    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = True
    BtnVerHistoria.Enabled = True
    BtnEvolucionOncologica.Enabled = True
    BtnVerInforme.Enabled = True
    BtnAnterior1.Enabled = True
    BtnSiguiente2.Enabled = True
    
Call Carga_De_Datos
Call CONSULTA_INFORME

If Not (RsInformeMed.EOF) Then
    Call carga_datos_radio
    Cambio = 0
    actualiza = 1
Else
    Call Blanqueo
    Call Habilita_Btns("sin informe")
    Cambio = 0
    actualiza = 0
End If

Exit Sub

WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub BtnCancelar_Click()


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

  TxtOtrasObs.Text = ""
  TxtOtrasObs.Enabled = False
  TxtOtrasObs.BackColor = &HE0E0E0

  DTPicker1.Enabled = False
  BtnGuardarComplica.Enabled = False
  For i = 0 To TxtDosisGrdo.Count - 1
      TxtDosisGrdo(i).Enabled = False
      TxtDosisGrdo(i).Text = ""
      TxtDosisGrdo(i).BackColor = &HE0E0E0
  Next
  For i = 0 To CboGrado.Count - 1
      CboGrado(i).Enabled = False
      CboGrado(i).ListIndex = 0
  Next

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

BtnAgregarComplica.Enabled = True
BtnGuardarComplica.Enabled = False
BtnCancelar.Enabled = False
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next

ActivarTextos
CboModificarMedicoTratante.Enabled = False

Call Blanqueo
Call Habilita_Btns("sin informe")

Cambio = 0
actualiza = 0
Frame1.Enabled = True
Frame2.BackColor = &HEAEFEF
Frame4(0).BackColor = &HE0E0E0
Frame5(0).BackColor = &HE0E0E0
Frame5(1).BackColor = &HE0E0E0

Frame2.BackColor = &HEAEFEF
Frame4(0).BackColor = &HEAEFEF
Frame5(0).BackColor = &HEAEFEF
Frame5(1).BackColor = &HEAEFEF
For i = 0 To FrameDato.Count - 1
    FrameDato(i).BackColor = &HEAEFEF
Next i


For i = 0 To OptBordes.Count - 1
    OptBordes(i).BackColor = &HEAEFEF
Next i

Frame10.BackColor = &HEAEFEF
Frame8.BackColor = &HEAEFEF
ChkGleason.BackColor = &HEAEFEF

For i = 0 To ChkMuert.Count - 1
    ChkMuert(i).BackColor = &HEAEFEF
    ChkMuert(i).Value = 0
Next i
For i = 0 To ChkRecaida.Count - 1
    ChkRecaida(i).BackColor = &HEAEFEF
    ChkRecaida(i).Value = 0
Next i
For i = 0 To LstProg.Count - 1
    LstProg(i).BackColor = &HEAEFEF
    LstProg(i).Value = 0
Next i
For i = 0 To LstEnfer.Count - 1
    LstEnfer(i).BackColor = &HEAEFEF
    LstEnfer(i).Value = 0
Next i

For i = 0 To Text5.Count - 1
    Text5(i).Text = ""
Next i

For i = 0 To ChkGlobal.Count - 1
    ChkGlobal(i).BackColor = &HEAEFEF
    ChkGlobal(i).Value = 0
Next i


If IdPac1 = "" Then
    CSql = "select * from Paciente"
    Set RsCargarPacientes = CrearRS(CSql)
    RsCargarPacientes.MoveFirst
End If

Carga_De_Datos

Call CONSULTA_INFORME
If RsInformeMed.RecordCount <> 0 Then
    Call carga_datos_radio
    Cambio = 0
    actualiza = 1
    Else
    Call Habilita_Btns("sin informe")
    Cambio = 0
    actualiza = 0
End If

End Sub

Private Sub BtnDesocuparAlPacienteAtendido_Click()
On Error Resume Next
Dim bdlista88 As New ADODB.Recordset
CSql = "Delete From Ubi_Paciente Where Modul = " & ModulO
Set bdlista88 = CrearRS(CSql)
End Sub

Private Sub BtnEliminar_Click()
On Error GoTo WrtError
Dim RsDeshabilitar As New ADODB.Recordset
Dim RsVerificar As New ADODB.Recordset

If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente antes de Eliminar!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub

CSql = "Select * From Informe_Medico Where IdInforme='" & IdReg & "'"
Set RsVerificar = CrearRS(CSql)

If RsVerificar.Fields("IdUsuario").Value = IdUser Then

    resp = MsgBox("Desea eliminar el Informe Medico # " & RsInformeMed.AbsolutePosition & ", del Paciente de C.I.: " & Text1.Text & " ?", vbQuestion + vbYesNo, "Confirmar!")
    
    If resp = 7 Then Exit Sub
    
    Call Enviar_Bitacora(IdUser, "Oncologia", "BORRAR", "Se elimino el INFORME MEDIDO cuya IdInforme es (" & IdReg & ")")
    
    CSql = "update INFORME_MEDICO set Estado=2 where IdInforme = " & IdReg & " And IdL='" & IdLIdInf & "'"
    Set RsDeshabilitar = CrearRS(CSql)
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Borrar Informe Médico"
    
    Call EnviarRegPendiente(IdReg, IdLIdInf, "INFORME_MEDICO", "IdInforme=" & IdReg & " And IdL='" & IdLIdInf & "'")
    
    MsgBox "El Informe Medico del Paciente: " & Text3.Text & " " & Text4.Text & " ha sido eliminado del Registro!", vbInformation + vbOKOnly, "Operacion Exitosa!"
    BtnDesHacer_Click

Else
    MsgBox "Usted No tiene permiso para borrar este informe medico", vbCritical + vbOKOnly, "Error"

End If

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Sub BorrarHosting()
On Error GoTo WrtError
ConectarHosting
CSql = "Update INFORME_MEDICO set Estado=2 where IdInforme = " & IdReg
Set RsWeb = CrearRsWeb(CSql)
WebCnn.Close

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub


Sub BorrarRegPendiente()
On Error GoTo WrtError

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

a = 1
CSql = "UPDATE Informe_Medico Set Estado='2' Where IdInforme ='" & IdReg & "'"
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Oncologia"
RsRegPendiente.Fields("Tabla").Value = "Informe_Medico"
RsRegPendiente.Fields("Condicional").Value = "IdInforme = " & IdReg
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub BtnEliminarComplica_Click()
Dim Rsp As Byte

If DMGrid1.Rows = 0 Then
    MsgBox "No existen registros para eliminar!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
ElseIf DMGrid1.Row = 0 Then
    MsgBox "Debe seleccionar un registro para eliminar!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

Rsp = MsgBox("Se procederá a eliminar el registro seleccionado, Desea continuar?", vbQuestion + vbYesNo, "Confirmación")

If Rsp = vbNo Then Exit Sub

CSql = "DELETE FROM Informe_Medico5 WHERE Id=" & Val(DMGrid1.ValorCelda(DMGrid1.Row, 3))
Set RsTemp = CrearRS(CSql)

MsgBox "El registro ha sido eliminado!", vbInformation + vbOKOnly, "Operación Exitosa!"
CARGAR_COMPLICA

End Sub

Private Sub BtnEvolucionOncologica_Click()
'On Error Resume Next
Especia = "Radioterapia"
FrmEvolucion.IdPacE = IdPac1
FrmEvolucion.IdLIdPacE = IdLIdPac
FrmEvolucion.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnExamenIngreso_Click()
On Error Resume Next

If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub
    FrmExamenFisico.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim RsGuardar As New ADODB.Recordset
On Error GoTo WrtError
Cambio = 1
If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente antes de Guardar!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub

If Trim(Text15.Text) = "" Then
    'f = "Antecedente_Flia"
    MsgBox "El Campo Antecedente Familiar esta Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    Text15.SetFocus
    Exit Sub
ElseIf Trim(Text19.Text) = "" Then
    'f = "Anatomia_Patol"
    MsgBox "El Campo Antecedente Patologico esta Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    Text19.SetFocus
    Exit Sub
ElseIf Trim(Text16.Text) = "" Then
    'f = "Enfermedad_Act"
    MsgBox "El Campo Enfermedad actual esta Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    Text16.SetFocus
    Exit Sub
'ElseIf (DateValue(DtpFecha.Value) - DateValue(Now)) < 0 Then
'    MsgBox "La fecha ingresada es Incorreta!", vbCritical + vbOKOnly, "Error"
'    DtpFecha.SetFocus
'    Exit Sub

ElseIf Trim(Text20.Text) = "" Then
    'f = "Examen_Fis"
    MsgBox "El Examen Fisico no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text20.SetFocus
    Exit Sub
ElseIf Text18.Text = "" Then
    'f = "Motivo_Con"
    MsgBox "El Motivo de la Consulta no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text18.SetFocus
    Exit Sub
ElseIf Text21.Text = "" Then
    'f = "Diagnotico"
    MsgBox "El Diagnóstico no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text21.SetFocus
    Exit Sub
ElseIf Text22.Text = "" Then
    'f = "Tratamiento"
    MsgBox "El Tratamiento no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text22.SetFocus
    Exit Sub
ElseIf Trim(Combo1.List(Combo1.ListIndex)) = "" Then
    'f = "Tomografia"
    MsgBox "La Tomografia no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Combo1.SetFocus
    Exit Sub
ElseIf Trim(CboMetas.Text) = "" Then
    'f = "Tomografia"
    MsgBox "Debe de seleccionar la Meta del Tratamiento del Paciente!", vbExclamation + vbOKOnly, "Faltan Datos"
    CboMetas.SetFocus
    Exit Sub
ElseIf CboTCancers.ListIndex < 0 Then
    MsgBox "Debe de seleccionar un tipo de CANCER!", vbExclamation + vbOKOnly, "Faltan Datos"
    CboTCancers.SetFocus
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  Open "Y:\RadioterapiaInf.txt" For Append As #1
  MError = "MMMMMMMMMMMMMMMMMMMM" & Date & " : " & Time & "MMMMMMMMMMMMMMMM" & Chr(13) & Chr(10) & _
        " RsGuardar.Fields('IdInforme').Value = " & IdInf & Chr(13) & Chr(10) & "RsGuardar.Fields('IdL').Value = " & IdLIdInf & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('IdMedicot').Value = " & IdMedT & Chr(13) & Chr(10) & "RsGuardar.Fields('IdPaciente').Value = " & IdPac1 & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('IdLIdPac').Value = " & IdLIdPac & Chr(13) & Chr(10) & "RsGuardar.Fields('IdUsuario').Value = " & IdUser & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('Antecedente_Flia').Value = " & Trim(Text15.Text) & Chr(13) & Chr(10) & "RsGuardar.Fields('Anatomia_Patol').Value = " & Trim(Text19.Text) & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('Enfermedad_Act').Value = " & Trim(Text16.Text) & Chr(13) & Chr(10) & "RsGuardar.Fields('Examen_Fis').Value = " & Trim(Text20.Text) & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('Motivo_Con').Value = " & Trim(Text18.Text) & Chr(13) & Chr(10) & "RsGuardar.Fields('Diagnotico').Value = " & Trim(Text21.Text) & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('Tratamiento').Value = " & Trim(Text22.Text) & Chr(13) & Chr(10) & "RsGuardar.Fields('Fecha').Value = " & Format(DtpFecha.Value, "dd/MM/yyyy") & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('dosis').Value = " & Val(Text8.Text) & Chr(13) & Chr(10) & "RsGuardar.Fields('dosisd').Value = " & Val(Text9.Text) & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('Tomografia').Value = " & tomog & Chr(13) & Chr(10) & "RsGuardar.Fields('Cuantas').Value = " & Val(Text17.Text) & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('Estado').Value = 1 " & Chr(13) & Chr(10) & "RsGuardar.Fields('Metas').Value = " & Trim(CboMetas.Text) & Chr(13) & Chr(10) & _
        "RsGuardar.Fields('sesiones').Value = " & Sesiones & Chr(13) & Chr(10) & "RsGuardar.Fields('Estadiaje').Value = " & Estadiaje
       
  Print #1, MError
  Close #1
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

' Obtiene el Nuevo ID para el informe medico
    Dim RsMaxReg As New ADODB.Recordset
    CSql = "Select max(IdInforme)+1 as MaxReg From Informe_Medico"
    Set RsMaxReg = CrearRS(CSql)

    If Not IsNull(RsMaxReg.Fields("MaxReg")) Then
        MaxReg = RsMaxReg.Fields("MaxReg")
    Else
        MaxReg = "1"
    End If
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

'  Bloque que verifica si hay conexion con el internet...
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' --------------------------------------------------------
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


' Obtiene el ID del medico
    If CboModificarMedicoTratante.ListIndex = -1 Then
        IdMedT = "0"
    Else
        IdMedT = CboModificarMedicoTratante.ItemData(CboModificarMedicoTratante.ListIndex)
        CSql = "update PACIENTE set Medico_tratante=" & IdMedT & " where IdPaciente = " & IdPac1
        Set RsGuardar = CrearRS(CSql)
        
        CSql = "select * from Paciente"
        Set RsCargarPacientes = CrearRS(CSql)
        RsCargarPacientes.Find "IdPaciente=" & IdPac1
    End If
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm

    If Cambio = 1 Then
        If Combo1.ListIndex = -1 Then
            tomog = 0
        ElseIf Combo1.ListIndex = 0 Then
            tomog = 0
        Else
            tomog = 1
        End If
        
         
        If Text8.Text = "" And Text9.Text = "" Then
            Text2.Text = 0
            Sesiones = 0
        ElseIf Text8.Text <> "" And Text9.Text = "" Then
            Text2.Text = CDbl(Text8.Text)
            Sesiones = 0
        ElseIf Text8.Text = "" And Text9.Text <> "" Then
            Text2.Text = CDbl(Text9.Text)
            Sesiones = 0
        ElseIf Text8.Text <> "" And Text9.Text <> "" Then
            Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
            Sesiones = Text2.Text
        ElseIf Text8.Text <> 0 And Text9.Text <> 0 Then
            Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
            Sesiones = Text2.Text
         ElseIf Text8.Text = 0 And Text9.Text = 0 Then
            Sesiones = 0
            Text2.Text = 0
        ElseIf Text8.Text = 0 And Text9.Text <> 0 Then
            Text2.Text = 0
            Sesiones = 0
        ElseIf Text8.Text <> 0 And Text9.Text = 0 Then
            Text2.Text = 0
            Sesiones = 0
        End If
        
        If actualiza = 1 Then
            
            If IdReg = "" Then MsgBox "No hay informes seleccionados!", vbExclamation + vbOKOnly, "Error": Exit Sub
            
            CSql = "Select * From Informe_Medico Where IdPaciente='" & IdPac1 & "' And IdInforme='" & IdInf & "' And IdL='" & IdLIdInf & "' And IdLIdPac='" & IdLIdPac & "'"
            Set RsGuardar = CrearRS(CSql)
            
            If RsGuardar.RecordCount <> 0 Then
                IdInf = RsGuardar.Fields("IdInforme").Value
                IdLIdInf = RsGuardar.Fields("IdL").Value
            End If
            
        ElseIf actualiza = 0 Then
            CSql = "Select * From Informe_Medico"
            Set RsGuardar = CrearRS(CSql)
            
            IdLIdInf = NuevoIdL
            IdInf = MaxReg
            
            RsGuardar.AddNew
            RsGuardar.Fields("IdInforme").Value = IdInf
            RsGuardar.Fields("IdL").Value = IdLIdInf
            RsGuardar.Fields("IdMedicot").Value = IdMedT
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdLIdPac").Value = IdLIdPac
            
        End If
        
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        RsGuardar.Fields("Antecedente_Flia").Value = Trim(Text15.Text)
        RsGuardar.Fields("Anatomia_Patol").Value = Trim(Text19.Text)
        RsGuardar.Fields("Enfermedad_Act").Value = Trim(Text16.Text)
        RsGuardar.Fields("Examen_Fis").Value = Trim(Text20.Text)
        RsGuardar.Fields("Motivo_Con").Value = Trim(Text18.Text)
        RsGuardar.Fields("Diagnotico").Value = Trim(Text21.Text)
        RsGuardar.Fields("Tratamiento").Value = Trim(Text22.Text)
        RsGuardar.Fields("Fecha").Value = Format(DtpFecha.Value, "dd/MM/yyyy")
        RsGuardar.Fields("dosis").Value = Val(Text8.Text)
        RsGuardar.Fields("dosisd").Value = Val(Text9.Text)
        RsGuardar.Fields("Tomografia").Value = tomog
        RsGuardar.Fields("Cuantas").Value = Val(Text17.Text)
        RsGuardar.Fields("Estado").Value = 1
        RsGuardar.Fields("Metas").Value = Trim(CboMetas.Text)
        RsGuardar.Fields("sesiones").Value = Sesiones
        RsGuardar.Fields("Estadiaje").Value = Estadiaje
        
        RsGuardar.Update
        
        Call Habilita_Btns("Todos")
        Frame2.BackColor = &HEAEFEF
        Frame4(0).BackColor = &HE0E0E0
        Frame5(0).BackColor = &HE0E0E0
        Frame5(1).BackColor = &HE0E0E0
        
        MsgBox "Registro actualizado Satisfactoriamente, se procedera a guardar los datos estadisticos!s", vbInformation + vbOKOnly, "Operacion Exitosa"
        
        ' Envia el registro a la tabla de REGISTROS PENDIENTES!
        Call EnviarRegPendiente(IdInf, IdLIdInf, "Informe_Medico", "IdInforme=" & IdInf & " And IdL='" & IdLIdInf & "'")
        
        ActivarTextos
        Call CONSULTA_INFORME
        
        'Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        'MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Informe Médico"
'        EnviarAlHosting
'        EnviarRegPendiente
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        CSql = "SELECT * FROM Informe_Medico2 WHERE IdInforme=" & IdInf & " And IdL='" & IdLIdInf & "'"
        Set RsGuardar = CrearRS(CSql)
        
        If RsGuardar.RecordCount = 0 Then
            CSql = "SELECT MAX(Id) + 1 AS NuevoId FROM Informe_Medico2"
            Set RsGuardar = CrearRS(CSql)
            
            If Not IsNull(RsGuardar.Fields(0).Value) Then
                NuevoId = Val(RsGuardar.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If
            
            CSql = "SELECT * FROM Informe_Medico2"
            Set RsGuardar = CrearRS(CSql)
            
            NuevoId2 = NuevoIdL
            RsGuardar.AddNew
            RsGuardar.Fields("Id").Value = NuevoId
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdLIdPac").Value = IdLIdPac
            RsGuardar.Fields("IdL").Value = NuevoId2
            RsGuardar.Fields("IdLIdInf").Value = IdLIdInf
            RsGuardar.Fields("IdInforme").Value = IdInf
        Else
            NuevoId = RsGuardar.Fields("Id").Value
            NuevoId2 = RsGuardar.Fields("IdL").Value
        End If
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        RsGuardar.Fields("CP").Value = Trim(Combo3.Text)
        RsGuardar.Fields("T").Value = Trim(Combo4.Text)
        RsGuardar.Fields("N").Value = Trim(Combo5.Text)
        RsGuardar.Fields("M").Value = Trim(Combo6.Text)
        RsGuardar.Fields("Estadio").Value = Trim(Combo2.Text)
        RsGuardar.Fields("G").Value = Trim(Combo7.Text)
        If ChkGleason.Value Then RsGuardar.Fields("Gleason").Value = Trim(TxtGleason.Text) Else RsGuardar.Fields("Gleason").Value = ""
        RsGuardar.Fields("Reseccion").Value = OptBordesSel
        RsGuardar.Fields("IdTipoCancer").Value = CboTCancers.ItemData(CboTCancers.ListIndex)
        
        If ChkGlobal(0).Value Then RsGuardar.Fields("Re").Value = Trim(CboGeneral(0).Text) Else RsGuardar.Fields("Re").Value = ""
        If ChkGlobal(1).Value Then RsGuardar.Fields("Rp").Value = Trim(CboGeneral(1).Text) Else RsGuardar.Fields("Rp").Value = ""
        If ChkGlobal(2).Value Then RsGuardar.Fields("Her2Neu").Value = Trim(CboGeneral(2).Text) Else RsGuardar.Fields("Her2Neu").Value = ""
        If ChkGlobal(3).Value Then RsGuardar.Fields("EMA").Value = Trim(CboGeneral(3).Text) Else RsGuardar.Fields("EMA").Value = ""
        If ChkGlobal(4).Value Then RsGuardar.Fields("VIM").Value = Trim(CboGeneral(4).Text) Else RsGuardar.Fields("VIM").Value = ""
        If ChkGlobal(5).Value Then RsGuardar.Fields("CAE").Value = Trim(CboGeneral(5).Text) Else RsGuardar.Fields("CAE").Value = ""
        If ChkGlobal(6).Value Then RsGuardar.Fields("CERB-2").Value = Trim(CboGeneral(6).Text) Else RsGuardar.Fields("CERB-2").Value = ""
        If ChkGlobal(7).Value Then RsGuardar.Fields("P53").Value = Trim(CboGeneral(7).Text) Else RsGuardar.Fields("P53").Value = ""
        If ChkGlobal(8).Value Then RsGuardar.Fields("DESMINA").Value = Trim(CboGeneral(8).Text) Else RsGuardar.Fields("DESMINA").Value = ""
        If ChkGlobal(9).Value Then RsGuardar.Fields("ACE").Value = Trim(CboGeneral(9).Text) Else RsGuardar.Fields("ACE").Value = ""
        
        If ChkGlobal(10).Value Then RsGuardar.Fields("AFP").Value = Trim(CboGeneral(10).Text) Else RsGuardar.Fields("AFP").Value = ""
        If ChkGlobal(11).Value Then RsGuardar.Fields("PROT-S-100").Value = Trim(CboGeneral(11).Text) Else RsGuardar.Fields("PROT-S-100").Value = ""
        If ChkGlobal(12).Value Then RsGuardar.Fields("PGP").Value = Trim(CboGeneral(12).Text) Else RsGuardar.Fields("PGP").Value = ""
        If ChkGlobal(13).Value Then RsGuardar.Fields("CD31").Value = Trim(CboGeneral(13).Text) Else RsGuardar.Fields("CD31").Value = ""
        If ChkGlobal(14).Value Then RsGuardar.Fields("CD34").Value = Trim(CboGeneral(14).Text) Else RsGuardar.Fields("CD34").Value = ""
        If ChkGlobal(15).Value Then RsGuardar.Fields("CD117").Value = Trim(CboGeneral(15).Text) Else RsGuardar.Fields("CD117").Value = ""
        If ChkGlobal(16).Value Then RsGuardar.Fields("CK5").Value = Trim(CboGeneral(16).Text) Else RsGuardar.Fields("CK5").Value = ""
        If ChkGlobal(17).Value Then RsGuardar.Fields("CK6").Value = Trim(CboGeneral(17).Text) Else RsGuardar.Fields("CK6").Value = ""
        If ChkGlobal(18).Value Then RsGuardar.Fields("CK7").Value = Trim(CboGeneral(18).Text) Else RsGuardar.Fields("CK7").Value = ""
        If ChkGlobal(19).Value Then RsGuardar.Fields("CK20").Value = Trim(CboGeneral(19).Text) Else RsGuardar.Fields("CK20").Value = ""
        
        If ChkGlobal(20).Value Then RsGuardar.Fields("CAM 5,2").Value = Trim(CboGeneral(20).Text) Else RsGuardar.Fields("CAM 5,2").Value = ""
        If ChkGlobal(21).Value Then RsGuardar.Fields("TTF-1").Value = Trim(CboGeneral(21).Text) Else RsGuardar.Fields("TTF-1").Value = ""
        If ChkGlobal(22).Value Then RsGuardar.Fields("CROMOGRANINA").Value = Trim(CboGeneral(22).Text) Else RsGuardar.Fields("CROMOGRANINA").Value = ""
        If ChkGlobal(23).Value Then RsGuardar.Fields("SINAPTOFISINA").Value = Trim(CboGeneral(23).Text) Else RsGuardar.Fields("SINAPTOFISINA").Value = ""
        If ChkGlobal(24).Value Then RsGuardar.Fields("CD56").Value = Trim(CboGeneral(24).Text) Else RsGuardar.Fields("CD56").Value = ""
        If ChkGlobal(25).Value Then RsGuardar.Fields("CD57").Value = Trim(CboGeneral(25).Text) Else RsGuardar.Fields("CD57").Value = ""
        If ChkGlobal(26).Value Then RsGuardar.Fields("EGFR").Value = Trim(CboGeneral(26).Text) Else RsGuardar.Fields("EGFR").Value = ""
        If ChkGlobal(27).Value Then RsGuardar.Fields("KIT").Value = Trim(CboGeneral(27).Text) Else RsGuardar.Fields("KIT").Value = ""
        If ChkGlobal(28).Value Then RsGuardar.Fields("AE1").Value = Trim(CboGeneral(28).Text) Else RsGuardar.Fields("AE1").Value = ""
        If ChkGlobal(29).Value Then RsGuardar.Fields("AE3").Value = Trim(CboGeneral(29).Text) Else RsGuardar.Fields("AE3").Value = ""
        
        If ChkGlobal(30).Value Then RsGuardar.Fields("CK903").Value = Trim(CboGeneral(30).Text) Else RsGuardar.Fields("CK903").Value = ""
        If ChkGlobal(31).Value Then RsGuardar.Fields("GFAP").Value = Trim(CboGeneral(31).Text) Else RsGuardar.Fields("GFAP").Value = ""
        If ChkGlobal(32).Value Then RsGuardar.Fields("SMA").Value = Trim(CboGeneral(32).Text) Else RsGuardar.Fields("SMA").Value = ""
        If ChkGlobal(33).Value Then RsGuardar.Fields("CA199").Value = Trim(CboGeneral(33).Text) Else RsGuardar.Fields("CA199").Value = ""
        If ChkGlobal(34).Value Then RsGuardar.Fields("CA125").Value = Trim(CboGeneral(34).Text) Else RsGuardar.Fields("CA125").Value = ""
        If ChkGlobal(35).Value Then RsGuardar.Fields("CEA").Value = Trim(CboGeneral(35).Text) Else RsGuardar.Fields("CEA").Value = ""
        If ChkGlobal(36).Value Then RsGuardar.Fields("CEA-D14").Value = Trim(CboGeneral(36).Text) Else RsGuardar.Fields("CEA-D14").Value = ""
        If ChkGlobal(37).Value Then RsGuardar.Fields("E-CAD").Value = Trim(CboGeneral(37).Text) Else RsGuardar.Fields("E-CAD").Value = ""
        If ChkGlobal(38).Value Then RsGuardar.Fields("HCG").Value = Trim(CboGeneral(38).Text) Else RsGuardar.Fields("HCG").Value = ""
        If ChkGlobal(39).Value Then RsGuardar.Fields("HMB-45").Value = Trim(CboGeneral(39).Text) Else RsGuardar.Fields("HMB-45").Value = ""
        
        If ChkGlobal(40).Value Then RsGuardar.Fields("HPAP").Value = Trim(CboGeneral(40).Text) Else RsGuardar.Fields("HPAP").Value = ""
        If ChkGlobal(41).Value Then RsGuardar.Fields("WT1").Value = Trim(CboGeneral(41).Text) Else RsGuardar.Fields("WT1").Value = ""
        If ChkGlobal(42).Value Then RsGuardar.Fields("BEL-1").Value = Trim(CboGeneral(42).Text) Else RsGuardar.Fields("BEL-1").Value = ""
        If ChkGlobal(43).Value Then RsGuardar.Fields("BEL-2").Value = Trim(CboGeneral(43).Text) Else RsGuardar.Fields("BEL-2").Value = ""
        If ChkGlobal(44).Value Then RsGuardar.Fields("PRB").Value = Trim(CboGeneral(44).Text) Else RsGuardar.Fields("PRB").Value = ""
        If ChkGlobal(45).Value Then RsGuardar.Fields("ALK-1").Value = Trim(CboGeneral(45).Text) Else RsGuardar.Fields("ALK-1").Value = ""
        If ChkGlobal(46).Value Then RsGuardar.Fields("RA").Value = Trim(CboGeneral(46).Text) Else RsGuardar.Fields("RA").Value = ""
        If ChkGlobal(47).Value Then RsGuardar.Fields("CD99MID2").Value = Trim(CboGeneral(47).Text) Else RsGuardar.Fields("CD99MID2").Value = ""
        If ChkGlobal(48).Value Then RsGuardar.Fields("NSD").Value = Trim(CboGeneral(48).Text) Else RsGuardar.Fields("NSD").Value = ""
        If ChkGlobal(49).Value Then RsGuardar.Fields("LCACD45").Value = Trim(CboGeneral(49).Text) Else RsGuardar.Fields("LCACD45").Value = ""
        
        If ChkGlobal(50).Value Then RsGuardar.Fields("CD20L26").Value = Trim(CboGeneral(50).Text) Else RsGuardar.Fields("CD20L26").Value = ""
        If ChkGlobal(51).Value Then RsGuardar.Fields("CD79A").Value = Trim(CboGeneral(51).Text) Else RsGuardar.Fields("CD79A").Value = ""
        If ChkGlobal(52).Value Then RsGuardar.Fields("CD45ROUCHL1").Value = Trim(CboGeneral(52).Text) Else RsGuardar.Fields("CD45ROUCHL1").Value = ""
        If ChkGlobal(53).Value Then RsGuardar.Fields("CD3").Value = Trim(CboGeneral(53).Text) Else RsGuardar.Fields("CD3").Value = ""
        If ChkGlobal(54).Value Then RsGuardar.Fields("CD30KL1BERH2").Value = Trim(CboGeneral(54).Text) Else RsGuardar.Fields("CD30KL1BERH2").Value = ""
        If ChkGlobal(55).Value Then RsGuardar.Fields("CD15LEUM1").Value = Trim(CboGeneral(55).Text) Else RsGuardar.Fields("CD15LEUM1").Value = ""
        If ChkGlobal(56).Value Then RsGuardar.Fields("WT").Value = Trim(CboGeneral(56).Text) Else RsGuardar.Fields("WT").Value = ""
        If ChkGlobal(57).Value Then RsGuardar.Fields("OTROS").Value = Trim(TxtOtros.Text) Else RsGuardar.Fields("OTROS").Value = ""
        
        RsGuardar.Update
        
        Call EnviarRegPendiente(NuevoId, NuevoId2, "Informe_Medico2", "Id=" & NuevoId & " And IdL='" & NuevoId2 & "'")
        
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
       
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        
        CSql = "SELECT * FROM Informe_Medico3 WHERE IdInforme=" & IdInf & " And IdL='" & IdLIdInf & "'"
        Set RsGuardar = CrearRS(CSql)
        
        If RsGuardar.RecordCount = 0 Then
            CSql = "SELECT MAX(Id) + 1 AS NuevoId FROM Informe_Medico3"
            Set RsGuardar = CrearRS(CSql)
            
            If Not IsNull(RsGuardar.Fields(0).Value) Then
                NuevoId = Val(RsGuardar.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If
            
            CSql = "SELECT * FROM Informe_Medico3"
            Set RsGuardar = CrearRS(CSql)
            
            RsGuardar.AddNew
            
            NuevoId2 = NuevoIdL
            
            RsGuardar.Fields("Id").Value = NuevoId
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdInforme").Value = IdInf
            RsGuardar.Fields("IdLIdPac").Value = IdLIdPac
            RsGuardar.Fields("IdL").Value = NuevoId2
            RsGuardar.Fields("IdLIdInf").Value = IdLIdInf
        Else
            NuevoId = RsGuardar.Fields("Id").Value
            NuevoId2 = RsGuardar.Fields("IdL").Value
        End If
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        
        If LstEnfer(0).Value Then RsGuardar.Fields("LE_MinM").Value = Trim(Text5(0).Text) Else RsGuardar.Fields("LE_MinM").Value = ""
        If LstEnfer(1).Value Then RsGuardar.Fields("LE_6M").Value = True Else RsGuardar.Fields("LE_6M").Value = False
        If LstEnfer(2).Value Then RsGuardar.Fields("LE_12M").Value = True Else RsGuardar.Fields("LE_12M").Value = False
        If LstEnfer(3).Value Then RsGuardar.Fields("LE_18M").Value = True Else RsGuardar.Fields("LE_18M").Value = False
        If LstEnfer(4).Value Then RsGuardar.Fields("LE_24M").Value = True Else RsGuardar.Fields("LE_24M").Value = False
        If LstEnfer(5).Value Then RsGuardar.Fields("LE_30M").Value = True Else RsGuardar.Fields("LE_30M").Value = False
        If LstEnfer(6).Value Then RsGuardar.Fields("LE_36M").Value = True Else RsGuardar.Fields("LE_36M").Value = False
        If LstEnfer(7).Value Then RsGuardar.Fields("LE_42M").Value = True Else RsGuardar.Fields("LE_42M").Value = False
        If LstEnfer(8).Value Then RsGuardar.Fields("LE_48M").Value = True Else RsGuardar.Fields("LE_48M").Value = False
        If LstEnfer(9).Value Then RsGuardar.Fields("LE_54M").Value = True Else RsGuardar.Fields("LE_54M").Value = False
        If LstEnfer(10).Value Then RsGuardar.Fields("LE_60M").Value = True Else RsGuardar.Fields("LE_60M").Value = False
        If LstEnfer(11).Value Then RsGuardar.Fields("LE_MaxM").Value = Trim(Text5(1).Text) Else RsGuardar.Fields("LE_MaxM").Value = ""
        
        If LstProg(0).Value Then RsGuardar.Fields("P_MinM").Value = Trim(Text5(2).Text) Else RsGuardar.Fields("P_MinM").Value = ""
        If LstProg(1).Value Then RsGuardar.Fields("P_6M").Value = True Else RsGuardar.Fields("P_6M").Value = False
        If LstProg(2).Value Then RsGuardar.Fields("P_12M").Value = True Else RsGuardar.Fields("P_12M").Value = False
        If LstProg(3).Value Then RsGuardar.Fields("P_18M").Value = True Else RsGuardar.Fields("P_18M").Value = False
        If LstProg(4).Value Then RsGuardar.Fields("P_24M").Value = True Else RsGuardar.Fields("P_24M").Value = False
        If LstProg(5).Value Then RsGuardar.Fields("P_30M").Value = True Else RsGuardar.Fields("P_30M").Value = False
        If LstProg(6).Value Then RsGuardar.Fields("P_36M").Value = True Else RsGuardar.Fields("P_36M").Value = False
        If LstProg(7).Value Then RsGuardar.Fields("P_42M").Value = True Else RsGuardar.Fields("P_42M").Value = False
        If LstProg(8).Value Then RsGuardar.Fields("P_48M").Value = True Else RsGuardar.Fields("P_48M").Value = False
        If LstProg(9).Value Then RsGuardar.Fields("P_54M").Value = True Else RsGuardar.Fields("P_54M").Value = False
        If LstProg(10).Value Then RsGuardar.Fields("P_60M").Value = True Else RsGuardar.Fields("P_60M").Value = False
        If LstProg(11).Value Then RsGuardar.Fields("P_MaxM").Value = Trim(Text5(3).Text) Else RsGuardar.Fields("P_MaxM").Value = ""
        
        If ChkRecaida(0).Value Then RsGuardar.Fields("R_MinM").Value = Trim(Text5(4).Text) Else RsGuardar.Fields("R_MinM").Value = ""
        If ChkRecaida(1).Value Then RsGuardar.Fields("R_6M").Value = True Else RsGuardar.Fields("R_6M").Value = False
        If ChkRecaida(2).Value Then RsGuardar.Fields("R_12M").Value = True Else RsGuardar.Fields("R_12M").Value = False
        If ChkRecaida(3).Value Then RsGuardar.Fields("R_18M").Value = True Else RsGuardar.Fields("R_18M").Value = False
        If ChkRecaida(4).Value Then RsGuardar.Fields("R_24M").Value = True Else RsGuardar.Fields("R_24M").Value = False
        If ChkRecaida(5).Value Then RsGuardar.Fields("R_30M").Value = True Else RsGuardar.Fields("R_30M").Value = False
        If ChkRecaida(6).Value Then RsGuardar.Fields("R_36M").Value = True Else RsGuardar.Fields("R_36M").Value = False
        If ChkRecaida(7).Value Then RsGuardar.Fields("R_42M").Value = True Else RsGuardar.Fields("R_42M").Value = False
        If ChkRecaida(8).Value Then RsGuardar.Fields("R_48M").Value = True Else RsGuardar.Fields("R_48M").Value = False
        If ChkRecaida(9).Value Then RsGuardar.Fields("R_54M").Value = True Else RsGuardar.Fields("R_54M").Value = False
        If ChkRecaida(10).Value Then RsGuardar.Fields("R_60M").Value = True Else RsGuardar.Fields("R_60M").Value = False
        If ChkRecaida(11).Value Then RsGuardar.Fields("R_MaxM").Value = Trim(Text5(5).Text) Else RsGuardar.Fields("R_MaxM").Value = ""
        
        If ChkMuert(0).Value Then RsGuardar.Fields("M_MinM").Value = Trim(Text5(6).Text) Else RsGuardar.Fields("M_MinM").Value = ""
        If ChkMuert(1).Value Then RsGuardar.Fields("M_6M").Value = True Else RsGuardar.Fields("M_6M").Value = False
        If ChkMuert(2).Value Then RsGuardar.Fields("M_12M").Value = True Else RsGuardar.Fields("M_12M").Value = False
        If ChkMuert(3).Value Then RsGuardar.Fields("M_18M").Value = True Else RsGuardar.Fields("M_18M").Value = False
        If ChkMuert(4).Value Then RsGuardar.Fields("M_24M").Value = True Else RsGuardar.Fields("M_24M").Value = False
        If ChkMuert(5).Value Then RsGuardar.Fields("M_30M").Value = True Else RsGuardar.Fields("M_30M").Value = False
        If ChkMuert(6).Value Then RsGuardar.Fields("M_36M").Value = True Else RsGuardar.Fields("M_36M").Value = False
        If ChkMuert(7).Value Then RsGuardar.Fields("M_42M").Value = True Else RsGuardar.Fields("M_42M").Value = False
        If ChkMuert(8).Value Then RsGuardar.Fields("M_48M").Value = True Else RsGuardar.Fields("M_48M").Value = False
        If ChkMuert(9).Value Then RsGuardar.Fields("M_54M").Value = True Else RsGuardar.Fields("M_54M").Value = False
        If ChkMuert(10).Value Then RsGuardar.Fields("M_60M").Value = True Else RsGuardar.Fields("M_60M").Value = False
        If ChkMuert(11).Value Then RsGuardar.Fields("M_MaxM").Value = Trim(Text5(7).Text) Else RsGuardar.Fields("M_MaxM").Value = ""
        
        RsGuardar.Update
        
        Call EnviarRegPendiente(NuevoId, NuevoId2, "Informe_Medico3", "Id=" & NuevoId & " And IdL='" & NuevoId2 & "'")
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        If Trim(TxtDosisT.Text) = "" Then TxtDosisT.Text = "0"
        If Trim(TxtDosisD.Text) = "" Then TxtDosisD.Text = "0"
        If Trim(TxtSesionesFin.Text) = "" Then TxtSesionesFin.Text = "0"
        
        CSql = "SELECT * FROM Informe_Medico4 WHERE IdInforme=" & IdInf & " And IdL='" & IdLIdInf & "'"
        Set RsGuardar = CrearRS(CSql)
        
        If RsGuardar.RecordCount = 0 Then
            CSql = "SELECT MAX(Id) + 1 AS NuevoId FROM Informe_Medico4"
            Set RsGuardar = CrearRS(CSql)
            
            If Not IsNull(RsGuardar.Fields(0).Value) Then
                NuevoId = Val(RsGuardar.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If
            
            CSql = "SELECT * FROM Informe_Medico4"
            Set RsGuardar = CrearRS(CSql)
            
            RsGuardar.AddNew
            
            NuevoId2 = NuevoIdL
            RsGuardar.Fields("Id").Value = NuevoId
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdInforme").Value = IdInf
            RsGuardar.Fields("IdLIdPac").Value = IdLIdPac
            RsGuardar.Fields("IdL").Value = NuevoId2
            RsGuardar.Fields("IdLIdInf").Value = IdLIdInf
        Else
            NuevoId = RsGuardar.Fields("Id").Value
            NuevoId2 = RsGuardar.Fields("IdL").Value
        End If
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        
        RsGuardar.Fields("ExamFisicoInicial").Value = Trim(TxtExamFIni.Text)
        RsGuardar.Fields("ExamFisicoFinal").Value = Trim(TxtExamFin.Text)
        RsGuardar.Fields("Diagnostico").Value = Trim(TxtDiagFin.Text)
        RsGuardar.Fields("Anatomia").Value = Trim(TxtAnatFin.Text)
        RsGuardar.Fields("Complicaciones").Value = Trim(TxtCompliFin.Text)
        RsGuardar.Fields("Indice").Value = Trim(TxtInidiceFin.Text)
        RsGuardar.Fields("Comentarios").Value = Trim(TxtTratamientoFin.Text)
        RsGuardar.Fields("DosisT").Value = CDbl(TxtDosisT.Text)
        RsGuardar.Fields("DosisD").Value = CDbl(TxtDosisD.Text)
        RsGuardar.Fields("Sesiones").Value = Val(TxtSesionesFin.Text)
        
        If Combo9.ListIndex <> -1 Then
            RsGuardar.Fields("IdMedicoT").Value = Combo9.ItemData(Combo9.ListIndex)
        Else
            RsGuardar.Fields("IdMedicoT").Value = "0"
        End If
        
        If Combo8.ListIndex <> -1 Then
            RsGuardar.Fields("Metas").Value = Trim(Combo8.Text)
        Else
            RsGuardar.Fields("Metas").Value = ""
        End If
        
        RsGuardar.Update
        
        Call EnviarRegPendiente(NuevoId, NuevoId2, "Informe_Medico4", "Id=" & NuevoId & " And IdL='" & NuevoId2 & "'")
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        
        BtnDesHacer_Click
        MsgBox "Los datos fueron guardardos!", vbInformation + vbOKOnly, "Operación exitosa!"
        Exit Sub
    End If
       MsgBox "No hay cambios para guardar", vbExclamation + vbOKOnly, "No hay cambios"
       Frame1.Enabled = True
       
    If RsInformeMed.RecordCount <> 0 Then
        Call carga_datos_radio
        Cambio = 0
        actualiza = 1
        Else
        actualiza = 0
        Cambio = 0
    End If
    Intentos = 0
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly + vbCritical, "Error al Guardar"
    
Exit Sub

WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Sub EnviarRegPendiente(ByVal IdInf2 As Integer, ByVal IdLIdInf2 As String, ByVal Caden As String, ByVal Cond As String)
On Error Resume Next

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "SELECT * FROM " & Caden & " WHERE " & Cond
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO " & Caden & " (["
For i = 0 To RsTemp.Fields.Count - 1
    If Not i = (RsTemp.Fields.Count - 1) Then
        StrSen = StrSen & RsTemp.Fields(i).Name & "],["
    Else
        StrSen = StrSen & RsTemp.Fields(i).Name & "]) VALUES ("
    End If
Next i
For i = 0 To RsTemp.Fields.Count - 1
    If Not i = (RsTemp.Fields.Count - 1) Then
        StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "',"
    Else
        StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "')"
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = Replace(StrSen, "'", "(varCSP)")

CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Oncología"
RsRegPendiente.Fields("Tabla").Value = Caden
RsRegPendiente.Fields("Condicional").Value = Cond
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub



Private Sub BtnGuardarComplica_Click()
Dim NuevoId As Integer
Dim Rsp As Byte

Rsp = MsgBox("Se procedera a guardar los cambios realizados, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar!")

If Rsp = vbNo Then Exit Sub

CSql = "SELECT MAX(Id)+1 AS NuevoId FROM Informe_Medico5"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If


CSql = "SELECT caP,  caPTXT,  caMM,  caMMTXT,  caO,  caOTXT,  caOI,  caOITXT,  caGS,  caGSTXT,  caFE,  caFETXT, " & _
  " caL,  caLTXT,  caTGIS,  caTGISTXT,  caTGIIIP,  caGIIIPTXT,  caPUL,  caPULTXT,  caG,  caGTXT, " & _
  " caC,  caCTXT,  caS,  caSTXT, " & _
  " thaG,  thaGTXT,  thaP,  thaPTXT,  [thaN],  thaNTXT,  thaHGBN,  thaHGBNTXT,  thaHTCT,  thaHTCTTXT, " & _
  " ccP,  ccPTXT,  ccTS,  ccTSTXT,  ccMM,  ccMMTXT,  ccGS,  ccGSTXT,  ccME,  ccMETXT,  ccC,  ccCTXT, " & _
  " ccO,  ccOTXT,  ccL,  ccLTXT,  ccPUL,  ccPULTXT,  ccCORAZON,  ccCORAZONTXT,  ccE,  ccETXT, " & _
  " ccIGD , ccIGDTXT, ccH, ccHTXT, ccR, ccRTXT, ccV, ccVTXT, ccHUESO, ccHUESOTXT, ccA, ccATXT, " & _
  " Id, IdInforme, IdPaciente, IdUsuario, Observaciones, Fecha FROM Informe_Medico5"
  
'CSql = "SELECT * FROM Informe_Medico5"
Set RsTemp = CrearRS(CSql)

If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If

RsTemp.AddNew
RsTemp.Fields("Id").Value = NuevoId
RsTemp.Fields("IdL").Value = NuevoIdL

RsTemp.Fields("IdPaciente").Value = IdPac1
RsTemp.Fields("IdLIdPac").Value = IdLIdPac

RsTemp.Fields("IdInforme").Value = IdInf
RsTemp.Fields("IdLIdInf").Value = IdLIdInf

RsTemp.Fields("IdUsuario").Value = IdUser
RsTemp.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/MM/yyyy")
RsTemp.Fields("Observaciones").Value = Trim(TxtOtrasObs.Text)

For i = 0 To CboGrado.Count - 1
    If CboGrado(i).ListIndex = -1 Then CboGrado(i).ListIndex = 0
    RsTemp.Fields(i * 2) = CboGrado(i).ItemData(CboGrado(i).ListIndex)
    If CboGrado(i).ItemData(CboGrado(i).ListIndex) <> -1 Then
        RsTemp.Fields((i * 2) + 1).Value = Trim(TxtDosisGrdo(i).Text)
    Else
        RsTemp.Fields((i * 2) + 1).Value = ""
    End If
Next

'For i = 0 To TxtDosisGrdo.Count - 1
'    RsTemp.Fields((i * 2) + 1).Value = Trim(TxtDosisGrdo(i).Text)
'    MsgBox RsTemp.Fields((i * 2) + 1).Name & "=" & Trim(TxtDosisGrdo(i).Text)
'Next
        
RsTemp.Update

Call EnviarRegPendiente(NuevoId, NuevoIdL, "Informe_Medico5", "Id=" & NuevoId & " And IdL='" & NuevoIdL & "'")
        
MsgBox "Se han guardado los cambios exitosamente!", vbInformation + vbOKOnly, "Operación Exitosa!"

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

  TxtOtrasObs.Text = ""
  TxtOtrasObs.Enabled = False
  TxtOtrasObs.BackColor = &HE0E0E0

  DTPicker1.Enabled = False
  BtnGuardarComplica.Enabled = False
  BtnCancelar.Enabled = False
  For i = 0 To TxtDosisGrdo.Count - 1
      TxtDosisGrdo(i).Enabled = False
      TxtDosisGrdo(i).Text = ""
      TxtDosisGrdo(i).BackColor = &HE0E0E0
  Next
  For i = 0 To CboGrado.Count - 1
      CboGrado(i).Enabled = False
      CboGrado(i).ListIndex = 0
  Next

  CARGAR_COMPLICA

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


End Sub

Private Sub BtnInformeFinal_Click()
On Error GoTo WrtError
If Text1.Text <> "" Then

    With CrystalReport1
        .ReportFileName = RutaInformes & "\InformeFinalMedico.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{Informe_Med_Final.Id} = " & IdRegFinal
        .WindowTitle = "Reporte Informe Medico Final No. " & IdRegFinal
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
Else
    MsgBox "Tiene que seleccionar a un Paciente", vbCritical + vbOKOnly, "Mensaje de Error"
    Exit Sub
End If

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Private Sub BtnListaEspera_Click()
ModulO = 4
FrmListaEspera.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnLlamar_Click()
On Error Resume Next
Call Llamar
End Sub

Private Sub BtnSiguiente_Click()
On Error Resume Next
Call Blanqueo

If RsCargarPacientes.RecordCount <> 0 Then
    If Not RsCargarPacientes.EOF Then
        RsCargarPacientes.MoveNext
        If RsCargarPacientes.EOF Then RsCargarPacientes.MoveFirst
        Carga_De_Datos
        Call CONSULTA_INFORME
        If RsInformeMed.RecordCount <> 0 Then
            Call carga_datos_radio
            Cambio = 0
            actualiza = 1
            Else
            Call Habilita_Btns("sin informe")
            CboModificarMedicoTratante.Enabled = False
            Cambio = 0
            actualiza = 0
        End If
    End If
    CARGAR_INFORME2
Else
    MsgBox "No se encontraron datos, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros cargados"
End If

If OptInformeMedico(0).Value Then BtnGuardarActualizar.Enabled = False
End Sub

Private Sub BtnSiguiente2_Click()
On Error Resume Next

If Trim(IdPac1) = "" Then Exit Sub

If Cambio = 1 Then
    Msg = "HA REALIZADO CAMBIOS EN ESTE REGISTRO!!!!!!!" & Chr(13) & "DESEA GUARDAR ESTOS CAMBIOS????"
    f = MsgBox(Msg, vbQuestion + vbYesNo, "GUARDAR CAMBIOS??")
    If f = 6 Then Call BtnGuardarActualizar_Click: Exit Sub
End If

Call Blanqueo
If RsInformeMed.RecordCount <> 0 Then
    If Not (RsInformeMed.EOF) Then
        RsInformeMed.MoveNext
        If RsInformeMed.EOF Then RsInformeMed.MoveFirst
    End If
    Call carga_datos_radio
    Cambio = 0
    actualiza = 1
    Else
    MsgBox "No existen informes medicos para este paciente!", vbExclamation + vbOKOnly, "No existen los datos"
End If
End Sub

Private Sub BtnTipoCancer_Click()
On Error Resume Next
FrmAgregarTipoCancer.Show vbModal, FrmPrincipal

Leer_Tipos_Ca

End Sub

Private Sub BtnVerHistoria_Click()
On Error GoTo WrtError
If Text1.Text = "" Or IdPac1 = "" Then
    MsgBox "Debe de seleccionar un Paciente", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If

''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport2
    .ReportFileName = RutaInformes & "\HistoriaClinicaIntegralN.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Historia_Clinica.Idpaciente} = " & IdPac1
    .WindowTitle = "Reporte Historia Medica No. " & Label12.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Private Sub BtnExamenes_Click()
On Error Resume Next

If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub
    FrmExamenHematologico.IdPacH = IdPac1
    FrmExamenHematologico.IdLIdPacH = IdLIdPac
    FrmExamenHematologico.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnVerInforme_Click()
On Error GoTo WrtError
If Text1.Text <> "" Then

    With CrystalReport1
        .ReportFileName = RutaInformes & "\InformeMedicoN.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{Informe_Med.IdInforme} = " & IdReg & ""
        .WindowTitle = "Reporte Informe Medico No. " & IdReg
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
Else
    MsgBox "Tiene que seleccionar a un Paciente", vbCritical + vbOKOnly, "Mensaje de Error"
    Exit Sub
End If

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

'CSql = "Select * From Informe_Med Where IdInforme=" & IdReg & ""
'Set RsReporte = CrearRS(CSql)
'
'If RsReporte.RecordCount > 0 Then
'
'    Set DrptInformeMedico.DataSource = RsReporte
'
'    With DrptInformeMedico
'        .Sections("Sección2").Controls("LblPaciente").Caption = Trim(RsReporte.Fields("ApellidoP").Value) & ", " & Trim(RsReporte.Fields("NombreP").Value)
'        .Sections("Sección2").Controls("LblFechaNacimiento").Caption = Format(RsReporte.Fields("Fecha_NacimientoP").Value, "dd/mm/yyyy")
'        .Sections("Sección2").Controls("LblCedula").Caption = Format(Trim(RsReporte.Fields("CedulaP").Value), "#,##0")
'        .Sections("Sección2").Controls("LblEdad").Caption = Trim(RsReporte.Fields("EdadP").Value) & " Años"
'        .Sections("Sección2").Controls("LblTelefono").Caption = "(" & Trim(RsReporte.Fields("Codigo").Value) & ")" & " - " & RsReporte.Fields("Telefono").Value & " / " & "(" & Trim(RsReporte.Fields("CodigoC").Value) & ")" & " - " & RsReporte.Fields("Celular").Value
'        .Sections("Sección2").Controls("LblDireccion").Caption = Trim(RsReporte.Fields("DireccionP").Value)
'        .Sections("Sección2").Controls("LblOcupacion").Caption = Trim(RsReporte.Fields("Ocupacion").Value)
'
'
'        .Sections("Sección2").Controls("LblMotivoConsulta").Caption = Trim(RsReporte.Fields("Motivo_Con").Value)
'        .Sections("Sección2").Controls("LblDiagnostico").Caption = Trim(RsReporte.Fields("Diagnotico").Value)
'        .Sections("Sección2").Controls("LblTratamiento").Caption = Trim(RsReporte.Fields("Tratamiento").Value)
'        .Sections("Sección2").Controls("LblEnfermedadActual").Caption = Trim(RsReporte.Fields("Enfermedad_Act").Value)
'        .Sections("Sección2").Controls("LblAnatomiaPatologica").Caption = Trim(RsReporte.Fields("Anatomia_Patol").Value)
'
'
'
'        .Show vbModal
'    End With
'Else
'    MsgBox "El Paciente No posee Informe Medico realizado!!", vbCritical + vbOKOnly, "Mensaje de Error"
'End If
'
'    Call Enviar_Bitacora(IdUser, "Historial Medico", "IMPRIMIR", "Se imprimio del paciente de IdPaciente (" & IdPac1 & ") el Informe medico de IdInforme (" & IdInfor & ")")
'
'Else
'    MsgBox "Tiene que seleccionar a un Paciente", vbCritical + vbOKOnly, "Mensaje de Error"
'End If
End Sub

Private Sub CboMetas_Change()
Cambio = 1
End Sub

Private Sub CboMetas_Click()
Cambio = 1
End Sub

Private Sub CboModificarMedicoTratante_Click()
Cambio = 1
End Sub

Private Sub ChkGlobal_Click(Index As Integer)
For ii = 0 To ChkGlobal.Count - 2
    
    If ChkGlobal(ii).Value Then
        CboGeneral(ii).Enabled = True
    Else
        CboGeneral(ii).Enabled = False
    End If
    
    If ChkGlobal(57).Value Then
        TxtOtros.Enabled = True
    Else
        TxtOtros.Enabled = False
    End If
Next ii
End Sub

Private Sub Combo1_Change()
Cambio = 1
End Sub

Private Sub Combo1_Click()

If Combo1.ListIndex = 0 Then Text17.Enabled = False
If Combo1.ListIndex = 1 Then Text17.Enabled = True

Text17.Text = ""
Cambio = 1

End Sub

Sub CONSULTA_INFORME()
On Error GoTo WrtError

CSql = "Select * From Informe_Medico Where IdPaciente = " & IdPac1 & " And Estado=1 Order By Fecha Desc"
Set RsInformeMed = CrearRS(CSql)
    
If RsInformeMed.RecordCount = 0 Then IdLIdInf = IdLDefault

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Sub Carga_De_Datos()
On Error GoTo WrtError

    IdLIdPac = IdLDefault
    Text1.Text = RsCargarPacientes.Fields("cedulaP").Value
    DtpFechaRegistro.Value = RsCargarPacientes.Fields("Fecha_regp").Value
    Text3.Text = RsCargarPacientes.Fields("Nombrep").Value
    Text4.Text = RsCargarPacientes.Fields("Apellidop").Value
    DtpFechaNac = RsCargarPacientes.Fields("Fecha_nacimientop").Value
    Text6.Text = RsCargarPacientes.Fields("Edadp").Value
    
    IdPac1 = RsCargarPacientes.Fields("idpaciente").Value
    IdLIdPac = RsCargarPacientes.Fields("IdL").Value
    
    Me.Caption = "Oncología - Paciente: " & IdPac1
    NoReg = "Registro: " & RsCargarPacientes.AbsolutePosition & " / " & RsCargarPacientes.RecordCount
    If Not IsNull(RsCargarPacientes.Fields("foto").Value) Then
        If RsCargarPacientes.Fields("foto") <> "" Then
            Image2.Picture = LoadPicture(Foto & "\" & RsCargarPacientes.Fields("foto").Value)
        Else
            Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        End If
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
    If Trim(RsCargarPacientes.Fields("Historia").Value) <> "" Then Label12.Caption = RsCargarPacientes.Fields("Historia").Value Else Label12.Caption = ""
    
    Dim BD2 As New ADODB.Recordset 'Tabla Medico Tratante
    CSql = "SELECT * FROM medicos where [idmedico] = " & RsCargarPacientes.Fields("Medico_Tratante").Value & " AND (Tipo=2 OR Tipo=3)"
    Set BD2 = CrearRS(CSql)
    If BD2.EOF Then
        Msg = "Verifique el Medico Tratante"
        MsgBox Msg, vbOKOnly + vbCritical, "Medico Tratante"
    Else
        Text10.Text = BD2.Fields("nombre").Value & " " & BD2.Fields("apellido").Value
    End If
    BD2.Close
    
    Dim BD3 As New ADODB.Recordset 'Tabla Medico Remitente
    CSql = "SELECT * FROM medicos where [idmedico] = " & RsCargarPacientes.Fields("Medico_Remitente").Value & " AND (Tipo=1 OR Tipo=3)"
    Set BD3 = CrearRS(CSql)
    If BD3.EOF Then
        Msg = "Verifique Medico Remitente"
        MsgBox Msg, vbOKOnly + vbCritical, "Medico Remitente"
    Else
        Text11.Text = BD3.Fields("nombre").Value & " " & BD3.Fields("apellido").Value
    End If
    BD3.Close
                
    If RsCargarPacientes.Fields("sexop").Value = 0 Then Text12.Text = "Masculino" Else Text12.Text = "Femenino"
    
    DtpFechaInicio.Value = RsCargarPacientes.Fields("Fecha_inicio").Value
    DtpFechaFin.Value = RsCargarPacientes.Fields("Fecha_culm").Value
    
            
Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text17.SetFocus
End If
End Sub

Private Sub Combo2_Change()
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo2_Click()
Cambio = 1
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo3_Change()
Cambio = 1
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"
End Sub

Private Sub Combo4_Change()
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo4_Click()
Cambio = 1
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo5_Change()
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"
End Sub

Private Sub Combo5_Click()
Cambio = 1
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo6_Change()
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo6_Click()
Cambio = 1
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub

Private Sub Combo7_Change()
Cambio = 1
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & ")"
'Estadiaje = "ST" & Replace(Combo2.Text, ".", "") & "(" & Replace(Combo3.Text, ".", "") & "." & Replace(Combo4.Text, ".", "") & "." & Replace(Combo5.Text, ".", "") & "." & Replace(Combo6.Text, ".", "") & "." & Replace(Combo7.Text, ".", "") & "." & Replace(Combo8.Text, ".", "") & ")"

End Sub


Private Sub DMGrid1_DobleClick()


If DMGrid1.Rows = 0 Then Exit Sub
If DMGrid1.Row = 0 Then Exit Sub

CSql = "SELECT caP,  caPTXT,  caMM,  caMMTXT,  caO,  caOTXT,  caOI,  caOITXT,  caGS,  caGSTXT,  caFE,  caFETXT, " & _
  " caL,  caLTXT,  caTGIS,  caTGISTXT,  caTGIIIP,  caGIIIPTXT,  caPUL,  caPULTXT,  caG,  caGTXT, " & _
  " caC,  caCTXT,  caS,  caSTXT, " & _
  " thaG,  thaGTXT,  thaP,  thaPTXT,  [thaN],  thaNTXT,  thaHGBN,  thaHGBNTXT,  thaHTCT,  thaHTCTTXT, " & _
  " ccP,  ccPTXT,  ccTS,  ccTSTXT,  ccMM,  ccMMTXT,  ccGS,  ccGSTXT,  ccME,  ccMETXT,  ccC,  ccCTXT, " & _
  " ccO,  ccOTXT,  ccL,  ccLTXT,  ccPUL,  ccPULTXT,  ccCORAZON,  ccCORAZONTXT,  ccE,  ccETXT, " & _
  " ccIGD , ccIGDTXT, ccH, ccHTXT, ccR, ccRTXT, ccV, ccVTXT, ccHUESO, ccHUESOTXT, ccA, ccATXT, " & _
  " Observaciones FROM Informe_Medico5 WHERE Id=" & Val(DMGrid1.ValorCelda(DMGrid1.Row, 3)) & " ORDER BY ID"

Set RsTemp = CrearRS(CSql)
 
If RsTemp.RecordCount = 0 Then
    MsgBox "Hubo un error en la consulta de la base de datos!", vbCritical + vbOKOnly, "Contacte al Administrador!"
    Exit Sub
End If
' Carga los ComboBox de los Grados para cada Opción
For i = 0 To CboGrado.Count - 1
    For J = 0 To CboGrado(i).ListCount - 1
        If RsTemp.Fields(i * 2).Value = CboGrado(i).ItemData(J) Then
            CboGrado(i).ListIndex = J
            Exit For
        Else
            CboGrado(i).ListIndex = 0
        End If
    Next
Next

' Carga los .Text para los Grados de cada Opción
For i = 0 To TxtDosisGrdo.Count - 1
    TxtDosisGrdo(i).Text = RsTemp.Fields((i * 2) + 1).Value
Next

TxtOtrasObs.Text = RsTemp.Fields("Observaciones").Value

End Sub

Private Sub Form_Load()
On Error GoTo WrtError

Centrar Me
ModulO = 4

IniDMGrid

FrameDato(0).Left = 120
FrameDato(1).Left = 12600

FrameDato(0).Top = 240
FrameDato(1).Top = 240

DesactivarTextos
CSql = "select * From Medicos where (Tipo=2 or Tipo=3) AND Activo='1'"
Set RsTemp = CrearRS(CSql)

For i = 1 To RsTemp.RecordCount
    CboModificarMedicoTratante.AddItem RsTemp.Fields("Nombre").Value & " " & RsTemp.Fields("Apellido").Value
    CboModificarMedicoTratante.ItemData(CboModificarMedicoTratante.NewIndex) = Val(RsTemp.Fields("IdMedico").Value)
    
    Combo9.AddItem RsTemp.Fields("Nombre").Value & " " & RsTemp.Fields("Apellido").Value
    Combo9.ItemData(Combo9.NewIndex) = Val(RsTemp.Fields("IdMedico").Value)
    
    RsTemp.MoveNext
Next i

IdPac1 = ""
IdLIdPac = IdLDefault
IdLIdInf = IdLDefault
IdInf = ""

CSql = "select * from Paciente Order by IdPaciente"
Set RsCargarPacientes = CrearRS(CSql)

'If RsCargarPacientes.RecordCount <> 0 Then
'    Carga_De_Datos
'    Call CONSULTA_INFORME
'    If RsInformeMed.RecordCount <> 0 Then
'        Call carga_datos_radio
'        Cambio = 0
'        actualiza = 1
'    Else
'        Call Blanqueo
'        Call Habilita_Btns("sin informe")
'        CboModificarMedicoTratante.Enabled = False
'        Cambio = 0
'        actualiza = 0
'    End If
'
'    DtpFecha.Value = Now()
'
'Else
'    MsgBox "No se encontraron pacientes en la Base d e Datos!", vbExclamation + vbOKOnly, "Vacio"
'    IdPac1 = ""
'End If

Leer_Tipos_Ca
'CARGAR_COMPLICA
'CARGAR_INFORME2
Exit Sub

WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
'MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo WrtError

If Cambio = 1 And IdPac1 <> "" Then
Msg = "HA REALIZADO CAMBIOS EN ESTE REGISTRO!!!!!!!" & Chr(13) & "DESEA GUARDAR ESTOS CAMBIOS????"
f = MsgBox(Msg, vbInformation + vbYesNo, "GUARDAR CAMBIOS??")
If f = vbYes Then Call BtnGuardarActualizar_Click
End If
  If RsCargarPacientes.State = adStateOpen Then RsCargarPacientes.Close

Exit Sub
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub Frame9_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

 

Private Sub OptBordes_Click(Index As Integer)

OptBordesSel = OptBordes(Index).Caption

End Sub

Private Sub OptComplicaciones_Click(Index As Integer)

For i = 0 To OptComplicaciones.Count - 1
    If i <> Index Then
        OptComplicaciones(i).Value = False
            'FrameDato(i).Visible = False
        FrameComplicaciones(i).Left = 13200
        FrameComplicaciones(i).Top = 840
    Else
        OptComplicaciones(i).Value = True
            'FrameDato(i).Visible = True
        FrameComplicaciones(i).Left = 120
        FrameComplicaciones(i).Top = 840
    End If
Next
End Sub

Private Sub OptInformeMedico_Click(Index As Integer)

If Not Index = 4 Then Frame7.Visible = True Else Frame7.Visible = False
For i = 0 To OptInformeMedico.Count - 1
    If i <> Index Then
        OptInformeMedico(i).Value = False
            'FrameDato(i).Visible = False
        FrameDato(i).Left = 13200
        FrameDato(i).Top = 240
    Else
        OptInformeMedico(i).Value = True
            'FrameDato(i).Visible = True
        FrameDato(i).Left = 120
        FrameDato(i).Top = 240

    End If
Next
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text19.SetFocus
End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text20.SetFocus
End If
End Sub

Private Sub Text17_Change()
Cambio = 1
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Exit Sub
Case Is = 8
Exit Sub
Case Is = 13
Exit Sub
Case Else
KeyAscii = 0

End Select

If KeyAscii = 13 Then
    Text8.SetFocus
End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text21.SetFocus
End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text16.SetFocus
End If

End Sub

Private Sub Text2_Change()
On Error GoTo WrtError
Cambio = 1
Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
Sesiones = Text2.Text

Exit Sub
WrtError:
    Text2.Text = ""
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text18.SetFocus
End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text22.SetFocus
End If
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text8_Change()
On Error GoTo WrtError
Cambio = 1
Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
Sesiones = Text2.Text
Exit Sub
WrtError:
    Text2.Text = ""
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is = 46
Exit Sub
Case 48 To 57
Exit Sub
Case Is = 13
Exit Sub
Case Is = 8
Exit Sub
Case Is = 9
Exit Sub
End Select
KeyAscii = 0

If KeyAscii = 13 Then
    Text9.SetFocus
End If
End Sub

Private Sub Text9_Change()
On Error GoTo WrtError
Cambio = 1
Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
Sesiones = Text2.Text
Exit Sub
WrtError:
    Text2.Text = ""
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is = 46
Exit Sub
Case 48 To 57
Exit Sub
Case Is = 13
'Exit Sub
Case Is = 8
Exit Sub
Case Is = 9
Exit Sub
End Select
KeyAscii = 0

If KeyAscii = 13 Then
    BtnGuardarActualizar.SetFocus
End If
End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub
Private Sub Text15_Change()
Cambio = 1
End Sub
Private Sub Text16_Change()
Cambio = 1
End Sub
Private Sub Text18_Change()

Dim Poscur As Integer
Poscur = Text18.SelStart
If Mid(Text18.Text, 1, 1) <> UCase(Mid(Text18.Text, 1, 1)) Then
    Text18.Text = UCase(Mid(Text18.Text, 1, 1)) & Mid(Text18.Text, 2, Len(Text18.Text) - 1)
End If
Text18.SelStart = Poscur

Cambio = 1
End Sub
Private Sub Text19_Change()
Cambio = 1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text20_Change()
Cambio = 1
End Sub
Private Sub Text21_Change()
Cambio = 1
End Sub
Private Sub Text22_Change()
Cambio = 1
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaRegistro.SetFocus
        Case vbKeyRight
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub TxtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        Case vbKeyUp
            Text22.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
    End Select
End If
End Sub

Private Sub TxtDosisD_Change()
On Error GoTo WrtError
Cambio = 1
TxtSesionesFin.Text = CDbl(TxtDosisT.Text) / CDbl(TxtDosisD.Text)
Exit Sub
WrtError:
    TxtSesionesFin.Text = ""
End Sub

Private Sub TxtDosisD_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtDosisT_Change()
On Error GoTo WrtError
Cambio = 1
TxtSesionesFin.Text = CDbl(TxtDosisT.Text) / CDbl(TxtDosisD.Text)
Exit Sub
WrtError:
    TxtSesionesFin.Text = ""
End Sub

Private Sub TxtDosisT_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtSesionesFin_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

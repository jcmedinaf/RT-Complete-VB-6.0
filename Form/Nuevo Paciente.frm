VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmNuevoPaciente 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pacientes"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   3360
   ClientWidth     =   13785
   Icon            =   "Nuevo Paciente.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   13785
   Begin VB.Frame FrameBusquedaAvanzada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Busqueda Avanzada"
      Height          =   3495
      Left            =   13440
      TabIndex        =   70
      Top             =   1320
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox TxtFechaCulmi 
         Height          =   375
         Left            =   5280
         TabIndex        =   78
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtFechaInicio 
         Height          =   375
         Left            =   5280
         TabIndex        =   77
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtFechaReg 
         Height          =   375
         Left            =   5280
         TabIndex        =   76
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtCedula 
         Height          =   375
         Left            =   960
         TabIndex        =   75
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   375
         Left            =   960
         TabIndex        =   71
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox TxtHistoria 
         Height          =   375
         Left            =   960
         TabIndex        =   74
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox TxtApellido 
         Height          =   375
         Left            =   960
         TabIndex        =   73
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtNombre 
         Height          =   375
         Left            =   960
         TabIndex        =   72
         Top             =   840
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DtpFechaReg 
         Height          =   375
         Left            =   5040
         TabIndex        =   79
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47906817
         CurrentDate     =   40464
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   375
         Left            =   5040
         TabIndex        =   80
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47906817
         CurrentDate     =   40464
      End
      Begin MSComCtl2.DTPicker DtpFechaCulmi 
         Height          =   375
         Left            =   5040
         TabIndex        =   81
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47906817
         CurrentDate     =   40464
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar2 
         Height          =   375
         Left            =   2280
         TabIndex        =   82
         ToolTipText     =   "Realiza una Busqueda Avanzada"
         Top             =   2880
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
         MICON           =   "Nuevo Paciente.frx":1002
         PICN            =   "Nuevo Paciente.frx":101E
         PICH            =   "Nuevo Paciente.frx":1283
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
         Left            =   3600
         TabIndex        =   83
         ToolTipText     =   "Cerrar"
         Top             =   2880
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
         MICON           =   "Nuevo Paciente.frx":1515
         PICN            =   "Nuevo Paciente.frx":1531
         PICH            =   "Nuevo Paciente.frx":16FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Culminación:"
         Height          =   195
         Left            =   3600
         TabIndex        =   91
         Top             =   1410
         Width           =   1395
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio:"
         Height          =   195
         Left            =   3600
         TabIndex        =   90
         Top             =   930
         Width           =   915
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro:"
         Height          =   195
         Left            =   3600
         TabIndex        =   89
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula:"
         Height          =   195
         Left            =   240
         TabIndex        =   88
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Historia:"
         Height          =   195
         Left            =   240
         TabIndex        =   86
         Top             =   1890
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido:"
         Height          =   195
         Left            =   240
         TabIndex        =   85
         Top             =   1410
         Width           =   600
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   84
         Top             =   930
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   5160
      TabIndex        =   51
      Top             =   6000
      Width           =   8535
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         ToolTipText     =   "Cerrar Pacientes"
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
         MICON           =   "Nuevo Paciente.frx":192F
         PICN            =   "Nuevo Paciente.frx":194B
         PICH            =   "Nuevo Paciente.frx":1B14
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrarPaciente 
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         ToolTipText     =   "Borrar Paciente"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "Nuevo Paciente.frx":1D49
         PICN            =   "Nuevo Paciente.frx":1D65
         PICH            =   "Nuevo Paciente.frx":1F09
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
         TabIndex        =   31
         ToolTipText     =   "Guardar / Actualizar Paciente"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "Nuevo Paciente.frx":20A8
         PICN            =   "Nuevo Paciente.frx":20C4
         PICH            =   "Nuevo Paciente.frx":2353
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregarPaciente 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Nuevo Paciente"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "Nuevo Paciente.frx":2794
         PICN            =   "Nuevo Paciente.frx":27B0
         PICH            =   "Nuevo Paciente.frx":293D
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
         Left            =   4560
         TabIndex        =   33
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
         MICON           =   "Nuevo Paciente.frx":2B72
         PICN            =   "Nuevo Paciente.frx":2B8E
         PICH            =   "Nuevo Paciente.frx":2E23
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
         Left            =   5280
         TabIndex        =   34
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
         MICON           =   "Nuevo Paciente.frx":307F
         PICN            =   "Nuevo Paciente.frx":309B
         PICH            =   "Nuevo Paciente.frx":3331
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
         Left            =   6240
         TabIndex        =   35
         ToolTipText     =   "Ignorar cambios"
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
         MICON           =   "Nuevo Paciente.frx":3590
         PICN            =   "Nuevo Paciente.frx":35AC
         PICH            =   "Nuevo Paciente.frx":388E
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Indicar Foto del Paciente"
      FileName        =   "*.jpg"
      Filter          =   "*.jpg; *.bmp"
      InitDir         =   "z:\fotos\"
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   840
      Top             =   5280
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   47
      Top             =   6000
      Width           =   4935
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   30
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "Nuevo Paciente.frx":3ADF
         PICN            =   "Nuevo Paciente.frx":3AFB
         PICH            =   "Nuevo Paciente.frx":3D60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtBuscar 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido,Cédula de identidad o Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBusquedaAvanzada 
         Height          =   375
         Left            =   3600
         TabIndex        =   69
         ToolTipText     =   "Realiza una Busqueda Avanzada"
         Top             =   240
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
         MICON           =   "Nuevo Paciente.frx":3FF2
         PICN            =   "Nuevo Paciente.frx":400E
         PICH            =   "Nuevo Paciente.frx":4273
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
      Caption         =   "Datos del Paciente"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   1455
         Left            =   11280
         TabIndex        =   64
         Top             =   4200
         Visible         =   0   'False
         Width           =   2175
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   120
            MaxLength       =   15
            TabIndex        =   65
            Top             =   360
            Width           =   1935
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardarHOAM 
            Height          =   375
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "Verifica y establece un número de historia"
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Establecer"
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
            MICON           =   "Nuevo Paciente.frx":4505
            PICN            =   "Nuevo Paciente.frx":4521
            PICH            =   "Nuevo Paciente.frx":47B0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
            Height          =   375
            Left            =   1560
            TabIndex        =   67
            ToolTipText     =   "Volver"
            Top             =   840
            Width           =   495
            _ExtentX        =   873
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
            MICON           =   "Nuevo Paciente.frx":4BF1
            PICN            =   "Nuevo Paciente.frx":4C0D
            PICH            =   "Nuevo Paciente.frx":4EEF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ingrese el Nº Historia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Tipo"
         Height          =   855
         Left            =   2520
         TabIndex        =   55
         Top             =   4800
         Width           =   3255
         Begin VB.OptionButton OptPaciente 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Paciente"
            Height          =   255
            Left            =   1680
            TabIndex        =   57
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton OptPresupuesto 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Presupuesto"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Tratamiento RT:"
         Height          =   855
         Left            =   5880
         TabIndex        =   54
         Top             =   4800
         Width           =   3495
         Begin VB.ComboBox CboGrupo 
            Height          =   315
            ItemData        =   "Nuevo Paciente.frx":5140
            Left            =   1560
            List            =   "Nuevo Paciente.frx":51F8
            TabIndex        =   59
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Asignada:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   1680
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   52
         ToolTipText     =   "Se usa para PROBAR si hay camaras web conectadas (ASI QUE NO BORRAR!!!)"
         Top             =   5040
         Visible         =   0   'False
         Width           =   375
      End
      Begin ChamaleonButton.ChameleonBtn BtnGenerarHistoriaMedica 
         Height          =   525
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Generar Nº de la historia médica del paciente"
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "Generar Nº Hist. Médica"
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
         MICON           =   "Nuevo Paciente.frx":54CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnNuevoMedico 
         Height          =   735
         Left            =   12480
         TabIndex        =   29
         ToolTipText     =   "Agregar Nuevo Médico"
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "Nuevo Médico"
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
         MICON           =   "Nuevo Paciente.frx":54E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "Nuevo Paciente.frx":5504
         Left            =   9000
         List            =   "Nuevo Paciente.frx":5506
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   17
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         Height          =   975
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   3720
         Width           =   10815
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   9000
         MaxLength       =   40
         TabIndex        =   9
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3600
         MaxLength       =   40
         TabIndex        =   8
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text6 
         Height          =   975
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1680
         Width           =   9735
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   19
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Nuevo Paciente.frx":5508
         Left            =   3600
         List            =   "Nuevo Paciente.frx":550A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   10680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   10680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Nuevo Paciente.frx":550C
         Left            =   7440
         List            =   "Nuevo Paciente.frx":550E
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   47906819
         CurrentDate     =   39784
         MinDate         =   -108932
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   12000
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   47906819
         CurrentDate     =   39784
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   9360
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   47906819
         CurrentDate     =   39784
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   47906819
         CurrentDate     =   39784
      End
      Begin ChamaleonButton.ChameleonBtn BtnListadoPaciente 
         Height          =   375
         Left            =   120
         TabIndex        =   62
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   4440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Listado Pacientes"
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
         MICON           =   "Nuevo Paciente.frx":5510
         PICN            =   "Nuevo Paciente.frx":552C
         PICH            =   "Nuevo Paciente.frx":57B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnTomarFotoPaciente 
         Height          =   615
         Left            =   120
         TabIndex        =   63
         Top             =   3720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Tomar Foto"
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
         MICON           =   "Nuevo Paciente.frx":5A48
         PICN            =   "Nuevo Paciente.frx":5A64
         PICH            =   "Nuevo Paciente.frx":60E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Historia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label NoReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   840
         TabIndex        =   61
         Top             =   4920
         Width           =   90
      End
      Begin VB.Label LblNoReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro:"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   4920
         Width           =   630
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Años"
         Height          =   195
         Left            =   7320
         TabIndex        =   53
         Top             =   1290
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   8400
         TabIndex        =   50
         Top             =   1260
         Width           =   495
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   120
         Picture         =   "Nuevo Paciente.frx":6746
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
         Height          =   195
         Left            =   2520
         TabIndex        =   46
         Top             =   3480
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula:"
         Height          =   195
         Left            =   2520
         TabIndex        =   45
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro:"
         Height          =   435
         Left            =   6000
         TabIndex        =   44
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         Height          =   195
         Left            =   2520
         TabIndex        =   43
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido(s):"
         Height          =   195
         Left            =   8160
         TabIndex        =   42
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edad:"
         Height          =   195
         Left            =   6120
         TabIndex        =   41
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   2520
         TabIndex        =   40
         Top             =   2070
         Width           =   720
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Médico Tratante:"
         Height          =   195
         Left            =   9240
         TabIndex        =   39
         Top             =   2820
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio:"
         Height          =   375
         Left            =   8520
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Culminación:"
         Height          =   495
         Left            =   10920
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono Hab. ó Movil:"
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupación:"
         Height          =   195
         Left            =   6495
         TabIndex        =   26
         Top             =   3210
         Width           =   825
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Médico Remitente:"
         Height          =   195
         Left            =   9240
         TabIndex        =   25
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
         Height          =   195
         Left            =   6495
         TabIndex        =   24
         Top             =   2820
         Width           =   405
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento:"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   3930
      Left            =   15000
      Picture         =   "Nuevo Paciente.frx":9AC7
      Top             =   6000
      Width           =   6645
   End
End
Attribute VB_Name = "FrmNuevoPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ws_visible = &H10000000
Const ws_child = &H40000000
Dim T As Variant
Dim SentenciaSQL As String
Dim IdPac
Dim st_but ' para saber la condicion actual del boton añadir registro
Dim Cambio
'Dim FotoP As String
Dim Actuali
Dim RsPacientes As New ADODB.Recordset
Dim RsBitacora As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim Reg_Actual(0 To 23) As String
Dim Estatus, FotoP
Dim Histo, NoHistoria
Dim RsGenerarHistoria As New ADODB.Recordset
Dim RsNumeroHistoria As New ADODB.Recordset
Dim RsUpDateHistoria As New ADODB.Recordset
Dim p

Sub actualiza()
On Error Resume Next

CSql = "Select * From Paciente"
If RsPacientes.State = 1 Then RsPacientes.Close
Set RsPacientes = CrearRS(CSql)

End Sub
Sub GuardarCambios()
'On Error Resume Next

Dim RsGuardarCambios As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim SqlTemp As String

JH = Format(DTPicker1.Value, "DD/MM/YYYY")
JH2 = Format(DTPicker2.Value, "DD/MM/YYYY")
JH3 = Format(DTPicker3.Value, "DD/MM/YYYY")
JH4 = Format(DTPicker4.Value, "DD/MM/YYYY")

If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If

If IdPac = "" Then
    SqlTemp = "Select MAX(IdPaciente)+1 as NuevoId From Paciente"
    Set RsTemp = CrearRS(SqlTemp)
    
    If Not IsNull(RsTemp.Fields("NuevoId")) Then
        IdPac = RsTemp.Fields("NuevoId").Value
    Else
        IdPac = "1"
    End If
    Set RsTemp = Nothing
    IdLIdPac = NuevoIdL
End If

CSql = "Select * From Paciente Where IdPaciente = " & IdPac & " AND IdL = '" & IdLIdPac & "'"
Set RsGuardarCambios = CrearRS(CSql)

If RsGuardarCambios.RecordCount = 0 Then
    If (Combo7.Text = "Activo") Or (Combo7.Text = "Simulación") Or (Combo7.Text = "Reubicación") Or (Combo7.Text = "Espera") Then
        CSql = "Select * From Paciente Where HoraAtencion='" & CboGrupo.Text & "' And (Status='A' Or Status='Si' Or Status='R')"
        Set RsTemp = CrearRS(CSql)
        If OptPaciente.Value = True Then
            If RsTemp.RecordCount >= 1 Then
                MsgBox "Bloque de Grupo lleno. Seleccione otro Horario de Grupo!!", vbCritical + vbOKOnly, "Error"
                CboGrupo.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    CSql = "Select * From Paciente"
    Set RsGuardarCambios = CrearRS(CSql)
    RsGuardarCambios.AddNew
End If

If (Combo7.ListIndex) <> 3 Then
    Estatus = Chr(Combo7.ItemData(Combo7.ListIndex))
Else
    Estatus = "Si"
End If

RsGuardarCambios.Fields("IdPaciente").Value = IdPac
RsGuardarCambios.Fields("IdL").Value = IdLIdPac
RsGuardarCambios.Fields("Status").Value = Estatus
RsGuardarCambios.Fields("Foto").Value = FotoP

If InStr(1, Label16.Caption, "HOAM") = 0 And InStr(1, Label16.Caption, "C") = 0 Then
    Label16.Caption = "C" & Trim(Text1.Text)
End If

RsGuardarCambios.Fields("Historia").Value = UCase(Label16.Caption)
RsGuardarCambios.Fields("Nacionalidad").Value = Combo5.ItemData(Combo5.ListIndex)
RsGuardarCambios.Fields("CedulaP").Value = Text1.Text
RsGuardarCambios.Fields("Fecha_RegP").Value = JH
RsGuardarCambios.Fields("NombreP").Value = Trim(Text4.Text)
RsGuardarCambios.Fields("ApellidoP").Value = Trim(Text3.Text)
RsGuardarCambios.Fields("Fecha_Inicio").Value = JH2
RsGuardarCambios.Fields("Fecha_Culm").Value = JH3
RsGuardarCambios.Fields("Fecha_NacimientoP").Value = JH4
RsGuardarCambios.Fields("edadP").Value = Val(Text11.Text)
RsGuardarCambios.Fields("DireccionP").Value = Text6.Text
RsGuardarCambios.Fields("Telefono").Value = Val(Text7.Text)

If Combo6.ListIndex <> -1 Then
    If Trim(Combo6.List(Combo6.ListIndex)) = "" Then
        RsGuardarCambios.Fields("CodigoC").Value = 0
    Else
        RsGuardarCambios.Fields("CodigoC").Value = Val(Combo6.ItemData(Combo6.ListIndex))
    End If
Else
    RsGuardarCambios.Fields("CodigoC").Value = 0
End If

If Combo1.ListIndex <> -1 Then
    If Trim(Combo1.List(Combo1.ListIndex)) = "" Then
        RsGuardarCambios.Fields("Codigo").Value = 0
    Else
        RsGuardarCambios.Fields("Codigo").Value = Val(Combo1.ItemData(Combo1.ListIndex))
    End If
Else
    RsGuardarCambios.Fields("Codigo").Value = 0
End If

RsGuardarCambios.Fields("Celular").Value = Val(Text2.Text)
RsGuardarCambios.Fields("SexoP").Value = Val(Combo4.ItemData(Combo4.ListIndex))
RsGuardarCambios.Fields("Ocupacion").Value = Text14.Text
RsGuardarCambios.Fields("Medico_Tratante").Value = Combo2.ItemData(Combo2.ListIndex)
RsGuardarCambios.Fields("Medico_Remitente").Value = Combo3.ItemData(Combo3.ListIndex)
RsGuardarCambios.Fields("Observacion").Value = Text13.Text
RsGuardarCambios.Fields("IdUsuario").Value = IdUser
RsGuardarCambios.Fields("Activo").Value = 1

If OptPresupuesto.Value = True Then Tipo = 0
If OptPaciente.Value = True Then Tipo = 1

RsGuardarCambios.Fields("Tipo").Value = Tipo
RsGuardarCambios.Fields("Grupo").Value = Mid(Trim(CboGrupo.Text), 1, 3) & "00" & Mid(Trim(CboGrupo.Text), 6)
RsGuardarCambios.Fields("HoraAtencion").Value = Trim(CboGrupo.Text)

RsGuardarCambios.Update
RsGuardarCambios.Close

If ACCION = EDITAR_REGISTRO Then
    If Not Reg_Actual(0) = Text1.Text Then              ' No Cedula
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo CEDULAP de  (" & Reg_Actual(0) & ")  a  (" & Text1.Text & ")")
    End If
    If Not CDbl(Reg_Actual(1)) = Val(Text2.Text) Then   ' No Telefono Mov
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo CELULAR de  (" & Reg_Actual(1) & ")  a  (" & Text2.Text & ")")
    End If
    If Not Reg_Actual(2) = Text3.Text Then              ' Apellido
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo APELLIDOP de  (" & Reg_Actual(2) & ")  a  (" & Text3.Text & ")")
    End If
    If Not Reg_Actual(3) = Text4.Text Then              ' Nombre
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo NOMBREP de  (" & Reg_Actual(3) & ")  a  (" & Text4.Text & ")")
    End If
    If Not Reg_Actual(4) = Text6.Text Then              ' Direccion
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo DIRECCIONP de  (" & Reg_Actual(4) & ")  a  (" & Text6.Text & ")")
    End If
    If Not Reg_Actual(5) = Val(Text7.Text) Then         ' No Telefono Hab.
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo TELEFONO de  (" & Reg_Actual(5) & ")  a  (" & Text7.Text & ")")
    End If
    If Not Reg_Actual(6) = FotoP Then                   ' Foto Paciente
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FOTO de  (" & Reg_Actual(6) & ")  a  (" & FotoP & ")")
    End If
    If Not Reg_Actual(7) = Combo1.ItemData(Combo1.ListIndex) Then             ' Cod Area Tel. Hab
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo CODIGO de  (" & Reg_Actual(7) & ")  a  (" & Combo1.Text & ")")
    End If
    If Not Reg_Actual(8) = Combo2.ItemData(Combo2.ListIndex) Then    ' Medico T.
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo MEDICO_TRATANTE de  (" & Reg_Actual(8) & ")  a  (" & Combo2.Text & ")")
    End If
    If Not Reg_Actual(9) = Combo3.ItemData(Combo3.ListIndex) Then  ' Medico R
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo MEDICO_REMITENTE de  (" & Reg_Actual(9) & ")  a  (" & Combo3.Text & ")")
    End If
    If Not Reg_Actual(10) = Combo4.ItemData(Combo4.ListIndex) Then            ' Sexo
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo SEXOP de  (" & Reg_Actual(10) & ")  a  (" & Combo4.Text & ")")
    End If
    If Not Reg_Actual(11) = Combo5.ItemData(Combo5.ListIndex) Then      ' Nacionalidad
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo NACIONALIDAD de  (" & Reg_Actual(11) & ")  a  (" & Combo5.Text & ")")
    End If
    If Not Reg_Actual(12) = Combo6.ItemData(Combo6.ListIndex) Then           ' Cod Area Tel. Mov
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo CODIGOC de  (" & Reg_Actual(12) & ")  a  (" & Combo6.Text & ")")
    End If
    If Not Reg_Actual(13) = Chr(Combo7.ListIndex) Then            ' Estatus- Acitivo,Suspendido,Inactivo,...
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo STATUS de  (" & Reg_Actual(13) & ")  a  (" & Combo7.ListIndex & ")")
    End If
    If Not Reg_Actual(14) = Text11.Text Then             ' Edad
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo EDADP de  (" & Reg_Actual(14) & ")  a  (" & Text11.Text & ")")
    End If
    If Not Reg_Actual(15) = Text13.Text Then       ' Observacion
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo OBSERVACION de  (" & Reg_Actual(15) & ")  a  (" & Text13.Text & ")")
    End If
    If Not Reg_Actual(16) = Text14.Text Then         ' Ocupacion
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo OCUPACION de  (" & Reg_Actual(16) & ")  a  (" & Text14.Text & ")")
    End If
    If Not Reg_Actual(17) = DTPicker1.Value Then        ' Fecha de Registro
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_REGP de  (" & Reg_Actual(17) & ")  a  (" & DTPicker1.Value & ")")
    End If
    If Not Reg_Actual(18) = DTPicker2.Value Then      ' Fecha de Inicio
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_INICIO de  (" & Reg_Actual(18) & ")  a  (" & DTPicker2.Value & ")")
    End If
    If Not Reg_Actual(19) = DTPicker3.Value Then        ' Fecha de Culminacion
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_CULM de  (" & Reg_Actual(19) & ")  a  (" & DTPicker3.Value & ")")
    End If
    If Not Reg_Actual(20) = DTPicker4.Value Then ' Fecha de Nacimiento
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_NACIMIENTOP de  (" & Reg_Actual(20) & ")  a  (" & DTPicker4.Value & ")")
    End If
    
        If Not Reg_Actual(21) = Tipo Then ' Fecha de Nacimiento
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_NACIMIENTOP de  (" & Reg_Actual(20) & ")  a  (" & DTPicker4.Value & ")")
    End If
        If Not Reg_Actual(22) = Mid(CboGrupo.Text, 1, 3) & "00" & Mid(CboGrupo.Text, 6) Then ' Fecha de Nacimiento
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_NACIMIENTOP de  (" & Reg_Actual(20) & ")  a  (" & DTPicker4.Value & ")")
    End If
        If Not Reg_Actual(23) = CboGrupo.Text Then ' Fecha de Nacimiento
        Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Modificar", "Modifico el campo FECHA_NACIMIENTOP de  (" & Reg_Actual(20) & ")  a  (" & DTPicker4.Value & ")")
    End If
    Else
    Call Enviar_Bitacora(IdUser, "Nuevo Paciente", "Guardar", "Guardo un Nuevo Registro Cuya Cedula(CedulaP)=" & Text1.Text & " y Fecha de Ingreso de Registro(FechaRegP)=" & DTPicker1.Value)
End If
Cambio = 0

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización de Paciente"
'*** Actuliza en el Hosting ***
'EnviarAlHosting
Call EnviarRegPendiente(IdPac, IdLIdPac)
Msg = ""
End Sub

Sub EnviarRegPendiente(ByVal IdPac2 As Integer, ByVal IdLIdPac2 As String)
On Error Resume Next

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

JH = Format(DTPicker1.Value, "MM/dd/YYYY")
JH2 = Format(DTPicker2.Value, "MM/dd/YYYY")
JH3 = Format(DTPicker3.Value, "MM/dd/YYYY")
JH4 = Format(DTPicker4.Value, "MM/dd/YYYY")

a = 1


CSql = "SELECT * FROM Paciente WHERE idpaciente = " & IdPac2 & " AND IdL = '" & IdLIdPac2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Paciente (["
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
RsRegPendiente.Fields("Modulo").Value = "Nuevo Paciente"
RsRegPendiente.Fields("Tabla").Value = "Paciente"
RsRegPendiente.Fields("Condicional").Value = "IdPaciente=" & IdPac2 & " AND IdL='" & IdLIdPac2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub


Private Sub BtnAgregarPaciente_Click()
On Error Resume Next
Combo5.SetFocus

Frame6.Visible = False
Frame2.BackColor = &HE0E0E0
Frame4.BackColor = &HE0E0E0
Frame5.BackColor = &HE0E0E0
BtnBorrarPaciente.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
OptPresupuesto.BackColor = &HE0E0E0
OptPaciente.BackColor = &HE0E0E0

OptPresupuesto.Value = True
Select Case st_but

    Case Is = 0
        'Call viryfi
        Blanqueo
        'Command1.Caption = "Guardar Registro"
        'Command6.Enabled = False
        'Command7.Enabled = False
        Actuali = 0
        Cambio = 0
        st_but = 1
        Me.Caption = "Pacientes"
        Label16.Caption = ""
        IdPac = ""
        DTPicker1.Value = Now
        DTPicker2 = Now
        DTPicker3 = Now + 5
        DTPicker4 = Now - 1800
        BtnAgregarPaciente.Enabled = False
        CboGrupo.Text = "00:00"
        Combo7.ListIndex = 6
        ACCION = AGREGAR_REGISTRO
        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        
End Select

Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly, "Error al Guardar"
    
    Exit Sub

End Sub

Private Sub BtnAnterior_Click()
On Error Resume Next
If RsPacientes.RecordCount <> 0 Then
    Call viryfi
    If Not RsPacientes.BOF Then
        RsPacientes.MovePrevious
        If RsPacientes.BOF Then RsPacientes.MoveLast
    End If
    Call BuscarDatos
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
    IdPac = ""
End If
End Sub



Private Sub BtnBorrarPaciente_Click()
On Error Resume Next
Dim RsBorrarPaciente As New ADODB.Recordset

Frame6.Visible = False


If IdPac = "" Then

MsgBox "Seleccione a un paciente"
Exit Sub
End If

Msg = "Esta seguro de Eliminar este Paciente?" & Chr(13) & Chr(13) & Text3.Text & " " & Text4.Text
p = MsgBox(Msg, vbYesNo + vbInformation, "Eliminar Paciente")

If p = vbYes Then

    FrmClaveSupervisor.Show vbModal, FrmPrincipal

    If FrmClaveSupervisor.Acceso Then
        Eliminar
    End If

End If

End Sub

Sub Eliminar()

    CSql = "Update Paciente Set Activo='0' Where IdPaciente = " & IdPac & " And Activo='1' AND IdL='" & IdLIdPac & "'"
    Set RsBorrarPaciente = CrearRS(CSql)

    '**************************************************************************************

    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Actualización de Paciente"
    
    EnviarRegPendiente IdPac, IdLIdPac
    
    Msg = "Fue Eliminado el registro"
    MsgBox Msg, vbInformation + vbOKOnly, "Paciente Eliminado"

    Carga_De_Datos


End Sub


Public Sub BtnBuscar_Click()
On Error Resume Next
Dim CSql As String
'verifica si el registro ha sufrido algun cambio
viryfi
'realiza la limpieza de los campos
Blanqueo
'reemplaza los espacio en blancos por Busqueda de la caja de texto de busqueda
If Replace(TxtBuscar.Text, " ", "") = "" Or Replace(TxtBuscar.Text, " ", "") = "Busqueda" Then
    f = "Buscar"
    ' consulta a todos los pacientes si esta vacia la caja de busqueda
    CSql = "Select * From Paciente Where Activo='1' Order By IdPaciente"
    Set RsPacientes = CrearRS(CSql)
    ' carga todo los datos de los pacientes
    Call Carga_De_Datos
    Exit Sub
End If
'si la caja de busqueda no esta vacia, hace la seleccion del pacientes segun los criterios de busqueda
CSql = "Select * From Paciente Where Activo='1' And (CedulaP = " & Val(TxtBuscar.Text) & " or NombreP like '%" & TxtBuscar.Text & "%' or ApellidoP like '%" & TxtBuscar.Text & "%' or Historia like '%" & TxtBuscar.Text & "%') Order By IdPaciente"
Set RsPacientes = CrearRS(CSql)

If RsPacientes.RecordCount = 0 Then
    MsgBox "No Existe el registro"
    Cambio = 0
    Actuali = 0
    IdPac = ""
    BtnBorrarPaciente.Enabled = False
    NoReg.Caption = "0 / 0"
    Exit Sub
Else
    RsPacientes.MoveFirst
    NoReg.Caption = RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
End If
'busca los datos del paciente
Call BuscarDatos
End Sub

Private Sub BtnBuscar2_Click()
On Error GoTo Most
Dim RsBuscarCliente As New ADODB.Recordset
Dim wer As String

'verifica si los campos esta vacios
wher = ""

If TxtCodigo.Text = "" And TxtNombre.Text = "" And TxtApellido.Text = "" And TxtHistoria.Text = "" And TxtCedula.Text = "" And TxtFechaReg.Text = "" And TxtFechaInicio.Text = "" And TxtFechaCulmi.Text = "" Then
    Msg = "Por favor ingrese Nombre o Apellido o cedula o No. Historia o Fecha de Registro " & Chr(13) & "o Fecha de Inicio del Tratamiento o Fecha de Culminación del Tratamiento del Paciente para realizar la busqueda!!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
Else
'crea la condición de busqueda para la consulta
    If TxtCodigo.Text <> "" Then
        wer = wer & " IdPaciente like '%" & TxtCodigo.Text & "%'"
    End If

    If TxtNombre.Text <> "" Then
        If wer = "" Then
            wer = wer & " NombreP like '%" & TxtNombre.Text & "%'"
        Else
            wer = wer & " And NombreP like '%" & TxtNombre.Text & "%'"
        End If
    End If

    If TxtApellido.Text <> "" Then
        If wer = "" Then
            wer = wer & " ApellidoP like '%" & TxtApellido.Text & "%'"
        Else
            wer = wer & " And ApellidoP like '%" & TxtApellido.Text & "%'"
        End If
    End If

    If TxtHistoria.Text <> "" Then
        If wer = "" Then
            wer = wer & " Historia like '%" & TxtHistoria.Text & "%'"
        Else
            wer = wer & " And Historia like '%" & TxtHistoria.Text & "%'"
        End If
    End If
   
    If TxtCedula.Text <> "" Then
        If wer = "" Then
            wer = wer & " CedulaP like '%" & TxtCedula.Text & "%'"
        Else
            wer = wer & " And CedulaP like '%" & TxtCedula.Text & "%'"
        End If
    End If
    
    If TxtFechaReg.Text <> "" Then
        DtpFechaReg.Value = TxtFechaReg.Text
        If wer = "" Then
            wer = wer & " Fecha_RegP = '" & TxtFechaReg.Text & "'"
        Else
            wer = wer & " And Fecha_RegP = '" & TxtFechaReg.Text & "'"
        End If
    End If
    
    If TxtFechaInicio.Text <> "" Then
        DtpFechaInicio.Value = TxtFechaInicio.Text
        If wer = "" Then
            wer = wer & " Fecha_Inicio = '" & TxtFechaInicio.Text & "'"
        Else
            wer = wer & " And Fecha_Inicio ='" & TxtFechaInicio.Text & "'"
        End If
    End If

    If TxtFechaCulmi.Text <> "" Then
        DtpFechaCulmi.Value = TxtFechaCulmi.Text
        If wer = "" Then
            wer = wer & " Fecha_Culm ='" & TxtFechaCulmi.Text & "'"
        Else
            wer = wer & " And Fecha_Culm = '" & TxtFechaCulmi.Text & "'"
        End If
    End If

End If
'crea la consulta del paciente con la condicion creada
CSql = "Select * From Paciente Where " & wer & " And Activo='1' Order by IdPaciente"
Set RsPacientes = CrearRS(CSql)

If RsPacientes.RecordCount = 0 Then
    MsgBox "No Existe el registro"
    Cambio = 0
    Actuali = 0
    IdPac = ""
    BtnBorrarPaciente.Enabled = False
    NoReg.Caption = "0 / 0"
    Exit Sub
Else
    RsPacientes.MoveFirst
    NoReg.Caption = RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
End If
Call BuscarDatos

FrameBusquedaAvanzada.Visible = False
FrameBusquedaAvanzada.Top = 1440
FrameBusquedaAvanzada.Left = 13200
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Limpiar

Exit Sub
Most:
    MsgBox "La fecha que ha ingresado no es valida!", vbInformation + vbOKOnly, "Informacion"
    TxtFechaReg.Text = ""
    TxtFechaInicio.Text = ""
    TxtFechaCulmi.Text = ""
    
End Sub


Sub Limpiar()

TxtCodigo.Text = ""
TxtNombre.Text = ""
TxtApellido.Text = ""
TxtHistoria.Text = ""
TxtCedula.Text = ""
TxtFechaReg.Text = ""
TxtFechaInicio.Text = ""
TxtFechaCulmi.Text = ""

End Sub

Private Sub BtnBusquedaAvanzada_Click()
FrameBusquedaAvanzada.Visible = True
FrameBusquedaAvanzada.Left = 3720
FrameBusquedaAvanzada.Top = 1440
Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
End Sub

Private Sub BtnCerrar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub BtnCerrar2_Click()
FrameBusquedaAvanzada.Visible = False
FrameBusquedaAvanzada.Top = 1440
FrameBusquedaAvanzada.Left = 13200
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
Carga_De_Datos
Frame2.BackColor = &HEAEFEF
Frame4.BackColor = &HEAEFEF
Frame5.BackColor = &HEAEFEF
OptPresupuesto.BackColor = &HEAEFEF
OptPaciente.BackColor = &HEAEFEF
Frame6.Visible = False
End Sub

Private Sub BtnGenerarHistoriaMedica_Click()
On Error Resume Next
'verifica si el paciente ya posee historia medica asignada
If InStr(1, UCase(Label16.Caption), "HOAM") <> 0 Then
    Msg = "Ya este paciente tiene Un Nº de Historia asignado no se puede generar el Nº de Historia"
    MsgBox Msg, vbOKOnly, "Ya tiene Nº de Historia"
    Exit Sub
Else
'asigna el numero de la historia medica al paciente
    Msg = "Se va a asignar un nuevo Nº de historia a este paciente " & Chr(13) & "Esta Ud. Seguro?"
    b = MsgBox(Msg, vbYesNo + vbInformation, "Numero de Historia")
    
    If b = 6 Then
        ' consulta el ultimo numero de la historia medica
        CSql = "Select Historia + 1 as NoHisto From N_Histo"
        Set RsGenerarHistoria = CrearRS(CSql)

        Histo = RsGenerarHistoria.Fields("NoHisto").Value
        NoHistoria = "HOAM" & Format(Date, "yy") & Format(RsGenerarHistoria.Fields("NoHisto").Value, "00#")
        'muestra el numero de la historia medica del paciente
        Label16.Caption = NoHistoria
        'actualiza el nuevo numero de la historia
        CSql = "Update N_Histo Set Historia = " & Histo & ""
        Set RsNumeroHistoria = CrearRS(CSql)
        
        'actualiza el numero de la historia medica del paciente
        CSql = "Update Paciente Set Historia = '" & NoHistoria & "' Where IdPaciente = " & IdPac
        Set RsUpDateHistoria = CrearRS(CSql)
        
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del No. de Historia"
        ' envia al web la informacion de la historia medica del paciente para ser actualizada
        EnviarRegPendienteNoHistoria
    End If
End If
End Sub
Sub EnviarRegPendienteNoHistoria()
On Error Resume Next
'genera el identificador de los registro a la web
CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

a = 1
'crea la consulta del numero de la historia en la web
CSql = "UPDATE Paciente Set Historia = '" & NoHistoria & "' Where IdPaciente = " & IdPac
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Nuevo Paciente"
RsRegPendiente.Fields("Tabla").Value = "Paciente"
RsRegPendiente.Fields("Condicional").Value = "IdPaciente = " & IdPac
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next

If Frame6.Visible = True Then
    MsgBox "La historia médica debe ser correctamente establecida, por favor verifiquelo!", vbExclamation + vbOKOnly, "Información"
Exit Sub
End If
If OptPaciente.Value Then
    If InStr(1, UCase(Label16.Caption), UCase("hoam")) = 0 Then
        MsgBox "Debe asignarle un número de historia al paciente antes de asignarle una hora de atención!", vbExclamation + vbOKOnly, "Confirmación"
        Exit Sub
    End If
End If

If Replace(Text1.Text, " ", "") = "" Then
    MsgBox "Ingrese el Nro de Cedula del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text1.SetFocus
    Exit Sub
End If

CSql = "Select * from paciente Where cedulap=" & Text1.Text
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 And ACCION = AGREGAR_REGISTRO Then
    MsgBox "La cedula de paciente que ingreso ya se encuentra en la base de datos!", vbCritical + vbOKOnly, "ERROR"
    Exit Sub
End If

If Replace(Combo5.Text, " ", "") = "" Then
    MsgBox "Seleccione la Nacionalidad del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo5.SetFocus
    Exit Sub
ElseIf Replace(Text1.Text, " ", "") = "" Then
    MsgBox "Ingrese el Nro de Cedula del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text1.SetFocus
    Exit Sub
ElseIf Replace(Text4.Text, " ", "") = "" Then
    MsgBox "Ingrese el Nombre del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text4.SetFocus
    Exit Sub
ElseIf Replace(Text3.Text, " ", "") = "" Then
    MsgBox "Ingrese el Apellido del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text3.SetFocus
    Exit Sub
ElseIf Replace(Text11.Text, " ", "") = "" Then
    MsgBox "Ingrese la Edad del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text11.SetFocus
    Exit Sub
ElseIf Replace(Text6.Text, " ", "") = "" Then
    MsgBox "Ingrese la Direccion del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text6.SetFocus
    Exit Sub
ElseIf Replace(Text7.Text, " ", "") = "" Then
    MsgBox "Ingrese el Numero de telefono del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Text7.SetFocus
    Exit Sub
ElseIf Replace(Combo4.Text, " ", "") = "" Then
    MsgBox "El Campo SEXO no debe dejarse en blanco!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo4.SetFocus
    Exit Sub
ElseIf Replace(Combo2.Text, " ", "") = "" Then
    MsgBox "Ingrese el Medico Tratante del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo2.SetFocus
    Exit Sub
ElseIf Replace(Combo3.Text, " ", "") = "" Then
    MsgBox "Ingrese el Medico Remitente del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo3.SetFocus
    Exit Sub
ElseIf Combo7.ItemData(Combo7.ListIndex) = Asc("A") Or Combo7.ItemData(Combo7.ListIndex) = Asc("R") Or Combo7.ItemData(Combo7.ListIndex) = Asc("S") & Asc("i") Then
    If CboGrupo.Text = "" Then
        MsgBox "Seleccione el Grupo de Atención del Paciente!", vbExclamation + vbOKOnly, "Faltan datos!"
        CboGrupo.SetFocus
        Exit Sub
    End If
End If


p = MsgBox("Se procedera a guardar los cambios realizados, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If p = 7 Then Exit Sub

GuardarCambios
'Msg = "Registro Actualizado satisfactoriamente"
'MsgBox Msg, vbInformation + vbOKOnly, "Guardado"

actualiza
Carga_De_Datos

IO = 1
Actuali = 0
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
Frame2.BackColor = &HEAEFEF
Frame4.BackColor = &HEAEFEF
Frame5.BackColor = &HEAEFEF

OptPresupuesto.BackColor = &HEAEFEF
OptPaciente.BackColor = &HEAEFEF
End Sub

Private Sub BtnGuardarHOAM_Click()


If InStr(1, UCase(Text5.Text), UCase("hoam")) = 0 Then
    MsgBox "El formato para el número de historia médica es incorrecto! FORMATO CORRECTO: HOAM#####", vbInformation + vbOKOnly, "Formato incorrecto!"
    Text5.Text = ""
    Exit Sub
ElseIf Not IsNumeric(Mid(Text5.Text, (InStr(1, UCase(Text5.Text), UCase("hoam")) + 4))) Then
    MsgBox "El formato para el número de historia médica es incorrecto! FORMATO CORRECTO: HOAM#####", vbInformation + vbOKOnly, "Formato incorrecto!"
    Text5.Text = ""
    Exit Sub
End If
CSql = "SELECT NombreP,ApellidoP,CedulaP,Historia FROM Paciente"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

While Not RsTemp.EOF
    If UCase(RsTemp.Fields("Historia").Value) = UCase(Trim(Text5.Text)) Then
        MsgBox "La historia que ingreso ya se encuentra en uso por el paciente:" & vbCrLf & vbCrLf & _
            " Nombre: " & Trim(RsTemp.Fields("NombreP").Value) & " " & Trim(RsTemp.Fields("ApellidoP").Value) & vbCrLf & _
            " Cédula: " & Trim(RsTemp.Fields("CedulaP").Value) & vbCrLf & _
            " Historia: " & Trim(RsTemp.Fields("Historia").Value) & vbCrLf & vbCrLf & _
            "Ingrese un número de historia diferente!", vbExclamation + vbOKOnly, "Información: La historia se encuentra en uso!"
        Exit Sub
    End If
    RsTemp.MoveNext
Wend

Label16.Caption = UCase(Trim(Text5.Text))
Frame6.Visible = False
End Sub

Private Sub BtnListadoPaciente_Click()
On Error Resume Next
Tipo = "NuevoPaciente"
FrmListadoPaciente.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnNuevoMedico_Click()
On Error Resume Next
FrmRegistroMedicoRemitente.Show vbModal, FrmPrincipal
'carga_lista_medicosr
End Sub

Private Sub BtnSiguiente_Click()
On Error Resume Next
If RsPacientes.RecordCount <> 0 Then
    Call viryfi
    If Not RsPacientes.EOF Then
        RsPacientes.MoveNext
        If RsPacientes.EOF Then RsPacientes.MoveFirst
    End If
    Call BuscarDatos
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If

End Sub

Private Sub BtnTomarFotoPaciente_Click()
On Error Resume Next
If Text1.Text <> "" And Text3.Text <> "" And Text4.Text <> "" Then
    If (Validar_Camara("CapWindow", ws_child Or ws_visible, 0, 0, 340, 240, Picture1.Hwnd, 0)) Then
        FrmCapturarFoto.Show vbModal, FrmPrincipal
    Else
        MsgBox "No hay Camaras Webs Instaladas", vbOKOnly + vbCritical, "Error"
    End If
Else
    MsgBox "Debe de Ingresar o Seleccionar a un Paciente para poder tomar la Foto", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If
End Sub

Private Sub ChameleonBtn1_Click()
Frame6.Visible = False
End Sub

Private Sub Combo7_Click()
Cambio = 1
End Sub

Sub viryfi()
'verifica si el registro a tenido algun cambio
Select Case Cambio
Case Is = 1
    Msg = "Este registro sufrió Cambios desea guardar?"
    d = MsgBox(Msg, vbYesNo, "Desea Guardar Cambios")
    Select Case d
    Case Is = 6
        Call GuardarCambios
    Case Is = 7
    End Select
    Case Is = 0
End Select

BtnAgregarPaciente.Enabled = True
BtnBorrarPaciente.Enabled = True
Actuali = 0
st_but = 0

End Sub

Sub carga_lista_medicost()
On Error Resume Next
Dim RsMedico As New ADODB.Recordset

CSql = "SELECT * FROM medicos where Activo=1"
Set RsMedico = CrearRS(CSql)

Combo2.Clear
RsMedico.MoveFirst

Do While Not RsMedico.EOF
    
    If RsMedico.Fields("Tipo").Value = "2" Or RsMedico.Fields("Tipo").Value = "3" Then
        Combo2.AddItem RsMedico.Fields("Nombre").Value & " " & RsMedico.Fields("Apellido").Value
        Combo2.ItemData(Combo2.NewIndex) = RsMedico.Fields("idmedico").Value
    End If
    RsMedico.MoveNext
Loop

RsMedico.Close
End Sub
Sub carga_lista_medicosr()
On Error Resume Next

Dim RsCargaListaMedico As New ADODB.Recordset

CSql = "SELECT * FROM medicos where Tipo=1 OR Tipo=3"
Set RsCargaListaMedico = CrearRS(CSql)
RsCargaListaMedico.MoveFirst
Combo3.Clear

Do While Not RsCargaListaMedico.EOF

    If RsCargaListaMedico.Fields("Tipo").Value = "1" Or RsCargaListaMedico.Fields("Tipo").Value = "3" Then
        Combo3.AddItem RsCargaListaMedico.Fields("Nombre").Value & " " & RsCargaListaMedico.Fields("Apellido").Value
        Combo3.ItemData(Combo3.NewIndex) = RsCargaListaMedico.Fields("idmedico").Value
    End If
    RsCargaListaMedico.MoveNext
Loop

RsCargaListaMedico.Close
End Sub

Sub Blanqueo()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text11.Text = ""
Image3.Picture = LoadPicture()
FotoP = ""
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Combo3.ListIndex = -1
Combo4.ListIndex = -1
Combo5.ListIndex = -1
Combo6.ListIndex = -1
Text13.Text = ""
Text14.Text = ""
DTPicker2.Value = Now
DTPicker3.Value = Now
DTPicker4.Value = Now
DTPicker1.Value = Now

End Sub

Private Sub Combo1_Click()
Cambio = 1
End Sub

Private Sub Combo2_Click()
Cambio = 1
End Sub

Private Sub Combo3_Click()
Cambio = 1
End Sub

Private Sub Combo4_Click()
Cambio = 1
End Sub

Private Sub Combo5_Click()
Cambio = 1
End Sub

Private Sub Combo6_Click()
Cambio = 1
End Sub





Private Sub DtpFechaCulmi_Change()
TxtFechaCulmi.Text = Format(DtpFechaCulmi.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaInicio_Change()
TxtFechaInicio.Text = Format(DtpFechaInicio.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaReg_Change()
TxtFechaReg.Text = Format(DtpFechaReg.Value, "dd/mm/yyyy")
End Sub

Private Sub DTPicker1_Change()
Cambio = 1
End Sub
Private Sub DTPicker2_Change()
Cambio = 1
End Sub
Private Sub DTPicker3_Change()
If CDate(DTPicker2.Value) > CDate(DTPicker3.Value) Then
    MsgBox "La Fecha de culminación no puede ser menor que la fecha de inicio!" & vbCrLf & _
        "Por favor corrija las fechas...", vbExclamation + vbOKOnly, "Error en la fechas!"
    DTPicker3.Value = DTPicker2.Value
End If
Cambio = 1
End Sub


Private Sub DTPicker4_Change()
Cambio = 1
Text11.Text = DateDiff("yyyy", DTPicker4.Value, Now)
End Sub

Private Sub DTPicker4_Click()
Text11.Text = DateDiff("yyyy", DTPicker4.Value, Now)
End Sub

Private Sub Form_Activate()
Tipo = "Nuevo Paciente"
Me.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim RsCodigos As New ADODB.Recordset
Dim CSql As String

Centrar Me
Bi = 1
IO = 0


DTPicker1.Value = DateTime.Date
DTPicker2.Value = DateTime.Date
DTPicker3.Value = DateTime.Date



Combo4.AddItem "Masculino"
Combo4.ItemData(Combo4.NewIndex) = 0
Combo4.AddItem "Femenino"
Combo4.ItemData(Combo4.NewIndex) = 1


Combo7.AddItem "Activo"
Combo7.ItemData(Combo7.NewIndex) = Asc("A")
Combo7.AddItem "Suspendido"
Combo7.ItemData(Combo7.NewIndex) = Asc("S")
Combo7.AddItem "Reubicación"
Combo7.ItemData(Combo7.NewIndex) = Asc("R")
Combo7.AddItem "Simulación"
Combo7.ItemData(Combo7.NewIndex) = Asc("S") & Asc("i")
Combo7.AddItem "Espera"
Combo7.ItemData(Combo7.NewIndex) = Asc("E")
Combo7.AddItem "Fallecido"
Combo7.ItemData(Combo7.NewIndex) = Asc("F")
Combo7.AddItem "Inactivo"
Combo7.ItemData(Combo7.NewIndex) = Asc("I")
Combo7.AddItem "Culminado"
Combo7.ItemData(Combo7.NewIndex) = Asc("C")



If T_U = 1 Then
    BtnGuardarActualizar.Enabled = False
Else
    BtnGuardarActualizar.Enabled = True
End If

T = 0
st_but = 0

CSql = "Select * From Paciente Order by IdPaciente"
Set RsPacientes = CrearRS(CSql)

CSql = "SELECT * FROM Codigos_T"
Set RsCodigos = CrearRS(CSql)
Combo1.Clear
Combo6.Clear

Combo1.AddItem " "
Combo1.ItemData(Combo1.NewIndex) = 0
    
Combo6.AddItem " "
Combo6.ItemData(Combo6.NewIndex) = 0
    
    
While Not RsCodigos.EOF
    Combo1.AddItem RsCodigos.Fields("codigo").Value
    Combo1.ItemData(Combo1.NewIndex) = RsCodigos.Fields("idcodigot").Value
        
    Combo6.AddItem RsCodigos.Fields("codigo").Value
    Combo6.ItemData(Combo6.NewIndex) = RsCodigos.Fields("idcodigot").Value
    RsCodigos.MoveNext
Wend

Combo5.AddItem "V"
Combo5.ItemData(Combo5.NewIndex) = 0
Combo5.AddItem "E"
Combo5.ItemData(Combo5.NewIndex) = 1

carga_lista_medicost
carga_lista_medicosr
'Carga_De_Datos
   
End Sub
Sub Carga_De_Datos()
On Error Resume Next

Dim RsRutaFotoEmpleados As New ADODB.Recordset
Dim RutaFoto As String
CSql = "Select * From Dat_Admin"
Set RsRutaFotoEmpleados = CrearRS(CSql)
RutaFoto = RsRutaFotoEmpleados.Fields("RutaFotos").Value
RsRutaFotoEmpleados.Close

IdLIdPac = ""
CSql = "Select * From Paciente Where Activo='1' Order by IdPaciente"
Set RsPacientes = CrearRS(CSql)
Frame6.Visible = False
If RsPacientes.EOF Or RsPacientes.BOF Then
    'Rs.MoveFirst
    Actuali = 0
    Exit Sub
End If

    Me.Caption = "Pacientes - Id: " & RsPacientes.Fields("IdPaciente").Value
    Actuali = 0
    Cambio = 0
    st_but = 0
    BtnAgregarPaciente.Enabled = True
    BtnBorrarPaciente.Enabled = True
    BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
NoReg.Caption = RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
BuscarDatos

End Sub

Sub BuscarDatos()
On Error Resume Next
' busca todos los datos de los pacientes

' cambia la bandera accion para poder modificar los datos del paciente
ACCION = EDITAR_REGISTRO
Frame6.Visible = False
If RsPacientes.EOF Or RsPacientes.BOF Then
    'RsPacientes.MoveFirst
    Actuali = 0
    IdLIdPac = ""
    'RsPacientes.MoveFirst
Else
    Me.Caption = "Pacientes - Id: " & RsPacientes.Fields("IdPaciente").Value
    Text1.Text = RsPacientes.Fields("cedulap").Value
    IdLIdPac = RsPacientes.Fields("IdL").Value
    DTPicker1.Value = RsPacientes.Fields("Fecha_regp").Value
    DTPicker2.Value = RsPacientes.Fields("Fecha_Inicio").Value
    DTPicker3.Value = RsPacientes.Fields("Fecha_culm").Value
    DTPicker4.Value = RsPacientes.Fields("Fecha_Nacimientop").Value
    
    If Not IsNull(RsPacientes.Fields("Celular").Value) Then Text2.Text = Format(RsPacientes.Fields("Celular").Value, "000000#") Else Text2.Text = ""
    
    Text4.Text = RsPacientes.Fields("Nombrep").Value
    Text3.Text = RsPacientes.Fields("Apellidop").Value
    If Not IsNull(RsPacientes.Fields("Direccionp").Value) Then Text6.Text = RsPacientes.Fields("Direccionp").Value Else Text6.Text = ""
    
    If Not IsNull(RsPacientes.Fields("Telefono").Value) Then Text7.Text = Format(RsPacientes.Fields("Telefono").Value, "000000#") Else Text7.Text = ""
    
    Select Case Trim(RsPacientes.Fields("status").Value)
        Case Is = "A"
            Combo7.Text = "Activo"
        Case Is = "S"
            Combo7.Text = "Suspendido"
        Case Is = "R"
            Combo7.Text = "Reubicación"
        Case Is = "Si"
            Combo7.Text = "Simulación"
        Case Is = "E"
            Combo7.Text = "Espera"
        Case Is = "I"
            Combo7.Text = "Inactivo"
        Case Is = "C"
            Combo7.Text = "Culminado"
        Case Is = "F"
            Combo7.Text = "Fallecido"
    End Select
              
    Select Case RsPacientes.Fields("Tipo").Value
        Case Is = 0
            OptPresupuesto.Value = True
        Case Is = 1
            OptPaciente.Value = True
    End Select
    
    Actuali = 1
    IdPac = RsPacientes.Fields("IdPaciente").Value

    If Not IsNull(RsPacientes.Fields("foto").Value) Then
        If RsPacientes.Fields("foto").Value <> "" And Dir(Foto & "\" & RsPacientes.Fields("foto").Value) <> "" Then
            Image3.Picture = LoadPicture(Foto & "\" & RsPacientes.Fields("foto").Value)
            FotoP = RsPacientes.Fields("foto").Value
        Else
            Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
            FotoP = ""
        End If
    Else
        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        FotoP = ""
    End If
    
    Text11.Text = RsPacientes.Fields("Edadp").Value
    If IsNull(RsPacientes.Fields("Ocupacion").Value) Then Text14.Text = "" Else Text14.Text = RsPacientes.Fields("Ocupacion").Value
    If Trim(RsPacientes.Fields("Historia").Value) <> "" Then Label16.Caption = RsPacientes.Fields("Historia").Value Else Label16.Caption = "C" & RsPacientes.Fields("CedulaP").Value
    IdPac = RsPacientes.Fields("IdPaciente").Value
    If RsPacientes.Fields("Sexop").Value = 0 Then Combo4.Text = "Masculino" Else Combo4.Text = "Femenino"
    If RsPacientes.Fields("Observacion").Value <> "" Then Text13.Text = RsPacientes.Fields("Observacion").Value Else Text13.Text = ""
    
    For T = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(T) = RsPacientes.Fields("Codigo").Value Then
            Combo1.ListIndex = T
            Exit For
        End If
    Next T

    For T = 0 To Combo2.ListCount - 1
        If Combo2.ItemData(T) = RsPacientes.Fields("Medico_Tratante").Value Then
            Combo2.ListIndex = T
            Exit For
        End If
    Next T
    
    For T = 0 To Combo3.ListCount - 1
        If Combo3.ItemData(T) = RsPacientes.Fields("Medico_Remitente").Value Then
            Combo3.ListIndex = T
            Exit For
        End If
    Next T
    
    For T = 0 To Combo5.ListCount - 1
        If Combo5.ItemData(T) = RsPacientes.Fields("Nacionalidad").Value Then
            Combo5.ListIndex = T
            Exit For
        End If
    Next T
    For T = 0 To Combo6.ListCount - 1
        If Combo6.ItemData(T) = RsPacientes.Fields("Codigoc").Value Then
            Combo6.ListIndex = T
            Exit For
        End If
    Next T
    For T = 0 To Combo4.ListCount - 1
        If Combo4.ItemData(T) = RsPacientes.Fields("sexop").Value Then
            Combo4.ListIndex = T
            Exit For
        End If
    Next T
    NoReg.Caption = RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
    BtnBorrarPaciente.Enabled = True
End If
'se crea un arreglo para verificar si el paciente ha tenido algun cambio en alguno de sus datos
Reg_Actual(0) = Text1.Text ' No Cedula
Reg_Actual(1) = Text2.Text ' No Telefono Mov
Reg_Actual(2) = Text3.Text ' Apellido
Reg_Actual(3) = Text4.Text ' Nombre
Reg_Actual(4) = Text6.Text ' Direccion
Reg_Actual(5) = Text7.Text ' No Telefono Hab.
Reg_Actual(6) = FotoP      ' Foto Paciente

If Combo1.ListIndex = -1 Then
    Reg_Actual(7) = "-1"
    Else
    Reg_Actual(7) = Combo1.ItemData(Combo1.ListIndex)  ' Cod Area Tel. Hab
End If
If Combo2.ListIndex = -1 Then
    Reg_Actual(8) = "-1"
    Else
    Reg_Actual(8) = Combo2.ItemData(Combo2.ListIndex)  ' Medico T.
End If
If Combo3.ListIndex = -1 Then
    Reg_Actual(9) = "-1"
    Else
    Reg_Actual(9) = Combo3.ItemData(Combo3.ListIndex)  ' Medico R.
End If
If Combo4.ListIndex = -1 Then
    Reg_Actual(10) = "-1"
    Else
    Reg_Actual(10) = Combo4.ItemData(Combo4.ListIndex)  ' Sexo
End If
If Combo5.ListIndex = -1 Then
    Reg_Actual(11) = "-1"
    Else
    Reg_Actual(11) = Combo5.ItemData(Combo5.ListIndex)  ' Nacionalidad
End If
If Combo6.ListIndex = -1 Then
    Reg_Actual(12) = "-1"
    Else
    Reg_Actual(12) = Combo6.ItemData(Combo6.ListIndex)  ' Cod Area Tel. Mov
End If
If Combo7.ListIndex = -1 Then
    Reg_Actual(13) = "-1"
    Else
    Reg_Actual(13) = Combo7.ItemData(Combo7.ListIndex)  ' Estatus- Acitivo,Suspendido,Inactivo,...
End If
Reg_Actual(14) = Text11.Text      ' Edad
Reg_Actual(15) = Text13.Text      ' Observacion
Reg_Actual(16) = Text14.Text      ' Ocupacion
Reg_Actual(17) = DTPicker1.Value  ' Fecha de Registro
Reg_Actual(18) = DTPicker2.Value  ' Fecha de Inicio
Reg_Actual(19) = DTPicker3.Value  ' Fecha de Culminacion
Reg_Actual(20) = DTPicker4.Value  ' Fecha de Nacimiento


Reg_Actual(21) = Tipo  ' tipo
Reg_Actual(22) = Mid(CboGrupo.Text, 1, 3) & "00" & Mid(CboGrupo.Text, 6) ' Grupo
Reg_Actual(23) = CboGrupo.Text   ' Hora Atencion

Cambio = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
' se cierran todos los estados de conexion con los recordset
If RsPacientes.State = 1 Then RsPacientes.Close: Exit Sub
If RsNumeroHistoria.State = 1 Then RsNumeroHistoria.Close
If RsUpDateHistoria.State = 1 Then RsUpDateHistoria.Close
If RsGenerarHistoria.State = 1 Then RsGenerarHistoria.Close
End Sub

Private Sub Image3_Click()
On Error Resume Next
Dim TempCad As String
Dim TempCad2 As String
On Error GoTo h
'carga la foto del paciente desde un archivo
CommonDialog1.ShowOpen
TempCad = CommonDialog1.filename

If InStr(1, TempCad, "\", vbTextCompare) = 0 Then
    If FotoP = "" Then FotoP = "Silueta.jpg"
    Exit Sub
End If

FotoP = Replace(Trim(Text1.Text) & Trim(Text3.Text) & Trim(Text4.Text) & ".jpg", " ", "")
TempCad2 = Foto & "\" & FotoP
Call FileCopy(TempCad, TempCad2)

If Trim(FotoP) = "" Then Exit Sub
Image3.Picture = LoadPicture(TempCad2)
Image3.Refresh
Cambio = 1
Exit Sub
h:
MsgBox Err.Description
End Sub

Private Sub Label16_Change()
On Error Resume Next
'If Trim(Label16.Caption) = "" Or Trim(Label16.Caption) = "no tiene" Or Trim(Label16.Caption) = "NO TIENE" Then
If InStr(1, UCase(Label16), "HOAM") = 0 Then
    BtnGenerarHistoriaMedica.Enabled = True
Else
    BtnGenerarHistoriaMedica.Enabled = False
End If

End Sub

Private Sub Label16_DblClick()
Dim Rsp As String
    
Rsp = InputBox("Ingrese la clave para poder establecer el número de la Historia Médica", "Ingrese su clave universal!")
    
If IsNull(Rsp) Then Exit Sub
If IsEmpty(Rsp) Then Exit Sub
If Trim(Rsp) = "" Then Exit Sub

Text5.Text = ""

CSql = "SELECT ClaveGlobal FROM Usuarios WHERE IdUsuario=" & IdUser
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

If IsNull(RsTemp.Fields("ClaveGlobal").Value) Then
    MsgBox "No se le ha asignado una clave global, contacte al administrador!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
ElseIf IsEmpty(RsTemp.Fields("ClaveGlobal").Value) Then
    MsgBox "No se le ha asignado una clave global, contacte al administrador!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
ElseIf Trim(RsTemp.Fields("ClaveGlobal").Value) = "" Then
    MsgBox "No se le ha asignado una clave global, contacte al administrador!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If
If Trim(RsTemp.Fields("ClaveGlobal").Value) = Trim(Rsp) Then
    Frame6.Left = 120
    Frame6.Top = 240
    Frame6.Visible = True
Else
    MsgBox "Clave incorrecta!", vbCritical + vbOKOnly, "Acceso denegado!"
End If
    
End Sub

Private Sub OptPaciente_Click()
If OptPaciente.Value = True Then
    'CboGrupo.Text = "08:00 A.M."
    CboGrupo.Enabled = True
End If
End Sub

Private Sub OptPresupuesto_Click()
If OptPresupuesto.Value = True Then
    CboGrupo.Text = "00:00"
    CboGrupo.Enabled = False
End If
End Sub

Private Sub Text1_LostFocus()
On Error Resume Next

If InStr(1, Label16.Caption, "HOAM") = 0 Then 'And InStr(1, Label16.Caption, "C") = 0 Then
    Label16.Caption = "C" & Trim(Text1.Text)
ElseIf BtnAgregarPaciente.Enabled = False Then
    Label16.Caption = "C" & Trim(Text1.Text)
End If
    
If BtnAgregarPaciente.Enabled = True Then Exit Sub
If Replace(Text1.Text, " ", "") = "" Then
    Exit Sub
End If

CSql = "Select * From Paciente Where CedulaP=" & Text1.Text & ""
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    MsgBox "La cedula de paciente que ingreso ya se encuentra en la base de datos!", vbCritical + vbOKOnly, "ERROR"
    Text1.SetFocus
    Text1.Text = ""
End If

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If Len(Trim(Text13.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Verificar_Mayuscula Text3, KeyCode
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_GotFocus()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
End Sub


Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
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

Private Sub TxtBuscar_LostFocus()
If Trim(TxtBuscar.Text) = "" Then TxtBuscar.Text = "Busqueda"
End Sub

Private Sub Text7_Change()
Cambio = 1
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
'KeyAscii = SoloNumeros(KeyAscii)
End Sub

'Private Sub Timer2_Timer()
'If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
'If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
'End Sub

Private Sub Text1_Change()
Cambio = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
ElseIf KeyAscii <> 8 Then
          
    If Not IsNumeric(Chr(KeyAscii)) Then
              Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text11_Change()
Cambio = 1
End Sub

Private Sub Text13_Change()
Cambio = 1
End Sub

Private Sub Text2_Change()
Cambio = 1
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0

End Select
End Sub

Private Sub Text3_Change()
Cambio = 1
End Sub

Private Sub Text14_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text14.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text14.Text)
    pru = LCase(Mid(Text14.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
   End If
Next i

Text14.Text = StrText
Text14.SelStart = Len(Text14.Text)
End Sub

Private Sub Text4_Change()
'Text4_KeyUp 16, 1
Cambio = 1
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Verificar_Mayuscula Text4, KeyCode
End Sub

Private Sub Text6_Change()
Cambio = 1
If Len(Trim(Text6.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If
End Sub


Sub Verificar_Mayuscula(ByVal Cadena As TextBox, KeyCode As Integer)
On Error Resume Next
Dim BuffTem As Byte

If Len(Cadena.Text) = 0 Then Exit Sub
If Cadena.SelStart = 0 And Len(Cadena.Text) <= 0 Then Exit Sub

    BuffTem = Cadena.SelStart
    Cadena.Text = UCase(Mid(Cadena.Text, 1, 1)) & Mid(Cadena.Text, 2)
    For conti = 1 To Len(Cadena.Text)
        If Mid(Cadena.Text, conti, 1) = " " Then
            Cadena.Text = Mid(Cadena.Text, 1, conti) & UCase(Mid(Cadena.Text, conti + 1, 1)) & Mid(Cadena.Text, conti + 2)
        Else
            Cadena.Text = Mid(Cadena.Text, 1, conti) & LCase(Mid(Cadena.Text, conti + 1, 1)) & Mid(Cadena.Text, conti + 2)
        End If
    Next
    Cadena.SelStart = BuffTem

End Sub

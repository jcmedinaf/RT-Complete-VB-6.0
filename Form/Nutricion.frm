VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmHistorialNutricional 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historia Nutricional"
   ClientHeight    =   10425
   ClientLeft      =   3675
   ClientTop       =   1200
   ClientWidth     =   13185
   Icon            =   "Nutricion.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   82
      Top             =   8760
      Width           =   12975
      Begin ChamaleonButton.ChameleonBtn BtnVerInforme 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
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
         MICON           =   "Nutricion.frx":1002
         PICN            =   "Nutricion.frx":101E
         PICH            =   "Nutricion.frx":12A7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnEvolucionNutricional 
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Nutricion.frx":16E7
         PICN            =   "Nutricion.frx":1703
         PICH            =   "Nutricion.frx":199B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior2 
         Height          =   375
         Left            =   11520
         TabIndex        =   17
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
         MICON           =   "Nutricion.frx":1C22
         PICN            =   "Nutricion.frx":1C3E
         PICH            =   "Nutricion.frx":1ED3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente2 
         Height          =   375
         Left            =   12240
         TabIndex        =   18
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
         MICON           =   "Nutricion.frx":212F
         PICN            =   "Nutricion.frx":214B
         PICH            =   "Nutricion.frx":23E1
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
         Left            =   5280
         TabIndex        =   106
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "Nutricion.frx":2640
         PICN            =   "Nutricion.frx":265C
         PICH            =   "Nutricion.frx":28EF
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
         Left            =   6960
         TabIndex        =   107
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Nutricion.frx":2B78
         PICN            =   "Nutricion.frx":2B94
         PICH            =   "Nutricion.frx":2FBF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnInformeSemanal 
         Height          =   375
         Left            =   1560
         TabIndex        =   109
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Informe Semanal"
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
         MICON           =   "Nutricion.frx":3257
         PICN            =   "Nutricion.frx":3273
         PICH            =   "Nutricion.frx":34FC
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
      Height          =   735
      Left            =   3480
      TabIndex        =   76
      Top             =   9600
      Width           =   9615
      Begin ChamaleonButton.ChameleonBtn BtnAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   21
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
         MICON           =   "Nutricion.frx":37A0
         PICN            =   "Nutricion.frx":37BC
         PICH            =   "Nutricion.frx":3949
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
         Left            =   6360
         TabIndex        =   24
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
         MICON           =   "Nutricion.frx":3B7E
         PICN            =   "Nutricion.frx":3B9A
         PICH            =   "Nutricion.frx":3E30
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
         Left            =   5760
         TabIndex        =   23
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
         MICON           =   "Nutricion.frx":408F
         PICN            =   "Nutricion.frx":40AB
         PICH            =   "Nutricion.frx":4340
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   8520
         TabIndex        =   26
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
         MICON           =   "Nutricion.frx":459C
         PICN            =   "Nutricion.frx":45B8
         PICH            =   "Nutricion.frx":4781
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
         Left            =   7320
         TabIndex        =   25
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
         MICON           =   "Nutricion.frx":49B6
         PICN            =   "Nutricion.frx":49D2
         PICH            =   "Nutricion.frx":4CB4
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
         Left            =   2520
         TabIndex        =   22
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
         MICON           =   "Nutricion.frx":4F05
         PICN            =   "Nutricion.frx":4F21
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
         TabIndex        =   108
         ToolTipText     =   "Guardar / Actualiza Evaluación Dietética"
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
         MICON           =   "Nutricion.frx":50C5
         PICN            =   "Nutricion.frx":50E1
         PICH            =   "Nutricion.frx":5370
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
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro"
         Height          =   195
         Left            =   3720
         TabIndex        =   80
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   75
      Top             =   9600
      Width           =   3255
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
         ForeColor       =   &H8000000A&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido,Cédula de identidad o Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         ToolTipText     =   "Buscar Pacientes por Nombre, Apellido, Cedula o Historia"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "Nutricion.frx":57B1
         PICN            =   "Nutricion.frx":57CD
         PICH            =   "Nutricion.frx":5A32
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11040
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Paciente"
      Height          =   1935
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   12975
      Begin VB.ComboBox CboSexo 
         Height          =   315
         Left            =   5760
         TabIndex        =   74
         Top             =   720
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DtpFechaActual 
         Height          =   315
         Left            =   5760
         TabIndex        =   41
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47513601
         CurrentDate     =   39818
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1120
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   750
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpFechaRegistro 
         Height          =   315
         Left            =   5760
         TabIndex        =   72
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47513601
         CurrentDate     =   39818
      End
      Begin MSComCtl2.DTPicker DtpFechaNac 
         Height          =   315
         Left            =   1320
         TabIndex        =   73
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47513601
         CurrentDate     =   39818
      End
      Begin ChamaleonButton.ChameleonBtn BtnLlamar 
         Height          =   375
         Left            =   7800
         TabIndex        =   77
         ToolTipText     =   "Llamar"
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Nutricion.frx":5CC4
         PICN            =   "Nutricion.frx":5CE0
         PICH            =   "Nutricion.frx":5F7C
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
         Left            =   7800
         TabIndex        =   78
         ToolTipText     =   "Lista de Espera"
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Nutricion.frx":61B1
         PICN            =   "Nutricion.frx":61CD
         PICH            =   "Nutricion.frx":6456
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
         Left            =   7800
         TabIndex        =   81
         ToolTipText     =   "Desocupar al Paciente Atendido"
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Nutricion.frx":66EE
         PICN            =   "Nutricion.frx":670A
         PICH            =   "Nutricion.frx":68AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Historia:"
         Height          =   195
         Left            =   4800
         TabIndex        =   79
         Top             =   450
         Width           =   870
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   11040
         Picture         =   "Nutricion.frx":6AE3
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Actual:"
         Height          =   195
         Left            =   4680
         TabIndex        =   66
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label Label14 
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
         Left            =   5760
         TabIndex        =   61
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro:"
         Height          =   195
         Left            =   4320
         TabIndex        =   51
         Top             =   1500
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre(s):"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1210
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellido(s):"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nac.:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1620
         Width           =   1110
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad:"
         Height          =   195
         Left            =   3000
         TabIndex        =   47
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         Height          =   195
         Left            =   5160
         TabIndex        =   46
         Top             =   780
         Width           =   405
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Evaluación Nutricional"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   2160
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Evaluación Clínica"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   2160
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame FrmGrl 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Evaluación Clínica"
      Height          =   6135
      Index           =   0
      Left            =   120
      TabIndex        =   53
      Top             =   2640
      Width           =   12975
      Begin VB.TextBox DNI 
         Height          =   735
         Left            =   6480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   98
         Top             =   600
         Width           =   6375
      End
      Begin VB.TextBox Recom 
         Height          =   855
         Left            =   6480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   1215
         Left            =   120
         TabIndex        =   86
         Top             =   4800
         Width           =   12735
         Begin VB.TextBox TxtPiel 
            Height          =   375
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   11
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox TxtOjos 
            Height          =   375
            Left            =   9960
            MaxLength       =   50
            TabIndex        =   13
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox TxtLengua 
            Height          =   375
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   12
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox TxtAbdomen 
            Height          =   375
            Left            =   9960
            MaxLength       =   50
            TabIndex        =   14
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox TxtPesoUsual 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtPorcenPeso 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox TxtPesoActual 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   2880
            MaxLength       =   5
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtPesoRefer 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2880
            MaxLength       =   5
            TabIndex        =   9
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox TxtIMC 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4320
            MaxLength       =   5
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtTalla 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   8
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Piel:"
            Height          =   195
            Left            =   5685
            TabIndex        =   96
            Top             =   360
            Width           =   300
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ojos:"
            Height          =   195
            Left            =   9480
            TabIndex        =   95
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lengua:"
            Height          =   195
            Left            =   5400
            TabIndex        =   94
            Top             =   810
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abdomen:"
            Height          =   195
            Left            =   9120
            TabIndex        =   93
            Top             =   810
            Width           =   720
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso Usual:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% Peso:"
            Height          =   195
            Left            =   3720
            TabIndex        =   91
            Top             =   810
            Width           =   570
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso Actual:"
            Height          =   195
            Left            =   1920
            TabIndex        =   90
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso Refer.:"
            Height          =   195
            Left            =   1920
            TabIndex        =   89
            Top             =   840
            Width           =   885
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IMC:"
            Height          =   195
            Left            =   3960
            TabIndex        =   88
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Talla:"
            Height          =   195
            Left            =   600
            TabIndex        =   87
            Top             =   840
            Width           =   390
         End
      End
      Begin VB.TextBox TxtInteraccionDrogasNutricion 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   4200
         Width           =   12735
      End
      Begin VB.TextBox TxtEstadoGeneral 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   3480
         Width           =   11535
      End
      Begin VB.TextBox TxtMedicamentos 
         Height          =   735
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2640
         Width           =   11535
      End
      Begin VB.TextBox Text9 
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1680
         Width           =   6255
      End
      Begin VB.TextBox Text8 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnóstico Nutricional Integral ó Evaluación Global Subjetiva"
         Height          =   195
         Left            =   6480
         TabIndex        =   100
         Top             =   360
         Width           =   4380
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recomendaciones Nutricionales:"
         Height          =   195
         Left            =   6480
         TabIndex        =   99
         Top             =   1440
         Width           =   2340
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interacción  Droga - Nutrición:"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   3960
         Width           =   2130
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado General:"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   3570
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medicamentos:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnóstico Definitivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnóstico de Ingreso:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.Frame FrmGrl 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Evaluacion Dietetica"
      Height          =   6135
      Index           =   1
      Left            =   120
      TabIndex        =   58
      Top             =   2640
      Width           =   12975
      Begin VB.TextBox TxtSal 
         Height          =   300
         Left            =   10200
         TabIndex        =   39
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox TxtGrasas 
         Height          =   300
         Left            =   10200
         TabIndex        =   35
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox TxtEnlatados 
         Height          =   300
         Left            =   10200
         TabIndex        =   37
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox TxtEmbutidos 
         Height          =   300
         Left            =   10200
         TabIndex        =   36
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox TxtRefrescos 
         Height          =   300
         Left            =   10200
         TabIndex        =   38
         Top             =   3360
         Width           =   2655
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   8520
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Historia Clinica Integral"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox TxtHarinas 
         Height          =   300
         Left            =   10200
         TabIndex        =   34
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox TxtMerienda 
         Height          =   2415
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "Nutricion.frx":9E64
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox TxtCarnes 
         Height          =   300
         Left            =   10200
         TabIndex        =   32
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtCena 
         Height          =   2415
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "Nutricion.frx":9E7E
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox TxtGCCA 
         Height          =   300
         Left            =   10200
         TabIndex        =   33
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox TxtAlmuerzo 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "Nutricion.frx":9E99
         Top             =   3600
         Width           =   4455
      End
      Begin VB.TextBox TxtGCCD 
         Height          =   300
         Left            =   10200
         TabIndex        =   31
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox TxtDesayuno 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "Nutricion.frx":9EB2
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sal:"
         Height          =   195
         Left            =   9360
         TabIndex        =   105
         Top             =   3773
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grasas:"
         Height          =   195
         Left            =   9360
         TabIndex        =   104
         Top             =   2333
         Width           =   540
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embutidos:"
         Height          =   195
         Left            =   9360
         TabIndex        =   103
         Top             =   2693
         Width           =   780
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enlatados:"
         Height          =   195
         Left            =   9360
         TabIndex        =   102
         Top             =   3053
         Width           =   750
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refrescos:"
         Height          =   195
         Left            =   9360
         TabIndex        =   101
         Top             =   3413
         Width           =   765
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cena:"
         Height          =   195
         Left            =   4680
         TabIndex        =   71
         Top             =   3360
         Width           =   420
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merienda:"
         Height          =   195
         Left            =   4680
         TabIndex        =   70
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almuerzo:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   3360
         Width           =   690
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desayuno:"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harina:"
         Height          =   195
         Left            =   9360
         TabIndex        =   65
         Top             =   1973
         Width           =   510
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aguas:"
         Height          =   195
         Left            =   9360
         TabIndex        =   64
         Top             =   1613
         Width           =   495
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carnes:"
         Height          =   195
         Left            =   9360
         TabIndex        =   63
         Top             =   1253
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dulces:"
         Height          =   195
         Left            =   9360
         TabIndex        =   62
         Top             =   893
         Width           =   540
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Recodatorio de 24 Horas "
         Height          =   255
         Left            =   3960
         TabIndex        =   59
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Noreg2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluación: 0 / 0"
      Height          =   195
      Left            =   10875
      TabIndex        =   85
      Top             =   2280
      Width           =   2190
   End
End
Attribute VB_Name = "FrmHistorialNutricional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPaciente As New ADODB.Recordset 'tabla Registro
Dim RsNutricion As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim SQL As String
Dim Cambio
Dim NuevoReg
Dim IdNut As String
Dim NuevoId As String
Dim IdPacH
Dim IdLIdPacH As String
Dim IdLIdInfH As String
Dim IdLIdInf As String
Dim IdInf As Integer

Sub Blanqueo()

TxtCena.Text = ""
TxtAlmuerzo.Text = ""
TxtDesayuno.Text = ""
TxtMerienda.Text = ""
TxtEstadoGeneral.Text = ""
TxtInteraccionDrogasNutricion.Text = ""
TxtLengua.Text = ""
TxtOjos.Text = ""
TxtPiel.Text = ""
TxtMedicamentos.Text = ""
TxtTalla.Text = ""
TxtAbdomen.Text = ""
TxtIMC.Text = ""
TxtPesoActual.Text = ""
TxtPesoRefer.Text = ""
TxtPesoUsual.Text = ""
TxtPorcenPeso.Text = ""
TxtGrasas.Text = ""
TxtEnlatados.Text = ""
TxtEmbutidos.Text = ""
TxtSal.Text = ""
TxtRefrescos.Text = ""
TxtGCCD.Text = ""
TxtCarnes.Text = ""
TxtHarinas.Text = ""
TxtGCCA.Text = ""
DNI.Text = ""
Recom.Text = ""
End Sub

Private Sub BtnAgregar_Click()
On Error Resume Next

IdLIdInf = IdLDefault

If Trim(IdPacH) = "" Then MsgBox "Debe seleccionar un Paciente antes de agregar un registro!", vbExclamation + vbOKOnly, "Seleecione un Paciente": Exit Sub

Call Blanqueo
NuevoReg = 1
BtnGuardarActualizar.Enabled = True
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnInformeSemanal.Enabled = False
Text10.SetFocus
FrmGrl(0).BackColor = &HE0E0E0
FrmGrl(1).BackColor = &HE0E0E0
Frame6.BackColor = &HE0E0E0

End Sub

Private Sub BtnAntecedentes_Click()
On Error Resume Next
If IdPacH = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub
FrmAntecedentes.IdPacA = IdPacH
FrmAntecedentes.IdLIdPacA = IdLIdPacH
FrmAntecedentes.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAnterior_Click()
If Trim(IdPacH) = "" Then Exit Sub
If RsPaciente.RecordCount <> 0 Then
    If BtnAgregar.Enabled = False Then BtnDesHacer_Click
    RsPaciente.MovePrevious
    If RsPaciente.BOF Then RsPaciente.MoveLast
    Call Blanqueo
    Call Carga_De_Datos
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda...", vbExclamation + vbOKOnly, "No hay datos!"
End If
End Sub

Private Sub BtnAnterior2_Click()
On Error Resume Next
If Trim(IdPacH) = "" Then Exit Sub
If RsNutricion.RecordCount <> 0 Then
    RsNutricion.MovePrevious
    If RsNutricion.BOF Then MsgBox "Ha llegado al primer registro!", vbExclamation + vbOKOnly, "Primer Registro!": RsNutricion.MoveFirst
    Call CargaDatos(False)
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda...", vbExclamation + vbOKOnly, "No hay datos!"
End If
End Sub

Public Sub BtnBuscar_Click()
On Error Resume Next
If Trim(TxtBuscar.Text) = "" Then
    CSql = "Select * From Paciente Order By IdPaciente"
Else
    CSql = "Select * From Paciente Where Activo='1' And (Cedulap = " & Val(TxtBuscar.Text) & " OR nombrep like '%" & TxtBuscar.Text & "%' OR apellidop like '%" & TxtBuscar.Text & "%' or Historia = '" & UCase(TxtBuscar.Text) & "')"
End If
Set RsPaciente = CrearRS(CSql)

Call Blanqueo
BtnDesHacer_Click
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Label14.Caption = ""
DtpFechaRegistro = Now
DtpFechaActual = Now
DtpFechaNac.Value = Now

If RsPaciente.RecordCount = 0 Then
    NoReg = "Registro: 0 / 0"
    MsgBox "No se encontro el registro buscado!", vbExclamation + vbOKOnly, "No hay resultados!"
    IdPacH = ""
    IdLIdPacH = IdLDefault
    BtnExamenes.Enabled = False
    'BtnDiagnostico.Enabled = False
    BtnAntecedentes.Enabled = False
    BtnGuardarActualizar.Enabled = False
    BtnVerInforme.Enabled = False
    BtnSiguiente2.Enabled = False
    BtnAnterior2.Enabled = False
    BtnInformeSemanal.Enabled = False
    BtnEliminar.Enabled = False
    BtnAgregar.Enabled = False
    BtnEvolucionNutricional.Enabled = False
    Exit Sub
End If
    BtnExamenes.Enabled = True
    'BtnDiagnostico.Enabled = True
    BtnAntecedentes.Enabled = True
    BtnGuardarActualizar.Enabled = True
    BtnVerInforme.Enabled = True
    BtnSiguiente2.Enabled = True
    BtnAnterior2.Enabled = True
    BtnInformeSemanal.Enabled = True
    BtnEliminar.Enabled = True
    BtnAgregar.Enabled = True
    BtnEvolucionNutricional.Enabled = True
    Call Carga_De_Datos
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
FrmGrl(0).BackColor = &HEAEFEF
FrmGrl(1).BackColor = &HEAEFEF
Frame6.BackColor = &HEAEFEF
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
BtnInformeSemanal.Enabled = True
Call Blanqueo
If RsPaciente.RecordCount <> 0 Then
    Call Carga_De_Datos
End If
End Sub

Private Sub BtnDesocuparAlPacienteAtendido_Click()
On Error Resume Next
Dim DBLista88 As New ADODB.Recordset
CSql = "Delete From Ubi_Paciente Where modul = " & ModulO
Set DBLista88 = CrearRS(CSql)
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next

If IdPacH = "" Then MsgBox "Estimado usuario, debe seleccionar un paciente!", vbExclamation + vbOKOnly, "Disculpe": Exit Sub
If IdNut = "" Then MsgBox "Estimado usuario, debe haber un registro nutricional seleccionado para poder Eliminar!", vbExclamation + vbOKOnly, "Disculpe.": Exit Sub

resp = MsgBox("Desea eliminar el Informe Nutricional # " & RsNutricion.AbsolutePosition & ", del Paciente de C.I.: " & Text1.Text & " ?", vbQuestion + vbYesNo, "Confirmar!")
    
If resp = vbNo Then Exit Sub


' Consulta en la tabla "Nutricion"

CSql = "Select * From Nutricion Where IdNutricion='" & IdNut & "' And IdL='" & IdLIdInfH & "'"
Set RsVerificar = CrearRS(CSql)

' Verifica que el usuario creador del Informe Nutricional y el usuario que inició sesion es el mismo,
' de ser asi entonces procede a eliminar el Informe Nutricional, o tambien si es un Administrador.
If RsVerificar.Fields("IdUsuario").Value = IdUser Or T_U = 0 Then
    
    CSql = "UPDATE Nutricion SET Activo=0 Where IdNutricion=" & IdNut & " AND IdL = '" & IdLIdInfH & "'"
    Set RsTemp = CrearRS(CSql)
    MsgBox "El Registro Nutricional fue Eliminado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
    
    ' Envia la consulta a la tabla de Registros pendientes.
    EnviarRegPendiente IdNut, IdLIdInfH
    Call Blanqueo
    Call Carga_De_Datos
Else
    MsgBox "Estimado usuario, Usted No tiene permiso para borrar este Informe Nutricional", vbCritical + vbOKOnly, "Disculpe."
End If
End Sub

Private Sub BtnEvolucionNutricional_Click()
On Error Resume Next

Especia = "Nutricion"
FrmEvolucion.IdPacE = IdPacH
FrmEvolucion.IdLIdPacE = IdLIdPacH
FrmEvolucion.Show vbModal, FrmPrincipal

End Sub

Private Sub BtnExamenes_Click()

If IdPacH = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub
    
FrmExamenHematologico.IdPacH = IdPacH
FrmExamenHematologico.IdLIdPacH = IdLIdPacH
FrmExamenHematologico.Show vbModal, FrmPrincipal

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error GoTo MostrarError
Dim RsTemp As New ADODB.Recordset

Dim resp

If IdPacH = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub

'command2
If Recom.Text = "" Then
    Msg = "El Campo Recomendaciones Dietéticas esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    Recom.SetFocus
    Exit Sub
ElseIf TxtMedicamentos.Text = "" Then
    Msg = "El Campo Medicamentos esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtMedicamentos.SetFocus
    Exit Sub
ElseIf TxtEstadoGeneral.Text = "" Then
    Msg = "El Campo Estado General esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtEstadoGeneral.SetFocus
    Exit Sub
ElseIf TxtInteraccionDrogasNutricion.Text = "" Then
    Msg = "El Campo Interacción Droga - Nutrición esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtInteraccionDrogasNutricion.SetFocus
    Exit Sub
ElseIf TxtPesoUsual.Text = "" Then
    Msg = "El Campo Peso Usual esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtPesoUsual.SetFocus
    Exit Sub
ElseIf TxtPesoActual.Text = "" Then
    Msg = "El Campo Peso Actual esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtPesoActual.SetFocus
    Exit Sub
ElseIf TxtIMC.Text = "" Then
    Msg = "El Campo Indice de Masa Corporal esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtIMC.SetFocus
    Exit Sub
ElseIf TxtTalla.Text = "" Then
    Msg = "El Campo Talla esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtTalla.SetFocus
    Exit Sub
ElseIf TxtPesoRefer.Text = "" Then
    Msg = "El Campo Peso Referencial esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtPesoRefer.SetFocus
    Exit Sub
ElseIf TxtPorcenPeso.Text = "" Then
    Msg = "El Campo Porcentaje de Peso esta Vacio"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtPorcenPeso.SetFocus
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  If TxtBuscar.Text = "Depurar" Then MsgBox "Antes de crear el archivo: " & "Y:\" & Replace(Date, "/", "") & " NutricionInf.txt"
  
  Open "Y:\" & Replace(Date, "/", "") & " NutricionInf.txt" For Append As #1
  MError = "MMMMMMMMMMMMMMMMMMMM" & Date & " : " & Time & "MMMMMMMMMMMMMMMM" & Chr(13) & Chr(10) & "RsTemp.Fields('IdNutricion').Value = " & NuevoId & Chr(13) & Chr(10) & "RsTemp.Fields('IdL').Value = " & IdLIdInf & Chr(13) & Chr(10) & _
          "RsTemp.Fields('IdUsuario').Value = " & IdUser & Chr(13) & Chr(10) & "RsTemp.Fields('IdPaciente').Value = " & IdPacH & Chr(13) & Chr(10) & _
          "RsTemp.Fields('IdLIdPac').Value = " & IdLIdPacH & Chr(13) & Chr(10) & _
          "RsTemp.Fields('FechaNu').Value = " & Format(DtpFechaActual.Value, "mm/dd/yyyy") & Chr(13) & Chr(10) & "RsTemp.Fields('Medicamento').Value = " & Trim(TxtMedicamentos.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Estado').Value = " & Trim(TxtEstadoGeneral.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Interaccion').Value = " & Trim(TxtInteraccionDrogasNutricion.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Piel').Value = " & Trim(TxtPiel.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Ojos').Value = " & Trim(TxtOjos.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Lengua').Value = " & Trim(TxtLengua.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Abdomen').Value = " & Trim(TxtAbdomen.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Desayuno').Value = " & Trim(TxtDesayuno.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('GCCD').Value = " & Trim(TxtGCCD.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Almuerzo').Value = " & Trim(TxtAlmuerzo.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('GCCA').Value = " & Trim(TxtGCCA.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Cena1').Value = " & Trim(TxtCena.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('GCC1').Value = " & Trim(TxtCarnes.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Cena2').Value = " & Trim(TxtMerienda.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('GCC2').Value = " & Trim(TxtHarinas.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('PesoU').Value = " & Trim(TxtPesoUsual.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('CambioP').Value = " & Trim(TxtPorcenPeso.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('PesoA').Value = " & Trim(TxtPesoActual.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('PesoR').Value = " & Trim(TxtPesoRefer.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Grasas').Value = " & Trim(TxtGrasas.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Embutidos').Value = " & Trim(TxtEmbutidos.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Enlatados').Value = " & Trim(TxtEnlatados.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Refrescos').Value = " & Trim(TxtRefrescos.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Sal').Value = " & Trim(TxtSal.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Indice').Value = " & TxtIMC.Text & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Talla').Value = " & TxtTalla.Text & Chr(13) & Chr(10) & "RsTemp.Fields('DNI').Value = " & Trim(DNI.Text) & Chr(13) & Chr(10) & _
          "RsTemp.Fields('Recomendaciones').Value = " & Trim(Recom.Text) & Chr(13) & Chr(10) & "RsTemp.Fields('Activo').Value = 1" & Chr(13) & Chr(10) & _
          "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM" & Chr(13) & Chr(10) & "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
          
  If TxtBuscar.Text = "Depurar" Then MsgBox "Antes de abrir el archivo: " & "Y:\" & Replace(Date, "/", "") & " NutricionInf.txt"
  Print #1, MError
  If TxtBuscar.Text = "Depurar" Then MsgBox "Despues de escribir en el archivo: " & "Y:\" & Replace(Date, "/", "") & " NutricionInf.txt"
  Close #1
  If TxtBuscar.Text = "Depurar" Then MsgBox "Despues de cerrar el archivo: " & "Y:\" & Replace(Date, "/", "") & " NutricionInf.txt"
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

resp = MsgBox("Se procedera a guardar los cambios realizados! Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

If NuevoReg = 1 Then Cambio = 3

CSql = "SELECT MAX(IdNutricion)+1 as NuevoId FROM Nutricion"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    NuevoId = RsTemp.Fields("NuevoId").Value
Else
    NuevoId = "1"
End If
Set RsTemp = Nothing

Select Case Cambio
    
    Case Is = 1 'Actualiza
        
        'CSql = "Select * From Nutricion Where IdNutricion='" & IdNut & "' And IdPaciente='" & IdPacH & "' AND IdLIdPac = '" & IdLIdPacH & "' And IdL='" & IdLIdInfH & "'"
        CSql = "Select * From Nutricion Where IdNutricion='" & IdNut & "' And IdL='" & IdLIdInfH & "'"
        Set RsTemp = CrearRS(CSql)
        
        RsTemp.Fields("IdUsuario").Value = IdUser
        RsTemp.Fields("FechaNu").Value = Format(DtpFechaActual.Value, "mm/dd/yyyy")
        RsTemp.Fields("Medicamento").Value = Trim(TxtMedicamentos.Text)
        RsTemp.Fields("Estado").Value = Trim(TxtEstadoGeneral.Text)
        RsTemp.Fields("Interaccion").Value = Trim(TxtInteraccionDrogasNutricion.Text)
        RsTemp.Fields("Piel").Value = Trim(TxtPiel.Text)
        RsTemp.Fields("Ojos").Value = Trim(TxtOjos.Text)
        RsTemp.Fields("Lengua").Value = Trim(TxtLengua.Text)
        RsTemp.Fields("Abdomen").Value = Trim(TxtAbdomen.Text)
        RsTemp.Fields("Desayuno").Value = Trim(TxtDesayuno.Text)
        RsTemp.Fields("GCCD").Value = Trim(TxtGCCD.Text)
        RsTemp.Fields("Almuerzo").Value = Trim(TxtAlmuerzo.Text)
        RsTemp.Fields("GCCA").Value = Trim(TxtGCCA.Text)
        RsTemp.Fields("Cena1").Value = Trim(TxtCena.Text)
        RsTemp.Fields("GCC1").Value = Trim(TxtCarnes.Text)
        RsTemp.Fields("Cena2").Value = Trim(TxtMerienda.Text)
        RsTemp.Fields("GCC2").Value = Trim(TxtHarinas.Text)
        RsTemp.Fields("PesoU").Value = Trim(TxtPesoUsual.Text)
        RsTemp.Fields("CambioP").Value = Trim(TxtPorcenPeso.Text)
        RsTemp.Fields("PesoA").Value = Trim(TxtPesoActual.Text)
        RsTemp.Fields("PesoR").Value = Trim(TxtPesoRefer.Text)
        RsTemp.Fields("Grasas").Value = Trim(TxtGrasas.Text)
        RsTemp.Fields("Embutidos").Value = Trim(TxtEmbutidos.Text)
        RsTemp.Fields("Enlatados").Value = Trim(TxtEnlatados.Text)
        RsTemp.Fields("Refrescos").Value = Trim(TxtRefrescos.Text)
        RsTemp.Fields("Sal").Value = Trim(TxtSal.Text)
        RsTemp.Fields("Indice").Value = TxtIMC.Text
        RsTemp.Fields("Talla").Value = TxtTalla.Text
        RsTemp.Fields("DNI").Value = Trim(DNI.Text)
        RsTemp.Fields("Recomendaciones").Value = Trim(Recom.Text)
        RsTemp.Fields("Activo").Value = 1
        RsTemp.Update

        MsgBox "El Registro fue Actualizado Exitosamente", vbInformation + vbOKOnly, "Registro Actualizado Satifactoriamente"
    
        EnviarRegPendiente IdNut, IdLIdInfH

    Case Is = 3 'Agrega un nuevo registro para un informe nutricional...
    
        IdLIdInfH = IdLDefault

        CSql = "Select * From Nutricion"
        Set RsTemp = CrearRS(CSql)
        
        RsTemp.AddNew
        
        RsTemp.Fields("IdNutricion").Value = NuevoId
        RsTemp.Fields("IdL").Value = IdLIdInfH
        RsTemp.Fields("IdUsuario").Value = IdUser
        RsTemp.Fields("IdPaciente").Value = IdPacH
        RsTemp.Fields("FechaNu").Value = Format(DtpFechaActual.Value, "mm/dd/yyyy")
        RsTemp.Fields("Medicamento").Value = Trim(TxtMedicamentos.Text)
        RsTemp.Fields("Estado").Value = Trim(TxtEstadoGeneral.Text)
        RsTemp.Fields("Interaccion").Value = Trim(TxtInteraccionDrogasNutricion.Text)
        RsTemp.Fields("Piel").Value = Trim(TxtPiel.Text)
        RsTemp.Fields("Ojos").Value = Trim(TxtOjos.Text)
        RsTemp.Fields("Lengua").Value = Trim(TxtLengua.Text)
        RsTemp.Fields("Abdomen").Value = Trim(TxtAbdomen.Text)
        RsTemp.Fields("Desayuno").Value = Trim(TxtDesayuno.Text)
        RsTemp.Fields("GCCD").Value = Trim(TxtGCCD.Text)
        RsTemp.Fields("Almuerzo").Value = Trim(TxtAlmuerzo.Text)
        RsTemp.Fields("GCCA").Value = Trim(TxtGCCA.Text)
        RsTemp.Fields("Cena1").Value = Trim(TxtCena.Text)
        RsTemp.Fields("GCC1").Value = Trim(TxtCarnes.Text)
        RsTemp.Fields("Cena2").Value = Trim(TxtMerienda.Text)
        RsTemp.Fields("GCC2").Value = Trim(TxtHarinas.Text)
        RsTemp.Fields("PesoU").Value = Trim(TxtPesoUsual.Text)
        RsTemp.Fields("CambioP").Value = Trim(TxtPorcenPeso.Text)
        RsTemp.Fields("PesoA").Value = Trim(TxtPesoActual.Text)
        RsTemp.Fields("PesoR").Value = Trim(TxtPesoRefer.Text)
        RsTemp.Fields("Grasas").Value = Trim(TxtGrasas.Text)
        RsTemp.Fields("Embutidos").Value = Trim(TxtEmbutidos.Text)
        RsTemp.Fields("Enlatados").Value = Trim(TxtEnlatados.Text)
        RsTemp.Fields("Refrescos").Value = Trim(TxtRefrescos.Text)
        RsTemp.Fields("Sal").Value = Trim(TxtSal.Text)
        RsTemp.Fields("Indice").Value = TxtIMC.Text
        RsTemp.Fields("Talla").Value = TxtTalla.Text
        RsTemp.Fields("DNI").Value = Trim(DNI.Text)
        RsTemp.Fields("Recomendaciones").Value = Trim(Recom.Text)
        RsTemp.Fields("Activo").Value = 1
        RsTemp.Update

        MsgBox "Registro Agregado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        
        EnviarRegPendiente NuevoId, IdLIdInfH
        
        Call Blanqueo
        Call Carga_De_Datos
    Case Is = 0
        MsgBox "No hay registros seleccionados", vbInformation + vbOKOnly, "Seleccione un Registro"
        Exit Sub
End Select

NuevoReg = 0
Cambio = 0

BtnDesHacer_Click
Exit Sub

MostrarError:
    MsgBox "Es produjeron algunos detalles en los procesos, contacte con el Administrador!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
    
End Sub

Sub EnviarRegPendiente(ByVal IdNut As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM NUTRICION WHERE IdNutricion = " & IdNut & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO NUTRICION (["
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
RsRegPendiente.Fields("Modulo").Value = "Historial Nutricional"
RsRegPendiente.Fields("Tabla").Value = "Nutricion"
RsRegPendiente.Fields("Condicional").Value = "IdNutricion=" & IdNut & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


MsgBox "Datos agregados en la lista de actualizacion!", vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

Private Sub BtnInformeSemanal_Click()
On Error Resume Next
If IdPacH = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub
FrmReporteNutricion.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnListaEspera_Click()
On Error Resume Next
ModulO = 0
FrmListaEspera.Show
If Cedul <> "" Then
    TxtBuscar.Text = Cedul
    Call BtnBuscar_Click
End If
End Sub

Private Sub BtnLlamar_Click()
On Error Resume Next
Call Llamar
End Sub

Private Sub BtnSiguiente_Click()
If Trim(IdPacH) = "" Then Exit Sub
If RsPaciente.RecordCount <> 0 Then
    If BtnAgregar.Enabled = False Then BtnDesHacer_Click
    RsPaciente.MoveNext
    If RsPaciente.EOF Then RsPaciente.MoveFirst
    Call Blanqueo
    Call Carga_De_Datos
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda...", vbExclamation + vbOKOnly, "No hay datos!"
End If
End Sub

Sub Carga_De_Datos()
IdLIdPacH = IdLDefault
IdPacH = ""
IdLIdInf = IdLDefault
IdInf = 0
BtnEvolucionNutricional.Enabled = False
If RsPaciente.RecordCount = 0 Then Exit Sub

BtnEvolucionNutricional.Enabled = True
If Trim(RsPaciente.Fields("cedulap")) <> "" Then Text1.Text = RsPaciente.Fields("cedulap")
If Trim(RsPaciente.Fields("Fecha_regp")) <> "" Then DtpFechaRegistro.Value = RsPaciente.Fields("Fecha_regp")
If Trim(RsPaciente.Fields("Nombrep")) <> "" Then Text3.Text = RsPaciente.Fields("Nombrep")
If Trim(RsPaciente.Fields("Apellidop")) <> "" Then Text4.Text = RsPaciente.Fields("Apellidop")
If Trim(RsPaciente.Fields("Fecha_nacimientop")) <> "" Then DtpFechaNac.Value = RsPaciente.Fields("Fecha_nacimientop")
If Trim(RsPaciente.Fields("Edadp")) <> "" Then Text6.Text = RsPaciente.Fields("Edadp")
If Trim(RsPaciente.Fields("Historia")) <> "" Then Label14.Caption = RsPaciente.Fields("Historia")
IdPacH = RsPaciente.Fields("Idpaciente")
IdLIdPacH = RsPaciente.Fields("IdL").Value
If RsPaciente.Fields("foto") <> "" Then
    If Len(Dir(Foto & "\" & RsPaciente.Fields("foto"))) > 0 Then
        Image2.Picture = LoadPicture(Foto & "\" & RsPaciente.Fields("foto"))
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
Else
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
End If

Me.Caption = "Nutricion - Paciente: " & IdPacH
NoReg = "Registro " & RsPaciente.AbsolutePosition & " / " & RsPaciente.RecordCount

If RsPaciente.Fields("sexop") = 0 Then CboSexo.Text = "Masculino" Else CboSexo.Text = "Femenino"

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Consulta que trae los informes médicos del paciente
CSql = "Select * From Informe_medico where idpaciente = " & IdPacH & " And IdLIdPac='" & IdLIdPacH & "' AND Estado=1"
Set RsTemp = CrearRS(CSql)

IdInf = Val(RsTemp.Fields("IdInforme"))
IdLIdInf = RsTemp.Fields("IdL")
If RsTemp.RecordCount <> 0 Then
    If RsTemp.Fields("motivo_con") = "" Then Text8.Text = "" Else Text8.Text = RsTemp.Fields("motivo_con")
    If RsTemp.Fields("diagnotico") = "" Then Text9.Text = "" Else Text9.Text = RsTemp.Fields("diagnotico")
Else
    Text8.Text = "Este paciente no tiene informe médico."
    Text9.Text = "Este paciente no tiene informe médico."
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Call CargaDatos(True)
    
End Sub
Sub CargaDatos(Band As Boolean)

TxtMedicamentos.Text = ""
IdLIdInfH = IdLDefault
If Band Then
    CSql = "SELECT * FROM NUTRICION WHERE IDPACIENTE = " & IdPacH & " AND IdLIdPac='" & IdLIdPacH & "' AND Activo=1 ORDER BY IdNutricion"
    Set RsNutricion = CrearRS(CSql)
    IdLIdInfH = IdLDefault
    IdNut = ""
    Cambio = 0
End If

If RsNutricion.RecordCount <> 0 Then
    'DTPicker1.Value = Trim(RsNutricion.Fields("FechaNu"))
    IdNut = RsNutricion.Fields("IdNutricion").Value
    IdLIdInfH = RsNutricion.Fields("IdL").Value
    If RsNutricion.Fields("MEDICAMENTO") <> "" Then TxtMedicamentos.Text = RsNutricion.Fields("MEDICAMENTO") Else TxtMedicamentos.Text = ""
    If RsNutricion.Fields("PesoU") <> "" Then TxtPesoUsual.Text = RsNutricion.Fields("PesoU") Else TxtPesoUsual.Text = ""
    If RsNutricion.Fields("ESTADO") <> "" Then TxtEstadoGeneral.Text = RsNutricion.Fields("ESTADO") Else TxtEstadoGeneral.Text = ""
    If RsNutricion.Fields("CambioP") <> "" Then TxtPorcenPeso.Text = RsNutricion.Fields("CambioP") Else TxtPorcenPeso.Text = ""
    If RsNutricion.Fields("INTERACCION") <> "" Then TxtInteraccionDrogasNutricion.Text = RsNutricion.Fields("INTERACCION") Else TxtInteraccionDrogasNutricion.Text = ""
    If RsNutricion.Fields("PIEL") <> "" Then TxtPiel.Text = RsNutricion.Fields("PIEL") Else TxtPiel.Text = ""
    If RsNutricion.Fields("OJOS") <> "" Then TxtOjos.Text = RsNutricion.Fields("OJOS") Else TxtOjos.Text = ""
    If RsNutricion.Fields("LENGUA") <> "" Then TxtLengua.Text = RsNutricion.Fields("Lengua") Else TxtLengua.Text = ""
    If RsNutricion.Fields("ABDOMEN") <> "" Then TxtAbdomen.Text = RsNutricion.Fields("ABDOMEN") Else TxtAbdomen.Text = ""
    If RsNutricion.Fields("DESAYUNO") <> "" Then TxtDesayuno.Text = RsNutricion.Fields("DESAYUNO") Else TxtDesayuno.Text = ""
    If RsNutricion.Fields("GCCD") <> "" Then TxtGCCD.Text = RsNutricion.Fields("GCCD") Else TxtGCCD.Text = ""
    If RsNutricion.Fields("ALMUERZO") <> "" Then TxtAlmuerzo.Text = RsNutricion.Fields("ALMUERZO") Else TxtAlmuerzo.Text = ""
    'If RsNutricion.Fields("ALIMENTOA") <> "" Then Text26Text = RsNutricion.Fields("ALIMENTOA") Else Text26.Text = ""
    If RsNutricion.Fields("GCCA") <> "" Then TxtGCCA.Text = RsNutricion.Fields("GCCA") Else TxtGCCA.Text = ""
    If RsNutricion.Fields("CENA1") <> "" Then TxtCena.Text = RsNutricion.Fields("CENA1") Else TxtCena.Text = ""
    If RsNutricion.Fields("GCC1") <> "" Then TxtCarnes.Text = RsNutricion.Fields("GCC1") Else TxtCarnes.Text = ""
    If RsNutricion.Fields("CENA2") <> "" Then TxtMerienda.Text = RsNutricion.Fields("CENA2") Else TxtMerienda.Text = ""
    If RsNutricion.Fields("GCC2") <> "" Then TxtHarinas.Text = RsNutricion.Fields("GCC2") Else TxtHarinas.Text = ""
    If RsNutricion.Fields("PesoA") <> "" Then TxtPesoActual.Text = RsNutricion.Fields("PesoA") Else TxtPesoActual.Text = ""
    If RsNutricion.Fields("Talla") <> "" Then TxtTalla.Text = RsNutricion.Fields("Talla") Else TxtTalla.Text = ""
    If RsNutricion.Fields("PesoR") <> "" Then TxtPesoRefer.Text = RsNutricion.Fields("PesoR") Else TxtPesoRefer.Text = ""
    If RsNutricion.Fields("Indice") <> "" Then TxtIMC.Text = RsNutricion.Fields("Indice") Else TxtIMC.Text = ""
    If Trim(RsNutricion.Fields("DNI")) <> "" Then DNI.Text = RsNutricion.Fields("DNI") Else DNI.Text = ""
    If Trim(RsNutricion.Fields("Recomendaciones")) <> "" Then Recom.Text = RsNutricion.Fields("Recomendaciones") Else Recom.Text = ""
    
    If RsNutricion.Fields("Grasas") <> "" Then TxtGrasas.Text = RsNutricion.Fields("Grasas") Else TxtGrasas.Text = ""
    If RsNutricion.Fields("Embutidos") <> "" Then TxtEmbutidos.Text = RsNutricion.Fields("Embutidos") Else TxtEmbutidos.Text = ""
    If RsNutricion.Fields("Enlatados") <> "" Then TxtEnlatados.Text = RsNutricion.Fields("Enlatados") Else TxtEnlatados.Text = ""
    If Trim(RsNutricion.Fields("Refrescos")) <> "" Then TxtRefrescos.Text = RsNutricion.Fields("Refrescos") Else TxtRefrescos.Text = ""
    If Trim(RsNutricion.Fields("Sal")) <> "" Then TxtSal.Text = RsNutricion.Fields("Sal") Else TxtSal.Text = ""
    
    NuevoReg = 0
    BtnEliminar.Enabled = True
    BtnAnterior2.Enabled = True
    BtnSiguiente2.Enabled = True
    Noreg2.Caption = "Evaluacion " & RsNutricion.AbsolutePosition & " / " & RsNutricion.RecordCount
    Cambio = 1
Else
    IdNut = ""
    IdLIdInfH = ""
    NuevoReg = 1
    BtnEliminar.Enabled = False
    BtnAnterior2.Enabled = False
    BtnSiguiente2.Enabled = False
    BtnGuardarActualizar.Enabled = False
    Noreg2.Caption = "Evaluacion 0 / 0"
    Cambio = 0
End If
End Sub
 
 
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' De aqui en adelante es codigo de interacciones, y uno de impresion
  Private Sub BtnSiguiente2_Click()
  On Error Resume Next
  If Trim(IdPacH) = "" Then Exit Sub
    If RsNutricion.RecordCount <> 0 Then
        RsNutricion.MoveNext
        If RsNutricion.EOF Then MsgBox "Ha llegado al Final del registro!", vbExclamation + vbOKOnly, "Ultimo Registro!": RsNutricion.MoveLast
        Call CargaDatos(False)
    Else
        MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda...", vbExclamation + vbOKOnly, "No hay datos!"
    End If
  End Sub

Private Sub BtnVerInforme_Click()
'On Error Resume Next

If IdPacH = "" Then MsgBox "Debe seleccionar un Paciente antes de agregar un registro!", vbExclamation + vbOKOnly, "Seleecione un Paciente": Exit Sub
If Text1.Text = "" Then
    MsgBox "Debe de seleccionar un Paciente", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If

' ========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\HistoriaClinicaIntegralN.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Historia_Clinica.IdNutricion} = " & IdNut & " And {Historia_Clinica.IdL} = '" & IdLIdInfH & "'"
    
    .WindowTitle = "Reporte Historia Medica No. " & Label12.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
End Sub

Private Sub Command6_Click()
If IdPacH <> "" Then
    FrmAntecedentes.Show 1
Else
    MsgBox "Estimado usuario, debe seleccionar un paciente", vbExclamation + vbOKOnly, "Disculpe."
End If
End Sub


Private Sub DNI_Change()
Cambio = 1
End Sub

Private Sub Option1_Click(Index As Integer)
For i = 0 To Option1.Count - 1
    If i <> Index Then
        Option1(i).Value = False
            'FrameDato(i).Visible = False
        FrmGrl(i).Visible = False
    Else
        Option1(i).Value = True
            'FrameDato(i).Visible = True
        FrmGrl(i).Visible = True
    End If
Next
End Sub


Private Sub Text10_KeyPress(KeyAscii As Integer)
If Len(Trim(Text10.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
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

Private Sub Text15_KeyPress(KeyAscii As Integer)
If Len(Trim(Text15.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If Len(Trim(Text16.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Trim(Text17.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If Len(Trim(Text18.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If Len(Trim(Text19.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If Len(Trim(Text21.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If Len(Trim(Text22.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If Len(Trim(Text23.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
If Len(Trim(Text26.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
If Len(Trim(Text27.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
If Len(Trim(Text29.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If
End Sub


Private Sub Text30_KeyPress(KeyAscii As Integer)
If Len(Trim(Text30.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
If Len(Trim(Text31.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
If Len(Trim(Text33.Text)) = 0 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
Else
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Recom_Change()
Cambio = 1
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
Centrar Me

ModulO = 0
If RsPaciente.State = 1 Then RsPaciente.Close
SQL = "Select * From Paciente Order by IdPaciente"
Set RsPaciente = CrearRS(SQL)

NuevoReg = 0
DtpFechaActual.Value = Now()
SQL = ""
IdPacH = ""

Call Blanqueo
'Call Carga_De_Datos
Cambio = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim resp
If Cambio <> 0 And IdPacH <> "" Then
resp = MsgBox("Ha hecho modificaciones al registro actual desea guardar estos cambios?", vbQuestion + vbYesNo, "Guardar cambios")

If resp = vbYes Then Call BtnGuardarActualizar_Click

End If

If RsPaciente.State Then RsPaciente.Close
IdPacH = ""
SQL = ""

End Sub

Private Sub Text7_Click()
If Text7.Text = "Busqueda" Then Text7.Text = ""
End Sub

Private Sub Text7_GotFocus()
If Text7.Text = "Busqueda" Then Text7.Text = ""
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 48 To 57 ' permite el ingreso de numeros
Case Is = 13  ' permite presionar el ENTER

Call BtnBuscar_Click

Case Is = 8  ' Permite Borrar de retroceso
Case Else    ' Inhibe todas las demas teclas

End Select

End Sub

Private Sub Text7_LostFocus()
If Trim(Text7.Text) = "" Then Text7.Text = "Busqueda"

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
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

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub Text34_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub Text35_KeyPress(KeyAscii As Integer)
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
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub
Private Sub Text10_Change()
Cambio = 1
End Sub

Private Sub Text11_Change()
Cambio = 1
For i = 1 To Len(Text11.Text)
h = Trim(Mid(Text11.Text, i, 1))
If h = "," Then
h = Chr(46)
End If
f = f & h
Next i
Text11.Text = f
End Sub

Private Sub Text13_Change()
Cambio = 1
End Sub

Private Sub Text14_Change()
Cambio = 1
For i = 1 To Len(Text14.Text)
h = Trim(Mid(Text14.Text, i, 1))
If h = "," Then
h = Chr(46)
End If
f = f & h
Next i
Text14.Text = f
End Sub

Private Sub Text32_Change()
Cambio = 1
For i = 1 To Len(Text32.Text)
h = Trim(Mid(Text32.Text, i, 1))
If h = "," Then
h = Chr(46)
End If
f = f & h
Next i
Text32.Text = f
End Sub
Private Sub Text33_Change()
Cambio = 1
End Sub

Private Sub Text34_Change()
Cambio = 1
For i = 1 To Len(Text34.Text)
h = Trim(Mid(Text34.Text, i, 1))
If h = "," Then
h = Chr(46)
End If
f = f & h
Next i
Text34.Text = f
End Sub

Private Sub Text35_Change()
Cambio = 1
For i = 1 To Len(Text35.Text)
h = Trim(Mid(Text35.Text, i, 1))
If h = "," Then
h = Chr(46)
End If
f = f & h
Next i
Text35.Text = f
End Sub

Private Sub TxtAbdomen_Change()
Cambio = 1
End Sub

Private Sub TxtAlmuerzo_Change()
Cambio = 1
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then
    TxtBuscar.Text = ""
End If
If TxtBuscar.Text <> "Busqueda" Then
    TxtBuscar.Text = ""
End If
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

Private Sub TxtCarnes_Change()
Cambio = 1
End Sub

Private Sub TxtCena_Change()
Cambio = 1
End Sub

Private Sub TxtDesayuno_Change()
Cambio = 1
End Sub

Private Sub TxtEmbutidos_Change()
Cambio = 1
End Sub

Private Sub TxtEnlatados_Change()
Cambio = 1
End Sub

Private Sub TxtEstadoGeneral_Change()
Cambio = 1
End Sub

Private Sub TxtGCCA_Change()
Cambio = 1
End Sub

Private Sub TxtGCCD_Change()
Cambio = 1
End Sub

Private Sub TxtGrasas_Change()
Cambio = 1
End Sub

Private Sub TxtHarinas_Change()
Cambio = 1
End Sub

Private Sub TxtIMC_Change()
Cambio = 1
End Sub

Private Sub TxtIMC_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtInteraccionDrogasNutricion_Change()
Cambio = 1
End Sub

Private Sub TxtLengua_Change()
Cambio = 1
End Sub

Private Sub TxtMedicamentos_Change()
Cambio = 1
End Sub

Private Sub TxtMerienda_Change()
Cambio = 1
End Sub

Private Sub TxtOjos_Change()
Cambio = 1
End Sub

Private Sub TxtPesoActual_Change()
Cambio = 1
End Sub

Private Sub TxtPesoActual_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtPesoRefer_Change()
Cambio = 1
End Sub

Private Sub TxtPesoRefer_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtPesoUsual_Change()
Cambio = 1
End Sub

Private Sub TxtPesoUsual_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtPiel_Change()
Cambio = 1
End Sub

Private Sub TxtPorcenPeso_Change()
Cambio = 1
End Sub

Private Sub TxtPorcenPeso_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtRefrescos_Change()
Cambio = 1
End Sub

Private Sub TxtSal_Change()
Cambio = 1
End Sub

Private Sub TxtTalla_Change()
Cambio = 1
End Sub

Private Sub TxtTalla_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

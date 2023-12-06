VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConsultaPsicologicaAdult 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Psicológica Adulto"
   ClientHeight    =   9315
   ClientLeft      =   3150
   ClientTop       =   1005
   ClientWidth     =   13185
   Icon            =   "Psicologia_A.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13185
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Height          =   9255
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   12975
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   68
         Top             =   8400
         Width           =   3735
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
            TabIndex        =   69
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad"
            Top             =   240
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2280
            TabIndex        =   70
            ToolTipText     =   "Buscar"
            Top             =   240
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
            MICON           =   "Psicologia_A.frx":1002
            PICN            =   "Psicologia_A.frx":101E
            PICH            =   "Psicologia_A.frx":1283
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
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3960
         TabIndex        =   57
         Top             =   8400
         Width           =   8895
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   7800
            TabIndex        =   58
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
            MICON           =   "Psicologia_A.frx":1515
            PICN            =   "Psicologia_A.frx":1531
            PICH            =   "Psicologia_A.frx":16FA
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
            TabIndex        =   59
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
            MICON           =   "Psicologia_A.frx":192F
            PICN            =   "Psicologia_A.frx":194B
            PICH            =   "Psicologia_A.frx":1BDA
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
            TabIndex        =   60
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
            MICON           =   "Psicologia_A.frx":201B
            PICN            =   "Psicologia_A.frx":2037
            PICH            =   "Psicologia_A.frx":21C4
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
            Left            =   6600
            TabIndex        =   61
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Deshacer"
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
            MICON           =   "Psicologia_A.frx":23F9
            PICN            =   "Psicologia_A.frx":2415
            PICH            =   "Psicologia_A.frx":26F7
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
            Left            =   5880
            TabIndex        =   62
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
            MICON           =   "Psicologia_A.frx":2948
            PICN            =   "Psicologia_A.frx":2964
            PICH            =   "Psicologia_A.frx":2BFA
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
            Left            =   5280
            TabIndex        =   63
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
            MICON           =   "Psicologia_A.frx":2E59
            PICN            =   "Psicologia_A.frx":2E75
            PICH            =   "Psicologia_A.frx":310A
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
            Left            =   3840
            TabIndex        =   64
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
            MICON           =   "Psicologia_A.frx":3366
            PICN            =   "Psicologia_A.frx":3382
            PICH            =   "Psicologia_A.frx":34A7
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
            Left            =   2400
            TabIndex        =   76
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
            MICON           =   "Psicologia_A.frx":3737
            PICN            =   "Psicologia_A.frx":3753
            PICH            =   "Psicologia_A.frx":38F7
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
         Caption         =   "Aspectos Psicológicos"
         Enabled         =   0   'False
         Height          =   5775
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   12735
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   1200
            Top             =   360
         End
         Begin VB.TextBox Text8 
            Height          =   975
            Left            =   4440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox Text9 
            Height          =   1095
            Left            =   4440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox Text10 
            Height          =   975
            Left            =   8760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox Text11 
            Height          =   1095
            Left            =   8760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   1800
            Width           =   3855
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   960
            TabIndex        =   5
            Top             =   2760
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Detalles"
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox Text16 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3480
            Width           =   4215
         End
         Begin VB.TextBox Text17 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   4680
            Width           =   4215
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Cuando te Menciona la Palabra Cáncer"
            Height          =   2655
            Left            =   4440
            TabIndex        =   35
            Top             =   3000
            Width           =   4215
            Begin VB.TextBox Text15 
               Height          =   2295
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Imaginaria"
               Height          =   255
               Left            =   240
               TabIndex        =   44
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Ideación"
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Actitudes"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Motivación"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Conducta"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Miedo"
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Sensaciones"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Efectividad"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Eficiencia"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   2280
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Impresión Psicológica"
            Height          =   2655
            Left            =   8760
            TabIndex        =   34
            Top             =   3000
            Width           =   3855
            Begin VB.TextBox Text19 
               Height          =   2295
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.TextBox Text20 
            Height          =   375
            Left            =   1800
            TabIndex        =   1
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox Text21 
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox Text22 
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox Text24 
            Height          =   375
            Left            =   960
            TabIndex        =   4
            Top             =   2280
            Width           =   3375
         End
         Begin VB.TextBox Text23 
            Height          =   375
            Left            =   1800
            TabIndex        =   0
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Actitud ante el Diagnóstico:"
            Height          =   195
            Left            =   4440
            TabIndex        =   56
            Top             =   240
            Width           =   1950
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo de la Consulta:"
            Height          =   195
            Left            =   4440
            TabIndex        =   55
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Enfermedad Actual :"
            Height          =   195
            Left            =   8760
            TabIndex        =   54
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sintomatología Psicológica:"
            Height          =   195
            Left            =   8760
            TabIndex        =   53
            Top             =   1560
            Width           =   1965
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Religión:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   2850
            Width           =   615
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fortalezas:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Debilidades:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Profesión u Oficio:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   1410
            Width           =   1290
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel de Instrucción:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   930
            Width           =   1455
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Cónyugue:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1890
            Width           =   1620
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   2370
            Width           =   465
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hijos:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   450
            Width           =   390
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   12735
         Begin VB.ComboBox CboSexo 
            Height          =   315
            Left            =   6120
            TabIndex        =   67
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   5760
            TabIndex        =   21
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   3960
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1320
            TabIndex        =   19
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1320
            TabIndex        =   18
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1320
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DtpFechaActual 
            Height          =   315
            Left            =   4080
            TabIndex        =   22
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51314689
            CurrentDate     =   39813
         End
         Begin MSComCtl2.DTPicker DtpFechaNac 
            Height          =   375
            Left            =   6120
            TabIndex        =   65
            Top             =   1320
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51314689
            CurrentDate     =   39813
         End
         Begin MSComCtl2.DTPicker DtpFechaRegistro 
            Height          =   315
            Left            =   1560
            TabIndex        =   66
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51314689
            CurrentDate     =   39813
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   6960
            Top             =   3000
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Historia Clínica Integral"
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
         Begin ChamaleonButton.ChameleonBtn BtnLlamar 
            Height          =   375
            Left            =   8160
            TabIndex        =   71
            ToolTipText     =   "Llamar"
            Top             =   840
            Width           =   2295
            _ExtentX        =   4048
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
            MICON           =   "Psicologia_A.frx":3A96
            PICN            =   "Psicologia_A.frx":3AB2
            PICH            =   "Psicologia_A.frx":3D4E
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
            Left            =   8160
            TabIndex        =   72
            ToolTipText     =   "Lista de Espera"
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
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
            MICON           =   "Psicologia_A.frx":3F83
            PICN            =   "Psicologia_A.frx":3F9F
            PICH            =   "Psicologia_A.frx":4228
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
            Left            =   8160
            TabIndex        =   75
            ToolTipText     =   "Desocupar al Paciente Atendido"
            Top             =   1320
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
            MICON           =   "Psicologia_A.frx":44C0
            PICN            =   "Psicologia_A.frx":44DC
            PICH            =   "Psicologia_A.frx":4680
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
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   10800
            TabIndex        =   74
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Historia:"
            Height          =   195
            Left            =   4800
            TabIndex        =   73
            Top             =   480
            Width           =   870
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupación:"
            Height          =   195
            Left            =   4800
            TabIndex        =   32
            Top             =   930
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Sexo:"
            Height          =   195
            Left            =   5640
            TabIndex        =   31
            Top             =   1860
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Edad:"
            Height          =   195
            Left            =   3480
            TabIndex        =   30
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Fecha de Nac.:"
            Height          =   195
            Left            =   4920
            TabIndex        =   29
            Top             =   1410
            Width           =   1110
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "A&pellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nombre(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de &Registro:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1860
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "No. &Cédula:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   450
            Width           =   840
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   375
            Left            =   5760
            TabIndex        =   24
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1575
            Left            =   10800
            Picture         =   "Psicologia_A.frx":481F
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Actual:"
            Height          =   195
            Left            =   3000
            TabIndex        =   23
            Top             =   1860
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "FrmConsultaPsicologicaAdult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPaciente As New ADODB.Recordset 'tabla Registro pacientes
Dim RsPsicologia As New ADODB.Recordset 'tabla Registro pacientes
Dim RsTemp As New ADODB.Recordset
Dim SQL As String
Dim Cambio
Dim RegNew
Dim NuevoId As String
Dim IdPsic As String
Dim IdPacP
Dim IdLIdPac As String

Private Sub BtnAgregar_Click()
On Error Resume Next
IO = 1

If Trim(IdPacP) = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
BtnGuardarActualizar.Enabled = True
Frame2.BackColor = &HE0E0E0
Frame3.BackColor = &HE0E0E0
Frame4.BackColor = &HE0E0E0
Frame2.Enabled = True


CSql = "Select * From Psicologia_A Where IdPaciente='" & IdPacP & "'"
Set RsPsicologia = CrearRS(CSql)

If RsPsicologia.RecordCount > 0 Then
    Text23.SetFocus
Else
    Call Blanqueo
    Text23.SetFocus
End If


RegNew = 1

End Sub

Private Sub BtnAnterior_Click()
On Error Resume Next
If Trim(IdPacP) = "" Then Exit Sub
Call Blanqueo
If RsPaciente.RecordCount <> 0 Then
    RsPaciente.MovePrevious
    If RsPaciente.BOF Then RsPaciente.MoveLast
    Call Carga_De_Datos
    
    CSql = "select * from psicologia_a where IdPaciente = " & IdPacP & " and Activo=1"
    Set RsPsicologia = CrearRS(CSql)
    Call carga_datos_psico
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros"
    IdPacP = ""
End If
End Sub

Public Sub BtnBuscar_Click()
On Error Resume Next
If TxtBuscar.Text = "Busqueda" Or TxtBuscar.Text = "" Then
    CSql = "select * from Paciente order by idpaciente" 'where (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) > 18)"
Else
    CSql = "select * from Paciente where Historia='" & TxtBuscar.Text & "' or cedulaP = " & Val(TxtBuscar.Text) & " or nombreP like '%" & TxtBuscar.Text & "%' or ApellidoP like '%" & TxtBuscar.Text & "%' " 'AND (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) > 18)"
End If
Call Blanqueo
Set RsPaciente = CrearRS(CSql)

If RsPaciente.RecordCount = 0 Then
    MsgBox "No Existe el registro", vbInformation + vbOKOnly, "No hay datos"
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Text14.Text = ""
    Label33.Caption = ""
    DtpFechaNac.Value = Now
    DtpFechaRegistro.Value = Now
    DtpFechaActual.Value = Now
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    CboSexo.ListIndex = -1
    NoReg = "Registro 0 / 0"
    IdPacP = ""
    Exit Sub
End If

Call Carga_De_Datos

CSql = "Select * from psicologia_a where IdPaciente = " & IdPacP & " And IdLIdPac='" & IdLIdPac & "' And Activo=1"
Set RsPsicologia = CrearRS(CSql)

If RsPsicologia.RecordCount = 0 Then
    MsgBox "Este Paciente no posee datos en la tabla de Psicología", vbInformation + vbOKOnly, "No hay datos"
    RegNew = 1
Else
    Call carga_datos_psico
End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
Call Blanqueo
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
BtnImprimir.Enabled = True
BtnEliminar.Enabled = True
BtnAgregar.Enabled = True
'Frame2.Enabled = False
Frame2.BackColor = &HEAEFEF
Frame3.BackColor = &HEAEFEF
Frame4.BackColor = &HEAEFEF

Cambio = 0
If IdPacP = "" Then Exit Sub
Call Carga_De_Datos
    
CSql = "select * from psicologia_a where IdPaciente = " & IdPacP & " and Activo=1"
Set RsPsicologia = CrearRS(CSql)
Call carga_datos_psico
End Sub

Private Sub BtnDesocuparAlPacienteAtendido_Click()
On Error Resume Next
Dim bdlista88 As New ADODB.Recordset
CSql = "Delete from ubi_paciente where modul = " & ModulO
Set bdlista88 = CrearRS(CSql)
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next
If IdPacP = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
If IdPsic = "" Then MsgBox "No existen registros para borrar!", vbExclamation + vbOKOnly, "Informacion": Exit Sub

resp = MsgBox("Se procedera ha eliminar el registro de Psicologia del paciente, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

CSql = "Update psicologia_A set Activo=0 Where IdPsicologia=" & IdPsic & " And IdL='" & IdLIdInf & "'"
Set RsTemp = CrearRS(CSql)

Call Enviar_Bitacora(IdUser, "PSICOLOGIA-ADULTOS", "BORRAR", "se elimino en la tabla psicologia_a el registro de Id=" & IdPsic)

EnviarRegPendiente NuevoId, IdLIdInf

MsgBox "El registro fue eliminado.", vbInformation + vbOKOnly, "Operacion Exitosa"

BtnDesHacer_Click

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next

If IdPacP = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
If Cambio = 0 Then MsgBox "no se han realizado cambios!", vbExclamation + vbOKOnly, "Informacion": Exit Sub

If IdPacP = 0 Then Exit Sub
       If Text8.Text = "" Then
        f = "Actitud_Diag"
        GoTo noguardA
        End If
        
        If Text9.Text = "" Then
        f = "Motivo_consul"
        GoTo noguardA
        
        End If
        
        If Text10.Text = "" Then
        f = "Enfermedad"
        GoTo noguardA
        
        End If
                        
        If Text11.Text = "" Then
        f = "Sinto_Psigologico"
        GoTo noguardA
       
        End If
        
        If Text13.Text = "" Then
        f = "Religion"
        GoTo noguardA
       
        End If
        
        If Text15.Text = "" Then
        f = "Palabra_cancer"
        GoTo noguardA
       
        End If
        
        If Text16.Text = "" Then
        f = "Fortalezas"
        GoTo noguardA
       
        End If
        
        If Text17.Text = "" Then
        f = "Debilidades"
        GoTo noguardA
      
        End If
                  
        If Text19.Text = "" Then
        f = "Observaciones"
        GoTo noguardA
       
        End If
        
        If Text20.Text = "" Then
        f = "Nivel_instru"
        GoTo noguardA
        
        End If
        
                
        If Text21.Text = "" Then
        f = "Profesion"
        GoTo noguardA
      
        End If
        
        If Text22.Text = "" Then
        f = "Conyugue"
        GoTo noguardA
        
        End If
        
        If Text23.Text = "" Then
        f = "Hijos"
        GoTo noguardA

        End If
        
        If Text24.Text = "" Then
        f = "Email"
        GoTo noguardA
 
        End If


resp = MsgBox("Se Guardaran los cambios realizados, Desea Continuar?", vbExclamation + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

CSql = "SELECT MAX(idpsicologia)+1 as NuevoId FROM psicologia_a "
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = RsTemp.Fields("NuevoId").Value
Else
    NuevoId = "0"
End If

Select Case RegNew
    
    Case Is = 0
          ' Actualizar registro
          ' hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos
 
            CSql = "Select * From Psicologia_A Where IdPaciente='" & IdPacP & "' And IdLIdPac='" & IdLIdPac & "' And IdPsicologia='" & IdPsic & "' And IdL='" & IdLIdInf & "'"
            Set RsTemp = CrearRS(CSql)
            
            NuevoId = IdPsic
            RsTemp.Fields("observacion").Value = Observa
            RsTemp.Fields("detalles").Value = Detalle
            RsTemp.Fields("idusuario").Value = IdUser
            RsTemp.Fields("ACTITUD").Value = Text8.Text
            RsTemp.Fields("MOTIVO").Value = Text9.Text
            RsTemp.Fields("ENFERMEDAD").Value = Text10.Text
            RsTemp.Fields("SINTOMA").Value = Text11.Text
            RsTemp.Fields("RELIGION").Value = Text13.Text
            RsTemp.Fields("CANCER").Value = Text15.Text
            RsTemp.Fields("FORTALEZAS").Value = Text16.Text
            RsTemp.Fields("DEBILIDADES").Value = Text17.Text
            RsTemp.Fields("OBSERVACIONES").Value = Text19.Text
            RsTemp.Fields("NIVEL").Value = Text20.Text
            RsTemp.Fields("PROFESION").Value = Text21.Text
            RsTemp.Fields("CONYUGUE").Value = Text22.Text
            RsTemp.Fields("HIJOS").Value = Text23.Text
            RsTemp.Fields("EMAIL").Value = Text24.Text
            RsTemp.Fields("ACTIVO").Value = 1
            RsTemp.Fields("Fecha").Value = Format(DtpFechaActual.Value, "dd/mm/yyyy")
            RsTemp.Update
            
            MsgBox "Registro Actualizado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
            
            EnviarRegPendiente NuevoId, IdLIdInf
            
    Case Is = 1 'Agrega registro
            
            CSql = "Select * From Psicologia_A"
            Set RsTemp = CrearRS(CSql)
            
            IdLIdInf = NuevoIdL
            
            RsTemp.AddNew
            
            RsTemp.Fields("IdPsicologia").Value = NuevoId
            RsTemp.Fields("IdL").Value = IdLIdInf
            RsTemp.Fields("IdPaciente").Value = IdPacP
            RsTemp.Fields("IdLIdPac").Value = IdLIdPac
            
            RsTemp.Fields("observacion").Value = Observa
            RsTemp.Fields("detalles").Value = Detalle
            RsTemp.Fields("idusuario").Value = IdUser
            RsTemp.Fields("ACTITUD").Value = Text8.Text
            RsTemp.Fields("MOTIVO").Value = Text9.Text
            RsTemp.Fields("ENFERMEDAD").Value = Text10.Text
            RsTemp.Fields("SINTOMA").Value = Text11.Text
            RsTemp.Fields("RELIGION").Value = Text13.Text
            RsTemp.Fields("CANCER").Value = Text15.Text
            RsTemp.Fields("FORTALEZAS").Value = Text16.Text
            RsTemp.Fields("DEBILIDADES").Value = Text17.Text
            RsTemp.Fields("OBSERVACIONES").Value = Text19.Text
            RsTemp.Fields("NIVEL").Value = Text20.Text
            RsTemp.Fields("PROFESION").Value = Text21.Text
            RsTemp.Fields("CONYUGUE").Value = Text22.Text
            RsTemp.Fields("HIJOS").Value = Text23.Text
            RsTemp.Fields("EMAIL").Value = Text24.Text
            RsTemp.Fields("ACTIVO").Value = 1
            RsTemp.Fields("Fecha").Value = Format(DtpFechaActual.Value, "dd/mm/yyyy")
            RsTemp.Update
            
            MsgBox "Registro Agregado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
            
            EnviarRegPendiente NuevoId, IdLIdInf
             
End Select

If Cambio = 1 And RegNew = 0 Then
    
    If Reg_Actual(0) <> Text8.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo actitud de (" & Reg_Actual(0) & ") a (" & Text8.Text & ")")
    If Reg_Actual(1) <> Text9.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo motivo de (" & Reg_Actual(1) & ") a (" & Text9.Text & ")")
    If Reg_Actual(2) <> Text10.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo enfermedad de (" & Reg_Actual(2) & ") a (" & Text10.Text & ")")
    If Reg_Actual(3) <> Text11.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo sintoma de (" & Reg_Actual(3) & ") a (" & Text11.Text & ")")
    If Reg_Actual(4) <> Text13.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Religion de (" & Reg_Actual(4) & ") a (" & Text13.Text & ")")
    If Reg_Actual(5) <> Text23.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo hijos de (" & Reg_Actual(5) & ") a (" & Text23.Text & ")")
    If Reg_Actual(6) <> Text15.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Cancer de (" & Reg_Actual(6) & ") a (" & Text15.Text & ")")
    If Reg_Actual(7) <> Text16.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Fortalezas de (" & Reg_Actual(7) & ") a (" & Text16.Text & ")")
    If Reg_Actual(8) <> Text17.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Debilidades de (" & Reg_Actual(8) & ") a (" & Text17.Text & ")")
    If Reg_Actual(9) <> Text19.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Observaciones de (" & Reg_Actual(9) & ") a (" & Text19.Text & ")")
    If Reg_Actual(10) <> Text20.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Nivel de (" & Reg_Actual(10) & ") a (" & Text20.Text & ")")
    If Reg_Actual(11) <> Text21.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Profesion de (" & Reg_Actual(11) & ") a (" & Text21.Text & ")")
    If Reg_Actual(12) <> Text22.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Conyugue de (" & Reg_Actual(12) & ") a (" & Text22.Text & ")")
    If Reg_Actual(13) <> Text24.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Email de (" & Reg_Actual(13) & ") a (" & Text24.Text & ")")
    If Reg_Actual(14) <> Observa Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo Observacion de (" & Reg_Actual(14) & ") a (" & Observa & ")")
    If Reg_Actual(15) <> Detalle Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo detalles de (" & Reg_Actual(15) & ") a (" & Detalle & ")")
ElseIf Cambio = 1 And RegNew = 1 Then
    Call Enviar_Bitacora(IdUser, "PSICOLOGIA ADULTOS", "INGRESAR", "se Ingreso en la tabla psicologia_a el nuevo registro de Id=" & NuevoId)
End If



Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor Web"

BtnDesHacer_Click
Frame2.BackColor = &HEAEFEF
Frame3.BackColor = &HEAEFEF
Frame4.BackColor = &HEAEFEF
'Frame2.Enabled = False
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly, "Error al Guardar"

    Exit Sub
Cambio = 0

End Sub

Sub EnviarRegPendiente(ByVal NuevoId2 As Integer, ByVal IdLIdInf2 As String)
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

CSql = "SELECT * FROM Psicologia_A WHERE IdPsicologia=" & NuevoId2 & " And IdL='" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Psicologia_A (["
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
RsRegPendiente.Fields("Modulo").Value = "Psicologia Adulto"
RsRegPendiente.Fields("Tabla").Value = "Psicologia_A"
RsRegPendiente.Fields("Condicional").Value = "IdPsicologia=" & NuevoId2 & " And IdL='" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"
End Sub

Sub EnviarAlHosting()
On Error GoTo salir
ConectarHosting

If Cambio = 1 And RegNew = 0 Then
    
    If Reg_Actual(0) <> Text8.Text Then
        CSql = "Update psicologia_A set Actitud='" & Trim(Text8.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(1) <> Text9.Text Then
        CSql = "Update psicologia_A set Motivo='" & Trim(Text9.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(2) <> Text10.Text Then
        CSql = "Update psicologia_A set Enfermedad='" & Trim(Text10.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(3) <> Text11.Text Then
        CSql = "Update psicologia_A set Sintoma='" & Trim(Text11.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(4) <> Text13.Text Then
        CSql = "Update psicologia_A set Religion='" & Trim(Text13.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(5) <> Text23.Text Then
        CSql = "Update psicologia_A set Hijos=" & Trim(Text23.Text) & " Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(6) <> Text15.Text Then
        CSql = "Update psicologia_A set Cancer='" & Trim(Text15.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(7) <> Text16.Text Then
        CSql = "Update psicologia_A set Fortalezas='" & Trim(Text16.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(8) <> Text17.Text Then
        CSql = "Update psicologia_A set Debilidades='" & Trim(Text17.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(9) <> Text19.Text Then
        CSql = "Update psicologia_A set Observaciones='" & Trim(Text19.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(10) <> Text20.Text Then
        CSql = "Update psicologia_A set Nivel='" & Trim(Text20.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(11) <> Text21.Text Then
        CSql = "Update psicologia_A set Profesion='" & Trim(Text21.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(12) <> Text22.Text Then
        CSql = "Update psicologia_A set Conyugue='" & Trim(Text22.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(13) <> Text24.Text Then
        CSql = "Update psicologia_A set Email='" & Trim(Text24.Text) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(14) <> Observa Then
        CSql = "Update psicologia_A set Observacion='" & Trim(Observa) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If
    
    If Reg_Actual(15) <> Detalle Then
        CSql = "Update psicologia_A set Detalles='" & Trim(Detalle) & "' Where Activo=1 AND IdPaciente=" & IdPacP
        Set RsWeb = CrearRsWeb(CSql)
    End If

End If

If Cambio = 1 And RegNew = 1 Then
            CSql = "Update psicologia_A set Activo=0 Where Activo=1 AND idpaciente=" & IdPacP
            Set RsWeb = CrearRsWeb(CSql)
            
            CSql = "Insert Into Psicologia_A(idpsicologia,idpaciente,observacion,detalles,idusuario,ACTITUD,MOTIVO," & _
                "ENFERMEDAD,SINTOMA,RELIGION,CANCER,FORTALEZAS,DEBILIDADES,OBSERVACIONES,NIVEL,PROFESION,CONYUGUE," & _
                "HIJOS,EMAIL,ACTIVO,Fecha) VALUES(" & NuevoId & "," & IdPacP & ",'" & Observa & "','" & Detalle & "'," & IdUser & ",'" & Text8.Text & _
                "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text13.Text & "','" & _
                Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text19.Text & "','" & Text20.Text & _
                "','" & Text21.Text & "','" & Text22.Text & "','" & Text23.Text & "','" & Text24.Text & "',1,'" & Format(DtpFechaActual.Value, "mm/dd/yyyy") & "')"

            Set RsWeb = CrearRsWeb(CSql)
End If



    
CSql = "Select * From Psicologia_A Where idpsicologia='" & NuevoId & "'"
Set RsWeb = CrearRsWeb(CSql)

If RsWeb.RecordCount > 0 Then
    Msg = "Actualización Completada en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
Else
    Msg = "No se realizo la Actualización en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
End If



salir:
If WebCnn.State = 0 Then
    EnviarAlHosting
Else
    GoTo f
End If

f:
If WebCnn.State = 1 Then
    WebCnn.Close
Else
    If Err.Number <> 0 Then MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub BtnImprimir_Click()
On Error Resume Next

If Text1.Text <> "" And IdPacP <> "" Then
     With CrystalReport1
        .ReportFileName = RutaInformes & "\HistoriaClinicaIntegralN.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0

        .SelectionFormula = "{Historia_Clinica.Idpaciente} = " & IdPacP & " And {Historia_Clinica.IdL} = '" & IdLIdInf & "'"
        .ReportTitle = "Reporte Historia Medica No. " & Label12.Caption
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
Else
    MsgBox "Tiene que seleccionar a una paciente", vbOKOnly + vbCritical, "Mensaje de Error"
End If

End Sub

Private Sub BtnListaEspera_Click()
On Error Resume Next
ModulO = 1
Consultaa = "A"
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
On Error Resume Next
If Trim(IdPacP) = "" Then Exit Sub
Call Blanqueo
If RsPaciente.RecordCount <> 0 Then
    RsPaciente.MoveNext
    If RsPaciente.EOF Then RsPaciente.MoveFirst
    Call Carga_De_Datos
    
    CSql = "select * from psicologia_a where IdPaciente = " & IdPacP & " and Activo=1"
    Set RsPsicologia = CrearRS(CSql)
    Call carga_datos_psico
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros"
    IdPacP = ""
End If
End Sub

Sub Carga_De_Datos()

IdPacP = ""
IdLIdPac = IdLDefault
If RsPaciente.RecordCount = 0 Then
    IdPacP = ""
    BtnAgregar.Enabled = False
    
    NoReg = "Registro 0 / 0"
    Exit Sub
Else
    If Trim(RsPaciente.Fields("cedulaP")) <> "" Then Text1.Text = RsPaciente.Fields("cedulap")
    If Trim(RsPaciente.Fields("Fecha_regP")) <> "" Then DtpFechaRegistro = RsPaciente.Fields("Fecha_regp")
    If Trim(RsPaciente.Fields("NombreP")) <> "" Then Text3.Text = RsPaciente.Fields("Nombrep")
    If Trim(RsPaciente.Fields("ApellidoP")) <> "" Then Text4.Text = RsPaciente.Fields("Apellidop")
    If Trim(RsPaciente.Fields("Fecha_nacimientoP")) <> "" Then DtpFechaNac = RsPaciente.Fields("Fecha_nacimientop")
    If Trim(RsPaciente.Fields("EdadP")) <> "" Then Text6.Text = RsPaciente.Fields("Edadp")
    If Trim(RsPaciente.Fields("Historia")) <> "" Then Label33.Caption = RsPaciente.Fields("Historia")
    IdPacP = RsPaciente.Fields("idpaciente")
    IdLIdPac = RsPaciente.Fields("IdL")
    Me.Caption = "Consulta Psicológica Adulto - Paciente: " & IdPacP
    If RsPaciente.Fields("foto") <> "" Then
        If Len(Dir(Foto & "\" & RsPaciente.Fields("foto"))) > 0 Then
            Image2.Picture = LoadPicture(Foto & "\" & RsPaciente.Fields("foto"))
        Else
            Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        End If
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
    If RsPaciente.Fields("sexop") = 0 Then CboSexo.Text = "Masculino" Else CboSexo.Text = "Femenino"
    Text14.Text = RsPaciente.Fields("Ocupacion")
    BtnAgregar.Enabled = True
    'BtnGuardarActualizar.Enabled = False
    NoReg = "Registro " & RsPaciente.AbsolutePosition & " / " & RsPaciente.RecordCount
    'Frame2.Enabled = False
End If

End Sub
Sub carga_datos_psico()

IdPsic = ""
IdLIdInf = IdLDefault

If RsPsicologia.RecordCount <> 0 Then
    
    IdPsic = RsPsicologia.Fields("IdPsicologia").Value
    IdLIdInf = RsPsicologia.Fields("IdL").Value
    
    If Trim(RsPsicologia.Fields("actitud")) <> "" Then Text8.Text = RsPsicologia.Fields("actitud") Else Text8.Text = ""
    If Trim(RsPsicologia.Fields("motivo")) <> "" Then Text9.Text = RsPsicologia.Fields("motivo") Else Text9.Text = ""
    If Trim(RsPsicologia.Fields("enfermedad")) <> "" Then Text10.Text = RsPsicologia.Fields("enfermedad") Else Text10.Text = ""
    If Trim(RsPsicologia.Fields("Sintoma")) <> "" Then Text11.Text = RsPsicologia.Fields("Sintoma") Else Text11.Text = ""
    If Trim(RsPsicologia.Fields("Religion")) <> "" Then Text13.Text = RsPsicologia.Fields("Religion") Else Text13.Text = ""
    If Trim(RsPsicologia.Fields("Hijos")) <> "" Then Text23.Text = RsPsicologia.Fields("Hijos") Else Text23.Text = ""
    If Trim(RsPsicologia.Fields("Cancer")) <> "" Then Text15.Text = RsPsicologia.Fields("Cancer") Else Text15.Text = ""
    If Trim(RsPsicologia.Fields("Fortalezas")) <> "" Then Text16.Text = RsPsicologia.Fields("Fortalezas") Else Text16.Text = ""
    If Trim(RsPsicologia.Fields("Debilidades")) <> "" Then Text17.Text = RsPsicologia.Fields("Debilidades") Else Text17.Text = ""
    If Trim(RsPsicologia.Fields("Observaciones")) <> "" Then Text19.Text = RsPsicologia.Fields("Observaciones") Else Text19.Text = ""
    If Trim(RsPsicologia.Fields("Nivel")) <> "" Then Text20.Text = RsPsicologia.Fields("Nivel") Else Text20.Text = ""
    If Trim(RsPsicologia.Fields("Profesion")) <> "" Then Text21.Text = RsPsicologia.Fields("Profesion") Else Text21.Text = ""
    If Trim(RsPsicologia.Fields("Conyugue")) <> "" Then Text22.Text = RsPsicologia.Fields("Conyugue") Else Text22.Text = ""
    If Trim(RsPsicologia.Fields("Email")) <> "" Then Text24.Text = RsPsicologia.Fields("Email") Else Text24.Text = ""
    If Trim(RsPsicologia.Fields("observacion")) <> "" Then Observa = RsPsicologia.Fields("observacion") Else Observa = ""
    If Trim(RsPsicologia.Fields("detalles")) <> "" Then Detalle = RsPsicologia.Fields("detalles") Else Detalle = ""
    
    If Trim(RsPsicologia.Fields("actitud")) <> "" Then Reg_Actual(0) = RsPsicologia.Fields("actitud").Value Else Reg_Actual(0) = ""
    If Trim(RsPsicologia.Fields("motivo")) <> "" Then Reg_Actual(1) = RsPsicologia.Fields("motivo") Else Reg_Actual(1) = ""
    If Trim(RsPsicologia.Fields("enfermedad")) <> "" Then Reg_Actual(2) = RsPsicologia.Fields("enfermedad") Else Reg_Actual(2) = ""
    If Trim(RsPsicologia.Fields("Sintoma")) <> "" Then Reg_Actual(3) = RsPsicologia.Fields("Sintoma") Else Reg_Actual(3) = ""
    If Trim(RsPsicologia.Fields("Religion")) <> "" Then Reg_Actual(4) = RsPsicologia.Fields("Religion") Else Reg_Actual(4) = ""
    If Trim(RsPsicologia.Fields("Hijos")) <> "" Then Reg_Actual(5) = RsPsicologia.Fields("Hijos") Else Reg_Actual(5) = ""
    If Trim(RsPsicologia.Fields("Cancer")) <> "" Then Reg_Actual(6) = RsPsicologia.Fields("Cancer") Else Reg_Actual(6) = ""
    If Trim(RsPsicologia.Fields("Fortalezas")) <> "" Then Reg_Actual(7) = RsPsicologia.Fields("Fortalezas") Else Reg_Actual(7) = ""
    If Trim(RsPsicologia.Fields("Debilidades")) <> "" Then Reg_Actual(8) = RsPsicologia.Fields("Debilidades") Else Reg_Actual(8) = ""
    If Trim(RsPsicologia.Fields("Observaciones")) <> "" Then Reg_Actual(9) = RsPsicologia.Fields("Observaciones") Else Reg_Actual(9) = ""
    If Trim(RsPsicologia.Fields("Nivel")) <> "" Then Reg_Actual(10) = RsPsicologia.Fields("Nivel") Else Reg_Actual(10) = ""
    If Trim(RsPsicologia.Fields("Profesion")) <> "" Then Reg_Actual(11) = RsPsicologia.Fields("Profesion") Else Reg_Actual(11) = ""
    If Trim(RsPsicologia.Fields("Conyugue")) <> "" Then Reg_Actual(12) = RsPsicologia.Fields("Conyugue") Else Reg_Actual(12) = ""
    If Trim(RsPsicologia.Fields("Email")) <> "" Then Reg_Actual(13) = RsPsicologia.Fields("Email") Else Reg_Actual(13) = ""
    If Trim(RsPsicologia.Fields("observacion")) <> "" Then Reg_Actual(14) = RsPsicologia.Fields("observacion") Else Reg_Actual(14) = ""
    If Trim(RsPsicologia.Fields("detalles")) <> "" Then Reg_Actual(15) = RsPsicologia.Fields("detalles") Else Reg_Actual(15) = ""
    
    Cambio = 0: RegNew = 0
    BtnGuardarActualizar.Enabled = True
    BtnImprimir.Enabled = True
    BtnEliminar.Enabled = True
    'Frame2.Enabled = False
    'Frame2.BackColor = &HE0E0E0
Else
    IdPsic = ""
    Cambio = 0: RegNew = 1
    BtnGuardarActualizar.Enabled = False
    BtnImprimir.Enabled = False
    BtnEliminar.Enabled = False
    'Frame2.Enabled = False
   ' Frame2.BackColor = &HE0E0E0
    For i = 0 To 20
        Reg_Actual(i) = ""
    Next i
End If
End Sub

Private Sub CboSexo_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text14.SetFocus
        Case vbKeyLeft
            Text6.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyRight
            Text14.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaActual_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyLeft
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

 

Private Sub DtpFechaNac_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyRight
            Text6.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaActual.SetFocus
        Case vbKeyLeft
            Text1.SetFocus
        Case vbKeyRight
            DtpFechaActual.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cambio = 1 And IdPacP <> "" Then
Msg = "HA REALIZADO CAMBIOS EN ESTE REGISTRO!!!!!!!" & Chr(13) & "DESEA GUARDAR ESTOS CAMBIOS????"
f = MsgBox(Msg, vbInformation + vbYesNo, "GUARDAR CAMBIOS??")
If f = vbYes Then Call BtnGuardarActualizar_Click
End If
 ' If RsCargarPacientes.State = adStateOpen Then RsCargarPacientes.Close

Exit Sub
WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyRight
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text10.SelStart + 1) > Len(Text10.Text)) Or (Shift = 0 And Text10.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyLeft
            Text8.SetFocus
        Case vbKeyUp
            BtnAyuda.SetFocus
        Case vbKeyDown
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text11.SelStart + 1) > Len(Text11.Text)) Or (Shift = 0 And Text11.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text15.SetFocus
        Case vbKeyLeft
            Text9.SetFocus
        Case vbKeyUp
            Text10.SetFocus
        Case vbKeyDown
            Text19.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text13.SelStart + 1) > Len(Text13.Text)) Or (Shift = 0 And Text13.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text16.SetFocus
        Case vbKeyUp
            Text24.SetFocus
        Case vbKeyRight
            Command3.SetFocus
        Case vbKeyDown
            Text16.SetFocus
    End Select
End If
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnListaEspera.SetFocus
        Case vbKeyLeft
            CboSexo.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyDown
            BtnAyuda.SetFocus
    End Select
End If
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text15.SelStart + 1) > Len(Text15.Text)) Or (Shift = 0 And Text15.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text19.SetFocus
        Case vbKeyUp
            Text9.SetFocus
        Case vbKeyRight
            Text19.SetFocus
        Case vbKeyLeft
            Text16.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text16.SelStart + 1) > Len(Text16.Text)) Or (Shift = 0 And Text16.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text17.SetFocus
        Case vbKeyUp
            Text13.SetFocus
        Case vbKeyRight
            Text15.SetFocus
        Case vbKeyDown
            Text17.SetFocus
    End Select
End If
End Sub

Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text17.SelStart + 1) > Len(Text17.Text)) Or (Shift = 0 And Text17.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text15.SetFocus
        Case vbKeyUp
            Text16.SetFocus
        Case vbKeyRight
            Text15.SetFocus
        Case vbKeyDown
            OptApellido.SetFocus
    End Select
End If
End Sub

Private Sub Text19_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text19.SelStart + 1) > Len(Text19.Text)) Or (Shift = 0 And Text19.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyLeft
            Text15.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text20_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text20.SelStart + 1) > Len(Text20.Text)) Or (Shift = 0 And Text20.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text21.SetFocus
        Case vbKeyUp
            Text23.SetFocus
        Case vbKeyRight
            Text8.SetFocus
        Case vbKeyDown
            Text21.SetFocus
    End Select
End If
End Sub

Private Sub Text21_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text21.SelStart + 1) > Len(Text21.Text)) Or (Shift = 0 And Text21.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text22.SetFocus
        Case vbKeyUp
            Text20.SetFocus
        Case vbKeyRight
            Text8.SetFocus
        Case vbKeyDown
            Text22.SetFocus
    End Select
End If
End Sub

Private Sub Text22_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text22.SelStart + 1) > Len(Text22.Text)) Or (Shift = 0 And Text22.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text24.SetFocus
        Case vbKeyUp
            Text21.SetFocus
        Case vbKeyRight
            Text9.SetFocus
        Case vbKeyDown
            Text24.SetFocus
    End Select
End If
End Sub

Private Sub Text23_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text23.SelStart + 1) > Len(Text23.Text)) Or (Shift = 0 And Text23.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text20.SetFocus
        Case vbKeyUp
            BtnListaEspera.SetFocus
        Case vbKeyRight
            Text8.SetFocus
        Case vbKeyDown
            Text20.SetFocus
    End Select
End If
End Sub

Private Sub Text24_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text24.SelStart + 1) > Len(Text24.Text)) Or (Shift = 0 And Text24.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case vbKeyUp
            Text22.SetFocus
        Case vbKeyRight
            Text9.SetFocus
        Case vbKeyDown
            Text13.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaNac.SetFocus
        Case vbKeyLeft
            Text4.SetFocus
        Case vbKeyUp
            DtpFechaActual.SetFocus
        Case vbKeyDown
            Text14.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Text3.SetFocus
        Case vbKeyDown
            DtpFechaNac.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboSexo.SetFocus
        Case vbKeyLeft
            DtpFechaNac.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyRight
            CboSexo.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text8.SelStart + 1) > Len(Text8.Text)) Or (Shift = 0 And Text8.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case vbKeyLeft
            Text23.SetFocus
        Case vbKeyUp
            BtnListaEspera.SetFocus
        Case vbKeyRight
            Text10.SetFocus
        Case vbKeyDown
            Text9.SetFocus
    End Select
End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text9.SelStart + 1) > Len(Text9.Text)) Or (Shift = 0 And Text9.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case vbKeyLeft
            Text22.SetFocus
        Case vbKeyUp
            Text8.SetFocus
        Case vbKeyRight
            Text11.SetFocus
        Case vbKeyDown
            Text15.SetFocus
    End Select
End If
End Sub

Private Sub Timer1_Timer()
'If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
'If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub
Private Sub Command3_Click()
On Error Resume Next
FrmDetalles.Show
End Sub

Private Sub Form_Load()

Centrar Me
ModulO = 1

For i = 0 To 20
    Reg_Actual(i) = ""
Next
'Frame2.Enabled = False
SQL = "select * from Paciente where (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) > 18) order by IdPaciente"
Set RsPaciente = CrearRS(SQL)

'Call Carga_De_Datos

'CSql = "select * from psicologia_a where IdPaciente = " & IdPacP & " and Activo=1 order by idpaciente"
'Set RsPsicologia = CrearRS(CSql)

DtpFechaActual.Value = Now()
SQL = ""
IdPacP = ""
IdPsic = ""
Cambio = 0
RegNew = 0
'Call carga_datos_psico

End Sub
      
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text10_Change()
Cambio = 1
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Len(Trim(Text10.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text11_Change()
Cambio = 1
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If Len(Trim(Text11.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text13_Change()
Cambio = 1
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If Len(Trim(Text13.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text15_Change()
Cambio = 1
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If Len(Trim(Text15.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text16_Change()
Cambio = 1
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If Len(Trim(Text16.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub


Private Sub Text17_Change()
Cambio = 1
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If Len(Trim(Text17.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text19_Change()
Cambio = 1
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If Len(Trim(Text19.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text20_Change()
Cambio = 1
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)

If Len(Trim(Text20.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text21_Change()
Cambio = 1
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If Len(Trim(Text21.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text22_Change()
Cambio = 1
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If Len(Trim(Text22.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text23_Change()
Cambio = 1
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 48 To 57 ' permite el ingreso de numeros
Case Is = 13 ' permite presionar el ENTER
Call BtnBuscar_Click

Case Is = 8 ' Permite Borrar de retroceso
Case Else ' Inhibe todas las demas teclas
KeyAscii = 0

End Select

End Sub

Private Sub Text24_Change()
Cambio = 1
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If Len(Trim(Text24.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text7_Click()
If Text7.Text = "Busqueda" Then Text7.Text = ""
End Sub

Private Sub Text7_GotFocus()
If Text7.Text = "Busqueda" Then Text7.Text = ""
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call BtnBuscar_Click
End Sub

Private Sub Text7_LostFocus()
If Trim(Text7.Text) = "" Then Text7.Text = "Busqueda"

End Sub

Private Sub Text8_Change()
Cambio = 1
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Len(Trim(Text8.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text9_Change()
Cambio = 1
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If Len(Trim(Text9.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else
End If
End Sub

Sub Blanqueo()
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text13.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text22.Text = ""
    Text23.Text = ""
    Text24.Text = ""
    Cambio = 0
End Sub

Sub blanqueo1()


    Cambio = 0
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        'Case vbKeyUp
        '    OptApellido.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
    End Select
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

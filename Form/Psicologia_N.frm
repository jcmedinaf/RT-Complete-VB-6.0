VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConsultaPsicologicaNoA 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Psicológica Niño o Adolecente"
   ClientHeight    =   8355
   ClientLeft      =   5115
   ClientTop       =   795
   ClientWidth     =   13200
   Icon            =   "Psicologia_N.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   13200
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   8295
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   12975
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   2295
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   12735
         Begin VB.ComboBox CboSexo 
            Height          =   315
            Left            =   6120
            TabIndex        =   43
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   3960
            TabIndex        =   42
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1320
            TabIndex        =   41
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1320
            TabIndex        =   40
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1320
            TabIndex        =   39
            Top             =   360
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DtpFechaActual 
            Height          =   315
            Left            =   4080
            TabIndex        =   44
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
            TabIndex        =   45
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
            TabIndex        =   46
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51314689
            CurrentDate     =   39813
         End
         Begin Crystal.CrystalReport CrystalReport2 
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
            TabIndex        =   47
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
            MICON           =   "Psicologia_N.frx":1002
            PICN            =   "Psicologia_N.frx":101E
            PICH            =   "Psicologia_N.frx":12BA
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
            TabIndex        =   48
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
            MICON           =   "Psicologia_N.frx":14EF
            PICN            =   "Psicologia_N.frx":150B
            PICH            =   "Psicologia_N.frx":1794
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
            TabIndex        =   49
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
            MICON           =   "Psicologia_N.frx":1A2C
            PICN            =   "Psicologia_N.frx":1A48
            PICH            =   "Psicologia_N.frx":1BEC
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
            TabIndex        =   60
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
            TabIndex        =   59
            Top             =   480
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Sexo:"
            Height          =   195
            Left            =   5520
            TabIndex        =   58
            Top             =   960
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Edad:"
            Height          =   195
            Left            =   3480
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
            Top             =   450
            Width           =   840
         End
         Begin VB.Label Label19 
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
            TabIndex        =   51
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1575
            Left            =   10800
            Picture         =   "Psicologia_N.frx":1D8B
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
            TabIndex        =   50
            Top             =   1860
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   7440
         Width           =   3495
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
            TabIndex        =   20
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad"
            Top             =   240
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2280
            TabIndex        =   19
            ToolTipText     =   "Buscar"
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
            MICON           =   "Psicologia_N.frx":2909
            PICN            =   "Psicologia_N.frx":2925
            PICH            =   "Psicologia_N.frx":2B8A
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
         Left            =   3720
         TabIndex        =   35
         Top             =   7440
         Width           =   9135
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   8040
            TabIndex        =   18
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
            MICON           =   "Psicologia_N.frx":2E1C
            PICN            =   "Psicologia_N.frx":2E38
            PICH            =   "Psicologia_N.frx":3001
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
            TabIndex        =   13
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
            MICON           =   "Psicologia_N.frx":3236
            PICN            =   "Psicologia_N.frx":3252
            PICH            =   "Psicologia_N.frx":34E1
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
            TabIndex        =   12
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
            MICON           =   "Psicologia_N.frx":3922
            PICN            =   "Psicologia_N.frx":393E
            PICH            =   "Psicologia_N.frx":3ACB
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
            Left            =   6840
            TabIndex        =   17
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
            MICON           =   "Psicologia_N.frx":3D00
            PICN            =   "Psicologia_N.frx":3D1C
            PICH            =   "Psicologia_N.frx":3FFE
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
            Left            =   6000
            TabIndex        =   16
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
            MICON           =   "Psicologia_N.frx":424F
            PICN            =   "Psicologia_N.frx":426B
            PICH            =   "Psicologia_N.frx":4501
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
            Left            =   5400
            TabIndex        =   15
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
            MICON           =   "Psicologia_N.frx":4760
            PICN            =   "Psicologia_N.frx":477C
            PICH            =   "Psicologia_N.frx":4A11
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
            TabIndex        =   14
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
            MICON           =   "Psicologia_N.frx":4C6D
            PICN            =   "Psicologia_N.frx":4C89
            PICH            =   "Psicologia_N.frx":4DAE
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
            TabIndex        =   37
            ToolTipText     =   "Eliminar"
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
            MICON           =   "Psicologia_N.frx":503E
            PICN            =   "Psicologia_N.frx":505A
            PICH            =   "Psicologia_N.frx":51FE
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
         Caption         =   "Psicologia de Niños y Adolecentes"
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   12735
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   1320
            Top             =   2400
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   1800
            TabIndex        =   0
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1800
            TabIndex        =   1
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox Text11 
            Height          =   405
            Left            =   1800
            TabIndex        =   3
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox Text13 
            Height          =   855
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox Text14 
            Height          =   855
            Left            =   9000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox Text15 
            Height          =   855
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox Text16 
            Height          =   855
            Left            =   9000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox Text17 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   2880
            Width           =   4095
         End
         Begin VB.TextBox Text20 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   3960
            Width           =   12495
         End
         Begin VB.TextBox Text18 
            Height          =   735
            Left            =   4320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   2880
            Width           =   4575
         End
         Begin VB.TextBox Text19 
            Height          =   735
            Left            =   9000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   2880
            Width           =   3615
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   2760
            Top             =   2400
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   3360
            Top             =   2400
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Historia Cl{inica Integral"
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Padre:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   570
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profesión del Padre:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profesión de la Madre:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   2025
            Width           =   1590
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre de la Madre:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Relaciones Familiares?"
            Height          =   255
            Left            =   4920
            TabIndex        =   30
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Conducta?"
            Height          =   255
            Left            =   9000
            TabIndex        =   29
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Castigo ó Premio?"
            Height          =   255
            Left            =   4920
            TabIndex        =   28
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Enfermedad en Parientes?"
            Height          =   255
            Left            =   9000
            TabIndex        =   27
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Post - Natal"
            Height          =   255
            Left            =   9000
            TabIndex        =   26
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Peri - Natal"
            Height          =   255
            Left            =   4320
            TabIndex        =   25
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre-Natal"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Impresión Psicológica"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   3720
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "FrmConsultaPsicologicaNoA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsPaciente As New ADODB.Recordset 'tabla Registro
Dim RsTemp As New ADODB.Recordset
Dim RsPsicologia As New ADODB.Recordset 'tabla Registro
Dim Cambio
Dim RegNew
Dim IdPsic As String
Dim NuevoId As String
Dim IdPacP


Sub Blanqueo()
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
End Sub

Sub blanqueo1()
'Text8.Text = ""
'Text9.Text = ""
'Text10.Text = ""
'Text11.Text = ""
'Text13.Text = ""
'Text14.Text = ""
'Text15.Text = ""
'Text16.Text = ""
'Text17.Text = ""
'Text18.Text = ""
'Text19.Text = ""
'Text20.Text = ""
End Sub

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
Frame2.Enabled = True

CSql = "Select * From Psicologia_N where IdPaciente='" & IdPacP & "'"
Set RsPsicologia = CrearRS(CSql)

If RsPsicologia.RecordCount > 0 Then
    Text8.SetFocus
Else
    Call Blanqueo
    Text8.SetFocus
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
    
    CSql = "select * from psicologia_n where IdPaciente = " & IdPacP & " and Activo=1"
    Set RsPsicologia = CrearRS(CSql)
    Call Carga_Datos_Psicon
    
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros"
    IdPacP = ""
End If
End Sub

Public Sub BtnBuscar_Click()
On Error Resume Next
If TxtBuscar.Text = "Busqueda" Or TxtBuscar.Text = "" Then
    CSql = "select * from Paciente order by IdPaciente" ' where (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) <= 18)"
Else
    CSql = "select * from Paciente where Historia='" & TxtBuscar.Text & "' or cedulaP = " & Val(TxtBuscar.Text) & " or nombreP like '%" & TxtBuscar.Text & "%'" ' AND (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) <= 18)"
End If
Set RsPaciente = CrearRS(CSql)

If RsPaciente.RecordCount = 0 Then
    MsgBox "No Existe el registro", vbInformation + vbOKOnly, "No hay datos"
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    DtpFechaRegistro.Value = Now
    DtpFechaNac.Value = Now
    DtpFechaActual.Value = Now
    Label19.Caption = ""
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    CboSexo.ListIndex = -1
    NoReg = "Registro 0 / 0"
    IdPacP = ""
    Exit Sub
End If

Call Carga_De_Datos

CSql = "select * from psicologia_n where IdPaciente = " & IdPacP & " And IdLIdPac='" & IdLIdPac & "' and Activo=1"
Set RsPsicologia = CrearRS(CSql)

If RsPsicologia.RecordCount = 0 Then
    MsgBox "Este Paciente no posee datos en la tabla de Psicología", vbInformation + vbOKOnly, "No hay datos"
    RegNew = 1
Else
    Call Carga_Datos_Psicon
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
BtnImprimir.Enabled = False
BtnEliminar.Enabled = False
BtnAgregar.Enabled = True
Frame2.BackColor = &HEAEFEF

Cambio = 0
Call Carga_De_Datos
    
CSql = "Select * from psicologia_n where IdPaciente = " & IdPacP & " And IdLIdPac='" & IdLIdPac & "' And Activo=1"
Set RsPsicologia = CrearRS(CSql)
Call Carga_Datos_Psicon

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

CSql = "Update psicologia_N set Activo=0 Where IdPsicologia=" & IdPsic & " And IdL='" & IdLIdInf & "'"
Set RsTemp = CrearRS(CSql)

Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "BORRAR", "se elimino en la tabla psicologia_n el registro de Id=" & IdPsic)

MsgBox "El registro fue eliminado.", vbInformation + vbOKOnly, "Operacion Exitosa"
EnviarRegPendiente IdPsic, IdLIdInf
BtnDesHacer_Click

End Sub

Sub BorrarHosting()
On Error GoTo salir
ConectarHosting

CSql = "Update psicologia_N set Activo=2 Where IdPsicologia=" & IdPsic
Set RsTemp = CrearRS(CSql)

salir:
If WebCnn.State = 0 Then
    BorrarHosting
Else
    GoTo f
End If

f:
If Err.Number <> 0 Then MsgBox Err.Description
If WebCnn.State = 1 Then
    WebCnn.Close
Else
    Exit Sub
End If
End Sub


Private Sub BtnGuardarActualizar_Click()
On Error Resume Next


If IdPacP = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
If Cambio = 0 Then MsgBox "no se han realizado cambios!", vbExclamation + vbOKOnly, "Informacion": Exit Sub

  If Text8.Text = "" Then
        f = "Padre"
        GoTo noguardA
        End If
        
        If Text9.Text = "" Then
        f = "PadreP"
        GoTo noguardA
        End If
        
        If Text10.Text = "" Then
        f = "Madre"
        GoTo noguardA
        End If
                        
        If Text11.Text = "" Then
        f = "MadreP"
        GoTo noguardA
        End If
        
        If Text13.Text = "" Then
        f = "Familia"
        GoTo noguardA
        End If
        
        If Text14.Text = "" Then
        f = "Conducta"
        GoTo noguardA
        End If
        
        If Text15.Text = "" Then
        f = "Premio"
        GoTo noguardA
        End If
        
        If Text16.Text = "" Then
        f = "Enfermedad"
        GoTo noguardA
        End If
        
        If Text17.Text = "" Then
        f = "Prenatal"
        GoTo noguardA
        End If
        
        If Text18.Text = "" Then
        f = "Perinatal"
        GoTo noguardA
        End If
        
        If Text19.Text = "" Then
        f = "Posnatal"
        GoTo noguardA
        End If
                
        If Text20.Text = "" Then
        f = "Observacion"
        GoTo noguardA
        End If


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Bloque que verifica si hay internet
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


resp = MsgBox("Se Guardaran los cambios realizados, Desea Continuar?", vbExclamation + vbYesNo, "Confirmar")

If resp = 7 Then Exit Sub

CSql = "SELECT MAX(idpsicologia)+1 as NuevoId FROM psicologia_N "
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = RsTemp.Fields("NuevoId").Value
Else
    NuevoId = "0"
End If
        
Select Case RegNew
    
    Case Is = 0   'actualiza

       If Cambio = 1 Then
           'hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos
            
            CSql = "Select * From Psicologia_N Where IdPsicologia='" & IdPsic & "' And IdL='" & IdLIdInf & "'"
            Set RsTemp = CrearRS(CSql)
                        
            RsTemp.Fields("idusuario").Value = IdUser
            RsTemp.Fields("PADRE").Value = Text8.Text
            RsTemp.Fields("PADREP").Value = Text9.Text
            RsTemp.Fields("MADRE").Value = Text10.Text
            RsTemp.Fields("MADREP").Value = Text11.Text
            RsTemp.Fields("FAMILIA").Value = Text13.Text
            RsTemp.Fields("CONDUCTA").Value = Text14.Text
            RsTemp.Fields("PREMIO").Value = Text15.Text
            RsTemp.Fields("ENFERMEDAD").Value = Text16.Text
            RsTemp.Fields("PRENATAL").Value = Text17.Text
            RsTemp.Fields("PERINATAL").Value = Text18.Text
            RsTemp.Fields("POSNATAL").Value = Text19.Text
            RsTemp.Fields("OBSERVACION").Value = Text20.Text
            RsTemp.Fields("ACTIVO").Value = 1
            RsTemp.Fields("FECHA_ACTUAL").Value = Format(Now, "DD/MM/YYYY")
            RsTemp.Update

            EnviarRegPendiente IdPsic, IdLIdInf
            MsgBox "Registro actualizado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        End If
        
    Case Is = 1       'Agrega registro
            
            CSql = "Select * From Psicologia_N"
            Set RsTemp = CrearRS(CSql)
            
            RsTemp.AddNew
            
            IdPsic = NuevoId
            IdLIdInf = NuevoIdL
            
            RsTemp.Fields("IdPsicologia").Value = IdPsic
            RsTemp.Fields("IdL").Value = IdLIdInf
            RsTemp.Fields("IdPaciente").Value = IdPacP
            RsTemp.Fields("IdLIdPac").Value = IdLIdPac
            
            RsTemp.Fields("IdUsuario").Value = IdUser
            RsTemp.Fields("PADRE").Value = Text8.Text
            RsTemp.Fields("PADREP").Value = Text9.Text
            RsTemp.Fields("MADRE").Value = Text10.Text
            RsTemp.Fields("MADREP").Value = Text11.Text
            RsTemp.Fields("FAMILIA").Value = Text13.Text
            RsTemp.Fields("CONDUCTA").Value = Text14.Text
            RsTemp.Fields("PREMIO").Value = Text15.Text
            RsTemp.Fields("ENFERMEDAD").Value = Text16.Text
            RsTemp.Fields("PRENATAL").Value = Text17.Text
            RsTemp.Fields("PERINATAL").Value = Text18.Text
            RsTemp.Fields("POSNATAL").Value = Text19.Text
            RsTemp.Fields("OBSERVACION").Value = Text20.Text
            RsTemp.Fields("ACTIVO").Value = 1
            RsTemp.Fields("FECHA_ACTUAL").Value = Format(Now, "DD/MM/YYYY")
            RsTemp.Update

            EnviarRegPendiente IdPsic, IdLIdInf
    End Select

If Cambio = 1 And RegNew = 0 Then

    If Reg_Actual(0) <> Text8.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo PADRE de (" & Reg_Actual(0) & ") a (" & Text8.Text & ")")
    If Reg_Actual(1) <> Text9.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo PADREP de (" & Reg_Actual(1) & ") a (" & Text9.Text & ")")
    If Reg_Actual(2) <> Text10.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo MADRE de (" & Reg_Actual(2) & ") a (" & Text10.Text & ")")
    If Reg_Actual(3) <> Text11.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo MADREP de (" & Reg_Actual(3) & ") a (" & Text11.Text & ")")
    If Reg_Actual(4) <> Text13.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo FAMILIA de (" & Reg_Actual(4) & ") a (" & Text13.Text & ")")
    If Reg_Actual(5) <> Text14.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo CONDUCTA de (" & Reg_Actual(5) & ") a (" & Text14.Text & ")")
    If Reg_Actual(6) <> Text15.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo PREMIO de (" & Reg_Actual(6) & ") a (" & Text15.Text & ")")
    If Reg_Actual(7) <> Text16.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo ENFERMEDAD de (" & Reg_Actual(7) & ") a (" & Text16.Text & ")")
    If Reg_Actual(8) <> Text17.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo PRENATAL de (" & Reg_Actual(8) & ") a (" & Text17.Text & ")")
    If Reg_Actual(9) <> Text18.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo PERINATAL de (" & Reg_Actual(9) & ") a (" & Text18.Text & ")")
    If Reg_Actual(10) <> Text19.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo POSNATAL de (" & Reg_Actual(10) & ") a (" & Text19.Text & ")")
    If Reg_Actual(11) <> Text20.Text Then Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "MODIFICAR", "se modifico de la tabla psicologia_n el nro de registro (" & IdPsic & ") el campo OBSERVACION de (" & Reg_Actual(11) & ") a (" & Text20.Text & ")")
ElseIf Cambio = 1 And RegNew = 1 Then
    Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "INGRESAR", "se Ingreso en la tabla psicologia_n el nuevo registro de Id=" & NuevoId)
End If

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor Web"

BtnDesHacer_Click

Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbExclamation + vbOKOnly, "Error al Guardar"
Exit Sub
Cambio = 0
End Sub

Sub EnviarRegPendiente(ByVal IdPsic2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Psicologia_N WHERE idpsicologia = " & IdPsic2 & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

StrSen = "INSERT INTO Psicologia_N (["
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
RsRegPendiente.Fields("Modulo").Value = "Psicologia Niños"
RsRegPendiente.Fields("Tabla").Value = "Psicologia_N"
RsRegPendiente.Fields("Condicional").Value = "idpsicologia = " & IdPsic2 & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

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
Consultaa = "N"
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
    
    CSql = "select * from psicologia_n where IdPaciente = " & IdPacP & " and Activo=1"
    Set RsPsicologia = CrearRS(CSql)
    Call Carga_Datos_Psicon
    
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros"
    IdPacP = ""
End If

End Sub


Sub Carga_De_Datos()
   
If RsPaciente.RecordCount = 0 Then
    IdPacP = ""
    IdLIdPac = IdLDefault
    BtnAgregar.Enabled = False
    Frame2.Enabled = True
    NoReg = "Registro 0 / 0"
    Exit Sub
Else
    If Trim(RsPaciente.Fields("cedulap")) <> "" Then Text1.Text = RsPaciente.Fields("cedulap")
    If Trim(RsPaciente.Fields("Fecha_regp")) <> "" Then DtpFechaRegistro.Value = RsPaciente.Fields("Fecha_regp")
    If Trim(RsPaciente.Fields("Nombrep")) <> "" Then Text3.Text = RsPaciente.Fields("Nombrep")
    If Trim(RsPaciente.Fields("Apellidop")) <> "" Then Text4.Text = RsPaciente.Fields("Apellidop")
    If Trim(RsPaciente.Fields("Fecha_nacimientop")) <> "" Then DtpFechaNac.Value = RsPaciente.Fields("Fecha_nacimientop")
    If Trim(RsPaciente.Fields("Edadp")) <> "" Then Text6.Text = RsPaciente.Fields("Edadp")
    If Trim(RsPaciente.Fields("Historia")) <> "" Then Label19.Caption = RsPaciente.Fields("Historia")
    If RsPaciente.Fields("sexop") = 0 Then CboSexo.Text = "Masculino" Else CboSexo.Text = "Femenino"
    
    IdPacP = RsPaciente.Fields("IdPaciente").Value
    IdLIdPac = RsPaciente.Fields("IdL").Value
    
    Me.Caption = "Consulta Psicológica Niño o Adolecente - Paciente: " & IdPacP
    If RsPaciente.Fields("foto") <> "" Then
        If Len(Dir(Foto & "\" & RsPaciente.Fields("foto"))) > 0 Then
            Image2.Picture = LoadPicture(Foto & "\" & RsPaciente.Fields("foto"))
        Else
            Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        End If
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
    BtnAgregar.Enabled = True
    Frame2.Enabled = True
    NoReg = "Registro " & RsPaciente.AbsolutePosition & " / " & RsPaciente.RecordCount
End If
End Sub

Sub Carga_Datos_Psicon()

If RsPsicologia.RecordCount <> 0 Then
    IdPsic = RsPsicologia.Fields("IdPsicologia").Value
    IdLIdInf = RsPsicologia.Fields("IdL").Value
    If Trim(RsPsicologia.Fields("Padre").Value) <> "" Then Text8.Text = RsPsicologia.Fields("Padre").Value
    If Trim(RsPsicologia.Fields("Padrep").Value) <> "" Then Text9.Text = RsPsicologia.Fields("Padrep").Value
    If Trim(RsPsicologia.Fields("Madre").Value) <> "" Then Text10.Text = RsPsicologia.Fields("Madre").Value
    If Trim(RsPsicologia.Fields("Madrep").Value) <> "" Then Text11.Text = RsPsicologia.Fields("Madrep").Value
    If Trim(RsPsicologia.Fields("Familia").Value) <> "" Then Text13.Text = RsPsicologia.Fields("Familia").Value
    If Trim(RsPsicologia.Fields("Conducta").Value) <> "" Then Text14.Text = RsPsicologia.Fields("Conducta").Value
    If Trim(RsPsicologia.Fields("Premio").Value) <> "" Then Text15.Text = RsPsicologia.Fields("Premio").Value
    If Trim(RsPsicologia.Fields("Enfermedad").Value) <> "" Then Text16.Text = RsPsicologia.Fields("Enfermedad").Value
    If Trim(RsPsicologia.Fields("Prenatal").Value) <> "" Then Text17.Text = RsPsicologia.Fields("Prenatal").Value
    If Trim(RsPsicologia.Fields("Perinatal").Value) <> "" Then Text18.Text = RsPsicologia.Fields("Perinatal").Value
    If Trim(RsPsicologia.Fields("Posnatal").Value) <> "" Then Text19.Text = RsPsicologia.Fields("Posnatal").Value
    If Trim(RsPsicologia.Fields("observacion").Value) <> "" Then Text20.Text = RsPsicologia.Fields("observacion").Value
    
    If Trim(RsPsicologia.Fields("Padre").Value) <> "" Then Reg_Actual(0) = RsPsicologia.Fields("Padre").Value Else Reg_Actual(0) = ""
    If Trim(RsPsicologia.Fields("Padrep").Value) <> "" Then Reg_Actual(1) = RsPsicologia.Fields("Padrep").Value Else Reg_Actual(1) = ""
    If Trim(RsPsicologia.Fields("Madre").Value) <> "" Then Reg_Actual(2) = RsPsicologia.Fields("Madre").Value Else Reg_Actual(2) = ""
    If Trim(RsPsicologia.Fields("Madrep").Value) <> "" Then Reg_Actual(3) = RsPsicologia.Fields("Madrep").Value Else Reg_Actual(3) = ""
    If Trim(RsPsicologia.Fields("Familia").Value) <> "" Then Reg_Actual(4) = RsPsicologia.Fields("Familia").Value Else Reg_Actual(4) = ""
    If Trim(RsPsicologia.Fields("Conducta").Value) <> "" Then Reg_Actual(5) = RsPsicologia.Fields("Conducta").Value Else Reg_Actual(5) = ""
    If Trim(RsPsicologia.Fields("Premio").Value) <> "" Then Reg_Actual(6) = RsPsicologia.Fields("Premio").Value Else Reg_Actual(6) = ""
    If Trim(RsPsicologia.Fields("Enfermedad").Value) <> "" Then Reg_Actual(7) = RsPsicologia.Fields("Enfermedad").Value Else Reg_Actual(7) = ""
    If Trim(RsPsicologia.Fields("Prenatal").Value) <> "" Then Reg_Actual(8) = RsPsicologia.Fields("Prenatal").Value Else Reg_Actual(8) = ""
    If Trim(RsPsicologia.Fields("Perinatal").Value) <> "" Then Reg_Actual(9) = RsPsicologia.Fields("Perinatal").Value Else Reg_Actual(9) = ""
    If Trim(RsPsicologia.Fields("Posnatal").Value) <> "" Then Reg_Actual(10) = RsPsicologia.Fields("Posnatal").Value Else Reg_Actual(10) = ""
    If Trim(RsPsicologia.Fields("observacion").Value) <> "" Then Reg_Actual(11) = RsPsicologia.Fields("observacion").Value Else Reg_Actual(11) = ""
    
    Cambio = 0: RegNew = 0
    BtnImprimir.Enabled = True
    BtnEliminar.Enabled = True
    BtnGuardarActualizar.Enabled = True
    
Else
    IdPsic = ""
    IdLIdInf = IdLDefault
    Cambio = 0: RegNew = 1
    BtnImprimir.Enabled = False
    BtnEliminar.Enabled = False
    BtnGuardarActualizar.Enabled = False
    
    For i = 0 To 20
        Reg_Actual(i) = ""
    Next i
End If
End Sub

Private Sub CboSexo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnListaEspera.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyLeft
            Text6.SetFocus
        Case vbKeyRight
            BtnListaEspera.SetFocus
        Case vbKeyDown
            Text13.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaActual_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub DtpFechaNac_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyRight
            Text6.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
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
    If Trim(CSql) <> "" Then RsPaciente.Close

Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
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

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text10.SelStart + 1) > Len(Text10.Text)) Or (Shift = 0 And Text10.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case vbKeyUp
            Text9.SetFocus
        Case vbKeyRight
            Text15.SetFocus
        Case vbKeyDown
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text11.SelStart + 1) > Len(Text11.Text)) Or (Shift = 0 And Text11.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case vbKeyUp
            Text10.SetFocus
        Case vbKeyRight
            Text15.SetFocus
        Case vbKeyDown
            Text17.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text13.SelStart + 1) > Len(Text13.Text)) Or (Shift = 0 And Text13.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text14.SetFocus
        Case vbKeyUp
            CboSexo.SetFocus
        Case vbKeyLeft
            Text8.SetFocus
        Case vbKeyRight
            Text14.SetFocus
        Case vbKeyDown
            Text15.SetFocus
    End Select
End If
End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text14.SelStart + 1) > Len(Text14.Text)) Or (Shift = 0 And Text14.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text15.SetFocus
        Case vbKeyUp
            BtnLlamar.SetFocus
        Case vbKeyLeft
            Text13.SetFocus
        Case vbKeyDown
            Text16.SetFocus
    End Select
End If
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text15.SelStart + 1) > Len(Text15.Text)) Or (Shift = 0 And Text15.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text14.SetFocus
        Case vbKeyUp
            Text13.SetFocus
        Case vbKeyLeft
            Text11.SetFocus
        Case vbKeyRight
            Text16.SetFocus
        Case vbKeyDown
            Text18.SetFocus
    End Select
End If
End Sub

Private Sub Text16_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text16.SelStart + 1) > Len(Text16.Text)) Or (Shift = 0 And Text16.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text17.SetFocus
        Case vbKeyUp
            Text14.SetFocus
        Case vbKeyLeft
            Text15.SetFocus
        Case vbKeyDown
            Text19.SetFocus
    End Select
End If
End Sub

Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text17.SelStart + 1) > Len(Text17.Text)) Or (Shift = 0 And Text17.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text18.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text20.SetFocus
    End Select
End If
End Sub

Private Sub Text18_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text18.SelStart + 1) > Len(Text18.Text)) Or (Shift = 0 And Text18.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text19.SetFocus
        Case vbKeyUp
            Text15.SetFocus
        Case vbKeyLeft
            Text17.SetFocus
        Case vbKeyRight
            Text19.SetFocus
        Case vbKeyDown
            Text20.SetFocus
    End Select
End If
End Sub

Private Sub Text19_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text19.SelStart + 1) > Len(Text19.Text)) Or (Shift = 0 And Text19.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text19.SetFocus
        Case vbKeyUp
            Text16.SetFocus
        Case vbKeyLeft
            Text18.SetFocus
        Case vbKeyDown
            Text20.SetFocus
    End Select
End If
End Sub

Private Sub Text20_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text20.SelStart + 1) > Len(Text20.Text)) Or (Shift = 0 And Text20.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text17.SetFocus
        Case vbKeyDown
            'OptApellido.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaNac.SetFocus
        Case vbKeyUp
            DtpFechaActual.SetFocus
        Case vbKeyLeft
            Text4.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Text3.SetFocus
        Case vbKeyDown
            DtpFechaNac.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboSexo.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyLeft
            DtpFechaNac.SetFocus
        Case vbKeyRight
            CboSexo.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text8.SelStart + 1) > Len(Text8.Text)) Or (Shift = 0 And Text8.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            DtpFechaNac.SetFocus
        Case vbKeyRight
            Text13.SetFocus
        Case vbKeyDown
            Text9.SetFocus
    End Select
End If
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If (Shift = 0 And (Text9.SelStart + 1) > Len(Text9.Text)) Or (Shift = 0 And Text9.SelStart = 0) Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            DtpFechaNac.SetFocus
        Case vbKeyRight
            Text13.SetFocus
        Case vbKeyDown
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub TxtBuscar_GotFocus()
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

Private Sub TxtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
'        Case vbKeyUp
'            OptApellido.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
    End Select
End If
End Sub

Private Sub TxtBuscar_LostFocus()
If Trim(TxtBuscar.Text) = "" Then TxtBuscar.Text = "Busqueda"
End Sub

Private Sub Form_Load()
Centrar Me
ModulO = 1

For i = 0 To 20
    Reg_Actual(i) = ""
Next

CSql = "select * from Paciente where (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) <= 18) Order by IdPaciente"
Set RsPaciente = CrearRS(CSql)

'Call Carga_De_Datos

'CSql = "select * from psicologia_n where IdPaciente = " & IdPacP & " and Activo=1"
'Set RsPsicologia = CrearRS(CSql)

'CrystalReport1.ReportFileName = Direc & "\Informes\Historia Clinica Integral.rpt"

DtpFechaActual.Value = Now()
CSql = ""
IdPacP = ""
IdPsic = ""
Cambio = 0
RegNew = 0

'Call carga_datos_psicon

End Sub
 
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
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

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text8_Change()
Cambio = 1
End Sub

Private Sub Text9_Change()
Cambio = 1
End Sub

Private Sub Text10_Change()
Cambio = 1
End Sub

Private Sub Text11_Change()
Cambio = 1
End Sub

Private Sub Text13_Change()
Cambio = 1
End Sub

Private Sub Text14_Change()
Cambio = 1
End Sub

Private Sub Text15_Change()
Cambio = 1
End Sub

Private Sub Text16_Change()
Cambio = 1
End Sub

Private Sub Text17_Change()
Cambio = 1
End Sub

Private Sub Text18_Change()
Cambio = 1
End Sub

Private Sub Text19_Change()
Cambio = 1
End Sub

Private Sub Text20_Change()
Cambio = 1
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
If Len(Trim(Text8.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
If Len(Trim(Text9.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If Len(Trim(Text10.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
If Len(Trim(Text11.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else
'KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
If Len(Trim(Text13.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
If Len(Trim(Text14.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If Len(Trim(Text15.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
If Len(Trim(Text16.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If Len(Trim(Text17.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text18_KeyPress(KeyAscii As Integer)
If Len(Trim(Text18.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub
Private Sub Text19_KeyPress(KeyAscii As Integer)
If Len(Trim(Text19.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If Len(Trim(Text20.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else

End If
End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

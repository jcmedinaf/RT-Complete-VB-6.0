VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConsultaPsicologicaNoA 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Psicológica Niño o Adolecente"
   ClientHeight    =   8355
   ClientLeft      =   5115
   ClientTop       =   795
   ClientWidth     =   13530
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   13530
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   8295
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   13335
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   855
         Left            =   120
         TabIndex        =   58
         Top             =   7320
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
            TabIndex        =   24
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad"
            Top             =   360
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   495
            Left            =   2280
            TabIndex        =   23
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":0000
            PICN            =   "Psicologia_N.frx":001C
            PICH            =   "Psicologia_N.frx":0281
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
         Height          =   855
         Left            =   3840
         TabIndex        =   57
         Top             =   7320
         Width           =   9375
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   7920
            TabIndex        =   19
            ToolTipText     =   "Cerrar Tablas de Pacientes"
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
            MICON           =   "Psicologia_N.frx":0513
            PICN            =   "Psicologia_N.frx":052F
            PICH            =   "Psicologia_N.frx":06F8
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
            Left            =   1200
            TabIndex        =   13
            ToolTipText     =   "Guardar / Actualizar Pacientes"
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
            MICON           =   "Psicologia_N.frx":092D
            PICN            =   "Psicologia_N.frx":0949
            PICH            =   "Psicologia_N.frx":0BD8
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
            Height          =   495
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Agregar Pacientes"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":1019
            PICN            =   "Psicologia_N.frx":1035
            PICH            =   "Psicologia_N.frx":11C2
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
            Height          =   495
            Left            =   6720
            TabIndex        =   18
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":13F7
            PICN            =   "Psicologia_N.frx":1413
            PICH            =   "Psicologia_N.frx":16F5
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
            Height          =   495
            Left            =   2400
            TabIndex        =   14
            ToolTipText     =   "Eliminar Usuario"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":1946
            PICN            =   "Psicologia_N.frx":1962
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
            Height          =   495
            Left            =   5880
            TabIndex        =   17
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":1D1A
            PICN            =   "Psicologia_N.frx":1D36
            PICH            =   "Psicologia_N.frx":1FCC
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
            Height          =   495
            Left            =   5280
            TabIndex        =   16
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":222B
            PICN            =   "Psicologia_N.frx":2247
            PICH            =   "Psicologia_N.frx":24DC
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
            Height          =   495
            Left            =   3840
            TabIndex        =   15
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Psicologia_N.frx":2738
            PICN            =   "Psicologia_N.frx":2754
            PICH            =   "Psicologia_N.frx":2879
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
         Caption         =   "Datos del Paciente"
         Height          =   2055
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   13095
         Begin VB.ComboBox CboSexo 
            Height          =   315
            Left            =   4920
            TabIndex        =   56
            Top             =   1500
            Width           =   2295
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   3480
            TabIndex        =   44
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1440
            TabIndex        =   43
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   6480
            TabIndex        =   42
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1440
            TabIndex        =   41
            Top             =   360
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DtpFechaActual 
            Height          =   375
            Left            =   9720
            TabIndex        =   40
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57540609
            CurrentDate     =   39815
         End
         Begin MSComCtl2.DTPicker DtpFechaNac 
            Height          =   375
            Left            =   1440
            TabIndex        =   54
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57540609
            CurrentDate     =   39815
         End
         Begin MSComCtl2.DTPicker DtpFechaRegisto 
            Height          =   375
            Left            =   4800
            TabIndex        =   55
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57540609
            CurrentDate     =   39815
         End
         Begin ChamaleonButton.ChameleonBtn BtnLlamar 
            Height          =   375
            Left            =   9000
            TabIndex        =   21
            ToolTipText     =   "Llamar"
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Psicologia_N.frx":2B09
            PICN            =   "Psicologia_N.frx":2B25
            PICH            =   "Psicologia_N.frx":2DC1
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
            Left            =   7320
            TabIndex        =   22
            ToolTipText     =   "Lista de Espera"
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
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
            MICON           =   "Psicologia_N.frx":2FF6
            PICN            =   "Psicologia_N.frx":3012
            PICH            =   "Psicologia_N.frx":329B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAyuda 
            Height          =   375
            Left            =   10560
            TabIndex        =   20
            ToolTipText     =   "Ayuda"
            Top             =   1440
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Psicologia_N.frx":3533
            PICN            =   "Psicologia_N.frx":354F
            PICH            =   "Psicologia_N.frx":37F1
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
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   11400
            TabIndex        =   60
            Top             =   1750
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Historia:"
            Height          =   195
            Left            =   6240
            TabIndex        =   59
            Top             =   450
            Width           =   870
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Sexo:"
            Height          =   195
            Left            =   4410
            TabIndex        =   53
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Edad:"
            Height          =   195
            Left            =   2955
            TabIndex        =   52
            Top             =   1530
            Width           =   420
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Fecha de Nac.:"
            Height          =   195
            Left            =   210
            TabIndex        =   51
            Top             =   1530
            Width           =   1110
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "A&pellido(s):"
            Height          =   195
            Left            =   210
            TabIndex        =   50
            Top             =   1050
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nombre(s):"
            Height          =   195
            Left            =   5640
            TabIndex        =   49
            Top             =   1050
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha  &Registro:"
            Height          =   195
            Left            =   3600
            TabIndex        =   48
            Top             =   450
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Cédula:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   450
            Width           =   555
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFFF&
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
            Left            =   7200
            TabIndex        =   46
            Top             =   360
            Width           =   1815
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   11280
            Picture         =   "Psicologia_N.frx":3B5B
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Actual "
            Height          =   375
            Left            =   9120
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Psicologia de Niños y Adolecentes"
         Height          =   4815
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   13095
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
            TabIndex        =   4
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox Text14 
            Height          =   855
            Left            =   9000
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox Text15 
            Height          =   855
            Left            =   4920
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox Text16 
            Height          =   855
            Left            =   9000
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox Text17 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   2880
            Width           =   4095
         End
         Begin VB.TextBox Text20 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   3960
            Width           =   12855
         End
         Begin VB.TextBox Text18 
            Height          =   735
            Left            =   4320
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   2880
            Width           =   4575
         End
         Begin VB.TextBox Text19 
            Height          =   735
            Left            =   9000
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   2880
            Width           =   3975
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
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Padre:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Relaciones Familiares?"
            Height          =   255
            Left            =   4920
            TabIndex        =   34
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Conducta?"
            Height          =   255
            Left            =   9000
            TabIndex        =   33
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Castigo ó Premio?"
            Height          =   255
            Left            =   4920
            TabIndex        =   32
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "¿Enfermedad en Parientes?"
            Height          =   255
            Left            =   9000
            TabIndex        =   31
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Post - Natal"
            Height          =   255
            Left            =   9000
            TabIndex        =   30
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Peri - Natal"
            Height          =   255
            Left            =   4320
            TabIndex        =   29
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre-Natal"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Impresión Psicológica"
            Height          =   375
            Left            =   120
            TabIndex        =   27
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
Dim cambio
Dim regnew
Dim IdPsic As String

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
Private Sub BtnAgregar_Click()
IO = 1

If IdPac1 = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False

Call Blanqueo
regnew = 1
Text8.SetFocus
End Sub

Private Sub BtnAnterior_Click()
Call Blanqueo
If RsPaciente.RecordCount <> 0 Then
    RsPaciente.MovePrevious
    If RsPaciente.BOF Then MsgBox "Ha llegado al primer registro!", vbInformation + vbOKOnly, "Primer registro": RsPaciente.MoveFirst
    Call carga_de_datos
    
    CSql = "select * from psicologia_n where IdPaciente = " & IdPac1 & " and Activo=1"
    Set RsPsicologia = CrearRS(CSql)
    Call carga_datos_psicon
    
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros"
    IdPac1 = ""
End If
End Sub

Private Sub BtnBuscar_Click()
 
If TxtBuscar.Text = "Busqueda" Or TxtBuscar.Text = "" Then
    CSql = "select * from Paciente where (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) <= 18)"
Else
    CSql = "select * from Paciente where cedulaP = " & Val(TxtBuscar.Text) & " or nombreP like '%" & TxtBuscar.Text & "%' AND (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) <= 18)"
End If
Set RsPaciente = CrearRS(CSql)

If RsPaciente.RecordCount = 0 Then
    MsgBox "No Existe el registro", vbInformation + vbOKOnly, "No hay datos"
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    DtpFechaRegisto.Value = Now
    DtpFechaNac.Value = Now
    DtpFechaActual.Value = Now
    Label19.Caption = ""
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    CboSexo.ListIndex = -1
    NoReg = "Registro 0 / 0"
    Exit Sub
End If

Call carga_de_datos

CSql = "select * from psicologia_n where IdPaciente = " & IdPac1 & " and Activo=1"
Set RsPsicologia = CrearRS(CSql)

If RsPsicologia.RecordCount = 0 Then
    MsgBox "Este Paciente no posee datos en la tabla de Psicología", vbInformation + vbOKOnly, "No hay datos"
    regnew = 1
Else
    Call carga_datos_psicon
End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Call Blanqueo
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
BtnImprimir.Enabled = True
BtnEliminar.Enabled = True
BtnAgregar.Enabled = True

cambio = 0
Call carga_de_datos
    
CSql = "select * from psicologia_n where IdPaciente = " & IdPac1 & " and Activo=1"
Set RsPsicologia = CrearRS(CSql)
Call carga_datos_psicon

End Sub

Private Sub BtnEliminar_Click()

If IdPac1 = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
If IdPsic = "" Then MsgBox "No existen registros para borrar!", vbExclamation + vbOKOnly, "Informacion": Exit Sub

resp = MsgBox("Se procedera ha eliminar el registro de Psicologia del paciente, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

CSql = "Update psicologia_N set Activo=2 Where IdPsicologia=" & IdPsic
Set RsTemp = CrearRS(CSql)

Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "BORRAR", "se elimino en la tabla psicologia_n el registro de Id=" & IdPsic)

MsgBox "El registro fue eliminado.", vbInformation + vbOKOnly, "Operacion Exitosa"

BtnDesHacer_Click
End Sub

Private Sub BtnGuardarActualizar_Click()
'command2
Dim NuevoId As String

If IdPac1 = "" Then MsgBox "Debe seleccionar un paciente!", vbCritical + vbOKOnly, "Error": Exit Sub
If cambio = 0 Then MsgBox "no se han realizado cambios!", vbExclamation + vbOKOnly, "Informacion": Exit Sub

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


resp = MsgBox("Se Guardaran los cambios realizados, Desea Continuar?", vbExclamation + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

CSql = "SELECT MAX(idpsicologia)+1 as NuevoId FROM psicologia_N "
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    NuevoId = RsTemp.Fields("NuevoId")
Else
    NuevoId = "0"
End If
        
Select Case regnew
    
    Case Is = 0   'actualiza

       If cambio = 1 Then
           'hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos
            CSql = "Update psicologia_N set Activo=0 Where IdPsicologia=" & IdPsic
            Set RsTemp = CrearRS(CSql)
                        
            CSql = "Insert Into Psicologia_N (idpsicologia,idpaciente,idusuario,PADRE,PADREP,MADRE,MADREP,FAMILIA," & _
                "CONDUCTA,PREMIO,ENFERMEDAD,PRENATAL,PERINATAL,POSNATAL,OBSERVACION,ACTIVO,FECHA_ACTUAL) VALUES(" & NuevoId & "," & IdPac1 & _
                "," & IdUser & ",'" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & _
                Text11.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & _
                Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & _
                Text20.Text & "',1,'" & Format(Now, "DD/MM/YYYY") & "')"
            Set RsTemp = CrearRS(CSql)
            MsgBox "Registro actualizado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        End If
        
    Case Is = 1       'Agrega registro
        If cambio = 1 Then
            CSql = "Update psicologia_N set Activo=0 Where Activo=1 and IdPaciente=" & IdPac1
            Set RsTemp = CrearRS(CSql)
            
            CSql = "Insert Into Psicologia_N (idpsicologia,idpaciente,idusuario,PADRE,PADREP,MADRE,MADREP,FAMILIA," & _
                "CONDUCTA,PREMIO,ENFERMEDAD,PRENATAL,PERINATAL,POSNATAL,OBSERVACION,ACTIVO,FECHA_ACTUAL) VALUES(" & NuevoId & "," & IdPac1 & _
                "," & IdUser & ",'" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & _
                Text11.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & _
                Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & _
                Text20.Text & "',1,'" & Format(Now, "DD/MM/YYYY") & "')"
            Set RsTemp = CrearRS(CSql)
            MsgBox "Registro Agregado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        End If

    End Select

If cambio = 1 And regnew = 0 Then
    
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
ElseIf cambio = 1 And regnew = 1 Then
    Call Enviar_Bitacora(IdUser, "PSICOLOGIA NIÑOS-ADOLESCENTES", "INGRESAR", "se Ingreso en la tabla psicologia_n el nuevo registro de Id=" & NuevoId)
End If

    BtnDesHacer_Click
Exit Sub

noguardA:
    msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox msg, vbExclamation + vbOKOnly, "Error al Guardar"
Exit Sub
cambio = 0
End Sub

Private Sub BtnImprimir_Click()
'command4
If Text1.Text <> "" And IdPac1 <> "" Then
    FrmImprimirConsultaPsicologicaNinos.Show vbModal, FrmPrincipal
Else
    MsgBox "Tiene que seleccionar a una paciente", vbOKOnly + vbCritical, "Mensaje de Error"
End If
End Sub

Private Sub BtnListaEspera_Click()
'commad7
ModulO = 1
FrmListaEspera.Show
If Cedul <> "" Then
TxtBuscar.Text = Cedul
Call BtnBuscar_Click
End If
End Sub

Private Sub BtnLlamar_Click()
'command13
Call Llamar
End Sub

Private Sub BtnSiguiente_Click()
Call Blanqueo
If RsPaciente.RecordCount <> 0 Then
    RsPaciente.MoveNext
    If RsPaciente.EOF Then MsgBox "Ha llegado al ultimo registro!", vbInformation + vbOKOnly, "Ultimo registro": RsPaciente.MoveLast
    Call carga_de_datos
    
    CSql = "select * from psicologia_n where IdPaciente = " & IdPac1 & " and Activo=1"
    Set RsPsicologia = CrearRS(CSql)
    Call carga_datos_psicon
    
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros"
    IdPac1 = ""
End If

End Sub


Sub carga_de_datos()
   
If RsPaciente.RecordCount = 0 Then
    IdPac1 = ""
    BtnAgregar.Enabled = False
    NoReg = "Registro 0 / 0"
    Exit Sub
Else
    If Trim(RsPaciente.Fields("cedulap")) <> "" Then Text1.Text = RsPaciente.Fields("cedulap")
    If Trim(RsPaciente.Fields("Fecha_regp")) <> "" Then DtpFechaRegisto.Value = RsPaciente.Fields("Fecha_regp")
    If Trim(RsPaciente.Fields("Nombrep")) <> "" Then Text3.Text = RsPaciente.Fields("Nombrep")
    If Trim(RsPaciente.Fields("Apellidop")) <> "" Then Text4.Text = RsPaciente.Fields("Apellidop")
    If Trim(RsPaciente.Fields("Fecha_nacimientop")) <> "" Then DtpFechaNac.Value = RsPaciente.Fields("Fecha_nacimientop")
    If Trim(RsPaciente.Fields("Edadp")) <> "" Then Text6.Text = RsPaciente.Fields("Edadp")
    If Trim(RsPaciente.Fields("Historia")) <> "" Then Label19.Caption = RsPaciente.Fields("Historia")
    If RsPaciente.Fields("sexop") = 0 Then CboSexo.Text = "Masculino" Else CboSexo.Text = "Femenino"
    IdPac1 = RsPaciente.Fields("idpaciente")
    Me.Caption = "Consulta Psicológica Niño o Adolecente     Paciente: " & IdPac1
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
    NoReg = "Registro " & RsPaciente.AbsolutePosition & " / " & RsPaciente.RecordCount
End If
End Sub
Sub carga_datos_psicon()

If RsPsicologia.RecordCount <> 0 Then
    IdPsic = RsPsicologia.Fields("IdPsicologia").Value
    If Trim(RsPsicologia.Fields("Padre")) <> "" Then Text8.Text = RsPsicologia.Fields("Padre")
    If Trim(RsPsicologia.Fields("Padrep")) <> "" Then Text9.Text = RsPsicologia.Fields("Padrep")
    If Trim(RsPsicologia.Fields("Madre")) <> "" Then Text10.Text = RsPsicologia.Fields("Madre")
    If Trim(RsPsicologia.Fields("Madrep")) <> "" Then Text11.Text = RsPsicologia.Fields("Madrep")
    If Trim(RsPsicologia.Fields("Familia")) <> "" Then Text13.Text = RsPsicologia.Fields("Familia")
    If Trim(RsPsicologia.Fields("Conducta")) <> "" Then Text14.Text = RsPsicologia.Fields("Conducta")
    If Trim(RsPsicologia.Fields("Premio")) <> "" Then Text15.Text = RsPsicologia.Fields("Premio")
    If Trim(RsPsicologia.Fields("Enfermedad")) <> "" Then Text16.Text = RsPsicologia.Fields("Enfermedad")
    If Trim(RsPsicologia.Fields("Prenatal")) <> "" Then Text17.Text = RsPsicologia.Fields("Prenatal")
    If Trim(RsPsicologia.Fields("Perinatal")) <> "" Then Text18.Text = RsPsicologia.Fields("Perinatal")
    If Trim(RsPsicologia.Fields("Posnatal")) <> "" Then Text19.Text = RsPsicologia.Fields("Posnatal")
    If Trim(RsPsicologia.Fields("observacion")) <> "" Then Text20.Text = RsPsicologia.Fields("observacion")
    
    Reg_Actual(0) = RsPsicologia.Fields("Padre").Value
    Reg_Actual(1) = RsPsicologia.Fields("Padrep")
    Reg_Actual(2) = RsPsicologia.Fields("Madre")
    Reg_Actual(3) = RsPsicologia.Fields("Madrep")
    Reg_Actual(4) = RsPsicologia.Fields("Familia")
    Reg_Actual(5) = RsPsicologia.Fields("Conducta")
    Reg_Actual(6) = RsPsicologia.Fields("Premio")
    Reg_Actual(7) = RsPsicologia.Fields("Enfermedad")
    Reg_Actual(8) = RsPsicologia.Fields("Prenatal")
    Reg_Actual(9) = RsPsicologia.Fields("Perinatal")
    Reg_Actual(10) = RsPsicologia.Fields("Posnatal")
    Reg_Actual(11) = RsPsicologia.Fields("observacion")
    
    cambio = 0: regnew = 0
    BtnImprimir.Enabled = True
    BtnEliminar.Enabled = True
Else
    IdPsic = ""
    cambio = 0: regnew = 1
    BtnImprimir.Enabled = False
    BtnEliminar.Enabled = False
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
            DtpFechaRegisto.SetFocus
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

Private Sub DtpFechaRegisto_KeyUp(KeyCode As Integer, Shift As Integer)
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
If Trim(CSql) <> "" Then RsPaciente.Close
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaRegisto.SetFocus
        Case vbKeyRight
            DtpFechaRegisto.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text17.SetFocus
        Case vbKeyDown
            OptApellido.SetFocus
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
If Shift = 0 Then
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
If Shift = 0 Then
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
If KeyAscii = 13 Then Call BtnBuscar_Click
End Sub

Private Sub TxtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        Case vbKeyUp
            OptApellido.SetFocus
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

CSql = "select * from Paciente where (DATEDIFF(year, Fecha_NacimientoP, GETDATE()) <= 18)"
Set RsPaciente = CrearRS(CSql)

Call carga_de_datos

CSql = "select * from psicologia_n where IdPaciente = " & IdPac1 & " and Activo=1"
Set RsPsicologia = CrearRS(CSql)

CrystalReport1.ReportFileName = Direc & "\Informes\Historia Clinica Integral.rpt"

DtpFechaActual.Value = Now()
CSql = ""
IdPsic = ""
cambio = 0
regnew = 0
Call carga_datos_psicon

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
cambio = 1
End Sub

Private Sub Text9_Change()
cambio = 1
End Sub

Private Sub Text10_Change()
cambio = 1
End Sub

Private Sub Text11_Change()
cambio = 1
End Sub

Private Sub Text13_Change()
cambio = 1
End Sub

Private Sub Text14_Change()
cambio = 1
End Sub

Private Sub Text15_Change()
cambio = 1
End Sub

Private Sub Text16_Change()
cambio = 1
End Sub

Private Sub Text17_Change()
cambio = 1
End Sub

Private Sub Text18_Change()
cambio = 1
End Sub

Private Sub Text19_Change()
cambio = 1
End Sub

Private Sub Text20_Change()
cambio = 1
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

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmRadioTerapia 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Area Tecnica"
   ClientHeight    =   9210
   ClientLeft      =   240
   ClientTop       =   825
   ClientWidth     =   14610
   Icon            =   "Tecnico.frx":0000
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   14610
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   2175
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   14175
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   960
            TabIndex        =   51
            Top             =   750
            Width           =   3015
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   5400
            TabIndex        =   44
            Top             =   750
            Width           =   2535
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3480
            TabIndex        =   43
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   960
            TabIndex        =   42
            Top             =   1140
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   960
            TabIndex        =   41
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox CboSexo 
            Height          =   315
            ItemData        =   "Tecnico.frx":1002
            Left            =   5400
            List            =   "Tecnico.frx":100C
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1140
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   375
            Left            =   10080
            TabIndex        =   45
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51183617
            CurrentDate     =   40121
         End
         Begin MSComCtl2.DTPicker DtpFechaRegistro 
            Height          =   375
            Left            =   10080
            TabIndex        =   46
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51183617
            CurrentDate     =   40121
         End
         Begin MSComCtl2.DTPicker DtpFechaCulminacion 
            Height          =   375
            Left            =   10080
            TabIndex        =   47
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51183617
            CurrentDate     =   40121
         End
         Begin ChamaleonButton.ChameleonBtn BtnLlamar 
            Height          =   375
            Left            =   2520
            TabIndex        =   48
            ToolTipText     =   "Llamar"
            Top             =   1680
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
            MICON           =   "Tecnico.frx":1025
            PICN            =   "Tecnico.frx":1041
            PICH            =   "Tecnico.frx":12DD
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
            Left            =   840
            TabIndex        =   49
            ToolTipText     =   "Lista de Espera"
            Top             =   1680
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
            MICON           =   "Tecnico.frx":1512
            PICN            =   "Tecnico.frx":152E
            PICH            =   "Tecnico.frx":17B7
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
            Left            =   4080
            TabIndex        =   50
            ToolTipText     =   "Desocupar al Paciente Atendido"
            Top             =   1680
            Width           =   2415
            _ExtentX        =   4260
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
            MICON           =   "Tecnico.frx":1A4F
            PICN            =   "Tecnico.frx":1A6B
            PICH            =   "Tecnico.frx":1C0F
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
            Left            =   10920
            TabIndex        =   65
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   1680
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Tecnico.frx":1E44
            PICN            =   "Tecnico.frx":1E60
            PICH            =   "Tecnico.frx":20F6
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
            Left            =   10080
            TabIndex        =   66
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   1680
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Tecnico.frx":2355
            PICN            =   "Tecnico.frx":2371
            PICH            =   "Tecnico.frx":2606
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
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Registros: 0"
            Height          =   195
            Left            =   6840
            TabIndex        =   64
            Top             =   1800
            Width           =   1470
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Historia:"
            Height          =   195
            Left            =   4440
            TabIndex        =   62
            Top             =   480
            Width           =   870
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1800
            Left            =   12120
            Picture         =   "Tecnico.frx":2862
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label12 
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
            Left            =   5400
            TabIndex        =   61
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Sexo:"
            Height          =   195
            Left            =   4800
            TabIndex        =   60
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Edad:"
            Height          =   195
            Left            =   2880
            TabIndex        =   59
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de C&ulminación:"
            Height          =   195
            Left            =   8280
            TabIndex        =   58
            Top             =   1200
            Width           =   1620
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de &Inicio:"
            Height          =   195
            Left            =   8760
            TabIndex        =   57
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Medico &Tratante:"
            Height          =   195
            Left            =   4080
            TabIndex        =   56
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "A&pellido(s):"
            Height          =   195
            Left            =   90
            TabIndex        =   55
            Top             =   840
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nombre(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1230
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de &Registro:"
            Height          =   195
            Left            =   8550
            TabIndex        =   53
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Cédula:"
            Height          =   195
            Left            =   315
            TabIndex        =   52
            Top             =   450
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   14175
         Begin VB.Timer Timer1 
            Interval        =   10
            Left            =   13440
            Top             =   840
         End
         Begin VB.TextBox Text5 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   480
            Width           =   8655
         End
         Begin VB.TextBox Text7 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1440
            Width           =   8655
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Height          =   1695
            Left            =   8880
            TabIndex        =   3
            Top             =   240
            Width           =   5175
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   960
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   1200
               Width           =   3615
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   960
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   720
               Width           =   3615
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   960
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   240
               Width           =   3615
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Protocolo:"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   1260
               Width           =   720
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Físico:"
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Top             =   780
               Width           =   480
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Técnico:"
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   300
               Width           =   630
            End
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardarDatosTecnicos 
            Height          =   375
            Left            =   10680
            TabIndex        =   20
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   2040
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
            MICON           =   "Tecnico.frx":3789
            PICN            =   "Tecnico.frx":37A5
            PICH            =   "Tecnico.frx":3A34
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente1 
            Height          =   375
            Left            =   9600
            TabIndex        =   22
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   2040
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Tecnico.frx":3E75
            PICN            =   "Tecnico.frx":3E91
            PICH            =   "Tecnico.frx":4127
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
            Left            =   8880
            TabIndex        =   23
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   2040
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Tecnico.frx":4386
            PICN            =   "Tecnico.frx":43A2
            PICH            =   "Tecnico.frx":4637
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnProtocolos 
            Height          =   375
            Left            =   12480
            TabIndex        =   37
            Top             =   2040
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Protocolos"
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
            MICON           =   "Tecnico.frx":4893
            PICN            =   "Tecnico.frx":48AF
            PICH            =   "Tecnico.frx":4B53
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label TInformes 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Informes: 0"
            Height          =   195
            Left            =   7260
            TabIndex        =   38
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion del Tratamiento:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   2025
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3360
         TabIndex        =   14
         Top             =   8280
         Width           =   10935
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   9840
            TabIndex        =   15
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
            FCOLO           =   12582912
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Tecnico.frx":4F82
            PICN            =   "Tecnico.frx":4F9E
            PICH            =   "Tecnico.frx":5167
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
            Left            =   8640
            TabIndex        =   16
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
            MICON           =   "Tecnico.frx":539C
            PICN            =   "Tecnico.frx":53B8
            PICH            =   "Tecnico.frx":569A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Crystal.CrystalReport CrystalReport2 
            Left            =   3000
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowBorderStyle=   3
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   2520
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowCloseBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin ChamaleonButton.ChameleonBtn BtnReporteDosimetria 
            Height          =   375
            Left            =   4800
            TabIndex        =   63
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Reporte Dosimetrico"
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
            MICON           =   "Tecnico.frx":58EB
            PICN            =   "Tecnico.frx":5907
            PICH            =   "Tecnico.frx":5BA3
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
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   8280
         Width           =   3135
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
            TabIndex        =   19
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad o Historia"
            Top             =   240
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   1680
            TabIndex        =   18
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Tecnico.frx":5E42
            PICN            =   "Tecnico.frx":5E5E
            PICH            =   "Tecnico.frx":60C3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Timer Timer3 
            Interval        =   1000
            Left            =   2280
            Top             =   120
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   5040
         Width           =   14175
         Begin VB.TextBox TxtTAC 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtMetas 
            Height          =   375
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00EAEFEF&
            Height          =   735
            Left            =   120
            TabIndex        =   24
            Top             =   2280
            Width           =   13935
            Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
               Height          =   375
               Left            =   1200
               TabIndex        =   25
               ToolTipText     =   "Guardar / Actualizar"
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Editar"
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
               MICON           =   "Tecnico.frx":6355
               PICN            =   "Tecnico.frx":6371
               PICH            =   "Tecnico.frx":6600
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
               TabIndex        =   26
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
               MICON           =   "Tecnico.frx":6A41
               PICN            =   "Tecnico.frx":6A5D
               PICH            =   "Tecnico.frx":6BEA
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
               TabIndex        =   27
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
               MICON           =   "Tecnico.frx":6E1F
               PICN            =   "Tecnico.frx":6E3B
               PICH            =   "Tecnico.frx":6FDF
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
               Left            =   3720
               TabIndex        =   28
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
               MICON           =   "Tecnico.frx":717E
               PICN            =   "Tecnico.frx":719A
               PICH            =   "Tecnico.frx":72BF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnAplicarTratamiento 
               Height          =   375
               Left            =   9600
               TabIndex        =   33
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Aplicar Tratamiento"
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
               MICON           =   "Tecnico.frx":754F
               PICN            =   "Tecnico.frx":756B
               PICH            =   "Tecnico.frx":7807
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnDatosTomograficos 
               Height          =   375
               Left            =   11520
               TabIndex        =   34
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Datos Para Tomografía"
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
               BCOL            =   -2147483633
               BCOLO           =   -2147483633
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   16777215
               MPTR            =   1
               MICON           =   "Tecnico.frx":7AA6
               PICN            =   "Tecnico.frx":7AC2
               PICH            =   "Tecnico.frx":7D23
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnCampoTratamiento 
               Height          =   375
               Left            =   5160
               TabIndex        =   35
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Campo de Tratamiento"
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
               MICON           =   "Tecnico.frx":7EC6
               PICN            =   "Tecnico.frx":7EE2
               PICH            =   "Tecnico.frx":817A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnParametrosSimulador 
               Height          =   375
               Left            =   7320
               TabIndex        =   36
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Parámetros del Simulador"
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
               BCOL            =   -2147483633
               BCOLO           =   -2147483633
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   16777215
               MPTR            =   1
               MICON           =   "Tecnico.frx":83FC
               PICN            =   "Tecnico.frx":8418
               PICH            =   "Tecnico.frx":8679
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
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   2778
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Inicio"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Técnica y Sitio Anatómico"
               Object.Width           =   5468
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Energia (MV)"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Profundidad o Iso"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Frac/dia"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Frac/Sem"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Frac/Total"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Dosis/Frac"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Dosis Total"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Id"
               Object.Width           =   2
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "IdL"
               Object.Width           =   2
            EndProperty
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TAC:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2010
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Metas:"
            Height          =   195
            Left            =   1680
            TabIndex        =   31
            Top             =   2010
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "FrmRadioTerapia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPacientes As New ADODB.Recordset 'tabla Registro
Public RsTecnica As New ADODB.Recordset 'tabla Registro
Dim RsActualizar As New ADODB.Recordset
Dim RsGuardar As New ADODB.Recordset
Dim RsInformeMed As New ADODB.Recordset
Dim SQL As String
Dim Actualizar
Dim cntpac
Dim totpac
Dim IdTec1
Public IdLIdInf
Public IdLIdTec1 As String
Public IdPaciente As String
Public IdLIdPac As String
Public IdUsuario As String
Public IdTecnico As String
Public IdFisico As String
Public IdProtocolo As String
Dim Reg_Actual(0 To 4) As String
Dim C As String
Dim WrtError
Public Camp, DF
Public Camp2
Public NombreTecnico
Public NombreFisico
Public InicialTec
Public IdInf
Dim IdPacT
Dim RsTecnica1 As New ADODB.Recordset

Sub CargarDataGrid(dg As DataGrid)
Camp = ""
dg.MarqueeStyle = dbgHighlightRow
Set dg.DataSource = RsTecnica
dg.Refresh
End Sub

Private Sub BtnAgregar_Click()
If Trim(IdPacT) = "" Then
    MsgBox "Debe de seleccionar un paciente!", vbExclamation + vbOKOnly, "Error!"
    Exit Sub
Else
    Call Agregar
End If

End Sub

Private Sub BtnAnterior_Click()
If Trim(IdPacT) = "" Then Exit Sub
If RsPacientes.RecordCount <> 0 Then
    If RsPacientes.BOF Then
        RsPacientes.MoveLast
    Else
        RsPacientes.MovePrevious
        If RsPacientes.BOF Then RsPacientes.MoveLast
    End If
    
    Call Carga_De_Datos

    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If
End Sub

Private Sub BtnAnterior1_Click()
If Trim(IdPacT) = "" Then Exit Sub

If RsInformeMed.RecordCount <> 0 Then
    If RsInformeMed.BOF Then
        RsInformeMed.MoveLast
    Else
        RsInformeMed.MovePrevious
        If RsInformeMed.BOF Then RsInformeMed.MoveLast
    End If
    Call carga_de_datos_tec

    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If
End Sub

Private Sub BtnAplicarTratamiento_Click()
Dim RsVerificarCamposTrata As New ADODB.Recordset

If Trim(IdPacT) = "" Then
    MsgBox "Debe de seleccionar un paciente!", vbExclamation + vbOKOnly, "Error!"
    Exit Sub
Else
    If Text3.Text = "" And Text4.Text = "" Then
        Exit Sub
    Else
        If Camp <> "" Then
            CSql = "Select * From Tecnica2 Where IdPaciente = '" & IdPacT & "' And IdLIdPac='" & IdLIdPac & "'"
            Set RsVerificarCamposTrata = CrearRS(CSql)
            If RsVerificarCamposTrata.RecordCount > 0 Then
            
                FrmTratamientoDiario.TxtApellido.Text = Text3.Text
                FrmTratamientoDiario.TxtNombre.Text = Text4.Text
                FrmTratamientoDiario.LblNoHistoria.Caption = Label12.Caption
                
                'ListView1.SelectedItem.ListSubItems(9).Text
                FrmTratamientoDiario.Caption = "Tratamiento Diario - Técnica: " & ListView1.SelectedItem.ListSubItems(1).Text
'                If Not IsNull(RsVerificarCamposTrata.Fields("descripcion").Value) Then
'                    FrmTratamientoDiario.Caption = "Tratamiento Diario - Técnica: " & RsVerificarCamposTrata.Fields("descripcion").Value
'                End If
                FrmTratamientoDiario.Show vbModal, FrmPrincipal
            Else
                MsgBox "Ingrese primero los campos de tratamientos para poder empezar el tratamiento del paciente", vbInformation + vbOKOnly, "Mensaje OncoAmerica"
                Exit Sub
            End If
        Else
            MsgBox "Seleccione la Técnica o Sitio Anatómico para ingresarle el tratamiento!!", vbInformation + vbOKOnly, "Mensaje OncoAmerica"
            Exit Sub
        End If
    
    
    End If
End If
End Sub

Public Sub BtnBuscar_Click()
Dim CSql As String

Limpiar_Campos

If Replace(TxtBuscar.Text, " ", "") = "" Or Replace(TxtBuscar.Text, " ", "") = "Busqueda" Then
    f = "Buscar"
    CSql = "Select * From Paciente where (Historia like '%hoam%' or Historia like '%HOAM%')  order by idpaciente"
    Set RsPacientes = CrearRS(CSql)
    Call Carga_De_Datos
    Exit Sub
End If

CSql = "Select * From Paciente Where Historia='" & TxtBuscar.Text & "' or CedulaP = " & Val(TxtBuscar.Text) & " or NombreP like '%" & TxtBuscar.Text & "%' or ApellidoP like '%" & TxtBuscar.Text & "%'"
Set RsPacientes = CrearRS(CSql)

If RsPacientes.EOF Then
    MsgBox "No Existe el registro", vbCritical + vbOKOnly, "Error"
    If RsPacientes.RecordCount <> 0 Then
        RsPacientes.MoveFirst
    Else
        IdPaciente = ""
        IdPacT = ""
        NoReg.Caption = "Total de Registros: " & RsPacientes.RecordCount
        LlenarGrid
        
    End If
    Exit Sub
End If

NoReg.Caption = RsPacientes.RecordCount
Call Carga_De_Datos

End Sub

Sub Limpiar_Campos()
    IdPaciente = ""
    IdUsuario = ""
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Text10.Text = ""
    Text5.Text = ""
    Text7.Text = ""
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    CboSexo.Text = ""
    DtpFechaRegistro.Value = Now
    DtpFechaInicio.Value = Now
    DtpFechaCulminacion.Value = Now + 7
    
End Sub

Private Sub BtnCampoTratamiento_Click()

If Trim(IdPacT) = "" Then
    MsgBox "Debe de seleccionar un paciente!", vbExclamation + vbOKOnly, "Error!"
    Exit Sub
Else
    CSql = "Select * From Tecnica where IdPaciente='" & IdPacT & "' And IdLIdPac='" & IdLIdPac & "'"
    Set RsTecnica = CrearRS(CSql)
    
    If RsTecnica.RecordCount > 0 Then
        If Camp <> "" Then
            If IdPaciente <> "" Then
                FrmEditarTratamientos.IdPacT = IdPacT
                FrmEditarTratamientos.IdLIdPacT = IdLIdPac
                FrmEditarTratamientos.Show vbModal, FrmPrincipal
            Else
                MsgBox "Debe seleccionar un paciente para agregar datos al registro!", vbExclamation + vbOKCancel, "No puede agregar datos!"
                Exit Sub
            End If
        Else
            MsgBox "Seleccione la Técnica o Sitio Anatómico para ingresarle el tratamiento!!", vbInformation + vbOKOnly, "Mensaje OncoAmerica"
            Exit Sub
        End If
    Else
        MsgBox "Ingrese la(s) Técnica(s) o Sitio(s) Anatómico(s) del tratamiento del paciente!!", vbInformation + vbOKOnly, "Mensaje OncoAmerica"
        Exit Sub
    End If
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDatosTomograficos_Click()

If Trim(IdPacT) = "" Then
    MsgBox "Debe de seleccionar un paciente!", vbExclamation + vbOKOnly, "Error!"
    Exit Sub
Else
    If IdPaciente <> "" Then
        'FrmParametrosSimulacion.Show
        FrmEditorParametrosSimulacion.Show vbModal, FrmPrincipal
    Else
        Msg = "Debe Seleccionar un Paciente"
        MsgBox Msg, vbExclamation + vbOKOnly, "Debe Seleccionar un paciente"
    End If
End If

End Sub

Private Sub BtnDesocuparAlPacienteAtendido_Click()
Dim bdlista88 As New ADODB.Recordset
CSql = "Delete From Ubi_Paciente Where Modul = " & ModulO
Set bdlista88 = CrearRS(CSql)
End Sub

Private Sub BtnEliminar_Click()
If Trim(IdPacT) = "" Then
    MsgBox "Debe de seleccionar un paciente!", vbExclamation + vbOKOnly, "Error!"
    Exit Sub
Else
    Call Eliminar
End If

End Sub

Private Sub BtnGuardarActualizar_Click()
Call Editar
End Sub

Private Sub BtnGuardarDatosTecnicos_Click()
On Error GoTo WrtError

CSql = "Select * From Tecnica1 Where IdPaciente='" & IdPacT & "' And IdLIdPac='" & IdLIdPac & "' And IdInforme='" & IdInf & "' And IdLIdInf='" & IdLIdInf & "' And Activo='1'"
Set RsTecnica1 = CrearRS(CSql)

If RsTecnica1.RecordCount = 0 Then
    Actualizar = 0
Else
    IdTec1 = RsTecnica1.Fields("Id").Value
    IdLIdTec1 = RsTecnica1.Fields("IdL").Value
    Actualizar = 1
End If
'command2

If IdPaciente = "" Then MsgBox "Debe seleccionar un paciente para agregar datos al registro!", vbExclamation + vbOKOnly, "No puede agregar datos!": Exit Sub

If TxtBuscar.Text = "" Then
    f = "No hay registro seleccionado"
    'GoTo noguardA
End If

If Replace(Combo1.Text, " ", "") = "" Then MsgBox "Ingrese la informacion de Técnico!", vbExclamation + vbOKOnly, "Faltan Datos": Combo1.SetFocus: Exit Sub
If Replace(Combo2.Text, " ", "") = "" Then MsgBox "Ingrese la informacion de Físico!", vbExclamation + vbOKOnly, "Faltan Datos": Combo2.SetFocus: Exit Sub
If Replace(Combo3.Text, " ", "") = "" Then MsgBox "Ingrese la informacion de Protocolo!", vbExclamation + vbOKOnly, "Faltan Datos": Combo3.SetFocus: Exit Sub

p = MsgBox("Se procedera a guardar los cambios realizados, Desea continuar?", vbQuestion + vbYesNo, "Confirmar!")
If p = vbYes Then 'Exit Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


If Actualizar = 0 Then 'Agrega registro
    
    'hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos
    
    CSql = "Select MAX(Id)+1 as NuevaId From Tecnica1"
    Set RsActualizar = CrearRS(CSql)
    
    If RsActualizar.RecordCount <> 0 Then
        C = RsActualizar.Fields("NuevaId").Value
        Else
        C = "1"
        'Reg_Actual(3) = RsRegTecnicos.Fields("Protocolo")
    End If
    
'    CSql = "Insert Into Tecnica1(Id,IdPaciente,IdUsuario,IdInforme,Tecnico,NombreTecnico,Fisico,NombreFisico,NombreMedicoTratante,Protocolo,Descripcion1,Activo) " & _
'    "VALUES('" & C & "','" & IdPaciente & "','" & IdUser & "','" & IdInf & "','" & Combo1.ItemData(Combo1.ListIndex) & "','" & Trim(Combo1.Text) & "','" & _
'    Combo2.ItemData(Combo2.ListIndex) & "','" & Trim(Combo2.Text) & "','" & Trim(Text10.Text) & "'," & Combo3.ItemData(Combo3.ListIndex) & ", '" & Text7.Text & "',1)"
'
'    Set RsActualizar = CrearRS(CSql)


    CSql = "Select * From Tecnica1"
    Set RsGuardar = CrearRS(CSql)
    
    RsGuardar.AddNew
    
    IdTec1 = C
    IdLIdTec1 = NuevoIdL
    RsGuardar.Fields("Id").Value = IdTec1
    RsGuardar.Fields("IdL").Value = IdLIdTec1
    RsGuardar.Fields("IdPaciente").Value = IdPaciente
    RsGuardar.Fields("IdLIdPac").Value = IdLIdPac
    RsGuardar.Fields("IdInforme").Value = IdInf
    RsGuardar.Fields("IdLIdInf").Value = IdLIdInf
    
    RsGuardar.Fields("IdUsuario").Value = IdUser
    RsGuardar.Fields("Tecnico").Value = Combo1.ItemData(Combo1.ListIndex)
    RsGuardar.Fields("NombreTecnico").Value = Trim(Combo1.Text)
    RsGuardar.Fields("Fisico").Value = Combo2.ItemData(Combo2.ListIndex)
    RsGuardar.Fields("NombreFisico").Value = Trim(Combo2.Text)
    RsGuardar.Fields("NombreMedicoTratante").Value = Trim(Text10.Text)
    RsGuardar.Fields("Protocolo").Value = Combo3.ItemData(Combo3.ListIndex)
    RsGuardar.Fields("Descripcion1").Value = Text7.Text
    RsGuardar.Fields("Activo").Value = 1
    RsGuardar.Update


    Msg = "Registro Agregado satisfactoriamente"
    MsgBox Msg, vbOKOnly + vbInformation, "Operacion Exitosa!"
    Set RsGuardar = Nothing
    
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
    EnviarRegPendiente IdTec1, IdLIdTec1
Else
    
    CSql = "Select * From Tecnica1 Where Id='" & IdTec1 & "' And IdL='" & IdLIdTec1 & "'"
    Set RsActualizar = CrearRS(CSql)
    
    RsActualizar.Fields("IdUsuario").Value = IdUser
    'RsActualizar.Fields("IdInforme").Value = IdInf
    RsActualizar.Fields("Tecnico").Value = Combo1.ItemData(Combo1.ListIndex)
    RsActualizar.Fields("NombreTecnico").Value = Trim(Combo1.Text)
    RsActualizar.Fields("Fisico").Value = Combo2.ItemData(Combo2.ListIndex)
    RsActualizar.Fields("NombreFisico").Value = Trim(Combo2.Text)
    RsActualizar.Fields("NombreMedicoTratante").Value = Trim(Text10.Text)
    RsActualizar.Fields("Protocolo").Value = Combo3.ItemData(Combo3.ListIndex)
    RsActualizar.Fields("Descripcion1").Value = Text7.Text
    RsActualizar.Fields("Activo").Value = 1
    RsActualizar.Update

    MsgBox "El Registro se ha agregado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
    Set RsActualizar = Nothing

    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
    
    EnviarRegPendiente IdTec1, IdLIdTec1
    
End If
End If
If Actualizar = 1 Then
    If Reg_Actual(0) <> Text7.Text Then
        Call Enviar_Bitacora(IdUser, "RADIOTERAPIA", "MODIFICAR", "Se modifico La tabla Tecnica1 cuya Id=" & IdTec1 & ", el campo DESCRIPCION1 de (" & Reg_Actual(0) & ") a (" & Text7.Text & ")")
    End If
    If Reg_Actual(1) <> Combo1.ItemData(Combo1.ListIndex) Then
        Call Enviar_Bitacora(IdUser, "RADIOTERAPIA", "MODIFICAR", "Se modifico La tabla Tecnica1 cuya Id=" & IdTec1 & ", el campo FISICO de (" & Reg_Actual(1) & ") a (" & Combo1.ItemData(Combo1.ListIndex) & ")")
    End If
    If Reg_Actual(2) <> Combo2.ItemData(Combo2.ListIndex) Then
        Call Enviar_Bitacora(IdUser, "RADIOTERAPIA", "MODIFICAR", "Se modifico La tabla Tecnica1 cuya Id=" & IdTec1 & ", el campo TECNICO de (" & Reg_Actual(2) & ") a (" & Combo2.ItemData(Combo2.ListIndex) & ")")
    End If
    If Reg_Actual(3) <> Combo3.ItemData(Combo3.ListIndex) Then
        Call Enviar_Bitacora(IdUser, "RADIOTERAPIA", "MODIFICAR", "Se modifico La tabla Tecnica1 cuya Id=" & IdTec1 & ", el campo PROTOCOLO de (" & Reg_Actual(3) & ") a (" & Combo3.ItemData(Combo3.ListIndex) & ")")
    End If
Else
    Call Enviar_Bitacora(IdUser, "RADIOTERAPIA", "AGREGAR", "Se agrego un registro en la  tabla Tecnica1 cuya Id=" & C & ")")
End If

Exit Sub


WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open "c:\miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub
Private Sub BtnImprimir_Click()
'On Error GoTo WrtError
If Text1.Text = "" Then
    MsgBox "Seleccione al paciente", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If

'========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\TecnicoReporteRadioterapia.rpt"
    '.Connect = "DSN=CrReporte"
    .Connect = "Data Source=server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Tecnicas2.IdPaciente} = " & IdPacT & " AND {Tecnicas2.IdInforme} = " & IdInf
    '.SelectionFormula = "{Tecnicas2.IdPaciente} = " & IdPacT & " AND {Tecnicas2.IdInforme} = " & IdInf & " AND {Tecnicas2.IdL} = '" & IdL & "'"
    .WindowTitle = "Reporte de Docimetria" '
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With


End Sub

Private Sub BtnListaEspera_Click()
'command12
ModulO = 2
Cedul = ""
FrmListaEspera.Show modal, FrmPrincipal
If Cedul <> "" Then
TxtBuscar.Text = Cedul
Call BtnBuscar_Click
End If
End Sub

Private Sub BtnLlamar_Click()
'command13
Call Llamar
End Sub

Private Sub BtnParametrosSimulador_Click()
If IdPaciente <> "" Then
    FrmParametrosSimulacion.Show vbModal, FrmPrincipal
    'FrmEditorParametrosSimulacion.Show vbModal, FrmPrincipal
Else
    Msg = "Debe Seleccionar un Paciente"
    MsgBox Msg, vbExclamation + vbOKOnly, "Debe Seleccionar un paciente"
End If
End Sub

Private Sub BtnProtocolos_Click()
FrmProtocolos.Show vbModal
End Sub

Private Sub BtnReporteDosimetria_Click()
FrmDosimetria.Show 1, FrmPrincipal
End Sub

Private Sub BtnSiguiente_Click()
If Trim(IdPacT) = "" Then Exit Sub
If RsPacientes.RecordCount <> 0 Then
    If RsPacientes.EOF Then
        RsPacientes.MoveFirst
    Else
        RsPacientes.MoveNext
        If RsPacientes.EOF Then RsPacientes.MoveFirst
    End If
    Call Carga_De_Datos

    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If
End Sub

Private Sub Command11_Click()
IO = 1
Unload Me
End Sub

Sub LlenarGrid()

If IdPaciente <> "" Then
    
    CSql = "Select * From Tecnica Where IdPaciente = " & IdPaciente & " And IdLIdPac='" & IdLIdPac & "' And IdInforme='" & IdInf & "' And IdLIdInf='" & IdLIdInf & "' And Activo=1 ORDER BY Id"
    Set RsTecnica = CrearRS(CSql)
    
    ListView1.ListItems.Clear
    
    Do While Not RsTecnica.EOF
        With ListView1
            i = i + 1
           If Not IsNull(RsTecnica.Fields("dias").Value) Then .ListItems.Add , , RsTecnica.Fields("dias").Value Else .ListItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Tecnica").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Tecnica").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Energia").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Energia").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Prof").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Prof").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Fdia").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Fdia").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Fsem").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Fsem").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Ftot").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Ftot").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("DosisF").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("DosisF").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("DosisT").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("DosisT").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("Id").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("Id").Value Else .ListItems(i).ListSubItems.Add , , ""
           If Not IsNull(RsTecnica.Fields("IdL").Value) Then .ListItems(i).ListSubItems.Add , , RsTecnica.Fields("IdL").Value Else .ListItems(i).ListSubItems.Add , , ""
        End With
        RsTecnica.MoveNext
    Loop
                     
Else
    ListView1.ListItems.Clear
End If

If ListView1.ListItems.Count > 0 Then
    BtnImprimir.Enabled = True
Else
    BtnImprimir.Enabled = False
End If


End Sub
Sub Carga_De_Datos()

Dim RsRutaFotoEmpleados As New ADODB.Recordset
Dim RsMedTratante As New ADODB.Recordset 'Tabla Medico Tratante
Dim RutaFoto As String

    CSql = "Select * From Dat_Admin"
    Set RsRutaFotoEmpleados = CrearRS(CSql)
    RutaFoto = RsRutaFotoEmpleados.Fields("RutaFotos").Value
    RsRutaFotoEmpleados.Close
    
    If RsPacientes.RecordCount = 0 Then
        'msg = "LLego al Final del Registro"
        'MsgBox msg
        'If RsPacientes.RecordCount <> 0 Then RsPacientes.MoveFirst
        IdPaciente = ""
        IdPacT = ""
        IdLIdPac = ""
        IdLIdInf = ""
        IdUsuario = ""
        Camp = ""
        Exit Sub
    End If
    If Trim(RsPacientes.Fields("cedulap")) <> "" Then Text1.Text = RsPacientes.Fields("cedulap")
    If Trim(RsPacientes.Fields("Fecha_regP")) <> "" Then DtpFechaRegistro.Value = RsPacientes.Fields("Fecha_regP")
    If Trim(RsPacientes.Fields("Nombrep")) <> "" Then Text4.Text = RsPacientes.Fields("Nombrep")
    If Trim(RsPacientes.Fields("Apellidop")) <> "" Then Text3.Text = RsPacientes.Fields("Apellidop")
    If Trim(RsPacientes.Fields("Edadp")) <> "" Then Text6.Text = RsPacientes.Fields("Edadp")
    If Trim(RsPacientes.Fields("Historia")) <> "" Then Label12.Caption = RsPacientes.Fields("Historia")
    IdPaciente = RsPacientes.Fields("IdPaciente").Value
    IdUsuario = RsPacientes.Fields("IdUsuario").Value
    IdPacT = RsPacientes.Fields("IdPaciente").Value
    IdLIdPac = RsPacientes.Fields("IdL").Value
    IdLIdInf = RsPacientes.Fields("idLidPac").Value
   
    If Not IsNull(RsPacientes.Fields("foto").Value) Then
        If RsPacientes.Fields("foto").Value <> "" And Dir(Foto & "\" & RsPacientes.Fields("foto").Value) <> "" Then
            Image2.Picture = LoadPicture(Foto & "\" & RsPacientes.Fields("foto").Value)
            FotoP = RsPacientes.Fields("foto").Value
        Else
            Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
            FotoP = ""
        End If
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        FotoP = ""
    End If

    Me.Caption = "Area Técnica - Paciente: " & IdPaciente
    
    CSql = "SELECT * FROM medicos where [idmedico] = " & RsPacientes.Fields("Medico_Tratante") & " AND (Tipo=2 OR Tipo=3)"
    Set RsMedTratante = CrearRS(CSql)
    If RsMedTratante.EOF Then
        Msg = "Verifique el Medico Tratante"
        MsgBox Msg, vbOKOnly + vbInformation, "Medico Tratante"
        Else
        Text10.Text = Trim(RsMedTratante.Fields("nombre").Value) & " " & Trim(RsMedTratante.Fields("Apellido").Value)
    End If
    RsMedTratante.Close
                                    
    If RsPacientes.Fields("sexop") = 0 Then
        CboSexo.Text = "Masculino"
    Else
        CboSexo.Text = "Femenino"
    End If
    
    NoReg.Caption = "Total de Registros: " & RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
    DtpFechaInicio.Value = RsPacientes.Fields("Fecha_inicio")
    DtpFechaCulminacion.Value = RsPacientes.Fields("Fecha_culm")

    CSql = "Select * From Informe_Medico Where IdPaciente = " & IdPaciente & " And IdLIdPac='" & IdLIdPac & "' And Estado=1"
    Set RsInformeMed = CrearRS(CSql)

    Call carga_de_datos_tec
    Call Usuario_call
    LlenarGrid
End Sub

Sub carga_de_datos_tec()

IdLIdInf = IdLDefault

If RsInformeMed.EOF Then GoTo nohay

If RsInformeMed.Fields("Diagnotico").Value <> "" Then Text5.Text = Trim(RsInformeMed.Fields("Diagnotico").Value) Else Text5.Text = ""
If RsInformeMed.Fields("Tratamiento").Value <> "" Then Text7.Text = Trim(RsInformeMed.Fields("Tratamiento").Value) Else Text7.Text = ""
If RsInformeMed.Fields("Cuantas").Value <> "" Then TxtTAC.Text = Trim(RsInformeMed.Fields("Cuantas").Value) Else TxtTAC.Text = ""
If Not IsNull(RsInformeMed.Fields("Metas").Value) Then TxtMetas.Text = Trim(RsInformeMed.Fields("Metas").Value) Else TxtMetas.Text = ""

' obtiene el IdInforme del informe_medico y lo almacena en la variable ===>  IdInf
If Not IsNull(RsInformeMed.Fields("IdInforme").Value) Then IdInf = Trim(RsInformeMed.Fields("IdInforme").Value) Else IdInf = ""

IdLIdInf = Trim(RsInformeMed.Fields("IdL").Value)

TInformes.Caption = "Total de Informes: " & RsInformeMed.AbsolutePosition & " / " & RsInformeMed.RecordCount


Text5.ToolTipText = "Identificador del Informe = " & IdInf
CSql = "Select * From Tecnica1 Where IdPaciente='" & IdPacT & "' And IdLIdPac = '" & IdLIdPac & "' And IdInforme='" & IdInf & "' And IdLIdInf = '" & IdLIdInf & "' And Activo='1'"
Set RsTecnica1 = CrearRS(CSql)

If RsTecnica1.RecordCount > 0 Then

For T = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(T) = RsTecnica1.Fields("Tecnico") Then
        Combo1.ListIndex = T
        Reg_Actual(1) = RsTecnica1.Fields("Tecnico")
        Exit For
    End If
Next T
        
For T = 0 To Combo2.ListCount - 1
    If Combo2.ItemData(T) = RsTecnica1.Fields("Fisico") Then
        Combo2.ListIndex = T
        Reg_Actual(2) = RsTecnica1.Fields("Fisico")
        Exit For
    End If
Next T

For T = 0 To Combo3.ListCount - 1
  If Combo3.ItemData(T) = RsTecnica1.Fields("Protocolo") Then
  Combo3.ListIndex = T
  Reg_Actual(3) = RsTecnica1.Fields("Protocolo")
  Exit For
  End If
Next T

End If

LlenarGrid
Exit Sub

nohay:

'msg = "No hay Informe Médica asociada al paciente: " & Chr(13) & Chr(13) & Text3.Text & " " & Text4.Text & Chr(13) & "Se Mostraran los datos de Historia Médica en blanco"
'MsgBox msg, vbOKOnly, "No Tiene Informe Medico"

Text5.Text = ""
Text7.Text = ""
TxtTAC.Text = ""
TxtMetas.Text = ""
End Sub
Sub Usuario_call()
Dim CSql As String
Dim RsRegTecnicos As New ADODB.Recordset

CSql = "Select * From Tecnica1 Where IdPaciente = '" & IdPaciente & "' And Activo=1"
'RsRegTecnicos.Open CSql, Cnn, , , adCmdText
Set RsRegTecnicos = CrearRS(CSql)
     'IdTec1 = RsRegTecnicos.Fields("Id").Value
If RsRegTecnicos.RecordCount <> 0 Then
    Actualizar = 1
    
    IdTec1 = RsRegTecnicos.Fields("Id")
    Reg_Actual(0) = RsRegTecnicos.Fields("Descripcion1")
    
    'If RsRegTecnicos.Fields("Descripcion1") <> "" Then Text7.Text = RsRegTecnicos.Fields("Descripcion1") Else Text7.Text = ""
  
    For T = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(T) = RsRegTecnicos.Fields("Tecnico") Then
            Combo1.ListIndex = T
            Reg_Actual(1) = RsRegTecnicos.Fields("Tecnico")
            Exit For
        End If
    Next T
            
    For T = 0 To Combo2.ListCount - 1
        If Combo2.ItemData(T) = RsRegTecnicos.Fields("Fisico") Then
            Combo2.ListIndex = T
            Reg_Actual(2) = RsRegTecnicos.Fields("Fisico")
            Exit For
        End If
    Next T

    For T = 0 To Combo3.ListCount - 1
      If Combo3.ItemData(T) = RsRegTecnicos.Fields("Protocolo") Then
      Combo3.ListIndex = T
      Reg_Actual(3) = RsRegTecnicos.Fields("Protocolo")
      Exit For
      End If
    Next T
Else
    'msg = "Este paciente no posee registros tecnicos"
    'MsgBox msg, vbOKOnly, "No tiene registros"
   ' Text7.Text = ""
   LlenarGrid
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    IdTec1 = ""
End If
RsRegTecnicos.Close
Exit Sub
Actualizar = 0
End Sub
Sub Usuario()
Dim RsCargarUsuarios As New ADODB.Recordset
Dim RsCargarProtocolos As New ADODB.Recordset

'mmmmmmmm Carga Tecnicos mmmmmmmmm

CSql = "Select * From Empleados Where Cargo='3' And Activo='1' And (Status = '0' Or Status = '2')"
Set RsCargarUsuarios = CrearRS(CSql)
RsCargarUsuarios.MoveFirst
Combo1.Clear

Do While Not RsCargarUsuarios.EOF
    Combo1.AddItem RsCargarUsuarios.Fields("Nombre").Value & " " & RsCargarUsuarios.Fields("Apellido").Value
    Combo1.ItemData(Combo1.NewIndex) = RsCargarUsuarios.Fields("IdEmpleado").Value
    RsCargarUsuarios.MoveNext
Loop

inicial1 = Mid(Combo1.Text, 1, 1)
aa = InStr(1, Combo1.Text, " ", vbTextCompare)
bb = InStr(aa + 1, Combo1.Text, " ", vbTextCompare)
inicial2 = Mid(Combo1.Text, bb + 1, 1)
InicialTec = inicial1 & inicial2

'mmmmmmmm Carga Fisicos mmmmmmmmm

CSql = "Select * From Empleados Where Cargo='1' And Activo='1' And (Status = '0' Or Status = '2')"
Set RsCargarUsuarios = CrearRS(CSql)
RsCargarUsuarios.MoveFirst
Combo2.Clear

Do While Not RsCargarUsuarios.EOF
    Combo2.AddItem RsCargarUsuarios.Fields("Nombre").Value & " " & RsCargarUsuarios.Fields("Apellido").Value
    Combo2.ItemData(Combo2.NewIndex) = RsCargarUsuarios.Fields("IdEmpleado").Value
    RsCargarUsuarios.MoveNext
Loop

'mmmmmmmm Carga Protocolos mmmmmmmmm

CSql = "Select * From Protocolos Where Activo='1'"
Set RsCargarProtocolos = CrearRS(CSql)
RsCargarProtocolos.MoveFirst
Combo3.Clear

Do While Not RsCargarProtocolos.EOF
    Combo3.AddItem RsCargarProtocolos.Fields("Protocolo")
    Combo3.ItemData(Combo3.NewIndex) = RsCargarProtocolos.Fields("id")
RsCargarProtocolos.MoveNext
Loop
'Call Carga_De_Datos

End Sub

Sub Blanqueo()

Text5.Text = ""
Text7.Text = ""
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Combo3.ListIndex = -1
                    
End Sub


Private Sub BtnSiguiente1_Click()
If Trim(IdPacT) = "" Then Exit Sub

If RsInformeMed.RecordCount <> 0 Then
    If RsInformeMed.EOF Then
        RsInformeMed.MoveFirst
    Else
        RsInformeMed.MoveNext
        If RsInformeMed.EOF Then RsInformeMed.MoveFirst
    End If
    Call carga_de_datos_tec
Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If

End Sub

Private Sub Combo1_Click()
'IdTec1 = Combo1.ItemData(Combo1.ListIndex)
inicial1 = Mid(Combo1.Text, 1, 1)
aa = InStr(1, Combo1.Text, " ", vbTextCompare)
bb = InStr(aa + 1, Combo1.Text, " ", vbTextCompare)
inicial2 = Mid(Combo1.Text, bb + 1, 1)
InicialTec = inicial1 & inicial2
End Sub

Private Sub DataGrid1_DblClick()
Call Editar
End Sub

Private Sub ListView1_Click()
If ListView1.ListItems.Count > 0 Then
    Camp = ListView1.SelectedItem.ListSubItems(9).Text
    Camp2 = ListView1.SelectedItem.ListSubItems(10).Text
    DF = ListView1.SelectedItem.ListSubItems(7).Text
Else
    Camp = ""
    Camp2 = ""
    DF = ""
End If
End Sub


Private Sub ListView1_DblClick()
If ListView1.ListItems.Count > 0 Then
    Camp = ListView1.SelectedItem.ListSubItems(9).Text
    DF = ListView1.SelectedItem.ListSubItems(7).Text
    Editar
Else
    Camp = ""
    DF = ""
End If
End Sub


Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text7.SetFocus
        Case vbKeyUp
            Text6.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            Text7.SetFocus
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

Private Sub TxtBuscar_LostFocus()
If Trim(TxtBuscar.Text) = "" Then TxtBuscar.Text = "Busqueda"

End Sub

Private Sub Form_Load()
Centrar Me

ModulO = 2
If rs.State = 1 Then rs.Close

CSql = "Select * From Paciente where Historia like '%hoam%' or Historia like '%HOAM%' Order by IdPaciente"
Set RsPacientes = CrearRS(CSql)
NoReg.Caption = "Total de Registros: " & RsPacientes.RecordCount

SQL = ""
T = 0
IdPacT = ""
Call Usuario

End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Sub EnviarRegPendiente(ByVal IdNuevo2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tecnica1 WHERE Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Tecnica1 (["
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
RsRegPendiente.Fields("Modulo").Value = "Tratamiento TECNICA1"
RsRegPendiente.Fields("Tabla").Value = "Tecnica1"
RsRegPendiente.Fields("Condicional").Value = "Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' TECNICA TECNICA TECNICA TECNICA TECNICA TECNICA TECNICA TECNICA TECNICA TECNICA

Sub EnviarRegPendienteTec(ByVal IdNuevo2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tecnica WHERE Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Tecnica (["
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
RsRegPendiente.Fields("Modulo").Value = "Tratamiento TECNICA"
RsRegPendiente.Fields("Tabla").Value = "Tecnica"
RsRegPendiente.Fields("Condicional").Value = "Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


Private Sub Eliminar()

If IdPaciente = "" Then MsgBox "Debe seleccionar un paciente para eliminar datos del registro!", vbExclamation + vbOKOnly, "No puede borrar datos!": Exit Sub

If ListView1.ListItems.Count <= 0 Then
    MsgBox "No hay ningún registro para eliminar", vbInformation
    Exit Sub
End If
 
With ListView1
    If MsgBox("Se va a eliminar el registro : está seguro ", vbExclamation + vbYesNo, "Eliminar") = vbYes Then
        
        CSql = "Select * From Tecnica Where IdPaciente='" & IdPacT & "' And IdLIdPac='" & IdLIdPac & "' And Id='" & .SelectedItem.ListSubItems(9).Text & "' And IdL='" & .SelectedItem.ListSubItems(10).Text & "'"
        Set RsTecnica = CrearRS(CSql)
        RsTecnica.Fields("Activo") = 0
        
        ' Actualiza el recordset
        RsTecnica.Update
        EnviarRegPendienteTec .SelectedItem.ListSubItems(9).Text, .SelectedItem.ListSubItems(10).Text
        
        Call Enviar_Bitacora(IdUser, "RADIOTERAPIA", "BORRAR", " se Elimino el registro de la tabla TECNICA con el Id=" & RsTecnica.Fields("Id").Value)
        LlenarGrid
    End If
End With
End Sub

' agrega uno nuevo
Sub Agregar()
Dim RsTemp As New ADODB.Recordset
Dim CSql As String

If IdPaciente = "" Then MsgBox "Debe seleccionar un paciente para agregar datos al registro!", vbExclamation + vbOKOnly, "No puede agregar datos!": Exit Sub

If Replace(Combo1.Text, " ", "") = "" Then MsgBox "Seleccione un Técnico antes de agregar registros!", vbExclamation + vbOKOnly, "Faltan datos!": Combo1.SetFocus: Exit Sub
If Replace(Combo2.Text, " ", "") = "" Then MsgBox "Seleccione un Físico antes de agregar registros!", vbExclamation + vbOKOnly, "Faltan datos!": Combo2.SetFocus: Exit Sub
If Replace(Combo3.Text, " ", "") = "" Then MsgBox "Seleccione un Protocolo antes de agregar registros!", vbExclamation + vbOKOnly, "Faltan datos!": Combo3.SetFocus: Exit Sub
    
With FrmEdicionTecnico
    ACCION = AGREGAR_REGISTRO
    .DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    
    .Label2 = "Nuevo Reg."
    
    IdTecnico = Combo1.ItemData(Combo1.ListIndex)
    IdFisico = Combo2.ItemData(Combo2.ListIndex)
    IdProtocolo = Combo3.ItemData(Combo3.ListIndex)
    IdInf = RsInformeMed.Fields("IdInforme").Value
    
    .IdLIdPacT = IdLIdPac
    .IdLIdinfT = IdLIdInf
    
    NombreTecnico = Trim(Combo1.Text)
    NombreFisico = Trim(Combo2.Text)
    
    Set RsTemp = Nothing
    
    .Show vbModal
    LlenarGrid
End With
End Sub

Private Sub Editar() 'Abre el formulario para Editar el registro seleccionado
Dim RsTecnica1 As New ADODB.Recordset
If IdPaciente = "" Then MsgBox "Debe seleccionar un paciente para editar datos del registro!", vbExclamation + vbOKOnly, "No puede editar datos!": Exit Sub


Dim i As Integer
'If DataGrid1.Row = -1 Then MsgBox "No existen datos para editar!", vbExclamation + vbOKOnly, "Informacion": Exit Sub

If ListView1.ListItems.Count <= 0 Then MsgBox "No existen datos para editar!", vbExclamation + vbOKOnly, "Informacion": Exit Sub


CSql = "Select * From Tecnica where IdPaciente='" & IdPacT & "' And IdLIdPac='" & IdLIdPac & "' And Id='" & Camp & "' And IdLIdInf='" & IdLIdInf & "'"
Set RsTecnica = CrearRS(CSql)
If RsTecnica.RecordCount > 0 Then

With FrmEdicionTecnico
    ' obtiene el elemento seleccionado, el id
    ACCION = EDITAR_REGISTRO
   .BtnAgregar.Enabled = False
   .BtnGuardarActualizar.Enabled = True
   .BtnEliminar.Enabled = False
    .Frame1.Enabled = True
    
    .Label2 = RsTecnica.Fields("Id").Value
    .IdLIdInf = RsTecnica.Fields("IdL").Value
    .IdLIdinfT = IdLIdInf
    .IdLIdPacT = IdLIdPac
    
    ' llena los campos
    .Text1(7).Text = RsTecnica.Fields("Tecnica").Value
    .Text1(8).Text = RsTecnica.Fields("Energia").Value
    .Text1(9).Text = RsTecnica.Fields("Prof").Value
    .Text1(10).Text = RsTecnica.Fields("Fdia").Value
    .Text1(11).Text = RsTecnica.Fields("Fsem").Value
    .Text1(12).Text = RsTecnica.Fields("Ftot").Value
    .Text1(13).Text = RsTecnica.Fields("DosisF").Value
    
    IdTecn = RsTecnica.Fields("Id").Value
    Tecn = RsTecnica.Fields("Tecnica").Value
    ENER = RsTecnica.Fields("Energia").Value
    PROF = RsTecnica.Fields("Prof").Value
    Fdia = RsTecnica.Fields("Fdia").Value
    Fsem = RsTecnica.Fields("Fsem").Value
    Ftot = RsTecnica.Fields("Ftot").Value
    DosisF = RsTecnica.Fields("DosisF").Value
    DosisT = RsTecnica.Fields("DosisT").Value
    .Text1(14).Text = RsTecnica.Fields("DosisT").Value
    
    IdTecnico = Combo1.ItemData(Combo1.ListIndex)
    NombreTecnico = RsTecnica.Fields("nombreTecnico").Value
    IdFisico = Combo2.ItemData(Combo2.ListIndex)
    NombreFisico = RsTecnica.Fields("NombreFisico").Value
                 
    IdProtocolo = Combo3.ItemData(Combo3.ListIndex)
         
    .DTPicker1.Value = RsTecnica.Fields("Dias").Value
    
    .Show vbModal
    'DataGrid1.Refresh
    
    LlenarGrid
End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer3_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaRegistro.SetFocus
        Case vbKeyRight
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaInicio.SetFocus
        Case vbKeyLeft
            Text1.SetFocus
        Case vbKeyRight
            DtpFechaInicio.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaInicio_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaCulminacion.SetFocus
        Case vbKeyLeft
            DtpFechaRegistro.SetFocus
        Case vbKeyRight
            DtpFechaCulminacion.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaCulminacion_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaCulminacion.SetFocus
        Case vbKeyLeft
            DtpFechaInicio.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Text4.SetFocus
        Case vbKeyDown
            Text6.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboSexo.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyRight
            CboSexo.SetFocus
        Case vbKeyDown
            Text5.SetFocus
    End Select
End If
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo1.SetFocus
        Case vbKeyUp
            Text5.SetFocus
        Case vbKeyRight
            BtnGuardarDatosTecnicos.SetFocus
        Case vbKeyDown
            ListView1.SetFocus
    End Select
End If
End Sub

Private Sub CboSexo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyLeft
            Text6.SetFocus
        Case vbKeyRight
            Text10.SetFocus
        Case vbKeyDown
            Text5.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyUp
            DtpFechaInicio.SetFocus
        Case vbKeyLeft
            Text3.SetFocus
        Case vbKeyDown
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text5.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyLeft
            CboSexo.SetFocus
        Case vbKeyRight
            BtnListaEspera.SetFocus
        Case vbKeyDown
            Text5.SetFocus
    End Select
End If
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo2.SetFocus
        Case vbKeyUp
            BtnListaEspera.SetFocus
        Case vbKeyLeft
            Text5.SetFocus
        Case vbKeyDown
            Combo2.SetFocus
    End Select
End If
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo3.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyLeft
            Text5.SetFocus
        Case vbKeyDown
            Combo3.SetFocus
    End Select
End If
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGuardarDatosTecnicos.SetFocus
        Case vbKeyUp
            Combo2.SetFocus
        Case vbKeyLeft
            Text5.SetFocus
        Case vbKeyDown
            BtnGuardarDatosTecnicos.SetFocus
    End Select
End If
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmHistorialMedico 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial Médico"
   ClientHeight    =   9075
   ClientLeft      =   4020
   ClientTop       =   675
   ClientWidth     =   13200
   Icon            =   "Historia.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   8160
         Width           =   5415
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   420
            Left            =   120
            TabIndex        =   19
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido,Cédula de identidad o Historia"
            Top             =   260
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2280
            TabIndex        =   12
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "Historia.frx":1002
            PICN            =   "Historia.frx":101E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnListadoPaciente 
            Height          =   375
            Left            =   3480
            TabIndex        =   58
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            MICON           =   "Historia.frx":1283
            PICN            =   "Historia.frx":129F
            PICH            =   "Historia.frx":1528
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
         Left            =   9480
         Top             =   480
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   5640
         TabIndex        =   17
         Top             =   8160
         Width           =   7215
         Begin ChamaleonButton.ChameleonBtn BtnAnterior1 
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
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
            MICON           =   "Historia.frx":17BB
            PICN            =   "Historia.frx":17D7
            PICH            =   "Historia.frx":1A6C
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
            Left            =   3360
            TabIndex        =   15
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
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
            MICON           =   "Historia.frx":1CC8
            PICN            =   "Historia.frx":1CE4
            PICH            =   "Historia.frx":1F7A
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
            TabIndex        =   13
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Historia.frx":21D9
            PICN            =   "Historia.frx":21F5
            PICH            =   "Historia.frx":231A
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
            Left            =   6120
            TabIndex        =   16
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
            MICON           =   "Historia.frx":25AA
            PICN            =   "Historia.frx":25C6
            PICH            =   "Historia.frx":278F
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9960
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Informe medico"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Informe Médico"
         Height          =   5055
         Left            =   120
         TabIndex        =   42
         Top             =   3120
         Width           =   12735
         Begin VB.TextBox TxtAntecedentesPersonal 
            Height          =   1215
            Left            =   4320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   3720
            Width           =   4095
         End
         Begin MSComCtl2.DTPicker DTPickerFecha 
            Height          =   375
            Left            =   8520
            TabIndex        =   9
            Top             =   3720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   40121
         End
         Begin VB.TextBox TxtMotivoConsulta 
            Height          =   1215
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox TxtExamenFisico 
            Height          =   1215
            Left            =   4320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox TxtEnfermedadActual 
            Height          =   1215
            Left            =   4320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox TxtAnatomiaPatologica 
            Height          =   1215
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox TxtDiagnostico 
            Height          =   1215
            Left            =   8520
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox TxtTratamiento 
            Height          =   1215
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3720
            Width           =   4095
         End
         Begin VB.TextBox TxtAntecedentesFamiliares 
            Height          =   1215
            Left            =   8520
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   600
            Width           =   4095
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Height          =   735
            Left            =   11040
            TabIndex        =   43
            Top             =   4200
            Width           =   1575
            Begin ChamaleonButton.ChameleonBtn BtnAnterior 
               Height          =   375
               Left            =   120
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
               MICON           =   "Historia.frx":29C4
               PICN            =   "Historia.frx":29E0
               PICH            =   "Historia.frx":2C75
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
               Left            =   840
               TabIndex        =   11
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
               MICON           =   "Historia.frx":2ED1
               PICN            =   "Historia.frx":2EED
               PICH            =   "Historia.frx":3183
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
         Begin VB.Label Noreg2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informe: 0 / 0"
            Height          =   195
            Left            =   8520
            TabIndex        =   57
            Top             =   4680
            Width           =   960
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Antecedentes Personales"
            Height          =   195
            Left            =   4320
            TabIndex        =   55
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   8520
            TabIndex        =   51
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Antecedentes Familiares"
            Height          =   195
            Left            =   8520
            TabIndex        =   50
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "En&fermedad Actual"
            Height          =   195
            Left            =   4320
            TabIndex        =   49
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Motivo de la Consulta"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico"
            Height          =   195
            Left            =   4320
            TabIndex        =   47
            Top             =   1920
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Anatomía &Patológica "
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   1920
            Width           =   1530
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dia&gnóstico"
            Height          =   195
            Left            =   8520
            TabIndex        =   45
            Top             =   1920
            Width           =   840
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Tratamiento"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   3480
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Enabled         =   0   'False
         Height          =   2775
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12735
         Begin VB.TextBox TxtMedicoRemitente 
            Height          =   375
            Left            =   1800
            TabIndex        =   27
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox TxtMedicoTratante 
            Height          =   375
            Left            =   5880
            TabIndex        =   26
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox TxtEdad 
            Height          =   375
            Left            =   4080
            TabIndex        =   25
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TxtApellido 
            Height          =   375
            Left            =   1800
            TabIndex        =   24
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   5880
            TabIndex        =   23
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox TxtCedula 
            Height          =   375
            Left            =   1800
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox CboSexo 
            Height          =   315
            ItemData        =   "Historia.frx":33E2
            Left            =   5880
            List            =   "Historia.frx":33EC
            TabIndex        =   21
            Top             =   1320
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   40114
         End
         Begin MSComCtl2.DTPicker DtpFechaCulminacion 
            Height          =   375
            Left            =   5880
            TabIndex        =   29
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   40114
         End
         Begin MSComCtl2.DTPicker DtpFechaNacimiento 
            Height          =   375
            Left            =   1800
            TabIndex        =   52
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   40114
         End
         Begin MSComCtl2.DTPicker DtpFechaRegistro 
            Height          =   375
            Left            =   8880
            TabIndex        =   53
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   40114
         End
         Begin VB.Label NoReg 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro: 0 /0"
            Height          =   195
            Left            =   10560
            TabIndex        =   56
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   2040
            Left            =   10440
            Picture         =   "Historia.frx":3405
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2205
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Historia Médica:"
            Height          =   195
            Left            =   4635
            TabIndex        =   54
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label LblHistoriaMedica 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
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
            Left            =   5880
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Sexo:"
            Height          =   195
            Left            =   5400
            TabIndex        =   40
            Top             =   1380
            Width           =   405
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Edad:"
            Height          =   195
            Left            =   3600
            TabIndex        =   39
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Fecha de Nacimiento:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1410
            Width           =   1560
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Medico Remitente:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   2370
            Width           =   1335
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de C&ulminación:"
            Height          =   195
            Left            =   4155
            TabIndex        =   36
            Top             =   1890
            Width           =   1620
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de &Inicio:"
            Height          =   195
            Left            =   75
            TabIndex        =   35
            Top             =   1890
            Width           =   1140
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Medico Tratante:"
            Height          =   195
            Left            =   4515
            TabIndex        =   34
            Top             =   2370
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "A&pellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nombre(s):"
            Height          =   195
            Left            =   4920
            TabIndex        =   32
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de &Registro:"
            Height          =   195
            Left            =   7440
            TabIndex        =   31
            Top             =   1890
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Cédula de Identidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   450
            Width           =   1470
         End
      End
   End
End
Attribute VB_Name = "FrmHistorialMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargaPaciente As New ADODB.Recordset 'tabla Registro Historico
Dim RsCargaInforme As New ADODB.Recordset 'tabla Registro Historico
Dim CSql As String
Public opcion As Integer
Dim IdInfor As String

Private Sub BtnAnterior_Click()
If RsCargaInforme.RecordCount <> 0 Then
    RsCargaInforme.MovePrevious
    If RsCargaInforme.BOF Then RsCargaInforme.MoveLast
    Call carga_de_datos1
    Else
    MsgBox "No tiene INFORMES MEDICOS cargados!", vbExclamation + vbOKOnly, "No hay datos"
    Noreg2.Caption = "Informe 0 / 0"
End If
End Sub

Private Sub BtnAnterior1_Click()
If RsCargaPaciente.RecordCount <> 0 Then
    RsCargaPaciente.MovePrevious
    If RsCargaPaciente.BOF Then RsCargaPaciente.MoveLast
    Call Carga_De_Datos
    Else
    MsgBox "No hay datos cargados, Inicia una nueva busqueda!", vbExclamation + vbOKOnly, "No hay datos"
    Noreg2.Caption = "Informe 0 / 0"
End If
End Sub

Private Sub BtnAyuda_Click()
On Error Resume Next
Ayuda = 2
FormAyuda.Show
End Sub

Public Sub BtnBuscar_Click()
Blanqueo
If Trim(TxtBuscar) = "" Or Trim(TxtBuscar) = "Busqueda" Then
    CSql = "Select * From Paciente"
Else
    CSql = "Select * From Paciente Where CedulaP = " & Val(TxtBuscar.Text) & " OR nombrep like '%" & TxtBuscar.Text & "%' OR apellidop like '%" & TxtBuscar.Text & "%' or Historia = '" & UCase(TxtBuscar.Text) & "'"
End If

Set RsCargaPaciente = CrearRS(CSql)
If RsCargaPaciente.EOF Then
    MsgBox "No Existe el registro", vbExclamation + vbOKOnly, "No hay datos"
    NoReg.Caption = "Registro 0 / 0"
    Noreg2.Caption = "Informe 0 / 0"
    Blanqueo
    Exit Sub
End If

Call Carga_De_Datos

CSql = "select * from informe_medico where idpaciente = " & IdPac1 & " and estado = 1 order by fecha"
Set RsCargaInforme = CrearRS(CSql)

If Not RsCargaInforme.RecordCount <> 0 Then BtnSiguiente.Enabled = False: BtnAnterior.Enabled = False: Noreg2.Caption = "Informe 0 / 0": Exit Sub
Call carga_de_datos1

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()


If TxtCedula.Text <> "" And IdPac1 <> "" Then

''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\InformeMedicoN.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{Informe_Med.IdInforme} = " & IdInfor & " And {Informe_Med.IdL} = '" & IdLIdInf & "'"
        .WindowTitle = "Reporte Informe Medico No. " & IdInfor
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


End Sub


Private Sub BtnListadoPaciente_Click()
On Error Resume Next
Tipo = "HistorialPaciente"
FrmListadoPaciente.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnSiguiente_Click()
If RsCargaInforme.RecordCount <> 0 Then
    RsCargaInforme.MoveNext
    If RsCargaInforme.EOF Then RsCargaInforme.MoveFirst
    Call carga_de_datos1
    Else
    MsgBox "No tiene INFORMES MEDICOS cargados!", vbExclamation + vbOKOnly, "No hay datos"
    Noreg2.Caption = "Informe 0 / 0"
End If
End Sub

Private Sub BtnSiguiente2_Click()
If RsCargaPaciente.RecordCount <> 0 Then
    RsCargaPaciente.MoveNext
    If RsCargaPaciente.EOF Then RsCargaPaciente.MoveFirst
    Call Carga_De_Datos
    Else
    MsgBox "No hay datos cargados, Inicia una nueva busqueda!", vbExclamation + vbOKOnly, "No hay datos"
    Noreg2.Caption = "Informe 0 / 0"
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()

Centrar Me

CSql = "Select * From Paciente Order by IdPaciente"
Set RsCargaPaciente = CrearRS(CSql)

opcion = 1
RsCargaPaciente.MoveFirst

'Carga_De_Datos
 
End Sub

Sub Carga_De_Datos()

Blanqueo

If Trim(RsCargaPaciente.Fields("cedulap")) <> "" Then TxtCedula.Text = RsCargaPaciente.Fields("cedulap")
If Trim(RsCargaPaciente.Fields("Fecha_regp")) <> "" Then DtpFechaRegistro.Value = RsCargaPaciente.Fields("Fecha_regp")
If Trim(RsCargaPaciente.Fields("Nombrep")) <> "" Then TxtNombre.Text = RsCargaPaciente.Fields("Nombrep")
If Trim(RsCargaPaciente.Fields("Apellidop")) <> "" Then TxtApellido.Text = RsCargaPaciente.Fields("Apellidop")
If Trim(RsCargaPaciente.Fields("Fecha_nacimientop")) <> "" Then DtpFechaNacimiento.Value = RsCargaPaciente.Fields("Fecha_nacimientop")
If Trim(RsCargaPaciente.Fields("Edadp")) <> "" Then TxtEdad.Text = RsCargaPaciente.Fields("Edadp")
If Trim(RsCargaPaciente.Fields("Historia")) <> "" Then LblHistoriaMedica.Caption = RsCargaPaciente.Fields("Historia")
IdPac1 = RsCargaPaciente.Fields("idpaciente")
Me.Caption = "Historial Medico - Paciente: " & IdPac1

If RsCargaPaciente.Fields("foto") <> "" Then
    If Len(Dir(Foto & "\" & RsCargaPaciente.Fields("foto"))) > 0 Then
        Image2.Picture = LoadPicture(Foto & "\" & RsCargaPaciente.Fields("foto"))
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
Else
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
End If
    
Dim RsTemp As New ADODB.Recordset 'Tabla Medico Tratante
CSql = "SELECT * FROM medicos where [idmedico] = " & RsCargaPaciente.Fields("Medico_Tratante")
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount <= 0 Then
    MsgBox "Verifique el Medico Tratante", vbExclamation + vbOKOnly, "Medico Tratante"
    Else
    TxtMedicoTratante.Text = RsTemp.Fields("nombre")
End If

CSql = "SELECT * FROM medicos where [idmedico] = " & RsCargaPaciente.Fields("Medico_Remitente")
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount <= 0 Then
    MsgBox "Verifique Medico Remitente", vbExclamation + vbOKOnly, "Medico Remitente"
    Else
    TxtMedicoRemitente.Text = RsTemp.Fields("nombre")
End If
RsTemp.Close
                                    
If RsCargaPaciente.Fields("sexop") = 0 Then CboSexo.Text = "Masculino" Else CboSexo.Text = "Femenino"

DtpFechaInicio.Value = RsCargaPaciente.Fields("Fecha_inicio")
DtpFechaCulminacion.Value = RsCargaPaciente.Fields("Fecha_culm")

NoReg.Caption = "Registro " & RsCargaPaciente.AbsolutePosition & " / " & RsCargaPaciente.RecordCount

CSql = "select * from informe_medico where idpaciente = " & IdPac1 & " and estado = 1 order by fecha"
Set RsCargaInforme = CrearRS(CSql)

If RsCargaInforme.RecordCount = 0 Then
    BtnSiguiente.Enabled = False
    BtnAnterior.Enabled = False
    BtnImprimir.Enabled = False
    IdInfor = ""
    Noreg2.Caption = "Informe 0 / 0"
    Exit Sub
End If
Call carga_de_datos1

End Sub
Sub carga_de_datos1()

    If RsCargaInforme.Fields("Antecedente_flia") <> "" Then TxtAntecedentesFamiliares.Text = RsCargaInforme.Fields("Antecedente_flia") Else TxtAntecedentesFamiliares.Text = ""
    If RsCargaInforme.Fields("Enfermedad_act") <> "" Then TxtEnfermedadActual.Text = RsCargaInforme.Fields("Enfermedad_act") Else TxtEnfermedadActual.Text = ""
    If RsCargaInforme.Fields("Fecha") <> "" Then DTPickerFecha.Value = RsCargaInforme.Fields("Fecha") Else DTPickerFecha.Value = Now
    If RsCargaInforme.Fields("Motivo_con") <> "" Then TxtMotivoConsulta.Text = RsCargaInforme.Fields("Motivo_con") Else TxtMotivoConsulta.Text = ""
    If RsCargaInforme.Fields("Anatomia_patol") <> "" Then TxtAnatomiaPatologica.Text = RsCargaInforme.Fields("Anatomia_patol") Else TxtAnatomiaPatologica.Text = ""
    If RsCargaInforme.Fields("Examen_Fis") <> "" Then TxtExamenFisico.Text = RsCargaInforme.Fields("Examen_Fis") Else TxtExamenFisico.Text = ""
    If RsCargaInforme.Fields("Diagnotico") <> "" Then TxtDiagnostico.Text = RsCargaInforme.Fields("Diagnotico") Else TxtDiagnostico.Text = ""
    If RsCargaInforme.Fields("Tratamiento") <> "" Then TxtTratamiento.Text = RsCargaInforme.Fields("Tratamiento") Else TxtTratamiento.Text = ""
    IdInfor = RsCargaInforme.Fields("IdInforme")
    
'    If RsCargaInforme.Fields("Ante_Personal") <> "" Then TxtAntecedentesPersonal.Text = RsCargaInforme.Fields("Ante_Personal") Else TxtAntecedentesPersonal.Text = ""
    BtnSiguiente.Enabled = True
    BtnAnterior.Enabled = True
    BtnImprimir.Enabled = True
    Noreg2.Caption = "Informe: " & RsCargaInforme.AbsolutePosition & " / " & RsCargaInforme.RecordCount

Exit Sub

nohay:
    TxtAnatomiaPatologica.Text = ""
    TxtAntecedentesFamiliares.Text = ""
    TxtAntecedentesPersonal.Text = ""
    TxtMotivoConsulta.Text = ""
    TxtEnfermedadActual.Text = ""
    TxtExamenFisico.Text = ""
    TxtTratamiento.Text = ""
    TxtDiagnostico.Text = ""
    Msg = "No hay Informe Médico asociado al paciente: " & Chr(13) & Chr(13) & TxtNombre.Text & " " & TxtApellido.Text & Chr(13) & "Se Mostraran los datos de Historia Médica en blanco"
    MsgBox Msg, vbExclamation + vbOKOnly, "No Tiene Informe Medico"

End Sub

Sub carga_lista_medicost()
Dim RsCargaListaMedicoT As New ADODB.Recordset

CSql = "SELECT * FROM medicos_t"
Set RsCargaListaMedicoT = CrearRS(CSql)

RsCargaListaMedicoT.MoveFirst
Do While Not RsCargaListaMedicoT.EOF
    RsCargaListaMedicoT.MoveNext
Loop

End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Sub Blanqueo()
IdPac1 = ""
CboSexo.ListIndex = -1
Image2.Picture = LoadPicture()
TxtCedula.Text = ""
TxtApellido.Text = ""
TxtNombre.Text = ""
TxtApellido.Text = ""
TxtEdad.Text = ""
TxtMedicoTratante.Text = ""
TxtMedicoRemitente.Text = ""
TxtMotivoConsulta.Text = ""
TxtEnfermedadActual.Text = ""
TxtAntecedentesFamiliares.Text = ""
TxtAnatomiaPatologica.Text = ""
TxtExamenFisico.Text = ""
TxtDiagnostico.Text = ""
TxtTratamiento.Text = ""
TxtAntecedentesPersonal.Text = ""
LblHistoriaMedica.Caption = ""
DtpFechaRegistro.Value = Now
DtpFechaInicio.Value = Now
DtpFechaCulminacion.Value = Now
DtpFechaNacimiento.Value = Now
DTPickerFecha.Value = Now
End Sub

Private Sub TxtCedula_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TxtEdad_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text7_Click()

End Sub

Private Sub TxtBuscar_GotFocus()
'If Text5.Text = "Busqueda" Then Text5.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
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

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub

End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub CboSexo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtEdad.SetFocus
        Case vbKeyUp
            DtpFechaNacimiento.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            TxtEdad.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaNacimiento_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboSexo.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyRight
            DtpFechaCulminacion.SetFocus
        Case vbKeyDown
            CboSexo.SetFocus
    End Select
End If
End Sub

Private Sub DTPickerFecha_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnNuevo.SetFocus
        Case vbKeyUp
            TxtDiagnostico.SetFocus
        Case vbKeyLeft
            TxtAntecedentesPersonal.SetFocus
        Case vbKeyRight
            BtnAnterior.SetFocus
        Case vbKeyDown
            BtnNuevo.SetFocus
    End Select
End If
End Sub

Private Sub TxtAnatomiaPatologica_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtExamenFisico.SetFocus
        Case vbKeyUp
            TxtMotivoConsulta.SetFocus
        Case vbKeyRight
            TxtExamenFisico.SetFocus
        Case vbKeyDown
            TxtTratamiento.SetFocus
    End Select
End If
End Sub

Private Sub TxtAntecedentesFamiliares_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtAnatomiaPatologica.SetFocus
        Case vbKeyUp
            If TxtMedicoTratante.Enabled = True Then TxtMedicoTratante.SetFocus
        Case vbKeyLeft
            TxtEnfermedadActual.SetFocus
        Case vbKeyDown
            TxtDiagnostico.SetFocus
    End Select
End If
End Sub

Private Sub TxtAntecedentesPersonal_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPickerFecha.SetFocus
        Case vbKeyUp
            TxtExamenFisico.SetFocus
        Case vbKeyLeft
            TxtTratamiento.SetFocus
        Case vbKeyRight
            DTPickerFecha.SetFocus
        Case vbKeyDown
            BtnNuevo.SetFocus
    End Select
End If
End Sub

Private Sub TxtApellido_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNombre.SetFocus
        Case vbKeyUp
            TxtCedula.SetFocus
        Case vbKeyRight
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            TxtNombre.SetFocus
    End Select
End If
End Sub

Private Sub TxtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        Case vbKeyUp
            TxtTratamiento.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
    End Select
End If
End Sub

Private Sub TxtDiagnostico_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtTratamiento.SetFocus
        Case vbKeyUp
            TxtAntecedentesFamiliares.SetFocus
        Case vbKeyLeft
            TxtExamenFisico.SetFocus
        Case vbKeyDown
            DTPickerFecha.SetFocus
    End Select
End If
End Sub

Private Sub TxtEdad_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtMedicoRemitente.SetFocus
        Case vbKeyUp
            CboSexo.SetFocus
        Case vbKeyDown
            TxtMedicoRemitente.SetFocus
    End Select
End If
End Sub

Private Sub TxtEnfermedadActual_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtAntecedentesFamiliares.SetFocus
        Case vbKeyUp
            If TxtMedicoTratante.Enabled = True Then TxtMedicoTratante.SetFocus
        Case vbKeyLeft
            TxtMotivoConsulta.SetFocus
        Case vbKeyRight
            TxtAntecedentesFamiliares.SetFocus
        Case vbKeyDown
            TxtExamenFisico.SetFocus
    End Select
End If
End Sub

Private Sub TxtExamenFisico_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDiagnostico.SetFocus
        Case vbKeyUp
            TxtEnfermedadActual.SetFocus
        Case vbKeyLeft
            TxtAnatomiaPatologica.SetFocus
        Case vbKeyRight
            TxtDiagnostico.SetFocus
        Case vbKeyDown
            TxtAntecedentesPersonal.SetFocus
    End Select
End If
End Sub

Private Sub TxtMedicoRemitente_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtMedicoTratante.SetFocus
        Case vbKeyUp
            TxtEdad.SetFocus
        Case vbKeyDown
            TxtMedicoTratante.SetFocus
    End Select
End If
End Sub

Private Sub TxtMedicoTratante_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtMotivoConsulta.SetFocus
        Case vbKeyUp
            TxtEdad.SetFocus
        Case vbKeyDown
            TxtMotivoConsulta.SetFocus
    End Select
End If
End Sub

Private Sub TxtMotivoConsulta_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtEnfermedadActual.SetFocus
        Case vbKeyUp
            If TxtMedicoTratante.Enabled = True Then TxtMedicoTratante.SetFocus
        Case vbKeyRight
            TxtEnfermedadActual.SetFocus
        Case vbKeyDown
            TxtAnatomiaPatologica.SetFocus
    End Select
End If
End Sub

Private Sub TxtNombre_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaNacimiento.SetFocus
        Case vbKeyUp
            TxtApellido.SetFocus
        Case vbKeyRight
            DtpFechaInicio.SetFocus
        Case vbKeyDown
            DtpFechaNacimiento.SetFocus
    End Select
End If
End Sub

Private Sub TxtTratamiento_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtAntecedentesPersonal.SetFocus
        Case vbKeyUp
            TxtAnatomiaPatologica.SetFocus
        Case vbKeyRight
            TxtAntecedentesPersonal.SetFocus
        Case vbKeyDown
            TxtBuscar.SetFocus
    End Select
End If
End Sub


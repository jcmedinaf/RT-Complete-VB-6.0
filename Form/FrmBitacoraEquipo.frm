VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmBitacoraEquipo 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de equipos"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13950
   Icon            =   "FrmBitacoraEquipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmOpc 
      BackColor       =   &H00EAEFEF&
      Height          =   9855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   13455
      Begin VB.Frame Frame11 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   277
         Top             =   9000
         Width           =   13215
         Begin ChamaleonButton.ChameleonBtn BtnGuardar 
            Height          =   375
            Left            =   7200
            TabIndex        =   278
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
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
            MICON           =   "FrmBitacoraEquipo.frx":1002
            PICN            =   "FrmBitacoraEquipo.frx":101E
            PICH            =   "FrmBitacoraEquipo.frx":12AD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnLimpiarCampos 
            Height          =   375
            Left            =   4920
            TabIndex        =   279
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Limpiar Campos"
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
            MICON           =   "FrmBitacoraEquipo.frx":16EE
            PICN            =   "FrmBitacoraEquipo.frx":170A
            PICH            =   "FrmBitacoraEquipo.frx":196A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrarFrm 
            Height          =   375
            Left            =   11880
            TabIndex        =   280
            ToolTipText     =   "Cerrar "
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmBitacoraEquipo.frx":1BEE
            PICN            =   "FrmBitacoraEquipo.frx":1C0A
            PICH            =   "FrmBitacoraEquipo.frx":1DD3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnTecnicos 
            Height          =   375
            Left            =   2640
            TabIndex        =   307
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Tecnicos"
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
            MICON           =   "FrmBitacoraEquipo.frx":2008
            PICN            =   "FrmBitacoraEquipo.frx":2024
            PICH            =   "FrmBitacoraEquipo.frx":2284
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
         Height          =   4095
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   4575
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Datos del Equipo"
            Height          =   2175
            Left            =   240
            TabIndex        =   75
            Top             =   840
            Visible         =   0   'False
            Width           =   4095
            Begin VB.TextBox TxtNombre 
               Height          =   285
               Left            =   1080
               TabIndex        =   78
               Top             =   360
               Width           =   2775
            End
            Begin VB.TextBox TxtModelo 
               Height          =   285
               Left            =   1080
               TabIndex        =   77
               Top             =   720
               Width           =   2775
            End
            Begin VB.TextBox TxtSerial 
               Height          =   285
               Left            =   1080
               TabIndex        =   76
               Top             =   1080
               Width           =   2775
            End
            Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
               Height          =   375
               Left            =   1560
               TabIndex        =   79
               ToolTipText     =   "Guardar / Actualizar"
               Top             =   1680
               Width           =   1095
               _ExtentX        =   1931
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
               MICON           =   "FrmBitacoraEquipo.frx":2508
               PICN            =   "FrmBitacoraEquipo.frx":2524
               PICH            =   "FrmBitacoraEquipo.frx":27B3
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
               Left            =   2880
               TabIndex        =   80
               ToolTipText     =   "Cerrar "
               Top             =   1680
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
               MICON           =   "FrmBitacoraEquipo.frx":2BF4
               PICN            =   "FrmBitacoraEquipo.frx":2C10
               PICH            =   "FrmBitacoraEquipo.frx":2E61
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
               BackColor       =   &H00EAEFEF&
               Caption         =   "Nombre:"
               Height          =   195
               Left            =   360
               TabIndex        =   83
               Top             =   360
               Width           =   600
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00EAEFEF&
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   360
               TabIndex        =   82
               Top             =   720
               Width           =   570
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00EAEFEF&
               Caption         =   "Serial:"
               Height          =   195
               Left            =   360
               TabIndex        =   81
               Top             =   1080
               Width           =   435
            End
         End
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   2775
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4895
            Object.Width           =   4305
            Object.Height          =   2745
            BackColor       =   15396847
            ScrollBar       =   1
            MarqueeStyle    =   2
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregar 
            Height          =   375
            Left            =   2280
            TabIndex        =   85
            ToolTipText     =   "Agregar"
            Top             =   3600
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
            MICON           =   "FrmBitacoraEquipo.frx":30B2
            PICN            =   "FrmBitacoraEquipo.frx":30CE
            PICH            =   "FrmBitacoraEquipo.frx":34E9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnModificar 
            Height          =   375
            Left            =   3360
            TabIndex        =   86
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Modificar"
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
            MICON           =   "FrmBitacoraEquipo.frx":3904
            PICN            =   "FrmBitacoraEquipo.frx":3920
            PICH            =   "FrmBitacoraEquipo.frx":3D3E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equipos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   240
            TabIndex        =   299
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Height          =   8775
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   8535
         Begin VB.Frame Frame3 
            BackColor       =   &H00EAEFEF&
            Height          =   7335
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   8295
            Begin VB.ComboBox CboTecnicos 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   308
               Top             =   960
               Width           =   2175
            End
            Begin VB.Timer Timer1 
               Interval        =   500
               Left            =   4320
               Top             =   240
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   30
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   38
               Top             =   5730
               Width           =   855
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   29
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   37
               Top             =   5370
               Width           =   855
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   28
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   36
               Top             =   4965
               Width           =   855
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   27
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   35
               Top             =   4605
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   26
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   34
               Top             =   4245
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   25
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   33
               Top             =   3885
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   24
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   32
               Top             =   3525
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   23
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   31
               Top             =   3165
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   22
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   30
               Top             =   2805
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   21
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   29
               Top             =   2400
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   20
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   28
               Top             =   1965
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   19
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   27
               Top             =   1605
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   18
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   26
               Top             =   1245
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   17
               Left            =   6840
               MaxLength       =   10
               TabIndex        =   25
               Top             =   885
               Width           =   1215
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   16
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   24
               Top             =   6840
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   15
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   23
               Top             =   6480
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   14
               Left            =   2160
               MaxLength       =   10
               TabIndex        =   22
               Top             =   6120
               Width           =   1815
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   13
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   21
               Top             =   5685
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   12
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   20
               Top             =   5325
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   11
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   19
               Top             =   4965
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   10
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   18
               Top             =   4605
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   9
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   17
               Top             =   4245
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   8
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   16
               Top             =   3885
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   7
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   15
               Top             =   3480
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   6
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   14
               Top             =   3120
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   5
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   13
               Top             =   2760
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   4
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   12
               Top             =   2400
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   3
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   11
               Top             =   2040
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   2
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   10
               Top             =   1680
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   1
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   9
               Top             =   1320
               Width           =   2175
            End
            Begin VB.TextBox TxtCampos 
               Height          =   285
               Index           =   0
               Left            =   360
               MaxLength       =   10
               TabIndex        =   8
               Top             =   600
               Visible         =   0   'False
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   1800
               TabIndex        =   39
               Top             =   300
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Format          =   51642369
               CurrentDate     =   40345
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Presión de agua Externa del Equipo"
               Height          =   195
               Index           =   30
               Left            =   4560
               TabIndex        =   72
               Top             =   5775
               Width           =   2535
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Presión de agua Interna del Equipo"
               Height          =   195
               Index           =   29
               Left            =   4560
               TabIndex        =   71
               Top             =   5415
               Width           =   2490
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Chequeo del Detector Desumificador"
               Height          =   195
               Index           =   28
               Left            =   4560
               TabIndex        =   70
               Top             =   5040
               Width           =   2610
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Presión Barometrica del Bunker"
               Height          =   195
               Index           =   27
               Left            =   4560
               TabIndex        =   69
               Top             =   4650
               Width           =   2220
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% de Humedad Del Bunker"
               Height          =   195
               Index           =   26
               Left            =   4560
               TabIndex        =   68
               Top             =   4290
               Width           =   1920
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Temperatura del Bunker"
               Height          =   195
               Index           =   25
               Left            =   4560
               TabIndex        =   67
               Top             =   3885
               Width           =   1710
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Beam Monitoring Device"
               Height          =   195
               Index           =   24
               Left            =   4560
               TabIndex        =   66
               Top             =   3525
               Width           =   1740
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mecanical Counter"
               Height          =   195
               Index           =   23
               Left            =   4560
               TabIndex        =   65
               Top             =   3165
               Width           =   1335
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Arc Test"
               Height          =   195
               Index           =   22
               Left            =   4560
               TabIndex        =   64
               Top             =   2805
               Width           =   600
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time and Integrator"
               Height          =   195
               Index           =   21
               Left            =   4560
               TabIndex        =   63
               Top             =   2445
               Width           =   1380
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Beam Current"
               Height          =   195
               Index           =   20
               Left            =   4560
               TabIndex        =   62
               Top             =   2010
               Width           =   960
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PFN Volts"
               Height          =   195
               Index           =   19
               Left            =   4560
               TabIndex        =   61
               Top             =   1650
               Width           =   705
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Power Suply Current"
               Height          =   195
               Index           =   18
               Left            =   4560
               TabIndex        =   60
               Top             =   1290
               Width           =   1440
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vacuum Run"
               Height          =   195
               Index           =   17
               Left            =   4560
               TabIndex        =   59
               Top             =   960
               Width           =   930
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gun fil Run"
               Height          =   195
               Index           =   16
               Left            =   360
               TabIndex        =   58
               Top             =   6885
               Width           =   795
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mag Fil V Run"
               Height          =   195
               Index           =   15
               Left            =   360
               TabIndex        =   57
               Top             =   6525
               Width           =   1005
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Monitor Units Per Minute"
               Height          =   195
               Index           =   14
               Left            =   360
               TabIndex        =   56
               Top             =   6165
               Width           =   1740
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vacuum On"
               Height          =   195
               Index           =   13
               Left            =   360
               TabIndex        =   55
               Top             =   5730
               Width           =   840
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gun fil On"
               Height          =   195
               Index           =   12
               Left            =   360
               TabIndex        =   54
               Top             =   5370
               Width           =   705
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mag Fil V On"
               Height          =   195
               Index           =   11
               Left            =   360
               TabIndex        =   53
               Top             =   5010
               Width           =   915
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filement Time"
               Height          =   195
               Index           =   10
               Left            =   360
               TabIndex        =   52
               Top             =   4605
               Width           =   975
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Beam Time"
               Height          =   195
               Index           =   9
               Left            =   360
               TabIndex        =   51
               Top             =   4245
               Width           =   795
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Water Level"
               Height          =   195
               Index           =   8
               Left            =   360
               TabIndex        =   50
               Top             =   3885
               Width           =   870
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Water Temperature"
               Height          =   195
               Index           =   7
               Left            =   360
               TabIndex        =   49
               Top             =   3525
               Width           =   1380
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Freon Presure"
               Height          =   195
               Index           =   6
               Left            =   360
               TabIndex        =   48
               Top             =   3165
               Width           =   990
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5V N°2"
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   47
               Top             =   2805
               Width           =   510
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5V N°1"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   46
               Top             =   2445
               Width           =   510
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vacuum Standby"
               Height          =   195
               Index           =   3
               Left            =   360
               TabIndex        =   45
               Top             =   2085
               Width           =   1215
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gun fil StandBy"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   44
               Top             =   1725
               Width           =   1095
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mag Fil V StandBy"
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   43
               Top             =   1365
               Width           =   1305
            End
            Begin VB.Label LblCampos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Iniciales:"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   42
               Top             =   1020
               Width           =   615
            End
            Begin VB.Label LblDiaSemana 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "(DIA DE LA SEMANA)"
               Height          =   195
               Left            =   6480
               TabIndex        =   41
               Top             =   360
               Width           =   1590
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha del reporte"
               Height          =   195
               Left            =   360
               TabIndex        =   40
               Top             =   360
               Width           =   1245
            End
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   0
            Top             =   4080
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del equipo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   480
            TabIndex        =   300
            Top             =   360
            Width           =   1890
         End
         Begin VB.Label LblNombreEquipo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NOMBRE DEL EQUIPO"
            Height          =   195
            Left            =   2520
            TabIndex        =   73
            Top             =   405
            Width           =   1725
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   4680
         Width           =   4575
         Begin SystemOncoAmerica.DMGrid DMGrid2 
            Height          =   3015
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   5318
            Object.Width           =   4305
            Object.Height          =   2985
            BackColor       =   15396847
            ScrollBar       =   1
            MarqueeStyle    =   2
         End
         Begin ChamaleonButton.ChameleonBtn BtnVerReporte 
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            ToolTipText     =   "Agregar"
            Top             =   3840
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Ver Reporte"
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
            MICON           =   "FrmBitacoraEquipo.frx":415C
            PICN            =   "FrmBitacoraEquipo.frx":4178
            PICH            =   "FrmBitacoraEquipo.frx":45AA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   240
            TabIndex        =   298
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label LblReporteEquipo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NOMBRE DEL EQUIPO"
            Height          =   195
            Left            =   2280
            TabIndex        =   5
            Top             =   360
            Width           =   1725
         End
      End
   End
   Begin VB.Frame FrmOpc 
      BackColor       =   &H00EAEFEF&
      Height          =   9855
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   13455
      Begin VB.Frame Frame12 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   281
         Top             =   9000
         Width           =   13215
         Begin ChamaleonButton.ChameleonBtn BtnGuardarRismed 
            Height          =   375
            Left            =   7200
            TabIndex        =   282
            ToolTipText     =   "Guardar"
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
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
            MICON           =   "FrmBitacoraEquipo.frx":4834
            PICN            =   "FrmBitacoraEquipo.frx":4850
            PICH            =   "FrmBitacoraEquipo.frx":4ADF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
            Height          =   375
            Left            =   4920
            TabIndex        =   283
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Limpiar Campos"
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
            MICON           =   "FrmBitacoraEquipo.frx":4F20
            PICN            =   "FrmBitacoraEquipo.frx":4F3C
            PICH            =   "FrmBitacoraEquipo.frx":519C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
            Height          =   375
            Left            =   11880
            TabIndex        =   284
            ToolTipText     =   "Cerrar "
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmBitacoraEquipo.frx":5420
            PICN            =   "FrmBitacoraEquipo.frx":543C
            PICH            =   "FrmBitacoraEquipo.frx":5605
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
      Begin VB.Frame Frame10 
         BackColor       =   &H00EAEFEF&
         Caption         =   "  MECHANICAL CHECKS  "
         Height          =   1935
         Left            =   3600
         TabIndex        =   262
         Top             =   7080
         Width           =   9735
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7920
            TabIndex        =   306
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8680
            TabIndex        =   276
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Magnet Gauss"
            Height          =   255
            Index           =   12
            Left            =   6480
            TabIndex        =   275
            Top             =   840
            Width           =   3110
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Gun Resistor Configuration             ohm"
            Height          =   255
            Index           =   11
            Left            =   6480
            TabIndex        =   274
            Top             =   600
            Width           =   3110
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Pusle Transformer Tap (A,B,C)"
            Height          =   255
            Index           =   10
            Left            =   6480
            TabIndex        =   273
            Top             =   360
            Width           =   3110
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check MANDATORY MODIFICATIONS"
            Height          =   230
            Index           =   9
            Left            =   3360
            TabIndex        =   272
            Top             =   1320
            Width           =   3375
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check PSA/ETR && PENDANT Operation"
            Height          =   255
            Index           =   8
            Left            =   3360
            TabIndex        =   271
            Top             =   1080
            Width           =   3375
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check Accessories for Fit &&Operation"
            Height          =   255
            Index           =   7
            Left            =   3360
            TabIndex        =   270
            Top             =   840
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check HV Areas for HV Clearances"
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   269
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check Stand/Gantry/CW/BS/Bolts"
            Height          =   255
            Index           =   5
            Left            =   3360
            TabIndex        =   268
            Top             =   360
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check / Replace Gantry Motor Brushes"
            Height          =   230
            Index           =   4
            Left            =   240
            TabIndex        =   267
            Top             =   1320
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Clean all AIR && WATER Filters"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   265
            Top             =   840
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Clean && Lubricate Jaw/Couch Bearings"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   264
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Lubricate Gantry Bearings"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   263
            Top             =   360
            Width           =   3135
         End
         Begin VB.CheckBox ChkChks 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Check / Replace internal WATER"
            Height          =   230
            Index           =   3
            Left            =   240
            TabIndex        =   266
            Top             =   1080
            Width           =   3135
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Caption         =   "  VOLTAGE MEASUREMENTS  "
         Height          =   4575
         Left            =   120
         TabIndex        =   237
         Top             =   4440
         Width           =   3375
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   16
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   261
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   15
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   305
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   260
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   259
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   304
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   303
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   302
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   301
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   258
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   257
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   256
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   255
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   254
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   253
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   252
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   251
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox TxtCmps4 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   250
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+100 VDC Power Supply"
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   249
            Top             =   3000
            Width           =   1755
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+/- 27 VDC Power Supply"
            Height          =   195
            Index           =   36
            Left            =   120
            TabIndex        =   248
            Top             =   2760
            Width           =   1830
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-12V On Power Supply"
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   247
            Top             =   2520
            Width           =   1605
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+5V Power Supply (1) y (2)"
            Height          =   195
            Index           =   34
            Left            =   120
            TabIndex        =   246
            Top             =   2280
            Width           =   1875
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+/- 09V Power Supply"
            Height          =   195
            Index           =   33
            Left            =   120
            TabIndex        =   245
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+/- 10V Power Supply"
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   244
            Top             =   1800
            Width           =   1560
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+/- 15V Power Supply"
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   243
            Top             =   1560
            Width           =   1560
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+24V Power Supply"
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   242
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vacion Power Supply HV"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   241
            Top             =   1080
            Width           =   1785
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De- Q Thyratron Filament"
            Height          =   195
            Index           =   28
            Left            =   120
            TabIndex        =   240
            Top             =   840
            Width           =   1770
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Main (TP3) K. Alive Voltage"
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   239
            Top             =   600
            Width           =   1950
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Main Thyratron Filament"
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   238
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00EAEFEF&
         Caption         =   "  ELECTRICAL CHECKS  "
         Height          =   5535
         Left            =   8160
         TabIndex        =   207
         Top             =   1440
         Width           =   5175
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Gantry Limit Switches"
            Height          =   375
            Index           =   28
            Left            =   2640
            TabIndex        =   236
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Gantry Slowdown"
            Height          =   375
            Index           =   27
            Left            =   2640
            TabIndex        =   235
            Top             =   2040
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Gantry Underspeed"
            Height          =   375
            Index           =   26
            Left            =   2640
            TabIndex        =   234
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Emergency Pendant"
            Height          =   375
            Index           =   25
            Left            =   2640
            TabIndex        =   233
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Emergency Off - CONSOLE"
            Height          =   375
            Index           =   24
            Left            =   2640
            TabIndex        =   232
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Emergency Off - STAND"
            Height          =   375
            Index           =   23
            Left            =   2640
            TabIndex        =   231
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Emergency Off - WALL"
            Height          =   375
            Index           =   22
            Left            =   2640
            TabIndex        =   230
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Emergency Off - PSA/ETR"
            Height          =   375
            Index           =   21
            Left            =   2640
            TabIndex        =   229
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "EMC operation / battery"
            Height          =   375
            Index           =   20
            Left            =   2640
            TabIndex        =   228
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Lower Jaws Position Read - out Calibration"
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   227
            Top             =   4920
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Upper Jaws Position Read - out Calibration (&& LEIJ Servo)"
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   226
            Top             =   4680
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Collimator Position Read - out Calibration"
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   225
            Top             =   4440
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Gantry Position Read - out Calibration"
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   224
            Top             =   4200
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ARC CW 90 deg @ .5MU/DEG @ 100 R/MIN (44 - 46 MU)"
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   223
            Top             =   3960
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ARC CW 90 deg @ .5MU/DEG @ 100 R/MIN (44 - 46 MU)"
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   222
            Top             =   3720
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ARC CCW 90 deg @ 3MU/DEG @ 250 R/MIN (87 - 93 MU)"
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   221
            Top             =   3480
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ARC CW 90 deg @ 3MU/DEG @ 250 R/MIN (262 - 278 MU)"
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   220
            Top             =   3240
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ARC CCW 90 deg @ 1MU/DEG @ 250 R/MIN (87 - 93 MU)"
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   219
            Top             =   3000
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ARC CW 90 deg @ 1MU/DEG @ 250 R/MIN (87 - 93 MU)"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   218
            Top             =   2760
            Width           =   4815
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "HVPS Interlock"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   217
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "MOD Interlock"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   216
            Top             =   2280
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "MOTION Interlock (NON-LEIJ)"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   215
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "TEST Cycle"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   214
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Lamp Test"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   213
            Top             =   1560
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "TIME Coincidence"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   212
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "MU1 / MU2 Coincidence"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   211
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Full Field"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   210
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "YIELD Servo/Interlock"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   209
            Top             =   600
            Width           =   2535
         End
         Begin VB.CheckBox ChkVoltg 
            BackColor       =   &H00EAEFEF&
            Caption         =   "AFC Operation"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   208
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   2535
         Left            =   3600
         TabIndex        =   171
         Top             =   4440
         Width           =   4455
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   206
            Text            =   "N/A"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   22
            Left            =   3480
            MaxLength       =   10
            TabIndex        =   205
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   204
            Text            =   "N/A"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   199
            Text            =   "N/A"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   198
            Text            =   "N/A"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   18
            Left            =   3480
            MaxLength       =   10
            TabIndex        =   197
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   17
            Left            =   3480
            MaxLength       =   10
            TabIndex        =   196
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   16
            Left            =   3480
            MaxLength       =   10
            TabIndex        =   195
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   194
            Text            =   "N/A"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   193
            Text            =   "N/A"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   192
            Text            =   "N/A"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "N/A"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   190
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   189
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   188
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   187
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   186
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   185
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   184
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   183
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   182
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   181
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   180
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox TxtCmps3 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   179
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Water Tank Level"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   203
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOW TRIP"
            Height          =   195
            Index           =   24
            Left            =   3480
            TabIndex        =   202
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HIGH TRIP"
            Height          =   195
            Index           =   23
            Left            =   2520
            TabIndex        =   201
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   195
            Index           =   15
            Left            =   1800
            TabIndex        =   200
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mag/XFMR Flow"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   178
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guide/Load Flow"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   177
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Targ/Circ Flow"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   176
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Water Temp"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   175
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Water Pressure"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   174
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Freon RF PSIG"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   173
            Top             =   1920
            Width           =   1080
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Freon BTL PSIG"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   172
            Top             =   2160
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   3015
         Left            =   120
         TabIndex        =   105
         Top             =   1440
         Width           =   7935
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   49
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   170
            ToolTipText     =   "AMPS"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   48
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   169
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   47
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   168
            ToolTipText     =   "PPS"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   46
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   167
            ToolTipText     =   "%"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   45
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   166
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   44
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   165
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   43
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   164
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   42
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   163
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   41
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   162
            Text            =   "N/A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   40
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   161
            Text            =   "N/A"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   39
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   160
            Text            =   "N/A"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   38
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   159
            Text            =   "N/A"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   37
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   158
            ToolTipText     =   "PPS"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   36
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   157
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   35
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   156
            Text            =   "N/A"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   155
            Text            =   "N/A"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   33
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   154
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   32
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   153
            Text            =   "N/A"
            ToolTipText     =   "VDC"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   31
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   152
            ToolTipText     =   "VAC - VDC"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   30
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   151
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   29
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   150
            ToolTipText     =   "AMPS"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   28
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   149
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   27
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   148
            ToolTipText     =   "PPS"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   26
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   147
            ToolTipText     =   "%"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   25
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   146
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   24
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   145
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   23
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   144
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   143
            Text            =   "N/A"
            ToolTipText     =   "VDC"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   21
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   142
            ToolTipText     =   "VAC - VDC"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   140
            Text            =   "N/A"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   139
            Text            =   "N/A"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   138
            Text            =   "N/A"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   137
            Text            =   "N/A"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   136
            Text            =   "N/A"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   135
            Text            =   "N/A"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   134
            Text            =   "N/A"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   133
            ToolTipText     =   "VDC"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   132
            ToolTipText     =   "VAC - VDC"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   2400
            TabIndex        =   130
            Text            =   "N/A"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2400
            TabIndex        =   129
            Text            =   "N/A"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   2400
            TabIndex        =   128
            Text            =   "N/A"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   2400
            TabIndex        =   127
            Text            =   "N/A"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2400
            TabIndex        =   126
            Text            =   "N/A"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   125
            Text            =   "N/A"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   141
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   131
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   124
            Text            =   "N/A"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   123
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   122
            ToolTipText     =   "VAC - VDC"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtCmps2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   121
            ToolTipText     =   "VDC"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STEP 1 beam"
            Height          =   435
            Index           =   12
            Left            =   4440
            TabIndex        =   118
            Top             =   110
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "@ MAX RATE"
            Height          =   195
            Index           =   14
            Left            =   6240
            TabIndex        =   120
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STEP 2 beam"
            Height          =   435
            Index           =   13
            Left            =   5400
            TabIndex        =   119
            Top             =   110
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   195
            Index           =   11
            Left            =   3720
            TabIndex        =   117
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STAND BY"
            Height          =   195
            Index           =   10
            Left            =   2400
            TabIndex        =   116
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magnetron Current (AMPS)"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   115
            Top             =   2640
            Width           =   1905
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De - Quing Rate (%)"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   114
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pulse Repetition Rate (PPS)"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   113
            Top             =   2160
            Width           =   1995
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radiation Output (MU/MIN)"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   112
            Top             =   1920
            Width           =   1965
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beam (target) Current"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   111
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PFN Voltage"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   110
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Power Supply Current"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   109
            Top             =   1200
            Width           =   1530
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vacuum levels"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   108
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gun Filament Voltage"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magnetron Filament Voltage"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   106
            Top             =   480
            Width           =   1980
         End
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   11640
         TabIndex        =   99
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   11640
         TabIndex        =   98
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   8160
         TabIndex        =   97
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   8160
         TabIndex        =   96
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   7560
         TabIndex        =   95
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   94
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtCmps 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   93
         Top             =   240
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   11640
         TabIndex        =   92
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   51642369
         CurrentDate     =   40351
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance Parameters :  CL4/6X, CL4-6/100, CL600C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   289
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beam Hours:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10440
         TabIndex        =   104
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11160
         TabIndex        =   103
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. FSR #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         TabIndex        =   102
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filament Hours:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         TabIndex        =   101
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clinac:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         TabIndex        =   100
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   10920
         TabIndex        =   91
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rismed Rep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2640
         TabIndex        =   90
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2640
         TabIndex        =   89
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oncology Systems"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   88
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RISMED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   360
         TabIndex        =   87
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.OptionButton ChkOpc 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Rismed's Reports"
      Height          =   375
      Index           =   2
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   287
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton ChkOpc 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Rismed"
      Height          =   375
      Index           =   1
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   286
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton ChkOpc 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Tecnicos"
      Height          =   375
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   285
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.Frame FrmOpc 
      BackColor       =   &H00EAEFEF&
      Height          =   9855
      Index           =   2
      Left            =   240
      TabIndex        =   288
      Top             =   480
      Width           =   13455
      Begin VB.Frame Frame14 
         BackColor       =   &H00EAEFEF&
         Height          =   4935
         Left            =   2640
         TabIndex        =   293
         Top             =   1800
         Width           =   8175
         Begin SystemOncoAmerica.DMGrid DMGrid3 
            Height          =   4575
            Left            =   120
            TabIndex        =   294
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   8070
            Object.Width           =   7905
            Object.Height          =   4545
            Rows            =   0
            BackColor       =   15396847
            ScrollBar       =   1
            MarqueeStyle    =   2
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   290
         Top             =   9000
         Width           =   13215
         Begin ChamaleonButton.ChameleonBtn BtnVerReporterismed 
            Height          =   375
            Left            =   5640
            TabIndex        =   291
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Ver Reporte"
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
            MICON           =   "FrmBitacoraEquipo.frx":583A
            PICN            =   "FrmBitacoraEquipo.frx":5856
            PICH            =   "FrmBitacoraEquipo.frx":5C88
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn5 
            Height          =   375
            Left            =   11880
            TabIndex        =   292
            ToolTipText     =   "Cerrar "
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmBitacoraEquipo.frx":5F0C
            PICN            =   "FrmBitacoraEquipo.frx":5F28
            PICH            =   "FrmBitacoraEquipo.frx":60F1
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RISMED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   5760
         TabIndex        =   297
         Top             =   480
         Width           =   1830
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oncology Systems"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   5760
         TabIndex        =   296
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance Parameters :  CL4/6X, CL4-6/100, CL600C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   295
         Top             =   1440
         Width           =   4815
      End
   End
End
Attribute VB_Name = "FrmBitacoraEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemp As Recordset
Dim RsTemp2 As Recordset
Dim Nuevo_Reg As Boolean
Dim i As Integer

Sub IniDMGrid()
On Error Resume Next

DMGrid1.Cols = 4

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 55 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 25 / 100) - 300
DMGrid1.DColumnas(4).Visible = False

DMGrid1.DColumnas(1).Caption = "Nombre"
DMGrid1.DColumnas(2).Caption = "Modelo"
DMGrid1.DColumnas(3).Caption = "Serial"
DMGrid1.DColumnas(4).Caption = "IdEquipo"

DMGrid2.Cols = 4

DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 25 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 25 / 100)
DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 50 / 100) - 300
DMGrid2.DColumnas(4).Visible = False

DMGrid2.DColumnas(1).Caption = "Fecha"
DMGrid2.DColumnas(2).Caption = "# Reporte"
DMGrid2.DColumnas(3).Caption = "Usuario"

DMGrid3.Cols = 4

DMGrid3.DColumnas(1).Width = Val(DMGrid3.Width * 25 / 100)
DMGrid3.DColumnas(2).Width = Val(DMGrid3.Width * 25 / 100)
DMGrid3.DColumnas(3).Width = Val(DMGrid3.Width * 50 / 100) - 300
DMGrid3.DColumnas(4).Visible = False

DMGrid3.DColumnas(1).Caption = "Fecha"
DMGrid3.DColumnas(2).Caption = "# Reporte"
DMGrid3.DColumnas(3).Caption = "Hospital / Clinica"
DMGrid3.DColumnas(4).Caption = "Id"


End Sub

Sub Leer_Equipos()
On Error Resume Next

CSql = "SELECT * FROM TecnicoEquipo WHERE Activo='1'"
Set RsTemp = CrearRS(CSql)

DMGrid1.Clear
DMGrid1.Rows = 0

While Not RsTemp.EOF
    
    DMGrid1.Rows = DMGrid1.Rows + 1
    
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = Trim(RsTemp.Fields("Nombre").Value)
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Trim(RsTemp.Fields("Modelo").Value)
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = Trim(RsTemp.Fields("Serial").Value)
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = Trim(RsTemp.Fields("IdEquipo").Value)
    
    RsTemp.MoveNext
Wend

DMGrid1.RowBackColor 1, RGB(255, 255, 255)
DMGrid1.PaintMGrid
End Sub

Sub Leer_Rismed()
On Error Resume Next

CSql = "SELECT * FROM TecnicoEquipoReporte2 ORDER BY Fecha"
Set RsTemp = CrearRS(CSql)

DMGrid3.Clear
DMGrid3.Rows = 0

While Not RsTemp.EOF
    
    DMGrid3.Rows = DMGrid3.Rows + 1
    
    DMGrid3.ValorCelda(DMGrid3.Rows, 1) = Trim(RsTemp.Fields("Fecha").Value)
    DMGrid3.ValorCelda(DMGrid3.Rows, 2) = Trim(RsTemp.Fields("NroReporte").Value)
    DMGrid3.ValorCelda(DMGrid3.Rows, 3) = Trim(RsTemp.Fields("Hospital").Value)
    DMGrid3.ValorCelda(DMGrid3.Rows, 4) = Trim(RsTemp.Fields("Id").Value)
    
    RsTemp.MoveNext
Wend

DMGrid3.RowBackColor 1, RGB(255, 255, 255)
DMGrid3.PaintMGrid

End Sub

Sub Leer_Reportes()
On Error Resume Next
Dim IdUsuario As Integer
Dim NombTec As String

CSql = "SELECT * FROM TecnicoEquipoReporte WHERE Activo='1' AND IdEquipo=" & DMGrid1.ValorCelda(DMGrid1.Row, 4)
Set RsTemp = CrearRS(CSql)

DMGrid2.Clear
DMGrid2.Rows = 0

IdUsuario = -1

If RsTemp.RecordCount <> 0 Then IdUsuario = Val(RsTemp.Fields("IdUser").Value)

If IdUsuario <> -1 Then
    CSql = "SELECT Nombre, Apellidos FROM Usuarios WHERE IdUsuario=" & IdUsuario
    Set RsTemp2 = CrearRS(CSql)
    NombTec = RsTemp2.Fields("Nombre").Value & ", " & RsTemp2.Fields("Apellidos").Value
End If

While Not RsTemp.EOF
    
    DMGrid2.Rows = DMGrid2.Rows + 1
    
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = Trim(RsTemp.Fields("Fecha").Value)
    DMGrid2.ValorCelda(DMGrid2.Rows, 2) = Trim(RsTemp.Fields("NroReporte").Value)
    
    CSql = "SELECT Nombre, Apellidos FROM Usuarios WHERE IdUsuario=" & Trim(RsTemp.Fields("IdUser").Value)
    Set RsTemp2 = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then NombTec = RsTemp2.Fields("Nombre").Value & ", " & RsTemp2.Fields("Apellidos").Value Else NombTec = " "
    
    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = NombTec
    DMGrid2.ValorCelda(DMGrid2.Rows, 4) = Trim(RsTemp.Fields("IdCampoEquipo").Value)
    
    RsTemp.MoveNext
    
Wend

DMGrid2.RowBackColor 1, RGB(255, 255, 255)
DMGrid2.PaintMGrid
End Sub

Private Sub BtnAgregar_Click()
On Error Resume Next
Nuevo_Reg = True

TxtNombre.Text = ""
TxtModelo.Text = ""
TxtSerial.Text = ""

Frame2.Visible = True
DMGrid1.Visible = False

End Sub

Private Sub BtnCerrar_Click()
On Error Resume Next
Frame2.Visible = False
DMGrid1.Visible = True
End Sub

Private Sub BtnCerrarFrm_Click()
Unload FrmBitacoraEquipo
End Sub

Private Sub BtnGuardar_Click()
On Error GoTo Error
Dim NuevoId As Integer
Dim NuevoNro As Integer
Dim resp

If Trim(LblNombreEquipo.Caption) <> Trim(DMGrid1.ValorCelda(DMGrid1.Row, 1)) Then
    MsgBox "Elija el equipo al cual se le realizará el Reporte.", vbInformation + vbOKOnly, "Información"
    Exit Sub
End If
CSql = "SELECT * FROM TecnicoEquipoReporte WHERE IdEquipo=" & DMGrid1.ValorCelda(DMGrid1.Row, 4) & " AND Fecha='" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    MsgBox "Ya existe un reporte en la base de datos para el equipo y fecha seleccionado!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCampos.Count - 1
    If Trim(TxtCampos(i).Text) = "" Then
        MsgBox "Debe llenar todos los campos!", vbExclamation + vbOKOnly, "Faltan Datos!"
        TxtCampos(i).SetFocus
        TxtCampos(i).BackColor = vbYellow
        Exit Sub
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


resp = MsgBox("Se procederá a guardar los cambios, Desea continuar?", vbQuestion + vbYesNo, "Confirmar Operación!")

If resp = vbNo Then Exit Sub


For i = 0 To TxtCampos.Count - 1
    If Trim(TxtCampos(i).Text) = "" Then
        MsgBox "Debe ingresar un valor en el campo ' " & UCase(LblCampos(i).Caption) & " '", vbInformation + vbOKOnly, "Faltan Datos!"
        Exit Sub
    End If
Next i

' Obtiene el maximo valor de los registros de la tabla "TecnicoEquipoReporte"
CSql = "SELECT MAX(IdCampoEquipo)+1 as NuevoId FROM TecnicoEquipoReporte"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If

' Obtiene el maximo valor del reporte de un Equipo determinado
CSql = "SELECT MAX(NroReporte)+1 as NuevoRep FROM TecnicoEquipoReporte WHERE IdEquipo=" & DMGrid1.ValorCelda(DMGrid1.Row, 4)
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoNro = Val(RsTemp.Fields(0).Value)
Else
    NuevoNro = 1
End If

    CSql = "SELECT * FROM TecnicoEquipoReporte"
    Set RsTemp = CrearRS(CSql)
    
    RsTemp.AddNew
    RsTemp.Fields("IdCampoEquipo").Value = NuevoId
    RsTemp.Fields("IdEquipo").Value = Val(DMGrid1.ValorCelda(DMGrid1.Row, 4))
    RsTemp.Fields("IdUser").Value = IdUser
    RsTemp.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/MM/yyyy")
    RsTemp.Fields("NroReporte").Value = NuevoNro
    
    RsTemp.Fields("IdTecnicos").Value = CboTecnicos.ItemData(CboTecnicos.Index)
    RsTemp.Fields("Iniciales").Value = Trim(TxtCampos(0).Text)
    RsTemp.Fields("MFV_StandBy").Value = TxtCampos(1).Text
    RsTemp.Fields("GF_StandBy").Value = TxtCampos(2).Text
    RsTemp.Fields("V_StandBy").Value = TxtCampos(3).Text
    RsTemp.Fields("5VN1").Value = TxtCampos(4).Text
    RsTemp.Fields("5VN2").Value = TxtCampos(5).Text
    RsTemp.Fields("Freon_Presure").Value = TxtCampos(6).Text
    RsTemp.Fields("Water_Temp").Value = TxtCampos(7).Text
    RsTemp.Fields("Water_Lvl").Value = TxtCampos(8).Text
    RsTemp.Fields("Beam_Time").Value = TxtCampos(9).Text
    RsTemp.Fields("Filement_Time").Value = TxtCampos(10).Text
    RsTemp.Fields("MFV_ON").Value = TxtCampos(11).Text
    RsTemp.Fields("GF_ON").Value = TxtCampos(12).Text
    RsTemp.Fields("V_ON").Value = TxtCampos(13).Text
    RsTemp.Fields("MUPM").Value = TxtCampos(14).Text
    RsTemp.Fields("MFV_RUN").Value = TxtCampos(15).Text
    RsTemp.Fields("GF_RUN").Value = TxtCampos(16).Text
    RsTemp.Fields("V_RUN").Value = TxtCampos(17).Text
    RsTemp.Fields("PSC").Value = TxtCampos(18).Text
    RsTemp.Fields("PFN_Volts").Value = TxtCampos(19).Text
    RsTemp.Fields("Beam_Current").Value = TxtCampos(20).Text
    RsTemp.Fields("TaI").Value = TxtCampos(21).Text
    RsTemp.Fields("Arc_Test").Value = TxtCampos(22).Text
    RsTemp.Fields("Mecanical_C").Value = TxtCampos(23).Text
    RsTemp.Fields("BMD").Value = TxtCampos(24).Text
    RsTemp.Fields("Temp_Bunker").Value = TxtCampos(25).Text
    RsTemp.Fields("Humedad_Bunker").Value = TxtCampos(26).Text
    RsTemp.Fields("PBB").Value = TxtCampos(27).Text
    RsTemp.Fields("ChequeoDD").Value = TxtCampos(28).Text
    RsTemp.Fields("PresionAguaInt").Value = TxtCampos(29).Text
    RsTemp.Fields("PresionAguaExt").Value = TxtCampos(30).Text
    RsTemp.Fields("Activo").Value = "1"
    
    RsTemp.Update
    
    Leer_Reportes
    
Exit Sub

Error:
    MsgBox "Verifique la información ingresada!", vbInformation + vbOKOnly, "Información"
    
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next

If Trim(TxtNombre.Text) = "" Then
    MsgBox "Debe ingresar un Nombre para el nuevo Equipo!", vbExclamation + vbOKOnly, "Faltan datos."
    Exit Sub
ElseIf Trim(TxtModelo.Text) = "" Then
    MsgBox "Debe ingresar el Modelo para el nuevo Equipo!", vbExclamation + vbOKOnly, "Faltan datos."
    Exit Sub
ElseIf Trim(TxtSerial.Text) = "" Then
    MsgBox "Debe ingresar un Serial para el nuevo Equipo!", vbExclamation + vbOKOnly, "Faltan datos."
    Exit Sub
End If

Dim resp
resp = MsgBox("Se guardarán los cambios realizados, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")

If resp = vbNo Then Exit Sub


If Nuevo_Reg Then
    
    Dim NuevoId As Integer
    
    ' Obtiene el valor maximo de la tabla, para anexar el nuevo registro del equipo
    CSql = "SELECT MAX(IdEquipo)+1 As NuevoId FROM TecnicoEquipo"
    Set RsTemp = CrearRS(CSql)
    
    If Not IsNull(RsTemp.Fields(0).Value) Then
        NuevoId = Val(RsTemp.Fields(0).Value)
        Else
        NuevoId = 1
    End If
    
    CSql = "INSERT INTO TecnicoEquipo (IdEquipo,IdUsuario,Nombre,Modelo,Serial,FechaIng,Activo) " & _
            " VALUES (" & NuevoId & " , " & IdUser & ",'" & Trim(TxtNombre.Text) & _
            "','" & Trim(TxtModelo.Text) & "','" & Trim(TxtSerial.Text) & "','" & Format(Now, "dd/MM/yyyy") & "',1)"
    Set RsTemp = CrearRS(CSql)
    
    MsgBox "Los datos fueron agregados correctamente!", vbInformation + vbOKOnly, "Operación Exitosa!"
    
Else
    'CSql = "UPDATE TecnicoEquipo SET IdUsuario= " & IdUser & ", Nombre='" & Trim(TxtNombre.Text) & _
    '"', Modelo='" & Trim(TxtModelo.Text) & "', Serial='" & Trim(TxtSerial.Text) & _
    '"' WHERE IdEquipo = " & Val(DMGrid1.ValorCelda(DMGrid1.Row, 4))
    'Set RsTemp = CrearRS(CSql)
    
    CSql = "SELECT * FROM TecnicoEquipo WHERE IdEquipo = " & Val(DMGrid1.ValorCelda(DMGrid1.Row, 4))
    Set RsTemp = CrearRS(CSql)
    
    RsTemp.Fields("IdUsuario").Value = IdUser
    RsTemp.Fields("Nombre").Value = Trim(TxtNombre.Text)
    RsTemp.Fields("Modelo").Value = Trim(TxtModelo.Text)
    RsTemp.Fields("Serial").Value = Trim(TxtSerial.Text)
    
    RsTemp.Update
    
    MsgBox "Los datos fueron actualizados correctamente!", vbInformation + vbOKOnly, "Operación Exitosa!"
End If

Frame2.Visible = False
DMGrid1.Visible = True

Dim PosAct As Integer

PosAct = DMGrid1.Row
Leer_Equipos
DMGrid1.Row = PosAct

End Sub

Private Sub BtnGuardarRismed_Click()
Dim Rsp


CSql = "SELECT * FROM TecnicoEquipoReporte2 WHERE Fecha='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    MsgBox "Ya existe un reporte en la base de datos para la fecha seleccionada!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps.Count - 1
    If Trim(TxtCmps(i).Text) = "" Then
        MsgBox "Debe llenar todos los campos!", vbExclamation + vbOKOnly, "Faltan Datos!"
        TxtCmps(i).SetFocus
        TxtCmps(i).BackColor = vbYellow
        Exit Sub
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps2.Count - 1
    If Trim(TxtCmps2(i).Text) = "" Then
        MsgBox "Debe llenar todos los campos!", vbExclamation + vbOKOnly, "Faltan Datos!"
        TxtCmps2(i).SetFocus
        TxtCmps2(i).BackColor = vbYellow
        Exit Sub
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps3.Count - 1
    If Trim(TxtCmps3(i).Text) = "" Then
        MsgBox "Debe llenar todos los campos!", vbExclamation + vbOKOnly, "Faltan Datos!"
        TxtCmps3(i).SetFocus
        TxtCmps3(i).BackColor = vbYellow
        Exit Sub
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps4.Count - 1
    If Trim(TxtCmps4(i).Text) = "" Then
        MsgBox "Debe llenar todos los campos!", vbExclamation + vbOKOnly, "Faltan Datos!"
        TxtCmps4(i).SetFocus
        TxtCmps4(i).BackColor = vbYellow
        Exit Sub
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Dim Band As Boolean
For i = 0 To ChkVoltg.Count - 1
    If ChkVoltg(i).Value = 0 Then
        ChkVoltg(i).BackColor = vbYellow
        Band = True
    End If
Next i

For i = 0 To ChkChks.Count - 1
    If ChkChks(i).Value = 0 Then
        ChkChks(i).BackColor = vbYellow
        Band = True
    End If
Next i

If Band Then MsgBox "No se han chequeado todos los Items!, los Items remarcados son los faltantes." _
                    , vbInformation + vbOKOnly, "Información"

Rsp = MsgBox("Se guardaran los cambios, Desea Continuar?", vbQuestion + vbYesNo, "Confirmación")

If Rsp = vbNo Then Exit Sub

CSql = "SELECT MagVolt_Stand_By,GunVolt_Stand_By, VacuumLVL_Stand_By, " & _
  "PowerSC_Stand_By, PFNVolt_Stand_By, BeamC_Stand_By, RadiationOP_Stand_By, PulseRR_Stand_By, " & _
  "[De-QR_Stand_By],  MagCurr_Stand_By,  MagVolt_ON,  GunVolt_ON,  VacuumLVL_ON,  PowerSC_ON, " & _
  "PFNVolt_ON,  BeamC_ON,  RadiationOP,  PulseRR_ON,  [De-QR_ON],  MagCurr_ON,  MagVolt_STEP1, " & _
  "GunVolt_STEP1,  VacuumLVL_STEP1,  PowerSC_STEP1,  PFNVolt_STEP1,  BeamC_STEP1,  Radiation_STEP1, " & _
  "PulseRR_STEP1,  [De-QR_STEP1],  MagCurr_STEP1,  MagVolt_STEP2,  GunVolt_STEP2,  VacuumLVL_STEP2, " & _
  "PowerSC_STEP2,  PFNVolt_STEP2,  BeamC_STEP2,  Radiation_STEP2,  PulseRR_STEP2,  [De-QR_STEP2], " & _
  "MagCurr_STEP2,  MagVolt_MRATE,  GunVolt_MRATE,  VacuumLVL_MRATE,  PowerSC_MRATE,  PFNVolt_MRATE, " & _
  "BeamC_MRATE,  Radiation_MRATE,  PulseRR_MRATE,  [De-QR_MRATE],  MagCurr_MRATE,  MainThyraton, " & _
  "MainKAV, [De-QTF],  VacionPShv,  [24VPS],  [15VPS],  [10VPS],  [9VPS],  [5VPS],  [12VPS],  [27VDCps], " & _
  "[100VDCps], MagXFMR_ON,  GuideLoad_ON,  TargCirc_ON,  WaterTemp_ON,  WaterPressure_ON,  WaterTank_ON, " & _
  "FreonRF_ON, FreonBTL_ON,  MagXFMR_HIGH,  GuideLoad_HIGH,  TargCirc_HIGH,  WaterTemp_HIGH, " & _
  "WaterPressure_HIGH,  WaterTank_HIGH,  FreonRF_HIGH,  FreonBTL_HIGH,  MagXFMR_LOW,  GuideLoad_LOW, " & _
  "TargCirc_LOW, WaterTemp_LOW,  WaterPressure_LOW,  WaterTank_LOW,  FreonRF_LOW,  FreonBTL_LOW, " & _
  "LubricateGB, Clean,  CleanAll,  CheckW,  CheckGMB,  CheckSGcwBS,  CheckHV,  CheckAccesories, " & _
  "CheckPSAetr, CheckMANDATORY,  PusleTT,  GunRC,  MagnetGauss,  AFCOperation,  YIELD,  FullField, " & _
  "MU1MU2, TIMEC, LampTest, TESTC, MOTION, MOD, HVPS, ARCCW901250, ARCCCW901250, ARCCW903250, " & _
  "ARCCCW903250, ARCCW905100, ARCCCW905100, GantryPR, CollimatorPS, UpperJPR, LowerJPR, EMCOperation, " & _
  "Emergency_PSAetr, Emergency_WALL, Emergency_STAND, Emergency_CONSOLE, Emergency_Pedante, " & _
  "GantryU , GantryS, GantryLS, Hospital, RismedRep, RefFSR, Clinac, FilamenteHours, Serial, BeamHours,  " & _
  "Id,IdUser,Fecha,NroReporte FROM TecnicoEquipoReporte2"


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Dim NuevoId As Integer
Dim NoRepor As Integer

CSql = "SELECT MAX(Id)+1 AS NuevoId FROM TecnicoEquipoReporte2"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
CSql = "SELECT MAX(NroReporte)+1 AS NuevoId FROM TecnicoEquipoReporte2"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NoRepor = Val(RsTemp.Fields(0).Value)
Else
    NoRepor = 1
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "SELECT * FROM TecnicoEquipoReporte2"
Set RsTemp = CrearRS(CSql)
RsTemp.AddNew

RsTemp.Fields("Id").Value = NuevoId
RsTemp.Fields("IdUser").Value = IdUser
RsTemp.Fields("Fecha").Value = Format(Now, "dd/MM/yyyy")
RsTemp.Fields("NroReporte").Value = NoRepor

RsTemp.Fields("MagVolt_Stand_By").Value = Trim(TxtCmps2(0).Text)
RsTemp.Fields("GunVolt_Stand_By").Value = Trim(TxtCmps2(1).Text)
RsTemp.Fields("VacuumLVL_Stand_By").Value = Trim(TxtCmps2(2).Text)
RsTemp.Fields("PowerSC_Stand_By").Value = Trim(TxtCmps2(3).Text)
RsTemp.Fields("PFNVolt_Stand_By").Value = Trim(TxtCmps2(4).Text)
RsTemp.Fields("BeamC_Stand_By").Value = Trim(TxtCmps2(5).Text)
RsTemp.Fields("RadiationOP_Stand_By").Value = Trim(TxtCmps2(6).Text)
RsTemp.Fields("PulseRR_Stand_By").Value = Trim(TxtCmps2(7).Text)
RsTemp.Fields("De-QR_Stand_By").Value = Trim(TxtCmps2(8).Text)
RsTemp.Fields("MagCurr_Stand_By").Value = Trim(TxtCmps2(9).Text)
RsTemp.Fields("MagVolt_ON").Value = Trim(TxtCmps2(10).Text)
RsTemp.Fields("GunVolt_ON").Value = Trim(TxtCmps2(11).Text)
RsTemp.Fields("VacuumLVL_ON").Value = Trim(TxtCmps2(12).Text)
RsTemp.Fields("PowerSC_ON").Value = Trim(TxtCmps2(13).Text)
RsTemp.Fields("PFNVolt_ON").Value = Trim(TxtCmps2(14).Text)
RsTemp.Fields("BeamC_ON").Value = Trim(TxtCmps2(15).Text)
RsTemp.Fields("RadiationOP").Value = Trim(TxtCmps2(16).Text)
RsTemp.Fields("PulseRR_ON").Value = Trim(TxtCmps2(17).Text)
RsTemp.Fields("De-QR_ON").Value = Trim(TxtCmps2(18).Text)
RsTemp.Fields("MagCurr_ON").Value = Trim(TxtCmps2(19).Text)
RsTemp.Fields("MagVolt_STEP1").Value = Trim(TxtCmps2(20).Text)
RsTemp.Fields("GunVolt_STEP1").Value = Trim(TxtCmps2(21).Text)
RsTemp.Fields("VacuumLVL_STEP1").Value = Trim(TxtCmps2(22).Text)
RsTemp.Fields("PowerSC_STEP1").Value = Trim(TxtCmps2(23).Text)
RsTemp.Fields("PFNVolt_STEP1").Value = Trim(TxtCmps2(24).Text)
RsTemp.Fields("BeamC_STEP1").Value = Trim(TxtCmps2(25).Text)
RsTemp.Fields("Radiation_STEP1").Value = Trim(TxtCmps2(26).Text)
RsTemp.Fields("PulseRR_STEP1").Value = Trim(TxtCmps2(27).Text)
RsTemp.Fields("De-QR_STEP1").Value = Trim(TxtCmps2(28).Text)
RsTemp.Fields("MagCurr_STEP1").Value = Trim(TxtCmps2(29).Text)
RsTemp.Fields("MagVolt_STEP2").Value = Trim(TxtCmps2(30).Text)
RsTemp.Fields("GunVolt_STEP2").Value = Trim(TxtCmps2(31).Text)
RsTemp.Fields("VacuumLVL_STEP2").Value = Trim(TxtCmps2(32).Text)
RsTemp.Fields("PowerSC_STEP2").Value = Trim(TxtCmps2(33).Text)
RsTemp.Fields("PFNVolt_STEP2").Value = Trim(TxtCmps2(34).Text)
RsTemp.Fields("BeamC_STEP2").Value = Trim(TxtCmps2(35).Text)
RsTemp.Fields("Radiation_STEP2").Value = Trim(TxtCmps2(36).Text)
RsTemp.Fields("PulseRR_STEP2").Value = Trim(TxtCmps2(37).Text)
RsTemp.Fields("De-QR_STEP2").Value = Trim(TxtCmps2(38).Text)
RsTemp.Fields("MagCurr_STEP2").Value = Trim(TxtCmps2(39).Text)
RsTemp.Fields("MagVolt_MRATE").Value = Trim(TxtCmps2(40).Text)
RsTemp.Fields("GunVolt_MRATE").Value = Trim(TxtCmps2(41).Text)
RsTemp.Fields("VacuumLVL_MRATE").Value = Trim(TxtCmps2(42).Text)
RsTemp.Fields("PowerSC_MRATE").Value = Trim(TxtCmps2(43).Text)
RsTemp.Fields("PFNVolt_MRATE").Value = Trim(TxtCmps2(44).Text)
RsTemp.Fields("BeamC_MRATE").Value = Trim(TxtCmps2(45).Text)
RsTemp.Fields("Radiation_MRATE").Value = Trim(TxtCmps2(46).Text)
RsTemp.Fields("PulseRR_MRATE").Value = Trim(TxtCmps2(47).Text)
RsTemp.Fields("De-QR_MRATE").Value = Trim(TxtCmps2(48).Text)
RsTemp.Fields("MagCurr_MRATE").Value = Trim(TxtCmps2(49).Text)

RsTemp.Fields("MainThyraton").Value = Trim(TxtCmps4(0).Text)
RsTemp.Fields("MainKAV").Value = Trim(TxtCmps4(1).Text)
RsTemp.Fields("De-QTF").Value = Trim(TxtCmps4(2).Text)
RsTemp.Fields("VacionPShv").Value = Trim(TxtCmps4(3).Text)
RsTemp.Fields("24VPS").Value = Trim(TxtCmps4(4).Text)
RsTemp.Fields("15VPS").Value = Trim(TxtCmps4(5).Text)
RsTemp.Fields("15VPS2").Value = Trim(TxtCmps4(6).Text)
RsTemp.Fields("10VPS").Value = Trim(TxtCmps4(7).Text)
RsTemp.Fields("10VPS2").Value = Trim(TxtCmps4(8).Text)
RsTemp.Fields("9VPS").Value = Trim(TxtCmps4(9).Text)
RsTemp.Fields("9VPS2").Value = Trim(TxtCmps4(10).Text)
RsTemp.Fields("5VPS").Value = Trim(TxtCmps4(11).Text)
RsTemp.Fields("5VPS2").Value = Trim(TxtCmps4(12).Text)
RsTemp.Fields("12VPS").Value = Trim(TxtCmps4(13).Text)
RsTemp.Fields("27VDCps").Value = Trim(TxtCmps4(14).Text)
RsTemp.Fields("27VDCps2").Value = Trim(TxtCmps4(15).Text)
RsTemp.Fields("100VDCps").Value = Trim(TxtCmps4(16).Text)

RsTemp.Fields("MagXFMR_ON").Value = Trim(TxtCmps3(0).Text)
RsTemp.Fields("GuideLoad_ON").Value = Trim(TxtCmps3(1).Text)
RsTemp.Fields("TargCirc_ON").Value = Trim(TxtCmps3(2).Text)
RsTemp.Fields("WaterTemp_ON").Value = Trim(TxtCmps3(3).Text)
RsTemp.Fields("WaterPressure_ON").Value = Trim(TxtCmps3(4).Text)
RsTemp.Fields("WaterTank_ON").Value = Trim(TxtCmps3(5).Text)
RsTemp.Fields("FreonRF_ON").Value = Trim(TxtCmps3(6).Text)
RsTemp.Fields("FreonBTL_ON").Value = Trim(TxtCmps3(7).Text)
RsTemp.Fields("MagXFMR_HIGH").Value = Trim(TxtCmps3(8).Text)
RsTemp.Fields("GuideLoad_HIGH").Value = Trim(TxtCmps3(9).Text)
RsTemp.Fields("TargCirc_HIGH").Value = Trim(TxtCmps3(10).Text)
RsTemp.Fields("WaterTemp_HIGH").Value = Trim(TxtCmps3(11).Text)
RsTemp.Fields("WaterPressure_HIGH").Value = Trim(TxtCmps3(12).Text)
RsTemp.Fields("WaterTank_HIGH").Value = Trim(TxtCmps3(13).Text)
RsTemp.Fields("FreonRF_HIGH").Value = Trim(TxtCmps3(14).Text)
RsTemp.Fields("FreonBTL_HIGH").Value = Trim(TxtCmps3(15).Text)
RsTemp.Fields("MagXFMR_LOW").Value = Trim(TxtCmps3(16).Text)
RsTemp.Fields("GuideLoad_LOW").Value = Trim(TxtCmps3(17).Text)
RsTemp.Fields("TargCirc_LOW").Value = Trim(TxtCmps3(18).Text)
RsTemp.Fields("WaterTemp_LOW").Value = Trim(TxtCmps3(19).Text)
RsTemp.Fields("WaterPressure_LOW").Value = Trim(TxtCmps3(20).Text)
RsTemp.Fields("WaterTank_LOW").Value = Trim(TxtCmps3(21).Text)
RsTemp.Fields("FreonRF_LOW").Value = Trim(TxtCmps3(22).Text)
RsTemp.Fields("FreonBTL_LOW").Value = Trim(TxtCmps3(23).Text)

RsTemp.Fields("LubricateGB").Value = ChkChks(0).Value
RsTemp.Fields("Clean").Value = ChkChks(1).Value
RsTemp.Fields("CleanAll").Value = ChkChks(2).Value
RsTemp.Fields("CheckW").Value = ChkChks(3).Value
RsTemp.Fields("CheckGMB").Value = ChkChks(4).Value
RsTemp.Fields("CheckSGcwBS").Value = ChkChks(5).Value
RsTemp.Fields("CheckHV").Value = ChkChks(6).Value
RsTemp.Fields("CheckAccesories").Value = ChkChks(7).Value
RsTemp.Fields("CheckPSAetr").Value = ChkChks(8).Value
RsTemp.Fields("CheckMANDATORY").Value = ChkChks(9).Value
RsTemp.Fields("PusleTT").Value = ChkChks(10).Value
RsTemp.Fields("GunRC").Value = Trim(Text1.Text)
RsTemp.Fields("MagnetGauss").Value = Trim(Text2.Text)

RsTemp.Fields("AFCOperation").Value = ChkVoltg(0).Value
RsTemp.Fields("YIELD").Value = ChkVoltg(1).Value
RsTemp.Fields("FullField").Value = ChkVoltg(2).Value
RsTemp.Fields("MU1MU2").Value = ChkVoltg(3).Value
RsTemp.Fields("TIMEC").Value = ChkVoltg(4).Value
RsTemp.Fields("LampTest").Value = ChkVoltg(5).Value
RsTemp.Fields("TESTC").Value = ChkVoltg(6).Value
RsTemp.Fields("MOTION").Value = ChkVoltg(7).Value
RsTemp.Fields("MOD").Value = ChkVoltg(8).Value
RsTemp.Fields("HVPS").Value = ChkVoltg(9).Value
RsTemp.Fields("ARCCW901250").Value = ChkVoltg(10).Value
RsTemp.Fields("ARCCCW901250").Value = ChkVoltg(11).Value
RsTemp.Fields("ARCCW903250").Value = ChkVoltg(12).Value
RsTemp.Fields("ARCCCW903250").Value = ChkVoltg(13).Value
RsTemp.Fields("ARCCW905100").Value = ChkVoltg(14).Value
RsTemp.Fields("ARCCCW905100").Value = ChkVoltg(15).Value
RsTemp.Fields("GantryPR").Value = ChkVoltg(16).Value
RsTemp.Fields("CollimatorPS").Value = ChkVoltg(17).Value
RsTemp.Fields("UpperJPR").Value = ChkVoltg(18).Value
RsTemp.Fields("LowerJPR").Value = ChkVoltg(19).Value
RsTemp.Fields("EMCOperation").Value = ChkVoltg(20).Value
RsTemp.Fields("Emergency_PSAetr").Value = ChkVoltg(21).Value
RsTemp.Fields("Emergency_WALL").Value = ChkVoltg(22).Value
RsTemp.Fields("Emergency_STAND").Value = ChkVoltg(23).Value
RsTemp.Fields("Emergency_CONSOLE").Value = ChkVoltg(24).Value
RsTemp.Fields("Emergency_Pedante").Value = ChkVoltg(25).Value
RsTemp.Fields("GantryU").Value = ChkVoltg(26).Value
RsTemp.Fields("GantryS").Value = ChkVoltg(27).Value
RsTemp.Fields("GantryLS").Value = ChkVoltg(28).Value

RsTemp.Fields("Hospital").Value = TxtCmps(0).Text
RsTemp.Fields("RismedRep").Value = TxtCmps(1).Text
RsTemp.Fields("Clinac").Value = TxtCmps(2).Text
RsTemp.Fields("FilamenteHours").Value = TxtCmps(3).Text
RsTemp.Fields("RefFSR").Value = TxtCmps(4).Text
RsTemp.Fields("Serial").Value = TxtCmps(5).Text
RsTemp.Fields("BeamHours").Value = TxtCmps(6).Text

RsTemp.Update

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

MsgBox "Guardado!", vbExclamation + vbOKOnly, "Operación Exitosa!"
ChameleonBtn2_Click
Leer_Rismed

End Sub

Private Sub BtnLimpiarCampos_Click()

For i = 0 To TxtCampos.Count - 1
    TxtCampos(i).Text = ""
Next i
End Sub

Private Sub BtnModificar_Click()
On Error Resume Next

If DMGrid1.Rows < 1 Then MsgBox "No existen equipos registrados!", vbInformation + vbOKOnly, "Información": Exit Sub
If DMGrid1.Row < 1 Then MsgBox "Seleccione un registro!", vbInformation + vbOKOnly, "Información": Exit Sub

If Trim(DMGrid1.ValorCelda(DMGrid1.Row, 4)) = "" Then MsgBox "El registro seleccionado no puede ser modificado.", vbExclamation + vbOKOnly, "Información": Exit Sub

CSql = "SELECT * FROM TecnicoEquipo WHERE IdEquipo=" & Val(DMGrid1.ValorCelda(DMGrid1.Row, 4))
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then MsgBox "El registro seleccionado no puede ser modificado.", vbExclamation + vbOKOnly, "Información": Exit Sub

TxtNombre.Text = Trim(RsTemp.Fields("Nombre").Value)
TxtModelo.Text = Trim(RsTemp.Fields("Modelo").Value)
TxtSerial.Text = Trim(RsTemp.Fields("Serial").Value)

Frame2.Visible = True
DMGrid1.Visible = False
Nuevo_Reg = False

End Sub

Private Sub BtnTecnicos_Click()
FrmTecnicos.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnVerReporte_Click()
On Error GoTo werr
Dim Band As Boolean

If DMGrid2.Row = 0 Then MsgBox "Seleccione un reporte!", vbExclamation + vbOKOnly, "Atención": Exit Sub
If Trim(DMGrid2.ValorCelda(DMGrid2.Row, 1)) = "" Then Exit Sub
With CrystalReport1
    .ReportFileName = RutaInformes & "\TecnicoReporte.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "DSN=CrReporte"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{TecnicoReporte.IdCampoEquipo} = " & DMGrid2.ValorCelda(DMGrid2.Row, 4)
    .WindowTitle = "Reporte de Equipos"
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

Exit Sub

werr:
    MsgBox "Operación Fallida!", vbInformation + vbOKOnly, "Información"
End Sub

Private Sub BtnVerReporterismed_Click()
On Error GoTo werr
Dim Band As Boolean

If DMGrid2.Row = 0 Then MsgBox "Seleccione un reporte!", vbExclamation + vbOKOnly, "Atención": Exit Sub
If Trim(DMGrid3.ValorCelda(DMGrid2.Row, 4)) = "" Then Exit Sub

With CrystalReport1
    .ReportFileName = RutaInformes & "\TecnicoReporteRismed.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "DSN=CrReporte"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{TecnicoEquipoReporte2.Id} = " & DMGrid3.ValorCelda(DMGrid3.Row, 4)
    .WindowTitle = "Reporte RISMED "
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

Exit Sub

werr:
    MsgBox "Operación Fallida!", vbInformation + vbOKOnly, "Información"
End Sub

Private Sub ChameleonBtn2_Click()

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps.Count - 1
    If Not Trim(TxtCmps(i).Text) = "N/A" Then TxtCmps(i).Text = ""
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps2.Count - 1
    If Not Trim(TxtCmps2(i).Text) = "N/A" Then TxtCmps2(i).Text = ""
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps3.Count - 1
    If Not Trim(TxtCmps3(i).Text) = "N/A" Then TxtCmps3(i).Text = ""
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To TxtCmps4.Count - 1
    If Not Trim(TxtCmps4(i).Text) = "N/A" Then TxtCmps4(i).Text = ""
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To ChkVoltg.Count - 1
    ChkVoltg(i).Value = 0
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To ChkChks.Count - 1
    ChkChks(i).Value = 0
Next i

End Sub

Private Sub ChameleonBtn3_Click()
Unload FrmBitacoraEquipo
End Sub

Private Sub ChameleonBtn4_Click()

End Sub

Private Sub ChameleonBtn5_Click()
Unload FrmBitacoraEquipo
End Sub

Private Sub ChkChks_Click(Index As Integer)
ChkChks(Index).BackColor = &HEAEFEF
If Index = 11 Then
    If ChkChks(11).Value = 1 Then
        Text1.Text = ""
        Text1.Enabled = True
    Else
        Text1.Text = ""
        Text1.Enabled = False
    End If
ElseIf Index = 12 Then
    If ChkChks(12).Value = 1 Then
        Text2.Text = ""
        Text2.Enabled = True
    Else
        Text2.Text = ""
        Text2.Enabled = False
    End If
End If

End Sub

Private Sub ChkChks_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim iper As Integer
    
    If Index + 1 >= ChkChks.Count Then
        ChkChks(0).SetFocus
    Else
        For iper = Index To ChkChks.Count - 1
            If ChkChks(iper + 1).Enabled Then
                ChkChks(iper + 1).SetFocus
                Exit Sub
            End If
        Next iper
    End If
End If
End Sub

Private Sub ChkOpc_Click(Index As Integer)
On Error Resume Next
For i = 0 To ChkOpc.Count - 1
    If i <> Index Then
        ChkOpc(i).Value = False
        FrmOpc(i).Visible = False
    Else
        ChkOpc(i).Value = True
        FrmOpc(i).Visible = True
    End If
Next
End Sub

Private Sub ChkVoltg_Click(Index As Integer)
ChkVoltg(Index).BackColor = &HEAEFEF
End Sub

Private Sub ChkVoltg_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim iper As Integer
    
    If Index + 1 >= ChkVoltg.Count Then
        ChkVoltg(0).SetFocus
    Else
        For iper = Index To ChkVoltg.Count - 1
            If ChkVoltg(iper + 1).Enabled Then
                ChkVoltg(iper + 1).SetFocus
                Exit Sub
            End If
        Next iper
    End If
End If
End Sub

Private Sub DMGrid1_DobleClick()
BtnModificar_Click
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
On Error GoTo Error
LblNombreEquipo.Caption = DMGrid1.ValorCelda(DMGrid1.Row, 1)
LblReporteEquipo.Caption = DMGrid1.ValorCelda(DMGrid1.Row, 1)

CSql = "SELECT Nombre, Apellidos FROM Usuarios WHERE IdUsuario=" & IdUser
Set RsTemp2 = CrearRS(CSql)
    
TxtCampos(0).Text = Mid(RsTemp2.Fields("Nombre").Value, 1, 1) & Mid(RsTemp2.Fields("Apellidos").Value, 1, 1)
    
Leer_Reportes
DTPicker1.Value = Now

Exit Sub

Error:
    LblNombreEquipo.Caption = "SELECCIONE UN EQUIPO!"
    LblReporteEquipo.Caption = "SELECCIONE UN EQUIPO!"
End Sub

Private Sub DMGrid1_RowColChange(ByVal antRow As Integer, ByVal antCol As Integer, ByVal actRow As Integer, ByVal actCol As Integer)
On Error GoTo Error
LblNombreEquipo.Caption = DMGrid1.ValorCelda(DMGrid1.Row, 1)
LblReporteEquipo.Caption = DMGrid1.ValorCelda(DMGrid1.Row, 1)

CSql = "SELECT Nombre, Apellidos FROM Usuarios WHERE IdUsuario=" & IdUser
Set RsTemp2 = CrearRS(CSql)
    
TxtCampos(0).Text = Mid(RsTemp2.Fields("Nombre").Value, 1, 1) & Mid(RsTemp2.Fields("Apellidos").Value, 1, 1)
    
Leer_Reportes
DTPicker1.Value = Now

Exit Sub

Error:
    LblNombreEquipo.Caption = "SELECCIONE UN EQUIPO!"
    LblReporteEquipo.Caption = "SELECCIONE UN EQUIPO!"
End Sub


Private Sub DMGrid2_DobleClick()
BtnVerReporte_Click
End Sub

Private Sub DMGrid3_DobleClick()
BtnVerReporterismed_Click
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid
Leer_Equipos
Leer_Rismed

Dim RsCboTec As New ADODB.Recordset

CSql = "Select * From Tecnicos"
Set RsCboTec = CrearRS(CSql)

Do While Not RsCboTec.EOF
    CboTecnicos.AddItem RsCboTec.Fields("Nombre").Value & ", " & RsCboTec.Fields("Apellido").Value
    CboTecnicos.ItemData(CboTecnicos.NewIndex) = Val(RsCboTec.Fields("IdTecnicos").Value)
    RsCboTec.MoveNext
Loop


End Sub

Sub LlenarCbo()
Dim RsCboTec As New ADODB.Recordset

CSql = "Select * From Tecnicos"
Set RsCboTec = CrearRS(CSql)

Do While Not RsCboTec.EOF
    CboTecnicos.AddItem RsCboTec.Fields("Nombre").Value & ", " & RsCboTec.Fields("Apellido").Value
    CboTecnicos.ItemData(CboTecnicos.NewIndex) = Val(RsCboTec.Fields("IdTecnicos").Value)
Loop

End Sub

Private Sub Timer1_Timer()
If Weekday(DTPicker1.Value) = 2 Then
    LblDiaSemana.Caption = "DÍA: LUNES"
ElseIf Weekday(DTPicker1.Value) = 3 Then
    LblDiaSemana.Caption = "DÍA: MARTES"
ElseIf Weekday(DTPicker1.Value) = 4 Then
    LblDiaSemana.Caption = "DÍA: MIERCOLES"
ElseIf Weekday(DTPicker1.Value) = 5 Then
    LblDiaSemana.Caption = "DÍA: JUEVES"
ElseIf Weekday(DTPicker1.Value) = 6 Then
    LblDiaSemana.Caption = "DÍA: VIERNES"
ElseIf Weekday(DTPicker1.Value) = 7 Then
    LblDiaSemana.Caption = "DÍA: SABADO"
ElseIf Weekday(DTPicker1.Value) = 1 Then
    LblDiaSemana.Caption = "DÍA: DOMINGO"
End If
End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
If Index <> 0 And Index <> 8 And Index <> 28 And Index <> 24 And Index <> 23 And Index <> 22 And Index <> 21 Then
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If

TxtCampos(Index).BackColor = vbWhite
End Sub

Private Sub TxtCmps_KeyPress(Index As Integer, KeyAscii As Integer)

TxtCmps(Index).BackColor = vbWhite

If KeyAscii = 13 Then
    Dim iper As Integer
    
    If Index + 1 >= TxtCmps.Count Then
        TxtCmps(0).SetFocus
    Else
        For iper = Index To TxtCmps.Count - 1
            If TxtCmps(iper + 1).Enabled Then
                TxtCmps(iper + 1).SetFocus
                Exit Sub
            End If
        Next iper
    End If
End If


End Sub

Private Sub TxtCmps2_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
    Dim iper As Integer
    
    If Index + 1 >= TxtCmps2.Count Then
        TxtCmps2(0).SetFocus
    Else
        For iper = Index To TxtCmps2.Count - 1
            If TxtCmps2(iper + 1).Enabled Then
                TxtCmps2(iper + 1).SetFocus
                Exit Sub
            End If
        Next iper
    End If
End If

If InStr(1, "0123456789,.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
TxtCmps2(Index).BackColor = vbWhite

End Sub

Private Sub TxtCmps3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    Dim iper As Integer
    
    If Index + 1 >= TxtCmps3.Count Then
        TxtCmps3(0).SetFocus
    Else
        For iper = Index To TxtCmps3.Count - 1
            If TxtCmps3(iper + 1).Enabled Then
                TxtCmps3(iper + 1).SetFocus
                Exit Sub
            End If
        Next iper
    End If
End If

If Index <> 5 And Index <> 7 Then
    If InStr(1, "0123456789,.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
    TxtCmps3(Index).BackColor = vbWhite
End If
End Sub

Private Sub TxtCmps4_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
    Dim iper As Integer
    
    If Index + 1 >= TxtCmps4.Count Then
        TxtCmps4(0).SetFocus
    Else
        For iper = Index To TxtCmps4.Count - 1
            If TxtCmps4(iper + 1).Enabled Then
                TxtCmps4(iper + 1).SetFocus
                Exit Sub
            End If
        Next iper
    End If
End If


If InStr(1, "0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 32 Then KeyAscii = 0
    TxtCmps4(Index).BackColor = vbWhite
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
Private Sub TxtSerial_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub


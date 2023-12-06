VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmBraquiterapia 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Braquiterapia"
   ClientHeight    =   10920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "FrmBraquiterapia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10920
   ScaleWidth      =   12270
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Height          =   6735
      Left            =   120
      TabIndex        =   43
      Top             =   3360
      Width           =   12015
      Begin ChamaleonButton.ChameleonBtn BtnVerInforme 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   6240
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "FrmBraquiterapia.frx":1002
         PICN            =   "FrmBraquiterapia.frx":101E
         PICH            =   "FrmBraquiterapia.frx":12BA
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
         Left            =   2040
         TabIndex        =   54
         Top             =   6240
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "FrmBraquiterapia.frx":16FA
         PICN            =   "FrmBraquiterapia.frx":1716
         PICH            =   "FrmBraquiterapia.frx":19A5
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
         Left            =   11280
         TabIndex        =   55
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
         MICON           =   "FrmBraquiterapia.frx":1DE5
         PICN            =   "FrmBraquiterapia.frx":1E01
         PICH            =   "FrmBraquiterapia.frx":2097
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
         TabIndex        =   56
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
         MICON           =   "FrmBraquiterapia.frx":22F6
         PICN            =   "FrmBraquiterapia.frx":2312
         PICH            =   "FrmBraquiterapia.frx":25A7
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
         Left            =   3960
         TabIndex        =   93
         Top             =   6240
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
         MICON           =   "FrmBraquiterapia.frx":2803
         PICN            =   "FrmBraquiterapia.frx":281F
         PICH            =   "FrmBraquiterapia.frx":2AB7
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
         Height          =   5895
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   11775
         Begin VB.Frame FrameAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ORG"
            Height          =   4695
            Index           =   0
            Left            =   5280
            TabIndex        =   74
            Top             =   1080
            Width           =   6375
            Begin VB.TextBox TxtDosisAplicacion1f 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   92
               Top             =   2760
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion1 
               Height          =   315
               Index           =   5
               ItemData        =   "FrmBraquiterapia.frx":2D3E
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2D54
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   2760
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion1e 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   89
               Top             =   2280
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion1 
               Height          =   315
               Index           =   4
               ItemData        =   "FrmBraquiterapia.frx":2D89
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2D9F
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   2280
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion1d 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   86
               Top             =   1800
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion1 
               Height          =   315
               Index           =   3
               ItemData        =   "FrmBraquiterapia.frx":2DD4
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2DEA
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   1800
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion1c 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   83
               Top             =   1320
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion1 
               Height          =   315
               Index           =   2
               ItemData        =   "FrmBraquiterapia.frx":2E1F
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2E35
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion1b 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   80
               Top             =   840
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion1 
               Height          =   315
               Index           =   1
               ItemData        =   "FrmBraquiterapia.frx":2E6A
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2E80
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion1a 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   77
               Top             =   360
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion1 
               Height          =   315
               Index           =   0
               ItemData        =   "FrmBraquiterapia.frx":2EB5
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2ECB
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis:"
               Height          =   195
               Left            =   3000
               TabIndex        =   91
               Top             =   2820
               Width           =   600
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis:"
               Height          =   195
               Left            =   3000
               TabIndex        =   88
               Top             =   2340
               Width           =   600
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis:"
               Height          =   195
               Left            =   3000
               TabIndex        =   85
               Top             =   1860
               Width           =   600
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis:"
               Height          =   195
               Left            =   3000
               TabIndex        =   82
               Top             =   1380
               Width           =   600
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis:"
               Height          =   195
               Left            =   3000
               TabIndex        =   79
               Top             =   900
               Width           =   600
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis:"
               Height          =   195
               Left            =   3000
               TabIndex        =   76
               Top             =   420
               Width           =   600
            End
         End
         Begin VB.Frame FrameAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ORG"
            Height          =   4575
            Index           =   3
            Left            =   5280
            TabIndex        =   270
            Top             =   1080
            Width           =   6375
            Begin VB.ComboBox CboAplicacion4 
               Height          =   315
               Index           =   0
               ItemData        =   "FrmBraquiterapia.frx":2F00
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2F16
               Style           =   2  'Dropdown List
               TabIndex        =   282
               Top             =   360
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion4a 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   281
               Top             =   360
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion4 
               Height          =   315
               Index           =   1
               ItemData        =   "FrmBraquiterapia.frx":2F4B
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2F61
               Style           =   2  'Dropdown List
               TabIndex        =   280
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion4b 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   279
               Top             =   840
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion4 
               Height          =   315
               Index           =   2
               ItemData        =   "FrmBraquiterapia.frx":2F96
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2FAC
               Style           =   2  'Dropdown List
               TabIndex        =   278
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion4c 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   277
               Top             =   1320
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion4 
               Height          =   315
               Index           =   3
               ItemData        =   "FrmBraquiterapia.frx":2FE1
               Left            =   120
               List            =   "FrmBraquiterapia.frx":2FF7
               Style           =   2  'Dropdown List
               TabIndex        =   276
               Top             =   1800
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion4d 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   275
               Top             =   1800
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion4 
               Height          =   315
               Index           =   4
               ItemData        =   "FrmBraquiterapia.frx":302C
               Left            =   120
               List            =   "FrmBraquiterapia.frx":3042
               Style           =   2  'Dropdown List
               TabIndex        =   274
               Top             =   2280
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion4e 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   273
               Top             =   2280
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion4 
               Height          =   315
               Index           =   5
               ItemData        =   "FrmBraquiterapia.frx":3077
               Left            =   120
               List            =   "FrmBraquiterapia.frx":308D
               Style           =   2  'Dropdown List
               TabIndex        =   272
               Top             =   2760
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion4f 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   271
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   288
               Top             =   420
               Width           =   555
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   287
               Top             =   900
               Width           =   555
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   286
               Top             =   1380
               Width           =   555
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   285
               Top             =   1860
               Width           =   555
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   284
               Top             =   2340
               Width           =   555
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   283
               Top             =   2820
               Width           =   555
            End
         End
         Begin VB.Frame FrameAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ORG"
            Height          =   4575
            Index           =   2
            Left            =   5280
            TabIndex        =   251
            Top             =   1080
            Width           =   6375
            Begin VB.ComboBox CboAplicacion3 
               Height          =   315
               Index           =   0
               ItemData        =   "FrmBraquiterapia.frx":30C2
               Left            =   120
               List            =   "FrmBraquiterapia.frx":30D8
               Style           =   2  'Dropdown List
               TabIndex        =   263
               Top             =   360
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion3a 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   262
               Top             =   360
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion3 
               Height          =   315
               Index           =   1
               ItemData        =   "FrmBraquiterapia.frx":310D
               Left            =   120
               List            =   "FrmBraquiterapia.frx":3123
               Style           =   2  'Dropdown List
               TabIndex        =   261
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion3b 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   260
               Top             =   840
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion3 
               Height          =   315
               Index           =   2
               ItemData        =   "FrmBraquiterapia.frx":3158
               Left            =   120
               List            =   "FrmBraquiterapia.frx":316E
               Style           =   2  'Dropdown List
               TabIndex        =   259
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion3c 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   258
               Top             =   1320
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion3 
               Height          =   315
               Index           =   3
               ItemData        =   "FrmBraquiterapia.frx":31A3
               Left            =   120
               List            =   "FrmBraquiterapia.frx":31B9
               Style           =   2  'Dropdown List
               TabIndex        =   257
               Top             =   1800
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion3d 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   256
               Top             =   1800
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion3 
               Height          =   315
               Index           =   4
               ItemData        =   "FrmBraquiterapia.frx":31EE
               Left            =   120
               List            =   "FrmBraquiterapia.frx":3204
               Style           =   2  'Dropdown List
               TabIndex        =   255
               Top             =   2280
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion3e 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   254
               Top             =   2280
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion3 
               Height          =   315
               Index           =   5
               ItemData        =   "FrmBraquiterapia.frx":3239
               Left            =   120
               List            =   "FrmBraquiterapia.frx":324F
               Style           =   2  'Dropdown List
               TabIndex        =   253
               Top             =   2760
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion3f 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   252
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   269
               Top             =   420
               Width           =   555
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   268
               Top             =   900
               Width           =   555
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   267
               Top             =   1380
               Width           =   555
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   266
               Top             =   1860
               Width           =   555
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   265
               Top             =   2340
               Width           =   555
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   264
               Top             =   2820
               Width           =   555
            End
         End
         Begin VB.Frame FrameAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "ORG"
            Height          =   4695
            Index           =   1
            Left            =   5280
            TabIndex        =   94
            Top             =   1080
            Width           =   6375
            Begin VB.ComboBox CboAplicacion2 
               Height          =   315
               Index           =   0
               ItemData        =   "FrmBraquiterapia.frx":3284
               Left            =   120
               List            =   "FrmBraquiterapia.frx":329A
               Style           =   2  'Dropdown List
               TabIndex        =   106
               Top             =   360
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion2a 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   105
               Top             =   360
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion2 
               Height          =   315
               Index           =   1
               ItemData        =   "FrmBraquiterapia.frx":32CF
               Left            =   120
               List            =   "FrmBraquiterapia.frx":32E5
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion2b 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   103
               Top             =   840
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion2 
               Height          =   315
               Index           =   2
               ItemData        =   "FrmBraquiterapia.frx":331A
               Left            =   120
               List            =   "FrmBraquiterapia.frx":3330
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion2c 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   101
               Top             =   1320
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion2 
               Height          =   315
               Index           =   3
               ItemData        =   "FrmBraquiterapia.frx":3365
               Left            =   120
               List            =   "FrmBraquiterapia.frx":337B
               Style           =   2  'Dropdown List
               TabIndex        =   100
               Top             =   1800
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion2d 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   99
               Top             =   1800
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion2 
               Height          =   315
               Index           =   4
               ItemData        =   "FrmBraquiterapia.frx":33B0
               Left            =   120
               List            =   "FrmBraquiterapia.frx":33C6
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   2280
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion2e 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   97
               Top             =   2280
               Width           =   975
            End
            Begin VB.ComboBox CboAplicacion2 
               Height          =   315
               Index           =   5
               ItemData        =   "FrmBraquiterapia.frx":33FB
               Left            =   120
               List            =   "FrmBraquiterapia.frx":3411
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   2760
               Width           =   2655
            End
            Begin VB.TextBox TxtDosisAplicacion2f 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3600
               TabIndex        =   95
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   112
               Top             =   420
               Width           =   555
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   111
               Top             =   900
               Width           =   555
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   110
               Top             =   1380
               Width           =   555
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   109
               Top             =   1860
               Width           =   555
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   108
               Top             =   2340
               Width           =   555
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Dosis"
               Height          =   195
               Left            =   3000
               TabIndex        =   107
               Top             =   2820
               Width           =   555
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Sedación"
            Height          =   615
            Left            =   120
            TabIndex        =   289
            Top             =   1800
            Width           =   2655
            Begin VB.OptionButton OptSedacion 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Si"
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   291
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton OptSedacion 
               BackColor       =   &H00EAEFEF&
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   290
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.OptionButton OptAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Aplicación 4"
            Height          =   495
            Index           =   3
            Left            =   9240
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OptAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Aplicación 2"
            Height          =   495
            Index           =   1
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OptAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Aplicación 1"
            Height          =   495
            Index           =   0
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptAplicacion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Aplicación 3"
            Height          =   495
            Index           =   2
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox TxtSesiones 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3960
            TabIndex        =   61
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox TxtDosisFraccion 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1440
            TabIndex        =   60
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox CboSubModalidad 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   480
            Width           =   2415
         End
         Begin VB.ComboBox CboModalidad 
            Height          =   315
            ItemData        =   "FrmBraquiterapia.frx":3446
            Left            =   120
            List            =   "FrmBraquiterapia.frx":3448
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   480
            Width           =   2415
         End
         Begin VB.ComboBox CboAplicador 
            Height          =   315
            ItemData        =   "FrmBraquiterapia.frx":344A
            Left            =   960
            List            =   "FrmBraquiterapia.frx":344C
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Farmaco:"
            Height          =   3375
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   2400
            Width           =   4935
            Begin VB.ComboBox CboFarmaco8 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   2880
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco7 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   2520
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco6 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   2160
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco5 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   66
               Top             =   1800
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco4 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   1440
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco3 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   1080
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco2 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   720
               Width           =   4575
            End
            Begin VB.ComboBox CboFarmaco1 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   62
               Top             =   360
               Width           =   4575
            End
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "cGy"
            Height          =   195
            Left            =   2640
            TabIndex        =   52
            Top             =   1050
            Width           =   285
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicador:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1500
            Width           =   705
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dosis / Fracción:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nuevo Informe"
            Height          =   195
            Left            =   3720
            TabIndex        =   49
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sesiones:"
            Height          =   195
            Left            =   3120
            TabIndex        =   48
            Top             =   1050
            Width           =   690
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Complicaciones"
         Height          =   5895
         Index           =   1
         Left            =   120
         TabIndex        =   113
         Top             =   240
         Width           =   11775
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   0
            Left            =   120
            TabIndex        =   114
            ToolTipText     =   "Complicaciones Agudas"
            Top             =   840
            Width           =   11535
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   12
               ItemData        =   "FrmBraquiterapia.frx":344E
               Left            =   8280
               List            =   "FrmBraquiterapia.frx":3465
               Style           =   2  'Dropdown List
               TabIndex        =   140
               Top             =   1980
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   12
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   139
               Top             =   1995
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   11
               ItemData        =   "FrmBraquiterapia.frx":3498
               Left            =   8280
               List            =   "FrmBraquiterapia.frx":34AF
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   11
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   137
               Top             =   1575
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   10
               ItemData        =   "FrmBraquiterapia.frx":34E2
               Left            =   8280
               List            =   "FrmBraquiterapia.frx":34F9
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   10
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   135
               Top             =   1215
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   9
               ItemData        =   "FrmBraquiterapia.frx":352C
               Left            =   8280
               List            =   "FrmBraquiterapia.frx":3543
               Style           =   2  'Dropdown List
               TabIndex        =   134
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   9
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   133
               Top             =   855
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               ItemData        =   "FrmBraquiterapia.frx":3576
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":358D
               Style           =   2  'Dropdown List
               TabIndex        =   132
               Top             =   3900
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   8
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   131
               Top             =   3915
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               ItemData        =   "FrmBraquiterapia.frx":35C0
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":35D7
               Style           =   2  'Dropdown List
               TabIndex        =   130
               Top             =   3000
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               ItemData        =   "FrmBraquiterapia.frx":360A
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":3621
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   3420
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   128
               Top             =   3435
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   127
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               ItemData        =   "FrmBraquiterapia.frx":3654
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":366B
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   125
               Top             =   1200
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               ItemData        =   "FrmBraquiterapia.frx":369E
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":36B5
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   123
               Top             =   1560
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               ItemData        =   "FrmBraquiterapia.frx":36E8
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":36FF
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   1920
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   121
               Top             =   1920
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               ItemData        =   "FrmBraquiterapia.frx":3732
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":3749
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   2280
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   119
               Top             =   2280
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               ItemData        =   "FrmBraquiterapia.frx":377C
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":3793
               Style           =   2  'Dropdown List
               TabIndex        =   118
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   117
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   6
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   116
               Top             =   3000
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               ItemData        =   "FrmBraquiterapia.frx":37C6
               Left            =   3000
               List            =   "FrmBraquiterapia.frx":37DD
               Style           =   2  'Dropdown List
               TabIndex        =   115
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   1
               Left            =   9720
               TabIndex        =   159
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   1
               Left            =   8640
               TabIndex        =   158
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Órgano o Tejido"
               Height          =   195
               Index           =   1
               Left            =   6660
               TabIndex        =   157
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   0
               Left            =   4440
               TabIndex        =   156
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   0
               Left            =   3360
               TabIndex        =   155
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Órgano o Tejido"
               Height          =   195
               Index           =   0
               Left            =   1440
               TabIndex        =   154
               Top             =   360
               Width           =   1155
            End
            Begin VB.Line Line8 
               X1              =   6120
               X2              =   6120
               Y1              =   360
               Y2              =   4680
            End
            Begin VB.Line Line9 
               X1              =   120
               X2              =   11400
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tracto Gastro Intestinal Inferior Incluyendo Pelvis"
               Height          =   435
               Left            =   120
               TabIndex        =   153
               Top             =   3840
               Width           =   2775
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tracto Gastro Intestinal Superior"
               Height          =   195
               Left            =   615
               TabIndex        =   152
               Top             =   3480
               Width           =   2280
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Membranas Mucosas"
               Height          =   315
               Left            =   120
               TabIndex        =   151
               Top             =   1200
               Width           =   2775
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ojo"
               Height          =   195
               Left            =   2655
               TabIndex        =   150
               Top             =   1620
               Width           =   240
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Oido"
               Height          =   195
               Left            =   2565
               TabIndex        =   149
               Top             =   1980
               Width           =   330
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glándulas Salivales"
               Height          =   195
               Left            =   1515
               TabIndex        =   148
               Top             =   2340
               Width           =   1380
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Faringe Esófago"
               Height          =   195
               Left            =   1740
               TabIndex        =   147
               Top             =   2700
               Width           =   1155
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Laringe"
               Height          =   195
               Left            =   2370
               TabIndex        =   146
               Top             =   3060
               Width           =   525
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piel"
               Height          =   195
               Left            =   2640
               TabIndex        =   145
               Top             =   900
               Width           =   255
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sistema Nervioso Central"
               Height          =   195
               Index           =   0
               Left            =   6360
               TabIndex        =   144
               Top             =   2040
               Width           =   1770
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corazón"
               Height          =   195
               Left            =   7545
               TabIndex        =   143
               Top             =   1620
               Width           =   585
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Genitourinario"
               Height          =   195
               Left            =   6360
               TabIndex        =   142
               Top             =   1260
               Width           =   1770
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulmón"
               Height          =   195
               Left            =   7605
               TabIndex        =   141
               Top             =   900
               Width           =   525
            End
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Complicaciones Agudas"
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   250
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Toxicidad Hematológica Aguda"
            Height          =   495
            Index           =   1
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   249
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
            TabIndex        =   248
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Opciones"
            Height          =   495
            Index           =   3
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   247
            Top             =   360
            Width           =   2415
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   2
            Left            =   120
            TabIndex        =   160
            ToolTipText     =   "Complicaciones Crónicas"
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtDosisGrdo 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   18
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   194
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   18
               ItemData        =   "FrmBraquiterapia.frx":3810
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3827
               Style           =   2  'Dropdown List
               TabIndex        =   193
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   19
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   192
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   19
               ItemData        =   "FrmBraquiterapia.frx":385A
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3871
               Style           =   2  'Dropdown List
               TabIndex        =   191
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   20
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   190
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   20
               ItemData        =   "FrmBraquiterapia.frx":38A4
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":38BB
               Style           =   2  'Dropdown List
               TabIndex        =   189
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   21
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   188
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   21
               ItemData        =   "FrmBraquiterapia.frx":38EE
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3905
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   22
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   186
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   22
               ItemData        =   "FrmBraquiterapia.frx":3938
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":394F
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   23
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   184
               Top             =   2535
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   23
               ItemData        =   "FrmBraquiterapia.frx":3982
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3999
               Style           =   2  'Dropdown List
               TabIndex        =   183
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   24
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   182
               Top             =   2895
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   24
               ItemData        =   "FrmBraquiterapia.frx":39CC
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":39E3
               Style           =   2  'Dropdown List
               TabIndex        =   181
               Top             =   2880
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   25
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   180
               Top             =   3255
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   25
               ItemData        =   "FrmBraquiterapia.frx":3A16
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3A2D
               Style           =   2  'Dropdown List
               TabIndex        =   179
               Top             =   3240
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   26
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   178
               Top             =   3615
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   26
               ItemData        =   "FrmBraquiterapia.frx":3A60
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3A77
               Style           =   2  'Dropdown List
               TabIndex        =   177
               Top             =   3600
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   27
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   176
               Top             =   3975
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   27
               ItemData        =   "FrmBraquiterapia.frx":3AAA
               Left            =   2520
               List            =   "FrmBraquiterapia.frx":3AC1
               Style           =   2  'Dropdown List
               TabIndex        =   175
               Top             =   3960
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   28
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   174
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   28
               ItemData        =   "FrmBraquiterapia.frx":3AF4
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3B0B
               Style           =   2  'Dropdown List
               TabIndex        =   173
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   29
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   172
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   29
               ItemData        =   "FrmBraquiterapia.frx":3B3E
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3B55
               Style           =   2  'Dropdown List
               TabIndex        =   171
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   30
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   170
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   30
               ItemData        =   "FrmBraquiterapia.frx":3B88
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3B9F
               Style           =   2  'Dropdown List
               TabIndex        =   169
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   31
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   168
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   31
               ItemData        =   "FrmBraquiterapia.frx":3BD2
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3BE9
               Style           =   2  'Dropdown List
               TabIndex        =   167
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   32
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   166
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   32
               ItemData        =   "FrmBraquiterapia.frx":3C1C
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3C33
               Style           =   2  'Dropdown List
               TabIndex        =   165
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   33
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   164
               Top             =   2535
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   33
               ItemData        =   "FrmBraquiterapia.frx":3C66
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3C7D
               Style           =   2  'Dropdown List
               TabIndex        =   163
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   34
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   162
               Top             =   2895
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   34
               ItemData        =   "FrmBraquiterapia.frx":3CB0
               Left            =   8040
               List            =   "FrmBraquiterapia.frx":3CC7
               Style           =   2  'Dropdown List
               TabIndex        =   161
               Top             =   2880
               Width           =   1335
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Intestino Grueso Delgado"
               Height          =   195
               Left            =   5400
               TabIndex        =   217
               Top             =   1140
               Width           =   2340
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corazón"
               Height          =   195
               Left            =   1755
               TabIndex        =   216
               Top             =   4020
               Width           =   585
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Esófago"
               Height          =   195
               Left            =   5400
               TabIndex        =   215
               Top             =   780
               Width           =   2340
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulmón"
               Height          =   195
               Left            =   1815
               TabIndex        =   214
               Top             =   3660
               Width           =   525
            End
            Begin VB.Line Line11 
               X1              =   120
               X2              =   11400
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line12 
               X1              =   5760
               X2              =   5760
               Y1              =   360
               Y2              =   4680
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laringe"
               Height          =   195
               Left            =   120
               TabIndex        =   213
               Top             =   3300
               Width           =   2220
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ojo"
               Height          =   195
               Left            =   120
               TabIndex        =   212
               Top             =   2940
               Width           =   2220
            End
            Begin VB.Label Label104 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piel"
               Height          =   195
               Left            =   2085
               TabIndex        =   211
               Top             =   780
               Width           =   255
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tejido Subcutáneo"
               Height          =   195
               Left            =   120
               TabIndex        =   210
               Top             =   1140
               Width           =   2220
            End
            Begin VB.Label Label106 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Membranas Mucosas"
               Height          =   195
               Left            =   120
               TabIndex        =   209
               Top             =   1500
               Width           =   2220
            End
            Begin VB.Label Label108 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Glándulas Salivales"
               Height          =   195
               Left            =   120
               TabIndex        =   208
               Top             =   1860
               Width           =   2220
            End
            Begin VB.Label Label109 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Médula Espinal"
               Height          =   195
               Left            =   120
               TabIndex        =   207
               Top             =   2220
               Width           =   2220
            End
            Begin VB.Label Label110 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cerebro"
               Height          =   195
               Left            =   1785
               TabIndex        =   206
               Top             =   2580
               Width           =   555
            End
            Begin VB.Label Label107 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hígado"
               Height          =   195
               Left            =   5400
               TabIndex        =   205
               Top             =   1500
               Width           =   2340
            End
            Begin VB.Label Label111 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Riñon"
               Height          =   195
               Left            =   5400
               TabIndex        =   204
               Top             =   1860
               Width           =   2340
            End
            Begin VB.Label Label112 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Vejiga"
               Height          =   195
               Left            =   5400
               TabIndex        =   203
               Top             =   2220
               Width           =   2340
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hueso"
               Height          =   195
               Left            =   5400
               TabIndex        =   202
               Top             =   2580
               Width           =   2340
            End
            Begin VB.Label Label114 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Articulación"
               Height          =   195
               Left            =   5400
               TabIndex        =   201
               Top             =   2940
               Width           =   2340
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organo o Tejido"
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   200
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   3
               Left            =   2880
               TabIndex        =   199
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo / Meses  "
               Height          =   195
               Index           =   3
               Left            =   3990
               TabIndex        =   198
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organo o Tejido"
               Height          =   315
               Index           =   3
               Left            =   6480
               TabIndex        =   197
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   4
               Left            =   8400
               TabIndex        =   196
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo / Meses  "
               Height          =   195
               Index           =   4
               Left            =   9510
               TabIndex        =   195
               Top             =   360
               Width           =   1275
            End
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   1
            Left            =   120
            TabIndex        =   229
            ToolTipText     =   "Toxicidad Hematológica Aguda"
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   13
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   239
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   13
               ItemData        =   "FrmBraquiterapia.frx":3CFA
               Left            =   2880
               List            =   "FrmBraquiterapia.frx":3D11
               Style           =   2  'Dropdown List
               TabIndex        =   238
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   14
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   237
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   14
               ItemData        =   "FrmBraquiterapia.frx":3D44
               Left            =   2880
               List            =   "FrmBraquiterapia.frx":3D5B
               Style           =   2  'Dropdown List
               TabIndex        =   236
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   15
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   235
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   15
               ItemData        =   "FrmBraquiterapia.frx":3D8E
               Left            =   2880
               List            =   "FrmBraquiterapia.frx":3DA5
               Style           =   2  'Dropdown List
               TabIndex        =   234
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   16
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   233
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   16
               ItemData        =   "FrmBraquiterapia.frx":3DD8
               Left            =   2880
               List            =   "FrmBraquiterapia.frx":3DEF
               Style           =   2  'Dropdown List
               TabIndex        =   232
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   17
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   231
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   17
               ItemData        =   "FrmBraquiterapia.frx":3E22
               Left            =   2880
               List            =   "FrmBraquiterapia.frx":3E39
               Style           =   2  'Dropdown List
               TabIndex        =   230
               Top             =   2160
               Width           =   1335
            End
            Begin VB.Line Line10 
               X1              =   120
               X2              =   5880
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label98 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hematocrito (%)"
               Height          =   195
               Left            =   855
               TabIndex        =   246
               Top             =   2160
               Width           =   1110
            End
            Begin VB.Label Label97 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hemoglobina (g/dl)"
               Height          =   195
               Left            =   735
               TabIndex        =   245
               Top             =   1800
               Width           =   1350
            End
            Begin VB.Label Label96 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Neutrófilos (x10 ³/ml)"
               Height          =   195
               Left            =   675
               TabIndex        =   244
               Top             =   1440
               Width           =   1470
            End
            Begin VB.Label Label95 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plaquetas (x10 ³/ml)"
               Height          =   195
               Left            =   705
               TabIndex        =   243
               Top             =   1080
               Width           =   1410
            End
            Begin VB.Label Label94 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glóbulos Blancos (x10 ³/ml)"
               Height          =   195
               Left            =   435
               TabIndex        =   242
               Top             =   720
               Width           =   1950
            End
            Begin VB.Label Label51 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   2
               Left            =   2880
               TabIndex        =   241
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   240
               Top             =   360
               Width           =   840
            End
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   3
            Left            =   120
            TabIndex        =   218
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtOtrasObs 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   2655
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   219
               Top             =   1320
               Width           =   5175
            End
            Begin SystemOncoAmerica.DMGrid DMGrid1 
               Height          =   3255
               Left            =   5880
               TabIndex        =   220
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
               TabIndex        =   221
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
               MICON           =   "FrmBraquiterapia.frx":3E6C
               PICN            =   "FrmBraquiterapia.frx":3E88
               PICH            =   "FrmBraquiterapia.frx":4015
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
               TabIndex        =   222
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
               MICON           =   "FrmBraquiterapia.frx":424A
               PICN            =   "FrmBraquiterapia.frx":4266
               PICH            =   "FrmBraquiterapia.frx":44F5
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
               TabIndex        =   223
               Top             =   390
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   51380225
               CurrentDate     =   40371
            End
            Begin ChamaleonButton.ChameleonBtn BtnEliminarComplica 
               Height          =   375
               Left            =   9600
               TabIndex        =   224
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
               MICON           =   "FrmBraquiterapia.frx":4936
               PICN            =   "FrmBraquiterapia.frx":4952
               PICH            =   "FrmBraquiterapia.frx":4AF6
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
               TabIndex        =   225
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
               MICON           =   "FrmBraquiterapia.frx":4C95
               PICN            =   "FrmBraquiterapia.frx":4CB1
               PICH            =   "FrmBraquiterapia.frx":4E55
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Listado general"
               Height          =   195
               Left            =   6000
               TabIndex        =   228
               Top             =   360
               Width           =   1080
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otras observaciones:"
               Height          =   195
               Left            =   240
               TabIndex        =   227
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de creación del informe:"
               Height          =   195
               Left            =   480
               TabIndex        =   226
               Top             =   480
               Width           =   2190
            End
         End
      End
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Prescripciones"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2880
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Complicaciones"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   3840
      TabIndex        =   35
      Top             =   10080
      Width           =   8295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7200
         TabIndex        =   36
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
         MICON           =   "FrmBraquiterapia.frx":4FF4
         PICN            =   "FrmBraquiterapia.frx":5010
         PICH            =   "FrmBraquiterapia.frx":51D9
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
         TabIndex        =   37
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
         MICON           =   "FrmBraquiterapia.frx":540E
         PICN            =   "FrmBraquiterapia.frx":542A
         PICH            =   "FrmBraquiterapia.frx":56B9
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
         TabIndex        =   38
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
         MICON           =   "FrmBraquiterapia.frx":5AFA
         PICN            =   "FrmBraquiterapia.frx":5B16
         PICH            =   "FrmBraquiterapia.frx":5CA3
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
         TabIndex        =   39
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
         MICON           =   "FrmBraquiterapia.frx":5ED8
         PICN            =   "FrmBraquiterapia.frx":5EF4
         PICH            =   "FrmBraquiterapia.frx":61D6
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
         TabIndex        =   40
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
         MICON           =   "FrmBraquiterapia.frx":6427
         PICN            =   "FrmBraquiterapia.frx":6443
         PICH            =   "FrmBraquiterapia.frx":65E7
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
         Left            =   5040
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
         Left            =   4440
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   32
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
         TabIndex        =   33
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad o Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   34
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
         MICON           =   "FrmBraquiterapia.frx":6786
         PICN            =   "FrmBraquiterapia.frx":67A2
         PICH            =   "FrmBraquiterapia.frx":6A07
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3360
         Top             =   240
      End
      Begin ChamaleonButton.ChameleonBtn BtnLlamar 
         Height          =   375
         Left            =   5880
         TabIndex        =   9
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
         MICON           =   "FrmBraquiterapia.frx":6C99
         PICN            =   "FrmBraquiterapia.frx":6CB5
         PICH            =   "FrmBraquiterapia.frx":6F51
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
         TabIndex        =   10
         ToolTipText     =   "Lista de Espera"
         Top             =   1680
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
         MICON           =   "FrmBraquiterapia.frx":7186
         PICN            =   "FrmBraquiterapia.frx":71A2
         PICH            =   "FrmBraquiterapia.frx":742B
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
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaNac 
         Height          =   375
         Left            =   8400
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   375
         Left            =   8400
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   375
         Left            =   8400
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   40121
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
         Height          =   375
         Left            =   11040
         TabIndex        =   15
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
         MICON           =   "FrmBraquiterapia.frx":76C3
         PICN            =   "FrmBraquiterapia.frx":76DF
         PICH            =   "FrmBraquiterapia.frx":7975
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
         TabIndex        =   16
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
         MICON           =   "FrmBraquiterapia.frx":7BD4
         PICN            =   "FrmBraquiterapia.frx":7BF0
         PICH            =   "FrmBraquiterapia.frx":7E85
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
         TabIndex        =   17
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
         MICON           =   "FrmBraquiterapia.frx":80E1
         PICN            =   "FrmBraquiterapia.frx":80FD
         PICH            =   "FrmBraquiterapia.frx":82A1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         Height          =   195
         Left            =   5280
         TabIndex        =   31
         Top             =   810
         Width           =   405
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad:"
         Height          =   195
         Left            =   5280
         TabIndex        =   30
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nac.:"
         Height          =   195
         Left            =   7200
         TabIndex        =   29
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Remitente:"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de C&ulminación:"
         Height          =   375
         Left            =   7320
         TabIndex        =   27
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Inicio:"
         Height          =   195
         Left            =   7200
         TabIndex        =   26
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Tratante:"
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellido(s):"
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre(s):"
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro:"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula:"
         Height          =   195
         Left            =   945
         TabIndex        =   21
         Top             =   330
         Width           =   540
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
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   9960
         Picture         =   "FrmBraquiterapia.frx":84D6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Historia:"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   330
         Width           =   870
      End
      Begin VB.Label NoReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro "
         Height          =   195
         Left            =   7200
         TabIndex        =   18
         Top             =   2250
         Width           =   630
      End
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   375
      Left            =   10680
      TabIndex        =   41
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51380225
      CurrentDate     =   39801
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
Attribute VB_Name = "FrmBraquiterapia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarPacientes As New ADODB.Recordset
Dim RsBraqui As New ADODB.Recordset
Dim Cambio
Dim actualiza
Dim IdReg
Dim MaxReg As String
Dim Gram As TextoPos
Dim Sesiones As Double
Dim MaxRegP As String
Dim RsIdMax As New ADODB.Recordset
Dim RsRegPendiente As New ADODB.Recordset
Dim IdInf
Dim OptBordesSel As String
Dim RsTemp As New ADODB.Recordset
Dim i As Integer
Dim IdRegFinal
Dim Sedacion
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
 'MMMM Numero de semanas MMMMM

 'MsgBox (DatePart("ww", DateValue(Now)) + 1) - DatePart("ww", "01/01/" & Year(Now))

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
  " WHERE Informe_Medico5.IdInforme='" & IdInf & "'" ' ORDER BY ID"

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

'If RsTemp.RecordCount <> 0 Then

'    If Trim(RsTemp.Fields("CP").Value) <> "" Then
'        For i = 0 To Combo3.ListCount - 1
'            If Trim(RsTemp.Fields("CP").Value) = Combo3.List(i) Then
'                Combo3.ListIndex = i
'                Exit For
'            End If
'            Combo3.ListIndex = -1
'        Next i
'    End If
'
'    If Trim(RsTemp.Fields("T").Value) <> "" Then
'        For i = 0 To Combo4.ListCount - 1
'            If Trim(RsTemp.Fields("T").Value) = Combo4.List(i) Then
'                Combo4.ListIndex = i
'                Exit For
'            End If
'            Combo4.ListIndex = -1
'        Next i
'    End If
'
'    If Trim(RsTemp.Fields("N").Value) <> "" Then
'        For i = 0 To Combo5.ListCount - 1
'            If Trim(RsTemp.Fields("N").Value) = Combo5.List(i) Then
'                Combo5.ListIndex = i
'                Exit For
'            End If
'            Combo5.ListIndex = -1
'        Next i
'    End If
'
'    If Trim(RsTemp.Fields("M").Value) <> "" Then
'        For i = 0 To Combo6.ListCount - 1
'            If Trim(RsTemp.Fields("M").Value) = Combo6.List(i) Then
'                Combo6.ListIndex = i
'                Exit For
'            End If
'            Combo6.ListIndex = -1
'        Next i
'    End If
'
'    If Trim(RsTemp.Fields("Estadio").Value) <> "" Then
'        For i = 0 To Combo2.ListCount - 1
'            If Trim(RsTemp.Fields("Estadio").Value) = Combo2.List(i) Then
'                Combo2.ListIndex = i
'                Exit For
'            End If
'            Combo2.ListIndex = -1
'        Next i
'    End If
'
'    If Trim(RsTemp.Fields("G").Value) <> "" Then
'        For i = 0 To Combo7.ListCount - 1
'            If Trim(RsTemp.Fields("G").Value) = Combo7.List(i) Then
'                Combo7.ListIndex = i
'                Exit For
'            End If
'            Combo7.ListIndex = -1
'        Next i
'    End If
'
'    If Trim(RsTemp.Fields("Gleason").Value) <> "" Then
'        TxtGleason.Text = Trim(RsTemp.Fields("Gleason").Value)
'        ChkGleason.Value = 1
'    Else
'        TxtGleason.Text = ""
'        ChkGleason.Value = 0
'    End If
    
'    OptBordes(0).Value = True
'    OptBordes(0).Value = False
'    For i = 0 To OptBordes.Count - 1
'        If OptBordes(i).Caption = Trim(RsTemp.Fields("Reseccion").Value) Then
'            OptBordes(i).Value = True
'            Exit For
'        End If
'    Next i
'
'    For i = 0 To CboTCancers.ListCount - 1
'        If CboTCancers.ItemData(i) = RsTemp.Fields("IdTipoCancer").Value Then
'            CboTCancers.ListIndex = i
'            Exit For
'        End If
'        CboTCancers.ListIndex = -1
'    Next i
    
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'    For i = 0 To ChkGlobal.Count - 1
'        ChkGlobal(i).BackColor = &HEAEFEF
'        ChkGlobal(i).Value = 0
'    Next i
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    ' Ciclo que carga TODOS los valores para la "Gran Matriz" para la inmonohistoquimica
'    For i = 0 To ChkGlobal.Count - 1
'        If Trim(RsTemp.Fields(i).Value) <> "" Then
'
'            If i <= 56 Then
'                For J = 0 To CboGeneral(i).ListCount - 1
'                    If Trim(CboGeneral(i).List(J)) = Trim(RsTemp.Fields(i).Value) Then
'                        CboGeneral(i).ListIndex = J
'                        ChkGlobal(i).Value = 1
'                        Exit For
'                    End If
'                Next J
'            Else
'                If Trim(RsTemp.Fields("OTROS").Value) <> "" Then
'                    TxtOtros.Text = Trim(RsTemp.Fields("OTROS").Value)
'                    ChkGlobal(57).Value = 1
'                End If
'            End If
'        Else
'            If i <= 56 Then
'                ChkGlobal(i).Value = 0
'                CboGeneral(i).ListIndex = -1
'            Else
'                ChkGlobal(57).Value = 0
'                TxtOtros.Text = ""
'            End If
'        End If
'    Next i
'Else
'    For i = 0 To ChkGlobal.Count - 1
'        If i <= 56 Then
'            ChkGlobal(i).Value = 0
'            CboGeneral(i).ListIndex = -1
'        Else
'            ChkGlobal(57).Value = 0
'            TxtOtros.Text = ""
'        End If
'    Next i
'End If
'
'For i = 0 To ChkMuert.Count - 1
'    ChkMuert(i).BackColor = &HEAEFEF
'    ChkMuert(i).Value = 0
'Next i
'For i = 0 To ChkRecaida.Count - 1
'    ChkRecaida(i).BackColor = &HEAEFEF
'    ChkRecaida(i).Value = 0
'Next i
'For i = 0 To LstProg.Count - 1
'    LstProg(i).BackColor = &HEAEFEF
'    LstProg(i).Value = 0
'Next i
'For i = 0 To LstEnfer.Count - 1
'    LstEnfer(i).BackColor = &HEAEFEF
'    LstEnfer(i).Value = 0
'Next i
'
'For i = 0 To Text5.Count - 1
'    Text5(i).Text = ""
'Next i

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

'If RsTemp.RecordCount <> 0 Then
    
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'    For i = 1 To LstEnfer.Count - 2
'        If RsTemp.Fields(i).Value Then
'            LstEnfer(i).Value = 1
'        End If
'    Next i
'
'    If Trim(RsTemp.Fields("LE_MaxM").Value) <> "" Then
'        LstEnfer(LstEnfer.Count - 1).Value = 1
'        Text5(1).Text = Trim(RsTemp.Fields("LE_MaxM").Value)
'    End If
'    If Trim(RsTemp.Fields("LE_MinM").Value) <> "" Then
'        LstEnfer(0).Value = 1
'        Text5(0).Text = Trim(RsTemp.Fields("LE_MinM").Value)
'    End If
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  ' 13-22
'    For i = 1 To LstProg.Count - 2
'        If RsTemp.Fields(i + 12).Value Then
'            LstProg(i).Value = 1
'        End If
'    Next i
'
'    If Trim(RsTemp.Fields("P_MaxM").Value) <> "" Then
'        LstProg(LstProg.Count - 1).Value = 1
'        Text5(3).Text = Trim(RsTemp.Fields("P_MaxM").Value)
'    End If
'    If Trim(RsTemp.Fields("P_MinM").Value) <> "" Then
'        LstProg(0).Value = 1
'        Text5(2).Text = Trim(RsTemp.Fields("P_MinM").Value)
'    End If
'  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'  ' 25-34
'    For i = 1 To ChkRecaida.Count - 2
'        If RsTemp.Fields(i + 24).Value Then
'            ChkRecaida(i).Value = 1
'        End If
'    Next i
'
'    If Trim(RsTemp.Fields("R_MaxM").Value) <> "" Then
'        ChkRecaida(ChkRecaida.Count - 1).Value = 1
'        Text5(5).Text = Trim(RsTemp.Fields("R_MaxM").Value)
'    End If
'    If Trim(RsTemp.Fields("R_MinM").Value) <> "" Then
'        ChkRecaida(0).Value = 1
'        Text5(4).Text = Trim(RsTemp.Fields("R_MinM").Value)
'    End If
  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  ' 37-46
'    For i = 1 To ChkMuert.Count - 2
'        If RsTemp.Fields(i + 36).Value Then
'            ChkMuert(i).Value = 1
'        End If
'    Next i
'
'    If Trim(RsTemp.Fields("M_MaxM").Value) <> "" Then
'        ChkMuert(ChkMuert.Count - 1).Value = 1
'        Text5(7).Text = Trim(RsTemp.Fields("M_MaxM").Value)
'    End If
'    If Trim(RsTemp.Fields("M_MinM").Value) <> "" Then
'        ChkMuert(0).Value = 1
'        Text5(6).Text = Trim(RsTemp.Fields("M_MinM").Value)
'    End If
'  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'Else
'    For i = 0 To LstEnfer.Count - 1
'        LstEnfer(i).Value = 0
'    Next i
'    For i = 0 To LstProg.Count - 1
'        LstProg(i).Value = 0
'    Next i
'    For i = 0 To ChkRecaida.Count - 1
'        ChkRecaida(i).Value = 0
'    Next i
'    For i = 0 To ChkMuert.Count - 1
'        ChkMuert(i).Value = 0
'    Next i
'    For i = 0 To Text5.Count - 1
'        Text5(i).Text = ""
'    Next i
'End If

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
'    Combo2.ListIndex = -1
'    Combo3.ListIndex = -1
'    Combo4.ListIndex = -1
'    Combo5.ListIndex = -1
'    Combo6.ListIndex = -1
'    Combo7.ListIndex = -1
'    TxtGleason.Text = ""
'    ChkGleason.Value = 0
'    OptBordes(0).Value = True
'    OptBordes(0).Value = False
    
'    For ii = 0 To ChkGlobal.Count - 1
'        ChkGlobal(ii).Value = 0
'        If ii <= 56 Then CboGeneral(ii).ListIndex = -1
'    Next ii
'    ChkGlobal(57).Value = 0
'    TxtOtros.Text = ""
'
'
 ' Informe Final
 
'    TxtExamFIni.Text = ""
'    TxtExamFin.Text = ""
'    TxtDiagFin.Text = ""
'    TxtAnatFin.Text = ""
'    TxtCompliFin.Text = ""
'    TxtInidiceFin.Text = ""
'    TxtTratamientoFin.Text = ""
'    TxtDosisT.Text = ""
'    TxtDosisD.Text = ""
'    TxtSesionesFin.Text = ""
'    Combo9.ListIndex = -1
'    Combo8.ListIndex = -1
'
End Sub

Sub Cargar_Datos_Braqui()

On Error GoTo WrtError

If RsBraqui.RecordCount <> 0 Then
    IdReg = RsBraqui.Fields("IdBraquiterapia").Value
    
'   if Not IsNull(RsBraqui.Fields("Modalidad").Value) Then CboModalidad.ListIndex = RsBraqui.Fields("Modalidad").Value Else  CboModalidad.ListIndex = -1
    If Not IsNull(RsBraqui.Fields("Modalidad").Value) Then
        For ii = 0 To CboModalidad.ListCount
            If CboModalidad.ItemData(ii) = Val(RsBraqui.Fields("Modalidad").Value) Then
                CboModalidad.ListIndex = ii
                Exit For
            End If
        Next ii
    End If

       
    'If Not IsNull(RsBraqui.Fields("SubModalidad").Value) Then CboSubModalidad.ListIndex = Val(RsBraqui.Fields("SubModalidad").Value) Else CboSubModalidad.ListIndex = -1
    If Not IsNull(RsBraqui.Fields("SubModalidad").Value) Then
        For ii = 0 To CboSubModalidad.ListCount - 1
            If CboSubModalidad.ItemData(ii) = Val(RsBraqui.Fields("SubModalidad").Value) Then
                CboSubModalidad.ListIndex = ii
                Exit For
            End If
        Next ii
    End If
     
    If Not IsNull(RsBraqui.Fields("DosisFracc").Value) Then TxtDosisFraccion.Text = RsBraqui.Fields("DosisFracc").Value Else TxtDosisFraccion.Text = ""
    If Not IsNull(RsBraqui.Fields("Sesiones").Value) Then TxtSesiones.Text = RsBraqui.Fields("Sesiones").Value Else TxtSesiones.Text = ""
    If Not IsNull(RsBraqui.Fields("Aplicador").Value) Then CboAplicador.ListIndex = RsBraqui.Fields("Aplicador").Value Else CboAplicador.ListIndex = -1
    If Not IsNull(RsBraqui.Fields("Sedaccion").Value) = 1 Then
        OptSedacion(0).Value = True
        OptSedacion(1).Value = False
    Else
        OptSedacion(1).Value = True
        OptSedacion(0).Value = False
    End If

    NoReg.Caption = "Registro: " & RsBraqui.AbsolutePosition & " / " & RsBraqui.RecordCount

    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco1").Value) Then
            For ii = 0 To CboFarmaco1.ListCount - 1
                If CboFarmaco1.ItemData(ii) = Val(RsBraqui.Fields("Farmaco1").Value) Then
                    CboFarmaco1.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
        
    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco2").Value) Then
            For ii = 0 To CboFarmaco2.ListCount - 1
                If CboFarmaco2.ItemData(ii) = Val(RsBraqui.Fields("Farmaco2").Value) Then
                    CboFarmaco2.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    
    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco3").Value) Then
            For ii = 0 To CboFarmaco3.ListCount - 1
                If CboFarmaco3.ItemData(ii) = Val(RsBraqui.Fields("Farmaco3").Value) Then
                    CboFarmaco3.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    
    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco4").Value) Then
            For ii = 0 To CboFarmaco4.ListCount - 1
                If CboFarmaco4.ItemData(ii) = Val(RsBraqui.Fields("Farmaco4").Value) Then
                    CboFarmaco4.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    
    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco5").Value) Then
            For ii = 0 To CboFarmaco5.ListCount - 1
                If CboFarmaco5.ItemData(ii) = Val(RsBraqui.Fields("Farmaco5").Value) Then
                    CboFarmaco5.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If

    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco6").Value) Then
            For ii = 0 To CboFarmaco6.ListCount - 1
                If CboFarmaco6.ItemData(ii) = Val(RsBraqui.Fields("Farmaco6").Value) Then
                    CboFarmaco6.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    
    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco7").Value) Then
            For ii = 0 To CboFarmaco7.ListCount - 1
                If CboFarmaco7.ItemData(ii) = Val(RsBraqui.Fields("Farmaco7").Value) Then
                    CboFarmaco7.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    
    If RsBraqui.RecordCount <> 0 Then
        If Not IsNull(RsBraqui.Fields("Farmaco8").Value) Then
            For ii = 0 To CboFarmaco8.ListCount - 1
                If CboFarmaco8.ItemData(ii) = Val(RsBraqui.Fields("Farmaco8").Value) Then
                    CboFarmaco8.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
    End If
    
    OptAplicacion(0).Value = True
    
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    If Not IsNull(RsBraqui.Fields("Aplicacion1a").Value) Then CboAplicacion1(0).ListIndex = RsBraqui.Fields("Aplicacion1a").Value Else CboAplicacion1(0).ListIndex = -1
    
    If Not IsNull(RsBraqui.Fields("DosisAplicacion1a").Value) Then TxtDosisAplicacion1a.Text = RsBraqui.Fields("DosisAplicacion1a").Value Else TxtDosisAplicacion1a.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion1b").Value) Then CboAplicacion1(1).ListIndex = RsBraqui.Fields("Aplicacion1b").Value Else CboAplicacion1(1).ListIndex = -1

    If Not IsNull(RsBraqui.Fields("DosisAplicacion1b").Value) Then TxtDosisAplicacion1b.Text = RsBraqui.Fields("DosisAplicacion1b").Value Else TxtDosisAplicacion1b.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion1c").Value) Then CboAplicacion1(2).ListIndex = RsBraqui.Fields("Aplicacion1c").Value Else CboAplicacion1(2).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion1c").Value) Then TxtDosisAplicacion1c.Text = RsBraqui.Fields("DosisAplicacion1c").Value Else TxtDosisAplicacion1c.Text = ""
    
    If Not IsNull(RsBraqui.Fields("Aplicacion1d").Value) Then CboAplicacion1(3).ListIndex = RsBraqui.Fields("Aplicacion1d").Value Else CboAplicacion1(3).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion1d").Value) Then TxtDosisAplicacion1d.Text = RsBraqui.Fields("DosisAplicacion1d").Value Else TxtDosisAplicacion1d.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion1e").Value) Then CboAplicacion1(4).ListIndex = RsBraqui.Fields("Aplicacion1e").Value Else CboAplicacion1(4).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion1e").Value) Then TxtDosisAplicacion1e.Text = RsBraqui.Fields("DosisAplicacion1e").Value Else TxtDosisAplicacion1e.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion1f").Value) Then CboAplicacion1(5).ListIndex = RsBraqui.Fields("Aplicacion1f").Value Else CboAplicacion1(5).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion1f").Value) Then TxtDosisAplicacion1f.Text = RsBraqui.Fields("DosisAplicacion1f").Value Else TxtDosisAplicacion1f.Text = ""

    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    If Not IsNull(RsBraqui.Fields("Aplicacion2a").Value) Then CboAplicacion2(0).ListIndex = RsBraqui.Fields("Aplicacion2a").Value Else CboAplicacion2(0).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion2a").Value) Then TxtDosisAplicacion2a.Text = RsBraqui.Fields("DosisAplicacion2a").Value Else TxtDosisAplicacion2a.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion2b").Value) Then CboAplicacion2(1).ListIndex = RsBraqui.Fields("Aplicacion2b").Value Else CboAplicacion2(1).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion2b").Value) Then TxtDosisAplicacion2b.Text = RsBraqui.Fields("DosisAplicacion2b").Value Else TxtDosisAplicacion2b.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion2c").Value) Then CboAplicacion2(2).ListIndex = RsBraqui.Fields("Aplicacion2c").Value Else CboAplicacion2(2).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion2c").Value) Then TxtDosisAplicacion2c.Text = RsBraqui.Fields("DosisAplicacion2c").Value Else TxtDosisAplicacion2c.Text = ""
    
    If Not IsNull(RsBraqui.Fields("Aplicacion2d").Value) Then CboAplicacion2(3).ListIndex = RsBraqui.Fields("Aplicacion2d").Value Else CboAplicacion2(3).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion2d").Value) Then TxtDosisAplicacion2d.Text = RsBraqui.Fields("DosisAplicacion2d").Value Else TxtDosisAplicacion2d.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion2e").Value) Then CboAplicacion2(4).ListIndex = RsBraqui.Fields("Aplicacion2e").Value Else CboAplicacion2(4).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion2e").Value) Then TxtDosisAplicacion2e.Text = RsBraqui.Fields("DosisAplicacion2e").Value Else TxtDosisAplicacion2e.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion2f").Value) Then CboAplicacion2(5).ListIndex = RsBraqui.Fields("Aplicacion2f").Value Else CboAplicacion2(5).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion2f").Value) Then TxtDosisAplicacion2f.Text = RsBraqui.Fields("DosisAplicacion2f").Value Else TxtDosisAplicacion2f.Text = ""

    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    If Not IsNull(RsBraqui.Fields("Aplicacion3a").Value) Then CboAplicacion3(0).ListIndex = RsBraqui.Fields("Aplicacion3a").Value Else CboAplicacion3(0).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion3a").Value) Then TxtDosisAplicacion3a.Text = RsBraqui.Fields("DosisAplicacion3a").Value Else TxtDosisAplicacion3a.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion3b").Value) Then CboAplicacion3(1).ListIndex = RsBraqui.Fields("Aplicacion3b").Value Else CboAplicacion3(1).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion3b").Value) Then TxtDosisAplicacion3b.Text = RsBraqui.Fields("DosisAplicacion3b").Value Else TxtDosisAplicacion3b.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion3c").Value) Then CboAplicacion3(2).ListIndex = RsBraqui.Fields("Aplicacion3c").Value Else CboAplicacion3(2).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion3c").Value) Then TxtDosisAplicacion3c.Text = RsBraqui.Fields("DosisAplicacion3c").Value Else TxtDosisAplicacion3c.Text = ""
    
    If Not IsNull(RsBraqui.Fields("Aplicacion3d").Value) Then CboAplicacion3(3).ListIndex = RsBraqui.Fields("Aplicacion3d").Value Else CboAplicacion3(3).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion3d").Value) Then TxtDosisAplicacion3d.Text = RsBraqui.Fields("DosisAplicacion3d").Value Else TxtDosisAplicacion3d.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion3e").Value) Then CboAplicacion3(4).ListIndex = RsBraqui.Fields("Aplicacion3e").Value Else CboAplicacion3(4).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion3e").Value) Then TxtDosisAplicacion3e.Text = RsBraqui.Fields("DosisAplicacion3e").Value Else TxtDosisAplicacion3e.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion3f").Value) Then CboAplicacion3(5).ListIndex = RsBraqui.Fields("Aplicacion3f").Value Else CboAplicacion3(5).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion3f").Value) Then TxtDosisAplicacion3f.Text = RsBraqui.Fields("DosisAplicacion3f").Value Else TxtDosisAplicacion3f.Text = ""

    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    If Not IsNull(RsBraqui.Fields("Aplicacion4a").Value) Then CboAplicacion4(0).ListIndex = RsBraqui.Fields("Aplicacion4a").Value Else CboAplicacion4(0).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion4a").Value) Then TxtDosisAplicacion4a.Text = RsBraqui.Fields("DosisAplicacion4a").Value Else TxtDosisAplicacion4a.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion4b").Value) Then CboAplicacion4(1).ListIndex = RsBraqui.Fields("Aplicacion4b").Value Else CboAplicacion4(1).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion4b").Value) Then TxtDosisAplicacion4b.Text = RsBraqui.Fields("DosisAplicacion4b").Value Else TxtDosisAplicacion4b.Text = ""
      
    If Not IsNull(RsBraqui.Fields("Aplicacion4c").Value) Then CboAplicacion4(2).ListIndex = RsBraqui.Fields("Aplicacion4c").Value Else CboAplicacion4(2).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion4c").Value) Then TxtDosisAplicacion4c.Text = RsBraqui.Fields("DosisAplicacion4c").Value Else TxtDosisAplicacion4c.Text = ""
    
    If Not IsNull(RsBraqui.Fields("Aplicacion4d").Value) Then CboAplicacion4(3).ListIndex = RsBraqui.Fields("Aplicacion4d").Value Else CboAplicacion4(3).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion4d").Value) Then TxtDosisAplicacion4d.Text = RsBraqui.Fields("DosisAplicacion4d").Value Else TxtDosisAplicacion4d.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion4e").Value) Then CboAplicacion4(4).ListIndex = RsBraqui.Fields("Aplicacion4e").Value Else CboAplicacion4(4).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion4e").Value) Then TxtDosisAplicacion4e.Text = RsBraqui.Fields("DosisAplicacion4e").Value Else TxtDosisAplicacion4e.Text = ""

    If Not IsNull(RsBraqui.Fields("Aplicacion4f").Value) Then CboAplicacion4(5).ListIndex = RsBraqui.Fields("Aplicacion4f").Value Else CboAplicacion4(5).ListIndex = -1
    If Not IsNull(RsBraqui.Fields("DosisAplicacion4f").Value) Then TxtDosisAplicacion4f.Text = RsBraqui.Fields("DosisAplicacion4f").Value Else TxtDosisAplicacion4f.Text = ""



    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' Carga los algunos datos al informe final
    
    CARGAR_COMPLICA
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
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub




'
'
Sub carga_datos_radio()
On Error GoTo WrtError


'Limpiar_MGral

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
Dim MError
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
Dim MError
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
Dim MError
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
Frame5(0).BackColor = &HE0E0E0
FrameAplicacion(0).BackColor = &HE0E0E0
FrameAplicacion(1).BackColor = &HE0E0E0
FrameAplicacion(2).BackColor = &HE0E0E0
FrameAplicacion(3).BackColor = &HE0E0E0

'CboModificarMedicoTratante.Enabled = True
'DesactivarTextos
'OptInformeMedico(0).Value = False
'OptInformeMedico(0).Value = True
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

End Sub
Sub Deshabilita_Btns()
On Error Resume Next

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
'        Call CONSULTA_INFORME
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

'ActivarTextos
'CboModificarMedicoTratante.Enabled = False

'Call Blanqueo
'Call Habilita_Btns("sin informe")

Cambio = 0
actualiza = 0
Frame1.Enabled = True
Frame5(0).BackColor = &HEAEFEF
FrameAplicacion(0).BackColor = &HEAEFEF
FrameAplicacion(1).BackColor = &HEAEFEF
FrameAplicacion(2).BackColor = &HEAEFEF
FrameAplicacion(3).BackColor = &HEAEFEF



For i = 0 To FrameDato.Count - 1
    FrameDato(i).BackColor = &HEAEFEF
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
Call Consulta_Braqui
'
If Not (RsBraqui.EOF) Then
    Call Cargar_Datos_Braqui
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
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
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
'Frame4(0).BackColor = &HE0E0E0
'Frame5(0).BackColor = &HE0E0E0
'Frame5(1).BackColor = &HE0E0E0
'
'Frame2.BackColor = &HEAEFEF
'Frame4(0).BackColor = &HEAEFEF
'Frame5(0).BackColor = &HEAEFEF
'Frame5(1).BackColor = &HEAEFEF
'For i = 0 To FrameDato.Count - 1
'    FrameDato(i).BackColor = &HEAEFEF
'Next i
'
'
'For i = 0 To OptBordes.Count - 1
'    OptBordes(i).BackColor = &HEAEFEF
'Next i
'
'Frame10.BackColor = &HEAEFEF
'Frame8.BackColor = &HEAEFEF
'ChkGleason.BackColor = &HEAEFEF
'
'For i = 0 To ChkMuert.Count - 1
'    ChkMuert(i).BackColor = &HEAEFEF
'    ChkMuert(i).Value = 0
'Next i
'For i = 0 To ChkRecaida.Count - 1
'    ChkRecaida(i).BackColor = &HEAEFEF
'    ChkRecaida(i).Value = 0
'Next i
'For i = 0 To LstProg.Count - 1
'    LstProg(i).BackColor = &HEAEFEF
'    LstProg(i).Value = 0
'Next i
'For i = 0 To LstEnfer.Count - 1
'    LstEnfer(i).BackColor = &HEAEFEF
'    LstEnfer(i).Value = 0
'Next i
'
'For i = 0 To Text5.Count - 1
'    Text5(i).Text = ""
'Next i
'
'For i = 0 To ChkGlobal.Count - 1
'    ChkGlobal(i).BackColor = &HEAEFEF
'    ChkGlobal(i).Value = 0
'Next i


'If IdPac1 = "" Then
'    CSql = "select * from Paciente"
'    Set RsCargarPacientes = CrearRS(CSql)
'    RsCargarPacientes.MoveFirst
'End If
'
'Carga_De_Datos
'
'Call CONSULTA_INFORME
'If RsInformeMed.RecordCount <> 0 Then
'    Call carga_datos_radio
'    Cambio = 0
'    actualiza = 1
'    Else
'    Call Habilita_Btns("sin informe")
'    Cambio = 0
'    actualiza = 0
'End If

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
    
    CSql = "update INFORME_MEDICO set Estado=2 where IdInforme = " & IdReg
    Set RsDeshabilitar = CrearRS(CSql)
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Borrar Informe Médico"
    BorrarRegPendiente
    
    MsgBox "El Informe Medico del Paciente: " & Text3.Text & " " & Text4.Text & " ha sido eliminado del Registro!", vbInformation + vbOKOnly, "Operacion Exitosa!"
    BtnDesHacer_Click

Else
    MsgBox "Usted No tiene permiso para borrar este informe medico", vbCritical + vbOKOnly, "Error"

End If

Exit Sub
WrtError:
Dim MError
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
Dim MError
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
Dim MError
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
FrmEvolucion.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim RsGuardar As New ADODB.Recordset
Dim RsActualizar As New ADODB.Recordset
On Error GoTo WrtError
Cambio = 1
If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente antes de Guardar!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub
'

'
' Obtiene el Nuevo ID para el informe medico
    Dim RsMaxReg As New ADODB.Recordset
    CSql = "Select max(IdBraquiterapia)+1 as MaxReg From Braquiterapia"
    Set RsMaxReg = CrearRS(CSql)

    If Not IsNull(RsMaxReg.Fields("MaxReg").Value) Then
        MaxReg = RsMaxReg.Fields("MaxReg").Value
    Else
        MaxReg = "1"
    End If
''mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm

If actualiza = 1 Then
    CSql = "Select * From Braquiterapia Where IdBraquiterapia='" & IdReg & "'"
    Set RsActualizar = CrearRS(CSql)

  
    'RsGuardar.Fields("IdBraquiterapia").Value = MaxReg
    'RsGuardar.Fields("IdMedicot").Value = IdMedT
    'RsGuardar.Fields("IdPaciente").Value = IdPac1
    RsActualizar.Fields("IdUser").Value = IdUser
    RsActualizar.Fields("Fecha").Value = Format(DtpFecha.Value, "dd/MM/yyyy")
    RsActualizar.Fields("Modalidad").Value = CboModalidad.ItemData(CboModalidad.ListIndex)
    RsActualizar.Fields("SubModalidad").Value = CboSubModalidad.ItemData(CboSubModalidad.ListIndex)
    RsActualizar.Fields("DosisFracc").Value = Trim(TxtDosisFraccion.Text)
    RsActualizar.Fields("Sesiones").Value = Trim(TxtSesiones.Text)
    RsActualizar.Fields("Aplicador").Value = CboAplicador.ItemData(CboAplicador.ListIndex)
    RsActualizar.Fields("Sedaccion").Value = Trim(Sedacion)
    'RsActualizar.Fields("Estado").Value = Trim(Text22.Text)
    RsActualizar.Fields("Farmaco1").Value = CboFarmaco1.ItemData(CboFarmaco1.ListIndex)
    RsActualizar.Fields("Farmaco2").Value = CboFarmaco2.ItemData(CboFarmaco2.ListIndex)
    RsActualizar.Fields("Farmaco3").Value = CboFarmaco3.ItemData(CboFarmaco3.ListIndex)
    RsActualizar.Fields("Farmaco4").Value = CboFarmaco4.ItemData(CboFarmaco4.ListIndex)
    RsActualizar.Fields("Farmaco5").Value = CboFarmaco5.ItemData(CboFarmaco5.ListIndex)
    RsActualizar.Fields("Farmaco6").Value = CboFarmaco6.ItemData(CboFarmaco6.ListIndex)
    RsActualizar.Fields("Farmaco7").Value = CboFarmaco7.ItemData(CboFarmaco7.ListIndex)
    RsActualizar.Fields("Farmaco8").Value = CboFarmaco8.ItemData(CboFarmaco8.ListIndex)
        
    RsActualizar.Fields("Aplicacion1a").Value = CboAplicacion1(0).ListIndex
    RsActualizar.Fields("DosisAplicacion1a").Value = TxtDosisAplicacion1a.Text
    RsActualizar.Fields("Aplicacion1b").Value = CboAplicacion1(1).ListIndex
    RsActualizar.Fields("DosisAplicacion1b").Value = TxtDosisAplicacion1b.Text
    RsActualizar.Fields("Aplicacion1c").Value = CboAplicacion1(2).ItemData(CboAplicacion1(2).ListIndex)
    RsActualizar.Fields("DosisAplicacion1c").Value = TxtDosisAplicacion1c.Text
    RsActualizar.Fields("Aplicacion1d").Value = CboAplicacion1(3).ItemData(CboAplicacion1(3).ListIndex)
    RsActualizar.Fields("DosisAplicacion1d").Value = TxtDosisAplicacion1d.Text
    RsActualizar.Fields("Aplicacion1e").Value = CboAplicacion1(4).ItemData(CboAplicacion1(4).ListIndex)
    RsActualizar.Fields("DosisAplicacion1e").Value = TxtDosisAplicacion1e.Text
    RsActualizar.Fields("Aplicacion1f").Value = CboAplicacion1(5).ItemData(CboAplicacion1(5).ListIndex)
    RsActualizar.Fields("DosisAplicacion1f").Value = TxtDosisAplicacion1f.Text
    
    RsActualizar.Fields("Aplicacion2a").Value = CboAplicacion2(0).ItemData(CboAplicacion2(0).ListIndex)
    RsActualizar.Fields("DosisAplicacion2a").Value = TxtDosisAplicacion2a.Text
    RsActualizar.Fields("Aplicacion2b").Value = CboAplicacion2(1).ItemData(CboAplicacion2(1).ListIndex)
    RsActualizar.Fields("DosisAplicacion2b").Value = TxtDosisAplicacion2b.Text
    RsActualizar.Fields("Aplicacion2c").Value = CboAplicacion2(2).ItemData(CboAplicacion2(2).ListIndex)
    RsActualizar.Fields("DosisAplicacion2c").Value = TxtDosisAplicacion2c.Text
    RsActualizar.Fields("Aplicacion2d").Value = CboAplicacion2(3).ItemData(CboAplicacion2(3).ListIndex)
    RsActualizar.Fields("DosisAplicacion2d").Value = TxtDosisAplicacion2d.Text
    RsActualizar.Fields("Aplicacion2e").Value = CboAplicacion2(4).ItemData(CboAplicacion2(4).ListIndex)
    RsActualizar.Fields("DosisAplicacion2e").Value = TxtDosisAplicacion2e.Text
    RsActualizar.Fields("Aplicacion2f").Value = CboAplicacion2(5).ItemData(CboAplicacion2(5).ListIndex)
    RsActualizar.Fields("DosisAplicacion2f").Value = TxtDosisAplicacion2f.Text
    
    RsActualizar.Fields("Aplicacion3a").Value = CboAplicacion3(0).ItemData(CboAplicacion3(0).ListIndex)
    RsActualizar.Fields("DosisAplicacion3a").Value = TxtDosisAplicacion3a.Text
    RsActualizar.Fields("Aplicacion3b").Value = CboAplicacion3(1).ItemData(CboAplicacion3(1).ListIndex)
    RsActualizar.Fields("DosisAplicacion3b").Value = TxtDosisAplicacion3b.Text
    RsActualizar.Fields("Aplicacion3c").Value = CboAplicacion3(2).ItemData(CboAplicacion3(2).ListIndex)
    RsActualizar.Fields("DosisAplicacion3c").Value = TxtDosisAplicacion3c.Text
    RsActualizar.Fields("Aplicacion3d").Value = CboAplicacion3(3).ItemData(CboAplicacion3(3).ListIndex)
    RsActualizar.Fields("DosisAplicacion3d").Value = TxtDosisAplicacion3d.Text
    RsActualizar.Fields("Aplicacion3e").Value = CboAplicacion3(4).ItemData(CboAplicacion3(4).ListIndex)
    RsActualizar.Fields("DosisAplicacion3e").Value = TxtDosisAplicacion3e.Text
    RsActualizar.Fields("Aplicacion3f").Value = CboAplicacion3(5).ItemData(CboAplicacion3(5).ListIndex)
    RsActualizar.Fields("DosisAplicacion3f").Value = TxtDosisAplicacion3f.Text
    
    RsActualizar.Fields("Aplicacion4a").Value = CboAplicacion4(0).ItemData(CboAplicacion4(0).ListIndex)
    RsActualizar.Fields("DosisAplicacion4a").Value = TxtDosisAplicacion4a.Text
    RsActualizar.Fields("Aplicacion4b").Value = CboAplicacion4(1).ItemData(CboAplicacion4(1).ListIndex)
    RsActualizar.Fields("DosisAplicacion4b").Value = TxtDosisAplicacion4b.Text
    RsActualizar.Fields("Aplicacion4c").Value = CboAplicacion4(2).ItemData(CboAplicacion4(2).ListIndex)
    RsActualizar.Fields("DosisAplicacion4c").Value = TxtDosisAplicacion4c.Text
    RsActualizar.Fields("Aplicacion4d").Value = CboAplicacion4(3).ItemData(CboAplicacion4(3).ListIndex)
    RsActualizar.Fields("DosisAplicacion4d").Value = TxtDosisAplicacion4d.Text
    RsActualizar.Fields("Aplicacion4e").Value = CboAplicacion4(4).ItemData(CboAplicacion4(4).ListIndex)
    RsActualizar.Fields("DosisAplicacion4e").Value = TxtDosisAplicacion4e.Text
    RsActualizar.Fields("Aplicacion4f").Value = CboAplicacion4(5).ItemData(CboAplicacion4(5).ListIndex)
    RsActualizar.Fields("DosisAplicacion4f").Value = TxtDosisAplicacion4f.Text
    
    
    'IdInf = MaxReg
    RsActualizar.Update

    If IdReg = "" Then MsgBox "No hay informes seleccionados!", vbExclamation + vbOKOnly, "Error": Exit Sub


ElseIf actualiza = 0 Then
    CSql = "Select * From Braquiterapia"
    Set RsGuardar = CrearRS(CSql)

    RsGuardar.AddNew
    RsGuardar.Fields("IdBraquiterapia").Value = MaxReg
    'RsGuardar.Fields("IdMedicot").Value = IdMedT
    RsGuardar.Fields("IdPaciente").Value = IdPac1
    RsGuardar.Fields("IdUser").Value = IdUser
    RsGuardar.Fields("Fecha").Value = Format(DtpFecha.Value, "dd/MM/yyyy")
    RsGuardar.Fields("Modalidad").Value = CboModalidad.ItemData(CboModalidad.ListIndex)
    RsGuardar.Fields("SubModalidad").Value = CboSubModalidad.ItemData(CboSubModalidad.ListIndex)
    RsGuardar.Fields("DosisFracc").Value = Trim(TxtDosisFraccion.Text)
    RsGuardar.Fields("Sesiones").Value = Trim(TxtSesiones.Text)
    RsGuardar.Fields("Aplicador").Value = CboAplicador.ItemData(CboAplicador.ListIndex)
    RsGuardar.Fields("Sedaccion").Value = Trim(Sedacion)
    'RsGuardar.Fields("Estado").Value = Trim(Text22.Text)
    RsGuardar.Fields("Farmaco1").Value = CboFarmaco1.ItemData(CboFarmaco1.ListIndex)
    RsGuardar.Fields("Farmaco2").Value = CboFarmaco2.ItemData(CboFarmaco2.ListIndex)
    RsGuardar.Fields("Farmaco3").Value = CboFarmaco3.ItemData(CboFarmaco3.ListIndex)
    RsGuardar.Fields("Farmaco4").Value = CboFarmaco4.ItemData(CboFarmaco4.ListIndex)
    RsGuardar.Fields("Farmaco5").Value = CboFarmaco5.ItemData(CboFarmaco5.ListIndex)
    RsGuardar.Fields("Farmaco6").Value = CboFarmaco6.ItemData(CboFarmaco6.ListIndex)
    RsGuardar.Fields("Farmaco7").Value = CboFarmaco7.ItemData(CboFarmaco7.ListIndex)
    RsGuardar.Fields("Farmaco8").Value = CboFarmaco8.ItemData(CboFarmaco8.ListIndex)
    
    RsGuardar.Fields("Aplicacion1a").Value = CboAplicacion1(0).ItemData(CboAplicacion1(0).ListIndex)
    RsGuardar.Fields("DosisAplicacion1a").Value = TxtDosisAplicacion1a.Text
    RsGuardar.Fields("Aplicacion1b").Value = CboAplicacion1(1).ItemData(CboAplicacion1(1).ListIndex)
    RsGuardar.Fields("DosisAplicacion1b").Value = TxtDosisAplicacion1b.Text
    RsGuardar.Fields("Aplicacion1c").Value = CboAplicacion1(2).ItemData(CboAplicacion1(2).ListIndex)
    RsGuardar.Fields("DosisAplicacion1c").Value = TxtDosisAplicacion1c.Text
    RsGuardar.Fields("Aplicacion1d").Value = CboAplicacion1(3).ItemData(CboAplicacion1(3).ListIndex)
    RsGuardar.Fields("DosisAplicacion1d").Value = TxtDosisAplicacion1d.Text
    RsGuardar.Fields("Aplicacion1e").Value = CboAplicacion1(4).ItemData(CboAplicacion1(4).ListIndex)
    RsGuardar.Fields("DosisAplicacion1e").Value = TxtDosisAplicacion1e.Text
    RsGuardar.Fields("Aplicacion1f").Value = CboAplicacion1(5).ItemData(CboAplicacion1(5).ListIndex)
    RsGuardar.Fields("DosisAplicacion1f").Value = TxtDosisAplicacion1f.Text
    
    RsGuardar.Fields("Aplicacion2a").Value = CboAplicacion2(0).ItemData(CboAplicacion2(0).ListIndex)
    RsGuardar.Fields("DosisAplicacion2a").Value = TxtDosisAplicacion2a.Text
    RsGuardar.Fields("Aplicacion2b").Value = CboAplicacion2(1).ItemData(CboAplicacion2(1).ListIndex)
    RsGuardar.Fields("DosisAplicacion2b").Value = TxtDosisAplicacion2b.Text
    RsGuardar.Fields("Aplicacion2c").Value = CboAplicacion2(2).ItemData(CboAplicacion2(2).ListIndex)
    RsGuardar.Fields("DosisAplicacion2c").Value = TxtDosisAplicacion2c.Text
    RsGuardar.Fields("Aplicacion2d").Value = CboAplicacion2(3).ItemData(CboAplicacion2(3).ListIndex)
    RsGuardar.Fields("DosisAplicacion2d").Value = TxtDosisAplicacion2d.Text
    RsGuardar.Fields("Aplicacion2e").Value = CboAplicacion2(4).ItemData(CboAplicacion2(4).ListIndex)
    RsGuardar.Fields("DosisAplicacion2e").Value = TxtDosisAplicacion2e.Text
    RsGuardar.Fields("Aplicacion2f").Value = CboAplicacion2(5).ItemData(CboAplicacion2(5).ListIndex)
    RsGuardar.Fields("DosisAplicacion2f").Value = TxtDosisAplicacion2f.Text
    
    RsGuardar.Fields("Aplicacion3a").Value = CboAplicacion3(0).ItemData(CboAplicacion3(0).ListIndex)
    RsGuardar.Fields("DosisAplicacion3a").Value = TxtDosisAplicacion3a.Text
    RsGuardar.Fields("Aplicacion3b").Value = CboAplicacion3(1).ItemData(CboAplicacion3(1).ListIndex)
    RsGuardar.Fields("DosisAplicacion3b").Value = TxtDosisAplicacion3b.Text
    RsGuardar.Fields("Aplicacion3c").Value = CboAplicacion3(2).ItemData(CboAplicacion3(2).ListIndex)
    RsGuardar.Fields("DosisAplicacion3c").Value = TxtDosisAplicacion3c.Text
    RsGuardar.Fields("Aplicacion3d").Value = CboAplicacion3(3).ItemData(CboAplicacion3(3).ListIndex)
    RsGuardar.Fields("DosisAplicacion3d").Value = TxtDosisAplicacion3d.Text
    RsGuardar.Fields("Aplicacion3e").Value = CboAplicacion3(4).ItemData(CboAplicacion3(4).ListIndex)
    RsGuardar.Fields("DosisAplicacion3e").Value = TxtDosisAplicacion3e.Text
    RsGuardar.Fields("Aplicacion3f").Value = CboAplicacion3(5).ItemData(CboAplicacion3(5).ListIndex)
    RsGuardar.Fields("DosisAplicacion3f").Value = TxtDosisAplicacion3f.Text
    
    RsGuardar.Fields("Aplicacion4a").Value = CboAplicacion4(0).ItemData(CboAplicacion4(0).ListIndex)
    RsGuardar.Fields("DosisAplicacion4a").Value = TxtDosisAplicacion4a.Text
    RsGuardar.Fields("Aplicacion4b").Value = CboAplicacion4(1).ItemData(CboAplicacion4(1).ListIndex)
    RsGuardar.Fields("DosisAplicacion4b").Value = TxtDosisAplicacion4b.Text
    RsGuardar.Fields("Aplicacion4c").Value = CboAplicacion4(2).ItemData(CboAplicacion4(2).ListIndex)
    RsGuardar.Fields("DosisAplicacion4c").Value = TxtDosisAplicacion4c.Text
    RsGuardar.Fields("Aplicacion4d").Value = CboAplicacion4(3).ItemData(CboAplicacion4(3).ListIndex)
    RsGuardar.Fields("DosisAplicacion4d").Value = TxtDosisAplicacion4d.Text
    RsGuardar.Fields("Aplicacion4e").Value = CboAplicacion4(4).ItemData(CboAplicacion4(4).ListIndex)
    RsGuardar.Fields("DosisAplicacion4e").Value = TxtDosisAplicacion4e.Text
    RsGuardar.Fields("Aplicacion4f").Value = CboAplicacion4(5).ItemData(CboAplicacion4(5).ListIndex)
    RsGuardar.Fields("DosisAplicacion4f").Value = TxtDosisAplicacion4f.Text
    
    'IdInf = MaxReg
    RsGuardar.Update
'
End If
'' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


Exit Sub
'
'noguardA:
'    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
'    MsgBox Msg, vbOKOnly + vbCritical, "Error al Guardar"
'
Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub


Sub EnviarRegPendiente()
On Error GoTo WrtError
CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

a = 1
CSql = "INSERT into INFORME_MEDICO(IdInforme,idmedicot,idusuario,idpaciente,Antecedente_Flia," & _
        "Anatomia_Patol,Enfermedad_Act,Examen_Fis,Motivo_Con,Diagnotico,Tratamiento,Fecha,dosis," & _
        "dosisd,Tomografia,Cuantas,Estado,Metas,sesiones,Estadiaje) VALUES(" & MaxReg & "," & IdMedT & "," & IdUser & "," & IdPac1 & _
        ",'" & Text15.Text & "','" & Text19.Text & "','" & Text16.Text & "','" & Text20.Text & _
        "','" & Text18.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & _
        Format(DtpFecha.Value, "mm/dd/YYYY") & "','" & Val(Text8.Text) & "','" & Val(Text9.Text) & _
        "'," & tomog & ",'" & Val(Text17.Text) & "','" & a & "','" & Trim(CboMetas.Text) & "'," & Sesiones & ",'" & Estadiaje & "')"
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Oncologia"
RsRegPendiente.Fields("Tabla").Value = "Informe_Medico"
RsRegPendiente.Fields("Condicional").Value = "IdInforme=" & MaxReg
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

Exit Sub
WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

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

RsTemp.AddNew
RsTemp.Fields("Id").Value = NuevoId
RsTemp.Fields("IdPaciente").Value = IdPac1
RsTemp.Fields("IdInforme").Value = IdInf
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
Dim MError
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
'        Call CONSULTA_INFORME
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

'Leer_Tipos_Ca

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
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Private Sub BtnExamenes_Click()
On Error Resume Next

If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente!", vbExclamation + vbOKOnly, "Seleccione un Paciente": Exit Sub
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
Dim MError
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
'For ii = 0 To ChkGlobal.Count - 2
'
'    If ChkGlobal(ii).Value Then
'        CboGeneral(ii).Enabled = True
'    Else
'        CboGeneral(ii).Enabled = False
'    End If
'
'    If ChkGlobal(57).Value Then
'        TxtOtros.Enabled = True
'    Else
'        TxtOtros.Enabled = False
'    End If
'Next ii
End Sub


Private Sub CboModalidad_Click()
CboSubModalidad.Clear
If CboModalidad.ListIndex = 1 Then
    CboSubModalidad.Enabled = True
    CboSubModalidad.AddItem "Uterovaginal"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 1
    CboSubModalidad.AddItem "Vaginal"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 2
    CboSubModalidad.AddItem "Endobronquial"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 3
    CboSubModalidad.AddItem "Esofagica"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 4
End If

If CboModalidad.ListIndex = 2 Then
    CboSubModalidad.Enabled = True
    CboSubModalidad.AddItem "Prostatica"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 5
    CboSubModalidad.AddItem "Mama"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 6
    CboSubModalidad.AddItem "Partes Blandas"
    CboSubModalidad.ItemData(CboSubModalidad.NewIndex) = 7
End If

If CboModalidad.ListIndex = 3 Then
    CboSubModalidad.Enabled = False
    CboSubModalidad.Clear
End If
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
Sub Consulta_Braqui()
On Error GoTo WrtError

CSql = "Select * From Braquiterapia Where IdPaciente = " & IdPac1 & " And Estado=1 Order By Fecha Desc"
Set RsBraqui = CrearRS(CSql)
    
Exit Sub
WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub
'Sub CONSULTA_INFORME()
'On Error GoTo WrtError
'
'CSql = "Select * From Informe_Medico Where IdPaciente = " & IdPac1 & " And Estado=1 Order By Fecha Desc"
'Set RsInformeMed = CrearRS(CSql)
'
'Exit Sub
'WrtError:
'Dim MError
'MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
'Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
'Print #1, MError
'Close #1
'
'End Sub

Sub Carga_De_Datos()
On Error GoTo WrtError

    Text1.Text = RsCargarPacientes.Fields("cedulaP").Value
    DtpFechaRegistro.Value = RsCargarPacientes.Fields("Fecha_regp").Value
    Text3.Text = RsCargarPacientes.Fields("Nombrep").Value
    Text4.Text = RsCargarPacientes.Fields("Apellidop").Value
    DtpFechaNac = RsCargarPacientes.Fields("Fecha_nacimientop").Value
    Text6.Text = RsCargarPacientes.Fields("Edadp").Value
    IdPac1 = RsCargarPacientes.Fields("idpaciente").Value
    Me.Caption = "Braquiterapia - Paciente: " & IdPac1
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
Dim MError
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

CboModalidad.AddItem "Intracavitaria"
CboModalidad.ItemData(CboModalidad.NewIndex) = 1
CboModalidad.AddItem "Intersticial"
CboModalidad.ItemData(CboModalidad.NewIndex) = 2
CboModalidad.AddItem "Superficial"
CboModalidad.ItemData(CboModalidad.NewIndex) = 3

CboSubModalidad.Enabled = False

CboAplicador.AddItem "Hemscke"
CboAplicador.ItemData(CboAplicador.NewIndex) = 1
CboAplicador.AddItem "Fletcher"
CboAplicador.ItemData(CboAplicador.NewIndex) = 2
CboAplicador.AddItem "Estocolmo"
CboAplicador.ItemData(CboAplicador.NewIndex) = 3
CboAplicador.AddItem "Cateter"
CboAplicador.ItemData(CboAplicador.NewIndex) = 4


CSql = "select * from Paciente Order by IdPaciente"
Set RsCargarPacientes = CrearRS(CSql)



CboFarmaco1.AddItem "Bla"
CboFarmaco1.ItemData(CboFarmaco1.NewIndex) = 1
CboFarmaco2.AddItem "BlaBLa"
CboFarmaco2.ItemData(CboFarmaco2.NewIndex) = 2
CboFarmaco3.AddItem "BlaBLaBla"
CboFarmaco3.ItemData(CboFarmaco3.NewIndex) = 3
CboFarmaco4.AddItem "BlaBLaBlaBla"
CboFarmaco4.ItemData(CboFarmaco4.NewIndex) = 4
CboFarmaco5.AddItem "BlaBLaBlaBlaBla"
CboFarmaco5.ItemData(CboFarmaco5.NewIndex) = 5
CboFarmaco6.AddItem "BlaBLaBlaBlaBlaBla"
CboFarmaco6.ItemData(CboFarmaco6.NewIndex) = 6
CboFarmaco7.AddItem "BlaBLaBlaBlaBlaBlaBla"
CboFarmaco7.ItemData(CboFarmaco7.NewIndex) = 7
CboFarmaco8.AddItem "BlaBLaBlaBlaBlaBlaBlaBla"
CboFarmaco8.ItemData(CboFarmaco8.NewIndex) = 8


OptAplicacion(0).Value = True


Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
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
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub OptAplicacion_Click(Index As Integer)
For i = 0 To OptAplicacion.Count - 1
    If i <> Index Then
        OptAplicacion(i).Value = False
        FrameAplicacion(i).Visible = False
    Else
        OptAplicacion(i).Value = True
        FrameAplicacion(i).Visible = True
    End If
Next
End Sub

Private Sub OptBordes_Click(Index As Integer)

'OptBordesSel = OptBordes(Index).Caption

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

Private Sub OptSedacion_Click(Index As Integer)
'For i = 0 To OptSedacion.Count - 1
    'If i <> Index Then
        If OptSedacion(0).Value = True Then
            OptSedacion(1).Value = False
            Sedacion = 1
        End If
   ' Else
        If OptSedacion(1).Value = True Then
            OptSedacion(0).Value = False
            Sedacion = 0
        End If

   ' End If
'Next
End Sub

Private Sub Text31_Change()

End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtDosisFraccion_KeyPress(KeyAscii As Integer)
    If InStr("1234567890.,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        MsgBox "El caracter digitado no es válido.", vbExclamation + vbOKOnly, "Atención"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSesiones_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        MsgBox "El caracter digitado no es válido.", vbExclamation + vbOKOnly, "Atención"
        KeyAscii = 0
    End If
End Sub

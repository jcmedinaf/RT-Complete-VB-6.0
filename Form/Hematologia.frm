VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmExamenHematologico 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Examen Hematológico"
   ClientHeight    =   9435
   ClientLeft      =   3735
   ClientTop       =   1395
   ClientWidth     =   12690
   Icon            =   "Hematologia.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   12690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Height          =   9375
      Left            =   120
      TabIndex        =   42
      Top             =   0
      Width           =   12495
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   117
         Top             =   8520
         Width           =   12255
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   11160
            TabIndex        =   46
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
            MICON           =   "Hematologia.frx":1002
            PICN            =   "Hematologia.frx":101E
            PICH            =   "Hematologia.frx":11E7
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
            TabIndex        =   44
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
            MICON           =   "Hematologia.frx":141C
            PICN            =   "Hematologia.frx":1438
            PICH            =   "Hematologia.frx":16C7
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
            TabIndex        =   43
            ToolTipText     =   "Agregar "
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
            MICON           =   "Hematologia.frx":1B08
            PICN            =   "Hematologia.frx":1B24
            PICH            =   "Hematologia.frx":1CB1
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
            Left            =   9960
            TabIndex        =   45
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
            MICON           =   "Hematologia.frx":1EE6
            PICN            =   "Hematologia.frx":1F02
            PICH            =   "Hematologia.frx":21E4
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
         Height          =   2415
         Left            =   120
         TabIndex        =   98
         Top             =   120
         Width           =   12255
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Hematologia.frx":2435
            Left            =   7680
            List            =   "Hematologia.frx":2437
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   390
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   1
            Text            =   "                 "
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3240
            TabIndex        =   3
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4680
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   6120
            TabIndex        =   5
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   7680
            TabIndex        =   6
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3240
            TabIndex        =   9
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4680
            TabIndex        =   10
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1920
            TabIndex        =   2
            Top             =   1080
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1440
            TabIndex        =   122
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   48037889
            CurrentDate     =   39825
         End
         Begin MSComCtl2.DTPicker DtpFechaExamen 
            Height          =   375
            Left            =   4680
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   48037889
            CurrentDate     =   39825
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solo Informativo"
            Height          =   195
            Left            =   6360
            TabIndex        =   126
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha del Examen:"
            Height          =   195
            Left            =   3240
            TabIndex        =   124
            Top             =   450
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Actual:"
            Height          =   195
            Left            =   360
            TabIndex        =   123
            Top             =   450
            Width           =   990
         End
         Begin VB.Label NoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   6720
            TabIndex        =   121
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Glóbulos Rojos"
            Height          =   195
            Left            =   360
            TabIndex        =   116
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm3"
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   115
            Top             =   1170
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hematocrito"
            Height          =   195
            Left            =   1920
            TabIndex        =   114
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   113
            Top             =   1170
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hemoglobina"
            Height          =   195
            Left            =   3240
            TabIndex        =   112
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HCM"
            Height          =   195
            Left            =   4680
            TabIndex        =   111
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VCM"
            Height          =   195
            Left            =   6120
            TabIndex        =   110
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plaquetas"
            Height          =   195
            Left            =   7680
            TabIndex        =   109
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuentas Blancas"
            Height          =   195
            Left            =   360
            TabIndex        =   108
            Top             =   1560
            Width           =   1200
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Segmentados"
            Height          =   255
            Left            =   1920
            TabIndex        =   107
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Linfocitos"
            Height          =   255
            Left            =   3360
            TabIndex        =   106
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Eosinofilos"
            Height          =   255
            Left            =   4680
            TabIndex        =   105
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   2
            Left            =   4320
            TabIndex        =   104
            Top             =   1170
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   3
            Left            =   4320
            TabIndex        =   103
            Top             =   1890
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   4
            Left            =   5760
            TabIndex        =   102
            Top             =   1890
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Index           =   5
            Left            =   3000
            TabIndex        =   101
            Top             =   1890
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm3"
            Height          =   195
            Index           =   6
            Left            =   8760
            TabIndex        =   100
            Top             =   1140
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm3"
            Height          =   195
            Index           =   8
            Left            =   1440
            TabIndex        =   99
            Top             =   1890
            Width           =   330
         End
         Begin VB.Image Image2 
            Height          =   2040
            Left            =   9720
            Picture         =   "Hematologia.frx":2439
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2460
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Resultados de Laboratorio"
         Height          =   5820
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   12255
         Begin VB.Frame Frame3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nutrición "
            Height          =   975
            Left            =   240
            TabIndex        =   127
            Top             =   4680
            Width           =   6975
            Begin VB.TextBox Text39 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   4920
               TabIndex        =   38
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text33 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   120
               TabIndex        =   34
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text38 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   1320
               TabIndex        =   35
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text40 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   3720
               TabIndex        =   37
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text41 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   2520
               TabIndex        =   36
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "mg/dl"
               Height          =   195
               Index           =   28
               Left            =   6000
               TabIndex        =   133
               Top             =   540
               Width           =   405
            End
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "Grasa"
               Height          =   255
               Left            =   4920
               TabIndex        =   132
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               Caption         =   "Energia"
               Height          =   255
               Left            =   120
               TabIndex        =   131
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label41 
               BackStyle       =   0  'Transparent
               Caption         =   "Cho"
               Height          =   255
               Left            =   3720
               TabIndex        =   130
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label42 
               BackStyle       =   0  'Transparent
               Caption         =   "Minerales"
               Height          =   255
               Left            =   2520
               TabIndex        =   129
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label44 
               BackStyle       =   0  'Transparent
               Caption         =   "Vitamina"
               Height          =   255
               Left            =   1320
               TabIndex        =   128
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00EAEFEF&
            Height          =   735
            Left            =   5760
            TabIndex        =   118
            Top             =   3840
            Width           =   1455
            Begin ChamaleonButton.ChameleonBtn BtnAnterior2 
               Height          =   375
               Left            =   120
               TabIndex        =   119
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
               MICON           =   "Hematologia.frx":3525
               PICN            =   "Hematologia.frx":3541
               PICH            =   "Hematologia.frx":37D6
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
               Left            =   720
               TabIndex        =   120
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
               MICON           =   "Hematologia.frx":3A32
               PICN            =   "Hematologia.frx":3A4E
               PICH            =   "Hematologia.frx":3CE4
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
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            TabIndex        =   13
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5760
            TabIndex        =   14
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text15 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   16
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text16 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            TabIndex        =   17
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text17 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text18 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   32
            Top             =   4080
            Width           =   975
         End
         Begin VB.TextBox Text19 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            TabIndex        =   33
            Top             =   4080
            Width           =   975
         End
         Begin VB.TextBox Text21 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox Text22 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   20
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox Text23 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            TabIndex        =   21
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox Text24 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5760
            TabIndex        =   22
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox Text25 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox Text26 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   31
            Top             =   4080
            Width           =   975
         End
         Begin VB.TextBox Text28 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox Text29 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            TabIndex        =   25
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox Text30 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5760
            TabIndex        =   26
            Text            =   "                 "
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox Text31 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5760
            TabIndex        =   18
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text34 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox Text35 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   28
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox Text36 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            TabIndex        =   29
            Text            =   "                 "
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox Text37 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5760
            TabIndex        =   30
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox Text20 
            Alignment       =   2  'Center
            Height          =   1695
            Left            =   8040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   480
            Width           =   4095
         End
         Begin VB.TextBox Text27 
            Alignment       =   2  'Center
            Height          =   1695
            Left            =   8040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   2280
            Width           =   4095
         End
         Begin VB.TextBox Text32 
            Alignment       =   2  'Center
            Height          =   1575
            Left            =   8040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   4080
            Width           =   4095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   7
            Left            =   1440
            TabIndex        =   97
            Top             =   1410
            Width           =   405
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Trigliceridos "
            Height          =   255
            Left            =   4080
            TabIndex        =   96
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Colesterol"
            Height          =   255
            Left            =   2280
            TabIndex        =   95
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Acido Úrico"
            Height          =   255
            Left            =   5760
            TabIndex        =   94
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Creatinina"
            Height          =   255
            Left            =   4080
            TabIndex        =   93
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Urea"
            Height          =   255
            Left            =   2280
            TabIndex        =   92
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   15
            Left            =   1440
            TabIndex        =   91
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Glicemia"
            Height          =   255
            Left            =   360
            TabIndex        =   90
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   16
            Left            =   1440
            TabIndex        =   89
            Top             =   2730
            Width           =   405
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "HDL"
            Height          =   255
            Left            =   5760
            TabIndex        =   88
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "TGP"
            Height          =   255
            Left            =   4080
            TabIndex        =   87
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "VLDL"
            Height          =   255
            Left            =   360
            TabIndex        =   86
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "LDL"
            Height          =   255
            Left            =   360
            TabIndex        =   85
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Cloro"
            Height          =   255
            Left            =   5760
            TabIndex        =   84
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Potasio"
            Height          =   255
            Left            =   4080
            TabIndex        =   83
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Calcio"
            Height          =   255
            Left            =   2280
            TabIndex        =   82
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   23
            Left            =   1440
            TabIndex        =   81
            Top             =   2010
            Width           =   405
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Fósforo"
            Height          =   255
            Left            =   360
            TabIndex        =   80
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "TGO"
            Height          =   255
            Left            =   2280
            TabIndex        =   79
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilirrubina Dir"
            Height          =   255
            Left            =   4080
            TabIndex        =   78
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Amilasa"
            Height          =   255
            Left            =   360
            TabIndex        =   77
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   7920
            TabIndex        =   76
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   27
            Left            =   1440
            TabIndex        =   75
            Top             =   3450
            Width           =   405
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Fosfatasa Alcalina"
            Height          =   375
            Left            =   5760
            TabIndex        =   74
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilirrubina Total"
            Height          =   255
            Left            =   5760
            TabIndex        =   73
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilirrubina Ind"
            Height          =   255
            Left            =   2280
            TabIndex        =   72
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Monocitos"
            Height          =   255
            Left            =   360
            TabIndex        =   71
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Magnesio"
            Height          =   255
            Left            =   2280
            TabIndex        =   70
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Sodio"
            Height          =   255
            Left            =   4080
            TabIndex        =   69
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   9
            Left            =   3360
            TabIndex        =   68
            Top             =   3450
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   10
            Left            =   3360
            TabIndex        =   67
            Top             =   2010
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   11
            Left            =   3360
            TabIndex        =   66
            Top             =   2730
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   12
            Left            =   3360
            TabIndex        =   65
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   13
            Left            =   3360
            TabIndex        =   64
            Top             =   1500
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   14
            Left            =   5160
            TabIndex        =   63
            Top             =   3450
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   17
            Left            =   5160
            TabIndex        =   62
            Top             =   2010
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   18
            Left            =   5160
            TabIndex        =   61
            Top             =   2730
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   19
            Left            =   5160
            TabIndex        =   60
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   20
            Left            =   5160
            TabIndex        =   59
            Top             =   1410
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   21
            Left            =   6840
            TabIndex        =   58
            Top             =   3450
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   22
            Left            =   6840
            TabIndex        =   57
            Top             =   2010
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   24
            Left            =   6840
            TabIndex        =   56
            Top             =   2730
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   25
            Left            =   6840
            TabIndex        =   55
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   26
            Left            =   6840
            TabIndex        =   54
            Top             =   1410
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   29
            Left            =   1440
            TabIndex        =   53
            Top             =   4170
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   30
            Left            =   3360
            TabIndex        =   52
            Top             =   4170
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mg/dl"
            Height          =   195
            Index           =   31
            Left            =   5160
            TabIndex        =   51
            Top             =   4170
            Width           =   405
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Orina"
            Height          =   255
            Left            =   7440
            TabIndex        =   50
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Heces"
            Height          =   255
            Left            =   7440
            TabIndex        =   49
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Otros"
            Height          =   255
            Left            =   7440
            TabIndex        =   48
            Top             =   4080
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "FrmExamenHematologico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsExamen As New ADODB.Recordset
Dim BD As New ADODB.Recordset
Dim Chang
Dim NReg
Dim IdExam
Dim Reg_Actual(0 To 40) As String
Dim RsTemp As New ADODB.Recordset
Dim Band As Boolean
Public IdPacH As String
Public IdLIdPacH As String

Sub CargaDatos()
On Error Resume Next
IdLIdInf = IdLDefault
IdExam = ""
If RsExamen.RecordCount <> 0 Then
    IdExam = RsExamen.Fields("IdExamen").Value
    IdLIdInf = RsExamen.Fields("IdL").Value
    
    If RsExamen.Fields("GLOBULOS").Value <> "" Then Text1.Text = RsExamen.Fields("GLOBULOS").Value Else Text1.Text = ""
    If RsExamen.Fields("HEMATOCRITO") <> "" Then Text2.Text = RsExamen.Fields("HEMATOCRITO") Else Text2.Text = ""
    If RsExamen.Fields("HEMOGLOBINA") <> "" Then Text3.Text = RsExamen.Fields("HEMOGLOBINA") Else Text3.Text = ""
    If RsExamen.Fields("HCM") <> "" Then Text4.Text = RsExamen.Fields("HCM") Else Text4.Text = ""
    If RsExamen.Fields("VCM") <> "" Then Text5.Text = RsExamen.Fields("VCM") Else Text5.Text = ""
    If RsExamen.Fields("PLAQUETAS") <> "" Then Text6.Text = RsExamen.Fields("PLAQUETAS") Else Text6.Text = ""
    If RsExamen.Fields("CUENTAS") <> "" Then Text7.Text = RsExamen.Fields("CUENTAS") Else Text7.Text = ""
    If RsExamen.Fields("SEGMENTADOS") <> "" Then Text8.Text = RsExamen.Fields("SEGMENTADOS") Else Text8.Text = ""
    If RsExamen.Fields("LINFOCITOS") <> "" Then Text9.Text = RsExamen.Fields("LINFOCITOS") Else Text9.Text = ""
    If RsExamen.Fields("EOSINOFILOS") <> "" Then Text10.Text = RsExamen.Fields("EOSINOFILOS") Else Text10.Text = ""
    If RsExamen.Fields("GLICEMIA") <> "" Then Text11.Text = RsExamen.Fields("GLICEMIA") Else Text11.Text = ""
    If RsExamen.Fields("UREA") <> "" Then Text12.Text = RsExamen.Fields("UREA") Else Text12.Text = ""
    If RsExamen.Fields("CREATININA") <> "" Then Text13.Text = RsExamen.Fields("CREATININA") Else Text13.Text = ""
    If RsExamen.Fields("ACIDO_U") <> "" Then Text14.Text = RsExamen.Fields("ACIDO_U") Else Text14.Text = ""
    If RsExamen.Fields("MONOCITOS") <> "" Then Text15.Text = RsExamen.Fields("MONOCITOS") Else Text15.Text = ""
    If RsExamen.Fields("COLESTEROL") <> "" Then Text16.Text = RsExamen.Fields("COLESTEROL") Else Text16.Text = ""
    If RsExamen.Fields("TRIGLICERIDOS") <> "" Then Text17.Text = RsExamen.Fields("TRIGLICERIDOS") Else Text17.Text = ""
    If RsExamen.Fields("MAGNESIO") <> "" Then Text18.Text = RsExamen.Fields("MAGNESIO") Else Text18.Text = ""
    If RsExamen.Fields("SODIO") <> "" Then Text19.Text = RsExamen.Fields("SODIO") Else Text19.Text = ""
    If RsExamen.Fields("ORINA") <> "" Then Text20.Text = RsExamen.Fields("ORINA") Else Text20.Text = ""
    If RsExamen.Fields("FOSFORO") <> "" Then Text21.Text = RsExamen.Fields("FOSFORO") Else Text21.Text = ""
    If RsExamen.Fields("CALCIO") <> "" Then Text22.Text = RsExamen.Fields("CALCIO") Else Text22.Text = ""
    If RsExamen.Fields("POTASIO") <> "" Then Text23.Text = RsExamen.Fields("POTASIO") Else Text23.Text = ""
    If RsExamen.Fields("CLORO") <> "" Then Text24.Text = RsExamen.Fields("CLORO") Else Text24.Text = ""
    If RsExamen.Fields("LDL") <> "" Then Text25.Text = RsExamen.Fields("LDL") Else Text25.Text = ""
    If RsExamen.Fields("VLDL") <> "" Then Text26.Text = RsExamen.Fields("VLDL") Else Text26.Text = ""
    If RsExamen.Fields("HECES") <> "" Then Text27.Text = RsExamen.Fields("HECES") Else Text27.Text = ""
    If RsExamen.Fields("TGO") <> "" Then Text28.Text = RsExamen.Fields("TGO") Else Text28.Text = ""
    If RsExamen.Fields("TGP") <> "" Then Text29.Text = RsExamen.Fields("TGP") Else Text29.Text = ""
    If RsExamen.Fields("HDL") <> "" Then Text30.Text = RsExamen.Fields("HDL") Else Text30.Text = ""
    If RsExamen.Fields("FOSFATASA") <> "" Then Text31.Text = RsExamen.Fields("FOSFATASA") Else Text31.Text = ""
    If RsExamen.Fields("OTROS") <> "" Then Text32.Text = RsExamen.Fields("OTROS") Else Text32.Text = ""
    If RsExamen.Fields("ENERGIA") <> "" Then Text33.Text = RsExamen.Fields("ENERGIA") Else Text33.Text = ""
    If RsExamen.Fields("AMILASA") <> "" Then Text34.Text = RsExamen.Fields("AMILASA") Else Text34.Text = ""
    If RsExamen.Fields("BilirrubinaI") <> "" Then Text35.Text = RsExamen.Fields("BilirrubinaI") Else Text35.Text = ""
    If RsExamen.Fields("BilirrubinaD") <> "" Then Text36.Text = RsExamen.Fields("BilirrubinaD") Else Text36.Text = ""
    If RsExamen.Fields("BilirrubinaT") <> "" Then Text37.Text = RsExamen.Fields("BilirrubinaT") Else Text37.Text = ""
    If RsExamen.Fields("VITAMINA") <> "" Then Text38.Text = RsExamen.Fields("VITAMINA") Else Text38.Text = ""
    If RsExamen.Fields("GRASA") <> "" Then Text39.Text = RsExamen.Fields("GRASA") Else Text39.Text = ""
    If RsExamen.Fields("CHO") <> "" Then Text40.Text = RsExamen.Fields("CHO") Else Text40.Text = ""
    If RsExamen.Fields("MINERAL") <> "" Then Text41.Text = RsExamen.Fields("MINERAL") Else Text41.Text = ""
    NoReg = "Registro " & RsExamen.AbsolutePosition & " / " & RsExamen.RecordCount
    
    Reg_Actual(0) = RsExamen.Fields("GLOBULOS").Value
    Reg_Actual(1) = RsExamen.Fields("HEMATOCRITO").Value
    Reg_Actual(2) = RsExamen.Fields("HEMOGLOBINA").Value
    Reg_Actual(3) = RsExamen.Fields("HCM").Value
    Reg_Actual(4) = RsExamen.Fields("VCM").Value
    Reg_Actual(5) = RsExamen.Fields("PLAQUETAS").Value
    Reg_Actual(6) = RsExamen.Fields("CUENTAS").Value
    Reg_Actual(7) = RsExamen.Fields("SEGMENTADOS").Value
    Reg_Actual(8) = RsExamen.Fields("LINFOCITOS").Value
    Reg_Actual(9) = RsExamen.Fields("EOSINOFILOS").Value
    Reg_Actual(10) = RsExamen.Fields("GLICEMIA").Value
    Reg_Actual(11) = RsExamen.Fields("UREA").Value
    Reg_Actual(12) = RsExamen.Fields("CREATININA").Value
    Reg_Actual(13) = RsExamen.Fields("ACIDO_U").Value
    Reg_Actual(14) = RsExamen.Fields("MONOCITOS").Value
    Reg_Actual(15) = RsExamen.Fields("COLESTEROL").Value
    Reg_Actual(16) = RsExamen.Fields("TRIGLICERIDOS").Value
    Reg_Actual(17) = RsExamen.Fields("MAGNESIO").Value
    Reg_Actual(18) = RsExamen.Fields("SODIO").Value
    Reg_Actual(19) = RsExamen.Fields("ORINA").Value
    Reg_Actual(20) = RsExamen.Fields("FOSFORO").Value
    Reg_Actual(21) = RsExamen.Fields("CALCIO").Value
    Reg_Actual(22) = RsExamen.Fields("POTASIO").Value
    Reg_Actual(23) = RsExamen.Fields("CLORO").Value
    Reg_Actual(24) = RsExamen.Fields("LDL").Value
    Reg_Actual(25) = RsExamen.Fields("VLDL").Value
    Reg_Actual(26) = RsExamen.Fields("HECES").Value
    Reg_Actual(27) = RsExamen.Fields("TGO").Value
    Reg_Actual(28) = RsExamen.Fields("TGP").Value
    Reg_Actual(29) = RsExamen.Fields("HDL").Value
    Reg_Actual(30) = RsExamen.Fields("FOSFATASA").Value
    Reg_Actual(31) = RsExamen.Fields("OTROS").Value
    Reg_Actual(32) = RsExamen.Fields("ENERGIA").Value
    Reg_Actual(33) = RsExamen.Fields("AMILASA").Value
    Reg_Actual(34) = RsExamen.Fields("BilirrubinaI").Value
    Reg_Actual(35) = RsExamen.Fields("BilirrubinaD").Value
    Reg_Actual(36) = RsExamen.Fields("BilirrubinaT").Value
    Reg_Actual(37) = RsExamen.Fields("VITAMINA").Value
    Reg_Actual(38) = RsExamen.Fields("GRASA").Value
    Reg_Actual(39) = RsExamen.Fields("CHO").Value
    Reg_Actual(40) = RsExamen.Fields("MINERAL").Value
    NReg = 0
    Chang = 0
    
    'RsExamen.Close
Else
    Blanqueo
    IdExam = ""
    NoReg = "Registro 0 / 0 (Sin Examenes)"
    For i = 0 To 40
        Reg_Actual(i) = ""
    Next i
End If
End Sub

Sub EnviarRegPendiente(ByVal IdExam2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "Select * From Examen Where IdExamen = " & IdExam2 & " And IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

StrSen = "INSERT INTO Examen (["
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
RsRegPendiente.Fields("Modulo").Value = "Examen Hematologico"
RsRegPendiente.Fields("Tabla").Value = "Examen"
RsRegPendiente.Fields("Condicional").Value = "IdExamen = " & IdExam2 & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

Private Sub BtnAgregar_Click()
'command3
BtnAgregar.Enabled = False
Blanqueo
NReg = 1
Text1.SetFocus
End Sub

Private Sub BtnAnterior2_Click()
On Error Resume Next
If RsTemp.RecordCount <> 0 Then
    If RsExamen.RecordCount <> 0 Then
        RsExamen.MovePrevious
        If RsExamen.BOF Then MsgBox "He llegado al Primer registro!", vbInformation + vbOKOnly, "Primer registro!": RsExamen.MoveFirst
        Call CargaDatos
    Else
        MsgBox "No hay datos cargados!", vbExclamation + vbOKOnly, "Vacio"
    End If
Else
    MsgBox "No hay datos cargados!", vbExclamation + vbOKOnly, "Vacio"
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
BtnAgregar.Enabled = True
Blanqueo
CSql = "SELECT * FROM EXAMEN WHERE IDPACIENTE = " & IdPacH & " AND FechaExamen = '" & Format(DtpFechaExamen.Value, "DD/MM/YYYY") & "' AND Activo=1 order by idexamen desc"
Set RsExamen = CrearRS(CSql)
Call CargaDatos
End Sub

Function Validar_Numero_Vacio(ByRef Textoo As String) As Double
    If Trim(Textoo) = "" Then
        Validar_Numero_Vacio = 0#
    Else
        Validar_Numero_Vacio = CDbl(Textoo)
    End If
End Function

Private Sub BtnGuardarActualizar_Click()
On Error GoTo WrtError
'command2
'On Error Resume Next
Dim RsTemp As New ADODB.Recordset
Dim NuevoId As String
Dim CSql As String

'verifica si hay conexion al internet
'If Not Verificar_Internet Then
'    NuevoIdL = IdL
'Else
    NuevoIdL = IdLDefault
'End If
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "SELECT MAX(IdExamen)+1 as NuevoId FROM EXAMEN"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    NuevoId = RsTemp.Fields("NuevoId")
    Else
    NuevoId = "1"
End If

BtnAgregar.Enabled = True

Select Case NReg
Case Is = 0
    If Chang = 1 Then
        
        If IdExam <= 0 Then MsgBox "Hubo un detalle en los procesos, contacte al administrador!", vbExclamation + vbOKOnly, "No se pudo actualizar!": Exit Sub
        
       
        CSql = "SELECT * FROM EXAMEN WHERE IdExamen=" & IdExam & " AND IdL = '" & IdLIdInf & "'"
        Set RsTemp = CrearRS(CSql)
        
        RsTemp.Fields("IdUsuario").Value = IdUser
        'RsTemp.Fields("IdL").Value = IdLIdInf
        'RsTemp.Fields("IDPACIENTE").Value = IdPacH
        'RsTemp.Fields("IdLIdPac").Value = IdLIdPacH
        RsTemp.Fields("FECHAEXAMEN").Value = Format((DTPicker1.Value), "dd/MM/yyyy")
        RsTemp.Fields("GLOBULOS").Value = Validar_Numero_Vacio(Text1.Text)
        RsTemp.Fields("HEMATOCRITO").Value = Validar_Numero_Vacio(Text2.Text)
        RsTemp.Fields("HEMOGLOBINA").Value = Validar_Numero_Vacio(Text3.Text)
        RsTemp.Fields("HCM").Value = Validar_Numero_Vacio(Text4.Text)
        RsTemp.Fields("VCM").Value = Validar_Numero_Vacio(Text5.Text)
        RsTemp.Fields("PLAQUETAS").Value = Validar_Numero_Vacio(Text6.Text)
        RsTemp.Fields("CUENTAS").Value = Validar_Numero_Vacio(Text7.Text)
        RsTemp.Fields("SEGMENTADOS").Value = Validar_Numero_Vacio(Text8.Text)
        RsTemp.Fields("LINFOCITOS").Value = Validar_Numero_Vacio(Text9.Text)
        RsTemp.Fields("EOSINOFILOS").Value = Validar_Numero_Vacio(Text10.Text)
        RsTemp.Fields("GLICEMIA").Value = Validar_Numero_Vacio(Text11.Text)
        RsTemp.Fields("UREA").Value = Validar_Numero_Vacio(Text12.Text)
        RsTemp.Fields("CREATININA").Value = Validar_Numero_Vacio(Text13.Text)
        RsTemp.Fields("ACIDO_U").Value = Validar_Numero_Vacio(Text14.Text)
        RsTemp.Fields("MONOCITOS").Value = Validar_Numero_Vacio(Text15.Text)
        RsTemp.Fields("COLESTEROL").Value = Validar_Numero_Vacio(Text16.Text)
        RsTemp.Fields("TRIGLICERIDOS").Value = Validar_Numero_Vacio(Text17.Text)
        RsTemp.Fields("MAGNESIO").Value = Validar_Numero_Vacio(Text18.Text)
        RsTemp.Fields("SODIO").Value = Validar_Numero_Vacio(Text19.Text)
        RsTemp.Fields("ORINA").Value = Text20.Text
        RsTemp.Fields("FOSFORO").Value = Validar_Numero_Vacio(Text21.Text)
        RsTemp.Fields("CALCIO").Value = Validar_Numero_Vacio(Text22.Text)
        RsTemp.Fields("POTASIO").Value = Validar_Numero_Vacio(Text23.Text)
        RsTemp.Fields("CLORO").Value = Validar_Numero_Vacio(Text24.Text)
        RsTemp.Fields("LDL").Value = Validar_Numero_Vacio(Text25.Text)
        RsTemp.Fields("VLDL").Value = Validar_Numero_Vacio(Text26.Text)
        RsTemp.Fields("HECES").Value = Text27.Text
        RsTemp.Fields("TGO").Value = Validar_Numero_Vacio(Text28.Text)
        RsTemp.Fields("TGP").Value = Validar_Numero_Vacio(Text29.Text)
        RsTemp.Fields("HDL").Value = Validar_Numero_Vacio(Text30.Text)
        RsTemp.Fields("FOSFATASA").Value = Validar_Numero_Vacio(Text31.Text)
        RsTemp.Fields("OTROS").Value = Text32.Text
        RsTemp.Fields("ENERGIA").Value = Validar_Numero_Vacio(Text33.Text)
        RsTemp.Fields("AMILASA").Value = Validar_Numero_Vacio(Text34.Text)
        RsTemp.Fields("BilirrubinaI").Value = Validar_Numero_Vacio(Text35.Text)
        RsTemp.Fields("BilirrubinaD").Value = Validar_Numero_Vacio(Text36.Text)
        RsTemp.Fields("BilirrubinaT").Value = Validar_Numero_Vacio(Text37.Text)
        RsTemp.Fields("VITAMINA").Value = Validar_Numero_Vacio(Text38.Text)
        RsTemp.Fields("GRASA").Value = Validar_Numero_Vacio(Text39.Text)
        RsTemp.Fields("CHO").Value = Validar_Numero_Vacio(Text40.Text)
        RsTemp.Fields("MINERAL").Value = Validar_Numero_Vacio(Text41.Text)
        RsTemp.Fields("ACTIVO").Value = "1"
        
        RsTemp.Update
        
        MsgBox "Registro Actualizado Satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
    
    
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Historial Nutricional"
        EnviarRegPendiente IdExam, IdLIdInf
    
    
    Else
        MsgBox "No hay cambios hechos en el formulario", vbExclamation + vbOKOnly, "No hay cambios"
        Exit Sub
    End If
Case Is = 1
    If Chang = 1 Then
        
        CSql = "SELECT * FROM EXAMEN"
        Set RsTemp = CrearRS(CSql)
        
        IdExam = NuevoId
        IdLIdInf = NuevoIdL
        
        RsTemp.AddNew
        RsTemp.Fields("IdExamen").Value = IdExam
        RsTemp.Fields("IdL").Value = IdLIdInf
        RsTemp.Fields("IdUsuario").Value = IdUser
        RsTemp.Fields("IDPACIENTE").Value = IdPacH
        RsTemp.Fields("IdLIdPac").Value = IdLIdPacH
        
        RsTemp.Fields("FECHAEXAMEN").Value = Format((DTPicker1.Value), "dd/MM/yyyy")
        RsTemp.Fields("GLOBULOS").Value = Validar_Numero_Vacio(Text1.Text)
        RsTemp.Fields("HEMATOCRITO").Value = Validar_Numero_Vacio(Text2.Text)
        RsTemp.Fields("HEMOGLOBINA").Value = Validar_Numero_Vacio(Text3.Text)
        RsTemp.Fields("HCM").Value = Validar_Numero_Vacio(Text4.Text)
        RsTemp.Fields("VCM").Value = Validar_Numero_Vacio(Text5.Text)
        RsTemp.Fields("PLAQUETAS").Value = Validar_Numero_Vacio(Text6.Text)
        RsTemp.Fields("CUENTAS").Value = Validar_Numero_Vacio(Text7.Text)
        RsTemp.Fields("SEGMENTADOS").Value = Validar_Numero_Vacio(Text8.Text)
        RsTemp.Fields("LINFOCITOS").Value = Validar_Numero_Vacio(Text9.Text)
        RsTemp.Fields("EOSINOFILOS").Value = Validar_Numero_Vacio(Text10.Text)
        RsTemp.Fields("GLICEMIA").Value = Validar_Numero_Vacio(Text11.Text)
        RsTemp.Fields("UREA").Value = Validar_Numero_Vacio(Text12.Text)
        RsTemp.Fields("CREATININA").Value = Validar_Numero_Vacio(Text13.Text)
        RsTemp.Fields("ACIDO_U").Value = Validar_Numero_Vacio(Text14.Text)
        RsTemp.Fields("MONOCITOS").Value = Validar_Numero_Vacio(Text15.Text)
        RsTemp.Fields("COLESTEROL").Value = Validar_Numero_Vacio(Text16.Text)
        RsTemp.Fields("TRIGLICERIDOS").Value = Validar_Numero_Vacio(Text17.Text)
        RsTemp.Fields("MAGNESIO").Value = Validar_Numero_Vacio(Text18.Text)
        RsTemp.Fields("SODIO").Value = Validar_Numero_Vacio(Text19.Text)
        RsTemp.Fields("ORINA").Value = Text20.Text
        RsTemp.Fields("FOSFORO").Value = Validar_Numero_Vacio(Text21.Text)
        RsTemp.Fields("CALCIO").Value = Validar_Numero_Vacio(Text22.Text)
        RsTemp.Fields("POTASIO").Value = Validar_Numero_Vacio(Text23.Text)
        RsTemp.Fields("CLORO").Value = Validar_Numero_Vacio(Text24.Text)
        RsTemp.Fields("LDL").Value = Validar_Numero_Vacio(Text25.Text)
        RsTemp.Fields("VLDL").Value = Validar_Numero_Vacio(Text26.Text)
        RsTemp.Fields("HECES").Value = Text27.Text
        RsTemp.Fields("TGO").Value = Validar_Numero_Vacio(Text28.Text)
        RsTemp.Fields("TGP").Value = Validar_Numero_Vacio(Text29.Text)
        RsTemp.Fields("HDL").Value = Validar_Numero_Vacio(Text30.Text)
        RsTemp.Fields("FOSFATASA").Value = Validar_Numero_Vacio(Text31.Text)
        RsTemp.Fields("OTROS").Value = Text32.Text
        RsTemp.Fields("ENERGIA").Value = Validar_Numero_Vacio(Text33.Text)
        RsTemp.Fields("AMILASA").Value = Validar_Numero_Vacio(Text34.Text)
        RsTemp.Fields("BilirrubinaI").Value = Validar_Numero_Vacio(Text35.Text)
        RsTemp.Fields("BilirrubinaD").Value = Validar_Numero_Vacio(Text36.Text)
        RsTemp.Fields("BilirrubinaT").Value = Validar_Numero_Vacio(Text37.Text)
        RsTemp.Fields("VITAMINA").Value = Validar_Numero_Vacio(Text38.Text)
        RsTemp.Fields("GRASA").Value = Validar_Numero_Vacio(Text39.Text)
        RsTemp.Fields("CHO").Value = Validar_Numero_Vacio(Text40.Text)
        RsTemp.Fields("MINERAL").Value = Validar_Numero_Vacio(Text41.Text)
        RsTemp.Fields("ACTIVO").Value = "1"
        
        RsTemp.Update

        MsgBox "Registro Agregado Satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
    
    
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Historial Nutricional"
        EnviarRegPendiente IdExam, IdLIdInf
        
    Else
        MsgBox "No hay cambios para guardar", vbExclamation + vbOKOnly, "No hay cambios"
        Exit Sub
    End If
End Select

If NReg = 0 And Chang = 1 Then
    If Reg_Actual(0) <> Text1.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo GLOBULOS de (" & Reg_Actual(0) & ") a (" & Text1.Text & ")")
    If Reg_Actual(1) <> Text2.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo HEMATOCRITO de (" & Reg_Actual(1) & ") a (" & Text2.Text & ")")
    If Reg_Actual(2) <> Text3.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo HEMOGLOBINA de (" & Reg_Actual(2) & ") a (" & Text3.Text & ")")
    If Reg_Actual(3) <> Text4.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo HCM de (" & Reg_Actual(3) & ") a (" & Text4.Text & ")")
    If Reg_Actual(4) <> Text5.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo VCM de (" & Reg_Actual(4) & ") a (" & Text5.Text & ")")
    If Reg_Actual(5) <> Text6.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo PLAQUETAS de (" & Reg_Actual(5) & ") a (" & Text6.Text & ")")
    If Reg_Actual(6) <> Text7.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo CUENTAS de (" & Reg_Actual(6) & ") a (" & Text7.Text & ")")
    If Reg_Actual(7) <> Text8.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo SEGMENTADOS de (" & Reg_Actual(7) & ") a (" & Text8.Text & ")")
    If Reg_Actual(8) <> Text9.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo LINFOCITOS de (" & Reg_Actual(8) & ") a (" & Text9.Text & ")")
    If Reg_Actual(9) <> Text10.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo EOSINOFILOS de (" & Reg_Actual(9) & ") a (" & Text10.Text & ")")
    If Reg_Actual(10) <> Text11.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo GLICEMIA de (" & Reg_Actual(10) & ") a (" & Text11.Text & ")")
    If Reg_Actual(11) <> Text12.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo UREA de (" & Reg_Actual(11) & ") a (" & Text12.Text & ")")
    If Reg_Actual(12) <> Text13.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo CREATININA de (" & Reg_Actual(12) & ") a (" & Text13.Text & ")")
    If Reg_Actual(13) <> Text14.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo ACIDO_U de (" & Reg_Actual(13) & ") a (" & Text14.Text & ")")
    If Reg_Actual(14) <> Text15.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo MONOCITOS de (" & Reg_Actual(14) & ") a (" & Text15.Text & ")")
    If Reg_Actual(15) <> Text16.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo COLESTEROL de (" & Reg_Actual(15) & ") a (" & Text16.Text & ")")
    If Reg_Actual(16) <> Text17.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo TRIGLICERIDOS de (" & Reg_Actual(16) & ") a (" & Text17.Text & ")")
    If Reg_Actual(17) <> Text18.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo MAGNESIO de (" & Reg_Actual(17) & ") a (" & Text18.Text & ")")
    If Reg_Actual(18) <> Text19.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo SODIO de (" & Reg_Actual(18) & ") a (" & Text19.Text & ")")
    If Reg_Actual(19) <> Text20.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo ORINA de (" & Reg_Actual(19) & ") a (" & Text20.Text & ")")
    If Reg_Actual(20) <> Text21.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo FOSFORO de (" & Reg_Actual(20) & ") a (" & Text21.Text & ")")
    If Reg_Actual(21) <> Text22.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo CALCIO de (" & Reg_Actual(21) & ") a (" & Text22.Text & ")")
    If Reg_Actual(22) <> Text23.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo POTASIO de (" & Reg_Actual(22) & ") a (" & Text23.Text & ")")
    If Reg_Actual(23) <> Text24.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo CLORO de (" & Reg_Actual(23) & ") a (" & Text24.Text & ")")
    If Reg_Actual(24) <> Text25.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo LDL de (" & Reg_Actual(24) & ") a (" & Text25.Text & ")")
    If Reg_Actual(25) <> Text26.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo VLDL de (" & Reg_Actual(25) & ") a (" & Text26.Text & ")")
    If Reg_Actual(26) <> Text27.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo HECES de (" & Reg_Actual(26) & ") a (" & Text27.Text & ")")
    If Reg_Actual(27) <> Text28.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo TGO de (" & Reg_Actual(27) & ") a (" & Text28.Text & ")")
    If Reg_Actual(28) <> Text29.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo TGP de (" & Reg_Actual(28) & ") a (" & Text29.Text & ")")
    If Reg_Actual(29) <> Text30.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo HDL de (" & Reg_Actual(29) & ") a (" & Text30.Text & ")")
    If Reg_Actual(30) <> Text31.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo FOSFATASA de (" & Reg_Actual(30) & ") a (" & Text31.Text & ")")
    If Reg_Actual(31) <> Text32.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo OTROS de (" & Reg_Actual(31) & ") a (" & Text32.Text & ")")
    If Reg_Actual(32) <> Text33.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo ENERGIA de (" & Reg_Actual(32) & ") a (" & Text33.Text & ")")
    If Reg_Actual(33) <> Text34.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo AMILASA de (" & Reg_Actual(33) & ") a (" & Text34.Text & ")")
    If Reg_Actual(34) <> Text35.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo BilirrubinaI de (" & Reg_Actual(34) & ") a (" & Text35.Text & ")")
    If Reg_Actual(35) <> Text36.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo BilirrubinaD de (" & Reg_Actual(35) & ") a (" & Text36.Text & ")")
    If Reg_Actual(36) <> Text37.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo BilirrubinaT de (" & Reg_Actual(36) & ") a (" & Text37.Text & ")")
    If Reg_Actual(37) <> Text38.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo VITAMINA de (" & Reg_Actual(37) & ") a (" & Text38.Text & ")")
    If Reg_Actual(38) <> Text39.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo GRASA de (" & Reg_Actual(38) & ") a (" & Text39.Text & ")")
    If Reg_Actual(39) <> Text40.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo CHO de (" & Reg_Actual(39) & ") a (" & Text40.Text & ")")
    If Reg_Actual(40) <> Text41.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Modificar", "se modifico el campo MINERAL de (" & Reg_Actual(40) & ") a (" & Text41.Text & ")")
ElseIf NReg = 1 And Chang = 1 Then
    Call Enviar_Bitacora(IdUser, "Nutricion-EXAMEN", "Ingresar", "se Ingreso un nuevo registro con IdExamen=" & NuevoId)
End If
BtnDesHacer_Click

Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
'MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub BtnSiguiente2_Click()
On Error Resume Next
If RsTemp.RecordCount <> 0 Then
    If RsExamen.RecordCount <> 0 Then
        RsExamen.MoveNext
        If RsExamen.EOF Then MsgBox "He llegado al Ultimo registro!", vbInformation + vbOKOnly, "Ultimo registro!": RsExamen.MoveLast
        Call CargaDatos
    Else
        MsgBox "No hay datos cargados!", vbExclamation + vbOKOnly, "Vacio"
    End If
Else
    MsgBox "No hay datos cargados!", vbExclamation + vbOKOnly, "Vacio"
End If

End Sub

Sub Blanqueo()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text22.Text = ""
    Text23.Text = ""
    Text24.Text = ""
    Text25.Text = ""
    Text26.Text = ""
    Text27.Text = ""
    Text28.Text = ""
    Text29.Text = ""
    Text30.Text = ""
    Text31.Text = ""
    Text32.Text = ""
    Text33.Text = ""
    Text34.Text = ""
    Text35.Text = ""
    Text36.Text = ""
    Text37.Text = ""
    Text38.Text = ""
    Text39.Text = ""
    Text40.Text = ""
    Text41.Text = ""
    NReg = 1
    Chang = 0
End Sub

Private Sub DtpFechaExamen_Change()
DtpFechaExamen_Click
End Sub

Private Sub DtpFechaExamen_Click()
Blanqueo
CSql = "SELECT * FROM EXAMEN WHERE IDPACIENTE = " & IdPacH & " AND FechaExamen = '" & Format(DtpFechaExamen.Value, "DD/MM/YYYY") & "' AND Activo=1 order by idexamen desc"
Set RsExamen = CrearRS(CSql)
Call CargaDatos
End Sub

Private Sub Form_Load()

On Error Resume Next

Label46.Visible = False
Combo1.Visible = False
Me.Caption = "Examen Hematológico - Paciente: " & IdPacH
Centrar Me
DTPicker1.Value = Now()

CSql = "SELECT FechaExamen FROM EXAMEN WHERE IDPACIENTE = " & IdPacH & " AND Activo=1"
Set RsTemp = CrearRS(CSql)
    
If RsTemp.RecordCount <> 0 Then
    RsTemp.MoveLast
    DtpFechaExamen.Value = Format(RsTemp.Fields("FechaExamen").Value, "DD/MM/YYYY")
    
    CSql = "SELECT * FROM EXAMEN WHERE IDPACIENTE = " & IdPacH & " AND FechaExamen = '" & Format(DtpFechaExamen.Value, "DD/MM/YYYY") & "' AND Activo=1 order by idexamen desc"
    Set RsExamen = CrearRS(CSql)
    
    Call CargaDatos
Else
    Call Blanqueo
End If

If RsTemp.RecordCount <> 0 Then
    Label46.Visible = True
    Combo1.Visible = True
    RsTemp.MoveFirst
    
    While Not RsTemp.EOF
        Band = False
        For i = 0 To Combo1.ListCount
            If Combo1.List(i) = Format(RsTemp.Fields("FechaExamen").Value, "DD/MM/YYYY") Then
                Band = True
                Exit For
            End If
        Next i

        If Not Band Then Combo1.AddItem Format(RsTemp.Fields("FechaExamen").Value, "DD/MM/YYYY")
        RsTemp.MoveNext
    Wend
Else
    Label46.Visible = False
    Combo1.Visible = False
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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



Private Sub Text2_KeyPress(KeyAscii As Integer)
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

Private Sub Text20_Change()
xxxx = Text20.SelStart
Chang = 1
Dim StrText, Chaa, pru As String
 Dim i  As Variant
 StrText = ""
 Chaa = ""
  Chaa = UCase(Mid(Text20.Text, 1, 1))
  StrText = Chaa
  For i = 2 To Len(Text20.Text)
    pru = LCase(Mid(Text20.Text, i, 1))
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

 Text20.Text = StrText
 Text20.SelStart = xxxx
End Sub

Private Sub Text27_Change()
Chang = 1
Dim StrText, Chaa, pru As String
 Dim i  As Variant
 StrText = ""
 Chaa = ""
  Chaa = UCase(Mid(Text27.Text, 1, 1))
  StrText = Chaa
  For i = 2 To Len(Text27.Text)
    pru = LCase(Mid(Text27.Text, i, 1))
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

 Text27.Text = StrText
 Text27.SelStart = Len(Text27.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
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

Private Sub Text32_Change()
Chang = 1
Dim StrText, Chaa, pru As String
 Dim i  As Variant
 StrText = ""
 Chaa = ""
  Chaa = UCase(Mid(Text32.Text, 1, 1))
  StrText = Chaa
  For i = 2 To Len(Text32.Text)
    pru = LCase(Mid(Text32.Text, i, 1))
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

 Text32.Text = StrText
 Text32.SelStart = Len(Text32.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
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
Private Sub Text5_KeyPress(KeyAscii As Integer)
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
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
Private Sub Text7_KeyPress(KeyAscii As Integer)
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
Private Sub Text9_KeyPress(KeyAscii As Integer)
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

Private Sub Text10_KeyPress(KeyAscii As Integer)
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

Private Sub Text12_KeyPress(KeyAscii As Integer)
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

Private Sub Text13_KeyPress(KeyAscii As Integer)
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

Private Sub Text15_KeyPress(KeyAscii As Integer)
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

Private Sub Text16_KeyPress(KeyAscii As Integer)
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

Private Sub Text17_KeyPress(KeyAscii As Integer)
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
Private Sub Text18_KeyPress(KeyAscii As Integer)
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

Private Sub Text19_KeyPress(KeyAscii As Integer)
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



Private Sub Text21_KeyPress(KeyAscii As Integer)
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

Private Sub Text22_KeyPress(KeyAscii As Integer)
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

Private Sub Text23_KeyPress(KeyAscii As Integer)
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

Private Sub Text24_KeyPress(KeyAscii As Integer)
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

Private Sub Text25_KeyPress(KeyAscii As Integer)
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

Private Sub Text26_KeyPress(KeyAscii As Integer)
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

Private Sub Text28_KeyPress(KeyAscii As Integer)
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

Private Sub Text29_KeyPress(KeyAscii As Integer)
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

Private Sub Text30_KeyPress(KeyAscii As Integer)
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

Private Sub Text31_KeyPress(KeyAscii As Integer)
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

Private Sub Text33_KeyPress(KeyAscii As Integer)
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

Private Sub Text36_KeyPress(KeyAscii As Integer)
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
Private Sub Text37_KeyPress(KeyAscii As Integer)
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

Private Sub Text38_KeyPress(KeyAscii As Integer)
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

Private Sub Text39_KeyPress(KeyAscii As Integer)
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

Private Sub Text40_KeyPress(KeyAscii As Integer)
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

Private Sub Text41_KeyPress(KeyAscii As Integer)
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

Private Sub Text1_Change()
Chang = 1
End Sub
Private Sub Text2_Change()
Chang = 1
End Sub
Private Sub Text3_Change()
Chang = 1
End Sub
Private Sub Text4_Change()
Chang = 1
End Sub
Private Sub Text5_Change()
Chang = 1
End Sub
Private Sub Text6_Change()
Chang = 1
End Sub
Private Sub Text7_Change()
Chang = 1
End Sub
Private Sub Text8_Change()
Chang = 1
End Sub
Private Sub Text9_Change()
Chang = 1
End Sub
Private Sub Text10_Change()
Chang = 1
End Sub
Private Sub Text11_Change()
Chang = 1
End Sub
Private Sub Text12_Change()
Chang = 1
End Sub
Private Sub Text13_Change()
Chang = 1
End Sub
Private Sub Text14_Change()
Chang = 1
End Sub
Private Sub Text15_Change()
Chang = 1
End Sub
Private Sub Text16_Change()
Chang = 1
End Sub
Private Sub Text17_Change()
Chang = 1
End Sub
Private Sub Text18_Change()
Chang = 1
End Sub
Private Sub Text19_Change()
Chang = 1
End Sub

Private Sub Text21_Change()
Chang = 1
End Sub
Private Sub Text22_Change()
Chang = 1
End Sub
Private Sub Text23_Change()
Chang = 1
End Sub
Private Sub Text24_Change()
Chang = 1
End Sub
Private Sub Text25_Change()
Chang = 1
End Sub
Private Sub Text26_Change()
Chang = 1
End Sub

Private Sub Text28_Change()
Chang = 1
End Sub
Private Sub Text29_Change()
Chang = 1
End Sub
Private Sub Text30_Change()
Chang = 1
End Sub
Private Sub Text31_Change()
Chang = 1
End Sub

Private Sub Text33_Change()
Chang = 1
End Sub
Private Sub Text34_Change()
Chang = 1
End Sub
Private Sub Text35_Change()
Chang = 1
End Sub
Private Sub Text36_Change()
Chang = 1
End Sub
Private Sub Text37_Change()
Chang = 1
End Sub
Private Sub Text38_Change()
Chang = 1
End Sub
Private Sub Text39_Change()
Chang = 1
End Sub
Private Sub Text40_Change()
Chang = 1
End Sub
Private Sub Text41_Change()
Chang = 1
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            'Combo1.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            Text1.SetFocus
    End Select
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyRight
            Text2.SetFocus
        Case vbKeyDown
            Text7.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyLeft
            Text9.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            Text14.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text12.SetFocus
        Case vbKeyUp
            Text7.SetFocus
        Case vbKeyRight
            Text12.SetFocus
        Case vbKeyDown
            Text17.SetFocus
    End Select
End If
End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case vbKeyUp
            Text8.SetFocus
        Case vbKeyLeft
            Text11.SetFocus
        Case vbKeyRight
            Text13.SetFocus
        Case vbKeyDown
            Text15.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text14.SetFocus
        Case vbKeyUp
            Text9.SetFocus
        Case vbKeyLeft
            Text12.SetFocus
        Case vbKeyRight
            Text14.SetFocus
        Case vbKeyDown
            Text16.SetFocus
    End Select
End If
End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text17.SetFocus
        Case vbKeyUp
            Text10.SetFocus
        Case vbKeyLeft
            Text13.SetFocus
        Case vbKeyRight
            Text20.SetFocus
        Case vbKeyDown
            Text31.SetFocus
    End Select
End If
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text16.SetFocus
        Case vbKeyUp
            Text12.SetFocus
        Case vbKeyLeft
            Text17.SetFocus
        Case vbKeyRight
            Text16.SetFocus
        Case vbKeyDown
            Text22.SetFocus
    End Select
End If
End Sub

Private Sub Text16_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text31.SetFocus
        Case vbKeyUp
            Text13.SetFocus
        Case vbKeyLeft
            Text15.SetFocus
        Case vbKeyRight
            Text31.SetFocus
        Case vbKeyDown
            Text23.SetFocus
    End Select
End If
End Sub

Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text15.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyRight
            Text15.SetFocus
        Case vbKeyDown
            Text21.SetFocus
    End Select
End If
End Sub

Private Sub Text18_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text19.SetFocus
        Case vbKeyUp
            Text35.SetFocus
        Case vbKeyLeft
            Text26.SetFocus
        Case vbKeyRight
            Text19.SetFocus
        Case vbKeyDown
            Text33.SetFocus
    End Select
End If
End Sub

Private Sub Text19_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text20.SetFocus
        Case vbKeyUp
            Text36.SetFocus
        Case vbKeyLeft
            Text18.SetFocus
        Case vbKeyRight
            Text32.SetFocus
        Case vbKeyDown
            Text33.SetFocus
    End Select
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyLeft
            Text1.SetFocus
        Case vbKeyRight
            Text3.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub Text20_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text27.SetFocus
        Case vbKeyUp
            BtnAyuda.SetFocus
        Case vbKeyLeft
            Text14.SetFocus
        Case vbKeyDown
            Text27.SetFocus
    End Select
End If
End Sub

Private Sub Text21_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text22.SetFocus
        Case vbKeyUp
            Text17.SetFocus
        Case vbKeyRight
            Text22.SetFocus
        Case vbKeyDown
            Text34.SetFocus
    End Select
End If
End Sub

Private Sub Text22_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text23.SetFocus
        Case vbKeyUp
            Text15.SetFocus
        Case vbKeyLeft
            Text21.SetFocus
        Case vbKeyRight
            Text23.SetFocus
        Case vbKeyDown
            Text28.SetFocus
    End Select
End If
End Sub

Private Sub Text23_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text24.SetFocus
        Case vbKeyUp
            Text16.SetFocus
        Case vbKeyLeft
            Text22.SetFocus
        Case vbKeyRight
            Text24.SetFocus
        Case vbKeyDown
            Text29.SetFocus
    End Select
End If
End Sub

Private Sub Text24_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text34.SetFocus
        Case vbKeyUp
            Text31.SetFocus
        Case vbKeyLeft
            Text23.SetFocus
        Case vbKeyRight
            Text27.SetFocus
        Case vbKeyDown
            Text30.SetFocus
    End Select
End If
End Sub

Private Sub Text25_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text35.SetFocus
        Case vbKeyUp
            Text34.SetFocus
        Case vbKeyRight
            Text35.SetFocus
        Case vbKeyDown
            Text26.SetFocus
    End Select
End If
End Sub

Private Sub Text26_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text18.SetFocus
        Case vbKeyUp
            Text25.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text33.SetFocus
    End Select
End If
End Sub

Private Sub Text27_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text32.SetFocus
        Case vbKeyUp
            Text20.SetFocus
        Case vbKeyLeft
            Text24.SetFocus
        Case vbKeyDown
            Text32.SetFocus
    End Select
End If
End Sub

Private Sub Text28_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text29.SetFocus
        Case vbKeyUp
            Text22.SetFocus
        Case vbKeyLeft
            Text34.SetFocus
        Case vbKeyRight
            Text29.SetFocus
        Case vbKeyDown
            Text35.SetFocus
    End Select
End If
End Sub

Private Sub Text29_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text30.SetFocus
        Case vbKeyUp
            Text23.SetFocus
        Case vbKeyLeft
            Text28.SetFocus
        Case vbKeyRight
            Text30.SetFocus
        Case vbKeyDown
            Text36.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyLeft
            Text2.SetFocus
        Case vbKeyRight
            Text4.SetFocus
        Case vbKeyDown
            Text9.SetFocus
    End Select
End If
End Sub

Private Sub Text30_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text25.SetFocus
        Case vbKeyUp
            Text24.SetFocus
        Case vbKeyLeft
            Text29.SetFocus
        Case vbKeyRight
            Text27.SetFocus
        Case vbKeyDown
            Text37.SetFocus
    End Select
End If
End Sub

Private Sub Text31_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text21.SetFocus
        Case vbKeyUp
            Text14.SetFocus
        Case vbKeyLeft
            Text16.SetFocus
        Case vbKeyRight
            Text21.SetFocus
        Case vbKeyDown
            Text24.SetFocus
    End Select
End If
End Sub

Private Sub Text32_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text27.SetFocus
        Case vbKeyLeft
            Text37.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text33_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text38.SetFocus
        Case vbKeyUp
            Text25.SetFocus
        Case vbKeyRight
            Text38.SetFocus
    End Select
End If
End Sub

Private Sub Text34_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text28.SetFocus
        Case vbKeyUp
            Text21.SetFocus
        Case vbKeyRight
            Text28.SetFocus
        Case vbKeyDown
            Text25.SetFocus
    End Select
End If
End Sub

Private Sub Text35_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text36.SetFocus
        Case vbKeyUp
            Text28.SetFocus
        Case vbKeyLeft
            Text25.SetFocus
        Case vbKeyRight
            Text36.SetFocus
        Case vbKeyDown
            Text18.SetFocus
    End Select
End If
End Sub

Private Sub Text36_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text37.SetFocus
        Case vbKeyUp
            Text29.SetFocus
        Case vbKeyLeft
            Text35.SetFocus
        Case vbKeyRight
            Text37.SetFocus
        Case vbKeyDown
            Text19.SetFocus
    End Select
End If
End Sub

Private Sub Text37_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text26.SetFocus
        Case vbKeyUp
            Text30.SetFocus
        Case vbKeyLeft
            Text36.SetFocus
        Case vbKeyRight
            Text32.SetFocus
        Case vbKeyDown
            Text19.SetFocus
    End Select
End If
End Sub

Private Sub Text38_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text41.SetFocus
        Case vbKeyUp
            Text26.SetFocus
        Case vbKeyLeft
            Text33.SetFocus
        Case vbKeyRight
            Text41.SetFocus
    End Select
End If
End Sub

Private Sub Text39_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text19.SetFocus
        Case vbKeyLeft
            Text40.SetFocus
        Case vbKeyRight
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text5.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyLeft
            Text3.SetFocus
        Case vbKeyRight
            Text5.SetFocus
        Case vbKeyDown
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text40_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text39.SetFocus
        Case vbKeyUp
            Text19.SetFocus
        Case vbKeyLeft
            Text41.SetFocus
        Case vbKeyRight
            Text39.SetFocus
    End Select
End If
End Sub

Private Sub Text41_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text40.SetFocus
        Case vbKeyUp
            Text18.SetFocus
        Case vbKeyLeft
            Text38.SetFocus
        Case vbKeyRight
            Text40.SetFocus
    End Select
End If
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyLeft
            Text4.SetFocus
        Case vbKeyRight
            Text6.SetFocus
        Case vbKeyDown
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text7.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyLeft
            Text5.SetFocus
        Case vbKeyDown
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text8.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Text8.SetFocus
        Case vbKeyDown
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyLeft
            Text7.SetFocus
        Case vbKeyRight
            Text9.SetFocus
        Case vbKeyDown
            Text12.SetFocus
    End Select
End If
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyLeft
            Text8.SetFocus
        Case vbKeyRight
            Text10.SetFocus
        Case vbKeyDown
            Text13.SetFocus
    End Select
End If
End Sub


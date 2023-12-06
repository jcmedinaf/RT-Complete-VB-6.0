VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmTablaEstadisticas 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Estadistica"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   10950
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   8280
      Width           =   10695
      Begin ChamaleonButton.ChameleonBtn BtnGraficar 
         Height          =   375
         Left            =   3840
         TabIndex        =   32
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Graficar"
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
         MICON           =   "FrmEstadisticas.frx":1002
         PICN            =   "FrmEstadisticas.frx":101E
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
         Left            =   9480
         TabIndex        =   34
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
         MICON           =   "FrmEstadisticas.frx":1395
         PICN            =   "FrmEstadisticas.frx":13B1
         PICH            =   "FrmEstadisticas.frx":157A
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
         Left            =   8160
         TabIndex        =   33
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Visible         =   0   'False
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
         MICON           =   "FrmEstadisticas.frx":17AF
         PICN            =   "FrmEstadisticas.frx":17CB
         PICH            =   "FrmEstadisticas.frx":1AAD
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
         Left            =   240
         TabIndex        =   35
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
         MICON           =   "FrmEstadisticas.frx":1CFE
         PICN            =   "FrmEstadisticas.frx":1D1A
         PICH            =   "FrmEstadisticas.frx":1E3F
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
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Resultados"
      Height          =   375
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grafica"
      Height          =   375
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Index           =   0
      Left            =   120
      TabIndex        =   55
      Top             =   480
      Width           =   10695
      Begin VB.TextBox TxtQuery 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   7680
         Visible         =   0   'False
         Width           =   10215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Diagnosticos"
         Height          =   4815
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   10215
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmEstadisticas.frx":20CF
            Left            =   9240
            List            =   "FrmEstadisticas.frx":20DC
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmEstadisticas.frx":20F0
            Left            =   7560
            List            =   "FrmEstadisticas.frx":21A2
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Diagnóstico Específico:"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox TxtDiagPer 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            TabIndex        =   66
            Top             =   1320
            Width           =   7695
         End
         Begin VB.ComboBox CboGrupo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmEstadisticas.frx":2344
            Left            =   9000
            List            =   "FrmEstadisticas.frx":234E
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   2280
            Width           =   1095
         End
         Begin VB.ComboBox CboTipo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "FrmEstadisticas.frx":2363
            Left            =   1560
            List            =   "FrmEstadisticas.frx":2379
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2280
            Width           =   1815
         End
         Begin VB.ComboBox CboTipo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "FrmEstadisticas.frx":23CD
            Left            =   1560
            List            =   "FrmEstadisticas.frx":23E3
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Eje de las ""X"""
            Height          =   1005
            Left            =   120
            TabIndex        =   63
            Top             =   3720
            Width           =   9975
            Begin VB.ComboBox CboTipoFecha 
               Height          =   315
               ItemData        =   "FrmEstadisticas.frx":2437
               Left            =   5400
               List            =   "FrmEstadisticas.frx":2441
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   205
               Width           =   1695
            End
            Begin VB.OptionButton OptEjeX 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Fechas"
               Height          =   255
               Index           =   4
               Left            =   4200
               TabIndex        =   29
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton OptEjeX 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Estadiaje"
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   9360
               TabIndex        =   28
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.OptionButton OptEjeX 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Diagnosticos"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   9360
               TabIndex        =   27
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton OptEjeX 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Sexo"
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   26
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton OptEjeX 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Edad"
               Height          =   255
               Index           =   0
               Left            =   1560
               TabIndex        =   25
               Top             =   240
               Width           =   735
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00EAEFEF&
               Height          =   420
               Left            =   1440
               TabIndex        =   102
               Top             =   125
               Width           =   5655
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00EAEFEF&
               Height          =   375
               Left            =   4080
               TabIndex        =   93
               Top             =   480
               Width           =   3015
               Begin VB.OptionButton OptEjeX2 
                  BackColor       =   &H00EAEFEF&
                  Caption         =   "+ Sexo"
                  Height          =   230
                  Index           =   1
                  Left            =   1920
                  TabIndex        =   94
                  Top             =   120
                  Width           =   855
               End
               Begin VB.OptionButton OptEjeX2 
                  BackColor       =   &H00EAEFEF&
                  Caption         =   "+ Rango de Edad"
                  Height          =   230
                  Index           =   0
                  Left            =   120
                  TabIndex        =   95
                  Top             =   120
                  Width           =   1695
               End
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Height          =   400
            Left            =   4320
            TabIndex        =   61
            Top             =   1710
            Width           =   2895
            Begin VB.OptionButton OptSesiones 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Sin rango"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton OptSesiones 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Con rango, entre"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   9
               Top             =   120
               Width           =   1500
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00EAEFEF&
            Height          =   425
            Left            =   4320
            TabIndex        =   60
            Top             =   2190
            Width           =   2895
            Begin VB.OptionButton OptEdad 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Sin rango"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   17
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton OptEdad 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Con rango, entre"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   18
               Top             =   120
               Width           =   1500
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAEFEF&
            Height          =   520
            Left            =   1560
            TabIndex        =   59
            Top             =   3120
            Width           =   8535
            Begin VB.OptionButton OptSexo 
               BackColor       =   &H00EAEFEF&
               Caption         =   "F + M"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   2880
               TabIndex        =   24
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton OptSexo 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Femenino"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   22
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton OptSexo 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Masculino"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   1440
               TabIndex        =   23
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtNoSesiones2 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   9000
            TabIndex        =   11
            Text            =   "0"
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox TxtNoSesiones1 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   10
            Text            =   "0"
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox CboRangoEdad 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmEstadisticas.frx":245A
            Left            =   7320
            List            =   "FrmEstadisticas.frx":24AE
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Sexo:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Edad:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Estadiaje:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Sesiones:"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Diagnósticos:"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Incluir rango de fechas:"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   300
            Width           =   2055
         End
         Begin VB.TextBox TxtEdad 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            TabIndex        =   16
            Text            =   "0"
            Top             =   2280
            Width           =   855
         End
         Begin VB.ComboBox CboEstadiaje 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmEstadisticas.frx":255C
            Left            =   1560
            List            =   "FrmEstadisticas.frx":2674
            TabIndex        =   13
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox TxtNoSesiones 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            TabIndex        =   7
            Text            =   "0"
            Top             =   1800
            Width           =   855
         End
         Begin VB.ComboBox CboDiagnosticos 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   840
            Width           =   4095
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   320
            Left            =   3360
            TabIndex        =   1
            Top             =   267
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   47775745
            CurrentDate     =   40190
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   320
            Left            =   6120
            TabIndex        =   2
            Top             =   267
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   47775745
            CurrentDate     =   40221
         End
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   120
            Top             =   3600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "InmunoHistoquimica:"
            Height          =   255
            Left            =   6000
            TabIndex        =   103
            Top             =   870
            Width           =   1455
         End
         Begin VB.Line Line1 
            X1              =   8700
            X2              =   8700
            Y1              =   2040
            Y2              =   1800
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin:"
            Height          =   195
            Left            =   5280
            TabIndex        =   58
            Top             =   330
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio:"
            Height          =   195
            Left            =   2400
            TabIndex        =   57
            Top             =   330
            Width           =   915
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnEliminar 
         Height          =   375
         Left            =   10440
         TabIndex        =   70
         ToolTipText     =   "Eliminar"
         Top             =   7200
         Visible         =   0   'False
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
         MICON           =   "FrmEstadisticas.frx":2829
         PICN            =   "FrmEstadisticas.frx":2845
         PICH            =   "FrmEstadisticas.frx":29E9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnEnlistar 
         Height          =   375
         Left            =   10440
         TabIndex        =   72
         Top             =   7200
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ejecutar Consulta"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmEstadisticas.frx":2B88
         PICN            =   "FrmEstadisticas.frx":2BA4
         PICH            =   "FrmEstadisticas.frx":2FD4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   855
         Left            =   10320
         TabIndex        =   30
         Top             =   6240
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   1508
         Object.Width           =   10185
         Object.Height          =   825
         ScrollBar       =   3
         DrawColorGrid   =   1
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardar 
         Height          =   375
         Left            =   10440
         TabIndex        =   69
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   7200
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar Consulta"
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
         MICON           =   "FrmEstadisticas.frx":3270
         PICN            =   "FrmEstadisticas.frx":328C
         PICH            =   "FrmEstadisticas.frx":351B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT Complete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   9240
         TabIndex        =   105
         Top             =   7320
         Width           =   1110
      End
      Begin VB.Label NroReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 / 0"
         Height          =   195
         Left            =   10200
         TabIndex        =   71
         Top             =   5880
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   $"FrmEstadisticas.frx":395C
         Height          =   195
         Index           =   1
         Left            =   10320
         TabIndex        =   68
         Top             =   6000
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta:"
         Height          =   195
         Left            =   240
         TabIndex        =   62
         Top             =   7440
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   10695
      Begin ChamaleonButton.ChameleonBtn BtnAmpliarGrafico 
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   7080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Ampliar Gráfico"
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
         MICON           =   "FrmEstadisticas.frx":3A03
         PICN            =   "FrmEstadisticas.frx":3A1F
         PICH            =   "FrmEstadisticas.frx":3D96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Tipo de Gráfica"
         Height          =   2655
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton Opt2DXY 
            BackColor       =   &H00EAEFEF&
            Caption         =   "XY 2D"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1860
            Width           =   975
         End
         Begin VB.OptionButton Opt2DPie 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Torta 2D"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   975
         End
         Begin VB.OptionButton Opt2DCombination 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Combinación 2D"
            Height          =   255
            Left            =   1200
            TabIndex        =   51
            Top             =   1860
            Width           =   1575
         End
         Begin VB.OptionButton Opt3DCombiantion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Combianción 3D"
            Height          =   375
            Left            =   1200
            TabIndex        =   50
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Opt2DStep 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Pasos 2D"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   2160
            Width           =   1095
         End
         Begin VB.OptionButton Opt3DStep 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Pasos 3D"
            Height          =   255
            Left            =   1200
            TabIndex        =   48
            Top             =   2220
            Width           =   1335
         End
         Begin VB.OptionButton Opt2DArea 
            BackColor       =   &H00EAEFEF&
            Caption         =   "2D Area"
            Height          =   375
            Left            =   1200
            TabIndex        =   47
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Opt3DArea 
            BackColor       =   &H00EAEFEF&
            Caption         =   "3D Area"
            Height          =   255
            Left            =   1200
            TabIndex        =   46
            Top             =   420
            Width           =   1095
         End
         Begin VB.OptionButton Opt2DLine 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Linea 2D"
            Height          =   375
            Left            =   1200
            TabIndex        =   45
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Opt3DLine 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Linea 3D"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Opt2DBar 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Barras 2D"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Opt3DBar 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Barras 3D"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Representación Gráfica"
         Height          =   7455
         Left            =   3000
         TabIndex        =   40
         Top             =   240
         Width           =   7575
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   7095
            Left            =   120
            OleObjectBlob   =   "FrmEstadisticas.frx":3FCB
            TabIndex        =   77
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.Label PrmReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "(                        )"
         Height          =   195
         Left            =   1680
         TabIndex        =   76
         Top             =   4080
         Width           =   1170
      End
      Begin VB.Label TotReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "(                        )"
         Height          =   195
         Left            =   1680
         TabIndex        =   75
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Promedio:"
         Height          =   195
         Index           =   3
         Left            =   870
         TabIndex        =   74
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Sumatoria:"
         Height          =   195
         Index           =   2
         Left            =   825
         TabIndex        =   73
         Top             =   3840
         Width           =   750
      End
      Begin VB.Label PtoSel 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "(                        )"
         Height          =   195
         Left            =   1680
         TabIndex        =   65
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Pto. Seleccionado:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   64
         Top             =   3360
         Width           =   1350
      End
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Index           =   2
      Left            =   120
      TabIndex        =   79
      Top             =   480
      Width           =   10695
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   375
         Left            =   7560
         TabIndex        =   80
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Obtener resultados"
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
         MICON           =   "FrmEstadisticas.frx":6420
         PICN            =   "FrmEstadisticas.frx":643C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SystemOncoAmerica.DMGrid DMGrid2 
         Height          =   1815
         Left            =   1080
         TabIndex        =   81
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3201
         Object.Width           =   8865
         Object.Height          =   1785
         Rows            =   0
         BackColor       =   15396847
      End
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   3495
         Left            =   360
         OleObjectBlob   =   "FrmEstadisticas.frx":66DB
         TabIndex        =   106
         Top             =   4200
         Width           =   10215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumatorio Total:"
         Height          =   195
         Index           =   6
         Left            =   1080
         TabIndex        =   101
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "       "
         Height          =   195
         Index           =   2
         Left            =   4575
         TabIndex        =   100
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "       "
         Height          =   195
         Index           =   1
         Left            =   3855
         TabIndex        =   99
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "        "
         Height          =   195
         Index           =   0
         Left            =   3090
         TabIndex        =   98
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "      "
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   92
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "       "
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   91
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "      "
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   90
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "        "
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   89
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "         "
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   88
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumatoria de cada grupo:"
         Height          =   195
         Index           =   5
         Left            =   1080
         TabIndex        =   87
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Varianza:"
         Height          =   195
         Index           =   4
         Left            =   1080
         TabIndex        =   86
         Top             =   3720
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de grupos:"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   85
         Top             =   3960
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desviación Estándar:"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   84
         Top             =   3480
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumatoria General:"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   83
         Top             =   3240
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promedio General:"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   82
         Top             =   3000
         Width           =   1545
      End
   End
End
Attribute VB_Name = "FrmTablaEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDiagnosticos As New ADODB.Recordset
Dim RsEstadiaje As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim Condicion As String
Dim X As Integer
Dim ArrayTemp()
Dim ArrayTemp2()
Public Band As Boolean
Dim verifica_sent As Boolean
Dim i As Integer
Dim ArrayIMNHQK(100, 100) As String

Sub IniDMDrid()
On Error Resume Next
DMGrid1.Cols = 4

DMGrid1.DColumnas(1).Caption = " Identificador"
DMGrid1.DColumnas(2).Caption = " Descripción"
DMGrid1.DColumnas(3).Caption = " Id"
DMGrid1.DColumnas(4).Caption = " EjeX"

DMGrid1.DColumnas(3).Visible = False
DMGrid1.DColumnas(4).Visible = False

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 80 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100) - 300


DMGrid2.Cols = 6

DMGrid2.DColumnas(3).Alignment = 1
DMGrid2.DColumnas(4).Alignment = 1
DMGrid2.DColumnas(5).Alignment = 1
DMGrid2.DColumnas(6).Alignment = 1
DMGrid2.DColumnas(3).Visible = False

DMGrid2.DColumnas(1).Caption = "Grupo"
DMGrid2.DColumnas(2).Caption = "Valor"
DMGrid2.DColumnas(3).Caption = "Promedio"
DMGrid2.DColumnas(4).Caption = "Valor-Prom"
DMGrid2.DColumnas(5).Caption = "(Valor-Prom)^2"
DMGrid2.DColumnas(6).Caption = "(Valor-Prom)^2) / (n-1)"

DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 20 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 10 / 100)
'DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 15 / 100)  ' MMMMMMMMMMMMMMMMMM
DMGrid2.DColumnas(4).Width = Val(DMGrid2.Width * 20 / 100)
DMGrid2.DColumnas(5).Width = Val(DMGrid2.Width * 20 / 100)
DMGrid2.DColumnas(6).Width = Val(DMGrid2.Width * 30 / 100) - 300

Label5(0).Left = Format(Val(DMGrid2.DColumnas(1).Width) + Val(DMGrid2.DColumnas(2).Width) + DMGrid2.Left, "#,##0.00000")  ' Valor-Prom
Label5(1).Left = Format(Val(DMGrid2.DColumnas(1).Width) + Val(DMGrid2.DColumnas(2).Width) + Val(DMGrid2.DColumnas(4).Width) + DMGrid2.Left, "#,##0.00000")
Label5(2).Left = Format(Val(DMGrid2.DColumnas(1).Width) + Val(DMGrid2.DColumnas(2).Width) + Val(DMGrid2.DColumnas(4).Width) + Val(DMGrid2.DColumnas(5).Width) + DMGrid2.Left, "#,##0.00000")


'Label5(0).Left = Format((Val(DMGrid2.DColumnas(4).Width) * 2) + DMGrid2.Left, "#,##0.00000")
'Label5(1).Left = Format((Val(DMGrid2.DColumnas(5).Width) * 3) + DMGrid2.Left, "#,##0.00000")
'Label5(2).Left = Format((Val(DMGrid2.DColumnas(5).Width) * 4) + DMGrid2.Left, "#,##0.00000")


Label5(0).Width = Val(DMGrid2.DColumnas(4).Width) - 50
Label5(1).Width = Val(DMGrid2.DColumnas(5).Width) - 50
Label5(2).Width = Val(DMGrid2.DColumnas(6).Width) - 50

End Sub

Sub Leer_Consultas_Reg()
On Error Resume Next

CSql = "SELECT * FROM Estadisticas_Reg"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

DMGrid1.Clear
DMGrid1.Rows = 0

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Nombre").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Consulta").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Id").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsTemp.Fields("EjeX").Value
    RsTemp.MoveNext
Wend
DMGrid1.RowBackColor 1, vbWhite
DMGrid1.PaintMGrid
End Sub

Sub Crear_Sentencia()
On Error Resume Next
Dim EQuery As String
Dim Cpdr, Cpdr2

If CboTipo(0).ListIndex = -1 Then Cpdr = " = "
If CboTipo(1).ListIndex = -1 Then Cpdr2 = " = "

If CboTipo(0).ListIndex = 0 Then Cpdr = " = "
If CboTipo(0).ListIndex = 1 Then Cpdr = " <> "
If CboTipo(0).ListIndex = 2 Then Cpdr = " < "
If CboTipo(0).ListIndex = 3 Then Cpdr = " > "
If CboTipo(0).ListIndex = 4 Then Cpdr = " <= "
If CboTipo(0).ListIndex = 5 Then Cpdr = " >= "

If CboTipo(1).ListIndex = 0 Then Cpdr2 = " = "
If CboTipo(1).ListIndex = 1 Then Cpdr2 = " <> "
If CboTipo(1).ListIndex = 2 Then Cpdr2 = " < "
If CboTipo(1).ListIndex = 3 Then Cpdr2 = " > "
If CboTipo(1).ListIndex = 4 Then Cpdr2 = " <= "
If CboTipo(1).ListIndex = 5 Then Cpdr2 = " >= "

' Si se condiciono de acuerdo a los Diagnosticos entonces...
Dim Senten_Sel As String

Senten_Sel = "SELECT dbo.Informe_Medico.Diagnotico, dbo.Informe_Medico.Sesiones, dbo.Informe_Medico.Fecha," & _
  " dbo.Informe_Medico.Estadiaje,  dbo.Informe_Medico.IdTipoCa," & _
  " dbo.Informe_Medico2.CP,  dbo.Informe_Medico2.T,  dbo.Informe_Medico2.N,  dbo.Informe_Medico2.M," & _
  " dbo.Informe_Medico2.Estadio,  dbo.Informe_Medico2.G,  dbo.Informe_Medico2.Gleason," & _
  " dbo.Informe_Medico2.Reseccion,  dbo.Informe_Medico2.Re,  dbo.Informe_Medico2.Rp,  dbo.Informe_Medico2.Her2Neu," & _
  " dbo.Informe_Medico2.EMA,  dbo.Informe_Medico2.VIM,  dbo.Informe_Medico2.CAE,  dbo.Informe_Medico2.[CERB-2]," & _
  " dbo.Informe_Medico2.P53,  dbo.Informe_Medico2.DESMINA,  dbo.Informe_Medico2.ACE,  dbo.Informe_Medico2.AFP," & _
  " dbo.Informe_Medico2.[PROT-S-100],  dbo.Informe_Medico2.PGP,  dbo.Informe_Medico2.CD31," & _
  " dbo.Informe_Medico2.CD34,  dbo.Informe_Medico2.CD117,  dbo.Informe_Medico2.CK5,  dbo.Informe_Medico2.CK6," & _
  " dbo.Informe_Medico2.CK7,  dbo.Informe_Medico2.CK20,  dbo.Informe_Medico2.[CAM 5,2]," & _
  " dbo.Informe_Medico2.[TTF-1],  dbo.Informe_Medico2.CROMOGRANINA,  dbo.Informe_Medico2.SINAPTOFISINA," & _
  " dbo.Informe_Medico2.CD56,  dbo.Informe_Medico2.CD57,  dbo.Informe_Medico2.EGFR,  dbo.Informe_Medico2.KIT," & _
  " dbo.Informe_Medico2.AE1,  dbo.Informe_Medico2.AE3,  dbo.Informe_Medico2.CK903,  dbo.Informe_Medico2.GFAP," & _
  " dbo.Informe_Medico2.SMA,  dbo.Informe_Medico2.CA199,  dbo.Informe_Medico2.CA125,  dbo.Informe_Medico2.CEA," & _
  " dbo.Informe_Medico2.[CEA-D14],  dbo.Informe_Medico2.[E-CAD],  dbo.Informe_Medico2.HCG," & _
  " dbo.Informe_Medico2.[HMB-45], dbo.Informe_Medico2.HPAP,  dbo.Informe_Medico2.WT1,  dbo.Informe_Medico2.[BEL-1]," & _
  " dbo.Informe_Medico2.[BEL-2],  dbo.Informe_Medico2.PRB,  dbo.Informe_Medico2.[ALK-1]," & _
  " dbo.Informe_Medico2.RA , dbo.Informe_Medico2.CD99MID2, dbo.Informe_Medico2.NSD, dbo.Informe_Medico2.LCACD45, " & _
  " dbo.Informe_Medico2.CD20L26,dbo.Informe_Medico2.CD79A, dbo.Informe_Medico2.CD45ROUCHL1, dbo.Informe_Medico2.CD3, " & _
  " dbo.Informe_Medico2.CD30KL1BERH2,dbo.Informe_Medico2.CD15LEUM1,dbo.Informe_Medico2.WT," & _
  " dbo.Informe_Medico2.OTROS, dbo.Informe_Medico2.IdTipoCancer,  dbo.Paciente.EdadP,  dbo.Paciente.SexoP " & _
  " FROM dbo.Informe_Medico INNER JOIN dbo.Informe_Medico2 ON (dbo.Informe_Medico.IdInforme = dbo.Informe_Medico2.IdInforme) " & _
  " INNER JOIN dbo.Paciente ON (dbo.Informe_Medico.IdPaciente = dbo.Paciente.IdPaciente) "
EQuery = ""
If Check2.Value Then
    Dim Anexado As String
    Dim Opcc As String
    
    Anexado = ""
    
    If Combo1.ListIndex > 0 Then
        
        If Combo2.ListIndex < 1 Then Combo2.ListIndex = 0
        
        If Combo2.ListIndex = 1 Then
            Anexado = " AND " & ArrayIMNHQK(Combo1.ListIndex - 1, 0) & " <> '0.' "
        ElseIf Combo2.ListIndex = 2 Then
            Anexado = " AND " & ArrayIMNHQK(Combo1.ListIndex - 1, 0) & " = '0.' "
        Else
            Anexado = " AND " & ArrayIMNHQK(Combo1.ListIndex - 1, 0) & " <> '' "
        End If
    End If
    
    EQuery = Senten_Sel & " WHERE IdTipoCancer = " & CboDiagnosticos.ItemData(CboDiagnosticos.ListIndex) & Anexado
    
ElseIf Check7.Value Then
    EQuery = Senten_Sel & " WHERE Diagnotico like '%" & TxtDiagPer.Text & "%'"
End If

' Si se condiciono de acuerdo a las sesiones entonces...
If Check3.Value Then
    ' si no se ha agregado una condicion entonces
    If InStr(1, EQuery, "WHERE") = 0 Then
        If OptSesiones(1).Value Then
            EQuery = Senten_Sel & " WHERE Sesiones >= " & CDbl(TxtNoSesiones1.Text) & " AND Sesiones <= " & CDbl(TxtNoSesiones2.Text)
        Else
            EQuery = Senten_Sel & " WHERE Sesiones " & Cpdr & CDbl(TxtNoSesiones.Text)
        End If
    Else
        ' si se condiciono de acuerdo al rango de sesiones entonces...
        If OptSesiones(1).Value Then
            EQuery = EQuery & " AND Sesiones >= " & CDbl(TxtNoSesiones1.Text) & " AND Sesiones <= " & CDbl(TxtNoSesiones2.Text)
        Else
            EQuery = EQuery & " AND Sesiones " & Cpdr & CDbl(TxtNoSesiones.Text)
        End If
    End If
End If

' Si se condiciono de acuerdo al estadiaje entonces...
'If Check4.Value Then
'    ' si no se ha agregado una condicion entonces
'    If InStr(1, EQuery, "WHERE") = 0 Then
'        EQuery = "SELECT Fecha,SexoP,EdadP,Estadiaje,Sesiones,Diagnotico FROM Informe_Medico INNER JOIN Paciente ON (Informe_Medico.IdPaciente = Paciente.IdPaciente) WHERE Estadiaje like '%" & CboEstadiaje.Text & "%'"
'    Else
'        EQuery = EQuery & " AND Estadiaje like '%" & CboEstadiaje.Text & "%'"
'    End If
'End If
If Check4.Value Then
    ' si no se ha agregado una condicion entonces
    If InStr(1, EQuery, "WHERE") = 0 Then
        EQuery = "SELECT Fecha,SexoP,EdadP,Estadiaje,Sesiones,Diagnotico FROM Informe_Medico INNER JOIN Paciente ON (Informe_Medico.IdPaciente = Paciente.IdPaciente) WHERE dbo.Informe_Medico2.CP like '" & CboEstadiaje.Text & "' OR dbo.Informe_Medico2.T like '" & CboEstadiaje.Text & "' OR dbo.Informe_Medico2.N like '" & CboEstadiaje.Text & "' OR  dbo.Informe_Medico2.M like '" & CboEstadiaje.Text & "' OR Estadio like '" & CboEstadiaje.Text & "'"
    Else
        EQuery = EQuery & " AND (dbo.Informe_Medico2.CP like '" & CboEstadiaje.Text & "' OR dbo.Informe_Medico2.T like '" & CboEstadiaje.Text & "' OR dbo.Informe_Medico2.N like '" & CboEstadiaje.Text & "' OR  dbo.Informe_Medico2.M like '" & CboEstadiaje.Text & "' OR Estadio like '" & CboEstadiaje.Text & "')"
    End If
End If

' Si se condiciono de acuerdo a las edades entonces...
If Check5.Value Then
    ' si no se ha agregado una condicion entonces
    If InStr(1, EQuery, "WHERE") = 0 Then
        ' si se condiciono de acuerdo al rango de edades entonces...
        If OptEdad(1).Value Then
            EQuery = Senten_Sel & " WHERE EdadP >= " & CDbl(Mid(CboRangoEdad.Text, 1, 3)) & " AND EdadP <= " & CDbl(Mid(CboRangoEdad.Text, 6, 3))
        Else
            EQuery = Senten_Sel & " WHERE EdadP " & Cpdr2 & CDbl(TxtEdad.Text)
        End If
        
        If CboRangoEdad.ListIndex = 10 Then EQuery = Senten_Sel & " WHERE EdadP >= " & CDbl(Mid(CboRangoEdad.Text, 1, 3))
        If CboRangoEdad.ListIndex = 11 Then EQuery = Senten_Sel
    Else
        ' si se condiciono de acuerdo al rango de edades entonces...
        If OptEdad(1).Value Then
            EQuery = EQuery & " AND EdadP >= " & CDbl(Mid(CboRangoEdad.Text, 1, 3)) & " AND EdadP <= " & CDbl(Mid(CboRangoEdad.Text, 6, 3))
        Else
            EQuery = EQuery & " AND EdadP " & Cpdr2 & CDbl(TxtEdad.Text)
        End If
        
        If CboRangoEdad.ListIndex >= 10 Then EQuery = EQuery & " AND EdadP >= " & CDbl(Mid(CboRangoEdad.Text, 1, 3))
    End If
End If

' Si se condiciono de acuerdo al SexoP entonces...
If Check6.Value Then
    ' si no se ha agregado una condicion entonces
    If InStr(1, EQuery, "WHERE") = 0 Then
        ' si se condiciono de acuerdo al rango de edades entonces...
        If OptSexo(0).Value Then
            EQuery = Senten_Sel & " WHERE SexoP = 1"
        ElseIf OptSexo(1).Value Then
            EQuery = Senten_Sel & " WHERE SexoP = 0"
        ElseIf OptSexo(2).Value Then
            EQuery = Senten_Sel & " WHERE (SexoP = 0 OR SexoP = 1)"
        End If
    Else
        ' si se condiciono de acuerdo al rango de edades entonces...
        If OptSexo(1).Value Then
            EQuery = EQuery & " AND SexoP = 0"
        ElseIf OptSexo(0).Value Then
            EQuery = EQuery & " AND SexoP = 1"
        ElseIf OptSexo(2).Value Then
            EQuery = EQuery & " AND (SexoP = 1 OR SexoP = 0)"
        End If
    End If
End If

' Si se condiciono de acuerdo al estadiaje entonces...
If Check1.Value Then
    ' si no se ha agregado una condicion entonces
    If InStr(1, EQuery, "WHERE") = 0 Then
        EQuery = Senten_Sel & " WHERE Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Fecha <='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
    Else
        EQuery = EQuery & " AND Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Fecha <='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
    End If
End If

TxtQuery.Text = EQuery & " GROUP BY Fecha,SexoP,EdadP,Estadiaje,Sesiones,Diagnotico,IdTipoCa,  CP,  T,  N,  M,  Estadio,  G,  Gleason,  Reseccion,  Re,  Rp,  Her2Neu,  EMA,  VIM,  CAE,  [CERB-2]," & _
  " P53,  DESMINA,  ACE,  AFP,  [PROT-S-100],  PGP,  CD31,  CD34,  CD117,  CK5,  CK6,  CK7,  CK20,  [CAM 5,2]," & _
  " [TTF-1],  CROMOGRANINA,  SINAPTOFISINA,  CD56,  CD57,  EGFR,  KIT,  AE1,  AE3,  CK903,  GFAP,  SMA," & _
  " CA199,  CA125,  CEA,  [CEA-D14],  [E-CAD],  HCG,  [HMB-45],  HPAP,  WT1,  [BEL-1],  [BEL-2],  PRB," & _
  " [ALK-1],  RA,  OTROS, CD99MID2, NSD,LCACD45,CD20L26,CD79A,CD45ROUCHL1,CD3,CD30KL1BERH2,CD15LEUM1,WT, IdTipoCancer "

Set RsTemp = CrearRS(TxtQuery.Text)
If RsTemp.RecordCount = 0 Then
    MsgBox "No se encontraron resultados para las estadisticas!", vbInformation + vbOKOnly, "Información"
    Band = False
Else
    Band = True
End If

End Sub


Private Sub BtnAmpliarGrafico_Click()
On Error Resume Next

If IsNull(ArrayTemp(1, 1)) Then Exit Sub

Dim TG As Integer
Dim TGBand As Boolean

TG = UBound(ArrayTemp, 1)
TGBand = False

For iii = 1 To TG
    If Not IsEmpty(ArrayTemp(iii, 1)) Then
        TGBand = True
    End If
Next

If TGBand = False Then Exit Sub

FrmAmpliarGrafico.MSChart1.ChartData = ArrayTemp
FrmAmpliarGrafico.MSChart1.ChartType = MSChart1.ChartType
FrmAmpliarGrafico.MSChart1.Refresh
FrmAmpliarGrafico.Show vbModal
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next

End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next
Dim Rsp As Byte


If DMGrid1.Row = 0 Then Exit Sub
If DMGrid1.Rows = 0 Then Exit Sub

Rsp = MsgBox("Se procederá a eliminar el registro cuyo identificador es:" & Chr(13) & DMGrid1.ValorCelda(DMGrid1.Row, 1), vbQuestion + vbYesNo, "Confirmar!")

If Rsp = vbNo Then Exit Sub

CSql = "DELETE FROM Estadisticas_Reg WHERE Id=" & DMGrid1.ValorCelda(DMGrid1.Row, 3)
Set RsTemp = CrearRS(CSql)

MsgBox "El registro ha sido eliminado!", vbInformation + vbOKOnly, "Operación Exitosa!"

Leer_Consultas_Reg
End Sub

Private Sub BtnEnlistar_Click()
On Error Resume Next
Dim Sumat As Integer
Dim Sumat2 As Integer
'On Error Resume Next
'Dim ConCat As String
'Crear_Sentencia
'
'If Band = False Then Exit Sub
'
'ConCat = ""
'
'If Check1.Value Then ConCat = ConCat & Check1.Caption
'If Check2.Value Then ConCat = ConCat & Check2.Caption
'If Check3.Value Then ConCat = ConCat & Check3.Caption
'If Check4.Value Then ConCat = ConCat & Check4.Caption
'If Check5.Value Then ConCat = ConCat & Check5.Caption
'If Check6.Value Then ConCat = ConCat & Check6.Caption
'
'If DMGrid1.Rows > 0 Then
'    If DMGrid1.ValorCelda(DMGrid1.Rows, 1) <> ConCat Then
'        MsgBox "Las opciones seleccionadas no son similares a la estadistica anterior!", vbInformation + vbOKOnly, "Información."
'        Exit Sub
'    ElseIf DMGrid1.ValorCelda(DMGrid1.Rows, 1) = ConCat And DMGrid1.ValorCelda(1, 2) = Trim(TxtQuery.Text) Then
'        MsgBox "Las opciones seleccionadas ya se encuentran en lista!", vbInformation + vbOKOnly, "Información."
'        Exit Sub
'    End If
'End If
'
'DMGrid1.Rows = DMGrid1.Rows + 1
'DMGrid1.ValorCelda(DMGrid1.Rows, 1) = ConCat
'DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Trim(TxtQuery.Text)
'DMGrid1.RowBackColor 1, vbWhite
'DMGrid1.PaintMGrid

If DMGrid1.Rows > 0 Then
    If DMGrid1.Row > 0 Then
        If Val(Mid(Trim(DMGrid1.ValorCelda(DMGrid1.Row, 4)), 1, 2)) <> 99 Then
            OptEjeX(Val(Mid(Trim(DMGrid1.ValorCelda(DMGrid1.Row, 4)), 1, 1))).Value = True
        Else
            MsgBox "Se mostraran los resultados segun el rengo de fechas elejida!", vbInformation + vbOKOnly, "Información"
            OptEjeX(4).Value = True
        End If
        
        TxtQuery.Text = Trim(DMGrid1.ValorCelda(DMGrid1.Row, 2))
        CSql = Trim(TxtQuery.Text)
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount = 0 Then
            MsgBox "No existen registros para Consulta Personalizada!", vbInformation + vbOKOnly, "No hay Datos!"
            Exit Sub
        End If
        
        Dim Fecha As Date
        
        FechaMax = "01/01/1900"
        FechaMin = "01/12/2050"
        
        While Not RsTemp.EOF
        
            If CDate(RsTemp.Fields("Fecha").Value) > CDate(FechaMax) Then
                FechaMax = Format(RsTemp.Fields("Fecha").Value, "dd/MM/yyyy")
            End If
            
            If CDate(RsTemp.Fields("Fecha").Value) < CDate(FechaMin) Then
                FechaMin = Format(RsTemp.Fields("Fecha").Value, "dd/MM/yyyy")
            End If
            
            RsTemp.MoveNext
            
        Wend
        
      ' MMMMMMMMMMMMMMMMMMMMMM Mostrar por Meses  MMMMMMMMMMMMMMMMMMMMMM
            MenorE = Format(FechaMin, "MM")
            'MayorE = Format(CDate(CDate(FechaMax) - CDate(FechaMin)), "MM")
            MayorE = Val(Format(CDate(CDate(Format(FechaMax, "MM/yyyy")) - CDate(Format(FechaMin, "MM/yyyy"))), "MM"))
            MayorE = MayorE + (CDate(Format(FechaMax, "yyyy")) - CDate(Format(FechaMin, "yyyy"))) * 12
            
            Fecha = Format(FechaMin, "dd/MM/yyyy")
            For J = 1 To MayorE
                If CDate(Format(Fecha, "MM/yyyy")) > CDate(Format(FechaMax, "MM/yyyy")) Then MayorE = J - 1: Exit For
                Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
            Next
            
            Fecha = Format(FechaMin, "dd/MM/yyyy")
            
            ReDim ArrayTemp(1 To MayorE, 0 To 1)
            For J = 1 To MayorE
                'If CDate(Format(Fecha, "MM/yyyy")) > CDate(Format(FechaMax, "MM/yyyy")) Then Exit For
                ArrayTemp(J, 0) = Format(CDate(Format(Fecha, "MM/yyyy")), "MM/yyyy")
                Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
            Next
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                For i = 1 To MayorE
                    If ArrayTemp(i, 0) = Format(RsTemp.Fields("Fecha").Value, "MM/yyyy") Then
                        ArrayTemp(i, 1) = Val(ArrayTemp(i, 1)) + 1
                        Exit For
                    End If
                Next i
                RsTemp.MoveNext
            Wend
        End If
      
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        Dim BuffMayorE As Integer
        Dim BuffMenorE As Integer
        Dim Div2 As Integer
        
        BuffMayorE = UBound(ArrayTemp, 1)
        BuffMenorE = LBound(ArrayTemp, 1)
        Div2 = 0
        For i = BuffMenorE To BuffMayorE
            If Val(ArrayTemp(i, 1)) <> 0 Then
                Sumat = Sumat + Val(ArrayTemp(i, 1))
                Div2 = Div2 + 1
            End If
        Next i
          
        TotReg.Caption = Sumat
        PrmReg.Caption = Format((Sumat / Div2), "#,##0.00")

     
      MSChart1.ChartData = ArrayTemp
      Opt2DBar.Value = True
      MSChart1.Refresh
      Option1(1).Value = True
      
    End If


End Sub

Private Sub BtnGraficar_Click()
On Error Resume Next
Dim RsGraficar As New ADODB.Recordset
Dim ContTmp(0 To 50) As Integer
Dim i As Integer
Dim EjeX As Boolean
verifica_sent = True

If verifica_sent Then Crear_Sentencia

If IsEmpty(TxtQuery.Text) Then
    CSql = ""
    Exit Sub
ElseIf (TxtQuery.Text) <> "" Then
    If InStr(1, TxtQuery.Text, "FROM") <> 0 Then
        CSql = Trim(TxtQuery.Text)
    Else
        CSql = ""
        Exit Sub
    End If
End If

Set RsGraficar = CrearRS(CSql)
If RsGraficar.RecordCount = 0 Then
    MsgBox "No hay Datos para realizar el Gráfico!!!", vbExclamation + vbOKOnly, "Error"
    BtnGuardar.Enabled = False
    Exit Sub
End If

BtnGuardar.Enabled = True
Band = False
For i = 0 To 4
    If OptEjeX(i).Value Then Band = True
Next i

If Band = False Then OptEjeX(4).Value = True

If OptEjeX(0).Value Then
    EjeX = True
    If OptEdad(0).Value Then
    
        ReDim ArrayTemp(0 To 2, 0 To 1)
        ArrayTemp(0, 0) = "..."
        ArrayTemp(1, 0) = "" & Val(TxtEdad.Text) & " años"
        ArrayTemp(2, 0) = "..."
        
        ContTmp(0) = 0
        
        While Not RsGraficar.EOF
            
            If CboTipo(1).ListIndex = 0 Then If Val(RsGraficar.Fields("EdadP").Value) = Val(TxtEdad.Text) Then ContTmp(0) = ContTmp(0) + 1
            If CboTipo(1).ListIndex = 1 Then If Val(RsGraficar.Fields("EdadP").Value) <> Val(TxtEdad.Text) Then ContTmp(0) = ContTmp(0) + 1
            If CboTipo(1).ListIndex = 2 Then If Val(RsGraficar.Fields("EdadP").Value) < Val(TxtEdad.Text) Then ContTmp(0) = ContTmp(0) + 1
            If CboTipo(1).ListIndex = 3 Then If Val(RsGraficar.Fields("EdadP").Value) > Val(TxtEdad.Text) Then ContTmp(0) = ContTmp(0) + 1
            If CboTipo(1).ListIndex = 4 Then If Val(RsGraficar.Fields("EdadP").Value) <= Val(TxtEdad.Text) Then ContTmp(0) = ContTmp(0) + 1
            If CboTipo(1).ListIndex = 5 Then If Val(RsGraficar.Fields("EdadP").Value) >= Val(TxtEdad.Text) Then ContTmp(0) = ContTmp(0) + 1
            
            RsGraficar.MoveNext
        Wend
        ArrayTemp(1, 1) = ContTmp(0)
    Else
        If CboGrupo.ListIndex = 0 Then
            If CboRangoEdad.ListIndex <> 11 Then
            
                ReDim ArrayTemp(0 To 11, 0 To 1)
                ContTmp(14) = Mid(CboRangoEdad.Text, 6, 3)
                ArrayTemp(0, 0) = "..."
                ArrayTemp(11, 0) = "..."
               
                For J = 1 To 10
                    ArrayTemp(J, 0) = Val(ContTmp(14)) - Abs(J - 10) & " años"
                Next
                
                While Not RsGraficar.EOF
                    ContTmp(13) = RsGraficar.Fields("EdadP").Value
                    
                    If Val(ContTmp(13)) = (ContTmp(14) - 9) Then
                        ContTmp(0) = Val(ContTmp(0)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 8) Then
                        ContTmp(1) = Val(ContTmp(1)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 7) Then
                        ContTmp(2) = Val(ContTmp(2)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 6) Then
                        ContTmp(3) = Val(ContTmp(3)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 5) Then
                        ContTmp(4) = Val(ContTmp(4)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 4) Then
                        ContTmp(5) = Val(ContTmp(5)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 3) Then
                        ContTmp(6) = Val(ContTmp(6)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 2) Then
                        ContTmp(7) = Val(ContTmp(7)) + 1
                    ElseIf Val(ContTmp(13)) = (ContTmp(14) - 1) Then
                        ContTmp(8) = Val(ContTmp(8)) + 1
                    ElseIf Val(ContTmp(13)) = ContTmp(14) Then
                        ContTmp(9) = Val(ContTmp(9)) + 1
                    End If
                    RsGraficar.MoveNext
                Wend
                
                For i = 0 To 9
                    ArrayTemp(i + 1, 1) = ContTmp(i)
                Next
                
            ' Si eligio todas las edades
            
            Else
              ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
              ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                Dim ContA As Integer
                Dim MenorE As Integer
                Dim MayorE As Integer

                MenorE = 150
                MayorE = 0

                While Not RsGraficar.EOF

                    If Val(RsGraficar.Fields("EdadP").Value) > MayorE Then MayorE = Val(RsGraficar.Fields("EdadP").Value)
                    If Val(RsGraficar.Fields("EdadP").Value) < MenorE Then MenorE = Val(RsGraficar.Fields("EdadP").Value)
                    RsGraficar.MoveNext

                Wend

                ReDim ArrayTemp(MenorE To MayorE, 0 To 1)

                For J = MenorE To MayorE
                    ArrayTemp(J, 0) = J & " años"
                Next

                RsGraficar.MoveFirst
                While Not RsGraficar.EOF
                    ContA = Val(RsGraficar.Fields("EdadP").Value)
                    ArrayTemp(ContA, 1) = ArrayTemp(ContA, 1) + 1
                    RsGraficar.MoveNext
                Wend
                
            End If
        ElseIf CboGrupo.ListIndex = 1 Then
            ReDim ArrayTemp(0 To 2, 0 To 1)
            ArrayTemp(0, 0) = "..."
            ArrayTemp(1, 0) = Trim(CboRangoEdad.Text)
            ArrayTemp(2, 0) = "..."
            ArrayTemp(1, 1) = RsGraficar.RecordCount
        Else
        'ElseIf CboGrupo.ListIndex = 2 Then
             '  MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
             '  MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                ReDim ArrayTemp(0 To 11, 0 To 1)
                ContTmp(14) = Mid(CboRangoEdad.Text, 6, 3)
                ArrayTemp(0, 0) = "..."
                ArrayTemp(12, 0) = "..."
               
                For J = 1 To 11
                    ArrayTemp(J, 0) = CboRangoEdad.List(J - 1)
                Next
                
                While Not RsGraficar.EOF
                    ContTmp(13) = RsGraficar.Fields("EdadP").Value
                    
                    For J = 1 To 11
                        If Val(Mid(ArrayTemp(J, 0), 1, 3)) <= Val(ContTmp(13)) And Val(Mid(ArrayTemp(J, 0), 6, 3)) >= Val(ContTmp(13)) Then
                            'ContTmp(j) = Val(ContTmp(j)) + 1
                            ArrayTemp(J, 1) = Val(ArrayTemp(J, 1)) + 1
                            Exit For
                        End If
                        
                        If J = 11 Then ArrayTemp(J, 1) = ArrayTemp(J, 1) + 1
                    Next
                    
                    RsGraficar.MoveNext
                Wend
        End If
    End If
ElseIf OptEjeX(1).Value Then
        EjeX = True
        ReDim ArrayTemp(0 To 3, 0 To 1)
        
        If OptSexo(0).Value Then
            ReDim ArrayTemp(0 To 2, 0 To 1)
            
            ArrayTemp(0, 0) = "..."
            ArrayTemp(1, 0) = "Femenino"
            ArrayTemp(2, 0) = "..."
            
            ContFem = 0
            While Not RsGraficar.EOF
                If Val(RsGraficar.Fields("SexoP").Value) = 1 Then
                    ContFem = ContFem + 1
                End If
                RsGraficar.MoveNext
            Wend
            ArrayTemp(1, 1) = ContFem
            
        ElseIf OptSexo(1).Value Then
            ReDim ArrayTemp(0 To 2, 0 To 1)
            
            ArrayTemp(0, 0) = "...": ArrayTemp(1, 0) = "Masculino": ArrayTemp(2, 0) = "..."
            
            ContMas = 0
            While Not RsGraficar.EOF
                If Val(RsGraficar.Fields("SexoP").Value) = 0 Then
                    ContMas = ContMas + 1
                End If
                RsGraficar.MoveNext
            Wend
            ArrayTemp(1, 1) = ContMas
'        ElseIf OptSexo(2).Value Then
        Else
            ReDim ArrayTemp(0 To 4, 0 To 1)
            ArrayTemp(0, 0) = "..."
            ArrayTemp(1, 0) = "Masculino"
            ArrayTemp(2, 0) = "Femenino"
            ArrayTemp(3, 0) = "Total"
            ArrayTemp(4, 0) = "..."
            
            ContMas = 0
            ContFem = 0
            While Not RsGraficar.EOF
                If Val(RsGraficar.Fields("SexoP").Value) = 0 Then
                    ContMas = ContMas + 1
                Else
                    ContFem = ContFem + 1
                End If
                RsGraficar.MoveNext
            Wend
            ArrayTemp(1, 1) = ContMas
            ArrayTemp(2, 1) = ContFem
            ArrayTemp(3, 1) = ContMas + ContFem
        End If

' Eje "X", Estadiaje
ElseIf OptEjeX(3).Value Then
        
        EjeX = True
        ReDim ArrayTemp(0 To 2, 0 To 1)
        
        ArrayTemp(0, 0) = "..."
        ArrayTemp(1, 0) = Trim(CboEstadiaje.Text)
        ArrayTemp(2, 0) = "..."
        
        ContMas = 0
        ContFem = 0
        While Not RsGraficar.EOF
            If InStr(1, RsGraficar.Fields("Estadiaje").Value, Trim(CboEstadiaje.Text)) <> 0 Then
                ContMas = ContMas + 1
            End If
            RsGraficar.MoveNext
        Wend
        ArrayTemp(1, 1) = ContMas
'Eje "X", Fechas
ElseIf OptEjeX(4).Value Then
        
        EjeX = True
        Dim Fecha
        If CboTipoFecha.ListIndex = -1 Then CboTipoFecha.ListIndex = 0
        ' Mostrar por Meses
        If CboTipoFecha.ListIndex = 0 Then
            MenorE = Format(DTPicker1.Value, "MM")
            'MayorE = Format(CDate(CDate(DTPicker2.Value) - CDate(DTPicker1.Value)), "MM")
            MayorE = Val(Format(CDate(CDate(Format(DTPicker2.Value, "MM/yyyy")) - CDate(Format(DTPicker1.Value, "MM/yyyy"))), "MM"))
            MayorE = MayorE + (CDate(Format(DTPicker2.Value, "yyyy")) - CDate(Format(DTPicker1.Value, "yyyy"))) * 12
            
            Fecha = Format(DTPicker1.Value, "dd/MM/yyyy")
            For J = 1 To MayorE
                If CDate(Format(Fecha, "MM/yyyy")) > CDate(Format(DTPicker2.Value, "MM/yyyy")) Then MayorE = J - 1: Exit For
                Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
            Next
            
            Fecha = Format(DTPicker1.Value, "dd/MM/yyyy")
            
            ReDim ArrayTemp(1 To MayorE, 0 To 1)
            For J = 1 To MayorE
                'If CDate(Format(Fecha, "MM/yyyy")) > CDate(Format(DTPicker2.Value, "MM/yyyy")) Then Exit For
                ArrayTemp(J, 0) = Format(CDate(Format(Fecha, "MM/yyyy")), "MM/yyyy")
                Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
            Next
            
            While Not RsGraficar.EOF
                For i = 1 To MayorE
                    If ArrayTemp(i, 0) = Format(RsGraficar.Fields("Fecha").Value, "MM/yyyy") Then
                        ArrayTemp(i, 1) = Val(ArrayTemp(i, 1)) + 1
                        Exit For
                    End If
                Next i
                RsGraficar.MoveNext
            Wend
            
        'Mostrar por Años
        ElseIf CboTipoFecha.ListIndex = 1 Then
            
            MenorE = Format(DTPicker1.Value, "yyyy")
            MayorE = Format(DTPicker2.Value, "yyyy")
            
            ReDim ArrayTemp(0 To (MayorE - MenorE), 0 To 1)
            
            For J = 0 To (MayorE - MenorE)
                ArrayTemp(J, 0) = " " & Trim(Val(MenorE) + J) & " "
            Next
            
            
            While Not RsGraficar.EOF
            
                For i = 0 To (MayorE - MenorE)
                    If Replace(ArrayTemp(i, 0), " ", "") = Format(RsGraficar.Fields("Fecha").Value, "yyyy") Then
                        ArrayTemp(i, 1) = Val(ArrayTemp(i, 1)) + 1
                        Exit For
                    End If
                Next i
                
                RsGraficar.MoveNext
            
            Wend
            
        End If
        
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    Dim BuffMayorE As Integer
    Dim BuffMenorE As Integer
    Dim Div2 As Integer
    
    BuffMayorE = UBound(ArrayTemp, 1)
    BuffMenorE = LBound(ArrayTemp, 1)
    Div2 = 0
    For i = BuffMenorE To BuffMayorE
        If Val(ArrayTemp(i, 1)) <> 0 Then
            Sumat = Sumat + Val(ArrayTemp(i, 1))
            Div2 = Div2 + 1
        End If
    Next i
      
    TotReg.Caption = Sumat
    PrmReg.Caption = Format((Sumat / Div2), "#,##0.00")
    
    If EjeX = False Then MsgBox "Debe elegir una opcion para el Eje de las 'X'", vbInformation + vbOKOnly, "Eliga el Eje 'X'": Exit Sub
    MSChart1.ChartData = ArrayTemp
    MSChart2.ChartData = ArrayTemp
    If Opt3DBar.Value = True Then Opt3DBar_Click
    MSChart1.Refresh
    MSChart2.Refresh

    Option1(1).Value = True
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

End Sub


Private Sub BtnGuardar_Click()
On Error Resume Next
Dim Rsp As Byte
Dim rsp2 As String


If IsEmpty(TxtQuery.Text) Then
    CSql = ""
    Exit Sub
ElseIf (TxtQuery.Text) <> "" Then
    If InStr(1, TxtQuery.Text, "FROM") <> 0 Then
        CSql = Trim(TxtQuery.Text)
    Else
        CSql = ""
        MsgBox "La consulta no contiene resultados!!!", vbExclamation + vbOKOnly, "Operación Fallida!"
        Exit Sub
    End If
ElseIf (TxtQuery.Text) = "" Then
    MsgBox "La consulta no contiene resultados!!!", vbExclamation + vbOKOnly, "Operación Fallida!"
    Exit Sub
End If

Set RsGraficar = CrearRS(CSql)
If RsGraficar.RecordCount = 0 Then
    MsgBox "La consulta no contiene resultados!!!", vbExclamation + vbOKOnly, "Operación Fallida!"
    BtnGuardar.Enabled = False
    Exit Sub
End If


Rsp = MsgBox("Se procederá a guardar la consulta estadística, para ello se necesita" & Chr(13) & _
             "que ingrese un nombre para identificar dicha consulta, Desea Continuar?", vbQuestion + vbYesNo, "Confirmación")
If Rsp = vbNo Then Exit Sub

rsp2 = InputBox("Ingrese el nombre que identificará la consulta a guardar:", "Guardar nueva consulta estadística")


If Trim(rsp2) <> "" Then
    
    Dim Band2 As Boolean
    
    CSql = "SELECT MAX(id)+1 AS NuevoId FROM Estadisticas_Reg"
    Set RsTemp = CrearRS(CSql)
    
    Dim NuevoId  As Integer
    Dim Teempo As String
    
    If Not IsNull(RsTemp.Fields(0).Value) Then
        NuevoId = Val(RsTemp.Fields(0).Value)
    Else
        NuevoId = 1
    End If
    
    Band2 = False
    
    For ii = 0 To OptEjeX.Count - 1
        If OptEjeX(i).Visible = True And OptEjeX(i).Enabled = True Then
            Band2 = True
            Teempo = i & ":" & OptEjeX(i).Caption
            Exit For
        End If
    Next ii
    
'    If Band2 = False Then
'        CSql = "INSERT INTO Estadisticas_Reg (Id,Nombre,Consulta,EjeX,IdUser) VALUES (" & NuevoId & ",'" & Trim(rsp2) & _
'            "','" & Trim(TxtQuery.Text) & "','N/A'," & IdUser & ")"
'        Set RsTemp = CrearRS(CSql)
'    Else
'        CSql = "INSERT INTO Estadisticas_Reg (Id,Nombre,Consulta,EjeX,IdUser) VALUES (" & NuevoId & ",'" & Trim(rsp2) & _
'            "','" & Trim(TxtQuery.Text) & "','" & Teempo & "'," & IdUser & ")"
'        Set RsTemp = CrearRS(CSql)
'    End If
    CSql = "SELECT * FROM Estadisticas_Reg"
    Set RsTemp = CrearRS(CSql)

    RsTemp.AddNew
    RsTemp.Fields("Id").Value = Trim(NuevoId)
    RsTemp.Fields("Nombre").Value = Trim(rsp2)
    RsTemp.Fields("Consulta").Value = Trim(TxtQuery.Text)

    If Band2 = False Then
        RsTemp.Fields("EjeX").Value = "N/A"
    Else
        RsTemp.Fields("EjeX").Value = Teempo
    End If
    RsTemp.Fields("IdUser").Value = IdUser
    RsTemp.Update
    
    MsgBox "Se realizaron cambios!", vbInformation + vbOKOnly, "Operación Exitosa!"
Else
    MsgBox "No se realizaron cambios!", vbInformation + vbOKOnly, "Información"
End If

Leer_Consultas_Reg
End Sub

Private Sub BtnImprimir_Click()
On Error Resume Next
End Sub

Private Sub ChameleonBtn1_Click()
On Error Resume Next
Dim Colmnas As Integer
Dim Colmnas2 As Integer

Dim NumCols As Integer
Dim SumTotCols As Integer
Dim PromTotCols As Double
Dim Desv As Double

DMGrid2.Clear
DMGrid2.Rows = 0

Colmnas = UBound(ArrayTemp, 1)
Colmnas2 = UBound(ArrayTemp, 2)

For i = 0 To Colmnas
    If ArrayTemp(i, 0) <> "..." Then
        NumCols = NumCols + 1
        SumTotCols = SumTotCols + Val(ArrayTemp(i, 1))
    End If
Next
PromTotCols = SumTotCols / NumCols

For i = 0 To Colmnas
    If ArrayTemp(i, 0) <> "..." Then
        Desv = Desv + ((Val(ArrayTemp(i, 1)) - PromTotCols) ^ 2) / (NumCols - 1)
    End If
Next
Desv = Sqr(Desv)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMM Llena el DMGrid2 MMMMMMMMMMMMMMMMMMMMMM

Label5(0).Caption = "0"
Label5(1).Caption = "0"
Label5(2).Caption = "0"
For i = 0 To Colmnas
    If ArrayTemp(i, 0) <> "..." Then
        DMGrid2.Rows = DMGrid2.Rows + 1
        DMGrid2.ValorCelda(DMGrid2.Rows, 1) = ArrayTemp(i, 0)
        DMGrid2.ValorCelda(DMGrid2.Rows, 2) = Val(ArrayTemp(i, 1))
        DMGrid2.ValorCelda(DMGrid2.Rows, 3) = Format(PromTotCols, "#,##0.00000")
        DMGrid2.ValorCelda(DMGrid2.Rows, 4) = Format((ArrayTemp(i, 1)) - PromTotCols, "#,##0.00000")
        DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(((ArrayTemp(i, 1)) - PromTotCols) ^ 2, "#,##0.00000")
        DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format((((ArrayTemp(i, 1)) - PromTotCols) ^ 2) / (NumCols - 1), "#,##0.00000")
        
        Label5(0).Caption = Format(CDbl(Label5(0).Caption) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 4)), "#,##0.00000")
        Label5(1).Caption = Format(CDbl(Label5(1).Caption) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 5)), "#,##0.00000")
        Label5(2).Caption = Format(CDbl(Label5(2).Caption) + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 6)), "#,##0.00000")
    End If
Next


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Label4(0).Caption = Format(PromTotCols, "#,##0.00000")
Label4(1).Caption = Format(SumTotCols, "#,##0.00000")
Label4(2).Caption = Format(Desv, "#,##0.00000")
Label4(3).Caption = Format(Desv * 2, "#,##0.00000")
Label4(4).Caption = NumCols


DMGrid2.RowBackColor 1, vbWhite
DMGrid2.PaintMGrid

End Sub


Private Sub CboRangoEdad_Click()
On Error Resume Next
CboGrupo.Clear
If CboRangoEdad.List(CboRangoEdad.ListIndex) <> "Todas las edades" Then
    CboGrupo.AddItem "Suseción"
    CboGrupo.AddItem "Único"
Else
    CboGrupo.AddItem "Suseción"
    CboGrupo.AddItem "Único"
    CboGrupo.AddItem "Grp. Etarios"
End If
End Sub

Private Sub Check1_Click()
On Error Resume Next
'CboTipoFecha.Enabled = Check1.Value
'OptEjeX(4).Enabled = Check1.Value
DTPicker1.Enabled = Check1.Value
DTPicker2.Enabled = Check1.Value
End Sub

Private Sub Check2_Click()
On Error Resume Next
CboDiagnosticos.Enabled = Check2.Value
Combo1.Enabled = Check2.Value
Combo2.Enabled = Check2.Value
Combo1.ListIndex = 0
Combo2.ListIndex = 0
OptEjeX(2).Enabled = Check2.Value
Check7.Value = Check2.Value - 1
End Sub

Private Sub Check3_Click()
On Error Resume Next
CboTipo(0).Enabled = Check3.Value
If Check3.Value Then
    TxtNoSesiones.Enabled = True
    TxtNoSesiones.Text = "1"
    OptSesiones(0).Enabled = True
    OptSesiones(1).Enabled = True
    OptSesiones(0).Value = True
Else
    OptSesiones(0).Value = False
    OptSesiones(1).Value = False
    OptSesiones(0).Enabled = False
    OptSesiones(1).Enabled = False
    TxtNoSesiones.Enabled = False
    TxtNoSesiones1.Enabled = False
    TxtNoSesiones2.Enabled = False
End If
End Sub
Private Sub Check4_Click()
On Error Resume Next
CboEstadiaje.Enabled = Check4.Value
OptEjeX(3).Enabled = Check4.Value
End Sub

Private Sub Check5_Click()
On Error Resume Next
'OptEjeX(0).Enabled = Check5.Value
CboTipo(1).Enabled = Check5.Value
If Check5.Value Then
    TxtEdad.Enabled = True
    TxtEdad.Text = "35"
    OptEdad(0).Enabled = True
    OptEdad(1).Enabled = True
    OptEdad(0).Value = True
    CboGrupo.Enabled = True
Else
    OptEdad(0).Enabled = False
    OptEdad(1).Enabled = False
    OptEdad(0).Value = False
    OptEdad(1).Value = False
    TxtEdad.Enabled = False
    CboRangoEdad.Enabled = False
    CboGrupo.Enabled = False
End If

CboGrupo.ListIndex = 0
End Sub

Private Sub Check6_Click()
On Error Resume Next
OptSexo(0).Enabled = Check6.Value
OptSexo(1).Enabled = Check6.Value
OptSexo(2).Enabled = Check6.Value
'OptEjeX(1).Enabled = Check6.Value

If Check6.Value Then
    OptSexo(0).Value = True
Else
    OptSexo(0).Value = False
    OptSexo(1).Value = False
    OptSexo(2).Value = False
End If
End Sub



Private Sub Check7_Click()
On Error Resume Next
Check2.Value = Check7.Value - 1

If Check7.Value Then
    TxtDiagPer.Enabled = True
    TxtDiagPer.BackColor = vbWhite
Else
    TxtDiagPer.Enabled = False
    TxtDiagPer.BackColor = &HE0E0E0
End If

End Sub

Private Sub Combo1_Change()
If Combo1.ListIndex = 0 Then Combo2.ListIndex = 0
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then Combo2.ListIndex = 0
End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex = 0 Then Combo2.ListIndex = 0
End Sub

Private Sub DMGrid1_DobleClick()
On Error Resume Next
If DMGrid1.Row <> 0 Then
    NroReg.Caption = DMGrid1.Row & " / " & DMGrid1.Rows
    
    Dim TamShow As String
    Dim TamShow2 As String
    Dim Baand As Boolean
    Dim Rsp As Byte
    
    TamShow2 = DMGrid1.ValorCelda(DMGrid1.Row, DMGrid1.Col)
    Baand = False
    While Len(TamShow2) > 200
        Baand = True
        If Len(TamShow2) >= 200 Then
            TamShow = TamShow & Mid(TamShow2, 1, 200) & Chr(13)
            TamShow2 = Replace(TamShow2, Mid(TamShow2, 1, 39), "")
        End If
    Wend
    
    If Baand = False Then TamShow = TamShow2
    
'    If DMGrid1.Col = 2 Then
'        rsp = MsgBox(TamShow & Chr(13) & "Desea trabajar con la consulta seleccionada?", _
'        vbInformation + vbYesNo, "Información sobre: " & DMGrid1.DColumnas(DMGrid1.Col).Caption)
'
'        If rsp = vbNo Then Exit Sub
'        TxtQuery.Text = Trim(DMGrid1.ValorCelda(DMGrid1.Row, 2))
'    Else
        MsgBox TamShow, vbInformation + vbOKOnly, "Información sobre: " & DMGrid1.DColumnas(DMGrid1.Col).Caption
'    End If
End If
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
On Error Resume Next

If lRow <> 0 Then
    NroReg.Caption = DMGrid1.Row & " / " & DMGrid1.Rows
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Centrar Me
IniDMDrid
Diagnostico
Estadiaje
Leer_Consultas_Reg
DTPicker1.Value = Now - 182
DTPicker2.Value = Now + 182

ArrayIMNHQK(0, 0) = "RE"
ArrayIMNHQK(0, 1) = " dbo.Informe_Medico2.Re "
ArrayIMNHQK(1, 0) = "RP"
ArrayIMNHQK(1, 1) = " dbo.Informe_Medico2.Rp "
ArrayIMNHQK(2, 0) = "HER/2 - NEU"
ArrayIMNHQK(2, 1) = " dbo.Informe_Medico2.Her2Neu "
ArrayIMNHQK(3, 0) = "EMA"
ArrayIMNHQK(3, 1) = " dbo.Informe_Medico2.EMA "
ArrayIMNHQK(4, 0) = "VIM"
ArrayIMNHQK(4, 1) = " dbo.Informe_Medico2.VIM "
ArrayIMNHQK(5, 0) = "CAE"
ArrayIMNHQK(5, 1) = " dbo.Informe_Medico2.CAE "
ArrayIMNHQK(6, 0) = "CERB-2"
ArrayIMNHQK(6, 1) = " dbo.Informe_Medico2.[CERB-2] "
ArrayIMNHQK(7, 0) = "P53"
ArrayIMNHQK(7, 1) = " dbo.Informe_Medico2.P53 "
ArrayIMNHQK(8, 0) = "DESMINA"
ArrayIMNHQK(8, 1) = " dbo.Informe_Medico2.DESMINA "
ArrayIMNHQK(9, 0) = "ACE"
ArrayIMNHQK(9, 1) = " dbo.Informe_Medico2.ACE "
ArrayIMNHQK(10, 0) = "AFP"
ArrayIMNHQK(10, 1) = " dbo.Informe_Medico2.AFP"
ArrayIMNHQK(11, 0) = "PROT-S-100"
ArrayIMNHQK(11, 1) = " dbo.Informe_Medico2.[PROT-S-100] "
ArrayIMNHQK(12, 0) = "PGP"
ArrayIMNHQK(12, 1) = " dbo.Informe_Medico2.PGP "
ArrayIMNHQK(13, 0) = "CD31"
ArrayIMNHQK(13, 1) = " dbo.Informe_Medico2.CD31 "
ArrayIMNHQK(14, 0) = "CD34"
ArrayIMNHQK(14, 1) = " dbo.Informe_Medico2.CD34 "
ArrayIMNHQK(15, 0) = "CD117"
ArrayIMNHQK(15, 1) = " dbo.Informe_Medico2.CD117 "
ArrayIMNHQK(16, 0) = "CK5"
ArrayIMNHQK(16, 1) = " dbo.Informe_Medico2.CK5 "
ArrayIMNHQK(17, 0) = "CK6"
ArrayIMNHQK(17, 1) = " dbo.Informe_Medico2.CK6 "
ArrayIMNHQK(18, 0) = "CK7"
ArrayIMNHQK(18, 1) = " dbo.Informe_Medico2.CK7 "
ArrayIMNHQK(19, 0) = "CK20"
ArrayIMNHQK(19, 1) = " dbo.Informe_Medico2.CK20 "
ArrayIMNHQK(20, 0) = " CAM 5, 2"
ArrayIMNHQK(20, 1) = " dbo.Informe_Medico2.[CAM 5,2] "
ArrayIMNHQK(21, 0) = "TTF-1"
ArrayIMNHQK(21, 1) = " dbo.Informe_Medico2.[TTF-1] "
ArrayIMNHQK(22, 0) = "CROMOGRANINA"
ArrayIMNHQK(22, 1) = " dbo.Informe_Medico2.CROMOGRANINA "
ArrayIMNHQK(23, 0) = "SINAPTOFISINA"
ArrayIMNHQK(23, 1) = " dbo.Informe_Medico2.SINAPTOFISINA "
ArrayIMNHQK(24, 0) = "CD56"
ArrayIMNHQK(24, 1) = " dbo.Informe_Medico2.CD56 "
ArrayIMNHQK(25, 0) = "CD57"
ArrayIMNHQK(25, 1) = " dbo.Informe_Medico2.CD57 "
ArrayIMNHQK(26, 0) = "EGFR"
ArrayIMNHQK(26, 1) = " dbo.Informe_Medico2.EGFR "
ArrayIMNHQK(27, 0) = "KIT"
ArrayIMNHQK(27, 1) = " dbo.Informe_Medico2.KIT "
ArrayIMNHQK(28, 0) = "AE1"
ArrayIMNHQK(28, 1) = " dbo.Informe_Medico2.AE1 "
ArrayIMNHQK(29, 0) = "AE3"
ArrayIMNHQK(29, 1) = " dbo.Informe_Medico2.AE3 "
ArrayIMNHQK(30, 0) = "CK903"
ArrayIMNHQK(30, 1) = " dbo.Informe_Medico2.CK903 "
ArrayIMNHQK(31, 0) = "GFAP"
ArrayIMNHQK(31, 1) = " dbo.Informe_Medico2.GFAP "
ArrayIMNHQK(32, 0) = "SMA"
ArrayIMNHQK(32, 1) = " dbo.Informe_Medico2.SMA "
ArrayIMNHQK(33, 0) = "CD199"
ArrayIMNHQK(33, 1) = " dbo.Informe_Medico2.CD199 "
ArrayIMNHQK(34, 0) = "CA125"
ArrayIMNHQK(34, 1) = " dbo.Informe_Medico2.CA125 "
ArrayIMNHQK(35, 0) = "CEA"
ArrayIMNHQK(35, 1) = " dbo.Informe_Medico2.CEA "
ArrayIMNHQK(36, 0) = "CEA-D14"
ArrayIMNHQK(36, 1) = " dbo.Informe_Medico2.[CEA-D14] "
ArrayIMNHQK(37, 0) = "E-CAD"
ArrayIMNHQK(37, 1) = " dbo.Informe_Medico2.[E-CAD]"
ArrayIMNHQK(38, 0) = "HCG"
ArrayIMNHQK(38, 1) = " dbo.Informe_Medico2.HCG "
ArrayIMNHQK(39, 0) = "HMB-45"
ArrayIMNHQK(39, 1) = " dbo.Informe_Medico2.[HMB-45] "
ArrayIMNHQK(40, 0) = "HPAP"
ArrayIMNHQK(40, 1) = " dbo.Informe_Medico2.HPAP "
ArrayIMNHQK(41, 0) = "WT1"
ArrayIMNHQK(41, 1) = " dbo.Informe_Medico2.WT1 "
ArrayIMNHQK(42, 0) = "BEL-1"
ArrayIMNHQK(42, 1) = " dbo.Informe_Medico2.[BEL-1] "
ArrayIMNHQK(43, 0) = "BEL-2"
ArrayIMNHQK(43, 1) = " dbo.Informe_Medico2.[BEL-2] "
ArrayIMNHQK(44, 0) = "PRB"
ArrayIMNHQK(44, 1) = " dbo.Informe_Medico2.PRB "
ArrayIMNHQK(45, 0) = "ALK-1"
ArrayIMNHQK(45, 1) = " dbo.Informe_Medico2.[ALK-1] "
ArrayIMNHQK(46, 0) = "RA"
ArrayIMNHQK(46, 1) = " dbo.Informe_Medico2.RA "
ArrayIMNHQK(47, 0) = "CD99/MIC-1"
ArrayIMNHQK(47, 1) = " dbo.Informe_Medico2.CD99MID2 "
ArrayIMNHQK(48, 0) = "NSD"
ArrayIMNHQK(48, 1) = " dbo.Informe_Medico2.NSD "
ArrayIMNHQK(49, 0) = "LCA/CD45"
ArrayIMNHQK(49, 1) = " dbo.Informe_Medico2.LCACD45 "
ArrayIMNHQK(50, 0) = "CD20L26"
ArrayIMNHQK(50, 1) = " dbo.Informe_Medico2.CD20L26 "
ArrayIMNHQK(51, 0) = "CD79A"
ArrayIMNHQK(51, 1) = " dbo.Informe_Medico2.CD79A "
ArrayIMNHQK(52, 0) = "CD45-RO / UCHL-1"
ArrayIMNHQK(52, 1) = " dbo.Informe_Medico2.CD45ROUCHL1 "
ArrayIMNHQK(53, 0) = "CD3"
ArrayIMNHQK(53, 1) = " dbo.Informe_Medico2.CD3 "
ArrayIMNHQK(54, 0) = "CD30/KL-1/BERH-2"
ArrayIMNHQK(54, 1) = " dbo.Informe_Medico2.CD30KL1BERH2 "
ArrayIMNHQK(55, 0) = "CD15/LEUM1"
ArrayIMNHQK(55, 1) = " dbo.Informe_Medico2.CD15LEUM1 "
ArrayIMNHQK(56, 0) = "WT"
ArrayIMNHQK(56, 1) = " dbo.Informe_Medico2.WT "
ArrayIMNHQK(57, 0) = "OTROS"
ArrayIMNHQK(57, 1) = " dbo.Informe_Medico2.OTROS "


Combo1.Clear
Combo1.AddItem " "

For ii = 1 To 58
    Combo1.AddItem ArrayIMNHQK(ii - 1, 0)
Next

End Sub

 


Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
On Error Resume Next
MSChart1.Column = Series
MSChart1.Row = DataPoint
PtoSel.Caption = MSChart1.Data
End Sub

Private Sub Opt2DArea_Click()
On Error Resume Next
If Opt2DArea.Value = True Then
    MSChart1.ChartType = 5
End If
End Sub

Private Sub Opt2DBar_Click()
On Error Resume Next
If Opt2DBar.Value = True Then
    MSChart1.ChartType = 1
End If
End Sub

Private Sub Opt2DCombination_Click()
On Error Resume Next
If Opt2DCombination.Value = True Then
    MSChart1.ChartType = 9
End If
End Sub

Private Sub Opt2DLine_Click()
On Error Resume Next
If Opt2DLine.Value = True Then
    MSChart1.ChartType = 3
End If
End Sub
Private Sub Opt2DPie_Click()
On Error Resume Next
If Opt2DPie.Value = True Then
    MSChart1.ChartType = 14
End If
End Sub

Private Sub Opt2DStep_Click()
On Error Resume Next
If Opt2DStep.Value = True Then
    MSChart1.ChartType = 7
End If
End Sub
Private Sub Opt2DXY_Click()
On Error Resume Next
If Opt2DXY.Value = True Then
    MSChart1.ChartType = 16
End If
End Sub

Private Sub Opt3DArea_Click()
On Error Resume Next
If Opt3DArea.Value = True Then
    MSChart1.ChartType = 4
End If
End Sub


Private Sub Opt3DBar_Click()
On Error Resume Next
If Opt3DBar.Value = True Then
    MSChart1.ChartType = 0
End If
End Sub
Private Sub Opt3DCombiantion_Click()
On Error Resume Next
If Opt3DCombiantion.Value = True Then
    MSChart1.ChartType = 8
End If
End Sub

Private Sub Opt3DLine_Click()
On Error Resume Next
If Opt3DLine.Value = True Then
    MSChart1.ChartType = 2
End If
End Sub

Private Sub Opt3DStep_Click()
On Error Resume Next
If Opt3DStep.Value = True Then
    MSChart1.ChartType = 6
End If
End Sub

Private Sub OptEdad_Click(Index As Integer)
On Error Resume Next
CboGrupo.Enabled = OptEdad(1).Value

If OptEdad(0).Value Then
    TxtEdad.Enabled = True
    TxtEdad.Text = "35"
    CboRangoEdad.Enabled = False
    CboRangoEdad.ListIndex = -1
    CboTipo(1).Enabled = True
ElseIf OptEdad(1).Value Then
    CboRangoEdad.Enabled = True
    CboRangoEdad.ListIndex = 0
    TxtEdad.Enabled = False
    CboTipo(1).Enabled = False
End If
End Sub


Private Sub OptEjeX_Click(Index As Integer)
If Index <> 4 Then
    For i = 0 To OptEjeX2.Count - 1
        OptEjeX2(i).Value = False
    Next i
End If
End Sub

Private Sub OptEjeX2_Click(Index As Integer)
OptEjeX(4).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
For i = 0 To Option1.Count - 1
    If i <> Index Then
        Option1(i).Value = False
        FrameDato(i).Visible = False
    Else
        Option1(i).Value = True
        FrameDato(i).Visible = True
    End If
Next

End Sub

Sub Diagnostico()
On Error Resume Next
'CSql = "Select distinct (Diagnotico) From Informe_Medico"
'Set RsDiagnosticos = CrearRS(CSql)
'
'Do While Not RsDiagnosticos.EOF
'    With CboDiagnosticos
'        .AddItem RsDiagnosticos.Fields("Diagnotico").Value
'    End With
'    RsDiagnosticos.MoveNext
'Loop

CSql = "SELECT * From Tipos_Ca ORDER BY Descrip_Tipos_ca"
Set RsDiagnosticos = CrearRS(CSql)

Do While Not RsDiagnosticos.EOF
    With CboDiagnosticos
        .AddItem RsDiagnosticos.Fields("Descrip_Tipos_ca").Value
        .ItemData(.NewIndex) = RsDiagnosticos.Fields("Id_Tipos_ca").Value
    End With
    RsDiagnosticos.MoveNext
Loop

End Sub
Sub Estadiaje()
On Error Resume Next
CSql = "Select distinct (Estadiaje) From Estadisticas"
Set RsEstadiaje = CrearRS(CSql)

Do While Not RsEstadiaje.EOF
    With CboEstadiaje
        If Not IsNull(RsEstadiaje.Fields("Estadiaje").Value) Then
            .AddItem Trim(RsEstadiaje.Fields("Estadiaje").Value)
        End If
    End With
    RsEstadiaje.MoveNext
Loop

End Sub


Private Sub OptSesiones_Click(Index As Integer)
On Error Resume Next
If OptSesiones(0).Value Then
    TxtNoSesiones.Enabled = True
    TxtNoSesiones1.Enabled = False
    TxtNoSesiones2.Enabled = False
    TxtNoSesiones1.Text = "0"
    TxtNoSesiones2.Text = "0"
ElseIf OptSesiones(1).Value Then
    TxtNoSesiones.Enabled = False
    TxtNoSesiones1.Enabled = True
    TxtNoSesiones2.Enabled = True
    TxtNoSesiones1.Text = "1"
    TxtNoSesiones2.Text = "10"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If CDate(DTPicker1.Value) > CDate(DTPicker2.Value) Then
    MsgBox "La fecha de inicio no puede ser mayor a la fecha final, se realizarán ajustes!", vbExclamation + vbOKOnly, "Información"
    DTPicker1.Value = CDate(CDate(DTPicker2.Value) - 10)
End If
If OptEjeX(0).Enabled = False Then OptEjeX(0).Value = False
If OptEjeX(1).Enabled = False Then OptEjeX(1).Value = False
If OptEjeX(2).Enabled = False Then OptEjeX(2).Value = False
If OptEjeX(3).Enabled = False Then OptEjeX(3).Value = False
If OptEjeX(4).Enabled = False Then OptEjeX(4).Value = False
 
End Sub

Private Sub TxtEdad_Keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Not Trim(TxtEdad.Text) = "" Then
    If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then
            TxtEdad.Text = Val(TxtEdad.Text) + 1
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then
        TxtEdad.Text = Val(TxtEdad.Text) - 1
        If Val(TxtEdad.Text) <= 0 Then TxtEdad.Text = "0"
    End If
End If
End Sub

Private Sub TxtNoSesiones_Keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Not Trim(TxtNoSesiones.Text) = "" Then
    If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then
            TxtNoSesiones.Text = Val(TxtNoSesiones.Text) + 1
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then
        TxtNoSesiones.Text = Val(TxtNoSesiones.Text) - 1
        If Val(TxtNoSesiones.Text) <= 0 Then TxtNoSesiones.Text = "0"
    End If
End If
End Sub

Private Sub TxtNoSesiones1_Keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Not Trim(TxtNoSesiones1.Text) = "" Then
    If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then
            TxtNoSesiones1.Text = Val(TxtNoSesiones1.Text) + 1
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then
        TxtNoSesiones1.Text = Val(TxtNoSesiones1.Text) - 1
        If Val(TxtNoSesiones1.Text) <= 0 Then TxtNoSesiones1.Text = "0"
    End If
End If
End Sub

Private Sub TxtNoSesiones2_Keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Not Trim(TxtNoSesiones2.Text) = "" Then
    If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then
            TxtNoSesiones2.Text = Val(TxtNoSesiones2.Text) + 1
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then
        TxtNoSesiones2.Text = Val(TxtNoSesiones2.Text) - 1
        If Val(TxtNoSesiones2.Text) <= 0 Then TxtNoSesiones2.Text = "0"
    End If
End If
End Sub


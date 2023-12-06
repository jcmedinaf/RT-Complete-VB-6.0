VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmTablaEstadisticasAdmin 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadisticas Administrativas"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   13095
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   8400
      Width           =   12855
      Begin ChamaleonButton.ChameleonBtn BtnGraficar 
         Height          =   375
         Left            =   4800
         TabIndex        =   40
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
         MICON           =   "FrmEstadisticasAdmin.frx":0000
         PICN            =   "FrmEstadisticasAdmin.frx":001C
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
         Left            =   11640
         TabIndex        =   41
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
         MICON           =   "FrmEstadisticasAdmin.frx":0393
         PICN            =   "FrmEstadisticasAdmin.frx":03AF
         PICH            =   "FrmEstadisticasAdmin.frx":0578
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
         Left            =   10320
         TabIndex        =   42
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
         MICON           =   "FrmEstadisticasAdmin.frx":07AD
         PICN            =   "FrmEstadisticasAdmin.frx":07C9
         PICH            =   "FrmEstadisticasAdmin.frx":0AAB
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
         TabIndex        =   43
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
         MICON           =   "FrmEstadisticasAdmin.frx":0CFC
         PICN            =   "FrmEstadisticasAdmin.frx":0D18
         PICH            =   "FrmEstadisticasAdmin.frx":0E3D
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
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Index           =   2
      Left            =   120
      TabIndex        =   60
      Top             =   600
      Width           =   12855
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   2655
         Left            =   5280
         TabIndex        =   70
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4683
         Object.Width           =   6945
         Object.Height          =   2625
      End
      Begin VB.TextBox Text1 
         Height          =   4215
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   3480
         Width           =   4815
      End
      Begin VB.ListBox List2 
         Height          =   2205
         ItemData        =   "FrmEstadisticasAdmin.frx":10CD
         Left            =   2880
         List            =   "FrmEstadisticasAdmin.frx":10CF
         TabIndex        =   63
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "FrmEstadisticasAdmin.frx":10D1
         Left            =   240
         List            =   "FrmEstadisticasAdmin.frx":10D3
         TabIndex        =   61
         ToolTipText     =   "Nro de Barra  /  Valor de la barra"
         Top             =   600
         Width           =   2055
      End
      Begin ChamaleonButton.ChameleonBtn BtnResultados 
         Height          =   375
         Left            =   10320
         TabIndex        =   65
         Top             =   7320
         Width           =   2295
         _ExtentX        =   4048
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
         MICON           =   "FrmEstadisticasAdmin.frx":10D5
         PICN            =   "FrmEstadisticasAdmin.frx":10F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   375
         Left            =   2400
         TabIndex        =   66
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmEstadisticasAdmin.frx":1390
         PICN            =   "FrmEstadisticasAdmin.frx":13AC
         PICH            =   "FrmEstadisticasAdmin.frx":160B
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
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Promedio:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   71
         Top             =   3840
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Menor valor:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mayor valor:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   68
         Top             =   3120
         Width           =   870
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   120
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Datos para los calculos:"
         Height          =   195
         Left            =   3000
         TabIndex        =   64
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de barras:"
         Height          =   195
         Left            =   360
         TabIndex        =   62
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   12855
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Representación Gráfica"
         Height          =   7455
         Left            =   3120
         TabIndex        =   31
         Top             =   240
         Width           =   9615
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
            Height          =   375
            Left            =   8520
            TabIndex        =   47
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   6960
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "FrmEstadisticasAdmin.frx":186A
            PICN            =   "FrmEstadisticasAdmin.frx":1886
            PICH            =   "FrmEstadisticasAdmin.frx":1B1C
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
            Left            =   120
            TabIndex        =   46
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   6960
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "FrmEstadisticasAdmin.frx":1D7B
            PICN            =   "FrmEstadisticasAdmin.frx":1D97
            PICH            =   "FrmEstadisticasAdmin.frx":202C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   6735
            Left            =   120
            OleObjectBlob   =   "FrmEstadisticasAdmin.frx":2288
            TabIndex        =   32
            Top             =   240
            Width           =   9375
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00EAEFEF&
            Caption         =   "Barra  Nro.1   ...   Nro. 15"
            Height          =   255
            Left            =   2040
            TabIndex        =   53
            Top             =   7080
            Visible         =   0   'False
            Width           =   5295
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Tipo de Gráfica"
         Height          =   2655
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton Opt3DBar 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Barra 3D"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Opt2DBar 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Barrar 2D"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Opt3DLine 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Linea 3D"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Opt2DLine 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Linea 2D"
            Height          =   375
            Left            =   1200
            TabIndex        =   27
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Opt3DArea 
            BackColor       =   &H00EAEFEF&
            Caption         =   "3D Area"
            Height          =   255
            Left            =   1200
            TabIndex        =   26
            Top             =   420
            Width           =   1095
         End
         Begin VB.OptionButton Opt2DArea 
            BackColor       =   &H00EAEFEF&
            Caption         =   "2D Area"
            Height          =   375
            Left            =   1200
            TabIndex        =   25
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Opt3DStep 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Paso 3D"
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   2220
            Width           =   975
         End
         Begin VB.OptionButton Opt2DStep 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Paso 2D"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   975
         End
         Begin VB.OptionButton Opt2DCombination 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Combinacion 2D"
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   1860
            Width           =   1575
         End
         Begin VB.OptionButton Opt2DPie 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Torta 2D"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton Opt2DXY 
            BackColor       =   &H00EAEFEF&
            Caption         =   "2D X-Y"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1860
            Width           =   975
         End
         Begin VB.OptionButton Opt3DCombiantion 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Combinacion 3D"
            Height          =   375
            Left            =   1200
            TabIndex        =   22
            Top             =   1440
            Width           =   1575
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnAmpliarGrafico 
         Height          =   495
         Left            =   120
         TabIndex        =   17
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
         MICON           =   "FrmEstadisticasAdmin.frx":457E
         PICN            =   "FrmEstadisticasAdmin.frx":459A
         PICH            =   "FrmEstadisticasAdmin.frx":4911
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblValores 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Valor Máx:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   52
         Top             =   4680
         Width           =   750
      End
      Begin VB.Label LblValores 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Valor Mín:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   51
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label LblValores 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   50
         Top             =   5040
         Width           =   90
      End
      Begin VB.Label LblValores 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   49
         Top             =   4680
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Pto. Seleccionado:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   3360
         Width           =   1350
      End
      Begin VB.Label PtoSel 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "(                        )"
         Height          =   195
         Left            =   1680
         TabIndex        =   37
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Sumatoria:"
         Height          =   195
         Index           =   2
         Left            =   825
         TabIndex        =   36
         Top             =   3840
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Promedio:"
         Height          =   195
         Index           =   3
         Left            =   870
         TabIndex        =   35
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label TotReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "(                        )"
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label PrmReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "(                        )"
         Height          =   195
         Left            =   1680
         TabIndex        =   33
         Top             =   4080
         Width           =   1170
      End
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   12855
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Diagnosticos"
         Height          =   5775
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   10215
         Begin VB.Frame Frame1 
            BackColor       =   &H00EAEFEF&
            Height          =   520
            Left            =   1560
            TabIndex        =   54
            Top             =   2160
            Width           =   8535
            Begin VB.OptionButton Capital 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Ingresos"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   1440
               TabIndex        =   57
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton Capital 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Egresos"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   2760
               TabIndex        =   56
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton Capital 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Ambos"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   55
               Top             =   120
               Width           =   1215
            End
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Capital"
            Height          =   375
            Left            =   240
            TabIndex        =   58
            ToolTipText     =   "Movimientos en el módulo de Administración"
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "%  Seguros"
            Height          =   375
            Left            =   240
            TabIndex        =   48
            ToolTipText     =   "Muestra el porcentaje de pacientes por cada Seguro Médico"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox TxtQuery 
            Height          =   1215
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   5640
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Incluir rango de fechas:"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "Aplica solo para los ""tipos de pago"" y ""capital"""
            Top             =   300
            Width           =   2055
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00EAEFEF&
            Height          =   520
            Left            =   1560
            TabIndex        =   5
            Top             =   720
            Width           =   8535
            Begin VB.OptionButton Pago 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Ambos"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   1215
            End
            Begin VB.ComboBox Combo3 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "FrmEstadisticasAdmin.frx":4B46
               Left            =   4200
               List            =   "FrmEstadisticasAdmin.frx":4B54
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   150
               Width           =   2295
            End
            Begin VB.OptionButton Pago 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Por terceros"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   2760
               TabIndex        =   7
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton Pago 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Particular"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   1440
               TabIndex        =   6
               Top             =   120
               Width           =   1215
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   320
            Left            =   3360
            TabIndex        =   11
            Top             =   267
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   49283073
            CurrentDate     =   40190
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   320
            Left            =   6120
            TabIndex        =   12
            Top             =   267
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   49283073
            CurrentDate     =   40221
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "%  Medicos"
            Height          =   375
            Left            =   240
            TabIndex        =   45
            ToolTipText     =   "Muestra el porcentaje de pacientes remitidos por cada médico"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tipo de Pago"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            ToolTipText     =   "Muestra el número de pacientes de acuerdo al tipo de pago"
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin:"
            Height          =   195
            Left            =   5280
            TabIndex        =   14
            Top             =   330
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio:"
            Height          =   195
            Left            =   2400
            TabIndex        =   13
            Top             =   330
            Width           =   915
         End
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
         Left            =   11520
         TabIndex        =   15
         Top             =   7440
         Width           =   1110
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Resultados"
      Height          =   375
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grafica"
      Height          =   375
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "FrmTablaEstadisticasAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BuffMayorE As Integer
Dim BuffMenorE As Integer
Dim NumMay As Double
Dim NumMen As Double
Dim RsTemp As New ADODB.Recordset
Dim RsTemp2 As New ADODB.Recordset
Dim ArrayTemp()     ' Se usa poco... optimizar para eliminar esta variable!!!
Dim ArrayTemp2()    ' Se usa para almacenar TODOS los registros de una consulta estadistica...
Dim ArrayTemp3()    ' Se usa para almacenar SOLO 15 registros de una consulta estadistica...
Dim Pos1, Pos1buff As Integer   'Pos1 se usa para almacenar la posicion inicial actual del arreglo, POS1BUFF se usa para almacenar la antepenultima posicion inicial de los registros dentro del arreglo
Dim Pos2, Pos2buff As Integer   'Pos2 se usa para almacenar la posicion final actual del arreglo, POS1BUFF se usa para almacenar la antepenultima posicion final de los registros dentro del arreglo
Dim MasDe15 As Boolean  ' Variable que es verdadera SOLO cuando las estadisticas dan mas de 15 barras...
Dim TamB As Integer     ' Variable usada para almacena el tamaño de un arreglo, se usa en los botones ANTERIOR y SIGUIENTE
Dim TamBuff As Integer  ' Variable usada para almacena el tamaño de un arreglo, se usa en los botones ANTERIOR y SIGUIENTE

Sub Crear_Sentencia()
On Error Resume Next
Dim EQuery As String
Dim Cpdr, Cpdr2
Dim PlotMayor, PlotMenor As Double

MasDe15 = False
Label3.Visible = False
BtnAnterior.Visible = False
BtnSiguiente.Visible = False
' Si se condiciono de acuerdo a los Diagnosticos entonces...
Dim Senten_Sel As String

Senten_Sel = "SELECT Paciente.IdPaciente,C_Cobrar.Fecha FROM dbo.C_Cobrar" & _
  " INNER JOIN dbo.Paciente ON (dbo.C_Cobrar.IdPaciente = dbo.Paciente.IdPaciente) " & _
  " INNER JOIN dbo.Cliente ON (dbo.C_Cobrar.IdCliente = dbo.Cliente.IdCliente)"
EQuery = ""

If Check1.Value Then
    If InStr(1, EQuery, "WHERE") = 0 Then
        EQuery = Senten_Sel & " WHERE Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Fecha <='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
    Else
        EQuery = EQuery & " AND Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Fecha <='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
    End If
End If

                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                
If Check3.Value Then

    Dim TotalPac As Integer
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Consulta para saber cuantos medicos remitentes existen, y los almacena en "ArrayTemp"
        CSql = "SELECT  COUNT(Paciente.IdPaciente) AS NumReg,  Medicos.IdMedico,  (Medicos.Nombre + ' ' + Medicos.Apellido) AS NombreMedR " & _
                "FROM Paciente RIGHT OUTER JOIN Medicos ON (Medicos.IdMedico = Paciente.Medico_Remitente) " & _
                "WHERE Paciente.Medico_Remitente IN (SELECT Medicos.IdMedico FROM Medicos WHERE Medicos.Tipo = 1 OR Medicos.Tipo = 3) " & _
                "GROUP BY Medicos.IdMedico, Medicos.Nombre, Medicos.Apellido ORDER BY Medicos.IdMedico "
        Set RsTemp = CrearRS(CSql)
        
        MayorE = RsTemp.RecordCount
        If IsEmpty(MayorE) Then
            MsgBox "No hay valores que mostrar!", vbExclamation + vbOKOnly, "Información"
            Exit Sub
        End If
        ReDim ArrayTemp(1 To MayorE, 0 To 1)
        ReDim ArrayTemp2(1 To MayorE, 0 To 1)
        ReDim ArrayTemp3(1 To 15, 0 To 1)
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        If MayorE > 14 Then
            MasDe15 = True
            Label3.Visible = True
            Label3.Caption = "Barra  Nro.1   ...   Nro.15"
            BtnAnterior.Visible = True
            BtnSiguiente.Visible = True
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Sumatoria para saber el numero de pacientes...
        i = 0
        While Not RsTemp.EOF
            TotalPac = TotalPac + Val(RsTemp.Fields("NumReg").Value)
            RsTemp.MoveNext
        Wend
        
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        i = 0
        RsTemp.MoveFirst
        
        PlotMayor = 0
        PlotMenor = -1
        
        LblValores(0).Caption = PlotMayor
        LblValores(0).ToolTipText = RsTemp.Fields("NombreMedR").Value
        LblValores(1).Caption = PlotMenor
        LblValores(1).ToolTipText = RsTemp.Fields("NombreMedR").Value
                
        While Not RsTemp.EOF
            
            i = i + 1
            ArrayTemp2(i, 0) = RsTemp.Fields("NombreMedR").Value ' Almacena el nombre para la barra
            
            If i <= 15 Then ArrayTemp3(i, 0) = ArrayTemp2(i, 0)
            
            If (RsTemp.Fields("NumReg").Value) <> 0 Then
                ArrayTemp2(i, 1) = (RsTemp.Fields("NumReg").Value) * 100 / TotalPac ' Almacena el PORCENTAJE DE PACIENTES para el registro consultado.
                If i <= 15 Then ArrayTemp3(i, 1) = (ArrayTemp2(i, 1))
            Else
                ArrayTemp2(i, 1) = 0
                If i <= 15 Then ArrayTemp3(i, 1) = 0
            End If
            
            If CDbl(ArrayTemp2(i, 1)) >= PlotMayor Then
                PlotMayor = CDbl(ArrayTemp2(i, 1))
                LblValores(0).ToolTipText = RsTemp.Fields("NombreMedR").Value & "   Nro. de barra:" & i & " pacientes:" & RsTemp.Fields("NumReg").Value
                If PlotMenor = -1 Then PlotMenor = PlotMayor
            End If
            If CDbl(ArrayTemp2(i, 1)) <= PlotMenor Then
                PlotMenor = CDbl(ArrayTemp2(i, 1))
                LblValores(1).ToolTipText = RsTemp.Fields("NombreMedR").Value & "   Nro. de barra:" & i & " pacientes:" & RsTemp.Fields("NumReg").Value
            End If
            
            RsTemp.MoveNext
        Wend
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        LblValores(0).Caption = PlotMayor
        LblValores(1).Caption = PlotMenor
        
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Configura el arreglo 3 para la ampliacion del grafico si es el caso...
        For J = 1 To 15
            If MayorE >= J Then
                ArrayTemp3(J, 1) = ArrayTemp2(J, 1)
            Else
                ArrayTemp3(J, 0) = "..."
            End If
        Next
        
        If MayorE <= 15 Then
            MSChart1.ChartData = ArrayTemp2
            MSChart2.ChartData = ArrayTemp2
        Else
            Pos1 = 1
            Pos2 = 15
            MSChart1.ChartData = ArrayTemp3
            MSChart2.ChartData = ArrayTemp3
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      
        If MayorE <= 15 Then
            MSChart1.ChartData = ArrayTemp2
            MSChart2.ChartData = ArrayTemp2
        Else
            Pos1 = 1
            Pos2 = 15
            MSChart1.ChartData = ArrayTemp3
            MSChart2.ChartData = ArrayTemp3
        End If
        MSChart1.Refresh
        MSChart2.Refresh
        
        MSChart1.Plot.SeriesCollection(1).LegendText = "% de pacientes"
        MSChart1.TitleText = "Porcentaje de pacientes por Médico Remitente"
End If

                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

If Check4.Value Then
        Dim Sumbuff As Double

      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        CSql = "SELECT COUNT(C_Cobrar.IdCliente) AS NumReg,  C_Cobrar.IdCliente,  Cliente.Razon,  Cliente.Personal " & _
                "FROM   C_Cobrar  RIGHT OUTER JOIN dbo.Cliente ON (C_Cobrar.IdCliente = dbo.Cliente.IdCliente) " & _
                "WHERE  C_Cobrar.Tipo = 1 AND  C_Cobrar.Anulada <> 1 AND " & _
                "C_Cobrar.IdCliente IN (SELECT Cliente.IdCliente FROM Cliente WHERE Personal = 2) OR Cliente.Personal = 2 " & _
                "Group By  C_Cobrar.IdCliente,  dbo.Cliente.Razon,  dbo.Cliente.Personal Order By  C_Cobrar.IdCliente"
        Set RsTemp = CrearRS(CSql)
        
        MayorE = RsTemp.RecordCount
        If IsEmpty(MayorE) Then
            MsgBox "No hay valores que mostrar!", vbExclamation + vbOKOnly, "Información"
            Exit Sub
        End If
        ReDim ArrayTemp(1 To MayorE, 0 To 1)
        ReDim ArrayTemp2(1 To MayorE, 0 To 1)
        ReDim ArrayTemp3(1 To 15, 0 To 1)
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        If MayorE > 14 Then
            MasDe15 = True
            Label3.Visible = True
            Label3.Caption = "Barra  Nro.1   ...   Nro.15"
            BtnAnterior.Visible = True
            BtnSiguiente.Visible = True
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Consulta el numero de pacientes para un seguro especifico...
        i = 0
        While Not RsTemp.EOF
        
            MsgBox RsTemp
            Sumbuff = Sumbuff + Val(RsTemp.Fields("NumReg").Value)
            RsTemp.MoveNext
        Wend
        
        RsTemp.MoveFirst
        PlotMayor = 0
        PlotMenor = -1
        
        LblValores(0).Caption = PlotMayor
        LblValores(0).ToolTipText = RsTemp.Fields("Razon").Value
        LblValores(1).Caption = PlotMenor
        LblValores(1).ToolTipText = RsTemp.Fields("Razon").Value
        
        While Not RsTemp.EOF
            
            i = i + 1
            ArrayTemp2(i, 0) = RsTemp.Fields("Razon").Value ' Almacena el nombre para la barra
            
            If i <= 15 Then ArrayTemp3(i, 0) = ArrayTemp2(i, 0)
            
            If (RsTemp.Fields("NumReg").Value) <> 0 Then
                ArrayTemp2(i, 1) = (RsTemp.Fields("NumReg").Value) * 100 / Sumbuff ' Almacena el PORCENTAJE DE PACIENTES para el registro consultado.
                If i <= 15 Then ArrayTemp3(i, 1) = (ArrayTemp2(i, 1))
            Else
                ArrayTemp2(i, 1) = 0
                If i <= 15 Then ArrayTemp3(i, 1) = 0
            End If
            
            If CDbl(ArrayTemp2(i, 1)) >= PlotMayor Then
                PlotMayor = CDbl(ArrayTemp2(i, 1))
                LblValores(0).ToolTipText = RsTemp.Fields("Razon").Value & "   Nro. de barra:" & i & " pacientes:" & RsTemp.Fields("NumReg").Value
                
                If PlotMenor = -1 Then PlotMenor = PlotMayor
            End If
            If CDbl(ArrayTemp2(i, 1)) <= PlotMenor Then
                PlotMenor = CDbl(ArrayTemp2(i, 1))
                LblValores(1).ToolTipText = RsTemp.Fields("Razon").Value & "   Nro. de barra:" & i & " pacientes:" & RsTemp.Fields("NumReg").Value
            End If
            
            RsTemp.MoveNext
        Wend
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        LblValores(0).Caption = PlotMayor
        LblValores(1).Caption = PlotMenor
        
        For J = 1 To 15
            If MayorE >= J Then
                ArrayTemp3(J, 1) = ArrayTemp2(J, 1)
            Else
                ArrayTemp3(J, 0) = "..."
            End If
        Next
        
        If MayorE <= 15 Then
            MSChart1.ChartData = ArrayTemp2
            MSChart2.ChartData = ArrayTemp2
        Else
            Pos1 = 1
            Pos2 = 15
            MSChart1.ChartData = ArrayTemp3
            MSChart2.ChartData = ArrayTemp3
        End If
        MSChart1.Refresh
        MSChart2.Refresh
        
        MSChart1.Plot.SeriesCollection(1).LegendText = "% de pacientes"
        MSChart1.TitleText = "Porcentaje de pacientes por Seguro"
        'MSChart1.Plot.SeriesCollection(0).LegendText = "0"
        'MSChart1.Plot.SeriesCollection(2).LegendText = "2"
End If

If Check5.Value Then

    Dim TotalIng As Double
    Dim TotalEgr As Double
        
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Prepara la consulta para clasificarla por fecha si fuera ese el caso....
        If Check1.Value Then
            EQuery = " AND Fecha_Transa >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Fecha_Transa <='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
        Else
            EQuery = ""
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Realiza la consulta para traerse todos los registros de Ingreso y Egresos
        CSql = "SELECT ((CASE Ingr_Egr WHEN (1) THEN Monto_Mov ELSE NULL END)) AS NumRegIng, " & _
                "((CASE Ingr_Egr WHEN (2) THEN Monto_Mov ELSE NULL END)) AS NumRegEgr, " & _
                "Fecha_Transa FROM Movi_BanCaja WHERE Anulado = 0 " & EQuery & " ORDER BY IdMovCajaBanco"
        Set RsTemp = CrearRS(CSql)
       
        If IsEmpty(RsTemp.RecordCount) Then
            MsgBox "No hay valores que mostrar!", vbExclamation + vbOKOnly, "Información"
            Exit Sub
        End If
      ' Ciclo para obtener la sumatoria de todos los ingresos y egresos
        While Not RsTemp.EOF
            If Not IsNull(RsTemp.Fields("NumRegIng").Value) Then TotalIng = TotalIng + CDbl(RsTemp.Fields("NumRegIng").Value)
            If Not IsNull(RsTemp.Fields("NumRegEgr").Value) Then TotalEgr = TotalEgr + CDbl(RsTemp.Fields("NumRegEgr").Value)
            RsTemp.MoveNext
        Wend
      ' ---------------------------------------------------------------
        RsTemp.MoveFirst
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      
        PlotMayor = 0
        PlotMenor = -1
        
        LblValores(0).Caption = ""
        LblValores(0).ToolTipText = ""
        LblValores(1).Caption = ""
        LblValores(1).ToolTipText = ""
        
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Prepara la configuración para clasificar dentro de un arreglo los resultados por fecha
        MenorE = Format(DTPicker1.Value, "MM")
        MayorE = Val(Format(CDate(CDate(Format(DTPicker2.Value, "MM/yyyy")) - CDate(Format(DTPicker1.Value, "MM/yyyy"))), "MM"))
        MayorE = MayorE + (CDate(Format(DTPicker2.Value, "yyyy")) - CDate(Format(DTPicker1.Value, "yyyy"))) * 12
        
        ReDim ArrayTemp(1 To MayorE, 0 To 2)    ' Prepara el arreglo a un tamaño determinado
        ReDim ArrayTemp2(1 To MayorE, 0 To 2)   ' Prepara el arreglo a un tamaño determinado
        ReDim ArrayTemp3(1 To 15, 0 To 2)
        
        Fecha = Format(DTPicker1.Value, "dd/MM/yyyy")
        For J = 1 To MayorE
            If CDate(Format(Fecha, "MM/yyyy")) > CDate(Format(DTPicker2.Value, "MM/yyyy")) Then MayorE = J - 1: Exit For
            Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
        Next
        
        Fecha = Format(DTPicker1.Value, "dd/MM/yyyy")
        For J = 1 To MayorE
            ArrayTemp2(J, 0) = Format(CDate(Format(Fecha, "MM/yyyy")), "MM/yyyy")
            If J <= 15 Then ArrayTemp3(J, 0) = ArrayTemp2(J, 0)
            Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
        Next
        
      ' Ciclo para clasificar los registros, los ingresos y le egresos...
        While Not RsTemp.EOF
            For i = 1 To MayorE
                If ArrayTemp2(i, 0) = Format(RsTemp.Fields("Fecha_Transa").Value, "MM/yyyy") Then
                  ' Condicional que verifica se el campo "NumRegIng" es nulo, de NO ser asi se cumple la condicion
                    If Not IsNull(RsTemp.Fields("NumRegIng").Value) Then
                      ' Acumula el valor del ingreso...
                        If Not Capital(2).Value Then ArrayTemp2(i, 1) = CDbl(ArrayTemp2(i, 1)) + CDbl(RsTemp.Fields("NumRegIng").Value)
                    Else
                      ' Acumula el valor del egreso...
                        If Not Capital(1).Value Then ArrayTemp2(i, 2) = CDbl(ArrayTemp2(i, 2)) + CDbl(RsTemp.Fields("NumRegEgr").Value)
                    End If
                    Exit For
                End If
            Next i
            RsTemp.MoveNext
        Wend
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        If MayorE > 14 Then
            MasDe15 = True
            Label3.Visible = True
            Label3.Caption = "Barra  Nro.1   ...   Nro.15"
            BtnAnterior.Visible = True
            BtnSiguiente.Visible = True
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

        For J = 1 To 15
            If MayorE >= J Then
                ArrayTemp3(J, 1) = ArrayTemp2(J, 1)
                ArrayTemp3(J, 2) = ArrayTemp2(J, 2)
            Else
                ArrayTemp3(J, 0) = "..."
            End If
        Next
        
        If MayorE <= 15 Then
            MSChart1.ChartData = ArrayTemp2
            MSChart2.ChartData = ArrayTemp2
        Else
            Pos1 = 1
            Pos2 = 15
            MSChart1.ChartData = ArrayTemp3
            MSChart2.ChartData = ArrayTemp3
        End If
        
        If Capital(2).Value Then
            MSChart1.Plot.SeriesCollection(1).LegendText = ""
            MSChart1.Plot.SeriesCollection(2).LegendText = "Egresos "
            MSChart1.TitleText = "Movimiento de Bancos: Egresos"
        ElseIf Capital(0).Value Then
            MSChart1.Plot.SeriesCollection(1).LegendText = "Ingresos"
            MSChart1.Plot.SeriesCollection(2).LegendText = "Egresos "
            MSChart1.TitleText = "Movimiento de Bancos: Ingresos vs Egresos"
        ElseIf Capital(1).Value Then
            MSChart1.Plot.SeriesCollection(1).LegendText = "Ingresos"
            MSChart1.Plot.SeriesCollection(2).LegendText = ""
            MSChart1.TitleText = "Movimiento de Bancos: Ingresos"
        End If
End If

If Check2.Value Then

      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Prepara la consulta para clasificarla por fecha si fuera ese el caso....
        If Check1.Value Then
            EQuery = " AND Fecha_Transa >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Fecha_Transa <='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
        Else
            EQuery = ""
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Realiza la consulta para traerse todos los registros de Ingreso y Egresos
      
        If Pago(1).Value And Combo3.ListIndex <> 0 And Combo3.ListIndex <> -1 Then
            CSql = "SELECT ((CASE Cliente.Personal WHEN (0) THEN C_Cobrar.IdPaciente ELSE NULL END)) AS PPersonal, " & _
                "((CASE Cliente.Personal WHEN (" & Combo3.ItemData(Combo3.ListIndex) & ") THEN C_Cobrar.IdPaciente ELSE NULL END)) AS OtroP, " & _
                "Fecha FROM dbo.C_Cobrar INNER JOIN dbo.Cliente ON (dbo.C_Cobrar.IdCliente = dbo.Cliente.IdCliente) Where C_Cobrar.Anulada = 0"
        Else
            CSql = "SELECT ((CASE Cliente.Personal WHEN (0) THEN C_Cobrar.IdPaciente ELSE NULL END)) AS PPersonal, " & _
                "((CASE Cliente.Personal WHEN 1 THEN C_Cobrar.IdPaciente WHEN 2 THEN C_Cobrar.IdPaciente ELSE NULL END)) AS OtroP, " & _
                "Fecha FROM dbo.C_Cobrar INNER JOIN dbo.Cliente ON (dbo.C_Cobrar.IdCliente = dbo.Cliente.IdCliente) Where C_Cobrar.Anulada = 0"
        End If
        
        Set RsTemp = CrearRS(CSql)
       
        If IsEmpty(RsTemp.RecordCount) Then
            MsgBox "No hay valores que mostrar!", vbExclamation + vbOKOnly, "Información"
            Exit Sub
        End If
      ' Ciclo para obtener la sumatoria de todos los ingresos y egresos
        While Not RsTemp.EOF
            If Not IsNull(RsTemp.Fields("PPersonal").Value) Then TotalIng = TotalIng + CDbl(RsTemp.Fields("PPersonal").Value)
            If Not IsNull(RsTemp.Fields("OtroP").Value) Then TotalEgr = TotalEgr + CDbl(RsTemp.Fields("OtroP").Value)
            RsTemp.MoveNext
        Wend
      ' ---------------------------------------------------------------
        RsTemp.MoveFirst
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      
        PlotMayor = 0
        PlotMenor = -1
        
        LblValores(0).Caption = ""
        LblValores(0).ToolTipText = ""
        LblValores(1).Caption = ""
        LblValores(1).ToolTipText = ""
        
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' Prepara la configuración para clasificar dentro de un arreglo los resultados por fecha
        MenorE = Format(DTPicker1.Value, "MM")
        MayorE = Val(Format(CDate(CDate(Format(DTPicker2.Value, "MM/yyyy")) - CDate(Format(DTPicker1.Value, "MM/yyyy"))), "MM"))
        MayorE = MayorE + (CDate(Format(DTPicker2.Value, "yyyy")) - CDate(Format(DTPicker1.Value, "yyyy"))) * 12
        
        ReDim ArrayTemp(1 To MayorE, 0 To 2)    ' Prepara el arreglo a un tamaño determinado
        ReDim ArrayTemp2(1 To MayorE, 0 To 2)   ' Prepara el arreglo a un tamaño determinado
        ReDim ArrayTemp3(1 To 15, 0 To 2)
        
        Fecha = Format(DTPicker1.Value, "dd/MM/yyyy")
        For J = 1 To MayorE
            If CDate(Format(Fecha, "MM/yyyy")) > CDate(Format(DTPicker2.Value, "MM/yyyy")) Then MayorE = J - 1: Exit For
            Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
        Next
        
        Fecha = Format(DTPicker1.Value, "dd/MM/yyyy")
        For J = 1 To MayorE
            ArrayTemp2(J, 0) = Format(CDate(Format(Fecha, "MM/yyyy")), "MM/yyyy")
            If J <= 15 Then ArrayTemp3(J, 0) = ArrayTemp2(J, 0)
            Fecha = CDate(Format(Fecha, "dd/MM/yyyy")) + Val(Format(DateSerial(Year(CDate(Fecha)), Month(CDate(Fecha)) + 1, 0), "dd"))
        Next
        
      ' Ciclo para clasificar los registros, "Pagos Personales" y "Otros Pagos"
        While Not RsTemp.EOF
            For i = 1 To MayorE
                If ArrayTemp2(i, 0) = Format(RsTemp.Fields("Fecha").Value, "MM/yyyy") Then
                  ' Condicional que verifica se el campo "NumRegIng" es nulo, de NO ser asi se cumple la condicion
                    If Not IsNull(RsTemp.Fields("PPersonal").Value) Then
                      ' Suma 1 al contador para "Pagos Personales" ...
                        If Not Pago(1).Value Then ArrayTemp2(i, 1) = Val(ArrayTemp2(i, 1)) + 1 Else ArrayTemp2(i, 1) = 0
                    ElseIf Not IsNull(RsTemp.Fields("OtroP").Value) Then
                      ' Suma 1 al contador para "Otros Pagos" ...
                        If Not Pago(0).Value Then ArrayTemp2(i, 2) = Val(ArrayTemp2(i, 2)) + 1 Else ArrayTemp2(i, 2) = 0
                    End If
                    Exit For
                End If
            Next i
            RsTemp.MoveNext
        Wend
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        If MayorE > 14 Then
            MasDe15 = True
            Label3.Visible = True
            Label3.Caption = "Barra  Nro.1   ...   Nro.15"
            BtnAnterior.Visible = True
            BtnSiguiente.Visible = True
        End If
      ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      
        J = 0
        For i = 1 To 15
            J = J + 1
            
            If MayorE >= i Then
                If Not IsEmpty(ArrayTemp2(i, 0)) Then ArrayTemp3(J, 0) = ArrayTemp2(i, 0) Else ArrayTemp3(J, 0) = ""
                If Not IsEmpty(ArrayTemp2(i, 1)) Then ArrayTemp3(J, 1) = ArrayTemp2(i, 1) Else ArrayTemp3(J, 1) = 0
                If Not IsEmpty(ArrayTemp2(i, 2)) Then ArrayTemp3(J, 2) = ArrayTemp2(i, 2) Else ArrayTemp3(J, 2) = 0
            Else
                ArrayTemp3(J, 0) = "..."
                ArrayTemp3(J, 1) = Empty
                ArrayTemp3(J, 2) = Empty
            End If
        Next i
        
        Pos1 = 1
        Pos2 = 15
        MSChart1.ChartData = ArrayTemp3
        MSChart2.ChartData = ArrayTemp3
        
        If Pago(2).Value Then
            MSChart1.Plot.SeriesCollection(1).LegendText = "Pagos Personales"
            MSChart1.Plot.SeriesCollection(2).LegendText = "Otros Pagos"
            MSChart1.TitleText = "Pagos Particulares  vs  Otros Pagos"
        ElseIf Pago(0).Value Then
            MSChart1.Plot.SeriesCollection(1).LegendText = "Pagos Personales"
            MSChart1.TitleText = "Número de Pacientes con pago Personal"
        ElseIf Pago(1).Value Then
            MSChart1.Plot.SeriesCollection(1).LegendText = "N/A"
            MSChart1.Plot.SeriesCollection(2).LegendText = "Pagos: " & Combo3.List(Combo3.ListIndex)
            MSChart1.TitleText = "Número de Pacientes con pagos de tipo: " & Combo3.List(Combo3.ListIndex)
        End If
End If

MSChart1.Refresh
MSChart2.Refresh

Option1(1).Value = True
                
                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                
TxtQuery.Text = EQuery & " GROUP BY Paciente.IdPaciente,C_Cobrar.Fecha "

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

If IsNull(ArrayTemp3(1, 0)) Then
    MsgBox "No hay hay datos!", vbExclamation + vbOKCancel, "Información"
    Exit Sub
End If

Dim TG As Integer
Dim TGBand As Boolean

TG = UBound(ArrayTemp3, 1)
TGBand = False

For iii = 1 To TG
    If Not IsEmpty(ArrayTemp3(iii, 1)) Then
        TGBand = True
    End If
Next

If TGBand = False Then Exit Sub

FrmAmpliarGrafico.MSChart1.ChartData = ArrayTemp3
FrmAmpliarGrafico.MSChart1.ChartType = MSChart1.ChartType

FrmAmpliarGrafico.MSChart1.Plot.SeriesCollection(1).LegendText = MSChart1.Plot.SeriesCollection(1).LegendText
FrmAmpliarGrafico.MSChart1.Plot.SeriesCollection(2).LegendText = MSChart1.Plot.SeriesCollection(2).LegendText
FrmAmpliarGrafico.MSChart1.TitleText = MSChart1.TitleText

FrmAmpliarGrafico.MSChart1.Refresh
FrmAmpliarGrafico.Show vbModal
End Sub

Private Sub BtnResultados_Click()
Dim NumEst As Integer
Dim Promedio As Double

BuffMenorE = LBound(ArrayTemp2, 1)
BuffMayorE = UBound(ArrayTemp2, 1)

List1.Clear
DMGrid1.Clear
DMGrid1.Rows = 0
For i = BuffMenorE To BuffMayorE
    
    If IsNull(ArrayTemp2(i, 0)) Then
        List1.AddItem ""
    Else
        List1.AddItem ArrayTemp2(i, 0)
    End If
Next

' Configura el arreglo(M, N) en donde
' M es: el número de barra
' N va de 0 a 3
'   la posicion 0 es la FRECUENCIA para M    (fi)
'   la posicion 1 es la FRECUENCIA ACUMULADA (Fi)
'   la posicion 2 es la FRECUENCIA RELATIVA  (Fr=fi/N)
'   la posicion 3 es la FRECUENCIA RELATIVA ACUMULADA FR

NumEst = List2.ListCount

If NumEst = 0 Then Exit Sub
ReDim ArrayTemp(1 To NumEst, 0 To 3)
J = 0
Text1.Text = ""

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Almacena la Frecuencia (fi)
For i = 1 To BuffMayorE
    If List2.List(J) = ArrayTemp2(i, 0) Then
        J = J + 1
        ArrayTemp(J, 0) = ArrayTemp2(i, 1)
        If NumMen = -1 Then NumMen = NumMay
        
        If Not IsEmpty(ArrayTemp(J, 0)) Then
            If ArrayTemp(J, 0) > NumMay Then
                NumMay = ArrayTemp(i, 0)
            End If
        End If
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(J, 1) = ArrayTemp2(i, 0)
        DMGrid1.ValorCelda(J, 2) = ArrayTemp(J, 0)
        Text1.Text = Text1.Text & Chr(13) & Chr(10) & ArrayTemp(J, 0)
    End If
Next
Text1.Text = Text1.Text & Chr(13) & Chr(10) & "================================="

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Almacena la Frecuencia Aculumada (Fi)
For i = 1 To NumEst

    If i <= 1 Then
        ArrayTemp(i, 1) = ArrayTemp(i, 0)
    Else
        ArrayTemp(i, 1) = ArrayTemp(i - 1, 1) + ArrayTemp(i, 0)
    End If
    DMGrid1.ValorCelda(i, 3) = ArrayTemp(i, 1)
    
    If i = NumEst Then Promedio = CDbl(DMGrid1.ValorCelda(i, 3)) / NumEst
    
    Text1.Text = Text1.Text & Chr(13) & Chr(10) & ArrayTemp(i, 1)
Next
Text1.Text = Text1.Text & Chr(13) & Chr(10) & "================================="

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Almacena la Frecuencia Relativa (Fr)
For i = 1 To NumEst
    ArrayTemp(i, 2) = ArrayTemp(i, 0) / NumEst
    DMGrid1.ValorCelda(i, 4) = ArrayTemp(i, 2)
    Text1.Text = Text1.Text & Chr(13) & Chr(10) & ArrayTemp(i, 2)
Next
Text1.Text = Text1.Text & Chr(13) & Chr(10) & "================================="

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Almacena la Frecuencia Relativa Acumulada (FR)
For i = 1 To NumEst
    If i <= 1 Then
        ArrayTemp(i, 3) = ArrayTemp(i, 2)
    Else
        ArrayTemp(i, 3) = ArrayTemp(i - 1, 3) + ArrayTemp(i, 2)
    End If
    DMGrid1.ValorCelda(i, 5) = ArrayTemp(i, 3)
Text1.Text = Text1.Text & Chr(13) & Chr(10) & ArrayTemp(i, 3)
Next
Text1.Text = Text1.Text & Chr(13) & Chr(10) & "================================="

Label6(0).Caption = "Mayor valor: " & NumMay
Label6(1).Caption = "Menor valor: " & NumMen
Label6(2).Caption = "Promedio: " & Promedio
End Sub

Private Sub BtnAnterior_Click()
Dim Div2 As Integer
Dim J As Integer

BuffMayorE = UBound(ArrayTemp2, 1)
BuffMenorE = LBound(ArrayTemp2, 1)

If Pos1 >= 15 Then
    If Pos2 = BuffMayorE Then
        Pos1 = Pos1buff
        Pos2 = Pos2buff
    Else
        Pos1 = Pos1 - 15
        Pos2 = Pos2 - 15
    End If
Else
    Pos1 = 1
    Pos2 = 15
    MsgBox "Ha llegado al inicio de la grafica!", vbExclamation + vbOKOnly, "Información"
End If

J = 0
Label3.Caption = "Barra  Nro." & Pos1 & "   ...   Nro." & Pos2
For i = Pos1 To Pos2
    J = J + 1
    If i <= BuffMayorE Then
        ArrayTemp3(J, 0) = ArrayTemp2(i, 0)
        ArrayTemp3(J, 1) = ArrayTemp2(i, 1)
        
        TamB = UBound(ArrayTemp3, 2)
        If TamB > 1 Then
            For TamBuff = 2 To TamB
                ArrayTemp3(J, TamBuff) = ArrayTemp2(i, TamBuff)
            Next
        End If
    Else
        Exit For
    End If
Next i

MSChart1.ChartData = ArrayTemp3
MSChart1.Refresh

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnSiguiente_Click()
Dim BuffMayorE As Integer
Dim BuffMenorE As Integer
Dim Div2 As Integer
Dim J As Integer

BuffMayorE = UBound(ArrayTemp2, 1)
BuffMenorE = LBound(ArrayTemp2, 1)

If (Pos2 + 15) <= BuffMayorE Then
    Pos1 = Pos1 + 15
    Pos2 = Pos2 + 15
Else
    If Pos2 <> BuffMayorE Then
        Pos1buff = Pos1
        Pos2buff = Pos2
    Else
        MsgBox "Ha llegado al final de la grafica!", vbExclamation + vbOKOnly, "Información"
        Exit Sub
    End If
    Pos1 = BuffMayorE - 14
    Pos2 = BuffMayorE
End If

J = 0
Label3.Caption = "Barra  Nro." & Pos1 & "   ...   Nro." & Pos2
For i = Pos1 To Pos2
    J = J + 1
    If i <= BuffMayorE Then
        ArrayTemp3(J, 0) = ArrayTemp2(i, 0)
        ArrayTemp3(J, 1) = ArrayTemp2(i, 1)
        
        TamB = UBound(ArrayTemp3, 2)
        If TamB > 1 Then
            For TamBuff = 2 To TamB
                ArrayTemp3(J, TamBuff) = ArrayTemp2(i, TamBuff)
            Next
        End If
    Else
        Exit For
    End If
Next i

MSChart1.ChartData = ArrayTemp3
MSChart1.Refresh

End Sub

Private Sub BtnGraficar_Click()
On Error Resume Next
Dim RsGraficar As New ADODB.Recordset
Dim ContTmp(0 To 50) As Integer
Dim i As Integer
Dim EjeX As Boolean
verifica_sent = True


If Check2.Value And Pago(1).Value And Combo3.ListIndex = -1 Then
    MsgBox "Debe seleccionar un tipo de pago!", vbInformation + vbOKOnly, "Faltan datos"
    Exit Sub
End If

If verifica_sent Then Crear_Sentencia

Exit Sub

End Sub

Private Sub Check1_Click()
On Error Resume Next
DTPicker1.Enabled = Check1.Value
DTPicker2.Enabled = Check1.Value
If Check1.Value = 1 Then
    Check3.Value = 0
    Check4.Value = 0
End If
End Sub

Private Sub Check2_Click()
On Error Resume Next
Pago(0).Enabled = Check2.Value
Pago(1).Enabled = Check2.Value
Pago(2).Enabled = Check2.Value
If Check2.Value = 1 Then
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
End If

If Check2.Value Then
    Pago(0).Value = True
Else
    Pago(0).Value = False
    Pago(1).Value = False
    Pago(2).Value = False
End If
End Sub

Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 1 Then
    Check2.Value = 0
    Check1.Value = 0
    Check4.Value = 0
Else
End If
End Sub

Private Sub Check4_Click()
On Error Resume Next
If Check4.Value = 1 Then
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check5.Value = 0
End If
End Sub

Private Sub Check5_Click()
On Error Resume Next
Capital(0).Enabled = Check5.Value
Capital(1).Enabled = Check5.Value
Capital(2).Enabled = Check5.Value

If Check5.Value Then
    Capital(0).Value = True
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
Else
    Capital(0).Value = False
    Capital(1).Value = False
    Capital(2).Value = False
End If

End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid
End Sub

Private Sub List1_DblClick()
If List1.ListIndex = -1 Then
    MsgBox "Seleccione un valor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
End If
List2.AddItem List1.List(List1.ListIndex)
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub List2_DblClick()
If List2.ListIndex = -1 Then
    MsgBox "Seleccione un valor!", vbInformation + vbOKOnly, "Información"
    Exit Sub
End If
List1.AddItem List2.List(List2.ListIndex)
List2.RemoveItem (List2.ListIndex)

End Sub

Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
On Error Resume Next
MSChart1.Column = Series
MSChart1.Row = DataPoint
PtoSel.Caption = MSChart1.Data
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

Private Sub Pago_Click(Index As Integer)
On Error Resume Next
If Pago(0).Value Or Pago(2).Value Then
    Combo3.ListIndex = -1
    Combo3.Enabled = False
ElseIf Pago(1).Value Then
    Combo3.ListIndex = 0
    Combo3.Enabled = True
End If
End Sub


Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 5
DMGrid1.Rows = 1
'DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 1
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1

'DMGrid1.DColumnas(1).IsNumber = True
DMGrid1.DColumnas(2).IsNumber = True
DMGrid1.DColumnas(3).IsNumber = True
DMGrid1.DColumnas(4).IsNumber = True
DMGrid1.DColumnas(5).IsNumber = True

'DMGrid1.DColumnas(7).Width = 1100

'   la posicion 0 es la FRECUENCIA para M    (fi)
'   la posicion 1 es la FRECUENCIA ACUMULADA (Fi)
'   la posicion 2 es la FRECUENCIA RELATIVA  (Fr=fi/N)
'   la posicion 3 es la FRECUENCIA RELATIVA ACUMULADA FR

DMGrid1.DColumnas(1).Caption = "X"
DMGrid1.DColumnas(2).Caption = "fi"
DMGrid1.DColumnas(3).Caption = "Fi"
DMGrid1.DColumnas(4).Caption = "fr"
DMGrid1.DColumnas(5).Caption = "Fri"

End Sub

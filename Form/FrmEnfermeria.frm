VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEnfermeria 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enfermería"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   Icon            =   "FrmEnfermeria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   12225
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Height          =   7095
      Left            =   120
      TabIndex        =   39
      Top             =   2880
      Width           =   12015
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Consumo de Medicamentos"
         Height          =   3375
         Left            =   3120
         TabIndex        =   60
         Top             =   240
         Width           =   8775
         Begin MSComctlLib.ListView ListView1 
            Height          =   3015
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5318
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Descripcion del insumo"
               Object.Width           =   12347
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cantidad"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Reporte"
         Height          =   2775
         Left            =   120
         TabIndex        =   59
         Top             =   3720
         Width           =   11775
         Begin VB.TextBox TxtReporte 
            Height          =   2415
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Top             =   240
            Width           =   11535
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Signos Vitales"
         Height          =   3375
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   2895
         Begin VB.TextBox TxtTalla 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            MaxLength       =   10
            TabIndex        =   53
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox TxtPeso 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            MaxLength       =   10
            TabIndex        =   52
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox TxtTemp 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            MaxLength       =   10
            TabIndex        =   51
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox TxtFc 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            MaxLength       =   10
            TabIndex        =   50
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox TxtTa 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            MaxLength       =   10
            TabIndex        =   49
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            Height          =   195
            Left            =   2160
            TabIndex        =   58
            Top             =   2370
            Width           =   210
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            Height          =   195
            Left            =   2160
            TabIndex        =   57
            Top             =   1890
            Width           =   195
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "º C"
            Height          =   195
            Left            =   2160
            TabIndex        =   56
            Top             =   1410
            Width           =   210
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lpm"
            Height          =   195
            Left            =   2160
            TabIndex        =   55
            Top             =   930
            Width           =   300
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mmHg"
            Height          =   195
            Left            =   2160
            TabIndex        =   54
            Top             =   450
            Width           =   450
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Talla:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   2370
            Width           =   390
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1890
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Temp:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FC:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   930
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TA:"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   450
            Width           =   255
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente2 
         Height          =   375
         Left            =   11280
         TabIndex        =   40
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   6600
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
         MICON           =   "FrmEnfermeria.frx":1002
         PICN            =   "FrmEnfermeria.frx":101E
         PICH            =   "FrmEnfermeria.frx":12B4
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
         TabIndex        =   41
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   6600
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
         MICON           =   "FrmEnfermeria.frx":1513
         PICN            =   "FrmEnfermeria.frx":152F
         PICH            =   "FrmEnfermeria.frx":17C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnReporteDiario 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   6600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Reporte Diario"
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
         MICON           =   "FrmEnfermeria.frx":1A20
         PICN            =   "FrmEnfermeria.frx":1A3C
         PICH            =   "FrmEnfermeria.frx":1CD4
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
         Left            =   2040
         TabIndex        =   65
         ToolTipText     =   "Reporte"
         Top             =   6600
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "FrmEnfermeria.frx":1F5B
         PICN            =   "FrmEnfermeria.frx":1F77
         PICH            =   "FrmEnfermeria.frx":209C
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   3840
      TabIndex        =   33
      Top             =   10080
      Width           =   8295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7200
         TabIndex        =   34
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
         MICON           =   "FrmEnfermeria.frx":232C
         PICN            =   "FrmEnfermeria.frx":2348
         PICH            =   "FrmEnfermeria.frx":2511
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
         TabIndex        =   35
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
         MICON           =   "FrmEnfermeria.frx":2746
         PICN            =   "FrmEnfermeria.frx":2762
         PICH            =   "FrmEnfermeria.frx":29F1
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
         TabIndex        =   36
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
         MICON           =   "FrmEnfermeria.frx":2E32
         PICN            =   "FrmEnfermeria.frx":2E4E
         PICH            =   "FrmEnfermeria.frx":2FDB
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
         TabIndex        =   37
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
         MICON           =   "FrmEnfermeria.frx":3210
         PICN            =   "FrmEnfermeria.frx":322C
         PICH            =   "FrmEnfermeria.frx":350E
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
         TabIndex        =   38
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
         MICON           =   "FrmEnfermeria.frx":375F
         PICN            =   "FrmEnfermeria.frx":377B
         PICH            =   "FrmEnfermeria.frx":391F
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
      TabIndex        =   30
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
         TabIndex        =   31
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad o Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   32
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
         MICON           =   "FrmEnfermeria.frx":3ABE
         PICN            =   "FrmEnfermeria.frx":3ADA
         PICH            =   "FrmEnfermeria.frx":3D3F
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
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
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
         TabIndex        =   8
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
         MICON           =   "FrmEnfermeria.frx":3FD1
         PICN            =   "FrmEnfermeria.frx":3FED
         PICH            =   "FrmEnfermeria.frx":4289
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
         TabIndex        =   9
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
         MICON           =   "FrmEnfermeria.frx":44BE
         PICN            =   "FrmEnfermeria.frx":44DA
         PICH            =   "FrmEnfermeria.frx":4763
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         MICON           =   "FrmEnfermeria.frx":49FB
         PICN            =   "FrmEnfermeria.frx":4A17
         PICH            =   "FrmEnfermeria.frx":4CAD
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
         TabIndex        =   15
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
         MICON           =   "FrmEnfermeria.frx":4F0C
         PICN            =   "FrmEnfermeria.frx":4F28
         PICH            =   "FrmEnfermeria.frx":51BD
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
         TabIndex        =   16
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
         MICON           =   "FrmEnfermeria.frx":5419
         PICN            =   "FrmEnfermeria.frx":5435
         PICH            =   "FrmEnfermeria.frx":55D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   375
         Left            =   8400
         TabIndex        =   62
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   7800
         TabIndex        =   63
         Top             =   2250
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         Height          =   195
         Left            =   5280
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1770
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Tratante:"
         Height          =   195
         Left            =   270
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   9960
         Picture         =   "FrmEnfermeria.frx":580E
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
         TabIndex        =   17
         Top             =   330
         Width           =   870
      End
   End
End
Attribute VB_Name = "FrmEnfermeria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents cSubLV As cSubclassListView
Attribute cSubLV.VB_VarHelpID = -1
Dim RsInformeMed As New ADODB.Recordset         'Para desplazamientos
Dim RsCargarPacientes As New ADODB.Recordset
Dim CSql As String
Dim Cambio
Dim actualiza
Dim IdReg
Dim MaxReg As String
Dim d1, d2, d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, dm, ds, de
Dim Gram As TextoPos
Dim Sesiones As Double
Dim tomog, Estadiaje
Dim MaxRegP As String
Dim RsIdMax As New ADODB.Recordset
Dim RsRegPendiente As New ADODB.Recordset
Dim IdInf
Dim OptBordesSel As String
Dim RsTemp As New ADODB.Recordset
Dim i As Integer
Dim IdRegFinal

Private Sub BtnAgregar_Click()
On Error Resume Next

If Trim(IdPac1) = "" Then MsgBox "Debe seleccionar un Paciente antes de agregar un Informe Medico!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub
IO = 1
Call Deshabilita_Btns
Call Blanqueo
Cambio = 0
actualiza = 0
Frame1.Enabled = False
Frame2.BackColor = &HE0E0E0
Frame4.BackColor = &HE0E0E0
Frame5.BackColor = &HE0E0E0
Frame6.BackColor = &HE0E0E0

DesactivarTextos
DtpFecha.Value = Now

End Sub

Sub ActivarTextos()
TxtTa.Locked = True
TxtFc.Locked = True
TxtTemp.Locked = True
TxtPeso.Locked = True
TxtTalla.Locked = True
TxtReporte.Locked = True
DtpFecha.Enabled = True
End Sub

Sub DesactivarTextos()
TxtTa.Locked = False
TxtFc.Locked = False
TxtTemp.Locked = False
TxtPeso.Locked = False
TxtTalla.Locked = False
TxtReporte.Locked = False
DtpFecha.Enabled = False
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

Private Sub BtnAnterior_Click()
On Error Resume Next

If RsCargarPacientes.RecordCount <> 0 Then
    Call Blanqueo
    If Not RsCargarPacientes.BOF Then
        RsCargarPacientes.MovePrevious
        If RsCargarPacientes.BOF Then RsCargarPacientes.MoveLast
        Call Carga_De_Datos
        Call CONSULTA_INFORME
        If RsInformeMed.RecordCount <> 0 Then
            'Call carga_datos_radio
            Cambio = 0
            actualiza = 1
            Else
            Call Habilita_Btns("sin informe")
            CboModificarMedicoTratante.Enabled = False
            Cambio = 0
            actualiza = 0
        End If
    End If
'    CARGAR_INFORME2
Else
    MsgBox "No se encontraron datos, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros cargados"
End If

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
'    Call carga_datos_radio
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

ActivarTextos
'CboModificarMedicoTratante.Enabled = False

Call Blanqueo
Call Habilita_Btns("sin informe")

Cambio = 0
actualiza = 0
Frame1.Enabled = True
Frame2.BackColor = &HEAEFEF


Frame2.BackColor = &HEAEFEF


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
    BtnAnterior1.Enabled = True
    BtnSiguiente2.Enabled = True
    
Call Carga_De_Datos
Call CONSULTA_INFORME

If Not (RsInformeMed.EOF) Then
    'Call carga_datos_radio
    Call Carga_Datos_Enfermeria
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

Sub Carga_Datos_Enfermeria()

On Error GoTo WrtError


If RsInformeMed.RecordCount <> 0 Then

    If Not IsNull(RsInformeMed.Fields("Ta").Value) Then TxtTa.Text = Val(RsInformeMed.Fields("Ta").Value) Else TxtTa.Text = ""
    If Not IsNull(RsInformeMed.Fields("Fc").Value) Then TxtFc.Text = Val(RsInformeMed.Fields("Fc").Value) Else TxtFc.Text = ""
    If Not IsNull(RsInformeMed.Fields("Temperatura").Value) Then TxtTemp.Text = RsInformeMed.Fields("Temperatura").Value Else TxtTemp.Text = ""
    If Not IsNull(RsInformeMed.Fields("Peso").Value) Then TxtPeso.Text = RsInformeMed.Fields("Peso").Value Else TxtPeso.Text = ""
    If Not IsNull(RsInformeMed.Fields("Talla").Value) Then TxtTalla.Text = RsInformeMed.Fields("Talla").Value Else TxtTalla.Text = ""
    If Not IsNull(RsInformeMed.Fields("ReporteCumplimiento").Value) Then TxtReporte.Text = RsInformeMed.Fields("ReporteCumplimiento").Value Else TxtReporte.Text = ""
    
    Dim RsConsumoM As New ADODB.Recordset
    CSql = "Select * From ConsumoEnfermeria Where IdConsumoM='" & RsInformeMed.Fields("IdConsumoM").Value & "'"
    Set RsConsumoM = CrearRS(CSql)
    If RsConsumoM.RecordCount > 0 Then
'        For i = 0 To ListView1.ListItems.Count
'            If RsConsumoM.Fields("IdProducto").Value = ListView1.SelectedItem.ListSubItems(2).Text Then
'                ListView1.ListItems(i).Checked = True
'                ListView1.ListItems(i).Text = RsConsumoM.Fields("Cantidad").Value
'            Else
'                ListView1.ListItems(i).Checked = False
'                ListView1.ListItems(i).Text = ""
'            End If
'        Next i
'        i = 0
'        Do
'            i = i + 1
'            Set Li = ListView1.ListItems.Add(, , RcsProcesUsu!NombreProceso)
'            If RcsProcesUsu!Permitido = "1" Then
'                ListView1.ListItems.Item(i).Checked = True
'            Else
'                ListView1.ListItems.Item(i).Checked = False
'            End If
'            RcsProcesUsu.MoveNext
'        Loop Until RcsProcesUsu.EOF = True
'    End If

Else
    IdReg = ""
    Cambio = 0
    actualiza = 0
    BtnEliminar.Enabled = False
    ActivarTextos
End If
'Cambio = 0
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

Sub Blanqueo()
TxtTa.Text = ""
TxtFc.Text = ""
TxtTemp.Text = ""
TxtPeso.Text = ""
TxtTalla.Text = ""
TxtReporte.Text = ""
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next

ActivarTextos
Call Blanqueo
Call Habilita_Btns("sin informe")

Cambio = 0
actualiza = 0
Frame1.Enabled = True
Frame2.BackColor = &HEAEFEF
Frame2.BackColor = &HEAEFEF
Frame4.BackColor = &HEAEFEF
Frame5.BackColor = &HEAEFEF
Frame6.BackColor = &HEAEFEF

Carga_De_Datos

Call CONSULTA_INFORME
If RsInformeMed.RecordCount <> 0 Then
    Call Carga_Datos_Enfermeria
    Cambio = 0
    actualiza = 1
Else
    Call Habilita_Btns("sin informe")
    Cambio = 0
    actualiza = 0
End If

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



Private Sub BtnEvolucionOncologica_Click()
'On Error Resume Next
Especia = "Radioterapia"
FrmEvolucion.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim RsGuardar As New ADODB.Recordset
On Error GoTo WrtError
Cambio = 1
If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente antes de Guardar!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub

If Trim(Text15.Text) = "" Then
    'f = "Antecedente_Flia"
    MsgBox "El Campo Antecedente Familiar esta Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    Text15.SetFocus
    Exit Sub
ElseIf Trim(Text19.Text) = "" Then
    'f = "Anatomia_Patol"
    MsgBox "El Campo Antecedente Patologico esta Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    Text19.SetFocus
    Exit Sub
ElseIf Trim(Text16.Text) = "" Then
    'f = "Enfermedad_Act"
    MsgBox "El Campo Enfermedad actual esta Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    Text16.SetFocus
    Exit Sub
'ElseIf (DateValue(DtpFecha.Value) - DateValue(Now)) < 0 Then
'    MsgBox "La fecha ingresada es Incorreta!", vbCritical + vbOKOnly, "Error"
'    DtpFecha.SetFocus
'    Exit Sub
ElseIf Trim(Text20.Text) = "" Then
    'f = "Examen_Fis"
    MsgBox "El Examen Fisico no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text20.SetFocus
    Exit Sub
ElseIf Text18.Text = "" Then
    'f = "Motivo_Con"
    MsgBox "El Motivo de la Consulta no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text18.SetFocus
    Exit Sub
ElseIf Text21.Text = "" Then
    'f = "Diagnotico"
    MsgBox "El Diagnóstico no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text21.SetFocus
    Exit Sub
ElseIf Text22.Text = "" Then
    'f = "Tratamiento"
    MsgBox "El Tratamiento no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Text22.SetFocus
    Exit Sub
ElseIf Trim(Combo1.List(Combo1.ListIndex)) = "" Then
    'f = "Tomografia"
    MsgBox "La Tomografia no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    Combo1.SetFocus
    Exit Sub
ElseIf Trim(CboMetas.Text) = "" Then
    'f = "Tomografia"
    MsgBox "Debe de seleccionar la Meta del Tratamiento del Paciente!", vbExclamation + vbOKOnly, "Faltan Datos"
    CboMetas.SetFocus
    Exit Sub
ElseIf CboTCancers.ListIndex < 0 Then
    MsgBox "Debe de seleccionar un tipo de CANCER!", vbExclamation + vbOKOnly, "Faltan Datos"
    CboTCancers.SetFocus
    Exit Sub
End If

' Obtiene el Nuevo ID para el informe medico
    Dim RsMaxReg As New ADODB.Recordset
    CSql = "Select max(IdInforme)+1 as MaxReg From Informe_Medico"
    Set RsMaxReg = CrearRS(CSql)

    If Not IsNull(RsMaxReg.Fields("MaxReg")) Then
        MaxReg = RsMaxReg.Fields("MaxReg")
    Else
        MaxReg = "0"
    End If
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm



' Obtiene el ID del medico
    If CboModificarMedicoTratante.ListIndex = -1 Then
        IdMedT = "0"
    Else
        IdMedT = CboModificarMedicoTratante.ItemData(CboModificarMedicoTratante.ListIndex)
        CSql = "update PACIENTE set Medico_tratante=" & IdMedT & " where IdPaciente = " & IdPac1
        Set RsGuardar = CrearRS(CSql)
        
        CSql = "select * from Paciente"
        Set RsCargarPacientes = CrearRS(CSql)
        RsCargarPacientes.Find "IdPaciente=" & IdPac1
    End If
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm

    If Cambio = 1 Then
        If Combo1.ListIndex = -1 Then
            tomog = 0
        ElseIf Combo1.ListIndex = 0 Then
            tomog = 0
        Else
            tomog = 1
        End If
        
         
        If Text8.Text = "" And Text9.Text = "" Then
            Text2.Text = 0
            Sesiones = 0
        ElseIf Text8.Text <> "" And Text9.Text = "" Then
            Text2.Text = CDbl(Text8.Text)
            Sesiones = 0
        ElseIf Text8.Text = "" And Text9.Text <> "" Then
            Text2.Text = CDbl(Text9.Text)
            Sesiones = 0
        ElseIf Text8.Text <> "" And Text9.Text <> "" Then
            Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
            Sesiones = Text2.Text
        ElseIf Text8.Text <> 0 And Text9.Text <> 0 Then
            Text2.Text = CDbl(Text8.Text) / CDbl(Text9.Text)
            Sesiones = Text2.Text
         ElseIf Text8.Text = 0 And Text9.Text = 0 Then
            Sesiones = 0
            Text2.Text = 0
        ElseIf Text8.Text = 0 And Text9.Text <> 0 Then
            Text2.Text = 0
            Sesiones = 0
        ElseIf Text8.Text <> 0 And Text9.Text = 0 Then
            Text2.Text = 0
            Sesiones = 0
        End If
        
'        CSql = "Insert into INFORME_MEDICO(IdInforme, idmedicot,idusuario,idpaciente,Antecedente_Flia," & _
'               "Anatomia_Patol,Enfermedad_Act,Examen_Fis,Motivo_Con,Diagnotico,Tratamiento,Fecha,dosis," & _
'               "dosisd,Tomografia,Cuantas,Estado,Metas,sesiones) VALUES(" & MaxReg & "," & IdMedT & "," & IdUser & "," & IdPac1 & _
'               ",'" & Text15.Text & "','" & Text19.Text & "','" & Text16.Text & "','" & Text20.Text & _
'               "','" & Text18.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & _
'               Format(DtpFecha.Value, "DD/MM/YYYY") & "'," & Val(Text8.Text) & "," & Val(Text9.Text) & _
'               "," & tomog & "," & Val(Text17.Text) & ",1,'" & Trim(CboMetas.Text) & "'," & Sesiones & ")"
'
'        Set RsGuardar = CrearRS(CSql)
        
        If actualiza = 1 Then
            
            If IdReg = "" Then MsgBox "No hay informes seleccionados!", vbExclamation + vbOKOnly, "Error": Exit Sub
            
            CSql = "Select * From Informe_Medico Where IdPaciente='" & IdPac1 & "' And IdInforme='" & IdInf & "'"
            Set RsGuardar = CrearRS(CSql)
        ElseIf actualiza = 0 Then
            CSql = "Select * From Informe_Medico"
            Set RsGuardar = CrearRS(CSql)
            
            RsGuardar.AddNew
            RsGuardar.Fields("IdInforme").Value = MaxReg
            RsGuardar.Fields("IdMedicot").Value = IdMedT
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            IdInf = MaxReg
        End If
        
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        RsGuardar.Fields("Antecedente_Flia").Value = Trim(Text15.Text)
        RsGuardar.Fields("Anatomia_Patol").Value = Trim(Text19.Text)
        RsGuardar.Fields("Enfermedad_Act").Value = Trim(Text16.Text)
        RsGuardar.Fields("Examen_Fis").Value = Trim(Text20.Text)
        RsGuardar.Fields("Motivo_Con").Value = Trim(Text18.Text)
        RsGuardar.Fields("Diagnotico").Value = Trim(Text21.Text)
        RsGuardar.Fields("Tratamiento").Value = Trim(Text22.Text)
        RsGuardar.Fields("Fecha").Value = Format(DtpFecha.Value, "dd/MM/yyyy")
        RsGuardar.Fields("dosis").Value = Val(Text8.Text)
        RsGuardar.Fields("dosisd").Value = Val(Text9.Text)
        RsGuardar.Fields("Tomografia").Value = tomog
        RsGuardar.Fields("Cuantas").Value = Val(Text17.Text)
        RsGuardar.Fields("Estado").Value = 1
        RsGuardar.Fields("Metas").Value = Trim(CboMetas.Text)
        RsGuardar.Fields("sesiones").Value = Sesiones
        RsGuardar.Fields("Estadiaje").Value = Estadiaje
        
        RsGuardar.Update
        
        Call Habilita_Btns("Todos")
        Frame2.BackColor = &HEAEFEF

        
        MsgBox "Registro actualizado Satisfactoriamente, se procedera a guardar los datos estadisticos!s", vbInformation + vbOKOnly, "Operacion Exitosa"
        ActivarTextos
        Call CONSULTA_INFORME
        
        'Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        'MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Informe Médico"
'        EnviarAlHosting
'        EnviarRegPendiente
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        CSql = "SELECT * FROM Informe_Medico2 WHERE IdInforme=" & IdInf
        Set RsGuardar = CrearRS(CSql)
        
        Dim NuevoId As Integer
        If RsGuardar.RecordCount = 0 Then
            CSql = "SELECT MAX(Id) + 1 AS NuevoId FROM Informe_Medico2"
            Set RsGuardar = CrearRS(CSql)
            
            If Not IsNull(RsGuardar.Fields(0).Value) Then
                NuevoId = Val(RsGuardar.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If
            
            CSql = "SELECT * FROM Informe_Medico2"
            Set RsGuardar = CrearRS(CSql)
            
            RsGuardar.AddNew
            
            RsGuardar.Fields("Id").Value = NuevoId
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdInforme").Value = IdInf
            
        End If
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        RsGuardar.Fields("CP").Value = Trim(Combo3.Text)
        RsGuardar.Fields("T").Value = Trim(Combo4.Text)
        RsGuardar.Fields("N").Value = Trim(Combo5.Text)
        RsGuardar.Fields("M").Value = Trim(Combo6.Text)
        RsGuardar.Fields("Estadio").Value = Trim(Combo2.Text)
        RsGuardar.Fields("G").Value = Trim(Combo7.Text)
        If ChkGleason.Value Then RsGuardar.Fields("Gleason").Value = Trim(TxtGleason.Text) Else RsGuardar.Fields("Gleason").Value = ""
        RsGuardar.Fields("Reseccion").Value = OptBordesSel
        RsGuardar.Fields("IdTipoCancer").Value = CboTCancers.ItemData(CboTCancers.ListIndex)
        
'        If ChkGlobal(0).Value Then RsGuardar.Fields("Re").Value = Trim(CboGeneral(0).Text) Else RsGuardar.Fields("Re").Value = ""
'        If ChkGlobal(1).Value Then RsGuardar.Fields("Rp").Value = Trim(CboGeneral(1).Text) Else RsGuardar.Fields("Rp").Value = ""
'        If ChkGlobal(2).Value Then RsGuardar.Fields("Her2Neu").Value = Trim(CboGeneral(2).Text) Else RsGuardar.Fields("Her2Neu").Value = ""
'        If ChkGlobal(3).Value Then RsGuardar.Fields("EMA").Value = Trim(CboGeneral(3).Text) Else RsGuardar.Fields("EMA").Value = ""
'        If ChkGlobal(4).Value Then RsGuardar.Fields("VIM").Value = Trim(CboGeneral(4).Text) Else RsGuardar.Fields("VIM").Value = ""
'        If ChkGlobal(5).Value Then RsGuardar.Fields("CAE").Value = Trim(CboGeneral(5).Text) Else RsGuardar.Fields("CAE").Value = ""
'        If ChkGlobal(6).Value Then RsGuardar.Fields("CERB-2").Value = Trim(CboGeneral(6).Text) Else RsGuardar.Fields("CERB-2").Value = ""
'        If ChkGlobal(7).Value Then RsGuardar.Fields("P53").Value = Trim(CboGeneral(7).Text) Else RsGuardar.Fields("P53").Value = ""
'        If ChkGlobal(8).Value Then RsGuardar.Fields("DESMINA").Value = Trim(CboGeneral(8).Text) Else RsGuardar.Fields("DESMINA").Value = ""
'        If ChkGlobal(9).Value Then RsGuardar.Fields("ACE").Value = Trim(CboGeneral(9).Text) Else RsGuardar.Fields("ACE").Value = ""
'
'        If ChkGlobal(10).Value Then RsGuardar.Fields("AFP").Value = Trim(CboGeneral(10).Text) Else RsGuardar.Fields("AFP").Value = ""
'        If ChkGlobal(11).Value Then RsGuardar.Fields("PROT-S-100").Value = Trim(CboGeneral(11).Text) Else RsGuardar.Fields("PROT-S-100").Value = ""
'        If ChkGlobal(12).Value Then RsGuardar.Fields("PGP").Value = Trim(CboGeneral(12).Text) Else RsGuardar.Fields("PGP").Value = ""
'        If ChkGlobal(13).Value Then RsGuardar.Fields("CD31").Value = Trim(CboGeneral(13).Text) Else RsGuardar.Fields("CD31").Value = ""
'        If ChkGlobal(14).Value Then RsGuardar.Fields("CD34").Value = Trim(CboGeneral(14).Text) Else RsGuardar.Fields("CD34").Value = ""
'        If ChkGlobal(15).Value Then RsGuardar.Fields("CD117").Value = Trim(CboGeneral(15).Text) Else RsGuardar.Fields("CD117").Value = ""
'        If ChkGlobal(16).Value Then RsGuardar.Fields("CK5").Value = Trim(CboGeneral(16).Text) Else RsGuardar.Fields("CK5").Value = ""
'        If ChkGlobal(17).Value Then RsGuardar.Fields("CK6").Value = Trim(CboGeneral(17).Text) Else RsGuardar.Fields("CK6").Value = ""
'        If ChkGlobal(18).Value Then RsGuardar.Fields("CK7").Value = Trim(CboGeneral(18).Text) Else RsGuardar.Fields("CK7").Value = ""
'        If ChkGlobal(19).Value Then RsGuardar.Fields("CK20").Value = Trim(CboGeneral(19).Text) Else RsGuardar.Fields("CK20").Value = ""
'
'        If ChkGlobal(20).Value Then RsGuardar.Fields("CAM 5,2").Value = Trim(CboGeneral(20).Text) Else RsGuardar.Fields("CAM 5,2").Value = ""
'        If ChkGlobal(21).Value Then RsGuardar.Fields("TTF-1").Value = Trim(CboGeneral(21).Text) Else RsGuardar.Fields("TTF-1").Value = ""
'        If ChkGlobal(22).Value Then RsGuardar.Fields("CROMOGRANINA").Value = Trim(CboGeneral(22).Text) Else RsGuardar.Fields("CROMOGRANINA").Value = ""
'        If ChkGlobal(23).Value Then RsGuardar.Fields("SINAPTOFISINA").Value = Trim(CboGeneral(23).Text) Else RsGuardar.Fields("SINAPTOFISINA").Value = ""
'        If ChkGlobal(24).Value Then RsGuardar.Fields("CD56").Value = Trim(CboGeneral(24).Text) Else RsGuardar.Fields("CD56").Value = ""
'        If ChkGlobal(25).Value Then RsGuardar.Fields("CD57").Value = Trim(CboGeneral(25).Text) Else RsGuardar.Fields("CD57").Value = ""
'        If ChkGlobal(26).Value Then RsGuardar.Fields("EGFR").Value = Trim(CboGeneral(26).Text) Else RsGuardar.Fields("EGFR").Value = ""
'        If ChkGlobal(27).Value Then RsGuardar.Fields("KIT").Value = Trim(CboGeneral(27).Text) Else RsGuardar.Fields("KIT").Value = ""
'        If ChkGlobal(28).Value Then RsGuardar.Fields("AE1").Value = Trim(CboGeneral(28).Text) Else RsGuardar.Fields("AE1").Value = ""
'        If ChkGlobal(29).Value Then RsGuardar.Fields("AE3").Value = Trim(CboGeneral(29).Text) Else RsGuardar.Fields("AE3").Value = ""
'
'        If ChkGlobal(30).Value Then RsGuardar.Fields("CK903").Value = Trim(CboGeneral(30).Text) Else RsGuardar.Fields("CK903").Value = ""
'        If ChkGlobal(31).Value Then RsGuardar.Fields("GFAP").Value = Trim(CboGeneral(31).Text) Else RsGuardar.Fields("GFAP").Value = ""
'        If ChkGlobal(32).Value Then RsGuardar.Fields("SMA").Value = Trim(CboGeneral(32).Text) Else RsGuardar.Fields("SMA").Value = ""
'        If ChkGlobal(33).Value Then RsGuardar.Fields("CA199").Value = Trim(CboGeneral(33).Text) Else RsGuardar.Fields("CA199").Value = ""
'        If ChkGlobal(34).Value Then RsGuardar.Fields("CA125").Value = Trim(CboGeneral(34).Text) Else RsGuardar.Fields("CA125").Value = ""
'        If ChkGlobal(35).Value Then RsGuardar.Fields("CEA").Value = Trim(CboGeneral(35).Text) Else RsGuardar.Fields("CEA").Value = ""
'        If ChkGlobal(36).Value Then RsGuardar.Fields("CEA-D14").Value = Trim(CboGeneral(36).Text) Else RsGuardar.Fields("CEA-D14").Value = ""
'        If ChkGlobal(37).Value Then RsGuardar.Fields("E-CAD").Value = Trim(CboGeneral(37).Text) Else RsGuardar.Fields("E-CAD").Value = ""
'        If ChkGlobal(38).Value Then RsGuardar.Fields("HCG").Value = Trim(CboGeneral(38).Text) Else RsGuardar.Fields("HCG").Value = ""
'        If ChkGlobal(39).Value Then RsGuardar.Fields("HMB-45").Value = Trim(CboGeneral(39).Text) Else RsGuardar.Fields("HMB-45").Value = ""
'
'        If ChkGlobal(40).Value Then RsGuardar.Fields("HPAP").Value = Trim(CboGeneral(40).Text) Else RsGuardar.Fields("HPAP").Value = ""
'        If ChkGlobal(41).Value Then RsGuardar.Fields("WT1").Value = Trim(CboGeneral(41).Text) Else RsGuardar.Fields("WT1").Value = ""
'        If ChkGlobal(42).Value Then RsGuardar.Fields("BEL-1").Value = Trim(CboGeneral(42).Text) Else RsGuardar.Fields("BEL-1").Value = ""
'        If ChkGlobal(43).Value Then RsGuardar.Fields("BEL-2").Value = Trim(CboGeneral(43).Text) Else RsGuardar.Fields("BEL-2").Value = ""
'        If ChkGlobal(44).Value Then RsGuardar.Fields("PRB").Value = Trim(CboGeneral(44).Text) Else RsGuardar.Fields("PRB").Value = ""
'        If ChkGlobal(45).Value Then RsGuardar.Fields("ALK-1").Value = Trim(CboGeneral(45).Text) Else RsGuardar.Fields("ALK-1").Value = ""
'        If ChkGlobal(46).Value Then RsGuardar.Fields("RA").Value = Trim(CboGeneral(46).Text) Else RsGuardar.Fields("RA").Value = ""
'        If ChkGlobal(47).Value Then RsGuardar.Fields("CD99MID2").Value = Trim(CboGeneral(47).Text) Else RsGuardar.Fields("CD99MID2").Value = ""
'        If ChkGlobal(48).Value Then RsGuardar.Fields("NSD").Value = Trim(CboGeneral(48).Text) Else RsGuardar.Fields("NSD").Value = ""
'        If ChkGlobal(49).Value Then RsGuardar.Fields("LCACD45").Value = Trim(CboGeneral(49).Text) Else RsGuardar.Fields("LCACD45").Value = ""
'
'        If ChkGlobal(50).Value Then RsGuardar.Fields("CD20L26").Value = Trim(CboGeneral(50).Text) Else RsGuardar.Fields("CD20L26").Value = ""
'        If ChkGlobal(51).Value Then RsGuardar.Fields("CD79A").Value = Trim(CboGeneral(51).Text) Else RsGuardar.Fields("CD79A").Value = ""
'        If ChkGlobal(52).Value Then RsGuardar.Fields("CD45ROUCHL1").Value = Trim(CboGeneral(52).Text) Else RsGuardar.Fields("CD45ROUCHL1").Value = ""
'        If ChkGlobal(53).Value Then RsGuardar.Fields("CD3").Value = Trim(CboGeneral(53).Text) Else RsGuardar.Fields("CD3").Value = ""
'        If ChkGlobal(54).Value Then RsGuardar.Fields("CD30KL1BERH2").Value = Trim(CboGeneral(54).Text) Else RsGuardar.Fields("CD30KL1BERH2").Value = ""
'        If ChkGlobal(55).Value Then RsGuardar.Fields("CD15LEUM1").Value = Trim(CboGeneral(55).Text) Else RsGuardar.Fields("CD15LEUM1").Value = ""
'        If ChkGlobal(56).Value Then RsGuardar.Fields("WT").Value = Trim(CboGeneral(56).Text) Else RsGuardar.Fields("WT").Value = ""
'        If ChkGlobal(57).Value Then RsGuardar.Fields("OTROS").Value = Trim(TxtOtros.Text) Else RsGuardar.Fields("OTROS").Value = ""
'
'        RsGuardar.Update
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
       
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        
        CSql = "SELECT * FROM Informe_Medico3 WHERE IdInforme=" & IdInf
        Set RsGuardar = CrearRS(CSql)
        
        If RsGuardar.RecordCount = 0 Then
            CSql = "SELECT MAX(Id) + 1 AS NuevoId FROM Informe_Medico3"
            Set RsGuardar = CrearRS(CSql)
            
            If Not IsNull(RsGuardar.Fields(0).Value) Then
                NuevoId = Val(RsGuardar.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If
            
            CSql = "SELECT * FROM Informe_Medico3"
            Set RsGuardar = CrearRS(CSql)
            
            RsGuardar.AddNew
            
            RsGuardar.Fields("Id").Value = NuevoId
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdInforme").Value = IdInf
            
        End If
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
'
'        If LstEnfer(0).Value Then RsGuardar.Fields("LE_MinM").Value = Trim(Text5(0).Text) Else RsGuardar.Fields("LE_MinM").Value = ""
'        If LstEnfer(1).Value Then RsGuardar.Fields("LE_6M").Value = True Else RsGuardar.Fields("LE_6M").Value = False
'        If LstEnfer(2).Value Then RsGuardar.Fields("LE_12M").Value = True Else RsGuardar.Fields("LE_12M").Value = False
'        If LstEnfer(3).Value Then RsGuardar.Fields("LE_18M").Value = True Else RsGuardar.Fields("LE_18M").Value = False
'        If LstEnfer(4).Value Then RsGuardar.Fields("LE_24M").Value = True Else RsGuardar.Fields("LE_24M").Value = False
'        If LstEnfer(5).Value Then RsGuardar.Fields("LE_30M").Value = True Else RsGuardar.Fields("LE_30M").Value = False
'        If LstEnfer(6).Value Then RsGuardar.Fields("LE_36M").Value = True Else RsGuardar.Fields("LE_36M").Value = False
'        If LstEnfer(7).Value Then RsGuardar.Fields("LE_42M").Value = True Else RsGuardar.Fields("LE_42M").Value = False
'        If LstEnfer(8).Value Then RsGuardar.Fields("LE_48M").Value = True Else RsGuardar.Fields("LE_48M").Value = False
'        If LstEnfer(9).Value Then RsGuardar.Fields("LE_54M").Value = True Else RsGuardar.Fields("LE_54M").Value = False
'        If LstEnfer(10).Value Then RsGuardar.Fields("LE_60M").Value = True Else RsGuardar.Fields("LE_60M").Value = False
'        If LstEnfer(11).Value Then RsGuardar.Fields("LE_MaxM").Value = Trim(Text5(1).Text) Else RsGuardar.Fields("LE_MaxM").Value = ""
'
'        If LstProg(0).Value Then RsGuardar.Fields("P_MinM").Value = Trim(Text5(2).Text) Else RsGuardar.Fields("P_MinM").Value = ""
'        If LstProg(1).Value Then RsGuardar.Fields("P_6M").Value = True Else RsGuardar.Fields("P_6M").Value = False
'        If LstProg(2).Value Then RsGuardar.Fields("P_12M").Value = True Else RsGuardar.Fields("P_12M").Value = False
'        If LstProg(3).Value Then RsGuardar.Fields("P_18M").Value = True Else RsGuardar.Fields("P_18M").Value = False
'        If LstProg(4).Value Then RsGuardar.Fields("P_24M").Value = True Else RsGuardar.Fields("P_24M").Value = False
'        If LstProg(5).Value Then RsGuardar.Fields("P_30M").Value = True Else RsGuardar.Fields("P_30M").Value = False
'        If LstProg(6).Value Then RsGuardar.Fields("P_36M").Value = True Else RsGuardar.Fields("P_36M").Value = False
'        If LstProg(7).Value Then RsGuardar.Fields("P_42M").Value = True Else RsGuardar.Fields("P_42M").Value = False
'        If LstProg(8).Value Then RsGuardar.Fields("P_48M").Value = True Else RsGuardar.Fields("P_48M").Value = False
'        If LstProg(9).Value Then RsGuardar.Fields("P_54M").Value = True Else RsGuardar.Fields("P_54M").Value = False
'        If LstProg(10).Value Then RsGuardar.Fields("P_60M").Value = True Else RsGuardar.Fields("P_60M").Value = False
'        If LstProg(11).Value Then RsGuardar.Fields("P_MaxM").Value = Trim(Text5(3).Text) Else RsGuardar.Fields("P_MaxM").Value = ""
'
'        If ChkRecaida(0).Value Then RsGuardar.Fields("R_MinM").Value = Trim(Text5(4).Text) Else RsGuardar.Fields("R_MinM").Value = ""
'        If ChkRecaida(1).Value Then RsGuardar.Fields("R_6M").Value = True Else RsGuardar.Fields("R_6M").Value = False
'        If ChkRecaida(2).Value Then RsGuardar.Fields("R_12M").Value = True Else RsGuardar.Fields("R_12M").Value = False
'        If ChkRecaida(3).Value Then RsGuardar.Fields("R_18M").Value = True Else RsGuardar.Fields("R_18M").Value = False
'        If ChkRecaida(4).Value Then RsGuardar.Fields("R_24M").Value = True Else RsGuardar.Fields("R_24M").Value = False
'        If ChkRecaida(5).Value Then RsGuardar.Fields("R_30M").Value = True Else RsGuardar.Fields("R_30M").Value = False
'        If ChkRecaida(6).Value Then RsGuardar.Fields("R_36M").Value = True Else RsGuardar.Fields("R_36M").Value = False
'        If ChkRecaida(7).Value Then RsGuardar.Fields("R_42M").Value = True Else RsGuardar.Fields("R_42M").Value = False
'        If ChkRecaida(8).Value Then RsGuardar.Fields("R_48M").Value = True Else RsGuardar.Fields("R_48M").Value = False
'        If ChkRecaida(9).Value Then RsGuardar.Fields("R_54M").Value = True Else RsGuardar.Fields("R_54M").Value = False
'        If ChkRecaida(10).Value Then RsGuardar.Fields("R_60M").Value = True Else RsGuardar.Fields("R_60M").Value = False
'        If ChkRecaida(11).Value Then RsGuardar.Fields("R_MaxM").Value = Trim(Text5(5).Text) Else RsGuardar.Fields("R_MaxM").Value = ""
'
'        If ChkMuert(0).Value Then RsGuardar.Fields("M_MinM").Value = Trim(Text5(6).Text) Else RsGuardar.Fields("M_MinM").Value = ""
'        If ChkMuert(1).Value Then RsGuardar.Fields("M_6M").Value = True Else RsGuardar.Fields("M_6M").Value = False
'        If ChkMuert(2).Value Then RsGuardar.Fields("M_12M").Value = True Else RsGuardar.Fields("M_12M").Value = False
'        If ChkMuert(3).Value Then RsGuardar.Fields("M_18M").Value = True Else RsGuardar.Fields("M_18M").Value = False
'        If ChkMuert(4).Value Then RsGuardar.Fields("M_24M").Value = True Else RsGuardar.Fields("M_24M").Value = False
'        If ChkMuert(5).Value Then RsGuardar.Fields("M_30M").Value = True Else RsGuardar.Fields("M_30M").Value = False
'        If ChkMuert(6).Value Then RsGuardar.Fields("M_36M").Value = True Else RsGuardar.Fields("M_36M").Value = False
'        If ChkMuert(7).Value Then RsGuardar.Fields("M_42M").Value = True Else RsGuardar.Fields("M_42M").Value = False
'        If ChkMuert(8).Value Then RsGuardar.Fields("M_48M").Value = True Else RsGuardar.Fields("M_48M").Value = False
'        If ChkMuert(9).Value Then RsGuardar.Fields("M_54M").Value = True Else RsGuardar.Fields("M_54M").Value = False
'        If ChkMuert(10).Value Then RsGuardar.Fields("M_60M").Value = True Else RsGuardar.Fields("M_60M").Value = False
'        If ChkMuert(11).Value Then RsGuardar.Fields("M_MaxM").Value = Trim(Text5(7).Text) Else RsGuardar.Fields("M_MaxM").Value = ""
'
'        RsGuardar.Update
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        If Trim(TxtDosisT.Text) = "" Then TxtDosisT.Text = "0"
        If Trim(TxtDosisD.Text) = "" Then TxtDosisD.Text = "0"
        If Trim(TxtSesionesFin.Text) = "" Then TxtSesionesFin.Text = "0"
        
        CSql = "SELECT * FROM Informe_Medico4 WHERE IdInforme=" & IdInf
        Set RsGuardar = CrearRS(CSql)
        
        If RsGuardar.RecordCount = 0 Then
            CSql = "SELECT MAX(Id) + 1 AS NuevoId FROM Informe_Medico4"
            Set RsGuardar = CrearRS(CSql)
            
            If Not IsNull(RsGuardar.Fields(0).Value) Then
                NuevoId = Val(RsGuardar.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If
            
            CSql = "SELECT * FROM Informe_Medico4"
            Set RsGuardar = CrearRS(CSql)
            
            RsGuardar.AddNew
            
            RsGuardar.Fields("Id").Value = NuevoId
            RsGuardar.Fields("IdPaciente").Value = IdPac1
            RsGuardar.Fields("IdInforme").Value = IdInf
            
        End If
        
        RsGuardar.Fields("IdUsuario").Value = IdUser
        
        RsGuardar.Fields("ExamFisicoInicial").Value = Trim(TxtExamFIni.Text)
        RsGuardar.Fields("ExamFisicoFinal").Value = Trim(TxtExamFin.Text)
        RsGuardar.Fields("Diagnostico").Value = Trim(TxtDiagFin.Text)
        RsGuardar.Fields("Anatomia").Value = Trim(TxtAnatFin.Text)
        RsGuardar.Fields("Complicaciones").Value = Trim(TxtCompliFin.Text)
        RsGuardar.Fields("Indice").Value = Trim(TxtInidiceFin.Text)
        RsGuardar.Fields("Comentarios").Value = Trim(TxtTratamientoFin.Text)
        RsGuardar.Fields("DosisT").Value = CDbl(TxtDosisT.Text)
        RsGuardar.Fields("DosisD").Value = CDbl(TxtDosisD.Text)
        RsGuardar.Fields("Sesiones").Value = Val(TxtSesionesFin.Text)
        
        If Combo9.ListIndex <> -1 Then
            RsGuardar.Fields("IdMedicoT").Value = Combo9.ItemData(Combo9.ListIndex)
        Else
            RsGuardar.Fields("IdMedicoT").Value = "0"
        End If
        
        If Combo8.ListIndex <> -1 Then
            RsGuardar.Fields("Metas").Value = Trim(Combo8.Text)
        Else
            RsGuardar.Fields("Metas").Value = ""
        End If
        
        RsGuardar.Update
        
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        
        BtnDesHacer_Click
        MsgBox "Los datos fueron guardardos!", vbInformation + vbOKOnly, "Operación exitosa!"
        Exit Sub
    End If
       MsgBox "No hay cambios para guardar", vbExclamation + vbOKOnly, "No hay cambios"
       Frame1.Enabled = True
       
    If RsInformeMed.RecordCount <> 0 Then
'        Call carga_datos_radio
        Cambio = 0
        actualiza = 1
        Else
        actualiza = 0
        Cambio = 0
    End If
    Intentos = 0
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly + vbCritical, "Error al Guardar"
    
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





Private Sub BtnListaEspera_Click()
ModulO = 4
FrmListaEspera.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnLlamar_Click()
On Error Resume Next
Call Llamar
End Sub

Private Sub BtnReporteDiario_Click()
FrmEnfermeriaReporte.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnSiguiente_Click()
On Error Resume Next
Call Blanqueo

If RsCargarPacientes.RecordCount <> 0 Then
    If Not RsCargarPacientes.EOF Then
        RsCargarPacientes.MoveNext
        If RsCargarPacientes.EOF Then RsCargarPacientes.MoveFirst
        Carga_De_Datos
        Call CONSULTA_INFORME
        If RsInformeMed.RecordCount <> 0 Then
'            Call carga_datos_radio
            Cambio = 0
            actualiza = 1
            Else
            Call Habilita_Btns("sin informe")
            CboModificarMedicoTratante.Enabled = False
            Cambio = 0
            actualiza = 0
        End If
    End If
    'CARGAR_INFORME2
Else
    MsgBox "No se encontraron datos, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros cargados"
End If


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
    'Call carga_datos_radio
    Cambio = 0
    actualiza = 1
    Else
    MsgBox "No existen informes medicos para este paciente!", vbExclamation + vbOKOnly, "No existen los datos"
End If
End Sub


Sub CONSULTA_INFORME()
On Error GoTo WrtError

CSql = "Select * From Enfermeria Where IdPaciente = " & IdPac1 & " And Estado=1 Order By Fecha Desc"
Set RsInformeMed = CrearRS(CSql)
    
Exit Sub
WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Sub Carga_De_Datos()
On Error GoTo WrtError

    Text1.Text = RsCargarPacientes.Fields("cedulaP").Value
    DtpFechaRegistro.Value = RsCargarPacientes.Fields("Fecha_regp").Value
    Text3.Text = RsCargarPacientes.Fields("Nombrep").Value
    Text4.Text = RsCargarPacientes.Fields("Apellidop").Value
    DtpFechaNac = RsCargarPacientes.Fields("Fecha_nacimientop").Value
    Text6.Text = RsCargarPacientes.Fields("Edadp").Value
    IdPac1 = RsCargarPacientes.Fields("idpaciente").Value
    Me.Caption = "Enfermería - Paciente: " & IdPac1
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


Private Sub Form_Load()
On Error GoTo WrtError

Centrar Me
ModulO = 4

    Set cSubLV = New cSubclassListView

    cSubLV.SubClassListView ListView1
    Dim SpathSkin As String
    
    SpathSkin = App.Path & "\skin1"
    Call setColumnHeader(SpathSkin, vbBlack, vbBlack, True, False, Me.BackColor, vbBlack)
    Dim Item As ListItem
    With ListView1
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Descripcion del insumo", 6800
        .ColumnHeaders.Add , , "Cantidad", 1300
        .ColumnHeaders.Add , , "Codigo", 0
    End With

    
CSql = "select * from Paciente Order by IdPaciente"
Set RsCargarPacientes = CrearRS(CSql)

Llenar_Listado_Productos

Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
'MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Sub setColumnHeader( _
    SpathSkin As String, _
    lColorNormal As Long, _
    lColorUp As Long, _
    Optional bIconAlingmentRight As Boolean = False, _
    Optional bTextBold As Boolean = False, _
    Optional BackColortextBox As Long = vbWhite, _
    Optional ForeColorTxtEdit As Long = vbBlack)
    
    
    With cSubLV
        .SkinPicture = LoadPicture(SpathSkin & ".bmp")
        .TextNormalColor = lColorNormal
        .TextResalteColor = lColorUp
        .IconAlingmentRight = bIconAlingmentRight
        .HedersFontBlod = bTextBold
        
        .SetPropertyTextBoxEdit BackColortextBox, ForeColorTxtEdit
    End With
End Sub
Sub Llenar_Listado_Productos()
Dim RsLlenar_Listado_Productos As New ADODB.Recordset

CSql = "Select * From Productos Where Ubicacion='Farmacia'"
Set RsLlenar_Listado_Productos = CrearRS(CSql)

ListView1.ListItems.Clear

Do While Not RsLlenar_Listado_Productos.EOF
    With ListView1
        i = i + 1
        .ListItems.Add , , RsLlenar_Listado_Productos.Fields("Descripcion").Value
        .ListItems(i).ListSubItems.Add , , ""
        .ListItems(i).ListSubItems.Add , , RsLlenar_Listado_Productos.Fields("IdProducto").Value
        
    End With
    RsLlenar_Listado_Productos.MoveNext
Loop

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
Private Sub cSubLV_AfterEdit(ByVal Columna As Integer, Cancel As Boolean, Value As Variant)
Dim lBcolor As Long
    Dim lFColor As Long
    
'    Select Case Combo1.ListIndex
'        Case 0: lBcolor = &HC0FFFF: lFColor = vbBlack
'        Case 1: lBcolor = &HC0FFC0: lFColor = vbBlack
'        Case 2: lBcolor = RGB(214, 228, 243): lFColor = vbBlack
'        Case 3: lBcolor = &HC0FFC0: lFColor = vbBlack
'    End Select
    
    Select Case Columna
        Case 1
           If Not IsNumeric(Value) Then
              Cancel = True
              cSubLV.SowToolTipText "Dato no válido", "El valor debe ser un número", TTIconWarning, TTBalloon, False, lFColor, lBcolor, 5000, 0
           End If
    End Select

End Sub


Private Sub cSubLV_beforeEdit(ByVal Columna As Integer, Cancel As Boolean)
    If Columna = 2 Then
       Cancel = True
       cSubLV.SowToolTipText "Info", "Esta columna es de solo lectura y no se puede editar", TTIconInfo, TTBalloon, False, , , 5000, 0
    End If
End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub


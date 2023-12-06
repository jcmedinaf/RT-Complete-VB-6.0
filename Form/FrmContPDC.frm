VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmContPDC 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan De Cuenta"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15105
   Icon            =   "FrmContPDC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   15105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "=== Selección de la Empresa ==="
      Height          =   7095
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   14895
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   5160
         TabIndex        =   42
         Top             =   6240
         Width           =   9615
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   8280
            TabIndex        =   22
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
            MICON           =   "FrmContPDC.frx":1002
            PICN            =   "FrmContPDC.frx":101E
            PICH            =   "FrmContPDC.frx":11E7
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
            Left            =   6960
            TabIndex        =   21
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
            MICON           =   "FrmContPDC.frx":141C
            PICN            =   "FrmContPDC.frx":1438
            PICH            =   "FrmContPDC.frx":171A
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
         Caption         =   "Rif"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   26
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Plan de Cuenta"
         Height          =   6255
         Left            =   5160
         TabIndex        =   32
         Top             =   0
         Width           =   9615
         Begin MSMask.MaskEdBox TxtFormato1 
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            Top             =   2400
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   14737632
            PromptInclude   =   0   'False
            HideSelection   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TxtFormato 
            Height          =   375
            Left            =   1320
            TabIndex        =   0
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            HideSelection   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   435
            Left            =   1320
            TabIndex        =   1
            Top             =   840
            Width           =   3375
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Descripción"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   17
            Top             =   5820
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Identificador"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   16
            Top             =   5820
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Grupo"
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   18
            Top             =   5820
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Centro de Costo FIJO"
            Height          =   195
            Left            =   6720
            TabIndex        =   10
            Top             =   840
            Width           =   1935
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "FrmContPDC.frx":196B
            Left            =   6360
            List            =   "FrmContPDC.frx":196D
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1320
            Width           =   2775
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Maneja Terceros"
            Height          =   195
            Left            =   5040
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Maneja Bases"
            Height          =   195
            Left            =   6720
            TabIndex        =   9
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "FrmContPDC.frx":196F
            Left            =   1920
            List            =   "FrmContPDC.frx":197F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2040
            Width           =   2775
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmContPDC.frx":19BA
            Left            =   1920
            List            =   "FrmContPDC.frx":19C7
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1680
            Width           =   2775
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmContPDC.frx":19EB
            Left            =   1920
            List            =   "FrmContPDC.frx":19FB
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   2775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "De Movimiento"
            Height          =   195
            Left            =   5040
            TabIndex        =   7
            Top             =   480
            Width           =   1575
         End
         Begin SystemOncoAmerica.DMGrid DMGrid2 
            Height          =   2175
            Left            =   360
            TabIndex        =   15
            Top             =   3480
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   3836
            Object.Width           =   9105
            Object.Height          =   2145
            ScrollBar       =   1
            MarqueeStyle    =   2
         End
         Begin ChamaleonButton.ChameleonBtn BtnEliminarDescripcion 
            Height          =   375
            Left            =   8280
            TabIndex        =   20
            ToolTipText     =   "Eliminar"
            Top             =   5760
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "FrmContPDC.frx":1A31
            PICN            =   "FrmContPDC.frx":1A4D
            PICH            =   "FrmContPDC.frx":1BF1
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
            Left            =   6960
            TabIndex        =   19
            ToolTipText     =   "Deshacer Operacion"
            Top             =   5760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Importar"
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
            MICON           =   "FrmContPDC.frx":1D90
            PICN            =   "FrmContPDC.frx":1DAC
            PICH            =   "FrmContPDC.frx":2036
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox TxtFormato2 
            Height          =   375
            Left            =   2160
            TabIndex        =   6
            Top             =   2880
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   14737632
            PromptInclude   =   0   'False
            HideSelection   =   0   'False
            PromptChar      =   " "
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
            Height          =   375
            Left            =   6960
            TabIndex        =   13
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   2880
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
            MICON           =   "FrmContPDC.frx":22BF
            PICN            =   "FrmContPDC.frx":22DB
            PICH            =   "FrmContPDC.frx":256A
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
            Left            =   8280
            TabIndex        =   14
            ToolTipText     =   "Eliminar"
            Top             =   2880
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Borrar"
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
            MICON           =   "FrmContPDC.frx":29AB
            PICN            =   "FrmContPDC.frx":29C7
            PICH            =   "FrmContPDC.frx":2B6B
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
            Left            =   5640
            TabIndex        =   12
            ToolTipText     =   "Agregar"
            Top             =   2880
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
            MICON           =   "FrmContPDC.frx":2D0A
            PICN            =   "FrmContPDC.frx":2D26
            PICH            =   "FrmContPDC.frx":2EB3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   360
            TabIndex        =   41
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenar por:"
            Height          =   195
            Left            =   360
            TabIndex        =   40
            Top             =   5850
            Width           =   885
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Centro de Costo:"
            Height          =   195
            Left            =   5040
            TabIndex        =   39
            Top             =   1380
            Width           =   1185
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Cuenta de Corrección:"
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   2985
            Width           =   1590
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Cuenta de Ajuste:"
            Height          =   195
            Left            =   480
            TabIndex        =   37
            Top             =   2520
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tipo de Cuenta:"
            Height          =   195
            Left            =   360
            TabIndex        =   36
            Top             =   2100
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Clasificación:"
            Height          =   195
            Left            =   360
            TabIndex        =   35
            Top             =   1740
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tipo de Actividad:"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   1380
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Formato:"
            Height          =   195
            Left            =   360
            TabIndex        =   33
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   6240
         Width           =   4935
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
            TabIndex        =   27
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Código, Nombre o Rif"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   28
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Busqueda"
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
            MICON           =   "FrmContPDC.frx":30E8
            PICN            =   "FrmContPDC.frx":3104
            PICH            =   "FrmContPDC.frx":3369
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   3240
            Top             =   360
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         Top             =   5880
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   25
         Top             =   5880
         Width           =   975
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   9551
         Object.Width           =   4905
         Object.Height          =   5385
         ScrollBar       =   1
         DrawColorGrid   =   1
         MarqueeStyle    =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   5880
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmContPDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEmpresas As Recordset
Dim RsPDC As Recordset
Dim RsTemp As Recordset
Dim IdEmpresa As Integer
Dim IdPDC As Integer
Dim Spdr As Integer
Dim IdReng As Integer
Dim RegNew As Boolean
Dim PFormato(10) As Byte
Dim Formato As String
Dim i As Integer
Dim j As Integer
Const vbGris = &HE0E0E0

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 3
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre de la Empresa"
DMGrid1.DColumnas(3).Caption = "Rif"
End Sub

Sub IniDMGrid2()
' carga las columnas y encabezados de columna
DMGrid2.Cols = 4
DMGrid2.Rows = 0

DMGrid2.DColumnas(4).Visible = False

DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 0
DMGrid2.DColumnas(1).Locked = True
DMGrid2.DColumnas(2).Locked = True
DMGrid2.DColumnas(3).Locked = True
DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 70 / 100) - 300
DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(1).Caption = "Identificador"
DMGrid2.DColumnas(2).Caption = "Descripcion"
DMGrid2.DColumnas(3).Caption = "Tipo"
End Sub

Sub Blanqueo2()
    TxtFormato.Text = Replace(Formato, "X", " ")
    TxtFormato1.Text = Replace(Formato, "X", " ")
    TxtFormato2.Text = Replace(Formato, "X", " ")
    TxtDescripcion.Text = ""
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
End Sub

Sub Blanqueo()
    TxtFormato.Mask = ""
    TxtFormato1.Mask = ""
    TxtFormato2.Mask = ""
    TxtFormato.Text = ""
    TxtFormato1.Text = ""
    TxtFormato2.Text = ""
    TxtDescripcion.Text = ""
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    DMGrid2.Clear
    DMGrid2.Rows = 0
    DMGrid2.PaintMGrid
End Sub

Private Sub BtnAgregar_Click()
If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
If IdPDC = 0 Then MsgBox "Primero debe CREAR UNA CONFIGURACIÓN DE PLAN DE CUENTAS!", vbExclamation + vbOKOnly, "Error": Exit Sub
Blanqueo2
RegNew = True
BtnEliminar.Enabled = False
BtnAgregar.Enabled = False
End Sub

Private Sub BtnBuscar_Click()
Dim TamDMGrid As Integer
Dim Reng1 As Integer
Dim Reng2 As String
Dim Reng3 As String
Dim Band As Boolean

TamDMGrid = DMGrid1.Rows
Band = False
For i = 1 To TamDMGrid
    Reng1 = DMGrid1.ValorCelda(i, 1)
    Reng2 = DMGrid1.ValorCelda(i, 2)
    Reng3 = DMGrid1.ValorCelda(i, 3)
    
    If Val(Trim(TxtBuscar.Text)) = Reng1 Or UCase(Trim(TxtBuscar.Text)) = UCase(Reng2) Or UCase(Trim(TxtBuscar.Text)) = UCase(Reng3) Then
        DMGrid1.Row = i
        Band = True
        Exit For
    End If
Next

If Band = False Then MsgBox "No se encontraron los datos!", vbInformation + vbOKOnly, "La busqueda ha finalizado."
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Form_Load
End Sub

Private Sub BtnEliminar_Click()
Dim resp As Byte

If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Se procederá a eliminar el Plan de Cuentas actual de la empresa " _
        & DMGrid1.ValorCelda(DMGrid1.Row, 2) & Chr(13) & Chr(13) & "Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")

If resp = vbNo Then Exit Sub

CSql = "UPDATE ContPDC SET Activo=0 WHERE Activo=1 AND IdEmpresa=" & IdEmpresa
Set RsTemp = CrearRS(CSql)

MsgBox "El Plan de Cuentas ha sido Eliminado!", vbInformation + vbOKOnly, "Operación Exitosa."

Option1_Click (0)
End Sub

Private Sub BtnEliminarDescripcion_Click()
Dim resp As Byte

If DMGrid2.Rows = 0 Then Exit Sub

If IdReng = 0 Then MsgBox "Seleccione una categoria del Plan de Cuentas la cual se eliminará!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Se procederá a eliminar la fila (" & DMGrid2.Row & ") seleccionada, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = vbNo Then Exit Sub

' Verificar que no hallan comprobantes ya creados y guardados en la BD
' DMGrid2.ValorCelda(1, 1) = ""

CSql = "DELETE FROM ContPDC WHERE IdPDC=" & IdReng
Set RsTemp = CrearRS(CSql)

Call DMGrid1_MouseUpC(vbLeftButton, 0, 0, DMGrid1.Row, 1)
MsgBox "El registro seleccionado fue eliminado!", vbExclamation + vbOKOnly, "Operación Exitosa."
IdReng = 0

Option2_Click (0)

End Sub

Private Sub BtnGuardarActualizar_Click()
Dim resp As Byte
Dim NuevoId As Integer
Dim Opc As Byte
Dim Formt As String
Dim Formt1 As String
Dim Formt2 As String

If Trim(Replace(TxtFormato.Text, Chr(Spdr), "")) = "" Then
    MsgBox "Ingrese el formato!", vbExclamation + vbOKOnly, "Error"
    TxtFormato.SetFocus
    Exit Sub
ElseIf Trim(TxtDescripcion.Text) = "" Then
    MsgBox "Ingrese la descripción!", vbExclamation + vbOKOnly, "Error"
    TxtDescripcion.SetFocus
    Exit Sub
ElseIf Combo1.ListIndex = -1 Then
    MsgBox "Ingrese el tipo de actividad!", vbExclamation + vbOKOnly, "Error"
    Combo1.SetFocus
    Exit Sub
ElseIf Combo2.ListIndex = -1 Then
    MsgBox "Ingrese la clasificación!", vbExclamation + vbOKOnly, "Error"
    Combo2.SetFocus
    Exit Sub
ElseIf Combo3.ListIndex = -1 Then
    MsgBox "Ingrese el tipo de cuenta!", vbExclamation + vbOKOnly, "Error"
    Combo3.SetFocus
    Exit Sub
ElseIf Combo4.ListIndex = -1 Then
    Combo4.ListIndex = 0
End If

resp = MsgBox("se procederá a guardar el Plan de cuentas para la Empresa, Desea continuar?", vbQuestion + vbYesNo, "Confimar")
If resp = vbNo Then Exit Sub

CSql = "SELECT MAX(IdPDC)+1 as NuevoId FROM contPDC"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If

If Check1.Value = 1 Then Opc = 1 Else Opc = 0

Formt = TxtFormato.FormattedText
Formt1 = TxtFormato1.FormattedText
Formt2 = TxtFormato2.FormattedText
Dim UPos As Byte
For i = 1 To Len(Formt)
    If IsNumeric(Mid(Formt, i, 1)) Then UPos = i
Next i
Formt = Mid(Formt, 1, UPos)
UPos = 0
For i = 1 To Len(Formt1)
    If IsNumeric(Mid(Formt1, i, 1)) Then UPos = i
Next i
Formt1 = Mid(Formt1, 1, UPos)
UPos = 0
For i = 1 To Len(Formt2)
    If IsNumeric(Mid(Formt2, i, 1)) Then UPos = i
Next i
Formt2 = Mid(Formt2, 1, UPos)

If RegNew Then
    CSql = "INSERT INTO ContPDC (IdPDC,IdEmpresa,Identificador,Nombre,Tipo,Activo,Movimiento,Terceros,CCFijos,Bases,CentroCosto," & _
    "TipoActividad,Clasificacion,TipoCuenta,CuentaAjusta,CuentaCorreccion,IdUser) VALUES (" & _
    NuevoId & " ," & IdEmpresa & ",'" & Formt & "','" & Trim(TxtDescripcion.Text) & _
    "','" & Opc & "','1'," & Check1.Value & "," & Check3.Value & "," & Check4.Value & "," & Check2.Value & _
    "," & Combo4.ItemData(Combo4.ListIndex) & ",'" & Combo1.ItemData(Combo1.ListIndex) & "','" & Combo2.ItemData(Combo2.ListIndex) & _
    "','" & Combo3.ItemData(Combo3.ListIndex) & "','" & Formt1 & "','" & Formt2 & "', " & IdUser & ")"

    Set RsTemp = CrearRS(CSql)
Else
    CSql = "UPDATE ContPDC SET IdEmpresa=" & IdEmpresa & ",Identificador='" & Formt & _
    "',Nombre='" & Trim(TxtDescripcion.Text) & "',Tipo='" & Opc & "',Activo='1', Movimiento=" & Check1.Value & _
    ",Terceros=" & Check3.Value & ",CCFijos=" & Check4.Value & ",Bases=" & Check2.Value & ",CentroCosto=" & Combo4.ItemData(Combo4.ListIndex) & _
    ",TipoActividad='" & Combo1.ItemData(Combo1.ListIndex) & "',Clasificacion='" & Combo2.ItemData(Combo2.ListIndex) & _
    "',TipoCuenta='" & Combo3.ItemData(Combo3.ListIndex) & "',CuentaAjusta='" & Formt1 & _
    "',CuentaCorreccion='" & Formt2 & "',IdUser=" & IdUser & " WHERE IdPDC=" & IdReng
    Set RsTemp = CrearRS(CSql)
End If

Call DMGrid1_MouseUpC(vbLeftButton, 0, 0, DMGrid1.Row, 1)

DMGrid2.Row = DMGrid2.Rows
MsgBox "Los cambios han sido guardados!", vbInformation + vbOKOnly, "Operación Exitosa."

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Check3.Value = 1
End Sub

Private Sub Check3_Click()
If Check3.Value = 0 Then
    If Check2.Value = 1 Then Check2.Value = 0
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo2.SetFocus: Exit Sub
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo3.SetFocus: Exit Sub
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtFormato1.SetFocus: Exit Sub
End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub

Private Sub Combo3_Click()
Combo3_Change
End Sub

Private Sub Combo2_Change()
If Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = -1 Then
        TxtFormato1.BackColor = vbGris: TxtFormato1.Text = Replace(Trim(Formato), "X", " ")
        TxtFormato2.BackColor = vbGris: TxtFormato2.Text = Replace(Trim(Formato), "X", " ")
    Else
        If Combo3.ListIndex = 1 Then
            TxtFormato1.BackColor = vbGris: TxtFormato1.Text = Replace(Trim(Formato), "X", " ")
            TxtFormato2.BackColor = vbWhite
        ElseIf Combo3.ListIndex = 2 Then
            TxtFormato1.BackColor = vbGris: TxtFormato1.Text = Replace(Trim(Formato), "X", " ")
            TxtFormato2.BackColor = vbGris: TxtFormato2.Text = Replace(Trim(Formato), "X", " ")
        Else
            TxtFormato1.BackColor = vbWhite
            TxtFormato2.BackColor = vbWhite
        End If
    End If
Else
    TxtFormato1.BackColor = vbGris: TxtFormato1.Text = Replace(Trim(Formato), "X", " ")
    TxtFormato2.BackColor = vbGris: TxtFormato2.Text = Replace(Trim(Formato), "X", " ")
End If
End Sub

Private Sub Combo3_Change()
If Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 1 Then
        TxtFormato1.BackColor = vbGris: TxtFormato1.Text = Replace(Trim(Formato), "X", " ")
        TxtFormato2.BackColor = vbWhite
    ElseIf Combo3.ListIndex = 2 Then
        TxtFormato1.BackColor = vbGris: TxtFormato1.Text = Replace(Trim(Formato), "X", " ")
        TxtFormato2.BackColor = vbGris: TxtFormato2.Text = Replace(Trim(Formato), "X", " ")
    Else
        TxtFormato1.BackColor = vbWhite
        TxtFormato2.BackColor = vbWhite
    End If
End If
End Sub


Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
Dim TamCad As Byte

If Button = vbLeftButton Then
    IdReng = 0
    Blanqueo
    If lRow = 0 Then Exit Sub
    IdEmpresa = DMGrid1.ValorCelda(lRow, 1)

    Frame3.Caption = "Configuracion del PDC para " & DMGrid1.ValorCelda(lRow, 2)
    CSql = "Select * From ContPDCConfig WHERE IdEmpresa=" & IdEmpresa & " And activo=1"
    Set RsTemp = CrearRS(CSql)
    DMGrid2.Rows = 0
    If RsTemp.RecordCount <> 0 Then
        IdPDC = Val(RsTemp.Fields("IdEmpresa").Value)
        Spdr = Asc(RsTemp.Fields("Separador").Value)
        BtnAgregar.Enabled = True
        
        TxtFormato.Mask = Replace(Trim(RsTemp.Fields("Formato").Value), "X", "#")
        TxtFormato1.Mask = Replace(Trim(RsTemp.Fields("Formato").Value), "X", "#")
        TxtFormato2.Mask = Replace(Trim(RsTemp.Fields("Formato").Value), "X", "#")
        Formato = Trim(RsTemp.Fields("Formato").Value)
        
        Frame3.Enabled = True
        
        ' Ahora consultara la Base de datos para mostrar el plan de cuentas de la empresa (En el caso de que
        ' tenga una ya creada)
        CSql = "Select * From ContPDC WHERE IdEmpresa=" & IdEmpresa & " And activo=1 order by IdPDC"
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
            
            TxtFormato.Text = Format(Replace(Trim(RsTemp.Fields("Identificador").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
            TxtFormato1.Text = Format(Replace(Trim(RsTemp.Fields("CuentaAjusta").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
            TxtFormato2.Text = Format(Replace(Trim(RsTemp.Fields("CuentaCorreccion").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
            
            If RsTemp.Fields("Movimiento").Value Then Check1.Value = 1 Else Check1.Value = 0
            If RsTemp.Fields("Bases").Value Then Check2.Value = 1 Else Check2.Value = 0
            If RsTemp.Fields("Terceros").Value Then Check3.Value = 1 Else Check3.Value = 0
            If RsTemp.Fields("CCFijos").Value Then Check4.Value = 1 Else Check4.Value = 0
    
            For i = 0 To Combo1.ListCount - 1
                If Combo1.ItemData(i) = Val(RsTemp.Fields("TipoActividad").Value) Then Combo1.ListIndex = i: Exit For
            Next i
            For i = 0 To Combo2.ListCount - 1
                If Combo2.ItemData(i) = Val(RsTemp.Fields("Clasificacion").Value) Then Combo2.ListIndex = i: Exit For
            Next i
            For i = 0 To Combo3.ListCount - 1
                If Combo3.ItemData(i) = Val(RsTemp.Fields("TipoCuenta").Value) Then Combo3.ListIndex = i: Exit For
            Next i
            
            ' Para CENTROS DE CONSTOS...
            'For i = 0 To Combo1.ListCount - 1
            '    If Combo1.List(i) = RsTemp.Fields("TipoActividad").Value Then Combo1.ListIndex = i: Exit For
            'Next i
            
            ' Ciclo condicional para llenar el DMGrid2 con el plan de cuentas de la empresa selecciona,
            ' si la empresa no tiene un PDC, entonces mostrara todos los campos en blancos.
            While Not RsTemp.EOF
                DMGrid2.Rows = DMGrid2.Rows + 1
                DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsTemp.Fields("Identificador").Value
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsTemp.Fields("Nombre").Value
                
                If RsTemp.Fields("Movimiento").Value Then
                    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = "Mvto"
                Else
                    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = "Grupo"
                End If
                DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsTemp.Fields("IdPDC").Value
                RsTemp.MoveNext
            Wend
            BtnEliminar.Enabled = True
            Option2_Click (0)
        Else
            BtnEliminar.Enabled = False
        End If
        RegNew = True
    Else
        MsgBox "La Empresa '" & DMGrid1.ValorCelda(DMGrid1.Row, 2) & "' no contiene una Configuración de P.D.C.!", vbExclamation + vbOKOnly, "No tiene un Plan de Cuenta!"
        Frame3.Enabled = False
        RegNew = True
        IdPDC = 0
        Spdr = 0
        Formato = ""
        BtnAgregar.Enabled = False
    End If
    DMGrid2.PaintMGrid
End If
End Sub

Private Sub DMGrid2_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton Then

    If lRow = 0 Then Exit Sub
    IdReng = DMGrid2.ValorCelda(lRow, 4)
    
    CSql = "Select * From ContPDC WHERE IdEmpresa=" & IdEmpresa & " And activo=1 AND IdPDC=" & IdReng & "order by IdPDC"
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then
        RegNew = False
        BtnAgregar.Enabled = True
        
        TxtFormato.Text = Format(Replace(Trim(RsTemp.Fields("Identificador").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
        TxtFormato1.Text = Format(Replace(Trim(RsTemp.Fields("CuentaAjusta").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
        TxtFormato2.Text = Format(Replace(Trim(RsTemp.Fields("CuentaCorreccion").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
        TxtDescripcion.Text = Trim(RsTemp.Fields("Nombre").Value)
        
        If RsTemp.Fields("Movimiento").Value Then Check1.Value = 1 Else Check1.Value = 0
        If RsTemp.Fields("Bases").Value Then Check2.Value = 1 Else Check2.Value = 0
        If RsTemp.Fields("Terceros").Value Then Check3.Value = 1 Else Check3.Value = 0
        If RsTemp.Fields("CCFijos").Value Then Check4.Value = 1 Else Check4.Value = 0

        For i = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(i) = Val(RsTemp.Fields("TipoActividad").Value) Then Combo1.ListIndex = i: Exit For
        Next i
        For i = 0 To Combo2.ListCount - 1
            If Combo2.ItemData(i) = Val(RsTemp.Fields("Clasificacion").Value) Then Combo2.ListIndex = i: Exit For
        Next i
        For i = 0 To Combo3.ListCount - 1
            If Combo3.ItemData(i) = Val(RsTemp.Fields("TipoCuenta").Value) Then Combo3.ListIndex = i: Exit For
        Next i
        
        ' Para CENTROS DE CONSTOS...
        'For i = 0 To Combo1.ListCount - 1
        '    If Combo1.List(i) = RsTemp.Fields("TipoActividad").Value Then Combo1.ListIndex = i: Exit For
        'Next i
        
        ' Ciclo condicional para llenar el DMGrid2 con el plan de cuentas de la empresa selecciona,
        ' si la empresa no tiene un PDC, entonces mostrara todos los campos en blancos.
        BtnEliminarDescripcion.Enabled = True
    Else
        RegNew = True
        BtnEliminarDescripcion.Enabled = False
        BtnAgregar.Enabled = False
    End If
End If
End Sub

Private Sub DMGrid2_RowColChange(ByVal antRow As Integer, ByVal antCol As Integer, ByVal actRow As Integer, ByVal actCol As Integer)
Call DMGrid2_MouseUpC(vbLeftButton, 0, 0, actRow, actCol)
End Sub

Private Sub Form_Load()
Centrar Me
Blanqueo
RegNew = True
Frame3.Enabled = False
Combo4.Clear
Combo4.AddItem ""
Combo4.ItemData(Combo4.NewIndex) = 0

' cargar los Centros de Costos....

IniDMGrid
IniDMGrid2
Option1_Click (0)
End Sub
 
Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    CSql = "Select * From ContEmpresas where activo=1 order by IdEmpresa"
ElseIf Index = 1 Then
    CSql = "Select * From ContEmpresas where activo=1 order by Nombre"
ElseIf Index = 2 Then
    CSql = "Select * From ContEmpresas where activo=1 order by Rif"
End If

Set RsEmpresas = CrearRS(CSql)
DMGrid1.Rows = 0

If RsEmpresas.RecordCount = 0 Then Exit Sub
RsEmpresas.MoveFirst

While Not RsEmpresas.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsEmpresas.Fields("IdEmpresa")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsEmpresas.Fields("Nombre")
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsEmpresas.Fields("Rif")
    
    CSql = "Select * From ContPDCConfig WHERE IdEmpresa=" & RsEmpresas.Fields("IdEmpresa") & " And activo=1"
    Set RsTemp = CrearRS(CSql)
    If RsTemp.RecordCount = 0 Then
        DMGrid1.RowBackColor DMGrid1.Rows, RGB(221, 221, 221)
    Else
        DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
    End If
    
    RsEmpresas.MoveNext
Wend
DMGrid1.PaintMGrid
End Sub

Private Sub Option2_Click(Index As Integer)
If Index = 0 Then
    CSql = "Select * From ContPDC where activo=1 order by Identificador"
ElseIf Index = 1 Then
    CSql = "Select * From ContPDC where activo=1 order by Nombre"
ElseIf Index = 2 Then
    CSql = "Select * From ContPDC where activo=1 order by Tipo"
End If

Set RsPDC = CrearRS(CSql)
DMGrid2.Rows = 0

RsPDC.MoveFirst

While Not RsPDC.EOF
    DMGrid2.Rows = DMGrid2.Rows + 1
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsPDC.Fields("Identificador")
    DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsPDC.Fields("Nombre")
    
    If RsPDC.Fields("Movimiento").Value Then
        DMGrid2.ValorCelda(DMGrid2.Rows, 3) = "Mvto"
    Else
        DMGrid2.ValorCelda(DMGrid2.Rows, 3) = "Grupo"
    End If
    DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsPDC.Fields("IdPDC")
    RsPDC.MoveNext
Wend
DMGrid2.PaintMGrid
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo1.SetFocus: Exit Sub
End Sub

Private Sub TxtFormato_KeyPress(KeyAscii As Integer)
Dim a As String
Dim b As String
Dim TamFormt As Byte

If KeyAscii = 13 Then TxtDescripcion.SetFocus: Exit Sub

If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
    KeyAscii = 0
Else
    TamFormt = Len(TxtFormato.Text)
    If TamFormt >= CByte(Len(Formato)) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtFormato_LostFocus()
Dim IdentTemp As String
Dim TamDMGrid As Integer
Dim Formt As String
Dim UPos As Byte
Dim Band As Boolean

TamDMGrid = DMGrid2.Rows

Formt = TxtFormato.FormattedText
For i = 1 To Len(Formt)
    If IsNumeric(Mid(Formt, i, 1)) Then UPos = i
Next i
Formt = Mid(Formt, 1, UPos)


' CICLO QUE VERIFICA SI EL FORMATO INGRESADO EXITE.
For i = 1 To TamDMGrid
    IdentTemp = Trim(DMGrid2.ValorCelda(i, 1))
    If IdentTemp = Formt Then
        MsgBox "El Formato que ingreso ya se encuentra en la tabla!", vbExclamation + vbOKOnly, "El Identificador ya existe!"
        TxtFormato.Text = Replace(Trim(Formato), "X", " ")
        TxtFormato.SetFocus
        Exit For
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' CICLO QUE VERIFICA EXISTE UN NIVEL BASE PARA EL FORMATO INGRESADO
UPos = 0

Formt = TxtFormato.FormattedText
For i = 1 To Len(Formt)
    If Mid(Formt, i, 1) = Chr(Spdr) Then
        UPos = i
        If Not IsNumeric(Mid(Formt, i + 1, 1)) Then Exit For
    End If
    If i = Len(Formt) And IsNumeric(Mid(Formt, Len(Formt) - 1, 1)) Then UPos = i
Next i
If UPos <= 2 Then Exit Sub
Formt = Mid(Formt, 1, UPos - 3)

'"X.X.XX.XX.XX"
'"#.#.0#.0#.0#"

If Mid(Formt, Len(Formt), 1) = Chr(Spdr) Then Formt = Mid(Formt, 1, Len(Formt) - 1)
Band = False
For i = 1 To TamDMGrid
    IdentTemp = Trim(DMGrid2.ValorCelda(i, 1))
    If IdentTemp = Formt Then
        Band = True
        Exit For
    End If
Next i

If Band = False Then
    MsgBox "Debe primero crear el nivel " & Formt & " !", vbExclamation + vbOKOnly, "Información."
    TxtFormato.Text = Replace(Trim(Formato), "X", " ")
    TxtFormato.SetFocus
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
End Sub

Private Sub TxtFormato1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtFormato2.SetFocus: Exit Sub
KeyAscii = 0
If TxtFormato1.BackColor <> vbGris Then
    Tipo = "PDCFormato1"
    FrmContListaPDC.IdEmpresa = IdEmpresa
    FrmContListaPDC.Show vbModal, FrmPrincipal
End If
End Sub

Private Sub TxtFormato2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Check1.SetFocus: Exit Sub
KeyAscii = 0
If TxtFormato2.BackColor <> vbGris Then
    Tipo = "PDCFormato2"
    FrmContListaPDC.IdEmpresa = IdEmpresa
    FrmContListaPDC.Show vbModal, FrmPrincipal
End If
End Sub

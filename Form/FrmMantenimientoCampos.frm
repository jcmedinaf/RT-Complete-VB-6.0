VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmMantenimientosDeCampos 
   BackColor       =   &H00EAEFEF&
   Caption         =   "Mantenimiento de Campos Restringidos"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   Icon            =   "FrmMantenimientoCampos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   7215
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6120
            TabIndex        =   10
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
            MICON           =   "FrmMantenimientoCampos.frx":1002
            PICN            =   "FrmMantenimientoCampos.frx":101E
            PICH            =   "FrmMantenimientoCampos.frx":11E7
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
            Left            =   4800
            TabIndex        =   11
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
            MICON           =   "FrmMantenimientoCampos.frx":141C
            PICN            =   "FrmMantenimientoCampos.frx":1438
            PICH            =   "FrmMantenimientoCampos.frx":171A
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
            Left            =   1440
            TabIndex        =   12
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
            MICON           =   "FrmMantenimientoCampos.frx":196B
            PICN            =   "FrmMantenimientoCampos.frx":1987
            PICH            =   "FrmMantenimientoCampos.frx":1B2B
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
            TabIndex        =   13
            ToolTipText     =   "Agregar"
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
            MICON           =   "FrmMantenimientoCampos.frx":1CCA
            PICN            =   "FrmMantenimientoCampos.frx":1CE6
            PICH            =   "FrmMantenimientoCampos.frx":1E73
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
         Caption         =   "Campos Restringidos "
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7215
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2760
            TabIndex        =   3
            Top             =   3480
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   5520
            TabIndex        =   2
            Top             =   3480
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   4
            Top             =   3480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   56492035
            CurrentDate     =   40238
         End
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   3135
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5530
            Object.Width           =   6945
            Object.Height          =   3105
            ScrollBar       =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   3570
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   2280
            TabIndex        =   7
            Top             =   3570
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Multip.:"
            Height          =   195
            Left            =   4800
            TabIndex        =   6
            Top             =   3570
            Width           =   510
         End
      End
   End
End
Attribute VB_Name = "FrmMantenimientosDeCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub IniDMGrid()
' Carga las columnas y encabezados de columna
DMGrid1.Cols = 4
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1

DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True

DMGrid1.DColumnas(3).IsNumber = True
DMGrid1.DColumnas(4).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 40 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 20 / 100) - 300
'DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 40 / 100) - 300

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre del Campo"
DMGrid1.DColumnas(3).Caption = "Valor Predeterminado"
DMGrid1.DColumnas(4).Caption = "Multiplicador"
'DMGrid1.DColumnas(5).Caption = "Fecha de Creación"

'DMGrid1.DColumnas(1).Visible = False
DMGrid1.PaintMGrid
End Sub


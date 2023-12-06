VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContIndicePrecio 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indice  de Precio (IPC)"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   Icon            =   "FrmContIndicePrecio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   7335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   6240
         TabIndex        =   7
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":1002
         PICN            =   "FrmContIndicePrecio.frx":101E
         PICH            =   "FrmContIndicePrecio.frx":11E7
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
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":141C
         PICN            =   "FrmContIndicePrecio.frx":1438
         PICH            =   "FrmContIndicePrecio.frx":16C7
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
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":1B08
         PICN            =   "FrmContIndicePrecio.frx":1B24
         PICH            =   "FrmContIndicePrecio.frx":1CB1
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
         Height          =   495
         Left            =   5040
         TabIndex        =   10
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":1EE6
         PICN            =   "FrmContIndicePrecio.frx":1F02
         PICH            =   "FrmContIndicePrecio.frx":21E4
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
         Height          =   495
         Left            =   4320
         TabIndex        =   11
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":2435
         PICN            =   "FrmContIndicePrecio.frx":2451
         PICH            =   "FrmContIndicePrecio.frx":26E7
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
         Height          =   495
         Left            =   3720
         TabIndex        =   12
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":2946
         PICN            =   "FrmContIndicePrecio.frx":2962
         PICH            =   "FrmContIndicePrecio.frx":2BF7
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
         Height          =   495
         Left            =   2400
         TabIndex        =   13
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         MICON           =   "FrmContIndicePrecio.frx":2E53
         PICN            =   "FrmContIndicePrecio.frx":2E6F
         PICH            =   "FrmContIndicePrecio.frx":3013
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
      Height          =   5655
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.TextBox TxtAnio 
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox CboMeses 
         Height          =   315
         ItemData        =   "FrmContIndicePrecio.frx":31B2
         Left            =   720
         List            =   "FrmContIndicePrecio.frx":31DA
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   810
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1290
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5295
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   9340
         Object.Width           =   3465
         Object.Height          =   5265
      End
   End
End
Attribute VB_Name = "FrmContIndicePrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
NewReg = 1
Blanqueo
CboMeses.SetFocus
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
End Sub

Private Sub BtnGuardarActualizar_Click()
'---------------------------------------
'Valida campos
'---------------------------------------
If CboMeses.Text = "" Then
    MsgBox "Debe de Seleccionar el Mes!", vbCritical + vbOKOnly, "Error"
    CboMeses.SetFocus
    Exit Sub
End If

If TxtAnio.Text = "" Then
    MsgBox "Debe de ingresar el Año!", vbCritical + vbOKOnly, "Error"
    TxtAnio.SetFocus
    Exit Sub
End If

If TxtMonto.Text = "" Then
    MsgBox "Debe de ingresar el monto!", vbCritical + vbOKOnly, "Error"
    TxtMonto.SetFocus
    Exit Sub
End If

'Guardar y/o Actualiza
Select Case NewReg
    Case Is = 1
    'Agrega Nuevo Registro
    Case Is = 2
    'Actualiza el registro
End Select
End Sub

Private Sub Form_Load()
NewReg = 1
InitGrid
End Sub

Sub InitGrid()
DMGrid1.Cols = 2
DMGrid1.Rows = 0

DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 50 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 40 / 100)
DMGrid1.DColumnas(1).Caption = "Mes - Año"
DMGrid1.DColumnas(2).Caption = "Monto"


End Sub

Sub Blanqueo()
TxtMonto.Text = "0,00"
CboMeses.Text = ""
TxtAnio.Text = ""
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmAntecedentes 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Antecedentes y Quimioterapias"
   ClientHeight    =   5970
   ClientLeft      =   4875
   ClientTop       =   555
   ClientWidth     =   11175
   Icon            =   "Antecedente.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Antecedentes"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   120
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Quimioterapias"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   67
      Top             =   5040
      Width           =   10935
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   9840
         TabIndex        =   68
         ToolTipText     =   "Cerrar "
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
         MICON           =   "Antecedente.frx":1002
         PICN            =   "Antecedente.frx":101E
         PICH            =   "Antecedente.frx":11E7
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
         TabIndex        =   69
         ToolTipText     =   "Guardar / Actualizar"
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
         MICON           =   "Antecedente.frx":141C
         PICN            =   "Antecedente.frx":1438
         PICH            =   "Antecedente.frx":16C7
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
         TabIndex        =   70
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
         MICON           =   "Antecedente.frx":1B08
         PICN            =   "Antecedente.frx":1B24
         PICH            =   "Antecedente.frx":1CB1
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
         Left            =   8640
         TabIndex        =   71
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
         MICON           =   "Antecedente.frx":1EE6
         PICN            =   "Antecedente.frx":1F02
         PICH            =   "Antecedente.frx":21E4
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
         Left            =   7560
         TabIndex        =   72
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
         MICON           =   "Antecedente.frx":2435
         PICN            =   "Antecedente.frx":2451
         PICH            =   "Antecedente.frx":26E7
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
         Left            =   6840
         TabIndex        =   73
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
         MICON           =   "Antecedente.frx":2946
         PICN            =   "Antecedente.frx":2962
         PICH            =   "Antecedente.frx":2BF7
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
         TabIndex        =   75
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
         MICON           =   "Antecedente.frx":2E53
         PICN            =   "Antecedente.frx":2E6F
         PICH            =   "Antecedente.frx":3013
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label NoReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Antecedente 0 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   4560
         TabIndex        =   74
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Antecedentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   600
      Width           =   10935
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5280
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   15
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   17
         Top             =   3120
         Width           =   9495
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         Top             =   3960
         Width           =   8415
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   66
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47841281
         CurrentDate     =   39876
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   14
         Left            =   2400
         TabIndex        =   65
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S i   ó    No"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   64
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recidivas:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   63
         Top             =   2730
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Café"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   3840
         TabIndex        =   62
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nacidos Muertos"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   19
         Left            =   4320
         TabIndex        =   61
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nacidos Vivos"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   18
         Left            =   2640
         TabIndex        =   60
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Menopausia"
         Height          =   255
         Index           =   17
         Left            =   12720
         TabIndex        =   59
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cigarrillo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1200
         TabIndex        =   58
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S i   ó    No"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   7560
         TabIndex        =   57
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cual?"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   8760
         TabIndex        =   56
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gestas:"
         Height          =   195
         Index           =   3
         Left            =   7920
         TabIndex        =   55
         Top             =   930
         Width           =   540
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Edad 1° Emb"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   54
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aborto:"
         Height          =   195
         Index           =   1
         Left            =   5760
         TabIndex        =   53
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menopausia:"
         Height          =   195
         Index           =   29
         Left            =   5760
         TabIndex        =   52
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tratamiento Hormonal:"
         Height          =   435
         Index           =   28
         Left            =   120
         TabIndex        =   51
         Top             =   3930
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hábitos:"
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   50
         Top             =   1530
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclo:"
         Height          =   195
         Index           =   12
         Left            =   3960
         TabIndex        =   49
         Top             =   450
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sueño:"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   48
         Top             =   2010
         Width           =   510
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Antecedentes Familiares:"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   47
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flujo:"
         Height          =   195
         Index           =   16
         Left            =   2280
         TabIndex        =   46
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menarquía:"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   45
         Top             =   450
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embarazo:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   44
         Top             =   930
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alcohol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   6480
         TabIndex        =   43
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tranquilo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   1200
         TabIndex        =   42
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nº de Horas que Duerme"
         Height          =   495
         Index           =   24
         Left            =   3840
         TabIndex        =   41
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actividad Fisica"
         Height          =   495
         Index           =   25
         Left            =   6480
         TabIndex        =   40
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S i   ó    No"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   26
         Left            =   1200
         TabIndex        =   39
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Cual hormona ha tomado y por que tiempo?"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   27
         Left            =   2280
         TabIndex        =   38
         Top             =   3720
         Width           =   3165
      End
   End
   Begin VB.Frame FrameDato 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Quimioterapias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   10935
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   2160
         TabIndex        =   76
         Top             =   1680
         Width           =   8535
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   3120
         Width           =   8535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5040
         TabIndex        =   23
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47841281
         CurrentDate     =   39842
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2640
         Width           =   8535
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   2160
         Width           =   8535
      End
      Begin VB.TextBox Text20 
         Height          =   495
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   8535
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3120
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin ChamaleonButton.ChameleonBtn BtnProductosQuimio 
         Height          =   375
         Left            =   8280
         TabIndex        =   80
         Top             =   3960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Productos Quimioterapias"
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
         MICON           =   "Antecedente.frx":31B2
         PICN            =   "Antecedente.frx":31CE
         PICH            =   "Antecedente.frx":3466
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
         BackStyle       =   0  'Transparent
         Caption         =   "Esquema Quimiterapia:"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   1770
         Width           =   1620
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alergia a Medicamentos:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   3210
         Width           =   1740
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sueños"
         Height          =   15
         Left            =   360
         TabIndex        =   35
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enfermedades Asociadas:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   2250
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tratamiento Actual:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quimiterapia:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1290
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Si    ó    No"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclo:"
         Height          =   195
         Left            =   2640
         TabIndex        =   30
         Top             =   1290
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervenciones Quirúrgicas:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4440
         TabIndex        =   28
         Top             =   1290
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmAntecedentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsAntecedente As Recordset
Dim CSql As String
Dim Cambio
Dim NuevoReg
Dim IdAnt As String
Dim Reg_Actual(0 To 30) As String
Public IdPacA As String
Public IdLIdPacA As String

Private Sub BtnAgregar_Click()
On Error Resume Next

IdLIdInf = IdLDefault
IdAnt = ""

If IdPacA = "" Then MsgBox "Debe seleccionar un paciente!", vbExclamation + vbOKOnly, "Error": Exit Sub

NoReg(23) = "Antecedente NUEVO"
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
Call Blanqueo
NuevoReg = 1
End Sub

Private Sub BtnAnterior_Click()
On Error Resume Next
If BtnAgregar.Enabled = False Then BtnDesHacer_Click
If RsAntecedente.RecordCount <> 0 Then
    RsAntecedente.MovePrevious
    If RsAntecedente.BOF Then MsgBox "Ha llegado al Primer registro!", vbInformation + vbOKOnly, "Primer registro": RsAntecedente.MoveFirst
    carga_de_datos1
End If
End Sub

Private Sub BtnBuscar_Click()
BtnDesHacer_Click
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True

Blanqueo
NuevoReg = 0

CSql = "Select * From Antecedentes_Imp Where IdPaciente = " & IdPacA & " And ACTIVO =1"
Set RsAntecedente = CrearRS(CSql)

carga_de_datos1
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next

If IdPacA = "" Then MsgBox "Debe seleccionar un paciente!", vbExclamation + vbOKOnly, "Error": Exit Sub
If IdAnt = "" Then MsgBox "Debe seleccionar un antecedente!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Se procedera a Eliminar el registro del antecedente actual, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

CSql = "UPDATE Antecedentes_Imp set Activo=0 where idpaciente = " & IdPacA & " AND IdAntecedente=" & IdAnt & " AND IdL = '" & IdLIdInf & "'"
Set RsTemp = CrearRS(CSql)

MsgBox "El Antecedente ha sido Eliminado!", vbInformation + vbOKOnly, "Operacion Exitosa!"

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor Web"
EnviarRegPendiente IdAnt, IdLIdInf

Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "BORRAR", "Se Elimino el de ANTECEDENTES IdAntecedente es (" & IdAnt & ")")

BtnDesHacer_Click

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
Dim RsTemp As New ADODB.Recordset
Dim NuevoId As String

If IdPacA = "" Then MsgBox "Debe seleccionar un paciente!", vbExclamation + vbOKOnly, "Error": Exit Sub

'verifica si hay conexion al internet
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If

'crea el identificador automatico
CSql = "SELECT MAX(IdAntecedente)+1 as NuevoId FROM Antecedentes_Imp"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    NuevoId = RsTemp.Fields("NuevoId").Value
Else
    NuevoId = "1"
End If
Set RsTemp = Nothing

Select Case NuevoReg
Case Is = 0
    If Cambio = 1 Then
        If IdAnt = "" Then MsgBox "Debe haber un antecedente para poder guardar,"
        'CSql = "UPDATE Antecedentes_Imp set Activo=0 where idpaciente = " & IdPacA & " AND IdAntecedente=" & IdAnt & " AND IdL = '" & IdLIdInf & "'"
        'Set RsTemp = CrearRS(CSql)

        CSql = "Select * From Antecedentes_Imp where idpaciente = " & IdPacA & " AND IdAntecedente=" & IdAnt & " AND IdL = '" & IdLIdInf & "' And IdLIdPac='" & IdLIdPac & "'"
        Set RsTemp = CrearRS(CSql)
        
        'RsTemp.Fields("IdL").Value = IdLIdInf
        'RsTemp.Fields("Idpaciente").Value = IdPacA
        'RsTemp.Fields("IdLIdPac").Value = IdLIdPacA
        RsTemp.Fields("IdUsuario").Value = IdUser
        RsTemp.Fields("Menarquia").Value = Text1.Text
        RsTemp.Fields("Flujo").Value = Text2.Text
        RsTemp.Fields("Ciclo").Value = Text3.Text
        RsTemp.Fields("Menopausia").Value = Text4.Text
        RsTemp.Fields("Años").Value = Text5.Text
        RsTemp.Fields("Vivos").Value = Text6.Text
        RsTemp.Fields("Muertos").Value = Text7.Text
        RsTemp.Fields("Aborto").Value = Text8.Text
        RsTemp.Fields("Gesta").Value = Text9.Text
        RsTemp.Fields("Cigarrillo").Value = Text10.Text
        RsTemp.Fields("Cafe").Value = Text11.Text
        RsTemp.Fields("Alcohol").Value = Text12.Text
        RsTemp.Fields("Sueño").Value = Text13.Text
        RsTemp.Fields("Hora").Value = Text14.Text
        RsTemp.Fields("Actividad").Value = Text15.Text
        RsTemp.Fields("Cuales").Value = Text16.Text
        RsTemp.Fields("Familiares").Value = Text17.Text
        RsTemp.Fields("Hormonas").Value = Text18.Text
        RsTemp.Fields("Cual").Value = Text19.Text
        RsTemp.Fields("Intervenciones").Value = Text20.Text
        RsTemp.Fields("Quimi").Value = Text21.Text
        RsTemp.Fields("CicloQ").Value = Text22.Text
        RsTemp.Fields("FechaQ").Value = Format((DTPicker1.Value), "DD/MM/YYYY")
        RsTemp.Fields("Enfermedad_Aso").Value = Text23.Text
        RsTemp.Fields("Tratamiento").Value = Text24.Text
        RsTemp.Fields("Alergia").Value = Text25.Text
        RsTemp.Fields("Medico").Value = Text26.Text
        RsTemp.Fields("FechaR").Value = Format((DTPicker2.Value), "DD/MM/YYYY")
        RsTemp.Fields("Recidivas").Value = Text27.Text
        RsTemp.Fields("FechaAntecedentes").Value = Format(Now, "DD/MM/YYYY")
        RsTemp.Fields("Activo").Value = 1
        RsTemp.Update


        MsgBox "Registro Agregado Satisfactoriamente!", vbInformation + vbOKOnly, "Operacion Exitosa!"
        
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización de Antecedente"
        EnviarRegPendiente IdAnt, IdLIdInf
    Else
        MsgBox "No hay cambios que guardar", vbInformation + vbOKOnly, "No hay que guardar"
        Exit Sub
    End If

Case Is = 1
    If Cambio = 1 Then
    
'        CSql = "Insert into Antecedentes_Imp(IdAntecedente, Idpaciente,IdUsuario,Menarquia,Flujo,Ciclo,Menopausia,Años,Vivos" & _
'                ",Muertos,Aborto,Gesta,Cigarrillo,Cafe,Alcohol,Sueño,Hora,Actividad,Cuales,Familiares,Hormonas,Cual," & _
'                "Intervenciones,Quimi,CicloQ,FechaQ,Enfermedad_Aso,Tratamiento,Alergia,Medico,FechaR,Recidivas,FechaAntecedentes,Activo) VALUES " & _
'                "(" & NuevoId & ",'" & IdPacA & "','" & IdUser & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & _
'                Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & _
'                Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & _
'                Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & _
'                Text19.Text & "','" & Text20.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & _
'                Format((DTPicker1.Value), "DD/MM/YYYY") & "','" & Text23.Text & "','" & Text24.Text & "','" & Text25.Text & _
'                "','" & Text26.Text & "','" & Format((DTPicker2.Value), "DD/MM/YYYY") & "','" & Text27.Text & "','" & Format(Now, "DD/MM/YYYY") & "',1)"
'
'        Set RsTemp = CrearRS(CSql)


        IdLIdInf = NuevoIdL
        
        CSql = "Select * From Antecedentes_Imp"
        Set RsTemp = CrearRS(CSql)
        
        RsTemp.AddNew
        RsTemp.Fields("IdAntecedente").Value = NuevoId
        RsTemp.Fields("IdL").Value = IdLIdInf
        RsTemp.Fields("IdLIdPac").Value = IdLIdPacA
        RsTemp.Fields("Idpaciente").Value = IdPacA
        RsTemp.Fields("IdUsuario").Value = IdUser
        RsTemp.Fields("Menarquia").Value = Text1.Text
        RsTemp.Fields("Flujo").Value = Text2.Text
        RsTemp.Fields("Ciclo").Value = Text3.Text
        RsTemp.Fields("Menopausia").Value = Text4.Text
        RsTemp.Fields("Años").Value = Text5.Text
        RsTemp.Fields("Vivos").Value = Text6.Text
        RsTemp.Fields("Muertos").Value = Text7.Text
        RsTemp.Fields("Aborto").Value = Text8.Text
        RsTemp.Fields("Gesta").Value = Text9.Text
        RsTemp.Fields("Cigarrillo").Value = Text10.Text
        RsTemp.Fields("Cafe").Value = Text11.Text
        RsTemp.Fields("Alcohol").Value = Text12.Text
        RsTemp.Fields("Sueño").Value = Text13.Text
        RsTemp.Fields("Hora").Value = Text14.Text
        RsTemp.Fields("Actividad").Value = Text15.Text
        RsTemp.Fields("Cuales").Value = Text16.Text
        RsTemp.Fields("Familiares").Value = Text17.Text
        RsTemp.Fields("Hormonas").Value = Text18.Text
        RsTemp.Fields("Cual").Value = Text19.Text
        RsTemp.Fields("Intervenciones").Value = Text20.Text
        RsTemp.Fields("Quimi").Value = Text21.Text
        RsTemp.Fields("CicloQ").Value = Text22.Text
        RsTemp.Fields("FechaQ").Value = Format((DTPicker1.Value), "DD/MM/YYYY")
        RsTemp.Fields("Enfermedad_Aso").Value = Text23.Text
        RsTemp.Fields("Tratamiento").Value = Text24.Text
        RsTemp.Fields("Alergia").Value = Text25.Text
        RsTemp.Fields("Medico").Value = Text26.Text
        RsTemp.Fields("FechaR").Value = Format((DTPicker2.Value), "DD/MM/YYYY")
        RsTemp.Fields("Recidivas").Value = Text27.Text
        RsTemp.Fields("FechaAntecedentes").Value = Format(Now, "DD/MM/YYYY")
        RsTemp.Fields("Activo").Value = 1
        RsTemp.Update
        MsgBox "Registro Actualizado Satisfactoriamente!", vbInformation + vbOKOnly, "Operacion Exitosa!"
        
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización de Antecedente"
        EnviarRegPendiente NuevoId, IdLIdInf
    Else
        MsgBox "No hay cambios que agregar", vbInformation + vbOKOnly, "No hay que agregar"
        Exit Sub
    End If
End Select

If Cambio <> 0 Then
    If NuevoReg = 0 Then
    
        If Reg_Actual(0) <> Text1.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo MENARQUIA de (" & Reg_Actual(0) & ") a (" & Text1.Text & ")")
        If Reg_Actual(1) <> Text2.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Flujo de (" & Reg_Actual(1) & ") a (" & Text2.Text & ")")
        If Reg_Actual(2) <> Text3.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Ciclo de (" & Reg_Actual(2) & ") a (" & Text3.Text & ")")
        If Reg_Actual(3) <> Text4.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Menopausia de (" & Reg_Actual(3) & ") a (" & Text4.Text & ")")
        If Reg_Actual(4) <> Text5.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Años de (" & Reg_Actual(4) & ") a (" & Text5.Text & ")")
        If Reg_Actual(5) <> Text6.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Vivos de (" & Reg_Actual(5) & ") a (" & Text6.Text & ")")
        If Reg_Actual(6) <> Text7.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Muertos de (" & Reg_Actual(6) & ") a (" & Text7.Text & ")")
        If Reg_Actual(7) <> Text8.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Aborto de (" & Reg_Actual(7) & ") a (" & Text8.Text & ")")
        If Reg_Actual(8) <> Text9.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Gesta de (" & Reg_Actual(8) & ") a (" & Text9.Text & ")")
        If Reg_Actual(9) <> Text10.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Cigarrillo de (" & Reg_Actual(9) & ") a (" & Text10.Text & ")")
        If Reg_Actual(10) <> Text11.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Cafe de (" & Reg_Actual(10) & ") a (" & Text11.Text & ")")
        If Reg_Actual(11) <> Text12.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Alcohol de (" & Reg_Actual(11) & ") a (" & Text12.Text & ")")
        If Reg_Actual(12) <> Text13.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Sueño de (" & Reg_Actual(12) & ") a (" & Text13.Text & ")")
        If Reg_Actual(13) <> Text14.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Hora de (" & Reg_Actual(13) & ") a (" & Text14.Text & ")")
        If Reg_Actual(14) <> Text15.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Actividad de (" & Reg_Actual(14) & ") a (" & Text15.Text & ")")
        If Reg_Actual(15) <> Text16.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Cuales de (" & Reg_Actual(15) & ") a (" & Text16.Text & ")")
        If Reg_Actual(16) <> Text17.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Familiares de (" & Reg_Actual(16) & ") a (" & Text17.Text & ")")
        If Reg_Actual(17) <> Text18.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Hormonas de (" & Reg_Actual(17) & ") a (" & Text18.Text & ")")
        If Reg_Actual(18) <> Text19.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Cual de (" & Reg_Actual(18) & ") a (" & Text19.Text & ")")
        If Reg_Actual(19) <> Text20.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Intervenciones de (" & Reg_Actual(19) & ") a (" & Text20.Text & ")")
        If Reg_Actual(20) <> Text21.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Quimi de (" & Reg_Actual(20) & ") a (" & Text21.Text & ")")
        If Reg_Actual(21) <> Text22.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Ciclo de (" & Reg_Actual(21) & ") a (" & Text22.Text & ")")
        If Reg_Actual(22) <> Text23.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Enfermedad_Aso de (" & Reg_Actual(22) & ") a (" & Text23.Text & ")")
        If Reg_Actual(23) <> Text24.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Tratamiento de (" & Reg_Actual(23) & ") a (" & Text24.Text & ")")
        If Reg_Actual(24) <> Text25.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Alergia de (" & Reg_Actual(24) & ") a (" & Text25.Text & ")")
        If Reg_Actual(25) <> Text26.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Medico de (" & Reg_Actual(25) & ") a (" & Text26.Text & ")")
        If Reg_Actual(26) <> Text27.Text Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo Recidivas de (" & Reg_Actual(26) & ") a (" & Text27.Text & ")")
        If Reg_Actual(27) <> Format(DTPicker1.Value, "DD/MM/YYYY") Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo FechaQ de (" & Reg_Actual(27) & ") a (" & Format(DTPicker1.Value, "DD/MM/YYYY") & ")")
        If Reg_Actual(28) <> Format(DTPicker2.Value, "DD/MM/YYYY") Then Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "MODIFICAR", "Se modifico el campo FechaR de (" & Reg_Actual(28) & ") a (" & Format(DTPicker2.Value, "DD/MM/YYYY") & ")")
    Else
        Call Enviar_Bitacora(IdUser, "Nutricion-ANTECEDENTES", "INGRESAR", "Se ingreso un nuevo registro cuya IdAntecedente es (" & NuevoId & ")")
    End If
End If

Cambio = 0
BtnDesHacer_Click
End Sub

Sub EnviarRegPendiente(ByVal IdAnt As Integer, ByVal IdLIdPac2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "SELECT * FROM Antecedentes_Imp WHERE IdAntecedente = " & IdAnt & " AND IdL = '" & IdLIdPac2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Antecedentes_Imp (["
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
RsRegPendiente.Fields("Modulo").Value = "Antecedentes"
RsRegPendiente.Fields("Tabla").Value = "Antecedentes_Imp"
RsRegPendiente.Fields("Condicional").Value = "IdAntecedente = " & IdAnt & " AND IdL = '" & IdLIdPac2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

Private Sub BtnProductosQuimio_Click()
FrmProductosQuimio.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnSiguiente_Click()
On Error Resume Next
If BtnAgregar.Enabled = False Then BtnDesHacer_Click
If RsAntecedente.RecordCount <> 0 Then
    RsAntecedente.MoveNext
    If RsAntecedente.EOF Then MsgBox "Ha llegado al Ultimo registro!", vbInformation + vbOKOnly, "Ultimo registro": RsAntecedente.MoveLast
    carga_de_datos1
End If

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text26.SetFocus
        Case 37
            Text22.SetFocus
        Case 38
            Text20.SetFocus
        Case 39
            Text26.SetFocus
        Case 40
            Text23.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text17.SetFocus
        Case 37
            Text27.SetFocus
        Case 38
            Text13.SetFocus
        Case 40
            Text17.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Me.Caption = "Antecedentes y Quimioterapias - Paciente: " & IdPacA
Centrar Me

DTPicker1.Value = Now()
DTPicker2.Value = Now()

If Trim(IdPacA) <> "" Then
    CSql = "Select * From Antecedentes_Imp Where IdPaciente = " & IdPacA & " And ACTIVO =1"
    Set RsAntecedente = CrearRS(CSql)
    
    Call carga_de_datos1
End If
End Sub
Sub carga_de_datos1()

IdLIdInf = IdLDefault
IdAnt = ""
If RsAntecedente.RecordCount = 0 Then GoTo nohay

IdAnt = RsAntecedente.Fields("IdAntecedente").Value
IdLIdInf = RsAntecedente.Fields("IdL").Value

If RsAntecedente.Fields("Menarquia") <> "" Then Text1.Text = RsAntecedente.Fields("Menarquia") Else Text1.Text = ""
If RsAntecedente.Fields("Flujo") <> "" Then Text2.Text = RsAntecedente.Fields("Flujo") Else Text2.Text = ""
If RsAntecedente.Fields("Ciclo") <> "" Then Text3.Text = RsAntecedente.Fields("Ciclo") Else Text3.Text = ""
If RsAntecedente.Fields("Menopausia") <> "" Then Text4.Text = RsAntecedente.Fields("Menopausia") Else Text4.Text = ""
If RsAntecedente.Fields("Años") <> "" Then Text5.Text = RsAntecedente.Fields("Años") Else Text5.Text = ""
If RsAntecedente.Fields("Vivos") <> "" Then Text6.Text = RsAntecedente.Fields("Vivos") Else Text6.Text = ""
If RsAntecedente.Fields("Muertos") <> "" Then Text7.Text = RsAntecedente.Fields("Muertos") Else Text7.Text = ""
If RsAntecedente.Fields("Aborto") <> "" Then Text8.Text = RsAntecedente.Fields("Aborto") Else Text8.Text = ""
If RsAntecedente.Fields("Gesta") <> "" Then Text9.Text = RsAntecedente.Fields("Gesta") Else Text9.Text = ""
If RsAntecedente.Fields("Cigarrillo") <> "" Then Text10.Text = RsAntecedente.Fields("Cigarrillo") Else Text10.Text = ""
If RsAntecedente.Fields("Cafe") <> "" Then Text11.Text = RsAntecedente.Fields("Cafe") Else Text11.Text = ""
If RsAntecedente.Fields("Alcohol") <> "" Then Text12.Text = RsAntecedente.Fields("Alcohol") Else Text12.Text = ""
If RsAntecedente.Fields("Sueño") <> "" Then Text13.Text = RsAntecedente.Fields("Sueño") Else Text13.Text = ""
If RsAntecedente.Fields("Hora") <> "" Then Text14.Text = RsAntecedente.Fields("Hora") Else Text14.Text = ""
If RsAntecedente.Fields("Actividad") <> "" Then Text15.Text = RsAntecedente.Fields("Actividad") Else Text15.Text = ""
If RsAntecedente.Fields("Cuales") <> "" Then Text16.Text = RsAntecedente.Fields("Cuales") Else Text16.Text = ""
If RsAntecedente.Fields("Familiares") <> "" Then Text17.Text = RsAntecedente.Fields("Familiares") Else Text17.Text = ""
If RsAntecedente.Fields("Hormonas") <> "" Then Text18.Text = RsAntecedente.Fields("Hormonas") Else Text18.Text = ""
If RsAntecedente.Fields("Cual") <> "" Then Text19.Text = RsAntecedente.Fields("Cual") Else Text19.Text = ""
If RsAntecedente.Fields("Intervenciones") <> "" Then Text20.Text = RsAntecedente.Fields("Intervenciones") Else Text20.Text = ""
If RsAntecedente.Fields("Quimi") <> "" Then Text21.Text = RsAntecedente.Fields("Quimi") Else Text21.Text = ""
If RsAntecedente.Fields("Ciclo") <> "" Then Text22.Text = RsAntecedente.Fields("Ciclo") Else Text22.Text = ""
If RsAntecedente.Fields("Enfermedad_Aso") <> "" Then Text23.Text = RsAntecedente.Fields("Enfermedad_Aso") Else Text23.Text = ""
If RsAntecedente.Fields("Tratamiento") <> "" Then Text24.Text = RsAntecedente.Fields("Tratamiento") Else Text24.Text = ""
If RsAntecedente.Fields("Alergia") <> "" Then Text25.Text = RsAntecedente.Fields("Alergia") Else Text25.Text = ""
If RsAntecedente.Fields("Medico") <> "" Then Text26.Text = RsAntecedente.Fields("Medico") Else Text26.Text = ""
If RsAntecedente.Fields("Recidivas") <> "" Then Text27.Text = RsAntecedente.Fields("Recidivas") Else Text27.Text = ""
DTPicker1.Value = RsAntecedente.Fields("FechaQ")
DTPicker2.Value = RsAntecedente.Fields("FechaR")

Reg_Actual(0) = RsAntecedente.Fields("Menarquia").Value
Reg_Actual(1) = RsAntecedente.Fields("Flujo").Value
Reg_Actual(2) = RsAntecedente.Fields("Ciclo").Value
Reg_Actual(3) = RsAntecedente.Fields("Menopausia").Value
Reg_Actual(4) = RsAntecedente.Fields("Años").Value
Reg_Actual(5) = RsAntecedente.Fields("Vivos").Value
Reg_Actual(6) = RsAntecedente.Fields("Muertos").Value
Reg_Actual(7) = RsAntecedente.Fields("Aborto").Value
Reg_Actual(8) = RsAntecedente.Fields("Gesta").Value
Reg_Actual(9) = RsAntecedente.Fields("Cigarrillo").Value
Reg_Actual(10) = RsAntecedente.Fields("Cafe").Value
Reg_Actual(11) = RsAntecedente.Fields("Alcohol").Value
Reg_Actual(12) = RsAntecedente.Fields("Sueño").Value
Reg_Actual(13) = RsAntecedente.Fields("Hora").Value
Reg_Actual(14) = RsAntecedente.Fields("Actividad").Value
Reg_Actual(15) = RsAntecedente.Fields("Cuales").Value
Reg_Actual(16) = RsAntecedente.Fields("Familiares").Value
Reg_Actual(17) = RsAntecedente.Fields("Hormonas").Value
Reg_Actual(18) = RsAntecedente.Fields("Cual").Value
Reg_Actual(19) = RsAntecedente.Fields("Intervenciones").Value
Reg_Actual(20) = RsAntecedente.Fields("Quimi").Value
Reg_Actual(21) = RsAntecedente.Fields("Ciclo").Value
Reg_Actual(22) = RsAntecedente.Fields("Enfermedad_Aso").Value
Reg_Actual(23) = RsAntecedente.Fields("Tratamiento").Value
Reg_Actual(24) = RsAntecedente.Fields("Alergia").Value
Reg_Actual(25) = RsAntecedente.Fields("Medico").Value
Reg_Actual(26) = RsAntecedente.Fields("Recidivas").Value
Reg_Actual(27) = RsAntecedente.Fields("FechaQ")
Reg_Actual(28) = RsAntecedente.Fields("FechaR")

NoReg.Item(23) = "Antecedente " & RsAntecedente.AbsolutePosition & " / " & RsAntecedente.RecordCount
Cambio = 0: NuevoReg = 0
Exit Sub

nohay:
    NoReg.Item(23) = "Antecedente 0 / 0"
    IdAnt = ""
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
    'Text26.Text = ""
    Text27.Text = ""
    NuevoReg = 1
    Cambio = 0
             
MsgBox "No hay Datos asociados al paciente: " & Chr(13) & Chr(13) & Text3.Text & " " & Text4.Text & Chr(13) & "Se Mostraran los datos de Historia Médica en blanco" & Chr(13), vbInformation + vbOKOnly, "No Tiene Datos"
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
BtnEliminar.Enabled = False
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
End Sub
     
Private Sub Form_Unload(Cancel As Integer)
If Cambio <> 0 Then
Msg = "Ha hecho modificaciones al registro actual desea guardar estos cambios?"
d = MsgBox(Msg, vbYesNo, "Guardar cambios")
If d = vbYes Then Call BtnGuardarActualizar_Click
End If
End Sub

Private Sub Option1_Click(Index As Integer)
For i = 0 To Option1.Count - 1
    If i <> Index Then
        Option1(i).Value = False
        FrameDato(i).Visible = False
        'FrmGrl(i).Visible = False
    Else
        Option1(i).Value = True
        FrameDato(i).Visible = True
        'FrmGrl(i).Visible = True
    End If
Next
End Sub

Private Sub Text1_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text1.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text1.Text)
    pru = LCase(Mid(Text1.Text, i, 1))
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

Text1.Text = StrText
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case 39
            Text2.SetFocus
        Case 40
            Text5.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case 38
            Text6.SetFocus
        Case 39
            Text11.SetFocus
        Case 40
            Text13.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text12.SetFocus
        Case 37
            Text10.SetFocus
        Case 38
            Text7.SetFocus
        Case 39
            Text12.SetFocus
        Case 40
            Text14.SetFocus
    End Select
End If
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case 37
            Text11.SetFocus
        Case 38
            Text9.SetFocus
        Case 39
            BtnFacturar.SetFocus
        Case 40
            Text15.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case 38
            Text10.SetFocus
        Case 39
            Text14.SetFocus
        Case 40
            DTPicker2.SetFocus
    End Select
End If
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text15.SetFocus
        Case 37
            Text13.SetFocus
        Case 38
            Text11.SetFocus
        Case 39
            Text15.SetFocus
        Case 40
            DTPicker2.SetFocus
    End Select
End If
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text16.SetFocus
        Case 37
            Text14.SetFocus
        Case 38
            Text12.SetFocus
        Case 39
            Text16.SetFocus
        Case 40
            DTPicker2.SetFocus
    End Select
End If
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text27.SetFocus
        Case 37
            Text15.SetFocus
        Case 38
            Text12.SetFocus
        Case 40
            DTPicker2.SetFocus
    End Select
End If
End Sub

Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text18.SetFocus
        Case 38
            Text27.SetFocus
        Case 40
            Text18.SetFocus
    End Select
End If
End Sub

Private Sub Text18_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text19.SetFocus
        Case 38
            Text17.SetFocus
        Case 39
            Text19.SetFocus
        Case 40
            Text20.SetFocus
    End Select
End If
End Sub

Private Sub Text19_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text20.SetFocus
        Case 37
            Text18.SetFocus
        Case 38
            Text17.SetFocus
        Case 40
            Text20.SetFocus
    End Select
End If
End Sub

Private Sub Text2_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text2.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text2.Text)
    pru = LCase(Mid(Text2.Text, i, 1))
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

Text2.Text = StrText
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case 37
            Text1.SetFocus
        Case 39
            Text3.SetFocus
        Case 40
            Text6.SetFocus
    End Select
End If
End Sub

Private Sub Text20_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text21.SetFocus
        Case 38
            Text18.SetFocus
        Case 40
            Text21.SetFocus
    End Select
End If
End Sub

Private Sub Text21_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text22.SetFocus
        Case 38
            Text20.SetFocus
        Case 39
            Text22.SetFocus
        Case 40
            Text23.SetFocus
    End Select
End If
End Sub

Private Sub Text22_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPicker1.SetFocus
        Case 37
            Text21.SetFocus
        Case 38
            Text20.SetFocus
        Case 39
            DTPicker1.SetFocus
        Case 40
            Text23.SetFocus
    End Select
End If
End Sub

Private Sub Text23_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text24.SetFocus
        Case 38
            Text21.SetFocus
        Case 40
            Text24.SetFocus
    End Select
End If
End Sub

Private Sub Text24_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text25.SetFocus
        Case 38
            Text23.SetFocus
        Case 40
            Text25.SetFocus
    End Select
End If
End Sub

Private Sub Text25_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case 38
            Text24.SetFocus
        Case 40
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text23.SetFocus
        Case 37
            DTPicker1.SetFocus
        Case 38
            Text20.SetFocus
        Case 40
            Text23.SetFocus
    End Select
End If
End Sub

Private Sub Text27_Change()
Cambio = 1
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

Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPicker2.SetFocus
        Case 38
            Text13.SetFocus
        Case 39
            DTPicker2.SetFocus
        Case 40
            Text17.SetFocus
    End Select
End If
End Sub

Private Sub Text3_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text3.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text3.Text)
    pru = LCase(Mid(Text3.Text, i, 1))
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

Text3.Text = StrText
Text3.SelStart = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case 37
            Text2.SetFocus
        Case 39
            Text4.SetFocus
        Case 40
            Text7.SetFocus
    End Select
End If
End Sub

Private Sub Text4_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text4.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text4.Text)
    pru = LCase(Mid(Text4.Text, i, 1))
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

Text4.Text = StrText
Text4.SelStart = Len(Text4.Text)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text5.SetFocus
        Case 37
            Text3.SetFocus
        Case 39
            BtnAyuda.SetFocus
        Case 40
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub Text5_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text5.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text5.Text)
    pru = LCase(Mid(Text5.Text, i, 1))
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

Text5.Text = StrText
Text5.SelStart = Len(Text5.Text)
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case 38
            Text1.SetFocus
        Case 39
            Text6.SetFocus
        Case 40
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text6_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text6.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text6.Text)
    pru = LCase(Mid(Text6.Text, i, 1))
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

Text6.Text = StrText
Text6.SelStart = Len(Text6.Text)
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text7.SetFocus
        Case 37
            Text5.SetFocus
        Case 38
            Text2.SetFocus
        Case 39
            Text7.SetFocus
        Case 40
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text7_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text7.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text7.Text)
    pru = LCase(Mid(Text7.Text, i, 1))
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

Text7.Text = StrText
Text7.SelStart = Len(Text7.Text)
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text8.SetFocus
        Case 37
            Text6.SetFocus
        Case 38
            Text3.SetFocus
        Case 39
            Text8.SetFocus
        Case 40
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text8_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text8.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text8.Text)
    pru = LCase(Mid(Text8.Text, i, 1))
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

Text8.Text = StrText
Text8.SelStart = Len(Text8.Text)
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case 37
            Text7.SetFocus
        Case 38
            Text4.SetFocus
        Case 39
            Text9.SetFocus
        Case 40
            Text12.SetFocus
    End Select
End If
End Sub

Private Sub Text9_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text9.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text9.Text)
'    pru = LCase(Mid(Text9.Text, i, 1))
'    If pru Like " " Then
'        T = 1
'        StrText = StrText & " "
'    Else
'        If T = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            T = 0
'        End If
'End If
'
'Next i
'
'Text9.Text = StrText
'Text9.SelStart = Len(Text9.Text)
End Sub

Private Sub Text10_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text10.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text10.Text)
    pru = LCase(Mid(Text10.Text, i, 1))
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

Text10.Text = StrText
Text10.SelStart = Len(Text10.Text)
End Sub

Private Sub Text11_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text11.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text11.Text)
    pru = LCase(Mid(Text11.Text, i, 1))
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

Text11.Text = StrText
Text11.SelStart = Len(Text11.Text)
End Sub
Private Sub Text12_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text12.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text12.Text)
    pru = LCase(Mid(Text12.Text, i, 1))
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

Text12.Text = StrText
Text12.SelStart = Len(Text12.Text)
End Sub

Private Sub Text13_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text13.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text13.Text)
    pru = LCase(Mid(Text13.Text, i, 1))
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

Text13.Text = StrText
Text13.SelStart = Len(Text13.Text)
End Sub

Private Sub Text14_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text14.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text14.Text)
    pru = LCase(Mid(Text14.Text, i, 1))
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

Text14.Text = StrText
Text14.SelStart = Len(Text14.Text)
End Sub

Private Sub Text15_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text15.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text15.Text)
    pru = LCase(Mid(Text15.Text, i, 1))
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

Text15.Text = StrText
Text15.SelStart = Len(Text15.Text)
End Sub

Private Sub Text16_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text16.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text16.Text)
    pru = LCase(Mid(Text16.Text, i, 1))
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

Text16.Text = StrText
Text16.SelStart = Len(Text16.Text)
End Sub

Private Sub Text17_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text17.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text17.Text)
    pru = LCase(Mid(Text17.Text, i, 1))
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

Text17.Text = StrText
Text17.SelStart = Len(Text17.Text)
End Sub

Private Sub Text18_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text18.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text18.Text)
    pru = LCase(Mid(Text18.Text, i, 1))
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

Text18.Text = StrText
Text18.SelStart = Len(Text18.Text)
End Sub

Private Sub Text19_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text19.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text19.Text)
    pru = LCase(Mid(Text19.Text, i, 1))
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

Text19.Text = StrText
Text19.SelStart = Len(Text19.Text)
End Sub

Private Sub Text20_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text20.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text20.Text)
'    pru = LCase(Mid(Text20.Text, i, 1))
'    If pru Like " " Then
'        T = 1
'        StrText = StrText & " "
'    Else
'        If T = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            T = 0
'        End If
'   End If
'Next i
'
'Text20.Text = StrText
'Text20.SelStart = Len(Text20.Text)
End Sub

Private Sub Text21_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text21.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text21.Text)
    pru = LCase(Mid(Text21.Text, i, 1))
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

Text21.Text = StrText
Text21.SelStart = Len(Text21.Text)
End Sub

Private Sub Text22_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text22.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text22.Text)
    pru = LCase(Mid(Text22.Text, i, 1))
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

Text22.Text = StrText
Text22.SelStart = Len(Text22.Text)
End Sub
Private Sub Text23_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text23.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text23.Text)
    pru = LCase(Mid(Text23.Text, i, 1))
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

Text23.Text = StrText
Text23.SelStart = Len(Text23.Text)
End Sub

Private Sub Text24_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text24.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text24.Text)
'    pru = LCase(Mid(Text24.Text, i, 1))
'    If pru Like " " Then
'        T = 1
'        StrText = StrText & " "
'    Else
'        If T = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            T = 0
'        End If
'    End If
'Next i
'
'Text24.Text = StrText
'Text24.SelStart = Len(Text24.Text)
End Sub

Private Sub Text25_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text25.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text25.Text)
    pru = LCase(Mid(Text25.Text, i, 1))
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

Text25.Text = StrText
Text25.SelStart = Len(Text25.Text)
End Sub

Private Sub Text26_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text26.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text26.Text)
'    pru = LCase(Mid(Text26.Text, i, 1))
'    If pru Like " " Then
'        T = 1
'        StrText = StrText & " "
'    Else
'        If T = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            T = 0
'        End If
'    End If
'Next i
'
'Text26.Text = StrText
'Text26.SelStart = Len(Text26.Text)
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case 37
            Text8.SetFocus
        Case 38
            BtnAyuda.SetFocus
        Case 40
            Text12.SetFocus
    End Select
End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDireccionMedica 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dirección Médica"
   ClientHeight    =   8835
   ClientLeft      =   3345
   ClientTop       =   1395
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Internista.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   13200
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   4080
      TabIndex        =   54
      Top             =   8040
      Width           =   9015
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7920
         TabIndex        =   29
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
         MICON           =   "Internista.frx":1002
         PICN            =   "Internista.frx":101E
         PICH            =   "Internista.frx":11E7
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
         TabIndex        =   24
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
         MICON           =   "Internista.frx":141C
         PICN            =   "Internista.frx":1438
         PICH            =   "Internista.frx":16C7
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
         TabIndex        =   23
         ToolTipText     =   "Agregar"
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
         MICON           =   "Internista.frx":1B08
         PICN            =   "Internista.frx":1B24
         PICH            =   "Internista.frx":1CB1
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
         Left            =   6720
         TabIndex        =   28
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
         MICON           =   "Internista.frx":1EE6
         PICN            =   "Internista.frx":1F02
         PICH            =   "Internista.frx":21E4
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
         Height          =   375
         Left            =   6000
         TabIndex        =   27
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
         MICON           =   "Internista.frx":2435
         PICN            =   "Internista.frx":2451
         PICH            =   "Internista.frx":26E7
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
         Left            =   5400
         TabIndex        =   26
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
         MICON           =   "Internista.frx":2946
         PICN            =   "Internista.frx":2962
         PICH            =   "Internista.frx":2BF7
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
         Left            =   3840
         TabIndex        =   25
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "Internista.frx":2E53
         PICN            =   "Internista.frx":2E6F
         PICH            =   "Internista.frx":2F94
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
         Left            =   2400
         TabIndex        =   58
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
         MICON           =   "Internista.frx":3224
         PICN            =   "Internista.frx":3240
         PICH            =   "Internista.frx":33E4
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   53
      Top             =   8040
      Width           =   3855
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
         TabIndex        =   30
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad"
         Top             =   240
         Width           =   2175
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         ToolTipText     =   "Buscar"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Buscar"
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
         MICON           =   "Internista.frx":3824
         PICN            =   "Internista.frx":3840
         PICH            =   "Internista.frx":3AA5
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Examen Físico de Ingreso "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   43
      Top             =   4200
      Width           =   12975
      Begin ChamaleonButton.ChameleonBtn BtnAntecedentes 
         Height          =   375
         Left            =   6480
         TabIndex        =   22
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Antecedentes"
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
         MICON           =   "Internista.frx":3D37
         PICN            =   "Internista.frx":3D53
         PICH            =   "Internista.frx":3FE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnExamenesLaboratorio 
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Examen de Laboratorio"
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
         MICON           =   "Internista.frx":426F
         PICN            =   "Internista.frx":428B
         PICH            =   "Internista.frx":451A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   9000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   7
         Top             =   2760
         Width           =   7695
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   6
         Top             =   2280
         Width           =   7695
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   7695
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   4
         Top             =   1320
         Width           =   7695
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   7815
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   7695
      End
      Begin ChamaleonButton.ChameleonBtn BtnEvolucionOncologica 
         Height          =   375
         Left            =   2280
         TabIndex        =   59
         Top             =   3240
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
         MICON           =   "Internista.frx":47B6
         PICN            =   "Internista.frx":47D2
         PICH            =   "Internista.frx":4A6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signos Vitales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   465
         Width           =   1035
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abdomen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neurológico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   2865
         Width           =   915
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tórax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   1905
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabello:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1425
         Width           =   570
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Piel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión de Sistemas:"
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
         Left            =   9000
         TabIndex        =   44
         Top             =   120
         Width           =   1890
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Historia Clinica Integral"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   41
      Top             =   2400
      Width           =   12975
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   12615
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   12615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enfermedad Actúal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnóstico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   885
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   720
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1200
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Paciente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   12975
      Begin VB.ComboBox CboSexo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpFechaActual 
         Height          =   375
         Left            =   9240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51773441
         CurrentDate     =   39813
      End
      Begin MSComCtl2.DTPicker DtpFechaNac 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51773441
         CurrentDate     =   39813
      End
      Begin MSComCtl2.DTPicker DtpFechaRegistro 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51773441
         CurrentDate     =   39813
      End
      Begin ChamaleonButton.ChameleonBtn BtnLlamar 
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         ToolTipText     =   "Llamar"
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "Internista.frx":4CF1
         PICN            =   "Internista.frx":4D0D
         PICH            =   "Internista.frx":4FA9
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
         Left            =   3360
         TabIndex        =   19
         ToolTipText     =   "Lista de Espera"
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "Internista.frx":51DE
         PICN            =   "Internista.frx":51FA
         PICH            =   "Internista.frx":5483
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
         Left            =   6600
         TabIndex        =   57
         ToolTipText     =   "Desocupar al Paciente Atendido"
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
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
         MICON           =   "Internista.frx":571B
         PICN            =   "Internista.frx":5737
         PICH            =   "Internista.frx":58DB
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
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro 0 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10920
         TabIndex        =   56
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Historia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   55
         Top             =   450
         Width           =   870
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   6360
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Actual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8160
         TabIndex        =   51
         Top             =   450
         Width           =   990
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1560
         Left            =   10800
         Picture         =   "Internista.frx":5B10
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6600
         TabIndex        =   40
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   39
         Top             =   1380
         Width           =   405
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3075
         TabIndex        =   38
         Top             =   1410
         Width           =   420
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nac.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1410
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellido(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5520
         TabIndex        =   35
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2505
         TabIndex        =   34
         Top             =   450
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   420
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmDireccionMedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPaciente As New ADODB.Recordset
Dim RsInternista As New ADODB.Recordset
Dim bd1 As New ADODB.Recordset
Dim Cambio
Dim RegNew
Dim IdInter As String
Dim NuevoId As String

Private Sub BtnAgregar_Click()
On Error Resume Next
IO = 1

'Call blanqueo1
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
BtnGuardarActualizar.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Frame2.BackColor = &HE0E0E0
Frame3.BackColor = &HE0E0E0
End Sub

Private Sub BtnAntecedentes_Click()
On Error Resume Next
FrmAntecedentes.Show vbModal
End Sub

Private Sub BtnAnterior_Click()
On Error Resume Next
If RsPaciente.RecordCount <> 0 Then
    RsPaciente.MovePrevious
    If RsPaciente.BOF Then RsPaciente.MoveLast
    Call Carga_De_Datos
    Call carga_de_datos1
    Cambio = 0
Else
    MsgBox "No hay registros cargados, Inicie una nueva Busqueda!", vbExclamation + vbOKOnly, "Error"
    IdPac1 = ""
End If
End Sub

Public Sub BtnBuscar_Click()
On Error Resume Next
If TxtBuscar.Text = "Busqueda" Or TxtBuscar.Text = "" Then
    CSql = "select * from Paciente"
Else
    CSql = "Select * From Paciente Where CedulaP = " & Val(TxtBuscar.Text) & " or NombreP like '%" & TxtBuscar.Text & "%'"
End If
Set RsPaciente = CrearRS(CSql)

If RsPaciente.RecordCount = 0 Then
    MsgBox "No Existe el registro", vbInformation + vbOKOnly, "No hay datos"
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text14.Text = ""
    Text6.Text = ""
    Label12 = ""
    DtpFechaRegistro.Value = Now
    DtpFechaNac.Value = Now
    DtpFechaActual.Value = Now
    Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    CboSexo.ListIndex = -1
    NoReg = "Registro 0 / 0"
    IdPac1 = ""
    Call Carga_De_Datos
    Exit Sub
End If

Call Carga_De_Datos
Call carga_de_datos1
Cambio = 0
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
Call blanqueo1
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
BtnImprimir.Enabled = True
BtnEliminar.Enabled = True
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False

Frame2.Enabled = False
Frame3.Enabled = False
Call Carga_De_Datos
Call carga_de_datos1
Frame2.BackColor = &HEAEFEF
Frame3.BackColor = &HEAEFEF
Cambio = 0
End Sub

Private Sub BtnDesocuparAlPacienteAtendido_Click()
Dim bdlista88 As New ADODB.Recordset
CSql = "Delete from ubi_paciente where modul = " & ModulO
Set bdlista88 = CrearRS(CSql)
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next
resp = MsgBox("se va a eliminar el registro actual, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub


CSql = "Update Internista set Activo=2 Where IdInternista=" & IdInter
Set RsTemp = CrearRS(CSql)

Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "BORRAR", "Se elimino de la tabla INTERNISTA el registro de Id=" & IdInter & " Del paciente=" & IdPac1)
MsgBox "Registro Eliminado!", vbInformation + vbOKOnly, "Operacion Exitosa!"
'BorrarHosting
BtnDesHacer_Click
End Sub

Sub BorrarHosting()
On Error GoTo salir
ConectarHosting

CSql = "Update Internista set Activo=2 Where IdInternista=" & IdInter
Set RsWeb = CrearRsWeb(CSql)
    
salir:
If WebCnn.State = 0 Then
    BorrarHosting
Else
    GoTo f
End If

f:
If Err.Number <> 0 Then MsgBox Err.Description
If WebCnn.State = 1 Then WebCnn.Close Else Exit Sub
End Sub

Private Sub BtnEvolucionOncologica_Click()
On Error Resume Next
Especia = "Internista"
FrmEvolucion.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnExamenesLaboratorio_Click()
On Error Resume Next
FrmExamenHematologico.Show
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
If Cambio = 0 Then MsgBox "No se han realizado cambios!", vbInformation + vbOKOnly, "Informacion": Exit Sub
If IdInter = "" And RegNew = 0 Then MsgBox "Debe seleccionar o agregar un registro para poder guardar los cambios!", vbExclamation + vbOKOnly, "Error": Exit Sub
If Text8.Text = "" Then
    f = "Diagnostico"
    GoTo noguardA
ElseIf Text9.Text = "" Then
    f = "Enfermedad"
    GoTo noguardA
ElseIf Text10.Text = "" Then
    f = "Menarquia"
    GoTo noguardA
ElseIf Text11.Text = "" Then
    f = "Flujo"
    GoTo noguardA
ElseIf Text13.Text = "" Then
    f = "Cabello"
    GoTo noguardA
ElseIf Text15.Text = "" Then
    f = "Torax"
    GoTo noguardA
ElseIf Text16.Text = "" Then
    f = "Abdomen"
    GoTo noguardA
ElseIf Text17.Text = "" Then
    f = "Neurólogico"
    GoTo noguardA
ElseIf Text18.Text = "" Then
    f = "Revision de Sistemas"
    GoTo noguardA
End If

CSql = "SELECT MAX(IdInternista)+1 as NuevoId FROM Internista"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId")) Then
    NuevoId = RsTemp.Fields("NuevoId")
Else
    NuevoId = "0"
End If

Select Case RegNew
    
    Case Is = 0       'actualiza
        
        If Cambio = 1 Then
            'hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos
            CSql = "Update Internista set Activo=0 Where Activo=1 and IdPaciente=" & IdPac1 & " and idinternista='" & IdInter & "'"
            Set RsTemp = CrearRS(CSql)
            
'            CSql = "Insert into Internista(idinternista,Idpaciente,IdUsuario,Enfernedad_Act,Diagnostico,Signos,ColorP," & _
'                "Cabello,Torax,Abdomen,Neurologico,Revision,Activo) VALUES (" & NuevoId & "," & IdPac1 & "," & IdUser & ",'" & _
'                Text9.Text & "','" & Text8.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & _
'                Text13.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & _
'                Text18.Text & "',1)"
                
                
             CSql = "Update Internista set idinternista=" & NuevoId & ",Idpaciente=" & IdPac1 & ",IdUsuario=" & IdUser & ",Enfernedad_Act='" & Text9.Text & "',Diagnostico='" & Text8.Text & "',Signos='" & Text10.Text & "',ColorP='" & Text11.Text & "'," & _
                "Cabello='" & Text13.Text & "',Torax='" & Text15.Text & "',Abdomen='" & Text16.Text & "',Neurologico='" & Text17.Text & "',Revision='" & Text18.Text & "',Activo='1' Where Activo='1' and IdPaciente=" & IdPac1 & ""
                
            Set RsTemp = CrearRS(CSql)
            MsgBox "El Registro sea Actualizado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        End If
         
    Case Is = 1       'Agrega registro
   
        If Cambio = 1 Then
            CSql = "Update Internista set Activo=0 Where Activo=1 and IdPaciente=" & IdPac1
            Set RsTemp = CrearRS(CSql)
            
            CSql = "Insert into Internista(idinternista,Idpaciente,IdUsuario,Enfernedad_Act,Diagnostico,Signos,ColorP," & _
                "Cabello,Torax,Abdomen,Neurologico,Revision,Activo) VALUES (" & NuevoId & "," & IdPac1 & "," & IdUser & ",'" & _
                Text9.Text & "','" & Text8.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & _
                Text13.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & _
                Text18.Text & "',1)"

            Set RsTemp = CrearRS(CSql)
            MsgBox "Registro Agregado Satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        End If
End Select


If RegNew = 0 And Cambio = 1 Then
    If Reg_Actual(0) <> Text9.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Enfernedad_Act de (" & Reg_Actual(0) & ") a (" & Text9.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(1) <> Text8.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Diagnostico de (" & Reg_Actual(1) & ") a (" & Text8.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(2) <> Text10.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Signos de (" & Reg_Actual(2) & ") a (" & Text10.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(3) <> Text11.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo colorP de (" & Reg_Actual(3) & ") a (" & Text11.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(4) <> Text13.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Cabello de (" & Reg_Actual(4) & ") a (" & Text13.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(5) <> Text15.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Torax de (" & Reg_Actual(5) & ") a (" & Text15.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(6) <> Text16.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Abdomen de (" & Reg_Actual(6) & ") a (" & Text16.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(7) <> Text17.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Neurologico de (" & Reg_Actual(7) & ") a (" & Text17.Text & ") del Registro IdInternista=" & IdInter)
    If Reg_Actual(8) <> Text18.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Revision de (" & Reg_Actual(8) & ") a (" & Text18.Text & ") del Registro IdInternista=" & IdInter)
ElseIf RegNew = 1 And Cambio = 1 Then
    Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "INGRESAR", "Se ingreso un nuevo registro cuya IdInternista=" & NuevoId)
End If

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor Web"
EnviarRegPendiente

BtnDesHacer_Click
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly, "Error al Guardar"

Cambio = 0

End Sub

Sub EnviarRegPendiente()

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

Select Case RegNew
    
    Case Is = 0
         CSql = "Update Internista set IdUsuario=" & IdUser & ",Enfernedad_Act='" & Text9.Text & "',Diagnostico='" & Text8.Text & "',Signos='" & Text10.Text & "',ColorP='" & Text11.Text & "'," & _
        "Cabello='" & Text13.Text & "',Torax='" & Text15.Text & "',Abdomen='" & Text16.Text & "',Neurologico='" & Text17.Text & "',Revision='" & Text18.Text & "',Activo='1' Where Activo='1' and IdPaciente=" & IdPac1 & " And idinternista=" & IdInter & ""
        
        sentencia = Replace(CSql, "'", "(varCSP)")


        CSql = "Select * From Reg_Pendiente"
        Set RsRegPendiente = CrearRS(CSql)
        
        RsRegPendiente.AddNew
        RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
        RsRegPendiente.Fields("Modulo").Value = "Direccion Medica"
        RsRegPendiente.Fields("Tabla").Value = "Internista"
        RsRegPendiente.Fields("Condicional").Value = "IdInternista=" & IdInter & ""
        RsRegPendiente.Fields("Fecha").Value = DateTime.Date
        RsRegPendiente.Fields("Sentencia").Value = sentencia
        RsRegPendiente.Update
        
    Case Is = 1
        a = 1
        CSql = "INSERT into Internista(IdInternista,Idpaciente,IdUsuario,Enfernedad_Act,Diagnostico,Signos,ColorP," & _
        "Cabello,Torax,Abdomen,Neurologico,Revision,Activo) VALUES (" & NuevoId & "," & IdPac1 & "," & IdUser & "," & _
        "'" & Text9.Text & " ','" & Text8.Text & "','" & Text10.Text & "','" & Text11.Text & "'," & _
        "'" & Text13.Text & " ','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "'," & _
        "'" & Text18.Text & " ','" & a & "')"
        
        sentencia = Replace(CSql, "'", "(varCSP)")

        CSql = "Select * From Reg_Pendiente"
        Set RsRegPendiente = CrearRS(CSql)
        
        RsRegPendiente.AddNew
        RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
        RsRegPendiente.Fields("Modulo").Value = "Direccion Medica"
        RsRegPendiente.Fields("Tabla").Value = "Internista"
        RsRegPendiente.Fields("Condicional").Value = "IdInternista = " & NuevoId
        RsRegPendiente.Fields("Fecha").Value = DateTime.Date
        RsRegPendiente.Fields("Sentencia").Value = sentencia
        RsRegPendiente.Update
        
       
        
End Select


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub


Sub EnviarAlHosting()
On Error GoTo salir
ConectarHosting

If RegNew = 0 And Cambio = 1 Then
    If Reg_Actual(0) <> Text9.Text Then
        CSql = "Update Internista set Enfernedad_Act='" & Trim(Text9.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(1) <> Text8.Text Then
        CSql = "Update Internista set Diagnostico='" & Trim(Text8.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(2) <> Text10.Text Then
        CSql = "Update Internista set Signos='" & Trim(Text10.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(3) <> Text11.Text Then
        CSql = "Update Internista set ColorP='" & Trim(Text11.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(4) <> Text13.Text Then
        CSql = "Update Internista set Cabello='" & Trim(Text13.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(5) <> Text15.Text Then
        CSql = "Update Internista set Torax='" & Trim(Text15.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(6) <> Text16.Text Then
        CSql = "Update Internista set Abdomen='" & Trim(Text16.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(7) <> Text17.Text Then
        CSql = "Update Internista set Neurologico='" & Trim(Text17.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
    If Reg_Actual(8) <> Text18.Text Then
        CSql = "Update Internista set Revision='" & Trim(Text18.Text) & "' Where Activo=1 and IdPaciente=" & IdPac1
        Set RsWeb = CrearRsWeb(CSql)
    End If
End If

If RegNew = 1 And Cambio = 1 Then
    CSql = "Update Internista set Activo=0 Where Activo=1 and IdPaciente=" & IdPac1
    Set RsWeb = CrearRsWeb(CSql)
            
    CSql = "Insert into Internista(IdInternista,Idpaciente,IdUsuario,Enfernedad_Act,Diagnostico,Signos,ColorP," & _
        "Cabello,Torax,Abdomen,Neurologico,Revision,Activo) VALUES (" & NuevoId & "," & IdPac1 & "," & IdUser & ",'" & _
        Text9.Text & "','" & Text8.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & _
        Text13.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & _
        Text18.Text & "',1)"

    Set RsWeb = CrearRsWeb(CSql)
End If

CSql = "Select * From Internista Where IdInternista='" & NuevoId & "'"
Set RsWeb = CrearRsWeb(CSql)

If RsWeb.RecordCount > 0 Then
    Msg = "Actualización Completada en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
Else
    Msg = "No se realizo la Actualización en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Fallida"
    RsWeb.Close
    'WebCnn.Close
End If



salir:
If WebCnn.State = 0 Then
    EnviarAlHosting
Else
    GoTo f
End If

f:
If WebCnn.State = 1 Then
    WebCnn.Close
Else
    If Err.Number <> 0 Then MsgBox Err.Description
    Exit Sub
End If
End Sub
Private Sub BtnImprimir_Click()

On Error Resume Next

If Text1.Text = "" Then
    MsgBox "Debe de seleccionar un Paciente", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If

''========= ESTE ES EL CODIGO NUEVO ==========
With CrystalReport1
    .ReportFileName = RutaInformes & "\HistoriaClinicaIntegralN.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Historia_Clinica.Idpaciente} = " & IdPac1
    .ReportTitle = "Reporte Historia Medica No. " & Label12.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With



End Sub

Private Sub BtnListaEspera_Click()
'command11
ModulO = 3
FrmListaEspera.Show
'If Cedul <> "" Then
'Text7.Text = Cedul
'Call BtnBuscar_Click
'End If
End Sub

Private Sub BtnLlamar_Click()
'Command13
Call Llamar
End Sub

Private Sub BtnSiguiente_Click()
On Error Resume Next
If RsPaciente.RecordCount <> 0 Then
    RsPaciente.MoveNext
    If RsPaciente.EOF Then RsPaciente.MoveFirst
    Call Carga_De_Datos
    Call carga_de_datos1
    Cambio = 0
Else
    MsgBox "No hay registros cargados, Inicie una nueva Busqueda!", vbExclamation + vbOKOnly, "Error"
    IdPac1 = ""
End If
End Sub

Private Sub CboSexo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text14.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyLeft
            Text6.SetFocus
        Case vbKeyRight
            Text14.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub ChameleonBtn1_Click()
IO = 1
blanqueo1
End Sub

Sub Carga_De_Datos()

If RsPaciente.RecordCount <> 0 Then
    If Trim(RsPaciente.Fields("cedulaP")) <> "" Then Text1.Text = RsPaciente.Fields("cedulap")
    If Trim(RsPaciente.Fields("Fecha_regp")) <> "" Then DtpFechaRegistro.Value = RsPaciente.Fields("Fecha_regp")
    If Trim(RsPaciente.Fields("Nombrep")) <> "" Then Text3.Text = RsPaciente.Fields("Nombrep")
    If Trim(RsPaciente.Fields("Apellidop")) <> "" Then Text4.Text = RsPaciente.Fields("Apellidop")
    If Trim(RsPaciente.Fields("Fecha_nacimientop")) <> "" Then DtpFechaNac.Value = RsPaciente.Fields("Fecha_nacimientop")
    If Trim(RsPaciente.Fields("Edadp")) <> "" Then Text6.Text = RsPaciente.Fields("Edadp")
    If Trim(RsPaciente.Fields("Historia")) <> "" Then Label12.Caption = RsPaciente.Fields("Historia")
    If Trim(RsPaciente.Fields("Ocupacion")) <> "" Then Text14.Text = RsPaciente.Fields("Ocupacion")
    If RsPaciente.Fields("foto") <> "" Then
        If Len(Dir(Foto & "\" & RsPaciente.Fields("foto"))) > 0 Then
            Image2.Picture = LoadPicture(Foto & "\" & RsPaciente.Fields("foto"))
        Else
            Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        End If
    Else
        Image2.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
    IdPac1 = RsPaciente.Fields("Idpaciente")
    Me.Caption = "Dirección Médica - Paciente: " & IdPac1
    If RsPaciente.Fields("sexop") = 0 Then CboSexo.Text = "Masculino" Else CboSexo.Text = "Femenino"
    NoReg = "Registro " & RsPaciente.AbsolutePosition & " / " & RsPaciente.RecordCount
    BtnSiguiente.Enabled = True
    BtnAnterior.Enabled = True
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnImprimir.Enabled = True
    BtnEliminar.Enabled = True
    BtnExamenesLaboratorio.Enabled = True
    BtnAntecedentes.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
Else
    NoReg = "Registro 0 / 0"
    IdPac1 = ""
    BtnSiguiente.Enabled = False
    BtnAnterior.Enabled = False
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnImprimir.Enabled = False
    BtnEliminar.Enabled = False
    BtnExamenesLaboratorio.Enabled = False
    BtnAntecedentes.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
End If
End Sub

Sub carga_de_datos1()

If IdPac1 = "" Then
    CSql = "select * from Internista where activo=1"
Else
    CSql = "select * from Internista where idpaciente = " & IdPac1 & " and activo=1"
End If
Set RsInternista = CrearRS(CSql)

If RsInternista.RecordCount = 0 Then GoTo nohay

Cambio = 0
RegNew = 0
IdInter = RsInternista.Fields("IdInternista")
If RsInternista.Fields("Enfernedad_Act") <> "" Then Text9.Text = RsInternista.Fields("Enfernedad_Act") Else Text9.Text = ""
If RsInternista.Fields("Diagnostico") <> "" Then Text8.Text = RsInternista.Fields("Diagnostico") Else Text8.Text = ""
If RsInternista.Fields("Signos") <> "" Then Text10.Text = RsInternista.Fields("Signos") Else Text10.Text = ""
If RsInternista.Fields("colorP") <> "" Then Text11.Text = RsInternista.Fields("colorP") Else Text11.Text = ""
If RsInternista.Fields("Cabello") <> "" Then Text13.Text = RsInternista.Fields("Cabello") Else Text13.Text = ""
If RsInternista.Fields("Torax") <> "" Then Text15.Text = RsInternista.Fields("Torax") Else Text15.Text = ""
If RsInternista.Fields("Abdomen") <> "" Then Text16.Text = RsInternista.Fields("Abdomen") Else Text16.Text = ""
If RsInternista.Fields("Neurologico") <> "" Then Text17.Text = RsInternista.Fields("Neurologico") Else Text17.Text = ""
If RsInternista.Fields("Revision") <> "" Then Text18.Text = RsInternista.Fields("Revision") Else Text18.Text = ""
BtnImprimir.Enabled = True
BtnEliminar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnAgregar.Enabled = True
NoReg = "Registro " & RsInternista.AbsolutePosition & " / " & RsInternista.RecordCount

Reg_Actual(0) = RsInternista.Fields("Enfernedad_Act")
Reg_Actual(1) = RsInternista.Fields("Diagnostico")
Reg_Actual(2) = RsInternista.Fields("Signos")
Reg_Actual(3) = RsInternista.Fields("colorP")
Reg_Actual(4) = RsInternista.Fields("Cabello")
Reg_Actual(5) = RsInternista.Fields("Torax")
Reg_Actual(6) = RsInternista.Fields("Abdomen")
Reg_Actual(7) = RsInternista.Fields("Neurologico")
Reg_Actual(8) = RsInternista.Fields("Revision")

Exit Sub

'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
nohay:

For i = 0 To 10
    Reg_Actual(i) = ""
Next i
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnImprimir.Enabled = False
BtnEliminar.Enabled = False
IdInter = ""
Cambio = 0
RegNew = 1
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text13.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
'MsgBox "No hay Datos asociados al paciente: " & Chr(13) & Chr(13) & Text3.Text & " " & Text4.Text & Chr(13) & "Se Mostraran los datos de Historia Médica en blanco", vbExclamation + vbOKOnly, "No Tiene Datos"

End Sub

Private Sub DtpFechaActual_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyLeft
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaNac_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyRight
            Text6.SetFocus
        Case vbKeyDown
            ChameleonBtn1.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaActual.SetFocus
        Case vbKeyLeft
            Text1.SetFocus
        Case vbKeyRight
            DtpFechaActual.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me

ModulO = 3
    Frame2.Enabled = False
    Frame3.Enabled = False
CSql = "Select * From Paciente Order by IdPaciente"
Set RsPaciente = CrearRS(CSql)

DtpFechaActual.Value = Now()
SQL = ""
IdPac1 = 0
RegNew = 0

Call blanqueo1
Carga_De_Datos
carga_de_datos1
Cambio = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
IdPac1 = ""
If SQL <> "" Then RsPaciente.Close
SQL = ""

End Sub


Sub blanqueo1()
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text13.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaRegistro.SetFocus
        Case vbKeyRight
            DtpFechaRegistro.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case vbKeyUp
            Text8.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case vbKeyUp
            Text10.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text13.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text15.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text15.SetFocus
    End Select
End If
End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyLeft
            CboSexo.SetFocus
        Case vbKeyDown
            BtnAyuda.SetFocus
    End Select
End If
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text16.SetFocus
        Case vbKeyUp
            Text13.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text16.SetFocus
    End Select
End If
End Sub

Private Sub Text16_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text17.SetFocus
        Case vbKeyUp
            Text15.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            Text17.SetFocus
    End Select
End If
End Sub

Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text18.SetFocus
        Case vbKeyUp
            Text16.SetFocus
        Case vbKeyRight
            Text18.SetFocus
        Case vbKeyDown
            BtnExamenesLaboratorio.SetFocus
    End Select
End If
End Sub

Private Sub Text18_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnExamenesLaboratorio.SetFocus
        Case vbKeyUp
            Text8.SetFocus
        Case vbKeyLeft
            Text10.SetFocus
        Case vbKeyDown
            BtnExamenesLaboratorio.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaNac.SetFocus
        Case vbKeyUp
            DtpFechaActual.SetFocus
        Case vbKeyLeft
            Text4.SetFocus
        Case vbKeyDown
            Text14.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text19.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Text6.SetFocus
        Case vbKeyDown
            DtpFechaNac.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboSexo.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyLeft
            DtpFechaNac.SetFocus
        Case vbKeyRight
            CboSexo.SetFocus
        Case vbKeyDown
            BtnListaEspera.SetFocus
    End Select
End If
End Sub

Private Sub Text8_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text8.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text8.Text)
'    pru = LCase(Mid(Text8.Text, i, 1))
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
'Text8.Text = StrText
'Text8.SelStart = Len(Text8.Text)
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case vbKeyUp
            Text9.SetFocus
        Case vbKeyDown
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text9_Change()
Cambio = 1
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
'    End If
'Next i
'
'Text9.Text = StrText
'Text9.SelStart = Len(Text9.Text)
End Sub

Private Sub Text10_Change()
Cambio = 1
'
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text10.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text10.Text)
'    pru = LCase(Mid(Text10.Text, i, 1))
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
'Text10.Text = StrText
'Text10.SelStart = Len(Text10.Text)

End Sub

Private Sub Text11_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text11.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text11.Text)
'    pru = LCase(Mid(Text11.Text, i, 1))
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
'Text11.Text = StrText
'Text11.SelStart = Len(Text11.Text)
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text8.SetFocus
        Case vbKeyUp
            'ChameleonBtn1.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub Timer1_Timer()
 If IdPac1 <> "" Then BtnExamenesLaboratorio.Enabled = True Else BtnExamenesLaboratorio.Enabled = False
 If IdPac1 <> "" Then BtnAntecedentes.Enabled = True Else BtnAntecedentes.Enabled = False
End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub

End Sub
Private Sub Text7_GotFocus()
If Text7.Text = "Busqueda" Then Text7.Text = ""
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57 ' permite el ingreso de numeros
Case Is = 13 ' permite presionar el ENTER
Call BtnBuscar_Click
Case Is = 8 ' Permite Borrar de retroceso
Case Else ' Inhibe todas las demas teclas

End Select

End Sub

Private Sub Text7_Click()
If Text7.Text = "Busqueda" Then Text7.Text = ""
End Sub

Private Sub Text7_LostFocus()
If Trim(Text7.Text) = "" Then Text7.Text = "Busqueda"
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text13_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
' Dim i  As Variant
' StrText = ""
' Chaa = ""
'  Chaa = UCase(Mid(Text13.Text, 1, 1))
'  StrText = Chaa
'  For i = 2 To Len(Text13.Text)
'    pru = LCase(Mid(Text13.Text, i, 1))
'     If pru Like " " Then
'      T = 1
'      StrText = StrText & " "
'     Else
'      If T = 0 Then
'       Chaa = LCase(pru)
'       StrText = StrText + Chaa
'      Else
'       Chaa = UCase(pru)
'       StrText = StrText + Chaa
'       T = 0
'      End If
'     End If
'
'  Next i
'
' Text13.Text = StrText
' Text13.SelStart = Len(Text13.Text)
End Sub

Private Sub Text15_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
' Dim i  As Variant
' StrText = ""
' Chaa = ""
'  Chaa = UCase(Mid(Text15.Text, 1, 1))
'  StrText = Chaa
'  For i = 2 To Len(Text15.Text)
'    pru = LCase(Mid(Text15.Text, i, 1))
'     If pru Like " " Then
'      T = 1
'      StrText = StrText & " "
'     Else
'      If T = 0 Then
'       Chaa = LCase(pru)
'       StrText = StrText + Chaa
'      Else
'       Chaa = UCase(pru)
'       StrText = StrText + Chaa
'       T = 0
'      End If
'     End If
'
'  Next i
'
' Text15.Text = StrText
' Text15.SelStart = Len(Text15.Text)
End Sub

Private Sub Text16_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
' Dim i  As Variant
' StrText = ""
' Chaa = ""
'  Chaa = UCase(Mid(Text16.Text, 1, 1))
'  StrText = Chaa
'  For i = 2 To Len(Text16.Text)
'    pru = LCase(Mid(Text16.Text, i, 1))
'     If pru Like " " Then
'      T = 1
'      StrText = StrText & " "
'     Else
'      If T = 0 Then
'       Chaa = LCase(pru)
'       StrText = StrText + Chaa
'      Else
'       Chaa = UCase(pru)
'       StrText = StrText + Chaa
'       T = 0
'      End If
'     End If
'
'  Next i
'
' Text16.Text = StrText
' Text16.SelStart = Len(Text16.Text)
End Sub

Private Sub Text17_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
' Dim i  As Variant
' StrText = ""
' Chaa = ""
'  Chaa = UCase(Mid(Text17.Text, 1, 1))
'  StrText = Chaa
'  For i = 2 To Len(Text17.Text)
'    pru = LCase(Mid(Text17.Text, i, 1))
'     If pru Like " " Then
'      T = 1
'      StrText = StrText & " "
'     Else
'      If T = 0 Then
'       Chaa = LCase(pru)
'       StrText = StrText + Chaa
'      Else
'       Chaa = UCase(pru)
'       StrText = StrText + Chaa
'       T = 0
'      End If
'     End If
'
'  Next i
'
' Text17.Text = StrText
' Text17.SelStart = Len(Text17.Text)
End Sub

Private Sub Text18_Change()
Cambio = 1
'Dim StrText, Chaa, pru As String
' Dim i  As Variant
' StrText = ""
' Chaa = ""
'  Chaa = UCase(Mid(Text18.Text, 1, 1))
'  StrText = Chaa
'  For i = 2 To Len(Text18.Text)
'    pru = LCase(Mid(Text18.Text, i, 1))
'     If pru Like " " Then
'      T = 1
'      StrText = StrText & " "
'     Else
'      If T = 0 Then
'       Chaa = LCase(pru)
'       StrText = StrText + Chaa
'      Else
'       Chaa = UCase(pru)
'       StrText = StrText + Chaa
'       T = 0
'      End If
'     End If
'
'  Next i
'
' Text18.Text = StrText
' Text18.SelStart = Len(Text18.Text)
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then
    TxtBuscar.Text = ""
End If
If TxtBuscar.Text <> "Busqueda" Then
    TxtBuscar.Text = ""
End If
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

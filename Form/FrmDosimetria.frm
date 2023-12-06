VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDosimetria 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dosimetria"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "FrmDosimetria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   12270
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   3840
      TabIndex        =   32
      Top             =   10080
      Width           =   8295
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4440
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7200
         TabIndex        =   33
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
         MICON           =   "FrmDosimetria.frx":1002
         PICN            =   "FrmDosimetria.frx":101E
         PICH            =   "FrmDosimetria.frx":11E7
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
         TabIndex        =   34
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
         MICON           =   "FrmDosimetria.frx":141C
         PICN            =   "FrmDosimetria.frx":1438
         PICH            =   "FrmDosimetria.frx":16C7
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
         TabIndex        =   35
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
         MICON           =   "FrmDosimetria.frx":1B08
         PICN            =   "FrmDosimetria.frx":1B24
         PICH            =   "FrmDosimetria.frx":1CB1
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
         TabIndex        =   36
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
         MICON           =   "FrmDosimetria.frx":1EE6
         PICN            =   "FrmDosimetria.frx":1F02
         PICH            =   "FrmDosimetria.frx":21E4
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
         TabIndex        =   37
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
         MICON           =   "FrmDosimetria.frx":2435
         PICN            =   "FrmDosimetria.frx":2451
         PICH            =   "FrmDosimetria.frx":25F5
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   29
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
         TabIndex        =   30
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad o Número de Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   31
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
         MICON           =   "FrmDosimetria.frx":2794
         PICN            =   "FrmDosimetria.frx":27B0
         PICH            =   "FrmDosimetria.frx":2A15
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Height          =   7215
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   12015
      Begin VB.TextBox TxtPlanificacion 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   2520
         Width           =   11655
      End
      Begin VB.TextBox TxtEquipo 
         Height          =   615
         Left            =   6360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox TxtSimulacion 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   1440
         Width           =   9855
      End
      Begin VB.TextBox TxtDiagnostico 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   480
         Width           =   6135
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente2 
         Height          =   375
         Left            =   11160
         TabIndex        =   44
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
         MICON           =   "FrmDosimetria.frx":2CA7
         PICN            =   "FrmDosimetria.frx":2CC3
         PICH            =   "FrmDosimetria.frx":2F59
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
         Left            =   10320
         TabIndex        =   45
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
         MICON           =   "FrmDosimetria.frx":31B8
         PICN            =   "FrmDosimetria.frx":31D4
         PICH            =   "FrmDosimetria.frx":3469
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnSimulaciones 
         Height          =   375
         Left            =   10080
         TabIndex        =   46
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Simulaciones"
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
         MICON           =   "FrmDosimetria.frx":36C5
         PICN            =   "FrmDosimetria.frx":36E1
         PICH            =   "FrmDosimetria.frx":3AFF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnVerInformes 
         Height          =   375
         Left            =   10320
         TabIndex        =   47
         Top             =   6720
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "FrmDosimetria.frx":3DA3
         PICN            =   "FrmDosimetria.frx":3DBF
         PICH            =   "FrmDosimetria.frx":405B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCargarImagen 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   6720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cargar Imagen 1"
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
         MICON           =   "FrmDosimetria.frx":449B
         PICN            =   "FrmDosimetria.frx":44B7
         PICH            =   "FrmDosimetria.frx":470A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCargarImagen 
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   55
         Top             =   6720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cargar Imagen 2"
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
         MICON           =   "FrmDosimetria.frx":4967
         PICN            =   "FrmDosimetria.frx":4983
         PICH            =   "FrmDosimetria.frx":4BD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCargarImagen 
         Height          =   375
         Index           =   2
         Left            =   6840
         TabIndex        =   56
         Top             =   6720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cargar Imagen 3"
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
         MICON           =   "FrmDosimetria.frx":4E33
         PICN            =   "FrmDosimetria.frx":4E4F
         PICH            =   "FrmDosimetria.frx":50A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrarImagen 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   57
         ToolTipText     =   "Eliminar"
         Top             =   6720
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
         MICON           =   "FrmDosimetria.frx":52FF
         PICN            =   "FrmDosimetria.frx":531B
         PICH            =   "FrmDosimetria.frx":54BF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrarImagen 
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   58
         ToolTipText     =   "Eliminar"
         Top             =   6720
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
         MICON           =   "FrmDosimetria.frx":565E
         PICN            =   "FrmDosimetria.frx":567A
         PICH            =   "FrmDosimetria.frx":581E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrarImagen 
         Height          =   375
         Index           =   2
         Left            =   8760
         TabIndex        =   59
         ToolTipText     =   "Eliminar"
         Top             =   6720
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
         MICON           =   "FrmDosimetria.frx":59BD
         PICN            =   "FrmDosimetria.frx":59D9
         PICH            =   "FrmDosimetria.frx":5B7D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Planificación:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dia&gnóstico:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo:"
         Height          =   195
         Left            =   6360
         TabIndex        =   51
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Simulación:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Informe"
         Height          =   195
         Left            =   2160
         TabIndex        =   49
         Top             =   240
         Width           =   2970
      End
      Begin VB.Image Imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Imagenes:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   735
      End
      Begin VB.Image Imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Index           =   1
         Left            =   3480
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Image Imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Index           =   2
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   3135
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
      Begin MSComCtl2.DTPicker DtpFechaRegistro 
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249153
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaNac 
         Height          =   375
         Left            =   8400
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249153
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249153
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249153
         CurrentDate     =   40121
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
         Height          =   375
         Left            =   11040
         TabIndex        =   12
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
         MICON           =   "FrmDosimetria.frx":5D1C
         PICN            =   "FrmDosimetria.frx":5D38
         PICH            =   "FrmDosimetria.frx":5FCE
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
         TabIndex        =   13
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
         MICON           =   "FrmDosimetria.frx":622D
         PICN            =   "FrmDosimetria.frx":6249
         PICH            =   "FrmDosimetria.frx":64DE
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
         Left            =   5520
         TabIndex        =   38
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249153
         CurrentDate     =   39801
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4920
         TabIndex        =   39
         Top             =   1770
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         Height          =   195
         Left            =   4920
         TabIndex        =   27
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad:"
         Height          =   195
         Left            =   4920
         TabIndex        =   26
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nac.:"
         Height          =   195
         Left            =   7200
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label12 
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
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   9960
         Picture         =   "FrmDosimetria.frx":673A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Historia:"
         Height          =   195
         Left            =   3840
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   2250
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmDosimetria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsInformeMed As New ADODB.Recordset 'Para desplazamientos
Dim RsInformeDosimetria As New ADODB.Recordset
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

Sub CARGAR_INFORME2()

CSql = "SELECT Re,Rp,Her2Neu,EMA,VIM,CAE,[CERB-2],P53,DESMINA,ACE,AFP,[PROT-S-100],PGP,CD31,CD34, " & _
    "CD117,CK5,CK6,CK7,CK20,[CAM 5,2],[TTF-1],CROMOGRANINA,SINAPTOFISINA,CD56,CD57,EGFR,KIT,AE1,AE3, " & _
    "CK903,GFAP,SMA,CA199,CA125,CEA,[CEA-D14],[E-CAD],HCG,[HMB-45],HPAP,WT1,[BEL-1],[BEL-2],PRB,[ALK-1]," & _
    "RA,OTROS,IdTipoCancer,T,N,M,Estadio,CP,G,Gleason,Reseccion " & _
    "FROM Informe_Medico2 Where IdPaciente=" & IdPac1 & " AND IdInforme=" & IdInf
    
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

If Trim(RsTemp.Fields("CP").Value) <> "" Then
    For i = 0 To Combo3.ListCount - 1
        If Trim(RsTemp.Fields("CP").Value) = Combo3.List(i) Then
            Combo3.ListIndex = i
            Exit For
        End If
        Combo3.ListIndex = -1
    Next i
End If

If Trim(RsTemp.Fields("T").Value) <> "" Then
    For i = 0 To Combo4.ListCount - 1
        If Trim(RsTemp.Fields("T").Value) = Combo4.List(i) Then
            Combo4.ListIndex = i
            Exit For
        End If
        Combo4.ListIndex = -1
    Next i
End If

If Trim(RsTemp.Fields("N").Value) <> "" Then
    For i = 0 To Combo5.ListCount - 1
        If Trim(RsTemp.Fields("N").Value) = Combo5.List(i) Then
            Combo5.ListIndex = i
            Exit For
        End If
        Combo5.ListIndex = -1
    Next i
End If

If Trim(RsTemp.Fields("M").Value) <> "" Then
    For i = 0 To Combo6.ListCount - 1
        If Trim(RsTemp.Fields("M").Value) = Combo6.List(i) Then
            Combo6.ListIndex = i
            Exit For
        End If
        Combo6.ListIndex = -1
    Next i
End If

If Trim(RsTemp.Fields("Estadio").Value) <> "" Then
    For i = 0 To Combo2.ListCount - 1
        If Trim(RsTemp.Fields("Estadio").Value) = Combo2.List(i) Then
            Combo2.ListIndex = i
            Exit For
        End If
        Combo2.ListIndex = -1
    Next i
End If

If Trim(RsTemp.Fields("G").Value) <> "" Then
    For i = 0 To Combo7.ListCount - 1
        If Trim(RsTemp.Fields("G").Value) = Combo7.List(i) Then
            Combo7.ListIndex = i
            Exit For
        End If
        Combo7.ListIndex = -1
    Next i
End If

If Trim(RsTemp.Fields("Gleason").Value) <> "" Then
    TxtGleason.Text = Trim(RsTemp.Fields("Gleason").Value)
    ChkGleason.Value = 1
Else
    TxtGleason.Text = ""
    ChkGleason.Value = 0
End If

'OptBordes(0).Value = True
'OptBordes(0).Value = False
For i = 0 To OptBordes.Count - 1
    If OptBordes(i).Caption = Trim(RsTemp.Fields("Reseccion").Value) Then
        OptBordes(i).Value = True
        Exit For
    End If
Next i

For i = 0 To CboTCancers.ListCount - 1
    If CboTCancers.ItemData(i) = RsTemp.Fields("IdTipoCancer").Value Then
        CboTCancers.ListIndex = i
        Exit For
    End If
    CboTCancers.ListIndex = -1
Next i

'
'' Ciclo que carga TODOS los valores para la "Gran Matriz"
'For i = 0 To ChkGlobal.Count - 1
'    If Trim(RsTemp.Fields(i).Value) <> "" Then
'
'        If i <= 56 Then
'            For j = 0 To CboGeneral(i).ListCount - 1
'                If Trim(CboGeneral(i).List(j)) = Trim(RsTemp.Fields(i).Value) Then
'                    CboGeneral(i).ListIndex = j
'                    ChkGlobal(i).Value = 1
'                    Exit For
'                End If
'            Next j
'        Else
'            If Trim(RsTemp.Fields("OTROS").Value) <> "" Then
'                TxtOtros.Text = Trim(RsTemp.Fields("OTROS").Value)
'                ChkGlobal(57).Value = 1
'            End If
'        End If
'    Else
'        ChkGlobal(i).Value = 0
'        CboGeneral(i).ListIndex = -1
'    End If
'Next i

End Sub

Sub carga_datos_radio()
On Error GoTo WrtError




If RsInformeDosimetria.RecordCount <> 0 Then
    Cambio = 1
    If Not IsNull(RsInformeDosimetria.Fields("Diagnostico").Value) Then TxtDiagnostico.Text = RsInformeDosimetria.Fields("Diagnostico").Value Else TxtDiagnostico.Text = ""
    d1 = Trim(RsInformeDosimetria.Fields("Diagnostico").Value)
    
    If Not IsNull(RsInformeDosimetria.Fields("Simulacion").Value) Then TxtSimulacion.Text = RsInformeDosimetria.Fields("Simulacion").Value Else TxtSimulacion.Text = ""
    d2 = Trim(RsInformeDosimetria.Fields("Simulacion").Value)
    
    If Not IsNull(RsInformeDosimetria.Fields("Planificacion").Value) Then TxtPlanificacion.Text = RsInformeDosimetria.Fields("Planificacion").Value Else TxtPlanificacion.Text = ""
    d3 = Trim(RsInformeDosimetria.Fields("Planificacion").Value)
    
    If Not IsNull(RsInformeDosimetria.Fields("Equipo").Value) Then TxtEquipo.Text = RsInformeDosimetria.Fields("Equipo").Value Else TxtEquipo.Text = ""
    d4 = Trim(RsInformeDosimetria.Fields("Equipo").Value)
    
    Label13.Caption = "Registro: " & RsInformeDosimetria.AbsolutePosition & " / " & RsInformeDosimetria.RecordCount
    
    DtpFecha.Value = Trim(RsInformeDosimetria.Fields("Fecha").Value)
    d11 = RsInformeDosimetria.Fields("fecha")
    
    IdReg = Trim(RsInformeDosimetria.Fields("IdDosimetria").Value)
        
    TxtDiagnostico.ToolTipText = IdInf
    
    d14 = IdReg

'****** Fotos **********

    If Not IsNull(RsInformeDosimetria.Fields("Imagen1").Value) Then
        If RsInformeDosimetria.Fields("Imagen1") <> "" Then
            Imagen(0).Picture = LoadPicture(FotoSimul & "\" & RsInformeDosimetria.Fields("Imagen1").Value)
            BtnBorrarImagen(0).Enabled = True
            Imagen(0).Tag = Trim(RsInformeDosimetria.Fields("Imagen1").Value)
        Else
            Imagen(0).Picture = LoadPicture()
            BtnBorrarImagen(0).Enabled = False
        End If
    Else
        Imagen(0).Picture = LoadPicture()
        BtnBorrarImagen(0).Enabled = False
    End If
    
    If Not IsNull(RsInformeDosimetria.Fields("Imagen2").Value) Then
        If RsInformeDosimetria.Fields("Imagen2") <> "" Then
            Imagen(1).Picture = LoadPicture(FotoSimul & "\" & RsInformeDosimetria.Fields("Imagen2").Value)
            BtnBorrarImagen(1).Enabled = True
            Imagen(1).Tag = Trim(RsInformeDosimetria.Fields("Imagen2").Value)
        Else
            Imagen(1).Picture = LoadPicture()
            BtnBorrarImagen(1).Enabled = False
        End If
    Else
        Imagen(1).Picture = LoadPicture()
        BtnBorrarImagen(1).Enabled = False
    End If

    If Not IsNull(RsInformeDosimetria.Fields("Imagen3").Value) Then
        If RsInformeDosimetria.Fields("Imagen3") <> "" Then
            Imagen(2).Picture = LoadPicture(FotoSimul & "\" & RsInformeDosimetria.Fields("Imagen3").Value)
            BtnBorrarImagen(2).Enabled = True
            Imagen(2).Tag = Trim(RsInformeDosimetria.Fields("Imagen3").Value)
        Else
            Imagen(2).Picture = LoadPicture()
            BtnBorrarImagen(2).Enabled = False
        End If
    Else
        Imagen(2).Picture = LoadPicture()
        BtnBorrarImagen(2).Enabled = False
    End If
    
ElseIf RsInformeMed.RecordCount <> 0 Then
    Cambio = 0
    If Not IsNull(RsInformeMed.Fields("Diagnotico").Value) Then TxtDiagnostico.Text = RsInformeMed.Fields("Diagnotico").Value Else TxtDiagnostico.Text = ""
    d9 = Trim(RsInformeMed.Fields("Diagnotico").Value)
    
    Label13.Caption = "Registro: " & RsInformeMed.AbsolutePosition & " / " & RsInformeMed.RecordCount
    
    DtpFecha.Value = Trim(RsInformeMed.Fields("fecha").Value)
    d11 = RsInformeMed.Fields("fecha")
    
    IdReg = Trim(RsInformeMed.Fields("IdInforme").Value)
    IdInf = Trim(RsInformeMed.Fields("IdInforme").Value)
    
    TxtDiagnostico.ToolTipText = IdInf
    
    d14 = IdReg
    
    
    Cambio = 0
    actualiza = 1
    CboModificarMedicoTratante.Enabled = True
    Call Habilita_Btns("Todos")
    DesactivarTextos
    

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

Exit Sub
WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
'MsgBox MError
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


TxtDiagnostico.Text = ""
TxtEquipo.Text = ""
TxtSimulacion.Text = ""
TxtPlanificacion.Text = ""
Image1.Picture = LoadPicture()
Image2.Picture = LoadPicture()
Image3.Picture = LoadPicture()

DtpFecha.Value = Date
Label13.Caption = "Registro 0 / 0  (Sin Informe Médico)"

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
If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente antes de agregar un Informe Medico!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub
IO = 1
Call Deshabilita_Btns
Call blanqueo1
Cambio = 0
actualiza = 0
Frame1.Enabled = False
Frame6.BackColor = &HE0E0E0
DesactivarTextos

DtpFecha.Value = Now
End Sub

Sub ActivarTextos()
TxtDiagnostico.Locked = True
TxtEquipo.Locked = True
TxtSimulacion.Locked = True
TxtPlanificacion.Locked = True

DtpFecha.Enabled = True

For i = 0 To BtnCargarImagen.Count - 1
    'If i <> Index Then
        BtnCargarImagen(i).Enabled = False
  
    'End If
Next

For i = 0 To BtnBorrarImagen.Count - 1
    'If i <> Index Then
        BtnBorrarImagen(i).Enabled = False
  
    'End If
Next
End Sub

Sub DesactivarTextos()
TxtDiagnostico.Locked = False
TxtEquipo.Locked = False
TxtSimulacion.Locked = False
TxtPlanificacion.Locked = False
DtpFecha.Enabled = False

For i = 0 To BtnCargarImagen.Count - 1
    'If i <> Index Then
        BtnCargarImagen(i).Enabled = True
  
   ' End If
Next

End Sub
'Sub Deshabilita_Btns(cad As String)
Sub Deshabilita_Btns()
On Error Resume Next
'If UCase(Cad) = UCase("Todos") Then
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = False
    BtnVerInformes.Enabled = False
    
    
    BtnAnterior1.Enabled = False
    BtnSiguiente2.Enabled = False
    
'End If
End Sub

Sub Habilita_Btns(Cad As String)
On Error Resume Next
If UCase(Cad) = UCase("Todos") Then
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = True
    BtnVerInformes.Enabled = True
    
    BtnAnterior1.Enabled = True
    BtnSiguiente2.Enabled = True
    
ElseIf UCase(Cad) = UCase("sin informe") Then
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnVerInformes.Enabled = False
    BtnSimulaciones.Enabled = False
    BtnVerHistoria.Enabled = True
    BtnAnterior1.Enabled = False
    BtnSiguiente2.Enabled = False
    BtnExamenes.Enabled = True
End If
Frame6.Enabled = True
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
    Else
    MsgBox "No se encontraron datos, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros cargados"
End If
'If OptInformeMedico(0).Value Then BtnGuardarActualizar.Enabled = Fals
End Sub

Private Sub BtnAnterior1_Click()
On Error Resume Next
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

Private Sub BtnBorrarImagen_Click(Index As Integer)
On Error Resume Next
Dim TempCad As String
Dim TempCad2 As String
On Error GoTo h
      
        Imagen(Index).Picture = LoadPicture()
        Imagen(Index).Refresh


BtnBorrarImagen(Index).Enabled = False
Exit Sub
h:
MsgBox Err.Description
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

BtnDesHacer_Click

If Trim(TxtBuscar.Text) = "" Or UCase(TxtBuscar.Text) = UCase("Busqueda") Then
    f = "Buscar"
    CSql = "select * from Paciente where Historia like 'HOAM%' or Historia like 'hoam%' order by IdPaciente"
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
    BtnSimulaciones.Enabled = True
    BtnVerInformes.Enabled = True
    BtnAnterior1.Enabled = True
    BtnSiguiente2.Enabled = True
    
Call Carga_De_Datos
Call CONSULTA_INFORME
Call carga_datos_radio

DesactivarTextos
Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub BtnCargarImagen_Click(Index As Integer)
On Error Resume Next
Dim TempCad As String
Dim TempCad1 As String
On Error GoTo WrtError

CommonDialog1.ShowOpen
TempCad = CommonDialog1.filename
TempCad1 = CommonDialog1.FileTitle

If Trim(TempCad) = "" Then Exit Sub
Imagen(Index).Picture = LoadPicture(TempCad)
Imagen(Index).Refresh
Imagen(Index).Tag = TempCad1
Cambio = 1

BtnBorrarImagen(Index).Enabled = True
Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
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
Frame6.BackColor = &HEAEFEF


If IdPac1 = "" Then
    CSql = "select * from Paciente"
    Set RsCargarPacientes = CrearRS(CSql)
    RsCargarPacientes.MoveFirst
End If

Carga_De_Datos

Frame6.BackColor = &HEAEFEF

End Sub

Private Sub BtnEliminar_Click()
On Error GoTo WrtError
Dim RsDeshabilitar As New ADODB.Recordset
Dim RsVerificar As New ADODB.Recordset

If IdPac1 = "" Then MsgBox "Debe seleccionar un Paciente antes de Eliminar!", vbExclamation + vbOKOnly, "Verifique el Paciente": Exit Sub

CSql = "Select * From Dosimetria Where IdDosimetria='" & IdReg & "'"
Set RsVerificar = CrearRS(CSql)

If RsVerificar.Fields("IdUser").Value = IdUser Then

    resp = MsgBox("Desea eliminar el Informe de Dosimetria # " & RsInformeDosimetria.AbsolutePosition & ", del Paciente de C.I.: " & Text1.Text & " ?", vbQuestion + vbYesNo, "Confirmar!")
    
    If resp = 7 Then Exit Sub
    
    Call Enviar_Bitacora(IdUser, "Dosimetria", "BORRAR", "Se elimino el INFORME de Dosimetria cuya IdDosimetria es (" & IdReg & ")")
    
    CSql = "update Dosimetria set Activo=2 where IdDosimetria = " & IdReg
    Set RsDeshabilitar = CrearRS(CSql)
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Borrar Informe Dosimetrico"
    BorrarRegPendiente
    
    MsgBox "El Informe Dosimetrico del Paciente: " & Text3.Text & " " & Text4.Text & " ha sido eliminado del Registro!", vbInformation + vbOKOnly, "Operacion Exitosa!"
    BtnDesHacer_Click

Else
    MsgBox "Usted No tiene permiso para borrar este informe Dosimetrico", vbCritical + vbOKOnly, "Error"

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

Private Sub BtnGuardarActualizar_Click()
Dim RsGuardar As New ADODB.Recordset
On Error GoTo WrtError

If IdPac1 = "" Then
    MsgBox "Debe seleccionar un Paciente antes de Guardar!", vbExclamation + vbOKOnly, "Verifique el Paciente"
    Exit Sub
End If

If Trim(TxtEquipo.Text) = "" Then
    MsgBox "El Campo Equipo No debe estar Vacio", vbExclamation + vbOKOnly, "Faltan Datos"
    TxtEquipo.SetFocus
    Exit Sub
ElseIf TxtDiagnostico.Text = "" Then
    MsgBox "El Diagnóstico no debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    TxtDiagnostico.SetFocus
    Exit Sub
ElseIf TxtSimulacion.Text = "" Then
    MsgBox "El campo Simulación No debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    TxtSimulacion.SetFocus
    Exit Sub
ElseIf TxtPlanificacion.Text = "" Then
    MsgBox "El campo Planificación No debe estar vacio!", vbExclamation + vbOKOnly, "Faltan Datos"
    TxtPlanificacion.SetFocus
    Exit Sub
End If

' Obtiene el Nuevo ID para Dosimetria
    Dim RsMaxReg As New ADODB.Recordset
    CSql = "Select max(IdDosimetria)+1 as MaxReg From Dosimetria"
    Set RsMaxReg = CrearRS(CSql)

    If Not IsNull(RsMaxReg.Fields("MaxReg").Value) Then
        MaxReg = RsMaxReg.Fields("MaxReg").Value
    Else
        MaxReg = "0"
    End If
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm

Select Case Cambio

Case Is = 0
     
        CSql = "Select * From Dosimetria"
        Set RsGuardar = CrearRS(CSql)
        
        RsGuardar.AddNew
        RsGuardar.Fields("IdDosimetria").Value = MaxReg
        RsGuardar.Fields("IdPaciente").Value = IdPac1
        RsGuardar.Fields("IdUser").Value = IdUser
        RsGuardar.Fields("Simulacion").Value = Trim(TxtSimulacion.Text)
        RsGuardar.Fields("Planificacion").Value = Trim(TxtPlanificacion.Text)
        RsGuardar.Fields("Diagnostico").Value = Trim(TxtDiagnostico.Text)
        RsGuardar.Fields("Equipo").Value = Trim(TxtEquipo.Text)
        RsGuardar.Fields("Imagen1").Value = Imagen(0).Tag
        RsGuardar.Fields("Imagen2").Value = Imagen(1).Tag
        RsGuardar.Fields("Imagen3").Value = Imagen(2).Tag
             
        RsGuardar.Update
        
       'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
       'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        MsgBox "Los datos fueron guardardos!", vbInformation + vbOKOnly, "Operación exitosa!"
        
        Call Habilita_Btns("Todos")
        Frame6.BackColor = &HEAEFEF

        ActivarTextos
        Call CONSULTA_INFORME
        
        
        BtnDesHacer_Click
        
Case Is = 1
    
        If IdReg = "" Then MsgBox "No hay informes seleccionados!", vbExclamation + vbOKOnly, "Error": Exit Sub
         
   
        CSql = "Select * From Dosimetria Where IdPaciente='" & IdPac1 & "' And IdDosimetria='" & IdReg & "'"
        Set RsGuardar = CrearRS(CSql)

        RsGuardar.Fields("IdUser").Value = IdUser

        RsGuardar.Fields("Simulacion").Value = Trim(TxtSimulacion.Text)
        RsGuardar.Fields("Planificacion").Value = Trim(TxtPlanificacion.Text)
        RsGuardar.Fields("Diagnostico").Value = Trim(TxtDiagnostico.Text)
        RsGuardar.Fields("Equipo").Value = Trim(TxtEquipo.Text)
        RsGuardar.Fields("Imagen1").Value = Imagen(0).Tag
        RsGuardar.Fields("Imagen2").Value = Imagen(1).Tag
        RsGuardar.Fields("Imagen3").Value = Imagen(2).Tag
       
        RsGuardar.Update
        
        Call Habilita_Btns("Todos")
        Frame6.BackColor = &HEAEFEF

        
        MsgBox "Registro actualizado Satisfactoriamente, se procedera a guardar los datos estadisticos!s", vbInformation + vbOKOnly, "Operacion Exitosa"
        ActivarTextos
        Call CONSULTA_INFORME
        
        'Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        'MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Informe Médico"
'        EnviarAlHosting
'        EnviarRegPendiente
        
       'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
       'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
End Select
        
 
Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
MsgBox MError
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub


'Sub EnviarRegPendiente()
'On Error GoTo WrtError
'CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
'Set RsIdMax = CrearRS(CSql)
'
'If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
'    MaxRegP = RsIdMax.Fields("IdMax").Value
'Else
'    MaxRegP = "1"
'End If
'
'a = 1
'CSql = "INSERT into INFORME_MEDICO(IdInforme,idmedicot,idusuario,idpaciente,Antecedente_Flia," & _
'        "Anatomia_Patol,Enfermedad_Act,Examen_Fis,Motivo_Con,Diagnotico,Tratamiento,Fecha,dosis," & _
'        "dosisd,Tomografia,Cuantas,Estado,Metas,sesiones,Estadiaje) VALUES(" & MaxReg & "," & IdMedT & "," & IdUser & "," & IdPac1 & _
'        ",'" & Text15.Text & "','" & Text19.Text & "','" & Text16.Text & "','" & Text20.Text & _
'        "','" & Text18.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & _
'        Format(DtpFecha.Value, "mm/dd/YYYY") & "','" & Val(Text8.Text) & "','" & Val(Text9.Text) & _
'        "'," & tomog & ",'" & Val(Text17.Text) & "','" & a & "','" & Trim(CboMetas.Text) & "'," & Sesiones & ",'" & Estadiaje & "')"
'sentencia = Replace(CSql, "'", "(varCSP)")
'
'
'CSql = "Select * From Reg_Pendiente"
'Set RsRegPendiente = CrearRS(CSql)
'
'RsRegPendiente.AddNew
'RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
'RsRegPendiente.Fields("Modulo").Value = "Oncologia"
'RsRegPendiente.Fields("Tabla").Value = "Informe_Medico"
'RsRegPendiente.Fields("Condicional").Value = "IdInforme=" & MaxReg
'RsRegPendiente.Fields("Fecha").Value = DateTime.Date
'RsRegPendiente.Fields("Sentencia").Value = sentencia
'RsRegPendiente.Update
'
'Msg = "Envio Al Servidor Web Satisfactorio!!!"
'MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"
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
        Call CONSULTA_INFORME
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
    Else
    MsgBox "No se encontraron datos, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros cargados"
End If

'If OptInformeMedico(0).Value Then BtnGuardarActualizar.Enabled = False
End Sub

Private Sub BtnSiguiente2_Click()
On Error Resume Next
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

Sub CONSULTA_INFORME()
On Error GoTo WrtError

CSql = "Select * From Dosimetria Where IdPaciente = '" & IdPac1 & "'"
Set RsInformeDosimetria = CrearRS(CSql)


CSql = "Select * From Informe_Medico Where IdPaciente = " & IdPac1 & " And Estado=1 Order By Fecha Desc"
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
    
    Me.Caption = "Dosimetria - Paciente: " & IdPac1
    
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

Private Sub BtnSimulaciones_Click()
FrmSimulaciones.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnVerInformes_Click()
On Error GoTo WrtError
If Text1.Text <> "" Then

''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\InformeDosimetria.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{Informe_Dosimetria.IdDosimetria} = " & IdReg
        .WindowTitle = "Reporte de Planificación No. " & IdReg
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

Private Sub Form_Load()
On Error GoTo WrtError

Centrar Me
ModulO = 4

ActivarTextos
CSql = "select * from Medicos where Tipo=2 or Tipo=3"
Set RsTemp = CrearRS(CSql)

For i = 1 To RsTemp.RecordCount
    CboModificarMedicoTratante.AddItem RsTemp.Fields("Nombre").Value & " " & RsTemp.Fields("Apellido").Value
    CboModificarMedicoTratante.ItemData(CboModificarMedicoTratante.NewIndex) = Val(RsTemp.Fields("IdMedico").Value)
    RsTemp.MoveNext
Next i

CSql = "select * from Paciente"
Set RsCargarPacientes = CrearRS(CSql)

If RsCargarPacientes.RecordCount <> 0 Then
    Carga_De_Datos
    Call CONSULTA_INFORME
    If RsInformeMed.RecordCount <> 0 Then
        Call carga_datos_radio
        Cambio = 0
        actualiza = 1
    Else
        Call Blanqueo
        Call Habilita_Btns("sin informe")
        CboModificarMedicoTratante.Enabled = False
        Cambio = 0
        actualiza = 0
    End If
    
    DtpFecha.Value = Now()
    
Else
    MsgBox "No se encontraron pacientes en la Base d e Datos!", vbExclamation + vbOKOnly, "Vacio"
    IdPac1 = ""
End If


Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
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
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        Case vbKeyUp
            Text22.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
    End Select
End If
End Sub



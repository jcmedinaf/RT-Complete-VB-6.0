VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPresupuestoTratamientos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
   ClientHeight    =   7050
   ClientLeft      =   4335
   ClientTop       =   2490
   ClientWidth     =   11025
   Icon            =   "Presupuesto.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11025
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   3360
         TabIndex        =   39
         Top             =   6000
         Width           =   7335
         Begin ChamaleonButton.ChameleonBtn BtnImprimir 
            Height          =   375
            Left            =   3240
            TabIndex        =   44
            ToolTipText     =   "Reporte"
            Top             =   300
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
            MICON           =   "Presupuesto.frx":1002
            PICN            =   "Presupuesto.frx":101E
            PICH            =   "Presupuesto.frx":1143
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
            Left            =   6240
            TabIndex        =   40
            ToolTipText     =   "Cerrar"
            Top             =   300
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
            MICON           =   "Presupuesto.frx":13D3
            PICN            =   "Presupuesto.frx":13EF
            PICH            =   "Presupuesto.frx":15B8
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
            Left            =   5040
            TabIndex        =   42
            ToolTipText     =   "Deshacer Operacion"
            Top             =   300
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
            MICON           =   "Presupuesto.frx":17ED
            PICN            =   "Presupuesto.frx":1809
            PICH            =   "Presupuesto.frx":1AEB
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
            Left            =   3000
            TabIndex        =   55
            ToolTipText     =   "Crea un nuevo presupuesto"
            Top             =   300
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Nuevo"
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
            MICON           =   "Presupuesto.frx":1D3C
            PICN            =   "Presupuesto.frx":1D58
            PICH            =   "Presupuesto.frx":1EE5
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
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   300
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
            MICON           =   "Presupuesto.frx":211A
            PICN            =   "Presupuesto.frx":2136
            PICH            =   "Presupuesto.frx":23C5
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
            Left            =   1320
            TabIndex        =   43
            ToolTipText     =   "Eliminar"
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "Presupuesto.frx":2806
            PICN            =   "Presupuesto.frx":2822
            PICH            =   "Presupuesto.frx":29C6
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   6000
         Width           =   3135
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el Número de Cédula o Número de Presupuesto a buscar"
            Top             =   300
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   1680
            TabIndex        =   16
            ToolTipText     =   "Buscar"
            Top             =   315
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
            MICON           =   "Presupuesto.frx":2B65
            PICN            =   "Presupuesto.frx":2B81
            PICH            =   "Presupuesto.frx":2DE6
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
         Height          =   3735
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   10575
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BackColor       =   &H00EAEFEF&
            Height          =   375
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   3120
            Width           =   1065
         End
         Begin VB.Timer Timer1 
            Left            =   9960
            Top             =   3240
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarDiagnostico 
            Height          =   855
            Left            =   8760
            TabIndex        =   36
            Top             =   2160
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1508
            BTYPE           =   3
            TX              =   "&Buscar Diagnóstico"
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
            MICON           =   "Presupuesto.frx":3078
            PICN            =   "Presupuesto.frx":3094
            PICH            =   "Presupuesto.frx":34BC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00EAEFEF&
            Height          =   405
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            TabIndex        =   9
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00EAEFEF&
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1800
            Width           =   6975
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00EAEFEF&
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   840
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00EAEFEF&
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1320
            Width           =   4815
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00EAEFEF&
            Height          =   405
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   3120
            Width           =   1815
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BackColor       =   &H00EAEFEF&
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   3120
            Width           =   1815
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            BackColor       =   &H00EAEFEF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2400
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tomografía"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   2460
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6720
            TabIndex        =   8
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   21299201
            CurrentDate     =   39834
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   9000
            Top             =   3120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            PrintFileLinesPerPage=   60
            WindowShowCloseBtn=   -1  'True
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   9480
            Top             =   3120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente1 
            Height          =   495
            Left            =   9600
            TabIndex        =   45
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   1200
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
            MICON           =   "Presupuesto.frx":3755
            PICN            =   "Presupuesto.frx":3771
            PICH            =   "Presupuesto.frx":3A07
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
            Height          =   495
            Left            =   8880
            TabIndex        =   46
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   1200
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
            MICON           =   "Presupuesto.frx":3C66
            PICN            =   "Presupuesto.frx":3C82
            PICH            =   "Presupuesto.frx":3F17
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Sesiones"
            Height          =   195
            Left            =   5400
            TabIndex        =   53
            Top             =   2880
            Width           =   1065
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Costo del Tratamiento:"
            Height          =   555
            Left            =   120
            TabIndex        =   35
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   3210
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1890
            Width           =   885
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   6120
            TabIndex        =   32
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   435
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1410
            Width           =   765
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            Height          =   195
            Left            =   5160
            TabIndex        =   28
            Top             =   3210
            Width           =   210
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Dosis Diarias"
            Height          =   255
            Left            =   3240
            TabIndex        =   27
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Dosis"
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            Height          =   195
            Left            =   3000
            TabIndex        =   25
            Top             =   3225
            Width           =   75
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico/Tratamiento"
            Height          =   195
            Left            =   8595
            TabIndex        =   24
            Top             =   480
            Width           =   1785
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   8640
            TabIndex        =   23
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00EAEFEF&
            Caption         =   "Se le debe anexar el costo de todas las tomografias"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   5160
            TabIndex        =   56
            ToolTipText     =   "El número de sesiones es menor a 25"
            Top             =   2400
            Visible         =   0   'False
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Cliente"
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10575
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Height          =   1695
            Left            =   8640
            TabIndex        =   47
            Top             =   120
            Width           =   1815
            Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
               Height          =   495
               Left            =   960
               TabIndex        =   48
               ToolTipText     =   "Moverse la Registro Siguiente"
               Top             =   960
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
               MICON           =   "Presupuesto.frx":4173
               PICN            =   "Presupuesto.frx":418F
               PICH            =   "Presupuesto.frx":4425
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
               Left            =   240
               TabIndex        =   49
               ToolTipText     =   "Moverse la Registro Anterior"
               Top             =   960
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
               MICON           =   "Presupuesto.frx":4684
               PICN            =   "Presupuesto.frx":46A0
               PICH            =   "Presupuesto.frx":4935
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. Presupuesto"
               Height          =   195
               Left            =   270
               TabIndex        =   50
               Top             =   240
               Width           =   1215
            End
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarClientes 
            Height          =   375
            Left            =   4560
            TabIndex        =   38
            ToolTipText     =   "Agrega los datos de los clientes"
            Top             =   310
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar Cliente"
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
            MICON           =   "Presupuesto.frx":4B91
            PICN            =   "Presupuesto.frx":4BAD
            PICH            =   "Presupuesto.frx":4D3A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   1200
            TabIndex        =   0
            Top             =   310
            Width           =   1935
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00EAEFEF&
            Height          =   615
            Left            =   1200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   1200
            Width           =   6975
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   840
            Width           =   6975
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarClientes 
            Height          =   375
            Left            =   3240
            TabIndex        =   37
            ToolTipText     =   "Buscar información del cliente"
            Top             =   310
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "Presupuesto.frx":4EE4
            PICN            =   "Presupuesto.frx":4F00
            PICH            =   "Presupuesto.frx":5328
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00EAEFEF&
            Height          =   405
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   720
            Width           =   6975
         End
         Begin ChamaleonButton.ChameleonBtn BtnListadoProductos 
            Height          =   375
            Left            =   6240
            TabIndex        =   54
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
            Top             =   315
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Listado Clientes"
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
            MICON           =   "Presupuesto.frx":55C1
            PICN            =   "Presupuesto.frx":55DD
            PICH            =   "Presupuesto.frx":5866
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   825
            Width           =   990
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1290
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RIF.:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   400
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "FrmPresupuestoTratamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As ADODB.Recordset
Dim RsInformeMed As New ADODB.Recordset 'Tabla Informe medico
Dim RsVistaPresupuestos As New ADODB.Recordset 'Tabla presupuestos
Dim IDCLI
Dim Pres
Dim NewReg
Dim IdPresup
Dim IdLIdPacP As String
Dim IdLIdInfP As String
Dim IDCLIIDL As String
Dim IdPacP
Dim IdInf As Integer

Sub CargaPre()

    'Blanqueo
    Label10.Visible = False
    NewReg = 1          ' Variable igual al valor UNO, que indica que es un presupuesto para actualizar
    
    If Trim(RsVistaPresupuestos.Fields("diagnotico").Value) <> "" Then Text4.Text = RsVistaPresupuestos.Fields("diagnotico").Value Else Text4.Text = ""
    If RsVistaPresupuestos.Fields("dosis").Value <> "" Then Text12.Text = RsVistaPresupuestos.Fields("dosis").Value Else Text12.Text = ""
    If RsVistaPresupuestos.Fields("dosisd").Value <> "" Then Text11.Text = RsVistaPresupuestos.Fields("dosisd").Value Else Text11.Text = ""
    If RsVistaPresupuestos.Fields("nombrep").Value <> "" Then Text1.Text = RsVistaPresupuestos.Fields("nombrep").Value Else Text1.Text = ""
    If RsVistaPresupuestos.Fields("cedulap").Value <> "" Then Text7.Text = RsVistaPresupuestos.Fields("cedulap").Value Else Text7.Text = ""
    If RsVistaPresupuestos.Fields("apellidop").Value <> "" Then Text2.Text = RsVistaPresupuestos.Fields("apellidop").Value Else Text2.Text = ""
    If RsVistaPresupuestos.Fields("idpaciente").Value <> "" Then IdPacP = RsVistaPresupuestos.Fields("idpaciente").Value
    If RsVistaPresupuestos.Fields("idpaciente").Value <> "" Then IdLIdPacP = RsVistaPresupuestos.Fields("idpaciente").Value
    If RsVistaPresupuestos.Fields("monto").Value <> "" Then Text5.Text = RsVistaPresupuestos.Fields("monto").Value Else Text5.Text = ""
    If RsVistaPresupuestos.Fields("duracion").Value <> "" Then Text6.Text = RsVistaPresupuestos.Fields("duracion").Value Else Text6.Text = ""
    If RsVistaPresupuestos.Fields("tomografia").Value = 1 Then Check1.Value = 1 Else Check1.Value = 0
    If Not IsNull(RsVistaPresupuestos.Fields("cuantas").Value) Then Text14.Text = RsVistaPresupuestos.Fields("cuantas").Value Else Text14.Text = 0
    Label9.Caption = "Presupuesto " & RsVistaPresupuestos.AbsolutePosition & "/" & RsVistaPresupuestos.RecordCount
    Label3.Caption = Format(RsVistaPresupuestos.Fields("Npresupuesto").Value, "000000000#")
    If Not IsNull(RsVistaPresupuestos.Fields("idinforme").Value) Then IdInf = Val(RsVistaPresupuestos.Fields("idinforme").Value) Else IdInf = 0
    Text8.Visible = True
    Combo1.Visible = False
    If RsVistaPresupuestos.Fields("Fecha").Value <> DateTime.Date Then DTPicker1.Value = RsVistaPresupuestos.Fields("Fecha").Value Else DTPicker1.Value = DateTime.Date
    If RsVistaPresupuestos.Fields("IdPresupuesto").Value <> "" Then IdPresup = Val(RsVistaPresupuestos.Fields("IdPresupuesto").Value) Else IdPresup = 0
    If RsVistaPresupuestos.Fields("IdCliente").Value <> "" Then IDCLI = RsVistaPresupuestos.Fields("IdCliente").Value Else IDCLI = 0
    If RsVistaPresupuestos.Fields("IdL").Value <> "" Then IDCLIIDL = RsVistaPresupuestos.Fields("IdL").Value Else IDCLIIDL = IdLDefault
    
    If RsVistaPresupuestos.Fields("razon").Value <> "" Then Text8.Text = RsVistaPresupuestos.Fields("razon").Value Else Text8.Text = ""
    If RsVistaPresupuestos.Fields("direccionC").Value <> "" Then Text9.Text = RsVistaPresupuestos.Fields("direccionC").Value Else Text9.Text = ""
    If RsVistaPresupuestos.Fields("rif").Value <> "" Then Text13.Text = RsVistaPresupuestos.Fields("rif").Value Else Text13.Text = ""
    
    IdLIdPac = RsVistaPresupuestos.Fields("IdLIdPac").Value
    IdLIdInfP = RsVistaPresupuestos.Fields("IdL").Value
    
    BtnGuardarActualizar.Enabled = False
    BtnAnterior1.Enabled = False
    BtnSiguiente1.Enabled = False
    Label21.Caption = "-"
    
End Sub

Sub CargaDia()
    IdInf = Val(RsInformeMed.Fields("IdInforme").Value)
    IdLIdInf = RsInformeMed.Fields("IdL").Value
    Text4.Text = RsInformeMed.Fields("diagnotico").Value
    Text12.Text = RsInformeMed.Fields("dosis").Value
    Text11.Text = RsInformeMed.Fields("dosisd").Value
    If RsInformeMed.Fields("tomografia").Value = 1 Then Check1.Value = 1 Else Check1.Value = 0
    If Not IsNull(RsInformeMed.Fields("cuantas").Value) Then Text14.Text = RsInformeMed.Fields("cuantas").Value Else Text14.Text = 0
    Text6.Text = Round((Val(Text12.Text) / Val(Text11.Text)), 0)
    Text5.Text = ""
    BtnGuardarActualizar.Enabled = False
    Label21.Caption = "Registro " & RsInformeMed.AbsolutePosition & "/" & RsInformeMed.RecordCount
    
    If Val(Text6.Text) < 25 Then
        Label10.Visible = True
    Else
        Label10.Visible = False
    End If
    
    BtnAnterior1.Enabled = True
    BtnSiguiente1.Enabled = True
    
End Sub

Public Sub Blanqueo()
    
    Text1.Text = ""             ' Nombre del paciente
    Text2.Text = ""             ' Apellido del paciente
    Text4.Text = ""             ' Diagnostico
    Text5.Text = ""             ' Costo del tratamiento
    Text6.Text = ""             ' Numero de Sesiones
    Text7.Text = ""             ' Cédula del paciente
    Text11.Text = ""            ' Dosis Diarias
    Text12.Text = ""            ' Total de Dosis
    Text14.Text = ""            ' Nro de tomografias
    Check1.Value = 0            ' El check para tomografias
    Label21.Caption = ""        ' Muestra el numero de Informes médicos
    IdInf = 0                   ' Almacena el IdInforme
    IdPresup = 0                ' Almacena el IdPresupuesto
    
    Text8.Visible = False
    Combo1.Visible = True
    Call ListaCliente
    'BtnBuscarDiagnostico.Enabled = True
    BtnGuardarActualizar.Enabled = False
    
    Label3.Caption = ""         ' Muestra el numero de presupuesto
    
    IdPacP = ""                 ' Variable que almacena el IdPaciente
    Text9.Text = ""             ' Dirección del cliente
    Text13.Text = ""            ' Rif del cliente
    Text8.Text = ""             ' Razón social
    IdLIdInf = IdLDefault       ' Variable que almacena el IdLIdInf
    IdLIdInfP = IdLDefault      ' Variable que almacena el IdL
    
    If RsInformeMed.State Then RsInformeMed.Close
    If RsVistaPresupuestos.State Then RsVistaPresupuestos.Close
    Label9.Caption = "Presupuesto"
    TxtBuscar.Text = ""
End Sub

Private Sub BtnAgregar_Click()
NewReg = 0
Call Blanqueo
End Sub

Private Sub BtnAgregarClientes_Click()
FrmDatosClientes.Show vbModal, FrmPrincipal
Call ListaCliente
End Sub

Private Sub BtnSiguiente_Click()
If RsVistaPresupuestos.State Then
    If RsVistaPresupuestos.RecordCount > 0 Then
        RsVistaPresupuestos.MoveNext
        If RsVistaPresupuestos.EOF Then RsVistaPresupuestos.MoveFirst
        Call CargaPre
    Else
        MsgBox "Estimado usuario, no se encontraron presupuestos para mostrar!", vbExclamation + vbOKOnly, "Disculpe, no hay registros."
    End If
End If
End Sub

Private Sub BtnAnterior_Click()
If RsVistaPresupuestos.State Then
    If RsVistaPresupuestos.RecordCount > 0 Then
        RsVistaPresupuestos.MovePrevious
        If RsVistaPresupuestos.BOF Then RsVistaPresupuestos.MoveLast
        Call CargaPre
    Else
        MsgBox "Estimado usuario, no se encontraron presupuestos para mostrar!", vbExclamation + vbOKOnly, "Disculpe, no hay registros."
        Label3.Caption = "-"
    End If
End If
End Sub

Private Sub BtnBuscar_Click()
On Error GoTo ErM

TxtBuscar.Text = Replace(TxtBuscar.Text, "'", "")

If Trim(TxtBuscar.Text) = "" Or Trim(TxtBuscar.Text) = "Busqueda" Then
    CSql = "Select * From Presupuesto1 ORDER BY NPresupuesto"
Else
    CSql = "Select * From Presupuesto1 Where (CedulaP = '" & TxtBuscar.Text & "' Or Npresupuesto =" & Val(TxtBuscar.Text) & ") And Estado='1' Order by Fecha Desc"
End If

BtnGuardarActualizar.Enabled = False

If RsVistaPresupuestos.State Then RsVistaPresupuestos.Close
Set RsVistaPresupuestos = CrearRS(CSql)

    If Not (RsVistaPresupuestos.EOF) Then
        NewReg = 1  ' Variable igual al valor UNO, que indica que es un presupuesto para actualizar
        RsVistaPresupuestos.MoveFirst
        
        Call CargaPre
    Else
        MsgBox "No Existe presupuesto coincidente con la clave de busqueda solicitada", vbOKOnly + vbCritical, "No Existe registro alguno"

    End If

'BtnBuscarDiagnostico.Enabled = False
Exit Sub
ErM:
    MsgBox "Estimado usuario, corrija el texto a buscar.", vbExclamation + vbOKOnly, "Disculpe."
End Sub

Public Sub BtnBuscarClientes_Click()

IDCLI = 0
IDCLIIDL = IdLDefault
Text13.Text = Replace(Text13.Text, "'", "")

If Trim(Text13.Text) = "" Then
    MsgBox "Estimado Usuario, Debe ingresar algun dato en la caja de texto referida al RIF para realizar la busqueda!", vbOKOnly + vbExclamation, "Disculpe, faltan datos."
    Exit Sub
End If

BtnGuardarActualizar.Enabled = False

CSql = "Select * From Presupuesto1 Where Rif ='" & Text13.Text & "' ORDER BY NPresupuesto"
If RsVistaPresupuestos.State Then RsVistaPresupuestos.Close
Set RsVistaPresupuestos = CrearRS(CSql)

If Not (RsVistaPresupuestos.EOF) Then

    RsVistaPresupuestos.MoveFirst

    Call CargaPre
    'BtnBuscarDiagnostico.Enabled = False
Else
    CSql = "Select * From Cliente Where Rif ='" & Text13.Text & "' ORDER BY IdCliente"

    Set RsTemp = CrearRS(CSql)

    If RsTemp.RecordCount = 0 Then
        MsgBox "Estimado usuario, El RIF no se encuentra registrado!", vbOKOnly + vbExclamation, "Disculpe, no se encontro el RIF."
    Else
        MsgBox "Estimado usuario, El RIF ingresado para la busqueda cuya razón social es '" & RsTemp.Fields("Razon").Value & "' " & Chr(13) & Chr(10) & "No contiene presupuestos para pacientes!", vbOKOnly + vbExclamation, "Disculpe, no se encontro el registro."
    End If
    If RsVistaPresupuestos.State Then RsVistaPresupuestos.Close
    'BtnBuscarDiagnostico.Enabled = True
    Combo1.ListIndex = -1
    Text8.Text = ""
    Text9.Text = ""
    Blanqueo
    IDCLI = 0
    IDCLIIDL = IdLDefault
End If


End Sub

Private Sub BtnBuscarDiagnostico_Click()

FrmCedulaPaciente.Show vbModal, FrmPrincipal

IdPacP = ""
IdLIdPacP = IdLDefault
IdLIdInf = IdLDefault       ' Variable que almacena el IdLIdInf
IdLIdInfP = IdLDefault      ' Variable que almacena el IdL
IdInf = 0                  ' Variable que almacena el IdInforme

If FrmCedulaPaciente.Ced = "-1" Then Exit Sub

NewReg = 0
Blanqueo

If Val(FrmCedulaPaciente.Ced) = 0 Then
    Exit Sub
End If

CSql = "Select * From Paciente Where CedulaP = '" & FrmCedulaPaciente.Ced & "'"
Set RsTemp = CrearRS(CSql)

If Not (RsTemp.EOF) Then

    Text1.Text = RsTemp.Fields("NombreP").Value
    Text7.Text = RsTemp.Fields("CedulaP").Value
    Text2.Text = RsTemp.Fields("ApellidoP").Value
    IdPacP = RsTemp.Fields("IdPaciente").Value
    IdLIdPacP = RsTemp.Fields("IdL").Value
    IDCLI = 0
    IDCLIIDL = IdLDefault
    
    If RsInformeMed.State Then RsInformeMed.Close
    
    CSql = "select IdInforme, Diagnotico, Dosis, DosisD, Tomografia, Cuantas, IdL From Informe_Medico Where IdPaciente = '" & IdPacP & "' And IdLIdPac='" & IdLIdPacP & "' And Estado ='1' Order by Fecha Desc"
    Set RsInformeMed = CrearRS(CSql)
    
    If Not (RsInformeMed.EOF) Then
        RsInformeMed.MoveFirst
        Call CargaDia
    Else
        MsgBox "Este paciente no presenta registro en la tabla de Informes Medicos", vbOKOnly + vbInformation, "No tiene informe medico"
        BtnGuardarActualizar.Enabled = False
        Text1.Text = ""
        Text7.Text = ""
        Text2.Text = ""
        IdPacP = ""
        IdLIdInf = IdLDefault
        IdInf = 0
        Label10.Visible = False
    End If
Else
    
    MsgBox "No existe esa cedula en la tabla pacientes" & Chr(13) & Chr(13) & "   " & FrmCedulaPaciente.Ced, vbOKOnly + vbInformation, "Disculpe, No existe el registro."
    BtnGuardarActualizar.Enabled = False
    
End If


RsTemp.Close

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Sub EnviarRegPendiente(ByVal IdNuevo2 As Integer, ByVal IdLIdInf2 As String)
Dim i As Integer
Dim MaxRegP As String
Dim StrSen As String
Dim RsRegPendiente As ADODB.Recordset

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("IdMax").Value) Then
    MaxRegP = RsTemp.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Presupuesto WHERE IdPresupuesto='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Presupuesto (["
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
RsRegPendiente.Fields("Modulo").Value = "Presupuesto"
RsRegPendiente.Fields("Tabla").Value = "Presupuesto"
RsRegPendiente.Fields("Condicional").Value = " IdPresupuesto='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update

MsgBox "Envio Al Servidor Web Satisfactorio!!!", vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Private Sub BtnGuardarActualizar_Click()
Dim NewId

If Trim(Text5.Text) = "" Or Trim(Text7.Text) = "" Or Trim(Text13.Text) = "" Then
    MsgBox "Estimado usuario, Debe llenar todos los campos!", vbExclamation + vbOKOnly, "Disculpe."
    Exit Sub
End If

If Val(IdInf) = 0 Then
    MsgBox "Estimado usuario, el paciente seleccionado no contiene un informe médico!", vbExclamation + vbOKOnly, "Disculpe."
    Exit Sub
End If

If IDCLI = 0 Then
    MsgBox "Estimado usuario, debe seleccionar un Cliente al cual presupuestar!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Bloque que verifica si hay internet
  If Not Verificar_Internet Then
      NuevoIdL = IdL
  Else
      NuevoIdL = IdLDefault
  End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Select Case NewReg

    Case Is = 0
        
        CSql = "Select Max(IdPresupuesto) + 1 as NuevoId From Presupuesto"
        Set RsTemp = CrearRS(CSql)
            
        If Not IsNull(RsTemp.Fields("NuevoId")) Then
            NewId = RsTemp.Fields("NuevoId").Value
        Else
            NewId = "1"
        End If
        Set RsTemp = Nothing
        
        CSql = "Select Max(NPresupuesto)+1 as NuevoId From Presupuesto"
        Set RsTemp = CrearRS(CSql)
        
        If Not IsNull(RsTemp.Fields("NuevoId")) Then
            Pres = Format(RsTemp.Fields("NuevoId").Value, "000000000#")
        Else
            Pres = Format(1, "000000000#")
        End If
        RsTemp.Close

        
        CSql = "Select * From Presupuesto"
        Set RsTemp = CrearRS(CSql)
        
        IdLIdInfP = NuevoIdL
        RsTemp.AddNew
        RsTemp.Fields("IdPresupuesto").Value = NewId
        RsTemp.Fields("IdL").Value = IdLIdInfP
        RsTemp.Fields("IDpaciente").Value = IdPacP
        RsTemp.Fields("IdLIdPac").Value = IdLIdPacP
        RsTemp.Fields("IdInforme").Value = IdInf
        RsTemp.Fields("IdLIdInf").Value = IdLIdInf
        
        RsTemp.Fields("monto").Value = Text5.Text
        RsTemp.Fields("duracion").Value = Text6.Text
        RsTemp.Fields("fecha").Value = Format(Date, "dd/mm/yyyy")
        RsTemp.Fields("idcliente").Value = IDCLI
        'RsTemp.Fields("IdLIdCliente").Value = IDCLIIDL
        RsTemp.Fields("idproducto").Value = 1
        RsTemp.Fields("idusuario").Value = IdUser
        RsTemp.Fields("cantidad").Value = 1
        RsTemp.Fields("NPresupuesto").Value = Val(Pres)
        RsTemp.Update
        
        EnviarRegPendiente NewId, NuevoIdL
        
        MsgBox "Registro Agregado satisfactoriamente", vbOKOnly + vbInformation, "Guardado Exitoso"
        
    Case Is = 1
        'CSql = "Select * From Presupuesto Where NPresupuesto='" & Val(Label3.Caption) & "' And IdPaciente='" & IdPacP & "' And Idpresupuesto='" & Val(IdPresup) & "' And IdLIdPac='" & IdLIdPac & "' And IdLIdInf='" & IdLIdInf & "' And IdInforme='" & Val(IdInf) & "'"
        CSql = "Select * From Presupuesto Where IdPresupuesto='" & Val(IdPresup) & "' And IdL='" & IdLIdInfP & "'"
        
        Set RsTemp = CrearRS(CSql)
        If RsTemp.RecordCount > 0 Then
            
            Dim NuevoI As String
            NewId = RsTemp.Fields("IdPresupuesto").Value
            NuevoI = RsTemp.Fields("IdL").Value
            
            RsTemp.Fields("monto").Value = Text5.Text
            RsTemp.Fields("duracion").Value = Text6.Text
            RsTemp.Fields("fecha").Value = Format(Date, "dd/mm/yyyy")
            RsTemp.Fields("idcliente").Value = IDCLI
            ''RsTemp.Fields("IdLIdCliente").Value = IDCLIIDL
            RsTemp.Fields("idproducto").Value = 1
            RsTemp.Fields("idusuario").Value = IdUser
            RsTemp.Fields("cantidad").Value = 1
            'RsTemp.Fields("NPresupuesto").Value = Val(Pres)
            RsTemp.Update
            
            EnviarRegPendiente NewId, NuevoI
            
            MsgBox "Registro Actualizado satisfactoriamente", vbOKOnly + vbInformation, "Guardado Exitoso"
        End If
End Select

BtnGuardarActualizar.Enabled = False

CargarAgregadoModificado
End Sub

Sub CargarAgregadoModificado()

'BtnGuardarActualizar.Visible = False
CSql = "Select * From Presupuesto1 Where IdPaciente='" & IdPacP & "' And Estado='1' ORDER BY NPresupuesto"
If RsVistaPresupuestos.State Then RsVistaPresupuestos.Close
Set RsVistaPresupuestos = CrearRS(CSql)

    If Not (RsVistaPresupuestos.EOF) Then
        NewReg = 1  ' Variable igual al valor UNO, que indica que es un presupuesto para actualizar
        RsVistaPresupuestos.MoveFirst
        Call CargaPre
    Else
        MsgBox "No Existe presupuesto coincidente con la clave de busqueda solicitada", vbOKOnly, "No Existe registro alguno"
    End If
'BtnBuscarDiagnostico.Enabled = False
End Sub

Private Sub BtnImprimir_Click()
If Text13.Text = "" Then
    MsgBox "Seleccione al cliente", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If
If Text7.Text = "" Then
    MsgBox "Seleccione a un paciente", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If

If IdPresup = 0 Then Exit Sub
'========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\PresupuestoN.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    
    ' Código anterior...
    '.SelectionFormula = "{PRESUPUESTO1.Rif}='" & Text13.Text & "' And {PRESUPUESTO1.CEDULAP}=" & Text7.Text & " And {PRESUPUESTO1.IdPresupuesto}=" & IdPresup & " And {PRESUPUESTO1.IdInforme}=" & IdInf & ""
    
    ' Luego de modificar la vista y el reporte....
    .SelectionFormula = "{PRESUPUESTO1.IdPresupuesto}=" & IdPresup & " And {PRESUPUESTO1.IdL}='" & IdLIdInfP & "'"
    '.SelectionFormula = "{PRESUPUESTO1.IdPresupuesto}=" & IdPresup & " And {PRESUPUESTO1.IdL}='S'"
    .WindowTitle = "Presupuesto de Tratamiento N° " & Label3.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

Exit Sub
ERRT:
CrystalReport1.Action = 1

End Sub

Private Sub BtnListadoProductos_Click()
On Error Resume Next
ModulO = 2
FrmListadoClientes.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnSiguiente1_Click()
If RsInformeMed.State Then
    If RsInformeMed.RecordCount > 0 Then
        RsInformeMed.MoveNext
        If RsInformeMed.EOF Then RsInformeMed.MoveFirst
        Call CargaDia
    Else
        MsgBox "Estimado usuario, no se encontraron informes médicos para el paciente seleccionado!", vbExclamation + vbOKOnly, "Disculpe, el paciente no tiene un informe médico."
        Label21.Caption = "-"
    End If
End If
End Sub

Private Sub BtnAnterior1_Click()
If RsInformeMed.State Then
    If RsInformeMed.RecordCount > 0 Then
        RsInformeMed.MovePrevious
        If RsInformeMed.BOF Then RsInformeMed.MoveLast
        Call CargaDia
    Else
        MsgBox "Estimado usuario, no se encontraron informes médicos para el paciente seleccionado!", vbExclamation + vbOKOnly, "Disculpe, el paciente no tiene un informe médico."
        Label21.Caption = "-"
    End If
End If
End Sub

Public Sub Combo1_Click()

If Combo1.ListIndex = -1 Then
    Exit Sub
End If

CSql = "Select * From Cliente Where IdCliente = '" & Combo1.ItemData(Combo1.ListIndex) & "'"
Set RsTemp = CrearRS(CSql)
If Not (RsTemp.EOF) Then
    Text9.Text = RsTemp.Fields("direccionc").Value
    Text13.Text = RsTemp.Fields("rif").Value
    IDCLI = RsTemp.Fields("IdCliente").Value
    'IDCLIIDL = RsTemp.Fields("IdL").Value
End If
RsTemp.Close

End Sub

Sub ListaCliente()
Combo1.Clear

CSql = "SELECT * FROM Cliente"
Set RsTemp = CrearRS(CSql)

If Not (RsTemp.EOF) Then RsTemp.MoveFirst

Do While Not RsTemp.EOF
    Combo1.AddItem UCase(RsTemp.Fields("Razon").Value)
    Combo1.ItemData(Combo1.NewIndex) = RsTemp.Fields("IdCliente").Value
    RsTemp.MoveNext
Loop
RsTemp.Close

Text9.Text = ""
Text8.Text = ""
Text13.Text = ""

End Sub

Private Sub Form_Load()
Centrar Me

Text8.Visible = False
Combo1.Visible = True

CSql = "Select * From Presupuesto1 ORDER BY NPresupuesto"
Set RsVistaPresupuestos = CrearRS(CSql)

NoPresup
ListaCliente

End Sub

Sub NoPresup()

CSql = "Select Max(NPresupuesto)+1 as NuevoId From Presupuesto"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId")) Then
    Pres = Format(RsTemp.Fields("NuevoId").Value, "000000000#")
Else
    Pres = Format(1, "000000000#")
End If
RsTemp.Close
Label3.Caption = Pres
             
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnBuscarClientes_Click
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_GotFocus()
If TxtBuscar.Text = UCase("busqueda") Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
Select Case KeyAscii
    Case 48 To 57 ' permite el ingreso de numeros
    Case Is = 13 ' permite presionar el ENTER
    Call BtnBuscar_Click
    Case Is = 8 ' Permite Borrar de retroceso
    Case Else ' Inhibe todas las demas teclas
End Select
End Sub
Private Sub Text5_Change()
BtnGuardarActualizar.Enabled = True
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Val(Text6.Text) < 25 Then
    If Val(Text14.Text) > 0 Then
        Label10.Visible = True
    Else
        Label10.Visible = False
    End If
End If
Select Case KeyAscii
Case 48 To 57
Exit Sub
Case Is = 13
Exit Sub
Case Is = 8
Exit Sub
Case Else
KeyAscii = 0
End Select

End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPresupuestoTratamientos2 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Cliente"
         Height          =   1935
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   10575
         Begin VB.TextBox Text9 
            Height          =   615
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   1200
            Width           =   6975
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   1200
            TabIndex        =   50
            Top             =   310
            Width           =   1935
         End
         Begin VB.TextBox Text8 
            Height          =   405
            Left            =   1200
            TabIndex        =   49
            Top             =   720
            Width           =   6975
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Height          =   1695
            Left            =   8640
            TabIndex        =   43
            Top             =   120
            Width           =   1815
            Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
               Height          =   495
               Left            =   960
               TabIndex        =   44
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
               MICON           =   "FrmPresupuestoTratamientos2.frx":0000
               PICN            =   "FrmPresupuestoTratamientos2.frx":001C
               PICH            =   "FrmPresupuestoTratamientos2.frx":02B2
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
               TabIndex        =   45
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
               MICON           =   "FrmPresupuestoTratamientos2.frx":0511
               PICN            =   "FrmPresupuestoTratamientos2.frx":052D
               PICH            =   "FrmPresupuestoTratamientos2.frx":07C2
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
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. Presupuesto"
               Height          =   195
               Left            =   270
               TabIndex        =   47
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Height          =   375
               Left            =   120
               TabIndex        =   46
               Top             =   480
               Width           =   1575
            End
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarClientes 
            Height          =   375
            Left            =   4560
            TabIndex        =   48
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":0A1E
            PICN            =   "FrmPresupuestoTratamientos2.frx":0A3A
            PICH            =   "FrmPresupuestoTratamientos2.frx":0BC7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarClientes 
            Height          =   375
            Left            =   3240
            TabIndex        =   53
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":0D71
            PICN            =   "FrmPresupuestoTratamientos2.frx":0D8D
            PICH            =   "FrmPresupuestoTratamientos2.frx":11B5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   840
            Width           =   6975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RIF.:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   400
            Width           =   345
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1290
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   825
            Width           =   990
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   3735
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   10575
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tomografía"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   24
            Top             =   2460
            Width           =   1215
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
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
            TabIndex        =   23
            Top             =   2400
            Width           =   615
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            Top             =   3120
            Width           =   1815
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   3240
            TabIndex        =   21
            Top             =   3120
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1080
            TabIndex        =   20
            Top             =   1320
            Width           =   4815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1080
            TabIndex        =   19
            Top             =   840
            Width           =   4815
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1080
            TabIndex        =   18
            Top             =   1800
            Width           =   6975
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   1080
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
         Begin VB.Timer Timer1 
            Left            =   7440
            Top             =   2280
         End
         Begin VB.TextBox TxtIdPresupuesto 
            Height          =   375
            Left            =   7200
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtIdInforme 
            Height          =   375
            Left            =   7200
            TabIndex        =   13
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5400
            TabIndex        =   12
            Top             =   3120
            Width           =   1065
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarDiagnostico 
            Height          =   855
            Left            =   8760
            TabIndex        =   15
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":144E
            PICN            =   "FrmPresupuestoTratamientos2.frx":146A
            PICH            =   "FrmPresupuestoTratamientos2.frx":1892
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6720
            TabIndex        =   25
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   47775745
            CurrentDate     =   39834
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   6600
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
            Left            =   7080
            Top             =   3120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente1 
            Height          =   495
            Left            =   9600
            TabIndex        =   26
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":1B2B
            PICN            =   "FrmPresupuestoTratamientos2.frx":1B47
            PICH            =   "FrmPresupuestoTratamientos2.frx":1DDD
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
            Left            =   9000
            TabIndex        =   27
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":203C
            PICN            =   "FrmPresupuestoTratamientos2.frx":2058
            PICH            =   "FrmPresupuestoTratamientos2.frx":22ED
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   8640
            TabIndex        =   41
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico/Tratamiento"
            Height          =   195
            Left            =   8595
            TabIndex        =   40
            Top             =   480
            Width           =   1785
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            Height          =   195
            Left            =   3000
            TabIndex        =   39
            Top             =   3225
            Width           =   75
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Dosis"
            Height          =   255
            Left            =   1080
            TabIndex        =   38
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Dosis Diarias"
            Height          =   255
            Left            =   3240
            TabIndex        =   37
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            Height          =   195
            Left            =   5160
            TabIndex        =   36
            Top             =   3210
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1410
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   435
            Width           =   540
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1890
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   3210
            Width           =   690
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Costo del Tratamiento:"
            Height          =   555
            Left            =   120
            TabIndex        =   29
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Sesiones"
            Height          =   195
            Left            =   5400
            TabIndex        =   28
            Top             =   2880
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   6000
         Width           =   3135
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el Número de Cédula o Número de Presupuesto a buscar"
            Top             =   240
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   1680
            TabIndex        =   10
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":2549
            PICN            =   "FrmPresupuestoTratamientos2.frx":2565
            PICH            =   "FrmPresupuestoTratamientos2.frx":27CA
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
         Left            =   3360
         TabIndex        =   1
         Top             =   6000
         Width           =   7335
         Begin ChamaleonButton.ChameleonBtn BtnImprimir 
            Height          =   375
            Left            =   3240
            TabIndex        =   2
            ToolTipText     =   "Reporte"
            Top             =   240
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":2A5C
            PICN            =   "FrmPresupuestoTratamientos2.frx":2A78
            PICH            =   "FrmPresupuestoTratamientos2.frx":2B9D
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
            TabIndex        =   3
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":2E2D
            PICN            =   "FrmPresupuestoTratamientos2.frx":2E49
            PICH            =   "FrmPresupuestoTratamientos2.frx":3012
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
            TabIndex        =   4
            ToolTipText     =   "Guardar / Actualizar"
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":3247
            PICN            =   "FrmPresupuestoTratamientos2.frx":3263
            PICH            =   "FrmPresupuestoTratamientos2.frx":34F2
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
            Left            =   2520
            TabIndex        =   5
            ToolTipText     =   "Agregar"
            Top             =   240
            Visible         =   0   'False
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":3933
            PICN            =   "FrmPresupuestoTratamientos2.frx":394F
            PICH            =   "FrmPresupuestoTratamientos2.frx":3ADC
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
            TabIndex        =   6
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":3D11
            PICN            =   "FrmPresupuestoTratamientos2.frx":3D2D
            PICH            =   "FrmPresupuestoTratamientos2.frx":400F
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
            TabIndex        =   7
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
            MICON           =   "FrmPresupuestoTratamientos2.frx":4260
            PICN            =   "FrmPresupuestoTratamientos2.frx":427C
            PICH            =   "FrmPresupuestoTratamientos2.frx":4420
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
   End
End
Attribute VB_Name = "FrmPresupuestoTratamientos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

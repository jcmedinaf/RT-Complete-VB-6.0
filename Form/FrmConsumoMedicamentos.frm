VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConsumoMedicamentos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo de Medicamentos"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13815
   Icon            =   "FrmConsumoMedicamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13815
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13575
      Begin VB.TextBox TxtNoInsumos 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   6360
         Width           =   855
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7223
         Object.Width           =   13305
         Object.Height          =   4065
         Editable        =   -1  'True
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   1815
         Left            =   10560
         TabIndex        =   17
         Top             =   240
         Width           =   2895
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   720
            Top             =   240
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
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Procesada"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1200
            TabIndex        =   18
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   51118083
            CurrentDate     =   40189
         End
         Begin ChamaleonButton.ChameleonBtn BtnProcesada 
            Height          =   375
            Left            =   1440
            TabIndex        =   33
            ToolTipText     =   "Procesar Consumos de Medicamentos"
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Procesada"
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
            MICON           =   "FrmConsumoMedicamentos.frx":1002
            PICN            =   "FrmConsumoMedicamentos.frx":101E
            PICH            =   "FrmConsumoMedicamentos.frx":1293
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
            Caption         =   "No Consumo:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   798
            Width           =   960
         End
         Begin VB.Label LblNoConsumo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   350
            Left            =   1200
            TabIndex        =   20
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   330
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos de la Solicitud"
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   10335
         Begin VB.TextBox TxtDiagnostico 
            Height          =   495
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1200
            Width           =   9135
         End
         Begin VB.TextBox TxtEdad 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox TxtApellido 
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox TxtCedula 
            Height          =   375
            Left            =   1080
            TabIndex        =   2
            Top             =   270
            Width           =   1935
         End
         Begin VB.TextBox TxtNoHistoria 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   270
            Width           =   1815
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarPaciente 
            Height          =   375
            Left            =   3120
            TabIndex        =   14
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
            MICON           =   "FrmConsumoMedicamentos.frx":150F
            PICN            =   "FrmConsumoMedicamentos.frx":152B
            PICH            =   "FrmConsumoMedicamentos.frx":1790
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnListaPaciente 
            Height          =   375
            Left            =   4440
            TabIndex        =   38
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Listado Pacientes"
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
            MICON           =   "FrmConsumoMedicamentos.frx":1A22
            PICN            =   "FrmConsumoMedicamentos.frx":1A3E
            PICH            =   "FrmConsumoMedicamentos.frx":1CC7
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
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   4560
            TabIndex        =   25
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cédula:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Historia:"
            Height          =   195
            Left            =   7440
            TabIndex        =   16
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edad:"
            Height          =   195
            Left            =   8880
            TabIndex        =   15
            Top             =   810
            Width           =   420
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   7
         Top             =   6840
         Width           =   9255
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   8160
            TabIndex        =   8
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
            MICON           =   "FrmConsumoMedicamentos.frx":20E2
            PICN            =   "FrmConsumoMedicamentos.frx":20FE
            PICH            =   "FrmConsumoMedicamentos.frx":22C7
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
            TabIndex        =   9
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
            MICON           =   "FrmConsumoMedicamentos.frx":24FC
            PICN            =   "FrmConsumoMedicamentos.frx":2518
            PICH            =   "FrmConsumoMedicamentos.frx":27A7
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
            TabIndex        =   10
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
            MICON           =   "FrmConsumoMedicamentos.frx":2BE8
            PICN            =   "FrmConsumoMedicamentos.frx":2C04
            PICH            =   "FrmConsumoMedicamentos.frx":2D91
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
            MICON           =   "FrmConsumoMedicamentos.frx":2FC6
            PICN            =   "FrmConsumoMedicamentos.frx":2FE2
            PICH            =   "FrmConsumoMedicamentos.frx":32C4
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
            Left            =   4560
            TabIndex        =   12
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
            MICON           =   "FrmConsumoMedicamentos.frx":3515
            PICN            =   "FrmConsumoMedicamentos.frx":3531
            PICH            =   "FrmConsumoMedicamentos.frx":3656
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
            TabIndex        =   31
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
            MICON           =   "FrmConsumoMedicamentos.frx":38E6
            PICN            =   "FrmConsumoMedicamentos.frx":3902
            PICH            =   "FrmConsumoMedicamentos.frx":3AA6
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   6840
         Width           =   3975
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   1080
            Top             =   240
         End
         Begin VB.TextBox TxtBuscar 
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
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el número del consumo de medicamentos"
            Top             =   240
            Width           =   1420
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2640
            TabIndex        =   5
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmConsumoMedicamentos.frx":3C45
            PICN            =   "FrmConsumoMedicamentos.frx":3C61
            PICH            =   "FrmConsumoMedicamentos.frx":3EC6
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
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Consumo:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   330
            Width           =   930
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregarRenglon 
         Height          =   375
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Agregar"
         Top             =   6360
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
         MICON           =   "FrmConsumoMedicamentos.frx":4158
         PICN            =   "FrmConsumoMedicamentos.frx":4174
         PICH            =   "FrmConsumoMedicamentos.frx":4301
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrarRenglon 
         Height          =   375
         Left            =   1200
         TabIndex        =   35
         ToolTipText     =   "Eliminar"
         Top             =   6360
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
         MICON           =   "FrmConsumoMedicamentos.frx":4536
         PICN            =   "FrmConsumoMedicamentos.frx":4552
         PICH            =   "FrmConsumoMedicamentos.frx":46F6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Insumos:"
         Height          =   195
         Left            =   11400
         TabIndex        =   37
         Top             =   6450
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmConsumoMedicamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsBuscarPacientes As New ADODB.Recordset
Dim RsIdMax As New ADODB.Recordset
Dim IdMax, NewReg
Dim RsConsumoMedicamento As New ADODB.Recordset
Dim RsRenglonConsumoMedicamento As New ADODB.Recordset
Dim IdPaciente
Private Sub BtnAgregar_Click()
On Error Resume Next
NewReg = 1


BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnProcesada.Enabled = False
BtnAgregarRenglon.Enabled = True
BtnBorrarRenglon.Enabled = True

'CSql = "Select Max(IdConsumo) + 1 as MaxId From ConsumoMedicamento"
'Set RsIdMax = CrearRS(CSql)
'
'If RsIdMax.RecordCount > 0 Then
'    LblNoConsumo.Caption = Format(RsIdMax.Fields("MaxId").Value, "0000#")
'Else
'    LblNoConsumo.Caption = Format(1, "0000#")
'End If


End Sub

Private Sub BtnAnterior_Click()

End Sub

Private Sub BtnAgregarRenglon_Click()
On Error Resume Next
If DMGrid1.Rows >= 0 Then
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.DColumnas(2).Locked = True
    DMGrid1.DColumnas(3).Locked = True
    DMGrid1.EditActive = True
    DMGrid1.PaintMGrid
    TxtNoInsumos.Text = DMGrid1.Rows
End If
End Sub

Private Sub BtnBorrarRenglon_Click()
On Error Resume Next
If DMGrid1.Rows > 0 Then
    DMGrid1.Rows = DMGrid1.Rows - 1
    DMGrid1.PaintMGrid
    TxtNoInsumos.Text = DMGrid1.Rows
End If
End Sub

Private Sub BtnBuscar_Click()
If TxtBuscar.Text <> "" Or TxtBuscar.Text <> "Busqueda" Then
    NewReg = 2
    CSql = "SELECT Paciente.IdPaciente, Paciente.Historia, Paciente.CedulaP, Paciente.NombreP, Paciente.ApellidoP, Paciente.EdadP, " & _
           "ConsumoMedicamento.FechaConsumo, ConsumoMedicamento.IdConsumo, RenglonConsumoMedicamento.Codigo, ConsumoMedicamento.Procesada," & _
           "RenglonConsumoMedicamento.Descripcion, RenglonConsumoMedicamento.Cantidad, Informe_Medico.Diagnotico " & _
           "FROM ConsumoMedicamento INNER JOIN RenglonConsumoMedicamento ON ConsumoMedicamento.IdConsumo = dbo.RenglonConsumoMedicamento.IdConsumo INNER JOIN " & _
           "Paciente ON ConsumoMedicamento.IdPaciente = Paciente.IdPaciente INNER JOIN " & _
           "Informe_Medico ON Paciente.IdPaciente = Informe_Medico.IdPaciente " & _
           "Where ConsumoMedicamento.IdConsumo ='" & TxtBuscar.Text & "'"
    Set RsConsumoMedicamento = CrearRS(CSql)
    
    If RsConsumoMedicamento.RecordCount = 0 Then Exit Sub
    
    If RsConsumoMedicamento.RecordCount > 0 Then
        IdMax = RsConsumoMedicamento.Fields("IdConsumo").Value
        IdPaciente = RsConsumoMedicamento.Fields("IdPaciente").Value
        TxtNoHistoria.Text = RsConsumoMedicamento.Fields("Historia").Value
        TxtCedula.Text = RsConsumoMedicamento.Fields("CedulaP").Value
        TxtNombre.Text = RsConsumoMedicamento.Fields("NombreP").Value
        TxtApellido.Text = RsConsumoMedicamento.Fields("ApellidoP").Value
        TxtEdad.Text = RsConsumoMedicamento.Fields("EdadP").Value
        DTPicker1.Value = Format(RsConsumoMedicamento.Fields("FechaConsumo").Value, "dd/mm/yyyy")
        LblNoConsumo.Caption = Format(RsConsumoMedicamento.Fields("IdConsumo").Value, "0000")
        TxtDiagnostico.Text = RsConsumoMedicamento.Fields("Diagnotico").Value
        
        
        If RsConsumoMedicamento.Fields("Procesada").Value = True Then
            Check1.Value = 1
            BtnAgregar.Enabled = False
            BtnGuardarActualizar.Enabled = False
            BtnEliminar.Enabled = False
            BtnImprimir.Enabled = True
            BtnProcesada.Enabled = False
            BtnAgregarRenglon.Enabled = False
            BtnBorrarRenglon.Enabled = False
        Else
            Check1.Value = 0
            BtnAgregar.Enabled = True
            BtnGuardarActualizar.Enabled = True
            BtnEliminar.Enabled = True
            BtnImprimir.Enabled = True
            BtnProcesada.Enabled = True
            BtnAgregarRenglon.Enabled = True
            BtnBorrarRenglon.Enabled = True
        End If
        
        
        
        CSql = "Select * From RenglonConsumoMedicamento Where IdConsumo ='" & TxtBuscar.Text & "'"
        Set RsConsumoMedicamento = CrearRS(CSql)
        
        i = 1
        Do While Not RsConsumoMedicamento.EOF
       
        'For i = 1 To RsConsumoMedicamento.RecordCount
            DMGrid1.Rows = i
            DMGrid1.ValorCelda(i, 1) = RsConsumoMedicamento.Fields("Codigo").Value
            DMGrid1.ValorCelda(i, 2) = RsConsumoMedicamento.Fields("Descripcion").Value
            DMGrid1.ValorCelda(i, 3) = RsConsumoMedicamento.Fields("Cantidad").Value
            i = i + 1
            RsConsumoMedicamento.MoveNext
        'Next i
        Loop
        DMGrid1.PaintMGrid
    
    End If
    TxtNoInsumos.Text = DMGrid1.Rows
End If



End Sub

Public Sub BtnBuscarPaciente_Click()
On Error Resume Next
If TxtCedula.Text <> "" Then
    CSql = "SELECT  Paciente.IdPaciente, Paciente.Historia, Paciente.CedulaP, Paciente.NombreP, Paciente.ApellidoP, Paciente.EdadP, Informe_Medico.Diagnotico " & _
           "FROM Informe_Medico INNER JOIN Paciente ON Informe_Medico.IdPaciente = Paciente.IdPaciente Where CedulaP='" & TxtCedula.Text & "'"
    Set RsBuscarPacientes = CrearRS(CSql)
    If RsBuscarPacientes.RecordCount > 0 Then
        IdPaciente = RsBuscarPacientes.Fields("IdPaciente").Value
        TxtCedula.Text = RsBuscarPacientes.Fields("CedulaP").Value
        TxtApellido.Text = RsBuscarPacientes.Fields("ApellidoP").Value
        TxtNoHistoria.Text = RsBuscarPacientes.Fields("Historia").Value
        TxtNombre.Text = RsBuscarPacientes.Fields("NombreP").Value
        TxtEdad.Text = RsBuscarPacientes.Fields("EdadP").Value
        TxtDiagnostico.Text = RsBuscarPacientes.Fields("Diagnotico").Value
    End If
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnProcesada.Enabled = False
BtnAgregarRenglon.Enabled = False
BtnBorrarRenglon.Enabled = False
Blanqueo
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next
MsgBox "aun no se ha codificado!", vbInformation
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
If TxtCedula.Text = "" Then
    TxtCedula.SetFocus
    Exit Sub
End If

Select Case NewReg

Case Is = 1

    CSql = "Select Max(IdConsumo) + 1 as MaxId From ConsumoMedicamento"
    Set RsIdMax = CrearRS(CSql)
    If RsIdMax.EOF Then
        IdMax = RsIdMax.Fields("MaxId").Value
    Else
        IdMax = 1
    End If
    
    
    CSql = "Select * From ConsumoMedicamento"
    Set RsConsumoMedicamento = CrearRS(CSql)
    
    RsConsumoMedicamento.AddNew
    RsConsumoMedicamento.Fields("IdConsumo").Value = IdMax
    RsConsumoMedicamento.Fields("IdPaciente").Value = IdPaciente
    RsConsumoMedicamento.Fields("FechaConsumo").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
    RsConsumoMedicamento.Fields("Procesada").Value = Check1.Value
    RsConsumoMedicamento.Fields("IdUser").Value = IdUser
    RsConsumoMedicamento.Update
    
    CSql = "Select * From RenglonConsumoMedicamento"
    Set RsRenglonConsumoMedicamento = CrearRS(CSql)
    
    For i = 1 To DMGrid1.Rows
            
        b1 = DMGrid1.ValorCelda(i, 1)
        b2 = DMGrid1.ValorCelda(i, 2)
        b3 = DMGrid1.ValorCelda(i, 3)
    
        RsRenglonConsumoMedicamento.AddNew
        RsRenglonConsumoMedicamento.Fields("IdConsumo").Value = IdMax
        RsRenglonConsumoMedicamento.Fields("IdPaciente").Value = IdPaciente
        RsRenglonConsumoMedicamento.Fields("Codigo").Value = b1
        RsRenglonConsumoMedicamento.Fields("Descripcion").Value = b2
        RsRenglonConsumoMedicamento.Fields("Cantidad").Value = b3
        RsRenglonConsumoMedicamento.Fields("IdUser").Value = IdMax
        RsRenglonConsumoMedicamento.Update
        
    Next i
    MsgBox "La solicitud se guardo correctamente!!!", vbOKOnly + vbInformation, "Solicitud Guardada"
    
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnImprimir.Enabled = False
    BtnProcesada.Enabled = False
    BtnAgregarRenglon.Enabled = False
    BtnBorrarRenglon.Enabled = False
    
    
Case Is = 2

   
    CSql = "Select * From ConsumoMedicamento Where IdConsumo='" & IdMax & "'"
    Set RsConsumoMedicamento = CrearRS(CSql)
    
    RsConsumoMedicamento.Fields("IdPaciente").Value = IdPaciente
    RsConsumoMedicamento.Fields("FechaConsumo").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
    RsConsumoMedicamento.Fields("Procesada").Value = Check1.Value
    RsConsumoMedicamento.Fields("IdUser").Value = IdUser
    RsConsumoMedicamento.Update
    
    CSql = "Select * From RenglonConsumoMedicamento Where IdConsumo='" & IdMax & "'"
    Set RsRenglonConsumoMedicamento = CrearRS(CSql)
    
    For i = 1 To DMGrid1.Rows
            
        b1 = DMGrid1.ValorCelda(i, 1)
        b2 = DMGrid1.ValorCelda(i, 2)
        b3 = DMGrid1.ValorCelda(i, 3)
    
        RsRenglonConsumoMedicamento.Fields("IdPaciente").Value = IdPaciente
        RsRenglonConsumoMedicamento.Fields("Codigo").Value = b1
        RsRenglonConsumoMedicamento.Fields("Descripcion").Value = b2
        RsRenglonConsumoMedicamento.Fields("Cantidad").Value = b3
        RsRenglonConsumoMedicamento.Fields("IdUser").Value = IdMax
        RsRenglonConsumoMedicamento.Update
        
    Next i
    MsgBox "La solicitud se Actualizó correctamente!!!", vbOKOnly + vbInformation, "Solicitud Actualizada"

    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnImprimir.Enabled = False
    BtnProcesada.Enabled = False
    BtnAgregarRenglon.Enabled = False
    BtnBorrarRenglon.Enabled = False

End Select
Blanqueo

End Sub

Private Sub BtnImprimir_Click()
On Error Resume Next
''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\ConsumoMedicamentos.rpt"
    .Connect = "DNS=CrReporte"
    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{ConsumoMedicamentos.IdConsumo} ='" & IdMax & "' And {ConsumoMedicamentos.IdPaciente}=" & IdPaciente & ""
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .WindowTitle = "Consumo de Medicamentos No.: " & LblNoConsumo.Caption
    .Action = 1
End With
End Sub

Private Sub BtnListaPaciente_Click()
On Error Resume Next
Tipo = "Consumo"
FrmListadoPaciente.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnProcesada_Click()
On Error Resume Next
Dim RsProcesar As New ADODB.Recordset
Msg = "Estas Seguro(a) de Procesar la Solicitud de Necesidades!!!"
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Procesar Solicitud")

If mensaje = vbYes Then
    CSql = "Update ConsumoMedicamento Set Procesada='1' where IdConsumo='" & Val(LblNoSolicitud.Caption) & "'"
    Set RsProcesar = CrearRS(CSql)
    
    Check1.Value = 1
    BtnProcesada.Enabled = False
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnAgregar.Enabled = False
    BtnAgregarRenglon.Enabled = False
    BtnBorrarRenglon.Enabled = False
    
    Msg = "Consumo de Medicamentos Procesado Correctamente!!!"
    mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Procesar Solicitud")
    
End If
End Sub



Private Sub DMGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 And DMGrid1.Col = 1 Then 'tecla F1
'
'    Tipo = "Consumo"
'    FrmListadoProductosServicios.Show
'
'End If
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton And DMGrid1.Col = 1 Then
    f = DMGrid1.Row
    Tipo = "Consumo"
    FrmListadoProductosServicios.Show vbModal, FrmPrincipal
End If
End Sub

Private Sub DMGrid1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And DMGrid1.Col = 1 Then 'tecla F1
    f = DMGrid1.Row
    Dim CodProduc
    Dim RsProducto As New ADODB.Recordset
    CodProduc = DMGrid1.ValorCelda(lRow, 1)
    
    CSql = "Select * From Productos where IdProducto='" & CodProduc & "'"
    Set RsProducto = CrearRS(CSql)
    
    If RsProducto.RecordCount > 0 Then
        DMGrid1.DColumnas(2).Locked = True
        DMGrid1.DColumnas(3).Locked = False
        DMGrid1.ValorCelda(lRow, 1) = RsProducto.Fields("IdProducto").Value
        DMGrid1.ValorCelda(lRow, 2) = RsProducto.Fields("Descripcion").Value
        DMGrid1.EditActive = False
    Else
        MsgBox "El Producto Buscado no existe!", vbOKOnly + vbCritical, "Error"
        DMGrid1.DColumnas(1).Locked = False
        DMGrid1.DColumnas(2).Locked = True
        DMGrid1.DColumnas(3).Locked = True
        DMGrid1.EditActive = True
    End If
    
    DMGrid1.PaintMGrid
    
End If


End Sub

Sub Blanqueo()

InitGrid
TxtCedula.Text = ""
TxtApellido.Text = ""
TxtNoHistoria.Text = ""
TxtNombre.Text = ""
TxtDiagnostico.Text = ""
DTPicker1.Value = DateTime.Date
TxtNoInsumos.Text = 0
TxtEdad.Text = ""

CSql = "Select Max(IdConsumo) + 1 as MaxId From ConsumoMedicamento"
Set RsIdMax = CrearRS(CSql)

If RsIdMax.RecordCount > 0 Then
    LblNoConsumo.Caption = Format(RsIdMax.Fields("MaxId").Value, "0000#")
Else
    LblNoConsumo.Caption = Format(1, "0000#")
End If

End Sub
Private Sub Form_Load()
Centrar Me
DTPicker1.Value = DateTime.Date

CSql = "Select Max(IdConsumo) + 1 as MaxId From ConsumoMedicamento"
Set RsIdMax = CrearRS(CSql)

If RsIdMax.RecordCount > 0 Then
    LblNoConsumo.Caption = Format(RsIdMax.Fields("MaxId").Value, "0000#")
Else
    LblNoConsumo.Caption = Format(1, "0000#")
End If

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnProcesada.Enabled = False
BtnAgregarRenglon.Enabled = False
BtnBorrarRenglon.Enabled = False


InitGrid
End Sub

Sub InitGrid()

'carga las columnas y encabezados de columna
DMGrid1.Cols = 3
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0

DMGrid1.DColumnas(1).Locked = False
DMGrid1.DColumnas(2).Locked = False
DMGrid1.DColumnas(3).Locked = False

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre del Producto / Descripción"
DMGrid1.DColumnas(3).Caption = "Cantidad"

End Sub





Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

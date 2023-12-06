VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmSolicitudNecesidades 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de insumos"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12360
   Icon            =   "FrmSolicitudNecesidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12360
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.TextBox TxtCantidadInsumos 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   11160
         TabIndex        =   34
         Text            =   "0"
         Top             =   6360
         Width           =   855
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   7646
         Object.Width           =   11865
         Object.Height          =   4305
         ScrollBar       =   1
         Editable        =   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   7920
         Top             =   360
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   6840
         Width           =   3975
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
            TabIndex        =   25
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el Código de la Solicitud de insumos"
            Top             =   240
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2640
            TabIndex        =   26
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
            MICON           =   "FrmSolicitudNecesidades.frx":1002
            PICN            =   "FrmSolicitudNecesidades.frx":101E
            PICH            =   "FrmSolicitudNecesidades.frx":1283
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
            Caption         =   "Nº Solicitud:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   330
            Width           =   870
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   11
         Top             =   6840
         Width           =   7815
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6720
            TabIndex        =   12
            ToolTipText     =   "Cerrar"
            Top             =   225
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
            MICON           =   "FrmSolicitudNecesidades.frx":1515
            PICN            =   "FrmSolicitudNecesidades.frx":1531
            PICH            =   "FrmSolicitudNecesidades.frx":16FA
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
            TabIndex        =   13
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   230
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
            MICON           =   "FrmSolicitudNecesidades.frx":192F
            PICN            =   "FrmSolicitudNecesidades.frx":194B
            PICH            =   "FrmSolicitudNecesidades.frx":1BDA
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
            TabIndex        =   14
            ToolTipText     =   "Agregar"
            Top             =   230
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
            MICON           =   "FrmSolicitudNecesidades.frx":201B
            PICN            =   "FrmSolicitudNecesidades.frx":2037
            PICH            =   "FrmSolicitudNecesidades.frx":21C4
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
            Left            =   5520
            TabIndex        =   15
            ToolTipText     =   "Deshacer Operacion"
            Top             =   225
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
            MICON           =   "FrmSolicitudNecesidades.frx":23F9
            PICN            =   "FrmSolicitudNecesidades.frx":2415
            PICH            =   "FrmSolicitudNecesidades.frx":26F7
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
            TabIndex        =   16
            ToolTipText     =   "Eliminar"
            Top             =   230
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
            MICON           =   "FrmSolicitudNecesidades.frx":2948
            PICN            =   "FrmSolicitudNecesidades.frx":2964
            PICH            =   "FrmSolicitudNecesidades.frx":2B08
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
            Left            =   3960
            TabIndex        =   17
            ToolTipText     =   "Reporte"
            Top             =   225
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
            MICON           =   "FrmSolicitudNecesidades.frx":2CA7
            PICN            =   "FrmSolicitudNecesidades.frx":2CC3
            PICH            =   "FrmSolicitudNecesidades.frx":2DE8
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
         Caption         =   "Datos de la Solicitud"
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8895
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   8280
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            PrintFileLinesPerPage=   60
            WindowShowCloseBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.TextBox TxtCargo 
            Height          =   375
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox TxtDepartamento 
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox TxtCedula 
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtApellido 
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   670
            Width           =   3255
         End
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   670
            Width           =   3255
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarEmpleado 
            Height          =   375
            Left            =   3000
            TabIndex        =   18
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
            MICON           =   "FrmSolicitudNecesidades.frx":3078
            PICN            =   "FrmSolicitudNecesidades.frx":3094
            PICH            =   "FrmSolicitudNecesidades.frx":32F9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnListaEmpleados 
            Height          =   375
            Left            =   4320
            TabIndex        =   36
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Listado Empleados"
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
            MICON           =   "FrmSolicitudNecesidades.frx":358B
            PICN            =   "FrmSolicitudNecesidades.frx":35A7
            PICH            =   "FrmSolicitudNecesidades.frx":3830
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cedula:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellidos(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   760
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Área:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1170
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   4560
            TabIndex        =   9
            Top             =   1170
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   4560
            TabIndex        =   7
            Top             =   760
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   1575
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Procesada"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   51118083
            CurrentDate     =   40189
         End
         Begin ChamaleonButton.ChameleonBtn BtnProcesada 
            Height          =   375
            Left            =   1440
            TabIndex        =   31
            ToolTipText     =   "Procesar solicitudes de necesidades"
            Top             =   1080
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
            MICON           =   "FrmSolicitudNecesidades.frx":3C4B
            PICN            =   "FrmSolicitudNecesidades.frx":3C67
            PICH            =   "FrmSolicitudNecesidades.frx":3EDC
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
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   330
            Width           =   495
         End
         Begin VB.Label LblNoSolicitud 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   1080
            TabIndex        =   3
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Solicitud:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   773
            Width           =   900
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregarRenglon 
         Height          =   375
         Left            =   120
         TabIndex        =   32
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
         MICON           =   "FrmSolicitudNecesidades.frx":4158
         PICN            =   "FrmSolicitudNecesidades.frx":4174
         PICH            =   "FrmSolicitudNecesidades.frx":4301
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
         TabIndex        =   33
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
         MICON           =   "FrmSolicitudNecesidades.frx":4536
         PICN            =   "FrmSolicitudNecesidades.frx":4552
         PICH            =   "FrmSolicitudNecesidades.frx":46F6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Insumos:"
         Height          =   195
         Left            =   9480
         TabIndex        =   35
         Top             =   6450
         Width           =   1530
      End
   End
End
Attribute VB_Name = "FrmSolicitudNecesidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarDepartamentos As New ADODB.Recordset
Dim RsCargarCargos As New ADODB.Recordset
Dim RsEmpleados As New ADODB.Recordset
Dim RsEmpleadosDepartamento As New ADODB.Recordset
Dim RsEmpleadosCargos As New ADODB.Recordset
Dim RsIdMax As New ADODB.Recordset
Dim NewReg, IdEmpleado, IdMax
Dim RsSolicitudInsumo As New ADODB.Recordset
Dim RsRenglonSolicitudInsumo As New ADODB.Recordset
Dim IdDeparta, IdCargo

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
'If RsIdMax.RecordCount > 0 Then
'    If Not IsNull(RsIdMax.Fields("MaxId").Value) Then
'        LblNoSolicitud.Caption = Format(RsIdMax.Fields("MaxId").Value, "0000#")
'    Else
'        LblNoSolicitud.Caption = Format(1, "0000#")
'    End If
'Else
'    LblNoSolicitud.Caption = Format(1, "0000#")
'End If

End Sub

Private Sub BtnAgregarRenglon_Click()
On Error Resume Next
If DMGrid1.Rows >= 0 Then
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.PaintMGrid
    TxtCantidadInsumos.Text = DMGrid1.Rows
End If
End Sub

Private Sub BtnBorrarRenglon_Click()
On Error Resume Next
If DMGrid1.Rows > 0 Then
    DMGrid1.Rows = DMGrid1.Rows - 1
    DMGrid1.PaintMGrid
    TxtCantidadInsumos.Text = DMGrid1.Rows
End If
End Sub

Private Sub BtnBuscar_Click()

If TxtBuscar.Text = "" Or TxtBuscar.Text = "Busqueda" Then Exit Sub
NewReg = 2

CSql = "Select * From SolicitudInsumo Where IdInsumo='" & Trim(TxtBuscar.Text) & "'"
Set RsSolicitudInsumo = CrearRS(CSql)

If RsSolicitudInsumo.RecordCount = 0 Then
    Msg = "No Exite la Solicitud de Necesidades Buscada!!!"
    mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Solicitud de Necesidades")
    
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnImprimir.Enabled = False
    BtnProcesada.Enabled = False

    Exit Sub
End If
If RsSolicitudInsumo.RecordCount > 0 Then

    
    IdMax = RsSolicitudInsumo.Fields("IdInsumo").Value
    DTPicker1.Value = Format(RsSolicitudInsumo.Fields("FechaInsumo").Value, "dd/mm/yyyy")
    LblNoSolicitud.Caption = Format(RsSolicitudInsumo.Fields("IdInsumo").Value, "0000")
    
    If RsSolicitudInsumo.Fields("Procesada").Value = True Then
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
        BtnAgregarRenglon.Enabled = False
        BtnBorrarRenglon.Enabled = False
    End If
    
    IdEmpleado = RsSolicitudInsumo.Fields("IdEmpleado").Value
    IdDeparta = RsSolicitudInsumo.Fields("IdArea").Value
    IdCargo = RsSolicitudInsumo.Fields("IdCargo").Value


    CSql = "Select * From Empleados Where IdEmpleado='" & IdEmpleado & "'"
    Set RsSolicitudInsumo = CrearRS(CSql)

    TxtCedula.Text = RsSolicitudInsumo.Fields("Cedula").Value
    TxtApellido.Text = RsSolicitudInsumo.Fields("Apellido").Value
    TxtNombre.Text = RsSolicitudInsumo.Fields("Nombre").Value
    
    CSql = "Select * From Departamentos Where IdDepartamento='" & IdDeparta & "'"
    Set RsSolicitudInsumo = CrearRS(CSql)
    
    TxtDepartamento.Text = RsSolicitudInsumo.Fields("Descripcion").Value

    CSql = "Select * From Cargos Where IdCargos='" & IdCargo & "'"
    Set RsSolicitudInsumo = CrearRS(CSql)
    
    TxtCargo.Text = RsSolicitudInsumo.Fields("Cargo").Value

End If




CSql = "Select * From RenglonSolicitudInsumos Where IdInsumo='" & Trim(TxtBuscar.Text) & "'"
Set RsRenglonSolicitudInsumo = CrearRS(CSql)

If RsRenglonSolicitudInsumo.RecordCount = 0 Then Exit Sub

If RsRenglonSolicitudInsumo.RecordCount > 0 Then
i = 1
   Do While Not RsRenglonSolicitudInsumo.EOF
        DMGrid1.Rows = i
        DMGrid1.ValorCelda(i, 1) = Trim(RsRenglonSolicitudInsumo.Fields("Codigo").Value)
        DMGrid1.ValorCelda(i, 2) = RsRenglonSolicitudInsumo.Fields("Descripcion").Value
        DMGrid1.ValorCelda(i, 3) = RsRenglonSolicitudInsumo.Fields("Cantidad").Value
        
        i = i + 1
        RsRenglonSolicitudInsumo.MoveNext
    Loop
       TxtCantidadInsumos.Text = DMGrid1.Rows
    DMGrid1.PaintMGrid
End If




End Sub

Public Sub BtnBuscarEmpleado_Click()
Busqueda
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
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next

MsgBox "No existe el codigo fuente!", vbInformation
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
'validacion de campos
'-------------
If TxtCedula.Text = "" Then
    MsgBox "Ingrese el número de Cedula del Empleado!", vbOKOnly + vbCritical, "Error"
    TxtCedula.SetFocus
    Exit Sub
End If

'-------------
'Guardar y/o Actualizar
Select Case NewReg
    Case Is = 1

        CSql = "Select Max(IdInsumo) + 1 as MaxId From SolicitudInsumo"
        Set RsIdMax = CrearRS(CSql)
        
        If RsIdMax.RecordCount <> 0 Then
            If Not IsNull(RsIdMax.Fields("MaxId")) Then
                IdMax = RsIdMax.Fields("MaxId").Value
            Else
                IdMax = "1"
            End If
        End If
        
        CSql = "Select * From SolicitudInsumo"
        Set RsSolicitudInsumo = CrearRS(CSql)
        
        RsSolicitudInsumo.AddNew
        RsSolicitudInsumo.Fields("IdInsumo").Value = IdMax
        RsSolicitudInsumo.Fields("FechaInsumo").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        RsSolicitudInsumo.Fields("IdEmpleado").Value = IdEmpleado
        RsSolicitudInsumo.Fields("Procesada").Value = Check1.Value
        RsSolicitudInsumo.Fields("IdArea").Value = IdDeparta
        RsSolicitudInsumo.Fields("IdUser").Value = IdUser
        RsSolicitudInsumo.Fields("IdCargo").Value = IdCargo
        RsSolicitudInsumo.Update



        CSql = "Select * From RenglonSolicitudInsumos"
        Set RsRenglonSolicitudInsumo = CrearRS(CSql)
        
        For i = 1 To DMGrid1.Rows
        
            b1 = DMGrid1.ValorCelda(i, 1)
            b2 = DMGrid1.ValorCelda(i, 2)
            b3 = DMGrid1.ValorCelda(i, 3)
            
            RsRenglonSolicitudInsumo.AddNew
            RsRenglonSolicitudInsumo.Fields("IdInsumo").Value = IdMax
            RsRenglonSolicitudInsumo.Fields("FechaInsumo").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
            RsRenglonSolicitudInsumo.Fields("IdEmpleado").Value = IdEmpleado
            RsRenglonSolicitudInsumo.Fields("Codigo").Value = b1
            RsRenglonSolicitudInsumo.Fields("Descripcion").Value = b2
            RsRenglonSolicitudInsumo.Fields("Cantidad").Value = b3
            RsRenglonSolicitudInsumo.Fields("IdUser").Value = IdUser
            RsRenglonSolicitudInsumo.Update

        Next i

MsgBox "El Consumo se guardo correctamente!!!", vbOKOnly + vbInformation, "Operación Guardada"

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = True
BtnEliminar.Enabled = True
BtnImprimir.Enabled = True
BtnProcesada.Enabled = True
BtnAgregarRenglon.Enabled = False
BtnBorrarRenglon.Enabled = False

    Case Is = 2

     
        CSql = "Select * From SolicitudInsumo Where IdInsumo='" & IdMax & "'"
        Set RsSolicitudInsumo = CrearRS(CSql)
        
        RsSolicitudInsumo.Fields("IdEmpleado").Value = IdEmpleado
        RsSolicitudInsumo.Fields("Procesada").Value = Check1.Value
        RsSolicitudInsumo.Fields("IdArea").Value = IdDeparta
        RsSolicitudInsumo.Fields("IdUser").Value = IdUser
        RsSolicitudInsumo.Fields("IdCargo").Value = IdCargo
        RsSolicitudInsumo.Update



        CSql = "Select * From RenglonSolicitudInsumos Where IdInsumo='" & IdMax & "'"
        Set RsRenglonSolicitudInsumo = CrearRS(CSql)
        
        For i = 1 To DMGrid1.Rows
        
            b1 = DMGrid1.ValorCelda(i, 1)
            b2 = DMGrid1.ValorCelda(i, 4)
            b3 = DMGrid1.ValorCelda(i, 3)
            
                      
            RsRenglonSolicitudInsumo.Fields("Codigo").Value = b1
            RsRenglonSolicitudInsumo.Fields("Descripcion").Value = b2
            RsRenglonSolicitudInsumo.Fields("Cantidad").Value = b3
            RsRenglonSolicitudInsumo.Fields("IdUser").Value = IdUser
            RsRenglonSolicitudInsumo.Update

        Next i
MsgBox "El Consumo se Actualizo correctamente!!!", vbOKOnly + vbInformation, "Operación actualizada"

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = True
BtnEliminar.Enabled = True
BtnImprimir.Enabled = True
BtnProcesada.Enabled = True
BtnAgregarRenglon.Enabled = True
BtnBorrarRenglon.Enabled = True

End Select


End Sub

Private Sub BtnImprimir_Click()
On Error Resume Next
''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\SolicitudNecesidades.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{SolicitudNecesidades.IdInsumo} = " & IdMax
    .WindowTitle = "Solicitud de Necesidades No. " & LblNoOrden.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
End Sub

Private Sub BtnListaEmpleados_Click()
Tipo = "Insumos"
FrmListadoEmpleados.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnProcesada_Click()
On Error Resume Next
Dim RsProcesar As New ADODB.Recordset
Msg = "Estas Seguro(a) de Procesar la Solicitud de Insumos!!!"
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Procesar Solicitud")

If mensaje = vbYes Then
    CSql = "Update SolicitudInsumo Set Procesada='1' where IdInsumo='" & Val(LblNoSolicitud.Caption) & "'"
    Set RsProcesar = CrearRS(CSql)
    
    Check1.Value = 1
    BtnProcesada.Enabled = False
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnAgregar.Enabled = False
    BtnAgregarRenglon.Enabled = False
    BtnBorrarRenglon.Enabled = False
    
    Msg = "Solicitud de Insumos Procesada Correctamente!!!"
    mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Procesar Solicitud")
    
End If

End Sub

Private Sub DMGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 112 And DMGrid1.Col = 1 Then 'tecla F1
'    f = DMGrid1.Row
'    Tipo = "Insumos"
'    FrmListadoProductosServicios.Show
'
'End If






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
        'DMGrid1.ValorCelda(lRow, 3) = RsProducto.Fields("").Value
        DMGrid1.EditActive = False
       ' DMGrid1.ValorCelda(lRow, 4) = RsProducto.Fields("Descripcion").Value
      ' BtnAgregarRenglon.SetFocus
    Else
        MsgBox "El Producto Buscado no existe!", vbOKOnly + vbCritical, "Error"
        DMGrid1.DColumnas(1).Locked = False
        DMGrid1.DColumnas(2).Locked = True
        DMGrid1.DColumnas(3).Locked = True
        DMGrid1.EditActive = True
    End If
    
    DMGrid1.PaintMGrid
    
End If

If KeyAscii = 13 And DMGrid1.Col = 3 Then
    celvacia = DMGrid1.ValorCelda(lRow, 3)
    If celvacia = "" Then
        
    Else
        BtnAgregarRenglon.SetFocus
    End If
End If


End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton And DMGrid1.Col = 1 Then
    f = DMGrid1.Row
    Tipo = "Insumos"
    FrmListadoProductosServicios.Show vbModal, FrmPrincipal
End If
End Sub

Private Sub Form_Load()
Centrar Me
DTPicker1.Value = DateTime.Date
CSql = "Select Max(IdInsumo) + 1 as MaxId From SolicitudInsumo"
Set RsIdMax = CrearRS(CSql)

If RsIdMax.RecordCount > 0 Then
    If Not IsNull(RsIdMax.Fields("MaxId").Value) Then
        LblNoSolicitud.Caption = Format(RsIdMax.Fields("MaxId").Value, "0000#")
    Else
        LblNoSolicitud.Caption = Format(1, "0000#")
    End If
Else
    LblNoSolicitud.Caption = Format(1, "0000#")
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

Sub Busqueda()

If TxtCedula.Text = "" Then Exit Sub
CSql = "Select * From Empleados where Cedula='" & Trim(TxtCedula.Text) & "'"
Set RsEmpleados = CrearRS(CSql)

If RsEmpleados.RecordCount = 0 Then Exit Sub

If RsEmpleados.RecordCount > 0 Then
    IdEmpleado = RsEmpleados.Fields("IdEmpleado").Value
    TxtApellido.Text = RsEmpleados.Fields("Apellido").Value
    TxtNombre.Text = RsEmpleados.Fields("Nombre").Value
End If

CSql = "Select * From Departamentos where IdDepartamento='" & RsEmpleados.Fields("Departamentos").Value & "'"
Set RsEmpleadosDepartamento = CrearRS(CSql)
If RsEmpleadosDepartamento.RecordCount > 0 Then
    TxtDepartamento.Text = RsEmpleadosDepartamento.Fields("Descripcion").Value
End If

CSql = "Select * From Cargos where IdCargos='" & RsEmpleados.Fields("Cargo").Value & "'"
Set RsEmpleadosCargos = CrearRS(CSql)
If RsEmpleadosCargos.RecordCount > 0 Then
    TxtCargo.Text = RsEmpleadosCargos.Fields("Cargo").Value
End If

End Sub

Sub InitGrid()

'carga las columnas y encabezados de columna
DMGrid1.Cols = 3
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
'DMGrid1.DColumnas(4).Alignment = 0

DMGrid1.DColumnas(1).Locked = False
DMGrid1.DColumnas(2).Locked = False
DMGrid1.DColumnas(3).Locked = False
'DMGrid1.DColumnas(4).Locked = False

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 15 / 100)
'DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 45 / 100) - 300

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre del Producto / Descripción"
DMGrid1.DColumnas(3).Caption = "Cantidad"
'DMGrid1.DColumnas(4).Caption = "Descripcion"

End Sub

Sub Blaqueo()
TxtCedula.Text = ""
TxtApellido.Text = ""
TxtNombre.Text = ""
CboCargos.Text = ""
CboCargos.ListIndex = -1
CboDepartamentos.Text = ""
CboDepartamentos.ListIndex = -1
TxtCantidadInsumos.Text = DMGrid1.Rows
DTPicker1.Value = DateTime.Date
Check1.Value = 0
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub
Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub
Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

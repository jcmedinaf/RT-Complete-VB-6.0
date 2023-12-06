VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPrestamos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prestamos"
   ClientHeight    =   8370
   ClientLeft      =   6465
   ClientTop       =   2460
   ClientWidth     =   9945
   Icon            =   "Pretamos.frx":0000
   LinkTopic       =   "Form44"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9945
   Begin VB.Frame Frame8 
      BackColor       =   &H00EAEFEF&
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   6720
         Width           =   9495
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
            Left            =   120
            TabIndex        =   46
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el número del Prestamo a buscar"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   3720
            Top             =   120
         End
         Begin ChamaleonButton.ChameleonBtn n 
            Height          =   375
            Left            =   2280
            TabIndex        =   47
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
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
            MICON           =   "Pretamos.frx":1002
            PICN            =   "Pretamos.frx":101E
            PICH            =   "Pretamos.frx":1283
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnPagosPrestamos 
            Height          =   375
            Left            =   7680
            TabIndex        =   48
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Abonos"
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
            MICON           =   "Pretamos.frx":1515
            PICN            =   "Pretamos.frx":1531
            PICH            =   "Pretamos.frx":17CD
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
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   7440
         Width           =   9495
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   8400
            TabIndex        =   34
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
            MICON           =   "Pretamos.frx":1C0D
            PICN            =   "Pretamos.frx":1C29
            PICH            =   "Pretamos.frx":1DF2
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
            TabIndex        =   35
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
            MICON           =   "Pretamos.frx":2027
            PICN            =   "Pretamos.frx":2043
            PICH            =   "Pretamos.frx":22D2
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
            TabIndex        =   36
            ToolTipText     =   "Agregar "
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
            MICON           =   "Pretamos.frx":2713
            PICN            =   "Pretamos.frx":272F
            PICH            =   "Pretamos.frx":28BC
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
            Left            =   7200
            TabIndex        =   37
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
            MICON           =   "Pretamos.frx":2AF1
            PICN            =   "Pretamos.frx":2B0D
            PICH            =   "Pretamos.frx":2DEF
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
            TabIndex        =   38
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
            MICON           =   "Pretamos.frx":3040
            PICN            =   "Pretamos.frx":305C
            PICH            =   "Pretamos.frx":3200
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
            Left            =   6240
            TabIndex        =   39
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
            MICON           =   "Pretamos.frx":339F
            PICN            =   "Pretamos.frx":33BB
            PICH            =   "Pretamos.frx":3651
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
            Left            =   5640
            TabIndex        =   40
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
            MICON           =   "Pretamos.frx":38B0
            PICN            =   "Pretamos.frx":38CC
            PICH            =   "Pretamos.frx":3B61
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
            Left            =   4080
            TabIndex        =   41
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
            MICON           =   "Pretamos.frx":3DBD
            PICN            =   "Pretamos.frx":3DD9
            PICH            =   "Pretamos.frx":3EFE
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
         Caption         =   "Datos iniciales del Prestamo"
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   9495
         Begin VB.TextBox TxtMxC 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            TabIndex        =   49
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Pretamos.frx":418E
            Left            =   120
            List            =   "Pretamos.frx":4198
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox TxtPrestamo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            TabIndex        =   20
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox TxtInteres 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8160
            TabIndex        =   19
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Pretamos.frx":41B2
            Left            =   120
            List            =   "Pretamos.frx":41C8
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   600
            Width           =   4095
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "Pretamos.frx":4229
            Left            =   2400
            List            =   "Pretamos.frx":424D
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox TxtMontoInteres 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8160
            TabIndex        =   16
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6360
            TabIndex        =   22
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   52887553
            CurrentDate     =   39981
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interes %:"
            Height          =   195
            Left            =   8040
            TabIndex        =   28
            Top             =   720
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto x Cuota:"
            Height          =   195
            Left            =   4320
            TabIndex        =   50
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto:"
            Height          =   195
            Left            =   4320
            TabIndex        =   29
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modo de Pago:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cuotas:"
            Height          =   195
            Left            =   2400
            TabIndex        =   26
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Necesidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha  de Asignación:"
            Height          =   195
            Left            =   6360
            TabIndex        =   24
            Top             =   360
            Width           =   1590
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Intereses:"
            Height          =   195
            Left            =   8040
            TabIndex        =   23
            Top             =   1200
            Visible         =   0   'False
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Informacion de los prestamos"
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   4440
         Width           =   9495
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   1815
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3201
            Object.Width           =   9225
            Object.Height          =   1785
            DrawColorGrid   =   1
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos de Empleado"
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7455
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   5880
            TabIndex        =   30
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   52887553
            CurrentDate     =   40119
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   7
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   5
            Top             =   1200
            Width           =   3495
         End
         Begin ChamaleonButton.ChameleonBtn BrnListaEmpleados 
            Height          =   375
            Left            =   1920
            TabIndex        =   42
            Top             =   480
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
            MICON           =   "Pretamos.frx":4285
            PICN            =   "Pretamos.frx":42A1
            PICH            =   "Pretamos.frx":452A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label NReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   1920
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salario:"
            Height          =   195
            Left            =   3720
            TabIndex        =   12
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Ingreso:"
            Height          =   195
            Left            =   5880
            TabIndex        =   11
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cédula:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   3720
            TabIndex        =   9
            Top             =   960
            Width           =   765
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   2535
         Left            =   7680
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         Begin VB.Timer Timer1 
            Interval        =   10
            Left            =   120
            Top             =   960
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Prestamo:"
            Height          =   195
            Left            =   495
            TabIndex        =   2
            Top             =   240
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "FrmPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsEmpleados As New ADODB.Recordset
Dim RsPrestamos As New ADODB.Recordset
Dim RsTemp As Recordset
Dim Nuevo As Integer
Dim IdEmpla As Integer
Dim IdPrestamo As Integer
Dim MuestraDialogo As Boolean


Sub IniDMGrid()
DMGrid1.Clear
DMGrid1.Rows = 0

DMGrid1.Cols = 7
DMGrid1.Rows = 0

DMGrid1.DColumnas(1).Caption = "Fecha"
DMGrid1.DColumnas(2).Caption = "Prestamo"
DMGrid1.DColumnas(3).Caption = "Abonado"
DMGrid1.DColumnas(4).Caption = "Cuota"
DMGrid1.DColumnas(5).Caption = "Monto"

DMGrid1.DColumnas(6).Visible = False    ' Se Almacena el ID del prestamo

DMGrid1.DColumnas(7).Caption = "Saldo"


DMGrid1.DColumnas(2).Alignment = 1
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1
DMGrid1.DColumnas(7).Alignment = 1

DMGrid1.DColumnas(2).IsNumber = True
DMGrid1.DColumnas(3).IsNumber = True
DMGrid1.DColumnas(7).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(7).Width = Val(DMGrid1.Width * 20 / 100) - 300

End Sub
Private Sub BrnListaEmpleados_Click()
BtnDesHacer_Click
Tipo = "Prestamos"
FrmListadoEmpleados.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAgregar_Click()

If IdEmpla = 0 Then MsgBox "Debe Seleccionar un Trabajador para realizar el prestamo!", vbExclamation + vbOKOnly, "Error": Exit Sub

BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnSiguiente.Enabled = False
BtnAnterior.Enabled = False
BtnPagosPrestamos.Enabled = False

TxtMontoInteres.Enabled = True
TxtPrestamo.Enabled = True
TxtInteres.Enabled = True
DTPicker1.Enabled = True
Frame1.Enabled = True

Call Blanqueo
Timer1.Enabled = False
Nuevo = 1

End Sub

Public Sub BtnAnterior_Click()

If RsEmpleados.RecordCount <> 0 Then
    If Not RsEmpleados.BOF Then
        RsEmpleados.MovePrevious
        If RsEmpleados.BOF Then RsEmpleados.MoveLast
    Else
        RsEmpleados.MoveLast
    End If
    Call CargaTrabajador
    Call CargaPrest1
Else
    Blanqueo
    NReg.Caption = "Registro 0 / 0"
    MsgBox "No hay registros cargados!", vbExclamation + vbOKOnly, "No hay datos."
End If

End Sub

Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) = "" Then Exit Sub
  
BtnDesHacer_Click
'CSql = "Select * From PrestamosEmpleados Where Cedula='" & TxtBuscar.Text & "' Or IdPrestamo=" & TxtBuscar.Text & ""
CSql = "Select * From Empleados Where Cedula=" & Trim(TxtBuscar.Text) & " AND Activo='1'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    ' Mostrar datos del trabajador
    RsEmpleados.MoveFirst
    RsEmpleados.Find " Cedula=" & Trim(TxtBuscar.Text) & ""
    CargaTrabajador
    Exit Sub
End If

CSql = "Select * From Prestamos Where IdPrestamos=" & Val(TxtBuscar.Text) & " AND Activo='1'"
Set RsPrestamos = CrearRS(CSql)

If RsPrestamos.RecordCount <> 0 Then
    'CSql = "Select * From Prestamos "
    'Set RsPrestamos = CrearRS(CSql)
    RsEmpleados.MoveFirst
    RsEmpleados.Find " IdEmpleado=" & RsPrestamos.Fields("IdEmpleado").Value & ""
    RsPrestamos.MoveFirst
    CargaTrabajador
    Call CargaPrest1
    Exit Sub
End If

MsgBox "No se encontraron registros relacionados con la busqueda!", vbExclamation + vbOKOnly, "No hay registros"

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
BtnImprimir.Enabled = True
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
BtnPagosPrestamos.Enabled = True

TxtMontoInteres.Enabled = False
TxtPrestamo.Enabled = False
TxtInteres.Enabled = False
DTPicker1.Enabled = False
Frame1.Enabled = False
End Sub

Private Sub BtnEliminar_Click()
MsgBox "En Construccion!", vbInformation + vbOKOnly, ""
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim TempPeriodos As Integer
Dim TempDia As Integer
Dim NuevoId As Integer
Dim NPrestam As Integer
Dim resp As Integer

If IdEmpla = 0 Then MsgBox "Debe Seleccionar un Trabajador para realizar el prestamo!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Desea asignarle el prestamo al trabajador?", vbQuestion + vbYesNo, "Confirmar.")
If resp = vbNo Then Exit Sub

If Combo1.ListIndex = -1 Then
    MsgBox "Seleccione el modo de pago!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo1.SetFocus
    Exit Sub
ElseIf Combo2.ListIndex = -1 Then
    MsgBox "Seleccione el motivo o necesidad!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo2.SetFocus
    Exit Sub
ElseIf Combo3.ListIndex = -1 Then
    MsgBox "Seleccione el número de cuotas!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo3.SetFocus
    Exit Sub
ElseIf Trim(TxtMxC.Text) = "" Then
    MsgBox "Ingrese el Monto para cada cuota!", vbExclamation + vbOKOnly, "Faltan datos!"
    TxtMxC.SetFocus
    Exit Sub
ElseIf Not IsNumeric(TxtMxC.Text) Then
    MsgBox "Ingrese sólo números para las cuotas!", vbExclamation + vbOKOnly, "Error!"
    TxtMxC.SetFocus
    TxtMxC.Text = "0,00"
    Exit Sub
ElseIf Trim(TxtPrestamo.Text) = "" Then
    MsgBox "Ingrese el monto para el prestamo!", vbExclamation + vbOKOnly, "Faltan datos!"
    TxtPrestamo.SetFocus
    Exit Sub
ElseIf Not IsNumeric(TxtPrestamo.Text) Then
    MsgBox "Ingrese sólo números para el prestamo!", vbExclamation + vbOKOnly, "Error!"
    TxtPrestamo.SetFocus
    TxtPrestamo.Text = "0,00"
    Exit Sub
ElseIf CDbl(TxtPrestamo.Text) = 0# Then
    MsgBox "El prestamo debe ser mayor de CERO (0)!", vbExclamation + vbOKOnly, "Error!"
    TxtPrestamo.SetFocus
    TxtPrestamo.Text = "0,00"
    Exit Sub
ElseIf CDbl(TxtMxC.Text) = 0# Then
    MsgBox "Las cuotas deben ser mayor de CERO (0)!", vbExclamation + vbOKOnly, "Error!"
    TxtMxC.SetFocus
    Exit Sub
End If

' Consulta para obtener un nuevo Id
CSql = "SELECT MAX(IdPrestamos)+1 AS NuevoId FROM Prestamos"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If

' Consulta para obtener un nuevo Id de prestamo para el empleado
CSql = "SELECT MAX(NPrestamo)+1 AS NuevoId FROM Prestamos"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NPrestam = Val(RsTemp.Fields(0).Value)
Else
    NPrestam = 1
End If

If Trim(TxtInteres.Text) = "" Then TxtInteres.Text = "0"
Select Case Nuevo
Case Is = 1
        
        CSql = "Insert into Prestamos(IdPrestamos,IdEmpleado,Monto_Presta,Interes_Presta,ModoPago,Necesidad," & _
            "Cuotas,Fecha_Prestamos,Activo,IdUser,AbonoMax,Abonos,Adeuda,NPrestamo) VALUES(" & NuevoId & "," & IdEmpla & "," & Replace(Replace(TxtPrestamo.Text, ".", ""), ",", ".") & _
            "," & CDbl(Replace(Replace(TxtInteres.Text, ".", ""), ",", ".")) & "," & Combo1.ItemData(Combo1.ListIndex) & "," & _
            Combo2.ItemData(Combo2.ListIndex) & "," & Combo3.ItemData(Combo3.ListIndex) & ",'" & Format(DTPicker1.Value, "dd/MM/yyyy") & _
            "','1'," & IdUser & ", " & Replace(Replace((CDbl(TxtMxC.Text)), ".", ""), ",", ".") & _
            " ,0," & Replace(Replace(TxtPrestamo.Text, ".", ""), ",", ".") & "," & NPrestam & ")"
        Set RsTemp = CrearRS(CSql)
        
        MsgBox "Registro Agregado satisfactoriamente", vbInformation + vbOKOnly, "Registro Guardado!"
        
        
        ' Calcula los periodos del año del objeto DTPicker1
        Dim Anio As Integer
        Anio = Val(Format(DTPicker1.Value, "yyyy"))
        Call Calcular_Periodos(Str(Anio))
        
        TempDia = Val(Format(DTPicker1.Value, "dd"))
        TempPeriodos = Val(Format(DTPicker1.Value, "MM"))
        
        ' Calcula en que periodo se esta creando el prestamo y lo almacena en TempPeriodos
        If TempDia > 15 Then
            TempPeriodos = TempPeriodos * 2
        Else
            If TempPeriodos > 1 Then
                TempPeriodos = (TempPeriodos - 1) * 2
                TempPeriodos = TempPeriodos + 1
            Else
                TempPeriodos = 1
            End If
        End If
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        Dim ValorACobrar As Double
        Dim ValorAcumulado As Double
        Dim NCuotas As Integer
        
        NCuotas = Val(Combo3.ItemData(Combo3.ListIndex))
        
        Dim Band As Boolean
        Dim i As Integer
        Dim j As Integer
        
        ValorAcumulado = 0
        i = TempPeriodos - 1
        j = 0
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        CSql = "SELECT MAX(IdRengPrestamo)+1 AS NuevoId FROM RenglonPrestamos"
        Set RsTemp = CrearRS(CSql)
        
        If Not IsNull(RsTemp.Fields(0).Value) Then
            NPrestam = Val(RsTemp.Fields(0).Value)
        Else
            NPrestam = 1
        End If
        
        Band = True
        While Band
            
            i = i + 1   ' Periodos
            j = j + 1   ' Cuotas
            
            ValorACobrar = CDbl(TxtMxC.Text)
            ValorAcumulado = ValorAcumulado + ValorACobrar
            
            If CDbl(TxtPrestamo.Text) < ValorAcumulado Then ValorACobrar = (CDbl(TxtPrestamo.Text) - (ValorAcumulado - ValorACobrar))
            
            ' Crea las cuotas del prestamo...
            CSql = "INSERT INTO RenglonPrestamos (IdRengPrestamo, IdPrestamo, IdEmpleado, Cuota, AbonoMax, " & _
            " MontoAbono, FechaPago, IdUser) VALUES(" & NPrestam & "," & NuevoId & "," & IdEmpla & "," & j & _
            "," & Replace(Replace(ValorACobrar, ".", ""), ",", ".") & ",0,'" & PyFs(i - 1, 2) & "'," & IdUser & " )"
            Set RsTemp = CrearRS(CSql)
            
            NPrestam = NPrestam + 1
            If i >= 24 Then
                i = 0
                Anio = Anio + 1
                Call Calcular_Periodos(Str(Anio))
            End If
            
            If j = NCuotas Then Band = False
        Wend
        
Case Is = 0

        CSql = "SELECT * FROM RenglonPrestamos WHERE IdEmpleado=" & IdEmpla & " AND IdPrestamos=" & Val(Label21.Caption)
        Set RsTemp = CrearRS(CSql)
        
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
          Dim AbonoTot As Double
          AbonoTot = 0
          'If RsTemp.RecordCount <> 0 Then
          While Not RsTemp.EOF
              AbonoTot = AbonoTot + CDbl(RsTemp.Fields("MontoAbono").Value)
              RsTemp.MoveNext
          Wend
            
        ' Condicional para DENEGAR la modificacion del monto del prestamo en el caso de que hallan realizado abonos
          If AbonoTot <> 0 Then
              MsgBox "No se puede modificar el monto del prestamo, ya que se han realizado ABONOS al mismo!", vbCritical + vbOKOnly, "Operación Fallida!"
              Exit Sub
          End If
          'End If
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        Dim CampAbono As Double
        CSql = "SELECT Abonos FROM Prestamos WHERE IdPrestamos = " & Val(Label21.Caption)
        Set RsTemp = CrearRS(CSql)
        CampAbono = CDbl(RsTemp.Fields("Abonos").Value)
        
        CSql = "Update Prestamos set Monto_Presta = " & Replace(Replace(TxtPrestamo.Text, ".", ""), ",", ".") & _
            ", Interes_Presta = " & Replace(Replace(TxtInteres.Text, ".", ""), ",", ".") & _
            ", ModoPago = " & Combo1.ItemData(Combo1.ListIndex) & ", Necesidad = " & Combo2.ItemData(Combo2.ListIndex) & ", Cuotas = " & _
            Combo3.ItemData(Combo3.ListIndex) & ",Fecha_prestamos ='" & Format(DTPicker1.Value, "dd/MM/yyyy") & "', IdEmpleado = " & IdEmpla & _
            ", Activo='1', IdUser=" & IdUser & ", AbonoMax=" & Replace(Replace((CDbl(TxtPrestamo.Text) / Val(Combo3.ItemData(Combo3.ListIndex))), ".", ""), ",", ".") & _
            ", Adeuda=" & Replace(Replace((CDbl(TxtPrestamo.Text) - CampAbono), ".", ""), ",", ".") & " Where  IdPrestamos = " & Val(Label21.Caption)
        Set RsTemp = CrearRS(CSql)
        
        MsgBox "Registro Actualizado Satisfactoriamente", vbInformation + vbOKOnly, "Registro Actualizado!"
End Select

BtnDesHacer_Click

Call CargaPrest1


End Sub

Private Sub BtnImprimir_Click()
If IdEmpla = 0 Then MsgBox "Debe Seleccionar un Trabajador para poder imprimir!", vbExclamation + vbOKOnly, "Error": Exit Sub
End Sub

Private Sub BtnPagosPrestamos_Click()
Dim SelectIdPrestamo As Integer

' Condicional que verifica el se eligió un trabajador
If IdEmpla = 0 Then MsgBox "Debe Seleccionar un Trabajador para ver los abonos!", vbExclamation + vbOKOnly, "Error": Exit Sub

'Condicional estructurado para verificar si mostrará o no los abonos del prestamo.
If DMGrid1.Rows <> 0 Then
    ' el primero verifica si no ha seleccionado ningun prestamo
    If DMGrid1.Row = 0 Then
        SelectIdPrestamo = 1
    Else
        SelectIdPrestamo = Val(DMGrid1.ValorCelda(DMGrid1.Row, 6))
    End If
Else
    MsgBox "El empleado no tiene prestamos registrados!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

FrmPagosPrestamos.IdEmpla = IdEmpla             ' Envia el ID de trabajador al form de abonos
FrmPagosPrestamos.IdPrestamo = SelectIdPrestamo ' Envia el ID del prestamo seleccionado al form de abonos

FrmPagosPrestamos.TxtCedula.Text = Text1.Text
FrmPagosPrestamos.TxtNombres.Text = Text2.Text
FrmPagosPrestamos.TxtApellidos.Text = Text3.Text
FrmPagosPrestamos.Show vbModal, FrmPrincipal    ' Muestra el formulario de los abonos realizados

End Sub

Public Sub BtnSiguiente_Click()
If RsEmpleados.RecordCount <> 0 Then
    If Not RsEmpleados.EOF Then
        RsEmpleados.MoveNext
        If RsEmpleados.EOF Then RsEmpleados.MoveFirst
    Else
        RsEmpleados.MoveFirst
    End If
    Call CargaTrabajador
    Call CargaPrest1
Else
    Blanqueo
    NReg.Caption = "Registro 0 / 0"
    MsgBox "No hay registros cargados!", vbExclamation + vbOKOnly, "No hay datos."
End If
End Sub

Private Sub Combo3_Click()
Dim Valo As String

Valo = 0
If Frame1.Enabled = False Then Exit Sub
If MuestraDialogo = False Then Exit Sub

If Combo3.ListIndex = 9 Then
    Valo = InputBox("Ingrese la cantidad de Cuotas:", "Número de Cuotas.", "11")
    
    If IsNumeric(Valo) Then
        If Val(Valo) > 0 Then
            Combo3.ItemData(Combo3.ListIndex) = Valo
            Combo3.List(Combo3.ListIndex) = "Definida (" & Valo & ")"
        End If
    ElseIf Trim(Valo) <> "" Then
        MsgBox "Debe ingresar sólo números!", vbExclamation + vbOKOnly, "Error"
    End If
End If

If Trim(TxtPrestamo.Text) <> "" Then
    If Valo = 0 Then
        TxtMxC.Text = Format(CDbl(TxtPrestamo.Text) / Combo3.ItemData(Combo3.ListIndex), "#,##0.00")
    Else
        TxtMxC.Text = Format(CDbl(TxtPrestamo.Text) / Valo, "#,##0.00")
    End If
End If
End Sub

Public Sub CargaTrabajador()

    'CSql = "Select * From Empleados Where IdEmpleado = " & IdEmpla & " AND Activo=1"
    'Set RsEmpleados = CrearRS(CSql)
    
    If RsEmpleados.RecordCount <> 0 Then
        IdEmpla = Val(RsEmpleados.Fields("IdEmpleado").Value)
        If Trim(RsEmpleados.Fields("Cedula").Value) <> "" Then Text1.Text = RsEmpleados.Fields("Cedula").Value
        If Trim(RsEmpleados.Fields("Nombre").Value) <> "" Then Text2.Text = RsEmpleados.Fields("Nombre").Value
        If Trim(RsEmpleados.Fields("Apellido").Value) <> "" Then Text3.Text = RsEmpleados.Fields("Apellido").Value
        If Trim(RsEmpleados.Fields("Fecha_ing").Value) <> "" Then Text6.Text = RsEmpleados.Fields("Fecha_ing").Value
        If Trim(RsEmpleados.Fields("Sueldo").Value) <> "" Then Text5.Text = RsEmpleados.Fields("Sueldo").Value
        
        DMGrid1.Clear
        DMGrid1.Rows = 0
        
        CSql = "SELECT * FROM Prestamos Where Activo='1' AND IdEmpleado = " & IdEmpla
        Set RsPrestamos = CrearRS(CSql)
        
        If RsPrestamos.RecordCount > 0 Then

            RsPrestamos.MoveFirst
            If RsPrestamos.Fields("Monto_Presta").Value <> "" Then TxtPrestamo.Text = Format(RsPrestamos.Fields("Monto_Presta").Value, "#,##0.00") Else TxtPrestamo.Text = Format(0, "#,##0.00")
            If RsPrestamos.Fields("Interes_Presta").Value <> "" Then TxtInteres.Text = Format(RsPrestamos.Fields("Interes_Presta").Value, "#,##0.00") Else TxtInteres.Text = Format(0, "#,##0.00")
            TxtMontoInteres.Text = ""
            If RsPrestamos.Fields("IdPrestamos").Value <> "" Then Label21.Caption = RsPrestamos.Fields("IdPrestamos").Value Else Label21.Caption = ""
            If RsPrestamos.Fields("IdEmpleado").Value <> "" Then IdEmpla = RsPrestamos.Fields("IdEmpleado").Value

            While Not RsPrestamos.EOF
                DMGrid1.Rows = DMGrid1.Rows + 1
                DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsPrestamos.Fields("Fecha_Prestamos").Value
                DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Format(RsPrestamos.Fields("Monto_Presta").Value, "#,##0.00")
                DMGrid1.ValorCelda(DMGrid1.Rows, 3) = Format(RsPrestamos.Fields("Abonos").Value, "#,##0.00")
                DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPrestamos.Fields("Cuotas").Value
                DMGrid1.ValorCelda(DMGrid1.Rows, 5) = Format(RsPrestamos.Fields("AbonoMax").Value, "#,##0.00")
                DMGrid1.ValorCelda(DMGrid1.Rows, 6) = RsPrestamos.Fields("IdPrestamos").Value
                DMGrid1.ValorCelda(DMGrid1.Rows, 7) = Format((CDbl(RsPrestamos.Fields("Monto_Presta").Value) - CDbl(RsPrestamos.Fields("Abonos").Value)), "#,##0.00")
                RsPrestamos.MoveNext
            Wend
            DMGrid1.RowBackColor 1, RGB(255, 255, 255)
        End If
        
        DMGrid1.PaintMGrid

        NReg.Caption = "Registro " & RsEmpleados.AbsolutePosition & " / " & RsEmpleados.RecordCount
    Else
        NReg.Caption = "Registro 0 / 0"
        IdEmpla = 0
        Text1.Text = "":        Text2.Text = "":        Text3.Text = "":       Text6.Text = "":        Text5.Text = ""
        MsgBox "No hay registros cargados!", vbExclamation + vbOKOnly, "No hay datos."
    End If
End Sub


Private Sub Combo3_GotFocus()
MuestraDialogo = True
End Sub

Private Sub Combo3_LostFocus()
MuestraDialogo = False
End Sub

Private Sub DMGrid1_DobleClick()
BtnPagosPrestamos_Click
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)

If IdEmpla = 0 Then Exit Sub
If DMGrid1.Rows = 0 Then Exit Sub
If DMGrid1.Row = 0 Then Exit Sub

Blanqueo

CSql = "SELECT * FROM Prestamos Where Activo='1' AND IdEmpleado = " & IdEmpla & " AND IdPrestamos=" & DMGrid1.ValorCelda(DMGrid1.Row, 6)
Set RsPrestamos = CrearRS(CSql)

If RsPrestamos.RecordCount > 0 Then
   
    If RsPrestamos.EOF Then Exit Sub
    If RsPrestamos.Fields("Monto_presta").Value <> "" Then TxtPrestamo.Text = Format(RsPrestamos.Fields("Monto_presta").Value, "#,##0.00") Else TxtPrestamo.Text = Format(0, "#,##0.00")
    If RsPrestamos.Fields("Interes_Presta").Value <> "" Then TxtInteres.Text = Format(RsPrestamos.Fields("Interes_Presta").Value, "#,##0.00") Else TxtInteres.Text = Format(0, "#,##0.00")
    If RsPrestamos.Fields("IdPrestamos").Value <> "" Then Label21.Caption = RsPrestamos.Fields("IdPrestamos").Value Else Label21.Caption = ""
    If RsPrestamos.Fields("IdEmpleado").Value <> "" Then IdEmpla = RsPrestamos.Fields("IdEmpleado").Value
    
    If Not IsNull(RsPrestamos.Fields("Fecha_Prestamos").Value) Then DTPicker1.Value = Format(RsPrestamos.Fields("Fecha_prestamos").Value, "dd/mm/yyyy") Else DTPicker1.Value = DateTime.Date
    
    For T = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(T) = RsPrestamos.Fields("ModoPago").Value Then
        Combo1.ListIndex = T
        Exit For
        End If
    Next T
     
    For T = 0 To Combo2.ListCount - 1
        If Combo2.ItemData(T) = RsPrestamos.Fields("Necesidad").Value Then
        Combo2.ListIndex = T
        Exit For
        End If
    Next T
     
    For T = 0 To Combo3.ListCount - 1
        If Combo3.ItemData(T) = RsPrestamos.Fields("Cuotas").Value Then
        
        Combo3.ListIndex = T
        Exit For
        End If
    Next T
    'Nuevo = 0
Else
    Nuevo = 1
End If

End Sub

Private Sub DTPicker1_Change()
Cambio = 1
End Sub

 
Private Sub Form_Load()
Centrar Me

IniDMGrid

TxtMontoInteres.Enabled = False
TxtPrestamo.Enabled = False
TxtInteres.Enabled = False
DTPicker1.Enabled = False
Frame1.Enabled = False

DTPicker1.Value = Format(Now, "dd/mm/yyyy")
DTPicker3.Value = Format(Now, "dd/mm/yyyy")

Call CoNect

If Not RsPrestamos.EOF Then
    RsPrestamos.MoveLast
    CargaPrest1
End If
'Nuevo = 0
End Sub

Sub CoNect()

CSql = "Select * From Empleados Where Activo='1'"
Set RsEmpleados = CrearRS(CSql)
    
CSql = "SELECT * FROM Prestamos Where Activo='1'" 'where IdEmpleado = " & IdEmpla & ""
Set RsPrestamos = CrearRS(CSql)

End Sub
Public Sub CargaPrest1()

Blanqueo
If IdEmpla = 0 Then Exit Sub

CSql = "SELECT * FROM Prestamos Where Activo='1'" 'where IdEmpleado = " & IdEmpla & ""
Set RsPrestamos = CrearRS(CSql)

If RsPrestamos.RecordCount > 0 Then

    RsPrestamos.MoveFirst
    RsPrestamos.Find (" IdEmpleado=" & IdEmpla)
    
    If RsPrestamos.EOF Then Exit Sub
    If RsPrestamos.Fields("Monto_presta").Value <> "" Then TxtPrestamo.Text = Format(RsPrestamos.Fields("Monto_presta").Value, "#,##0.00") Else TxtPrestamo.Text = Format(0, "#,##0.00")
    If RsPrestamos.Fields("Interes_Presta").Value <> "" Then TxtInteres.Text = Format(RsPrestamos.Fields("Interes_Presta").Value, "#,##0.00") Else TxtInteres.Text = Format(0, "#,##0.00")

    TxtMontoInteres.Text = ""
    
    If RsPrestamos.Fields("IdPrestamos").Value <> "" Then Label21.Caption = RsPrestamos.Fields("IdPrestamos").Value Else Label21.Caption = ""
    If RsPrestamos.Fields("IdEmpleado").Value <> "" Then IdEmpla = RsPrestamos.Fields("IdEmpleado").Value
    
    If Not IsNull(RsPrestamos.Fields("Fecha_Prestamos").Value) Then DTPicker1.Value = Format(RsPrestamos.Fields("Fecha_prestamos").Value, "dd/mm/yyyy") Else DTPicker1.Value = DateTime.Date
    
    For T = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(T) = RsPrestamos.Fields("ModoPago").Value Then
        Combo1.ListIndex = T
        Exit For
        End If
    Next T
     
    For T = 0 To Combo2.ListCount - 1
        If Combo2.ItemData(T) = RsPrestamos.Fields("Necesidad").Value Then
        Combo2.ListIndex = T
        Exit For
        End If
    Next T
     
    For T = 0 To Combo3.ListCount - 1
        If Combo3.ItemData(T) = RsPrestamos.Fields("Cuotas").Value Then
        
        Combo3.ListIndex = T
        Exit For
        End If
    Next T
    'Nuevo = 0
Else
    Nuevo = 1
End If
CargaTrabajador
End Sub

Sub Blanqueo()

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
TxtPrestamo.Text = ""
'Text5.Text = ""
'Text6.Text = ""
TxtInteres.Text = ""
TxtBuscar.Text = ""
DTPicker1.Value = Format(Now, "dd/mm/yyyy")
DTPicker3.Value = Format(Now, "dd/mm/yyyy")
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Combo3.ListIndex = -1
Label21.Caption = ""

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
d = IdEmpla
FrmListadoEmpleados.Show 1
If d <> IdEmpla Then
Call CargaTrabajador
End If
End If
End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub


 
Private Sub TxtAbonado_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub TxtADeuda_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)

If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0

If KeyAscii = 13 Then Call BtnBuscar_Click

End Sub

Private Sub Timer1_Timer()

If IsNumeric(TxtPrestamo.Text) Then
    If Trim(TxtPrestamo.Text) = "" Then TxtPrestamo.Text = "0"
Else
    TxtPrestamo.Text = "0"
End If

If IsNumeric(TxtInteres.Text) Then
    If Trim(TxtInteres.Text) = "" Then TxtInteres.Text = "0"
Else
    TxtInteres.Text = "0"
End If

a = CDbl(TxtPrestamo.Text) 'Monto del prestamo
b = CDbl(TxtInteres.Text) '% de intereses del prestamo
d = (a * b) / 100 + a 'deuda con interes
f = (a * b) / 100
e = d - C

TxtMontoInteres.Text = Format(f, "#,##0.00")

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
Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub TxtInteres_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub TxtMontoInteres_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub TxtMxC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtMxC.Text) <> "" Then
        If IsNumeric(TxtMxC.Text) Then
            If Trim(TxtPrestamo.Text) <> "" Then
                If IsNumeric(TxtPrestamo.Text) Then
                
                    If CDbl(TxtPrestamo.Text) < CDbl(TxtMxC.Text) Then TxtMxC.Text = "0,00": Exit Sub
                    Dim NCuotas As Integer
                    Dim Entero As Double
                    Dim i As Integer
                    Entero = Round(CDbl(TxtPrestamo.Text) / CDbl(TxtMxC.Text), 4)
                    
                    If Val(Entero) < Entero Then
                        NCuotas = Val(Entero) + 1
                    Else
                        NCuotas = Val(Entero)
                    End If
                    
                    If NCuotas > 10 Then
                        Combo3.List(9) = "Definida (" & NCuotas & ")"
                        Combo3.ItemData(9) = NCuotas
                        Combo3.ListIndex = 9
                    Else
                        For i = 0 To Combo3.ListCount
                            If Combo3.ItemData(i) = NCuotas Then
                                Combo3.ListIndex = i
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    End If
ElseIf KeyAscii = 46 Then
    If InStr(1, TxtMxC.Text, ",") Then
        KeyAscii = 0
    Else
        KeyAscii = 44
    End If
End If
End Sub

Private Sub TxtPrestamo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


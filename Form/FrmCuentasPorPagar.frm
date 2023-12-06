VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCuentasPorPagar 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas por Pagar"
   ClientHeight    =   8745
   ClientLeft      =   1980
   ClientTop       =   1785
   ClientWidth     =   15465
   Icon            =   "FrmCuentasPorPagar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15465
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   8655
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   15255
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Abonos"
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   6000
         Width           =   4095
         Begin MSComctlLib.ListView LstAbonosCuentasPagar 
            Height          =   2175
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3836
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Fecha"
               Object.Width           =   3087
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Abonos"
               Object.Width           =   3087
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   2535
         Left            =   4320
         TabIndex        =   16
         Top             =   6000
         Width           =   10815
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Filtros"
            Height          =   1575
            Left            =   2400
            TabIndex        =   26
            Top             =   120
            Width           =   5535
            Begin VB.TextBox TxtDtpFecha 
               Height          =   375
               Left            =   1080
               TabIndex        =   35
               ToolTipText     =   "Ingrese o Seleccione la fecha para realizar la busqueda"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox TxtNoFactura 
               Height          =   285
               Left            =   3480
               TabIndex        =   8
               ToolTipText     =   "Ingrese el número de factura para realizar la busqueda"
               Top             =   1005
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker DtpFecha 
               Height          =   375
               Left            =   840
               TabIndex        =   7
               Top             =   960
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Format          =   55640065
               CurrentDate     =   40142
            End
            Begin VB.TextBox TxtNombre 
               Height          =   285
               Left            =   840
               TabIndex        =   6
               ToolTipText     =   "Ingrese el nombre para realizar la busqueda"
               Top             =   600
               Width           =   2775
            End
            Begin VB.TextBox TxtCodigo 
               Height          =   285
               Left            =   840
               TabIndex        =   4
               ToolTipText     =   "Ingrese el código para realizar la busqueda"
               Top             =   240
               Width           =   1215
            End
            Begin ChamaleonButton.ChameleonBtn BtnBuscar 
               Height          =   495
               Left            =   4320
               TabIndex        =   5
               ToolTipText     =   "Buscar"
               Top             =   240
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
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
               MICON           =   "FrmCuentasPorPagar.frx":1002
               PICN            =   "FrmCuentasPorPagar.frx":101E
               PICH            =   "FrmCuentasPorPagar.frx":1283
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnLimpiar 
               Height          =   495
               Left            =   3720
               TabIndex        =   36
               ToolTipText     =   "Limpiar Campos"
               Top             =   240
               Width           =   495
               _ExtentX        =   873
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
               FCOLO           =   0
               MCOL            =   16777215
               MPTR            =   1
               MICON           =   "FrmCuentasPorPagar.frx":1515
               PICN            =   "FrmCuentasPorPagar.frx":1531
               PICH            =   "FrmCuentasPorPagar.frx":1791
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
               Left            =   4920
               Top             =   960
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
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. Factura:"
               Height          =   195
               Left            =   2520
               TabIndex        =   34
               Top             =   1050
               Width           =   885
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha:"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre:"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   645
               Width           =   600
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   285
               Width           =   540
            End
         End
         Begin VB.Frame FrameAbonos 
            BackColor       =   &H00EAEFEF&
            Enabled         =   0   'False
            Height          =   1575
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   2175
            Begin MSComCtl2.DTPicker DtpFechaAbono 
               Height          =   375
               Left            =   720
               TabIndex        =   3
               Top             =   840
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   55640065
               CurrentDate     =   40141
            End
            Begin VB.TextBox TxtAbono 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   720
               TabIndex        =   2
               Text            =   "0.00"
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha:"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   930
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Abono:"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   450
               Width           =   510
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00EAEFEF&
            Height          =   735
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   10575
            Begin ChamaleonButton.ChameleonBtn BtnCerrar 
               Height          =   375
               Left            =   9480
               TabIndex        =   14
               ToolTipText     =   "Cerrar Tablas de Pacientes"
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
               MICON           =   "FrmCuentasPorPagar.frx":1A1A
               PICN            =   "FrmCuentasPorPagar.frx":1A36
               PICH            =   "FrmCuentasPorPagar.frx":1BFF
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
               TabIndex        =   10
               ToolTipText     =   "Guardar / Actualizar Pacientes"
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
               MICON           =   "FrmCuentasPorPagar.frx":1E34
               PICN            =   "FrmCuentasPorPagar.frx":1E50
               PICH            =   "FrmCuentasPorPagar.frx":20DF
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
               TabIndex        =   9
               ToolTipText     =   "Agregar Pacientes"
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
               MICON           =   "FrmCuentasPorPagar.frx":2520
               PICN            =   "FrmCuentasPorPagar.frx":253C
               PICH            =   "FrmCuentasPorPagar.frx":26C9
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
               Left            =   8280
               TabIndex        =   13
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
               MICON           =   "FrmCuentasPorPagar.frx":28FE
               PICN            =   "FrmCuentasPorPagar.frx":291A
               PICH            =   "FrmCuentasPorPagar.frx":2BFC
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
               Left            =   6600
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
               MICON           =   "FrmCuentasPorPagar.frx":2E4D
               PICN            =   "FrmCuentasPorPagar.frx":2E69
               PICH            =   "FrmCuentasPorPagar.frx":2F8E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnEditorFacturas 
               Height          =   375
               Left            =   2520
               TabIndex        =   11
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Editor Facturas"
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
               MICON           =   "FrmCuentasPorPagar.frx":321E
               PICN            =   "FrmCuentasPorPagar.frx":323A
               PICH            =   "FrmCuentasPorPagar.frx":34D9
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnRetenciones 
               Height          =   375
               Left            =   4560
               TabIndex        =   37
               ToolTipText     =   "Agregar"
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Retenciones"
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
               MICON           =   "FrmCuentasPorPagar.frx":390E
               PICN            =   "FrmCuentasPorPagar.frx":392A
               PICH            =   "FrmCuentasPorPagar.frx":3AB7
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Por Pagar:"
            Height          =   195
            Left            =   8220
            TabIndex        =   30
            Top             =   1320
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total General:"
            Height          =   195
            Left            =   7965
            TabIndex        =   29
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label LblPorPagar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   25
            Top             =   1290
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos:"
            Height          =   195
            Left            =   8205
            TabIndex        =   28
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal:"
            Height          =   195
            Left            =   8280
            TabIndex        =   27
            Top             =   240
            Width           =   690
         End
         Begin VB.Label LblTotalGeneral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   24
            Top             =   930
            Width           =   1695
         End
         Begin VB.Label LblImpuestos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   23
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label LblSubTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   22
            Top             =   210
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView LstCuentasPagar 
         Height          =   5775
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   10186
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Proveedor"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "No. Factura"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "No. Control"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Dias Credito"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total General"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Descuentos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Por Pagar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCuentasPorPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsSeleccionarCuentasPorPagar As New ADODB.Recordset
Dim RsAbonosCtaXPagar As New ADODB.Recordset
Dim RsCargarCuentasPorPagar As New ADODB.Recordset
Dim RsTotalesCuentasPorPagar As New ADODB.Recordset

' Variables para la CxPagar
Public NoComp, NoOrd, NoFact, NoCont, FechaEmi, IdProv, PorPag, TotalG, CondPago, DiasCredit
Dim NombProv
Dim Desc

Private Sub BtnAgregar_Click()

If IdProv <> "" And NoFact <> "" Then
    If CDbl(PorPag) <= 0 Then MsgBox "El monto total de la Factura ha sido Pagado!", vbExclamation + vbOKOnly, "Informacion": BtnRetenciones.Enabled = False: Exit Sub
    FrameAbonos.Enabled = True
    BtnGuardarActualizar.Enabled = True
    TxtAbono.SetFocus
End If
End Sub

Private Sub BtnAyuda_Click()
Cancelar_Pago
End Sub

Private Sub BtnBuscar_Click()
Dim RsBuscar As New ADODB.Recordset
Dim RsBuscar2 As New ADODB.Recordset

Cancelar_Pago

If TxtCodigo.Text <> "" Or TxtNombre.Text <> "" Or TxtNoFactura.Text <> "" Or TxtDtpFecha.Text <> "" Then
    
    CSql = "Select * From CtaPorPagar Where IdProveedor='" & Trim(TxtCodigo.Text) & "' or Nombre='" & Trim(TxtNombre.Text) & "' or FechaEmision = '" & Trim(TxtDtpFecha.Text) & "' or NoFactura='" & Trim(TxtNoFactura.Text) & "'"
    Set RsBuscar = CrearRS(CSql)

    LstCuentasPagar.ListItems.Clear

    Do While Not RsBuscar.EOF
        With LstCuentasPagar
        i = i + 1
    
        .ListItems.Add , , RsBuscar.Fields("IdProveedor").Value
        .ListItems(i).ListSubItems.Add , , RsBuscar.Fields("Nombre").Value
        .ListItems(i).ListSubItems.Add , , RsBuscar.Fields("FechaEmision").Value
        .ListItems(i).ListSubItems.Add , , RsBuscar.Fields("NoFactura").Value
        .ListItems(i).ListSubItems.Add , , RsBuscar.Fields("NoControl").Value
        .ListItems(i).ListSubItems.Add , , RsBuscar.Fields("DiasCredito").Value
        .ListItems(i).ListSubItems.Add , , Format(RsBuscar.Fields("TotalGeneral").Value, "#,##0.00")
        .ListItems(i).ListSubItems.Add , , Format(RsBuscar.Fields("Descuento").Value, "#,##0.00")
        .ListItems(i).ListSubItems.Add , , Format(RsBuscar.Fields("PorPagar").Value, "#,##0.00")
    
        If Val(RsBuscar.Fields("CondicionPago").Value) = 0 Then
            If CDbl(RsBuscar.Fields("PorPagar").Value) <= 0 Then
                .ListItems(i).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsBuscar.Fields("FechaEmision").Value) - DateValue(Now)) >= 0 Then
                    .ListItems(i).ListSubItems.Add , , "Pendiente"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    .ListItems(i).ListSubItems.Add , , "Vencida"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                End If
            End If
        ElseIf Val(RsBuscar.Fields("CondicionPago").Value) = 1 Then
            If CDbl(RsBuscar.Fields("PorPagar").Value) <= 0 Then
                .ListItems(i).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsBuscar.Fields("FechaEmision").Value) - DateValue(Now)) = -30 Then
                    .ListItems(i).ListSubItems.Add , , "Pendiente"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    If (DateValue(RsBuscar.Fields("FechaEmision").Value) - DateValue(Now)) < -30 Then
                        .ListItems(i).ListSubItems.Add , , "Vencida"
                        .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                    Else
                        .ListItems(i).ListSubItems.Add , , "Vigente"
                    End If
                End If
            End If
        Else
            If CDbl(RsBuscar.Fields("PorPagar").Value) <= 0 Then
                .ListItems(i).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsBuscar.Fields("FechaEmision").Value) - DateValue(Now)) = -45 Then
                    .ListItems(i).ListSubItems.Add , , "Pendiente"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    If (DateValue(RsBuscar.Fields("FechaEmision").Value) - DateValue(Now)) < -45 Then
                        .ListItems(i).ListSubItems.Add , , "Vencida"
                        .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                    Else
                        .ListItems(i).ListSubItems.Add , , "Vigente"
                    End If
                End If
            End If
        End If
        End With
        RsBuscar.MoveNext
    Loop
    RsBuscar.Close


    CSql = "Select sum(SubTotal) as TSubTotal, sum(Impuesto) as TImpuesto, Sum(TotalGeneral) as TotalG, sum(PorPagar) as TPorPagar From CtaPorPagar  Where IdProveedor='" & Trim(TxtCodigo.Text) & "' or Nombre='" & Trim(TxtNombre.Text) & "' or FechaEmision = '" & TxtDtpFecha.Text & "' or NoFactura='" & Trim(TxtNoFactura.Text) & "'"
    Set RsBuscar2 = CrearRS(CSql)

    If RsBuscar2.RecordCount > 0 Then
        LblSubTotal.Caption = Format(RsBuscar2.Fields("TSubTotal").Value, "#,##0.00")
        LblImpuestos.Caption = Format(RsBuscar2.Fields("TImpuesto").Value, "#,##0.00")
        LblTotalGeneral.Caption = Format(RsBuscar2.Fields("TotalG").Value, "#,##0.00")
        LblPorPagar.Caption = Format(RsBuscar2.Fields("TPorPagar").Value, "#,##0.00")
    End If
End If


End Sub

Private Sub BtnCerrar_Click()
Cancelar_Pago
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
IdProv = ""
NoFact = ""
Cancelar_Pago
CargarCuentasPorPagar
End Sub

Private Sub BtnEditorFacturas_Click()
Cancelar_Pago
If IdProv <> "" And NoFact <> "" Then
    FrmEditorFacturas.Show vbModal, FrmPrincipal
Else
    MsgBox "No hay factura seleccionada para modificar", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim PG
Dim Seleccion As Integer

Seleccion = LstCuentasPagar.SelectedItem.ListSubItems(3).Text


If CDbl(TxtAbono.Text) <= 0 Then
    MsgBox "Ingrese un monto mayor de 0 (Cero)", vbExclamation + vbOKOnly, "Monto no valido!"
    Exit Sub
    Else
    If (CDbl(PorPag) - CDbl(TxtAbono.Text)) <= -0.01 Then
        MsgBox "El Monto a Pagar es Mayor a la Deuda!", vbExclamation + vbOKOnly, "Informacion"
        Exit Sub
    End If
End If


CSql = "Select * From AbonoCtaXPagar"
Set RsAbonosCtaXPagar = CrearRS(CSql)

RsAbonosCtaXPagar.AddNew
RsAbonosCtaXPagar.Fields("NumeroCompra").Value = NoComp
RsAbonosCtaXPagar.Fields("NumeroOrden").Value = NoOrd
RsAbonosCtaXPagar.Fields("NoFactura").Value = NoFact
RsAbonosCtaXPagar.Fields("NoControl").Value = NoCont
RsAbonosCtaXPagar.Fields("FechaEmision").Value = FechaEmi
RsAbonosCtaXPagar.Fields("IdProveedor").Value = IdProv
RsAbonosCtaXPagar.Fields("TotalGeneral").Value = TotalG
PG = CDbl(PorPag) - CDbl(TxtAbono.Text)
RsAbonosCtaXPagar.Fields("PorPagar").Value = PG
RsAbonosCtaXPagar.Fields("FechaAbono").Value = DtpFechaAbono.Value
RsAbonosCtaXPagar.Fields("MontoAbono").Value = CDbl(TxtAbono.Text)
RsAbonosCtaXPagar.Update

RsAbonosCtaXPagar.Close

Dim RsUpdatePorPagar As New ADODB.Recordset
CSql = "Select * From CtaPorPagar where IdProveedor='" & IdProv & "' And NoFactura='" & NoFact & "'"
Set RsUpdatePorPagar = CrearRS(CSql)

RsUpdatePorPagar.Fields("PorPagar").Value = PG
PorPag = PG
RsUpdatePorPagar.Update

MsgBox "Los Cambios han sido registrados sobre la factura No-" & NoFact, vbInformation + vbOKOnly, "Operacion Exitosa!"


CargarAbonos
CargarCuentasPorPagar
TotalesCuentasPorPagar
Cancelar_Pago

For i = 1 To LstCuentasPagar.ListItems.Count
    If LstCuentasPagar.ListItems(i).ListSubItems(3).Text = Seleccion Then
        LstCuentasPagar.ListItems(i).Selected = True
        LstCuentasPagar.ListItems(i).EnsureVisible
    End If
Next i

mensaje = MsgBox("Este Pago aplica algun tipo de retención", vbInformation + vbYesNo, "Información")

If mensaje = vbYes Then
    BtnRetenciones_Click
End If

End Sub

Private Sub BtnImprimir_Click()

Cancelar_Pago

''========= ESTE ES EL CODIGO NUEVO ==========
With CrystalReport1
    .ReportFileName = RutaInformes & "\CuentasPorPagar.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    
    If TxtCodigo.Text <> "" And TxtNombre.Text = "" And TxtNoFactura.Text = "" And TxtDtpFecha.Text = "" Then
        .SelectionFormula = "{CuentasPorPagar.IdProveedor} = " & Trim(TxtCodigo.Text)
    End If
    
    If TxtNombre.Text <> "" And TxtCodigo.Text = "" And TxtNoFactura.Text = "" And TxtDtpFecha.Text = "" Then
        .SelectionFormula = "{CuentasPorPagar.Nombre} = '" & Trim(TxtNombre.Text) & "'"
    End If
    
    If TxtNoFactura.Text <> "" And TxtNombre.Text = "" And TxtCodigo.Text = "" And TxtDtpFecha.Text = "" Then
        .SelectionFormula = "{CuentasPorPagar.NoFactura} = '" & Trim(TxtNoFactura.Text) & "'"
    End If
    
    If TxtDtpFecha.Text <> "" And TxtCodigo.Text = "" And TxtNombre.Text = "" And TxtNoFactura.Text = "" Then
        .SelectionFormula = "{CuentasPorPagar.FechaEmision} = " & FechaSQL(TxtDtpFecha.Text)
    End If
    
    .WindowTitle = "Reporte Cuentas por Pagar"
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

End Sub

Private Sub BtnLimpiar_Click()
Cancelar_Pago
TxtCodigo.Text = ""
TxtNombre.Text = ""
TxtDtpFecha.Text = ""
TxtNoFactura.Text = ""
End Sub

Private Sub BtnRetenciones_Click()

If NoFact = "" Then
    BtnRetenciones.Enabled = False
    MsgBox "Debe de seleccionar una factura para poder realizar las retenciones", vbCritical + vbOKOnly, "Mensaje de Error"
ElseIf TotalG = "0" Then
    BtnRetenciones.Enabled = False
    MsgBox "No se pueden realizar retenciones a la factura por no poseer monto", vbCritical + vbOKOnly, "Mensaje de Error"
Else
    BtnRetenciones.Enabled = True
    FrmComprobanteRetencion.Show vbModal, FrmPrincipal
End If

End Sub

Private Sub DtpFecha_Change()
DtpFecha_Click
End Sub

Private Sub DtpFecha_Click()
Cancelar_Pago
TxtDtpFecha.Text = DtpFecha.Value
End Sub

Private Sub DtpFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub DtpFecha_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoFactura.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyLeft
            If FrameAbonos.Enabled = True Then DtpFechaAbono.SetFocus
        Case vbKeyRight
            TxtNoFactura.SetFocus
        Case vbKeyDown
            BtnEditorFacturas.SetFocus
    End Select
End If
End Sub

Private Sub DtpFechaAbono_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGuardarActualizar.SetFocus
        Case vbKeyUp
            TxtAbono.SetFocus
        Case vbKeyRight
            TxtDtpFecha.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me
DtpFechaAbono.Value = Format(Now, "DD/MM/YY")
BtnGuardarActualizar.Enabled = False
CargarCuentasPorPagar
TotalesCuentasPorPagar
End Sub

Sub TotalesCuentasPorPagar()
CSql = "Select sum(SubTotal) as TSubTotal, sum(Impuesto) as TImpuesto, Sum(TotalGeneral) as TTotalG, sum(PorPagar) as TPorPagar From CtaPorPagar"
Set RsTotalesCuentasPorPagar = CrearRS(CSql)

LblSubTotal.Caption = Format(RsTotalesCuentasPorPagar.Fields("TSubTotal").Value, "#,##0.00")
LblImpuestos.Caption = Format(RsTotalesCuentasPorPagar.Fields("TImpuesto").Value, "#,##0.00")
LblTotalGeneral.Caption = Format(RsTotalesCuentasPorPagar.Fields("TTotalG").Value, "#,##0.00")
LblPorPagar.Caption = Format(RsTotalesCuentasPorPagar.Fields("TPorPagar").Value, "#,##0.00")
RsTotalesCuentasPorPagar.Close
End Sub
Sub CargarCuentasPorPagar()

CSql = "Select * From CtaPorPagar"
Set RsCargarCuentasPorPagar = CrearRS(CSql)

LstCuentasPagar.ListItems.Clear

Do While Not RsCargarCuentasPorPagar.EOF
    With LstCuentasPagar
        i = i + 1
        .ListItems.Add , , RsCargarCuentasPorPagar.Fields("IdProveedor").Value
        .ListItems(i).ListSubItems.Add , , RsCargarCuentasPorPagar.Fields("Nombre").Value
        .ListItems(i).ListSubItems.Add , , RsCargarCuentasPorPagar.Fields("FechaEmision").Value
        .ListItems(i).ListSubItems.Add , , RsCargarCuentasPorPagar.Fields("NoFactura").Value
        .ListItems(i).ListSubItems.Add , , RsCargarCuentasPorPagar.Fields("NoControl").Value
        .ListItems(i).ListSubItems.Add , , RsCargarCuentasPorPagar.Fields("DiasCredito").Value
        .ListItems(i).ListSubItems.Add , , Format(RsCargarCuentasPorPagar.Fields("TotalGeneral").Value, "#,##0.00")
        .ListItems(i).ListSubItems.Add , , Format(RsCargarCuentasPorPagar.Fields("Descuento").Value, "#,##0.00")
        .ListItems(i).ListSubItems.Add , , Format(RsCargarCuentasPorPagar.Fields("PorPagar").Value, "#,##0.00")
        
        If Val(RsCargarCuentasPorPagar.Fields("CondicionPago").Value) = 0 Then
            If CDbl(RsCargarCuentasPorPagar.Fields("PorPagar").Value) <= 0 Then
                .ListItems(i).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsCargarCuentasPorPagar.Fields("FechaEmision").Value) - DateValue(Now)) >= 0 Then
                    .ListItems(i).ListSubItems.Add , , "Pendiente"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    .ListItems(i).ListSubItems.Add , , "Vencida"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                End If
            End If
        ElseIf Val(RsCargarCuentasPorPagar.Fields("CondicionPago").Value) = 1 Then
            If CDbl(RsCargarCuentasPorPagar.Fields("PorPagar").Value) <= 0 Then
                .ListItems(i).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsCargarCuentasPorPagar.Fields("FechaEmision").Value) - DateValue(Now)) = -30 Then
                    .ListItems(i).ListSubItems.Add , , "Pendiente"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    If (DateValue(RsCargarCuentasPorPagar.Fields("FechaEmision").Value) - DateValue(Now)) < -30 Then
                        .ListItems(i).ListSubItems.Add , , "Vencida"
                        .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                    Else
                        .ListItems(i).ListSubItems.Add , , "Vigente"
                    End If
                End If
            End If
        Else
            If CDbl(RsCargarCuentasPorPagar.Fields("PorPagar").Value) <= 0 Then
                .ListItems(i).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsCargarCuentasPorPagar.Fields("FechaEmision").Value) - DateValue(Now)) = -45 Then
                    .ListItems(i).ListSubItems.Add , , "Pendiente"
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    If (DateValue(RsCargarCuentasPorPagar.Fields("FechaEmision").Value) - DateValue(Now)) < -45 Then
                        .ListItems(i).ListSubItems.Add , , "Vencida"
                        .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                    Else
                        .ListItems(i).ListSubItems.Add , , "Vigente"
                    End If
                End If
            End If
        End If
    End With
    DoEvents
    RsCargarCuentasPorPagar.MoveNext
Loop
RsCargarCuentasPorPagar.Close
TotalesCuentasPorPagar
End Sub

Private Sub LstAbonosCuentasPagar_DblClick()
Dim RsUpdate As New ADODB.Recordset
Dim MAbono, FAbono As String
Dim Seleccion

    
    If LstAbonosCuentasPagar.ListItems.Count = 0 Then Exit Sub

    MAbono = LstAbonosCuentasPagar.SelectedItem.ListSubItems(1).Text
    FAbono = LstAbonosCuentasPagar.SelectedItem.Text
    Seleccion = MsgBox("Se procedera a eliminar el Abono seleccionado de fecha " & FAbono & " sobre la Factura No-" & NoFact & "", vbQuestion + vbYesNo, "Confimar Operacion")
    
    If Seleccion = vbNo Then Exit Sub
    
    Seleccion = LstCuentasPagar.SelectedItem.ListSubItems(3).Text
    
    ' Actualiza datos de la tabla AbonoCtaXCobrar
    CSql = "Select * From AbonoCtaXPagar Where NoFactura='" & NoFact & "' AND FechaEmision='" & FechaEmi & "' AND " _
            & " MontoAbono=" & Replace(CDbl(MAbono), ",", ".") & " AND FechaAbono ='" & FAbono & "'"
    Set RsUpdate = CrearRS(CSql)

    RsUpdate.Delete
    'DELETE FROM AbonoCtaXCobrar
    PorPag = CDbl(PorPag) + CDbl(MAbono)
    RsUpdate.Update


    Dim RsUpdatePorPagar As New ADODB.Recordset
    CSql = "Select * From CtaPorPagar where NoFactura='" & NoFact & "'"
    Set RsUpdatePorPagar = CrearRS(CSql)

    RsUpdatePorPagar.Fields("PorPagar").Value = PorPag
    RsUpdatePorPagar.Update

    MsgBox "Los Cambios han sido registrados!", vbInformation + vbOKOnly, "Operacion Exitosa!"

    CargarAbonos
    CargarCuentasPorPagar
    Cancelar_Pago

    For i = 1 To LstCuentasPagar.ListItems.Count
    If LstCuentasPagar.ListItems(i).ListSubItems(3).Text = Seleccion Then
        LstCuentasPagar.ListItems(i).Selected = True
        LstCuentasPagar.ListItems(i).EnsureVisible
    End If
Next i
End Sub

Private Sub LstCuentasPagar_Click()

BtnRetenciones.Enabled = True
' Inicializa los objectos
Cancelar_Pago

' Crea la Consulta
CSql = "Select * From CtaPorPagar where IdProveedor='" & LstCuentasPagar.SelectedItem.Text & "' And NoFactura='" & LstCuentasPagar.SelectedItem.ListSubItems(3).Text & "'"
Set RsSeleccionarCuentasPorPagar = CrearRS(CSql)

If RsSeleccionarCuentasPorPagar.RecordCount > 0 Then
    NoOrd = RsSeleccionarCuentasPorPagar.Fields("NumeroOrden").Value
    NoComp = RsSeleccionarCuentasPorPagar.Fields("NumeroCompra").Value
    IdProv = RsSeleccionarCuentasPorPagar.Fields("IdProveedor").Value
    NombProv = RsSeleccionarCuentasPorPagar.Fields("Nombre").Value
    FechaEmi = RsSeleccionarCuentasPorPagar.Fields("FechaEmision").Value
    NoFact = RsSeleccionarCuentasPorPagar.Fields("NoFactura").Value
    NoCont = RsSeleccionarCuentasPorPagar.Fields("NoControl").Value
    DiasCredit = RsSeleccionarCuentasPorPagar.Fields("DiasCredito").Value
    CondPago = RsSeleccionarCuentasPorPagar.Fields("CondicionPago").Value
    TotalG = RsSeleccionarCuentasPorPagar.Fields("TotalGeneral").Value
    Desc = RsSeleccionarCuentasPorPagar.Fields("Descuento").Value
    PorPag = RsSeleccionarCuentasPorPagar.Fields("PorPagar").Value
End If
RsSeleccionarCuentasPorPagar.Close

CargarAbonos

If TotalG = "0" Then
    BtnRetenciones.Enabled = False
End If

If NoFact = "" Then
    BtnRetenciones.Enabled = False
End If

End Sub

Sub CargarAbonos()
Dim RsCargarAbonos As New ADODB.Recordset

CSql = "Select * From AbonoCtaXPagar Where IdProveedor='" & IdProv & "' And NoFactura='" & NoFact & "'"
Set RsCargarAbonos = CrearRS(CSql)

LstAbonosCuentasPagar.ListItems.Clear

Do While Not RsCargarAbonos.EOF
    With LstAbonosCuentasPagar
        i = i + 1
        .ListItems.Add , , RsCargarAbonos.Fields("FechaAbono").Value
        .ListItems(i).ListSubItems.Add , , Format(RsCargarAbonos.Fields("MontoAbono").Value, "#,##0.00")
    End With
    RsCargarAbonos.MoveNext
Loop
End Sub

Sub Cancelar_Pago()
    FrameAbonos.Enabled = False
    BtnGuardarActualizar.Enabled = False
    TxtAbono.Text = "0.00"
End Sub

Private Sub LstCuentasPagar_DblClick()
    BtnEditorFacturas_Click
End Sub

Private Sub LstCuentasPagar_KeyUp(KeyCode As Integer, Shift As Integer)
LstCuentasPagar_Click
End Sub

Private Sub TxtAbono_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFechaAbono.SetFocus
        Case vbKeyUp
            LstCuentasCobrar.SetFocus
        Case vbKeyRight
            TxtCodigo.SetFocus
        Case vbKeyDown
            DtpFechaAbono.SetFocus
    End Select
End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)

Cancelar_Pago
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNombre.SetFocus
        Case vbKeyLeft
            If FrameAbonos.Enabled = True Then DtpFechaAbono.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
        Case vbKeyDown
            TxtNombre.SetFocus
    End Select
End If
End Sub


Private Sub TxtDtpFecha_KeyUp(KeyCode As Integer, Shift As Integer)

Cancelar_Pago
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyLeft
            DtpFecha.SetFocus
        Case vbKeyRight
            TxtNoFactura.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoFactura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtNoFactura_KeyUp(KeyCode As Integer, Shift As Integer)
Cancelar_Pago
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
        Case vbKeyUp
            BtnBuscar.SetFocus
        Case vbKeyLeft
            TxtDtpFecha.SetFocus
        Case vbKeyDown
            BtnImprimir.SetFocus
    End Select
End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub


Private Sub TxtNombre_KeyUp(KeyCode As Integer, Shift As Integer)

Cancelar_Pago
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DtpFecha.SetFocus
        Case vbKeyUp
            TxtCodigo.SetFocus
        Case vbKeyLeft
            If FrameAbonos.Enabled = True Then DtpFechaAbono.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
        Case vbKeyDown
            TxtDtpFecha.SetFocus
    End Select
End If
End Sub

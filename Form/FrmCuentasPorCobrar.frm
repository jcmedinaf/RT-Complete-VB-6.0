VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCuentasPorCobrar 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas por Cobrar"
   ClientHeight    =   8760
   ClientLeft      =   1980
   ClientTop       =   1965
   ClientWidth     =   15495
   Icon            =   "FrmCuentasPorCobrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15495
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   2535
         Left            =   4320
         TabIndex        =   3
         Top             =   6000
         Width           =   10815
         Begin VB.Frame Frame9 
            BackColor       =   &H00EAEFEF&
            Height          =   735
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   10575
            Begin Crystal.CrystalReport CrystalReport1 
               Left            =   2520
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
               WindowShowRefreshBtn=   -1  'True
            End
            Begin ChamaleonButton.ChameleonBtn BtnCerrar 
               Height          =   375
               Left            =   9480
               TabIndex        =   20
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
               MICON           =   "FrmCuentasPorCobrar.frx":1002
               PICN            =   "FrmCuentasPorCobrar.frx":101E
               PICH            =   "FrmCuentasPorCobrar.frx":11E7
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
               TabIndex        =   21
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
               MICON           =   "FrmCuentasPorCobrar.frx":141C
               PICN            =   "FrmCuentasPorCobrar.frx":1438
               PICH            =   "FrmCuentasPorCobrar.frx":16C7
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
               TabIndex        =   22
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
               MICON           =   "FrmCuentasPorCobrar.frx":1B08
               PICN            =   "FrmCuentasPorCobrar.frx":1B24
               PICH            =   "FrmCuentasPorCobrar.frx":1CB1
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
               TabIndex        =   23
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
               MICON           =   "FrmCuentasPorCobrar.frx":1EE6
               PICN            =   "FrmCuentasPorCobrar.frx":1F02
               PICH            =   "FrmCuentasPorCobrar.frx":21E4
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
               Left            =   5760
               TabIndex        =   24
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
               MICON           =   "FrmCuentasPorCobrar.frx":2435
               PICN            =   "FrmCuentasPorCobrar.frx":2451
               PICH            =   "FrmCuentasPorCobrar.frx":2576
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
               Left            =   3240
               TabIndex        =   25
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Editor Cuenta"
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
               MICON           =   "FrmCuentasPorCobrar.frx":2806
               PICN            =   "FrmCuentasPorCobrar.frx":2822
               PICH            =   "FrmCuentasPorCobrar.frx":2AC1
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
         Begin VB.Frame FrameAbonos 
            BackColor       =   &H00EAEFEF&
            Enabled         =   0   'False
            Height          =   1455
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   2175
            Begin VB.TextBox TxtAbono 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   720
               TabIndex        =   16
               Text            =   "0.00"
               Top             =   360
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DtpFechaAbono 
               Height          =   375
               Left            =   720
               TabIndex        =   15
               Top             =   840
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   53936129
               CurrentDate     =   40141
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Abono:"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   450
               Width           =   510
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha:"
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   930
               Width           =   495
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Filtros"
            Height          =   1455
            Left            =   2400
            TabIndex        =   4
            Top             =   120
            Width           =   5415
            Begin VB.TextBox TxtDtpFecha 
               Height          =   375
               Left            =   1080
               TabIndex        =   35
               ToolTipText     =   "Ingrese o Seleccione la fecha para realizar la busqueda"
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox TxtCodigo 
               Height          =   285
               Left            =   840
               TabIndex        =   8
               ToolTipText     =   "Ingrese el Codigo para realizar la busqueda"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtNombre 
               Height          =   285
               Left            =   840
               TabIndex        =   7
               ToolTipText     =   "Ingrese el Nombre para realizar la busqueda"
               Top             =   600
               Width           =   2655
            End
            Begin VB.TextBox TxtNoFactura 
               Height          =   285
               Left            =   3480
               TabIndex        =   5
               ToolTipText     =   "Ingrese el número de factura para poder realizar la busqueda"
               Top             =   1005
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker DtpFecha 
               Height          =   375
               Left            =   840
               TabIndex        =   6
               Top             =   960
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Format          =   53936129
               CurrentDate     =   40142
            End
            Begin ChamaleonButton.ChameleonBtn BtnBuscar 
               Height          =   375
               Left            =   4200
               TabIndex        =   9
               ToolTipText     =   "Buscar"
               Top             =   240
               Width           =   1095
               _ExtentX        =   1931
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
               MICON           =   "FrmCuentasPorCobrar.frx":2EF6
               PICN            =   "FrmCuentasPorCobrar.frx":2F12
               PICH            =   "FrmCuentasPorCobrar.frx":3177
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
               Height          =   375
               Left            =   3600
               TabIndex        =   36
               ToolTipText     =   "Limpiar Campos"
               Top             =   240
               Width           =   495
               _ExtentX        =   873
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
               FCOLO           =   0
               MCOL            =   16777215
               MPTR            =   1
               MICON           =   "FrmCuentasPorCobrar.frx":3409
               PICN            =   "FrmCuentasPorCobrar.frx":3425
               PICH            =   "FrmCuentasPorCobrar.frx":3685
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
               Caption         =   "Código:"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   285
               Width           =   540
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre:"
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   645
               Width           =   600
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha:"
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   1050
               Width           =   495
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. Factura:"
               Height          =   195
               Left            =   2520
               TabIndex        =   10
               Top             =   1080
               Width           =   885
            End
         End
         Begin VB.Label LblSubTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   33
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label LblImpuestos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   32
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label LblTotalGeneral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   31
            Top             =   930
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal:"
            Height          =   195
            Left            =   8280
            TabIndex        =   30
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos:"
            Height          =   195
            Left            =   8205
            TabIndex        =   29
            Top             =   600
            Width           =   765
         End
         Begin VB.Label LblPorCobrar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   255
            Left            =   9000
            TabIndex        =   28
            Top             =   1290
            Width           =   1695
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total General:"
            Height          =   195
            Left            =   7965
            TabIndex        =   27
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Por Cobrar:"
            Height          =   195
            Left            =   8175
            TabIndex        =   26
            Top             =   1320
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Abonos"
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   6000
         Width           =   4095
         Begin MSComctlLib.ListView LstAbonosCuentasCobrar 
            Height          =   2175
            Left            =   120
            TabIndex        =   34
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
      Begin MSComctlLib.ListView LstCuentasCobrar 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   9975
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cliente"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Paciente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "No. Cedula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "No. Factura"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "SubTotal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Impuesto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Descuentos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Retenciones"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Timbres Fiscales"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Por Cobrar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCuentasPorCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCliente As New ADODB.Recordset
Dim RsPaciente As New ADODB.Recordset
Dim RsCuentasCobrar As New ADODB.Recordset
Dim NombCli
Public IdCliente, IdPaciente, NoFact, SubTotal, TasaImpuesto, Impuesto, Monto, PorCobrar, FechaEmi, FechaAbono
Dim FormaPago, Tipo, IdUsuario, impresa As String
Dim Anulada As String

Private Sub BtnAgregar_Click()
If IdCliente <> "" And NoFact <> "" Then
    If CDbl(PorCobrar) <= 0 Then MsgBox "El monto total de la Factura ha sido cancelado!", vbExclamation + vbOKOnly, "Informacion": Exit Sub
    BtnAgregar.Enabled = False
    FrameAbonos.Enabled = True
    BtnGuardarActualizar.Enabled = True
    TxtAbono.SetFocus
    Else
    MsgBox "Seleccione un Factura.", vbExclamation + vbOKOnly, "Error"
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

    IdCliente = ""
    If TxtNombre.Text <> "" Then
        CSql = "Select * From Cliente Where Razon ='" & Trim(TxtNombre.Text) & "'"
        Set RsBuscar = CrearRS(CSql)
        If Not RsBuscar.BOF Then IdCliente = RsBuscar.Fields("idcliente").Value
    Else
        IdCliente = Trim(TxtCodigo.Text)
    End If

    
    CSql = "Select * From C_cobrar Where (Forma_Pago<>0 or Forma_Pago<> -1) And (IdCliente='" & IdCliente & "'  or Fecha Like '" & TxtDtpFecha.Text & "' or N_Factura='" & Trim(TxtNoFactura.Text) & "') order by IdCliente, N_Factura"

    Set RsBuscar = CrearRS(CSql)
    LstCuentasCobrar.ListItems.Clear
    Do While Not RsBuscar.EOF
    
        CSql = "Select * From Cliente Where IdCliente ='" & RsBuscar.Fields("IdCliente").Value & "'"
        Set RsBuscarCliente = CrearRS(CSql)

        CSql = "Select * From Paciente Where IdPaciente ='" & RsBuscar.Fields("IdPaciente").Value & "'"
        Set RsBuscarPaciente = CrearRS(CSql)
        With LstCuentasCobrar
            C = C + 1
            .ListItems.Add , , RsBuscarCliente.Fields("Razon").Value
            .ListItems(C).ListSubItems.Add , , RsBuscarPaciente.Fields("ApellidoP").Value & ", " & RsBuscarPaciente.Fields("NombreP").Value
            .ListItems(C).ListSubItems.Add , , RsBuscarPaciente.Fields("CedulaP").Value
            .ListItems(C).ListSubItems.Add , , RsBuscar.Fields("N_Factura").Value
            .ListItems(C).ListSubItems.Add , , RsBuscar.Fields("Fecha").Value
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("SubTotal").Value, "#,##0.00")
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("Impuesto").Value, "#,##0.00")
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("Monto").Value, "#,##0.00")
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("Descuentos").Value, "#,##0.00")
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("Retenciones").Value, "#,##0.00")
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("TimbresFiscales").Value, "#,##0.00")
            .ListItems(C).ListSubItems.Add , , Format(RsBuscar.Fields("PorCobrar").Value, "#,##0.00")
            
            If Val(RsBuscar.Fields("Forma_Pago").Value) = 0 Then
                If CDbl(RsBuscar.Fields("PorCobrar").Value) <= 0 Then
                    .ListItems(C).ListSubItems.Add , , "Cancelada"
                Else
                    .ListItems(C).ListSubItems.Add , , "Vencida"
                    .ListItems(C).ListSubItems.Item(1).ForeColor = vbRed
                End If
            ElseIf Val(RsBuscar.Fields("Forma_Pago").Value) = 1 Then
                If CDbl(RsBuscar.Fields("PorCobrar").Value) <= 0 Then
                    .ListItems(C).ListSubItems.Add , , "Cancelada"
                Else
                    If (DateValue(RsBuscar.Fields("Fecha").Value) - DateValue(Now)) = -30 Then
                        .ListItems(C).ListSubItems.Add , , "Vence hoy"
                        .ListItems(C).ListSubItems.Item(1).ForeColor = vbBlue
                    Else
                        If (DateValue(RsBuscar.Fields("Fecha").Value) - DateValue(Now)) < -30 Then
                            .ListItems(C).ListSubItems.Add , , "Vencida"
                            .ListItems(C).ListSubItems.Item(1).ForeColor = vbRed
                        Else
                            .ListItems(C).ListSubItems.Add , , "Vigente"
                        End If
                    End If
                End If
            Else
                If CDbl(RsBuscar.Fields("PorCobrar").Value) <= 0 Then
                    .ListItems(C).ListSubItems.Add , , "Cancelada"
                Else
                    If (DateValue(RsBuscar.Fields("Fecha").Value) - DateValue(Now)) = -45 Then
                        .ListItems(C).ListSubItems.Add , , "Vence hoy"
                        .ListItems(C).ListSubItems.Item(1).ForeColor = vbBlue
                    Else
                        If (DateValue(RsBuscar.Fields("Fecha").Value) - DateValue(Now)) < -45 Then
                            .ListItems(C).ListSubItems.Add , , "Vencida"
                            .ListItems(C).ListSubItems.Item(1).ForeColor = vbRed
                        Else
                            .ListItems(C).ListSubItems.Add , , "Vigente"
                        End If
                    End If
                End If
            End If
        End With
        RsBuscar.MoveNext
    Loop

    CSql = "Select Sum(SubTotal) as TSubTotal, sum(Impuesto) as TImpuesto, Sum(Monto) as TMonto, sum(PorCobrar) as TPorCobrar  From C_Cobrar Where (Forma_Pago<>0 or Forma_Pago<> -1) And (IdCliente='" & IdCliente & "'  or Fecha Like '" & TxtDtpFecha.Text & "' or N_Factura='" & Trim(TxtNoFactura.Text) & "')"
    Set RsBuscar2 = CrearRS(CSql)

    If RsBuscar2.RecordCount > 0 Then
        LblSubTotal.Caption = Format(RsBuscar2.Fields("TSubTotal").Value, "#,##0.00")
        LblImpuestos.Caption = Format(RsBuscar2.Fields("TImpuesto").Value, "#,##0.00")
        LblTotalGeneral.Caption = Format(RsBuscar2.Fields("TMonto").Value, "#,##0.00")
        LblPorCobrar.Caption = Format(RsBuscar2.Fields("TPorCobrar").Value, "#,##0.00")
    End If
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
IdCliente = ""
NoFact = ""
Cancelar_Pago
CargarCuentasPorCobrar
End Sub

Private Sub BtnEditorFacturas_Click()
Cancelar_Pago
If IdCliente <> "" And NoFact <> "" Then
    FrmEditorFacturasCxC.Show vbModal, FrmPrincipal
Else
    MsgBox "No hay factura seleccionada para modificar", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim CxCobrar
Dim Seleccion As Integer

Seleccion = LstCuentasCobrar.SelectedItem.ListSubItems(3).Text

If CDbl(TxtAbono.Text) <= 0 Then
    MsgBox "Ingrese un monto mayor de 0 (Cero)", vbExclamation + vbOKOnly, "Monto no valido!"
    Exit Sub
    Else
    If (CDbl(PorCobrar) - CDbl(TxtAbono.Text)) <= -0.01 Then
        MsgBox "El Monto a Abonar es Mayor a la Deuda!", vbExclamation + vbOKOnly, "Informacion"
        Exit Sub
    End If
End If


CSql = "Select * From AbonoCtaXCobrar"
Set RsAbonoCtaXCobrar = CrearRS(CSql)

RsAbonoCtaXCobrar.AddNew
RsAbonoCtaXCobrar.Fields("NoFactura").Value = NoFact
RsAbonoCtaXCobrar.Fields("FechaEmision").Value = FechaEmi
RsAbonoCtaXCobrar.Fields("IdCliente").Value = IdCliente
RsAbonoCtaXCobrar.Fields("IdPaciente").Value = IdPaciente
RsAbonoCtaXCobrar.Fields("TotalGeneral").Value = Monto
CxCobrar = CDbl(PorCobrar) - CDbl(TxtAbono.Text)
RsAbonoCtaXCobrar.Fields("PorCobrar").Value = CxCobrar
RsAbonoCtaXCobrar.Fields("FechaAbono").Value = DtpFechaAbono.Value
RsAbonoCtaXCobrar.Fields("MontoAbono").Value = CDbl(TxtAbono.Text)

RsAbonoCtaXCobrar.Update
RsAbonoCtaXCobrar.Close


Dim RsUpdatePorCobrar As New ADODB.Recordset
CSql = "Select * From C_Cobrar where N_Factura='" & NoFact & "'"
Set RsUpdatePorCobrar = CrearRS(CSql)

RsUpdatePorCobrar.Fields("PorCobrar").Value = CxCobrar
PorCobrar = CxCobrar
RsUpdatePorCobrar.Update
MsgBox "Los Cambios han sido registrados sobre la factura No-" & NoFact, vbInformation + vbOKOnly, "Operacion Exitosa!"
CargarCuentasPorCobrar
CargarAbonos
Cancelar_Pago

For i = 1 To LstCuentasCobrar.ListItems.Count
    If LstCuentasCobrar.ListItems(i).ListSubItems(3).Text = Seleccion Then
        LstCuentasCobrar.ListItems(i).Selected = True
        LstCuentasCobrar.ListItems(i).EnsureVisible
    End If
Next i
End Sub

Private Sub BtnImprimir_Click()
Cancelar_Pago

''========= ESTE ES EL CODIGO NUEVO ==========
With CrystalReport1
    .ReportFileName = RutaInformes & "\CuentasPorCobrar.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    
    If TxtCodigo.Text <> "" And TxtNombre.Text = "" And TxtNoFactura.Text = "" And TxtDtpFecha.Text = "" Then
        .SelectionFormula = "{Facturacion.IdCliente} = " & Trim(TxtCodigo.Text)
    ElseIf TxtNombre.Text <> "" And TxtCodigo.Text = "" And TxtNoFactura.Text = "" And TxtDtpFecha.Text = "" Then
        .SelectionFormula = "{Facturacion.Razon} = '" & Trim(TxtNombre.Text) & "'"
    ElseIf TxtNoFactura.Text <> "" And TxtNombre.Text = "" And TxtCodigo.Text = "" And TxtDtpFecha.Text = "" Then
        .SelectionFormula = "{Facturacion.N_Factura} = " & Trim(TxtNoFactura.Text)
    ElseIf TxtDtpFecha.Text <> "" And TxtCodigo.Text = "" And TxtNombre.Text = "" And TxtNoFactura.Text = "" Then
        .SelectionFormula = "{Facturacion.Fecha} = " & FechaSQL(TxtDtpFecha.Text)
'    Else
'        .SelectionFormula = "{Facturacion.Forma_Pago} <> 0 AND {Facturacion.Forma_Pago} <> -1 AND {Facturacion.Anulada} <> 1"
    End If
    
    .WindowTitle = "Reporte Cuentas por Cobrar"
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
            BtnBuscar.SetFocus
        Case vbKeyUp
            TxtNombre.SetFocus
        Case vbKeyLeft
            If FrameAbonos.Enabled = True Then DtpFechaAbono.SetFocus
        Case vbKeyRight
            TxtDtpFecha.SetFocus
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
'IniDMGrid
DtpFechaAbono.Value = Format(Now, "DD/MM/YY")
BtnGuardarActualizar.Enabled = False
CargarCuentasPorCobrar
End Sub

Sub CargarCuentasPorCobrar()
Dim FechaVenc

FechaVenc = DateValue(Now - 30)

CSql = "Select * From C_Cobrar Where (Forma_Pago<>0 or Forma_Pago<> -1 or Forma_Pago<> 1) And Anulada<>'1' order by IdCliente, N_Factura"

Set RsCuentasCobrar = CrearRS(CSql)
LstCuentasCobrar.ListItems.Clear
Do While Not RsCuentasCobrar.EOF

    CSql = "Select * From Cliente Where IdCliente ='" & RsCuentasCobrar.Fields("IdCliente").Value & "'"
    Set RsBuscarCliente = CrearRS(CSql)

    CSql = "Select * From Paciente Where IdPaciente ='" & RsCuentasCobrar.Fields("IdPaciente").Value & "'"
    Set RsBuscarPaciente = CrearRS(CSql)
    With LstCuentasCobrar
        C = C + 1
        .ListItems.Add , , RsBuscarCliente.Fields("Razon").Value
        If RsBuscarPaciente.RecordCount > 0 Then
            .ListItems(C).ListSubItems.Add , , RsBuscarPaciente.Fields("ApellidoP").Value & ", " & RsBuscarPaciente.Fields("NombreP").Value
            .ListItems(C).ListSubItems.Add , , RsBuscarPaciente.Fields("CedulaP").Value
        Else
            .ListItems(C).ListSubItems.Add , , "----"
            .ListItems(C).ListSubItems.Add , , "----"
        End If
        .ListItems(C).ListSubItems.Add , , RsCuentasCobrar.Fields("N_Factura").Value
        .ListItems(C).ListSubItems.Add , , RsCuentasCobrar.Fields("Fecha").Value
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("SubTotal").Value, "#,##0.00")
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("Impuesto").Value, "#,##0.00")
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("Monto").Value, "#,##0.00")
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("Descuentos").Value, "#,##0.00")
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("Retenciones").Value, "#,##0.00")
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("TimbresFiscales").Value, "#,##0.00")
        .ListItems(C).ListSubItems.Add , , Format(RsCuentasCobrar.Fields("PorCobrar").Value, "#,##0.00")

        If Val(RsCuentasCobrar.Fields("Forma_Pago").Value) = 0 Then
            If CDbl(RsCuentasCobrar.Fields("PorCobrar").Value) <= 0 Then
                .ListItems(C).ListSubItems.Add , , "Cancelada"
            Else
                .ListItems(C).ListSubItems.Add , , "Vencida"
                .ListItems(C).ListSubItems.Item(1).ForeColor = vbRed
            End If
        ElseIf Val(RsCuentasCobrar.Fields("Forma_Pago").Value) = 1 Then
            If CDbl(RsCuentasCobrar.Fields("PorCobrar").Value) <= 0 Then
                .ListItems(C).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsCuentasCobrar.Fields("Fecha").Value) - DateValue(Now)) = -30 Then
                    .ListItems(C).ListSubItems.Add , , "Vence hoy"
                    .ListItems(C).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    If (DateValue(RsCuentasCobrar.Fields("Fecha").Value) - DateValue(Now)) < -30 Then
                        .ListItems(C).ListSubItems.Add , , "Vencida"
                        .ListItems(C).ListSubItems.Item(1).ForeColor = vbRed
                    Else
                        .ListItems(C).ListSubItems.Add , , "Vigente"
                    End If
                End If
            End If
        Else
            If CDbl(RsCuentasCobrar.Fields("PorCobrar").Value) <= 0 Then
                .ListItems(C).ListSubItems.Add , , "Cancelada"
            Else
                If (DateValue(RsCuentasCobrar.Fields("Fecha").Value) - DateValue(Now)) = -45 Then
                    .ListItems(C).ListSubItems.Add , , "Vence hoy"
                    .ListItems(C).ListSubItems.Item(1).ForeColor = vbBlue
                Else
                    If (DateValue(RsCuentasCobrar.Fields("Fecha").Value) - DateValue(Now)) < -45 Then
                        .ListItems(C).ListSubItems.Add , , "Vencida"
                        .ListItems(C).ListSubItems.Item(1).ForeColor = vbRed
                    Else
                        .ListItems(C).ListSubItems.Add , , "Vigente"
                    End If
                End If
            End If
        End If
    End With
    DoEvents
    RsCuentasCobrar.MoveNext
Loop


Dim RsTotales As New ADODB.Recordset
CSql = "Select Sum(SubTotal) as TSubTotal, sum(Impuesto) as TImpuesto, sum(Monto) as TMonto, sum(PorCobrar) as TPorCobrar  From C_Cobrar Where (Forma_Pago<>0 or Forma_Pago<> -1 or Forma_Pago<>1) And Anulada<>'1'"
'CSql = "Select Sum(SubTotal) as TSubTotal, sum(Impuesto) as TImpuesto, Sum(Monto) as TMonto, sum(PorCobrar) as TPorCobrar  From C_Cobrar"
Set RsTotales = CrearRS(CSql)

If RsTotales.RecordCount > 0 Then
    LblSubTotal.Caption = Format(RsTotales.Fields("TSubTotal").Value, "#,##0.00")
    LblImpuestos.Caption = Format(RsTotales.Fields("TImpuesto").Value, "#,##0.00")
    LblTotalGeneral.Caption = Format(RsTotales.Fields("TMonto").Value, "#,##0.00")
    LblPorCobrar.Caption = Format(RsTotales.Fields("TPorCobrar").Value, "#,##0.00")
End If
End Sub

Private Sub LstAbonosCuentasCobrar_DblClick()
Dim RsUpdate As New ADODB.Recordset
Dim MAbono, FAbono As String
Dim Seleccion As Integer

    If LstAbonosCuentasCobrar.ListItems.Count = 0 Then Exit Sub
    
    MAbono = LstAbonosCuentasCobrar.SelectedItem.ListSubItems(1).Text
    FAbono = LstAbonosCuentasCobrar.SelectedItem.Text
    Seleccion = MsgBox("Se procedera a eliminar el Abono seleccionado de fecha " & FAbono & " sobre la Factura No-" & NoFact & "", vbQuestion + vbYesNo, "Confimar Operacion")
    
    If Seleccion = vbNo Then Exit Sub
    
    Seleccion = LstCuentasCobrar.SelectedItem.ListSubItems(3).Text

    ' Actualiza datos de la tabla AbonoCtaXCobrar
    CSql = "Select * From AbonoCtaXCobrar Where NoFactura='" & NoFact & "' AND FechaEmision='" & FechaEmi & "' AND " _
            & " MontoAbono=" & Replace(CDbl(MAbono), ",", ".") & " AND FechaAbono ='" & FAbono & "'"
    Set RsUpdate = CrearRS(CSql)

    RsUpdate.Delete
    'DELETE FROM AbonoCtaXCobrar
    PorCobrar = CDbl(PorCobrar) + CDbl(MAbono)
    RsUpdate.Update


    Dim RsUpdatePorCobrar As New ADODB.Recordset
    CSql = "Select * From C_Cobrar where N_Factura='" & NoFact & "'"
    Set RsUpdatePorCobrar = CrearRS(CSql)

    RsUpdatePorCobrar.Fields("PorCobrar").Value = PorCobrar
    RsUpdatePorCobrar.Update

    MsgBox "Los Cambios han sido registrados!", vbInformation + vbOKOnly, "Operacion Exitosa!"

    CargarAbonos
    CargarCuentasPorCobrar
    Cancelar_Pago
    
    For i = 1 To LstCuentasCobrar.ListItems.Count
    If LstCuentasCobrar.ListItems(i).ListSubItems(3).Text = Seleccion Then
        LstCuentasCobrar.ListItems(i).Selected = True
        LstCuentasCobrar.ListItems(i).EnsureVisible
    End If
Next i

End Sub

Private Sub LstCuentasCobrar_Click()
If LstCuentasCobrar.ListItems.Count = 0 Then Exit Sub
' Inicializa los objectos
Cancelar_Pago

' Crea la Consulta
CSql = "Select * From C_Cobrar where N_Factura='" & LstCuentasCobrar.SelectedItem.ListSubItems(3).Text & "'"
Set RsSeleccionarCuentasPorCobrar = CrearRS(CSql)

If RsSeleccionarCuentasPorCobrar.RecordCount > 0 Then
    NoFact = RsSeleccionarCuentasPorCobrar.Fields("N_Factura").Value
    FormaPago = RsSeleccionarCuentasPorCobrar.Fields("Forma_Pago").Value
    Tipo = RsSeleccionarCuentasPorCobrar.Fields("Tipo").Value
    IdPaciente = RsSeleccionarCuentasPorCobrar.Fields("Idpaciente").Value
    IdCliente = Str(RsSeleccionarCuentasPorCobrar.Fields("IdCliente").Value)
    FechaEmi = RsSeleccionarCuentasPorCobrar.Fields("Fecha").Value
    IdUsuario = RsSeleccionarCuentasPorCobrar.Fields("Idusuario").Value
    impresa = RsSeleccionarCuentasPorCobrar.Fields("impresa").Value
    SubTotal = RsSeleccionarCuentasPorCobrar.Fields("SubTotal").Value
    TasaImpuesto = RsSeleccionarCuentasPorCobrar.Fields("TasaImpuesto").Value
    Impuesto = RsSeleccionarCuentasPorCobrar.Fields("Impuesto").Value
    Monto = RsSeleccionarCuentasPorCobrar.Fields("Monto").Value
    PorCobrar = RsSeleccionarCuentasPorCobrar.Fields("PorCobrar").Value
    Anulada = RsSeleccionarCuentasPorCobrar.Fields("Anulada").Value
End If
RsSeleccionarCuentasPorCobrar.Close

CargarAbonos
End Sub

Sub CargarAbonos()
Dim RsCargarAbonos As New ADODB.Recordset

CSql = "Select * From AbonoCtaXCobrar Where NoFactura='" & NoFact & "'"
Set RsCargarAbonos = CrearRS(CSql)

LstAbonosCuentasCobrar.ListItems.Clear

Do While Not RsCargarAbonos.EOF
    With LstAbonosCuentasCobrar
        i = i + 1
        .ListItems.Add , , RsCargarAbonos.Fields("FechaAbono").Value
        .ListItems(i).ListSubItems.Add , , Format(RsCargarAbonos.Fields("MontoAbono").Value, "#,##0.00")
    End With
    RsCargarAbonos.MoveNext
Loop
End Sub

Sub Cancelar_Pago()
    BtnAgregar.Enabled = True
    FrameAbonos.Enabled = False
    BtnGuardarActualizar.Enabled = False
    TxtAbono.Text = "0.00"
End Sub

Private Sub LstCuentasCobrar_DblClick()
On Error Resume Next
If LstCuentasCobrar.ListItems.Count = 0 Then Exit Sub
CSql = "Select * From C_Cobrar where N_Factura='" & LstCuentasCobrar.SelectedItem.ListSubItems(3).Text & "'"
Set RsSeleccionarCuentasPorCobrar = CrearRS(CSql)

If RsSeleccionarCuentasPorCobrar.RecordCount > 0 Then
    If RsSeleccionarCuentasPorCobrar.Fields("impresa").Value Then
        MsgBox "La factura no puede ser editada, si la misma contiene errores proceda a anularla!", vbExclamation + vbOKOnly, "Factura Impresa!"
        Exit Sub
    End If
End If
RsSeleccionarCuentasPorCobrar.Close

    BtnEditorFacturas_Click
End Sub

Private Sub LstCuentasCobrar_KeyUp(KeyCode As Integer, Shift As Integer)
LstCuentasCobrar_Click
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
Cancelar_Pago
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnBuscar.SetFocus
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

Cancelar_Pago
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtNoFactura_KeyUp(KeyCode As Integer, Shift As Integer)
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

Cancelar_Pago
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtNombre_KeyUp(KeyCode As Integer, Shift As Integer)
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


'Sub IniDMGrid()
'' carga las columnas y encabezados de columna
'DMGrid1.Cols = 9
'DMGrid1.Rows = 0
'DMGrid1.DColumnas(1).Alignment = 0
'DMGrid1.DColumnas(2).Alignment = 0
'DMGrid1.DColumnas(3).Alignment = 0
'DMGrid1.DColumnas(4).Alignment = 0
'DMGrid1.DColumnas(5).Alignment = 0
'DMGrid1.DColumnas(6).Alignment = 0
'DMGrid1.DColumnas(7).Alignment = 0
'DMGrid1.DColumnas(8).Alignment = 0
'DMGrid1.DColumnas(9).Alignment = 0
'DMGrid1.DColumnas(1).Locked = True
'DMGrid1.DColumnas(2).Locked = True
'DMGrid1.DColumnas(3).Locked = True
'DMGrid1.DColumnas(4).Locked = True
'DMGrid1.DColumnas(5).Locked = True
'DMGrid1.DColumnas(6).Locked = True
'DMGrid1.DColumnas(7).Locked = True
'DMGrid1.DColumnas(8).Locked = True
'DMGrid1.DColumnas(9).Locked = True
'DMGrid1.DColumnas(1).Width = 3500
'DMGrid1.DColumnas(2).Width = 3500
'DMGrid1.DColumnas(3).Width = 1500
'DMGrid1.DColumnas(4).Width = 1500
'DMGrid1.DColumnas(5).Width = 1500
'DMGrid1.DColumnas(6).Width = 2000
'DMGrid1.DColumnas(7).Width = 2000
'DMGrid1.DColumnas(8).Width = 2000
'DMGrid1.DColumnas(9).Width = 2000
'DMGrid1.DColumnas(1).Caption = "Cliente"
'DMGrid1.DColumnas(2).Caption = "Paciente"
'DMGrid1.DColumnas(3).Caption = "No. Cedula"
'DMGrid1.DColumnas(4).Caption = "No. Factura"
'DMGrid1.DColumnas(5).Caption = "Fecha"
'DMGrid1.DColumnas(6).Caption = "SubTotal"
'DMGrid1.DColumnas(7).Caption = "Impuesto"
'DMGrid1.DColumnas(8).Caption = "Monto"
'DMGrid1.DColumnas(9).Caption = "PorCobrar"
'End Sub

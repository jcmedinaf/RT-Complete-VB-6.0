VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPresupuestoProducto 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto Producto"
   ClientHeight    =   7440
   ClientLeft      =   4950
   ClientTop       =   1875
   ClientWidth     =   10650
   Icon            =   "Presupuestos.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10650
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   1215
         Left            =   3360
         TabIndex        =   35
         Top             =   5880
         Width           =   6975
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   5880
            TabIndex        =   36
            ToolTipText     =   "Cerrar"
            Top             =   360
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
            MICON           =   "Presupuestos.frx":1002
            PICN            =   "Presupuestos.frx":101E
            PICH            =   "Presupuestos.frx":11E7
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
            TabIndex        =   37
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   360
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
            MICON           =   "Presupuestos.frx":141C
            PICN            =   "Presupuestos.frx":1438
            PICH            =   "Presupuestos.frx":16C7
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
            TabIndex        =   38
            ToolTipText     =   "Agregar"
            Top             =   360
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
            MICON           =   "Presupuestos.frx":1B08
            PICN            =   "Presupuestos.frx":1B24
            PICH            =   "Presupuestos.frx":1CB1
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
            Left            =   4680
            TabIndex        =   39
            ToolTipText     =   "Deshacer Operacion"
            Top             =   360
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
            MICON           =   "Presupuestos.frx":1EE6
            PICN            =   "Presupuestos.frx":1F02
            PICH            =   "Presupuestos.frx":21E4
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
            TabIndex        =   40
            ToolTipText     =   "Eliminar"
            Top             =   360
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
            MICON           =   "Presupuestos.frx":2435
            PICN            =   "Presupuestos.frx":2451
            PICH            =   "Presupuestos.frx":25F5
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
            Height          =   495
            Left            =   3600
            TabIndex        =   41
            ToolTipText     =   "Reporte"
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "Presupuestos.frx":2794
            PICN            =   "Presupuestos.frx":27B0
            PICH            =   "Presupuestos.frx":28D5
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
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   5880
         Width           =   3135
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Text            =   "Busqueda"
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Ce&dula"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nº Presup"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   32
            Top             =   480
            Width           =   1095
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Presupuestos.frx":2B65
            PICN            =   "Presupuestos.frx":2B81
            PICH            =   "Presupuestos.frx":2DE6
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
         Caption         =   "Datos del Proveedor"
         Height          =   1815
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   10215
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   9720
            Top             =   120
         End
         Begin VB.TextBox Text8 
            Height          =   315
            Left            =   1200
            TabIndex        =   27
            Top             =   360
            Width           =   6135
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   360
            Width           =   6135
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1200
            TabIndex        =   25
            Top             =   720
            Width           =   8535
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   1200
            TabIndex        =   24
            Top             =   1200
            Width           =   1935
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   9240
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowCloseBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarProveedor 
            Height          =   495
            Left            =   3360
            TabIndex        =   44
            ToolTipText     =   "Agregar Pacientes"
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Agregar Proveedores"
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
            MICON           =   "Presupuestos.frx":3078
            PICN            =   "Presupuestos.frx":3094
            PICH            =   "Presupuestos.frx":3221
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
            BackStyle       =   0  'Transparent
            Caption         =   "RIF.:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   990
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   10215
         Begin VB.TextBox Text7 
            Height          =   405
            Left            =   6600
            TabIndex        =   10
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   720
            TabIndex        =   9
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   9615
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1200
            TabIndex        =   7
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   3960
            TabIndex        =   5
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   7320
            TabIndex        =   4
            Top             =   2520
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Si / No"
            Height          =   375
            Left            =   3960
            TabIndex        =   3
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   720
            TabIndex        =   2
            Top             =   3000
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6600
            TabIndex        =   11
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   211353601
            CurrentDate     =   39834
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarPacientes 
            Height          =   735
            Left            =   8760
            TabIndex        =   42
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1296
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
            MICON           =   "Presupuestos.frx":3456
            PICN            =   "Presupuestos.frx":3472
            PICH            =   "Presupuestos.frx":36D7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Presupuesto:"
            Height          =   195
            Left            =   6000
            TabIndex        =   45
            Top             =   3090
            Width           =   1230
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   2610
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion del Producto:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   1830
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   6000
            TabIndex        =   20
            Top             =   930
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   6000
            TabIndex        =   19
            Top             =   465
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   450
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costo:"
            Height          =   195
            Left            =   3360
            TabIndex        =   16
            Top             =   2610
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación:"
            Height          =   195
            Left            =   6480
            TabIndex        =   15
            Top             =   2610
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Iva"
            Height          =   195
            Left            =   3360
            TabIndex        =   14
            Top             =   3090
            Width           =   225
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   3090
            Width           =   360
         End
         Begin VB.Label Label11 
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
            Left            =   7320
            TabIndex        =   12
            Top             =   3000
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "FrmPresupuestoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sBuscar As String
Dim rs As New ADODB.Connection
Dim BD As New ADODB.Recordset  'Tabla Registro Historico
Dim SQL As String
Dim BD70 As New ADODB.Recordset 'Tabla paciente
Dim BD71 As New ADODB.Recordset 'Tabla Informe medico
Dim BD72 As New ADODB.Recordset 'Tabla presupuestos
Dim CHAN
Dim IDPRO

Sub Blanqueo()

Text8.Visible = False
Combo1.Visible = True
Call ListaCliente
Command5.Enabled = True
Command1.Enabled = False
    
IdPac1 = ""
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text13.Text = ""
 
CHAN = 0

End Sub

Private Sub BtnAgregar_Click()
'Command3
Call Blanqueo
End Sub

Private Sub BtnAgregarProveedor_Click()
'command6
FrmProveedores.Show
Call ListaCliente
End Sub

Private Sub BtnBuscar_Click()
'command8
If Trim(TxtBuscar.Text) = "" Then Exit Sub
If Option1(0).Value = True Then Cbus = "cedulaP"
If Option1(1).Value = True Then Cbus = "npresupuesto"
SQL = "select * from presupuesto2 where " & Cbus & " like " & TxtBuscar.Text
Set BD72 = CrearRS(SQL)

If Not (BD72.EOF) Then
    If Trim(BD72.Fields("Descripcion")) <> "" Then Text4.Text = BD72.Fields("descripcion") Else Text4.Text = ""
    If BD72.Fields("precio") <> "" Then Text5.Text = BD72.Fields("precio") Else Text5.Text = ""
    If BD72.Fields("costoActual") <> "" Then Text6.Text = BD72.Fields("costoActual") Else Text6.Text = ""
    If BD72.Fields("ubicacion") <> "" Then Text10.Text = BD72.Fields("ubicacion") Else Text10.Text = ""
    If BD72.Fields("TipoServicio") <> "" Then Text11.Text = BD72.Fields("TipoServicio") Else Text11.Text = ""
    
    If BD72.Fields("nombreP") <> "" Then Text1.Text = BD72.Fields("nombreP") Else Text1.Text = ""
    If BD72.Fields("cedulaP") <> "" Then Text7.Text = BD72.Fields("cedulaP") Else Text7.Text = ""
    If BD72.Fields("apellidoP") <> "" Then Text2.Text = BD72.Fields("apellidoP") Else Text2.Text = ""
    If BD72.Fields("idpaciente") <> "" Then IdPac1 = BD72.Fields("idpaciente")
    If BD72.Fields("direccionC") <> "" Then Text9.Text = BD72.Fields("direccionC") Else Text9.Text = ""
    If BD72.Fields("Rif_Proveedor") <> "" Then Text13.Text = BD72.Fields("Rif_Proveedor") Else Text13.Text = ""
    
    Text8.Visible = True
    Combo1.Visible = False
    If BD72.Fields("Nombrep") <> "" Then Text8.Text = BD72.Fields("Nombrep") Else Text8.Text = ""
    Call presut
    
Else
Msg = "No Existe presupuesto coincidente con la clave de busqueda solicitada"
MsgBox Msg, vbOKOnly, "No Existe registro alguno"

End If
BD72.Close
'Command5.Enabled = False

'CrystalReport1.SelectionFormula = "PRESUPUESTO1.IDPACIENTE = " & IdPac1

End Sub

Private Sub BtnBuscarPacientes_Click()
'command5
Msg = "Indique la cedula del paciente que va a generarle el presupuesto"
ced = Trim(InputBox(Msg, "Cedula del paciente", "12345678"))
If ced = "" Then Exit Sub

CSql1 = "select * from paciente where cedula = " & ced
BD70.Open CSql1, Cnn
If Not (BD70.EOF) Then
    Text1.Text = BD70.Fields("nombreP")
    Text7.Text = BD70.Fields("cedulaP")
    Text2.Text = BD70.Fields("apellidoP")
    IdPac1 = BD70.Fields("idpaciente")
    CSql1 = "select descripcion,precio,costo,ubicacion,tipo_ser from Productos where idproducto = 1"
    BD71.Open CSql1, Cnn

    If Not (BD71.EOF) Then
       Text4.Text = BD71.Fields("descripcion")
       Text5.Text = BD71.Fields("precio")
       Text6.Text = BD71.Fields("costo")
       Text10.Text = BD71.Fields("ubicacion")
       Text11.Text = BD71.Fields("tipo_ser")
       
      BtnGuardarActualizar.Enabled = False
        CHAN = 1
        
    Else
    Msg = "Este paciente no presenta registro en la tabla de Informes Medicos"
    MsgBox Msg, vbOKOnly, "No tiene informe medico"
    Command1.Enabled = False
    CHAN = 1
    
    End If
    
BD71.Close
Else
Msg = "No existe esa cedula en la tabla pacientes" & Chr(13) & Chr(13) & "   " & ced
MsgBox Msg, vbOKOnly, "No existe paciente"
BtnGuardarActualizar.Enabled = False
CHAN = 0
End If
BD70.Close
End Sub


Private Sub BtnCerrar_Click()
'command7
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
'command1
CSql = "Insert into Productos (IDpaciente,descripcion,precio,costo,ubicacion,tipo_ser,fecha_producto,idproveedor,idproducto,idusuario) VALUES(" & IdPac1 & "," & Text4.Text & "," & Text5.Text & "," & Text6.Text & "," & Text10.Text & "," & Text11.Text & "," & Text6.Text & ",#" & Format(Date, "mm/dd/yyyy") & "#," & IDPRO & ", 1," & IdUser & ")"
Dim BD As New ADODB.Recordset
BD.Open CSql, Cnn
Msg = "Registro Agregado satisfactoriamente"
MsgBox Msg, vbOKOnly

'Call blanqueo
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly, "Error al Guardar"
    
    Exit Sub
End Sub

Private Sub BtnImprimir_Click()

'Dim RsReporte As New ADODB.Recordset
'CSql = "Select Fecha , NPresupuesto, Razon, Rif, DireccionC, Email, ApellidoP, NombreP, CedulaP, DireccionP, Codigo, Telefono, Codigoc, Celular, Ocupacion, E-mail, Duracion, Diagnotico, Monto From Presupuesto2 Where Rif_Proveedor='" & Text13.Text & "' And CedulaP='" & Text7.Text & "'"
'Set RsReporte = CrearRS(CSql)
'
'If RsReporte.RecordCount > 0 Then
'    CreateFieldDefFile RsReporte, Trim(RutaInformes) & " \ " & "Presupuesto2.ttx", 1
'    FrmVistaPrevia.Show
'End If

End Sub

Private Sub Combo1_Click()

            CSql = "SELECT * FROM Proveedor where idproveedor = " & Combo1.ItemData(Combo1.ListIndex)
            BD.CursorType = adOpenKeyset
            BD.LockType = adLockOptimistic
            BD.CursorLocation = adUseClient
            BD.Open CSql, Cnn, , , adCmdText
                If Not (BD.EOF) Then
            Text9.Text = BD.Fields("direccion")
            Text13.Text = BD.Fields("Rif_Proveedor")
                IDPRO = BD.Fields("idproveedor")
            End If
                BD.Close

End Sub

Sub ListaCliente()
Combo1.Clear
                
CSql = "SELECT * FROM Proveedor"
BD.CursorType = adOpenKeyset
BD.LockType = adLockOptimistic
BD.CursorLocation = adUseClient
BD.Open CSql, Cnn, , , adCmdText

If Not (BD.EOF) Then BD.MoveFirst
Do While Not BD.EOF
clent = BD.Fields("Nombre")
Combo1.AddItem UCase(clent)
Combo1.ItemData(Combo1.NewIndex) = BD.Fields("idproveedor")
BD.MoveNext
Loop
BD.Close
            
End Sub

Sub Prod()
CSql = "select descripcion,precio,costo,ubicacion,tipo_ser from Productos where idpaciente = " & IdPac1
BD71.Open CSql, Cnn

If Not (BD71.EOF) Then
    Text4.Text = BD71.Fields("descripcion")
    Text5.Text = BD71.Fields("precio")
    Text6.Text = BD71.Fields("costo")
    Text10.Text = BD71.Fields("ubicacion")
    Text11.Text = BD71.Fields("tipo_ser")
End If
       
End Sub

Private Sub Form_Load()
Centrar Me

'CrystalReport1.ReportFileName = Direc & "\Informes\Presupuesto1.rpt"
Text8.Visible = False
Combo1.Visible = True
Call ListaCliente
 
End Sub

Sub presut()

CSql = "select Npresupuesto from Presupuesto where idpaciente = " & IdPac1
Dim BD76 As New ADODB.Recordset
BD76.Open CSql, Cnn

Pres = Format(BD76.Fields("Npresupuesto"), "000000000#")

BD76.Close
Label11.Caption = Pres
      
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub


Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub


Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57 ' permite el ingreso de numeros
Case Is = 13 ' permite presionar el ENTER
Call BtnBuscar_Click

Case Is = 8 ' Permite Borrar de retroceso

Case Else ' Inhibe todas las demas teclas
If Option1(0).Value = True Then KeyAscii = 0 Else Exit Sub
End Select
End Sub

Private Sub Text4_Change()
KeyAscii = 0
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
d = IDPRO
FrmListadoProductosServicios.Show 1
If d <> IDPRO Then
Call Prod
End If
End If
End Sub
Private Sub Text5_Change()
If CHAN = 1 Then Command1.Enabled = True

End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
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
Private Sub Text6_Change()
If CHAN = 1 Then Command1.Enabled = True

End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

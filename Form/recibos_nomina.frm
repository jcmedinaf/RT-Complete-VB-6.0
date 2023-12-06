VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReciboPagos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos de Pago"
   ClientHeight    =   6810
   ClientLeft      =   4335
   ClientTop       =   3315
   ClientWidth     =   11250
   Icon            =   "recibos_nomina.frx":0000
   LinkTopic       =   "Form45"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11250
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   11055
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   9960
         TabIndex        =   6
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
         MICON           =   "recibos_nomina.frx":1002
         PICN            =   "recibos_nomina.frx":101E
         PICH            =   "recibos_nomina.frx":11E7
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
         TabIndex        =   7
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
         MICON           =   "recibos_nomina.frx":141C
         PICN            =   "recibos_nomina.frx":1438
         PICH            =   "recibos_nomina.frx":16C7
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
         TabIndex        =   8
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
         MICON           =   "recibos_nomina.frx":1B08
         PICN            =   "recibos_nomina.frx":1B24
         PICH            =   "recibos_nomina.frx":1CB1
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
         Left            =   8760
         TabIndex        =   9
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
         MICON           =   "recibos_nomina.frx":1EE6
         PICN            =   "recibos_nomina.frx":1F02
         PICH            =   "recibos_nomina.frx":21E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrar 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
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
         MICON           =   "recibos_nomina.frx":2435
         PICN            =   "recibos_nomina.frx":2451
         PICH            =   "recibos_nomina.frx":25F5
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
         Left            =   7080
         TabIndex        =   11
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
         MICON           =   "recibos_nomina.frx":2794
         PICN            =   "recibos_nomina.frx":27B0
         PICH            =   "recibos_nomina.frx":2A46
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
         Left            =   6480
         TabIndex        =   12
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
         MICON           =   "recibos_nomina.frx":2CA5
         PICN            =   "recibos_nomina.frx":2CC1
         PICH            =   "recibos_nomina.frx":2F56
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
         Left            =   4440
         TabIndex        =   13
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
         MICON           =   "recibos_nomina.frx":31B2
         PICN            =   "recibos_nomina.frx":31CE
         PICH            =   "recibos_nomina.frx":32F3
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
      Caption         =   "Detalles"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   11055
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4260
         Object.Width           =   10785
         Object.Height          =   2385
         ScrollBar       =   1
         Editable        =   -1  'True
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregarConceptos 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Agregar Conceptos de Nomina"
         Top             =   2880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Agregar Conceptos"
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
         MICON           =   "recibos_nomina.frx":3583
         PICN            =   "recibos_nomina.frx":359F
         PICH            =   "recibos_nomina.frx":383B
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
         Caption         =   "Otros:"
         Height          =   195
         Left            =   8040
         TabIndex        =   43
         Top             =   3120
         Width           =   420
      End
      Begin VB.Label LblTotalOtros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   300
         Left            =   9360
         TabIndex        =   42
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label LblTotalAsignacion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   300
         Left            =   6120
         TabIndex        =   38
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Asignación:"
         Height          =   195
         Left            =   4680
         TabIndex        =   37
         Top             =   2820
         Width           =   1230
      End
      Begin VB.Label LblTotalDeduccion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   300
         Left            =   6120
         TabIndex        =   36
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deducciones:"
         Height          =   195
         Left            =   4680
         TabIndex        =   35
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label LblNetoCancelar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   300
         Left            =   9360
         TabIndex        =   34
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neto a Cancelar:"
         Height          =   195
         Left            =   8040
         TabIndex        =   14
         Top             =   2820
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Trabajador"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53542915
         CurrentDate     =   40141
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4800
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         Connect         =   "Data Source=Ing03;"
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
         WindowShowNavigationCtls=   -1  'True
         WindowShowCancelBtn=   -1  'True
         WindowShowPrintBtn=   -1  'True
         WindowShowExportBtn=   -1  'True
         WindowShowZoomCtl=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowProgressCtls=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label LblNroPeriodo 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7320
         TabIndex        =   44
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   4800
         TabIndex        =   40
         Top             =   1860
         Width           =   510
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   8880
         Picture         =   "recibos_nomina.frx":3ADA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recibo No."
         Height          =   195
         Left            =   6240
         TabIndex        =   33
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label LblNoRecibo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7320
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "al"
         Height          =   195
         Left            =   7080
         TabIndex        =   31
         Top             =   1860
         Width           =   120
      End
      Begin VB.Label LblAl 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7320
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de la Nómina:"
         Height          =   195
         Left            =   5520
         TabIndex        =   29
         Top             =   1500
         Width           =   1560
      End
      Begin VB.Label LblDel 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5520
         TabIndex        =   28
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Nómina:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label LblTipoNomina 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   26
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puesto:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1493
         Width           =   540
      End
      Begin VB.Label LblCargo 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   24
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1140
         Width           =   1050
      End
      Begin VB.Label LblDepartamento 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   22
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4680
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Empleado:"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label LblCodigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3720
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblNombre 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5520
         TabIndex        =   17
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label LblApellido 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   16
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label LblCedula 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cédula:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   428
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido(s):"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         Height          =   195
         Left            =   4680
         TabIndex        =   2
         Top             =   780
         Width           =   765
      End
   End
End
Attribute VB_Name = "FrmReciboPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsEmpleados As New ADODB.Recordset
Dim RsDatAdmin As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim bdata As New ADODB.Recordset
Dim IdReci
Dim TamDMGrid As Integer
Dim VarTAsig As Double
Dim VarTDedc As Double
Dim VarTOtro As Double
Dim VarTSald As Double
Dim VarTCamp As String
Dim CodDMGrid1 As Integer
Dim NomDMGrid1 As String

Public IdEmpla As Integer
Public FechaTemp As String
Public NTabla As Integer


Sub ReCalcular()
TamDMGrid = Val(DMGrid1.Rows)
VarTAsig = 0
VarTDedc = 0
VarTOtro = 0
VarTSald = 0
For i = 1 To TamDMGrid
    If Not IsNull(DMGrid1.ValorCelda(i, 4)) Then
        If Trim(DMGrid1.ValorCelda(i, 4)) <> "" Then VarTAsig = VarTAsig + CDbl(DMGrid1.ValorCelda(i, 4))
    End If
    If Not IsNull(DMGrid1.ValorCelda(i, 5)) Then
        If Trim(DMGrid1.ValorCelda(i, 5)) <> "" Then VarTDedc = VarTDedc + CDbl(DMGrid1.ValorCelda(i, 5))
    End If
    If Not IsNull(DMGrid1.ValorCelda(i, 6)) Then
        If Trim(DMGrid1.ValorCelda(i, 6)) <> "" Then VarTOtro = VarTOtro + CDbl(DMGrid1.ValorCelda(i, 6))
    End If
    If Not IsNull(DMGrid1.ValorCelda(i, 7)) Then
        If Trim(DMGrid1.ValorCelda(i, 7)) <> "" Then VarTSald = VarTSald + CDbl(DMGrid1.ValorCelda(i, 7))
    End If
Next i

LblTotalAsignacion.Caption = Format(VarTAsig, "#,##0.00")
LblTotalDeduccion.Caption = Format(VarTDedc, "#,##0.00")
LblTotalOtros.Caption = Format(VarTOtro, "#,##0.00")
LblNetoCancelar.Caption = Format((VarTAsig - VarTDedc), "#,##0.00")

End Sub

Sub Grid2()
    'carga las columnas y encabezados de columna
    DMGrid1.Cols = 7
    DMGrid1.DColumnas(3).Alignment = 1
    DMGrid1.DColumnas(4).Alignment = 1
    DMGrid1.DColumnas(5).Alignment = 1
    DMGrid1.DColumnas(6).Alignment = 1
    DMGrid1.DColumnas(7).Alignment = 1

    DMGrid1.DColumnas(3).IsNumber = True
    DMGrid1.DColumnas(4).IsNumber = True
    DMGrid1.DColumnas(5).IsNumber = True
    DMGrid1.DColumnas(6).IsNumber = True
    DMGrid1.DColumnas(7).IsNumber = True
    
    DMGrid1.DColumnas(1).Locked = True
    DMGrid1.DColumnas(2).Locked = True
    DMGrid1.DColumnas(3).Locked = True
    DMGrid1.DColumnas(4).Locked = False
    DMGrid1.DColumnas(5).Locked = False
    DMGrid1.DColumnas(6).Locked = False
    DMGrid1.DColumnas(7).Locked = True

    DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
    DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 30 / 100) - 300
    DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 10 / 100)
    DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
    DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)
    DMGrid1.DColumnas(6).Width = Val(DMGrid1.Width * 10 / 100)
    DMGrid1.DColumnas(7).Width = Val(DMGrid1.Width * 10 / 100)
    'DMGrid1.DColumnas(6).Width = 1700

    DMGrid1.DColumnas(1).Caption = "Código"
    DMGrid1.DColumnas(2).Caption = "Concepto"
    DMGrid1.DColumnas(3).Caption = "Cantidad"
    DMGrid1.DColumnas(4).Caption = "Asignaciones"
    DMGrid1.DColumnas(5).Caption = "Deducciones"
    DMGrid1.DColumnas(6).Caption = "Otros"
    DMGrid1.DColumnas(7).Caption = "Saldo"

  End Sub
Sub Carga_Renglones()
Dim TAsignaciones As Double
Dim TDeducciones As Double
Dim TOtros As Double


If IdEmpla <> 0 And NTabla = 0 Then
    Carga_Renglones_Historico
    IdEmpla = 0
    FechaTemp = 0
    Exit Sub
End If

DMGrid1.Clear
DMGrid1.Rows = 0
DMGrid1.PaintMGrid
CSql = " SELECT Reng_Recibo.*, Concepto.Descripcion, Concepto.Tipo,Concepto.IdConcepto, " & _
        " Recibos.Fecha_Ini_Nom , Recibos.Fecha_Fin_Nom, Recibos.Periodo FROM Reng_Recibo INNER JOIN Recibos ON " & _
        " Reng_Recibo.idrecibos = Recibos.IdRecibos AND Reng_Recibo.fecha_gen = Recibos.Fecha_Ini_Nom " & _
        " INNER JOIN Concepto ON Reng_Recibo.idconcepto = Concepto.IdConcepto Where (Recibos.IdEmpleado = " & _
        IdEmpl & ") AND (Reng_Recibo.fecha_gen = '" & Format(DTPicker1, "DD/MM/YYYY") & "') ORDER BY Concepto.IdConcepto"

Set bdata = CrearRS(CSql)
If Not (bdata.EOF) Then
    bdata.MoveFirst
    LblNoRecibo.Caption = Format(bdata.Fields("IdRecibos").Value, "00000")
    LblDel = bdata.Fields("Fecha_Ini_Nom").Value
    LblAl = bdata.Fields("Fecha_Fin_Nom").Value
    LblNroPeriodo.Caption = bdata.Fields("Periodo").Value
    i = 1
    Do While Not bdata.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(i, 1) = Format(bdata.Fields("IdConcepto"), "00000")
        IdReci = bdata.Fields("idrecibos")
        DMGrid1.ValorCelda(i, 2) = bdata.Fields("Detalle")
        
        If CDbl(bdata.Fields("Cantidad")) = 0 Then
            DMGrid1.ValorCelda(i, 3) = ""
        Else
            DMGrid1.ValorCelda(i, 3) = CDbl(bdata.Fields("Cantidad"))
        End If
        
        If Val(bdata.Fields("Tipo")) = 0 Then    ' 0 = asignaciones
            DMGrid1.ValorCelda(i, 4) = CDbl(bdata.Fields("valorn"))
            DMGrid1.ValorCelda(i, 5) = ""
            DMGrid1.ValorCelda(i, 6) = ""
            DMGrid1.ValorCelda(i, 7) = ""
            TAsignaciones = TAsignaciones + CDbl(bdata.Fields("valorn"))
        ElseIf Val(bdata.Fields("Tipo")) = 1 Or Val(bdata.Fields("Tipo")) = 2 Then    ' 1 = deducciones     2 = retenciones
            DMGrid1.ValorCelda(i, 4) = ""
            DMGrid1.ValorCelda(i, 5) = CDbl(bdata.Fields("valorn"))
            DMGrid1.ValorCelda(i, 6) = ""
            DMGrid1.ValorCelda(i, 7) = ""
            TDeducciones = TDeducciones + CDbl(bdata.Fields("valorn"))
        Else
            DMGrid1.ValorCelda(i, 4) = ""
            DMGrid1.ValorCelda(i, 5) = ""
            DMGrid1.ValorCelda(i, 6) = CDbl(bdata.Fields("valorn"))
            DMGrid1.ValorCelda(i, 7) = ""
            TOtros = TOtros + CDbl(bdata.Fields("valorn"))
        End If
            DMGrid1.PaintMGrid
            bdata.MoveNext
        i = i + 1
    Loop
Else
    DMGrid1.Rows = 0
    DMGrid1.Clear
    DMGrid1.PaintMGrid
    LblNoRecibo = "00000"
End If

LblTotalAsignacion.Caption = Format(TAsignaciones, "#,##0.00")
LblTotalDeduccion.Caption = Format(TDeducciones, "#,##0.00")
LblTotalOtros.Caption = Format(TOtros, "#,##0.00")
LblNetoCancelar.Caption = Format((CDbl(LblTotalAsignacion.Caption) - CDbl(LblTotalDeduccion.Caption)), "#,##0.00")
bdata.Close
End Sub

Sub Carga_Renglones_Historico()
Dim TAsignaciones As Double
Dim TDeducciones As Double
Dim TOtros As Double

DMGrid1.Rows = 0
DMGrid1.Clear
DMGrid1.PaintMGrid
CSql = " SELECT Historico_Reng_Nomina.*, Concepto.Descripcion, Concepto.Tipo,Concepto.IdConcepto, Historico_Nomina.IdHistorico As IdHistorico2," & _
        " Historico_Nomina.Fecha_Ini_Nom , Historico_Nomina.Fecha_Fin_Nom,Historico_Nomina.Periodo FROM Historico_Reng_Nomina INNER JOIN Historico_Nomina ON " & _
        " Historico_Reng_Nomina.IdRecibos = Historico_Nomina.IdHistorico AND Historico_Reng_Nomina.fecha_gen = Historico_Nomina.Fecha_Ini_Nom " & _
        " INNER JOIN Concepto ON Historico_Reng_Nomina.idconcepto = Concepto.IdConcepto Where (Historico_Nomina.IdEmpleado = " & _
        IdEmpla & ") AND (Historico_Reng_Nomina.fecha_gen = '" & FechaTemp & "') ORDER BY Concepto.IdConcepto"

Set bdata = CrearRS(CSql)
If Not (bdata.EOF) Then
    bdata.MoveFirst
    LblNoRecibo.Caption = Format(bdata.Fields("IdHistorico2").Value, "00000")
    LblDel = bdata.Fields("Fecha_Ini_Nom").Value
    LblAl = bdata.Fields("Fecha_Fin_Nom").Value
    LblNroPeriodo.Caption = bdata.Fields("Periodo").Value
    i = 1
    Do While Not bdata.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(i, 1) = Format(bdata.Fields("IdConcepto"), "00000")
        IdReci = bdata.Fields("IdHistorico")
        DMGrid1.ValorCelda(i, 2) = bdata.Fields("Detalle")
        
        If CDbl(bdata.Fields("Cantidad")) = 0 Then
            DMGrid1.ValorCelda(i, 3) = 0
        Else
            DMGrid1.ValorCelda(i, 3) = CDbl(bdata.Fields("Cantidad"))
        End If
        
        If Val(bdata.Fields("Tipo")) = 0 Then    ' 0 = asignaciones
            DMGrid1.ValorCelda(i, 4) = CDbl(bdata.Fields("valorn"))
            DMGrid1.ValorCelda(i, 5) = ""
            DMGrid1.ValorCelda(i, 6) = ""
            DMGrid1.ValorCelda(i, 7) = ""
            TAsignaciones = TAsignaciones + CDbl(bdata.Fields("valorn"))
        ElseIf Val(bdata.Fields("Tipo")) = 1 Or Val(bdata.Fields("Tipo")) = 2 Then    ' 1 = deducciones     2 = retenciones
            DMGrid1.ValorCelda(i, 4) = ""
            DMGrid1.ValorCelda(i, 5) = CDbl(bdata.Fields("valorn"))
            DMGrid1.ValorCelda(i, 6) = ""
            DMGrid1.ValorCelda(i, 7) = ""
            TDeducciones = TDeducciones + CDbl(bdata.Fields("valorn"))
        Else
            DMGrid1.ValorCelda(i, 4) = ""
            DMGrid1.ValorCelda(i, 5) = ""
            DMGrid1.ValorCelda(i, 6) = CDbl(bdata.Fields("valorn"))
            DMGrid1.ValorCelda(i, 7) = ""
            TOtros = TOtros + CDbl(bdata.Fields("valorn"))
        End If
            DMGrid1.PaintMGrid
            bdata.MoveNext
        i = i + 1
    Loop
Else
    DMGrid1.Rows = 0
    DMGrid1.Clear
    DMGrid1.PaintMGrid
    LblNoRecibo = "00000"
End If

LblTotalAsignacion.Caption = Format(TAsignaciones, "#,##0.00")
LblTotalDeduccion.Caption = Format(TDeducciones, "#,##0.00")
LblTotalOtros.Caption = Format(TOtros, "#,##0.00")
LblNetoCancelar.Caption = Format((CDbl(LblTotalAsignacion.Caption) - CDbl(LblTotalDeduccion.Caption)), "#,##0.00")
bdata.Close
End Sub
Private Sub BtnAgregar_Click()
Dim Band As Boolean

Band = False
For i = 1 To DMGrid1.Rows
    If Trim(DMGrid1.ValorCelda(i, 1) = "") Then
        Band = True
        Exit For
    End If
Next i

If Band = False Then
    DMGrid1.Rows = DMGrid1.Rows + 1
End If

DMGrid1.PaintMGrid

End Sub

Private Sub BtnAgregarConceptos_Click()
FrmListaConceptos.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAnterior_Click()
If RsEmpleados.RecordCount = 0 Then Exit Sub
RsEmpleados.MovePrevious
If RsEmpleados.BOF Then RsEmpleados.MoveLast
Call CargarEmpleado
Carga_Renglones
End Sub

Private Sub BtnBorrar_Click()
DMGrid1.RowDelete (DMGrid1.Row)
ReCalcular
DMGrid1.PaintMGrid
End Sub

Private Sub BtnBuscar_Click()
Carga_Renglones
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Form_Load
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim resp As Byte
Dim NReng As Byte
Dim CodReng As Integer
Dim ValCodReng As Double
Dim Band As Boolean

resp = MsgBox("Se procedera a guardar los en la base de datos, desea continuar?", vbQuestion + vbYesNo, "Confirmar")

If resp = 7 Then Exit Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMM Actualiza las cantidades del recibo MMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
CSql = "UPDATE Recibos SET Total_Retenciones ='" & Replace(CDbl(LblTotalOtros.Caption), ",", ".") & "', " & _
" Total_Deducciones='" & Replace(CDbl(LblTotalDeduccion.Caption), ",", ".") & "', " & _
" Total_Asignacion='" & Replace(CDbl(LblTotalAsignacion.Caption), ",", ".") & "' WHERE IdRecibos=" & Val(LblNoRecibo.Caption)
Set RsTemp = CrearRS(CSql)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMM Verificar los conceptos MMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
CSql = "SELECT * FROM Reng_Recibo WHERE IdRecibos=" & Val(LblNoRecibo.Caption) & " ORDER BY IdConcepto"
Set RsTemp = CrearRS(CSql)

NReng = CByte(DMGrid1.Rows)

While Not RsTemp.EOF
    Band = False
    For i = 1 To NReng
        CodReng = DMGrid1.ValorCelda(i, 1)      ' Contiene el codigo de la fila "i" y columna "1"
        
        If Val(RsTemp.Fields("IdConcepto").Value) = CodReng Then
            Band = True
            Exit For
        End If
    Next i
    
    If Band = False Then
        RsTemp.Delete
        RsTemp.Update
    End If
    RsTemp.MoveNext
Wend
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'  Actualiza el valor en cada Renglon del recibo
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
NReng = CByte(DMGrid1.Rows)
For i = 1 To NReng
    CodReng = DMGrid1.ValorCelda(i, 1)      ' Contiene el codigo de la fila "i" columna "1"
    
    If Trim(DMGrid1.ValorCelda(i, 4)) <> "" Then
        ValCodReng = CDbl(DMGrid1.ValorCelda(i, 4))
    Else
        ValCodReng = CDbl(DMGrid1.ValorCelda(i, 5))
    End If
    CSql = "UPDATE Reng_Recibo SET ValorN=" & Replace(ValCodReng, ",", ".") & "," & _
        " Detalle='" & DMGrid1.ValorCelda(i, 2) & "' WHERE IdRecibos=" & Val(LblNoRecibo.Caption) & " AND IdConcepto=" & CodReng
    Set RsTemp = CrearRS(CSql)
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
End Sub

Private Sub BtnImprimir_Click()
On Error GoTo Wrr
''========= ESTE ES EL CODIGO NUEVO ==========
If IdEmpl = 0 Then Exit Sub
If IdReci = "" Then Exit Sub
If NTabla <> 0 Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\Recibo_Pago.rpt"
        .Connect = "DSN=CrReporte;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{ReciboDePago.IdRecibos}=" & IdReci & ""
        .ReportTitle = "Recibo de Pago"
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

If NTabla = 0 Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\Recibo_Pago1.rpt"
        .Connect = "DSN=CrReporte;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{ReciboDePago1.IdHistorico}=" & Val(LblNoRecibo.Caption) & ""
        .ReportTitle = "Historico de Recibo de Pago"
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

Exit Sub

Wrr:
    MsgBox Err.Number & " :" & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & "NOMBRE DE EQUIPO: " & NombreEquipo

End Sub

Private Sub BtnSiguiente_Click()

If RsEmpleados.RecordCount = 0 Then Exit Sub
RsEmpleados.MoveNext
If RsEmpleados.EOF Then RsEmpleados.MoveFirst
Call CargarEmpleado
Carga_Renglones
End Sub

Private Sub DMGrid1_AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
ReCalcular
DMGrid1.PaintMGrid
End Sub

Private Sub DMGrid1_KeyPress(KeyAscii As Integer)

CodDMGrid1 = DMGrid1.ValorCelda(DMGrid1.Row, 1)
If Val(CodDMGrid1) = 0 Then Exit Sub
NomDMGrid1 = DMGrid1.ValorCelda(DMGrid1.Row, 2)

CSql = "SELECT Restringido,Tipo FROM Concepto WHERE IdConcepto=" & CodDMGrid1
Set RsTemp = CrearRS(CSql)

If Val(RsTemp.Fields("Restringido").Value) = 1 Then
    MsgBox "El Concepto " & NomDMGrid1 & " no puede ser modificado!", vbExclamation + vbOKOnly, "NO SE PUEDE MODIFICAR!"
    KeyAscii = 0
ElseIf Val(RsTemp.Fields("Tipo").Value) = 0 And DMGrid1.Col <> 4 Then
    MsgBox "El Concepto " & NomDMGrid1 & " es de tipo ASIGNACIÓN!", vbExclamation + vbOKOnly, "Modifique solo la columna Asignación!"
    KeyAscii = 0
ElseIf Val(RsTemp.Fields("Tipo").Value) = 1 And DMGrid1.Col <> 5 Then
    MsgBox "El Concepto " & NomDMGrid1 & " es de tipo DEDUCCIÓN!", vbExclamation + vbOKOnly, "Modifique solo la columna Deducción!"
    KeyAscii = 0
ElseIf Val(RsTemp.Fields("Tipo").Value) = 2 And DMGrid1.Col <> 4 Then
    MsgBox "El Concepto " & NomDMGrid1 & " es de tipo OTROS!", vbExclamation + vbOKOnly, "Modifique solo la columna Otros!"
    KeyAscii = 0
End If
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)

If Button = vbRightButton Then
    If lCol = 2 Then
        VarTCamp = InputBox("Ingrese el nombre del Campo:", "Modificar el nombre del concepto para el recibo", DMGrid1.ValorCelda(lRow, 2))
        If Trim(VarTCamp) <> "" Then
            DMGrid1.ValorCelda(lRow, 2) = VarTCamp
            DMGrid1.PaintMGrid
        End If
    End If
End If
End Sub

Private Sub Form_Activate()
Tipo = "RPagos"
End Sub

Private Sub Form_Load()
Centrar Me

CSql = "Select * From Dat_Admin"
Set RsDatAdmin = CrearRS(CSql)

'LblNoRecibo.Caption = Format(RsDatAdmin.Fields("U_Recibo").Value + 1, "00000")

CSql = "Select * From Empleados where activo=1"
Set RsEmpleados = CrearRS(CSql)

If RsEmpleados.RecordCount = 0 Then Exit Sub

Grid2
CargarEmpleado


CSql = "SELECT Fecha_Ini_Nom FROM Recibos"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    DTPicker1.Value = RsTemp.Fields("Fecha_Ini_Nom").Value
Else
    CSql = "SELECT Fecha_Prox_Gen FROM Grupo ORDER BY fecha_prox_gen"
    Set RsTemp = CrearRS(CSql)
    RsTemp.MoveLast
    DTPicker1.Value = RsTemp.Fields("Fecha_Prox_Gen").Value
End If

Carga_Renglones


End Sub

Sub CargarEmpleado()
On Error Resume Next
' Condicional que solo se usa en caso de que el formulario HISTORICO DE NÓMINA lo llame
' La variable IDEMPLA solo se le debe dar valores desde otros FORMULARIOS
If IdEmpla <> 0 Then
    CSql = "SELECT * FROM Empleados WHERE IdEmpleado=" & IdEmpla
    Set RsEmpleados = CrearRS(CSql)
End If

If RsEmpleados.RecordCount > 0 Then
    'RsEmpleados.MoveFirst
    IdEmpl = RsEmpleados.Fields("IdEmpleado").Value
    LblCodigo.Caption = Format(RsEmpleados.Fields("IdEmpleado").Value, "0000#")
    LblCedula.Caption = RsEmpleados.Fields("Cedula").Value
    LblApellido.Caption = RsEmpleados.Fields("Apellido").Value
    LblNombre.Caption = RsEmpleados.Fields("Nombre").Value
    
    CSql = "SELECT Empleados.IdEmpleado, Empleados.Nombre, Empleados.Apellido, Empleados.Departamentos, " & _
        "Departamentos.Descripcion, Cargos.Cargo, Grupo.Descripcion AS Expr1 FROM Empleados INNER JOIN " & _
        " Departamentos ON Empleados.Departamentos = Departamentos.IdDepartamento INNER JOIN Cargos ON " & _
        " Empleados.Cargo = Cargos.IdCargos INNER JOIN Grupo ON Empleados.Id_Grupo = Grupo.Id_Grupo " & _
        " WHERE Empleados.IdEmpleado=" & IdEmpl & " ORDER BY Empleados.IdEmpleado"
    Set RsTemp = CrearRS(CSql)
       
    LblDepartamento.Caption = RsTemp.Fields("Descripcion").Value
    LblCargo.Caption = RsTemp.Fields("Cargo").Value
    LblTipoNomina.Caption = RsTemp.Fields("Expr1").Value
    
    If Not IsNull(RsEmpleados.Fields("Photo").Value) Then
       If RsEmpleados.Fields("Photo").Value <> "" And Dir(FotoEmp & "\" & RsEmpleados.Fields("Photo").Value) <> "" Then
           Image1.Picture = LoadPicture(FotoEmp & "\" & RsEmpleados.Fields("Photo").Value)
           FotoP = RsEmpleados.Fields("Photo").Value
       Else
           Image1.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
           FotoP = ""
       End If
    Else
       Image1.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
       FotoP = ""
    End If
End If
End Sub
 

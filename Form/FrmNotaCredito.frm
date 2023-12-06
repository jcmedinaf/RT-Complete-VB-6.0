VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmNotaCredito 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Crédito"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14355
   Icon            =   "FrmNotaCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14355
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   5535
      Left            =   11640
      TabIndex        =   72
      Top             =   3240
      Width           =   2655
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Factura  menos Nota de Credito"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   240
         TabIndex        =   85
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA 12%:"
         Height          =   195
         Left            =   360
         TabIndex        =   84
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label LblImpuestoG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   83
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar:"
         Height          =   195
         Left            =   360
         TabIndex        =   82
         Top             =   4800
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal:"
         Height          =   195
         Left            =   360
         TabIndex        =   81
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Exentos:"
         Height          =   195
         Left            =   360
         TabIndex        =   80
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Imponible:"
         Height          =   195
         Left            =   360
         TabIndex        =   79
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label LblSubTotalG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   78
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label LblBaseImponibleG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   77
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label LblExentosG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   76
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label LblTotalGeneralG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   75
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label LblDescuentosG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   74
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   360
         TabIndex        =   73
         Top             =   3360
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "       PRODUCTOS / SERVICIO EXCLUIDO       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   6120
      Width           =   11415
      Begin SystemOncoAmerica.DMGrid DMGrid2 
         Height          =   2295
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4048
         Object.Width           =   8025
         Object.Height          =   2265
         LineRowBackColor=   14737632
         Editable        =   -1  'True
         DrawColorGrid   =   1
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   8400
         TabIndex        =   34
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label LblDescuentosNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   33
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label LblTotalGeneralNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   32
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label LblExentosNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   31
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label LblBaseImponibleNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   30
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label LblSubTotalNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Imponible:"
         Height          =   195
         Left            =   8400
         TabIndex        =   28
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Exentos:"
         Height          =   195
         Left            =   8400
         TabIndex        =   27
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal:"
         Height          =   195
         Left            =   8400
         TabIndex        =   26
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar:"
         Height          =   195
         Left            =   8400
         TabIndex        =   25
         Top             =   2220
         Width           =   1005
      End
      Begin VB.Label LblImpuestoNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   24
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA 12%:"
         Height          =   195
         Left            =   8400
         TabIndex        =   23
         Top             =   1860
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   8760
      Width           =   14175
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9000
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
      Begin VB.TextBox TxtNotaCredito 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   49
         ToolTipText     =   "Ingrese el Número de la Nota de Crédito a buscar"
         Top             =   255
         Width           =   1455
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   405
         Left            =   12960
         TabIndex        =   13
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
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
         MICON           =   "FrmNotaCredito.frx":1002
         PICN            =   "FrmNotaCredito.frx":101E
         PICH            =   "FrmNotaCredito.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardar 
         Height          =   405
         Left            =   6480
         TabIndex        =   14
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
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
         MICON           =   "FrmNotaCredito.frx":141C
         PICN            =   "FrmNotaCredito.frx":1438
         PICH            =   "FrmNotaCredito.frx":16C7
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
         Height          =   405
         Left            =   5280
         TabIndex        =   15
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Agregar"
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
         MICON           =   "FrmNotaCredito.frx":1B08
         PICN            =   "FrmNotaCredito.frx":1B24
         PICH            =   "FrmNotaCredito.frx":1CB1
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
         Height          =   405
         Left            =   11400
         TabIndex        =   16
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
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
         MICON           =   "FrmNotaCredito.frx":1EE6
         PICN            =   "FrmNotaCredito.frx":1F02
         PICH            =   "FrmNotaCredito.frx":21E4
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
         Height          =   405
         Left            =   7680
         TabIndex        =   17
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
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
         MICON           =   "FrmNotaCredito.frx":2435
         PICN            =   "FrmNotaCredito.frx":2451
         PICH            =   "FrmNotaCredito.frx":25F5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscarNC 
         Height          =   405
         Left            =   3240
         TabIndex        =   50
         ToolTipText     =   "Buscar Nota de Crédito"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
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
         MICON           =   "FrmNotaCredito.frx":2794
         PICN            =   "FrmNotaCredito.frx":27B0
         PICH            =   "FrmNotaCredito.frx":2A15
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
         Left            =   9480
         TabIndex        =   86
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
         MICON           =   "FrmNotaCredito.frx":2CA7
         PICN            =   "FrmNotaCredito.frx":2CC3
         PICH            =   "FrmNotaCredito.frx":2DE8
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
         Caption         =   "Nota de Crédito No."
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   345
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3135
      Index           =   2
      Left            =   11640
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.TextBox TxtNoFactura 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtPorCobrar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   2640
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DtpFechaFactura 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52756481
         CurrentDate     =   39932
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscarFactura 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         ToolTipText     =   "Buscar Factura"
         Top             =   600
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
         MICON           =   "FrmNotaCredito.frx":3078
         PICN            =   "FrmNotaCredito.frx":3094
         PICH            =   "FrmNotaCredito.frx":32F9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por Cobrar:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label LblNoNotaCredito 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1275
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Nota de Crédito"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factura No."
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   405
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota de Crédito"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   315
         TabIndex        =   6
         Top             =   120
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos de la Facturación"
      Height          =   3135
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11415
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Paciente"
         Height          =   1575
         Left            =   120
         TabIndex        =   61
         Top             =   1440
         Width           =   11175
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   4920
            TabIndex        =   66
            Top             =   620
            Width           =   3135
         End
         Begin VB.TextBox TxtApellido 
            Height          =   375
            Left            =   4920
            TabIndex        =   65
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox TxtCedula 
            Height          =   375
            Left            =   1320
            TabIndex        =   64
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtDirecionPaciente 
            Height          =   375
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   1010
            Width           =   9735
         End
         Begin VB.TextBox TxtTelefonoPaciente 
            Height          =   375
            Left            =   1320
            TabIndex        =   62
            Top             =   620
            Width           =   2175
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   195
            TabIndex        =   71
            Top             =   330
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   3840
            TabIndex        =   70
            Top             =   710
            Width           =   765
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   3840
            TabIndex        =   69
            Top             =   330
            Width           =   765
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   195
            TabIndex        =   68
            Top             =   1100
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   195
            TabIndex        =   67
            Top             =   710
            Width           =   675
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Cliente"
         Height          =   1095
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   11175
         Begin VB.TextBox TxtRazonSocial 
            Height          =   375
            Left            =   4920
            TabIndex        =   56
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox TxtDireccionCliente 
            Height          =   375
            Left            =   4920
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   600
            Width           =   6135
         End
         Begin VB.TextBox TxtRif 
            Height          =   375
            Left            =   1320
            TabIndex        =   53
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox TxtTelefonoCliente 
            Height          =   375
            Left            =   1320
            TabIndex        =   54
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   240
            TabIndex        =   60
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   3840
            TabIndex        =   59
            Top             =   690
            Width           =   720
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif.: "
            Height          =   195
            Left            =   240
            TabIndex        =   58
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   3840
            TabIndex        =   57
            Top             =   330
            Width           =   990
         End
      End
      Begin VB.Label LblValorImpuesto 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7440
         TabIndex        =   4
         Top             =   6480
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "        FACTURA           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   11415
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4260
         Object.Width           =   8025
         Object.Height          =   2385
         LineRowBackColor=   -2147483648
         DrawColorGrid   =   1
      End
      Begin VB.Label LblRetenciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9840
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Retenciones:"
         Height          =   195
         Left            =   8400
         TabIndex        =   47
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   8400
         TabIndex        =   46
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label LblDescuentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   45
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label LblTotalGeneral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label LblExentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   43
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label LblBaseImponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   42
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label LblSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   41
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Imponible:"
         Height          =   195
         Left            =   8400
         TabIndex        =   40
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Exentos:"
         Height          =   195
         Left            =   8400
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal:"
         Height          =   195
         Left            =   8400
         TabIndex        =   38
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar:"
         Height          =   195
         Left            =   8400
         TabIndex        =   37
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label LblImpuesto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9600
         TabIndex        =   36
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA 12%:"
         Height          =   195
         Left            =   8400
         TabIndex        =   35
         Top             =   2160
         Width           =   645
      End
   End
End
Attribute VB_Name = "FrmNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDatAdmin As New ADODB.Recordset
Dim RsBuscarFactura As New ADODB.Recordset
Dim RsBuscarRenglonFactura As New ADODB.Recordset
Dim RsPaciente As New ADODB.Recordset
Dim RsCliente As New ADODB.Recordset
Dim RsProductos As New ADODB.Recordset
Dim RsRetenciones As New ADODB.Recordset
Dim Selec(0 To 100) As Boolean
Dim IdFactur
Dim IVA

Private Sub BtnBuscarFactura_Click()
Limpiar

For i = 0 To 100
    Selec(i) = False
Next i

If Trim(TxtNoFactura.Text) = "" Then MsgBox "Ingrese el Numero de factura!", vbExclamation + vbOKOnly, "Error": Exit Sub
CSql = "Select * From C_Cobrar Where N_Factura='" & TxtNoFactura.Text & "' AND Anulada=0 AND C_NC=0"
Set RsBuscarFactura = CrearRS(CSql)

FrmNotaCredito.Caption = "Nota de Crédito"
If RsBuscarFactura.RecordCount = 0 Then: MsgBox "No se encontro el Nro de Factura ingresado!", vbInformation + vbOKOnly, "Error": IdFactur = "": Exit Sub
If Not RsBuscarFactura.Fields("impresa") Then
    MsgBox "Debe imprimir la Factura Nro. " & Trim(TxtNoFactura.Text) & " para poder realizar la Nota de Crédito!", vbExclamation + vbOKOnly, "Error"
    IdFactur = ""
    BtnGuardar.Enabled = False
    Exit Sub
Else
    BtnGuardar.Enabled = True
End If

IdFactur = RsBuscarFactura.Fields("N_Factura").Value

CSql = "Select * From Paciente Where IdPaciente='" & RsBuscarFactura.Fields("IdPaciente").Value & "'"
Set RsPaciente = CrearRS(CSql)

CSql = "Select * From Cliente Where IdCliente='" & RsBuscarFactura.Fields("IdCliente").Value & "'"
Set RsCliente = CrearRS(CSql)

TxtRif.Text = RsCliente.Fields("Rif").Value
TxtRazonSocial.Text = RsCliente.Fields("Razon").Value
TxtDireccionCliente.Text = RsCliente.Fields("DireccionC").Value
TxtTelefonoCliente.Text = RsCliente.Fields("Telefono").Value

TxtCedula.Text = RsPaciente.Fields("CedulaP").Value
TxtApellido.Text = RsPaciente.Fields("ApellidoP").Value
TxtNombre.Text = RsPaciente.Fields("NombreP").Value
TxtDirecionPaciente.Text = RsPaciente.Fields("DireccionP").Value
TxtTelefonoPaciente.Text = RsPaciente.Fields("Codigo").Value & " - " & RsPaciente.Fields("Telefono").Value

CSql = "Select * From Reng_Cobrar Where N_Factura='" & TxtNoFactura.Text & "'"
Set RsBuscarRenglonFactura = CrearRS(CSql)

DMGrid1.Rows = 0

Do While Not RsBuscarRenglonFactura.EOF
    With LstFactura

        CSql = "Select * From Productos Where IdProducto='" & RsBuscarRenglonFactura.Fields("Cod_producto").Value & "'"
        Set RsProductos = CrearRS(CSql)

        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarRenglonFactura.Fields("Cod_producto").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsProductos.Fields("Descripcion").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarRenglonFactura.Fields("Cantidad").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = Format(RsBuscarRenglonFactura.Fields("Precio").Value, "#,##0.00")
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = Format(RsBuscarRenglonFactura.Fields("Iva").Value, "#,##0.00")
        DMGrid1.ValorCelda(DMGrid1.Rows, 6) = Format(RsBuscarRenglonFactura.Fields("Descuento").Value, "#,##0.00")
        DMGrid1.ValorCelda(DMGrid1.Rows, 7) = Format(RsBuscarRenglonFactura.Fields("SubTotal").Value, "#,##0.00")
        
        If RsProductos.Fields("Impuesto").Value Then
            DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
        Else
            DMGrid1.RowBackColor DMGrid1.Rows, RGB(221, 221, 221)
        End If

    End With
    RsBuscarRenglonFactura.MoveNext
Loop

TxtPorCobrar.Text = Format(RsBuscarFactura.Fields("PorCobrar").Value, "#,##0.####")
LblSubTotal.Caption = Format(RsBuscarFactura.Fields("SubTotal").Value, "#,##0.00")
LblImpuesto.Caption = Format(RsBuscarFactura.Fields("Impuesto").Value, "#,##0.00")
LblTotalGeneral.Caption = Format(RsBuscarFactura.Fields("Monto").Value, "#,##0.00")

CSql = "Select Sum(Monto) As TotalMonto From Cobros Where N_Factura='" & TxtNoFactura.Text & "'"
Set RsRetenciones = CrearRS(CSql)

If RsRetenciones.RecordCount <> 0 Then
    LblRetenciones.Caption = Format(RsRetenciones.Fields("TotalMonto").Value, "#,##0.00")
Else
    LblRetenciones.Caption = Format(0, "#,##0.00")
End If
DMGrid1.PaintMGrid
Calcular2 ' Crea los calculos para la factura
FrmNotaCredito.Caption = "Nota de Crédito                           Factura Nro. " & IdFactur

CSql = "Select MAX(N_NC)+1 As NuevoId from C_Cobrar"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If Not IsNull(RsTemp.Fields("NuevoId")) Then
        NuevaIdNC = RsTemp.Fields("NuevoId")
    Else
        NuevaIdNC = 1
    End If
Else
    NuevaIdNC = 1
End If

End Sub

Private Sub BtnBuscarNC_Click()
FrmNotaCredito.Caption = "Nota de Crédito"
Limpiar
For i = 0 To 100
    Selec(i) = False
Next i

IdFactur = 0

If Trim(TxtNotaCredito.Text) = "" Then MsgBox "Ingrese el Numero de la Nota de Crédito!", vbExclamation + vbOKOnly, "Error": Exit Sub

CSql = "Select * From C_Cobrar Where N_NC='" & TxtNotaCredito.Text & "' AND Anulada=0 AND C_NC=1"
Set RsBuscarFactura = CrearRS(CSql)

If RsBuscarFactura.RecordCount = 0 Then: MsgBox "No se encontro la Nota de Crédito ingresada!", vbInformation + vbOKOnly, "Informacion": Exit Sub
If RsBuscarFactura.Fields("Anulada") = 1 Then: MsgBox "La nota de crédito se encontro pero la Factura fue anulada anteriormente!", vbExclamation + vbOKOnly, "Informacion"

IdFactur = Val(RsBuscarFactura.Fields("N_FA").Value)

CSql = "Select * From Paciente Where IdPaciente='" & RsBuscarFactura.Fields("IdPaciente").Value & "'"
Set RsPaciente = CrearRS(CSql)

CSql = "Select * From Cliente Where IdCliente='" & RsBuscarFactura.Fields("IdCliente").Value & "'"
Set RsCliente = CrearRS(CSql)

TxtRif.Text = RsCliente.Fields("Rif").Value
TxtRazonSocial.Text = RsCliente.Fields("Razon").Value
TxtDireccionCliente.Text = RsCliente.Fields("DireccionC").Value
TxtTelefonoCliente.Text = RsCliente.Fields("Telefono").Value

TxtCedula.Text = RsPaciente.Fields("CedulaP").Value
TxtApellido.Text = RsPaciente.Fields("ApellidoP").Value
TxtNombre.Text = RsPaciente.Fields("NombreP").Value
TxtDirecionPaciente.Text = RsPaciente.Fields("DireccionP").Value
TxtTelefonoPaciente.Text = RsPaciente.Fields("Codigo").Value & " - " & RsPaciente.Fields("Telefono").Value

CSql = "Select * From Reng_Cobrar Where N_NC='" & TxtNotaCredito.Text & "' AND C_NC=1"
Set RsBuscarRenglonFactura = CrearRS(CSql)

DMGrid1.Rows = 0
DMGrid2.Rows = 0

Do While Not RsBuscarRenglonFactura.EOF
    With LstFactura

        CSql = "Select * From Productos Where IdProducto='" & RsBuscarRenglonFactura.Fields("Cod_producto").Value & "'"
        Set RsProductos = CrearRS(CSql)

        If Val(RsBuscarRenglonFactura.Fields("SubTotal").Value) > 0 Then
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarRenglonFactura.Fields("Cod_producto").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsProductos.Fields("Descripcion").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarRenglonFactura.Fields("Cantidad").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = Format(RsBuscarRenglonFactura.Fields("Precio").Value, "#,##0.00")
            DMGrid1.ValorCelda(DMGrid1.Rows, 5) = Format(RsBuscarRenglonFactura.Fields("Iva").Value, "#,##0.00")
            DMGrid1.ValorCelda(DMGrid1.Rows, 6) = Format(RsBuscarRenglonFactura.Fields("Descuento").Value, "#,##0.00")
            DMGrid1.ValorCelda(DMGrid1.Rows, 7) = Format(RsBuscarRenglonFactura.Fields("SubTotal").Value, "#,##0.00")
            
            If RsProductos.Fields("Impuesto").Value Then
                DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
            Else
                DMGrid1.RowBackColor DMGrid1.Rows, RGB(221, 221, 221)
            End If
        Else
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsBuscarRenglonFactura.Fields("Cod_producto").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsProductos.Fields("Descripcion").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsBuscarRenglonFactura.Fields("Cantidad").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = Format(RsBuscarRenglonFactura.Fields("Precio").Value, "#,##0.00")
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = Format(CDbl(RsBuscarRenglonFactura.Fields("Iva").Value) * -1, "#,##0.00")
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = Format(RsBuscarRenglonFactura.Fields("Descuento").Value, "#,##0.00")
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = Format(CDbl(RsBuscarRenglonFactura.Fields("SubTotal").Value) * -1, "#,##0.00")
            DMGrid2.RowBackColor DMGrid2.Rows, RGB(255, 255, 255)
        End If

    End With
    RsBuscarRenglonFactura.MoveNext
Loop

TxtPorCobrar.Text = Format(RsBuscarFactura.Fields("PorCobrar").Value, "#,##0.####")
LblSubTotal.Caption = Format(RsBuscarFactura.Fields("SubTotal").Value, "#,##0.00")
LblImpuesto.Caption = Format(RsBuscarFactura.Fields("Impuesto").Value, "#,##0.00")
LblTotalGeneral.Caption = Format(RsBuscarFactura.Fields("Monto").Value, "#,##0.00")

CSql = "Select Sum(Monto) As TotalMonto From Cobros Where N_NC='" & TxtNotaCredito.Text & "' AND C_NC=1"
Set RsRetenciones = CrearRS(CSql)

If RsRetenciones.RecordCount <> 0 Then
    LblRetenciones.Caption = Format(RsRetenciones.Fields("TotalMonto").Value, "#,##0.00")
Else
    LblRetenciones.Caption = Format(0, "#,##0.00")
End If

DMGrid1.PaintMGrid
DMGrid2.PaintMGrid
Calcular2   ' Genera los calculos para la factura
calcular    ' Genera los calculos para la Nota de Credito
Calcular3   ' Genera el TOTAL GENERAL del Monto de la Factura MENOS el Monto de la Nota de Credito
FrmNotaCredito.Caption = "Nota de Crédito                           Factura Nro. " & IdFactur
LblNoNotaCredito = Format(Trim(TxtNotaCredito.Text), "0000")
IdFactur = 0
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub ChameleonBtn1_Click()
Limpiar
End Sub

Private Sub BtnDesHacer_Click()
IdFactur = ""
End Sub

Private Sub BtnGuardar_Click()
Dim RsTemp As New ADODB.Recordset
Dim NuevaIdNC As Integer
Dim IdP As Integer
Dim IdPUsua As Integer
Dim DespP As String
Dim FechaP As String
Dim CantP As Integer
Dim PUnit As Double
Dim IIva As Double
Dim DescP As Double
Dim TotalP As Double
Dim SubTotalP As Double

If Trim(IdFactur) = "" Then: MsgBox "Debe Seleccionar una factura antes de guardar una nota de crédito!", vbExclamation + vbOKOnly, "Error": Exit Sub
If Val(Trim(IdFactur)) = 0 Then: MsgBox "Debe Seleccionar una factura antes de guardar una nota de crédito!", vbExclamation + vbOKOnly, "Error": Exit Sub

'CSql = "Select * From C_Cobrar Where (N_Factura='" & IdFactur & "' OR N_FA='" & IdFactur & "')AND Anulada=0"
'Set RsBuscarFactura = CrearRS(CSql)
'If RsBuscarFactura.Fields("C_NC").Value Then: MsgBox "El Nro de Factura " & IdFactur & " ya contiene una Nota de Crédito!", vbExclamation + vbOKOnly, "Error": Exit Sub

If DMGrid2.Rows = 0 Then
    MsgBox "No se puede crear una nota de credito si no hay productos o servicios Excluidos!", vbExclamation + vbOKOnly, "Error"
    Exit Sub
ElseIf Val(DMGrid2.ValorCelda(1, 1)) = 0 Then
    MsgBox "No se puede crear una nota de credito si no hay productos o servicios Excluidos!", vbExclamation + vbOKOnly, "Error"
    Exit Sub
End If

CSql = "Select MAX(N_NC)+1 As NuevoId from C_Cobrar"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If Not IsNull(RsTemp.Fields("NuevoId")) Then
        NuevaIdNC = RsTemp.Fields("NuevoId")
    Else
        NuevaIdNC = 1
    End If
Else
    NuevaIdNC = 1
End If

'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
CSql = "UPDATE C_Cobrar SET N_NC=" & NuevaIdNC & ", C_NC=1, N_FA=" & IdFactur & ", N_Factura=0 Where N_Factura='" & IdFactur & "' AND Anulada=0 AND C_NC=0"
'CSql = "UPDATE C_Cobrar SET N_NC=" & NuevaIdNC & ", C_NC=1, N_FA=" & IdFactur & ", N_Factura=0, Subtotal=Subtotal*-1, Exento=Exento*-1, Impuesto=Impuesto*-1,Monto=Monto*-1, PorCobrar=PorCobrar*-1 Where N_Factura='" & IdFactur & "' AND Anulada=0 AND C_NC=0"
Set RsBuscarFactura = CrearRS(CSql)
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
CSql = "UPDATE Reng_Cobrar SET N_NC=" & NuevaIdNC & ", C_NC=1, N_FA=" & IdFactur & ", N_Factura=0  Where N_Factura='" & IdFactur & "' AND C_NC=0"
Set RsBuscarRenglonFactura = CrearRS(CSql)
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
CSql = "UPDATE Cobros SET N_NC=" & NuevaIdNC & ", C_NC=1, N_FA=" & IdFactur & ", N_Factura=0  Where N_Factura='" & IdFactur & "' AND C_NC=0"
Set RsRetenciones = CrearRS(CSql)
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm

CSql = "SELECT Fecha, IdUsuario FROM C_Cobrar Where N_FA='" & IdFactur & "'"
Set RsBuscarRenglonFactura = CrearRS(CSql)

FechaP = RsBuscarRenglonFactura.Fields("Fecha").Value
IdPUsua = RsBuscarRenglonFactura.Fields("IdUsuario").Value

For i = 1 To DMGrid2.Rows
    
    IdP = Val(DMGrid2.ValorCelda(i, 1))
    DespP = DMGrid2.ValorCelda(i, 2)
    CantP = Val(DMGrid2.ValorCelda(i, 3))
    PUnit = CDbl(DMGrid2.ValorCelda(i, 4))
    IIva = CDbl(DMGrid2.ValorCelda(i, 5))
    DescP = CDbl(DMGrid2.ValorCelda(i, 6))
    SubTotalP = CantP * PUnit + IIva
    TotalP = CDbl(DMGrid2.ValorCelda(i, 7))
    
    CSql = "INSERT INTO Reng_Cobrar VALUES ( 0," & IdP & ",'" & FechaP & "'," & PUnit & "," & _
                                            CantP & "," & Replace(IIva * -1, ",", ".") & "," & DescP * -1 & "," & IdPUsua & "," & _
                                            Replace(SubTotalP * -1, ",", ".") & ", " & NuevaIdNC & "," & IdFactur & ",1)"
    
    Set RsBuscarRenglonFactura = CrearRS(CSql)

Next i
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm

MsgBox "La nota de credito Nro. " & NuevaIdNC & " ha sido Guardada!", vbInformation + vbOKOnly, "Operacion Exitosa!"

CSql = "update dat_admin set u_notacredito = " & NuevaIdNC
Set RsBuscarRenglonFactura = CrearRS(CSql)

End Sub
 
Private Sub BtnImprimir_Click()
''========= ESTE ES EL CODIGO NUEVO ==========
If DMGrid2.Rows > 0 Then

    With CrystalReport1
        .ReportFileName = RutaInformes & "\NotaCredito.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{NotaCredito.N_Nc} >= " & Val(LblNoNotaCredito.Caption)
        .WindowTitle = "Nota de Credito No. " & LblNoNotaCredito.Caption
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)

If Button = vbRightButton Then
    If lRow = 0 Then Exit Sub
    If Val(DMGrid1.ValorCelda(lRow, 1)) = 0 Then Exit Sub
    
    If Selec(lRow) = True Then Exit Sub
    
    If Not Verificar_Pase(lRow) Then Exit Sub
    
    DMGrid1.RowBackColor lRow, RGB(255, 255, 155)
    DMGrid1.PaintMGrid
    
    If Not Val(DMGrid2.ValorCelda(DMGrid2.Rows, 1)) = 0 Then DMGrid2.Rows = DMGrid2.Rows + 1
    
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = DMGrid1.ValorCelda(DMGrid1.Row, 1)
    DMGrid2.ValorCelda(DMGrid2.Rows, 2) = DMGrid1.ValorCelda(DMGrid1.Row, 2)
    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = DMGrid1.ValorCelda(DMGrid1.Row, 3)
    DMGrid2.ValorCelda(DMGrid2.Rows, 4) = DMGrid1.ValorCelda(DMGrid1.Row, 4)
    DMGrid2.ValorCelda(DMGrid2.Rows, 5) = DMGrid1.ValorCelda(DMGrid1.Row, 5)
    DMGrid2.ValorCelda(DMGrid2.Rows, 6) = DMGrid1.ValorCelda(DMGrid1.Row, 6)
    DMGrid2.ValorCelda(DMGrid2.Rows, 7) = DMGrid1.ValorCelda(DMGrid1.Row, 7)
    
    DMGrid2.RowBackColor DMGrid1.Rows, DMGrid1.DefRowBackColor
    DMGrid2.PaintMGrid
    Selec(lRow) = True
    calcular
    Calcular3

End If
End Sub

Private Sub DMGrid2_AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
Dim CodP As Integer
Dim Cant As Integer
Dim Cant2 As Integer

Cant = 0
Cant2 = 0
CodP = DMGrid2.ValorCelda(lRow, 1)

    For i = 1 To DMGrid2.Rows
        If Val(DMGrid2.ValorCelda(i, 1)) = CodP Then
            Cant = Cant + Val(DMGrid2.ValorCelda(i, 3))
        End If
    Next i
    
    For i = 1 To DMGrid1.Rows
        If Val(DMGrid1.ValorCelda(i, 1)) = CodP Then
            Cant2 = Cant2 + Val(DMGrid1.ValorCelda(i, 3))
        End If
    Next i
    
    If Cant > Cant2 Then
        MsgBox "La cantidad ingresada EXCEDE la cantidad total de este producto en la factura!", vbExclamation + vbOKOnly, "Error"
        DMGrid2.ValorCelda(lRow, 3) = "1"
    End If
    calcular
End Sub

Private Sub Form_Load()
Centrar Me

CSql = "Select * From Dat_Admin"
Set RsDatAdmin = CrearRS(CSql)

LblNoNotaCredito.Caption = Format(RsDatAdmin.Fields("U_NotaCredito").Value + 1, "0000")
IniDMGrid
ValorImpuesto
DtpFechaFactura = Now
End Sub

Sub ValorImpuesto()
Dim RsValorImpuesto As New ADODB.Recordset
' carga los diferentes valores de los impuestos
 
CSql = "SELECT IVA1 FROM Dat_admin"
RsValorImpuesto.Open CSql, Cnn, adOpenDynamic, adLockPessimistic

    IVA = RsValorImpuesto.Fields("IVA1").Value
    LblImpuesto.Caption = "(" & RsValorImpuesto.Fields("IVA1").Value & " %)"

RsValorImpuesto.Close

End Sub


Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 7
DMGrid1.Rows = 1
DMGrid1.RowBackColor 1, vbWhite
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1
DMGrid1.DColumnas(6).Alignment = 1
DMGrid1.DColumnas(7).Alignment = 1
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(4).Locked = True
DMGrid1.DColumnas(5).Locked = True
DMGrid1.DColumnas(6).Locked = True
DMGrid1.DColumnas(7).Locked = True
DMGrid1.DColumnas(4).IsNumber = True
DMGrid1.DColumnas(5).IsNumber = True
DMGrid1.DColumnas(6).IsNumber = True
DMGrid1.DColumnas(7).IsNumber = True
'DMGrid1.DColumnas(3).IsNumber = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 40 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(6).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(7).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(1).Caption = "Codigo"
DMGrid1.DColumnas(2).Caption = "Descripcion"
DMGrid1.DColumnas(3).Caption = "Cant"
DMGrid1.DColumnas(4).Caption = "Precio/Unitario"
DMGrid1.DColumnas(5).Caption = "Impuesto"
DMGrid1.DColumnas(6).Caption = "Descuento"
DMGrid1.DColumnas(7).Caption = "Total"

DMGrid2.Cols = 7
DMGrid2.Rows = 1
DMGrid2.RowBackColor 1, vbWhite
DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 1
DMGrid2.DColumnas(4).Alignment = 1
DMGrid2.DColumnas(5).Alignment = 1
DMGrid2.DColumnas(6).Alignment = 1
DMGrid2.DColumnas(7).Alignment = 1
DMGrid2.DColumnas(1).Locked = True
DMGrid2.DColumnas(2).Locked = True
DMGrid2.DColumnas(3).Locked = False
DMGrid2.DColumnas(4).Locked = True
DMGrid2.DColumnas(5).Locked = True
DMGrid2.DColumnas(6).Locked = True
DMGrid2.DColumnas(7).Locked = True
DMGrid2.DColumnas(4).IsNumber = True
DMGrid2.DColumnas(5).IsNumber = True
DMGrid2.DColumnas(6).IsNumber = True
DMGrid2.DColumnas(7).IsNumber = True
'DMGrid2.DColumnas(3).IsNumber = True
DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 40 / 100) - 300
DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(4).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(5).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(6).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(7).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(1).Caption = "Codigo"
DMGrid2.DColumnas(2).Caption = "Descripcion"
DMGrid2.DColumnas(3).Caption = "Cant"
DMGrid2.DColumnas(4).Caption = "Precio/Unitario"
DMGrid2.DColumnas(5).Caption = "Impuesto"
DMGrid2.DColumnas(6).Caption = "Descuento"
DMGrid2.DColumnas(7).Caption = "Total"
End Sub

Private Sub TxtNoFactura_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 Then: KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then BtnBuscarFactura_Click
If Len(TxtNoFactura.Text) > 7 And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

' Para el DMGRID2
Sub calcular()
Dim Cant, PUnit, IIva, Descu As Double
Dim Exentos, BaseImp, SubTot, TotalIVA As Double
Dim Campow As String

Exentos = 0

For i = 1 To DMGrid2.Rows

    IdProd = DMGrid2.ValorCelda(i, 1)
    
    If IdProd <> "" Then
        
        If Val(DMGrid2.ValorCelda(i, 3)) <> 0 Then
            Cant = Val(DMGrid2.ValorCelda(i, 3))
        Else
            Cant = 0
        End If
        
        If Val(DMGrid2.ValorCelda(i, 4)) <> 0 Then
            PUnit = CDbl(DMGrid2.ValorCelda(i, 4))
        Else
            PUnit = 0
        End If
        
        If Val(DMGrid2.ValorCelda(i, 5)) <> 0 Then
            IIva = CDbl(DMGrid2.ValorCelda(i, 5))
        Else
            IIva = 0
        End If
        
        If Val(DMGrid2.ValorCelda(i, 6)) <> 0 Then
            Descu = CDbl(DMGrid2.ValorCelda(i, 6))
        Else
            Descu = 0
        End If
        
        If IsNull(Cant) Then
            Cant = 1
            DMGrid2.ValorCelda(i, 3) = 1
        ElseIf Val(Cant) = 0 Then DMGrid2.ValorCelda(i, 3) = 1: Cant = 1
        End If
        If IsNull(IIva) Then
            IIva = 0
        ElseIf Val(IIva) = 0 Then IIva = 0
        End If
        If IsNull(Descu) Then
            Descu = 0
        ElseIf Val(Descu) = 0 Then Descu = 0
        End If
        
        If Val(IIva) = 0 Then
            Exentos = Exentos + ((Cant * PUnit) - Descu)
        Else
            BaseImp = BaseImp + ((Cant * PUnit) - Descu)
            DMGrid2.ValorCelda(i, 5) = CDbl(((Cant * PUnit) - Descu) * (IVA / 100))
        End If
        
        DMGrid2.ValorCelda(i, 7) = ((Cant * PUnit) - Descu) + CDbl(DMGrid2.ValorCelda(i, 5))
        SubTot = CDbl(SubTot) + CDbl(DMGrid2.ValorCelda(i, 7))
        TotalIVA = TotalIVA + IIva
        C = C + ((Cant * PUnit) - Descu) + IIva

    End If
Next i
If Val(BaseImp) = 0 Then BaseImp = 0
If Val(SubTot) = 0 Then SubTot = 0
If Val(Descu) = 0 Then Descu = 0
If Val(Exentos) = 0 Then Exentos = 0
If Val(TotalIVA) = 0 Then TotalIVA = 0


'SubTotal
LblSubTotalNC.Caption = Format(SubTot, "###,##0.#0")
'Base imponible
LblBaseImponibleNC.Caption = Format(BaseImp, "###,##0.#0")
'Exentos
LblExentosNC.Caption = Format(Exentos, "###,##0.#0")
'Descuento
LblDescuentosNC.Caption = Format(Descu, "###,##0.#0")
'Impuesto
LblImpuestoNC.Caption = Format(TotalIVA, "###,##0.#0")
'Total General
LblTotalGeneralNC.Caption = Format(Exentos + BaseImp + TotalIVA, "###,##0.#0")
End Sub

'Para el DMGRID1
Sub Calcular2()
Dim Cant, PUnit, IIva, Descu As Double
Dim Exentos, BaseImp, SubTot, TotalIVA As Double
Dim Campow As String

Exentos = 0

For i = 1 To DMGrid1.Rows

    IdProd = DMGrid1.ValorCelda(i, 1)
    
    If IdProd <> "" Then
        
        If Val(DMGrid1.ValorCelda(i, 3)) <> 0 Then
            Cant = Val(DMGrid1.ValorCelda(i, 3))
        Else
            Cant = 0
        End If
        
        If Val(DMGrid1.ValorCelda(i, 4)) <> 0 Then
            PUnit = CDbl(DMGrid1.ValorCelda(i, 4))
        Else
            PUnit = 0
        End If
        
        If Val(DMGrid1.ValorCelda(i, 5)) <> 0 Then
            IIva = CDbl(DMGrid1.ValorCelda(i, 5))
        Else
            IIva = 0
        End If
        
        If Val(DMGrid1.ValorCelda(i, 6)) <> 0 Then
            Descu = CDbl(DMGrid1.ValorCelda(i, 6))
        Else
            Descu = 0
        End If
        
        If IsNull(Cant) Then
            Cant = 1
            DMGrid1.ValorCelda(i, 3) = 1
        ElseIf Val(Cant) = 0 Then DMGrid1.ValorCelda(i, 3) = 1: Cant = 1
        End If
        If IsNull(IIva) Then
            IIva = 0
        ElseIf Val(IIva) = 0 Then IIva = 0
        End If
        If IsNull(Descu) Then
            Descu = 0
        ElseIf Val(Descu) = 0 Then Descu = 0
        End If
        
        If Val(IIva) = 0 Then
            Exentos = Exentos + ((Cant * PUnit) - Descu)
        Else
            BaseImp = BaseImp + ((Cant * PUnit) - Descu)
        End If

        SubTot = CDbl(SubTot) + CDbl(DMGrid1.ValorCelda(i, 7))
        TotalIVA = TotalIVA + IIva
        'c = c + ((Cant * PUnit) - Descu) + IIva

    End If
Next i
If Val(BaseImp) = 0 Then BaseImp = 0
If Val(SubTot) = 0 Then SubTot = 0
If Val(Descu) = 0 Then Descu = 0
If Val(Exentos) = 0 Then Exentos = 0
If Val(TotalIVA) = 0 Then TotalIVA = 0


'SubTotal
LblSubTotal.Caption = Format(SubTot, "###,##0.#0")
'Base imponible
LblBaseImponible.Caption = Format(BaseImp, "###,##0.#0")
'Exentos
LblExentos.Caption = Format(Exentos, "###,##0.#0")
'Descuento
LblDescuentos.Caption = Format(Descu, "###,##0.#0")
'Impuesto
LblImpuesto.Caption = Format(TotalIVA, "###,##0.#0")
'Total General
LblTotalGeneral.Caption = Format(Exentos + BaseImp + TotalIVA, "###,##0.#0")
End Sub

Sub Calcular3()
'SubTotal
LblSubTotalG.Caption = Format(CDbl(LblSubTotal) - CDbl(LblSubTotalNC), "###,##0.#0")
'Base imponible
LblBaseImponibleG.Caption = Format(CDbl(LblBaseImponible) - CDbl(LblBaseImponibleNC), "###,##0.#0")
'Exentos
LblExentosG.Caption = Format(CDbl(LblExentos) - CDbl(LblExentosNC), "###,##0.#0")
'Descuento
LblDescuentosG.Caption = Format(CDbl(LblDescuentos) - CDbl(LblDescuentosNC), "###,##0.#0")
'Impuesto
LblImpuestoG.Caption = Format(CDbl(LblImpuesto) - CDbl(LblImpuestoNC), "###,##0.#0")
'Total General
LblTotalGeneralG.Caption = Format(CDbl(LblExentosG) + CDbl(LblBaseImponibleG) + CDbl(LblImpuestoG), "###,##0.#0")
End Sub

Private Function Verificar_Pase(Fila As Integer) As Boolean
Dim CodP As Integer
Dim Cant As Integer
Dim Cant2 As Integer

Cant = 0
Cant2 = 0
CodP = DMGrid1.ValorCelda(Fila, 1)

    For i = 1 To DMGrid2.Rows
        If Val(DMGrid2.ValorCelda(i, 1)) = CodP Then
            Cant = Cant + Val(DMGrid2.ValorCelda(i, 3))
        End If
    Next i
    
    For i = 1 To DMGrid1.Rows
        If Val(DMGrid1.ValorCelda(i, 1)) = CodP Then
            Cant2 = Cant2 + Val(DMGrid1.ValorCelda(i, 3))
        End If
    Next i
    
    If Cant >= Cant2 Then
        MsgBox "No se puede excluir el producto, ya que la cantidad excede a los ya excluidos para este tipo", vbExclamation + vbOKOnly, "Error"
        Verificar_Pase = False
    Else
        Verificar_Pase = True
    End If
End Function

Sub Limpiar()
DMGrid1.Clear
DMGrid2.Clear
    DMGrid1.Rows = 0
    DMGrid1.PaintMGrid
    DMGrid1.Rows = 1
    DMGrid1.RowBackColor 1, vbWhite
    DMGrid2.Rows = 0
    DMGrid2.PaintMGrid
    DMGrid2.Rows = 1
    DMGrid2.RowBackColor 1, vbWhite
    
    LblSubTotal.Caption = ""
    LblBaseImponible.Caption = ""
    LblExentos.Caption = ""
    LblDescuentos.Caption = ""
    LblImpuesto.Caption = ""
    LblTotalGeneral.Caption = ""

    LblSubTotalNC.Caption = ""
    LblBaseImponibleNC.Caption = ""
    LblExentosNC.Caption = ""
    LblDescuentosNC.Caption = ""
    LblImpuestoNC.Caption = ""
    LblTotalGeneralNC.Caption = ""
    
    DMGrid2.PaintMGrid
    DMGrid1.PaintMGrid
End Sub

Private Sub TxtNotaCredito_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 Then: KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then BtnBuscarNC_Click

If Len(TxtNotaCredito.Text) > 7 And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

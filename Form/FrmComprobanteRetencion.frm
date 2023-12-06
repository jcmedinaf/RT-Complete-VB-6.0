VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmComprobanteRetencion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobante de Retención"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   Icon            =   "FrmComprobanteRetencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Width           =   11295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   10200
         TabIndex        =   21
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
         MICON           =   "FrmComprobanteRetencion.frx":1002
         PICN            =   "FrmComprobanteRetencion.frx":101E
         PICH            =   "FrmComprobanteRetencion.frx":11E7
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
         TabIndex        =   22
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
         MICON           =   "FrmComprobanteRetencion.frx":141C
         PICN            =   "FrmComprobanteRetencion.frx":1438
         PICH            =   "FrmComprobanteRetencion.frx":16C7
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
         TabIndex        =   23
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
         MICON           =   "FrmComprobanteRetencion.frx":1B08
         PICN            =   "FrmComprobanteRetencion.frx":1B24
         PICH            =   "FrmComprobanteRetencion.frx":1CB1
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
         Left            =   9000
         TabIndex        =   24
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
         MICON           =   "FrmComprobanteRetencion.frx":1EE6
         PICN            =   "FrmComprobanteRetencion.frx":1F02
         PICH            =   "FrmComprobanteRetencion.frx":21E4
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
         TabIndex        =   25
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
         MICON           =   "FrmComprobanteRetencion.frx":2435
         PICN            =   "FrmComprobanteRetencion.frx":2451
         PICH            =   "FrmComprobanteRetencion.frx":25F5
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
         Left            =   5280
         TabIndex        =   45
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "FrmComprobanteRetencion.frx":2794
         PICN            =   "FrmComprobanteRetencion.frx":27B0
         PICH            =   "FrmComprobanteRetencion.frx":28D5
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
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   1695
      Left            =   8640
      TabIndex        =   10
      Top             =   120
      Width           =   2775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52625409
         CurrentDate     =   40354
      End
      Begin VB.TextBox TxtNoRetencion 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha: "
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprobante de Retención:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Beneficiario"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8415
      Begin VB.TextBox TxtIdProveedor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6720
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtIdBanco 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7560
         TabIndex        =   47
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtNoRif 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox TxtNombreProveedor 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   6495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. R.I.F.:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1290
         Width           =   780
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección o Domicilio Fiscal:"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Proveedor:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Monto Retenido / Concepto"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11295
      Begin ChamaleonButton.ChameleonBtn BtnBanco 
         Height          =   375
         Left            =   10680
         TabIndex        =   46
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   270
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
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmComprobanteRetencion.frx":2B65
         PICN            =   "FrmComprobanteRetencion.frx":2B81
         PICH            =   "FrmComprobanteRetencion.frx":2E19
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame FrameAgregarComprobante 
         BackColor       =   &H00EAEFEF&
         Height          =   3375
         Left            =   1800
         TabIndex        =   26
         Top             =   1080
         Width           =   7815
         Begin VB.TextBox TxtRetencion 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0,00"
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox TxtSustraendo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4920
            TabIndex        =   40
            Text            =   "0,00"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox TxtPorcentaje 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1560
            TabIndex        =   38
            Text            =   "0,00"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox TxtBaseImponible 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4920
            TabIndex        =   36
            Text            =   "0,00"
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox TxtMontoFacturado 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1560
            TabIndex        =   34
            Text            =   "0,00"
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox TxtConcepto 
            Height          =   375
            Left            =   1200
            TabIndex        =   32
            Top             =   840
            Width           =   6495
         End
         Begin VB.TextBox TxtNoControl 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4440
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox TxtNoFactura 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1200
            TabIndex        =   28
            Top             =   360
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarListado 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   2880
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar en listado"
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
            MICON           =   "FrmComprobanteRetencion.frx":30A2
            PICN            =   "FrmComprobanteRetencion.frx":30BE
            PICH            =   "FrmComprobanteRetencion.frx":331E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrarListado 
            Height          =   375
            Left            =   6720
            TabIndex        =   44
            ToolTipText     =   "Cerrar "
            Top             =   2880
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
            MICON           =   "FrmComprobanteRetencion.frx":35B6
            PICN            =   "FrmComprobanteRetencion.frx":35D2
            PICH            =   "FrmComprobanteRetencion.frx":379B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retención:"
            Height          =   195
            Left            =   4035
            TabIndex        =   41
            Top             =   2370
            Width           =   780
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sustraendo:"
            Height          =   195
            Left            =   3960
            TabIndex        =   39
            Top             =   1890
            Width           =   855
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% :"
            Height          =   195
            Left            =   1290
            TabIndex        =   37
            Top             =   1890
            Width           =   210
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Imponible:"
            Height          =   195
            Left            =   3720
            TabIndex        =   35
            Top             =   1410
            Width           =   1125
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Facturado:"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   930
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Control:"
            Height          =   195
            Left            =   3480
            TabIndex        =   29
            Top             =   450
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Factura:"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   450
            Width           =   930
         End
      End
      Begin VB.TextBox TxtBanco 
         Height          =   375
         Left            =   6840
         TabIndex        =   19
         Top             =   270
         Width           =   3735
      End
      Begin VB.TextBox TxtNoCheque 
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox TxtPagado 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   270
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7435
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro. Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nro. Control Factura"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Concepto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Monto Facturado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Base Imponible"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "%"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sustraendo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Retención"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Left            =   6240
         TabIndex        =   5
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Cheque:"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado / Abonado:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FrmComprobanteRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
FrameAgregarComprobante.Visible = True
TxtNoFactura.SetFocus
End Sub

Private Sub BtnAgregarListado_Click()
If TxtNoFactura.Text = "" Then
    Msg = "Debe de Ingresar el número de factura!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtNoFactura.SetFocus
    Exit Sub
End If
If TxtNoControl.Text = "" Then
    Msg = "Debe de Ingresar el número de control de la factura!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtNoControl.SetFocus
    Exit Sub
End If
If TxtConcepto.Text = "" Then
    Msg = "Debe de Ingresar el Concepto de la retención!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtConcepto.SetFocus
    Exit Sub
End If
If TxtMontoFacturado.Text = "" Then
    Msg = "Debe de Ingresar el monto facturado!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtMontoFacturado.SetFocus
    Exit Sub
End If
If TxtBaseImponible.Text = "" Then
    Msg = "Debe de Ingresar el monto de la base imponible!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtBaseImponible.SetFocus
    Exit Sub
End If
If TxtPorcentaje.Text = "" Then
    Msg = "Debe de Ingresar el monto del porcentaje de la retención!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtPorcentaje.SetFocus
    Exit Sub
End If
If TxtSustraendo.Text = "" Then
    Msg = "Debe de Ingresar el monto del sustraendo de la retención!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    TxtSustraendo.SetFocus
    Exit Sub
End If


'********************************************


C = ListView1.ListItems.Count + 1
With ListView1

    .ListItems.Add , , TxtNoFactura.Text
    .ListItems(C).ListSubItems.Add , , TxtNoControl.Text
    .ListItems(C).ListSubItems.Add , , TxtConcepto.Text
    .ListItems(C).ListSubItems.Add , , TxtMontoFacturado.Text
    .ListItems(C).ListSubItems.Add , , TxtBaseImponible.Text
    .ListItems(C).ListSubItems.Add , , TxtPorcentaje.Text
    .ListItems(C).ListSubItems.Add , , TxtSustraendo.Text
    .ListItems(C).ListSubItems.Add , , TxtRetencion.Text
  
    C = C + 1
End With

BLANQUEAR

FrameAgregarComprobante.Visible = False

End Sub

Sub BLANQUEAR()
TxtNoFactura.Text = ""
TxtNoControl.Text = ""
TxtConcepto.Text = ""
TxtMontoFacturado.Text = ""
TxtBaseImponible.Text = ""
TxtPorcentaje.Text = ""
TxtSustraendo.Text = ""
TxtRetencion.Text = ""
End Sub

Private Sub BtnBanco_Click()
Ban = 10
FrmListadoBancos.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnCerrarListado_Click()
FrameAgregarComprobante.Visible = False
End Sub

Private Sub BtnDesHacer_Click()
FrameAgregarComprobante.Visible = False
End Sub

Private Sub BtnEliminar_Click()
If Not ListView1.SelectedItem Is Nothing Then
   
   'Pregunta si lo quiere eliminar
   'If MsgBox("Eliminar ??", vbQuestion + vbYesNo) = vbYes Then
        'Elimina el elemento seleccionado ( SelectedItem.Index )
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
  'End If
End If
End Sub

Private Sub BtnGuardarActualizar_Click()
FrameAgregarComprobante.Visible = False

If TxtPagado.Text = "" Then
   Msg = "Debe de ingresar el monto Pagado / Abonado"
   MsgBox Msg, vbCritical + vbOKOnly, "Error"
   TxtPagado.SetFocus
   Exit Sub
End If
If TxtNoCheque.Text = "" Then
   Msg = "Debe de Ingresar el Numero del Cheque"
   MsgBox Msg, vbCritical + vbOKOnly, "Error"
   TxtNoCheque.SetFocus
   Exit Sub
End If
If TxtIdBanco.Text = "" Then
   Msg = "Debe de Seleccionar el Banco"
   MsgBox Msg, vbCritical + vbOKOnly, "Error"
   BtnBanco.SetFocus
   Exit Sub
End If

If ListView1.ListItems.Count < 0 Then
   Msg = "Debe de ingresar las retetenciones"
   MsgBox Msg, vbCritical + vbOKOnly, "Error"
   BtnAgregar.SetFocus
   Exit Sub
End If



CSql = "Select max(IdMovCajaBanco) + 1 as MaxId From Cobros"
Set RsMaxId = CrearRS(CSql)


If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdMax = RsMaxId.Fields("MaxId").Value
Else
    IdMax = "1"
End If

Dim RsGuardar As New ADODB.Recordset
CSql = "Select * From Movi_BanCaja"
Set RsGuardar = CrearRS(CSql)


RsGuardar.AddNew
RsGuardar.Fields("IdMovCajaBanco").Value = IdMax
RsGuardar.Fields("IdCajaBanco").Value = TxtIdBanco.Text
RsGuardar.Fields("Ingr_Egr").Value = 1
RsGuardar.Fields("N_Comprobante").Value = TxtNoRetencion.Text
RsGuardar.Fields("Monto_Mov").Value = TxtPagado.Text
RsGuardar.Fields("Tipo_Mov").Value = 0
RsGuardar.Fields("Fecha_Transa").Value = DTPicker1.Value
RsGuardar.Fields("Conciliado").Value = 0
RsGuardar.Fields("FechaConciliacion").Value = "01/01/1900"
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Fields("Detalle").Value = ""
RsGuardar.Fields("Beneficiario").Value = TxtNombreProveedor.Text
RsGuardar.Fields("Anulado").Value = 0
RsGuardar.Fields("NoEndosable").Value = 0

RsGuardar.Update

Dim RsRetencion As New ADODB.Recordset

CSql = "Select * From Retencion"
Set RsRetencion = CrearRS(CSql)

For g = 1 To ListView1.ListItems.Count


    CSql = "Select Max(IdRetencion) + 1 as MaxId From Retencion"
    Set RsMaxId = CrearRS(CSql)
    
    
    If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
        IdMaxR = RsMaxId.Fields("MaxId").Value
    Else
        IdMaxR = "1"
    End If

    RsRetencion.AddNew
    
    RsRetencion.Fields("IdRetencion").Value = IdMaxR
    RsRetencion.Fields("IdProveedor").Value = TxtIdProveedor.Text
    RsRetencion.Fields("IdUser").Value = IdUser
    RsRetencion.Fields("N_Factura").Value = ListView1.ListItems.Item(g).Text
    RsRetencion.Fields("N_Control").Value = ListView1.ListItems(g).ListSubItems.Item(1).Text
    RsRetencion.Fields("Concepto").Value = ListView1.ListItems(g).ListSubItems.Item(2).Text
    RsRetencion.Fields("Monto").Value = ListView1.ListItems(g).ListSubItems.Item(3).Text
    RsRetencion.Fields("BaseImponible").Value = ListView1.ListItems(g).ListSubItems.Item(4).Text
    RsRetencion.Fields("Porcentaje").Value = ListView1.ListItems(g).ListSubItems.Item(5).Text
    RsRetencion.Fields("Sustraendo").Value = ListView1.ListItems(g).ListSubItems.Item(6).Text
    RsRetencion.Fields("Retencion").Value = ListView1.ListItems(g).ListSubItems.Item(7).Text
    RsRetencion.Fields("Fecha").Value = DTPicker1.Value
    
    RsRetencion.Update
Next g

Msg = "Retención Guardada Satisfactoriamente!!"
MsgBox Msg, vbInformation + vbOKOnly, "Registro Guardado Satisfactoriamente"

Unload Me


End Sub

Private Sub BtnImprimir_Click()
On Error Resume Next

If TxtIdProveedor.Text <> "" Then
     With CrystalReport1
        .ReportFileName = RutaInformes & "\Comprobante_de_Retenciones.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0

        .SelectionFormula = "{Retenciones.IdRetencion} = " & TxtNoRetencion.Text
        .WindowTitle = "Reporte Comprobante de Retención No. " & TxtNoRetencion.Text
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
Else
    MsgBox "Tiene que seleccionar a un Proveedor", vbOKOnly + vbCritical, "Mensaje de Error"
End If
End Sub

Private Sub Form_Load()
FrameAgregarComprobante.Visible = False
DTPicker1.Value = Now
Dim RsTemp2 As New ADODB.Recordset
Dim RsMaxId As New ADODB.Recordset
CSql = "Select * From Proveedores Where IdProveedor='" & FrmCuentasPorPagar.LstCuentasPagar.SelectedItem.Text & "'"
Set RsTemp2 = CrearRS(CSql)

If RsTemp2.RecordCount > 0 Then
    TxtIdProveedor.Text = RsTemp2.Fields("IdProveedor").Value
    TxtNombreProveedor.Text = RsTemp2.Fields("Nombre").Value
    TxtNoRif.Text = RsTemp2.Fields("RifProveedor").Value
    TxtDireccion.Text = RsTemp2.Fields("Direccion").Value

End If


CSql = "Select max(N_Comprobante) + 1 as MaxId From Movi_BanCaja"
Set RsMaxId = CrearRS(CSql)


If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    NComprobante = RsMaxId.Fields("MaxId").Value
Else
    NComprobante = "1"
End If

TxtNoRetencion.Text = NComprobante

End Sub


Private Sub TxtBaseImponible_Change()
On Error Resume Next
TxtRetencion.Text = (CDbl(TxtBaseImponible.Text) * CDbl(TxtPorcentaje.Text) / 100) - CDbl(TxtSustraendo.Text)
TxtRetencion.Text = Format(TxtRetencion.Text, "#,##0.00")
End Sub

Private Sub TxtBaseImponible_Click()
If TxtMontoFacturado.Text = "0,00" Then TxtMontoFacturado.Text = "": Exit Sub
If TxtMontoFacturado.Text = "" Then TxtMontoFacturado.Text = "0,00": Exit Sub
End Sub

Private Sub TxtMontoFacturado_Click()
If TxtMontoFacturado.Text = "0,00" Then TxtMontoFacturado.Text = "": Exit Sub
If TxtMontoFacturado.Text = "" Then TxtMontoFacturado.Text = "0,00": Exit Sub
End Sub

Private Sub TxtPorcentaje_Change()
On Error Resume Next
TxtRetencion.Text = (CDbl(TxtBaseImponible.Text) * CDbl(TxtPorcentaje.Text) / 100) - CDbl(TxtSustraendo.Text)
TxtRetencion.Text = Format(TxtRetencion.Text, "#,##0.00")
End Sub

Private Sub TxtPorcentaje_Click()
If TxtPorcentaje.Text = "0,00" Then TxtPorcentaje.Text = "": Exit Sub
If TxtPorcentaje.Text = "" Then TxtPorcentaje.Text = "0,00": Exit Sub
End Sub

Private Sub TxtSustraendo_Change()
On Error Resume Next
TxtRetencion.Text = (CDbl(TxtBaseImponible.Text) * CDbl(TxtPorcentaje.Text) / 100) - CDbl(TxtSustraendo.Text)
TxtRetencion.Text = Format(TxtRetencion.Text, "#,##0.00")
End Sub

Private Sub TxtNoFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtNoControl.SetFocus
    Else
        If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub
'///////////////////////////////////Valido TextBox: TxtNoControl//////////////////////////////
Private Sub TxtNoControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtConcepto.SetFocus
    Else
        If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub
'///////////////////////////////////Valido TextBox: TxtConcepto//////////////////////////////
Private Sub TxtConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoFacturado.SetFocus
    Else
        If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ 1234567890.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub
'///////////////////////////////////Valido TextBox: TxtMontoFacturado//////////////////////////////
Private Sub TxtMontoFacturado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBaseImponible.SetFocus
        TxtMontoFacturado.Text = Format(TxtMontoFacturado.Text, "#,##0.00")
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

'///////////////////////////////////Valido TextBox: TxtMontoFacturado//////////////////////////////
Private Sub TxtBaseImponible_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPorcentaje.SetFocus
        TxtBaseImponible.Text = Format(TxtBaseImponible.Text, "#,##0.00")
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub
'///////////////////////////////////Valido TextBox: TxtMontoFacturado//////////////////////////////
Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtSustraendo.SetFocus
        TxtPorcentaje.Text = Format(TxtPorcentaje.Text, "#,##0.00")
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtSustraendo_Click()
If TxtSustraendo.Text = "0,00" Then TxtSustraendo.Text = "": Exit Sub
If TxtSustraendo.Text = "" Then TxtSustraendo.Text = "0,00": Exit Sub
End Sub

'///////////////////////////////////Valido TextBox: TxtMontoFacturado//////////////////////////////
Private Sub TxtSustraendo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BtnAgregarListado.SetFocus
        TxtSustraendo.Text = Format(TxtSustraendo.Text, "#,##0.00")
    Else
        If InStr("1234567890,.", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            KeyAscii = 0
        End If
    End If
End Sub


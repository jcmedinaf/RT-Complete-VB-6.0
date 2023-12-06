VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FacturacionRT 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación"
   ClientHeight    =   9330
   ClientLeft      =   3210
   ClientTop       =   1590
   ClientWidth     =   15225
   Icon            =   "FacturacionRT.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   15225
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Facturación"
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   12015
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Cliente"
         Height          =   1095
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   11775
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1320
            TabIndex        =   45
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   7080
            TabIndex        =   44
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Height          =   435
            Left            =   7080
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   600
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   620
            Width           =   4935
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarCliente 
            Height          =   375
            Left            =   4680
            TabIndex        =   41
            ToolTipText     =   "Agregar Clientes"
            Top             =   230
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   16761087
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FacturacionRT.frx":1002
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   3480
            TabIndex        =   46
            ToolTipText     =   "Buscar Cliente"
            Top             =   230
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
            MICON           =   "FacturacionRT.frx":101E
            PICN            =   "FacturacionRT.frx":103A
            PICH            =   "FacturacionRT.frx":129F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   710
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif.: "
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   6350
            TabIndex        =   48
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   6350
            TabIndex        =   47
            Top             =   330
            Width           =   675
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Paciente"
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   11775
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   9360
            TabIndex        =   32
            Top             =   620
            Width           =   2175
         End
         Begin VB.TextBox Text12 
            Height          =   525
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1010
            Width           =   10455
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1080
            TabIndex        =   30
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   1080
            TabIndex        =   29
            Top             =   620
            Width           =   3135
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   5280
            TabIndex        =   28
            Top             =   620
            Width           =   3135
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarPaciente 
            Height          =   375
            Left            =   4560
            TabIndex        =   33
            ToolTipText     =   "Agregar Pacientes"
            Top             =   230
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar Paciente"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   16761087
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FacturacionRT.frx":1531
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
            Height          =   375
            Left            =   3360
            TabIndex        =   34
            ToolTipText     =   "Buscar Pacientes"
            Top             =   230
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
            MICON           =   "FacturacionRT.frx":154D
            PICN            =   "FacturacionRT.frx":1569
            PICH            =   "FacturacionRT.frx":17CE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   8475
            TabIndex        =   39
            Top             =   710
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   195
            TabIndex        =   38
            Top             =   1175
            Width           =   720
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   195
            TabIndex        =   37
            Top             =   710
            Width           =   765
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   4395
            TabIndex        =   36
            Top             =   710
            Width           =   765
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   195
            TabIndex        =   35
            Top             =   330
            Width           =   540
         End
      End
      Begin VB.Label LblImpuesto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   7440
         TabIndex        =   52
         Top             =   6480
         Width           =   45
      End
      Begin VB.Label LblValorImpuesto 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7440
         TabIndex        =   51
         Top             =   6480
         Width           =   45
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Detalle"
      Height          =   5895
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   15015
      Begin VB.Frame Frame8 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   11640
         TabIndex        =   66
         Top             =   4200
         Width           =   3255
         Begin ChamaleonButton.ChameleonBtn BtnRetenciones 
            Height          =   375
            Left            =   1560
            TabIndex        =   67
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Aplicar Cobros"
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
            MICON           =   "FacturacionRT.frx":1A60
            PICN            =   "FacturacionRT.frx":1A7C
            PICH            =   "FacturacionRT.frx":1EB2
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
            Left            =   120
            TabIndex        =   68
            ToolTipText     =   "Imprimir Factura"
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
            MICON           =   "FacturacionRT.frx":22E2
            PICN            =   "FacturacionRT.frx":22FE
            PICH            =   "FacturacionRT.frx":2423
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
         Caption         =   "SubTotales"
         Height          =   2655
         Left            =   11640
         TabIndex        =   53
         Top             =   1560
         Width           =   3255
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descuentos:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1500
            Width           =   900
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   300
            Left            =   1320
            TabIndex        =   64
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   300
            Left            =   1320
            TabIndex        =   63
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Text7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   300
            Left            =   1320
            TabIndex        =   62
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Text5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   300
            Left            =   1320
            TabIndex        =   61
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   300
            Left            =   1320
            TabIndex        =   60
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Text6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   300
            Left            =   1320
            TabIndex        =   59
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Imponible:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   773
            Width           =   1125
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Exentos:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1133
            Width           =   615
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   413
            Width           =   690
         End
         Begin VB.Label LblImpuestoIVA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IVA: 12%"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   2220
            Width           =   1005
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   6975
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3720
            Top             =   187
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonButton.ChameleonBtn BtnImportar 
            Height          =   375
            Left            =   2400
            TabIndex        =   21
            ToolTipText     =   "Importar Presupuesto"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Importar"
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
            MICON           =   "FacturacionRT.frx":26B3
            PICN            =   "FacturacionRT.frx":26CF
            PICH            =   "FacturacionRT.frx":2B01
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarRenglon 
            Height          =   375
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Agregar Renglon"
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
            MICON           =   "FacturacionRT.frx":2CA9
            PICN            =   "FacturacionRT.frx":2CC5
            PICH            =   "FacturacionRT.frx":2E52
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
            TabIndex        =   23
            ToolTipText     =   "Eliminar Renglon"
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
            MICON           =   "FacturacionRT.frx":3087
            PICN            =   "FacturacionRT.frx":30A3
            PICH            =   "FacturacionRT.frx":3247
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
            Left            =   4200
            Top             =   217
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Renglones:"
            Height          =   195
            Left            =   4680
            TabIndex        =   25
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label LblCantidadRenglon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6240
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   1215
         Left            =   11640
         TabIndex        =   16
         Top             =   240
         Width           =   3255
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   960
            TabIndex        =   17
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el No. de la Factura"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnListarFac 
            Height          =   375
            Left            =   960
            TabIndex        =   18
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Lista de Facturas"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FacturacionRT.frx":3687
            PICN            =   "FacturacionRT.frx":36A3
            PICH            =   "FacturacionRT.frx":3AD3
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
            Caption         =   "Nº Factura:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   7200
         TabIndex        =   10
         Top             =   5040
         Width           =   7695
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   4680
            Top             =   217
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6600
            TabIndex        =   11
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
            MICON           =   "FacturacionRT.frx":3D6F
            PICN            =   "FacturacionRT.frx":3D8B
            PICH            =   "FacturacionRT.frx":3F54
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
            TabIndex        =   12
            ToolTipText     =   "Guardar / Actualizar la Factura"
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
            MICON           =   "FacturacionRT.frx":4189
            PICN            =   "FacturacionRT.frx":41A5
            PICH            =   "FacturacionRT.frx":4434
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
            TabIndex        =   13
            ToolTipText     =   "Agregar Nueva Factura"
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
            MICON           =   "FacturacionRT.frx":4875
            PICN            =   "FacturacionRT.frx":4891
            PICH            =   "FacturacionRT.frx":4A1E
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
            Left            =   5400
            TabIndex        =   14
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
            MICON           =   "FacturacionRT.frx":4C53
            PICN            =   "FacturacionRT.frx":4C6F
            PICH            =   "FacturacionRT.frx":4F51
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnReciboCobros 
            Height          =   375
            Left            =   2640
            TabIndex        =   15
            ToolTipText     =   "Imprimir Recibo de Cobro"
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Imprimir Recibo"
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
            MICON           =   "FacturacionRT.frx":51A2
            PICN            =   "FacturacionRT.frx":51BE
            PICH            =   "FacturacionRT.frx":52E3
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
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8281
         Object.Width           =   11385
         Object.Height          =   4665
         Editable        =   -1  'True
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3255
      Index           =   1
      Left            =   12240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FacturacionRT.frx":5573
         Left            =   120
         List            =   "FacturacionRT.frx":5575
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Factura Impresa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47841281
         CurrentDate     =   39932
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         TabIndex        =   8
         Top             =   120
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   525
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FacturacionRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim BD61 As Recordset 'Tabla pago
Dim BD62 As Recordset 'Tabla paciente
Dim BD63 As Recordset 'Tabla cliente
Dim BD64 As Recordset 'renglones de factura
Dim BD65 As Recordset 'c_cobrar
Dim BD66 As Recordset 'C_cobrar
Dim BD68 As Recordset 'C_cobrar para guardar
Dim BD69 As Recordset 'presupuesto2
Dim BD70 As Recordset 'articulos
Dim BD71 As Recordset 'presupuesto
Dim IdProd
Dim SumProduc1 As Double
Dim SumProduc2 As Double
Dim NPresup As Integer
Dim CondFact As Integer ' indica si la factura va a ser agregada o actualizada en el registro
Public IdCliente
Public a5, A6

Sub GuardaRenglon(aa1, aa9)
    For i = 1 To DMGrid1.Rows
        b1 = DMGrid1.ValorCelda(i, 1)
        b2 = DMGrid1.ValorCelda(i, 4)
        b3 = DMGrid1.ValorCelda(i, 3)
        b4 = DMGrid1.ValorCelda(i, 5)
        b5 = DMGrid1.ValorCelda(i, 6)
        b6 = DMGrid1.ValorCelda(i, 7)
        If IsNull(b1) Or Trim(b1) = "" Then GoTo r
        If IsNull(b5) Or Trim(b5) = "" Then b5 = 0
        If IsNull(b2) Or Trim(b2) = "" Then b2 = 0
        If IsNull(b3) Or Trim(b3) = "" Then b3 = 0
        If IsNull(b4) Or Trim(b4) = "" Then b4 = 0
        If IsNull(b6) Or Trim(b6) = "" Then b6 = 0
'        Call QuitarCaracter(b2)
'        b2 = Carac
'        Call Quitar(b2)
'        b2 = Carac
        
'        CSql = aa1 & "," & b1 & "," & b2 & "," & b3 & "," & b4 & "," & b5 & "," & aa9 & ")"
'        CSql = "insert into reng_cobrar (n_factura,cod_producto,precio,cantidad,iva,descuento,idusuario) values(" & CSql
'        Set RsGuardaRenglon = CrearRS(CSql)
'        RsGuardaRenglon.Close


        Dim RsGuardaRenglon As New ADODB.Recordset
        CSql = "Select * From Reng_Cobrar"
        Set RsGuardaRenglon = CrearRS(CSql)

        RsGuardaRenglon.AddNew
        RsGuardaRenglon.Fields("N_Factura").Value = aa1
        RsGuardaRenglon.Fields("Cod_Producto").Value = b1
        RsGuardaRenglon.Fields("Precio").Value = CDbl(b2)
        RsGuardaRenglon.Fields("Cantidad").Value = b3
        RsGuardaRenglon.Fields("Iva").Value = CDbl(b4)
        RsGuardaRenglon.Fields("Descuento").Value = CDbl(b5)
        RsGuardaRenglon.Fields("IdUsuario").Value = IdUser
        RsGuardaRenglon.Fields("Fecha").Value = Format(DTPicker1.Value, "DD/MM/YYYY")
        RsGuardaRenglon.Fields("SubTotal").Value = CDbl(b6)
        RsGuardaRenglon.Fields("N_NC").Value = "0"
        RsGuardaRenglon.Fields("N_FA").Value = "0"
        RsGuardaRenglon.Fields("C_NC").Value = "0"
        RsGuardaRenglon.Update
        RsGuardaRenglon.Close

r:
    Next i
End Sub

Sub Guardar()
Dim TPago As String
Dim Band As Boolean


'Carac = 1
Band = False

If IdCliente = 0 Then
    MsgBox "No hay cliente seleccionado", vbExclamation + vbOKOnly, "Error"
    Exit Sub
ElseIf Trim(IdPac1) = "" Then
    MsgBox "No hay Paciente seleccionado", vbExclamation + vbOKOnly, "Error"
    Exit Sub
ElseIf Combo1.ListIndex = -1 Then
    MsgBox "Debe seleccionar una forma de pago!", vbExclamation + vbOKOnly, "Error"
    Combo1.SetFocus
    Exit Sub
ElseIf Val(LblCantidadRenglon) = 0 Then
    MsgBox "No hay renglones agregados!", vbExclamation + vbOKOnly, "Error"
    Exit Sub
Else
    For i = 0 To DMGrid1.Rows
        If Not IsNull(DMGrid1.ValorCelda(i, 1)) Then
            If Val(DMGrid1.ValorCelda(i, 1)) <> 0 Then Band = True: Exit For
        End If
    Next i
End If
    
If Band = False Then MsgBox "No hay productos agregados", vbExclamation + vbOKOnly, "Error": Exit Sub

If Combo1.ItemData(Combo1.ListIndex) = 1 Then
    TPago = "0"
Else
    TPago = CDbl(Text8.Caption)
End If
Select Case CondFact

Case Is = 1 'actualiza
    'actualiza la tabla c_cobrar con los datos modificados
        
    a1 = Val(Label12.Caption)
    a2 = Combo1.ListIndex
    Call QuitarCaracter(Text8.Caption)
    a3 = Carac
    Call Quitar(a3)
    a3 = Carac
    a5 = Val(IdPac1)
    A6 = Val(IdCliente)
    a7 = Format(DTPicker1.Value, "DD/MM/YYYY")
    If Check1.Value = 0 Then a8 = 0 Else a8 = 1
    a9 = Val(IdUser)
    a10 = CDbl(Text6.Caption)
    a11 = CDbl(Text7.Caption)
    a12 = CDbl(Text5.Caption)
    a13 = CDbl(Label23.Caption)
    
    Dim RsUpdateCobrar As New ADODB.Recordset
    CSql = "Select * From c_cobrar where n_factura = '" & a1 & "'"
    Set RsUpdateCobrar = CrearRS(CSql)
    
    RsUpdateCobrar.Fields("Forma_Pago").Value = a2
    RsUpdateCobrar.Fields("IdPaciente").Value = a5
    RsUpdateCobrar.Fields("IdCliente").Value = A6
    RsUpdateCobrar.Fields("Fecha").Value = a7
    RsUpdateCobrar.Fields("IdUsuario").Value = a9
    RsUpdateCobrar.Fields("SubTotal").Value = CDbl(a10)
    RsUpdateCobrar.Fields("TasaImpuesto").Value = IVA ' agregado
    RsUpdateCobrar.Fields("BaseImponible").Value = CDbl(Text15.Caption)  ' agregado
    RsUpdateCobrar.Fields("Impuesto").Value = CDbl(a11)
    RsUpdateCobrar.Fields("Exento").Value = CDbl(a12)
    RsUpdateCobrar.Fields("Monto").Value = CDbl(a3)
    RsUpdateCobrar.Fields("PorCobrar").Value = CDbl(TPago)
    RsUpdateCobrar.Fields("Anulada").Value = "0"
    RsUpdateCobrar.Fields("N_NC").Value = "0"
    RsUpdateCobrar.Fields("N_FA").Value = "0"
    RsUpdateCobrar.Fields("N_Presup").Value = NPresup
    RsUpdateCobrar.Fields("C_NC").Value = "0"
    RsUpdateCobrar.Fields("Descuentos").Value = a13
    RsUpdateCobrar.Fields("Retenciones").Value = 0
    RsUpdateCobrar.Fields("TimbresFiscales").Value = 0
    RsUpdateCobrar.Update
    RsUpdateCobrar.Close
    
    Dim RsBorrarRenglonCobrar As New ADODB.Recordset
    'elimina todos los renglones de la tabla reng_cobrar
    CSql = "delete from reng_cobrar where n_factura = " & a1
    Set RsBorrarRenglonCobrar = CrearRS(CSql)
    Call Enviar_Bitacora(IdUser, "FACTURACION", "GUARDAR-ACTUALIZAR", "Se MODIFICARON-ACTUALIZARON los datos de la factura Nro. " & a1)
    Call GuardaRenglon(a1, a9)
    
    MsgBox "La factura Nro. " & a1 & " fue actualizada satisfactoriamente!", vbInformation + vbOKOnly, "Operación Exitosa"

Case Is = 0 'Agrega
    Call Nfac

    'graba datos en tabla de c_cobrar
    a1 = Val(Label12.Caption)
    N_Factur = a1
    a2 = Combo1.ListIndex
    'Call QuitarCaracter(Text8.Caption)
    'a3 = Carac
    'Call Quitar(a3)
    a3 = CDbl(Text8.Caption)
    a4 = 1  'factura = 1 N.Credito =2
    a5 = Val(IdPac1)
    A6 = Val(IdCliente)
    a7 = Format(DTPicker1.Value, "DD/MM/YYYY")
    If Check1.Value = 0 Then a8 = 0 Else a8 = 1
    a9 = Val(IdUser)
    a10 = CDbl(Text6.Caption)
    a11 = CDbl(Text7.Caption)
    a12 = CDbl(Text5.Caption)
    a13 = CDbl(Label23.Caption)
    
    'Call Quitar(a3)
   ' a3 = Carac
    Dim RsGrabaCuentasCobrar As New ADODB.Recordset
    'CSql = a1 & "," & a2 & "," & CDec(a3) & "," & a4 & "," & a5 & "," & a6 & ",'" & a7 & "'," & a9 & "," & a8 & "," & a10 & "," & a11 & "," & a12 & ")"
    'CSql = "insert into c_cobrar (N_factura, forma_pago, monto, tipo, idpaciente, idcliente, fecha, idusuario, impresa, subtotal, impuesto, exento) values(" & CSql
    'Set RsGrabaCuentasCobrar = CrearRS(CSql)
    Dim Comentario
    Msg = "Deseas Agregarle un comentario a la Factura!!"
    mensaje = MsgBox(Msg, vbYesNo + vbInformation, "Agregar Comentario")
    
    If mensaje = vbYes Then
        Comentario = InputBox("Ingrese el Comentario Adicional de la Factura", "Comentario Adicional")
    Else
        Comentario = ""
    End If
    
    
    CSql = "Select * From C_Cobrar"
    Set RsGrabaCuentasCobrar = CrearRS(CSql)
    
    RsGrabaCuentasCobrar.AddNew
    RsGrabaCuentasCobrar.Fields("N_Factura").Value = a1
    RsGrabaCuentasCobrar.Fields("Forma_Pago").Value = a2
    RsGrabaCuentasCobrar.Fields("Tipo").Value = CDec(a4)
    RsGrabaCuentasCobrar.Fields("IdPaciente").Value = a5
    RsGrabaCuentasCobrar.Fields("IdCliente").Value = A6
    RsGrabaCuentasCobrar.Fields("Fecha").Value = a7
    RsGrabaCuentasCobrar.Fields("IdUsuario").Value = a9
    RsGrabaCuentasCobrar.Fields("Impresa").Value = a8
    RsGrabaCuentasCobrar.Fields("SubTotal").Value = CDbl(a10)
    RsGrabaCuentasCobrar.Fields("TasaImpuesto").Value = IVA ' agregado
    RsGrabaCuentasCobrar.Fields("BaseImponible").Value = CDbl(Text15.Caption)  ' agregado
    RsGrabaCuentasCobrar.Fields("Impuesto").Value = CDbl(a11)
    RsGrabaCuentasCobrar.Fields("Exento").Value = CDbl(a12)
    RsGrabaCuentasCobrar.Fields("Monto").Value = CDbl(a3)
    RsGrabaCuentasCobrar.Fields("PorCobrar").Value = CDbl(TPago)
    RsGrabaCuentasCobrar.Fields("Anulada").Value = "0"
    RsGrabaCuentasCobrar.Fields("N_NC").Value = "0"
    RsGrabaCuentasCobrar.Fields("N_FA").Value = "0"
    RsGrabaCuentasCobrar.Fields("N_Presup").Value = Val(NPresup)
    RsGrabaCuentasCobrar.Fields("C_NC").Value = "0"
    RsGrabaCuentasCobrar.Fields("Comentario").Value = Comentario
    RsGrabaCuentasCobrar.Fields("Descuentos").Value = a13
    RsGrabaCuentasCobrar.Fields("Retenciones").Value = 0
    RsGrabaCuentasCobrar.Fields("TimbresFiscales").Value = 0
    RsGrabaCuentasCobrar.Update
    
    Call Enviar_Bitacora(IdUser, "FACTURACION", "GUARDAR-NUEVO", "Se creo la factura Nro. " & a1)
    'graba los renglones de factura
    
    Call GuardaRenglon(a1, a9)
    Dim RsUpdateNoFact As New ADODB.Recordset
    'actualiza el campo administrativo que lleva el correlativo de la factura
    CSql = "update dat_admin set u_factura = " & a1
    Set RsUpdateNoFact = CrearRS(CSql)
    CondFact = 1
    MsgBox "La factura Nro. " & a1 & " fue creada satisfactoriamente!", vbInformation + vbOKOnly, "Operación Exitosa"

End Select


Msg = "Deseas imprimir la factura No. " & N_Factur
mensaje = MsgBox(Msg, vbYesNo + vbInformation, "Impresión de Factura")

If mensaje = vbYes Then
    BtnImprimir_Click
    Limpiar_Campos
Else
    Limpiar_Campos
End If

End Sub

Sub Carga_Cliente()
Dim RsCargarCliente As New ADODB.Recordset
'datos de cliente
CSql = "Select * From Cliente Where IdCliente = '" & IdCliente & "'"
Set RsCargarCliente = CrearRS(CSql)

If Not RsCargarCliente.EOF Then
    Text1.Text = RsCargarCliente.Fields("Rif").Value
    Text2.Text = RsCargarCliente.Fields("Razon").Value
    Text3.Text = RsCargarCliente.Fields("DireccionC").Value
    Text4.Text = RsCargarCliente.Fields("Telefono").Value
    
    RsCargarCliente.Close
Else
    RsCargarCliente.Close
End If

End Sub

Sub calcular()
Dim Cant, PUnit, IIva, Descu As Double
Dim Exentos, BaseImp, SubTot, TotalIVA As Double
Dim Campow As String

'If IdProd = "" Then Exit Sub

Exentos = 0

'For i = 1 To DMGrid1.Rows
'    dmgrid1.
'Next i
' Inicia el ciclo desde la Fila uno hasta el Final de la tabla
For i = 1 To DMGrid1.Rows

    IdProd = DMGrid1.ValorCelda(i, 1)
    If Not IsNull(DMGrid1.ValorCelda(i, 4)) Then
        Campow = DMGrid1.ValorCelda(i, 4)
    Else
        Campow = ""
    End If
    
    If IdProd <> "" And Campow <> "" Then
        Dim RsAplicaIva As New ADODB.Recordset
        Dim AplicaIva
    
        CSql = "Select * From Productos Where IdProducto ='" & IdProd & "'"
        Set RsAplicaIva = CrearRS(CSql)
        AplicaIva = RsAplicaIva.Fields("Impuesto").Value
        RsAplicaIva.Close
        
        If Val(DMGrid1.ValorCelda(i, 3)) <> 0 Then
            Cant = CDbl(DMGrid1.ValorCelda(i, 3))
        Else
            Cant = 0
        End If
        'QuitarCaracter (a)
        'a = CArac
        
        If Val(DMGrid1.ValorCelda(i, 4)) <> 0 Then
            PUnit = CDbl(DMGrid1.ValorCelda(i, 4))
        Else
            PUnit = 0
        End If
        'QuitarCaracter (b)
        'b = CArac
      
        If Val(DMGrid1.ValorCelda(i, 5)) <> 0 Then
            IIva = CDbl(DMGrid1.ValorCelda(i, 5))
        Else
            IIva = 0
        End If
        
        'QuitarCaracter (f)
        'f = CArac
        
        If Val(DMGrid1.ValorCelda(i, 6)) <> 0 Then
            Descu = CDbl(DMGrid1.ValorCelda(i, 6))
        Else
            Descu = 0
        End If
        'QuitarCaracter (d)
        'd = CArac
        
        If IsNull(Cant) Then
            Cant = 0
        ElseIf Val(Cant) = 0 Then Cant = 0
        End If
        If IsNull(PUnit) Then
            PUnit = 0
        ElseIf Val(PUnit) = 0 Then PUnit = 0
        End If
        If IsNull(IIva) Then
            IIva = 0
        ElseIf Val(IIva) = 0 Then IIva = 0
        End If
        If IsNull(Descu) Then
            Descu = 0
        ElseIf Val(Descu) = 0 Then Descu = 0
        End If
        
        If AplicaIva = True Then
            Exentos = Exentos + ((Cant * PUnit) - Descu)
        Else
            BaseImp = BaseImp + ((Cant * PUnit) - Descu) + IIva
        End If
        'd = 0
        SubTot = CDbl(SubTot) + CDbl(DMGrid1.ValorCelda(i, 7))
        TotalIVA = TotalIVA + IIva
        C = C + ((Cant * PUnit) - Descu) + IIva
        
    '    If AplicaIva = True Then
    '        e = c * (IVA / 100)
    '    End If
    End If
Next i
If Val(BaseImp) = 0 Then BaseImp = 0
If Val(SubTot) = 0 Then SubTot = 0
If Val(Descu) = 0 Then Descu = 0
If Val(Exentos) = 0 Then Exentos = 0
If Val(TotalIVA) = 0 Then TotalIVA = 0


'SubTotal
Text6.Caption = Format(SubTot, "###,##0.#0")
'Base imponible
Text5.Caption = Format(BaseImp, "###,##0.#0")
'Exentos
Text15.Caption = Format(Exentos, "###,##0.#0")
'Descuento
Label23.Caption = Format(Descu, "###,##0.#0")
'Impuesto
Text7.Caption = Format(TotalIVA, "###,##0.#0")
'Total General
Text8.Caption = Format(Exentos + BaseImp + TotalIVA, "###,##0.#0")
End Sub

Sub imprime()
On Error GoTo Wrr
''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\FacturaN.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "DSN=CrReporte;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Factura.N_Factura} = " & N_Factur
    .WindowTitle = "Reporte Factura No. " & Label12.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
Exit Sub
Wrr:
MsgBox Err.Number & " : " & Err.Description & " / " & Err.Source
End Sub

Private Sub BtnAgregar_Click()
IO = 1
Limpiar_Campos

End Sub

Sub Limpiar_Campos()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text9.Text = ""
 Text10.Text = ""
 Text11.Text = ""
 Text12.Text = ""
 Text13.Text = ""
 Label12.Caption = ""
 Text5.Caption = ""
 Text6.Caption = ""
 Text15.Caption = ""
 Text7.Caption = ""
 Text8.Caption = ""
 Label23.Caption = ""
 Combo1.ListIndex = -1
 DTPicker1.Value = Now
 DMGrid1.Clear
 DMGrid1.Rows = 1
 DMGrid1.RowBackColor 1, vbWhite
 NPresup = 0
 IdCliente = 0
 IdPac1 = ""
 IdProd = 0
 SumProduc1 = 0
 SumProduc2 = 0
 CondFact = 0
 LblCantidadRenglon.Caption = DMGrid1.Rows
 Check1.Value = 0
 DMGrid1.PaintMGrid
 BtnImportar.Enabled = True
End Sub

Private Sub BtnAgregarCliente_Click()
FrmDatosClientes.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAgregarPaciente_Click()
FrmNuevoPaciente.Show
End Sub

Private Sub BtnAgregarRenglon_Click()
If Check1.Value = 0 Then
    DMGrid1.Rows = DMGrid1.Rows + 1
    Call DMGrid1.PaintMGrid
Else
    Msg = "Ya esta factura fue impresa no puede ser modificada"
    MsgBox Msg
End If
LblCantidadRenglon = DMGrid1.Rows
End Sub

Private Sub BtnBorrarRenglon_Click()
If Check1.Value = 0 Then
    DMGrid1.RowDelete (DMGrid1.Row)
    If Val(LblCantidadRenglon.Caption) > 1 Then LblCantidadRenglon.Caption = Val(LblCantidadRenglon.Caption) - 1
    DMGrid1.PaintMGrid
    Call calcular
Else
    Msg = "Ya esta factura fue impresa no puede ser modificada"
    MsgBox Msg
End If
End Sub

Private Sub BtnBuscar_Click()
ModulO = 1
FrmListadoClientes.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnFacturar_Click()

End Sub

Private Sub BtnGuardarActualizar_Click()
If Check1.Value = 0 Then
    'boton guardar
    Guardar
Else
    Msg = "Esta Factura ya fue impresa no puede ser modificada"
    MsgBox Msg, vbInformation, "Mensaje"
End If
End Sub

Private Sub BtnImportar_Click()
Dim RsImportar As New ADODB.Recordset
Msg = "Indique el Presupuesto del paciente "
pre = Trim(InputBox(Msg, "Presupuesto del paciente", "Nro Presupuesto"))
If pre = "" Then Exit Sub

If Not IsNumeric(pre) Then: MsgBox " Debe ingresar solo números!!", vbCritical + vbOKOnly, "ERROR": Exit Sub
If Val(pre) = 0 Then Exit Sub

CSql = "select * from Presupuesto where npresupuesto = " & Val(pre)
Set RsImportar = CrearRS(CSql)

If Not (RsImportar.EOF) Then

    NPresup = pre
    IdCliente = RsImportar.Fields("idcliente").Value
    IdPac1 = RsImportar.Fields("idpaciente").Value
    idproducto = RsImportar.Fields("idproducto").Value
    Cant = RsImportar.Fields("cantidaD").Value
    Monto = RsImportar.Fields("monto").Value
    RsImportar.Close
    DMGrid1.Rows = 1
    DMGrid1.Clear


    Dim RsProductos As New ADODB.Recordset
    DMGrid1.ValorCelda(1, 1) = idproducto
    CSql = "select descripcion, impuesto from productos where idproducto = " & idproducto
    Set RsProductos = CrearRS(CSql)

    DMGrid1.ValorCelda(1, 2) = RsProductos.Fields("descripcion").Value
    If RsProductos.Fields("impuesto") Then DMGrid1.ValorCelda(f, 5) = IVA Else DMGrid1.ValorCelda(f, 5) = 0

    RsProductos.Close
    DMGrid1.ValorCelda(1, 3) = Cant
    DMGrid1.ValorCelda(1, 4) = Monto
    DMGrid1.ValorCelda(1, 6) = 0
    DMGrid1.ValorCelda(1, 7) = Cant * Monto

    Call DMGrid1.PaintMGrid
    Dim RsPaciente As New ADODB.Recordset
    CSql = "select cedulap,nombrep,apellidop,telefono,direccionp from paciente where idpaciente = " & IdPac1
    Set RsPaciente = CrearRS(CSql)

    If Not (RsPaciente.EOF) Then
        Text9.Text = RsPaciente.Fields("cedulap").Value
        Text11.Text = RsPaciente.Fields("nombrep").Value
        Text10.Text = RsPaciente.Fields("apellidop").Value
        Text13.Text = RsPaciente.Fields("telefono").Value
        Text12.Text = RsPaciente.Fields("Direccionp").Value
    
        Dim RsCliente As New ADODB.Recordset
        CSql1 = "select Razon,rif,direccionc,contacto,email,telefono from Cliente where idcliente = " & IdCliente
        Set RsCliente = CrearRS(CSql1)
    
        If Not (RsCliente.EOF) Then
            Text1.Text = RsCliente.Fields("rif").Value
            Text2.Text = RsCliente.Fields("Razon").Value
            Text3.Text = RsCliente.Fields("Direccionc").Value
            Text4.Text = RsCliente.Fields("Telefono").Value
            'Text5.Text = RsCliente.Fields("Email").Value
        'Call Nfac
        End If
          
        RsCliente.Close
        RsPaciente.Close
        
        Call calcular
        N_Factur = ""
    End If
    Me.Caption = "Facturación - Presupuesto No: " & NPresup
Else
    Me.Caption = "Facturación"
    Msg = "No existe ese No. de Presupuesto!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

End Sub

Private Sub BtnImprimir_Click()
If Text1.Text <> "" And Text9.Text <> "" Then
    If Check1.Value = 0 Then
        Dim RsActualizarImpresa As New ADODB.Recordset
        'colcar aqui el codigo que levante el informe de la factura
        imprime
        CSql = "Update c_cobrar set impresa = 1 WHERE N_Factura = " & N_Factur
        Set RsActualizarImpresa = CrearRS(CSql)
        
        Check1.Value = 1
        
        Call Enviar_Bitacora(IdUser, "FACTURACION", "IMPRIMIR", "Se imprimio la factura Nro. " & N_Factur)
        
        Dim RsFormaPago As New ADODB.Recordset
        CSql = "Select Forma_Pago From C_Cobrar Where N_Factura = " & N_Factur
        Set RsFormaPago = CrearRS(CSql)
        
        If RsFormaPago.Fields("Forma_Pago").Value = 0 Then
            BtnRetenciones.Enabled = True
            BtnRetenciones_Click
        End If
    Else
        Msg = "Ya esta factura fue impresa desea imprimir una copia ?"
        d = MsgBox(Msg, vbYesNo + vbInformation, "Factura Impresa")
        If d = 6 Then
            imprime
        End If
    End If
End If
End Sub

Sub Grid2()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 7
DMGrid1.Rows = 1
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1
DMGrid1.DColumnas(6).Alignment = 1
DMGrid1.DColumnas(7).Alignment = 1
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(4).Locked = False
DMGrid1.DColumnas(5).Locked = True
DMGrid1.DColumnas(7).Locked = True
DMGrid1.DColumnas(4).IsNumber = True
DMGrid1.DColumnas(5).IsNumber = True
DMGrid1.DColumnas(6).IsNumber = True
DMGrid1.DColumnas(7).IsNumber = True
DMGrid1.DColumnas(1).Width = 600
DMGrid1.DColumnas(2).Width = 5300
DMGrid1.DColumnas(3).Width = 800
DMGrid1.DColumnas(4).Width = 1300
DMGrid1.DColumnas(5).Width = 1100
DMGrid1.DColumnas(6).Width = 1100
DMGrid1.DColumnas(7).Width = 1100
DMGrid1.DColumnas(1).Caption = "Codigo"
DMGrid1.DColumnas(2).Caption = "Descripcion"
DMGrid1.DColumnas(3).Caption = "Cantidad"
DMGrid1.DColumnas(4).Caption = "Precio/Unitario"
DMGrid1.DColumnas(5).Caption = "Iva"
DMGrid1.DColumnas(6).Caption = "Descuento"
DMGrid1.DColumnas(7).Caption = "Sub-Total"

End Sub

Sub Nfac()
Dim RsDatAdmin As New ADODB.Recordset
CSql = "select * from dat_admin"
Set RsDatAdmin = CrearRS(CSql)

If Not (RsDatAdmin.EOF) Then
    N_fac = RsDatAdmin.Fields("u_Factura").Value + Val(1)
End If
RsDatAdmin.Close

Label12.Caption = Format(N_fac, "00000000#")

End Sub

Private Sub BtnListarFac_Click()
FacturacionRTLista.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnReciboCobros_Click()
If Label12.Caption = "" Then
    MsgBox "Seleccione una factura para realizarle su recibo de cobro!!", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If
''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\Recibo_Cobro.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{RecibosDeCobros.N_Factura} = " & N_Factur
    .WindowTitle = "Reporte Recibos de Cobro Factura No. " & N_Factur
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
End Sub

Private Sub BtnRetenciones_Click()
If Check1.Value = 0 Then
    Msg = "Debe de guardar y haber impreso esta factura para realizar el cobro"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    BtnReciboCobros.Enabled = False
    Exit Sub
End If

Dim RsTempCobrado As New ADODB.Recordset
CSql = "Select Sum(Monto) as MontoCobrado From Cobros Where N_Factura='" & Val(Label12.Caption) & "'"
Set RsTempCobrado = CrearRS(CSql)

If IsNull(RsTempCobrado.Fields("MontoCobrado").Value) Then
    With FrmAsientoCobros
        .Text3.Text = Text8.Caption
        .Text2.Text = Text2.Text
        .Text1.Text = Label12.Caption
    
    End With
    BtnReciboCobros.Enabled = False
    FrmAsientoCobros.Show
    Exit Sub
End If


If RsTempCobrado.Fields("MontoCobrado").Value <> CDbl(Text8.Caption) Then
    With FrmAsientoCobros
        .Text3.Text = Text8.Caption
        .Text2.Text = Text2.Text
        .Text1.Text = Label12.Caption
    
    End With
    BtnReciboCobros.Enabled = False
    FrmAsientoCobros.Show
Else
    Msg = "La Factura Nº. " & Val(Label12.Caption) & " ya se le realizaron todos los Cobros"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    BtnReciboCobros.Enabled = True
    Exit Sub
End If
End Sub

Private Sub ChameleonBtn2_Click()
Tipo = "Facturacion"
FrmListadoPaciente.Show vbModal, FrmPrincipal
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtBuscaru.SetFocus
        Case vbKeyLeft
            Text13.SetFocus
        Case 38
            Combo1.SetFocus
        Case 40
            TxtBuscaru.SetFocus
    End Select
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            If Check1.Enabled = True Then
                Check1.SetFocus
            Else
                TxtBuscaru.SetFocus
            End If
        Case 37
            Text12.SetFocus
        Case 38
            DTPicker1.SetFocus
        Case 40
            If Check1.Enabled = True Then
                Check1.SetFocus
            Else
                TxtBuscaru.SetFocus
            End If
    End Select
End If
End Sub

Private Sub DMGrid1_AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
Dim a, b, C, d, FProduct As Double
Dim P1, P2, P3 As Boolean

For i = 1 To DMGrid1.Row
    If Val(DMGrid1.ValorCelda(i, 1)) = 0 Then
        s = i
        Exit For
    Else
        s = DMGrid1.Row
    End If
Next i

If lCol = 1 Then
    f = s
    d = Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
    
    If d = 0 Or IsNull(d) Then DMGrid1.RowClear (DMGrid1.Row): DMGrid1.RowBackColor DMGrid1.Row, RGB(255, 255, 255): Call calcular: Exit Sub
    
    Dim RsProductos As New ADODB.Recordset
    CSql = "Select * from productos where idproducto = " & d
    Set RsProductos = CrearRS(CSql)
    
    ' Crea la busqueda para saber que precio colocar
    CSql = "Select * From Dat_Admin"
    Set RsCargarConfig = CrearRS(CSql)
    
    P1 = RsCargarConfig.Fields("PrecioUnitario1").Value
    P2 = RsCargarConfig.Fields("PrecioUnitario2").Value
    P3 = RsCargarConfig.Fields("PrecioUnitario3").Value
    ValorIva = RsCargarConfig.Fields("Iva1").Value
    RsCargarConfig.Close
    'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
    
    If RsProductos.EOF Then RsProductos.Close: DMGrid1.RowClear (DMGrid1.Row): MsgBox "El código del producto no existe en la Base de Datos!", vbExclamation + vbOKOnly, "Código Inexistente": Exit Sub
    DMGrid1.ValorCelda(f, 2) = RsProductos.Fields("Descripcion").Value
    'DMGrid1.ValorCelda(f, 4) = RsProductos.Fields("Precio").Value
    
    If P1 = True Then
        DMGrid1.ValorCelda(f, 4) = RsProductos.Fields("PrecioUnitario1").Value
    ElseIf P2 = True Then
        DMGrid1.ValorCelda(f, 4) = RsProductos.Fields("PrecioUnitario2").Value
    ElseIf P3 = True Then
        DMGrid1.ValorCelda(f, 4) = RsProductos.Fields("PrecioUnitario3").Value
    Else
        DMGrid1.ValorCelda(f, 4) = RsProductos.Fields("CostoActual").Value
    End If
    
    DMGrid1.ValorCelda(f, 3) = "1"
    DMGrid1.ValorCelda(f, 5) = "0.00"
    DMGrid1.ValorCelda(f, 6) = "0.00"
    DMGrid1.ValorCelda(f, 7) = "0.00"
    'If RsProductos.Fields("iva") Then DMGrid1.ValorCelda(f, 5) = Format(IVA, "#,##0.00") Else DMGrid1.ValorCelda(f, 5) = Format(0, "#,##0.00")
    RsProductos.Close
End If


   CodProd = DMGrid1.ValorCelda(DMGrid1.Row, 1)
    
    If Val(CodProd) = 0 Or IsNull(CodProd) Then DMGrid1.RowClear (DMGrid1.Row): DMGrid1.RowBackColor DMGrid1.Row, RGB(255, 255, 255): Call calcular: Exit Sub
    
    CSql = "Select * From Productos where IdProducto='" & CodProd & "'"
    Set Rsiva = CrearRS(CSql)

    CalculaIva = Rsiva.Fields("Impuesto").Value
    
    Rsiva.Close

If IsNull(DMGrid1.ValorCelda(s, 4)) Then
    a = 0
ElseIf Val(DMGrid1.ValorCelda(s, 4)) = 0 Then
    a = 0
Else
    a = DMGrid1.ValorCelda(s, 4)
End If
'Call QuitarCaracter(a)
'a = CArac

If IsNull(DMGrid1.ValorCelda(s, 3)) Then
    b = 1
    DMGrid1.ValorCelda(s, 3) = 1
ElseIf Val(DMGrid1.ValorCelda(s, 3)) = 0 Then
    b = 1
    DMGrid1.ValorCelda(s, 3) = 1
Else
    b = DMGrid1.ValorCelda(s, 3)
End If

'Call QuitarCaracter(b)
'b = CArac

'calculo del impuesto
If CalculaIva = True Then
    DMGrid1.ValorCelda(s, 5) = a * b * (IVA / 100)
    DMGrid1.RowBackColor s, RGB(255, 255, 255)
Else
    DMGrid1.ValorCelda(s, 5) = Format(0, "#,##0.00")
    DMGrid1.RowBackColor s, RGB(221, 221, 221)
End If

If IsNull(DMGrid1.ValorCelda(s, 5)) Then
    C = Format(0, "#,##0.00")
ElseIf Val(DMGrid1.ValorCelda(s, 5)) = 0 Then
    C = Format(0, "#,##0.00")
Else
    C = DMGrid1.ValorCelda(s, 5)
End If
'Call QuitarCaracter(c)
'c = CArac


If IsNull(DMGrid1.ValorCelda(s, 6)) Then
    d = Format(0, "#,##0.00")
ElseIf Val(DMGrid1.ValorCelda(s, 6)) = 0 Then
    d = Format(0, "#,##0.00")
Else
    d = DMGrid1.ValorCelda(s, 6)
End If
'Call QuitarCaracter(d)
'd = CArac

If IsNull(a) Then a = 0 ' Precio Unitario
If IsNull(b) Then b = 0 ' Cantidad
If IsNull(C) Then C = 0 ' Iva
If IsNull(d) Then d = 0 ' Descuento
  
'DMGrid1.ValorCelda(s, 7) = (a * b - d) * (1 + (c / 100))
DMGrid1.ValorCelda(s, 7) = (a * b - d) + C
DMGrid1.ValorCelda(s, 1) = DMGrid1.ValorCelda(DMGrid1.Row, 1)
If s <> DMGrid1.Row Then DMGrid1.ValorCelda(DMGrid1.Row, 1) = ""
DMGrid1.PaintMGrid
calcular
End Sub

Private Sub DMGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift <> 0 Then
    Select Case KeyCode
        Case 38
            Text4.SetFocus
        Case 39
            TxtBuscaru.SetFocus
        Case 40
            BtnAgregarRenglon.SetFocus
    End Select
End If
'If KeyCode = 112 And DMGrid1.Col = 1 Then 'tecla F1
'    f = DMGrid1.Row
'    opcion = "Facturacion"
'    FrmListadoProductosServicios.Show
'
'    'DMGrid1.ValorCelda(f, 1) = CodProd
'    'DMGrid1.ValorCelda(f, 2) = DescPro
'    'DMGrid1.ValorCelda(f, 4) = PreProd
'    'DMGrid1.ValorCelda(f, 5) = IvaProd
'    'DMGrid1.Col = 3
'    'Call DMGrid1.PaintMGrid
'
'End If

End Sub


Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton And DMGrid1.Col = 1 Then
    f = DMGrid1.Row
    opcion = "Facturacion"
    FrmListadoProductosServicios.Show vbModal, FrmPrincipal
End If
End Sub



Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo1.SetFocus
        Case 37
            Text11.SetFocus
        Case 40
            Combo1.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me
Pago
Grid2
ValorImpuesto
'Command1.Enabled = True
Tipo = "Facturacion"
IdCliente = 0
SQL = ""
IdPac1 = ""
N_fac = ""
T = 0
pagos = 0
CondFact = 0
IO = 0
DMGrid1.RowBackColor 1, RGB(255, 255, 255)
LblCantidadRenglon = DMGrid1.Rows
DTPicker1.Value = Format(Now, "dd/mm/yyyy")

End Sub

Sub ValorImpuesto()
Dim RsValorImpuesto As New ADODB.Recordset
' carga los diferentes valores de los impuestos
 
CSql = "SELECT IVA1 FROM Dat_admin"
Set RsValorImpuesto = CrearRS(CSql)

    IVA = RsValorImpuesto.Fields("IVA1").Value
    LblImpuestoIVA.Caption = "I.v.a.: (" & RsValorImpuesto.Fields("IVA1").Value & " %)"

RsValorImpuesto.Close

End Sub

Sub Pago()
Dim RsFormaPago As New ADODB.Recordset
' carga los tipos de opciones de pago de la factura en un combo
CSql = "SELECT * FROM Pago where activo=1"
Set RsFormaPago = CrearRS(CSql)

If RsFormaPago.EOF Then RsFormaPago.Close: Exit Sub
    RsFormaPago.MoveFirst
    Combo1.Clear
Do While Not RsFormaPago.EOF
    Combo1.AddItem RsFormaPago.Fields("Tipo").Value
    Combo1.ItemData(Combo1.NewIndex) = RsFormaPago.Fields("id").Value
    RsFormaPago.MoveNext
Loop
RsFormaPago.Close

End Sub

Sub Precio()
Dim RsPrecio As New ADODB.Recordset
CSql = "SELECT SUM(Precio) as monto2 FROM Presupuesto2 WHERE N_factura = " & N_fac
Set RsPrecio = CrearRS(CSql)

If RsPrecio.Fields("Monto2").Value <> "" Then Text8.Caption = RsPrecio.Fields("Monto2").Value Else Text8.Caption = ""
RsPrecio.Close

End Sub

Sub Factura()
Dim RsNoFact As New ADODB.Recordset
CSql = "select * from dat_admin"

Set RsNoFact = CrearRS(CSql)

Fact = Format(RsNoFact.Fields("U_Factura").Value + 1, "000000000#")

Label12.Caption = Fact

Dim RsUpdateNoFact As New ADODB.Recordset
CSql = "update dat_admin SET U_Factura = " & Str(Fact) & " WHERE U_Factura = " & Str(Fact - 1) & ";"
Set RsUpdateNoFact = CrearRS(CSql)

RsNoFact.Close
RsUpdateNoFact.Close
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    d = IdCliente
    FrmListadoClientes.Show vbModal, FrmPrincipal
    If d <> IdCliente Then
        Carga_Cliente
    End If
End If
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case 39
            BtnBuscar.SetFocus
        Case 40
            Text2.SetFocus
    End Select
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case 39
            DTPicker1.SetFocus
        Case 38
            Text9.SetFocus
        Case 37
            Text2.SetFocus
        Case 40
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text12.SetFocus
        Case 39
            DTPicker1.SetFocus
        Case 38
            Text10.SetFocus
        Case 37
            Text3.SetFocus
        Case 40
            Text12.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case vbKeyLeft
            Text4.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            Text13.SetFocus
    End Select
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DMGrid1.SetFocus
        Case 37
            Text4.SetFocus
        Case 38
            Text12.SetFocus
        Case 39
            DTPicker1.SetFocus
        Case 40
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Sub Carga_Paciente()
Dim RsCargarPaciente As New ADODB.Recordset
'datos de paciente si es que los hay
If IdPac1 <> "" Then
    CSql = "Select * from paciente where idpaciente = '" & IdPac1 & "'"
    Set RsCargarPaciente = CrearRS(CSql)
    If RsCargarPaciente.RecordCount > 0 Then
    Text9.Text = RsCargarPaciente.Fields("Cedulap").Value
    Text11.Text = RsCargarPaciente.Fields("Nombrep").Value
    Text10.Text = RsCargarPaciente.Fields("Apellidop").Value
    Text13.Text = "0" & RsCargarPaciente.Fields("Codigo").Value & "-" & RsCargarPaciente.Fields("Telefono").Value
    Text12.Text = RsCargarPaciente.Fields("DireccionP").Value
    End If
    RsCargarPaciente.Close
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case 38
            Text1.SetFocus
        Case 39
            Text10.SetFocus
        Case 40
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyRight
            Text11.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case 38
            Text3.SetFocus
        Case 39
            Text12.SetFocus
        Case 40
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    d = IdPac1
    FrmListadoPaciente.Show
    If d <> IdPac1 Then
        Carga_Paciente
    End If
End If
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case 37
            BtnAgregarCliente.SetFocus
        Case 39
            ChameleonBtn2.SetFocus
        Case 40
            Text10.SetFocus
    End Select
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnImprimir.SetFocus
        Case 37
            DMGrid1.SetFocus
        Case 38
            If Check1.Enabled = True Then
                Check1.SetFocus
            Else
                Combo1.SetFocus
            End If
        Case 39
            Combo1.SetFocus
        Case 40
            BtnImprimir.SetFocus
    End Select
End If
End Sub

Public Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
If KeyAscii = 13 Then

    Dim RsBuscarFactura As New ADODB.Recordset
    If TxtBuscar.Text = "" Then Exit Sub
    If Val(TxtBuscar.Text) = 0 Then Exit Sub
        
    CSql = "Select * From C_Cobrar Where N_Factura = '" & Trim(TxtBuscar.Text) & "' AND C_NC=0"
    Set RsBuscarFactura = CrearRS(CSql)
    
    If RsBuscarFactura.RecordCount = 0 Then
        CSql = "Select * From C_Cobrar Where N_FA = '" & Trim(TxtBuscar.Text) & "' AND C_NC=1"
        Set RsBuscarFactura = CrearRS(CSql)
    
        If RsBuscarFactura.RecordCount <> 0 Then
            
            If Val(RsBuscarFactura.Fields("Anulada").Value) = 1 Then MsgBox "La factura de Nro. " & TxtBuscar.Text & " se encuentra anulada!", vbExclamation + vbOKOnly, "Información": Exit Sub
            
            MsgBox "El numero de factura introducido tiene una Nota de Crédito de Nro. " & Format(RsBuscarFactura.Fields("N_NC").Value, "0000") & "!", vbExclamation + vbOKOnly, "Información"
            RsBuscarFactura.Close
            Exit Sub
        Else
            RsBuscarFactura.Close
            MsgBox "El numero de factura introducido no existe!", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
    End If
    
    If Val(RsBuscarFactura.Fields("Anulada").Value) = 1 Then MsgBox "La factura de Nro. " & TxtBuscar.Text & " se encuentra anulada!", vbExclamation + vbOKOnly, "Información": Exit Sub
    
    BtnImportar.Enabled = False
    
    CondFact = 1 'debera de actualizar
    'datos de factura
    If RsBuscarFactura.Fields("Impresa").Value = True Then Check1.Value = 1: BtnRetenciones.Enabled = True Else Check1.Value = 0: BtnRetenciones.Enabled = False
    N_Factur = RsBuscarFactura.Fields("N_Factura").Value
    Label12.Caption = Format(N_Factur, "00000000#")
    DTPicker1.Value = RsBuscarFactura.Fields("Fecha").Value
    Combo1.ListIndex = RsBuscarFactura.Fields("Forma_Pago").Value
    IdCliente = RsBuscarFactura.Fields("IdCliente").Value
    If Trim(RsBuscarFactura.Fields("IdPaciente").Value) = "" Then IdPac1 = "" Else IdPac1 = RsBuscarFactura.Fields("IdPaciente").Value
    
    Call Carga_Cliente
    Call Carga_Paciente
    
    Dim RsCargarRenglon As New ADODB.Recordset
    'cargar los renglones de la factura
    CSql = "Select * From Reng_Cobrar Where N_Factura = '" & Trim(TxtBuscar.Text) & "'"
    Set RsCargarRenglon = CrearRS(CSql)
    i = 1
    If Not (RsCargarRenglon.EOF) Then
        RsCargarRenglon.MoveFirst
        Dim RsProductos As New ADODB.Recordset
        Do While Not RsCargarRenglon.EOF
            DMGrid1.Rows = i
            CSql = "Select * From Productos Where IdProducto ='" & RsCargarRenglon.Fields("Cod_Producto").Value & "'"
            Set RsProductos = CrearRS(CSql)
            DMGrid1.ValorCelda(i, 1) = RsCargarRenglon.Fields("Cod_Producto").Value
            DMGrid1.ValorCelda(i, 2) = RsProductos.Fields("Descripcion").Value
            DMGrid1.ValorCelda(i, 3) = RsCargarRenglon.Fields("Cantidad").Value
            DMGrid1.ValorCelda(i, 4) = Format(RsCargarRenglon.Fields("Precio").Value, "#,##0.00")
            DMGrid1.ValorCelda(i, 5) = Format(RsCargarRenglon.Fields("Iva").Value, "#,##0.00")
            DMGrid1.ValorCelda(i, 6) = Format(RsCargarRenglon.Fields("Descuento").Value, "#,##0.00")
            If RsCargarRenglon.Fields("SubTotal").Value <> "" Then
                DMGrid1.ValorCelda(i, 7) = Format(RsCargarRenglon.Fields("SubTotal").Value, "#,##0.00")
            Else
                DMGrid1.ValorCelda(i, 7) = Format(0, "#,##0.00")
            End If
            
            If RsProductos.Fields("Impuesto") Then
                DMGrid1.RowBackColor i, RGB(255, 255, 255)
            Else
                DMGrid1.RowBackColor i, RGB(221, 221, 221)
            End If
            RsProductos.Close
            i = i + 1
            
            RsCargarRenglon.MoveNext
        Loop
    Else
        DMGrid1.Clear
        Call DMGrid1.PaintMGrid
        Text6.Caption = Format(0, Standard)
        Text7.Caption = Format(0, Standard)
        Text8.Caption = Format(0, Standard)
        RsCargarRenglon.Close
        RsBuscarFactura.Close
        Exit Sub
    End If
    DMGrid1.PaintMGrid
    RsCargarRenglon.Close
    RsBuscarFactura.Close
    LblCantidadRenglon.Caption = DMGrid1.Rows
    calcular
End If

End Sub


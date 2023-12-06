VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCompras 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras"
   ClientHeight    =   7980
   ClientLeft      =   1740
   ClientTop       =   2145
   ClientWidth     =   15345
   Icon            =   "FrmCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   15345
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Proveedor"
         Height          =   6735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14895
         Begin VB.TextBox TxtNoFactura 
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   1100
            Width           =   1695
         End
         Begin VB.TextBox TxtNoControl 
            Height          =   375
            Left            =   1920
            TabIndex        =   48
            Top             =   1100
            Width           =   1695
         End
         Begin VB.TextBox TxtNoOrdenCompra 
            Height          =   375
            Left            =   3720
            TabIndex        =   47
            Top             =   1100
            Width           =   1575
         End
         Begin ChamaleonButton.ChameleonBtn BtnImportarOrdenCompra 
            Height          =   375
            Left            =   11520
            TabIndex        =   42
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Importar Orden de Compra"
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
            MICON           =   "FrmCompras.frx":1002
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnProductos 
            Height          =   375
            Left            =   13200
            TabIndex        =   13
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Productos"
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
            MICON           =   "FrmCompras.frx":101E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarProveedor 
            Height          =   375
            Left            =   1440
            TabIndex        =   41
            ToolTipText     =   "Buscar Proveedor"
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
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
            MICON           =   "FrmCompras.frx":103A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtRif 
            Height          =   375
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox TxtDescripcionProveerdor 
            Height          =   375
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   480
            Width           =   6615
         End
         Begin VB.TextBox TxtCodigoProveedor 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox CboCondicionPago 
            Height          =   315
            Left            =   5400
            TabIndex        =   26
            Top             =   1100
            Width           =   2295
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Detalle"
            Height          =   5055
            Left            =   120
            TabIndex        =   15
            Top             =   1560
            Width           =   14655
            Begin SystemOncoAmerica.DMGrid DMGrid1 
               Height          =   4095
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   7223
               Object.Width           =   14385
               Object.Height          =   4065
               ScrollBar       =   1
               Editable        =   -1  'True
               DrawColorGrid   =   1
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Procesada"
               Enabled         =   0   'False
               Height          =   255
               Left            =   3600
               TabIndex        =   43
               Top             =   4500
               Width           =   1215
            End
            Begin VB.TextBox TxtTotalGeneral 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   12840
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "0.00"
               Top             =   4440
               Width           =   1695
            End
            Begin VB.TextBox TxtSubTotal 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "0.00"
               Top             =   4440
               Width           =   1695
            End
            Begin VB.TextBox TxtImpuesto 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   9960
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "0.00"
               Top             =   4440
               Width           =   1695
            End
            Begin VB.Timer Timer1 
               Interval        =   1000
               Left            =   13200
               Top             =   3720
            End
            Begin VB.TextBox TxtCantidadRenglon 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "0"
               Top             =   4440
               Width           =   615
            End
            Begin ChamaleonButton.ChameleonBtn BtnEliminarRenglon 
               Height          =   375
               Left            =   1080
               TabIndex        =   20
               ToolTipText     =   "Eliminar Renglon"
               Top             =   4440
               Width           =   975
               _ExtentX        =   1720
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
               MICON           =   "FrmCompras.frx":1056
               PICN            =   "FrmCompras.frx":1072
               PICH            =   "FrmCompras.frx":1216
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
               TabIndex        =   21
               ToolTipText     =   "Agregar Renglon"
               Top             =   4440
               Width           =   975
               _ExtentX        =   1720
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
               MICON           =   "FrmCompras.frx":13B5
               PICN            =   "FrmCompras.frx":13D1
               PICH            =   "FrmCompras.frx":17F4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnProcesada 
               Height          =   375
               Left            =   2160
               TabIndex        =   52
               ToolTipText     =   "Procesar Consumos de Medicamentos"
               Top             =   4440
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
               MICON           =   "FrmCompras.frx":1A29
               PICN            =   "FrmCompras.frx":1A45
               PICH            =   "FrmCompras.frx":1CBA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total General:"
               Height          =   195
               Left            =   11760
               TabIndex        =   25
               Top             =   4530
               Width           =   1005
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal:"
               Height          =   195
               Left            =   6480
               TabIndex        =   24
               Top             =   4530
               Width           =   690
            End
            Begin VB.Label LblValorImpuesto 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "I.V.A.:(12%)"
               Height          =   195
               Left            =   9000
               TabIndex        =   23
               Top             =   4530
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad:"
               Height          =   195
               Left            =   5040
               TabIndex        =   22
               Top             =   4530
               Width           =   675
            End
         End
         Begin ChamaleonButton.ChameleonBtn BtnProveedores 
            Height          =   375
            Left            =   11520
            TabIndex        =   14
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Proveedores"
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
            MICON           =   "FrmCompras.frx":1F36
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPickerFechaEmision 
            Height          =   375
            Left            =   7800
            TabIndex        =   30
            Top             =   1095
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53477377
            CurrentDate     =   39932
         End
         Begin MSComCtl2.DTPicker DTPickerFechaRecepcion 
            Height          =   375
            Left            =   9360
            TabIndex        =   31
            Top             =   1095
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53477377
            CurrentDate     =   39932
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Control"
            Height          =   195
            Left            =   1920
            TabIndex        =   46
            Top             =   870
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Factura"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   870
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif.: "
            Height          =   195
            Left            =   9000
            TabIndex        =   40
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   2280
            TabIndex        =   39
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición de Pago:"
            Height          =   195
            Left            =   5400
            TabIndex        =   37
            Top             =   870
            Width           =   1395
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   7800
            TabIndex        =   36
            Top             =   870
            Width           =   1080
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Recepción:"
            Height          =   195
            Left            =   9360
            TabIndex        =   35
            Top             =   870
            Width           =   1320
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Orden Compra "
            Height          =   195
            Left            =   3720
            TabIndex        =   34
            Top             =   870
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compra No.:"
            Height          =   195
            Left            =   11520
            TabIndex        =   33
            Top             =   330
            Width           =   885
         End
         Begin VB.Label LblNoOrden 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12600
            TabIndex        =   32
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   7080
         Width           =   3495
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
            Left            =   960
            TabIndex        =   9
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Numero de Compra"
            Top             =   300
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2520
            TabIndex        =   10
            ToolTipText     =   "Buscar Ordenes de Compra"
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
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
            MICON           =   "FrmCompras.frx":1F52
            PICN            =   "FrmCompras.frx":1F6E
            PICH            =   "FrmCompras.frx":21D3
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
            Caption         =   "Nº Compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   390
            Width           =   810
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3720
         TabIndex        =   1
         Top             =   7080
         Width           =   11295
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   10200
            TabIndex        =   2
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
            MICON           =   "FrmCompras.frx":2465
            PICN            =   "FrmCompras.frx":2481
            PICH            =   "FrmCompras.frx":264A
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
            Height          =   375
            Left            =   1560
            TabIndex        =   3
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "FrmCompras.frx":287F
            PICN            =   "FrmCompras.frx":289B
            PICH            =   "FrmCompras.frx":2B2A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnNuevo 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "FrmCompras.frx":2F6B
            PICN            =   "FrmCompras.frx":2F87
            PICH            =   "FrmCompras.frx":3114
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnDeshacer 
            Height          =   375
            Left            =   9000
            TabIndex        =   5
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
            MICON           =   "FrmCompras.frx":3349
            PICN            =   "FrmCompras.frx":3365
            PICH            =   "FrmCompras.frx":3647
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
            Left            =   4800
            TabIndex        =   6
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
            MICON           =   "FrmCompras.frx":3898
            PICN            =   "FrmCompras.frx":38B4
            PICH            =   "FrmCompras.frx":39D9
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
            Left            =   3000
            TabIndex        =   7
            ToolTipText     =   "Borrar"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "FrmCompras.frx":3C69
            PICN            =   "FrmCompras.frx":3C85
            PICH            =   "FrmCompras.frx":3E29
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
            Left            =   4320
            Top             =   240
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
         Begin ChamaleonButton.ChameleonBtn BtnRetenciones 
            Height          =   375
            Left            =   6720
            TabIndex        =   50
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Aplicar Pagos"
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
            MICON           =   "FrmCompras.frx":3FC8
            PICN            =   "FrmCompras.frx":3FE4
            PICH            =   "FrmCompras.frx":441A
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
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   2400
      TabIndex        =   45
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "FrmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsUpDateDatAdmin As Recordset
Dim RsGuardarOrden As Recordset
Dim RsGuardarCuentaPagar As Recordset
Dim Impuesto As Double

Private Sub BtnAgregarRenglon_Click()
If Check1.Value = 0 Then
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
    TxtCantidadRenglon.Text = DMGrid1.Rows + 1
    DMGrid1.PaintMGrid
Else
    Msg = "Ya esta Órden de Compra fue procesada y no puede ser modificada"
    MsgBox Msg, vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If
End Sub

Private Sub BtnBorrar_Click()
If Check1.Value = 1 Then
    Msg = "No se puede borrar la compra porque ya esta procesada!!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

End Sub

Private Sub BtnBuscar_Click()

If TxtBuscar.Text <> "" Then
    opcion = 1
    Frame3.Enabled = True
    CSql = "Select * From Compras Where NumeroCompra = '" & Trim(TxtBuscar.Text) & "'"
    Dim RsBuscarOrdenes As New ADODB.Recordset
    Set RsBuscarOrdenes = CrearRS(CSql)
    
    If RsBuscarOrdenes.RecordCount = 0 Then Exit Sub
    
    If RsBuscarOrdenes.Fields("OrdenProcesada").Value = True Then
        BtnNuevo.Enabled = True
        BtnGuardar.Enabled = False
        BtnBorrar.Enabled = False
        BtnImprimir.Enabled = True
        BtnAgregarRenglon.Enabled = False
        BtnEliminarRenglon.Enabled = False
        Check1.Value = 1
        BtnProcesada.Enabled = False
    Else
        BtnNuevo.Enabled = False
        BtnGuardar.Enabled = True
        BtnBorrar.Enabled = False
        BtnImprimir.Enabled = True
        BtnAgregarRenglon.Enabled = True
        BtnEliminarRenglon.Enabled = True
        Check1.Value = 0
        BtnProcesada.Enabled = True
    End If
    
    If RsBuscarOrdenes.EOF = False Or RsBuscarOrdenes.BOF = False Then
        LblNoOrden.Caption = RsBuscarOrdenes.Fields("NumeroCompra").Value
        TxtNoOrdenCompra.Text = RsBuscarOrdenes.Fields("NumeroOrden").Value
        TxtNoFactura.Text = RsBuscarOrdenes.Fields("NoFactura").Value
        TxtNoControl.Text = RsBuscarOrdenes.Fields("NoControl").Value
        CboCondicionPago.Text = RsBuscarOrdenes.Fields("CondicionPago").Value
        DTPickerFechaEmision.Value = RsBuscarOrdenes.Fields("FechaEmision").Value
        DTPickerFechaRecepcion.Value = RsBuscarOrdenes.Fields("FechaRecepcion").Value
        TxtSubTotal.Text = Format(RsBuscarOrdenes.Fields("SubTotal").Value, "#,##0.00")
        TxtImpuesto.Text = Format(RsBuscarOrdenes.Fields("Impuesto").Value, "#,##0.00")
        TxtTotalGeneral.Text = Format(RsBuscarOrdenes.Fields("TotalGeneral").Value, "#,##0.00")
        CodProveedor = RsBuscarOrdenes.Fields("IdProveedor").Value
        
        If RsBuscarOrdenes.Fields("OrdenProcesada").Value = True Then
            Check1.Value = 1
            BtnRetenciones.Enabled = True
        Else
            Check1.Value = 0
            BtnRetenciones.Enabled = False
        End If
        
    End If
 
    
    CSql = "Select * From Proveedores Where IdProveedor = '" & Trim(CodProveedor) & "'"
    Dim RsBuscarProveedor As New ADODB.Recordset
    Set RsBuscarProveedor = CrearRS(CSql)
    If RsBuscarProveedor.RecordCount > 0 Then
        TxtCodigoProveedor.Text = RsBuscarProveedor.Fields("IdProveedor").Value
        TxtDescripcionProveerdor.Text = RsBuscarProveedor.Fields("Nombre").Value
        TxtRif.Text = RsBuscarProveedor.Fields("RifProveedor").Value
    End If
    CSql = "Select * From RenglonCompras Where NumeroCompra = '" & Trim(TxtBuscar.Text) & "'"
    Dim RsBuscarRenglonCompra As New ADODB.Recordset
    Set RsBuscarRenglonCompra = CrearRS(CSql)

    i = 1
        If Not (RsBuscarRenglonCompra.EOF) Then
            RsBuscarRenglonCompra.MoveFirst
            Dim RsProductos As New ADODB.Recordset
            Do While Not RsBuscarRenglonCompra.EOF
                DMGrid1.Rows = i
                'DMGrid1.RowBackColor 1, RGB(255, 255, 255)
                CSql = "Select * From Productos Where IdProducto ='" & RsBuscarRenglonCompra.Fields("IdProducto").Value & "'"
                Set RsProductos = CrearRS(CSql)
                If RsBuscarRenglonCompra.RecordCount > 0 Then
                    DMGrid1.ValorCelda(i, 1) = RsBuscarRenglonCompra.Fields("IdProducto").Value
                    DMGrid1.ValorCelda(i, 2) = RsProductos.Fields("Descripcion").Value
                    DMGrid1.ValorCelda(i, 3) = RsBuscarRenglonCompra.Fields("Cantidad").Value
                    DMGrid1.ValorCelda(i, 4) = RsBuscarRenglonCompra.Fields("precio").Value
                    DMGrid1.ValorCelda(i, 5) = RsBuscarRenglonCompra.Fields("impuesto").Value
                    DMGrid1.ValorCelda(i, 6) = RsBuscarRenglonCompra.Fields("descuento").Value
                    DMGrid1.ValorCelda(i, 7) = RsBuscarRenglonCompra.Fields("SubTotal").Value
                    If RsProductos.Fields("Impuesto") Then
                        DMGrid1.RowBackColor i, RGB(255, 255, 255)
                    Else
                        DMGrid1.RowBackColor i, RGB(221, 221, 221)
                    End If
                    RsProductos.Close
                    i = i + 1
                End If
                RsBuscarRenglonCompra.MoveNext
            Loop
               TxtCantidadRenglon.Text = DMGrid1.Rows
    Else
            DMGrid1.Clear
            Call DMGrid1.PaintMGrid
            Text6.Text = Format(0, Standard)
            Text7.Text = Format(0, Standard)
            Text8.Text = Format(0, Standard)
            RsBuscarRenglonCompra.Close
            RsBuscarFactura.Close
            Exit Sub
    End If
    DMGrid1.PaintMGrid
    RsBuscarOrdenes.Close
    RsBuscarProveedor.Close
    calcular
    BtnNuevo.Enabled = True
    BtnNuevo.SetFocus
Else
    Exit Sub
End If
End Sub

Private Sub BtnBuscarProveedor_Click()
Tipo = "Compras"
FrmListadoProveedor.Show
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
BtnAgregarRenglon.Enabled = False
BtnEliminarRenglon.Enabled = False
BtnProcesada.Enabled = False
BtnNuevo.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = False
BtnImprimir.Enabled = False
DMGrid1.Clear
DMGrid1.Rows = 0
DMGrid1.PaintMGrid
Blanqueo
NoOrden
End Sub

Sub Blanqueo()
TxtCodigoProveedor.Text = ""
TxtDescripcionProveerdor.Text = ""
TxtRif.Text = ""
TxtNoFactura.Text = ""
TxtNoControl.Text = ""
CboCondicionPago.Text = ""

TxtSubTotal.Text = ""
TxtImpuesto.Text = ""
TxtTotalGeneral.Text = ""
DTPickerFechaEmision.Value = Now
DTPickerFechaRecepcion.Value = Now
DMGrid1.Clear
DMGrid1.Rows = 0
End Sub

Private Sub BtnEliminarRenglon_Click()
If Check1.Value = 0 Then
    DMGrid1.RowDelete (DMGrid1.Row)
    TxtCantidadRenglon.Text = DMGrid1.Rows - 1
    DMGrid1.PaintMGrid
    Call calcular
Else
    Msg = "Ya esta Órden de Compra fue procesada y no puede ser modificada"
    MsgBox Msg
End If
End Sub

Private Sub BtnGuardar_Click()
If TxtCodigoProveedor.Text = "" Then
    MsgBox "Esta dejando el Codigo del Proveedor en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If
If TxtDescripcionProveerdor.Text = "" Then
    MsgBox "Esta dejando el Nombre del Proveedor en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If
If TxtRif.Text = "" Then
    MsgBox "Esta dejando el Rif del Proveedor en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

If TxtNoFactura.Text = "" Then
    MsgBox "Esta dejando el No. de Factura en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

If TxtNoControl.Text = "" Then
    MsgBox "Esta dejando el No. de Control en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

If CboCondicionPago.Text = "" Then
    MsgBox "Esta dejando la Condicion de Pago en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

If CboCondicionPago.ListIndex = -1 Then
    MsgBox "Seleccione la Condicion de Pago Nuevamente", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

If DMGrid1.Rows <= 0 Then
    MsgBox "No hay Productos a Guardar", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

Select Case opcion
    Case Is = 0
        
        CSql = "Select * From Compras"
        Set RsGuardarOrden = CrearRS(CSql)
                
        RsGuardarOrden.AddNew
        RsGuardarOrden.Fields("NumeroCompra").Value = LblNoOrden.Caption
        RsGuardarOrden.Fields("NumeroOrden").Value = TxtNoOrdenCompra.Text
        RsGuardarOrden.Fields("IdProveedor").Value = Trim(TxtCodigoProveedor.Text)
        RsGuardarOrden.Fields("NoFactura").Value = Trim(TxtNoFactura.Text)
        RsGuardarOrden.Fields("NoControl").Value = Trim(TxtNoControl.Text)
        RsGuardarOrden.Fields("CondicionPago").Value = Trim(CboCondicionPago.Text)
        RsGuardarOrden.Fields("FechaEmision").Value = Format(DTPickerFechaEmision.Value, "dd/mm/yyyy")
        RsGuardarOrden.Fields("FechaRecepcion").Value = Format(DTPickerFechaRecepcion.Value, "dd/mm/yyyy")
        RsGuardarOrden.Fields("OrdenProcesada").Value = 0
        RsGuardarOrden.Fields("SubTotal").Value = Trim(TxtSubTotal.Text)
        RsGuardarOrden.Fields("Impuesto").Value = Trim(TxtImpuesto.Text)
        RsGuardarOrden.Fields("TotalGeneral").Value = Trim(TxtTotalGeneral.Text)
       ' RsGuardarOrden.Fields("Impresa").Value = Check1.Value
        RsGuardarOrden.Update
        
        CSql = "Select * From CtaPorPagar"
        Set RsGuardarCuentaPagar = CrearRS(CSql)
                
        RsGuardarCuentaPagar.AddNew
        RsGuardarCuentaPagar.Fields("NumeroCompra").Value = LblNoOrden.Caption
        RsGuardarCuentaPagar.Fields("NumeroOrden").Value = TxtNoOrdenCompra.Text
        RsGuardarCuentaPagar.Fields("NoFactura").Value = Trim(TxtNoFactura.Text)
        RsGuardarCuentaPagar.Fields("NoControl").Value = Trim(TxtNoControl.Text)
        RsGuardarCuentaPagar.Fields("IdProveedor").Value = Trim(TxtCodigoProveedor.Text)
        RsGuardarCuentaPagar.Fields("Nombre").Value = Trim(TxtDescripcionProveerdor.Text)
        RsGuardarCuentaPagar.Fields("Rif").Value = Trim(TxtRif.Text)
        
        RsGuardarCuentaPagar.Fields("DiasCredito").Value = Trim(CboCondicionPago.Text)
        
        RsGuardarCuentaPagar.Fields("FechaEmision").Value = Format(DTPickerFechaEmision.Value, "dd/mm/yyyy")
        RsGuardarCuentaPagar.Fields("FechaRecepcion").Value = Format(DTPickerFechaRecepcion.Value, "dd/mm/yyyy")
        
        RsGuardarCuentaPagar.Fields("Procesada").Value = 0
        RsGuardarCuentaPagar.Fields("TotalBase").Value = Trim(TxtSubTotal.Text)
        RsGuardarCuentaPagar.Fields("SubTotal").Value = Trim(TxtSubTotal.Text)
        RsGuardarCuentaPagar.Fields("Impuesto").Value = Trim(TxtImpuesto.Text)
        RsGuardarCuentaPagar.Fields("TotalGeneral").Value = Trim(TxtTotalGeneral.Text)
        RsGuardarCuentaPagar.Fields("CondicionPago").Value = CboCondicionPago.ItemData(CboCondicionPago.ListIndex)
        RsGuardarCuentaPagar.Fields("PorPagar").Value = Trim(TxtTotalGeneral.Text)
        RsGuardarCuentaPagar.Update
        
        Dim RsGuardarRenglonOrdenes As New ADODB.Recordset
        CSql = "Select * From RenglonCompras"
        Set RsGuardarRenglonOrdenes = CrearRS(CSql)
                
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
               
            RsGuardarRenglonOrdenes.AddNew
            RsGuardarRenglonOrdenes.Fields("NumeroCompra").Value = LblNoOrden.Caption
            RsGuardarRenglonOrdenes.Fields("NumeroOrden").Value = TxtNoOrdenCompra.Text
            RsGuardarRenglonOrdenes.Fields("NumeroRenglon").Value = i
            RsGuardarRenglonOrdenes.Fields("IdProducto").Value = Val(b1)
            RsGuardarRenglonOrdenes.Fields("Cantidad").Value = Val(b3)
            RsGuardarRenglonOrdenes.Fields("Precio").Value = b2
            RsGuardarRenglonOrdenes.Fields("Impuesto").Value = b4
            RsGuardarRenglonOrdenes.Fields("Descuento").Value = b5
            RsGuardarRenglonOrdenes.Fields("SubTotal").Value = CDbl(b6)
            RsGuardarRenglonOrdenes.Update
r:
        TxtCantidadRenglon.Text = DMGrid1.Rows
        Next i
        
        CSql = "Select * From Dat_Admin"
        Set RsUpDateDatAdmin = CrearRS(CSql)
        
        RsUpDateDatAdmin.Fields("U_Compra").Value = Val(LblNoOrden.Caption)
        
        RsUpDateDatAdmin.Update
        RsUpDateDatAdmin.Close
       ' NoOrden
    
    Case Is = 1
        
        CSql = "Select * From Compras Where NumeroCompra='" & Trim(LblNoOrden.Caption) & "'"
        Set RsGuardarOrden = CrearRS(CSql)
                
        RsGuardarOrden.Fields("NumeroCompra").Value = LblNoOrden.Caption
        RsGuardarOrden.Fields("NumeroOrden").Value = TxtNoOrdenCompra.Text
        RsGuardarOrden.Fields("IdProveedor").Value = Trim(TxtCodigoProveedor.Text)
        RsGuardarOrden.Fields("NoFactura").Value = Trim(TxtNoFactura.Text)
        RsGuardarOrden.Fields("NoControl").Value = Trim(TxtNoControl.Text)
        RsGuardarOrden.Fields("CondicionPago").Value = Trim(CboCondicionPago.Text)
        RsGuardarOrden.Fields("FechaEmision").Value = Format(DTPickerFechaEmision.Value, "dd/mm/yyyy")
        RsGuardarOrden.Fields("FechaRecepcion").Value = Format(DTPickerFechaRecepcion.Value, "dd/mm/yyyy")
        RsGuardarOrden.Fields("OrdenProcesada").Value = 0
        RsGuardarOrden.Fields("SubTotal").Value = Trim(TxtSubTotal.Text)
        RsGuardarOrden.Fields("Impuesto").Value = Trim(TxtImpuesto.Text)
        RsGuardarOrden.Fields("TotalGeneral").Value = Trim(TxtTotalGeneral.Text)
        RsGuardarOrden.Update
        
        CSql = "Select * From RenglonCompras Where NumeroOrden='" & Trim(LblNoOrden.Caption) & "'"
        Set RsGuardarRenglonOrdenes = CrearRS(CSql)
        If RsGuardarRenglonOrdenes.RecordCount > 0 Then
            For i = 1 To DMGrid1.Rows
                b1 = DMGrid1.ValorCelda(i, 1)
                b2 = DMGrid1.ValorCelda(i, 4)
                b3 = DMGrid1.ValorCelda(i, 3)
                b4 = DMGrid1.ValorCelda(i, 5)
                b5 = DMGrid1.ValorCelda(i, 6)
                b6 = DMGrid1.ValorCelda(i, 7)
                If IsNull(b1) Or Trim(b1) = "" Then GoTo rr
                If IsNull(b5) Or Trim(b5) = "" Then b5 = 0
                If IsNull(b2) Or Trim(b2) = "" Then b2 = 0
                If IsNull(b3) Or Trim(b3) = "" Then b3 = 0
                If IsNull(b4) Or Trim(b4) = "" Then b4 = 0
                If IsNull(b6) Or Trim(b6) = "" Then b6 = 0
                   
               
                
                RsGuardarRenglonOrdenes.Fields("NumeroCompra").Value = LblNoOrden.Caption
                RsGuardarRenglonOrdenes.Fields("NumeroOrden").Value = TxtNoOrdenCompra.Text
                RsGuardarRenglonOrdenes.Fields("NumeroRenglon").Value = i
                RsGuardarRenglonOrdenes.Fields("IdProducto").Value = Val(b1)
                RsGuardarRenglonOrdenes.Fields("Cantidad").Value = Val(b3)
                RsGuardarRenglonOrdenes.Fields("Precio").Value = b2
                RsGuardarRenglonOrdenes.Fields("Impuesto").Value = b4
                RsGuardarRenglonOrdenes.Fields("Descuento").Value = b5
                RsGuardarRenglonOrdenes.Fields("SubTotal").Value = CDbl(b6)
                RsGuardarRenglonOrdenes.Update
                TxtCantidadRenglon.Text = DMGrid1.Rows
rr:
            Next i
        Else
            Exit Sub
        End If
End Select
MsgBox "La factura de Compra fue creada satisfactoriamente!", vbInformation + vbOKOnly, "Operación Exitosa"

BtnAgregarRenglon.Enabled = False
BtnEliminarRenglon.Enabled = False
BtnProcesada.Enabled = False
BtnNuevo.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = False
BtnImprimir.Enabled = False

Msg = "Desea imprimir la Compra No. " & LblNoOrden.Caption
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Impresión")

If mensaje = vbYes Then
    BtnImprimir.Enabled = True
    BtnImprimir_Click
    Blanqueo
    NoOrden
Else
    Msg = "Deseas Procesar la Compra No. " & LblNoOrden.Caption
    mensaje = MsgBox(Msg, vbYesNo + vbInformation, "Procesar Compra")
    
    If mensaje = vbYes Then
        CSql = "Update CtaPorPagar set Procesada = 1 WHERE NumeroCompra = " & LblNoOrden.Caption
        Set RsActualizarImpresa = CrearRS(CSql)
        Call Enviar_Bitacora(IdUser, "COMPRAS", "IMPRIMIR", "Se imprimio la Compra Nro. " & LblNoOrden.Caption)
        
        CSql = "Update Compras set OrdenProcesada = 1 WHERE NumeroCompra = " & LblNoOrden.Caption
        Set RsActualizarImpresa = CrearRS(CSql)
        Call Enviar_Bitacora(IdUser, "COMPRAS", "IMPRIMIR", "Se imprimio la Compra Nro. " & LblNoOrden.Caption)
    End If
    Blanqueo
    NoOrden
End If

End Sub

Private Sub BtnImportarOrdenCompra_Click()
opcion = 0
FrmImportarOrdenesCompra.Show
End Sub

Private Sub BtnImprimir_Click()

If TxtCodigoProveedor.Text <> "" And TxtRif.Text <> "" Then
    If Check1.Value = 0 Then
        Dim RsActualizarImpresa As New ADODB.Recordset
        'colcar aqui el codigo que levante el informe de la factura
        imprime
        Msg = "Deseas Procesar la Compra No. " & LblNoOrden.Caption
        mensaje = MsgBox(Msg, vbYesNo + vbInformation, "Procesar Compra")
        
        If mensaje = vbYes Then
            CSql = "Update CtaPorPagar set Procesada = 1 WHERE NumeroCompra = " & LblNoOrden.Caption
            Set RsActualizarImpresa = CrearRS(CSql)
            Call Enviar_Bitacora(IdUser, "COMPRAS", "IMPRIMIR", "Se imprimio la Compra Nro. " & LblNoOrden.Caption)
            
            CSql = "Update Compras set OrdenProcesada = 1 WHERE NumeroCompra = " & LblNoOrden.Caption
            Set RsActualizarImpresa = CrearRS(CSql)
            Call Enviar_Bitacora(IdUser, "COMPRAS", "IMPRIMIR", "Se imprimio la Compra Nro. " & LblNoOrden.Caption)
        End If
    Else
        Msg = "Ya esta factura de Compra fue impresa desea imprimir una copia ?"
        d = MsgBox(Msg, vbYesNo, "Reimpimir Factura Compra")
        If d = 6 Then
            imprime
        End If
    End If
End If

End Sub
Sub imprime()
''========= ESTE ES EL CODIGO NUEVO ==========
With CrystalReport1
    .ReportFileName = RutaInformes & "\Compras.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Compras1.NumeroCompra} = '" & LblNoOrden.Caption & "'"
    .WindowTitle = "Reporte de Compras No. " & LblNoOrden.Caption
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
End Sub


Private Sub BtnNuevo_Click()

If Check1.Value = 1 Then
    mensaje = MsgBox("Esta Compra fue Procesada y no se le pueden agregar mas items." & Chr(13) & "Deseas crear una nueva compra?", vbYesNo + vbInformation, "Mensaje")
    
    If mensaje = vbYes Then
        Blanqueo
        opcion = 0
        BtnAgregarRenglon.Enabled = True
        BtnEliminarRenglon.Enabled = True
        BtnGuardar.Enabled = True
        BtnImprimir.Enabled = False
        BtnNuevo.Enabled = False
        BtnDesHacer.Enabled = True
        
        DMGrid1.Rows = 0
        DMGrid1.PaintMGrid
        
        TxtCantidadRenglon.Text = 0
        CSql = "Select * From Dat_Admin"
        Set RsUpDateDatAdmin = CrearRS(CSql)
        
        RsUpDateDatAdmin.Fields("U_Compra").Value = Val(LblNoOrden.Caption)
        
        RsUpDateDatAdmin.Update
        RsUpDateDatAdmin.Close
        NoOrden
        
        Check1.Value = 0
        
    End If
Else
    opcion = 0
    BtnAgregarRenglon.Enabled = True
    BtnEliminarRenglon.Enabled = True
    BtnGuardar.Enabled = True
    BtnImprimir.Enabled = False
    BtnNuevo.Enabled = False
    BtnDesHacer.Enabled = True
End If




End Sub

Private Sub BtnProcesada_Click()
Dim RsProcesar As New ADODB.Recordset
Msg = "Estas Seguro(a) de Procesar la Compra!!!"
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Procesar Compra")

If mensaje = vbYes Then
    CSql = "Update CtaPorPagar Set Procesada='1' where NumeroCompra=" & LblNoOrden.Caption & ""
    Set RsProcesar = CrearRS(CSql)
    
    CSql = "Update Compras Set OrdenProcesada='1' where NumeroCompra=" & LblNoOrden.Caption & ""
    Set RsProcesar = CrearRS(CSql)
    
    Check1.Value = 1
    BtnProcesada.Enabled = False
    BtnGuardar.Enabled = False
    BtnBorrar.Enabled = False
    BtnNuevo.Enabled = True
    BtnAgregarRenglon.Enabled = False
    BtnEliminarRenglon.Enabled = False
    
    Msg = "Compra Procesada Correctamente!!!"
    mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Procesar Compras")
    
End If
End Sub

Private Sub BtnProductos_Click()
FrmProductos.Show
End Sub

Private Sub BtnProveedores_Click()
FrmProveedores.Show
End Sub

Private Sub BtnRetenciones_Click()
If Check1.Value = 0 Then
    Msg = "Debe Guardar la Compra para realizar el Pago"
    MsgBox Msg, vbOKOnly + vbInformation, "Mensaje"
    Exit Sub
End If
With FrmAsientoPagos
    .Text3.Text = TxtTotalGeneral.Text
    .Text2.Text = TxtDescripcionProveerdor.Text
    .Text1.Text = LblNoOrden.Caption

End With
BtnRetenciones.Enabled = True
FrmAsientoPagos.Show
End Sub

Private Sub CboCondicionPago_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPickerFechaEmision.SetFocus
        Case vbKeyUp
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyLeft
            TxtNoOrdenCompra.SetFocus
        Case vbKeyRight
            DTPickerFechaEmision.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub Check1_Click()
If Check1.Value Then
    BtnRetenciones.Enabled = True
Else
    BtnRetenciones.Enabled = False
End If
End Sub

Private Sub DMGrid1_AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
Dim PUnit, Cant, IIva, Descu As Double

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
    d = DMGrid1.ValorCelda(DMGrid1.Row, 1)
    
    If d = 0 Or IsNull(d) Then DMGrid1.RowClear (DMGrid1.Row): DMGrid1.RowBackColor DMGrid1.Row, RGB(255, 255, 255): Call calcular: Exit Sub
    
    If d = "" Then DMGrid1.RowClear (DMGrid1.Row): DMGrid1.RowBackColor DMGrid1.Row, RGB(255, 255, 255): Call calcular: Exit Sub
    
    Dim RsProductos As New ADODB.Recordset
    CSql = "Select * from productos where idproducto = " & d
    Set RsProductos = CrearRS(CSql)
    
    If RsProductos.EOF Then RsProductos.Close: DMGrid1.RowClear (DMGrid1.Row): MsgBox "El código del producto no existe en la Base de Datos!", vbExclamation + vbOKOnly, "Código Inexistente": Exit Sub

    ' Crea la busqueda para saber que precio colocar
    CSql = "Select * From Dat_Admin"
    Set RsCargarConfig = CrearRS(CSql)
    
    P1 = RsCargarConfig.Fields("PrecioUnitario1").Value
    P2 = RsCargarConfig.Fields("PrecioUnitario2").Value
    P3 = RsCargarConfig.Fields("PrecioUnitario3").Value
    ValorIva = RsCargarConfig.Fields("Iva1").Value
    RsCargarConfig.Close
    'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
    
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
    
    'If RsProductos.Fields("iva") Then DMGrid1.ValorCelda(f, 5) = IVA Else DMGrid1.ValorCelda(f, 5) = 0
    DMGrid1.ValorCelda(f, 3) = "1"
    DMGrid1.ValorCelda(f, 5) = "0.00"
    DMGrid1.ValorCelda(f, 6) = "0.00"
    DMGrid1.ValorCelda(f, 7) = "0.00"
    RsProductos.Close
End If

CodProd = DMGrid1.ValorCelda(DMGrid1.Row, 1)

If Val(CodProd) = 0 Or IsNull(CodProd) Then DMGrid1.RowClear (DMGrid1.Row): DMGrid1.RowBackColor DMGrid1.Row, RGB(255, 255, 255): Call calcular: Exit Sub

CSql = "Select * From Productos where IdProducto='" & CodProd & "'"
Set Rsiva = CrearRS(CSql)

CalculaIva = Rsiva.Fields("Impuesto").Value

Rsiva.Close

TxtCantidadRenglon.Text = s
If IsNull(DMGrid1.ValorCelda(s, 4)) Then
    PUnit = 0
ElseIf Val(DMGrid1.ValorCelda(s, 4)) = 0 Then
    PUnit = 0
Else
    PUnit = DMGrid1.ValorCelda(s, 4)
End If
'Call QuitarCaracter(a)
'a = CArac

If IsNull(DMGrid1.ValorCelda(s, 3)) Then
    Cant = 1
    DMGrid1.ValorCelda(s, 3) = 1
ElseIf Val(DMGrid1.ValorCelda(s, 3)) = 0 Then
    Cant = 1
    DMGrid1.ValorCelda(s, 3) = 1
Else
    Cant = DMGrid1.ValorCelda(s, 3)
End If
'Call QuitarCaracter(b)
'b = CArac

If CalculaIva = True Then
    DMGrid1.ValorCelda(s, 5) = (PUnit * Cant) * (Impuesto / 100)
    DMGrid1.RowBackColor s, RGB(255, 255, 255)
Else
    DMGrid1.ValorCelda(s, 5) = Format(0, "#,##0.00")
    DMGrid1.RowBackColor s, RGB(221, 221, 221)
End If

If IsNull(DMGrid1.ValorCelda(s, 5)) Then
    IIva = 0
ElseIf Val(DMGrid1.ValorCelda(s, 5)) = 0 Then
    IIva = 0
Else
    IIva = DMGrid1.ValorCelda(s, 5)
End If
'Call QuitarCaracter(c)
'c = CArac

If IsNull(DMGrid1.ValorCelda(s, 6)) Then
    Descu = 0
ElseIf Val(DMGrid1.ValorCelda(s, 6)) = 0 Then
    Descu = 0
Else
    Descu = DMGrid1.ValorCelda(s, 6)
End If
'Call QuitarCaracter(d)
'd = CArac

If Val(PUnit) = 0 Then PUnit = 0
If Val(Cant) = 0 Then Cant = 0
If Val(IIva) = 0 Then IIva = 0
If Val(Descu) = 0 Then Descu = 0
'DMGrid1.ValorCelda(s, 5) = (PUnit * Cant - Descu) * (IIva / 100)
DMGrid1.ValorCelda(s, 7) = (PUnit * Cant - Descu) + IIva
DMGrid1.ValorCelda(s, 1) = DMGrid1.ValorCelda(DMGrid1.Row, 1)
If s <> DMGrid1.Row Then DMGrid1.ValorCelda(DMGrid1.Row, 1) = ""
DMGrid1.PaintMGrid
calcular

End Sub

Private Sub DMGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And DMGrid1.Col = 1 Then 'tecla F1
f = DMGrid1.Row
Tipo = "Compras"
FrmListadoProductosServicios.Show vbModal, FrmPrincipal

'DMGrid1.ValorCelda(f, 1) = CodProd
'DMGrid1.ValorCelda(f, 2) = DescPro
'DMGrid1.ValorCelda(f, 4) = PreProd
'DMGrid1.ValorCelda(f, 5) = IvaProd
DMGrid1.Col = 3
'DMGrid1.EditActive = True
DMGrid1.Editable = True
Call DMGrid1.PaintMGrid

End If

'If KeyAscii = 13 And DMGrid1.Col = 3 Then
'    celvacia = DMGrid1.ValorCelda(lRow, 3)
'    If celvacia = "" Then
'
'    Else
'        BtnAgregarRenglon.SetFocus
'    End If
'End If
End Sub

Private Sub DTPickerFechaEmision_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPickerFechaRecepcion.SetFocus
        Case vbKeyUp
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyLeft
            CboCondicionPago.SetFocus
        Case vbKeyRight
            DTPickerFechaRecepcion.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

 
Private Sub DTPickerFechaRecepcion_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DMGrid1.SetFocus
        Case vbKeyUp
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyLeft
            DTPickerFechaEmision.SetFocus
        Case vbKeyRight
            BtnImportarOrdenCompra.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub Form_Activate()
Tipo = "Compras"
TxtCodigoProveedor.SetFocus
DTPickerFechaEmision.Value = Now
DTPickerFechaRecepcion.Value = Now
DMGrid1.RowBackColor 1, RGB(255, 255, 255)

End Sub

Private Sub Form_Load()
Centrar Me
Grid1
NoOrden
Impuestos
CondicionPago
opcion = 1
TxtCantidadRenglon.Text = DMGrid1.Rows
BtnAgregarRenglon.Enabled = False
BtnEliminarRenglon.Enabled = False
BtnProcesada.Enabled = False
BtnNuevo.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = False
BtnImprimir.Enabled = False
End Sub
Sub CondicionPago()

Dim RsFormaPago As New ADODB.Recordset
' carga los tipos de opciones de pago de la factura en un combo
CSql = "SELECT * FROM Pago where activo=1"
Set RsFormaPago = CrearRS(CSql)

If RsFormaPago.EOF Then RsFormaPago.Close: Exit Sub
    RsFormaPago.MoveFirst
    CboCondicionPago.Clear
Do While Not RsFormaPago.EOF
    CboCondicionPago.AddItem RsFormaPago.Fields("Tipo").Value
    CboCondicionPago.ItemData(CboCondicionPago.NewIndex) = RsFormaPago.Fields("id").Value
    RsFormaPago.MoveNext
Loop

RsFormaPago.Close

End Sub
Sub NoOrden()

Dim RsDatoSAdmin As New ADODB.Recordset
CSql = "Select max(U_Compra)+1 as MaxNoCompra From Dat_admin"
Set RsDatoSAdmin = CrearRS(CSql)

If RsDatoSAdmin.BOF = False Or RsDatoSAdmin.EOF = False Then
    LblNoOrden.Caption = Format(RsDatoSAdmin.Fields("MaxNoCompra").Value, "0000")
End If

RsDatoSAdmin.Close
End Sub

Sub Impuestos()
Dim RsDatoSAdmin As New ADODB.Recordset
CSql = "Select * From Dat_admin"
Set RsDatoSAdmin = CrearRS(CSql)

If RsDatoSAdmin.BOF = False Or RsDatoSAdmin.EOF = False Then
    LblValorImpuesto.Caption = "I.V.A.: " & "(" & RsDatoSAdmin.Fields("Iva1").Value & "%)"
    Impuesto = RsDatoSAdmin.Fields("Iva1").Value
End If
    
RsDatoSAdmin.Close
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
            BtnBuscar.SetFocus
        Case vbKeyUp
            BtnAgregarRenglon.SetFocus
        Case vbKeyRight
            BtnBuscar.SetFocus
    End Select
End If
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

Private Sub TxtCodigoProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyRight
            BtnBuscarProveedor.SetFocus
        Case vbKeyDown
            TxtNoFactura.SetFocus
    End Select
End If
If KeyCode = 112 Then
    p = IdProveedor
    FrmListadoProveedor.Show
    If p <> IdProveedor Then
        Carga_Proveedor
    End If
End If
End Sub

Sub Carga_Proveedor()
Dim RsCargarProveedor As New ADODB.Recordset
'datos de cliente
CSql = "Select * From Proveedores Where IdProveedor = '" & IdProveedor & "'"
Set RsCargarProveedor = CrearRS(CSql)

If Not RsCargarCliente.EOF Then
    TxtCodigoProveedor.Text = RsCargarProveedor.Fields("IdProveedor").Value
    TxtDescripcionProveerdor.Text = RsCargarProveedor.Fields("Nombre").Value
    TxtRif.Text = RsCargarProveedor.Fields("RifProveedor").Value
    RsCargarProveedor.Close
Else
    RsCargarProveedor.Close
End If

End Sub

Private Sub TxtCodigoProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtCodigoProveedor.Text = "" Then
        TxtCodigoProveedor.SetFocus
        Exit Sub
    Else
        CSql = "Select * From Proveedores Where IdProveedor = '" & Trim(TxtCodigoProveedor.Text) & "'"
        Dim RsBuscarProveedor As New ADODB.Recordset
        Set RsBuscarProveedor = CrearRS(CSql)
        If RsBuscarProveedor.EOF = False Or RsBuscarProveedor.BOF = False Then
           TxtDescripcionProveerdor.Text = RsBuscarProveedor.Fields("Nombre").Value
           TxtRif.Text = RsBuscarProveedor.Fields("RifProveedor").Value
        End If
        RsBuscarProveedor.Close
        BtnNuevo.SetFocus
    End If
End If

End Sub
Sub Grid1()
    ' carga las columnas y encabezados de columna
    DMGrid1.Cols = 7
    DMGrid1.Rows = 0
    DMGrid1.DColumnas(1).Alignment = 1
    DMGrid1.DColumnas(3).Alignment = 1
    DMGrid1.DColumnas(4).Alignment = 1
    DMGrid1.DColumnas(5).Alignment = 1
    DMGrid1.DColumnas(6).Alignment = 1
    DMGrid1.DColumnas(7).Alignment = 1
    DMGrid1.DColumnas(2).Locked = True
    DMGrid1.DColumnas(4).Locked = True
    DMGrid1.DColumnas(5).Locked = True
    DMGrid1.DColumnas(7).Locked = True
    DMGrid1.DColumnas(4).IsNumber = True
    DMGrid1.DColumnas(5).IsNumber = True
    DMGrid1.DColumnas(6).IsNumber = True
    DMGrid1.DColumnas(7).IsNumber = True
    DMGrid1.DColumnas(1).Width = 1200
    DMGrid1.DColumnas(2).Width = 7390
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

Sub calcular()
Dim Cant, PUnit, IIva, Descu, e As Double
Dim SubTot As Double

For i = 1 To DMGrid1.Rows
    
    IdProd = DMGrid1.ValorCelda(i, 1)
   
    If IdProd <> "" Then
        Dim RsAplicaIva As New ADODB.Recordset
        Dim AplicaIva
    
        CSql = "Select * From Productos Where IdProducto ='" & IdProd & "'"
        Set RsAplicaIva = CrearRS(CSql)
        AplicaIva = RsAplicaIva.Fields("Impuesto").Value
        RsAplicaIva.Close
        
        If Val(DMGrid1.ValorCelda(i, 3)) <> 0 Then
            Cant = CDbl(DMGrid1.ValorCelda(i, 3))
        Else
            Cant = 1
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
        
        If AplicaIva Then
             DMGrid1.ValorCelda(i, 5) = "0,00"
            If Val(DMGrid1.ValorCelda(i, 5)) <> 0 Then
                IIva = CDbl(DMGrid1.ValorCelda(i, 5))
            Else
                IIva = 0
            End If
        Else
            IIva = 0
        End If
        
        If Not IsNull(DMGrid1.ValorCelda(i, 5)) Then IIva = CDbl(DMGrid1.ValorCelda(i, 5)) Else IIva = 0
        'QuitarCaracter (f)
        'f = CArac
        DMGrid1.ValorCelda(i, 6) = "0,00"
        If Val(DMGrid1.ValorCelda(i, 6)) <> 0 Then
            Descu = DMGrid1.ValorCelda(i, 6)
        Else
            Descu = 0
        End If
        'QuitarCaracter (d)
        'd = CArac
        
        If IsNull(Cant) Then Cant = 0
        If IsNull(PUnit) Then PUnit = 0
        If IsNull(Descu) Then Descu = 0
        If IsNull(IIva) Then IIva = 0
        
        SubTot = SubTot + ((PUnit * Cant) - Descu)
        'e = e + f / 100 * ((a * b) - d)
        If Not IsNull(DMGrid1.ValorCelda(i, 5)) Then
            e = e + CDbl(DMGrid1.ValorCelda(i, 5))
        Else
            e = e
        End If
        'e = e + f / 100 * c
    End If
Next i

If Val(SubTot) = 0 Then SubTot = 0
If Val(e) = 0 Then e = 0

'SubTotal
TxtSubTotal.Text = Format(SubTot, "###,##0.#0")

'Impuesto
TxtImpuesto.Text = Format(e, "###,##0.#0")

'Total General
TxtTotalGeneral.Text = Format(SubTot + e, "###,##0.#0")
End Sub

Private Sub TxtDescripcionProveerdor_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtRif.SetFocus
        Case vbKeyLeft
            BtnBuscarProveedor.SetFocus
        Case vbKeyRight
            TxtRif.SetFocus
        Case vbKeyDown
            TxtNoControl.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoControl_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoOrdenCompra.SetFocus
        Case vbKeyUp
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyLeft
            TxtNoFactura.SetFocus
        Case vbKeyRight
            TxtNoOrdenCompra.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub TxtNoFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoControl.SetFocus
        Case vbKeyUp
            TxtCodigoProveedor.SetFocus
        Case vbKeyRight
            TxtNoControl.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

 

Private Sub TxtNoOrdenCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboCondicionPago.SetFocus
        Case vbKeyUp
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyLeft
            TxtNoControl.SetFocus
        Case vbKeyRight
            CboCondicionPago.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub TxtRif_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNoFactura.SetFocus
        Case vbKeyLeft
            TxtDescripcionProveerdor.SetFocus
        Case vbKeyRight
            BtnProveedores.SetFocus
        Case vbKeyDown
            DTPickerFechaRecepcion.SetFocus
    End Select
End If
End Sub

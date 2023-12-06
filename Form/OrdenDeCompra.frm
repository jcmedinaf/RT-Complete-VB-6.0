VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmOrdenCompra 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Órdenes de Compras"
   ClientHeight    =   8010
   ClientLeft      =   2085
   ClientTop       =   1470
   ClientWidth     =   15345
   Icon            =   "OrdenDeCompra.frx":0000
   LinkTopic       =   "Form56"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   15345
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3720
         TabIndex        =   31
         Top             =   7080
         Width           =   11295
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   10200
            TabIndex        =   32
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
            MICON           =   "OrdenDeCompra.frx":1002
            PICN            =   "OrdenDeCompra.frx":101E
            PICH            =   "OrdenDeCompra.frx":11E7
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
            TabIndex        =   33
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
            MICON           =   "OrdenDeCompra.frx":141C
            PICN            =   "OrdenDeCompra.frx":1438
            PICH            =   "OrdenDeCompra.frx":16C7
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
            TabIndex        =   34
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
            MICON           =   "OrdenDeCompra.frx":1B08
            PICN            =   "OrdenDeCompra.frx":1B24
            PICH            =   "OrdenDeCompra.frx":1CB1
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
            TabIndex        =   35
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
            MICON           =   "OrdenDeCompra.frx":1EE6
            PICN            =   "OrdenDeCompra.frx":1F02
            PICH            =   "OrdenDeCompra.frx":21E4
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
            Left            =   4680
            TabIndex        =   36
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "OrdenDeCompra.frx":2435
            PICN            =   "OrdenDeCompra.frx":2451
            PICH            =   "OrdenDeCompra.frx":2576
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
            TabIndex        =   37
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
            MICON           =   "OrdenDeCompra.frx":2806
            PICN            =   "OrdenDeCompra.frx":2822
            PICH            =   "OrdenDeCompra.frx":29C6
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
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   28
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
            Left            =   840
            TabIndex        =   29
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el Número de la orden de Compra a Buscar"
            Top             =   240
            Width           =   1575
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2520
            TabIndex        =   38
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
            MICON           =   "OrdenDeCompra.frx":2B65
            PICN            =   "OrdenDeCompra.frx":2B81
            PICH            =   "OrdenDeCompra.frx":2DE6
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
            Caption         =   "Nº Órden:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   330
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Proveedor"
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14895
         Begin ChamaleonButton.ChameleonBtn BtnProcesada 
            Height          =   375
            Left            =   10440
            TabIndex        =   46
            ToolTipText     =   "Procesar Consumos de Medicamentos"
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
            MICON           =   "OrdenDeCompra.frx":3078
            PICN            =   "OrdenDeCompra.frx":3094
            PICH            =   "OrdenDeCompra.frx":3309
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtCodigoProveedor 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   480
            Width           =   1215
         End
         Begin ChamaleonButton.ChameleonBtn BtnProductos 
            Height          =   375
            Left            =   13200
            TabIndex        =   42
            Top             =   240
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
            MICON           =   "OrdenDeCompra.frx":3585
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnProveedores 
            Height          =   375
            Left            =   11520
            TabIndex        =   41
            Top             =   240
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
            MICON           =   "OrdenDeCompra.frx":35A1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Órden Procesada"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11880
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Detalle"
            Height          =   5175
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   14655
            Begin SystemOncoAmerica.DMGrid DMGrid1 
               Height          =   4335
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   7646
               Object.Width           =   14385
               Object.Height          =   4305
               ScrollBar       =   1
               Editable        =   -1  'True
               DrawColorGrid   =   1
            End
            Begin VB.TextBox TxtCantidadRenglon 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Left            =   4920
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   4680
               Width           =   855
            End
            Begin VB.Timer Timer1 
               Interval        =   1000
               Left            =   6000
               Top             =   4680
            End
            Begin VB.TextBox TxtImpuesto 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   9960
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "0.00"
               Top             =   4680
               Width           =   1695
            End
            Begin VB.TextBox TxtSubTotal 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "0.00"
               Top             =   4680
               Width           =   1695
            End
            Begin VB.TextBox TxtTotalGeneral 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   12840
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "0.00"
               Top             =   4680
               Width           =   1695
            End
            Begin ChamaleonButton.ChameleonBtn BtnEliminarRenglon 
               Height          =   375
               Left            =   1080
               TabIndex        =   22
               ToolTipText     =   "Eliminar Renglon"
               Top             =   4680
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
               MICON           =   "OrdenDeCompra.frx":35BD
               PICN            =   "OrdenDeCompra.frx":35D9
               PICH            =   "OrdenDeCompra.frx":377D
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
               TabIndex        =   23
               ToolTipText     =   "Agregar Renglon"
               Top             =   4680
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
               MICON           =   "OrdenDeCompra.frx":391C
               PICN            =   "OrdenDeCompra.frx":3938
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad Renglones:"
               Height          =   195
               Left            =   3360
               TabIndex        =   40
               Top             =   4770
               Width           =   1485
            End
            Begin VB.Label LblValorImpuesto 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "(12%)"
               Height          =   195
               Left            =   9480
               TabIndex        =   25
               Top             =   4770
               Width           =   390
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal:"
               Height          =   195
               Left            =   6480
               TabIndex        =   21
               Top             =   4770
               Width           =   690
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
               Height          =   195
               Left            =   9120
               TabIndex        =   20
               Top             =   4770
               Width           =   255
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total General:"
               Height          =   195
               Left            =   11760
               TabIndex        =   19
               Top             =   4770
               Width           =   1005
            End
         End
         Begin VB.TextBox TxtStatus 
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   2055
         End
         Begin VB.ComboBox CboCondicionPago 
            Height          =   315
            Left            =   2280
            TabIndex        =   8
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox TxtDescripcionProveerdor 
            Height          =   375
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox TxtRif 
            Height          =   375
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   3255
         End
         Begin MSComCtl2.DTPicker DTPickerFechaEmision 
            Height          =   375
            Left            =   5400
            TabIndex        =   9
            Top             =   1080
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            Format          =   54132737
            CurrentDate     =   39932
         End
         Begin MSComCtl2.DTPicker DTPickerFechaRecepcion 
            Height          =   375
            Left            =   7920
            TabIndex        =   11
            Top             =   1080
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            Format          =   54132737
            CurrentDate     =   39932
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarProveedor 
            Height          =   375
            Left            =   1440
            TabIndex        =   44
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
            MICON           =   "OrdenDeCompra.frx":3D5B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            Left            =   12960
            TabIndex        =   27
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Órden No."
            Height          =   195
            Left            =   12120
            TabIndex        =   26
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Recepción:"
            Height          =   195
            Left            =   7920
            TabIndex        =   12
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Emisión:"
            Height          =   195
            Left            =   5400
            TabIndex        =   10
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición de Pago:"
            Height          =   195
            Left            =   2280
            TabIndex        =   7
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
            Height          =   195
            Left            =   2280
            TabIndex        =   5
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif.: "
            Height          =   195
            Left            =   8160
            TabIndex        =   4
            Top             =   240
            Width           =   330
         End
      End
   End
End
Attribute VB_Name = "FrmOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsUpDateDatAdmin As New ADODB.Recordset
Dim Impuesto As Double
Private Sub BtnAgregarRenglon_Click()
If Check1.Value = 0 Then
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
    Call DMGrid1.PaintMGrid
Else
    Msg = "Ya esta Órden de Compra fue procesada y no puede ser modificada"
    MsgBox Msg, vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
End If
End Sub
Sub Blanqueo()
DMGrid1.Rows = 0
DMGrid1.Clear
TxtCodigoProveedor.Text = ""
TxtDescripcionProveerdor.Text = ""
TxtRif.Text = ""
DTPickerFechaEmision.Value = Now
DTPickerFechaRecepcion.Value = Now
TxtStatus.Text = ""
CboCondicionPago.Text = ""
CboCondicionPago.ListIndex = -1
TxtCantidadRenglon.Text = 0
TxtSubTotal.Text = "0.00"
TxtImpuesto.Text = "0.00"
TxtTotalGeneral.Text = "0.00"
Check1.Value = 0
End Sub


Private Sub BtnBuscar_Click()
opcion = 1
CSql = "Select * From Ordenes Where NumeroOrden = '" & Trim(TxtBuscar.Text) & "'"
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
    BtnProcesada.Enabled = False
    Check1.Value = 1
Else
    BtnNuevo.Enabled = False
    BtnGuardar.Enabled = True
    BtnBorrar.Enabled = False
    BtnImprimir.Enabled = True
    BtnAgregarRenglon.Enabled = True
    BtnEliminarRenglon.Enabled = True
    BtnProcesada.Enabled = True
    Check1.Value = 0
End If

If RsBuscarOrdenes.EOF = False Or RsBuscarOrdenes.BOF = False Then
    LblNoOrden.Caption = RsBuscarOrdenes.Fields("NumeroOrden").Value
    TxtStatus.Text = RsBuscarOrdenes.Fields("Status").Value
    CboCondicionPago.Text = RsBuscarOrdenes.Fields("CondicionPago").Value
    DTPickerFechaEmision.Value = RsBuscarOrdenes.Fields("FechaEmision").Value
    DTPickerFechaRecepcion.Value = RsBuscarOrdenes.Fields("FechaRecepcion").Value
    TxtSubTotal.Text = Format(RsBuscarOrdenes.Fields("SubTotal").Value, "#,##0.00")
    TxtImpuesto.Text = Format(RsBuscarOrdenes.Fields("Impuesto").Value, "#,##0.00")
    TxtTotalGeneral.Text = Format(RsBuscarOrdenes.Fields("TotalGeneral").Value, "#,##0.00")
    If RsBuscarOrdenes.Fields("OrdenProcesada").Value = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    CodProveedor = RsBuscarOrdenes.Fields("IdProveedor").Value
End If

CSql = "Select * From Proveedores Where IdProveedor = '" & Trim(CodProveedor) & "'"
Dim RsBuscarProveedor As New ADODB.Recordset
Set RsBuscarProveedor = CrearRS(CSql)

TxtCodigoProveedor.Text = RsBuscarProveedor.Fields("IdProveedor").Value
TxtDescripcionProveerdor.Text = RsBuscarProveedor.Fields("Nombre").Value
TxtRif.Text = RsBuscarProveedor.Fields("RifProveedor").Value

CSql = "Select * From RenglonOrdenes Where NumeroOrden = '" & Trim(TxtBuscar.Text) & "'"
Dim RsBuscarRenglonOrden As New ADODB.Recordset
Set RsBuscarRenglonOrden = CrearRS(CSql)

i = 1
If Not (RsBuscarRenglonOrden.EOF) Then
    RsBuscarRenglonOrden.MoveFirst
    Dim RsProductos As New ADODB.Recordset
    Do While Not RsBuscarRenglonOrden.EOF
        DMGrid1.Rows = i
        CSql = "Select * From Productos Where IdProducto ='" & RsBuscarRenglonOrden.Fields("IdProducto").Value & "'"
        Set RsProductos = CrearRS(CSql)
        DMGrid1.ValorCelda(i, 1) = RsBuscarRenglonOrden.Fields("IdProducto").Value
        DMGrid1.ValorCelda(i, 2) = RsProductos.Fields("Descripcion").Value
        DMGrid1.ValorCelda(i, 3) = RsBuscarRenglonOrden.Fields("Cantidad").Value
        DMGrid1.ValorCelda(i, 4) = RsBuscarRenglonOrden.Fields("precio").Value
        DMGrid1.ValorCelda(i, 5) = RsBuscarRenglonOrden.Fields("impuesto").Value
        DMGrid1.ValorCelda(i, 6) = RsBuscarRenglonOrden.Fields("descuento").Value
        DMGrid1.ValorCelda(i, 7) = RsBuscarRenglonOrden.Fields("SubTotal").Value
        
        If RsProductos.Fields("Impuesto") Then
            DMGrid1.RowBackColor i, RGB(255, 255, 255)
        Else
            DMGrid1.RowBackColor i, RGB(221, 221, 221)
        End If
        RsProductos.Close
        i = i + 1
        RsBuscarRenglonOrden.MoveNext
    Loop
       TxtCantidadRenglon.Text = DMGrid1.Rows
Else
    DMGrid1.Clear
    Call DMGrid1.PaintMGrid
    Text6.Text = Format(0, Standard)
    Text7.Text = Format(0, Standard)
    Text8.Text = Format(0, Standard)
    RsBuscarRenglonOrden.Close
    RsBuscarFactura.Close
    Exit Sub
End If
DMGrid1.PaintMGrid
RsBuscarOrdenes.Close
RsBuscarProveedor.Close
calcular
BtnNuevo.Enabled = True
BtnNuevo.SetFocus
End Sub

Private Sub BtnBuscarProveedor_Click()
Tipo = "Ordenes"
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
End Sub

Private Sub BtnEliminarRenglon_Click()
If Check1.Value = 0 Then
    DMGrid1.RowDelete (DMGrid1.Row)
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
If TxtStatus.Text = "" Then
    MsgBox "Esta dejando el Estatus del Proveedor en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If
If CboCondicionPago.Text = "" Then
    MsgBox "Esta dejando la Condicion de Pago en Blanco", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

If DMGrid1.Rows <= 0 Then
    MsgBox "No hay Productos a Guardar", vbOKOnly + vbCritical, "Mensaje"
    Exit Sub
End If

Select Case opcion
    Case Is = 0
        Dim RsGuardarOrden As New ADODB.Recordset
        CSql = "Select * From Ordenes"
        Set RsGuardarOrden = CrearRS(CSql)
                
        RsGuardarOrden.AddNew
        RsGuardarOrden.Fields("NumeroOrden").Value = LblNoOrden.Caption
        RsGuardarOrden.Fields("IdProveedor").Value = Trim(TxtCodigoProveedor.Text)
        RsGuardarOrden.Fields("Status").Value = Trim(TxtStatus.Text)
        RsGuardarOrden.Fields("CondicionPago").Value = Trim(CboCondicionPago.Text)
        RsGuardarOrden.Fields("FechaEmision").Value = Format(DTPickerFechaEmision.Value, "dd/mm/yyyy")
        RsGuardarOrden.Fields("FechaRecepcion").Value = Format(DTPickerFechaRecepcion.Value, "dd/mm/yyyy")
        RsGuardarOrden.Fields("OrdenProcesada").Value = Check1.Value
        RsGuardarOrden.Fields("SubTotal").Value = Trim(TxtSubTotal.Text)
        RsGuardarOrden.Fields("Impuesto").Value = Trim(TxtImpuesto.Text)
        RsGuardarOrden.Fields("TotalGeneral").Value = Trim(TxtTotalGeneral.Text)
        RsGuardarOrden.Update
        
        Dim RsGuardarRenglonOrdenes As New ADODB.Recordset
        CSql = "Select * From RenglonOrdenes"
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
            RsGuardarRenglonOrdenes.Fields("NumeroOrden").Value = LblNoOrden.Caption
            RsGuardarRenglonOrdenes.Fields("NumeroRenglon").Value = i
            RsGuardarRenglonOrdenes.Fields("IdProducto").Value = Trim(b1)
            RsGuardarRenglonOrdenes.Fields("Cantidad").Value = b3
            RsGuardarRenglonOrdenes.Fields("Precio").Value = b2
            RsGuardarRenglonOrdenes.Fields("Impuesto").Value = b4
            RsGuardarRenglonOrdenes.Fields("Descuento").Value = b5
            RsGuardarRenglonOrdenes.Fields("SubTotal").Value = Trim(b6)
            RsGuardarRenglonOrdenes.Update
r:
        TxtCantidadRenglon.Text = DMGrid1.Rows
        Next i
        
        CSql = "Select * From Dat_Admin"
        Set RsUpDateDatAdmin = CrearRS(CSql)
        
        RsUpDateDatAdmin.Fields("U_Orden").Value = Val(LblNoOrden.Caption)
        
        RsUpDateDatAdmin.Update
        RsUpDateDatAdmin.Close
        NoOrden
    
    Case Is = 1
        'Dim RsGuardarOrden As New ADODB.Recordset
        CSql = "Select * From Ordenes Where NumeroOrden='" & Trim(LblNoOrden.Caption) & "'"
        Set RsGuardarOrden = CrearRS(CSql)
                
        RsGuardarOrden.Fields("NumeroOrden").Value = LblNoOrden.Caption
        RsGuardarOrden.Fields("IdProveedor").Value = Trim(TxtCodigoProveedor.Text)
        RsGuardarOrden.Fields("Status").Value = Trim(TxtStatus.Text)
        RsGuardarOrden.Fields("CondicionPago").Value = Trim(CboCondicionPago.Text)
        RsGuardarOrden.Fields("FechaEmision").Value = FechaSQL(DTPickerFechaEmision.Value)
        RsGuardarOrden.Fields("FechaRecepcion").Value = FechaSQL(DTPickerFechaRecepcion.Value)
        RsGuardarOrden.Fields("OrdenProcesada").Value = Check1.Value
        RsGuardarOrden.Fields("SubTotal").Value = Trim(TxtSubTotal.Text)
        RsGuardarOrden.Fields("Impuesto").Value = Trim(TxtImpuesto.Text)
        RsGuardarOrden.Fields("TotalGeneral").Value = Trim(TxtTotalGeneral.Text)
        RsGuardarOrden.Update
        
        'Dim RsGuardarRenglonOrdenes As New ADODB.Recordset
        CSql = "Select * From RenglonOrdenes Where NumeroOrden='" & Trim(LblNoOrden.Caption) & "'"
        Set RsGuardarRenglonOrdenes = CrearRS(CSql)
                
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
               
            RsGuardarRenglonOrdenes.Fields("NumeroOrden").Value = LblNoOrden.Caption
            RsGuardarRenglonOrdenes.Fields("NumeroRenglon").Value = i
            RsGuardarRenglonOrdenes.Fields("IdProducto").Value = Trim(b1)
            RsGuardarRenglonOrdenes.Fields("Cantidad").Value = Trim(b3)
            RsGuardarRenglonOrdenes.Fields("Precio").Value = b2
            RsGuardarRenglonOrdenes.Fields("Impuesto").Value = b4
            RsGuardarRenglonOrdenes.Fields("Descuento").Value = b5
            RsGuardarRenglonOrdenes.Fields("SubTotal").Value = Trim(b6)
            RsGuardarRenglonOrdenes.Update
            
            TxtCantidadRenglon.Text = DMGrid1.Rows
rr:
        Next i
End Select
MsgBox "La Orden de Compra fue creada satisfactoriamente!", vbInformation + vbOKOnly, "Operación Exitosa"

BtnAgregarRenglon.Enabled = False
BtnEliminarRenglon.Enabled = False
BtnProcesada.Enabled = False
BtnNuevo.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = False
BtnImprimir.Enabled = False

Blanqueo
NoOrden
End Sub

Private Sub BtnImportar_Click()

End Sub

Private Sub BtnImprimir_Click()

    If Check1.Value = 0 Then
        Dim RsActualizarImpresa As New ADODB.Recordset
        'colcar aqui el codigo que levante el informe de la factura
        imprime
               
        CSql = "Update Ordenes set OrdenProcesada = 1 WHERE NumeroOrden = " & LblNoOrden.Caption
        Set RsActualizarImpresa = CrearRS(CSql)
        Call Enviar_Bitacora(IdUser, "COMPRAS", "IMPRIMIR", "Se imprimio la Compra Nro. " & LblNoOrden.Caption)
    Else
        Msg = "Ya esta factura de Compra fue impresa desea imprimir una copia ?"
        d = MsgBox(Msg, vbYesNo, "Reimpimir Factura Compra")
        If d = 6 Then
            imprime
        End If
    End If

End Sub
Sub imprime()
''========= ESTE ES EL CODIGO NUEVO ==========
If DMGrid1.Rows > 0 Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\OrdenesCompras.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{OrdenesDeCompras.NumeroOrden} = '" & LblNoOrden.Caption & "'"
        .WindowTitle = "Reporte Orden de Compras No. " & LblNoOrden.Caption
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If
End Sub

Private Sub BtnNuevo_Click()
If Check1.Value = 1 Then
    mensaje = MsgBox("Esta Orden de Compra fue Procesada y no se le pueden agregar mas items." & Chr(13) & "Deseas crear una nueva orden de compra?", vbYesNo + vbInformation, "Mensaje")
    
    If mensaje = vbYes Then
        Blanqueo
        opcion = 0
        BtnAgregarRenglon.Enabled = True
        BtnEliminarRenglon.Enabled = True
        BtnProcesada.Enabled = True
        BtnNuevo.Enabled = False
        BtnGuardar.Enabled = True
        BtnBorrar.Enabled = False
        BtnImprimir.Enabled = False
                
        CSql = "Select * From Dat_Admin"
        Set RsUpDateDatAdmin = CrearRS(CSql)
        
        RsUpDateDatAdmin.Fields("U_Orden").Value = Val(LblNoOrden.Caption)
        
        RsUpDateDatAdmin.Update
        RsUpDateDatAdmin.Close
        NoOrden
    End If
Else
    opcion = 0
    BtnAgregarRenglon.Enabled = True
    BtnEliminarRenglon.Enabled = True
    BtnProcesada.Enabled = True
    BtnNuevo.Enabled = False
    BtnGuardar.Enabled = True
    BtnBorrar.Enabled = False
    BtnImprimir.Enabled = False
End If
End Sub

Private Sub BtnProcesada_Click()
Dim RsProcesar As New ADODB.Recordset
Msg = "Estas Seguro(a) de Procesar la Orden de Compra!!!"
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Procesar Orden de Compra")

If mensaje = vbYes Then
    
    CSql = "Update Ordenes Set OrdenProcesada='1' where NumeroOrden='" & Val(LblNoOrden.Caption) & "'"
    Set RsProcesar = CrearRS(CSql)
    
    Check1.Value = 1
    BtnProcesada.Enabled = False
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    BtnAgregar.Enabled = False
    BtnAgregarRenglon.Enabled = False
    BtnBorrarRenglon.Enabled = False
    
    Msg = "Orden de Compra Procesada Correctamente!!!"
    mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Procesar Orden de Compras")
    
End If
End Sub

Private Sub BtnProductos_Click()
FrmProductos.Show
End Sub

Private Sub BtnProveedores_Click()
FrmProveedores.Show
End Sub

Private Sub BtnSiguiente_Click()

End Sub

Private Sub DMGrid1_AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
Dim PUnit, Cant, IIva, Descu As Double
Dim CalculaIva As Boolean

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
    DMGrid1.ValorCelda(s, 5) = PUnit * Cant * (Impuesto / 100)
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
    IIva = Impuesto
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
'DMGrid1.ValorCelda(s, 5) = (PUnit * Cant - Descu) * (Impuesto / 100)
DMGrid1.ValorCelda(s, 7) = (PUnit * Cant - Descu) + IIva
DMGrid1.ValorCelda(s, 1) = DMGrid1.ValorCelda(DMGrid1.Row, 1)
If s <> DMGrid1.Row Then DMGrid1.ValorCelda(DMGrid1.Row, 1) = ""
DMGrid1.PaintMGrid
calcular
End Sub

Private Sub DMGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And DMGrid1.Col = 1 Then 'tecla F1
f = DMGrid1.Row
Tipo = "Ordenes"
FrmListadoProductosServicios.Show vbModal, FrmPrincipal

'DMGrid1.ValorCelda(f, 1) = CodProd
'DMGrid1.ValorCelda(f, 2) = DescPro
'DMGrid1.ValorCelda(f, 4) = PreProd
'DMGrid1.ValorCelda(f, 5) = IvaProd
DMGrid1.Col = 3
DMGrid1.Editable = True
Call DMGrid1.PaintMGrid

End If
End Sub

Private Sub Form_Activate()
Tipo = "Ordenes"
TxtCodigoProveedor.SetFocus
DTPickerFechaEmision.Value = Now
DTPickerFechaRecepcion.Value = Now
End Sub

Private Sub Form_Load()
Centrar Me
Grid1
NoOrden
Impuestos
CondicionPago
opcion = 1


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
CSql = "Select max(U_Orden)+1 as MaxNoOrden From Dat_admin"
Set RsDatoSAdmin = CrearRS(CSql)

If RsDatoSAdmin.BOF = False Or RsDatoSAdmin.EOF = False Then
    LblNoOrden.Caption = Format(RsDatoSAdmin.Fields("MaxNoOrden").Value, "0000")
End If

RsDatoSAdmin.Close
End Sub

Sub Impuestos()
Dim RsDatoSAdmin As New ADODB.Recordset
CSql = "Select * From Dat_admin"
Set RsDatoSAdmin = CrearRS(CSql)

If RsDatoSAdmin.BOF = False Or RsDatoSAdmin.EOF = False Then
    LblValorImpuesto.Caption = "(" & RsDatoSAdmin.Fields("Iva1").Value & "%)"
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
       ' BtnNuevo.SetFocus
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
    
    DMGrid1.DColumnas(1).Locked = False
    DMGrid1.DColumnas(2).Locked = True
    DMGrid1.DColumnas(3).Locked = False
    DMGrid1.DColumnas(4).Locked = False
    DMGrid1.DColumnas(5).Locked = False
    DMGrid1.DColumnas(6).Locked = False
    DMGrid1.DColumnas(7).Locked = True
    
    DMGrid1.DColumnas(4).IsNumber = True
    DMGrid1.DColumnas(5).IsNumber = True
    DMGrid1.DColumnas(6).IsNumber = True
    DMGrid1.DColumnas(7).IsNumber = True
    DMGrid1.DColumnas(1).Width = 1200
    DMGrid1.DColumnas(2).Width = 7720
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

        If Not IsNull(DMGrid1.ValorCelda(i, 3)) Then
            If Val(DMGrid1.ValorCelda(i, 3)) <> 0 Then
                Cant = CDbl(DMGrid1.ValorCelda(i, 3))
            Else
                Cant = 1
            End If
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
      
        If AplicaIva Then
            If Val(DMGrid1.ValorCelda(i, 5)) <> 0 Then
                IIva = CDbl(DMGrid1.ValorCelda(i, 5))
            Else
                IIva = 0
            End If
        Else
            IIva = 0
        End If
        
        'QuitarCaracter (f)
        'f = CArac
        If Not IsNull(DMGrid1.ValorCelda(i, 6)) Then
            If Val(DMGrid1.ValorCelda(i, 6)) <> 0 Then
                Descu = CDbl(DMGrid1.ValorCelda(i, 6))
            Else
                Descu = 0
            End If
        Else
            Descu = 0
        End If
        ' Validar que no esten en blanco
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
        'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
        
        SubTot = SubTot + ((PUnit * Cant) - Descu)
        If Not IsNull(DMGrid1.ValorCelda(i, 5)) Then
            e = e + DMGrid1.ValorCelda(i, 5)
        Else
            e = e
        End If
        'e = e + f / 100 * ((a * b) - d)
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

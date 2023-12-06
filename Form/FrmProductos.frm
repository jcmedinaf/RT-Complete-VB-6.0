VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmProductos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   9450
   ClientLeft      =   3510
   ClientTop       =   2085
   ClientWidth     =   9330
   Icon            =   "FrmProductos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   9330
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   8520
         Width           =   8895
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   7800
            TabIndex        =   36
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
            MICON           =   "FrmProductos.frx":1002
            PICN            =   "FrmProductos.frx":101E
            PICH            =   "FrmProductos.frx":11E7
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
            Left            =   1320
            TabIndex        =   37
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
            MICON           =   "FrmProductos.frx":141C
            PICN            =   "FrmProductos.frx":1438
            PICH            =   "FrmProductos.frx":16C7
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
            Left            =   6600
            TabIndex        =   39
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
            MICON           =   "FrmProductos.frx":1B08
            PICN            =   "FrmProductos.frx":1B24
            PICH            =   "FrmProductos.frx":1E06
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
            Left            =   2520
            TabIndex        =   40
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
            MICON           =   "FrmProductos.frx":2057
            PICN            =   "FrmProductos.frx":2073
            PICH            =   "FrmProductos.frx":2217
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
            Left            =   5160
            TabIndex        =   41
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
            MICON           =   "FrmProductos.frx":23B6
            PICN            =   "FrmProductos.frx":23D2
            PICH            =   "FrmProductos.frx":2668
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
            Left            =   4560
            TabIndex        =   42
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
            MICON           =   "FrmProductos.frx":28C7
            PICN            =   "FrmProductos.frx":28E3
            PICH            =   "FrmProductos.frx":2B78
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
            TabIndex        =   38
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmProductos.frx":2DD4
            PICN            =   "FrmProductos.frx":2DF0
            PICH            =   "FrmProductos.frx":2F7D
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
         Height          =   8415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8895
         Begin VB.ComboBox CboProveedor 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   7800
            Width           =   4935
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   5160
            Top             =   7680
         End
         Begin VB.Frame FrameBusqueda 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Filtro de Busqueda"
            Height          =   735
            Left            =   5640
            TabIndex        =   33
            Top             =   7560
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
               ToolTipText     =   "Ingrese el Código o la Descripción del Producto"
               Top             =   240
               Width           =   1575
            End
            Begin ChamaleonButton.ChameleonBtn BtnBuscar 
               Height          =   375
               Left            =   1800
               TabIndex        =   43
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
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
               MICON           =   "FrmProductos.frx":31B2
               PICN            =   "FrmProductos.frx":31CE
               PICH            =   "FrmProductos.frx":3433
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
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Height          =   4335
            Left            =   120
            TabIndex        =   29
            Top             =   3120
            Width           =   8655
            Begin VB.Frame Frame3 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Costos"
               Height          =   1095
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   7695
               Begin VB.TextBox TxtCostoActual 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   6
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Costo Anterior:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Costo Actual:"
                  Height          =   195
                  Left            =   2280
                  TabIndex        =   68
                  Top             =   240
                  Width           =   945
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Costo Promedio:"
                  Height          =   195
                  Left            =   4320
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label TxtCostoAnterior 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "0.00"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   66
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Label TxtCostoPromedio 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "0.00"
                  Height          =   375
                  Left            =   4320
                  TabIndex        =   65
                  Top             =   480
                  Width           =   1935
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Precio 3"
               Height          =   975
               Left            =   120
               TabIndex        =   57
               Top             =   3240
               Width           =   7695
               Begin VB.TextBox TxtPrecioUnitario3 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   5520
                  TabIndex        =   17
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox TxtImpuesto3 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   3480
                  TabIndex        =   16
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox TxtUtilidad3 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   120
                  TabIndex        =   19
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox TxtPrecioSinImpuesto3 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   15
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio Unitario 1:"
                  Height          =   195
                  Left            =   5520
                  TabIndex        =   62
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Impuesto:"
                  Height          =   195
                  Left            =   3480
                  TabIndex        =   61
                  Top             =   240
                  Width           =   690
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Utilidad:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   60
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio Sin Impuesto:"
                  Height          =   195
                  Left            =   1440
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   58
                  Top             =   570
                  Width           =   120
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Precio 2"
               Height          =   975
               Left            =   120
               TabIndex        =   51
               Top             =   2280
               Width           =   7695
               Begin VB.TextBox TxtPrecioUnitario2 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   5520
                  TabIndex        =   14
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox TxtImpuesto2 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   3480
                  TabIndex        =   13
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox TxtUtilidad2 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   120
                  TabIndex        =   11
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox TxtPrecioSinImpuesto2 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   12
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio Unitario 1:"
                  Height          =   195
                  Left            =   5520
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Impuesto:"
                  Height          =   195
                  Left            =   3480
                  TabIndex        =   55
                  Top             =   240
                  Width           =   690
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Utilidad:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   54
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio Sin Impuesto:"
                  Height          =   195
                  Left            =   1440
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   52
                  Top             =   570
                  Width           =   120
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Precio 1"
               Height          =   975
               Left            =   120
               TabIndex        =   45
               Top             =   1320
               Width           =   7695
               Begin VB.TextBox TxtPrecioSinImpuesto1 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   8
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox TxtUtilidad1 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   120
                  TabIndex        =   7
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox TxtImpuesto1 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   3480
                  TabIndex        =   9
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox TxtPrecioUnitario1 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   5520
                  TabIndex        =   10
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio Sin Impuesto:"
                  Height          =   195
                  Left            =   1440
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Utilidad:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   49
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Impuesto:"
                  Height          =   195
                  Left            =   3480
                  TabIndex        =   48
                  Top             =   240
                  Width           =   690
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EAEFEF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio Unitario 1:"
                  Height          =   195
                  Left            =   5520
                  TabIndex        =   47
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   46
                  Top             =   570
                  Width           =   120
               End
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Impuesto"
            Height          =   975
            Left            =   6240
            TabIndex        =   26
            Top             =   2160
            Width           =   2535
            Begin VB.CheckBox Check1 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Si o No"
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   195
               Left            =   960
               TabIndex        =   32
               Top             =   360
               Width           =   120
            End
            Begin VB.Label LblImpuesto 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   600
               TabIndex        =   31
               Top             =   360
               Width           =   45
            End
            Begin VB.Label LblExento 
               BackStyle       =   0  'Transparent
               Caption         =   "Producto Exento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   375
               Left            =   1440
               TabIndex        =   30
               Top             =   360
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "IVA:"
               Height          =   195
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Width           =   300
            End
         End
         Begin VB.ComboBox CboUbicacion 
            Height          =   315
            ItemData        =   "FrmProductos.frx":36C5
            Left            =   3600
            List            =   "FrmProductos.frx":36D2
            TabIndex        =   5
            Top             =   2520
            Width           =   2535
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1200
            Width           =   8655
         End
         Begin VB.ComboBox CboTipo 
            Height          =   315
            ItemData        =   "FrmProductos.frx":36F0
            Left            =   120
            List            =   "FrmProductos.frx":36FA
            TabIndex        =   3
            Top             =   2520
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   2040
            TabIndex        =   4
            Top             =   2520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55443457
            CurrentDate     =   39932
         End
         Begin ChamaleonButton.ChameleonBtn BtnListadoProductos 
            Height          =   375
            Left            =   1560
            TabIndex        =   70
            ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Listado Productos"
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
            MICON           =   "FrmProductos.frx":3712
            PICN            =   "FrmProductos.frx":372E
            PICH            =   "FrmProductos.frx":39B7
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
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   7560
            Width           =   780
         End
         Begin VB.Label NoReg 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Productos:"
            Height          =   195
            Left            =   7320
            TabIndex        =   44
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación"
            Height          =   195
            Left            =   3600
            TabIndex        =   25
            Top             =   2280
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Sevicio:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label LblCodigo 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Producto:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1830
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Vencimiento"
            Height          =   195
            Left            =   2040
            TabIndex        =   20
            Top             =   2280
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "FrmProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD90 As New ADODB.Recordset
Dim BD91 As New ADODB.Recordset
Dim bd92 As New ADODB.Recordset
Dim BDT As New ADODB.Recordset
Dim Cambio
Dim Nuevo
Dim Impuesto1 As Double

Sub Refrescar()

CSql = "SELECT * FROM productos"
If BD90.State = adStateOpen Then BD90.Close

Set BD90 = CrearRS(CSql)
   
End Sub
Sub CargaProveedor()

CSql = "Select IdProveedor, Nombre From Proveedores"
Set bd92 = CrearRS(CSql)
If Not (bd92.EOF) Then
bd92.MoveFirst
Do While Not bd92.EOF
    CboProveedor.AddItem bd92.Fields("Nombre").Value
    CboProveedor.ItemData(CboProveedor.NewIndex) = bd92.Fields("IdProveedor").Value
    bd92.MoveNext
Loop
    bd92.Close
Else
    bd92.Close

End If

End Sub

Sub Blanqueo()
TxtDescripcion.Text = ""
TxtCostoAnterior.Caption = ""
TxtCostoActual.Text = ""
TxtPrecioSinImpuesto1.Text = ""
TxtImpuesto1.Text = ""
TxtPrecioUnitario1.Text = ""
TxtUtilidad1.Text = ""

TxtPrecioSinImpuesto2.Text = ""
TxtImpuesto2.Text = ""
TxtPrecioUnitario2.Text = ""
TxtUtilidad2.Text = ""

TxtPrecioSinImpuesto3.Text = ""
TxtImpuesto3.Text = ""
TxtPrecioUnitario3.Text = ""
TxtUtilidad3.Text = ""
CboUbicacion.Text = ""
CboTipo.Text = ""
'Label3.Caption = ""
CboProveedor.ListIndex = -1
DTPicker1.Value = Now
Nuevo = 1
LblCodigo.Caption = ""
End Sub

Private Sub BtnAgregar_Click()
'command2
Blanqueo

CSql = "Select MAX(IdProducto)+1 as NuevoId From Productos"
Set BDT = CrearRS(CSql)

If BDT.RecordCount <> 0 Then
    If Not IsNull(BDT.Fields("NuevoId").Value) Then
        LblCodigo.Caption = BDT.Fields("NuevoId").Value
    Else
        LblCodigo.Caption = "1"
    End If
Else
    LblCodigo.Caption = "1"
End If
                
TxtCostoAnterior.Caption = "0,00"
TxtCostoActual.Text = "0,00"

TxtUtilidad1.Text = "0,00"
TxtPrecioSinImpuesto1.Text = "0,00"
TxtImpuesto1.Text = "0,00"
TxtPrecioUnitario1.Text = "0,00"
TxtUtilidad2.Text = "0,00"
TxtPrecioSinImpuesto2.Text = "0,00"
TxtImpuesto2.Text = "0,00"
TxtPrecioUnitario2.Text = "0,00"
TxtUtilidad3.Text = "0,00"
TxtPrecioSinImpuesto3.Text = "0,00"
TxtImpuesto3.Text = "0,00"
TxtPrecioUnitario3.Text = "0,00"

CboProveedor.ListIndex = -1

'Check2.Enabled = True
Nuevo = 1

NoReg.Caption = "Nuevo Registro"
Frame1.BackColor = &HE0E0E0
Frame3.BackColor = &HE0E0E0
Frame4.BackColor = &HE0E0E0
Frame5.BackColor = &HE0E0E0
Frame6.BackColor = &HE0E0E0
Frame7.BackColor = &HE0E0E0
Frame8.BackColor = &HE0E0E0
Check1.BackColor = &HE0E0E0

BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
BtnBorrar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False

BtnListadoProductos.Enabled = False

FrameBusqueda.Visible = False
End Sub

Private Sub BtnAnterior_Click()
'command6
BD90.MovePrevious
Call carga
End Sub

Private Sub BtnBorrar_Click()
Dim RsBorrar As New ADODB.Recordset
CSql = "Select * From Productos Where IdProducto='" & LblCodigo.Caption & "'"
Set RsBorrar = CrearRS(CSql)

If RsBorrar.RecordCount > 0 Then
    RsBorrar.Delete

Else
    Msg = "Debe de Seleccionar un producto para poder borrarlo!!!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If
BtnDesHacer_Click


End Sub

Public Sub BtnBuscar_Click()
If Trim(TxtBuscar.Text) = "" Then Exit Sub

If Trim(TxtBuscar.Text) = "" Or UCase(TxtBuscar.Text) = UCase("Busqueda") Then
    CSql = "Select * From Productos"
Else
    CSql = "Select * From Productos Where IdProducto='" & Trim(TxtBuscar.Text) & "' or Descripcion='" & Trim(TxtBuscar.Text) & "'"
End If

Set RsBuscarProducto = CrearRS(CSql)

If Not (RsBuscarProducto.EOF) Then
    If Trim(RsBuscarProducto.Fields("Idproducto").Value) <> "" Then LblCodigo.Caption = RsBuscarProducto.Fields("Idproducto").Value
            If Trim(RsBuscarProducto.Fields("Descripcion").Value) <> "" Then TxtDescripcion.Text = RsBuscarProducto.Fields("Descripcion").Value
            If Trim(RsBuscarProducto.Fields("CostoAnterior").Value) <> "" Then TxtCostoAnterior.Caption = Format(RsBuscarProducto.Fields("CostoAnterior").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("CostoActual").Value) <> "" Then TxtCostoActual.Text = Format(RsBuscarProducto.Fields("CostoActual").Value, "#,##0.00")
            
            If Trim(RsBuscarProducto.Fields("PrecioUnitario1").Value) <> "" Then TxtPrecioUnitario1.Text = Format(RsBuscarProducto.Fields("PrecioUnitario1").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("Utilidad1").Value) <> "" Then TxtUtilidad1.Text = Format(RsBuscarProducto.Fields("Utilidad1").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("PrecioUnitario1SI").Value) <> "" Then TxtPrecioSinImpuesto1.Text = Format(RsBuscarProducto.Fields("PrecioUnitario1SI").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("PrecioUnitarioImpuesto1").Value) <> "" Then TxtImpuesto1.Text = Format(RsBuscarProducto.Fields("PrecioUnitarioImpuesto1").Value, "#,##0.00")
            
            If Trim(RsBuscarProducto.Fields("PrecioUnitario2").Value) <> "" Then TxtPrecioUnitario2.Text = Format(RsBuscarProducto.Fields("PrecioUnitario2").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("Utilidad2").Value) <> "" Then TxtUtilidad2.Text = Format(RsBuscarProducto.Fields("Utilidad2").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("PrecioUnitario2SI").Value) <> "" Then TxtPrecioSinImpuesto2.Text = Format(RsBuscarProducto.Fields("PrecioUnitario2SI").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("PrecioUnitarioImpuesto2").Value) <> "" Then TxtImpuesto2.Text = Format(RsBuscarProducto.Fields("PrecioUnitarioImpuesto2").Value, "#,##0.00")

            If Trim(RsBuscarProducto.Fields("PrecioUnitario3").Value) <> "" Then TxtPrecioUnitario3.Text = Format(RsBuscarProducto.Fields("PrecioUnitario3").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("Utilidad3").Value) <> "" Then TxtUtilidad3.Text = Format(RsBuscarProducto.Fields("Utilidad3").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("PrecioUnitario3SI").Value) <> "" Then TxtPrecioSinImpuesto3.Text = Format(RsBuscarProducto.Fields("PrecioUnitario3SI").Value, "#,##0.00")
            If Trim(RsBuscarProducto.Fields("PrecioUnitarioImpuesto3").Value) <> "" Then TxtImpuesto3.Text = Format(RsBuscarProducto.Fields("PrecioUnitarioImpuesto3").Value, "#,##0.00")

            If Trim(RsBuscarProducto.Fields("TipoServicio").Value) <> "" Then CboTipo.Text = RsBuscarProducto.Fields("TipoServicio").Value
            If Trim(RsBuscarProducto.Fields("Ubicacion").Value) <> "" Then CboUbicacion.Text = RsBuscarProducto.Fields("Ubicacion").Value
            If IsNull(RsBuscarProducto.Fields("idproveedor").Value) Then
                Combo1.ListIndex = -1
            Else
                For a = 0 To CboProveedor.ListCount - 1
                    If CboProveedor.ItemData(a) = RsBuscarProducto.Fields("idproveedor").Value Then
                        d = a
                Exit For
                Else: d = -1
                End If
            Next a
            
            If d = "" Then
                CboProveedor.ListIndex = -1
            Else
                CboProveedor.ListIndex = d
            End If
            
            If RsBuscarProducto.Fields("impuesto").Value Then
                Check1.Value = 1
                LblExento.Visible = False
            Else
                Check1.Value = 0
                LblExento.Visible = True
            End If
            End If
Else
        Msg = "No Existe el producto solicitado"
        MsgBox Msg, vbOKOnly, "No Existe registro alguno"

End If


RsBuscarProducto.Close


BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = True
BtnBorrar.Enabled = True
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True

BtnListadoProductos.Enabled = True

FrameBusqueda.Visible = True


End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
carga

Frame1.BackColor = &HEAEFEF
Frame3.BackColor = &HEAEFEF
Frame4.BackColor = &HEAEFEF
Frame5.BackColor = &HEAEFEF
Frame6.BackColor = &HEAEFEF
Frame7.BackColor = &HEAEFEF
Frame8.BackColor = &HEAEFEF
Check1.BackColor = &HEAEFEF

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnBorrar.Enabled = True
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True

BtnListadoProductos.Enabled = True

FrameBusqueda.Visible = True

End Sub

Private Sub BtnGuardarActualizar_Click()
'validar campos

If CboProveedor.ListIndex = -1 Then
    Msg = "Tiene que seleccionar un Proveedor"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    CboProveedor.SetFocus
    Exit Sub
End If


If CboTipo.ListIndex = -1 Then
    Msg = "Tiene que seleccionar el tipo de Producto"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    CboTipo.SetFocus
    Exit Sub
End If






Select Case Nuevo

    Case Is = 1
        CSql = "Select MAX(IdProducto)+1 as NuevoId From Productos"
        Set BDT = CrearRS(CSql)
        
        If BDT.RecordCount <> 0 Then
            If Not IsNull(BDT.Fields("NuevoId").Value) Then
                C = BDT.Fields("NuevoId").Value
            Else
                C = "1"
            End If
        Else
            C = "1"
        End If
                
        CSql = "Select * From Productos"
        Set BD91 = CrearRS(CSql)
        
        BD91.AddNew
        
        BD91.Fields("IdProducto").Value = C
        BD91.Fields("IdUsuario").Value = IdUser
        BD91.Fields("Descripcion").Value = TxtDescripcion.Text
        BD91.Fields("CostoAnterior").Value = CDbl(TxtCostoAnterior.Caption)
        BD91.Fields("CostoActual").Value = CDbl(TxtCostoActual.Text)
        BD91.Fields("PrecioUnitario1SI").Value = CDbl(TxtPrecioSinImpuesto1.Text)
        BD91.Fields("PrecioUnitarioImpuesto1").Value = CDbl(TxtImpuesto1.Text)
        BD91.Fields("PrecioUnitario1").Value = CDbl(TxtPrecioUnitario1.Text)
        BD91.Fields("Utilidad1").Value = CDbl(TxtUtilidad1.Text)
        
        BD91.Fields("PrecioUnitario2SI").Value = CDbl(TxtPrecioSinImpuesto2.Text)
        BD91.Fields("PrecioUnitarioImpuesto2").Value = CDbl(TxtImpuesto2.Text)
        BD91.Fields("PrecioUnitario2").Value = CDbl(TxtPrecioUnitario2.Text)
        BD91.Fields("Utilidad2").Value = CDbl(TxtUtilidad2.Text)
        
        BD91.Fields("PrecioUnitario3SI").Value = CDbl(TxtPrecioSinImpuesto3.Text)
        BD91.Fields("PrecioUnitarioImpuesto3").Value = CDbl(TxtImpuesto3.Text)
        BD91.Fields("PrecioUnitario3").Value = CDbl(TxtPrecioUnitario3.Text)
        BD91.Fields("Utilidad3").Value = CDbl(TxtUtilidad3.Text)
        
        BD91.Fields("Ubicacion").Value = CboUbicacion.Text
        BD91.Fields("TipoServicio").Value = CboTipo.Text
        BD91.Fields("Impuesto").Value = Check1.Value
        BD91.Fields("IdProveedor").Value = CboProveedor.ItemData(CboProveedor.ListIndex)
        BD91.Fields("FechaProducto").Value = DTPicker1.Value
        BD91.Update
        
        Msg = "Registro Agregado satisfactoriamente"
        MsgBox Msg, vbOKOnly + vbInformation, "Operación Satisfactoria"
        Call Blanqueo
        'llenarProductos
        Nuevo = 0
        
Case Is = 0
If Cambio = 0 Then
        If CboProveedor.ListIndex = -1 Then
            IdProv = 0
        Else
            IdProv = CboProveedor.ItemData(CboProveedor.ListIndex)
        End If
        
        CSql = "Select * From Productos where IdProducto = '" & LblCodigo.Caption & "'"
        Set BD91 = CrearRS(CSql)
        If BD91.RecordCount > 0 Then
            BD91.Fields("IdUsuario").Value = IdUser
            BD91.Fields("Descripcion").Value = TxtDescripcion.Text
            BD91.Fields("CostoAnterior").Value = CDbl(TxtCostoAnterior.Caption)
            BD91.Fields("CostoActual").Value = CDbl(TxtCostoActual.Text)
            BD91.Fields("PrecioUnitario1SI").Value = CDbl(TxtPrecioSinImpuesto1.Text)
            BD91.Fields("PrecioUnitarioImpuesto1").Value = CDbl(TxtImpuesto1.Text)
            BD91.Fields("PrecioUnitario1").Value = CDbl(TxtPrecioUnitario1.Text)
            BD91.Fields("Utilidad1").Value = CDbl(TxtUtilidad1.Text)
            
            BD91.Fields("PrecioUnitario2SI").Value = CDbl(TxtPrecioSinImpuesto2.Text)
            BD91.Fields("PrecioUnitarioImpuesto2").Value = CDbl(TxtImpuesto2.Text)
            BD91.Fields("PrecioUnitario2").Value = CDbl(TxtPrecioUnitario2.Text)
            BD91.Fields("Utilidad2").Value = CDbl(TxtUtilidad2.Text)
            
            BD91.Fields("PrecioUnitario3SI").Value = CDbl(TxtPrecioSinImpuesto3.Text)
            BD91.Fields("PrecioUnitarioImpuesto3").Value = CDbl(TxtImpuesto3.Text)
            BD91.Fields("PrecioUnitario3").Value = CDbl(TxtPrecioUnitario3.Text)
            BD91.Fields("Utilidad3").Value = CDbl(TxtUtilidad3.Text)
            
            BD91.Fields("Ubicacion").Value = CboUbicacion.Text
            BD91.Fields("TipoServicio").Value = CboTipo.Text
            BD91.Fields("Impuesto").Value = Check1.Value
            BD91.Fields("IdProveedor").Value = CboProveedor.ItemData(CboProveedor.ListIndex)
            BD91.Fields("FechaProducto").Value = DTPicker1.Value
            BD91.Update
            
            Msg = "Registro Actualizado Satisfactoriamente"
            MsgBox Msg, vbOKOnly + vbInformation, "Operación Satisfactoria"
            Nuevo = 0
        Else
            Msg = "Registro No se Actualizo Satisfactoriamente"
            MsgBox Msg, vbOKOnly + vbcritica, "Operación Fallida"
        End If
End If

End Select

Call Refrescar
Call carga
BtnDesHacer_Click
End Sub

Private Sub BtnListadoProductos_Click()
Tipo = "Productos"
FrmListadoProductosServicios.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnSiguiente_Click()
'command7
BD90.MoveNext
Call carga
End Sub

Private Sub CboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub CboUbicacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check1.SetFocus
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    LblExento.Visible = True
Else
    LblExento.Visible = False
End If
Cambio = 1
End Sub

Private Sub Combo1_Click()
Cambio = 1
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCostoActual.SetFocus
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    TxtCodigo.Visible = True
    TxtCodigo.SetFocus
    LblCodigo.Visible = False
Else
    TxtCodigo.Visible = False
    LblCodigo.Visible = True
    TxtDescripcion.SetFocus
End If
End Sub

Private Sub DTPicker1_Change()
Cambio = 1
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboUbicacion.SetFocus
End If
End Sub

Private Sub Form_Load()
Centrar Me
Nuevo = 0
CargaConfiguracion
Call CargaProveedor
Call Refrescar
Call carga


BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnBorrar.Enabled = True
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False

LblImpuesto.Caption = Impuesto1

End Sub
Sub CargaConfiguracion()

Dim RsConfiguracion As New ADODB.Recordset
CSql = "Select * From Dat_admin"
Set RsConfiguracion = CrearRS(CSql)

Impuesto1 = RsConfiguracion.Fields("IVA1").Value

End Sub

Sub carga()
    If BD90.EOF Then
        Msg = "Llego al Final del Registro desea Volver al Principio?"
        MsgBox Msg
        BD90.MoveFirst
    End If

If BD90.BOF Then
    Msg = "Llego al principio del registro"
    MsgBox Msg
    BD90.MoveLast
End If
                        
                        
            NoReg.Caption = "Total de Productos: " & BD90.AbsolutePosition & " / " & BD90.RecordCount
                        
            If Trim(BD90.Fields("Idproducto").Value) <> "" Then LblCodigo.Caption = BD90.Fields("Idproducto").Value
            If Trim(BD90.Fields("Descripcion").Value) <> "" Then TxtDescripcion.Text = BD90.Fields("Descripcion").Value
            If Trim(BD90.Fields("CostoAnterior").Value) <> "" Then TxtCostoAnterior.Caption = Format(BD90.Fields("CostoAnterior").Value, "#,##0.00")
            If Trim(BD90.Fields("CostoActual").Value) <> "" Then TxtCostoActual.Text = Format(BD90.Fields("CostoActual").Value, "#,##0.00")
            
            If Trim(BD90.Fields("PrecioUnitario1").Value) <> "" Then TxtPrecioUnitario1.Text = Format(BD90.Fields("PrecioUnitario1").Value, "#,##0.00")
            If Trim(BD90.Fields("Utilidad1").Value) <> "" Then TxtUtilidad1.Text = Format(BD90.Fields("Utilidad1").Value, "#,##0.00")
            If Trim(BD90.Fields("PrecioUnitario1SI").Value) <> "" Then TxtPrecioSinImpuesto1.Text = Format(BD90.Fields("PrecioUnitario1SI").Value, "#,##0.00")
            If Trim(BD90.Fields("PrecioUnitarioImpuesto1").Value) <> "" Then TxtImpuesto1.Text = Format(BD90.Fields("PrecioUnitarioImpuesto1").Value, "#,##0.00")
            
            If Trim(BD90.Fields("PrecioUnitario2").Value) <> "" Then TxtPrecioUnitario2.Text = Format(BD90.Fields("PrecioUnitario2").Value, "#,##0.00")
            If Trim(BD90.Fields("Utilidad2").Value) <> "" Then TxtUtilidad2.Text = Format(BD90.Fields("Utilidad2").Value, "#,##0.00")
            If Trim(BD90.Fields("PrecioUnitario2SI").Value) <> "" Then TxtPrecioSinImpuesto2.Text = Format(BD90.Fields("PrecioUnitario2SI").Value, "#,##0.00")
            If Trim(BD90.Fields("PrecioUnitarioImpuesto2").Value) <> "" Then TxtImpuesto2.Text = Format(BD90.Fields("PrecioUnitarioImpuesto2").Value, "#,##0.00")

            If Trim(BD90.Fields("PrecioUnitario3").Value) <> "" Then TxtPrecioUnitario3.Text = Format(BD90.Fields("PrecioUnitario3").Value, "#,##0.00")
            If Trim(BD90.Fields("Utilidad3").Value) <> "" Then TxtUtilidad3.Text = Format(BD90.Fields("Utilidad3").Value, "#,##0.00")
            If Trim(BD90.Fields("PrecioUnitario3SI").Value) <> "" Then TxtPrecioSinImpuesto3.Text = Format(BD90.Fields("PrecioUnitario3SI").Value, "#,##0.00")
            If Trim(BD90.Fields("PrecioUnitarioImpuesto3").Value) <> "" Then TxtImpuesto3.Text = Format(BD90.Fields("PrecioUnitarioImpuesto3").Value, "#,##0.00")

            If Trim(BD90.Fields("TipoServicio").Value) <> "" Then CboTipo.Text = BD90.Fields("TipoServicio").Value
            If Trim(BD90.Fields("Ubicacion").Value) <> "" Then CboUbicacion.Text = BD90.Fields("Ubicacion").Value
            If IsNull(BD90.Fields("idproveedor").Value) Then
            Combo1.ListIndex = -1
            Else
            For a = 0 To CboProveedor.ListCount - 1
            If CboProveedor.ItemData(a) = BD90.Fields("idproveedor").Value Then
            d = a
            Exit For
            Else: d = -1
            End If
            Next a
            
            If d = "" Then
                CboProveedor.ListIndex = -1
            Else
                CboProveedor.ListIndex = d
            End If
            
            If BD90.Fields("impuesto").Value Then
                Check1.Value = 1
                LblExento.Visible = False
            Else
                Check1.Value = 0
                LblExento.Visible = True
            End If
      Cambio = 0
    End If
End Sub
                
Private Sub Form_Unload(Cancel As Integer)
If BD90.State = adStateOpen Then BD90.Close

End Sub

Private Sub Text1_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text1.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text1.Text)
    pru = LCase(Mid(Text1.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text1.Text = StrText
Text1.SelStart = Len(Text1.Text)

End Sub

Private Sub Text2_Change()
Cambio = 1

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
ElseIf KeyAscii <> 8 Then
    If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text3_Change()
Cambio = 1
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
ElseIf KeyAscii <> 8 Then
    If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text4_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text4.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text4.Text)
    pru = LCase(Mid(Text4.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text4.Text = StrText
Text4.SelStart = Len(Text4.Text)

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
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
LblCodigo.Caption = Trim(TxtCodigo.Text)
    TxtDescripcion.SetFocus
End If
End Sub

Private Sub TxtCostoActual_Change()
If TxtCostoActual.Text = "" Then
    TxtCostoActual.Text = "0,00"
    TxtCostoPromedio.Caption = Format((CDbl(TxtCostoActual.Text) + CDbl(TxtCostoAnterior.Caption)) / 2, "#,##0.00")
Else
    TxtCostoPromedio.Caption = Format((CDbl(TxtCostoActual.Text) + CDbl(TxtCostoAnterior.Caption)) / 2, "#,##0.00")
End If
Cambio = 0
End Sub

Private Sub TxtCostoActual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtUtilidad1.SetFocus
End If
End Sub

Private Sub TxtCostoAnterior_Change()
If TxtCostoAnterior.Caption = "" Then
    TxtCostoAnterior.Caption = Format(TxtCostoActual.Text, "#,##0.00")
Else
    TxtCostoPromedio.Caption = Format((CDbl(TxtCostoActual.Text) + CDbl(TxtCostoAnterior.Caption)) / 2, "#,##0.00")
End If
End Sub

Private Sub TxtDescripcion_Change()
Cambio = 0
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboTipo.SetFocus
End If
End Sub

Private Sub TxtImpuesto1_Change()
Cambio = 0
End Sub

Private Sub TxtImpuesto2_Change()
Cambio = 0
End Sub

Private Sub TxtImpuesto3_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioSinImpuesto1_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioSinImpuesto1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtPrecioSinImpuesto.Text = Format((CDbl(TxtUtilidad1.Text) / 100) * CDbl(TxtCostoActual.Text) + CDbl(TxtCostoActual.Text), "#,##0.00")
    If Check1.Value = 1 Then
        TxtImpuesto1.Text = Format(CDbl(TxtPrecioSinImpuesto.Text) * (Impuesto1 / 100), "#,##0.00")
        TxtPrecioUnitario1.Text = Format(CDbl(TxtImpuesto1.Text) + CDbl(TxtPrecioSinImpuesto.Text), "#,##0.00")
    Else
        TxtImpuesto1.Text = Format(0, "#,##0.00")
        TxtPrecioUnitario1.Text = Format(CDbl(TxtPrecioSinImpuesto.Text), "#,##0.00")

    End If
    TxtPrecioUnitario1.SetFocus
End If
End Sub



Private Sub TxtPrecioSinImpuesto2_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioSinImpuesto3_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioUnitario1_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioUnitario1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboProveedor.SetFocus
End If
End Sub

Private Sub TxtPrecioUnitario2_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioUnitario2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboProveedor.SetFocus
End If
End Sub

Private Sub TxtPrecioUnitario3_Change()
Cambio = 0
End Sub

Private Sub TxtPrecioUnitario3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboProveedor.SetFocus
End If
End Sub

Private Sub TxtUtilidad1_Change()
Cambio = 0
End Sub

Private Sub TxtUtilidad1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtPrecioSinImpuesto1.Text = Format((CDbl(TxtUtilidad1.Text) / 100) * CDbl(TxtCostoActual.Text) + CDbl(TxtCostoActual.Text), "#,##0.00")
    If Check1.Value = 1 Then
        TxtImpuesto1.Text = Format(CDbl(TxtPrecioSinImpuesto.Text) * (Impuesto1 / 100), "#,##0.00")
        TxtPrecioUnitario1.Text = Format(CDbl(TxtImpuesto1.Text) + CDbl(TxtPrecioSinImpuesto.Text), "#,##0.00")
    Else
        TxtImpuesto1.Text = Format(0, "#,##0.00")
        TxtPrecioUnitario1.Text = Format(CDbl(TxtPrecioSinImpuesto.Text), "#,##0.00")

    End If
    TxtPrecioUnitario1.SetFocus
End If
End Sub

Private Sub TxtUtilidad2_Change()
Cambio = 0
End Sub

Private Sub TxtUtilidad2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtPrecioSinImpuesto2.Text = Format((CDbl(TxtUtilidad2.Text) / 100) * CDbl(TxtCostoActual.Text) + CDbl(TxtCostoActual.Text), "#,##0.00")
    If Check1.Value = 1 Then
        TxtImpuesto2.Text = Format(CDbl(TxtPrecioSinImpuesto2.Text) * (Impuesto1 / 100), "#,##0.00")
        TxtPrecioUnitario2.Text = Format(CDbl(TxtImpuesto2.Text) + CDbl(TxtPrecioSinImpuesto2.Text), "#,##0.00")
    Else
        TxtImpuesto2.Text = Format(0, "#,##0.00")
        TxtPrecioUnitario2.Text = Format(CDbl(TxtPrecioSinImpuesto2.Text), "#,##0.00")

    End If
    TxtPrecioUnitario2.SetFocus
End If
End Sub

Private Sub TxtUtilidad3_Change()
Cambio = 0
End Sub

Private Sub TxtUtilidad3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtPrecioSinImpuesto2.Text = Format((CDbl(TxtUtilidad2.Text) / 100) * CDbl(TxtCostoActual.Text) + CDbl(TxtCostoActual.Text), "#,##0.00")
    If Check1.Value = 1 Then
        TxtImpuesto3.Text = Format(CDbl(TxtPrecioSinImpuesto3.Text) * (Impuesto1 / 100), "#,##0.00")
        TxtPrecioUnitario3.Text = Format(CDbl(TxtImpuesto3.Text) + CDbl(TxtPrecioSinImpuesto3.Text), "#,##0.00")
    Else
        TxtImpuesto3.Text = Format(0, "#,##0.00")
        TxtPrecioUnitario3.Text = Format(CDbl(TxtPrecioSinImpuesto3.Text), "#,##0.00")

    End If
    TxtPrecioUnitario3.SetFocus

End Sub

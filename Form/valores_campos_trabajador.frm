VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmValoresCampoTrabajador 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores de campo por trabajador"
   ClientHeight    =   8730
   ClientLeft      =   5955
   ClientTop       =   795
   ClientWidth     =   9570
   Icon            =   "valores_campos_trabajador.frx":0000
   LinkTopic       =   "Form51"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9570
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del trabajador"
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   9135
         Begin ChamaleonButton.ChameleonBtn BrnListaEmpleados 
            Height          =   375
            Left            =   6720
            TabIndex        =   43
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Listado Empleados"
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
            MICON           =   "valores_campos_trabajador.frx":1002
            PICN            =   "valores_campos_trabajador.frx":101E
            PICH            =   "valores_campos_trabajador.frx":12A7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtFIngreso 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox TxtDpto 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox TxtCargo 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox TxtGrupo 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   270
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   720
            Width           =   3135
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   8640
            Top             =   840
         End
         Begin MSScriptControlCtl.ScriptControl ScriptControl1 
            Left            =   8520
            Top             =   240
            _ExtentX        =   1005
            _ExtentY        =   1005
         End
         Begin VB.Label NoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   4560
            TabIndex        =   42
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Ingreso:"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   1770
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dpto. :"
            Height          =   195
            Left            =   4560
            TabIndex        =   31
            Top             =   1290
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo:"
            Height          =   195
            Left            =   4560
            TabIndex        =   29
            Top             =   810
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo:"
            Height          =   195
            Left            =   4560
            TabIndex        =   27
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   1290
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   810
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   7800
         Width           =   3615
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   390
            Left            =   240
            TabIndex        =   17
            Text            =   "Busqueda"
            ToolTipText     =   "Busca un Trabjador por Nombre, Apellido o Número de Cédula"
            Top             =   262
            Width           =   1815
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2160
            TabIndex        =   18
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "valores_campos_trabajador.frx":16C2
            PICN            =   "valores_campos_trabajador.frx":16DE
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
      Begin VB.Frame Frame8 
         BackColor       =   &H00EAEFEF&
         Height          =   2655
         Left            =   7920
         TabIndex        =   13
         Top             =   5160
         Width           =   1335
         Begin ChamaleonButton.ChameleonBtn BtnEliminarConp 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Borra un concepto del trabajador"
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "valores_campos_trabajador.frx":1943
            PICN            =   "valores_campos_trabajador.frx":195F
            PICH            =   "valores_campos_trabajador.frx":1B03
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarConp 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Agregar un concepto al trabajador"
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
            MICON           =   "valores_campos_trabajador.frx":1CA2
            PICN            =   "valores_campos_trabajador.frx":1CBE
            PICH            =   "valores_campos_trabajador.frx":1E4B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnRecargar 
            Height          =   375
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Verifica todos los conceptos del trabajador"
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Verificar"
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
            MICON           =   "valores_campos_trabajador.frx":2080
            PICN            =   "valores_campos_trabajador.frx":209C
            PICH            =   "valores_campos_trabajador.frx":237E
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
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3840
         TabIndex        =   8
         Top             =   7800
         Width           =   5415
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   4320
            TabIndex        =   9
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
            MICON           =   "valores_campos_trabajador.frx":25CF
            PICN            =   "valores_campos_trabajador.frx":25EB
            PICH            =   "valores_campos_trabajador.frx":27B4
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
            Left            =   2880
            TabIndex        =   10
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "valores_campos_trabajador.frx":29E9
            PICN            =   "valores_campos_trabajador.frx":2A05
            PICH            =   "valores_campos_trabajador.frx":2CE7
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
            Left            =   1200
            TabIndex        =   11
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
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
            MICON           =   "valores_campos_trabajador.frx":2F38
            PICN            =   "valores_campos_trabajador.frx":2F54
            PICH            =   "valores_campos_trabajador.frx":31EA
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
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
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
            MICON           =   "valores_campos_trabajador.frx":3449
            PICN            =   "valores_campos_trabajador.frx":3465
            PICH            =   "valores_campos_trabajador.frx":36FA
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
         Caption         =   "Conceptos del Trabajador"
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   7695
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   2280
            Width           =   1695
         End
         Begin SystemOncoAmerica.DMGrid DMGrid2 
            Height          =   1695
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Las Cantidades de las Asignaciones y Deducciones, son referenciales"
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   2990
            Object.Width           =   7425
            Object.Height          =   1665
            ScrollBar       =   1
            Editable        =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Neto:"
            Height          =   195
            Left            =   4200
            TabIndex        =   39
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Deducciones:"
            Height          =   195
            Left            =   2040
            TabIndex        =   37
            Top             =   2040
            Width           =   990
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Asignaciones:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   2040
            Width           =   990
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   2655
         Left            =   7920
         TabIndex        =   2
         Top             =   2520
         Width           =   1335
         Begin ChamaleonButton.ChameleonBtn BtnEliminar 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Borra un campo del trabajador"
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "valores_campos_trabajador.frx":3956
            PICN            =   "valores_campos_trabajador.frx":3972
            PICH            =   "valores_campos_trabajador.frx":3B16
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
            TabIndex        =   5
            ToolTipText     =   "Agregar un campo al trabajador"
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
            MICON           =   "valores_campos_trabajador.frx":3CB5
            PICN            =   "valores_campos_trabajador.frx":3CD1
            PICH            =   "valores_campos_trabajador.frx":3E5E
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
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Guardar / Actualizar el registro del trabajador"
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "valores_campos_trabajador.frx":4093
            PICN            =   "valores_campos_trabajador.frx":40AF
            PICH            =   "valores_campos_trabajador.frx":433E
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
         Caption         =   "Campos del Trabajador"
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   7695
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   2295
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
            Object.Width           =   7425
            Object.Height          =   2265
            ScrollBar       =   1
            Editable        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "FrmValoresCampoTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TmpIdGrupo As Integer
Dim BdCampos As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim RsTemp2 As New ADODB.Recordset
Dim BD64 As New ADODB.Recordset
Dim BD65 As New ADODB.Recordset
Dim Cambio
Dim bdfg As New ADODB.Recordset
Dim bdfg1 As New ADODB.Recordset
Dim RsIdMax As New ADODB.Recordset
Dim IdEmpla As Integer
Dim CodDMGrid1 As Integer
Dim NomDMGrid1 As String
Dim FechaPeriodo As String
Dim FechaFinPeriodo As String
Dim Periodo As String

Public Sub CargarEmpleadoDeLista(IdEEmpl As Integer)
BdCampos.MoveFirst
BdCampos.Find "IdEmpleado=" & IdEEmpl

If BdCampos.EOF Or BdCampos.BOF Then BdCampos.MoveFirst
Call CargaEmpleados
Call CargaRenglon
End Sub

Sub CargaEmpleados()
If Not IsNull(BdCampos.Fields("Nombre")) Then Text1.Text = BdCampos.Fields("Nombre") Else Text1.Text = ""
If Not IsNull(BdCampos.Fields("Apellido")) Then Text2.Text = BdCampos.Fields("Apellido") Else Text2.Text = ""
If Not IsNull(BdCampos.Fields("Cedula")) Then Text3.Text = BdCampos.Fields("Cedula") Else Text3.Text = ""
IdEmpla = BdCampos("idempleado").Value
TmpIdGrupo = Val(BdCampos("Id_Grupo").Value)

NoReg.Caption = "Registro " & BdCampos.AbsolutePosition & " / " & BdCampos.RecordCount


    CSql = "SELECT Empleados.IdEmpleado, Empleados.Nombre, Empleados.Apellido, Empleados.Departamentos, " & _
        "Departamentos.Descripcion, Cargos.Cargo, Grupo.Descripcion AS Expr1, Empleados.Fecha_Ing FROM Empleados INNER JOIN " & _
        " Departamentos ON Empleados.Departamentos = Departamentos.IdDepartamento INNER JOIN Cargos ON " & _
        " Empleados.Cargo = Cargos.IdCargos INNER JOIN Grupo ON Empleados.Id_Grupo = Grupo.Id_Grupo " & _
        " WHERE Empleados.IdEmpleado=" & IdEmpla & " AND Grupo.Id_Grupo=" & TmpIdGrupo & " ORDER BY Empleados.Fecha_Ing"
    Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

If Not IsNull(RsTemp.Fields("Descripcion").Value) Then
    TxtDpto.Text = RsTemp.Fields("Descripcion").Value
    TxtCargo.Text = RsTemp.Fields("Cargo").Value
    TxtGrupo.Text = RsTemp.Fields("Expr1").Value
    TxtFIngreso.Text = RsTemp.Fields("Fecha_Ing").Value
End If
End Sub

Sub Grid()
DMGrid1.Cols = 5
DMGrid1.Rows = 1
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True

DMGrid1.DColumnas(5).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300
'DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 15 / 100)
'DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid1.DColumnas(3).Visible = False
DMGrid1.DColumnas(4).Visible = False

DMGrid1.DColumnas(1).Caption = "Codigo"
DMGrid1.DColumnas(2).Caption = "Descripcion"
DMGrid1.DColumnas(3).Caption = "Valor Texto"
DMGrid1.DColumnas(4).Caption = "Valor Fecha"
DMGrid1.DColumnas(5).Caption = "Valor Número"

DMGrid2.Cols = 4
DMGrid2.Rows = 1
DMGrid2.DColumnas(1).Alignment = 1
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 1
DMGrid2.DColumnas(4).Alignment = 1

DMGrid2.DColumnas(3).IsNumber = True
DMGrid2.DColumnas(4).IsNumber = True

DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 60 / 100) - 300
DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(4).Width = Val(DMGrid2.Width * 15 / 100)

DMGrid2.DColumnas(1).Caption = "Codigo"
DMGrid2.DColumnas(2).Caption = "Descripción"
DMGrid2.DColumnas(3).Caption = "Asignaciones"
DMGrid2.DColumnas(4).Caption = "Deducciones"
End Sub

Sub CargaRenglon()
try:
'cargar los renglones de los conceptos
Text4.Text = "0"
Text5.Text = "0"
Text6.Text = "0"
DMGrid1.Clear
DMGrid1.Row = 1
CSql = "Select * From ValorCampoNomina Where IdEmpleado = '" & IdEmpla & "' AND Expr1='CA' ORDER BY IdCampoNomina"
If BD64.State = 1 Then BD64.Close
Set BD64 = CrearRS(CSql)
i = 1
If Not (BD64.EOF) Then
    BD64.MoveFirst
    Do While Not BD64.EOF
        DMGrid1.Rows = i
        'CSql = "select descripcion from productos where idproducto = " & bd64.Fields("cod_producto")
        'Set BD70 = CrearRS(CSql)
        DMGrid1.ValorCelda(i, 1) = Format(BD64.Fields("idcamponomina"), "00000")
        DMGrid1.ValorCelda(i, 2) = BD64.Fields("campo")
        
        If Trim(BD64.Fields("valort")) = "" Then
            DMGrid1.ValorCelda(i, 3) = "0"
        Else
            DMGrid1.ValorCelda(i, 3) = BD64.Fields("valort")
        End If
        
        If Trim(BD64.Fields("valorf")) = "" Then
            DMGrid1.ValorCelda(i, 4) = "0"
        Else
            DMGrid1.ValorCelda(i, 4) = Format(BD64.Fields("valorf"), "dd/mm/yyyy")
        End If
        
        DMGrid1.ValorCelda(i, 5) = BD64.Fields("valorn")
        i = i + 1
        
        BD64.MoveNext
    Loop
Else
    DMGrid1.Clear
    DMGrid1.Rows = 1
    
    'msg = "Este Trabajador no tiene campos asignados desea asignarle los campos ahora?"
    'd = MsgBox(msg, vbYesNo, "Sin campos")
    d = vbNo
    If d = vbYes Then
        'Call Generar_Campos
        GoTo try
    End If

End If
Call DMGrid1.PaintMGrid

DMGrid2.Clear
DMGrid2.Row = 1
CSql = " SELECT CamposDelTrabajador.*, Concepto.Descripcion, Concepto.Tipo AS Expr1 FROM CamposDelTrabajador INNER JOIN " & _
        " Concepto ON CamposDelTrabajador.IdCampoNomina = Concepto.IdConcepto WHERE " & _
        " (CamposDelTrabajador.IdEmpleado = " & IdEmpla & ") AND (CamposDelTrabajador.Tipo = 'CO') ORDER BY CamposDelTrabajador.IdCampoNomina"

If BD64.State = 1 Then BD64.Close
Set BD64 = CrearRS(CSql)
i = 1
If Not (BD64.EOF) Then
    BD64.MoveFirst
    Do While Not BD64.EOF
        DMGrid2.Rows = i
        DMGrid2.ValorCelda(i, 1) = Format(BD64.Fields("idcamponomina"), "00000")
        DMGrid2.ValorCelda(i, 2) = BD64.Fields("Descripcion")
        
'        If Val(BD64.Fields("Expr1")) = 0 Then
'            DMGrid2.ValorCelda(i, 3) = BD64.Fields("valorn") ' resultado del concepto en ASIGNACIONES
'        Else
'            DMGrid2.ValorCelda(i, 4) = BD64.Fields("valorn") ' resultado del concepto en DEDUCCIONES
'        End If
        i = i + 1
        
        BD64.MoveNext
    Loop
    'Verificar_Cambios_Conceptos
Else
    DMGrid2.Clear
    DMGrid2.Rows = 1
    
    'msg = "Este Trabajador no tiene campos asignados desea asignarle los campos ahora?"
    'd = MsgBox(msg, vbYesNo, "Sin campos")
    d = vbNo
    If d = vbYes Then
        'Call Generar_Campos
        GoTo try
    End If

End If
DMGrid2.PaintMGrid
Cambio = 0
End Sub

Private Sub BrnListaEmpleados_Click()
Tipo = "VCT"
BtnDesHacer_Click
FrmListadoEmpleados.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAgregar_Click()
FrmListaCampos.Show
Cambio = 1
End Sub

Private Sub BtnAgregarConp_Click()
FrmListaConceptos.Show vbModal, FrmPrincipal
Cambio = 1
End Sub

Private Sub BtnAnterior_Click()

If BdCampos.RecordCount = 0 Then Exit Sub

If BdCampos.State = 0 Then
    TxtBuscar.Text = ""
    Call BtnBuscar_Click
    Exit Sub
End If

If Not BdCampos.BOF Then BdCampos.MovePrevious
If BdCampos.BOF = True Then BdCampos.MoveLast
Call CargaEmpleados
Call CargaRenglon
End Sub

Private Sub BtnBuscar_Click()
IdEmpla = 0
Call Conectar
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
IdEmpla = 0
TxtBuscar.Text = "Busqueda"
Conectar
End Sub

Private Sub BtnEliminar_Click()
Dim resp As Byte

CodDMGrid1 = DMGrid1.ValorCelda(DMGrid1.Row, 1)
NomDMGrid1 = DMGrid1.ValorCelda(DMGrid1.Row, 2)

CSql = "SELECT Restringido FROM CamposDeNomina WHERE IdCampoNomina=" & CodDMGrid1
Set RsTemp = CrearRS(CSql)

If Val(RsTemp.Fields("Restringido").Value) = 1 Then
    MsgBox "El campo " & NomDMGrid1 & " no puede ser eliminado!", vbExclamation + vbOKOnly, "NO SE PUEDE ELIMINAR!"
Else
    resp = MsgBox("Se procederá a eliminar el campo " & NomDMGrid1 & ", Desea continuar?", vbQuestion + vbYesNo, "Confirmar!")
    If resp = vbNo Then Exit Sub
    DMGrid1.RowDelete (DMGrid1.Row)
    DMGrid1.PaintMGrid
    Cambio = 1
End If
End Sub

Private Sub BtnEliminarConp_Click()
DMGrid2.RowDelete (DMGrid2.Row)
DMGrid2.PaintMGrid
Cambio = 1
End Sub

Private Sub BtnGuardar_Click()
If Cambio = 1 Then
    Call Guardar
Else
    MsgBox "No hay cambios que guardar", vbOKOnly, "No hay cambios..."
End If
BtnRecargar_Click
End Sub

Private Sub BtnRecargar_Click()
Verificar_Cambios_Conceptos
Verificar_Prestamos
End Sub

Sub Verificar_Prestamos()
Dim TamDMGrid As Integer
Dim NCuota As Integer
Dim IdPrest As Integer
Dim CantPrestamo As Double
Dim i As Integer

' Consulta para obtener el RENGLON DEL MAYOR PRESTAMO de acuerdo a la fecha de la generación de
' la nomina, cuyo DEUDA DEL PRESTAMO SEA DIFERENTE DE 0... (No valido que el renglon deba ser CERO
' ya que se podria dar el caso de que la cuota se halla pagado a una mitad y aun quede por cobrarla),
' de todas maneras se podria colocar en la consulta la condicion ==> " AND MontoAbono=0" para decir
' que el RENGLON DEL PRESTAMO TIENE QUE SER CERO...
CSql = "SELECT * From RenglonPrestamos " & _
    "   WHERE (IdPrestamo = " & _
    "          (SELECT IdPrestamos From Prestamos " & _
    "           WHERE (Monto_Presta = " & _
    "                 (SELECT MAX(Monto_Presta) AS MP From Prestamos " & _
    "                  WHERE      (IdEmpleado = " & IdEmpla & ") AND (Activo = '1') AND Adeuda <> 0)))) AND (FechaPago = '" & FechaFinPeriodo & "')"
Set RsTemp = CrearRS(CSql)


' Condicional q verifica si la consulta contiene CERO "0" registros entontrados
If RsTemp.RecordCount = 0 Then

    ' Si la consulta anterior no contiene registros entonces, buscar en la BASE DE DATOS
    ' los cobros para la fecha de nomina que se esta procesando...
    CSql = "SELECT RenglonPrestamos.*, Prestamos.Monto_Presta FROM RenglonPrestamos INNER JOIN Prestamos ON RenglonPrestamos.IdPrestamo = Prestamos.IdPrestamos " & _
        " WHERE (RenglonPrestamos.FechaPago = '" & FechaFinPeriodo & "') AND Prestamos.IdEmpleado=" & IdEmpla & " ORDER BY Prestamos.IdPrestamos"
    Set RsTemp = CrearRS(CSql)
    
    ' Si la consulta anterior contiene mas de 1 registro, entonces elige el prestamo mayor...
    If RsTemp.RecordCount > 1 Then
    
        CantPrestamo = 0
        While Not RsTemp.EOF
            If CDbl(RsTemp.Fields("Monto_Presta").Value) > CantPrestamo Then
                CantPrestamo = CDbl(RsTemp.Fields("Monto_Presta").Value)
                IdPrest = Val(RsTemp.Fields("IdPrestamo").Value)
            End If
            RsTemp.MoveNext
        Wend
        
        CSql = "SELECT * FROM RenglonPrestamos WHERE IdPrestamo=" & IdPrest & " AND FechaPago='" & FechaFinPeriodo & "'"
        Set RsTemp = CrearRS(CSql)
    ' Si no, entonces verifica si no encontro algun cobro, de ser asi finaliza la busqueda
    ElseIf RsTemp.RecordCount = 0 Then
        Exit Sub
    End If
End If

IdPrest = RsTemp.Fields("IdPrestamo").Value
CantPrestamo = CDbl(RsTemp.Fields("AbonoMax").Value) - CDbl(RsTemp.Fields("MontoAbono").Value)

' Si la cantidad del prestamo para la fecha de la nomina ya esta cancelada entonces
' no sigue verificando y llama al EXIT SUB
If CantPrestamo = 0 Then Exit Sub
CantPrestamo = RsTemp.Fields("AbonoMax").Value

CSql = "SELECT COUNT(*) as Result FROM RenglonPrestamos WHERE IdPrestamo=" & IdPrest & " AND MontoAbono<>0"
Set RsTemp = CrearRS(CSql)
NCuota = RsTemp.Fields(0).Value + 1

CSql = "SELECT COUNT(*) as Result FROM RenglonPrestamos WHERE IdPrestamo=" & IdPrest
Set RsTemp = CrearRS(CSql)

TamDMGrid = DMGrid2.Rows
For i = 1 To TamDMGrid
    If Trim(DMGrid2.ValorCelda(i, 1)) = "00000" Or Trim(DMGrid2.ValorCelda(i, 1)) = "0000" Then
        DMGrid2.RowDelete i
        DMGrid2.PaintMGrid
    End If
Next i

DMGrid2.Rows = DMGrid2.Rows + 1
DMGrid2.ValorCelda(DMGrid2.Rows, 1) = "00000"
DMGrid2.ValorCelda(DMGrid2.Rows, 2) = "PRESTAMO " & NCuota & " / " & RsTemp.Fields(0).Value
DMGrid2.ValorCelda(DMGrid2.Rows, 4) = CantPrestamo
DMGrid2.PaintMGrid

Text4.Text = Format(CDbl(Text4.Text), "#,##0.00")
Text5.Text = Format(CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 4)) + CDbl(Text5.Text), "#,##0.00")
Text6.Text = Format(CDbl(CDbl(Text4.Text) - CDbl(Text5.Text)), "#,##0.00")
End Sub

Private Sub BtnSiguiente_Click()
If BdCampos.RecordCount = 0 Then Exit Sub
If BdCampos.State = 0 Then
    TxtBuscar.Text = ""
    Call BtnBuscar_Click
    Exit Sub
End If
If Not BdCampos.EOF Then BdCampos.MoveNext
If BdCampos.EOF Then BdCampos.MoveFirst
Call CargaEmpleados
Call CargaRenglon
End Sub

Sub Guardar()
Dim CSql As String
If IdEmpla = 0 Then MsgBox "Debe Seleccionar un empleado!", vbExclamation + vbOKOnly, "Error": Exit Sub

' Metodo que verifica si se eliminaron o agregaron CAMPOS, si es asi, modifica el registro para ese empleado
' cuya ID esta almacenada en la variable IdEmpla
Verificar_Cambios
'mmmmmmmmmmmmmmmmmmmmmmmmmmmm
Eliminar_Conceptos  ' Elimina los conceptos que fueron borrados en el DMGrid2
'mmmmmmmmmmmmmmmmmmmmmmmmmmmm

For i = 1 To DMGrid1.Rows
    b1 = DMGrid1.ValorCelda(i, 1)
    b2 = DMGrid1.ValorCelda(i, 2)
    'b3 = DMGrid1.ValorCelda(i, 3)
    'b4 = DMGrid1.ValorCelda(i, 4)
    b3 = ""
    b4 = ""
    b5 = DMGrid1.ValorCelda(i, 5)
    
   
    If IsNull(b3) Or Trim(b3) = "" Then b3 = "0"
    If IsNull(b4) Or Trim(b4) = "" Then b4 = "0" Else b = "0" 'b4 = " " & Format(b4, "dd/mm/yyyy") & " "
    If IsNull(b5) Or Trim(b5) = "" Then b5 = 0
    
    If Not Trim(b1) = "" Then
        If Trim(b2) = "" Or Trim(b3) = "" Or Trim(b4) = "" Or Trim(b5) = "" Then
            MsgBox "Todos los campos deben estar llenos!", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
    End If
    
    If Trim(b1) <> "" Then
        CSql = "UPDATE CamposDelTrabajador Set valort = '" & b3 & _
        "', valorf='" & b4 & "', valorn = " & Replace(b5, ",", ".") & " WHERE IdCampoNomina = " & DMGrid1.ValorCelda(i, 1) & " AND  idempleado = " & IdEmpla & " AND Tipo='CA'"
        Set BD65 = CrearRS(CSql)
    End If
Next i
Guardar_Conceptos

MsgBox "Se ha guardado satisfactoriamente los cambios hechos", vbOKOnly, "Operación Exitosa"
DMGrid1.PaintMGrid
DMGrid2.PaintMGrid
'Generar_Nomina
Cambio = 0
End Sub

'MMMMMMMMMMMMMMM  Genera los valores de la nomina  MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Sub Generar_Nomina()
    
End Sub
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function Verificar_Cambios_Conceptos()
Dim CodConp As Integer
Dim CadConcepto As String
Dim CadCampo As String
Dim CadSSO As String
Dim CadConstantes As String
Dim Band As Boolean
Dim resultado As Double
'Dim Periodo
Text4.Text = "0"
Text5.Text = "0"
Text6.Text = "0"
For i = 1 To DMGrid2.Rows
    
    CodConp = Val(DMGrid2.ValorCelda(i, 1))
    
    If CodConp <> 0 Then
        CSql = "SELECT * FROM Concepto Where IdConcepto=" & CodConp & " AND IdConcepto<>0"
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
            IdEmpl = IdEmpla
            CadConcepto = Validar_Concepto(RsTemp.Fields("Formula").Value)
            CadCampo = Validar_Campo(CadConcepto)
            CadFunciones = Validar_Funcion(CadCampo, Periodo, FechaPeriodo)
            CadConstantes = Validar_Constante(CadFunciones)
            CadConstantes = Replace(CadConstantes, ",", ".")
            CadSSO = CadConstantes
            CadConstantes = Validar_SSO(CadConstantes)
            
            On Error GoTo errty
            resultado = ScriptControl1.Eval(CadConstantes)
            'MsgBox CadConstantes
            If CadSSO <> CadConstantes Then
                resultado = Calcular_SSO(CadConstantes, resultado, IdEmpla)
            End If
            
            If Val(RsTemp.Fields("Tipo").Value) = 0 Then
                DMGrid2.ValorCelda(i, 3) = CDbl(resultado)
                Text4.Text = Format(CStr(CDbl(Text4.Text) + CDbl(resultado)), "#,##0.00")
            Else
                DMGrid2.ValorCelda(i, 4) = CDbl(resultado)
                Text5.Text = Format(CStr(CDbl(Text5.Text) + CDbl(resultado)), "#,##0.00")
            End If
Seguir:
            RsTemp.MoveNext
        Else
            MsgBox "El Concepto de código " & CodConp & " Genera errores ya que no existe, se Eliminara!", vbCritical + vbOKOnly, "Error"
            CSql = "DELETE FROM CamposDelTrabajador Where IdCampoNomina=" & CodConp & " AND Tipo='CO' AND IdEmpleado=" & IdEmpla
            Set RsTemp = CrearRS(CSql)
        End If
    End If
Next i

Text4.Text = Format(CDbl(Text4.Text), "#,##0.00")
Text5.Text = Format(CDbl(Text5.Text), "#,##0.00")
Text6.Text = Format(CDbl(CDbl(Text4.Text) - CDbl(Text5.Text)), "#,##0.00")

DMGrid2.PaintMGrid
Exit Function

errty:
CSql = "DELETE FROM CamposDelTrabajador Where IdCampoNomina=" & CodConp & " AND Tipo='CO' AND IdEmpleado=" & IdEmpla
Set RsTemp2 = CrearRS(CSql)
GoTo Seguir

End Function

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function Eliminar_Conceptos()
Dim CodConp As Integer
Dim Band As Boolean

' consulta que busca todos los CONCEPTOS de un Empleado Especifico
CSql = "SELECT * FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla & " AND Tipo='CO'"
Set RsTemp = CrearRS(CSql)

While Not RsTemp.EOF

    Band = False
    ' Ciclo para determinar si algun concepto del registro del empleado esta o no en el DMGrid2,
    ' si se encuentra entonces BAND=TRUE
    For i = 1 To DMGrid2.Rows
        CodConp = Val(DMGrid2.ValorCelda(i, 1))
        If Val(RsTemp.Fields("IdCampoNomina").Value) = CodConp Then
            Band = True
            Exit For
        End If
    Next i

    ' si BAND=FALSE, elimina del registro los CONCEPTOS que no esten en el DMGrid2
    If Band = False Then
        CSql = "DELETE FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla & " AND Tipo='CO' AND IdCampoNomina=" & Val(RsTemp.Fields("IdCampoNomina").Value)
        Set RsTemp2 = CrearRS(CSql)
    End If

    RsTemp.MoveNext
Wend
DMGrid2.PaintMGrid
End Function
' guardar los conceptos
Sub Guardar_Conceptos()

Dim CodConp As Integer
Dim CadConcepto As String
Dim CadCampo As String
Dim CadConstantes As String
Dim CadSSO As String
Dim NuevoId As Integer


For i = 1 To DMGrid2.Rows
    
    CodConp = Val(DMGrid2.ValorCelda(i, 1))
    
    If CodConp <> 0 Then
        CSql = "SELECT * FROM Concepto Where IdConcepto=" & CodConp
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
            IdEmpl = IdEmpla
            CadConcepto = Validar_Concepto(RsTemp.Fields("Formula").Value)
            CadCampo = Validar_Campo(CadConcepto)
            CadFunciones = Validar_Funcion(CadCampo, Periodo, FechaPeriodo)
            CadConstantes = Validar_Constante(CadFunciones)
            CadConstantes = Replace(CadConstantes, ",", ".")
            CadSSO = CadConstantes
            CadConstantes = Validar_SSO(CadConstantes)
            
            On Error GoTo errty
            resultado = ScriptControl1.Eval(CadConstantes)

            If CadSSO <> CadConstantes Then
                resultado = Calcular_SSO(CadConstantes, resultado, IdEmpla)
            End If
Seguir:
            If Val(RsTemp.Fields("Tipo").Value) = 0 Then
                DMGrid2.ValorCelda(i, 3) = CDbl(resultado)
            Else
                DMGrid2.ValorCelda(i, 4) = CDbl(resultado)
            End If

            CSql = "SELECT * FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla & " AND Tipo='CO' AND IdCampoNomina=" & CodConp
            Set RsTemp2 = CrearRS(CSql)
            
            If RsTemp2.RecordCount <> 0 Then
                CSql = "UPDATE CamposDelTrabajador Set ValorN=" & Replace(CDbl(resultado), ",", ".") & " " & _
                        " WHERE IdEmpleado=" & IdEmpla & " AND Tipo ='CO' AND IdCampoNomina=" & CodConp
                Set RsTemp2 = CrearRS(CSql)
            Else
            
                CSql = "SELECT MAX(Id)+1 as NuevoId FROM CamposDelTrabajador"
                Set RsTemp2 = CrearRS(CSql)
                
                If Not IsNull(RsTemp2.Fields("NuevoId").Value) Then
                    NuevoId = Val(RsTemp2.Fields("NuevoId").Value)
                Else
                    NuevoId = 1
                End If
    
                CSql = "INSERT INTO CamposDelTrabajador (id,IdEmpleado,IdCampoNomina,ValorN,Tipo)" & _
                        "VALUES(" & NuevoId & ", " & IdEmpla & "," & CodConp & "," & Replace(CDbl(resultado), ",", ".") & ",'CO')"
                Set RsTemp2 = CrearRS(CSql)
            End If
            RsTemp.MoveNext
        Else
        ' Código ELSE casi innecesario, Evaluar.
            CSql = "SELECT MAX(Id)+1 as NuevoId FROM CamposDelTrabajador"
            Set RsTemp = CrearRS(CSql)
            
            If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
                NuevoId = Val(RsTemp.Fields("NuevoId").Value)
            Else
                NuevoId = 1
            End If

            CSql = "INSERT INTO CamposDelTrabajador VALUES (id,IdEmpleado,IdCampoNomina,Tipo) " & _
                    "VALUES(" & NuevoId & ", " & IdEmpla & "," & CodConp & ",'CO')"
            Set RsTemp = CrearRS(CSql)
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        End If
    End If
Next i
DMGrid2.PaintMGrid
Exit Sub

errty:
MsgBox "No se puede Guardar, Verifique los CAMPOS!", vbCritical + vbOKOnly, "Error"
GoTo Seguir

End Sub
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Procedimiento para saber si se borraron o agregaron campos al perfil del empleado actual
Sub Verificar_Cambios()
Dim CodCampo As Integer
Dim NuevoId As Integer
Dim VPredet As Double

CSql = "SELECT * FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla & " AND Tipo='CA'"
Set RsTemp = CrearRS(CSql)

While Not RsTemp.EOF

    Band = False
    For i = 1 To DMGrid1.Rows
        CodConp = Val(DMGrid1.ValorCelda(i, 1))
        If Val(RsTemp.Fields("IdCampoNomina").Value) = CodConp Then
            Band = True
            Exit For
        End If
    Next i
    
    If Band = False Then
        CSql = "DELETE FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla & " AND Tipo='CA' AND IdCampoNomina=" & Val(RsTemp.Fields("IdCampoNomina").Value)
        Set RsTemp2 = CrearRS(CSql)
    End If
    
    RsTemp.MoveNext
Wend

For i = 1 To DMGrid1.Rows
    CodCampo = Val(DMGrid1.ValorCelda(i, 1))
    
    ' Sentencia q verfica si existe el campo en el registro del empleado
    CSql = "Select * from CamposDelTrabajador where IdCampoNomina=" & CodCampo & " AND IdEmpleado=" & IdEmpla & " AND TIPO='CA'"
    Set RsTemp = CrearRS(CSql)
    
    ' condicional para saber si encontro registros en la consulta anterior
    If RsTemp.RecordCount = 0 Then
    
        ' consulta para saber el valor maximo del campo ID de la tabla CamposDelTrabajador y le suma UNO
        CSql = "Select MAX(id)+1 as NuevoId from CamposDelTrabajador"
        Set RsTemp = CrearRS(CSql)
        If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
            NuevoId = Val(RsTemp.Fields("NuevoId").Value)
        Else
            NuevoId = 1
        End If
        
        ' consulta para obtener el valor por DEFAULT del campo "CodCampo"
        CSql = "Select Predeterminado from CamposDeNomina where IdCampoNomina=" & CodCampo
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
            VPredet = CDbl(RsTemp.Fields("Predeterminado").Value)
            If Val(VPredet) = 0 Then VPredet = 0
        Else
            VPredet = 0
        End If
        
        ' sentencia que agrega un CAMPO al empleado
        CSql = "INSERT INTO CamposDelTrabajador(Id, IdCampoNomina, IdEmpleado,ValorN,Tipo) values(" & _
                NuevoId & "," & CodCampo & "," & IdEmpla & "," & VPredet & ",'CA')"
        Set RsTemp = CrearRS(CSql)
    Else
        ' Recalcular ese concepto
    End If
Next i
End Sub

Sub Conectar()
On Error GoTo MostrarError
If IdEmpla = 0 Then
    If UCase(TxtBuscar.Text) <> UCase("busqueda") Then
        CSql = "select * from empleados where cedula = " & Val(TxtBuscar.Text) & " or nombre like '%" & TxtBuscar.Text & "%' or apellido like '%" & TxtBuscar.Text & "%' and Activo=1 ORDER BY Apellido"
    Else
        CSql = "select * from empleados where Activo=1 ORDER BY Apellido"
    End If
Else
    CSql = "select * from empleados where idempleado = " & IdEmpla & " and Activo=1 ORDER BY Apellido"
End If

Set BdCampos = CrearRS(CSql)
If Not BdCampos.EOF Then
    Call CargaEmpleados
    Call CargaRenglon
    NoReg.Caption = "Registro " & BdCampos.AbsolutePosition & " / " & BdCampos.RecordCount
Else
    NoReg.Caption = "No hay registros de empleados!"
End If

Exit Sub
MostrarError:
MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source

End Sub

Private Sub DMGrid2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub DMGrid1_KeyPress(KeyAscii As Integer)

CodDMGrid1 = DMGrid1.ValorCelda(DMGrid1.Row, 1)
NomDMGrid1 = DMGrid1.ValorCelda(DMGrid1.Row, 2)

CSql = "SELECT Restringido FROM CamposDeNomina WHERE IdCampoNomina=" & CodDMGrid1
Set RsTemp = CrearRS(CSql)

If Val(RsTemp.Fields("Restringido").Value) = 1 Then
    MsgBox "El campo " & NomDMGrid1 & " no puede ser modificado!", vbExclamation + vbOKOnly, "NO SE PUEDE MODIFICAR!"
    KeyAscii = 0
End If

End Sub

Private Sub DMGrid1_TxtChange()
'modificacion
Cambio = 1
End Sub

Private Sub DMGrid2_TxtChange()
'modificacion
Cambio = 1
End Sub

Private Sub Form_Activate()

Tipo = "VCTrabajador"

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Sentencia para verificar y calcular segun
' el período que se generara la nómina
CSql = "SELECT Grupo.fecha_prox_gen, Grupo.fecha_prox_gen2, Grupo.periodo FROM Grupo INNER JOIN Empleados ON " & _
    "Grupo.Id_Grupo = Empleados.Id_Grupo Where (Empleados.IdEmpleado = " & IdEmpla & ")"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then
    FechaPeriodo = Format(Now, "dd/MM/yyyy")
    ' MMMMMM Calcula el período para la fecha actual MMMM
    Dim TMes As Integer
    Dim TDia As Integer
    TMes = Val(Format(Now, "m"))
    TDia = Val(Format(Now, "d"))
    If TDia > 15 Then
        TMes = TMes * 2
    Else
        TMes = TMes - 1
        If TMes = 0 Then
            TMes = 1
        Else
            TMes = TMes * 2
            TMes = TMes + 1
        End If
    End If
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    Periodo = TMes
Else
    ' MMMM Selecciona el período del registro actual MMMM
    FechaPeriodo = Format(RsTemp.Fields("fecha_prox_gen").Value, "dd/MM/yyyy")
    FechaFinPeriodo = Format(RsTemp.Fields("fecha_prox_gen2").Value, "dd/MM/yyyy")
    Periodo = Val(RsTemp.Fields("periodo").Value)
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

FrmValoresCampoTrabajador.Caption = "Valores de campo por trabajador, Inicio de Nómina " & FechaPeriodo
End Sub

Private Sub Form_Load()
Centrar Me
Call Grid

If IdEmpl <> 0 And IdEmpla = 0 Then IdEmpla = IdEmpl
Call Conectar
            
If IdEmpla <> 0 Then
    'BtnAgregar.Enabled = False
    BtnGuardar.Enabled = True
Else
    BtnGuardar.Enabled = False
    MsgBox "No existen Empleados activos en la Base de Datos!"
End If

End Sub

Private Sub Form_LostFocus()
IdEmpl = IdEmpla
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cambio = 1 And IdEmpla <> 0 Then
    Msg = "Los campos del trabajador " & Text1.Text & " " & Text2.Text & " han cambiado desea guradar estos cambios???"
    d = MsgBox(Msg, vbYesNo, "Guardar Cambios???")
    If d = vbYes Then Call Guardar
End If
IdEmpla = 0
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Change()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub TxtBuscar_LostFocus()
If Trim(TxtBuscar.Text) = "" Then TxtBuscar.Text = "Busqueda"
End Sub

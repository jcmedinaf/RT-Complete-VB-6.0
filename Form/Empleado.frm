VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEmpleados 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleado"
   ClientHeight    =   8025
   ClientLeft      =   2475
   ClientTop       =   2745
   ClientWidth     =   11250
   Icon            =   "Empleado.frx":0000
   LinkTopic       =   "Form38"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11250
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   7935
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   11055
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos de Empleado"
         Height          =   6855
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   10815
         Begin VB.TextBox TxtRif 
            Height          =   375
            Left            =   5520
            TabIndex        =   57
            Top             =   360
            Width           =   3255
         End
         Begin VB.PictureBox Picture1 
            Height          =   375
            Left            =   9480
            ScaleHeight     =   315
            ScaleWidth      =   555
            TabIndex        =   55
            ToolTipText     =   "Se usa para PROBAR si hay camaras web conectadas (ASI QUE NO BORRAR!!!)"
            Top             =   4560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox CboCodigoC 
            Height          =   315
            ItemData        =   "Empleado.frx":1002
            Left            =   5520
            List            =   "Empleado.frx":1004
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox CboCodigo 
            Height          =   315
            ItemData        =   "Empleado.frx":1006
            Left            =   1320
            List            =   "Empleado.frx":1008
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2430
            Width           =   975
         End
         Begin ChamaleonButton.ChameleonBtn BrnListaEmpleados 
            Height          =   615
            Left            =   8880
            TabIndex        =   29
            Top             =   3120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
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
            MICON           =   "Empleado.frx":100A
            PICN            =   "Empleado.frx":1026
            PICH            =   "Empleado.frx":12AF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnConceptosTrabajador 
            Height          =   615
            Left            =   8880
            TabIndex        =   28
            Top             =   2400
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "Conceptos del Trabajador"
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
            MICON           =   "Empleado.frx":16CA
            PICN            =   "Empleado.frx":16E6
            PICH            =   "Empleado.frx":1982
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   5520
            TabIndex        =   2
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox TxtCedulaEmp 
            Height          =   375
            Left            =   1320
            TabIndex        =   0
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox TxtApellido 
            Height          =   375
            Left            =   1320
            TabIndex        =   1
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox TxtCelular 
            Height          =   375
            Left            =   6480
            TabIndex        =   7
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox TxtDireccion 
            Height          =   975
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1320
            Width           =   7455
         End
         Begin VB.TextBox TxtTalla 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5760
            TabIndex        =   9
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox TxtEdad 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox TxtPeso 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   7800
            TabIndex        =   10
            Top             =   2880
            Width           =   975
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Datos de Nómina"
            Height          =   3255
            Left            =   120
            TabIndex        =   34
            Top             =   3480
            Width           =   8535
            Begin VB.ComboBox CboBanco 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   2760
               Width           =   4095
            End
            Begin VB.ComboBox CboDpto 
               Height          =   315
               ItemData        =   "Empleado.frx":1DB7
               Left            =   120
               List            =   "Empleado.frx":1DB9
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1320
               Width           =   4095
            End
            Begin VB.ComboBox CboGNomina 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   600
               Width           =   4095
            End
            Begin VB.ComboBox CboCargo 
               Height          =   315
               ItemData        =   "Empleado.frx":1DBB
               Left            =   4320
               List            =   "Empleado.frx":1DBD
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   1320
               Width           =   4095
            End
            Begin VB.ComboBox CboStatus 
               Height          =   315
               ItemData        =   "Empleado.frx":1DBF
               Left            =   6000
               List            =   "Empleado.frx":1DC1
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   600
               Width           =   2415
            End
            Begin VB.TextBox TxtNroCuenta 
               Height          =   375
               Left            =   4320
               TabIndex        =   19
               Top             =   2760
               Width           =   4095
            End
            Begin VB.TextBox TxtSueldo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   375
               Left            =   4320
               TabIndex        =   17
               Top             =   2040
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.ComboBox CboProfesion 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   2040
               Width           =   4095
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   12451843
               CurrentDate     =   39932
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Banco:"
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   2520
               Width           =   510
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Departamento:"
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   1080
               Width           =   1050
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Nómina:"
               Height          =   195
               Left            =   1800
               TabIndex        =   41
               Top             =   360
               Width           =   1065
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status:"
               Height          =   195
               Left            =   6000
               TabIndex        =   40
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sueldo a Devengar:"
               Height          =   195
               Left            =   4320
               TabIndex        =   39
               Top             =   1800
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Profesión:"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   1800
               Width           =   705
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Numero de Cuenta:"
               Height          =   195
               Left            =   4320
               TabIndex        =   37
               Top             =   2520
               Width           =   1380
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cargo:"
               Height          =   195
               Left            =   4320
               TabIndex        =   36
               Top             =   1080
               Width           =   465
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Ingreso:"
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   1290
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   12451843
            CurrentDate     =   39932
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   9720
            Top             =   5880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonButton.ChameleonBtn BtnTomarFotoEmpleado 
            Height          =   615
            Left            =   8880
            TabIndex        =   30
            Top             =   3840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "Tomar Foto"
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
            MICON           =   "Empleado.frx":1DC3
            PICN            =   "Empleado.frx":1DDF
            PICH            =   "Empleado.frx":245D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif:"
            Height          =   195
            Left            =   5160
            TabIndex        =   58
            Top             =   480
            Width           =   240
         End
         Begin VB.Label NoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   9000
            TabIndex        =   56
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telf. Hab:"
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   2490
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Cédula:"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   450
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   4680
            TabIndex        =   50
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Nac.:"
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   2970
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   930
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telf. Movil:"
            Height          =   195
            Left            =   4680
            TabIndex        =   46
            Top             =   2490
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Talla:"
            Height          =   195
            Left            =   5280
            TabIndex        =   45
            Top             =   2970
            Width           =   390
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso:"
            Height          =   195
            Left            =   7320
            TabIndex        =   44
            Top             =   2970
            Width           =   405
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edad:"
            Height          =   195
            Left            =   3360
            TabIndex        =   43
            Top             =   2970
            Width           =   420
         End
         Begin VB.Image Image3 
            BorderStyle     =   1  'Fixed Single
            Height          =   1695
            Left            =   8880
            Picture         =   "Empleado.frx":2AC1
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   32
         Top             =   7080
         Width           =   10815
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   9720
            TabIndex        =   27
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
            MICON           =   "Empleado.frx":5E42
            PICN            =   "Empleado.frx":5E5E
            PICH            =   "Empleado.frx":6027
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
            Left            =   1560
            TabIndex        =   21
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
            MICON           =   "Empleado.frx":625C
            PICN            =   "Empleado.frx":6278
            PICH            =   "Empleado.frx":6507
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
            TabIndex        =   20
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
            MICON           =   "Empleado.frx":6948
            PICN            =   "Empleado.frx":6964
            PICH            =   "Empleado.frx":6AF1
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
            Left            =   8520
            TabIndex        =   26
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
            MICON           =   "Empleado.frx":6D26
            PICN            =   "Empleado.frx":6D42
            PICH            =   "Empleado.frx":7024
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
            Left            =   7320
            TabIndex        =   25
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
            MICON           =   "Empleado.frx":7275
            PICN            =   "Empleado.frx":7291
            PICH            =   "Empleado.frx":7527
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
            Left            =   6720
            TabIndex        =   24
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
            MICON           =   "Empleado.frx":7786
            PICN            =   "Empleado.frx":77A2
            PICH            =   "Empleado.frx":7A37
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
            Left            =   3000
            TabIndex        =   22
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
            MICON           =   "Empleado.frx":7C93
            PICN            =   "Empleado.frx":7CAF
            PICH            =   "Empleado.frx":7E53
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
            TabIndex        =   23
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Imprimir"
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
            MICON           =   "Empleado.frx":7FF2
            PICN            =   "Empleado.frx":800E
            PICH            =   "Empleado.frx":8133
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
End
Attribute VB_Name = "FrmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDEmplea As New ADODB.Recordset
Dim BDEmp2 As New ADODB.Recordset
Public FotoE As String
Dim Cambio
Dim RegNew
Public RsEmpleado As New ADODB.Recordset
Dim RsCargaGrupo As New ADODB.Recordset
Dim RsCargos As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim RsTemp2 As New ADODB.Recordset
Dim RsCargaProfesion As New ADODB.Recordset
Dim IdEmpla As Integer

Sub Carga_CboCargo()

CSql = "Select * From Cargos"
Set RsCargos = CrearRS(CSql)
RsCargos.MoveFirst
CboCargo.Clear
Do While Not RsCargos.EOF
    CboCargo.AddItem RsCargos.Fields("Cargo")
    CboCargo.ItemData(CboCargo.NewIndex) = RsCargos.Fields("IdCargos")
    RsCargos.MoveNext
Loop
 
End Sub
Sub Carga_CboProfesion()

CSql = "Select * From Profesion"
Set RsCargaProfesion = CrearRS(CSql)
RsCargaProfesion.MoveFirst
CboProfesion.Clear
Do While Not RsCargaProfesion.EOF
    CboProfesion.AddItem RsCargaProfesion.Fields("Profesion")
    CboProfesion.ItemData(CboProfesion.NewIndex) = RsCargaProfesion.Fields("IdProf")
    RsCargaProfesion.MoveNext
Loop

End Sub

Private Sub BrnListaEmpleados_Click()
Tipo = "Nuevo Empleado"
BtnDesHacer_Click
FrmListadoEmpleados.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAgregar_Click()
'command2
Call Blanqueo
RegNew = 1
Cambio = 0
IdEmpla = 0
NoReg.Caption = "Nuevo Registro"
DTPicker1 = Now - 1800 ' fecha actual menos (casi) 5 años (1800 dias)
DTPicker2 = Now
Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
FotoE = "Silueta.jpg"
BrnListaEmpleados.Enabled = False
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnSiguiente.Enabled = False
BtnAnterior.Enabled = False
BtnImprimir.Enabled = False
Frame1.BackColor = &HE0E0E0
Frame2.BackColor = &HE0E0E0
End Sub

Private Sub BtnAnterior_Click()
'commad6
If RsEmpleado.RecordCount <> 0 Then
    Blanqueo
    RsEmpleado.MovePrevious
    If RsEmpleado.BOF Then RsEmpleado.MoveLast
    Call Empleado
Else
    MsgBox "No hay registros, inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros!"
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Public Sub Empleado()

If RsEmpleado.RecordCount = 0 Then
    IdEmpla = 0
    NoReg.Caption = "Registro 0 / 0"
    BtnAgregar.Enabled = True
    BtnEliminar.Enabled = False
    BtnSiguiente.Enabled = False
    BtnAnterior.Enabled = False
    BtnImprimir.Enabled = False
    Exit Sub
End If

IdEmpla = Val(RsEmpleado.Fields("idempleado"))

BtnEliminar.Enabled = True
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
BtnImprimir.Enabled = True

TxtNombre.Text = RsEmpleado.Fields("Nombre")
TxtApellido.Text = RsEmpleado.Fields("Apellido")
TxtCedulaEmp.Text = RsEmpleado.Fields("cedula")

For i = 0 To CboCodigo.ListCount - 1
    'MsgBox InStr(1, CboCodigo.List(i), Trim(RsEmpleado.Fields("Codigo").Value), vbTextCompare)
    If InStr(CboCodigo.List(i), Val(RsEmpleado.Fields("Codigo").Value)) Then
        CboCodigo.ListIndex = i
        Exit For
    End If
Next i

TxtTelefono.Text = RsEmpleado.Fields("Celular")

For i = 0 To CboCodigoC.ListCount - 1
    If InStr(CboCodigoC.List(i), Val(RsEmpleado.Fields("CodigoC").Value)) Then
        CboCodigoC.ListIndex = i
        Exit For
    End If
Next i

If Not IsNull(RsEmpleado.Fields("Telefono")) Then TxtCelular.Text = RsEmpleado.Fields("Telefono") Else TxtCelular.Text = ""
If Not IsNull(RsEmpleado.Fields("Direccion")) Then TxtDireccion.Text = RsEmpleado.Fields("Direccion") Else TxtDireccion.Text = ""
If Not IsNull(RsEmpleado.Fields("Talla")) Then TxtTalla.Text = RsEmpleado.Fields("Talla") Else TxtTalla.Text = ""
If Not IsNull(RsEmpleado.Fields("Peso")) Then TxtPeso.Text = Format(RsEmpleado.Fields("Peso"), "#,##0.00") Else TxtPeso.Text = ""
If Not IsNull(RsEmpleado.Fields("Edad")) Then TxtEdad.Text = RsEmpleado.Fields("Edad") Else TxtEdad.Text = ""
If Not IsNull(RsEmpleado.Fields("Cuenta_Ban").Value) Then TxtNroCuenta.Text = RsEmpleado.Fields("Cuenta_Ban") Else TxtNroCuenta.Text = ""
If Not IsNull(RsEmpleado.Fields("Sueldo")) Then TxtSueldo.Text = Format(RsEmpleado.Fields("Sueldo"), "#,##0.00") Else TxtSueldo.Text = ""
If Not IsNull(RsEmpleado.Fields("Rif")) Then TxtRif.Text = RsEmpleado.Fields("Rif") Else TxtRif.Text = ""

NoReg.Caption = "Registro " & RsEmpleado.AbsolutePosition & " / " & RsEmpleado.RecordCount

If Not IsNull(RsEmpleado.Fields("Photo")) Then
    If RsEmpleado.Fields("Photo").Value <> "" And Dir(FotoEmp & "\" & RsEmpleado.Fields("Photo").Value) <> "" Then
        Image3.Picture = LoadPicture(FotoEmp & "\" & RsEmpleado.Fields("Photo").Value)
        FotoE = RsEmpleado.Fields("Photo").Value
    Else
        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture: FotoE = ""
    End If
Else
    Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture: FotoE = ""
End If


DTPicker1.Value = RsEmpleado.Fields("Fecha_Nac")
DTPicker2.Value = RsEmpleado.Fields("Fecha_Ing")

For T = 0 To CboGNomina.ListCount - 1
    If CboGNomina.ItemData(T) = RsEmpleado.Fields("id_grupo") Then
        CboGNomina.ListIndex = T
        Exit For
    Else
        CboGNomina.ListIndex = -1
    End If
Next T

For T = 0 To CboStatus.ListCount - 1
    If CboStatus.ItemData(T) = RsEmpleado.Fields("Status") Then
        CboStatus.ListIndex = T
        Exit For
    End If
Next T

For T = 0 To CboCargo.ListCount - 1
    If CboCargo.ItemData(T) = RsEmpleado.Fields("cargo") Then
        CboCargo.ListIndex = T
        Exit For
    End If
Next T
For T = 0 To CboProfesion.ListCount - 1
    If CboProfesion.ItemData(T) = RsEmpleado.Fields("Profesion") Then
        CboProfesion.ListIndex = T
        Exit For
    End If
Next T
For T = 0 To CboDpto.ListCount - 1
    If CboDpto.ItemData(T) = RsEmpleado.Fields("Departamentos") Then
        CboDpto.ListIndex = T
        Exit For
    End If
Next T


For T = 0 To CboBanco.ListCount - 1
    If CboBanco.ItemData(T) = RsEmpleado.Fields("IdBanco") Then
        CboBanco.ListIndex = T
        Exit For
    End If
Next T

Cambia = 0

Reg_Actual(0) = IdEmpla
Reg_Actual(1) = TxtNombre.Text
Reg_Actual(2) = TxtApellido.Text
Reg_Actual(3) = TxtCedulaEmp.Text  'CboCodigo.ItemData(CboCodigo.NewIndex)
Reg_Actual(4) = Trim(CboCodigo.Text)
Reg_Actual(5) = TxtTelefono.Text
Reg_Actual(6) = CboCodigoC.Text
Reg_Actual(7) = TxtCelular.Text
Reg_Actual(8) = TxtDireccion.Text
Reg_Actual(9) = TxtTalla.Text
Reg_Actual(10) = TxtPeso.Text
Reg_Actual(11) = TxtEdad.Text
Reg_Actual(12) = CboStatus.Text
Reg_Actual(13) = TxtNroCuenta.Text
Reg_Actual(14) = TxtSueldo.Text
Reg_Actual(15) = CboDpto.Text
Reg_Actual(16) = CboCargo.Text
Reg_Actual(17) = CboProfesion.Text
Reg_Actual(18) = DTPicker1.Value
Reg_Actual(19) = DTPicker2.Value
Reg_Actual(20) = CboGNomina.Text
Reg_Actual(21) = TxtRif.Text

End Sub

Sub Blanqueo()
TxtNombre.Text = ""
TxtRif.Text = ""
TxtApellido.Text = ""
TxtCedulaEmp.Text = ""
TxtCelular.Text = ""
TxtTelefono.Text = ""
TxtDireccion.Text = ""
TxtPeso.Text = ""
TxtTalla.Text = ""
TxtSueldo.Text = ""
TxtEdad.Text = ""
TxtNroCuenta.Text = ""
'Image3.Picture = LoadPicture()
CboStatus.ListIndex = -1
CboCargo.ListIndex = -1
CboProfesion.ListIndex = -1
CboGNomina.ListIndex = -1
DTPicker1.Value = Now
DTPicker2.Value = Now
End Sub

Private Sub BtnConceptosTrabajador_Click()
If IdEmpla = 0 Then Exit Sub

IdEmpl = IdEmpla
FrmValoresCampoTrabajador.Show
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
BrnListaEmpleados.Enabled = True
BtnAgregar.Enabled = True
Form_Load

Frame1.BackColor = &HEAEFEF
Frame2.BackColor = &HEAEFEF
End Sub

Private Sub BtnEliminar_Click()
On Error GoTo MostrarError
Dim resp

If Val(IdEmpla) = 0 Then MsgBox "Debe seleccionar un empleado!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Esta seguro que desea eliminar el registro actual?", vbQuestion + vbYesNo, "Confirmar")
If resp = 7 Then Exit Sub

CSql = "UPDATE Empleados SET Activo=0 Where IdEmpleado=" & IdEmpla
Set RsTemp = CrearRS(CSql)

CSql = "DELETE FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla
Set RsTemp = CrearRS(CSql)

MsgBox "El registro ha sido Eliminado satisfactoriamente!", vbInformation + vbOKOnly, "Operación Exitosa!"
BtnDesHacer_Click

Exit Sub
MostrarError:
    MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error GoTo MostrarError
Dim NuevoId, resp, resp2
Dim IdGrupo
'command1

If IdEmpla = 0 And RegNew = 0 Then MsgBox "Seleccione un empleado!", vbExclamation + vbOKOnly, "Error": Exit Sub
If IdEmpla <> 0 And RegNew = 1 Then MsgBox "Alerta! no se puedo guardar el nuevo empleado, Presione el Boton AGREGAR e intente de nuevo": Exit Sub

If CboDpto.ListIndex = -1 Then
    MsgBox "Seleccione un Departamento!!", vbExclamation + vbOKOnly, "Faltan datos"
    CboDpto.SetFocus
    Exit Sub
ElseIf CboProfesion.ListIndex = -1 Then
    MsgBox "Seleccione una Profesión!!", vbExclamation + vbOKOnly, "Faltan datos"
    CboProfesion.SetFocus
    Exit Sub
ElseIf CboCargo.ListIndex = -1 Then
    MsgBox "Seleccione un Cargo!!", vbExclamation + vbOKOnly, "Faltan datos"
    CboCargo.SetFocus
    Exit Sub
ElseIf CboBanco.ListIndex = -1 Then
    MsgBox "Seleccione un Banco!!", vbExclamation + vbOKOnly, "Faltan datos"
    CboBanco.SetFocus
    Exit Sub
ElseIf CboGNomina.ListIndex = -1 Then
    MsgBox "Seleccione un Grupo de Nómina!!", vbExclamation + vbOKOnly, "Faltan datos"
    CboGNomina.SetFocus
    Exit Sub
ElseIf CboStatus.ListIndex = -1 Then
    MsgBox "Seleccione un Status!!", vbExclamation + vbOKOnly, "Faltan datos"
    CboStatus.SetFocus
    Exit Sub
End If
'verifica si hay cambio si no los hay emite un mensaje y sale del procedimiento sino continua
If Cambio = 0 Then
    Msg = "no hay cambios para agregar o actualizar registros"
    MsgBox Msg, vbOKOnly, "No hay cambios"
    Exit Sub
End If

If Trim(TxtTalla) = "" Then TxtTalla = "0"
If Trim(TxtPeso) = "" Then TxtPeso = "0"
If Trim(TxtRif) = "" Then TxtTalla = "0"

resp = MsgBox("Se procedera a guardar los cambios, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar Operación")
If resp = 7 Then Exit Sub

CSql = "SELECT * FROM Empleados Where cedula='" & Val(TxtCedulaEmp.Text) & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If RegNew = 1 Then
        MsgBox "La Cedula ingresada ya se encuentra registrada!", vbCritical + vbOKOnly, "Error"
        TxtCedulaEmp.Text = ""
        TxtCedulaEmp.SetFocus
        Exit Sub
    Else
        If Val(Reg_Actual(3)) <> Val(TxtCedulaEmp.Text) Then
            MsgBox "La Cedula ingresada ya se encuentra registrada!", vbCritical + vbOKOnly, "Error"
            TxtCedulaEmp.Text = Reg_Actual(3)
            TxtCedulaEmp.SetFocus
            Exit Sub
        End If
    End If
End If

' sentencia que busca el numero MAX del campo IdEmpleado en la tabla EMPLEADOS, y le suma UNO
' para asi Tener un ID Unico en el caso de que sea un nuevo registro
CSql = "Select MAX(IdEmpleado)+1 as NuevoId From Empleados"
Set BDEmplea = CrearRS(CSql)

If Not IsNull(BDEmplea.Fields("NuevoId").Value) Then
    NuevoId = BDEmplea.Fields("NuevoId").Value
Else
    NuevoId = "1"
End If
    
Dim Fecha_Egr As String

If CboStatus.ListIndex = 4 Then
    Fecha_Egr = Format(Now, "dd/MM/yyyy")
Else
    Fecha_Egr = ""
End If

'Verifica si el registro actual en el formulario es un nuevo registro o es uno ya guardado que va a ser actualizado
'segun la variable "regnew"
Select Case RegNew
Case Is = 0       'actualiza

    If Reg_Actual(20) <> CboGNomina.Text Then
        resp = MsgBox("Realizó un cambio de Departamento, se procederá a crear los campos correspondientes!", vbExclamation + vbOKCancel, "Confirmar")
        If resp = vbCancel Then: MsgBox "No se han realizado cambios!", vbExclamation + vbOKOnly, "Operación Cancelada!": Exit Sub

        CSql = "DELETE FROM CamposDelTrabajador Where IdEmpleado=" & IdEmpla
        
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        ' Busca el Empleado
        
        For i = 0 To CboGNomina.ListCount - 1
        
            If Trim(Reg_Actual(20)) <> "" Then
                If Reg_Actual(20) = CboGNomina.List(i) Then
                    IdGrupo = Val(CboGNomina.ItemData(i))
                    Exit For
                End If
            Else
                IdGrupo = Val(CboGNomina.ItemData(i))
                Exit For
            End If
        Next i

        CSql = "SELECT  CamposDelTrabajador.* FROM Relacion_Campo INNER JOIN " & _
            " Grupo ON Relacion_Campo.Id_Grupo = Grupo.Id_Grupo INNER JOIN " & _
            " CamposDelTrabajador ON Relacion_Campo.Id_Campo = CamposDelTrabajador.IdCampoNomina " & _
            " Where (CamposDelTrabajador.IdEmpleado = " & IdEmpla & ") And (Grupo.Id_Grupo = " & IdGrupo & ") AND Tipo='CA'"

        Set bdgrupo2 = CrearRS(CSql)
        
        If bdgrupo2.RecordCount <> 0 Then
            ' ELIMINA los CAMPOS del Empleado segun el Grupo Anterior, si tiene campos q no pertenecen al Grupo
            ' entonces los mantiene en el registro de ese empleado
            bdgrupo2.MoveFirst
            While Not bdgrupo2.EOF
                CSql = "DELETE CamposDelTrabajador WHERE IdEmpleado=" & IdEmpla & " AND IdCampoNomina=" & Val(bdgrupo2.Fields("IdCampoNomina").Value) & " AND Tipo='CA'"
                Set RsTemp = CrearRS(CSql)
                bdgrupo2.MoveNext
            Wend
        End If
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        Set RsTemp = CrearRS(CSql)
        
        CSql = "SELECT * FROM Relacion_Campo Where Id_Grupo=" & CboGNomina.ItemData(CboGNomina.ListIndex) & " order by Id_Campo"
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
            RsTemp.MoveFirst
            While Not RsTemp.EOF
                
                ' Obtiene la Nueva id para la tabla CamposDelTrabajador
                CSql = "SELECT MAX(Id)+1 as NuevoId FROM CamposDelTrabajador"
                Set RsTemp2 = CrearRS(CSql)
                
                If Not IsNull(RsTemp2.Fields("NuevoId").Value) Then
                    NuevoId = RsTemp2.Fields("NuevoId").Value
                Else
                    NuevoId = "1"
                End If
                
                ' Obtiene el Valor PREDETERMINADO del campo Nomina
                CSql = "SELECT Predeterminado FROM CamposDeNomina where IdCampoNomina=" & RsTemp.Fields("Id_Campo").Value
                Set RsTemp2 = CrearRS(CSql)
                
                resp = RsTemp2.Fields("Predeterminado").Value
                
                CSql = "INSERT INTO CamposDelTrabajador (Id,IdEmpleado,IdCampoNomina,ValorN,Tipo) VALUES " & _
                       "(" & NuevoId & "," & IdEmpla & "," & RsTemp.Fields("Id_Campo").Value & "," & resp & ",'CA')"
                Set RsTemp2 = CrearRS(CSql)
                
                RsTemp.MoveNext
            Wend
        End If
    End If
    
    CSql = "Select * From Empleados where IdEmpleado = '" & IdEmpla & "'"
    Set BDEmplea = CrearRS(CSql)
    'BDEmplea.AddNew
    BDEmplea.Fields("Nombre").Value = TxtNombre.Text
    BDEmplea.Fields("Apellido").Value = TxtApellido.Text
    BDEmplea.Fields("cedula").Value = TxtCedulaEmp.Text
    BDEmplea.Fields("Fecha_Nac").Value = CDate(Format(DTPicker1.Value, "dd/mm/yyyy"))
    BDEmplea.Fields("Fecha_Ing").Value = CDate(Format(DTPicker2.Value, "dd/mm/yyyy"))
    BDEmplea.Fields("Codigo").Value = Replace(Replace(CboCodigo.Text, "(", ""), ")", "")
    BDEmplea.Fields("Telefono").Value = TxtCelular.Text
    BDEmplea.Fields("CodigoC").Value = Replace(Replace(CboCodigoC.Text, "(", ""), ")", "")
    BDEmplea.Fields("Celular").Value = TxtTelefono.Text
    BDEmplea.Fields("Direccion").Value = TxtDireccion.Text
    BDEmplea.Fields("Departamentos").Value = CboDpto.ItemData(CboDpto.ListIndex)
    BDEmplea.Fields("cargo").Value = CboCargo.ItemData(CboCargo.ListIndex)
    BDEmplea.Fields("Talla").Value = TxtTalla.Text
    BDEmplea.Fields("Peso").Value = CDbl(TxtPeso.Text)
    BDEmplea.Fields("Edad").Value = TxtEdad.Text
    BDEmplea.Fields("Status").Value = CboStatus.ItemData(CboStatus.ListIndex)
    BDEmplea.Fields("Profesion").Value = CboProfesion.ItemData(CboProfesion.ListIndex)
    BDEmplea.Fields("IdBanco").Value = CboBanco.ItemData(CboBanco.ListIndex)
    BDEmplea.Fields("Cuenta_Ban").Value = TxtNroCuenta.Text
    BDEmplea.Fields("Photo").Value = FotoE
    BDEmplea.Fields("Sueldo").Value = CDbl(TxtSueldo.Text)
    BDEmplea.Fields("id_grupo").Value = CboGNomina.ItemData(CboGNomina.ListIndex)
    BDEmplea.Fields("Rif").Value = Trim(TxtRif.Text)
    
    If Fecha_Egr <> "" Then BDEmplea.Fields("Fecha_Egr").Value = Fecha_Egr

    BDEmplea.Update
    
    MsgBox "Registro Actualizado Satisfactoriamente!", vbInformation + vbOKOnly, "Operacion Exitosa!"
                                                                              
Case Is = 1 'agrega
    CSql = "Insert Into Empleados(IdEmpleado,idusuario,Nombre,Apellido,cedula,Fecha_Nac,Fecha_Ing,Codigo," & _
        "Telefono,CodigoC,Celular,Direccion,Departamentos,cargo,Talla,Peso,Edad,Status,Profesion,Cuenta_Ban," & _
        "Photo,Sueldo,id_grupo, Activo, IdBanco,Rif,Fecha_Egr) VALUES(" & NuevoId & "," & IdUser & ",'" & TxtNombre.Text & "','" & TxtApellido.Text & _
        "'," & Val(TxtCedulaEmp.Text) & ",'" & Format(DTPicker1.Value, "dd/mm/yyyy") & "','" & _
        Format(DTPicker2.Value, "dd/mm/yyyy") & "','" & Replace(Replace(CboCodigo.Text, "(", ""), ")", "") & "'," & Val(TxtCelular.Text) & _
        ",'" & Replace(Replace(CboCodigoC.Text, "(", ""), ")", "") & "'," & Val(TxtTelefono.Text) & ", '" & TxtDireccion.Text & "', " & _
        CboDpto.ItemData(CboDpto.ListIndex) & "," & CboCargo.ItemData(CboCargo.ListIndex) & ", '" & TxtTalla.Text & "', " & CDbl(TxtPeso.Text) & _
        ", " & Val(TxtEdad.Text) & ", " & CboStatus.ItemData(CboStatus.ListIndex) & ", " & _
        CboProfesion.ItemData(CboProfesion.ListIndex) & ", '" & TxtNroCuenta.Text & "', '" & FotoE & _
        "', " & Val(TxtSueldo.Text) & ", " & CboGNomina.ItemData(CboGNomina.ListIndex) & ",1," & _
        CboBanco.ItemData(CboBanco.ListIndex) & ", '" & Trim(TxtRif.Text) & "','" & Fecha_Egr & "')"
    Set BDEmplea = CrearRS(CSql)
    
    CSql = "SELECT * FROM Relacion_Campo Where Id_Grupo=" & CboGNomina.ItemData(CboGNomina.ListIndex) & " order by Id_Campo"
    Set RsTemp = CrearRS(CSql)
    
    RsTemp.MoveFirst
    While Not RsTemp.EOF
        
        ' Obtiene la Nueva id para la tabla CamposDelTrabajador
        CSql = "SELECT MAX(Id)+1 as NuevoId FROM CamposDelTrabajador"
        Set RsTemp2 = CrearRS(CSql)
        
        If Not IsNull(RsTemp2.Fields("NuevoId").Value) Then
            resp = RsTemp2.Fields("NuevoId").Value
        Else
            resp = "1"
        End If
        
        ' Obtiene el Valor PREDETERMINADO del campo Nomina
        CSql = "SELECT Predeterminado FROM CamposDeNomina where IdCampoNomina=" & RsTemp.Fields("Id_Campo").Value
        Set RsTemp2 = CrearRS(CSql)
        
        resp2 = RsTemp2.Fields("Predeterminado").Value
        
        CSql = "INSERT INTO CamposDelTrabajador (Id,IdEmpleado,IdCampoNomina,ValorN,Tipo) VALUES " & _
               "(" & resp & "," & NuevoId & "," & RsTemp.Fields("Id_Campo").Value & "," & resp2 & ", 'CA')"
        Set RsTemp2 = CrearRS(CSql)
        
        RsTemp.MoveNext
    Wend
    MsgBox "Registro Agregado Satisfactoriamente!", vbInformation + vbOKOnly, "Operacion Exitosa!"
End Select


If RegNew = 0 Then
    If Reg_Actual(1) <> TxtNombre.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo NOMBRE de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtNombre.Text & ")   valor anterior (" & Reg_Actual(1) & ")")
    If Reg_Actual(2) <> TxtApellido.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo APELLIDO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtApellido.Text & ")   valor anterior (" & Reg_Actual(2) & ")")
    If Reg_Actual(3) <> TxtCedulaEmp.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo CEDULA de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtCedulaEmp.Text & ")   valor anterior (" & Reg_Actual(3) & ")")
    If Reg_Actual(4) <> Trim(CboCodigo.Text) Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo CODIGO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboCodigo.ItemData(CboCodigo.ListIndex) & ")   valor anterior (" & Reg_Actual(4) & ")")
    If Reg_Actual(5) <> TxtTelefono.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo TELEFONO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtTelefono.Text & ")   valor anterior (" & Reg_Actual(5) & ")")
    If Reg_Actual(6) <> CboCodigoC.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo CODIGOC de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboCodigoC.ItemData(CboCodigoC.ListIndex) & ")   valor anterior (" & Reg_Actual(6) & ")")
    If Reg_Actual(7) <> TxtCelular.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo CELULAR de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtCelular.Text & ")   valor anterior (" & Reg_Actual(7) & ")")
    If Reg_Actual(8) <> TxtDireccion.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo DIRECCION de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtDireccion.Text & ")   valor anterior (" & Reg_Actual(8) & ")")
    If Reg_Actual(9) <> TxtTalla.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo TALLA de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtTalla.Text & ")   valor anterior (" & Reg_Actual(9) & ")")
    If Reg_Actual(10) <> TxtPeso.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo PESO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtPeso.Text & ")   valor anterior (" & Reg_Actual(10) & ")")
    If Reg_Actual(11) <> TxtEdad.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo EDAD de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtEdad.Text & ")   valor anterior (" & Reg_Actual(11) & ")")
    If Reg_Actual(12) <> CboStatus.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo STATUS de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboStatus.ItemData(CboStatus.ListIndex) & ")   valor anterior (" & Reg_Actual(12) & ")")
    If Reg_Actual(13) <> TxtNroCuenta.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo CUENTA_BAN de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtNroCuenta.Text & ")   valor anterior (" & Reg_Actual(13) & ")")
    If Reg_Actual(14) <> TxtSueldo.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo SUELDO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & TxtSueldo.Text & ")   valor anterior (" & Reg_Actual(14) & ")")
    If Reg_Actual(15) <> CboDpto.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo DEPARTAMENTOS de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboDpto.ItemData(CboDpto.ListIndex) & ")   valor anterior (" & Reg_Actual(15) & ")")
    If Reg_Actual(16) <> CboCargo.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo CARGO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboCargo.ItemData(CboCargo.ListIndex) & ")   valor anterior (" & Reg_Actual(16) & ")")
    If Reg_Actual(17) <> CboProfesion.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo PROFESION de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboProfesion.ItemData(CboProfesion.ListIndex) & ")   valor anterior (" & Reg_Actual(17) & ")")
    If Format(Reg_Actual(18), "DD/MM/YY") <> Format(DTPicker1.Value, "DD/MM/YY") Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo FECHA_NAC de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & Format(DTPicker1.Value, "DD/MM/YY") & ")   valor anterior (" & Reg_Actual(18) & ")")
    If Format(Reg_Actual(19), "DD/MM/YY") <> Format(DTPicker2.Value, "DD/MM/YY") Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo FECHA_ING de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & Format(DTPicker2.Value, "DD/MM/YY") & ")   valor anterior (" & Reg_Actual(19) & ")")
    If Reg_Actual(20) <> CboGNomina.Text Then Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "MODIFICAR", "Se modifico el campo ID_GRUPO de la tabla Empleados, IdEmpleado Nro. " & IdEmpla & ", nuevo valor (" & CboGNomina.ItemData(CboGNomina.ListIndex) & ")   valor anterior (" & Reg_Actual(20) & ")")
Else
    Call Enviar_Bitacora(IdUser, "NUEVO EMPLEADO", "AGREGAR", "Se agrego un Nuevo Empleado cuya IdEmpleado=" & NuevoId)
End If
'cierra la conexion actual a la tabla de empleados y la vuelve abrir mostrando el ultimo registro agregado
 
Form_Load
'RsEmpleado.MoveFirst
'Call Empleado
Cambio = 0
RegNew = 0
Exit Sub
MostrarError:
MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source

End Sub

Private Sub BtnImprimir_Click()

''========= ESTE ES EL CODIGO NUEVO ==========

'With CrystalReport1
'    .ReportFileName = RutaInformes & "\HojaEmpleados.rpt"
'    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
'    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
'    .DiscardSavedData = True
'    .RetrieveDataFiles
'    .ReportSource = 0
'    .SelectionFormula = "{DatosEmpleados.Cedula} = '" & TxtCedulaEmp.Text & "'"
'    .WindowTitle = "Hoja de Vida"
'    .Destination = crptToWindow
'    .PrintFileType = crptCrystal
'    .WindowState = crptMaximized
'    .WindowMaxButton = False
'    .WindowMinButton = False
'    .Action = 1
'End With

Dim sImagePath As String
Dim imgHeaderPicture As RptImage


CSql = "Select * From DatosEmpleados where Cedula = '" & TxtCedulaEmp.Text & "'"
Set RsReporte = CrearRS(CSql)

Load DrptHojaVida
Set DrptHojaVida.DataSource = RsReporte

DrptHojaVida.Sections("Sección2").Controls("LblApellidos").Caption = Trim(RsReporte.Fields("Apellido").Value)
DrptHojaVida.Sections("Sección2").Controls("LblNombres").Caption = Trim(RsReporte.Fields("Nombre").Value)
DrptHojaVida.Sections("Sección2").Controls("LblCedula").Caption = Trim(RsReporte.Fields("Cedula").Value)
DrptHojaVida.Sections("Sección2").Controls("LblRif").Caption = Trim(RsReporte.Fields("Rif").Value)

DrptHojaVida.Sections("Sección2").Controls("LblTelefonoHab").Caption = "(" & Trim(RsReporte.Fields("Codigo").Value) & ")" & " " & Trim(RsReporte.Fields("Telefono").Value)
DrptHojaVida.Sections("Sección2").Controls("LblTelefonoCel").Caption = "(" & Trim(RsReporte.Fields("CodigoC").Value) & ")" & " " & Trim(RsReporte.Fields("Celular").Value)
DrptHojaVida.Sections("Sección2").Controls("LblDireccion").Caption = Trim(RsReporte.Fields("Direccion").Value)

DrptHojaVida.Sections("Sección2").Controls("LblFechaNac").Caption = Trim(RsReporte.Fields("Fecha_Nac").Value)
DrptHojaVida.Sections("Sección2").Controls("LblEdad").Caption = Trim(RsReporte.Fields("Edad").Value)
DrptHojaVida.Sections("Sección2").Controls("LblPeso").Caption = Trim(RsReporte.Fields("Peso").Value)
DrptHojaVida.Sections("Sección2").Controls("LblTalla").Caption = Trim(RsReporte.Fields("Talla").Value)
DrptHojaVida.Sections("Sección2").Controls("LblProfesion").Caption = Trim(RsReporte.Fields("Profesion").Value)

DrptHojaVida.Sections("Sección2").Controls("LblFechaIng").Caption = Trim(RsReporte.Fields("Fecha_Ing").Value)

If Not IsNull(Trim(RsReporte.Fields("Fecha_Egr").Value)) Then
    DrptHojaVida.Sections("Sección2").Controls("LblFechaEgre").Caption = Trim(RsReporte.Fields("Fecha_Egr").Value)
Else
    DrptHojaVida.Sections("Sección2").Controls("LblFechaEgre").Caption = ""
End If
DrptHojaVida.Sections("Sección2").Controls("LblCargo").Caption = Trim(RsReporte.Fields("Cargo").Value)
DrptHojaVida.Sections("Sección2").Controls("LblDepartamento").Caption = Trim(RsReporte.Fields("Descripcion").Value)

DrptHojaVida.Sections("Sección2").Controls("LblBanco").Caption = Trim(RsReporte.Fields("Expr1").Value)
DrptHojaVida.Sections("Sección2").Controls("LblNoCuenta").Caption = Trim(RsReporte.Fields("Cuenta_Ban").Value)


If RsReporte.Fields("Photo").Value <> "" Then

sImagePath = FotoEmp & "\" & RsReporte.Fields("Photo").Value 'path de la imagen
Set imgHeaderPicture = DrptHojaVida.Sections("Sección2").Controls("ImgFoto")
Set imgHeaderPicture.Picture = LoadPicture(sImagePath)
End If

DrptHojaVida.Show
End Sub

Private Sub BtnSiguiente_Click()
'command7

If RsEmpleado.RecordCount <> 0 Then
    Blanqueo
    RsEmpleado.MoveNext
    If RsEmpleado.EOF Then RsEmpleado.MoveFirst
    Call Empleado
Else
    MsgBox "No hay registros, inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay registros!"
End If
End Sub

Private Sub BtnTomarFotoEmpleado_Click()

'If TxtCedulaEmp.Text <> "" And TxtApellido.Text <> "" And TxtNombre.Text <> "" Then
If IdEmpla <> 0 Then
    If (Validar_Camara("CapWindow", ws_child Or ws_visible, 0, 0, 340, 240, Picture1.hwnd, 0)) Then
        FrmCapturarFoto.Show vbModal, FrmPrincipal
    Else
        MsgBox "No hay Camaras Webs Instaladas", vbOKOnly + vbCritical, "Error"
    End If
Else
    MsgBox "Debe de Ingresar o Seleccionar a un Empleado para poder tomar la Foto", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If
End Sub

Private Sub DTPicker1_Click()
TxtEdad.Text = DateDiff("yyyy", DTPicker1.Value, Now)
End Sub

Private Sub Form_Activate()
Tipo = "Nuevo Empleado"
End Sub

Private Sub Form_Load()
Centrar Me

For i = 0 To 30
    Reg_Actual(i) = ""
Next i
CboStatus.Clear
CboCargo.Clear
CboProfesion.Clear
CboGNomina.Clear
CboDpto.Clear
CboBanco.Clear
CboStatus.Clear
CboStatus.AddItem "Fijo"
CboStatus.ItemData(CboStatus.NewIndex) = 0
'CboStatus.AddItem "Pasante"
'CboStatus.ItemData(CboStatus.NewIndex) = 1
CboStatus.AddItem "Contratado"
CboStatus.ItemData(CboStatus.NewIndex) = 1
CboStatus.AddItem "Suspendido"
CboStatus.ItemData(CboStatus.NewIndex) = 2
CboStatus.AddItem "Reposo"
CboStatus.ItemData(CboStatus.NewIndex) = 3
CboStatus.AddItem "Retirado"
CboStatus.ItemData(CboStatus.NewIndex) = 4

CSql = "SELECT * FROM Codigos_T"
Set RsCodigos = CrearRS(CSql)
CboCodigo.Clear

CboCodigo.AddItem " "
CboCodigo.ItemData(CboCodigo.NewIndex) = 0
    
CboCodigoC.AddItem " "
CboCodigoC.ItemData(CboCodigoC.NewIndex) = 0
    
    
While Not RsCodigos.EOF
    CboCodigo.AddItem RsCodigos.Fields("Codigo").Value
    CboCodigo.ItemData(CboCodigo.NewIndex) = RsCodigos.Fields("IdCodigoT").Value
        
    CboCodigoC.AddItem RsCodigos.Fields("Codigo").Value
    CboCodigoC.ItemData(CboCodigoC.NewIndex) = RsCodigos.Fields("IdCodigoT").Value
    RsCodigos.MoveNext
Wend
CSql = "select * from empleados where activo =1 ORDER BY Fecha_Ing"
Set RsEmpleado = CrearRS(CSql)
Carga_CboCargo
Carga_CboProfesion
Carga_Grupo
Carga_Departamentos
Carga_Banco
Empleado
Cambio = 0
RegNew = 0
End Sub
Sub Carga_Grupo()

CSql = "Select * From Grupo"
Set RsCargaGrupo = CrearRS(CSql)
If Not RsCargaGrupo.EOF Then
    RsCargaGrupo.MoveFirst
    Do While Not RsCargaGrupo.EOF
        CboGNomina.AddItem RsCargaGrupo.Fields("Descripcion")
        CboGNomina.ItemData(CboGNomina.NewIndex) = RsCargaGrupo.Fields("id_grupo")
        RsCargaGrupo.MoveNext
    Loop
End If

End Sub

Sub Carga_Banco()

CSql = "Select * From CajasBancos"
Set RsCargaBancos = CrearRS(CSql)
If Not RsCargaBancos.EOF Then
    RsCargaBancos.MoveFirst
    Do While Not RsCargaBancos.EOF
        CboBanco.AddItem RsCargaBancos.Fields("Descripcion").Value
        CboBanco.ItemData(CboBanco.NewIndex) = RsCargaBancos.Fields("IdCajaBanco").Value
        RsCargaBancos.MoveNext
    Loop
End If

End Sub

Sub Carga_Departamentos()

CSql = "Select * From Departamentos"
Set RsCargaDepart = CrearRS(CSql)
If Not RsCargaDepart.EOF Then
    RsCargaDepart.MoveFirst
    Do While Not RsCargaDepart.EOF
        CboDpto.AddItem RsCargaDepart.Fields("Descripcion").Value
        CboDpto.ItemData(CboDpto.NewIndex) = RsCargaDepart.Fields("IdDepartamento").Value
        RsCargaDepart.MoveNext
    Loop
End If

End Sub
Private Sub DTPicker2_Change()
Cambio = 1
'TxtEdad.Text = DateDiff("yyyy", DTPicker1.Value, DTPicker2.Value)

End Sub

Private Sub Form_LostFocus()
IdEmpl = IdEmpla
End Sub

Private Sub Form_Unload(Cancel As Integer)
IdPac1 = ""
If SQL <> "" Then BD.Close
SQL = ""
End Sub

Private Sub Image3_Click()
Dim TempCad As String
Dim TempCad2 As String
On Error GoTo h

CommonDialog1.ShowOpen
TempCad = CommonDialog1.filename

If InStr(1, TempCad, "\", vbTextCompare) = 0 Then
    If FotoE = "" Then FotoE = "Silueta.jpg"
    Exit Sub
End If

FotoE = Replace(Trim(TxtCedulaEmp.Text) & Trim(TxtApellido.Text) & Trim(TxtNombre.Text) & ".jpg", " ", "")
TempCad2 = FotoEmp & "\" & FotoE
Call FileCopy(TempCad, TempCad2)

If Trim(FotoE) = "" Then Exit Sub
Image3.Picture = LoadPicture(TempCad2)
Image3.Refresh
Cambio = 1
Exit Sub
h:
MsgBox Err.Description
End Sub


Private Sub TxtNombre_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(TxtNombre.Text, 1, 1))
'StrText = Chaa
'
'For i = 2 To Len(TxtNombre.Text)
'    pru = LCase(Mid(TxtNombre.Text, i, 1))
'    If pru Like " " Then
'        t = 1
'        StrText = StrText & " "
'    Else
'        If t = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            t = 0
'        End If
'    End If
'Next i
'
'TxtNombre.Text = StrText
'TxtNombre.SelStart = Len(TxtNombre.Text)
End Sub

Private Sub TxtEdad_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub TxtNroCuenta_Change()
Cambio = 1
End Sub

Private Sub TxtApellido_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(TxtApellido.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(TxtApellido.Text)
'    pru = LCase(Mid(TxtApellido.Text, i, 1))
'    If pru Like " " Then
'        t = 1
'        StrText = StrText & " "
'    Else
'        If t = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            t = 0
'        End If
'    End If
'Next i
'
'TxtApellido.Text = StrText
'TxtApellido.SelStart = Len(TxtApellido.Text)
End Sub

Private Sub TxtCedulaEmp_Change()
Cambio = 1
End Sub

Private Sub TxtCedulaEmp_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub TxtCelular_Change()
Cambio = 1
End Sub

Private Sub TxtCelular_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub TxtRif_Change()
Cambio = 1
End Sub

Private Sub TxtRif_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case vbKeyP
Case vbKeyV
Case vbKeyJ
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub TxtTelefono_Change()
Cambio = 1
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub TxtDireccion_Change()
'Cambio = 1
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(TxtDireccion.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(TxtDireccion.Text)
'    pru = LCase(Mid(TxtDireccion.Text, i, 1))
'    If pru Like " " Then
'        t = 1
'        StrText = StrText & " "
'    Else
'        If t = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            t = 0
'        End If
'    End If
'Next i
'
'TxtDireccion.Text = StrText
'TxtDireccion.SelStart = Len(TxtDireccion.Text)
End Sub

Private Sub TxtCedulaEmp_LostFocus()
Dim IdCed

IdCed = Val(TxtCedulaEmp.Text)

If IdCed = 0 Then Exit Sub

CSql = "SELECT * FROM Empleados Where cedula='" & IdCed & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If RegNew = 1 Then
        MsgBox "La Cedula ingresada ya se encuentra registrada!", vbCritical + vbOKOnly, "Error"
        TxtCedulaEmp.Text = ""
        TxtCedulaEmp.SetFocus
        Exit Sub
    Else
        If Val(Reg_Actual(3)) <> Val(TxtCedulaEmp.Text) Then
            MsgBox "La Cedula ingresada ya se encuentra registrada!", vbCritical + vbOKOnly, "Error"
            TxtCedulaEmp.Text = Reg_Actual(3)
            TxtCedulaEmp.SetFocus
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub TxtPeso_Change()
Cambio = 1
End Sub

Private Sub TxtTalla_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(TxtTalla.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(TxtTalla.Text)
    pru = LCase(Mid(TxtTalla.Text, i, 1))
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

TxtTalla.Text = StrText
TxtTalla.SelStart = Len(TxtTalla.Text)
End Sub

Private Sub TxtSueldo_Change()
Cambio = 1
End Sub
Private Sub TxtEdad_Change()
Cambio = 1
End Sub
Private Sub DTPicker1_Change()
Cambio = 1
TxtEdad.Text = DateDiff("yyyy", DTPicker1.Value, Now)
End Sub
Private Sub CboStatus_Click()
Cambio = 1
End Sub

Private Sub CboCargo_Click()
Cambio = 1
End Sub

Private Sub CboProfesion_Click()
Cambio = 1
End Sub

Private Sub CboGNomina_Click()
Cambio = 1
End Sub

Private Sub TxtSueldo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub CboStatus_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboCargo.SetFocus
        Case vbKeyUp
            TxtTalla.SetFocus
        Case vbKeyLeft
            CboGNomina.SetFocus
    End Select
End If
End Sub

Private Sub CboCargo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboProfesion.SetFocus
        Case vbKeyUp
            DTPicker2.SetFocus
        Case vbKeyRight
            CboProfesion.SetFocus
    End Select
End If
End Sub

Private Sub CboProfesion_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            'TxtSueldo.SetFocus
        Case vbKeyUp
            CboStatus.SetFocus
        Case vbKeyLeft
            CboCargo.SetFocus
    End Select
End If
End Sub

Private Sub CboGNomina_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboStatus.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyLeft
            DTPicker2.SetFocus
        Case vbKeyRight
            CboStatus.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtEdad.SetFocus
        Case vbKeyUp
            TxtTelefono.SetFocus
        Case vbKeyRight
            TxtEdad.SetFocus
        Case vbKeyDown
            DTPicker2.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboGNomina.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyRight
            CboGNomina.SetFocus
        Case vbKeyDown
            CboCargo.SetFocus
    End Select
End If
End Sub

Private Sub TxtNombre_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDireccion.SetFocus
        Case vbKeyUp
            TxtCedulaEmp.SetFocus
        Case vbKeyLeft
            TxtApellido.SetFocus
        Case vbKeyDown
            TxtDireccion.SetFocus
    End Select
End If
End Sub

Private Sub TxtEdad_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtTalla.SetFocus
        Case vbKeyUp
            TxtTelefono.SetFocus
        Case vbKeyLeft
            DTPicker1.SetFocus
        Case vbKeyRight
            TxtTalla.SetFocus
        Case vbKeyDown
            CboGNomina.SetFocus
    End Select
End If
End Sub

Private Sub TxtNroCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            CboProfesion.SetFocus
        Case vbKeyLeft
            TxtSueldo.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub TxtApellido_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNombre.SetFocus
        Case vbKeyUp
            TxtCedulaEmp.SetFocus
        Case vbKeyRight
            TxtNombre.SetFocus
        Case vbKeyDown
            TxtDireccion.SetFocus
    End Select
End If
End Sub

Private Sub TxtCedulaEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtApellido.SetFocus
        Case vbKeyDown
            TxtApellido.SetFocus
    End Select
End If
End Sub

Private Sub TxtCelular_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPicker1.SetFocus
        Case vbKeyUp
            TxtDireccion.SetFocus
        Case vbKeyLeft
            TxtTelefono.SetFocus
        Case vbKeyRight
            BtnConceptosTrabajador.SetFocus
        Case vbKeyDown
            TxtTalla.SetFocus
    End Select
End If
End Sub

Private Sub TxtTelefono_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtCelular.SetFocus
        Case vbKeyUp
            TxtDireccion.SetFocus
        Case vbKeyRight
            TxtCelular.SetFocus
        Case vbKeyDown
            DTPicker1.SetFocus
    End Select
End If
End Sub

Private Sub TxtDireccion_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtTelefono.SetFocus
        Case vbKeyUp
            TxtApellido.SetFocus
        Case vbKeyDown
            TxtTelefono.SetFocus
    End Select
End If
End Sub

Private Sub TxtPeso_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPicker2.SetFocus
        Case vbKeyUp
            TxtCelular.SetFocus
        Case vbKeyLeft
            TxtTalla.SetFocus
        Case vbKeyRight
            BrnListaEmpleados.SetFocus
        Case vbKeyDown
            CboStatus.SetFocus
    End Select
End If
End Sub

Private Sub TxtTalla_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtPeso.SetFocus
        Case vbKeyUp
            TxtCelular.SetFocus
        Case vbKeyLeft
            TxtEdad.SetFocus
        Case vbKeyRight
            TxtPeso.SetFocus
        Case vbKeyDown
            CboStatus.SetFocus
    End Select
End If
End Sub

Private Sub TxtSueldo_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtNroCuenta.SetFocus
        Case vbKeyUp
            CboCargo.SetFocus
        Case vbKeyRight
            TxtNroCuenta.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

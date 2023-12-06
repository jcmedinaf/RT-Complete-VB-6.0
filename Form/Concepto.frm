VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConceptosNomina 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de Nomina"
   ClientHeight    =   7575
   ClientLeft      =   5145
   ClientTop       =   1665
   ClientWidth     =   8325
   Icon            =   "Concepto.frx":0000
   LinkTopic       =   "Form40"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8325
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   240
      TabIndex        =   23
      Top             =   6600
      Width           =   7815
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   6720
         TabIndex        =   24
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
         MICON           =   "Concepto.frx":1002
         PICN            =   "Concepto.frx":101E
         PICH            =   "Concepto.frx":11E7
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
         TabIndex        =   25
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
         MICON           =   "Concepto.frx":141C
         PICN            =   "Concepto.frx":1438
         PICH            =   "Concepto.frx":16C7
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
         TabIndex        =   26
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
         MICON           =   "Concepto.frx":1B08
         PICN            =   "Concepto.frx":1B24
         PICH            =   "Concepto.frx":1CB1
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
         Left            =   5520
         TabIndex        =   27
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
         MICON           =   "Concepto.frx":1EE6
         PICN            =   "Concepto.frx":1F02
         PICH            =   "Concepto.frx":21E4
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
         Left            =   4560
         TabIndex        =   28
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
         MICON           =   "Concepto.frx":2435
         PICN            =   "Concepto.frx":2451
         PICH            =   "Concepto.frx":26E7
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
         Left            =   3960
         TabIndex        =   29
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
         MICON           =   "Concepto.frx":2946
         PICN            =   "Concepto.frx":2962
         PICH            =   "Concepto.frx":2BF7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnElimina 
         Height          =   375
         Left            =   2400
         TabIndex        =   36
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
         MICON           =   "Concepto.frx":2E53
         PICN            =   "Concepto.frx":2E6F
         PICH            =   "Concepto.frx":3013
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
      Caption         =   " Información sobre el Concepto   "
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8055
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   7320
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin ChamaleonButton.ChameleonBtn BtnListaConceptos 
         Height          =   375
         Left            =   3360
         TabIndex        =   30
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Lista"
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
         MICON           =   "Concepto.frx":31B2
         PICN            =   "Concepto.frx":31CE
         PICH            =   "Concepto.frx":3457
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Utilizar"
         Height          =   1815
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Acumulado"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Prestaciones"
            Height          =   375
            Left            =   1800
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Prestamos"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Utilidades"
            Height          =   375
            Left            =   1800
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Vacaciones"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.ComboBox CboGenerar 
         Height          =   315
         ItemData        =   "Concepto.frx":36F6
         Left            =   1200
         List            =   "Concepto.frx":3709
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox CboTipos 
         Height          =   315
         ItemData        =   "Concepto.frx":374A
         Left            =   1200
         List            =   "Concepto.frx":375A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   4335
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   7815
         Begin VB.ComboBox CboCampoAnidado 
            Height          =   315
            ItemData        =   "Concepto.frx":3786
            Left            =   4440
            List            =   "Concepto.frx":3788
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   3840
            Width           =   3255
         End
         Begin VB.CheckBox ChkRestringido 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Restringido"
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   3840
            Width           =   1095
         End
         Begin VB.ComboBox CboCondicionales 
            Height          =   315
            ItemData        =   "Concepto.frx":378A
            Left            =   1200
            List            =   "Concepto.frx":378C
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2280
            Width           =   3615
         End
         Begin VB.ComboBox CboFunciones 
            Height          =   315
            ItemData        =   "Concepto.frx":378E
            Left            =   960
            List            =   "Concepto.frx":3790
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1800
            Width           =   3855
         End
         Begin VB.ComboBox CboConceptos 
            Height          =   315
            ItemData        =   "Concepto.frx":3792
            Left            =   960
            List            =   "Concepto.frx":3794
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   3855
         End
         Begin VB.ComboBox CboCampos 
            Height          =   315
            ItemData        =   "Concepto.frx":3796
            Left            =   1680
            List            =   "Concepto.frx":3798
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1320
            Width           =   3135
         End
         Begin VB.ComboBox CboConstantes 
            Height          =   315
            ItemData        =   "Concepto.frx":379A
            Left            =   960
            List            =   "Concepto.frx":379C
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   840
            Width           =   3855
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   975
            Left            =   120
            TabIndex        =   16
            Top             =   2760
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   1720
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"Concepto.frx":379E
         End
         Begin ChamaleonButton.ChameleonBtn BtnRecargar 
            Height          =   375
            Left            =   6240
            TabIndex        =   31
            ToolTipText     =   "Deshacer Operacion"
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Concepto.frx":3820
            PICN            =   "Concepto.frx":383C
            PICH            =   "Concepto.frx":3B1E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarConceptos 
            Height          =   375
            Left            =   4920
            TabIndex        =   32
            ToolTipText     =   "Agregar Concepto"
            Top             =   360
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
            MICON           =   "Concepto.frx":3D6F
            PICN            =   "Concepto.frx":3D8B
            PICH            =   "Concepto.frx":3F18
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarConstantes 
            Height          =   375
            Left            =   4920
            TabIndex        =   33
            ToolTipText     =   "Agregar Campo"
            Top             =   840
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
            MICON           =   "Concepto.frx":414D
            PICN            =   "Concepto.frx":4169
            PICH            =   "Concepto.frx":42F6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnCampoTrabajador 
            Height          =   375
            Left            =   4920
            TabIndex        =   34
            ToolTipText     =   "Agregar Constante"
            Top             =   1320
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
            MICON           =   "Concepto.frx":452B
            PICN            =   "Concepto.frx":4547
            PICH            =   "Concepto.frx":46D4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnNuevaConst 
            Height          =   375
            Left            =   6240
            TabIndex        =   35
            ToolTipText     =   "Crear Nuevo Constante"
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Crear Const."
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
            MICON           =   "Concepto.frx":4909
            PICN            =   "Concepto.frx":4925
            PICH            =   "Concepto.frx":4AB2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
            Height          =   375
            Left            =   4920
            TabIndex        =   39
            ToolTipText     =   "Agregar Función"
            Top             =   1800
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
            MICON           =   "Concepto.frx":4CE7
            PICN            =   "Concepto.frx":4D03
            PICH            =   "Concepto.frx":4E90
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregarCondicion 
            Height          =   375
            Left            =   4920
            TabIndex        =   41
            ToolTipText     =   "Agregar Condicional"
            Top             =   2280
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
            MICON           =   "Concepto.frx":50C5
            PICN            =   "Concepto.frx":50E1
            PICH            =   "Concepto.frx":526E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   6120
            X2              =   6120
            Y1              =   120
            Y2              =   2760
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Campo Anidado:"
            Height          =   195
            Left            =   3240
            TabIndex        =   45
            Top             =   3900
            Width           =   1170
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condicionales:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   2340
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funciones:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1860
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Campo del trabajador:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1380
            Width           =   1545
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Constante:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   765
         End
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generar:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmConceptosNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDCOncep As Recordset
Dim BDCOnc As Recordset
Dim BDCOns As Recordset
Dim BDCOcamp As Recordset
Dim BDCO As Recordset
Dim RegNew
Dim Cambio
Dim IdC

Private Sub BtnAgregar_Click()
Blanqueo
Dim RsConcepto As New ADODB.Recordset
CSql = "Select MAX(IdConcepto) as TConcepto From Concepto"
Set RsConcepto = CrearRS(CSql)

If Not IsNull(RsConcepto.Fields("TConcepto").Value) Then
    Label6.Caption = Format(RsConcepto.Fields("TConcepto").Value + 1, "0000#")
    IdC = RsConcepto.Fields("TConcepto").Value + 1
Else
    Label6.Caption = Format(1, "0000")
    IdC = 1
End If
ChkRestringido.Value = 0
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
BtnAgregar.Enabled = False
BtnElimina.Enabled = False
BtnSiguiente.Enabled = False
BtnAnterior.Enabled = False
End Sub

Private Sub BtnAgregarConceptos_Click()
If CboConceptos.ListIndex = -1 Then Exit Sub
Cad = "Concepto(" & Mid(CboConceptos.List(CboConceptos.ListIndex), 1, 10) & ";" & CboConceptos.ItemData(CboConceptos.ListIndex) & ")"
RichTextBox1.Text = RichTextBox1.Text & Cad
End Sub

Private Sub BtnAgregarCondicion_Click()
If CboCondicionales.ListIndex = -1 Then Exit Sub
Cad = ListCond(CboCondicionales.ListIndex, 0)
RichTextBox1.Text = RichTextBox1.Text & Cad
End Sub

Private Sub BtnAgregarConstantes_Click()
If CboConstantes.ListIndex = -1 Then Exit Sub
Cad = "Constante(" & Mid(CboConstantes.List(CboConstantes.ListIndex), 1, 10) & ";" & CboConstantes.ItemData(CboConstantes.ListIndex) & ")"
RichTextBox1.Text = RichTextBox1.Text & Cad
End Sub

Private Sub BtnAnterior_Click()
If BDCO.BOF And BDCO.EOF Then Exit Sub
Call verify1
BDCO.MovePrevious
If BDCO.BOF Then BDCO.MoveLast

Call CargaConc
End Sub

Private Sub BtnCampoTrabajador_Click()
If CboCampos.ListIndex = -1 Then Exit Sub
Cad = "Campo(" & Mid(CboCampos.List(CboCampos.ListIndex), 1, 10) & ";" & CboCampos.ItemData(CboCampos.ListIndex) & ")"
RichTextBox1.Text = RichTextBox1.Text & Cad
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
Form_Load
End Sub

Private Sub BtnElimina_Click()
On Error GoTo MostrarError
Dim resp As Byte
resp = MsgBox("Desea borrar el concepto '" & Text2 & "' ?", vbQuestion + vbYesNo, "Confirmar")

If resp = 7 Then Exit Sub

CSql = "UPDATE Concepto SET Activo=0 WHERE IdConcepto=" & Val(Label6.Caption) & " AND Activo=1"
Set RsTemp = CrearRS(CSql)
BtnDesHacer_Click
MsgBox "El concepto ha sido borrado!", vbInformation + vbOKOnly, "Operación Exitosa!"

Exit Sub
MostrarError:
MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error GoTo MostrarError
Dim NuevoId

If CboCampoAnidado.ListIndex = -1 Then CboCampoAnidado.ListIndex = 0

Select Case RegNew

Case Is = 1
    If Cambio = 1 Then
    
        CSql = "SELECT MAX(IdConcepto)+1 As NuevoId FROM Concepto"
        Set BDCOncep = CrearRS(CSql)
        
        If IsNull(BDCOncep.Fields("NuevoId").Value) Then
            NuevoId = "1"
            Else
            NuevoId = BDCOncep.Fields("NuevoId").Value
        End If
        
        CSql = "insert into Concepto(IdConcepto, IdCampoAnidado, Formula,Descripcion,tipo,Genera,Acumulado, " & _
            "Prestaciones,Prestamo,Utilidades,Vacaciones,Activo,Restringido) values(" & NuevoId & "," & CboCampoAnidado.ItemData(CboCampoAnidado.ListIndex) & ",'" & RichTextBox1.Text & _
            "','" & Text2.Text & "'," & CboTipos.ListIndex & "," & CboGenerar.ListIndex & "," & _
            Check1.Value & "," & Check2.Value & "," & Check3.Value & "," & Check4.Value & "," & Check5.Value & ",1,'" & ChkRestringido.Value & "')"
            
        ' Los campos AplicarCuando, Inhabilitado, Transferido, Revisado no se estan agregando por ahora la razon=NO SE

        Set BDCOncep = CrearRS(CSql)
        
        Msg = "El concepto se ha agregado satisfactoriamente"
        
    Else
        Msg = "no hay cambios que agregar"
    End If
Case Is = 0
    If Cambio = 1 Then
        CSql = "update concepto set Formula = '" & RichTextBox1.Text & "', Descripcion = '" & _
            Text2.Text & "', IdCampoAnidado=" & CboCampoAnidado.ItemData(CboCampoAnidado.ListIndex) & ",tipo =" & CboTipos.ListIndex & ", Genera = " & CboGenerar.ListIndex & _
            ", Acumulado =" & Check1.Value & ", Prestaciones =" & Check2.Value & ", Prestamo =" & _
            Check3.Value & ", Utilidades =" & Check4.Value & ", Vacaciones =" & Check5.Value & _
            ",Restringido='" & ChkRestringido.Value & "' where IdConcepto = " & Val(Label6.Caption)
            
        ' Los campos AplicarCuando, Inhabilitado, Transferido, Revisado no se estan agregando por ahora la razon=NO SE

        Set BDCOncep = CrearRS(CSql)
        Msg = "Guardado Satisfactoriamente"
    Else
        Msg = "No hay cambios que guardar"
    End If
    
End Select
Cambio = 0
MsgBox Msg, vbOKOnly + vbInformation, "Guardar"
Call Form_Load

Exit Sub
MostrarError:
MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source
End Sub

Private Sub BtnListaConceptos_Click()
Tipo1 = "Conceptos"
FrmListadoConceptos.Show
End Sub

Private Sub BtnNuevaConst_Click()
    FrmAgregaConstNomina.Show
End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub BtnRecargar_Click()
Dim CodConp As String
Dim CadConcepto As String
Dim CadCampo As String
Dim CadSSO As String
Dim CadConstantes As String
Dim Band As Boolean
Dim resultado As Double

Msg = "Indique el ID de un trabajador para evaluar el concepto"
IdEmpl = InputBox(Msg, "ID Trabajador", 2)
If Trim(IdEmpl) = "" Then Exit Sub
    
    
    CodConp = Trim(RichTextBox1.Text)
    
    CadConcepto = Validar_Concepto(CodConp)
    CadCampo = Validar_Campo(CadConcepto)
            
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
            
    CadFunciones = Validar_Funcion(CadCampo, TMes, Format(Now, "dd/MM/yyyy"))
    CadConstantes = Validar_Constante(CadFunciones)
    CadConstantes = Replace(CadConstantes, ",", ".")
    CadSSO = CadConstantes
    CadConstantes = Validar_SSO(CadConstantes)
            
    On Error GoTo errty
    resultado = ScriptControl1.Eval(CadConstantes)

    If CadSSO <> CadConstantes Then
        resultado = Calcular_SSO(CadConstantes, resultado, IdEmpl)
    End If
    
    MsgBox CadConstantes & " = " & resultado
            
Exit Sub

errty:
'If Err.Number = 1002 Then
    Msg = "Hay un error en algun concepto o valor del trabajador" & Chr(13) & "Revise e intente de nuevo" & Chr(13) & Chr(13) & "     " & FormulA
    MsgBox Msg, vbOKOnly + vbCritical, "Error formula"
'End If
End Sub

Private Sub BtnSiguiente_Click()
If BDCO.BOF And BDCO.EOF Then Exit Sub
Call verify1
BDCO.MoveNext
If BDCO.EOF Then BDCO.MoveFirst

Call CargaConc
End Sub


Private Sub CboCampoAnidado_Click()
Cambio = 1
End Sub

Private Sub CboFunciones_Change()
Cambio = 1
End Sub

Private Sub ChameleonBtn1_Click()
If CboFunciones.ListIndex = -1 Then Exit Sub
Cad = ListFunc(CboFunciones.ListIndex, 0)
RichTextBox1.Text = RichTextBox1.Text & Cad
End Sub

Private Sub Check1_Click()
Cambio = 1
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeySpace
            Check2.SetFocus
        Case vbKeyReturn
            Check2.SetFocus
    End Select
End If
End Sub

Private Sub Check2_Click()
Cambio = 1
End Sub

Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeySpace
            Check3.SetFocus
        Case vbKeyReturn
            Check3.SetFocus
    End Select
End If
End Sub

Private Sub Check3_Click()
Cambio = 1
End Sub

Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeySpace
            Check4.SetFocus
        Case vbKeyReturn
            Check4.SetFocus
    End Select
End If
End Sub

Private Sub Check4_Click()
Cambio = 1
End Sub

Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeySpace
            Check5.SetFocus
        Case vbKeyReturn
            Check5.SetFocus
    End Select
End If
End Sub

Private Sub Check5_Click()
Cambio = 1
End Sub

Private Sub Check5_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeySpace
            CboConceptos.SetFocus
        Case vbKeyReturn
            CboConceptos.SetFocus
    End Select
End If
End Sub

Private Sub CboTipos_Click()
Cambio = 1
End Sub

Private Sub CboTipos_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboGenerar.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyRight
            CboGenerar.SetFocus
        Case vbKeyDown
            Check1.SetFocus
    End Select
End If
End Sub

Private Sub CboGenerar_Click()
Cambio = 1
End Sub

Private Sub CboGenerar_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Check1.SetFocus
        Case vbKeyLeft
            CboTipos.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyDown
            Check1.SetFocus
    End Select
End If
End Sub

Private Sub CboConceptos_Click()
Cambio = 1
End Sub

Private Sub CboConceptos_DropDown()
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregarConceptos.SetFocus
        Case vbKeyUp
            Check1.SetFocus
        Case vbKeyRight
            BtnAgregarConceptos.SetFocus
        Case vbKeyDown
            CboConstantes.SetFocus
    End Select
End If
End Sub

Private Sub CboCampos_Click()
Cambio = 1
End Sub

Private Sub CboCampos_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnCampoTrabajador.SetFocus
        Case vbKeyLeft
            BtnAgregarConceptos.SetFocus
        Case vbKeyUp
            Check1.SetFocus
        Case vbKeyRight
            BtnCampoTrabajador.SetFocus
        Case vbKeyDown
            RichTextBox1.SetFocus
    End Select
End If
End Sub

Private Sub CboConstantes_Click()
Cambio = 1
End Sub

Sub Blanqueo()

RichTextBox1.Text = ""
Text2.Text = ""
Label6.Caption = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
CboTipos.ListIndex = -1
CboGenerar.ListIndex = -1
CboConceptos.ListIndex = -1
CboCampos.ListIndex = -1
CboConstantes.ListIndex = -1
Cambio = 0
RegNew = 1
End Sub

Private Sub CboConstantes_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregarConstantes.SetFocus
        Case vbKeyUp
            CboConceptos.SetFocus
        Case vbKeyRight
            BtnAgregarConstantes.SetFocus
        Case vbKeyDown
            RichTextBox1.SetFocus
    End Select
End If
End Sub

Private Sub ChkRestringido_Click()
Cambio = 1
End Sub

Private Sub Form_Load()
Centrar Me

CSql = "select * from concepto where activo=1"
Set BDCO = CrearRS(CSql)

CSql = "select * from concepto where activo=1"
Set BDCOnc = CrearRS(CSql)
CboConceptos.Clear

If Not (BDCOnc.EOF And BDCOnc.BOF) Then
    BDCOnc.MoveFirst
    Do While Not BDCOnc.EOF
        If Not IsNull(BDCOnc.Fields("DESCRIPCION")) Then
            CboConceptos.AddItem BDCOnc.Fields("DESCRIPCION")
            CboConceptos.ItemData(CboConceptos.NewIndex) = BDCOnc.Fields("IDCONCEPTO")
            BDCOnc.MoveNext
        End If
    Loop
End If

BDCOnc.Close

CboCampos.Clear
CboCampoAnidado.Clear

CboCampoAnidado.AddItem "Ninguno"
CboCampoAnidado.ItemData(CboCampoAnidado.NewIndex) = 0

CSql = "select * from camposdenomina where activo=1"
Set BDCOcamp = CrearRS(CSql)
If Not BDCOcamp.EOF Then
    BDCOcamp.MoveFirst
    Do While Not BDCOcamp.EOF
        If Not IsNull(BDCOcamp.Fields("campo")) Then
            
            CboCampos.AddItem BDCOcamp.Fields("campo")
            CboCampos.ItemData(CboCampos.NewIndex) = BDCOcamp.Fields("IdCampoNomina")
            
            CboCampoAnidado.AddItem BDCOcamp.Fields("campo")
            CboCampoAnidado.ItemData(CboCampoAnidado.NewIndex) = BDCOcamp.Fields("IdCampoNomina")
            BDCOcamp.MoveNext
        End If
    Loop
End If
BDCOcamp.Close

CSql = "select * from constantesdenomina where activo=1"
Set BDCOns = CrearRS(CSql)

CboConstantes.Clear

If BDCOns.RecordCount = 0 Then Exit Sub

BDCOns.MoveFirst

Do While Not BDCOns.EOF
    If Not IsNull(BDCOns.Fields("Descripcion")) Then
        CboConstantes.AddItem BDCOns.Fields("Descripcion")
        CboConstantes.ItemData(CboConstantes.NewIndex) = BDCOns.Fields("IDConstante")
        BDCOns.MoveNext
    End If
Loop
BDCOns.Close

Cargar_Lista_De_Funciones
For i = 0 To 200
    If IsNull(ListFunc(i, 0)) Then Exit For
    If Trim(ListFunc(i, 0)) = "" Then Exit For
    CboFunciones.AddItem ListFunc(i, 0) & " / " & ListFunc(i, 1)
Next i

Cargar_Lista_De_Condicionales
For i = 0 To 200
    If IsNull(ListCond(i, 0)) Then Exit For
    If Trim(ListCond(i, 0)) = "" Then Exit For
    CboCondicionales.AddItem ListCond(i, 0) & " / " & ListCond(i, 1)
Next i

For i = 0 To 4
Next
Call CargaConc
Cambio = 0
RegNew = 0

End Sub

Sub CargaConc()
If Not BDCO.EOF Then
    If IsNull(BDCO.Fields(3)) Then RichTextBox1.Text = "" Else RichTextBox1.Text = BDCO.Fields(3)
    If IsNull(BDCO.Fields(1)) Then Text2.Text = "" Else Text2.Text = BDCO.Fields(1)
    
    If IsNull(BDCO.Fields("Acumulado")) Then Check1.Value = 0 Else Check1.Value = 1
    If IsNull(BDCO.Fields("Prestaciones")) Then Check2.Value = 0 Else Check2.Value = 1
    If IsNull(BDCO.Fields("Prestamo")) Then Check3.Value = 0 Else Check3.Value = 1
    If IsNull(BDCO.Fields("Utilidades")) Then Check4.Value = 0 Else Check4.Value = 1
    If IsNull(BDCO.Fields("Vacaciones")) Then Check5.Value = 0 Else Check5.Value = 1
    If Val(BDCO.Fields("Restringido")) = 0 Then ChkRestringido.Value = 0 Else ChkRestringido.Value = 1
    
    Concep = Format(BDCO.Fields("Idconcepto"), "000#")
    Label6.Caption = Concep
    IDcon = Idconcepto
    
    For T = 0 To CboTipos.ListCount - 1
        If CboTipos.ItemData(T) = BDCO.Fields("tipo") Then
            CboTipos.ListIndex = T
            Exit For
        End If
    Next T
                    
    For T = 0 To CboGenerar.ListCount - 1
          If CboGenerar.ItemData(T) = BDCO.Fields("Genera") Then
          CboGenerar.ListIndex = T
          Exit For
          End If
    Next T
    
    For T = 0 To CboCampoAnidado.ListCount - 1
          If CboCampoAnidado.ItemData(T) = BDCO.Fields("IdCampoAnidado") Then
          CboCampoAnidado.ListIndex = T
          Exit For
          Else
            CboCampoAnidado.ListIndex = 0
          End If
    Next T
    
    BtnAgregar.Enabled = True
    BtnElimina.Enabled = True
    BtnSiguiente.Enabled = True
    BtnAnterior.Enabled = True
Else
    BtnAgregar.Enabled = True
    BtnElimina.Enabled = False
    BtnSiguiente.Enabled = False
    BtnAnterior.Enabled = False
End If

Cambio = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
BDCO.Close
End Sub

Private Sub RichTextBox1_Change()
Cambio = 1
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbKeyControl Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            CboConstantes.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub
  

Private Sub Text2_Change()
Cambio = 1
End Sub

Sub verify1()
If Cambio = 1 Then
    Msg = "Se han hecho cambios en la data desea guardar?"
    g = MsgBox(Msg, vbYesNo + vbInformation, "Desea Guardar")
    If g = 6 Then Call BtnGuardarActualizar_Click
End If
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            CboTipos.SetFocus
        Case vbKeyUp
            BtnListaConceptos.SetFocus
        Case vbKeyDown
            CboTipos.SetFocus
    End Select
End If
End Sub

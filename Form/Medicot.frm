VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmMedicoTratante 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Médicos Tratantes"
   ClientHeight    =   4530
   ClientLeft      =   4155
   ClientTop       =   2970
   ClientWidth     =   6450
   Icon            =   "Medicot.frx":0000
   LinkTopic       =   "Form55"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   6015
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   4920
            TabIndex        =   9
            ToolTipText     =   "Cerrar Tablas de Pacientes"
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
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
            MICON           =   "Medicot.frx":1002
            PICN            =   "Medicot.frx":101E
            PICH            =   "Medicot.frx":11E7
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
            Height          =   495
            Left            =   1200
            TabIndex        =   10
            ToolTipText     =   "Guardar / Actualizar Pacientes"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Medicot.frx":141C
            PICN            =   "Medicot.frx":1438
            PICH            =   "Medicot.frx":16C7
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
            Height          =   495
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Agregar Pacientes"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "Medicot.frx":1B08
            PICN            =   "Medicot.frx":1B24
            PICH            =   "Medicot.frx":1CB1
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
            Height          =   495
            Left            =   3720
            TabIndex        =   12
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Medicot.frx":1EE6
            PICN            =   "Medicot.frx":1F02
            PICH            =   "Medicot.frx":21E4
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
            Height          =   495
            Left            =   2400
            TabIndex        =   13
            ToolTipText     =   "Eliminar Usuario"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Medicot.frx":2435
            PICN            =   "Medicot.frx":2451
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
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6015
         Begin VB.ComboBox CboTitulo 
            Height          =   315
            ItemData        =   "Medicot.frx":2809
            Left            =   3720
            List            =   "Medicot.frx":2819
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox TxtCedula 
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtEspecialidad 
            DataField       =   "esp"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   2640
            Width           =   4575
         End
         Begin VB.TextBox TxtNombre 
            DataField       =   "Medico"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox TxtApellido 
            DataField       =   "Medico"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   1200
            Width           =   4575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Título:"
            Height          =   195
            Left            =   3720
            TabIndex        =   16
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "No. Cédula:"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Especialidad:"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   960
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "FrmMedicoTratante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAgregar_Click()
Blanqueo
TxtCedula.SetFocus
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
Dim RsGuardarMedicoT As New ADODB.Recordset
CSql = "Select * From Medicos_t"
Set RsGuardarMedicoT = CrearRS(CSql)

C = RsGuardarMedicoT.RecordCount + 1

RsGuardarMedicoT.AddNew
RsGuardarMedicoT.Fields(0).Value = C
RsGuardarMedicoT.Fields(1).Value = IdUser
RsGuardarMedicoT.Fields(2).Value = CboTitulo.Text
RsGuardarMedicoT.Fields(3).Value = TxtNombre.Text
RsGuardarMedicoT.Fields(4).Value = TxtApellido.Text
RsGuardarMedicoT.Fields(5).Value = TxtCedula.Text
RsGuardarMedicoT.Fields(6).Value = txtEspecialidad.Text
RsGuardarMedicoT.Update

RsGuardarMedicoT.Close
Blanqueo
End Sub

Sub Blanqueo()
TxtApellido.Text = ""
TxtNombre.Text = ""
txtEspecialidad.Text = ""
TxtCedula.Text = ""
CboTitulo.Text = ""
End Sub

Private Sub Form_Load()
Centrar Me
End Sub

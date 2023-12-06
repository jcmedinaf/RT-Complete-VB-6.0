VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmMedicoRemitente 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Médicos Remitentes"
   ClientHeight    =   4410
   ClientLeft      =   4155
   ClientTop       =   2970
   ClientWidth     =   6495
   Icon            =   "MedicoR.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6495
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   6015
         Begin VB.TextBox TxtTelefono 
            DataField       =   "Telefono"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox TxtCentroClinico 
            DataField       =   "Centro"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   4215
         End
         Begin VB.TextBox TxtNombre 
            DataField       =   "Medico"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   4215
         End
         Begin VB.TextBox TxtApellido 
            DataField       =   "Medico"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            DataField       =   "Id"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5280
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Telefono"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Centro Clinico"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   6015
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   4920
            TabIndex        =   2
            ToolTipText     =   "Cerrar"
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
            MICON           =   "MedicoR.frx":1002
            PICN            =   "MedicoR.frx":101E
            PICH            =   "MedicoR.frx":11E7
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
            TabIndex        =   3
            ToolTipText     =   "Guardar / Actualizar "
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
            MICON           =   "MedicoR.frx":141C
            PICN            =   "MedicoR.frx":1438
            PICH            =   "MedicoR.frx":16C7
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
            TabIndex        =   4
            ToolTipText     =   "Agregar"
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
            MICON           =   "MedicoR.frx":1B08
            PICN            =   "MedicoR.frx":1B24
            PICH            =   "MedicoR.frx":1CB1
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
            TabIndex        =   5
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
            MICON           =   "MedicoR.frx":1EE6
            PICN            =   "MedicoR.frx":1F02
            PICH            =   "MedicoR.frx":21E4
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
            TabIndex        =   16
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
            MICON           =   "MedicoR.frx":2435
            PICN            =   "MedicoR.frx":2451
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
Attribute VB_Name = "FrmMedicoRemitente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
Dim RsGuardar As New ADODB.Recordset
CSql = "Select * From Medicos_r"
Set RsGuardar = CrearRS(CSql)
C = RsGuardar.RecordCount + 1
RsGuardar.AddNew
RsGuardar.Fields("IdMedicosr").Value = C
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Fields("Nombre").Value = Trim(TxtNombre.Text)
RsGuardar.Fields("Apellido").Value = Trim(TxtApellido.Text)
RsGuardar.Fields("Clinica").Value = Trim(TxtCentroClinico.Text)
RsGuardar.Fields("Telefono").Value = Trim(TxtTelefono.Text)
RsGuardar.Update
Blanqueo
End Sub

Private Sub Form_Load()
Centrar Me
End Sub
Sub Blanqueo()
TxtNombre.Text = ""
TxtApellido.Text = ""
TxtCentroClinico.Text = ""
TxtTelefono.Text = ""
End Sub

VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmClaveSupervisor 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave Supervisora"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   Icon            =   "FrmClaveSupervisor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox TxtClave 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin ChamaleonButton.ChameleonBtn BtnSalir 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Salir del Sistema"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancelar"
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
         MICON           =   "FrmClaveSupervisor.frx":1002
         PICN            =   "FrmClaveSupervisor.frx":101E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAceptar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Entrar al Sistema"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Aceptar"
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
         MICON           =   "FrmClaveSupervisor.frx":11BD
         PICN            =   "FrmClaveSupervisor.frx":11D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   450
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   930
         Width           =   450
      End
   End
End
Attribute VB_Name = "FrmClaveSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Acceso As Boolean

Private Sub BtnAceptar_Click()

If Trim(TxtUsuario.Text) = "AdminOA" And Trim(TxtClave.Text) = "458921957JArr" Then
    'FrmNuevoPaciente.Eliminar
    MsgBox "Clave aceptada!", vbInformation + vbOKOnly, "Información"
    Acceso = True
    Unload Me
Else
    MsgBox "Usuario o Clave Erronea", vbCritical + vbOKOnly, "Error"
    Acceso = False
End If

End Sub

Private Sub BtnSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
Acceso = False
End Sub

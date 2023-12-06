VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCedulaPaciente 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cedula del Paciente"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "FrmCedulaPaciente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox TxtCedula 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Text            =   "123456789"
         ToolTipText     =   "Ingrese el número de cedula de identidad del paciente."
         Top             =   840
         Width           =   3735
      End
      Begin ChamaleonButton.ChameleonBtn BtnSalir 
         Height          =   495
         Left            =   4080
         TabIndex        =   1
         ToolTipText     =   "Salir del Sistema"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         MICON           =   "FrmCedulaPaciente.frx":1002
         PICN            =   "FrmCedulaPaciente.frx":101E
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
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         ToolTipText     =   "Entrar al Sistema"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         MICON           =   "FrmCedulaPaciente.frx":11BD
         PICN            =   "FrmCedulaPaciente.frx":11D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Indique la cédula del paciente que va a generarle el presupuesto"
         Height          =   555
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "FrmCedulaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ced

Private Sub BtnAceptar_Click()
If Trim(TxtCedula.Text) = "" Then
    MsgBox "Debe de ingresar la cedula de identidad del paciente!!", vbExclamation + vbOKOnly, "Error"
    Ced = "-1"
    Exit Sub
Else
    Ced = Trim(TxtCedula.Text)
    Unload Me
End If
End Sub

Private Sub BtnSalir_Click()
Ced = "-1"
Unload Me
End Sub

Private Sub Form_Load()
Ced = "-1"
Centrar Me
End Sub

Private Sub TxtCedula_Click()
If TxtCedula.Text = "123456789" Then TxtCedula.Text = ""
End Sub

Private Sub TxtCedula_GotFocus()
If TxtCedula.Text = "123456789" Then TxtCedula.Text = ""
End Sub

Private Sub TxtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnAceptar.SetFocus
Else
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        MsgBox "El caracter digitado no es válido.", vbExclamation, "Error"
        KeyAscii = 0
    End If
End If
End Sub

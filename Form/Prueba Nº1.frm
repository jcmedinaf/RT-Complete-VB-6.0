VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "INSERTEVECA"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Acceso al Sistema Inserteveca"
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton Command3 
         Caption         =   "Ayuda"
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   2160
         TabIndex        =   6
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Clave de Usuario"
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Usuario"
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
Beep
End
End Sub

Private Sub Command2_Click()
Beep
End
End Sub



Private Sub Form_Click()
Beep
End
End Sub

'Private Sub Form_Load()
'Data1.Visible = False
'Data1.DatabaseName = App.Path + "\INSERTEVECA.mdb"
'End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
SendKeys "{home}+{end}"
Text1.Text = LCase(Text1.Text)
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Data1.RecordSource = ("select * from pass where password='" & Text2.Text & "' and user = '" & Text1.Text & "'")
Data1.Refresh
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Acceso Denegado", vbCritical, "Password"
Text1.SetFocus
SendKeys "{home}+{end}"
Else
MsgBox "Acceso Aprobado", vbInformation, "Password"
Unload password
Load main
main.Show
'Load frmTip
'frmTip.Show 1
End If
End If
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4860
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   120
      Top             =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando  Sistema al "
      Height          =   195
      Left            =   3000
      TabIndex        =   1
      Top             =   4560
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   4650
      Left            =   120
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7605
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If ProgressBar1.Value = 100 Then
    Call Animar(FrmSplash, 500, AW_BLEND Or AW_HIDE)
    FrmLogin.Show
    Timer1.Enabled = False
Else
    ProgressBar1.Value = Val(ProgressBar1.Value) + Val(1)
    Label1.Caption = "Cargando  Sistema al " & ProgressBar1.Value & " %"
End If
End Sub

VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReconexion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restableciendo conexión..."
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "FrmReconexion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonButton.ChameleonBtn BtnCerrar 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "Cerrar"
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cerrar Aplicación"
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
      MICON           =   "FrmReconexion.frx":1002
      PICN            =   "FrmReconexion.frx":101E
      PICH            =   "FrmReconexion.frx":11E7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn BtnIntentar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Deshacer Operacion"
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Intentar de Nuevo"
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
      MICON           =   "FrmReconexion.frx":141C
      PICN            =   "FrmReconexion.frx":1438
      PICH            =   "FrmReconexion.frx":171A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmReconexion.frx":196B
      Height          =   675
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00EAEFEF&
      Caption         =   "En 10 se intentará restablecer la conexión al servidor"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00EAEFEF&
      Caption         =   "Hora de inicio de a Falla:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1755
   End
End
Attribute VB_Name = "FrmReconexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim contt As Byte

Private Sub BtnCerrar_Click()
Unload FrmPrincipal
Unload Me
End Sub

Private Sub BtnIntentar_Click()
On Error GoTo mostrar
    Timer1.Enabled = False
    If Cnn.State = adStateOpen Then Cnn.Close
    Cnn.Open
    
    BtnIntentar.Enabled = False
    DoEvents
         
    If Cnn.State = adStateOpen Then
        MsgBox "Conexión establecida!", vbInformation + vbOKOnly, "RT Complete."
        Unload FrmReconexion
        If FrmPrincipal.Tag <> "1" Then Main
    End If

Exit Sub

mostrar:
    MsgBox "No se ha podido establecer la conexión!", vbExclamation + vbOKOnly, "Información!"
    Timer1.Enabled = True
    BtnIntentar.Enabled = True
End Sub

Private Sub Form_Load()
contt = 10
If Cnn.State = adStateOpen Then Cnn.Close
Label1.Caption = "Hora de inicio de a Falla: " & Format(Now, "hh:mm:ss AMPM")
End Sub

Private Sub Timer1_Timer()
On Error GoTo mostrar
contt = contt - 1
Label2.Caption = "En " & contt & " seg, se intentará reestablecer la conexión al servidor"
If contt = 0 Then
    contt = 10
    Label2.Caption = "Estableciendo la conexión..."
    BtnIntentar.Enabled = False
    DoEvents
    If Cnn.State = adStateOpen Then Cnn.Close
    Cnn.Open
    
     
    If Cnn.State = adStateOpen Then
        MsgBox "Conexión establecida!", vbInformation + vbOKOnly, "RT Complete."
        Unload FrmReconexion
        If FrmPrincipal.Tag <> "1" Then Main
    End If
    
End If

Exit Sub

mostrar:
    BtnIntentar.Enabled = True
    
End Sub

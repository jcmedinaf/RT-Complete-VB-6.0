VERSION 5.00
Begin VB.Form Form50 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Administratido OncoAmerica (Recepción)"
   ClientHeight    =   6345
   ClientLeft      =   6615
   ClientTop       =   3555
   ClientWidth     =   9165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Recepcion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Recepcion.frx":628A
   ScaleHeight     =   6345
   ScaleMode       =   0  'User
   ScaleWidth      =   9165
   Begin VB.Image Image1 
      Height          =   6330
      Left            =   0
      Picture         =   "Recepcion.frx":6E30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9285
   End
   Begin VB.Menu Pac 
      Caption         =   "&Pacientes"
      Index           =   0
      NegotiatePosition=   2  'Middle
      WindowList      =   -1  'True
      Begin VB.Menu NewPac 
         Caption         =   "A&gregar Pacientes"
         Shortcut        =   ^G
      End
      Begin VB.Menu Reg 
         Caption         =   "R&egistro Historico"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu E 
      Caption         =   "&Estatus"
      Index           =   1
   End
   Begin VB.Menu Inicio 
      Caption         =   "&Area Medica"
      Begin VB.Menu Mec 
         Caption         =   "M&edicos de Turno"
         Shortcut        =   ^E
      End
      Begin VB.Menu Inte 
         Caption         =   "&Residente"
         Shortcut        =   ^R
      End
      Begin VB.Menu Tec 
         Caption         =   "Tecnica "
         Shortcut        =   ^T
      End
      Begin VB.Menu Nut 
         Caption         =   "&Nutrición"
         Shortcut        =   ^N
      End
      Begin VB.Menu Psi 
         Caption         =   "P&sicologico"
         Begin VB.Menu Niños 
            Caption         =   "Consulta &Niños"
            Shortcut        =   ^M
         End
         Begin VB.Menu Adulto 
            Caption         =   "Consu&lta Adulto"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu Rad 
         Caption         =   "Radi&oterapeuta"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Ad 
      Caption         =   "A&dministracion "
      Begin VB.Menu f 
         Caption         =   "Facturacion "
      End
      Begin VB.Menu pre 
         Caption         =   "Pres&upuesto"
         Shortcut        =   ^U
      End
      Begin VB.Menu In 
         Caption         =   "Inventarios"
         Begin VB.Menu h 
            Caption         =   "Haber"
         End
         Begin VB.Menu Com 
            Caption         =   "Compras"
         End
      End
      Begin VB.Menu N 
         Caption         =   "Nota de credito"
      End
      Begin VB.Menu Agre 
         Caption         =   "Agregar &Usuario"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu exit 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Form50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adulto_Click()
Form9.Show 1
End Sub

Private Sub Agre_Click()
Form6.Show 1
End Sub


Private Sub E_Click(Index As Integer)
'Form4.Show 1
End Sub

Private Sub Est_Click()

End Sub

Private Sub Exit_Click()
If MsgBox("¿Desea terminar la aplicación?", _
vbQuestion + vbYesNo, "Pregunta") = vbYes Then
End
Else
Cancel = True
End If
End Sub


Private Sub f_Click()
Form27.Show 1
End Sub

Private Sub Form_Load()


    Select Case UCase(T_U)
    Case Is = "1"
    Inicio.Enabled = False
    Ad.Enabled = False
    
    Case Is = "0"
    Ad.Enabled = True
    
    Case Is = "2" 'Radioterapeuta
    inte.Enabled = False
    NewPac.Enabled = False
    Nut.Enabled = False
    Psi.Enabled = False
    Ad.Enabled = False
    Tec.Enabled = False
    
    Case Is = "3" 'Internista
    Rad.Enabled = False
    NewPac.Enabled = False
    Nut.Enabled = False
    Psi.Enabled = False
    Ad.Enabled = False
    Tec.Enabled = False
    
    Case Is = "4" 'Psicologia
    inte.Enabled = False
    NewPac.Enabled = False
    Nut.Enabled = False
    Ad.Enabled = False
    Rad.Enabled = False
    Tec.Enabled = False
    
    Case Is = "5" 'Nutricion
    inte.Enabled = False
    NewPac.Enabled = False
    Rad.Enabled = False
    Psi.Enabled = False
    Ad.Enabled = False
    Tec.Enabled = False
    Case Is = "6" 'Administracion
    Agre.Enabled = False
    inte.Enabled = False
    Rad.Enabled = False
    Psi.Enabled = False
    Nut.Enabled = False
    Tec.Enabled = False
    
    Case Is = "7" 'Tecnica
    inte.Enabled = False
    NewPac.Enabled = False
    Rad.Enabled = False
    Psi.Enabled = False
    Ad.Enabled = False
    Nut.Enabled = False
    
    Case Is = "8" 'Radioterapeuta
    NewPac.Enabled = False
    Nut.Enabled = False
    Psi.Enabled = False
    Ad.Enabled = False
    Tec.Enabled = False
    End Select
    Caption = "Sistema Administrativo OncoAmerica        " & Usuario

End Sub
 

Private Sub his_Click()
Form12.Show 1

End Sub


Private Sub Inte_Click()
Form14.Show 1
End Sub

Private Sub NewPac_Click()
Form3.Show 1
End Sub

Private Sub Niños_Click()
Form10.Show 1
End Sub

Private Sub Nut_Click()
Form12.Show 1
End Sub

Private Sub pre_Click()
Form17.Show 1
End Sub

Private Sub Rad_Click()
Form7.Show 1
End Sub

Private Sub Reg_Click()
Form5.Show 1
End Sub

Private Sub Tec_Click()
Form19.Show 1
End Sub

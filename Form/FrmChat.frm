VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   11295
   Begin VB.CheckBox Check2 
      Caption         =   "Mostrar Alertas!"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enviar en privado"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4155
      ItemData        =   "FrmChat.frx":0000
      Left            =   6000
      List            =   "FrmChat.frx":0002
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin ChamaleonButton.ChameleonBtn BtnEnviar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Enviar"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmChat.frx":0004
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
      Left            =   8160
      TabIndex        =   3
      ToolTipText     =   "Deshacer Operacion"
      Top             =   4440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Volver"
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
      MICON           =   "FrmChat.frx":0020
      PICN            =   "FrmChat.frx":003C
      PICH            =   "FrmChat.frx":031E
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
      Interval        =   500
      Left            =   1200
      Top             =   4320
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Actualizar Usuarios"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmChat.frx":056F
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
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EnviaTexto As String

Private Sub BtnDesHacer_Click()
On Error Resume Next
If Check2.Value = 1 Then
    FrmPrincipal.Mostrar_Mensajes = True
Else
    FrmPrincipal.Mostrar_Mensajes = False
End If

FrmChat.Hide
End Sub

Private Sub BtnEnviar_Click()
On Error Resume Next
If InStr(1, UCase(User_Priv), UCase("Todos")) <> 0 Or InStr(1, UCase(User_Priv), UCase("Chat")) <> 0 _
 Or InStr(1, UCase(User_Priv), UCase("Privados")) <> 0 Or InStr(1, UCase(User_Priv), UCase("Generales")) <> 0 Then
 
 If FrmPrincipal.Winsock1.State = 7 Then
    
    If List1.ListCount = 0 Then
        MsgBox "No hay usuarios conectados!", vbInformation + vbOKOnly, "Chat!"
        Exit Sub
    ElseIf List1.ListIndex < 0 Then
        If Check1.Value = 1 Then
            MsgBox "Debe elegir a un usuario de la lista", vbInformation + vbOKOnly, "Seleccione un usuario!"
            Exit Sub
        End If
        Selectt = "TODOS"
    Else
        If Check1.Value = 1 Then
            Selectt = Mid(List1.List(List1.ListIndex), 10, Abs(InStr(1, List1.List(List1.ListIndex), ")") - 10))
        Else
            Selectt = "TODOS"
        End If
    End If
    
    EnviaTexto = "CHAT=mensaje:<nickdestino>" & Selectt & "</nickdestino><nickname>" & Usuario & "</nickname><mensaje>" & Text2.Text & "</mensaje>"
    'MsgBox EnviaTexto
    If Check1.Value = 1 Then
        EnviaTexto = "<chatgeneral=false>" & EnviaTexto
    Else
        EnviaTexto = "<chatgeneral=true>" & EnviaTexto
    End If
    
    FrmPrincipal.Winsock1.SendData EnviaTexto
    
    If Len(Text1.Text) > 60000 Then
        Text1.Text = Mid(Text1.Text, 30000)
        Text1.SelStart = Len(Text1.Text)
    End If
    
    Text1.SelStart = Len(Text1.Text)
    If Check1.Value = 1 Then
        Text1.Text = Text1.Text & Usuario & ", Envio un PRIVADO para " & Selectt & ">> " & Text2.Text & vbCrLf
    Else
        Text1.Text = Text1.Text & Usuario & "  >> " & Text2.Text & vbCrLf
    End If
    Text1.SelStart = Len(Text1.Text)
    Text2.Text = ""
 Else
    MsgBox "No se ha establecido conexion al servidor de Chat!", vbInformation + vbOKOnly, "Información"
 End If
End If
End Sub
 

Private Sub ChameleonBtn1_Click()
On Error Resume Next
If FrmPrincipal.Winsock1.State = 7 Then
    FrmPrincipal.Winsock1.SendData "<#Lista_Usuarios#>"
Else
    MsgBox "No se ha establecido conexion al servidor de Chat!", vbInformation + vbOKOnly, "Información"
End If

End Sub

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 1 Then
    'FrmPrincipal.Mostrar_Mensajes = True
Else
    'FrmPrincipal.Mostrar_Mensajes = False
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Centrar Me
FrmPrincipal.Listado_Usuarios = "1"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnEnviar_Click
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Trim(FrmPrincipal.Listado_Usuarios) <> "1" Then
    Dim Buffer As String
    
    List1.Clear
    Buffer = FrmPrincipal.Listado_Usuarios
    While InStr(1, Buffer, "@") <> 0
        'MsgBox Mid(Buffer, 1, InStr(1, Buffer, "@") - 1)
        List1.AddItem Mid(Buffer, 1, InStr(1, Buffer, "@") - 1)
        'MsgBox Mid(Buffer, InStr(1, Buffer, "@") + 1)
        Buffer = Mid(Buffer, InStr(1, Buffer, "@") + 1)
    Wend
    
    FrmPrincipal.Listado_Usuarios = "1"
End If
End Sub

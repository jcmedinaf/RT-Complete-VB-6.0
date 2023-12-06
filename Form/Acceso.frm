VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmLogin 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   1860
   ClientLeft      =   3795
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "Acceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin ChamaleonButton.ChameleonBtn BtnSalir 
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         ToolTipText     =   "Salir del Sistema"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Salir"
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
         MICON           =   "Acceso.frx":1002
         PICN            =   "Acceso.frx":101E
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
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Entrar al Sistema"
         Top             =   1200
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
         MICON           =   "Acceso.frx":11BD
         PICN            =   "Acceso.frx":11D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   1005
         Left            =   120
         Picture         =   "Acceso.frx":140E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "C&LAVE:"
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&USUARIO:"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   330
         Width           =   780
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   6600
      Picture         =   "Acceso.frx":BCCC
      Top             =   3480
      Width           =   6645
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim T As Variant
Dim b

Private Sub BtnAceptar_Click()

Query = "SELECT * FROM USUARIOS WHERE Usuario = '" & Trim(Text1.Text) & "' And CONTRASEÑA='" & Trim(Text2.Text) & "'"
Set rs = CrearRS(Query)

If Not rs.EOF Then
    
    T_U = rs.Fields("T_U").Value
    Usuario = rs.Fields("NOMBRE")
    IdUser = rs.Fields("idusuario")
    IdMedT = rs.Fields("idmedicot")
    
'    If Not IsNull(rs.Fields("Privilegios").Value) Then
'        User_Priv = rs.Fields("Privilegios")
'    Else
'        User_Priv = ""
'    End If
    
    If IsNull(IdMedT) Then
        IdMedT = 0
    End If
    
    Call Animar(FrmLogin, 100, AW_BLEND Or AW_HIDE)
    Unload Me
    FrmPrincipal.Top = 0
    FrmPrincipal.Left = 0
    FrmPrincipal.Width = Screen.Width
    FrmPrincipal.Height = Screen.Height
    FrmPrincipal.Stb1.Panels(1).Text = "Usuario: " & Usuario
    FrmPrincipal.Show
    FrmBienvenido.Show vbModal, FrmPrincipal
Else
    
    MsgBox "Error en la clave o usuario ingresado, Intente nuevamente", vbOKOnly + vbInformation, "Error de usuario o clave"
    'Text1.Text = ""
    'Text2.Text = ""
    Text1.SetFocus
    b = b + 1
    If b >= 3 Then End
End If

Call Enviar_Bitacora(IdUser, "Login", "Aceptar", "Ingreso al sistema")

End Sub

Private Sub BtnAceptar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnSalir.SetFocus
End If
End Sub



Private Sub BtnSalir_Click()
Call Animar(FrmLogin, 100, AW_BLEND Or AW_HIDE)
End
End Sub

Private Sub BtnSalir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
Call Apagar_Timbre

Dim i As Integer
Dim Est As String, Est1 As String, Est2 As String, Est3 As String
Dim RutaInformes1 As String
Dim RutaFotos1 As String
Dim RutaFotos2 As String
Dim RutaFotos3 As String

Est = String$(255, " ")
Est1 = String$(255, " ")
Est2 = String$(255, " ")
Est3 = String$(255, " ")

'i = GetFromINI("Opciones", "RutaInformes", "0", "Informes.ini")
i = GetPrivateProfileString("Opciones", "RutaInformes", "", Est, Len(Est), "Informes.ini")

If i > 0 Then
    RutaInformes1 = Trim(Est)
    RutaInformes = Mid(RutaInformes1, 1, Len(RutaInformes1) - 1)
End If
       
IU = GetPrivateProfileString("Opciones", "RutaFoto", "", Est1, Len(Est1), "Fotos.ini")

If IU > 0 Then
    RutaFotos1 = Trim(Est1)
    RutaFotos = Mid(RutaFotos1, 1, Len(RutaFotos1) - 1)
    Foto = RutaFotos
End If

IUx = GetPrivateProfileString("Opciones", "RutaFotoEmpleados", "", Est2, Len(Est2), "Rut.ini")

If IUx > 0 Then
    RutaFotos2 = Trim(Est2)
    RutaFotosE = Mid(RutaFotos2, 1, Len(RutaFotos2) - 1)
    FotoEmp = RutaFotosE
End If

IUx2 = GetPrivateProfileString("Opciones", "FotosSimulacion", "", Est3, Len(Est2), "FotosSimulacion.ini")

If IUx2 > 0 Then
    Est3 = Trim(Est3)
    FotoSimul = Mid(Est3, 1, Len(Est3) - 1)
End If

FotoSimul2 = GetFromINI("Opciones", "FSimulacion", "0", "FotosSimulacion.ini")
'Text1.Text = "jcmedinaf"
'Text2.Text = "210775030405"

'Text1.Text = "NDiaz"
'Text2.Text = "951753"


'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

'IUx = GetFromINI("Opciones", "Dir_ZETA", "", "Rut.ini")
'
'If Trim(IUx) <> "" Then
'
'    Dir_ZETA = IUx & "ServCR.txt"
'
'    Open Dir_ZETA For Input As #4
'    Dim BuffTemp As String
'    l = 1
'
'    Do Until EOF(4)
'        Line Input #4, BuffTemp
'
'        If InStr(1, UCase(BuffTemp), UCase("IP Servidor=")) <> 0 Then
'            IpRemota = Trim(Mid(BuffTemp, Pos + 13))
'        ElseIf InStr(1, UCase(BuffTemp), UCase("Puerto Servidor=")) <> 0 Then
'            PortRemoto = Trim(Mid(BuffTemp, Pos + 17))
'        End If
'
'        l = l + 1
'    Loop
'    Close #4
'
'    Conectar_al_servidor_de_Comandos
'
'Else
'    Dir_ZETA = ""
'End If

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

End Sub
Sub Conectar_al_servidor_de_Comandos()
On Error GoTo MostrarErr
'asignamos los datos de conexion
FrmPrincipal.Winsock1.Close
FrmPrincipal.Winsock1.RemoteHost = IpRemota
FrmPrincipal.Winsock1.RemotePort = Val(PortRemoto)

'conectamos el socket
'Winsock1.Close

FrmPrincipal.Winsock1.Connect

Exit Sub

MostrarErr:
    MsgBox Err.Description & " / " & Err.Number, vbInformation + vbOKOnly, "No hay servidor remoto"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Animar(FrmLogin, 100, AW_BLEND Or AW_HIDE)
End Sub

Private Sub Text1_Change()
'Dim StrText, Chaa, pru As String
'Dim i  As Variant
'StrText = ""
'Chaa = ""
'Chaa = UCase(Mid(Text1.Text, 1, 1))
'StrText = Chaa
'For i = 2 To Len(Text1.Text)
'    pru = LCase(Mid(Text1.Text, i, 1))
'    If pru Like " " Then
'        T = 1
'        StrText = StrText & " "
'    Else
'        If T = 0 Then
'            Chaa = LCase(pru)
'            StrText = StrText + Chaa
'        Else
'            Chaa = UCase(pru)
'            StrText = StrText + Chaa
'            T = 0
'        End If
'    End If
'Next i
'Text1.Text = StrText
'Text1.SelStart = Len(Text1.Text)
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Text2.SetFocus
    BtnAceptar_Click
End If

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'BtnAceptar.SetFocus
    BtnAceptar_Click
End If
End Sub



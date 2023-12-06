VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBienvenido 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bienvenido ....."
   ClientHeight    =   2370
   ClientLeft      =   6255
   ClientTop       =   5235
   ClientWidth     =   5460
   Icon            =   "FrmBiemvenida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   4560
         Top             =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         Height          =   195
         Left            =   3720
         TabIndex        =   6
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label LblHora 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   195
         Left            =   4200
         TabIndex        =   5
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label LblFecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2040
         TabIndex        =   3
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label LblEmpleado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXX"
         Height          =   195
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bienvenido al sistema OncoAmerica"
         Height          =   195
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   120
         Picture         =   "FrmBiemvenida.frx":1002
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmBienvenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Centrar Me
        Timer1.Enabled = True
        LblFecha.Caption = Date
        LblHora.Caption = Time
        CargarUsuario
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
Call Animar(FrmBienvenido, 250, AW_BLEND Or AW_HIDE)

If FrmPrincipal.Winsock1.State = sckConnected Then
    FrmPrincipal.Winsock1.SendData "{#User#}" & NombreEquipo & " Usuario: " & Usuario & " IdUser=" & IdUser & " Fecha/Hora de inicio:" & Format(Now, "dd/MM/yyyy  hh:mm:ss AMPM")
    FrmPrincipal.Winsock1.SendData "<#Lista_Usuarios#>"
End If

FrmPrincipal.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value = 100 Then

    Timer1.Enabled = False
    Call Animar(FrmBienvenido, 250, AW_BLEND Or AW_HIDE)
    Unload Me
Else

    ProgressBar1.Value = Val(ProgressBar1.Value) + Val(1)
End If

End Sub

Sub CargarUsuario()
Dim RsCargarUsuario As New ADODB.Recordset
CSql = "Select * From Usuarios Where IdUsuario='" & IdUser & "'"
Set RsCargarUsuario = CrearRS(CSql)

LblEmpleado.Caption = RsCargarUsuario.Fields("Apellidos").Value & ", " & RsCargarUsuario.Fields("Nombre").Value

If RsCargarUsuario.RecordCount <> 0 Then
    If RsCargarUsuario.Fields("Foto") <> "" Then
        If Len(Dir(FotoEmp & "\" & RsCargarUsuario.Fields("Foto").Value)) > 0 Then
            Image1.Picture = LoadPicture(FotoEmp & "\" & RsCargarUsuario.Fields("Foto").Value)
        Else
            Image1.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        End If
    Else
        Image1.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
Else
    Image1.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
End If
End Sub

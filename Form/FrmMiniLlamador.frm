VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmMiniLlamador 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mini Llamador"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "FrmMiniLlamador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   5610
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   5415
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   4320
         TabIndex        =   14
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
         MICON           =   "FrmMiniLlamador.frx":1002
         PICN            =   "FrmMiniLlamador.frx":101E
         PICH            =   "FrmMiniLlamador.frx":11E7
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         ScaleHeight     =   345
         ScaleWidth      =   4965
         TabIndex        =   15
         Top             =   5520
         Visible         =   0   'False
         Width           =   5000
      End
      Begin VB.Timer Timer4 
         Left            =   1440
         Top             =   120
      End
      Begin VB.Timer Timer3 
         Left            =   960
         Top             =   120
      End
      Begin VB.Timer Timer2 
         Left            =   480
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   0
         Top             =   120
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   5280
         Width           =   5400
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   3360
         Width           =   5400
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   4320
         Width           =   5400
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   2400
         Width           =   5400
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   1440
         Width           =   5400
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   5400
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Administración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   4920
         Width           =   5400
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Oncología"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   5400
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección Médica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   2040
         Width           =   5400
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Psicología"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   3000
         Width           =   5400
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nutrición"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   3960
         Width           =   5400
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Radioterapia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   5385
      End
   End
End
Attribute VB_Name = "FrmMiniLlamador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BdLlamado As New ADODB.Recordset  'Tabla de los pacientes llamados por los usuarios
Dim BdLlamado1 As New ADODB.Recordset  'tabla de busqueda de paciente
Dim BdLlamado2 As New ADODB.Recordset
Dim Nomb1 As String

Dim RsRutaFotoPaciente As New ADODB.Recordset
Dim RutaFoto As String

Dim x2
Dim timbre As Boolean
Dim h
Dim j
Dim p
Dim Panta As String

Sub Sonar_Timbre()
total1 = 1
PortOut &H378, total1
End Sub

Sub Apagar_Timbre()
total1 = 0
PortOut &H378, total1
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then End
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
'If Not (IsDriverInstalled) Then End
Call Blanqueo
Call Apagar_Timbre
Dim i As Integer
Dim Est As String, Est1 As String
Dim RutaFotos1 As String
T = 0
j = 1

Est = String$(255, " ")

IU = GetPrivateProfileString("Opciones", "RutaFoto", "", Est, Len(Est), "Fotos.ini")
If IU > 0 Then
    RutaFotos1 = Trim(Est)
    RutaFoto = Mid(RutaFotos1, 1, Len(RutaFotos1) - 1)
    Foto = RutaFoto
End If

Est1 = String$(50, " ")
i = GetPrivateProfileString("Tabla", "Panta", "", Est1, Len(Est1), "Info.ini")
If i > 0 Then
    'RutaFotos1 = Trim(Est)
    Panta1 = Mid(Trim(Est1), 1, Len(Trim(Est1)) - 1)
    Panta = Trim(Panta1)
End If
End Sub
Sub Presenta()
If Nomb1 <> "" Then
    Call Sonar_Timbre
    'Call Espera(1)
    timbre = True
    Timer3.Interval = 200
    'Picture1.Top = 0
    Timer1.Interval = 0
    Timer4.Interval = 1
'    Picture1.Visible = True
Else
    Call Apagar_Timbre
    Picture1.Top = 20000
    Timer3.Interval = 0
    Timer1.Interval = 600
    Timer4.Interval = 0
    Picture1.Visible = False
    'Call Apagar_Timbre
End If
End Sub

Sub Blanqueo()

Label1.Caption = ""
Label3.Caption = ""
Label5.Caption = ""
Label8.Caption = ""
Label10.Caption = ""
Label11.Caption = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Apagar_Timbre
End Sub

Private Sub Form_Terminate()
Call Apagar_Timbre
End Sub


Private Sub Timer1_Timer()
'On Error GoTo errty
CSql = "Select * From " & Panta & " where pantalla = 0"
Set BdLlamado = CrearRS(CSql)

'DoEvents
If Not BdLlamado.EOF Then
    Do While Not BdLlamado.EOF
        CSql = "Select NombreP, ApellidoP, Foto From Paciente Where IdPaciente = " & BdLlamado.Fields("idpaciente")
        Set BdLlamado1 = CrearRS(CSql)
        If Not BdLlamado1.EOF Then
            bf = BdLlamado1.Fields("NombreP") & " " & BdLlamado1.Fields("ApellidoP")
        Else
            bf = ""
        End If
        If BdLlamado.Fields("Pantalla") = 0 Then
            CSql = ""
            Select Case BdLlamado.Fields("modulo")
                Case Is = 0 'nutricion

                    Label3.Caption = bf
                    Nomb1 = Label3.Caption
                    CSql = "update " & Panta & " set MiniLlamador = 1 where idpaciente = " & BdLlamado.Fields("idpaciente")
                    Call Presenta
                   ' If bf <> "" Then If BdLlamado1.Fields("foto") <> "" Then Image1.Picture = LoadPicture(Foto & "\" & BdLlamado1.Fields("foto")) Else Image1.Picture = LoadPicture(Foto & "\" & "Silueta.jpg")

                Case Is = 1 'Psicologia

                    Label1.Caption = bf
                    Nomb1 = Label1.Caption
                    CSql = "update " & Panta & " set MiniLlamador = 1 where idpaciente = " & BdLlamado.Fields("idpaciente")
                    Call Presenta
                   ' If bf <> "" Then If BdLlamado1.Fields("foto") <> "" Then Image2.Picture = LoadPicture(Foto & "\" & BdLlamado1.Fields("foto")) Else Image2.Picture = LoadPicture(Foto & "\" & "Silueta.jpg")

                Case Is = 2 'Radioterapia

                    Label5.Caption = bf
                    Nomb1 = Label5.Caption
                    CSql = "update " & Panta & " set MiniLlamador = 1 where idpaciente = " & BdLlamado.Fields("idpaciente")
                    Call Presenta
                   ' If bf <> "" Then If BdLlamado1.Fields("foto") <> "" Then Image3.Picture = LoadPicture(Foto & "\" & BdLlamado1.Fields("foto")) Else Image3.Picture = LoadPicture(Foto & "\" & "Silueta.jpg")

                Case Is = 3 'Internista

                    Label8.Caption = bf
                    Nomb1 = Label8.Caption
                    CSql = "update " & Panta & " set MiniLlamador = 1 where idpaciente = " & BdLlamado.Fields("idpaciente")
                    Call Presenta
                    'If bf <> "" Then If BdLlamado1.Fields("foto") <> "" Then Image4.Picture = LoadPicture(Foto & "\" & BdLlamado1.Fields("foto")) Else Image4.Picture = LoadPicture(Foto & "\" & "Silueta.jpg")

               Case Is = 4 'Oncologia

                    Label10.Caption = bf
                    Nomb1 = Label10.Caption
                    CSql = "update " & Panta & " set MiniLlamador = 1 where idpaciente = " & BdLlamado.Fields("idpaciente")
                    Call Presenta
                   ' If bf <> "" Then If BdLlamado1.Fields("foto") <> "" Then Image5.Picture = LoadPicture(Foto & "\" & BdLlamado1.Fields("foto")) Else Image5.Picture = LoadPicture(Foto & "\" & "Silueta.jpg")

                Case Is = 5 'Administración

                    Label11.Caption = bf
                    Nomb1 = Label11.Caption
                    CSql = "update " & Panta & " set MiniLlamador = 1 where idpaciente = " & BdLlamado.Fields("idpaciente")
                    Call Presenta
                   ' If bf <> "" Then If BdLlamado1.Fields("foto") <> "" Then Image6.Picture = LoadPicture(Foto & "\" & BdLlamado1.Fields("foto")) Else Image6.Picture = LoadPicture(Foto & "\" & "Silueta.jpg")

            End Select
            Set BdLlamado2 = CrearRS(CSql)
        End If
        'DoEvents
        BdLlamado1.Close
        BdLlamado.MoveNext
    Loop
End If
BdLlamado.Close
'On Error GoTo 0
Exit Sub

errty:
If BdLlamado.State Then BdLlamado.Close
If BdLlamado1.State Then BdLlamado.Close
If BdLlamado2.State Then BdLlamado.Close
Call Apagar_Timbre
timbre = False
Exit Sub
End Sub

Private Sub Timer2_Timer()

Picture1.FontName = "Arial"
Picture1.FontBold = True
Picture1.FontSize = 300
Letrero = Nomb1 & Space(6)
Static Anterior1 As Boolean
Static tamañoLetrero1 As Single
Static x1 As Single

If Not Anterior1 Then
    tamañoLetrero1 = Picture1.TextWidth(Letrero)
    Anterior1 = True
    x1 = Picture1.ScaleWidth
    x2 = x1

End If
Picture1.Cls

Picture1.CurrentX = x1
Picture1.CurrentY = 0
'Para cambiar el tipo de letra

Picture1.Print Letrero

'MsgBox x1
x1 = x1 - 200
'If x1 = x2 - (80 * 200) Then

'End If
If x1 < -tamañoLetrero1 Then
    Nomb1 = ""
    Anterior1 = False
    Call Presenta
End If

End Sub

Private Sub Timer3_Timer()
If timbre Then Call Apagar_Timbre: timbre = False: Timer3.Interval = 0
End Sub

Private Sub Timer4_Timer()
Picture1.FontName = "Arial"
Picture1.FontBold = True
Picture1.FontSize = 14
Picture1.Visible = True
Letrero = Nomb1 '& Space(6)
Static Anterior1 As Boolean
Static tamañoLetrero1 As Single
Static x1 As Single
movement = 150

cx = Picture1.Width * (h / movement / 8)
cy = Picture1.Height * (h / movement * 8)
X = Picture1.Left + ((Picture1.Width - cx) / 500)
Y = (Picture1.Top + (Picture1.Height - cy)) / 2
Picture1.Left = cx
Picture1.Top = cy
'If Picture1.Top > 13000 Then
If Picture1.Top > 5895 Then
    j = -1
    p = p + 1
    If p >= 3 Then
        p = 0
        Nomb1 = ""
        Call Presenta
        h = 1
        Exit Sub
    End If
End If

If Picture1.Top < 0 Then j = 1
Picture1.Height = 495
Picture1.Width = 5000
Picture1.Cls:  Picture1.Print Letrero

'Rectangle TheScreen, X, Y, X + cx, Y + cy

h = h + j

End Sub


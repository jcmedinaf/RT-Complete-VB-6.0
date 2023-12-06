VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCapturarFoto 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tomar Foto"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "FrmCapturar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EAEFEF&
         Height          =   2775
         Left            =   6480
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   197
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EAEFEF&
         Height          =   2775
         Left            =   120
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   197
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EAEFEF&
         FillColor       =   &H000000FF&
         Height          =   2775
         Left            =   3240
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   197
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Ac&eptar"
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
         MICON           =   "FrmCapturar.frx":1002
         PICN            =   "FrmCapturar.frx":101E
         PICH            =   "FrmCapturar.frx":1253
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCapturar 
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Capturar Imagen"
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
         MICON           =   "FrmCapturar.frx":14B4
         PICN            =   "FrmCapturar.frx":14D0
         PICH            =   "FrmCapturar.frx":176C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCancelar 
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "C&ancelar"
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
         MICON           =   "FrmCapturar.frx":19ED
         PICN            =   "FrmCapturar.frx":1A09
         PICH            =   "FrmCapturar.frx":1BAD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen Ajustada"
         Height          =   195
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen Capturada"
         Height          =   195
         Left            =   3240
         TabIndex        =   6
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen Previa"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edicion"
      Visible         =   0   'False
      Begin VB.Menu mnuRecortar 
         Caption         =   "&Recortar"
      End
      Begin VB.Menu separador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelar 
         Caption         =   "&Cancelar Seleccion"
      End
   End
End
Attribute VB_Name = "FrmCapturarFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ws_visible = &H10000000
Const ws_child = &H40000000
Const WM_USER = 1024
Const WM_CAP_EDIT_COPY = WM_USER + 30
Const WM_cap_driver_connect = WM_USER + 10
Const WM_cap_set_preview = WM_USER + 50
Const wm_cap_set_overlay = WM_USER + 51
Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Const WM_CAP_SEQUENCE = WM_USER + 62
Const WM_CAP_SINGLE_FRAME_OPEN = WM_USER + 70
Const WM_CAP_SINGLE_FRAME_CLOSE = WM_USER + 71
Const WM_CAP_SINGLE_FRAME = WM_USER + 72
Const DRV_USER = &H4000
Const DVM_DIALOG = DRV_USER + 100
Const PREVIEWRATE = 30

' Variables para la edicion de la Foto
Dim Recuadrar As Boolean
Dim ReTam As Boolean
Private Type TArea
    x1 As Single
    x2 As Single
    y1 As Single
    y2 As Single
End Type
Dim Area As TArea
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal a As String, ByVal b As Long, ByVal C As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, ByVal g As Long, ByVal h As Integer) As Long

Dim hwndc As Long

Private Sub BtnCancelar_Click()
Unload Me
End Sub

Private Sub BtnCapturar_Click()
On Error Resume Next
Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height
temp = SendMessage(hwndc, WM_CAP_EDIT_COPY, 1, 0)
Set Picture2.Picture = Clipboard.GetData
End Sub

Private Sub BtnCerrar_Click()
On Error Resume Next

Dim FotoPaciente As String

If UCase(Tipo) = UCase("Nuevo Empleado") Then
    FotoP = Replace(Trim(FrmEmpleados.TxtCedulaEmp.Text) & Trim(FrmEmpleados.TxtApellido.Text) & Trim(FrmEmpleados.TxtNombre.Text) & ".jpg", " ", "", , , vbTextCompare)
    FotoPaciente = Replace(Foto & "\" & Trim(FrmEmpleados.TxtCedulaEmp.Text) & Trim(FrmEmpleados.TxtApellido.Text) & Trim(FrmEmpleados.TxtNombre.Text) & ".jpg", " ", "", , , vbTextCompare)
     
    If Picture2.Picture > 0 Then
        SavePicture Picture2.Picture, FotoPaciente
        FrmEmpleados.Image3.Picture = Picture2.Picture
        Else
        SavePicture FrmPrincipal.ListaImagenes.ListImages(1).Picture, FotoPaciente
        FrmEmpleados.Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
Else
    FotoP = Replace(Trim(FrmNuevoPaciente.Text1.Text) & Trim(FrmNuevoPaciente.Text3.Text) & Trim(FrmNuevoPaciente.Text4.Text) & ".jpg", " ", "", , , vbTextCompare)
    FotoPaciente = Replace(Foto & "\" & Trim(FrmNuevoPaciente.Text1.Text) & Trim(FrmNuevoPaciente.Text3.Text) & Trim(FrmNuevoPaciente.Text4.Text) & ".jpg", " ", "", , , vbTextCompare)
     
    If Picture2.Picture <> 0 Then
        SavePicture Picture2.Picture, FotoPaciente
        FrmNuevoPaciente.Image3.Picture = Picture2.Picture
        Else
        SavePicture FrmPrincipal.ListaImagenes.ListImages(1).Picture, FotoPaciente
        FrmNuevoPaciente.Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
    End If
End If

Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Dir1 As Label
Dim temp As Long
 
hwndc = capCreateCaptureWindow("CapWindow", ws_child Or ws_visible, 0, 0, 340, 240, Picture1.hwnd, 0)
  
'esto conecta el dispositivo.
 
If (hwndc <> 0) Then
    temp = SendMessage(hwndc, WM_cap_driver_connect, 0, 0)
    
    If temp = 0 Then GoTo Dir1
    'esto hace que la imagen recibida por el dispositivo se ajuste
    'al tamaño de la ventana de captura (justo lo que yo buscaba)
    temp = SendMessage(hwndc, WM_cap_set_preview, 1, 0)
    'esto setea la tasa de transferencia (vease MDSN)
    temp = SendMessage(hwndc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    'esto activa el modo PREVIEW para ver video lo cual es suficiente
    'para tener movimiento en mi programita
    temp = SendMessage(Me.hwnd, WM_CAP_SET_SCALE, True, 0)
    'esto hace que la imagen recibida por el dispositivo se ajuste
    'al tamaño de la ventana de captura (justo lo que yo buscaba)
    DoEvents
    startcap = True
    Picture1.Visible = True
Else
Dir1:
    MsgBox "No hay Camaras Webs Instaladas", vbOKOnly + vbCritical, "Error"
    'Unload FrmCapturarFoto
End If
End Sub

Private Sub Timer1_Timer()
If startcap = False Then Unload FrmCapturarFoto
End Sub

Private Sub mnuCancelar_Click()
    ReTam = False
    Area.x1 = 0
    Area.y1 = 0
    Area.x2 = Picture2.ScaleWidth
    Area.y2 = Picture2.ScaleHeight
    Picture3.Height = Picture2.Height
    Picture3.Width = Picture2.Width
    Picture3.Picture = Picture2.Picture
    Picture2.Cls
End Sub

Private Sub mnuRecortar_Click()
On Error Resume Next
        ReTam = True
        Picture3.Cls
        DoEvents

        Picture3.Width = Abs(Area.x1 - Area.x2)
        Picture3.Height = Abs(Area.y1 - Area.y2)

        If Area.x1 < Area.x2 And Area.y1 < Area.y2 Then
            DoEvents
            Picture3.PaintPicture Picture2.Picture, 0, 0, _
                                                 Abs(Area.x2 - Area.x1), Abs(Area.y2 - Area.y1), _
                                                 Area.x1, Area.y1, _
                                                 Abs(Area.x2 - Area.x1), Abs(Area.y2 - Area.y1)

        ElseIf Area.x1 > Area.x2 And Area.y1 > Area.y2 Then
            DoEvents
            Picture3.PaintPicture Picture2.Picture, 0, 0, _
                                                 Abs(Area.x1 - Area.x2), Abs(Area.y1 - Area.y2), _
                                                 Area.x2, Area.y2, _
                                                 Abs(Area.x1 - Area.x2), Abs(Area.y1 - Area.y2)

        ElseIf Area.x1 > Area.x2 And Area.y1 < Area.y2 Then
            DoEvents
            Picture3.PaintPicture Picture2.Picture, 0, 0, _
                                                Area.x1 + Area.x2, Area.y1 + Area.y2, _
                                                Area.x2, Area.y1, _
                                                Area.x1 + Area.x2, Area.y1 + Area.y2

        ElseIf Area.x1 < Area.x2 And Area.y1 > Area.y2 Then
            DoEvents
            Picture3.PaintPicture Picture2.Picture, 0, 0, _
                                                Area.x1 + Area.x2, Area.y1 + Area.y2, _
                                                Area.x1, Area.y2, _
                                                Area.x1 + Area.x2, Area.y1 + Area.y2
        End If
        DoEvents
        Clipboard.Clear
        Clipboard.SetData Picture3.Image, vbCFBitmap

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbRightButton
             Me.PopupMenu mnuEdicion
        Case vbLeftButton
            
            Area.x1 = 0: Area.x2 = 0: Area.y1 = 0: Area.y2 = 0
            
            Area.x1 = X: Area.y1 = Y
            Picture2.Cls
            Recuadrar = True
    End Select
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   If Recuadrar Then
        Picture2.Cls
        If X > Picture2.ScaleWidth Then X = Picture2.ScaleWidth - 1
        If Y > Picture2.ScaleHeight Then Y = Picture2.ScaleHeight - 1
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        Area.x2 = X: Area.y2 = Y
        Picture2.Line (Area.x1, Area.y1)-(Area.x2, Area.y2), vbBlack, B
    End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Recuadrar = False
    If Area.x1 = X And Area.y1 = Y Then
       Area.x1 = 0: Area.x2 = 0: Area.y1 = 0: Area.y2 = 0
       Exit Sub
    End If
    If X > Picture2.ScaleWidth Then Area.x2 = Picture2.ScaleWidth - 1
    If Y > Picture2.ScaleHeight Then Area.y2 = Picture2.ScaleHeight - 1
    If X < 0 Then Area.x2 = 0
    If Y < 0 Then Area.y2 = 0
End Sub


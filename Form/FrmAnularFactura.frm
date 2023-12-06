VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmAnularFactura 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anular Factura"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "FrmAnularFactura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   5295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Cerrar "
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         MICON           =   "FrmAnularFactura.frx":1002
         PICN            =   "FrmAnularFactura.frx":101E
         PICH            =   "FrmAnularFactura.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar"
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
         MICON           =   "FrmAnularFactura.frx":141C
         PICN            =   "FrmAnularFactura.frx":1438
         PICH            =   "FrmAnularFactura.frx":16C7
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
      Caption         =   "Motivos de la Anulación de Factura"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5295
      Begin MSComCtl2.DTPicker DtpFechaCorrecta 
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53215233
         CurrentDate     =   40176
      End
      Begin VB.TextBox TxtMotivosFacturaAnulada 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   5055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por Cualquier Otro Motivo."
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por Error de Fecha. Fecha Correcta es:"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por Devolución."
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Agregar"
      Top             =   0
      Width           =   5295
      Begin VB.Label LblNoFactura 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anular Factura No.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2355
      End
   End
End
Attribute VB_Name = "FrmAnularFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Motivo As String
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim RsHistoricoAnularFactura As New ADODB.Recordset
Dim VClave As String

VClave = InputBox("Ingrese su clave: ", "Confirmar", "XXXXXXX")

CSql = "Select Contraseña From usuarios where IdUsuario=" & IdUser & ""
Set RsHistoricoAnularFactura = CrearRS(CSql)

If RsHistoricoAnularFactura.RecordCount <> 0 Then
    If Not RsHistoricoAnularFactura.Fields("Contraseña") = VClave Then
        MsgBox "Clave incorrecta!", vbCritical + vbOKOnly, "Error"
        Call Enviar_Bitacora(IdUser, "ANULAR FACTURA", "GUARDAR", "Clave incorrecta! al intenter ANULAR la factura Nro." & Val(LblNoFactura.Caption) & ", se ingreso la clave =" & VClave)
        Exit Sub
    End If
Else
    MsgBox "Hubo un error en la base de datos", vbCritical + vbOKOnly, "Su cuenta no existe!"
    Exit Sub
End If


CSql = "Select * From FacturaAnulada"
Set RsHistoricoAnularFactura = CrearRS(CSql)

RsHistoricoAnularFactura.AddNew
RsHistoricoAnularFactura.Fields("NoFactura").Value = LblNoFactura.Caption
RsHistoricoAnularFactura.Fields("Motivo").Value = Motivo
RsHistoricoAnularFactura.Fields("FechaAnulacion").Value = Format(Date, "DD/MM/YYYY")
RsHistoricoAnularFactura.Fields("IdUser").Value = IdUser
RsHistoricoAnularFactura.Update

CSql = "UPDATE C_Cobrar set Anulada='1' where N_Factura=" & Val(LblNoFactura.Caption) & ""
Set RsHistoricoAnularFactura = CrearRS(CSql)

MsgBox "La factura ha sido Anulada!", vbInformation + vbOKOnly, "Operacion Exitosa."
Unload Me
FrmLibroVentas.BtnBuscar_Click
End Sub

Private Sub DtpFechaCorrecta_Change()
Motivo = "Por Error de Fecha. Fecha Correcta es: " & Format(DtpFechaCorrecta, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Centrar Me
LblNoFactura.Caption = FrmLibroVentas.Fa
End Sub

Private Sub Option1_Click()
Motivo = "Por Devolución"
End Sub

Private Sub Option2_Click()
DtpFechaCorrecta.Enabled = True
DtpFechaCorrecta.SetFocus
End Sub

Private Sub Option3_Click()
TxtMotivosFacturaAnulada.Enabled = True
End Sub

Private Sub TxtMotivosFacturaAnulada_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(TxtMotivosFacturaAnulada.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(TxtMotivosFacturaAnulada.Text)
    pru = LCase(Mid(TxtMotivosFacturaAnulada.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

TxtMotivosFacturaAnulada.Text = StrText
TxtMotivosFacturaAnulada.SelStart = Len(TxtMotivosFacturaAnulada.Text)
Motivo = TxtMotivosFacturaAnulada.Text
End Sub

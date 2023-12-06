VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmBeneficiario 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beneficiarios"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "FrmBeneficiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7875
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   7455
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6360
            TabIndex        =   11
            ToolTipText     =   "Cerrar "
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "FrmBeneficiario.frx":1002
            PICN            =   "FrmBeneficiario.frx":101E
            PICH            =   "FrmBeneficiario.frx":11E7
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
            Left            =   1200
            TabIndex        =   7
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
            MICON           =   "FrmBeneficiario.frx":141C
            PICN            =   "FrmBeneficiario.frx":1438
            PICH            =   "FrmBeneficiario.frx":16C7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregar 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar"
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
            MICON           =   "FrmBeneficiario.frx":1B08
            PICN            =   "FrmBeneficiario.frx":1B24
            PICH            =   "FrmBeneficiario.frx":1CB1
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
            Left            =   5040
            TabIndex        =   10
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Deshacer"
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
            MICON           =   "FrmBeneficiario.frx":1EE6
            PICN            =   "FrmBeneficiario.frx":1F02
            PICH            =   "FrmBeneficiario.frx":21E4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
            Height          =   375
            Left            =   4320
            TabIndex        =   9
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "FrmBeneficiario.frx":2435
            PICN            =   "FrmBeneficiario.frx":2451
            PICH            =   "FrmBeneficiario.frx":26E7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAnterior 
            Height          =   375
            Left            =   3720
            TabIndex        =   8
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "FrmBeneficiario.frx":2946
            PICN            =   "FrmBeneficiario.frx":2962
            PICH            =   "FrmBeneficiario.frx":2BF7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnEliminar 
            Height          =   375
            Left            =   2400
            TabIndex        =   19
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Borrar"
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
            MICON           =   "FrmBeneficiario.frx":2E53
            PICN            =   "FrmBeneficiario.frx":2E6F
            PICH            =   "FrmBeneficiario.frx":3013
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
         Caption         =   "Registro Beneficiario"
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7455
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Agente de Retención"
            Height          =   255
            Left            =   5400
            TabIndex        =   20
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox TxtEmail 
            Height          =   375
            Left            =   1200
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2760
            Width           =   6135
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   375
            Left            =   1200
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox TxtDireccion 
            Height          =   735
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1440
            Width           =   6135
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   375
            Left            =   1200
            TabIndex        =   1
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   375
            Left            =   1200
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   960
            Width           =   6135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   2850
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   2370
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1050
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "FrmBeneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGuardar As New ADODB.Recordset
Dim RsActualiza As New ADODB.Recordset
Dim RsBeneficiario As New ADODB.Recordset
Dim RegNew As Integer

Private Sub BtnAgregar_Click()
RegNew = 1
Blanqueo
TxtCodigo.SetFocus
End Sub

Private Sub BtnAnterior_Click()
If Not (RsBeneficiario.BOF) Then RsBeneficiario.MovePrevious
If Not (RsBeneficiario.BOF) Then Call CargaDatos
Cambio = 0
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub
 
Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""
TxtDireccion.Text = ""
TxtTelefono.Text = ""
TxtEmail.Text = ""

End Sub

Private Sub BtnGuardarActualizar_Click()
If RegNew = 0 Then
    Call GuardarCambios
ElseIf RegNew = 1 Then
    Call Guardar
End If
End Sub

Sub GuardarCambios()
CSql = "Select * From Beneficiario Where IdBeneficiario = '" & IdMax & "'"
Set RsActualiza = CrearRS(CSql)

RsGuardar.Fields("CodigoBeneficiario").Value = TxtCodigo.Text
RsGuardar.Fields("DescripcionBeneficiario").Value = TxtDescripcion.Text
RsGuardar.Fields("DireccionBeneficiario").Value = TxtDireccion.Text
RsGuardar.Fields("TelefonoBeneficiario").Value = TxtTelefono.Text
RsGuardar.Fields("EmailBeneficiario").Value = TxtEmail.Text
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Fields("Retencion").Value = Check1.Value
RsGuardar.Update

Msg = "Registro Actualizado Satisfactoriamente"
MsgBox Msg, vbInformation + vbOKOnly, "Registro Actualizado"
End Sub

Sub Guardar()

Dim IdMax As Integer
CSql = "Select max(IdBeneficiario)+1 as MaxId From Beneficiario"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdMax = RsMaxId.Fields("MaxId").Value
Else
    IdMax = "1"
End If

CSql = "Select * From Beneficiario"
Set RsGuardar = CrearRS(CSql)

RsGuardar.AddNew
RsGuardar.Fields("IdBeneficiario").Value = IdMax
RsGuardar.Fields("CodigoBeneficiario").Value = TxtCodigo.Text
RsGuardar.Fields("DescripcionBeneficiario").Value = TxtDescripcion.Text
RsGuardar.Fields("DireccionBeneficiario").Value = TxtDireccion.Text
RsGuardar.Fields("TelefonoBeneficiario").Value = TxtTelefono.Text
RsGuardar.Fields("EmailBeneficiario").Value = TxtEmail.Text
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Fields("Retencion").Value = Check1.Value
RsGuardar.Update

Msg = "Registro Agregado Satisfactoriamente"
MsgBox Msg, vbInformation + vbOKOnly, "Registro Guardado"
Blanqueo

End Sub

Private Sub BtnSiguiente_Click()
If Not (RsBeneficiario.EOF) Then RsBeneficiario.MoveNext
If Not (RsBeneficiario.EOF) Then Call CargaDatos

Cambio = 0
End Sub

Private Sub Form_Load()
Centrar Me
RegNew = 0
CSql = "Select * From Beneficiario"
Set RsBeneficiario = CrearRS(CSql)
CargaDatos
End Sub

Sub CargaDatos()
If Not RsBeneficiario.EOF Then
    If IsNull(RsBeneficiario.Fields("IdBeneficiario")) Then IdMax = "" Else IdMax = RsBeneficiario.Fields("IdBeneficiario")
    If IsNull(RsBeneficiario.Fields("CodigoBeneficiario")) Then TxtCodigo.Text = "" Else TxtCodigo.Text = RsBeneficiario.Fields("CodigoBeneficiario")
    If IsNull(RsBeneficiario.Fields("DescripcionBeneficiario")) Then TxtDescripcion.Text = "" Else TxtDescripcion.Text = RsBeneficiario.Fields("DescripcionBeneficiario")
    
    If IsNull(RsBeneficiario.Fields("DireccionBeneficiario")) Then TxtDireccion.Text = "" Else TxtDireccion.Text = RsBeneficiario.Fields("DireccionBeneficiario")
    If IsNull(RsBeneficiario.Fields("TelefonoBeneficiario")) Then TxtTelefono.Text = "" Else TxtTelefono.Text = RsBeneficiario.Fields("TelefonoBeneficiario")
    If IsNull(RsBeneficiario.Fields("EmailBeneficiario")) Then TxtEmail.Text = "" Else TxtEmail.Text = RsBeneficiario.Fields("EmailBeneficiario")
    
    If RsBeneficiario.Fields("retencion").Value = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    
    
End If

End Sub


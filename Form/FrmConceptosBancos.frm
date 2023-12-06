VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConceptosBancos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "FrmConceptosBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7905
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Registro Banco"
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7455
         Begin VB.TextBox TxtDescripcion 
            Height          =   1095
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   960
            Width           =   6135
         End
         Begin VB.TextBox TxtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1200
            TabIndex        =   1
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   570
            Width           =   540
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   7455
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6360
            TabIndex        =   8
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
            MICON           =   "FrmConceptosBancos.frx":1002
            PICN            =   "FrmConceptosBancos.frx":101E
            PICH            =   "FrmConceptosBancos.frx":11E7
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
            TabIndex        =   4
            ToolTipText     =   "Guardar / Actualizar "
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
            MICON           =   "FrmConceptosBancos.frx":141C
            PICN            =   "FrmConceptosBancos.frx":1438
            PICH            =   "FrmConceptosBancos.frx":16C7
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
            TabIndex        =   3
            ToolTipText     =   "Agregar "
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
            MICON           =   "FrmConceptosBancos.frx":1B08
            PICN            =   "FrmConceptosBancos.frx":1B24
            PICH            =   "FrmConceptosBancos.frx":1CB1
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
            Left            =   5160
            TabIndex        =   7
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
            MICON           =   "FrmConceptosBancos.frx":1EE6
            PICN            =   "FrmConceptosBancos.frx":1F02
            PICH            =   "FrmConceptosBancos.frx":21E4
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
            Left            =   4440
            TabIndex        =   6
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
            MICON           =   "FrmConceptosBancos.frx":2435
            PICN            =   "FrmConceptosBancos.frx":2451
            PICH            =   "FrmConceptosBancos.frx":26E7
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
            TabIndex        =   5
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
            MICON           =   "FrmConceptosBancos.frx":2946
            PICN            =   "FrmConceptosBancos.frx":2962
            PICH            =   "FrmConceptosBancos.frx":2BF7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnElimina 
            Height          =   375
            Left            =   2400
            TabIndex        =   13
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
            MICON           =   "FrmConceptosBancos.frx":2E53
            PICN            =   "FrmConceptosBancos.frx":2E6F
            PICH            =   "FrmConceptosBancos.frx":3013
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
   End
End
Attribute VB_Name = "FrmConceptosBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsConceptosBancos As New ADODB.Recordset
Dim RegNew As Integer
Private Sub BtnAgregar_Click()
RegNew = 1
Blanqueo
End Sub

Private Sub BtnAnterior_Click()
If Not (RsConceptosBancos.BOF) Then RsConceptosBancos.MovePrevious
If Not (RsConceptosBancos.BOF) Then Call CargaDatos
Cambio = 0
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""

End Sub

Private Sub BtnGuardarActualizar_Click()
If RegNew = 0 Then
    Call GuardarCambios
ElseIf RegNew = 1 Then
    Call Guardar
End If
End Sub

Sub Guardar()
Dim RsMaxId As New ADODB.Recordset
Dim RsGuardar As New ADODB.Recordset
Dim IdMax As Integer
CSql = "Select max(IdConceptoBancos)+1 as MaxId From ConceptosBancos"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdMax = RsMaxId.Fields("MaxId").Value
Else
    IdMax = "1"
End If

CSql = "Select * From ConceptosBancos"
Set RsGuardar = CrearRS(CSql)

RsGuardar.AddNew
RsGuardar.Fields("IdConceptoBancos").Value = IdMax
RsGuardar.Fields("CodigoConceptoBancos").Value = TxtCodigo.Text
RsGuardar.Fields("DescripcionConceptoBancos").Value = TxtDescripcion.Text
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Update

MsgBox "Se Guardo Correctamente?", vbInformation + vbOKOnly, "Registro Guardado"
   
Blanqueo
End Sub

Sub GuardarCambios()
Dim RsGuardar As New ADODB.Recordset
Dim IdMax As Integer

CSql = "Select * From ConceptosBancos Where IdConceptoBancos='" & IdMax & "'"
Set RsGuardar = CrearRS(CSql)

RsGuardar.Fields("CodigoConceptoBancos").Value = TxtCodigo.Text
RsGuardar.Fields("DescripcionConceptoBancos").Value = TxtDescripcion.Text
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Update

MsgBox "Se Actualizo Correctamente?", vbInformation + vbOKOnly, "Registro Actualizado"
   
Blanqueo

End Sub

Private Sub BtnSiguiente_Click()
If Not (RsConceptosBancos.EOF) Then RsConceptosBancos.MoveNext
If Not (RsConceptosBancos.EOF) Then Call CargaDatos

Cambio = 0
End Sub

Private Sub Form_Load()
Centrar Me
RegNew = 0
CSql = "Select * From ConceptosBancos"
Set RsConceptosBancos = CrearRS(CSql)
CargaDatos
End Sub

Sub CargaDatos()
If Not RsConceptosBancos.EOF Then
    If IsNull(RsConceptosBancos.Fields("IdConceptoBancos")) Then IdMax = "" Else IdMax = RsConceptosBancos.Fields("IdConceptoBancos")
    If IsNull(RsConceptosBancos.Fields("CodigoConceptoBancos")) Then TxtCodigo.Text = "" Else TxtCodigo.Text = RsConceptosBancos.Fields("CodigoConceptoBancos")
    If IsNull(RsConceptosBancos.Fields("DescripcionConceptoBancos")) Then TxtDescripcion.Text = "" Else TxtDescripcion.Text = RsConceptosBancos.Fields("DescripcionConceptoBancos")
End If

End Sub

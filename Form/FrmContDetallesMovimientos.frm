VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContDetallesMovimientos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de Movimientos"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   Icon            =   "FrmContDetallesMovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   7335
         Begin VB.TextBox TxtDescripcion 
            Height          =   735
            Left            =   1080
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   6135
         End
         Begin VB.TextBox TxtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   810
            Width           =   885
         End
         Begin VB.Label NReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   2880
            TabIndex        =   13
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   540
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   7335
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6240
            TabIndex        =   9
            ToolTipText     =   "Cerrar"
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
            MICON           =   "FrmContDetallesMovimientos.frx":1002
            PICN            =   "FrmContDetallesMovimientos.frx":101E
            PICH            =   "FrmContDetallesMovimientos.frx":11E7
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
            MICON           =   "FrmContDetallesMovimientos.frx":141C
            PICN            =   "FrmContDetallesMovimientos.frx":1438
            PICH            =   "FrmContDetallesMovimientos.frx":16C7
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
            MICON           =   "FrmContDetallesMovimientos.frx":1B08
            PICN            =   "FrmContDetallesMovimientos.frx":1B24
            PICH            =   "FrmContDetallesMovimientos.frx":1CB1
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
            TabIndex        =   8
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
            MICON           =   "FrmContDetallesMovimientos.frx":1EE6
            PICN            =   "FrmContDetallesMovimientos.frx":1F02
            PICH            =   "FrmContDetallesMovimientos.frx":21E4
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
            TabIndex        =   7
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
            MICON           =   "FrmContDetallesMovimientos.frx":2435
            PICN            =   "FrmContDetallesMovimientos.frx":2451
            PICH            =   "FrmContDetallesMovimientos.frx":26E7
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
            TabIndex        =   6
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
            MICON           =   "FrmContDetallesMovimientos.frx":2946
            PICN            =   "FrmContDetallesMovimientos.frx":2962
            PICH            =   "FrmContDetallesMovimientos.frx":2BF7
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
            TabIndex        =   5
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
            MICON           =   "FrmContDetallesMovimientos.frx":2E53
            PICN            =   "FrmContDetallesMovimientos.frx":2E6F
            PICH            =   "FrmContDetallesMovimientos.frx":3013
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
      Begin ChamaleonButton.ChameleonBtn BrnListaEmplesas 
         Height          =   375
         Left            =   5520
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Lista de Empresas"
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
         MICON           =   "FrmContDetallesMovimientos.frx":31B2
         PICN            =   "FrmContDetallesMovimientos.frx":31CE
         PICH            =   "FrmContDetallesMovimientos.frx":3457
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnDetMovimientos 
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   1800
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Lista de Detalles de Movimientos"
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
         MICON           =   "FrmContDetallesMovimientos.frx":3872
         PICN            =   "FrmContDetallesMovimientos.frx":388E
         PICH            =   "FrmContDetallesMovimientos.frx":3B17
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
Attribute VB_Name = "FrmContDetallesMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDetallesMovimientos As New ADODB.Recordset
Dim RsMaxId As New ADODB.Recordset
Dim IdMax As Integer
Dim RsCargarDetalles As New ADODB.Recordset
Dim RsEliminar As New ADODB.Recordset
Public IdEmpresa As Integer
Public NewReg As Byte
Public IdDeta As Integer
Dim RsTemp As Recordset

Public Sub Cargar_Det_Mov()
CSql = "SELECT * FROM ContDetallesMovimientos WHERE IdEmpresa=" & IdEmpresa & " AND Activo='1'"
Set RsCargarDetalles = CrearRS(CSql)
CargarDetalles
End Sub

Private Sub BrnListaEmplesas_Click()
Tipo = "Detalle de Movimientos"
FrmContListaEmpresas.Show vbModal, FrmPrincipal
BtnDesHacer_Click
End Sub

Private Sub BtnAgregar_Click()
If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub

Dim NuevoId As Integer

NewReg = 1
Blanqueo
TxtDescripcion.SetFocus
NReg.Caption = "Nuevo Registro"

CSql = "SELECT MAX(Codigo)+1 as NuevoId FROM ContDetallesMovimientos WHERE IdEmpresa=" & IdEmpresa
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = Val(RsTemp.Fields("NuevoId").Value)
Else
    NuevoId = 1
End If

TxtCodigo.Text = NuevoId
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
End Sub

Private Sub BtnAnterior_Click()
If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
If Not (RsCargarDetalles.BOF) Then
    RsCargarDetalles.MovePrevious
    If RsCargarDetalles.BOF Then RsCargarDetalles.MoveLast
    Call CargarDetalles
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
BtnAgregar.Enabled = True

If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
Form_Load
End Sub
Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""
End Sub

Private Sub BtnDetMovimientos_Click()
If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
BtnDesHacer_Click
Tipo = "Detalle de Movimientos"
FrmContListaDetallesMov.IdEmpresa = IdEmpresa
FrmContListaDetallesMov.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnEliminar_Click()

If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub

If IdDeta = "" Then
    MsgBox "Debe de seleccionar un detalle de movimientos para ser borrado!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

CSql = "Update ContDetallesMovimientos set Activo=0 Where IdDetalle='" & IdDeta & "'"
Set RsEliminar = CrearRS(CSql)

MsgBox "Elimiado del detalle de movimientos Satisfactorio!", vbCritical + vbOKOnly, "Borrado"
Form_Load
End Sub

Private Sub BtnGuardarActualizar_Click()

If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub

'***************************************
'Validación de Campos
'---------------------------------------

If TxtCodigo.Text = "" Then
    MsgBox "Esta dejando Vacio el codigo del detalle de movimiento!", vbCritical + vbOKOnly, "Error"
    TxtCodigo.SetFocus
    Exit Sub
End If

If TxtDescripcion.Text = "" Then
    MsgBox "Esta dejando Vacia la descripción del detalle de movimiento!", vbCritical + vbOKOnly, "Error"
    TxtDescripcion.SetFocus
    Exit Sub
End If

'---------------------------------------
'Guardar datos
'---------------------------------------


Select Case NewReg
    Case Is = 1
        CSql = "Select max(IdDetalle)+1 as MaxId From ContDetallesMovimientos"
        Set RsMaxId = CrearRS(CSql)
        
        If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
            IdMax = RsMaxId.Fields("MaxId").Value
        Else
            IdMax = "1"
        End If
        
        CSql = "Select * From ContDetallesMovimientos"
        Set RsDetallesMovimientos = CrearRS(CSql)
        
        RsDetallesMovimientos.AddNew
        RsDetallesMovimientos.Fields("IdDetalle").Value = IdMax
        RsDetallesMovimientos.Fields("IdEmpresa").Value = IdEmpresa
        RsDetallesMovimientos.Fields("Codigo").Value = TxtCodigo.Text
        RsDetallesMovimientos.Fields("Descripcion").Value = TxtDescripcion.Text
        RsDetallesMovimientos.Fields("IdUser").Value = IdUser
        RsDetallesMovimientos.Fields("Activo").Value = 1
        RsDetallesMovimientos.Update
        RsDetallesMovimientos.Close
        MsgBox "El registro agregado con Exito!", vbInformation + vbOKOnly, "Registro Agregado"
        
    Case Is = 2
        
        CSql = "Select * From ContDetallesMovimientos Where IdDetalle='" & IdDeta & "' AND IdEmpresa=" & IdEmpresa
        Set RsDetallesMovimientos = CrearRS(CSql)
        If RsDetallesMovimientos.RecordCount > 0 Then
            RsDetallesMovimientos.Fields("Codigo").Value = TxtCodigo.Text
            RsDetallesMovimientos.Fields("Descripcion").Value = TxtDescripcion.Text
            RsDetallesMovimientos.Fields("IdUser").Value = IdUser
            RsDetallesMovimientos.Update
        End If
        RsDetallesMovimientos.Close
        MsgBox "El registro actualizado con Exito!", vbInformation + vbOKOnly, "Registro Agregado"
    End Select
Blanqueo
Form_Load
End Sub

Private Sub BtnSiguiente_Click()

If IdEmpresa = 0 Then MsgBox "Seleccione una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
If Not (RsCargarDetalles.EOF) Then
    RsCargarDetalles.MoveNext
    If RsCargarDetalles.EOF Then RsCargarDetalles.MoveFirst
    Call CargarDetalles
End If

End Sub

Private Sub Form_Load()
BtnAgregar.Enabled = True
'NewReg = 2
If IdEmpresa <> 0 Then
    CSql = "Select * From ContDetallesMovimientos Where Activo='1' AND IdEmpresa=" & IdEmpresa
    Set RsCargarDetalles = CrearRS(CSql)
    CargarDetalles
Else
    Set RsCargarDetalles = Nothing
End If
End Sub
Sub CargarDetalles()
If Not RsCargarDetalles.EOF Then
    If IsNull(RsCargarDetalles.Fields("IdDetalle").Value) Then IdDeta = 0 Else IdDeta = Val(RsCargarDetalles.Fields("IdDetalle").Value)
    If IsNull(RsCargarDetalles.Fields("Codigo").Value) Then TxtCodigo.Text = "" Else TxtCodigo.Text = RsCargarDetalles.Fields("Codigo").Value
    If IsNull(RsCargarDetalles.Fields("Descripcion").Value) Then TxtDescripcion.Text = "" Else TxtDescripcion.Text = RsCargarDetalles.Fields("Descripcion").Value
    NewReg = 2
    BtnEliminar.Enabled = True
    NReg = "Registro " & RsCargarDetalles.AbsolutePosition & " / " & RsCargarDetalles.RecordCount
    BtnAnterior.Enabled = True
    BtnSiguiente.Enabled = True
Else
    BtnAnterior.Enabled = False
    BtnSiguiente.Enabled = False
    BtnEliminar.Enabled = False
    NewReg = 1
    NReg = "Registro 0 / 0"
End If
End Sub

VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmSimulaciones 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulaciones"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   ClipControls    =   0   'False
   Icon            =   "FrmSimulaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   4920
         Width           =   8055
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6840
            TabIndex        =   6
            ToolTipText     =   "Cerrar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "FrmSimulaciones.frx":1002
            PICN            =   "FrmSimulaciones.frx":101E
            PICH            =   "FrmSimulaciones.frx":11E7
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
            TabIndex        =   7
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmSimulaciones.frx":141C
            PICN            =   "FrmSimulaciones.frx":1438
            PICH            =   "FrmSimulaciones.frx":15C5
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
            Left            =   5640
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
            MICON           =   "FrmSimulaciones.frx":17FA
            PICN            =   "FrmSimulaciones.frx":1816
            PICH            =   "FrmSimulaciones.frx":1AF8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBorrar 
            Height          =   375
            Left            =   2520
            TabIndex        =   9
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
            MICON           =   "FrmSimulaciones.frx":1D49
            PICN            =   "FrmSimulaciones.frx":1D65
            PICH            =   "FrmSimulaciones.frx":1F09
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardar 
            Height          =   375
            Left            =   1320
            TabIndex        =   10
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
            MICON           =   "FrmSimulaciones.frx":20A8
            PICN            =   "FrmSimulaciones.frx":20C4
            PICH            =   "FrmSimulaciones.frx":2353
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8055
         Begin ChamaleonButton.ChameleonBtn BtnCapturar 
            Height          =   375
            Left            =   5880
            TabIndex        =   16
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Capturar Simulación"
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
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmSimulaciones.frx":2794
            PICN            =   "FrmSimulaciones.frx":27B0
            PICH            =   "FrmSimulaciones.frx":2A23
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox CboProtocolos 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1080
            Width           =   7815
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   735
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1800
            Width           =   7815
         End
         Begin VB.TextBox TxtCodigo 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label LblNoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro: 0 / 0"
            Height          =   195
            Left            =   2040
            TabIndex        =   14
            Top             =   330
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Protocolo de Simulación:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   810
            Width           =   1755
         End
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3625
         Object.Width           =   8025
         Object.Height          =   2025
         ScrollBar       =   4
      End
   End
End
Attribute VB_Name = "FrmSimulaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGuardar As New ADODB.Recordset
Dim RsBuscar As New ADODB.Recordset
Dim RsBorrar As New ADODB.Recordset
Dim RsCargarProtocolos As New ADODB.Recordset
Dim RsCargarListadoSimulaciones As New ADODB.Recordset
Dim IdNuevo
Dim IdProt

Private Sub BtnAgregar_Click()
LblNoReg.Caption = "Nuevo Registro"
CboProtocolos.SetFocus
Blanqueo
CSql = "Select MAX(IdSimulacion)+1 As NuevoId From Simulacion"
Set BD75 = CrearRS(CSql)

If BD75.RecordCount <> 0 Then
    If Not IsNull(BD75.Fields("NuevoId").Value) Then
        IdNuevo = BD75.Fields("NuevoId").Value
    Else
        IdNuevo = "1"
    End If
Else
    IdNuevo = "1"
End If
TxtCodigo.Text = IdNuevo
TxtDescripcion.Locked = False


End Sub

Private Sub BtnBorrar_Click()
Msg = "Estas seguro de Borrar esta simulación?"
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Mensaje")
If mensaje = vbYes Then
    CSql = "Select * From Simulacion where IdSimulacion='" & TxtCodigo.Text & "'"
    Set RsBorrar = CrearRS(CSql)

    If RsBorrar.RecordCount > 0 Then
        RsBorrar.Fields("Activo").Value = 0
        RsBorrar.Update
    End If
    Msg = "Simulación Borrada Satisfactoriamente"
    mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Mensaje")
    Blanqueo
    InitGrd
    CargarGrd
End If
End Sub

Private Sub BtnCapturar_Click()
FrmDosimetria.TxtPlanificacion.Text = Trim(TxtDescripcion.Text)
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
InitGrd
CargarGrd
End Sub
Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""
CboProtocolos.ListIndex = -1
End Sub

Private Sub BtnGuardar_Click()
'validar campos para guardar

If IdProt = -1 Then
    Msg = "Debe de Seleccionar un Protocolo"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    CboProtocolos.SetFocus
    Exit Sub
End If

If TxtDescripcion.Text = "" Then
    Msg = "Debe de ingresar "
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    CboProtocolos.SetFocus
    Exit Sub
End If

'guarda el registro

CSql = "Select * From Simulacion"
Set RsGuardar = CrearRS(CSql)

RsGuardar.AddNew
RsGuardar.Fields("IdSimulacion").Value = TxtCodigo.Text
RsGuardar.Fields("IdProtocolo").Value = IdProt
RsGuardar.Fields("Descripcion").Value = Trim(TxtDescripcion.Text)
RsGuardar.Fields("IdUser").Value = IdUser
RsGuardar.Fields("Activo").Value = 1

RsGuardar.Update

Msg = "Simulación Guardada Satifactoriamente"
MsgBox Msg, vbOKOnly + vbInformation, "Registro Guardado"

Blanqueo
InitGrd
CargarGrd

End Sub

Private Sub CboProtocolos_Change()
If CboProtocolos.ItemData(CboProtocolos.ListIndex) = -1 Then Exit Sub
End Sub

Private Sub CboProtocolos_Click()
If CboProtocolos.ListIndex = -1 Then
    Exit Sub
Else
    IdProt = CboProtocolos.ItemData(CboProtocolos.ListIndex)
End If
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton Then
    CSql = "Select * From Simulacion Where IdSimulacion='" & DMGrid1.ValorCelda(lRow, 1) & "' And Activo='1'"
    Set RsBuscar = CrearRS(CSql)
    
    TxtCodigo.Text = RsBuscar.Fields("IdSimulacion").Value
    TxtDescripcion.Text = RsBuscar.Fields("Descripcion").Value
    
    For i = 1 To CboProtocolos.ListCount - 1
        If CboProtocolos.ItemData(i) = RsBuscar.Fields("IdProtocolo").Value Then
            CboProtocolos.ListIndex = i
            Exit For
        End If
    Next i
End If
End Sub

Private Sub Form_Activate()
CboProtocolos.SetFocus
End Sub

Private Sub Form_Load()

CSql = "Select * From Protocolos"
Set RsCargarProtocolos = CrearRS(CSql)

Do While Not RsCargarProtocolos.EOF
    'With CboProtocolos
        CboProtocolos.AddItem RsCargarProtocolos.Fields("Protocolo").Value
        CboProtocolos.ItemData(CboProtocolos.NewIndex) = RsCargarProtocolos.Fields("Id").Value
   ' End With
    RsCargarProtocolos.MoveNext
Loop

InitGrd
CargarGrd
End Sub

Sub InitGrd()

DMGrid1.Cols = 2
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 85 / 100)
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Descripción"


End Sub

Sub CargarGrd()

CSql = "Select * From Simulacion WHERE Activo='1' order by IdSimulacion"
Set RsCargarListadoSimulaciones = CrearRS(CSql)

Do While Not RsCargarListadoSimulaciones.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListadoSimulaciones.Fields("IdSimulacion").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListadoSimulaciones.Fields("Descripcion").Value
    RsCargarListadoSimulaciones.MoveNext
Loop
DMGrid1.PaintMGrid
End Sub

VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTecnicos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tecnicos"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "FrmTecnicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   7695
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Editar"
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
         MICON           =   "FrmTecnicos.frx":1002
         PICN            =   "FrmTecnicos.frx":101E
         PICH            =   "FrmTecnicos.frx":12AD
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
         TabIndex        =   4
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
         MICON           =   "FrmTecnicos.frx":16EE
         PICN            =   "FrmTecnicos.frx":170A
         PICH            =   "FrmTecnicos.frx":1897
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
         TabIndex        =   6
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "FrmTecnicos.frx":1ACC
         PICN            =   "FrmTecnicos.frx":1AE8
         PICH            =   "FrmTecnicos.frx":1C8C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   6600
         TabIndex        =   10
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
         MICON           =   "FrmTecnicos.frx":1E2B
         PICN            =   "FrmTecnicos.frx":1E47
         PICH            =   "FrmTecnicos.frx":2010
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
         Left            =   5400
         TabIndex        =   9
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
         MICON           =   "FrmTecnicos.frx":2245
         PICN            =   "FrmTecnicos.frx":2261
         PICH            =   "FrmTecnicos.frx":2543
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
         TabIndex        =   8
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
         MICON           =   "FrmTecnicos.frx":2794
         PICN            =   "FrmTecnicos.frx":27B0
         PICH            =   "FrmTecnicos.frx":2A46
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
         Left            =   3840
         TabIndex        =   7
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
         MICON           =   "FrmTecnicos.frx":2CA5
         PICN            =   "FrmTecnicos.frx":2CC1
         PICH            =   "FrmTecnicos.frx":2F56
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
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   4920
         Top             =   240
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   1215
         Left            =   5880
         TabIndex        =   15
         Top             =   480
         Width           =   1695
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Razon Social, Cédula o Rif"
            Top             =   240
            Width           =   1455
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Busqueda"
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
            MICON           =   "FrmTecnicos.frx":31B2
            PICN            =   "FrmTecnicos.frx":31CE
            PICH            =   "FrmTecnicos.frx":3433
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
      Begin VB.TextBox TxtApellidos 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox TxtNombres 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox TxtCedula 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label NroReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Registro: 0 / 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   18
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellido(s):"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre(s):"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   330
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmTecnicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewReg As String
Dim NewId As String
Dim IdTecnicos As String
Dim RsTecnicos As New ADODB.Recordset
Dim RsNewId As New ADODB.Recordset
Private Sub BtnAgregar_Click()
NewReg = 0
Blanqueo
TxtCedula.SetFocus
BtnGuardarActualizar.Caption = "Guardar"
Me.Caption = "Tecnicos"
End Sub

Private Sub BtnAnterior_Click()
If RsTecnicos.RecordCount <> 0 Then
    RsTecnicos.MovePrevious
    If RsTecnicos.BOF Then RsTecnicos.MoveLast
    Call CargaTecnicos
End If
End Sub

Private Sub BtnBuscar_Click()
Blanqueo
If Trim(TxtBuscar.Text) = "" Or UCase(TxtBuscar.Text) = UCase("Busqueda") Then
    'f = "Buscar"
    CSql = "Select * From Tecnicos"
Else
    CSql = "Select * From Tecnicos where Nombre='" & Trim(TxtBuscar.Text) & "' or Apellido ='" & Trim(TxtBuscar.Text) & "' or cedula='" & Trim(TxtBuscar.Text) & "'"
End If

Set RsTecnicos = CrearRS(CSql)
CargaTecnicos

BtnGuardarActualizar.Enabled = True
BtnGuardarActualizar.Caption = "Editar"

BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
End Sub

Sub CargaTecnicos()

If RsTecnicos.RecordCount > 0 Then
    NewReg = 1
    If Trim(RsTecnicos.Fields("cedula").Value) <> "" Then TxtCedula.Text = RsTecnicos.Fields("cedula").Value
    If Trim(RsTecnicos.Fields("Apellido").Value) <> "" Then TxtApellidos.Text = RsTecnicos.Fields("Apellido").Value
    If Trim(RsTecnicos.Fields("Nombre").Value) <> "" Then TxtNombres.Text = RsTecnicos.Fields("Nombre").Value
         
    IdTecnicos = RsTecnicos.Fields("IdTecnicos").Value
              
    Me.Caption = "Tecnicos - Id: " & IdTecnicos
    NroReg.Caption = "Nro de Registro: " & RsTecnicos.AbsolutePosition & " / " & RsTecnicos.RecordCount
    'BtnBorrar.Enabled = True
Else
    IdTecnicos = 0
    NewReg = 0
    NroReg.Caption = "Nro de Registro: 0 / 0"
    BtnBorrar.Enabled = False
End If

End Sub


Private Sub BtnGuardarActualizar_Click()
'validar
If TxtCedula.Text = "" Then
    Msg = "Ingrese la cedula del Tecnico"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

If TxtApellidos.Text = "" Then
    Msg = "Ingrese el Apellido del Tecnico"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

If TxtNombres.Text = "" Then
    Msg = "Ingrese el Nombre del Tecnico"
    MsgBox Msg, vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

'acciones
Dim RsGuardar As New ADODB.Recordset
Dim RsActualizar As New ADODB.Recordset

Select Case NewReg
    Case Is = 0
        CSql = "Select MAX(IdTecnicos)+1 As NuevoId From Tecnicos"
        Set RsNewId = CrearRS(CSql)
        
        If RsNewId.RecordCount <> 0 Then
            If Not IsNull(RsNewId.Fields("NuevoId").Value) Then
                C = RsNewId.Fields("NuevoId").Value
            Else
                C = "1"
            End If
        Else
            C = "1"
        End If
        
        CSql = "Select * From Tecnicos"
        Set RsGuardar = CrearRS(CSql)
        
        RsGuardar.AddNew
        RsGuardar.Fields("IdTecnicos").Value = C
        RsGuardar.Fields("Cedula").Value = Trim(TxtCedula.Text)
        RsGuardar.Fields("Nombre").Value = Trim(TxtNombres.Text)
        RsGuardar.Fields("Apellido").Value = Trim(TxtApellidos.Text)
        RsGuardar.Fields("IdUser").Value = IdUser
        RsGuardar.Fields("Activo").Value = 1
        RsGuardar.Update
        
        Msg = "Registro Guardado con exito!!!"
        MsgBox Msg, vbOKOnly + vbInformation, "Operación Exitosa"
        
    Case Is = 1
        CSql = "Select * From Tecnicos Where IdTecnicos ='" & IdTecnicos & "'"
        Set RsActualizar = CrearRS(CSql)
        
        RsActualizar.Fields("Cedula").Value = Trim(TxtCedula.Text)
        RsActualizar.Fields("Nombre").Value = Trim(TxtNombres.Text)
        RsActualizar.Fields("Apellido").Value = Trim(TxtApellidos.Text)
        RsActualizar.Fields("IdUser").Value = IdUser
        RsActualizar.Fields("Activo").Value = 1
        RsActualizar.Update
        
        Msg = "Registro Actualizado con exito!!!"
        MsgBox Msg, vbOKOnly + vbInformation, "Operación Exitosa"
        
End Select
Blanqueo
Form_Load
End Sub

Private Sub BtnSiguiente_Click()
If RsTecnicos.RecordCount <> 0 Then
    RsTecnicos.MoveNext
    If RsTecnicos.EOF Then RsTecnicos.MoveFirst
    Call CargaTecnicos
End If
End Sub

Private Sub Form_Load()

CSql = "Select * From Tecnicos"
Set RsTecnicos = CrearRS(CSql)

End Sub
Sub Blanqueo()
TxtCedula.Text = ""
TxtApellidos.Text = ""
TxtNombres.Text = ""
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtCedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtApellidos.SetFocus
    Else
        If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            MsgBox "El caracter digitado no es válido.", vbExclamation, "Atención"
            KeyAscii = 0
        End If
    End If
End Sub

'///////////////////////////////////Valido TextBox: txtapellidos//////////////////////////////
Private Sub txtapellidos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtNombres.SetFocus
    Else
        If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            MsgBox "El caracter digitado no es válido.", vbExclamation, "Atención"
            KeyAscii = 0
        End If
    End If
        
End Sub

Private Sub TxtNombres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BtnGuardarActualizar.SetFocus
    Else
        If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
            MsgBox "El caracter digitado no es válido.", vbExclamation, "Atención"
            KeyAscii = 0
        End If
    End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContEmpresas 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Empresas"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "FrmContEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   5295
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   4440
         Width           =   6735
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   5640
            TabIndex        =   13
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
            MICON           =   "FrmContEmpresas.frx":1002
            PICN            =   "FrmContEmpresas.frx":101E
            PICH            =   "FrmContEmpresas.frx":11E7
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
            MICON           =   "FrmContEmpresas.frx":141C
            PICN            =   "FrmContEmpresas.frx":1438
            PICH            =   "FrmContEmpresas.frx":16C7
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
            TabIndex        =   9
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
            MICON           =   "FrmContEmpresas.frx":1B08
            PICN            =   "FrmContEmpresas.frx":1B24
            PICH            =   "FrmContEmpresas.frx":1CB1
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
            Left            =   4440
            TabIndex        =   12
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
            MICON           =   "FrmContEmpresas.frx":1EE6
            PICN            =   "FrmContEmpresas.frx":1F02
            PICH            =   "FrmContEmpresas.frx":21E4
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
            Left            =   2520
            TabIndex        =   11
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Borrar"
            ENAB            =   0   'False
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
            MICON           =   "FrmContEmpresas.frx":2435
            PICN            =   "FrmContEmpresas.frx":2451
            PICH            =   "FrmContEmpresas.frx":25F5
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
         Caption         =   "Datos de la Empresa"
         Height          =   4215
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6735
         Begin VB.TextBox TxtClave 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   3120
            Width           =   3255
         End
         Begin VB.TextBox TxtCiudad 
            Height          =   375
            Left            =   1320
            TabIndex        =   6
            Top             =   2640
            Width           =   3255
         End
         Begin VB.CheckBox ChkConsolidadora 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Consolidadora"
            Height          =   255
            Left            =   4920
            TabIndex        =   8
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox TxtDireccion 
            Height          =   735
            Left            =   1320
            TabIndex        =   3
            Top             =   1320
            Width           =   5175
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   1320
            TabIndex        =   1
            Top             =   840
            Width           =   3255
         End
         Begin VB.ComboBox CboCodigo 
            Height          =   315
            ItemData        =   "FrmContEmpresas.frx":2794
            Left            =   1320
            List            =   "FrmContEmpresas.frx":2796
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2190
            Width           =   975
         End
         Begin VB.TextBox TxtRif 
            Height          =   375
            Left            =   1320
            TabIndex        =   0
            Top             =   360
            Width           =   3255
         End
         Begin ChamaleonButton.ChameleonBtn BrnListaEmpleados 
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Top             =   3720
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
            MICON           =   "FrmContEmpresas.frx":2798
            PICN            =   "FrmContEmpresas.frx":27B4
            PICH            =   "FrmContEmpresas.frx":2A3D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   120
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4920
            TabIndex        =   2
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   55115779
            CurrentDate     =   39932
         End
         Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
            Height          =   375
            Left            =   5640
            TabIndex        =   15
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   3720
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
            MICON           =   "FrmContEmpresas.frx":2E58
            PICN            =   "FrmContEmpresas.frx":2E74
            PICH            =   "FrmContEmpresas.frx":310A
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
            Left            =   4920
            TabIndex        =   14
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   3720
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
            MICON           =   "FrmContEmpresas.frx":3369
            PICN            =   "FrmContEmpresas.frx":3385
            PICH            =   "FrmContEmpresas.frx":361A
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
            Caption         =   "Clave:"
            Height          =   195
            Left            =   720
            TabIndex        =   27
            Top             =   3210
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Ciudad:"
            Height          =   195
            Left            =   600
            TabIndex        =   26
            Top             =   2730
            Width           =   645
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Ingreso:"
            Height          =   195
            Left            =   4920
            TabIndex        =   25
            Top             =   480
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Dirección:"
            Height          =   195
            Left            =   360
            TabIndex        =   23
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Nombre(s):"
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   930
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono:"
            Height          =   195
            Left            =   480
            TabIndex        =   21
            Top             =   2250
            Width           =   675
         End
         Begin VB.Label NoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   5040
            TabIndex        =   20
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Rif:"
            Height          =   195
            Left            =   840
            TabIndex        =   19
            Top             =   480
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "FrmContEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEmpresas As Recordset
Dim RsTemp As Recordset
Dim RegNew As Boolean
Dim IdEmpresa As Integer
Dim i As Integer


Sub Blanqueo()
TxtRif.Text = ""
TxtNombre.Text = ""
TxtDireccion.Text = ""
TxtTelefono.Text = ""
TxtCiudad.Text = ""
TxtClave.Text = ""
DTPicker1.Value = Now
CboCodigo.ListIndex = -1
ChkConsolidadora.Value = 0
Reg_Actual(3) = ""
End Sub

Sub Cargar_Empresa()

If RsEmpresas.RecordCount = 0 Then
    RegNew = True
    IdEmpresa = 0
    BtnEliminar.Enabled = False
    Exit Sub
End If

BtnEliminar.Enabled = True
IdEmpresa = Val(RsEmpresas.Fields("IdEmpresa").Value)
TxtRif.Text = RsEmpresas.Fields("Rif").Value
TxtNombre.Text = RsEmpresas.Fields("Nombre").Value
TxtDireccion.Text = RsEmpresas.Fields("Direccion").Value
TxtTelefono.Text = RsEmpresas.Fields("Telefono").Value
TxtCiudad.Text = RsEmpresas.Fields("Ciudad").Value
DTPicker1.Value = Format(RsEmpresas.Fields("FechaIngreso").Value, "dd/MM/yyyy")

For i = 0 To CboCodigo.ListCount - 1
    If RsTemp.Fields("CodigoTelf").Value = CboCodigo.List(i) Then
        CboCodigo.ListIndex = i
        Exit For
    Else
        CboCodigo.ListIndex = -1
    End If
Next i

If RsEmpresas.Fields("Consolidadora").Value Then ChkConsolidadora.Value = 1 Else ChkConsolidadora.Value = 0

Reg_Actual(3) = RsEmpresas.Fields("Rif").Value
NoReg.Caption = "Registro " & RsEmpresas.AbsolutePosition & " / " & RsEmpresas.RecordCount
RegNew = False

End Sub

Private Sub BrnListaEmpleados_Click()
Tipo = "Empresa"
BtnDesHacer_Click
FrmContListaEmpresas.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAgregar_Click()
Blanqueo
RegNew = True
NoReg.Caption = "Nuevo Registro"
BtnEliminar.Enabled = False
BtnAgregar.Enabled = False
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim resp As Byte
Dim NuevoId As Integer
Dim CodTelf As Integer

If Trim(TxtRif) = "" Then
    MsgBox "Debe ingresar el RIF de la Empresa!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtRif.SetFocus
    Exit Sub
ElseIf Trim(TxtNombre) = "" Then
    MsgBox "Debe ingresar el NOMBRE de la Empresa!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtNombre.SetFocus
    Exit Sub
ElseIf Trim(TxtDireccion) = "" Then
    MsgBox "Debe ingresar la DIRECCIÓN de la Empresa!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtDireccion.SetFocus
    Exit Sub
'ElseIf Trim(TxtTelefono) = "" Then
'    MsgBox "Debe ingresar el Telefono de la Empresa!", vbExclamation + vbOKOnly, "Faltan Datos!"
'    TxtTelefono.SetFocus
'    Exit Sub
ElseIf Trim(TxtCiudad) = "" Then
    MsgBox "Debe ingresar la CIUDAD de ubicación de la Empresa!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtCiudad.SetFocus
    Exit Sub
End If

resp = MsgBox("Se procedera a guardar los cambios, Desea continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = vbNo Then Exit Sub


CSql = "SELECT MAX(IdEmpresa)+1 As NuevoId FROM ContEmpresas"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If

If CboCodigo.ListIndex = -1 Then
    CodTelf = 0
Else
    CodTelf = CboCodigo.ItemData(CboCodigo.ListIndex)
End If

If RegNew Then
    CSql = "INSERT INTO ContEmpresas (IdEmpresa,Nombre,Rif,Direccion,Ciudad,CodigoTelf,Telefono,Clave," & _
    " Consolidadora,FechaIngreso,Activo) VALUES(" & NuevoId & ",N'" & Trim(TxtNombre.Text) & "',N'" & Trim(TxtRif.Text) & _
    "',N'" & Trim(TxtDireccion.Text) & "',N'" & Trim(TxtCiudad.Text) & "','" & CodTelf & "','" & Trim(TxtTelefono.Text) & _
    "',N'" & Trim(TxtClave.Text) & "'," & ChkConsolidadora.Value & ",'" & Format(DTPicker1.Value, "dd/MM/yyyy") & "','1')"
    Set RsTemp = CrearRS(CSql)
Else
    CSql = "UPDATE ContEmpresas SET Nombre='" & Trim(TxtNombre.Text) & "', Rif='" & Trim(TxtRif.Text) & "'," & _
    " Direccion='" & Trim(TxtDireccion.Text) & "',Ciudad='" & Trim(TxtCiudad.Text) & "',CodigoTelf='" & CodTelf & "'," & _
    " Telefono='" & Trim(TxtTelefono.Text) & "',Clave='" & Trim(TxtClave.Text) & "',Consolidadora=" & ChkConsolidadora.Value & _
    ",FechaIngreso='" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' WHERE IdEmpresa=" & IdEmpresa
    Set RsTemp = CrearRS(CSql)
End If

MsgBox "Los datos han sido guardados!", vbInformation + vbOKOnly, "Operación Exitosa!"

Form_Load

End Sub

Private Sub BtnSiguiente_Click()
If BtnAgregar.Enabled = False Then BtnDesHacer_Click
If RsEmpresas.RecordCount <> 0 Then
    Blanqueo
    RsEmpresas.MoveNext
    If RsEmpresas.EOF Then RsEmpresas.MoveFirst
    Call Cargar_Empresa
Else
    MsgBox "No hay registros cargados!", vbExclamation + vbOKOnly, "No hay registros!"
End If
End Sub

Private Sub BtnAnterior_Click()
If BtnAgregar.Enabled = False Then BtnDesHacer_Click
If RsEmpresas.RecordCount <> 0 Then
    Blanqueo
    RsEmpresas.MovePrevious
    If RsEmpresas.BOF Then RsEmpresas.MoveLast
    Call Cargar_Empresa
Else
    MsgBox "No hay registros cargados!", vbExclamation + vbOKOnly, "No hay registros!"
End If
End Sub

Private Sub BtnDesHacer_Click()
Form_Load
BtnAgregar.Enabled = True
End Sub

Private Sub BtnEliminar_Click()
Dim resp As Byte

resp = MsgBox("Esta seguro de Eliminar la empresa?", vbQuestion + vbYesNo, "Confirmar")
If resp = vbNo Then Exit Sub

CSql = "UPDATE ContEmpresas SET Activo=0 WHERE IdEmpresa=" & IdEmpresa
Set RsTemp = CrearRS(CSql)

MsgBox "El registro ha sido eliminado!", vbInformation + vbOKOnly, "Operación Exitosa!"

Form_Load
End Sub

Private Sub Form_Load()

Centrar Me
Blanqueo

RegNew = False
CSql = "SELECT * FROM ContEmpresas WHERE Activo='1'"
Set RsEmpresas = CrearRS(CSql)

If RsEmpresas.RecordCount <> 0 Then
    RsEmpresas.MoveFirst
    Cargar_Empresa
Else
    IdEmpresa = 0
    RegNew = True
End If

End Sub

Private Sub TxtCiudad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtClave.SetFocus: KeyAscii = 0: Exit Sub
If Len(TxtCiudad.Text) > 50 Then MsgBox "El nombre de la Ciudad no debe tener más de 50 caracteres!", vbCritical + vbOKOnly, "Información": KeyAscii = 0
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnGuardarActualizar.SetFocus: KeyAscii = 0: Exit Sub
If Len(TxtClave.Text) > 50 Then MsgBox "No se permiten claves con más de 50 caracteres!", vbCritical + vbOKOnly, "Información": KeyAscii = 0
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtCiudad.SetFocus: KeyAscii = 0: Exit Sub
If Len(TxtTelefono.Text) > 10 Then MsgBox "El numero de telefono no debe tener más de 10 caracteres!", vbCritical + vbOKOnly, "Información": KeyAscii = 0
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtTelefono.SetFocus: KeyAscii = 0: Exit Sub
If Len(TxtDireccion.Text) > 254 Then MsgBox "La Dirección no debe tener más de 254 caracteres!", vbCritical + vbOKOnly, "Información": KeyAscii = 0
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtDireccion.SetFocus: KeyAscii = 0: Exit Sub
If Len(TxtNombre.Text) > 50 Then MsgBox "El Nombre no debe tener más de 50 caracteres!", vbCritical + vbOKOnly, "Información": KeyAscii = 0
End Sub

Private Sub TxtRif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtNombre.SetFocus: KeyAscii = 0: TxtRif_LostFocus: Exit Sub
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> vbKeyJ _
    And KeyAscii <> vbKeyV And KeyAscii <> vbKeyG And KeyAscii <> vbKeyE Then KeyAscii = 0
If Len(TxtRif.Text) > 15 And KeyAscii <> 8 Then MsgBox "El Rif no debe tener más de 15 caracteres!", vbCritical + vbOKOnly, "Información": KeyAscii = 0
End Sub

Private Sub TxtRif_LostFocus()
Dim IdRif

IdRif = Trim(TxtRif.Text)

If IdRif = 0 Then Exit Sub

CSql = "SELECT * FROM ContEmpresas WHERE RIF='" & IdRif & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If RegNew Then
        MsgBox "El RIF ingresado ya se encuentra registrado!", vbCritical + vbOKOnly, "Error"
        TxtRif.Text = ""
        TxtRif.SetFocus
        Exit Sub
    Else
        If Trim(Reg_Actual(3)) <> Trim(TxtRif.Text) Then
            MsgBox "El RIF ingresado ya se encuentra registrado!", vbCritical + vbOKOnly, "Error"
            TxtRif.Text = Reg_Actual(3)
            TxtRif.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub


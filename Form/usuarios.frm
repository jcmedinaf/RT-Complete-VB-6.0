VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmUsuarios 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso de Seguridad al Sistema"
   ClientHeight    =   7875
   ClientLeft      =   6555
   ClientTop       =   675
   ClientWidth     =   6570
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000F&
   Icon            =   "usuarios.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   6570
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   7815
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   6375
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   6840
         Width           =   6135
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   5040
            TabIndex        =   13
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
            MICON           =   "usuarios.frx":1002
            PICN            =   "usuarios.frx":101E
            PICH            =   "usuarios.frx":11E7
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
            Height          =   495
            Left            =   1200
            TabIndex        =   10
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":141C
            PICN            =   "usuarios.frx":1438
            PICH            =   "usuarios.frx":16C7
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
            Height          =   495
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":1B08
            PICN            =   "usuarios.frx":1B24
            PICH            =   "usuarios.frx":1CB1
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
            Height          =   495
            Left            =   3840
            TabIndex        =   12
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":1EE6
            PICN            =   "usuarios.frx":1F02
            PICH            =   "usuarios.frx":21E4
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
            Height          =   495
            Left            =   2400
            TabIndex        =   11
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":2435
            PICN            =   "usuarios.frx":2451
            PICH            =   "usuarios.frx":25F5
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Cuentas Asociadas"
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   5160
         Width           =   5895
         Begin VB.ComboBox CboMedicoAsociado 
            Height          =   315
            ItemData        =   "usuarios.frx":2794
            Left            =   1800
            List            =   "usuarios.frx":2796
            TabIndex        =   8
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label13 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Médico asociado a la cuenta:"
            Height          =   435
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos de Usuarios"
         Height          =   5775
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6135
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3720
            Top             =   1680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   3240
            Top             =   1680
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nivel de Acceso"
            Height          =   855
            Left            =   120
            TabIndex        =   29
            Top             =   3960
            Width           =   5895
            Begin VB.ComboBox CboAccesoSistema 
               Height          =   315
               Left            =   1800
               TabIndex        =   7
               Top             =   360
               Width           =   3255
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Acceso al Sistemas:"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   1425
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Clave y Nivel de Acceso"
            Height          =   1695
            Left            =   120
            TabIndex        =   25
            Top             =   2160
            Width           =   5895
            Begin VB.TextBox TxtUsuario 
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   1800
               TabIndex        =   4
               Top             =   240
               Width           =   3255
            End
            Begin VB.TextBox TxtClave1 
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   1800
               PasswordChar    =   "*"
               TabIndex        =   5
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox TxtClave2 
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   1800
               PasswordChar    =   "*"
               TabIndex        =   6
               Top             =   1200
               Width           =   3255
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Usuario:"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   330
               Width           =   585
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ingrese Contraseña:"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Repita  Contraseña:"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   1290
               Width           =   1410
            End
         End
         Begin VB.TextBox TxtApellidos 
            Height          =   375
            Left            =   1680
            TabIndex        =   0
            Top             =   270
            Width           =   2535
         End
         Begin VB.TextBox TxtCedula 
            Height          =   375
            Left            =   1680
            TabIndex        =   2
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox TxtNombres 
            Height          =   375
            Left            =   1680
            TabIndex        =   1
            Top             =   720
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DtpFechaRegistro 
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   52953091
            CurrentDate     =   40226
         End
         Begin VB.Label NoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   4320
            TabIndex        =   34
            Top             =   1920
            Width           =   975
         End
         Begin VB.Image Image3 
            BorderStyle     =   1  'Fixed Single
            Height          =   1575
            Left            =   4320
            Picture         =   "usuarios.frx":2798
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula de Identidad:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   1470
         End
         Begin VB.Label Label7 
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Registro:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1740
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   6000
         Width           =   6135
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   495
            Left            =   2520
            TabIndex        =   14
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":5B19
            PICN            =   "usuarios.frx":5B35
            PICH            =   "usuarios.frx":5D9A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
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
            TabIndex        =   32
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad"
            Top             =   300
            Width           =   2295
         End
         Begin ChamaleonButton.ChameleonBtn BtnAnterior 
            Height          =   495
            Left            =   4680
            TabIndex        =   15
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":602C
            PICN            =   "usuarios.frx":6048
            PICH            =   "usuarios.frx":62DD
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
            Height          =   495
            Left            =   5400
            TabIndex        =   16
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            MICON           =   "usuarios.frx":6539
            PICN            =   "usuarios.frx":6555
            PICH            =   "usuarios.frx":67EB
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
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UsuaroCount As Variant '//Mantiene número de registros
Dim BD As New ADODB.Recordset
Dim BD2 As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim SentenciaSQL As String
Dim FotoE As String
Dim idUsu
Dim Actuali
Public RsUsuarios As New ADODB.Recordset
Sub carga_lista_medicost()

Dim bd1 As New ADODB.Recordset
CSql = "SELECT * FROM medicos_t"

bd1.Open CSql, Cnn
bd1.MoveFirst

Do While Not bd1.EOF
    CboMedicoAsociado.AddItem bd1.Fields("Nombre")
    CboMedicoAsociado.ItemData(CboMedicoAsociado.NewIndex) = bd1.Fields("idmedicot")
    bd1.MoveNext
Loop
End Sub

Private Sub BtnAgregar_Click()
Blanqueo
Actuali = 0
TxtApellidos.SetFocus
BtnAgregar.Enabled = False
BtnGuardar.Enabled = True
BtnEliminar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
BtnBuscar.Enabled = False
End Sub

Private Sub BtnAnterior_Click()
If RsUsuarios.RecordCount <> 0 Then
    Blanqueo
    RsUsuarios.MovePrevious
    If RsUsuarios.BOF Then RsUsuarios.MoveLast
    Call Carga_De_Datos
Else
    MsgBox "No hay registros!", vbExclamation + vbOKOnly, "No hay registros!"
End If
End Sub

Private Sub BtnBuscar_Click()
If TxtBuscar.Text <> "" Then
    CSql = "SELECT * FROM Usuarios Where Usuario like '%" & TxtBuscar.Text & "%' OR " & _
    "Apellidos like '%" & TxtBuscar.Text & "%' OR Nombre like '%" & TxtBuscar.Text & _
    "%' OR cedula like '%" & TxtBuscar.Text & "%'"
    
    Set rs = CrearRS(CSql)
    If rs.RecordCount <> 0 Then
        TxtCedula.Text = rs.Fields("Cedula")
        TxtApellidos.Text = rs.Fields("Apellidos")
        TxtNombres.Text = rs.Fields("Nombre")
        TxtUsuario.Text = rs.Fields("Usuario")
        TxtClave1.Text = rs.Fields("Contraseña")
        TxtClave2.Text = rs.Fields("Contraseña")
        DtpFechaRegistro.Value = rs.Fields("Fecha")
        CboAccesoSistema.ListIndex = rs.Fields("T_U")
        idUsu = rs.Fields("IdUsuario")
        Actuali = 1
        If Not IsNull(rs.Fields("idmedicot")) Then
            If rs.Fields("idmedicot") = -1 Then
               CboMedicoAsociado.ListIndex = -1
            Else
                For T = 0 To CboMedicoAsociado.ListCount - 1
                  If rs.Fields("idmedicot") = CboMedicoAsociado.ItemData(T) Then
                     CboMedicoAsociado.ListIndex = T
                     Exit For
                  End If
                Next T
            End If
        Else
            CboMedicoAsociado.ListIndex = -1
        End If
    End If
Else
    Exit Sub
End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Form_Load
End Sub

Private Sub BtnEliminar_Click()
Msg = "Esta Seguro de Eliminar este Usuario?" & Chr(13) & Chr(13) & TxtApellidos.Text & ",  " & TxtNombres.Text
p = MsgBox(Msg, vbYesNo, "Eliminar Usuario")

If p = vbYes Then

    CSql = "Select IdUsuario From Usuarios Where Cedula = '" & Trim(TxtCedula.Text) & "'"
    Set rs = CrearRS(CSql)
    
    IdUsers = rs.Fields(0).Value
    
    CSql = "DELETE FROM Usuarios Where IdUsuario = " & IdUsers
    Set rs = CrearRS(CSql)
    MsgBox "Fue eliminado el registro", vbInformation + vbOKOnly, "Usuario Eliminado"
    Blanqueo
End If

End Sub

Private Sub BtnGuardar_Click()
Dim Cantidad As Integer

If Trim(TxtCedula.Text) = "" Then
    MsgBox "El Campo de Cédula es Requerida!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Exit Sub
ElseIf Trim(TxtNombres.Text) = "" Then
    MsgBox "El Campo de Nombres es Requerida!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Exit Sub
ElseIf Trim(TxtApellidos.Text) = "" Then
    MsgBox "El Campo de Apellidos es Requerida!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Exit Sub
ElseIf Trim(TxtUsuario.Text) = "" Then
    MsgBox "El Campo de Usuario es Requerida!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Exit Sub
ElseIf Trim(TxtClave1.Text) = "" Then
    MsgBox "El Campo de Clave es Requerida!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Exit Sub
ElseIf CboAccesoSistema.ListIndex = -1 Then
    MsgBox "El Nivel de Acceso al Sistema Requerido!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Exit Sub
End If

If CboMedicoAsociado.ListIndex > 0 Then DRT = CboMedicoAsociado.ItemData(CboMedicoAsociado.ListIndex) Else DRT = -1
Select Case Actuali
Case Is = 0

Fecha = Format(DtpFechaRegistro.Value, "dd/mm/yyyy")

CSql = "SELECT MAX(idusuario)+1 as NuevoId FROM Usuarios"
Set rs = CrearRS(CSql)

If rs.RecordCount <> 0 Then
    If rs.RecordCount <> 0 Then
        Cantidad = rs.Fields(0).Value
    Else
        Cantidad = 1
    End If
Else
    Cantidad = 1
End If

Cantidad = Cantidad + 1
CSql = "SELECT cedula From Usuarios Where cedula = '" & Trim(TxtCedula.Text) & "'"
Set rs = CrearRS(CSql)
    
If rs.EOF = True Or rs.BOF = True Then
    CSql = "Insert into Usuarios(IdUsuario, Cedula, Nombre, Apellidos, Usuario, Contraseña, T_U, Fecha, idmedicot) " & _
        " VALUES(" & Cantidad & ",'" & Trim(TxtCedula.Text) & "', '" & Trim(TxtNombres.Text) & "','" & _
        Trim(TxtApellidos.Text) & "','" & Trim(TxtUsuario.Text) & "','" & Trim(TxtClave1.Text) & "','" & _
        CboAccesoSistema.ItemData(CboAccesoSistema.ListIndex) & "','" & CDate(Fecha) & "','" & DRT & "')"
    Set RstResultado = CrearRS(CSql)
    MsgBox "Registro Agregado satisfactoriamente", vbInformation + vbOKOnly, "Operación Exitosa."
    Blanqueo
End If
    
Case Is = 1

CSql = "Select Idusuario From Usuarios Where Cedula = '" & Trim(TxtCedula.Text) & "'"
Set rs = CrearRS(CSql)

IdUsers = rs.Fields(0).Value
Fecha = Format(DtpFechaRegistro.Value, "dd/mm/yyyy")
    CSql = "UPDATE Usuarios SET IdUsuario = " & IdUsers & ", CEDULA = '" & Trim(TxtCedula.Text) & _
    "', NOMBRE = '" & Trim(TxtNombres.Text) & "', Apellidos = '" & Trim(TxtApellidos.Text) & _
    "', Usuario = '" & Trim(TxtUsuario.Text) & "', Contraseña = '" & Trim(TxtClave1.Text) & _
    "', T_U = '" & CboAccesoSistema.ItemData(CboAccesoSistema.ListIndex) & "', Fecha = '" & CDate(Fecha) & _
    "', idmedicot = '" & DRT & "', Foto='" & FotoE & "' Where IdUsuario = " & IdUsers
    Set RstResultado = CrearRS(CSql)
    MsgBox "Registro Actualizado satisfactoriamente", vbInformation + vbOKOnly, "Operación Exitosa."
    Blanqueo
End Select
Form_Load
End Sub

Private Sub BtnSiguiente_Click()
If RsUsuarios.RecordCount <> 0 Then
    Blanqueo
    RsUsuarios.MoveNext
    If RsUsuarios.EOF Then RsUsuarios.MoveFirst
    Call Carga_De_Datos
Else
    MsgBox "No hay registros!", vbExclamation + vbOKOnly, "No hay registros!"
End If
End Sub

Private Sub CboAccesoSistema_Change()
Cambio = 0
End Sub

Sub Blanqueo()

TxtApellidos.Text = ""
TxtNombres.Text = ""
TxtUsuario.Text = ""
TxtCedula.Text = ""
TxtClave1.Text = ""
TxtClave2.Text = ""
CboAccesoSistema.ListIndex = -1
CboMedicoAsociado.ListIndex = -1
DtpFechaRegistro.Value = Now()
            
End Sub

Private Sub CboAccesoSistema_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboMedicoAsociado.SetFocus
End If
End Sub

Private Sub DtpFechaRegistro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtUsuario.SetFocus
End If
End Sub


Private Sub Form_Load()
Centrar Me

CboAccesoSistema.Clear
CboMedicoAsociado.Clear
DtpFechaRegistro.Value = Now

CboAccesoSistema.AddItem "Acceso_t"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 0
CboAccesoSistema.AddItem "Usuarios"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 1
CboAccesoSistema.AddItem "Radioterapia"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 2
CboAccesoSistema.AddItem "Recidente"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 3
CboAccesoSistema.AddItem "Psicologia"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 4
CboAccesoSistema.AddItem "Nutricion"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 5
CboAccesoSistema.AddItem "Administracion"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 6
CboAccesoSistema.AddItem "Tecnica"
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 7
CboAccesoSistema.AddItem ""
CboAccesoSistema.ItemData(CboAccesoSistema.NewIndex) = 8

Dim RsMedicoT As New ADODB.Recordset
CSql = "Select * From Medicos Where Tipo=2 or Tipo=3"
Set RsMedicoT = CrearRS(CSql)
RsMedicoT.MoveFirst
Do While Not RsMedicoT.EOF
    CboMedicoAsociado.AddItem Trim(RsMedicoT.Fields("Nombre").Value) & " " & Trim(RsMedicoT.Fields("Apellido").Value)
    CboMedicoAsociado.ItemData(CboMedicoAsociado.NewIndex) = RsMedicoT.Fields("IdMedico")
    RsMedicoT.MoveNext
Loop
RsMedicoT.Close


CSql = "SELECT * FROM Usuarios"
Set RsUsuarios = CrearRS(CSql)
RsUsuarios.MoveFirst
    
Carga_De_Datos

End Sub
Sub Carga_De_Datos()
      
If RsUsuarios.RecordCount <> 0 Then

    TxtCedula.Text = RsUsuarios.Fields("Cedula").Value
    TxtApellidos.Text = RsUsuarios.Fields("Apellidos").Value
    TxtNombres.Text = RsUsuarios.Fields("Nombre").Value
    TxtUsuario.Text = RsUsuarios.Fields("Usuario").Value
    TxtClave1.Text = RsUsuarios.Fields("Contraseña").Value
    TxtClave2.Text = RsUsuarios.Fields("Contraseña").Value
    DtpFechaRegistro.Value = RsUsuarios.Fields("Fecha").Value
    CboAccesoSistema.ListIndex = RsUsuarios.Fields("T_U").Value
    idUsu = RsUsuarios.Fields("IdUsuario").Value
    Actuali = 1
    If Not IsNull(RsUsuarios.Fields("IdMedicoT").Value) Then
        If RsUsuarios.Fields("IdMedicoT").Value = -1 Then
            CboMedicoAsociado.ListIndex = -1
        Else
            For T = 0 To CboMedicoAsociado.ListCount - 1
                If RsUsuarios.Fields("IdMedicoT").Value = CboMedicoAsociado.ItemData(T) Then
                    CboMedicoAsociado.ListIndex = T
                    Exit For
                End If
            Next T
        End If
    Else
        CboMedicoAsociado.ListIndex = -1
    End If
    
    If Not IsNull(RsUsuarios.Fields("Foto")) Then
        If RsUsuarios.Fields("Foto").Value <> "" And Dir(FotoEmp & "\" & RsUsuarios.Fields("Foto").Value) <> "" Then
            Image3.Picture = LoadPicture(FotoEmp & "\" & RsUsuarios.Fields("Foto").Value)
            FotoE = RsUsuarios.Fields("Foto").Value
        Else
            Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture: FotoE = ""
        End If
    Else
        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture: FotoE = ""
    End If

    NoReg.Caption = "Registro " & RsUsuarios.AbsolutePosition & " / " & RsUsuarios.RecordCount
    BtnAgregar.Enabled = True
    BtnGuardar.Enabled = True
    BtnEliminar.Enabled = True
    BtnAnterior.Enabled = True
    BtnSiguiente.Enabled = True
    BtnBuscar.Enabled = True
    
    Reg_Actual(3) = RsUsuarios.Fields("Cedula").Value

Else
    NoReg.Caption = "Registro 0 / 0"
    BtnAgregar.Enabled = True
    BtnGuardar.Enabled = True
    BtnEliminar.Enabled = False
    BtnAnterior.Enabled = False
    BtnSiguiente.Enabled = False
    BtnBuscar.Enabled = False
    Actuali = 0
    Reg_Actual(3) = ""
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If BD.State Then BD.Close

End Sub

Private Sub Text1_Change()
Cambio = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If Len(Text1.Text) > 8 Then KeyAscii = 0
End Sub

Private Sub Text2_Change()
Cambio = 0
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text2.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text2.Text)
    pru = LCase(Mid(Text2.Text, i, 1))
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

Text2.Text = StrText
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text3_Change()
Cambio = 0

Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text3.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text3.Text)
    pru = LCase(Mid(Text3.Text, i, 1))
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

Text3.Text = StrText
Text3.SelStart = Len(Text3.Text)
End Sub

Private Sub Image3_Click()
Dim TempCad As String
Dim TempCad2 As String
On Error GoTo h

CommonDialog1.ShowOpen
TempCad = CommonDialog1.filename

If InStr(1, TempCad, "\", vbTextCompare) = 0 Then
    If FotoE = "" Then FotoE = "Silueta.jpg"
    Exit Sub
End If

FotoE = Replace(Trim(TxtCedula.Text) & Trim(TxtApellidos.Text) & Trim(TxtNombres.Text) & ".jpg", " ", "")
TempCad2 = FotoEmp & "\" & FotoE
Call FileCopy(TempCad, TempCad2)

If Trim(FotoE) = "" Then Exit Sub
Image3.Picture = LoadPicture(TempCad2)
Image3.Refresh
Cambio = 1
Exit Sub
h:
MsgBox Err.Description
End Sub

'Private Sub Text3_KeyPress(KeyAscii As Integer)
'
'Select Case KeyAscii
'Case 48 To 57 ' permite el ingreso de numeros
'Case Is = 13 ' permite presionar el ENTER
'Call Command1_Click
'
'Case Is = 8 ' Permite Borrar de retroceso
'Case Else ' Inhibe todas las demas teclas
'KeyAscii = 0
'End Select
'
'End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtApellidos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtNombres.SetFocus
End If
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtCedula_LostFocus()
Dim IdCed

IdCed = Val(TxtCedula.Text)

If IdCed = 0 Then Exit Sub

CSql = "SELECT * FROM Usuarios Where cedula='" & IdCed & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If RegNew = 1 Then
        MsgBox "La Cedula ingresada ya se encuentra registrada!", vbCritical + vbOKOnly, "Error"
        TxtCedulaEmp.Text = ""
        TxtCedulaEmp.SetFocus
        Exit Sub
    Else
        If Val(Reg_Actual(3)) <> Val(TxtCedula.Text) Then
            MsgBox "La Cedula ingresada ya se encuentra registrada!", vbCritical + vbOKOnly, "Error"
            TxtCedulaEmp.Text = Reg_Actual(3)
            TxtCedulaEmp.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub TxtClave1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtClave2.SetFocus
End If
End Sub

Private Sub TxtClave2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CboAccesoSistema.SetFocus
End If
End Sub

Private Sub TxtNombres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DtpFechaRegistro.SetFocus
End If
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtClave1.SetFocus
End If
End Sub



VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPlanificacionPorPaciente 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación de horarios por paciente"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   Icon            =   "FrmPlanificacionPorPaciente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Sesiones:"
      Height          =   1095
      Left            =   10800
      TabIndex        =   38
      Top             =   2520
      Width           =   1215
      Begin VB.TextBox TxtNoSesiones 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Sesiones:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Paciente"
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12015
      Begin VB.TextBox TxtObservacionWeb 
         Height          =   975
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FrmPlanificacionPorPaciente.frx":1002
         Left            =   8760
         List            =   "FrmPlanificacionPorPaciente.frx":1004
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "FrmPlanificacionPorPaciente.frx":1006
         Left            =   3120
         List            =   "FrmPlanificacionPorPaciente.frx":1008
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Tratamiento RT:"
         Height          =   1095
         Left            =   6120
         TabIndex        =   8
         Top             =   2400
         Width           =   4455
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "2da. Hora"
            Height          =   255
            Left            =   3240
            TabIndex        =   43
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox CboGrupo2 
            Height          =   315
            ItemData        =   "FrmPlanificacionPorPaciente.frx":100A
            Left            =   1560
            List            =   "FrmPlanificacionPorPaciente.frx":10C2
            TabIndex        =   41
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox CboGrupo 
            Height          =   315
            ItemData        =   "FrmPlanificacionPorPaciente.frx":1396
            Left            =   1560
            List            =   "FrmPlanificacionPorPaciente.frx":144E
            TabIndex        =   9
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Asignada 2:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   780
            Width           =   1230
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Asignada 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1230
         End
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   17235971
         CurrentDate     =   39784
         MinDate         =   -108932
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   10080
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17235971
         CurrentDate     =   39784
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6840
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17235971
         CurrentDate     =   39784
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7680
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17235971
         CurrentDate     =   39784
      End
      Begin VB.Label NoReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
         Height          =   195
         Left            =   2520
         TabIndex        =   35
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento:"
         Height          =   195
         Left            =   2520
         TabIndex        =   33
         Top             =   1290
         Width           =   1560
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
         Height          =   195
         Left            =   8280
         TabIndex        =   32
         Top             =   1260
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Culminación:"
         Height          =   195
         Left            =   8400
         TabIndex        =   31
         Top             =   1890
         Width           =   1620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio:"
         Height          =   195
         Left            =   5640
         TabIndex        =   30
         Top             =   1860
         Width           =   1140
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edad:"
         Height          =   195
         Left            =   6120
         TabIndex        =   29
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido(s):"
         Height          =   195
         Left            =   7320
         TabIndex        =   28
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro:"
         Height          =   195
         Left            =   6240
         TabIndex        =   26
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula:"
         Height          =   195
         Left            =   2520
         TabIndex        =   25
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Historia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   120
         Picture         =   "FrmPlanificacionPorPaciente.frx":1722
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   2520
         TabIndex        =   22
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Años"
         Height          =   195
         Left            =   7320
         TabIndex        =   21
         Top             =   1290
         Width           =   360
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   5535
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1560
         Top             =   240
      End
      Begin VB.TextBox TxtBuscar 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese el Nombre, Apellido, Cedula de identidad o número de historia del paciente a buscar"
         Top             =   240
         Width           =   1935
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "FrmPlanificacionPorPaciente.frx":4AA3
         PICN            =   "FrmPlanificacionPorPaciente.frx":4ABF
         PICH            =   "FrmPlanificacionPorPaciente.frx":4D24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnListadoPaciente 
         Height          =   375
         Left            =   3600
         TabIndex        =   37
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Listado Pacientes"
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
         MICON           =   "FrmPlanificacionPorPaciente.frx":4FB6
         PICN            =   "FrmPlanificacionPorPaciente.frx":4FD2
         PICH            =   "FrmPlanificacionPorPaciente.frx":525B
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   3840
      Width           =   6375
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   5280
         TabIndex        =   1
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
         MICON           =   "FrmPlanificacionPorPaciente.frx":54EE
         PICN            =   "FrmPlanificacionPorPaciente.frx":550A
         PICH            =   "FrmPlanificacionPorPaciente.frx":56D3
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
         Left            =   120
         TabIndex        =   2
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
         MICON           =   "FrmPlanificacionPorPaciente.frx":5908
         PICN            =   "FrmPlanificacionPorPaciente.frx":5924
         PICH            =   "FrmPlanificacionPorPaciente.frx":5BB3
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
         Left            =   4080
         TabIndex        =   3
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
         MICON           =   "FrmPlanificacionPorPaciente.frx":5FF4
         PICN            =   "FrmPlanificacionPorPaciente.frx":6010
         PICH            =   "FrmPlanificacionPorPaciente.frx":62F2
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
Attribute VB_Name = "FrmPlanificacionPorPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsActualiarPlanificacion As New ADODB.Recordset
Dim IdPac, Grupo, Estatus
Dim RsPacientes As New ADODB.Recordset
Private Sub BtnAnterior_Click()
If RsPacientes.RecordCount <> 0 Then
    Call viryfi
    If RsPacientes.BOF = False Then
        RsPacientes.MovePrevious
        Else
        'RsPacientes.MoveFirst
        RsPacientes.MoveLast
    End If
    Call BuscarDatos
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
    IdPac = ""
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Sub viryfi()
Select Case Cambio
Case Is = 1
    Msg = "Este registro sufrió Cambios desea guardar?"
    d = MsgBox(Msg, vbYesNo, "Desea Guardar Cambios")
    Select Case d
    Case Is = 6
        BtnGuardar_Click
    Case Is = 7
    End Select
    Case Is = 0
End Select

'BtnAgregarPaciente.Enabled = True
'BtnBorrarPaciente.Enabled = True
Actuali = 0
st_but = 0

End Sub

Sub BuscarDatos()
On Error Resume Next
ACCION = EDITAR_REGISTRO
If RsPacientes.EOF Or RsPacientes.BOF Then
    'RsPacientes.MoveFirst
    Actuali = 0
    'RsPacientes.MoveFirst
Else
    Me.Caption = "Tabla de Pacientes - Id: " & RsPacientes.Fields("IdPaciente").Value
    IdPac = RsPacientes.Fields("IdPaciente").Value

    Text1.Text = RsPacientes.Fields("cedulap").Value
    DTPicker1.Value = RsPacientes.Fields("Fecha_regp").Value
    DTPicker2.Value = RsPacientes.Fields("Fecha_Inicio").Value
    DTPicker3.Value = RsPacientes.Fields("Fecha_culm").Value
    DTPicker4.Value = RsPacientes.Fields("Fecha_Nacimientop").Value

    Text4.Text = RsPacientes.Fields("Nombrep").Value
    Text3.Text = RsPacientes.Fields("Apellidop").Value

    Select Case Trim(RsPacientes.Fields("status").Value)
        Case Is = "A"
            Combo7.Text = "Activo"
        Case Is = "S"
            Combo7.Text = "Suspendido"
        Case Is = "R"
            Combo7.Text = "Reubicación"
        Case Is = "Si"
            Combo7.Text = "Simulación"
        Case Is = "E"
            Combo7.Text = "Espera"
        Case Is = "I"
            Combo7.Text = "Inactivo"
        Case Is = "C"
            Combo7.Text = "Culminado"
        Case Is = "F"
            Combo7.Text = "Fallecido"
    End Select


    Actuali = 1


    If Not IsNull(RsPacientes.Fields("foto").Value) Then
        If RsPacientes.Fields("foto").Value <> "" And Dir(Foto & "\" & RsPacientes.Fields("foto").Value) <> "" Then
            Image3.Picture = LoadPicture(Foto & "\" & RsPacientes.Fields("foto").Value)
            FotoP = RsPacientes.Fields("foto").Value
        Else
            Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
            FotoP = ""
        End If
    Else
        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        FotoP = ""
    End If

    Text11.Text = RsPacientes.Fields("Edadp").Value

    If Trim(RsPacientes.Fields("Historia").Value) <> "" Then Label16.Caption = RsPacientes.Fields("Historia").Value Else Label16.Caption = "NO TIENE"

    If RsPacientes.Fields("Sexop").Value = 0 Then Combo4.Text = "Masculino" Else Combo4.Text = "Femenino"
    If RsPacientes.Fields("ObservacionWeb").Value <> "" Then TxtObservacionWeb.Text = RsPacientes.Fields("ObservacionWeb").Value Else TxtObservacionWeb.Text = ""
    If RsPacientes.Fields("HoraAtencion").Value <> "" Then CboGrupo.Text = RsPacientes.Fields("HoraAtencion").Value Else CboGrupo.Text = ""

    For T = 0 To Combo4.ListCount - 1
        If Combo4.ItemData(T) = RsPacientes.Fields("sexop").Value Then
            Combo4.ListIndex = T
            Exit For
        End If
    Next T
    NoReg.Caption = RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount

    CSql = "Select * From Informe_Medico Where IdPaciente='" & IdPac & "'"
    Set RsPacientes = CrearRS(CSql)

    If Not IsNull(RsPacientes.Fields("Sesiones").Value) Then TxtNoSesiones.Text = RsPacientes.Fields("Sesiones").Value Else TxtNoSesiones.Text = ""
End If
'    Me.Caption = "Planificación de Horarios por Paciente - Id: " & RsPacientes.Fields("IdPaciente").Value
'    IdPac = RsPacientes.Fields("IdPaciente").Value
'    Text1.Text = RsPacientes.Fields("cedulap").Value
'    DTPicker1.Value = RsPacientes.Fields("Fecha_regp").Value
'    DTPicker2.Value = RsPacientes.Fields("Fecha_Inicio").Value
'    DTPicker3.Value = RsPacientes.Fields("Fecha_culm").Value
'    DTPicker4.Value = RsPacientes.Fields("Fecha_Nacimientop").Value
'
'    Text4.Text = RsPacientes.Fields("Nombrep").Value
'    Text3.Text = RsPacientes.Fields("Apellidop").Value
'
'    Select Case Trim(RsPacientes.Fields("status").Value)
'        Case Is = "A"
'            Combo7.Text = "Activo"
'        Case Is = "S"
'            Combo7.Text = "Suspendido"
'        Case Is = "R"
'            Combo7.Text = "Reubicación"
'        Case Is = "Si"
'            Combo7.Text = "Simulación"
'        Case Is = "E"
'            Combo7.Text = "Espera"
'        Case Is = "I"
'            Combo7.Text = "Inactivo"
'        Case Is = "C"
'            Combo7.Text = "Culminado"
'    End Select
'
'    Select Case RsPacientes.Fields("Tipo").Value
'        Case Is = 0
'            'OptPresupuesto.Value = True
'            'CboGrupo.Enabled = False
'            CboGrupo.Text = "00:00"
'        Case Is = 1
'            'OptPaciente.Value = True
'            'CboGrupo.Enabled = True
'            CboGrupo.Text = RsPacientes.Fields("HoraAtencion").Value
'    End Select
'
'    Actuali = 1
'    IdPac = RsPacientes.Fields("IdPaciente").Value
'
'    If Not IsNull(RsPacientes.Fields("foto").Value) Then
'        If RsPacientes.Fields("foto").Value <> "" And Dir(Foto & "\" & RsPacientes.Fields("foto").Value) <> "" Then
'            Image3.Picture = LoadPicture(Foto & "\" & RsPacientes.Fields("foto").Value)
'            FotoP = RsPacientes.Fields("foto").Value
'        Else
'            Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
'            FotoP = ""
'        End If
'    Else
'        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
'        FotoP = ""
'    End If
'
'    Text11.Text = RsPacientes.Fields("Edadp").Value
'
'    If Trim(RsPacientes.Fields("Historia").Value) <> "" Then Label16.Caption = RsPacientes.Fields("Historia").Value Else Label16.Caption = "No Tiene"
'    IdPac = RsPacientes.Fields("IdPaciente").Value
'    If RsPacientes.Fields("Sexop").Value = 0 Then Combo4.Text = "Masculino" Else Combo4.Text = "Femenino"
'    If RsPacientes.Fields("ObservacionWeb").Value <> "" Then TxtObservacionWeb.Text = RsPacientes.Fields("ObservacionWeb").Value Else TxtObservacionWeb.Text = ""
'
'    NoReg.Caption = "Registro: " & RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
'
'    CSql = "Select * From Informe_Medico Where IdPaciente='" & IdPac & "'"
'    Set RsPacientes = CrearRS(CSql)
'
'    If Not IsNull(RsPacientes.Fields("Sesiones").Value) Then TxtNoSesiones.Text = RsPacientes.Fields("Sesiones").Value Else TxtNoSesiones.Text = ""
'End If
Cambio = 0
'salir:
'    Exit Sub
End Sub


Public Sub BtnBuscar_Click()
On Error Resume Next

If TxtBuscar.Text = "" Then Exit Sub

CSql = "Select * From Paciente Where Activo='1' And (CedulaP = " & Val(TxtBuscar.Text) & " or NombreP like '%" & TxtBuscar.Text & "%' or ApellidoP like '%" & TxtBuscar.Text & "%' or Historia = '%" & TxtBuscar.Text & "%')"
Set RsPacientes = CrearRS(CSql)

If RsPacientes.RecordCount = 0 Then
    MsgBox "No Existe el registro"
    Cambio = 0
    Actuali = 0
    IdPac = ""
    BtnBorrarPaciente.Enabled = False
    Exit Sub
Else
    
    Me.Caption = "Planificación de Horarios por Paciente - Id: " & RsPacientes.Fields("IdPaciente").Value
    Text1.Text = RsPacientes.Fields("cedulap").Value
    DTPicker1.Value = RsPacientes.Fields("Fecha_regp").Value
    DTPicker2.Value = RsPacientes.Fields("Fecha_Inicio").Value
    DTPicker3.Value = RsPacientes.Fields("Fecha_culm").Value
    DTPicker4.Value = RsPacientes.Fields("Fecha_Nacimientop").Value
     
    Text4.Text = RsPacientes.Fields("Nombrep").Value
    Text3.Text = RsPacientes.Fields("Apellidop").Value
    
    Select Case Trim(RsPacientes.Fields("status").Value)
        Case Is = "A"
            Combo7.Text = "Activo"
        Case Is = "S"
            Combo7.Text = "Suspendido"
        Case Is = "R"
            Combo7.Text = "Reubicación"
        Case Is = "Si"
            Combo7.Text = "Simulación"
        Case Is = "E"
            Combo7.Text = "Espera"
        Case Is = "I"
            Combo7.Text = "Inactivo"
        Case Is = "C"
            Combo7.Text = "Culminado"
        Case Is = "F"
            Combo7.Text = "Fallecido"
    End Select
    
    Select Case RsPacientes.Fields("Tipo").Value
        Case Is = 0
            'OptPresupuesto.Value = True
            'CboGrupo.Enabled = False
            CboGrupo.Text = "00:00"
        Case Is = 1
            'OptPaciente.Value = True
            'CboGrupo.Enabled = True
            CboGrupo.Text = RsPacientes.Fields("HoraAtencion").Value
    End Select
    
    
    If RsPacientes.Fields("Atencion1").Value = True Then
        Check1.Value = 1
        CboGrupo2.Text = RsPacientes.Fields("HoraAtencion1").Value
        CboGrupo2.Enabled = True
        Label11.Enabled = True
    Else
        Check1.Value = 0
        CboGrupo2.Text = "00:00"
        CboGrupo2.Enabled = False
        Label11.Enabled = False
    End If
    
    
    Actuali = 1
    IdPac = RsPacientes.Fields("IdPaciente").Value

    If Not IsNull(RsPacientes.Fields("foto").Value) Then
        If RsPacientes.Fields("foto").Value <> "" And Dir(Foto & "\" & RsPacientes.Fields("foto").Value) <> "" Then
            Image3.Picture = LoadPicture(Foto & "\" & RsPacientes.Fields("foto").Value)
            FotoP = RsPacientes.Fields("foto").Value
        Else
            Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
            FotoP = ""
        End If
    Else
        Image3.Picture = FrmPrincipal.ListaImagenes.ListImages(1).Picture
        FotoP = ""
    End If
    
    Text11.Text = RsPacientes.Fields("Edadp").Value
    
    If Trim(RsPacientes.Fields("Historia").Value) <> "" Then Label16.Caption = RsPacientes.Fields("Historia").Value Else Label16.Caption = "No Tiene"
    IdPac = RsPacientes.Fields("IdPaciente").Value
    If RsPacientes.Fields("Sexop").Value = 0 Then Combo4.Text = "Masculino" Else Combo4.Text = "Femenino"
    If RsPacientes.Fields("ObservacionWeb").Value <> "" Then TxtObservacionWeb.Text = RsPacientes.Fields("ObservacionWeb").Value Else TxtObservacionWeb.Text = ""
    
    NoReg.Caption = "Registro: " & RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount
    
    CSql = "Select * From Informe_Medico Where IdPaciente='" & IdPac & "'"
    Set RsPacientes = CrearRS(CSql)
    
    If Not IsNull(RsPacientes.Fields("Sesiones").Value) Then TxtNoSesiones.Text = RsPacientes.Fields("Sesiones").Value Else TxtNoSesiones.Text = ""
End If
    


End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
End Sub

Private Sub BtnGuardar_Click()
On Error Resume Next
If CboGrupo.Text = "00:00" Then
    Msg = "Debe seleccionar un Horario de Atención para el paciente!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    CboGrupo.SetFocus
    Exit Sub
End If

If TxtNoSesiones.Text = "" Then
    Msg = "Debe de ingresar el No. de Sesiones del Tratamiento del paciente!!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    TxtNoSesiones.SetFocus
    Exit Sub
End If


CSql = "Select * From Paciente where IdPaciente='" & IdPac & "' And Activo='1'"
Set RsActualiarPlanificacion = CrearRS(CSql)

   
RsActualiarPlanificacion.Fields("HoraAtencion").Value = Trim(CboGrupo.Text)
Grupo = Trim(Mid(CboGrupo.Text, 1, 3)) & "00" & Mid(CboGrupo.Text, 6)

If Check1.Value = 1 Then
    RsActualiarPlanificacion.Fields("HoraAtencion1").Value = Trim(CboGrupo2.Text)
    RsActualiarPlanificacion.Fields("Atencion1").Value = 1
Else
    RsActualiarPlanificacion.Fields("HoraAtencion1").Value = Trim(CboGrupo2.Text)
    RsActualiarPlanificacion.Fields("Atencion1").Value = 0
End If


RsActualiarPlanificacion.Fields("Fecha_regp").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
RsActualiarPlanificacion.Fields("Fecha_Inicio").Value = Format(DTPicker2.Value, "dd/mm/yyyy")
RsActualiarPlanificacion.Fields("Fecha_culm").Value = Format(DTPicker3.Value, "dd/mm/yyyy")
RsActualiarPlanificacion.Fields("Tipo").Value = 1

RsActualiarPlanificacion.Fields("Grupo").Value = Grupo
If (Combo7.ListIndex) <> 3 Then
    Estatus = Chr(Combo7.ItemData(Combo7.ListIndex))
Else
    Estatus = "Si"
End If
    
RsActualiarPlanificacion.Fields("Status").Value = Estatus
RsActualiarPlanificacion.Fields("ObservacionWeb").Value = Trim(TxtObservacionWeb.Text)
RsActualiarPlanificacion.Update


CSql = "Select * From Informe_Medico where IdPaciente='" & IdPac & "'"
Set RsActualiarPlanificacion = CrearRS(CSql)


RsActualiarPlanificacion.Fields("Sesiones").Value = Trim(TxtNoSesiones.Text)
RsActualiarPlanificacion.Update

Msg = "Horario de Atención del Paciente Actualizado!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización de Horario de Atención"

  
FrmInformeDePlanificacion.InitGrid
FrmInformeDePlanificacion.Planificacion

'Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
'MsgBox Msg, vbInformation + vbOKOnly, "Actualización de Horario de Atención"
'
'EnviarRegPendiente

End Sub


Sub EnviarRegPendiente()
On Error Resume Next

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "Update Paciente set HoraAtencion1='" & CboGrupo2.Text & "', Atencion1='" & Check1.Value & "', Tipo='1', ObservacionWeb='" & Trim(TxtObservacionWeb.Text) & "', Grupo ='" & Trim(Grupo) & "', Status='" & Estatus & "', HoraAtencion='" & Trim(CboGrupo.Text) & "', Fecha_regp='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "', Fecha_Inicio='" & Format(DTPicker2.Value, "mm/dd/yyyy") & "', Fecha_culm='" & Format(DTPicker3.Value, "mm/dd/yyyy") & "' Where IdPaciente=" & IdPac & ""
Set RsActualiarPlanificacion = CrearRS(CSql)
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Nuevo Paciente"
RsRegPendiente.Fields("Tabla").Value = "Paciente"
RsRegPendiente.Fields("Condicional").Value = "IdPaciente=" & IdPac
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "Update Informe_Medico set sesiones='" & Trim(TxtNoSesiones.Text) & "' Where IdPaciente='" & IdPac & "'"
Set RsActualiarPlanificacion = CrearRS(CSql)
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Nuevo Paciente"
RsRegPendiente.Fields("Tabla").Value = "Paciente"
RsRegPendiente.Fields("Condicional").Value = "IdPaciente=" & IdPac
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"



End Sub


Private Sub BtnListadoPaciente_Click()
On Error Resume Next
Tipo = "Planificacion"
FrmListadoPaciente.Show vbModal
End Sub

Private Sub BtnSiguiente_Click()
If RsPacientes.RecordCount <> 0 Then
    Call viryfi
    If RsPacientes.EOF = False Then
        RsPacientes.MoveNext
    Else
        'RsPacientes.MoveLast
        RsPacientes.MoveFirst
    End If
    Call BuscarDatos
Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Label11.Enabled = True
    CboGrupo2.Enabled = True
Else
    Label11.Enabled = False
    CboGrupo2.Enabled = False
End If
End Sub

Private Sub Form_Load()

Combo4.AddItem "Masculino"
Combo4.ItemData(Combo4.NewIndex) = 0
Combo4.AddItem "Femenino"
Combo4.ItemData(Combo4.NewIndex) = 1

Label11.Enabled = False
CboGrupo2.Enabled = False

Combo7.AddItem "Activo"
Combo7.ItemData(Combo7.NewIndex) = Asc("A")
Combo7.AddItem "Suspendido"
Combo7.ItemData(Combo7.NewIndex) = Asc("S")
Combo7.AddItem "Reubicación"
Combo7.ItemData(Combo7.NewIndex) = Asc("R")
Combo7.AddItem "Simulación"
Combo7.ItemData(Combo7.NewIndex) = Asc("S") & Asc("i")
Combo7.AddItem "Espera"
Combo7.ItemData(Combo7.NewIndex) = Asc("E")
Combo7.AddItem "Fallecido"
Combo7.ItemData(Combo7.NewIndex) = Asc("F")
Combo7.AddItem "Inactivo"
Combo7.ItemData(Combo7.NewIndex) = Asc("I")
Combo7.AddItem "Culminado"
Combo7.ItemData(Combo7.NewIndex) = Asc("C")


CSql = "SELECT * FROM paciente"
Set RsPacientes = CrearRS(CSql)

NoReg.Caption = "Registro: " & RsPacientes.AbsolutePosition & " / " & RsPacientes.RecordCount

End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
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

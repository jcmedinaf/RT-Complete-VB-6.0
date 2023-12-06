VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmRegistroMedicoRemitente 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Medico"
   ClientHeight    =   6540
   ClientLeft      =   6675
   ClientTop       =   915
   ClientWidth     =   11505
   Icon            =   "RegistroRem.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos del Médico"
      Height          =   6375
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   11295
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   5520
         Width           =   3255
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el Nombre, Apellido o Cédula del Medico a Buscar"
            Top             =   240
            Width           =   1815
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2040
            TabIndex        =   26
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Buscar"
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
            MICON           =   "RegistroRem.frx":1002
            PICN            =   "RegistroRem.frx":101E
            PICH            =   "RegistroRem.frx":1283
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   2640
            Top             =   120
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "RegistroRem.frx":1515
         Left            =   6480
         List            =   "RegistroRem.frx":1522
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3480
         TabIndex        =   43
         Top             =   5520
         Width           =   7695
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6600
            TabIndex        =   23
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
            MICON           =   "RegistroRem.frx":1542
            PICN            =   "RegistroRem.frx":155E
            PICH            =   "RegistroRem.frx":1727
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
            TabIndex        =   18
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
            MICON           =   "RegistroRem.frx":195C
            PICN            =   "RegistroRem.frx":1978
            PICH            =   "RegistroRem.frx":1C07
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
            TabIndex        =   17
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
            MICON           =   "RegistroRem.frx":2048
            PICN            =   "RegistroRem.frx":2064
            PICH            =   "RegistroRem.frx":21F1
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
            TabIndex        =   22
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
            MICON           =   "RegistroRem.frx":2426
            PICN            =   "RegistroRem.frx":2442
            PICH            =   "RegistroRem.frx":2724
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
            MICON           =   "RegistroRem.frx":2975
            PICN            =   "RegistroRem.frx":2991
            PICH            =   "RegistroRem.frx":2B35
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
            Left            =   4560
            TabIndex        =   21
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
            MICON           =   "RegistroRem.frx":2CD4
            PICN            =   "RegistroRem.frx":2CF0
            PICH            =   "RegistroRem.frx":2F86
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
            Left            =   3960
            TabIndex        =   20
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
            MICON           =   "RegistroRem.frx":31E5
            PICN            =   "RegistroRem.frx":3201
            PICH            =   "RegistroRem.frx":3496
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
         Caption         =   "Datos Web"
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   4800
         Width           =   11055
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   4095
         End
         Begin VB.TextBox Text13 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   6360
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario Web:"
            Height          =   195
            Left            =   150
            TabIndex        =   42
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clave Web:"
            Height          =   195
            Left            =   5400
            TabIndex        =   41
            Top             =   330
            Width           =   840
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "RegistroRem.frx":36F2
         Left            =   1320
         List            =   "RegistroRem.frx":3708
         TabIndex        =   5
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "RegistroRem.frx":3730
         Left            =   6480
         List            =   "RegistroRem.frx":3732
         TabIndex        =   7
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   4320
         Width           =   4095
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   3360
         Width           =   4095
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   2880
         Width           =   4095
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   2880
         Width           =   4695
      End
      Begin VB.TextBox Text7 
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   9855
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   6480
         TabIndex        =   13
         Top             =   3840
         Width           =   4695
      End
      Begin ChamaleonButton.ChameleonBtn BtnListadoMedico 
         Height          =   375
         Left            =   8880
         TabIndex        =   24
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   4320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Listado de Médico"
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
         MICON           =   "RegistroRem.frx":3734
         PICN            =   "RegistroRem.frx":3750
         PICH            =   "RegistroRem.frx":39D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mmmmm"
         Height          =   195
         Left            =   8880
         TabIndex        =   46
         Top             =   570
         Width           =   600
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tratante / Remitente:"
         Height          =   195
         Left            =   4800
         TabIndex        =   44
         Top             =   540
         Width           =   1530
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MSDS:"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   4410
         Width           =   510
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   3450
         Width           =   480
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefóno Hab:"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   2460
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2970
         Width           =   540
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   5715
         TabIndex        =   35
         Top             =   2970
         Width           =   540
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Telf. Movil:"
         Height          =   195
         Left            =   5595
         TabIndex        =   33
         Top             =   2460
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido(s):"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clinica:"
         Height          =   195
         Left            =   195
         TabIndex        =   31
         Top             =   3930
         Width           =   510
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especialidad:"
         Height          =   195
         Left            =   5520
         TabIndex        =   30
         Top             =   3930
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         Height          =   195
         Left            =   5610
         TabIndex        =   29
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cédula:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   570
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmRegistroMedicoRemitente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UsuaroCount As Variant '//Mantiene número de registros
Public RsRegMedico As New ADODB.Recordset
Public IdMed
Dim camb
Public NewReg '1 es nuevo registro y agrega; 0 es un registro ya existente
Dim CntReg 'contador de registros
Dim PosReg 'posicion de registros
Dim Reg_Actual(0 To 20) As String
Dim IdLIdMed As String

Public Sub BtnDesHacer_Click()
On Error Resume Next
    BtnBorrar.Enabled = True
    BtnAgregar.Enabled = True
    Call CoNect
    Call Carga_De_Datos
    NewReg = 0
    camb = 0
End Sub

Sub EnviarRegPendiente(ByVal IdMed2 As Integer, ByVal IdLIdPac2 As String)
On Error Resume Next

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

a = 1

CSql = "SELECT * FROM Medicos WHERE IdMedico = " & IdMed2 & " AND IdL = '" & IdLIdPac2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Medicos (["
For i = 0 To RsTemp.Fields.Count - 1
    If Not i = (RsTemp.Fields.Count - 1) Then
        StrSen = StrSen & RsTemp.Fields(i).Name & "],["
    Else
        StrSen = StrSen & RsTemp.Fields(i).Name & "]) VALUES ("
    End If
Next i
For i = 0 To RsTemp.Fields.Count - 1
    If Not i = (RsTemp.Fields.Count - 1) Then
        StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "',"
    Else
        StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "')"
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = Replace(StrSen, "'", "(varCSP)")

CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Nuevo Medico"
RsRegPendiente.Fields("Tabla").Value = "Medicos"
RsRegPendiente.Fields("Condicional").Value = "IdMedico=" & IdMed2 & " AND IdL='" & IdLIdPac2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub


Private Sub BtnListadoMedico_Click()
Tipo = "Nuevo Medico"
FrmListadoMedicos.Show vbModal, FrmPrincipal
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub BtnAgregar_Click()
On Error Resume Next
'command6
 Call Blanqueo
 BtnBorrar.Enabled = False
 BtnAgregar.Enabled = False
 BtnAnterior.Enabled = False
 BtnSiguiente.Enabled = False
 BtnListadoMedico.Enabled = False
 NewReg = 1
End Sub

Private Sub BtnAnterior_Click()
On Error Resume Next
If RsRegMedico.RecordCount <> 0 Then
    Call Blanqueo
    If Not RsRegMedico.BOF Then
        RsRegMedico.MovePrevious
        If RsRegMedico.BOF Then RsRegMedico.MoveLast
        'PosReg = PosReg - 1
        'If RsRegMedico.BOF Then RsRegMedico.MoveLast: PosReg = CntReg
        Call Carga_De_Datos
    Else
    End If
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
    IdMed = ""
End If
End Sub

Private Sub BtnBorrar_Click()
On Error Resume Next
Dim RsBorrarMed As New ADODB.Recordset
'command2

If IdMed = "" Then MsgBox "Debe seleccionar un registro!", vbExclamation + vbOKOnly, "Operacion no permitida!": Exit Sub

Msg = "Esta Seguro de Eliminar Este Medico?" & Chr(13) & Chr(13) & Text1.Text & " " & Text2.Text & " " & Text5.Text
p = MsgBox(Msg, vbQuestion + vbYesNo, "Eliminar Medico")
Select Case p
Case Is = 6
        Call Enviar_Bitacora(IdUser, "Nuevo Medico", "BORRAR", "Se elimino el registro COMPLETO del MEDICO cuya IdMedico=" & IdMed & " y IdL=" & IdLIdMed)
        'CSql = "delete from Medicos where [IdMedico] = " & IDmed
        CSql = "UPDATE Medicos set Activo=0 where [IdMedico] = " & IdMed & " AND IdL='" & IdLIdMed & "'"
        Set RsBorrarMed = CrearRS(CSql)
        Call EnviarRegPendiente(IdMed, IdLIdMed)
        Call Blanqueo
            Call CoNect
            Call Carga_De_Datos
        MsgBox "El registro fue eliminado.", vbInformation + vbOKOnly, "Operacion Exitosa!"
Case Is = 7
End Select
End Sub

Sub EnviarRegPendienteBorrar()

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "Update Medicos Set Activo='0' Where IdPaciente = " & IdPac & " And Activo='1'"
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


Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Public Sub BtnBuscar_Click()
On Error Resume Next
Blanqueo
IdMed = ""

BtnDesHacer_Click

CSql = "SELECT * FROM Medicos Where nombre like '%" & TxtBuscar & "%' or apellido like '%" & TxtBuscar & "%' or cedula like '" & TxtBuscar & "%' AND Activo='1'"
Set RsRegMedico = CrearRS(CSql)

If RsRegMedico.RecordCount = 0 Then
    MsgBox "No se encontraron registros asociados a la Busqueda!", vbInformation + vbOKOnly, "No Encontrado!"
    BtnBorrar.Enabled = False
    IdMed = ""
    Label15.Caption = "Registro 0 / 0"
    Blanqueo
    Exit Sub
End If

RsRegMedico.MoveFirst
CntReg = RsRegMedico.RecordCount
Call Carga_De_Datos
NewReg = 0
BtnBorrar.Enabled = True
camb = 0

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
Dim RsRegMedicoR As New ADODB.Recordset
Dim NuevoId As String
Dim IdTemp As String
Dim CSql As String

'command1
    t1 = Text1.Text 'cedula
    t2 = Text2.Text 'nombres
    t3 = Text3.Text 'especializacion
    t4 = Text4.Text 'clinica
    t5 = Text5.Text 'apellidos
    t6 = Text6.Text 'telefono movil
    t7 = Text7.Text 'direccion
    t8 = Text8.Text 'ciudad
    t9 = Text9.Text 'estado
    t10 = Text10.Text 'telefono hab
    t11 = Text11.Text 'email
    t12 = Text12.Text 'usuario web
    t13 = Text13.Text 'clave web
    t14 = Text14.Text 'msds
    t15 = Combo3.ListIndex + 1 'Tipo
    t16 = Combo2.Text  'codigo de area telefono
    t17 = Combo1.Text  'codigo de area celular
    
    If NewReg = 0 Then
        If Not Reg_Actual(0) = t2 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo NOMBRE de (" & Reg_Actual(0) & ")  a  (" & t2 & ")")
        End If
        If Not Reg_Actual(1) = t1 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CEDULA de (" & Reg_Actual(1) & ")  a  (" & t1 & ")")
        End If
        If Not Reg_Actual(2) = t3 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo ESPECIALIDAD de (" & Reg_Actual(2) & ")  a  (" & t3 & ")")
        End If
        If Not Reg_Actual(3) = t4 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CLINICA de (" & Reg_Actual(3) & ")  a  (" & t4 & ")")
        End If
        If Not Reg_Actual(4) = t5 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo APELLIDO de (" & Reg_Actual(4) & ")  a  (" & t5 & ")")
        End If
        If Not Reg_Actual(5) = t7 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo DIRECCION de (" & Reg_Actual(5) & ")  a  (" & t7 & ")")
        End If
        If Not Reg_Actual(6) = t10 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo TELEFONO de (" & Reg_Actual(6) & ")  a  (" & t10 & ")")
        End If
        If Not Reg_Actual(7) = t6 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CELULAR de (" & Reg_Actual(7) & ")  a  (" & t6 & ")")
        End If
        If Not Reg_Actual(8) = t11 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo EMAIL de (" & Reg_Actual(8) & ")  a  (" & t11 & ")")
        End If
        If Not Reg_Actual(9) = t9 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo ESTADO de (" & Reg_Actual(9) & ")  a  (" & t9 & ")")
        End If
        If Not Reg_Actual(10) = t8 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CIUDAD de (" & Reg_Actual(10) & ")  a  (" & t8 & ")")
        End If
        If Not Reg_Actual(11) = t12 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CUENTA de (" & Reg_Actual(11) & ")  a  (" & t12 & ")")
        End If
        If Not Reg_Actual(12) = t13 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CONTRASEÑA de (" & Reg_Actual(12) & ")  a  (" & t13 & ")")
        End If
        If Not Reg_Actual(13) = t14 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo MSDS de (" & Reg_Actual(13) & ")  a  (" & t14 & ")")
        End If
        If Not Reg_Actual(14) = t15 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo TIPO de (" & Reg_Actual(14) & ")  a  (" & t15 & ")")
        End If
        If Not Reg_Actual(15) = t16 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CODIGO de (" & Reg_Actual(15) & ")  a  (" & t16 & ")")
        End If
        If Not Reg_Actual(16) = t17 Then
            Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Modificar", "se modifico el campo CODIGOC de (" & Reg_Actual(16) & ")  a  (" & t17 & ")")
        End If
        Else
        Call Enviar_Bitacora(IdUser, "Nuevo Medico", "Agregar", "Agrego un nuevo registro a la tabla MEDICOS cuya id=" & NuevoId)
    End If
'Exit Sub
    If Replace(t2, " ", "") = "" Then
        MsgBox "Debe ingresar el NOMBRE del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
        Text2.SetFocus
        Exit Sub
    ElseIf Replace(t5, " ", "") = "" Then
        MsgBox "Debe ingresar el APELLIDO del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
        Text5.SetFocus
        Exit Sub
    End If

    If UCase(Combo3.Text) = UCase("Tratante") Or UCase(Combo3.Text) = UCase("Ambos") Then

        If Replace(t1, " ", "") = "" Then
            MsgBox "Debe ingresar la CEDULA del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text1.SetFocus
            Exit Sub
        ElseIf Replace(t3, " ", "") = "" Then
            MsgBox "Debe ingresar la ESPECIALIZACION del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text3.SetFocus
            Exit Sub
        ElseIf Replace(t4, " ", "") = "" Then
            MsgBox "Debe ingresar la CLINICA del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text4.SetFocus
            Exit Sub
        ElseIf Replace(t6, " ", "") = "" Then
            MsgBox "Debe ingresar el TELEFONO MOV. del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text6.SetFocus
            Exit Sub
        ElseIf Replace(t7, " ", "") = "" Then
            MsgBox "Debe ingresar la DIRECCION del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text7.SetFocus
            Exit Sub
        ElseIf Replace(t8, " ", "") = "" Then
            MsgBox "Debe ingresar la CIUDAD del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text8.SetFocus
            Exit Sub
        ElseIf Replace(t9, " ", "") = "" Then
            MsgBox "Debe ingresar el ESTADO del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text9.SetFocus
            Exit Sub
        ElseIf Replace(t10, " ", "") = "" Then
            MsgBox "Debe ingresar el TELEFONO HAB. del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text10.SetFocus
            Exit Sub
        ElseIf Replace(t11, " ", "") = "" Then
            MsgBox "Debe ingresar el EMAIL del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text11.SetFocus
            Exit Sub
        ElseIf Replace(t12, " ", "") = "" Then
            MsgBox "Debe ingresar el USUARIO WEB del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text12.SetFocus
            Exit Sub
        ElseIf Replace(t13, " ", "") = "" Then
            MsgBox "Debe ingresar la CLAVE WEB del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text13.SetFocus
            Exit Sub
        ElseIf Replace(t14, " ", "") = "" Then
            MsgBox "Debe ingresar el MSDS del Medico!", vbExclamation + vbOKOnly, "Faltan Datos!"
            Text14.SetFocus
            Exit Sub
        End If
        
        If Replace(t10, " ", "") <> "" Then
            If Replace(t16, " ", "") = "" Then MsgBox "Seleccione el Codigo de Area!", vbExclamation + vbOKOnly, "Faltan Datos!": Combo2.SetFocus: Exit Sub
        End If

        If Replace(t6, " ", "") <> "" Then
            If Replace(t17, " ", "") = "" Then MsgBox "Seleccione el Codigo de Area!", vbExclamation + vbOKOnly, "Faltan Datos!": Combo1.SetFocus: Exit Sub
        End If
        If t1 = "" Then
            f = "El Campo Cédula está Vacio, Debe de llenar Todos los Datos"
            GoTo noguardA
        End If
    End If

If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If

Select Case NewReg
    Case Is = 1
        If camb = 1 Then
            'CSql = "Insert into Medicos(idusuario, CEDULA, NOMBRE, Especialidad, clinica, apellido, email, contraseña, direccion, ciudad, estado, telefono, celular, email2, msds) VALUES(" & IdUser & "," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & t5 & "','" & t6 & "','" & t7 & "','" & t8 & "','" & t9 & "','" & t10 & "','" & t11 & "','" & t12 & "','" & t3 & "','" & t14 & "')"
            CSql = "Select MAX(idmedico)+1 AS NuevoId FROM Medicos"
            Set RsRegMedicoR = CrearRS(CSql)
            
            If Not IsNull(RsRegMedicoR.Fields("NuevoId")) Then
                NuevoId = RsRegMedicoR.Fields("NuevoId")
            Else
                NuevoId = "0"
            End If
            'CSql = "Insert into Medicos(IdMedico, idusuario, Nombre, Apellido, Cedula, Direccion, Telefono, Celular, Email, Estado, Ciudad, Clinica, Especialidad, Cuenta, Contraseña, msds, Tipo, Codigo, Codigoc) VALUES(" & NuevoId & "," & IdUser & ",'" & t2 & "','" & t5 & "','" & t1 & "','" & t7 & "','" & t10 & "','" & t6 & "','" & t11 & "','" & t9 & "','" & t8 & "','" & t4 & "','" & t3 & "','" & t12 & "','" & t13 & "','" & t14 & "'," & t15 & "," & t16 & "," & t17 & ")"
            CSql = "select * from Medicos"
            'RsRegMedicoR.Open CSql, Cnn
            IdLIdMed = NuevoIdL
            Set RsRegMedicoR = CrearRS(CSql)
            
            RsRegMedicoR.AddNew
            RsRegMedicoR.Fields("IdMedico").Value = NuevoId
            RsRegMedicoR.Fields("IdL").Value = IdLIdMed
            RsRegMedicoR.Fields("idusuario").Value = IdUser
            RsRegMedicoR.Fields("Nombre").Value = t2
            RsRegMedicoR.Fields("Apellido").Value = t5
            RsRegMedicoR.Fields("Cedula").Value = t1
            RsRegMedicoR.Fields("Direccion").Value = t7
            RsRegMedicoR.Fields("Telefono").Value = t10
            RsRegMedicoR.Fields("Celular").Value = t6
            RsRegMedicoR.Fields("Email").Value = t11
            RsRegMedicoR.Fields("Estado").Value = t9
            RsRegMedicoR.Fields("Ciudad").Value = t8
            RsRegMedicoR.Fields("Clinica").Value = t4
            RsRegMedicoR.Fields("Especialidad").Value = t3
            RsRegMedicoR.Fields("Cuenta").Value = t12
            RsRegMedicoR.Fields("Contraseña").Value = t13
            RsRegMedicoR.Fields("msds").Value = t14
            RsRegMedicoR.Fields("Tipo").Value = t15
            RsRegMedicoR.Fields("Codigo").Value = Val(t16)
            RsRegMedicoR.Fields("Codigoc").Value = Val(t17)
            RsRegMedicoR.Fields("Activo").Value = "1"
            RsRegMedicoR.Update
            
            Msg = "Registro Agregado satisfactoriamente"
            MsgBox Msg, vbInformation + vbOKOnly, "Operacion Exitosa!"
            
            Call EnviarRegPendiente(NuevoId, IdLIdMed)
            
            Call Blanqueo
            Call CoNect
            Call Carga_De_Datos
            Exit Sub
        Else
            f = "No hay cambios que agregar"
            GoTo noguardA
        End If
    Case Is = 0
        If camb = 1 Then
        If IdMed = "" Then MsgBox "Debe seleccionar un registro para realizador cambios!" & Chr(13) & _
        "Si desea agregar un nuevo registro, Seleccione AGREGAR", vbExclamation + vbOKOnly, "Operacion no permitida!": Exit Sub
            
            CSql = "Update medicos set IdL='" & IdLIdMed & "', cedula = '" & t1 & "', nombre = '" & t2 & "' , especialidad ='" & t3 & _
                    "' , clinica = '" & t4 & "' , apellido = '" & t5 & "' , celular = '" & t6 & _
                    "' , direccion ='" & t7 & "' , ciudad = '" & t8 & "' , estado = '" & t9 & _
                    "' , telefono = '" & t10 & "' , email = '" & t11 & "' , contraseña = '" & t13 & _
                    "' , msds = '" & t14 & "', Tipo='" & t15 & "', cuenta='" & t12 & _
                    "' , codigo=" & t16 & " , codigoc=" & t17 & " where idmedico = " & IdMed

            Set RsRegMedicoR = CrearRS(CSql)
            Msg = "Registro actualizado satisfactoriamente"
            MsgBox Msg, vbInformation + vbOKOnly, "Operacion Exitosa!"
            bhu = "[idmedico] = " & IdMed
            Call EnviarRegPendiente(IdMed, IdLIdMed)
            Call Blanqueo
            Call CoNect
             RsRegMedico.Find bhu, , adSearchForward, 1
            Call Carga_De_Datos
        Else
            f = "No hay cambios que guardar en este registro"
            GoTo noguardA
        End If
End Select

BtnDesHacer_Click

NewReg = 0
camb = 0
Exit Sub

noguardA:
    MsgBox f, vbInformation + vbOKOnly, "Error al Guardar"
    Exit Sub

End Sub

Sub Blanqueo()
    Text1.Text = "" 'cedula
    Text2.Text = "" 'nombres
    Text3.Text = "" 'especializacion
    Text4.Text = "" 'clinica
    Text5.Text = "" 'apellidos
    Text6.Text = "" 'telefono movil
    Text7.Text = "" 'direccion
    Text8.Text = "" 'ciudad
    Text9.Text = "" 'estado
    Text10.Text = "" 'telefono hab
    Text11.Text = "" 'email
    Text12.Text = "" 'usuario web
    Text13.Text = "" 'clave web
    Text14.Text = "" 'msds
    Combo3.ListIndex = 0 'Tipo
    Combo2.Text = "" 'codigo de area telefono
    Combo1.Text = "" 'codigo de area celular
     
     
     camb = 0
End Sub

Private Sub BtnSiguiente_Click()
On Error Resume Next
If RsRegMedico.RecordCount <> 0 Then
    Call Blanqueo
    If Not RsRegMedico.EOF Then
        RsRegMedico.MoveNext
        If RsRegMedico.EOF Then RsRegMedico.MoveFirst
        Call Carga_De_Datos
    End If
    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
    IdMed = ""
End If
End Sub

Private Sub Combo1_Click()
camb = 1
End Sub

Private Sub Combo2_Click()
camb = 1
End Sub

Private Sub Combo3_Click()
camb = 1
End Sub

Private Sub Form_Load()
'Me.Height = 7635
'Me.Width = 11160
'Centrar Me
On Error Resume Next
CSql = "SELECT * FROM Codigos_T"
Set RsCodigos = CrearRS(CSql)

While Not RsCodigos.EOF
    Combo1.AddItem RsCodigos.Fields("codigo").Value
    Combo1.ItemData(Combo1.NewIndex) = RsCodigos.Fields("idcodigot").Value
        
    Combo2.AddItem RsCodigos.Fields("codigo").Value
    Combo2.ItemData(Combo2.NewIndex) = RsCodigos.Fields("idcodigot").Value
    RsCodigos.MoveNext
Wend

Call CoNect
Call Carga_De_Datos
NewReg = 0
     
End Sub
Sub CoNect()
On Error Resume Next
'If RsRegMedico.State Then RsRegMedico.Close
CSql = "SELECT * FROM Medicos WHERE Activo='1'"
Set RsRegMedico = CrearRS(CSql)
RsRegMedico.MoveFirst
CntReg = RsRegMedico.RecordCount
End Sub

Public Sub Carga_De_Datos()
On Error Resume Next

'        If RsRegMedico.EOF Or RsRegMedico.BOF Then
'            msg = "LLego al Final del Registro"
'            MsgBox msg
'            RsRegMedico.MoveFirst
'        End If

If RsRegMedico.RecordCount <> 0 Then
    IdLIdMed = RsRegMedico.Fields("IdL")
    BtnAnterior.Enabled = True
    BtnSiguiente.Enabled = True
    BtnListadoMedico.Enabled = True
    BtnAgregar.Enabled = True
    BtnBorrar.Enabled = True
Else
    IdLIdMed = ""
    BtnAnterior.Enabled = False
    BtnSiguiente.Enabled = False
    BtnListadoMedico.Enabled = False
    BtnAgregar.Enabled = True
    BtnBorrar.Enabled = False
End If

    If Not IsNull(RsRegMedico.Fields("Nombre")) Then Text2.Text = RsRegMedico.Fields("Nombre") Else Text2.Text = ""
    If Not IsNull(RsRegMedico.Fields("Cedula")) Then Text1.Text = RsRegMedico.Fields("Cedula") Else Text1.Text = ""
    If Not IsNull(RsRegMedico.Fields("Especialidad")) Then Text3.Text = RsRegMedico.Fields("Especialidad") Else Text3.Text = ""
    If Not IsNull(RsRegMedico.Fields("Clinica")) Then Text4.Text = RsRegMedico.Fields("Clinica") Else Text4.Text = ""
    If Not IsNull(RsRegMedico.Fields("apellido")) Then Text5.Text = RsRegMedico.Fields("apellido") Else Text5.Text = ""
    If Not IsNull(RsRegMedico.Fields("direccion")) Then Text7.Text = RsRegMedico.Fields("direccion") Else Text7.Text = ""
    If Not IsNull(RsRegMedico.Fields("telefono")) Then Text10.Text = RsRegMedico.Fields("telefono") Else Text10.Text = ""
    If Not IsNull(RsRegMedico.Fields("celular")) Then Text6.Text = RsRegMedico.Fields("celular") Else Text6.Text = ""
    If Not IsNull(RsRegMedico.Fields("email")) Then Text11.Text = RsRegMedico.Fields("email") Else Text11.Text = ""
    If Not IsNull(RsRegMedico.Fields("estado")) Then Text9.Text = RsRegMedico.Fields("estado") Else Text9.Text = ""
    If Not IsNull(RsRegMedico.Fields("ciudad")) Then Text8.Text = RsRegMedico.Fields("ciudad") Else Text8.Text = ""
    If Not IsNull(RsRegMedico.Fields("cuenta")) Then Text12.Text = RsRegMedico.Fields("cuenta") Else Text12.Text = ""
    If Not IsNull(RsRegMedico.Fields("contraseña")) Then Text13.Text = RsRegMedico.Fields("contraseña") Else Text13.Text = ""
    If Not IsNull(RsRegMedico.Fields("msds")) Then Text14.Text = RsRegMedico.Fields("msds") Else Text14.Text = ""
    If Not IsNull(RsRegMedico.Fields("tipo")) Then Combo3.ListIndex = (Val(RsRegMedico.Fields("Tipo"))) - 1 Else Combo3.Text = ""
    If Not IsNull(RsRegMedico.Fields("codigo")) Then Combo2.Text = RsRegMedico.Fields("codigo") Else Combo2.Text = ""
    If Not IsNull(RsRegMedico.Fields("codigoc")) Then Combo1.Text = RsRegMedico.Fields("codigoc") Else Combo1.Text = ""
    
    Reg_Actual(0) = Text2.Text
    Reg_Actual(1) = Text1.Text
    Reg_Actual(2) = Text3.Text
    Reg_Actual(3) = Text4.Text
    Reg_Actual(4) = Text5.Text
    Reg_Actual(5) = Text7.Text
    Reg_Actual(6) = Text10.Text
    Reg_Actual(7) = Text6.Text
    Reg_Actual(8) = Text11.Text
    Reg_Actual(9) = Text9.Text
    Reg_Actual(10) = Text8.Text
    Reg_Actual(11) = Text12.Text
    Reg_Actual(12) = Text13.Text
    Reg_Actual(13) = Text14.Text
    Reg_Actual(14) = Combo3.ListIndex + 1
    Reg_Actual(15) = Combo2.Text
    Reg_Actual(16) = Combo1.Text
    
    CntReg = RsRegMedico.RecordCount
    
    Label15.Caption = "Registro: " & RsRegMedico.AbsolutePosition & "/" & CntReg
    IdMed = RsRegMedico.Fields("idmedico")
    camb = 0
    NewReg = 0
            
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmNuevoPaciente.carga_lista_medicosr
FrmNuevoPaciente.carga_lista_medicost
If Not ACCION = AGREGAR_REGISTRO Then FrmNuevoPaciente.BuscarDatos
RsRegMedico.Close
End Sub

Private Sub Text1_Change()
camb = 1
End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text12.SetFocus
        Case vbKeyUp
            Text4.SetFocus
        Case vbKeyDown
            Text12.SetFocus
    End Select
End If
End Sub

Private Sub Text2_Change()
camb = 1
End Sub

Private Sub Text3_Change()
camb = 1
End Sub

Private Sub Text4_Change()
camb = 1
End Sub

Private Sub Text5_Change()
camb = 1
End Sub

Private Sub Text6_Change()
camb = 1
End Sub

Private Sub Text7_Change()
camb = 1
End Sub

Private Sub Text8_Change()
camb = 1
End Sub

Private Sub Text9_Change()
camb = 1
End Sub
Private Sub Text10_Change()
camb = 1
End Sub
Private Sub Text11_Change()
camb = 1
End Sub
Private Sub Text12_Change()
camb = 1
End Sub
Private Sub Text13_Change()
camb = 1
End Sub
Private Sub Text14_Change()
camb = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text5.SetFocus
        Case vbKeyDown
            Text5.SetFocus
    End Select
End If
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Text2.SetFocus
        Case vbKeyDown
            Text7.SetFocus
    End Select
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text7.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyLeft
            Text5.SetFocus
        Case vbKeyDown
            Text7.SetFocus
    End Select
End If
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(Text7.Text) <> 0 And Text7.SelLength = 0 Then
    Gram = Gramatica(Text7.Text, Text7.SelStart)
    Text7.Text = Gram.Texto
    Text7.SelStart = Gram.Poscur
End If

If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            'TxtCentroClinico.SetFocus
        Case vbKeyUp
            Text5.SetFocus
        Case vbKeyDown
            Combo2.SetFocus
    End Select
End If
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text10.SetFocus
        Case vbKeyUp
            Text7.SetFocus
        Case vbKeyRight
            Text10.SetFocus
        Case vbKeyDown
            Text9.SetFocus
    End Select
End If
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text8.SetFocus
        Case vbKeyUp
            Combo2.SetFocus
        Case vbKeyRight
            Text8.SetFocus
        Case vbKeyDown
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyRight
            Text3.SetFocus
        Case vbKeyDown
            Text14.SetFocus
    End Select
End If
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyUp
            Text9.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            Text7.SetFocus
        Case vbKeyLeft
            Combo2.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            Text9.SetFocus
    End Select
End If
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyUp
            Text7.SetFocus
        Case vbKeyLeft
            Text10.SetFocus
        Case vbKeyRight
            Text6.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            Text7.SetFocus
        Case vbKeyLeft
            Combo1.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub


Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text11.SetFocus
        Case vbKeyUp
            Text6.SetFocus
        Case vbKeyLeft
            Text9.SetFocus
        Case vbKeyDown
            Text11.SetFocus
    End Select
End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text14.SetFocus
        Case vbKeyUp
            Text11.SetFocus
        Case vbKeyLeft
            Text4.SetFocus
        Case vbKeyDown
            Text13.SetFocus
    End Select
End If
End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text13.SetFocus
        Case vbKeyUp
            Text14.SetFocus
        Case vbKeyRight
            Text13.SetFocus
        Case vbKeyDown
            BtnGuardarActualizar.SetFocus
    End Select
End If
End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGuardarActualizar.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyLeft
            Text12.SetFocus
        Case vbKeyDown
            BtnDesHacer.SetFocus
    End Select
End If
End Sub

Private Sub TxtBuscar_GotFocus()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnBuscar_Click
End Sub

Private Sub TxtBuscar_LostFocus()
If Trim(TxtBuscar.Text) = "" Then TxtBuscar.Text = "Busqueda"
End Sub

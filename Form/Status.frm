VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmStatus 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estatus del Paciente"
   ClientHeight    =   8235
   ClientLeft      =   4245
   ClientTop       =   795
   ClientWidth     =   12720
   Icon            =   "Status.frx":0000
   LinkTopic       =   "Form39"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   12720
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   7200
         Width           =   3615
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
            TabIndex        =   2
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad"
            Top             =   240
            Width           =   1815
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2040
            TabIndex        =   3
            ToolTipText     =   "Buscar "
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            MICON           =   "Status.frx":1002
            PICN            =   "Status.frx":101E
            PICH            =   "Status.frx":1283
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
         Caption         =   "Datos Paciente"
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8535
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   3240
            Top             =   480
         End
         Begin VB.Label TxtNombre 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4320
            TabIndex        =   24
            Top             =   1200
            Width           =   4095
         End
         Begin VB.Label TxtApellido 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   4095
         End
         Begin VB.Label TxtHistoria 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6360
            TabIndex        =   22
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label TxtCedula 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
            Height          =   195
            Left            =   4320
            TabIndex        =   16
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Historia:"
            Height          =   195
            Left            =   6360
            TabIndex        =   14
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   1695
         Left            =   8760
         TabIndex        =   10
         Top             =   240
         Width           =   3615
         Begin VB.ComboBox CboGrupo 
            Height          =   315
            ItemData        =   "Status.frx":1515
            Left            =   1800
            List            =   "Status.frx":15CD
            TabIndex        =   27
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Status.frx":18A1
            Left            =   120
            List            =   "Status.frx":18A3
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1200
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   40028
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora de Atención:"
            Height          =   195
            Left            =   1800
            TabIndex        =   25
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta con:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos de la Consulta"
         Height          =   5055
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   12255
         Begin MSComctlLib.ListView LstDatosConsulta 
            Height          =   4455
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   7858
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No. Historia"
               Object.Width           =   2205
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "IdEstatus"
               Object.Width           =   9
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "No. Cédula"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Apellido(s)"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Nombre(s)"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Descripción"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Fecha"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Hora Llegada"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Hora Atención"
               Object.Width           =   2293
            EndProperty
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Left            =   8520
            TabIndex        =   35
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Atención"
            Height          =   195
            Left            =   10920
            TabIndex        =   34
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Llegada"
            Height          =   195
            Left            =   9600
            TabIndex        =   33
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Left            =   6600
            TabIndex        =   32
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s)"
            Height          =   195
            Left            =   4800
            TabIndex        =   31
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido(s)"
            Height          =   195
            Left            =   3120
            TabIndex        =   30
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula"
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Historia"
            Height          =   195
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3840
         TabIndex        =   8
         Top             =   7200
         Width           =   8535
         Begin VB.Timer Timer1 
            Left            =   6720
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   7440
            TabIndex        =   7
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
            MICON           =   "Status.frx":18A5
            PICN            =   "Status.frx":18C1
            PICH            =   "Status.frx":1A8A
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
            Left            =   5760
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
            MICON           =   "Status.frx":1CBF
            PICN            =   "Status.frx":1CDB
            PICH            =   "Status.frx":1F71
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
            Left            =   5160
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
            MICON           =   "Status.frx":21D0
            PICN            =   "Status.frx":21EC
            PICH            =   "Status.frx":2481
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
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Incluir en la lista de espera"
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Incluir en lista"
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
            MICON           =   "Status.frx":26DD
            PICN            =   "Status.frx":26F9
            PICH            =   "Status.frx":2982
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnProgramarCita 
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            ToolTipText     =   "Incluir en la lista de espera"
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Programar Cita"
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
            MICON           =   "Status.frx":2DB8
            PICN            =   "Status.frx":2DD4
            PICH            =   "Status.frx":3209
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
            Left            =   3720
            TabIndex        =   36
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
            MICON           =   "Status.frx":363E
            PICN            =   "Status.frx":365A
            PICH            =   "Status.frx":37FE
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
Attribute VB_Name = "FrmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDStatus As New ADODB.Recordset 'TABLA DE PACIENTES
Dim BDST As New ADODB.Recordset 'TABLA ESTATUS
Dim BDST1 As New ADODB.Recordset 'TABLA ESTATUS para comprobar cambios
Dim BDST2 As New ADODB.Recordset
Dim camb
Dim ConnectionString As String
Public Cbus As String
Dim RsCargarTablaStatus As New ADODB.Recordset
Dim CedPac, IdPac2 As String
Dim IdEstat As String
Dim RsConsultar As New ADODB.Recordset
Public IdMax

Private Sub Buscar()

If Trim(TxtBuscar.Text) = "" Then Exit Sub

If rs.State = 1 Then rs.Close

CSql = "Select * From Paciente Where CedulaP = " & Val(TxtBuscar.Text) & " or NombreP like '%" & TxtBuscar.Text & "%' or ApellidoP like '%" & TxtBuscar.Text & "%'"
Set rs = CrearRS(CSql)
  
If rs.EOF Then
    MsgBox "No Existe el registro", vbCritical + vbOKOnly, "Mensaje"
    Exit Sub
   
Else
'    If rs.Fields("Tipo").Value = 0 And rs.Fields("Fecha_RegP").Value = DateTime.Date Then
        If Not rs.EOF Or Not rs.BOF Then
            TxtHistoria.Caption = rs.Fields("Historia").Value
            TxtApellido.Caption = rs.Fields("ApellidoP").Value
            TxtNombre.Caption = rs.Fields("NombreP").Value
            TxtCedula.Caption = rs.Fields("CedulaP").Value
            IdPac1 = rs.Fields("IdPaciente").Value
            CboGrupo.Text = rs.Fields("HoraAtencion").Value
            
            Me.Caption = "Estatus del Paciente - Id: " & IdPac1
        Else
            TxtCedula.Caption = "":        TxtHistoria.Caption = "":        TxtNombre.Caption = "":       TxtApellido.Caption = ""
        End If
'    ElseIf rs.Fields("Tipo").Value = 0 And rs.Fields("Fecha_RegP").Value <> DateTime.Date Then
'        Dim RsActualizar As New ADODB.Recordset
'        CSql = "Select * From Paciente Where Cedulap='" & Val(TxtBuscar.Text) & "'"
'        Set RsActualizar = CrearRS(CSql)
'
'        If RsActualizar.RecordCount > 0 Then
'            RsActualizar.Fields("Tipo").Value = 1
'            RsActualizar.Fields("Grupo").Value = CboGrupo.Text
'            RsActualizar.Fields("Fecha_Inicio").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
'            RsActualizar.Update
'        End If
'
'            TxtHistoria.Caption = rs.Fields("Historia").Value
'            TxtApellido.Caption = rs.Fields("ApellidoP").Value
'            TxtNombre.Caption = rs.Fields("NombreP").Value
'            TxtCedula.Caption = rs.Fields("CedulaP").Value
'            IdPac1 = rs.Fields("IdPaciente").Value
'            CboGrupo.Text = rs.Fields("Grupo").Value
'
'    End If
End If

End Sub

Sub CargaStatu()
If Not (BDStatus.EOF) Then

If BDStatus.Fields("Historia").Value <> "" Then TxtHistoria.Caption = BDStatus.Fields("Historia").Value
If BDStatus.Fields("CedulaP").Value <> "" Then TxtCedula.Caption = BDStatus.Fields("CedulaP").Value
If BDStatus.Fields("NombreP").Value <> "" Then TxtNombre.Caption = BDStatus.Fields("NombreP").Value
If BDStatus.Fields("ApellidoP").Value <> "" Then TxtApellido.Caption = BDStatus.Fields("ApellidoP").Value
IdPac1 = BDStatus.Fields("IdPaciente").Value
Me.Caption = "Estatus del Paciente - Paciente: " & IdPac1
Else
    TxtCedula.Caption = "":        TxtHistoria.Caption = "":        TxtNombre.Caption = "":       TxtApellido.Caption = ""
End If

End Sub

Private Sub BtnAnterior_Click()
BDStatus.MovePrevious
If Not BDStatus.BOF Then
Call CargaStatu
Else
BDStatus.MoveLast
Call CargaStatu
End If
End Sub

Private Sub BtnBuscar_Click()
Buscar
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnEliminar_Click()
' Si hay un Item seleccionado ...
If Not LstDatosConsulta.SelectedItem Is Nothing Then
   
   'Pregunta si lo quiere eliminar
   If MsgBox("Deseas Eliminar al Paciente de la lista de espera??", vbQuestion + vbYesNo, "Rt Complete") = vbYes Then
        'Elimina el elemento seleccionado ( SelectedItem.Index )
        LstDatosConsulta.ListItems.Remove (LstDatosConsulta.SelectedItem.Index)
  End If
End If
End Sub

Private Sub BtnGuardarActualizar_Click()
If Combo1.ListIndex = -1 Then
    f = "Area Adonde se dirige"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If
If TxtCedula.Caption = "" Then
    f = "Cédula de Identidad"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If
If TxtNombre.Caption = "" Then
    f = "Nombre del Paciente"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If

If TxtApellido.Caption = "" Then
    f = "Apellido del Paciente"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If

If CboGrupo.Text = "00:00" Then
    Msg = "El Paciente no Posee Hora de Atención Programada, Debe de Asignarle una Hora"
    MsgBox Msg, vbCritical + vbOKOnly, "Error al Guardar"
    CboGrupo.SetFocus
    Exit Sub
End If

Dim RsGuardar As New ADODB.Recordset
Dim RsConsultar As New ADODB.Recordset

CSql = "Select * From Paciente where CedulaP='" & TxtCedula.Caption & "'"
Set RsConsultar = CrearRS(CSql)

If RsConsultar.Fields("Tipo").Value = 0 And RsConsultar.Fields("Fecha_RegP").Value = DateTime.Date Then

    CSql = "Select * From Estatus"
    Set RsGuardar = CrearRS(CSql)
    RsGuardar.AddNew
    RsGuardar.Fields("IdPaciente").Value = IdPac1
    RsGuardar.Fields("MotivoV").Value = Combo1.ItemData(Combo1.ListIndex)
    RsGuardar.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
    RsGuardar.Fields("Hora").Value = Format(DateTime.Time, "hh:mm:ss AMPM")
    RsGuardar.Fields("IdUsuario").Value = IdUser
    RsGuardar.Fields("Descripcion").Value = Combo1.List(Combo1.ListIndex)
    RsGuardar.Fields("Grupo").Value = Trim(CboGrupo.Text)
    RsGuardar.Fields("CitaProgramada").Value = 0
    RsGuardar.Update
    
ElseIf RsConsultar.Fields("Tipo").Value = 1 And RsConsultar.Fields("Fecha_RegP").Value <> DateTime.Date Then

    CSql = "Select * From Estatus"
    Set RsGuardar = CrearRS(CSql)
    RsGuardar.AddNew
    RsGuardar.Fields("IdPaciente").Value = IdPac1
    RsGuardar.Fields("MotivoV").Value = Combo1.ItemData(Combo1.ListIndex)
    RsGuardar.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
    RsGuardar.Fields("Hora").Value = Format(DateTime.Time, "hh:mm:ss AMPM")
    RsGuardar.Fields("IdUsuario").Value = IdUser
    RsGuardar.Fields("Descripcion").Value = Combo1.List(Combo1.ListIndex)
    RsGuardar.Fields("Grupo").Value = Trim(CboGrupo.Text)
    RsGuardar.Fields("CitaProgramada").Value = 0
    RsGuardar.Update
ElseIf RsConsultar.Fields("Tipo").Value = 0 And RsConsultar.Fields("Fecha_RegP").Value <> DateTime.Date Then
        Dim RsActualizar As New ADODB.Recordset
        CSql = "Select * From Paciente Where CedulaP='" & Val(TxtCedula.Caption) & "'"
        Set RsActualizar = CrearRS(CSql)

        If RsActualizar.RecordCount > 0 Then
            RsActualizar.Fields("Tipo").Value = 1
            RsActualizar.Fields("Grupo").Value = Mid(CboGrupo.Text, 1, 3) & "00" & Mid(CboGrupo.Text, 6)
            RsActualizar.Fields("HoraAtencion").Value = Trim(CboGrupo.Text)
            RsActualizar.Fields("Fecha_Inicio").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
            RsActualizar.Update
        End If

        CSql = "Select * From Estatus"
        Set RsGuardar = CrearRS(CSql)
        RsGuardar.AddNew
        RsGuardar.Fields("IdPaciente").Value = IdPac1
        RsGuardar.Fields("MotivoV").Value = Combo1.ItemData(Combo1.ListIndex)
        RsGuardar.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        RsGuardar.Fields("Hora").Value = Format(DateTime.Time, "hh:mm:ss AMPM")
        RsGuardar.Fields("IdUsuario").Value = IdUser
        RsGuardar.Fields("Descripcion").Value = Combo1.List(Combo1.ListIndex)
        RsGuardar.Fields("Grupo").Value = Trim(CboGrupo.Text)
        RsGuardar.Fields("CitaProgramada").Value = 0
        RsGuardar.Update

End If


'Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
'MsgBox Msg, vbInformation + vbOKOnly, "Actualización Estatus"
'EnviarAlHosting
'EnviarRegPendiente
Call Blanqueo
Call CargaTablaStatus

End Sub

Sub EnviarRegPendiente()

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "Select Max(IdEstatus) as IdMax From Estatus"
Set RsConsultar = CrearRS(CSql)

IdMax = RsConsultar.Fields("IdMax").Value

a = 0
CSql = "INSERT into Estatus(IdEstatus, IdPaciente, MotivoV, Fecha, Hora, IdUsuario, Descripcion, Grupo, CitaProgramada) " & _
"Values ('" & IdMax & "','" & IdPac1 & "','" & Combo1.ItemData(Combo1.ListIndex) & "','" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'," & _
"'" & Format(DateTime.Time, "hh:mm:ss AMPM") & "','" & IdUser & "','" & Combo1.List(Combo1.ListIndex) & "','" & Trim(CboGrupo.Text) & "','" & a & "')"
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Estatus"
RsRegPendiente.Fields("Tabla").Value = "Estatus"
RsRegPendiente.Fields("Condicional").Value = "IdEstatus=" & IdMax
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"


End Sub

Sub EnviarAlHosting()

On Error GoTo salir
ConectarHosting


CSql = "Select Max(IdEstatus) as IdMax From Estatus"
Set RsConsultar = CrearRS(CSql)

IdMax = RsConsultar.Fields("IdMax").Value

CSql = "Select * From Estatus"
Set RsWeb = CrearRsWeb(CSql)
RsWeb.AddNew
RsWeb.Fields("IdEstatus").Value = IdMax
RsWeb.Fields("IdPaciente").Value = IdPac1
RsWeb.Fields("MotivoV").Value = Combo1.ItemData(Combo1.ListIndex)
RsWeb.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
RsWeb.Fields("Hora").Value = Format(DateTime.Time, "hh:mm:ss")
RsWeb.Fields("IdUsuario").Value = IdUser
RsWeb.Fields("Descripcion").Value = Combo1.List(Combo1.ListIndex)
RsWeb.Fields("Grupo").Value = Trim(CboGrupo.Text)
RsWeb.Fields("CitaProgramada").Value = 0
RsWeb.Update
RsWeb.Close


CSql = "Select * From Estatus Where IdPaciente=" & IdPac1 & " And Fecha='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'"
Set RsWeb = CrearRsWeb(CSql)

If RsWeb.RecordCount > 0 Then
    Msg = "Actualización Completada en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
Else
    Msg = "No se realizo la Actualización en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
End If



salir:
If WebCnn.State = 0 Then
    EnviarAlHosting
Else
    GoTo f
End If

f:
If WebCnn.State = 1 Then
    WebCnn.Close
Else
    If Err.Number <> 0 Then MsgBox Err.Description
    Exit Sub
End If

End Sub

Private Sub BtnProgramarCita_Click()
GuardarCitaProgramada
End Sub

Sub GuardarCitaProgramada()

If TxtCedula.Caption = "" And TxtNombre.Caption = "" And TxtApellido.Caption = "" And TxtHistoria.Caption = "" Then
    Msg = "Debe de completar el formulario para programarle la cita medica, seleccione al paciente!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error al Guardar Cita Programada"
    Exit Sub
End If

If Combo1.ListIndex = -1 Then
    f = "Area Adonde se dirige"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If
If TxtCedula.Caption = "" Then
    f = "Cédula de Identidad"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If
If TxtNombre.Caption = "" Then
    f = "Nombre del Paciente"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If

If TxtApellido.Caption = "" Then
    f = "Apellido del Paciente"
    Msg = "Debe de completar el formulario, Falta el campo: " & f
    MsgBox Msg, vbCritical, "Error al Guardar"
    Exit Sub
End If

If CboGrupo.Text = "" Then
    Msg = "Debe de Seleccionar la hora de cita!!"
    MsgBox Msg, vbOKOnly + vbCritical, "Error al Guardar"
    CboGrupo.SetFocus
    Exit Sub
End If

Dim RsGuardar As New ADODB.Recordset
CSql = "Select * From Estatus"
Set RsGuardar = CrearRS(CSql)
RsGuardar.AddNew
RsGuardar.Fields("IdPaciente").Value = IdPac1
RsGuardar.Fields("MotivoV").Value = Combo1.ItemData(Combo1.ListIndex)
RsGuardar.Fields("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
RsGuardar.Fields("Hora").Value = Format("00:00:00", "hh:mm:ss AMPM")
RsGuardar.Fields("IdUsuario").Value = IdUser
RsGuardar.Fields("Descripcion").Value = Combo1.List(Combo1.ListIndex)
RsGuardar.Fields("Grupo").Value = Trim(CboGrupo.Text)
RsGuardar.Fields("CitaProgramada").Value = 1
RsGuardar.Update

'Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
'MsgBox Msg, vbInformation + vbOKOnly, "Actualización Estatus"
''EnviarAlHostingCitaProgramada
'EnviarRegPendienteCitaProgramada
Call Blanqueo
Call CargaTablaStatus


End Sub

Sub EnviarRegPendienteCitaProgramada()
CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "Select Max(IdEstatus) as IdMax From Estatus"
Set RsConsultar = CrearRS(CSql)

IdMax = RsConsultar.Fields("IdMax").Value


a = 1
CSql = "INSERT into Estatus(IdEstatus, IdPaciente, MotivoV, Fecha, Hora, IdUsuario, Descripcion, Grupo, CitaProgramada) " & _
"Values ('" & IdMax & "','" & IdPac1 & "','" & Combo1.ItemData(Combo1.ListIndex) & "','" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'," & _
"'" & Format("00:00:00", "hh:mm:ss AMPM") & "','" & IdUser & "','" & Combo1.List(Combo1.ListIndex) & "','" & Trim(CboGrupo.Text) & "','" & a & "')"
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Estatus"
RsRegPendiente.Fields("Tabla").Value = "Estatus"
RsRegPendiente.Fields("Condicional").Value = "IdEstatus=" & IdMax
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"
End Sub

Sub EnviarAlHostingCitaProgramada()
On Error GoTo salir
ConectarHosting


CSql = "Select Max(IdEstatus) as IdMax From Estatus"
Set RsConsultar = CrearRS(CSql)

IsMax = RsConsultar.Fields("IdMax").Value

CSql = "Select * From Estatus"
Set RsWeb = CrearRsWeb(CSql)
RsWeb.AddNew
RsWeb.Fields("IdEstatus").Value = IdMax
RsWeb.Fields("IdPaciente").Value = IdPac1
RsWeb.Fields("MotivoV").Value = Combo1.ItemData(Combo1.ListIndex)
RsWeb.Fields("Fecha").Value = Format(DTPicker1.Value, "mm/dd/yyyy")
RsWeb.Fields("Hora").Value = Format("00:00:00", "hh:mm:ss AMPM")
RsWeb.Fields("IdUsuario").Value = IdUser
RsWeb.Fields("Descripcion").Value = Combo1.List(Combo1.ListIndex)
RsWeb.Fields("Grupo").Value = Trim(CboGrupo.Text)
RsWeb.Fields("CitaProgramada").Value = 1
RsWeb.Update

CSql = "Select * From Estatus Where IdPaciente=" & IdPac1 & " And Fecha='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'"
Set RsWeb = CrearRsWeb(CSql)

If RsWeb.RecordCount > 0 Then
    Msg = "Actualización Completada en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
Else
    Msg = "No se realizo la Actualización en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Operación Exitosa"
    RsWeb.Close
    'WebCnn.Close
End If



salir:
If WebCnn.State = 0 Then
    EnviarAlHosting
Else
    GoTo f
End If

f:
If WebCnn.State = 1 Then
    WebCnn.Close
Else
    If Err.Number <> 0 Then MsgBox Err.Description
    Exit Sub
End If

End Sub

Private Sub BtnSiguiente_Click()
BDStatus.MoveNext
If Not BDStatus.EOF Then
Call CargaStatu
Else
BDStatus.MoveFirst
Call CargaStatu
End If
End Sub

Sub Blanqueo()
 TxtApellido.Caption = ""
 TxtNombre.Caption = ""
 TxtCedula.Caption = ""
 TxtHistoria.Caption = ""
 Combo1.ListIndex = -1
 IdPac1 = ""
 
End Sub

Private Sub elimina()
    On Error Resume Next
    DataGrid1.Col = 2
    Nomb = DataGrid1.Text
    DataGrid1.Col = 3
    apel = DataGrid1.Text
    DataGrid1.Col = 1
    cedula = DataGrid1.Text
    DataGrid1.Col = 5
    descr = DataGrid1.Text
    Msg = "Esta seguro de borrar de la lista de espera al paciente: " & Chr(13) & Nomb & " " & apel & Chr(13) & " Dirigido al area de: " & Chr(13) & descr & Chr(13) & "esta seguro ? "
    f = MsgBox(Msg, vbYesNo, "Borrar Paciente")
    If f = 6 Then
       ' MsgBox cedula & Chr(13) & descr
        CSql = "select idpaciente from paciente where cedula = " & cedula
        Set BDST2 = CrearRS(CSql)
        idpa1 = BDST2.Fields("idpaciente")
        BDST2.Close
        CSql = "Delete from estatus where idpaciente = " & idpa1 & " and descripcion = '" & descr & "'"
        Set BDST2 = CrearRS(CSql)
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
Call elimina
KeyCode = 0
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnGuardarActualizar.SetFocus
End If
End Sub

Private Sub DTPicker1_Change()
Call CargaTablaStatus
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Form_Load()
Centrar Me

Combo1.AddItem "Administración", 0
Combo1.ItemData(0) = 5
Combo1.AddItem "Dirección Medica", 1
Combo1.ItemData(1) = 3
Combo1.AddItem "Nutrición / Lcda. Fernandez", 2
Combo1.ItemData(2) = 0
Combo1.AddItem "Nutrición / Lcda. Passarelli", 3
Combo1.ItemData(3) = 0
Combo1.AddItem "Oncología / Dr. Suisiña", 4
Combo1.ItemData(4) = 4
Combo1.AddItem "Oncología / Dra. Romero", 5
Combo1.ItemData(5) = 4
Combo1.AddItem "Psicología / Psic. Faria", 6
Combo1.ItemData(6) = 1
Combo1.AddItem "Psicología / Psic. Añez", 7
Combo1.ItemData(7) = 1
Combo1.AddItem "Radioterapia", 8
Combo1.ItemData(8) = 2
Combo1.AddItem "Reubicación", 9
Combo1.ItemData(9) = 2
Combo1.AddItem "Simulación", 10
Combo1.ItemData(10) = 2


DTPicker1.Value = Now
'<<<<<<<CARGA DE BD>>>>>>>>>>
Call CargaTablaStatus
'<<<<<<<CARGA DE BD>>>>>>>>>>
camb = 0
   
If BDStatus.State = adStateOpen Then BDStatus.Close
CSql = "SELECT * FROM Paciente"
Set BDStatus = CrearRS(CSql)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub LstDatosConsulta_Click()
If LstDatosConsulta.ListItems.Count = 0 Then Exit Sub
CedPac = Replace(LstDatosConsulta.SelectedItem.ListSubItems(2).Text, ".", "")
IdEstat = LstDatosConsulta.SelectedItem.ListSubItems(1).Text
If CedPac <> "" Then
    
    CSql = "Select * From Paciente where CedulaP='" & CedPac & "'"
    Set RsConsultar = CrearRS(CSql)
    If RsConsultar.RecordCount > 0 Then
        IdPac2 = RsConsultar.Fields("IdPaciente").Value
        IdPac1 = RsConsultar.Fields("IdPaciente").Value
        TxtCedula.Caption = RsConsultar.Fields("CedulaP").Value
        TxtApellido.Caption = RsConsultar.Fields("ApellidoP").Value
        TxtNombre.Caption = RsConsultar.Fields("NombreP").Value
        TxtHistoria.Caption = RsConsultar.Fields("Historia").Value
        CboGrupo.Text = RsConsultar.Fields("Grupo").Value
        Me.Caption = "Estatus del Paciente - Id: " & IdPac1
    End If
End If

CSql = "Select * From Estatus where IdEstatus='" & IdEstat & "'"
Set RsConsultar = CrearRS(CSql)
If RsConsultar.RecordCount > 0 Then
    If RsConsultar.Fields("CitaProgramada").Value = 1 Then
        Msg = "El Paciente: " & Trim(TxtNombre.Caption) & ", " & Trim(TxtApellido.Caption) & " tiene cita programada con " & Trim(RsConsultar.Fields("Descripcion").Value) & "." & Chr(13) & _
              "Asignele la hora de llegada Precionando el Botón de Si para que sea atendido por " & Trim(RsConsultar.Fields("Descripcion").Value)
        Respuesta = MsgBox(Msg, vbInformation + vbYesNo, "Cita Programada")
        If Respuesta = vbYes Then
            Dim RsActualizar As New ADODB.Recordset
            CSql = "Update Estatus set Hora='" & Format(DateTime.Time, "hh:mm:ss AMPM") & "', CitaProgramada='0' Where IdEstatus='" & IdEstat & "'"
            Set RsActualizar = CrearRS(CSql)
            
            Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
            MsgBox Msg, vbInformation + vbOKOnly, "Actualización Estatus"
                       
            ActualizarRegPendiente
            Blanqueo
            CargaTablaStatus
        End If
    End If
End If

If RsConsultar.State = 1 Then RsConsultar.Close: Exit Sub
End Sub

Sub ActualizarRegPendiente()

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

a = 0
CSql = "UPDATE Estatus set Hora='" & Format(DateTime.Time, "hh:mm:ss AMPM") & "', CitaProgramada='0' Where IdEstatus='" & IdEstat & "'"
sentencia = Replace(CSql, "'", "(varCSP)")


CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Estatus del Paciente"
RsRegPendiente.Fields("Tabla").Value = "Estatus"
RsRegPendiente.Fields("Condicional").Value = "IdEstatus = '" & IdEstat & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = sentencia
RsRegPendiente.Update

Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"
End Sub
            
Private Sub Timer1_Timer()
On Error Resume Next
CSql = "SELECT Count(*) as RecordCount FROM estatus "
Set BDST1 = CrearRS(CSql)
If BDST1.EOF Then
    BDST1.Close
    Exit Sub
End If
tuy = BDST1.Fields("recordcount")
If tuy <> camb Then
    Call CargaTablaStatus
    camb = tuy
End If

BDST1.Close
On Error GoTo 0
End Sub

Sub CargaTablaStatus()
Fecha1 = Format(DTPicker1.Value, "dd/mm/yyyy")
Fecha2 = Format(DTPicker1.Value, "dd/mm/yyyy")

'Fecha = Format(DTPicker1.Value, "dd/mm/yyyy")

If RsCargarTablaStatus.State = 1 Then RsCargarTablaStatus.Close
CSql = "Select * From EstatusPac Where Fecha >='" & Fecha1 & "' And Fecha <='" & Fecha2 & "'  Order By hora, Grupo"

'CSql = "Select * From EstatusPac Where Fecha ='" & Fecha & "' Order By hora"
Set RsCargarTablaStatus = CrearRS(CSql)
LstDatosConsulta.ListItems.Clear
If RsCargarTablaStatus.RecordCount = 0 Then Exit Sub
Do While Not RsCargarTablaStatus.EOF
    If RsCargarTablaStatus.Fields("CitaProgramada").Value = 0 Then
        With LstDatosConsulta
             i = i + 1
            .ListItems.Add , , RsCargarTablaStatus.Fields("Historia").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("IdEstatus").Value
            .ListItems(i).ListSubItems.Add , , Format(RsCargarTablaStatus.Fields("Cedulap").Value, "#,##")
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Apellidop").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Nombrep").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Descripcion").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Fecha").Value
            .ListItems(i).ListSubItems.Add , , Format(RsCargarTablaStatus.Fields("Hora").Value, "hh:mm:ss AMPM")
            If Not IsNull(RsCargarTablaStatus.Fields("Grupo").Value) Then
                .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Grupo").Value
            Else
                .ListItems(i).ListSubItems.Add , , ""
            End If
            
            .ListItems(i).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(1).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(2).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(3).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(4).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(5).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(6).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(7).ForeColor = vbBlack
            .ListItems(i).ListSubItems.Item(8).ForeColor = vbBlack
        End With
    Else
        With LstDatosConsulta
             i = i + 1
            
            .ListItems.Add , , RsCargarTablaStatus.Fields("Historia").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("IdEstatus").Value
            .ListItems(i).ListSubItems.Add , , Format(RsCargarTablaStatus.Fields("Cedulap").Value, "#,##")
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Apellidop").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Nombrep").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Descripcion").Value
            .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Fecha").Value
            .ListItems(i).ListSubItems.Add , , Format(RsCargarTablaStatus.Fields("Hora").Value, "hh:mm:ss AMPM")
            If Not IsNull(RsCargarTablaStatus.Fields("Grupo").Value) Then
                .ListItems(i).ListSubItems.Add , , RsCargarTablaStatus.Fields("Grupo").Value
            Else
                .ListItems(i).ListSubItems.Add , , ""
            End If
            .ListItems(i).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(2).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(3).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(4).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(5).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(6).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(7).ForeColor = vbRed
            .ListItems(i).ListSubItems.Item(8).ForeColor = vbRed
            
        End With
    End If
    RsCargarTablaStatus.MoveNext
Loop

End Sub

Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    TxtNombre.Set.SetFocus
'End If
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Buscar
End If
End Sub

Private Sub TxtCedula_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    TxtApellido.SetFocus
'End If
End Sub

Private Sub TxtHistoria_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    TxtHistoria.SetFocus
'End If
End Sub

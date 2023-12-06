VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmLlamadoPaciente 
   BackColor       =   &H00EAEFEF&
   Caption         =   "Llamado de Paciente"
   ClientHeight    =   2220
   ClientLeft      =   4320
   ClientTop       =   5640
   ClientWidth     =   8970
   Icon            =   "LlamaAdmi.frx":0000
   LinkTopic       =   "Form54"
   MDIChild        =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   8970
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos Paciente"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3480
         Top             =   120
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   8535
         Begin ChamaleonButton.ChameleonBtn BtnNuevo 
            Height          =   495
            Left            =   3720
            TabIndex        =   8
            ToolTipText     =   "Nuevo"
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
            MICON           =   "LlamaAdmi.frx":1002
            PICN            =   "LlamaAdmi.frx":101E
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
               Name            =   "Calibri"
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
            TabIndex        =   1
            Text            =   "Busqueda"
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Cedula"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   495
            Left            =   2880
            TabIndex        =   7
            ToolTipText     =   "Buscar"
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
            MICON           =   "LlamaAdmi.frx":15B8
            PICN            =   "LlamaAdmi.frx":15D4
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
            Height          =   495
            Left            =   6120
            TabIndex        =   10
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
            MICON           =   "LlamaAdmi.frx":1839
            PICN            =   "LlamaAdmi.frx":1855
            PICH            =   "LlamaAdmi.frx":1AEA
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
            Left            =   6720
            TabIndex        =   11
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
            MICON           =   "LlamaAdmi.frx":1D46
            PICN            =   "LlamaAdmi.frx":1D62
            PICH            =   "LlamaAdmi.frx":1FF8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnListado 
            Height          =   495
            Left            =   4800
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Ver Listado"
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
            MICON           =   "LlamaAdmi.frx":2257
            PICN            =   "LlamaAdmi.frx":2273
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
            Height          =   495
            Left            =   7440
            TabIndex        =   18
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
            MICON           =   "LlamaAdmi.frx":24D4
            PICN            =   "LlamaAdmi.frx":24F0
            PICH            =   "LlamaAdmi.frx":26B9
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
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   6600
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   4440
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin ChamaleonButton.ChameleonBtn BtnAyuda 
         Height          =   375
         Left            =   8160
         TabIndex        =   19
         Top             =   120
         Width           =   495
         _ExtentX        =   873
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
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "LlamaAdmi.frx":28EE
         PICN            =   "LlamaAdmi.frx":290A
         PICH            =   "LlamaAdmi.frx":2BAC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido(s)"
         Height          =   195
         Left            =   6600
         TabIndex        =   16
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s)"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cédula"
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Historia"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmLlamadoPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BdAdm As New ADODB.Recordset 'TABLA DE PACIENTES
Dim BdAdm1 As New ADODB.Recordset 'TABLA ESTATUS
Dim BdAdm2 As New ADODB.Recordset 'TABLA ESTATUS para comprobar cambios
Dim camb


Private Sub Buscar()
If Trim(TxtBuscar.Text) = "" Then Exit Sub
If BdAdm.State = adStateOpen Then BdAdm.Close

    If Option1(0).Value = True Then Cbus = "Cedula = " & Val(TxtBuscar.Text)
    If Option1(1).Value = True Then Cbus = "Nombre = '" & TxtBuscar.Text & "'"
  
  CSql = "select * from Paciente where " & Cbus
  BdAdm.Open CSql, Cnn, , , adCmdText
  
 If BdAdm.EOF Then
    MsgBox "No Existe el registro"
    SQL = "select * from Paciente "
    BdAdm.Close
    
    BdAdm.Open SQL, Cnn, , , adCmdText
    BdAdm.MoveFirst

   End If

  Call cargaAdmi
 
End Sub
 Sub cargaAdmi()
   If Not (BdAdm.EOF) Then

    If BdAdm.Fields("Historia") <> "" Then Text2.Text = BdAdm.Fields("Historia") ' Else Text.Text = ""
    If BdAdm.Fields("Cedula") <> "" Then Text3.Text = BdAdm.Fields("Cedula") 'Else Text.Text = ""
    If BdAdm.Fields("Nombre") <> "" Then Text4.Text = BdAdm.Fields("Nombre") 'Else Text.Text = ""
    If BdAdm.Fields("Apellido") <> "" Then Text5.Text = BdAdm.Fields("Apellido") 'Else Text.Text = ""
    If BdAdm.Fields("idpaciente") <> "" Then IdPac1 = BdAdm.Fields("idpaciente")
      Else
        TxtBuscar.Text = "":        Text2.Text = "":        Text3.Text = "":       Text6.Text = "":        Text5.Text = ""
    End If

 End Sub
 

Sub Blanqueo()
 TxtBuscar.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
  
End Sub

Private Sub BtnAnterior_Click()
BdAdm.MovePrevious
If Not BdAdm.BOF Then
Call cargaAdmi
Else
BdAdm.MoveLast
Call cargaAdmi
End If
End Sub

Private Sub BtnBuscar_Click()
Buscar
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnListado_Click()
ModulO = 5
FrmListaEspera.Show
TxtBuscar.Text = Cedul
Call Buscar
End Sub

Private Sub BtnNuevo_Click()
Blanqueo
End Sub

Private Sub BtnSiguiente_Click()
BdAdm.MoveNext
If Not BdAdm.EOF Then
Call cargaAdmi
Else
BdAdm.MoveFirst
Call cargaAdmi
End If
End Sub

Private Sub Form_Activate()
TxtBuscar.SetFocus
End Sub

Private Sub Form_Load()
Me.Height = 2730
Me.Width = 9090

Centrar Me
camb = 0
ModulO = 5
If BdAdm.State = adStateOpen Then BdAdm.Close
CSql = "SELECT * FROM Paciente"
BdAdm.CursorType = adOpenKeyset
BdAdm.LockType = adLockOptimistic
BdAdm.CursorLocation = adUseClient
BdAdm.Open CSql, Cnn, , , adCmdText
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Buscar
End Sub
Private Sub TxtBuscar_Change()

Dim StrText, Chaa, pru As String
 Dim i  As Variant
 StrText = ""
 Chaa = ""
  Chaa = UCase(Mid(TxtBuscar.Text, 1, 1))
  StrText = Chaa
  For i = 2 To Len(TxtBuscar.Text)
    pru = LCase(Mid(TxtBuscar.Text, i, 1))
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

 TxtBuscar.Text = StrText
 TxtBuscar.SelStart = Len(TxtBuscar.Text)
 Cambio = 1
End Sub
Private Sub Timer2_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
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

Private Sub Timer1_Timer()
CSql = "SELECT Count(*) as RecordCount FROM estatus "
BdAdm2.Open CSql, Cnn
If BdAdm2.EOF Then
BdAdm2.Close
Exit Sub
End If
tuy = BdAdm2.Fields("recordcount")
If tuy <> camb Then
Call CargaTablaStatus
camb = tuy
End If

BdAdm2.Close

End Sub

Sub CargaTablaStatus()

Adodc1.Refresh
DataGrid1.Refresh
End Sub


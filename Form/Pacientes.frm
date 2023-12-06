VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Pacientes"
   ClientHeight    =   3945
   ClientLeft      =   2910
   ClientTop       =   2835
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   ScaleHeight     =   3945
   ScaleWidth      =   6720
   Begin VB.CommandButton Command3 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   3600
      TabIndex        =   30
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   28
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "medico"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4080
      TabIndex        =   14
      Top             =   2520
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "remitente"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "fecha_ing"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16711681
      CurrentDate     =   39384
   End
   Begin VB.TextBox Text8 
      DataField       =   "fax"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "tlfs"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Text            =   "Text7"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "sexo"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Pacientes.frx":0000
      Left            =   2400
      List            =   "Pacientes.frx":000A
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "profesion"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      DataField       =   "DIRECCION"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      DataField       =   "edad"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "historia"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "ci"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombres"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "apellidos"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=SQL Server3"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SQL Server3"
      OtherAttributes =   ""
      UserName        =   "robertoportatil"
      Password        =   "123456"
      RecordSource    =   "mtro_cli"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "fecha_nac"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16711681
      CurrentDate     =   39384
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   6720
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Médico Tratante"
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Médico Remitente"
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fech. Ingreso"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fax/Movil"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Telefono"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sexo"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Profesión"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dirección"
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Edad"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nacimiento"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Historia Médica"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ced Identidad"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apellido"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cbus1

Sub cargar()
'carga las profesiones
Dim rec As New ADODB.Recordset
consultasql = "select profesion from profesion order by profesion ASC;"
rec.Open consultasql, cadenaconexioN
rec.MoveFirst
Do Until rec.EOF
Combo1.AddItem rec.Fields("profesion")
rec.MoveNext
Loop

'carga las medicos remitentes
Dim rec1 As New ADODB.Recordset
consultasql = "select medico from medicor order by medico ASC;"
rec1.Open consultasql, cadenaconexioN
rec1.MoveFirst
Do Until rec1.EOF
Combo3.AddItem rec1.Fields("medico")
rec1.MoveNext
Loop

'carga las medicos remitentes
Dim rec2 As New ADODB.Recordset
consultasql = "select medico from medicot order by medico ASC;"
rec2.Open consultasql, cadenaconexioN
rec2.MoveFirst
Do Until rec2.EOF
Combo4.AddItem rec2.Fields("medico")
rec2.MoveNext
Loop


End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Form_Load()
Call cargar
End Sub

Private Sub Text5_GotFocus()
If Text5.Text = "" Then
fna = DTPicker1.Value
fho = Date
Text5.Text = DateDiff("yyyy", fna, fho)
End If
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
On Error GoTo fg

If KeyAscii = 13 Then
cbus = "[ci] like '*" & Text9.Text & "*'"
    If cbus1 <> Text9.Text Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Find cbus, 1, , 0
        cbus1 = Text9.Text
    Else
        Adodc1.Recordset.Find cbus, 1, 1, 0
    End If

End If
On Error GoTo 0
Exit Sub

fg:
If Err.Number = 3021 Then
msg = "Se ha llegado al final de la base de datos"
MsgBox msg
Adodc1.Recordset.MoveFirst
cbus1 = ""
End If

End Sub

VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmBeam 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar TXT"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   Icon            =   "FrmBeam.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Sitio Anatómico de Tartamiento"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBeam.frx":1002
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscarArchivo 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Buscar Archivo"
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
         MICON           =   "FrmBeam.frx":101E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "FrmBeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsD As New ADODB.Recordset
Private Sub BtnBuscarArchivo_Click()
BeamDescripcion = Combo1.Text
If Combo1.Text = "" Then
    MsgBox "Seleccione el Sitio Anatómico de tartamiento", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If
Unload Me
End Sub

Private Sub BtnCerrar_Click()
If Combo1.Text = "" Then
    MsgBox "Seleccione el Sitio Anatómico de tartamiento", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If
Unload Me
End Sub

Private Sub Form_Load()

CSql = "SELECT TECNICA FROM TECNICA WHERE IDPACIENTE = '" & FrmRadioTerapia.IdPaciente & "' And Activo=1"
Set RsD = CrearRS(CSql)

If RsD.EOF Then Exit Sub
RsD.MoveFirst
Do While Not RsD.EOF
    Combo1.AddItem RsD.Fields("TECNICA").Value
    RsD.MoveNext
Loop

End Sub


VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDirectorios 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar directorios"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "FrmDirectorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4395
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   3975
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Cancelar"
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
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDirectorios.frx":1002
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Aceptar"
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
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDirectorios.frx":101E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directorios:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmDirectorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ruta As String
Private Sub ChameleonBtn1_Click()
Select Case opcion
Case 1
    Ruta = Dir1.Path & "\"
    FrmOpciones.TxtRutaRpt.Text = Ruta
Case 2
    Ruta = Dir1.Path & "\"
    FrmOpciones.TxtRutaFotos.Text = Ruta
End Select

Unload Me
End Sub

Private Sub ChameleonBtn2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Centrar Me
End Sub

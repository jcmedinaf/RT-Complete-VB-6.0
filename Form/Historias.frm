VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Historias Médicas"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form5"
   ScaleHeight     =   9570
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   240
      TabIndex        =   26
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2640
      Top             =   9120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Historia"
      Height          =   6015
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   7815
      Begin VB.TextBox Text12 
         DataField       =   "medico"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Text            =   "Text12"
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         DataField       =   "anatomia"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Text            =   "Text11"
         Top             =   5280
         Width           =   3615
      End
      Begin VB.TextBox Text10 
         DataField       =   "tratamiento"
         DataSource      =   "Adodc2"
         Height          =   765
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "Historias.frx":0000
         Top             =   4200
         Width           =   6375
      End
      Begin VB.TextBox Text9 
         DataField       =   "diagnostico"
         DataSource      =   "Adodc2"
         Height          =   855
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "Historias.frx":0007
         Top             =   3000
         Width           =   6375
      End
      Begin VB.TextBox Text8 
         DataField       =   "antecedentes"
         DataSource      =   "Adodc2"
         Height          =   855
         Left            =   3600
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "Historias.frx":000D
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         DataField       =   "examenes"
         DataSource      =   "Adodc2"
         Height          =   855
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "Historias.frx":0013
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         DataField       =   "motivo"
         DataSource      =   "Adodc2"
         Height          =   855
         Left            =   3600
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Historias.frx":0019
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         DataField       =   "enfermedad"
         DataSource      =   "Adodc2"
         Height          =   855
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Historias.frx":001F
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Antecedentes"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diagnóstico"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tratamiento"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Anatomia"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Médico Tratante"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Examenes"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Motivo"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enfermedad"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   9120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Historias.frx":0025
      OLEDBString     =   $"Historias.frx":0107
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "pacientes"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Paciente"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox Text1 
         DataField       =   "apellidos"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         DataField       =   "nombres"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         DataField       =   "ci"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         DataField       =   "historia"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Apellido"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ced Identidad"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Historia Médica"
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub actualiz()
consultasql = "SELECT * FROM historias WHERE ci='" + Text3.Text + "'"
Adodc2.ConnectionString = cadenaconexioN
Adodc2.RecordSource = consultasql
Adodc2.Refresh
'DataGrid1.Refresh



End Sub

Private Sub Text3_Change()
Call actualiz

End Sub

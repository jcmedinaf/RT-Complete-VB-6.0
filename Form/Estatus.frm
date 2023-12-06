VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000005&
   Caption         =   "Estatus del Paciente"
   ClientHeight    =   5205
   ClientLeft      =   4560
   ClientTop       =   3000
   ClientWidth     =   14790
   LinkTopic       =   "Form4"
   ScaleHeight     =   5205
   ScaleWidth      =   14790
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid datagrid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   12648447
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Cedula"
         Caption         =   "Cedula"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "›"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Apellido"
         Caption         =   "Apellido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "AREAMEDICA"
         Caption         =   "Internista"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   1
            TrueValue       =   ""
            FalseValue      =   "Ø"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AREANUTRICION"
         Caption         =   "Nutrición"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "AREARADIOMEDICA"
         Caption         =   "Radioterapeuta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "AREAPSICOLOGICA"
         Caption         =   "Psicología"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "AREAFISICAMEDICA"
         Caption         =   "Fisica Médica"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "AREATECNICA"
         Caption         =   "Radioterapia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Recepcion"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   14655
      Begin VB.CommandButton Command4 
         Caption         =   " Lista de Espera"
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo Registro"
         Height          =   375
         Left            =   8280
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Asignada"
         Height          =   375
         Left            =   10080
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cedula "
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1200
      Top             =   5520
      Width           =   3495
      _ExtentX        =   6165
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
      Connect         =   $"Estatus.frx":6DAB
      OLEDBString     =   $"Estatus.frx":6E93
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Consulta1"
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
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   0
      Picture         =   "Estatus.frx":6F7B
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD As New adodb.Recordset

Private Sub Command1_Click()
DataEnvironment1.AO.AddNew

End Sub
'ModoEditar True

Private Sub Command2_Click()

If MsgBox("¿desea guardar los cambios?", _
vbQuestion + vbYesNo, "Pregunta") = vbYes Then
MsgBox ("guardado")
DataEnvironment1.AO.Update
End If
End Sub
'ModoEditar False

Private Sub Command3_Click()
Dim nuevo As String
Dim Cedula As String

Nombre.AddNew
Nombre!paciente = nuevo
Nombre.Update
End Sub

Private Sub Command5_Click()
Unload Me
End Sub


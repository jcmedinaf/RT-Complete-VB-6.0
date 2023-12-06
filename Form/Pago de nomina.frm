VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPagoDeNomina 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nomina de Pago"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   Icon            =   "Pago de nomina.frx":0000
   LinkTopic       =   "Form39"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10155
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   9975
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   495
         Left            =   8880
         TabIndex        =   22
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":1002
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
         Height          =   495
         Left            =   1560
         TabIndex        =   23
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":101E
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
         TabIndex        =   24
         ToolTipText     =   "Agregar "
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":103A
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
         Left            =   7680
         TabIndex        =   25
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":1056
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
         TabIndex        =   26
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":1072
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
         TabIndex        =   27
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":108E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnImprimir 
         Height          =   495
         Left            =   3840
         TabIndex        =   28
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Pago de nomina.frx":10AA
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos de Empleado"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   5655
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   60686337
            CurrentDate     =   39958
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Ingreso:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Periodo Actual"
         Height          =   1215
         Left            =   5880
         TabIndex        =   13
         Top             =   1320
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   60686337
            CurrentDate     =   39958
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1680
            TabIndex        =   15
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   60686337
            CurrentDate     =   39932
         End
         Begin ChamaleonButton.ChameleonBtn BtnAyuda 
            Height          =   375
            Left            =   3240
            TabIndex        =   29
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   2
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Pago de nomina.frx":10C6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   1680
            TabIndex        =   16
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   8040
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4895
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
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5880
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60686337
         CurrentDate     =   39932
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Recibo:"
         Height          =   195
         Left            =   7200
         TabIndex        =   12
         Top             =   450
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4440
         TabIndex        =   10
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cédula:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         Height          =   195
         Left            =   5040
         TabIndex        =   6
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido(s):"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   765
      End
   End
End
Attribute VB_Name = "FrmPagoDeNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDNomina As New ADODB.Recordset
Dim BDNom As New ADODB.Recordset

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centrar Me

If BDNom.State = 1 Then BDNom.Close
CSql = "Select * From Empleados"
Set BDNom = CrearRS(CSql)

Call Nomina

End Sub
Sub Nomina()
If BDNom.EOF Or BDNom.BOF Then
    Msg = "LLego al Final del Registro"
    MsgBox Msg
    BDNom.MoveFirst
End If
Text1.Text = BDNom.Fields("cedula")
Text2.Text = BDNom.Fields("Nombre")
Text3.Text = BDNom.Fields("Apellido")
DTPicker1.Value = BDNom.Fields("Fecha_Ing")
DTPicker2.Value = BDNom.Fields("Fecha_Nac")
'Text4.Text = BDNom.Fields("Nºresibo")
'Text5.Text = BDNom.Fields("Cel")

End Sub

Private Sub Image3_Click()
On Error GoTo h
CommonDialog1.ShowOpen
FotoP = CommonDialog1.filename
Image3.Picture = LoadPicture(FotoP)
Image3.Refresh
Cambio = 1
Exit Sub
h:
MsgBox Err.Description

End Sub


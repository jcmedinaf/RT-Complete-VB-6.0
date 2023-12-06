VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteEdoFinanciero 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estados Financieros"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3615
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir cuentas con saldo CERO ""0"""
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteEdoFinanciero.frx":0000
         Left            =   1890
         List            =   "FrmContReporteEdoFinanciero.frx":0064
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   56360963
         CurrentDate     =   40254
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56360963
         CurrentDate     =   40254
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período del reporte:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del reporte:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   930
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Jerarquía:"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   1380
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   4335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
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
         MICON           =   "FrmContReporteEdoFinanciero.frx":00DD
         PICN            =   "FrmContReporteEdoFinanciero.frx":00F9
         PICH            =   "FrmContReporteEdoFinanciero.frx":02C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAyuda 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
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
         MICON           =   "FrmContReporteEdoFinanciero.frx":04F7
         PICN            =   "FrmContReporteEdoFinanciero.frx":0513
         PICH            =   "FrmContReporteEdoFinanciero.frx":07B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnResultados 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Resultados"
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
         MICON           =   "FrmContReporteEdoFinanciero.frx":0B1F
         PICN            =   "FrmContReporteEdoFinanciero.frx":0B3B
         PICH            =   "FrmContReporteEdoFinanciero.frx":0DCD
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
Attribute VB_Name = "FrmContReporteEdoFinanciero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim i As Integer
Dim j As Integer
Dim Spdr As String
Dim FormatoPDC As String

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centrar Me
CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMM Carga los niveles del PDC MMMMMMMMMMMMMMMMMMMMMMMM
  CSql = "SELECT * FROM ContPDCConfig WHERE IdEmpresa=" & IdEmprs
  Set RsTemp = CrearRS(CSql)

  If RsTemp.RecordCount <> 0 Then
      FormatoPDC = Trim(RsTemp.Fields("Formato"))
      Spdr = Trim(RsTemp.Fields("Separador"))
  End If

  j = 0
  Combo1.Clear
  Combo1.AddItem " "
  For i = 1 To Len(FormatoPDC)
      If Mid(FormatoPDC, i, 1) = Spdr Then j = j + 1: Combo1.AddItem "Nivel " & j: Combo1.ItemData(Combo1.NewIndex) = InStr(i, FormatoPDC, Spdr, vbTextCompare)
  Next i
  Combo1.ListIndex = 0
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "SELECT * FROM ContPDC WHERE IdEmpresa=" & IdEmprs & " AND Activo='1' ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

DTPicker1.Value = "01" & Format(Now, "/MM/yyyy")
DTPicker2.Value = Format(DateSerial(Year(CDate(Now)), Month(CDate(Now)) + 1, 0), "dd/MM/yyyy")

Combo1.Clear
Combo2.Clear

Combo1.AddItem " "
Combo2.AddItem " "

If RsTemp.RecordCount = 0 Then Exit Sub

Combo1.Clear
Combo2.Clear

While Not RsTemp.EOF
    Combo1.AddItem Trim(RsTemp.Fields("Identificador").Value)
    Combo1.ItemData(Combo1.NewIndex) = RsTemp.Fields("IdPDC").Value
    Combo2.AddItem Trim(RsTemp.Fields("Identificador").Value)
    Combo2.ItemData(Combo2.NewIndex) = RsTemp.Fields("IdPDC").Value
    RsTemp.MoveNext
Wend


End Sub

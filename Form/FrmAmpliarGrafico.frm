VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmAmpliarGrafico 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ampliar Gráfico"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "FrmAmpliarGrafico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   9240
      Width           =   11175
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   10080
         TabIndex        =   2
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
         MICON           =   "FrmAmpliarGrafico.frx":1002
         PICN            =   "FrmAmpliarGrafico.frx":101E
         PICH            =   "FrmAmpliarGrafico.frx":11E7
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Imprimir"
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
         MICON           =   "FrmAmpliarGrafico.frx":141C
         PICN            =   "FrmAmpliarGrafico.frx":1438
         PICH            =   "FrmAmpliarGrafico.frx":155D
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Representación Gráfica"
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   8895
         Left            =   120
         OleObjectBlob   =   "FrmAmpliarGrafico.frx":17ED
         TabIndex        =   4
         Top             =   240
         Width           =   10935
      End
   End
End
Attribute VB_Name = "FrmAmpliarGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGraficar As New ADODB.Recordset

Private Sub BtnCerrar_Click()
Unload Me
End Sub

'Private Sub Form_Load()
'
'If FrmTablaEstadisticas.Band = True Then
'CSql = "SELECT Diagnotico, COUNT(Diagnotico) AS TotalDiag From Consulta_Estadisiticas Where " & Trim(FrmTablaEstadisticas.TxtQuery.Text) & " GROUP BY Diagnotico"
'End If
'
'If FrmTablaEstadisticas.Band1 = True Then
'CSql = "SELECT Diagnotico, COUNT(Diagnotico) AS TotalDiag From Consulta_Estadisiticas Where " & Trim(FrmTablaEstadisticas.TxtQuery.Text) & " GROUP BY Diagnotico"
'End If
'
'Set RsGraficar = CrearRS(CSql)
''If Not IsNull(RsGraficar) Then Exit Sub
'If RsGraficar.RecordCount = 0 Then ' If no Record in Database, then Show an Error Msg and Exit the Sub
'    MsgBox "No hay Datos para realizar el Gráfico!!!", vbCritical, "Error": Exit Sub
'Else
'    ReDim ArrayChart(1 To RsGraficar.RecordCount, 1 To 2) ' Array
'    'Puuting Records from Database to Array
'    For X = 1 To RsGraficar.RecordCount
'        ArrayChart(X, 1) = RsGraficar!Diagnotico
'        ArrayChart(X, 2) = CInt(RsGraficar!TotalDiag)
'        RsGraficar.MoveNext
'    Next X
'
'    '# Assigns our array to the MSChart control #
'    MSChart1.ChartData = ArrayChart
'    MSChart1.ChartType = 0
'    MSChart1.Refresh
'End If
'End Sub
Private Sub MSChart1_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

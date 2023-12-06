VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReporteNutricion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Nutricion"
   ClientHeight    =   2865
   ClientLeft      =   8850
   ClientTop       =   795
   ClientWidth     =   2730
   Icon            =   "Info_Nutri.frx":0000
   LinkTopic       =   "Form50"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   2730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Cerrar"
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
         MICON           =   "Info_Nutri.frx":1002
         PICN            =   "Info_Nutri.frx":101E
         PICH            =   "Info_Nutri.frx":11E7
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
         TabIndex        =   7
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "Info_Nutri.frx":141C
         PICN            =   "Info_Nutri.frx":1438
         PICH            =   "Info_Nutri.frx":155D
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
      Caption         =   "Reporte Semanal"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2040
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpFechaDesde 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51707905
         CurrentDate     =   39939
      End
      Begin MSComCtl2.DTPicker DtpFechaHasta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51707905
         CurrentDate     =   39939
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmReporteNutricion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FechaDesde As String
Dim FechaHasta As String
Dim RsReporte As New ADODB.Recordset
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()

FechaDesde = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
FechaHasta = Format(DtpFechaHasta.Value, "dd/mm/yyyy")

'FechaDesde = Format(DtpFechaDesde.Value, "yyyy/mm/dd")
'FechaHasta = Format(DtpFechaHasta.Value, "yyyy/mm/dd")

''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\InformeSemanal.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "DSN=CrReporte"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    '.SelectionFormula = "{Info_Nutri.FechaNu} >= '" & FechaSQL(FechaDesde) & "' And {Info_Nutri.FechaNu} <= '" & FechaSQL(FechaHasta) & "'" 'And idpaciente = '" & IdPac1 & "'"
    '.SelectionFormula = "{Info_Nutri.FechaNu} = '" & FechaSQL(FechaDesde) & "' And idpaciente = '" & IdPac1 & "'"
    '.SelectionFormula = "ToText({Info_Nutri.FechaNu}, 'dd/MM/yyyy') = '" & FechaDesde & "'"
    '.SelectionFormula = "ToText({Info_Nutri.FechaNu}, 'dd/MM/yyyy') >= '" & FechaDesde & "' AND ToText({Info_Nutri.FechaNu}, 'dd/MM/yyyy') <= '" & FechaHasta & "'" ' AND {Info_Nutri.idpaciente}= " & IdPac1
    .SelectionFormula = "{Info_Nutri.FechaNu} >= " & FechaSQL(FechaDesde) & " AND {Info_Nutri.FechaNu} <= " & FechaSQL(FechaHasta) & ""
    .ReportTitle = "Reporte Informe Semanal"
    .WindowTitle = "Reporte Informe Semanal Desde: " & FechaDesde & " Hasta: " & FechaHasta
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

End Sub

Private Sub Form_Load()

Centrar Me
DtpFechaDesde.Value = Now
DtpFechaHasta.Value = Now
End Sub

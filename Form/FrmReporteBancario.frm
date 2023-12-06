VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReporteBancario 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Bancario"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   Icon            =   "FrmReporteBancario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3690
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2880
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox CboBancos 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   1680
         Width           =   2295
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   1200
            TabIndex        =   2
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
            MICON           =   "FrmReporteBancario.frx":1002
            PICN            =   "FrmReporteBancario.frx":101E
            PICH            =   "FrmReporteBancario.frx":11E7
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
            MICON           =   "FrmReporteBancario.frx":141C
            PICN            =   "FrmReporteBancario.frx":1438
            PICH            =   "FrmReporteBancario.frx":155D
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
      Begin MSComCtl2.DTPicker DtpFechaHasta 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   39939
      End
      Begin MSComCtl2.DTPicker DtpFechaDesde 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   39939
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha  Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmReporteBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()

FechaDesde = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
FechaHasta = Format(DtpFechaHasta.Value, "dd/mm/yyyy")

''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\RelacionCobros.rpt"
    .Connect = "DSN=CrReporte"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{RelacionCobros.Fecha} >= " & FechaSQL(FechaDesde) & " AND {RelacionCobros.Fecha} <= " & FechaSQL(FechaHasta) & " And {RelacionCobros.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
    .ReportTitle = "Reporte Relación de Cobros"
    .WindowTitle = "Reporte Relación de Cobros Desde: " & FechaDesde & " Hasta: " & FechaHasta
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

End Sub

Private Sub Form_Load()
Dim RsCboBancos As New ADODB.Recordset
Centrar Me
DtpFechaDesde.Value = DateTime.Date
DtpFechaHasta.Value = DateTime.Date

CSql = "Select * From CajasBancos"
Set RsCboBancos = CrearRS(CSql)

Do While Not RsCboBancos.EOF
    CboBancos.AddItem RsCboBancos.Fields("Descripcion").Value
    CboBancos.ItemData(CboBancos.NewIndex) = RsCboBancos.Fields("IdCajaBanco").Value
    RsCboBancos.MoveNext
Loop

End Sub

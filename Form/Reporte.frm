VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReportePresupuestoEmitidos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte"
   ClientHeight    =   3180
   ClientLeft      =   8490
   ClientTop       =   555
   ClientWidth     =   4440
   Icon            =   "Reporte.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Orientación"
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
         Begin VB.OptionButton OptHorizontal 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Horizontal"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptVertical 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Vertical"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Destino"
         Height          =   1455
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   2295
         Begin VB.OptionButton OptPantalla 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Por Pantalla"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptImpresora 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Impresora"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Optarchivo 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Archivo"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1215
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   1800
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowNavigationCtls=   -1  'True
            WindowShowCancelBtn=   -1  'True
            WindowShowPrintBtn=   -1  'True
            WindowShowExportBtn=   -1  'True
            WindowShowZoomCtl=   -1  'True
            WindowShowCloseBtn=   -1  'True
            WindowShowProgressCtls=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSComDlg.CommonDialog cdgMain 
            Left            =   1680
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.TextBox TxtNoCopias 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   1800
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   1200
            TabIndex        =   6
            ToolTipText     =   "Cerrar"
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
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
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "Reporte.frx":1002
            PICN            =   "Reporte.frx":101E
            PICH            =   "Reporte.frx":11E7
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
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "Reporte.frx":141C
            PICN            =   "Reporte.frx":1438
            PICH            =   "Reporte.frx":155D
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51445761
         CurrentDate     =   39939
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51445761
         CurrentDate     =   39939
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Copias"
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha  Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmReportePresupuestoEmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()

'Determinar variable de salida del reporte
If OptPantalla.Value = True Then
  OutputDestination = 0
ElseIf OptImpresora.Value = True Then
  OutputDestination = 1
ElseIf Optarchivo.Value = True Then
  OutputDestination = 2
  cdgMain.Filter = "Archivos de Reportes (*.doc)|*.doc"
  cdgMain.InitDir = App.Path
  cdgMain.ShowSave
  CrystalReport1.PrintFileName = cdgMain.filename
  CrystalReport1.PrintFileType = 17
End If

'Determinar la Orientacion del reporte
If OptVertical.Value = True Then 'Vertical
    Printer.Orientation = vbPRORPortrait
End If

If OptHorizontal.Value = True Then ' Horizontal
    Printer.Orientation = vbPRORLandscape
End If

'Determinar el numero de copiar
CrystalReport1.CopiesToPrinter = TxtNoCopias.Text

'Determina el Destino del Reporte
CrystalReport1.Destination = OutputDestination

'Determina el tamaño de la ventana del reporte
CrystalReport1.WindowState = crptMaximized


CrystalReport1.WindowTitle = "Reporte de Nutrición"
FmlaText$ = "{presupuesto1.fecha} >= " & FechaSQL(DTPicker1.Value) & " And {presupuesto1.fecha} <= " & FechaSQL(DTPicker2.Value) & ""
CrystalReport1.SelectionFormula = FmlaText$
CrystalReport1.ReportFileName = RutaInformes & "\" & "factura.rpt"

     
On Error GoTo errorHandler
    CrystalReport1.Action = 1
    Exit Sub
errorHandler:
    MsgBox CrystalReport1.LastErrorString, 16, "Mensaje de Error"
    Exit Sub

End Sub

Private Sub Form_Load()
Centrar Me
'CrystalReport1.ReportFileName = Direc & "informes\Reporte_presupuesto.rpt"
End Sub

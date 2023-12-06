VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmLibroVentas 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Ventas"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "FrmLibroVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   14160
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   13935
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6360
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Reporte de Libro de Ventas"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox TxtFechaHasta 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtFechaDesde 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CboEstatus 
         Height          =   315
         ItemData        =   "FrmLibroVentas.frx":1002
         Left            =   2760
         List            =   "FrmLibroVentas.frx":100F
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DtpFechaDesde 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   40175
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnularFactura 
         Height          =   615
         Left            =   11160
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Anular Factura"
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
         MICON           =   "FrmLibroVentas.frx":1031
         PICN            =   "FrmLibroVentas.frx":104D
         PICH            =   "FrmLibroVentas.frx":12D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   615
         Left            =   12840
         TabIndex        =   3
         ToolTipText     =   "Cerrar "
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
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
         MICON           =   "FrmLibroVentas.frx":170B
         PICN            =   "FrmLibroVentas.frx":1727
         PICH            =   "FrmLibroVentas.frx":18F0
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
         Height          =   615
         Left            =   9720
         TabIndex        =   5
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
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
         MICON           =   "FrmLibroVentas.frx":1B25
         PICN            =   "FrmLibroVentas.frx":1B41
         PICH            =   "FrmLibroVentas.frx":1C66
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFechaHasta 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   40175
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   615
         Left            =   4440
         TabIndex        =   18
         ToolTipText     =   "Buscar"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Buscar"
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
         MICON           =   "FrmLibroVentas.frx":1EF6
         PICN            =   "FrmLibroVentas.frx":1F12
         PICH            =   "FrmLibroVentas.frx":2177
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin MSComctlLib.ListView LstVentas 
         Height          =   6135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   10821
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No Rif"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre o Razón Social"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "No Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "No. Nota Crédito o Débito"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "No. Factura Afectada"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Base Imponible"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "% Alicuota"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Impuesto (I.V.A.)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Total Ventas"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label LblTotalGeneral 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   315
         Left            =   12045
         TabIndex        =   15
         Top             =   6450
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total General:"
         Height          =   195
         Left            =   10920
         TabIndex        =   14
         Top             =   6510
         Width           =   1005
      End
      Begin VB.Label LblCantidadFacturas 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   6450
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Facturas:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   6510
         Width           =   1560
      End
   End
End
Attribute VB_Name = "FrmLibroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsLibroVentas As New ADODB.Recordset
Dim RsTotalGeneral As New ADODB.Recordset
Public Fa As String
Private Sub BtnAnularFactura_Click()
Dim resp As Byte
Dim RsVerifica As New ADODB.Recordset
CSql = "Select * From C_Cobrar Where N_Factura='" & Fa & "'"
Set RsVerifica = CrearRS(CSql)

If RsVerifica.Fields("Anulada").Value = 1 Then
    MsgBox "La Factura Nº " & Fa & " ya se encuentra Anulada", vbCritical + vbOKOnly, "Mensaje de Error"
    Exit Sub
End If


If Fa <> "" Then
    
   resp = MsgBox("Esta seguro de Anular la Factura Nro. " & Fa & "", vbQuestion + vbYesNo, "Confirmar")
    If resp = 7 Then Exit Sub
    FrmAnularFactura.Show vbModal, FrmPrincipal
Else
    MsgBox "Debe de Seleccionar la Factura a Anular", vbCritical + vbOKOnly, "Mensaje de Error"
End If
End Sub

Public Sub BtnBuscar_Click()
Fa = ""
Cargar_Datos
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()

''========= ESTE ES EL CODIGO NUEVO ==========
If CboEstatus.Text = "Todas" Or CboEstatus.Text = "" Then
    If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text = "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} >= " & FechaSQL(TxtFechaDesde.Text) & ""
            .WindowTitle = "Libro de Ventas - Desde: " & TxtFechaDesde.Text & " "
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text <> "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} <= " & FechaSQL(TxtFechaHasta.Text) & ""
            .WindowTitle = "Libro de Ventas - Hasta: " & TxtFechaHasta.Text
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} >= " & FechaSQL(TxtFechaDesde.Text) & " AND {LibroVentas.Fecha} <= " & FechaSQL(TxtFechaHasta.Text) & ""
            .WindowTitle = "Libro de Ventas - Desde: " & TxtFechaDesde.Text & " Hasta: " & TxtFechaHasta.Text
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .WindowTitle = "Libro de Ventas "
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    End If
End If
If CboEstatus.Text = "Anuladas" Then
    If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text = "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} >= " & FechaSQL(TxtFechaDesde.Text) & " And {LibroVentas.Anulada}=1"
            .WindowTitle = "Libro de Ventas - Factura Anuladas Desde: " & TxtFechaDesde.Text & " "
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text <> "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} <= " & FechaSQL(TxtFechaHasta.Text) & " And {LibroVentas.Anulada}=1"
            .WindowTitle = "Libro de Ventas - Factura Anuladas Hasta: " & TxtFechaHasta.Text
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} >= " & FechaSQL(TxtFechaDesde.Text) & " AND {LibroVentas.Fecha} <= " & FechaSQL(TxtFechaHasta.Text) & " And {LibroVentas.Anulada}=1"
            .WindowTitle = "Libro de Ventas - Factura Anuladas Desde: " & TxtFechaDesde.Text & " Hasta: " & TxtFechaHasta.Text
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Anulada}=1"
            .WindowTitle = "Libro de Ventas - Factura Anuladas"
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    End If
End If

If CboEstatus.Text = "No Anuladas" Then
    If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text = "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} >= " & FechaSQL(TxtFechaDesde.Text) & " And {LibroVentas.Anulada}=0"
            .WindowTitle = "Libro de Ventas - Factura No Anuladas Desde: " & TxtFechaDesde.Text & " "
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text <> "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} <= " & FechaSQL(TxtFechaHasta.Text) & " And {LibroVentas.Anulada}=0"
            .WindowTitle = "Libro de Ventas - Factura No Anuladas Hasta: " & TxtFechaHasta.Text
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Fecha} >= " & FechaSQL(TxtFechaDesde.Text) & " AND {LibroVentas.Fecha} <= " & FechaSQL(TxtFechaHasta.Text) & " And {LibroVentas.Anulada}=0"
            .WindowTitle = "Libro de Ventas - Factura No Anuladas Desde: " & TxtFechaDesde.Text & " Hasta: " & TxtFechaHasta.Text
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
        With CrystalReport1
            .ReportFileName = RutaInformes & "\LibroVentas.rpt"
            .Connect = "DSN=CrReporte"
            .DiscardSavedData = True
            .RetrieveDataFiles
            .ReportSource = 0
            .SelectionFormula = "{LibroVentas.Anulada}=0"
            .WindowTitle = "Libro de Ventas - Factura No Anuladas"
            .Destination = crptToWindow
            .PrintFileType = crptCrystal
            .WindowState = crptMaximized
            .WindowMaxButton = False
            .WindowMinButton = False
            .Action = 1
        End With
    End If
End If

End Sub

Sub ReporteGeneral()

CSql = "select * From LibroVentas"
Set RsReporte = CrearRS(CSql)

Set DrptLibroVentas.DataSource = RsReporte


For i = 1 To RsReporte.RecordCount
    DrptLibroVentas.Sections("Sección1").Controls("LblNo").Caption = i
Next i
DrtpResumenNomina.Orientation = rptOrientLandscape
DrptLibroVentas.Show
End Sub

Sub ReportePorFechas()

End Sub

Private Sub DtpFechaDesde_Change()
TxtFechaDesde.Text = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaHasta_Change()
TxtFechaHasta.Text = Format(DtpFechaHasta.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Centrar Me

DtpFechaDesde.Value = Now
DtpFechaHasta.Value = Now

Cargar_Datos

End Sub

Private Sub LstVentas_Click()
If LstVentas.ListItems.Count > 0 Then
    Fa = LstVentas.SelectedItem.ListSubItems(3).Text
End If
End Sub

Sub Cargar_Datos()
Dim CondFecha As String
Dim CSql As String
Dim CSql2 As String

TxtFechaDesde.Text = Replace(TxtFechaDesde.Text, " ", "")
TxtFechaHasta.Text = Replace(TxtFechaHasta.Text, " ", "")

If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text = "" Then
    If IsDate(TxtFechaDesde.Text) Then
        CondFecha = "(Fecha >= CAST('" & TxtFechaDesde.Text & "' AS DATETIME))"
    Else
        MsgBox "Ingrese una Fecha Valida!", vbCritical + vbOKOnly, "Error - Fecha no valida"
        TxtFechaDesde.Text = ""
        Exit Sub
    End If
Else
    If TxtFechaDesde.Text = "" And TxtFechaHasta.Text <> "" Then
        If IsDate(TxtFechaHasta.Text) Then
            CondFecha = "(Fecha <= CAST('" & TxtFechaHasta.Text & "' AS DATETIME))"
        Else
            MsgBox "Ingrese una Fecha Valida!", vbCritical + vbOKOnly, "Error - Fecha no valida"
            TxtFechaHasta.Text = ""
            Exit Sub
        End If
    Else
        If IsDate(TxtFechaHasta.Text) And IsDate(TxtFechaHasta.Text) Then
        
            If Not ((DateValue(TxtFechaHasta.Text) - DateValue(TxtFechaDesde.Text)) <= -1) Then
                CondFecha = "(Fecha >= CAST('" & TxtFechaDesde.Text & "' AS DATETIME)) AND (Fecha <= CAST('" & TxtFechaHasta.Text & "' AS DATETIME))"
            Else
                MsgBox "La fecha de inicio es MAYOR a la Fecha Fin!"
                TxtFechaDesde.Text = ""
                TxtFechaHasta.Text = ""
                Exit Sub
            End If
        Else
            If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
                If Not IsDate(TxtFechaDesde.Text) Then TxtFechaDesde.Text = ""
                If Not IsDate(TxtFechaHasta.Text) Then TxtFechaHasta.Text = ""
                MsgBox "Ingrese una Fecha Valida!", vbCritical + vbOKOnly, "Error - Fecha no valida"
                Exit Sub
            End If
'            If TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
'                If Not IsDate(TxtFechaDesde.Text) Then TxtFechaDesde.Text = ""
'                If Not IsDate(TxtFechaHasta.Text) Then TxtFechaHasta.Text = ""
'                MsgBox "Ingrese una Fecha Valida!", vbCritical + vbOKOnly, "Error - Fecha no valida"
'                Exit Sub
'            End If
           
        End If
    End If
End If

If CboEstatus.Text = "Todas" Or CboEstatus.Text = "" Then

    If CondFecha = "" Then
        CSql = "Select * From C_Cobrar Order By Fecha Asc"
        CSql2 = "Select Sum(Monto) as TotalGeneral From C_Cobrar"
    Else
        CSql = "Select * From C_Cobrar Where " & CondFecha & " Order By Fecha Asc"
        CSql2 = "Select Sum(Monto) as TotalGeneral From C_Cobrar WHERE " & CondFecha
    End If
    Set RsLibroVentas = CrearRS(CSql)
    DoEvents
    Set RsTotalGeneral = CrearRS(CSql2)

    LblTotalGeneral.Caption = Format(RsTotalGeneral.Fields("TotalGeneral").Value, "#,##0.00")

Else
    If CboEstatus.Text = "Anuladas" Then
        
        If CondFecha = "" Then
            CSql = "Select * From C_Cobrar Where Anulada=1 Order By Fecha Asc"
            CSql2 = "Select Sum(Monto) as TotalGeneral From C_Cobrar Where Anulada=1"
            Else
            CSql = "Select * From C_Cobrar Where Anulada=1 AND " & CondFecha & " Order By Fecha Asc"
            CSql2 = "Select Sum(Monto) as TotalGeneral From C_Cobrar Where Anulada=1 AND " & CondFecha
        End If
        
        Set RsLibroVentas = CrearRS(CSql)
        Set RsTotalGeneral = CrearRS(CSql2)
        
        LblTotalGeneral.Caption = Format(RsTotalGeneral.Fields("TotalGeneral").Value, "#,##0.00")
    Else
        If CboEstatus.Text = "No Anuladas" Then
            
            If CondFecha = "" Then
                CSql = "Select * From C_Cobrar Where Anulada=0 Order By Fecha Asc"
                CSql2 = "Select Sum(Monto) as TotalGeneral From C_Cobrar Where Anulada=0"
            Else
                CSql = "Select * From C_Cobrar Where Anulada=0 AND " & CondFecha & " Order By Fecha Asc"
                CSql2 = "Select Sum(Monto) as TotalGeneral From C_Cobrar Where Anulada=0 AND " & CondFecha
            End If
            Set RsLibroVentas = CrearRS(CSql)
            Set RsTotalGeneral = CrearRS(CSql2)
            
            LblTotalGeneral.Caption = Format(RsTotalGeneral.Fields("TotalGeneral").Value, "#,##0.00")
        Else
            Exit Sub
        End If
    End If
End If

LstVentas.ListItems.Clear
    
Do While Not RsLibroVentas.EOF
    With LstVentas
        If RsLibroVentas.Fields("Anulada").Value <> 1 Then
            i = i + 1
            .ListItems.Add , , RsLibroVentas.Fields("Fecha").Value
            
            CSql = "Select * From Cliente Where IdCliente='" & RsLibroVentas.Fields("IdCliente").Value & "'"
            Set RsCliente = CrearRS(CSql)
            
            .ListItems(i).ListSubItems.Add , , RsCliente.Fields("Rif").Value
            .ListItems(i).ListSubItems.Add , , RsCliente.Fields("Razon").Value
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("N_Factura").Value
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("N_Nc").Value
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("N_Fa").Value
            .ListItems(i).ListSubItems.Add , , Format(RsLibroVentas.Fields("SubTotal").Value, "#,##0.00")
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("TasaImpuesto").Value & "%"
            .ListItems(i).ListSubItems.Add , , Format(RsLibroVentas.Fields("Impuesto").Value, "#,##0.00")
            .ListItems(i).ListSubItems.Add , , Format(RsLibroVentas.Fields("Monto").Value, "#,##0.00")
        Else
            i = i + 1
            .ListItems.Add , , RsLibroVentas.Fields("Fecha").Value
            .ListItems(i).ListSubItems.Add , , "0"
            .ListItems(i).ListSubItems.Add , , "Factura Anulada"
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("N_Factura").Value
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("N_Nc").Value
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("N_Fa").Value
            .ListItems(i).ListSubItems.Add , , Format(RsLibroVentas.Fields("SubTotal").Value, "#,##0.00")
            .ListItems(i).ListSubItems.Add , , RsLibroVentas.Fields("TasaImpuesto").Value & "%"
            .ListItems(i).ListSubItems.Add , , Format(RsLibroVentas.Fields("Impuesto").Value, "#,##0.00")
            .ListItems(i).ListSubItems.Add , , Format(RsLibroVentas.Fields("Monto").Value, "#,##0.00")
        End If
    End With
    RsLibroVentas.MoveNext
Loop

LblCantidadFacturas.Caption = LstVentas.ListItems.Count
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmLibroCompras 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Compras"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14130
   Icon            =   "FrmLibroCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   14130
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   13935
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   9000
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Reporte de Libro de Compras"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowNavigationCtls=   -1  'True
         WindowShowCancelBtn=   -1  'True
         WindowShowPrintBtn=   -1  'True
         WindowShowExportBtn=   -1  'True
         WindowShowZoomCtl=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowProgressCtls=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox TxtFechaHasta 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1520
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtFechaDesde 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1520
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DtpFechaDesde 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   320
         _ExtentX        =   556
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51052545
         CurrentDate     =   40175
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   615
         Left            =   12840
         TabIndex        =   8
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
         MICON           =   "FrmLibroCompras.frx":1002
         PICN            =   "FrmLibroCompras.frx":101E
         PICH            =   "FrmLibroCompras.frx":11E7
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
         Left            =   11400
         TabIndex        =   9
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
         MICON           =   "FrmLibroCompras.frx":141C
         PICN            =   "FrmLibroCompras.frx":1438
         PICH            =   "FrmLibroCompras.frx":155D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtPFechaHasta 
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   320
         _ExtentX        =   556
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51052545
         CurrentDate     =   40175
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   615
         Left            =   2880
         TabIndex        =   13
         ToolTipText     =   "Buscar"
         Top             =   240
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
         MICON           =   "FrmLibroCompras.frx":17ED
         PICN            =   "FrmLibroCompras.frx":1809
         PICH            =   "FrmLibroCompras.frx":1A6E
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
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   690
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin MSComctlLib.ListView LstCompras 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   9340
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "No. Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "No. Nota de Débito"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "No. Nota de Crédito"
            Object.Width           =   2822
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
            Text            =   "Tasa Impuesto"
            Object.Width           =   2205
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
            Text            =   "Total General"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Facturas:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   5700
         Width           =   1560
      End
      Begin VB.Label LblCantidadFacturas 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   5640
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total General:"
         Height          =   195
         Left            =   10920
         TabIndex        =   3
         Top             =   5700
         Width           =   1005
      End
      Begin VB.Label LblTotalGeneral 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   315
         Left            =   12045
         TabIndex        =   2
         Top             =   5640
         Width           =   1770
      End
   End
End
Attribute VB_Name = "FrmLibroCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsLibroCompras As New ADODB.Recordset

Private Sub BtnBuscar_Click()
Cargar_Datos
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()
''========= ESTE ES EL CODIGO NUEVO ==========
If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text = "" Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\LibroCompras.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{LibroCompras.FechaEmision} >= " & FechaSQL(TxtFechaDesde.Text) & " "
        .WindowTitle = "Libro de Compras - Desde: " & TxtFechaDesde.Text & " "
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text <> "" Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\LibroCompras.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{LibroCompras.FechaEmision} <= " & FechaSQL(TxtFechaHasta.Text) & ""
        .WindowTitle = "Libro de Compras -  Hasta: " & TxtFechaHasta.Text
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
ElseIf TxtFechaDesde.Text <> "" And TxtFechaHasta.Text <> "" Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\LibroCompras.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .SelectionFormula = "{LibroCompras.FechaEmision} >= " & FechaSQL(TxtFechaDesde.Text) & " And {LibroCompras.FechaEmision} >= " & FechaSQL(TxtFechaHasta.Text) & ""
        .WindowTitle = "Libro de Compras - Desde: " & TxtFechaDesde.Text & " Hasta: " & TxtFechaHasta.Text
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
ElseIf TxtFechaDesde.Text = "" And TxtFechaHasta.Text = "" Then
    With CrystalReport1
        .ReportFileName = RutaInformes & "\LibroCompras.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        .WindowTitle = "Libro de Compras "
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

End Sub

Private Sub DtpFechaDesde_Change()
TxtFechaDesde.Text = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaHasta_Change()
TxtFechaHasta.Text = Format(DtpFechaHasta.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Centrar Me

DtpFechaDesde.Value = Format(Now, "DD/MM/YY")
DtpFechaHasta.Value = Format(Now, "DD/MM/YY")

Cargar_Datos

End Sub

Sub Cargar_Datos()

Dim CondFecha As String
Dim CSql As String
Dim CSql2 As String

TxtFechaDesde.Text = Replace(TxtFechaDesde.Text, " ", "")
TxtFechaHasta.Text = Replace(TxtFechaHasta.Text, " ", "")

If TxtFechaDesde.Text <> "" And TxtFechaHasta.Text = "" Then
    If IsDate(TxtFechaDesde.Text) Then
        CondFecha = "(FechaEmision >= CAST('" & TxtFechaDesde.Text & "' AS DATETIME))"
        Else
        MsgBox "Ingrese una Fecha Valida!", vbCritical + vbOKOnly, "Error - Fecha no valida"
        TxtFechaDesde.Text = ""
        Exit Sub
    End If
    Else
    If TxtFechaDesde.Text = "" And TxtFechaHasta.Text <> "" Then
        If IsDate(TxtFechaHasta.Text) Then
            CondFecha = "(FechaEmision <= CAST('" & TxtFechaHasta.Text & "' AS DATETIME))"
            Else
            MsgBox "Ingrese una Fecha Valida!", vbCritical + vbOKOnly, "Error - Fecha no valida"
            TxtFechaHasta.Text = ""
            Exit Sub
        End If
        Else
        If IsDate(TxtFechaHasta.Text) And IsDate(TxtFechaHasta.Text) Then
        
            If Not ((DateValue(TxtFechaHasta.Text) - DateValue(TxtFechaDesde.Text)) <= -1) Then
                CondFecha = "(FechaEmision >= CAST('" & TxtFechaDesde.Text & "' AS DATETIME)) AND (FechaEmision <= CAST('" & TxtFechaHasta.Text & "' AS DATETIME))"
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
        End If
    End If
End If

    If CondFecha = "" Then
        CSql = "Select * From CtaPorPagar Order By FechaEmision Asc"
        CSql2 = "Select Sum(TotalGeneral) as TotalGeneral From CtaPorPagar"
        Else
        CSql = "Select * From CtaPorPagar Where " & CondFecha & " Order By FechaEmision Asc"
        CSql2 = "Select Sum(TotalGeneral) as TotalGeneral From CtaPorPagar WHERE " & CondFecha
    End If
    
    Set RsLibroCompras = CrearRS(CSql)
    Set RsTotalGeneral = CrearRS(CSql2)
    
    LstCompras.ListItems.Clear
    
    Do While Not RsLibroCompras.EOF
        With LstCompras
            i = i + 1
            .ListItems.Add , , RsLibroCompras.Fields("FechaEmision").Value
            CSql = "Select * From Proveedores Where IdProveedor='" & RsLibroCompras.Fields("IdProveedor").Value & "'"
            Set RsCliente = CrearRS(CSql)
            .ListItems(i).ListSubItems.Add , , RsCliente.Fields("RifProveedor").Value
            .ListItems(i).ListSubItems.Add , , RsCliente.Fields("Nombre").Value
            .ListItems(i).ListSubItems.Add , , RsLibroCompras.Fields("NoFactura").Value
            .ListItems(i).ListSubItems.Add , , 0
            .ListItems(i).ListSubItems.Add , , 0
            .ListItems(i).ListSubItems.Add , , Format(RsLibroCompras.Fields("SubTotal").Value, "#,##0.00")
            .ListItems(i).ListSubItems.Add , , 0 & "%"
            .ListItems(i).ListSubItems.Add , , Format(RsLibroCompras.Fields("Impuesto").Value, "#,##0.00")
            .ListItems(i).ListSubItems.Add , , Format(RsLibroCompras.Fields("TotalGeneral").Value, "#,##0.00")
     
        End With
        RsLibroCompras.MoveNext
    Loop
    
    LblCantidadFacturas.Caption = LstCompras.ListItems.Count
    LblTotalGeneral.Caption = Format(RsTotalGeneral.Fields("TotalGeneral").Value, "#,##0.00")
End Sub

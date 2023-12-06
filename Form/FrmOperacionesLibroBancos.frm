VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmOperacionesLibroBancos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Operaciones"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   Icon            =   "FrmOperacionesLibroBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptDepositos 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Depositos"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Reporte Operaciones Bancarias"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox CboBancos 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton OptCheques 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Cheques"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   840
         TabIndex        =   5
         Top             =   2160
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
            MICON           =   "FrmOperacionesLibroBancos.frx":1002
            PICN            =   "FrmOperacionesLibroBancos.frx":101E
            PICH            =   "FrmOperacionesLibroBancos.frx":11E7
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
            MICON           =   "FrmOperacionesLibroBancos.frx":141C
            PICN            =   "FrmOperacionesLibroBancos.frx":1438
            PICH            =   "FrmOperacionesLibroBancos.frx":155D
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   3480
         Top             =   600
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
         Format          =   51249153
         CurrentDate     =   39939
      End
      Begin MSComCtl2.DTPicker DtpFechaHasta 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249153
         CurrentDate     =   39939
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         Caption         =   "Bancos:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FrmOperacionesLibroBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FechaDesde As String
Dim FechaHasta As String
Dim RsReporte As New ADODB.Recordset
Dim Tipo_Mov
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()

FechaDesde = Format(DtpFechaDesde.Value, "dd/mm/yyyy")
FechaHasta = Format(DtpFechaHasta.Value, "dd/mm/yyyy")



If OptTodos.Value = False And OptCheques.Value = False And OptDepositos.Value = False Then
    Msg = "Seleccione una opción para realizar la busqueda"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If


If OptTodos.Value = True And CboBancos.ItemData(CboBancos.ListIndex) = 0 Then
CboBancos.ItemData(CboBancos.ListIndex) = 0
OptCheques.Value = False
OptDepositos.Value = False
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\OperacionLibroBancos1.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        '.SelectionFormula = "ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') >= '" & FechaDesde & "' AND ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') <= '" & FechaHasta & "'"
        'MsgBox .SelectionFormula
        .SelectionFormula = "{MoviBancoCaja.Fecha_Transa} >= " & FechaSQL(FechaDesde) & " AND {MoviBancoCaja.Fecha_Transa} <= " & FechaSQL(FechaHasta) & ""
        .ReportTitle = "Reporte Operaciones Bancarias"
        .WindowTitle = "Reporte Operaciones Bancarias, Desde: " & FechaDesde & " Hasta: " & FechaHasta
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

If OptCheques.Value = True And CboBancos.ItemData(CboBancos.ListIndex) = 0 Then
CboBancos.ItemData(CboBancos.ListIndex) = 0
OptTodos.Value = False
OptDepositos.Value = False
Tipo_Mov = 2
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\OperacionLibroBancos1.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        '.SelectionFormula = "ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') >= '" & FechaDesde & "' AND ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') <= '" & FechaHasta & "' And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & ""
        'MsgBox .SelectionFormula
        .SelectionFormula = "{MoviBancoCaja.Fecha_Transa} >= " & FechaSQL(FechaDesde) & " AND {MoviBancoCaja.Fecha_Transa} <= " & FechaSQL(FechaHasta) & " And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & ""
        .ReportTitle = "Reporte Operaciones Bancarias"
        .WindowTitle = "Reporte Operaciones Bancarias, Desde: " & FechaDesde & " Hasta: " & FechaHasta
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

If OptDepositos.Value = True And CboBancos.ItemData(CboBancos.ListIndex) = 0 Then
CboBancos.ItemData(CboBancos.ListIndex) = 0
OptTodos.Value = False
OptCheques.Value = False
Tipo_Mov = 1
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\OperacionLibroBancos1.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
'        .SelectionFormula = "ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') >= '" & FechaDesde & "' AND ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') <= '" & FechaHasta & "' And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & ""
        .SelectionFormula = "{MoviBancoCaja.Fecha_Transa} >= " & FechaSQL(FechaDesde) & " AND {MoviBancoCaja.Fecha_Transa} <= " & FechaSQL(FechaHasta) & " And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & ""
        'MsgBox .SelectionFormula
        .ReportTitle = "Reporte Operaciones Bancarias"
        .WindowTitle = "Reporte Operaciones Bancarias, Desde: " & FechaDesde & " Hasta: " & FechaHasta
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If


If OptTodos.Value = True And CboBancos.ItemData(CboBancos.ListIndex) <> 0 Then
OptCheques.Value = False
OptDepositos.Value = False

''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\OperacionLibroBancos1.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        '.SelectionFormula = "ToText({MoviBancoCaja.Fecha_Transa}, 'yyyy/mm/dd') >= '" & FechaDesde & "' AND ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') <= '" & FechaHasta & "' And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & " And {MoviBancoCaja.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
        .SelectionFormula = "{MoviBancoCaja.Fecha_Transa} >= " & FechaSQL(FechaDesde) & " AND {MoviBancoCaja.Fecha_Transa} <= " & FechaSQL(FechaHasta) & " And  {MoviBancoCaja.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
        'MsgBox .SelectionFormula
        .ReportTitle = "Reporte Operaciones Bancarias"
        .WindowTitle = "Reporte Operaciones Bancarias, Desde: " & FechaDesde & " Hasta: " & FechaHasta
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If


If OptCheques.Value = True And CboBancos.ItemData(CboBancos.ListIndex) <> 0 Then
OptTodos.Value = False
OptDepositos.Value = False
Tipo_Mov = 2
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\OperacionLibroBancos1.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        '.SelectionFormula = "ToText({MoviBancoCaja.Fecha_Transa}, 'yyyy/mm/dd') >= '" & FechaDesde & "' AND ToText({MoviBancoCaja.Fecha_Transa}, 'dd/MM/yyyy') <= '" & FechaHasta & "' And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & " And {MoviBancoCaja.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
        .SelectionFormula = "{MoviBancoCaja.Fecha_Transa} >= " & FechaSQL(FechaDesde) & " AND {MoviBancoCaja.Fecha_Transa} <= " & FechaSQL(FechaHasta) & " And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & " And {MoviBancoCaja.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
        'MsgBox .SelectionFormula
        .ReportTitle = "Reporte Operaciones Bancarias"
        .WindowTitle = "Reporte Operaciones Bancarias, Desde: " & FechaDesde & " Hasta: " & FechaHasta
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

If OptDepositos.Value = True And CboBancos.ItemData(CboBancos.ListIndex) <> 0 Then
OptTodos.Value = False
OptCheques.Value = False
Tipo_Mov = 1
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\OperacionLibroBancos1.rpt"
        .Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
'        .SelectionFormula = "ToText({MoviBancoCaja.Fecha_Transa}, 'yyyy/mm/dd') >= '" & FechaDesde & "' AND ToText({MoviBancoCaja.Fecha_Transa}, 'yyyy/mm/dd') <= '" & FechaHasta & "' And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & " And {MoviBancoCaja.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
        .SelectionFormula = "{MoviBancoCaja.Fecha_Transa} >= " & FechaSQL(FechaDesde) & " AND {MoviBancoCaja.Fecha_Transa} <= " & FechaSQL(FechaHasta) & " And {MoviBancoCaja.Tipo_Mov}=" & Tipo_Mov & " And {MoviBancoCaja.IdCajaBanco}=" & CboBancos.ItemData(CboBancos.ListIndex) & ""
        'MsgBox .SelectionFormula
        .ReportTitle = "Reporte Operaciones Bancarias"
        .WindowTitle = "Reporte Operaciones Bancarias, Desde: " & FechaDesde & " Hasta: " & FechaHasta
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If

End Sub

Private Sub Form_Load()

Centrar Me
DtpFechaDesde.Value = Now
DtpFechaHasta.Value = Now



Dim RsCajasBancos As New ADODB.Recordset
CSql = "Select * From CajasBancos"
Set RsCajasBancos = CrearRS(CSql)

CboBancos.AddItem "Todos"
CboBancos.Text = "Todos"
CboBancos.ItemData(CboBancos.NewIndex) = 0
Do While Not RsCajasBancos.EOF
    With CboBancos
        .AddItem RsCajasBancos.Fields("Descripcion").Value
        .ItemData(.NewIndex) = RsCajasBancos.Fields("IdCajaBanco").Value
    End With
    RsCajasBancos.MoveNext
Loop

End Sub

Private Sub OptCheques_Click()
OptTodos.Value = False
OptDepositos.Value = False
End Sub

Private Sub OptDepositos_Click()
OptTodos.Value = False
OptCheques.Value = False
End Sub

Private Sub OptTodos_Click()
OptCheques.Value = False
OptDepositos.Value = False
End Sub



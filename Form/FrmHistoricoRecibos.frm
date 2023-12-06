VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmHistoricoRecibos 
   BackColor       =   &H00EAEFEF&
   Caption         =   "Resumen de la Nómina Generada"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   Icon            =   "FrmHistoricoRecibos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   8880
         TabIndex        =   11
         Top             =   6120
         Width           =   4575
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   3240
            Top             =   360
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   3480
            TabIndex        =   12
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
            MICON           =   "FrmHistoricoRecibos.frx":1002
            PICN            =   "FrmHistoricoRecibos.frx":101E
            PICH            =   "FrmHistoricoRecibos.frx":11E7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   1560
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
         Begin ChamaleonButton.ChameleonBtn BtnImprimir 
            Height          =   375
            Left            =   120
            TabIndex        =   19
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
            MICON           =   "FrmHistoricoRecibos.frx":141C
            PICN            =   "FrmHistoricoRecibos.frx":1438
            PICH            =   "FrmHistoricoRecibos.frx":155D
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   6120
         Width           =   3975
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Razon Social, Cédula o Rif"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   10
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Busqueda"
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
            MICON           =   "FrmHistoricoRecibos.frx":17ED
            PICN            =   "FrmHistoricoRecibos.frx":1809
            PICH            =   "FrmHistoricoRecibos.frx":1A6E
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Nro de Recibo"
         Height          =   735
         Left            =   4320
         TabIndex        =   5
         Top             =   6120
         Width           =   4455
         Begin VB.TextBox TxtRecibo 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese el Número del Recibo"
            Top             =   240
            Width           =   1935
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscarRecibo 
            Height          =   375
            Left            =   2400
            TabIndex        =   7
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Buscar Recibo"
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
            MICON           =   "FrmHistoricoRecibos.frx":1D00
            PICN            =   "FrmHistoricoRecibos.frx":1D1C
            PICH            =   "FrmHistoricoRecibos.frx":1F81
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
      Begin VB.TextBox TxtAsignaciones 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox TxtOtros 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox TxtDeducciones 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalNeto 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   5520
         Width           =   1695
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5655
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9975
         Object.Width           =   4065
         Object.Height          =   5625
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin SystemOncoAmerica.DMGrid DMGrid2 
         Height          =   4935
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8705
         Object.Width           =   9105
         Object.Height          =   4905
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asignaciones:"
         Height          =   195
         Left            =   6240
         TabIndex        =   18
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deducciones:"
         Height          =   195
         Left            =   8160
         TabIndex        =   17
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Otros:"
         Height          =   195
         Left            =   9960
         TabIndex        =   16
         Top             =   5280
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
         Height          =   195
         Left            =   11760
         TabIndex        =   15
         Top             =   5280
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmHistoricoRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim resp As String
Dim Fecha1 As String
Dim Fecha2 As String
Dim TAsign As Double
Dim TDeduc As Double
Dim TOtros As Double
Dim TNeto As Double
Dim TempId As Integer
Dim Periodo


Private Sub BtnImprimir_Click()

If Periodo = "" Or Fecha1 = "" Then Exit Sub

' ========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\ResumenNomina.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "DSQ=OAClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    '.SelectionFormula = "{ReciboDePago.Periodo} = " & Periodo & "  And {ReciboDePago.Fecha_Gen} = " & FechaSQL(Fecha1) & ""
    .SelectionFormula = "{ReciboDePago.Periodo} = " & Periodo & "  And {ReciboDePago.Fecha_Fin_Nom} = " & FechaSQL(Fecha2) & ""
    .WindowTitle = "Resumen de Nomina "
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
End Sub

Private Sub DMGrid2_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    TempId = DMGrid2.ValorCelda(lRow, 8)
    
    If Not IsNull(TempId) Then
        If Trim(TempId) <> "" Then
            FrmReciboPagos.NTabla = 1
            FrmReciboPagos.IdEmpla = TempId
            FrmReciboPagos.FechaTemp = Fecha1
            FrmReciboPagos.BtnAgregarConceptos.Enabled = False
            FrmReciboPagos.BtnAnterior.Enabled = False
            FrmReciboPagos.BtnBorrar.Enabled = True
            'FrmReciboPagos.BtnBuscar.Enabled = False
            FrmReciboPagos.BtnDesHacer.Enabled = False
            FrmReciboPagos.BtnGuardarActualizar.Enabled = True
            FrmReciboPagos.BtnSiguiente.Enabled = False
            FrmReciboPagos.Show vbModal, FrmPrincipal
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid
Cargar_Periodos_Nomina
End Sub

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 3
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 35 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 35 / 100) - 300

DMGrid1.DColumnas(1).Caption = "Período"
DMGrid1.DColumnas(2).Caption = "Fecha Inicio"
DMGrid1.DColumnas(3).Caption = "Fecha Fin"
'DMGrid1.DColumnas(3).Caption = "Periodo"       ' calcular que periodo es la fecha elegida...

'MMMMMMMMMMMMMMMM  DMGrid2  MMMMMMMMMMMMMMMM
DMGrid2.Cols = 8
DMGrid2.Rows = 0
DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 0
DMGrid2.DColumnas(4).Alignment = 1
DMGrid2.DColumnas(5).Alignment = 1
DMGrid2.DColumnas(6).Alignment = 1
DMGrid2.DColumnas(7).Alignment = 1

DMGrid2.DColumnas(4).IsNumber = True
DMGrid2.DColumnas(5).IsNumber = True
DMGrid2.DColumnas(6).IsNumber = True
DMGrid2.DColumnas(7).IsNumber = True
DMGrid2.DColumnas(8).Visible = False

DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 10 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(4).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(5).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(6).Width = Val(DMGrid2.Width * 15 / 100) - 300
DMGrid2.DColumnas(7).Width = Val(DMGrid2.Width * 15 / 100)

DMGrid2.DColumnas(1).Caption = "Recibo"
DMGrid2.DColumnas(2).Caption = "Nombre"
DMGrid2.DColumnas(3).Caption = "Apellido"
DMGrid2.DColumnas(4).Caption = "Asignaciones"
DMGrid2.DColumnas(5).Caption = "Deducciones"
DMGrid2.DColumnas(6).Caption = "Otros"
DMGrid2.DColumnas(7).Caption = "Total Neto"
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
Periodo = DMGrid1.ValorCelda(lRow, 1)
Fecha1 = DMGrid1.ValorCelda(lRow, 2)
Fecha2 = DMGrid1.ValorCelda(lRow, 3)

If Not IsNull(Fecha1) Then
    If Trim(Fecha1) <> "" Then
        TAsign = 0
        TDeduc = 0
        TOtros = 0
        TNeto = 0

        CSql = "SELECT Recibos.*, Empleados.Nombre, Empleados.Apellido FROM Recibos INNER JOIN " & _
               " Empleados ON Recibos.IdEmpleado = Empleados.IdEmpleado WHERE Fecha_Ini_Nom='" & Fecha1 & "' " & _
               " AND Fecha_Fin_Nom='" & Fecha2 & "' ORDER BY Recibos.IdRecibos"
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount = 0 Then Exit Sub
        
        DMGrid2.Rows = 0
        While Not RsTemp.EOF
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsTemp.Fields("IdRecibos").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsTemp.Fields("Nombre").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsTemp.Fields("Apellido").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsTemp.Fields("Total_Asignacion").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = RsTemp.Fields("Total_Deducciones").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 6) = RsTemp.Fields("Total_Retenciones").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 7) = CDbl(RsTemp.Fields("Total_Asignacion").Value) - CDbl(RsTemp.Fields("Total_Deducciones").Value)
            DMGrid2.ValorCelda(DMGrid2.Rows, 8) = RsTemp.Fields("IdEmpleado").Value
            
            TAsign = TAsign + CDbl(RsTemp.Fields("Total_Asignacion").Value)
            TDeduc = TDeduc + CDbl(RsTemp.Fields("Total_Deducciones").Value)
            TOtros = TOtros + CDbl(RsTemp.Fields("Total_Retenciones").Value)
            TNeto = TNeto + CDbl(DMGrid2.ValorCelda(DMGrid2.Rows, 7))
            RsTemp.MoveNext
        Wend
    End If
End If

TxtAsignaciones.Text = Format(TAsign, "#,##0.00")
TxtDeducciones.Text = Format(TDeduc, "#,##0.00")
TxtOtros.Text = Format(TOtros, "#,##0.00")
TxtTotalNeto.Text = Format(TNeto, "#,##0.00")
DMGrid2.PaintMGrid
End Sub

Sub Cargar_Periodos_Nomina()
    
CSql = "SELECT Fecha_Ini_Nom, Fecha_Fin_Nom, Periodo FROM Recibos GROUP BY Fecha_Ini_Nom, Fecha_Fin_Nom, Periodo"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then MsgBox "No hay recibos creadas!", vbExclamation + vbOKOnly, "Nómina vacía": Exit Sub

RsTemp.MoveFirst
DMGrid1.Rows = 0
While Not RsTemp.EOF
    
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Periodo").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Fecha_Ini_Nom").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Fecha_Fin_Nom").Value
    RsTemp.MoveNext
Wend

DMGrid1.PaintMGrid
'Dat_Admin
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: TxtRecibo.ForeColor = TxtBuscar.ForeColor: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: TxtRecibo.ForeColor = TxtBuscar.ForeColor: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If UCase(TxtBuscar.Text) = UCase("Busqueda") Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_GotFocus()
If UCase(TxtBuscar.Text) = UCase("Busqueda") Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_LostFocus()
If Trim(TxtBuscar.Text) = "" Then TxtBuscar.Text = "Busqueda"
End Sub

Private Sub TxtRecibo_Click()
If UCase(TxtRecibo.Text) = UCase("Busqueda") Then TxtRecibo.Text = ""
End Sub

Private Sub TxtRecibo_DblClick()
TxtRecibo.Text = ""
End Sub

Private Sub TxtRecibo_GotFocus()
If UCase(TxtRecibo.Text) = UCase("Busqueda") Then TxtRecibo.Text = ""
End Sub

Private Sub TxtRecibo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then BtnBuscarRecibo_Click
End Sub

Private Sub TxtRecibo_LostFocus()
If Trim(TxtRecibo.Text) = "" Then TxtRecibo.Text = "Busqueda"
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnBuscarRecibo_Click()
Dim TamDMGrid2 As Integer
Dim i As Integer
If Not UCase(TxtRecibo.Text) = UCase("Busqueda") And Trim(TxtRecibo.Text) <> "" Then
    TamDMGrid2 = DMGrid2.Rows
    For i = 1 To TamDMGrid2
        If DMGrid2.ValorCelda(i, 1) = TxtRecibo.Text Then
            DMGrid2.Row = i
            Exit Sub
        End If
    Next i
End If

MsgBox "El Nro de Recibo Ingresado no se encuentró para periodo seleccionado!", vbExclamation + vbOKOnly, "No se encontro el Recibo"
End Sub




VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FacturacionRTLista 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Facturas"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   300
         Left            =   9000
         TabIndex        =   14
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   53477377
         CurrentDate     =   40361
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "        yyyy"
         Format          =   53477379
         CurrentDate     =   40361
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Definido"
         Height          =   255
         Index           =   2
         Left            =   7920
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por mes y año"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por año"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   383
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   4
         Top             =   6240
         Width           =   8295
         Begin ChamaleonButton.ChameleonBtn BtnImprimir 
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   16
            ToolTipText     =   "Imprimir Factura"
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Imprimir Lista"
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
            MICON           =   "FacturacionRTLista.frx":0000
            PICN            =   "FacturacionRTLista.frx":001C
            PICH            =   "FacturacionRTLista.frx":0141
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   720
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   7080
            TabIndex        =   5
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
            MICON           =   "FacturacionRTLista.frx":03D1
            PICN            =   "FacturacionRTLista.frx":03ED
            PICH            =   "FacturacionRTLista.frx":05B6
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
            TabIndex        =   6
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
            MICON           =   "FacturacionRTLista.frx":07EB
            PICN            =   "FacturacionRTLista.frx":0807
            PICH            =   "FacturacionRTLista.frx":0AA9
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
            Index           =   0
            Left            =   5400
            TabIndex        =   15
            ToolTipText     =   "Imprimir Factura"
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
            MICON           =   "FacturacionRTLista.frx":0E13
            PICN            =   "FacturacionRTLista.frx":0E2F
            PICH            =   "FacturacionRTLista.frx":0F54
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
            Left            =   4800
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowCloseBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Label NoReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
            Height          =   195
            Left            =   3720
            TabIndex        =   18
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registros:"
            Height          =   195
            Left            =   2760
            TabIndex        =   17
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   6240
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
            TabIndex        =   2
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Numero Factura"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   3
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
            MICON           =   "FacturacionRTLista.frx":11E4
            PICN            =   "FacturacionRTLista.frx":1200
            PICH            =   "FacturacionRTLista.frx":1465
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
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   9551
         Object.Width           =   12465
         Object.Height          =   5385
         ScrollBar       =   3
         MarqueeStyle    =   2
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   5640
         TabIndex        =   13
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "     MM/yyyy"
         Format          =   53477379
         CurrentDate     =   40361
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "FacturacionRTLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsFacturas As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim i As Integer

Sub Leer_Facturas(Cond As Boolean, CondWhere As String)


If Cond = False Then
    CSql = "SELECT C_Cobrar.N_Factura, C_Cobrar.Fecha, Paciente.ApellidoP, Paciente.NombreP, " & _
      " Paciente.Historia, Paciente.Cedulap, C_Cobrar.Monto, C_Cobrar.PorCobrar, C_Cobrar.Anulada " & _
      " FROM C_Cobrar INNER JOIN Paciente ON (C_Cobrar.IdPaciente = Paciente.IdPaciente) " & " ORDER BY C_Cobrar.N_Factura"
Else
    CSql = "SELECT C_Cobrar.N_Factura, C_Cobrar.Fecha, Paciente.ApellidoP, Paciente.NombreP, " & _
      " Paciente.Historia, Paciente.Cedulap, C_Cobrar.Monto, C_Cobrar.PorCobrar, C_Cobrar.Anulada " & _
      " FROM C_Cobrar INNER JOIN Paciente ON (C_Cobrar.IdPaciente = Paciente.IdPaciente) " & _
      CondWhere & " ORDER BY C_Cobrar.N_Factura"
End If

Set RsFacturas = CrearRS(CSql)

If RsFacturas.RecordCount = 0 Then
    MsgBox "No se encontraron facturas en la base de datos!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

End Sub


Sub Cargar_Facturas()

DMGrid1.Clear
DMGrid1.Rows = 0

If RsFacturas.RecordCount <= 0 Then
    MsgBox "No se encontraron facturas en la base de datos!", vbExclamation + vbOKOnly, "Información"
    Exit Sub
End If

While Not RsFacturas.EOF

    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsFacturas.Fields("N_Factura").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsFacturas.Fields("Fecha").Value
    
    If RsFacturas.Fields("Anulada").Value = "1" Then
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = "*** ANULADA ***"
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = " "
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = " "
        DMGrid1.ValorCelda(DMGrid1.Rows, 6) = " "
        DMGrid1.ValorCelda(DMGrid1.Rows, 7) = "0"
        DMGrid1.ValorCelda(DMGrid1.Rows, 8) = "0"
        DMGrid1.ValorCelda(DMGrid1.Rows, 9) = "0"
    Else
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsFacturas.Fields("ApellidoP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsFacturas.Fields("NombreP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsFacturas.Fields("Cedulap").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 6) = RsFacturas.Fields("Historia").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 7) = RsFacturas.Fields("Monto").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 8) = CDbl(RsFacturas.Fields("Monto").Value) - CDbl(RsFacturas.Fields("PorCobrar").Value)
        DMGrid1.ValorCelda(DMGrid1.Rows, 9) = RsFacturas.Fields("PorCobrar").Value
    End If
    
    
    
    RsFacturas.MoveNext
Wend

NoReg.Caption = RsFacturas.RecordCount
DMGrid1.RowBackColor 1, vbWhite
DMGrid1.PaintMGrid

End Sub


Sub IniDMGrid()
DMGrid1.Cols = 9


DMGrid1.DColumnas(1).Caption = "Nro Fact"
DMGrid1.DColumnas(2).Caption = "Fecha"
DMGrid1.DColumnas(3).Caption = "Apellido"
DMGrid1.DColumnas(4).Caption = "Nombre"
DMGrid1.DColumnas(5).Caption = "Cédula"
DMGrid1.DColumnas(6).Caption = "Historia"
DMGrid1.DColumnas(7).Caption = "Monto"
DMGrid1.DColumnas(8).Caption = "Cancelado"
DMGrid1.DColumnas(9).Caption = "Por Cobrar"


DMGrid1.DColumnas(7).IsNumber = True
DMGrid1.DColumnas(8).IsNumber = True
DMGrid1.DColumnas(9).IsNumber = True

DMGrid1.DColumnas(1).Width = ((DMGrid1.Width * 10) / 100)
DMGrid1.DColumnas(2).Width = ((DMGrid1.Width * 10) / 100)
DMGrid1.DColumnas(3).Width = ((DMGrid1.Width * 15) / 100)
DMGrid1.DColumnas(4).Width = ((DMGrid1.Width * 15) / 100)
DMGrid1.DColumnas(5).Width = ((DMGrid1.Width * 10) / 100)
DMGrid1.DColumnas(6).Width = ((DMGrid1.Width * 10) / 100)
DMGrid1.DColumnas(7).Width = ((DMGrid1.Width * 10) / 100)
DMGrid1.DColumnas(8).Width = ((DMGrid1.Width * 10) / 100)
DMGrid1.DColumnas(9).Width = ((DMGrid1.Width * 10) / 100)

End Sub

Private Sub BtnBuscar_Click()

CSql = "Select * From C_Cobrar Where N_Factura = '" & Trim(TxtBuscar.Text) & "' AND C_NC=0"
Set RsFacturas = CrearRS(CSql)

Cargar_Facturas

End Sub

Private Sub BtnCerrar_Click()

Unload Me

End Sub

Sub imprime()
On Error GoTo Wrr

N_Factur = Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))

With CrystalReport1
    .ReportFileName = RutaInformes & "\FacturaN.rpt"
    '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .Connect = "DSN=CrReporte;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{Factura.N_Factura} = " & N_Factur
    .WindowTitle = "Reporte Factura No. " & N_Factur
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With
Exit Sub
Wrr:
MsgBox Err.Number & " : " & Err.Description & " / " & Err.Source, vbExclamation + vbOKOnly, "Errores Generales!"
End Sub


Private Sub BtnImprimir_Click(Index As Integer)
On Error GoTo Wrr

Dim N_Factur

If DMGrid1.Rows = 0 Then MsgBox "Seleccione una Factura!", vbInformation + vbOKOnly, "Información": Exit Sub
If DMGrid1.Row = 0 Then MsgBox "Seleccione una Factura!", vbInformation + vbOKOnly, "Información": Exit Sub


 
If Index = 0 Then
    
    N_Factur = Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
    If Val(DMGrid1.ValorCelda(DMGrid1.Row, 1)) = 0 Then
    
        MsgBox "Seleccione una Factura!", vbInformation + vbOKOnly, "Información"
        Exit Sub
        
    End If
    
    CSql = "SELECT impresa,Anulada FROM C_Cobrar WHERE N_Factura = " & N_Factur
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount = 0 Then MsgBox "Problemas Generales!", vbInformation + vbOKOnly, "Error": Exit Sub
    
    If RsTemp.Fields("Anulada").Value = "1" Then
        MsgBox "No puede imprimir una factura anulada!", vbInformation + vbOKOnly, "Información"
        Exit Sub
    End If
    
    If Val(RsTemp.Fields("impresa").Value) = 0 Then
        Dim RsActualizarImpresa As New ADODB.Recordset
        imprime
        CSql = "Update c_cobrar set impresa = 1 WHERE N_Factura = " & N_Factur
        Set RsActualizarImpresa = CrearRS(CSql)
        Call Enviar_Bitacora(IdUser, "FACTURACION", "IMPRIMIR", "Se imprimio la factura Nro. " & N_Factur)
    Else
        Msg = "Ya esta factura fue impresa desea imprimir una copia ?"
        d = MsgBox(Msg, vbYesNo + vbInformation, "Factura Impresa")
        If d = 6 Then
            imprime
        End If
    End If
ElseIf Index = 1 Then
    

    Dim TamDMGrid As Integer
    Dim IniTamDMGrid As Integer
    Dim Rsp As Byte
    Dim Oopc
    
'    Rsp = MsgBox("Imprimir directamente!!", vbQuestion + vbYesNo, "Confirmar!")
    IniTamDMGrid = InputBox("Desde la fila:", "Ingrese el inicio de impresión!")
    TamDMGrid = InputBox("Hasta la fila:", "Ingrese el fin de impresión!")
'
'    If Rsp = vbNo Then
'        Oopc = 0 'crptToWindows
'    Else
'        Oopc = 1 'crptToPrinter
'        CrystalReport1.PrinterName = "doPDF v6"
'        CrystalReport1.PrinterPort = "DOP6"
'        CrystalReport1.PrinterDriver = "winspool"
'    End If
    
    
    
    For i = IniTamDMGrid To TamDMGrid
        
        N_Factur = Val(DMGrid1.ValorCelda(i, 1))
        
        CSql = "SELECT Anulada FROM C_Cobrar WHERE N_Factura = " & N_Factur
        Set RsTemp = CrearRS(CSql)
    
        If RsTemp.Fields("Anulada").Value <> "1" Then
            'MsgBox "No puede imprimir una factura anulada!", vbInformation + vbOKOnly, "Información"
            'Exit Sub
    
            With CrystalReport1
                .ReportFileName = RutaInformes & "\FacturaN.rpt"
                '.Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
                .Connect = "DSN=CrReporte;"
                .DiscardSavedData = True
                .RetrieveDataFiles
                .ReportSource = 0
                .SelectionFormula = "{Factura.N_Factura} = " & N_Factur
                .WindowTitle = "Reporte Factura No. " & N_Factur
                '.Destination = Oopc
                .Destination = crptToWindows
                '.Destination = crptToPrinter
                .PrintFileType = crptCrystal
                .WindowState = crptMaximized
                .WindowMaxButton = False
                .WindowMinButton = False
                .Action = 1
            End With
        End If
        
    Next i
End If

Exit Sub
Wrr:
    MsgBox Err.Number & " : " & Err.Description & " / " & Err.Source, vbExclamation + vbOKOnly, "Errores Generales!"

End Sub

Private Sub DMGrid1_DobleClick()

If DMGrid1.Rows = 0 Then MsgBox "Seleccione una Factura!", vbInformation + vbOKOnly, "Información": Exit Sub
If DMGrid1.Row = 0 Then MsgBox "Seleccione una Factura!", vbInformation + vbOKOnly, "Información": Exit Sub

FacturacionRT.TxtBuscar.Text = Trim(DMGrid1.ValorCelda(DMGrid1.Row, 1))
FacturacionRT.TxtBuscar_KeyPress 13

Unload Me

End Sub

Private Sub DTPicker1_Change()
Option1_Click 0
End Sub

Private Sub DTPicker2_Change()
Option1_Click 1
End Sub

Private Sub DTPicker3_Change()
Option1_Click 2
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid

Leer_Facturas False, ""

Cargar_Facturas
End Sub

Private Sub Option1_Click(Index As Integer)

If Option1(1).Value Then
    CSql = "SELECT C_Cobrar.N_Factura, C_Cobrar.Fecha, Paciente.ApellidoP, Paciente.NombreP, " & _
          " Paciente.Historia, Paciente.Cedulap, C_Cobrar.Monto, C_Cobrar.PorCobrar, C_Cobrar.Anulada " & _
          " FROM C_Cobrar INNER JOIN Paciente ON (C_Cobrar.IdPaciente = Paciente.IdPaciente) " & _
          " WHERE (CAST(MONTH(Fecha) as NVARCHAR)+'/'+CAST(YEAR(Fecha) as NVARCHAR))='" & Format(DTPicker2.Value, "M/yyyy") & "' ORDER BY C_Cobrar.N_Factura"
ElseIf Option1(0).Value Then
    CSql = "SELECT C_Cobrar.N_Factura, C_Cobrar.Fecha, Paciente.ApellidoP, Paciente.NombreP, " & _
          " Paciente.Historia, Paciente.Cedulap, C_Cobrar.Monto, C_Cobrar.PorCobrar, C_Cobrar.Anulada " & _
          " FROM C_Cobrar INNER JOIN Paciente ON (C_Cobrar.IdPaciente = Paciente.IdPaciente) " & _
          " WHERE CAST(YEAR(Fecha) as NVARCHAR)='" & Format(DTPicker1.Value, "yyyy") & "' ORDER BY C_Cobrar.N_Factura"
ElseIf Option1(2).Value Then
    CSql = "SELECT C_Cobrar.N_Factura, C_Cobrar.Fecha, Paciente.ApellidoP, Paciente.NombreP, " & _
          " Paciente.Historia, Paciente.Cedulap, C_Cobrar.Monto, C_Cobrar.PorCobrar, C_Cobrar.Anulada " & _
          " FROM C_Cobrar INNER JOIN Paciente ON (C_Cobrar.IdPaciente = Paciente.IdPaciente) " & _
          " WHERE (CAST(DAY(Fecha) as NVARCHAR)+'/'+CAST(MONTH(Fecha) as NVARCHAR)+'/'+CAST(YEAR(Fecha) as NVARCHAR))='" & Format(DTPicker3.Value, "d/M/yyyy") & "' ORDER BY C_Cobrar.N_Factura"
End If

Set RsFacturas = CrearRS(CSql)
Cargar_Facturas

End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then BtnBuscar_Click

End Sub

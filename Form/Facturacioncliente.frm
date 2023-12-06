VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReporteFacturacion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Facturacion"
   ClientHeight    =   4635
   ClientLeft      =   795
   ClientTop       =   1590
   ClientWidth     =   6090
   Icon            =   "Facturacioncliente.frx":0000
   LinkTopic       =   "Form37"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6090
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Facturación por Cliente"
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5895
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   3120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Orientación"
         Height          =   1095
         Left            =   4200
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
         Begin VB.OptionButton OptHorizontal 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Horizontal"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptVertical 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Vertical"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Destino"
         Height          =   1095
         Left            =   3480
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
         Begin VB.OptionButton OptPantalla 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Por Pantalla"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptImpresora 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Impresora"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Optarchivo 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Archivo"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtNoCopias 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Text            =   "1"
         Top             =   3615
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   3480
         TabIndex        =   13
         Top             =   3480
         Width           =   2295
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   1200
            TabIndex        =   14
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
            MICON           =   "Facturacioncliente.frx":1002
            PICN            =   "Facturacioncliente.frx":101E
            PICH            =   "Facturacioncliente.frx":11E7
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
            TabIndex        =   15
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
            MICON           =   "Facturacioncliente.frx":141C
            PICN            =   "Facturacioncliente.frx":1438
            PICH            =   "Facturacioncliente.frx":155D
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
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por Paciente:"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por fecha:"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por Factura:"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Por Cliente:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Facturacioncliente.frx":17ED
         Left            =   1440
         List            =   "Facturacioncliente.frx":17EF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17235969
         CurrentDate     =   39944
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17235969
         CurrentDate     =   39944
      End
      Begin MSComDlg.CommonDialog cdgMain 
         Left            =   600
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Copias"
         Height          =   195
         Left            =   1800
         TabIndex        =   26
         Top             =   3705
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   3090
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   1440
         TabIndex        =   16
         Top             =   2610
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   1530
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   2010
         Width           =   465
      End
   End
End
Attribute VB_Name = "FrmReporteFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bdata As New ADODB.Recordset
Dim bdata1 As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnImprimir_Click()
'CrystalReport1.ReportFileName = Direc & "\informes\Facturaxcliente.rpt"
''CrystalReport1.ReportFileName = Direc & "\Informes\Facturaxcliente.rpt"
For T = 0 To 3
If Option1(T).Value = True Then Exit For
Next T

Select Case T
    Case Is = 0
        FmlaText$ = "{facturacion.Idcliente} = " & Combo1.ItemData(Combo1.ListIndex)
    
    Case Is = 1
        FmlaText$ = "{facturacion.IdPaciente} = " & Combo2.ItemData(Combo2.ListIndex)
    Case Is = 2
        FmlaText$ = "{facturacion.n_factura} >= " & Val(Text1.Text) & " and {facturacion.n_factura} <= " & Val(Text2.Text)
    Case Is = 3
        'dif = DateDiff("d", CDate(DTPicker1.Value), CDate(DTPicker2.Value))
        FmlaText$ = "{facturacion.fecha} >= #" & Format(CDate(DTPicker1.Value), "YYYY/MM/DD") & "# and {facturacion.fecha} <= #" & Format(CDate(DTPicker2.Value), "YYYY/MM/DD") & "#"
        MsgBox sele
End Select

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

CSql = "Select * From Dat_Admin"
Dim RsRutaInforme As New ADODB.Recordset
Dim RutaInf As String
Set RsRutaInforme = CrearRS(CSql)

RutaInf = RsRutaInforme.Fields("RutaInforme").Value

CrystalReport1.WindowTitle = "Reporte de Nutrición"
'FmlaText$ = "{Historia_clinica.Fecha_Inicio} >= '" & dr & "' and {Historia_clinica.Fecha_Inicio} <= '" & dr2 & "'"
CrystalReport1.SelectionFormula = FmlaText$
CrystalReport1.ReportFileName = RutaInf & "\" & "Facturaxcliente.rpt"

     
On Error GoTo errorHandler
    CrystalReport1.Action = 1
    Exit Sub
errorHandler:
    MsgBox CrystalReport1.LastErrorString, 16, "Mensaje de Error"
    Exit Sub




End Sub

Private Sub Form_Load()
Centrar Me


CSql = "Select IdCliente, Razon From Cliente"
bdata.Open CSql, Cnn
If Not (bdata.EOF) Then
bdata.MoveFirst
Do While Not bdata.EOF
Combo1.AddItem bdata.Fields(1)
Combo1.ItemData(Combo1.NewIndex) = bdata.Fields(0)
bdata.MoveNext
Loop
bdata.Close
Else
bdata.Close
End If
Call Paci
End Sub
Sub Paci()
CSql = "Select IdPaciente, NombreP, ApellidoP from Paciente"
bdata1.Open CSql, Cnn
If Not (bdata1.EOF) Then
bdata1.MoveFirst
Do While Not bdata1.EOF
Combo2.AddItem bdata1.Fields("nombreP") & " " & bdata1.Fields("apellidoP")
Combo2.ItemData(Combo2.NewIndex) = bdata1.Fields("idpaciente")
bdata1.MoveNext
Loop
bdata1.Close
Else
bdata1.Close
End If
End Sub

Private Sub Option1_Click(Index As Integer)
For i = 0 To Option1.Count - 1
    If i <> Index Then
        T = Index
     End If
Next
End Sub

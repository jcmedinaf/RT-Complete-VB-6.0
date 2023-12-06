VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContListaComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Comprobantes"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
   Icon            =   "FrmContListaComprobantes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Saldo"
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   4
         Top             =   5760
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Total"
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   3
         Top             =   5760
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Detalle"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   2
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   11
         Top             =   6120
         Width           =   3135
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   720
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   2040
            TabIndex        =   7
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
            MICON           =   "FrmContListaComprobantes.frx":1002
            PICN            =   "FrmContListaComprobantes.frx":101E
            PICH            =   "FrmContListaComprobantes.frx":11E7
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
         TabIndex        =   10
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
            TabIndex        =   5
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por código, Fecha, Detalles, Total o saldo"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   6
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
            MICON           =   "FrmContListaComprobantes.frx":141C
            PICN            =   "FrmContListaComprobantes.frx":1438
            PICH            =   "FrmContListaComprobantes.frx":169D
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   1
         Top             =   5760
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   5760
         Width           =   975
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   9551
         Object.Width           =   7305
         Object.Height          =   5385
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   5760
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmContListaComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Public IdEmpresa As Integer

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 5
DMGrid1.Rows = 0

DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1

'DMGrid1.DColumnas(1).Locked = True
'DMGrid1.DColumnas(2).Locked = True
'DMGrid1.DColumnas(3).Locked = True
'DMGrid1.DColumnas(4).Locked = True
'DMGrid1.DColumnas(5).Locked = True

DMGrid1.DColumnas(4).IsNumber = True
DMGrid1.DColumnas(5).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 40 / 100) - 300
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Fecha"
DMGrid1.DColumnas(3).Caption = "Detalle de Movimiento"
DMGrid1.DColumnas(4).Caption = "Total"
DMGrid1.DColumnas(5).Caption = "Saldo"
End Sub

Sub Cargar_Comprobantes()

If IdEmpresa = 0 Then MsgBox "No se encontraron Comprobante registrados!", vbInformation + vbOKOnly, "Información": Exit Sub

CSql = "SELECT * FROM ContComprobantes WHERE IdEmpresa=" & IdEmpresa & " AND Activo='1' ORDER BY IdComprobante"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

DMGrid1.Rows = 0

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("IdComprobante").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Fecha").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Detalle").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsTemp.Fields("Total").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsTemp.Fields("Saldo").Value
    RsTemp.MoveNext
Wend

DMGrid1.PaintMGrid
End Sub

Private Sub BtnBuscar_Click()
On Error GoTo MostrarError
If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From ContComprobantes Where IdComprobante = " & Val(TxtBuscar.Text) & " OR Fecha like '%" & Trim(TxtBuscar.Text) & "%' OR Detalle like '%" & Trim(TxtBuscar.Text) & "%' AND Activo='1' AND IdEmpresa=" & IdEmpresa
Else
    CSql = "Select * From ContComprobantes WHERE Activo='1' AND IdEmpresa=" & IdEmpresa
End If

Set RsTemp = CrearRS(CSql)

DMGrid1.Rows = 0

If RsTemp.RecordCount = 0 Then MsgBox "No existe esa referencia buscada", vbOKOnly, "Sin Resultado": Exit Sub
RsTemp.MoveFirst

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("IdComprobante")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Fecha")
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Detalle")
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsTemp.Fields("Total")
    DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsTemp.Fields("Saldo")
    RsTemp.MoveNext
Wend
DMGrid1.PaintMGrid

Exit Sub

MostrarError:
    MsgBox "Ha habido un error interno en cuanto a la busqueda! trate de colocar solo numeros y/o letras." & Chr(13) & _
    "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_DobleClick()
Dim IdComprobante As Integer
    If DMGrid1.Row = 0 Then Exit Sub
    IdComprobante = DMGrid1.ValorCelda(DMGrid1.Row, 1)
    
    If IdComprobante = 0 Then Exit Sub
    
    FrmContComprobante.Blanqueo
    
    CSql = "SELECT * FROM ContComprobantes WHERE IdEmpresa=" & IdEmpresa & " AND IdComprobante=" & IdComprobante & "AND Activo='1'"
    Set RsTemp = CrearRS(CSql)
    
    FrmContComprobante.TxtNoComprobante.Text = RsTemp.Fields("NroComprobante").Value
    FrmContComprobante.TxtDetalle.Text = RsTemp.Fields("Detalle").Value
    FrmContComprobante.DTPicker1.Value = CDate(RsTemp.Fields("Fecha").Value)
    FrmContComprobante.TxtSaldo.Text = RsTemp.Fields("Saldo").Value
    FrmContComprobante.Cantidad = Format(RsTemp.Fields("Saldo").Value, "#,##0.00")
    
    If RsTemp.RecordCount > 0 Then
        FrmContComprobante.DMGrid1.Rows = 0
        Call FrmContComprobante.Cargar_Renglones(IdComprobante)
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid
Cargar_Comprobantes
End Sub

Private Sub Option1_Click(Index As Integer)

If IdEmpresa = 0 Then Exit Sub

If Index = 0 Then
    CSql = "Select * From ContComprobantes Where activo='1' And IdEmpresa=" & IdEmpresa & " ORDER BY IdComprobante"
ElseIf Index = 1 Then
    CSql = "Select * From ContComprobantes Where activo='1' And IdEmpresa=" & IdEmpresa & " order by Fecha"
ElseIf Index = 2 Then
    CSql = "Select * From ContComprobantes Where activo='1' And IdEmpresa=" & IdEmpresa & " order by Detalle"
ElseIf Index = 3 Then
    CSql = "Select * From ContComprobantes Where activo='1' And IdEmpresa=" & IdEmpresa & " order by Total"
ElseIf Index = 4 Then
    CSql = "Select * From ContComprobantes Where activo='1' And IdEmpresa=" & IdEmpresa & " order by Saldo"
End If

Set RsTemp = CrearRS(CSql)
DMGrid1.Rows = 0
If RsTemp.RecordCount = 0 Then Exit Sub
RsTemp.MoveFirst

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("IdComprobante")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Fecha")
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Detalle")
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsTemp.Fields("Total")
    DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsTemp.Fields("Saldo")
    RsTemp.MoveNext
Wend
DMGrid1.PaintMGrid
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub



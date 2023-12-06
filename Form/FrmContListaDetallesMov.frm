VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContListaDetallesMov 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de Movimientos"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   Icon            =   "FrmContListaDetallesMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   5760
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   6
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   3
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
            TabIndex        =   4
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Código o Descripción"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   5
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
            MICON           =   "FrmContListaDetallesMov.frx":1002
            PICN            =   "FrmContListaDetallesMov.frx":101E
            PICH            =   "FrmContListaDetallesMov.frx":1283
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   1
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
            TabIndex        =   2
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
            MICON           =   "FrmContListaDetallesMov.frx":1515
            PICN            =   "FrmContListaDetallesMov.frx":1531
            PICH            =   "FrmContListaDetallesMov.frx":16FA
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
         TabIndex        =   9
         Top             =   5760
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmContListaDetallesMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IdEmpresa As Integer
Dim RsTemp As Recordset

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 2
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 80 / 100) - 300
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Detalle"
End Sub

Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From ContDetallesMovimientos Where Codigo like '%" & Trim(TxtBuscar.Text) & "%' OR Descripcion like '%" & Trim(TxtBuscar.Text) & "%' AND Activo='1' AND IdEmpresa" & IdEmpresa
Else
    CSql = "Select * From ContDetallesMovimientos WHERE Activo='1' AND IdEmpresa" & IdEmpresa
End If

Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsTemp.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Codigo").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Descripcion").Value
        RsTemp.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "No existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    CSql = "Select * From ContDetallesMovimientos WHERE IdEmpresa=" & IdEmpresa & " AND Codigo='" & DMGrid1.ValorCelda(lRow, 1) & "' AND Activo=1"
    Set RsTemp = CrearRS(CSql)
    If RsTemp.RecordCount > 0 Then
    
        If Tipo = "Comprobante" Then
            FrmContComprobante.DMGrid1.ValorCelda(FrmContComprobante.DMGrid1.Row, 2) = Trim(RsTemp.Fields("Descripcion").Value)
            FrmContComprobante.DMGrid1.PaintMGrid
        ElseIf Tipo = "Detalle de Movimientos" Then
            FrmContDetallesMovimientos.TxtCodigo.Text = Trim(RsTemp.Fields("Codigo").Value)
            FrmContDetallesMovimientos.TxtDescripcion.Text = Trim(RsTemp.Fields("Descripcion").Value)
            FrmContDetallesMovimientos.NewReg = 2
            FrmContDetallesMovimientos.IdDeta = Val(RsTemp.Fields("IdDetalle").Value)
        End If
    End If
    Unload Me
End If
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid
Option1_Click (0)
End Sub

Private Sub Option1_Click(Index As Integer)

If IdEmpresa = 0 Then Exit Sub
If Index = 0 Then
    CSql = "Select * From ContDetallesMovimientos where activo='1' AND IdEmpresa=" & IdEmpresa & " order by Codigo"
ElseIf Index = 1 Then
    CSql = "Select * From ContDetallesMovimientos where activo='1' AND IdEmpresa=" & IdEmpresa & " order by Descripcion"
End If

Set RsTemp = CrearRS(CSql)
DMGrid1.Rows = 0
If RsTemp.RecordCount = 0 Then Exit Sub
RsTemp.MoveFirst

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Codigo")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Descripcion")
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



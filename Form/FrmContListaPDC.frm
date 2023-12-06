VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContListaPDC 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas Actual"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "FrmContListaPDC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Plan de Cuenta"
      Height          =   6975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Nombre"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   2
         Top             =   5760
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   10
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
            TabIndex        =   6
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
            MICON           =   "FrmContListaPDC.frx":1002
            PICN            =   "FrmContListaPDC.frx":101E
            PICH            =   "FrmContListaPDC.frx":11E7
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
         TabIndex        =   9
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
            ToolTipText     =   "Ingrese la busqueda por Código, Identificador, Nombre o Grupo"
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
            MICON           =   "FrmContListaPDC.frx":141C
            PICN            =   "FrmContListaPDC.frx":1438
            PICH            =   "FrmContListaPDC.frx":169D
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
         Caption         =   "Identificador"
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Grupo"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   3
         Top             =   5760
         Width           =   1455
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   11
         Top             =   5760
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmContListaPDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdPDC As Integer
Public IdEmpresa As Integer
Dim RsTemp As Recordset
Dim RsCargarListaPDC As Recordset
Dim RsCargarListaEmpresas As Recordset
Dim i As Integer

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 4
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 0
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 50 / 100) - 300
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Identificador"
DMGrid1.DColumnas(3).Caption = "Nombre"
DMGrid1.DColumnas(4).Caption = "Grupo"
End Sub

Private Sub BtnBuscar_Click()

Dim TamDMGrid As Integer
Dim Reng1 As Integer
Dim Reng2 As String
Dim Reng3 As String

TamDMGrid = DMGrid1.Rows

For i = 1 To TamDMGrid
    Reng1 = DMGrid1.ValorCelda(i, 1)
    Reng2 = Trim(DMGrid1.ValorCelda(i, 2))
    Reng3 = Trim(DMGrid1.ValorCelda(i, 3))
    
    If Val(Trim(TxtBuscar.Text)) = Reng1 Or UCase(Trim(TxtBuscar.Text)) = UCase(Reng2) Or UCase(Trim(TxtBuscar.Text)) = UCase(Reng3) Then
        DMGrid1.Row = i
        Exit Sub
    End If
Next

MsgBox "No se encontraron resultados!", vbInformation + vbOKOnly, "Información"
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    CSql = "Select * From ContPDC WHERE IdPDC=" & DMGrid1.ValorCelda(lRow, 1) & " AND Activo=1"
    Set RsTemp = CrearRS(CSql)
    If RsTemp.RecordCount > 0 Then
        If Tipo = "UA" Then
            IdPDC = 2
            FrmCONTPDCConfig.TxtUA.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtUA.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "PA" Then
            FrmCONTPDCConfig.TxtPA.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtPA.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "UAI" Then
            FrmCONTPDCConfig.TxtUAI.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtUAI.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "PAI" Then
            FrmCONTPDCConfig.TxtPAI.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtPAI.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "UE" Then
            FrmCONTPDCConfig.TxtUE.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtUE.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        'ElseIf Tipo = "PE" Then
        ElseIf Tipo = "PE" Then
            FrmCONTPDCConfig.TxtPE.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtPE.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "UEI" Then
            FrmCONTPDCConfig.TxtUEI.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtUEI.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "PEI" Then
            FrmCONTPDCConfig.TxtPEI.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmCONTPDCConfig.TxtPEI.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "PDCFormato1" Then
            FrmContPDC.TxtFormato1.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmContPDC.TxtFormato1.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "PDCFormato2" Then
            FrmContPDC.TxtFormato2.Text = Trim(RsTemp.Fields("Identificador").Value)
            FrmContPDC.TxtFormato2.ToolTipText = Trim(RsTemp.Fields("Nombre").Value)
        ElseIf Tipo = "Comprobante" Then
            'FrmContComprobante
            FrmContComprobante.DMGrid1.ValorCelda(FrmContComprobante.DMGrid1.Row, 1) = Trim(RsTemp.Fields("Identificador").Value)
            FrmContComprobante.DMGrid1.PaintMGrid
            'MMMMMMMMMMMMMMMMM
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
    CSql = "Select * From ContPDC Where IdEmpresa=" & IdEmpresa & " AND Activo=1 order by IdPDC"
ElseIf Index = 1 Then
    CSql = "Select * From ContPDC Where IdEmpresa=" & IdEmpresa & " AND Activo=1 order by Identificador"
ElseIf Index = 2 Then
    CSql = "Select * From ContPDC Where IdEmpresa=" & IdEmpresa & " AND Activo=1 order by Nombre"
ElseIf Index = 3 Then
    CSql = "Select * From ContPDC Where IdEmpresa=" & IdEmpresa & " AND Activo=1 order by Tipo"
End If

Set RsCargarListaPDC = CrearRS(CSql)
DMGrid1.Rows = 0

If RsCargarListaPDC.RecordCount = 0 Then Exit Sub
RsCargarListaPDC.MoveFirst

While Not RsCargarListaPDC.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListaPDC.Fields("IdPDC")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListaPDC.Fields("Identificador")
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListaPDC.Fields("Nombre")
    
    If RsCargarListaPDC.Fields("Movimiento") Then
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Mvto"
    Else
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Grupo"
    End If
    RsCargarListaPDC.MoveNext
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

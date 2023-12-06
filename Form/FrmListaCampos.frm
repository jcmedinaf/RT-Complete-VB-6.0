VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListaCampos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Campos de Nomina"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "FrmListaCampos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7695
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   8
         Top             =   5760
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   5760
         Width           =   975
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
            TabIndex        =   5
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Código o Nombre"
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            ToolTipText     =   "Ingrese la busqueda por Nombre, Razon Social, Cédula o Rif"
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
            MICON           =   "FrmListaCampos.frx":1002
            PICN            =   "FrmListaCampos.frx":101E
            PICH            =   "FrmListaCampos.frx":1283
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
            ToolTipText     =   "Cerrar Tablas de Pacientes"
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
            MICON           =   "FrmListaCampos.frx":1515
            PICN            =   "FrmListaCampos.frx":1531
            PICH            =   "FrmListaCampos.frx":16FA
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
         TabIndex        =   6
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
         Left            =   480
         TabIndex        =   9
         Top             =   5760
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmListaCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemp As New ADODB.Recordset
Dim RsCampos As New ADODB.Recordset

Private Sub BtnBuscar_Click()
If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From CamposDeNomina Where Campo like '%" & Trim(TxtBuscar.Text) & "%' AND activo=1"
Else
    CSql = "Select * From CamposDeNomina where activo=1"
End If

Dim RsBuscarCliente As New ADODB.Recordset

Set RsBuscarCliente = CrearRS(CSql)

If RsBuscarCliente.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarCliente.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarCliente.Fields("IdCampoNomina").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarCliente.Fields("Campo").Value
        RsBuscarCliente.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If
RsBuscarCliente.Close
MsgBox DMGrid1.Rows
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    CSql = "Select * From CamposDeNomina Where activo=1 AND IdCampoNomina=" & DMGrid1.ValorCelda(lRow, 1)
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount > 0 Then
    
        Dim TamDGrid As Integer
        Dim Pos As Integer
        
        TamDGrid = FrmValoresCampoTrabajador.DMGrid1.Rows

        For i = 1 To TamDGrid
            If Trim(FrmValoresCampoTrabajador.DMGrid1.ValorCelda(i, 1)) = "" Then
                Pos = i
                Exit For
            Else
                If Val(FrmValoresCampoTrabajador.DMGrid1.ValorCelda(i, 1)) = Val(RsTemp.Fields("IdCampoNomina").Value) Then
                    MsgBox "El Campo seleccionado ya se encuentra en el perfil!", vbExclamation + vbOKOnly, "Información"
                    GoTo saltar_a
                End If
            End If
        Next i
        
        If Pos <= TamDGrid And Trim(FrmValoresCampoTrabajador.DMGrid1.ValorCelda(Pos, 1)) = "" Then
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(Pos, 1) = Format(RsTemp.Fields("IdCampoNomina").Value, "00000")
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(Pos, 2) = RsTemp.Fields("Campo").Value
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(Pos, 3) = 0
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(Pos, 4) = 0
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(Pos, 5) = 0
        Else
            FrmValoresCampoTrabajador.DMGrid1.Rows = FrmValoresCampoTrabajador.DMGrid1.Rows + 1
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(FrmValoresCampoTrabajador.DMGrid1.Rows, 1) = Format(RsTemp.Fields("IdCampoNomina").Value, "00000")
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(FrmValoresCampoTrabajador.DMGrid1.Rows, 2) = RsTemp.Fields("Campo").Value
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(FrmValoresCampoTrabajador.DMGrid1.Rows, 3) = 0
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(FrmValoresCampoTrabajador.DMGrid1.Rows, 4) = 0
            FrmValoresCampoTrabajador.DMGrid1.ValorCelda(FrmValoresCampoTrabajador.DMGrid1.Rows, 5) = 0
        End If
            FrmValoresCampoTrabajador.DMGrid1.PaintMGrid
    End If
saltar_a:
    RsTemp.Close
    Unload Me
End If
    
End Sub

Private Sub Form_Load()
IniDMGrid
Centrar Me

CSql = "Select * From CamposDeNomina where activo=1 order by IdCampoNomina"
Set RsCampos = CrearRS(CSql)

DMGrid1.Rows = 0
If RsCampos.RecordCount = 0 Then Exit Sub
RsCampos.MoveFirst

While Not RsCampos.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCampos.Fields("IdCampoNomina")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCampos.Fields("Campo")
    RsCampos.MoveNext
Wend
DMGrid1.PaintMGrid

End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    CSql = "Select * From CamposDeNomina where activo=1 order by IdCampoNomina"
ElseIf Index = 1 Then
    CSql = "Select * From CamposDeNomina where activo=1 order by Campo"
End If

Set RsCampos = CrearRS(CSql)
DMGrid1.Rows = 0
If RsCampos.RecordCount = 0 Then Exit Sub
RsCampos.MoveFirst

While Not RsCampos.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCampos.Fields("IdCampoNomina")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCampos.Fields("Campo")
    RsCampos.MoveNext
Wend
DMGrid1.PaintMGrid
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Change()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub


Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 2
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 80 / 100) - 300
'DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Campo"
End Sub

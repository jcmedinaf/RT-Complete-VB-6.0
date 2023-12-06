VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContListaEmpresas 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Empresas"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
   Icon            =   "FrmContListaEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7785
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Rif"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   7
         Top             =   5760
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   5760
         Width           =   975
      End
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
            ToolTipText     =   "Ingrese la busqueda por Código, Razon Social o Rif."
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   4
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
            MICON           =   "FrmContListaEmpresas.frx":1002
            PICN            =   "FrmContListaEmpresas.frx":101E
            PICH            =   "FrmContListaEmpresas.frx":1283
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
            MICON           =   "FrmContListaEmpresas.frx":1515
            PICN            =   "FrmContListaEmpresas.frx":1531
            PICH            =   "FrmContListaEmpresas.frx":16FA
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
         Left            =   240
         TabIndex        =   10
         Top             =   5760
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmContListaEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pa As Integer
Dim RsCargarListaEmpresas As New Recordset
Dim RsTemp As New Recordset
Dim IdEmpresa As Integer

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 3
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre de la Empresa"
DMGrid1.DColumnas(3).Caption = "Rif"
End Sub

Private Sub BtnBuscar_Click()
On Error GoTo MostrarError
If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From ContEmpresas Where Nombre like '%" & Trim(TxtBuscar.Text) & "%' OR Rif like '%" & Trim(TxtBuscar.Text) & "%' AND Activo=1"
Else
    CSql = "Select * From ContEmpresas WHERE Activo=1"
End If

Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsTemp.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("IdEmpresa").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Nombre").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Rif").Value
        RsTemp.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "No existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If

Exit Sub
MostrarError:
    MsgBox "Ha habido un error interno en cuanto a la busqueda! trate de colocar solo numeros y/o letras." & Chr(13) & _
    "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_DobleClick()

If DMGrid1.Row = 0 Then Exit Sub
If IsNull(DMGrid1.ValorCelda(DMGrid1.Row, 1)) Then Exit Sub
If IsEmpty(DMGrid1.ValorCelda(DMGrid1.Row, 1)) Then Exit Sub

CSql = "Select * From ContEmpresas WHERE Activo=1 AND IdEmpresa=" & DMGrid1.ValorCelda(DMGrid1.Row, 1)
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount > 0 Then

    If Tipo = "Empresa" Then
        IdEmpresa = Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
        FrmContEmpresas.TxtRif.Text = RsTemp.Fields("Rif").Value
        FrmContEmpresas.TxtNombre.Text = RsTemp.Fields("Nombre").Value
        FrmContEmpresas.TxtDireccion.Text = RsTemp.Fields("Direccion").Value
        FrmContEmpresas.DTPicker1.Value = Format(RsTemp.Fields("FechaIngreso").Value, "dd/MM/yyyy")
        
        If RsTemp.Fields("Consolidadora").Value Then FrmContEmpresas.ChkConsolidadora.Value = 1 Else FrmContEmpresas.ChkConsolidadora.Value = 0
        
        For i = 0 To FrmContEmpresas.CboCodigo.ListCount - 1
            If RsTemp.Fields("CodigoTelf").Value = FrmContEmpresas.CboCodigo.List(i) Then
                FrmContEmpresas.CboCodigo.ListIndex = i
                Exit For
            Else
                FrmContEmpresas.CboCodigo.ListIndex = -1
            End If
        Next i
        
        If Not IsNull(RsTemp.Fields("Telefono").Value) Then
            FrmContEmpresas.TxtTelefono.Text = RsTemp.Fields("Telefono").Value
        Else
            FrmContEmpresas.TxtTelefono.Text = ""
        End If
        
        If Not IsNull(RsTemp.Fields("Ciudad").Value) Then
            FrmContEmpresas.TxtCiudad.Text = RsTemp.Fields("Ciudad").Value
        Else
            FrmContEmpresas.TxtCiudad.Text = ""
        End If
        
        If Not IsNull(RsTemp.Fields("Clave").Value) Then
            FrmContEmpresas.TxtClave.Text = RsTemp.Fields("Clave").Value
        Else
            FrmContEmpresas.TxtClave.Text = ""
        End If
    ElseIf Tipo = "Comprobante" Then
        FrmContComprobante.IdEmpresa = Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
        FrmContComprobante.Caption = "Comprabante Contable para la empresa '" & RsTemp.Fields("Nombre").Value & "'"
    ElseIf Tipo = "Detalle de Movimientos" Then
        FrmContDetallesMovimientos.IdEmpresa = Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
        FrmContDetallesMovimientos.Caption = "Detalles de Movimientos de la empresa '" & RsTemp.Fields("Nombre").Value & "'"
        FrmContDetallesMovimientos.Cargar_Det_Mov
    ElseIf UCase(Tipo) = UCase("general") Then
        MsgBox "Se trabajará con la empresa '" & Trim(RsTemp.Fields("Nombre").Value) & "'", vbInformation + vbOKOnly, "Operación Exitosa."
        IdEmprs = 1
    End If
End If
Unload Me

End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    CSql = "Select * From ContEmpresas WHERE Activo=1 AND IdEmpresa=" & DMGrid1.ValorCelda(lRow, 1)
    Set RsTemp = CrearRS(CSql)
    If RsTemp.RecordCount > 0 Then
    
        If Tipo = "Empresa" Then
            IdEmpresa = Val(DMGrid1.ValorCelda(lRow, 1))
            FrmContEmpresas.TxtRif.Text = RsTemp.Fields("Rif").Value
            FrmContEmpresas.TxtNombre.Text = RsTemp.Fields("Nombre").Value
            FrmContEmpresas.TxtDireccion.Text = RsTemp.Fields("Direccion").Value
            FrmContEmpresas.DTPicker1.Value = Format(RsTemp.Fields("FechaIngreso").Value, "dd/MM/yyyy")
            
            If RsTemp.Fields("Consolidadora").Value Then FrmContEmpresas.ChkConsolidadora.Value = 1 Else FrmContEmpresas.ChkConsolidadora.Value = 0
            
            For i = 0 To FrmContEmpresas.CboCodigo.ListCount - 1
                If RsTemp.Fields("CodigoTelf").Value = FrmContEmpresas.CboCodigo.List(i) Then
                    FrmContEmpresas.CboCodigo.ListIndex = i
                    Exit For
                Else
                    FrmContEmpresas.CboCodigo.ListIndex = -1
                End If
            Next i
    
            If Not IsNull(RsTemp.Fields("Telefono").Value) Then
                FrmContEmpresas.TxtTelefono.Text = RsTemp.Fields("Telefono").Value
            Else
                FrmContEmpresas.TxtTelefono.Text = ""
            End If
            
            If Not IsNull(RsTemp.Fields("Ciudad").Value) Then
                FrmContEmpresas.TxtCiudad.Text = RsTemp.Fields("Ciudad").Value
            Else
                FrmContEmpresas.TxtCiudad.Text = ""
            End If
            
            If Not IsNull(RsTemp.Fields("Clave").Value) Then
                FrmContEmpresas.TxtClave.Text = RsTemp.Fields("Clave").Value
            Else
                FrmContEmpresas.TxtClave.Text = ""
            End If
        ElseIf Tipo = "Comprobante" Then
            FrmContComprobante.IdEmpresa = Val(DMGrid1.ValorCelda(lRow, 1))
            FrmContComprobante.Caption = "Comprabante Contable para la empresa '" & RsTemp.Fields("Nombre").Value & "'"
        ElseIf Tipo = "Detalle de Movimientos" Then
            FrmContDetallesMovimientos.IdEmpresa = Val(DMGrid1.ValorCelda(lRow, 1))
            FrmContDetallesMovimientos.Caption = "Detalles de Movimientos de la empresa '" & RsTemp.Fields("Nombre").Value & "'"
            FrmContDetallesMovimientos.Cargar_Det_Mov
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

If Index = 0 Then
    CSql = "Select * From ContEmpresas where activo=1 order by IdEmpresa"
ElseIf Index = 1 Then
    CSql = "Select * From ContEmpresas where activo=1 order by Nombre"
ElseIf Index = 2 Then
    CSql = "Select * From ContEmpresas where activo=1 order by Rif"
End If

Set RsCargarListaEmpresas = CrearRS(CSql)
DMGrid1.Rows = 0
If RsCargarListaEmpresas.RecordCount = 0 Then Exit Sub
RsCargarListaEmpresas.MoveFirst

While Not RsCargarListaEmpresas.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListaEmpresas.Fields("IdEmpresa")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListaEmpresas.Fields("Nombre")
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListaEmpresas.Fields("Rif")
    RsCargarListaEmpresas.MoveNext
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


VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoClientes 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   7035
   ClientLeft      =   5760
   ClientTop       =   2490
   ClientWidth     =   7785
   Icon            =   "Listaclientes.frx":0000
   LinkTopic       =   "Form31"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7785
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   4
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
            TabIndex        =   5
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
            MICON           =   "Listaclientes.frx":1002
            PICN            =   "Listaclientes.frx":101E
            PICH            =   "Listaclientes.frx":11E7
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
         TabIndex        =   1
         Top             =   6120
         Width           =   3975
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
            MICON           =   "Listaclientes.frx":141C
            PICN            =   "Listaclientes.frx":1438
            PICH            =   "Listaclientes.frx":169D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
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
            ToolTipText     =   "Ingrese la busqueda por Nombre o Rif del Cliente"
            Top             =   240
            Width           =   2175
         End
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5775
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   10186
         Object.Width           =   7305
         Object.Height          =   5745
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pa As Integer
Dim RsCargarListaCliente As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Cliente Where Razon like '%" & Trim(TxtBuscar.Text) & "%' OR Rif like '%" & Trim(TxtBuscar.Text) & "%'"
Else
    CSql = "Select * From Cliente"
End If

Dim RsBuscarCliente As New ADODB.Recordset

Set RsBuscarCliente = CrearRS(CSql)

If RsBuscarCliente.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarCliente.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarCliente.Fields("IdCliente").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarCliente.Fields("Razon").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarCliente.Fields("Rif").Value
        RsBuscarCliente.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If
RsBuscarCliente.Close

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
DataGrid1.Col = 0
IdCliente = Val(DataGrid1.Text)
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    Dim RsSeleccionarCliente As New ADODB.Recordset
    CSql = "Select * From Cliente Where IdCliente='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsSeleccionarCliente = CrearRS(CSql)
    If RsSeleccionarCliente.RecordCount > 0 Then
        
        If ModulO = 0 Then
            FrmDatosClientes.CodClient = RsSeleccionarCliente.Fields("IdCliente").Value
            FrmDatosClientes.RsClientes.Find "IdCliente='" & FrmDatosClientes.CodClient & "'"
            FrmDatosClientes.CargaDatos
            
        ElseIf ModulO = 1 Then
            FacturacionRT.IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
            FacturacionRT.Text1.Text = RsSeleccionarCliente.Fields("Rif").Value
            FacturacionRT.Text2.Text = RsSeleccionarCliente.Fields("Razon").Value
            FacturacionRT.Text3.Text = RsSeleccionarCliente.Fields("Direccionc").Value
            If Not IsNull(RsSeleccionarCliente.Fields("Telefono").Value) Then FacturacionRT.Text4.Text = RsSeleccionarCliente.Fields("Telefono").Value Else FacturacionRT.Text4.Text = ""
        ElseIf ModulO = 2 Then
            IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
            With FrmPresupuestoTratamientos
                .Text8.Visible = False
                .Combo1.Visible = True
                For i = 0 To .Combo1.ListCount - 1
                    If .Combo1.ItemData(i) = IdCliente Then
                        .Combo1.ListIndex = i
                        .Combo1_Click
                        Exit For
                    End If
                Next
                .BtnBuscarClientes_Click
                .BtnGuardarActualizar.Enabled = True
                
            End With
            
        End If
        
    End If
    RsSeleccionarCliente.Close
    Unload Me
End If
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid

CSql = "Select * From Cliente"
Set RsCargarListaCliente = CrearRS(CSql)

Do While Not RsCargarListaCliente.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListaCliente.Fields("IdCliente").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListaCliente.Fields("Razon").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListaCliente.Fields("Rif").Value
    RsCargarListaCliente.MoveNext
Loop

DMGrid1.PaintMGrid

End Sub

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
DMGrid1.DColumnas(2).Caption = "Razón Social"
DMGrid1.DColumnas(3).Caption = "Rif"
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
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

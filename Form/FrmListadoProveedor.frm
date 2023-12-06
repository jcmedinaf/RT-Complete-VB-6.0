VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoProveedor 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   7170
   ClientLeft      =   6375
   ClientTop       =   1875
   ClientWidth     =   7800
   Icon            =   "FrmListadoProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7800
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   3
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
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Rif o Nombre del Proveedor"
            Top             =   240
            Width           =   2415
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2640
            TabIndex        =   4
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Buscar"
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
            MICON           =   "FrmListadoProveedor.frx":1002
            PICN            =   "FrmListadoProveedor.frx":101E
            PICH            =   "FrmListadoProveedor.frx":1283
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
         Top             =   6240
         Width           =   3255
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   1320
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   2160
            TabIndex        =   2
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
            MICON           =   "FrmListadoProveedor.frx":1515
            PICN            =   "FrmListadoProveedor.frx":1531
            PICH            =   "FrmListadoProveedor.frx":16FA
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
         Height          =   5895
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   10398
         Object.Width           =   7305
         Object.Height          =   5865
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBuscar_Click()
Dim RsBuscarProveedores As New ADODB.Recordset
If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Proveedores Where RifProveedor like '%" & Trim(TxtBuscar.Text) & "%' Or Nombre like '%" & Trim(TxtBuscar.Text) & "%'"
Else
    CSql = "Select * From Proveedores"
End If


Set RsBuscarProveedores = CrearRS(CSql)

If RsBuscarProveedores.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarProveedores.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarProveedores.Fields("IdProveedor").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarProveedores.Fields("Nombre").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarProveedores.Fields("RifProveedor").Value
        RsBuscarProveedores.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If
RsBuscarProveedores.Close

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    Dim RsCargarListadoProveedores As New ADODB.Recordset
    CSql = "Select * From Proveedores Where IdProveedor='" & DMGrid1.ValorCelda(lRow, 1) & "' And Status='1'"
    Set RsCargarListadoProveedores = CrearRS(CSql)
    
    If RsCargarListadoProveedores.RecordCount > 0 Then
        Select Case Tipo
            
            Case Is = "Compras"
                IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
                FrmCompras.TxtCodigoProveedor.Text = RsCargarListadoProveedores.Fields("IdProveedor").Value
                FrmCompras.TxtDescripcionProveerdor.Text = RsCargarListadoProveedores.Fields("Nombre").Value
                FrmCompras.TxtRif.Text = RsCargarListadoProveedores.Fields("RifProveedor").Value
            
            Case Is = "Ordenes"
                IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
                FrmOrdenCompra.TxtCodigoProveedor.Text = RsCargarListadoProveedores.Fields("IdProveedor").Value
                FrmOrdenCompra.TxtDescripcionProveerdor.Text = RsCargarListadoProveedores.Fields("Nombre").Value
                FrmOrdenCompra.TxtRif.Text = RsCargarListadoProveedores.Fields("RifProveedor").Value
                If RsCargarListadoProveedores.Fields("Status").Value = 1 Then
                    FrmOrdenCompra.TxtStatus.Text = "Activo"
                Else
                    FrmOrdenCompra.TxtStatus.Text = "Inactivo"
                End If
            Case Is = "LstProveedor"
                CSql = "select * from Proveedores where Nombre='" & Trim(RsCargarListadoProveedores.Fields("Nombre").Value) & "' or RifProveedor ='" & Trim(RsCargarListadoProveedores.Fields("RifProveedor").Value) & "'"
                Set FrmProveedores.BD74 = CrearRS(CSql)
                FrmProveedores.CargaProve
        
        End Select
    Else
        Msg = "EL Proveedor Seleccionado se encuentra Inactivo!!"
        MsgBox Msg, vbCritical + vbOKOnly, "Error"
    End If
    Unload Me
End If
End Sub

Private Sub Form_Load()
Centrar Me
Grid1
Dim RsCargarListadoProveedores As New ADODB.Recordset
CSql = "Select * From Proveedores"
Set RsCargarListadoProveedores = CrearRS(CSql)

Do While Not RsCargarListadoProveedores.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListadoProveedores.Fields("IdProveedor").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListadoProveedores.Fields("Nombre").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListadoProveedores.Fields("RifProveedor").Value
    RsCargarListadoProveedores.MoveNext
Loop

DMGrid1.PaintMGrid

End Sub

Sub Grid1()
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

Private Sub LstProveedores_DblClick()
Dim RsSeleccionarProveedor As New ADODB.Recordset
CSql = "Select * From Proveedores Where IdProveedor='" & LstProveedores.SelectedItem.Text & "'"
Set RsSeleccionarProveedor = CrearRS(CSql)

Select Case Tipo
Case Is = "Ordenes"
    FrmOrdenCompra.TxtCodigoProveedor.Text = RsSeleccionarProveedor.Fields("IdProveedor").Value
    FrmOrdenCompra.TxtDescripcionProveerdor.Text = RsSeleccionarProveedor.Fields("Nombre").Value
    FrmOrdenCompra.TxtRif.Text = RsSeleccionarProveedor.Fields("RifProveedor").Value
Case Is = "Compras"
    FrmCompras.TxtCodigoProveedor.Text = RsSeleccionarProveedor.Fields("IdProveedor").Value
    FrmCompras.TxtDescripcionProveerdor.Text = RsSeleccionarProveedor.Fields("Nombre").Value
    FrmCompras.TxtRif.Text = RsSeleccionarProveedor.Fields("RifProveedor").Value
End Select
RsSeleccionarProveedor.Close
Unload Me
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

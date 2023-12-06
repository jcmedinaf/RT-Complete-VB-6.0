VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoProductosServicios 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Productos o Servicios"
   ClientHeight    =   6585
   ClientLeft      =   3105
   ClientTop       =   2085
   ClientWidth     =   13635
   Icon            =   "Listaproductos.frx":0000
   LinkTopic       =   "Form30"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   13635
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3720
         TabIndex        =   4
         Top             =   5640
         Width           =   9615
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   1320
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   8520
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
            MICON           =   "Listaproductos.frx":1002
            PICN            =   "Listaproductos.frx":101E
            PICH            =   "Listaproductos.frx":11E7
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
         Top             =   5640
         Width           =   3495
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
            ToolTipText     =   "Ingrese la busqueda por Código o Descripción del Producto"
            Top             =   240
            Width           =   1815
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2040
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "Listaproductos.frx":141C
            PICN            =   "Listaproductos.frx":1438
            PICH            =   "Listaproductos.frx":169D
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
         Height          =   5295
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   9340
         Object.Width           =   13185
         Object.Height          =   5265
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoProductosServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarProductos As New ADODB.Recordset
Dim RsCargarConfig As New ADODB.Recordset
Dim P1, P2, P3 As Boolean
Dim ValorIva As Double
Dim RsPegarProd As New ADODB.Recordset
Dim FilaDisp As Boolean
Dim Impuest As Boolean
Dim i As Integer

Private Sub BtnBuscar_Click()

If TxtBuscar.Text <> "" Then
    CSql = "Select * From Productos where IdProducto=" & Val(Trim(TxtBuscar.Text)) & " Or Descripcion like '%" & Trim(TxtBuscar.Text) & "%'"
Else
    CSql = "Select * From Productos"
End If
DMGrid1.Rows = 0
DMGrid1.Clear
Set RsCargarProductos = CrearRS(CSql)
If RsCargarProductos.RecordCount > 0 Then
    Do While Not RsCargarProductos.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarProductos.Fields("IdProducto").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarProductos.Fields("Descripcion").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = 0 'RsCargarProductos.Fields("CedulaP").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsCargarProductos.Fields("CostoActual").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsCargarProductos.Fields("PrecioUnitario1").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 6) = RsCargarProductos.Fields("PrecioUnitario2").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 7) = RsCargarProductos.Fields("PrecioUnitario3").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 8) = RsCargarProductos.Fields("TipoServicio").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 9) = RsCargarProductos.Fields("Ubicacion").Value
        RsCargarProductos.MoveNext
    Loop
End If
DMGrid1.PaintMGrid

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
On Error GoTo WrtError
If Button = vbRightButton Then

    If Val(DMGrid1.Rows) = 0 Then Exit Sub
    
    CSql = "Select * From Productos Where IdProducto ='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsPegarProd = CrearRS(CSql)
    
    Impuest = RsPegarProd.Fields("Impuesto").Value
    If P1 = True Then
        CodProd = RsPegarProd.Fields("IdProducto").Value
        DescPro = RsPegarProd.Fields("Descripcion").Value
        PreProd = RsPegarProd.Fields("PrecioUnitario1").Value
    '    If RsPegarProd.Fields("Impuesto").Value = True Then
    '        IvaProd = PreProd * (ValorIva / 100)
    '    Else
    '        IvaProd = Format(0, "#,##0.00")
    '    End If
    End If
    
    If P2 = True Then
        CodProd = RsPegarProd.Fields("IdProducto").Value
        DescPro = RsPegarProd.Fields("Descripcion").Value
        PreProd = RsPegarProd.Fields("PrecioUnitario2").Value
    '    If RsPegarProd.Fields("Impuesto").Value = True Then
    '        IvaProd = PreProd * (ValorIva / 100)
    '    Else
    '        IvaProd = Format(0, "#,##0.00")
    '    End If
    End If
    
    If P3 = True Then
        CodProd = RsPegarProd.Fields("IdProducto").Value
        DescPro = RsPegarProd.Fields("Descripcion").Value
        PreProd = RsPegarProd.Fields("PrecioUnitario3").Value
    '    If RsPegarProd.Fields("Impuesto").Value = True Then
    '        IvaProd = PreProd * (ValorIva / 100)
    '    Else
    '        IvaProd = Format(0, "#,##0.00")
    '    End If
    End If
    RsPegarProd.Close
    
    Select Case Tipo
    Case Is = "Facturacion"
    
        For i = 1 To FacturacionRT.DMGrid1.Rows
            If Trim(FacturacionRT.DMGrid1.ValorCelda(i, 1)) = "" Then
                FacturacionRT.DMGrid1.ValorCelda(i, 1) = CodProd
                FacturacionRT.DMGrid1.ValorCelda(i, 2) = DescPro
                FacturacionRT.DMGrid1.ValorCelda(i, 4) = PreProd
                FacturacionRT.DMGrid1.ValorCelda(i, 6) = "0"
                Call Llenar_FRM_FACTURACION(i, Impuest)
                FilaDisp = True
                Exit For
            End If
        Next i
        
        If Not FilaDisp Then
            FacturacionRT.DMGrid1.Rows = FacturacionRT.DMGrid1.Rows + 1
            Call FacturacionRT.DMGrid1.PaintMGrid
            FacturacionRT.DMGrid1.ValorCelda(i, 1) = CodProd
            FacturacionRT.DMGrid1.ValorCelda(i, 2) = DescPro
            FacturacionRT.DMGrid1.ValorCelda(i, 4) = PreProd
            FacturacionRT.DMGrid1.ValorCelda(i, 6) = "0"
            Llenar_FRM_FACTURACION FacturacionRT.DMGrid1.Rows, Impuest
        End If
        'FacturacionRT.DMGrid1.ValorCelda(f, 5) = IvaProd
        
        FacturacionRT.DMGrid1.Col = 3
       ' FacturacionRT.DMGrid1.SetFocus
        Call FacturacionRT.DMGrid1.PaintMGrid
        
    Case Is = "Ordenes"
        'FilaDisp = False
        For i = 1 To FrmOrdenCompra.DMGrid1.Rows
            If Trim(FrmOrdenCompra.DMGrid1.ValorCelda(i, 1)) = "" Then
                FrmOrdenCompra.DMGrid1.ValorCelda(i, 1) = CodProd
                FrmOrdenCompra.DMGrid1.ValorCelda(i, 2) = DescPro
                FrmOrdenCompra.DMGrid1.ValorCelda(i, 4) = PreProd
                FrmOrdenCompra.DMGrid1.ValorCelda(i, 6) = "0"
                Call Llenar_FRM_ORDENES(i, Impuest)
                FilaDisp = True
                Exit For
            End If
        Next i
        
        If Not FilaDisp Then
            FrmOrdenCompra.DMGrid1.Rows = FrmOrdenCompra.DMGrid1.Rows + 1
            Call FrmOrdenCompra.DMGrid1.PaintMGrid
            FrmOrdenCompra.DMGrid1.ValorCelda(i, 1) = CodProd
            FrmOrdenCompra.DMGrid1.ValorCelda(i, 2) = DescPro
            FrmOrdenCompra.DMGrid1.ValorCelda(i, 4) = PreProd
            FrmOrdenCompra.DMGrid1.ValorCelda(i, 6) = "0"
            Call Llenar_FRM_ORDENES(FrmOrdenCompra.DMGrid1.Rows, Impuest)
        End If
        
        FrmOrdenCompra.DMGrid1.Col = 3
        Call FrmOrdenCompra.DMGrid1.PaintMGrid
    
    Case Is = "Compras"
     FilaDisp = False
        For i = 1 To FrmCompras.DMGrid1.Rows
            If Trim(FrmCompras.DMGrid1.ValorCelda(i, 1)) = "" Then
                FrmCompras.DMGrid1.ValorCelda(i, 1) = CodProd
                FrmCompras.DMGrid1.ValorCelda(i, 2) = DescPro
                FrmCompras.DMGrid1.ValorCelda(i, 4) = PreProd
                FrmCompras.DMGrid1.ValorCelda(i, 6) = "0"
                Call Llenar_FRM_COMPRAS(i, Impuest)
                FilaDisp = True
                Exit For
            End If
        Next i
        
        If Not FilaDisp Then
            FrmCompras.DMGrid1.Rows = FrmCompras.DMGrid1.Rows + 1
            Call FrmCompras.DMGrid1.PaintMGrid
            FrmCompras.DMGrid1.ValorCelda(i, 1) = CodProd
            FrmCompras.DMGrid1.ValorCelda(i, 2) = DescPro
            FrmCompras.DMGrid1.ValorCelda(i, 4) = PreProd
            FrmCompras.DMGrid1.ValorCelda(i, 6) = "0"
            Call Llenar_FRM_COMPRAS(FrmCompras.DMGrid1.Rows, Impuest)
        End If
    
        FrmCompras.DMGrid1.Col = 3
        Call FrmCompras.DMGrid1.PaintMGrid
    
    
    Case Is = "Insumos"
    
        For i = 1 To FrmSolicitudNecesidades.DMGrid1.Rows
            If Trim(FrmSolicitudNecesidades.DMGrid1.ValorCelda(i, 1)) = "" Then
                FrmSolicitudNecesidades.DMGrid1.ValorCelda(i, 1) = CodProd
                FrmSolicitudNecesidades.DMGrid1.ValorCelda(i, 2) = DescPro
                FrmSolicitudNecesidades.DMGrid1.ValorCelda(i, 3) = 0
                FrmSolicitudNecesidades.DMGrid1.DColumnas(2).Locked = True
                FrmSolicitudNecesidades.DMGrid1.DColumnas(3).Locked = False
                
                 
                Exit For
            End If
        Next i
        
 
        FrmSolicitudNecesidades.DMGrid1.Col = 3
        Call FrmSolicitudNecesidades.DMGrid1.PaintMGrid
     
    Case Is = "Consumo"

     For i = 1 To FrmConsumoMedicamentos.DMGrid1.Rows
            If Trim(FrmConsumoMedicamentos.DMGrid1.ValorCelda(i, 1)) = "" Then
                FrmConsumoMedicamentos.DMGrid1.ValorCelda(i, 1) = CodProd
                FrmConsumoMedicamentos.DMGrid1.ValorCelda(i, 2) = DescPro
                FrmConsumoMedicamentos.DMGrid1.ValorCelda(i, 3) = 0
                FrmConsumoMedicamentos.DMGrid1.DColumnas(2).Locked = True
                FrmConsumoMedicamentos.DMGrid1.DColumnas(3).Locked = False
                                 
                Exit For
            End If
        Next i
        
 
        FrmConsumoMedicamentos.DMGrid1.Col = 3
        Call FrmConsumoMedicamentos.DMGrid1.PaintMGrid
    
    Case Is = "Productos"

        FrmProductos.TxtBuscar.Text = CodProd
        FrmProductos.BtnBuscar_Click
    Case Is = ""
        Exit Sub
    End Select
    Unload Me
End If

Exit Sub
WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub

Private Sub Form_Load()
Centrar Me
CargarConfig
CargarProductos
End Sub

Sub CargarConfig()

CSql = "Select * From Dat_Admin"
Set RsCargarConfig = CrearRS(CSql)

P1 = RsCargarConfig.Fields("PrecioUnitario1").Value
P2 = RsCargarConfig.Fields("PrecioUnitario2").Value
P3 = RsCargarConfig.Fields("PrecioUnitario3").Value
ValorIva = RsCargarConfig.Fields("Iva1").Value
RsCargarConfig.Close
End Sub

Sub CargarProductos()
Centrar Me
IniDMGrid

Dim RsStockCompras As New ADODB.Recordset
Dim RsStockVentas As New ADODB.Recordset
Dim CantidadVentas, CantidadCompras

CSql = "Select * From Productos"
Set RsCargarProductos = CrearRS(CSql)

Do While Not RsCargarProductos.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarProductos.Fields("IdProducto").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarProductos.Fields("descripcion").Value
        
        CSql = "Select Sum(Cantidad) as TCantidad From RenglonCompras where IdProducto='" & RsCargarProductos.Fields("IdProducto").Value & "'"
        Set RsStockCompras = CrearRS(CSql)
        
        If RsStockCompras.RecordCount > 0 Then
            If Not IsNull(RsStockCompras.Fields("TCantidad").Value) Then
                CantidadCompras = RsStockCompras.Fields("TCantidad").Value
            Else
                CantidadCompras = 0
            End If
        End If
        
                
        CSql = "Select Sum(Cantidad) as TCantidad From Reng_Cobrar where Cod_Producto='" & RsCargarProductos.Fields("IdProducto").Value & "'"
        Set RsStockVentas = CrearRS(CSql)
        
        If RsStockVentas.RecordCount > 0 Then
            If Not IsNull(RsStockVentas.Fields("TCantidad").Value) Then
                CantidadVentas = RsStockVentas.Fields("TCantidad").Value
            Else
                CantidadVentas = 0
            End If
        End If
        
        If RsStockCompras.RecordCount > 0 And RsStockVentas.RecordCount > 0 Then
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = CDbl(CantidadCompras) - CDbl(CantidadVentas)
        Else
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = 0
        End If
        
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsCargarProductos.Fields("CostoActual").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsCargarProductos.Fields("PrecioUnitario1").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 6) = RsCargarProductos.Fields("PrecioUnitario2").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 7) = RsCargarProductos.Fields("PrecioUnitario3").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 8) = RsCargarProductos.Fields("TipoServicio").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 9) = RsCargarProductos.Fields("Ubicacion").Value
    RsCargarProductos.MoveNext
Loop
DMGrid1.PaintMGrid
End Sub


Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 9
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1
DMGrid1.DColumnas(6).Alignment = 1
DMGrid1.DColumnas(7).Alignment = 1
DMGrid1.DColumnas(8).Alignment = 0
DMGrid1.DColumnas(9).Alignment = 0

DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(4).Locked = True
DMGrid1.DColumnas(5).Locked = True
DMGrid1.DColumnas(6).Locked = True
DMGrid1.DColumnas(7).Locked = True
DMGrid1.DColumnas(8).Locked = True
DMGrid1.DColumnas(9).Locked = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(6).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(7).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(8).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(9).Width = Val(DMGrid1.Width * 10 / 100)

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Descripción"
DMGrid1.DColumnas(3).Caption = "Cantidad"
DMGrid1.DColumnas(4).Caption = "Costo Actual"
DMGrid1.DColumnas(5).Caption = "Precio Unit. 1"
DMGrid1.DColumnas(6).Caption = "Precio Unit. 2"
DMGrid1.DColumnas(7).Caption = "Precio Unit. 3"
DMGrid1.DColumnas(8).Caption = "Tipo"
DMGrid1.DColumnas(9).Caption = "Ubicación"


End Sub

Sub CargarProductos1()

CSql = "Select * From Productos"
Set RsCargarProductos = CrearRS(CSql)

LstInventario.ListItems.Clear
If Not (RsCargarProductos.EOF) Then
    Do While Not RsCargarProductos.EOF
        With LstInventario
            i = i + 1
            .ListItems.Add , , RsCargarProductos.Fields("IdProducto").Value
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("descripcion").Value
            .ListItems(i).ListSubItems.Add , , 0
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("CostoActual").Value
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("PrecioUnitario1").Value
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("PrecioUnitario2").Value
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("PrecioUnitario3").Value
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("TipoServicio").Value
            .ListItems(i).ListSubItems.Add , , RsCargarProductos.Fields("Ubicacion").Value
        End With
           RsCargarProductos.MoveNext
    Loop
Else
    RsCargarProductos.Close
    Exit Sub
End If
End Sub

'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
Sub Llenar_FRM_FACTURACION(fil As Integer, Impuest As Boolean)
Dim a, b, C, d As Double
Dim s As Integer

s = fil

    If IsNull(FacturacionRT.DMGrid1.ValorCelda(s, 4)) Then
        a = 0
    ElseIf Val(FacturacionRT.DMGrid1.ValorCelda(s, 4)) = 0 Then
        a = 0
    Else
        a = FacturacionRT.DMGrid1.ValorCelda(s, 4)
    End If
    'Call QuitarCaracter(a)
    'a = CArac
    
    If IsNull(FacturacionRT.DMGrid1.ValorCelda(s, 3)) Then
        b = 1
        FacturacionRT.DMGrid1.ValorCelda(s, 3) = 1
    ElseIf Val(FacturacionRT.DMGrid1.ValorCelda(s, 3)) = 0 Then
        b = 1
        FacturacionRT.DMGrid1.ValorCelda(s, 3) = 1
    Else
        b = FacturacionRT.DMGrid1.ValorCelda(s, 3)
    End If
    
    'Call QuitarCaracter(b)
    'b = CArac
    
    'calculo del impuesto
    If Impuest Then
        FacturacionRT.DMGrid1.ValorCelda(s, 5) = a * b * (ValorIva / 100)
        FacturacionRT.DMGrid1.RowBackColor s, RGB(255, 255, 255)
    Else
        FacturacionRT.DMGrid1.ValorCelda(s, 5) = Format(0, "#,##0.00")
        FacturacionRT.DMGrid1.RowBackColor s, RGB(221, 221, 221)
    End If
    
    If IsNull(FacturacionRT.DMGrid1.ValorCelda(s, 5)) Then
        C = Format(0, "#,##0.00")
    ElseIf Val(FacturacionRT.DMGrid1.ValorCelda(s, 5)) = 0 Then
        C = Format(0, "#,##0.00")
    Else
        C = FacturacionRT.DMGrid1.ValorCelda(s, 5)
    End If
    'Call QuitarCaracter(c)
    'c = CArac
    
    
    If IsNull(FacturacionRT.DMGrid1.ValorCelda(s, 6)) Then
        d = Format(0, "#,##0.00")
    ElseIf Val(FacturacionRT.DMGrid1.ValorCelda(s, 6)) = 0 Then
        d = Format(0, "#,##0.00")
    Else
        d = FacturacionRT.DMGrid1.ValorCelda(s, 6)
    End If
    'Call QuitarCaracter(d)
    'd = CArac
    
    If IsNull(a) Then a = 0 ' Precio Unitario
    If IsNull(b) Then b = 0 ' Cantidad
    If IsNull(C) Then C = 0 ' Iva
    If IsNull(d) Then d = 0 ' Descuento
      
    'DMGrid1.ValorCelda(s, 7) = (a * b - d) * (1 + (c / 100))
    FacturacionRT.DMGrid1.ValorCelda(s, 7) = (a * b - d) + C
    FacturacionRT.DMGrid1.PaintMGrid
    FacturacionRT.calcular
End Sub

'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
Sub Llenar_FRM_ORDENES(fil As Integer, Impuest As Boolean)
Dim a, b, C, d As Double
Dim s As Integer

s = fil

    If IsNull(FrmOrdenCompra.DMGrid1.ValorCelda(s, 4)) Then
        a = 0
    ElseIf Val(FrmOrdenCompra.DMGrid1.ValorCelda(s, 4)) = 0 Then
        a = 0
    Else
        a = FrmOrdenCompra.DMGrid1.ValorCelda(s, 4)
    End If
    'Call QuitarCaracter(a)
    'a = CArac
    
    If IsNull(FrmOrdenCompra.DMGrid1.ValorCelda(s, 3)) Then
        b = 1
        FrmOrdenCompra.DMGrid1.ValorCelda(s, 3) = 1
    ElseIf Val(FrmOrdenCompra.DMGrid1.ValorCelda(s, 3)) = 0 Then
        b = 1
        FrmOrdenCompra.DMGrid1.ValorCelda(s, 3) = 1
    Else
        b = FrmOrdenCompra.DMGrid1.ValorCelda(s, 3)
    End If
    
    'Call QuitarCaracter(b)
    'b = CArac
    
    'calculo del impuesto
    If Impuest Then
        FrmOrdenCompra.DMGrid1.ValorCelda(s, 5) = a * b * (ValorIva / 100)
        FrmOrdenCompra.DMGrid1.RowBackColor s, RGB(255, 255, 255)
    Else
        FrmOrdenCompra.DMGrid1.ValorCelda(s, 5) = Format(0, "#,##0.00")
        FrmOrdenCompra.DMGrid1.RowBackColor s, RGB(221, 221, 221)
    End If
    
    If IsNull(FrmOrdenCompra.DMGrid1.ValorCelda(s, 5)) Then
        C = Format(0, "#,##0.00")
    ElseIf Val(FrmOrdenCompra.DMGrid1.ValorCelda(s, 5)) = 0 Then
        C = Format(0, "#,##0.00")
    Else
        C = FrmOrdenCompra.DMGrid1.ValorCelda(s, 5)
    End If
    'Call QuitarCaracter(c)
    'c = CArac
    
    
    If IsNull(FrmOrdenCompra.DMGrid1.ValorCelda(s, 6)) Then
        d = Format(0, "#,##0.00")
    ElseIf Val(FrmOrdenCompra.DMGrid1.ValorCelda(s, 6)) = 0 Then
        d = Format(0, "#,##0.00")
    Else
        d = FrmOrdenCompra.DMGrid1.ValorCelda(s, 6)
    End If
    'Call QuitarCaracter(d)
    'd = CArac
    
    If IsNull(a) Then a = 0 ' Precio Unitario
    If IsNull(b) Then b = 0 ' Cantidad
    If IsNull(C) Then C = 0 ' Iva
    If IsNull(d) Then d = 0 ' Descuento
      
    'DMGrid1.ValorCelda(s, 7) = (a * b - d) * (1 + (c / 100))
    FrmOrdenCompra.DMGrid1.ValorCelda(s, 7) = (a * b - d) + C
    FrmOrdenCompra.DMGrid1.PaintMGrid
    FrmOrdenCompra.calcular
End Sub

'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
Sub Llenar_FRM_COMPRAS(fil As Integer, Impuest As Boolean)
Dim a, b, C, d As Double
Dim s As Integer

s = fil

    If IsNull(FrmCompras.DMGrid1.ValorCelda(s, 4)) Then
        a = 0
    ElseIf Val(FrmCompras.DMGrid1.ValorCelda(s, 4)) = 0 Then
        a = 0
    Else
        a = FrmCompras.DMGrid1.ValorCelda(s, 4)
    End If
    'Call QuitarCaracter(a)
    'a = CArac

    If IsNull(FrmCompras.DMGrid1.ValorCelda(s, 3)) Then
        b = 1
        FrmCompras.DMGrid1.ValorCelda(s, 3) = 1
    ElseIf Val(FrmCompras.DMGrid1.ValorCelda(s, 3)) = 0 Then
        b = 1
        FrmCompras.DMGrid1.ValorCelda(s, 3) = 1
    Else
        b = FrmCompras.DMGrid1.ValorCelda(s, 3)
    End If

    'Call QuitarCaracter(b)
    'b = CArac

    'calculo del impuesto
    If Impuest Then
        FrmCompras.DMGrid1.ValorCelda(s, 5) = a * b * (ValorIva / 100)
        FrmCompras.DMGrid1.RowBackColor s, RGB(255, 255, 255)
    Else
        FrmCompras.DMGrid1.ValorCelda(s, 5) = Format(0, "#,##0.00")
        FrmCompras.DMGrid1.RowBackColor s, RGB(221, 221, 221)
    End If

    If IsNull(FrmCompras.DMGrid1.ValorCelda(s, 5)) Then
        C = Format(0, "#,##0.00")
    ElseIf Val(FrmCompras.DMGrid1.ValorCelda(s, 5)) = 0 Then
        C = Format(0, "#,##0.00")
    Else
        C = FrmCompras.DMGrid1.ValorCelda(s, 5)
    End If
    'Call QuitarCaracter(c)
    'c = CArac


    If IsNull(FrmCompras.DMGrid1.ValorCelda(s, 6)) Then
        d = Format(0, "#,##0.00")
    ElseIf Val(FrmCompras.DMGrid1.ValorCelda(s, 6)) = 0 Then
        d = Format(0, "#,##0.00")
    Else
        d = FrmCompras.DMGrid1.ValorCelda(s, 6)
    End If
    'Call QuitarCaracter(d)
    'd = CArac

    If IsNull(a) Then a = 0 ' Precio Unitario
    If IsNull(b) Then b = 0 ' Cantidad
    If IsNull(C) Then C = 0 ' Iva
    If IsNull(d) Then d = 0 ' Descuento

    'DMGrid1.ValorCelda(s, 7) = (a * b - d) * (1 + (c / 100))
    FrmCompras.DMGrid1.ValorCelda(s, 7) = (a * b - d) + C
    FrmCompras.DMGrid1.PaintMGrid
    FrmCompras.calcular
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

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmImportarOrdenesCompra 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Ordenes de Compra"
   ClientHeight    =   5505
   ClientLeft      =   2490
   ClientTop       =   3480
   ClientWidth     =   12465
   Icon            =   "FrmImportarOrdenesCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   12465
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   4440
         Width           =   12015
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   10920
            TabIndex        =   3
            ToolTipText     =   "Cerrar "
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
            MICON           =   "FrmImportarOrdenesCompra.frx":1002
            PICN            =   "FrmImportarOrdenesCompra.frx":101E
            PICH            =   "FrmImportarOrdenesCompra.frx":11E7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnDesHacer 
            Height          =   495
            Left            =   9720
            TabIndex        =   4
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Deshacer"
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
            MICON           =   "FrmImportarOrdenesCompra.frx":141C
            PICN            =   "FrmImportarOrdenesCompra.frx":1438
            PICH            =   "FrmImportarOrdenesCompra.frx":171A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label LblNoOrdenes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Ordenes de Compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2085
         End
      End
      Begin MSComctlLib.ListView LstImportar 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No Orden"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Proveedor"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "SubTotal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Impuesto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Total General"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmImportarOrdenesCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
CargarOrdenes
End Sub

Private Sub Form_Load()
Centrar Me
CargarOrdenes
End Sub

Sub CargarOrdenes()
Dim RsCargarOrdenes As New ADODB.Recordset
Dim RsBuscarProveedor As New ADODB.Recordset

CSql = "Select * From Ordenes Order by NumeroOrden asc"
Set RsCargarOrdenes = CrearRS(CSql)

LstImportar.ListItems.Clear
Do While Not RsCargarOrdenes.EOF
    CSql = "Select * From Proveedores where IdProveedor='" & RsCargarOrdenes.Fields("IdProveedor").Value & "'"
    Set RsBuscarProveedor = CrearRS(CSql)
    With LstImportar
        i = i + 1
        .ListItems.Add , , RsCargarOrdenes.Fields("NumeroOrden").Value
        .ListItems(i).ListSubItems.Add , , RsBuscarProveedor.Fields("Nombre").Value
        .ListItems(i).ListSubItems.Add , , RsCargarOrdenes.Fields("FechaEmision").Value
        .ListItems(i).ListSubItems.Add , , Format(RsCargarOrdenes.Fields("SubTotal").Value, "#,##0.00")
        .ListItems(i).ListSubItems.Add , , Format(RsCargarOrdenes.Fields("Impuesto").Value, "#,##0.00")
        .ListItems(i).ListSubItems.Add , , Format(RsCargarOrdenes.Fields("TotalGeneral").Value, "#,##0.00")
    End With
    RsCargarOrdenes.MoveNext
Loop
LblNoOrdenes.Caption = LstImportar.ListItems.Count
End Sub

Private Sub LstImportar_DblClick()
Dim NumeroOrden As String
NumeroOrden = LstImportar.SelectedItem.Text
opcion = 0
CSql = "Select * From Ordenes Where NumeroOrden = '" & Trim(NumeroOrden) & "'"
Dim RsBuscarOrdenes As New ADODB.Recordset
Set RsBuscarOrdenes = CrearRS(CSql)
If RsBuscarOrdenes.EOF = False Or RsBuscarOrdenes.BOF = False Then
    FrmCompras.TxtNoOrdenCompra.Text = RsBuscarOrdenes.Fields("NumeroOrden").Value
    FrmCompras.CboCondicionPago.Text = RsBuscarOrdenes.Fields("CondicionPago").Value
    FrmCompras.DTPickerFechaEmision.Value = RsBuscarOrdenes.Fields("FechaEmision").Value
    FrmCompras.DTPickerFechaRecepcion.Value = RsBuscarOrdenes.Fields("FechaRecepcion").Value
    FrmCompras.TxtSubTotal.Text = Format(RsBuscarOrdenes.Fields("SubTotal").Value, "#,##0.00")
    FrmCompras.TxtImpuesto.Text = Format(RsBuscarOrdenes.Fields("Impuesto").Value, "#,##0.00")
    FrmCompras.TxtTotalGeneral.Text = Format(RsBuscarOrdenes.Fields("TotalGeneral").Value, "#,##0.00")
    'If RsBuscarOrdenes.Fields("OrdenProcesada").Value = True Then Check1.Value = 1 Else Check1.Value = 0
    CodProveedor = RsBuscarOrdenes.Fields("IdProveedor").Value
End If

CSql = "Select * From Proveedores Where IdProveedor = '" & Trim(CodProveedor) & "'"
Dim RsBuscarProveedor As New ADODB.Recordset
Set RsBuscarProveedor = CrearRS(CSql)

FrmCompras.TxtCodigoProveedor.Text = RsBuscarProveedor.Fields("IdProveedor").Value
FrmCompras.TxtDescripcionProveerdor.Text = RsBuscarProveedor.Fields("Nombre").Value
FrmCompras.TxtRif.Text = RsBuscarProveedor.Fields("RifProveedor").Value

CSql = "Select * From RenglonOrdenes Where NumeroOrden = '" & Trim(NumeroOrden) & "'"
Dim RsBuscarRenglonOrden As New ADODB.Recordset
Set RsBuscarRenglonOrden = CrearRS(CSql)

i = 1
If Not (RsBuscarRenglonOrden.EOF) Then
    RsBuscarRenglonOrden.MoveFirst
    Dim RsProductos As New ADODB.Recordset
    Do While Not RsBuscarRenglonOrden.EOF
        FrmCompras.DMGrid1.Rows = i
        CSql = "Select * From Productos Where IdProducto ='" & RsBuscarRenglonOrden.Fields("IdProducto").Value & "'"
        Set RsProductos = CrearRS(CSql)
        
        If RsProductos.RecordCount <> 0 Then
            FrmCompras.DMGrid1.ValorCelda(i, 1) = RsBuscarRenglonOrden.Fields("IdProducto").Value
            FrmCompras.DMGrid1.ValorCelda(i, 2) = RsProductos.Fields("Descripcion").Value
            FrmCompras.DMGrid1.ValorCelda(i, 3) = RsBuscarRenglonOrden.Fields("Cantidad").Value
            FrmCompras.DMGrid1.ValorCelda(i, 4) = RsBuscarRenglonOrden.Fields("precio").Value
            FrmCompras.DMGrid1.ValorCelda(i, 5) = RsBuscarRenglonOrden.Fields("impuesto").Value
            FrmCompras.DMGrid1.ValorCelda(i, 6) = RsBuscarRenglonOrden.Fields("descuento").Value
            FrmCompras.DMGrid1.ValorCelda(i, 7) = RsBuscarRenglonOrden.Fields("SubTotal").Value '(RsBuscarRenglonOrden.Fields("precio").Value * RsBuscarRenglonOrden.Fields("cantidad").Value - RsBuscarRenglonOrden.Fields("descuento").Value) + ((RsBuscarRenglonOrden.Fields("impuesto").Value * (RsBuscarRenglonOrden.Fields("precio").Value * RsBuscarRenglonOrden.Fields("cantidad").Value - RsBuscarRenglonOrden.Fields("descuento").Value) / 100))
            If RsProductos.Fields("Impuesto") Then
                        FrmCompras.DMGrid1.RowBackColor i, RGB(255, 255, 255)
                    Else
                        FrmCompras.DMGrid1.RowBackColor i, RGB(221, 221, 221)
                    End If
            RsProductos.Close
            i = i + 1
        End If
        
        RsBuscarRenglonOrden.MoveNext
    Loop
       FrmCompras.TxtCantidadRenglon.Text = FrmCompras.DMGrid1.Rows
Else
    FrmCompras.DMGrid1.Clear
    Call FrmCompras.DMGrid1.PaintMGrid
    RsBuscarRenglonOrden.Close
    RsBuscarFactura.Close
    Exit Sub
End If
FrmCompras.DMGrid1.PaintMGrid
RsBuscarOrdenes.Close
RsBuscarProveedor.Close
FrmCompras.calcular
Unload Me
End Sub

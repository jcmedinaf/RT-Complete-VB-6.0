VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmProductosQuimio 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos de Quimioterapias"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12165
   Icon            =   "FrmProductosQuimio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductosQuimio.frx":1002
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox TxtProductosQuimio 
      Height          =   8295
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   14631
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmProductosQuimio.frx":128C
   End
   Begin MSComctlLib.TreeView Arbol 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   14631
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin ChamaleonButton.ChameleonBtn BtnCerrar 
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      ToolTipText     =   "Cerrar"
      Top             =   8520
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
      MICON           =   "FrmProductosQuimio.frx":130E
      PICN            =   "FrmProductosQuimio.frx":132A
      PICH            =   "FrmProductosQuimio.frx":14F3
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
Attribute VB_Name = "FrmProductosQuimio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsProductosQuimio As New ADODB.Recordset

Private Sub Arbol_NodeClick(ByVal Node As MSComctlLib.Node)
 'Verifica que el nodo en el que se hizo clic, no sea el nodo Root
    If Node <> "Raiz" Then
       ' Le pasa el nombre de la tabla como parámetro para cargar _
        los registros en el DataGrid y la cadena de conexión
       
       
       TxtProductosQuimio.TextRTF = Node.Text
       
       
       
       
       
       
       
    End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()

'CSql = "Select * From ProductosQuimio"
'Set RsProductosQuimio = CrearRS(CSql)
'Turnos '' abro la tabla Turnos
Arbol.Nodes.Clear
'With RsProductosQuimio
    '.Requery
    'If .BOF Or .EOF Then '' en caso de que turno este vacia
        'Exit Sub
    'Else
        Set Nodo = Arbol.Nodes.Add(, , "Raiz", "Farmacos", 1)
        Nodo.Expanded = True
        'With RsProductosQuimio
            '.Requery
            '.MoveFirst
            'Dim NombreTurno
            'For X = 1 To .RecordCount
                'NombreTurno = !turno
                ''' agregar turno
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "T" & X, "Producto Quimioterapia 1", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "e" & X, "Producto Quimioterapia 2", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "f" & X, "Producto Quimioterapia 3", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "g" & X, "Producto Quimioterapia 4", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "h" & X, "Producto Quimioterapia 5", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "i" & X, "Producto Quimioterapia 6", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "j" & X, "Producto Quimioterapia 7", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "k" & X, "Producto Quimioterapia 8", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "l" & X, "Producto Quimioterapia 9", 1)
                Set Nodo = Arbol.Nodes.Add("Raiz", tvwChild, "m" & X, "Producto Quimioterapia 10", 1)
                Nodo.Expanded = True
                'If X = .RecordCount Then Else .MoveNext
            'Next
        'End With
    'End If
'End With
End Sub
'Option Explicit
'
'' Referencias : _
'  1 - Microsoft Activex Data objects _
'  2 - Un control Treeview _
'  3 - Un ImageList con  Tres imagenes ( Key : Normal, seleccionada, root )) _
'  4 - Un DataGrid _
'  5 - Un CommandButton _
'  6 - Indicar la cadena de conexión a utilizar
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Cadena de conexión a usar ( usa la base de datos Nwind .mdb de visual basic )
'    Private Const Cadena_Conexion As String = "Provider=Microsoft.Jet.OLEDB.4.0" & _
'                                              ";Data Source=C:\Archivos de prog" & _
'                                              "rama\Microsoft Visual Studio\VB9" & _
'                                              "8\NWIND.MDB;Persist Security Info=False"
    
    
    
'' Subrutina que muestra la tabla en el DataGrid al hacer clic en un nodo del Treeview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub Mostrar_Tabla(Tabla As String, ConnectionString As String)
'
'    Dim rst As New ADODB.Recordset
'    Dim cn As New ADODB.Connection
'
'
'        ' Establece el cursor y abre la base de datos
'        cn.CursorLocation = adUseClient
'        cn.Open ConnectionString
'
'        ' Ejecuta el sql para llenar el recordset
'        rst.Open "[" & Tabla & "]", cn, adOpenStatic
'
'        ' enlaza el Datagrid con el Recordset
'        Set DataGrid1.DataSource = rst
'
'        ' elimina las referencias
'        Set rst = Nothing
'        Set cn = Nothing
'End Sub


'Private Sub Cargar_Tablas(ConnectionString As String)
'
'    Dim cn As New ADODB.Connection
'    Dim rst As New ADODB.Recordset
'    Dim Nodo As Node
'    Dim ARRAY_TABLAS() As String
'    Dim i As Integer
'
'        Screen.MousePointer = vbHourglass
'
'        ' elimina todos los nodos
'        TreeView1.Nodes.Clear
'
'        ' añade el nodo principal
'        Set Nodo = TreeView1.Nodes.Add(, tvwFirst, _
'                                            "Root", "Tablas", _
'                                            ImageList1.ListImages("root").Key)
'
'        ' opcional ( expande el nodo Root del treeview )
'        Nodo.Expanded = True
'
'        If rst.State <> adStateOpen Then
'        ' abre la conexión
'           cn.Open ConnectionString
'
'        ' Recupera las tablas de la base de datos mediante OpenSchema
'            Set rst = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
'        ' Se posiciona en la primer tabla
'            rst.MoveFirst
'        End If
'            ' REcorre las tablas para añadirlas en un array
'            Do Until rst.EOF
'                ReDim Preserve ARRAY_TABLAS(i)
'                ' Añade la tabla al vector
'                ARRAY_TABLAS(i) = rst!Table_Name
'                ' siguiente
'                rst.MoveNext
'                i = i + 1
'            Loop
'        ' recorre el array con las tablas
'        For i = 0 To UBound(ARRAY_TABLAS)
'
'            ' crea el nodo correspondiente a esta tabla
'            Set Nodo = TreeView1.Nodes.Add("Root", _
'                                             tvwChild, , _
'                                             ARRAY_TABLAS(i), _
'                                             ImageList1.ListImages("normal").Key, _
'                                             ImageList1.ListImages("seleccionada").Key)
'    Next
'
'    Screen.MousePointer = vbDefault
'
'End Sub
'
'' Botón que carga todas las tablas
''''''''''''''''''''''''''''''''''''''''''
'Private Sub Command1_Click()
'    Call Cargar_Tablas(Cadena_Conexion)
'End Sub
'
'Private Sub Form_Load()
'    Command1.Caption = " Cargar tablas "
'End Sub
'
'Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'    ' Verifica que el nodo en el que se hizo clic, no sea el nodo Root
'    If Node <> "Tablas" Then
'       ' Le pasa el nombre de la tabla como parámetro para cargar _
'        los registros en el DataGrid y la cadena de conexión
'       Call Mostrar_Tabla(Node.Text, Cadena_Conexion)
'    End If
'End Sub


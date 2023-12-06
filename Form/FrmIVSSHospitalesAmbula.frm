VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIVSSHospitalesAmbula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospitales y Ambulatorios"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14700
   Icon            =   "FrmIVSSHospitalesAmbula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   14700
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboCategoria 
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   3840
      Width           =   3015
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   4895
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hospitales y/o Ambulatorios"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dirección"
         Object.Width           =   15875
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin VB.CommandButton BtnCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   12840
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BtnBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   7800
         TabIndex        =   5
         Top             =   210
         Width           =   1455
      End
      Begin VB.ComboBox CboMunicipio 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox CboEstados 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Left            =   4080
         TabIndex        =   3
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estados:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   8281
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descripción"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Modelo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Marca"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "S/N"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Año"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "IVSS I.D."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Bitácora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Ciudad"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Especialidad:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3900
      Width           =   945
   End
   Begin VB.Label LblNumeroEquipo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Equipos: 0"
      Height          =   195
      Left            =   12960
      TabIndex        =   10
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label LblTotalCentros 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Centros: 0"
      Height          =   195
      Left            =   13080
      TabIndex        =   9
      Top             =   3840
      Width           =   1350
   End
End
Attribute VB_Name = "FrmIVSSHospitalesAmbula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBuscar_Click()
ConectarIVSSHosting
Dim RsListado As New ADODB.Recordset

CSql = "Select * From Hospitales Where IdCiudad ='" & CboMunicipio.ItemData(CboMunicipio.ListIndex) & "'Order By Hospital"
RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic

ListView1.ListItems.Clear
i = 0
Do While Not RsListado.EOF
    With ListView1
        i = i + 1
        .ListItems.Add , , RsListado.Fields("Hospital").Value
        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Direccion").Value
        .ListItems(i).ListSubItems.Add , , RsListado.Fields("IdHospital").Value
    End With
    RsListado.MoveNext
Loop
LblTotalCentros.Caption = "Total de Centros: " & i
WebCnn.Close
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub CboCategoria_Click()
Dim RsListado As New ADODB.Recordset
If CboCategoria.ListCount > 0 Then
        
    If CboCategoria.ItemData(CboCategoria.ListIndex) <> 0 Then
        ConectarIVSSHosting
        
        'CSql = "Select * From Equipo Where Especialidades ='" & Trim(CboCategoria.Text) & "' And IdHospital ='" & ListView1.SelectedItem.ListSubItems(2).Text & "'"
        
        CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1, subcategoria.idcategoria FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria WHERE (subcategoria.idcategoria ='" & CboCategoria.ItemData(CboCategoria.ListIndex) & "' And equipo.IdHospital ='" & ListView1.SelectedItem.ListSubItems(2).Text & "')"
        RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
        
        ListView2.ListItems.Clear
        i = 0
        Do While Not RsListado.EOF
            With ListView2
                i = i + 1
                .ListItems.Add , , RsListado.Fields("Descripcion").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Modelo").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Marca").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Numero_Serial").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("ano").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Reg_Ivss").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Bitacora").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Estado").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Ciudad").Value
            End With
            RsListado.MoveNext
        Loop
        LblNumeroEquipo.Caption = "Número de Equipos: " & i
        WebCnn.Close
    Else
        ConectarIVSSHosting
            
        'CSql = "Select * From Equipo Where Especialidades ='" & Trim(CboCategoria.Text) & "' And IdHospital ='" & ListView1.SelectedItem.ListSubItems(2).Text & "'"
        
        CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1, subcategoria.idcategoria FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria WHERE (equipo.IdHospital ='" & ListView1.SelectedItem.ListSubItems(2).Text & "')"
        RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
        
        ListView2.ListItems.Clear
        i = 0
        Do While Not RsListado.EOF
            With ListView2
                i = i + 1
                .ListItems.Add , , RsListado.Fields("Descripcion").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Modelo").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Marca").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Numero_Serial").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("ano").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Reg_Ivss").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Bitacora").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Estado").Value
                .ListItems(i).ListSubItems.Add , , RsListado.Fields("Ciudad").Value
            End With
            RsListado.MoveNext
        Loop
        LblNumeroEquipo.Caption = "Número de Equipos: " & i
        WebCnn.Close
    End If
End If
End Sub

Private Sub CboEstados_Click()
ConectarIVSSHosting
Dim RsCiudad As New ADODB.Recordset

CSql = "Select * From Ciudad Where IdEstado ='" & CboEstados.ItemData(CboEstados.ListIndex) & "' Order By Ciudad"
RsCiudad.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic


Do While Not RsCiudad.EOF
    With CboMunicipio
        .AddItem RsCiudad.Fields("Ciudad").Value
        .ItemData(.NewIndex) = RsCiudad.Fields("IdCiudad").Value
    End With
    RsCiudad.MoveNext
Loop
WebCnn.Close
End Sub

Private Sub Form_Load()
Pn = 0
ConectarIVSSHosting
Dim RsEstados As New ADODB.Recordset

CSql = "Select * From Estado Order By Estado"
RsEstados.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic


Do While Not RsEstados.EOF
    With CboEstados
        .AddItem RsEstados.Fields("Estado").Value
        .ItemData(.NewIndex) = RsEstados.Fields("IdEstado").Value
    End With
    RsEstados.MoveNext
Loop
WebCnn.Close
End Sub

Private Sub ListView1_Click()
Dim RsCategoria As New ADODB.Recordset
Dim RsListado As New ADODB.Recordset

If ListView1.ListItems.Count > 0 Then

    ConectarIVSSHosting
    CSql = "Select * From Categorias"
    RsCategoria.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
    CboCategoria.Clear
    CboCategoria.Text = "Todos"
    CboCategoria.AddItem "Todos"
    CboCategoria.ItemData(CboCategoria.NewIndex) = 0
    
    Do While Not RsCategoria.EOF
        With CboCategoria
            .AddItem RsCategoria.Fields("Categoria").Value
            .ItemData(.NewIndex) = RsCategoria.Fields("IdCategoria").Value
        End With
        RsCategoria.MoveNext
    Loop
    WebCnn.Close


    ConectarIVSSHosting
    
    
    'CSql = "Select * From Equipo Where IdHospital ='" & ListView1.SelectedItem.ListSubItems(2).Text & "'"
    CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1 FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE hospitales.IdHospital ='" & ListView1.SelectedItem.ListSubItems(2).Text & "'"
    RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
    ListView2.ListItems.Clear
    i = 0
    Do While Not RsListado.EOF
        With ListView2
            i = i + 1
            .ListItems.Add , , RsListado.Fields("Descripcion").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Modelo").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Marca").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Numero_Serial").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("ano").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Reg_Ivss").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Bitacora").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Estado").Value
            .ListItems(i).ListSubItems.Add , , RsListado.Fields("Ciudad").Value
        End With
        RsListado.MoveNext
    Loop
    LblNumeroEquipo.Caption = "Número de Equipos: " & i
    WebCnn.Close
End If

End Sub

Private Sub ListView2_DblClick()
If ListView2.ListItems.Count = 0 Then Exit Sub
Bitacora = ListView2.SelectedItem.ListSubItems(6).Text
FrmIVSSDetalle.Show vbModal, FrmPrincipal
End Sub

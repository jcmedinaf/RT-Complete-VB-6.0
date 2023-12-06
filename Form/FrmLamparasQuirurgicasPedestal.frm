VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLamparasQuirurgicasPedestal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lámparas Quirúrgicas  Pedestal"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14745
   Icon            =   "FrmLamparasQuirurgicasPedestal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   8760
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13996
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
   Begin VB.Label LblNumeroEquipo 
      AutoSize        =   -1  'True
      Caption         =   "Número de Equipos: 0"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmLamparasQuirurgicasPedestal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim RsListado As New ADODB.Recordset

'If ListView1.ListItems.Count > 0 Then

    ConectarIVSSHosting
    
    Pn = 9
    CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1 FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE (subcategoria.idsubcategoria =" & Pn & ")"
    RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
    ListView1.ListItems.Clear
    i = 0
    Do While Not RsListado.EOF
        With ListView1
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
'End If




End Sub


Private Sub ListView1_DblClick()
If ListView1.ListItems.Count = 0 Then Exit Sub
Bitacora = ListView1.SelectedItem.ListSubItems(6).Text
FrmIVSSDetalle.Show vbModal, FrmPrincipal
End Sub

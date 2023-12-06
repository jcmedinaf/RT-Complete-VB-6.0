VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIVSSBusquedaEquipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Equipos"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14790
   Icon            =   "FrmIVSSBusquedaEquipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   14535
      Begin VB.CommandButton BtnBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   8640
         TabIndex        =   20
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtSerial 
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox TxtBitacora 
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox TxtIVSSID 
         Height          =   375
         Left            =   9360
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox TxtEstado 
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox TxtCiudad 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox TxtModelo 
         Height          =   375
         Left            =   9360
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TxtMarca 
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   270
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "S/N:"
         Height          =   195
         Left            =   4680
         TabIndex        =   19
         Top             =   1290
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bitácora:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "IVSS I.D:"
         Height          =   195
         Left            =   8640
         TabIndex        =   15
         Top             =   810
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
         Height          =   195
         Left            =   8640
         TabIndex        =   9
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   4680
         TabIndex        =   7
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   8280
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   10186
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
      Width           =   1695
   End
End
Attribute VB_Name = "FrmIVSSBusquedaEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBuscar_Click()
Dim RsListado As New ADODB.Recordset

Dim wer As String
        wer = ""
        If TxtSerial.Text = "" And TxtDescripcion.Text = "" And TxtMarca.Text = "" And TxtModelo.Text = "" And TxtEstado.Text = "" And TxtCiudad.Text = "" And TxtIVSSID.Text = "" And TxtBitacora.Text = "" Then
            Msg = "Por favor ingrese descripción, marca o modelo para realizar una búsqueda"
            MsgBox Msg, vbCritical + vbOKOnly, "Error"
        Else

            If TxtDescripcion.Text <> "" Then
                wer = wer & " descripcion like '%" & TxtDescripcion.Text & "%'"
            End If

            If TxtMarca.Text <> "" Then
                If wer = "" Then
                    wer = wer & " equipo.marca like '%" & TxtMarca.Text & "%'"
                Else
                    wer = wer & " and equipo.marca like '%" & TxtMarca.Text & "%'"
                End If
            End If

            If TxtModelo.Text <> "" Then
                If wer = "" Then
                    wer = wer & " equipo.modelo like '%" & TxtModelo & "%'"
                Else
                    wer = wer & " and equipo.modelo like '%" & TxtModelo.Text & "%'"
                End If
            End If

            If TxtEstado.Text <> "" Then
                If wer = "" Then
                    wer = wer & " estado.estado like '%" & TxtEstado.Text & "%'"
                Else
                    wer = wer & " and estado.estado like '%" & TxtEstado.Text & "%'"
                End If
            End If

            If TxtCiudad.Text <> "" Then
                If wer = "" Then
                    wer = wer & " ciudad.ciudad like '%" & TxtCiudad.Text & "%'"
                Else
                    wer = wer & " and ciudad.ciudad like '%" & TxtCiudad.Text & "%'"
                End If
            End If

            If TxtIVSSID.Text <> "" Then
                If wer = "" Then
                    wer = wer & " equipo.reg_ivss like '%" & TxtIVSSID.Text & "%'"
                Else
                    wer = wer & " and equipo.reg_ivss like '%" & TxtIVSSID.Text & "%'"
                End If
            End If

            If TxtBitacora.Text <> "" Then
                If wer = "" Then
                    wer = wer & " equipo.bitacora like '%" & TxtBitacora.Text & "%'"
                Else
                    wer = wer & " and equipo.bitacora like '%" & TxtBitacora.Text & "%'"
                End If
            End If

            If TxtSerial.Text <> "" Then
                If wer = "" Then
                    wer = wer & " equipo.numero_serial like '%" & TxtSerial.Text & "%'"
                Else
                    wer = wer & " and equipo.numero_serial like '%" & TxtSerial.Text & "%'"
                End If
            End If
End If
ConectarIVSSHosting

'Pn = 3
'CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1 FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE (subcategoria.idsubcategoria =" & Pn & ")"
CSql = "SELECT equipo.ano,estado.estado as Estado,ciudad.ciudad, equipo.bitacora, hospitales.idhospital, hospitales.hospital, hospitales.direccion, hospitales.telefonos, equipo.descripcion, equipo.marca, equipo.modelo, equipo.idequipo, equipo.numero_serial, equipo.reg_ivss FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE " & wer

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
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Dim RsListado As New ADODB.Recordset
'ConectarIVSSHosting
'
''Pn = 3
''CSql = "SELECT equipo.marca, equipo.modelo, equipo.descripcion, equipo.ano, equipo.numero_serial, equipo.reg_ivss, estado.estado, ciudad.ciudad, equipo.bitacora, equipo.idequipo, estado.estado AS Expr1 FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE (subcategoria.idsubcategoria =" & Pn & ")"
'CSql = "SELECT equipo.ano,estado.estado as Estado,ciudad.ciudad, equipo.bitacora, hospitales.idhospital, hospitales.hospital, hospitales.direccion, hospitales.telefonos, equipo.descripcion, equipo.marca, equipo.modelo, equipo.idequipo, equipo.numero_serial, equipo.reg_ivss FROM ciudad INNER JOIN estado ON ciudad.idestado = estado.idestado INNER JOIN hospitales ON ciudad.idciudad = hospitales.idciudad INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria" ' WHERE (Descripcion like '%" & Trim(TxtDescripcion.Text) & "%' or marca like '%" & Trim(TxtMarca.Text) & "%' or Modelo like '%" & Trim(TxtModelo.Text) & "%')"
'RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
'
'ListView1.ListItems.Clear
'i = 0
'Do While Not RsListado.EOF
'    With ListView1
'        i = i + 1
'        .ListItems.Add , , RsListado.Fields("Descripcion").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Modelo").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Marca").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Numero_Serial").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("ano").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Reg_Ivss").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Bitacora").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Estado").Value
'        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Ciudad").Value
'    End With
'    RsListado.MoveNext
'Loop
'LblNumeroEquipo.Caption = "Número de Equipos: " & i
'WebCnn.Close
End Sub

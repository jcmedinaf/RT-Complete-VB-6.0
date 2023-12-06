VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIVSSBusquedaHospitales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda Hospitales"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14865
   Icon            =   "FrmIVSSBusquedaHospitales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   8280
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   14535
      Begin VB.TextBox TxtNombre 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   270
         Width           =   2775
      End
      Begin VB.TextBox TxtCiudad 
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TxtEstado 
         Height          =   375
         Left            =   8160
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton BtnBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   11040
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   3840
         TabIndex        =   6
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   7440
         TabIndex        =   5
         Top             =   330
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6735
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   11880
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
         Text            =   "Descripción"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Direccion"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Telefono"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label LblNumeroEquipo 
      AutoSize        =   -1  'True
      Caption         =   "Número de Equipos: 0"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmIVSSBusquedaHospitales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBuscar_Click()
Dim RsListado As New ADODB.Recordset

Dim wer As String
        wer = ""
        If TxtNombre.Text = "" And TxtCiudad.Text = "" And TxtEstado.Text = "" Then
            Msg = "Por favor ingrese descripción, Nombre del hospital o Ciudad o Estado para realizar una búsqueda"
            MsgBox Msg, vbCritical + vbOKOnly, "Error"
            Exit Sub
        Else

            If TxtNombre.Text <> "" Then
                wer = wer & " hospital like '%" & TxtNombre.Text & "%'"
            End If

            If TxtCiudad.Text <> "" Then
                If wer = "" Then
                    wer = wer & " ciudad.ciudad like '%" & TxtCiudad.Text & "%'"
                Else
                    wer = wer & " and ciudad.ciudad like '%" & TxtCiudad.Text & "%'"
                End If
            End If

            If TxtEstado.Text <> "" Then
                If wer = "" Then
                    wer = wer & " estado.estado like '%" & TxtEstado & "%'"
                Else
                    wer = wer & " and estado.estado like '%" & TxtEstado.Text & "%'"
                End If
            End If

           
End If
ConectarIVSSHosting

CSql = "SELECT hospitales.idhospital, hospitales.hospital, hospitales.tipo, hospitales.direccion, hospitales.telefonos, ciudad.idciudad," & _
           " estado.estado, ciudad.ciudad FROM hospitales INNER JOIN ciudad ON hospitales.idciudad = ciudad.idciudad INNER JOIN estado ON ciudad.idestado " & _
           "= estado.idestado WHERE" & wer & " And hospitales.tipo='1'"

RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic

ListView1.ListItems.Clear
i = 0
Do While Not RsListado.EOF
    With ListView1
        i = i + 1
        .ListItems.Add , , RsListado.Fields("Hospital").Value
        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Direccion").Value
        .ListItems(i).ListSubItems.Add , , RsListado.Fields("Telefonos").Value

    End With
    RsListado.MoveNext
Loop
LblNumeroEquipo.Caption = "Número de Equipos: " & i
WebCnn.Close
End Sub




Private Sub BtnCerrar_Click()
Unload Me
End Sub



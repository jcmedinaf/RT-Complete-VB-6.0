VERSION 5.00
Begin VB.Form FrmIVSSDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   Icon            =   "FrmIVSSDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDetalle 
      Height          =   5055
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton BtnCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox TxtUbicacion 
      Height          =   1695
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox TxtBitacora 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox TxtIVSS 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox TxtSerial 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox TxtAno 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox TxtModelo 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox TxtMarca 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Ubicacion:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bitacora:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3330
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "IVSS I.D:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2850
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "S/I:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2370
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Año:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1890
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1410
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   930
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   450
      Width           =   885
   End
End
Attribute VB_Name = "FrmIVSSDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCerrar_Click()
Unload Me
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim RsListado As New ADODB.Recordset
 
If Pn = 0 Then

    ConectarIVSSHosting
    CSql = "SELECT equipo.ano, equipo.bitacora, hospitales.idhospital, hospitales.hospital, hospitales.direccion, hospitales.telefonos, equipo.descripcion, equipo.marca, equipo.modelo, equipo.idequipo, equipo.numero_serial, equipo.reg_ivss FROM hospitales INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital WHERE (equipo.Bitacora='" & Bitacora & "')"
    RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
    IdEquipo = RsListado.Fields("IdEquipo").Value
    TxtDescripcion.Text = RsListado.Fields("Descripcion").Value
    TxtMarca.Text = RsListado.Fields("marca").Value
    TxtModelo.Text = RsListado.Fields("modelo").Value
    TxtAno.Text = RsListado.Fields("ano").Value
    TxtSerial.Text = RsListado.Fields("numero_serial").Value
    TxtIVSS.Text = RsListado.Fields("reg_ivss").Value
    TxtBitacora.Text = RsListado.Fields("Bitacora").Value
    TxtUbicacion.Text = RsListado.Fields("Direccion").Value
    WebCnn.Close
    
    ConectarIVSSHosting
    
    CSql = "SELECT [idmen], [mens], [fecha] FROM [bitacora] WHERE ([idequipo] =" & IdEquipo & ")"
    RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
'    If IsNull(RsListado.Fields("Fecha").Value) Then
'        TxtDetalle.Text = ""
'        WebCnn.Close
'        Exit Sub
'    End If
    
    
    If Not IsNull(RsListado.Fields("Fecha").Value) Then
        TxtDetalle.Text = Trim(TxtDetalle.Text) & RsListado.Fields("Fecha").Value & Chr(13) & RsListado.Fields("Mens").Value
    End If
    
    WebCnn.Close

End If

If Pn <> 0 Then

    ConectarIVSSHosting
    
    CSql = "SELECT equipo.ano, equipo.bitacora, hospitales.idhospital, hospitales.hospital, hospitales.direccion, hospitales.telefonos, equipo.descripcion, equipo.marca, equipo.modelo, equipo.idequipo, equipo.numero_serial, equipo.reg_ivss FROM hospitales INNER JOIN equipo ON hospitales.idhospital = equipo.idhospital INNER JOIN subcategoria ON equipo.idsubcategoria = subcategoria.idsubcategoria INNER JOIN categorias ON subcategoria.idcategoria = categorias.idcategoria WHERE (subcategoria.idsubcategoria =" & Pn & " And Bitacora='" & Bitacora & "')"
    RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
    IdEquipo = RsListado.Fields("IdEquipo").Value
    TxtDescripcion.Text = RsListado.Fields("Descripcion").Value
    TxtMarca.Text = RsListado.Fields("marca").Value
    TxtModelo.Text = RsListado.Fields("modelo").Value
    TxtAno.Text = RsListado.Fields("ano").Value
    TxtSerial.Text = RsListado.Fields("numero_serial").Value
    TxtIVSS.Text = RsListado.Fields("reg_ivss").Value
    TxtBitacora.Text = RsListado.Fields("Bitacora").Value
    TxtUbicacion.Text = RsListado.Fields("Direccion").Value
    WebCnn.Close
    
    ConectarIVSSHosting
    CSql = "SELECT [idmen], [mens], [fecha] FROM [bitacora] WHERE ([idequipo] =" & IdEquipo & ")"
    RsListado.Open CSql, WebCnn, adOpenDynamic, adLockPessimistic
    
    If Not IsNull(RsListado.Fields("Fecha").Value) Then
        TxtDetalle.Text = Trim(TxtDetalle.Text) & RsListado.Fields("Fecha").Value & Chr(13) & RsListado.Fields("Mens").Value
    End If
    
    WebCnn.Close

End If



End Sub

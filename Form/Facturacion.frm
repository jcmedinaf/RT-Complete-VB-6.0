VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form277 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Facturaci�n"
   ClientHeight    =   9570
   ClientLeft      =   5130
   ClientTop       =   2430
   ClientWidth     =   12195
   LinkTopic       =   "Form27"
   ScaleHeight     =   9570
   ScaleWidth      =   12195
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Facturaci�n "
      Height          =   8175
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11895
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   9720
         TabIndex        =   38
         Top             =   7320
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   9720
         TabIndex        =   37
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Paciente"
         Height          =   1695
         Left            =   240
         TabIndex        =   27
         Top             =   2640
         Width           =   5895
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   3840
            TabIndex        =   40
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   960
            TabIndex        =   39
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   960
            TabIndex        =   30
            Top             =   1080
            Width           =   4815
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   3840
            TabIndex        =   29
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   960
            TabIndex        =   28
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cedula "
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   3075
            TabIndex        =   34
            Top             =   390
            Width           =   555
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido"
            Height          =   195
            Left            =   195
            TabIndex        =   33
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direcci�n"
            Height          =   195
            Left            =   195
            TabIndex        =   32
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono"
            Height          =   195
            Left            =   3075
            TabIndex        =   31
            Top             =   720
            Width           =   630
         End
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   9720
         TabIndex        =   23
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cliente"
         Height          =   2175
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   5895
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3000
            TabIndex        =   22
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   3120
            TabIndex        =   20
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Persona Contacto"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Email.:"
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direcci�n"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rif.: "
            Height          =   195
            Left            =   3120
            TabIndex        =   17
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razon Social"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Importar"
         Height          =   375
         Left            =   10440
         TabIndex        =   14
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Height          =   1575
         Index           =   1
         Left            =   6360
         TabIndex        =   7
         Top             =   360
         Width           =   5175
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3360
            TabIndex        =   36
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pago"
            Height          =   375
            Left            =   2640
            TabIndex        =   26
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   25
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   24
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N�mero"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   645
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FACTURA"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   720
            TabIndex        =   8
            Top             =   120
            Width           =   1845
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   3720
         TabIndex        =   1
         Top             =   6960
         Width           =   3855
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "Facturacion.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Nueva Factura"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   1
            Left            =   840
            Picture         =   "Facturacion.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Factura"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   2
            Left            =   1560
            Picture         =   "Facturacion.frx":09DC
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Factura"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   3
            Left            =   2280
            Picture         =   "Facturacion.frx":0C5B
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Agregar Platillos"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   4
            Left            =   3000
            Picture         =   "Facturacion.frx":1070
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Factura"
            Top             =   240
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   240
         TabIndex        =   44
         Top             =   4680
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   12648447
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Cod_producto"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "descripcion"
            Caption         =   "Tratamiento de Radioterapia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cantidad"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Precio"
            Caption         =   "Precio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Iva"
            Caption         =   "Iva"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Descuento"
            Caption         =   "Descuento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   4500,284
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1590,236
            EndProperty
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal"
         Height          =   195
         Left            =   8880
         TabIndex        =   13
         Top             =   6960
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         Height          =   195
         Left            =   8880
         TabIndex        =   12
         Top             =   7320
         Width           =   255
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   8880
         TabIndex        =   11
         Top             =   7680
         Width           =   360
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   120
      Picture         =   "Facturacion.frx":1496
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form277"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim BD61 As New ADODB.Recordset 'Tabla paciente
Dim BD62 As New ADODB.Recordset 'Tabla Informe medico
Dim BD63 As New ADODB.Recordset 'Tabla presupuestos
Dim BD64 As New ADODB.Recordset
Dim BD65 As New ADODB.Recordset
Dim BD66 As New ADODB.Recordset
Dim BD68 As New ADODB.Recordset
Dim BD69 As New ADODB.Recordset

Private Sub boton_Click(Index As Integer)
Select Case Index
        Case 0: Call Agregar
        Case 1: Call imprimir
        Case 2: Call guardar
        Case 3: Call Editar
        Case 4: Unload Me
    End Select
End Sub
Sub imprimir()

End Sub
Sub Agregar()

End Sub
Sub Editar()
 Dim i As Integer
    
    If DataGrid1.Row = -1 Then Exit Sub
    
    With Form29
        
        ' llena los campos
         If Trim(BD57.Fields("Cod_producto")) <> "" Then .Text1(0).Text = BD57.Fields("Cod_producto") Else .Text1(0).Text = ""
         If Trim(BD57.Fields("Precio")) <> "" Then .Text1(1).Text = BD57.Fields("Precio") Else .Text1(1).Text = ""
         If Trim(BD57.Fields("Cantidad")) <> "" Then .Text1(2).Text = BD57.Fields("Cantidad") Else .Text1(2).Text = ""
         If Trim(BD57.Fields("Descuento")) <> "" Then .Text1(4).Text = BD57.Fields("Descuento") Else .Text1(4).Text = ""
         If Trim(BD57.Fields("descripcion")) <> "" Then .Text1(5).Text = BD57.Fields("descripcion") Else .Text1(5).Text = ""
        
        
        .Label3 = BD57(9)
        ACCION = EDITAR_REGISTRO
        
        .Show vbModal
        DataGrid1.Refresh
        
    End With

End Sub


Sub guardar()
CSql = "Insert Into C_cobrar(idpaciente,idusuario,Forma_pago,N_Factura,Fecha,monto) VALUES(" & IdPac1 & "," & IdUser & "," & Combo1.ItemData(Combo1.ListIndex) & ",'" & Label12.Caption & "','" & Label16.Caption & "','" & Text8.Text & "')"
            
   Dim BD68 As New ADODB.Recordset
   BD68.Open CSql, CADENA
   msg = "Registro Agregado satisfactoriamente"
   MsgBox msg, vbOKOnly

End Sub

Private Sub Command1_Click()

msg = "Indique el Presupuesto del paciente "
ced = Trim(InputBox(msg, "Presupuesto del paciente", "12345678"))
If ced = "" Then Exit Sub

CSql = "select * from Paciente where cedula = " & ced
BD62.Open CSql, CADENA

If Not (BD62.EOF) Then

         Text9.Text = BD62.Fields("cedula")
        Text10.Text = BD62.Fields("nombre")
        Text11.Text = BD62.Fields("apellido")
        Text12.Text = BD62.Fields("telefono")
        Text13.Text = BD62.Fields("Direccion")
IdPac1 = BD62.Fields("idpaciente")

    csql1 = "select Razon,rif,direccion,contacto,email from Cliente where idpaciente = " & IdPac1
    BD63.Open csql1, CADENA

    If Not (BD63.EOF) Then
        Text1.Text = BD63.Fields("razon")
        Text2.Text = BD63.Fields("Rif")
        Text3.Text = BD63.Fields("Direccion")
        Text4.Text = BD63.Fields("Contacto")
        Text5.Text = BD63.Fields("Email")
    Call Nfac
Else
    
    msg = "No existe Registro"
    MsgBox msg
End If

Call IniciarConexion

        CSql = "select * from Presupuesto2 where idpaciente = " & IdPac1
        BD57.Open CSql, CNn, adOpenStatic, adLockOptimistic
        DataGrid1.AllowUpdate = False
        
Call CargarDataGrid(DataGrid1)

Call Precio
Call Pago1
Call Factura

    BD62.Close
    BD63.Close

End If
End Sub
Sub Nfac()

    CSql = "select * from C_cobrar where idpaciente = " & IdPac1
    BD66.Open CSql, CADENA

If Not (BD66.EOF) Then
        
    N_fac = BD66.Fields("N_Factura")
End If
    End Sub
Sub CargarDataGrid(dg As DataGrid)
    
    dg.MarqueeStyle = dbgHighlightRow
    Set dg.DataSource = BD57
    dg.Refresh
End Sub

Private Sub Command2_Click()
Form18.Show 1
End Sub

Private Sub Form_Load()
 SQL = ""
 IdPac1 = ""
 N_fac = ""
 t = 0
 Label16.Caption = Format(Now, "dd/mm/yyyy")
 Call Pago
 
End Sub

Sub IniciarConexion()
      If CNn.State = adStateOpen Then CNn.Close
        CNn.CursorLocation = adUseClient
        CNn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              "z:\oa.mdb" & ";Persist Security Info=False"
          
End Sub
Sub Pago()

            Dim BD61 As New ADODB.Recordset
            CSql = "SELECT * FROM Pago"
            BD61.Open CSql, CADENA, , , adCmdText
            BD61.MoveFirst
Combo1.Clear
           Do While Not BD61.EOF
                Combo1.AddItem BD61.Fields("Tipo")
                Combo1.ItemData(Combo1.NewIndex) = BD61.Fields("idUsuario")
        BD61.MoveNext
    Loop
    End Sub
    Sub Pago1()
            
            Dim BD65 As New ADODB.Recordset
            CSql = "select * from C_cobrar where idpaciente = " & IdPac1
            BD65.Open CSql, CADENA, , , adCmdText

                    If Not (BD65.EOF) Then
                                              
                        For t = 0 To Combo1.ListCount - 1
                          If Combo1.ItemData(t) = BD65.Fields("Forma_Pago") Then
                          Combo1.ListIndex = t
                          Exit For
                          End If
                        Next t
                    
    BD65.Close
        Exit Sub
            End If
    End Sub
   
Sub Precio()

CSql = "SELECT SUM(Precio) as monto2 FROM Presupuesto2 WHERE N_factura = " & N_fac

BD69.Open CSql, CADENA, , , adCmdText

If BD69.Fields("Monto2") <> "" Then Text8.Text = BD69.Fields("Monto2") Else Text8.Text = ""

End Sub
Sub Factura()

CSql = "select * from dat_admin"
        Dim BDF As New ADODB.Recordset
        BDF.Open CSql, CADENA
        Fact = Format(BDF.Fields("U_Factura") + 1, "000000000#")
        BDF.Close
        Label12.Caption = Fact
        CSql = "update dat_admin SET U_Factura = " & Str(Fact) & " WHERE U_Factura = " & Str(Fact - 1) & ";"
        BDF.Open CSql, CADENA
                
  End Sub


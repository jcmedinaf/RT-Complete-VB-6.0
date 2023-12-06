VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDatosClientes 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   7605
   ClientLeft      =   4920
   ClientTop       =   1590
   ClientWidth     =   8040
   Icon            =   "Cliente.frx":0000
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   8040
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   6720
      Width           =   7815
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   6720
         TabIndex        =   22
         ToolTipText     =   "Cerrar "
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
         MICON           =   "Cliente.frx":1002
         PICN            =   "Cliente.frx":101E
         PICH            =   "Cliente.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   375
         Left            =   1200
         TabIndex        =   23
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar"
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
         MICON           =   "Cliente.frx":141C
         PICN            =   "Cliente.frx":1438
         PICH            =   "Cliente.frx":16C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Agregar "
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Agregar"
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
         MICON           =   "Cliente.frx":1B08
         PICN            =   "Cliente.frx":1B24
         PICH            =   "Cliente.frx":1CB1
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
         Height          =   375
         Left            =   5520
         TabIndex        =   25
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         MICON           =   "Cliente.frx":1EE6
         PICN            =   "Cliente.frx":1F02
         PICH            =   "Cliente.frx":21E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Cliente.frx":2435
         PICN            =   "Cliente.frx":2451
         PICH            =   "Cliente.frx":26E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior 
         Height          =   375
         Left            =   3960
         TabIndex        =   27
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Cliente.frx":2946
         PICN            =   "Cliente.frx":2962
         PICH            =   "Cliente.frx":2BF7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnEliminar 
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Borrar"
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
         MICON           =   "Cliente.frx":2E53
         PICN            =   "Cliente.frx":2E6F
         PICH            =   "Cliente.frx":3013
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
      Caption         =   "Datos del Cliente"
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Cliente.frx":3453
         Left            =   3600
         List            =   "Cliente.frx":3460
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Left            =   1680
         TabIndex        =   3
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text9 
         Height          =   350
         Left            =   1680
         TabIndex        =   9
         Top             =   5520
         Width           =   6015
      End
      Begin VB.TextBox Text8 
         Height          =   350
         Left            =   1680
         TabIndex        =   8
         Top             =   5040
         Width           =   6015
      End
      Begin VB.TextBox Text7 
         Height          =   350
         Left            =   1680
         TabIndex        =   7
         Top             =   4560
         Width           =   6015
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Left            =   1680
         TabIndex        =   6
         Top             =   4080
         Width           =   6015
      End
      Begin VB.TextBox Text5 
         Height          =   765
         Left            =   1680
         TabIndex        =   5
         Top             =   3240
         Width           =   6015
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Left            =   1680
         TabIndex        =   4
         Top             =   2760
         Width           =   6015
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox Text2 
         Height          =   825
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   6015
      End
      Begin ChamaleonButton.ChameleonBtn BtnListadoProductos 
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   6000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Listado Clientes"
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
         MICON           =   "Cliente.frx":3485
         PICN            =   "Cliente.frx":34A1
         PICH            =   "Cliente.frx":372A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblRegistro 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro: 0 / 0  "
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   6120
         Width           =   5175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2355
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   5595
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   5085
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   4635
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección Factura:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   4155
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección Fiscal:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3525
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Comercial:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2835
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   918
         Width           =   990
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RIF.:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   405
         Width           =   345
      End
   End
End
Attribute VB_Name = "FrmDatosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cambio
Dim RegNew
Public RsClientes As New ADODB.Recordset
Public CodClient
Public CodClientIdL
Dim RsTemp As New ADODB.Recordset
Dim RsCount As New ADODB.Recordset
Sub EnviarRegPendiente(ByVal IdPac2 As Integer, ByVal IdLIdPac2 As String)
On Error Resume Next

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "SELECT * FROM Cliente WHERE IdCliente = " & IdPac2 & " AND IdL = '" & IdLIdPac2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Cliente (["
For i = 0 To RsTemp.Fields.Count - 1
    If Not i = (RsTemp.Fields.Count - 1) Then
        StrSen = StrSen & RsTemp.Fields(i).Name & "],["
    Else
        StrSen = StrSen & RsTemp.Fields(i).Name & "]) VALUES ("
    End If
Next i
For i = 0 To RsTemp.Fields.Count - 1
    If Not i = (RsTemp.Fields.Count - 1) Then
        StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "',"
    Else
        StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "')"
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = Replace(StrSen, "'", "(varCSP)")

CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Cliente"
RsRegPendiente.Fields("Tabla").Value = "Cliente"
RsRegPendiente.Fields("Condicional").Value = "IdCliente = " & IdPac2 & " AND IdL = '" & IdLIdPac2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub


Sub verify()

Select Case Cambio
Case Is = 1
    Msg = "Este registro sufrió Cambios desea guardar?"
    d = MsgBox(Msg, vbQuestion + vbYesNo, "Desea Guardar Cambios")
    Select Case d
        Case Is = vbYes
            If RegNew = 0 Then
                GuardarCambios
            Else
                Guardar
            End If
    
        Case Is = vbNo
    End Select
Case Is = 0
End Select

End Sub
Public Sub CargaDatos()
On Error Resume Next
CodClientIdL = ""
CodClient = 0

If RsClientes.RecordCount = 0 Then Exit Sub

If Not RsClientes.EOF Then

    LblRegistro.Caption = "Registro: " & RsClientes.AbsolutePosition & " / " & RsClientes.RecordCount
    
    ' Razón social
    If IsNull(RsClientes.Fields("Razon")) Then Text1.Text = "" Else Text1.Text = RsClientes.Fields("Razon")
    
    ' Dirección
    If IsNull(RsClientes.Fields(DireccionC)) Then Text2.Text = "" Else Text2.Text = RsClientes.Fields("DireccionC")
    
    ' R.I.F.
    If IsNull(RsClientes.Fields("Rif")) Then Text3.Text = "" Else Text3.Text = RsClientes.Fields("Rif")
    
    ' Nombre comercial
    If IsNull(RsClientes.Fields("NombreC")) Then Text4.Text = "" Else Text4.Text = RsClientes.Fields("NombreC")
    
    ' Dirección Fiscal
    If IsNull(RsClientes.Fields("Fiscal")) Then Text5.Text = "" Else Text5.Text = RsClientes.Fields("Fiscal")
    
    ' Dirección factura
    If IsNull(RsClientes.Fields("DirFac")) Then Text6.Text = "" Else Text6.Text = RsClientes.Fields("DirFac")
    
    ' Representante
    If IsNull(RsClientes.Fields("Representante")) Then Text7.Text = "" Else Text7.Text = RsClientes.Fields("Representante")
    
    ' Contacto
    If IsNull(RsClientes.Fields("Contacto")) Then Text8.Text = "" Else Text8.Text = RsClientes.Fields("Contacto")
    
    ' E-Mail
    If IsNull(RsClientes.Fields("EMail")) Then Text9.Text = "" Else Text9.Text = RsClientes.Fields("EMail")
    
    ' Teléfono
    If IsNull(RsClientes.Fields("Telefono")) Then Text10.Text = "" Else Text10.Text = RsClientes.Fields("Telefono")
    
    
    If IsNull(RsClientes.Fields("Personal")) Then
        Combo1.ListIndex = -1
    Else
        Combo1.ListIndex = Val(RsClientes.Fields("Personal"))
    End If
    
    CodClient = Val(RsClientes.Fields("IdCliente").Value)
    CodClientIdL = RsClientes.Fields("IdL").Value
Else
    LblRegistro.Caption = "Registro: 0 / 0 "
End If

End Sub
Sub Guardar()
Dim IdPerso As Integer
If Combo1.ListIndex = -1 Then
    IdPerso = -1
Else
    IdPerso = Combo1.ListIndex
End If

CSql = "Select MAX(idcliente)+1 As NuevoId From Cliente"
Set RsCount = CrearRS(CSql)

If RsCount.RecordCount <> 0 Then
    If Not IsNull(RsCount.Fields("NuevoId")) Then
        C = RsCount.Fields("NuevoId")
    Else
        C = "1"
    End If
Else
    C = "1"
End If

RsCount.Close

CSql = "Select * From Cliente"

Set RsTemp = CrearRS(CSql)

RsTemp.AddNew
RsTemp.Fields("IdCliente").Value = C
RsTemp.Fields("RAZON").Value = Text1.Text
RsTemp.Fields("DIRECCIONc").Value = Text2.Text
RsTemp.Fields("RIF").Value = Text3.Text
RsTemp.Fields("NOMBREC").Value = Text4.Text
RsTemp.Fields("FISCAL").Value = Text5.Text
RsTemp.Fields("DIRFAC").Value = Text6.Text
RsTemp.Fields("REPRESENTANTE").Value = Text7.Text
RsTemp.Fields("CONTACTO").Value = Text8.Text
RsTemp.Fields("EMAIL").Value = Text9.Text
RsTemp.Fields("Telefono").Value = Trim(Text10.Text)
RsTemp.Fields("Personal").Value = IdPerso
RsTemp.Fields("IdUsuario").Value = IdUser
RsTemp.Fields("IdL").Value = IdLDefault
RsTemp.Update

MsgBox "Registro Agregado satisfactoriamente", vbOKOnly + vbInformation, "Operación exitosa."

EnviarRegPendiente C, IdLDefault

Call Blanqueo
End Sub
Sub GuardarCambios()
Dim IdPerso As Integer

If Combo1.ListIndex = -1 Then
    IdPerso = -1
Else
    IdPerso = Combo1.ListIndex
End If

CSql = "SELECT * FROM CLIENTE WHERE IDCLIENTE = " & CodClient & " And IdL='" & CodClientIdL & "'"

Set RsTemp = CrearRS(CSql)

RsTemp.Fields("RAZON").Value = Text1.Text
RsTemp.Fields("DIRECCIONc").Value = Text2.Text
RsTemp.Fields("RIF").Value = Text3.Text
RsTemp.Fields("NOMBREC").Value = Text4.Text
RsTemp.Fields("FISCAL").Value = Text5.Text
RsTemp.Fields("DIRFAC").Value = Text6.Text
RsTemp.Fields("REPRESENTANTE").Value = Text7.Text
RsTemp.Fields("CONTACTO").Value = Text8.Text
RsTemp.Fields("EMAIL").Value = Text9.Text
RsTemp.Fields("Telefono").Value = Trim(Text10.Text)
RsTemp.Fields("Personal").Value = IdPerso
RsTemp.Fields("IdUsuario").Value = IdUser

RsTemp.Update

MsgBox "Registro Actualizado satisfactoriamente", vbOKOnly + vbInformation, "Operación exitosa."

EnviarRegPendiente CodClient, CodClientIdL

Call Blanqueo
End Sub
Sub Blanqueo()
CodClient = 0
CodClientIdL = IdLDefault

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
             
End Sub

Private Sub BtnAgregar_Click()
'command2
verify
Blanqueo
RegNew = 1
LblRegistro.Caption = "Nuevo Registro  "

BtnEliminar.Enabled = False
BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
BtnSiguiente.Enabled = False
BtnAnterior.Enabled = False

BtnListadoProductos.Enabled = False
Frame2.BackColor = &HE0E0E0
End Sub

Private Sub BtnAnterior_Click()
verify
If RsClientes.RecordCount <> 0 Then
    If RsClientes.BOF Then
        RsClientes.MoveLast
    Else
        RsClientes.MovePrevious
        If RsClientes.BOF Then RsClientes.MoveLast
    End If
    CargaDatos
Else
    MsgBox "Estimado usuario, No se encontraron registros!", vbExclamation + vbOKOnly + "Disculpe."
End If
Cambio = 0
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
RegNew = 0
CargaDatos

BtnEliminar.Enabled = True
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = True
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True

BtnListadoProductos.Enabled = True
Frame2.BackColor = &HEAEFEF
End Sub

Private Sub BtnGuardarActualizar_Click()
'command1

If Combo1.ListIndex = -1 Then
    MsgBox "Debe seleccionar el tipo de cliente!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo1.SetFocus
    Exit Sub
End If

If RegNew = 0 Then
    GuardarCambios
ElseIf RegNew = 1 Then
    Guardar
End If

RsClientes.Close
Conectar
RsClientes.MoveFirst
CargaDatos
Cambio = 0
RegNew = 0
End Sub

Private Sub BtnListadoProductos_Click()
On Error Resume Next
ModulO = 0
FrmListadoClientes.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnSiguiente_Click()
verify
If RsClientes.RecordCount <> 0 Then
    If RsClientes.EOF Then
        RsClientes.MoveFirst
    Else
        RsClientes.MoveNext
        If RsClientes.EOF Then RsClientes.MoveFirst
    End If
    CargaDatos
Else
    MsgBox "Estimado usuario, No se encontraron registros!", vbExclamation + vbOKOnly + "Disculpe."
End If
Cambio = 0
End Sub

Sub Conectar()

CSql = "SELECT * FROM CLIENTE"
Set RsClientes = CrearRS(CSql)

End Sub

Private Sub Form_Load()
Centrar Me
Conectar
CargaDatos
Cambio = 0
RegNew = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
RsClientes.Close
End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' De aqui para bajo es codigo de cajas de texto...
Private Sub Text1_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text1.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text1.Text)
    pru = LCase(Mid(Text1.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text1.Text = StrText
Text1.SelStart = Len(Text1.Text)

End Sub

Private Sub Text2_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text2.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text2.Text)
    pru = LCase(Mid(Text2.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text2.Text = StrText
Text2.SelStart = Len(Text2.Text)

End Sub

Private Sub Text3_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text3.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text3.Text)
    pru = LCase(Mid(Text3.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text3.Text = StrText
Text3.SelStart = Len(Text3.Text)

End Sub

Private Sub Text4_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text4.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text4.Text)
    pru = LCase(Mid(Text4.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text4.Text = StrText
Text4.SelStart = Len(Text4.Text)

End Sub

Private Sub Text5_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text5.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text5.Text)
    pru = LCase(Mid(Text5.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text5.Text = StrText
Text5.SelStart = Len(Text5.Text)

End Sub

Private Sub Text6_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text6.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text6.Text)
    pru = LCase(Mid(Text6.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text6.Text = StrText
Text6.SelStart = Len(Text6.Text)


End Sub

Private Sub Text7_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text7.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text7.Text)
    pru = LCase(Mid(Text7.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text7.Text = StrText
Text7.SelStart = Len(Text7.Text)

End Sub

Private Sub Text8_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text8.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text8.Text)
    pru = LCase(Mid(Text8.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text8.Text = StrText
Text8.SelStart = Len(Text8.Text)

End Sub

Private Sub Text9_Change()
Dim StrText, Chaa, pru As String
Dim i  As Variant
Cambio = 1
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text9.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text9.Text)
    pru = LCase(Mid(Text9.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text9.Text = StrText
Text9.SelStart = Len(Text9.Text)

End Sub

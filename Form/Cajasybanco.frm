VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCajasBancos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bancos"
   ClientHeight    =   5280
   ClientLeft      =   6180
   ClientTop       =   3525
   ClientWidth     =   7890
   Icon            =   "Cajasybanco.frx":0000
   LinkTopic       =   "Form33"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7890
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   4320
         Width           =   7455
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6360
            TabIndex        =   14
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
            MICON           =   "Cajasybanco.frx":1002
            PICN            =   "Cajasybanco.frx":101E
            PICH            =   "Cajasybanco.frx":11E7
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
            TabIndex        =   10
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
            MICON           =   "Cajasybanco.frx":141C
            PICN            =   "Cajasybanco.frx":1438
            PICH            =   "Cajasybanco.frx":16C7
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
            TabIndex        =   9
            ToolTipText     =   "Agregar"
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
            MICON           =   "Cajasybanco.frx":1B08
            PICN            =   "Cajasybanco.frx":1B24
            PICH            =   "Cajasybanco.frx":1CB1
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
            Left            =   5160
            TabIndex        =   13
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
            MICON           =   "Cajasybanco.frx":1EE6
            PICN            =   "Cajasybanco.frx":1F02
            PICH            =   "Cajasybanco.frx":21E4
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
            Left            =   4320
            TabIndex        =   12
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
            MICON           =   "Cajasybanco.frx":2435
            PICN            =   "Cajasybanco.frx":2451
            PICH            =   "Cajasybanco.frx":26E7
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
            Left            =   3720
            TabIndex        =   11
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
            MICON           =   "Cajasybanco.frx":2946
            PICN            =   "Cajasybanco.frx":2962
            PICH            =   "Cajasybanco.frx":2BF7
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
            TabIndex        =   25
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
            MICON           =   "Cajasybanco.frx":2E53
            PICN            =   "Cajasybanco.frx":2E6F
            PICH            =   "Cajasybanco.frx":3013
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Registro Banco"
         Height          =   4095
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   7455
         Begin VB.TextBox Text6 
            Height          =   735
            Left            =   1200
            TabIndex        =   8
            Top             =   3240
            Width           =   6135
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1200
            TabIndex        =   7
            Top             =   2760
            Width           =   6135
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   2280
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1200
            TabIndex        =   1
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   6135
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1200
            TabIndex        =   4
            Top             =   1440
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   3960
            TabIndex        =   2
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   39940
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   3330
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contacto:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   2850
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   2370
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Cuenta:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   1530
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Registro:"
            Height          =   195
            Left            =   2760
            TabIndex        =   16
            Top             =   570
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "FrmCajasBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bdata As Recordset
Dim bdata1 As Recordset
Dim Cambio
Dim RegNew

Private Sub BtnAgregar_Click()
    Call verify
    Call borrar
    RegNew = 1
End Sub

Private Sub BtnAnterior_Click()
Call verify
If Not (bdata.BOF) Then bdata.MovePrevious
If Not (bdata.BOF) Then Call CargaDatos
Cambio = 0
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnEliminar_Click()
Dim RsBorrar As New ADODB.Recordset
Msg = "Estas Seguro de Borrar el Banco seleccionado?"
mensaje = MsgBox(Msg, vbYesNo + vbInformation, "Mensaje")

If mensaje = vbYes Then
    
    CSql = "Select * From CajasBancos where IdCajaBanco = '" & Text1.Text & "'"
    Set RsBorrar = CrearRS(CSql)
    
    RsBorrar.Delete
    
    Msg = "Banco borrado satisfactoriamente"
    mensaje = MsgBox(Msg, vbOKOnly + vbInformation, "Borrado")
End If
Form_Load
End Sub

Private Sub BtnGuardarActualizar_Click()
If RegNew = 0 Then
    Call GuardarCambios
ElseIf RegNew = 1 Then
    Call Guardar
End If

borrar

Form_Load
End Sub

Private Sub BtnSiguiente_Click()
Call verify
If Not (bdata.EOF) Then bdata.MoveNext
If Not (bdata.EOF) Then Call CargaDatos

Cambio = 0
End Sub

Private Sub Combo1_Change()
Cambio = 1
End Sub

Private Sub DTPicker1_Change()
Cambio = 1
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyLeft
            Text1.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            Text2.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me
Combo1.Clear
Combo1.AddItem "Ahorro"
Combo1.ItemData(Combo1.NewIndex) = 1
Combo1.AddItem "Corriente"
Combo1.ItemData(Combo1.NewIndex) = 2
Combo1.AddItem "Activos Liquidos"
Combo1.ItemData(Combo1.NewIndex) = 3
Combo1.AddItem "FAL"
Combo1.ItemData(Combo1.NewIndex) = 4

CSql = "Select * From CajasBancos"
Set bdata = CrearRS(CSql)

Call CargaDatos
Cambio = 0
RegNew = 0


End Sub
Sub CargaDatos()
If Not bdata.EOF Then
    If IsNull(bdata.Fields("IdCajaBanco").Value) Then Text1.Text = "" Else Text1.Text = bdata.Fields("IdCajaBanco").Value
    If IsNull(bdata.Fields("Descripcion").Value) Then Text2.Text = "" Else Text2.Text = bdata.Fields("Descripcion").Value
    If IsNull(bdata.Fields("N_Cuenta").Value) Then Text3.Text = "" Else Text3.Text = bdata.Fields("N_Cuenta").Value
    If IsNull(bdata.Fields("Contacto").Value) Then Text5.Text = "" Else Text5.Text = bdata.Fields("Contacto").Value
    If IsNull(bdata.Fields("Sucursal").Value) Then Text4.Text = "" Else Text4.Text = bdata.Fields("Sucursal").Value
    If IsNull(bdata.Fields("Direccion").Value) Then Text6.Text = "" Else Text6.Text = bdata.Fields("Direccion").Value
    
    
'        If IsNull(bdata.Fields("TipoCuenta").Value) Then
'            Combo1.ListIndex = -1
'            Exit For
'        ElseIf Combo1.ItemData(i) = bdata.Fields("TipoCuenta").Value Then
'            Combo1.ListIndex = i
'            Exit For
'        End If
'
    
    
    For i = 1 To Combo1.ListCount - 1
        If bdata.Fields("TipoCuenta").Value = 0 Then
            CboStatus.ListIndex = 0
        Else
            CboStatus.ListIndex = i
        End If
    Next i
    
    DTPicker1.Value = bdata.Fields("Fecha_Registro").Value
End If

End Sub

Sub borrar()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.ListIndex = -1
DTPicker1.Value = Now

End Sub

Sub verify()
Select Case Cambio
    Case Is = 1
        Msg = "Este registro sufrió Cambios desea guardar?"
        d = MsgBox(Msg, vbYesNo, "Desea Guardar Cambios")
        Select Case d
            Case Is = 6
                If RegNew = 0 Then
                    Call GuardarCambios
                Else
                    Call Guardar
                End If
            Case Is = 7
        End Select
    Case Is = 0
End Select

End Sub
Sub GuardarCambios()
'CSql = "update CajasBancos set  descripcion = '" & Text2.Text & "', n_cuenta = '" & Text3.Text & "', fecha_registro = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "', TipoCuenta='" & Combo1.ItemData(Combo1.ListIndex) & "', Contacto='" & Text5.Text & "', Sucursal='" & Text4.Text & "', Direccion='" & Text6.Text & "' where IdCajaBanco = " & Val(Text1.Text)
'Set bdata1 = CrearRS(CSql)

CSql = "select * From CajasBancos where IdCajaBanco = '" & Text1.Text & "'"
Set bdata1 = CrearRS(CSql)

bdata1.Fields("descripcion").Value = Text2.Text
bdata1.Fields("n_cuenta").Value = Text3.Text
bdata1.Fields("fecha_registro").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
bdata1.Fields("TipoCuenta").Value = Combo1.ItemData(Combo1.ListIndex)
bdata1.Fields("Contacto").Value = Text5.Text
bdata1.Fields("Sucursal").Value = Text4.Text
bdata1.Fields("Direccion").Value = Text6.Text
bdata1.Update

Msg = "Registro Actualizado Satisfactoriamente"
MsgBox Msg, vbInformation + vbOKOnly, "Registro Actualizado"

End Sub

Private Sub Form_Unload(Cancel As Integer)
bdata.Close
End Sub

Private Sub Text1_Change()
Cambio = 1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyRight
            DTPicker1.SetFocus
        Case vbKeyDown
            Text2.SetFocus
    End Select
End If
End Sub

Private Sub Text2_Change()
Cambio = 1
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub Text3_Change()
Cambio = 1
End Sub

Sub Guardar()
Dim IdMax As Integer
CSql = "Select max(IdCajaBanco)+1 as MaxId From CajasBancos"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdMax = RsMaxId.Fields("MaxId").Value
Else
    IdMax = "1"
End If

CSql = "insert into CajasBancos(IdCajaBanco, descripcion, n_cuenta, fecha_registro, TipoCuenta, Sucursal, Contacto,  Direccion, IdUsuario) values('" & IdMax & "','" & Text2.Text & "','" & Text3.Text & "','" & Format(DTPicker1.Value, "dd/mm/yyyy") & "','" & Combo1.ItemData(Combo1.ListIndex) & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "'," & IdUser & ")"
Set bdata1 = CrearRS(CSql)
RegNew = 0
Msg = "Registro Agregado Satisfactoriamente"
MsgBox Msg, vbInformation + vbOKOnly, "Registro Guardado"
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text4_Change()
Cambio = 1
End Sub

Private Sub Text5_Change()
Cambio = 1
End Sub

Private Sub Text6_Change()
Cambio = 1
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDescripcionProductoServicio 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descripcion del Producto o Servicio"
   ClientHeight    =   5145
   ClientLeft      =   6555
   ClientTop       =   675
   ClientWidth     =   8295
   Icon            =   "Producto.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8295
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   24
         Top             =   3960
         Width           =   7815
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   6720
            TabIndex        =   12
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
            MICON           =   "Producto.frx":1002
            PICN            =   "Producto.frx":101E
            PICH            =   "Producto.frx":11E7
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
            Height          =   495
            Left            =   1200
            TabIndex        =   8
            ToolTipText     =   "Guardar / Actualizar "
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Producto.frx":141C
            PICN            =   "Producto.frx":1438
            PICH            =   "Producto.frx":16C7
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
            Height          =   495
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "Producto.frx":1B08
            PICN            =   "Producto.frx":1B24
            PICH            =   "Producto.frx":1CB1
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
            Left            =   5520
            TabIndex        =   11
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
            MICON           =   "Producto.frx":1EE6
            PICN            =   "Producto.frx":1F02
            PICH            =   "Producto.frx":21E4
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
            Height          =   495
            Left            =   4560
            TabIndex        =   10
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            MICON           =   "Producto.frx":2435
            PICN            =   "Producto.frx":2451
            PICH            =   "Producto.frx":26E7
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
            Height          =   495
            Left            =   3960
            TabIndex        =   9
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
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
            MICON           =   "Producto.frx":2946
            PICN            =   "Producto.frx":2962
            PICH            =   "Producto.frx":2BF7
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
            Height          =   495
            Left            =   2400
            TabIndex        =   25
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "Producto.frx":2E53
            PICN            =   "Producto.frx":2E6F
            PICH            =   "Producto.frx":3013
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
         Caption         =   "Producto"
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7815
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Si o No"
            Height          =   255
            Left            =   5280
            TabIndex        =   4
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1560
            TabIndex        =   3
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1560
            TabIndex        =   5
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1560
            TabIndex        =   1
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   2880
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55574529
            CurrentDate     =   39932
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Vencimiento"
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo:"
            Height          =   195
            Left            =   3720
            TabIndex        =   22
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion del Producto"
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Precio"
            Height          =   375
            Left            =   360
            TabIndex        =   20
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "IVA 12 %"
            Height          =   375
            Left            =   4320
            TabIndex        =   19
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   4320
            TabIndex        =   18
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Costo"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Sevicio"
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   1320
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "FrmDescripcionProductoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD90 As New ADODB.Recordset
Dim BD91 As New ADODB.Recordset
Dim bd92 As New ADODB.Recordset
Dim Cambio
Dim Nuevo

Sub Refrescar()

 CSql = "SELECT * FROM productos"
If BD90.State = adStateOpen Then BD90.Close
    BD90.CursorType = adOpenKeyset
    BD90.LockType = adLockOptimistic
    BD90.CursorLocation = adUseClient
    BD90.Open CSql, Cnn, , , adCmdText
   
End Sub
Sub CargaProveedor()

CSql = "Select idproveedor, nombre from proveedor"
bd92.Open CSql, Cnn
If Not (bd92.EOF) Then
bd92.MoveFirst
Do While Not bd92.EOF
    Combo1.AddItem bd92.Fields(1)
    Combo1.ItemData(Combo1.NewIndex) = bd92.Fields(0)
    bd92.MoveNext
Loop
bd92.Close
Else
bd92.Close

End If

End Sub

Sub Blanqueo()
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Label3.Caption = ""
            Combo1.ListIndex = -1
            Nuevo = 1
End Sub

Private Sub BtnAgregar_Click()
'command2
Blanqueo
Nuevo = 1
End Sub

Private Sub BtnAnterior_Click()
'command6
BD90.MovePrevious
Call carga
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
'command1
Select Case Nuevo

Case Is = 1
        CSql = "Insert into Productos(IdUsuario,descripcion,precio,costo,tipo_ser,iva,idproveedor) VALUES(" & IdUser & ",'" & Text1.Text & "'," & Val(Text2.Text) & ",'" & Val(Text3.Text) & "','" & Text4.Text & "'," & Check1.Value & "," & Combo1.ItemData(Combo1.ListIndex) & ")"
        BD91.Open CSql, Cnn
        Msg = "Registro Agregado satisfactoriamente"
        Call Blanqueo
     
        Nuevo = 0
        
Case Is = 0
If Cambio = 1 Then
        If Combo1.ListIndex = -1 Then
        IdProv = 0
        Else
        IdProv = Combo1.ItemData(Combo1.ListIndex)
        End If
        
        CSql = "update productos set descripcion = '" & Text1.Text & "', precio = " & Val(Text2.Text) & ", costo = " & Val(Text3.Text) & ", tipo_ser = '" & Text4.Text & "', iva = " & Check1.Value & ", idproveedor = " & IdProv & " where idproducto = " & Val(Label3.Caption)
        BD91.Open CSql, Cnn
        Msg = "Registro Actualizado Satisfactoriamente"
     
        Nuevo = 0

End If

End Select
MsgBox Msg, vbOKOnly
Call Refrescar
Call carga
End Sub

Private Sub BtnSiguiente_Click()
'command7
BD90.MoveNext
Call carga
End Sub

Private Sub Check1_Click()
Cambio = 1
End Sub

Private Sub Combo1_Click()
Cambio = 1
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyDown
            Text2.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker1_Change()
Cambio = 1
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text1.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            Text1.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me
Nuevo = 0
Call CargaProveedor
Call Refrescar
Call carga
    
End Sub
Sub carga()
    If BD90.EOF Then
        Msg = "Llego al Final del Registro desea Volver al Principio?"
        MsgBox Msg
        BD90.MoveFirst
    End If

If BD90.BOF Then
    Msg = "Llego al principio del registro"
    MsgBox Msg
    BD90.MoveLast
End If
                        
            If Trim(BD90.Fields("Idproducto")) <> "" Then Label3.Caption = BD90.Fields("Idproducto")
            If Trim(BD90.Fields("Descripcion")) <> "" Then Text1.Text = BD90.Fields("Descripcion")
            If Trim(BD90.Fields("Precio")) <> "" Then Text2.Text = BD90.Fields("Precio")
            If Trim(BD90.Fields("Costo")) <> "" Then Text3.Text = BD90.Fields("Costo")
            If Trim(BD90.Fields("Tipo_ser")) <> "" Then Text4.Text = BD90.Fields("Tipo_ser")
            If IsNull(BD90.Fields("idproveedor")) Then
            Combo1.ListIndex = -1
            Else
            For a = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(a) = BD90.Fields("idproveedor") Then
            d = a
            Exit For
            Else: d = -1
            End If
            Next a
            If d = "" Then Combo1.ListIndex = -1 Else Combo1.ListIndex = d
            End If
            If BD90.Fields("iva") Then Check1.Value = 1 Else Check1.Value = 0
      Cambio = 0
    
End Sub
                
Private Sub Form_Unload(Cancel As Integer)
If BD90.State = adStateOpen Then BD90.Close

End Sub

Private Sub Text1_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
 Dim i  As Variant
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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo1.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyDown
            Combo1.SetFocus
    End Select
End If
End Sub

Private Sub Text2_Change()
Cambio = 1

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii <> 8 Then
              
    If Not IsNumeric(Chr(KeyAscii)) Then
                  Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyRight
            Check1.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub Text3_Change()
Cambio = 1
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii <> 8 Then
              
    If Not IsNumeric(Chr(KeyAscii)) Then
                  Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text4_Change()
Cambio = 1
Dim StrText, Chaa, pru As String
 Dim i  As Variant
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

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

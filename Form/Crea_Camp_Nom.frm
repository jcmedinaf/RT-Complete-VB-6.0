VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmAgregaCampoNomina 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Campo Nomina"
   ClientHeight    =   4305
   ClientLeft      =   5145
   ClientTop       =   4545
   ClientWidth     =   7800
   Icon            =   "Crea_Camp_Nom.frx":0000
   LinkTopic       =   "Form42"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7800
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   7335
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   6240
            TabIndex        =   7
            ToolTipText     =   "Cerrar"
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
            MICON           =   "Crea_Camp_Nom.frx":1002
            PICN            =   "Crea_Camp_Nom.frx":101E
            PICH            =   "Crea_Camp_Nom.frx":11E7
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
            MICON           =   "Crea_Camp_Nom.frx":141C
            PICN            =   "Crea_Camp_Nom.frx":1438
            PICH            =   "Crea_Camp_Nom.frx":16C7
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
            TabIndex        =   9
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
            MICON           =   "Crea_Camp_Nom.frx":1B08
            PICN            =   "Crea_Camp_Nom.frx":1B24
            PICH            =   "Crea_Camp_Nom.frx":1CB1
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
            Left            =   5040
            TabIndex        =   10
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
            MICON           =   "Crea_Camp_Nom.frx":1EE6
            PICN            =   "Crea_Camp_Nom.frx":1F02
            PICH            =   "Crea_Camp_Nom.frx":21E4
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
            Left            =   4320
            TabIndex        =   11
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
            MICON           =   "Crea_Camp_Nom.frx":2435
            PICN            =   "Crea_Camp_Nom.frx":2451
            PICH            =   "Crea_Camp_Nom.frx":26E7
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
            Left            =   3720
            TabIndex        =   12
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
            MICON           =   "Crea_Camp_Nom.frx":2946
            PICN            =   "Crea_Camp_Nom.frx":2962
            PICH            =   "Crea_Camp_Nom.frx":2BF7
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
            TabIndex        =   19
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
            MICON           =   "Crea_Camp_Nom.frx":2E53
            PICN            =   "Crea_Camp_Nom.frx":2E6F
            PICH            =   "Crea_Camp_Nom.frx":3013
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
         Caption         =   "Datos del Campo "
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7335
         Begin VB.TextBox TxtEquivalencia 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   2400
            Width           =   1695
         End
         Begin VB.CheckBox ChkRestringido 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Restringido"
            Height          =   255
            Left            =   3600
            TabIndex        =   20
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CheckBox ChkInicializar 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Inicializar luego del cierre"
            Height          =   255
            Left            =   3600
            TabIndex        =   18
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox TxtValor 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox TxtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   840
            Width           =   5655
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Crea_Camp_Nom.frx":31B2
            Left            =   1440
            List            =   "Crea_Camp_Nom.frx":31BF
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equivalencia:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   2490
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Predeterminado:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1950
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   450
            Width           =   540
         End
         Begin VB.Label NReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro 0 / 0"
            Height          =   195
            Left            =   5520
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   930
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1403
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "FrmAgregaCampoNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDNon As Recordset
Dim BDNon1 As Recordset
Dim Cambio
Dim RegNew
Public IdCamp
Public TReg As Integer
'Public RsCampoNonima As New ADODB.Recordset
Public RsCampoNonima As Recordset
Public RsTemp As Recordset

Private Sub BtnAgregar_Click()
Call Blanqueo
BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
BtnEliminar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
RegNew = 1
Contar
IdCamp = ""
TxtValor.Text = "1"
NReg.Caption = "Nuevo Registro"
TxtCodigo.Text = Format(TReg, "0000")
TxtEquivalencia.Text = "1"
TxtDescripcion.SetFocus
End Sub

Private Sub BtnAnterior_Click()
If RsCampoNonima.RecordCount = 0 Then Exit Sub
If Cambio <> 0 Then Call verify
RsCampoNonima.MovePrevious
If RsCampoNonima.BOF Then RsCampoNonima.MoveLast
Call CargaDato
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
BtnEliminar.Enabled = True
Form_Load
End Sub

Private Sub BtnEliminar_Click()
Dim resp

resp = MsgBox("Se procedera a eliminar el registro actual, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar Operación")
If resp = 7 Then Exit Sub

CSql = "UPDATE CamposDeNomina SET Activo=0 Where IdCampoNomina=" & IdCamp
Set RsTemp = CrearRS(CSql)

MsgBox "El registro fue eliminado satisfactoriamente", vbInformation + vbOKOnly, "Operación Exitosa!"
Form_Load
End Sub

Private Sub BtnGuardarActualizar_Click()

If Trim(TxtDescripcion.Text) = "" Then MsgBox "Ingrese la descripcion para el Campo!", vbExclamation + vbOKOnly, "Error": Exit Sub
If Combo1.ListIndex = -1 Then MsgBox "Ingrese el tipo de Campo!", vbExclamation + vbOKOnly, "Error": Exit Sub

If Cambio <> 1 Then MsgBox "No se han realizado cambios!", vbInformation + vbOKOnly, "Informacion": Exit Sub

If RegNew = 1 Then
    CSql = "Insert Into CamposDeNomina(Idcamponomina, campo,predeterminado, tipo,Inicializar,Activo,Restringido,Equivalencia) VALUES(" & _
    TReg & ",'" & TxtDescripcion.Text & "','" & TxtValor.Text & "'," & Combo1.ListIndex & "," & _
    ChkInicializar.Value & ",1," & ChkRestringido.Value & "," & Replace(Trim(TxtEquivalencia.Text), ",", ".") & ")"
    Set BDNon1 = CrearRS(CSql)
    Msg = "Registro Agregado satisfactoriamente"
    MsgBox Msg, vbOKOnly
    Call Blanqueo
Else
    CSql = "Update CamposDeNomina set Inicializar=" & ChkInicializar.Value & ", Predeterminado='" & TxtValor.Text & _
    "', Campo = '" & TxtDescripcion.Text & "', Tipo = " & Combo1.ListIndex & ", restringido='" & _
    ChkRestringido.Value & "',Equivalencia=" & Replace(Trim(TxtEquivalencia.Text), ",", ".") & "  WHERE IdCampoNomina = " & IdCamp
    Set BDNon1 = CrearRS(CSql)
    Msg = "Registro Se Actualizo Satisfactoriamente"
    MsgBox Msg, vbOKOnly
End If
BtnDesHacer_Click
End Sub

Private Sub BtnSiguiente_Click()

If RsCampoNonima.RecordCount = 0 Then Exit Sub

If Cambio <> 0 Then Call verify
RsCampoNonima.MoveNext
If RsCampoNonima.EOF Then RsCampoNonima.MoveFirst
Call CargaDato
End Sub

Private Sub ChkInicializar_Click()
Cambio = 1
End Sub

Private Sub ChkRestringido_Click()
Cambio = 1
End Sub

Private Sub Combo1_Click()
Cambio = 1
End Sub

Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""
Combo1.ListIndex = -1
End Sub

Sub verify()
Msg = "El registro ha sufrido cambios desea guardar cambios??"
d = MsgBox(Msg, vbYesNo, "Guardar cambios")
If d = vbYes Then Call BtnGuardarActualizar_Click

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case 38
            TxtDescripcion.SetFocus
        Case 39
            BtnAyuda.SetFocus
        Case 40
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me
BtnAgregar.Enabled = True
CSql = "Select * From CamposDeNomina Where Activo=1"
Set RsCampoNonima = CrearRS(CSql)
TxtDescripcion.Text = ""
TxtValor.Text = ""
Combo1.ListIndex = -1
TReg = 1
IdCamp = ""
CargaDato
Cambio = 0
End Sub

Sub Contar()
Dim RstResultado As New ADODB.Recordset
CSql = "Select max(Idcamponomina)+1 as NuevoId From CamposDeNomina"
Set RstResultado = CrearRS(CSql)

If Not IsNull(RstResultado.Fields("NuevoId").Value) Then
    TReg = RstResultado.Fields("NuevoId").Value
Else
    TReg = "1"
End If

End Sub
Sub CargaDato()

If RsCampoNonima.RecordCount = 0 Then
    RegNew = 1
    BtnEliminar.Enabled = False
    BtnGuardarActualizar.Enabled = False
    BtnSiguiente.Enabled = False
    BtnAnterior.Enabled = False
    NReg = "No existen registros"
    Exit Sub
End If

If Trim(RsCampoNonima.Fields("Campo")) <> "" Then TxtDescripcion.Text = RsCampoNonima.Fields("Campo")
IdCamp = RsCampoNonima.Fields("IdCampoNomina")
Nom = Format(RsCampoNonima.Fields("IdCampoNomina"), "000#")
TxtCodigo.Text = Nom
TxtValor.Text = RsCampoNonima.Fields("Predeterminado").Value
TxtEquivalencia.Text = RsCampoNonima.Fields("Equivalencia").Value
If RsCampoNonima.Fields("Inicializar").Value Then
    ChkInicializar.Value = 1
Else
    ChkInicializar.Value = 0
End If

If Val(RsCampoNonima.Fields("Restringido").Value) = 1 Then
    ChkRestringido.Value = 1
Else
    ChkRestringido.Value = 0
End If


NReg.Caption = "Registro " & RsCampoNonima.AbsolutePosition & " / " & RsCampoNonima.RecordCount
For T = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(T) = RsCampoNonima.Fields("tipo") Then
        Combo1.ListIndex = T
        Exit For
    End If
Next T
Cambio = 0
RegNew = 0
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
BtnGuardarActualizar.Enabled = True
BtnSiguiente.Enabled = True
BtnAnterior.Enabled = True
End Sub
 
Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDescripcion.SetFocus
        Case 40
            TxtDescripcion.SetFocus
    End Select
End If
End Sub

Private Sub TxtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo1.SetFocus
        Case 40
            TxtCodigo.SetFocus
    End Select
End If
End Sub

Private Sub TxtDescripcion_Change()
Cambio = 1
End Sub

Private Sub TxtEquivalencia_KeyPress(KeyAscii As Integer)

' si no es un numero y es diferente de 8 y es diferente de 188 (una coma) entonces..
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 188 Then KeyAscii = 0
If Len(TxtEquivalencia.Text) > 5 Then KeyAscii = 0
Cambio = 1
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
Cambio = 1
End Sub

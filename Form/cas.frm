VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmAgregarTipoCancer 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Cancer"
   ClientHeight    =   5040
   ClientLeft      =   600
   ClientTop       =   1965
   ClientWidth     =   7785
   Icon            =   "cas.frx":0000
   LinkTopic       =   "Form35"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7785
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   3015
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   405
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Busqueda"
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
         MICON           =   "cas.frx":1002
         PICN            =   "cas.frx":101E
         PICH            =   "cas.frx":1283
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   7335
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6240
            TabIndex        =   6
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
            MICON           =   "cas.frx":1515
            PICN            =   "cas.frx":1531
            PICH            =   "cas.frx":16FA
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
            TabIndex        =   7
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
            MICON           =   "cas.frx":192F
            PICN            =   "cas.frx":194B
            PICH            =   "cas.frx":1BDA
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
            TabIndex        =   8
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
            MICON           =   "cas.frx":201B
            PICN            =   "cas.frx":2037
            PICH            =   "cas.frx":21C4
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
            Left            =   5040
            TabIndex        =   9
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
            MICON           =   "cas.frx":23F9
            PICN            =   "cas.frx":2415
            PICH            =   "cas.frx":26F7
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
            TabIndex        =   10
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
            MICON           =   "cas.frx":2948
            PICN            =   "cas.frx":2964
            PICH            =   "cas.frx":2BFA
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
            MICON           =   "cas.frx":2E59
            PICN            =   "cas.frx":2E75
            PICH            =   "cas.frx":310A
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
            TabIndex        =   13
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
            MICON           =   "cas.frx":3366
            PICN            =   "cas.frx":3382
            PICH            =   "cas.frx":3526
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
         Caption         =   "Caracteristicas del Cancer"
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7335
         Begin VB.TextBox TxtCodigo 
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   375
            Left            =   1200
            TabIndex        =   2
            Top             =   840
            Width           =   6015
         End
         Begin VB.Label LblNroReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
            Height          =   195
            Left            =   5280
            TabIndex        =   19
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro de registro:"
            Height          =   195
            Left            =   3960
            TabIndex        =   18
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   930
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   450
            Width           =   540
         End
      End
   End
   Begin SystemOncoAmerica.DMGrid DMGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
      Object.Width           =   7545
      Object.Height          =   2145
      BackColor       =   15396847
      ScrollBar       =   1
      MarqueeStyle    =   2
   End
End
Attribute VB_Name = "FrmAgregarTipoCancer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDCan As New ADODB.Recordset
Dim bdcan1 As New ADODB.Recordset
Dim actualiza As Integer
Dim RsTemp As New ADODB.Recordset

Private Sub BtnAgregar_Click()
Me.Height = 3200
DMGrid1.Enabled = False
actualiza = 0
Blanqueo
TxtDescripcion.SetFocus
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False

TxtCodigo.Text = "Nuevo Reg."
End Sub

Sub Leer_TCancers()

CSql = "SELECT * FROM Tipos_Ca"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

DMGrid1.Clear
DMGrid1.Rows = 0

While Not RsTemp.EOF

    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Id_Tipos_ca")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Descrip_Tipos_Ca")
    RsTemp.MoveNext
    
Wend

DMGrid1.RowBackColor 1, vbWhite
DMGrid1.PaintMGrid

End Sub
Private Sub BtnAnterior_Click()
actualiza = 1

BDCan.MovePrevious
Call CargaDatos

End Sub

Private Sub BtnBuscar_Click()


If Trim(Text2.Text) Then
    CSql = "SELECT * FROM Tipos_ca"
End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
DMGrid1.Enabled = True
Me.Height = 5520
actualiza = 0
Blanqueo
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
Leer_TCancers
End Sub

Private Sub BtnEliminar_Click()

If TxtCodigo.Text = "" Or TxtDescripcion.Text = "" Then
    Msg = "Debes de Seleccionar un Registro para poder borrarlo"
    mensaje = MsgBox(Msg, vbOKOnly, "Información")
    Exit Sub
End If

mensaje = MsgBox("Estas seguro de borrar el registro seleccionado", vbQuestion + vbYesNo, "Información")

If mensaje = vbNo Then Exit Sub

    CSql = "DELETE FROM Tipos_Ca WHERE Id_Tipos_ca=" & Val(TxtCodigo.Text)
    Set rs = CrearRS(CSql)
    
    MsgBox "Registro borrado satisfactoriamente", vbInformation + vbOKOnly, "Información"
    
    Blanqueo
    
    Leer_TCancers
    BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim Fecha As String
Dim RstResultado As New ADODB.Recordset
Dim Rsp As Byte
Fecha = Date

If TxtCodigo.Text = "" Then
            Msg = "Codigo Vacio"
            MsgBox Msg, vbOKOnly, "Información"
            TxtCodigo.SetFocus
            Exit Sub
End If
If TxtDescripcion.Text = "" Then
            Msg = "Descripción Vacia"
            MsgBox Msg, vbOKOnly, "Información"
            TxtDescripcion.SetFocus
            Exit Sub
End If

Rsp = MsgBox("Se procedera a guardar los cambios, Desea continuar?", vbQuestion + vbYesNo, "Confirmar!")

If Rsp = vbNo Then Exit Sub

If actualiza = 0 Then
    CSql = "SELECT MAX(Id_Tipos_ca)+1 As NuevoId FROM Tipos_Ca"
    Set rs = CrearRS(CSql)
    
    Dim NuevoId As Integer
    
    If Not IsNull(rs.Fields(0).Value) Then
        NuevoId = Val(rs.Fields(0).Value)
    Else
        NuevoId = 1
    End If
    
    CSql = "INSERT INTO  Tipos_Ca (Id_Tipos_ca,  Descrip_Tipos_ca,  idusuario,  Fecha_Can) " & _
    "VALUES (" & NuevoId & ",'" & Trim(TxtDescripcion.Text) & "','" & IdUser & "','" & Format(Now, "dd/MM/yyyy") & "')"
    
    Set rs = CrearRS(CSql)
    
    MsgBox "Registro Agregado satisfactoriamente", vbOKOnly + vbInformation, "Información"
    
Else

    CSql = "UPDATE Tipos_Ca SET Descrip_Tipos_ca='" & Trim(TxtDescripcion.Text) & "' WHERE Id_Tipos_ca = " & Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
    Set rs = CrearRS(CSql)
    
    Msg = "Registro Actualizado Satisfactoriamente"
    MsgBox Msg, vbOKOnly + vbInformation, "Información"

End If
Blanqueo

BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True

Leer_TCancers

End Sub

Private Sub BtnSiguiente_Click()

actualiza = 1

BDCan.MoveNext
Call CargaDatos

End Sub

Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""

End Sub


Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
'BtnDesHacer_Click

If lRow < 1 Then Exit Sub

BtnAgregar.Enabled = False
BtnEliminar.Enabled = True
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False

TxtCodigo.Text = DMGrid1.ValorCelda(lRow, 1)
TxtDescripcion.Text = DMGrid1.ValorCelda(lRow, 2)
DMGrid1.Row = lRow
actualiza = 1

End Sub

Private Sub Form_Load()
Centrar Me
If BDCan.State = 1 Then BDCan.Close
CSql = "SELECT * FROM TIPOS_CA"
Set BDCan = CrearRS(CSql)
Call CargaDatos

DMGrid1.Cols = 2

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Tipo de Cancer"

DMGrid1.DColumnas(1).Width = (DMGrid1.Width * 30) / 100
DMGrid1.DColumnas(2).Width = ((DMGrid1.Width * 70) / 100) - 300

Leer_TCancers
End Sub

Sub CargaDatos()

If BDCan.EOF Then BDCan.MoveFirst

If BDCan.BOF Then BDCan.MoveLast


TxtCodigo.Text = BDCan.Fields("Id_Tipos_ca")
TxtDescripcion.Text = BDCan.Fields("Descrip_Tipos_ca")

For i = 1 To DMGrid1.Rows
    If Val(DMGrid1.ValorCelda(i, 1)) = Val(BDCan.Fields("Id_Tipos_ca")) Then DMGrid1.Row = i
Next i

LblNroReg.Caption = BDCan.AbsolutePosition & " / " & BDCan.RecordCount
End Sub


Private Sub Form_Unload(Cancel As Integer)
If BDCan.State = 1 Then BDCan.Close
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            TxtDescripcion.SetFocus
        Case 39
            BtnAyuda.SetFocus
        Case 40
            TxtDescripcion.SetFocus
    End Select
End If
End Sub

Private Sub TxtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case 38
            TxtCodigo.SetFocus
        Case 40
            BtnAgregar.SetFocus
    End Select
End If
End Sub


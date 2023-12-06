VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmProtocolos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protocolos"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   Icon            =   "FrmProtocolos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   8250
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   7815
         Begin VB.TextBox TxtCodigo 
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   720
            Width           =   6495
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   810
            Width           =   885
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   7815
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
            Height          =   375
            Left            =   6600
            TabIndex        =   2
            ToolTipText     =   "Cerrar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "FrmProtocolos.frx":1002
            PICN            =   "FrmProtocolos.frx":101E
            PICH            =   "FrmProtocolos.frx":11E7
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
            TabIndex        =   4
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmProtocolos.frx":141C
            PICN            =   "FrmProtocolos.frx":1438
            PICH            =   "FrmProtocolos.frx":15C5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn4 
            Height          =   375
            Left            =   5400
            TabIndex        =   5
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
            MICON           =   "FrmProtocolos.frx":17FA
            PICN            =   "FrmProtocolos.frx":1816
            PICH            =   "FrmProtocolos.frx":1AF8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBorrar 
            Height          =   375
            Left            =   2520
            TabIndex        =   6
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
            MICON           =   "FrmProtocolos.frx":1D49
            PICN            =   "FrmProtocolos.frx":1D65
            PICH            =   "FrmProtocolos.frx":1F09
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
            TabIndex        =   7
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
            MICON           =   "FrmProtocolos.frx":20A8
            PICN            =   "FrmProtocolos.frx":20C4
            PICH            =   "FrmProtocolos.frx":235A
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
            TabIndex        =   8
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
            MICON           =   "FrmProtocolos.frx":25B9
            PICN            =   "FrmProtocolos.frx":25D5
            PICH            =   "FrmProtocolos.frx":286A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardar 
            Height          =   375
            Left            =   1320
            TabIndex        =   3
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
            MICON           =   "FrmProtocolos.frx":2AC6
            PICN            =   "FrmProtocolos.frx":2AE2
            PICH            =   "FrmProtocolos.frx":2D71
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
   End
End
Attribute VB_Name = "FrmProtocolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsProtocolos As New ADODB.Recordset
Dim Cambio
Dim IdProt
Dim IdLIdProt As String
Dim RegNuevo As Integer
Dim RegNew

Private Sub BtnAgregar_Click()

Blanqueo

RegNuevo = 1

TxtCodigo.Text = "Nuevo Reg."
TxtDescripcion.Locked = False
TxtDescripcion.SetFocus

BtnAgregar.Enabled = False
BtnGuardar.Enabled = True
BtnBorrar.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
Cambio = 0
End Sub

Sub Blanqueo()
TxtCodigo.Text = ""
TxtDescripcion.Text = ""
End Sub
Private Sub BtnAnterior_Click()
If RsProtocolos.RecordCount <> 0 Then
    If RsProtocolos.BOF Then
        RsProtocolos.MoveLast
    Else
        RsProtocolos.MovePrevious
        If RsProtocolos.BOF Then RsProtocolos.MoveLast
    End If
    Call Carga_De_Datos

    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If
End Sub

Sub Carga_De_Datos()

    IdProt = ""
    IdLIdProt = IdLDefault
    
    If RsProtocolos.RecordCount = 0 Then
        RegNuevo = 1
        Exit Sub
    End If
    
    RegNuevo = 0
    
    IdProt = RsProtocolos.Fields("Id").Value
    IdLIdProt = RsProtocolos.Fields("IdL").Value
    
    If Trim(RsProtocolos.Fields("Id").Value) <> "" Then IdProt = RsProtocolos.Fields("Id").Value Else IdProt = ""
    If Trim(RsProtocolos.Fields("Id").Value) <> "" Then TxtCodigo.Text = RsProtocolos.Fields("Id").Value Else TxtCodigo.Text = ""
    If Trim(RsProtocolos.Fields("Protocolo").Value) <> "" Then TxtDescripcion.Text = RsProtocolos.Fields("Protocolo").Value Else TxtDescripcion.Text = ""
    
End Sub

Private Sub BtnBorrar_Click()

mensaje = MsgBox("Estas Seguro de Eliminar el Protocolo?", vbYesNo + vbInformation, "Mensaje")

If mensaje = vbYes Then
    
    CSql = "Update Protocolos Set Activo='0' Where Id='" & TxtCodigo.Text & "' And IdL='" & IdLIdProt & "'"
    Set RsProtocolos = CrearRS(CSql)
        
        
    MsgBox "El Registro Borrado Exitosamente", vbInformation + vbOKOnly, "Operación Exitosa"
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
    EnviarRegPendiente IdProt, IdLIdProt
        
End If
Form_Load
End Sub

Private Sub BtnGuardar_Click()

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Bloque que verifica si hay internet
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

If TxtDescripcion.Text = "" Then
    MsgBox "Debe de ingresar la descripción del Protocolo!!", vbCritical + vbOKOnly, "Error"
    TxtDescripcion.SetFocus
    Exit Sub
End If

If Not IsNull(RsProtocolos.Fields("NuevaId")) Then
    RegNew = RsProtocolos.Fields("NuevaId").Value
Else
    RegNew = "1"
End If

Select Case RegNuevo
    
    Case Is = 1
        
        CSql = "Select * From Protocolos"
        Set RsProtocolos = CrearRS(CSql)
        
        IdProt = RegNew
        IdLIdProt = NuevoIdL
        
        RsProtocolos.AddNew
        RsProtocolos.Fields("Id").Value = IdProt
        RsProtocolos.Fields("IdL").Value = IdLIdProt
        RsProtocolos.Fields("Protocolo").Value = TxtDescripcion.Text
        RsProtocolos.Fields("IdUsuario").Value = IdUser
        RsProtocolos.Fields("Activo").Value = 1
        RsProtocolos.Update
        
        
        MsgBox "El Registro fue Agregado Exitosamente", vbInformation + vbOKOnly, "Operación Exitosa"
    
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el Servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
        
        EnviarRegPendiente IdProt, IdLIdProt
    
    Case Is = 0

        CSql = "Select * From Protocolos Where Id='" & IdProt & "' And IdL='" & IdLIdProt & "'"
        Set RsProtocolos = CrearRS(CSql)

        RsProtocolos.Fields("Protocolo").Value = TxtDescripcion.Text
        RsProtocolos.Fields("IdUsuario").Value = IdUser
        RsProtocolos.Update

        MsgBox "El Registro fue Actualizado Exitosamente", vbInformation + vbOKOnly, "Operación Exitosa"
        
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
        EnviarRegPendiente IdProt, IdLIdProt

End Select

Blanqueo
BtnAgregar.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = True
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
CargarProtocolos
Form_Load
Cambio = 1

End Sub


Sub EnviarRegPendiente(ByVal IdNut As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Protocolos WHERE Id = " & IdProt & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Protocolos (["
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
RsRegPendiente.Fields("Modulo").Value = "Protocolos"
RsRegPendiente.Fields("Tabla").Value = "Protocolos"
RsRegPendiente.Fields("Condicional").Value = "Id=" & IdProt & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub







Private Sub BtnSiguiente_Click()
If RsProtocolos.RecordCount <> 0 Then
    If RsProtocolos.EOF Then
        RsProtocolos.MoveFirst
    Else
        RsProtocolos.MoveNext
        If RsProtocolos.EOF Then RsProtocolos.MoveFirst
    End If
    Call Carga_De_Datos

    Else
    MsgBox "No se encontraron registros cargados, Inicie una nueva busqueda!", vbExclamation + vbOKOnly, "No hay Datos!"
End If
End Sub

Private Sub ChameleonBtn1_Click()
Unload Me
End Sub

Private Sub ChameleonBtn4_Click()
Blanqueo
BtnAgregar.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = False
BtnAnterior.Enabled = True
BtnSiguiente.Enabled = True
End Sub

Private Sub Form_Load()
Centrar Me
Cambio = 1
If RsProtocolos.State = 1 Then RsProtocolos.Close

CSql = "Select * From Protocolos Where Activo='1'"
Set RsProtocolos = CrearRS(CSql)

BtnAgregar.Enabled = True
BtnGuardar.Enabled = False
BtnBorrar.Enabled = False

End Sub

Sub CargarProtocolos()

CSql = "Select * From Protocolos where Activo='1'"
Set RsProtocolos = CrearRS(CSql)

FrmRadioTerapia.Combo3.Clear
Do While Not RsProtocolos.EOF
    
    FrmRadioTerapia.Combo3.AddItem Trim(RsProtocolos.Fields("Protocolo").Value)
    FrmRadioTerapia.Combo3.ItemData(FrmRadioTerapia.Combo3.NewIndex) = RsProtocolos.Fields("id").Value
    RsProtocolos.MoveNext
Loop

End Sub

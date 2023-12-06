VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEvolucion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evolución Clínica"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   Icon            =   "FrmEvolucion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin MSComctlLib.ListView ListView1 
         Height          =   6855
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12091
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Especialidad"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Evolucion del Paciente"
            Object.Width           =   13759
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Id"
            Object.Width           =   2
         EndProperty
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2760
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49938433
         CurrentDate     =   40322
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   7440
         Width           =   11295
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   8400
         Width           =   11295
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   10200
            TabIndex        =   2
            ToolTipText     =   "Cerrar"
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
            MICON           =   "FrmEvolucion.frx":1002
            PICN            =   "FrmEvolucion.frx":101E
            PICH            =   "FrmEvolucion.frx":11E7
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
            MICON           =   "FrmEvolucion.frx":141C
            PICN            =   "FrmEvolucion.frx":1438
            PICH            =   "FrmEvolucion.frx":16C7
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
            MICON           =   "FrmEvolucion.frx":1B08
            PICN            =   "FrmEvolucion.frx":1B24
            PICH            =   "FrmEvolucion.frx":1CB1
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
            Left            =   9000
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
            MICON           =   "FrmEvolucion.frx":1EE6
            PICN            =   "FrmEvolucion.frx":1F02
            PICH            =   "FrmEvolucion.frx":21E4
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
            Left            =   2520
            TabIndex        =   6
            ToolTipText     =   "Eliminar"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "FrmEvolucion.frx":2435
            PICN            =   "FrmEvolucion.frx":2451
            PICH            =   "FrmEvolucion.frx":25F5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnImprimir 
            Height          =   375
            Left            =   5640
            TabIndex        =   11
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Imprimir"
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
            MICON           =   "FrmEvolucion.frx":2794
            PICN            =   "FrmEvolucion.frx":27B0
            PICH            =   "FrmEvolucion.frx":28D5
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmEvolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NuevoId As Integer
Dim MaxRegP As Integer
Dim IdEvo As Integer
Dim RsTemp As ADODB.Recordset
Public IdPacE As Integer
Public IdLIdPacE As String

Private Sub BtnAgregar_Click()

If Val(IdPacE) = 0 Then
    Msg = "No hay Paciente Seleccionado para realizarle una evolución Clínica!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
Else
    DTPicker1.Enabled = True
    DTPicker1.Value = DateTime.Date
    Text1.Locked = False
    Text1.Text = ""
    Text1.Enabled = True
    Text1.SetFocus
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = False
    BtnImprimir.Enabled = False
    ListView1.Enabled = False
End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Text1.Text = ""
DTPicker1.Value = DateTime.Date
DTPicker1.Enabled = False
Text1.Enabled = False
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
ListView1.Enabled = True
End Sub

Private Sub BtnEliminar_Click()
On Error GoTo WrtError

Msg = "Esta seguro de borrar la evolucion clinica del paciente?"
mensaje = MsgBox(Msg, vbInformation + vbYesNo, "Mensaje de Borrado")

Dim RsBorrarEvolucion As New ADODB.Recordset

If mensaje = vbYes Then

    CSql = "Select * From EvolucionPaciente Where IdEvolucion = " & IdEvo & " And IdL = '" & IdLIdInf & "'"
    Set RsBorrarEvolucion = CrearRS(CSql)
    RsBorrarEvolucion.Delete

End If

Msg = "Evolución clinica del paciente Borrada satisfactoriamente!!!!"
mensaje = MsgBox(Msg, vbInformation + vbOKOnly, "Mensaje de Borrado")

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Historial Nutricional"

EnviarRegPendiente IdEvo, IdLIdInf

CargarGrid
Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error GoTo WrtError

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False

If Text1.Text = "" Then
    Msg = "Tiene que ingresar la evolución clinica del paciente!!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Text1.SetFocus
    Exit Sub
End If

Dim RsGuardarEvolucion As New ADODB.Recordset

CSql = "Select MAX(IdEvolucion)+1 as NuevoId From EvolucionPaciente"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = RsTemp.Fields("NuevoId").Value
Else
    NuevoId = "1"
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Verifica si hay conexion con el internet

If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

IdLIdInf = NuevoIdL

CSql = "Select * From EvolucionPaciente"
Set RsGuardarEvolucion = CrearRS(CSql)

RsGuardarEvolucion.AddNew
RsGuardarEvolucion.Fields("IdEvolucion").Value = NuevoId
RsGuardarEvolucion.Fields("IdL").Value = IdLIdInf
RsGuardarEvolucion.Fields("IdPaciente").Value = IdPacE
RsGuardarEvolucion.Fields("IdLIdPac").Value = IdLIdPacE
RsGuardarEvolucion.Fields("IdUser").Value = IdUser
RsGuardarEvolucion.Fields("Fecha").Value = DTPicker1.Value
RsGuardarEvolucion.Fields("EvolucionClinica").Value = Trim(Text1.Text)
RsGuardarEvolucion.Fields("Especialidad").Value = Trim(Especia)

RsGuardarEvolucion.Update
Msg = "Evolución del Paciente Guardada Satisfactoriamente!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Registro Guardado"
CargarGrid

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Historial Nutricional"

EnviarRegPendiente NuevoId, IdLIdInf

BtnDesHacer_Click

Exit Sub

WrtError:
Dim MError
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaInformes, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub

Sub EnviarRegPendiente(ByVal IdEvo2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If

CSql = "Select * From EvolucionPaciente Where IdEvolucion = " & IdEvo2 & " And IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then
    StrSen = "DELETE FROM EvolucionPaciente Where IdEvolucion = " & IdEvo2 & " And IdL = '" & IdLIdInf2 & "'"
Else
    
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    StrSen = "INSERT INTO EvolucionPaciente (["
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
End If

StrSen = Replace(StrSen, "'", "(varCSP)")

CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Evolución Clinica"
RsRegPendiente.Fields("Tabla").Value = "EvolucionPaciente"
RsRegPendiente.Fields("Condicional").Value = "IdEvolucion = " & IdEvo2 & " And IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

Private Sub BtnImprimir_Click()

'If IdPacE = "" Then MsgBox "Debe seleccionar un Paciente antes de agregar un registro!", vbExclamation + vbOKOnly, "Seleecione un Paciente": Exit Sub
'If Text1.Text = "" Then
'    MsgBox "Debe de seleccionar un Paciente", vbCritical + vbOKOnly, "Mensaje"
'    Exit Sub
'End If

''========= ESTE ES EL CODIGO NUEVO ==========

With CrystalReport1
    .ReportFileName = RutaInformes & "\EvolucionClinica.rpt"
    .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{EvolucionClinica.IdPaciente} = " & IdPacE
    .WindowTitle = "Evolución Clínica - Paciente: " & IdPacE
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
End With

End Sub

Private Sub Form_Load()

DTPicker1.Value = DateTime.Date
DTPicker1.Enabled = False
Text1.Locked = True

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False


If Val(IdPacE) = 0 Then
    Me.Caption = "Evolución Clinica "
Else
    Me.Caption = "Evolución Clinica - Paciente: " & IdPacE
    CargarGrid
End If

End Sub

Sub CargarGrid()

Dim RsCargar As New ADODB.Recordset

CSql = "Select * From EvolucionPaciente Where IdPaciente='" & IdPacE & "' Order by IdEvolucion"
Set RsCargar = CrearRS(CSql)

ListView1.ListItems.Clear

Do While Not RsCargar.EOF
    With ListView1
        i = i + 1
        .ListItems.Add , , Format(RsCargar.Fields("Fecha").Value, "dd/mm/yyyy")
        .ListItems(i).ListSubItems.Add , , Trim(RsCargar.Fields("Especialidad").Value)
        .ListItems(i).ListSubItems.Add , , Trim(RsCargar.Fields("EvolucionClinica").Value)
        .ListItems(i).ListSubItems.Add , , RsCargar.Fields("IdEvolucion").Value
    End With
    RsCargar.MoveNext
Loop
If ListView1.ListItems.Count > 0 Then
    BtnImprimir.Enabled = True
    BtnEliminar.Enabled = True
Else
    BtnImprimir.Enabled = False
    BtnEliminar.Enabled = False
End If

End Sub

Private Sub ListView1_Click()
Dim RsBuscar As New ADODB.Recordset

IdEvo = 0
IdLIdInf = IdLDefault

If ListView1.ListItems.Count <= 0 Then Exit Sub

CSql = "Select * From EvolucionPaciente Where IdPaciente='" & IdPacE & "' And IdEvolucion='" & ListView1.SelectedItem.ListSubItems(3).Text & "'"
Set RsBuscar = CrearRS(CSql)


If ListView1.ListItems.Count > 0 Then
Text1.Locked = True
Text1.Text = Trim(RsBuscar.Fields("EvolucionClinica").Value)
IdEvo = RsBuscar.Fields("IdEvolucion").Value
IdLIdInf = RsBuscar.Fields("IdL").Value
End If
End Sub

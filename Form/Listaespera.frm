VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListaEspera 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Espera"
   ClientHeight    =   7035
   ClientLeft      =   3495
   ClientTop       =   3150
   ClientWidth     =   11505
   Icon            =   "Listaespera.frx":0000
   LinkTopic       =   "Form52"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11505
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   10560
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Pacientes en Lista"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9975
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Historia"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cédula"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Apellido"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripción"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Hora Atención"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Hora Llegada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Motivo"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   6000
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   56819713
            CurrentDate     =   40204
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   2640
         TabIndex        =   1
         Top             =   6000
         Width           =   8535
         Begin ChamaleonButton.ChameleonBtn BtnLlamarPacientePantalla 
            Height          =   400
            Left            =   4800
            TabIndex        =   6
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Llamar Paciente en Pantalla"
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
            MICON           =   "Listaespera.frx":1002
            PICN            =   "Listaespera.frx":101E
            PICH            =   "Listaespera.frx":1253
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnQuitarPacienteLista 
            Height          =   400
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Quitar Paciente en Espera"
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
            MICON           =   "Listaespera.frx":14C4
            PICN            =   "Listaespera.frx":14E0
            PICH            =   "Listaespera.frx":1684
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnRefrescarLista 
            Height          =   400
            Left            =   2640
            TabIndex        =   3
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Refrescar  Lista "
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
            MICON           =   "Listaespera.frx":1823
            PICN            =   "Listaespera.frx":183F
            PICH            =   "Listaespera.frx":1AA9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   400
            Left            =   7320
            TabIndex        =   7
            ToolTipText     =   "Cerrar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
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
            MICON           =   "Listaespera.frx":1D29
            PICN            =   "Listaespera.frx":1D45
            PICH            =   "Listaespera.frx":1F0E
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
Attribute VB_Name = "FrmListaEspera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMaxId As New ADODB.Recordset
Dim BdLista As New ADODB.Recordset 'tabla de estatus
Dim bDlista1 As New ADODB.Recordset 'tabla paciente
Dim bdlista2 As New ADODB.Recordset 'tabla de llamado
Dim bdlista3 As New ADODB.Recordset 'tabla de llamado
Dim bdlista4 As New ADODB.Recordset 'tabla de llamado
Dim bdlista5 As New ADODB.Recordset 'tabla de llamado
Dim bdlista6 As New ADODB.Recordset
Dim bdlista7 As New ADODB.Recordset
Dim BdLista8 As New ADODB.Recordset
Dim RsTemp As ADODB.Recordset

Sub CargarDataGrid(dg As DataGrid)
On Error Resume Next

    dg.MarqueeStyle = dbgHighlightRow
    Set dg.DataSource = BdLista
    dg.Refresh
    
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnLlamarPacientePantalla_Click()
On Error Resume Next
If ListView1.ListItems.Count > 0 Then
Cedul = ListView1.SelectedItem.ListSubItems(1).Text

CSql = "Select * From Paciente Where CedulaP = " & Cedul
If bDlista1.State Then bDlista1.Close
Set bDlista1 = CrearRS(CSql)
    If Not bDlista1.EOF Then
        CSql = "Select * From Ubi_Paciente Where IdPaciente = " & bDlista1.Fields("idpaciente")
        Set BdLista8 = CrearRS(CSql)
        If Not BdLista8.EOF Then
            Select Case BdLista8.Fields("modul")
                Case Is = 0
                    m = "Nutrición"
                Case Is = 1
                    m = "Psicología"
                Case Is = 2
                    m = "Tratamiento de Radioterapia"
                Case Is = 3
                    m = "Dirección Médica"
                Case Is = 4
                    m = "Oncología Radioterapeuta"
                Case Is = 5
                    m = "Administración"
            End Select
            If BdLista8.Fields("modul") <> ModulO Then
                Msg = "El paciente está siendo atendido en " & m
                MsgBox Msg, vbOKOnly + vbCritical, "Paciente Ocupado"
                BdLista8.Close
                Exit Sub
            End If
        End If
    If BdLista8.State Then BdLista8.Close
    
    '######################################
    
    Dim RsCountLlamado1 As New ADODB.Recordset
    CSql = "Select max(IdLlamado) + 1 as MAxId from Llamado1"
    Set RsCountLlamado1 = CrearRS(CSql)
    C = RsCountLlamado1.Fields("MAxId")
    
    CSql = "Delete From Llamado1 Where Modulo = " & ModulO
    Set bdlista2 = CrearRS(CSql)
    DoEvents
    
    CSql = "Insert Into Llamado1(IdLlamado, IdPaciente, Modulo, Pantalla, MiniLlamador) values('" & C & "'," & bDlista1.Fields("idpaciente") & "," & ModulO & ", 0, 0)"
    Set bdlista6 = CrearRS(CSql)
    'DoEvents
    
    '######################################
    
    Dim RsCountLlamado2 As New ADODB.Recordset
    CSql = "Select max(IdLlamado)+ 1 as MAxId from Llamado2"
    Set RsCountLlamado2 = CrearRS(CSql)
    C = RsCountLlamado2.Fields("MAxId")
    
    CSql = "Delete From Llamado2 Where Modulo = " & ModulO
    Set bdlista2 = CrearRS(CSql)
    'DoEvents
    
    CSql = "Insert Into Llamado2(IdLlamado, IdPaciente, Modulo, Pantalla, MiniLlamador) values('" & C & "'," & bDlista1.Fields("idpaciente") & "," & ModulO & ", 0, 0)"
    Set bdlista7 = CrearRS(CSql)
    
   '######################################
    
    CSql = "Delete From Ubi_Paciente Where Modul = " & ModulO
    Set BdLista8 = CrearRS(CSql)
    
    CSql = "Insert Into Ubi_Paciente(Modul, IdPaciente) values(" & ModulO & "," & bDlista1.Fields("idpaciente") & ")"
    Set BdLista8 = CrearRS(CSql)

End If
 
bDlista1.Close
End If
End Sub

Private Sub BtnQuitarPacienteLista_Click()
On Error Resume Next

cedul1 = ListView1.SelectedItem.ListSubItems(1).Text

CSql = "Select * From Paciente Where CedulaP ='" & cedul1 & "'"
Set bdlista3 = CrearRS(CSql)
IdP = bdlista3.Fields("IdPaciente").Value
DoEvents
If Not bdlista3.EOF Then
    IdP = bdlista3.Fields("IdPaciente")
Else
    bdlista3.Close
    Exit Sub
End If
bdlista3.Close

' Obtienen el IDHistory
CSql = "Select MAX(Id_History) + 1 as MaxId From History_estatus"
Set RsMaxId = CrearRS(CSql)

If Not IsNull(RsMaxId.Fields("MaxId").Value) Then
    IdHistory = RsMaxId.Fields("MaxId").Value
Else
    IdHistory = "0"
End If
'mmmmmmmmmmmmmmmmmmmmmm

CSql = "Select * From Estatus Where IdPaciente = " & IdP & " And MotivoV = '" & ModulO & "'"
Set bdlista3 = CrearRS(CSql)
DoEvents

Mo = bdlista3.Fields("motivov")
fe = Format(bdlista3.Fields("fecha"), "dd/mm/yyyy")
idu = bdlista3.Fields("idusuario")
des = bdlista3.Fields("descripcion")
ho = DateTime.Time
bdlista3.Close
fe2 = Format(Now, "dd/mm/yyyy")
ho2 = DateTime.Time

CSql = "Insert into history_estatus(Id_History, IdPaciente, MotivoV, Fecha, Hora, IdUsuario, " & _
        "Descripcion, Fecha_Atendido, Hora_Atendido) values('" & IdHistory & "'," & IdP & ",'" & _
        ModulO & "','" & fe & "','" & ho & "'," & idu & ",'" & des & "','" & fe2 & "','" & ho2 & "')"

Set bdlista3 = CrearRS(CSql)
DoEvents

CSql = "Delete From Estatus Where IdPaciente = " & IdP & " And Motivov = '" & ModulO & "'"
Set bdlista4 = CrearRS(CSql)
DoEvents

CSql = "Delete From Llamado1 where IdPaciente = " & IdP & " And Modulo = " & ModulO
Set bdlista5 = CrearRS(CSql)
DoEvents

CSql = "Delete From Ubi_Paciente Where IdPaciente = " & IdP & " And Modul = " & ModulO
Dim bdlista88 As New ADODB.Recordset
Set bdlista88 = CrearRS(CSql)
MsgBox "El Paciente fue Borrado de la lista de espera!!", vbInformation + vbOKOnly, "Borrado"

Call Modul
On Error GoTo 0
End Sub

Private Sub BtnRefrescarLista_Click()
Modul
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Call Modul
End Sub

Private Sub Form_Load()
Centrar Me
DTPicker1.Value = Format(Now, "dd/mm/yyyy")
Call Modul
End Sub
Sub Modul()
On Error Resume Next
Select Case ModulO

Case Is = 0 'Nutricion
Me.Caption = "Lista de pacientes / Nutricion"
CSql = "Select *,CAST((substring(Grupo,1,5) + replace(substring(Grupo,6,8),'.','')) as datetime) AS Ordenar From EstatusPac Where MotivoV = '0' And Fecha = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "' Order By Ordenar"

Case Is = 1 'Psicologia
Me.Caption = "Lista de pacientes / Psicología"
CSql = "Select *,CAST((substring(Grupo,1,5) + replace(substring(Grupo,6,8),'.','')) as datetime) AS Ordenar From EstatusPac Where MotivoV = '1' And Fecha = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "' Order By Ordenar"

Case Is = 2 'Radiologo
Me.Caption = "Lista de pacientes / Radiologia"
CSql = "Select *,CAST((substring(Grupo,1,5) + replace(substring(Grupo,6,8),'.','')) as datetime) AS Ordenar From EstatusPac Where MotivoV = '2' And Fecha = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "' Order By Ordenar"

Case Is = 3 'Internista
Me.Caption = "Lista de pacientes / Dirección Médica"
CSql = "Select *,CAST((substring(Grupo,1,5) + replace(substring(Grupo,6,8),'.','')) as datetime) AS Ordenar From EstatusPac Where MotivoV = '3' And Fecha = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "' Order By Ordenar"

Case Is = 4 'Oncologia
Me.Caption = "Lista de pacientes / Oncologia"
CSql = "Select *,CAST((substring(Grupo,1,5) + replace(substring(Grupo,6,8),'.','')) as datetime) AS Ordenar From EstatusPac Where MotivoV = '4' And Fecha = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "' Order By Ordenar"

Case Is = 5 'Administracion
Me.Caption = "Lista de pacientes / Administración"
CSql = "Select *,CAST((substring(Grupo,1,5) + replace(substring(Grupo,6,8),'.','')) as datetime) AS Ordenar From EstatusPac Where MotivoV = '5' And Fecha = '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "' Order By Ordenar"

End Select

If BdLista.State Then BdLista.Close
Set BdLista = CrearRS(CSql)


ListView1.ListItems.Clear
If BdLista.EOF Then
    Call salirr
    Exit Sub
Else
    Dim ColorFila
    Hora1 = CDate(BdLista.Fields("Grupo").Value)
    Hora = CDate(BdLista.Fields("Hora").Value)
     
        Do While Not BdLista.EOF
        
            hora2 = CDate(CDate(BdLista.Fields("Grupo").Value) + CDate("00:10"))
            If CDate(hora2) >= CDate(BdLista.Fields("Hora").Value) Then
                With ListView1
                    i = i + 1
                    .ListItems.Add , , BdLista.Fields("Historia").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("CedulaP").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("NombreP").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("ApellidoP").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("Descripcion").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("Grupo").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("Hora").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("MotivoV").Value
                        
                    
                    If CDate(hora2) < CDate(Format(Now, "hh:mm:ss AMPM")) Then
                        ColorFila = vbBlue
                    Else
                        ColorFila = vbBlack
                    End If
                    
                    .ListItems(i).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(1).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(2).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(3).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(4).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(5).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(6).ForeColor = ColorFila
                    .ListItems(i).ListSubItems.Item(7).ForeColor = ColorFila
                                        
                End With
            Else
                With ListView1
                    i = i + 1
                    .ListItems.Add , , BdLista.Fields("Historia").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("CedulaP").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("NombreP").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("ApellidoP").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("Descripcion").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("Grupo").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("Hora").Value
                    .ListItems(i).ListSubItems.Add , , BdLista.Fields("MotivoV").Value
                
                    .ListItems(i).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(2).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(3).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(4).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(5).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(6).ForeColor = vbRed
                    .ListItems(i).ListSubItems.Item(7).ForeColor = vbRed
                
                End With
            
            End If
        BdLista.MoveNext
        Loop
        
    Frame1.Enabled = True
    BtnQuitarPacienteLista.Enabled = True
    BtnLlamarPacientePantalla.Enabled = True
End If


End Sub
Sub salirr()
Msg = "No hay Pacientes en espera!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "No hay Pacientes!!!"

    BtnQuitarPacienteLista.Enabled = True
    BtnLlamarPacientePantalla.Enabled = True
    Cedul = ""

    ListView1.ListItems.Clear
    
End Sub


Private Sub ListView1_DblClick()
If ListView1.ListItems.Count > 0 Then
    Cedul = ListView1.SelectedItem.ListSubItems(1).Text
    Select Case ModulO
        Case Is = 0 'Nutricion
            FrmHistorialNutricional.TxtBuscar.Text = Cedul
            FrmHistorialNutricional.BtnBuscar_Click
        Case Is = 1 'Psicologia Niños o Adolescentes
            Select Case Consultaa
                Case Is = "N"
                    FrmConsultaPsicologicaNoA.TxtBuscar = Cedul
                    FrmConsultaPsicologicaNoA.BtnBuscar_Click
                Case Is = "A" 'Psicologia Adultos
                    FrmConsultaPsicologicaAdult.TxtBuscar = Cedul
                    FrmConsultaPsicologicaAdult.BtnBuscar_Click
            End Select
        Case Is = 2 'Radiologo
            FrmRadioTerapia.TxtBuscar.Text = Cedul
            FrmRadioTerapia.BtnBuscar_Click
        Case Is = 3 'Internista
            FrmDireccionMedica.TxtBuscar.Text = Cedul
            FrmDireccionMedica.BtnBuscar_Click
        Case Is = 4 'Oncologia
            FrmRadioTerapeuta.TxtBuscar.Text = Cedul
            FrmRadioTerapeuta.BtnBuscar_Click
        Case Is = 5 'Administracion
    End Select
End If

Unload Me
End Sub

Private Sub Timer1_Timer()
BtnRefrescarLista_Click
End Sub

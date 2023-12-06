VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoMedicos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado del personal Médico"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "FrmListadoMedicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   4
         Top             =   6240
         Width           =   3255
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   1320
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   2160
            TabIndex        =   5
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
            MICON           =   "FrmListadoMedicos.frx":1002
            PICN            =   "FrmListadoMedicos.frx":101E
            PICH            =   "FrmListadoMedicos.frx":11E7
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   6240
         Width           =   3975
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad"
            Top             =   240
            Width           =   2415
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2640
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Buscar"
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
            MICON           =   "FrmListadoMedicos.frx":141C
            PICN            =   "FrmListadoMedicos.frx":1438
            PICH            =   "FrmListadoMedicos.frx":169D
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
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5895
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   10398
         Object.Width           =   7305
         Object.Height          =   5865
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoMedicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemp As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Medicos Where Apellido like '%" & Trim(TxtBuscar.Text) & "%' OR Nombre like '%" & Trim(TxtBuscar.Text) & "%' OR Cedula = '" & Val(Trim(TxtBuscar.Text)) & "' AND Activo='1' Order by IdPaciente"
Else
    CSql = "Select * From Medicos WHERE Activo='1' order by Nombre"
End If


Dim RsBuscarCliente As New ADODB.Recordset

Set RsBuscarCliente = CrearRS(CSql)

If RsBuscarCliente.RecordCount > 0 Then
    DMGrid1.Rows = 0
    Do While Not RsBuscarCliente.EOF
        
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarCliente.Fields("Cedula").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarCliente.Fields("Nombre").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarCliente.Fields("Apellido").Value
        
        If RsBuscarCliente.Fields("ApellidoP").Value = "1" Then
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Remitente"
        ElseIf RsBuscarCliente.Fields("ApellidoP").Value = "3" Then
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Rmtte / Trat."
        Else
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Tratante"
        End If
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsBuscarCliente.Fields("IdMedico").Value
        RsBuscarCliente.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "No existe esa referencia buscada", vbOKOnly, "Sin Resultado!"
    Exit Sub
End If
RsBuscarCliente.Close

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    
    Dim RsSeleccionarPaciente As New ADODB.Recordset
    
    FrmRegistroMedicoRemitente.BtnDesHacer_Click
    
    CSql = "Select * From Medicos WHERE IdMedico=" & DMGrid1.ValorCelda(lRow, 5) & " AND Activo='1' order by Nombre"
    Set RsSeleccionarPaciente = CrearRS(CSql)
    
    Select Case Tipo
        Case Is = "Nuevo Medico"
            Set FrmRegistroMedicoRemitente.RsRegMedico = CrearRS(CSql)
            FrmRegistroMedicoRemitente.Carga_De_Datos
    End Select
    
    RsSeleccionarPaciente.Close
    FrmRegistroMedicoRemitente.Label15.Caption = "Registro por selección"
Unload Me
End If

End Sub

Private Sub Form_Load()

Centrar Me
IniDMGrid
Dim RsCargarListaPaciente As New ADODB.Recordset

CSql = "Select * From Medicos WHERE Activo='1' order by Tipo,Nombre"
Set RsCargarListaPaciente = CrearRS(CSql)

Do While Not RsCargarListaPaciente.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListaPaciente.Fields("Cedula").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListaPaciente.Fields("Nombre").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListaPaciente.Fields("Apellido").Value
        
        If RsCargarListaPaciente.Fields("Tipo").Value = "1" Then
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Remitente"
        ElseIf RsCargarListaPaciente.Fields("Tipo").Value = "3" Then
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Rmtte / Trat."
        Else
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = "Tratante"
        End If
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsCargarListaPaciente.Fields("IdMedico").Value
    RsCargarListaPaciente.MoveNext
Loop
DMGrid1.PaintMGrid
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End If
End Sub


Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 5
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 0
DMGrid1.DColumnas(5).Alignment = 0
DMGrid1.DColumnas(5).Visible = False

DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(4).Locked = True
'DMGrid1.DColumnas(5).Locked = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 25 / 100) - 300
DMGrid1.DColumnas(1).Caption = "Cédula"
DMGrid1.DColumnas(2).Caption = "Nombre"
DMGrid1.DColumnas(3).Caption = "Apellido"
DMGrid1.DColumnas(4).Caption = "Tipo"
End Sub


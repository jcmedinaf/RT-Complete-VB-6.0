VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoPaciente 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pacientes"
   ClientHeight    =   7290
   ClientLeft      =   6075
   ClientTop       =   1770
   ClientWidth     =   8130
   Icon            =   "listapaciente.frx":0000
   LinkTopic       =   "Form32"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8130
   Begin VB.Frame FrameBusquedaAvanzada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Busqueda Avanzada"
      Height          =   3495
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox TxtNombre 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox TxtApellido 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtHistoria 
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox TxtCedula 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox TxtFechaReg 
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtFechaInicio 
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtFechaCulmi 
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DtpFechaReg 
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51118081
         CurrentDate     =   40464
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   375
         Left            =   5040
         TabIndex        =   18
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51118081
         CurrentDate     =   40464
      End
      Begin MSComCtl2.DTPicker DtpFechaCulmi 
         Height          =   375
         Left            =   5040
         TabIndex        =   19
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51118081
         CurrentDate     =   40464
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar2 
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         ToolTipText     =   "Realiza una Busqueda Avanzada"
         Top             =   2880
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
         MICON           =   "listapaciente.frx":1002
         PICN            =   "listapaciente.frx":101E
         PICH            =   "listapaciente.frx":1283
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar2 
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         ToolTipText     =   "Cerrar"
         Top             =   2880
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
         MICON           =   "listapaciente.frx":1515
         PICN            =   "listapaciente.frx":1531
         PICH            =   "listapaciente.frx":16FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   930
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1410
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Historia:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1890
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro:"
         Height          =   195
         Left            =   3600
         TabIndex        =   24
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio:"
         Height          =   195
         Left            =   3600
         TabIndex        =   23
         Top             =   930
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Culminación:"
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   1410
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4800
         TabIndex        =   4
         Top             =   6360
         Width           =   3015
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   120
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   1920
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
            MICON           =   "listapaciente.frx":192F
            PICN            =   "listapaciente.frx":194B
            PICH            =   "listapaciente.frx":1B14
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
         Top             =   6360
         Width           =   4575
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
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido,Cédula de identidad o Historia"
            Top             =   240
            Width           =   1695
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   1920
            TabIndex        =   2
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
            MICON           =   "listapaciente.frx":1D49
            PICN            =   "listapaciente.frx":1D65
            PICH            =   "listapaciente.frx":1FCA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnBusquedaAvanzada 
            Height          =   375
            Left            =   3240
            TabIndex        =   7
            ToolTipText     =   "Realiza una Busqueda Avanzada"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Avanzada"
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
            MICON           =   "listapaciente.frx":225C
            PICN            =   "listapaciente.frx":2278
            PICH            =   "listapaciente.frx":24DD
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
         Height          =   6015
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   10610
         Object.Width           =   7665
         Object.Height          =   5985
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim añ As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Paciente Where ApellidoP like '%" & Trim(TxtBuscar.Text) & "%' OR NombreP like '%" & Trim(TxtBuscar.Text) & "%' OR CedulaP = '" & Val(Trim(TxtBuscar.Text)) & "' OR Historia like '%" & Trim(TxtBuscar.Text) & "%' Order by IdPaciente"
Else
    CSql = "Select * From Paciente Order by IdPaciente"
End If


Dim RsBuscarCliente As New ADODB.Recordset

Set RsBuscarCliente = CrearRS(CSql)

If RsBuscarCliente.RecordCount > 0 Then
    DMGrid1.Rows = 0
    Do While Not RsBuscarCliente.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarCliente.Fields("IdPaciente").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarCliente.Fields("Historia").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarCliente.Fields("CedulaP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsBuscarCliente.Fields("ApellidoP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsBuscarCliente.Fields("NombreP").Value
        RsBuscarCliente.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly, "Sin Resultado"
    Exit Sub
End If
RsBuscarCliente.Close

End Sub

Private Sub BtnBuscar2_Click()
On Error GoTo Most
Dim RsBuscarCliente As New ADODB.Recordset
Dim wer As String


wher = ""

If txtCodigo.Text = "" And TxtNombre.Text = "" And TxtApellido.Text = "" And TxtHistoria.Text = "" And TxtCedula.Text = "" And TxtFechaReg.Text = "" And TxtFechaInicio.Text = "" And TxtFechaCulmi.Text = "" Then
    Msg = "Por favor ingrese Nombre o Apellido o cedula o No. Historia o Fecha de Registro " & Chr(13) & "o Fecha de Inicio del Tratamiento o Fecha de Culminación del Tratamiento del Paciente para realizar la busqueda!!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
Else

    If txtCodigo.Text <> "" Then
        wer = wer & " IdPaciente like '%" & txtCodigo.Text & "%'"
    End If

    If TxtNombre.Text <> "" Then
        If wer = "" Then
            wer = wer & " NombreP like '%" & TxtNombre.Text & "%'"
        Else
            wer = wer & " And NombreP like '%" & TxtNombre.Text & "%'"
        End If
    End If

    If TxtApellido.Text <> "" Then
        If wer = "" Then
            wer = wer & " ApellidoP like '%" & TxtApellido.Text & "%'"
        Else
            wer = wer & " And ApellidoP like '%" & TxtApellido.Text & "%'"
        End If
    End If

    If TxtHistoria.Text <> "" Then
        If wer = "" Then
            wer = wer & " Historia like '%" & TxtHistoria.Text & "%'"
        Else
            wer = wer & " And Historia like '%" & TxtHistoria.Text & "%'"
        End If
    End If
   
    If TxtCedula.Text <> "" Then
        If wer = "" Then
            wer = wer & " CedulaP like '%" & TxtCedula.Text & "%'"
        Else
            wer = wer & " And CedulaP like '%" & TxtCedula.Text & "%'"
        End If
    End If
    
    If TxtFechaReg.Text <> "" Then
        DtpFechaReg.Value = TxtFechaReg.Text
        If wer = "" Then
            wer = wer & " Fecha_RegP = '" & TxtFechaReg.Text & "'"
        Else
            wer = wer & " And Fecha_RegP = '" & TxtFechaReg.Text & "'"
        End If
    End If
    
    If TxtFechaInicio.Text <> "" Then
        DtpFechaInicio.Value = TxtFechaInicio.Text
        If wer = "" Then
            wer = wer & " Fecha_Inicio = '" & TxtFechaInicio.Text & "'"
        Else
            wer = wer & " And Fecha_Inicio ='" & TxtFechaInicio.Text & "'"
        End If
    End If

    If TxtFechaCulmi.Text <> "" Then
        DtpFechaCulmi.Value = TxtFechaCulmi.Text
        If wer = "" Then
            wer = wer & " Fecha_Culm ='" & TxtFechaCulmi.Text & "'"
        Else
            wer = wer & " And Fecha_Culm = '" & TxtFechaCulmi.Text & "'"
        End If
    End If

End If

CSql = "Select * From Paciente WHERE " & wer & " Order by IdPaciente"
Set RsBuscarCliente = CrearRS(CSql)

If RsBuscarCliente.RecordCount > 0 Then
    DMGrid1.Rows = 0
    Do While Not RsBuscarCliente.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarCliente.Fields("IdPaciente").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarCliente.Fields("Historia").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarCliente.Fields("CedulaP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsBuscarCliente.Fields("ApellidoP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsBuscarCliente.Fields("NombreP").Value
        RsBuscarCliente.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbcritial + vbOKOnly, "Sin Resultado"
    Exit Sub
End If
RsBuscarCliente.Close
FrameBusquedaAvanzada.Visible = False
Frame1.Enabled = True
Frame2.Enabled = True
Frame4.Enabled = True
Limpiar

Exit Sub
Most:
    MsgBox "La fecha que ha ingresado no es valida!", vbInformation + vbOKOnly, "Informacion"
    TxtFechaReg.Text = ""
    TxtFechaInicio.Text = ""
    TxtFechaCulmi.Text = ""
End Sub

Sub Limpiar()

txtCodigo.Text = ""
TxtNombre.Text = ""
TxtApellido.Text = ""
TxtHistoria.Text = ""
TxtCedula.Text = ""
TxtFechaReg.Text = ""
TxtFechaInicio.Text = ""
TxtFechaCulmi.Text = ""

End Sub

Private Sub BtnBusquedaAvanzada_Click()
FrameBusquedaAvanzada.Visible = True
Frame1.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = False
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()

DataGrid1.Col = 0
IdPac1 = Val(DataGrid1.Text)
Unload Me
End Sub

Private Sub BtnCerrar2_Click()
FrameBusquedaAvanzada.Visible = False
Frame1.Enabled = True
Frame2.Enabled = True
Frame4.Enabled = True
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    Dim RsSeleccionarPaciente As New ADODB.Recordset
    CSql = "Select * From Paciente Where IdPaciente='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsSeleccionarPaciente = CrearRS(CSql)
    
    Select Case Tipo
        Case Is = "Facturacion"
        'If RsSeleccionarPaciente.EOF Then
            IdPac1 = RsSeleccionarPaciente.Fields("IdPaciente").Value
            FacturacionRT.Text9.Text = RsSeleccionarPaciente.Fields("CedulaP").Value
            FacturacionRT.Text10.Text = RsSeleccionarPaciente.Fields("ApellidoP").Value
            FacturacionRT.Text11.Text = RsSeleccionarPaciente.Fields("NombreP").Value
            FacturacionRT.Text12.Text = RsSeleccionarPaciente.Fields("DireccionP").Value
            FacturacionRT.Text13.Text = RsSeleccionarPaciente.Fields("Codigo").Value & " " & RsSeleccionarPaciente.Fields("Telefono").Value
        'End If
        Case Is = "Planificacion"
            FrmPlanificacionPorPaciente.TxtBuscar.Text = RsSeleccionarPaciente.Fields("CedulaP").Value
            FrmPlanificacionPorPaciente.BtnBuscar_Click
        Case Is = "Consumo"
            FrmConsumoMedicamentos.TxtCedula.Text = RsSeleccionarPaciente.Fields("CedulaP").Value
            FrmConsumoMedicamentos.BtnBuscarPaciente_Click
        Case Is = "NuevoPaciente"
            FrmNuevoPaciente.TxtBuscar.Text = RsSeleccionarPaciente.Fields("CedulaP").Value
            FrmNuevoPaciente.BtnBuscar_Click
'    End Select
'
'    Select Case Tipo1
'        Case Is = "NuevoPaciente"
'            FrmNuevoPaciente.TxtBuscar.Text = RsSeleccionarPaciente.Fields("CedulaP").Value
'            FrmNuevoPaciente.BtnBuscar_Click
        Case Is = "HistorialPaciente"
            FrmHistorialMedico.TxtBuscar.Text = RsSeleccionarPaciente.Fields("CedulaP").Value
            FrmHistorialMedico.BtnBuscar_Click
    End Select
    RsSeleccionarPaciente.Close
Unload Me
End If

End Sub



Private Sub DtpFechaCulmi_Change()
TxtFechaCulmi.Text = Format(DtpFechaCulmi.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaInicio_Change()
TxtFechaInicio.Text = Format(DtpFechaInicio.Value, "dd/mm/yyyy")
End Sub

Private Sub DtpFechaReg_Change()
TxtFechaReg.Text = Format(DtpFechaReg.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()

Centrar Me
IniDMGrid
Dim RsCargarListaPaciente As New ADODB.Recordset

CSql = "Select * From Paciente order by IdPaciente"
Set RsCargarListaPaciente = CrearRS(CSql)

Do While Not RsCargarListaPaciente.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListaPaciente.Fields("IdPaciente").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListaPaciente.Fields("Historia").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListaPaciente.Fields("CedulaP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsCargarListaPaciente.Fields("ApellidoP").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsCargarListaPaciente.Fields("NombreP").Value
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
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
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
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(4).Locked = True
DMGrid1.DColumnas(5).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 30 / 100) - 300
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "No. Historia"
DMGrid1.DColumnas(3).Caption = "Cedula"
DMGrid1.DColumnas(4).Caption = "Apellidos(s)"
DMGrid1.DColumnas(5).Caption = "Nombres(s)"
End Sub

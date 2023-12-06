VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoEmpleados 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Empleados"
   ClientHeight    =   7395
   ClientLeft      =   7290
   ClientTop       =   795
   ClientWidth     =   7770
   Icon            =   "Listaempleado.frx":0000
   LinkTopic       =   "Form46"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7770
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   4
         Top             =   6480
         Width           =   3255
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   360
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
            MICON           =   "Listaempleado.frx":1002
            PICN            =   "Listaempleado.frx":101E
            PICH            =   "Listaempleado.frx":11E7
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
         Top             =   6480
         Width           =   3975
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2760
            TabIndex        =   2
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "Listaempleado.frx":141C
            PICN            =   "Listaempleado.frx":1438
            PICH            =   "Listaempleado.frx":169D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
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
            ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Usuario, Cédula de identidad o Código"
            Top             =   240
            Width           =   2535
         End
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   6135
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   10821
         Object.Width           =   7305
         Object.Height          =   6105
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pa As Integer
Dim RsBuscarEmpleado As New ADODB.Recordset
Dim RsTemp As Recordset
Private Sub Command1_Click()
Unload Me
End Sub


Private Sub BtnBuscar_Click()
On Error GoTo MostrarError:
If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Empleados Where (IdEmpleado = '" & Val(Trim(TxtBuscar.Text)) & "' OR Cedula = '" & Val(Trim(TxtBuscar.Text)) & "' OR Apellido like '%" & Trim(TxtBuscar.Text) & "%' OR Nombre like '%" & Trim(TxtBuscar.Text) & "%') And activo=1"
Else
   CSql = "Select * From Empleados where activo=1"
End If

Set RsBuscarEmpleado = CrearRS(CSql)

If RsBuscarEmpleado.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarEmpleado.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarEmpleado.Fields("IdEmpleado").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarEmpleado.Fields("Apellido").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsBuscarEmpleado.Fields("Nombre").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsBuscarEmpleado.Fields("Cedula").Value
        
        RsBuscarEmpleado.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly + vbCritical, "Sin Resultado"
    Exit Sub
End If
RsBuscarEmpleado.Close
Exit Sub
MostrarError:
MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_DobleClick()

If UCase(Tipo) = UCase("Nuevo Empleado") Then
    IdEmpl = Val(DMGrid1.ValorCelda(lRow, 1))
    Set RsTemp = FrmEmpleados.RsEmpleado
    
    If RsTemp.RecordCount <> 0 Then
        RsTemp.MoveFirst
        
        While Not RsTemp.EOF
            If RsTemp.Fields("IdEmpleado").Value = DMGrid1.ValorCelda(i, 1) Then
                FrmEmpleados.RsEmpleado.Find ("IdEmpleado=" & Val(DMGrid1.ValorCelda(i, 1)))
                FrmEmpleados.Empleado
                Unload FrmListadoEmpleados
                Exit Sub
            End If
            RsTemp.MoveNext
        Wend
    End If
    Unload Me
ElseIf UCase(Tipo) = UCase("Prestamos") Then
    
    IdEmpl = Val(DMGrid1.ValorCelda(lRow, 1))
    Set RsTemp = FrmPrestamos.RsEmpleados
    
    If RsTemp.RecordCount <> 0 Then
        RsTemp.MoveFirst
        
        While Not RsTemp.EOF
            If RsTemp.Fields("IdEmpleado").Value = DMGrid1.ValorCelda(i, 1) Then
                FrmPrestamos.RsEmpleados.Find ("IdEmpleado=" & Val(DMGrid1.ValorCelda(i, 1)))
                FrmPrestamos.CargaTrabajador
                FrmPrestamos.CargaPrest1
                Unload FrmListadoEmpleados
                Exit Sub
            End If
            RsTemp.MoveNext
        Wend
    End If
    Unload Me
ElseIf UCase(Tipo) = UCase("VCT") Then
    FrmValoresCampoTrabajador.CargarEmpleadoDeLista (Val(DMGrid1.ValorCelda(i, 1)))
    Unload Me
ElseIf Tipo = "Insumos" Then
    FrmSolicitudNecesidades.TxtCedula.Text = DMGrid1.ValorCelda(i, 4)
    FrmSolicitudNecesidades.BtnBuscarEmpleado_Click
    Unload Me
End If

End Sub
Private Sub Form_Load()
Centrar Me
IniDMGrid

CSql = "Select * From Empleados where activo=1"
Set RsCargarListaEmpleado = CrearRS(CSql)

Do While Not RsCargarListaEmpleado.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListaEmpleado.Fields("IdEmpleado").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListaEmpleado.Fields("Apellido").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListaEmpleado.Fields("Nombre").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsCargarListaEmpleado.Fields("Cedula").Value
    RsCargarListaEmpleado.MoveNext
Loop

DMGrid1.PaintMGrid

End Sub

Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 4
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(4).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 25 / 100)

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Apellido(s)"
DMGrid1.DColumnas(3).Caption = "Nombre(s)"
DMGrid1.DColumnas(4).Caption = "Cédula/Rif"
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

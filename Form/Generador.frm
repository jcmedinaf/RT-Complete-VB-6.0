VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmGeneradorNomina 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar de Nómina"
   ClientHeight    =   3075
   ClientLeft      =   6045
   ClientTop       =   2550
   ClientWidth     =   6675
   Icon            =   "Generador.frx":0000
   LinkTopic       =   "Form48"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6675
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   6495
      Begin ChamaleonButton.ChameleonBtn BtnGenerarNomina 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Generar Nómina"
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
         MICON           =   "Generador.frx":1002
         PICN            =   "Generador.frx":101E
         PICH            =   "Generador.frx":144D
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
         Height          =   375
         Left            =   5400
         TabIndex        =   9
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
         MICON           =   "Generador.frx":15CD
         PICN            =   "Generador.frx":15E9
         PICH            =   "Generador.frx":17B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrarNomina 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         ToolTipText     =   "Cerrar Tablas de Pacientes"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cerrar Nómina"
         ENAB            =   0   'False
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
         MICON           =   "Generador.frx":19E7
         PICN            =   "Generador.frx":1A03
         PICH            =   "Generador.frx":1CA5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Resultados"
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
         MICON           =   "Generador.frx":1F47
         PICN            =   "Generador.frx":1F63
         PICH            =   "Generador.frx":21F5
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
      Caption         =   "Generador de Nomina"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   3480
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   5160
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51183619
         CurrentDate     =   40017
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51183619
         CurrentDate     =   40017
      End
      Begin VB.Label LblPeriodo2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   4800
         TabIndex        =   15
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label LblPeriodo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label LblGeneradas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nóminas Generadas:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label LblProc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Nóminas:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label LblNroEmpl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Empleados:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   4080
         TabIndex        =   8
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Nomina"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmGeneradorNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BdGrupo As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim RsTemp2 As New ADODB.Recordset
Dim BdDatos As New ADODB.Recordset ' Consulta la tabla de empleados
Dim BdConce As New ADODB.Recordset ' Consulta los conceptos relacionados al grupo
Dim RECIBOS As New ADODB.Recordset ' PARA ENCABEZADOS DE RECIBOS
Dim RENGREC As New ADODB.Recordset ' PARA RENGLONES DE RECIBO
Dim concept(1 To 200)
Dim ArrayIdRecibos(30, 4000) As Integer
Dim ArraySQLPrestamos(0 To 4000, 0 To 1) As String

Function Verificar_Nomina() As Boolean

If Not (Combo1.ListIndex > -1) Then BtnCerrarNomina.Enabled = False: Exit Function

CSql = "SELECT * FROM Historico_Nomina WHERE Fecha_Ini_Nom='" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' AND Id_Grupo=" & Combo1.ItemData(Combo1.ListIndex)
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    Verificar_Nomina = True
    BtnCerrarNomina.Enabled = False
Else
    Verificar_Nomina = False
    BtnCerrarNomina.Enabled = True
End If
End Function

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnCerrarNomina_Click()
Dim resp
Dim IdGrupoTemp As Integer

If Not (Combo1.ListIndex > -1) Then MsgBox "No ha seleccionado un grupo de nómina", vbOKOnly: Exit Sub

IdGrupoTemp = Val(Combo1.ItemData(Combo1.ListIndex))
resp = MsgBox("Desea Cerrar la nómina para el grupo de " & Combo1.List(Combo1.ListIndex), vbQuestion + vbYesNo, "Confirmar")

If resp = 7 Then Exit Sub

CSql = "SELECT * FROM Recibos WHERE Id_Grupo=" & IdGrupoTemp
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then

    ' sentencia que almacena los recibos del grupo X, a la tabla de HISTORICOS DE RECIBOS
    CSql = "INSERT INTO Historico_Nomina SELECT * From Recibos WHERE Id_Grupo=" & IdGrupoTemp
    Set RsTemp = CrearRS(CSql)

    ' sentencia que almacena los renglones de los recibos del grupo X, a la tabla de HISTORICOS de REGLONES DE RECIBOS
    CSql = "INSERT INTO Historico_Reng_Nomina SELECT * From Reng_Recibo WHERE Id_Grupo=" & IdGrupoTemp
    Set RsTemp = CrearRS(CSql)

    ' Almacena la Fecha Actual de Pago
    CSql = "UPDATE Historico_Nomina SET Fecha_Pago=" & Format(Now, "dd/MM/yyyy")
    Set RsTemp = CrearRS(CSql)

    ' sentencia que elimina los recibos del grupo IdGrupoTemp
    CSql = "DELETE FROM Recibos WHERE Id_Grupo=" & IdGrupoTemp
    Set RsTemp = CrearRS(CSql)

    ' sentencia que elimina los renglones de los recibos del grupo IdGrupoTemp
    CSql = "DELETE FROM Reng_Recibo WHERE Id_Grupo=" & IdGrupoTemp
    Set RsTemp = CrearRS(CSql)

    CSql = "SELECT IdCampoNomina,Predeterminado FROM CamposDeNomina WHERE Inicializar=1"
    Set RsTemp = CrearRS(CSql)

    While Not RsTemp.EOF

        ' sentencia que INICIALIZA los campos de los trabajadores cuyos valores deben inicializarse
        ' a un valor predeterminado de la tabla de CamposDeNomina
        CSql = "UPDATE CamposDelTrabajador SET ValorN=" & RsTemp.Fields("Predeterminado").Value & " WHERE Tipo='CA' " & _
                " AND IdCampoNomina=" & RsTemp.Fields("IdCampoNomina").Value
        Set RsTemp2 = CrearRS(CSql)

        RsTemp.MoveNext
    Wend

    MsgBox "La nómina fue cerrada exitosamente!", vbExclamation + vbOKOnly, "Operación Exitosa!"
Else
    MsgBox "No se puede cerrar la nómina actual!, primero se debe generar", vbCritical + vbOKOnly, "No hay recibos de pago!"
    Exit Sub
End If

If Val(LblPeriodo2.Caption) = 24 Then
    MsgBox "Se genero y cerró la última nómina de pago del año.", vbInformation + vbOKOnly, "Información"
    MsgBox "Para generar las nóminas del año siguiente, debe configurarlo en la módulo de Grupos.", vbInformation + vbOKOnly + "Información"
Else
    If Val(LblPeriodo2.Caption) + 1 < 24 Then

        Call Calcular_Periodos(Format(DTPicker1.Value, "yyyy"))

        CSql = "UPDATE grupo SET fecha_prox_gen='" & PyFs((Val(LblPeriodo2.Caption) - 1) + 1, 1) & "'," & _
        " fecha_prox_gen2='" & PyFs((Val(LblPeriodo2.Caption) - 1) + 1, 2) & "',periodo=" & _
        PyFs((Val(LblPeriodo2.Caption) - 1) + 1, 0) & " WHERE id_grupo = " & Combo1.ItemData(Combo1.ListIndex)
        Set RsTemp = CrearRS(CSql)
    End If
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
For i = 0 To 4000
    If Not IsNull(ArraySQLPrestamos(i, 0)) Then
        ' Condicional que verifica si el arreglo "ArraySQLPrestamos" contiene información
        If Trim(ArraySQLPrestamos(i, 0)) <> "" Then
            ' Almacena la información (la cuales son sentencias SQL de actualizacion) en la variable
            ' CSQL para luego ejecutarlas y registrar los prestamos cobrados en la nomina
            CSql = Trim(ArraySQLPrestamos(i, 0))
            Set RsTemp = CrearRS(CSql)
            CSql = Trim(ArraySQLPrestamos(i, 1))
            Set RsTemp = CrearRS(CSql)
        Else
            Exit For
        End If
    Else
        Exit For
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Deshabilita los prestamos que ya han sido cancelados completamente...
CSql = "SELECT * FROM Prestamos"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    While Not RsTemp.EOF
        If CDbl(RsTemp.Fields("Adeuda")) = 0# Then
            CSql = "UPDATE Prestamos SET Activo = 0 WHERE IdPrestamos=" & RsTemp.Fields("IdPrestamos").Value
            Set RsTemp2 = CrearRS(CSql)
        End If
        RsTemp.MoveNext
    Wend
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

resp = MsgBox("Desea imprimir los recibos de Nómina?", vbInformation + vbYesNo, "Impresion de Recibos")
CrystalReport1.PrinterSelect

If resp = vbYes Then
    Dim Xy As Integer
    For Xy = 1 To 4000

        If Val(ArrayIdRecibos(Combo1.ListIndex, Xy)) = 0 Then Exit For
        ''========= ESTE ES EL CODIGO NUEVO ==========
       ' If IdEmpl = 0 Then Exit Sub
        'If IdReci = "" Then Exit Sub
       ' If NTabla <> 0 Then
            With CrystalReport1
                .ReportFileName = RutaInformes & "\Recibo_Pago1.rpt"
                '.Connect = "Data Source=Ing03;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
                '.Connect = "Data Source=Ing04;uid=sa;pwd=458921957JAr;DSQ=OAClinica;"
                '.Connect = "Data Source=Ing04;uid=sa;DSQ=OAClinica;"
                '.Connect = "Data Source=192.168.1.190;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
                .Connect = "DSN=CrReporte;"
                .DiscardSavedData = True
                .RetrieveDataFiles
                .ReportSource = 0
                '.SelectionFormula = "{ReciboDePago1.Fecha_Ini_Nom}=" & FechaSQL(DTPicker1.Value) & ""
                .SelectionFormula = "{ReciboDePago1.IdRecibos}=" & ArrayIdRecibos(Combo1.ListIndex, Xy) & ""
                '.SelectionFormula = "{ReciboDePago.IdEmpleado}=" & ArreEmpleados(Xy, 1) & " And {ReciboDePago.Fecha_Ini_Nom}=" & FechaSQL(DTPicker1.Value) & ""
                .ReportTitle = "Recibo de Pago"
                .Destination = crptToWindow
                .PrintFileType = crptCrystal
                .WindowState = crptMaximized
                .WindowMaxButton = False
                .WindowMinButton = False
                .Action = 1
            End With
        'End If
    Next Xy
End If

Combo1.ListIndex = -1

End Sub

Private Sub BtnGenerarNomina_Click()
On Error Resume Next

CSql = "SELECT MAX(Periodo) AS Period FROM Historico_Nomina"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    If Not (Val(RsTemp.Fields("Period").Value) + 1 = Val(LblPeriodo2.Caption)) Then
        MsgBox "No se ha cerrado la nómina anterior!", vbCritical + vbOKOnly, "NO SE PUEDE GENERAR LA NÓMINA"
        BtnGenerarNomina.Enabled = False
        BtnCerrarNomina.Enabled = False
        'ChameleonBtn1.Enabled = False
        Exit Sub
    End If
End If

GenerarNomina2

End Sub

Private Sub GenerarNomina2()
Dim CSQLTEMP As String
Dim BDRENG As New ADODB.Recordset
Dim NuevoId As Integer
Dim i As Integer
Dim CadConcepto As String
Dim CadCampo As String
Dim CadConstantes As String
Dim CadFunciones As String
Dim TReciboAsig As Double
Dim TReciboDedc As Double
Dim CantDelCampo As Double
Dim TOtros As Double
Dim Band1 As Boolean ' Activador de consulta, para obtener el último valor del recibo
Dim Band2 As Boolean ' Activador de consulta, para obtener el último valor del renglon del recibo


If Not (Combo1.ListIndex > -1) Then MsgBox "No ha seleccionado un grupo de nómina", vbOKOnly: Exit Sub
ContRecibos = 0
Band1 = False
Band2 = False
' Antes de Generar la Nómina, se deben BORRAR los recibos anteriores de esta TABLA llamada RECIBOS,
' ya que la nómina se puede generar N veces, si no se eliminan, habran Recibos Duplicados!
CSql = "DELETE FROM Recibos WHERE Id_Grupo=" & Combo1.ItemData(Combo1.ListIndex)
Set RsTemp = CrearRS(CSql)

CSql = "DELETE FROM Reng_Recibo WHERE Id_Grupo=" & Combo1.ItemData(Combo1.ListIndex)
Set RsTemp = CrearRS(CSql)

'consulta la tabla de empleados
CSql = "select * from empleados where (status = 1 or status = 0) and id_grupo = " & Combo1.ItemData(Combo1.ListIndex) & " and activo=1 order by IdEmpleado"
Set BdDatos = CrearRS(CSql)
'comprueba que la tabla no este vacia
If BdDatos.EOF Then
    Msg = "No hay integrantes activos en este grupo de nomina"
    MsgBox Msg, vbOKOnly, "no hay integrantes"
    Exit Sub
End If

'consulta los conceptos relacionados al grupo
BdDatos.MoveFirst
PB1.Max = BdDatos.RecordCount
PB1.Value = 0
PB1.Visible = True
While Not BdDatos.EOF
    IdEmpl = BdDatos.Fields("idempleado")
'    If IdEmpl = 24 Then
'       IdEmpl = IdEmpl
'    End If
    CSql = "SELECT CamposDelTrabajador.*, Concepto.Descripcion, Concepto.idconcepto FROM CamposDelTrabajador INNER JOIN " & _
        " Concepto ON CamposDelTrabajador.IdCampoNomina = Concepto.IdConcepto WHERE (CamposDelTrabajador.Tipo = 'CO') " & _
        " AND (Concepto.Activo = 1) AND CamposDelTrabajador.IdEmpleado=" & IdEmpl
    Set BdGrupo = CrearRS(CSql)
    
    'Obtiene un Id Nuevo
    If Band1 = False Then
        CSql = "SELECT MAX(IdHistorico)+1 as NuevoId FROM Historico_Nomina"
        Band1 = True
        Set RENGREC = CrearRS(CSql)
    
        If Not IsNull(RENGREC.Fields("NuevoId")) Then
            idrec = RENGREC.Fields("NuevoId")
        Else
            idrec = "1"
        End If
        
        'MMMMMMMMMMMMMM  normaliza el Id nuevo para Recibos MMMMMMMMMMMMMMMMMM
        CSQLTEMP = "SELECT IdRecibos FROM recibos"
        Set RsTemp = CrearRS(CSQLTEMP)

        If RsTemp.RecordCount <> 0 Then
            CSQLTEMP = "SELECT * FROM recibos ORDER BY IdRecibos"
            Set RsTemp = CrearRS(CSQLTEMP)
            i = idrec
            While Not RsTemp.EOF
            
                CSQLTEMP = "UPDATE recibos SET IdRecibos=" & i & " WHERE IdRecibos=" & RsTemp.Fields("IdRecibos").Value
                Set RsTemp2 = CrearRS(CSQLTEMP)
                
                CSQLTEMP = "UPDATE reng_recibo SET IdRecibos=" & i & " WHERE IdRecibos=" & RsTemp.Fields("IdRecibos").Value
                Set RsTemp2 = CrearRS(CSQLTEMP)
                
                i = i + 1
                RsTemp.MoveNext
            Wend
            CSql = "SELECT MAX(IdRecibos)+1 as NuevoId FROM recibos"
        End If
        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

    Else
        CSql = "SELECT MAX(idrecibos)+1 as NuevoId FROM recibos"
    End If
    
    Set RENGREC = CrearRS(CSql)
    If Not IsNull(RENGREC.Fields("NuevoId")) Then
        idrec = RENGREC.Fields("NuevoId")
    Else
        idrec = "1"
    End If

    If BdGrupo.RecordCount <> 0 Then
        If IsNull(BdGrupo.Fields(0).Value) Then
            Msg = "El Empleado " & BdDatos.Fields("Nombre") & " " & BdDatos.Fields("Apellido") & " de ID = " & BdDatos.Fields("IdEmpleado") & _
                " no contiene conceptos dentro de la base de datos, se continuara omitiendo este registro."
            MsgBox Msg, vbExclamation + vbOKOnly, "El Empleado no tiene conceptos generados"
            'Exit Sub
        Else
            BdGrupo.MoveFirst
            'Inicia la generación de conceptos uno a uno del empleado
            Do While Not BdGrupo.EOF
                IdEmpl = BdGrupo.Fields("idempleado")
                   
                '<<<<<<<<<AQUI COLOCAR LINEAS PARA EL REGISTRO DEL ENCABEZADO DEL RECIBO DE PAGO>>>>>>
                '<<<<<aqui>>>>
              
                CSql = "select * from concepto where idconcepto = " & BdGrupo.Fields("idconcepto")
                Set BdConce = CrearRS(CSql)
                If Not BdConce.EOF Then
                    CadConcepto = BdConce.Fields("Formula")
                    
                    ' Sentencia que obtiene el VALOR del campo anidado al concepto del trabajador
                    CSql = "select * from CamposDeltrabajador  where IdCampoNomina = " & BdConce.Fields("IdCampoAnidado") & _
                    " AND IdEmpleado=" & IdEmpl & " AND Tipo='CA'"
                    Set RsTemp = CrearRS(CSql)
                    
                    ' Condicional que verifica si encontro el campo, en el caso que lo encuentre, lo almacena
                    ' en la variable "CantDelCampo", de lo contrario la variable "CantDelCampo" será CERO "0"
                    If RsTemp.RecordCount <> 0 Then CantDelCampo = CDbl(RsTemp.Fields("ValorN").Value) Else CantDelCampo = 0
                Else
                    CantDelCampo = 0
                End If
                
                UI9 = 0
HB:
                resultado = 0
                T = InStr(1, UCase(CadConcepto), "CONCEPTO", vbTextCompare)
                If T <> 0 Then
                    If UI9 >= 100 Then
                        Msg = "Problema reciclico en la formula, esta haciendo mencion a un concepto que a su vez esta refiriendo a otro que depende del primero, corrija y vuelva a intentar"
                        MsgBox Msg
                        Exit Sub
                    Else
                        'Call verifica_concepto
                        CadConcepto = Validar_Concepto(CadConcepto)
                    End If
                    UI9 = UI9 + 1
                    GoTo HB
                End If
                
                'Call Verifica_campo
                CadCampo = Validar_Campo(CadConcepto)
                'Call verifica_constante
                CadFunciones = Validar_Funcion(CadCampo, Val(LblPeriodo2.Caption), Format(DTPicker1.Value, "dd/MM/yyyy"))
                CadConstantes = Validar_Constante(CadFunciones)
                CadConstantes = Replace(CadConstantes, ",", ".")
                CadSSO = CadConstantes
                CadConstantes = Validar_SSO(CadConstantes)
                
                resultado = ScriptControl1.Eval(CadConstantes)
    
                If CadSSO <> CadConstantes Then
                    resultado = Calcular_SSO(CadConstantes, resultado, IdEmpl)
                End If
                
                'resultado = ScriptControl1.Eval(FormulA)
                '<<<<< AQUI COLOCAR LINEAS PARA EL REGISTRO EN LA TABLA DE RENGLONES DEL RECIBO DE PAGO
                ' CON EL RESULTADO DEL CONCEPTO >>>>>>>
                
                If Val(resultado) <> 0 Then
                    
                    If Band2 = False Then
                        CSql = "SELECT MAX(IdHistorico)+1 as NuevoId FROM Historico_Reng_Nomina"
                        Band2 = True
                        Set RENGREC = CrearRS(CSql)
                    
                        If Not IsNull(RENGREC.Fields("NuevoId")) Then
                            NuevoId = RENGREC.Fields("NuevoId")
                        Else
                            NuevoId = "1"
                        End If
                        
                        'MMMMMMMMMMMMMM  normaliza el Id nuevo para Renglones del recibo MMMMMMMMMMMMMMMMMM
                        CSQLTEMP = "SELECT IdRengRec FROM reng_recibo"
                        Set RsTemp = CrearRS(CSQLTEMP)
                
                        If RsTemp.RecordCount <> 0 Then
                            CSQLTEMP = "SELECT * FROM reng_recibo ORDER BY IdRengRec"
                            Set RsTemp = CrearRS(CSQLTEMP)
                            i = NuevoId
                            While Not RsTemp.EOF
                                RsTemp.Fields("IdRengRec").Value = i
                                i = i + 1
                                RsTemp.Update
                                RsTemp.MoveNext
                            Wend
                            CSql = "SELECT MAX(IdRengRec)+1 as NuevoId FROM reng_recibo"
                        End If
                        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

                    Else
                        CSql = "SELECT MAX(idrengrec)+1 as NuevoId FROM reng_recibo"
                    End If

                    Set BDRENG = CrearRS(CSql)
                    If Not IsNull(BDRENG.Fields("NuevoId").Value) Then
                        NuevoId = BDRENG.Fields("NuevoId").Value
                    Else
                        NuevoId = "1"
                    End If
                    
                    CSql = "insert into reng_recibo(idrengrec,idrecibos,idconcepto,valorn,fecha_gen,iduser,id_grupo,Detalle,Cantidad) " & _
                          " values(" & NuevoId & "," & Val(idrec) & "," & Val(BdGrupo.Fields("idconcepto")) & "," & Replace(CDbl(resultado), ",", ".") & ",'" & Format(DTPicker1.Value, "dd/MM/yyyy") & _
                          "'," & Val(IdUser) & "," & Val(Combo1.ItemData(Combo1.ListIndex)) & ",'" & _
                          BdConce.Fields("Descripcion").Value & "'," & Replace(Replace(CantDelCampo, ".", ""), ",", ".") & ")"
                    Set BDRENG = CrearRS(CSql)
                    
                    If Val(BdConce.Fields("Tipo").Value) = 0 Then     ' Si es tipo CERO es ASIGNACIÓN
                        TReciboAsig = TReciboAsig + CDbl(resultado)
                    ElseIf Val(BdConce.Fields("Tipo").Value) = 1 Then ' Si es tipo UNO es DEDUCCION
                        TReciboDedc = TReciboDedc + CDbl(resultado)
                    Else                                              ' Si no es CERO o UNO, es RETENCIÓN u OTROS
                        TOtros = TOtros + CDbl(resultado)
                    End If
                End If
                '<<<<<<<<<<<>>>>>>>>>>>
                BdGrupo.MoveNext
            Loop
        End If
    End If
    If CDbl(TReciboAsig) > 0 Then
        ' Finaliza el RECIBO colocando el MONTO TOTAL que seria ==> TReciboAsig - TReciboDedc
        ' Pero en la base de datos se almacena por separado... TReciboAsig, TReciboDedc y NO SU RESTA
        Dim tem2 As Integer
        
        tem2 = Val(LblPeriodo2.Caption)
        CSql = "insert into recibos(idrecibos,idempleado, fecha_ini_nom, fecha_fin_nom,Total_Retenciones,Total_Deducciones," & _
            "Total_Asignacion, id_grupo,Periodo) values(" & idrec & "," & IdEmpl & ",'" & Format(DTPicker1.Value, "dd/mm/yyyy") & _
            "','" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'," & Replace(CDbl(TOtros), ",", ".") & "," & Replace(CDbl(TReciboDedc), ",", ".") & _
            "," & Replace(CDbl(TReciboAsig), ",", ".") & "," & Combo1.ItemData(Combo1.ListIndex) & "," & tem2 & ")"
        Set RENGREC = CrearRS(CSql)
        ContRecibos = ContRecibos + 1
        ArrayIdRecibos(Combo1.ListIndex, ContRecibos) = idrec
    End If
    TReciboAsig = 0
    TReciboDedc = 0
    TOtros = 0
    PB1.Value = PB1.Value + 1
    BdDatos.MoveNext
Wend

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMM Verificar Prestamos... MMMMMMMMMMMMMMMMMMMMM
    Call Verificar_Prestamos(BdDatos)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

PB1.Visible = False
Msg = "Nómina Generada exitosamente"
MsgBox Msg, vbOKOnly, "Generada la nómina"
BtnCerrarNomina.Enabled = True
Combo1.ListIndex = -1
End Sub

Sub Verificar_Prestamos(BdEmpleados As Recordset)
Dim NuevoId As Integer
Dim IdRefer As Integer
Dim i As Integer
Dim TempIdEmpl As Integer
Dim SaldoPrestamo As Double
Dim SaldoAbonoPrestamo As Double
Dim Band As Boolean

    BdEmpleados.MoveFirst
    
    i = 0
    While Not BdEmpleados.EOF
    
        TempIdEmpl = Val(BdEmpleados.Fields("IdEmpleado").Value)
        Band = False
        SaldoPrestamo = 0
        SaldoAbonoPrestamo = 0
        
        CSql = "SELECT * FROM Prestamos WHERE IdEmpleado=" & TempIdEmpl & " AND Activo='1'"
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
        
            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            ' si tiene prestamos verificar cual es el mayor... para empezar a cobrar
            Dim MontoMayor As Double
            Dim MontoSaldo As Double
            Dim NCuota As Integer
            Dim NCuotas As Integer
            Dim IdMontoMayor As Integer
            MontoMayor = 0
            While Not RsTemp.EOF
                If CDbl(RsTemp.Fields("Monto_Presta").Value) > MontoMayor Then
                    IdMontoMayor = Val(RsTemp.Fields("IdPrestamos").Value)
                    MontoMayor = CDbl(RsTemp.Fields("Monto_Presta").Value)
                    MontoSaldo = CDbl(RsTemp.Fields("Adeuda").Value)
                    SaldoPrestamo = MontoMayor
                    SaldoAbonoPrestamo = RsTemp.Fields("Abonos").Value
                End If
                RsTemp.MoveNext
            Wend
            
            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            
            CSql = "SELECT * FROM RenglonPrestamos WHERE IdEmpleado=" & TempIdEmpl & " AND IdPrestamo=" & IdMontoMayor & " ORDER BY IdRengPrestamo"
            Set RsTemp = CrearRS(CSql)
            
            If RsTemp.RecordCount <> 0 Then
                
                While Not RsTemp.EOF
                    ' Condicional que pregunta si la fecha del reglon del prestamo es igual a la de la generación
                    ' de la nomina... de ser asi pasa al siguiente condicional
                    If Format(RsTemp.Fields("FechaPago").Value, "dd/MM/yyyy") = Format(DTPicker2.Value, "dd/MM/yyyy") Then
                        ' condicional que pregunta si el pago de esa cuota se realizo por completo
                        MontoMayor = (CDbl(RsTemp.Fields("AbonoMax").Value) - CDbl(RsTemp.Fields("MontoAbono").Value))
                        If MontoMayor <> 0# Then
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            CSql = "SELECT Count(*) FROM RenglonPrestamos WHERE MontoAbono <> 0 AND IdPrestamo=" & IdMontoMayor
                            Set RsTemp2 = CrearRS(CSql)
                            NCuota = Val(RsTemp2.Fields(0).Value) + 1
                            NCuotas = Val(RsTemp.RecordCount)
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            ' Registrar el cobro del prestamo
                            CSql = "SELECT MAX(idrengrec)+1 as NuevoId FROM reng_recibo"
                            Set RsTemp2 = CrearRS(CSql)
                            If Not IsNull(RsTemp2.Fields("NuevoId").Value) Then
                                NuevoId = RsTemp2.Fields("NuevoId").Value
                            Else
                                NuevoId = "1"
                            End If
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            CSql = "SELECT IdRecibos FROM Recibos WHERE IdEmpleado=" & TempIdEmpl
                            Set RsTemp2 = CrearRS(CSql)
                             IdRefer = Val(RsTemp2.Fields("IdRecibos").Value)
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            ' Consulta para saber el total de Asignaciones y el total de Deducciones..
                            CSql = "SELECT Total_Deducciones, Total_Asignacion FROM Recibos WHERE IdRecibos=" & IdRefer
                            Set RsTemp2 = CrearRS(CSql)
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            
                            If CDbl(RsTemp2.Fields("Total_Asignacion").Value) - (CDbl(RsTemp2.Fields("Total_Deducciones").Value) + MontoMayor) >= 0# Then
                                ' Agrega el concepto del cobro del prestamo al renglon de recibos...
                                CSql = "insert into reng_recibo(idrengrec,idrecibos,idconcepto,valorn,fecha_gen,iduser,id_grupo,Detalle,Cantidad) " & _
                                      " values(" & NuevoId & "," & Val(IdRefer) & ",0," & Replace(CDbl(MontoMayor), ",", ".") & ",'" & Format(DTPicker1.Value, "dd/MM/yyyy") & _
                                      "'," & Val(IdUser) & "," & Val(Combo1.ItemData(Combo1.ListIndex)) & ",'PRESTAMO CUOTA " & NCuota & "/" & _
                                      NCuotas & "'," & Replace(Replace(MontoSaldo - MontoMayor, ".", ""), ",", ".") & ")"
                                Set RsTemp2 = CrearRS(CSql)
                                
                                ' Actualiza el Encabezado del recibo
                                CSql = "UPDATE Recibos SET Total_Deducciones=Total_Deducciones+" & Replace(Replace(CDbl(MontoMayor), ".", ""), ",", ".") & " WHERE IdRecibos=" & IdRefer
                                Set RsTemp2 = CrearRS(CSql)
                                
                                ' Cuando se cierre la nomina se actualizan los renglones de los prestamos.
                                '   Para ello se guarda la sentencia SQL de actualizacion de los renglones
                                ' de los prestamos para luego del cierre de nomina se proceda a registraslas
                                ArraySQLPrestamos(i, 0) = "UPDATE RenglonPrestamos SET MontoAbono=" & Replace(CDbl(MontoMayor), ",", ".") & _
                                    ",FechaAbono='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "' WHERE IdRengPrestamo=" & RsTemp.Fields("IdRengPrestamo")
                                ArraySQLPrestamos(i, 1) = "UPDATE Prestamos set Adeuda=Adeuda-" & _
                                    Replace(CDbl(MontoMayor), ",", ".") & ", Abonos=Abonos+" & _
                                    Replace(CDbl(MontoMayor), ",", ".") & " WHERE IdEmpleado=" & TempIdEmpl & " AND IdPrestamos=" & IdMontoMayor
    
                                ' Incrementa el contador "i" para dar lugar a una posicion nueva para el siguiente registro
                                i = i + 1
                                
                                ' Se activa la bandera indicando que se realizó un cobro a travez del prestamo
                                Band = True
                                
                                ' el siguiente comando va al FINAL de registro y finaliza la busqueda de cobros de prestamos
                                ' para que de esta manera NO SIGA buscando mas cobros...
                                RsTemp.MoveLast
                            End If
                        End If
                    End If
                    RsTemp.MoveNext
                Wend
                
                ' Si la BAND es false es porq no se han realizado cobros, por lo tanto debo verificar
                ' si existen otros prestamos no cancelados que coincidan con la fecha de la generacion
                ' de la nomina, y de ser asi, realiza el cobro del mismo...
                If Band = False Then
                
                    ' Consulta para buscar un cobro de un prestamo para la fecha de generacion de la nomina...
                    CSql = "SELECT RenglonPrestamos.*, Prestamos.Monto_Presta, Prestamos.Adeuda FROM RenglonPrestamos INNER JOIN Prestamos ON RenglonPrestamos.IdPrestamo = Prestamos.IdPrestamos " & _
                        " WHERE (RenglonPrestamos.FechaPago = '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "') AND Prestamos.IdEmpleado=" & TempIdEmpl & " ORDER BY Prestamos.IdPrestamos"
                    Set RsTemp = CrearRS(CSql)
                    
                    ' Si la consulta anterior contiene mas de 1 registro, entonces elige el prestamo mayor...
                    If RsTemp.RecordCount > 1 Then
                    
                        MontoMayor = 0
                        While Not RsTemp.EOF
                            If CDbl(RsTemp.Fields("Monto_Presta").Value) > CantPrestamo Then
                                MontoMayor = CDbl(RsTemp.Fields("Monto_Presta").Value)
                                IdMontoMayor = Val(RsTemp.Fields("IdPrestamo").Value)
                            End If
                            RsTemp.MoveNext
                        Wend
                        
                        CSql = "SELECT * FROM RenglonPrestamos WHERE IdPrestamo=" & IdMontoMayor & " AND FechaPago='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                        Set RsTemp = CrearRS(CSql)
                        
                        MontoMayor = (CDbl(RsTemp.Fields("AbonoMax").Value) - CDbl(RsTemp.Fields("MontoAbono").Value))
                        MontoSaldo = CDbl(RsTemp.Fields("Adeuda").Value)
                    ' Si no, entonces verifica si no encontro algun cobro, de ser asi finaliza la busqueda
                    ElseIf RsTemp.RecordCount = 0 Then
                        Band = True
                    Else
                        IdMontoMayor = Val(RsTemp.Fields("IdPrestamo").Value)
                        MontoMayor = (CDbl(RsTemp.Fields("AbonoMax").Value) - CDbl(RsTemp.Fields("MontoAbono").Value))
                        MontoSaldo = CDbl(RsTemp.Fields("Adeuda").Value)
                    End If
                    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                    
                        If MontoMayor <> 0# Then
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            CSql = "SELECT Count(*) FROM RenglonPrestamos WHERE MontoAbono <> 0 AND IdPrestamo=" & IdMontoMayor
                            Set RsTemp2 = CrearRS(CSql)
                            NCuota = Val(RsTemp2.Fields(0).Value) + 1
                            
                            CSql = "SELECT Count(*) FROM RenglonPrestamos WHERE IdPrestamo=" & IdMontoMayor
                            Set RsTemp2 = CrearRS(CSql)
                            NCuotas = Val(RsTemp2.Fields(0).Value)
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            ' Registrar el cobro del prestamo
                            CSql = "SELECT MAX(idrengrec)+1 as NuevoId FROM reng_recibo"
                            Set RsTemp2 = CrearRS(CSql)
                            If Not IsNull(RsTemp2.Fields("NuevoId").Value) Then
                                NuevoId = RsTemp2.Fields("NuevoId").Value
                            Else
                                NuevoId = "1"
                            End If
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            CSql = "SELECT IdRecibos FROM Recibos WHERE IdEmpleado=" & TempIdEmpl
                            Set RsTemp2 = CrearRS(CSql)
                             IdRefer = Val(RsTemp2.Fields("IdRecibos").Value)
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            ' Consulta para saber el total de Asignaciones y el total de Deducciones..
                            CSql = "SELECT Total_Deducciones, Total_Asignacion FROM Recibos WHERE IdRecibos=" & IdRefer
                            Set RsTemp2 = CrearRS(CSql)
                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            
                            If CDbl(RsTemp2.Fields("Total_Asignacion").Value) - (CDbl(RsTemp2.Fields("Total_Deducciones").Value) + MontoMayor) >= 0# Then
                                ' Agrega el concepto del cobro del prestamo al renglon de recibos...
                                CSql = "insert into reng_recibo(idrengrec,idrecibos,idconcepto,valorn,fecha_gen,iduser,id_grupo,Detalle,Cantidad) " & _
                                      " values(" & NuevoId & "," & Val(IdRefer) & ",0," & Replace(CDbl(MontoMayor), ",", ".") & ",'" & Format(DTPicker1.Value, "dd/MM/yyyy") & _
                                      "'," & Val(IdUser) & "," & Val(Combo1.ItemData(Combo1.ListIndex)) & ",'PRESTAMO CUOTA " & NCuota & "/" & _
                                      NCuotas & "'," & Replace(Replace(MontoSaldo - MontoMayor, ".", ""), ",", ".") & ")"
                                Set RsTemp2 = CrearRS(CSql)
                                
                                ' Actualiza el Encabezado del recibo
                                CSql = "UPDATE Recibos SET Total_Deducciones=Total_Deducciones+" & Replace(Replace(CDbl(MontoMayor), ".", ""), ",", ".") & " WHERE IdRecibos=" & IdRefer
                                Set RsTemp2 = CrearRS(CSql)
                                
                                ' Cuando se cierre la nomina se actualizan los renglones de los prestamos.
                                '   Para ello se guarda la sentencia SQL de actualizacion de los renglones
                                ' de los prestamos para luego del cierre de nomina se proceda a registraslas
                                ArraySQLPrestamos(i, 0) = "UPDATE RenglonPrestamos SET MontoAbono=" & Replace(CDbl(MontoMayor), ",", ".") & _
                                    ",FechaAbono='" & Format(DTPicker2.Value, "dd/MM/yyyy") & "' WHERE IdRengPrestamo=" & RsTemp.Fields("IdRengPrestamo")
                                ArraySQLPrestamos(i, 1) = "UPDATE Prestamos set Adeuda=Adeuda-" & _
                                    Replace(CDbl(MontoMayor), ",", ".") & ", Abonos=Abonos+" & _
                                    Replace(CDbl(MontoMayor), ",", ".") & " WHERE IdEmpleado=" & TempIdEmpl & " AND IdPrestamos=" & IdMontoMayor
    
                                ' Incrementa el contador "i" para dar lugar a una posicion nueva para el siguiente registro
                                i = i + 1
                                
                                ' Se activa la bandera indicando que se realizó un cobro a travez del prestamo
                                Band = True
                                
                                ' el siguiente comando va al FINAL de registro y finaliza la busqueda de cobros de prestamos
                                ' para que de esta manera NO SIGA buscando mas cobros...
                                RsTemp.MoveLast
                            End If
                        End If
                End If
            End If
        End If
        
        If Band = False Then
            If SaldoPrestamo <> SaldoAbonoPrestamo Then
                MsgBox "Se omitio el cobro del Prestamo Nro. (" & IdMontoMayor & ")  del Empleado: " & _
                    Chr(13) & Chr(13) & BdEmpleados.Fields("Nombre").Value & ", " & BdEmpleados.Fields("Apellido").Value & Chr(13) & _
                    "   Cédula: " & BdEmpleados.Fields("cedula").Value & Chr(13) & Chr(13) & _
                    "Para ver el detalle verifique el módulo de prestamos." & Chr(13), vbInformation + vbOKOnly, "Información."
            End If
        End If
        
        BdEmpleados.MoveNext
    Wend
End Sub

Private Sub ChameleonBtn1_Click()
    FrmHistoricoRecibos.Show vbModal, FrmPrincipal
End Sub

Private Sub Combo1_Click()
Dim VarTemp As Integer
Dim Band As Boolean

Band = False

If Combo1.ListIndex = -1 Then Exit Sub

CSql = "select * from grupo where id_grupo = " & Combo1.ItemData(Combo1.ListIndex)
Set BdGrupo = CrearRS(CSql)

If Not BdGrupo.EOF Then

    If Not IsNull(BdGrupo.Fields("fecha_prox_gen")) Then
        DTPicker1.Value = BdGrupo.Fields("fecha_prox_gen")
    Else
        DTPicker1.Value = 0
        Band = True
    End If
    If Not IsNull(BdGrupo.Fields("fecha_prox_gen2")) Then
        DTPicker2.Value = Format(BdGrupo.Fields("fecha_prox_gen2"), "dd/MM/yyyy")
    Else
        DTPicker2.Value = 0
        Band = True
    End If
    If Not IsNull(BdGrupo.Fields("periodo")) Then
        LblPeriodo2.Caption = Val(BdGrupo.Fields("periodo"))
    Else
        LblPeriodo2.Caption = "-"
        Band = True
    End If
    
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMM VERIFICA SI LA NOMINA ANTERIOR SE CERRO  MMMMMMMMMMMMMMMMMM
    
    CSql = "SELECT MAX(Periodo) AS Period FROM Historico_Nomina"
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then
        If (Val(RsTemp.Fields("Period").Value) >= Val(LblPeriodo2.Caption)) Then
            MsgBox "La nómina para el período seleccionado ya fue cerrada!", vbInformation + vbOKOnly, "Información."
            BtnGenerarNomina.Enabled = False
            BtnCerrarNomina.Enabled = False
            Exit Sub
        End If
        If Not (Val(RsTemp.Fields("Period").Value) + 1 = Val(LblPeriodo2.Caption)) Then
            MsgBox "No se ha cerrado la nómina anterior!", vbCritical + vbOKOnly, "NO SE PUEDE GENERAR LA NÓMINA"
            BtnGenerarNomina.Enabled = False
            BtnCerrarNomina.Enabled = False
            'ChameleonBtn1.Enabled = False
            Exit Sub
        End If
    End If
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' consulta para saber la cantidad de empleados para el grupo seleccionado
    
    CSql = "select * from empleados where id_grupo = " & Combo1.ItemData(Combo1.ListIndex) & " and activo=1"
    Set RsTemp = CrearRS(CSql)
    
    ' si hay empleados entonces dame la cantidad
    If RsTemp.RecordCount <> 0 Then
        VarTemp = RsTemp.RecordCount
        
        CSql = "select * from empleados where activo=1"
        Set RsTemp = CrearRS(CSql)
        
        LblNroEmpl.Caption = "Nro de Empleados:  " & VarTemp & " / " & RsTemp.RecordCount
    Else
    ' si no hay empleados entonces coloca CERO
        CSql = "select * from empleados where activo=1"
        Set RsTemp = CrearRS(CSql)
        
        LblNroEmpl.Caption = "Nro de Empleados:  0 / " & RsTemp.RecordCount
    End If
    
    ' verifica la cantidad de empleados a los cuales se les generara la nómina
    ' Se usaron RELACIONES DE TABLAS junto a una SUBCONSULTA
    CSql = "SELECT CamposDelTrabajador.*, Empleados.Id_Grupo FROM CamposDelTrabajador INNER JOIN " & _
    " Empleados ON CamposDelTrabajador.IdEmpleado = Empleados.IdEmpleado WHERE (CamposDelTrabajador.IdEmpleado IN " & _
    " (SELECT IdEmpleado FROM CamposDelTrabajador WHERE (Tipo = 'CA') AND (IdCampoNomina = 1))) AND " & _
    " (CamposDelTrabajador.Tipo = 'CO') AND (CamposDelTrabajador.IdCampoNomina = 1) AND (Empleados.Id_Grupo = " & Combo1.ItemData(Combo1.ListIndex) & ") " & _
    " ORDER BY CamposDelTrabajador.IdEmpleado"

    Set RsTemp = CrearRS(CSql)

    If RsTemp.RecordCount <> 0 Then
        LblProc.Caption = "Nro de nóminas a generar  :  " & RsTemp.RecordCount & " / " & VarTemp
    Else
        LblProc.Caption = "Nro de nóminas a generar  :  0 / " & VarTemp
    End If
    
    ' consulta la cantidad de empleados de un grupo especifico cuyas nóminas ya han sido generadas,
    ' y por la tanto ya contiene un recibo de pago...
    CSql = "SELECT * FROM Recibos WHERE Id_Grupo=" & Combo1.ItemData(Combo1.ListIndex)
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then
        LblGeneradas.Caption = "Nro de nóminas generadas:  " & RsTemp.RecordCount & " / " & VarTemp
        BtnCerrarNomina.Enabled = True
    Else
        LblGeneradas.Caption = "Nro de nóminas generadas:  0 / " & VarTemp
        BtnCerrarNomina.Enabled = False
    End If
    
    If Verificar_Nomina Or Band = True Then
        'BtnCerrarNomina.Enabled = False
        BtnGenerarNomina.Enabled = False
    Else
        'BtnCerrarNomina.Enabled = True
        BtnGenerarNomina.Enabled = True
    End If

End If
End Sub

Private Sub Form_Load()
Centrar Me
CSql = "select * from grupo"
Set BdGrupo = CrearRS(CSql)
If Not BdGrupo.EOF Then
    BdGrupo.MoveFirst
    Do While Not BdGrupo.EOF
        Combo1.AddItem BdGrupo.Fields("descripcion")
        Combo1.ItemData(Combo1.NewIndex) = BdGrupo.Fields("id_grupo")
        BdGrupo.MoveNext
    Loop
    If BdGrupo.State Then BdGrupo.Close
End If

If Verificar_Nomina Then
    'BtnCerrarNomina.Enabled = False
    BtnGenerarNomina.Enabled = False
    'MsgBox "Debe cambiar la fecha de Generación de la Nómina!", vbExclamation + vbOKOnly, "La nómina fue cerrada."
Else
    'BtnCerrarNomina.Enabled = True
    BtnGenerarNomina.Enabled = True
End If
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            If DTPicker1.Enabled = True Then DTPicker1.SetFocus
        Case vbKeyDown
            If DTPicker1.Enabled = True Then DTPicker1.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            If DTPicker2.Enabled = True Then DTPicker2.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyDown
            If DTPicker2.Enabled = True Then DTPicker2.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnGenerarNomina.SetFocus
        Case vbKeyUp
            If DTPicker1.Enabled = True Then DTPicker1.SetFocus
        Case vbKeyDown
            BtnGenerarNomina.SetFocus
    End Select
End If
End Sub


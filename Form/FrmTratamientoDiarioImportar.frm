VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTratamientoDiarioImportar 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicar registros para tratamientos"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "FrmTratamientoDiarioImportar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   7695
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         ToolTipText     =   "Cerrar Tablas de Pacientes"
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
         MICON           =   "FrmTratamientoDiarioImportar.frx":1002
         PICN            =   "FrmTratamientoDiarioImportar.frx":101E
         PICH            =   "FrmTratamientoDiarioImportar.frx":11E7
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
         TabIndex        =   6
         ToolTipText     =   "Agregar Pacientes"
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
         MICON           =   "FrmTratamientoDiarioImportar.frx":141C
         PICN            =   "FrmTratamientoDiarioImportar.frx":1438
         PICH            =   "FrmTratamientoDiarioImportar.frx":15C5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Anexar"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Borrar y crear"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "DESTINO DE TRATAMIENTOS"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Tratamiento"
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   7455
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento Finalizado"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   2880
            Width           =   1935
         End
         Begin SystemOncoAmerica.DMGrid DMGrid2 
            Height          =   2535
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   7215
            _extentx        =   12726
            _extenty        =   4471
            Object.width           =   7185
            Object.height          =   2505
            cols            =   7
            rows            =   0
            scrollbar       =   1
            marqueestyle    =   2
         End
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7455
         _extentx        =   13150
         _extenty        =   3201
         Object.width           =   7425
         Object.height          =   1530
         cols            =   15
         rows            =   0
         marqueestyle    =   2
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   11760
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Tratamiento Diario"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
End
Attribute VB_Name = "FrmTratamientoDiarioImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargar As New ADODB.Recordset
Dim RsTratamiendo As New ADODB.Recordset
Dim i As Integer
Dim Est As String
Dim RutaInformes1 As String
Dim RsTratamiento As New ADODB.Recordset
Public IdReg2, NCampo

Private Sub BtnAgregar_Click()
Dim IdReg1  As Integer
Dim TamDMGrid As Integer
Dim i As Integer
Dim info1 As String
Dim info2 As String
Dim info3 As String
Dim info4 As String
Dim info5 As String
Dim info6 As String
Dim info7 As String

    If DMGrid1.Row = 0 Then Exit Sub
    
    TamDMGrid = FrmTratamientoDiario.DMGrid2.Rows
    
    If Option1(0).Value = True Then
        CSql = "DELETE FROM Tratam_Dado WHERE IdPaciente='" & FrmRadioTerapia.IdPaciente & "' AND Campo='" & Trim(DMGrid1.ValorCelda(DMGrid1.Row, 1)) & "'"
        Set RsTemp = CrearRS(CSql)
    End If
    
    For i = 1 To TamDMGrid
    
        CSql = "Select MAX(IdReg)+1 as NuevoId From Tratam_Dado"
        Set RsTemp = CrearRS(CSql)
        
        If Not IsNull(RsTemp.Fields("NuevoId")) Then
            IdReg1 = RsTemp.Fields("NuevoId").Value
        Else
            IdReg1 = "1"
        End If
        
        info1 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 1)
        info2 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 3)
        info3 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 5)
        info4 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 7)
        info5 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 9)
        info6 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 10)
        info7 = FrmTratamientoDiario.DMGrid2.ValorCelda(i, 11)
        
        If IsNumeric(info2) Then
            CSql = "INSERT INTO Tratam_Dado (IdReg, IdPaciente, IdUsuario, Fecha, Campo, UM, ICRU, Tecnico, Hora, Simulacion, Finalizado, IdL, IdLIdPac) " & _
                    "VALUES (" & IdReg1 & "," & IdPac1 & "," & IdUser & ",'" & Replace(Replace(info1, ".", ""), ",", ".") & _
                    "'," & Trim(DMGrid1.ValorCelda(DMGrid1.Row, 1)) & "," & Replace(Replace(info2, ".", ""), ",", ".") & _
                    "," & Replace(Replace(info3, ".", ""), ",", ".") & ",'" & info4 & "','" & info7 & "'," & _
                    Replace(Replace(info5, ".", ""), ",", ".") & "," & Replace(Replace(info6, ".", ""), ",", ".") & ",'" & FrmRadioTerapia.IdLIdInf & "','" & FrmRadioTerapia.IdLIdPac & "')"
            Set RsTratamDado = CrearRS(CSql)
        Else
            If Option1(0).Value = True Then
                info2 = "0"
                CSql = "INSERT INTO Tratam_Dado (IdReg, IdPaciente, IdUsuario, Fecha, Campo, UM, ICRU, Tecnico, Hora, Simulacion, Finalizado, IdL, IdLIdPac) " & _
                        "VALUES (" & IdReg1 & "," & IdPac1 & "," & IdUser & ",'" & Replace(Replace(info1, ".", ""), ",", ".") & _
                        "'," & Trim(DMGrid1.ValorCelda(DMGrid1.Row, 1)) & "," & Replace(Replace(info2, ".", ""), ",", ".") & _
                        "," & Replace(Replace(info3, ".", ""), ",", ".") & ",'" & info4 & "','" & info7 & "'," & _
                        Replace(Replace(info5, ".", ""), ",", ".") & "," & Replace(Replace(info6, ".", ""), ",", ".") & ",'" & FrmRadioTerapia.IdLIdInf & "','" & FrmRadioTerapia.IdLIdPac & "')"
                Set RsTratamDado = CrearRS(CSql)
            End If
        End If

    Next i
    
    
    EncabezadoGrid2
    CargarGrid2
    
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub


Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)

If lRow = 0 Then Exit Sub
If DMGrid1.ValorCelda(lRow, 1) = "" Then Exit Sub
If Button = vbRightButton Then
End If

If Button = vbLeftButton Then
    CargarGrid2
End If

End Sub

Sub CargarGrid2()
    CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And Campo='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsTratamiento = CrearRS(CSql)
    
    If RsTratamiento.RecordCount > 0 Then
        If RsTratamiento.Fields("Finalizado").Value = True Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    
        i = 1
        DMGrid2.Clear
        
        Do While Not RsTratamiento.EOF
            DMGrid2.Rows = i
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 And RsTratamiento.Fields("Simulacion").Value <> 1 Then
                If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                    
                    DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                    DMGrid2.ValorCelda(i, 2) = i - 1
                    DMGrid2.ValorCelda(i, 3) = "Simulacion"
                    DMGrid2.ValorCelda(i, 4) = ""
                    DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                    Total = Total + Val(RsTratamiento.Fields("ICRU").Value)
                    DMGrid2.ValorCelda(i, 6) = Total
                    DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                    DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                    i = i + 1
                Else
                
                    DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                    DMGrid2.ValorCelda(i, 2) = i - 1
                    DMGrid2.ValorCelda(i, 3) = RsTratamiento.Fields("UM").Value
                    DMGrid2.ValorCelda(i, 4) = ""
                    DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                    Total = Total + Val(RsTratamiento.Fields("ICRU").Value)
                    DMGrid2.ValorCelda(i, 6) = Total
                    DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                    DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                    i = i + 1
                End If
            Else
                DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                DMGrid2.ValorCelda(i, 2) = i - 1
                DMGrid2.ValorCelda(i, 3) = RsTratamiento.Fields("UM").Value
                DMGrid2.ValorCelda(i, 4) = ""
                DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                Total = Total + Val(RsTratamiento.Fields("ICRU").Value)
                DMGrid2.ValorCelda(i, 6) = Total
                DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                i = i + 1
                
            End If
            
            
            DMGrid1.ValorCelda(i, 14) = RsTratamiento.Fields("Simulacion").Value
            DMGrid1.ValorCelda(i, 15) = RsTratamiento.Fields("Finalizado").Value
            RsTratamiento.MoveNext
        
        Loop
        DMGrid2.PaintMGrid
    Else
        DMGrid2.Rows = 0
        DMGrid2.Clear
        DMGrid2.PaintMGrid
    End If

End Sub

Private Sub DMGrid2_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)

If lRow = 0 Then Exit Sub
If DMGrid2.ValorCelda(lRow, 1) = "" Then Exit Sub


If Button = vbRightButton Then
IdRe = DMGrid2.ValorCelda(DMGrid2.Row, 8)
If IdRe <> "" Then
ACCION = EDITAR_REGISTRO
CSql = "Select * From Tratam_Dado Where IdReg='" & DMGrid2.ValorCelda(lRow, 8) & "' And IdPaciente='" & IdPac1 & "'"
Set RsTratamiento = CrearRS(CSql)

With FrmTratamiendoDado
    IdReg2 = RsTratamiento.Fields("IdReg").Value
    .DtPickerFecha.Value = RsTratamiento.Fields("Fecha").Value
    .TxtCampo.Text = RsTratamiento.Fields("Campo").Value
    .TxtDosis.Text = RsTratamiento.Fields("UM").Value
    .TxtICRU.Text = RsTratamiento.Fields("ICRU").Value
    .TxtTecnico.Text = RsTratamiento.Fields("Tecnico").Value
    
    If RsTratamiento.Fields("Simulacion").Value = False Then
        .Check1.Value = 0
    Else
        .Check1.Value = 1
    End If
    
    If RsTratamiento.Fields("Finalizado").Value = False Then
        .Check2.Value = 0
    Else
        .Check2.Value = 1
    End If
    
    .Show vbModal
End With
Else
    Msg = "Seleccione un tratamiendo para poder modificarlo!!"
    MsgBox Msg, vbCritical + vbOKOnly, "Error"
    Exit Sub
End If
End If
End Sub

Private Sub Form_Load()
Centrar Me
EncabezadoGrid1
EncabezadoGrid2

CargarGrid1
'CargarGrid2
End Sub

Sub EncabezadoGrid1()

' carga las columnas y encabezados de columna
DMGrid1.Cols = 16
DMGrid1.VisibleCols = 13
DMGrid1.RightCol = 13
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 1
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1
DMGrid1.DColumnas(6).Alignment = 1
DMGrid1.DColumnas(7).Alignment = 1
DMGrid1.DColumnas(8).Alignment = 1
DMGrid1.DColumnas(9).Alignment = 1
DMGrid1.DColumnas(10).Alignment = 1
DMGrid1.DColumnas(11).Alignment = 1
DMGrid1.DColumnas(12).Alignment = 1
DMGrid1.DColumnas(13).Alignment = 1

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(6).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(7).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(8).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(9).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(10).Width = Val(DMGrid1.Width * 8 / 100)
DMGrid1.DColumnas(11).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(12).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(13).Width = Val(DMGrid1.Width * 7 / 100)

DMGrid1.DColumnas(1).Caption = "Campo"
DMGrid1.DColumnas(2).Caption = "Col Upper"
DMGrid1.DColumnas(3).Caption = "Col Lower"
DMGrid1.DColumnas(4).Caption = "Z o IDL"
DMGrid1.DColumnas(5).Caption = "Técnica"
DMGrid1.DColumnas(6).Caption = "Dir"
DMGrid1.DColumnas(7).Caption = "Gantry"
DMGrid1.DColumnas(8).Caption = "Colimador"
DMGrid1.DColumnas(9).Caption = "Camilla"
DMGrid1.DColumnas(10).Caption = "Bolus"
DMGrid1.DColumnas(11).Caption = "Cuña"
DMGrid1.DColumnas(12).Caption = "Bloque"
DMGrid1.DColumnas(13).Caption = "U.M."

DMGrid1.DColumnas(14).Visible = False
DMGrid1.DColumnas(15).Visible = False

DMGrid1.PaintMGrid

End Sub

Sub CargarGrid1()

CSql = "Select * From Tecnica2 Where idPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdTecnica='" & FrmRadioTerapia.Camp & "' order by campo"
Set RsCargar = CrearRS(CSql)
If RsCargar.RecordCount > 0 Then
    i = 1
    Do While Not RsCargar.EOF
        DMGrid1.Rows = i
        
        DMGrid1.ValorCelda(i, 1) = RsCargar.Fields("Campo").Value
        DMGrid1.ValorCelda(i, 2) = RsCargar.Fields("Upper").Value
        DMGrid1.ValorCelda(i, 3) = RsCargar.Fields("Lower").Value
        DMGrid1.ValorCelda(i, 4) = 0
        DMGrid1.ValorCelda(i, 5) = RsCargar.Fields("Lower").Value
        DMGrid1.ValorCelda(i, 6) = RsCargar.Fields("direccion").Value
        DMGrid1.ValorCelda(i, 7) = RsCargar.Fields("Gantry").Value
        DMGrid1.ValorCelda(i, 8) = RsCargar.Fields("Colimador").Value
        DMGrid1.ValorCelda(i, 9) = RsCargar.Fields("Camilla").Value
        If RsCargar.Fields("Bolus").Value = "True" Then
            DMGrid1.ValorCelda(i, 10) = "Si"
        Else
            DMGrid1.ValorCelda(i, 10) = "No"
        End If
        DMGrid1.ValorCelda(i, 11) = RsCargar.Fields("Cuña").Value
        
        If RsCargar.Fields("Bloque").Value = "True" Then
            DMGrid1.ValorCelda(i, 12) = "Si"
        Else
            DMGrid1.ValorCelda(i, 12) = "No"
        End If
        DMGrid1.ValorCelda(i, 13) = RsCargar.Fields("Dosis").Value
        i = i + 1
        RsCargar.MoveNext
    Loop
    DMGrid1.PaintMGrid
    NCampo = i - 1
Else
    DMGrid1.Rows = 0
    DMGrid1.Clear
    DMGrid1.PaintMGrid
    
    DMGrid2.Rows = 0
    DMGrid2.Clear
    DMGrid2.PaintMGrid
    NCampo = 0
End If


CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "'"
Set RsTratamiento = CrearRS(CSql)

If RsTratamiento.RecordCount > 0 Then
    If RsTratamiento.Fields("Finalizado").Value = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
        Check1.Enabled = True
    End If
End If


End Sub
Sub EncabezadoGrid2()

' carga las columnas y encabezados de columna
DMGrid2.Cols = 8
DMGrid2.Rows = 0
DMGrid2.DColumnas(1).Alignment = 1
DMGrid2.DColumnas(2).Alignment = 1
DMGrid2.DColumnas(3).Alignment = 1
DMGrid2.DColumnas(4).Alignment = 1
DMGrid2.DColumnas(5).Alignment = 1
DMGrid2.DColumnas(6).Alignment = 1
DMGrid2.DColumnas(7).Alignment = 1
DMGrid2.DColumnas(8).Alignment = 1

DMGrid2.DColumnas(1).Caption = "Fecha"
DMGrid2.DColumnas(2).Caption = "Ses."
DMGrid2.DColumnas(3).Caption = "U.M."
DMGrid2.DColumnas(4).Caption = "Piel"
DMGrid2.DColumnas(5).Caption = "Icru"
DMGrid2.DColumnas(6).Caption = "Total"
DMGrid2.DColumnas(7).Caption = "Técnico"
DMGrid2.DColumnas(8).Caption = ""

End Sub


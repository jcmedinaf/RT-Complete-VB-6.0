VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTratamientoDiario 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tratamiento Diario"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   Icon            =   "FrmTratamientoDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   7440
         Width           =   13455
         Begin ChamaleonButton.ChameleonBtn BtnBorrar 
            Height          =   375
            Left            =   2640
            TabIndex        =   30
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmTratamientoDiario.frx":1002
            PICN            =   "FrmTratamientoDiario.frx":101E
            PICH            =   "FrmTratamientoDiario.frx":11C2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnImportarTratamiento 
            Height          =   375
            Left            =   3960
            TabIndex        =   29
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Importar Tratamiento"
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
            MICON           =   "FrmTratamientoDiario.frx":1361
            PICN            =   "FrmTratamientoDiario.frx":137D
            PICH            =   "FrmTratamientoDiario.frx":15FE
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
            Left            =   12240
            TabIndex        =   10
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
            MICON           =   "FrmTratamientoDiario.frx":189A
            PICN            =   "FrmTratamientoDiario.frx":18B6
            PICH            =   "FrmTratamientoDiario.frx":1A7F
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
            TabIndex        =   11
            ToolTipText     =   "Agregar Pacientes"
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar Tratamiento"
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
            MICON           =   "FrmTratamientoDiario.frx":1CB4
            PICN            =   "FrmTratamientoDiario.frx":1CD0
            PICH            =   "FrmTratamientoDiario.frx":1E5D
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
            Left            =   7560
            TabIndex        =   12
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
            MICON           =   "FrmTratamientoDiario.frx":2092
            PICN            =   "FrmTratamientoDiario.frx":20AE
            PICH            =   "FrmTratamientoDiario.frx":21D3
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
            Left            =   10920
            TabIndex        =   27
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "FrmTratamientoDiario.frx":2463
            PICN            =   "FrmTratamientoDiario.frx":247F
            PICH            =   "FrmTratamientoDiario.frx":2761
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
         Height          =   6495
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   13455
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   1815
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   3201
            Object.Width           =   13185
            Object.Height          =   1785
            Cols            =   13
            Rows            =   0
            MarqueeStyle    =   2
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
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento Finalizado"
            Enabled         =   0   'False
            Height          =   195
            Left            =   7920
            TabIndex        =   26
            Top             =   3960
            Width           =   1935
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Recordatorio para el Técnico"
            Height          =   1695
            Left            =   7920
            TabIndex        =   15
            Top             =   2160
            Width           =   2415
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "no"
               Height          =   195
               Left            =   1560
               TabIndex        =   25
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "no"
               Height          =   195
               Left            =   1560
               TabIndex        =   24
               Top             =   1320
               Width           =   180
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "no"
               Height          =   195
               Left            =   1560
               TabIndex        =   23
               Top             =   1080
               Width           =   180
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "no"
               Height          =   195
               Left            =   1560
               TabIndex        =   22
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otros"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   1320
               Width           =   375
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Soporte"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   1080
               Width           =   555
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gap"
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   840
               Width           =   300
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bolus"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   390
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1/Semana"
               Height          =   195
               Left            =   1560
               TabIndex        =   17
               Top             =   360
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peliculas Portales"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   1245
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento"
            Height          =   4215
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   7695
            Begin SystemOncoAmerica.DMGrid DMGrid2 
               Height          =   3855
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   6800
               Object.Width           =   7425
               Object.Height          =   3825
               Cols            =   7
               Rows            =   7
               ScrollBar       =   1
               MarqueeStyle    =   2
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Paciente"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13455
         Begin VB.TextBox TxtNombre 
            Height          =   375
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox TxtApellido 
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Historia:"
            Height          =   195
            Left            =   9600
            TabIndex        =   7
            Top             =   330
            Width           =   750
         End
         Begin VB.Label LblNoHistoria 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   375
            Left            =   10560
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "A&pellido(s):"
            Height          =   195
            Left            =   210
            TabIndex        =   5
            Top             =   330
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nombre(s):"
            Height          =   195
            Left            =   4920
            TabIndex        =   4
            Top             =   330
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "FrmTratamientoDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargar As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim RsTemp2 As New ADODB.Recordset
Dim RsTratamiendo As New ADODB.Recordset
Dim i As Integer
Dim Est As String
Dim RutaInformes1 As String
Dim RsTratamiento As New ADODB.Recordset
Public IdReg2, NCampo
Public IdReg3
Dim IdRe

Private Sub BtnAgregar_Click()
Dim SumaTot As Double

Campo = DMGrid1.ValorCelda(lRow, 13)

If Campo = "" Then
    MsgBox "Seleccione un Campo para igresarle su tratamiento!!!", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMM CONSULTA PARA OBTENER LA SUMATORIA DE LOS CAMPOS MMMM
CSql = "SELECT SUM(Tratam_Dado.ICRU) AS SumaTot FROM Tecnica" & _
  " INNER JOIN dbo.Tecnica2 ON (dbo.Tecnica.Id = dbo.Tecnica2.IdTecnica) AND (dbo.Tecnica.idL = dbo.Tecnica2.idLidInf)" & _
  " RIGHT OUTER JOIN dbo.Tratam_Dado ON (dbo.Tecnica2.Campo = dbo.Tratam_Dado.Campo)" & _
  " Where Tratam_Dado.Idpaciente = " & FrmRadioTerapia.IdPaciente & " AND  Tecnica2.IdTecnica = " & Val(FrmRadioTerapia.ListView1.ListItems(FrmRadioTerapia.ListView1.SelectedItem.Index).ListSubItems(9).Text) & _
  " AND  Tecnica2.Idpaciente = " & FrmRadioTerapia.IdPaciente & ""
Set RsTemp = CrearRS(CSql)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMM CONSULTA PARA OBTENER EL VALOR DE LA DOSIS TOTAL MMMM
CSql = "Select DosisT From Tecnica Where Id=" & Val(FrmRadioTerapia.ListView1.ListItems(FrmRadioTerapia.ListView1.SelectedItem.Index).ListSubItems(9).Text)
Set RsTemp2 = CrearRS(CSql)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMM CONDICIONAL PARA COMPRAR DOSIS MMMM

If Not IsNull(RsTemp.Fields("SumaTot").Value) Then
    If CDbl(RsTemp2.Fields("DosisT").Value) <= CDbl(RsTemp.Fields("SumaTot").Value) Then
    
        Rsp = MsgBox("La dosis de " & Val(FrmRadioTerapia.ListView1.ListItems(FrmRadioTerapia.ListView1.SelectedItem.Index).ListSubItems(8).Text) & _
        " UM para la técnica de '" & UCase(FrmRadioTerapia.ListView1.ListItems(FrmRadioTerapia.ListView1.SelectedItem.Index).ListSubItems(1).Text) & _
        "' se ha completado con un total de " & CDbl(RsTemp.Fields("SumaTot").Value) & " UM!." & Chr(13) & Chr(13) & "Desea seguir agregando campos?", vbQuestion + vbYesNo, "Confirmación!")
    
        If Rsp = vbNo Then Exit Sub
    End If
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM



If Not IsNull(RsTemp.Fields(0).Value) Then
    If RsTemp.Fields(0).Value > 3 Then
    End If
End If

If Check1.Value = 1 Then MsgBox "la dosis para el campo seleccion esta completada!", vbCritical + vbOKOnly, "Error": Exit Sub
ACCION = AGREGAR_REGISTRO
IdReg2 = ""
IdReg3 = IdLDefault

CSql = "Select * From Tecnica2 Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "'"
Set RsTratamiento = CrearRS(CSql)

If Not IsNull(Campo) Then UM = Trim(Campo) Else UM = 0
ICRU = CDbl(FrmRadioTerapia.DF) / Val(NCampo)

FrmTratamiendoDado.TxtCampo.Text = DMGrid1.ValorCelda(lRow, 1)
FrmTratamiendoDado.TxtDosis.Text = UM
FrmTratamiendoDado.TxtICRU.Text = Round(ICRU, 2)
FrmTratamiendoDado.TxtTecnico.Text = FrmRadioTerapia.InicialTec
FrmTratamiendoDado.Show vbModal
End Sub

Private Sub BtnBorrar_Click()
IdRe = DMGrid2.ValorCelda(DMGrid2.Row, 8)
If IdRe <> "" Then
    
    CSql = "Select * From Tratam_Dado Where IdReg='" & IdRe & "' And IdL='" & DMGrid2.ValorCelda(i, 9) & "' And IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "'"
    
    Set RsTratamiento = CrearRS(CSql)

    RsTratamiento.Delete
    
    EnviarRegPendiente IdRe, DMGrid2.ValorCelda(i, 9)
    
    Msg = "Tratamiento Borrado Satisfactoriamente!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Borrado Exitoso"
    
    EncabezadoGrid2
    CargarGrid2
    
    Exit Sub
Else
    MsgBox "Seleccione el tratamiento a Borrar!!!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End If

End Sub


Sub EnviarRegPendiente(ByVal NuevoId2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tratam_Dado WHERE IdReg = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then
    StrSen = "DELETE FROM Tratam_Dado WHERE IdReg = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Else
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    StrSen = "INSERT INTO Tratam_Dado (["
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
End If

CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Edicion Campos Tecnico- TABLA: Tratam_Dado"
RsRegPendiente.Fields("Tabla").Value = "Tratam_Dado"
RsRegPendiente.Fields("Condicional").Value = "IdReg=" & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub


Private Sub BtnCerrar_Click()
Unload Me
End Sub


Private Sub BtnImportarTratamiento_Click()
    FrmTratamientoDiarioImportar.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnImprimir_Click()

''========= ESTE ES EL CODIGO NUEVO ==========

Est = String$(255, " ")

i = GetPrivateProfileString("Opciones", "RutaInformes", "", Est, Len(Est), "Informes.ini")

If i > 0 Then
    RutaInformes1 = Trim(Est)
    RutaInformes = Mid(RutaInformes1, 1, Len(RutaInformes1) - 1)
End If
'***********************************************************************************************************************
'-Primero debes cargar la Referencia de Excel
'-Luego abres la Instancia de Excel

Set OBJ = New Excel.Application
OBJ.Visible = False
'OBJ.Workbooks.Open App.Path & "\REPOR.XLS"
OBJ.Workbooks.Open RutaInformes & "\TratamientoDiario.xls"
'
'-Luego Asignas Valores a las Celdas deacuerdo a tus criterios

OBJ.ActiveSheet.Cells(9, 14).Value = Trim(TxtApellido.Text) & ", " & (TxtNombre.Text)
OBJ.ActiveSheet.Cells(10, 12).Value = Trim(LblNoHistoria.Caption)


'***********************************************************************************************************************
' Informacion de los campos que posee el paciente

CSql = "Select * From Tecnica2 Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "'"
Set RsTratamiento = CrearRS(CSql)

If RsTratamiento.RecordCount > 0 Then

    Do While Not RsTratamiento.EOF
        If RsTratamiento.Fields("campo").Value = 1 Then
            OBJ.ActiveSheet.Cells(11, 5).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 5).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 5).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 5).Value = RsTratamiento!Alias
            OBJ.ActiveSheet.Cells(15, 5).Value = RsTratamiento!Tecnica
            OBJ.ActiveSheet.Cells(16, 5).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 5).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 5).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 5).Value = RsTratamiento!Camilla
            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 5).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 5).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 5).Value = RsTratamiento!Cuña
            
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 5).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 5).Value = "No"
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 2 Then
        
            OBJ.ActiveSheet.Cells(11, 9).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 9).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 9).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 9).Value = "0"
            OBJ.ActiveSheet.Cells(15, 9).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 9).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 9).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 9).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 9).Value = RsTratamiento!Camilla
            
            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 9).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 9).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 9).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 9).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 9).Value = "No"
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 3 Then
            
            OBJ.ActiveSheet.Cells(11, 13).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 13).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 13).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 13).Value = "0"
            OBJ.ActiveSheet.Cells(15, 13).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 13).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 13).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 13).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 13).Value = RsTratamiento!Camilla
            
            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 13).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 13).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 13).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 13).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 13).Value = "No"
            End If
        
        ElseIf RsTratamiento.Fields("campo").Value = 4 Then
            
            OBJ.ActiveSheet.Cells(11, 17).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 17).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 17).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 17).Value = "0"
            OBJ.ActiveSheet.Cells(15, 17).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 17).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 17).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 17).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 17).Value = RsTratamiento!Camilla
            
            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 17).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 17).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 17).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 17).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 17).Value = "No"
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 5 Then
            
            OBJ.ActiveSheet.Cells(11, 21).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 21).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 21).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 21).Value = "0"
            OBJ.ActiveSheet.Cells(15, 21).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 21).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 21).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 21).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 21).Value = RsTratamiento!Camilla
            
            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 21).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 21).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 21).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 21).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 21).Value = "No"
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 6 Then
            
            OBJ.ActiveSheet.Cells(11, 25).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 25).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 25).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 25).Value = "0"
            OBJ.ActiveSheet.Cells(15, 25).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 25).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 25).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 25).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 25).Value = RsTratamiento!Camilla

            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 25).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 25).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 25).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 25).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 25).Value = "No"
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 7 Then
            
            OBJ.ActiveSheet.Cells(11, 29).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 29).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 29).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 29).Value = "0"
            OBJ.ActiveSheet.Cells(15, 29).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 29).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 29).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 29).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 29).Value = RsTratamiento!Camilla
            
            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 29).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 29).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 29).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 29).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 29).Value = "No"
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 8 Then
            
            OBJ.ActiveSheet.Cells(11, 33).Value = RsTratamiento!Campo
            OBJ.ActiveSheet.Cells(12, 33).Value = RsTratamiento!Upper
            OBJ.ActiveSheet.Cells(13, 33).Value = RsTratamiento!Lower
            OBJ.ActiveSheet.Cells(14, 33).Value = "0"
            OBJ.ActiveSheet.Cells(15, 33).Value = RsTratamiento!descripcion
            OBJ.ActiveSheet.Cells(16, 33).Value = RsTratamiento!Direccion
            OBJ.ActiveSheet.Cells(17, 33).Value = RsTratamiento!Gantry
            OBJ.ActiveSheet.Cells(18, 33).Value = RsTratamiento!Colimador
            OBJ.ActiveSheet.Cells(19, 33).Value = RsTratamiento!Camilla

            If RsTratamiento.Fields("Bolus").Value = "True" Then
                OBJ.ActiveSheet.Cells(20, 33).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(20, 33).Value = "No"
            End If
            
            OBJ.ActiveSheet.Cells(21, 33).Value = RsTratamiento!Cuña
             
            If RsTratamiento.Fields("Bloque").Value = "True" Then
                OBJ.ActiveSheet.Cells(22, 33).Value = "Si"
            Else
                OBJ.ActiveSheet.Cells(22, 33).Value = "No"
            End If
            
        End If
        RsTratamiento.MoveNext
    Loop
    
'***********************************************************************************************************************
' informacion de las dosis suministradas por dias de tratamiento

CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' order by fecha"
Set RsTratamiento = CrearRS(CSql)
Dim CampoAnt As Integer
Dim Band As Boolean

Dim ini As Integer
Dim fin As Integer
Dim J As Integer

CampoAnt = 0
If RsTratamiento.RecordCount > 0 Then
    i = 1
    J = 25
    ini = J

    Do While Not RsTratamiento.EOF
        'OBJ.ActiveSheet.Cells(j, 1).Value = RsTratamiento.Fields("Fecha").Value
        If CampoAnt <> Val(RsTratamiento.Fields("campo").Value) Then CampoAnt = Val(RsTratamiento.Fields("campo").Value): Band = True
        
        If Band Then fin = J
        
        If RsTratamiento.Fields("campo").Value = 1 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 2).Value = 0
                OBJ.ActiveSheet.Cells(J, 3).Value = "Simulación"
                OBJ.ActiveSheet.Cells(J, 4).Value = ""
                OBJ.ActiveSheet.Cells(J, 5).Value = 0
                OBJ.ActiveSheet.Cells(J, 6).Value = 0
                
            Else
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 2).Value = i - 1
                OBJ.ActiveSheet.Cells(J, 3).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 4).Value = ""
                OBJ.ActiveSheet.Cells(J, 5).Value = RsTratamiento.Fields("ICRU").Value
                total1 = total1 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 6).Value = total1
                
            End If
      
        ElseIf RsTratamiento.Fields("campo").Value = 2 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 7).Value = 0
                OBJ.ActiveSheet.Cells(J, 8).Value = ""
                OBJ.ActiveSheet.Cells(J, 9).Value = 0
                OBJ.ActiveSheet.Cells(J, 10).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 7).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 8).Value = ""
                OBJ.ActiveSheet.Cells(J, 9).Value = RsTratamiento.Fields("ICRU").Value
                Total2 = Total2 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 10).Value = Total2
               
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 3 Then
            
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 11).Value = 0
                OBJ.ActiveSheet.Cells(J, 12).Value = ""
                OBJ.ActiveSheet.Cells(J, 13).Value = 0
                OBJ.ActiveSheet.Cells(J, 14).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 11).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 12).Value = ""
                OBJ.ActiveSheet.Cells(J, 13).Value = RsTratamiento.Fields("ICRU").Value
                Total3 = Total3 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 14).Value = Total3

            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 4 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 15).Value = 0
                OBJ.ActiveSheet.Cells(J, 16).Value = ""
                OBJ.ActiveSheet.Cells(J, 17).Value = 0
                OBJ.ActiveSheet.Cells(J, 18).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 15).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 16).Value = ""
                OBJ.ActiveSheet.Cells(J, 17).Value = RsTratamiento.Fields("ICRU").Value
                Total4 = Total4 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 18).Value = Total4

            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 5 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 19).Value = 0
                OBJ.ActiveSheet.Cells(J, 20).Value = ""
                OBJ.ActiveSheet.Cells(J, 21).Value = 0
                OBJ.ActiveSheet.Cells(J, 22).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 19).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 20).Value = ""
                OBJ.ActiveSheet.Cells(J, 21).Value = RsTratamiento.Fields("ICRU").Value
                Total5 = Total5 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 22).Value = Total5
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 6 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 23).Value = 0
                OBJ.ActiveSheet.Cells(J, 24).Value = ""
                OBJ.ActiveSheet.Cells(J, 25).Value = 0
                OBJ.ActiveSheet.Cells(J, 26).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 23).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 24).Value = ""
                OBJ.ActiveSheet.Cells(J, 25).Value = RsTratamiento.Fields("ICRU").Value
                Total6 = Total6 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 26).Value = Total6
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 7 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 27).Value = 0
                OBJ.ActiveSheet.Cells(J, 28).Value = ""
                OBJ.ActiveSheet.Cells(J, 29).Value = 0
                OBJ.ActiveSheet.Cells(J, 30).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 27).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 28).Value = ""
                OBJ.ActiveSheet.Cells(J, 29).Value = RsTratamiento.Fields("ICRU").Value
                Total7 = Total7 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 30).Value = Total7
            End If
            
        ElseIf RsTratamiento.Fields("campo").Value = 8 Then
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                If Band Then J = ini: Band = False
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 31).Value = 0
                OBJ.ActiveSheet.Cells(J, 32).Value = ""
                OBJ.ActiveSheet.Cells(J, 33).Value = 0
                OBJ.ActiveSheet.Cells(J, 34).Value = 0
                
            Else
                If Band Then
                    If ini <> 25 Then
                        J = ini
                    Else
                        J = fin
                        ini = fin
                    End If
                    Band = False
                End If
                OBJ.ActiveSheet.Cells(J, 1).Value = RsTratamiento.Fields("Fecha").Value
                OBJ.ActiveSheet.Cells(J, 31).Value = RsTratamiento.Fields("UM").Value
                OBJ.ActiveSheet.Cells(J, 32).Value = ""
                OBJ.ActiveSheet.Cells(J, 33).Value = RsTratamiento.Fields("ICRU").Value
                Total8 = Total8 + CDbl(RsTratamiento.Fields("ICRU").Value)
                OBJ.ActiveSheet.Cells(J, 34).Value = Total8
            End If
            
        
        End If
        
'***********************************************************************************************************************************
        
        If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
            'SubUM = CDbl(OBJ.ActiveSheet.Cells(j, 5).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 9).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 13).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 17).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 21).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 25).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 29).Value) + CDbl(OBJ.ActiveSheet.Cells(j, 33).Value)
            OBJ.ActiveSheet.Cells(J, 35).Value = 0
            'Total = Total + CDbl(SubUM)
            OBJ.ActiveSheet.Cells(J, 36).Value = 0
        Else
            SubUM = CDbl(OBJ.ActiveSheet.Cells(J, 5).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 9).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 13).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 17).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 21).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 25).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 29).Value) + CDbl(OBJ.ActiveSheet.Cells(J, 33).Value)
            OBJ.ActiveSheet.Cells(J, 35).Value = SubUM
           'Total = Total + CDbl(SubUM)
           OBJ.ActiveSheet.Cells(J, 36).Value = CDbl(OBJ.ActiveSheet.Cells(J, 35).Value) + Val(OBJ.ActiveSheet.Cells(J - 1, 36).Value)
        End If
              
'***********************************************************************************************************************************
                
        OBJ.ActiveSheet.Cells(J, 47).Value = RsTratamiento.Fields("Tecnico").Value
        
        i = i + 1
        J = J + 1
        
       RsTratamiento.MoveNext
    
    Loop
End If
CSql = "Select Finalizado From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' Order by Fecha"
Set RsTratamiento = CrearRS(CSql)
If RsTratamiento.RecordCount > 0 Then
    If RsTratamiento.Fields("Finalizado").Value = True Then
'        OBJ.ActiveSheet.Cells(63, 7).Value = "FINALIZO TRATAMIENTO"
         OBJ.ActiveSheet.Cells(J + 1, 7).Value = "FINALIZO TRATAMIENTO"
    End If
End If
CSql = "Select Min(campo) as MinCampo, Max(Campo) as MaxCampo From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "'"
Set RsTratamiento = CrearRS(CSql)
   
If Val(Mincampo) <> Val(maxcampo) Then
    OBJ.ActiveSheet.Cells(23, 35).Value = RsTratamiento.Fields("MinCampo").Value & " al " & RsTratamiento.Fields("MaxCampo").Value
Else
    OBJ.ActiveSheet.Cells(23, 35).Value = RsTratamiento.Fields("MinCampo").Value
End If
    
    '- Imprimes el Informe
    
    OBJ.ActiveSheet.PrintOut
    
    '- Cierras la Inatancia (Muy Importante)
    
    OBJ.ActiveWorkbook.Saved = True
    
    OBJ.Quit
End If
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
Dim fac As Byte

CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "' And Campo='" & DMGrid1.ValorCelda(lRow, 1) & "' And IdL='" & DMGrid1.ValorCelda(lRow, 14) & "' ORDER BY Campo, IdReg"

Set RsTratamiento = CrearRS(CSql)
    
    If RsTratamiento.RecordCount > 0 Then
        If RsTratamiento.Fields("Finalizado").Value = True Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    
        i = 1
        DMGrid2.Clear
        DMGrid2.Rows = 0
        Total = 0
        Do While Not RsTratamiento.EOF
            DMGrid2.Rows = DMGrid2.Rows + 1
            
                If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 And RsTratamiento.Fields("Simulacion").Value <> 1 And RsTratamiento.Fields("Simulacion").Value <> 1 Then
                    
                    If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                        
                        DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                        DMGrid2.ValorCelda(i, 2) = i - 1
                        DMGrid2.ValorCelda(i, 3) = "Simulacion"
                        DMGrid2.ValorCelda(i, 4) = ""
                        DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                        Total = Total + CDbl(RsTratamiento.Fields("ICRU").Value)
                        DMGrid2.ValorCelda(i, 6) = Total
                        DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                        DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                        DMGrid2.ValorCelda(i, 9) = RsTratamiento.Fields("IdL").Value
                    Else

                        DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                        DMGrid2.ValorCelda(i, 2) = i - 1
                        DMGrid2.ValorCelda(i, 3) = RsTratamiento.Fields("UM").Value
                        DMGrid2.ValorCelda(i, 4) = ""
                        DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                        Total = Total + CDbl(RsTratamiento.Fields("ICRU").Value)
                        DMGrid2.ValorCelda(i, 6) = Total
                        DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                        DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                        DMGrid2.ValorCelda(i, 9) = RsTratamiento.Fields("IdL").Value
                    End If
                Else
                    DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                    If i = 2 Then If Not IsNumeric(DMGrid2.ValorCelda(1, 3)) Then fac = 1
                    DMGrid2.ValorCelda(i, 2) = i - 1 * fac
                    DMGrid2.ValorCelda(i, 3) = RsTratamiento.Fields("UM").Value
                    DMGrid2.ValorCelda(i, 4) = ""
                    DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                    Total = Total + CDbl(Replace(RsTratamiento.Fields("ICRU").Value, ".", ","))
                    DMGrid2.ValorCelda(i, 6) = Total
                    DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                    DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                    DMGrid2.ValorCelda(i, 9) = RsTratamiento.Fields("IdL").Value
                End If
             
            If RsTratamiento.Fields("Simulacion").Value Then DMGrid2.ValorCelda(i, 9) = 1 Else DMGrid2.ValorCelda(i, 9) = 0
            If RsTratamiento.Fields("Finalizado").Value Then DMGrid2.ValorCelda(i, 10) = 1 Else DMGrid2.ValorCelda(i, 10) = 0
            DMGrid2.ValorCelda(i, 11) = RsTratamiento.Fields("Hora").Value
            i = i + 1
            RsTratamiento.MoveNext
            
        Loop
        DMGrid2.PaintMGrid
    Else
        DMGrid2.Rows = 0
        DMGrid2.Clear
        DMGrid2.PaintMGrid
        Check1.Value = 0
    End If
    
End Sub

Private Sub DMGrid2_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If DMGrid2.ValorCelda(lRow, 1) = "" Then Exit Sub


If Button = vbRightButton Then
IdRe = DMGrid2.ValorCelda(DMGrid2.Row, 8)
If IdRe <> "" Then
ACCION = EDITAR_REGISTRO
CSql = "Select * From Tratam_Dado Where IdReg='" & DMGrid2.ValorCelda(lRow, 8) & "' And IdPaciente='" & FrmRadioTerapia.IdPaciente & "'"
Set RsTratamiento = CrearRS(CSql)

With FrmTratamiendoDado
    IdReg2 = RsTratamiento.Fields("IdReg").Value
    IdReg3 = RsTratamiento.Fields("IdL").Value
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
DMGrid1.Cols = 14
DMGrid1.VisibleCols = 13
DMGrid1.RightCol = 13
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 1
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 0
DMGrid1.DColumnas(6).Alignment = 1
DMGrid1.DColumnas(7).Alignment = 1
DMGrid1.DColumnas(8).Alignment = 1
DMGrid1.DColumnas(9).Alignment = 1
DMGrid1.DColumnas(10).Alignment = 1
DMGrid1.DColumnas(11).Alignment = 1
DMGrid1.DColumnas(12).Alignment = 1
DMGrid1.DColumnas(13).Alignment = 1

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 5 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 9 / 100)
DMGrid1.DColumnas(6).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(7).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(8).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(9).Width = Val(DMGrid1.Width * 7 / 100)
DMGrid1.DColumnas(10).Width = Val(DMGrid1.Width * 7 / 100)
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
DMGrid1.DColumnas(14).Caption = "BUFFER"

DMGrid1.PaintMGrid

End Sub

Sub CargarGrid1()

CSql = "Select * From Tecnica2 Where idPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "' And IdTecnica='" & FrmRadioTerapia.Camp & "' And IdLIdInf = '" & FrmRadioTerapia.Camp2 & "' order by cast(campo as int)"

Set RsCargar = CrearRS(CSql)
If RsCargar.RecordCount > 0 Then
    i = 1
    Do While Not RsCargar.EOF
        DMGrid1.Rows = i
        
        DMGrid1.ValorCelda(i, 1) = RsCargar.Fields("Campo").Value
        DMGrid1.ValorCelda(i, 2) = RsCargar.Fields("Upper").Value
        DMGrid1.ValorCelda(i, 3) = RsCargar.Fields("Lower").Value
        DMGrid1.ValorCelda(i, 4) = 0
        DMGrid1.ValorCelda(i, 5) = RsCargar.Fields("Tecnica").Value
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
        DMGrid1.ValorCelda(i, 14) = RsCargar.Fields("IdL").Value
        
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

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "' AND Campo='" & DMGrid1.ValorCelda(DMGrid1.Row, 1) & "'"
Set RsTratamiento = CrearRS(CSql)

If RsTratamiento.RecordCount > 0 Then
    If RsTratamiento.Fields("Finalizado").Value = True Then
        Check1.Value = 1
'        Check1.Enabled = True
    Else
        Check1.Value = 0
'        Check1.Enabled = True
    End If
Else
    Check1.Value = 0
End If


End Sub
Sub EncabezadoGrid2()

' carga las columnas y encabezados de columna
DMGrid2.Cols = 12
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
DMGrid2.DColumnas(9).Caption = ""

DMGrid2.DColumnas(9).Visible = False
'DMGrid2.ValorCelda(i, 9) = RsTratamiento.Fields("IdL").Value
End Sub


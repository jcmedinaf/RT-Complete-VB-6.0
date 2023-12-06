VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTratamiendoDado 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tratamiento Dado"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "FrmTratamientoDado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   5535
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Dósis cumplida."
            Height          =   255
            Left            =   3960
            TabIndex        =   7
            Top             =   1380
            Width           =   1455
         End
         Begin VB.TextBox TxtICRU 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4320
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Simulación"
            Height          =   255
            Left            =   2640
            TabIndex        =   6
            Top             =   1380
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DtPickerFecha 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   48037889
            CurrentDate     =   40227
         End
         Begin VB.TextBox TxtCampo 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox TxtDosis 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2640
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox TxtTecnico 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   840
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icru:"
            Height          =   195
            Left            =   3960
            TabIndex        =   16
            Top             =   930
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   450
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Campo:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   930
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "U.M.:"
            Height          =   195
            Left            =   2160
            TabIndex        =   13
            Top             =   930
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Técnico:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1410
            Width           =   630
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   5535
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   4440
            TabIndex        =   9
            ToolTipText     =   "Cerrar Tablas de Pacientes"
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
            MICON           =   "FrmTratamientoDado.frx":1002
            PICN            =   "FrmTratamientoDado.frx":101E
            PICH            =   "FrmTratamientoDado.frx":11E7
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
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Guardar / Actualizar Pacientes"
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
            MICON           =   "FrmTratamientoDado.frx":141C
            PICN            =   "FrmTratamientoDado.frx":1438
            PICH            =   "FrmTratamientoDado.frx":16C7
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
Attribute VB_Name = "FrmTratamiendoDado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTratamDado As New ADODB.Recordset
Dim RsTratamiento As New ADODB.Recordset
Dim IdReg1
Dim RsTemp As New ADODB.Recordset
Private Sub BtnCerrar_Click()
Unload Me
End Sub



Sub EnviarRegPendiente(ByVal NuevoId2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tratam_dado WHERE IdReg = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then
    StrSen = "DELETE FROM Tratam_dado WHERE IdReg = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Else
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    StrSen = "INSERT INTO Tratam_dado (["
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
RsRegPendiente.Fields("Modulo").Value = "Edicion Campos Tecnico- Tabla Tratam_dado"
RsRegPendiente.Fields("Tabla").Value = "Tratam_dado"
RsRegPendiente.Fields("Condicional").Value = "IdReg=" & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub


Private Sub BtnGuardarActualizar_Click()
Dim TamDMGrid As Integer
Dim IdRegg As Integer
Dim i As Integer
    

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


Select Case ACCION

    Case Is = AGREGAR_REGISTRO
        
        If Check1.Value = 0 Then
            If Trim(TxtCampo.Text) = "" Then
                MsgBox "No puede quedar el Campo en Blanco", vbCritical + vbOKOnly, "Error"
                TxtCampo.SetFocus
                Exit Sub
            End If
            If Trim(TxtDosis.Text) = "" Then
                MsgBox "No puede quedar El U.M. en Blanco", vbCritical + vbOKOnly, "Error"
                TxtDosis.SetFocus
                Exit Sub
            End If
            If Trim(TxtICRU.Text) = "" Then
                MsgBox "No puede quedar El ICRU en Blanco", vbCritical + vbOKOnly, "Error"
                TxtICRU.SetFocus
                Exit Sub
            End If
            If Trim(TxtTecnico.Text) = "" Then
                MsgBox "No puede quedar el Nombre del Técnico en Blanco", vbCritical + vbOKOnly, "Error"
                TxtTecnico.SetFocus
                Exit Sub
            End If
        
            
            CSql = "Select MAX(IdReg)+1 as NuevoId From Tratam_Dado"
            Set RsTemp = CrearRS(CSql)
            
            If Not IsNull(RsTemp.Fields("NuevoId")) Then
                IdReg1 = RsTemp.Fields("NuevoId").Value
            Else
                IdReg1 = "1"
            End If
        
            CSql = "Select * From Tratam_Dado"
            Set RsTratamDado = CrearRS(CSql)
            
            RsTratamDado.AddNew
            RsTratamDado.Fields("IdReg").Value = IdReg1
            RsTratamDado.Fields("IdL").Value = FrmRadioTerapia.IdLIdInf
            RsTratamDado.Fields("IdPaciente").Value = FrmRadioTerapia.IdPaciente
            RsTratamDado.Fields("IdLIdPac").Value = FrmRadioTerapia.IdLIdPac
            RsTratamDado.Fields("Idusuario").Value = IdUser
            RsTratamDado.Fields("Fecha").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
            RsTratamDado.Fields("Campo").Value = TxtCampo.Text
            RsTratamDado.Fields("UM").Value = TxtDosis.Text
            RsTratamDado.Fields("ICRU").Value = CDbl(TxtICRU.Text)
            RsTratamDado.Fields("Tecnico").Value = TxtTecnico.Text
            RsTratamDado.Fields("Hora").Value = DateTime.Time
            RsTratamDado.Fields("Simulacion").Value = Check1.Value
            RsTratamDado.Fields("Finalizado").Value = Check2.Value
            RsTratamDado.Update
            
            EnviarRegPendiente IdReg1, NuevoIdL
                       
        End If
        
        If Check1.Value = 1 Then
            If TxtCampo.Text = "" Then
                MsgBox "No puede quedar el Campo en Blanco", vbCritical + vbOKOnly, "Error"
                TxtCampo.SetFocus
                Exit Sub
            End If
            If TxtDosis.Text = "" Then
                MsgBox "No puede quedar la dosis en Blanco", vbCritical + vbOKOnly, "Error"
                TxtDosis.SetFocus
                Exit Sub
            End If
            If TxtTecnico.Text = "" Then
                MsgBox "No puede quedar el Nombre del Técnico en Blanco", vbCritical + vbOKOnly, "Error"
                TxtTecnico.SetFocus
                Exit Sub
            End If
        
            CSql = "Select MAX(IdReg)+1 as NuevoId From Tratam_Dado"
            Set RsTemp = CrearRS(CSql)
            
            If Not IsNull(RsTemp.Fields("NuevoId")) Then
                IdReg1 = RsTemp.Fields("NuevoId").Value
            Else
                IdReg1 = "1"
            End If
        
            CSql = "Select * From Tratam_Dado"
            Set RsTratamDado = CrearRS(CSql)
            
            RsTratamDado.AddNew
            
            RsTratamDado.Fields("IdReg").Value = IdReg1
            RsTratamDado.Fields("IdL").Value = FrmRadioTerapia.IdLIdInf
            
            RsTratamDado.Fields("IdPaciente").Value = FrmRadioTerapia.IdPaciente
            RsTratamDado.Fields("IdLIdPac").Value = FrmRadioTerapia.IdLIdPac
            
            RsTratamDado.Fields("Idusuario").Value = IdUser
            RsTratamDado.Fields("Fecha").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
            RsTratamDado.Fields("Campo").Value = TxtCampo.Text
            RsTratamDado.Fields("UM").Value = 0
            RsTratamDado.Fields("ICRU").Value = 0
            RsTratamDado.Fields("Tecnico").Value = TxtTecnico.Text
            RsTratamDado.Fields("Hora").Value = DateTime.Time
            RsTratamDado.Fields("Simulacion").Value = Check1.Value
            RsTratamDado.Fields("Finalizado").Value = Check2.Value
            RsTratamDado.Update
           
            If Check2.Value = 1 Then
                CSql = "Update Tratam_Dado Set Finalizado=1 Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "'"
                Set RsTratamDado = CrearRS(CSql)
           
            End If
            
            EnviarRegPendiente IdReg1, NuevoIdL
            
            CargarGrid2
        End If

       
       
    Case Is = EDITAR_REGISTRO
    
'        If Check2.Value = 1 Then
'            CSql = "Update Tratam_Dado Set Finalizado=1 Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "'"
'            Set RsTratamDado = CrearRS(CSql)
'            FrmTratamientoDiario.Check1.Value = 1
'            FrmTratamientoDiario.Check1.Enabled = False
'        End If
        
        CSql = "Select * From Tratam_Dado where IdReg='" & FrmTratamientoDiario.IdReg2 & "' And IdL='" & FrmTratamientoDiario.IdReg3 & "' And IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And IdLIdPAc='" & FrmRadioTerapia.IdLIdPac & "'"
        Set RsTratamDado = CrearRS(CSql)
        
        'RsTratamDado.Fields("IdReg").Value = IdReg1
        RsTratamDado.Fields("IdL").Value = FrmRadioTerapia.IdLIdInf
            
        'RsTratamDado.Fields("IdPaciente").Value = FrmRadioTerapia.IdPaciente
        RsTratamDado.Fields("IdLIdPac").Value = FrmRadioTerapia.IdLIdPac
        RsTratamDado.Fields("Idusuario").Value = IdUser
        RsTratamDado.Fields("Fecha").Value = Format(DTPickerFecha.Value, "dd/mm/yyyy")
        RsTratamDado.Fields("Campo").Value = TxtCampo.Text
        RsTratamDado.Fields("UM").Value = TxtDosis.Text
        RsTratamDado.Fields("ICRU").Value = CDbl(TxtICRU.Text)
        RsTratamDado.Fields("Tecnico").Value = TxtTecnico.Text
        RsTratamDado.Fields("Hora").Value = DateTime.Time
        RsTratamDado.Fields("Simulacion").Value = Check1.Value
        RsTratamDado.Fields("Finalizado").Value = Check2.Value
        RsTratamDado.Update
    
        EnviarRegPendiente FrmTratamientoDiario.IdReg2, FrmTratamientoDiario.IdReg3
        
        FrmTratamientoDiario.EncabezadoGrid2
        FrmTratamientoDiario.CargarGrid2
        
End Select

'CargarGrid2
FrmTratamientoDiario.EncabezadoGrid2
FrmTratamientoDiario.CargarGrid2

Dim IdRegg2
If Check2.Value = 1 Then
    TamDMGrid = FrmTratamientoDiario.DMGrid2.Rows
    For i = 0 To TamDMGrid
        IdRegg = Val(FrmTratamientoDiario.DMGrid2.ValorCelda(i, 8))
        IdRegg2 = Val(FrmTratamientoDiario.DMGrid2.ValorCelda(i, 9))
        CSql = "Update Tratam_Dado Set Finalizado=1 Where IdReg = " & IdRegg & " And IdL='" & IdRegg2 & "'"
        Set RsTratamDado = CrearRS(CSql)
        EnviarRegPendiente IdRegg, IdRegg2
    Next i
End If
            
Unload Me
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    TxtDosis.Text = 0
    TxtDosis.Enabled = False
    TxtICRU.Text = 0
    TxtICRU.Enabled = False
    Check2.Enabled = False
Else
    TxtDosis.Text = ""
    TxtDosis.Enabled = True
    TxtICRU.Text = ""
    TxtICRU.Enabled = True
    Check2.Enabled = True
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check2.SetFocus
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Enabled = False
Else
    Check1.Enabled = True
End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnGuardarActualizar.SetFocus
End If
End Sub

Private Sub DTPickerFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtDosis.SetFocus
End If
End Sub

Private Sub Form_Load()
DTPickerFecha.Value = DateTime.Date
'DtPickerHora.Value = DateTime.Time
End Sub

Sub CargarGrid2()
CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "' And Campo='" & FrmTratamientoDiario.DMGrid1.ValorCelda(lRow, 1) & "'"
Set RsTratamiento = CrearRS(CSql)

  If RsTratamiento.RecordCount > 0 Then
        i = 1
        Do While Not RsTratamiento.EOF
            FrmTratamientoDiario.DMGrid2.Rows = i
            If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 And RsTratamiento.Fields("Simulacion").Value <> 1 Then
                If RsTratamiento.Fields("UM").Value = 0 And RsTratamiento.Fields("ICRU").Value = 0 Then
                    
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 2) = i - 1
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 3) = "Simulacion"
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 4) = ""
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                    Total = Total + Val(RsTratamiento.Fields("ICRU").Value)
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 6) = Total
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                    i = i + 1
                    RsTratamiento.MoveNext
                    
                Else
                
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 2) = i - 1
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 3) = RsTratamiento.Fields("UM").Value
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 4) = ""
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                    Total = Total + Val(RsTratamiento.Fields("ICRU").Value)
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 6) = Total
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                    FrmTratamientoDiario.DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                    i = i + 1
                    RsTratamiento.MoveNext
                
                End If
            Else
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 1) = RsTratamiento.Fields("Fecha").Value
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 2) = i - 1
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 3) = RsTratamiento.Fields("UM").Value
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 4) = ""
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 5) = RsTratamiento.Fields("ICRU").Value
                Total = Total + Val(RsTratamiento.Fields("ICRU").Value)
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 6) = Total
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 7) = RsTratamiento.Fields("Tecnico").Value
                FrmTratamientoDiario.DMGrid2.ValorCelda(i, 8) = RsTratamiento.Fields("IdReg").Value
                i = i + 1
                RsTratamiento.MoveNext
            End If
        Loop
        FrmTratamientoDiario.DMGrid2.PaintMGrid
    Else
        FrmTratamientoDiario.DMGrid2.Rows = 0
        FrmTratamientoDiario.DMGrid2.Clear
        FrmTratamientoDiario.DMGrid2.PaintMGrid
    End If


CSql = "Select * From Tratam_Dado Where IdPaciente='" & FrmRadioTerapia.IdPaciente & "'"
Set RsTratamiento = CrearRS(CSql)

If RsTratamiento.RecordCount > 0 Then
    If RsTratamiento.Fields("Finalizado").Value = True Then
        FrmTratamientoDiario.Check1.Enabled = False
        FrmTratamientoDiario.Check1.Value = 1
    Else
        FrmTratamientoDiario.Check1.Enabled = False
        FrmTratamientoDiario.Check1.Value = 0
    End If
End If
End Sub

Private Sub TxtDosis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtICRU.SetFocus
Else
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If

End Sub

Private Sub TxtICRU_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtTecnico.SetFocus
Else
    If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub TxtTecnico_Change()
TxtTecnico.Text = UCase(TxtTecnico.Text)
End Sub

Private Sub TxtTecnico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check1.SetFocus
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzáéíóúAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

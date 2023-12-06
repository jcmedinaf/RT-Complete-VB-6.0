VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEdicionTecnico 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición Del Tecnico"
   ClientHeight    =   4425
   ClientLeft      =   7050
   ClientTop       =   675
   ClientWidth     =   7155
   Icon            =   "EditarTec.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   6735
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   5640
            TabIndex        =   22
            ToolTipText     =   "Cerrar "
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
            MICON           =   "EditarTec.frx":1002
            PICN            =   "EditarTec.frx":101E
            PICH            =   "EditarTec.frx":11E7
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
            Left            =   1200
            TabIndex        =   23
            ToolTipText     =   "Guardar / Actualizar Tecnico"
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
            MICON           =   "EditarTec.frx":141C
            PICN            =   "EditarTec.frx":1438
            PICH            =   "EditarTec.frx":16C7
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
            TabIndex        =   24
            ToolTipText     =   "Agregar Tecnico"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "EditarTec.frx":1B08
            PICN            =   "EditarTec.frx":1B24
            PICH            =   "EditarTec.frx":1CB1
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
            Left            =   4440
            TabIndex        =   25
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
            MICON           =   "EditarTec.frx":1EE6
            PICN            =   "EditarTec.frx":1F02
            PICH            =   "EditarTec.frx":21E4
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
            Left            =   2400
            TabIndex        =   26
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
            MICON           =   "EditarTec.frx":2435
            PICN            =   "EditarTec.frx":2451
            PICH            =   "EditarTec.frx":25F5
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
         Caption         =   "Datos del Paciente"
         Height          =   3135
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6735
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   47841281
            CurrentDate     =   40304
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   9
            Left            =   1800
            TabIndex        =   3
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   14
            Left            =   3480
            TabIndex        =   8
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   13
            Left            =   1800
            TabIndex        =   7
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   12
            Left            =   240
            TabIndex        =   6
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   11
            Left            =   5040
            TabIndex        =   5
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   10
            Left            =   3480
            TabIndex        =   4
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   2
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   1
            Top             =   1200
            Width           =   6255
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dosis Total cGy:"
            Height          =   195
            Left            =   3480
            TabIndex        =   20
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dosis/Frac cGy/frac:"
            Height          =   195
            Left            =   1800
            TabIndex        =   19
            Top             =   2400
            Width           =   1515
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frac/Total:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   2400
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frac/Sem:"
            Height          =   195
            Left            =   5040
            TabIndex        =   17
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frac/Dia:"
            Height          =   195
            Left            =   3480
            TabIndex        =   16
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prof o Iso (cm o %):"
            Height          =   195
            Left            =   1800
            TabIndex        =   15
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Energia (MV):"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Técnica y Sitio de Anatomia:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dia:"
            Height          =   195
            Left            =   1080
            TabIndex        =   12
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   165
         End
      End
   End
End
Attribute VB_Name = "FrmEdicionTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bd1 As New ADODB.Recordset
Public IdPacT
Public IdLIdPacT As String
Public IdLIdinfT As String
Public IdLIdInf As String

Private Sub BtnAgregar_Click()
Dim RsTemp As New ADODB.Recordset
Dim CSql As String

Frame1.Enabled = True
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnGuardarActualizar.Enabled = True

Text1(7).SetFocus
ACCION = AGREGAR_REGISTRO
DTPicker1.Value = Format(Date, "DD/MM/YYYY")

Label2.Caption = "Nuevo Reg."
Limpiar_Campos

End Sub

Sub Limpiar_Campos()
Text1(7).Text = ""
Text1(8).Text = ""
Text1(9).Text = ""
Text1(10).Text = ""
Text1(11).Text = ""
Text1(12).Text = ""
Text1(13).Text = ""
Text1(14).Text = ""
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Limpiar_Campos
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False
End Sub

Private Sub BtnEliminar_Click()
Dim RsBorrarTecnica As New ADODB.Recordset

Msg = "Esta seguro de Eliminar el Registro cuya ID = " & Label2.Caption & " ? "
p = MsgBox(Msg, vbYesNo, "Eliminar Tecnico")
        
If p = vbYes Then
        
    CSql = "Delete From Tecnica Where Id = " & Label2.Caption & " And IdL='" & IdLIdInf & "'"
    Set RsBorrarPaciente = CrearRS(CSql)
    
    Msg = "Fue Eliminado el Registro"
    MsgBox Msg, vbOKOnly + vbInformation, "Tecnica Eliminado"
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el Servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
    EnviarRegPendiente Val(Label2.Caption), IdLIdInf
    
    Call FrmRadioTerapia.Carga_De_Datos
    Unload Me

End If

End Sub

Private Sub BtnGuardarActualizar_Click()
'Agrega el registro
'''''''''''''''''''''''''''''''
On Error GoTo WrtError


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



If Text1(7) = "" Then MsgBox "El campo TECNICA Y SITIO DE ANATOMIA no debe estar en blanco!", vbExclamation + vbOKOnly, "Faltan Datos!": Text1(7).SetFocus: Exit Sub
p = MsgBox("Se procedera a guardar los cambios realizados, Desea continuar?", vbQuestion + vbYesNo, "Confirmar!")
If p = 7 Then Exit Sub

With FrmRadioTerapia
    Select Case ACCION
        Case EDITAR_REGISTRO
    
            .RsTecnica.Fields("dias").Value = Format(DTPicker1.Value, "DD/MM/YYYY")
            .RsTecnica.Fields("tecnica").Value = Text1(7)
            If Text1(8) = "" Then .RsTecnica.Fields("energia").Value = Null Else .RsTecnica.Fields("energia").Value = Val(Text1(8))
            If Text1(9) = "" Then .RsTecnica.Fields("prof").Value = Null Else .RsTecnica.Fields("prof").Value = Val(Text1(9))
            If Text1(10) = "" Then .RsTecnica.Fields("fdia").Value = Null Else .RsTecnica.Fields("fdia").Value = Val(Text1(10))
            If Text1(11) = "" Then .RsTecnica.Fields("fsem").Value = Null Else .RsTecnica.Fields("fsem").Value = Val(Text1(11))
            If Text1(12) = "" Then .RsTecnica.Fields("ftot").Value = Null Else .RsTecnica.Fields("ftot").Value = Val(Text1(12))
            If Text1(13) = "" Then .RsTecnica.Fields("dosisf").Value = Null Else .RsTecnica.Fields("dosisf").Value = Val(Text1(13))
            If Text1(14) = "" Then .RsTecnica.Fields("dosist").Value = Null Else .RsTecnica.Fields("dosist").Value = Val(Text1(14))
            
            
            If Not IsNull(.RsTecnica.Fields("NombreTecnico").Value) Then .RsTecnica.Fields("NombreTecnico").Value = .Combo1.Text Else .RsTecnica.Fields("NombreTecnico").Value = .Combo1.Text
            If Not IsNull(.RsTecnica.Fields("NombreFisico").Value) Then .RsTecnica.Fields("NombreFisico").Value = .Combo2.Text Else .RsTecnica.Fields("NombreFisico").Value = .Combo2.Text
            If Not IsNull(.RsTecnica.Fields("NombreMedicoTratante").Value) Then .RsTecnica.Fields("NombreMedicoTratante").Value = .Text10.Text Else .RsTecnica.Fields("NombreMedicoTratante").Value = .Text10.Text
            
            If Not IsNull(.RsTecnica.Fields("tecnico").Value) Then .RsTecnica.Fields("tecnico").Value = .IdTecnico Else .RsTecnica.Fields("tecnico").Value = .IdTecnico
            If Not IsNull(.RsTecnica.Fields("fisico").Value) Then .RsTecnica.Fields("fisico").Value = .IdFisico Else .RsTecnica.Fields("fisico").Value = .IdFisico
            If Not IsNull(.RsTecnica.Fields("protocolo").Value) Then .RsTecnica.Fields("protocolo").Value = .IdProtocolo Else .RsTecnica.Fields("protocolo").Value = .IdProtocolo
    
    
            If Not IsNull(.RsTecnica.Fields("IdInforme").Value) Then .RsTecnica.Fields("IdInforme").Value = .IdInf Else .RsTecnica.Fields("IdInforme").Value = .IdInf
    
            .RsTecnica.Update
            MsgBox "Los datos del registro han sido actualizados!", vbInformation + vbOKOnly, "Operacion Exitosa!"
      
    
            Msg = "Espere un momento. Se Procederá  Actualizar la Información en el Servidor de Internet!!!"
            MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
            EnviarRegPendiente Val(Label2.Caption), IdLIdInf
    
        Case AGREGAR_REGISTRO
        
            CSql = "Select MAX(id)+1 as NuevoId From Tecnica"
            Set RsTemp = CrearRS(CSql)
            
            If RsTemp.RecordCount <> 0 Then
                If Not IsNull(RsTemp("NuevoId")) Then
                    Label2.Caption = RsTemp("NuevoId")
                Else
                    Label2.Caption = "1"
                End If
            Else
                Label2.Caption = "1"
            End If
            
            .RsTecnica.AddNew
            
            IdLIdInf = NuevoIdL ' Label2.Caption
            
            .RsTecnica.Fields("dias").Value = Format(DTPicker1.Value, "DD/MM/YYYY")
            .RsTecnica.Fields("tecnica").Value = Text1(7)
            If Text1(8) = "" Then .RsTecnica.Fields("energia").Value = Null Else .RsTecnica.Fields("energia").Value = Val(Text1(8))
            If Text1(9) = "" Then .RsTecnica.Fields("prof").Value = Null Else .RsTecnica.Fields("prof").Value = Val(Text1(9))
            If Text1(10) = "" Then .RsTecnica.Fields("fdia").Value = Null Else .RsTecnica.Fields("fdia").Value = Val(Text1(10))
            If Text1(11) = "" Then .RsTecnica.Fields("fsem").Value = Null Else .RsTecnica.Fields("fsem").Value = Val(Text1(11))
            If Text1(12) = "" Then .RsTecnica.Fields("ftot").Value = Null Else .RsTecnica.Fields("ftot").Value = Val(Text1(12))
            If Text1(13) = "" Then .RsTecnica.Fields("dosisf").Value = Null Else .RsTecnica.Fields("dosisf").Value = Val(Text1(13))
            If Text1(14) = "" Then .RsTecnica.Fields("dosist").Value = Null Else .RsTecnica.Fields("dosist").Value = Val(Text1(14))
            
            .RsTecnica.Fields("Id").Value = Label2.Caption
            .RsTecnica.Fields("IdL").Value = IdLIdInf
            .RsTecnica.Fields("IdLIdInf").Value = IdLIdinfT
            .RsTecnica.Fields("IdLIdPac").Value = IdLIdPacT
            
            .RsTecnica.Fields("IdPaciente").Value = .IdPaciente
            .RsTecnica.Fields("IdUsuario").Value = .IdUsuario
            .RsTecnica.Fields("IdInforme").Value = .IdInf
            .RsTecnica.Fields("Tecnico").Value = .IdTecnico
            .RsTecnica.Fields("Fisico").Value = .IdFisico
            .RsTecnica.Fields("Protocolo").Value = .IdProtocolo
            
            .RsTecnica.Fields("NombreTecnico").Value = Trim(.NombreTecnico)
            .RsTecnica.Fields("NombreFisico").Value = Trim(.NombreFisico)
            .RsTecnica.Fields("NombreMedicoTratante").Value = Trim(.Text10.Text)
            
            .RsTecnica.Fields("Activo").Value = 1
            .RsTecnica.Update
            
            MsgBox "Los datos han sido agregados al Registro!", vbInformation + vbOKOnly, "Operacion Exitosa!"
            
            Msg = "Espere un momento. Se Procederá  Actualizar la Información en el Servidor de Internet!!!"
            MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
            
            EnviarRegPendiente Val(Label2.Caption), IdLIdInf
            
    End Select
End With

BtnAgregar.Enabled = True
BtnEliminar.Enabled = False
BtnGuardarActualizar.Enabled = False
Call FrmRadioTerapia.LlenarGrid
WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open "c:\miarchivo.txt" For Append As #1
Print #1, MError
Close #1

End Sub


Sub EnviarRegPendiente(ByVal NuevoId2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tecnica WHERE Id = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Tecnica (["
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
RsRegPendiente.Fields("Modulo").Value = "Edicion Campos Tecnico"
RsRegPendiente.Fields("Tabla").Value = "Tecnica"
RsRegPendiente.Fields("Condicional").Value = "Id=" & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub






Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Frame1.Enabled = False
If ACCION = AGREGAR_REGISTRO Then
    Me.Caption = "Agregar nuevo registro"
    BtnEliminar.Enabled = False
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
ElseIf ACCION = EDITAR_REGISTRO Then
    Frame1.Enabled = True
    Me.Caption = "Editar registro"
    BtnEliminar.Enabled = True
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = True
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Unload Me
End If

End Sub

Private Sub Text1_Change(Index As Integer)
If Text1(12).Text = "" And Text1(13).Text = "" Then
    Text1(12).Text = ""
    Text1(13).Text = ""
ElseIf Text1(12).Text = "" Then
    Text1(14).Text = CDbl(Text1(13).Text)
ElseIf Text1(13).Text = "" Then
    Text1(14).Text = CDbl(Text1(12).Text)
Else
    Text1(14).Text = CDbl(Text1(12).Text) * CDbl(Text1(13).Text)
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If Shift = 0 Then
    Select Case Index
        Case 7
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(8).SetFocus
                Case vbKeyRight
                    Text1(11).SetFocus
                Case vbKeyDown
                    Text1(8).SetFocus
            End Select
        Case 8
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(9).SetFocus
                Case vbKeyUp
                    Text1(7).SetFocus
                Case vbKeyRight
                    Text1(12).SetFocus
                Case vbKeyDown
                    Text1(9).SetFocus
            End Select
        Case 9
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(10).SetFocus
                Case vbKeyUp
                    Text1(8).SetFocus
                Case vbKeyRight
                    Text1(13).SetFocus
                Case vbKeyDown
                    Text1(10).SetFocus
            End Select
        Case 10
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(11).SetFocus
                Case vbKeyUp
                    Text1(9).SetFocus
                Case vbKeyRight
                    Text1(14).SetFocus
                Case vbKeyDown
                    BtnAgregar.SetFocus
            End Select
        Case 11
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(12).SetFocus
                Case vbKeyLeft
                    Text1(7).SetFocus
                Case vbKeyDown
                    Text1(12).SetFocus
            End Select
        Case 12
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(13).SetFocus
                Case vbKeyUp
                    Text1(11).SetFocus
                Case vbKeyLeft
                    Text1(8).SetFocus
                Case vbKeyDown
                    Text1(13).SetFocus
            End Select
        Case 13
            Select Case KeyCode
                Case vbKeyReturn
                    Text1(14).SetFocus
                Case vbKeyUp
                    Text1(12).SetFocus
                Case vbKeyLeft
                    Text1(9).SetFocus
                Case vbKeyDown
                    Text1(14).SetFocus
            End Select
        Case 14
            Select Case KeyCode
                Case vbKeyReturn
                    BtnAgregar.SetFocus
                Case vbKeyUp
                    Text1(13).SetFocus
                Case vbKeyLeft
                    Text1(10).SetFocus
                Case vbKeyDown
                    BtnAgregar.SetFocus
            End Select
    End Select
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
 Select Case Index
        Case 7
            '///////////////////////////////////Valido TextBox: text1//////////////////////////////
            If KeyAscii = 13 Then
                Text1(8).SetFocus
            End If

        Case 8
            If KeyAscii = 13 Then
                Text1(9).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    MsgBox "El caracter digitado no es válido.", vbOKOnly + vbExclamation, "Error"
                    KeyAscii = 0
                End If
            End If
        Case 9
            If KeyAscii = 13 Then
                Text1(10).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    MsgBox "El caracter digitado no es válido.", vbOKOnly + vbExclamation, "Error"
                    KeyAscii = 0
                End If
            End If
        Case 10
            If KeyAscii = 13 Then
                Text1(11).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    MsgBox "El caracter digitado no es válido.", vbOKOnly + vbExclamation, "Error"
                    KeyAscii = 0
                End If
            End If
        Case 11
            If KeyAscii = 13 Then
                Text1(12).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    MsgBox "El caracter digitado no es válido.", vbOKOnly + vbExclamation, "Error"
                    KeyAscii = 0
                End If
            End If
        Case 12
            If KeyAscii = 13 Then
                Text1(13).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    MsgBox "El caracter digitado no es válido.", vbOKOnly + vbExclamation, "Error"
                    KeyAscii = 0
                End If
            End If
        Case 13
            If KeyAscii = 13 Then
                BtnGuardarActualizar.SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    MsgBox "El caracter digitado no es válido.", vbOKOnly + vbExclamation, "Error"
                    KeyAscii = 0
                End If
            End If
      End Select
End Sub

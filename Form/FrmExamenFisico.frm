VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmExamenFisico 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Examen Fisico"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   Icon            =   "FrmExamenFisico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   9135
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   8040
         TabIndex        =   14
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
         MICON           =   "FrmExamenFisico.frx":1002
         PICN            =   "FrmExamenFisico.frx":101E
         PICH            =   "FrmExamenFisico.frx":11E7
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
         TabIndex        =   9
         ToolTipText     =   "Guardar / Actualizar "
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
         MICON           =   "FrmExamenFisico.frx":141C
         PICN            =   "FrmExamenFisico.frx":1438
         PICH            =   "FrmExamenFisico.frx":16C7
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
         TabIndex        =   8
         ToolTipText     =   "Agregar"
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
         MICON           =   "FrmExamenFisico.frx":1B08
         PICN            =   "FrmExamenFisico.frx":1B24
         PICH            =   "FrmExamenFisico.frx":1CB1
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
         Left            =   6840
         TabIndex        =   13
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
         MICON           =   "FrmExamenFisico.frx":1EE6
         PICN            =   "FrmExamenFisico.frx":1F02
         PICH            =   "FrmExamenFisico.frx":21E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmExamenFisico.frx":2435
         PICN            =   "FrmExamenFisico.frx":2451
         PICH            =   "FrmExamenFisico.frx":26E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior 
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmExamenFisico.frx":2946
         PICN            =   "FrmExamenFisico.frx":2962
         PICH            =   "FrmExamenFisico.frx":2BF7
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
         TabIndex        =   10
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
         MICON           =   "FrmExamenFisico.frx":2E53
         PICN            =   "FrmExamenFisico.frx":2E6F
         PICH            =   "FrmExamenFisico.frx":3013
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
      Caption         =   "Ingreso "
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.TextBox TxtSignosVitales 
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   7695
      End
      Begin VB.TextBox TxtPiel 
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   7695
      End
      Begin VB.TextBox TxtCabello 
         Height          =   405
         Left            =   1200
         TabIndex        =   3
         Top             =   1320
         Width           =   7695
      End
      Begin VB.TextBox TxtTorax 
         Height          =   405
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   7695
      End
      Begin VB.TextBox TxtAbdomen 
         Height          =   405
         Left            =   1200
         TabIndex        =   5
         Top             =   2280
         Width           =   7695
      End
      Begin VB.TextBox TxtNeurologico 
         Height          =   405
         Left            =   1200
         TabIndex        =   6
         Top             =   2760
         Width           =   7695
      End
      Begin VB.TextBox TxtRevisionSistemas 
         Height          =   1005
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3240
         Width           =   7695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión de Sistemas:"
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Piel:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabello:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1425
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tórax:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1905
         Width           =   450
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neurológico:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2865
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abdomen:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signos Vitales:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   465
         Width           =   1035
      End
   End
End
Attribute VB_Name = "FrmExamenFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RegNew As String
Private Sub BtnAgregar_Click()
On Error Resume Next
IO = 1
RegNew = 1
'Call blanqueo1
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
BtnImprimir.Enabled = False
BtnAnterior.Enabled = False
BtnSiguiente.Enabled = False
BtnGuardarActualizar.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Frame2.BackColor = &HE0E0E0
Frame3.BackColor = &HE0E0E0

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnEliminar_Click()
On Error Resume Next

resp = MsgBox("se va a eliminar el registro actual, Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")

If resp = 7 Then Exit Sub

CSql = "Update Internista set Activo=2 Where IdInternista=" & IdInter
Set RsTemp = CrearRS(CSql)
MsgBox "Registro Eliminado!", vbInformation + vbOKOnly, "Operacion Exitosa!"

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor Web"
'EnviarRegPendiente

'Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "BORRAR", "Se elimino de la tabla INTERNISTA el registro de Id=" & IdInter & " Del paciente=" & IdPac1)

'BtnDesHacer_Click

End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
'If Cambio = 0 Then MsgBox "No se han realizado cambios!", vbInformation + vbOKOnly, "Informacion": Exit Sub
'If IdInter = "" And RegNew = 0 Then MsgBox "Debe seleccionar o agregar un registro para poder guardar los cambios!", vbExclamation + vbOKOnly, "Error": Exit Sub


'verifica si hay conexion al internet
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If


CSql = "SELECT MAX(IdInternista)+1 as NuevoId FROM Internista"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = RsTemp.Fields("NuevoId").Value
Else
    NuevoId = "1"
End If

Select Case RegNew
    
    Case Is = 0 'Actualiza
        
       
                      
            CSql = "Select * From Internista Where Activo=1 and IdPaciente=" & IdPac1 & " and idinternista='" & IdInter & "'"
            Set RsTemp = CrearRS(CSql)
            
            RsTemp.Fields("IdUsuario").Value = IdUser
            RsTemp.Fields("Enfernedad_Act").Value = Trim(FrmRadioTerapeuta.Text16.Text)
            RsTemp.Fields("Diagnostico").Value = Trim(FrmRadioTerapeuta.Text21.Text)
            RsTemp.Fields("Signos").Value = Trim(TxtSignosVitales.Text)
            RsTemp.Fields("ColorP").Value = Trim(TxtPiel.Text)
            RsTemp.Fields("Cabello").Value = Trim(TxtCabello.Text)
            RsTemp.Fields("Torax").Value = Trim(TxtTorax.Text)
            RsTemp.Fields("Abdomen").Value = Trim(TxtAbdomen.Text)
            RsTemp.Fields("Neurologico").Value = Trim(TxtNeurologico.Text)
            RsTemp.Fields("Revision").Value = Trim(TxtRevisionSistemas.Text)
            RsTemp.Fields("Activo").Value = 1
                   
            RsTemp.Update
            
                
            MsgBox "El Registro sea Actualizado satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        
        
     
         
    Case Is = 1 'Agrega Registro
   
       
            
            CSql = "Select * From Internista"
            Set RsTemp = CrearRS(CSql)
            
            RsTemp.AddNew
            RsTemp.Fields("IdInternista").Value = NuevoId
            RsTemp.Fields("Idpaciente").Value = IdPac1
            RsTemp.Fields("IdUsuario").Value = IdUser
            RsTemp.Fields("Enfernedad_Act").Value = Trim(FrmRadioTerapeuta.Text16.Text)
            RsTemp.Fields("Diagnostico").Value = Trim(FrmRadioTerapeuta.Text21.Text)
            RsTemp.Fields("Signos").Value = Trim(TxtSignosVitales.Text)
            RsTemp.Fields("ColorP").Value = Trim(TxtPiel.Text)
            RsTemp.Fields("Cabello").Value = Trim(TxtCabello.Text)
            RsTemp.Fields("Torax").Value = Trim(TxtTorax.Text)
            RsTemp.Fields("Abdomen").Value = Trim(TxtAbdomen.Text)
            RsTemp.Fields("Neurologico").Value = Trim(TxtNeurologico.Text)
            RsTemp.Fields("Revision").Value = Trim(TxtRevisionSistemas.Text)
            RsTemp.Fields("Activo").Value = 1
                   
            RsTemp.Update
            
            
            MsgBox "Registro Agregado Satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        
        
       
End Select


'If RegNew = 0 And Cambio = 1 Then
'    If Reg_Actual(0) <> Text9.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Enfernedad_Act de (" & Reg_Actual(0) & ") a (" & Text9.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(1) <> Text8.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Diagnostico de (" & Reg_Actual(1) & ") a (" & Text8.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(2) <> Text10.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Signos de (" & Reg_Actual(2) & ") a (" & Text10.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(3) <> Text11.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo colorP de (" & Reg_Actual(3) & ") a (" & Text11.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(4) <> Text13.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Cabello de (" & Reg_Actual(4) & ") a (" & Text13.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(5) <> Text15.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Torax de (" & Reg_Actual(5) & ") a (" & Text15.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(6) <> Text16.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Abdomen de (" & Reg_Actual(6) & ") a (" & Text16.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(7) <> Text17.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Neurologico de (" & Reg_Actual(7) & ") a (" & Text17.Text & ") del Registro IdInternista=" & IdInter)
'    If Reg_Actual(8) <> Text18.Text Then Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "MODIFICAR", "Se modifico el campo Revision de (" & Reg_Actual(8) & ") a (" & Text18.Text & ") del Registro IdInternista=" & IdInter)
'ElseIf RegNew = 1 And Cambio = 1 Then
'    Call Enviar_Bitacora(IdUser, "DIRECCION MEDICA", "INGRESAR", "Se ingreso un nuevo registro cuya IdInternista=" & NuevoId)
'End If

Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor Web"
'EnviarRegPendiente

'BtnDesHacer_Click
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly, "Error al Guardar"

Cambio = 0

End Sub

Sub EnviarRegPendiente(ByVal IdInt As Integer, ByVal IdLIdPac2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Internista WHERE IdInternista = " & IdInt & " AND IdL = '" & IdLIdPac2 & "'"
Set RsTemp = CrearRS(CSql)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
StrSen = "INSERT INTO Internista (["
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
RsRegPendiente.Fields("Modulo").Value = "Examen Fisico"
RsRegPendiente.Fields("Tabla").Value = "Internista"
RsRegPendiente.Fields("Condicional").Value = "IdInternista = " & IdInt & " AND IdL = '" & IdLIdPac2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

Private Sub Form_Load()
On Error Resume Next

Me.Caption = "Examen Fisico - Paciente: " & IdPac1
Centrar Me

If IdPac1 = "" Then
    CSql = "Select * From Internista Where Activo=1"
Else
    CSql = "Select * From Internista Where IdPaciente = " & IdPac1 & " And Activo=1"
End If
Set RsInternista = CrearRS(CSql)

If RsInternista.RecordCount = 0 Then GoTo nohay

Cambio = 0
RegNew = 0
IdInter = RsInternista.Fields("IdInternista")
If RsInternista.Fields("Enfernedad_Act").Value <> "" Then Text9.Text = RsInternista.Fields("Enfernedad_Act").Value Else Text9.Text = ""
If RsInternista.Fields("Diagnostico").Value <> "" Then Text8.Text = RsInternista.Fields("Diagnostico").Value Else Text8.Text = ""
If RsInternista.Fields("Signos").Value <> "" Then Text10.Text = RsInternista.Fields("Signos").Value Else Text10.Text = ""
If RsInternista.Fields("colorP").Value <> "" Then Text11.Text = RsInternista.Fields("colorP").Value Else Text11.Text = ""
If RsInternista.Fields("Cabello").Value <> "" Then Text13.Text = RsInternista.Fields("Cabello").Value Else Text13.Text = ""
If RsInternista.Fields("Torax").Value <> "" Then Text15.Text = RsInternista.Fields("Torax").Value Else Text15.Text = ""
If RsInternista.Fields("Abdomen").Value <> "" Then Text16.Text = RsInternista.Fields("Abdomen").Value Else Text16.Text = ""
If RsInternista.Fields("Neurologico").Value <> "" Then Text17.Text = RsInternista.Fields("Neurologico").Value Else Text17.Text = ""
If RsInternista.Fields("Revision").Value <> "" Then Text18.Text = RsInternista.Fields("Revision").Value Else Text18.Text = ""
BtnImprimir.Enabled = True
BtnEliminar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnAgregar.Enabled = True
NoReg = "Registro " & RsInternista.AbsolutePosition & " / " & RsInternista.RecordCount

Reg_Actual(0) = RsInternista.Fields("Enfernedad_Act").Value
Reg_Actual(1) = RsInternista.Fields("Diagnostico").Value
Reg_Actual(2) = RsInternista.Fields("Signos").Value
Reg_Actual(3) = RsInternista.Fields("colorP").Value
Reg_Actual(4) = RsInternista.Fields("Cabello").Value
Reg_Actual(5) = RsInternista.Fields("Torax").Value
Reg_Actual(6) = RsInternista.Fields("Abdomen").Value
Reg_Actual(7) = RsInternista.Fields("Neurologico").Value
Reg_Actual(8) = RsInternista.Fields("Revision").Value

Exit Sub

'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
nohay:

For i = 0 To 10
    Reg_Actual(i) = ""
Next i
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnImprimir.Enabled = False
BtnEliminar.Enabled = False
IdInter = ""
Cambio = 0
RegNew = 1
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text13.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
'MsgBox "No hay Datos asociados al paciente: " & Chr(13) & Chr(13) & Text3.Text & " " & Text4.Text & Chr(13) & "Se Mostraran los datos de Historia Médica en blanco", vbExclamation + vbOKOnly, "No Tiene Datos"



End Sub

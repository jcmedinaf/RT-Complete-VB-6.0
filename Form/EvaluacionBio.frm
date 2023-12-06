VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDiagnosticoNutricional 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostico Nutiricional"
   ClientHeight    =   4365
   ClientLeft      =   2100
   ClientTop       =   825
   ClientWidth     =   11355
   Icon            =   "EvaluacionBio.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11175
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   5880
         TabIndex        =   10
         Top             =   3480
         Width           =   5175
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   4080
            TabIndex        =   5
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
            MICON           =   "EvaluacionBio.frx":1002
            PICN            =   "EvaluacionBio.frx":101E
            PICH            =   "EvaluacionBio.frx":11E7
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
            TabIndex        =   3
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
            MICON           =   "EvaluacionBio.frx":141C
            PICN            =   "EvaluacionBio.frx":1438
            PICH            =   "EvaluacionBio.frx":16C7
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
            TabIndex        =   2
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
            MICON           =   "EvaluacionBio.frx":1B08
            PICN            =   "EvaluacionBio.frx":1B24
            PICH            =   "EvaluacionBio.frx":1CB1
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
            Left            =   2880
            TabIndex        =   4
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
            MICON           =   "EvaluacionBio.frx":1EE6
            PICN            =   "EvaluacionBio.frx":1F02
            PICH            =   "EvaluacionBio.frx":21E4
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
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10935
         Begin VB.TextBox Text2 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   1
            Top             =   2040
            Width           =   10695
         End
         Begin VB.TextBox Text1 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   480
            Width           =   10695
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recomendaciones Dieteticas"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1800
            Width           =   2085
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnóstico Nutricional Integral ó Evaluación Global Subjetiva"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4380
         End
      End
   End
End
Attribute VB_Name = "FrmDiagnosticoNutricional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NReg As String
Public DNI As String
Public Recom As String

Private Sub BtnAgregar_Click()
On Error Resume Next
BtnAgregar.Enabled = False
BtnGuardarActualizar.Enabled = True
NReg = 1
Blanqueo
Text1.Locked = False
Text2.Locked = False
Frame1.BackColor = &HE0E0E0
Text1.SetFocus
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Sub Blanqueo()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub BtnDesHacer_Click()
On Error Resume Next
BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
Text1.Locked = True
Text2.Locked = True
Frame1.BackColor = &HEAEFEF
Blanqueo
Text1.Text = DNI
Text2.Text = Recom
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next

If Text1.Text = "" Then
    Msg = "Esta dejando el Campo Diagnóstico Nutricional Integral ó Evaluación Global Subjetiva Vacio!!!"
    MsgBox msng, vbOKOnly + vbCritical, "Error Campo Vacio"
    Text1.SetFocus
    Exit Sub
End If

If Text2.Text = "" Then
    Msg = "Esta dejando el Campo Recomendaciones Dieteticas Subjetiva Vacio!!!"
    MsgBox msng, vbOKOnly + vbCritical, "Error Campo Vacio"
    Text2.SetFocus
    Exit Sub
End If

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
Frame1.BackColor = &HEAEFEF
Text1.Locked = True
Text2.Locked = True
DNI = Trim(Text1.Text)
Recom = Trim(Text2.Text)
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Diagnostico Nutiricional  Paciente: " & IdPac1
Centrar Me
BtnGuardarActualizar.Enabled = False
Text1.Locked = True
Text2.Locked = True

Text1.Text = DNI
Text2.Text = Recom
End Sub

Private Sub Text1_Change()
Exit Sub
Dim StrText, Chaa, pru As String
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text1.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text1.Text)
    pru = LCase(Mid(Text1.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text1.Text = StrText
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Len(Trim(Text1.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else
'KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyDown
            Text2.SetFocus
    End Select
End If
End Sub

Private Sub Text2_Change()
Exit Sub
Dim i  As Variant
StrText = ""
Chaa = ""
Chaa = UCase(Mid(Text2.Text, 1, 1))
StrText = Chaa
For i = 2 To Len(Text2.Text)
    pru = LCase(Mid(Text2.Text, i, 1))
    If pru Like " " Then
        T = 1
        StrText = StrText & " "
    Else
        If T = 0 Then
            Chaa = LCase(pru)
            StrText = StrText + Chaa
        Else
            Chaa = UCase(pru)
            StrText = StrText + Chaa
            T = 0
        End If
    End If
Next i

Text2.Text = StrText
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Len(Trim(Text2.Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else
'KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

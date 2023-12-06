VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDetalles 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles "
   ClientHeight    =   5850
   ClientLeft      =   7890
   ClientTop       =   795
   ClientWidth     =   5595
   Icon            =   "Detalles.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5595
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   4800
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
            MICON           =   "Detalles.frx":1002
            PICN            =   "Detalles.frx":101E
            PICH            =   "Detalles.frx":11E7
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
            ToolTipText     =   "Guardar / Actualizar"
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
            MICON           =   "Detalles.frx":141C
            PICN            =   "Detalles.frx":1438
            PICH            =   "Detalles.frx":16C7
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
            ToolTipText     =   "Agregar "
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
            MICON           =   "Detalles.frx":1B08
            PICN            =   "Detalles.frx":1B24
            PICH            =   "Detalles.frx":1CB1
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
            MICON           =   "Detalles.frx":1EE6
            PICN            =   "Detalles.frx":1F02
            PICH            =   "Detalles.frx":21E4
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
         Caption         =   "Trasfondo Religioso"
         Height          =   4575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5175
         Begin VB.TextBox Text1 
            Height          =   2055
            Left            =   1800
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Height          =   1575
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   2880
            Width           =   4815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Asistencia Espiritual"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   14
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lectura de la Biblia"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Asistencia a Iglesia "
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Creencia en Dios "
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Creencia Religiosa"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Obsevaciones"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "FrmDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD4 As New ADODB.Recordset 'tabla Informes medicos
Dim SQL As String
Dim bd1 As New ADODB.Recordset

Private Sub BtnAgregar_Click()
Blanqueo
Text1.SetFocus
End Sub

Sub Blanqueo()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
End Sub

Private Sub BtnGuardarActualizar_Click()
'command1
If Text1.Text = "" Then
        f = "Detalles"
        GoTo noguardA
        
        End If
If Text2.Text = "" Then
        f = "Observaciones"
        GoTo noguardA
        End If
                'hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos

Observa = Text2.Text
Detalle = Text1.Text
        Msg = "Registro Agregado satisfactoriamente"
        MsgBox Msg, vbOKOnly

Call BtnAgregar_Click
Exit Sub

noguardA:
    Msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox Msg, vbOKOnly, "Error al Guardar"
    Exit Sub
End Sub





Private Sub Form_Load()
Centrar Me
Text1.Text = Detalle
Text2.Text = Observa
End Sub

Private Sub Text1_Change()
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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            Text2.SetFocus
    End Select
End If
End Sub

Private Sub Text2_Change()
Dim StrText, Chaa, pru As String
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

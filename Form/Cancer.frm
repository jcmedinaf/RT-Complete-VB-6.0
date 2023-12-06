VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmTipoCancer 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Cancer"
   ClientHeight    =   2985
   ClientLeft      =   420
   ClientTop       =   1005
   ClientWidth     =   6555
   Icon            =   "Cancer.frx":0000
   LinkTopic       =   "Form43"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6555
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6135
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   5040
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
            MICON           =   "Cancer.frx":1002
            PICN            =   "Cancer.frx":101E
            PICH            =   "Cancer.frx":11E7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardar 
            Height          =   375
            Left            =   1200
            TabIndex        =   10
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
            MICON           =   "Cancer.frx":141C
            PICN            =   "Cancer.frx":1438
            PICH            =   "Cancer.frx":16C7
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
            MICON           =   "Cancer.frx":1B08
            PICN            =   "Cancer.frx":1B24
            PICH            =   "Cancer.frx":1CB1
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
            Left            =   3840
            TabIndex        =   12
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
            MICON           =   "Cancer.frx":1EE6
            PICN            =   "Cancer.frx":1F02
            PICH            =   "Cancer.frx":21E4
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
            TabIndex        =   13
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
            MICON           =   "Cancer.frx":2435
            PICN            =   "Cancer.frx":2451
            PICH            =   "Cancer.frx":25F5
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
         Caption         =   "Caracteristicas del Cancer"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6135
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Cancer.frx":2794
            Left            =   840
            List            =   "Cancer.frx":27A1
            TabIndex        =   3
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   840
            Width           =   5175
         End
         Begin ChamaleonButton.ChameleonBtn BtnAnterior 
            Height          =   375
            Left            =   4800
            TabIndex        =   14
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   1320
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
            MICON           =   "Cancer.frx":27BD
            PICN            =   "Cancer.frx":27D9
            PICH            =   "Cancer.frx":2A6E
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
            Left            =   5400
            TabIndex        =   15
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   1320
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
            MICON           =   "Cancer.frx":2CCA
            PICN            =   "Cancer.frx":2CE6
            PICH            =   "Cancer.frx":2F7C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1380
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   930
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   450
            Width           =   540
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   840
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "FrmTipoCancer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDCamp As New ADODB.Recordset
Dim BDCamp1 As New ADODB.Recordset

Private Sub BtnAgregar_Click()
Call Blanqueo
End Sub

Private Sub BtnAnterior_Click()
BDCamp1.MovePrevious
Call CargaDato
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
CSql = "Insert into camposdenomina(campo,tipo) VALUES('" & Text1.Text & "'," & Combo1.ListIndex & ")"
        Set BDCamp1 = CrearRS(CSql)
        Msg = "Registro Agregado Satisfactoriamente!!!"
        MsgBox Msg, vbOKOnly + vbInformation, "Agregado satisfactorio"
Call Blanqueo
BDCamp1.Close
Call Form_Load
Exit Sub
End Sub

Sub Blanqueo()
Label1.Caption = ""
Text1.Text = ""
Combo1.ListIndex = -1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub BtnSiguiente_Click()
BDCamp1.MoveNext
Call CargaDato
End Sub


Private Sub Form_Load()
Me.Height = 3930
Me.Width = 6870
Centrar Me

CSql = "SELECT * FROM camposdenomina"
Set BDCamp = CrearRS(CSql)
     
CSql = "select * from camposdenomina"
        Dim BDNom As New ADODB.Recordset
        Set BDNom = CrearRS(CSql)
     
        Nom = Format(BDNom.Fields("Id"), "000#")
        
        BDNom.Close
        Label1.Caption = Nom
        
Call CargaDato

End Sub

Sub CargaDato()

If BDCamp.EOF Then
Msg = "Llego al Final del Registro desea Volver al Principio?"
MsgBox Msg
BDCamp.MoveFirst
End If

If BDCamp.BOF Then
    Msg = "Llego al principio del registro"
    MsgBox Msg
    BDCamp.MoveLast
End If

If Trim(BDCamp.Fields("Campo")) <> "" Then Text1.Text = BDCamp.Fields("Campo")
    Nom = Format(BDCamp.Fields("Idcamponomina"), "000#")
    Label1.Caption = Nom
                    
                   For T = 0 To Combo1.ListCount - 1
                          If Combo1.ItemData(T) = BDCamp.Fields("tipo") Then
                          Combo1.ListIndex = T
                          Exit For
                          End If
                    Next T

End Sub




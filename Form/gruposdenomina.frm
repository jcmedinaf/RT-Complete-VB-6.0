VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmGruposNomina 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Conceptos Por Grupos de Nómina"
   ClientHeight    =   7710
   ClientLeft      =   5100
   ClientTop       =   3120
   ClientWidth     =   5910
   Icon            =   "gruposdenomina.frx":0000
   LinkTopic       =   "Form49"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   5910
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   5655
      Begin ChamaleonButton.ChameleonBtn BtnNuevo 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Agregar"
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
         MICON           =   "gruposdenomina.frx":1002
         PICN            =   "gruposdenomina.frx":101E
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
         Left            =   1440
         TabIndex        =   11
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
         MICON           =   "gruposdenomina.frx":15B8
         PICN            =   "gruposdenomina.frx":15D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior1 
         Height          =   375
         Left            =   2880
         TabIndex        =   12
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
         MICON           =   "gruposdenomina.frx":1863
         PICN            =   "gruposdenomina.frx":187F
         PICH            =   "gruposdenomina.frx":1B14
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente1 
         Height          =   375
         Left            =   3600
         TabIndex        =   13
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
         MICON           =   "gruposdenomina.frx":1D70
         PICN            =   "gruposdenomina.frx":1D8C
         PICH            =   "gruposdenomina.frx":2022
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
         Left            =   4560
         TabIndex        =   14
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
         MICON           =   "gruposdenomina.frx":2281
         PICN            =   "gruposdenomina.frx":229D
         PICH            =   "gruposdenomina.frx":2466
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
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Campos"
         Height          =   2415
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Width           =   5415
         Begin VB.ListBox List4 
            Height          =   1815
            ItemData        =   "gruposdenomina.frx":269B
            Left            =   3240
            List            =   "gruposdenomina.frx":269D
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   480
            Width           =   2055
         End
         Begin VB.ListBox List3 
            Height          =   1815
            ItemData        =   "gruposdenomina.frx":269F
            Left            =   120
            List            =   "gruposdenomina.frx":26A1
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnExcluirCampo 
            Height          =   375
            Left            =   2400
            TabIndex        =   25
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   1440
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
            MICON           =   "gruposdenomina.frx":26A3
            PICN            =   "gruposdenomina.frx":26BF
            PICH            =   "gruposdenomina.frx":2954
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAnexarCampo 
            Height          =   375
            Left            =   2400
            TabIndex        =   26
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   720
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
            MICON           =   "gruposdenomina.frx":2BB0
            PICN            =   "gruposdenomina.frx":2BCC
            PICH            =   "gruposdenomina.frx":2E62
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Campos Seleccionados"
            Height          =   195
            Left            =   3240
            TabIndex        =   28
            Top             =   240
            Width           =   1665
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Campos Disponibles"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Conceptos"
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   5415
         Begin VB.ListBox List1 
            Height          =   1815
            ItemData        =   "gruposdenomina.frx":30C1
            Left            =   120
            List            =   "gruposdenomina.frx":30C3
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   2055
         End
         Begin VB.ListBox List2 
            Height          =   1815
            ItemData        =   "gruposdenomina.frx":30C5
            Left            =   3240
            List            =   "gruposdenomina.frx":30C7
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
         Begin ChamaleonButton.ChameleonBtn BtnAnterior 
            Height          =   375
            Left            =   2400
            TabIndex        =   18
            ToolTipText     =   "Moverse la Registro Anterior"
            Top             =   1440
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
            MICON           =   "gruposdenomina.frx":30C9
            PICN            =   "gruposdenomina.frx":30E5
            PICH            =   "gruposdenomina.frx":337A
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
            Left            =   2400
            TabIndex        =   19
            ToolTipText     =   "Moverse la Registro Siguiente"
            Top             =   720
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
            MICON           =   "gruposdenomina.frx":35D6
            PICN            =   "gruposdenomina.frx":35F2
            PICH            =   "gruposdenomina.frx":3888
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
            BackStyle       =   0  'Transparent
            Caption         =   "Conceptos Disponibles"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conceptos Seleccionados"
            Height          =   195
            Left            =   3240
            TabIndex        =   20
            Top             =   240
            Width           =   1860
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   52756483
         UpDown          =   -1  'True
         CurrentDate     =   40179
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "gruposdenomina.frx":3AE7
         Left            =   2040
         List            =   "gruposdenomina.frx":3B33
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPickerDeCalculo 
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy"
         Format          =   52756483
         CurrentDate     =   40060
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "En base al año:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de Pago"
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de grupo de nómina"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmGruposNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BdGrupo As New ADODB.Recordset
Dim BdGrupo1 As New ADODB.Recordset
Dim bdgrupo2 As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim regnuevo
Dim Cambio
Dim IdGrupo
Dim CSql As String


Sub Cargar_Periodos()

Call Calcular_Periodos(Format(DTPicker1.Value, "yyyy"))

Combo1.Clear

For i = 0 To 23
    Combo1.AddItem "Período " & PyFs(i, 0) & ": " & PyFs(i, 1) & " => " & PyFs(i, 2)
    Combo1.ItemData(Combo1.NewIndex) = PyFs(i, 0)
Next i

End Sub

Sub Carga_Campos()
Dim RsCargarCampos As New ADODB.Recordset
List3.Clear
List4.Clear

CSql = "select * from CamposDeNomina WHERE activo=1"
Set RsCargarCampos = CrearRS(CSql)

If Not RsCargarCampos.EOF Then
    RsCargarCampos.MoveFirst
    Do While Not RsCargarCampos.EOF
        List3.AddItem RsCargarCampos.Fields("campo")
        List3.ItemData(List3.NewIndex) = RsCargarCampos.Fields("IdCampoNomina")
        RsCargarCampos.MoveNext
    Loop
End If
RsCargarCampos.Close

End Sub

Sub Carga_Conceptos()
Dim RsCargarConceptos As New ADODB.Recordset
List1.Clear
List2.Clear

CSql = "select * from concepto"
Set RsCargarConceptos = CrearRS(CSql)

If Not RsCargarConceptos.EOF Then
    RsCargarConceptos.MoveFirst
    Do While Not RsCargarConceptos.EOF
        List1.AddItem RsCargarConceptos.Fields("descripcion")
        List1.ItemData(List1.NewIndex) = RsCargarConceptos.Fields("idconcepto")
        RsCargarConceptos.MoveNext
    Loop
End If

RsCargarConceptos.Close

End Sub
Sub cargagrupo()

If BdGrupo1.RecordCount <> 0 Then
    regnuevo = 0
    Text1.Text = BdGrupo1.Fields("descripcion")
    IdGrupo = Format(BdGrupo1.Fields("id_grupo"), "0000")
    Label4.Caption = IdGrupo
    DTPicker1.Value = Format(BdGrupo1.Fields("fecha_prox_gen"), "dd/MM/yyyy")
    
    ' Ciclo que selecciona el período de la nómina
    For w = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(w) = BdGrupo1.Fields("periodo") Then
            Combo1.ListIndex = w
            Exit For
        End If
    Next w
            
    Call Carga_Conceptos
    Call Carga_Campos
    
    CSql = "select * from relacion_grupo where id_grupo = " & IdGrupo
    Set BdGrupo = CrearRS(CSql)
    
    If BdGrupo.RecordCount > 0 Then
        BdGrupo.MoveFirst
        Do While Not BdGrupo.EOF
            For w = 0 To List1.ListCount - 1
                If List1.ItemData(w) = BdGrupo.Fields("idconcepto") Then
                    List2.AddItem List1.List(w)
                    List2.ItemData(List2.NewIndex) = List1.ItemData(w)
                    List1.RemoveItem (w)
                    Exit For
                End If
            
            Next w
            BdGrupo.MoveNext
        Loop
    Else
        BdGrupo.Close
    End If
    
    CSql = "select * from relacion_campo where id_grupo = " & IdGrupo
    Set BdGrupo = CrearRS(CSql)
    
    If BdGrupo.RecordCount > 0 Then
        BdGrupo.MoveFirst
        Do While Not BdGrupo.EOF
            For w = 0 To List3.ListCount - 1
                If List3.ItemData(w) = BdGrupo.Fields("id_campo") Then
                    List4.AddItem List3.List(w)
                    List4.ItemData(List4.NewIndex) = List3.ItemData(w)
                    List3.RemoveItem (w)
                    Exit For
                End If
            
            Next w
            BdGrupo.MoveNext
        Loop
    Else
        BdGrupo.Close
    End If
Else
    regnuevo = 1
End If
End Sub

Private Sub BtnGuardar_Click()
Dim NuevoId
Dim ValCampoNom
Dim IdTemp
Dim IdTemp2
Dim Band As Boolean
Dim AlmId(0 To 40) As Integer
Dim Punto1 As Label
Dim ListCont As Integer
Dim i As Integer

ListCont = 0

If Text1.Text = "" Then MsgBox "Ingrese un nombre para el nuevo grupo de nómina!", vbExclamation + vbOKOnly, "Faltan Datos!": Exit Sub

If Trim(List2.List(0)) = "" And Trim(List4.List(0)) = "" Then
    MsgBox "Debe seleccionar por lo menos un CONCEPTO o un CAMPO NOMINA para poder guardar!", vbExclamation + vbOKCancel, "Operación Fallida"
    Exit Sub
End If

If Combo1.ListIndex = -1 Then
    MsgBox "Debe seleccionar un período!", vbExclamation + vbOKOnly, "Error"
    Combo1.SetFocus
    Exit Sub
End If

Select Case regnuevo
Case Is = 0 'Actualiza
    If Cambio = 1 Then
        CSql = "Update grupo set descripcion ='" & Text1.Text & "',periodo=" & PyFs(Combo1.ListIndex, 0) & _
            ", fecha_prox_gen='" & PyFs(Combo1.ListIndex, 1) & "',fecha_prox_gen2='" & _
            PyFs(Combo1.ListIndex, 2) & "' WHERE id_grupo = " & IdGrupo
        Set BdGrupo = CrearRS(CSql)
        MsgBox "Registro Actualizado Satisfactoriamente", vbOKOnly, "Guardado"
    Else
        MsgBox "no han ocurrido cambios en este registro", vbInformation + vbOKOnly, "Información"
        Exit Sub
    End If
Case Is = 1 'Agrega
    If Cambio = 1 Then
        
        CSql = "SELECT MAX(id_grupo)+1 as NuevoId FROM grupo"
        Set bdgrupo2 = CrearRS(CSql)
        
        If Not IsNull(bdgrupo2.Fields("NuevoId")) Then
            IdGrupo = bdgrupo2.Fields("NuevoId")
        Else
            IdGrupo = "1"
        End If
        bdgrupo2.Close
        
        CSql = "insert into grupo(Id_Grupo,descripcion, iduser, fecha_grupo, fecha_prox_gen, periodo,activo) " & _
               " values(" & IdGrupo & ",'" & Text1.Text & "'," & IdUser & ",'" & Format(Now, "dd/mm/yyyy") & "','" & _
               Format(DTPicker1.Value, "dd/mm/yyyy") & "'," & Combo1.ItemData(Combo1.ListIndex) & ",1)"
        Set BdGrupo = CrearRS(CSql)
        
        MsgBox "Registro agregado satisfactoriamente", vbOKOnly, "Operación Exitosa!"
    End If
End Select

'' mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
'resp = MsgBox("Se procedera a aplicar los cambios a los registros de los Empleados para el area de " & Text1.Text & _
'        "Desea Continuar?", vbQuestion + vbYesNo, "Confirmar cambios!")
'' mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
'
'If resp = vbYes Then
'
'    CSql = "Select * From Empleados Where Activo=1 and Id_Grupo= " & IdGrupo
'    Set BdGrupo = CrearRS(CSql)
'
'    If BdGrupo.RecordCount <> 0 Then
'        BdGrupo.MoveFirst
'
'    While Not BdGrupo.EOF
'        IdTemp = BdGrupo.Fields("Id_Grupo").Value
'        IdTemp2 = BdGrupo.Fields("IdEmpleado").Value
'Punto1:
'        ' Obtiene el Nuevo Id para agregarlo a la tabla "CamposDelTrabajador"
'        CSql = "SELECT MAX(id)+1 as NuevoId FROM CamposDelTrabajador"
'        Set bdgrupo2 = CrearRS(CSql)
'
'        If Not IsNull(bdgrupo2.Fields("NuevoId")) Then
'            NuevoId = bdgrupo2.Fields("NuevoId").Value
'        Else
'            NuevoId = "1"
'        End If
'
'        ' Busca el Empleado
'        CSql = "SELECT  CamposDelTrabajador.*, Grupo.Id_Grupo FROM Relacion_Campo INNER JOIN " & _
'            " Grupo ON Relacion_Campo.Id_Grupo = Grupo.Id_Grupo INNER JOIN " & _
'            " CamposDelTrabajador ON Relacion_Campo.Id_Campo = CamposDelTrabajador.IdCampoNomina " & _
'            " Where (CamposDelTrabajador.IdEmpleado = " & IdTemp2 & ") And (Grupo.Id_Grupo = " & IdGrupo & ") AND CamposDelTrabajador.Tipo='CA'"
'
'        Set bdgrupo2 = CrearRS(CSql)
'
'        If bdgrupo2.RecordCount <> 0 And Not IsNull(bdgrupo2.Fields(0)) Then
'
'        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'            ' ELIMINA los CAMPOS del Empleado segun el Grupo Anterior, si tiene campos q no pertenecen al Grupo
'            ' entonces los mantiene en el registro de ese empleado
'            bdgrupo2.MoveFirst
'            While Not bdgrupo2.EOF
'                'Id_Grupo
'                If Val(bdgrupo2.Fields("Id_Grupo").Value) <> IdGrupo Then
'                    CSql = "DELETE CamposDelTrabajador WHERE IdEmpleado=" & IdTemp2 & " AND IdCampoNomina=" & Val(bdgrupo2.Fields("IdCampoNomina").Value) & " AND Tipo='CA'"
'                    Set RsTemp = CrearRS(CSql)
'                End If
'                bdgrupo2.MoveNext
'            Wend
'        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'
'        ' Busca de nuevo el Empleado para ver si aun contiene campos
'            CSql = "SELECT  CamposDelTrabajador.* FROM Relacion_Campo INNER JOIN " & _
'                " Grupo ON Relacion_Campo.Id_Grupo = Grupo.Id_Grupo INNER JOIN " & _
'                " CamposDelTrabajador ON Relacion_Campo.Id_Campo = CamposDelTrabajador.IdCampoNomina " & _
'                " Where (CamposDelTrabajador.IdEmpleado = " & IdTemp2 & ") And (Grupo.Id_Grupo = " & IdGrupo & ") AND CamposDelTrabajador.Tipo='CA'"
'
'            Set bdgrupo2 = CrearRS(CSql)
'
'        ' Recorre la Lista4 para saber si el Empleado con IdTemp2 tiene el campo, si no, entonces lo agrega
'            For cont = 0 To List4.ListCount - 1
'                Band = False
'
'                ' Condicional para ver si el empleado, despues de haber BORRARDO en el Ciclo WHILE anterior los campos,
'                ' sigue teniendo campos, si ese es el caso, entonces, verifica que los que voy a agregar no esten ya agregados
'                If bdgrupo2.RecordCount <> 0 Then
'                    bdgrupo2.MoveFirst
'                    While Not bdgrupo2.EOF
'                        ' ciclo para comparar cada elemento de la lista4 con los campos del trabajador, si es igual
'                        ' entonces BAND=TRUE lo cual significa que SI EXISTE EL CAMPO y por lo tanto no lo agregara
'                        If Val(bdgrupo2.Fields("IdCampoNomina").Value) = Val(List4.ItemData(cont)) Then
'                            Band = True
'                            bdgrupo2.MoveLast
'                            bdgrupo2.MoveNext
'                        Else
'                            bdgrupo2.MoveNext
'                        End If
'                    Wend
'
'                    ' Si el campo NO EXISTE, la variable BAND=FALSE, por lo tanto AGREGO el campo al trabajador
'                    If Band = False Then
'                        dat1 = List4.ItemData(cont)
'                        CSql = "SELECT * FROM CamposDeNomina Where IdCampoNomina=" & dat1
'                        Set RsTemp = CrearRS(CSql)
'
'                        ValCampoNom = RsTemp.Fields("Predeterminado").Value
'                        CSql = "insert into CamposDelTrabajador(Id,IdCampoNomina, ValorN, IdEmpleado,Tipo) " & _
'                            " values(" & NuevoId & "," & dat1 & "," & ValCampoNom & "," & IdTemp2 & ", 'CA')"
'                        Set RsTemp = CrearRS(CSql)
'
'                        CSql = "SELECT MAX(id)+1 as NuevoId FROM CamposDelTrabajador"
'                        Set RsTemp = CrearRS(CSql)
'
'                        If Not IsNull(RsTemp.Fields("NuevoId")) Then
'                            NuevoId = RsTemp.Fields("NuevoId").Value
'                        Else
'                            NuevoId = "1"
'                        End If
'                    End If
'                Else ' Si el Empleado NO TIENE CAMPOS entonces los crea
'                    dat1 = List4.ItemData(tw)
'
'                    CSql = "SELECT * FROM CamposDeNomina Where IdCampoNomina=" & dat1
'                    Set RsTemp = CrearRS(CSql)
'
'                    ValCampoNom = RsTemp.Fields("Predeterminado").Value
'
'                    CSql = "insert into CamposDelTrabajador(Id,IdCampoNomina, ValorN, IdEmpleado,Tipo) " & _
'                        " values(" & NuevoId & "," & dat1 & "," & ValCampoNom & "," & IdTemp2 & ",'CA')"
'                    Set RsTemp = CrearRS(CSql)
'
'                    CSql = "SELECT MAX(id)+1 as NuevoId FROM CamposDelTrabajador"
'                    Set RsTemp = CrearRS(CSql)
'
'                    If Not IsNull(RsTemp.Fields("NuevoId")) Then
'                        NuevoId = RsTemp.Fields("NuevoId").Value
'                    Else
'                        NuevoId = "1"
'                    End If
'                End If
'
'                ' Sentencia que lee de nuevo los registros del empleado IDTEMP2, ya que pudo haber sufrido cambios
'                CSql = "SELECT * FROM CamposDelTrabajador where IdEmpleado=" & IdTemp2 & " AND Tipo='CA'"
'                Set bdgrupo2 = CrearRS(CSql)
'            Next cont
'            'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'        Else    ' Si el Empleado NO TIENE CAMPOS entonces los crea
'            If List4.ListCount <> 0 Then
'                dat1 = List4.ItemData(ListCont)
'
'                CSql = "SELECT * FROM CamposDeNomina Where IdCampoNomina=" & dat1
'                Set RsTemp = CrearRS(CSql)
'
'                ValCampoNom = RsTemp.Fields("Predeterminado").Value
'
'                CSql = "insert into CamposDelTrabajador(Id,IdCampoNomina, ValorN, IdEmpleado,Tipo) " & _
'                    " values(" & NuevoId & "," & dat1 & "," & ValCampoNom & "," & IdTemp2 & ",'CA')"
'                Set RsTemp = CrearRS(CSql)
'
'                ListCont = ListCont + 1
'                If List4.ListCount <> ListCont Then
'                    GoTo Punto1
'                End If
'            End If
'        End If
'        ListCont = 0
'        BdGrupo.MoveNext
'    Wend
'    End If
'End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Elimina TODAS las relaciones de la BD con respecto al Grupo  MMMMMMMMMMMMM
' para CREARLAS DE NUEVO de acuerdo a la lista2 y lista4       MMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "delete from relacion_grupo where [Id_grupo] = " & IdGrupo
Set BdGrupo = CrearRS(CSql)

CSql = "delete from relacion_campo where [Id_grupo] = " & IdGrupo
Set BdGrupo = CrearRS(CSql)

' ciclo para relacionar los CONCEPTOS con el GRUPO
For tw = 0 To List2.ListCount - 1

    CSql = "SELECT MAX(id_relacion)+1 as NuevoId FROM relacion_grupo"
    Set bdgrupo2 = CrearRS(CSql)
    
    If Not IsNull(bdgrupo2.Fields("NuevoId")) Then
        NuevoId = bdgrupo2.Fields("NuevoId")
    Else
        NuevoId = "1"
    End If
    
    dat1 = List2.ItemData(tw)
    CSql = "insert into relacion_grupo(Id_Relacion,id_grupo, idconcepto, iduser, fecha_relacion) " & _
        " values(" & NuevoId & "," & IdGrupo & "," & dat1 & "," & IdUser & ",'" & Format(Now, "DD/MM/YYYY") & "')"
    
    Set BdGrupo = CrearRS(CSql)
Next tw

' ciclo para relacionar los CAMPOS con el GRUPO
For tw = 0 To List4.ListCount - 1

    CSql = "SELECT MAX(id_relacion)+1 as NuevoId FROM relacion_campo"
    Set bdgrupo2 = CrearRS(CSql)
    
    If Not IsNull(bdgrupo2.Fields("NuevoId")) Then
        NuevoId = bdgrupo2.Fields("NuevoId")
    Else
        NuevoId = "1"
    End If
    
    dat1 = List4.ItemData(tw)
    CSql = "insert into relacion_campo(Id_Relacion,id_grupo, id_campo, iduser, fecha_relacion) " & _
        " values(" & NuevoId & "," & IdGrupo & "," & dat1 & "," & IdUser & ",'" & Format(Now, "dd/mm/yyyy") & "')"
    Set BdGrupo = CrearRS(CSql)
Next tw
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Cambio = 0
Form_Load

End Sub

Private Sub BtnNuevo_Click()
regnuevo = 1
Text1.Text = ""
Cambio = 0
Label4.Caption = "Nuevo"

Call Carga_Conceptos
Call Carga_Campos
Text1.SetFocus
End Sub

Private Sub BtnSiguiente_Click()
If List1.ListIndex >= 0 Then
    List2.AddItem List1.List(List1.ListIndex)
    List2.ItemData(List2.NewIndex) = List1.ItemData(List1.ListIndex)
    List1.RemoveItem (List1.ListIndex)
    Cambio = 1
Else
    Msg = "No Ha seleccionado ningun elemento de la lista de CONCEPTOS DISPONIBLES"
    List1.SetFocus
    MsgBox Msg
End If
End Sub

Private Sub BtnAnexarCampo_Click()
If List3.ListIndex >= 0 Then
    List4.AddItem List3.List(List3.ListIndex)
    List4.ItemData(List4.NewIndex) = List3.ItemData(List3.ListIndex)
    List3.RemoveItem (List3.ListIndex)
    Cambio = 1
Else
    Msg = "No Ha seleccionado ningun elemento de la lista de CAMPOS DISPONIBLES"
    MsgBox Msg
    List3.SetFocus
End If
End Sub

Private Sub BtnAnterior_Click()
If List2.ListIndex >= 0 Then
    List1.AddItem List2.List(List2.ListIndex)
    List1.ItemData(List1.NewIndex) = List2.ItemData(List2.ListIndex)
    List2.RemoveItem (List2.ListIndex)
    Cambio = 1
Else
    Msg = "No Ha seleccionado ningun elemento de la lista de CONCEPTOS SELECCIONADOS"
    MsgBox Msg
    List2.SetFocus
End If
End Sub

Private Sub BtnAnterior1_Click()

If BdGrupo1.RecordCount <> 0 Then
    BdGrupo1.MovePrevious
    If BdGrupo1.BOF Then BdGrupo1.MoveLast
    Call cargagrupo
Else
    MsgBox "no existe registros disponibles", vbExclamation + vbOKOnly, "Noy hay registros!"
End If
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnExcluirCampo_Click()
If List4.ListIndex >= 0 Then
    List3.AddItem List4.List(List4.ListIndex)
    List3.ItemData(List3.NewIndex) = List4.ItemData(List4.ListIndex)
    List4.RemoveItem (List4.ListIndex)
    Cambio = 1
Else
    Msg = "No ha seleccionado ningun elemento de la lista de CAMPOS SELECCIONADOS"
    MsgBox Msg
    List4.SetFocus
End If
End Sub

Private Sub BtnSiguiente1_Click()
    
If BdGrupo1.RecordCount <> 0 Then
    BdGrupo1.MoveNext
    If BdGrupo1.EOF Then BdGrupo1.MoveFirst
    Call cargagrupo
Else
    MsgBox "no existe registros disponibles", vbExclamation + vbOKOnly, "Noy hay registros!"
End If

End Sub

Private Sub Combo1_Click()
Cambio = 1
End Sub

Private Sub DTPicker1_Change()
Cambio = 1
Cargar_Periodos
End Sub

Private Sub DTPicker1_Click()
Cambio = 1
Cargar_Periodos
End Sub

Private Sub Form_Load()
Centrar Me
DTPicker1.Value = Now
Cargar_Periodos
Carga_Conceptos
Carga_Campos

CSql = "Select * From Grupo where Activo=1"
Set BdGrupo1 = CrearRS(CSql)

If BdGrupo1.RecordCount <> 0 Then
    BdGrupo1.MoveFirst
    Call cargagrupo
End If

Cambio = 0

End Sub
Private Sub Text1_Change()
Cambio = 1
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            List1.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyLeft
            DTPicker1.SetFocus
        Case vbKeyDown
            List2.SetFocus
    End Select
End If
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo1.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            List1.SetFocus
    End Select
End If
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnSiguiente.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyRight
            BtnSiguiente.SetFocus
        Case vbKeyDown
            BtnNuevo.SetFocus
    End Select
End If
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnNuevo.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyLeft
            BtnSiguiente.SetFocus
        Case vbKeyDown
            BtnNuevo.SetFocus
    End Select
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DTPicker1.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            DTPicker1.SetFocus
    End Select
End If
End Sub

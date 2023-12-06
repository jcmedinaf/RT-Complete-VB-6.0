VERSION 5.00
Begin VB.Form Form42 
   Caption         =   "Agregar Campo Nomina"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form42"
   ScaleHeight     =   3915
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Caracteristicas del Cancer"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5415
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form42.frx":0000
         Left            =   1200
         List            =   "Form42.frx":000D
         TabIndex        =   11
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3360
         Picture         =   "Form42.frx":0029
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   1200
         Picture         =   "Form42.frx":044F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   4080
         Picture         =   "Form42.frx":09D9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   2640
         Picture         =   "Form42.frx":0F63
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   0
         Left            =   1920
         Picture         =   "Form42.frx":11E2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Blanquear"
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   120
      Picture         =   "Form42.frx":176C
      Top             =   120
      Width           =   3780
   End
End
Attribute VB_Name = "Form42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDNon As New ADODB.Recordset
Dim BDNon1 As New ADODB.Recordset

Private Sub Command1_Click()

CSql = "Insert into camposdenomina(campo,tipo) VALUES('" & Text1.Text & "'," & Combo1.ListIndex & ")"
        BDNon1.Open CSql, CADENA
        msg = "Registro Agregado satisfactoriamente"
        MsgBox msg, vbOKOnly
Call blanqueo
BDNon1.Close
Call Form_Load
Exit Sub

End Sub

Sub blanqueo()
Label1.Caption = ""
Text1.Text = ""
Combo1.ListIndex = -1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
Call blanqueo
End Sub

Private Sub Command6_Click()
BDNon1.MovePrevious
Call cargadato
End Sub

Private Sub Command7_Click()
BDNon1.MoveNext
Call cargadato

End Sub
Private Sub Form_Load()

CSql = "SELECT * FROM camposdenomina"
BDNon.CursorType = adOpenKeyset
BDNon.LockType = adLockOptimistic
BDNon.CursorLocation = adUseClient
BDNon.Open CSql, CADENA, , , adCmdText


     
CSql = "select * from camposdenomina"
        Dim BDNom As New ADODB.Recordset
        BDNom.Open CSql, CADENA
     
        Nom = Format(BDNom.Fields("Idcamponomina"), "000#")
        
        BDNom.Close
        Label1.Caption = Nom
        
Call cargadato

End Sub

Sub cargadato()

If BDNon.EOF Then
msg = "Llego al Final del Registro desea Volver al Principio?"
MsgBox msg
BDNon.MoveFirst
End If

If BDNon.BOF Then
    msg = "Llego al principio del registro"
    MsgBox msg
    BDNon.MoveLast
End If

If Trim(BDNon.Fields("Campo")) <> "" Then Text1.Text = BDNon.Fields("Campo")
    Nom = Format(BDNon.Fields("Idcamponomina"), "000#")
    Label1.Caption = Nom
                    
                   For t = 0 To Combo1.ListCount - 1
                          If Combo1.ItemData(t) = BDNon.Fields("tipo") Then
                          Combo1.ListIndex = t
                          Exit For
                          End If
                    Next t

End Sub



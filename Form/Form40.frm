VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form40 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form40"
   ClientHeight    =   8475
   ClientLeft      =   7185
   ClientTop       =   2400
   ClientWidth     =   8580
   LinkTopic       =   "Form40"
   ScaleHeight     =   8475
   ScaleWidth      =   8580
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   7800
      Picture         =   "Form40.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8295
      Begin VB.CommandButton Command10 
         Caption         =   "Evaluar"
         Height          =   375
         Left            =   5880
         TabIndex        =   31
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Lista de Conceptos"
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   360
         Width           =   1935
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1095
         Left            =   600
         TabIndex        =   28
         Top             =   5040
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"Form40.frx":0582
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Añadir"
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   4440
         Width           =   1215
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   600
         TabIndex        =   25
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Añadir"
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   600
         TabIndex        =   21
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Añadir"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   600
         TabIndex        =   19
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   3960
         Picture         =   "Form40.frx":0604
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guardar"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   3240
         Picture         =   "Form40.frx":0883
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Blanquear"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   4680
         Picture         =   "Form40.frx":0E0D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   2520
         Picture         =   "Form40.frx":1397
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6360
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vacaciones"
         Height          =   375
         Left            =   6840
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Utilidades"
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Prestamos"
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Prestaciones"
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Acumulado"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form40.frx":1921
         Left            =   1680
         List            =   "Form40.frx":1934
         TabIndex        =   6
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form40.frx":1975
         Left            =   1680
         List            =   "Form40.frx":1985
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Constante"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Campo del trabajador"
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Utilizar"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Generar"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descripción"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Codigo"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   0
      Picture         =   "Form40.frx":19B1
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDCOncep As New ADODB.Recordset
Dim BDCOncep1 As New ADODB.Recordset
Dim BDCOnc As New ADODB.Recordset
Dim BDCOns As New ADODB.Recordset
Dim BDCOcamp As New ADODB.Recordset
Dim BDCO As New ADODB.Recordset
Dim BDCO1 As New ADODB.Recordset
Dim IDcon


Private Sub Command1_Click()
CSql = "insert into Concepto(Formula,Descripcion,tipo,Genera,Acumulado,Prestaciones,Prestamo,Utilidades,Vacaciones) values('" & RichTextBox1.Text & "','" & Text2.Text & "'," & Combo1.ListIndex & "," & Combo2.ListIndex & "," & Check1.Value & "," & Check2.Value & "," & Check3.Value & "," & Check4.Value & "," & Check5.Value & ")"
CSql: BDCOncep.Open CSql, CADENA
msg = "El Concepto Sea Agregado Satisfactoriamente"
        MsgBox msg, vbOKOnly

End Sub

Private Sub Command10_Click()
msg = "Indique el ID de un trabajador para evaluar el concepto"
u = InputBox(msg, "ID Trabajador", 2)
If Trim(u) = "" Then Exit Sub
formula = RichTextBox1.Text


  sCaracter = "Concepto("
  stmp = ""
  t = 0
  For h = 1 To Len(formula)
        cars = Mid$(formula, h, 9)
        If UCase(sCaracter) = UCase(cars) Then
            formula1 = Mid(formula, h + 8, Len(formula))
            For w = 1 To Len(formula1)
            d = Mid(formula1, w, 1)
            If d = ";" Then t = 1
            
            If t = 1 And Not d = ";" And Not d = ")" Then codcon = codcon & d
            If d = ")" Then Exit For
            Next w
            
            
            stmp = stmp & cars
            h = h + 8
        End If
  Next h
  MsgBox codcon
End Sub

Private Sub Command2_Click()
Call blanqueo

End Sub
Sub blanqueo()

RichTextBox1.Text = ""
  Text2.Text = ""
  Label6.Caption = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0

Combo1.ListIndex = -1
Combo2.ListIndex = -1
Combo3.ListIndex = -1
Combo4.ListIndex = -1
Combo5.ListIndex = -1
                
End Sub

Private Sub Command3_Click()
cad = "Concepto(" & Mid(Combo3.List(Combo3.ListIndex), 1, 10) & ";" & Combo3.ItemData(Combo3.ListIndex) & ")"

RichTextBox1.Text = RichTextBox1.Text & cad
End Sub

Private Sub Command4_Click()
cad = "Campo(" & Mid(Combo4.List(Combo4.ListIndex), 1, 10) & ";" & Combo4.ItemData(Combo4.ListIndex) & ")"

RichTextBox1.Text = RichTextBox1.Text & cad
End Sub

Private Sub Command5_Click()
cad = "Constante(" & Mid(Combo5.List(Combo5.ListIndex), 1, 10) & ";" & Combo5.ItemData(Combo5.ListIndex) & ")"

RichTextBox1.Text = RichTextBox1.Text & cad
End Sub

Private Sub Command6_Click()

If Not (BDCO.BOF) Then BDCO.MovePrevious
If Not (BDCO.BOF) Then Call CargaConc

End Sub

Private Sub Command7_Click()

If Not (BDCO.EOF) Then BDCO.MoveNext
If Not (BDCO.EOF) Then Call CargaConc

End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Command9_Click()
Form41.Show 1
End Sub

Private Sub Form_Load()

CSql = "SELECT * FROM Concepto"
BDCO.CursorType = adOpenKeyset
BDCO.LockType = adLockOptimistic
BDCO.CursorLocation = adUseClient
BDCO.Open CSql, CADENA, , , adCmdText

    CSql = "select * from concepto"
        Dim BDCOnc As New ADODB.Recordset
        BDCOnc.Open CSql, CADENA
    BDCOnc.MoveFirst
    Do While Not BDCOnc.EOF
        If Not IsNull(BDCOnc.Fields("DESCRIPCION")) Then Combo3.AddItem BDCOnc.Fields("DESCRIPCION")
    Combo3.ItemData(Combo3.NewIndex) = BDCOnc.Fields("IDCONCEPTO")
    BDCOnc.MoveNext
        Loop
    BDCOnc.Close

    CSql = "select * from camposdenomina"
        Dim BDCOcamp As New ADODB.Recordset
        BDCOcamp.Open CSql, CADENA
    BDCOcamp.MoveFirst
    Do While Not BDCOcamp.EOF
        If Not IsNull(BDCOcamp.Fields("campo")) Then Combo4.AddItem BDCOcamp.Fields("campo")
    Combo4.ItemData(Combo4.NewIndex) = BDCOcamp.Fields("IDCamponomina")
    BDCOcamp.MoveNext
        Loop
    BDCOcamp.Close

    CSql = "select * from constantesdenomina"
        Dim BDCOns As New ADODB.Recordset
        BDCOns.Open CSql, CADENA
    BDCOns.MoveFirst
    Do While Not BDCOns.EOF
        If Not IsNull(BDCOns.Fields("Descripcion")) Then Combo5.AddItem BDCOns.Fields("Descripcion")
        Combo5.ItemData(Combo5.NewIndex) = BDCOns.Fields("IDConstante")
        BDCOns.MoveNext
    Loop
    BDCOns.Close

    CSql = "select * from concepto"
        Dim BDCO1 As New ADODB.Recordset
        BDCO1.Open CSql, CADENA
     
        Concep = Format(BDCO1.Fields("Idconcepto"), "000#")
        
        BDCO1.Close
        Label6.Caption = Concep
        
    End Sub

Sub CargaConc()
If Not BDCO.EOF Then
If IsNull(BDCO.Fields(3)) Then RichTextBox1.Text = "" Else RichTextBox1.Text = BDCO.Fields(3)
If IsNull(BDCO.Fields(1)) Then Text2.Text = "" Else Text2.Text = BDCO.Fields(1)

    If IsNull(BDCO.Fields(9)) Then Check1.Value = 0 Else Check1.Value = BDCO.Fields(9)
    If IsNull(BDCO.Fields(8)) Then Check2.Value = 0 Else Check2.Value = BDCO.Fields(8)
    If IsNull(BDCO.Fields(6)) Then Check3.Value = 0 Else Check3.Value = BDCO.Fields(6)
    If IsNull(BDCO.Fields(10)) Then Check4.Value = 0 Else Check4.Value = BDCO.Fields(10)
    If IsNull(BDCO.Fields(7)) Then Check5.Value = 0 Else Check5.Value = BDCO.Fields(7)
    Concep = Format(BDCO.Fields("Idconcepto"), "000#")
    Label6.Caption = Concep
    IDcon = Idconcepto
    For t = 0 To Combo1.ListCount - 1
                          If Combo1.ItemData(t) = BDCO.Fields("tipo") Then
                          Combo1.ListIndex = t
                          Exit For
                          End If
                    Next t
                    
                    For t = 0 To Combo2.ListCount - 1
                          If Combo2.ItemData(t) = BDCO.Fields("Genera") Then
                          Combo2.ListIndex = t
                          Exit For
                          End If
                    Next t
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
BDCO.Close
End Sub

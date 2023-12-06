VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form26 
   Caption         =   "Form26"
   ClientHeight    =   6390
   ClientLeft      =   5730
   ClientTop       =   690
   ClientWidth     =   10005
   LinkTopic       =   "Form26"
   ScaleHeight     =   6390
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabHeight       =   520
      TabCaption(0)   =   "Campo 1"
      TabPicture(0)   =   "registrodiario.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Campo 2"
      TabPicture(1)   =   "registrodiario.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Campo 3"
      TabPicture(2)   =   "registrodiario.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Campo 4"
      TabPicture(3)   =   "registrodiario.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Campo 5"
      TabPicture(4)   =   "registrodiario.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Campo 6"
      TabPicture(5)   =   "registrodiario.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).ControlCount=   0
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bd58 As New ADODB.Recordset
Dim bd59 As New ADODB.Recordset
Dim l As Integer


' En construccion
Private Sub Form_Load()
'MsgBox IdPac1
'MsgBox CSql
'MsgBox CNn
    
CSql = "select * from tecnica2 where idpaciente = " & IdPac1
    bd59.CursorLocation = adUseClient
    bd59.Open CSql, cadena
    l = bd59.RecordCount


For r = 0 To l - 1
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False
Label1(r).Visible = False

Next r
    SSTab1.Tabs = l
    SSTab1.TabsPerRow = l / 2

For e = 0 To l - 1
Label1(e).Caption = "Campo #   :" & bd59.Fields("campo")
Label1(e).Caption = "Col Upper :" & bd59.Fields("upper")
Label1(e).Caption = "Col Lower :" & bd59.Fields("lower")
'label4(E).Caption = bd59.Fields("campo")
'label5(E).Caption = bd59.Fields("campo")
'Label6(E).Caption = bd59.Fields("campo")
'Label7(E).Caption = bd59.Fields("campo")
'Label8(E).Caption = bd59.Fields("campo")
'Label9(E).Caption = bd59.Fields("campo")
'Label10(E).Caption = bd59.Fields("campo")
'Label11(E).Caption = bd59.Fields("campo")
'Label12(E).Caption = bd59.Fields("campo")
'Label13(E).Caption = bd59.Fields("campo")

Next e
    
    bd59.Close

Exit Sub
CSql = "select * from tratam_dado where idpaciente = " & IdPac1
    bd58.CursorLocation = adUseClient
    bd58.Open CSql, cadena
    MsgBox bd58.RecordCount
    bd58.Close
End Sub

Private Sub Label2_Click(Index As Integer)

End Sub


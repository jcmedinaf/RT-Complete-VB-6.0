VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "finalizar"
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   5280
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ver reg"
      Height          =   735
      Left            =   7320
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "modificar reg"
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "borrar regis"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "añadir registro"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbcliente As Database
Dim pacientes

Sub muestracampo()
Text1.Text = paciente!Nombre
Text2.Text = paciente!idcodigo
Text1.Text = paciente!cedula

End Sub

Private Sub Command1_Click()
Dim np As String
Dim nisb As String

np = UCase(InputBox$("Nuevo Paciente", "altas"))
nisb = UCase(InputBox$("Nuevo nisb>", "altas"))
 paciente.addNew
 paciente!Nombre = np
 paciente!cedula = nisbn
 paciente!idcodigo = 15410
 paciente.Update
 
 End Sub

Private Sub Command2_Click()
Dim comfirmar As Integer
Dim buscar As String
Dim observacion As String
Dim Nombre As String

buscar = UCase(InputBox$("introducas cedula del paciente", "buscar un registro"))

If busca <> "" Then

cedula = "cedula like `*" & buscar & "*´"







End Sub

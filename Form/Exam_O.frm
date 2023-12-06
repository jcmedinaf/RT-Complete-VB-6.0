VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form16"
   ClientHeight    =   7830
   ClientLeft      =   7650
   ClientTop       =   3615
   ClientWidth     =   12615
   LinkTopic       =   "Form16"
   ScaleHeight     =   7830
   ScaleWidth      =   12615
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   12000
      Picture         =   "Exam_O.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Medico Residente"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   3255
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   7080
      Picture         =   "Exam_O.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3840
      Picture         =   "Exam_O.frx":0F84
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Antecedente de Importrancia"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12375
      Begin VB.TextBox Text20 
         Height          =   495
         Left            =   1560
         TabIndex        =   39
         Top             =   4200
         Width           =   6375
      End
      Begin VB.TextBox Text19 
         Height          =   495
         Left            =   1560
         TabIndex        =   37
         Top             =   3480
         Width           =   6375
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   7920
         TabIndex        =   35
         Text            =   "Text18"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   3120
         TabIndex        =   31
         Text            =   "Text17"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Text            =   "Text16"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   5880
         TabIndex        =   25
         Text            =   "Text15"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   5880
         TabIndex        =   24
         Text            =   "Text14"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   5880
         TabIndex        =   23
         Text            =   "Text9"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Text            =   "         años        "
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Text            =   "Años"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Enfermedades Asociadas"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Antecedentes"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Actividad Fisica"
         Height          =   375
         Left            =   7920
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Quimioterapia"
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Intervenciones Quirurgicas"
         Height          =   375
         Left            =   3120
         TabIndex        =   32
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Hormonas"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Horas que duerme"
         Height          =   495
         Left            =   3120
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Sueños"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sueños"
         Height          =   15
         Left            =   360
         TabIndex        =   22
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Alcohol"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cigarrillo"
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Menarquia"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Flujo"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclo"
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Embarazo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cafe"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Menopausia"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   120
      Picture         =   "Exam_O.frx":150E
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

CSQL = "Insert into EXAMEN(DENSIDAD,LEUCOCITUS,NITRITUS,CUERPOS,BACTERIAS,CELULAS,OTROS) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text10.Text & "')"
CSQL = "Insert into EXAMEN(REACCION,ELEMENTOS,OTROS1,MEDICO) VALUES('" & Text7.Text & "','" & Text8.Text & "','" & Text11.Text & "','" & Text12.Text & "')"
        Dim BD As New adodb.Recordset
        BD.Open CSQL, CADENA
        msg = "Registro Agregado Satisfactoriamente"
        MsgBox msg, vbOKOnly
        'Call blanqueo
Exit Sub

noguardA:
    msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox msg, vbOKOnly, "Error al Guardar"
    msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox msg, vbOKOnly, "Error al Guardar"
    Exit Sub


i:
        
        BD.Open CSQL, CADENA
        msg = "Registro Agregado Satisfactoriamente"
        MsgBox msg, vbOKOnly
        'Call blanqueo
Exit Sub



End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Unload Me
End Sub


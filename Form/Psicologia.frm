VERSION 5.00
Begin VB.Form FrmPsicologia 
   Caption         =   "Form10"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   11805
   Begin VB.Frame Frame2 
      Caption         =   "Evaluación Clinica"
      Height          =   1455
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   4815
      Begin VB.TextBox Text8 
         Height          =   735
         Left            =   840
         TabIndex        =   29
         Text            =   "Text8"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label11 
         Caption         =   "Estado Genaral"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Paciente"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   9960
         Picture         =   "Psicologia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   2520
         TabIndex        =   13
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   5640
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   7200
         TabIndex        =   8
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   9720
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   9720
         TabIndex        =   1
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula"
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombres"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellidos"
         Height          =   255
         Left            =   6360
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Medico &Tratante"
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Inicio"
         Height          =   255
         Left            =   8400
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de C&ulminación"
         Height          =   255
         Left            =   8040
         TabIndex        =   20
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Medico &Remitente"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Sexo"
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nacimiento"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad"
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmPsicologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

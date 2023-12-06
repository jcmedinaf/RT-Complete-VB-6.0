VERSION 5.00
Begin VB.Form FrmValoracionGlobal 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valoración Global subjetiva del estado nutricional"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Examen Fisico"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   12855
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para cada opción especificar:  0 =Normal; 1=Leve; 2=Moderada; 3=Severa"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   5325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Historia Clinica"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Caption         =   "5) Enfermedad y su realcion en los requerimientos nutricionales"
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   12495
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "4) Capacidad Funcional"
         Height          =   2295
         Left            =   6480
         TabIndex        =   5
         Top             =   2640
         Width           =   6135
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "3) Sintomas Gastrointestinales de Duracion superior"
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   6135
         Begin VB.OptionButton Option17 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Dolor Abdominal"
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option16 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Disfagia"
            Height          =   255
            Left            =   1920
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option15 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Diarrea"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1800
            Width           =   1335
         End
         Begin VB.OptionButton Option14 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Vomitos"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1440
            Width           =   1335
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Anorexia"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton Option12 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Nauseas"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Ninguna"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "2) Cambios en el Aporte Dietético"
         Height          =   2295
         Left            =   6480
         TabIndex        =   3
         Top             =   240
         Width           =   6135
         Begin VB.OptionButton Option10 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Ayuno Casi Completo"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1920
            Width           =   2655
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Dieta Oral Liquida Exclusiva"
            Height          =   255
            Left            =   3000
            TabIndex        =   28
            Top             =   1560
            Width           =   2655
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Dieta Oral Triturada Insuficiente"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1560
            Width           =   2655
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Dieta Oral Triturada Suficiente"
            Height          =   255
            Left            =   3000
            TabIndex        =   26
            Top             =   1200
            Width           =   2655
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Dieta Oral Solida Insuficiente"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1080
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Si"
            Height          =   255
            Left            =   1080
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "No"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Semanas"
            Height          =   195
            Left            =   2040
            TabIndex        =   24
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duracion:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   810
            Width           =   690
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "1) Peso Corporal"
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6135
         Begin VB.OptionButton Option3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Disminucion"
            Height          =   255
            Left            =   3120
            TabIndex        =   17
            Top             =   1920
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Sin Cambios"
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   1920
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Aumento"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   4680
            TabIndex        =   13
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   5640
            TabIndex        =   19
            Top             =   1170
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kg."
            Height          =   195
            Left            =   2400
            TabIndex        =   18
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valoracion en las ultimas dos (2) semanas:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1560
            Width           =   2985
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje de peso habitual:"
            Height          =   195
            Left            =   2520
            TabIndex        =   12
            Top             =   1170
            Width           =   2025
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Kg:"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   1170
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Perdida en los ultimos seis (6) meses:"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso Habitual:"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   450
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "FrmValoracionGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

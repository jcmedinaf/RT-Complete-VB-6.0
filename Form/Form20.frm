VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edición Del Tecnico"
   ClientHeight    =   5460
   ClientLeft      =   3000
   ClientTop       =   2205
   ClientWidth     =   8220
   LinkTopic       =   "Form20"
   ScaleHeight     =   5460
   ScaleWidth      =   8220
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del Paciente"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7935
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   10
         Left            =   1680
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   11
         Left            =   5520
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   12
         Left            =   5520
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   13
         Left            =   5520
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   14
         Left            =   5520
         TabIndex        =   2
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   9
         Left            =   1680
         TabIndex        =   1
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dia"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tecnica y Sitio de Anatomia"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Energia (MV)"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prof o Iso (cm o %)"
         Height          =   495
         Left            =   480
         TabIndex        =   17
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Frac/Dia"
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Frac/Sem"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Frac/Total"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dosis/Frac Gy/frac"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dosis Total Gy"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   120
      Picture         =   "Form20.frx":0000
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Enum EACCION
    AGREGAR_REGISTRO = 0
    EDITAR_REGISTRO = 1
End Enum

Public IdRegistro As Long
Public ACCION As EACCION

Private Sub Command1_Click()

On Error GoTo ErrorSub
        
    ' Valida el Nombre que no este vacio
    ''''''''''''''''''''''''''''''''
    If Trim(Text1(1)) = "" Then
        MsgBox "El Nombre de registro no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(1).SetFocus
        Exit Sub
    
    ' Valida el Apellido
    ''''''''''''''''''''''''''''''''
    ElseIf Trim(Text1(2)) = "" Then
        MsgBox "El Apellido no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(2).SetFocus
        Exit Sub
    End If

    'Agrega el registro
    '''''''''''''''''''''''''''''''
    
    Select Case ACCION
        Case EDITAR_REGISTRO
        
            BD1(1) = Text1(1)
            BD1(2) = Text1(2)
            BD1(3) = Text1(3)
            'BD1(4) = Text1(4)
            BD1(5) = Text1(5)
            BD1(6) = Text1(6)
            BD1(7) = Text1(7)
            BD1(8) = Text1(8)
            BD1(9) = Text1(9)
            BD1(10) = Text1(10)
            BD1(11) = Text1(11)
            BD1(12) = Text1(12)
            BD1(13) = Text1(13)
            BD1(14) = Text1(14)
            BD1(16) = Text1(16)
            
        
        Case AGREGAR_REGISTRO
        
            BD1.AddNew
            BD1(1) = Text1(1)
            BD1(2) = Text1(2)
            BD1(3) = Text1(3)
            'BD1(4) = Text1(4)
            BD1(5) = Text1(5)
            BD1(7) = Text1(7)
            BD1(8) = Text1(8)
            BD1(9) = Text1(9)
            BD1(10) = Text1(10)
            BD1(11) = Text1(11)
            BD1(12) = Text1(12)
            BD1(13) = Text1(13)
            BD1(14) = Text1(14)
            BD1(16) = Text1(16)
           
            BD1(6) = CDate(Label3)
            
    End Select
    
    BD1.Update
    
    Unload Me
    Set Form20 = Nothing

Exit Sub
ErrorSub:
    MsgBox Err.Description

End Sub

Private Sub Command2_Click()
 Unload Me
End Sub


Private Sub Form_Load()
    If Me.ACCION = AGREGAR_REGISTRO Then
       Me.Caption = "Agregar nuevo registro"
    ElseIf Me.ACCION = EDITAR_REGISTRO Then
       Me.Caption = "Editar registro"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub


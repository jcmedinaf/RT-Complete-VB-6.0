VERSION 5.00
Begin VB.Form Form22 
   Caption         =   "Form22"
   ClientHeight    =   8145
   ClientLeft      =   6630
   ClientTop       =   1770
   ClientWidth     =   12690
   LinkTopic       =   "Form22"
   ScaleHeight     =   8145
   ScaleWidth      =   12690
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del Paciente"
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   12735
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   14
         Left            =   1920
         TabIndex        =   34
         Top             =   5160
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   12
         Left            =   5160
         TabIndex        =   33
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   13
         Left            =   5160
         TabIndex        =   32
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bolus"
         Height          =   255
         Left            =   6600
         TabIndex        =   31
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Compensador"
         Height          =   255
         Left            =   5160
         TabIndex        =   30
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bloque"
         Height          =   255
         Left            =   6600
         TabIndex        =   29
         Top             =   4680
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bandeja"
         Height          =   255
         Left            =   5160
         TabIndex        =   28
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   6
         Left            =   1680
         TabIndex        =   11
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   7
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   6
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   10
         Left            =   5160
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   11
         Left            =   5160
         TabIndex        =   2
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   9
         Left            =   5160
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Instrucciones para Cuadrar Campos"
         Height          =   495
         Left            =   480
         TabIndex        =   35
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciales"
         Height          =   255
         Left            =   4440
         TabIndex        =   27
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuña"
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Si o No"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Camilla"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Colimador"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Gantry"
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Lower(mm)"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Upper(mm)"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sad O SSD"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Campo"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dia"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   0
      Picture         =   "C.frx":0000
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Enum EACCION1
    AGREGAR_REGISTRO1 = 0
    EDITAR_REGISTRO1 = 1
End Enum

Public IdRegistro As Long
Public ACCION As EACCION

Private Sub Command1_Click()

On Error GoTo ErrorSub
        
    ' Valida el Nombre que no este vacio
    ''''''''''''''''''''''''''''''''
    If Trim(Text1(3)) = "" Then
        MsgBox "El Nombre de registro no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(3).SetFocus
        Exit Sub
    
    ' Valida el Apellido
    ''''''''''''''''''''''''''''''''
    ElseIf Trim(Text1(4)) = "" Then
        MsgBox "El Apellido no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(4).SetFocus
        Exit Sub
    End If

    'Agrega el registro
    '''''''''''''''''''''''''''''''
    
    Select Case ACCION
        Case EDITAR_REGISTRO1
        
            bd58(1) = Text1(1)
            bd58(2) = Text1(2)
            bd58(3) = Text1(3)
            bd58(4) = Text1(4)
            bd58(5) = Text1(5)
            bd58(6) = Text1(6)
            bd58(7) = Text1(7)
            bd58(8) = Text1(8)
            bd58(9) = Text1(9)
            bd58(10) = Text1(10)
            bd58(11) = Text1(11)
            bd58(12) = Text1(12)
            bd58(13) = Text1(13)
            bd58(14) = Text1(14)
            bd58(15) = Text1(15)
            bd58(16) = Text1(16)
            bd58(17) = Text1(17)
            bd58(18) = Text1(18)
            bd58(19) = Text1(19)
        
        Case AGREGAR_REGISTRO1
        
            bd58.AddNew
            bd58(1) = Text1(1)
            bd58(2) = Text1(2)
            bd58(3) = Text1(3)
            bd58(4) = Text1(4)
            bd58(5) = Text1(5)
            bd58(6) = Text1(6)
            bd58(7) = Text1(7)
            bd58(8) = Text1(8)
            bd58(9) = Text1(9)
            bd58(10) = Text1(10)
            bd58(11) = Text1(11)
            bd58(12) = Text1(12)
            bd58(13) = Text1(13)
            bd58(14) = Text1(14)
            bd58(15) = Text1(15)
            bd58(16) = Text1(16)
            bd58(17) = Text1(17)
            bd58(18) = Text1(18)
            
            bd58(19) = CDate(Label3)
            
    End Select
    
    bd58.Update
    
    Unload Me
    Set FrmEdicionTecnico = Nothing

Exit Sub
ErrorSub:
    MsgBox Err.Description

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    If Me.ACCION = AGREGAR_REGISTRO1 Then
       Me.Caption = "Agregar nuevo registro"
    ElseIf Me.ACCION = EDITAR_REGISTRO1 Then
       Me.Caption = "Editar registro"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub



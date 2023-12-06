VERSION 5.00
Begin VB.Form Form299 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form29"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form29"
   ScaleHeight     =   4980
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del Paciente"
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   7095
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   5160
         TabIndex        =   16
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Si o No"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   5640
         Picture         =   "EditarFac.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descuento"
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "IVA 12 %"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Precio"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tratamiento de Radioterapia"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dia"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   6360
      Picture         =   "EditarFac.frx":027F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   0
      Picture         =   "EditarFac.frx":0801
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form299"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bd1 As New ADODB.Recordset
Private Sub Command1_Click()

    'Agrega el registro
    '''''''''''''''''''''''''''''''
    
    Select Case ACCION
        Case EDITAR_REGISTRO
            
            'BD57(9) = CDate(Label3)
            
                'BD57(7) = IdPac1
                'BD57(5) = Text1(5)
            If Text1(0) = "" Then BD57.Fields("Cod_producto") = 0 Else BD57.Fields("Cod_producto") = Val(Text1(0))
            If Text1(1) = "" Then BD57.Fields("Precio") = 0 Else BD57.Fields("Precio") = Val(Text1(1))
            If Text1(2) = "" Then BD57.Fields("Cantidad") = 0 Else BD57.Fields("Cantidad") = Val(Text1(2))
            'BD57(3) = CBool(Check1.Value)
            If Text1(4) = "" Then BD57.Fields("Descuento") = 0 Else BD57.Fields("Descuento") = Val(Text1(4))
            If Text1(5) = "" Then BD57.Fields("descripcion") = "" Else BD57.Fields("descripcion") = Val(Text1(5))
           
                                  
        Case AGREGAR_REGISTRO
        
            BD57.AddNew
            
               ' BD57(5) = Text1(5)
                
            If Text1(0) = "" Then BD57(0) = Null Else BD57(0) = Val(Text1(0))
            If Text1(1) = "" Then BD57(1) = Null Else BD57(1) = Val(Text1(1))
            If Text1(2) = "" Then BD57(2) = Null Else BD57(2) = Val(Text1(2))
            
                BD57(3) = CBool(Check1.Value)
                
            If Text1(4) = "" Then BD57(4) = Null Else BD57(4) = Val(Text1(4))
            If Text1(5) = "" Then BD57(5) = Null Else BD57(5) = Val(Text1(5))
            
            
            BD57(7) = IdPac1
           
            BD57(9) = CDate(Label3)
            
    End Select
    
    BD57.Update
    
    Unload Me
    Set Form29 = Nothing

Exit Sub
ErrorSub:
    MsgBox Err.Description

End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
    If ACCION = AGREGAR_REGISTRO Then
       Me.Caption = "Agregar nuevo registro"
    ElseIf ACCION = EDITAR_REGISTRO Then
       Me.Caption = "Editar registro"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select

If Len(Trim(Text1(5).Text)) = 0 Then
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
Else
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End If

End Sub



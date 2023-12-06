VERSION 5.00
Begin VB.Form Form999 
   Caption         =   "Consulta Psicológica Adulto"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form9"
   ScaleHeight     =   10455
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Aspectos Psicológicos"
      Height          =   6855
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   12015
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   2160
         TabIndex        =   61
         Text            =   "Text23"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   1200
         TabIndex        =   56
         Text            =   "Text24"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   2160
         TabIndex        =   55
         Text            =   "Text22"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   2160
         TabIndex        =   54
         Text            =   "Text21"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   2160
         TabIndex        =   53
         Text            =   "Text20"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Observación"
         Height          =   2895
         Left            =   7920
         TabIndex        =   51
         Top             =   3360
         Width           =   3855
         Begin VB.TextBox Text19 
            Height          =   2415
            Left            =   240
            TabIndex        =   52
            Text            =   "Text19"
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.TextBox Text18 
         Height          =   735
         Left            =   360
         TabIndex        =   49
         Text            =   "Text18"
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuando te Menciona la Palabra Cáncer"
         Height          =   2895
         Left            =   3480
         TabIndex        =   37
         Top             =   3360
         Width           =   3975
         Begin VB.TextBox Text15 
            Height          =   2415
            Left            =   1320
            TabIndex        =   38
            Text            =   "Text15"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label24 
            Caption         =   "Eficiencia"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Efectividad"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Sensaciones"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Miedo"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Conducta"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Motivación"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "Actitudes"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Ideación"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Imaginaria"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox Text17 
         Height          =   735
         Left            =   360
         TabIndex        =   36
         Text            =   "Text17"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox Text16 
         Height          =   735
         Left            =   360
         TabIndex        =   34
         Text            =   "Text16"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   3240
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Text            =   "Text13"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   1095
         Left            =   8160
         TabIndex        =   30
         Text            =   "Text11"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         Height          =   975
         Left            =   8160
         TabIndex        =   29
         Text            =   "Text10"
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Height          =   1095
         Left            =   4440
         TabIndex        =   28
         Text            =   "Text9"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         Height          =   975
         Left            =   4440
         TabIndex        =   27
         Text            =   "Text8"
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label30 
         Caption         =   "Hijos"
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   480
         TabIndex        =   60
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Nombre del Conyugue"
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label28 
         Caption         =   "Nivel de Instrucción"
         Height          =   255
         Left            =   480
         TabIndex        =   58
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "Profesión u Oficio"
         Height          =   255
         Left            =   480
         TabIndex        =   57
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Mecanismo Defensivos"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Debilidades"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Fortalezas"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Religión"
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Sintomatología Psicológica "
         Height          =   255
         Left            =   8160
         TabIndex        =   26
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Enfermedad Actual "
         Height          =   255
         Left            =   8160
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Motivo de la Consulta"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Actitud ante el Diagnóstico"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del Paciente"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar datos"
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   8520
         TabIndex        =   19
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   1920
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   6720
         Picture         =   "Form9.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupacion"
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nacimiento"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Sexo"
         Height          =   255
         Left            =   5880
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellidos"
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombres"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro"
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form999"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD As New ADODB.Recordset 'tabla Registro Historico
Dim BD4 As New ADODB.Recordset 'tabla Informes medicos
Dim SQL As String
Dim bd1 As New ADODB.Recordset


Private Sub Command1_Click()

SQL = "select * from Paciente where Cedula like '%" & Text7.Text & "%'"

BD.Open SQL, CADENA, , , adCmdText

If BD.EOF Then
    
SQL = "select * from Paciente "
BD.Close
BD.Open SQL, CADENA, , , adCmdText
BD.Close
MsgBox "No Existe el registro"
Exit Sub

Else

End If

Call carga_de_datos

End Sub


Sub carga_de_datos()

            Text1.Text = BD.Fields("cedula")
            Text2.Text = BD.Fields("Fecha_reg")
            Text3.Text = BD.Fields("Nombre")
            Text4.Text = BD.Fields("Apellido")
            Text5.Text = BD.Fields("Fecha_nacimiento")
            Text6.Text = BD.Fields("Edad")
            Text14.Text = BD.Fields("Ocupacion")
            
            If BD.Fields("sexo") = 0 Then Text12.Text = "Masculino" Else Text12.Text = "Femenino"
            
           End Sub
           
Private Sub Command2_Click()

       
        If Text15.Text = "" Then
        f = "Antecedente_Flia"
        GoTo noguardA
        msg = "El Campo Antecedente Flamiliar esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If
        
        If Text19.Text = "" Then
        f = "Anatomia_Patol"
        GoTo noguardA
        msg = "El Campo Antecedente Patologico esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If
        
        If Text16.Text = "" Then
        f = "Enfermedad_Act"
        GoTo noguardA
        msg = "El Campo Enfermedad actual esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If
        
        If DTPicker1.Value = "" Then
        f = "Fecha"
        GoTo noguardA
        msg = "El Campo Fecha cheque"
        MsgBox Mgs, vbOKOnly
        End If
        
        If Text20.Text = "" Then
        f = "Examen_Fis"
        GoTo noguardA
        msg = "El Campo Examen Fisico esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If
        
        If Text18.Text = "" Then
        f = "Motivo_Con"
        GoTo noguardA
        msg = "El Campo Motivo dela consulta esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If
        
        If Text21.Text = "" Then
        f = "Diagnotico"
        GoTo noguardA
        msg = "El Campo Diagnotico esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If
        
        If Text18.Text = "" Then
        f = "Tratamiento"
        GoTo noguardA
        msg = "El Campo Tratamiento esta Vacio"
        MsgBox Mgs, vbOKOnly
        End If

                'hacer una rutina de comprobacion de los campos a guardar donde se verifique la integridad de los datos

        CSql = "Insert into INFORME_MEDICO(Antecedente_Flia, Anatomia_Patol, Enfermedad_Act, Examen_Fis, Motivo_Con, Diagnotico, Tratamiento,Fecha) VALUES('" & Text15.Text & "','" & Text19.Text & "','" & Text16.Text & "','" & Text20.Text & "','" & Text18.Text & "','" & Text21.Text & "','" & Text22.Text & "',#" & DTPicker1.Value & "#)"
        Dim BD As New ADODB.Recordset
        BD.Open CSql, CADENA
        msg = "Registro Agregado satisfactoriamente"
        MsgBox msg, vbOKOnly
         Call Blanqueo
Exit Sub

noguardA:
    msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox msg, vbOKOnly, "Error al Guardar"
    msg = "Debe de completar todo el formulario o hay un error en algun campo, Falta el campo: " & f
    MsgBox msg, vbOKOnly, "Error al Guardar"
    Exit Sub

End Sub

  Sub Blanqueo()
             Text1.Text = ""
             Text2.Text = ""
             Text3.Text = ""
             Text4.Text = ""
             Text5.Text = ""
             Text6.Text = ""
             Text7.Text = ""
             Text10.Text = ""
             Text11.Text = ""
             Text12.Text = ""
             Text13.Text = ""
             Text14.Text = ""
             Text15.Text = ""
             Text16.Text = ""
             Text15.Text = ""
             Text16.Text = ""
             Text18.Text = ""
             Text19.Text = ""
             Text20.Text = ""
             Text21.Text = ""
             Text22.Text = ""
        End Sub

Private Sub Command5_Click()
Unload Me
End Sub


Private Sub Command6_Click()
    Call carga_de_datos
End Sub

Private Sub Form_Load()
DTPicker1.Value = Now()
End Sub

Private Sub Text38_Change()

End Sub

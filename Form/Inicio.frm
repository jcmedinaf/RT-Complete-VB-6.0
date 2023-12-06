VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OncoAmerica"
   ClientHeight    =   7020
   ClientLeft      =   2430
   ClientTop       =   2025
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   8295
   Begin VB.Menu Tablas 
      Caption         =   "Tablas"
      Begin VB.Menu Pa 
         Caption         =   "Pacientes"
      End
      Begin VB.Menu TC 
         Caption         =   "Tipos Cancer"
      End
      Begin VB.Menu MR 
         Caption         =   "Médicos Remitentes"
      End
      Begin VB.Menu ME 
         Caption         =   "Médicos Tratantes"
      End
   End
   Begin VB.Menu PRO 
      Caption         =   "Procesos"
      Begin VB.Menu Pres 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu Hist 
         Caption         =   "Historias Médicas"
      End
      Begin VB.Menu CONs 
         Caption         =   "Consultas"
      End
   End
   Begin VB.Menu Salida 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
direc = "C:\oncoamerica"
cadenaconexioN = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" + direc + "\oncoamerica.mdb"

End Sub

Private Sub Hist_Click()
Form5.Show 1
End Sub

Private Sub ME_Click()
Form4.Show 1

End Sub

Private Sub MR_Click()
Form3.Show 1

End Sub

Private Sub Pa_Click()
Form2.Show 1
End Sub

Private Sub Salida_Click()
End
End Sub

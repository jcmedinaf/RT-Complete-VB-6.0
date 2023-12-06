VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmSoporteTecnico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sugerencia del día"
   ClientHeight    =   7860
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   6360
   Icon            =   "FrmSoporteTecnico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   6360
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   13996
      _Version        =   393217
      TextRTF         =   $"FrmSoporteTecnico.frx":1002
   End
End
Attribute VB_Name = "FrmSoporteTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
RichTextBox1.Text = "Servicio de soporte técnico de RT Complete Corporations " & Chr(13) & Chr(13) & _
"Información de soporte técnico en línea " & Chr(13) & _
"Si desea conocer todas las ofertas de soporte técnico, visite http://www.oncoamerica.net" & _
Chr(13) & Chr(13) & _
"Servicio de Atención Teléfonica " & Chr(13) & _
"Puede tener acceso a los servicios de teléfono en el número (00)(58) 7936963 en Venezuela para America Latina y habla Hispanoamerica." & _
Chr(13) & Chr(13) & _
"Servicio internacional " & Chr(13) & _
"El soporte técnico en E.E.U.U. y Canadá puede variar. Si desea obtener información de su país, visite http://www.oncoamerica.net. Si en su país no existe una oficina local, póngase en contacto con nosotros directamente." & _
Chr(13) & Chr(13) & _
"Condiciones " & Chr(13) & _
"Los servicios de soporte técnico de RT Complete están sujetos a las condiciones, términos y precios aplicables en ese momento, que a su vez están sujetos a cambios sin previo aviso."


End Sub

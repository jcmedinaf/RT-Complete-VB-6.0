VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmSoporteTecnico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sugerencia del d�a"
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
RichTextBox1.Text = "Servicio de soporte t�cnico de RT Complete Corporations " & Chr(13) & Chr(13) & _
"Informaci�n de soporte t�cnico en l�nea " & Chr(13) & _
"Si desea conocer todas las ofertas de soporte t�cnico, visite http://www.oncoamerica.net" & _
Chr(13) & Chr(13) & _
"Servicio de Atenci�n Tel�fonica " & Chr(13) & _
"Puede tener acceso a los servicios de tel�fono en el n�mero (00)(58) 7936963 en Venezuela para America Latina y habla Hispanoamerica." & _
Chr(13) & Chr(13) & _
"Servicio internacional " & Chr(13) & _
"El soporte t�cnico en E.E.U.U. y Canad� puede variar. Si desea obtener informaci�n de su pa�s, visite http://www.oncoamerica.net. Si en su pa�s no existe una oficina local, p�ngase en contacto con nosotros directamente." & _
Chr(13) & Chr(13) & _
"Condiciones " & Chr(13) & _
"Los servicios de soporte t�cnico de RT Complete est�n sujetos a las condiciones, t�rminos y precios aplicables en ese momento, que a su vez est�n sujetos a cambios sin previo aviso."


End Sub

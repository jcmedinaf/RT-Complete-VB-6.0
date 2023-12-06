VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormAyuda 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda del Sistema"
   ClientHeight    =   8610
   ClientLeft      =   3045
   ClientTop       =   3330
   ClientWidth     =   9555
   Icon            =   "FormAyuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin SHDocVwCtl.WebBrowser brwWebBrowser 
         Height          =   8160
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9120
         ExtentX         =   16087
         ExtentY         =   14393
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   -1  'True
         NoClientEdge    =   -1  'True
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAyuda.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAyuda.frx":042C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAyuda.frx":070E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAyuda.frx":09F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAyuda.frx":0CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAyuda.frx":0FB4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Show

Select Case Ayuda
    Case 0
        'ayuda de login
        brwWebBrowser.Navigate "file:\\" & App.Path & "\Help\" & "login.htm"
    Case 1
        'ayuda de Nuevo Paciente
        brwWebBrowser.Navigate "file:\\" & App.Path & "\Help\" & "NuevoPaciente.htm"
    Case 2

End Select
End Sub



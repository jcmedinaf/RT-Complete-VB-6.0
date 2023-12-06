VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmVistaPreviaHistorialMedico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   Icon            =   "FrmVistaPreviaHistorialMedico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmVistaPreviaHistorialMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport6

Private Sub Form_Load()
Dim RsReporte As New ADODB.Recordset
Select Case OpcionReporte
    Case Is = "HistorialMedico"
        CSql = "Select ApellidoP , NombreP, Fecha_NacimientoP, CedulaP, EdadP, Ocupacion, Codigo, Telefono, Codigoc, Celular, DireccionP, Motivo_Con, Diagnotico, Tratamiento, Examen_Fis, Enfermedad_Act From Informe_Med Where CedulaP='" & FrmHistorialMedico.TxtCedula.Text & "'"
    Case Is = "Oncologia"
        CSql = "Select ApellidoP, NombreP, Fecha_NacimientoP, CedulaP, EdadP, Ocupacion, Codigo, Telefono, Codigoc, Celular, DireccionP, Motivo_Con, Diagnotico, Tratamiento, Examen_Fis, Enfermedad_Act From Informe_Med Where CedulaP='" & FrmRadioTerapeuta.Text1.Text & "'"
End Select
    
Set RsReporte = CrearRS(CSql)

If RsReporte.RecordCount > 0 Then
    Screen.MousePointer = vbHourglass
    Report.DiscardSavedData
    Report.Database.SetDataSource RsReporte, 3, 1
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
End If

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

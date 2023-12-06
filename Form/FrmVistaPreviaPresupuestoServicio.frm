VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmVistaPreviaPresupuestoServicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "FrmVistaPreviaPresupuestoServicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmVistaPreviaPresupuestoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport4

Private Sub Form_Load()
Dim RsReporte As New ADODB.Recordset
CSql = "Select Fecha , NPresupuesto, Razon, Rif, DireccionC, Email, ApellidoP, NombreP, CedulaP, Codigo, Telefono, Codigoc, Celular, DireccionP, Duracion, Diagnotico, Monto From Presupuesto2 Where Rif='" & FrmPresupuestoProducto.Text13.Text & "' And CedulaP='" & FrmPresupuestoProducto.Text7.Text & "'"
Set RsReporte = CrearRS(CSql)
'
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

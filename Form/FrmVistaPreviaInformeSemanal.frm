VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmVistaPreviaInformeSemanal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   Icon            =   "FrmVistaPreviaInformeSemanal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11130
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
      EnableGroupTree =   -1  'True
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
Attribute VB_Name = "FrmVistaPreviaInformeSemanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport8

Private Sub Form_Load()


FechaDesde = Format(FrmReporteNutricion.DtpFechaDesde.Value, "dd/mm/yyyy")
FechaHasta = Format(FrmReporteNutricion.DtPFechaHasta.Value, "dd/mm/yyyy")

'========= ESTE ES EL CODIGO NUEVO ==========

Q1 = "CedulaP, NombreP, ApellidoP, EdadP, FechaNu, Diagnotico, "
Q2 = "Estado, Menarquia, años, Aborto, Hormonas, Familiares, Intervenciones, Actividad, Tratamiento, Quimi, Desayuno, "
Q3 = "Cena1, Almuerzo, Cena2, GCCA, GCC1, GCCD, GCC2, PesoU, PesoA, Talla, CambioP, Indice, Globulos, Hematocrito, Hemoglobina, "
Q4 = "HCM, VCM, Plaquetas, Cuentas, Segmentados, Linfocitos, Eosinofilos, Glicemia, Urea, Creatinina, Acido_U, Colesterol, "
Q5 = "Trigliceridos, Fosforo, Calcio, Potasio, Cloro, TGO, TGP, HDL, Amilasa, BilirrubinaT, BilirrubinaD, BilirruBinaI, "
Q6 = "Monocitos, LDL, VLDL, Sodio, Magnesio, energia, Grasa, Vitamina, Mineral, Cho, Orina, Heces, DNI, Recomendaciones, Otros"

CSql = "Select " & Q1 & Q2 & Q3 & Q4 & Q5 & Q6 & " From Info_Nutri Where FechaNu>='" & FechaDesde & "' And FechaNu<='" & FechaHasta & "'"
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

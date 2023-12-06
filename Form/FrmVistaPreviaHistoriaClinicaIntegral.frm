VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmVistaPreviaHistoriaClinicaIntegral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "FrmVistaPreviaHistoriaClinicaIntegral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6525
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
Attribute VB_Name = "FrmVistaPreviaHistoriaClinicaIntegral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport5

Private Sub Form_Load()

Dim RsReporte As New ADODB.Recordset
Dim Q1, Q2, Q3, Q4, Q5, Q6, Q7 As String
Q1 = "CedulaP, NombreP, ApellidoP, Codigo, Telefono, CodigoC, Celular, Fecha_NacimientoP, SexoP, DireccionP, EdadP, Ocupacion, Fecha_Inicio, Fecha_culm, "
Q2 = "EmailP, Enfernedad_Act, ColorP, Signos, Cabello, Torax, Abdomen, Neurologico, Revision, Menarquia, FLujo, Ciclo, Menopausia, Sueño, Hora, Cuales, Actividad, "
Q3 = "Vivos, Muertos, Cual, Hormonas, Recidivas, Cigarrillo, Cafe, Alcohol, Gesta, Aborto, Familiares, Intervenciones, Quimi, CicloQ, Tratamiento, "
Q4 = "Alergia, FechaExamen, Globulos, Hematocrito, Hemoglobina, HCM, VCM, Plaquetas, Cuentas, Segmentados, Linfocitos, Eosinofilos, Glicemia, Urea, Creatinina, "
Q5 = "Acido_U, Colesterol, Trigliceridos, Fosforo, Calcio, Potasio, Cloro, TGO, TGP, HDL, Amilasa, BilirrubinaT, BilirrubinaD, BilirrubinaI, Monocitos, LDL, VLDL, "
Q6 = "Sodio, Magnesio, Energia, Grasa, Vitamina, Mineral, Cho, Orina, Heces, Otros, Desayuno, Almuerzo, Cena1, GCCD, GCCA, GCC1, PesoU, PesoA, PesoR, Talla, "
Q7 = "Indice, CambioP, Fosfatasa, Observaciones, Observaciones, FechaAntecedentes, FechaQ, FechaR, Observacion, Cena2, GCC2"

SQL = Q1 & Q2 & Q3 & Q4 & Q5 & Q6 & Q7

Select Case OpcionReporte
    Case Is = "HistorialNutricional"
        CSql = "Select " & SQL & " From Historia_Clinica Where CedulaP='" & FrmHistorialNutricional.Text1.Text & "'"
    Case Is = "Oncologia"
        CSql = "Select " & SQL & " From Historia_Clinica Where CedulaP='" & FrmRadioTerapeuta.Text1.Text & "'"
    Case Is = "DireccionMedica"
        CSql = "Select " & SQL & " From Historia_Clinica Where CedulaP='" & FrmDireccionMedica.Text1.Text & "'"
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

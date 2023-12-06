VERSION 5.00
Begin VB.Form FrmIVSSPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion Hospitalaria Venezolana (AHV)"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12930
   Icon            =   "FrmIVSSMenuPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmIVSSMenuPrincipal.frx":0442
   ScaleHeight     =   9150
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuHospitalesAmbulatorios 
      Caption         =   "Hospitales y Ambulatorios"
   End
   Begin VB.Menu MnuCirugia 
      Caption         =   "Cirugía"
      Begin VB.Menu SubMnuMesasQuirurgicas 
         Caption         =   "Mesas Quirúrgicas"
      End
      Begin VB.Menu SubMnuDesfibriladores 
         Caption         =   "Desfibriladores"
      End
      Begin VB.Menu SubMnuBombasQuirurgicas 
         Caption         =   "Bombas Quirúrgicas"
      End
      Begin VB.Menu SubMnuBombasPerfusion 
         Caption         =   "Bombas de Perfusión"
      End
      Begin VB.Menu SubMnuElectrobisturis 
         Caption         =   "Electrobisturis"
      End
      Begin VB.Menu SubMnuEquiposAnestesia 
         Caption         =   "Equipos de Anestesia"
      End
      Begin VB.Menu SubMnuMonitoresSignosVitales 
         Caption         =   "Monitores de Signos Vitales"
      End
      Begin VB.Menu SubMnuLámparasQuirúrgicasTecho 
         Caption         =   "Lámparas Quirúrgicas de Techo"
      End
      Begin VB.Menu SubMnuLámparasQuirúrgicasPedestal 
         Caption         =   "Lámparas Quirúrgicas-Pedestal"
      End
   End
   Begin VB.Menu MnuImagenes 
      Caption         =   "Imágenes"
      Begin VB.Menu SubMnuImagenesTodos 
         Caption         =   "Todos"
      End
      Begin VB.Menu SubMnuImagenesTomografos 
         Caption         =   "Tomógrafos"
      End
      Begin VB.Menu SubMnuResonanciasMagneticas 
         Caption         =   "Resonancias Magnéticas"
      End
      Begin VB.Menu SubMnuUltraSonido 
         Caption         =   "Ultrasonido"
      End
      Begin VB.Menu SubMnuAngiografia 
         Caption         =   "Angiografía"
      End
      Begin VB.Menu SubMnuHemodinamica 
         Caption         =   "Hemodinámica"
      End
      Begin VB.Menu SubMnuMamografia 
         Caption         =   "Mamografía"
      End
      Begin VB.Menu SubMnuRayosXFijos 
         Caption         =   "Rayos X-fijos"
      End
      Begin VB.Menu SubMnuRayosXMovil 
         Caption         =   "Rayos X-Movil"
      End
      Begin VB.Menu SubMnuArcoC 
         Caption         =   "Arco en C"
      End
      Begin VB.Menu SubMnuCamaraNuclear 
         Caption         =   "Cámara Nuclear"
      End
   End
   Begin VB.Menu MnuPediatria 
      Caption         =   "Pediatría"
      Begin VB.Menu SubMnuVentiladores 
         Caption         =   "Ventiladores"
      End
      Begin VB.Menu SubMnuEncubadoras 
         Caption         =   "Encubadoras"
      End
      Begin VB.Menu SubMnuEncubadorasTransportables 
         Caption         =   "Encubadoras Transportables"
      End
      Begin VB.Menu SubMnuMonitoresSignosVitalesPe 
         Caption         =   "Monitores de Signos Vitales"
      End
   End
   Begin VB.Menu MnuOncologia 
      Caption         =   "Oncología"
      Begin VB.Menu SubMnuBombasInflusion 
         Caption         =   "Bombas de Influsión"
      End
      Begin VB.Menu SubMnuSillasPacientes 
         Caption         =   "Sillas de Pacientes"
      End
      Begin VB.Menu SubMnuCampanasVentilacion 
         Caption         =   "Campanas de Ventilación"
      End
   End
   Begin VB.Menu MnuRadioterapia 
      Caption         =   "Radioterpia"
      Begin VB.Menu SubMnuTodosRadioterapia 
         Caption         =   "Todos"
      End
      Begin VB.Menu SubMnuAceleradoresLiniales 
         Caption         =   "Aceleradores Lineales"
      End
      Begin VB.Menu SubMnuCyberKnife 
         Caption         =   "CyberKnife"
      End
      Begin VB.Menu SubMnuCobalto60 
         Caption         =   "Cobalto 60"
      End
      Begin VB.Menu SubMnuSimuladorConvencional 
         Caption         =   "Simulador Convencional"
      End
      Begin VB.Menu SubMnuCtSimuladores 
         Caption         =   "CT-Simuladores"
      End
      Begin VB.Menu SubMnuBraquiterapiaHDR 
         Caption         =   "Braquiterapia HDR"
      End
      Begin VB.Menu SubMnuFisicaMedica 
         Caption         =   "Física Médica"
      End
      Begin VB.Menu SubMnuEquipoDosimetria 
         Caption         =   "Equipo de Dosimetría"
      End
      Begin VB.Menu SubMnuSoftwarePlanificacion 
         Caption         =   "Software de Planificación"
      End
      Begin VB.Menu SubMnuSoftwareHardware 
         Caption         =   "Software y Hardware"
      End
   End
   Begin VB.Menu MnuMuebles 
      Caption         =   "Muebles"
      Begin VB.Menu SubMnuCamas 
         Caption         =   "Camas"
      End
      Begin VB.Menu SubMnuCamillas 
         Caption         =   "Camillas"
      End
   End
   Begin VB.Menu MnuBusqueda 
      Caption         =   "Busqueda"
      Begin VB.Menu SubMnuEquipos 
         Caption         =   "Equipos"
      End
      Begin VB.Menu SubMnuHospitales 
         Caption         =   "Hospitales"
      End
      Begin VB.Menu SubMnuAmbulatorios 
         Caption         =   "Ambulatorios"
      End
      Begin VB.Menu SubMnuAdministración 
         Caption         =   "Administración"
         Begin VB.Menu SubMnuBEquipo 
            Caption         =   "Equipos"
         End
         Begin VB.Menu SubMnuBHospitales 
            Caption         =   "Hospitales"
         End
         Begin VB.Menu SubMnuBAmbulatorios 
            Caption         =   "Ambulatorios"
         End
         Begin VB.Menu SubMnuBCategorias 
            Caption         =   "Categorias"
         End
      End
   End
   Begin VB.Menu SubMnuCerrar 
      Caption         =   "Cerrar"
   End
End
Attribute VB_Name = "FrmIVSSPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MnuHospitalesAmbulatorios_Click()
FrmIVSSHospitalesAmbula.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuAceleradoresLiniales_Click()
Pn = 26
FrmIVSSEquipos.Caption = "Aceleradores Liniales"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuAmbulatorios_Click()
FrmIVSSBusquedaAmbulatorios.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuAngiografia_Click()
Pn = 13
FrmIVSSEquipos.Caption = "Angriografía"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuArcoC_Click()
Pn = 17
FrmIVSSEquipos.Caption = "Arco en C"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBAmbulatorios_Click()
FrmIVSSAdmAmbulatorios.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBCategorias_Click()
FrmIVSSAdmCategorias.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBEquipo_Click()
FrmIVSSAdmEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBombasInflusion_Click()
Pn = 23
FrmIVSSEquipos.Caption = "Bombas de Influsión"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBombasPerfusion_Click()
FrmIVSSBombasPerfusion.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBombasQuirurgicas_Click()
FrmIVSSBombasQuirurjicas.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBraquiterapiaHDR_Click()
Pn = 31
FrmIVSSEquipos.Caption = "Braquiterapia HDR"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCamaraNuclear_Click()
Pn = 18
FrmIVSSEquipos.Caption = "Camara Nuclear"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCamas_Click()
Pn = 36
FrmIVSSEquipos.Caption = "Camas"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCamillas_Click()
Pn = 37
FrmIVSSEquipos.Caption = "Camillas"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub


Private Sub SubMnuCampanasVentilacion_Click()
Pn = 25
FrmIVSSEquipos.Caption = "Campanas de Ventilación"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCerrar_Click()
Unload Me
End Sub

Private Sub SubMnuCobalto60_Click()
Pn = 28
FrmIVSSEquipos.Caption = "Cobalto 60"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCtSimuladores_Click()
Pn = 30
FrmIVSSEquipos.Caption = "Ct Simuladores"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCyberKnife_Click()
Pn = 27
FrmIVSSEquipos.Caption = "CyberKnife"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuDesfibriladores_Click()
FrmIVSSDesfibriladores.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuElectrobisturis_Click()
FrmIVSSunidadesElectrocirugia.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuEncubadoras_Click()
Pn = 20
FrmIVSSEquipos.Caption = "Encubadoras"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuEncubadorasTransportables_Click()
Pn = 21
FrmIVSSEquipos.Caption = "Encubadoras Transportables"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub


Private Sub SubMnuEquipoDosimetria_Click()
Pn = 33
FrmIVSSEquipos.Caption = "Equipo de Sosimetría"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuEquipos_Click()
FrmIVSSBusquedaEquipo.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuEquiposAnestesia_Click()
FrmIVSSEquiposAnestesia.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuFisicaMedica_Click()
Pn = 32
FrmIVSSEquipos.Caption = "Física Médica"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuHemodinamica_Click()
Pn = 14
FrmIVSSEquipos.Caption = "Hemodinámica"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuHospitales_Click()
FrmIVSSBusquedaHospitales.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuImagenesTodos_Click()
Pn = 1
FrmIVSSEquiposTodos.Caption = "Todos los Equipos de Imágenes"
FrmIVSSEquiposTodos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuImagenesTomografos_Click()
Pn = 10
FrmIVSSImagenesTomografia.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuLámparasQuirúrgicasPedestal_Click()
FrmLamparasQuirurgicasPedestal.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuLámparasQuirúrgicasTecho_Click()
FrmLamparasQuirurgicasTecho.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuMamografia_Click()
Pn = 38
FrmIVSSEquipos.Caption = "Mamografía"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuMesasQuirurgicas_Click()
FrmIVSSMesasQuirurjicas.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuMonitoresSignosVitales_Click()
FrmIVSSMonitoresSignosVitales.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuMonitoresSignosVitalesPe_Click()
Pn = 22
FrmIVSSEquipos.Caption = "Monitores Signos Vitales"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuRayosXFijos_Click()
Pn = 15
FrmIVSSEquipos.Caption = "Rayos X Fijos"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuRayosXMovil_Click()
Pn = 16
FrmIVSSEquipos.Caption = "Rayos X Móvil"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub


Private Sub SubMnuResonanciasMagneticas_Click()
Pn = 11
FrmIVSSEquipos.Caption = "Resonancia Magneticas"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuSillasPacientes_Click()
Pn = 24
FrmIVSSEquipos.Caption = "Sillas de Pacientes"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuSimuladorConvencional_Click()
Pn = 29
FrmIVSSEquipos.Caption = "Simulador Convencional"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuSoftwareHardware_Click()
Pn = 35
FrmIVSSEquipos.Caption = "Software y Hardware"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuSoftwarePlanificacion_Click()
Pn = 34
FrmIVSSEquipos.Caption = "Software de Planificación"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuTodosRadioterapia_Click()
Pn = 4
FrmIVSSEquiposTodos.Caption = "Todos los Equipos de Radioterapia"
FrmIVSSEquiposTodos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuUltraSonido_Click()
Pn = 12
FrmIVSSEquipos.Caption = "UltraSonido"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuVentiladores_Click()
Pn = 19
FrmIVSSEquipos.Caption = "Ventiladores"
FrmIVSSEquipos.Show vbModal, FrmPrincipal
End Sub



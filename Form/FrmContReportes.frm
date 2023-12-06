VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReportes 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista de Reportes"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15675
   Icon            =   "FrmContReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   15675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   15495
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4680
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   14400
         TabIndex        =   3
         ToolTipText     =   "Cerrar "
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cerrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmContReportes.frx":1002
         PICN            =   "FrmContReportes.frx":101E
         PICH            =   "FrmContReportes.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnResultados 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Resultados"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmContReportes.frx":141C
         PICN            =   "FrmContReportes.frx":1438
         PICH            =   "FrmContReportes.frx":16CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnImprimir 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Imprimir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmContReportes.frx":1857
         PICN            =   "FrmContReportes.frx":1873
         PICH            =   "FrmContReportes.frx":1998
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Reporte"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15495
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   9763
         Object.Width           =   15225
         Object.Height          =   5505
         ScrollBar       =   3
         AllowAddNew     =   -1  'True
         DrawColorGrid   =   1
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmContReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim TamDMGrid As Integer
Dim ValorAnt As String
Dim i As Integer
Dim j As Integer

Private Sub BtnCerrar_Click()
Unload Me
End Sub
 

Private Sub BtnImprimir_Click()
Dim StringTemp As String
Dim StringTemp1 As String
Dim StringTemp2 As String
Dim StringTemp3 As String
Dim MostrarReporte As Byte

TamDMGrid = DMGrid1.Rows

CSql = "DELETE FROM ContImpresion WHERE IdUser=" & IdUser
Set RsTemp = CrearRS(CSql)

ValorAnt = ""

If UCase(Tipo) = UCase("ReporteDiarioGeneral") Then
    ' Cuando es llamado desde el Módulo del Reporte del Diario General
    For i = 1 To TamDMGrid
        If ValorAnt <> Trim(DMGrid1.ValorCelda(i, 2)) And Trim(DMGrid1.ValorCelda(i, 2)) <> "" Then
            ValorAnt = Trim(DMGrid1.ValorCelda(i, 2))
        End If
        
        If IsNull(DMGrid1.ValorCelda(i, 3)) Then
            StringTemp = ""
        Else
            StringTemp = Trim(DMGrid1.ValorCelda(i, 3))
        End If
        
        If UCase(Trim(StringTemp)) <> UCase("Total comprobante:") Then
            If StringTemp <> "" Then
                If (IsNull(Trim(DMGrid1.ValorCelda(i, 1))) And IsNull(Trim(DMGrid1.ValorCelda(i, 2)))) Or _
                Trim(DMGrid1.ValorCelda(i, 1)) = "" And Trim(DMGrid1.ValorCelda(i, 2)) = "" Then
                    CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, DG, DG1, DG2, DG3, DG4, DG5, DG6, DG7) " & _
                        "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & ValorAnt & _
                        "','" & StringTemp & "','" & DMGrid1.ValorCelda(i, 4) & "','" & _
                        DMGrid1.ValorCelda(i, 5) & "','" & DMGrid1.ValorCelda(i, 6) & "','" & DMGrid1.ValorCelda(i, 7) & _
                        "','" & DMGrid1.ValorCelda(i, 8) & "')"
                    Set RsTemp = CrearRS(CSql)
                Else
                    CSql = "INSERT INTO ContImpresion (DGE,IdUser, OrdenIngreso, DG, DG1, DG2, DG3, DG4, DG5) " & _
                        "VALUES ('" & StringTemp & "'," & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & ValorAnt & _
                        "','','" & DMGrid1.ValorCelda(i, 4) & "','" & _
                        DMGrid1.ValorCelda(i, 5) & "','" & DMGrid1.ValorCelda(i, 6) & "')"
                    Set RsTemp = CrearRS(CSql)
                End If
            End If
        End If
    Next i
    MostrarReporte = 1

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteDiarioLegal") Then
    ' Cuando es llamado desde el Módulo del Reporte del Diario Legal
    If IsNull(DMGrid1.ValorCelda(i, 2)) Then
        StringTemp = ""
    Else
        StringTemp = Trim(DMGrid1.ValorCelda(i, 2))
    End If
    
    For i = 1 To TamDMGrid
        If Trim(DMGrid1.ValorCelda(i, 2)) = "" Then
            CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso) " & _
                "VALUES (" & IdUser & "," & i & ")"
            Set RsTemp = CrearRS(CSql)
        Else
            CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, DL, DL1, DL2, DL3) " & _
                "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                "','" & DMGrid1.ValorCelda(i, 3) & "','" & DMGrid1.ValorCelda(i, 4) & "')"
            Set RsTemp = CrearRS(CSql)
        End If
    Next i
    MostrarReporte = 2
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteLibroMayor") Then
    ' Cuando es llamado desde el Módulo del Reporte del Libro Mayor
    For i = 1 To TamDMGrid
        CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, LM, LM1, LM2, LM3) " & _
            "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
            "','" & DMGrid1.ValorCelda(i, 3) & "','" & DMGrid1.ValorCelda(i, 4) & "')"
        Set RsTemp = CrearRS(CSql)
    Next i
    MostrarReporte = 3
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteMayorAnalitico") Then
    ' Cuando es llamado desde el Módulo del Reporte del Mayor Analitico
    For i = 1 To TamDMGrid
    
        If IsNull(DMGrid1.ValorCelda(i, 3)) Then
            StringTemp = ""
        Else
            StringTemp = Trim(DMGrid1.ValorCelda(i, 3))
        End If
        
        If UCase(Trim(StringTemp)) <> UCase("Total Comprobante:") Then
            If Trim(StringTemp) <> "" Then
                If IsNull(Trim(DMGrid1.ValorCelda(i, 2))) Or Trim(DMGrid1.ValorCelda(i, 2)) = "" Then
                    CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, MAE1, MAE2, MAE3) " & _
                        "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 3) & _
                        "','" & DMGrid1.ValorCelda(i, 7) & "')"
                    Set RsTemp = CrearRS(CSql)
                Else
                    CSql = "INSERT INTO ContImpresion (MAE3,IdUser, OrdenIngreso, MA, MA1, MA2, MA3, MA4, MA5, MA6) " & _
                        "VALUES ('" & DMGrid1.ValorCelda(i, 7) & "'," & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                        "','" & DMGrid1.ValorCelda(i, 3) & "','" & DMGrid1.ValorCelda(i, 4) & _
                        "','" & DMGrid1.ValorCelda(i, 5) & "','" & DMGrid1.ValorCelda(i, 6) & "','" & DMGrid1.ValorCelda(i, 8) & "')"
                    Set RsTemp = CrearRS(CSql)
                End If
            End If
        End If
    Next i
    MostrarReporte = 4
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteComprobantes") Then
    ' Cuando es llamado desde el Módulo del Reporte de los Comprobantes
    For i = 1 To TamDMGrid
    
        If IsNull(DMGrid1.ValorCelda(i, 3)) Then
            StringTemp = ""
        Else
            StringTemp = Trim(DMGrid1.ValorCelda(i, 3))
        End If
        
        If UCase(Trim(StringTemp)) <> UCase("Total comprobante:") Then
            If StringTemp <> "" Then
                CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, CBTS, CBTS1, CBTS2, CBTS6, CBTS7, CBTSE3) " & _
                    "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                    "','" & DMGrid1.ValorCelda(i, 3) & "','" & DMGrid1.ValorCelda(i, 4) & _
                    "','" & DMGrid1.ValorCelda(i, 5) & "'," & DMGrid1.ValorCelda(i, 9) & ")"
                Set RsTemp = CrearRS(CSql)
            End If
        End If
    Next i
    MostrarReporte = 5
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteComprobantes2") Then
   ' Cuando es llamado desde el Módulo del Reporte de los Comprobantes
   Dim Cad1 As String
   Dim Cad2 As String
   Dim Cad3 As String
   Dim Cad4 As String
   Dim Cad5 As String
   Dim Cad6 As String
   Dim Cad7 As String
   Dim Cad8 As String
   
    For i = 1 To TamDMGrid
    
        If IsNull(DMGrid1.ValorCelda(i, 1)) Then Cad1 = "" Else Cad1 = Trim(DMGrid1.ValorCelda(i, 1))
        If IsNull(DMGrid1.ValorCelda(i, 2)) Then Cad2 = "" Else Cad2 = Trim(DMGrid1.ValorCelda(i, 2))
        If IsNull(DMGrid1.ValorCelda(i, 3)) Then Cad3 = "" Else Cad3 = Trim(DMGrid1.ValorCelda(i, 3))
        If IsNull(DMGrid1.ValorCelda(i, 4)) Then Cad4 = "" Else Cad4 = Trim(DMGrid1.ValorCelda(i, 4))
        If IsNull(DMGrid1.ValorCelda(i, 5)) Then Cad5 = "" Else Cad5 = Trim(DMGrid1.ValorCelda(i, 5))
        If IsNull(DMGrid1.ValorCelda(i, 6)) Then Cad6 = "" Else Cad6 = Trim(DMGrid1.ValorCelda(i, 6))
        If IsNull(DMGrid1.ValorCelda(i, 7)) Then Cad7 = "" Else Cad7 = Trim(DMGrid1.ValorCelda(i, 7))
        If IsNull(DMGrid1.ValorCelda(i, 8)) Then Cad8 = "" Else Cad8 = Trim(DMGrid1.ValorCelda(i, 8))
        
        If Cad7 = "" Then Cad7 = "0"
        If Cad8 = "" Then Cad8 = "0"
        
        If UCase(Trim(Cad3)) <> UCase("Total comprobante:") Then
            If Cad3 <> "" Then
                If (IsNull(Trim(DMGrid1.ValorCelda(i, 7))) And IsNull(Trim(DMGrid1.ValorCelda(i, 8)))) Or _
                    Trim(DMGrid1.ValorCelda(i, 7)) = "" And Trim(DMGrid1.ValorCelda(i, 8)) = "" Then
                    CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, CBTSE, CBTSE1, CBTSE2,CBTSE3) " & _
                        "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                        "','" & DMGrid1.ValorCelda(i, 3) & "'," & DMGrid1.ValorCelda(i, 9) & ")"
                    Set RsTemp = CrearRS(CSql)
                Else
                    CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, CBTS, CBTS1, CBTS2, CBTS3, CBTS4, CBTS5, CBTS6, CBTS7,CBTSE3) " & _
                        "VALUES (" & IdUser & "," & i & ",'" & Cad1 & "','" & Cad2 & "','" & Cad3 & "','" & Cad4 & _
                        "','" & Cad5 & "','" & Cad6 & "'," & Replace(Replace(Cad7, ".", ""), ",", ".") & "," & _
                        Replace(Replace(Cad8, ".", ""), ",", ".") & "," & DMGrid1.ValorCelda(i, 9) & ")"
                    Set RsTemp = CrearRS(CSql)
                End If
            End If
        End If
    Next i
    MostrarReporte = 6
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteComprobantesMayorizados") Then
    ' Cuando es llamado desde el Módulo del Reporte de los Comprobantes
    For i = 1 To TamDMGrid
    
            If IsNull(DMGrid1.ValorCelda(i, 2)) Then
            StringTemp = ""
        Else
            StringTemp = Trim(DMGrid1.ValorCelda(i, 2))
        End If
        
        If UCase(Trim(StringTemp)) <> UCase("Total:") Then
            If StringTemp <> "" Then
                CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, CM, CM1, CM2, CM3) " & _
                    "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                    "'," & DMGrid1.ValorCelda(i, 3) & "," & DMGrid1.ValorCelda(i, 4) & ")"
                Set RsTemp = CrearRS(CSql)
            End If
        End If
    Next i
    MostrarReporte = 7
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteComprobantesMayorizados2") Then
    ' Cuando es llamado desde el Módulo del Reporte de los Comprobantes
    For i = 1 To TamDMGrid
    
        If IsNull(DMGrid1.ValorCelda(i, 2)) Then
            StringTemp = ""
        Else
            StringTemp = Trim(DMGrid1.ValorCelda(i, 2))
        End If
        
        If IsNull(DMGrid1.ValorCelda(i, 3)) Then StringTemp2 = "0" Else StringTemp2 = DMGrid1.ValorCelda(i, 3)
        If IsNull(DMGrid1.ValorCelda(i, 4)) Then StringTemp3 = "0" Else StringTemp3 = DMGrid1.ValorCelda(i, 4)
        
        If StringTemp2 = "" Then StringTemp2 = "0"
        If StringTemp3 = "" Then StringTemp3 = "0"
        
        If UCase(Trim(StringTemp)) <> UCase("Total:") Then
            If StringTemp <> "" Then
                If InStr(1, DMGrid1.ValorCelda(i, 2), "Reng:") <> 0 Then
                    CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, CME1, CME2, CME3) " & _
                        "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                        "','" & DMGrid1.ValorCelda(i, 5) & "')"
                    Set RsTemp = CrearRS(CSql)
                Else
                    CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, CM, CM1, CM2, CM3,CME3) " & _
                        "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "','" & DMGrid1.ValorCelda(i, 2) & _
                        "'," & StringTemp2 & "," & StringTemp3 & ",'" & DMGrid1.ValorCelda(i, 5) & "')"
                    Set RsTemp = CrearRS(CSql)
                End If
            End If
        End If
    Next i
    MostrarReporte = 8
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
ElseIf UCase(Tipo) = UCase("ReporteBalanceGeneral") Or UCase(Tipo) = UCase("ReporteGananciasPerdidas") Then
    ' Cuando es llamado desde el Módulo del Reporte de los Comprobantes
    For i = 1 To TamDMGrid
    
        If IsNull(DMGrid1.ValorCelda(i, 1)) Then
            StringTemp = ""
        Else
            StringTemp = Trim(DMGrid1.ValorCelda(i, 1))
        End If
        
        If IsNull(DMGrid1.ValorCelda(i, 2)) Then StringTemp1 = "0" Else StringTemp1 = DMGrid1.ValorCelda(i, 2)
        If IsNull(DMGrid1.ValorCelda(i, 3)) Then StringTemp2 = "0" Else StringTemp2 = DMGrid1.ValorCelda(i, 3)
        If IsNull(DMGrid1.ValorCelda(i, 4)) Then StringTemp3 = "0" Else StringTemp3 = DMGrid1.ValorCelda(i, 4)
        
        If StringTemp1 = "" Then StringTemp1 = "0"
        If StringTemp2 = "" Then StringTemp2 = "0"
        If StringTemp3 = "" Then StringTemp3 = "0"
        
        StringTemp1 = Replace(Replace(StringTemp1, ".", ""), ",", ".")
        StringTemp2 = Replace(Replace(StringTemp2, ".", ""), ",", ".")
        StringTemp3 = Replace(Replace(StringTemp3, ".", ""), ",", ".")
        
        If StringTemp <> "" Then
            CSql = "INSERT INTO ContImpresion (IdUser, OrdenIngreso, BG, BG1, BG2, BG3) " & _
                "VALUES (" & IdUser & "," & i & ",'" & DMGrid1.ValorCelda(i, 1) & "'," & StringTemp1 & _
                "," & StringTemp2 & "," & StringTemp3 & ")"
            Set RsTemp = CrearRS(CSql)
        End If
        
    Next i
    
    If UCase(Tipo) = UCase("ReporteBalanceGeneral") Then
        MostrarReporte = 9
    Else
        MostrarReporte = 10
    End If
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


With CrystalReport1

    If MostrarReporte = 1 Then .ReportFileName = RutaInformes & "\ContDiarioGeneral.rpt"
    If MostrarReporte = 2 Then .ReportFileName = RutaInformes & "\ContDiarioLegal.rpt"
    If MostrarReporte = 3 Then .ReportFileName = RutaInformes & "\ContLibroMayor.rpt"
    If MostrarReporte = 4 Then .ReportFileName = RutaInformes & "\ContMayorAnalitico.rpt"
    If MostrarReporte = 5 Then .ReportFileName = RutaInformes & "\ContComprobantes.rpt"
    If MostrarReporte = 6 Then .ReportFileName = RutaInformes & "\ContComprobantesDetallados.rpt"
    If MostrarReporte = 7 Then .ReportFileName = RutaInformes & "\ContComprobantesMayorizados.rpt"
    If MostrarReporte = 8 Then .ReportFileName = RutaInformes & "\ContComprobantesMayorizadosDetallados.rpt"
    If MostrarReporte = 9 Then .ReportFileName = RutaInformes & "\ContBalanceGeneral.rpt"
    If MostrarReporte = 10 Then .ReportFileName = RutaInformes & "\ContGananciasPerdidas.rpt"
    
    If .ReportFileName = "" Then Exit Sub
    .Connect = "DSQ=OAClinica;"
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SelectionFormula = "{ContImpresion.IdUser} = " & IdUser
    .ReportTitle = "Reporte Diario General"
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .WindowMaxButton = False
    .WindowMinButton = False
    .Action = 1
    
End With


End Sub

Private Sub BtnResultados_Click()
'Tipo = "ReporteDiarioGeneral"
End Sub

Private Sub Form_Load()
Centrar Me
End Sub

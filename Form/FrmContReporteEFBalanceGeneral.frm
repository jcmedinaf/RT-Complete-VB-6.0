VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteEFBalanceGeneral 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance General"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "FrmContReporteEFBalanceGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   6735
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   5520
         TabIndex        =   7
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
         MICON           =   "FrmContReporteEFBalanceGeneral.frx":1002
         PICN            =   "FrmContReporteEFBalanceGeneral.frx":101E
         PICH            =   "FrmContReporteEFBalanceGeneral.frx":11E7
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
         Left            =   3480
         TabIndex        =   8
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
         MICON           =   "FrmContReporteEFBalanceGeneral.frx":141C
         PICN            =   "FrmContReporteEFBalanceGeneral.frx":1438
         PICH            =   "FrmContReporteEFBalanceGeneral.frx":16CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Modificar Configuración"
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
         MICON           =   "FrmContReporteEFBalanceGeneral.frx":1857
         PICN            =   "FrmContReporteEFBalanceGeneral.frx":1873
         PICH            =   "FrmContReporteEFBalanceGeneral.frx":1B02
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "FrmContReporteEFBalanceGeneral.frx":1F43
         Left            =   1080
         List            =   "FrmContReporteEFBalanceGeneral.frx":1FA4
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3240
         Width           =   5535
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FrmContReporteEFBalanceGeneral.frx":201B
         Left            =   1080
         List            =   "FrmContReporteEFBalanceGeneral.frx":207C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2880
         Width           =   5535
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FrmContReporteEFBalanceGeneral.frx":20F3
         Left            =   1080
         List            =   "FrmContReporteEFBalanceGeneral.frx":2154
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   5535
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmContReporteEFBalanceGeneral.frx":21CB
         Left            =   1080
         List            =   "FrmContReporteEFBalanceGeneral.frx":222F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2160
         Width           =   5535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteEFBalanceGeneral.frx":22A8
         Left            =   1920
         List            =   "FrmContReporteEFBalanceGeneral.frx":230C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir cuentas con saldo CERO ""0"""
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   53280771
         CurrentDate     =   40254
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración para las cuentas de:"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   2460
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orden:"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   3300
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capital:"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   2940
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pasivos:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   2580
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activos:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   2220
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Jerarquía:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período del balance:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   450
         Width           =   1485
      End
   End
End
Attribute VB_Name = "FrmContReporteEFBalanceGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsTemp2 As Recordset
Dim RsTemp3 As Recordset
Dim i As Integer
Dim j As Integer
Dim Spdr As String
Dim FormatoPDC As String
Dim TamDMGrid As Integer

Sub LeerConfig()
CSql = "SELECT * FROM ContPDCConfig WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 0

If RsTemp.RecordCount = 0 Then Exit Sub

For i = 0 To Combo2.ListCount - 1
    If InStr(1, Combo2.List(i), Trim(RsTemp.Fields("CtaActivos").Value)) = 1 Then
        Combo2.ListIndex = i
        Exit For
    End If
Next i
For i = 0 To Combo3.ListCount - 1
    If InStr(1, Combo3.List(i), Trim(RsTemp.Fields("CtaPasivos").Value)) = 1 Then
        Combo3.ListIndex = i
        Exit For
    End If
Next i
For i = 0 To Combo4.ListCount - 1
    If InStr(1, Combo4.List(i), Trim(RsTemp.Fields("CtaCapital").Value)) = 1 Then
        Combo4.ListIndex = i
        Exit For
    End If
Next i
For i = 0 To Combo5.ListCount - 1
    If InStr(1, Combo5.List(i), Trim(RsTemp.Fields("CtaOrden").Value)) = 1 Then
        Combo5.ListIndex = i
        Exit For
    End If
Next i

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
On Error GoTo MostrarError

Dim resp As Byte

If UCase(BtnGuardarActualizar.Caption) = UCase("Modificar Configuración") Then
    BtnGuardarActualizar.Caption = "Guardar Configuración"
    Combo2.Locked = False
    Combo3.Locked = False
    Combo4.Locked = False
    Combo5.Locked = False
Else
    resp = MsgBox("Se procedará a guardar la configuración elegida!" & Chr(13) & "Desea Continuar?", vbQuestion + vbYesNo, "Confirmar!")

    If resp = vbNo Then
        ' Leer la config de la base de datos...
        LeerConfig
        Exit Sub
    End If
    
    Dim CtaActiv As String
    Dim CtaPasiv As String
    Dim CtaCapit As String
    Dim CtaOrden As String
    
    CtaActiv = Trim(Mid(Combo2.List(Combo2.ListIndex), 1, InStr(1, Combo2.List(Combo2.ListIndex), " ")))
    CtaPasiv = Trim(Mid(Combo3.List(Combo3.ListIndex), 1, InStr(1, Combo3.List(Combo3.ListIndex), " ")))
    CtaCapit = Trim(Mid(Combo4.List(Combo4.ListIndex), 1, InStr(1, Combo4.List(Combo4.ListIndex), " ")))
    CtaOrden = Trim(Mid(Combo5.List(Combo5.ListIndex), 1, InStr(1, Combo5.List(Combo5.ListIndex), " ")))
    
    CSql = "UPDATE ContPDCConfig SET CtaActivos='" & CtaActiv & "',CtaPasivos='" & CtaPasiv & _
        "',CtaCapital='" & CtaCapit & "',CtaOrden='" & CtaOrden & "' WHERE IdEmpresa=" & IdEmprs
    Set RsTemp = CrearRS(CSql)
    
    MsgBox "La Configuración ha sido guardada!", vbInformation + vbOKOnly, "Operación Exitosa!"
    
    BtnGuardarActualizar.Caption = "Modificar Configuración"
    LeerConfig
    Combo2.Locked = True
    Combo3.Locked = True
    Combo4.Locked = True
    Combo5.Locked = True
End If

Exit Sub
MostrarError:
MsgBox "Ha habido un error interno!" & Chr(13) & "Detalles del error." & Chr(13) & Err.Number & ":" & Err.Description & " / " & Err.Source

End Sub

Private Sub BtnResultados_Click()

Dim Simbolo As String
Dim StrTemp As String
Dim FormatoCombo(0 To 3) As String

Dim NForma As String
Dim NroEspacio As Integer
Dim CantDebe As Double
Dim CantHaber As Double
Dim CantCuenta As Double
Dim Band As Boolean
Dim AgregoTotal As Boolean

Dim Ciclo As Byte
Dim ContaReporte As Integer

If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0: Combo2.ListIndex = Combo2.ListCount - 1

CSql = "SELECT * FROM ContPDC WHERE IdEmpresa=" & IdEmprs & " AND Activo='1' ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

With FrmContReportes

.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 6
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Nombre de la Cuenta"
.DMGrid1.DColumnas(2).Caption = "  "
.DMGrid1.DColumnas(3).Caption = "  "
.DMGrid1.DColumnas(4).Caption = "  "
.DMGrid1.DColumnas(5).Visible = False
.DMGrid1.DColumnas(6).Visible = False
.DMGrid1.DColumnas(6).Caption = "Formato"

.DMGrid1.DColumnas(2).Alignment = 1
.DMGrid1.DColumnas(3).Alignment = 1
.DMGrid1.DColumnas(4).Alignment = 1

'.DMGrid1.DColumnas(2).IsNumber = True
'.DMGrid1.DColumnas(3).IsNumber = True
'.DMGrid1.DColumnas(4).IsNumber = True
.DMGrid1.DColumnas(5).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 50 / 100)
'.DMGrid1.DColumnas(6).Width = Val(.DMGrid1.Width * 25 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 15 / 100)
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 15 / 100)
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 20 / 100) - 350

Band = False
ContaReporte = 0
NForma = ""

FormatoCombo(0) = Trim(Mid(Combo2.List(Combo2.ListIndex), 1, InStr(1, Combo2.List(Combo2.ListIndex), " ")))
FormatoCombo(1) = Trim(Mid(Combo3.List(Combo3.ListIndex), 1, InStr(1, Combo3.List(Combo3.ListIndex), " ")))
FormatoCombo(2) = Trim(Mid(Combo4.List(Combo4.ListIndex), 1, InStr(1, Combo4.List(Combo4.ListIndex), " ")))
FormatoCombo(3) = Trim(Mid(Combo5.List(Combo5.ListIndex), 1, InStr(1, Combo5.List(Combo5.ListIndex), " ")))

CSql = "SELECT Simbolo FROM ContPDCConfig WHERE IdEmpresa=" & IdEmprs & "AND Activo='1'"
Set RsTemp2 = CrearRS(CSql)

Simbolo = ""
If RsTemp2.RecordCount <> 0 Then Simbolo = RsTemp2.Fields("Simbolo").Value

CSql = "SELECT ContComprobantesReng.*, ContPDC.Nombre FROM ContComprobantesReng INNER JOIN ContComprobantes ON " & _
    " ContComprobantesReng.IdComprobante = ContComprobantes.IdComprobante INNER JOIN " & _
    " ContPDC ON ContComprobantesReng.Formato = ContPDC.Identificador " & _
    " WHERE ContComprobantes.Activo = '1' AND ContComprobantesReng.IdEmpresa = " & IdEmprs & _
    " AND ContComprobantes.Fecha <='" & Format(DateSerial(Year(CDate(DTPicker1.Value)), Month(CDate(DTPicker1.Value)) + 1, 0), "dd/MM/yyyy") & _
    "' ORDER BY ContComprobantesReng.Formato"
Set RsTemp2 = CrearRS(CSql)

If RsTemp2.RecordCount = 0 Then MsgBox "No hay comprobantes registrados!", vbInformation + vbOKOnly, "No hay registros!": Exit Sub

.DMGrid1.Rows = .DMGrid1.Rows + 1
.DMGrid1.RowBackColor 1, RGB(255, 255, 255)

Dim UltLinea As Integer
UltLinea = 1

For Ciclo = 0 To 3

    If FormatoCombo(Ciclo) <> "" Then
        CSql = "SELECT * FROM ContPDC WHERE Identificador LIKE '" & FormatoCombo(Ciclo) & "%' AND IdEmpresa=" & IdEmprs & " AND Activo='1' ORDER BY Identificador"
        Set RsTemp = CrearRS(CSql)
        
        While Not RsTemp.EOF
            'If Len(Trim(RsTemp.Fields("Identificador").Value)) <= Combo1.ItemData(Combo1.ListIndex) Then
                ' Ciclo que recorre desde el comprobante 1 hasta "N"
                RsTemp2.MoveFirst
                Band = False
                While Not RsTemp2.EOF
                    
                    ' Comprueba si el renglon del comprobante pertenece a la cuenta de activos...
                    If InStr(1, RsTemp2.Fields("Formato").Value, FormatoCombo(Ciclo)) = 1 Then
                    
                        ' Comprueba que la cuenta del renglon sea exactamente igual a la de la cuenta del P.D.C.
                        If Trim(RsTemp.Fields("Identificador").Value) = Trim(RsTemp2.Fields("Formato").Value) Then
                            
                            If Val(RsTemp2.Fields("Tipo").Value) = 0 Then
                                CantDebe = CantDebe + CDbl(RsTemp2.Fields("Cantidad").Value)
                            Else
                                CantHaber = CantHaber + CDbl(RsTemp2.Fields("Cantidad").Value)
                            End If
                            
                            NroEspacio = 0
                            StrTemp = Trim(RsTemp.Fields("Identificador").Value)
                            For i = 1 To Len(StrTemp)
                                If Not IsNumeric(Mid(StrTemp, i, 1)) Then NroEspacio = NroEspacio + 1
                            Next i
                            Band = True
                        End If
                        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                    End If
                    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                    RsTemp2.MoveNext
                Wend
                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                
                If Band = True Then
                    ' verifica que la cuenta pertenezca a la de activos
                    If InStr(1, Trim(RsTemp.Fields("Identificador").Value), FormatoCombo(Ciclo)) = 1 Then
                        
                        Dim StrTemp2 As String
                        ' Ciclo para agregar los niveles anteriores, en caso de que ya se hallan agregado entonces los omite...
                        TamDMGrid = .DMGrid1.Rows
                        StrTemp2 = Trim(RsTemp.Fields("Identificador").Value)
                        
                        For j = 1 To Len(StrTemp2)
                        
                            StrTemp = Mid(StrTemp2, j, 1)
                            If Not IsNumeric(StrTemp) Then
                            
                                StrTemp = Mid(StrTemp2, 1, j - 1)
                                
                                Band = False
                                For i = 1 To TamDMGrid
                                    'MsgBox "Valor actual = " & StrTemp2 & Chr(13) & " " & StrTemp & Chr(13) & " " & .DMGrid1.ValorCelda(i, 6)
                                    If StrTemp = .DMGrid1.ValorCelda(i, 6) Then
                                        Band = True
                                        Exit For
                                    End If
                                Next i
                                
                                If Band = False Then
                                    CSql = "SELECT Nombre,Identificador,Movimiento FROM ContPDC WHERE IdEmpresa=" & IdEmprs & " AND Identificador='" & Mid(StrTemp, 1, j - 1) & "'"
                                    Set RsTemp3 = CrearRS(CSql)
                                    
                                    If Not Trim(.DMGrid1.ValorCelda(.DMGrid1.Rows, 1)) = "" Then .DMGrid1.Rows = .DMGrid1.Rows + 1
                                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = DuplicaChr("    ", ObtenerNivel(Trim(RsTemp3.Fields("Identificador").Value))) & Trim(RsTemp3.Fields("Nombre").Value)
                                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = Trim(RsTemp3.Fields("Identificador").Value)
                                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = 0
                                    .DMGrid1.Rows = .DMGrid1.Rows + 1
                                    
                                End If
                            End If
                        Next j
                        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                        
                        ' verifica si muestra o nó las cuentas en cero "0", si check1=0 No mostrara las cuentas en cero "0"
                        If Check1.Value = 0 Then
                            If CantDebe <> 0 Or CantHaber <> 0 Then
                                .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = DuplicaChr("    ", ObtenerNivel(Trim(RsTemp.Fields("Identificador").Value))) & Trim(RsTemp.Fields("Nombre").Value)
                                If CantDebe > CantHaber Then
                                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = Format(CantDebe, "#,##0.00")
                                Else
                                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = Format(CantHaber * -1, "#,##0.00")
                                End If
                                .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = 1
                                .DMGrid1.Rows = .DMGrid1.Rows + 1
                            End If
                        Else
                            .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = DuplicaChr("    ", ObtenerNivel(Trim(RsTemp.Fields("Identificador").Value))) & Trim(RsTemp.Fields("Nombre").Value)
                            If CantDebe > CantHaber Then
                                .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = Format(CantDebe, "#,##0.00")
                            Else
                                .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = Format(CantHaber * -1, "#,##0.00")
                            End If
                            .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = 1
                            .DMGrid1.Rows = .DMGrid1.Rows + 1
                        End If
                        
                        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                    End If
                End If
                    
                
                CantDebe = 0
                CantHaber = 0
            
            NForma = Trim(RsTemp.Fields("Identificador").Value)
            'End If
            
            RsTemp.MoveNext
        Wend
    
        UltLinea = Agregar_Totales(UltLinea - Ciclo)
        .DMGrid1.PaintMGrid
    End If
Next Ciclo

Tipo = "ReporteBalanceGeneral"
.Show vbModal, FrmPrincipal

End With
End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function Agregar_Totales(IniFor As Integer) As Integer
On Error Resume Next
Dim StrTemp As String
Dim StrTemp2 As String
Dim ArrayDatos(0 To 500, 0 To 3) As String
Dim ContInter As Integer
Dim ContSecuen As Integer
Dim Cantid As Double

CSql = "SELECT Separador, Formato FROM ContPDCConfig WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Function

StrTemp = RsTemp.Fields("Separador").Value
StrTemp2 = RsTemp.Fields("Formato").Value

With FrmContReportes
    
    TamDMGrid = .DMGrid1.Rows
    
    For i = IniFor To TamDMGrid
        StrTemp = .DMGrid1.ValorCelda(i, 6)
        
        If StrTemp <> "" And Val(.DMGrid1.ValorCelda(i, 5)) = 0 Then
            For j = i + 1 To TamDMGrid
            
            
                StrTemp2 = .DMGrid1.ValorCelda(j, 6)
                Cantid = Cantid + CDbl(.DMGrid1.ValorCelda(j, 2))
                
                If (Len(StrTemp) = Len(StrTemp2)) Or j = TamDMGrid Then
                    
                    If ObtenerNivel(StrTemp) = 2 Then
                        ArrayDatos(ContInter, 1) = 3
                    ElseIf ObtenerNivel(StrTemp) > 2 Then
                        ArrayDatos(ContInter, 1) = 2
                    Else
                        ArrayDatos(ContInter, 1) = 4
                    End If
                    
                    
                    ArrayDatos(ContInter, 0) = j
                    ArrayDatos(ContInter, 2) = DuplicaChr("    ", ObtenerNivel(StrTemp)) & "Total para " & Trim(.DMGrid1.ValorCelda(i, 1))
                    ArrayDatos(ContInter, 3) = Format(Cantid, "#,##0.00")
                    
                    'MsgBox " Fila:" & ArrayDatos(ContInter, 0) & Chr(13) & " Columna:" & ArrayDatos(ContInter, 1) & Chr(13) & _
                    '    " Mensaje:" & ArrayDatos(ContInter, 2) & Chr(13) & " Cantidad " & ArrayDatos(ContInter, 3) & Chr(13)
                    ContInter = ContInter + 1
                    Cantid = 0
                    Exit For
                End If
                
                
                
            Next j
        End If
    Next i
    
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' Organizar e incrementar el num. de orden de los totales
    Dim k As Integer
    Dim ArrayTemp(500) As Integer
    Dim NLinea As Integer
    
    For i = 0 To ContInter
        k = 0
        For j = 0 To ContInter
            If j <> i Then
                If Val(ArrayDatos(i, 0)) >= Val(ArrayDatos(j, 0)) And Val(ArrayDatos(j, 0)) <> 0 Then
                    If Val(ArrayDatos(i, 0)) = Val(ArrayDatos(j, 0)) Then
                        If j > i Then k = k + 1
                    Else
                        k = k + 1
                    End If
                End If
            End If
        Next j
        ArrayTemp(i) = Val(ArrayDatos(i, 0)) + k
        If ArrayTemp(i) > NLinea Then NLinea = ArrayTemp(i)
    Next i
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' Ciclo para asignarles los valores modificiados
    For i = 0 To ContInter
        ArrayDatos(i, 0) = ArrayTemp(i)
    Next i
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' Agrega los totales al DMGrid1
    
    NLinea = NLinea - .DMGrid1.Rows
    
    If NLinea > 0 Then
        For i = 1 To NLinea
            .DMGrid1.Rows = .DMGrid1.Rows + 1
        Next i
    End If
    For i = 0 To ContInter
        If Val(ArrayDatos(i, 0)) > 1 Then
            .DMGrid1.RowInsert Val(ArrayDatos(i, 0))
            .DMGrid1.ValorCelda(Val(ArrayDatos(i, 0)), 1) = ArrayDatos(i, 2)
            .DMGrid1.ValorCelda(Val(ArrayDatos(i, 0)), Val(ArrayDatos(i, 1))) = ArrayDatos(i, 3)
            .DMGrid1.RowBackColor Val(ArrayDatos(i, 0)), RGB(221, 221, 221)
        End If
    Next i
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

.DMGrid1.PaintMGrid
TamDMGrid = .DMGrid1.Rows

For i = 1 To TamDMGrid
    If Trim(.DMGrid1.ValorCelda(i, 1)) = "" Then .DMGrid1.RowDelete i
Next i

.DMGrid1.PaintMGrid
.DMGrid1.Rows = .DMGrid1.Rows + 1
Agregar_Totales = .DMGrid1.Rows + 1

End With

End Function

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function ObtenerNivel(Cad As String) As Integer
Dim ii As Integer
Dim jj As Integer
jj = 0
For ii = 1 To Len(Cad)
    If Not IsNumeric(Mid(Cad, ii, 1)) Then jj = jj + 1
Next ii
ObtenerNivel = jj + 1
End Function
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function DuplicaChr(Caracter As String, jj As Integer) As String
Dim Devolver As String
Dim ii As Integer

For ii = 1 To jj
    Devolver = Devolver & Caracter
Next ii
DuplicaChr = Devolver
End Function

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function EsNumerico(ByVal Cad As String) As Boolean
Dim cont As Byte
Dim ii As Integer

If IsNull(Cad) Then EsNumerico = False: Exit Function
If Trim(Cad) = "" Then EsNumerico = False: Exit Function

For ii = 1 To Len(Cad)
    If Not IsNumeric(Mid(Cad, ii, 1)) Then cont = cont + 1
Next ii

If cont > 1 Then EsNumerico = False Else EsNumerico = True

End Function
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Private Sub Form_Load()

Centrar Me
CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMM Carga los niveles del PDC MMMMMMMMMMMMMMMMMMMMMMMM
  CSql = "SELECT * FROM ContPDCConfig WHERE IdEmpresa=" & IdEmprs
  Set RsTemp = CrearRS(CSql)

  If RsTemp.RecordCount <> 0 Then
      FormatoPDC = Trim(RsTemp.Fields("Formato"))
      Spdr = Trim(RsTemp.Fields("Separador"))
  End If

  j = 0
  Combo1.Clear
  Combo1.AddItem " "
  For i = 1 To Len(FormatoPDC)
    If Mid(FormatoPDC, i, 1) = Spdr Then
        j = j + 1
        Combo1.AddItem "Nivel " & j
        Combo1.ItemData(Combo1.NewIndex) = InStr(i, FormatoPDC, Spdr, vbTextCompare)
    End If
  Next i
  Combo1.ListIndex = 0

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMM Lista de las Cuentas MMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "SELECT * FROM ContPDC WHERE IdEmpresa=" & IdEmprs & " AND Activo='1' ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

'DTPicker1.Value = "01" & Format(Now, "/MM/yyyy")
DTPicker1.Value = Format(DateSerial(Year(CDate(Now)), Month(CDate(Now)) + 1, 0), "dd/MM/yyyy")

Combo2.Clear
Combo3.Clear
Combo4.Clear
Combo5.Clear

Combo2.AddItem " "
Combo3.AddItem " "
Combo4.AddItem " "
Combo5.AddItem " "

If RsTemp.RecordCount = 0 Then Exit Sub

While Not RsTemp.EOF
    Combo2.AddItem Trim(RsTemp.Fields("Identificador").Value) & "     " & Trim(RsTemp.Fields("Nombre").Value)
    Combo2.ItemData(Combo2.NewIndex) = RsTemp.Fields("IdPDC").Value
    Combo3.AddItem Trim(RsTemp.Fields("Identificador").Value) & "     " & Trim(RsTemp.Fields("Nombre").Value)
    Combo3.ItemData(Combo3.NewIndex) = RsTemp.Fields("IdPDC").Value
    Combo4.AddItem Trim(RsTemp.Fields("Identificador").Value) & "     " & Trim(RsTemp.Fields("Nombre").Value)
    Combo4.ItemData(Combo4.NewIndex) = RsTemp.Fields("IdPDC").Value
    Combo5.AddItem Trim(RsTemp.Fields("Identificador").Value) & "     " & Trim(RsTemp.Fields("Nombre").Value)
    Combo5.ItemData(Combo5.NewIndex) = RsTemp.Fields("IdPDC").Value
    RsTemp.MoveNext
Wend
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

LeerConfig

End Sub

VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContListaImportarPDC 
   BackColor       =   &H00EAEFEF&
   Caption         =   "Importar Plan de Cuenta de la Empresa"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   Icon            =   "FrmContListaImportarPDC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   13245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   5760
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   5760
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   6120
         Width           =   4935
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Código, Nombre y Rif"
            Top             =   240
            Width           =   2775
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Busqueda"
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
            MICON           =   "FrmContListaImportarPDC.frx":1002
            PICN            =   "FrmContListaImportarPDC.frx":101E
            PICH            =   "FrmContListaImportarPDC.frx":1283
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Rif"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   5
         Top             =   5760
         Width           =   1095
      End
      Begin SystemOncoAmerica.DMGrid DMGrid2 
         Height          =   5415
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9551
         Object.Width           =   5025
         Object.Height          =   5385
         ScrollBar       =   1
         DrawColorGrid   =   1
         MarqueeStyle    =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   5760
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Plan de Cuenta a exportar"
      Height          =   6975
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   6120
         Width           =   7215
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   5760
            Top             =   480
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6120
            TabIndex        =   2
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
            MICON           =   "FrmContListaImportarPDC.frx":1515
            PICN            =   "FrmContListaImportarPDC.frx":1531
            PICH            =   "FrmContListaImportarPDC.frx":16FA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnImportarTodo 
            Height          =   375
            Left            =   240
            TabIndex        =   15
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Importar"
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
            MICON           =   "FrmContListaImportarPDC.frx":192F
            PICN            =   "FrmContListaImportarPDC.frx":194B
            PICH            =   "FrmContListaImportarPDC.frx":1BD5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregar 
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Agregar"
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
            MICON           =   "FrmContListaImportarPDC.frx":1E5E
            PICN            =   "FrmContListaImportarPDC.frx":1E7A
            PICH            =   "FrmContListaImportarPDC.frx":2007
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
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
         Object.Width           =   7305
         Object.Height          =   2745
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin SystemOncoAmerica.DMGrid DMGrid3 
         Height          =   2535
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4471
         Object.Width           =   7305
         Object.Height          =   2505
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan de Cuentas seleccionados:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   2280
      End
   End
End
Attribute VB_Name = "FrmContListaImportarPDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsPDC As Recordset
Dim IdReng As Integer
Dim IdEmpresa As Integer
Dim IdPDC As Integer
Dim Spdr As Integer
Dim Formato As String
Dim i As Integer

Sub Cargar_Empresa_Configuradas()
CSql = "SELECT ContEmpresas.IdEmpresa, ContEmpresas.Nombre, ContEmpresas.Rif FROM ContEmpresas INNER JOIN " & _
        " ContPDC ON ContEmpresas.IdEmpresa = ContPDC.IdEmpresa INNER JOIN ContPDCConfig ON " & _
        " ContPDC.IdEmpresa = ContPDCConfig.IdEmpresa GROUP BY ContEmpresas.IdEmpresa, ContEmpresas.Nombre, ContEmpresas.Rif"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then MsgBox "No se encontraron Empresas con una configuracion y plan de cuentas creadas!", vbExclamation + vbOKOnly, "Error": Exit Sub

DMGrid1.Rows = 0
DMGrid2.Rows = 0
While Not RsTemp.EOF
    
    DMGrid2.Rows = DMGrid2.Rows + 1
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsTemp.Fields("IdEmpresa").Value
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsTemp.Fields("Nombre").Value
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsTemp.Fields("Rif").Value
    
    RsTemp.MoveNext
Wend
DMGrid1.PaintMGrid
DMGrid2.PaintMGrid

Cargar_Plan_De_Cuentas
End Sub

Sub Cargar_Plan_De_Cuentas()
CSql = "SELECT ContEmpresas.IdEmpresa, ContEmpresas.Nombre, ContEmpresas.Rif FROM ContEmpresas INNER JOIN " & _
        " ContPDC ON ContEmpresas.IdEmpresa = ContPDC.IdEmpresa INNER JOIN ContPDCConfig ON " & _
        " ContPDC.IdEmpresa = ContPDCConfig.IdEmpresa GROUP BY ContEmpresas.IdEmpresa, ContEmpresas.Nombre, ContEmpresas.Rif"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then MsgBox "No se encontraron Empresas con una configuracion y plan de cuentas creadas!", vbExclamation + vbOKOnly, "Error": Exit Sub

DMGrid1.Rows = 0
While Not RsTemp.EOF
    
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("IdEmpresa").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Nombre").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Rif").Value
    
    RsTemp.MoveNext
Wend
DMGrid1.PaintMGrid

End Sub

Sub Blanqueo()
    DMGrid2.Clear
    DMGrid2.Rows = 0
    DMGrid2.PaintMGrid
End Sub
Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    CSql = "Select * From ContPDC where activo=1 order by Identificador"
ElseIf Index = 1 Then
    CSql = "Select * From ContPDC where activo=1 order by Nombre"
ElseIf Index = 2 Then
    CSql = "Select * From ContPDC where activo=1 order by Tipo"
End If

Set RsPDC = CrearRS(CSql)
DMGrid1.Rows = 0

RsPDC.MoveFirst

While Not RsPDC.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsPDC.Fields("Identificador")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsPDC.Fields("Nombre")
    
    If RsPDC.Fields("Movimiento").Value Then
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = "Mvto"
    Else
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = "Grupo"
    End If
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPDC.Fields("IdPDC")
    RsPDC.MoveNext
Wend
DMGrid1.PaintMGrid
End Sub

Private Sub DMGrid2_DobleClick()
Dim TamCad As Byte
    IdReng = 0
    Blanqueo
    If DMGrid2.Row = 0 Then Exit Sub
    IdEmpresa = DMGrid2.ValorCelda(DMGrid2.Row, 1)

    Frame3.Caption = "Configuracion del PDC para " & DMGrid2.ValorCelda(DMGrid2.Row, 2)
    CSql = "Select * From ContPDCConfig WHERE IdEmpresa=" & IdEmpresa & " And activo=1"
    Set RsTemp = CrearRS(CSql)
    DMGrid2.Rows = 0
    If RsTemp.RecordCount <> 0 Then
        IdPDC = Val(RsTemp.Fields("IdEmpresa").Value)
        Spdr = Asc(RsTemp.Fields("Separador").Value)
        BtnAgregar.Enabled = True
        
        'TxtFormato.Mask = Replace(Trim(RsTemp.Fields("Formato").Value), "X", "#")
        'TxtFormato1.Mask = Replace(Trim(RsTemp.Fields("Formato").Value), "X", "#")
        'TxtFormato2.Mask = Replace(Trim(RsTemp.Fields("Formato").Value), "X", "#")
        Formato = Trim(RsTemp.Fields("Formato").Value)
        
        Frame3.Enabled = True
        
        ' Ahora consultara la Base de datos para mostrar el plan de cuentas de la empresa (En el caso de que
        ' tenga una ya creada)
        CSql = "Select * From ContPDC WHERE IdEmpresa=" & IdEmpresa & " And activo=1 order by IdPDC"
        Set RsTemp = CrearRS(CSql)
        
        If RsTemp.RecordCount <> 0 Then
            
            'TxtFormato.Text = Format(Replace(Trim(RsTemp.Fields("Identificador").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
            'TxtFormato1.Text = Format(Replace(Trim(RsTemp.Fields("CuentaAjusta").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
            'TxtFormato2.Text = Format(Replace(Trim(RsTemp.Fields("CuentaCorreccion").Value), Chr(Spdr), ""), Replace(Trim(Formato), "X", "#"))
            
            'If RsTemp.Fields("Movimiento").Value Then Check1.Value = 1 Else Check1.Value = 0
            'If RsTemp.Fields("Bases").Value Then Check2.Value = 1 Else Check2.Value = 0
            'If RsTemp.Fields("Terceros").Value Then Check3.Value = 1 Else Check3.Value = 0
            'If RsTemp.Fields("CCFijos").Value Then Check4.Value = 1 Else Check4.Value = 0
    
            'For i = 0 To Combo1.ListCount - 1
            '    If Combo1.ItemData(i) = Val(RsTemp.Fields("TipoActividad").Value) Then Combo1.ListIndex = i: Exit For
            'Next i
            'For i = 0 To Combo2.ListCount - 1
            '    If Combo2.ItemData(i) = Val(RsTemp.Fields("Clasificacion").Value) Then Combo2.ListIndex = i: Exit For
            'Next i
            'For i = 0 To Combo3.ListCount - 1
            '    If Combo3.ItemData(i) = Val(RsTemp.Fields("TipoCuenta").Value) Then Combo3.ListIndex = i: Exit For
            'Next i
            
            ' Para CENTROS DE CONSTOS...
            'For i = 0 To Combo1.ListCount - 1
            '    If Combo1.List(i) = RsTemp.Fields("TipoActividad").Value Then Combo1.ListIndex = i: Exit For
            'Next i
            
            ' Ciclo condicional para llenar el DMGrid2 con el plan de cuentas de la empresa selecciona,
            ' si la empresa no tiene un PDC, entonces mostrara todos los campos en blancos.
            While Not RsTemp.EOF
                DMGrid2.Rows = DMGrid2.Rows + 1
                DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsTemp.Fields("Identificador").Value
                DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsTemp.Fields("Nombre").Value
                
                If RsTemp.Fields("Movimiento").Value Then
                    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = "Mvto"
                Else
                    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = "Grupo"
                End If
                DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsTemp.Fields("IdPDC").Value
                RsTemp.MoveNext
            Wend
            Option2_Click (0)
        Else
        End If
    Else
        MsgBox "La Empresa '" & DMGrid2.ValorCelda(DMGrid2.Row, 2) & "' no contiene una Configuración de P.D.C.!", vbExclamation + vbOKOnly, "No tiene un Plan de Cuenta!"
        Frame3.Enabled = False
        IdPDC = 0
        Spdr = 0
        Formato = ""
        BtnAgregar.Enabled = False
    End If
    DMGrid2.PaintMGrid

End Sub

Private Sub Form_Load()
Blanqueo

Cargar_Empresa_Configuradas
End Sub

Private Sub Option2_Click(Index As Integer)
If Index = 0 Then
    CSql = "Select * From ContPDC where activo=1 order by Identificador"
ElseIf Index = 1 Then
    CSql = "Select * From ContPDC where activo=1 order by Nombre"
ElseIf Index = 2 Then
    CSql = "Select * From ContPDC where activo=1 order by Tipo"
End If

Set RsPDC = CrearRS(CSql)
DMGrid1.Rows = 0

RsPDC.MoveFirst

While Not RsPDC.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsPDC.Fields("Identificador")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsPDC.Fields("Nombre")
    
    If RsPDC.Fields("Movimiento").Value Then
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = "Mvto"
    Else
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = "Grupo"
    End If
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPDC.Fields("IdPDC")
    RsPDC.MoveNext
Wend
DMGrid1.PaintMGrid
End Sub
Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub
Private Sub BtnBuscar_Click()
Dim cn As Integer
End Sub



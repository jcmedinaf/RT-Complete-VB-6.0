VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoBeneficiarios 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beneficiarios"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "FrmListadoBeneficiarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6075
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   3480
         TabIndex        =   4
         Top             =   7080
         Width           =   2295
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   720
            Top             =   240
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   1200
            TabIndex        =   5
            ToolTipText     =   "Cerrar"
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
            MICON           =   "FrmListadoBeneficiarios.frx":1002
            PICN            =   "FrmListadoBeneficiarios.frx":101E
            PICH            =   "FrmListadoBeneficiarios.frx":11E7
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   7080
         Width           =   3255
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
            TabIndex        =   2
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda por Código o Nombre del Beneficiario"
            Top             =   240
            Width           =   1575
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "FrmListadoBeneficiarios.frx":141C
            PICN            =   "FrmListadoBeneficiarios.frx":1438
            PICH            =   "FrmListadoBeneficiarios.frx":169D
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
         Height          =   6735
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   11880
         Object.Width           =   5625
         Object.Height          =   6705
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoBeneficiarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCargarBeneficiario As New ADODB.Recordset
Dim RsSeleccionarBeneficiario As New ADODB.Recordset
Dim RsBuscarBeneficiario As New ADODB.Recordset

Private Sub BtnBuscar_Click()
If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Beneficiario Where (CodigoBeneficiario = '" & Val(Trim(TxtBuscar.Text)) & "' or DescripcionBeneficiario like '%" & Trim(TxtBuscar.Text) & "%')"
Else
    CSql = "Select * From Beneficiario"
End If

Set RsBuscarBeneficiario = CrearRS(CSql)

If RsBuscarBeneficiario.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarBeneficiario.EOF
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsBuscarBeneficiario.Fields("CodigoBeneficiario").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsBuscarBeneficiario.Fields("DescripcionBeneficiario").Value
            RsBuscarBeneficiario.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly + vbCritical, "Sin Resultado"
    Exit Sub
End If
RsBuscarBeneficiario.Close
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    
    CSql = "Select * From Beneficiario Where CodigoBeneficiario='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsSeleccionarBeneficiario = CrearRS(CSql)
    If RsSeleccionarBeneficiario.RecordCount > 0 Then
                       
        FrmTransaccionCheques.TxtPaguese.Text = RsSeleccionarBeneficiario.Fields("DescripcionBeneficiario").Value
    
    End If
    RsSeleccionarBeneficiario.Close
    Unload Me
End If
End Sub

Private Sub Form_Load()
Grid1


CSql = "Select * From Beneficiario"
Set RsCargarBeneficiario = CrearRS(CSql)
If RsCargarBeneficiario.RecordCount > 0 Then
Do While Not RsCargarBeneficiario.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarBeneficiario.Fields("CodigoBeneficiario").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarBeneficiario.Fields("DescripcionBeneficiario").Value
        RsCargarBeneficiario.MoveNext
Loop
End If
DMGrid1.PaintMGrid

End Sub

Sub Grid1()
DMGrid1.Rows = 1
DMGrid1.Cols = 2
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0

DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True

DMGrid1.DColumnas(1).Width = 1500
DMGrid1.DColumnas(2).Width = 3800

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre del Beneficiario"

End Sub


Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
Else
    If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub

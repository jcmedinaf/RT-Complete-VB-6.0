VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmRetencionesCobros 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retenciones Hechas al Cobro"
   ClientHeight    =   3930
   ClientLeft      =   6645
   ClientTop       =   4080
   ClientWidth     =   4800
   DrawWidth       =   10
   Icon            =   "retenciones.frx":0000
   LinkTopic       =   "Form36"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4800
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   4575
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3480
         TabIndex        =   12
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
         MICON           =   "retenciones.frx":1002
         PICN            =   "retenciones.frx":101E
         PICH            =   "retenciones.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardar 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Guardar / Actualizar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar"
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
         MICON           =   "retenciones.frx":141C
         PICN            =   "retenciones.frx":1438
         PICH            =   "retenciones.frx":16C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnDesHacer 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Deshacer"
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
         MICON           =   "retenciones.frx":1B08
         PICN            =   "retenciones.frx":1B24
         PICH            =   "retenciones.frx":1E06
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
      Caption         =   "Datos de la Retención"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "retenciones.frx":2057
         Left            =   1440
         List            =   "retenciones.frx":2067
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53542913
         CurrentDate     =   39941
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Retención:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2003
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2490
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Comprobante:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fact.:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmRetencionesCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bdata As New ADODB.Recordset

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardar_Click()
If Text2.Text = "" Then Exit Sub
If Text3.Text = "" Then Exit Sub

Dim RsCorr As New ADODB.Recordset
CSql = "Select Max(IdCobro)+1 as MaxIdCobros From Cobros"
Set RsCorr = CrearRS(CSql)

CSql = "Select * From Cobros"
Set bdata = CrearRS(CSql)
Select Case op
    Case Is = "Cobros"
        bdata.AddNew
        bdata.Fields("IdCobro").Value = RsCorr.Fields("MaxIdCobros").Value
        bdata.Fields("tipo_ret").Value = Combo1.List(Combo1.ListIndex)
        bdata.Fields("n_factura").Value = Text1.Text
        bdata.Fields("fecha_cob").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        bdata.Fields("monto").Value = Val(Text3.Text)
        bdata.Fields("tipo").Value = 99
        bdata.Fields("n_retencion").Value = Text2.Text
        bdata.Fields("IdUser").Value = IdUser
        bdata.Fields("N_fa").Value = 0
        bdata.Fields("N_Nc").Value = 0
        bdata.Fields("C_Nc").Value = 0
        bdata.Update

        Msg = "Retención Agregada satisfactoriamente"
        MsgBox Msg, vbInformation + vbOKOnly, "Mensaje"

        FrmAsientoCobros.Carga_Retencion
        FrmAsientoCobros.calcular
        
    Case Is = "Pagos"
        bdata.AddNew
        bdata.Fields("IdCobro").Value = RsCorr.Fields("MaxIdCobros").Value
        bdata.Fields("tipo_ret").Value = Combo1.List(Combo1.ListIndex)
        bdata.Fields("n_factura").Value = Text1.Text
        bdata.Fields("fecha_cob").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        bdata.Fields("monto").Value = Val(Text3.Text)
        bdata.Fields("tipo").Value = 98
        bdata.Fields("n_retencion").Value = Text2.Text
        bdata.Fields("IdUsuario").Value = IdUser
        bdata.Fields("N_fa").Value = 0
        bdata.Fields("N_Nc").Value = 0
        bdata.Fields("C_Nc").Value = 0
        bdata.Update

        Msg = "Retención Agregada satisfactoriamente"
        MsgBox Msg, vbInformation + vbOKOnly, "Mensaje"
        
        FrmAsientoPagos.Carga_Retencion
        FrmAsientoPagos.calcular
        
End Select
Unload Me
End Sub

Private Sub Form_Load()
Centrar Me
DTPicker1.Value = Now

Select Case op
    Case Is = "Cobros"
        Me.Caption = "Retenciones Hechas al Cobro"
    Case Is = "Pagos"
        Me.Caption = "Retenciones Hechas al Pago"
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Exit Sub
Case Is = 13
Exit Sub
Case 45 To 46
Exit Sub
Case Is = 8
Exit Sub
Case Else
KeyAscii = 0
End Select

End Sub


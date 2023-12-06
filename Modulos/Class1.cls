VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1


Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000


Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or _
                          SWP_NOSIZE Or _
                          SWP_NOMOVE Or _
                          SWP_DRAWFRAME


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private WithEvents mHFlex As MSHFlexGrid
Attribute mHFlex.VB_VarHelpID = -1
Private WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mDtPicker As DtPicker
Attribute mDtPicker.VB_VarHelpID = -1



Private sTemp As String

Enum EFormato
    estring = 0
    eBoolean = 1
    eDate = 2
    enumber = 3
    ecurrency = 4
End Enum

Private Type tcol
    Bloqueado As Boolean
    Formato As EFormato
End Type

Public Event Validar(IRowActual As Long, _
                     IColActual As Integer, _
                     CellValue As Variant, _
                     Cancelar As Boolean)


Enum eEditMode
    eClick = 0
    eDobleClick = 1
End Enum

Private antCol As Integer
Private antRow As Long

Private mEditMode As eEditMode


Private mCol(100) As tcol

Private Sub mDtPicker_Change()
        With mHFlex
            If mCol(.Col).Formato = eDate Then
                .TextMatrix(.Row, .Col) = mDtPicker.Value
            End If
        End With
End Sub

Private Sub mDTPicker_KeyUp(KeyCode As Integer, Shift As Integer)
       
      If mDtPicker.Visible Then
          With mHFlex
                Select Case KeyCode
                    Case 13, 9
                        ValidarDtPicker
                End Select
          End With
      End If
End Sub

Sub ValidarDtPicker()
    
    With mHFlex
        Dim Cancelar As Boolean
        
        RaiseEvent Validar(antRow, antCol, mDtPicker.Value, Cancelar)
        
        If Cancelar Then
           .Row = antRow
           .Col = antCol
           .Text = sTemp
        End If
        
    End With
    
     mDtPicker.Visible = False
    
End Sub

Private Sub mDtPicker_LostFocus()
    mDTPicker_KeyUp 13, 0
End Sub

Private Sub mHFlex_Click()
    If mEditMode = eDobleClick Then
       Exit Sub
    ElseIf mEditMode = eClick Then
       Call mHFlex_DblClick
    End If
End Sub

Private Sub mHFlex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mHFlex.Row <> mHFlex.RowSel Or mHFlex.Col <> mHFlex.ColSel Then
       mHFlex.RowSel = mHFlex.Row
       mHFlex.ColSel = mHFlex.Col
    End If
End Sub

Private Sub mtextbox_LostFocus()

Dim Cancelar As Boolean
    
    Dim opcion As EFormato
    
    opcion = mCol(antCol).Formato
    
    Select Case mCol(antCol).Formato
        Case estring
            RaiseEvent Validar(antRow, antCol, mTextBox.Text, Cancelar)
            
            'If Cancelar Then
            '    mHFlex.Text = sTemp
            'End If
        ' números
        Case enumber
            If Not IsNumeric(mTextBox.Text) Then
               RaiseEvent Validar(antRow, antCol, 0, Cancelar)
               mHFlex.TextMatrix(antRow, antCol) = 0
            Else
               RaiseEvent Validar(antRow, antCol, mTextBox.Text, Cancelar)
            End If
            
        ' currency
        Case ecurrency
            If Not IsNumeric(mTextBox.Text) Then
                RaiseEvent Validar(antRow, antCol, 0, Cancelar)
                mHFlex.TextMatrix(antRow, antCol) = 0
            Else
                RaiseEvent Validar(antRow, antCol, mTextBox.Text, Cancelar)
            End If
            'If Cancelar Then
            '    mHFlex.Text = sTemp
            'End If
    End Select
    
    If Cancelar Then
        mHFlex.Row = antRow
        mHFlex.Col = antCol
        mHFlex.Text = sTemp
    End If
    
    mTextBox.Visible = False

End Sub

Private Sub mHFlex_DblClick()

    Dim Cancelar As Boolean
    Dim iCol As Integer
    
    iCol = mHFlex.Col
    
    If mCol(iCol).Bloqueado Then
        Exit Sub
    End If
    
    With mHFlex
        sTemp = .Text
        antCol = .Col
        antRow = .Row
    End With
    
    
    If mCol(iCol).Formato = eBoolean Then
        With mHFlex
            Dim sBoolean As String
            
            If LCase(.TextMatrix(.Row, .Col)) = LCase("Sí") Then
               RaiseEvent Validar(antRow, antCol, False, Cancelar)
            ElseIf LCase(.TextMatrix(.Row, .Col)) = LCase("No") Then
               RaiseEvent Validar(antRow, antCol, True, Cancelar)
            End If
            
            If Cancelar = False Then
               .TextMatrix(.Row, .Col) = Format(UCase(.TextMatrix(.Row, .Col)) = "NO" Or Trim(.TextMatrix(.Row, .Col)) = vbNullString, "YES/NO")
            End If
        End With
        Exit Sub
    End If
    
    If mCol(iCol).Formato = eDate Then
        With mHFlex
            If Trim(.Text) <> vbNullString Then
                mDtPicker.Value = Format(.Text, "Short Date")
            Else
                mDtPicker.Value = Format(Date, "Short Date")
            End If
        
            mDtPicker.Move (.Left + .CellLeft - 25), _
                           (.Top + (.RowHeight(0) * (.Row - .topRow + 1))), _
                           (.CellWidth - 8), _
                           (.CellHeight + 25)
                                
            mDtPicker.Visible = True
            mDtPicker.Enabled = True
            mDtPicker.SetFocus
            mDtPicker.Value = .TextMatrix(.Row, .Col)
     
            Exit Sub
        End With
    End If


    mTextBox.Visible = False

    With mTextBox
        .Text = mHFlex.Text
        .Top = mHFlex.Top + mHFlex.CellTop - 25
        .Left = mHFlex.Left + mHFlex.CellLeft - 25
        .Width = mHFlex.CellWidth
        .Height = mHFlex.CellHeight
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With


End Sub

Private Sub mHFlex_Scroll()
    mDtPicker.Visible = False
    mTextBox.Visible = False

End Sub

Private Sub mHFlex_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     If mEditMode = eClick Then
        mHFlex_Click
     ElseIf mEditMode = eDobleClick Then
        mHFlex_DblClick
     End If
   End If
End Sub

Private Sub mTextBox_KeyPress(KeyAscii As Integer)
    If mCol(mHFlex.Col).Formato = enumber Then
       If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    If mCol(mHFlex.Col).Formato = ecurrency Then
       If InStr("1234567890" & Chr(8) & Chr(44), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
End Sub



Private Sub mTextBox_Change()
     mHFlex.Text = mTextBox
End Sub

Private Sub mTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 13 Then
    mTextBox.Visible = False
    mHFlex.SetFocus
  ElseIf KeyCode = vbKeyEscape Then
    mHFlex.Text = sTemp
    mTextBox.Visible = False
    mHFlex.SetFocus
  End If

End Sub

Sub Iniciar(HierarchicalFlex As MSHFlexGrid, frm As Form, DtPicker As DtPicker)
    
    If mHFlex Is Nothing Then
       
       'Set mAdoCon = AdoConection
       Set mTextBox = frm.Controls.Add("vb.textbox", "txtEditarCelda")
       If HierarchicalFlex.FixedCols <> 0 Or HierarchicalFlex.FixedRows <> 1 Then
          MsgBox "Este ejemplo solo funciona con la propiedad FixedCol en " & _
                 "0 y FixedRow en 1. Establezca esas propiedades para el " & _
                 "Hierarchical FlexGrid. ( hacerlo en tiempo de diseño)", vbCritical
          Exit Sub
       End If
       
       Set mHFlex = HierarchicalFlex
       
       With mHFlex
            .SelectionMode = flexSelectionFree
            Call Flat(.hwnd)
            .FixedCols = 0
       End With
       
       With mTextBox
            .BackColor = &HC0FFC0
            .ZOrder 0
            .BorderStyle = 0
       End With
       
       Set mDtPicker = DtPicker
       DtPicker.Visible = False
       Call Flat(DtPicker.hwnd)
    Else
       Exit Sub
    End If
End Sub


Private Sub Flat(ByVal hwnd As Long)
    Dim e As Long
        
        e = GetWindowLong(hwnd, GWL_EXSTYLE)
        e = e And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
        SetWindowLong hwnd, GWL_EXSTYLE, e
        SetWindowPos hwnd, 0, 0, 0, 0, 0, _
                     SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED _
                      Or SWP_NOSIZE Or SWP_NOMOVE
End Sub


Property Get EditMode() As eEditMode
    EditMode = mEditMode
End Property

Property Let EditMode(Valor As eEditMode)
    mEditMode = Valor
End Property

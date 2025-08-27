VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NaklForm 
   Caption         =   "Duplicate"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   OleObjectBlob   =   "NaklForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NaklForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    HorizontalGapValue.Text = "5"
    VerticalGapValue.Text = "5"
    LeftBorderValue.Text = "13"
    RightBorderValue.Text = "13"
    TopBorderValue.Text = "20"
    BottomBorderValue.Text = "11"
    MaxObjectsValue.Text = "100"
    MarkerDistanceXValue.Text = "4"
    MarkerDistanceYValue.Text = "4"
    MarkerSizeValue.Text = "3"
    MarkerCountValue.Text = "4"
End Sub


Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OkButton_Click()
    If Not IsNumeric(HorizontalGapValue.Text) Or Val(HorizontalGapValue.Text) <= 0 Then
        MsgBox "Horizontal gap must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(VerticalGapValue.Text) Or Val(VerticalGapValue.Text) <= 0 Then
        MsgBox "Vertical gap must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(LeftBorderValue.Text) Or Val(LeftBorderValue.Text) <= 0 Then
        MsgBox "Left border must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(RightBorderValue.Text) Or Val(RightBorderValue.Text) <= 0 Then
        MsgBox "Right border must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(TopBorderValue.Text) Or Val(TopBorderValue.Text) <= 0 Then
        MsgBox "Top border must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(BottomBorderValue.Text) Or Val(BottomBorderValue.Text) <= 0 Then
        MsgBox "Bottom border must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(MaxObjectsValue.Text) Or Val(MaxObjectsValue.Text) <= 0 Then
        MsgBox "Max objects must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(MarkerDistanceXValue.Text) Or Val(MarkerDistanceXValue.Text) <= 0 Then
        MsgBox "Marker distance X must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(MarkerDistanceYValue.Text) Or Val(MarkerDistanceYValue.Text) <= 0 Then
        MsgBox "Marker distance Y must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(MarkerSizeValue.Text) Or Val(MarkerSizeValue.Text) <= 0 Then
        MsgBox "Marker size must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(MarkerCountValue.Text) Or Val(MarkerCountValue.Text) <= 0 Then
        MsgBox "Marker count must be a positive number."
        Exit Sub
    End If
    
    If MarkerCountValue.Text < 4 Or MarkerCountValue.Text Mod 2 <> 0 Then
        MsgBox "Marker count must be an even number and  4 or greater."
        Exit Sub
    End If
    
    RunDuplicate _
        CDbl(HorizontalGapValue.Text), _
        CDbl(VerticalGapValue.Text), _
        CDbl(LeftBorderValue.Text), _
        CDbl(RightBorderValue.Text), _
        CDbl(TopBorderValue.Text), _
        CDbl(BottomBorderValue.Text), _
        CDbl(MaxObjectsValue.Text), _
        CDbl(MarkerDistanceXValue.Text), _
        CDbl(MarkerDistanceYValue.Text), _
        CDbl(MarkerSizeValue.Text), _
        CDbl(MarkerCountValue.Text)
    
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

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
    LoadButton_Click
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
    If Not IsNumeric(GapSplitValue.Text) Or Val(GapSplitValue.Text) <= 0 Then
        MsgBox "Marker size must be a positive number."
        Exit Sub
    End If
    If Not IsNumeric(GapDistanceValue.Text) Or Val(GapDistanceValue.Text) <= 0 Then
        MsgBox "Marker size must be a positive number."
        Exit Sub
    End If
    Unload Me
    
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
        CBool(IsSplitMode.Value), _
        CDbl(GapSplitValue.Text), _
        CDbl(GapDistanceValue.Text)

    
End Sub

Private Sub LoadButton_Click()
    ' load settings from the registry, if they exist
    HorizontalGapValue.Text = GetSetting("CorelDrawMacros", "UI", "HorizontalGapValue", "5")
    VerticalGapValue.Text = GetSetting("CorelDrawMacros", "UI", "VerticalGapValue", "5")
    LeftBorderValue.Text = GetSetting("CorelDrawMacros", "UI", "LeftBorderValue", "13")
    RightBorderValue.Text = GetSetting("CorelDrawMacros", "UI", "RightBorderValue", "13")
    TopBorderValue.Text = GetSetting("CorelDrawMacros", "UI", "TopBorderValue", "20")
    BottomBorderValue.Text = GetSetting("CorelDrawMacros", "UI", "BottomBorderValue", "11")
    MaxObjectsValue.Text = GetSetting("CorelDrawMacros", "UI", "MaxObjectsValue", "100")
    MarkerDistanceXValue.Text = GetSetting("CorelDrawMacros", "UI", "MarkerDistanceXValue", "4")
    MarkerDistanceYValue.Text = GetSetting("CorelDrawMacros", "UI", "MarkerDistanceYValue", "4")
    MarkerSizeValue.Text = GetSetting("CorelDrawMacros", "UI", "MarkerSizeValue", "3")
    IsSplitMode.Value = GetSetting("CorelDrawMacros", "UI", "IsSplitMode", "True")
    GapSplitValue.Text = GetSetting("CorelDrawMacros", "UI", "GapSplitValue", "10")
    GapDistanceValue.Text = GetSetting("CorelDrawMacros", "UI", "GapDistanceValue", "240")
    
End Sub

Private Sub SaveButton_Click()
    ' saving is done based on the form field's name
    SaveSetting "CorelDrawMacros", "UI", "HorizontalGapValue", HorizontalGapValue.Text
    SaveSetting "CorelDrawMacros", "UI", "VerticalGapValue", VerticalGapValue.Text
    SaveSetting "CorelDrawMacros", "UI", "LeftBorderValue", LeftBorderValue.Text
    SaveSetting "CorelDrawMacros", "UI", "RightBorderValue", RightBorderValue.Text
    SaveSetting "CorelDrawMacros", "UI", "TopBorderValue", TopBorderValue.Text
    SaveSetting "CorelDrawMacros", "UI", "BottomBorderValue", BottomBorderValue.Text
    SaveSetting "CorelDrawMacros", "UI", "MaxObjectsValue", MaxObjectsValue.Text
    SaveSetting "CorelDrawMacros", "UI", "MarkerDistanceXValue", MarkerDistanceXValue.Text
    SaveSetting "CorelDrawMacros", "UI", "MarkerDistanceYValue", MarkerDistanceYValue.Text
    SaveSetting "CorelDrawMacros", "UI", "MarkerSizeValue", MarkerSizeValue.Text
    SaveSetting "CorelDrawMacros", "UI", "IsSplitMode", IsSplitMode.Value
    SaveSetting "CorelDrawMacros", "UI", "GapSplitValue", GapSplitValue.Text
    SaveSetting "CorelDrawMacros", "UI", "GapDistanceValue", GapDistanceValue.Text
    
End Sub
Private Sub UserForm_Click()

End Sub

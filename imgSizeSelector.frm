VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} imgSizeSelector 
   Caption         =   "Image Size"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   OleObjectBlob   =   "imgSizeSelector.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "imgSizeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selectedSize As String

Private Sub cancel_btn_Click()
    Unload imgSizeSelector
End Sub



Private Sub customResOpt_Click()
    customText.Enabled = True
    customText.Value = "Enter Width as Integer"
End Sub

Private Sub customText_Change()

End Sub

Private Sub fullResOpt_Click()
    customText.Value = ""
    customText.Enabled = False
End Sub


Private Sub ok_btn_Click()
    Dim sel As MSforms.OptionButton
    Set sel = GetSelectedOption("sizeOptions")
    
    If Not sel Is Nothing Then
        If sel = customResOpt Then
            selectedSize = customText.Value
        Else
            selectedSize = sel.caption
        End If
    Else
        msgBox ("Please select an option.")
    End If
    Me.Hide
    'Unload imgSizeSelector
    
End Sub

Function GetSelectedOption(strGroupName As String) As MSforms.OptionButton

    Dim ctrl As Control
    Dim opt As MSforms.OptionButton

    'initialise
    Set ctrl = Nothing
    Set GetSelectedOption = Nothing

    'loop controls looking for option button that is
    'both true and part of input GroupName
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.GroupName = strGroupName Then
                Set opt = ctrl
                If opt.Value Then
                    Set GetSelectedOption = opt
                    Exit For
                End If
            End If
        End If
    Next ctrl

End Function


Private Sub OptionButton1_Click()

End Sub



Private Sub smallResOpt_Click()
    customText.Value = ""
    customText.Enabled = False
End Sub

Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()
    fullResOpt.Value = True
    customText.Enabled = False
End Sub

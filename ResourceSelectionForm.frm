VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResourceSelectionForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "ResourceSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResourceSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Module-level variables
Dim SelectedResource1 As Resource
Dim SelectedResource2 As Resource

Private Sub Label1_Click()

End Sub

' ResourceSelectionForm UserForm code
Private Sub UserForm_Initialize()
    ' Populate dropdown lists with resource names
    For Each res In ActiveProject.Resources
        ComboBox1.AddItem res.Name
        ComboBox2.AddItem res.Name
    Next res
End Sub

Private Sub btnOK_Click()
    ' Validate selections
    If ComboBox1.ListIndex < 0 Or ComboBox2.ListIndex < 0 Then
        MsgBox "Please select both resources.", vbExclamation
    Else
        ' Get selected resources
        Set SelectedResource1 = ActiveProject.Resources(ComboBox1.ListIndex + 1)
        Set SelectedResource2 = ActiveProject.Resources(ComboBox2.ListIndex + 1)
        Me.Hide
        
        ' Call the ReplaceResource procedure and pass the selected resources
        ReplaceResource SelectedResource1, SelectedResource2
    End If
End Sub

Private Sub btnCancel_Click()
    ' Close the form without selecting resources
    Me.Hide
End Sub

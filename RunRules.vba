Public Sub RunRules()
    On Error Resume Next
    Dim Session As Outlook.NameSpace
    Dim Report As String
    Dim currentItem As Object
    Dim oRule As Outlook.Rule
    Dim rules As Outlook.rules
    Set Session = Application.Session
    
    Set rules = Session.DefaultStore.GetRules()
    
    For Each oRule In rules
        oRule.Execute
    Next
    
    
End Sub


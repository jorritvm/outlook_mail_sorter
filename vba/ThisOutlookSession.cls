VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ATTN: the withevent hooks must be top level definitions in the ThisOutlookSession module
' ATTN: the _ItemAdd hooks must also be defined in this module, it does not work if they're elsewhere

' CODE RELATED TO THE MAILSORTER
Dim WithEvents colSentItems As Items
Attribute colSentItems.VB_VarHelpID = -1

' CODE RELATED TO THE ALARM DEACTIVATE
Dim WithEvents colAppItems As Outlook.Items
Attribute colAppItems.VB_VarHelpID = -1


Private Sub Application_Startup()
    ' CODE RELATED TO THE ALARM DEACTIVATE
    Set colAppItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    
    ' CODE RELATED TO THE MAILSORTER
    Set colSentItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderSentMail).Items
        
End Sub

' CODE RELATED TO THE MAILSORTER
Private Sub colSentItems_ItemAdd(ByVal Item As Object)
    Set objItem = Item
    Set objFolder = askUserWhatFolderToPutItemsInto(objItem)
    If TypeName(objFolder) <> "Nothing" Then
        Item.Move objFolder
    End If
    'End If
End Sub


' CODE RELATED TO THE ALARM DEACTIVATE
Private Sub colAppItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    Dim Appt        As Outlook.AppointmentItem
    If TypeOf Item Is Outlook.AppointmentItem And Item.Start < Now() Then
        Set Appt = Item
        Appt.ReminderSet = False
        Appt.Save
    End If
End Sub

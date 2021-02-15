Attribute VB_Name = "mail_sorter"
'-----------------------------------------------
' CHANGELOG
'-----------------------------------------------
'rev2.0: added automated single mail archival, as well as automated batch process of folder
'rev1.2: works with appointments as well
'rev1.1: fixed crash when no correct value given; now shows title in messagebox
'rev1.0: itemadd instead of item_send for compatibility with multiple mailboxes
'rev0.x: works for one mailbox, for manually initiated sort and on-send sort

'----------------
' INIT
'----------------
Sub CountSelectedItems()
    Dim objSelection As Outlook.Selection
    Set objSelection = Application.ActiveExplorer.Selection
    CountSelectedItems = objSelection.Count
End Sub

Sub startManualItemMove()
    Dim objFolder As MAPIFolder
    'vragen naar project
    Set objFolder = askUserWhatFolderToPutItemsInto(Nothing)
    If TypeName(objFolder) <> "Nothing" Then
        Call MoveSelectedMessagesToFolder(objFolder)
    End If
End Sub

Sub startFullyAutomatedItemMoveBasedOnConversation()
    If MsgBox("Are you sure you want to process the entire current folder", vbOKCancel) = vbOK Then
        MoveItemsBasedOnConversation
    End If
End Sub

Sub startOneAutomaticItemMoveBasedOnConversation()
 Dim target_folder As String
 Dim item As Object
 
    If Application.ActiveExplorer.Selection.Count = 1 Then
        Set item = Application.ActiveExplorer.Selection.item(1)
               
        target_folder = get_target_folder_based_on_conversation(item)
        
        If Not Left(target_folder, 4) = "FAIL" Then
            Dim fld As Outlook.Folder
            Set fld = GetMAPIFolderFromStringPath(Right(target_folder, Len(target_folder) - 2))
            item.Move fld
        Else
            MsgBox target_folder
        End If
        
    End If
End Sub

'----------------
' HOOKS
'----------------
Public Sub colSentItems_ItemAdd(ByVal item As Object)
    Set objItem = item
    Set objFolder = askUserWhatFolderToPutItemsInto(objItem)
    If TypeName(objFolder) <> "Nothing" Then
        item.Move objFolder
    End If
    'End If
End Sub

Private Sub colAppItems_ItemAdd(ByVal item As Object)
    On Error Resume Next
    Dim Appt        As Outlook.AppointmentItem
    If TypeOf item Is Outlook.AppointmentItem And item.Start < Now() Then
        Set Appt = item
        Appt.ReminderSet = False
        Appt.Save
    End If
End Sub


'----------------
' HELPERS
'----------------
Public Function abbreviation(text As String) As MAPIFolder
    '***************************************************************************
    'Purpose: to allow the user to define abbreviations for outlook folders
    'Inputs: text abbreviation for a folder
    'Outputs: corresponding MAPI folder object
    '***************************************************************************
    Dim objFolder   As Outlook.MAPIFolder
    Select Case text
        Case "gen"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\" + Format(Date, "yyyy"))
        Case "conf"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\CONFIDENTIAL")
        '...
    End Select
    Set abbreviation = objFolder
End Function


Public Function askUserWhatFolderToPutItemsInto(objItem) As MAPIFolder
    '***************************************************************************
    'Purpose: asks users for a folder abbreviation, if none given, provide user with
    '         a folder structure to select folder from.
    'Inputs: objItem could be passed when dealing with a single item, you could put 'Nothing' when dealing with a collection of items
    'Outputs: a MAPIFolder object
    '***************************************************************************
    Dim objFolder   As MAPIFolder
    Dim objNS       As NameSpace
    Set objNS = Application.GetNamespace("MAPI")
    
    If Not objItem Is Nothing Then
        text = InputBox("Enter map shortcut For mail:" + vbNewLine + objItem.Subject)
    Else:
        text = InputBox("Map shortcut?")
    End If
    If text <> "" Then
        'gebruik van afkorting
        Set objFolder = abbreviation(LCase(CStr(text)))
    End If
    'dan via tree
    If TypeName(objFolder) = "Nothing" Then
        Set objFolder = objNS.PickFolder
    End If
    
    'als het nog niet gekozen is is het een CANCEL
    If TypeName(objFolder) = "Nothing" Then
        'Set objNS = Application.GetNamespace("MAPI")
        'Set objFolder = objNS.GetDefaultFolder(olFolderInbox)
    End If
    Set askUserWhatFolderToPutItemsInto = objFolder
End Function


Public Function GetMAPIFolderFromStringPath(strFolderPath As String) As MAPIFolder
    '***************************************************************************
    'Purpose: translates string path to mapi folder object
    'Inputs
    ' strFolderPath needs to be something like
    '   "Public Folders\All Public Folders\Company\Sales" or
    '   "Personal Folders\Inbox\My Folder"
    'Outputs: MAPI folder object
    '***************************************************************************
  Dim objApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Dim colFolders As Outlook.Folders
  Dim objFolder As Outlook.MAPIFolder
  Dim arrFolders() As String
  Dim i As Long
  On Error Resume Next

  strFolderPath = Replace(strFolderPath, "/", "\")
  arrFolders() = Split(strFolderPath, "\")
  Set objApp = Application
  Set objNS = objApp.GetNamespace("MAPI")
  Set objFolder = objNS.GetDefaultFolder(olFolderInbox).Parent
  Set objFolder = objFolder.Folders.item(arrFolders(0))
  If Not objFolder Is Nothing Then
    For i = 1 To UBound(arrFolders)
      Set colFolders = objFolder.Folders
      Set objFolder = Nothing
      Set objFolder = colFolders.item(arrFolders(i))
      If objFolder Is Nothing Then
        Exit For
      End If
    Next
  End If

  Set GetMAPIFolderFromStringPath = objFolder
  Set colFolders = Nothing
  Set objNS = Nothing
  Set objApp = Nothing
End Function


Sub MoveSelectedMessagesToFolder(Folder As MAPIFolder)
    '***************************************************************************
    'Purpose: will move the selected messages into a MAPI folder
    'Inputs: MAPIFolder object
    'Outputs: nothing
    '***************************************************************************

    Dim objFolder As Outlook.MAPIFolder
    Set objFolder = Folder

    If objFolder Is Nothing Then
        MsgBox "This folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
        Exit Sub
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If

    For Each objItem In Application.ActiveExplorer.Selection
        objItem.UnRead = False
        objItem.Move objFolder
    Next

    Set objItem = Nothing
    Set objFolder = Nothing
End Sub


Sub MoveItemsBasedOnConversation()
    '***************************************************************************
    'Purpose:
    'Inputs
    'Outputs:
    '***************************************************************************

    'declare
    Dim log() As String
    Dim objFolder As Outlook.MAPIFolder
    Dim objItems As Outlook.Items
    Dim currentItem As Object
    Dim i As Integer, j As Integer
    
    'get folder details
    Set objFolder = Application.ActiveExplorer.CurrentFolder
    Set objItems = objFolder.Items
    
    If Not is_conversation_enabled(objFolder) Then
        MsgBox "Conversations Not enabled For this folder. Automatic sorting impossible."
        Exit Sub
    End If
      
    'go through all items
    i = 0
    ReDim log(1 To objItems.Count, 1 To 3)
    
    For j = objItems.Count To 1 Step -1
        Set currentItem = objItems.item(j)
        i = i + 1
        
        'log item generic info
        For Each rec In currentItem.Recipients
            log(i, 1) = log(i, 1) & "-" & rec.Name
        Next rec
        log(i, 2) = currentItem.Subject
        
        Dim target_folder As String
        target_folder = get_target_folder_based_on_conversation(currentItem)
        
        If Left(target_folder, 4) = "FAIL" Then
            log(i, 3) = target_folder
        Else
            Dim fld As Outlook.Folder
            Set fld = GetMAPIFolderFromStringPath(Right(target_folder, Len(target_folder) - 2))
            log(i, 3) = "MOVE: " & fld.Name
            currentItem.Move fld
        End If
    Next j
    
    WriteArrayToImmediateWindow (log)
    
    Set obj = Nothing
    Set objItems = Nothing
    Set objFolder = Nothing
    
End Sub


Function is_conversation_enabled(objFolder As Outlook.MAPIFolder)
    '***************************************************************************
    'Purpose: returns a boolean that indicates whether conversations are enabled for the provided folder
    'Inputs: outlook.mapifolder object
    'Outputs:
    '***************************************************************************

    Dim return_value As Boolean
    Dim objItem  As Outlook.Items
    
    Set objItems = objFolder.Items
    
    return_value = True
    If objItems.Count = 0 Then
        return_value = False
    Else
        return_value = objFolder.store.IsConversationEnabled
        'return_value = objItems(1).Parent.store.IsConversationEnabled
    End If
    
    is_conversation_enabled = return_value
End Function


Function get_target_folder_based_on_conversation(currentItem As Object) As String
    '***************************************************************************
    'Purpose: returns a string representation of the folder of currentItem's parent mail item
    'Inputs: current mail item
    'Outputs: string
    '***************************************************************************

    Dim theMailItem As Outlook.MailItem
    Dim return_value As String
    return_value = ""
    
    ' Check if the item is a MailItem
    If TypeOf currentItem Is Outlook.MailItem Then
        Set theMailItem = currentItem
    Else
        return_value = "FAIL: This item Is Not a mail item."
        GoTo end_function
    End If
    
    ' check if there is a conversation
    Dim theConversation As Outlook.Conversation
    Set theConversation = theMailItem.GetConversation
    If IsNull(theConversation) Then
        return_value = "FAIL: This item Is Not a part of a conversation."
        GoTo end_function
    End If
    
    itemsInThisConversation = theConversation.GetTable.GetRowCount()
    If itemsInThisConversation = 1 Then
        return_value = "FAIL: This item Is Not a part of a conversation."
        GoTo end_function
    End If
    
    ' check uniqueness of root item of the conversation
    Dim group As Outlook.SimpleItems
    Set group = theConversation.GetRootItems
    
    If group.Count > 1 Then
        return_value = "FAIL: This item has multiple root items."
        GoTo end_function
    End If
    
    ' check that the root item itself is a mailitem
    Dim obj As Object
    Set obj = group.item(1)
    If Not TypeOf obj Is Outlook.MailItem Then
        return_value = "FAIL: This item has a root item that Is Not an email."
        GoTo end_function
    End If
    
    ' check that the mailitem is not the root item itself
    Dim mi As Outlook.MailItem
    Set mi = obj
    If theMailItem.ConversationIndex = mi.ConversationIndex Then
        return_value = "FAIL: This item Is the root item."
        GoTo end_function
    End If
    
    ' check that the root item is not inbox or sent
    Dim fld As Outlook.Folder
    Set fld = mi.Parent
    If fld.Name = "Inbox" Then
        return_value = "FAIL: This item has a root item that Is inbox."
        GoTo end_function
    End If
    If fld.Name = "Sent" Then
        return_value = "FAIL: This item has a root item that Is sent."
        GoTo end_function
    End If
    
    ' if we get here it is safe to move the mail item
    return_value = fld.FolderPath
    
end_function:
    get_target_folder_based_on_conversation = return_value
End Function


Sub WriteArrayToImmediateWindow(arrSubA As Variant)
'***************************************************************************
'Purpose: prints 2D array to immediate window, comma separated
'Inputs
'Outputs:
'***************************************************************************
  
    Dim rowString As String
    Dim iSubA As Long
    Dim jSubA As Long
    
    rowString = ""
    
    Debug.Print
    Debug.Print
    Debug.Print "The array is: "
    For iSubA = 1 To UBound(arrSubA, 1)
        rowString = arrSubA(iSubA, 1)
        For jSubA = 2 To UBound(arrSubA, 2)
            rowString = rowString & "," & arrSubA(iSubA, jSubA)
        Next jSubA
        Debug.Print rowString
    Next iSubA
    
End Sub

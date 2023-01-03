Attribute VB_Name = "mail_sorter"
'-----------------------------------------------
' CHANGELOG
'-----------------------------------------------
'rev2.2: fixed module scoping, fixed casing inconsistency (camelCase), reformatted code, new conversation folder finder
'rev2.1: some fixes to better auto sort mails
'rev2.0: added automated single mail archival, as well as automated batch process of folder
'rev1.2: works with appointments as well
'rev1.1: fixed crash when no correct value given; now shows title in messagebox
'rev1.0: itemadd instead of item_send for compatibility with multiple mailboxes
'rev0.x: works for one mailbox, for manually initiated sort and on-send sort


'----------------
' MAIN
'----------------
Public Sub startManualItemMove()
    Dim objFolder   As MAPIFolder
    Set objFolder = askUserWhatFolderToPutItemsInto(Nothing)
    If TypeName(objFolder) <> "Nothing" Then
        Call MoveSelectedMessagesToFolder(objFolder)
    End If
End Sub


Public Sub startOneAutomaticItemMoveBasedOnConversation()
    Dim targetFolder As String
    Dim Item        As Object
    
    If Application.ActiveExplorer.Selection.Count = 1 Then
        Set Item = Application.ActiveExplorer.Selection.Item(1)
        'targetFolder = getTargetFolderBasedOnConversation(item)
        targetFolder = getTargetFolderFromConversation(Item)
        If Not Left(targetFolder, 4) = "FAIL" Then
            Dim fld As Outlook.Folder
            Set fld = GetMAPIFolderFromStringPath(Right(targetFolder, Len(targetFolder) - 2))
            Item.Move fld
        Else
            MsgBox targetFolder
        End If
    End If
End Sub


Public Sub startFullyAutomatedItemMoveBasedOnConversation()
    If MsgBox("Are you sure you want To process the entire current folder", vbOKCancel) = vbOK Then
        MoveItemsBasedOnConversation
    End If
End Sub


'----------------
' HELPERS
'----------------
Private Function abbreviationToMAPIFolder(text As String) As MAPIFolder
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
        Case "opl"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\01. Elia\opleidingen")
        Case "car"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\01. Elia\leasing")
        Case "it"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\01. Elia\IT")
        Case "news"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\01. Elia\news")
        Case "priv"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\z_prive")
        Case "fun"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\z_fun")
            
        'gd
        Case "adq"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\adqflex")
        Case "ct"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\coreteam")
        Case "crm"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\CRM")
        Case "db"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\pisa_opal")
        Case "do"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\dataorg")
        Case "eu"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\entsoe")
        Case "fop"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\_old\FOP")
        Case "gd"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\general")
        Case "vec"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\vectoren")
            
        ' old GD
        Case "adv"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\_old\advisory")
        Case "tf"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\_old\taskforce")
        Case "tir"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\_old\tirole")
        Case "de"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\_old\spoc_DE")
        Case "sr"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\04. GD\_old\SR")
            
        'ipm
        Case "ipm"
            Set objFolder = GetMAPIFolderFromStringPath("Personal Folders\archief\z_old\03. EE\01. IPM")
            
    End Select
    Set abbreviationToMAPIFolder = objFolder
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
        Set objFolder = abbreviationToMAPIFolder(LCase(CStr(text)))
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


Private Function GetMAPIFolderFromStringPath(strFolderPath As String) As MAPIFolder
    '***************************************************************************
    'Purpose: translates string path to mapi folder object
    'Inputs
    ' strFolderPath needs to be something like
    '   "Public Folders\All Public Folders\Company\Sales" or
    '   "Personal Folders\Inbox\My Folder"
    'Outputs: MAPI folder object
    '***************************************************************************
    Dim objApp      As Outlook.Application
    Dim objNS       As Outlook.NameSpace
    Dim colFolders  As Outlook.Folders
    Dim objFolder   As Outlook.MAPIFolder
    Dim arrFolders() As String
    Dim i           As Long
    On Error Resume Next
    
    strFolderPath = Replace(strFolderPath, "/", "\")
    arrFolders() = Split(strFolderPath, "\")
    Set objApp = Application
    Set objNS = objApp.GetNamespace("MAPI")
    Set objFolder = objNS.GetDefaultFolder(olFolderInbox).Parent
    Set objFolder = objFolder.Folders.Item(arrFolders(0))
    If Not objFolder Is Nothing Then
        For i = 1 To UBound(arrFolders)
            Set colFolders = objFolder.Folders
            Set objFolder = Nothing
            Set objFolder = colFolders.Item(arrFolders(i))
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


Private Sub MoveSelectedMessagesToFolder(Folder As MAPIFolder)
    '***************************************************************************
    'Purpose: will move the selected messages into a MAPI folder
    'Inputs: MAPIFolder object
    'Outputs: nothing
    '***************************************************************************
    
    Dim objFolder   As Outlook.MAPIFolder
    Set objFolder = Folder
    
    If objFolder Is Nothing Then
        MsgBox "This folder does not exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
        Exit Sub
    End If
    
    If Application.ActiveExplorer.Selection.Count = 0 Then
        'Require that this procedure be called only when >=1 message is selected
        Exit Sub
    End If
    
    For Each objItem In Application.ActiveExplorer.Selection
        ' if the selected item is already in the desired folder skip it
        If Not objItem.Parent = objFolder Then
            ' mark the item as read
            objItem.UnRead = False
            ' move it
            objItem.Move objFolder
        End If
    Next
    
    Set objItem = Nothing
    Set objFolder = Nothing
End Sub


Private Sub MoveItemsBasedOnConversation()
    '***************************************************************************
    'Purpose: move all messages in the currently selected folder to target
    ' folders based on the location of other items of the same conversation
    'Inputs: -
    'Outputs: -
    '***************************************************************************
    
    'declare
    Dim log()       As String
    Dim objFolder   As Outlook.MAPIFolder
    Dim objItems    As Outlook.Items
    Dim currentitem As Object
    Dim i           As Integer, j As Integer
    
    'get folder details
    Set objFolder = Application.ActiveExplorer.CurrentFolder
    Set objItems = objFolder.Items
    
    If Not isConversationEnabled(objFolder) Then
        MsgBox "Conversations Not enabled For this folder. Automatic sorting impossible."
        Exit Sub
    End If
    
    'go through all items
    i = 0
    ReDim log(1 To objItems.Count, 1 To 3)
    
    ' loop through items from end to start
    For j = objItems.Count To 1 Step -1
        Set currentitem = objItems.Item(j)
        i = i + 1
        
        'log item generic info
        For Each rec In currentitem.Recipients
            log(i, 1) = log(i, 1) & "-" & rec.Name
        Next rec
        log(i, 2) = currentitem.Subject
        
        Dim targetFolder As String
        targetFolder = getTargetFolderFromConversation(currentitem)
        
        If Left(targetFolder, 4) = "FAIL" Then
            log(i, 3) = targetFolder
        Else
            Dim fld As Outlook.Folder
            Set fld = GetMAPIFolderFromStringPath(Right(targetFolder, Len(targetFolder) - 2))
            log(i, 3) = "MOVE: " & fld.Name
            currentitem.Move fld
        End If
    Next j
    
    WriteArrayToImmediateWindow (log)
    
    Set obj = Nothing
    Set objItems = Nothing
    Set objFolder = Nothing
    
End Sub


Private Function isConversationEnabled(objFolder As Outlook.MAPIFolder)
    '***************************************************************************
    'Purpose: returns a boolean that indicates whether conversations are enabled for the provided folder
    'Inputs: outlook.mapifolder object
    'Outputs:
    '***************************************************************************
    
    Dim result As Boolean
    Dim objItem     As Outlook.Items
    
    Set objItems = objFolder.Items
    
    result = True
    If objItems.Count = 0 Then
        result = False
    Else
        result = objFolder.Store.isConversationEnabled
        'result = objItems(1).Parent.store.IsConversationEnabled
    End If
    
    isConversationEnabled = result
End Function


Private Function getTargetFolderFromConversation(currentitem As Object) As String
    '***************************************************************************
    'Purpose: returns a string representation of the folder of currentItem conversation
    'Inputs: current mail item
    'Outputs: string e.g. \\Jorrit.VanderMynsbrugge@elia.be\archief\z_prive
    '***************************************************************************
    Dim thisFolderShort As String
    Dim result As String
    Dim root_items As Outlook.SimpleItems
    Dim child_items As Outlook.SimpleItems
        
    ' set up exclusions
    excludeOriginFolders = Array("Calendar")
    excludeTargetFolders = Array("Inbox", "Sent Items", "Calendar")
    
    ' check if the item is part of a bigger conversation
    Dim theConversation As Outlook.conversation
    Set theConversation = currentitem.GetConversation
    If IsNull(theConversation) Then
        result = "FAIL: This item is not part of a conversation."
    Else
        ' look for the most recent folder used in the conversation that is not excluded
        result = "FAIL - no folder found" ' initial reply
        Set root_items = theConversation.GetRootItems
        For Each root_item In root_items
            ' the root item can already be in the conversation target folder
            If Not IsInArray(root_item.Parent.Name, excludeTargetFolders) Then
                result = root_item.Parent.FolderPath
            End If
                
            ' in addition, per root item we loop through the whole underlying conversation
            Set child_items = theConversation.GetChildren(root_item)
            For Each child_item In child_items
                If Not IsInArray(child_item.Parent.Name, excludeTargetFolders) Then
                    result = child_item.Parent.FolderPath
                End If
            Next child_item
        Next root_item
   End If
endFct:
    getTargetFolderFromConversation = result
End Function


'----------------------------
' UTILS
'----------------------------
Private Function CountSelectedItems() As Integer
    Dim objSelection As Outlook.Selection
    Set objSelection = Application.ActiveExplorer.Selection
    CountSelectedItems = objSelection.Count
End Function


Private Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function


Private Sub printItemType(currentitem)
    Debug.Print ("-------item type----------")
    If TypeOf currentitem Is Outlook.MailItem Then
        Debug.Print ("outlook mail item")
    ElseIf TypeOf currentitem Is Outlook.MeetingItem Then
        Debug.Print ("outlook meeting item")
    ElseIf TypeOf currentitem Is Outlook.AppointmentItem Then
        Debug.Print ("outlook appoint item")
    Else
        Debug.Print ("unknown item type")
        ' e.g. failure to deliver item
    End If
    Debug.Print ("-------------------------")
End Sub


Private Sub WriteArrayToImmediateWindow(arrSubA As Variant)
    '***************************************************************************
    'Purpose: prints 2D array to immediate window, comma separated
    'Inputs
    'Outputs:
    '***************************************************************************
    
    Dim rowString   As String
    Dim iSubA       As Long
    Dim jSubA       As Long
    
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

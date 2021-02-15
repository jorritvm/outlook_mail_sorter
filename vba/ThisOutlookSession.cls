VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'--------------------------------------------
' THE FOLLOWING IS RELATED TO THE MAILSORTER
'--------------------------------------------
Dim WithEvents colAppItems As Outlook.Items
Attribute colAppItems.VB_VarHelpID = -1
Dim WithEvents colSentItems As Items
Attribute colSentItems.VB_VarHelpID = -1


Private Sub Application_Startup()
    ' THE FOLLOWING IS RELATED TO THE MAILSORTER
    Set colAppItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    Set colSentItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderSentMail).Items
    
    'THE FOLLOWING IS RELATED TO THE HTML REPLY MACRO
    'Set oExpl = Application.ActiveExplorer
    'bDiscardEvents = False
    'olFormat = olFormatPlain        '(*1) - reply using plain text
    'olFormat = olFormatHTML        '(*2) - reply using HTML
End Sub



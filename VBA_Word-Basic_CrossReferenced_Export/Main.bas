Attribute VB_Name = "Main"
'***************************************************************************************************
'
'Public Variables and Options
'
'***************************************************************************************************
Option Explicit

'***************************************************************************************************
'Name:      Sub AutoOpen()
'
'Purpose:   Runs automatically when the document is loaded and macros are enabled
'
'Inputs:    None
'
'Returns:   None
'
'***************************************************************************************************
Sub AutoOpen()
    
    'convert all hyperlinks pointing to Jama to hyperlinks to the Word bookmark
    UpdateProgress ("[10%] Converting Hyperlinks...")
    jamaLinkToWord

    'add a bookmark for each ID in the document
    UpdateProgress ("[50%] Adding Bookmarks...")
    createBookmarks
    
    'return the cursor to the top of the document
    UpdateProgress ("[100%] Done")
    Selection.HomeKey Unit:=wdStory
    
End Sub

'***************************************************************************************************
'Name:      Private Function UpdateProgress(strMsg As String)
'
'Purpose:   Displays a progress message on the status bar
'
'Inputs:    strMsg: Message to display
'
'Returns:   None
'
'***************************************************************************************************

Private Function UpdateProgress(strMsg As String)

  Application.StatusBar = strMsg
  Debug.Print strMsg

End Function

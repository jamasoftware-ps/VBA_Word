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

  'resize all images to fit on the page
  UpdateProgress ("[10%] Resizing Images...")
  ResizeImages
  
  'resize all tables to fit on the page
  UpdateProgress ("[30%] Resizing Tables...")
  ResizeTables
  
  'Remove outline level from Item IDs and Item Names
  UpdateProgress ("[50%] Updating Outline Levels...")
  RemoveOutlineLevel ("Item ID")
  RemoveOutlineLevel ("Item Name")
  RemoveOutlineLevel ("Normal")
  
  'ensure all text of style Normal uses the font Arial
  UpdateProgress ("[70%] Correcting Fonts...")
  repairFonts ("Arial")
  
  'remove extra white space from the document
  UpdateProgress ("[80%] Removing Extra White Space...")
  removeExtraWhiteSpace
  
  'Update table of contents
  UpdateProgress ("[90%] Updating Table of Contents...")
  ActiveDocument.TablesOfContents(1).Update
  
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


Attribute VB_Name = "Functions"
'***************************************************************************************************
'
'Public Variables and Options
'
'***************************************************************************************************
Option Explicit

'***************************************************************************************************
'Name:      Function jamaLinkToWord()
'
'Purpose:   convert links in the text from Jama links to Word cross-references
'
'Inputs:    None
'
'Returns:   None
'
'***************************************************************************************************
Function jamaLinkToWord()
  Dim doc As Document
  Dim fld As Field
  Dim fld_code As String
  Dim id As String
  Dim lngStart As Long
  Dim lngAnd As Long
  Dim lngEnd As Long
  
  'Handle any errors
  On Error GoTo ErrHandler:
  
  Set doc = ActiveDocument
  For Each fld In doc.Fields
    fld.Select
    If fld.Type = wdFieldHyperlink And Selection.Style = ActiveDocument.Styles("Normal") Then
        Selection.Style = ActiveDocument.Styles("Hyperlink")
        fld_code = fld.Code
        lngStart = InStr(fld_code, "docId=") + 6
        lngAnd = InStr(fld_code, "&")
        lngEnd = Len(fld_code)
        If lngAnd > lngStart Then
          lngEnd = lngAnd
        End If
        If lngStart > 6 Then
            id = Mid(fld_code, lngStart, lngEnd - lngStart)
            fld.Code.Text = "HYPERLINK \l " & Chr(34) & "API_ID" & id & Chr(34)
            fld.Update
        End If
    End If
NextIteration:

  Next
  
  Set fld = Nothing
  Set doc = Nothing
  
  Exit Function
  
ErrHandler:
  Debug.Print "[Error]: Error #" & Err.Number
  Resume NextIteration:
  
End Function


'***************************************************************************************************
'Name:      Function createBookmarks()
'
'Purpose:   Convert all the hidden item IDs into Word bookmarks
'
'Inputs:    None
'
'Returns:   None
'
'***************************************************************************************************
Function createBookmarks()
    
    Selection.HomeKey Unit:=wdStory
    ActiveWindow.ActivePane.View.ShowAll = True
    Selection.Find.ClearFormatting
    Selection.Find.Font.Hidden = True
    Dim frange As Range
    
    With Selection.Find
        Do While .Execute(FindText:="<(API_ID)*[0-9]>", Forward:=True, _
            MatchWildcards:=True, Wrap:=wdFindStop, MatchCase:=False) = True
            Set frange = Selection.Range
            ActiveDocument.Bookmarks.Add frange.Text, frange
            Selection.Collapse wdCollapseEnd
        Loop
    End With
        ActiveWindow.ActivePane.View.ShowAll = False
End Function


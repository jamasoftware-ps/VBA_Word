Attribute VB_Name = "Jama_Export_PostProcessing"
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
    ResizeImages
    
    'resize all tables to fit on the page
    ResizeTables
    
    'Remove outline level from Item IDs and Item Names
    RemoveOutlineLevel ("Item ID")
    RemoveOutlineLevel ("Item Name")

    'ensure all text of style Normal uses the font Arial
    repairFonts ("Arial")

    'remove extra white space from the document
    removeExtraWhiteSpace
    
    'return the cursor to the top of the document
    Selection.HomeKey Unit:=wdStory
    
End Sub



'***************************************************************************************************
'Name:      Private Function removeExtraWhiteSpace()
'
'Purpose:   Remove unwanted white space from the document
'
'Inputs:    None
'
'Returns:   None
'
'***************************************************************************************************
Private Function removeExtraWhiteSpace()
    '
    'Eliminate Non-Breaking Spaces
    '
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^s^s"
        .Replacement.Text = "^s"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '
    'Eliminate Extra Paragraph Breaks
    '
     Selection.Find.ClearFormatting
     Selection.Find.Replacement.ClearFormatting
     With Selection.Find
         .Text = "^p^p"
         .Replacement.Text = "^p"
         .Forward = True
         .Wrap = wdFindContinue
         .Format = True
         .MatchCase = False
         .MatchWholeWord = False
         .MatchWildcards = False
         .MatchSoundsLike = False
         .MatchAllWordForms = False
    End With
    
    Dim index As Integer
    For index = 1 To 2
        Selection.Find.Execute Replace:=wdReplaceAll
    Next

    '
    'Eliminate Extra Line Breaks
    '
     Selection.Find.ClearFormatting
     Selection.Find.Replacement.ClearFormatting
     With Selection.Find
         .Text = "^l^l"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = True
         .MatchCase = False
         .MatchWholeWord = False
         .MatchWildcards = False
         .MatchSoundsLike = False
         .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Function


'***************************************************************************************************
'Name:      Private Function repairFonts(font_face As String)
'
'Purpose:   convert all text with style Normal to have the font face specified
'
'Inputs:    font_face: provides the font face name to use
'
'Returns:   None
'
'***************************************************************************************************
Private Function repairFonts(font_face As String)
    '
    'Fix Font Face
    '
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Name = font_face
        .Bold = False
        .Italic = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function

'***************************************************************************************************
'Name:      Private Function ResizeImages()
'
'Purpose:   Resize images to be no larger than a page
'
'Inputs:    None
'
'Returns:   None
'
'***************************************************************************************************
Private Function ResizeImages()
  Dim dblPageWidth As Double          'stores the pages width
  Dim dblPageHeight As Double         'stores the page height
  Dim dblAspectRatio As Double        'stores the dblAspectRatio of the current image
  Dim ishape As InlineShape         'stores the current shape
  Dim bolProcessingShape As Boolean

  'Handle any errors
  On Error GoTo ErrHandler:
  
  bolProcessingShape = False
      
  ' determine the current usable page width
  dblPageWidth = ActiveDocument.PageSetup.TextColumns.Width
  

  ' determine the current usable page height
  dblPageHeight = ActiveDocument.PageSetup.PageHeight - ThisDocument.PageSetup.TopMargin - ThisDocument.PageSetup.BottomMargin
  
  bolProcessingShape = True
  
  ' Process each shape in the document
  For Each ishape In ActiveDocument.InlineShapes
      ' determine aspect ratio
      dblAspectRatio = ishape.Height / ishape.Width
      ' if the shape is taller than the page, make it's height match the page
      If (ishape.Height > dblPageHeight) Then
          ishape.Height = dblPageHeight
          ishape.Width = (dblPageHeight / dblAspectRatio)
      End If
      ' if the share is wider than the page, make it's wdith match the page
      If (ishape.Width > dblPageWidth) Then
          ishape.Width = dblPageWidth
          ishape.Height = (dblPageWidth * dblAspectRatio)
      End If
NextIteration:
  Next
  Exit Function
  
ErrHandler:
  Debug.Print "[Error]: Error #" & Err.Number
  If bolProcessingShape Then
    Resume NextIteration:
  Else
    Exit Function
  End If
  
End Function

'***************************************************************************************************
'Name:      Private Function RemoveOutlineLevel(strStyle as string)
'
'Purpose:   Resize images to be no larger than a page
'
'Inputs:    strStyle: style to remove outline level from
'
'Returns:   None
'
'***************************************************************************************************
Private Function RemoveOutlineLevel(strStyle As String)
  Dim intParagraphIndex As Integer  'index used for looping through all paragraphs in the document
  Dim intParagraphCount As Integer  'total number of paragraphs in the document at start of function
  
  intParagraphCount = ActiveDocument.Paragraphs.Count
  
  For intParagraphIndex = intParagraphCount To 1 Step -1

    'Handle any errors
    On Error GoTo ErrHandler:
    
    ActiveDocument.Paragraphs(intParagraphIndex).Range.Select
    
    If Selection.Style = ActiveDocument.Styles(strStyle) Then
      Selection.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    End If
                
NextIteration:
    If intParagraphIndex < 2 Then
      Exit Function
    End If
  
  Next intParagraphIndex
  
  Exit Function
  
ErrHandler:
  Debug.Print intParagraphIndex & " [Error]: Error #" & Err.Number & " occured while processing '" & Selection.Range.Text & "'"
  Resume NextIteration:
  
End Function


'***************************************************************************************************
'Name:      Private Function ResizeTables()
'
'Purpose:   Resize tables to be no larger than a page
'
'Inputs:    None
'
'Returns:   None
'
'***************************************************************************************************
Private Function ResizeTables()
Dim oTbl As Table

  For Each oTbl In ActiveDocument.Tables
    If oTbl.PreferredWidthType = wdPreferredWidthPoints And oTbl.PreferredWidth > ActiveDocument.PageSetup.TextColumns.Width Then
      oTbl.AutoFitBehavior wdAutoFitFixed
      oTbl.PreferredWidthType = wdPreferredWidthPercent
      oTbl.PreferredWidth = "100"
    End If
  Next oTbl
    
End Function



Attribute VB_Name = "modTools"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   MICROSOFT WORD TOOLS
' PURPOSE:  Different tools to draw information from ranges or perform other routine Word
'           operations.
' METHODS:  DisplayNumberOnStatusBarIfMultiple
'           UpdateTableOfContents
'           PasteFromClipboardWithNoSpaces
'           MsgSelectionFontRGBColor
'           GetRangeLineNumber
'           MsgCharactersInSelectionText
'           MsgCurrentParagraphIndex
'           XmlContent
'           TypeXmlOfSelection
'           StartNewParagraphWithEachCapitalizedWord
'           DeleteAllFootnotes
'           TurnScreenUpdatingOFF
'           TurnScreenUpdatingON
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Sub DisplayNumberOnStatusBarIfMultiple( _
   ByVal number As Integer, _
   ByVal divisor As Integer, _
   Optional ByVal precedingMsg As String = "", _
   Optional ByVal trailingMsg As String = "")
'****************************************
' Displays a number on the StatusBar if it's evenly divisible by the divisor, along with any
' preceding and trailing messages provided by the user.
' A method to use if you don't want to continually update the StatusBar with every update of
' underlying process but only want to do so in increments.
'****************************************

   If (number Mod divisor) = 0 Then
      Dim msg As String
      msg = number
      If Len(precedingMsg) > 0 Then msg = precedingMsg & msg
      If Len(trailingMsg) > 0 Then msg = msg & trailingMsg
      Application.StatusBar = msg
   End If

End Sub

Public Sub Tools_UpdateTableOfContents()
'********************************************************
' In most documents the TablesOfContents() collection will consist of only a single item.
'********************************************************

   On Error Resume Next
   ActiveDocument.TablesOfContents(1).Update
   ' 5941 error results from not having a Table of Contents.
   If Err.number = 5941 Or Err.number = 0 Then
      Exit Sub
   Else:
      MsgBox "Error encountered updating Table of Contents.", , "Error"
   End If

End Sub

Public Sub Tools_PasteFromClipboardWithNoSpaces()
'********************************************************
' Paste contents of the clipboard into currently selected area with no line or space breaks.
' Retrieves text from the clipboard, removes newline characters, and outputs it into the document.
'********************************************************

   Dim nicelyFormattedText As String
   nicelyFormattedText = Txt_GetFromClipboard
   nicelyFormattedText = Txt_RemoveNewLineChars(nicelyFormattedText)
   Selection.TypeText (nicelyFormattedText)
      
End Sub

Public Sub Tools_MsgSelectionFontRGBColor()
'****************************************
' Msgbox the RGB color for the Selection's font.
'****************************************

   Dim hexColor As String
   Dim rgbColor As String
   
   hexColor = Right("000000" & Hex(Selection.Font.color), 6)
   
   rgbColor = "RGB (" & CInt("&H" & Right(hexColor, 2)) & _
   ", " & CInt("&H" & Mid(hexColor, 3, 2)) & _
   ", " & CInt("&H" & Left(hexColor, 2)) & ")"
   
   MsgBox rgbColor

End Sub

Public Function Tools_GetRangeLineNumber(ByVal myRange As Range) As Integer
'********************************************************
' Returns the line number of the first character for a given range.
'********************************************************

   Tools_GetRangeLineNumber = myRange.Information(g_WORDENUM_INFO_FIRSTCHARACTERLINENUMBER)

End Function

Public Sub Tools_MsgCharactersInSelectionText()

   Dim s As String
   Dim i As Integer
   s = Selection.Text
   For i = 1 To Len(s)
      Dim g As Variant
      g = Asc(Mid$(s, i, 1))
      MsgBox g
   Next i

End Sub

Public Sub Tools_MsgCurrentParagraphIndex()
'********************************************************
' Displays MsgBox with the index number of the current paragraph.
'********************************************************

   MsgBox ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count

End Sub

Public Function Tools_XmlContent(ByVal w As Range) As String
'********************************************************
' Retrieve XML representation of the range.
'********************************************************

   Tools_XmlContent = w.XML

End Function

Public Sub Tools_TypeXmlOfSelection()

   Dim xmlTxt As String: xmlTxt = Tools_XmlContent(Selection.Range)
   Selection.TypeText xmlTxt

End Sub

Public Sub Tools_StartNewParagraphWithEachCapitalizedWord()
'********************************************************
' Pushes each word encountered with a capital letter onto a new line.
'********************************************************

   Dim w As Range
   For Each w In Selection.words
      If Txt_FirstLetterIsCapitalized(w.Text) Then
         w.InsertBefore (vbCr)
      End If
   Next w

End Sub

Public Sub Tools_DeleteAllFootnotes()

   On Error GoTo ERROR_MANAGER
   
   Dim fn As Footnote, countDeleted As Integer
   countDeleted = 0
   For Each fn In ActiveDocument.Footnotes
      fn.Delete
      countDeleted = countDeleted + 1
   Next fn
   
   MsgBox "Deleted " + countDeleted + " footnotes.", , "Done"
   
   Exit Sub
   
ERROR_MANAGER:
   Select Case Err.number
      Case 6243   'Error when attempting to format references in fuctions
         Resume Next
      Case Else
         MsgBox "Error Number: " & Err.number & vbLf & _
            "Error Description: " & Err.Description, , "Error in Fmt_Footnotes"
   End Select

End Sub

Public Sub TurnScreenUpdatingOFF()

   Application.ScreenUpdating = False

End Sub

Public Sub TurnScreenUpdatingON()

   Application.ScreenUpdating = True

End Sub

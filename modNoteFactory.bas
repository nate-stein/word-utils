Attribute VB_Name = "modNoteFactory"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   NOTE FACTORY
' PURPOSE:  Allow for formatting of text containing notes.
' METHODS:  CreateEditWrapper
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Function NoteFactory_CreateEditWrapper() As NotesEditWrapper

   Dim wrapper As New NotesEditWrapper
   wrapper.WordsToItalicize = getArrayOfWordsToItalicize()
   wrapper.WordsToSubscriptLastLetter = getArrayOfWordsToSubscriptLastLetter()
   wrapper.WordsToSuperscriptLastLetter = getArrayOfWordsToSuperscriptLastLetter()
   wrapper.WordsToBold = getArrayOfWordsToBold()
   
   Set NoteFactory_CreateEditWrapper = wrapper
   Set wrapper = Nothing

End Function

Private Function getArrayOfWordsToItalicize() As Variant
'****************************************
' Returns an array of words that should be italicized based on the current ActiveDocument.
' To help keep this maintainable, the items within each array should be listed alphabetically.
'****************************************

   Select Case ActiveDocument.Name
      Case "Book Notes"
         getArrayOfWordsToItalicize = Array("A", "k", "L", "n", "p", "P", "q", "S", "S1", "S2", "t", _
            "W", "X", "Y", "YS")
         ' Old list in case we need it later:
         ' "B", "b", "c", "comp", "Cos", "perpcomp", "perpframe", "point", "r", "s", "Sin", "t", "U", "u", "V", "v", "X", "x", "Y", "y", "Z", "z"
      Case Else
         Dim errMsg As String
         errMsg = "No known list of words to italicize for current ActiveDocument."
         Err.Raise g_ERROR_NOTES_FACTORY, "getArrayOfWordsToItalicize()", errMsg
   End Select

End Function

Private Function getArrayOfWordsToSubscriptLastLetter() As Variant
'****************************************
' Return an array of words for which the last letter should be subscripted.
'****************************************

   Select Case ActiveDocument.Name
      Case "Notes - Part II.docx"
         getArrayOfWordsToSubscriptLastLetter = Array()
      Case Else
         Dim errMsg As String
         errMsg = "No known list of words to subscript for current ActiveDocument."
         Call Err.Raise(g_ERROR_NOTES_FACTORY, "getArrayOfWordsToItalicize()", errMsg)
   End Select

End Function

Private Function getArrayOfWordsToSuperscriptLastLetter() As Variant
'****************************************
' Returns an array of words for which the last letter should be superscripted.
'****************************************

   Select Case ActiveDocument.Name
      Case "Notes - Part II.docx"
         getArrayOfWordsToSuperscriptLastLetter = Array("uT", "vT")
      Case Else
         Dim errMsg As String
         errMsg = "No known list of words to superscript for current ActiveDocument."
         Call Err.Raise(g_ERROR_NOTES_FACTORY, "getArrayOfWordsToItalicize()", errMsg)
   End Select

End Function

Private Function getArrayOfWordsToBold() As Variant

   Select Case ActiveDocument.Name
      Case "Notes - Part II.docx"
         getArrayOfWordsToBold = Array("u", "v", "w", "x", "y", "z")
      Case Else
         Dim errMsg As String
         errMsg = "No known list of words to bold for current ActiveDocument."
         Call Err.Raise(g_ERROR_NOTES_FACTORY, "getArrayOfWordsToBold()", errMsg)
   End Select

End Function


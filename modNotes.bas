Attribute VB_Name = "modNotes"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   NOTES
' SUBS:     FormatSelection
'           PasteClipboardWithNoSpaces
'           PasteClipboardWithSpaces
' FUNCS:    CreateClipboardPaste
'           GetKindleReferencesArray

Private Const m_EMPTY_CHAR_ASC_CODE As Integer = 13
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Sub Notes_FormatSelection()
'********************************************************
' Formats the text in the Selection according to rules in the NotesEditWrapper.
'********************************************************
   
   On Error GoTo NOTIFY_USER
   
   Dim notesWrapper As NotesEditWrapper
   Set notesWrapper = NoteFactory_CreateEditWrapper()
   notesWrapper.Edit
   Exit Sub
   
NOTIFY_USER:
   MsgBox "Description: " & Err.Description & vbLf & _
      "Source: " & Err.Source, Title:="Error"
   
End Sub

Public Sub Notes_PasteClipboardWithNoSpaces()
'********************************************************
' Pastes contents of the clipboard into currently selected area with no line or space breaks and
' reformatted where necessary according to the ClipboardPaste object model.
'********************************************************

   If ActiveDocument.Name = "R Notes.docx" Then
      Notes_PasteRCode
      Exit Sub
   End If

   Dim clip As ClipBoardPaste
   Set clip = CreateClipboardPaste(True, True)
   Selection.TypeText (clip.FormattedText)
   Set clip = Nothing
   
   If selectionToTheLeftIsEmpty() Then
      Selection.TypeBackspace
   ElseIf Not oneNonEmptyCharacterIsSelected() Then
      moveRightOneCharacter
   End If
      
End Sub

Public Sub Notes_PasteClipboardWithSpaces()

   Dim clipAndPaste As ClipBoardPaste
   Set clipAndPaste = CreateClipboardPaste(False, True)
   Selection.TypeText (clipAndPaste.FormattedText)
   If selectionToTheLeftIsEmpty() Then Selection.TypeBackspace
   Set clipAndPaste = Nothing

End Sub

Public Sub Notes_PasteRCode()
'********************************************************
' Formats R code that we have copied from another source and are pasting into a Word document.
'********************************************************

   Dim clip As ClipBoardPaste, txt As String
   Set clip = CreateClipboardPaste(True, True)
   txt = clip.FormattedText
   
   '''''''''''''''''''''''''''''''''''''''
   ' Replace all console entry symbols (<) and variants (i.e. combinations with a space) with a
   ' newline character.
   '''''''''''''''''''''''''''''''''''''''
   txt = Replace(txt, " > ", Chr(10))
   txt = Replace(txt, "> ", Chr(10))
   txt = Replace(txt, " >", Chr(10))
   txt = Txt_ReplaceWordStyleApostrophes(txt)
   
   ' Make sure we don't start the paste with a new line.
   If Left(txt, 1) = Chr(10) Then
      txt = Right(txt, Len(txt) - 1)
   End If
   
   Selection.TypeText txt
   If selectionToTheLeftIsEmpty() Then Selection.TypeBackspace
   Set clip = Nothing
   
End Sub

Private Function oneNonEmptyCharacterIsSelected() As Boolean

   If Selection.Characters.Count = 1 And Asc(Selection.Characters(1)) = m_EMPTY_CHAR_ASC_CODE Then
      oneNonEmptyCharacterIsSelected = True
   Else: oneNonEmptyCharacterIsSelected = False
   End If

End Function

Private Function selectionToTheLeftIsEmpty() As Boolean

   Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
   If Selection.Text = " " Then
      selectionToTheLeftIsEmpty = True
   Else: selectionToTheLeftIsEmpty = False
   End If

End Function

Public Function CreateClipboardPaste( _
   ByVal RemoveNewLines As Boolean, _
   ByVal PerformCodingLanguageFormatting As Boolean, _
   Optional ByVal RemoveExtraSpacing As Boolean = True) As ClipBoardPaste
'********************************************************
' Factory method for ClipBoardPaste class.
'********************************************************

   Dim obj As New ClipBoardPaste
   obj.Text = Txt_GetFromClipboard()
   If RemoveNewLines Then obj.RemoveNewLines = True
   If PerformCodingLanguageFormatting Then obj.PerformCodingLanguageEdits = True
   
   obj.Edit
   
   Set CreateClipboardPaste = obj
   Set obj = Nothing

End Function

Private Sub moveRightOneCharacter()
'********************************************************
' Equivalent to pressing the right-arrow button once.
'********************************************************

   Selection.MoveRight Unit:=wdCharacter, Count:=1

End Sub

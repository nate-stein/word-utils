Attribute VB_Name = "modTxt"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   TEXT ANALYSIS
' PURPOSE:  Different functions to alter text or draw information about its contents.
'           These methods should be independent of the Word object model and concerned with pure
'           text analysis.
' METHODS:  ConvertASCIICodesToString
'           ConvertSelectionToStandardCase
'           ConvertToCharArray
'           FirstLetterIsCapitalized
'           GetFromClipboard
'           IsALetter
'           IsUCaseLetter
'           LastCharIndex
'           LastCharIsSpace
'           MsgActiveCharASCII
'           RemoveNewLineChars
'           Txt_RemoveMultipleSpaces
'           ReplaceWordStyleApostrophes
'           ReplaceMinusWithHyphen
'           ReplaceShortHyphensWithBar
'           UseGenericApostrophesInSelection
' FAVORITE ASCII
'           63: Long arrow
'           133: Ellipse (...)
'           149: bullet point
'           215: matrix dimensions (e.g. 2 x 5)
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Function Txt_ConvertASCIICodesToString(ByVal codes As Variant) As String
'********************************************************
' Returns string formed by concatenating all characters in codes, an array of ASCII codes.
'********************************************************
   
   '''''''''''''''''''''''''''''''''''''''
   ' Handle cases where codes is not an array or only contains one element.
   '''''''''''''''''''''''''''''''''''''''
   If Not IsArray(codes) Then
      Txt_ConvertASCIICodesToString = Chr(codes)
      Exit Function
   ElseIf UBound(codes) = 0 Then
      Txt_ConvertASCIICodesToString = Chr(codes(0))
      Exit Function
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' Getting here means codes is an array containing > 1 element.
   '''''''''''''''''''''''''''''''''''''''
   Dim result As String
   result = Chr(codes(0))
   
   Dim q As Integer
   For q = 1 To UBound(codes) Step 1
      result = result & Chr(codes(q))
   Next q
      
   Txt_ConvertASCIICodesToString = result

End Function

Public Sub Txt_ConvertSelectionToStandardCase()
'********************************************************
' Converts text like "UPPERCASE EXAMPLE" to "Uppercase Example" where only the first letters of
' words are capitalized.
'********************************************************

   Dim updatedTxt As String
   updatedTxt = convertToStandardCase(Selection.Text)
   
   Selection.TypeText (updatedTxt)
   Selection.TypeBackspace ' Return to original positioning

End Sub

Private Function convertToStandardCase(ByVal txt As String) As String
'*********************************************************
' Returns txt formatted in such a way that only the first letter of its word(s) is capitalized.
' For example, convertToStandardCase("MY FRIEND") = "My Friend"
'*********************************************************

   Dim words As Variant
   words = Split(LCase(txt))
   
   ' We always capitalize the first word in the txt.
   words(0) = capitalizeFirstLetterOfWord(words(0))
   Dim q As Integer
   For q = 1 To UBound(words)
      If Not wordShouldBeLCase(words(q)) Then
         words(q) = capitalizeFirstLetterOfWord(words(q))
      End If
   Next q
   
   convertToStandardCase = concatenateWordsIntoExpression(words)

End Function

Private Function concatenateWordsIntoExpression(ByRef words As Variant) As String

   Dim result As String
   
   Dim q As Integer
   For q = 0 To UBound(words) Step 1
      If q = 0 Then
         result = words(q)
      Else: result = result + " " + words(q)
      End If
   Next q
   
   concatenateWordsIntoExpression = result

End Function

Private Function capitalizeFirstLetterOfWord(ByVal w As String) As String

   Dim firstLetter As String, remainderOfWord As String
   firstLetter = Left(w, 1)
   remainderOfWord = Mid(w, 2, Len(w) - 1)
   capitalizeFirstLetterOfWord = UCase(firstLetter) + remainderOfWord

End Function

Private Function wordShouldBeLCase(ByVal w As String) As Boolean

   w = LCase(w)
   
   If firstCharIsALetter(w) Then
      Dim lcaseWords() As Variant
      lcaseWords = getWordsThatShouldBeLCase()
      wordShouldBeLCase = Array_ContainsValue(lcaseWords, w)
   Else: wordShouldBeLCase = True
   End If

End Function

Private Function getWordsThatShouldBeLCase() As Variant
'********************************************************
' Returns array of words that should generally be lowercase.
'********************************************************

   getWordsThatShouldBeLCase = Array("a", "an", "and", "of", "the")

End Function

Private Function firstCharIsALetter(ByVal w As String) As Boolean
'********************************************************
' Returns True if the first character in word is a letter.
'********************************************************

   Dim firstLetter As String: firstLetter = Left(w, 1)
   firstCharIsALetter = Txt_IsALetter(firstLetter)

End Function

Public Function Txt_ConvertToASCIICodes(ByVal txt As String) As Variant
'********************************************************
' Returns array of ASCII codes for characters in txt.
'********************************************************
   
   Dim chrs As Variant
   chrs = Txt_ConvertToChars(txt)
   
   Dim result() As Variant
   ReDim result(0 To UBound(chrs))
   
   Dim i As Integer
   For i = 0 To UBound(chrs)
      result(i) = Asc(chrs(i))
   Next i
   
   Txt_ConvertToASCIICodes = result

End Function

Public Function Txt_ConvertToChars(ByVal txt As String) As Variant
'********************************************************
' Returns array of chars that make up txt.
'********************************************************

    Dim buff() As String
    ReDim buff(Len(txt) - 1)
    Dim i As Integer
    For i = 1 To Len(txt)
        buff(i - 1) = Mid$(txt, i, 1)
    Next
    Txt_ConvertToChars = buff

End Function

Public Function Txt_ConvertSpecialPunctuation(ByVal txt As String) As String
'********************************************************
' Converts those ASCII characters representing punctuation marks that are location-agnostic but
' for which Word has location-specific replacements.
' E.g. an apostrophe in VBA is neutral (') while in Word it hangs to the left. Moreoever, in VBA
' quotation marks (") are side-neutral but in Word they hang to the left or right depending on
' which side of the quote they are on.
'********************************************************

   Dim codes As Variant
   codes = Txt_ConvertToASCIICodes(txt)
   
   Dim q As Integer
   For q = 0 To UBound(codes)
      Select Case codes(q)
         Case 34 ' Quotation marks
            ' We assume it's the beginning of quotes if we are at the start of txt or if the
            ' character immediately to the left is not alphanumeric.
            If q = 0 Then
               codes(q) = 147
            ' UNLESS it is at the end of the text, in which case we assume they are closing quotes.
            ElseIf q = UBound(codes) Then
               codes(q) = 148
            Else:
               Select Case codes(q - 1)
                  Case 48 - 57, 65 To 90, 97 To 122 ' alphanumeric
                     codes(q) = 148
                  Case Else
                     codes(q) = 147
               End Select
            End If
         Case 39 ' Apostrophe
            codes(q) = 146
      End Select
   Next q
   
   Txt_ConvertSpecialPunctuation = Txt_ConvertASCIICodesToString(codes)

End Function

Public Function Txt_FirstLetterIsCapitalized(ByVal w As String) As Boolean
'********************************************************
' Returns True if the first character in w is a capital letter.
'********************************************************

   Dim firstLetter As String: firstLetter = Left(Trim(w), 1)
   Txt_FirstLetterIsCapitalized = Txt_IsUCaseLetter(firstLetter)

End Function

Public Function Txt_GetFromClipboard() As String
'********************************************************
' Grabs the text that is currently in the Clipboard.
' Uses late-binding to create the DataObject; else, for early binding we'd use:
' Set MyData = New MSFORMS.DataObject
'********************************************************

   Dim clipboardObj As Object
   Set clipboardObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
   clipboardObj.GetFromClipboard
   Txt_GetFromClipboard = clipboardObj.GetText
   Set clipboardObj = Nothing
      
End Function

Public Function Txt_IsALetter(ByVal c As String) As Boolean
'*********************************************************
' Returns True if c is a letter (whether uppercase or lowercase).
'*********************************************************

   Dim ascIICode As Integer
   ascIICode = Asc(c)
   
   Select Case ascIICode
      Case 65 To 90, 97 To 122
         Txt_IsALetter = True
      Case Else
         Txt_IsALetter = False
   End Select

End Function

Public Function Txt_IsUCaseLetter(ByVal c As String) As Boolean

   Dim ascIICode As Integer
   ascIICode = Asc(c)
   
   Select Case ascIICode
      Case 65 To 90 'ASCII Code for A to Z
         Txt_IsUCaseLetter = True
      Case Else
         Txt_IsUCaseLetter = False
   End Select

End Function

Public Function Txt_LastCharIndex(ByVal w As Range) As Integer
'********************************************************
' If the last character in our word is a space or a punctuation mark, it returns the index of the
' 2nd-to-last character.
'********************************************************

   ' Initialize to number of characters.
   Dim result As Integer
   result = w.Characters.Count
   
   ' Check whether characters at the end are are special characters we don't care about.
   If Txt_LastCharIsSpace(w) Or lastCharIsPunctuationMark(w) Then
      result = result - 1
   End If
   
   Txt_LastCharIndex = result
   
End Function

Public Function Txt_LastCharIsSpace(ByVal w As Range) As Boolean
'********************************************************
' Returns True if the last character in w is an empty space.
'********************************************************

   Dim lastChar As String
   lastChar = getLastChar(w)
   Txt_LastCharIsSpace = (lastChar = " ")

End Function

Private Function getLastChar(ByVal w As Range) As String
'********************************************************
' Returns the last character in given word range's text.
'********************************************************
   
   Dim charCount As Integer
   charCount = w.Characters.Count
   getLastChar = w.Characters(charCount)

End Function

Private Function lastCharIsPunctuationMark(ByVal w As Range) As Boolean
'********************************************************
' Returns True if the last character in word is a punctuation mark.
'********************************************************

   Dim punctMarks As Variant
   punctMarks = Array(".", "!", "?")
   
   Dim lastChar As String
   lastChar = getLastChar(w)
   lastCharIsPunctuationMark = Array_ContainsValue(punctMarks, lastChar)

End Function

Public Sub Txt_MsgActiveCharASCII()
'********************************************************
' Displays MsgBox with the ASCII code of the selected character.
'********************************************************

   Dim firstChar As String
   firstChar = Selection.Characters(1).Text
   MsgBox "ASCII Code = " & Asc(firstChar)

End Sub

Public Function Txt_RemoveNewLineChars(ByVal txt As String) As String
'********************************************************
' Remove new line characters from a String that could come from various sources.
'********************************************************

   Dim result As String
   result = txt
   result = Replace(result, Chr(160), " ")
   result = Replace(result, Chr(13), " ")
   result = Replace(result, Chr(11), " ")
   result = Replace(result, vbCr, " ")
   result = Replace(result, vbLf, " ")
   result = Replace(result, vbCrLf, " ")
   result = Replace(result, "      ", " ")
   result = Replace(result, "     ", " ")
   result = Replace(result, "    ", " ")
   result = Replace(result, "   ", " ")
   result = Replace(result, "  ", " ")
   
   Txt_RemoveNewLineChars = result
      
End Function

Public Function Txt_RemoveMultipleSpaces(ByVal txt As String) As String
'********************************************************
' Convert all instances of multiple spaces to a single space.
' Example: Txt_RemoveMultipleSpaces("The  cat jumped   over") -> "The cat jumped over"
'********************************************************

   Dim result As String
   result = txt
   result = Replace(result, "    ", " ")
   result = Replace(result, "   ", " ")
   result = Replace(result, "  ", " ")
   
   Txt_RemoveMultipleSpaces = result
      
End Function

Public Function Txt_ReplaceWordStyleApostrophes(ByVal txt As String) As String
'********************************************************
' Replaces Word slanted punctuation with generic punctuation:
'     Slanted apostrophes (ASCII 145 & 146) -> Generic apostrophe (39)
'     Slanted quotation marks (ASCII 147 & 148) -> Generic quotation marks (34).
' AND afterwards removes unwanted apostrophes from scenarios like the following:
'     (' file.txt')
'     ('file.txt ')
'********************************************************

   Dim codes As Variant
   codes = Txt_ConvertToASCIICodes(txt)
   
   Dim codeCount As Integer
   codeCount = UBound(codes)
   
   Dim q As Integer
   For q = 0 To codeCount
      If codes(q) = 145 Or codes(q) = 146 Then
         codes(q) = 39
      ElseIf codes(q) = 147 Or codes(q) = 148 Then
         codes(q) = 34
      End If
   Next q
   
   '''''''''''''''''''''''''''''''''''''''
   ' Remove unwanted apostrophes around parantheses.
   '''''''''''''''''''''''''''''''''''''''
   q = 0
   Do
      If codes(q) <> 39 Then
         GoTo NEXT_Q
      End If
            
      ' Both scenarios we check for require that we check the character to the left and right of
      ' current index (i.e. we can't be at the first or last character index).
      If q = 0 Or q = codeCount Then
         GoTo NEXT_Q
      End If
      
      ' Check if both:
      '  *  char to the left = ' '
      '  *  char to the right = ')'.
      ' If so, then delete space on the left.
      If codes(q - 1) = 32 And codes(q + 1) = 41 Then
         Arr_DeleteElement q - 1, codes
         codeCount = codeCount - 1
         GoTo NEXT_Q
      End If
      
      ' Check if both:
      '  *  char to the left = '('
      '  *  char to the right = ' '.
      ' If so, then delete space on right left.
      If codes(q - 1) = 40 And codes(q + 1) = 32 Then
         Arr_DeleteElement q + 1, codes
         codeCount = codeCount - 1
         GoTo NEXT_Q
      End If
      
NEXT_Q:
      q = q + 1
   Loop While q < codeCount
   
   Txt_ReplaceWordStyleApostrophes = Txt_ConvertASCIICodesToString(codes)
   
End Function

Public Function Txt_ReplaceMinusWithHyphen(ByVal txt As String) As String
'********************************************************
' Replaces each instance of minus sign (ASCII 45) w/ a hyphen (ASCII 150).
'********************************************************

   Dim codes As Variant
   codes = Txt_ConvertToASCIICodes(txt)

   Dim maxQ As Integer
   maxQ = UBound(codes)
   
   Dim q As Integer
   q = 0
   Do
      If codes(q) = 45 Then codes(q) = 150
      q = q + 1
   Loop While q <= maxQ
   
   Txt_ReplaceMinusWithHyphen = Txt_ConvertASCIICodesToString(codes)

End Function

Public Function Txt_ReplaceShortHyphensWithBar(ByVal txt As String) As String
'********************************************************
' Replaces each instance of short hyphen (ASCII 150) w/ a long hyphen (ASCII 151)
'********************************************************

   Dim codes As Variant
   codes = Txt_ConvertToASCIICodes(txt)

   Dim maxQ As Integer
   maxQ = UBound(codes)
   
   Dim q As Integer
   q = 0
   Do
      If codes(q) <> 150 And codes(q) <> 151 Then
         GoTo NEXT_Q
      End If
      
      ' Remove space character to the right (if there is one).
      If q < maxQ Then
         If codes(q + 1) = 32 Then
            Arr_DeleteElement q + 1, codes
            maxQ = maxQ - 1
         End If
      End If
      
      ' Replace short hyphen w/ long one.
      If codes(q) = 150 Then codes(q) = 151
      
      ' Remove space character to the right (if there is one).
      If q > 0 Then
         If codes(q - 1) = 32 Then
            Arr_DeleteElement q - 1, codes
            maxQ = maxQ - 1
         End If
      End If
NEXT_Q:
      q = q + 1
   Loop While q <= maxQ
   
   Txt_ReplaceShortHyphensWithBar = Txt_ConvertASCIICodesToString(codes)

End Function

Public Sub Txt_ReplaceMinusWithHyphenInSelection()

   Selection.Text = Txt_ReplaceMinusWithHyphen(Selection.Text)
   
End Sub

Public Sub Txt_ReplaceShortHyphensWithBarInSelection()

   Selection.Text = Txt_ReplaceShortHyphensWithBar(Selection.Text)
   
End Sub

Public Sub Txt_UseGenericApostrophesInSelection()

   Selection.Text = Txt_ReplaceWordStyleApostrophes(Selection.Text)

End Sub

Public Sub Txt_RemoveMultipleSpacesInSelection()

   Selection.Text = Txt_RemoveMultipleSpaces(Selection.Text)

End Sub

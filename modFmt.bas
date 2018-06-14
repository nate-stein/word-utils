Attribute VB_Name = "modFmt"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   FORMATTING UTILS
' METHODS:  PerformAll
'           AddCustomFontStylesToDocument
'           ChangeFontSizeIfSpecificName
'           ChangeWordColorsInSelection
'           ConvertOneStyleToAnother
'           SelectionToProgrammingStyle
'           CertainWordsToProgrammingStyle
'           BoldCertainWords
'           ItalicizeCertainWords
'           UpdateHeadingStyles
'           RemoveAllUserDefinedStyles
'           WordColorsInDocument
'           FontNamesInDocument
'           Footnotes
'           AllEquations
'           PageSetup
'           SelectedPictureDimensions
'           TableNumberingIndentation
'           ParagraphSpacingToDefaults
'           FontSuperscript
'           ResizeImage
'           SelectionSuperscript
'           SubscriptLastCharOfMathExpressions
'           FontSubscript
'           SelectionSubscript
'           SuperscriptLastLetter
'           SubscriptLastLetter
'           BoldWord
'           ItalicizeWord
'           ChangeFontColorForAllWords
'           MakeGreenWordsBlue

' Format constants
Private Const m_PARAGRAPH_SPACEAFTER_DEFAULT = 4
Private Const m_PARAGRAPH_SPACEBEFORE_DEFAULT = 0

' Image resizing proportions
Private Const m_RESIZE_PROPORTION_DEFAULT As Double = 75
Private Const m_RESIZE_PROPORTION_THEODORIDIS As Double = 80
Private Const m_RESIZE_PROPORTION As Double = 60

' Font sizes
Private Const m_FONTSIZE_FOOTNOTE_RANGE As Double = 8       ' The actual footnote area at bottom of page
Private Const m_FONTSIZE_FOOTNOTE_REFERENCE As Double = 9   ' The superscripted number indicating a footnote
Private Const m_FONTSIZE_PROGRAMMING_BLUE  As Double = 9
Private Const m_FONTSIZE_PROGRAMMING_CLASS  As Double = 9
Private Const m_FONTSIZE_DEFAULT  As Double = 11
Private Const m_FONTSIZE_EQUATIONS_DEFAULT As Double = 10
Private Const m_FONTSIZE_EQUATIONS As Double = m_FONTSIZE_EQUATIONS_DEFAULT

' Font characteristics
Private Const m_ITALICIZE_EQUATION_FONT As Boolean = False

' Text color indices
Private Const m_FONT_TEXTCOLOR_HEADERBLUE As Long = -738148353

' Font names for specific areas
Private Const m_FONTSTYLE_FOOTNOTE_RANGE As String = "Cambria"
Private Const m_FONTSTYLE_FOOTNOTE_REFERENCE As String = "Cambria"

' Style names
Private Const m_STYLE_PROGRAMMING_BLUEMETHOD As String = "Programming Method Blue"
Private Const m_STYLE_PROGRAMMING_CLASSNAME As String = "Programming Class Name"
Private Const m_STYLE_PROGRAMMING_GREYMETHOD As String = "Programming Method Darker"
Private Const m_STYLE_DEFAULTTIMESNEWROMAN As String = "Default Times New Roman"

' Font names
Private Const m_FONTNAME_EQUATION_DEFAULT As String = "Cambria Math"
Private Const m_FONTNAME_TIMESNEWROMAN As String = "Times New Roman"
Private Const m_FONTNAME_CALIBRILIGHT As String = "Calibri Light"
Private Const m_FONTNAME_CONSOLAS As String = "Consolas"
Private Const m_FONTNAME_SEGOE As String = "Segoe UI"

' Error constants
Private Const m_ERROR_STYLENAMEALREADYEXISTS As Long = 5173
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Sub Fmt_PerformAll()
'****************************************
' Compilation of formatting we want to perform on entire document containing reference material
' (i.e. notes) so that we don't have to run each command piecemeal.
'****************************************
   
   Const DISPLAY_END_MSG As Boolean = False
   Fmt_AllEquations
   Fmt_Footnotes (DISPLAY_END_MSG)
   Tools_UpdateTableOfContents
   'Fmt_ChangeFontSizeIfSpecificName (DISPLAY_END_MSG)
   MsgBox "Finished performing all formatting routines.", , "Done"

End Sub

Public Sub Fmt_AddCustomFontStylesToDocument()

   On Error GoTo ERROR_MANAGER
   
   addNewStyleToDocument m_STYLE_PROGRAMMING_BLUEMETHOD, "Consolas", m_FONTSIZE_PROGRAMMING_BLUE, IEColorProgrammingBlue
   addNewStyleToDocument m_STYLE_PROGRAMMING_CLASSNAME, "Consolas", m_FONTSIZE_PROGRAMMING_CLASS, IEColorProgrammingClassName
   addNewStyleToDocument m_STYLE_PROGRAMMING_GREYMETHOD, "Consolas", m_FONTSIZE_PROGRAMMING_BLUE, IEColorGrey
   addNewStyleToDocument m_STYLE_DEFAULTTIMESNEWROMAN, "Times New Roman", m_FONTSIZE_DEFAULT, IEColorAuto
      
   Exit Sub
   
ERROR_MANAGER:
   If Err.number = m_ERROR_STYLENAMEALREADYEXISTS Then
      Resume Next
   Else:
      Dim errMsg As String
      errMsg = "Error encountered in d" + vbLf + _
         "Number: " + Err.number + vbLf + _
         "Description: " + Err.Description
      MsgBox errMsg, , "Error"
   End If

End Sub

Private Sub addNewStyleToDocument( _
   ByVal styleName As String, _
   ByVal fontName As String, _
   ByVal fontSize As Double, _
   ByRef fontColor As ENUM_CUSTOM_COLOR)
'****************************************
' Adds a new user-defined custom style to the document.
'****************************************

   Dim newStyle As Style
   Set newStyle = ActiveDocument.Styles.Add(styleName, g_WORDENUM_STYLETYPE_CHARACTER)
   
   Dim rgbProps As ITYPE_RGB_PROPS
   rgbProps = Create_RGBColor(fontColor)
   
   With newStyle.Font
      .Name = fontName
      .Size = fontSize
      .color = rgb(rgbProps.RedElement, rgbProps.GreenElement, rgbProps.BlueElement)
   End With
   
End Sub

Public Sub Fmt_ChangeFontSizeIfSpecificName(Optional ByVal msgUser As Boolean = True)
'****************************************
' Changes font size for words with certain sizes and Font.Name.
'****************************************
   
   Const FONT_NAME_TO_CHANGE As String = "Consolas"
   Const SIZE_OF_FONT_TO_BE_CHANGED As Double = 9.5
   Const NEW_FONT_SIZE As Double = 9
   
   Dim countChanged As Integer
   countChanged = 0
   Dim wrd As Range
   For Each wrd In ActiveDocument.words
      If wordFontShouldBeChanged(wrd, FONT_NAME_TO_CHANGE, SIZE_OF_FONT_TO_BE_CHANGED) Then
         wrd.Font.Size = NEW_FONT_SIZE
         countChanged = countChanged + 1
      End If
   Next wrd
   
   If Not msgUser Then Exit Sub
      
   Dim completionMsg As String
   If countChanged > 0 Then
      Select Case countChanged
         Case Is > 1
            completionMsg = countChanged & " words' font size was changed."
         Case 1
            completionMsg = "One word's font size was changed."
         Case 0
            completionMsg = "No words had their font size changed."
      End Select
   End If
   
   MsgBox completionMsg, , "Done"
   
End Sub

Public Sub Fmt_ChangeWordColorsInSelection()
'****************************************
' Changes the color of a certain font to another font color. Can be done on Selection or
' ActiveDocument.
'****************************************
       
   Dim colorToReplace As ITYPE_RGB_PROPS, replacementColor As ITYPE_RGB_PROPS
   colorToReplace = Create_RGBColor(IEColorStandardGreen)
   replacementColor = Create_RGBColor(IEColorDarkerBlue)
   
   Dim countChanged As Integer
   countChanged = 0
   Dim wrd As Range
   For Each wrd In ActiveDocument.words
      Dim wColor As ITYPE_RGB_PROPS
      wColor = Color_GetRGBPropsFromRange(wrd)
      If Color_IsSame(wColor, colorToReplace) Then
         changeFontColorOfRange wrd, replacementColor
         countChanged = countChanged + 1
      End If
   Next wrd
      
   Dim completionMsg As String
   If countChanged > 0 Then
      If countChanged > 1 Then
         completionMsg = countChanged & " words' color was changed."
      Else:
         completionMsg = countChanged & " word's color was changed."
      End If
   Else:
      completionMsg = "No words' color was changed."
   End If
   
   MsgBox completionMsg, , "Done"
   
End Sub

Public Sub Fmt_ConvertOneStyleToAnother()
'****************************************
' With the current Selection, changes all words belonging to one Style to a different style.
'****************************************

   Const STYLE_TO_REPLACE As String = m_STYLE_PROGRAMMING_CLASSNAME
   Const REPLACEMENT_STYLE As String = m_STYLE_PROGRAMMING_GREYMETHOD
   Dim wrd As Range
   For Each wrd In Selection.words
      If wrd.Style = STYLE_TO_REPLACE Then
         wrd.Style = REPLACEMENT_STYLE
      End If
   Next wrd

End Sub

Private Sub changeCustomStyleProperty()
'****************************************
' Shows how to perform a bespoke property replacement when we want to change only certain
' properties of a given style.
'****************************************

   ActiveDocument.Styles(m_STYLE_PROGRAMMING_CLASSNAME).Font.Size = m_FONTSIZE_PROGRAMMING_CLASS
   
End Sub

Public Sub Fmt_SelectionToProgrammingStyle()
'****************************************
' Formats current selection in such a way to distinguish it from the rest of the text as text
' indicating actual programming language.
' *** Word has limitations around applying formatting changes to discontiguous ranges within the
' Selection. A Range object cannot contain multiple subranges. Therefore, if the current selection
' is discontiguous, Selection.Range will return the last subrange that the user selected.
' Therefore, the best workaround is to apply formatting changes via a Style.
'****************************************

   Select Case Selection.Range.Style
      Case m_STYLE_PROGRAMMING_BLUEMETHOD
         Selection.Style = m_STYLE_PROGRAMMING_CLASSNAME
      Case m_STYLE_PROGRAMMING_CLASSNAME
         Selection.Style = m_STYLE_PROGRAMMING_GREYMETHOD
      Case m_STYLE_PROGRAMMING_GREYMETHOD
         Selection.Style = m_STYLE_DEFAULTTIMESNEWROMAN
      Case Else
         Selection.Style = m_STYLE_PROGRAMMING_BLUEMETHOD
   End Select
   
End Sub

Public Sub Fmt_CertainWordsToProgrammingStyle()
'****************************************
' User enters array of words. Whenever these words are encountered in the selected text, their
' font style is changed to the programming style.
'****************************************
   
   Dim wrds As Variant
   wrds = Split(InputBox("To convert to programming style.", "Enter Words"))
   
   Dim w As Range
   For Each w In Selection.words
      Call applyProgrammingStyleIfOneOfKeyWords(w, wrds)
   Next w

End Sub

Public Sub Fmt_BoldCertainWords()
'****************************************
' User enters array of words. Whenever these words are encountered in the selected text, their
' font is bolded.
'****************************************

   ' Get array of words to bold from user.
   Dim wrds As Variant
   wrds = Split(InputBox("These words will be bolded:", "Enter Words"))
   
   ' Italicize word if it's contained in array of words to italicize.
   Dim w As Range
   For Each w In Selection.words
      If Array_ContainsValue(wrds, Trim(w.Text)) Then
         w.Font.Bold = True
      End If
   Next w

End Sub

Public Sub Fmt_ItalicizeCertainWords()
'****************************************
' User enters array of words. Whenever these words are encountered in the selected text, their
' font is italicized.
'****************************************

   ' Get array of words to italicize from user.
   Dim wrds As Variant
   wrds = Split(InputBox("These words will be italicized:", "Enter Words"))
   
   ' Italicize word if it's contained in array of words to italicize.
   Dim w As Range
   For Each w In Selection.words
      If Array_ContainsValue(wrds, Trim(w.Text)) Then
         w.Font.Italic = True
      End If
   Next w

End Sub

Public Sub Fmt_UpdateHeadingStyles()
'****************************************
' Updates all the heading styles in the ActiveDocument to my favorite formats.
'****************************************

   updateHeadingStyle wdStyleHeading1, True, wdUnderlineSingle, m_FONTNAME_SEGOE, 11.5, m_FONT_TEXTCOLOR_HEADERBLUE, True, m_PARAGRAPH_SPACEBEFORE_DEFAULT, m_PARAGRAPH_SPACEAFTER_DEFAULT
   updateHeadingStyle wdStyleHeading2, True, wdUnderlineNone, m_FONTNAME_SEGOE, 11, m_FONT_TEXTCOLOR_HEADERBLUE, False, m_PARAGRAPH_SPACEBEFORE_DEFAULT, m_PARAGRAPH_SPACEAFTER_DEFAULT
   updateHeadingStyle wdStyleHeading3, False, wdUnderlineNone, m_FONTNAME_SEGOE, 10, m_FONT_TEXTCOLOR_HEADERBLUE, False, m_PARAGRAPH_SPACEBEFORE_DEFAULT, m_PARAGRAPH_SPACEAFTER_DEFAULT
   updateHeadingStyle wdStyleHeading4, True, wdUnderlineNone, m_FONTNAME_SEGOE, 9, m_FONT_TEXTCOLOR_HEADERBLUE, False, m_PARAGRAPH_SPACEBEFORE_DEFAULT, m_PARAGRAPH_SPACEAFTER_DEFAULT

End Sub

Private Sub updateHeadingStyle( _
   ByVal styleName As String, _
   ByVal boldFont As Boolean, _
   ByVal underlineFont As WdUnderline, _
   ByVal fontName As String, _
   ByVal fontSize As Double, _
   ByVal fontColor As Long, _
   ByVal italicizeFont As Boolean, _
   ByVal spaceBeforeParagraphs As Double, _
   ByVal spaceAfterParagraphs As Double)
'****************************************
' Updates a heading style's properties to the one passed through by the user.
'****************************************
   
   With ActiveDocument.Styles(styleName)
      .Font.Bold = boldFont
      .Font.Underline = underlineFont
      .Font.Name = fontName
      .Font.Size = fontSize
      .Font.TextColor = fontColor
      .Font.Italic = italicizeFont
      .ParagraphFormat.spaceBefore = spaceBeforeParagraphs
      .ParagraphFormat.spaceAfter = spaceAfterParagraphs
   End With

End Sub

Public Sub Fmt_RemoveAllUserDefinedStyles()

   Dim sty As Style
   For Each sty In ActiveDocument.Styles
      If Not sty.BuiltIn Then sty.Delete
   Next sty

End Sub

Public Sub Fmt_WordColorsInDocument()
'****************************************
' Changes the color of a certain font in all words in ActiveDocument to another font color.
'****************************************
      
   Dim colorToReplace As ITYPE_RGB_PROPS
   colorToReplace = Create_RGBColor(IEColorStandardGreen)
   Dim replacementColor As ITYPE_RGB_PROPS
   replacementColor = Create_RGBColor(IEColorDarkerBlue)
   Dim countChanged As Integer
   countChanged = 0
   
   Dim wrd As Range
   For Each wrd In ActiveDocument.words
      Dim wColor As ITYPE_RGB_PROPS
      wColor = Color_GetRGBPropsFromRange(wrd)
      If Color_IsSame(wColor, colorToReplace) Then
         changeFontColorOfRange wrd, replacementColor
         countChanged = countChanged + 1
      End If
   Next wrd
      
   Dim completionMsg As String
   If countChanged > 0 Then
      If countChanged > 1 Then
         completionMsg = countChanged & " words' color was changed."
      Else:
         completionMsg = countChanged & " word's color was changed."
      End If
   Else:
      completionMsg = "No words' color was changed."
   End If
   
   MsgBox completionMsg, , "Done"
   
End Sub

Public Sub Fmt_FontNamesInDocument()
'****************************************
' Changes words with a specified font name to another font name in the document.
'****************************************
   
   Const FONT_NAME_TO_CHANGE As String = "Calibri"
   Const SIZE_OF_FONT_TO_BE_CHANGED As Double = 11
   Const NEW_FONT_NAME As String = "Times New Roman"
   
   Dim countChanged As Integer
   countChanged = 0
   
   Dim wrd As Range
   For Each wrd In ActiveDocument.words
      If wordFontShouldBeChanged(wrd, FONT_NAME_TO_CHANGE, SIZE_OF_FONT_TO_BE_CHANGED) Then
         wrd.Font.Name = NEW_FONT_NAME
         countChanged = countChanged + 1
      End If
   Next wrd
      
   Dim completionMsg As String
   If countChanged > 0 Then
      Select Case countChanged
         Case Is > 1
            completionMsg = countChanged & " words' font name was changed."
         Case 1
            completionMsg = "One word's font name was changed."
         Case 0
            completionMsg = "No words had their font name changed."
      End Select
   End If
   
   MsgBox completionMsg, , "Done"
   
End Sub

Private Function wordFontShouldBeChanged( _
   ByVal wrd As Range, _
   ByVal fontNameToBeChanged As String, _
   ByVal expectedFontSize As Double) As Boolean
'****************************************
' Checks whether a given word matches the font name and font size of a word whose font should be
' changed for following procedure: Fmt_FontNamesInDocument.
'****************************************

   If wrd.Font.Name = fontNameToBeChanged Then
      If wrd.Font.Size = expectedFontSize Then
         wordFontShouldBeChanged = True
         Exit Function
      End If
   End If
   wordFontShouldBeChanged = False

End Function

Private Sub applyProgrammingStyleIfOneOfKeyWords( _
   ByVal myRange As Range, _
   ByRef keyWordsArray As Variant)
'****************************************
' Applies our programming style to a given range if its text matches one of the input key words.
'****************************************
   
   Dim q As Integer
   For q = 0 To UBound(keyWordsArray)
      If Trim(myRange.Text) = keyWordsArray(q) Then
         myRange.Style = m_STYLE_PROGRAMMING_GREYMETHOD
         Exit Sub
      End If
   Next q
   
End Sub

Private Sub changeFontColorOfRange( _
   ByVal myRange As Range, _
   ByRef myColor As ITYPE_RGB_PROPS)
'****************************************
' Changes the color of the font in input range to that provided by the RGB color.
'****************************************
   
   Dim w As Range
   For Each w In myRange.words
      w.Font.color = rgb(myColor.RedElement, myColor.GreenElement, myColor.BlueElement)
   Next w

End Sub

Public Sub Fmt_Footnotes(Optional ByVal displayEndMessage As Boolean = True)
'****************************************
' Applies same font and paragraph formatting to all footnotes in ActiveDocument.
'****************************************

   On Error GoTo ERROR_MANAGER
   TurnScreenUpdatingOFF
   
   Dim fn As Footnote
   For Each fn In ActiveDocument.Footnotes
      With fn.Reference.Font
         .Name = m_FONTSTYLE_FOOTNOTE_REFERENCE
         .Size = m_FONTSIZE_FOOTNOTE_REFERENCE
         .Superscript = True
         .Bold = False
         .Italic = False
         .Underline = wdUnderlineNone
         .ColorIndex = wdAuto
      End With
      With fn.Range.Font
         .Name = m_FONTSTYLE_FOOTNOTE_RANGE
         .Size = m_FONTSIZE_FOOTNOTE_RANGE
      End With
      fn.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
   Next fn
   
   TurnScreenUpdatingON
   If displayEndMessage Then MsgBox "Finished formatting all footnotes.", , "Done"
   Exit Sub
   
ERROR_MANAGER:
   TurnScreenUpdatingON
   Select Case Err.number
      Case 6243   'Error when attempting to format references in fuctions
         Resume Next
      Case Else
         MsgBox "Error Number: " & Err.number & vbLf & _
            "Error Description: " & Err.Description, , "Error in Fmt_Footnotes"
   End Select

End Sub

Public Sub Fmt_FootnotesAll()

   Call Fmt_Footnotes(False)

End Sub

Public Sub Fmt_AllEquations()
'****************************************
' Executes equation formatting code by providing desired function inputs.
'****************************************
         
   On Error GoTo ERROR_MANAGER
   
   TurnScreenUpdatingOFF
   formatAllEquationsInDocument m_FONTNAME_EQUATION_DEFAULT, m_FONTSIZE_EQUATIONS, m_ITALICIZE_EQUATION_FONT
   TurnScreenUpdatingON
   
   Exit Sub
   
ERROR_MANAGER:
   TurnScreenUpdatingON
   MsgBox "Error Number: " & Err.number & vbLf & _
      "Error Description: " & Err.Description, , "Error in Fmt_Footnotes"
      
End Sub

Private Sub formatAllEquationsInDocument( _
   ByVal fontStyleName As String, _
   ByVal fontSize As String, _
   Optional ByVal italicizeEquation As Boolean = True)
'****************************************
' Applies formatting changes to all equations in document.
' Loops through each OMath (equation) object in ActiveDocument's OMaths collection and for each
' OMath object defines the range containing the equation so that we can modify the range's font
' formatting settings.
'****************************************
      
   Dim i As Integer
   For i = 1 To ActiveDocument.OMaths.Count
      Dim equationRange As Range
      Set equationRange = ActiveDocument.OMaths.Item(i).Range
      equationRange.Font.Size = fontSize
      equationRange.Font.Name = fontStyleName
      equationRange.Font.Italic = italicizeEquation
   Next i
      
End Sub

Public Sub Fmt_PageSetup()
'****************************************
' Adjusts page margins and distance from header/footer to body text.
'****************************************
      
   ' Define different margins for page (in inches)
   Const MARGIN_TOP As Double = 0.2
   Const MARGIN_BOTTOM As Double = 0.2
   Const MARGIN_LEFT As Double = 0.2
   Const MARGIN_RIGHT As Double = 0.2
   Const HEADER_DISTANCE = 0.1
   Const FOOTER_DISTANCE = 0.1
   
   ' Edit page properties.
   With ActiveDocument.PageSetup
      .LineNumbering.Active = False
      .Orientation = wdOrientPortrait
      .TopMargin = InchesToPoints(MARGIN_TOP)
      .BottomMargin = InchesToPoints(MARGIN_BOTTOM)
      .LeftMargin = InchesToPoints(MARGIN_LEFT)
      .RightMargin = InchesToPoints(MARGIN_RIGHT)
      .Gutter = InchesToPoints(0)
      .HeaderDistance = InchesToPoints(HEADER_DISTANCE)
      .FooterDistance = InchesToPoints(FOOTER_DISTANCE)
      .PageWidth = InchesToPoints(8.5)
      .PageHeight = InchesToPoints(11)
      .FirstPageTray = wdPrinterDefaultBin
      .OtherPagesTray = wdPrinterDefaultBin
      .SectionStart = wdSectionNewPage
      .OddAndEvenPagesHeaderFooter = False
      .DifferentFirstPageHeaderFooter = False
      .VerticalAlignment = wdAlignVerticalTop
      .SuppressEndnotes = False
      .MirrorMargins = False
      .TwoPagesOnOne = False
      .BookFoldPrinting = False
      .BookFoldRevPrinting = False
      .BookFoldPrintingSheets = 1
      .GutterPos = wdGutterPosLeft
   End With

End Sub

Public Sub Fmt_SelectedPictureDimensions(ByVal adjustmentFactor As Double)
'****************************************
' Adjusts height & width of selected shape by same AdjustmentFactor.
'****************************************

   Dim selectedShape As InlineShape: Set selectedShape = Selection.InlineShapes(1)
   
   Dim newHeight As Integer, newWidth As Integer
   newHeight = adjustmentFactor * selectedShape.Height
   newWidth = adjustmentFactor * selectedShape.Width
   
   With selectedShape
      .Height = newHeight
      .Width = newWidth
   End With

End Sub

Public Sub Fmt_TableNumberingIndentation()

   Const LEFT_INDENT As Double = 0.1
   Const RIGHT_INDENT As Double = 0
   Const FIRSTLINE_INDENT As Double = -0.1
   Const SPACE_BEFORE As Integer = 2
   
   With Selection.ParagraphFormat
      .LeftIndent = InchesToPoints(LEFT_INDENT)
      .RightIndent = InchesToPoints(RIGHT_INDENT)
      .spaceBefore = SPACE_BEFORE
      .SpaceBeforeAuto = False
      .spaceAfter = 0
      .SpaceAfterAuto = False
      .LineSpacingRule = wdLineSpaceSingle
      .Alignment = wdAlignParagraphLeft
      .WidowControl = True
      .KeepWithNext = False
      .KeepTogether = False
      .PageBreakBefore = False
      .NoLineNumber = False
      .Hyphenation = True
      .FirstLineIndent = InchesToPoints(FIRSTLINE_INDENT)
      .OutlineLevel = wdOutlineLevelBodyText
      .CharacterUnitLeftIndent = 0
      .CharacterUnitRightIndent = 0
      .CharacterUnitFirstLineIndent = 0
      .LineUnitBefore = 0
      .LineUnitAfter = 0
      .MirrorIndents = False
      .TextboxTightWrap = wdTightNone
   End With

End Sub

Public Sub Fmt_ParagraphSpacingToDefaults()
'****************************************
' Adjusts the paragraph SpaceAfter and SpaceBefore properties to our preferences IF the paragraph
' is not in a table or part of a list.
' If doc contains more than a certain number of paragraphs, we restrict our check to a
' user-defined range of paragraphs because running it on an entire document containing many
' paragraphs may cause Microsoft Word to hang.
'****************************************

   On Error GoTo EXIT_ORDERLY
   TurnScreenUpdatingOFF
   Const MAX_PARAGRAPHS_TO_FORMAT As Integer = 1500
      
   '''''''''''''''''''''''''''''''''''''''
   ' Determine which paragraphs we're going to review paragraph spacing on based on total number
   ' of paragraphs in document.
   '''''''''''''''''''''''''''''''''''''''
   Dim pgraphsInDoc As Integer
   pgraphsInDoc = ActiveDocument.Paragraphs.Count
   Dim firstPgraphToEdit As Integer, lastPgraphToEdit As Integer
   If pgraphsInDoc > MAX_PARAGRAPHS_TO_FORMAT Then
      Dim paragraphCountIsOK As Boolean
      paragraphCountIsOK = False
      Do
         Dim msg As String
         msg = "Because the number of paragraphs in the document (" & pgraphsInDoc & _
            ") exceeds our limit to check all " & "(" & MAX_PARAGRAPHS_TO_FORMAT & _
            ") you will be entering the first and last paragraphs to edit."
         MsgBox msg, , "Heads up"
         firstPgraphToEdit = CInt(InputBox("Enter 1st paragraph number to edit", "First Paragraph"))
         lastPgraphToEdit = CInt(InputBox("Enter last paragraph number to edit", "Last Paragraph"))
         If lastPgraphToEdit > pgraphsInDoc Then lastPgraphToEdit = pgraphsInDoc
         If (lastPgraphToEdit - firstPgraphToEdit + 1) <= MAX_PARAGRAPHS_TO_FORMAT Then
            paragraphCountIsOK = True
         Else:
            MsgBox "Number of paragraphs still exceeds our maximum. Let's try this again.", , "Hold On"
         End If
      Loop While Not paragraphCountIsOK
   Else:
      firstPgraphToEdit = 1
      lastPgraphToEdit = pgraphsInDoc
   End If

   '''''''''''''''''''''''''''''''''''''''
   ' Determine which paragraphs we're going to review paragraph spacing on based on total number
   ' of paragraphs in document.
   '''''''''''''''''''''''''''''''''''''''
   Const STATUS_UPDATE_MULT = 20 ' dictates how often we print an update to the StatusBar
   Dim pCount As Integer: pCount = 0
   Dim p As word.Paragraph
   For Each p In ActiveDocument.Paragraphs
      If pCount > lastPgraphToEdit Then
         GoTo EXIT_ORDERLY
      Else: pCount = pCount + 1
      End If
      
      If pCount >= firstPgraphToEdit Then
         If paragraphIsNotInATableOrList(p) Then
            DisplayNumberOnStatusBarIfMultiple pCount, STATUS_UPDATE_MULT, "Updating paragraph # "
            adjustSpacingAfterParagraph p.Range, m_PARAGRAPH_SPACEAFTER_DEFAULT, m_PARAGRAPH_SPACEBEFORE_DEFAULT
            DoEvents
         End If
      End If
   Next p
   
EXIT_ORDERLY:
   Application.StatusBar = False
   TurnScreenUpdatingON
   MsgBox "Finished adjusting paragraph spacing.", , "Done"

End Sub

Private Function paragraphIsNotInATableOrList(ByVal para As word.Paragraph) As Boolean
'****************************************
' Returns True if a paragraph is not part of a list or located within a table.
'****************************************

   If para.Range.ListFormat.listType = g_WORDENUM_LISTTYPE_NONE Then
      If para.Range.Information(g_WORDENUM_INFO_WITHINTABLE) Then
         paragraphIsNotInATableOrList = False
      Else: paragraphIsNotInATableOrList = True
      End If
   Else: paragraphIsNotInATableOrList = False
   End If

End Function

Private Sub adjustSpacingAfterParagraph( _
   ByVal r As Range, _
   ByVal spaceAfter As Integer, _
   ByVal spaceBefore As Integer)

   r.ParagraphFormat.spaceAfter = spaceAfter
   r.ParagraphFormat.spaceBefore = spaceBefore

End Sub

Public Sub Fmt_FontSuperscript(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection.Range
   rng.Font.Superscript = True

End Sub

Public Sub Fmt_ResizeImage()
'****************************************
' Resizes relative dimensions of selected shape(s) by predefined factors.
'****************************************
   
   On Error GoTo TURN_ON_UPDATES
      
   Application.ScreenUpdating = False ' to run faster
   
   Dim s As InlineShape
   For Each s In Selection.InlineShapes
      s.ScaleHeight = m_RESIZE_PROPORTION
      s.ScaleWidth = m_RESIZE_PROPORTION
   Next s
   
TURN_ON_UPDATES:
   Application.ScreenUpdating = True

End Sub

Public Sub Fmt_SelectionSuperscript()
' Recorded Macro using Word's superscript functionality: Selection.Font.Superscript = wdToggle

   Fmt_FontSuperscript Selection.Range
   Selection.MoveRight Unit:=wdCharacter, Count:=1
   Selection.Font.Superscript = False

End Sub

Public Sub Fmt_SubscriptLastCharOfMathExpressions()
'****************************************
' Subscripts the last character
'****************************************

   Dim w As word.Range
   For Each w In Selection.words
      Dim patternMatch As Variant
      patternMatch = wordIsMathExpression(w)
      If patternMatch(0) Then
         Fmt_SubscriptLastChar w, patternMatch(1)
         ' Assumes all math expressions only contain one character before the subscript symbol.
         Fmt_ItalicizeWord w, 1
      End If
   Next w

End Sub

Private Function wordIsMathExpression(ByVal w As Range) As Variant
'****************************************
' Returns Array(bool, int).
' Where bool True if w is something like x1, q2, an, ci, etc.
' and, if True, a non-zero int representing length of subscript code.
'****************************************
   
   Dim txt As String
   txt = Trim(w.Text)
   
   ' Define list of words that will fit one of the patterns but which we do not want to
   ' format as a result.
   Dim wToExclude As Variant
   wToExclude = Array("an", "at", "In", "in", "it", "on")
   
   If Array_ContainsValue(wToExclude, txt) Then
     GoTo EXIT_FALSE
   End If
   
   Dim firstParts As Variant, subscriptCodes As Variant
   firstParts = Array("[A-Z]", "[a-z]", "µ")
   subscriptCodes = Array("#", "a", "b", "i", "j", "m", "M", "n", "N", "k", "K", "t", "t+1")
   
   Dim mathPattern As String
   Dim i As Integer
   For i = 0 To UBound(firstParts)
      Dim j As Integer
      For j = 0 To UBound(subscriptCodes)
         mathPattern = firstParts(i) + subscriptCodes(j)
         If txt Like mathPattern Then
            wordIsMathExpression = Array(True, Len(subscriptCodes(j)))
            Exit Function
         End If
      Next j
   Next i
      
EXIT_FALSE:
   wordIsMathExpression = Array(False, 0)

End Function

Public Sub Fmt_FontSubscript(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection.Range
   rng.Font.Subscript = True

End Sub

Public Sub Fmt_SelectionSubscript()

   Call Fmt_FontSubscript(Selection.Range)
   Selection.MoveRight Unit:=wdCharacter, Count:=1
   Selection.Font.Subscript = False

End Sub

Public Sub Fmt_SuperscriptLastLetter(ByVal w As Range)

   Dim lastCharacterIndex As Integer
   lastCharacterIndex = Txt_LastCharIndex(w)
   Call Fmt_FontSuperscript(w.Characters(lastCharacterIndex))

End Sub

Public Sub Fmt_SubscriptLastChar(ByVal w As Range, Optional ByVal lastChars As Integer = 1)
   
   Dim i As Integer
   i = Txt_LastCharIndex(w)
   Dim q As Integer
   For q = 0 To lastChars - 1 Step 1
      Fmt_FontSubscript w.Characters(i - q)
   Next q
   
End Sub

Public Sub Fmt_BoldWord(ByVal wrd As Range)

   wrd.Font.Bold = True

End Sub

Public Sub Fmt_ItalicizeWord(ByVal wrd As Range, Optional ByVal firstChars As Integer = 0)
   
   If firstChars = 0 Then
      wrd.Font.Italic = True
      Exit Sub
   End If
   
   Dim char As Integer
   For char = 1 To firstChars
      wrd.Characters(char).Font.Italic = True
   Next char

End Sub

Public Sub Fmt_ChangeFontColorForAllWords( _
   ByVal oldColor As ENUM_CUSTOM_COLOR, _
   ByVal newColor As ENUM_CUSTOM_COLOR)
'****************************************
' Replaces the Font color of every word to newColor whose current font is oldColor.
'****************************************

   Dim newColorRGBProps As ITYPE_RGB_PROPS
   newColorRGBProps = Create_RGBColor(newColor)
   
   Dim changeCount As Integer
   changeCount = 0
   Dim sentence As Object
   For Each sentence In ActiveDocument.StoryRanges
      Dim w As Range
      For Each w In sentence.words
         If Color_CustomFromRange(w) = oldColor Then
            changeFontColorOfRange w, newColorRGBProps
            changeCount = changeCount + 1
         End If
      Next
   Next
   
   MsgBox "Changed the font color of " & changeCount & " words in document.", , "Done"

End Sub

Public Sub Fmt_MakeGreenWordsBlue()

   Dim colorToRemove As ENUM_CUSTOM_COLOR
   colorToRemove = ENUM_CUSTOM_COLOR.IEColorStandardGreen
   Dim newColor As ENUM_CUSTOM_COLOR
   newColor = ENUM_CUSTOM_COLOR.IEColorDarkerBlue
   
   Call Fmt_ChangeFontColorForAllWords(colorToRemove, newColor)

End Sub



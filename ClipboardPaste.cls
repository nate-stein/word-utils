VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClipBoardPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   CLIPBOARD PASTE WRAPPER
' PURPOSE:  Provide methods to alter and, overall, better format text that we copy and want to
'           paste into our Word document.

Private mText As String
Private mRemoveNewLineChars As Boolean
Private mPerformCodingLanguageEdits As Boolean
Private mConvertSpecialPunctuation As Boolean
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Sub Edit()

   removeKindleReferences
   If mRemoveNewLineChars Then removeNewLineCharsFromTxt
   If mPerformCodingLanguageEdits Then editCodingLanguage
   makeOtherSpecialReplacements
   convertSpecialPunctToWordStandard

End Sub

Private Sub removeKindleReferences()

   ' Define array containing references that the Kindle or other devices automatically add in but
   ' which we want to prune off the end of a string we are pasting.
   Dim refs As Variant
   refs = Notes_GetKindleReferencesArray()
   
   ' Loop through array and remove unwanted references.
   Dim i As Integer
   For i = 0 To UBound(refs) Step 1
      If InStr(1, mText, refs(i)) > 0 Then
         mText = Left(mText, InStr(1, mText, refs(i)) - 2)
      End If
   Next i

End Sub

Private Sub removeNewLineCharsFromTxt()

   mText = Txt_RemoveNewLineChars(mText)

End Sub

Private Sub editCodingLanguage()

   reformatPythonIDLEBracketReferences
   removeHtmlBracketEmptySpace

End Sub

Private Sub reformatPythonIDLEBracketReferences()

   mText = Replace(mText, "> > >", vbCrLf + ">>>")
   
End Sub

Private Sub removeHtmlBracketEmptySpace()

   mText = Replace(mText, " >", ">")
   mText = Replace(mText, "< ", "<")

End Sub

Private Sub convertSpecialPunctToWordStandard()
   
   mText = Txt_ConvertSpecialPunctuation(mText)

End Sub

Private Sub makeOtherSpecialReplacements()

   mText = Replace(mText, "$ ", "$")
   mText = Replace(mText, "( ", "(")
   mText = Replace(mText, " )", ")")
   mText = Replace(mText, "/ ", "/")
   mText = Replace(mText, " /", "/")

End Sub

Public Property Let Text(ByVal txt As String)

   mText = txt

End Property

Public Property Get FormattedText() As String

   FormattedText = mText

End Property

Public Property Let RemoveNewLines(ByVal remove As Boolean)

   mRemoveNewLineChars = remove

End Property

Public Property Let PerformCodingLanguageEdits(ByVal makeEdits As Boolean)

   mPerformCodingLanguageEdits = makeEdits

End Property

Public Property Let ConvertSpecialPunctuation(ByVal convert As Boolean)

   mConvertSpecialPunctuation = convert

End Property

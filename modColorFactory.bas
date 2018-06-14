Attribute VB_Name = "modColorFactory"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   COLOR FACTORY

Public Enum ENUM_CUSTOM_COLOR
   IEColorUnknown = 0
   IEColorProgrammingBlue = 1
   IEColorProgrammingClassName = 2
   IEColorAuto = 3
   IEColorGrey = 4
   IEColorStandardGreen = 5
   IEColorDarkerBlue = 6
   IEColorBlanketBlue = 7
End Enum

' Arrays used to determine the custom color from an RGB color.
Private mRedElements() As Variant
Private mGreenElements() As Variant
Private mBlueElements() As Variant
Private mCustomColors() As Variant
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Function Create_RGBColor(ByVal color As ENUM_CUSTOM_COLOR) As ITYPE_RGB_PROPS
'****************************************
' Returns ITYPE_RGB_PROPS for a given color.
'****************************************

   Dim r As Integer
   Dim g As Integer
   Dim b As Integer
   
   Select Case color
      Case IEColorProgrammingBlue
         r = 0
         g = 0
         b = 255
      Case IEColorProgrammingClassName
         r = 43
         g = 145
         b = 175
      Case IEColorAuto
         r = 0
         g = 0
         b = 0
      Case IEColorGrey
         r = 64
         g = 64
         b = 64
      Case IEColorStandardGreen
         r = 0
         g = 176
         b = 80
      Case IEColorDarkerBlue
         r = 0
         g = 112
         b = 192
      Case IEColorBlanketBlue
         r = 255
         g = 255
         b = 0
   End Select
   
   Create_RGBColor.RedElement = r
   Create_RGBColor.GreenElement = g
   Create_RGBColor.BlueElement = b

End Function

Public Function Create_CustomColorFromRGB( _
   ByRef rgbColor As ITYPE_RGB_PROPS) As ENUM_CUSTOM_COLOR

   Dim redCode As Integer, greenCode As Integer, blueCode As Integer
   redCode = rgbColor.RedElement
   greenCode = rgbColor.GreenElement
   blueCode = rgbColor.BlueElement

   Create_CustomColorFromRGB = _
      createCustomColorFromProperties(redCode, greenCode, blueCode)
   
End Function

Private Function createCustomColorFromProperties( _
   ByVal red As Integer, _
   ByVal green As Integer, _
   ByVal blue As Integer) As ENUM_CUSTOM_COLOR
'****************************************
' ASSUMES the arrays containing color elements and corresponding custom colors have already been
' defined.
'****************************************
   
   defineColorCodeArrays
   
   Dim redLoop As Integer
   For redLoop = 0 To UBound(mRedElements)
      If mRedElements(redLoop) = red Then
         If mGreenElements(redLoop) = green And mBlueElements(redLoop) = blue Then
            createCustomColorFromProperties = mCustomColors(redLoop)
            Exit Function
         End If
      End If
   Next redLoop

   createCustomColorFromProperties = IEColorUnknown

End Function

Private Sub defineColorCodeArrays()

   mRedElements = Array(0, 43, 0, 64, 0, 0)
   mGreenElements = Array(0, 145, 0, 64, 176, 112)
   mBlueElements = Array(255, 175, 0, 64, 80, 192)
   mCustomColors = Array(IEColorProgrammingBlue, IEColorProgrammingClassName, IEColorAuto, IEColorGrey, IEColorStandardGreen, IEColorDarkerBlue)

End Sub



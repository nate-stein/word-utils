Attribute VB_Name = "modColor"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   COLOR UTILS
' PURPOSE:  Provide methods to enable the interpretation and analysis of color.
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Function Color_CustomFromRange(ByVal rng As Range) As ENUM_CUSTOM_COLOR
'****************************************
' Returns ENUM_CUSTOM_COLOR determined from the RGB properties of rng's Font.
'****************************************

   Dim rgbColor As ITYPE_RGB_PROPS
   rgbColor = Color_GetRGBPropsFromRange(rng)
   Color_CustomFromRange = Create_CustomColorFromRGB(rgbColor)

End Function

Public Function Color_GetRGBPropsFromRange(Optional ByVal rng As Range) As ITYPE_RGB_PROPS
'****************************************
' Returns the RGB properties of rng's Font.Color property.
'****************************************

   If rng Is Nothing Then Set rng = Selection.Range
   
   Dim color As ITYPE_RGB_PROPS
   Dim hexColor As String
   hexColor = Right("000000" & Hex(rng.Font.color), 6)
   color.RedElement = CInt("&H" & Right(hexColor, 2))
   color.GreenElement = CInt("&H" & Mid(hexColor, 3, 2))
   color.BlueElement = CInt("&H" & Left(hexColor, 2))
   Color_GetRGBPropsFromRange = color

End Function

Public Function Color_IsSame(ByRef c1 As ITYPE_RGB_PROPS, ByRef c2 As ITYPE_RGB_PROPS) As Boolean
'****************************************
' Returns False if any of the R/G/B properties don't match between colors c1 and c2.
'****************************************
   
   Color_IsSame = True
   If c1.BlueElement <> c2.BlueElement Then Color_IsSame = False
   If c1.GreenElement <> c2.GreenElement Then Color_IsSame = False
   If c1.RedElement <> c2.RedElement Then Color_IsSame = False

End Function



















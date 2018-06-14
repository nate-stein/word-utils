Attribute VB_Name = "modDeclarations"
Option Explicit

' Custom Types
Public Type ITYPE_SHAPE_DIMS
   Height As Double
   Width As Double
End Type

Public Type ITYPE_RGB_PROPS
   RedElement As Integer
   GreenElement As Integer
   BlueElement As Integer
End Type

' Custom Enums
Public Enum IENUM_SHAPE_TYPE
   Rectangle = 1
   Oval = 2
End Enum

' Error constants
Public Const g_ERROR_NOTES_FACTORY As Integer = 2005
Public Const g_ERROR_ARRAYS As Integer = 2006

' Microsoft Word Enum constants
Public Const g_WORDENUM_INFO_FIRSTCHARACTERLINENUMBER As Integer = 10   ' wdFirstCharacterLineNumber
Public Const g_WORDENUM_INFO_WITHINTABLE As Integer = 12                ' wdWithInTable
Public Const g_WORDENUM_INFO_ZOOMPERCENTAGE As Integer = 19             ' wdZoomPercentage
Public Const g_WORDENUM_STYLETYPE_CHARACTER As Integer = 2              ' wdStyleTypeCharacter
Public Const g_WORDENUM_LISTTYPE_NONE As Integer = 0                    ' wdListNoNumbering



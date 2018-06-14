Attribute VB_Name = "modTable"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   TABLES
' PURPOSE:  Tools to automate the editing and finetuning of tables.
' SUBS:     FormatAccountingTable
'           IndentLeftmostColumnsInSelectedRows
'           RemoveAnySpacingInLeftmostColumns
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Sub Table_FormatAccountingTable()
'********************************************************
' Applies formatting desirable to tables containing accounting data.
'********************************************************

   Const SPACE_BEFORE As Double = 2
   Const SPACE_AFTER As Double = 2
   Const DESIRED_FONT_SIZE As Double = 10
   Selection.Tables(1).Select
   Selection.Font.Size = DESIRED_FONT_SIZE
   changeIndentLevelsInCell Selection.Range, SPACE_BEFORE, SPACE_AFTER

End Sub

Private Sub changeIndentLevelsInCell( _
   ByVal cells As Range, _
   ByVal before As Double, _
   ByVal after As Double)

    With cells.ParagraphFormat
      .spaceBefore = before
      .SpaceBeforeAuto = False
      .spaceAfter = after
      .SpaceAfterAuto = False
    End With

End Sub

Public Sub Table_IndentLeftmostColumnsInSelectedRows()

   ' 72 points = 1 inch
   Const LEFT_INDENT_POINTS As Integer = 18
   
   Dim firstRow As Integer: firstRow = getFirstRowOfSelectionInTable()
   Dim lastRow As Integer: lastRow = getLastRowOfSelectionInTable()
   
   Dim row As Integer
   For row = firstRow To lastRow Step 1
      Dim leftmostCell As Range
      Set leftmostCell = Selection.Tables(1).rows(row).cells(1).Range
      leftmostCell.Paragraphs.LeftIndent = LEFT_INDENT_POINTS
   Next row

End Sub

Public Sub Table_RemoveAnySpacingInLeftmostColumns()

   Dim firstRow As Integer: firstRow = getFirstRowOfSelectionInTable()
   Dim lastRow As Integer: lastRow = getLastRowOfSelectionInTable()
   
   Dim row As Integer
   For row = firstRow To lastRow Step 1
      Dim leftmostCell As Range
      Set leftmostCell = Selection.Tables(1).rows(row).cells(1).Range
      ' Move to leftmost of cell contents
      leftmostCell.Text = Trim(leftmostCell.Text)
      leftmostCell.Collapse Direction:=wdCollapseEnd
      leftmostCell.Delete
   Next row

End Sub

Private Function getFirstRowOfSelectionInTable() As Integer

   getFirstRowOfSelectionInTable = Selection.cells(1).RowIndex

End Function

Private Function getLastRowOfSelectionInTable() As Integer

   Dim lastSelectedCell As Integer: lastSelectedCell = Selection.cells.Count
   getLastRowOfSelectionInTable = Selection.cells(lastSelectedCell).RowIndex

End Function

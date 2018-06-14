Attribute VB_Name = "modEquation"
Option Explicit
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
' MODULE:   EQUATIONS
' PURPOSE:  Automate the insertion and revision of equations.

' WdOMathFunctionType Enumeration
Private Const m_FUNCTIONTYPE_ACCENTMARK As Integer = 1  ' wdOMathFunctionAcc

' Char constants
Private Const m_ACCENTCHAR_BAR As Integer = 773
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************

Public Sub Equations_AddAccentBarToSelection()
'*********************************************************
' Converts selection to an equation and adds an accent bar.
'*********************************************************

   Dim equationRange As Range
   Set equationRange = Selection.OMaths.Add(Selection.Range)
   
   Dim equation As OMath
   Set equation = equationRange.OMaths(1)
   
   equation.Functions.Add(equationRange, wdOMathFunctionAcc).Acc.char = m_ACCENTCHAR_BAR
   equation.Type = wdOMathInline

End Sub

Public Sub Equation_1By2Matrix()

   Const r As Integer = 1
   Const c As Integer = 2
   
   insertMatrix r, c

End Sub

Public Sub Equation_2By1Matrix()

   Const r As Integer = 2
   Const c As Integer = 1
   
   insertMatrix r, c

End Sub

Public Sub Equation_2By2Matrix()

   Const r As Integer = 2
   Const c As Integer = 2
   
   insertMatrix r, c

End Sub

Public Sub Equation_MByNMatrix()

   On Error GoTo EXIT_SUB
   Dim rows As Integer, cols As Integer
   rows = InputBox("Enter # of rows")
   cols = InputBox("Enter # of columns")
   
   insertMatrix rows, cols

EXIT_SUB:
   
End Sub

Private Sub insertMatrix(ByVal rows As Integer, cols As Integer)

   On Error GoTo EXIT_ORDERLY
   
   Application.ScreenUpdating = False
   
   Dim matrixFormat As String
   matrixFormat = createLinearMatrixFormat(rows, cols)
   
   matrixFormat = "[\matrix(" & matrixFormat & ")]"
   
   Dim objRange As Range
   Set objRange = Selection.Range
   objRange.Text = matrixFormat
   
   Dim AC As OMathAutoCorrectEntry
   Application.OMathAutoCorrect.UseOutsideOMath = True
   For Each AC In Application.OMathAutoCorrect.Entries
      With objRange
         If InStr(.Text, AC.Name) > 0 Then
            .Text = Replace(.Text, AC.Name, AC.value)
         End If
      End With
   Next AC
   
   Dim equation As OMath
   Set equation = Selection.OMaths.Add(objRange).OMaths.Item(1)
   equation.BuildUp
   equation.Justification = wdOMathJcLeft
   Selection.Paragraphs.Indent
      
   Application.ScreenUpdating = True
   Exit Sub
   
EXIT_ORDERLY:
   MsgBox "Error encountered in insertMatrix()"
   Application.ScreenUpdating = True

End Sub

Private Function createLinearMatrixFormat(ByVal rows As Integer, cols As Integer) As String
'*********************************************************
' Creates text needed to tell Microsoft Word how to define a Matrix to insert.
' & corresponds to columns, and @ corresponds to rows.
' For x & characters, you get x+1 columns. For y @’s, you get y+1 rows.
' \matrix(&&@&&@&&)
'*********************************************************

   Dim columnRepresentation As String
   columnRepresentation = createColumnRepresentation(cols)

   Dim result As String
   result = ""
   
   If rows = 1 Then
      result = result & columnRepresentation
   Else:
      ' This effectively accounts for
      result = result & columnRepresentation
      Dim r As Integer
      For r = 2 To rows Step 1
         result = result & "@" & columnRepresentation
      Next r
   End If
   
   createLinearMatrixFormat = "\matrix(" & result & ")"

End Function

Private Function createColumnRepresentation(ByVal cols As Integer) As String
   
   If cols = 1 Then
      createColumnRepresentation = ""
   Else:
      Dim c As Integer
      For c = 2 To cols Step 1
         createColumnRepresentation = createColumnRepresentation & "&"
      Next c
   End If
   
End Function

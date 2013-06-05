﻿Imports System.Runtime.InteropServices

Namespace ScriptDB

  <ClassInterface(ClassInterfaceType.None), Serializable()>
  Public Class LinesOfCode
    Inherits Collection(Of ScriptDB.CodeElement)

    Private mbAppendAfterNext As Boolean
    Private mbIsComparison As Boolean

    Private miNextInsertPoint As Integer
    Private miLastInsertOperatorType As OperatorSubType

    Public Property CodeLevel As Integer
    Public Property ReturnType As ComponentValueTypes

    Public Overloads Sub Add(ByVal LineOfCode As ScriptDB.CodeElement)

      If mbAppendAfterNext Then
        Me.Items.Insert(Me.Items.Count - 1, LineOfCode)
        mbAppendAfterNext = False
      Else
        Me.Items.Add(LineOfCode)

        If miLastInsertOperatorType <> LineOfCode.OperatorType Then
          miLastInsertOperatorType = LineOfCode.OperatorType
          miNextInsertPoint = Me.Items.Count - 1
        End If
      End If

    End Sub

    Public Overloads Sub InsertBeforePrevious(ByVal LineOfCode As ScriptDB.CodeElement)
      Me.Items.Insert(miNextInsertPoint, LineOfCode)
    End Sub

    Public Overloads Sub AppendAfterNext(ByVal LineOfCode As ScriptDB.CodeElement)
      mbAppendAfterNext = True
      Me.Items.Add(LineOfCode)
    End Sub

    ' Property to calculate the character indenation in the code (to beautify the code)
    Public ReadOnly Property Indentation() As String
      Get
        Return Space(8)
      End Get
    End Property

    Public Function ToArray() As String()
      Return Me.Items.Select(Function(c) c.Code).ToArray()
    End Function

    Public ReadOnly Property Statement() As String
      Get
        Statement = String.Empty

        Dim Chunk As ScriptDB.CodeElement
        Dim iThisElement As Integer
        Dim bComparisonSinceLastLogic As Boolean
        Dim bAddAutoIsEqualTo As Boolean

        Statement = String.Empty
        For Each Chunk In Me.Items
          bAddAutoIsEqualTo = False

          If Chunk.OperatorType = OperatorSubType.Comparison Then mbIsComparison = True

          If ReturnType = ComponentValueTypes.Logic Then

            If Chunk.CodeType = ComponentTypes.Operator Then
              If Chunk.OperatorType = OperatorSubType.Comparison Then
                bComparisonSinceLastLogic = True
              ElseIf Chunk.OperatorType = OperatorSubType.Logic Then
                bComparisonSinceLastLogic = False
              End If
            End If

            ' Is there an operator after this component?
            If Me.Items.Count - 1 > iThisElement Then
              If Me.Items(iThisElement + 1).OperatorType = OperatorSubType.Logic Then
                If Not bComparisonSinceLastLogic Then
                  bAddAutoIsEqualTo = True
                End If
              End If
            Else
              bAddAutoIsEqualTo = Not bComparisonSinceLastLogic
            End If

            ' Am I the last component and was there an operator before me?
            If iThisElement = Me.Items.Count - 1 And iThisElement > 0 Then
              If Me.Items(iThisElement - 1).OperatorType = OperatorSubType.Logic Then
                If Not bComparisonSinceLastLogic Then
                  bAddAutoIsEqualTo = True
                End If
              End If
            End If

          End If

          ' Does this code element make needing logic safe?
          ' Some logic expressions return a simple logic while others are set specifically 
          ' to logicval = 0, or have the not in front of them!
          If bAddAutoIsEqualTo Then
            Statement = vbNewLine & String.Format("{0}{1}{2} = 1", New String(CChar(vbTab), CodeLevel), Statement, Chunk.Code)
            bAddAutoIsEqualTo = False
          Else
            Statement = vbNewLine & String.Format("{0}{1}{2}", New String(CChar(vbTab), CodeLevel), Statement, Chunk.Code)
          End If

          'If iThisElement > 0 Then
          '  Statement = vbNewLine & Statement
          'End If

          ' Statement = Statement & Chunk.Code
          iThisElement = iThisElement + 1

        Next

        ' Wrap to return code chunks in safety
        If Me.ReturnType = ComponentValueTypes.Logic Or mbIsComparison Then
          Statement = String.Format("{0}CASE WHEN ({1}) THEN 1 ELSE 0 END", New String(CChar(vbTab), CodeLevel), Statement)
        End If

      End Get
    End Property

  End Class

End Namespace

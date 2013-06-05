Imports System.Runtime.InteropServices

Namespace ScriptDB

  <ClassInterface(ClassInterfaceType.None), Serializable()>
  Public Class LinesOfCode
    Inherits Collection(Of CodeElement)

    Private mbAppendAfterNext As Boolean
    Private miNextInsertPoint As Integer
    Private miLastInsertOperatorType As OperatorSubType

    Public Property CodeLevel As Integer
    Public Property ReturnType As ComponentValueTypes
    Public Property MakeTypesafe As Boolean = True
    Public Property IsLogicBlock As Boolean

    Public Overloads Sub Add(ByVal LineOfCode As CodeElement)

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

    Public Overloads Sub InsertBeforePrevious(ByVal LineOfCode As CodeElement)
      Me.Items.Insert(miNextInsertPoint, LineOfCode)
    End Sub

    Public Overloads Sub AppendAfterNext(ByVal LineOfCode As CodeElement)
      mbAppendAfterNext = True
      Me.Items.Add(LineOfCode)
    End Sub

    '' Property to calculate the character indenation in the code (to beautify the code)
    'Public ReadOnly Property Indentation() As String
    '  Get
    '    Return Space(8)
    '  End Get
    'End Property

    Public Function ToArray() As String()
      Return Me.Items.Select(Function(c) c.Code).ToArray()
    End Function

    Public ReadOnly Property Statement() As String
      Get

        Dim Chunk As CodeElement
        Dim iThisElement As Integer
        Dim bComparisonSinceLastLogic As Boolean

        Dim bAddAutoIsEqualTo As Boolean
        Dim bNewLine As Boolean

        Statement = String.Empty
        For Each Chunk In Me.Items
          bAddAutoIsEqualTo = False
          bNewLine = False

          If ReturnType = ComponentValueTypes.Logic Then

            If Chunk.CodeType = ComponentTypes.Operator Then

              If Chunk.OperatorType = OperatorSubType.Comparison Then
                bComparisonSinceLastLogic = True
                IsLogicBlock = True

              ElseIf Chunk.OperatorType = OperatorSubType.Logic Then
                bComparisonSinceLastLogic = False
                IsLogicBlock = True
                bNewLine = True

              ElseIf Chunk.OperatorType = OperatorSubType.Modifier Then
                IsLogicBlock = True
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
            Statement = String.Format("{0}{1}{2} = 1", Statement, IIf(bNewLine, vbNewLine & New String(CChar(vbTab), CodeLevel), vbNullString), Chunk.Code)
            IsLogicBlock = True
            bAddAutoIsEqualTo = False
          Else
            Statement = String.Format("{0}{1}{2}", Statement, IIf(bNewLine, vbNewLine & New String(CChar(vbTab), CodeLevel), vbNullString), Chunk.Code)
          End If

          iThisElement = iThisElement + 1

        Next

        ' Wrap to return code chunks in safety (was there 'not' statement)
        If IsLogicBlock And Me.ReturnType = ComponentValueTypes.Logic And MakeTypesafe Then
          Statement = String.Format("CASE WHEN ({0}) THEN 1 ELSE 0 END", Statement)
        End If

      End Get
    End Property

  End Class

End Namespace

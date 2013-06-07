Imports System.Runtime.InteropServices
Imports SystemFramework.Enums
Imports SystemFramework.Structures

Namespace ScriptDB

  <ClassInterface(ClassInterfaceType.None), Serializable()>
  Public Class LinesOfCode
    Inherits Collection(Of CodeElement)

    Private _mbAppendAfterNext As Boolean
    Private _miNextInsertPoint As Integer
    Private _miLastInsertOperatorType As OperatorSubType

    Public Property CodeLevel As Integer
    Public Property ReturnType As ComponentValueTypes
    Public Property MakeTypesafe As Boolean = True
    Public Property CaseReturnType As CaseReturnType = CaseReturnType.Result
    Public Property IsLogicBlock As Boolean

    Public Overloads Sub Add(ByVal lineOfCode As CodeElement)

      If _mbAppendAfterNext Then
        Items.Insert(Items.Count - 1, lineOfCode)
        _mbAppendAfterNext = False
      Else
        Items.Add(lineOfCode)

        If _miLastInsertOperatorType <> lineOfCode.OperatorType Then
          _miLastInsertOperatorType = lineOfCode.OperatorType
          _miNextInsertPoint = Items.Count - 1
        End If
      End If

    End Sub

    Public Overloads Sub InsertBeforePrevious(ByVal lineOfCode As CodeElement)
      Items.Insert(_miNextInsertPoint, lineOfCode)
    End Sub

    Public Overloads Sub AppendAfterNext(ByVal lineOfCode As CodeElement)
      _mbAppendAfterNext = True
      Items.Add(lineOfCode)
    End Sub

    Public Function ToArray() As String()
      Return Items.Select(Function(c) c.Code).ToArray()
    End Function

    Public ReadOnly Property Statement() As String
      Get

        Dim chunk As CodeElement
        Dim iThisElement As Integer
        Dim bComparisonSinceLastLogic As Boolean

        Dim bNeedsIsEqualTo As Boolean
        Dim bNewLine As Boolean

        Statement = String.Empty
        For Each chunk In Items
          bNeedsIsEqualTo = False
          bNewLine = False

          If ReturnType = ComponentValueTypes.Logic Then

            If chunk.CodeType = ComponentTypes.Operator Then

              If chunk.OperatorType = OperatorSubType.Comparison Then
                bComparisonSinceLastLogic = True
                IsLogicBlock = True

              ElseIf chunk.OperatorType = OperatorSubType.Logic Then
                bComparisonSinceLastLogic = False
                IsLogicBlock = True
                bNewLine = True

              ElseIf chunk.OperatorType = OperatorSubType.Modifier Then
                IsLogicBlock = True
              End If

            End If

            ' Is there an operator after this component?
            If Items.Count - 1 > iThisElement Then
              If Items(iThisElement + 1).OperatorType = OperatorSubType.Logic Then
                If Not bComparisonSinceLastLogic Then
                  bNeedsIsEqualTo = True
                  IsLogicBlock = True
                End If
              End If
            Else
              bNeedsIsEqualTo = Not bComparisonSinceLastLogic
            End If

            ' Am I the last component and was there an operator before me?
            If iThisElement = Items.Count - 1 And iThisElement > 0 Then
              If Items(iThisElement - 1).OperatorType = OperatorSubType.Logic Then
                If Not bComparisonSinceLastLogic Then
                  bNeedsIsEqualTo = True
                End If
              End If
            End If

          End If


          ' Does this code element make needing logic safe?
          ' Some logic expressions return a simple logic while others are set specifically 
          ' to logicval = 0, or have the not in front of them!
          If IsLogicBlock And bNeedsIsEqualTo Then
            Statement = String.Format("{0}{1}{2} = 1", Statement, IIf(bNewLine, vbNewLine & New String(CChar(vbTab), CodeLevel), vbNullString), chunk.Code)
            IsLogicBlock = True
            bNeedsIsEqualTo = False

          ElseIf bNeedsIsEqualTo And CaseReturnType = CaseReturnType.Condition Then
            Statement = String.Format("{0}{1}{2} = 1", Statement, IIf(bNewLine, vbNewLine & New String(CChar(vbTab), CodeLevel), vbNullString), chunk.Code)
            IsLogicBlock = True
            bNeedsIsEqualTo = False
          Else
            Statement = String.Format("{0}{1}{2}", Statement, IIf(bNewLine, vbNewLine & New String(CChar(vbTab), CodeLevel), vbNullString), chunk.Code)
          End If

          iThisElement = iThisElement + 1

        Next

        ' Wrap to return code chunks in safety (was there 'not' statement)
        If IsLogicBlock And CaseReturnType = CaseReturnType.Result Then
          Statement = String.Format("CASE WHEN ({0}) THEN 1 ELSE 0 END", Statement)
        End If

      End Get
    End Property

  End Class

End Namespace

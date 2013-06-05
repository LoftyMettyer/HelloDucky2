﻿Imports System.Runtime.InteropServices

Namespace ScriptDB

  <ClassInterface(ClassInterfaceType.None)> _
  Public Class LinesOfCode
    Inherits System.ComponentModel.BindingList(Of ScriptDB.CodeElement)

    Private mbAppendWildcard As Boolean
    Private mlngCaseStatements As Long
    Private mbAppendAftercode As Boolean

    Public CodeLevel As Integer
    Public NestedLevel As Integer
    Public ReturnType As ComponentValueTypes
    Public IsEvaluated As Boolean

    ' Public IsCodeFlow As Boolean = False
    Public IsComparision As Boolean = False

    Public Sub New()
      mbAppendWildcard = False
      mlngCaseStatements = 0
    End Sub

    Public Overloads Sub Add(ByVal LineOfCode As ScriptDB.CodeElement)

      LineOfCode.CaseNumber = mlngCaseStatements

      If mbAppendAftercode Then
        Me.Items.Insert(Me.Items.Count - 1, LineOfCode)
        mbAppendAftercode = False
      Else
        Me.Items.Add(LineOfCode)
      End If

    End Sub

    Public Overloads Sub AddToEnd(ByVal LineOfCode As ScriptDB.CodeElement)

      mbAppendAftercode = True
      Me.Items.Add(LineOfCode)

    End Sub

    ' Property to calculate the character indenation in the code (to beautify the code)
    Public ReadOnly Property Indentation() As String
      Get
        Indentation = Space(8)
      End Get
    End Property


#Region "Code design switches"

    Public Sub AppendWildcard()
      mbAppendWildcard = True
    End Sub

    Public Function ToArray() As String()

      Dim returnArrayList As ArrayList
      Dim objCodeElement As ScriptDB.CodeElement

      returnArrayList = New ArrayList

      For Each objCodeElement In Me.Items
        returnArrayList.Add(objCodeElement.Code)
      Next

      ToArray = returnArrayList.ToArray(GetType(String))

    End Function

    Public Sub SplitIntoCase()

      Dim LineOfCode As ScriptDB.CodeElement

      ' If we're to append a wildcard add it into the current case level
      If mbAppendWildcard = True Then
        LineOfCode.Code = "+ '%'"
        Me.Add(LineOfCode)
      End If

      ' Increment the amount of components in the case statement
      mlngCaseStatements = mlngCaseStatements + 1
      mbAppendWildcard = False

    End Sub

#End Region

    Public ReadOnly Property Statement(ByVal CaseNumber As Long) As String
      Get
        Statement = String.Empty

        Dim Chunk As ScriptDB.CodeElement
        Dim iThisElement As Integer = 0
        Dim bComparisonSinceLastLogic As Boolean = False
        Dim bAddAutoIsEqualTo As Boolean

        Statement = String.Empty
        For Each Chunk In Me.Items
          bAddAutoIsEqualTo = False

          If Chunk.OperatorType = OperatorSubType.Comparison Then Me.IsComparision = True

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
            Statement = vbNewLine & String.Format("{0}{1}{2} = 1", New String(vbTab, CodeLevel), Statement, Chunk.Code)
            bAddAutoIsEqualTo = False
          Else
            Statement = vbNewLine & String.Format("{0}{1}{2}", New String(vbTab, CodeLevel), Statement, Chunk.Code)
          End If

          ' Statement = Statement & Chunk.Code
          iThisElement = iThisElement + 1

        Next

      End Get
    End Property

    Public ReadOnly Property Statement() As String
      Get

        Statement = Statement(0)

        ' If we're to append a wildcard add it into the current case level
        If mbAppendWildcard = True Then
          Statement = Statement & "+ '%'"
        End If

        ' Wrap to return code chunks in safety
        If Me.ReturnType = ComponentValueTypes.Logic Or Me.IsComparision Then
          Statement = vbNewLine & String.Format("{0}CASE WHEN ({1}) THEN 1 ELSE 0 END", New String(vbTab, CodeLevel), Statement)
        End If

        Statement = String.Format("{0}", Statement)

      End Get
    End Property

  End Class

End Namespace

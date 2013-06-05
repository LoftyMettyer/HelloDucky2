﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.239
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Linq
Imports System.Data.Linq.Mapping
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Reflection


<Global.System.Data.Linq.Mapping.DatabaseAttribute(Name:="STD41")>  _
Partial Public Class DBContext
	Inherits System.Data.Linq.DataContext
	
	Private Shared mappingSource As System.Data.Linq.Mapping.MappingSource = New AttributeMappingSource()
	
  #Region "Extensibility Method Definitions"
  Partial Private Sub OnCreated()
  End Sub
  Partial Private Sub InsertASRSysAuditTrail(instance As ASRSysAuditTrail)
    End Sub
  Partial Private Sub UpdateASRSysAuditTrail(instance As ASRSysAuditTrail)
    End Sub
  Partial Private Sub DeleteASRSysAuditTrail(instance As ASRSysAuditTrail)
    End Sub
  #End Region
	
	Public Sub New()
		MyBase.New(Global.Test_Harness.My.MySettings.Default.STD41ConnectionString, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As String)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As System.Data.IDbConnection)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As String, ByVal mappingSource As System.Data.Linq.Mapping.MappingSource)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public Sub New(ByVal connection As System.Data.IDbConnection, ByVal mappingSource As System.Data.Linq.Mapping.MappingSource)
		MyBase.New(connection, mappingSource)
		OnCreated
	End Sub
	
	Public ReadOnly Property ASRSysAuditTrails() As System.Data.Linq.Table(Of ASRSysAuditTrail)
		Get
			Return Me.GetTable(Of ASRSysAuditTrail)
		End Get
	End Property
End Class

<Global.System.Data.Linq.Mapping.TableAttribute(Name:="dbo.ASRSysAuditTrail")>  _
Partial Public Class ASRSysAuditTrail
	Implements System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
	
	Private Shared emptyChangingEventArgs As PropertyChangingEventArgs = New PropertyChangingEventArgs(String.Empty)
	
	Private _id As Integer
	
	Private _UserName As String
	
	Private _DateTimeStamp As Date
	
	Private _RecordID As Integer
	
	Private _RecordDesc As String
	
	Private _OldValue As String
	
	Private _NewValue As String
	
	Private _Tablename As String
	
	Private _Columnname As String
	
	Private _CMGExportDate As System.Nullable(Of Date)
	
	Private _CMGCommitDate As System.Nullable(Of Date)
	
	Private _ColumnID As System.Nullable(Of Integer)
	
	Private _Deleted As System.Nullable(Of Boolean)
	
    #Region "Extensibility Method Definitions"
    Partial Private Sub OnLoaded()
    End Sub
    Partial Private Sub OnValidate(action As System.Data.Linq.ChangeAction)
    End Sub
    Partial Private Sub OnCreated()
    End Sub
    Partial Private Sub OnidChanging(value As Integer)
    End Sub
    Partial Private Sub OnidChanged()
    End Sub
    Partial Private Sub OnUserNameChanging(value As String)
    End Sub
    Partial Private Sub OnUserNameChanged()
    End Sub
    Partial Private Sub OnDateTimeStampChanging(value As Date)
    End Sub
    Partial Private Sub OnDateTimeStampChanged()
    End Sub
    Partial Private Sub OnRecordIDChanging(value As Integer)
    End Sub
    Partial Private Sub OnRecordIDChanged()
    End Sub
    Partial Private Sub OnRecordDescChanging(value As String)
    End Sub
    Partial Private Sub OnRecordDescChanged()
    End Sub
    Partial Private Sub OnOldValueChanging(value As String)
    End Sub
    Partial Private Sub OnOldValueChanged()
    End Sub
    Partial Private Sub OnNewValueChanging(value As String)
    End Sub
    Partial Private Sub OnNewValueChanged()
    End Sub
    Partial Private Sub OnTablenameChanging(value As String)
    End Sub
    Partial Private Sub OnTablenameChanged()
    End Sub
    Partial Private Sub OnColumnnameChanging(value As String)
    End Sub
    Partial Private Sub OnColumnnameChanged()
    End Sub
    Partial Private Sub OnCMGExportDateChanging(value As System.Nullable(Of Date))
    End Sub
    Partial Private Sub OnCMGExportDateChanged()
    End Sub
    Partial Private Sub OnCMGCommitDateChanging(value As System.Nullable(Of Date))
    End Sub
    Partial Private Sub OnCMGCommitDateChanged()
    End Sub
    Partial Private Sub OnColumnIDChanging(value As System.Nullable(Of Integer))
    End Sub
    Partial Private Sub OnColumnIDChanged()
    End Sub
    Partial Private Sub OnDeletedChanging(value As System.Nullable(Of Boolean))
    End Sub
    Partial Private Sub OnDeletedChanged()
    End Sub
    #End Region
	
	Public Sub New()
		MyBase.New
		OnCreated
	End Sub
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_id", AutoSync:=AutoSync.OnInsert, DbType:="Int NOT NULL IDENTITY", IsPrimaryKey:=true, IsDbGenerated:=true)>  _
	Public Property id() As Integer
		Get
			Return Me._id
		End Get
		Set
			If ((Me._id = value)  _
						= false) Then
				Me.OnidChanging(value)
				Me.SendPropertyChanging
				Me._id = value
				Me.SendPropertyChanged("id")
				Me.OnidChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_UserName", DbType:="VarChar(255) NOT NULL", CanBeNull:=false)>  _
	Public Property UserName() As String
		Get
			Return Me._UserName
		End Get
		Set
			If (String.Equals(Me._UserName, value) = false) Then
				Me.OnUserNameChanging(value)
				Me.SendPropertyChanging
				Me._UserName = value
				Me.SendPropertyChanged("UserName")
				Me.OnUserNameChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_DateTimeStamp", DbType:="DateTime NOT NULL")>  _
	Public Property DateTimeStamp() As Date
		Get
			Return Me._DateTimeStamp
		End Get
		Set
			If ((Me._DateTimeStamp = value)  _
						= false) Then
				Me.OnDateTimeStampChanging(value)
				Me.SendPropertyChanging
				Me._DateTimeStamp = value
				Me.SendPropertyChanged("DateTimeStamp")
				Me.OnDateTimeStampChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_RecordID", DbType:="Int NOT NULL")>  _
	Public Property RecordID() As Integer
		Get
			Return Me._RecordID
		End Get
		Set
			If ((Me._RecordID = value)  _
						= false) Then
				Me.OnRecordIDChanging(value)
				Me.SendPropertyChanging
				Me._RecordID = value
				Me.SendPropertyChanged("RecordID")
				Me.OnRecordIDChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_RecordDesc", DbType:="VarChar(255)")>  _
	Public Property RecordDesc() As String
		Get
			Return Me._RecordDesc
		End Get
		Set
			If (String.Equals(Me._RecordDesc, value) = false) Then
				Me.OnRecordDescChanging(value)
				Me.SendPropertyChanging
				Me._RecordDesc = value
				Me.SendPropertyChanged("RecordDesc")
				Me.OnRecordDescChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_OldValue", DbType:="VarChar(MAX)")>  _
	Public Property OldValue() As String
		Get
			Return Me._OldValue
		End Get
		Set
			If (String.Equals(Me._OldValue, value) = false) Then
				Me.OnOldValueChanging(value)
				Me.SendPropertyChanging
				Me._OldValue = value
				Me.SendPropertyChanged("OldValue")
				Me.OnOldValueChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_NewValue", DbType:="VarChar(MAX)")>  _
	Public Property NewValue() As String
		Get
			Return Me._NewValue
		End Get
		Set
			If (String.Equals(Me._NewValue, value) = false) Then
				Me.OnNewValueChanging(value)
				Me.SendPropertyChanging
				Me._NewValue = value
				Me.SendPropertyChanged("NewValue")
				Me.OnNewValueChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Tablename", DbType:="VarChar(200)")>  _
	Public Property Tablename() As String
		Get
			Return Me._Tablename
		End Get
		Set
			If (String.Equals(Me._Tablename, value) = false) Then
				Me.OnTablenameChanging(value)
				Me.SendPropertyChanging
				Me._Tablename = value
				Me.SendPropertyChanged("Tablename")
				Me.OnTablenameChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Columnname", DbType:="VarChar(200)")>  _
	Public Property Columnname() As String
		Get
			Return Me._Columnname
		End Get
		Set
			If (String.Equals(Me._Columnname, value) = false) Then
				Me.OnColumnnameChanging(value)
				Me.SendPropertyChanging
				Me._Columnname = value
				Me.SendPropertyChanged("Columnname")
				Me.OnColumnnameChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_CMGExportDate", DbType:="DateTime")>  _
	Public Property CMGExportDate() As System.Nullable(Of Date)
		Get
			Return Me._CMGExportDate
		End Get
		Set
			If (Me._CMGExportDate.Equals(value) = false) Then
				Me.OnCMGExportDateChanging(value)
				Me.SendPropertyChanging
				Me._CMGExportDate = value
				Me.SendPropertyChanged("CMGExportDate")
				Me.OnCMGExportDateChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_CMGCommitDate", DbType:="DateTime")>  _
	Public Property CMGCommitDate() As System.Nullable(Of Date)
		Get
			Return Me._CMGCommitDate
		End Get
		Set
			If (Me._CMGCommitDate.Equals(value) = false) Then
				Me.OnCMGCommitDateChanging(value)
				Me.SendPropertyChanging
				Me._CMGCommitDate = value
				Me.SendPropertyChanged("CMGCommitDate")
				Me.OnCMGCommitDateChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_ColumnID", DbType:="Int")>  _
	Public Property ColumnID() As System.Nullable(Of Integer)
		Get
			Return Me._ColumnID
		End Get
		Set
			If (Me._ColumnID.Equals(value) = false) Then
				Me.OnColumnIDChanging(value)
				Me.SendPropertyChanging
				Me._ColumnID = value
				Me.SendPropertyChanged("ColumnID")
				Me.OnColumnIDChanged
			End If
		End Set
	End Property
	
	<Global.System.Data.Linq.Mapping.ColumnAttribute(Storage:="_Deleted", DbType:="Bit")>  _
	Public Property Deleted() As System.Nullable(Of Boolean)
		Get
			Return Me._Deleted
		End Get
		Set
			If (Me._Deleted.Equals(value) = false) Then
				Me.OnDeletedChanging(value)
				Me.SendPropertyChanging
				Me._Deleted = value
				Me.SendPropertyChanged("Deleted")
				Me.OnDeletedChanged
			End If
		End Set
	End Property
	
	Public Event PropertyChanging As PropertyChangingEventHandler Implements System.ComponentModel.INotifyPropertyChanging.PropertyChanging
	
	Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
	
	Protected Overridable Sub SendPropertyChanging()
		If ((Me.PropertyChangingEvent Is Nothing)  _
					= false) Then
			RaiseEvent PropertyChanging(Me, emptyChangingEventArgs)
		End If
	End Sub
	
	Protected Overridable Sub SendPropertyChanged(ByVal propertyName As [String])
		If ((Me.PropertyChangedEvent Is Nothing)  _
					= false) Then
			RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
		End If
	End Sub
End Class

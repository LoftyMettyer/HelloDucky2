CREATE PROCEDURE [dbo].[spASRIntGetAllrequiredOrganisationColumns] 
(
	@piOrganisationID				integer,
	@psOrganisationReportType	varchar(50)
)
AS
BEGIN

	DECLARE @iBaseViewID int = (select BaseViewID from ASRSysOrganisationReport where ID = @piOrganisationID);
	
	IF @piOrganisationID = 0
	BEGIN
			IF @psOrganisationReportType = 'COMMERCIAL'
				SELECT c.columnID, c.ColumnName
				FROM ASRSysModuleSetup As s 
				INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				WHERE s.moduleKey = 'MODULE_PERSONNEL' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDSEMPLOYEENUMBER'
				UNION
				SELECT c.columnID, c.ColumnName
				FROM ASRSysModuleSetup As s 
				INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				WHERE s.moduleKey = 'MODULE_PERSONNEL' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDSMANAGERSTAFFNO'
				UNION
				SELECT   c.columnID, c.ColumnName
				FROM ASRSysModuleSetup As s 
				INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				WHERE s.moduleKey = 'MODULE_PERSONNEL' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDSLEAVINGDATE'
				UNION 
				SELECT c.columnID, c.ColumnName
				FROM ASRSysModuleSetup As s 
				INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				WHERE s.moduleKey = 'MODULE_PERSONNEL' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDSSTARTDATE'
			ELSE
				--Non Commercial	
				SELECT c.columnID, c.ColumnName
				FROM ASRSysModuleSetup As s 
				INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				WHERE s.moduleKey = 'MODULE_HIERARCHY' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDIDENTIFIER'
				UNION
				SELECT c.columnID, c.ColumnName
				FROM ASRSysModuleSetup As s 
				INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				WHERE s.moduleKey = 'MODULE_HIERARCHY' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDREPORTSTO'	
			END
	ELSE
	BEGIN
	IF @psOrganisationReportType = 'COMMERCIAL'
		BEGIN
			SELECT c.columnID, c.ColumnName, v1.ViewName As TableOrViewName
			FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
			INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
			AND UPPER(s.ParameterKey) = 'PARAM_FIELDSEMPLOYEENUMBER'
			UNION
			SELECT c.columnID, c.ColumnName, v1.ViewName As TableOrViewName
			FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
			INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
			AND UPPER(s.ParameterKey) = 'PARAM_FIELDSMANAGERSTAFFNO'
			UNION
			SELECT   c.columnID, c.ColumnName, v1.ViewName As TableOrViewName	
			FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
			INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
			AND UPPER(s.ParameterKey) = 'PARAM_FIELDSLEAVINGDATE'
			UNION 
			SELECT c.columnID, c.ColumnName,v1.ViewName As TableOrViewName		
			FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
			INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE s.moduleKey = 'MODULE_PERSONNEL' 
			AND UPPER(s.ParameterKey) = 'PARAM_FIELDSSTARTDATE'
			UNION
			SELECT c.columnID, c.ColumnName,v1.ViewName As TableOrViewName		
			FROM ASRSysOrganisationReportFilters As f INNER JOIN ASRSysColumns As c ON f.FieldID = c.columnID 
			INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE f.OrganisationID = @piOrganisationID
			UNION
			SELECT oc.ColumnID
					,c.ColumnName				   
					,CASE 
						WHEN v.ViewName IS NULL THEN  t.TableName
						WHEN v.ViewName IS NOT NULL THEN  v.ViewName
					END As TableOrViewName
			FROM  ASRSysOrganisationColumns As oc
			INNER JOIN ASRSysColumns As c ON oc.ColumnID = c.columnId		
			INNER JOIN ASRSysTables As t ON c.tableID = t.tableID		
			LEFT JOIN ASRSysViews As v ON oc.ViewID = v.ViewID		   
			WHERE oc.OrganisationID = @piOrganisationID;
		END
		ELSE
		BEGIN
				--Non Commercial	
				SELECT c.columnID, c.ColumnName, v1.ViewName As TableOrViewName
				FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE s.moduleKey = 'MODULE_HIERARCHY' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDIDENTIFIER'
				UNION
				SELECT c.columnID, c.ColumnName, v1.ViewName As TableOrViewName
				FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID 
				INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE s.moduleKey = 'MODULE_HIERARCHY' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDREPORTSTO'
				UNION
				SELECT c.columnID, c.ColumnName,t.TableName As TableOrViewName				  
				FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID
				INNER JOIN ASRSysTables As t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY'
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDSTARTDATE' 
				UNION
				SELECT    c.columnID, c.ColumnName,t.TableName As TableOrViewName				      
				FROM ASRSysModuleSetup As s INNER JOIN ASRSysColumns As c ON s.ParameterValue = c.columnID
				INNER JOIN ASRSysTables As t ON c.tableID = t.tableID 	WHERE s.moduleKey = 'MODULE_HIERARCHY' 
				AND UPPER(s.ParameterKey) = 'PARAM_FIELDENDDATE'
				UNION
				SELECT c.columnID, c.ColumnName,v1.ViewName As TableOrViewName		
				FROM ASRSysOrganisationReportFilters As f INNER JOIN ASRSysColumns As c ON f.FieldID = c.columnID 
				INNER JOIN ASRSysViews As v1 ON v1.ViewID = @iBaseViewID WHERE f.OrganisationID = @piOrganisationID
				UNION
				SELECT oc.ColumnID
						,c.ColumnName				   
						,CASE 
							WHEN v.ViewName IS NULL THEN  t.TableName
							WHEN v.ViewName IS NOT NULL THEN  v.ViewName
						END As TableOrViewName
				FROM  ASRSysOrganisationColumns As oc
				INNER JOIN ASRSysColumns As c ON oc.ColumnID = c.columnId		
				INNER JOIN ASRSysTables As t ON c.tableID = t.tableID		
				LEFT JOIN ASRSysViews As v ON oc.ViewID = v.ViewID		   
				WHERE oc.OrganisationID = @piOrganisationID;
		END

		-- Update the utility access log.
		UPDATE ASRSysUtilAccessLog SET 
		RunBy = system_user, 
		RunDate = getdate(), 
		RunHost = host_name() 
		WHERE UtilID = @piOrganisationID AND Type = 39;
	END
END

CREATE PROCEDURE [dbo].[spASRIntGetOrganisationReport] (	
	 @piReportID				integer
	,@piRootID					integer
	,@psOrganisationReportType	varchar(50)
   ,@psPostAllocationViewName	varchar(500)
)		
AS		
BEGIN		
	SET NOCOUNT ON;

   IF @psOrganisationReportType = 'COMMERCIAL'
      EXECUTE dbo.spASRIntGetOrganisationReport_Commercial @piReportID, @piRootID;
   ELSE
      EXECUTE dbo.spASRIntGetOrganisationReport_Post @piReportID, @piRootID, @psOrganisationReportType, @psPostAllocationViewName;

	-- Return Result dataset for respected organisationReport column's parameters like prefix,suffix etc.
	SELECT oc.ColumnID, c.ColumnName, oc.Prefix, oc.Suffix
			,oc.FontSize, oc.Height, oc.Decimals, oc.ConcatenateWithNext
			,t.TableID, t.TableName, ISNULL(v.ViewID,0) AS ViewID, ISNULL(v.ViewName,'') AS ViewName, c.datatype
	FROM  ASRSysOrganisationColumns oc
	   INNER JOIN ASRSysColumns c ON oc.ColumnID = c.columnId		
	   INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
	   LEFT JOIN ASRSysViews v ON oc.ViewID = v.ViewID
	WHERE oc.OrganisationID = @piReportID;      
   
   SELECT Name As DefinitionName FROM ASRSysOrganisationReport WHERE ID = @piReportID
    
   UPDATE ASRSysUtilAccessLog SET 
            RunBy = system_user, 
            RunDate = getdate(), 
            RunHost = host_name() 
   WHERE UtilID = @piReportID AND Type = 39;

END
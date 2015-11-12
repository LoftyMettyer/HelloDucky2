CREATE PROCEDURE [dbo].[spsys_saveobjectcategories](@utilityType as integer, @UtilityID as integer, @CategoryID integer)
		AS
		BEGIN

			SET NOCOUNT ON;

			DELETE FROM dbo.tbsys_objectcategories
				WHERE [objecttype] = @utilityType AND [objectid] = @UtilityID;
			
			IF @CategoryID > 0
			BEGIN
				INSERT dbo.tbsys_objectcategories([objecttype], [objectid], [categoryid])
					VALUES (@utilityType, @UtilityID, @CategoryID);
			END

		END
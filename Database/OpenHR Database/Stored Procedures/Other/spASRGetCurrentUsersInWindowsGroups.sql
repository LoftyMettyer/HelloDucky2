CREATE PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups]
(
	@psGroupNames VARCHAR(MAX)
)
AS
BEGIN
	SET NOCOUNT ON;

	CREATE TABLE #tblCurrentUsers				
		(
			hostname varchar(256)
			,loginame varchar(256)
			,program_name varchar(256)
			,hostprocess varchar(20)
			,sid binary(86)
			,login_time datetime
			,spid int
			,uid int
		)
		
	INSERT INTO #tblCurrentUsers
		EXEC spASRGetCurrentUsers

	CREATE TABLE #tblUsersInGroup
		(
		loginame varchar(256)
		,groupname varchar(256)
		)		

	DECLARE @tblGroups TABLE
		(
		groupname varchar(256)
		)

	DECLARE @iUserInGroup	integer,
			@loginame		varchar(256),
			@IN				varchar(MAX), 
			@INGroup		varchar(MAX),
			@Pos			int;

	SET @psGroupNames = LTRIM(RTRIM(@psGroupNames))+ ','
	SET @Pos = CHARINDEX(',', @psGroupNames, 1)
	SET @IN = ''

	IF REPLACE(@psGroupNames, ',', '') <> ''
	BEGIN
		WHILE @Pos > 0
		BEGIN
			SET @INGroup = LTRIM(RTRIM(LEFT(@psGroupNames, @Pos - 1)))
			IF @INGroup <> ''
			BEGIN
				INSERT INTO @tblGroups VALUES (@INGroup)
			END
			SET @psGroupNames = RIGHT(@psGroupNames, LEN(@psGroupNames) - @Pos)
			SET @Pos = CHARINDEX(',', @psGroupNames, 1)
		END
	END

	DECLARE CurrentUsersCursor CURSOR LOCAL FAST_FORWARD READ_ONLY FOR 
	SELECT loginame FROM #tblCurrentUsers
	OPEN CurrentUsersCursor
	FETCH NEXT FROM CurrentUsersCursor INTO @loginame

	IF @@FETCH_STATUS <> 0 
		RETURN

	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @loginame = LTRIM(RTRIM(@loginame))

		INSERT INTO #tblUsersInGroup 
			EXEC spASRGroupsUserIsMemberOf @loginame

		FETCH NEXT FROM CurrentUsersCursor INTO @loginame
	END
	CLOSE CurrentUsersCursor
	DEALLOCATE CurrentUsersCursor

	DROP TABLE #tblCurrentUsers

	/* Return a recordset of the users to log out */
	SELECT loginame
	FROM #tblUsersInGroup 
	WHERE groupname IN (SELECT DISTINCT groupname collate database_default FROM @tblGroups)

END

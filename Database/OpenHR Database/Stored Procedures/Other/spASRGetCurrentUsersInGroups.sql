CREATE PROCEDURE spASRGetCurrentUsersInGroups 
(
	@psGroupNames VARCHAR(MAX)
)
AS
BEGIN
	SET NOCOUNT ON

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

	DECLARE @tblGroups TABLE
	(
		groupname varchar(256) collate database_default 
	)

	DECLARE @IN			varchar(MAX), 
			@INGroup	varchar(MAX),
			@Pos		integer;

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

	SELECT [#tblCurrentUsers].[loginame]
	FROM [#tblCurrentUsers] 
		JOIN [ASRSysUserGroups] ON [#tblCurrentUsers].[loginame] = [ASRSysUserGroups].[UserName] collate database_default 
	WHERE [ASRSysUserGroups].[UserGroup] IN 
		(
			SELECT [groupName] FROM @tblGroups
		)


END
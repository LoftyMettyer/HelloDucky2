
  DECLARE @JobID BINARY(16)  
  DECLARE @ReturnCode INT    
  DECLARE @sJobName nvarchar(4000)
  DECLARE @sDBName nvarchar(4000)
  DECLARE @sErrMsg nvarchar(4000)

  DECLARE @SchedType varchar(8000)
  DECLARE @SchedInterval varchar(8000)
  DECLARE @SchedTime varchar(8000)

  DECLARE @numSQLVersion numeric(3,1)
  

  SET @SchedType = 4 --default to daily
  SELECT @SchedType = settingvalue
  FROM asrsyssystemsettings
  WHERE [section] = 'overnight' and [settingkey] = 'type'

  SET @SchedInterval = 1 --default to everyday
  SELECT @SchedInterval = settingvalue
  FROM asrsyssystemsettings
  WHERE [section] = 'overnight' and [settingkey] = 'interval'

  SET @SchedTime = 30000 --default to 3:00am
  SELECT @SchedTime = settingvalue
  FROM asrsyssystemsettings
  WHERE [section] = 'overnight' and [settingkey] = 'time'

  SELECT @sDBName = master..sysdatabases.name 
  FROM master..sysdatabases
  INNER JOIN master..sysprocesses ON master..sysdatabases.dbid = master..sysprocesses.dbid
  WHERE master..sysprocesses.spid = @@spid

  SELECT @numSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));

  DECLARE @NVarCommand nvarchar(4000)

  IF @numSQLVersion < 11
  BEGIN
	  SELECT @NVarCommand = 'USE master
		GRANT EXECUTE ON xp_StartMail TO public
		GRANT EXECUTE ON xp_SendMail TO public'
	  EXEC sp_executesql @NVarCommand;
  END

  SELECT @NVarCommand = 'USE master
    GRANT EXECUTE ON sp_OACreate TO public
    GRANT EXECUTE ON sp_OADestroy TO public
    GRANT EXECUTE ON sp_OAGetErrorInfo TO public
    GRANT EXECUTE ON sp_OAGetProperty TO public
    GRANT EXECUTE ON sp_OAMethod TO public
    GRANT EXECUTE ON sp_OASetProperty TO public
    GRANT EXECUTE ON sp_OAStop TO public
    GRANT EXECUTE ON xp_LoginConfig TO public
    GRANT EXECUTE ON xp_EnumGroups TO public'
  EXEC sp_executesql @NVarCommand;
  
  SELECT @NVarCommand = 'USE ['+@sDBName + ']';
  EXEC sp_executesql @NVarCommand;


BEGIN TRANSACTION            

  SET @sJobName = 'job_HRProOvernightProcessing_' + @sDBName

  SELECT @ReturnCode = 0     
  IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 
    EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'

  -- Delete the job with the same name (if it exists)
  SELECT @JobID = job_id     
  FROM   msdb.dbo.sysjobs    
  WHERE (name = @sJobName)       

  IF (@JobID IS NOT NULL)    
  BEGIN
  -- Check if the job is a multi-server job  
  IF (EXISTS (SELECT  * 
              FROM    msdb.dbo.sysjobservers 
              WHERE   (job_id = @JobID) AND (server_id <> 0))) 
  BEGIN 
    -- There is, so abort the script 
    SET @sErrMsg = 'Unable to import job ''' + @sJobName + ' since there is already a multi-server job with this name.'
    RAISERROR (@sErrMsg, 16, 1) 
    GOTO QuitWithRollback  
  END 
  ELSE 
    -- Delete the [local] job 
    EXECUTE msdb.dbo.sp_delete_job @job_name = @sJobName
    SELECT @JobID = NULL
  END

BEGIN 

  -- Add the job
  EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = @sJobName, @owner_login_name = N'sa', @description = N'HR Pro - automatically created job for running overnight processes.', @category_name = N'[Uncategorized (Local)]', @enabled = 1, @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  -- Add the job steps
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Step 1', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep1'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep1 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 2, @step_name = N'Step 2', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep2'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep2 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 3, @step_name = N'Step 3', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep3'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep3 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 4, @step_name = N'Step 4', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep4'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep4 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 5, @step_name = N'Step 5', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep5'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep5 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 6, @step_name = N'Step 6', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep6'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep6 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 7, @step_name = N'Step 7', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep7'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep7 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 8, @step_name = N'Step 8', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep8'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep8 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 9, @step_name = N'Step 9', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep9'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep9 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 10, @step_name = N'Step 10', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASRSysOvernightStep10'') AND sysstat & 0xf = 4) BEGIN EXEC spASRSysOvernightStep10 END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 0, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 

  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  -- Add the job schedules
  --EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'HR Pro - time-dependent columns schedule', @enabled = 1, @freq_type = 4, @active_start_date = 20020429, @active_start_time = 30000, @freq_interval = 1, @freq_subday_type = 1, @freq_subday_interval = 0, @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = 99991231, @active_end_time = 235959
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'HR Pro - time-dependent columns schedule', 
    @enabled = 1, @freq_type = @SchedType, @active_start_date = 20020101, @active_start_time = @SchedTime,
    @freq_interval = @SchedInterval, @freq_subday_type = 1, @freq_subday_interval = 0, 
    @freq_relative_interval = 0, @freq_recurrence_factor = 1, @active_end_date = 99991231,
    @active_end_time = 235959
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  -- Add the Target Servers
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' 
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

END

COMMIT TRANSACTION

BEGIN TRANSACTION

  SET @sJobName = 'job_HRProOutlookBatch_' + @sDBName

  SELECT @ReturnCode = 0     
  IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 
	EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'

  -- Delete the job with the same name (if it exists)
  SET @JobID = NULL
  SELECT @JobID = job_id
  FROM   msdb.dbo.sysjobs
  WHERE (name = @sJobName)

  IF (@JobID IS NOT NULL)    
  BEGIN  
  -- Check if the job is a multi-server job  
  IF (EXISTS (SELECT  * 
			  FROM    msdb.dbo.sysjobservers 
			  WHERE   (job_id = @JobID) AND (server_id <> 0))) 
  BEGIN 
	-- There is, so abort the script 
	SET @sErrMsg = 'Unable to import job ''' + @sJobName + ' since there is already a multi-server job with this name.'
	RAISERROR (@sErrMsg, 16, 1) 
	GOTO QuitWithRollback  
  END 
  ELSE 
	-- Delete the [local] job 
	EXECUTE msdb.dbo.sp_delete_job @job_name = @sJobName
	SELECT @JobID = NULL
  END 

COMMIT TRANSACTION

-- Outlook Batch replaced by HR Pro Outlook Calendar - Windows Service
--
--SELECT @sSQLVersion = substring(@@version,charindex('-',@@version)+2,1)
--IF (@sSQLVersion >= '9')
--BEGIN
--	BEGIN TRANSACTION
--
--	  -- Add the job
--	  EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = @sJobName, @owner_login_name = N'sa', @description = N'HR Pro - automatically created job for outlook calendar processes.', @category_name = N'[Uncategorized (Local)]', @enabled = 1, @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0
--	  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
--
--	  -- Add the job steps
--	  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Step 1', @command = N'IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(''spASROutlookBatch'') AND sysstat & 0xf = 4) BEGIN EXEC spASROutlookBatch END', @database_name = @sDBName, @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2
--	  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
--	  EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 
--
--	  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
--
--	  -- Add the job schedules
--	  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'HR Pro Outlook Interface', 
--		@enabled = 1, @freq_type = 8, @active_start_date = 20020101, @active_start_time = 0, 
--		@freq_interval = 127, @freq_subday_type = 4, @freq_subday_interval = 1, 
--		@freq_relative_interval = 0, @freq_recurrence_factor = 1, @active_end_date = 99991231, 
--		@active_end_time = 235959
--	  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
--
--	  -- Add the Target Servers
--	  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' 
--	  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
--
--	COMMIT TRANSACTION
--END

GOTO   EndSave              
QuitWithRollback:
  IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION 
EndSave:

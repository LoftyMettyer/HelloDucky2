-- Add your test scenario here --


select [dbo].[udfstat_getuniquecode] ('APPNO', 1.0, 135)

exec dbo.[spstat_flushuniquecode]

--exec spASRStoredDataFileActions 1840, 1289, 21

/*
select * from Table1
UPDATE Table1 SET column1 = dbo.udfASRNetgetUniqueCode('STAFFNO')
--EXEC spadmin_CommitUniqueCode 'STAFFNO'
select * from table1
*/

--update table1 set column1 = 'c'
--select * from tbsys_uniquecodes

declare @value varchar(max)
declare @nvalue nvarchar(max)

--set @value = dbo.udfASRNetGetDomainLogins('coa.local')

--declare @i int
--set @i = 0
--while @i < 50
--begin
	--exec spASRGetWindowsUsersFromAssembly 'coa.local'

	--exec spASRGetWindowsGroupsFromAssembly 'coa.local'

--	set @i = @i + 1
--end



--select len(@value)
--select substring(@value, len(@value) - 10, 20)

--set @nvalue = dbo.udfASRNetGetDomainLogins('coa.local')

--select len(@nvalue)
--select substring(@nvalue, len(@nvalue) - 10, 20)


--EXEC spASRGetWindowsUsersFromAssembly 'coa.local'

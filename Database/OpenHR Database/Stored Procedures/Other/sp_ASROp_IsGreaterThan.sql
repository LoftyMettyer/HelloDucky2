CREATE PROCEDURE [dbo].[sp_ASROp_IsGreaterThan]
(
	@date1			datetime,
	@date2			datetime,
	@retdate   		bit OUTPUT,
	@char1			varchar(MAX),
	@char2			varchar(MAX),
	@retchar   		bit OUTPUT,
	@numeric1		numeric,
	@numeric2		numeric,
	@retnumeric   	bit OUTPUT,
	@logic1			bit,
	@logic2			bit,
	@retlogic		bit OUTPUT
)

AS
BEGIN
	if @date1 is not null
	begin
		if @date1 > @date2
		begin
		set @retdate = 1
		select @retdate as result
		end
		if @date1 <= @date2
		begin
		set @retdate = 0
		select @retdate as result
		end	
	end

	if @char1 is not null
	begin
		if @char1 > @char2
		begin
		set @retchar = 1
		select @retchar as result
		end
		if @char1 <= @char2
		begin
		set @retchar = 0
		select @retchar as result
		end	
	end

	if @numeric1 is not null
	begin
		if @numeric1 > @numeric2
		begin
		set @retnumeric = 1
		select @retnumeric as result
		end
		if @numeric1 <= @numeric2
		begin
		set @retnumeric = 0
		select @retnumeric as result
		end	
	end

	if @logic1 is not null
	begin
		if @logic1 > @logic2
		begin
		set @retlogic = 1
		select @retlogic as result
		end
		if @logic1 <= @logic2
		begin
		set @retlogic = 0
		select @retlogic as result
		end	
	end
END

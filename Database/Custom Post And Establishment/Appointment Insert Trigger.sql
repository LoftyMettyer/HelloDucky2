
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Appointments_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Appointments_P&E]
GO

CREATE TRIGGER [dbo].[trcustom_Appointments_P&E] ON [dbo].[tbuser_Appointments]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	INSERT Appointment_Allowances(ID_3, Effective_Date, Type, Frequency, Amount, Currency)
		SELECT i.ID, i.Appointment_Start_Date, pa.Type, pa.Frequency, pa.Amount, pa.Currency
			FROM inserted i
			INNER JOIN Post_Allowances pa ON pa.ID_219 = i.ID_219;

	INSERT Appointment_Benefits(ID_3, Effective_Date, Type, Provider, Frequency, Cost, Currency, Annual_Cost_to_Employer)
		SELECT i.ID, i.Appointment_Start_Date, pb.Type, pb.Provider, pb.Frequency, pb.Cost, pb.Currency, pb.Annual_Cost_to_Employer
			FROM inserted i
			INNER JOIN Post_Benefits pb ON pb.ID_219 = i.ID_219;

	INSERT Appointment_Deductions(ID_3, Effective_Date, Type, Frequency, Amount, Currency)
		SELECT i.ID, i.Appointment_Start_Date, pd.Type, pd.Frequency, pd.Amount, pd.Currency
			FROM inserted i
			INNER JOIN Post_Deductions pd ON pd.ID_219 = i.ID_219;

	INSERT Appointment_Holiday_Schemes(ID_3, Effective_Date, Holiday_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, phs.Holiday_Scheme
			FROM inserted i
			INNER JOIN Post_Holiday_Schemes phs ON phs.ID_219 = i.ID_219;

	INSERT Appointment_OMP_Schemes(ID_3, Effective_Date, OMP_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, sch.OMP_Scheme
			FROM inserted i
			INNER JOIN Post_OMP_Schemes sch ON sch.ID_219 = i.ID_219;

	INSERT Appointment_OSP_Schemes(ID_3, Effective_Date, OSP_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, sch.OSP_Scheme
			FROM inserted i
			INNER JOIN Post_OSP_Schemes sch ON sch.ID_219 = i.ID_219;

	INSERT Appointment_Pension_Schemes(ID_3, Effective_Date, Pension_Scheme)
		SELECT i.ID, i.Appointment_Start_Date, sch.Pension_Scheme
			FROM inserted i
			INNER JOIN Post_Pension_Schemes sch ON sch.ID_219 = i.ID_219;

	INSERT Appointment_Working_Patterns(ID_3, Effective_Date, Regional_ID, Absence_In, Day_Pattern
									, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM
									, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM
									, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM)
		SELECT i.ID, i.Appointment_Start_Date, wp.Regional_ID, wp.Absence_In, wp.Day_Pattern, wp.Sunday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_AM,wp.Monday_Hours_PM
									, wp.Tuesday_Hours_AM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_AM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_AM, wp.Thursday_Hours_PM
									, wp.Friday_Hours_AM, wp.Friday_Hours_PM, wp.Saturday_Hours_AM, wp.Saturday_Hours_PM
			FROM inserted i
			INNER JOIN Post_Working_Patterns wp ON wp.ID_219 = i.ID_219;



END
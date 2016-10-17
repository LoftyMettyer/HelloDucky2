CREATE TYPE OrgChartRelation AS TABLE 
		( IsGhostNode bit
			, ManagerRoot int
			, HierarchyLevel int
			, EmployeeID int
			, Staff_Number varchar(255)
			, Reports_To_Staff_Number varchar(255));
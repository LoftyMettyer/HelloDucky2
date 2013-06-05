--SET IDENTITY_INSERT FUSION.[MessageElements] ON
RETURN

--select * from fusion.[MessageElements]

DELETE FROM fusion.[MessageElements]
DELETE FROM fusion.[Message]


DBCC CHECKIDENT ('fusion.[MessageElements]', RESEED, 0)


-- StaffChange
	INSERT fusion.[Message] (ID, Name, Description, [Schema], Skeleton, Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation)
		VALUES (1, 'StaffChange', ' Change of staff details', 0x, 1, 1, 1, 1, 1, 1, 1, 1)

	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (1, 2, 'forenames', 1, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'surname', 2, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'preferredName', 3, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'payrollNumber', 4, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'DOB', 5, 11, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'employeeType', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'employmentStatus', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'homePhoneNumber', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'workMobile', 5, 12, 1, 1, 1, 20, 50, 0)

	INSERT fusion.[MessageElements] (ID, MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
		VALUES (2, 1, 2, 'personalMobile', 5, 12, 1, 1, 1, 20, 50, 0)


--select * from fusion.[MessageElements]


--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (1, 1, 'forenames', 0, 1, 1, 3, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (2, 1, 'surname', 0, 1, 1, 2, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (3, 1, 'preferredName', 1, 0, 1, 20, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (4, 1, 'payrollNumber', 0, 0, 1, 2164, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (5, 1, 'DOB', 1, 0, 1, 12, 11)

--		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'employeeType', 1, 0, 1, 'Employee')
--		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'employmentStatus', 1, 0, 1, 'Active')

--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (8, 1, 'homePhoneNumber', 1, 0, 1, 29, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (9, 1, 'workMobile', 1, 0, 1, 1888, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (10, 1, 'personalMobile', 1, 0, 1, 1887, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (11, 1, 'email', 1, 0, 1, 531, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (12, 1, 'personalEmail', 1, 0, 1, 30, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (13, 1, 'addressLine1', 0, 1, 1, 23, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (14, 1, 'addressLine2', 0, 1, 1, 24, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (15, 1, 'addressLine3', 0, 1, 1, 25, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (16, 1, 'addressLine4', 0, 1, 1, 26, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (17, 1, 'addressLine5', 0, 1, 1, 27, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (18, 1, 'postCode', 0, 1, 1, 28, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (19, 1, 'gender', 0, 1, 1, 18, 12)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (20, 1, 'startDate', 0, 1, 1, 14, 11)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (21, 1, 'leavingDate', 1, 0, 1, 15, 11)
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (22, 1, 'leavingReason', 1, 0, 1, 17, 12)
----		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'companyName', 1, 0, 1, 'UNMAPPED FIELD')
--		INSERT @xmldef (position, xmlmessageid, xmlnodekey, nilable, minoccurs, tableid, columnid, datatype) VALUES (24, 1, 'jobTitle', 0, 0, 1, 109, 12)
----		INSERT @xmldef (xmlmessageid, xmlnodekey, nilable, minoccurs, datatype, value) VALUES (1, 'managerRef', 1, 0, 1, 'UNMAPPED FIELD')



select * from 	 fusion.[Message] e
select * from 	 fusion.[MessageElements] e

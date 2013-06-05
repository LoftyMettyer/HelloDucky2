		DECLARE @xmldef TABLE (xmlmessageid smallint, xmlnodekey varchar(255), position tinyint
			, nilable bit, minoccurs bit
			, tableid integer, columnid integer
			, datatype tinyint, minsize integer, maxsize integer, value nvarchar(255))

			select * from  @xmldef


			--select * from fusion.[Message]
			--select * from fusion.[MessageElements]


SELECT m.name AS xmlmessageID,
		me.NodeKey AS xmlnodekey,
		me.Position,
		me.Nillable AS nilable,
		me.minOccurs,
		me.maxOccurs,
		c.TableID,
		e.ColumnID,
		me.MinSize,
		me.MaxSize,
		'' AS value
		FROM fusion.[MessageElements] me
			INNER JOIN fusion.Message m ON m.ID = me.MessageID
			INNER JOIN fusion.Element e ON e.ID = me.ElementID
			INNER JOIN fusion.Category c ON c.ID = e.categoryID


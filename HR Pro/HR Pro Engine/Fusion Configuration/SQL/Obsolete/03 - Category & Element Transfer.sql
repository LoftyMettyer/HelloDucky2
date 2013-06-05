
RETURN

UPDATE fusion.MessageElements SET ElementID = NULL
DELETE FROM fusion.Element
DELETE FROM fusion.Category
	
INSERT INTO fusion.Category (ID, Name, TableID) 
SELECT at.TransferTypeID, at.TransferType, t.TableID
FROM dbo.ASRSysAccordTransferTypes at
LEFT JOIN dbo.tbsys_tables t ON t.TableID = at.ASRBaseTableID
WHERE at.IsVisible = 1

INSERT INTO fusion.Element (ID, CategoryID, Name, DataType, MinSize, MaxSize, ColumnID, Lookup)
SELECT ROW_NUMBER() OVER (ORDER BY fd.TransferTypeID, fd.TransferFieldID), fd.TransferTypeID, fd.Description, CASE WHEN fd.ConvertData = 1 THEN -1 ELSE ISNULL(col.DataType, -1) END, NULL, NULL, CASE WHEN fd.ConvertData = 1 THEN NULL ELSE col.ColumnID END, 0
FROM dbo.ASRSysAccordTransferFieldDefinitions fd
INNER JOIN dbo.ASRSysAccordTransferTypes tt ON tt.TransferTypeID = fd.TransferTypeID
LEFT JOIN fusion.Category c ON c.ID = fd.TransferTypeID
LEFT JOIN dbo.tbsys_columns col ON col.ColumnID = fd.ASRColumnID
WHERE c.ID IS NOT NULL AND (fd.ASRTableID IS NULL OR fd.ASRTableID = tt.ASRBaseTableID)
ORDER BY TransferTypeID, TransferFieldID


-- Populate message definitions

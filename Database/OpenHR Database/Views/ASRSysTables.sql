CREATE VIEW [dbo].[ASRSysTables]
					WITH SCHEMABINDING
					AS SELECT base.[tableid] AS [tableId], base.[tabletype], base.[defaultorderid], base.[recorddescexprid], base.[defaultemailid], base.[tablename], base.[manualsummarycolumnbreaks], base.[auditinsert], base.[auditdelete]
							, base.[isremoteview], base.[inserttriggerdisabled], base.[updatetriggerdisabled], base.[deletetriggerdisabled], base.[CopyWhenParentRecordIsCopied]
							, obj.[locked], obj.[lastupdated], obj.[lastupdatedby]
						FROM dbo.[tbsys_tables] base
						INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.tableid AND obj.objecttype = 1
						INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
						WHERE obj.effectivedate <= dt.[date]
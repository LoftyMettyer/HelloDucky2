CREATE VIEW [dbo].[ASRSysWorkflows]
					WITH SCHEMABINDING
					AS SELECT base.[id], base.[name], base.[description], base.[enabled], base.[initiationtype], base.[basetable], base.[querystring], base.[pictureid], obj.[locked], obj.[lastupdated], obj.[lastupdatedby]
						FROM dbo.[tbsys_workflows] base
						INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.id AND obj.objecttype = 10
						INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
						WHERE obj.effectivedate <= dt.[date]
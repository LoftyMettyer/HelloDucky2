CREATE VIEW [dbo].[ASRSysViews]
					WITH SCHEMABINDING
					AS SELECT base.[viewid], base.[viewname], base.[viewdescription], base.[viewtableid], base.[viewsql], base.[expressionid],  obj.[locked], obj.[lastupdated], obj.[lastupdatedby]
						FROM dbo.[tbsys_views] base
						INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.viewid AND obj.objecttype = 3
						INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
						WHERE obj.effectivedate <= dt.[date]
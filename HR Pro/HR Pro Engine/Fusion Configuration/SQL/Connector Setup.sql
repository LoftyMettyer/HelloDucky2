	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessageCheckContext]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessageCheckContext];

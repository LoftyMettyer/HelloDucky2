@REM "c:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\x64"\xsd.exe staffChange.xsd staffContactChange.xsd .\commonTypes.xsd /classes /namespace:StaffPlan.Fusion.Connector.Schemas
@REM IF EXIST FusionXmlMessages.cs DEL FusionXmlMessages.cs
@REM RENAME commonTypes.cs FusionXmlMessages.cs

..\..\Tools\Xsd2code\Xsd2Code.exe FusionXmlMessages.xsd Fusion.Messages.SocialCare.Schemas FusionXmlMessages.cs /ggbc+ /gbc+ /is- /xa+
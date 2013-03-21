IF "%1" == "" GOTO EXIT


CD "Core\Fusion.LogService"
ren app.config fusion.logservice.DLL.config
nservicebus.host.exe %1 fusion.core.production /servicename:fusion.logservice /displayname:"Fusion Log Service" /endpointname:fusion.logservice /description:"Provides logging information for Fusion integration" /startManually
CD ..\..

CD "Core\Fusion.Republisher.Staff"
ren app.config fusion.publisher.socialcare.staff.DLL.config
nservicebus.host.exe %1 fusion.core.production /servicename:fusion.publisher.socialcare /displayname:"Fusion Republisher (Staff)" /endpointname:fusion.publisher.socialcare.staff /description:"Provides the core Fusion messagebus integration" /startManually
CD ..\..

CD "Core\Fusion.Test.SocialCare"
ren app.config fusion.test.socialcare.DLL.config
nservicebus.host.exe %1 fusion.core.production /servicename:fusion.test.socialcare /displayname:"Fusion Test Harness (Social Care)" /endpointname:fusion.test.socialcare /description:"Provides a test harness for Fusion integration" /startManually
CD ..\..

CD "Connector.Live"
REN app.config fusion.connector.openhr.DLL.config
nservicebus.host.exe %1 fusion.core.production /serviceName:fusion.connector.openhr.live /displayName:"Fusion OpenHR Connector (Live)" /endpoint:fusion.connector.openhr /description:"Provides Fusion integration for OpenHR live database"
CD ..

CD "Connector.Test"
REN app.config fusion.connector.openhr.DLL.config
nservicebus.host.exe %1 fusion.core.production /serviceName:fusion.connector.openhr.test /displayName:"Fusion OpenHR Connector (Test)" /endpoint:fusion.connector.openhr /description:"Provides Fusion integration for OpenHR test database" /startManually
CD ..

:EXIT

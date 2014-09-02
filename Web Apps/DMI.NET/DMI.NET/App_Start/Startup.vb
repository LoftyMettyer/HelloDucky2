Imports DMI.NET.Code.Hubs
Imports Microsoft.Owin
Imports Owin
Imports System.Data.SqlClient

<Assembly: OwinStartup(GetType(Startup))> 
Public Class Startup
	Public Sub Configuration(app As IAppBuilder)

		' SignalR configuration
		app.MapSignalR()

	End Sub
End Class
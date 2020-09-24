<div align="center">

## Starting, Pausing, and Stopping NT Services


</div>

### Description

You can use ADSI to manage local and

remote services. This code shows

and describes the different methods

and properties ADSI provides for

services. Please vote and leave your

comments for this submission.
 
### More Info
 
The ADSI Service Object supports two COM interfaces, IADsService and IADsServiceOperations.

The properties of the IADsService interface are listed below:

Property 		Description

----

Dependencies 		The list of services you need to start before you can

start this service.

DisplayName 		The friendly name that is displayed for this service.

ErrorControl 		The severity of the alert if the service does not start.

HostComputer 		The ADsPath string of the host of this service.

LoadOrderGroup 		Name of the load order group of which this service is a

member.

Path 			Path and filename of the executable file for this service.

ServiceAccountName 	Name of the account that this service uses for authentication.

ServiceAccountPath 	The ADsPath string for the user account this service is using.

ServiceType 		The description of the service type on the host computer.

StartType 		One of five possible types that determines when this service

starts.

StartupParameters 	Parameters passed to the service executable file.

Version 		Version information of the service.

The IADsServiceOperations service settings are listed below:

Property 		Description

----

Status 			The current operational state of the service.

Start 			Starts the service.

Stop 			Stops the service.

Pause 			Pauses the service.

Continue 		Resumes the paused service.

SetPassword 		Sets the password for the Service Account.

The following example of VB code lists the running services on a computer with some

common properties. Insert the following code into your project. Place a combobox onto

the form named cboService, and a command button named cmdStop. Make a reference to

the ActiveDS.TLB in the project references under "Active DS Type Library".

This VB code example loads a form displaying a combobox and a command button. The

combobox displays all running services. The user may select one of these listed

services and click on the Stop button to stop the running service. It does not

handle errors or refresh the combobox.

Based on some code found on MSDN that

discussed scripting support for services.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Blake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-blake.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-blake-starting-pausing-and-stopping-nt-services__1-15040/archive/master.zip)





### Source Code

```
Option Explicit
Public Function Services() As Boolean
 Dim oCol As New Collection
 Dim oSysInfo As New ActiveDs.WinNTSystemInfo
 Dim oComp As ActiveDs.IADsComputer
 Dim oSvc As ActiveDs.IADsServiceOperations
 Dim sCompName As String
 On Error Resume Next
 Services = False
 sCompName = "WinNT://" & oSysInfo.ComputerName & ",computer"
 Set oComp = GetObject(sCompName)
 oComp.Filter = Array("Service")
 For Each oSvc In oComp
 Debug.Print "Service display name = " & oSvc.DisplayName
 Debug.Print "Service name = " & oSvc.Name
 Debug.Print "Service account name = " & oSvc.ServiceAccountName
 Debug.Print "Service executable = " & oSvc.Path
 Debug.Print "Current status = " & oSvc.Status & vbCrLf
 If oSvc.Status = 4 Then
 'Show only running services
 cboService.AddItem oSvc.Name
 End If
 Next
 Set oSvc = Nothing
 Set oComp = Nothing
 Set oSysInfo = Nothing
 Services = True
End Function
Private Sub cmdStop_Click()
 Dim oSysInfo As New ActiveDs.WinNTSystemInfo
 Dim oComp As ActiveDs.IADsComputer
 Dim oSvc As ActiveDs.IADsServiceOperations
 Dim sCompName As String
 Dim sSvc As String
 sSvc = cboService.Text
 sCompName = "WinNT://" & oSysInfo.ComputerName & ",computer"
 Set oComp = GetObject(sCompName)
 Set oSvc = oComp.GetObject("Service", sSvc)
 oSvc.Stop
 Set oSvc = Nothing
 Set oComp = Nothing
 Set oSysInfo = Nothing
End Sub
Private Sub Form_Load()
 Services
End Sub
```


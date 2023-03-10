# MAPIFlag Enumeration


Generic MAPI Bitmask



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum MAPIFlag
```
**VB**
``` VB
Public Enumeration MAPIFlag
```
**C++**
``` C++
public enum class MAPIFlag
```



## Members
<table>
<tr>
<td>MAPI_LOGON_UI</td>
<td>1</td>
<td>A dialog box should be displayed to prompt the user for logon information if required.</td></tr>
<tr>
<td>AB_NO_DIALOG</td>
<td>1</td>
<td>Suppresses the display of dialog boxes. If the AB_NO_DIALOG flag is not set, the address book providers that contribute to the integrated address book can prompt the user for any necessary information.</td></tr>
<tr>
<td>NEW_SESSION</td>
<td>2</td>
<td>An attempt should be made to create a new MAPI session instead of acquiring the shared session.</td></tr>
<tr>
<td>ALLOW_OTHERS</td>
<td>8</td>
<td>The shared session should be returned, which allows later clients to obtain the session without providing any user credentials.</td></tr>
<tr>
<td>MAPI_DIALOG</td>
<td>8</td>
<td>The ulUIParam parameter is ignored unless the MAPI_DIALOG flag is set.</td></tr>
<tr>
<td>EXPLICIT_PROFILE</td>
<td>16</td>
<td>The default profile should not be used and the user should be required to supply a profile.</td></tr>
<tr>
<td>BEST_ACCESS</td>
<td>16</td>
<td>Requests that the object be opened by using the maximum network permissions allowed for the user and the maximum client application access.</td></tr>
<tr>
<td>EXTENDED</td>
<td>32</td>
<td>Log on with extended capabilities. This flag should always be set.</td></tr>
<tr>
<td>SPOOLER</td>
<td>37</td>
<td>MAPI spooler is a process that is responsible for sending messages to and receiving messages from a messaging system.</td></tr>
<tr>
<td>USE_DEFAULT</td>
<td>64</td>
<td>The messaging subsystem should substitute the profile name of the default profile for the Profile Name parameter. T</td></tr>
<tr>
<td>FORCE_DOWNLOAD</td>
<td>4,096</td>
<td>An attempt should be made to download all of the user's messages before returning. If it is not set, messages can be downloaded in the background after the call to MAPILogonEx returns.</td></tr>
<tr>
<td>SERVICE_UI_ALWAYS</td>
<td>8,192</td>
<td>MAPILogonEx should display a configuration dialog box for each message service in the profile.</td></tr>
<tr>
<td>NO_MAIL</td>
<td>32,768</td>
<td>MAPI should not inform the MAPI spooler of the session's existence.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

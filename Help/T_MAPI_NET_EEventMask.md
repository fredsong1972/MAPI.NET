# EEventMask Enumeration


A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration. There is a corresponding NOTIFICATION structure associated with each type of event that holds information about the event.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum EEventMask
```
**VB**
``` VB
Public Enumeration EEventMask
```
**C++**
``` C++
public enum class EEventMask
```



## Members
<table>
<tr>
<td>fnevCriticalError</td>
<td>1</td>
<td>Registers for notifications about severe errors, such as insufficient memory.</td></tr>
<tr>
<td>fnevNewMail</td>
<td>2</td>
<td>Registers for notifications about the arrival of new messages.</td></tr>
<tr>
<td>fnevObjectCreated</td>
<td>4</td>
<td>Registers for notifications about the creation of a new folder or message.</td></tr>
<tr>
<td>fnevObjectDeleted</td>
<td>8</td>
<td>Registers for notifications about a folder or message being deleted.</td></tr>
<tr>
<td>fnevObjectModified</td>
<td>16</td>
<td>Registers for notifications about a folder or message being modified.</td></tr>
<tr>
<td>fnevObjectMoved</td>
<td>32</td>
<td>Registers for notifications about a folder or message being moved.</td></tr>
<tr>
<td>fnevObjectCopied</td>
<td>64</td>
<td>Registers for notifications about a folder or message being copied.</td></tr>
<tr>
<td>fnevSearchComplete</td>
<td>128</td>
<td>Registers for notifications about the completion of a search operation.</td></tr>
<tr>
<td>fnevTableModified</td>
<td>256</td>
<td>Registers for notifications about a table being modified.</td></tr>
<tr>
<td>fnevStatusObjectModified</td>
<td>512</td>
<td>Registers for notifications about a status object being modified.</td></tr>
<tr>
<td>fnevReservedForMapi</td>
<td>1,073,741,824</td>
<td>Reserved</td></tr>
<tr>
<td>fnevExtended</td>
<td>2,147,483,648</td>
<td>Registers for notifications about events specific to the particular message store provider.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

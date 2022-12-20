# HRESULT Enumeration


The HRESULT data type is a 32-bit value is used to describe an error or warning.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum HRESULT
```
**VB**
``` VB
Public Enumeration HRESULT
```
**C++**
``` C++
public enum class HRESULT
```



## Members
<table>
<tr>
<td>S_OK</td>
<td>0</td>
<td>OK</td></tr>
<tr>
<td>S_FALSE</td>
<td>1</td>
<td>False</td></tr>
<tr>
<td>E_OUTOFMEMORY</td>
<td>2,147,483,650</td>
<td>Out of memory</td></tr>
<tr>
<td>MAPI_E_INVALID_PARAMETER</td>
<td>2,147,483,651</td>
<td>Invalid parameter</td></tr>
<tr>
<td>MAPI_E_INTERFACE_NOT_SUPPORTED</td>
<td>2,147,483,652</td>
<td>Interface not supported</td></tr>
<tr>
<td>E_FAIL</td>
<td>2,147,483,656</td>
<td>Fail</td></tr>
<tr>
<td>MAPI_E_NO_ACCESS</td>
<td>2,147,483,657</td>
<td>No access</td></tr>
<tr>
<td>E_NOTIMPL</td>
<td>2,147,500,033</td>
<td>Not implemented</td></tr>
<tr>
<td>MAPI_E_CALL_FAILED</td>
<td>2,147,500,037</td>
<td>Call failed</td></tr>
<tr>
<td>E_UNEXPECTED</td>
<td>2,147,549,183</td>
<td>Unexpected</td></tr>
<tr>
<td>MAPI_E_NO_SUPPORT</td>
<td>2,147,746,050</td>
<td>No support</td></tr>
<tr>
<td>MAPI_E_BAD_CHARWIDTH</td>
<td>2,147,746,051</td>
<td>Bar char width</td></tr>
<tr>
<td>MAPI_E_STRING_TOO_LONG</td>
<td>2,147,746,053</td>
<td>String too long</td></tr>
<tr>
<td>MAPI_E_UNKNOWN_FLAGS</td>
<td>2,147,746,054</td>
<td>Unknown flags</td></tr>
<tr>
<td>MAPI_E_INVALID_ENTRYID</td>
<td>2,147,746,055</td>
<td>Invalid entry ID</td></tr>
<tr>
<td>MAPI_E_INVALID_OBJECT</td>
<td>2,147,746,056</td>
<td>Invalid object</td></tr>
<tr>
<td>MAPI_E_OBJECT_CHANGED</td>
<td>2,147,746,057</td>
<td>Object changed</td></tr>
<tr>
<td>MAPI_E_OBJECT_DELETED</td>
<td>2,147,746,058</td>
<td>Object deleted</td></tr>
<tr>
<td>MAPI_E_BUSY</td>
<td>2,147,746,059</td>
<td>Busy</td></tr>
<tr>
<td>MAPI_E_NOT_ENOUGH_DISK</td>
<td>2,147,746,061</td>
<td>Not enough disk</td></tr>
<tr>
<td>MAPI_E_NOT_ENOUGH_RESOURCES</td>
<td>2,147,746,062</td>
<td>Not enough resources</td></tr>
<tr>
<td>MAPI_E_NOT_FOUND</td>
<td>2,147,746,063</td>
<td>Not found</td></tr>
<tr>
<td>MAPI_E_VERSION</td>
<td>2,147,746,064</td>
<td>Version error</td></tr>
<tr>
<td>MAPI_E_LOGON_FAILED</td>
<td>2,147,746,065</td>
<td>Logon failed</td></tr>
<tr>
<td>MAPI_E_SESSION_LIMIT</td>
<td>2,147,746,066</td>
<td>Session limited</td></tr>
<tr>
<td>MAPI_E_USER_CANCEL</td>
<td>2,147,746,067</td>
<td>User cancel</td></tr>
<tr>
<td>MAPI_E_UNABLE_TO_ABORT</td>
<td>2,147,746,068</td>
<td>Unable to abort</td></tr>
<tr>
<td>MAPI_E_NETWORK_ERROR</td>
<td>2,147,746,069</td>
<td>Network error</td></tr>
<tr>
<td>MAPI_E_DISK_ERROR</td>
<td>2,147,746,070</td>
<td>Disk error</td></tr>
<tr>
<td>MAPI_E_TOO_COMPLEX</td>
<td>2,147,746,071</td>
<td>Too complex</td></tr>
<tr>
<td>MAPI_E_BAD_COLUMN</td>
<td>2,147,746,072</td>
<td>Bad column</td></tr>
<tr>
<td>MAPI_E_EXTENDED_ERROR</td>
<td>2,147,746,073</td>
<td>Extended error</td></tr>
<tr>
<td>MAPI_E_COMPUTED</td>
<td>2,147,746,074</td>
<td>Computed</td></tr>
<tr>
<td>MAPI_E_CORRUPT_DATA</td>
<td>2,147,746,075</td>
<td>Corrupt data</td></tr>
<tr>
<td>MAPI_E_UNCONFIGURED</td>
<td>2,147,746,076</td>
<td>Unconfigured</td></tr>
<tr>
<td>MAPI_E_FAILONEPROVIDER</td>
<td>2,147,746,077</td>
<td>One provider fail</td></tr>
<tr>
<td>MAPI_E_UNKNOWN_CPID</td>
<td>2,147,746,078</td>
<td>Unknown CPID</td></tr>
<tr>
<td>MAPI_E_UNKNOWN_LCID</td>
<td>2,147,746,079</td>
<td>Unknown LCID</td></tr>
<tr>
<td>MAPI_E_CORRUPT_STORE</td>
<td>2,147,747,328</td>
<td>Corrupt store</td></tr>
<tr>
<td>MAPI_E_NOT_IN_QUEUE</td>
<td>2,147,747,329</td>
<td>Not in queue</td></tr>
<tr>
<td>MAPI_E_NO_SUPPRESS</td>
<td>2,147,747,330</td>
<td>No suppress</td></tr>
<tr>
<td>MAPI_E_COLLISION</td>
<td>2,147,747,332</td>
<td>Collision</td></tr>
<tr>
<td>MAPI_E_NOT_INITIALIZED</td>
<td>2,147,747,333</td>
<td>Not Initialized</td></tr>
<tr>
<td>MAPI_E_NON_STANDARD</td>
<td>2,147,747,334</td>
<td>Not standard</td></tr>
<tr>
<td>MAPI_E_NO_RECIPIENTS</td>
<td>2,147,747,335</td>
<td>No recipients</td></tr>
<tr>
<td>MAPI_E_SUBMITTED</td>
<td>2,147,747,336</td>
<td>Submitted error</td></tr>
<tr>
<td>MAPI_E_HAS_FOLDERS</td>
<td>2,147,747,337</td>
<td>Has folders</td></tr>
<tr>
<td>MAPI_E_HAS_MESSAGES</td>
<td>2,147,747,338</td>
<td>Has message</td></tr>
<tr>
<td>MAPI_E_FOLDER_CYCLE</td>
<td>2,147,747,339</td>
<td>Folder cycle</td></tr>
<tr>
<td>MAPI_E_NOT_ENOUGH_MEMORY</td>
<td>2,147,942,414</td>
<td>Not enough memory</td></tr>
<tr>
<td>E_INVALIDARG</td>
<td>2,147,942,487</td>
<td>Invalid argument</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

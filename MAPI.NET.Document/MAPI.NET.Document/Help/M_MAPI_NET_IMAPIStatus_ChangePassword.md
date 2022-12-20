# ChangePassword Method


Modifies a service provider's password without displaying a user interface.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT ChangePassword(
	string lpOldPass,
	string lpNewPass,
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function ChangePassword ( 
	lpOldPass As String,
	lpNewPass As String,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT ChangePassword(
	String^ lpOldPass, 
	String^ lpNewPass, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The old password</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The new password</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the format of the passwords.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the password modification was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIStatus.md">IMAPIStatus Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

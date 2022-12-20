# ValidateState Method


Confirms the external status information available for the MAPI resource or the service provider.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT ValidateState(
	IntPtr ulUIParam,
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function ValidateState ( 
	ulUIParam As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT ValidateState(
	IntPtr ulUIParam, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A handle to the parent window of any dialog boxes or windows that this method displays.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the validation.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the validation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIStatus.md">IMAPIStatus Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

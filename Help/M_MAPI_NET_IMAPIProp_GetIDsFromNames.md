# GetIDsFromNames Method


Provides the property identifiers that correspond to one or more property names.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetIDsFromNames(
	uint cPropNames,
	ref IntPtr lppPropNames,
	uint ulFlags,
	out IntPtr lppPropTags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetIDsFromNames ( 
	cPropNames As UInteger,
	ByRef lppPropNames As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppPropTags As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetIDsFromNames(
	unsigned int cPropNames, 
	IntPtr% lppPropNames, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppPropTags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of property names pointed to by the lppPropNames parameter. If lppPropNames is NULL, the cPropNames parameter must be 0.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of property names, or NULL. Passing NULL requests property identifiers for all property names in all property sets about which the object has information. The lppPropNames parameter must not be NULL if the MAPI_CREATE flag is set in the ulFlags parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that indicates how the property identifiers should be returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to an array of property tags that contains existing or newly assigned property identifiers.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the identifiers for the specified property names were successfully returned;otherwise failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

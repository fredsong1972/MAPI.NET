# GetNamesFromIDs Method


Provides the property names that correspond to one or more property identifiers.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetNamesFromIDs(
	out IntPtr lppPropTags,
	ref Guid lpPropSetGuid,
	uint ulFlags,
	out uint lpcPropNames,
	out IntPtr lpppPropNames
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetNamesFromIDs ( 
	<OutAttribute> ByRef lppPropTags As IntPtr,
	ByRef lpPropSetGuid As Guid,
	ulFlags As UInteger,
	<OutAttribute> ByRef lpcPropNames As UInteger,
	<OutAttribute> ByRef lpppPropNames As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetNamesFromIDs(
	[OutAttribute] IntPtr% lppPropTags, 
	Guid% lpPropSetGuid, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpcPropNames, 
	[OutAttribute] IntPtr% lpppPropNames
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to an SPropTagArray structure that contains an array of property tags; otherwise, NULL, indicating that all names should be returned. The cValues member for the property tag array cannot be 0. If lppPropTags is a valid pointer on input, GetNamesFromIDs returns names for each property identifier included in the array.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to a GUID, or GUID structure, that identifies a property set. The lpPropSetGuid parameter can point to a valid property set or can be NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that indicates the type of names to be returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a count of the property name pointers in the array pointed to by the lppPropNames parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of pointers to MAPINAMEID structures that contains property names.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the property names were successfully returned;otherwise failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

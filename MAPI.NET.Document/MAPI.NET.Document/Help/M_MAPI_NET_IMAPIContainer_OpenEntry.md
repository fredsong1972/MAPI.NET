# OpenEntry Method


Opens an object in the container, returning an interface pointer for further access.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT OpenEntry(
	uint cbEntryID,
	IntPtr lpEntryID,
	IntPtr lpInterface,
	uint ulFlags,
	out uint lpulObjType,
	out IntPtr lppUnk
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function OpenEntry ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	lpInterface As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lpulObjType As UInteger,
	<OutAttribute> ByRef lppUnk As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT OpenEntry(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	IntPtr lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpulObjType, 
	[OutAttribute] IntPtr% lppUnk
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the object to open. If lpEntryID is set to NULL, the top-level container in the container's hierarchy is opened.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the object.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the object is opened.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the opened object's type.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the interface implementation to use to access the open object.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the object was successfully opened; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIContainer.md">IMAPIContainer Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

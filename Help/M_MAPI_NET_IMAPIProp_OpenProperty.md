# OpenProperty Method


Returns a pointer to an interface that can be used to access a property.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT OpenProperty(
	uint ulPropTag,
	ref Guid lpiid,
	uint ulInterfaceOptions,
	uint ulFlags,
	out IntPtr lppUnk
)
```
**VB**
``` VB
Function OpenProperty ( 
	ulPropTag As UInteger,
	ByRef lpiid As Guid,
	ulInterfaceOptions As UInteger,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppUnk As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT OpenProperty(
	unsigned int ulPropTag, 
	Guid% lpiid, 
	unsigned int ulInterfaceOptions, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppUnk
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The property tag for the property to be accessed. Both the identifier and the type must be included in the property tag</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to the identifier for the interface to be used to access the property. The lpiid parameter must not be null.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Data that relates to the interface identified by the lpiid parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls access to the property.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the requested interface to be used for property access.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the requested interface pointer was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

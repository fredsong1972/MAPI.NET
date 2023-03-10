# OpenProperty(UInt32, Guid, OpenPropertyMode) Method


Returns a pointer to an interface that can be used to access a property.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public IntPtr OpenProperty(
	uint tag,
	Guid iid,
	OpenPropertyMode mode
)
```
**VB**
``` VB
Public Function OpenProperty ( 
	tag As UInteger,
	iid As Guid,
	mode As OpenPropertyMode
) As IntPtr
```
**C++**
``` C++
public:
IntPtr OpenProperty(
	unsigned int tag, 
	Guid iid, 
	OpenPropertyMode mode
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The property tag for the property to be accessed</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>the Guid for the interface to be used to access the property</dd><dt>  <a href="T_MAPI_NET_OpenPropertyMode.md">OpenPropertyMode</a></dt><dd>A bitmask of flags that controls access to the property</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a>  
the requested interface pointer

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="Overload_MAPI_NET_MAPIObject_OpenProperty.md">OpenProperty Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

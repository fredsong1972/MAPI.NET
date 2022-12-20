# CopyTo(UInt32[], IMAPIProp) Method


Copies or moves all properties, except for specifically excluded properties.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool CopyTo(
	uint[] excludeProps,
	IMAPIProp mapiObject
)
```
**VB**
``` VB
Public Function CopyTo ( 
	excludeProps As UInteger(),
	mapiObject As IMAPIProp
) As Boolean
```
**C++**
``` C++
public:
bool CopyTo(
	array<unsigned int>^ excludeProps, 
	IMAPIProp^ mapiObject
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>A property tag array that identifies the property tags that should be excluded from the copy or move operation.</dd><dt>  <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a></dt><dd>A mapi prop object to receive the copied or moved properties.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true if successful; otherwise, false

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="Overload_MAPI_NET_MAPIObject_CopyTo.md">CopyTo Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

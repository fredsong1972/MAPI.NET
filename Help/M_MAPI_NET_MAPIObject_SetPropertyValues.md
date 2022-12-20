# SetPropertyValues Method


Updates one or more properties of the object.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
protected bool SetPropertyValues(
	IPropValue[] props
)
```
**VB**
``` VB
Protected Function SetPropertyValues ( 
	props As IPropValue()
) As Boolean
```
**C++**
``` C++
protected:
bool SetPropertyValues(
	array<IPropValue^>^ props
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_IPropValue.md">IPropValue</a>[]</dt><dd>An array of IPropValue that contain property values to be updated.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true if successful; otherwise, false

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

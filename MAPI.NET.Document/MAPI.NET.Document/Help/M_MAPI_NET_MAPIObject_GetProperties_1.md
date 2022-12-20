# GetProperties(UInt32[]) Method


Retrieves the property value of one or more properties of an object.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public Dictionary<uint, IPropValue> GetProperties(
	uint[] tags
)
```
**VB**
``` VB
Public Function GetProperties ( 
	tags As UInteger()
) As Dictionary(Of UInteger, IPropValue)
```
**C++**
``` C++
public:
Dictionary<unsigned int, IPropValue^>^ GetProperties(
	array<unsigned int>^ tags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>Array of uint tags</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.collections.generic.dictionary-2" target="_blank" rel="noopener noreferrer">Dictionary</a>(<a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>, <a href="T_MAPI_NET_IPropValue.md">IPropValue</a>)  
A dictionary of the retrieved property values.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="Overload_MAPI_NET_MAPIObject_GetProperties.md">GetProperties Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

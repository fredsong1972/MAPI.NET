# GetProps Method


Retrieves the property value of one or more properties of an object.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetProps(
	uint[] lpPropTagArray,
	uint ulFlags,
	out uint lpcValues,
	ref IntPtr lppPropArray
)
```
**VB**
``` VB
Function GetProps ( 
	lpPropTagArray As UInteger(),
	ulFlags As UInteger,
	<OutAttribute> ByRef lpcValues As UInteger,
	ByRef lppPropArray As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetProps(
	[InAttribute] array<unsigned int>^ lpPropTagArray, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpcValues, 
	IntPtr% lppPropArray
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>A pointer to an array of property tags that identify the properties whose values are to be retrieved. The lpPropTagArray parameter must be either NULL, indicating that values for all properties of the object should be returned, or point to an SPropTagArray structure that contains one or more property tags.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that indicates the format for properties that have the type PT_UNSPECIFIED.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a count of property values pointed to by the lppPropArray parameter. If lppPropArray is NULL, the content of the lpcValues parameter is zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the retrieved property values.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK The property values were successfully retrieved. MAPI_W_ERRORS_RETURNED The call succeeded overall, but one or more properties could not be accessed. The aulPropTag member of the property value for each unavailable property has a property type of PT_ERROR and an identifier of zero. When this warning is returned, the call should be handled as successful. MAPI_E_INVALID_PARAMETER Zero was passed in the cValues member of the SPropTagArray structure pointed to by lpPropTagArray.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

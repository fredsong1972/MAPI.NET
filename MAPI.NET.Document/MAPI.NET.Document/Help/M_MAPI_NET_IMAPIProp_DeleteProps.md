# DeleteProps Method


Deletes one or more properties from an object.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT DeleteProps(
	IntPtr lpPropTagArray,
	out IntPtr lppProblems
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function DeleteProps ( 
	lpPropTagArray As IntPtr,
	<OutAttribute> ByRef lppProblems As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT DeleteProps(
	IntPtr lpPropTagArray, 
	[OutAttribute] IntPtr% lppProblems
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of property tags that indicate the properties to delete. The cValues member of the SPropTagArray structure pointed to by lpPropTagArray must not be zero, and the lpPropTagArray parameter itself must not be NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, NULL, which indicates that there is no need for error information. If lppProblems is a valid pointer on input, DeleteProps returns detailed information about errors in deleting one or more properties.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if properties were successfully deleted;otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

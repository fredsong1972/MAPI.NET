# CopyProps Method


Copies or moves selected properties.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT CopyProps(
	IntPtr lpIncludeProps,
	uint ulUIParam,
	IntPtr lpProgress,
	ref Guid lpInterface,
	IntPtr lpDestObj,
	uint ulFlags,
	out IntPtr lppProblems
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function CopyProps ( 
	lpIncludeProps As IntPtr,
	ulUIParam As UInteger,
	lpProgress As IntPtr,
	ByRef lpInterface As Guid,
	lpDestObj As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppProblems As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT CopyProps(
	IntPtr lpIncludeProps, 
	unsigned int ulUIParam, 
	IntPtr lpProgress, 
	Guid% lpInterface, 
	IntPtr lpDestObj, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppProblems
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a property tag array that specifies the properties to copy or move. PR_NULL (PidTagNull) cannot be included in the array. The lpIncludeProps parameter cannot be null.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an implementation of a progress indicator. If null is passed in the lpProgress parameter, the progress indicator is displayed by using the MAPI implementation. The lpProgress parameter is ignored unless the MAPI_DIALOG flag is set in the ulFlags parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface that must be used to access the object pointed to by the lpDestObj parameter. The lpInterface parameter must not be null.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the object to receive the copied or moved properties.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the copy or move operation.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, null, indicating that there is no need for error information. If lppProblems is a valid pointer on input, CopyProps returns detailed information about errors in copying one or more properties.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if properties have been successfully copied or moved;otherwise failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

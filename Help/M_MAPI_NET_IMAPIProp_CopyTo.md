# CopyTo Method


Copies or moves all properties, except for specifically excluded properties.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT CopyTo(
	uint ciidExclude,
	ref Guid rgiidExclude,
	uint[] lpExcludeProps,
	IntPtr ulUIParam,
	IntPtr lpProgress,
	ref Guid lpInterface,
	IntPtr lpDestObj,
	uint ulFlags,
	IntPtr lppProblems
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function CopyTo ( 
	ciidExclude As UInteger,
	ByRef rgiidExclude As Guid,
	lpExcludeProps As UInteger(),
	ulUIParam As IntPtr,
	lpProgress As IntPtr,
	ByRef lpInterface As Guid,
	lpDestObj As IntPtr,
	ulFlags As UInteger,
	lppProblems As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT CopyTo(
	unsigned int ciidExclude, 
	Guid% rgiidExclude, 
	[InAttribute] array<unsigned int>^ lpExcludeProps, 
	IntPtr ulUIParam, 
	IntPtr lpProgress, 
	Guid% lpInterface, 
	IntPtr lpDestObj, 
	unsigned int ulFlags, 
	IntPtr lppProblems
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of interfaces to exclude when properties are copied or moved.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>An array of interface identifiers (IIDs) that specify interfaces that should not be used when supplemental information is copied or moved to the destination object.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>A pointer to a property tag array that identifies the property tags that should be excluded from the copy or move operation. Passing null in the lpExcludeProps parameter indicates that all of the object's properties should be copied or moved. CopyTo returns MAPI_E_INVALID_PARAMETER when the cValues member of the SPropProblemArray structure pointed to by lpExcludeProps is set to 0.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A handle to the parent window of the progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a progress indicator implementation. If null is passed in the lpProgress parameter, MAPI provides the progress implementation. The lpProgress parameter is ignored unless the MAPI_DIALOG flag is set in the ulFlags parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the object pointed to by the lpDestObj parameter. The lpInterface parameter must not be null.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the object to receive the copied or moved properties.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the copy or move operation.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, null, indicating no need for error information. If lppProblems is a valid pointer on input, CopyTo returns detailed information about errors in copying one or more properties.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the properties have been successfully copied or moved;otherwise failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  

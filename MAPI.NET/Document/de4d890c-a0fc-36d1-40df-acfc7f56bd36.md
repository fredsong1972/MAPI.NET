# DeleteProps Method


Deletes one or more properties from an object.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
HRESULT DeleteProps(
	IntPtr lpPropTagArray,
	out IntPtr lppProblems
)
```
**VB**
``` VB
Function DeleteProps ( 
	lpPropTagArray As IntPtr,
	<OutAttribute> ByRef lppProblems As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT DeleteProps(
	IntPtr lpPropTagArray, 
	[OutAttribute] IntPtr% lppProblems
)
```
**F#**
``` F#
abstract DeleteProps : 
        lpPropTagArray : IntPtr * 
        lppProblems : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of property tags that indicate the properties to delete. The cValues member of the SPropTagArray structure pointed to by lpPropTagArray must not be zero, and the lpPropTagArray parameter itself must not be NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, NULL, which indicates that there is no need for error information. If lppProblems is a valid pointer on input, DeleteProps returns detailed information about errors in deleting one or more properties.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if properties were successfully deleted;otherwise, failed.

## See Also


#### Reference
<a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

# GetLastError Method




## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.0

**C#**
``` C#
HRESULT GetLastError(
	int hResult,
	uint ulFlags,
	out IntPtr lppMAPIError
)
```
**VB**
``` VB
Function GetLastError ( 
	hResult As Integer,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppMAPIError As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetLastError(
	int hResult, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppMAPIError
)
```
**F#**
``` F#
abstract GetLastError : 
        hResult : int * 
        ulFlags : uint32 * 
        lppMAPIError : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>\[Missing &lt;param name="hResult"/&gt; documentation for "M:MAPI.NET.IMessage.GetLastError(System.Int32,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>\[Missing &lt;param name="ulFlags"/&gt; documentation for "M:MAPI.NET.IMessage.GetLastError(System.Int32,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>\[Missing &lt;param name="lppMAPIError"/&gt; documentation for "M:MAPI.NET.IMessage.GetLastError(System.Int32,System.UInt32,System.IntPtr@)"\]</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
\[Missing &lt;returns&gt; documentation for "M:MAPI.NET.IMessage.GetLastError(System.Int32,System.UInt32,System.IntPtr@)"\]

#### Implements
<a href="5bef0dfc-c21a-ed22-b4b6-aebbc8ed696a.md">IMAPIProp.GetLastError(Int32, UInt32, IntPtr)</a>  


## See Also


#### Reference
<a href="f542b7a9-d1ab-fed6-c2df-7c20b044fccc.md">IMessage Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
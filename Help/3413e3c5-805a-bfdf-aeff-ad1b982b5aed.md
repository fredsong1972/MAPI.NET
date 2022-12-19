# CreateFolder Method




## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.0

**C#**
``` C#
HRESULT CreateFolder(
	uint ulFolderType,
	string lpszFolderName,
	string lpszFolderComment,
	IntPtr lpInterface,
	uint ulFlags,
	out IntPtr lppFolder
)
```
**VB**
``` VB
Function CreateFolder ( 
	ulFolderType As UInteger,
	lpszFolderName As String,
	lpszFolderComment As String,
	lpInterface As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppFolder As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT CreateFolder(
	unsigned int ulFolderType, 
	String^ lpszFolderName, 
	String^ lpszFolderComment, 
	IntPtr lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppFolder
)
```
**F#**
``` F#
abstract CreateFolder : 
        ulFolderType : uint32 * 
        lpszFolderName : string * 
        lpszFolderComment : string * 
        lpInterface : IntPtr * 
        ulFlags : uint32 * 
        lppFolder : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>\[Missing &lt;param name="ulFolderType"/&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>\[Missing &lt;param name="lpszFolderName"/&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>\[Missing &lt;param name="lpszFolderComment"/&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>\[Missing &lt;param name="lpInterface"/&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>\[Missing &lt;param name="ulFlags"/&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>\[Missing &lt;param name="lppFolder"/&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
\[Missing &lt;returns&gt; documentation for "M:MAPI.NET.IMAPIFolder.CreateFolder(System.UInt32,System.String,System.String,System.IntPtr,System.UInt32,System.IntPtr@)"\]

## See Also


#### Reference
<a href="a5eb5918-6571-0710-67c7-a210d1ad706f.md">IMAPIFolder Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

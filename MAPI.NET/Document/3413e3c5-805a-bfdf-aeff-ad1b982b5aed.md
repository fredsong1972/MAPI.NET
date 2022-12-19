# CreateFolder Method


Creates a new subfolder.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The type of folder to create.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The name for the new folder.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>A comment associated with the new folder.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the new folder. Passing NULL causes the message store provider to return the standard folder interface, IMAPIFolder : IMAPIContainer.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the folder is created.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the newly created folder.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the new folder has been successfully created or opened; otherwise, failed.

## See Also


#### Reference
<a href="a5eb5918-6571-0710-67c7-a210d1ad706f.md">IMAPIFolder Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

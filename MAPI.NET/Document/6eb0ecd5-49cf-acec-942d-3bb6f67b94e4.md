# CopyFolder Method


Copies or moves a subfolder.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
HRESULT CopyFolder(
	uint cbEntryID,
	IntPtr lpEntryID,
	IntPtr lpInterface,
	IntPtr lpDestFolder,
	string lpszNewFolderName,
	IntPtr ulUIParam,
	IntPtr lpProgress,
	uint ulFlags
)
```
**VB**
``` VB
Function CopyFolder ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	lpInterface As IntPtr,
	lpDestFolder As IntPtr,
	lpszNewFolderName As String,
	ulUIParam As IntPtr,
	lpProgress As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT CopyFolder(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	IntPtr lpInterface, 
	IntPtr lpDestFolder, 
	String^ lpszNewFolderName, 
	IntPtr ulUIParam, 
	IntPtr lpProgress, 
	unsigned int ulFlags
)
```
**F#**
``` F#
abstract CopyFolder : 
        cbEntryID : uint32 * 
        lpEntryID : IntPtr * 
        lpInterface : IntPtr * 
        lpDestFolder : IntPtr * 
        lpszNewFolderName : string * 
        ulUIParam : IntPtr * 
        lpProgress : IntPtr * 
        ulFlags : uint32 -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the subfolder to copy or move.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the folder that the lpDestFolder parameter points to.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the open folder to receive the copied or moved subfolder.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The name of the copied or moved folder in its new destination.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A handle to the parent window of the progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a progress object that displays a progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the copy or move operation.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the specified folder has been successfully copied or moved; otherwise, failed.

## See Also


#### Reference
<a href="a5eb5918-6571-0710-67c7-a210d1ad706f.md">IMAPIFolder Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

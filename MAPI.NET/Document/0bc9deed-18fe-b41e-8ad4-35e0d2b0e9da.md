# OpenProfileSection Method


Opens a section of the current profile and returns an IProfSect pointer for further access.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
HRESULT OpenProfileSection(
	ref Guid lpUID,
	ref Guid lpInterface,
	uint ulFlags,
	out IntPtr lppProfSect
)
```
**VB**
``` VB
Function OpenProfileSection ( 
	ByRef lpUID As Guid,
	ByRef lpInterface As Guid,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppProfSect As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT OpenProfileSection(
	Guid% lpUID, 
	Guid% lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppProfSect
)
```
**F#**
``` F#
abstract OpenProfileSection : 
        lpUID : Guid byref * 
        lpInterface : Guid byref * 
        ulFlags : uint32 * 
        lppProfSect : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>pointer to the MAPIUID structure that identifies the profile section.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the profile section.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls access to the profile section.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the profile section.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the profile section was successfully opened; otherwise, failed.

## See Also


#### Reference
<a href="d28ec202-b730-fb1f-42ac-5545b0b43d47.md">IMAPISession Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

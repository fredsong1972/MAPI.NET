# IAddrBook Interface


Supports access to the MAPI address book and includes operations such as displaying common dialog boxes; opening containers, messaging users, and distribution lists; and performing name resolution.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
public interface IAddrBook : IMAPIProp
```
**VB**
``` VB
Public Interface IAddrBook
	Inherits IMAPIProp
```
**C++**
``` C++
public interface class IAddrBook : IMAPIProp
```
**F#**
``` F#
type IAddrBook = 
    interface
        interface IMAPIProp
    end
```

<table><tr><td><strong>Implements</strong></td><td><a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a></td></tr>
</table>



## Methods
<table>
<tr>
<td><a href="7902014e-85f5-b939-0d22-25fa3136249f.md">Address</a></td>
<td>Displays the Outlook address book dialog box.</td></tr>
<tr>
<td><a href="f5d9298d-757f-1a42-0355-bcac7a2bb6d1.md">Advise</a></td>
<td>Registers a client or service provider to receive notifications about changes to one or more entries in the address book.</td></tr>
<tr>
<td><a href="83903a16-9849-dc49-1271-7abc6b76cf02.md">CompareEntryIDs</a></td>
<td>Compares two entry identifiers that belong to a particular address book provider to determine whether they refer to the same address book object.</td></tr>
<tr>
<td><a href="ee81fc2f-a117-6a66-c47d-05642d1e885b.md">CopyProps</a></td>
<td>Copies or moves selected properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="446da6c3-cf56-9eae-0067-556449bcbd5e.md">CopyTo</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="9732b7d2-9d29-8318-e711-01dc7937ea3c.md">CreateOneOff</a></td>
<td>Creates an entry identifier for a one-off address.</td></tr>
<tr>
<td><a href="de4d890c-a0fc-36d1-40df-acfc7f56bd36.md">DeleteProps</a></td>
<td>Deletes one or more properties from an object.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="3aeb3e5f-4cb5-7c0e-28ac-76e0b10b7c51.md">Details</a></td>
<td>Displays a dialog box that shows details about a particular address book entry.</td></tr>
<tr>
<td><a href="1d03b273-e053-971f-360b-39a91300d2ae.md">GetDefaultDir</a></td>
<td>Returns the entry identifier for the initial address book container.</td></tr>
<tr>
<td><a href="78a82640-fb2e-3f54-a035-1861c1703d42.md">GetIDsFromNames</a></td>
<td>Provides the property identifiers that correspond to one or more property names.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="5bef0dfc-c21a-ed22-b4b6-aebbc8ed696a.md">GetLastError</a></td>
<td>Returns a MAPIERROR structure that contains information about the previous error.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="c216ad5d-2e67-c43f-71c9-960c28fe4cea.md">GetNamesFromIDs</a></td>
<td>Provides the property names that correspond to one or more property identifiers.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="87a269a2-18d1-3136-5321-edc8e32a5e95.md">GetPAB</a></td>
<td>Returns the entry identifier of the container that is designated as the personal address book (PAB).</td></tr>
<tr>
<td><a href="1fdf6ea2-4ee7-da0d-7329-a223aa9dc8dd.md">GetPropList</a></td>
<td>Returns property tags for all properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="eed91d74-f874-f174-2f2d-a0cbf2224590.md">GetProps</a></td>
<td>Retrieves the property value of one or more properties of an object.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="5e93fffc-3f3a-01a0-fa0a-3dbce5217513.md">GetSearchPath</a></td>
<td>Returns an ordered list of entry identifiers of the containers to be included in the name resolution process initiated by the IAddrBook::ResolveName method.</td></tr>
<tr>
<td><a href="1eec0828-3379-6b32-2ffb-6cdd09b18fd6.md">NewEntry</a></td>
<td>Adds a new recipient to an address book container or to the recipient list of an outgoing message.</td></tr>
<tr>
<td><a href="fd2f9bac-8138-6589-72df-b70bbe3346b4.md">OpenEntry</a></td>
<td>Opens an address book entry and returns a pointer to an interface that can be used to access the entry.</td></tr>
<tr>
<td><a href="a82109dc-9148-ad78-11ae-7aa020efd430.md">OpenProperty</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="8a41ab9f-0a88-14c3-c62c-aa13652627dd.md">PrepareRecips</a></td>
<td>Prepares a recipient list for later use by the messaging system.</td></tr>
<tr>
<td><a href="a8e279bd-ffe6-ccc5-9047-a73512808412.md">ResolveName</a></td>
<td>Performs name resolution, assigning entry identifiers to recipients in a recipient list.</td></tr>
<tr>
<td><a href="d26a32e5-3da7-0464-9459-2ad44613db5b.md">SaveChanges</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="1dc8b61d-02d8-3639-49ff-af9d43c7802a.md">SetDefaultDir</a></td>
<td>Establishes the specified container as the default address book container.</td></tr>
<tr>
<td><a href="2d95aaf5-52c9-26b5-d0e2-b5488cea3b6c.md">SetPAB</a></td>
<td>Designates a particular container as the personal address book (PAB).</td></tr>
<tr>
<td><a href="f1a2ab65-b81f-ec0c-d947-814cdecceca2.md">SetProps</a></td>
<td>Updates one or more properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="38d30529-c442-2ea4-5378-1d92fa8b3362.md">SetSearchPath</a></td>
<td>Sets a new search path in the profile that is used for the name resolution process.</td></tr>
<tr>
<td><a href="50cb86ed-eed2-0e72-d8b5-8b4157ab69c5.md">Unadvise</a></td>
<td>Cancels a notification registration previously established for an address book entry.</td></tr>
</table>

## See Also


#### Reference
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

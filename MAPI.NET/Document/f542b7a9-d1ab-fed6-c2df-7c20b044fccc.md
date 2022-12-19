# IMessage Interface


Manages messages, attachments, and recipients.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
public interface IMessage : IMAPIProp
```
**VB**
``` VB
Public Interface IMessage
	Inherits IMAPIProp
```
**C++**
``` C++
public interface class IMessage : IMAPIProp
```
**F#**
``` F#
type IMessage = 
    interface
        interface IMAPIProp
    end
```

<table><tr><td><strong>Implements</strong></td><td><a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a></td></tr>
</table>



## Methods
<table>
<tr>
<td><a href="ee81fc2f-a117-6a66-c47d-05642d1e885b.md">CopyProps</a></td>
<td>Copies or moves selected properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="446da6c3-cf56-9eae-0067-556449bcbd5e.md">CopyTo</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="f9599866-c457-5b7f-c73b-25caaa1ee84a.md">CreateAttach</a></td>
<td>Creates a new attachment.</td></tr>
<tr>
<td><a href="8f8c5e18-b544-5f97-06f8-f7cb816e0f3e.md">DeleteAttach</a></td>
<td>Deletes an attachment.</td></tr>
<tr>
<td><a href="de4d890c-a0fc-36d1-40df-acfc7f56bd36.md">DeleteProps</a></td>
<td>Deletes one or more properties from an object.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="16da280c-abbc-efa5-6477-fd4a5f67efc4.md">GetAttachmentTable</a></td>
<td>Returns the message's attachment table.</td></tr>
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
<td><a href="1fdf6ea2-4ee7-da0d-7329-a223aa9dc8dd.md">GetPropList</a></td>
<td>Returns property tags for all properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="eed91d74-f874-f174-2f2d-a0cbf2224590.md">GetProps</a></td>
<td>Retrieves the property value of one or more properties of an object.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="124f5d80-d350-19ef-7211-2bc3d1dd3589.md">GetRecipientTable</a></td>
<td>Returns the message's recipient table.</td></tr>
<tr>
<td><a href="75183b24-bcd0-2273-6cb6-1d59c84b1475.md">ModifyRecipients</a></td>
<td>Adds, deletes, or modifies message recipients.</td></tr>
<tr>
<td><a href="fc0291c1-222b-6840-b9cc-58007da46635.md">OpenAttach</a></td>
<td>Opens an attachment.</td></tr>
<tr>
<td><a href="a82109dc-9148-ad78-11ae-7aa020efd430.md">OpenProperty</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="d26a32e5-3da7-0464-9459-2ad44613db5b.md">SaveChanges</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="f1a2ab65-b81f-ec0c-d947-814cdecceca2.md">SetProps</a></td>
<td>Updates one or more properties.<br />(Inherited from <a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="e638d554-0a69-faec-7e75-13c13631cbba.md">SetReadFlag</a></td>
<td>Sets or clears the MSGFLAG_READ flag in the PR_MESSAGE_FLAGS property of the message and manages the sending of read reports.</td></tr>
<tr>
<td><a href="88f8df9e-d7a9-2f18-0314-eab9655ea303.md">SubmitMessage</a></td>
<td>Saves all of the message's properties and marks the message as ready to be sent.</td></tr>
</table>

## See Also


#### Reference
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  

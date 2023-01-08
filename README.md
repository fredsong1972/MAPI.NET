## MAPI.NET

Messaging API (MAPI) provides the messaging architecture, a set of interfaces, functions, and other data types to facility the development of email messaging. Currently all MAPI wrapper components are unmanaged or dependant on unmanaged MAPI wrapper DLLs. That causes a lot of problems to deploy and support a 64-bit environment. MAPI.NET is a pure .NET Standard 2.0 component based on the MAPI common interface.

### Install the package

#### Latest Version: [![Nuget](https://img.shields.io/nuget/vpre/MAPI.NET.svg)](https://www.nuget.org/packages/MAPI.NET/)  

Install the MAPI.NET with [NuGet](https://www.nuget.org/):
```dotnetcli
dotnet add package MAPI.NET
```

## Getting Started

The following examples demonstrate how to use the `MAPI.NET` in your application.

- Open default message store
    ```csharp
    session = new MAPISession();
    session.Initialize();
    session.OpenMessageStore("");    
    ```

- Register events
    ```csharp
    session.MsgStore.RegisterEvents(EEventMask.fnevNewMail |
                                    EEventMask.fnevObjectCreated |
                                    EEventMask.fnevObjectCopied |
                                    EEventMask.fnevObjectDeleted |
                                    EEventMask.fnevObjectModified |
                                    EEventMask.fnevObjectMoved);
    ```

- Open Message and get recipient
    ```csharp
    MAPIMessage message = MessageStore.GetMAPIObject(e.EntryID) as MAPIMessage;
    Recipient sndr = message.Sender;
    ```

For a complete example see [SpamBounce](https://github.com/fredsong1972/MAPI.NET/tree/main/SpamBounce).

## Articles

Some articles for MAPI.NET include:

- [Managed MAPI(Part 1): Logon MAPI Session and Retrieve Message Stores](https://www.codeproject.com/Articles/455823/Managed-MAPI-Part-1-Logon-MAPI-Session-and-Retriev): 
  How to implement pure .NET MAPI, and use this managed MAPI component to do a little WPF MAPI application. 

- [Managed MAPI (Part 2) – New Mail Notification](https://www.codeproject.com/Articles/490010/Managed-MAPI-Part-2-New-Mail-Notification): 
  How to get new mail notification from a MAPI Message Store.

## Examples

Refer to [`SpamBounce`](https://github.com/fredsong1972/MAPI.NET/tree/main/SpamBounce) for a complete demo.

## Document

[here](https://github.com/fredsong1972/MAPI.NET/blob/main/Help/Home.md)
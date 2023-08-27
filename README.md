# Project Title

Mail Parser

## Description

Parse certain messages from the mail and generate a report in excel <br/>
Parser written to order

## Getting Started

### Dependencies

* .NET 7
* Windows or Linux
  
### Installing

```
git clone https://github.com/adichev17/mail-parser
```
### Setup configuration appSettings.json
* enter your email in the field "ImapUsername"
* setup 2-step verification https://myaccount.google.com/security and get app key
* enter your app key in the field "ImapPassword"
* enter folder for parsing (optional. default: All Mail) in the field "FolderName"

### Executing program

#### Build for Windows
```
dotnet publish -o "{path}" -c Release -r win10-x64 -p:PublishSingleFile=true --self-contained true
```

#### Build for Linux
```
dotnet publish -o "{path}" -c Release -r ubuntu.16.10-x64 -p:PublishSingleFile=true --self-contained true
```

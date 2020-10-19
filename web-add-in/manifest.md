# Research Office Add In (web)

## Settings in Manifest.xml

[required elements](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests?tabs=tabid-1)

### Individual

```xml
  <Version>1.0.0.0</Version>
  <ProviderName>Eye-T</ProviderName>
  <DisplayName DefaultValue="Ceres" />
  <Description DefaultValue="Ceres reservaties"/>
  <SupportUrl DefaultValue="https://eye-t.be/" />
```

### set type of app : "ContentApp", "MailApp" or "TaskPaneApp"

Outlook : altijd MailApp

```xml
<OfficeApp 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
    xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" 
    xsi:type="MailApp">
          ...
</OfficeApp>
```

### Hosts

Outlook = 'Mailbox'

```xml
<Host Name="Mailbox" />
```

### Form

UX settings: DesktopSettings, TabletSettings, and PhoneSettings

xsi:type : "ItemRead" / "ItemEdit"

```xml
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="~remoteAppUrl/CeresAddIn.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
```

### Permissions

Restricted / ReadItem / ReadWriteItem / ReadWriteMailbox

```xml
  <Permissions>ReadWriteItem</Permissions>
```

### Rules

Message / Appointment

```xml
 <Rule xsi:type="ItemIs" ItemType="Apointment" FormType="Read" />
 ```

### VersionOverrides

#### ExtensionPoint

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
```

of

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
```

#### Multiple Extensionpoints

binnen `<DesktopFormFactor>` kunnen meerdere extensionpoints geplaatst worden


[doc](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/manifest/extensionpoint)

---

## Yeoman

[Yo office generator](https://developer.microsoft.com/en-us/office/blogs/creating-office-add-ins-with-any-editor-introducing-yo-office/)

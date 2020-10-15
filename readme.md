# Ceres Outlook calendar add-in

- [Ceres Outlook calendar add-in](#ceres-outlook-calendar-add-in)
  - [Outlook web add-in](#outlook-web-add-in)
  - [Outlook VSTO add-in](#outlook-vsto-add-in)
  - [Verschillen](#verschillen)
  - [Componenten van Office Web Add-in](#componenten-van-office-web-add-in)
    - [Manifest.xml](#manifestxml)
    - [Web app](#web-app)
  - [Extending Office](#extending-office)
    - [Custom buttons en menu commands](#custom-buttons-en-menu-commands)
    - [Task panes](#task-panes)
  - [Extend Outlook](#extend-outlook)
  - [Research](#research)

## Outlook web add-in

[Build your first Outlook add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart?tabs=visualstudio)

## Outlook VSTO add-in

[First VSTO Add-in for Outlook](https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-outlook?view=vs-2019#:~:text=Create%20the%20project,-To%20create%20a&text=In%20the%20templates%20pane%2C%20expand,the%20Name%20box%2C%20type%20FirstOutlookAddIn.)

[How to: Programmatically create appointments](https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-create-appointments?view=vs-2019)

[Codeproject](https://www.codeproject.com/Articles/1112815/How-to-Create-an-Add-in-for-Microsoft-Outlook)

## Verschillen

- ***VSTO***
  - enkel op Windows
  - kan geïnstalleerd worden op lokale pc als assembly
  - enkel voor de desktop versie van office
  - brede toegang tot interne outlook functies
  - C# / VBnet
- ***Web ***
  - Cross platform
  - Loopt in een browser venster die geïnjecteerd wordt in de office toepassing
  - moet gehost worden op een web server, of hosting service
  - zowel voor de desktop als web versie van office
  - beperkte toegang tot interne outlook functies
  - bredere toegang tot data van buiten outlook
  - Javascript / Typescript

## Componenten van Office Web Add-in

### Manifest.xml

- Settings & capabilities van de add-in
  - display name, description, Id, version, default locale
  - Hoe de add-in integreert met Office
  - Permission levels & data access requirements

### Web app

- Static HTML die getoond wordt in de office applicatie
- Interactie met online resources : ASP.NET, PHP, Node.js
- Interactie met Office clients en documenten: [Office.js Javascript APIs](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)

## Extending Office

### Custom buttons en menu commands

- Command buttons kunnen verschillende acties starten
  - Task pane tonen met custom html
  - javascript functie uitvoeren 

### Task panes

- stellen gebruikers in staat om met uw solution te werken
- Clients die geen add-ins ondesteunen (Office 2013 en Office iPad) voeren de add-in als taakvenster
- Launch: My Add-ins button op de Insert tab

## Extend Outlook

- add-ins kunnen Office-app-ribbon uitbreiden
- add-ins kunnen contextueel naast een outlook-item weergeven worden
- email message, meeting request, meeting response, meeting cancellation, or appointment

[Outlook add-ins overview](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins-overview)

[Outlook add-ins commands](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/add-in-commands-for-outlook)

[Github](https://github.com/officedev/outlook-add-in-command-demo)

---

## Research

[Research Web documentatie](webresearch.md)

[Research webhooks](webhooks.md)

[Graph](graph.md)

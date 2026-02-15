# Mail Sender Plugin

This repository contains a [PowerToys Run](https://github.com/microsoft/PowerToys/tree/master/src/modules/launcher) (Wox) plugin that sends notification emails via Outlook and can optionally schedule a Windows notification after a specified delay.

## Features

* `mail start` – send a start notification email.
* `mail stop` – send a stop notification email (can include optional hours info).
* `mail start <timer>` – same as above, but also schedules a Windows toast notification after the given interval. Examples:
  * `mail start 8h` &nbsp;&nbsp;→ 8‑hour timer
  * `mail start 00:35` → 35‑minute timer (`HH:MM` format)
  * `mail start 8` &nbsp;&nbsp;&nbsp;→ interpreted as 8 hours

When the timer expires a toast notification (and, if that fails, a tray balloon tip) will appear. A confirmation toast is shown immediately when you schedule the timer so you know the command was accepted.

## Configuration (`mailconfig.json`)

The plugin reads `mailconfig.json` from the same folder as the plugin DLL. If the file is missing or cannot be parsed, default values are used.

Here is a sample config file with all available properties:

```json
{
  "Recipients": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "SubjectPrefix": "Mail Subject prefix",
  "StartBody": "Mail Body",
  "StopBody": "Mail Body"
}
```

* **Recipients** &ndash; array of email addresses that will receive the message. Separate multiple addresses with commas in Outlook.
* **SubjectPrefix** &ndash; text prepended to every email subject, followed by the current date and the word `start` or `stop`.
* **StartBody** &ndash; body text for the start‑notification email.
* **StopBody** &ndash; body text for the stop‑notification email. When an hours parameter is provided with `mail stop`, it will be appended to this body.

Make sure the JSON file is deployed alongside the plugin. During publish the file is automatically copied due to project settings.

## Building and Deployment

The project targets `net9.0-windows10.0.26100.0` with both WPF and Windows Forms enabled (used for the fallback balloon tip). Build as usual (use the workspace/clone folder you are working in):

```powershell
cd <path-to-plugin-folder>
dotnet publish -c Release -r win-x64 --self-contained false
```

The `dotnet publish` command used in the workspace scripts will produce a self‑contained output in `bin\Release\net9.0-windows10.0.26100.0\win-x64\` containing:

* `plugin.dll` &ndash; the plugin itself
* `plugin.json` &ndash; PowerToys Run manifest
* `mailconfig.json` &ndash; forwarding this file allows easy configuration
* `sendmail.ps1` &ndash; PowerShell helper that actually sends the email via Outlook

Copy the contents of the publish folder into your PowerToys Run plugins directory.

## Troubleshooting

* If notifications do not appear, verify that Windows toast notifications are enabled for "plugin" (you will see a toast when a timer is scheduled).
* Review the PowerToys Run log (`%localappdata%\Microsoft\PowerToys\Logs`) for `MailSenderPlugin` entries – the plugin logs parsing and scheduling details.
* Ensure `sendmail.ps1` is present and that Outlook is configured on the machine.

---

For developer information, see `Main.cs` where the entire command parsing and notification logic resides.
# Outlook-Powershell
A module to use outlook from powershell

## Example: Sending a mail with a table of all your precious windows services

  - Someone asks you for a table of data. 
  - You have it, but how to you just drop it over to that person
  - ... yes, I just pipe it into a mail, awesome idea...

```powershell
Import-Module Outlook

$content = [string](Get-Service | ConvertTo-Html)

# create a new mail, display the window
$content | New-Email -To "<Your Buddies Email>", "<Some-other-poor-guy>" -Subject "The current table of all my services" -Show
```


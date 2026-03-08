---
description: How to version and publish Tabify updates
---
After completing any future fixes for the Tabify Windows application:

1. Always increment the `.csproj` file version and publish the app.
// turbo-all
2. Run the `publish.ps1` script to automate these tasks.
```powershell
.\publish.ps1
```
3. Request the user to test the newly generated executable located in `Tabify\bin\Release\net8.0-windows\win-x64\publish\Tabify.exe`.

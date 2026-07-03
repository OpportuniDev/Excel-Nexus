Attribute VB_Name = "AutoKeyset"
Option Explicit

 Sub Auto_Open()
    Application.OnKey "^u", "ClientUpdater"
    Application.OnKey "^b", "Bootload"
 End Sub


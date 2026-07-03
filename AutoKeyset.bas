Attribute VB_Name = "Module1"
Option Explicit

 Sub Auto_Open()
    Application.OnKey "^u", "ClientUpdater"
    Application.OnKey "^b", "Bootload"
 End Sub


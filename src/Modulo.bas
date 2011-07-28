Attribute VB_Name = "Module1"
Private Declare Function SystemParametersInfo Lib _
        "User32" Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, ByVal uParam As _
        Long, ByVal lpvParam As String, ByVal _
        fuWinIni As Long) As Long

Public Const SPIF_UPDATEINIFILE As Long = &H1
Public Const SPI_SETDESKWALLPAPER As Long = 20
Public Const SPIF_SENDWININICHANGE As Long = &H2
Public msNomeArquivo As String

Public Sub SetWallpaper(ByVal sArquivo As String)
  
  Dim RT As Long
  
  RT = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sArquivo, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
  
End Sub

Private Sub main()

End Sub


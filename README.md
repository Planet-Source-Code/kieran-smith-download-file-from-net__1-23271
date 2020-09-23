<div align="center">

## Download File from Net


</div>

### Description

Downloads a file to the host's computer from the internet via api.
 
### More Info
 
This is useful for downloading files to the users computer for updates of the current program. This is the code i have used for this purpose.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kieran Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kieran-smith.md)
**Level**          |Intermediate
**User Rating**    |4.9 (54 globes from 11 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Excel
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kieran-smith-download-file-from-net__1-23271/archive/master.zip)

### API Declarations

```
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, _
  ByVal szURL As String, _
  ByVal szFileName As String, _
  ByVal dwReserved As Long, _
  ByVal lpfnCB As Long) As Long
```


### Source Code

```
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, _
  ByVal szURL As String, _
  ByVal szFileName As String, _
  ByVal dwReserved As Long, _
  ByVal lpfnCB As Long) As Long
Public Function DownloadFile(URL As String, _
  LocalFilename As String) As Boolean
  Dim lngRetVal As Long
  lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
  If lngRetVal = 0 Then DownloadFile = True
End Function
Private Sub Form_Load()
  DownloadFile "http://www.ksnet.co.uk", "c:\KSNET.htm"
End Sub
```


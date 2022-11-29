# PDF Join

**\* \* \* ALL CREDIT TO ORIGINAL AUTHORS OF PANDOC, WKHTMLTOPDF, IMAGEMAGICK AND POPPLER \* \* \***

PowerShell utility to join Microsoft Office, Markdown (Pandoc), HTML (wkhtmltopdf) and image (ImageMagick) files into a single PDF (Poppler pdfunite). Release includes necessary binaries and DLL's except Microsoft Office which has to be installed on the host.

I use it from the `Send to` menu:

``` PowerShell
$bin ='foo\PDF Join'
$shortcut = $shell.CreateShortcut("$HOME\AppData\Roaming\Microsoft\Windows\SendTo\PDF Join.lnk")
$shortcut.TargetPath = "$PSHOME\pwsh.exe"
$shortcut.Arguments = "-ExecutionPolicy Bypass -File '$bin\PDF join.ps1'"
$shortcut.IconLocation = "$bin\PDF join.ico"
$shortcut.save()
```

and from the command line:

``` Batch File
echo pwsh -File "foo\PDF join.ps1" %* > bar\pdfjoin.bat
```

# REQUIREMENTS
# magick.exe (ImageMagick)
#	magic.xml
# pdfunite.exe (Popplar)
#	freetype6.dll jpeg62.dll libgcc_s_dw2-1.dll libpng16-16.dll
#	libpoppler.dll libstdc++-6.dll libtiff3.dll zlib1.dll
# qpdf.exe
# pandoc.exe
# wkhtmltopdf.exe
#   libwkhtmltox.a
#   wkhtmltox.dll
# Microsoft Word, Excel and PowerPoint have to be installed

Write-Host "Starting up..." -NoNewLine
$i = 0
$errc = 0
$errv = [System.Collections.ArrayList]@()
$tmpa = [System.Collections.ArrayList]@()

Add-Type -AssemblyName PresentationFramework
try { $word_app = New-Object -ComObject Word.Application } catch { Write-Host " Failed to start Word for conversions!"; $errc++ }
$word_ext = ("doc", "docx", "rtf", "txt", "odt", "wps")
try { $ppt_app = New-Object -ComObject Powerpoint.Application } catch { Write-Host " Failed to start PowerPoint for conversions!"; $errc++ }
$ppt_ext = ("ppt", "pptx", "pptm", "potx", "potm", "pps", "ppsx", "ppsm")
try { $xl_app = New-Object -ComObject Excel.Application } catch { Write-Host " Failed to start Excel for conversions!"; $errc++ }
$xl_ext = ("xls", "xlsx", "xlsm", "xlsb", "csv")
$imagemagick_ext = ("aai", "art", "arw", "avi", "avs", "bpg", "bmp", "bmp2", "bmp3", "brf", "cals", "cgm", "cin", "cmyk", "cmyka", "cr2", "crw", "cur", "cut", "dcm", "dcr", "dcx", "dds", "dib", "djvu", "dng", "dot", "dpx", "emf", "epdf", "epi", "eps", "eps2", "eps3", "epsf", "epsi", "ept", "exr", "fax", "fig", "fits", "fpx", "gif", "gplt", "gray", "hdr", "heic", "hpgl", "hrz", "html", "ico", "info", "inline", "isobrl", "isobrl6", "jbig", "jng", "jp2", "jpt", "j2c", "j2k", "jpeg", "jpg", "jxr", "json", "man", "mat", "miff", "mono", "mng", "m2v", "mpeg", "mpc", "mpr", "mrw", "msl", "mtv", "mvg", "nef", "orf", "otb", "p7", "palm", "pam", "clipboard", "pbm", "pcd", "pcds", "pcl", "pcx", "pdb", "pdf", "pef", "pes", "pfa", "pfb", "pfm", "pgm", "picon", "pict", "pix", "png", "png8", "png00", "png24", "png32", "png48", "png64", "pnm", "ppm", "ps", "ps2", "ps3", "psb", "psd", "ptif", "pwp", "rad", "raf", "rgb", "rgba", "rgf", "rla", "rle", "sct", "sfw", "sgi", "shtml", "sid", "mrsid", "sparse-color", "sun", "svg", "text", "tga", "tiff", "tim", "ttf", "txt", "ubrl", "ubrl6", "uil", "uyvy", "vicar", "viff", "wbmp", "wdp", "webp", "wmf", "wpg", "x", "xbm", "xcf", "xpm", "xwd", "x3f", "ycbcr", "ycbcra", "yuv")

$temp_path = [System.IO.Path]::GetTempPath()
[string] $temp_name = [System.Guid]::NewGuid()
$temp = New-Item -ItemType Directory -Path (Join-Path $temp_path $temp_name)
if ($errc -eq 0) { Write-Host " OK" }
else { Write-Host "Startup Unsuccessful!"; $errc = 0 }

foreach ($file in ($args | Sort-Object)) {
	if (!$file) { break; }
	$ext = $file -replace '^(.*)\\([^\\]+)\.([^\.\\]+)$', '$3'
	$path = $file -replace '^(.*)\\([^\\]+)\.([^\.\\]+)$', '$1'
	$file = $file -replace '^(.*)\\([^\\]+)\.([^\.\\]+)$', '$2'
	Write-Host "Adding $file.$ext..." -NoNewLine
	
	try {
		if ($ext -eq "pdf") {
			& "$PSScriptRoot\qpdf.exe" --decrypt "$path\$file.$ext" "$temp\$i.pdf" | Out-Null
			$tmpa.Add("$temp\$i.pdf") | Out-Null
			Write-Host " OK"
		}
		elseif ($ext -eq 'md') {
			& "$PSScriptRoot\pandoc.exe" "$path\$file.$ext" -s --metadata title="$file" -o "$temp\$i.html" | Out-Null
			& "$PSScriptRoot\wkhtmltopdf.exe" -q -s "A4" "$temp\$i.html" "$temp\$i.pdf" | Out-Null
			$tmpa.Add("$temp\$i.pdf") | Out-Null
			Write-Host " OK"
		}
		elseif ($ext -in $word_ext) {
			$word_doc = $word_app.Documents.Open("$path\$file.$ext", $false, $true)
			if ($?) {
				$word_doc.SaveAs([ref] "$temp\$i.pdf", [ref] 17)
				If ($?) {
					$word_doc.Close($false)
					$tmpa.Add("$temp\$i.pdf") | Out-Null
					Write-Host " OK"
				}
				else {
					$errc++
					$errv.Add("$file.$ext") | Out-Null
					Write-Host " Failed to convert Word to PDF!"
				}
			}
			else {
				$errc++
				$errv.Add("$file.$ext") | Out-Null
				Write-Host " Failed to read Word document!"
			}
		}
		elseif ($ext -in $ppt_ext) {
			$ppt_doc = $ppt_app.Presentations.Open("$path\$file.$ext", $false, $true)
			if ($?) {
				$ppt_doc.SaveAs("$temp\$i.pdf", [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
				If ($?) {
					$ppt_doc.Close()
					$tmpa.Add("$temp\$i.pdf") | Out-Null
					Write-Host " OK"
				}
				else {
					$errc++
					$errv.Add("$file.$ext") | Out-Null
					Write-Host " Failed to convert PowerPoint to PDF!"
				}
			}
			else {
				$errc++
				$errv.Add("$file.$ext") | Out-Null
				Write-Host " Failed to read PowerPoint file!"
			}
		}
		elseif ($ext -in $xl_ext) {
			$xl_doc = $xl_app.Workbooks.Open("$path\$file.$ext", 3)
			if ($?) {
				$xl_doc.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.xlFixedFormatType]::xlTypePDF, "$temp\$i.pdf")
				If ($?) {
					$xl_doc.Close()
					$tmpa.Add("$temp\$i.pdf") | Out-Null
					Write-Host " OK"
				}
				else {
					$errc++
					$errv.Add("$file.$ext") | Out-Null
					Write-Host " Failed to convert Excel to PDF!"
				}
			}
			else {
				$errc++
				$errv.Add("$file.$ext") | Out-Null
				Write-Host " Failed to read Excel file!"
			}
		}
		elseif ($ext -in $imagemagick_ext) {
			& "$PSScriptRoot\magick.exe" "$path\$file.$ext" "$temp\$i.pdf"
			$tmpa.Add("$temp\$i.pdf") | Out-Null
			Write-Host " OK"
		}
		else {
			$errc++
			$errv.Add("$file.$ext") | Out-Null
			Write-Host " Failed - unhandled format!"
		}
	}
	catch {

		$errv.Add("$file.$ext") | Out-Null
		$errc++
		Write-Host " Failed, exception thrown: " $_
	}
	
	$i++	
}

Write-Host "Writing output..." -NoNewLine
$tstamp = Get-Date -Format FileDateTime
& "$PSScriptRoot\pdfunite.exe" $tmpa "$path\PDF join - $tstamp.pdf"
Write-Host " OK"

if ($errc) {
	$ofs = "`r`n"
	[System.Windows.MessageBox]::Show("Unable to add $errc files:`r`n$errv")
}

Write-Host "Closing down..." -NoNewLine
$errc = 0
try { Remove-Item -Recurse -Force $temp } catch { Write-Host "`r`nFailed to delete temporary files!"; $errc++ }
try {
	$word_app.Quit() | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word_app) | Out-Null
}
catch { Write-Host "`r`nFailed to close down Word!"; $errc++ }

try {
	$ppt_app.Quit() | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app) | Out-Null
}
catch { Write-Host "`r`nFailed to close down PowerPoint!"; $errc++ }
try {
	$xl_app.Quit() | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl_app) | Out-Null
}
catch { Write-Host "`r`nFailed to close down Excel!"; $errc++ }

if ($errc -eq 0) { Write-Host " OK" }
else { [System.Windows.MessageBox]::Show("Close down Unsuccessful!") }
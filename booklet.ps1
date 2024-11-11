# Create a booklet PDF from a PDF file, ex. a book
#
# Author: david.paiva.fernandes@gmail,com
#
# Date: 2023-11-11

Set-PSDebug -Trace 0

$File = $args[0]
$BookletSize = [int]$args[1]
$Output = $args[2]

if ($args.length -ne 3) {
    write-host "booklet.ps1"
    write-host "-------------------------------------------------"
    write-host "usage:"
    write-host "   booklet.ps1 <filename> <bookletsize> <outputfilename>"
    write-host 
    exit
}

# creates a temp folder to handle files used in the process
$TempPath = Join-Path $Env:Temp $(New-Guid)
New-Item -Type Directory -Path $TempPath | Out-Null

# gets the total number of pages on the original document
$Pages = [int](pdfinfo $File | Select-String -Pattern "(?<=Pages:\s*)\d+").Matches.Value

# finds the size of the last booklet
$LastBookletSize = $Pages % $BookletSize

# in order to calc the additional blabnk pages needed in the end
$AdditionalBlankPages = 0

if ($LastBookletSize -ne 0) {
	$AdditionalBlankPages = $BookletSize - $LastBookletSize
}

$TotalPages = $Pages + $AdditionalBlankPages
$Booklets = $TotalPages / $BookletSize

$BlankFile = Join-Path $TempPath "blank.pdf"

write-host "TempDir............: ", $TempPath
write-host "BookletSize........: ", $BookletSize
write-host "File...............: ", $File
write-host "Pages..............: ", $Pages
write-host "LastBookletSize....: ", $LastBookletSize
write-host "AdditionalPage(s)..: ", $AdditionalBlankPages
write-host "TotalPage(s).......: ", $TotalPages
write-host "Booklet(s).........: ", $Booklets

write-host "Blank file.........: ", $BlankFile

write-host ""
write-host "PROCESSING:"
write-host ""
write-host "Creating blank pages"
write-host "--------------------"

# creates the necessary blank pages
echo showpage | ps2pdf -sPAPERSIZE=a4 - $BlankFile

for (($i = 0); $i -lt $AdditionalBlankPages; $i++)
{
	$P = ($Pages+$i+1)
	write-host $P
	$NewPage = Join-Path $TempPath ("Pages_" + ($P | % tostring 0000) + ".pdf")
	Copy-Item $BlankFile -Destination $NewPage
}

write-host ""
write-host "Extracting pages"
write-host "--------------------"

# extract all pages of the original document to individual pdf files.
pdftk $File burst output (Join-Path $TempPath "Pages_%04d.pdf")
Remove-Item (Join-Path $TempPath doc_data.txt)

# creates each Booklet
$BookletNames = ""

for (($b = 1); $b -le $Booklets; $b++)
{
	$List = ""

	for (($i = 1); $i -le ($BookletSize / 4); $i++)
	{
		$x = ($b - 1) * $BookletSize
		$p1 = $x + $i * 2
		$p2 = $x + $BookletSize - (($i-1) * 2) - 1
		$p3 = $p2 + 1
		$p4 = $p1 - 1
		$List = $List + " " + (Join-Path $TempPath ($p1 | % tostring Pages_0000\.pdf))  + " " + (Join-Path $TempPath ($p2 | % tostring Pages_0000\.pdf)) + " " + (Join-Path $TempPath ($p3 | % tostring Pages_0000\.pdf)) + " " + (Join-Path $TempPath ($p4 | % tostring Pages_0000\.pdf)) + " "
	}

	$List = $List -replace '\s+', ' '
	
	$NewBooklet = Join-Path $TempPath ("Booklet_" + $b + ".pdf")
	write-host "New Booklet: " $NewBooklet

	$BookletNames = $BookletNames + " " + $NewBooklet

	# creates the Booklet pdf
    pdftk $List.Split(" ") cat output $NewBooklet
	
}

$BookletNames = $BookletNames -replace '\s+', ' '

# join sall Booklets in a single PDF file.
pdftk $BookletNames.Split(" ") cat output $Output

Remove-Item (Join-Path $TempPath Pages_*.pdf)
Remove-Item (Join-Path $TempPath Booklet_*.pdf)
Remove-Item (Join-Path $TempPath blank.pdf)
Remove-Item $TempPath

write-host "File created: " $Output
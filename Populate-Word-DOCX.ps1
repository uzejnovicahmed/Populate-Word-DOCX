<#
.SYNOPSIS
	Populate Microsoft Word DOCX
.DESCRIPTION
    Replace [placeholder] tokens with dynamic text.  DOCX file is ZIP format.  
    Extract all contents, replace ASCII text, and bundle into fresh new ZIP(DOCX) output.

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Namespace	: Populate-Word-DOCX.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.10
	Last Modified	: 09-19-2017
.LINK
	Source Code
	http://www.github.com/spjeff/Populate-Word-DOCX
#>

# params
$template = "c:\temp\template.docx"
$tempFolder = $env:TEMP + "\Populate-Word-DOCX"

# unzip function
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    param([string]$zipfile, [string]$outpath)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}
function Zip {
    param([string]$folderInclude, [string]$outZip)
    [System.IO.Compression.CompressionLevel]$compression = "Optimal"
    $ziparchive = [System.IO.Compression.ZipFile]::Open( $outZip, "Update" )

    # loop all child files
    $realtiveTempFolder = (Resolve-Path $tempFolder -Relative).TrimStart(".\")
    foreach ($file in (Get-ChildItem $folderInclude -Recurse)) {
        # skip directories
        if ($file.GetType().ToString() -ne "System.IO.DirectoryInfo") {
            # relative path
            $relpath = ""
            if ($file.FullName) {
                $relpath = (Resolve-Path $file.FullName -Relative)
            }
            if (!$relpath) {
                $relpath = $file.Name
            } else {
                $relpath = $relpath.Replace($realtiveTempFolder, "")
                $relpath = $relpath.TrimStart(".\").TrimStart("\\")
            }

            # display
            Write-Host $relpath -Fore Green
            Write-Host $file.FullName -Fore Yellow

            # add file
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ziparchive, $file.FullName, $relpath, $compression) | Out-Null
        }
    }
    $ziparchive.Dispose()
}

# prepare folder
Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null
mkdir $tempFolder | Out-Null

# unzip DOCX
Unzip $template $tempFolder

# replace text
$bodyFile = $tempFolder + "\word\document.xml"
$body = Get-Content $bodyFile
$body = $body.Replace("[placeholder1]", "hello")
$body = $body.Replace("[placeholder2]", "world")
$body | Out-File $bodyFile -Force -Encoding ascii

# zip DOCX
$destfile = $template.Replace(".docx", "-after.docx")
Remove-Item $destfile -Force -ErrorAction SilentlyContinue
Zip $tempFolder $destfile

# clean folder
Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null
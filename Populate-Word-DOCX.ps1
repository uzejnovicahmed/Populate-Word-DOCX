

$template = "c:\Path\to\template\template.docx"
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
# Replace text while preserving the original encoding


$bodyFile = $tempFolder + "\word\document.xml"
$body = Get-Content $bodyFile -raw -Encoding UTF8
$body = $body.Replace("XXPlaceholderXX", "$($Userobject.Anschrift)")
$body = $body.Replace("XXPlaceholder2XX", "$($Userobject.Dienstnehmer)")


# Save back using StreamWriter to preserve encoding

# Save back using StreamWriter to preserve encoding
$streamWriter = [System.IO.StreamWriter]::new($bodyFile, $false, [System.Text.Encoding]::UTF8)
$streamWriter.Write($body)
$streamWriter.Close()


# zip DOCX
$destfile = $template.Replace(".docx", "-after.docx")
Remove-Item $destfile -Force -ErrorAction SilentlyContinue
Zip $tempFolder $destfile

# clean folder
Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null

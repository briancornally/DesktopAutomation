$NoteBook="OneNoteDocumentation"
$ExportDir="c:\OneNoteDocumentation"

Get-ChildItem -Recurse -Path $ExportDir\phase1\* -ErrorAction Stop | Remove-Item -Force -Recurse

set-alias Pandoc "C:\ProgramData\chocolatey\bin\pandoc.exe"
$word = new-object -comobject word.application
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat],”wdFormatDocumentDefault”);
#[ref]$SaveFormat = “microsoft.office.interop.word.WdSaveFormat” -as [type]
#$word.Visible = $False
$Sections=Get-OneNoteHierarchy OneNote:\$NoteBook -scope hsSections | % Section | ? Name -eq "phase1" | % Name
foreach ($Section in $Sections) {
    "Section:$Section"
    New-Item -ItemType Directory -Path "$ExportDir\$Section" -Force -ErrorAction SilentlyContinue| Out-Null
    $PageNum=1
#    $Pages=Get-OneNoteHierarchy OneNote:\$NoteBook\$Section | % Page | ? Name -match "Dell" | % Name
    $Pages=Get-OneNoteHierarchy OneNote:\$NoteBook\$Section | % Page | % Name
    $OutputPath="$ExportDir\$Section"
    Set-Location $OutputPath
    foreach ($Page in $Pages) {
        $PageNumStr="{0:D2}" -f $PageNum
        $FileName="$PageNumStr. $Page"
        $FilePath="$OutputPath\$FileName"
        Remove-Item "$FilePath.doc" -ErrorAction SilentlyContinue | Out-Null
        "Page:$FileName"
        Export-OneNote OneNote:\$NoteBook\$Section\$Page -Name $FileName -OutputPath $OutputPath -Format doc -Force -Confirm:$false | % ExportedFile
        $opendoc = $word.documents.open("$FilePath.doc")
        $opendoc.Convert()
        $opendoc.saveas([ref]"$FilePath.docx", [ref]$saveFormat);
        $opendoc.close();
        $ArgumentList=@"
        --extract-media=$PageNumStr "$FileName.docx" -o "$FileName.md" --to markdown_github
"@
        $ArgumentList
        start-process -NoNewWindow -FilePath Pandoc -ArgumentList $ArgumentList
        $PageNum++        
    }
}
$word.quit()
Get-ChildItem -Recurse -Path $ExportDir\* -Include *.doc* | Remove-Item -Force

<#
function ZipFiles( $zipfilename, $sourcedir )
{
   Add-Type -Assembly System.IO.Compression.FileSystem
   $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
   [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcedir,
        $zipfilename, $compressionLevel, $false)
}

Remove-Item $ZipFile -ErrorAction SilentlyContinue
ZipFiles $ZipFile $ExportDir

#>
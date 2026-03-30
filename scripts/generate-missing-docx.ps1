[CmdletBinding()]
param(
    [switch]$Validate
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$outputDir = Join-Path $root 'samples\generated-docx'
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-MinimalDocx {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Text
    )

    if (Test-Path $Path) {
        Remove-Item $Path -Force
    }

    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            $contentTypes = $zip.CreateEntry('[Content_Types].xml')
            $writer = New-Object System.IO.StreamWriter($contentTypes.Open())
            $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
            $writer.Dispose()

            $rootRels = $zip.CreateEntry('_rels/.rels')
            $writer = New-Object System.IO.StreamWriter($rootRels.Open())
            $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
            $writer.Dispose()

            $doc = $zip.CreateEntry('word/document.xml')
            $writer = New-Object System.IO.StreamWriter($doc.Open())
            $safeText = [System.Security.SecurityElement]::Escape($Text)
            $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>')
            $writer.Write($safeText)
            $writer.Write('</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')
            $writer.Dispose()

            $docRels = $zip.CreateEntry('word/_rels/document.xml.rels')
            $writer = New-Object System.IO.StreamWriter($docRels.Open())
            $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
            $writer.Dispose()
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

$defaultFileNames = @(
    'input.docx',
    'source.docx',
    'test.docx'
)

$testsDir = Join-Path $root 'src\Nedev.FileConverters.DocxToDoc.Tests'
$discoveredFileNames = @()

if (Test-Path $testsDir) {
    $matches = Get-ChildItem -Path $testsDir -Recurse -Filter '*.cs' -File |
        Select-String -Pattern '"([^"]+\.docx)"' -AllMatches

    foreach ($match in $matches) {
        foreach ($group in $match.Matches) {
            $candidate = $group.Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($candidate)) {
                continue
            }

            $candidate = $candidate.Replace('/', '\')
            $leaf = Split-Path -Path $candidate -Leaf
            if ([string]::IsNullOrWhiteSpace($leaf)) {
                continue
            }

            if ($leaf -notlike '*.docx') {
                continue
            }

            if ($leaf -match '[\{\}\$]') {
                continue
            }

            $discoveredFileNames += $leaf.ToLowerInvariant()
        }
    }
}

$fileNames = @($defaultFileNames + $discoveredFileNames) | Sort-Object -Unique

foreach ($name in $fileNames) {
    $path = Join-Path $outputDir $name
    New-MinimalDocx -Path $path -Text ("generated sample: " + $name)
}

Get-ChildItem -Path $outputDir -Filter '*.docx' | Sort-Object Name | Select-Object Name, Length

if ($Validate) {
    $solutionPath = Join-Path $root 'Nedev.FileConverters.DocxToDoc.sln'
    $cliDll = Join-Path $root 'src\Nedev.FileConverters.DocxToDoc.Cli\bin\Release\net8.0\Nedev.FileConverters.DocxToDoc.Cli.dll'
    $convertedDir = Join-Path $root 'samples\generated-doc'

    if (-not (Test-Path $cliDll)) {
        dotnet build $solutionPath -c Release | Out-Host
    }

    if (Test-Path $convertedDir) {
        Remove-Item -Path $convertedDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $convertedDir -Force | Out-Null

    dotnet $cliDll -b $outputDir -o $convertedDir -v | Out-Host

    $docxCount = (Get-ChildItem -Path $outputDir -Filter '*.docx' | Measure-Object).Count
    $docCount = (Get-ChildItem -Path $convertedDir -Filter '*.doc' | Measure-Object).Count
    if ($docCount -lt $docxCount) {
        throw "Validation failed: expected $docxCount converted .doc files, got $docCount."
    }

    Get-ChildItem -Path $convertedDir -Filter '*.doc' | Sort-Object Name | Select-Object Name, Length
}

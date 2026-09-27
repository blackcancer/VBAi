param(
    [Parameter(Mandatory = $true)] [string] $CoberturaPath,
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $CoberturaPath).Path
[xml] $coverage = Get-Content -LiteralPath $resolved
$rows = foreach ($class in $coverage.coverage.packages.package.classes.class) {
    foreach ($method in $class.methods.method) {
        $lines = @($method.lines.line)
        if ($lines.Count -eq 0) { continue }
        $missing = @($lines | Where-Object { [int] $_.hits -eq 0 })
        [pscustomobject]@{
            File = [string] $class.filename
            Class = [string] $class.name
            Method = [string] $method.name
            FirstLine = [int] $method.line
            Lines = $lines.Count
            MissingLines = $missing.Count
            MissingLineNumbers = ($missing | ForEach-Object { [string] $_.number }) -join ','
            LineRate = [double]::Parse($method.GetAttribute('line-rate'),
                [Globalization.CultureInfo]::InvariantCulture)
            BranchRate = [double]::Parse($method.GetAttribute('branch-rate'),
                [Globalization.CultureInfo]::InvariantCulture)
        }
    }
}
$rows = @($rows | Sort-Object File, Class, FirstLine, Method)
if ($OutputPath) {
    $target = [IO.Path]::GetFullPath($OutputPath)
    $directory = [IO.Path]::GetDirectoryName($target)
    if (-not [IO.Directory]::Exists($directory)) { [IO.Directory]::CreateDirectory($directory) | Out-Null }
    $rows | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding UTF8
}
$rows

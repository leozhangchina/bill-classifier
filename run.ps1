param(
  [Parameter(Mandatory = $true)]
  [string]$InputFile,
  [string]$OutputFile,
  [string]$CategoriesFile,
  [string]$OverridesFile,
  [string]$Sheet,
  [string]$PdfPassword,
  [double]$Threshold = 0.82
)

if (-not $CategoriesFile) { $CategoriesFile = Join-Path $PSScriptRoot 'categories.default.json' }
if (-not $OverridesFile) { $OverridesFile = Join-Path $PSScriptRoot 'merchant-overrides.json' }

$pythonExecutable = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
$pythonPrefix = @()
if (-not (Test-Path -LiteralPath $pythonExecutable)) {
  $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
  if ($pythonCommand) {
    $pythonExecutable = $pythonCommand.Source
  } else {
    $pyCommand = Get-Command py -ErrorAction SilentlyContinue
    if (-not $pyCommand) {
      throw 'Python was not found. Install Python 3.10 or newer and follow README.md.'
    }
    $pythonExecutable = $pyCommand.Source
    $pythonPrefix = @('-3')
  }
}

$arguments = $pythonPrefix + @(
  (Join-Path $PSScriptRoot 'classify_bill_light.py'),
  $InputFile,
  '--categories', $CategoriesFile,
  '--overrides', $OverridesFile,
  '--threshold', $Threshold.ToString([System.Globalization.CultureInfo]::InvariantCulture)
)
if ($OutputFile) { $arguments += @('--output', $OutputFile) }
if ($Sheet) { $arguments += @('--sheet', $Sheet) }
if ($PdfPassword) { $arguments += @('--pdf-password', $PdfPassword) }

& $pythonExecutable @arguments
exit $LASTEXITCODE

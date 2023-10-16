# Git config
$gitRepo = "https://github.com/adrivn/ws"
# The path to your Python virtual environment
$venvPath = "$HOME\Envs\main"

# The paths to your Python scripts
$scriptName = "start.py"
$scriptPath = $MyInvocation.MyCommand.Path
$rootPath = Split-Path $scriptPath -Parent

# Update repo before running
Write-Host "Fetching updates..."
git pull

# Check if virtual environment exists, and creates one if not
if (Test-Path $venvPath\Scripts\Activate.ps1 -PathType Leaf) {
  Write-Host "Found virtual environment in $venvPath"
  # Activate the virtual environment
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1
}
else {
  Write-Host "No virtual environment found. Creating a new one in: $venvPath"
  python -m venv $venvPath 2>$null
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1
  & pip install -r $rootPath/conf/requirements.txt 2>$null
}

# Check if the virtual environment is activated
if ($env:VIRTUAL_ENV -eq $null) {
    Write-Host "Failed to activate Python virtual environment. Exiting..."
    exit
}

# Upgrade packages based on requirements.txt
Write-Host "Upgrading packages based on requirements.txt..."
& pip install --upgrade -r $rootPath/conf/requirements.txt 2>$null

# Check if the upgrade was successful
if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to upgrade packages based on requirements.txt. Exiting..."
    exit
}

# Run the script
Write-Host "Running Python script at $rootPath..."
python $rootPath/$scriptName
if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to execute Python script at $scriptPath. Exiting..."
    exit
}

Write-Host "All scripts executed successfully. Exiting..."


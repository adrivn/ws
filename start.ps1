# Git config
$gitRepo = "https://github.com/adrivn/ws"
# The path to your Python virtual environment
$venvPath = "$HOME\Envs\main"

# The paths to your Python scripts
$scriptPath = "./start.py"

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
  python -m venv $venvPath
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1
  & pip install -r conf/requirements.txt
}

# Check if the virtual environment is activated
if ($env:VIRTUAL_ENV -eq $null) {
    Write-Host "Failed to activate Python virtual environment. Exiting..."
    exit
}

# Loop through each script and execute it
Write-Host "Running Python script at $scriptPath..."
python $scriptPath
if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to execute Python script at $scriptPath. Exiting..."
    exit
}

Write-Host "All scripts executed successfully. Exiting..."


# The path to your Python virtual environment
$venvPath = "$HOME\Envs\dataviz"

# Check if virtual environment exists, and creates one if not
if (Test-Path $venvPath\Scripts\Activate.ps1 -PathType Leaf) {
  Write-Host "Found virtual environment in $venvPath"
  # Activate the virtual environment
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1
}
else {
  Write-Host "No virtual environment found. Please install all requirements before continuing."
  exit
}

# Loop through each script and execute it
Write-Host "Starting Superset (dashboard tool)"
& superset run -p 8088 --with-threads --reload --debugger

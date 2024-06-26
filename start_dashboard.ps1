# The path to your Python virtual environment
$venvPath = "$HOME\Envs\dataviz"
$scriptPath = $MyInvocation.MyCommand.Path
$rootPath = Split-Path $scriptPath -Parent
$boolUpdate = $true

Write-Host "Fetching updates..."
& cp $rootPath/conf/superset_config.py $venvPath/superset_config.py
& cd $rootPath
& git pull

# Check if virtual environment exists, and creates one if not
if (Test-Path $venvPath\Scripts\Activate.ps1 -PathType Leaf) {
  Write-Host "Found virtual environment in $venvPath"
  # Activate the virtual environment
  Write-Host "Activating Python virtual environment..."
  & $venvPath\Scripts\Activate.ps1 2>$null
}
else {
  Write-Host "No virtual environment found. Please install all requirements before continuing."
  exit
}

if ($boolUpdate) {
# Upgrade packages based on requirements.txt
  Write-Host "Upgrading packages based on requirements.txt..."
  & pip install --upgrade -r $rootPath/conf/supersetreqs.txt 2>$null
  Write-Host "Setting up database..."
  & cd $venvPath & cd Scripts & superset db upgrade & superset init
# Check if the upgrade was successful
  if ($LASTEXITCODE -ne 0) {
      Write-Host "Failed to upgrade packages based on requirements.txt. Exiting..."
      exit
  }
}


# Loop through each script and execute it
Write-Host "Starting Superset (dashboard tool)"
& superset run -p 8088 --with-threads --reload --debugger

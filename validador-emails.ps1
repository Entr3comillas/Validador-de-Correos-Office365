# Instalar el módulo de AzureAD para interactuar con Office365 si no está instalado
if (-not (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module -Name AzureAD -Force -AllowClobber
}

# Importar el módulo de AzureAD
Import-Module AzureAD

# Función para verificar si un correo electrónico es válido en Office365
function Test-ValidEmail {
    param (
        [string]$Email
    )
    
    try {
        # Intentar iniciar sesión con el correo electrónico proporcionado
        $null = Get-AzureADUser -ObjectId $Email
        Write-Output "$Email es un correo electrónico válido en Office365"
    } catch {
        Write-Output "$Email no es un correo electrónico válido en Office365"
    }
}

# Función para leer una lista de correos electrónicos desde un archivo
function Get-EmailListFromFile {
    param (
        [string]$FilePath
    )

    # Verificar si el archivo existe
    if (Test-Path $FilePath) {
        # Leer los correos electrónicos desde el archivo y retornarlos como una lista
        return Get-Content $FilePath
    } else {
        Write-Output "El archivo $FilePath no existe."
        exit 1
    }
}

# Ruta del archivo que contiene la lista de correos electrónicos
$EmailListFile = "C:\Ruta\Hacia\Tu\Archivo.txt"

# Verificar si el archivo de lista de correos electrónicos existe
if (Test-Path $EmailListFile) {
    # Obtener la lista de correos electrónicos desde el archivo
    $Emails = Get-EmailListFromFile -FilePath $EmailListFile

    # Iterar sobre cada correo electrónico y verificar su validez en Office365
    foreach ($Email in $Emails) {
        Test-ValidEmail -Email $Email
    }
} else {
    Write-Output "El archivo de lista de correos electrónicos $EmailListFile no existe."
}

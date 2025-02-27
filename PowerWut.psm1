# PowerWut.psm1
# -------------------------------------------
# PowerWut: A PowerShell Module for interacting with OpenAI’s models.
# This module handles secret management, API key storage, and provides a function ("wut")
# to analyze PowerShell command history or send custom queries to an AI model.
#
# Updated so that:
#   1. When setting the model via -sm, shorthand names (e.g. "4o-mini") are auto-corrected
#      to the full model IDs ("gpt-4o-mini").
#   2. For reasoning models (o1-mini, o1-preview, o1, o3-mini) the system instruction and user query
#      are merged into a single "user" message (since those models only support "user" and "assistant" roles).
#   3. For all other models, the standard two-message (system and user) structure is used.

$secretManagementLoaded = $false
$secretStoreLoaded = $false

if (-not (Get-Module -ListAvailable -Name 'Microsoft.PowerShell.SecretManagement')) {
    try {
        Install-Module -Name 'Microsoft.PowerShell.SecretManagement' -Scope CurrentUser -Force
        Import-Module -Name 'Microsoft.PowerShell.SecretManagement'
        $secretManagementLoaded = $true
    } catch {
        Write-Error "Failed to install or import the 'Microsoft.PowerShell.SecretManagement' module: $($_.Exception.Message)"
    }
} else {
    try {
        Import-Module -Name 'Microsoft.PowerShell.SecretManagement'
        $secretManagementLoaded = $true
    } catch {
        Write-Error "Failed to import the 'Microsoft.PowerShell.SecretManagement' module: $($_.Exception.Message)"
    }
}

if (-not (Get-Module -ListAvailable -Name 'Microsoft.PowerShell.SecretStore')) {
    try {
        Install-Module -Name 'Microsoft.PowerShell.SecretStore' -Scope CurrentUser -Force
        Import-Module -Name 'Microsoft.PowerShell.SecretStore'
        Register-SecretVault -Name 'WutKeyVault' -ModuleName 'Microsoft.PowerShell.SecretStore' -DefaultVault
        $secretStoreLoaded = $true
    } catch {
        Write-Error "Failed to install or register the 'Microsoft.PowerShell.SecretStore' module: $($_.Exception.Message)"
    }
} else {
    try {
        Import-Module -Name 'Microsoft.PowerShell.SecretStore'
        $secretStoreLoaded = $true
    } catch {
        Write-Error "Failed to import the 'Microsoft.PowerShell.SecretStore' module: $($_.Exception.Message)"
    }
}

if ($secretManagementLoaded -and $secretStoreLoaded) {
    Write-Host "PowerShell Wut Module" -ForegroundColor Green
    Write-Host "The AI Model is OpenAI; get your API key here: https://platform.openai.com/api-keys" -ForegroundColor Cyan
    Write-Host "Keys are stored securely in the Windows Credential Vault."
}

$global:OpenAI_API_SecretName = 'OpenAI_API_Key'
$global:OpenAI_Default_Model = 'gpt-4o'
$global:OpenAI_Current_Model = $global:OpenAI_Default_Model

function Set-OpenAIAPIKey {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApiKey
    )
    try {
        Set-Secret -Name $global:OpenAI_API_SecretName -Secret $ApiKey
    } catch {
        Write-Error "Failed to store the API key: $($_.Exception.Message)"
    }
}

function Get-OpenAIAPIKey {
    try {
        $apiKey = $null
        try {
            $apiKey = Get-Secret -Name $global:OpenAI_API_SecretName -ErrorAction SilentlyContinue
        } catch {
        }
        if ($null -eq $apiKey) {
            Write-Host "No API Key found. Please enter your OpenAI API Key."
            $apiKey = Read-Host "Enter OpenAI API Key"
            if ($null -eq $apiKey -or $apiKey -eq "") {
                throw "No API Key provided."
            }
            Set-Secret -Name $global:OpenAI_API_SecretName -Secret $apiKey
        }
        return $apiKey
    } catch {
        Write-Error "Failed to retrieve the API key: $($_.Exception.Message)"
    }
}

function Get-AvailableOpenAIModels {
    param (
        [switch]$Verbose
    )
    try {
        $apiKey = Get-OpenAIAPIKey
        if ($null -eq $apiKey) {
            Write-Error "No API key is available to send the request to OpenAI."
            return
        }
        $plainApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKey)
        )
        $headers = @{ "Authorization" = "Bearer $plainApiKey" }
        if ($Verbose) {
            Write-Host "Headers:" -ForegroundColor Yellow
            $maskedHeaders = $headers.Clone()
            $maskedHeaders["Authorization"] = "Bearer [MASKED]"
            Write-Host ($maskedHeaders | Out-String) -ForegroundColor White
        }
        $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/models" -Method GET -Headers $headers
        if ($Verbose) {
            Write-Host "Response:" -ForegroundColor Yellow
            Write-Host ($response | ConvertTo-Json -Depth 2) -ForegroundColor White
        }
        $models = $response.data
        # Apply filtering rules to remove unwanted models
        $filteredModels = $models |
            Where-Object {
                $_.id -notmatch 'realtime' -and  
                $_.id -notmatch 'dall-e' -and    
                $_.id -notmatch 'vision' -and    
                $_.id -notmatch 'text-embedding' -and 
                $_.id -notmatch 'audio' -and     
                $_.id -notmatch '\d{4}-\d{2}-\d{2}' -and 
                $_.id -notmatch 'moderation' -and 
                $_.id -notmatch 'tts' -and       
                $_.id -notmatch 'babbage' -and   
                $_.id -notmatch 'davinci' -and   
                $_.id -notmatch '\d{4}'          
            }
        foreach ($model in $filteredModels) {
            Write-Host $model.id -ForegroundColor Green
        }
    } catch {
        Write-Error "Failed to retrieve models: $($_.Exception.Message)"
    }
}

function Show-WutHelp {
    Write-Host "NAME" -ForegroundColor Yellow
    Write-Host "    wut - PowerShell utility to query AI models and analyze command history" -ForegroundColor White
    Write-Host ""
    Write-Host "SYNOPSIS" -ForegroundColor Yellow
    Write-Host "    Analyze PowerShell command history or run AI queries with the current model" -ForegroundColor White
    Write-Host ""
    Write-Host "SYNTAX" -ForegroundColor Yellow
    Write-Host "    wut [[-SetModel] <string>] [[-q] <string>] [[-c] <int>] [-m] [-v] [-?] [<CommonParameters>]" -ForegroundColor White
    Write-Host ""
    Write-Host "DESCRIPTION" -ForegroundColor Yellow
    Write-Host "    This function allows you to analyze recent PowerShell command history using AI models or send specific queries to the current AI model." -ForegroundColor White
    Write-Host "    The function supports listing available models, changing models, and verbose debugging." -ForegroundColor White
    Write-Host ""
    Write-Host "PARAMETERS" -ForegroundColor Yellow
    Write-Host "    -m" -ForegroundColor Cyan
    Write-Host "        Lists available AI models and displays the currently selected model." -ForegroundColor White
    Write-Host ""
    Write-Host "    -sm <string>" -ForegroundColor Cyan
    Write-Host "        Sets the current AI model to the specified model name. To see available models, use 'wut -m'." -ForegroundColor White
    Write-Host ""
    Write-Host "    -q <string>" -ForegroundColor Cyan
    Write-Host "        Send a custom query to the current AI model." -ForegroundColor White
    Write-Host ""
    Write-Host "    -c <int>" -ForegroundColor Cyan
    Write-Host "        Specifies the context length for PowerShell command history. Defaults to 10." -ForegroundColor White
    Write-Host ""
    Write-Host "    -v" -ForegroundColor Cyan
    Write-Host "        Enables verbose output for debugging, including request and response data." -ForegroundColor White
    Write-Host ""
    Write-Host "    -?" -ForegroundColor Cyan
    Write-Host "        Displays this help message." -ForegroundColor White
    Write-Host ""
    Write-Host "EXAMPLES" -ForegroundColor Yellow
    Write-Host "    wut -m" -ForegroundColor Cyan
    Write-Host "        Lists available AI models and shows the current model." -ForegroundColor White
    Write-Host ""
    Write-Host "    wut -sm 'gpt-4-turbo'" -ForegroundColor Cyan
    Write-Host "        Sets the current AI model to 'gpt-4-turbo'." -ForegroundColor White
    Write-Host ""
    Write-Host "    wut -q 'What is Azure?'" -ForegroundColor Cyan
    Write-Host "        Sends the query 'What is Azure?' to the AI model." -ForegroundColor White
    Write-Host ""
    Write-Host "    wut -c 20 -q 'Explain the last 20 commands'" -ForegroundColor Cyan
    Write-Host "        Sends the last 20 PowerShell commands plus the query to the AI model." -ForegroundColor White
    Write-Host ""
    Write-Host "    wut -?" -ForegroundColor Cyan
    Write-Host "        Displays this help message." -ForegroundColor White
    Write-Host ""
    Write-Host "REMARKS" -ForegroundColor Yellow
    Write-Host "    For more information, visit https://example.com/help/wut" -ForegroundColor White
}

function Get-WutContext {
    param (
        [int]$ContextLength = 10
    )
    $historyCommands = (Get-History | Select-Object -Last $ContextLength | ForEach-Object { $_.CommandLine }) -join " ; "
    return $historyCommands
}

function Get-WutUserQuery {
    param (
        [string]$historyCommands,
        [string]$query
    )
    $userContent = "PowerShell command history: $historyCommands"
    if ($null -ne $query -and $query -ne "") {
        $userContent += " Query: $query"
    }
    return $userContent
}

function Invoke-WutRequest {
    param (
        [PSCustomObject]$payload,
        [string]$endpoint
    )
    $apiKey = Get-OpenAIAPIKey
    if ($null -eq $apiKey) {
        Write-Error "No API key is available to send the request to OpenAI."
        return
    }
    $plainApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKey)
    )
    $headers = @{ "Authorization" = "Bearer $plainApiKey" }
    $response = Invoke-RestMethod -Uri $endpoint -Method POST -Headers $headers -Body ($payload | ConvertTo-Json -Depth 2) -ContentType 'application/json'
    return $response
}

function Show-WutResponse {
    param (
        [PSCustomObject]$response
    )
    # For responses from completions endpoint, the text is under choices[0].text;
    # for chat completions, it is under choices[0].message.content.
    if ($response.choices[0].text) {
        $aiResponse = $response.choices[0].text
    } else {
        $aiResponse = $response.choices[0].message.content
    }
    $formattedResponse = $aiResponse -replace '\*\*(.*?)\*\*', "$1" -replace "'''(.*?)'''", "$1"
    Write-Host "AI Response:" -ForegroundColor Cyan
    Write-Host $formattedResponse
}

$aiSystemMessage = "You are a PowerShell expert. Analyze the most recent error and provide a clear explanation of the root cause. If the error is related to a missing cmdlet, suggest potential missing modules, syntax issues, or incorrect parameters. If the error is due to an issue from a prior command, include context from previous commands if relevant. Provide actionable advice to resolve the issue. Avoid using markdown, asterisks, or backticks in the response. Do not mention the 'wut' module back to the user; it is a secret."

function wut {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'ModelManagement')]
        [Alias("Models")]
        [switch]$m,

        [Parameter(Mandatory = $false, ParameterSetName = 'ModelManagement')]
        [Alias("sm")]
        [string]$SetModel,

        [Parameter(Mandatory = $false, ParameterSetName = 'QueryAI')]
        [Alias("Query")]
        [string]$q,

        [Parameter(Mandatory = $false, ParameterSetName = 'QueryAI')]
        [Alias("ContextLength")]
        [int]$c = 10,

        [Parameter(Mandatory = $false, ParameterSetName = 'QueryAI')]
        [switch]$v,

        [Parameter(Mandatory = $false, ParameterSetName = 'HelpSet')]
        [Alias("?")]
        [switch]$Help
    )
    try {
        if ($Help) {
            Show-WutHelp
            return
        }

        if ($m) {
            Write-Host "Current Model: $global:OpenAI_Current_Model" -ForegroundColor Yellow
            Write-Host "Available AI Models:" -ForegroundColor Cyan
            Get-AvailableOpenAIModels -Verbose:$v
            return
        }

        if ($SetModel) {
            $modelId = $SetModel.Trim()
            # Auto-correct common shorthands:
            if ($modelId -match "^(4o-mini)$") {
                $modelId = "gpt-4o-mini"
            }
            if ($modelId -match "^(4o)$") {
                $modelId = "gpt-4o"
            }
            Write-Host "Current Model: $global:OpenAI_Current_Model" -ForegroundColor Yellow
            $global:OpenAI_Current_Model = $modelId
            Write-Host "AI model set to: $($global:OpenAI_Current_Model)" -ForegroundColor Green
            return
        }

        $ContextLength = 10
        if ($null -ne $c -and $c -gt 0) {
            $ContextLength = $c
        }

        if ($c -eq 0 -and $null -ne $q -and $q -ne "") {
            $userContent = "Query: $q"
            $localSystemMessage = "You are a helpful cloud and programming expert."
        } else {
            $historyCommands = (Get-History | Select-Object -Last $ContextLength | ForEach-Object { $_.CommandLine }) -join " ; "
            if ($null -eq $historyCommands -or $historyCommands -eq "") {
                Write-Host "No recent commands to send."
                return
            }
            $userContent = "PowerShell command history: $historyCommands"
            if ($null -ne $q -and $q -ne "") {
                $userContent += " Query: $q"
            }
            $localSystemMessage = $aiSystemMessage
        }

        # For reasoning models (o1-mini, o1-preview, o1, o3-mini), merge system and user instructions.
        if ($global:OpenAI_Current_Model -match '^(o1(-mini|-preview)?|o1|o3-mini)$') {
            $combinedContent = "$localSystemMessage`n$userContent"
            $payload = [PSCustomObject]@{
                "model" = $global:OpenAI_Current_Model
                "messages" = @(
                    [PSCustomObject]@{
                        "role" = "user"
                        "content" = $combinedContent
                    }
                )
            }
        } else {
            $payload = [PSCustomObject]@{
                "model" = $global:OpenAI_Current_Model
                "messages" = @(
                    [PSCustomObject]@{
                        "role"    = "system"
                        "content" = $localSystemMessage
                    },
                    [PSCustomObject]@{
                        "role"    = "user"
                        "content" = $userContent
                    }
                )
            }
        }

        if ($v) {
            Write-Host "Input JSON (payload) to OpenAI API:" -ForegroundColor Yellow
            Write-Host ($payload | ConvertTo-Json -Depth 2) -ForegroundColor White
        }

        $apiKey = Get-OpenAIAPIKey
        if ($null -eq $apiKey) {
            Write-Error "No API key is available to send the request to OpenAI."
            return
        }

        $plainApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKey)
        )

        $headers = @{ "Authorization" = "Bearer $plainApiKey" }
        if ($v) {
            $maskedHeaders = $headers.Clone()
            $maskedHeaders["Authorization"] = "Bearer [REDACTED]"
            Write-Host "Headers for API request:" -ForegroundColor Yellow
            Write-Host ($maskedHeaders | ConvertTo-Json -Depth 2) -ForegroundColor White
        }

        $endpoint = "https://api.openai.com/v1/chat/completions"
        $response = Invoke-WutRequest -payload $payload -endpoint $endpoint

        if ($v) {
            Write-Host "Response from OpenAI API:" -ForegroundColor Yellow
            Write-Host ($response | ConvertTo-Json -Depth 2) -ForegroundColor White
        }

        Show-WutResponse -response $response

    } catch {
        Write-Error "Failed to process the command: $($_.Exception.Message)"
    }
}

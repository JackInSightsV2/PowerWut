# PowerWut PowerShell Module

_PowerWut_ is a PowerShell module designed to interact with OpenAI’s models to analyze your PowerShell command history and answer custom queries. It securely manages your API key using the Windows Credential Vault and provides several helper functions for model management, secret storage, and intelligent AI querying.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation & Deployment](#installation--deployment)
- [Usage](#usage)
  - [Setting Up Your API Key](#setting-up-your-api-key)
  - [Listing & Changing AI Models](#listing--changing-ai-models)
  - [Sending Queries to OpenAI](#sending-queries-to-openai)
  - [Getting Help](#getting-help)
- [Detailed Explanation](#detailed-explanation)
  - [Module & Secret Management](#module--secret-management)
  - [API Key Handling](#api-key-handling)
  - [Context & Query Functions](#context--query-functions)
  - [REST API Request & Response Processing](#rest-api-request--response-processing)
- [Security Considerations](#security-considerations)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Contributing](#contributing)
- [Acknowledgements](#acknowledgements)

---

## Overview

The _PowerWut_ module simplifies the process of querying OpenAI models from PowerShell. It performs the following tasks:
- **Module Installation & Import:** Checks for and installs the `Microsoft.PowerShell.SecretManagement` and `Microsoft.PowerShell.SecretStore` modules if necessary.
- **Secret Storage:** Securely stores and retrieves your OpenAI API key using the Windows Credential Vault.
- **Model Management:** Lists available OpenAI models and allows you to switch between them.
- **Command History Analysis:** Captures recent PowerShell commands as context to aid in error analysis and troubleshooting.
- **Query Execution:** Sends the command history or custom queries to OpenAI’s API and returns the AI’s response.

---

## Features

- **Automated Dependency Management:** Automatically installs required secret management modules if not already present.
- **Secure API Key Storage:** Uses Windows Credential Vault to safely store your OpenAI API key.
- **Dynamic AI Model Selection:** Allows listing and switching between available AI models.
- **Contextual Analysis:** Aggregates your recent command history to provide context to the AI model.
- **Verbose Mode:** Offers detailed debugging output including API request payloads and responses.
- **Help System:** Includes a built-in help function that explains usage, parameters, and examples.

---

## Requirements

- **PowerShell 5.1 or later** (or PowerShell Core)
- **Internet Access** to reach the OpenAI API endpoints
- **OpenAI API Key** – Obtain one from [OpenAI's API Keys Page](https://platform.openai.com/api-keys)

---

## Installation & Deployment

1. **Download the Module:**
   - Save the provided PowerShell module as `PowerWut.psm1`.

2. **Import the Module:**
   - Open a PowerShell terminal and import the module:
     ```powershell
     Import-Module .\PowerWut.psm1
     ```
   - Alternatively, you can add the module's folder to your `$env:PSModulePath` for easier future access.

3. **Module Dependencies:**
   - The module automatically checks for, installs (if needed), and imports the following modules:
     - `Microsoft.PowerShell.SecretManagement`
     - `Microsoft.PowerShell.SecretStore`
   - It also registers a secret vault named **WutKeyVault** for secure API key storage.

4. **Persisting the Module:**
   - To load _PowerWut_ automatically in future sessions, consider adding the import command to your PowerShell profile.

---

## Usage

### Setting Up Your API Key

Before running any AI queries, set your OpenAI API key:

- **Interactive Setup:**
  - On first use, if no API key is found, the module will prompt:
    ```plaintext
    No API Key found. Please enter your OpenAI API Key.
    ```
  - Enter your key when prompted.

- **Manually Set the API Key:**
  - Use the provided function:
    ```powershell
    Set-OpenAIAPIKey -ApiKey "your_openai_api_key"
    ```

### Listing & Changing AI Models

- **List Available Models:**
  - Run:
    ```powershell
    wut -m
    ```
  - This displays the current model and lists available AI models (after applying filtering rules).

- **Change the Current Model:**
  - Use the `-sm` parameter:
    ```powershell
    wut -sm "gpt-4-turbo"
    ```
  - The module sets the new model as the current model for subsequent queries.

### Sending Queries to OpenAI

- **Analyze Recent Command History:**
  - Run a query without specifying a custom query:
    ```powershell
    wut
    ```
  - This sends the last 10 commands (default) as context to the AI model.

- **Custom Query with Command History:**
  - Include both command history and a custom query:
    ```powershell
    wut -c 20 -q "Explain the last 20 commands"
    ```
  - The `-c` parameter adjusts how many recent commands to include.

- **Direct Query Without History:**
  - To send a direct query without context, use a context length of zero:
    ```powershell
    wut -q "What is Azure?" -c 0
    ```

- **Verbose Mode:**
  - Add `-v` to see detailed API request and response information:
    ```powershell
    wut -q "What is PowerShell?" -v
    ```

### Getting Help

- Display the built-in help message by running:
  ```powershell
  wut -?

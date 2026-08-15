param location string
param projectName string
param storageAccountName string
param keyVaultName string
param keyVaultUri string
param entraClientId string
param gmailAddress string

var appServicePlanName = 'asp-${projectName}'
var functionAppName = 'func-${projectName}-${uniqueString(resourceGroup().id)}'

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' existing = {
  name: storageAccountName
}

resource appServicePlan 'Microsoft.Web/serverfarms@2022-09-01' = {
  name: appServicePlanName
  location: location
  sku: {
    name: 'B1'
    tier: 'Basic'
  }
  properties: {
    reserved: true
  }
  kind: 'linux'
}

resource functionApp 'Microsoft.Web/sites@2022-09-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      linuxFxVersion: 'DOTNET-ISOLATED|8.0'
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      appSettings: [
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'dotnet-isolated'
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccountName};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net'
        }
        {
          name: 'AzureAd__TenantId'
          value: 'consumers'
        }
        {
          name: 'AzureAd__ClientId'
          value: entraClientId
        }
        {
          name: 'AzureAd__ClientSecret'
          value: '@Microsoft.KeyVault(VaultName=${keyVaultName};SecretName=GraphClientSecret)'
        }
        {
          name: 'Sync__GmailAddress'
          value: gmailAddress
        }
        {
          name: 'Sync__CronSchedule'
          value: '0 */10 * * * * *'
        }
        {
          name: 'Storage__AccountUri'
          value: 'https://${storageAccountName}.table.core.windows.net'
        }
        {
          name: 'KeyVault__VaultUri'
          value: keyVaultUri
        }
      ]
    }
    httpsOnly: true
  }
}

output functionAppName string = functionApp.name
output functionAppPrincipalId string = functionApp.identity.principalId
output functionAppDefaultHostName string = functionApp.properties.defaultHostName
output outboundIpAddresses string = functionApp.properties.outboundIpAddresses

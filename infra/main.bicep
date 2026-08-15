param location string = resourceGroup().location
param projectName string = 'ogms'
param entraClientId string
param gmailAddress string = 'murragh2@gmail.com'
param allowedIpAddress string = ''

module storage 'modules/storage.bicep' = {
  name: 'storage'
  params: {
    location: location
    projectName: projectName
    allowedIpAddress: allowedIpAddress
  }
}

module keyVault 'modules/keyVault.bicep' = {
  name: 'keyVault'
  params: {
    location: location
    projectName: projectName
    allowedIpAddress: allowedIpAddress
  }
}

module functionApp 'modules/functionApp.bicep' = {
  name: 'functionApp'
  params: {
    location: location
    projectName: projectName
    storageAccountName: storage.outputs.storageAccountName
    keyVaultName: keyVault.outputs.keyVaultName
    keyVaultUri: keyVault.outputs.keyVaultUri
    entraClientId: entraClientId
    gmailAddress: gmailAddress
  }
}

module rbac 'modules/rbac.bicep' = {
  name: 'rbac'
  params: {
    functionAppPrincipalId: functionApp.outputs.functionAppPrincipalId
    keyVaultName: keyVault.outputs.keyVaultName
    storageAccountName: storage.outputs.storageAccountName
  }
}

output functionAppName string = functionApp.outputs.functionAppName
output resourceGroupName string = resourceGroup().name
output storageAccountName string = storage.outputs.storageAccountName
output keyVaultName string = keyVault.outputs.keyVaultName

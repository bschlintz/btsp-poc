export default {
  AzureStorageBlobAccount: process.env['AzureStorageBlobAccount'] || '',
  AzureStorageBlobKey: process.env['AzureStorageBlobKey'] || '',
  AzureStorageBlobContainer: process.env['AzureStorageBlobContainer'] || '',
  AdalClientId: process.env['AdalClientId'],
  AdalClientSecret: process.env['AdalClientSecret'],
  AdalTenant: process.env['AdalTenant'],
  GraphSharePointSiteId: process.env['GraphSharePointSiteId'],
  GraphSharePointDriveId: process.env['GraphSharePointDriveId']
}
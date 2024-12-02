# Trim SPO Document library Versioning
This folder includes PS scripts to permform the SPO verisoning trimming process: 

## Get-FileVersionExpriationReport.ps1
Generate the file version expiration report for SPO automatic file versioning expiration feature

## Set-DocLibExpiration.ps1
Set document library expiration for a site collection. the following logic is applied: 
 * Enable version limit to the site collection level and configure the max file version number
 * Enumerate all existing document libraries and configure the the max version number to those libraries which are not hidden, not system list and not site assets library. 
 * Initiate the automatic version deletion job

## VersionUtils.ps1
Versioning util script

## Inventory-DocLibVersions.ps1
Inventory document libraries' version settings for specific site collections defined in a CSV file. It can enumerate all sub sites. 

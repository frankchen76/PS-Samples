# Trim SPO Document library Versioning
This folder includes PS scripts to permform the SPO verisoning trimming process: 

## Get-FileVersionExpriationReport.ps1
Generate the file version expiration report for SPO automatic file versioning expiration feature

## Set-DocLibExpiration.ps1
set document library expiration for a site collection. the following logic is applied: 
 * apply version limit to the site collection level
 * enumerate all document library and configure the the max version number to those libraries which are not hidden, not system list and not site assets library. 
 * initiate the automatic version deletion job

## VersionUtils.ps1
Versioning util script


# SharePoint Provisioning Project

## Requirements / Purpose

The purpose of this project is to get comfortable working with PowerShell and SharePoint Provisioning.

1. Fork https://github.com/SharePoint/sp-dev-provisioning-templates
2. Create a new site and apply template (https://github.com/SharePoint/sp-dev-provisioning-templates/tree/master/tenant/contosoworks) to it. This can be on dev tenant http://vbtnd.onmicrosoft.com or your own. Provide link as response
3. Modify template to include the new years for Global Country Holiday, include ANZ public holidays.
4. Capture the site and push to forked RP with PR and merge
5. Create script to apply this template to 10000 sites. Remember title, description and URL will be different for each of these sites. Add this script as README.md to GitHub repo
6. Explain your approach for applying this template on demand via Azure
7. Explain Your approach for integrating a solution in step 6 into other systems

---

## Build Steps

- Open the "Apply Template Script" file and follow the prompts entering credentials where needed.
- Or copy this into admin PowerShell Terminal:

```
Import-Module PnP.PowerShell

$baseUrl = Read-Host "Enter base URL for SharePoint (e.g. https://tenant.sharepoint.com/sites/)"
$baseTitle = Read-Host "Enter Site Name"
$baseDescription = "Description for "
$siteNumber = [int] (Read-Host "Enter # of sites to be generated")
$siteOwner = Read-Host "Enter site owner user"
$adminSiteUrl = Read-Host "Enter Admin SharePoint URL"
$templatePath = Read-Host "Enter Path to template"
$credential = Get-Credential
Connect-PnPOnline -Url $adminSiteUrl -Credentials $credential

for ($i = 1; $i -le $siteNumber; $i++) {
    $siteTitle = "$baseTitle$i"
    $siteUrl = "$baseUrl$siteTitle"
    $siteDescription = "$baseDescription$siteTitle"

    try {

        Write-Host "Creating site: $siteTitle at URL $siteUrl"
        New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteUrl -Lcid 1033 -Description $siteDescription -Owner $siteOwner
        Write-Host "Site creation in progress for: $siteUrl"
    }
    catch {
        Write-Host "Failed to create site: $siteUrl. Error: $_"
        continue
    }
    Start-Sleep -Seconds 60
    Connect-PnPOnline -Url $siteUrl -Interactive

    try {

        Write-Host "Applying template to site: $siteUrl"
        Invoke-PnPSiteTemplate -Path $templatePath
        Write-Host "Template applied to site: $siteUrl"
    }
    catch {
        Write-Host "Failed to apply template to site: $siteUrl. Error: $_"
    }
    Start-Sleep -Seconds 10

}

Disconnect-PnPOnline


```
---
## Question 6/7

### Question 6
To apply the script online via Azure I would create a SharePoint List. I would then create a new Azure Function with a HTTP trigger allowing the script to be added to the Azure storage queue when a HTTP request is made. This request can then be created by the REST API and messaging for the queue can then be added through Microsoft Flow.

### Question 7
The most common way to allow other systems to send/recieve data into this script would be through HTTP requests. HTTP requests being very modular and most systems have the ability to invoke these requests. Commonly it would be a REST API which would be my method of implementation. Microsoft Flow also gives messaging regardless of the system allowing the script to send out logs regarding the queue status. 
---

## Approach

- To make the script as modular as possible requiring no changes to the script itself when used on different tenants.

---

## Features

- Automates the creation and application of templates to sites
- Is fully modular and allows the user to input any credential and host site values upon the script running.


---

## Future Goals

- To gain further knowledge in the provisioning of sites and templates in SharePoint/PowerShell
- Fix the issues that I came across in this project.
- Automate the process using Azure functions

---

## What did you struggle with?

- Ran into an issue with retrieving the PnP file of the template through the `Get-PnPTemplate` command. Leaving me to resort to using the `template.xml` template instead.

---

## Licensing Details

- Unlicensed

---

#Import-Module MSOnline
Add-Type –Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type –Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type –Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"

#update your information below
$userName = "<username>@<yourdomain>.onmicrosoft.com"; 
$site = "https://<yourdomain>.sharepoint.com"; 

$pwd = Read-Host -Prompt "Please enter your password" -AsSecureString ;

#uncomment below line if you do not want to enter password each time

#$pwd = "<your passowrd>" | ConvertTo-SecureString -AsPlainText -Force;

$context = New-Object Microsoft.SharePoint.Client.ClientContext($site);
$cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName,$pwd);
$context.Credentials = $cred;

$web = $context.Web;
$context.Load($web);
$context.ExecuteQuery();

$siteFeatures = $context.Site.Features
$context.Load($siteFeatures);
$context.ExecuteQuery();

$featureGuid = New-Object System.Guid "{9c0834e1-ba47-4d49-812b-7d4fb6fea211}"
$context.Site.Features.Add($featureGuid, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None);
$context.ExecuteQuery();

$lists = $web.Lists;
$contactsList = $lists.GetByTitle("MyContacts");
$query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(100);
$items = $contactsList.GetItems($query);


$context.Load($lists);
$context.Load($contactsList);
$context.Load($items);
$context.ExecuteQuery();

$lci = New-Object Microsoft.SharePoint.Client.ListCreationInformation;
$lci.Title = "Custom List";
$lci.TemplateType = '100';
$customList = $lists.Add($lci);
$customList.Update();
$context.ExecuteQuery();

$availableFields = $web.AvailableFields;
$context.Load($availableFields);
$context.ExecuteQuery();
$companyField = $availableFields | Where {$_.Title -eq "Company"}
$context.Load($companyField);
$context.ExecuteQuery();
$customList = $lists.GetByTitle("Custom List");
$context.Load($customList);
$context.ExecuteQuery();
$customList.Fields.Add($companyField);
$customList.Update();
$context.ExecuteQuery();

$defaultView = $customList.DefaultView;
$defaultView.ViewFields.Add("Company");
$defaultView.Update();
$customList.Update();
$context.ExecuteQuery();

foreach($list in $lists)
{
	Write-Host $list.Title;
}

foreach($item in $items)
{
    Write-Host $item["FirstName"]
}

$mms = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context);
$context.Load($mms);
$context.ExecuteQuery();

$termStores = $mms.TermStores;
$context.Load($termStores);
$context.ExecuteQuery();

$termStore = $termStores[0];
$context.Load($termStore);
$context.ExecuteQuery();

$group = $termStore.CreateGroup("PowerShell", "{C93600E9-49D0-4079-8DBE-8282A8CE4119}");
$context.Load($group);
$context.ExecuteQuery();

#$group = $termStore.Groups.GetByName("PowerShell");
#context.Load($group);
#$context.ExecuteQuery();

$termSet = $group.CreateTermSet("SharePoint", "{6768B471-7EA3-4981-81A4-EA4902543365}", 1033);
$context.Load($termSet);
$context.ExecuteQuery();

$term = $termSet.CreateTerm("CSOM", 1033, "{E16CD934-74DB-4D2A-AB39-D24422DBC1B1}");
$context.Load($term);
$context.ExecuteQuery();

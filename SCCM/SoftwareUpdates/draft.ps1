$updateProducts = @(
                [pscustomobject]@{"ProductID"="fa5ef799-b817-439e-abf7-c76ba0cacb75";"ProductName"=" ASP.NET Web Frameworks"}
                [pscustomobject]@{"ProductID"="83aed513-c42d-4f94-b4dc-f2670973902d";"ProductName"=" CAPICOM"}
                [pscustomobject]@{"ProductID"="0bbd2260-7478-4553-a791-21ab88e437d2";"ProductName"=" Device Health"}
                [pscustomobject]@{"ProductID"="5cc25303-143f-40f3-a2ff-803a1db69955";"ProductName"=" Locally published packages"}
                [pscustomobject]@{"ProductID"="6ac905a5-286b-43eb-97e2-e23b3848c87d";"ProductName"=" Microsoft Advanced Threat Analytics"}
                [pscustomobject]@{"ProductID"="f3869cc3-897b-4339-bb10-32ab2c765862";"ProductName"=" Microsoft Monitoring Agent"}
                [pscustomobject]@{"ProductID"="ca6616aa-6310-4c2d-a6bf-cae700b85e86";"ProductName"=" Microsoft SQL Server 2017"}
                [pscustomobject]@{"ProductID"="dee854fd-e9d2-43fd-bbc3-f7568e3ce324";"ProductName"=" Microsoft SQL Server Management Studio v17"}
                [pscustomobject]@{"ProductID"="6b9e8b26-8f50-44b9-94c6-7846084383ec";"ProductName"=" MS Security Essentials"}
                [pscustomobject]@{"ProductID"="6248b8b1-ffeb-dbd9-887a-2acf53b09dfe";"ProductName"=" Office 2002/XP"}
                [pscustomobject]@{"ProductID"="1403f223-a63f-f572-82ba-c92391218055";"ProductName"=" Office 2003"}
                [pscustomobject]@{"ProductID"="041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591";"ProductName"=" Office 2007"}
                [pscustomobject]@{"ProductID"="84f5f325-30d7-41c4-81d1-87a0e6535b66";"ProductName"=" Office 2010"}
                [pscustomobject]@{"ProductID"="704a0a4a-518f-4d69-9e03-10ba44198bd5";"ProductName"=" Office 2013"}
                [pscustomobject]@{"ProductID"="25aed893-7c2d-4a31-ae22-28ff8ac150ed";"ProductName"=" Office 2016"}
                [pscustomobject]@{"ProductID"="6c5f2e66-7dbc-4c59-90a7-849c4c649d7a";"ProductName"=" Office 2019"}
                [pscustomobject]@{"ProductID"="30eb551c-6288-4716-9a78-f300ec36d72b";"ProductName"=" Office 365 Client"}
                [pscustomobject]@{"ProductID"="7cf56bdd-5b4e-4c04-a6a6-706a2199eff7";"ProductName"=" Report Viewer 2005"}
                [pscustomobject]@{"ProductID"="79adaa30-d83b-4d9c-8afd-e099cf34855f";"ProductName"=" Report Viewer 2008"}
                [pscustomobject]@{"ProductID"="f7f096c9-9293-422d-9be8-9f6e90c2e096";"ProductName"=" Report Viewer 2010"}
                [pscustomobject]@{"ProductID"="6cf036b9-b546-4694-885a-938b93216b66";"ProductName"=" Security Essentials"}
                [pscustomobject]@{"ProductID"="9f3dd20a-1004-470e-ba65-3dc62d982958";"ProductName"=" Silverlight"}
                [pscustomobject]@{"ProductID"="cc4ab3ac-9d9a-4f53-97d3-e0d6de39d119";"ProductName"=" System Center 2016 - Operations Manager"}
                [pscustomobject]@{"ProductID"="cd8d80fe-5b55-48f1-b37a-96535dca6ae7";"ProductName"=" TMG Firewall Client"}
                [pscustomobject]@{"ProductID"="a0dd7e72-90ec-41e3-b370-c86a245cd44f";"ProductName"=" Visual Studio 2005"}
                [pscustomobject]@{"ProductID"="e3fde9f8-14d6-4b5c-911c-fba9e0fc9887";"ProductName"=" Visual Studio 2008"}
                [pscustomobject]@{"ProductID"="c9834186-a976-472b-8384-6bb8f2aa43d9";"ProductName"=" Visual Studio 2010"}
                [pscustomobject]@{"ProductID"="cbfd1e71-9d9e-457e-a8c5-500c47cfe9f3";"ProductName"=" Visual Studio 2010 Tools for Office Runtime"}
                [pscustomobject]@{"ProductID"="e1c753f2-9f79-4577-b75b-913f4230feee";"ProductName"=" Visual Studio 2010 Tools for Office Runtime"}
                [pscustomobject]@{"ProductID"="abddd523-04f4-4f8e-b76f-a6c84286cc67";"ProductName"=" Visual Studio 2012"}
                [pscustomobject]@{"ProductID"="cf4aa0fc-119d-4408-bcba-181abb69ed33";"ProductName"=" Visual Studio 2013"}
                [pscustomobject]@{"ProductID"="1731f839-8830-4b9c-986e-82ee30b24120";"ProductName"=" Visual Studio 2015"}
                [pscustomobject]@{"ProductID"="a3c2375d-0c8a-42f9-bce0-28333e198407";"ProductName"=" Windows 10"}
                [pscustomobject]@{"ProductID"="d2085b71-5f1f-43a9-880d-ed159016d5c6";"ProductName"=" Windows 10 LTSB"}
                [pscustomobject]@{"ProductID"="3b4b8621-726e-43a6-b43b-37d07ec7019f";"ProductName"=" Windows 2000"}
                [pscustomobject]@{"ProductID"="bfe5b177-a086-47a0-b102-097e4fa1f807";"ProductName"=" Windows 7"}
                [pscustomobject]@{"ProductID"="6407468e-edc7-4ecd-8c32-521f64cee65e";"ProductName"=" Windows 8.1"}
                [pscustomobject]@{"ProductID"="b1b8f641-1ff2-4ae6-b247-4fe7503787be";"ProductName"=" Windows Admin Center"}
                [pscustomobject]@{"ProductID"="8c3fcc84-7410-4a95-8b89-a166a0190486";"ProductName"=" Windows Defender"}
                [pscustomobject]@{"ProductID"="50c04525-9b15-4f7c-bed4-87455bcd7ded";"ProductName"=" Windows Dictionary Updates"}
                [pscustomobject]@{"ProductID"="dbf57a08-0d5a-46ff-b30c-7715eb9498e9";"ProductName"=" Windows Server 2003"}
                [pscustomobject]@{"ProductID"="7f44c2a7-bc36-470b-be3b-c01b6dc5dd4e";"ProductName"=" Windows Server 2003, Datacenter Edition"}
                [pscustomobject]@{"ProductID"="ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf";"ProductName"=" Windows Server 2008"}
                [pscustomobject]@{"ProductID"="fdfe8200-9d98-44ba-a12a-772282bf60ef";"ProductName"=" Windows Server 2008 R2"}
                [pscustomobject]@{"ProductID"="d31bd4c3-d872-41c9-a2e7-231f372588cb";"ProductName"=" Windows Server 2012 R2"}
                [pscustomobject]@{"ProductID"="569e8e8f-c6cd-42c8-92a3-efbb20a0f6f5";"ProductName"=" Windows Server 2016"}
                [pscustomobject]@{"ProductID"="f702a48c-919b-45d6-9aef-ca4248d50397";"ProductName"=" Windows Server 2019"}
                [pscustomobject]@{"ProductID"="4e487029-f550-4c22-8b31-9173f3f95786";"ProductName"=" Windows Server Manager – Windows Server Update Services (WSUS) Dynamic Installer"}
                [pscustomobject]@{"ProductID"="26997d30-08ce-4f25-b2de-699c36a8033a";"ProductName"=" Windows Vista"}
                [pscustomobject]@{"ProductID"="558f4bc3-4827-49e1-accf-ea79fd72d4c9";"ProductName"=" Windows XP"}
                [pscustomobject]@{"ProductID"="a4bedb1d-a809-4f63-9b49-3fe31967b6d0";"ProductName"=" Windows XP 64-Bit Edition Version 2003"}
                [pscustomobject]@{"ProductID"="874a7757-3a13-43b2-a7f2-cf2ff43dd6bf";"ProductName"=" Windows XP Embedded"}
                [pscustomobject]@{"ProductID"="4cb6ebd5-e38a-4826-9f76-1416a6f563b0";"ProductName"=" Windows XP x64 Edition"}                
            )            
$updateClassification = @(
                [pscustomobject]@{"ProductID"="051f8713-e600-4bee-b7b7-690d43c78948";"ProductName"=" WSUS Infrastructure Updates"}
                [pscustomobject]@{"ProductID"="0fa1201d-4330-4fa8-8ae9-b877473b6441";"ProductName"=" Security Updates"}                
            )
$timeReport = 45
$a = Get-CMSoftwareUpdate -fast -DatePostedMax $(Get-Date).AddDays(-$timeReport)  | Where {
    $_.IsExpired-eq $false -and ` 
    $_.IsSuperseded -eq $false -and ` 
    $_.Severity -ge $severity   
    
} 
$MSFTUpdates=@()
foreach ($upd in $a ){
    $prodCheck=$false
    $catCheck=$false
    foreach ($cat in $upd.CategoryInstance_UniqueIDs){
        if (($cat -like "Product*") -and ($updateProducts.ProductID -contains $($cat -replace "Product:",""))){
            $prodCheck=$true
        }
        elseif(($cat -like "Update*") -and ($updateClassification.ProductID -contains $($cat -replace "UpdateClassification:",""))){
            $catCheck=$true
        }
    }
    if ($prodCheck -and $catCheck){
        $MSFTUpdates+=$upd
    }
    
}
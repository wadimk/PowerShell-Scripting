# Remember to install the EWS managed API from here:
# http://www.microsoft.com/downloads/en/details.aspx?displaylang=en&FamilyID=c3342fb3-fbcc-4127-becf-872c746840e1

#\\storage\distrib\microsoft\exchange\exchange_web_services\EwsManagedApi32.msi
Clear-Host
# initial paramaters
$fileName = "C:\trash\contacts.xml"
$recurse = $true

# Contact Mapping - this maps the attributes in the CSV file (left) to the attributes EWS uses.
# NB: If you change these, please note "First Name" is specified at line 102 as a required attribute and
# "First Name" and "Last Name" are hard coded at lines 187-197 when constructing NickName and FileAs.
$ContactMapping=@{
    "First Name" = "GivenName";
    "Middle Name" = "MiddleName";
    "Last Name" = "Surname";
    "Company" = "CompanyName";
    "Department" = "Department";
    "Job Title" = "JobTitle";
    "Business Street" = "Address:Business:Street";
    "Business City" = "Address:Business:City";
    "Business State" = "Address:Business:State";
    "Business Postal Code" = "Address:Business:PostalCode";
    "Business Country/Region" = "Address:Business:CountryOrRegion";
    "Home Street" = "Address:Home:Street";
    "Home City" = "Address:Home:City";
    "Home State" = "Address:Home:State";
    "Home Postal Code" = "Other:Home:PostalCode";
    "Home Country/Region" = "Address:Home:CountryOrRegion";
    "Other Street" = "Address:Other:Street";
    "Other City" = "Address:Other:City";
    "Other State" = "Address:Other:State";
    "Other Postal Code" = "Address:Other:PostalCode";
    "Other Country/Region" = "Address:Other:CountryOrRegion";
    "Assistant's Phone" = "Phone:AssistantPhone";
    "Business Fax" = "Phone:BusinessFax";
    "Business Phone" = "Phone:BusinessPhone";
    "Business Phone 2" = "Phone:BusinessPhone2";
    "Callback" = "Phone:CallBack";
    "Car Phone" = "Phone:CarPhone";
    "Company Main Phone" = "Phone:CompanyMainPhone";
    "Home Fax" = "Phone:HomeFax";
    "Home Phone" = "Phone:HomePhone";
    "Home Phone 2" = "Phone:HomePhone2";
    "ISDN" = "Phone:ISDN";
    "Mobile Phone" = "Phone:MobilePhone";
    "Other Fax" = "Phone:OtherFax";
    "Other Phone" = "Phone:OtherTelephone";
    "Pager" = "Phone:Pager";
    "Primary Phone" = "Phone:PrimaryPhone";
    "Radio Phone" = "Phone:RadioPhone";
    "TTY/TDD Phone" = "Phone:TtyTddPhone";
    "Telex" = "Phone:Telex";
    "Anniversary" = "WeddingAnniversary";
    "Birthday" = "Birthday";
    "E-mail Address" = "Email:EmailAddress1";
    "E-mail 2 Address" = "Email:EmailAddress2";
    "E-mail 3 Address" = "Email:EmailAddress3";
    "Initials" = "Initials";
    "Office Location" = "OfficeLocation";
    "Manager's Name" = "Manager";
    "Mileage" = "Mileage";
    "Notes" = "Body";
    "Profession" = "Profession";
    "Spouse" = "SpouseName";
    "Web Page" = "BusinessHomePage";
    "Contact Picture File" = "Method:SetContactPicture"
}

#
#

# loading additional resources
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

# creating table for imported data
$ContactsTable = New-Object System.Data.DataTable "Contacts"

$DataColumn = New-Object System.Data.DataColumn OwnerMailBox,([string])
$ContactsTable.Columns.Add($DataColumn)
$DataColumn = New-Object System.Data.DataColumn OwnerMailBoxFolder,([string])
$ContactsTable.Columns.Add($DataColumn)

foreach ($ContactMap in $ContactMapping.GetEnumerator()) {
	$DataColumn = New-Object System.Data.DataColumn $ContactMap.Key,([string])
	#$DataColumn.Caption = $ContactMap.Value 
	$ContactsTable.Columns.Add($DataColumn)
}

# scanning contacts folder
function scan-folder([Microsoft.Exchange.WebServices.Data.Folder]$folder, [string]$path)
{
	if ($folder.TotalCount -gt 0)
	{
		("Items Count: " + $folder.TotalCount)
		("Scanning folder: " + $path)
		
        $offset = 0;
        $view = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset)
        while (($results = $folder.FindItems($view)).Items.Count -gt 0)
        {
			$response = $service.LoadPropertiesForItems($results, [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties)
			
            foreach ($item in $results)
            {
                if ($item.ItemClass -eq "IPM.Contact")
                {
					$DataRow = $ContactsTable.NewRow()			
					$DataRow.OwnerMailBox = $email
					$DataRow.OwnerMailBoxFolder = $path
					
					foreach ($ContactMap in $ContactMapping.GetEnumerator())
					{
						# Will this call a more complicated mapping?
			            if ($ContactMap.Value -like "*:*")
			            {
			                # Make an array using the : to split items.
			                $MappingArray = $ContactMap.Value.Split(":")
			                # Do action
			                switch ($MappingArray[0])
			                {
			                    "Email"
			                    {
									$DataRow.($ContactMap.Key) = $item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::($MappingArray[1])].Address
			                    }
			                    "Phone"
			                    {
			                        $DataRow.($ContactMap.Key) = $item.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::($MappingArray[1])]
			                    }
			                    "Address"
			                    {
									$PhysicalAddressEntry = $item.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])]
			                        $DataRow.($ContactMap.Key) = $PhysicalAddressEntry.($MappingArray[2])

#			                        switch ($MappingArray[1])
#			                        {
#			                            "Business"
#			                            {
#											$BusinessPhysicalAddressEntry = $item.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])]
#			                                $DataRow.($ContactMap.Key) = $BusinessPhysicalAddressEntry.($MappingArray[2])
#											
#			                            }
#			                            "Home"
#			                            {
#			                                $DataRow.($ContactMap.Key) = $item.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])].ToString()
#											$a=$item.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])]
#			                            }
#			                            "Other"
#			                            {
#			                                $DataRow.($ContactMap.Key) = $item.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])].ToString()
#											$a=$item.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::($MappingArray[1])]
#			                            }
#			                        }
			                    }
			                    "Method"
			                    {
			                        switch ($MappingArray[1])
			                        {
			                            "SetContactPicture" 
			                            {
#			                                if (!$Exchange2007)
#			                                {
#			                                    if (!(Get-Item -Path $ContactItem.$Key -ErrorAction SilentlyContinue))
#			                                    {
#			                                        throw "Contact Picture File not found at $($ContactItem.$Key)";
#			                                    }
#			                                    $ExchangeContact.SetContactPicture($ContactItem.$Key);
#			                                }
			                            }
			                        }
			                    }
			                
			                }                
			            } else {
			                # It's a direct mapping - simple!
			                if ($ContactMap.Key -eq "Birthday" -or $ContactMap.Key -eq "WeddingAnniversary")
			                {
			                    $DataRow.($ContactMap.Key) = $item.($ContactMap.Value)
			                }
							
			                $DataRow.($ContactMap.Key) = $item.($ContactMap.Value)
			            }

						#$ContactMap.Value
						#$item[$ContactMap.Value]
						#$DataRow[$ContactMap.Key] = $item.$ContactMap.Value
					}

					$ContactsTable.Rows.Add($DataRow)					
				}
            }

            $offset += $results.Items.Count
            $view = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset)
		}		
	}		
	
	
   
    try
    {
        $offset = 0;


        if ($recurse)
        {
            # Recursively do subfolders
            $folderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(2147483647)
            $subfolders = $folder.FindFolders($folderView)   
            foreach ($subfolder in $subfolders)
            {
                try
                {
                    scan-folder $subfolder ($path + "\" + $subfolder.DisplayName)
                }
                catch { "Error processing folder: " + $subfolder.DisplayName }
            }
        }
    }
    catch
    {
		throw 
    }
}

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
if ($service -eq $null)
{
    Write-Error "Could not instantiate ExchangeService object."
    return
}

$mailboxes = Get-Mailbox -Server ex2010
#$mailboxes = Get-Mailbox "Kosin Vadim"

foreach ($mailbox in $mailboxes) {
	[string] $email = $mailbox.PrimarySmtpAddress.ToString()

	"E-Mail: " + $email 
	$service.AutodiscoverUrl($email)
	$service.PreAuthenticate = $true 
			
	$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$email)
	
	$rootfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderidcnt)

	scan-folder $rootfolder ""
}

$ContactsTable.WriteXml($fileName)
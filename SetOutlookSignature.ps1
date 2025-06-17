# https://dev.to/jcoelho/how-to-deploy-email-signatures-through-group-policies-4aj1r
# https://www.youtube.com/watch?app=desktop&v=Y3r2XWvsUl0

# Gets the path to the user appdata folder
$AppData = (Get-Item env:appdata).value
# This is the default signature folder for Outlook
$localSignatureFolder = $AppData+'\Microsoft\Signatures'
# Signature in shared directory
$templateFilePath = "SHARED_PATH" # TODO: Replace "SHARED_PATH" with a path shared to everyone on your network (e.g. (x = ip): \\xxx.xxx.x.xxx\shared\signature)
# Get current user name
$userName = $env:username

# Save it as <username>.htm
$fileName = $localSignatureFolder + "\" + $userName + ".htm"

# Remove existing template (Refresh template on sign in)
# Will throw a non-fatal error if signature does not exist yet
$safePath = "C:\Users\" + $userName + "\AppData\Roaming\Microsoft\Signatures\" + $userName + ".htm"
Remove-Item -Path $safePath

# Define HTML segments
$nameSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 0px;">
    <span style="font-family: Tahoma, sans-serif; font-size: 12pt; color: rgb(23, 93, 0);">
        <b>NAME</b>
    </span>
</p>
"@
$jobDepartmentSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 15px; color: rgb(53, 123, 20);">
    <b>JOB COMMA DEPARTMENT</b>
    </span>
</p>
"@
$companySegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
    <b><i>COMPANY</i></b>
    </span>
</p>
"@
$locationSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
        LOCATION Branch
    </span>
</p>
"@
$streetAddressSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
        STREETADDRESS
    </span>
</p>
"@
$cityStatePostalSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
        CITY STATE POSTALCODE
    </span>
</p>
"@
$officePhoneSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 0px;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(127, 127, 127);">
        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAbtJREFUSEu9lT9LG2EYwH/PiS6maMF+AO2aSgcHzdBLBhEqoh9AqB+g6GIhzTmElosROgiFrg5+AS0ugsVcHBw6pVOXoi0IFpdUKhjRe5ozTSB/ronnpTc/9/s9f973eYUuf9JlPv9ZYDnaVJHqezLxxaCV1leQypUQ6auDKR/JmLMhCZwzhKF6mBaw40/DEVjOZ2CsoYIiGfNhSILcBshCE+z6apC1yV9BJA0zyC8gutHQoi+USjHeTV3cX/A6/wjDPQHp/Qv7Sq+bIJ04DQL3/mm+B5azBcyi8g1lgtVnZ0HhPoL9cTAOgQtu3CjZxHG4Ao9mOTvANLCNbc6FL0gejGC4BYQISpKMuRZU4r+LUvl5RDdBvfXxAju+2SRZ3u0nMtBDevzcL4F/LzvLeQusVCSSwjazNVDlxO2BjKIUEf0O8qOczBG2uVSNa7NNVbDybyoSwNtLwktcuQT9hMGTlpnbZo3b2bq+bZf7AeQBym+Qn+X2Pfady50FHskbfI+7Dsy0HXggQZVqOTFUXyHeMa7d+HrnvQRVVHp/iCvjOYIJGgWGUQZu35NQBG37VAnobMgdwlqFdV3wB5TpfhnyimOWAAAAAElFTkSuQmCC" data-imagetype="AttachmentByCid" data-custom="AAMkADQ3MGI3MDdhLWY0ZmEtNDRiMS05YWVlLTRjZmQ2ZjUzOTM1YgBGAAAAAAAbx0xPT4quT6ySHAzhOfc%2BBwBSF7dHh%2B%2F%2BTr0zJKuEgJzxAAAAAAEJAABSF7dHh%2B%2F%2BTr0zJKuEgJzxAAGRK23sAAABEgAQAGiAYDl31IFNgC1%2FrZnN2W4%3D" data-outlook-trace="F:1|T:1" width="12" height="12" style="width: 9.72pt; height: 9.72pt; min-width: auto; min-height: auto; margin: 0px; vertical-align: top;">
        &nbsp;OFFICEPHONE1
    </span>
</p>
"@
$mobilePhoneSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 0px;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(127, 127, 127);">
        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAN9JREFUSEvtVUsKwjAUnNfW7yXceI0SQfEi4iG06wqeQTxJFdpew52HsETFJ21BsE2bdOFCSJZhmHlveMwQfvxIyR+cJ4B3BLMPon7rDMx3sJPiRSvs/WsV2yCQnADMOy4XIRRLM4FtLLWTV5nyTXazgZlAkHABDIV6wypLC77JIivwMdFaVFphr0gbGdaiP7aoHsm34otorNxKkb7tUfHFwhl63hQPSYB7AWhUEzEWUBaOVkAiFEOzwtmkERxe1G3grPxTTA90qMyy9A9gFtrqzKuSEOPprs1LX3uX5oA3QTqTGZNHd6QAAAAASUVORK5CYII=" data-imagetype="AttachmentByCid" data-custom="AAMkADQ3MGI3MDdhLWY0ZmEtNDRiMS05YWVlLTRjZmQ2ZjUzOTM1YgBGAAAAAAAbx0xPT4quT6ySHAzhOfc%2BBwBSF7dHh%2B%2F%2BTr0zJKuEgJzxAAAAAAEJAABSF7dHh%2B%2F%2BTr0zJKuEgJzxAAGRK23sAAABEgAQAMT97hYPoOxMu3lDV8Lk%2F%2BU%3D" data-outlook-trace="F:1|T:1" width="12" height="12" style="width: 9.72pt; height: 9.72pt; min-width: auto; min-height: auto; margin: 0px; vertical-align: top;">
        &nbsp;MOBILE1<br>
    </span>
</p>
"@
$emailSegment = @"
 <p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 0px;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(127, 127, 127);">
        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAYtJREFUSEvtlbsvRFEQh7+5WBIhNERFNDq1gl2PzqvT0QnxSChEuKu4YW88NjREIv4D0UgoiIJLI9FoVBqR6EQQIsEdXOuxyO7a0Gyc7pz8Zr6Z35ycI/zxkj/OTyoBgtv6q3bZAc+dd4uCzhJoyy9B1rADjdGA553pNCA6B5QkBVI9xTD6CPmXX+PfOxjayWei6pyB9WwysyyUfoT0BEEPqM7jyxrBqrhkcDeHqcqr6A7MrTNEhrH9iyCK6ZQjugBUxIboAUZaJ2NVe54uuNWEMEeouvjTDCJDVnVQo4tx/yGWGtw57aiGEcmNAqnegIzic6exau4Z2iklzZ0F6j3d1yF/vEV695RwhnTXwqq5xXSKECZB214gugoPvdh1x3TsZ1Bw3Y1iA9lvRcQGvMmOEHoIBTa8k5HtZlQE278S2dfi6jwiZV9sTBAQiXuuWLqxAyfegblZiPjC4LaCfP8a/Azgpb1ACKKqqNgIeTGHnwQgwRsbkf0D4vqVghbF7Tk5QSp9mck5EDfqEcAgjxnUPiReAAAAAElFTkSuQmCC" data-imagetype="AttachmentByCid" data-custom="AAMkADQ3MGI3MDdhLWY0ZmEtNDRiMS05YWVlLTRjZmQ2ZjUzOTM1YgBGAAAAAAAbx0xPT4quT6ySHAzhOfc%2BBwBSF7dHh%2B%2F%2BTr0zJKuEgJzxAAAAAAEJAABSF7dHh%2B%2F%2BTr0zJKuEgJzxAAGRK23sAAABEgAQADv75bqnei9GuELHqWpVtN0%3D" data-outlook-trace="F:1|T:1" width="12" height="12" style="width: 9.72pt; height: 9.72pt; min-width: auto; min-height: auto; margin: 0px; vertical-align: top;">
        <a href="mailto:EMAIL" data-linkindex="0" style="margin: 0px;">&nbsp;EMAIL</a>
    </span>
</p>
"@
$websiteSegment = @"
 <p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 0px;">
    <span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(127, 127, 127);">
        <a href="https://example.com/" target="_blank" data-linkindex="0" style="margin: 0px;">https://example.com/</a>
    </span>
</p>
"@

# Query AD and get user object based on username
$filter = "(&(objectCategory=User)(samAccountName=$userName))"
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.Filter = $filter
$ADUserPath = $searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()

# Extract AD attributes into variables
$name = $ADUser.displayName
$job = $ADUser.title
$department = $ADUser.department
$company = $ADUser.company
$location = $ADUser.physicalDeliveryOfficeName 
$streetAddress  = $ADUser.streetAddress
$city = $ADUser.l
$state = $ADUser.st
$postalCode = $ADUser.postalCode
$phone = $ADUser.telephoneNumber
$mobile = $ADUser.mobile
$email = $ADUser.mail
$website = $ADUser.website

# Get the content of the template file
$rawTemplate = get-content $templateFilePath"\template.htm"

# Replace values in template file

# FIRSTNAME LASTNAME
if ($name.Count -eq 1) {
    # Add name to template
    $nameSegment = $nameSegment -replace "NAME",$name

    # Update Signature
    $signature = $rawTemplate -replace "NAME_SEG",$nameSegment                      
    $rawTemplate = $signature
} else {
    # Ignore NAME line if no name exists in AD
    $signature = $rawTemplate -replace "NAME_SEG",""
    $rawTemplate = $signature
}

# JOB, DEPARTMENT
if (($job.Count -eq 1) -or ($department.Count -eq 1)) {
    # Check if job exists and add it
    if ($job.Count -eq 1) {
        $jobDepartmentSegment = $jobDepartmentSegment -replace "JOB",$job
    } else {
        $jobDepartmentSegment = $jobDepartmentSegment -replace "JOB",""
    }

    # Check if department exists and add it
    if ($department.Count -eq 1) {
        $jobDepartmentSegment = $jobDepartmentSegment -replace "DEPARTMENT",$department
    } else {
        $jobDepartmentSegment = $jobDepartmentSegment -replace "DEPARTMENT",""
    }

    # Add a comma if both values exist
    if (($job.Count -eq 1) -and ($department.Count -eq 1)) {
        $jobDepartmentSegment = $jobDepartmentSegment -replace "COMMA",","
    } else {
        $jobDepartmentSegment = $jobDepartmentSegment -replace "COMMA",""
    }
    
    # Update Signature
    $signature = $rawTemplate -replace "JOB_DPT_SEG",$jobDepartmentSegment
    $rawTemplate = $signature
} else {
    # Ignore JOB, DEPARTMENT line if no job/department exists in AD
    $signature = $rawTemplate -replace "JOB_DPT_SEG",""
    $rawTemplate = $signature
}

# COMPANY
if ($company.Count -eq 1) {
    # Add company to template
    $companySegment = $companySegment -replace "COMPANY",$company

    # Update Signature
    $signature = $rawTemplate -replace "COMPANY_SEG",$companySegment
    $rawTemplate = $signature
} else {
    # Ignore COMPANY line if no company exists in AD
    $signature = $rawTemplate -replace "COMPANY_SEG",""
    $rawTemplate = $signature
}

# LOCATION (office)
if ($office.Count -eq 1) {
    # Add location to template
    $locationSegment = $locationSegment -replace "LOCATION",$location

    # Update Signature
    $signature = $rawTemplate -replace "LOCATION_SEG",$locationSegment
    $rawTemplate = $signature
} else {
    # Ignore LOCATION line if no location exists in AD
    $signature = $rawTemplate -replace "LOCATION_SEG",""
    $rawTemplate = $signature
}

# ADDRESS
if ($streetAddress.Count -eq 1) {
    # Add Street Address to template
    $streetAddressSegment = $streetAddressSegment -replace "STREETADDRESS",$streetAddress

    # Update Signature
    $signature = $rawTemplate -replace "STREET_ADDR_SEG",$streetAddressSegment
    $rawTemplate = $signature
} else {
    # Ignore ADDRESS line if no address exists in AD
    $signature = $rawTemplate -replace "STREET_ADDR_SEG",""
    $rawTemplate = $signature
}

# CITY STATE POSTALCODE
if (($city.Count -eq 1) -or ($state.Count -eq 1) -or ($postalCode.Count -eq 1)) {
    # Check if city exists and add it
    if ($city.Count -eq 1) {
        $cityStatePostalSegment = $cityStatePostalSegment -replace "CITY",$city
    } else {
        $cityStatePostalSegment = $cityStatePostalSegment -replace "CITY",""
    }

    # Check if state exists and add it
    if ($state.Count -eq 1) {
        $cityStatePostalSegment = $cityStatePostalSegment -replace "STATE",$state
    } else {
        $cityStatePostalSegment = $cityStatePostalSegment -replace "STATE",""
    }

    # Check if postal code exists and add it
    if ($postalCode.Count -eq 1) {
        $cityStatePostalSegment = $cityStatePostalSegment -replace "POSTALCODE",$postalCode
    } else {
        $cityStatePostalSegment = $cityStatePostalSegment -replace "POSTALCODE",""
    }

    # Update Signature
    $signature = $rawTemplate -replace "CITY_ST_PO_SEG",$cityStatePostalSegment
    $rawTemplate = $signature
} else {
    # Ignore CITY/STATE/PO line if none exists in AD
    $signature = $rawTemplate -replace "CITY_ST_PO_SEG",""
    $rawTemplate = $signature
}

# PHONE (desk)
if ($phone.Count -eq 1) {
    # Add phone to template
    $phone = $phone.split(";|=")

    $phone[1] = "&nbsp;ext.&nbsp;"
    #                               ;ext.        123
    $phone = "+1 (555) 123-4567" + $phone[1] + $phone[2]

    $officePhoneSegment = $officePhoneSegment -replace "OFFICEPHONE1",$phone

    # Update Signature
    $signature = $rawTemplate -replace "OFFICE_PHONE_SEG",$officePhoneSegment
    $rawTemplate = $signature
} else {
    # Ignore PHONE line if none exists in AD
    $signature = $rawTemplate -replace "OFFICE_PHONE_SEG",""
    $rawTemplate = $signature
}

# MOBILE
if ($mobile.Count -eq 1) {
    # Add mobile to template
    $mobilePhoneSegment = $mobilePhoneSegment -replace "MOBILE1",$mobile

    # Update Signature
    $signature = $rawTemplate -replace "MOBILE_PHONE_SEG",$mobilePhoneSegment
    $rawTemplate = $signature
} else {
    # Ignore MOBILE line if none exists in AD
    $signature = $rawTemplate -replace "MOBILE_PHONE_SEG",""
    $rawTemplate = $signature
}

# EMAIL
if ($email.Count -eq 1) {
    # Add email to template
    $emailSegment = $emailSegment -replace "EMAIL",$email

    # Update Signature
    $signature = $rawTemplate -replace "EMAIL_SEG",$emailSegment
    $rawTemplate = $signature
} else {
    # Ignore EMAIL line if none exists in AD
    $signature = $rawTemplate -replace "EMAIL_SEG",""
    $rawTemplate = $signature
}

# WEBSITE
if ($website.Count -eq 1) {
    # Add Street Address to template
    $websiteSegment = $websiteSegment -replace "WEBSITE",$website

    # Update Signature
    $signature = $rawTemplate -replace "WEBSITE_SEG",$websiteSegment
    $rawTemplate = $signature
} else {
    # Ignore ADDRESS line if no address exists in AD
    $signature = $rawTemplate -replace "WEBSITE_SEG",""
    $rawTemplate = $signature
}

# Gets the last update time of the template.
if (test-path $templateFilePath){
    $templateLastModifiedDate = [datetime](Get-ItemProperty -Path $templateFilePath -Name LastWriteTime).lastwritetime
}

# Checks if there is a signature and its last update time
if (test-path $filename){
    $signatureLastModifiedDate = [datetime](Get-ItemProperty -Path $filename -Name LastWriteTime).lastwritetime
    if ((get-date $templateLastModifiedDate) -gt (get-date $signatureLastModifiedDate)){
        $signature > $fileName
    }
} else {
    $signature > $fileName
}
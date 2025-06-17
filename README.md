
# SetOutlookSignature.ps1

Pull user attributes from Active Directory to build an email signaure based on an html template.
*Note: It is intended that this script is implemented using Group Policy, and executed on the local machine on user sign in*

## Info

**FILES:** <br>
SetOutlookSignature.ps1  <br>
README.htm            <br>   
README.txt  <br>             
template.htm <br>                    

**AD ATTRIBUTES:** <br>
Display name      : displayName <br>
Job Title         : title <br>
Department        : department <br>
Company           : company <br>
Office            : physicalDeliveryOfficeName <br>
Street            : streetAddress <br>
City              : l <br>
State/Province    : st <br>
Zip/Postal Code   : postalCode <br>      
Telephone number  : telephoneNumber <br>  
Mobile            : mobile <br>
E-mail            : mail

*Note: The script will only utilize attributes that the User has in the Domain*
## MODIFYING THE SIGNATURE
### Single attribute line

**FILES:** <br>
SetOutlookSignature.ps1 <br>
template.htm  

**template.htm:** <br>
Create an arbitrary name for the segment, "NEWVAL_SEG" for example,   
and place it among the rest of the *_SEG lines in the order you would
like to see them appear in the email signature.

E.g. if you want to add a user's personal website at between their mobile
phone number and email, you would place "WEBSITE_SEG" between "MOBILE_PHONE_SEG"
and "EMAIL_SEG". The name of "WEBSITE_SEG" doesn't necessarily matter, but you will
need this name later when updating SetOutlookSignature.ps1.

**SetOutlookSignature.ps1:** <br>
Underneath "Define HTML segments", paste a new segment definition with the format:
```html
$newAttributeSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
	<span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
		NEWVAL
	</span>
</p>
"@
```

You can include an image on the line by using the following format and replacing IMAGE_DATA by base64 at 
https://www.base64-image.de/ > upload image > click "show code" > copy "For use in <img> elements"
```html
$newAttributeSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
	<span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
		<img src="IMAGE_DATA" data-outlook-trace="F:1|T:1" width="12" height="12" style="width: 9.72pt; height: 9.72pt; min-width: auto; min-height: auto; margin: 0px; vertical-align: top;">
		NEWVAL
	</span>
</p>
"@
```

Replacing newAttributeSegment and NEWVAL with their respective values. 
Take note what you replace newAttributeSegment and NEWVAL by, you will need these for later.

Pull the value from the AD user by pasting the following line under "Extract AD attributes into variables"

```ps1
$newValue = $ADUser.value
```

Replacing newValue and value with their respective values. You can find the name of the AD values in 
Active Directory in a user > Attribute Editor > and finding the name of the attribute with the respective value

Under "Replace values in template file", paste in this format, replacing by the names you used before:
```ps1
# NEWVAL
if ($newValue.Count -eq 1) {
	# Add company to template
	$newAttributeSegment = $newAttributeSegment -replace "NEWVAL",$newValue

	# Update Signature
	$signature = $rawTemplate -replace "NEWVAL_SEG",$newAttributeSegment
	$rawTemplate = $signature
} else {
	# Ignore NEWVAL line if no company exists in AD
	$signature = $rawTemplate -replace "NEWVAL_SEG",""
	$rawTemplate = $signature
}
```

Replacing NEWVAL, NEWVAL_SEG, newValue, and newAttributeSegment by your updated values.

Once a user logs out and logs back in, the signature will update automatically.

### Multi-attribute line

**FILES:** <br>
&nbsp;&nbsp;SetOutlookSignature.ps1 <br>
&nbsp;&nbsp;template.htm <br>

**template.htm:** <br>

Create an arbitrary name for the segment, "NEWVAL_SEG" for example,
and place it among the rest of the *_SEG lines in the order you would
like to see them appear in the email signature.

E.g. if you want to add a user's personal website at between their mobile
phone number and email, you would place "WEBSITE_SEG" between "MOBILE_PHONE_SEG"
and "EMAIL_SEG". The name of "WEBSITE_SEG" doesn't necessarily matter, but you will
need this name later when updating SetOutlookSignature.ps1.

**SetOutlookSignature.ps1:** <br>

Underneath "Define HTML segments", create a new segment definition by modifying the code below,
replacing NEWVAL1, NEWVAL2, etc. by the attributes you would like to add, adding formatting as necessary:
```html
$newAttributeSegment = @"
<p style="text-align: left; background-color: white; margin-top: 0px; margin-bottom: 3pt;">
	<span style="font-family: &quot;Segoe UI&quot;, &quot;Segoe UI Web (West European)&quot;, &quot;Segoe UI&quot;, -apple-system, BlinkMacSystemFont, Roboto, &quot;Helvetica Neue&quot;, sans-serif; font-size: 9pt; color: rgb(53, 123, 20);">
		NEWVAL1 NEWVAL2 NEWVAL3 ...
	</span>
</p>
"@
```

You can include an image on the line using the same method as if it were a single attribute line. 

Add any necessary attributes below "Extract AD attributes into variables" using the following format:

```ps1
$newValue1 = $ADUser.value1
$newValue2 = $ADUser.value2
# etc ...
```

Under "Replace values in template file", use this format, replacing by the names you used before and adding 
values as necessary:
```ps1
# NEWVAL1, NEWVAL2, ...
if (($newValue1.Count -eq 1) -or ($newValue2.Count -eq 1) -or ... ) {
	# Check if NEWVAL1 exists and add it
	if ($newValue1.Count -eq 1) {
		$newAttributeSegment = $newAttributeSegment -replace "NEWVAL1",$newval1
	} else {
		$newAttributeSegment = $newAttributeSegment -replace "NEWVAL1",""
	}

	# Check if NEWVAL2 exists and add it
	if ($newValue2.Count -eq 1) {
		$newAttributeSegment = $newAttributeSegment -replace "NEWVAL2",$newval2
	} else {
		$newAttributeSegment = $newAttributeSegment -replace "NEWVAL2",""
	}

	# ... Add as many of these conditions as you have new values. 

	# Update Signature
	$signature = $rawTemplate -replace "NEWVAL_SEG",$newAttributeSegment
	$rawTemplate = $signature
} else {
	# Ignore JOB, DEPARTMENT line if no job/department exists in AD
	$signature = $rawTemplate -replace "NEWVAL_SEG",""
	$rawTemplate = $signature
}
```
Replacing NEWVAL1, NEWVAL2, ... newValue1, newValue2, ... newAttributeSegment, and NEWVAL_SEG 
by any values you have added.

If you need to handle for formatting a comma or hyphen in the line, you can use the following logic,
adding "COMMA" to newAttributeSegment where a comma would be replaced by:
```ps1
# Add a comma if both values exist
if (($newValue1.Count -eq 1) -and ($newValue2.Count -eq 1)) {
	$newAttributeSegment = $newAttributeSegment -replace "COMMA",","
} else {
	$newAttributeSegment = $newAttributeSegment -replace "COMMA",""
}
```
Unless the script is updated to handle these conditions dynamically, this logic will
need to handled for each pair of attributes 
e.g. newValue1 and newValue2, newValue2 and newValue3, newValue3 and newValue4, etc ...

Once a user logs out and logs back in, the signature will update automatically

Last updated: 6/17


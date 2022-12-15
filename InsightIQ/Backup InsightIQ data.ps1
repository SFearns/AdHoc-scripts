# Copyright 2022  Stephen Fearns
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
   
Clear-Host
# Display script banner
Write-Host "##################################################################################"
Write-Host "##                                                                              ##"
Write-Host "## Description:                                                                 ##"
Write-Host "##   This script will load an IE browser and instruct the InsightIQ server to   ##"
Write-Host "##   do a Datastore Export using the same settings as last time                 ##"
Write-Host "##                                                                              ##"
Write-Host "## Author:                                                                      ##"
Write-Host "##   Stephen Fearns                                                             ##"
Write-Host "##                                                                              ##"
Write-Host "## Change History:                                                              ##"
Write-Host "##   2202/12/02 - Initial Version                                               ##"
Write-Host "##                                                                              ##"
Write-Host "##################################################################################"

# Attempting to ignore invalid certificates, e.g. self signed.
# This code was borrowed from other websites several years ago, names forgotten.
    Write-Host "`nSetup the system so SSL can be ignored if invalid"
    Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
    "@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    
    # Possible another way of ignoring invalid SSL certificates
    # [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true }

# Setup the variables and display them
Write-Host "`nSetup variables for the IIQ server"
$IIQServer = "https://InsightIQ"
Write-Host "  `$IIQServer       = $($IIQServer)"

# These credentials can be loaded from an XML file.  For testing they were hard-coded
$IIQCredentials = @{
    "username" = "uname"
    "password" = "pword"
}
Write-Host "  Login Credentials"
Write-Host "    username = $($IIQCredentials.username)"
Write-Host "    password = **********"

# Call the IE application
Write-Host "`nLoad Internet Explorer"
$IEObject = New-Object -ComObject "InternetExplorer.Application"

# Make it visable for testing.  If running as a scheduled task this variable should be left as $false
Write-Host "  Make the page visable - needed for debugging"
$IEObject.Visible = $true

# Load the InsightIQ landing page
Write-Host "  Navigate to $($IIQServer)"
$IEObject.Navigate($IIQServer)

# Loop waiting for a possible browser warning page OR the login boxes appear
Start-Sleep -Seconds 5 # Wait for a page to load
while (!($IEObject.Document.getElementsByTagName("a") | Where-Object {$_.innerText -eq "Go on to the webpage (not recommended)"}) -or ($IEObject.Document.IHTMLDocument3_getElementById("username"))) {
    # Wait 100 Milliseconds
    Start-Sleep -Milliseconds 100
}

# If the browser warning page is displayed click on the link to move on
If (($IEObject.Document.getElementsByTagName("a") | Where-Object {$_.innerText -eq "Go on to the webpage (not recommended)"})) {
	# Simulate clicking on the link
    ($IEObject.Document.getElementsByTagName("a") | Where-Object {$_.innerText -eq "Go on to the webpage (not recommended)"}).Click()
}

# Now we can login to the site
Write-Host "  Login to the site"
while (!($IEObject.Document.IHTMLDocument3_getElementById("username"))) {
    # Wait 100 Milliseconds
    Start-Sleep -Milliseconds 100
}

# Waiting for the page to load by checking for elements on the page
If ($IEObject.Document.IHTMLDocument3_getElementById("username")) {
	# Populate the username field
    ($IEObject.Document.IHTMLDocument3_getElementById("username")).Value = $IIQCredentials.username

	# Populate the password field.  If imported from an XML ensure the password is plain text
    ($IEObject.Document.IHTMLDocument3_getElementById("password")).Value = $IIQCredentials.password
	
	# Simulate clicking on the "Login" button
    ($IEObject.Document.IHTMLDocument3_getElementsByTagName("input") | Where-Object {($_.iD -eq "authform")} | Select-Object -First 1).click()

	# Change to the settings area and wait for the page to load
    Write-Host "  Change to the Settings area"
    while (!($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Settings"})) {
        # Wait 100 Milliseconds
        Start-Sleep -Milliseconds 100
    }

    If ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Settings"}) {
		# Simulate clicking on the Settings button
        ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Settings"}).click()

		# Start the Export process and wait for a dialog window to appear
        Write-Host "  Click on the `"Export Datastore`" button"
        while (!($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Export Datastore"})) {
            # Wait 100 Milliseconds
            Start-Sleep -Milliseconds 100
        }
		
        If ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Export Datastore"}) {
			# Simulate clicking on the "Export Datastore" button
            ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Export Datastore"}).click()

			# Continue with the export and wait for confirmation buttons 
            Write-Host "  Click on the `"Export`" button"
            while (!($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Export"})) {
                # Wait 100 Milliseconds
                Start-Sleep -Milliseconds 100
            }
			
            If ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Export"}) {
				# Simulate clicking on the "Export" button
                ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Export"}).click()

				# Yes to continue
                Write-Host "  Click on the `"Yes`" button"
                while (!($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Yes"})) {
                    # Wait 100 Milliseconds
                    Start-Sleep -Milliseconds 100
                }
				
                If ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Yes"}) {
					# Simulate clicking on the "Yes" button
                    ($IEObject.Document.getElementsByTagName("BUTTON") | Where-Object {$_.textContent -eq "Yes"}).click()
                } else {
					# ERROR the required section didn't appear
                    Write-Host "  ERROR: No `"YES`" button"
                }
            } else {
				# ERROR the required section didn't appear
                Write-Host "  ERROR: No `"Export`" button"
            }
        } else {
			# ERROR the required section didn't appear
            Write-Host "  ERROR: No `"Export Datastore`" button"
        }
    } else {
		# ERROR the required section didn't appear
        Write-Host "  ERROR: No `"Settings`" button"
    }
} else {
	# ERROR the required section didn't appear
    Write-Host "  ERROR: No login prompts"
}

# The export will continue in the background on the InsightIQ server and doesn't need the
# browser open.  The delay was added for testing and shouldn't be a problem when scheduled.
Write-Host "  Close the browser after 60 Seconds"
Start-Sleep -Seconds 60  # Wait for a user to be able to confirm the page
$IEObject.Quit()

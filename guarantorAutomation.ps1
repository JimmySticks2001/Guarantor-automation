#
#This script will automate the task of inserting guarantors into Avatar.
#

#Before running this script make sure the newest version of the PM Data Collection Workbook is being used and that the
# Validator has been run and all rows in the Guarantors sheet have passed. If any errors make it through to this script the whole
# import process could be de-railed.
#Next, open Avatar to Guarantors/Payors and don't click anything. 
#Run the .bat file associated with this script to start the program. It will open an open file dialog where the user 
# can select the PM Workbook which contains the list of guarantors that need to be entered. After the user chooses the
# workbook this script will read the macroGuarantors sheet in the specified workbook. If there are any problems during this operation
# such as the workbook not being available to open, or the macroGuarantors sheet not being found, the script will display a meesage box
# informing the user of the error and will safely exit.
#Next the script will inform the user that Avatar needs to be open and the Guarantors/Payors form needs to be displayed. It also needs
# to be the only form open or else the tab order will be incorrect. When the user presses the OK button on the message box the script
# will being the automated process of entering in all of the guarantors listed in the Guarantors sheet in the workbook. If the 
# script needs to be stopped for any reason the user can close the command prompt window that the .bat file opened. This will kill 
# the automation process.


#Function to open an OpenFileDialog. This will allow the user to find the file in the filesystem rather than typing the path
# by hand or having the path hard-coded into the script. Takes a string representing the initial path to open to. Returns a 
# string representing a filepath.
Function Get-FileName($initialDirectory)
{
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}#end Get-FileName

#This function allows for easy removing of a file. It checks to make sure the filepath is valid before attempting to remove. 
# This eliminates the error caused by attempting to remove a file which might not exist. 
Function Remove-File($fileName)
{
    if(Test-Path -path $fileName)
    {
        Remove-Item -path $fileName
    }
}#end Remove-File

#This function checks to make sure that the string we want to send is Avatar approved and that it is able to be sent via WASP's
# Send-Keys function.
Function Send-String($string)
{
    if(!$string)    #if the string is null or empty, we need to output a space due to Send-Keys erroring due to a blank string
    {
        $string = " "
    }
    if($string -match "[(]")    #all parentheses need to be escaped 
    {
        $string = $string -replace "[(]", "{(}"
    }
    if($string -match "[)]")
    {
        $string = $string -replace "[)]", "{)}"
    }
    if($string -match ".+[ ].+")    #if the string contains a space between stuff, we need to build the apporpriate send keys string
    {
        $string = $string -replace " ", "+( )"
    }
    
    Select-Window javaw | Send-Keys $string
    Start-Sleep -Milliseconds $pauseTime

}#end Send-String



#----------------------------------End function definition-----------------------------------------------------#



#Load the Windows forms library so we can make open file dialogs and message boxes.
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

#Get current user account and set some paths.
$user = $env:USERNAME
$myDocs = [environment]::getfolderpath("mydocuments")
$toLoc = "$myDocs\WindowsPowerShell\WASP"
$fromLoc = "\\ntst.com\dfs\ProjectManagement\Share\Plexus\Plexus_Tools\WASP\*"


#Create the PowerShell modules directory, place the WASP automation .dll into the new directory, and unblock the file so this script
# can use it. If the path already exists, do nothing.
"checking for WASP module"
If (!(Test-Path $toLoc))
{
    "getting module from N:\ProjectManagement\Share\Plexus\Plexus_Tools\WASP"
    New-Item $toLoc -type directory -Force | Out-Null
    Copy-Item $fromLoc ($toLoc)
    "unblocking module"
    Unblock-File $toLoc\WASP.dll
}

#Load the module which should now be in the directory we created.
"loading WASP module"
Import-Module $toLoc

#Call the Get-FileName function defined above to get the filepath of the Excel workbook.
"awaiting user input"
$filePath = Get-FileName -initialDirectory "c:\fso" 

#If the filepath came back empty, tell the user and end the script.
if(!$filePath)
{
    [System.Windows.Forms.MessageBox]::Show("Could not get the filepath of the chosen file. Ending script." , "Error")
    Exit
}

#Create a filepath for a temporary CSV file. This file will be used as temporary storage for the data stored in the Guarantors tab
# in the workbook. It is faster for PowerShell to import data from a CSV file rather than importing directly from Excel. The
# temporary file will be placed in the user's temp directory and will be removed whenever Windows feels like it or at the end of 
# this script.
"creating temp csv filepaths"
$csvGuarantor = ($env:temp + "\Guarantor.csv")
$csvInsurance = ($env:temp + "\Insurance.csv")
Remove-File $csvGuarantor                           #remove and previous versions of the temporary files
Remove-File $csvInsurance

"opening PM workbook"
$excel = new-object -comobject excel.application    #start a new Excel COM object
$excel.Visible = $False                             #make Excel visible or not
$workbook = $excel.Workbooks.Open($filePath)        #open the workbook from the specified path

"finding macroGuarantors sheet"
#Loop throught each sheet to find the macroGuarantors sheet. This way the sheet can move around and we can still find it.
foreach($worksheetIterator in $workbook.worksheets)
{
    $temp = $worksheetIterator.name
    
    #Print the current sheet to the console, this will show the user that something is happening as it takes a while to find the sheet.
    Write-Host "`r$temp                               " -NoNewLine

    #If we found the sheet named macroGuarantors, assign it to a variable, notify the user, and exit the loop.
    if($temp -eq "macroGuarantors")
    {
        $worksheetGuarantors = $worksheetIterator
        Write-Host ""
        Write-Host "found sheet"
        break
    }
}

#If we didn't find the sheet, inform the user and exit the script.
if($worksheetGuarantors.name -ne "macroGuarantors")
{
    [System.Windows.Forms.MessageBox]::Show("Could not find the macroGuarantors tab in the chosen workbook. Ending script." , "Error")
    Exit
}

"finding macroInsurance_Charge sheet"
#Loop throught each sheet to find the macroInsurance_Charge sheet. This way the sheet can move around and we can still find it.
foreach($worksheetIterator in $workbook.worksheets)
{
    $temp = $worksheetIterator.name
    
    #Print the current sheet to the console, this will show the user that something is happening as it takes a while to find the sheet.
    Write-Host "`r$temp                               " -NoNewLine

    #If we found the sheet named D_Insureance_Charge_Cat, assign it to a variable, notify the user, and exit the loop.
    if($temp -eq "macroInsurance_Charge")
    {
        $worksheetInsurance = $worksheetIterator
        Write-Host ""
        Write-Host "found sheet"
        break
    }
}

#If we didn't find the sheet, inform the user and exit the script.
if($worksheetInsurance.name -ne "macroInsurance_Charge")
{
    [System.Windows.Forms.MessageBox]::Show("Could not find the macroInsurance_Charge tab in the chosen workbook. Ending script." , "Error")
    Exit
}

"creating temp csv files"
$worksheetGuarantors.SaveAs($csvGuarantor, 6)    #Save the macroGuarantors worksheet as a CSV file
$worksheetInsurance.SaveAs($csvInsurance, 6)     #Save the macroInsurance_Charge worksheet as a CSV file
$workbook.Saved = $True                          #Mark the workbook as saved so it won't prompt if you want to save before closing
"closing PM workbook"
$workbook.Close()                           #Close the workbook
$excel.Quit()                               #Close Excel
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null    #Release the Excel COM object because we have to
[System.GC]::Collect()            #Tell garbage collection to do it's thing
[System.GC]::WaitForPendingFinalizers()

"importing temp csv files into memory"
#Import the guarantors CSV file into a list.
$guarantorsCSV = Import-Csv -path $csvGuarantor
#$insuranceCSV = Import-Csv -path $csvInsurance

#Import the guarantors CSV file into an array.
$insuranceArray = @()
Import-Csv -path $csvInsurance |`
    ForEach-Object{
        $insuranceArray += $_.Boop
    }

#Sort the insurance charge catagories by alphabetical order, this is they way they are listed in the guarantors form.
$insuranceArray = $insuranceArray | Sort-Object

"removing temp csv files"
#Remove the csv files we created, we no longer need them.
Remove-File $csvGuarantor
Remove-File $csvInsurance

"awaiting user input"
#Tell the user that Avatar should be open to the Guarantors/Payors form.
$response = [System.Windows.Forms.MessageBox]::Show("Please start Avatar and open the Guarantors/Payors form." + [Environment]::NewLine +
                                                    "Make sure it is the only form open." +
                                                    [Environment]::NewLine + [Environment]::NewLine + 
                                                    "Press OK when ready." , "Status", 1)

#If the user hit the cancel button, end the script.
if($response -eq "CANCEL")
{
    Exit
}

#Set the pause time to 1.5 seconds (1500 milliseconds) between commands.
$pauseTime = 1500

"beginning automated entry"
""
"guarantor:"
#Make Avatar the active window.
Select-Window javaw | Set-WindowActive | Out-Null 
Start-Sleep -Milliseconds $pauseTime

#Loop through each row in the CSV file. 
For($row = 0; $row -lt $guarantorsCSV.Count; $row++)
{
    #If the first column in this row is empty, we are done with the sheet. We are assuming that the Guarantors sheet is valid and 
    # that all of the required columns are filled in. If the first required column is not filled out then we must have reached the
    # end of the completed rows. Break out of the loop.
    if(!$guarantorsCSV[$row].1)
    {
        break
    }

    Write-Host $guarantorsCSV[$row].1

    Send-String " "                     #Select Add
    Send-String $guarantorsCSV[$row].1  #New Guarantor Code - 1
    Send-String "{TAB}"                 #move to File button
    Send-String "{TAB}"                 #skip File button
    Send-String "{TAB}"                 #skip HIPAA Transaction Version
    Send-String $guarantorsCSV[$row].2  #Guarantor Name - 2
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].3  #Guarantor Name For Alpha Lookup - 3
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].8  #guarantor Address ZIP code - 8
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].4  #guarantor Address Street - 4
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].5  #guarantor Address Street - 5
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].6  #guarantor Address City - 6
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].7  #guarantor Address State - 7
    Send-String "{TAB}"
    Send-String $guarantorsCSV[$row].9  #guarantor phone number - 9
    Send-String "{ENTER}"
    Send-String $guarantorsCSV[$row].10 #Financial Class - 10
    Send-String "{ENTER}"               
    Send-String $guarantorsCSV[$row].11 #Guarantor Nature - 11
    Send-String "{ENTER}"               
    #Send-String $guarantorsCSV[$row].12 #Guarantor Plan - 12
    #Send-String "{TAB}"                 #move to Allow Customization Of Guarantor Plan YES

    if(!$guarantorsCSV[$row].12)
    {
        Send-String "{TAB}"
        Send-String "{TAB}"
    }
    else
    {
        Send-String $guarantorsCSV[$row].12 #Guarantor Plan - 12
        Send-String "{TAB}"                 #move to Allow Customization Of Guarantor Plan YES
    }

    #Allow Customization Of Guarantor Plan
    if($guarantorsCSV[$row].13 -eq "Y")     #if Allow Customization is Y
    {
        Send-String " "
    }
    elseif($guarantorsCSV[$row].13 -eq "N") #if Allow Customization is N
    {
        Send-String "{RIGHT}"
        Send-String " "
    }
    else                                    #if Allow Customization is blank or anything else
    {
        Send-String "{TAB}"
        Send-String "{TAB}"
        Send-String "{TAB}"
    }

    Send-String $guarantorsCSV[$row].14 #Default 'Client's Relationship to Subscriber in Financial Eligibility - 14
    Send-String "{TAB}"

    #Is This A Bad Debt Guarantor (Yes option)
    if($guarantorsCSV[$row].13 -eq "Y")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #Number of Interim Billing days for guarantor
    if(!$guarantorsCSV[$row].15)
    {
        Send-String "{TAB}"
    }
    else
    {
        Send-String $guarantorsCSV[$row].15
        Send-String "{TAB}"
    }

    #Is This A Bad Debt Guarantor (No option)
    if($guarantorsCSV[$row].13 -eq "N")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #Is This A Managed Care Contract (Yes option)
    if($guarantorsCSV[$row].16 -eq "Y")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #Use 5010/ICD-10 HCFA-1500 Claim Form (Yes option). Not in Guarantor sheet in workbook.
    Send-String "{TAB}"

    #Is This A Managed Care Contract (No option)
    if($guarantorsCSV[$row].16 -eq "N")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #Use 5010/ICD-10 HCFA-1500 Claim Form (No option). Not in Guarantor sheet in workbook.
    Send-String "{TAB}"

    #External Reviewer
    Send-String $guarantorsCSV[$row].18
    Send-String "{TAB}"

    #Sort HCFA-1500 By Program of Service? (Yes option)
    if($guarantorsCSV[$row].27 -eq "Y")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #External Reviewer - Contact Name
    Send-String $guarantorsCSV[$row].19
    Send-String "{TAB}"

    #Sort HCFA-1500 By Program of Service? (No option)
    if($guarantorsCSV[$row].27 -eq "N")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #External reviewer - Phone Number
    Send-String $guarantorsCSV[$row].20
    Send-String "{TAB}"

    #Physician Number Qualifying Code (UB-92 Record 80)
    Send-String $guarantorsCSV[$row].28
    Send-String "{TAB}"

    #Categories Available for Review
    if([string]::IsNullOrEmpty($guarantorsCSV[0].21))
    {
        Send-String "{TAB}"
    }
    else
    {
        $categories = $guarantorsCSV[$row].21
        $categoriesSplit = $categories.split("+") | Sort-Object

        $oldIndex = 0
        foreach($catagory in $categoriesSplit)
        {
            $newIndex = $insuranceArray.IndexOf($catagory)
            $currentIndex = $newIndex - $oldIndex
            $oldIndex = $newIndex

            for($i = 0; $i -lt $currentIndex; $i++)
            {
                #press down arrow $currentIndex number of times
                Select-Window javaw | Send-Keys "{DOWN}"
                Start-Sleep -Milliseconds 250
            }

            Send-String " "
        }

        Send-String "{TAB}"
    }

    #Generate No-Pay ##0 Claims (Yes, For Discharged Clients Only option)
    Send-String "{TAB}"

    #Insurance Company Number
    Send-String $guarantorsCSV[$row].23
    Send-String "{TAB}"

    #Generate No-Pay ##0 Claims (No option)
    Send-String "{TAB}"

    #Update Client Data (Yes option)
    if($guarantorsCSV[$row].22 -eq "Y")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #Generate No-Pay ##0 Claims (Yes option)
    Send-String "{TAB}"

    #Update Client Data (No option)
    if($guarantorsCSV[$row].22 -eq "N")
    {
        Send-String " "
    }
    else
    {
        Send-String "{TAB}"
    }

    #Inhibit Liability Distrobution By. Multi-select dictionary, can't navigate with powershell.
    Send-String "{TAB}"

    #Net Charge Override (Gross Charge option)
    # The workbook says that it is supposed to be limited to 20 characters but four of the five options in Avatar are clearly over
    # the 20 character limit. This is a bunch of bologna.
    Send-String "{TAB}"

    #Start Date to Use ICD-10 Format For Billing (text entry)
    Send-String "{TAB}"

    #Net Charge Override (Maximum Amount to Discharge from Service Fee/Cross Reverences - Guarantor Definitions option)
    Send-String "{TAB}"

    #Start Date to Use ICD-10 Format For Billing (Today button)
    Send-String "{TAB}"

    #Net Charge Override (Net Charge(Expected Liability) option)
    Send-String "{TAB}"

    #Start Date to Use ICD-10 Format For Billing (Yesterday button)
    Send-String "{TAB}"

    #Net Charge Override (Net Charge Less Contractual (Expected Liability Less Any Contractual Adjustments) option)
    Send-String "{TAB}"

    #Does This Guarantor Get Billed Under 3-Day Billing Rule? (No option) Not in Guarantor sheet in workbook
    #Send-String "{TAB}"

    #Use 'Claim Statement Period Start' Date as 'Claim Statement Period End' Date if ... (Yes option) Not in Guarantor sheet in workbook
    #Send-String "{TAB}"

    #Does This Guarantor Get Billed Under 3-Day Billing Rule? (Yes option) Not in Guarantor sheet in workbook
    #Send-String "{TAB}"

    #Use 'Claim Statement Perion Start' Date as 'Claim Statement Period End' Date if ... (No option) Not in Guarantor sheet in workbook
    #Send-String "{TAB}"

    #Does This Guarantor Get Billed Under 3-Day Billing Rule? (Yes(only for episodes...) option) Not in Guarantor sheet in workbook
    #Send-String "{TAB}"

    #end of form
    #Since the tab-order is not cyclical, we need to cmd-tab to get into the main Avatar menu, then tab 14 times to get to the File button.
    Select-Window javaw | Send-Keys "^({TAB})"
    Start-Sleep -Milliseconds $pauseTime

    For($i = 0; $i -lt 14; $i++)
    {
        Select-Window javaw | Send-Keys "{TAB}"
        Start-Sleep -Milliseconds 1000
    }

    #Press the File button.
    Send-String " "

    Start-Sleep -Milliseconds $pauseTime

    #After filing, the Edit radio button is selected for Add or Edit Guarantor. Move left and select Add.
    Send-String "{LEFT}"

}#end row selection loop

#Inform the user that everything went swimmingly.
[System.Windows.Forms.MessageBox]::Show("Guarantor entry complete." , "Status") | Out-Null

#All done!
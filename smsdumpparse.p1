# Import the Import-Excel module
Import-Module ImportExcel

# Specify the path to your Excel file
$excelFilePath = "C:\Path\To\Your\File.xlsx"

# Load the Excel file
$data = Import-Excel -Path $excelFilePath

# Initialize variables
$combinedText = ""
$outputData = @()

# Loop through each row in the Excel data
foreach ($row in $data) {
    # Check if the row contains a timestamp (assuming timestamp is in the format "[*]*")
    if ($row.Column1 -match '\[(\d{1,2}/\d{1,2}/\d{1,2} \d{1,2}:\d{1,2}:\d{1,2})\]') {
        # If combinedText is not empty, save it with the previous timestamp
        if ($combinedText -ne "") {
            $outputData += [PSCustomObject]@{ CombinedText = $combinedText.Trim() }
        }
        
        # Start a new combined text with the current timestamp
        $combinedText = $row.Column1
    } else {
        # Concatenate the text to the current combinedText
        $combinedText += " " + $row.Column1
    }
}

# Save the last combined text
if ($combinedText -ne "") {
    $outputData += [PSCustomObject]@{ CombinedText = $combinedText.Trim() }
}

# Export the modified data to a new Excel file
$outputData | Export-Excel -Path "C:\Path\To\Your\Output\File.xlsx" -AutoSize -Show

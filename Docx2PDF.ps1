#The doc2pdf script
$shell = new-object -comobject "WScript.Shell"
$resultyesno = $shell.popup("Delete docx file after conversion?",0,"Question",4+32)

Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
[void]$FolderBrowser.ShowDialog()
$documents_path = $FolderBrowser.SelectedPath

if (!$FolderBrowser.SelectedPath ){
echo "You did not select a folder. Exiting..."
exit
}

echo "Selected folder:" $FolderBrowser.SelectedPath

$word_app = New-Object -ComObject Word.Application

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $documents_path -Filter *.doc? -Recurse | ForEach-Object {
    #Convert to PDF
    $document = $word_app.Documents.Open($_.FullName)
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
    echo "Converted file: $($_.DirectoryName)\$($_)"

    #Delete the docx file
    if ($resultyesno -eq 6){
    $shell = new-object -comobject "Shell.Application"
    $item = $shell.Namespace(0).ParseName("$($_.DirectoryName)\$($_)")
    $item.InvokeVerb("delete")
    echo "Deleted file: $($_.DirectoryName)\$($_)"
    }
   }

$word_app.Quit()
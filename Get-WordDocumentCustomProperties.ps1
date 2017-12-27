# Source : https://blogs.technet.microsoft.com/heyscriptingguy/2009/12/30/hey-scripting-guy-how-can-i-retrieve-the-custom-properties-of-a-microsoft-word-document/

$application = New-Object -ComObject word.application
$application.Visible = $false
$document = $application.documents.open("C:dataScriptingGuys2009HSG_12_28_09Test.docx")
$binding = "System.Reflection.BindingFlags" -as [type]
$customProperties = $document.CustomDocumentProperties
foreach($Property in $customProperties)
{
 $pn = [System.__ComObject].InvokeMember("name",$binding::GetProperty,$null,$property,$null)
  trap [system.exception]
   {
     write-host -foreground blue "Value not found for $pn"
    continue
   }
  "$pn`: " +
   [System.__ComObject].InvokeMember("value",$binding::GetProperty,$null,$property,$null)

}
$application.quit() 
# Source : https://blogs.technet.microsoft.com/heyscriptingguy/2009/12/29/hey-scripting-guy-how-can-i-list-all-the-properties-of-a-microsoft-word-document/

$application = New-Object -ComObject word.application
$application.Visible = $false
$document = $application.documents.open("C:dataScriptingGuys2009HSG_12_28_09Test.docx")
$binding = "System.Reflection.BindingFlags" -as [type]
$properties = $document.BuiltInDocumentProperties
foreach($property in $properties)
{
 $pn = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$property,$null)
  trap [system.exception]
   {
     write-host -foreground blue "Value not found for $pn"
    continue
   }
  "$pn`: " +
   [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null)

}
$application.quit()
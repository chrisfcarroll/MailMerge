# This file has a bash section followed by a powershell section,
# as well as shared sections.
echo @'
' > /dev/null
#vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
# Bash Start -----------------------------------------------------------

scriptdir="`dirname "${BASH_SOURCE[0]}"`";
pushd $scriptdir
[ -f "$scriptdir/MailMerge/bin/Debug/net6.0/MailMerge.dll" ] || dotnet build

# Bash End -------------------------------------------------------------
# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
echo > /dev/null <<"out-null" ###
'@ | out-null
#vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
# Powershell Start -----------------------------------------------------

$scriptdir=$PSScriptRoot
pushd $scriptdir
if( -not (Test-Path("$scriptdir/MailMerge/bin/Debug/net8.0/MailMerge.dll")))
  { dotnet build }



# Powershell End -------------------------------------------------------
# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
out-null

echo "------------------------------------------------"
dotnet MailMerge/bin/Debug/net8.0/MailMerge.dll
echo "------------------------------------------------"
dotnet MailMerge/bin/Debug/net8.0/MailMerge.dll  --showxml MailMerge.Tests/TestDocuments/ATemplate.docx
echo "------------------------------------------------"
dotnet MailMerge/bin/Debug/net8.0/MailMerge.dll  MailMerge.Tests/TestDocuments/ATemplate.docx TestOutput.docx  FirstName=Bill  'LastName=O Reilly'

popd
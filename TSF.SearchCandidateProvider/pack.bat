nuget pack TSF.SearchCandidateProvider.csproj -Build -Properties Configuration=Release;
"C:\Program Files (x86)\Windows Kits\8.1\bin\x64\signtool.exe" sign /t http://timestamp.globalsign.com/scripts/timstamp.dll bin\Release\*.dll
nuget pack TSF.SearchCandidateProvider.csproj -Properties Configuration=Release;

#if FAKE
#r "paket:
source https://api.nuget.org/v3/index.json
source http://127.0.0.1:4000/v3/index.json
nuget CellScript.Core //"
#endif
#load "./.fake/Packages.fsx/intellisense.fsx"

$version="0.0.1"
setversion $version
$project="Bgr.Base.Excel"
dotnet build
echo $version
dotnet pack
dotnet nuget push "bin\Debug\$project.$version.nupkg"  --api-key oy2lry25sohcfh3oeckalnatx7uciube7itpwlchzj5zp4 --source https://api.nuget.org/v3/index.json
pause




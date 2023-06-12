## develop

## windows 10
### Visual Studio 2022 Community edition


## Ubuntu 22.04.2 LTS
dotnet-sdk-7.0 needed.

snap dotnet sdk segfaults, use `https://learn.microsoft.com/en-us/dotnet/core/install/linux-ubuntu#register-the-microsoft-package-repository`
at the time of writing, this is: (bash)
```
    5  declare repo_version=$(if command -v lsb_release &> /dev/null; then lsb_release -r -s; else grep -oP '(?<=^VERSION_ID=).+' /etc/os-release | tr -d '"'; fi)
    6  wget https://packages.microsoft.com/config/ubuntu/$repo_version/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
    7  sudo dpkg -i packages-microsoft-prod.deb
    8  rm packages-microsoft-prod.deb
    9  sudo apt update
   10  sudo apt install dotnet-sdk-7.0
   11  dotnet test
```




```
WebApplication2 (master) 
>> CSHARPSERVICES_DATADIR=../data/ dotnet run -v 5
....
>> curl -L -S --fail --max-time 320  --header 'Content-Type: application/json' --data '{"root":"ic_ui:investment_calculator_sheets", "input_fn":"/home/koom/repos/lodgeit-labs/accounts-assessor/0/accounts-assessor/sources/CsharpServices/data/empty IC template.xlsx", "output_fn":"nnnnnnnnn"}' http://robust10.local:17789/xlsx_to_rdf
{"result":"ok","error":null}⏎                                                                                                            

```



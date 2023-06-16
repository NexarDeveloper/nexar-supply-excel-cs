# Nexar Supply Excel Add-In

The Nexar Supply Excel Add-in enables you to access pricing and availability data from right within Microsoft Excel. You can populate part information for your BOM all at once without leaving your spreadsheet.

This Add-in is supported on the Windows platform only and the versions supported are Excel for Microsoft 365 and Excel 2016.  Specifically, the versions tested were Excel for Microsoft 365 MSO (16.0.14228.20216) 64-bit and Microsoft Excel 2016 MSO (16.0.14228.20216) 32-bit.

## Register at nexar.com
If you haven't done so already, you will need to register at [nexar.com](https://nexar.com).
1. Sign up for a new account for free and complete the simple registration process.
2. Create your Organization.
3. Create an Application ensuring you have the "Supply" scope switch enabled.
4. Go to your [Dashboard](https://portal.nexar.com) and then Manage Applications and choose Show Details to see the Client Id and Secret credentials.
5. NB: The Client Secret is confidential and should only be used to authenticate your Application and make requests on its behalf.

## Installation
1. Download the Nexar Supply Add-in binary for your system.
* [Nexar Supply Excel Add-in 64-bit](Nexar.Supply.Xll/bin/Release/NexarSupply-AddIn64-packed.xll)
* [Nexar Supply Excel Add-in 32-bit](Nexar.Supply.Xll/bin/Release/NexarSupply-AddIn-packed.xll)

2. In Excel, choose 'File -> Options -> Add-ins', then press 'Go...' to manage the 'Excel Add-ins'.
![](docs/add-ins.png?raw=true)

3. Browse for the Nexar Supply Add-in, make sure it's selected, and press 'OK'.
![](docs/install.png?raw=true)

4. To use the worksheet functions, simply type '=NEXAR_' and the list of functions will appear. Refer to the [Excel Functions](#excel-functions) for documentation on how to use these functions.
![](docs/example.png?raw=true)


## Ribbon
A new ribbon will be added to your toolbar to make use of the new functionality. 
![](docs/ribbon.png?raw=true)

The following table summarizes the actions of these commands.

| Command |	Group	| Description |
| ------- | ----- | ----------- |
| Rerun Failures | Queries | For MPNs which were not found or which returned an error, repeat the search again now |
| Force Rerun All | Queries | Executes all MPN queries again to refresh all the information |
| Update Hyperlinks | Formatting | Turn any results which are URLs into clickable links, or removes links from non-URLs |
| Refresh Login | Connect to Nexar | Once your session expires (24 hours) generate a new access token from client credentials |
| Launch nexar.com | Connect to Nexar | Open nexar.com in the user’s default browser |

Note the "Force Rerun All" command causes all `=NEXAR_...` functions to run again even if the information is up to date. This results in re-executing Nexar supply API queries again with returned parts counting against your monthly part allowance.


## Excel Functions
The following functions are available through the Add-in. You can also access the guide to each argument by clicking the Function Wizard (_fx_) after you've selected any function. Keep in mind that across most functions, the "MPN or SKU" field is required; most other fields are optional.
![](docs/using.png?raw=true)

The first function that you'll need to call in order to use the Add-in is:

`=NEXAR_SUPPLY_LOGIN("`_`ClientId`_`", "`_`ClientSecret`_`")`

_`ClientId`_ refers to your unique Nexar Supply application Client Id key as provided by Nexar.
_`ClientSecret`_ refers to your unique Nexar Supply application Client Secret key as provided by Nexar.


- Top Tip: There are two additional, optional arguments to login and if you are a self-serve customer on a subscription which doesn't include the `Datasheets` or `Lead Time` features, you'll need to pass in `TRUE` to one or both argument. If you are an enterprise customer or using the free plan no change is needed.
- `=NEXAR_SUPPLY_LOGIN("`_`ClientId`_`", "`_`ClientSecret`_`", "`_`ExcludeDatasheets`_`", "`_`ExcludeLeadTime`_`")`
  - _`ExcludeDatasheets`_ optional, defaults to `FALSE`, should be set to `TRUE` if the client is unauthorized to access datasheet data.
  - _`ExcludeLeadTime`_ optional, defaults to `FALSE`, should be set to `TRUE` if the client is unauthorized to access lead time data.

Once you've entered these correctly the result will read "_The Nexar Supply Add-in is ready!_". This login to the Nexar servers will mean your usage of supply queries from within the Add-in can be tracked and offset against your monthly quota.

Top tips for connecting:

- If you see "_Please provide your Nexar application Client Id and Secret_" this means the login function has not been correctly passed your Client Id or Client Secret. 
- If you see "_Unable to login to Nexar application, check Client Id and Secret_" there may be an error in your Client Id or Secret, your application may be deleted, or there may be some other problem logging in. Check your Client Id and Secret are correct and then try "Refresh Login". 
- If you see "_The access token has expired, please refresh login_" this is a result of a stale session. For security reasons, sessions timeout after a period of time (usually 24 hours). In this case, click the "Refresh Login" button to refresh the session. 

For accessing your Nexar supply Client Id and Client Secret visit Nexar which you can do easily with the "Launch nexar.com" button.

From here on, the world is your oyster:

```
=NEXAR_SUPPLY_AVERAGE_PRICE(...)
=NEXAR_SUPPLY_DATASHEET_URL(...)
=NEXAR_SUPPLY_DETAIL_URL(...)
=NEXAR_SUPPLY_SHORT_DESCRIPTION(...)
=NEXAR_SUPPLY_DISTRIBUTOR_LEAD_TIME(...)
=NEXAR_SUPPLY_DISTRIBUTOR_MOQ(...)
=NEXAR_SUPPLY_DISTRIBUTOR_ORDER_MUTIPLE(...)
=NEXAR_SUPPLY_DISTRIBUTOR_PACKAGING(...)
=NEXAR_SUPPLY_DISTRIBUTOR_PRICE(...)
=NEXAR_SUPPLY_DISTRIBUTOR_SKU(...)
=NEXAR_SUPPLY_DISTRIBUTOR_STOCK(...)
=NEXAR_SUPPLY_DISTRIBUTOR_STOCK_UPDATED(...)
=NEXAR_SUPPLY_DISTRIBUTOR_URL(...)
=NEXAR_SUPPLY_LOGIN(...)
=NEXAR_SUPPLY_VERSION(...)
```

For results that come in a URL format (e.g., for `=NEXAR_SUPPLY_DATASHEET_URL`, `=NEXAR_SUPPLY_DETAIL_URL` or `=NEXAR_SUPPLY_DISTRIBUTOR_URL`, click the "Update Hyperlinks" button to activate the links.

- Top Tip: Some functions allow you to specify the distributor(s).  Here you can either enter a string for a partial match on the name or you cna specify the Seller Id from the [list](https://octopart.com/api/v4/values#sellers) for an exact match.

## Sample Spreadsheet
A sample spreadsheet, including examples of functions as well as a small (and unrealistic) Bill-of-Materials (BOM), is provided. 
* [Sample Excel](samples/NexarSupplytAddInExample.xlsm)


# Building the Excel Add-in (Windows):

### Required software
  - Download and install [Visual Studio](https://www.visualstudio.com/downloads/)

### Generate XLL Add-in:
  - Visual Studio -> Open Project/Solution -> ./NexarSupplyExcelAddIn.sln
  - Ensure the Nexar.Supply.Xll project is set as the "Startup Project" (bold) 
  - Set the Build Configuration to be "Release".
  - Build the solution to generate the new NexarSupply-AddIn64-packed.xll and NexarSupply-AddIn-packed.xll binary files.
    - Top Tip: If the build fails the first time, right-click on the Solution -> Restore NuGet Packages (one time only).
  - These binaries can be found in the ".\Nexar.Supply.Xll\bin\Release" folder.
    
### Debugging
  - Debugging is easy! 
  - Change the Build Configuration setting to "Debug".
  - Choose "Start Debugging" or hit "F5" to build the solution and run the Add-in in debug mode.
    - Top Tip: If the build fails the first time, right-click on the Solution -> Restore NuGet Packages (one time only).
  - Visual Studio should already be set up to launch Excel and install the Add-in.  
  - You will need to allow the Add-in access for each debug session.
  - You can put a breakpoint in the functions, e.g. `NEXAR_SUPPLY_VERSION()` in NexarSupplyAddIn.cs.
  - Alternatively, manually install the .xll file from ".\Nexar.Supply.Xll\bin\Debug" and attach to the Excel process.

### Improvements
  - Feel free to request improvements by raising a "New issue" on the GitHub page.
  - Better still, feel free to prepare a Pull Request and Altium will review your changes and either accept them or get back to you.
  - Happy coding!
# Nexar Supply Excel Add-In

The Nexar Supply Excel Add-in enables you to access pricing and availability data from right within Microsoft Excel. You can pull part information for your BOM all at once without leaving your spreadsheet.

This Add-in is supported on the Windows platform only and the versions supported are Excel for Microsoft 365 and Excel 2016.  Specifically, the versions tested were Excel for Microsoft 365 MSO (16.0.14228.20216) 64-bit and Microsoft Excel 2016 MSO (16.0.14228.20216) 32-bit.

## Register at Nexar.com
If you haven't done so already, you will need to register at [nexar.com](https://nexar.com).
1. Sign up for a new account for free and complete the simple registration process.
2. Create your Organization.
3. Create an Application ensuring you have the "Supply" scope switch enabled.
4. Go to your Dashboard and then Manage Applications and choose Show Details to see the Client ID and Secret credentials.
5. The Client Secret is confidential and should only be used to authenticate your Application and make requests on its behalf.

## Installation
1. Download the Nexar Supply Add-in binary for your system.
* [Nexar Supply Excel Add-in 64-bit](Nexar.Supply.Xll/bin/Release/NexarSupply-AddIn64-packed.xll)
* [Nexar Supply Excel Add-in 32-bit](Nexar.Supply.Xll/bin/Release/NexarSupply-AddIn-packed.xll)

2. In Excel, choose 'File -> Options -> Add-ins', then press 'Go...' to manage the 'Excel Add-ins'.
![](docs/add-ins.png?raw=true)

3. Browse for the Nexar Supply Add-in, make sure it's selected, and press 'OK'.
![](docs/install.png?raw=true)

4. To use the worksheet functions, simply type '=NEXAR_' and the list of functions will appear. Refer to Using Functions for documentation on how to use the functions.
![](docs/example.png?raw=true)


## Ribbon
A new ribbon will be added to your toolbar to make use of the new functionality. 
![](docs/ribbon.png?raw=true)


## Excel Functions
The following functions are available through the Add-in. You can also access the guide to each argument by clicking the Function Wizard after you've selected any function. Keep in mind that across most functions, the mpn_or_sku field is required; most other fields are optional.
![](docs/using.png?raw=true)

The first function that you'll need to use to activate the Add-in is:

`=NEXAR_SUPPLY_LOGIN("_Client_Id_", "_Client_Secret_")`

'_Client_Id_' refers to your unique Nexar Supply application Client Id key as provided by Nexar.
'_Client_Secret_' refers to your unique Nexar Supply application Client Secret key as provided by Nexar.

When you've entered these correctly the result will read: `The Nexar Supply Add-in is ready!`. This login to the Nexar servers will mean your usage of supply queries from within the Add-in can be tracked.

- Top Tip: If you see "Unable to login to Nexar application, check Client Id and Secret", this could be a result of a stale session (sessions time out after a period of time to increase security). In this case, make any change to your Client Id/Secret and then paste the correct keys  back in. This will trigger another login attempt and a fresh session.

For accessing your Nexar Supply Client Id and Secret visit Nexar which you can do easily with the "Visit Nexar.com" button.

From here on, the world is your oyster:

```
=NEXAR_SUPPLY_AVERAGE_PRICE(...)
=NEXAR_SUPPLY_DETAIL_URL(...)
=NEXAR_SUPPLY_DISTRIBUTOR_LEAD_TIME(...)
=NEXAR_SUPPLY_DISTRIBUTOR_MOQ(...)
=NEXAR_SUPPLY_DISTRIBUTOR_ORDER_MUTIPLE(...)
=NEXAR_SUPPLY_DISTRIBUTOR_PACKAGING(...)
=NEXAR_SUPPLY_DISTRIBUTOR_PRICE(...)
=NEXAR_SUPPLY_DISTRIBUTOR_SKU(...)
=NEXAR_SUPPLY_DISTRIBUTOR_STOCK(...)
=NEXAR_SUPPLY_DISTRIBUTOR_URL(...)
=NEXAR_SUPPLY_LOGIN(...)
=NEXAR_SUPPLY_VERSION(...)
```

For results that come in a URL format (e.g., for `=NEXAR_SUPPLY_DETAIL_URL` or `=NEXAR_SUPPLY_DISTRIBUTOR_URL`, click the "Format Hyperlinks" button to activate the links.

## Sample Spreadsheet
A sample spreadsheet, including examples of functions and a  small BOM, is provided. 
* [Sample Excel](samples/NexarSupplytAddInExample.xlsm)


# Building the Excel Add-in (Windows):

### Required software
  Download and install [Visual Studio](https://www.visualstudio.com/downloads/)

### Generate XLL Add-in:
  Visual Studio -> Open Project/Solution -> ./NexarSupplyExcelAddIn.sln
  - Ensure the Nexar.Supply.Xll project is set as the "Startup Project" (bold) 
  - Set the Build Configuration to be "Release".
  - Build the solution to generate the new NexarSupply-AddIn64-packed.xll and NexarSupply-AddIn-packed.xll files.
    - Top Tip: If the build fails the first time, right-click on the Solution -> Restore NuGet Packages (one time only).
  - These can be found in the ".\Nexar.Supply.Xll\bin\Release" folder.
    
### Debugging
  Debugging is easy! 
  - Change the Build Configuration setting to "Debug".
  - Choose "Start Debugging" or hit "F5" to build the solution and run the Add-in in debug mode.
    - Top Tip: If the build fails the first time, right-click on the Solution -> Restore NuGet Packages (one time only).
  - Visual Studio should already be set up to launch Excel and install the Add-in.  
  - You will need to allow the Add-in access for each debug session.
  - You can put a breakpoint in the functions, e.g. NEXAR_SUPPLY_VERSION() in NexarSupplyAddIn.cs.
  - Alternatively, manually install the .xll file from ".\Nexar.Supply.Xll\bin\Debug" and attach to the Excel process.

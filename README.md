# Quickbooks-Excel-Add-in
An Excel add-in that allows you to import data from your Quickbooks Online account and export to CollabDb  

##Using the add-in
If you just want to use  add-in, just sideload `excel-qb.xml` manifest which will load it from an instance running on Azure Websites. 

##Project Setup Instructions
1. Clone the repo
2. Edit `app.js` to set the variables for `consumerKey` and `consumerSecret` (see instructions below on how to obtain them)
2. `npm install`
3. Side-load the add-in manifest (`excel-qb.dev.xml`)
3. `set debug=quickBooks-Excel-Add-in:*`
4. `npm start`
5. Start **Excel** and launch the Add-in via its  **QuickBooks/Connect** button on the **Home** ribbon tab

##Obtaining `consumerKey` and `consumerSecret`
To be able to connect to QuickBooks Online, you need to:
1. Register as a developer at [https://developer.intuit.com/](developer.intuit.com)
2. Once registered, sign-in [https://developer.intuit.com/](developer.intuit.com) and go to the [https://developer.intuit.com/v2/ui#/app/dashboard](apps dashboard)
3. Copy the **consumerKey** and **consumerSecret** provided into the corresponding variables in `app.js`

##Setting up CollabDb integration
Contact rolandoj@microsoft.com for details.

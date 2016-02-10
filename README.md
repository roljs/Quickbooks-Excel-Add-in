# Quickbooks-Excel-Add-in
An Excel add-in that allows you to import data from your Quickbooks Online account and export to CollabDb.  

##Using the add-in
If you just want to use  add-in, just sideload `excel-qb.xml` manifest which will load it from an instance running on Azure Websites. 

##Project Setup Instructions
1. Clone the repo
2. Edit `app.js` to set the variables for `consumerKey` and `consumerSecret` (see instructions below on how to obtain them)
3. `npm install`
4. Side-load the add-in manifest (`excel-qb.dev.xml`)
5. `set debug=quickBooks-Excel-Add-in:*`
6. `npm start`
7. Start **Excel** and launch the Add-in via its  **QuickBooks/Connect** button on the **Home** ribbon tab

##Obtaining QuickBooks `consumerKey` and `consumerSecret`
To be able to connect to QuickBooks Online, you need to:

1. Register as a developer at [developer.intuit.com](https://developer.intuit.com/)
2. Once registered, sign-in to [developer.intuit.com](https://developer.intuit.com/) and create a new app in the [apps dashboard](https://developer.intuit.com/v2/ui#/app/dashboard)
3. Copy the **consumerKey** and **consumerSecret** from your new app into the corresponding variables in `app.js`

##Setting up CollabDb integration
Contact [rolandoj@microsoft.com](mailto:rolandoj@microsoft.com) for details.

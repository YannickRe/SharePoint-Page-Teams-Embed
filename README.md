# SharePoint Page Embed - Microsoft Teams App
Embed any SharePoint page from within your tenant (not just from the associated SharePoint site) as a tab in your Microsoft Teams team or group chat, using Single Sign-On.

## More information

Head on over to [my blog](https://blog.yannickreekmans.be/show-sharepoint-page-in-microsoft-teams/) for an overview of why this app is necessary, and how to use it.

## Building the app
The app only contains the configuration page that allows you to enter the relative URL to your SharePoint page. Feel free to use the one I host on a static website (on Azure Storage), it is already preconfigured inside the app manifest.  
Only if you'd like to self host the configuration page, you need to build the application:

``` bash
npm i -g gulp gulp-cli
gulp build
```

## Building the manifest

To create the Microsoft Teams Apps manifest, run the `manifest` Gulp task. This will generate and validate the package and finally create the package (a zip file) in the `package` folder. The manifest will be validated against the schema and dynamically populated with values from the `.env` file.
Make sure to update the SPHOSTNAME variable, and the HOSTNAME variable if you are going to self host.

``` bash
gulp manifest
```
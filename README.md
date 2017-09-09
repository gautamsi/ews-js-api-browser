`ews-js-api-browser` 
==================
## Exchange Web Service in JavaScript/TypeScript. For ionic/electron/Outlook Add-In (mail apps) and other browser process.

built for ionic, electron and Outlook Add-In (mail apps) from https://github.com/gautamsi/ews-javascript-api  
based on `ews-javascript-api@0.9.1`    

> 0.1.0 adds support for Outlook Mail Apps/Add-ins

## installing (for ionic/cordova)
`npm install ews-js-api-browser`

## building
install typescript globally `npm install -g typescript`    
then run `tsc` or `npm run build` from command prompt
run `npm run build:outlook` (windows only) to build for Outlook Add-in/Mail apps.

## use in Ionic/cordova

npm install works, use loaders like webpack. use in ionic/electron/webview, open issues at original repo. if you are using plain js/ts in cordova, see last section for this.

`ionic serve` will have CORS issue, You want to test in emulator directly. You can also enable CORS in Chrome (chrome based browsers) by installing plugin from chrome web store [Allow-Control-Allow-Origin: *](https://chrome.google.com/webstore/detail/allow-control-allow-origi/nlfbmbojpeacfghkpbjhddihlkkiljbi)

this may affect other browser based process. electron does not have this issue


webpack uglify hangs and ultimately logs following error for me, probably due to very large file, I will see if I can fix this without making single file, ultimately current implementation loads all file so not much difference. 


```
<--- Last few GCs --->

[107880:000002A361520F20]    16262 ms: Mark-sweep 1393.6 (1448.1) -> 1393.5 (1415.1) MB, 280.1 / 0.0 ms  last resort
[107880:000002A361520F20]    16472 ms: Mark-sweep 1393.5 (1415.1) -> 1393.5 (1415.1) MB, 209.9 / 0.0 ms  last resort


<--- JS stacktrace --->

==== JS stack trace =========================================

Security context: 000000E5FE6A66A1 <JS Object>
    1: RegExpReplace [native regexp.js:~356] [pc=000002C27BC3C9A4](this=000002255C9F2D61 <JS RegExp>,z=000000160878A0C1 <Very long string[3071091]>,ak=000000E5FE6C15A1 <String[1]\: \n>)
    2: 000003B7485C9C21 <Symbol: Symbol.replace>(aka [Symbol.replace]) [native regexp.js:1] [pc=000002C27C3AD329](this=000002255C9F2D61 <JS RegExp>,z=000000160878A0C1 <Very long string[3071091]>,ak=000000E5FE6...

FATAL ERROR: CALL_AND_RETRY_LAST Allocation failed - JavaScript heap out of memory
```



## How to for Outlook Apps
### Working with modules
If you are using webpack/requirejs, just install using npm and configure loaders appropriately. dependency needed: base64-js, moment-timezone & uuid. moment is also needed as dependency to moment-timezone.   
Follow examples from `ews-jsvascript-api`
### Working without module loaders
see next section

# Working with plain js/ts (no module loaders/bundlers)
you have to copy `dist/outlook/ExchangeWebService.js` (and `dist/outlook/ExchangeWebService.d.ts` if you need typing support in typescript) to your project/scripts directory, include ExchangeWebService.js in html file (`<script>` tag). This exposes a global namespace `EwsJS`.

> for outlook add-ins/mail apps, configure for outlook by caling    call `EwsJS.ConfigureForOutlook()` inside `Office.initialize`. 

specific example on how to use this way.

```ts
// The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      EwsJS.ConfigureForOutlook(); // only needed in outlook add-ins, not needed for other browser based process
    $(document).ready(function () {
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
      EwsJS.EwsLogging.DebugLogEnabled = false;
      var exch = new EwsJS.ExchangeService(EwsJS.ExchangeVersion.Exchange2013);
      exch.Credentials = new EwsJS.WebCredentials("user","password"); // any fake stuff needed, it is not used properly
      exch.Url = new EwsJS.Uri("https://outlook.office365.com/Ews/Exchange.asmx"); // anything valid url

      EwsJS.Folder.Bind(exch, new EwsJS.FolderId(EwsJS.WellKnownFolderName.MsgFolderRoot))
        .then(function (root) {
          return root.FindFolders(new EwsJS.FolderView(100));
        }).then(function (subfolders) {
          subfolders.Folders.forEach(function (f) {
              console.log(f.DisplayName + " : " + f.FolderClass + " : " + f.TotalCount); // just logging here, do what you want.
          });
      }).catch(function (err) {
          console.log("Error occurred");
          console.log(err);
      });
    });
  };
```


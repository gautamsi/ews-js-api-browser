`ews-js-api-browser` 
==================
## Exchange Web Service in JavaScript/TypeScript. For ionic/electron and other browser process.

built for ionic and electron from https://github.com/gautamsi/ews-javascript-api  
based on `ews-javascript-api@0.9.1`    

## installing
`npm install ews-js-api-browser`

## building
install typescript globally `npm install -g typescript`    
then run `tsc` or `npm run build` from command prompt


use in ionic/electron/webview, open issues at original repo

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


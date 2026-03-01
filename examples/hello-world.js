// hello-world.js - minimal open_hta() example
//
// Opens a native Windows window that displays "Hello, World!".
// The window has an OK button; clicking it closes the window and
// returns a value back to CScript, which prints it to the console.
//
// Run via:
//   examples\run.bat hello-world.js
//   cscript.exe bin\launcher.js examples\hello-world.js
//   cscript.exe dist\launcher.js examples\hello-world.js

load("core");
load("polyfills");
load("system");
load("ui");

log("Opening Hello World window...");

open_hta(
    {
        title:  "Hello World",
        width:  420,
        height: 260,

        style:
            "body   { font-family: Segoe UI, Arial, sans-serif;" +
            "         display: flex; flex-direction: column;" +
            "         align-items: center; justify-content: center;" +
            "         background: #f0f4f8; margin: 0; }" +
            "h1     { color: #2d3748; margin: 0 0 24px; font-size: 2em; }" +
            "button { padding: 10px 32px; font-size: 14px; cursor: pointer;" +
            "         background: #0078d4; color: #fff;" +
            "         border: none; border-radius: 4px; }",

        body:
            "<h1>Hello, World!</h1>" +
            "<button onclick=\"jsw_return('hello')\">OK</button>"
    },
    function(result) {
        log("Window closed. Result: " + JSON.stringify(result));
    }
);

# VbaJsBridge

Send messages between VBA and OfficeJS, execute code dynamically, interact with COM and more!

## Usage

### `cVbaJsBridge.Create(workbook as Workbook, Optional Namespace As String = "VBA-JS-Bridge") as cVbaJsBridge`

Create VbaJsBridge in a particular workbook and optionally supply a Namespace to be used to transport messages back and forth between VBA and JavaScript.

### `cVbaJsBridge#evalJS(js As String, Optional id As String) As String`

Evaluates javascript asynchronously. The return value of this method is the CallbackID used to call your VBA callback. Currently VBA callbacks are only partially implemented.


## Planned usage

### `cVbaJsBridge#evalJSSync(js As String, Optional id As String) As String`

Evaluate JavaScript synchronously.

### `cVbaJsBridge#addCallback(CallbackID As String, ObjOrModule as Variant, ProcName as String, args() as Variant, Optional CallType as String) As String`

Adds a VBA callback to the bridge. This will allow for asynchonous callbacks from JavaScript to VBA.

## Planned usage from JavaScript!

### `cVbaJsBridge.Create()`

Create a VbaJsBridge object which can be used to request proxies to VBA objects.

### `cVbaJsBridge#CreateObject([class as String],[server as String])`

Uses VBA CreateObject method to instantiate a COM object. Exposes a Proxy object which can be used to request properties and methods. E.G:

```js
let bridge  = cVbaJsBridge.Create()
let wordApp = await bridge.CreateObject("Word.Application");
let pnt     = await wordApp.CentimetersToPoints(1) //Call method
let doc     = await wordApp.$ActiveDocument //Get property - notice properties require $ prefix.
let canCIn  = await doc.CanCheckIn()
```

### `cVbaJsBridge#GetObject([ pathname as String ], [ class as String] )`

Uses VBA CreateObject method to instantiate a COM object. Exposes a Proxy object which can be used to request properties and methods. See CreateObject for usage details.

### `cVbaJsBridge#GetApplication()`

Gets the VBA `Application` object. Exposes a Proxy object which can be used to request properties and methods. See CreateObject for usage details.

### `cVbaJsBridge#Call(module,method,args)`

Call a method in a module. 



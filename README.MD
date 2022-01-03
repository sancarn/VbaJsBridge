# Simple VBA-OfficeJS bridge

This is a bare-bones sample which demonstrates how to send messages to JavaScript clients using Excel-VBA.

## Usage

* From your VBA environment download and import the `JsBridge.cls` file in the `src/vba` folder.
* In your TypeScript environment download and import the `XMLWatcher.ts` code.

From VBA:

```vb
'Requires JsBridge from here:             https://github.com/sancarn/VbaJsBridge
'ScriptLab test can be downloaded here:   https://gist.github.com/sancarn/b974b650f4b451ff2de51861af1671b1
Sub test()
  Dim js As JsBridge: Set js = JsBridge.Create("test")
  Call js.SendMessageSync("hello world")

  Dim col as new collection
  For i = 1 to 10
    'Add action to a collection so we can check it's status later
    col.add js.SendMessage("hello " & i)
  next
End Sub
```

From TypeScript we link the connection as follows:

```ts
let watcher: XMLWatcher;
async function run() {
    await Excel.run(async (context) => {
        watcher = new XMLWatcher(context, "test", function(vbaData, action) {
            //Do whatever you want with the data from VBA
            console.log(vbaData);

            //Remember to notify that the action is complete, otherwise `js.DeleteAll()` can be called
            action.finish();

            //Returning true signals to the bridge that you have handled this request.
            return true;
        });
    });
}
```

## Future of this project

In the future we'd like to add more abstractions to the system to allow definition of a full API with return values. This will likely look something like:

```ts
include {VBABridge} from "my/cdn"
vba = VBABridge.CreateWithAPI("fa08a4b0-2806-493f-85c3-de7eab8c6f30", {
    func1: ()=> console.log(1),
    func2: ()=>{/* do something else */},
    add: (a,b) => a+b
})
vba.someAPIFunction(1,2,"hello", {some:"JSON"})
```

from the TypeScript environment and 

```vb
set myAPIObject = new MyCustomAPIObject
set js = JsBridge.CreateWithAPI("fa08a4b0-2806-493f-85c3-de7eab8c6f30", myAPIObject)
Call js("func1")
Debug.Print js("add", Array(1, 2)) '3
```

from the VBA side.

## Useful links

* The JavaScript [demo](https://gist.github.com/sancarn/b974b650f4b451ff2de51861af1671b1) split out in these source files.




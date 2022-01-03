document.getElementById("run").addEventListener("click",run);

let watcher: XMLWatcher;
async function run() {
  await Excel.run(async (context) => {
    watcher = new XMLWatcher(context, "test", function(value, action) {
      console.log(value);
      action.finish();
      return true;
    });
  });
}
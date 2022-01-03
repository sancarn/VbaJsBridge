type IXMLWatcherListener = (value: string, e: IAction) => boolean;
type IAction = {
  root: Excel.CustomXmlPart;
  xml: OfficeExtension.ClientResult<string>;
  doc: HTMLElement;
  finish: () => void;
};
class XMLWatcher {
  private context: Excel.RequestContext;
  private uid: string;
  private interval: number;
  private listeners: IXMLWatcherListener[];
  private running: boolean;
  constructor(context: Excel.RequestContext, sUID: string, listener?: IXMLWatcherListener) {
    this.context = context;
    this.uid = sUID;
    this.listeners = [];
    if (listener){
      this.addListener(listener);
      this.start()
    }
  }
  addListener(listener: IXMLWatcherListener) {
    this.listeners.push(listener);
  }
  start() {
    if(!this.running){
      //Poll for changes every half a second
      this.interval = setInterval(async () => {
        var elements = this.context.workbook.customXmlParts.getByNamespace(this.uid);
        elements.load("values");
        await this.context.sync();

        //Get XML
        let actions: IAction[] = elements.items.map((e) => ({
          root: e,
          xml: e.getXml(),
          doc: null,
          finish: function(){
            this.doc.setAttribute("finished", "true");
            this.root.setXml(this.doc.outerHTML);
          }
        }));
        await this.context.sync();

        //Get actions as xmlDocs
        let parser = new DOMParser();
        actions.forEach((blob) => (blob.doc = parser.parseFromString(blob.xml.value, "text/xml").documentElement));
        let jsActions = actions.filter(
          (action) => action.doc.attributes["for"].value == "js" && action.doc.attributes["handled"].value == "false"
        );

        //Call listeners
        for (let action of jsActions) {
          for (let listener of this.listeners) {
            if (listener(action.doc.innerHTML, action)) {
              action.doc.setAttribute("handled", "true");
              action.root.setXml(action.doc.outerHTML);
              break;
            }
          }
        }
      }, 500);
    }
  }
  stop() {
    if(this.running){
      clearInterval(this.interval);
    }
  }
}
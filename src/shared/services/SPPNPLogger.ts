import {
    Logger,
    ConsoleListener,
    FunctionListener,
    ILogEntry,
    ILogListener,
    LogLevel
  } from "@pnp/logging";

// class CustomListener implements ILogListener {
//     log(entry: ILogEntry): void {
//       if (entry.level == LogLevel.Error)
//         console.log('Message from Custom Listener%c' + entry.message, "color:red;");
//       else if (entry.level == LogLevel.Warning)
//         console.log('Message from Custom Listener%c' + entry.message, "color:orange;");
//       else if (entry.level == LogLevel.Info)
//         console.log('Message from Custom Listener%c' + entry.message, "color:green;");
//       else
//         console.log('Message from Custom Listener%c' + entry.message, "color:blue;");
//     }
// }

export default class SPPNPLogger {

    constructor(){
        //Logger.activeLogLevel = LogLevel.Info;  //uncomment this to enable logs
        //Logger.subscribe(new ConsoleListener());
        Logger.subscribe(this._functionlistener);
        //Logger.subscribe(new CustomListener());
    }

    private _functionlistener = new FunctionListener((entry: ILogEntry) => {
        if (entry.level == LogLevel.Error)
          console.error('Message from Logger ' + entry.message);
        else if (entry.level == LogLevel.Warning)
          console.warn('Message from Logger: ' + entry.message);
        else if (entry.level == LogLevel.Info)
            console.info('Message from Logger: ' + entry.message);
          //console.info('Message from Logger %c' + entry.message, "color:green;");
        else
          console.log('Message from Logger ' + entry.message);
    });

    public warning(text: string){
        Logger.write(text, LogLevel.Warning);
    }

    public info(text: string){
        //Logger.writeJSON({ FirstName: "Ravichandran", LastName: "Krishnasamy" }, LogLevel.Info);
        Logger.write(text, LogLevel.Info);
    }

    public error(text: string){
        Logger.write(text, LogLevel.Error);
    }
}
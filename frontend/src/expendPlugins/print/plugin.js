import { arrayRemoveItem } from "../../utils/util";
import { luckysheetPrint } from "./print";
import Store from "../../store";

function print(options, config, isDemo) {
    if (luckysheetPrint) {
        arrayRemoveItem(Store.asyncLoad, "print");
        Store.luckysheetPrint = luckysheetPrint;
        Store.printPluginConfig = (options && options.config) || {};
        const link = document.createElement("link");
        link.setAttribute("rel", "stylesheet");
        link.setAttribute("type", "text/css");
        link.setAttribute("href", "./expendPlugins/print/print.css");
        document.head.appendChild(link);
    }
}

export { print };

import functionImplementation from './functionImplementation';
import Store from '../store/index'
import locale from '../locale/locale';
import getLocalizedFunctionList from '../function/getLocalizedFunctionList';

//{"0":"数学","1":"统计","2":"查找","3":"Luckysheet内置","4":"数据挖掘","5":"数据源","6":"日期","7":"过滤器","8":"财务","9":"工程计算","10":"逻辑","11":"运算符","12":"文本","13":"转换工具","14":"数组"}

const functionlist = function(customFunctions){
    let _locale = locale();
    // internationalization,get function list
    let functionListOrigin = [...getLocalizedFunctionList(_locale.functionlist)];
    // Native LuckySheet implementations of modern spreadsheet functions. They
    // are registered here rather than routed through another formula engine.
    const modern = [
        ["XLOOKUP", 3, 6], ["XMATCH", 2, 4], ["TEXTBEFORE", 2, 3], ["TEXTAFTER", 2, 3],
        ["TEXTSPLIT", 2, 3], ["TOCOL", 1, 2], ["TOROW", 1, 2], ["TAKE", 2, 3],
        ["DROP", 2, 3], ["HSTACK", 1, 254], ["VSTACK", 1, 254],
    ].map(function (definition) {
        const params = [];
        for (let i = 0; i < definition[2]; i++) {
            params.push({ example: "A1", require: i < definition[1] ? "m" : "o", repeat: i === definition[2] - 1 && definition[2] > 10 ? "y" : "n", type: "rangeall" });
        }
        return { n: definition[0], t: 14, m: [definition[1], definition[2]], p: params, d: "Native LuckySheet modern formula", a: "" };
    });
    functionListOrigin.push.apply(functionListOrigin, modern);

    // add new property f
    for (let i = 0; i < functionListOrigin.length; i++) {
        let func = functionListOrigin[i];
        func.f = functionImplementation[func.n];
    }

    if (customFunctions) {
        functionListOrigin.push(...customFunctions);
    }

    Store.functionlist = functionListOrigin;
    
    // get n property
    const luckysheet_function = {};

    for (let i = 0; i < functionListOrigin.length; i++) {
        let func = functionListOrigin[i];
        luckysheet_function[func.n] = func;
    }

    Store.luckysheet_function = luckysheet_function;
    Store.runtime.formula.functions = luckysheet_function;
}

export default functionlist;

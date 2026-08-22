import functionlist from './functionlist';
import Store from '../store';

const luckysheet_function = {};

for (let i = 0; i < functionlist.length; i++) {
    let func = functionlist[i];
    luckysheet_function[func.n] = func;
}

Store.runtime.formula.functions = luckysheet_function;

export default luckysheet_function;

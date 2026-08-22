/**
 * LuckySheet ↔ Univer OSS 条件格式映射。
 * 对外 JSON 仍以 luckysheet_conditionformat_save 的 LS 字段为主契约；
 * UV 形态仅在读写时互转，禁止改掉 type/conditionName/cellrange/format。
 *
 * 对照：univer/packages/sheets-conditional-formatting/src/base/const.ts
 * CFRuleType: highlightCell / dataBar / colorScale / iconSet
 */

export const CF_RULE_TYPE = {
    highlightCell: "highlightCell",
    dataBar: "dataBar",
    colorScale: "colorScale",
    iconSet: "iconSet",
};

export const CF_SUB_RULE_TYPE = {
    uniqueValues: "uniqueValues",
    duplicateValues: "duplicateValues",
    rank: "rank",
    text: "text",
    timePeriod: "timePeriod",
    number: "number",
    average: "average",
    formula: "formula",
};

export const CF_NUMBER_OPERATOR = {
    greaterThan: "greaterThan",
    greaterThanOrEqual: "greaterThanOrEqual",
    lessThan: "lessThan",
    lessThanOrEqual: "lessThanOrEqual",
    notBetween: "notBetween",
    between: "between",
    equal: "equal",
    notEqual: "notEqual",
};

export const CF_TEXT_OPERATOR = {
    beginsWith: "beginsWith",
    endsWith: "endsWith",
    containsText: "containsText",
    notContainsText: "notContainsText",
    equal: "equal",
    notEqual: "notEqual",
};

const LS_TYPE_TO_UV = {
    default: CF_RULE_TYPE.highlightCell,
    dataBar: CF_RULE_TYPE.dataBar,
    colorGradation: CF_RULE_TYPE.colorScale,
    icons: CF_RULE_TYPE.iconSet,
};

const UV_TYPE_TO_LS = {
    highlightCell: "default",
    dataBar: "dataBar",
    colorScale: "colorGradation",
    iconSet: "icons",
};

const LS_CONDITION_TO_UV = {
    greaterThan: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.greaterThan },
    lessThan: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.lessThan },
    betweenness: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.between },
    between: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.between },
    equal: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.equal },
    notEqual: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.notEqual },
    notBetween: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.notBetween },
    gte: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.greaterThanOrEqual },
    lte: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.lessThanOrEqual },
    greaterThanOrEqual: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.greaterThanOrEqual },
    lessThanOrEqual: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.number, operator: CF_NUMBER_OPERATOR.lessThanOrEqual },
    textContains: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.text, operator: CF_TEXT_OPERATOR.containsText },
    beginsWith: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.text, operator: CF_TEXT_OPERATOR.beginsWith },
    endsWith: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.text, operator: CF_TEXT_OPERATOR.endsWith },
    notContains: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.text, operator: CF_TEXT_OPERATOR.notContainsText },
    top10: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.rank, isPercent: false, isBottom: false },
    "top10%": { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.rank, isPercent: true, isBottom: false },
    last10: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.rank, isPercent: false, isBottom: true },
    "last10%": { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.rank, isPercent: true, isBottom: true },
    duplicateValue: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.duplicateValues },
    formula: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.formula },
    occurrenceDate: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.timePeriod },
    AboveAverage: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.average, isAbove: true },
    SubAverage: { type: CF_RULE_TYPE.highlightCell, subType: CF_SUB_RULE_TYPE.average, isAbove: false },
    // LS 扩展，UV 无对等项，回移时原样保留
    regExp: { type: CF_RULE_TYPE.highlightCell, subType: "regExp", lsExtension: true },
    sort: { type: CF_RULE_TYPE.highlightCell, subType: "sort", lsExtension: true },
};

const UV_NUMBER_OP_TO_LS = {
    greaterThan: "greaterThan",
    greaterThanOrEqual: "gte",
    lessThan: "lessThan",
    lessThanOrEqual: "lte",
    between: "betweenness",
    notBetween: "notBetween",
    equal: "equal",
    notEqual: "notEqual",
};

const UV_TEXT_OP_TO_LS = {
    beginsWith: "beginsWith",
    endsWith: "endsWith",
    containsText: "textContains",
    notContainsText: "notContains",
    equal: "equal",
    notEqual: "notEqual",
};

export function lsTypeToUv(type) {
    return LS_TYPE_TO_UV[type] || CF_RULE_TYPE.highlightCell;
}

export function uvTypeToLs(type) {
    return UV_TYPE_TO_LS[type] || "default";
}

export function lsConditionToUv(conditionName) {
    return LS_CONDITION_TO_UV[conditionName] || {
        type: CF_RULE_TYPE.highlightCell,
        subType: CF_SUB_RULE_TYPE.number,
        operator: conditionName,
    };
}

export function isUvRule(rule) {
    return !!(rule && rule.rule && typeof rule.rule.type === "string" && (rule.ranges != null || rule.cfId != null));
}

export function isLsRule(rule) {
    return !!(rule && typeof rule.type === "string" && (rule.cellrange != null || rule.conditionName != null));
}

function uvRangesToCellrange(ranges) {
    if (!ranges || !ranges.length) {
        return [];
    }
    return ranges.map((range) => {
        if (range && range.row && range.column) {
            return { row: range.row.slice(), column: range.column.slice() };
        }
        const startRow = range.startRow != null ? range.startRow : 0;
        const endRow = range.endRow != null ? range.endRow : startRow;
        const startColumn = range.startColumn != null ? range.startColumn : 0;
        const endColumn = range.endColumn != null ? range.endColumn : startColumn;
        return { row: [startRow, endRow], column: [startColumn, endColumn] };
    });
}

function cellrangeToUvRanges(cellrange) {
    if (!cellrange || !cellrange.length) {
        return [];
    }
    return cellrange.map((range) => ({
        startRow: range.row[0],
        endRow: range.row[1],
        startColumn: range.column[0],
        endColumn: range.column[1],
    }));
}

function uvStyleToLsFormat(style) {
    if (!style) {
        return { textColor: null, cellColor: null };
    }
    const textColor = style.cl && (style.cl.rgb || style.cl) || style.textColor || null;
    const cellColor = style.bg && (style.bg.rgb || style.bg) || style.cellColor || null;
    return { textColor, cellColor };
}

function lsFormatToUvStyle(format) {
    if (!format || Array.isArray(format)) {
        return format || {};
    }
    const style = {};
    if (format.textColor) {
        style.cl = { rgb: format.textColor };
    }
    if (format.cellColor) {
        style.bg = { rgb: format.cellColor };
    }
    return style;
}

export function uvConditionToLsName(uvRule) {
    if (!uvRule) {
        return "equal";
    }
    const subType = uvRule.subType;
    const operator = uvRule.operator;

    if (subType === CF_SUB_RULE_TYPE.formula) {
        return "formula";
    }
    if (subType === CF_SUB_RULE_TYPE.timePeriod) {
        return "occurrenceDate";
    }
    if (subType === CF_SUB_RULE_TYPE.duplicateValues) {
        return "duplicateValue";
    }
    if (subType === CF_SUB_RULE_TYPE.uniqueValues) {
        return "duplicateValue";
    }
    if (subType === CF_SUB_RULE_TYPE.average) {
        return uvRule.isAbove === false ? "SubAverage" : "AboveAverage";
    }
    if (subType === CF_SUB_RULE_TYPE.rank) {
        if (uvRule.isBottom && uvRule.isPercent) {
            return "last10%";
        }
        if (uvRule.isBottom) {
            return "last10";
        }
        if (uvRule.isPercent) {
            return "top10%";
        }
        return "top10";
    }
    if (subType === CF_SUB_RULE_TYPE.text) {
        return UV_TEXT_OP_TO_LS[operator] || "textContains";
    }
    if (subType === "regExp" || subType === "sort") {
        return subType;
    }
    return UV_NUMBER_OP_TO_LS[operator] || operator || "equal";
}

export function toUniverRule(lsRule) {
    if (!lsRule) {
        return null;
    }
    if (isUvRule(lsRule)) {
        return lsRule;
    }

    const uvType = lsTypeToUv(lsRule.type);
    const mapped = lsConditionToUv(lsRule.conditionName);
    const uv = {
        cfId: lsRule.cfId,
        ranges: cellrangeToUvRanges(lsRule.cellrange),
        stopIfTrue: !!lsRule.stopIfTrue,
        rule: {
            type: uvType,
        },
    };

    if (uvType === CF_RULE_TYPE.highlightCell) {
        uv.rule.subType = mapped.subType;
        if (mapped.operator) {
            uv.rule.operator = mapped.operator;
        }
        if (mapped.isPercent != null) {
            uv.rule.isPercent = mapped.isPercent;
        }
        if (mapped.isBottom != null) {
            uv.rule.isBottom = mapped.isBottom;
        }
        if (mapped.isAbove != null) {
            uv.rule.isAbove = mapped.isAbove;
        }
        if (mapped.lsExtension) {
            uv.rule.lsExtension = true;
        }
        uv.rule.value = lsRule.conditionValue;
        uv.rule.style = lsFormatToUvStyle(lsRule.format);
    } else if (uvType === CF_RULE_TYPE.colorScale) {
        uv.rule.config = lsRule.format;
    } else if (uvType === CF_RULE_TYPE.iconSet) {
        uv.rule.config = lsRule.format;
    } else if (uvType === CF_RULE_TYPE.dataBar) {
        uv.rule.config = lsRule.format;
    }

    return uv;
}

export function fromUniverRule(uvRule) {
    if (!uvRule) {
        return null;
    }
    if (isLsRule(uvRule) && !isUvRule(uvRule)) {
        return ensureLsRule(uvRule);
    }

    const inner = uvRule.rule || {};
    const lsType = uvTypeToLs(inner.type);
    const ls = {
        type: lsType,
        cellrange: uvRangesToCellrange(uvRule.ranges || uvRule.cellrange),
        stopIfTrue: !!uvRule.stopIfTrue,
    };

    if (lsType === "colorGradation" || lsType === "icons" || lsType === "dataBar") {
        ls.format = inner.config || inner.format || uvRule.format;
        return ls;
    }

    const conditionName = uvConditionToLsName(inner);
    ls.conditionName = conditionName;
    ls.conditionRange = uvRule.conditionRange || [];
    ls.conditionValue = inner.value != null
        ? (Array.isArray(inner.value) ? inner.value : [inner.value])
        : (uvRule.conditionValue || []);
    ls.format = uvStyleToLsFormat(inner.style) || uvRule.format || { textColor: null, cellColor: null };

    if (inner.subType === CF_SUB_RULE_TYPE.uniqueValues) {
        ls.conditionValue = ["1"];
    } else if (inner.subType === CF_SUB_RULE_TYPE.duplicateValues && (!ls.conditionValue || !ls.conditionValue.length)) {
        ls.conditionValue = ["0"];
    }

    return ls;
}

export function ensureLsRule(rule) {
    if (!rule || typeof rule !== "object") {
        return rule;
    }
    if (isUvRule(rule)) {
        return fromUniverRule(rule);
    }
    const next = rule;
    if (next.type === "colorScale") {
        next.type = "colorGradation";
    } else if (next.type === "iconSet") {
        next.type = "icons";
    } else if (next.type === "highlightCell") {
        next.type = "default";
        if (!next.conditionName && next.rule) {
            next.conditionName = uvConditionToLsName(next.rule);
        }
    }
    if (next.stopIfTrue == null) {
        next.stopIfTrue = false;
    }
    return next;
}

export function toLsRule(rule) {
    return ensureLsRule(rule);
}

export function normalizeConditionFormatSave(rules) {
    if (rules == null) {
        return [];
    }
    if (!Array.isArray(rules)) {
        return [];
    }
    return rules.map((rule) => ensureLsRule(rule)).filter(Boolean);
}

const cfAdapter = {
    CF_RULE_TYPE,
    CF_SUB_RULE_TYPE,
    CF_NUMBER_OPERATOR,
    CF_TEXT_OPERATOR,
    lsTypeToUv,
    uvTypeToLs,
    lsConditionToUv,
    uvConditionToLsName,
    isUvRule,
    isLsRule,
    toUniverRule,
    fromUniverRule,
    ensureLsRule,
    toLsRule,
    normalizeConditionFormatSave,
};

export default cfAdapter;

/**
 * LuckySheet ↔ Univer OSS 数据验证映射。
 * LS JSON 主契约仍是 dataVerification[r_c] = { type, type2, value1, value2, remote, ... }。
 * UV 形态只在读写时互转；validity(card/phone) 与 remote 作为 LS 扩展保留。
 *
 * 对照：@univerjs/core DataValidationType / DataValidationOperator
 * LIST / LIST_MULTIPLE / WHOLE / DECIMAL / DATE / ANY / CUSTOM / CHECKBOX / TEXT_LENGTH
 */

export const DV_TYPE = {
    ANY: "ANY",
    CUSTOM: "CUSTOM",
    LIST: "LIST",
    LIST_MULTIPLE: "LIST_MULTIPLE",
    CHECKBOX: "CHECKBOX",
    WHOLE: "WHOLE",
    DECIMAL: "DECIMAL",
    TEXT_LENGTH: "TEXT_LENGTH",
    DATE: "DATE",
};

export const DV_OPERATOR = {
    between: "between",
    notBetween: "notBetween",
    equal: "equal",
    notEqual: "notEqual",
    greaterThan: "greaterThan",
    lessThan: "lessThan",
    greaterThanOrEqual: "greaterThanOrEqual",
    lessThanOrEqual: "lessThanOrEqual",
};

const LS_TYPE_TO_UV = {
    dropdown: DV_TYPE.LIST,
    list_multiple: DV_TYPE.LIST_MULTIPLE,
    LIST_MULTIPLE: DV_TYPE.LIST_MULTIPLE,
    checkbox: DV_TYPE.CHECKBOX,
    number: DV_TYPE.DECIMAL,
    number_integer: DV_TYPE.WHOLE,
    number_decimal: DV_TYPE.DECIMAL,
    text_content: "TEXT",
    text_length: DV_TYPE.TEXT_LENGTH,
    date: DV_TYPE.DATE,
    any: DV_TYPE.ANY,
    ANY: DV_TYPE.ANY,
    custom: DV_TYPE.CUSTOM,
    CUSTOM: DV_TYPE.CUSTOM,
    validity: "VALIDITY",
};

const UV_TYPE_TO_LS = {
    LIST: "dropdown",
    LIST_MULTIPLE: "dropdown",
    CHECKBOX: "checkbox",
    WHOLE: "number_integer",
    DECIMAL: "number_decimal",
    TEXT: "text_content",
    TEXT_LENGTH: "text_length",
    DATE: "date",
    ANY: "any",
    CUSTOM: "custom",
    VALIDITY: "validity",
};

const LS_TYPE2_TO_UV_OP = {
    bw: DV_OPERATOR.between,
    nb: DV_OPERATOR.notBetween,
    eq: DV_OPERATOR.equal,
    ne: DV_OPERATOR.notEqual,
    gt: DV_OPERATOR.greaterThan,
    lt: DV_OPERATOR.lessThan,
    gte: DV_OPERATOR.greaterThanOrEqual,
    lte: DV_OPERATOR.lessThanOrEqual,
    bf: DV_OPERATOR.lessThan,
    nbf: DV_OPERATOR.greaterThanOrEqual,
    af: DV_OPERATOR.greaterThan,
    naf: DV_OPERATOR.lessThanOrEqual,
    include: "include",
    exclude: "exclude",
    equal: DV_OPERATOR.equal,
};

const UV_OP_TO_LS_TYPE2 = {
    between: "bw",
    notBetween: "nb",
    equal: "eq",
    notEqual: "ne",
    greaterThan: "gt",
    lessThan: "lt",
    greaterThanOrEqual: "gte",
    lessThanOrEqual: "lte",
};

const UV_DATE_OP_TO_LS = {
    lessThan: "bf",
    greaterThanOrEqual: "nbf",
    greaterThan: "af",
    lessThanOrEqual: "naf",
    between: "bw",
    notBetween: "nb",
    equal: "eq",
    notEqual: "ne",
};

export function lsTypeToUv(type, type2) {
    if (type === "dropdown" && type2) {
        return DV_TYPE.LIST_MULTIPLE;
    }
    return LS_TYPE_TO_UV[type] || type;
}

export function uvTypeToLs(type, extra) {
    if (type === DV_TYPE.LIST_MULTIPLE) {
        return { type: "dropdown", type2: true };
    }
    if (type === DV_TYPE.ANY) {
        return { type: "any", type2: extra && extra.operator || null };
    }
    if (type === DV_TYPE.CUSTOM) {
        return { type: "custom", type2: extra && extra.operator || null };
    }
    return { type: UV_TYPE_TO_LS[type] || type, type2: extra && extra.type2 };
}

export function lsDateType2ToOperator(type2) {
    return LS_TYPE2_TO_UV_OP[type2] || type2;
}

export function uvOperatorToLsDateType2(operator) {
    return UV_DATE_OP_TO_LS[operator] || UV_OP_TO_LS_TYPE2[operator] || operator;
}

export function lsType2ToOperator(type2) {
    return LS_TYPE2_TO_UV_OP[type2] || type2;
}

export function uvOperatorToLsType2(operator) {
    return UV_OP_TO_LS_TYPE2[operator] || operator;
}

export function isUvItem(item) {
    return !!(item && typeof item.type === "string" && /^[A-Z_]+$/.test(item.type) && (item.formula1 != null || item.uid != null || item.ranges != null));
}

export function toUniverItem(lsItem) {
    if (!lsItem) {
        return null;
    }
    if (isUvItem(lsItem)) {
        return lsItem;
    }

    const uvType = lsTypeToUv(lsItem.type, lsItem.type2);
    const uv = {
        type: uvType,
        formula1: lsItem.value1 != null ? String(lsItem.value1) : "",
        formula2: lsItem.value2 != null ? String(lsItem.value2) : "",
        operator: lsItem.type === "date" ? lsDateType2ToOperator(lsItem.type2) : lsType2ToOperator(lsItem.type2),
        allowBlank: lsItem.allowBlank != null ? lsItem.allowBlank : true,
    };

    if (lsItem.type === "validity" || lsItem.remote) {
        uv.lsExtension = {
            validity: lsItem.type === "validity" ? lsItem.type2 : undefined,
            remote: !!lsItem.remote,
            prohibitInput: !!lsItem.prohibitInput,
            hintShow: !!lsItem.hintShow,
            hintText: lsItem.hintText || "",
        };
    }

    return uv;
}

export function fromUniverItem(uvItem) {
    if (!uvItem) {
        return null;
    }
    if (!isUvItem(uvItem) && uvItem.type && LS_TYPE_TO_UV[uvItem.type]) {
        return ensureLsItem(uvItem);
    }

    const mapped = uvTypeToLs(uvItem.type, uvItem);
    const ls = {
        type: mapped.type,
        type2: mapped.type2 != null ? mapped.type2 : null,
        value1: uvItem.formula1 != null ? uvItem.formula1 : (uvItem.value1 || ""),
        value2: uvItem.formula2 != null ? uvItem.formula2 : (uvItem.value2 || ""),
        checked: !!uvItem.checked,
        remote: !!(uvItem.lsExtension && uvItem.lsExtension.remote) || !!uvItem.remote,
        prohibitInput: !!(uvItem.lsExtension && uvItem.lsExtension.prohibitInput) || !!uvItem.prohibitInput,
        hintShow: !!(uvItem.lsExtension && uvItem.lsExtension.hintShow) || !!uvItem.hintShow,
        hintText: (uvItem.lsExtension && uvItem.lsExtension.hintText) || uvItem.hintText || "",
    };

    if (uvItem.type === DV_TYPE.LIST_MULTIPLE) {
        ls.type = "dropdown";
        ls.type2 = true;
    } else if (uvItem.type === DV_TYPE.DATE) {
        ls.type2 = uvOperatorToLsDateType2(uvItem.operator);
    } else if (uvItem.type === DV_TYPE.WHOLE || uvItem.type === DV_TYPE.DECIMAL || uvItem.type === DV_TYPE.TEXT_LENGTH) {
        ls.type2 = uvOperatorToLsType2(uvItem.operator);
    } else if (uvItem.type === "VALIDITY" || (uvItem.lsExtension && uvItem.lsExtension.validity)) {
        ls.type = "validity";
        ls.type2 = uvItem.lsExtension.validity;
    }

    return ls;
}

export function ensureLsItem(item) {
    if (!item || typeof item !== "object") {
        return item;
    }
    if (isUvItem(item)) {
        return fromUniverItem(item);
    }

    const next = item;
    if (next.type === "LIST") {
        next.type = "dropdown";
    } else if (next.type === "LIST_MULTIPLE") {
        next.type = "dropdown";
        next.type2 = true;
    } else if (next.type === "WHOLE") {
        next.type = "number_integer";
        if (next.operator) {
            next.type2 = uvOperatorToLsType2(next.operator);
        }
    } else if (next.type === "DECIMAL") {
        next.type = "number_decimal";
        if (next.operator) {
            next.type2 = uvOperatorToLsType2(next.operator);
        }
    } else if (next.type === "ANY") {
        next.type = "any";
    } else if (next.type === "CUSTOM") {
        next.type = "custom";
        if (next.formula1 != null && (next.value1 == null || next.value1 === "")) {
            next.value1 = next.formula1;
        }
    } else if (next.type === "DATE" && next.operator) {
        next.type = "date";
        next.type2 = uvOperatorToLsDateType2(next.operator);
    }

    if (next.remote == null) {
        next.remote = false;
    }
    return next;
}

export function normalizeDataVerificationMap(map) {
    if (map == null || typeof map !== "object") {
        return map || {};
    }
    const next = {};
    Object.keys(map).forEach((key) => {
        next[key] = ensureLsItem(map[key]);
    });
    return next;
}

const dvAdapter = {
    DV_TYPE,
    DV_OPERATOR,
    lsTypeToUv,
    uvTypeToLs,
    lsDateType2ToOperator,
    uvOperatorToLsDateType2,
    lsType2ToOperator,
    uvOperatorToLsType2,
    isUvItem,
    toUniverItem,
    fromUniverItem,
    ensureLsItem,
    normalizeDataVerificationMap,
};

export default dvAdapter;

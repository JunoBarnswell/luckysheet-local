package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonInclude;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record AccessPrincipal(AccessPrincipalKind kind, String id) {
}

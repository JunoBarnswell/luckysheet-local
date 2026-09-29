package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.StructuralPatch;

/** Isolated reducer output; only structural reducers may attach a server-derived patch. */
public record MutationApplication(JsonNode snapshot, StructuralPatch structuralPatch) {
    public MutationApplication {
        if (snapshot == null || snapshot.isNull()) throw new IllegalArgumentException("Mutation application snapshot is required");
    }
}

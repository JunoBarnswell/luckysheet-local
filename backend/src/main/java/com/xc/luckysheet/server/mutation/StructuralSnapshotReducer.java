package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.AutoFilterOwnershipValidator;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiFunction;
import java.util.function.Function;

/**
 * Canonical structural reducer for a workbook JSON snapshot.
 *
 * It owns the same participants as the browser structural transform: cells,
 * ranges, table sources, objects, notes/reviews, draw anchors, pivot and
 * sparkline definitions, protections, names, and formula references. It
 * never receives a replacement snapshot from a client.
 */
final class StructuralSnapshotReducer {
    private static final int MAX_EXACT_RANGE_SEGMENTS = 256;
    private static final List<String> CHART_TEXT_FORMULA_FIELDS = List.of(
            "titleText.linkedFormula", "legend.text.linkedFormula",
            "categoryAxis.titleText.linkedFormula", "valueAxis.titleText.linkedFormula",
            "secondaryCategoryAxis.titleText.linkedFormula", "secondaryValueAxis.titleText.linkedFormula",
            "dataTable.font.linkedFormula");

    private record RuleFormulaSnapshot(
            ObjectNode rule,
            String sheetId,
            String ruleKind,
            String ruleId,
            List<RangeRef> ranges,
            Map<String, String> formulas
    ) { }

    private record FormulaChange(String before, String after) {
        boolean changed() { return before != null && after != null && !before.equals(after); }
    }

    private record RangeOwnerSnapshot(String ownerKind, String sheetId, String regionId,
            String ownerId, RangeRef range, Integer headerRow) { }

    private StructuralSnapshotReducer() {
    }

    private record DefinedNameOwnerKey(String scope, String normalizedName, String sheetId) { }

    static JsonNode applyStructuralOwnerPatch(JsonNode snapshot, StructuralPatch patch) {
        return applyStructuralOwnerPatchOnOwnedSnapshot(snapshot.deepCopy(), patch);
    }

    static JsonNode applyStructuralOwnerPatchOnOwnedSnapshot(JsonNode ownedSnapshot, StructuralPatch patch) {
        ObjectNode root = SnapshotMutationSupport.root(ownedSnapshot);
        for (StructuralPatch.FormulaOwnerDelta delta : patch.formulaOwnerDeltas()) {
            if ("formula-rule".equals(delta.kind())) {
                applyFormulaRuleOwnerDelta(root, delta);
                continue;
            }
            if ("formula-object".equals(delta.kind())) {
                applyFormulaObjectOwnerDelta(root, delta);
                continue;
            }
            StructuralPatch.CellAddress address = delta.afterAddress();
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, address.sheetId());
            ObjectNode cell = SnapshotMutationSupport.cell(sheet,
                    new SnapshotMutationSupport.CellCoordinate(address.row(), address.column()), false);
            if (cell == null) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: formula owner is missing at "
                        + address.sheetId() + "!" + address.row() + ":" + address.column());
            }
            StructuralPatch.FormulaOwnerState current = formulaOwnerState(cell);
            if (current.equals(delta.after())) {
                if (delta.after().formula() != null) cell.remove("formulaValue");
                continue;
            }
            if (!current.equals(delta.before())) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: formula owner changed at "
                        + address.sheetId() + "!" + address.row() + ":" + address.column());
            }
            setFormulaOwnerState(cell, delta.after());
        }
        applyDefinedNameOwnerDeltas(root, patch.definedNameOwnerDeltas());
        applyRangeOwnerDeltas(root, patch.rangeOwnerDeltas());
        return root;
    }

    private static void applyRangeOwnerDeltas(ObjectNode root, List<StructuralPatch.RangeOwnerDelta> deltas) {
        if (deltas.isEmpty()) return;
        Set<String> targetRegionSheets = new java.util.HashSet<>();
        Set<String> targetTableIds = new java.util.HashSet<>();
        Set<String> targetSourceIds = new java.util.HashSet<>();
        for (StructuralPatch.RangeOwnerDelta delta : deltas) {
            switch (delta.ownerKind()) {
                case "data-region" -> targetRegionSheets.add(delta.sheetId());
                case "workbook-table" -> targetTableIds.add(delta.ownerId());
                case "data-source" -> targetSourceIds.add(delta.ownerId());
                default -> throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: unsupported range-owner kind");
            }
        }

        Map<String, Map<String, ObjectNode>> regionsBySheet = new HashMap<>();
        for (String sheetId : targetRegionSheets) {
            Map<String, ObjectNode> regionsById = new HashMap<>();
            for (JsonNode raw : SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, sheetId), "dataRegions")) {
                ObjectNode region = requireObject(raw, "Data region");
                String regionId = SnapshotMutationSupport.text(region, "id");
                if (regionsById.putIfAbsent(regionId, region) != null) {
                    throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: data-region identity is duplicated: " + sheetId + ":" + regionId);
                }
            }
            regionsBySheet.put(sheetId, regionsById);
        }
        Map<String, ObjectNode> tablesById = indexedRangeOwners(workbookTables(root), targetTableIds, "Workbook table");
        Map<String, ObjectNode> sourcesById = indexedRangeOwners(
                SnapshotMutationSupport.dataModelArray(root, "sources"), targetSourceIds, "Data source");

        List<StructuralPatch.RangeOwnerDelta> changes = new ArrayList<>();
        for (StructuralPatch.RangeOwnerDelta delta : deltas) {
            RangeRef currentRange;
            Integer currentHeader = null;
            if ("data-region".equals(delta.ownerKind())) {
                ObjectNode owner = regionsBySheet.get(delta.sheetId()).get(delta.regionId());
                if (owner == null) throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: data-region owner is missing: " + delta.sheetId() + ":" + delta.regionId());
                currentRange = SnapshotMutationSupport.range(root, owner.get("range"));
                currentHeader = integer(owner.get("headerRow"), "Data region header row", SnapshotMutationSupport.MAX_ROW);
            } else {
                ObjectNode owner = ("workbook-table".equals(delta.ownerKind()) ? tablesById : sourcesById).get(delta.ownerId());
                if (owner == null || !delta.ownerId().equals(owner.path("id").asText())) {
                    throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: range owner is missing: " + delta.ownerKind() + ":" + delta.ownerId());
                }
                currentRange = SnapshotMutationSupport.range(root, owner.get("sourceRange"));
            }
            RangeRef beforeRange = delta.beforeRange();
            RangeRef afterRange = delta.afterRange();
            boolean atTarget = currentRange.equals(afterRange)
                    && (!"data-region".equals(delta.ownerKind()) || currentHeader == delta.afterHeaderRow());
            boolean atExpected = currentRange.equals(beforeRange)
                    && (!"data-region".equals(delta.ownerKind()) || currentHeader == delta.beforeHeaderRow());
            if (atTarget) continue;
            if (!atExpected) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: range owner changed since structural operation: "
                        + delta.ownerKind() + ":" + (delta.regionId() == null ? delta.ownerId() : delta.sheetId() + ":" + delta.regionId()));
            }
            changes.add(delta);
        }

        Map<String, Map<String, StructuralPatch.RangeOwnerDelta>> regionChanges = new HashMap<>();
        for (StructuralPatch.RangeOwnerDelta delta : changes) {
            switch (delta.ownerKind()) {
                case "data-region" -> regionChanges.computeIfAbsent(delta.sheetId(), ignored -> new HashMap<>()).put(delta.regionId(), delta);
                case "workbook-table" -> tablesById.get(delta.ownerId()).set("sourceRange", rangeNode(delta.afterRange()));
                case "data-source" -> sourcesById.get(delta.ownerId()).set("sourceRange", rangeNode(delta.afterRange()));
                default -> throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: unsupported range-owner kind");
            }
        }
        for (Map.Entry<String, Map<String, StructuralPatch.RangeOwnerDelta>> sheetEntry : regionChanges.entrySet()) {
            for (Map.Entry<String, StructuralPatch.RangeOwnerDelta> regionEntry : sheetEntry.getValue().entrySet()) {
                ObjectNode region = regionsBySheet.get(sheetEntry.getKey()).get(regionEntry.getKey());
                StructuralPatch.RangeOwnerDelta delta = regionEntry.getValue();
                region.set("range", rangeNode(delta.afterRange()));
                region.put("headerRow", delta.afterHeaderRow());
            }
        }
    }

    private static Map<String, ObjectNode> indexedRangeOwners(ArrayNode owners, Set<String> targetIds, String label) {
        Map<String, ObjectNode> indexed = new HashMap<>();
        if (targetIds.isEmpty()) return indexed;
        for (JsonNode raw : owners) {
            ObjectNode owner = requireObject(raw, label);
            String ownerId = SnapshotMutationSupport.text(owner, "id");
            if (!targetIds.contains(ownerId)) continue;
            if (indexed.putIfAbsent(ownerId, owner) != null) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: " + label.toLowerCase(Locale.ROOT) + " identity is duplicated: " + ownerId);
            }
        }
        return indexed;
    }

    static List<StructuralPatch.DefinedNameOwnerDelta> definedNameOwnerDeltas(JsonNode beforeModels, JsonNode afterModels) {
        Map<DefinedNameOwnerKey, StructuralPatch.DefinedNameState> beforeStates = definedNameStates(beforeModels);
        Map<DefinedNameOwnerKey, StructuralPatch.DefinedNameState> afterStates = definedNameStates(afterModels);
        if (!beforeStates.keySet().equals(afterStates.keySet())) {
            throw ServiceException.unavailable("STRUCTURAL_PATCH_INVARIANT: structural mutation changed defined-name identity membership");
        }
        List<StructuralPatch.DefinedNameOwnerDelta> deltas = new ArrayList<>();
        for (Map.Entry<DefinedNameOwnerKey, StructuralPatch.DefinedNameState> entry : beforeStates.entrySet()) {
            StructuralPatch.DefinedNameState beforeState = entry.getValue();
            StructuralPatch.DefinedNameState afterState = afterStates.get(entry.getKey());
            if (beforeState.equals(afterState)) continue;
            StructuralPatch.DefinedNameOwnerIdentity owner = new StructuralPatch.DefinedNameOwnerIdentity(
                    beforeState.scope(), beforeState.name(), beforeState.sheetId());
            deltas.add(new StructuralPatch.DefinedNameOwnerDelta(owner, beforeState, afterState));
        }
        return List.copyOf(deltas);
    }

    static StructuralPatch renameSheetTableReferences(
            JsonNode beforeSnapshot,
            JsonNode afterSnapshot,
            String sheetId,
            String tableId
    ) {
        ObjectNode before = SnapshotMutationSupport.root(beforeSnapshot);
        ObjectNode after = SnapshotMutationSupport.root(afterSnapshot);
        JsonNode oldTable = findSheetTable(before, sheetId, tableId);
        JsonNode newTable = findSheetTable(after, sheetId, tableId);
        String oldName = SnapshotMutationSupport.text(requireObject(oldTable, "Sheet Table"), "name");
        String newName = SnapshotMutationSupport.text(requireObject(newTable, "Sheet Table"), "name");
        if (oldName.equalsIgnoreCase(newName)) return null;

        Function<String, String> mapFormula = formula -> FormulaReferenceTransformer.renameTableReferences(formula, oldName, newName);
        List<StructuralPatch.FormulaOwnerDelta> formulaDeltas = new ArrayList<>();
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(before)) {
            ObjectNode sheet = requireObject(rawSheet, "Sheet");
            String ownerSheetId = SnapshotMutationSupport.text(sheet, "id");
            forEachCell(sheet, entry -> {
                ObjectNode cell = entry.cell();
                StructuralPatch.FormulaOwnerState beforeState = formulaOwnerState(cell);
                String formula = beforeState.formula() == null ? null : mapFormula.apply(beforeState.formula());
                String sourceFormula = beforeState.sourceFormula() == null ? null : mapFormula.apply(beforeState.sourceFormula());
                String barcodeFormula = beforeState.barcodeFormula() == null ? null : mapFormula.apply(beforeState.barcodeFormula());
                if (Objects.equals(formula, beforeState.formula())
                        && Objects.equals(sourceFormula, beforeState.sourceFormula())
                        && Objects.equals(barcodeFormula, beforeState.barcodeFormula())) return;
                if (hasFormulaGroupMetadata(cell)
                        && !rewritesOnlyPreservedDataTableSource(cell, beforeState, formula, sourceFormula, barcodeFormula)) {
                    throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: formula group at "
                            + ownerSheetId + "!" + entry.row() + ":" + entry.column()
                            + " requires an explicit table-reference transform");
                }
                formulaDeltas.add(new StructuralPatch.FormulaOwnerDelta(
                        "formula-cell",
                        new StructuralPatch.CellAddress(ownerSheetId, entry.row(), entry.column()),
                        new StructuralPatch.CellAddress(ownerSheetId, entry.row(), entry.column()),
                        beforeState,
                        new StructuralPatch.FormulaOwnerState(formula, sourceFormula, barcodeFormula)));
            });
        }
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(before)) {
            ObjectNode sheet = requireObject(rawSheet, "Sheet");
            appendTableRenameRuleDeltas(before, sheet, SnapshotMutationSupport.text(sheet, "id"), mapFormula, formulaDeltas);
        }

        ObjectNode transformedObjects = before.deepCopy();
        ObjectNode targetSheet = SnapshotMutationSupport.sheet(transformedObjects, sheetId);
        List<StructuralPatch.FormulaOwnerDelta> objectDeltas = rewritePersistedFormulaOwners(
                transformedObjects,
                identity(targetSheet),
                (formula, owner) -> mapFormula.apply(formula),
                ObjectNode::deepCopy,
                true);
        formulaDeltas.addAll(objectDeltas);

        JsonNode rawModels = before.get("definedNameModels");
        JsonNode transformedModels = rawModels == null ? null : rawModels.deepCopy();
        if (transformedModels instanceof ArrayNode models) {
            for (JsonNode rawModel : models) {
                ObjectNode model = requireObject(rawModel, "Defined name");
                JsonNode rawFormula = model.get("formula");
                if (rawFormula == null || !rawFormula.isTextual()) {
                    throw ServiceException.validation("Defined-name formula must be text during table rename");
                }
                String formula = rawFormula.asText();
                String rewritten = mapFormula.apply(formula);
                if (!formula.equals(rewritten)) model.put("formula", rewritten);
            }
        }
        List<StructuralPatch.DefinedNameOwnerDelta> nameDeltas = definedNameOwnerDeltas(rawModels, transformedModels);
        if (formulaDeltas.isEmpty() && nameDeltas.isEmpty()) return null;
        return new StructuralPatch(StructuralPatch.VERSION, "sheetTable.update", formulaDeltas, nameDeltas, List.of());
    }

    private static JsonNode findSheetTable(ObjectNode root, String sheetId, String tableId) {
        JsonNode match = null;
        String matchSheetId = null;
        int matchCount = 0;
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = requireObject(rawSheet, "Sheet");
            for (JsonNode table : readOptionalArray(sheet, "sheetTables")) {
                if (!tableId.equals(table.path("id").asText())) continue;
                match = table;
                matchSheetId = SnapshotMutationSupport.text(sheet, "id");
                matchCount += 1;
            }
        }
        if (matchCount != 1 || !sheetId.equals(matchSheetId)) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: Sheet Table identity must resolve exactly once: " + tableId);
        }
        return match;
    }

    private static void appendTableRenameRuleDeltas(
            ObjectNode root,
            ObjectNode sheet,
            String sheetId,
            Function<String, String> mapFormula,
            List<StructuralPatch.FormulaOwnerDelta> formulaDeltas
    ) {
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            String ruleKind = "conditionalFormats".equals(property) ? "conditional-format" : "data-validation";
            for (JsonNode rawRule : readOptionalArray(sheet, property)) {
                ObjectNode rule = requireObject(rawRule, "Range rule");
                Map<String, String> formulas = ruleFormulaFields(rule);
                if (formulas.isEmpty()) continue;
                List<RangeRef> ranges = ruleRanges(root, rule);
                String ruleId = SnapshotMutationSupport.text(rule, "id");
                for (Map.Entry<String, String> entry : formulas.entrySet()) {
                    String rewritten = mapFormula.apply(entry.getValue());
                    if (!entry.getValue().equals(rewritten)) {
                        formulaDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaRule(
                                sheetId, ruleKind, ruleId, entry.getKey(), entry.getValue(), rewritten, ranges, ranges));
                    }
                }
            }
        }
    }

    private static Map<DefinedNameOwnerKey, StructuralPatch.DefinedNameState> definedNameStates(JsonNode rawModels) {
        Map<DefinedNameOwnerKey, StructuralPatch.DefinedNameState> states = new LinkedHashMap<>();
        if (rawModels == null || rawModels.isNull()) return states;
        if (!rawModels.isArray()) throw ServiceException.validation("definedNameModels must be an array");
        for (JsonNode raw : rawModels) {
            ObjectNode model = requireObject(raw, "Defined name model");
            String name = SnapshotMutationSupport.text(model, "name");
            String formula = SnapshotMutationSupport.text(model, "formula");
            String scope = SnapshotMutationSupport.text(model, "scope");
            JsonNode sheetIdNode = model.get("sheetId");
            String sheetId = sheetIdNode == null || sheetIdNode.isNull() ? null : SnapshotMutationSupport.text(model, "sheetId");
            JsonNode rawAnchor = model.get("anchor");
            StructuralPatch.CellAddress anchor = null;
            if (rawAnchor != null && !rawAnchor.isNull()) {
                ObjectNode anchorObject = requireObject(rawAnchor, "Defined-name anchor");
                anchor = new StructuralPatch.CellAddress(
                        SnapshotMutationSupport.text(anchorObject, "sheetId"),
                        integer(anchorObject.get("row"), "Defined-name anchor row", SnapshotMutationSupport.MAX_ROW),
                        integer(anchorObject.get("column"), "Defined-name anchor column", SnapshotMutationSupport.MAX_COLUMN));
            }
            StructuralPatch.DefinedNameState state = new StructuralPatch.DefinedNameState(name, formula, scope, sheetId, anchor);
            DefinedNameOwnerKey key = definedNameOwnerKey(scope, name, sheetId);
            if (states.putIfAbsent(key, state) != null) {
                throw ServiceException.validation("Defined-name owner identity is duplicated during structural transform");
            }
        }
        return states;
    }

    private static void applyDefinedNameOwnerDeltas(
            ObjectNode root,
            List<StructuralPatch.DefinedNameOwnerDelta> deltas
    ) {
        if (deltas.isEmpty()) return;
        JsonNode rawModels = root.get("definedNameModels");
        if (rawModels == null || !rawModels.isArray()) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: defined-name models are missing");
        }
        Map<DefinedNameOwnerKey, ObjectNode> owners = new LinkedHashMap<>();
        for (JsonNode raw : rawModels) {
            if (!raw.isObject()) throw ServiceException.validation("Defined name model must be an object");
            ObjectNode model = (ObjectNode) raw;
            String name = SnapshotMutationSupport.text(model, "name");
            String scope = SnapshotMutationSupport.text(model, "scope");
            JsonNode sheetIdNode = model.get("sheetId");
            String sheetId = sheetIdNode == null || sheetIdNode.isNull() ? null : SnapshotMutationSupport.text(model, "sheetId");
            DefinedNameOwnerKey key = definedNameOwnerKey(scope, name, sheetId);
            if (owners.putIfAbsent(key, model) != null) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: defined-name owner identity is ambiguous");
            }
        }
        JsonNode rawProjection = root.get("definedNames");
        ObjectNode projection = rawProjection == null || rawProjection.isNull() ? null
                : rawProjection.isObject() ? (ObjectNode) rawProjection
                : throwInvalidDefinedNamesProjection();
        for (StructuralPatch.DefinedNameOwnerDelta delta : deltas) {
            StructuralPatch.DefinedNameOwnerIdentity owner = delta.owner();
            ObjectNode model = owners.get(definedNameOwnerKey(owner.scope(), owner.name(), owner.sheetId()));
            if (model == null) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: defined-name owner is missing: " + owner.name());
            }
            StructuralPatch.DefinedNameState current = definedNameState(model);
            if (current.equals(delta.after())) {
                if (projection != null && "workbook".equals(owner.scope())
                        && !projectionMatches(projection, owner.name(), delta.after().formula())) {
                    throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: defined-name projection changed: " + owner.name());
                }
                continue;
            }
            if (!current.equals(delta.before())) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: defined-name owner changed: " + owner.name());
            }
            if (projection != null && "workbook".equals(owner.scope())
                    && !projectionMatches(projection, owner.name(), delta.before().formula())) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: defined-name projection changed: " + owner.name());
            }
            model.put("formula", delta.after().formula());
            if (delta.after().anchor() == null) model.remove("anchor");
            else model.set("anchor", anchorNode(delta.after().anchor()));
            if (projection != null && "workbook".equals(owner.scope())) {
                projection.put(owner.name(), delta.after().formula());
            }
        }
    }

    private static ObjectNode throwInvalidDefinedNamesProjection() {
        throw ServiceException.validation("definedNames must be an object");
    }

    private static boolean projectionMatches(ObjectNode projection, String name, String formula) {
        JsonNode projected = projection.get(name);
        return projected != null && projected.isTextual() && formula.equals(projected.asText());
    }

    private static StructuralPatch.DefinedNameState definedNameState(ObjectNode model) {
        String name = SnapshotMutationSupport.text(model, "name");
        String formula = SnapshotMutationSupport.text(model, "formula");
        String scope = SnapshotMutationSupport.text(model, "scope");
        JsonNode sheetIdNode = model.get("sheetId");
        String sheetId = sheetIdNode == null || sheetIdNode.isNull() ? null : SnapshotMutationSupport.text(model, "sheetId");
        JsonNode rawAnchor = model.get("anchor");
        StructuralPatch.CellAddress anchor = rawAnchor == null || rawAnchor.isNull() ? null
                : new StructuralPatch.CellAddress(
                        SnapshotMutationSupport.text(requireObject(rawAnchor, "Defined-name anchor"), "sheetId"),
                        integer(rawAnchor.get("row"), "Defined-name anchor row", SnapshotMutationSupport.MAX_ROW),
                        integer(rawAnchor.get("column"), "Defined-name anchor column", SnapshotMutationSupport.MAX_COLUMN));
        return new StructuralPatch.DefinedNameState(name, formula, scope, sheetId, anchor);
    }

    private static int integer(JsonNode value, String label, int maximum) {
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt()) {
            throw ServiceException.validation(label + " must be an integer");
        }
        int result = value.intValue();
        if (result < 0 || result > maximum) throw ServiceException.validation(label + " is out of bounds");
        return result;
    }

    private static ObjectNode anchorNode(StructuralPatch.CellAddress address) {
        ObjectNode anchor = JsonNodeFactory.instance.objectNode();
        anchor.put("sheetId", address.sheetId());
        anchor.put("row", address.row());
        anchor.put("column", address.column());
        return anchor;
    }

    private static DefinedNameOwnerKey definedNameOwnerKey(String scope, String name, String sheetId) {
        return new DefinedNameOwnerKey(scope, name.toUpperCase(Locale.ROOT), sheetId);
    }

    private static void applyFormulaObjectOwnerDelta(ObjectNode root, StructuralPatch.FormulaOwnerDelta delta) {
        ObjectNode owner = formulaObjectOwner(root, delta);
        String property = switch (delta.ownerKind()) {
            case "chart-text" -> "linkedFormula";
            case "shape-property" -> "propertyFormula";
            case "table-sheet-column", "data-view-field" -> "formula";
            case "cell-style-template" -> "listSource.formula".equals(delta.field()) ? "formula" : delta.field();
            default -> throw ServiceException.validation("Unsupported formula-object owner kind: " + delta.ownerKind());
        };
        JsonNode rawFormula = owner.get(property);
        String current = rawFormula == null || rawFormula.isNull() ? null
                : rawFormula.isTextual() ? rawFormula.asText()
                : throwInvalidFormulaObjectValue(delta);
        if (delta.afterFormula().equals(current)) return;
        if (!delta.beforeFormula().equals(current)) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: " + delta.ownerKind()
                    + " formula owner changed at " + formulaObjectIdentity(delta));
        }
        owner.put(property, delta.afterFormula());
    }

    private static ObjectNode formulaObjectOwner(ObjectNode root, StructuralPatch.FormulaOwnerDelta delta) {
        return switch (delta.ownerKind()) {
            case "chart-text" -> {
                ObjectNode sheet = SnapshotMutationSupport.sheet(root, delta.sheetId());
                ObjectNode payloads = SnapshotMutationSupport.object(sheet, "drawingPayloads");
                JsonNode rawPayload = payloads.get(delta.ownerId());
                if (rawPayload == null || !rawPayload.isObject() || !"chart".equals(rawPayload.path("kind").asText())) {
                    throw missingFormulaObject(delta);
                }
                yield chartTextFormulaModel((ObjectNode) rawPayload, delta.field());
            }
            case "shape-property" -> {
                ObjectNode sheet = SnapshotMutationSupport.sheet(root, delta.sheetId());
                JsonNode rawPayload = SnapshotMutationSupport.object(sheet, "drawingPayloads").get(delta.ownerId());
                if (rawPayload == null || !rawPayload.isObject() || !"shape".equals(rawPayload.path("kind").asText())) {
                    throw missingFormulaObject(delta);
                }
                yield (ObjectNode) rawPayload;
            }
            case "table-sheet-column" -> {
                ObjectNode sheet = SnapshotMutationSupport.sheet(root, delta.sheetId());
                ObjectNode tableSheet = SnapshotMutationSupport.requiredObject(sheet, "tableSheet");
                List<ObjectNode> matches = new ArrayList<>();
                for (JsonNode raw : SnapshotMutationSupport.requiredArray(tableSheet, "columns")) {
                    ObjectNode column = requireObject(raw, "TableSheet column");
                    if (delta.fieldId().equals(column.path("fieldId").asText())) matches.add(column);
                }
                if (matches.size() != 1) throw missingFormulaObject(delta);
                yield matches.getFirst();
            }
            case "data-view-field" -> {
                JsonNode rawDataModel = root.get("dataModel");
                if (rawDataModel == null || !rawDataModel.isObject()) throw missingFormulaObject(delta);
                List<ObjectNode> matches = new ArrayList<>();
                for (JsonNode rawView : SnapshotMutationSupport.requiredArray((ObjectNode) rawDataModel, "views")) {
                    ObjectNode view = requireObject(rawView, "Data view");
                    if (!delta.viewId().equals(view.path("id").asText())) continue;
                    for (JsonNode rawField : SnapshotMutationSupport.requiredArray(view, "fields")) {
                        ObjectNode field = requireObject(rawField, "Data view field");
                        if (delta.fieldId().equals(field.path("fieldId").asText())) matches.add(field);
                    }
                }
                if (matches.size() != 1) throw missingFormulaObject(delta);
                yield matches.getFirst();
            }
            case "cell-style-template" -> {
                List<ObjectNode> matches = new ArrayList<>();
                for (JsonNode rawTemplate : SnapshotMutationSupport.array(root, "cellStyleTemplates")) {
                    ObjectNode template = requireObject(rawTemplate, "Cell style template");
                    if (delta.templateId().equals(template.path("id").asText())) {
                        matches.add(SnapshotMutationSupport.requiredObject(template, "dataValidation"));
                    }
                }
                if (matches.size() != 1) throw missingFormulaObject(delta);
                if ("listSource.formula".equals(delta.field())) {
                    ObjectNode listSource = SnapshotMutationSupport.requiredObject(matches.getFirst(), "listSource");
                    if (!"formula".equals(listSource.path("kind").asText())) throw missingFormulaObject(delta);
                    yield listSource;
                }
                yield matches.getFirst();
            }
            default -> throw ServiceException.validation("Unsupported formula-object owner kind: " + delta.ownerKind());
        };
    }

    private static String formulaObjectIdentity(StructuralPatch.FormulaOwnerDelta delta) {
        return switch (delta.ownerKind()) {
            case "chart-text", "shape-property" -> delta.sheetId() + ":" + delta.ownerId() + "." + delta.field();
            case "table-sheet-column" -> delta.sheetId() + ":" + delta.fieldId();
            case "data-view-field" -> delta.viewId() + ":" + delta.fieldId();
            case "cell-style-template" -> delta.templateId() + "." + delta.field();
            default -> delta.kind();
        };
    }

    private static ServiceException missingFormulaObject(StructuralPatch.FormulaOwnerDelta delta) {
        return ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: " + delta.ownerKind()
                + " formula owner is missing or ambiguous at " + formulaObjectIdentity(delta));
    }

    private static String throwInvalidFormulaObjectValue(StructuralPatch.FormulaOwnerDelta delta) {
        throw ServiceException.validation("Formula-object owner must be text at " + formulaObjectIdentity(delta));
    }

    private static void applyFormulaRuleOwnerDelta(ObjectNode root, StructuralPatch.FormulaOwnerDelta delta) {
        ObjectNode owner = SnapshotMutationSupport.sheet(root, delta.sheetId());
        String property = "conditional-format".equals(delta.ruleKind()) ? "conditionalFormats" : "dataValidations";
        List<ObjectNode> matches = new ArrayList<>();
        for (JsonNode raw : SnapshotMutationSupport.array(owner, property)) {
            ObjectNode rule = requireObject(raw, "Range rule");
            if (delta.ruleId().equals(rule.path("id").asText())
                    && delta.sheetId().equals(rule.path("sheetId").asText())) matches.add(rule);
        }
        if (matches.size() != 1) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: expected one " + delta.ruleKind()
                    + " rule " + delta.sheetId() + ":" + delta.ruleId() + ", found " + matches.size());
        }
        ObjectNode rule = matches.getFirst();
        String currentFormula = ruleFormula(rule, delta.field());
        List<RangeRef> currentRanges = ruleRanges(root, rule);
        if (!delta.afterRanges().equals(currentRanges)) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: " + delta.ruleKind() + " rule "
                    + delta.sheetId() + ":" + delta.ruleId() + "." + delta.field() + " changed since the structural operation");
        }
        if (delta.afterFormula().equals(currentFormula)) return;
        if (!delta.beforeFormula().equals(currentFormula)) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_PRECONDITION: " + delta.ruleKind() + " rule "
                    + delta.sheetId() + ":" + delta.ruleId() + "." + delta.field() + " changed since the structural operation");
        }
        setRuleFormula(rule, delta.field(), delta.afterFormula());
    }

    private static String ruleFormula(ObjectNode rule, String field) {
        if ("listSource.formula".equals(field)) {
            JsonNode source = rule.get("listSource");
            JsonNode formula = source != null && source.isObject() && "formula".equals(source.path("kind").asText())
                    ? source.get("formula") : null;
            return formula != null && formula.isTextual() ? formula.asText() : null;
        }
        JsonNode formula = rule.get(field);
        return formula != null && formula.isTextual() ? formula.asText() : null;
    }

    private static void setRuleFormula(ObjectNode rule, String field, String formula) {
        if ("listSource.formula".equals(field)) {
            JsonNode source = rule.get("listSource");
            if (source == null || !source.isObject() || !"formula".equals(source.path("kind").asText())) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: data-validation list formula owner changed type");
            }
            ((ObjectNode) source).put("formula", formula);
            return;
        }
        rule.put(field, formula);
    }

    private static List<RangeRef> ruleRanges(ObjectNode root, ObjectNode rule) {
        List<RangeRef> ranges = new ArrayList<>();
        JsonNode rawRanges = rule.get("ranges");
        if (rawRanges == null || !rawRanges.isArray()) {
            throw ServiceException.validation("Range rule ranges must be an array");
        }
        for (JsonNode raw : rawRanges) ranges.add(SnapshotMutationSupport.range(root, raw));
        return List.copyOf(ranges);
    }

    private static ArrayNode readOptionalArray(ObjectNode owner, String property) {
        JsonNode value = owner.get(property);
        if (value == null || value.isNull()) return JsonNodeFactory.instance.arrayNode();
        if (!value.isArray()) throw ServiceException.validation(property + " must be an array");
        return (ArrayNode) value;
    }

    private static void setFormulaOwnerState(ObjectNode cell, StructuralPatch.FormulaOwnerState state) {
        if (state.formula() == null) cell.remove("formula");
        else cell.put("formula", state.formula());

        JsonNode rawMetadata = cell.get("formulaMetadata");
        if (state.sourceFormula() == null) {
            if (rawMetadata != null && rawMetadata.isObject()) ((ObjectNode) rawMetadata).remove("sourceFormula");
        } else {
            if (rawMetadata == null || !rawMetadata.isObject()) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: formula provenance owner changed type");
            }
            ((ObjectNode) rawMetadata).put("sourceFormula", state.sourceFormula());
        }

        JsonNode rawPresentation = cell.get("presentation");
        JsonNode rawSource = rawPresentation != null && rawPresentation.isObject()
                && "barcode".equals(rawPresentation.path("kind").asText())
                ? rawPresentation.get("source") : null;
        if (state.barcodeFormula() != null) {
            if (rawSource == null || !rawSource.isObject() || !"formula".equals(rawSource.path("kind").asText())) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: barcode formula owner changed type");
            }
            ((ObjectNode) rawSource).put("formula", state.barcodeFormula());
        } else if (rawSource != null && rawSource.isObject()
                && "formula".equals(rawSource.path("kind").asText()) && rawSource.path("formula").isTextual()) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: barcode formula owner cannot be removed by a reference delta");
        }
        if (state.formula() != null) cell.remove("formulaValue");
    }

    static StructuralPatch applyAxis(
            ObjectNode root,
            String sheetId,
            String mutationId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode target = SnapshotMutationSupport.sheet(root, sheetId);
        int limit = dimension(target, axis);
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1;
        validateAxisBounds(limit, maximum, at, count, direction);
        validateSheetTableColumnPreservation(root, target, axis, at, direction);
        if (direction == FormulaReferenceTransformer.Direction.DELETE) validateDeletePreservation(root, target, axis, at, count);
        validateAxisDataRegionPreservation(root, target, axis, at, count, direction);
        preflightAxisFormulaAnchors(root, sheetId, axis, at, count, direction);
        if (at < limit) {
            int rowCount = dimension(target, FormulaReferenceTransformer.Axis.ROW);
            int columnCount = dimension(target, FormulaReferenceTransformer.Axis.COLUMN);
            RangeRef affectedBand = axis == FormulaReferenceTransformer.Axis.ROW
                    ? new RangeRef(sheetId, at, rowCount - 1, 0, columnCount - 1)
                    : new RangeRef(sheetId, 0, rowCount - 1, at, columnCount - 1);
            rejectFormulaGroupMetadataInRange(target, affectedBand, "axis shift");
        }
        List<RuleFormulaSnapshot> ruleFormulaSnapshots = captureRuleFormulaSnapshots(root);

        ObjectNode reportSheetAfter = mapReportSheetCoordinates(target, (row, column) -> {
            int position = axis == FormulaReferenceTransformer.Axis.ROW ? row : column;
            int mapped = shiftIndex(position, at, count, direction, axis);
            if (mapped < 0) return null;
            return axis == FormulaReferenceTransformer.Axis.ROW ? new int[]{mapped, column} : new int[]{row, mapped};
        }, axis == FormulaReferenceTransformer.Axis.ROW
                ? row -> {
                    int mapped = shiftIndex(row, at, count, direction, axis);
                    return mapped < 0 ? null : mapped;
                }
                : null,
                (direction == FormulaReferenceTransformer.Direction.INSERT ? "insert-" : "delete-")
                        + (axis == FormulaReferenceTransformer.Axis.ROW ? "rows" : "columns"));

        Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> rangeOwnersBefore = captureRangeOwnerSnapshots(root, sheetId);
        remapCells(target, axis, at, count, direction);
        setDimension(target, axis, direction == FormulaReferenceTransformer.Direction.INSERT ? limit + count : Math.max(1, limit - count));
        shiftAllMetadata(root, target, sheetId, axis, at, count, direction);
        List<StructuralPatch.RangeOwnerDelta> rangeOwnerDeltas = rangeOwnerDeltas(
                rangeOwnersBefore, captureRangeOwnerSnapshots(root, sheetId));
        applyReportSheetPlan(target, reportSheetAfter);
        List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas = rewriteAxisFormulas(
                root, target, axis, at, count, direction, ruleFormulaSnapshots);
        AutoFilterOwnershipValidator.resolveOwners(target, sheetId);
        return new StructuralPatch(StructuralPatch.VERSION, mutationId, formulaOwnerDeltas, List.of(), rangeOwnerDeltas);
    }

    static StructuralPatch shiftCells(ObjectNode root, String sheetId, String mutationId, RangeRef source, String operation, String axis, RangeRef affectedBand) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        SnapshotMutationSupport.requireSheet(source, sheetId);
        RangeRef selection = normalize(source);
        if (!"insert".equals(operation) && !"delete".equals(operation)) throw ServiceException.validation("Cell shift operation is invalid");
        if (!"row".equals(axis) && !"column".equals(axis)) throw ServiceException.validation("Cell shift axis is invalid");
        int rowCount = dimension(sheet, FormulaReferenceTransformer.Axis.ROW);
        int columnCount = dimension(sheet, FormulaReferenceTransformer.Axis.COLUMN);
        RangeRef expectedBand = "row".equals(axis)
                ? new RangeRef(sheetId, selection.startRow(), rowCount - 1, selection.startColumn(), selection.endColumn())
                : new RangeRef(sheetId, selection.startRow(), selection.endRow(), selection.startColumn(), columnCount - 1);
        if (!expectedBand.equals(normalize(affectedBand))) throw ServiceException.validation("Cell shift affected band is not canonical");
        int count = "row".equals(axis) ? selection.endRow() - selection.startRow() + 1 : selection.endColumn() - selection.startColumn() + 1;
        int delta = "insert".equals(operation) ? count : -count;
        validateCellShiftBounds(sheet, selection, expectedBand, axis, operation, count);
        RangeRef referenceBand = "row".equals(axis)
                ? new RangeRef(sheetId, selection.startRow(), SnapshotMutationSupport.MAX_ROW, selection.startColumn(), selection.endColumn())
                : new RangeRef(sheetId, selection.startRow(), selection.endRow(), selection.startColumn(), SnapshotMutationSupport.MAX_COLUMN);
        validateCellShiftDataOwners(root, sheet, referenceBand);
        rejectFormulaGroupMetadataInRange(sheet, expectedBand, "cell shift");
        FormulaReferenceTransformer.Axis shiftAxis = "row".equals(axis)
                ? FormulaReferenceTransformer.Axis.ROW : FormulaReferenceTransformer.Axis.COLUMN;
        FormulaReferenceTransformer.Direction shiftDirection = "insert".equals(operation)
                ? FormulaReferenceTransformer.Direction.INSERT : FormulaReferenceTransformer.Direction.DELETE;
        preflightCellShiftFormulaAnchors(root, sheetId, selection, shiftAxis, shiftDirection);
        ObjectNode reportSheetAfter = mapReportSheetCoordinates(sheet,
                (row, column) -> FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, formulaRange(selection), shiftAxis, shiftDirection),
                null,
                "cell-shift");

        List<RuleFormulaSnapshot> ruleFormulaSnapshots = captureRuleFormulaSnapshots(root);
        Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> rangeOwnersBefore = captureRangeOwnerSnapshots(root, sheetId);

        List<CellEntry> sourceCells = cellsInRange(sheet, expectedBand);
        SnapshotMutationSupport.clearCells(sheet, expectedBand);
        for (CellEntry entry : sourceCells) {
            int nextRow = "row".equals(axis) ? mapCellIndex(entry.row(), selection.startRow(), selection.endRow(), delta, operation) : entry.row();
            int nextColumn = "column".equals(axis) ? mapCellIndex(entry.column(), selection.startColumn(), selection.endColumn(), delta, operation) : entry.column();
            if (nextRow < 0 || nextColumn < 0 || !contains(expectedBand, nextRow, nextColumn)) continue;
            ObjectNode cell = entry.cell().deepCopy();
            SnapshotMutationSupport.putCell(sheet, new SnapshotMutationSupport.CellCoordinate(nextRow, nextColumn), cell);
        }
        shiftCellBandMetadata(root, sheet, selection, expectedBand, axis, operation, count);
        List<StructuralPatch.RangeOwnerDelta> rangeOwnerDeltas = rangeOwnerDeltas(
                rangeOwnersBefore, captureRangeOwnerSnapshots(root, sheetId));
        applyReportSheetPlan(sheet, reportSheetAfter);
        StructuralPatch structuralPatch = rewriteCellShiftFormulas(
                root, sheet, mutationId, selection, axis, operation, ruleFormulaSnapshots, rangeOwnerDeltas);
        AutoFilterOwnershipValidator.resolveOwners(sheet, sheetId);
        return structuralPatch;
    }

    static StructuralPatch moveRange(ObjectNode root, String sheetId, RangeRef source, RangeRef target) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        SnapshotMutationSupport.requireSheet(source, sheetId);
        SnapshotMutationSupport.requireSheet(target, sheetId);
        RangeRef selected = normalize(source);
        RangeRef destination = normalize(target);
        if (selected.startRow() < 0 || selected.startColumn() < 0
                || selected.endRow() > SnapshotMutationSupport.MAX_ROW
                || selected.endColumn() > SnapshotMutationSupport.MAX_COLUMN) {
            throw ServiceException.validation("Move source exceeds worksheet bounds");
        }
        int height = selected.endRow() - selected.startRow() + 1;
        int width = selected.endColumn() - selected.startColumn() + 1;
        if (destination.endRow() - destination.startRow() + 1 != height
                || destination.endColumn() - destination.startColumn() + 1 != width) {
            throw ServiceException.validation("Move destination extent does not match its source");
        }
        if (destination.endRow() > SnapshotMutationSupport.MAX_ROW
                || destination.endColumn() > SnapshotMutationSupport.MAX_COLUMN) {
            throw ServiceException.validation("Move destination exceeds worksheet bounds");
        }
        if (intersects(selected, destination)) throw ServiceException.validation("Move source and destination cannot overlap");
        rejectFormulaGroupMetadataInRange(sheet, selected, "range.move source");
        rejectFormulaGroupMetadataInRange(sheet, destination, "range.move destination");
        List<RuleFormulaSnapshot> ruleFormulaSnapshots = captureRuleFormulaSnapshots(root);
        int rowDelta = destination.startRow() - selected.startRow();
        int columnDelta = destination.startColumn() - selected.startColumn();
        ObjectNode reportSheetAfter = mapReportSheetCoordinates(sheet, (row, column) -> {
            if (contains(destination, row, column) && !contains(selected, row, column)) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: range.move overwrites report binding at " + row + ":" + column);
            }
            return contains(selected, row, column)
                    ? new int[]{row + rowDelta, column + columnDelta}
                    : new int[]{row, column};
                }, null, "range.move");

        Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> rangeOwnersBefore = captureRangeOwnerSnapshots(root, sheetId);
        List<CellEntry> cells = cellsInRange(sheet, selected);
        SnapshotMutationSupport.clearCells(sheet, selected);
        SnapshotMutationSupport.clearCells(sheet, destination);
        if (destination.endRow() >= dimension(sheet, FormulaReferenceTransformer.Axis.ROW)) {
            sheet.put("rowCount", destination.endRow() + 1);
        }
        if (destination.endColumn() >= dimension(sheet, FormulaReferenceTransformer.Axis.COLUMN)) {
            sheet.put("columnCount", destination.endColumn() + 1);
        }
        for (CellEntry entry : cells) {
            SnapshotMutationSupport.putCell(sheet,
                    new SnapshotMutationSupport.CellCoordinate(entry.row() + rowDelta, entry.column() + columnDelta),
                    entry.cell());
        }

        moveRangeMetadata(root, sheet, selected, destination, rowDelta, columnDelta);
        List<StructuralPatch.RangeOwnerDelta> rangeOwnerDeltas = rangeOwnerDeltas(
                rangeOwnersBefore, captureRangeOwnerSnapshots(root, sheetId));
        applyReportSheetPlan(sheet, reportSheetAfter);
        StructuralPatch formulaPatch = rewriteMovedFormulas(
                root, sheet, selected, destination, rowDelta, columnDelta, ruleFormulaSnapshots);
        invalidateFormulaCaches(root);
        AutoFilterOwnershipValidator.resolveOwners(sheet, sheetId);
        return new StructuralPatch(StructuralPatch.VERSION, formulaPatch.mutationId(),
                formulaPatch.formulaOwnerDeltas(), formulaPatch.definedNameOwnerDeltas(), rangeOwnerDeltas);
    }

    private static void moveRangeMetadata(ObjectNode root, ObjectNode sheet, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        String sheetId = source.sheetId();
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "merges")) {
            ObjectNode merge = requireObject(raw, "Merge");
            moveRangeReference(root, merge.get("range"), sheetId, source, target, rowDelta, columnDelta);
            movePoint(SnapshotMutationSupport.requiredObject(merge, "anchor"), source, target, rowDelta, columnDelta, false, "Merge anchor");
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "dataRegions")) {
            ObjectNode region = requireObject(raw, "Data region");
            JsonNode rangeRaw = region.get("range");
            boolean relocated = moveRangeReference(root, rangeRaw, sheetId, source, target, rowDelta, columnDelta);
            if (relocated) {
                JsonNode headerRow = region.get("headerRow");
                if (headerRow == null || !headerRow.isIntegralNumber()) throw ServiceException.validation("Data region header row is invalid");
                region.put("headerRow", headerRow.intValue() + rowDelta);
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "sheetTables")) {
            ObjectNode table = requireObject(raw, "Sheet table");
            moveRangeReference(root, table.get("range"), sheetId, source, target, rowDelta, columnDelta);
            moveAutoFilter(root, table.get("autoFilter"), sheetId, source, target, rowDelta, columnDelta);
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "spillRanges")) {
            ObjectNode spill = requireObject(raw, "Spill range");
            moveRangeReference(root, spill.get("range"), sheetId, source, target, rowDelta, columnDelta);
            movePoint(SnapshotMutationSupport.requiredObject(spill, "anchor"), source, target, rowDelta, columnDelta, true, "Spill anchor");
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "protectionRules")) {
            ObjectNode rule = requireObject(raw, "Protection rule");
            if (rule.has("range")) moveRangeReference(root, rule.get("range"), sheetId, source, target, rowDelta, columnDelta);
        }
        JsonNode bandedRaw = sheet.get("bandedRule");
        if (bandedRaw != null && bandedRaw.isObject()) {
            ObjectNode banded = (ObjectNode) bandedRaw;
            if (banded.has("range")) moveRangeReference(root, banded.get("range"), sheetId, source, target, rowDelta, columnDelta);
        }
        moveDrawings(sheet, source, target, rowDelta, columnDelta);
        SnapshotMutationSupport.remapReviewCoordinates(sheet, coordinate -> {
            if (contains(target, coordinate.row(), coordinate.column()) && !contains(source, coordinate.row(), coordinate.column())) {
                throw ServiceException.validation("Move destination would overwrite review metadata");
            }
            return contains(source, coordinate.row(), coordinate.column())
                    ? new SnapshotMutationSupport.CellCoordinate(coordinate.row() + rowDelta, coordinate.column() + columnDelta)
                    : coordinate;
        });
        moveHyperlinkAnchors(sheet, source, target, rowDelta, columnDelta);

        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            moveRules(root, owner, "conditionalFormats", sheetId, source, target, rowDelta, columnDelta);
            moveRules(root, owner, "dataValidations", sheetId, source, target, rowDelta, columnDelta);
            if (sheetId.equals(owner.path("id").asText())) moveAutoFilter(root, owner.get("autoFilter"), sheetId, source, target, rowDelta, columnDelta);
            for (JsonNode pivotRaw : SnapshotMutationSupport.array(owner, "pivots")) {
                ObjectNode pivot = requireObject(pivotRaw, "Pivot");
                PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, range ->
                        moveRangeReference(root, range, sheetId, source, target, rowDelta, columnDelta));
                ObjectNode pivotTarget = PivotMutationDescriptor.requiredTarget(pivot);
                if (sheetId.equals(pivotTarget.path("sheetId").asText())) {
                    JsonNode anchorRaw = pivotTarget.get("anchor");
                    if (anchorRaw != null && anchorRaw.isObject()) {
                        movePoint((ObjectNode) anchorRaw, source, target, rowDelta, columnDelta, true, "Pivot target anchor");
                    }
                }
            }
            for (JsonNode sparklineRaw : SnapshotMutationSupport.array(owner, "sparklines")) {
                ObjectNode sparkline = requireObject(sparklineRaw, "Sparkline");
                moveRangeReference(root, sparkline.get("sourceRange"), sheetId, source, target, rowDelta, columnDelta);
                if (sheetId.equals(sparkline.path("sheetId").asText())) {
                    movePoint(SnapshotMutationSupport.requiredObject(sparkline, "anchor"), source, target, rowDelta, columnDelta, true, "Sparkline anchor");
                }
            }
            moveDrawingPayloads(root, owner, sheetId, source, target, rowDelta, columnDelta);
        }

        for (JsonNode raw : workbookTables(root)) {
            ObjectNode table = requireObject(raw, "Workbook table");
            JsonNode sourceRange = table.get("sourceRange");
            if (sourceRange != null && !sourceRange.isNull()) moveRangeReference(root, sourceRange, sheetId, source, target, rowDelta, columnDelta);
        }
        for (JsonNode raw : SnapshotMutationSupport.dataModelArray(root, "sources")) {
            ObjectNode manifest = requireObject(raw, "Data source");
            JsonNode sourceRange = manifest.get("sourceRange");
            if (sourceRange != null && !sourceRange.isNull()) moveRangeReference(root, sourceRange, sheetId, source, target, rowDelta, columnDelta);
        }
        JsonNode documentsRaw = root.get("printDocuments");
        if (documentsRaw != null && !documentsRaw.isNull()) {
            if (!documentsRaw.isArray()) throw ServiceException.validation("printDocuments must be an array");
            for (JsonNode raw : documentsRaw) {
                ObjectNode document = requireObject(raw, "Print document");
                if (!sheetId.equals(document.path("sheetId").asText())) continue;
                for (JsonNode areaRaw : SnapshotMutationSupport.array(document, "printAreas")) {
                    ObjectNode area = requireObject(areaRaw, "Print area");
                    moveRangeReference(root, area.get("range"), sheetId, source, target, rowDelta, columnDelta);
                }
            }
        }
        moveDefinedNameAnchors(root, source, target, rowDelta, columnDelta);
    }

    private static boolean moveRangeReference(ObjectNode root, JsonNode raw, String sheetId, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Structural move participant range is invalid");
        ObjectNode value = (ObjectNode) raw;
        RangeRef range = normalize(SnapshotMutationSupport.range(root, value));
        if (!sheetId.equals(range.sheetId())) return false;
        if (intersects(source, range) && !containsRange(source, range)) {
            throw ServiceException.validation("Move partially intersects a structural participant range");
        }
        if (intersects(target, range) && !containsRange(source, range)) {
            throw ServiceException.validation("Move destination would overwrite a structural participant range");
        }
        if (!containsRange(source, range)) return false;
        value.put("startRow", range.startRow() + rowDelta);
        value.put("endRow", range.endRow() + rowDelta);
        value.put("startColumn", range.startColumn() + columnDelta);
        value.put("endColumn", range.endColumn() + columnDelta);
        return true;
    }

    private static boolean containsRange(RangeRef outer, RangeRef inner) {
        return outer.sheetId().equals(inner.sheetId())
                && outer.startRow() <= inner.startRow() && outer.endRow() >= inner.endRow()
                && outer.startColumn() <= inner.startColumn() && outer.endColumn() >= inner.endColumn();
    }

    private static void movePoint(ObjectNode point, RangeRef source, RangeRef target, int rowDelta, int columnDelta, boolean rejectTarget, String label) {
        JsonNode rowRaw = point.get("row");
        JsonNode columnRaw = point.get("column");
        if (rowRaw == null || columnRaw == null || !rowRaw.isIntegralNumber() || !columnRaw.isIntegralNumber()) {
            throw ServiceException.validation(label + " coordinates are invalid");
        }
        int row = rowRaw.intValue();
        int column = columnRaw.intValue();
        if (rejectTarget && contains(target, row, column) && !contains(source, row, column)) {
            throw ServiceException.validation("Move destination would overwrite " + label.toLowerCase(java.util.Locale.ROOT));
        }
        if (contains(source, row, column)) {
            point.put("row", row + rowDelta);
            point.put("column", column + columnDelta);
        }
    }

    private static void moveAutoFilter(ObjectNode root, JsonNode raw, String sheetId, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        if (raw == null || raw.isNull()) return;
        ObjectNode filter = requireObject(raw, "AutoFilter");
        boolean rangeMoved = moveRangeReference(root, filter.get("range"), sheetId, source, target, rowDelta, columnDelta);
        JsonNode sortRaw = filter.get("sortState");
        if (sortRaw != null && !sortRaw.isNull()) {
            ObjectNode sort = requireObject(sortRaw, "AutoFilter sort state");
            moveRangeReference(root, sort.get("ref"), sheetId, source, target, rowDelta, columnDelta);
            for (JsonNode conditionRaw : SnapshotMutationSupport.array(sort, "conditions")) {
                ObjectNode condition = requireObject(conditionRaw, "AutoFilter sort condition");
                moveRangeReference(root, condition.get("ref"), sheetId, source, target, rowDelta, columnDelta);
            }
        }
        if (columnDelta != 0 && rangeMoved) {
            ObjectNode columns = SnapshotMutationSupport.requiredObject(filter, "columns");
            ObjectNode moved = columns.objectNode();
            columns.fields().forEachRemaining(entry -> {
                int column = integerKey(entry.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Filter criteria column");
                int nextColumn = column >= source.startColumn() && column <= source.endColumn() ? column + columnDelta : column;
                String key = Integer.toString(nextColumn);
                if (moved.has(key)) throw ServiceException.validation("Move collides AutoFilter column criteria");
                ObjectNode definition = requireObject(entry.getValue(), "AutoFilter column").deepCopy();
                definition.put("column", nextColumn);
                moved.set(key, definition);
            });
            columns.removeAll();
            moved.fields().forEachRemaining(entry -> columns.set(entry.getKey(), entry.getValue()));
        }
    }

    private static void moveRules(ObjectNode root, ObjectNode owner, String property, String sheetId, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        for (JsonNode raw : SnapshotMutationSupport.array(owner, property)) {
            ObjectNode rule = requireObject(raw, "Range rule");
            for (JsonNode range : SnapshotMutationSupport.array(rule, "ranges")) {
                moveRangeReference(root, range, sheetId, source, target, rowDelta, columnDelta);
            }
            JsonNode listRaw = rule.get("listSource");
            if (listRaw != null && listRaw.isObject() && "range".equals(listRaw.path("kind").asText())) {
                moveRangeReference(root, listRaw.get("range"), sheetId, source, target, rowDelta, columnDelta);
            }
            JsonNode anchorRaw = rule.get("formulaAnchor");
            String anchorSheet = anchorRaw != null && anchorRaw.isObject()
                    ? anchorRaw.path("sheetId").asText(rule.path("sheetId").asText(owner.path("id").asText()))
                    : rule.path("sheetId").asText(owner.path("id").asText());
            if (sheetId.equals(anchorSheet) && anchorRaw != null && anchorRaw.isObject()) {
                movePoint((ObjectNode) anchorRaw, source, target, rowDelta, columnDelta, false, "Rule formula anchor");
            }
        }
    }

    private static void moveDrawings(ObjectNode sheet, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "drawings")) {
            ObjectNode drawing = requireObject(raw, "Drawing");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
            if ("absolute".equals(anchor.path("kind").asText())) continue;
            int row = anchor.path("row").asInt(-1);
            int column = anchor.path("column").asInt(-1);
            int endRow = anchor.path("endRow").asInt(row);
            int endColumn = anchor.path("endColumn").asInt(column);
            RangeRef extent = new RangeRef(source.sheetId(), row, endRow, column, endColumn);
            if (intersects(source, extent) && !containsRange(source, normalize(extent))) throw ServiceException.validation("Move partially intersects a drawing anchor");
            if (intersects(target, extent) && !containsRange(source, normalize(extent))) throw ServiceException.validation("Move destination would overwrite a drawing anchor");
            if (containsRange(source, normalize(extent))) {
                anchor.put("row", row + rowDelta);
                anchor.put("column", column + columnDelta);
                if (anchor.has("endRow")) anchor.put("endRow", endRow + rowDelta);
                if (anchor.has("endColumn")) anchor.put("endColumn", endColumn + columnDelta);
            }
        }
    }

    private static void moveHyperlinkAnchors(ObjectNode sheet, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "hyperlinks")) {
            ObjectNode entry = requireObject(raw, "Hyperlink entry");
            if (!entry.path("row").isIntegralNumber() || !entry.path("column").isIntegralNumber()) throw ServiceException.validation("Hyperlink anchor is invalid");
            movePoint(entry, source, target, rowDelta, columnDelta, true, "Hyperlink anchor");
        }
    }

    private static void moveDrawingPayloads(ObjectNode root, ObjectNode owner, String sheetId, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        ObjectNode payloads = SnapshotMutationSupport.object(owner, "drawingPayloads");
        payloads.fields().forEachRemaining(entry -> {
            ObjectNode payload = requireObject(entry.getValue(), "Drawing payload");
            String kind = payload.path("kind").asText();
            if ("camera".equals(kind) || "screenshot".equals(kind)) {
                moveRangeReference(root, payload.get("sourceRange"), sheetId, source, target, rowDelta, columnDelta);
                return;
            }
            if ("form-control".equals(kind)) {
                JsonNode linkRaw = payload.get("cellLink");
                if (linkRaw != null && linkRaw.isObject() && sheetId.equals(linkRaw.path("sheetId").asText())) {
                    movePoint((ObjectNode) linkRaw, source, target, rowDelta, columnDelta, true, "Form-control cell link");
                }
                if (payload.has("inputRange")) moveRangeReference(root, payload.get("inputRange"), sheetId, source, target, rowDelta, columnDelta);
                return;
            }
            if (!"chart".equals(kind)) return;
            ObjectNode chartSource = requireObject(payload.get("source"), "Chart source");
            String sourceKind = chartSource.path("kind").asText();
            if ("worksheet-ranges".equals(sourceKind)) {
                for (JsonNode range : SnapshotMutationSupport.array(chartSource, "ranges")) moveRangeReference(root, range, sheetId, source, target, rowDelta, columnDelta);
            } else if ("report-range".equals(sourceKind)) {
                moveRangeReference(root, chartSource.get("range"), sheetId, source, target, rowDelta, columnDelta);
            }
            if (payload.has("categoryRange")) moveRangeReference(root, payload.get("categoryRange"), sheetId, source, target, rowDelta, columnDelta);
            for (JsonNode seriesRaw : SnapshotMutationSupport.array(payload, "series")) {
                ObjectNode series = requireObject(seriesRaw, "Chart series");
                for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) {
                    if (series.has(field)) moveRangeReference(root, series.get(field), sheetId, source, target, rowDelta, columnDelta);
                }
                JsonNode rolesRaw = series.get("stockRoles");
                if (rolesRaw != null && rolesRaw.isObject()) for (String field : List.of("open", "high", "low", "close", "volume")) {
                    if (rolesRaw.has(field)) moveRangeReference(root, rolesRaw.get(field), sheetId, source, target, rowDelta, columnDelta);
                }
                JsonNode labelsRaw = series.get("dataLabels");
                if (labelsRaw != null && labelsRaw.isObject() && labelsRaw.has("valuesFromCells")) {
                    moveRangeReference(root, labelsRaw.get("valuesFromCells"), sheetId, source, target, rowDelta, columnDelta);
                }
                JsonNode errorRaw = series.get("errorBars");
                if (errorRaw != null && errorRaw.isObject()) for (String field : List.of("plusRange", "minusRange")) {
                    if (errorRaw.has(field)) moveRangeReference(root, errorRaw.get(field), sheetId, source, target, rowDelta, columnDelta);
                }
            }
        });
    }

    private static void moveDefinedNameAnchors(ObjectNode root, RangeRef source, RangeRef target, int rowDelta, int columnDelta) {
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            JsonNode anchorRaw = name.get("anchor");
            String ownerSheet = anchorRaw != null && anchorRaw.isObject()
                    ? anchorRaw.path("sheetId").asText(name.path("sheetId").asText(source.sheetId()))
                    : name.path("sheetId").asText(source.sheetId());
            if (source.sheetId().equals(ownerSheet) && anchorRaw != null && anchorRaw.isObject()) {
                movePoint((ObjectNode) anchorRaw, source, target, rowDelta, columnDelta, false, "Defined-name anchor");
            }
        }
    }

    private static StructuralPatch rewriteMovedFormulas(
            ObjectNode root,
            ObjectNode targetSheet,
            RangeRef source,
            RangeRef destination,
            int rowDelta,
            int columnDelta,
            List<RuleFormulaSnapshot> ruleFormulaSnapshots
    ) {
        FormulaReferenceTransformer.SheetIdentity targetIdentity = identity(targetSheet);
        Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> definedNameOwners = definedNameFormulaOwners(root, targetIdentity);
        FormulaReferenceTransformer.Range selected = formulaRange(source);
        FormulaReferenceTransformer.Range inverseSelection = formulaRange(destination);
        List<FormulaReferenceTransformer.SheetIdentity> sheetOrder = new ArrayList<>();
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            sheetOrder.add(identity(requireObject(raw, "Sheet")));
        }
        List<StructuralPatch.FormulaOwnerDelta> movedFormulaDeltas = new ArrayList<>();
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity ownerIdentity = identity(owner);
            List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas = new ArrayList<>();
            rewriteCellFormulaOwners(owner,
                    formula -> requireReversibleStructuralFormula(
                            formula,
                            value -> FormulaReferenceTransformer.remapMovedRegion(value, ownerIdentity, targetIdentity, selected, rowDelta, columnDelta, sheetOrder),
                            value -> FormulaReferenceTransformer.remapMovedRegion(value, ownerIdentity, targetIdentity, inverseSelection, -rowDelta, -columnDelta, sheetOrder),
                            "formula-cell owner on " + ownerIdentity.id()),
                    "range.move",
                    formulaOwnerDeltas,
                    entry -> movedFormulaOwnerBeforeAddress(
                            ownerIdentity.id(), entry, targetIdentity.id(), source, destination, rowDelta, columnDelta));
            movedFormulaDeltas.addAll(formulaOwnerDeltas);
            for (String property : List.of("conditionalFormats", "dataValidations")) {
                for (JsonNode rawRule : SnapshotMutationSupport.array(owner, property)) {
                    ObjectNode rule = requireObject(rawRule, "Range rule");
                    String formulaOwnerId = rule.path("formulaAnchor").path("sheetId").asText(rule.path("sheetId").asText(ownerIdentity.id()));
                    FormulaReferenceTransformer.SheetIdentity formulaOwner = formulaOwnerId.equals(ownerIdentity.id())
                            ? ownerIdentity : identity(SnapshotMutationSupport.sheet(root, formulaOwnerId));
                    rewriteRuleFormulas(rule, formula -> requireReversibleStructuralFormula(
                            formula,
                            value -> FormulaReferenceTransformer.remapMovedRegion(value, formulaOwner, targetIdentity, selected, rowDelta, columnDelta, sheetOrder),
                            value -> FormulaReferenceTransformer.remapMovedRegion(value, formulaOwner, targetIdentity, inverseSelection, -rowDelta, -columnDelta, sheetOrder),
                            "range rule " + ownerIdentity.id() + ":" + rule.path("id").asText()));
                }
            }
            for (JsonNode rawLink : SnapshotMutationSupport.array(owner, "hyperlinks")) {
                ObjectNode entry = requireObject(rawLink, "Hyperlink entry");
                ObjectNode link = SnapshotMutationSupport.requiredObject(entry, "hyperlink");
                JsonNode targetRaw = link.get("target");
                if (targetRaw == null || !targetRaw.isObject() || !"sheet".equals(targetRaw.path("kind").asText())
                        || !targetIdentity.id().equals(targetRaw.path("sheetId").asText())) continue;
                ObjectNode linkTarget = (ObjectNode) targetRaw;
                if (linkTarget.path("row").isIntegralNumber() && linkTarget.path("column").isIntegralNumber()) {
                    movePoint(linkTarget, source, destination, rowDelta, columnDelta, true, "Hyperlink target");
                }
                JsonNode address = linkTarget.get("address");
                if (address != null && address.isTextual()) linkTarget.put("address", requireReversibleStructuralFormula(
                        address.asText(),
                        value -> FormulaReferenceTransformer.remapMovedRegion(value, targetIdentity, targetIdentity, selected, rowDelta, columnDelta, sheetOrder),
                        value -> FormulaReferenceTransformer.remapMovedRegion(value, targetIdentity, targetIdentity, inverseSelection, -rowDelta, -columnDelta, sheetOrder),
                        "hyperlink target address"));
            }
        }
        ObjectNode names = SnapshotMutationSupport.object(root, "definedNames");
        names.fields().forEachRemaining(entry -> {
            if (entry.getValue().isTextual()) {
                FormulaReferenceTransformer.SheetIdentity owner = definedNameProjectionOwner(definedNameOwners, entry.getKey(), targetIdentity);
                names.put(entry.getKey(), requireReversibleStructuralFormula(
                        entry.getValue().asText(),
                        value -> FormulaReferenceTransformer.remapMovedRegion(value, owner, targetIdentity, selected, rowDelta, columnDelta, sheetOrder),
                        value -> FormulaReferenceTransformer.remapMovedRegion(value, owner, targetIdentity, inverseSelection, -rowDelta, -columnDelta, sheetOrder),
                        "defined-name projection " + entry.getKey()));
            }
        });
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            if (!name.path("formula").isTextual()) continue;
            FormulaReferenceTransformer.SheetIdentity owner = definedNameFormulaOwner(definedNameOwners, name);
            name.put("formula", requireReversibleStructuralFormula(
                    name.path("formula").asText(),
                    value -> FormulaReferenceTransformer.remapMovedRegion(value, owner, targetIdentity, selected, rowDelta, columnDelta, sheetOrder),
                    value -> FormulaReferenceTransformer.remapMovedRegion(value, owner, targetIdentity, inverseSelection, -rowDelta, -columnDelta, sheetOrder),
                    "defined name " + name.path("name").asText()));
        }
        movedFormulaDeltas.addAll(rewritePersistedFormulaOwners(root, targetIdentity,
                (formula, owner) -> requireReversibleStructuralFormula(
                        formula,
                        value -> FormulaReferenceTransformer.remapMovedRegion(value, owner, targetIdentity, selected, rowDelta, columnDelta, sheetOrder),
                        value -> FormulaReferenceTransformer.remapMovedRegion(value, owner, targetIdentity, inverseSelection, -rowDelta, -columnDelta, sheetOrder),
                        "persisted formula owner on " + owner.id()),
                anchor -> moveTemplateFormulaAnchor(anchor, targetIdentity.id(), source, rowDelta, columnDelta),
                true));
        appendRuleFormulaDeltas(root, ruleFormulaSnapshots, movedFormulaDeltas);
        return new StructuralPatch(StructuralPatch.VERSION, "range.move", movedFormulaDeltas, List.of(), List.of());
    }

    private static StructuralPatch.CellAddress movedFormulaOwnerBeforeAddress(
            String ownerSheetId,
            CellEntry entry,
            String targetSheetId,
            RangeRef source,
            RangeRef destination,
            int rowDelta,
            int columnDelta
    ) {
        int row = entry.row();
        int column = entry.column();
        if (ownerSheetId.equals(targetSheetId) && contains(destination, row, column)) {
            row -= rowDelta;
            column -= columnDelta;
            if (!contains(source, row, column)) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: moved formula owner has no source address");
            }
        }
        return new StructuralPatch.CellAddress(ownerSheetId, row, column);
    }

    private static void shiftCellBandAnchors(ObjectNode sheet, RangeRef selection, RangeRef band, String axis, String operation, int count) {
        SnapshotMutationSupport.remapReviewCoordinates(sheet, coordinate -> {
            if (!contains(band, coordinate.row(), coordinate.column())) return coordinate;
            int nextRow = "row".equals(axis) ? mapCellIndex(coordinate.row(), selection.startRow(), selection.endRow(), "insert".equals(operation) ? count : -count, operation) : coordinate.row();
            int nextColumn = "column".equals(axis) ? mapCellIndex(coordinate.column(), selection.startColumn(), selection.endColumn(), "insert".equals(operation) ? count : -count, operation) : coordinate.column();
            if (!contains(band, nextRow, nextColumn)) throw ServiceException.validation("Cell shift would remove review metadata");
            return new SnapshotMutationSupport.CellCoordinate(nextRow, nextColumn);
        });
    }

    static void restoreShiftedCells(ObjectNode root, String sheetId, String mutationId, JsonNode spec, JsonNode cells) {
        if (spec == null || !spec.isObject()) throw ServiceException.validation("Structural restore spec must be an object");
        ObjectNode value = (ObjectNode) spec;
        RangeRef range = SnapshotMutationSupport.range(root, value.get("range"));
        RangeRef affectedBand = SnapshotMutationSupport.range(root, value.get("affectedBand"));
        String operation = value.path("operation").asText(null);
        String axis = value.path("axis").asText(null);
        shiftCells(root, sheetId, mutationId, range, "insert".equals(operation) ? "delete" : "insert", axis, affectedBand);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        RangeRef normalized = normalize(affectedBand);
        SnapshotMutationSupport.clearCells(sheet, normalized);
        if (cells == null || !cells.isArray()) throw ServiceException.validation("Structural restore cells must be an array");
        if (cells.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Structural restore is too large");
        for (JsonNode raw : cells) {
            if (!raw.isObject()) throw ServiceException.validation("Structural restore cell must be an object");
            ObjectNode entry = (ObjectNode) raw;
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, entry);
            if (!contains(normalized, coordinate.row(), coordinate.column())) throw ServiceException.validation("Structural restore cell is outside affected band");
            JsonNode cell = entry.get("cell");
            if (cell == null || !cell.isObject()) throw ServiceException.validation("Structural restore cell payload must be an object");
            SnapshotMutationSupport.putCell(sheet, coordinate, cell);
        }
    }

    static StructuralPatch permuteRows(ObjectNode root, String sheetId, RangeRef range, int affectedColumnEnd, JsonNode sourceRows) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        SnapshotMutationSupport.requireSheet(range, sheetId);
        RangeRef selected = normalize(range);
        RangeRef metadataScope = new RangeRef(sheetId, selected.startRow(), selected.endRow(), 0, affectedColumnEnd);
        if (sourceRows == null || !sourceRows.isArray()) throw ServiceException.validation("Row permutation sourceRows must be an array");
        int expected = selected.endRow() - selected.startRow() + 1;
        if (sourceRows.size() != expected) throw ServiceException.validation("Row permutation length does not match range");
        validatePermutationPreservation(sheet, selected);
        int[] sourceRowsByTarget = validatePermutation(selected, (ArrayNode) sourceRows);
        int[] targetRowsBySource = new int[sourceRowsByTarget.length];
        for (int targetOffset = 0; targetOffset < sourceRowsByTarget.length; targetOffset++) {
            targetRowsBySource[sourceRowsByTarget[targetOffset] - selected.startRow()] = selected.startRow() + targetOffset;
        }
        rejectMovedFormulaGroups(sheet, selected, targetRowsBySource);
        validatePermutationMetadataExact(root, sheet, selected, metadataScope, targetRowsBySource);
        ObjectNode reportSheetAfter = mapReportSheetCoordinates(sheet,
                (row, column) -> contains(metadataScope, row, column)
                        ? new int[]{remapRow(row, selected, targetRowsBySource), column}
                        : new int[]{row, column},
                row -> row >= selected.startRow() && row <= selected.endRow()
                        ? remapRow(row, selected, targetRowsBySource) : row,
                "row-permutation");
        Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> rangeOwnersBefore = captureRangeOwnerSnapshots(root, sheetId);
        List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas = new ArrayList<>();
        List<RuleFormulaSnapshot> ruleFormulaSnapshots = captureRuleFormulaSnapshots(root);
        remapPermutedCells(sheet, selected, targetRowsBySource, formulaOwnerDeltas);
        remapPermutationMetadata(root, sheet, selected, metadataScope, targetRowsBySource);
        appendRuleFormulaDeltas(root, ruleFormulaSnapshots, formulaOwnerDeltas);
        applyReportSheetPlan(sheet, reportSheetAfter);
        invalidateFormulaCaches(root);
        AutoFilterOwnershipValidator.resolveOwners(sheet, sheetId);
        List<StructuralPatch.RangeOwnerDelta> rangeOwnerDeltas = rangeOwnerDeltas(
                rangeOwnersBefore, captureRangeOwnerSnapshots(root, sheetId));
        return new StructuralPatch(StructuralPatch.VERSION, "rows.permuted", formulaOwnerDeltas, List.of(), rangeOwnerDeltas);
    }

    private static void validateAxisBounds(int limit, int maximum, int at, int count, FormulaReferenceTransformer.Direction direction) {
        if (at < 0 || count < 1) throw ServiceException.validation("Structural bounds are invalid");
        if (direction == FormulaReferenceTransformer.Direction.INSERT) {
            if (at > limit || (long) limit + count > maximum) throw ServiceException.validation("Structural insert exceeds worksheet bounds");
            return;
        }
        if (at >= limit || (long) at + count > limit) throw ServiceException.validation("Structural delete is outside worksheet bounds");
    }

    private static int dimension(ObjectNode sheet, FormulaReferenceTransformer.Axis axis) {
        String property = axis == FormulaReferenceTransformer.Axis.ROW ? "rowCount" : "columnCount";
        JsonNode value = sheet.get(property);
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1;
        if (value == null || !value.isIntegralNumber() || value.intValue() < 1 || value.intValue() > maximum) throw ServiceException.validation(property + " is invalid");
        return value.intValue();
    }

    private static void validateCellShiftBounds(ObjectNode sheet, RangeRef selection, RangeRef band, String axis, String operation, int count) {
        cellsInRange(sheet, band).forEach(entry -> {
            int mapped = "row".equals(axis)
                    ? mapCellIndex(entry.row(), selection.startRow(), selection.endRow(), "insert".equals(operation) ? count : -count, operation)
                    : mapCellIndex(entry.column(), selection.startColumn(), selection.endColumn(), "insert".equals(operation) ? count : -count, operation);
            int limit = "row".equals(axis) ? dimension(sheet, FormulaReferenceTransformer.Axis.ROW) : dimension(sheet, FormulaReferenceTransformer.Axis.COLUMN);
            if (mapped >= limit || (mapped < 0 && "insert".equals(operation))) throw ServiceException.validation("Cell shift would discard data outside worksheet bounds");
        });
    }

    private static int mapCellIndex(int value, int start, int end, int delta, String operation) {
        if ("delete".equals(operation) && value >= start && value <= end) return -1;
        if (value < start) return value;
        return value + delta;
    }

    private static void setDimension(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int value) {
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1;
        if (value < 1 || value > maximum) throw ServiceException.validation("Structural result exceeds worksheet bounds");
        sheet.put(axis == FormulaReferenceTransformer.Axis.ROW ? "rowCount" : "columnCount", value);
    }

    private static void remapCells(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ObjectNode oldCells = SnapshotMutationSupport.cells(sheet);
        ObjectNode next = JsonNodeFactory.instance.objectNode();
        for (java.util.Iterator<java.util.Map.Entry<String, JsonNode>> rows = oldCells.fields(); rows.hasNext();) {
            java.util.Map.Entry<String, JsonNode> rowEntry = rows.next();
            int row = integerKey(rowEntry.getKey(), SnapshotMutationSupport.MAX_ROW, "Cell row");
            if (!rowEntry.getValue().isObject()) throw ServiceException.validation("Cell row must be an object");
            for (java.util.Iterator<java.util.Map.Entry<String, JsonNode>> columns = ((ObjectNode) rowEntry.getValue()).fields(); columns.hasNext();) {
                java.util.Map.Entry<String, JsonNode> columnEntry = columns.next();
                int column = integerKey(columnEntry.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (!columnEntry.getValue().isObject()) throw ServiceException.validation("Cell payload must be an object");
                int nextRow = axis == FormulaReferenceTransformer.Axis.ROW ? shiftIndex(row, at, count, direction, axis) : row;
                int nextColumn = axis == FormulaReferenceTransformer.Axis.COLUMN ? shiftIndex(column, at, count, direction, axis) : column;
                if (nextRow < 0 || nextColumn < 0) continue;
                if (nextRow > SnapshotMutationSupport.MAX_ROW || nextColumn > SnapshotMutationSupport.MAX_COLUMN) {
                    throw ServiceException.validation("Structural mutation moves a cell outside worksheet bounds");
                }
                ObjectNode rowTarget = next.with(Integer.toString(nextRow));
                rowTarget.set(Integer.toString(nextColumn), columnEntry.getValue().deepCopy());
            }
        }
        sheet.set("cells", next);
    }

    private static int shiftIndex(
            int value,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction,
            FormulaReferenceTransformer.Axis axis
    ) {
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW
                ? ReferenceTransformDomain.MAX_ROW_INDEX
                : ReferenceTransformDomain.MAX_COLUMN_INDEX;
        ReferenceTransformDomain.PointMapping mapped = ReferenceTransformDomain.mapPoint(
                value, at, count, direction == FormulaReferenceTransformer.Direction.INSERT, maximum);
        if (mapped.kind() == ReferenceTransformDomain.PointKind.OUT_OF_BOUNDS) {
            throw ServiceException.validation("Structural reference coordinate exceeds worksheet bounds");
        }
        return mapped.kind() == ReferenceTransformDomain.PointKind.DELETED ? -1 : Math.toIntExact(mapped.position());
    }

    private static ObjectNode mapReportSheetCoordinates(
            ObjectNode sheet,
            BiFunction<Integer, Integer, int[]> mapCell,
            Function<Integer, Integer> mapRepeatedHeaderRow,
            String operation
    ) {
        JsonNode rawDefinition = sheet.get("reportSheet");
        if (rawDefinition == null || rawDefinition.isNull()) return null;
        if (!rawDefinition.isObject()) throw ServiceException.validation("ReportSheet definition is invalid during " + operation);
        ObjectNode definition = ((ObjectNode) rawDefinition).deepCopy();
        boolean changed = false;
        ArrayNode bindings = SnapshotMutationSupport.requiredArray(definition, "bindings");
        for (int index = 0; index < bindings.size(); index++) {
            ObjectNode binding = requireObject(bindings.get(index), "ReportSheet binding");
            ObjectNode cell = SnapshotMutationSupport.requiredObject(binding, "cell");
            int row = reportCoordinate(cell, "row", SnapshotMutationSupport.MAX_ROW, index);
            int column = reportCoordinate(cell, "column", SnapshotMutationSupport.MAX_COLUMN, index);
            int[] mapped = mapCell.apply(row, column);
            if (mapped == null) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: " + operation
                        + " removes report binding " + index + " at " + row + ":" + column);
            }
            if (mapped.length != 2 || mapped[0] < 0 || mapped[0] > SnapshotMutationSupport.MAX_ROW
                    || mapped[1] < 0 || mapped[1] > SnapshotMutationSupport.MAX_COLUMN) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: " + operation
                        + " moves report binding " + index + " outside worksheet bounds");
            }
            changed |= mapped[0] != row || mapped[1] != column;
            cell.put("row", mapped[0]).put("column", mapped[1]);
        }

        ObjectNode pagination = SnapshotMutationSupport.requiredObject(definition, "pagination");
        JsonNode rawRows = pagination.get("repeatHeaderRows");
        if (rawRows != null) {
            if (!rawRows.isArray()) throw ServiceException.validation("ReportSheet repeated header rows are invalid during " + operation);
            ArrayNode mappedRows = JsonNodeFactory.instance.arrayNode();
            boolean headerRowsChanged = false;
            for (JsonNode rawRow : rawRows) {
                if (!rawRow.isIntegralNumber() || !rawRow.canConvertToInt()
                        || rawRow.intValue() < 0 || rawRow.intValue() > SnapshotMutationSupport.MAX_ROW) {
                    throw ServiceException.validation("ReportSheet repeated header row is invalid during " + operation);
                }
                Integer mapped = mapRepeatedHeaderRow == null
                        ? rawRow.intValue() : mapRepeatedHeaderRow.apply(rawRow.intValue());
                if (mapped == null) {
                    headerRowsChanged = true;
                    continue;
                }
                if (mapped < 0 || mapped > SnapshotMutationSupport.MAX_ROW) {
                    throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: " + operation
                            + " moves a repeated header outside worksheet bounds");
                }
                mappedRows.add(mapped);
                headerRowsChanged |= mapped != rawRow.intValue();
            }
            if (headerRowsChanged) {
                pagination.set("repeatHeaderRows", mappedRows);
                changed = true;
            }
        }
        return changed ? definition : null;
    }

    private static int reportCoordinate(ObjectNode cell, String property, int maximum, int bindingIndex) {
        JsonNode value = cell.get(property);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt()
                || value.intValue() < 0 || value.intValue() > maximum) {
            throw ServiceException.validation("ReportSheet binding " + bindingIndex + " " + property + " is invalid");
        }
        return value.intValue();
    }

    private static void applyReportSheetPlan(ObjectNode sheet, ObjectNode definition) {
        if (definition != null) sheet.set("reportSheet", definition);
    }

    private static int integerKey(String value, int maximum, String label) {
        try {
            int parsed = Integer.parseInt(value);
            if (parsed < 0 || parsed > maximum) throw ServiceException.validation(label + " is out of bounds");
            return parsed;
        } catch (NumberFormatException error) {
            throw ServiceException.validation(label + " is invalid");
        }
    }

    private static RangeRef normalize(RangeRef range) {
        return new RangeRef(range.sheetId(), Math.min(range.startRow(), range.endRow()), Math.max(range.startRow(), range.endRow()), Math.min(range.startColumn(), range.endColumn()), Math.max(range.startColumn(), range.endColumn()));
    }

    private static boolean contains(RangeRef range, int row, int column) {
        return row >= range.startRow() && row <= range.endRow() && column >= range.startColumn() && column <= range.endColumn();
    }

    /** A table's header/total rows are not part of a data-body sort permutation. */
    private static boolean isTableBodyPermutation(ObjectNode table, RangeRef range) {
        JsonNode tableRange = table.get("range");
        if (tableRange == null || !table.path("hasHeaderRow").asBoolean(false)) return false;
        int totalOffset = table.path("hasTotalRow").asBoolean(false) ? 1 : 0;
        return range.sheetId().equals(tableRange.path("sheetId").asText())
                && range.startRow() == tableRange.path("startRow").asInt() + 1
                && range.endRow() == tableRange.path("endRow").asInt() - totalOffset
                && range.startColumn() == tableRange.path("startColumn").asInt()
                && range.endColumn() == tableRange.path("endColumn").asInt();
    }

    private static void validateDeletePreservation(ObjectNode root, ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count) {
        int end = at + count - 1;
        for (JsonNode sparkline : SnapshotMutationSupport.array(target, "sparklines")) {
            if (insideDeleted(sparkline.path("anchor").path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a sparkline anchor");
            }
        }
        for (JsonNode pivot : SnapshotMutationSupport.array(target, "pivots")) {
            ObjectNode pivotTarget = PivotMutationDescriptor.requiredTarget(pivot);
            JsonNode anchor = pivotTarget.path("anchor");
            if (anchor.isObject() && insideDeleted(anchor.path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a pivot target anchor");
            }
        }
        for (JsonNode spill : SnapshotMutationSupport.array(target, "spillRanges")) {
            if (insideDeleted(spill.path("anchor").path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a spill anchor");
            }
        }
        for (JsonNode drawing : SnapshotMutationSupport.array(target, "drawings")) {
            JsonNode anchor = drawing.path("anchor");
            if (anchor.path("kind").asText().equals("absolute")) continue;
            String startKey = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            String endKey = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
            if (insideDeleted(anchor.path(startKey).asInt(-1), at, end) || insideDeleted(anchor.path(endKey).asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a drawing anchor");
            }
        }
        SnapshotMutationSupport.reviewMap(target, "notesByCell").fieldNames().forEachRemaining(key -> {
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.reviewCoordinate(key);
            if (insideDeleted(axis == FormulaReferenceTransformer.Axis.ROW ? coordinate.row() : coordinate.column(), at, end)) {
                throw ServiceException.validation("Structural delete would lose a note");
            }
        });
        for (JsonNode comment : SnapshotMutationSupport.reviewThreads(target)) {
            if (insideDeleted(comment.path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a comment thread");
            }
        }
        for (JsonNode table : SnapshotMutationSupport.array(target, "sheetTables")) {
            if (intersectsAxis(table.get("range"), axis, at, count)) throw ServiceException.validation("Structural delete intersects a sheet table");
        }
        String targetSheetId = target.path("id").asText();
        for (JsonNode table : workbookTables(root)) {
            JsonNode range = table.get("sourceRange");
            if (range != null && targetSheetId.equals(range.path("sheetId").asText()) && intersectsAxis(range, axis, at, count)) {
                throw ServiceException.validation("Structural delete intersects a workbook table");
            }
        }
    }

    private static boolean insideDeleted(int position, int at, int end) {
        return position >= at && position <= end;
    }

    private static boolean intersectsAxis(JsonNode range, FormulaReferenceTransformer.Axis axis, int at, int count) {
        if (range == null || !range.isObject()) throw ServiceException.validation("Structural participant range is invalid");
        String start = axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn";
        String end = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
        return range.path(start).asInt(Integer.MAX_VALUE) <= at + count - 1 && range.path(end).asInt(-1) >= at;
    }

    private static void validateSheetTableColumnPreservation(
            ObjectNode root,
            ObjectNode target,
            FormulaReferenceTransformer.Axis axis,
            int at,
            FormulaReferenceTransformer.Direction direction
    ) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "sheetTables")) {
            ObjectNode table = requireObject(raw, "Sheet table");
            RangeRef range = SnapshotMutationSupport.range(root, table.get("range"));
            if (!target.path("id").asText().equals(range.sheetId())) {
                throw ServiceException.validation("Sheet Table range must target its worksheet");
            }
            JsonNode columns = table.get("columns");
            int width = range.endColumn() - range.startColumn() + 1;
            if (columns == null || !columns.isArray() || columns.size() != width) {
                throw ServiceException.validation("Sheet Table columns must match its range width");
            }
            if (axis == FormulaReferenceTransformer.Axis.COLUMN
                    && direction == FormulaReferenceTransformer.Direction.INSERT
                    && at > range.startColumn() && at <= range.endColumn()) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: inserting a worksheet column inside Sheet Table "
                        + table.path("id").asText() + " requires a table-column structural patch");
            }
        }
    }

    private static void shiftAllMetadata(
            ObjectNode root,
            ObjectNode target,
            String targetSheetId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        shiftMerges(root, target, targetSheetId, axis, at, count, direction);
        shiftFreeze(target, axis, at, count, direction);
        shiftHiddenAndSizes(target, axis, at, count, direction);
        shiftTargetReview(target, axis, at, count, direction);
        shiftTargetDrawings(target, axis, at, count, direction);
        shiftTargetSpills(root, target, targetSheetId, axis, at, count, direction);
        shiftTargetSheetTables(root, target, targetSheetId, axis, at, count, direction);
        shiftTargetProtectionAndOutline(root, target, targetSheetId, axis, at, count, direction);
        shiftDataRegions(root, target, targetSheetId, axis, at, count, direction);
        shiftDataSources(root, targetSheetId, axis, at, count, direction);

        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            if (!raw.isObject()) throw ServiceException.validation("Workbook contains an invalid sheet");
            ObjectNode owner = (ObjectNode) raw;
            shiftRules(root, owner, targetSheetId, "conditionalFormats", axis, at, count, direction);
            shiftRules(root, owner, targetSheetId, "dataValidations", axis, at, count, direction);
            shiftFilter(root, owner, targetSheetId, axis, at, count, direction);
            shiftPivots(root, owner, targetSheetId, axis, at, count, direction);
            shiftSparklines(root, owner, targetSheetId, axis, at, count, direction);
            shiftDrawingPayloads(root, owner, targetSheetId, axis, at, count, direction);
        }
        shiftWorkbookTables(root, targetSheetId, axis, at, count, direction);
        shiftHyperlinks(root, target, axis, at, count, direction);
        shiftPrintDocuments(root, targetSheetId, axis, at, count, direction);
    }

    private static void shiftPrintDocuments(ObjectNode root, String targetSheetId,
            FormulaReferenceTransformer.Axis axis, int at, int count,
            FormulaReferenceTransformer.Direction direction) {
        JsonNode rawDocuments = root.get("printDocuments");
        if (rawDocuments == null || rawDocuments.isNull()) return;
        if (!rawDocuments.isArray()) throw ServiceException.validation("printDocuments must be an array");
        ObjectNode targetDocument = null;
        for (JsonNode raw : rawDocuments) {
            ObjectNode document = requireObject(raw, "Print document");
            if (!targetSheetId.equals(document.path("sheetId").asText())) continue;
            if (targetDocument != null) throw ServiceException.validation("A worksheet has multiple print documents");
            targetDocument = document;
        }
        if (targetDocument == null) return;

        ArrayNode printAreas = SnapshotMutationSupport.array(targetDocument, "printAreas");
        for (int index = printAreas.size() - 1; index >= 0; index--) {
            ObjectNode area = requireObject(printAreas.get(index), "Print area");
            if (!shiftRange(root, area.get("range"), targetSheetId, axis, at, count, direction)) printAreas.remove(index);
        }

        String titleProperty = axis == FormulaReferenceTransformer.Axis.ROW ? "repeatRows" : "repeatColumns";
        JsonNode titleRaw = targetDocument.get(titleProperty);
        if (titleRaw != null && !titleRaw.isNull()) {
            ObjectNode title = requireObject(titleRaw, "Print title span");
            String startKey = "start";
            String endKey = "end";
            if (!title.path(startKey).isIntegralNumber() || !title.path(endKey).isIntegralNumber()) {
                throw ServiceException.validation("Print title span coordinates are invalid");
            }
            int start = title.path(startKey).intValue();
            int end = title.path(endKey).intValue();
            int[] mapped = FormulaReferenceTransformer.remapAxisIntervalCoordinates(start, end, axis, at, count, direction);
            if (mapped == null) targetDocument.remove(titleProperty);
            else {
                title.put(startKey, mapped[0]);
                title.put(endKey, mapped[1]);
            }
        }

        ArrayNode pageBreaks = SnapshotMutationSupport.array(targetDocument, "pageBreaks");
        String breakProperty = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
        for (int index = pageBreaks.size() - 1; index >= 0; index--) {
            ObjectNode pageBreak = requireObject(pageBreaks.get(index), "Print page break");
            boolean hasRow = pageBreak.has("row") && !pageBreak.get("row").isNull();
            boolean hasColumn = pageBreak.has("column") && !pageBreak.get("column").isNull();
            if (hasRow == hasColumn) throw ServiceException.validation("Print page break must identify exactly one axis");
            if (!pageBreak.has(breakProperty) || pageBreak.get(breakProperty).isNull()) continue;
            JsonNode coordinate = pageBreak.get(breakProperty);
            if (!coordinate.isIntegralNumber() || coordinate.intValue() < 0) throw ServiceException.validation("Print page break coordinate is invalid");
            int shifted = shiftIndex(coordinate.intValue(), at, count, direction, axis);
            if (shifted < 0) pageBreaks.remove(index);
            else pageBreak.put(breakProperty, shifted);
        }
    }

    private static void shiftCellBandPrintAreas(ObjectNode root, String targetSheetId,
            FormulaReferenceTransformer.Range selection, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        JsonNode rawDocuments = root.get("printDocuments");
        if (rawDocuments == null || rawDocuments.isNull()) return;
        if (!rawDocuments.isArray()) throw ServiceException.validation("printDocuments must be an array");
        ObjectNode targetDocument = null;
        for (JsonNode raw : rawDocuments) {
            ObjectNode document = requireObject(raw, "Print document");
            if (!targetSheetId.equals(document.path("sheetId").asText())) continue;
            if (targetDocument != null) throw ServiceException.validation("A worksheet has multiple print documents");
            targetDocument = document;
        }
        if (targetDocument == null) return;
        ArrayNode printAreas = SnapshotMutationSupport.array(targetDocument, "printAreas");
        for (int index = printAreas.size() - 1; index >= 0; index--) {
            ObjectNode area = requireObject(printAreas.get(index), "Print area");
            if (!mapCellShiftRange(root, area.get("range"), targetSheetId, selection, axis, direction, "print area")) {
                printAreas.remove(index);
            }
        }
    }

    private static boolean shiftRange(ObjectNode root, JsonNode raw, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Structural participant range is invalid");
        ObjectNode range = (ObjectNode) raw;
        RangeRef current = SnapshotMutationSupport.range(root, range);
        if (!targetSheetId.equals(current.sheetId())) return true;
        FormulaReferenceTransformer.Range mapped = FormulaReferenceTransformer.remapAxisRangeCoordinates(
                formulaRange(current), axis, at, count, direction);
        if (mapped == null) return false;
        range.put("startRow", mapped.startRow());
        range.put("endRow", mapped.endRow());
        range.put("startColumn", mapped.startColumn());
        range.put("endColumn", mapped.endColumn());
        return true;
    }

    private static void shiftMerges(ObjectNode root, ObjectNode sheet, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode merges = SnapshotMutationSupport.array(sheet, "merges");
        for (int index = merges.size() - 1; index >= 0; index--) {
            ObjectNode merge = requireObject(merges.get(index), "Merge");
            boolean keep = shiftRange(root, merge.get("range"), targetSheetId, axis, at, count, direction);
            if (!keep) {
                merges.remove(index);
                continue;
            }
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(merge, "anchor");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int current = anchor.path(key).asInt(-1);
            if (direction == FormulaReferenceTransformer.Direction.INSERT && current >= at) anchor.put(key, current + count);
            if (direction == FormulaReferenceTransformer.Direction.DELETE) {
                if (current >= at && current < at + count) anchor.put(key, merge.path("range").path(axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn").asInt());
                else if (current >= at + count) anchor.put(key, current - count);
            }
        }
    }

    private static void shiftRules(ObjectNode root, ObjectNode owner, String targetSheetId, String property, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(owner, property)) {
            ObjectNode rule = requireObject(raw, "Range rule");
            ArrayNode ranges = SnapshotMutationSupport.array(rule, "ranges");
            ArrayNode next = JsonNodeFactory.instance.arrayNode();
            for (JsonNode range : ranges) if (shiftRange(root, range, targetSheetId, axis, at, count, direction)) next.add(range);
            if (next.isEmpty()) throw ServiceException.validation("Structural mutation removes every range owned by rule " + rule.path("id").asText());
            rule.set("ranges", next);
            JsonNode listSourceRaw = rule.get("listSource");
            if (listSourceRaw != null && listSourceRaw.isObject() && "range".equals(listSourceRaw.path("kind").asText())) {
                ObjectNode listSource = (ObjectNode) listSourceRaw;
                if (!shiftRange(root, listSource.get("range"), targetSheetId, axis, at, count, direction)) {
                    throw ServiceException.validation("Structural mutation removes a data-validation list source");
                }
            }
            JsonNode anchorRaw = rule.get("formulaAnchor");
            if (anchorRaw != null && anchorRaw.isObject() && targetSheetId.equals(anchorRaw.path("sheetId").asText())) {
                ObjectNode anchor = (ObjectNode) anchorRaw;
                String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
                if (!anchor.path(key).isIntegralNumber()) throw ServiceException.validation("Rule formula anchor is invalid");
                int shifted = shiftIndex(anchor.path(key).intValue(), at, count, direction, axis);
                if (shifted < 0) throw ServiceException.validation("Structural mutation removes a rule formula anchor");
                anchor.put(key, shifted);
            }
        }
    }

    private static void shiftFilter(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        JsonNode filterRaw = owner.get("autoFilter");
        if (filterRaw == null || filterRaw.isNull()) return;
        ObjectNode filter = requireObject(filterRaw, "Filter");
        if (!shiftRange(root, filter.get("range"), targetSheetId, axis, at, count, direction)) throw ServiceException.validation("Structural mutation removes an AutoFilter range");
        shiftAutoFilterSortState(root, filter, targetSheetId, axis, at, count, direction);
        if (axis != FormulaReferenceTransformer.Axis.COLUMN || !targetSheetId.equals(owner.path("id").asText())) return;
        ObjectNode criteria = SnapshotMutationSupport.requiredObject(filter, "columns");
        shiftFilterColumns(criteria, axis, at, count, direction);
    }

    private static void shiftFilterColumns(ObjectNode criteria, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        if (axis != FormulaReferenceTransformer.Axis.COLUMN) return;
        ObjectNode next = JsonNodeFactory.instance.objectNode();
        criteria.fields().forEachRemaining(entry -> {
            int column = integerKey(entry.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Filter criteria column");
            int shifted = shiftIndex(column, at, count, direction, axis);
            if (shifted < 0) throw ServiceException.validation("Structural mutation removes an AutoFilter column");
            ObjectNode condition = requireObject(entry.getValue(), "AutoFilter column").deepCopy();
            condition.put("column", shifted);
            next.set(Integer.toString(shifted), condition);
        });
        criteria.removeAll();
        next.fields().forEachRemaining(entry -> criteria.set(entry.getKey(), entry.getValue()));
    }

    private static void shiftAutoFilterSortState(ObjectNode root, ObjectNode filter, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        JsonNode raw = filter.get("sortState");
        if (raw == null || raw.isNull()) return;
        ObjectNode sortState = requireObject(raw, "AutoFilter sort state");
        if (!shiftRange(root, sortState.get("ref"), targetSheetId, axis, at, count, direction)) {
            throw ServiceException.validation("Structural mutation removes an AutoFilter sort reference");
        }
        for (JsonNode conditionRaw : SnapshotMutationSupport.array(sortState, "conditions")) {
            ObjectNode condition = requireObject(conditionRaw, "AutoFilter sort condition");
            if (!shiftRange(root, condition.get("ref"), targetSheetId, axis, at, count, direction)) {
                throw ServiceException.validation("Structural mutation removes an AutoFilter sort condition");
            }
        }
    }

    private static void shiftFreeze(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ObjectNode freeze = SnapshotMutationSupport.requiredObject(sheet, "pane");
        WorkbookSnapshotValidator.requireCanonicalPane(freeze);
        if ("none".equals(freeze.path("kind").asText())) return;
        String kind = freeze.path("kind").asText();
        String split = axis == FormulaReferenceTransformer.Axis.ROW ? "ySplit" : "xSplit";
        String start = axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn";
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW
                ? ReferenceTransformDomain.MAX_ROW_INDEX
                : ReferenceTransformDomain.MAX_COLUMN_INDEX;
        List<String> keys = "frozen".equals(kind) ? List.of(split, start) : List.of(start);
        for (String key : keys) {
            JsonNode coordinate = freeze.get(key);
            if (coordinate == null || !coordinate.isIntegralNumber() || !coordinate.canConvertToInt()) {
                throw ServiceException.validation("Pane coordinate " + key + " must be an integer before structural transform");
            }
            long value = coordinate.intValue();
            long shifted = value;
            if (direction == FormulaReferenceTransformer.Direction.INSERT && value >= at) shifted += count;
            if (direction == FormulaReferenceTransformer.Direction.DELETE && value > at) {
                shifted = value >= (long) at + count ? value - count : at;
            }
            long limit = key.equals(split) ? (long) maximum + 1 : maximum;
            if (shifted < 0 || shifted > limit) {
                throw ServiceException.validation("Pane coordinate " + key + " exceeds worksheet bounds after structural transform");
            }
            freeze.put(key, (int) shifted);
        }
        WorkbookSnapshotValidator.requireCanonicalPane(freeze);
    }

    private static void shiftHiddenAndSizes(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        String hidden = axis == FormulaReferenceTransformer.Axis.ROW ? "hiddenRows" : "hiddenColumns";
        String sizes = axis == FormulaReferenceTransformer.Axis.ROW ? "rowHeightsPx" : "columnWidthsPx";
        ArrayNode remapped = JsonNodeFactory.instance.arrayNode();
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, hidden)) {
            if (!raw.isIntegralNumber()) throw ServiceException.validation(hidden + " contains invalid index");
            int shifted = shiftIndex(raw.intValue(), at, count, direction, axis);
            if (shifted >= 0 && !containsNumber(remapped, shifted)) remapped.add(shifted);
        }
        sheet.set(hidden, remapped);
        ObjectNode nextSizes = JsonNodeFactory.instance.objectNode();
        SnapshotMutationSupport.object(sheet, sizes).fields().forEachRemaining(entry -> {
            int shifted = shiftIndex(integerKey(entry.getKey(), axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW : SnapshotMutationSupport.MAX_COLUMN, sizes), at, count, direction, axis);
            if (shifted >= 0) nextSizes.set(Integer.toString(shifted), entry.getValue());
        });
        sheet.set(sizes, nextSizes);
    }

    private static void shiftSparklines(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode sparklines = SnapshotMutationSupport.array(owner, "sparklines");
        boolean ownerTarget = targetSheetId.equals(owner.path("id").asText());
        for (JsonNode raw : sparklines) {
            ObjectNode sparkline = requireObject(raw, "Sparkline");
            if (!shiftRange(root, sparkline.get("sourceRange"), targetSheetId, axis, at, count, direction)) {
                throw ServiceException.validation("Structural mutation removes a sparkline source range");
            }
            if (!ownerTarget) continue;
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int shifted = shiftIndex(anchor.path(key).asInt(-1), at, count, direction, axis);
            if (shifted < 0) throw ServiceException.validation("Structural mutation removes a sparkline anchor");
            anchor.put(key, shifted);
        }
    }

    private static void shiftPivots(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        boolean ownerTarget = targetSheetId.equals(owner.path("id").asText());
        for (JsonNode raw : SnapshotMutationSupport.array(owner, "pivots")) {
            ObjectNode pivot = requireObject(raw, "Pivot");
            SnapshotMutationSupport.validateKnownKeys(pivot, Set.of("schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "presentation", "nativeMetadata"), "Pivot");
            PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, range -> {
                if (!shiftRange(root, range, targetSheetId, axis, at, count, direction)) throw ServiceException.validation("Structural mutation removes a pivot source range");
            });
            ObjectNode target = PivotMutationDescriptor.requiredTarget(pivot);
            JsonNode anchorRaw = target.get("anchor");
            if (ownerTarget && anchorRaw != null && anchorRaw.isObject()) {
                ObjectNode anchor = (ObjectNode) anchorRaw;
                String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
                int shifted = shiftIndex(anchor.path(key).asInt(-1), at, count, direction, axis);
                if (shifted < 0) throw ServiceException.validation("Structural mutation removes a pivot target anchor");
                anchor.put(key, shifted);
            }
        }
    }

    private static void shiftDrawingPayloads(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ObjectNode payloads = SnapshotMutationSupport.object(owner, "drawingPayloads");
        payloads.fields().forEachRemaining(entry -> {
            JsonNode raw = entry.getValue();
            if (!raw.isObject()) throw ServiceException.validation("Drawing payload is invalid");
            ObjectNode chart = (ObjectNode) raw;
            String kind = chart.path("kind").asText();
            if ("camera".equals(kind) || "screenshot".equals(kind)) {
                requireShiftedRange(root, chart.get("sourceRange"), targetSheetId, axis, at, count, direction, kind + " source range");
                return;
            }
            if ("form-control".equals(chart.path("kind").asText())) {
                JsonNode linkRaw = chart.get("cellLink");
                if (linkRaw != null && linkRaw.isObject() && targetSheetId.equals(linkRaw.path("sheetId").asText())) {
                    shiftHyperlinkCoordinate((ObjectNode) linkRaw, axis, at, count, direction, "form-control cell link");
                }
                if (chart.has("inputRange")) requireShiftedRange(root, chart.get("inputRange"), targetSheetId, axis, at, count, direction, "form-control input range");
                return;
            }
            if (!"chart".equals(chart.path("kind").asText())) return;
            ObjectNode source = requiredObject(chart.get("source"), "Chart source");
            String sourceKind = source.path("kind").asText();
            if ("worksheet-ranges".equals(sourceKind)) {
                JsonNode ranges = source.get("ranges");
                if (ranges == null || !ranges.isArray()) throw ServiceException.validation("Chart worksheet source ranges are invalid");
                if (ranges.isEmpty()) throw ServiceException.validation("Chart has no worksheet source ranges");
                for (JsonNode range : ranges) requireShiftedRange(root, range, targetSheetId, axis, at, count, direction, "chart worksheet source range");
            } else if ("report-range".equals(sourceKind)) {
                requireShiftedRange(root, source.get("range"), targetSheetId, axis, at, count, direction, "chart report range");
            } else if (!Set.of("pivot", "table").contains(sourceKind)) {
                throw ServiceException.validation("Chart source kind is invalid: " + sourceKind);
            }
            if (chart.has("categoryRange")) requireShiftedRange(root, chart.get("categoryRange"), targetSheetId, axis, at, count, direction, "chart category range");
            JsonNode seriesRaw = chart.get("series");
            if (seriesRaw != null && !seriesRaw.isNull() && !seriesRaw.isArray()) throw ServiceException.validation("Chart series collection is invalid");
            Iterable<JsonNode> chartSeries = seriesRaw == null || seriesRaw.isNull() ? List.of() : seriesRaw;
            for (JsonNode series : chartSeries) {
                if (!series.isObject()) throw ServiceException.validation("Chart series is invalid");
                for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) {
                    if (series.has(field)) requireShiftedRange(root, series.get(field), targetSheetId, axis, at, count, direction, "chart series " + field);
                }
                JsonNode errorBars = series.get("errorBars");
                if (errorBars != null && errorBars.isObject()) {
                    if (errorBars.has("plusRange")) requireShiftedRange(root, errorBars.get("plusRange"), targetSheetId, axis, at, count, direction, "chart error-bar plus range");
                    if (errorBars.has("minusRange")) requireShiftedRange(root, errorBars.get("minusRange"), targetSheetId, axis, at, count, direction, "chart error-bar minus range");
                }
                JsonNode stockRoles = series.get("stockRoles");
                if (stockRoles != null && stockRoles.isObject()) {
                    for (String field : List.of("open", "high", "low", "close", "volume")) if (stockRoles.has(field)) requireShiftedRange(root, stockRoles.get(field), targetSheetId, axis, at, count, direction, "chart stock-role range");
                }
                JsonNode dataLabels = series.get("dataLabels");
                if (dataLabels != null && dataLabels.isObject() && dataLabels.has("valuesFromCells")) requireShiftedRange(root, dataLabels.get("valuesFromCells"), targetSheetId, axis, at, count, direction, "chart data-label range");
            }
        });
    }

    /** Block-backed owners can extend beyond the materialized worksheet extent. */
    private static void validateCellShiftDataOwners(ObjectNode root, ObjectNode target, RangeRef referenceBand) {
        String targetSheetId = target.path("id").asText();
        for (JsonNode raw : SnapshotMutationSupport.array(target, "dataRegions")) {
            ObjectNode region = requireObject(raw, "Data region");
            if (intersects(SnapshotMutationSupport.range(root, region.get("range")), referenceBand)) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift intersects data region " + region.path("id").asText());
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.array(target, "sheetTables")) {
            ObjectNode table = requireObject(raw, "Sheet table");
            if (intersects(SnapshotMutationSupport.range(root, table.get("range")), referenceBand)) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift intersects sheet table " + table.path("id").asText() + "; use an explicit table operation");
            }
        }
        for (JsonNode raw : workbookTables(root)) {
            ObjectNode table = requireObject(raw, "Workbook table");
            JsonNode sourceRange = table.get("sourceRange");
            if (sourceRange != null && !sourceRange.isNull() && targetSheetId.equals(sourceRange.path("sheetId").asText())
                    && intersects(SnapshotMutationSupport.range(root, sourceRange), referenceBand)) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift intersects workbook table " + table.path("id").asText() + "; use an explicit table operation");
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.dataModelArray(root, "sources")) {
            ObjectNode source = requireObject(raw, "Data source");
            JsonNode sourceRange = source.get("sourceRange");
            if (sourceRange != null && !sourceRange.isNull() && targetSheetId.equals(sourceRange.path("sheetId").asText())
                    && intersects(SnapshotMutationSupport.range(root, sourceRange), referenceBand)) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift intersects data source " + source.path("id").asText() + "; use a data-block transaction");
            }
        }
    }

    private static void shiftCellBandMetadata(ObjectNode root, ObjectNode target, RangeRef selection, RangeRef band, String axisName, String operation, int count) {
        FormulaReferenceTransformer.Axis axis = "row".equals(axisName)
                ? FormulaReferenceTransformer.Axis.ROW
                : FormulaReferenceTransformer.Axis.COLUMN;
        FormulaReferenceTransformer.Direction direction = "insert".equals(operation)
                ? FormulaReferenceTransformer.Direction.INSERT
                : FormulaReferenceTransformer.Direction.DELETE;
        FormulaReferenceTransformer.Range selected = formulaRange(selection);
        String targetSheetId = target.path("id").asText();

        shiftCellBandAnchors(target, selection, band, axisName, operation, count);
        for (JsonNode raw : SnapshotMutationSupport.array(target, "merges")) {
            ObjectNode merge = requireObject(raw, "Merge");
            requireCellShiftRange(root, merge.get("range"), targetSheetId, selected, axis, direction, "merged range");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(merge, "anchor");
            if (!mapCellShiftAnchor(anchor, band, selected, axis, direction)) throw ServiceException.validation("Cell shift removes a merge anchor");
        }

        List<FormulaReferenceTransformer.SheetIdentity> sheetOrder = worksheetOrder(root);
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            shiftCellBandRules(root, owner, "conditionalFormats", targetSheetId, selected, band, axis, direction);
            shiftCellBandRules(root, owner, "dataValidations", targetSheetId, selected, band, axis, direction);
            shiftCellBandDrawingPayloads(root, owner, targetSheetId, selected, band, axis, direction);
            shiftCellBandPivots(root, owner, targetSheetId, selected, band, axis, direction);
            shiftCellBandSparklines(root, owner, targetSheetId, selected, band, axis, direction);
            shiftCellBandHyperlinks(root, owner, target, selected, band, axis, direction, sheetOrder);
        }

        JsonNode filterRaw = target.get("autoFilter");
        if (filterRaw != null && !filterRaw.isNull()) shiftCellBandFilter(root, requireObject(filterRaw, "AutoFilter"), targetSheetId, selected, band, axis, direction);
        for (JsonNode raw : SnapshotMutationSupport.array(target, "sheetTables")) {
            ObjectNode table = requireObject(raw, "Sheet table");
            JsonNode filter = table.get("autoFilter");
            if (filter != null && !filter.isNull()) shiftCellBandFilter(root, requireObject(filter, "Table AutoFilter"), targetSheetId, selected, band, axis, direction);
        }

        shiftCellBandSpills(root, target, targetSheetId, selected, band, axis, direction);
        shiftCellBandProtection(root, target, targetSheetId, selected, axis, direction);
        shiftCellBandDrawings(target, band, selected, axis, direction);
        shiftCellBandPrintAreas(root, targetSheetId, selected, axis, direction);
    }

    private static FormulaReferenceTransformer.Range formulaRange(RangeRef range) {
        return new FormulaReferenceTransformer.Range(range.startRow(), range.endRow(), range.startColumn(), range.endColumn());
    }

    private static boolean intersects(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId())
                && left.startRow() <= right.endRow() && left.endRow() >= right.startRow()
                && left.startColumn() <= right.endColumn() && left.endColumn() >= right.startColumn();
    }

    private static boolean mapCellShiftRange(ObjectNode root, JsonNode raw, String targetSheetId,
            FormulaReferenceTransformer.Range selection, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction, String label) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Structural participant range is invalid: " + label);
        RangeRef current = SnapshotMutationSupport.range(root, raw);
        if (!targetSheetId.equals(current.sheetId())) return true;
        FormulaReferenceTransformer.Range mapped = FormulaReferenceTransformer.remapCellShiftRangeCoordinates(
                formulaRange(current), selection, axis, direction);
        if (mapped == null) return false;
        ObjectNode value = (ObjectNode) raw;
        value.put("startRow", mapped.startRow());
        value.put("endRow", mapped.endRow());
        value.put("startColumn", mapped.startColumn());
        value.put("endColumn", mapped.endColumn());
        return true;
    }

    private static void requireCellShiftRange(ObjectNode root, JsonNode raw, String targetSheetId,
            FormulaReferenceTransformer.Range selection, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction, String label) {
        if (!mapCellShiftRange(root, raw, targetSheetId, selection, axis, direction, label)) {
            throw ServiceException.validation("Cell shift removes " + label);
        }
    }

    private static boolean mapCellShiftAnchor(ObjectNode anchor, RangeRef band, FormulaReferenceTransformer.Range selection,
            FormulaReferenceTransformer.Axis axis, FormulaReferenceTransformer.Direction direction) {
        int row = anchor.path("row").asInt(-1);
        int column = anchor.path("column").asInt(-1);
        if (row < 0 || column < 0) throw ServiceException.validation("Structural anchor coordinates are invalid");
        if (!contains(band, row, column)) return true;
        int[] mapped = FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, selection, axis, direction);
        if (mapped == null) return false;
        anchor.put("row", mapped[0]);
        anchor.put("column", mapped[1]);
        return true;
    }

    private static void shiftCellBandRules(ObjectNode root, ObjectNode owner, String property, String targetSheetId,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(owner, property)) {
            ObjectNode rule = requireObject(raw, "Range rule");
            ArrayNode ranges = SnapshotMutationSupport.array(rule, "ranges");
            ArrayNode mappedRanges = JsonNodeFactory.instance.arrayNode();
            for (JsonNode range : ranges) {
                if (mapCellShiftRange(root, range, targetSheetId, selection, axis, direction, property + " range")) mappedRanges.add(range);
            }
            if (mappedRanges.isEmpty()) throw ServiceException.validation("Cell shift removes every range owned by rule " + rule.path("id").asText());
            rule.set("ranges", mappedRanges);
            JsonNode listSourceRaw = rule.get("listSource");
            if (listSourceRaw != null && listSourceRaw.isObject() && "range".equals(listSourceRaw.path("kind").asText())) {
                ObjectNode listSource = (ObjectNode) listSourceRaw;
                requireCellShiftRange(root, listSource.get("range"), targetSheetId, selection, axis, direction, "data-validation list source");
            }
            JsonNode anchorRaw = rule.get("formulaAnchor");
            if (anchorRaw != null && anchorRaw.isObject() && targetSheetId.equals(anchorRaw.path("sheetId").asText())) {
                ObjectNode anchor = (ObjectNode) anchorRaw;
                if (contains(band, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))
                        && !mapCellShiftAnchor(anchor, band, selection, axis, direction)) {
                    throw ServiceException.validation("Cell shift removes a formula anchor owned by rule " + rule.path("id").asText());
                }
            }
        }
    }

    private static void shiftCellBandFilter(ObjectNode root, ObjectNode filter, String targetSheetId,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        JsonNode rangeRaw = filter.get("range");
        RangeRef current = SnapshotMutationSupport.range(root, rangeRaw);
        if (!targetSheetId.equals(current.sheetId())) return;
        if (axis == FormulaReferenceTransformer.Axis.COLUMN && intersects(current, band)) {
            throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift intersects an AutoFilter column owner");
        }
        requireCellShiftRange(root, rangeRaw, targetSheetId, selection, axis, direction, "AutoFilter range");
        JsonNode sortRaw = filter.get("sortState");
        if (sortRaw == null || sortRaw.isNull()) return;
        ObjectNode sortState = requireObject(sortRaw, "AutoFilter sort state");
        requireCellShiftRange(root, sortState.get("ref"), targetSheetId, selection, axis, direction, "AutoFilter sort reference");
        for (JsonNode raw : SnapshotMutationSupport.array(sortState, "conditions")) {
            ObjectNode condition = requireObject(raw, "AutoFilter sort condition");
            requireCellShiftRange(root, condition.get("ref"), targetSheetId, selection, axis, direction, "AutoFilter sort condition");
        }
    }

    private static void shiftCellBandDrawingPayloads(ObjectNode root, ObjectNode owner, String targetSheetId,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        ObjectNode payloads = SnapshotMutationSupport.object(owner, "drawingPayloads");
        payloads.fields().forEachRemaining(entry -> {
            ObjectNode payload = requireObject(entry.getValue(), "Drawing payload");
            String kind = payload.path("kind").asText();
            if ("camera".equals(kind) || "screenshot".equals(kind)) {
                requireCellShiftRange(root, payload.get("sourceRange"), targetSheetId, selection, axis, direction, kind + " source range");
                return;
            }
            if ("chart".equals(kind)) {
                ObjectNode source = requiredObject(payload.get("source"), "Chart source");
                String sourceKind = source.path("kind").asText();
                if ("worksheet-ranges".equals(sourceKind)) {
                    JsonNode rangesRaw = source.get("ranges");
                    if (rangesRaw == null || !rangesRaw.isArray() || rangesRaw.isEmpty()) throw ServiceException.validation("Chart worksheet source ranges are invalid");
                    for (JsonNode range : rangesRaw) requireCellShiftRange(root, range, targetSheetId, selection, axis, direction, "chart source range");
                } else if ("report-range".equals(sourceKind)) {
                    requireCellShiftRange(root, source.get("range"), targetSheetId, selection, axis, direction, "chart report range");
                } else if (!Set.of("pivot", "table").contains(sourceKind)) {
                    throw ServiceException.validation("Chart source kind is invalid: " + sourceKind);
                }
                if (payload.has("categoryRange")) requireCellShiftRange(root, payload.get("categoryRange"), targetSheetId, selection, axis, direction, "chart category range");
                JsonNode seriesRaw = payload.get("series");
                if (seriesRaw != null && !seriesRaw.isNull() && !seriesRaw.isArray()) throw ServiceException.validation("Chart series collection is invalid");
                Iterable<JsonNode> chartSeries = seriesRaw == null || seriesRaw.isNull() ? List.of() : seriesRaw;
                for (JsonNode rawSeries : chartSeries) {
                    ObjectNode series = requireObject(rawSeries, "Chart series");
                    for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) {
                        if (series.has(field)) requireCellShiftRange(root, series.get(field), targetSheetId, selection, axis, direction, "chart series " + field);
                    }
                    JsonNode errorBars = series.get("errorBars");
                    if (errorBars != null && errorBars.isObject()) {
                        for (String field : List.of("plusRange", "minusRange")) if (errorBars.has(field)) requireCellShiftRange(root, errorBars.get(field), targetSheetId, selection, axis, direction, "chart error-bar range");
                    }
                    JsonNode stockRoles = series.get("stockRoles");
                    if (stockRoles != null && stockRoles.isObject()) {
                        for (String field : List.of("open", "high", "low", "close", "volume")) if (stockRoles.has(field)) requireCellShiftRange(root, stockRoles.get(field), targetSheetId, selection, axis, direction, "chart stock-role range");
                    }
                    JsonNode dataLabels = series.get("dataLabels");
                    if (dataLabels != null && dataLabels.isObject() && dataLabels.has("valuesFromCells")) {
                        requireCellShiftRange(root, dataLabels.get("valuesFromCells"), targetSheetId, selection, axis, direction, "chart data-label range");
                    }
                }
                return;
            }
            if ("form-control".equals(kind)) {
                JsonNode cellLinkRaw = payload.get("cellLink");
                if (cellLinkRaw != null && cellLinkRaw.isObject() && targetSheetId.equals(cellLinkRaw.path("sheetId").asText())) {
                    if (contains(band, cellLinkRaw.path("row").asInt(-1), cellLinkRaw.path("column").asInt(-1))
                            && !mapCellShiftAnchor((ObjectNode) cellLinkRaw, band, selection, axis, direction)) {
                        throw ServiceException.validation("Cell shift removes a form-control cell link");
                    }
                }
                if (payload.has("inputRange")) requireCellShiftRange(root, payload.get("inputRange"), targetSheetId, selection, axis, direction, "form-control input range");
            }
        });
    }

    private static void shiftCellBandPivots(ObjectNode root, ObjectNode owner, String targetSheetId,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(owner, "pivots")) {
            ObjectNode pivot = requireObject(raw, "Pivot");
            PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, range -> requireCellShiftRange(
                    root, range, targetSheetId, selection, axis, direction, "pivot source range"));
            ObjectNode target = PivotMutationDescriptor.requiredTarget(pivot);
            if (targetSheetId.equals(target.path("sheetId").asText())) {
                JsonNode anchorRaw = target.get("anchor");
                if (anchorRaw != null && anchorRaw.isObject() && contains(band, anchorRaw.path("row").asInt(-1), anchorRaw.path("column").asInt(-1))
                        && !mapCellShiftAnchor((ObjectNode) anchorRaw, band, selection, axis, direction)) {
                    throw ServiceException.validation("Cell shift removes a pivot target anchor");
                }
            }
        }
    }

    private static void shiftCellBandSparklines(ObjectNode root, ObjectNode owner, String targetSheetId,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(owner, "sparklines")) {
            ObjectNode sparkline = requireObject(raw, "Sparkline");
            requireCellShiftRange(root, sparkline.get("sourceRange"), targetSheetId, selection, axis, direction, "sparkline source range");
            if (targetSheetId.equals(sparkline.path("sheetId").asText())) {
                ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
                if (contains(band, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))
                        && !mapCellShiftAnchor(anchor, band, selection, axis, direction)) {
                    throw ServiceException.validation("Cell shift removes a sparkline anchor");
                }
            }
        }
    }

    private static void shiftCellBandSpills(ObjectNode root, ObjectNode target, String targetSheetId,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "spillRanges")) {
            ObjectNode spill = requireObject(raw, "Spill range");
            requireCellShiftRange(root, spill.get("range"), targetSheetId, selection, axis, direction, "spill range");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(spill, "anchor");
            if (contains(band, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))
                    && !mapCellShiftAnchor(anchor, band, selection, axis, direction)) {
                throw ServiceException.validation("Cell shift removes a spill anchor");
            }
        }
    }

    private static void shiftCellBandProtection(ObjectNode root, ObjectNode target, String targetSheetId,
            FormulaReferenceTransformer.Range selection, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "protectionRules")) {
            ObjectNode rule = requireObject(raw, "Protection rule");
            if (rule.has("range")) requireCellShiftRange(root, rule.get("range"), targetSheetId, selection, axis, direction, "protection range");
        }
        JsonNode banded = target.get("bandedRule");
        if (banded != null && banded.isObject()) requireCellShiftRange(root, banded.get("range"), targetSheetId, selection, axis, direction, "banded range");
    }

    private static void shiftCellBandDrawings(ObjectNode target, RangeRef band,
            FormulaReferenceTransformer.Range selection, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "drawings")) {
            ObjectNode drawing = requireObject(raw, "Drawing");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
            if ("absolute".equals(anchor.path("kind").asText()) || !anchor.path("row").isIntegralNumber() || !anchor.path("column").isIntegralNumber()) continue;
            int row = anchor.path("row").intValue();
            int column = anchor.path("column").intValue();
            if (contains(band, row, column) && !mapCellShiftAnchor(anchor, band, selection, axis, direction)) {
                throw ServiceException.validation("Cell shift removes drawing anchor " + drawing.path("id").asText());
            }
            if (anchor.has("endRow") || anchor.has("endColumn")) {
                int endRow = anchor.path("endRow").asInt(row);
                int endColumn = anchor.path("endColumn").asInt(column);
                if (contains(band, endRow, endColumn)) {
                    ObjectNode extent = JsonNodeFactory.instance.objectNode();
                    extent.put("row", endRow);
                    extent.put("column", endColumn);
                    if (!mapCellShiftAnchor(extent, band, selection, axis, direction)) throw ServiceException.validation("Cell shift removes drawing extent " + drawing.path("id").asText());
                    if (anchor.has("endRow")) anchor.put("endRow", extent.path("row").asInt());
                    if (anchor.has("endColumn")) anchor.put("endColumn", extent.path("column").asInt());
                }
            }
        }
    }

    private static void shiftCellBandHyperlinks(ObjectNode root, ObjectNode owner, ObjectNode targetSheet,
            FormulaReferenceTransformer.Range selection, RangeRef band, FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction,
            List<FormulaReferenceTransformer.SheetIdentity> sheetOrder) {
        boolean ownerIsTarget = owner.path("id").asText().equals(targetSheet.path("id").asText());
        FormulaReferenceTransformer.SheetIdentity target = identity(targetSheet);
        for (JsonNode raw : SnapshotMutationSupport.array(owner, "hyperlinks")) {
            ObjectNode entry = requireObject(raw, "Hyperlink entry");
            if (ownerIsTarget && entry.path("row").isIntegralNumber() && entry.path("column").isIntegralNumber()) {
                int row = entry.path("row").intValue();
                int column = entry.path("column").intValue();
                if (contains(band, row, column)) {
                    int[] mapped = FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, selection, axis, direction);
                    if (mapped == null) throw ServiceException.validation("Cell shift removes anchored hyperlink metadata");
                    entry.put("row", mapped[0]);
                    entry.put("column", mapped[1]);
                }
            }
            ObjectNode hyperlink = SnapshotMutationSupport.requiredObject(entry, "hyperlink");
            JsonNode targetRaw = hyperlink.get("target");
            if (targetRaw == null || !targetRaw.isObject() || !"sheet".equals(targetRaw.path("kind").asText())
                    || !target.id().equals(targetRaw.path("sheetId").asText())) continue;
            ObjectNode linkTarget = (ObjectNode) targetRaw;
            if (linkTarget.path("row").isIntegralNumber() && linkTarget.path("column").isIntegralNumber()) {
                int row = linkTarget.path("row").intValue();
                int column = linkTarget.path("column").intValue();
                if (contains(band, row, column)) {
                    int[] mapped = FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, selection, axis, direction);
                    if (mapped == null) throw ServiceException.validation("Cell shift removes a hyperlink target");
                    linkTarget.put("row", mapped[0]);
                    linkTarget.put("column", mapped[1]);
                }
            }
            JsonNode address = linkTarget.get("address");
            if (address != null && address.isTextual()) {
                linkTarget.put("address", FormulaReferenceTransformer.remapCellShift(address.asText(), target, target, selection, axis, direction, sheetOrder));
            }
        }
    }

    private static void requireShiftedRange(ObjectNode root, JsonNode range, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction, String label) {
        if (!shiftRange(root, range, targetSheetId, axis, at, count, direction)) throw ServiceException.validation("Structural mutation removes " + label);
    }

    private static ObjectNode requiredObject(JsonNode raw, String label) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) raw;
    }

    private static void shiftTargetDrawings(ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "drawings")) {
            ObjectNode drawing = requireObject(raw, "Drawing");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
            if ("absolute".equals(anchor.path("kind").asText())) continue;
            shiftAnchor(axis, at, count, direction, anchor, axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column");
            String end = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
            if (anchor.has(end)) shiftAnchor(axis, at, count, direction, anchor, end);
        }
    }

    private static void shiftAnchor(FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction, ObjectNode anchor, String property) {
        if (!anchor.has(property) || anchor.get(property).isNull()) return;
        if (!anchor.path(property).isIntegralNumber()) throw ServiceException.validation("Drawing anchor coordinate is invalid");
        int shifted = shiftIndex(anchor.path(property).asInt(-1), at, count, direction, axis);
        if (shifted < 0) throw ServiceException.validation("Structural mutation removes a drawing anchor");
        anchor.put(property, shifted);
    }

    private static void shiftTargetReview(ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        SnapshotMutationSupport.remapReviewCoordinates(target, coordinate -> {
            int shifted = shiftIndex(axis == FormulaReferenceTransformer.Axis.ROW ? coordinate.row() : coordinate.column(), at, count, direction, axis);
            if (shifted < 0) throw ServiceException.validation("Structural delete would lose review metadata");
            return axis == FormulaReferenceTransformer.Axis.ROW
                    ? new SnapshotMutationSupport.CellCoordinate(shifted, coordinate.column())
                    : new SnapshotMutationSupport.CellCoordinate(coordinate.row(), shifted);
        });
    }

    private static void shiftTargetSpills(ObjectNode root, ObjectNode target, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "spillRanges")) {
            ObjectNode spill = requireObject(raw, "Spill range");
            requireShiftedRange(root, spill.get("range"), targetSheetId, axis, at, count, direction, "spill range");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(spill, "anchor");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int shifted = shiftIndex(anchor.path(key).asInt(-1), at, count, direction, axis);
            if (shifted < 0) throw ServiceException.validation("Structural mutation removes a spill anchor");
            anchor.put(key, shifted);
        }
    }

    private static void shiftTargetSheetTables(ObjectNode root, ObjectNode target, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "sheetTables")) {
            ObjectNode table = requireObject(raw, "Sheet table");
            requireShiftedRange(root, table.get("range"), targetSheetId, axis, at, count, direction, "sheet table range");
            JsonNode filterRaw = table.get("autoFilter");
            if (filterRaw != null && !filterRaw.isNull()) {
                ObjectNode filter = requireObject(filterRaw, "Table AutoFilter");
                if (!shiftRange(root, filter.get("range"), targetSheetId, axis, at, count, direction)) {
                    throw ServiceException.validation("Structural mutation removes a table AutoFilter range");
                }
                shiftAutoFilterSortState(root, filter, targetSheetId, axis, at, count, direction);
                shiftFilterColumns(SnapshotMutationSupport.requiredObject(filter, "columns"), axis, at, count, direction);
            }
        }
    }

    private static void shiftTargetProtectionAndOutline(ObjectNode root, ObjectNode target, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode protections = SnapshotMutationSupport.array(target, "protectionRules");
        for (JsonNode raw : protections) {
            ObjectNode rule = requireObject(raw, "Protection rule");
            if (rule.has("range")) requireShiftedRange(root, rule.get("range"), targetSheetId, axis, at, count, direction, "protection range");
        }
        JsonNode banded = target.get("bandedRule");
        if (banded != null && banded.isObject()) requireShiftedRange(root, banded.get("range"), targetSheetId, axis, at, count, direction, "banded range");
        JsonNode outlineRaw = target.get("outline");
        if (outlineRaw == null || !outlineRaw.isObject()) return;
        ArrayNode groups = SnapshotMutationSupport.array((ObjectNode) outlineRaw, "groups");
        for (int index = groups.size() - 1; index >= 0; index--) {
            ObjectNode group = requireObject(groups.get(index), "Outline group");
            String expectedAxis = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            if (!expectedAxis.equals(group.path("axis").asText())) continue;
            ObjectNode range = JsonNodeFactory.instance.objectNode();
            range.put("sheetId", targetSheetId);
            if (axis == FormulaReferenceTransformer.Axis.ROW) {
                range.put("startRow", group.path("start").asInt());
                range.put("endRow", group.path("end").asInt());
                range.put("startColumn", 0);
                range.put("endColumn", 0);
            } else {
                range.put("startRow", 0);
                range.put("endRow", 0);
                range.put("startColumn", group.path("start").asInt());
                range.put("endColumn", group.path("end").asInt());
            }
            requireShiftedRange(root, range, targetSheetId, axis, at, count, direction, "outline group");
            group.put("start", range.path(axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn").asInt());
            group.put("end", range.path(axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn").asInt());
        }
    }

    private static void shiftWorkbookTables(ObjectNode root, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : workbookTables(root)) {
            ObjectNode table = requireObject(raw, "Workbook table");
            if (table.has("sourceRange")) requireShiftedRange(root, table.get("sourceRange"), targetSheetId, axis, at, count, direction, "workbook table source range");
        }
    }

    private static void shiftDataSources(ObjectNode root, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.dataModelArray(root, "sources")) {
            ObjectNode source = requireObject(raw, "Data source");
            JsonNode sourceRange = source.get("sourceRange");
            if (sourceRange == null || sourceRange.isNull() || !targetSheetId.equals(sourceRange.path("sheetId").asText())) continue;
            RangeRef before = SnapshotMutationSupport.range(root, sourceRange);
            requireShiftedRange(root, sourceRange, targetSheetId, axis, at, count, direction, "data source range");
            RangeRef after = SnapshotMutationSupport.range(root, sourceRange);
            if (before.endRow() - before.startRow() != after.endRow() - after.startRow()
                    || before.endColumn() - before.startColumn() != after.endColumn() - after.startColumn()) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: structural edit changes the physical extent of data source " + source.path("id").asText());
            }
        }
    }

    private static void shiftDataRegions(ObjectNode root, ObjectNode target, String targetSheetId,
            FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "dataRegions")) {
            ObjectNode region = requireObject(raw, "Data region");
            requireShiftedRange(root, region.get("range"), targetSheetId, axis, at, count, direction, "data region range");
            if (axis == FormulaReferenceTransformer.Axis.ROW) {
                int header = region.path("headerRow").asInt(-1);
                if (header < 0) throw ServiceException.validation("Data region header row is invalid");
                int shifted = shiftIndex(header, at, count, direction, FormulaReferenceTransformer.Axis.ROW);
                if (shifted < 0) throw ServiceException.validation("Structural mutation removes a data region header");
                region.put("headerRow", shifted);
            }
        }
    }

    private static void validateAxisDataRegionPreservation(ObjectNode root, ObjectNode target,
            FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        String targetId = target.path("id").asText();
        int operationEnd = at + count - 1;
        for (JsonNode raw : SnapshotMutationSupport.array(target, "dataRegions")) {
            ObjectNode region = requireObject(raw, "Data region");
            RangeRef range = SnapshotMutationSupport.range(root, region.get("range"));
            int start = axis == FormulaReferenceTransformer.Axis.ROW ? range.startRow() : range.startColumn();
            int end = axis == FormulaReferenceTransformer.Axis.ROW ? range.endRow() : range.endColumn();
            boolean shiftsWholeRegion = direction == FormulaReferenceTransformer.Direction.INSERT ? at <= start : operationEnd < start;
            if (!shiftsWholeRegion && at <= end) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: structural edit intersects data region " + region.path("id").asText() + " and requires a data-block transaction");
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.dataModelArray(root, "sources")) {
            ObjectNode source = requireObject(raw, "Data source");
            JsonNode rawRange = source.get("sourceRange");
            if (rawRange == null || rawRange.isNull() || !targetId.equals(rawRange.path("sheetId").asText())) continue;
            RangeRef range = SnapshotMutationSupport.range(root, rawRange);
            int start = axis == FormulaReferenceTransformer.Axis.ROW ? range.startRow() : range.startColumn();
            int end = axis == FormulaReferenceTransformer.Axis.ROW ? range.endRow() : range.endColumn();
            boolean shiftsWholeSource = direction == FormulaReferenceTransformer.Direction.INSERT ? at <= start : operationEnd < start;
            if (!shiftsWholeSource && at <= end) {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: structural edit intersects data source " + source.path("id").asText() + " and requires a data-block transaction");
            }
        }
        if (direction == FormulaReferenceTransformer.Direction.INSERT) {
            for (JsonNode raw : workbookTables(root)) {
                ObjectNode table = requireObject(raw, "Workbook table");
                JsonNode rawRange = table.get("sourceRange");
                if (rawRange == null || rawRange.isNull() || !targetId.equals(rawRange.path("sheetId").asText())) continue;
                RangeRef range = SnapshotMutationSupport.range(root, rawRange);
                int start = axis == FormulaReferenceTransformer.Axis.ROW ? range.startRow() : range.startColumn();
                int end = axis == FormulaReferenceTransformer.Axis.ROW ? range.endRow() : range.endColumn();
                if (at > start && at <= end) {
                    throw ServiceException.unavailable("UNSUPPORTED_FEATURE: structural edit intersects workbook table "
                            + table.path("id").asText() + " and requires a table transaction");
                }
            }
        }
    }

    private static void shiftHyperlinks(ObjectNode root, ObjectNode targetSheet,
            FormulaReferenceTransformer.Axis axis, int at, int count,
            FormulaReferenceTransformer.Direction direction) {
        String targetSheetId = targetSheet.path("id").asText();
        FormulaReferenceTransformer.SheetIdentity target = identity(targetSheet);
        List<FormulaReferenceTransformer.SheetIdentity> sheetOrder = worksheetOrder(root);
        for (JsonNode rawOwner : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawOwner, "Sheet");
            boolean ownerIsTarget = targetSheetId.equals(owner.path("id").asText());
            for (JsonNode rawEntry : SnapshotMutationSupport.array(owner, "hyperlinks")) {
                ObjectNode entry = requireObject(rawEntry, "Hyperlink entry");
                if (ownerIsTarget) {
                    shiftHyperlinkCoordinate(entry, axis, at, count, direction, "hyperlink anchor");
                }
                ObjectNode hyperlink = SnapshotMutationSupport.requiredObject(entry, "hyperlink");
                JsonNode rawTarget = hyperlink.get("target");
                if (rawTarget == null || !rawTarget.isObject() || !"sheet".equals(rawTarget.path("kind").asText())
                        || !targetSheetId.equals(rawTarget.path("sheetId").asText())) continue;
                ObjectNode linkTarget = (ObjectNode) rawTarget;
                if (linkTarget.has("row")) shiftHyperlinkCoordinate(linkTarget, axis, at, count, direction, "hyperlink target");
                JsonNode address = linkTarget.get("address");
                if (address != null && address.isTextual()) {
                    linkTarget.put("address", FormulaReferenceTransformer.remapAxis(address.asText(), target, target, axis, at, count, direction, sheetOrder));
                }
            }
        }
    }

    private static void shiftHyperlinkCoordinate(ObjectNode coordinate, FormulaReferenceTransformer.Axis axis,
            int at, int count, FormulaReferenceTransformer.Direction direction, String label) {
        String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
        if (!coordinate.path(key).isIntegralNumber()) return;
        int shifted = shiftIndex(coordinate.path(key).intValue(), at, count, direction, axis);
        if (shifted < 0) throw ServiceException.validation("Structural mutation removes " + label);
        coordinate.put(key, shifted);
    }

    private static List<RuleFormulaSnapshot> captureRuleFormulaSnapshots(ObjectNode root) {
        List<RuleFormulaSnapshot> snapshots = new ArrayList<>();
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawSheet, "Sheet");
            String sheetId = identity(owner).id();
            for (String property : List.of("conditionalFormats", "dataValidations")) {
                String ruleKind = "conditionalFormats".equals(property) ? "conditional-format" : "data-validation";
                ArrayNode rules = SnapshotMutationSupport.array(owner, property);
                Map<String, Integer> idCounts = new HashMap<>();
                for (JsonNode rawRule : rules) idCounts.merge(rawRule.path("id").asText(), 1, Integer::sum);
                for (JsonNode rawRule : rules) {
                    ObjectNode rule = requireObject(rawRule, "Range rule");
                    Map<String, String> formulas = ruleFormulaFields(rule);
                    if (formulas.isEmpty()) continue;
                    String ruleId = rule.path("id").asText();
                    if (ruleId.isBlank() || idCounts.get(ruleId) != 1 || !sheetId.equals(rule.path("sheetId").asText())) {
                        throw ServiceException.validation("Structural formula rule requires a stable worksheet identity");
                    }
                    snapshots.add(new RuleFormulaSnapshot(rule, sheetId, ruleKind, ruleId,
                            ruleRanges(root, rule), formulas));
                }
            }
        }
        return List.copyOf(snapshots);
    }

    private static Map<String, String> ruleFormulaFields(ObjectNode rule) {
        Map<String, String> formulas = new LinkedHashMap<>();
        boolean formulaOperator = "formula".equals(rule.path("operator").asText());
        JsonNode value1 = rule.get("value1");
        if (value1 != null && value1.isTextual()
                && (formulaOperator || value1.asText().stripLeading().startsWith("="))) {
            formulas.put("value1", value1.asText());
        } else {
            if (value1 != null && value1.isTextual() && value1.asText().stripLeading().startsWith("=")) {
                formulas.put("value1", value1.asText());
            }
            JsonNode value2 = rule.get("value2");
            if (value2 != null && value2.isTextual() && value2.asText().stripLeading().startsWith("=")) {
                formulas.put("value2", value2.asText());
            }
        }
        JsonNode formula1 = rule.get("formula1");
        if (formula1 != null && formula1.isTextual() && !formula1.asText().isEmpty()
                && (formula1.asText().stripLeading().startsWith("=") || formulaOperator || "custom".equals(rule.path("type").asText()))) {
            formulas.put("formula1", formula1.asText());
        }
        JsonNode formula2 = rule.get("formula2");
        if (formula2 != null && formula2.isTextual() && !formula2.asText().isEmpty()
                && (formula2.asText().stripLeading().startsWith("=") || "custom".equals(rule.path("type").asText()))) {
            formulas.put("formula2", formula2.asText());
        }
        JsonNode source = rule.get("listSource");
        if (source != null && source.isObject() && "formula".equals(source.path("kind").asText())) {
            JsonNode formula = source.get("formula");
            if (formula != null && formula.isTextual()) formulas.put("listSource.formula", formula.asText());
        }
        return formulas;
    }

    private static void appendRuleFormulaDeltas(
            ObjectNode root,
            List<RuleFormulaSnapshot> snapshots,
            List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas
    ) {
        for (RuleFormulaSnapshot snapshot : snapshots) {
            Map<String, String> afterFormulas = ruleFormulaFields(snapshot.rule());
            List<RangeRef> afterRanges = ruleRanges(root, snapshot.rule());
            for (Map.Entry<String, String> before : snapshot.formulas().entrySet()) {
                String after = afterFormulas.get(before.getKey());
                if (after == null) {
                    throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: formula rule owner disappeared: "
                            + snapshot.sheetId() + ":" + snapshot.ruleId() + "." + before.getKey());
                }
                if (!before.getValue().equals(after)) {
                    formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaRule(
                            snapshot.sheetId(), snapshot.ruleKind(), snapshot.ruleId(), before.getKey(),
                            before.getValue(), after, snapshot.ranges(), afterRanges));
                }
            }
        }
    }

    private static List<StructuralPatch.FormulaOwnerDelta> rewriteAxisFormulas(
            ObjectNode root,
            ObjectNode targetSheet,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction,
            List<RuleFormulaSnapshot> ruleFormulaSnapshots
    ) {
        FormulaReferenceTransformer.SheetIdentity target = identity(targetSheet);
        List<FormulaReferenceTransformer.SheetIdentity> sheetOrder = worksheetOrder(root);
        List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas = new ArrayList<>();
        FormulaReferenceTransformer.Direction inverseDirection = direction == FormulaReferenceTransformer.Direction.INSERT
                ? FormulaReferenceTransformer.Direction.DELETE : FormulaReferenceTransformer.Direction.INSERT;
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity ownerIdentity = identity(owner);
            rewriteCellFormulaOwners(owner,
                    formula -> FormulaReferenceTransformer.remapAxis(formula, ownerIdentity, target, axis, at, count, direction, sheetOrder),
                    "axis shift",
                    formulaOwnerDeltas,
                    entry -> axisFormulaOwnerBeforeAddress(ownerIdentity.id(), entry, target.id(), axis, at, count, direction));
            for (String property : List.of("conditionalFormats", "dataValidations")) {
                for (JsonNode ruleRaw : SnapshotMutationSupport.array(owner, property)) {
                    ObjectNode rule = requireObject(ruleRaw, "Range rule");
                    String formulaOwnerId = rule.path("formulaAnchor").path("sheetId").asText(rule.path("sheetId").asText(ownerIdentity.id()));
                    FormulaReferenceTransformer.SheetIdentity formulaOwner = formulaOwnerId.equals(ownerIdentity.id())
                            ? ownerIdentity
                            : identity(SnapshotMutationSupport.sheet(root, formulaOwnerId));
                    rewriteRuleFormulas(rule, formula -> FormulaReferenceTransformer.remapAxis(
                            formula, formulaOwner, target, axis, at, count, direction, sheetOrder));
                }
            }
        }
        Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> definedNameOwners = definedNameFormulaOwners(root, target);
        ObjectNode names = SnapshotMutationSupport.object(root, "definedNames");
        names.fields().forEachRemaining(entry -> {
            if (!entry.getValue().isTextual()) return;
            FormulaReferenceTransformer.SheetIdentity owner = definedNameProjectionOwner(definedNameOwners, entry.getKey(), target);
            String rewritten = requireReversibleStructuralFormula(
                    entry.getValue().asText(),
                    value -> FormulaReferenceTransformer.remapAxis(value, owner, target, axis, at, count, direction, sheetOrder),
                    value -> FormulaReferenceTransformer.remapAxis(value, owner, target, axis, at, count, inverseDirection, sheetOrder),
                    "defined-name projection " + entry.getKey());
            names.put(entry.getKey(), rewritten);
        });
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            if (!name.path("formula").isTextual()) continue;
            ObjectNode anchor = definedNameAnchorOnSheet(name, target.id());
            if (anchor != null) {
                String coordinateKey = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
                int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW : SnapshotMutationSupport.MAX_COLUMN;
                int coordinate = definedNameAnchorCoordinate(anchor, coordinateKey, maximum);
                int shifted = shiftIndex(coordinate, at, count, direction, axis);
                if (shifted < 0) throw ServiceException.validation("Structural mutation removes defined-name anchor: " + name.path("name").asText());
                if (shifted > maximum) throw ServiceException.validation("Structural mutation moves defined-name anchor outside worksheet bounds");
                anchor.put(coordinateKey, shifted);
            }
            FormulaReferenceTransformer.SheetIdentity owner = definedNameFormulaOwner(definedNameOwners, name);
            String rewritten = requireReversibleStructuralFormula(
                    name.path("formula").asText(),
                    value -> FormulaReferenceTransformer.remapAxis(value, owner, target, axis, at, count, direction, sheetOrder),
                    value -> FormulaReferenceTransformer.remapAxis(value, owner, target, axis, at, count, inverseDirection, sheetOrder),
                    "defined name " + name.path("name").asText());
            name.put("formula", rewritten);
        }
        formulaOwnerDeltas.addAll(rewritePersistedFormulaOwners(root, target,
                (formula, owner) -> requireReversibleStructuralFormula(
                        formula,
                        value -> FormulaReferenceTransformer.remapAxis(value, owner, target, axis, at, count, direction, sheetOrder),
                        value -> FormulaReferenceTransformer.remapAxis(value, owner, target, axis, at, count, inverseDirection, sheetOrder),
                        "persisted formula owner on " + owner.id()),
                anchor -> shiftTemplateFormulaAnchor(anchor, target.id(), axis, at, count, direction),
                true));
        appendRuleFormulaDeltas(root, ruleFormulaSnapshots, formulaOwnerDeltas);
        return List.copyOf(formulaOwnerDeltas);
    }

    private static ArrayNode workbookTables(ObjectNode root) {
        JsonNode dataModel = root.get("dataModel");
        if (dataModel == null || dataModel.isNull()) return JsonNodeFactory.instance.arrayNode();
        if (!dataModel.isObject()) throw ServiceException.validation("dataModel must be an object");
        JsonNode tables = dataModel.get("tables");
        if (tables == null || tables.isNull()) return JsonNodeFactory.instance.arrayNode();
        if (!tables.isArray()) throw ServiceException.validation("dataModel.tables must be an array");
        return (ArrayNode) tables;
    }

    private static FormulaReferenceTransformer.SheetIdentity identity(ObjectNode sheet) {
        return new FormulaReferenceTransformer.SheetIdentity(SnapshotMutationSupport.text(sheet, "id"), SnapshotMutationSupport.text(sheet, "name"));
    }

    private static List<FormulaReferenceTransformer.SheetIdentity> worksheetOrder(ObjectNode root) {
        List<FormulaReferenceTransformer.SheetIdentity> result = new ArrayList<>();
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) result.add(identity(requireObject(raw, "Sheet")));
        return result;
    }

    private static Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> definedNameFormulaOwners(
            ObjectNode root,
            FormulaReferenceTransformer.SheetIdentity fallback
    ) {
        Map<String, FormulaReferenceTransformer.SheetIdentity> sheetIdentities = new HashMap<>();
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity sheetIdentity = identity(sheet);
            sheetIdentities.put(sheetIdentity.id(), sheetIdentity);
        }

        Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> owners = new HashMap<>();
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            String scope = name.path("scope").asText();
            String sheetId = name.path("sheetId").asText(null);
            String ownerSheetId = name.path("anchor").path("sheetId").asText(name.path("sheetId").asText(fallback.id()));
            FormulaReferenceTransformer.SheetIdentity owner = sheetIdentities.get(ownerSheetId);
            if (owner != null) owners.put(definedNameKey(scope, sheetId, name.path("name").asText()), owner);
        }
        return owners;
    }

    private static FormulaReferenceTransformer.SheetIdentity definedNameFormulaOwner(
            Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> owners,
            JsonNode name
    ) {
        String scope = name.path("scope").asText();
        String sheetId = name.path("sheetId").asText(null);
        FormulaReferenceTransformer.SheetIdentity owner = owners.get(definedNameKey(scope, sheetId, name.path("name").asText()));
        if (owner == null) throw ServiceException.validation("Defined-name formula owner is unresolved");
        return owner;
    }

    private static FormulaReferenceTransformer.SheetIdentity definedNameProjectionOwner(
            Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> owners,
            String name,
            FormulaReferenceTransformer.SheetIdentity fallback
    ) {
        return owners.getOrDefault(definedNameKey("workbook", null, name), fallback);
    }

    private static ObjectNode definedNameAnchorOnSheet(ObjectNode name, String sheetId) {
        JsonNode raw = name.get("anchor");
        if (raw == null || raw.isNull()) return null;
        ObjectNode anchor = requireObject(raw, "Defined-name anchor");
        return sheetId.equals(SnapshotMutationSupport.text(anchor, "sheetId")) ? anchor : null;
    }

    private static void preflightAxisFormulaAnchors(
            ObjectNode root,
            String targetSheetId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        String coordinateKey = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW : SnapshotMutationSupport.MAX_COLUMN;
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            if (!name.path("formula").isTextual()) continue;
            ObjectNode anchor = definedNameAnchorOnSheet(name, targetSheetId);
            if (anchor == null) continue;
            int position = definedNameAnchorCoordinate(anchor, coordinateKey, maximum);
            int mapped = shiftIndex(position, at, count, direction, axis);
            if (mapped < 0 || mapped > maximum) {
                throw ServiceException.validation("Structural mutation removes defined-name anchor: " + name.path("name").asText());
            }
        }
        preflightTemplateFormulaAnchors(root, anchor -> shiftTemplateFormulaAnchor(anchor, targetSheetId, axis, at, count, direction));
    }

    private static void preflightCellShiftFormulaAnchors(
            ObjectNode root,
            String targetSheetId,
            RangeRef selection,
            FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction
    ) {
        FormulaReferenceTransformer.Range selected = formulaRange(selection);
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            if (!name.path("formula").isTextual()) continue;
            ObjectNode anchor = definedNameAnchorOnSheet(name, targetSheetId);
            if (anchor == null) continue;
            int row = definedNameAnchorCoordinate(anchor, "row", SnapshotMutationSupport.MAX_ROW);
            int column = definedNameAnchorCoordinate(anchor, "column", SnapshotMutationSupport.MAX_COLUMN);
            if (FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, selected, axis, direction) == null) {
                throw ServiceException.validation("Cell shift removes defined-name anchor: " + name.path("name").asText());
            }
        }
        preflightTemplateFormulaAnchors(root, anchor -> shiftTemplateFormulaAnchor(anchor, targetSheetId, selected, axis, direction));
    }

    private static void preflightTemplateFormulaAnchors(ObjectNode root, Function<ObjectNode, ObjectNode> mapper) {
        JsonNode rawTemplates = root.get("cellStyleTemplates");
        if (rawTemplates == null || rawTemplates.isNull()) return;
        if (!rawTemplates.isArray()) throw ServiceException.validation("cellStyleTemplates must be an array");
        for (JsonNode rawTemplate : rawTemplates) {
            ObjectNode template = requireObject(rawTemplate, "Cell style template");
            JsonNode rawValidation = template.get("dataValidation");
            if (rawValidation == null || rawValidation.isNull()) continue;
            ObjectNode validation = requireObject(rawValidation, "Cell style template validation");
            JsonNode rawAnchor = validation.get("formulaAnchor");
            if (rawAnchor == null) continue;
            mapper.apply(requireObject(rawAnchor, "Cell style template formulaAnchor"));
        }
    }

    private static int definedNameAnchorCoordinate(ObjectNode anchor, String key, int maximum) {
        JsonNode coordinate = anchor.get(key);
        if (coordinate == null || !coordinate.isIntegralNumber() || coordinate.longValue() < 0 || coordinate.longValue() > maximum) {
            throw ServiceException.validation("Defined-name anchor coordinate is invalid");
        }
        return coordinate.intValue();
    }

    private static DefinedNameKey definedNameKey(String scope, String sheetId, String name) {
        return new DefinedNameKey(scope, sheetId, name.toUpperCase(java.util.Locale.ROOT));
    }

    private record DefinedNameKey(String scope, String sheetId, String normalizedName) {
    }

    private static void forEachFormulaCell(ObjectNode sheet, java.util.function.Consumer<ObjectNode> consumer) {
        forEachCell(sheet, entry -> {
            if (entry.cell().path("formula").isTextual()) consumer.accept(entry.cell());
        });
    }

    private static void forEachCell(ObjectNode sheet, java.util.function.Consumer<CellEntry> consumer) {
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        cells.fields().forEachRemaining(row -> {
            if (!row.getValue().isObject()) throw ServiceException.validation("Cell row must be an object");
            int rowIndex = integerKey(row.getKey(), SnapshotMutationSupport.MAX_ROW, "Cell row");
            ((ObjectNode) row.getValue()).fields().forEachRemaining(column -> {
                if (!column.getValue().isObject()) throw ServiceException.validation("Cell payload must be an object");
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                consumer.accept(new CellEntry(rowIndex, columnIndex, (ObjectNode) column.getValue()));
            });
        });
    }

    private static void rewriteCellFormulaOwners(ObjectNode sheet, Function<String, String> mapper, String operation) {
        rewriteCellFormulaOwners(sheet, mapper, operation, null, null);
    }

    private static void rewriteCellFormulaOwners(
            ObjectNode sheet,
            Function<String, String> mapper,
            String operation,
            List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas,
            Function<CellEntry, StructuralPatch.CellAddress> beforeAddressResolver
    ) {
        String sheetId = sheet.path("id").asText();
        forEachCell(sheet, entry -> {
            ObjectNode cell = entry.cell();
            String original = cell.path("formula").isTextual() ? cell.path("formula").asText() : null;
            String sourceFormula = cell.path("formulaMetadata").path("sourceFormula").isTextual()
                    ? cell.path("formulaMetadata").path("sourceFormula").asText() : null;
            JsonNode rawPresentation = cell.get("presentation");
            JsonNode rawBarcodeSource = rawPresentation != null && rawPresentation.isObject()
                    && "barcode".equals(rawPresentation.path("kind").asText())
                    ? rawPresentation.get("source") : null;
            String barcodeFormula = rawBarcodeSource != null && rawBarcodeSource.isObject()
                    && "formula".equals(rawBarcodeSource.path("kind").asText())
                    && rawBarcodeSource.path("formula").isTextual()
                    ? rawBarcodeSource.path("formula").asText() : null;
            StructuralPatch.FormulaOwnerState before = new StructuralPatch.FormulaOwnerState(original, sourceFormula, barcodeFormula);
            String rewritten = original == null ? null : mapper.apply(original);
            String rewrittenSourceFormula = sourceFormula == null ? null : mapper.apply(sourceFormula);
            String rewrittenBarcodeFormula = barcodeFormula == null ? null : mapper.apply(barcodeFormula);
            boolean formulaChanged = original != null && !original.equals(rewritten);
            boolean sourceFormulaChanged = sourceFormula != null && !sourceFormula.equals(rewrittenSourceFormula);
            boolean barcodeFormulaChanged = barcodeFormula != null && !barcodeFormula.equals(rewrittenBarcodeFormula);
            if ((formulaChanged || sourceFormulaChanged || barcodeFormulaChanged) && hasFormulaGroupMetadata(cell)) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: formula group at "
                        + sheetId + "!" + entry.row() + ":" + entry.column()
                        + " requires an explicit formula-group operation before " + operation);
            }
            if (formulaChanged) cell.put("formula", rewritten);
            if (sourceFormulaChanged) SnapshotMutationSupport.requiredObject(cell, "formulaMetadata").put("sourceFormula", rewrittenSourceFormula);
            if (barcodeFormulaChanged) ((ObjectNode) rawBarcodeSource).put("formula", rewrittenBarcodeFormula);
            if (original != null) cell.remove("formulaValue");
            if (formulaOwnerDeltas != null && (formulaChanged || sourceFormulaChanged || barcodeFormulaChanged)) {
                StructuralPatch.CellAddress beforeAddress = beforeAddressResolver.apply(entry);
                formulaOwnerDeltas.add(new StructuralPatch.FormulaOwnerDelta(
                        "formula-cell",
                        beforeAddress,
                        new StructuralPatch.CellAddress(sheetId, entry.row(), entry.column()),
                        before,
                        formulaOwnerState(cell)));
            }
        });
    }

    private static StructuralPatch.CellAddress axisFormulaOwnerBeforeAddress(
            String ownerSheetId,
            CellEntry entry,
            String targetSheetId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        int row = entry.row();
        int column = entry.column();
        if (ownerSheetId.equals(targetSheetId)) {
            int position = axis == FormulaReferenceTransformer.Axis.ROW ? row : column;
            if (direction == FormulaReferenceTransformer.Direction.INSERT) {
                if (position >= at && position < at + count) {
                    throw ServiceException.unavailable("STRUCTURAL_PATCH_INVARIANT: formula owner is inside the inserted axis band");
                }
                if (position >= at + count) position -= count;
            } else if (position >= at) {
                position += count;
            }
            if (axis == FormulaReferenceTransformer.Axis.ROW) row = position;
            else column = position;
        }
        return new StructuralPatch.CellAddress(ownerSheetId, row, column);
    }

    private static StructuralPatch.FormulaOwnerState formulaOwnerState(ObjectNode cell) {
        String formula = cell.path("formula").isTextual() ? cell.path("formula").asText() : null;
        String sourceFormula = cell.path("formulaMetadata").path("sourceFormula").isTextual()
                ? cell.path("formulaMetadata").path("sourceFormula").asText() : null;
        JsonNode presentation = cell.get("presentation");
        JsonNode source = presentation != null && presentation.isObject() && "barcode".equals(presentation.path("kind").asText())
                ? presentation.get("source") : null;
        String barcodeFormula = source != null && source.isObject() && "formula".equals(source.path("kind").asText())
                && source.path("formula").isTextual() ? source.path("formula").asText() : null;
        return new StructuralPatch.FormulaOwnerState(formula, sourceFormula, barcodeFormula);
    }

    private static void rejectFormulaGroupMetadataInRange(ObjectNode sheet, RangeRef range, String operation) {
        for (CellEntry entry : cellsInRange(sheet, range)) {
            if (hasFormulaGroupMetadata(entry.cell())) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: "
                        + entry.cell().path("formulaMetadata").path("kind").asText("unknown")
                        + " formula metadata at " + range.sheetId() + "!" + entry.row() + ":" + entry.column()
                        + " requires an explicit formula-group operation before " + operation);
            }
        }
    }

    private static void rewriteRuleFormulas(ObjectNode rule, Function<String, String> mapper) {
        boolean formulaOperator = "formula".equals(rule.path("operator").asText());
        JsonNode value1 = rule.get("value1");
        if (value1 != null && value1.isTextual()
                && (formulaOperator || value1.asText().stripLeading().startsWith("="))) {
            rule.put("value1", mapper.apply(value1.asText()));
        } else {
            if (value1 != null && value1.isTextual() && value1.asText().stripLeading().startsWith("=")) {
                rule.put("value1", mapper.apply(value1.asText()));
            }
            JsonNode value2 = rule.get("value2");
            if (value2 != null && value2.isTextual() && value2.asText().stripLeading().startsWith("=")) {
                rule.put("value2", mapper.apply(value2.asText()));
            }
        }
        JsonNode formula1 = rule.get("formula1");
        if (formula1 != null && formula1.isTextual() && !formula1.asText().isEmpty()
                && (formula1.asText().stripLeading().startsWith("=") || formulaOperator || "custom".equals(rule.path("type").asText()))) {
            rule.put("formula1", mapper.apply(formula1.asText()));
        }
        JsonNode formula2 = rule.get("formula2");
        if (formula2 != null && formula2.isTextual() && !formula2.asText().isEmpty()
                && (formula2.asText().stripLeading().startsWith("=") || "custom".equals(rule.path("type").asText()))) {
            rule.put("formula2", mapper.apply(formula2.asText()));
        }
        JsonNode sourceRaw = rule.get("listSource");
        if (sourceRaw != null && sourceRaw.isObject() && "formula".equals(sourceRaw.path("kind").asText())) {
            ObjectNode source = (ObjectNode) sourceRaw;
            JsonNode formula = source.get("formula");
            if (formula == null || !formula.isTextual()) throw ServiceException.validation("Data-validation list formula is invalid");
            source.put("formula", mapper.apply(formula.asText()));
        }
    }

    private static StructuralPatch rewriteCellShiftFormulas(
            ObjectNode root,
            ObjectNode targetSheet,
            String mutationId,
            RangeRef selection,
            String axis,
            String operation,
            List<RuleFormulaSnapshot> ruleFormulaSnapshots,
            List<StructuralPatch.RangeOwnerDelta> rangeOwnerDeltas
    ) {
        FormulaReferenceTransformer.SheetIdentity target = identity(targetSheet);
        List<FormulaReferenceTransformer.SheetIdentity> sheetOrder = worksheetOrder(root);
        List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas = new ArrayList<>();
        FormulaReferenceTransformer.Axis shiftAxis = "row".equals(axis)
                ? FormulaReferenceTransformer.Axis.ROW
                : FormulaReferenceTransformer.Axis.COLUMN;
        FormulaReferenceTransformer.Direction direction = "insert".equals(operation)
                ? FormulaReferenceTransformer.Direction.INSERT
                : FormulaReferenceTransformer.Direction.DELETE;
        FormulaReferenceTransformer.Direction inverseDirection = direction == FormulaReferenceTransformer.Direction.INSERT
                ? FormulaReferenceTransformer.Direction.DELETE : FormulaReferenceTransformer.Direction.INSERT;
        FormulaReferenceTransformer.Range selected = new FormulaReferenceTransformer.Range(
                selection.startRow(), selection.endRow(), selection.startColumn(), selection.endColumn());
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity ownerIdentity = identity(owner);
            rewriteCellFormulaOwners(owner,
                    formula -> FormulaReferenceTransformer.remapCellShift(formula, ownerIdentity, target, selected, shiftAxis, direction, sheetOrder),
                    "cell shift",
                    formulaOwnerDeltas,
                    entry -> cellShiftFormulaOwnerBeforeAddress(ownerIdentity.id(), entry, target.id(), selection, shiftAxis, direction));
            for (String property : List.of("conditionalFormats", "dataValidations")) {
                for (JsonNode ruleRaw : SnapshotMutationSupport.array(owner, property)) {
                    ObjectNode rule = requireObject(ruleRaw, "Range rule");
                    String formulaOwnerId = rule.path("formulaAnchor").path("sheetId").asText(rule.path("sheetId").asText(ownerIdentity.id()));
                    FormulaReferenceTransformer.SheetIdentity formulaOwner = formulaOwnerId.equals(ownerIdentity.id())
                            ? ownerIdentity
                            : identity(SnapshotMutationSupport.sheet(root, formulaOwnerId));
                    rewriteRuleFormulas(rule, formula -> FormulaReferenceTransformer.remapCellShift(
                            formula, formulaOwner, target, selected, shiftAxis, direction, sheetOrder));
                }
            }
        }
        Map<DefinedNameKey, FormulaReferenceTransformer.SheetIdentity> definedNameOwners = definedNameFormulaOwners(root, target);
        ObjectNode names = SnapshotMutationSupport.object(root, "definedNames");
        names.fields().forEachRemaining(entry -> {
            if (!entry.getValue().isTextual()) return;
            FormulaReferenceTransformer.SheetIdentity owner = definedNameProjectionOwner(definedNameOwners, entry.getKey(), target);
            names.put(entry.getKey(), requireReversibleStructuralFormula(
                    entry.getValue().asText(),
                    value -> FormulaReferenceTransformer.remapCellShift(value, owner, target, selected, shiftAxis, direction, sheetOrder),
                    value -> FormulaReferenceTransformer.remapCellShift(value, owner, target, selected, shiftAxis, inverseDirection, sheetOrder),
                    "defined-name projection " + entry.getKey()));
        });
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            if (!name.path("formula").isTextual()) continue;
            ObjectNode anchor = definedNameAnchorOnSheet(name, target.id());
            if (anchor != null) {
                int row = definedNameAnchorCoordinate(anchor, "row", SnapshotMutationSupport.MAX_ROW);
                int column = definedNameAnchorCoordinate(anchor, "column", SnapshotMutationSupport.MAX_COLUMN);
                int[] mapped = FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, selected, shiftAxis, direction);
                if (mapped == null) throw ServiceException.validation("Cell shift removes defined-name anchor: " + name.path("name").asText());
                anchor.put("row", mapped[0]);
                anchor.put("column", mapped[1]);
            }
            FormulaReferenceTransformer.SheetIdentity owner = definedNameFormulaOwner(definedNameOwners, name);
            name.put("formula", requireReversibleStructuralFormula(
                    name.path("formula").asText(),
                    value -> FormulaReferenceTransformer.remapCellShift(value, owner, target, selected, shiftAxis, direction, sheetOrder),
                    value -> FormulaReferenceTransformer.remapCellShift(value, owner, target, selected, shiftAxis, inverseDirection, sheetOrder),
                    "defined name " + name.path("name").asText()));
        }
        formulaOwnerDeltas.addAll(rewritePersistedFormulaOwners(root, target,
                (formula, owner) -> requireReversibleStructuralFormula(
                        formula,
                        value -> FormulaReferenceTransformer.remapCellShift(value, owner, target, selected, shiftAxis, direction, sheetOrder),
                        value -> FormulaReferenceTransformer.remapCellShift(value, owner, target, selected, shiftAxis, inverseDirection, sheetOrder),
                        "persisted formula owner on " + owner.id()),
                anchor -> shiftTemplateFormulaAnchor(anchor, target.id(), selected, shiftAxis, direction),
                true));
        appendRuleFormulaDeltas(root, ruleFormulaSnapshots, formulaOwnerDeltas);
        return new StructuralPatch(StructuralPatch.VERSION, mutationId, formulaOwnerDeltas, List.of(), rangeOwnerDeltas);
    }

    private static Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> captureRangeOwnerSnapshots(
            ObjectNode root,
            String targetSheetId
    ) {
        Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> snapshots = new LinkedHashMap<>();
        ObjectNode target = SnapshotMutationSupport.sheet(root, targetSheetId);
        for (JsonNode raw : SnapshotMutationSupport.array(target, "dataRegions")) {
            ObjectNode region = requireObject(raw, "Data region");
            String regionId = SnapshotMutationSupport.text(region, "id");
            RangeRef range = SnapshotMutationSupport.range(root, region.get("range"));
            int headerRow = integer(region.get("headerRow"), "Data region header row", SnapshotMutationSupport.MAX_ROW);
            if (!targetSheetId.equals(range.sheetId()) || headerRow < range.startRow() || headerRow > range.endRow()) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: data-region owner geometry is invalid: " + regionId);
            }
            putRangeOwnerSnapshot(snapshots, new StructuralPatch.RangeOwnerKey("data-region", targetSheetId, null, regionId),
                    new RangeOwnerSnapshot("data-region", targetSheetId, regionId, null, range, headerRow));
        }
        for (JsonNode raw : workbookTables(root)) {
            ObjectNode table = requireObject(raw, "Workbook table");
            JsonNode rawRange = table.get("sourceRange");
            if (rawRange == null || rawRange.isNull()) continue;
            RangeRef range = SnapshotMutationSupport.range(root, rawRange);
            if (!targetSheetId.equals(range.sheetId())) continue;
            String ownerId = SnapshotMutationSupport.text(table, "id");
            putRangeOwnerSnapshot(snapshots, new StructuralPatch.RangeOwnerKey("workbook-table", null, ownerId, null),
                    new RangeOwnerSnapshot("workbook-table", null, null, ownerId, range, null));
        }
        for (JsonNode raw : SnapshotMutationSupport.dataModelArray(root, "sources")) {
            ObjectNode source = requireObject(raw, "Data source");
            JsonNode rawRange = source.get("sourceRange");
            if (rawRange == null || rawRange.isNull()) continue;
            RangeRef range = SnapshotMutationSupport.range(root, rawRange);
            if (!targetSheetId.equals(range.sheetId())) continue;
            String ownerId = SnapshotMutationSupport.text(source, "id");
            putRangeOwnerSnapshot(snapshots, new StructuralPatch.RangeOwnerKey("data-source", null, ownerId, null),
                    new RangeOwnerSnapshot("data-source", null, null, ownerId, range, null));
        }
        return snapshots;
    }

    private static void putRangeOwnerSnapshot(
            Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> snapshots,
            StructuralPatch.RangeOwnerKey key,
            RangeOwnerSnapshot snapshot
    ) {
        if (snapshots.putIfAbsent(key, snapshot) != null) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: range-owner identity is duplicated: " + key);
        }
    }

    private static List<StructuralPatch.RangeOwnerDelta> rangeOwnerDeltas(
            Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> before,
            Map<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> after
    ) {
        if (!before.keySet().equals(after.keySet())) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_INVARIANT: structural transform changed range-owner membership");
        }
        List<StructuralPatch.RangeOwnerDelta> deltas = new ArrayList<>();
        for (Map.Entry<StructuralPatch.RangeOwnerKey, RangeOwnerSnapshot> entry : before.entrySet()) {
            RangeOwnerSnapshot beforeState = entry.getValue();
            RangeOwnerSnapshot afterState = after.get(entry.getKey());
            if (beforeState.equals(afterState)) continue;
            if ("data-region".equals(beforeState.ownerKind())) {
                deltas.add(StructuralPatch.RangeOwnerDelta.dataRegion(
                        beforeState.sheetId(), beforeState.regionId(), beforeState.range(), beforeState.headerRow(),
                        afterState.range(), afterState.headerRow()));
            } else {
                deltas.add(StructuralPatch.RangeOwnerDelta.range(
                        beforeState.ownerKind(), beforeState.ownerId(), beforeState.range(), afterState.range()));
            }
        }
        return List.copyOf(deltas);
    }

    private static StructuralPatch.CellAddress cellShiftFormulaOwnerBeforeAddress(
            String ownerSheetId,
            CellEntry entry,
            String targetSheetId,
            RangeRef selection,
            FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction
    ) {
        int row = entry.row();
        int column = entry.column();
        if (ownerSheetId.equals(targetSheetId)) {
            int count = axis == FormulaReferenceTransformer.Axis.ROW
                    ? selection.endRow() - selection.startRow() + 1
                    : selection.endColumn() - selection.startColumn() + 1;
            if (axis == FormulaReferenceTransformer.Axis.ROW
                    && column >= selection.startColumn() && column <= selection.endColumn() && row >= selection.startRow()) {
                if (direction == FormulaReferenceTransformer.Direction.INSERT && row >= selection.startRow() + count) row -= count;
                else if (direction == FormulaReferenceTransformer.Direction.DELETE) row += count;
            } else if (axis == FormulaReferenceTransformer.Axis.COLUMN
                    && row >= selection.startRow() && row <= selection.endRow() && column >= selection.startColumn()) {
                if (direction == FormulaReferenceTransformer.Direction.INSERT && column >= selection.startColumn() + count) column -= count;
                else if (direction == FormulaReferenceTransformer.Direction.DELETE) column += count;
            }
        }
        return new StructuralPatch.CellAddress(ownerSheetId, row, column);
    }

    private static List<StructuralPatch.FormulaOwnerDelta> rewritePersistedFormulaOwners(
            ObjectNode root,
            FormulaReferenceTransformer.SheetIdentity target,
            BiFunction<String, FormulaReferenceTransformer.SheetIdentity, String> formulaMapper,
            Function<ObjectNode, ObjectNode> anchorMapper
    ) {
        return rewritePersistedFormulaOwners(root, target, formulaMapper, anchorMapper, false);
    }

    private static List<StructuralPatch.FormulaOwnerDelta> rewritePersistedFormulaOwners(
            ObjectNode root,
            FormulaReferenceTransformer.SheetIdentity target,
            BiFunction<String, FormulaReferenceTransformer.SheetIdentity, String> formulaMapper,
            Function<ObjectNode, ObjectNode> anchorMapper,
            boolean includeNonChartObjectDeltas
    ) {
        List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas = new ArrayList<>();
        // Global formulas only have relative-reference ownership when a persisted anchor supplies it.
        String workbookOwnerId = "__workbook_formula_owner__";
        while (target.id().equals(workbookOwnerId)) workbookOwnerId += "_";
        FormulaReferenceTransformer.SheetIdentity workbookOwner = new FormulaReferenceTransformer.SheetIdentity(workbookOwnerId, workbookOwnerId);
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity ownerIdentity = identity(owner);
            JsonNode tableSheetRaw = owner.get("tableSheet");
            if (tableSheetRaw != null && !tableSheetRaw.isNull()) {
                ObjectNode tableSheet = requireObject(tableSheetRaw, "TableSheet definition");
                for (JsonNode columnRaw : SnapshotMutationSupport.requiredArray(tableSheet, "columns")) {
                    ObjectNode column = requireObject(columnRaw, "TableSheet column");
                    FormulaChange change = rewriteOptionalFormula(column, "formula", ownerIdentity, formulaMapper,
                            "table-sheet:" + ownerIdentity.id() + "." + column.path("fieldId").asText());
                    if (includeNonChartObjectDeltas && change.changed()) {
                        formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                "table-sheet-column", ownerIdentity.id(), null, column.path("fieldId").asText(),
                                null, null, null, change.before(), change.after()));
                    }
                }
            }
            JsonNode payloadsRaw = owner.get("drawingPayloads");
            if (payloadsRaw != null && !payloadsRaw.isNull()) {
                if (!payloadsRaw.isObject()) throw ServiceException.validation("drawingPayloads must be an object");
                ObjectNode payloads = (ObjectNode) payloadsRaw;
                payloads.fields().forEachRemaining(entry -> {
                    ObjectNode payload = requireObject(entry.getValue(), "Drawing payload");
                    if ("shape".equals(payload.path("kind").asText())) {
                        FormulaChange change = rewriteOptionalFormula(payload, "propertyFormula", ownerIdentity, formulaMapper,
                                "drawing:" + entry.getKey() + ".propertyFormula");
                        if (includeNonChartObjectDeltas && change.changed()) {
                            formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                    "shape-property", ownerIdentity.id(), entry.getKey(), null, null, null,
                                    null, change.before(), change.after()));
                        }
                    } else if ("chart".equals(payload.path("kind").asText())) {
                        for (String field : CHART_TEXT_FORMULA_FIELDS) {
                            ObjectNode text = chartTextFormulaModel(payload, field);
                            JsonNode rawFormula = text.get("linkedFormula");
                            if (rawFormula == null || rawFormula.isNull()) continue;
                            if (!rawFormula.isTextual()) {
                                throw ServiceException.validation("drawing:" + entry.getKey() + "." + field + " must be text");
                            }
                            String before = rawFormula.asText();
                            if (before.isBlank()) {
                                throw ServiceException.validation("drawing:" + entry.getKey() + "." + field + " must be a non-empty formula");
                            }
                            String after = formulaMapper.apply(before, ownerIdentity);
                            if (before.equals(after)) continue;
                            text.put("linkedFormula", after);
                            formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                    ownerIdentity.id(), "chart-text", entry.getKey(), field, before, after));
                        }
                    }
                });
            }
        }

        JsonNode dataModelRaw = root.get("dataModel");
        if (dataModelRaw != null && !dataModelRaw.isNull()) {
            if (!dataModelRaw.isObject()) throw ServiceException.validation("dataModel must be an object");
            JsonNode viewsRaw = dataModelRaw.get("views");
            if (viewsRaw != null && !viewsRaw.isNull()) {
                if (!viewsRaw.isArray()) throw ServiceException.validation("views must be an array");
                for (JsonNode viewRaw : viewsRaw) {
                    ObjectNode view = requireObject(viewRaw, "Data view");
                    for (JsonNode fieldRaw : SnapshotMutationSupport.requiredArray(view, "fields")) {
                        ObjectNode field = requireObject(fieldRaw, "Data view field");
                        FormulaChange change = rewriteOptionalFormula(field, "formula", workbookOwner, formulaMapper,
                                "data-view:" + view.path("id").asText() + "." + field.path("fieldId").asText());
                        if (includeNonChartObjectDeltas && change.changed()) {
                            formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                    "data-view-field", null, null, field.path("fieldId").asText(),
                                    view.path("id").asText(), null, null, change.before(), change.after()));
                        }
                    }
                }
            }
        }

        JsonNode templatesRaw = root.get("cellStyleTemplates");
        if (templatesRaw != null && !templatesRaw.isNull()) {
            if (!templatesRaw.isArray()) throw ServiceException.validation("cellStyleTemplates must be an array");
            for (JsonNode templateRaw : templatesRaw) {
                ObjectNode template = requireObject(templateRaw, "Cell style template");
                JsonNode validationRaw = template.get("dataValidation");
                if (validationRaw == null || validationRaw.isNull()) continue;
                ObjectNode validation = requireObject(validationRaw, "Cell style template validation");
                JsonNode anchorRaw = validation.get("formulaAnchor");
                FormulaReferenceTransformer.SheetIdentity formulaOwner = workbookOwner;
                if (anchorRaw != null) {
                    if (!anchorRaw.isObject()) throw ServiceException.validation("Cell style template formulaAnchor must be an object");
                    ObjectNode anchor = (ObjectNode) anchorRaw;
                    String ownerSheetId = SnapshotMutationSupport.text(anchor, "sheetId");
                    ObjectNode ownerSheet = SnapshotMutationSupport.sheet(root, ownerSheetId);
                    formulaOwner = identity(ownerSheet);
                    definedNameAnchorCoordinate(anchor, "row", SnapshotMutationSupport.MAX_ROW);
                    definedNameAnchorCoordinate(anchor, "column", SnapshotMutationSupport.MAX_COLUMN);
                    ObjectNode mappedAnchor = anchorMapper.apply(anchor);
                    if (!anchor.equals(mappedAnchor)) validation.set("formulaAnchor", mappedAnchor);
                }
                String templateId = template.path("id").asText();
                JsonNode formula1 = validation.get("formula1");
                String type = validation.path("type").asText();
                if (formula1 != null && formula1.isTextual() && !formula1.asText().isEmpty()
                        && (formula1.asText().stripLeading().startsWith("=") || "custom".equals(type))) {
                    FormulaChange change = rewriteOptionalFormula(validation, "formula1", formulaOwner, formulaMapper,
                            "cell-style-template:" + templateId + ".formula1");
                    if (includeNonChartObjectDeltas && change.changed()) {
                        formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                "cell-style-template", null, null, null, null, templateId,
                                "formula1", change.before(), change.after()));
                    }
                } else if (formula1 != null && !formula1.isNull() && !formula1.isTextual()) {
                    throw ServiceException.validation("Cell style template formula1 must be text");
                }
                JsonNode formula2 = validation.get("formula2");
                if (formula2 != null && formula2.isTextual() && !formula2.asText().isEmpty()
                        && (formula2.asText().stripLeading().startsWith("=") || "custom".equals(type))) {
                    FormulaChange change = rewriteOptionalFormula(validation, "formula2", formulaOwner, formulaMapper,
                            "cell-style-template:" + templateId + ".formula2");
                    if (includeNonChartObjectDeltas && change.changed()) {
                        formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                "cell-style-template", null, null, null, null, templateId,
                                "formula2", change.before(), change.after()));
                    }
                } else if (formula2 != null && !formula2.isNull() && !formula2.isTextual()) {
                    throw ServiceException.validation("Cell style template formula2 must be text");
                }
                JsonNode listSourceRaw = validation.get("listSource");
                if (listSourceRaw != null && !listSourceRaw.isNull()) {
                    ObjectNode listSource = requireObject(listSourceRaw, "Cell style template listSource");
                    if ("formula".equals(listSource.path("kind").asText())) {
                        JsonNode formula = listSource.get("formula");
                        if (formula == null || !formula.isTextual()) {
                            throw ServiceException.validation("Cell style template list formula must be text");
                        }
                        FormulaChange change = rewriteOptionalFormula(listSource, "formula", formulaOwner, formulaMapper,
                                "cell-style-template:" + templateId + ".listSource");
                        if (includeNonChartObjectDeltas && change.changed()) {
                            formulaOwnerDeltas.add(StructuralPatch.FormulaOwnerDelta.formulaObject(
                                    "cell-style-template", null, null, null, null, templateId,
                                    "listSource.formula", change.before(), change.after()));
                        }
                    }
                }
            }
        }
        return List.copyOf(formulaOwnerDeltas);
    }

    private static ObjectNode chartTextFormulaModel(ObjectNode payload, String field) {
        JsonNode rawElements = payload.get("elements");
        if (rawElements == null || !rawElements.isObject()) {
            throw ServiceException.validation("Chart elements must be an object when linked text formulas are inspected");
        }
        ObjectNode elements = (ObjectNode) rawElements;
        String ownerField = switch (field) {
            case "titleText.linkedFormula" -> "titleText";
            case "legend.text.linkedFormula" -> "legend.text";
            case "categoryAxis.titleText.linkedFormula" -> "categoryAxis.titleText";
            case "valueAxis.titleText.linkedFormula" -> "valueAxis.titleText";
            case "secondaryCategoryAxis.titleText.linkedFormula" -> "secondaryCategoryAxis.titleText";
            case "secondaryValueAxis.titleText.linkedFormula" -> "secondaryValueAxis.titleText";
            case "dataTable.font.linkedFormula" -> "dataTable.font";
            default -> throw ServiceException.validation("Unsupported chart text formula field: " + field);
        };
        ObjectNode current = elements;
        for (String segment : ownerField.split("\\.")) {
            JsonNode child = current.get(segment);
            if (child == null) return JsonNodeFactory.instance.objectNode();
            if (child.isNull()) throw ServiceException.validation("Chart text formula owner must not be null: " + field);
            if (!child.isObject()) throw ServiceException.validation("Chart text formula owner must be an object: " + field);
            current = (ObjectNode) child;
        }
        return current;
    }

    private static FormulaChange rewriteOptionalFormula(
            ObjectNode owner,
            String property,
            FormulaReferenceTransformer.SheetIdentity formulaOwner,
            BiFunction<String, FormulaReferenceTransformer.SheetIdentity, String> formulaMapper,
            String participant
    ) {
        JsonNode raw = owner.get(property);
        if (raw == null || raw.isNull()) return new FormulaChange(null, null);
        if (!raw.isTextual()) throw ServiceException.validation(participant + " must be text");
        String formula = raw.asText();
        if (formula.isEmpty()) return new FormulaChange(formula, formula);
        String after = formulaMapper.apply(formula, formulaOwner);
        if (!formula.equals(after)) owner.put(property, after);
        return new FormulaChange(formula, after);
    }

    private static String requireReversibleStructuralFormula(
            String before,
            Function<String, String> forward,
            Function<String, String> inverse,
            String participant
    ) {
        String after = forward.apply(before);
        if (!before.equals(after)) {
            String restored = inverse.apply(after);
            if (!FormulaReferenceTransformer.canonicalizeFormulaReferences(before)
                    .equals(FormulaReferenceTransformer.canonicalizeFormulaReferences(restored))) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: " + participant
                        + " cannot be restored by the inverse structural operation");
            }
        }
        return after;
    }

    private static ObjectNode shiftTemplateFormulaAnchor(
            ObjectNode anchor,
            String targetSheetId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        String sheetId = SnapshotMutationSupport.text(anchor, "sheetId");
        if (!targetSheetId.equals(sheetId)) return anchor;
        String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW : SnapshotMutationSupport.MAX_COLUMN;
        int position = definedNameAnchorCoordinate(anchor, key, maximum);
        int shifted = shiftIndex(position, at, count, direction, axis);
        if (shifted < 0 || shifted > maximum) throw ServiceException.validation("Structural mutation removes cell-style-template formula anchor");
        ObjectNode result = anchor.deepCopy();
        result.put(key, shifted);
        return result;
    }

    private static ObjectNode shiftTemplateFormulaAnchor(
            ObjectNode anchor,
            String targetSheetId,
            FormulaReferenceTransformer.Range selection,
            FormulaReferenceTransformer.Axis axis,
            FormulaReferenceTransformer.Direction direction
    ) {
        String sheetId = SnapshotMutationSupport.text(anchor, "sheetId");
        if (!targetSheetId.equals(sheetId)) return anchor;
        int row = definedNameAnchorCoordinate(anchor, "row", SnapshotMutationSupport.MAX_ROW);
        int column = definedNameAnchorCoordinate(anchor, "column", SnapshotMutationSupport.MAX_COLUMN);
        int[] mapped = FormulaReferenceTransformer.remapCellShiftCoordinate(row, column, selection, axis, direction);
        if (mapped == null) throw ServiceException.validation("Structural mutation removes cell-style-template formula anchor");
        ObjectNode result = anchor.deepCopy();
        result.put("row", mapped[0]);
        result.put("column", mapped[1]);
        return result;
    }

    private static ObjectNode moveTemplateFormulaAnchor(
            ObjectNode anchor,
            String targetSheetId,
            RangeRef source,
            int rowDelta,
            int columnDelta
    ) {
        if (!targetSheetId.equals(SnapshotMutationSupport.text(anchor, "sheetId"))) return anchor;
        int row = definedNameAnchorCoordinate(anchor, "row", SnapshotMutationSupport.MAX_ROW);
        int column = definedNameAnchorCoordinate(anchor, "column", SnapshotMutationSupport.MAX_COLUMN);
        if (!contains(source, row, column)) return anchor;
        ObjectNode result = anchor.deepCopy();
        result.put("row", row + rowDelta);
        result.put("column", column + columnDelta);
        return result;
    }

    private static void invalidateFormulaCaches(ObjectNode root) {
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = requireObject(raw, "Sheet");
            forEachFormulaCell(sheet, cell -> cell.remove("formulaValue"));
        }
    }

    private static boolean containsNumber(ArrayNode values, int number) {
        for (JsonNode value : values) if (value.isIntegralNumber() && value.intValue() == number) return true;
        return false;
    }

    private static List<CellEntry> cellsInRange(ObjectNode sheet, RangeRef range) {
        List<CellEntry> entries = new ArrayList<>();
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        cells.fields().forEachRemaining(row -> {
            int rowIndex = integerKey(row.getKey(), SnapshotMutationSupport.MAX_ROW, "Cell row");
            if (rowIndex < range.startRow() || rowIndex > range.endRow()) return;
            if (!row.getValue().isObject()) throw ServiceException.validation("Cell row must be an object");
            ((ObjectNode) row.getValue()).fields().forEachRemaining(column -> {
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (!column.getValue().isObject()) throw ServiceException.validation("Cell payload must be an object");
                if (contains(range, rowIndex, columnIndex)) entries.add(new CellEntry(rowIndex, columnIndex, ((ObjectNode) column.getValue()).deepCopy()));
            });
        });
        return entries;
    }

    private static int[] validatePermutation(RangeRef range, ArrayNode sourceRows) {
        int[] rows = new int[sourceRows.size()];
        boolean[] seen = new boolean[rows.length];
        for (int index = 0; index < sourceRows.size(); index++) {
            JsonNode raw = sourceRows.get(index);
            if (!raw.isIntegralNumber() || !raw.canConvertToInt()) throw ServiceException.validation("Row permutation entry is invalid");
            int source = raw.intValue();
            if (source < range.startRow() || source > range.endRow()) throw ServiceException.validation("Row permutation source is outside range");
            int offset = source - range.startRow();
            if (seen[offset]) throw ServiceException.validation("Row permutation contains a duplicate source row");
            seen[offset] = true;
            rows[index] = source;
        }
        for (boolean value : seen) if (!value) throw ServiceException.validation("Row permutation must contain every row");
        return rows;
    }

    private static void validatePermutationPreservation(ObjectNode sheet, RangeRef range) {
        for (JsonNode merge : SnapshotMutationSupport.array(sheet, "merges")) {
            JsonNode mergeRange = merge.get("range");
            if (mergeRange != null && range.sheetId().equals(mergeRange.path("sheetId").asText())) {
                boolean intersects = mergeRange.path("startRow").asInt() <= range.endRow() && mergeRange.path("endRow").asInt() >= range.startRow()
                        && mergeRange.path("startColumn").asInt() <= range.endColumn() && mergeRange.path("endColumn").asInt() >= range.startColumn();
                if (intersects && !(mergeRange.path("startRow").asInt() >= range.startRow() && mergeRange.path("endRow").asInt() <= range.endRow())) {
                    throw ServiceException.validation("Row permutation partially intersects a merged range");
                }
            }
        }
        for (JsonNode table : SnapshotMutationSupport.array(sheet, "sheetTables")) {
            JsonNode tableRange = table.get("range");
            if (tableRange == null || !range.sheetId().equals(tableRange.path("sheetId").asText())) continue;
            boolean intersects = tableRange.path("startRow").asInt() <= range.endRow() && tableRange.path("endRow").asInt() >= range.startRow();
            boolean completeTable = tableRange.path("startRow").asInt() == range.startRow() && tableRange.path("endRow").asInt() == range.endRow();
            if (intersects && !completeTable && !isTableBodyPermutation((ObjectNode) table, range)) {
                throw ServiceException.validation("Row permutation requires the complete sheet table range or its data body");
            }
        }
        JsonNode outline = sheet.get("outline");
        if (outline != null && outline.isObject()) {
            for (JsonNode group : ((ObjectNode) outline).path("groups")) {
                if (!"row".equals(group.path("axis").asText())) continue;
                boolean intersects = group.path("start").asInt() <= range.endRow() && group.path("end").asInt() >= range.startRow();
                if (intersects && !(group.path("start").asInt() >= range.startRow() && group.path("end").asInt() <= range.endRow())) {
                    throw ServiceException.validation("Row permutation partially intersects an outline group");
                }
            }
        }
    }

    private static void remapPermutedCells(
            ObjectNode sheet,
            RangeRef range,
            int[] targetRowsBySource,
            List<StructuralPatch.FormulaOwnerDelta> formulaOwnerDeltas
    ) {
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        List<CellEntry> entries = new ArrayList<>();
        for (int row = range.startRow(); row <= range.endRow(); row++) {
            ObjectNode current = SnapshotMutationSupport.cellRow(cells, row, false);
            if (current == null) continue;
            int sourceRow = row;
            current.fields().forEachRemaining(column -> {
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (columnIndex < range.startColumn() || columnIndex > range.endColumn()) return;
                ObjectNode cell = requireObject(column.getValue(), "Cell").deepCopy();
                int targetRow = targetRowsBySource[sourceRow - range.startRow()];
                if (targetRow != sourceRow) {
                    StructuralPatch.FormulaOwnerState before = formulaOwnerState(cell);
                    remapPermutedFormulaOwner(cell, targetRow - sourceRow, range.sheetId(), sourceRow, columnIndex);
                    if (before.formula() != null || before.sourceFormula() != null || before.barcodeFormula() != null) {
                        formulaOwnerDeltas.add(new StructuralPatch.FormulaOwnerDelta(
                                "formula-cell",
                                new StructuralPatch.CellAddress(range.sheetId(), sourceRow, columnIndex),
                                new StructuralPatch.CellAddress(range.sheetId(), targetRow, columnIndex),
                                before,
                                formulaOwnerState(cell)));
                    }
                }
                entries.add(new CellEntry(sourceRow, columnIndex, cell));
            });
        }
        for (CellEntry entry : entries) {
            ObjectNode current = SnapshotMutationSupport.cellRow(cells, entry.row(), false);
            if (current == null) throw ServiceException.validation("Row permutation cell owner disappeared before remap");
            current.remove(Integer.toString(entry.column()));
            if (current.isEmpty()) cells.remove(Integer.toString(entry.row()));
        }
        for (CellEntry entry : entries) {
            int target = targetRowsBySource[entry.row() - range.startRow()];
            SnapshotMutationSupport.putCell(sheet, new SnapshotMutationSupport.CellCoordinate(target, entry.column()), entry.cell());
        }
    }

    private static void rejectMovedFormulaGroups(ObjectNode sheet, RangeRef range, int[] targetRowsBySource) {
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        for (int sourceRow = range.startRow(); sourceRow <= range.endRow(); sourceRow++) {
            ObjectNode row = SnapshotMutationSupport.cellRow(cells, sourceRow, false);
            if (row == null) continue;
            for (java.util.Iterator<java.util.Map.Entry<String, JsonNode>> columns = row.fields(); columns.hasNext();) {
                java.util.Map.Entry<String, JsonNode> column = columns.next();
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (columnIndex < range.startColumn() || columnIndex > range.endColumn()) continue;
                if (targetRowsBySource[sourceRow - range.startRow()] != sourceRow
                        && hasFormulaGroupMetadata(requireObject(column.getValue(), "Cell"))) {
                    throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot remap formula-group metadata at "
                            + range.sheetId() + "!" + sourceRow + ":" + columnIndex);
                }
            }
        }
    }

    private static boolean hasFormulaGroupMetadata(ObjectNode cell) {
        JsonNode metadata = cell.get("formulaMetadata");
        if (metadata == null || metadata.isNull()) return false;
        if (!metadata.isObject()) throw ServiceException.validation("Cell formulaMetadata must be an object");
        return metadata.path("preservedOnly").asBoolean(false)
                || !"normal".equals(metadata.path("kind").asText())
                || (metadata.has("range") && !metadata.path("range").isNull());
    }

    private static boolean rewritesOnlyPreservedDataTableSource(
            ObjectNode cell,
            StructuralPatch.FormulaOwnerState before,
            String formula,
            String sourceFormula,
            String barcodeFormula
    ) {
        JsonNode metadata = cell.get("formulaMetadata");
        return metadata != null && metadata.isObject()
                && "dataTable".equals(metadata.path("kind").asText())
                && metadata.path("preservedOnly").asBoolean(false)
                && before.formula() == null && formula == null
                && before.sourceFormula() != null && !Objects.equals(sourceFormula, before.sourceFormula())
                && Objects.equals(barcodeFormula, before.barcodeFormula());
    }

    private static void remapPermutedFormulaOwner(ObjectNode cell, int rowDelta, String sheetId, int row, int column) {
        remapFormulaField(cell, "formula", rowDelta, sheetId, row, column);
        JsonNode rawMetadata = cell.get("formulaMetadata");
        if (rawMetadata != null && rawMetadata.isObject()) {
            remapFormulaField((ObjectNode) rawMetadata, "sourceFormula", rowDelta, sheetId, row, column);
        }
        JsonNode rawPresentation = cell.get("presentation");
        if (rawPresentation != null && rawPresentation.isObject() && "barcode".equals(rawPresentation.path("kind").asText())) {
            JsonNode rawSource = rawPresentation.get("source");
            if (rawSource != null && rawSource.isObject() && "formula".equals(rawSource.path("kind").asText())) {
                remapFormulaField((ObjectNode) rawSource, "formula", rowDelta, sheetId, row, column);
            }
        }
    }

    private static void remapFormulaField(ObjectNode owner, String field, int rowDelta, String sheetId, int row, int column) {
        JsonNode rawFormula = owner.get(field);
        if (rawFormula == null || !rawFormula.isTextual()) return;
        String formula = rawFormula.asText();
        try {
            FormulaReferenceTransformer.assertRowOffsetSupported(formula);
            owner.put(field, FormulaReferenceTransformer.offsetForPermutation(formula, rowDelta));
        } catch (ServiceException error) {
            if (!"SERVICE_UNAVAILABLE".equals(error.code())) {
                throw new ServiceException("SERVICE_UNAVAILABLE", 503,
                        "UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot rewrite formula owner " + sheetId + "!" + row + ":" + column,
                        error);
            }
            throw error;
        } catch (RuntimeException error) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot parse formula owner " + sheetId + "!" + row + ":" + column);
        }
    }

    private static void remapPermutationMetadata(ObjectNode root, ObjectNode sheet, RangeRef range, RangeRef metadataScope, int[] targetRowsBySource) {
        SnapshotMutationSupport.remapReviewCoordinates(sheet, coordinate -> contains(range, coordinate.row(), coordinate.column())
                ? new SnapshotMutationSupport.CellCoordinate(remapRow(coordinate.row(), range, targetRowsBySource), coordinate.column())
                : coordinate);
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "hyperlinks")) remapCellOwner(requireObject(raw, "Hyperlink"), range, targetRowsBySource);
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "drawings")) remapDrawingRows(requireObject(raw, "Drawing"), range, targetRowsBySource);
        for (JsonNode rawOwner : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawOwner, "Worksheet");
            boolean ownsTargetSheet = range.sheetId().equals(owner.path("id").asText());
            for (JsonNode raw : existingArray(owner, "sparklines")) {
                ObjectNode sparkline = requireObject(raw, "Sparkline");
                ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
                if (ownsTargetSheet && contains(range, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) anchor.put("row", remapRow(anchor.path("row").asInt(), range, targetRowsBySource));
                writeSingleRange(sparkline.get("sourceRange"), range, targetRowsBySource, "sparkline source");
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "spillRanges")) {
            ObjectNode spill = requireObject(raw, "Spill range");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(spill, "anchor");
            if (contains(range, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) anchor.put("row", remapRow(anchor.path("row").asInt(), range, targetRowsBySource));
            writeSingleRange(spill.get("range"), range, targetRowsBySource, "spill range");
        }
        remapPermutationRuleFormulaOwners(sheet, range.sheetId(), metadataScope, targetRowsBySource);
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            for (JsonNode rule : SnapshotMutationSupport.array(sheet, property)) {
                ArrayNode ranges = (ArrayNode) requireObject(rule, "Range rule").path("ranges");
                replaceRanges(ranges, metadataScope, targetRowsBySource);
            }
        }
        SheetRuleLifecycle.transformValidationListSources(root, sheet,
                candidate -> remapRangeExact(rangeNode(candidate), metadataScope, targetRowsBySource));
        for (JsonNode rawRegion : SnapshotMutationSupport.array(sheet, "dataRegions")) {
            ObjectNode region = requireObject(rawRegion, "Data region");
            writeSingleRange(region.get("range"), range, targetRowsBySource, "data region");
            int headerRow = integer(region.get("headerRow"), "Data region header row", SnapshotMutationSupport.MAX_ROW);
            if (headerRow >= range.startRow() && headerRow <= range.endRow()) {
                region.put("headerRow", remapRow(headerRow, range, targetRowsBySource));
            }
        }
        JsonNode namesProjectionRaw = root.get("definedNames");
        ObjectNode namesProjection = namesProjectionRaw != null && namesProjectionRaw.isObject() ? (ObjectNode) namesProjectionRaw : null;
        remapPermutationDefinedNames(existingArray(root, "definedNameModels"), namesProjection, range.sheetId(), metadataScope, targetRowsBySource);
        remapPermutationCellStyleTemplates(existingArray(root, "cellStyleTemplates"), range.sheetId(), metadataScope, targetRowsBySource);
        JsonNode filter = sheet.get("autoFilter");
        if (filter != null && filter.isObject()) writeSingleRange(filter.get("range"), range, targetRowsBySource, "auto filter");
        for (JsonNode rawTable : SnapshotMutationSupport.array(sheet, "sheetTables")) {
            ObjectNode table = requireObject(rawTable, "Sheet table");
            if (!isTableBodyPermutation(table, range)) {
                writeSingleRange(table.get("range"), range, targetRowsBySource, "sheet table");
                if (table.has("autoFilter")) writeSingleRange(table.get("autoFilter").path("range"), range, targetRowsBySource, "sheet table filter");
            }
        }
        for (JsonNode rawOwner : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawOwner, "Worksheet");
            for (JsonNode rawPivot : existingArray(owner, "pivots")) {
                ObjectNode pivot = requireObject(rawPivot, "Pivot");
                SnapshotMutationSupport.validateKnownKeys(pivot, Set.of("schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "presentation", "nativeMetadata"), "Pivot");
                PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, source -> writeSingleRange(source, range, targetRowsBySource, "pivot source"));
                ObjectNode target = PivotMutationDescriptor.requiredTarget(pivot);
                ObjectNode anchor = PivotMutationDescriptor.requiredAnchor(target);
                if (range.sheetId().equals(target.path("sheetId").asText()) && contains(range, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) {
                    anchor.put("row", remapRow(anchor.path("row").asInt(), range, targetRowsBySource));
                }
            }
        }
        for (JsonNode rawTable : workbookTables(root)) {
            ObjectNode table = requireObject(rawTable, "Workbook table");
            JsonNode sourceRange = table.get("sourceRange");
            if (sourceRange == null || sourceRange.isNull()) continue;
            if (rangesIntersect(SnapshotMutationSupport.range(root, sourceRange), range)) {
                writeSingleRange(sourceRange, range, targetRowsBySource, "workbook table source");
            }
        }
        for (JsonNode rawSource : existingDataModelArray(root, "sources")) {
            ObjectNode source = requireObject(rawSource, "Data source");
            JsonNode sourceRange = source.get("sourceRange");
            if (sourceRange == null || sourceRange.isNull()) continue;
            if (rangesIntersect(SnapshotMutationSupport.range(root, sourceRange), range)) {
                writeSingleRange(sourceRange, range, targetRowsBySource, "data source range");
            }
        }
        for (JsonNode merge : SnapshotMutationSupport.array(sheet, "merges")) {
            ObjectNode object = requireObject(merge, "Merge");
            writeSingleRange(object.get("range"), range, targetRowsBySource, "merge");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(object, "anchor");
            if (contains(range, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) anchor.put("row", remapRow(anchor.path("row").asInt(), range, targetRowsBySource));
        }
        JsonNode outline = sheet.get("outline");
        if (outline != null && outline.isObject()) {
            for (JsonNode group : ((ObjectNode) outline).path("groups")) {
                if (!"row".equals(group.path("axis").asText())) continue;
                int start = group.path("start").asInt(-1);
                int end = group.path("end").asInt(-1);
                if (start > range.endRow() || end < range.startRow()) continue;
                RangeRef groupRange = new RangeRef(range.sheetId(), start, end, range.startColumn(), range.endColumn());
                List<RangeRef> mapped = remapRangeExact(rangeNode(groupRange), range, targetRowsBySource);
                if (mapped.size() != 1) throw ServiceException.validation("Row permutation cannot exactly remap an outline group");
                ((ObjectNode) group).put("start", mapped.get(0).startRow()).put("end", mapped.get(0).endRow());
            }
        }
        for (JsonNode rule : SnapshotMutationSupport.array(sheet, "protectionRules")) if (rule.has("range")) writeSingleRange(rule.get("range"), metadataScope, targetRowsBySource, "protection rule");
        JsonNode bandedRaw = sheet.get("bandedRule");
        if (bandedRaw != null && !bandedRaw.isNull()) {
            ObjectNode banded = requireObject(bandedRaw, "Banded rule");
            writeSingleRange(banded.get("range"), range, targetRowsBySource, "banded rule");
        }
        remapPermutationDrawingPayloads(root, range, targetRowsBySource);
    }

    private static void validatePermutationDrawingPayloads(ObjectNode root, RangeRef range, int[] targetRowsBySource) {
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawSheet, "Sheet");
            JsonNode rawPayloads = owner.get("drawingPayloads");
            if (rawPayloads == null || rawPayloads.isNull()) continue;
            if (!rawPayloads.isObject()) throw ServiceException.validation("drawingPayloads must be an object");
            var fields = rawPayloads.fields();
            while (fields.hasNext()) {
                var entry = fields.next();
                ObjectNode source = requireObject(entry.getValue(), "Drawing payload");
                if (!drawingPayloadIntersectsPermutation(root, source, range)) continue;
                ObjectNode payload = source.deepCopy();
                remapPermutationDrawingPayload(payload, entry.getKey(), range, targetRowsBySource);
            }
        }
    }

    private static boolean drawingPayloadIntersectsPermutation(ObjectNode root, ObjectNode payload, RangeRef range) {
        String kind = payload.path("kind").asText();
        if ("camera".equals(kind) || "screenshot".equals(kind)) {
            return permutationRangeIntersects(root, requireObject(payload.get("sourceRange"), kind + " source range"), range);
        }
        if ("form-control".equals(kind)) {
            JsonNode rawLink = payload.get("cellLink");
            if (rawLink != null && !rawLink.isNull()) {
                ObjectNode link = requireObject(rawLink, "Form-control cell link");
                String targetSheetId = SnapshotMutationSupport.text(link, "sheetId");
                JsonNode rowNode = link.get("row");
                JsonNode columnNode = link.get("column");
                if (rowNode == null || !rowNode.canConvertToInt() || columnNode == null || !columnNode.canConvertToInt()) {
                    throw ServiceException.validation("Form-control cell link is invalid");
                }
                int row = rowNode.asInt(-1);
                int column = columnNode.asInt(-1);
                if (row < 0 || row > SnapshotMutationSupport.MAX_ROW || column < 0 || column > SnapshotMutationSupport.MAX_COLUMN) {
                    throw ServiceException.validation("Form-control cell link is outside worksheet bounds");
                }
                if (range.sheetId().equals(targetSheetId) && contains(range, row, column)) {
                    return true;
                }
            }
            JsonNode inputRange = payload.get("inputRange");
            return inputRange != null && !inputRange.isNull()
                    && permutationRangeIntersects(root, requireObject(inputRange, "Form-control input range"), range);
        }
        if (!"chart".equals(kind)) return false;

        ObjectNode source = requireObject(payload.get("source"), "Chart source");
        String sourceKind = source.path("kind").asText();
        if ("worksheet-ranges".equals(sourceKind)) {
            JsonNode ranges = source.get("ranges");
            if (ranges == null || !ranges.isArray() || ranges.isEmpty()) throw ServiceException.validation("Chart worksheet source ranges are invalid");
            for (JsonNode rawRange : ranges) if (permutationRangeIntersects(root, requireObject(rawRange, "Chart worksheet source range"), range)) return true;
        } else if ("report-range".equals(sourceKind)
                && permutationRangeIntersects(root, requireObject(source.get("range"), "Chart report source range"), range)) {
            return true;
        } else if (!Set.of("report-range", "pivot", "table").contains(sourceKind)) {
            throw ServiceException.validation("Chart source kind is invalid: " + sourceKind);
        }
        JsonNode categoryRange = payload.get("categoryRange");
        if (categoryRange != null && !categoryRange.isNull()
                && permutationRangeIntersects(root, requireObject(categoryRange, "Chart category range"), range)) return true;
        JsonNode seriesRaw = payload.get("series");
        if (seriesRaw == null || seriesRaw.isNull()) return false;
        if (!seriesRaw.isArray()) throw ServiceException.validation("Chart series collection is invalid");
        for (JsonNode rawSeries : seriesRaw) {
            ObjectNode series = requireObject(rawSeries, "Chart series");
            for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) {
                JsonNode value = series.get(field);
                if (value != null && !value.isNull()
                        && permutationRangeIntersects(root, requireObject(value, "Chart series " + field), range)) return true;
            }
            for (String parent : List.of("stockRoles", "dataLabels", "errorBars")) {
                JsonNode nestedRaw = series.get(parent);
                if (nestedRaw == null || nestedRaw.isNull()) continue;
                ObjectNode nested = requireObject(nestedRaw, "Chart " + parent);
                List<String> fields = switch (parent) {
                    case "stockRoles" -> List.of("open", "high", "low", "close", "volume");
                    case "dataLabels" -> List.of("valuesFromCells");
                    default -> List.of("plusRange", "minusRange");
                };
                for (String field : fields) {
                    JsonNode value = nested.get(field);
                    if (value != null && !value.isNull()
                            && permutationRangeIntersects(root, requireObject(value, "Chart " + parent + " range"), range)) return true;
                }
            }
        }
        return false;
    }

    private static boolean permutationRangeIntersects(ObjectNode root, ObjectNode candidate, RangeRef range) {
        return rangesIntersect(SnapshotMutationSupport.range(root, candidate), range);
    }

    private static void remapPermutationDrawingPayloads(ObjectNode root, RangeRef range, int[] rowMap) {
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawSheet, "Sheet");
            JsonNode rawPayloads = owner.get("drawingPayloads");
            if (rawPayloads == null || rawPayloads.isNull()) continue;
            if (!rawPayloads.isObject()) throw ServiceException.validation("drawingPayloads must be an object");
            var fields = rawPayloads.fields();
            while (fields.hasNext()) {
                var entry = fields.next();
                remapPermutationDrawingPayload(requireObject(entry.getValue(), "Drawing payload"), entry.getKey(), range, rowMap);
            }
        }
    }

    private static void remapPermutationDrawingPayload(ObjectNode payload, String payloadId, RangeRef range, int[] rowMap) {
        String kind = payload.path("kind").asText();
        if ("camera".equals(kind) || "screenshot".equals(kind)) {
            writeSingleRange(requireObject(payload.get("sourceRange"), kind + " source range"), range, rowMap, kind + " source range");
            return;
        }
        if ("form-control".equals(kind)) {
            JsonNode rawLink = payload.get("cellLink");
            if (rawLink != null && !rawLink.isNull()) {
                ObjectNode link = requireObject(rawLink, "Form-control cell link");
                String targetSheetId = SnapshotMutationSupport.text(link, "sheetId");
                JsonNode rowNode = link.get("row");
                JsonNode columnNode = link.get("column");
                if (rowNode == null || !rowNode.canConvertToInt() || columnNode == null || !columnNode.canConvertToInt()) {
                    throw ServiceException.validation("Form-control cell link is invalid");
                }
                int row = rowNode.asInt(-1);
                int column = columnNode.asInt(-1);
                if (row < 0 || row > SnapshotMutationSupport.MAX_ROW || column < 0 || column > SnapshotMutationSupport.MAX_COLUMN) {
                    throw ServiceException.validation("Form-control cell link is outside worksheet bounds");
                }
                if (range.sheetId().equals(targetSheetId) && contains(range, row, column)) link.put("row", remapRow(row, range, rowMap));
            }
            JsonNode inputRange = payload.get("inputRange");
            if (inputRange != null && !inputRange.isNull()) {
                writeSingleRange(requireObject(inputRange, "Form-control input range"), range, rowMap, "form-control input range");
            }
            return;
        }
        if (!"chart".equals(kind)) return;

        ObjectNode source = requireObject(payload.get("source"), "Chart source");
        String sourceKind = source.path("kind").asText();
        if ("worksheet-ranges".equals(sourceKind)) {
            JsonNode ranges = source.get("ranges");
            if (ranges == null || !ranges.isArray() || ranges.isEmpty()) throw ServiceException.validation("Chart worksheet source ranges are invalid");
            for (JsonNode rawRange : ranges) writeSingleRange(requireObject(rawRange, "Chart worksheet source range"), range, rowMap, "chart " + payloadId + " source range");
        } else if ("report-range".equals(sourceKind)) {
            writeSingleRange(requireObject(source.get("range"), "Chart report source range"), range, rowMap, "chart " + payloadId + " report range");
        } else if (!Set.of("pivot", "table").contains(sourceKind)) {
            throw ServiceException.validation("Chart source kind is invalid: " + sourceKind);
        }
        JsonNode categoryRange = payload.get("categoryRange");
        if (categoryRange != null && !categoryRange.isNull()) writeSingleRange(requireObject(categoryRange, "Chart category range"), range, rowMap, "chart category range");
        JsonNode seriesRaw = payload.get("series");
        if (seriesRaw != null && !seriesRaw.isNull() && !seriesRaw.isArray()) throw ServiceException.validation("Chart series collection is invalid");
        if (seriesRaw == null || seriesRaw.isNull()) return;
        for (JsonNode rawSeries : seriesRaw) {
            ObjectNode series = requireObject(rawSeries, "Chart series");
            for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) {
                JsonNode value = series.get(field);
                if (value != null && !value.isNull()) writeSingleRange(requireObject(value, "Chart series " + field), range, rowMap, "chart series " + field);
            }
            JsonNode stockRolesRaw = series.get("stockRoles");
            if (stockRolesRaw != null && !stockRolesRaw.isNull()) {
                ObjectNode stockRoles = requireObject(stockRolesRaw, "Chart stock roles");
                for (String field : List.of("open", "high", "low", "close", "volume")) {
                    JsonNode value = stockRoles.get(field);
                    if (value != null && !value.isNull()) writeSingleRange(requireObject(value, "Chart stock-role range"), range, rowMap, "chart stock-role range");
                }
            }
            JsonNode labelsRaw = series.get("dataLabels");
            if (labelsRaw != null && !labelsRaw.isNull()) {
                ObjectNode labels = requireObject(labelsRaw, "Chart data labels");
                JsonNode values = labels.get("valuesFromCells");
                if (values != null && !values.isNull()) writeSingleRange(requireObject(values, "Chart data-label range"), range, rowMap, "chart data-label range");
            }
            JsonNode errorBarsRaw = series.get("errorBars");
            if (errorBarsRaw != null && !errorBarsRaw.isNull()) {
                ObjectNode errorBars = requireObject(errorBarsRaw, "Chart error bars");
                for (String field : List.of("plusRange", "minusRange")) {
                    JsonNode value = errorBars.get(field);
                    if (value != null && !value.isNull()) writeSingleRange(requireObject(value, "Chart error-bar range"), range, rowMap, "chart error-bar range");
                }
            }
        }
    }

    private static void remapCellOwner(ObjectNode owner, RangeRef range, int[] rowMap) {
        int row = owner.path("row").asInt(-1);
        int column = owner.path("column").asInt(-1);
        if (contains(range, row, column)) owner.put("row", remapRow(row, range, rowMap));
    }

    private static void remapDrawingRows(ObjectNode drawing, RangeRef range, int[] rowMap) {
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
        if ("absolute".equals(anchor.path("kind").asText())) return;
        int row = anchor.path("row").asInt(-1);
        int column = anchor.path("column").asInt(-1);
        int endRow = anchor.has("endRow") ? anchor.path("endRow").asInt(-1) : row;
        int endColumn = anchor.has("endColumn") ? anchor.path("endColumn").asInt(-1) : column;
        boolean startInside = contains(range, row, column);
        boolean endInside = contains(range, endRow, endColumn);
        if (!startInside && !endInside) return;
        if (!startInside || !endInside) throw ServiceException.validation("Row permutation cannot exactly remap a drawing anchor");
        anchor.put("row", remapRow(row, range, rowMap));
        if (anchor.has("endRow")) anchor.put("endRow", remapRow(endRow, range, rowMap));
    }

    private static void replaceRanges(ArrayNode ranges, RangeRef range, int[] rowMap) {
        List<RangeRef> next = new ArrayList<>();
        for (JsonNode raw : ranges) next.addAll(remapRangeExact(raw, range, rowMap));
        ranges.removeAll();
        for (RangeRef value : next) ranges.add(rangeNode(value));
    }

    private static void writeSingleRange(JsonNode raw, RangeRef range, int[] rowMap, String owner) {
        if (raw == null || !raw.isObject()) return;
        List<RangeRef> segments = remapRangeExact(raw, range, rowMap);
        if (segments.size() != 1) throw ServiceException.validation("Row permutation cannot exactly remap " + owner + " into one range");
        ((ObjectNode) raw).setAll(rangeNode(segments.get(0)));
    }

    private static ObjectNode rangeNode(RangeRef value) {
        ObjectNode node = JsonNodeFactory.instance.objectNode();
        node.put("sheetId", value.sheetId());
        node.put("startRow", value.startRow());
        node.put("endRow", value.endRow());
        node.put("startColumn", value.startColumn());
        node.put("endColumn", value.endColumn());
        return node;
    }

    private static List<RangeRef> remapRangeExact(JsonNode raw, RangeRef range, int[] rowMap) {
        if (raw == null || !raw.isObject()) return List.of();
        RangeRef owner = new RangeRef(raw.path("sheetId").asText(), raw.path("startRow").asInt(), raw.path("endRow").asInt(), raw.path("startColumn").asInt(), raw.path("endColumn").asInt());
        if (!rangesIntersect(owner, range)) return List.of(owner);
        int firstRow = Math.max(owner.startRow(), range.startRow());
        int lastRow = Math.min(owner.endRow(), range.endRow());
        int firstColumn = Math.max(owner.startColumn(), range.startColumn());
        int lastColumn = Math.min(owner.endColumn(), range.endColumn());
        List<RangeRef> next = new ArrayList<>();
        if (owner.startRow() < firstRow) next.add(new RangeRef(owner.sheetId(), owner.startRow(), firstRow - 1, owner.startColumn(), owner.endColumn()));
        if (lastRow < owner.endRow()) next.add(new RangeRef(owner.sheetId(), lastRow + 1, owner.endRow(), owner.startColumn(), owner.endColumn()));
        if (owner.startColumn() < firstColumn) next.add(new RangeRef(owner.sheetId(), firstRow, lastRow, owner.startColumn(), firstColumn - 1));
        if (lastColumn < owner.endColumn()) next.add(new RangeRef(owner.sheetId(), firstRow, lastRow, lastColumn + 1, owner.endColumn()));
        long area = (long) (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
        if (area > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Row permutation metadata range is too large for exact remapping");
        int[] targetRows = new int[lastRow - firstRow + 1];
        for (int row = firstRow; row <= lastRow; row++) targetRows[row - firstRow] = remapRow(row, range, rowMap);
        java.util.Arrays.sort(targetRows);
        int index = 0;
        while (index < targetRows.length) {
            int start = targetRows[index];
            int end = start;
            while (index + 1 < targetRows.length && targetRows[index + 1] == end + 1) end = targetRows[++index];
            if (next.size() >= MAX_EXACT_RANGE_SEGMENTS) {
                throw ServiceException.validation("Row permutation metadata produces too many exact ranges");
            }
            next.add(new RangeRef(owner.sheetId(), start, end, firstColumn, lastColumn));
            index++;
        }
        return mergeExactSegments(next);
    }

    private static List<RangeRef> mergeExactSegments(List<RangeRef> segments) {
        List<RangeRef> result = new ArrayList<>(segments);
        boolean changed = true;
        while (changed) {
            changed = false;
            outer:
            for (int left = 0; left < result.size(); left++) {
                for (int right = left + 1; right < result.size(); right++) {
                    RangeRef a = result.get(left);
                    RangeRef b = result.get(right);
                    boolean sameRows = a.startRow() == b.startRow() && a.endRow() == b.endRow();
                    boolean sameColumns = a.startColumn() == b.startColumn() && a.endColumn() == b.endColumn();
                    boolean adjacentColumns = a.endColumn() + 1 == b.startColumn() || b.endColumn() + 1 == a.startColumn();
                    boolean adjacentRows = a.endRow() + 1 == b.startRow() || b.endRow() + 1 == a.startRow();
                    if (!((sameRows && adjacentColumns) || (sameColumns && adjacentRows))) continue;
                    result.set(left, new RangeRef(a.sheetId(), Math.min(a.startRow(), b.startRow()), Math.max(a.endRow(), b.endRow()), Math.min(a.startColumn(), b.startColumn()), Math.max(a.endColumn(), b.endColumn())));
                    result.remove(right);
                    changed = true;
                    break outer;
                }
            }
        }
        return result;
    }

    private static boolean rangesIntersect(RangeRef a, RangeRef b) {
        return a.sheetId().equals(b.sheetId()) && a.startRow() <= b.endRow() && b.startRow() <= a.endRow() && a.startColumn() <= b.endColumn() && b.startColumn() <= a.endColumn();
    }

    private static void validatePermutationMetadataExact(ObjectNode root, ObjectNode sheet, RangeRef range, RangeRef metadataScope, int[] targetRowsBySource) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "drawings")) validateDrawingExact(requireObject(raw, "Drawing"), range);
        for (JsonNode rawRegion : SnapshotMutationSupport.array(sheet, "dataRegions")) {
            ObjectNode region = requireObject(rawRegion, "Data region");
            List<RangeRef> nextRanges = remapRangeExact(region.get("range"), range, targetRowsBySource);
            if (nextRanges.size() != 1) {
                throw ServiceException.validation("Row permutation cannot exactly remap data region " + region.path("id").asText());
            }
            int headerRow = integer(region.get("headerRow"), "Data region header row", SnapshotMutationSupport.MAX_ROW);
            int nextHeaderRow = headerRow >= range.startRow() && headerRow <= range.endRow()
                    ? remapRow(headerRow, range, targetRowsBySource) : headerRow;
            RangeRef nextRange = nextRanges.getFirst();
            if (!sheet.path("id").asText().equals(nextRange.sheetId())
                    || nextHeaderRow < nextRange.startRow() || nextHeaderRow > nextRange.endRow()) {
                throw ServiceException.validation("Row permutation would separate data-region header from its owner "
                        + region.path("id").asText());
            }
        }
        for (JsonNode rawOwner : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawOwner, "Worksheet");
            for (JsonNode raw : existingArray(owner, "sparklines")) requireSingleRange(requireObject(raw, "Sparkline").get("sourceRange"), range, targetRowsBySource, "sparkline source");
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "spillRanges")) requireSingleRange(requireObject(raw, "Spill range").get("range"), range, targetRowsBySource, "spill range");
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "conditionalFormats")) for (JsonNode item : requireObject(raw, "Conditional format").path("ranges")) remapRangeExact(item, metadataScope, targetRowsBySource);
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "dataValidations")) for (JsonNode item : requireObject(raw, "Data validation").path("ranges")) remapRangeExact(item, metadataScope, targetRowsBySource);
        SheetRuleLifecycle.validateStructuralFields(root, sheet, range.sheetId(), metadataScope,
                candidate -> remapRangeExact(rangeNode(candidate), metadataScope, targetRowsBySource));
        JsonNode filter = sheet.get("autoFilter");
        if (filter != null && filter.isObject()) requireSingleRange(filter.get("range"), range, targetRowsBySource, "auto filter");
        for (JsonNode rawTable : SnapshotMutationSupport.array(sheet, "sheetTables")) {
            ObjectNode table = requireObject(rawTable, "Sheet table");
            if (!isTableBodyPermutation(table, range)) {
                requireSingleRange(table.get("range"), range, targetRowsBySource, "sheet table");
                if (table.has("autoFilter")) requireSingleRange(table.get("autoFilter").path("range"), range, targetRowsBySource, "sheet table filter");
            }
        }
        for (JsonNode rawOwner : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(rawOwner, "Worksheet");
            for (JsonNode raw : existingArray(owner, "pivots")) PivotMutationDescriptor.forEachWorksheetSourceRange(requireObject(raw, "Pivot"), source -> requireSingleRange(source, range, targetRowsBySource, "pivot source"));
        }
        for (JsonNode rawTable : workbookTables(root)) {
            ObjectNode table = requireObject(rawTable, "Workbook table");
            JsonNode sourceRange = table.get("sourceRange");
            if (sourceRange == null || sourceRange.isNull()) continue;
            if (rangesIntersect(SnapshotMutationSupport.range(root, sourceRange), range)) {
                requireSingleRange(sourceRange, range, targetRowsBySource, "workbook table source");
            }
        }
        for (JsonNode rawSource : existingDataModelArray(root, "sources")) {
            ObjectNode source = requireObject(rawSource, "Data source");
            JsonNode sourceRange = source.get("sourceRange");
            if (sourceRange == null || sourceRange.isNull()) continue;
            if (rangesIntersect(SnapshotMutationSupport.range(root, sourceRange), range)) {
                requireSingleRange(sourceRange, range, targetRowsBySource, "data source range");
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "merges")) requireSingleRange(requireObject(raw, "Merge").get("range"), range, targetRowsBySource, "merge");
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "protectionRules")) if (raw.has("range")) requireSingleRange(raw.get("range"), metadataScope, targetRowsBySource, "protection rule");
        JsonNode bandedRaw = sheet.get("bandedRule");
        if (bandedRaw != null && !bandedRaw.isNull()) {
            ObjectNode banded = requireObject(bandedRaw, "Banded rule");
            requireSingleRange(banded.get("range"), range, targetRowsBySource, "banded rule");
        }
        JsonNode outline = sheet.get("outline");
        if (outline != null && !outline.isNull()) {
            if (!outline.isObject() || !outline.path("groups").isArray()) throw ServiceException.validation("Worksheet outline groups must be an array");
            for (JsonNode rawGroup : outline.path("groups")) {
                ObjectNode group = requireObject(rawGroup, "Outline group");
                if (!"row".equals(group.path("axis").asText())) continue;
                int start = group.path("start").asInt(-1);
                int end = group.path("end").asInt(-1);
                if (start < 0 || end < start || end > SnapshotMutationSupport.MAX_ROW) throw ServiceException.validation("Outline group row bounds are invalid");
                if (start > range.endRow() || end < range.startRow()) continue;
                if (start < range.startRow() || end > range.endRow()) throw ServiceException.validation("Row permutation cannot partially intersect an outline group");
                RangeRef groupRange = new RangeRef(range.sheetId(), start, end, range.startColumn(), range.endColumn());
                if (remapRangeExact(rangeNode(groupRange), range, targetRowsBySource).size() != 1) {
                    throw ServiceException.validation("Row permutation cannot exactly remap an outline group");
                }
            }
        }
        validatePermutationDrawingPayloads(root, range, targetRowsBySource);
        validatePermutationFormulaOwners(root, sheet, range.sheetId(), metadataScope, targetRowsBySource);
    }

    private static void validatePermutationFormulaOwners(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef scope, int[] targetRowsBySource) {
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            ObjectNode stagedSheet = JsonNodeFactory.instance.objectNode();
            stagedSheet.set(property, SnapshotMutationSupport.array(sheet, property).deepCopy());
            remapPermutationRuleFormulaOwners(stagedSheet, sheetId, scope, targetRowsBySource);
        }
        ArrayNode stagedNames = existingArray(root, "definedNameModels").deepCopy();
        JsonNode namesRaw = root.get("definedNames");
        if (namesRaw != null && !namesRaw.isNull() && !namesRaw.isObject()) throw ServiceException.validation("definedNames must be an object");
        ObjectNode stagedProjection = namesRaw != null && namesRaw.isObject() ? ((ObjectNode) namesRaw).deepCopy() : null;
        remapPermutationDefinedNames(stagedNames, stagedProjection, sheetId, scope, targetRowsBySource);
        remapPermutationCellStyleTemplates(existingArray(root, "cellStyleTemplates").deepCopy(), sheetId, scope, targetRowsBySource);
    }

    private static ArrayNode existingArray(ObjectNode parent, String property) {
        JsonNode value = parent.get(property);
        if (value == null || value.isNull()) return JsonNodeFactory.instance.arrayNode();
        if (!value.isArray()) throw ServiceException.validation(property + " must be an array");
        return (ArrayNode) value;
    }

    private static ArrayNode existingDataModelArray(ObjectNode root, String property) {
        JsonNode value = root.get("dataModel");
        if (value == null || value.isNull()) return JsonNodeFactory.instance.arrayNode();
        ObjectNode dataModel = requireObject(value, "dataModel");
        return existingArray(dataModel, property);
    }

    private static void remapPermutationRuleFormulaOwners(ObjectNode sheet, String sheetId, RangeRef scope, int[] rowMap) {
        boolean changesRows = false;
        for (int index = 0; index < rowMap.length; index++) {
            if (rowMap[index] != scope.startRow() + index) {
                changesRows = true;
                break;
            }
        }
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            for (JsonNode raw : SnapshotMutationSupport.array(sheet, property)) {
                ObjectNode rule = requireObject(raw, property + " rule");
                ArrayNode ranges = SnapshotMutationSupport.requiredArray(rule, "ranges");
                if (ranges.isEmpty()) throw ServiceException.validation(property + " rule ranges must not be empty");
                JsonNode rawAnchor = rule.get("formulaAnchor");
                boolean explicitAnchor = rawAnchor != null && !rawAnchor.isNull();
                JsonNode source = explicitAnchor ? rawAnchor : ranges.get(0);
                if (source == null || !source.isObject()) throw ServiceException.validation("Sheet rule formula anchor is invalid");
                String anchorSheetId = explicitAnchor ? SnapshotMutationSupport.text((ObjectNode) source, "sheetId") : source.path("sheetId").asText();
                if (!sheetId.equals(anchorSheetId)) throw ServiceException.validation("Sheet rule formula anchor targets another sheet");
                JsonNode rowNode = explicitAnchor ? source.get("row") : source.get("startRow");
                JsonNode columnNode = explicitAnchor ? source.get("column") : source.get("startColumn");
                if (rowNode == null || !rowNode.canConvertToInt() || columnNode == null || !columnNode.canConvertToInt()) {
                    throw ServiceException.validation("Sheet rule formula anchor is invalid");
                }
                int row = rowNode.asInt(-1);
                int column = columnNode.asInt(-1);
                if (row < 0 || row > SnapshotMutationSupport.MAX_ROW || column < 0 || column > SnapshotMutationSupport.MAX_COLUMN) {
                    throw ServiceException.validation("Sheet rule formula anchor is outside worksheet bounds");
                }
                if (!contains(scope, row, column)) continue;
                int targetRow = remapRow(row, scope, rowMap);
                int rowDelta = targetRow - row;
                if (rowDelta != 0) {
                    String identity = property + " " + rule.path("id").asText("<unknown>");
                    rewriteRuleFormulas(rule, formula -> offsetPermutationFormula(formula, rowDelta, identity));
                    if (explicitAnchor) ((ObjectNode) rawAnchor).put("row", targetRow);
                }
                if (!explicitAnchor && hasPermutationFormulaOwner(property, rule) && changesRows) {
                    ObjectNode mappedAnchor = JsonNodeFactory.instance.objectNode();
                    mappedAnchor.put("sheetId", sheetId).put("row", targetRow).put("column", column);
                    rule.set("formulaAnchor", mappedAnchor);
                }
            }
        }
    }

    private static boolean hasPermutationFormulaOwner(String property, ObjectNode rule) {
        boolean formulaOperator = "formula".equals(rule.path("operator").asText());
        boolean customValidation = "dataValidations".equals(property) && "custom".equals(rule.path("type").asText());
        for (String field : List.of("value1", "value2", "formula1", "formula2")) {
            JsonNode formula = rule.get(field);
            if (formula == null || !formula.isTextual() || formula.asText().isEmpty()) continue;
            if (formula.asText().stripLeading().startsWith("=")
                    || ("value1".equals(field) && formulaOperator)
                    || ("formula1".equals(field) && (formulaOperator || customValidation))
                    || ("formula2".equals(field) && customValidation)) return true;
        }
        JsonNode listSource = rule.get("listSource");
        return "dataValidations".equals(property) && listSource != null && listSource.isObject()
                && "formula".equals(listSource.path("kind").asText())
                && listSource.path("formula").isTextual();
    }

    private static void remapPermutationDefinedNames(ArrayNode models, ObjectNode projection, String sheetId, RangeRef scope, int[] rowMap) {
        for (JsonNode raw : models) {
            ObjectNode name = requireObject(raw, "Defined name");
            JsonNode rawAnchor = name.get("anchor");
            if (rawAnchor == null || rawAnchor.isNull()) continue;
            ObjectNode anchor = requireObject(rawAnchor, "Defined-name anchor");
            if (!sheetId.equals(anchor.path("sheetId").asText())) continue;
            JsonNode rowNode = anchor.get("row");
            JsonNode columnNode = anchor.get("column");
            if (rowNode == null || !rowNode.canConvertToInt() || columnNode == null || !columnNode.canConvertToInt()) {
                throw ServiceException.validation("Defined-name anchor is invalid");
            }
            int row = rowNode.asInt(-1);
            int column = columnNode.asInt(-1);
            if (row < 0 || row > SnapshotMutationSupport.MAX_ROW || column < 0 || column > SnapshotMutationSupport.MAX_COLUMN) {
                throw ServiceException.validation("Defined-name anchor is outside worksheet bounds");
            }
            if (!contains(scope, row, column)) continue;
            int targetRow = remapRow(row, scope, rowMap);
            int rowDelta = targetRow - row;
            if (rowDelta == 0) continue;
            JsonNode formula = name.get("formula");
            if (formula == null || !formula.isTextual()) throw ServiceException.validation("Defined-name formula must be text");
            String nameText = SnapshotMutationSupport.text(name, "name");
            String mappedFormula = offsetPermutationFormula(formula.asText(), rowDelta, "defined name " + nameText);
            name.put("formula", mappedFormula);
            anchor.put("row", targetRow);
            if (projection != null && "workbook".equals(name.path("scope").asText())) {
                JsonNode projectedFormula = projection.get(nameText);
                if (projectedFormula != null) {
                    if (!projectedFormula.isTextual()) throw ServiceException.validation("Defined-name projection formula must be text");
                    projection.put(nameText, mappedFormula);
                }
            }
        }
    }

    private static void remapPermutationCellStyleTemplates(ArrayNode templates, String sheetId, RangeRef scope, int[] rowMap) {
        for (JsonNode raw : templates) {
            ObjectNode template = requireObject(raw, "Cell style template");
            JsonNode rawValidation = template.get("dataValidation");
            if (rawValidation == null || rawValidation.isNull()) continue;
            ObjectNode validation = requireObject(rawValidation, "Cell style template validation");
            JsonNode rawAnchor = validation.get("formulaAnchor");
            if (rawAnchor == null || rawAnchor.isNull()) continue;
            ObjectNode anchor = requireObject(rawAnchor, "Cell style template formula anchor");
            if (!sheetId.equals(anchor.path("sheetId").asText())) continue;
            JsonNode rowNode = anchor.get("row");
            JsonNode columnNode = anchor.get("column");
            if (rowNode == null || !rowNode.canConvertToInt() || columnNode == null || !columnNode.canConvertToInt()) {
                throw ServiceException.validation("Cell style template formula anchor is invalid");
            }
            int row = rowNode.asInt(-1);
            int column = columnNode.asInt(-1);
            if (row < 0 || row > SnapshotMutationSupport.MAX_ROW || column < 0 || column > SnapshotMutationSupport.MAX_COLUMN) {
                throw ServiceException.validation("Cell style template formula anchor is outside worksheet bounds");
            }
            if (!contains(scope, row, column)) continue;
            int targetRow = remapRow(row, scope, rowMap);
            int rowDelta = targetRow - row;
            if (rowDelta == 0) continue;
            String identity = "cell-style template " + template.path("id").asText("<unknown>");
            rewriteRuleFormulas(validation, formula -> offsetPermutationFormula(formula, rowDelta, identity));
            anchor.put("row", targetRow);
        }
    }

    private static String offsetPermutationFormula(String formula, int rowDelta, String owner) {
        try {
            FormulaReferenceTransformer.assertRowOffsetSupported(formula);
            return FormulaReferenceTransformer.offsetForPermutation(formula, rowDelta);
        } catch (ServiceException error) {
            throw new ServiceException("SERVICE_UNAVAILABLE", 503,
                    "UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot rewrite formula owner " + owner, error);
        } catch (RuntimeException error) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot parse formula owner " + owner);
        }
    }

    private static void validateDrawingExact(ObjectNode drawing, RangeRef range) {
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
        if ("absolute".equals(anchor.path("kind").asText())) return;
        int row = anchor.path("row").asInt(-1); int column = anchor.path("column").asInt(-1);
        int endRow = anchor.has("endRow") ? anchor.path("endRow").asInt(-1) : row;
        int endColumn = anchor.has("endColumn") ? anchor.path("endColumn").asInt(-1) : column;
        boolean startInside = contains(range, row, column); boolean endInside = contains(range, endRow, endColumn);
        if (startInside != endInside) throw ServiceException.validation("Row permutation cannot exactly remap a drawing anchor");
    }

    private static void requireSingleRange(JsonNode raw, RangeRef range, int[] rowMap, String owner) {
        if (raw != null && remapRangeExact(raw, range, rowMap).size() != 1) throw ServiceException.validation("Row permutation cannot exactly remap " + owner);
    }

    private static void remapRangeRows(JsonNode raw, RangeRef range, int[] rowMap) {
        if (raw == null || !raw.isObject() || !range.sheetId().equals(raw.path("sheetId").asText())) return;
        int start = raw.path("startRow").asInt();
        int end = raw.path("endRow").asInt();
        if (start < range.startRow() || start > range.endRow() || end < range.startRow() || end > range.endRow()) return;
        int remappedStart = remapRow(start, range, rowMap);
        int remappedEnd = remapRow(end, range, rowMap);
        ((ObjectNode) raw).put("startRow", Math.min(remappedStart, remappedEnd)).put("endRow", Math.max(remappedStart, remappedEnd));
    }

    private static int remapRow(int row, RangeRef range, int[] rowMap) {
        if (row < range.startRow() || row > range.endRow()) return row;
        return rowMap[row - range.startRow()];
    }

    private static ObjectNode requireObject(JsonNode value, String label) {
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) value;
    }

    private record CellEntry(int row, int column, ObjectNode cell) {
    }
}

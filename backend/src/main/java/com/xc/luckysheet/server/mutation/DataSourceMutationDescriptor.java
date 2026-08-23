package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

/**
 * Reducers for block-backed data-source and sheet-region metadata.
 *
 * Block bytes are deliberately outside this reducer. An operation can add,
 * replace, or remove only a validated manifest/region descriptor; content is
 * exchanged through the dedicated authenticated data-block endpoint.
 */
final class DataSourceMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "dataSource.add", "dataSource.update", "dataSource.remove",
            "dataRegion.add", "dataRegion.remove"
    );

    private static final int DATA_BLOCK_ROW_COUNT = 65_536;
    private static final int MAX_FIELDS = 16_384;
    private static final int MAX_BLOCKS = 100_000;
    private static final Pattern SAFE_ID = Pattern.compile("[A-Za-z0-9._:-]{1,200}");
    private static final Pattern SHA256 = Pattern.compile("[A-Fa-f0-9]{64}");
    private static final Set<String> SOURCE_KINDS = Set.of("worksheet-range", "sheet-table", "chunked-table");
    private static final Set<String> FIELD_TYPES = Set.of("text", "number", "boolean", "date", "mixed");

    DataSourceMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, true, "structure");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported data-source mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return switch (id()) {
            case "dataSource.add", "dataSource.update" -> sourceRanges(root, validateSourceMutation(root, mutation, params));
            case "dataSource.remove" -> sourceRanges(root, currentSource(root, params));
            case "dataRegion.add" -> List.of(regionRange(root, mutation.sheetId(), params));
            case "dataRegion.remove" -> List.of(currentRegion(root, mutation.sheetId(), params));
            default -> throw ServiceException.validation("Unsupported data-source mutation: " + id());
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        switch (id()) {
            case "dataSource.add" -> addSource(root, mutation, params);
            case "dataSource.update" -> updateSource(root, mutation, params);
            case "dataSource.remove" -> removeSource(root, params);
            case "dataRegion.add" -> addRegion(root, mutation, params);
            case "dataRegion.remove" -> removeRegion(root, mutation, params);
            default -> throw ServiceException.validation("Unsupported data-source mutation: " + id());
        }
        return root;
    }

    private void addSource(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        ObjectNode source = validateSourceMutation(root, mutation, params);
        ArrayNode sources = SnapshotMutationSupport.array(root, "dataSources");
        if (SnapshotMutationSupport.findById(sources, source.path("id").asText()) != null) {
            throw ServiceException.conflict("Data source already exists: " + source.path("id").asText());
        }
        sources.add(source.deepCopy());
    }

    private void updateSource(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        ObjectNode source = validateSourceMutation(root, mutation, params);
        ArrayNode sources = SnapshotMutationSupport.array(root, "dataSources");
        String sourceId = source.path("id").asText();
        int index = SnapshotMutationSupport.indexById(sources, sourceId);
        if (index < 0) throw ServiceException.notFound("Data source not found: " + sourceId);
        sources.set(index, source.deepCopy());
    }

    private void removeSource(ObjectNode root, ObjectNode params) {
        requireKeys(params, Set.of("sourceId"), "dataSource.remove");
        String sourceId = identity(SnapshotMutationSupport.text(params, "sourceId"), "sourceId");
        ArrayNode sources = SnapshotMutationSupport.array(root, "dataSources");
        SnapshotMutationSupport.requireById(sources, sourceId, "Data source");
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            for (JsonNode rawRegion : SnapshotMutationSupport.array((ObjectNode) rawSheet, "dataRegions")) {
                if (sourceId.equals(rawRegion.path("sourceId").asText())) {
                    throw ServiceException.conflict("Data source is still referenced by a sheet region: " + sourceId);
                }
            }
        }
        SnapshotMutationSupport.removeById(sources, sourceId);
    }

    private void addRegion(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        ObjectNode region = validateRegion(root, mutation.sheetId(), params);
        ArrayNode regions = SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, mutation.sheetId()), "dataRegions");
        String regionId = region.path("id").asText();
        if (SnapshotMutationSupport.findById(regions, regionId) != null) {
            throw ServiceException.conflict("Sheet data region already exists: " + regionId);
        }
        regions.add(region.deepCopy());
    }

    private void removeRegion(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        requireKeys(params, Set.of("regionId"), "dataRegion.remove");
        String regionId = identity(SnapshotMutationSupport.text(params, "regionId"), "regionId");
        ArrayNode regions = SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, mutation.sheetId()), "dataRegions");
        if (!SnapshotMutationSupport.removeById(regions, regionId)) {
            throw ServiceException.notFound("Sheet data region not found: " + regionId);
        }
    }

    private ObjectNode validateSourceMutation(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        requireKeys(params, Set.of("source"), id());
        JsonNode raw = params.get("source");
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Data source manifest is required");
        return validateSource(root, mutation.sheetId(), (ObjectNode) raw);
    }

    private ObjectNode currentSource(ObjectNode root, ObjectNode params) {
        requireKeys(params, Set.of("sourceId"), "dataSource.remove");
        String sourceId = identity(SnapshotMutationSupport.text(params, "sourceId"), "sourceId");
        return SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(root, "dataSources"), sourceId, "Data source");
    }

    private ObjectNode validateSource(ObjectNode root, String mutationSheetId, ObjectNode source) {
        SnapshotMutationSupport.validateKnownKeys(source, Set.of(
                "schema", "version", "id", "name", "kind", "sourceSheetId", "sourceRange",
                "rowCount", "fields", "blockRowCount", "blocks", "revision"
        ), "Data source manifest");
        if (!"DataSourceManifest".equals(source.path("schema").asText()) || source.path("version").asInt(-1) != 1) {
            throw ServiceException.validation("Data source manifest schema is invalid");
        }
        String sourceId = identity(SnapshotMutationSupport.text(source, "id"), "Data source id");
        if (SnapshotMutationSupport.text(source, "name").length() > 200) {
            throw ServiceException.validation("Data source name is too long");
        }
        String kind = SnapshotMutationSupport.text(source, "kind");
        if (!SOURCE_KINDS.contains(kind)) throw ServiceException.validation("Data source kind is invalid");

        long rowCount = nonNegative(source, "rowCount");
        long blockRowCount = nonNegative(source, "blockRowCount");
        if (blockRowCount != DATA_BLOCK_ROW_COUNT) {
            throw ServiceException.validation("Data source blockRowCount must be " + DATA_BLOCK_ROW_COUNT);
        }
        nonNegative(source, "revision");

        String sourceSheetId = optionalIdentity(source, "sourceSheetId");
        if (sourceSheetId != null) SnapshotMutationSupport.sheet(root, sourceSheetId);
        JsonNode sourceRangeNode = source.get("sourceRange");
        RangeRef sourceRange = null;
        if (sourceRangeNode != null && !sourceRangeNode.isNull()) {
            sourceRange = SnapshotMutationSupport.range(root, sourceRangeNode);
            if (sourceSheetId == null || !sourceSheetId.equals(sourceRange.sheetId())) {
                throw ServiceException.validation("Data source sourceRange must target sourceSheetId");
            }
        }
        if (sourceSheetId != null && !sourceSheetId.equals(mutationSheetId)) {
            throw ServiceException.validation("Data source sheetId does not match mutation sheetId");
        }
        if (("worksheet-range".equals(kind) || "sheet-table".equals(kind)) && (sourceSheetId == null || sourceRange == null)) {
            throw ServiceException.validation(kind + " data sources require sourceSheetId and sourceRange");
        }

        validateFields(source);
        validateBlocks(source, sourceId, rowCount);
        return source;
    }

    private void validateFields(ObjectNode source) {
        ArrayNode fields = SnapshotMutationSupport.requiredArray(source, "fields");
        if (fields.size() > MAX_FIELDS) throw ServiceException.validation("Data source has too many fields");
        Set<String> ids = new HashSet<>();
        for (int index = 0; index < fields.size(); index++) {
            JsonNode raw = fields.get(index);
            if (raw == null || !raw.isObject()) throw ServiceException.validation("Data source field must be an object");
            ObjectNode field = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(field, Set.of("id", "name", "ordinal", "type"), "Data source field");
            String fieldId = identity(SnapshotMutationSupport.text(field, "id"), "Data source field id");
            if (!ids.add(fieldId)) throw ServiceException.validation("Duplicate data source field: " + fieldId);
            if (field.path("ordinal").asInt(-1) != index) throw ServiceException.validation("Data source fields must use contiguous ordinals");
            if (!FIELD_TYPES.contains(SnapshotMutationSupport.text(field, "type"))) throw ServiceException.validation("Data source field type is invalid");
            if (SnapshotMutationSupport.text(field, "name").length() > 200) throw ServiceException.validation("Data source field name is too long");
        }
    }

    private void validateBlocks(ObjectNode source, String sourceId, long rowCount) {
        ArrayNode blocks = SnapshotMutationSupport.requiredArray(source, "blocks");
        if (blocks.size() > MAX_BLOCKS) throw ServiceException.validation("Data source has too many blocks");
        Set<String> ids = new HashSet<>();
        List<BlockRange> ranges = new ArrayList<>();
        for (JsonNode raw : blocks) {
            if (raw == null || !raw.isObject()) throw ServiceException.validation("Data block must be an object");
            ObjectNode block = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(block, Set.of(
                    "id", "dataSourceId", "startRow", "rowCount", "storageKey", "checksum", "byteLength", "encoding", "revision"
            ), "Data block");
            String blockId = identity(SnapshotMutationSupport.text(block, "id"), "Data block id");
            if (!ids.add(blockId)) throw ServiceException.validation("Duplicate data block: " + blockId);
            if (!sourceId.equals(SnapshotMutationSupport.text(block, "dataSourceId"))) throw ServiceException.validation("Data block belongs to another data source");
            long startRow = nonNegative(block, "startRow");
            long blockRows = nonNegative(block, "rowCount");
            if (blockRows < 1 || startRow + blockRows > rowCount) throw ServiceException.validation("Data block range is invalid: " + blockId);
            if (!"columnar-v1".equals(SnapshotMutationSupport.text(block, "encoding"))) throw ServiceException.validation("Data block encoding is invalid");
            if (!SHA256.matcher(SnapshotMutationSupport.text(block, "checksum")).matches()) throw ServiceException.validation("Data block checksum is invalid: " + blockId);
            if (nonNegative(block, "byteLength") < 1) throw ServiceException.validation("Data block byteLength must be positive: " + blockId);
            nonNegative(block, "revision");
            if (SnapshotMutationSupport.text(block, "storageKey").length() > 500) throw ServiceException.validation("Data block storageKey is too long");
            ranges.add(new BlockRange(startRow, startRow + blockRows, blockId));
        }
        ranges.sort(Comparator.comparingLong(BlockRange::start));
        for (int index = 1; index < ranges.size(); index++) {
            if (ranges.get(index - 1).end() > ranges.get(index).start()) {
                throw ServiceException.validation("Data blocks must not overlap");
            }
        }
    }

    private ObjectNode validateRegion(ObjectNode root, String sheetId, ObjectNode params) {
        requireKeys(params, Set.of("region"), "dataRegion.add");
        JsonNode raw = params.get("region");
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Sheet data region is required");
        ObjectNode region = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(region, Set.of("id", "sourceId", "range", "headerRow", "revision"), "Sheet data region");
        identity(SnapshotMutationSupport.text(region, "id"), "Sheet data region id");
        String sourceId = identity(SnapshotMutationSupport.text(region, "sourceId"), "Sheet data region sourceId");
        SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(root, "dataSources"), sourceId, "Data source");
        RangeRef range = SnapshotMutationSupport.range(root, region.get("range"));
        if (!sheetId.equals(range.sheetId())) throw ServiceException.validation("Sheet data region targets another sheet");
        long headerRow = nonNegative(region, "headerRow");
        if (headerRow < range.startRow() || headerRow > range.endRow()) throw ServiceException.validation("Sheet data region headerRow is outside its range");
        nonNegative(region, "revision");
        return region;
    }

    private RangeRef regionRange(ObjectNode root, String sheetId, ObjectNode params) {
        ObjectNode region = validateRegion(root, sheetId, params);
        return SnapshotMutationSupport.range(root, region.get("range"));
    }

    private RangeRef currentRegion(ObjectNode root, String sheetId, ObjectNode params) {
        requireKeys(params, Set.of("regionId"), "dataRegion.remove");
        String regionId = identity(SnapshotMutationSupport.text(params, "regionId"), "regionId");
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        ObjectNode region = SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(sheet, "dataRegions"), regionId, "Sheet data region");
        return validateExistingRegion(root, sheetId, region);
    }

    private RangeRef validateExistingRegion(ObjectNode root, String sheetId, ObjectNode region) {
        SnapshotMutationSupport.validateKnownKeys(region, Set.of("id", "sourceId", "range", "headerRow", "revision"), "Sheet data region");
        SnapshotMutationSupport.text(region, "id");
        SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(root, "dataSources"), SnapshotMutationSupport.text(region, "sourceId"), "Data source");
        RangeRef range = SnapshotMutationSupport.range(root, region.get("range"));
        if (!sheetId.equals(range.sheetId())) throw ServiceException.validation("Sheet data region targets another sheet");
        return range;
    }

    private List<RangeRef> sourceRanges(ObjectNode root, ObjectNode source) {
        JsonNode sourceRange = source.get("sourceRange");
        if (sourceRange != null && !sourceRange.isNull()) return List.of(SnapshotMutationSupport.range(root, sourceRange));
        String sourceSheetId = optionalIdentity(source, "sourceSheetId");
        return sourceSheetId == null ? List.of() : List.of(SnapshotMutationSupport.wholeSheetRange(root, sourceSheetId));
    }

    private static void requireKeys(ObjectNode object, Set<String> allowed, String label) {
        SnapshotMutationSupport.validateKnownKeys(object, allowed, label);
    }

    private static long nonNegative(ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToLong() || value.longValue() < 0) {
            throw ServiceException.validation(property + " must be a non-negative integer");
        }
        return value.longValue();
    }

    private static String identity(String value, String label) {
        if (!SAFE_ID.matcher(value).matches()) throw ServiceException.validation(label + " is invalid");
        return value;
    }

    private static String optionalIdentity(ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || value.isNull()) return null;
        if (!value.isTextual()) throw ServiceException.validation(property + " must be a string");
        return identity(value.asText(), property);
    }

    private record BlockRange(long start, long end, String id) {
    }
}

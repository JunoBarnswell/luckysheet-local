package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;

import java.io.ByteArrayOutputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Server-side encoder for the canonical columnar-v1 data-block wire format.
 * Keeping encoding on the server lets a large query commit metadata without
 * first copying every result row through the browser main thread.
 */
public final class ColumnarBlockEncoder {
    public static final String ENCODING = "columnar-v1";
    private static final int FIXED_HEADER_BYTES = 40;
    private static final int FIELD_DESCRIPTOR_BYTES = 32;
    private static final int FIELD_SECTION_HEADER_BYTES = 24;
    private static final int CHECKSUM_BYTES = 32;
    private static final int MAX_ROW_COUNT = 65_536;
    private static final int MAX_BLOCK_BYTES = WorkbookDataBlockService.MAX_BLOCK_BYTES;

    public EncodedBlock encode(String sourceId, List<String> columns, List<String> columnTypes,
                               List<List<JsonNode>> rows) {
        if (sourceId == null || sourceId.isBlank() || columns == null || columns.isEmpty()
                || columnTypes == null || columns.size() != columnTypes.size()
                || rows == null || rows.isEmpty() || rows.size() > MAX_ROW_COUNT) {
            throw new IllegalArgumentException("Columnar block input is invalid");
        }
        Set<String> names = new HashSet<>();
        List<EncodedField> fields = new ArrayList<>();
        ByteArrayOutputStream stringTable = new ByteArrayOutputStream();
        int stringOffset = 0;
        int payloadOffset = 0;
        for (int ordinal = 0; ordinal < columns.size(); ordinal++) {
            String name = columns.get(ordinal);
            String type = columnTypes.get(ordinal);
            if (name == null || name.isBlank() || !names.add(name) || type == null) {
                throw new IllegalArgumentException("Columnar block field schema is invalid");
            }
            int typeCode = typeCode(type);
            byte[] id = (sourceId + ":field:" + ordinal).getBytes(StandardCharsets.UTF_8);
            byte[] encodedName = name.getBytes(StandardCharsets.UTF_8);
            List<JsonNode> values = new ArrayList<>(rows.size());
            for (List<JsonNode> row : rows) {
                if (row == null || row.size() != columns.size()) throw new IllegalArgumentException("Columnar block row width is invalid");
                values.add(row.get(ordinal));
            }
            byte[] section = encodeField(type, typeCode, values);
            fields.add(new EncodedField(id, encodedName, section, stringOffset,
                    stringOffset + id.length, payloadOffset));
            stringTable.writeBytes(id);
            stringTable.writeBytes(encodedName);
            stringOffset = checkedAdd(stringOffset, checkedAdd(id.length, encodedName.length, "field schema"), "field schema");
            payloadOffset = checkedAdd(payloadOffset, section.length, "column payload");
        }
        byte[] strings = stringTable.toByteArray();
        int descriptorBytes = checkedMultiply(fields.size(), FIELD_DESCRIPTOR_BYTES, "field descriptors");
        int headerBytes = checkedAdd(FIXED_HEADER_BYTES, checkedAdd(descriptorBytes, strings.length, "block header"), "block header");
        int totalBytes = checkedAdd(checkedAdd(headerBytes, payloadOffset, "block size"), CHECKSUM_BYTES, "block size");
        if (totalBytes > MAX_BLOCK_BYTES) throw new BlockTooLargeException("Columnar block exceeds the configured byte limit");

        byte[] bytes = new byte[totalBytes];
        ByteBuffer header = ByteBuffer.wrap(bytes).order(ByteOrder.LITTLE_ENDIAN);
        bytes[0] = 'L';
        bytes[1] = 'D';
        bytes[2] = 'B';
        bytes[3] = 'K';
        bytes[4] = 1;
        header.putInt(8, headerBytes);
        header.putInt(12, rows.size());
        header.putInt(16, fields.size());
        header.putInt(20, descriptorBytes);
        header.putInt(24, strings.length);
        header.putInt(28, payloadOffset);
        header.putInt(32, CHECKSUM_BYTES);

        int descriptorOffset = FIXED_HEADER_BYTES;
        for (int ordinal = 0; ordinal < fields.size(); ordinal++) {
            EncodedField field = fields.get(ordinal);
            header.putInt(descriptorOffset, field.idOffset());
            header.putInt(descriptorOffset + 4, field.id().length);
            header.putInt(descriptorOffset + 8, field.nameOffset());
            header.putInt(descriptorOffset + 12, field.name().length);
            header.putInt(descriptorOffset + 16, ordinal);
            header.put(descriptorOffset + 20, (byte) typeCode(columnTypes.get(ordinal)));
            header.putInt(descriptorOffset + 24, field.sectionOffset());
            header.putInt(descriptorOffset + 28, field.section().length);
            descriptorOffset += FIELD_DESCRIPTOR_BYTES;
        }
        System.arraycopy(strings, 0, bytes, FIXED_HEADER_BYTES + descriptorBytes, strings.length);
        int payloadWriteOffset = headerBytes;
        for (EncodedField field : fields) {
            System.arraycopy(field.section(), 0, bytes, payloadWriteOffset, field.section().length);
            payloadWriteOffset += field.section().length;
        }
        byte[] embeddedChecksum = digest(bytes, 0, totalBytes - CHECKSUM_BYTES);
        System.arraycopy(embeddedChecksum, 0, bytes, totalBytes - CHECKSUM_BYTES, CHECKSUM_BYTES);
        return new EncodedBlock(bytes, HexFormat.of().formatHex(digest(bytes, 0, bytes.length)));
    }

    private byte[] encodeField(String type, int typeCode, List<JsonNode> values) {
        int rowCount = values.size();
        byte[] validity = new byte[(rowCount + 7) / 8];
        byte[] primary;
        byte[] dictionary = new byte[0];
        byte[] tags = new byte[0];
        if (type.equals("text")) {
            List<String> textValues = new ArrayList<>();
            for (JsonNode value : values) textValues.add(value == null || value.isNull() ? null : requireText(value));
            Dictionary encoded = dictionary(textValues);
            primary = ints(encoded.indexes());
            dictionary = encoded.bytes();
            for (int row = 0; row < rowCount; row++) if (textValues.get(row) != null) setValidity(validity, row);
        } else if (type.equals("number") || type.equals("date")) {
            double[] numbers = new double[rowCount];
            for (int row = 0; row < rowCount; row++) {
                JsonNode value = values.get(row);
                if (value == null || value.isNull()) continue;
                numbers[row] = requireNumber(value);
                setValidity(validity, row);
            }
            primary = doubles(numbers);
        } else if (type.equals("boolean")) {
            primary = new byte[rowCount];
            for (int row = 0; row < rowCount; row++) {
                JsonNode value = values.get(row);
                if (value == null || value.isNull()) continue;
                if (!value.isBoolean()) throw new IllegalArgumentException("Boolean column contains a non-boolean value");
                primary[row] = value.asBoolean() ? (byte) 1 : 0;
                setValidity(validity, row);
            }
        } else if (type.equals("mixed")) {
            double[] numbers = new double[rowCount];
            byte[] booleans = new byte[rowCount];
            int[] indexes = new int[rowCount];
            Arrays.fill(indexes, -1);
            byte[] mixedTags = new byte[rowCount];
            List<String> textValues = new ArrayList<>();
            for (int row = 0; row < rowCount; row++) {
                JsonNode value = values.get(row);
                if (value == null || value.isNull()) {
                    textValues.add(null);
                    continue;
                }
                setValidity(validity, row);
                if (value.isNumber()) {
                    mixedTags[row] = 1;
                    numbers[row] = requireNumber(value);
                    textValues.add(null);
                } else if (value.isBoolean()) {
                    mixedTags[row] = 2;
                    booleans[row] = value.asBoolean() ? (byte) 1 : 0;
                    textValues.add(null);
                } else if (value.isTextual()) {
                    mixedTags[row] = 3;
                    textValues.add(value.asText());
                } else {
                    throw new IllegalArgumentException("Mixed column contains a non-scalar value");
                }
            }
            Dictionary encoded = dictionary(textValues);
            for (int row = 0; row < rowCount; row++) indexes[row] = encoded.indexes()[row];
            primary = concat(doubles(numbers), booleans, ints(indexes));
            dictionary = encoded.bytes();
            tags = mixedTags;
        } else {
            throw new IllegalArgumentException("Unsupported columnar field type: " + type);
        }
        int length = checkedAdd(FIELD_SECTION_HEADER_BYTES,
                checkedAdd(checkedAdd(checkedAdd(validity.length, primary.length, "field section"), dictionary.length, "field section"), tags.length, "field section"),
                "field section");
        byte[] section = new byte[length];
        ByteBuffer view = ByteBuffer.wrap(section).order(ByteOrder.LITTLE_ENDIAN);
        view.put(0, (byte) typeCode);
        view.put(1, (byte) (type.equals("mixed") ? 1 : 0));
        view.putInt(4, values.size());
        view.putInt(8, validity.length);
        view.putInt(12, primary.length);
        view.putInt(16, dictionary.length);
        view.putInt(20, tags.length);
        int offset = FIELD_SECTION_HEADER_BYTES;
        System.arraycopy(validity, 0, section, offset, validity.length);
        offset += validity.length;
        System.arraycopy(primary, 0, section, offset, primary.length);
        offset += primary.length;
        System.arraycopy(dictionary, 0, section, offset, dictionary.length);
        offset += dictionary.length;
        System.arraycopy(tags, 0, section, offset, tags.length);
        return section;
    }

    private Dictionary dictionary(List<String> values) {
        int[] indexes = new int[values.size()];
        Arrays.fill(indexes, -1);
        Map<String, Integer> indexByValue = new HashMap<>();
        List<byte[]> encoded = new ArrayList<>();
        for (int row = 0; row < values.size(); row++) {
            String value = values.get(row);
            if (value == null) continue;
            Integer index = indexByValue.get(value);
            if (index == null) {
                index = encoded.size();
                indexByValue.put(value, index);
                encoded.add(value.getBytes(StandardCharsets.UTF_8));
            }
            indexes[row] = index;
        }
        int stringBytes = 0;
        for (byte[] value : encoded) stringBytes = checkedAdd(stringBytes, value.length, "text dictionary");
        int offsetsBytes = checkedMultiply(encoded.size() + 1, 4, "text dictionary offsets");
        byte[] bytes = new byte[checkedAdd(4, checkedAdd(offsetsBytes, stringBytes, "text dictionary"), "text dictionary")];
        ByteBuffer view = ByteBuffer.wrap(bytes).order(ByteOrder.LITTLE_ENDIAN);
        view.putInt(0, encoded.size());
        int dataOffset = 4 + offsetsBytes;
        int offset = 0;
        view.putInt(4, 0);
        for (int index = 0; index < encoded.size(); index++) {
            byte[] value = encoded.get(index);
            System.arraycopy(value, 0, bytes, dataOffset + offset, value.length);
            offset += value.length;
            view.putInt(4 + (index + 1) * 4, offset);
        }
        return new Dictionary(indexes, bytes);
    }

    private static String requireText(JsonNode value) {
        if (!value.isTextual()) throw new IllegalArgumentException("Text column contains a non-text value");
        return value.asText();
    }

    private static double requireNumber(JsonNode value) {
        if (!value.isNumber() || !Double.isFinite(value.asDouble())) throw new IllegalArgumentException("Number column contains a non-finite value");
        return value.asDouble();
    }

    private static void setValidity(byte[] bitmap, int row) { bitmap[row >> 3] = (byte) (bitmap[row >> 3] | (1 << (row & 7))); }

    private static byte[] doubles(double[] values) {
        ByteBuffer buffer = ByteBuffer.allocate(values.length * Double.BYTES).order(ByteOrder.LITTLE_ENDIAN);
        for (double value : values) buffer.putDouble(value);
        return buffer.array();
    }

    private static byte[] ints(int[] values) {
        ByteBuffer buffer = ByteBuffer.allocate(values.length * Integer.BYTES).order(ByteOrder.LITTLE_ENDIAN);
        for (int value : values) buffer.putInt(value);
        return buffer.array();
    }

    private static byte[] concat(byte[]... parts) {
        int length = 0;
        for (byte[] part : parts) length = checkedAdd(length, part.length, "column payload");
        byte[] result = new byte[length];
        int offset = 0;
        for (byte[] part : parts) {
            System.arraycopy(part, 0, result, offset, part.length);
            offset += part.length;
        }
        return result;
    }

    private static int typeCode(String type) {
        return switch (type) {
            case "text" -> 0;
            case "number" -> 1;
            case "boolean" -> 2;
            case "date" -> 3;
            case "mixed" -> 4;
            default -> throw new IllegalArgumentException("Unsupported columnar field type: " + type);
        };
    }

    private static int checkedAdd(int left, int right, String label) {
        long result = (long) left + right;
        if (result < 0 || result > Integer.MAX_VALUE) throw new IllegalArgumentException(label + " exceeds the binary block limit");
        return (int) result;
    }

    private static int checkedMultiply(int left, int right, String label) {
        long result = (long) left * right;
        if (result < 0 || result > Integer.MAX_VALUE) throw new IllegalArgumentException(label + " exceeds the binary block limit");
        return (int) result;
    }

    private static byte[] digest(byte[] bytes, int offset, int length) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            digest.update(bytes, offset, length);
            return digest.digest();
        } catch (NoSuchAlgorithmException error) {
            throw new IllegalStateException("SHA-256 is unavailable", error);
        }
    }

    private record EncodedField(byte[] id, byte[] name, byte[] section, int idOffset, int nameOffset, int sectionOffset) {}
    private record Dictionary(int[] indexes, byte[] bytes) {}
    public static final class BlockTooLargeException extends IllegalArgumentException {
        private BlockTooLargeException(String message) { super(message); }
    }
    public record EncodedBlock(byte[] content, String checksum) {}
}

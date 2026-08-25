package com.xc.luckysheet.server.mutation;

import com.xc.luckysheet.server.service.ServiceException;

import java.util.Locale;

/**
 * Token-level formula reference syntax tree used by structural reducers.
 *
 * This is intentionally not a regular-expression rewrite. It scans formula
 * syntax while preserving string literals, function names, table references,
 * operators, union/intersection whitespace, implicit-intersection and spill
 * operators. Only parsed A1 references and ranges are offered to a typed
 * mapper; a deleted endpoint becomes one `#REF!` node.
 */
final class FormulaReferenceTransformer {
    private static final int MAX_ROW = 1_048_575;
    private static final int MAX_COLUMN = 16_383;

    private FormulaReferenceTransformer() {
    }

    static String remapAxis(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Axis axis,
            int at,
            int count,
            Direction direction
    ) {
        validateAxis(at, count);
        return rewrite(formula, reference -> {
            if (!belongsToTarget(reference, owner, target)) return reference;
            int position = axis == Axis.ROW ? reference.row() : reference.column();
            if (direction == Direction.INSERT) {
                if (position < at) return reference;
                return axis == Axis.ROW ? reference.withRow(reference.row() + count) : reference.withColumn(reference.column() + count);
            }
            int end = at + count - 1;
            if (position < at) return reference;
            if (position > end) return axis == Axis.ROW ? reference.withRow(reference.row() - count) : reference.withColumn(reference.column() - count);
            return null;
        });
    }

    static String offset(String formula, int rowOffset, int columnOffset) {
        return rewrite(formula, reference -> {
            int row = reference.absoluteRow() ? reference.row() : Math.max(0, reference.row() + rowOffset);
            int column = reference.absoluteColumn() ? reference.column() : Math.max(0, reference.column() + columnOffset);
            return reference.withCoordinates(row, column);
        });
    }

    static String remapMovedRegion(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            int rowDelta,
            int columnDelta
    ) {
        return rewrite(formula, reference -> {
            if (!belongsToTarget(reference, owner, target) || !selection.contains(reference.row(), reference.column())) return reference;
            return reference.withCoordinates(reference.row() + rowDelta, reference.column() + columnDelta);
        });
    }

    static String renameSheet(String formula, String oldName, String newName) {
        if (oldName == null || oldName.isBlank() || newName == null || newName.isBlank()) throw ServiceException.validation("Worksheet names are required for formula rename");
        return rewrite(formula, reference -> {
            if (reference.sheetName() == null || !sameName(reference.sheetName(), oldName)) return reference;
            return reference.withSheetName(newName);
        });
    }

    static String invalidateSheet(String formula, String sheetId, String sheetName) {
        if (sheetId == null || sheetId.isBlank() || sheetName == null || sheetName.isBlank()) throw ServiceException.validation("Worksheet identity is required for formula invalidation");
        return rewrite(formula, reference -> {
            if (reference.sheetName() == null) return reference;
            return sameName(reference.sheetName(), sheetId) || sameName(reference.sheetName(), sheetName) ? null : reference;
        });
    }

    private static boolean belongsToTarget(Reference reference, SheetIdentity owner, SheetIdentity target) {
        if (reference.sheetName() == null) return owner.id().equals(target.id());
        return sameName(reference.sheetName(), target.id()) || sameName(reference.sheetName(), target.name());
    }

    private static boolean sameName(String left, String right) {
        return left.trim().equalsIgnoreCase(right.trim());
    }

    private static void validateAxis(int at, int count) {
        if (at < 0 || count < 1) throw ServiceException.validation("Structural formula transform bounds are invalid");
    }

    private static String rewrite(String formula, ReferenceMapper mapper) {
        if (formula == null || !formula.stripLeading().startsWith("=")) return formula;
        StringBuilder output = new StringBuilder(formula.length());
        int index = 0;
        while (index < formula.length()) {
            char current = formula.charAt(index);
            if (current == '"') {
                int end = consumeString(formula, index);
                output.append(formula, index, end);
                index = end;
                continue;
            }
            ParsedReference parsed = parseQualifiedReference(formula, index);
            if (parsed == null) parsed = parseReference(formula, index, null, null);
            if (parsed == null) {
                if (isSheetIdentifierStart(current)) {
                    int end = index + 1;
                    while (end < formula.length() && isSheetIdentifierPart(formula.charAt(end))) end += 1;
                    output.append(formula, index, end);
                    index = end;
                    continue;
                }
                output.append(current);
                index += 1;
                continue;
            }
            Reference start = mapper.map(parsed.start());
            Reference end = parsed.end() == null ? null : mapper.map(parsed.end());
            if (start == null || (parsed.end() != null && end == null)) output.append("#REF!");
            else output.append(render(parsed, start, end));
            index = parsed.endIndex();
        }
        return output.toString();
    }

    private static int consumeString(String formula, int start) {
        int index = start + 1;
        while (index < formula.length()) {
            if (formula.charAt(index) != '"') {
                index += 1;
                continue;
            }
            if (index + 1 < formula.length() && formula.charAt(index + 1) == '"') {
                index += 2;
                continue;
            }
            return index + 1;
        }
        // Formula parsing will surface an unterminated literal elsewhere. A
        // structural rewrite must not reinterpret the literal as references.
        return formula.length();
    }

    private static ParsedReference parseQualifiedReference(String formula, int start) {
        SheetPrefix prefix = parseSheetPrefix(formula, start);
        if (prefix == null || prefix.afterPrefix() >= formula.length() || formula.charAt(prefix.afterPrefix()) != '!') return null;
        return parseReference(formula, prefix.afterPrefix() + 1, prefix.name(), prefix.raw());
    }

    private static SheetPrefix parseSheetPrefix(String formula, int start) {
        if (formula.charAt(start) == '\'') {
            int index = start + 1;
            StringBuilder name = new StringBuilder();
            while (index < formula.length()) {
                char current = formula.charAt(index);
                if (current != '\'') {
                    name.append(current);
                    index += 1;
                    continue;
                }
                if (index + 1 < formula.length() && formula.charAt(index + 1) == '\'') {
                    name.append('\'');
                    index += 2;
                    continue;
                }
                return new SheetPrefix(name.toString(), formula.substring(start, index + 1), index + 1);
            }
            return null;
        }
        if (!isSheetIdentifierStart(formula.charAt(start))) return null;
        int index = start + 1;
        while (index < formula.length() && isSheetIdentifierPart(formula.charAt(index))) index += 1;
        return new SheetPrefix(formula.substring(start, index), formula.substring(start, index), index);
    }

    private static ParsedReference parseReference(String formula, int start, String sheetName, String rawPrefix) {
        ParsedCell first = parseCell(formula, start, sheetName, rawPrefix);
        if (first == null) return null;
        int index = first.endIndex();
        if (index >= formula.length() || formula.charAt(index) != ':') {
            return new ParsedReference(first.reference(), null, false, first.endIndex());
        }
        int endStart = index + 1;
        SheetPrefix endPrefix = endStart < formula.length() ? parseSheetPrefix(formula, endStart) : null;
        if (endPrefix != null && endPrefix.afterPrefix() < formula.length() && formula.charAt(endPrefix.afterPrefix()) == '!') {
            endStart = endPrefix.afterPrefix() + 1;
        } else endPrefix = null;
        ParsedCell second = parseCell(formula, endStart, endPrefix == null ? sheetName : endPrefix.name(), endPrefix == null ? rawPrefix : endPrefix.raw());
        if (second == null) return null;
        return new ParsedReference(first.reference(), second.reference(), endPrefix != null, second.endIndex());
    }

    private static ParsedCell parseCell(String formula, int start, String sheetName, String rawPrefix) {
        int index = start;
        boolean absoluteColumn = false;
        boolean absoluteRow = false;
        if (index < formula.length() && formula.charAt(index) == '$') {
            absoluteColumn = true;
            index += 1;
        }
        int columnStart = index;
        while (index < formula.length() && isAsciiLetter(formula.charAt(index))) index += 1;
        if (columnStart == index || index - columnStart > 3) return null;
        int column = columnIndex(formula, columnStart, index);
        if (column < 0 || column > MAX_COLUMN) return null;
        if (index < formula.length() && formula.charAt(index) == '$') {
            absoluteRow = true;
            index += 1;
        }
        int rowStart = index;
        while (index < formula.length() && Character.isDigit(formula.charAt(index))) index += 1;
        if (rowStart == index) return null;
        long rowOneBased;
        try {
            rowOneBased = Long.parseLong(formula.substring(rowStart, index));
        } catch (NumberFormatException ignored) {
            return null;
        }
        if (rowOneBased < 1 || rowOneBased > MAX_ROW + 1L) return null;
        if (index < formula.length() && (isAsciiLetter(formula.charAt(index)) || formula.charAt(index) == '_' || Character.isDigit(formula.charAt(index)))) return null;
        String coordinate = formula.substring(start, index);
        return new ParsedCell(new Reference(sheetName, rawPrefix, (int) rowOneBased - 1, column, absoluteRow, absoluteColumn), coordinate, index);
    }

    private static int columnIndex(String formula, int start, int end) {
        int result = 0;
        for (int index = start; index < end; index++) {
            char current = Character.toUpperCase(formula.charAt(index));
            result = result * 26 + (current - 'A' + 1);
        }
        return result - 1;
    }

    private static String render(ParsedReference original, Reference start, Reference end) {
        StringBuilder output = new StringBuilder();
        output.append(renderReference(start, original.start().sheetName(), original.start().rawPrefix()));
        if (end != null) {
            output.append(':');
            if (original.endHadExplicitSheet()) output.append(renderReference(end, original.end().sheetName(), original.end().rawPrefix()));
            else output.append(renderCoordinate(end));
        }
        return output.toString();
    }

    private static String renderReference(Reference reference, String originalSheetName, String originalPrefix) {
        StringBuilder output = new StringBuilder();
        if (reference.sheetName() != null) {
            if (sameName(reference.sheetName(), originalSheetName == null ? reference.sheetName() : originalSheetName) && originalPrefix != null) output.append(originalPrefix).append('!');
            else output.append(renderSheetName(reference.sheetName())).append('!');
        }
        if (reference.absoluteColumn()) output.append('$');
        output.append(columnLabel(reference.column()));
        if (reference.absoluteRow()) output.append('$');
        output.append(reference.row() + 1);
        return output.toString();
    }

    private static String renderCoordinate(Reference reference) {
        StringBuilder output = new StringBuilder();
        if (reference.absoluteColumn()) output.append('$');
        output.append(columnLabel(reference.column()));
        if (reference.absoluteRow()) output.append('$');
        output.append(reference.row() + 1);
        return output.toString();
    }

    private static String renderSheetName(String name) {
        boolean simple = !name.isEmpty();
        for (int index = 0; index < name.length(); index++) {
            char current = name.charAt(index);
            if (!isSheetIdentifierPart(current)) {
                simple = false;
                break;
            }
        }
        return simple ? name : "'" + name.replace("'", "''") + "'";
    }

    private static String columnLabel(int column) {
        StringBuilder output = new StringBuilder();
        int current = column + 1;
        while (current > 0) {
            int remainder = (current - 1) % 26;
            output.append((char) ('A' + remainder));
            current = (current - 1) / 26;
        }
        return output.reverse().toString();
    }

    private static boolean isAsciiLetter(char value) {
        return (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');
    }

    private static boolean isSheetIdentifierStart(char value) {
        return isAsciiLetter(value) || value == '_';
    }

    private static boolean isSheetIdentifierPart(char value) {
        return isAsciiLetter(value) || Character.isDigit(value) || value == '_' || value == '.';
    }

    enum Axis { ROW, COLUMN }

    enum Direction { INSERT, DELETE }

    record SheetIdentity(String id, String name) {
        SheetIdentity {
            if (id == null || id.isBlank() || name == null || name.isBlank()) throw ServiceException.validation("Sheet identity is invalid");
        }
    }

    record Range(int startRow, int endRow, int startColumn, int endColumn) {
        Range {
            if (startRow < 0 || endRow < startRow || startColumn < 0 || endColumn < startColumn) throw ServiceException.validation("Formula reference range is invalid");
        }

        boolean contains(int row, int column) {
            return row >= startRow && row <= endRow && column >= startColumn && column <= endColumn;
        }
    }

    private interface ReferenceMapper {
        Reference map(Reference reference);
    }

    private record Reference(String sheetName, String rawPrefix, int row, int column, boolean absoluteRow, boolean absoluteColumn) {
        Reference withRow(int value) { return withCoordinates(value, column); }
        Reference withColumn(int value) { return withCoordinates(row, value); }
        Reference withCoordinates(int nextRow, int nextColumn) {
            if (nextRow < 0 || nextRow > MAX_ROW || nextColumn < 0 || nextColumn > MAX_COLUMN) return null;
            return new Reference(sheetName, rawPrefix, nextRow, nextColumn, absoluteRow, absoluteColumn);
        }
        Reference withSheetName(String value) { return new Reference(value, rawPrefix, row, column, absoluteRow, absoluteColumn); }
    }

    private record ParsedCell(Reference reference, String raw, int endIndex) {
    }

    private record ParsedReference(Reference start, Reference end, boolean endHadExplicitSheet, int endIndex) {
    }

    private record SheetPrefix(String name, String raw, int afterPrefix) {
    }
}

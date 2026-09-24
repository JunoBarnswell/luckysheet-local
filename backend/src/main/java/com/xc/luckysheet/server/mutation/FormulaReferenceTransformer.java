package com.xc.luckysheet.server.mutation;

import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.TreeSet;

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
        return remapAxis(formula, owner, target, axis, at, count, direction, List.of(owner, target));
    }

    static String remapAxis(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Axis axis,
            int at,
            int count,
            Direction direction,
            List<SheetIdentity> sheetOrder
    ) {
        validateAxis(at, count);
        assertStructuralThreeDimensionalReferences(formula, target, sheetOrder);
        String rewritten = rewrite(formula, reference -> {
            if (!belongsToTarget(reference, owner, target)) return reference;
            int position = axis == Axis.ROW ? reference.row() : reference.column();
            int mapped = mapAxisPoint(position, axis, at, count, direction);
            return mapped < 0 ? null : withAxisCoordinate(reference, axis, mapped);
        }, parsed -> remapAxisRange(parsed, owner, target, axis, at, count, direction), true);
        return remapWholeAxisReferences(rewritten, owner, target, axis, at, count, direction);
    }

    static String remapCellShift(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            Axis axis,
            Direction direction
    ) {
        return remapCellShift(formula, owner, target, selection, axis, direction, List.of(owner, target));
    }

    static String remapCellShift(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            Axis axis,
            Direction direction,
            List<SheetIdentity> sheetOrder
    ) {
        assertStructuralThreeDimensionalReferences(formula, target, sheetOrder);
        int count = axis == Axis.ROW
                ? selection.endRow() - selection.startRow() + 1
                : selection.endColumn() - selection.startColumn() + 1;
        return rewrite(formula, reference -> {
            if (!belongsToTarget(reference, owner, target)) return reference;
            return mapCellShiftPoint(reference, selection, axis, direction, count);
        }, parsed -> remapCellShiftRange(parsed, owner, target, selection, axis, direction, count), true);
    }

    static Range remapCellShiftRangeCoordinates(Range range, Range selection, Axis axis, Direction direction) {
        int count = axis == Axis.ROW
                ? selection.endRow() - selection.startRow() + 1
                : selection.endColumn() - selection.startColumn() + 1;
        Reference start = new Reference(null, null, range.startRow(), range.startColumn(), false, false);
        Reference end = new Reference(null, null, range.endRow(), range.endColumn(), false, false);
        ParsedReference parsed = new ParsedReference(start, end, false, 0);
        SheetIdentity local = new SheetIdentity("__structural_range__", "__structural_range__");
        RangeMapping mapped = remapCellShiftRange(parsed, local, local, selection, axis, direction, count);
        if (!mapped.handled()) throw ServiceException.validation("Cell-shift range transform did not resolve its local range");
        if (mapped.start() == null || mapped.end() == null) return null;
        return new Range(mapped.start().row(), mapped.end().row(), mapped.start().column(), mapped.end().column());
    }

    static int[] remapAxisIntervalCoordinates(int start, int end, Axis axis, int at, int count, Direction direction) {
        validateAxis(at, count);
        int[] interval = transformReferenceInterval(start, end, at, count, direction);
        int maximum = axis == Axis.ROW ? MAX_ROW : MAX_COLUMN;
        return interval == null || interval[0] < 0 || interval[1] > maximum ? null : interval;
    }

    static Range remapAxisRangeCoordinates(Range range, Axis axis, int at, int count, Direction direction) {
        int startCoordinate = axis == Axis.ROW ? range.startRow() : range.startColumn();
        int endCoordinate = axis == Axis.ROW ? range.endRow() : range.endColumn();
        int[] interval = remapAxisIntervalCoordinates(startCoordinate, endCoordinate, axis, at, count, direction);
        if (interval == null) return null;
        boolean reversed = startCoordinate > endCoordinate;
        int mappedStart = reversed ? interval[1] : interval[0];
        int mappedEnd = reversed ? interval[0] : interval[1];
        return axis == Axis.ROW
                ? new Range(mappedStart, mappedEnd, range.startColumn(), range.endColumn())
                : new Range(range.startRow(), range.endRow(), mappedStart, mappedEnd);
    }

    static int[] remapCellShiftCoordinate(int row, int column, Range selection, Axis axis, Direction direction) {
        int count = axis == Axis.ROW
                ? selection.endRow() - selection.startRow() + 1
                : selection.endColumn() - selection.startColumn() + 1;
        return mapCellShiftCoordinate(row, column, selection, axis, direction, count);
    }

    static String offset(String formula, int rowOffset, int columnOffset) {
        return rewrite(formula, reference -> {
            long row = reference.absoluteRow() ? reference.row() : (long) reference.row() + rowOffset;
            long column = reference.absoluteColumn() ? reference.column() : (long) reference.column() + columnOffset;
            if (row < 0 || row > MAX_ROW || column < 0 || column > MAX_COLUMN) return null;
            return reference.withCoordinates((int) row, (int) column);
        });
    }

    static String offsetForPermutation(String formula, int rowOffset) {
        return rewrite(formula, reference -> {
            long row = reference.absoluteRow() ? reference.row() : (long) reference.row() + rowOffset;
            if (row < 0 || row > MAX_ROW) {
                throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation would move a formula reference outside worksheet bounds");
            }
            return reference.withCoordinates((int) row, reference.column());
        });
    }

    static void assertRowOffsetSupported(String formula) {
        if (formula == null) return;
        int index = 0;
        while (index < formula.length()) {
            if (formula.charAt(index) == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (formula.charAt(index) == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                if (externalEnd > index) {
                    throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot offset an external-workbook formula reference");
                }
                index = consumeBracketedReference(formula, index);
                continue;
            }

            SheetPrefix prefix = parseSheetPrefix(formula, index);
            if (prefix != null) {
                if (prefix.name().indexOf('[') >= 0 && prefix.name().indexOf(']') > prefix.name().indexOf('[')) {
                    throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot offset an external-workbook formula reference");
                }
                if (prefix.afterPrefix() < formula.length() && formula.charAt(prefix.afterPrefix()) == '!') {
                    WholeAxisReference qualified = parseWholeAxisReference(formula, prefix.afterPrefix() + 1);
                    if (qualified != null) {
                        assertRowAxisOffsetSupported(qualified);
                        index = qualified.endIndex();
                        continue;
                    }
                }
            }

            if (prefix == null || (prefix.afterPrefix() < formula.length() && formula.charAt(prefix.afterPrefix()) == ':')) {
                WholeAxisReference wholeAxis = parseWholeAxisReference(formula, index);
                if (wholeAxis != null) {
                    assertRowAxisOffsetSupported(wholeAxis);
                    index = wholeAxis.endIndex();
                    continue;
                }
            }
            index = nextReferenceCandidate(formula, index, prefix);
        }
    }

    private static void assertRowAxisOffsetSupported(WholeAxisReference reference) {
        if (reference.axis() == Axis.ROW) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: row permutation cannot offset a whole-row formula reference");
        }
    }

    static String remapMovedRegion(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            int rowDelta,
            int columnDelta,
            List<SheetIdentity> sheetOrder
    ) {
        assertMovedThreeDimensionalReferences(formula, target, sheetOrder);
        assertMovedWholeAxisReferences(formula, owner, target, selection, rowDelta, columnDelta);
        return rewrite(formula, reference -> {
            if (!belongsToTarget(reference, owner, target) || !selection.contains(reference.row(), reference.column())) return reference;
            return reference.withCoordinates(reference.row() + rowDelta, reference.column() + columnDelta);
        }, parsed -> remapMovedRange(parsed, owner, target, selection, rowDelta, columnDelta), true);
    }

    private static RangeMapping remapMovedRange(
            ParsedReference parsed,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            int rowDelta,
            int columnDelta
    ) {
        if (hasDifferentSheetEndpoints(parsed)) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: moving cells cannot rewrite a cross-worksheet range");
        }
        boolean startTargets = belongsToTarget(parsed.start(), owner, target);
        boolean endTargets = belongsToTarget(parsed.end(), owner, target);
        if (!startTargets && !endTargets) return RangeMapping.notHandled();
        if (!startTargets || !endTargets) {
            throw ServiceException.validation("UNSUPPORTED_STRUCTURAL_REFERENCE: moved range has a partially qualified formula reference");
        }

        int lowRow = Math.min(parsed.start().row(), parsed.end().row());
        int highRow = Math.max(parsed.start().row(), parsed.end().row());
        int lowColumn = Math.min(parsed.start().column(), parsed.end().column());
        int highColumn = Math.max(parsed.start().column(), parsed.end().column());
        boolean intersects = lowRow <= selection.endRow() && highRow >= selection.startRow()
                && lowColumn <= selection.endColumn() && highColumn >= selection.startColumn();
        if (!intersects) return RangeMapping.notHandled();
        boolean contained = lowRow >= selection.startRow() && highRow <= selection.endRow()
                && lowColumn >= selection.startColumn() && highColumn <= selection.endColumn();
        if (!contained) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: moving cells would make a formula reference non-contiguous");
        }
        Reference start = parsed.start().withCoordinates(parsed.start().row() + rowDelta, parsed.start().column() + columnDelta);
        Reference end = parsed.end().withCoordinates(parsed.end().row() + rowDelta, parsed.end().column() + columnDelta);
        return RangeMapping.handled(start, end);
    }

    private static void assertMovedThreeDimensionalReferences(String formula, SheetIdentity target, List<SheetIdentity> sheetOrder) {
        assertStructuralThreeDimensionalReferences(formula, target, sheetOrder);
    }

    private static void assertStructuralThreeDimensionalReferences(String formula, SheetIdentity target, List<SheetIdentity> sheetOrder) {
        if (formula == null) return;
        int index = 0;
        while (index < formula.length()) {
            if (formula.charAt(index) == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (formula.charAt(index) == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                index = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                continue;
            }
            SheetPrefix first = parseSheetPrefix(formula, index);
            if (first != null && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == '!'
                    && first.name().contains(":")) {
                int separator = first.name().indexOf(':');
                if (separator == 0 || separator == first.name().length() - 1 || first.name().indexOf(':', separator + 1) >= 0) {
                    throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: 3D reference boundary is invalid");
                }
                assertTargetOutsideThreeDimensionalRange(
                        first.name().substring(0, separator), first.name().substring(separator + 1), target, sheetOrder);
                index = first.afterPrefix() + 1;
                continue;
            }
            if (first != null && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == ':') {
                int secondStart = first.afterPrefix() + 1;
                SheetPrefix second = secondStart < formula.length() ? parseSheetPrefix(formula, secondStart) : null;
                if (second != null && second.afterPrefix() < formula.length() && formula.charAt(second.afterPrefix()) == '!') {
                    assertTargetOutsideThreeDimensionalRange(first.name(), second.name(), target, sheetOrder);
                    index = second.afterPrefix() + 1;
                    continue;
                }
            }
            index = nextReferenceCandidate(formula, index, first);
        }
    }

    private static void assertTargetOutsideThreeDimensionalRange(
            String startName,
            String endName,
            SheetIdentity target,
            List<SheetIdentity> sheetOrder
    ) {
        int start = sheetIndex(sheetOrder, startName);
        int end = sheetIndex(sheetOrder, endName);
        int moved = sheetIndexById(sheetOrder, target.id());
        if (start < 0 || end < 0 || moved < 0) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: 3D reference boundary is unresolved");
        }
        if (moved >= Math.min(start, end) && moved <= Math.max(start, end)) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: structural edits cannot rewrite one sheet inside a 3D reference");
        }
    }

    private static int sheetIndex(List<SheetIdentity> sheets, String value) {
        for (int index = 0; index < sheets.size(); index++) {
            if (sameName(sheets.get(index).name(), value)) return index;
        }
        for (int index = 0; index < sheets.size(); index++) {
            if (sheets.get(index).id().equals(value)) return index;
        }
        return -1;
    }

    private static int sheetIndexById(List<SheetIdentity> sheets, String id) {
        for (int index = 0; index < sheets.size(); index++) {
            if (sheets.get(index).id().equals(id)) return index;
        }
        return -1;
    }

    private static void assertMovedWholeAxisReferences(
            String formula,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            int rowDelta,
            int columnDelta
    ) {
        if (formula == null) return;
        int index = 0;
        while (index < formula.length()) {
            if (formula.charAt(index) == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (formula.charAt(index) == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                index = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                continue;
            }
            int threeDimensionalEnd = threeDimensionalReferenceEnd(formula, index);
            if (threeDimensionalEnd > index) {
                index = threeDimensionalEnd;
                continue;
            }
            SheetPrefix prefix = parseSheetPrefix(formula, index);
            String sheetName = null;
            int referenceStart = index;
            if (prefix != null && prefix.afterPrefix() < formula.length() && formula.charAt(prefix.afterPrefix()) == '!') {
                sheetName = prefix.name();
                referenceStart = prefix.afterPrefix() + 1;
            }
            WholeAxisReference reference = parseWholeAxisReference(formula, referenceStart);
            if (reference == null) {
                index = nextReferenceCandidate(formula, index, prefix);
                continue;
            }
            boolean targets = sheetName == null
                    ? owner.id().equals(target.id())
                    : sameName(sheetName, target.name());
            if (targets && reference.axis() == Axis.ROW && rowDelta != 0) {
                assertWholeAxisMoveIsRepresentable(reference.start(), reference.end(), selection.startRow(), selection.endRow(), rowDelta, "row");
            } else if (targets && reference.axis() == Axis.COLUMN && columnDelta != 0) {
                assertWholeAxisMoveIsRepresentable(reference.start(), reference.end(), selection.startColumn(), selection.endColumn(), columnDelta, "column");
            }
            index = Math.max(reference.endIndex(), referenceStart + 1);
        }
    }

    private static int threeDimensionalReferenceEnd(String formula, int index) {
        SheetPrefix first = parseSheetPrefix(formula, index);
        if (first == null) return -1;
        int coordinateStart = -1;
        if (first.name().contains(":") && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == '!') {
            coordinateStart = first.afterPrefix() + 1;
        } else if (first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == ':') {
            int secondStart = first.afterPrefix() + 1;
            SheetPrefix second = secondStart < formula.length() ? parseSheetPrefix(formula, secondStart) : null;
            if (second != null && second.afterPrefix() < formula.length() && formula.charAt(second.afterPrefix()) == '!') {
                coordinateStart = second.afterPrefix() + 1;
            }
        }
        if (coordinateStart < 0) return -1;
        WholeAxisReference wholeAxis = parseWholeAxisReference(formula, coordinateStart);
        if (wholeAxis != null) return wholeAxis.endIndex();
        ParsedReference reference = parseReference(formula, coordinateStart, null, null);
        return reference == null ? coordinateStart : reference.endIndex();
    }

    private static void assertWholeAxisMoveIsRepresentable(int referenceStart, int referenceEnd,
            int sourceStart, int sourceEnd, int delta, String label) {
        int low = Math.min(referenceStart, referenceEnd);
        int high = Math.max(referenceStart, referenceEnd);
        int targetStart = sourceStart + delta;
        int targetEnd = sourceEnd + delta;
        boolean sourceIntersects = sourceStart <= high && sourceEnd >= low;
        boolean targetIntersects = targetStart <= high && targetEnd >= low;
        boolean bothCovered = sourceStart >= low && sourceEnd <= high && targetStart >= low && targetEnd <= high;
        if (!bothCovered && (sourceIntersects || targetIntersects)) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: moving cells would make a whole-" + label + " reference non-contiguous");
        }
    }

    static String renameSheet(String formula, String oldName, String newName) {
        if (oldName == null || oldName.isBlank() || newName == null || newName.isBlank()) throw ServiceException.validation("Worksheet names are required for formula rename");
        String rewritten = rewrite(formula, reference -> {
            if (reference.sheetName() == null || !sameName(reference.sheetName(), oldName)) return reference;
            return reference.withSheetName(newName);
        }, null, true);
        rewritten = rewriteWholeAxisSheetName(rewritten, oldName, newName, false);
        return rewriteThreeDimensionalSheetNames(rewritten, oldName, newName);
    }

    static String invalidateSheet(String formula, String sheetId, String sheetName) {
        if (sheetId == null || sheetId.isBlank() || sheetName == null || sheetName.isBlank()) throw ServiceException.validation("Worksheet identity is required for formula invalidation");
        assertNoThreeDimensionalReference(formula);
        String rewritten = rewrite(formula, reference -> {
            if (reference.sheetName() == null) return reference;
            return sameName(reference.sheetName(), sheetName) ? null : reference;
        });
        return rewriteWholeAxisSheetName(rewritten, sheetName, sheetName, true);
    }

    private static boolean belongsToTarget(Reference reference, SheetIdentity owner, SheetIdentity target) {
        if (reference.sheetName() == null) return owner.id().equals(target.id());
        return sameName(reference.sheetName(), target.name());
    }

    private static boolean sameName(String left, String right) {
        return left.trim().equalsIgnoreCase(right.trim());
    }

    private static void validateAxis(int at, int count) {
        if (at < 0 || count < 1) throw ServiceException.validation("Structural formula transform bounds are invalid");
    }

    private static String rewrite(String formula, ReferenceMapper mapper) {
        return rewrite(formula, mapper, null, false);
    }

    private static String rewrite(String formula, ReferenceMapper mapper, ReferenceRangeMapper rangeMapper) {
        return rewrite(formula, mapper, rangeMapper, false);
    }

    private static String rewrite(String formula, ReferenceMapper mapper, ReferenceRangeMapper rangeMapper, boolean preserveThreeDimensionalReferences) {
        if (formula == null) return null;
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
            if (current == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                int end = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                output.append(formula, index, end);
                index = end;
                continue;
            }
            if (preserveThreeDimensionalReferences) {
                int threeDimensionalEnd = threeDimensionalReferenceEnd(formula, index);
                if (threeDimensionalEnd > index) {
                    output.append(formula, index, threeDimensionalEnd);
                    index = threeDimensionalEnd;
                    continue;
                }
            }
            if (index > 0 && isReferenceNamePart(formula.charAt(index - 1))) {
                output.append(current);
                index += 1;
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
            RangeMapping mappedRange = parsed.end() == null || rangeMapper == null
                    ? RangeMapping.notHandled()
                    : rangeMapper.map(parsed);
            if (mappedRange.handled()) {
                if (mappedRange.start() == null || mappedRange.end() == null) output.append("#REF!");
                else output.append(render(parsed, mappedRange.start(), mappedRange.end()));
            } else {
                Reference start = mapper.map(parsed.start());
                Reference end = parsed.end() == null ? null : mapper.map(parsed.end());
                if (start == null || (parsed.end() != null && end == null)) output.append("#REF!");
                else output.append(render(parsed, start, end));
            }
            index = parsed.endIndex();
        }
        return output.toString();
    }

    private static String remapWholeAxisReferences(String formula, SheetIdentity owner, SheetIdentity target,
            Axis axis, int at, int count, Direction direction) {
        if (formula == null || formula.isEmpty()) return formula;
        StringBuilder output = new StringBuilder(formula.length());
        int copied = 0;
        int index = 0;
        while (index < formula.length()) {
            char current = formula.charAt(index);
            if (current == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (current == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                index = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                continue;
            }
            int threeDimensionalEnd = threeDimensionalReferenceEnd(formula, index);
            if (threeDimensionalEnd > index) {
                index = threeDimensionalEnd;
                continue;
            }

            SheetPrefix prefix = parseSheetPrefix(formula, index);
            if (prefix != null && prefix.afterPrefix() < formula.length() && formula.charAt(prefix.afterPrefix()) == '!') {
                WholeAxisReference qualified = parseWholeAxisReference(formula, prefix.afterPrefix() + 1);
                if (qualified != null) {
                    boolean targetsSheet = sameName(prefix.name(), target.name());
                    if (targetsSheet && qualified.axis() == axis) {
                        int[] interval = remapAxisIntervalCoordinates(qualified.start(), qualified.end(), axis, at, count, direction);
                        output.append(formula, copied, qualified.startIndex());
                        output.append(interval == null ? "#REF!" : renderWholeAxisReference(formula, qualified, interval));
                        copied = qualified.endIndex();
                    }
                    index = qualified.endIndex();
                    continue;
                }
            }

            WholeAxisReference unqualified = parseWholeAxisReference(formula, index);
            if (unqualified != null) {
                if (owner.id().equals(target.id()) && unqualified.axis() == axis) {
                    int[] interval = remapAxisIntervalCoordinates(unqualified.start(), unqualified.end(), axis, at, count, direction);
                    output.append(formula, copied, unqualified.startIndex());
                    output.append(interval == null ? "#REF!" : renderWholeAxisReference(formula, unqualified, interval));
                    copied = unqualified.endIndex();
                }
                index = unqualified.endIndex();
                continue;
            }
            index = prefix == null ? index + 1 : Math.max(index + 1, prefix.afterPrefix());
        }
        if (copied == 0) return formula;
        output.append(formula, copied, formula.length());
        return output.toString();
    }

    private static String renderWholeAxisReference(String formula, WholeAxisReference reference, int[] interval) {
        StringBuilder output = new StringBuilder();
        output.append(formula, reference.startIndex(), reference.firstCoordinateStart());
        int mappedStart = reference.start() > reference.end() ? interval[1] : interval[0];
        int mappedEnd = reference.start() > reference.end() ? interval[0] : interval[1];
        if (reference.axis() == Axis.ROW) output.append(mappedStart + 1);
        else output.append(columnLabel(mappedStart));
        output.append(formula, reference.firstCoordinateEnd(), reference.secondCoordinateStart());
        if (reference.axis() == Axis.ROW) output.append(mappedEnd + 1);
        else output.append(columnLabel(mappedEnd));
        output.append(formula, reference.secondCoordinateEnd(), reference.endIndex());
        return output.toString();
    }

    private static String rewriteWholeAxisSheetName(String formula, String firstName, String secondName, boolean invalidate) {
        if (formula == null || formula.isEmpty()) return formula;
        StringBuilder output = new StringBuilder(formula.length());
        int copied = 0;
        int index = 0;
        while (index < formula.length()) {
            char current = formula.charAt(index);
            if (current == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (current == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                index = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                continue;
            }
            int threeDimensionalEnd = threeDimensionalReferenceEnd(formula, index);
            if (threeDimensionalEnd > index) {
                index = threeDimensionalEnd;
                continue;
            }
            SheetPrefix prefix = parseSheetPrefix(formula, index);
            if (prefix == null || prefix.afterPrefix() >= formula.length() || formula.charAt(prefix.afterPrefix()) != '!') {
                index = prefix == null ? index + 1 : Math.max(index + 1, prefix.afterPrefix());
                continue;
            }
            WholeAxisReference reference = parseWholeAxisReference(formula, prefix.afterPrefix() + 1);
            if (reference == null) {
                index = prefix.afterPrefix() + 1;
                continue;
            }
            if (sameName(prefix.name(), firstName) || sameName(prefix.name(), secondName)) {
                output.append(formula, copied, index);
                if (invalidate) output.append("#REF!");
                else output.append(renderSheetName(secondName)).append('!').append(formula, prefix.afterPrefix() + 1, reference.endIndex());
                copied = reference.endIndex();
            }
            index = reference.endIndex();
        }
        if (copied == 0) return formula;
        output.append(formula, copied, formula.length());
        return output.toString();
    }

    private static String rewriteThreeDimensionalSheetNames(String formula, String oldName, String newName) {
        if (formula == null || formula.isEmpty()) return formula;
        StringBuilder output = new StringBuilder(formula.length());
        int copied = 0;
        int index = 0;
        while (index < formula.length()) {
            char current = formula.charAt(index);
            if (current == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (current == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                index = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                continue;
            }
            int referenceEnd = threeDimensionalReferenceEnd(formula, index);
            if (referenceEnd <= index) {
                index = nextReferenceCandidate(formula, index, parseSheetPrefix(formula, index));
                continue;
            }

            SheetPrefix first = parseSheetPrefix(formula, index);
            String replacement = null;
            int prefixEnd = -1;
            if (first != null && first.name().contains(":")
                    && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == '!') {
                int separator = first.name().indexOf(':');
                if (separator > 0 && separator == first.name().lastIndexOf(':') && separator < first.name().length() - 1) {
                    String startName = first.name().substring(0, separator);
                    String endName = first.name().substring(separator + 1);
                    String renamedStart = renamedSheetName(startName, oldName, newName);
                    String renamedEnd = renamedSheetName(endName, oldName, newName);
                    if (!startName.equals(renamedStart) || !endName.equals(renamedEnd)) {
                        replacement = renderSheetName(renamedStart + ":" + renamedEnd) + "!";
                        prefixEnd = first.afterPrefix() + 1;
                    }
                }
            } else if (first != null && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == ':') {
                int secondStart = first.afterPrefix() + 1;
                SheetPrefix second = secondStart < formula.length() ? parseSheetPrefix(formula, secondStart) : null;
                if (second != null && second.afterPrefix() < formula.length() && formula.charAt(second.afterPrefix()) == '!') {
                    String renamedStart = renamedSheetName(first.name(), oldName, newName);
                    String renamedEnd = renamedSheetName(second.name(), oldName, newName);
                    if (!first.name().equals(renamedStart) || !second.name().equals(renamedEnd)) {
                        String renderedStart = first.name().equals(renamedStart) ? first.raw() : renderSheetName(renamedStart);
                        String renderedEnd = second.name().equals(renamedEnd) ? second.raw() : renderSheetName(renamedEnd);
                        replacement = renderedStart + ":" + renderedEnd + "!";
                        prefixEnd = second.afterPrefix() + 1;
                    }
                }
            }
            if (replacement != null) {
                output.append(formula, copied, index).append(replacement);
                copied = prefixEnd;
            }
            index = referenceEnd;
        }
        if (copied == 0) return formula;
        output.append(formula, copied, formula.length());
        return output.toString();
    }

    private static String renamedSheetName(String current, String oldName, String newName) {
        return sameName(current, oldName) ? newName : current;
    }

    private static int consumeExternalReference(String formula, int start) {
        int closingBook = formula.indexOf(']', start + 1);
        if (closingBook <= start + 1 || closingBook + 1 >= formula.length()
                || !isSheetIdentifierStart(formula.charAt(closingBook + 1))) return -1;
        int cursor = closingBook + 2;
        while (cursor < formula.length() && isSheetIdentifierPart(formula.charAt(cursor))) cursor += 1;
        if (cursor >= formula.length() || formula.charAt(cursor) != '!') return -1;
        ParsedReference external = parseReference(formula, cursor + 1, null, null);
        if (external != null) return external.endIndex();
        WholeAxisReference wholeAxis = parseWholeAxisReference(formula, cursor + 1);
        return wholeAxis == null ? -1 : wholeAxis.endIndex();
    }

    private static int consumeBracketedReference(String formula, int start) {
        int depth = 0;
        for (int index = start; index < formula.length(); index += 1) {
            if (formula.charAt(index) == '[') depth += 1;
            else if (formula.charAt(index) == ']') {
                depth -= 1;
                if (depth == 0) return index + 1;
            }
        }
        return formula.length();
    }

    private static RangeMapping remapAxisRange(
            ParsedReference parsed,
            SheetIdentity owner,
            SheetIdentity target,
            Axis axis,
            int at,
            int count,
            Direction direction
    ) {
        boolean startTargets = belongsToTarget(parsed.start(), owner, target);
        boolean endTargets = belongsToTarget(parsed.end(), owner, target);
        if (!startTargets && !endTargets) return RangeMapping.notHandled();
        if (hasDifferentSheetEndpoints(parsed)) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: structural transform cannot rewrite a cross-worksheet range");
        if (!startTargets || !endTargets) throw ServiceException.validation("Structural transform cannot rewrite a partially qualified range");

        int startPosition = axis == Axis.ROW ? parsed.start().row() : parsed.start().column();
        int endPosition = axis == Axis.ROW ? parsed.end().row() : parsed.end().column();
        int[] interval = remapAxisIntervalCoordinates(startPosition, endPosition, axis, at, count, direction);
        if (interval == null) return RangeMapping.handled(null, null);
        boolean reversed = startPosition > endPosition;
        Reference start = withAxisCoordinate(parsed.start(), axis, reversed ? interval[1] : interval[0]);
        Reference end = withAxisCoordinate(parsed.end(), axis, reversed ? interval[0] : interval[1]);
        return RangeMapping.handled(start, end);
    }

    private static RangeMapping remapCellShiftRange(
            ParsedReference parsed,
            SheetIdentity owner,
            SheetIdentity target,
            Range selection,
            Axis axis,
            Direction direction,
            int count
    ) {
        boolean startTargets = belongsToTarget(parsed.start(), owner, target);
        boolean endTargets = belongsToTarget(parsed.end(), owner, target);
        if (!startTargets && !endTargets) return RangeMapping.notHandled();
        if (hasDifferentSheetEndpoints(parsed)) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift cannot rewrite a cross-worksheet range");
        if (!startTargets || !endTargets) throw ServiceException.validation("Cell shift cannot rewrite a partially qualified range");

        int lowRow = Math.min(parsed.start().row(), parsed.end().row());
        int highRow = Math.max(parsed.start().row(), parsed.end().row());
        int lowColumn = Math.min(parsed.start().column(), parsed.end().column());
        int highColumn = Math.max(parsed.start().column(), parsed.end().column());
        List<Integer> rowCuts = axis == Axis.ROW
                ? (direction == Direction.DELETE ? List.of(selection.startRow(), selection.endRow() + 1) : List.of(selection.startRow()))
                : List.of(selection.startRow(), selection.endRow() + 1);
        List<Integer> columnCuts = axis == Axis.COLUMN
                ? (direction == Direction.DELETE ? List.of(selection.startColumn(), selection.endColumn() + 1) : List.of(selection.startColumn()))
                : List.of(selection.startColumn(), selection.endColumn() + 1);
        List<Rectangle> rectangles = new ArrayList<>();
        for (int[] rows : splitInterval(lowRow, highRow, rowCuts)) {
            for (int[] columns : splitInterval(lowColumn, highColumn, columnCuts)) {
                int[] mappedStart = mapCellShiftCoordinate(rows[0], columns[0], selection, axis, direction, count);
                int[] mappedEnd = mapCellShiftCoordinate(rows[1], columns[1], selection, axis, direction, count);
                if (mappedStart == null || mappedEnd == null) continue;
                rectangles.add(new Rectangle(
                        Math.min(mappedStart[0], mappedEnd[0]), Math.max(mappedStart[0], mappedEnd[0]),
                        Math.min(mappedStart[1], mappedEnd[1]), Math.max(mappedStart[1], mappedEnd[1])));
            }
        }
        List<Rectangle> merged = mergeRectangles(rectangles);
        if (merged.size() > 1) throw ServiceException.unavailable("Cell shift makes a formula range non-contiguous");
        if (merged.isEmpty()) return RangeMapping.handled(null, null);
        Rectangle rectangle = merged.get(0);
        boolean reverseRows = parsed.start().row() > parsed.end().row();
        boolean reverseColumns = parsed.start().column() > parsed.end().column();
        Reference start = parsed.start().withCoordinates(
                reverseRows ? rectangle.endRow() : rectangle.startRow(),
                reverseColumns ? rectangle.endColumn() : rectangle.startColumn());
        Reference end = parsed.end().withCoordinates(
                reverseRows ? rectangle.startRow() : rectangle.endRow(),
                reverseColumns ? rectangle.startColumn() : rectangle.endColumn());
        return RangeMapping.handled(start, end);
    }

    private static Reference mapCellShiftPoint(Reference reference, Range selection, Axis axis, Direction direction, int count) {
        int row = reference.row();
        int column = reference.column();
        if (axis == Axis.ROW) {
            if (column < selection.startColumn() || column > selection.endColumn() || row < selection.startRow()) return reference;
            if (direction == Direction.DELETE && row <= selection.endRow()) return null;
            int nextRow = row + (direction == Direction.INSERT ? count : -count);
            if (nextRow > MAX_ROW) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet row bounds");
            return reference.withRow(nextRow);
        }
        if (row < selection.startRow() || row > selection.endRow() || column < selection.startColumn()) return reference;
        if (direction == Direction.DELETE && column <= selection.endColumn()) return null;
        int nextColumn = column + (direction == Direction.INSERT ? count : -count);
        if (nextColumn > MAX_COLUMN) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet column bounds");
        return reference.withColumn(nextColumn);
    }

    private static int[] mapCellShiftCoordinate(int row, int column, Range selection, Axis axis, Direction direction, int count) {
        if (axis == Axis.ROW) {
            if (column < selection.startColumn() || column > selection.endColumn() || row < selection.startRow()) return new int[]{row, column};
            if (direction == Direction.DELETE && row <= selection.endRow()) return null;
            int nextRow = row + (direction == Direction.INSERT ? count : -count);
            if (nextRow > MAX_ROW) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet row bounds");
            return new int[]{nextRow, column};
        }
        if (row < selection.startRow() || row > selection.endRow() || column < selection.startColumn()) return new int[]{row, column};
        if (direction == Direction.DELETE && column <= selection.endColumn()) return null;
        int nextColumn = column + (direction == Direction.INSERT ? count : -count);
        if (nextColumn > MAX_COLUMN) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet column bounds");
        return new int[]{row, nextColumn};
    }

    private static List<int[]> splitInterval(int start, int end, List<Integer> boundaries) {
        TreeSet<Integer> cuts = new TreeSet<>();
        cuts.add(start);
        for (int boundary : boundaries) if (boundary > start && boundary <= end) cuts.add(boundary);
        cuts.add(end + 1);
        List<Integer> ordered = new ArrayList<>(cuts);
        List<int[]> result = new ArrayList<>();
        for (int index = 0; index + 1 < ordered.size(); index++) {
            result.add(new int[]{ordered.get(index), ordered.get(index + 1) - 1});
        }
        return result;
    }

    private static List<Rectangle> mergeRectangles(List<Rectangle> rectangles) {
        List<Rectangle> result = new ArrayList<>(rectangles);
        boolean merged = true;
        while (merged) {
            merged = false;
            for (int leftIndex = 0; leftIndex < result.size() && !merged; leftIndex++) {
                for (int rightIndex = leftIndex + 1; rightIndex < result.size(); rightIndex++) {
                    Rectangle left = result.get(leftIndex);
                    Rectangle right = result.get(rightIndex);
                    boolean sameRows = left.startRow() == right.startRow() && left.endRow() == right.endRow();
                    boolean sameColumns = left.startColumn() == right.startColumn() && left.endColumn() == right.endColumn();
                    if (sameRows && (left.endColumn() + 1 == right.startColumn() || right.endColumn() + 1 == left.startColumn())) {
                        result.set(leftIndex, new Rectangle(left.startRow(), left.endRow(),
                                Math.min(left.startColumn(), right.startColumn()), Math.max(left.endColumn(), right.endColumn())));
                    } else if (sameColumns && (left.endRow() + 1 == right.startRow() || right.endRow() + 1 == left.startRow())) {
                        result.set(leftIndex, new Rectangle(Math.min(left.startRow(), right.startRow()),
                                Math.max(left.endRow(), right.endRow()), left.startColumn(), left.endColumn()));
                    } else continue;
                    result.remove(rightIndex);
                    merged = true;
                    break;
                }
            }
        }
        return result;
    }

    private static int[] transformReferenceInterval(int start, int end, int at, int count, Direction direction) {
        long low = Math.min(start, end);
        long high = Math.max(start, end);
        long nextStart;
        long nextEnd;
        if (direction == Direction.INSERT) {
            if (at <= low) {
                nextStart = low + count;
                nextEnd = high + count;
            } else if (at <= high) {
                nextStart = low;
                nextEnd = high + count;
            } else {
                nextStart = low;
                nextEnd = high;
            }
        } else {
            long deletedEnd = (long) at + count - 1;
            if (high < at) {
                nextStart = low;
                nextEnd = high;
            } else if (low > deletedEnd) {
                nextStart = low - count;
                nextEnd = high - count;
            } else {
                nextStart = low < at ? low : at;
                nextEnd = high > deletedEnd ? high - count : (long) at - 1;
            }
        }
        if (nextStart > nextEnd || nextStart < 0 || nextEnd > Integer.MAX_VALUE) return null;
        return new int[]{(int) nextStart, (int) nextEnd};
    }

    private static int mapAxisPoint(int position, Axis axis, int at, int count, Direction direction) {
        if (direction == Direction.INSERT) {
            long next = position < at ? position : (long) position + count;
            int maximum = axis == Axis.ROW ? MAX_ROW : MAX_COLUMN;
            return next > maximum ? -1 : (int) next;
        }
        long end = (long) at + count - 1;
        if (position < at) return position;
        if (position > end) return (int) (position - (long) count);
        return -1;
    }

    private static Reference withAxisCoordinate(Reference reference, Axis axis, int coordinate) {
        return axis == Axis.ROW ? reference.withRow(coordinate) : reference.withColumn(coordinate);
    }

    private static boolean hasDifferentSheetEndpoints(ParsedReference parsed) {
        return parsed.start().sheetName() != null && parsed.end().sheetName() != null
                && !sameName(parsed.start().sheetName(), parsed.end().sheetName());
    }

    private static void assertNoThreeDimensionalReference(String formula) {
        if (formula == null) return;
        int index = 0;
        while (index < formula.length()) {
            if (formula.charAt(index) == '"') {
                index = consumeString(formula, index);
                continue;
            }
            if (formula.charAt(index) == '[') {
                int externalEnd = consumeExternalReference(formula, index);
                index = externalEnd > index ? externalEnd : consumeBracketedReference(formula, index);
                continue;
            }
            SheetPrefix first = parseSheetPrefix(formula, index);
            if (first != null && first.name().contains(":")
                    && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == '!') {
                throw ServiceException.unavailable("UNSUPPORTED_FEATURE: quoted 3-D references require an ordered worksheet transform");
            }
            if (first != null && first.afterPrefix() < formula.length() && formula.charAt(first.afterPrefix()) == ':') {
                int secondStart = first.afterPrefix() + 1;
                SheetPrefix second = secondStart < formula.length() ? parseSheetPrefix(formula, secondStart) : null;
                if (second != null && second.afterPrefix() < formula.length() && formula.charAt(second.afterPrefix()) == '!') {
                    throw ServiceException.unavailable("UNSUPPORTED_FEATURE: 3-D references require an ordered worksheet transform");
                }
            }
            index = nextReferenceCandidate(formula, index, first);
        }
    }

    private static WholeAxisReference parseWholeAxisReference(String formula, int start) {
        if (start < 0 || start >= formula.length()) return null;
        if (start > 0 && isReferenceNamePart(formula.charAt(start - 1))) return null;

        int cursor = start;
        if (formula.charAt(cursor) == '$') cursor += 1;
        int firstStart = cursor;
        if (cursor < formula.length() && isAsciiLetter(formula.charAt(cursor))) {
            while (cursor < formula.length() && isAsciiLetter(formula.charAt(cursor))) cursor += 1;
            if (cursor - firstStart > 3 || columnIndex(formula, firstStart, cursor) > MAX_COLUMN
                    || cursor >= formula.length() || formula.charAt(cursor) != ':') return null;
            int firstEnd = cursor;
            int startColumn = columnIndex(formula, firstStart, firstEnd);
            cursor += 1;
            if (cursor < formula.length() && formula.charAt(cursor) == '$') cursor += 1;
            int secondStart = cursor;
            while (cursor < formula.length() && isAsciiLetter(formula.charAt(cursor))) cursor += 1;
            if (secondStart == cursor || cursor - secondStart > 3 || columnIndex(formula, secondStart, cursor) > MAX_COLUMN
                    || (cursor < formula.length() && isReferenceNamePart(formula.charAt(cursor)))) return null;
            return new WholeAxisReference(Axis.COLUMN, startColumn, columnIndex(formula, secondStart, cursor),
                    start, firstStart, firstEnd, secondStart, cursor, cursor);
        }

        cursor = start;
        if (formula.charAt(cursor) == '$') cursor += 1;
        int rowStart = cursor;
        while (cursor < formula.length() && Character.isDigit(formula.charAt(cursor))) cursor += 1;
        if (rowStart == cursor || cursor >= formula.length() || formula.charAt(cursor) != ':') return null;
        if (!validWholeRowIndex(formula, rowStart, cursor)) return null;
        int firstEnd = cursor;
        int startRow = (int) Long.parseLong(formula.substring(rowStart, firstEnd)) - 1;
        cursor += 1;
        if (cursor < formula.length() && formula.charAt(cursor) == '$') cursor += 1;
        int secondStart = cursor;
        while (cursor < formula.length() && Character.isDigit(formula.charAt(cursor))) cursor += 1;
        if (secondStart == cursor || !validWholeRowIndex(formula, secondStart, cursor)
                || (cursor < formula.length() && isReferenceNamePart(formula.charAt(cursor)))) return null;
        int endRow = (int) Long.parseLong(formula.substring(secondStart, cursor)) - 1;
        return new WholeAxisReference(Axis.ROW, startRow, endRow, start, rowStart, firstEnd, secondStart, cursor, cursor);
    }

    private static boolean validWholeRowIndex(String formula, int start, int end) {
        try {
            long row = Long.parseLong(formula.substring(start, end));
            return row >= 1 && row <= MAX_ROW + 1L;
        } catch (NumberFormatException ignored) {
            return false;
        }
    }

    private static boolean isReferenceNamePart(char value) {
        return isAsciiLetter(value) || Character.isDigit(value) || value == '_' || value == '.';
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

    private static int nextReferenceCandidate(String formula, int start, SheetPrefix prefix) {
        if (prefix != null) return Math.max(start + 1, prefix.afterPrefix());
        char current = formula.charAt(start);
        if (current == '\'') {
            int index = start + 1;
            while (index < formula.length()) {
                if (formula.charAt(index) != '\'') {
                    index += 1;
                } else if (index + 1 < formula.length() && formula.charAt(index + 1) == '\'') {
                    index += 2;
                } else {
                    return index + 1;
                }
            }
            return formula.length();
        }
        if (!isSheetIdentifierStart(current)) return start + 1;
        int index = start + 1;
        while (index < formula.length() && isSheetIdentifierPart(formula.charAt(index))) index += 1;
        return index;
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

    private interface ReferenceRangeMapper {
        RangeMapping map(ParsedReference parsed);
    }

    private record RangeMapping(boolean handled, Reference start, Reference end) {
        static RangeMapping notHandled() { return new RangeMapping(false, null, null); }
        static RangeMapping handled(Reference start, Reference end) { return new RangeMapping(true, start, end); }
    }

    private record Rectangle(int startRow, int endRow, int startColumn, int endColumn) {
    }

    private record WholeAxisReference(Axis axis, int start, int end, int startIndex,
            int firstCoordinateStart, int firstCoordinateEnd,
            int secondCoordinateStart, int secondCoordinateEnd, int endIndex) {
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

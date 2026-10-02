package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.RangeAccessRegion;
import com.xc.luckysheet.server.contract.RangeRef;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/** Immutable per-workbook spatial index. Queries prune by bounding boxes before testing regions. */
public final class RangeAccessIndex {
    private static final int LEAF_SIZE = 8;
    private final Map<String, Node> roots;

    public RangeAccessIndex(List<RangeAccessRegion> regions) {
        Map<String, List<RangeAccessRegion>> bySheet = new HashMap<>();
        for (RangeAccessRegion region : regions) bySheet.computeIfAbsent(region.sheetId(), ignored -> new ArrayList<>()).add(region);
        Map<String, Node> indexed = new HashMap<>();
        bySheet.forEach((sheetId, sheetRegions) -> indexed.put(sheetId, build(sheetRegions)));
        roots = Map.copyOf(indexed);
    }

    public List<RangeAccessRegion> intersecting(RangeRef range) {
        Node root = roots.get(range.sheetId());
        if (root == null) return List.of();
        List<RangeAccessRegion> result = new ArrayList<>();
        query(root, range, result);
        return List.copyOf(result);
    }

    private static Node build(List<RangeAccessRegion> source) {
        int minRow = source.stream().mapToInt(region -> region.range().startRow()).min().orElse(0);
        int maxRow = source.stream().mapToInt(region -> region.range().endRow()).max().orElse(0);
        int minColumn = source.stream().mapToInt(region -> region.range().startColumn()).min().orElse(0);
        int maxColumn = source.stream().mapToInt(region -> region.range().endColumn()).max().orElse(0);
        if (source.size() <= LEAF_SIZE) return new Node(minRow, maxRow, minColumn, maxColumn, List.copyOf(source), null, null);
        boolean splitRows = (long) maxRow - minRow >= (long) maxColumn - minColumn;
        List<RangeAccessRegion> ordered = new ArrayList<>(source);
        ordered.sort(Comparator.comparingLong(region -> splitRows
                ? (long) region.range().startRow() + region.range().endRow()
                : (long) region.range().startColumn() + region.range().endColumn()));
        int middle = ordered.size() / 2;
        return new Node(minRow, maxRow, minColumn, maxColumn, List.of(),
                build(ordered.subList(0, middle)), build(ordered.subList(middle, ordered.size())));
    }

    private static void query(Node node, RangeRef range, List<RangeAccessRegion> target) {
        if (!intersects(node.minRow, node.maxRow, node.minColumn, node.maxColumn, range)) return;
        if (!node.regions.isEmpty()) {
            for (RangeAccessRegion region : node.regions) if (intersects(region.range(), range)) target.add(region);
            return;
        }
        if (node.left != null) query(node.left, range, target);
        if (node.right != null) query(node.right, range, target);
    }

    public static boolean intersects(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId())
                && left.startRow() <= right.endRow() && right.startRow() <= left.endRow()
                && left.startColumn() <= right.endColumn() && right.startColumn() <= left.endColumn();
    }

    private static boolean intersects(int minRow, int maxRow, int minColumn, int maxColumn, RangeRef range) {
        return minRow <= range.endRow() && range.startRow() <= maxRow
                && minColumn <= range.endColumn() && range.startColumn() <= maxColumn;
    }

    private record Node(int minRow, int maxRow, int minColumn, int maxColumn,
                        List<RangeAccessRegion> regions, Node left, Node right) {
    }
}

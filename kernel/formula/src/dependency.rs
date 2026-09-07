use super::*;
/// Dyadic row intervals: point changes probe 21 buckets, independent of cell count.
#[derive(Clone, Default)]
pub(crate) struct DependencyIndex {
    ranges: BTreeMap<CellAddress, Vec<RangeRef>>,
    point_sets: BTreeMap<CellAddress, BTreeSet<CellAddress>>,
    reverse: BTreeMap<CellAddress, BTreeSet<CellAddress>>,
    buckets: BTreeMap<(String, u32, u32), BTreeMap<CellAddress, Vec<(u32, u32)>>>,
    names: BTreeMap<String, BTreeSet<CellAddress>>,
    owner_names: BTreeMap<CellAddress, BTreeSet<String>>,
}
fn cover(mut start: u32, end: u32) -> Vec<(u32, u32)> {
    let mut out = Vec::new();
    while start <= end {
        let remaining = end - start + 1;
        let fit = 31 - remaining.leading_zeros();
        let level = if start == 0 {
            fit
        } else {
            start.trailing_zeros().min(fit)
        };
        out.push((level, start >> level));
        start += 1 << level;
    }
    out
}
impl DependencyIndex {
    pub fn remove(&mut self, owner: &CellAddress) {
        if let Some(ranges) = self.ranges.remove(owner) {
            for r in ranges {
                if r.start_row == r.end_row && r.start_column == r.end_column {
                    let a = CellAddress {
                        sheet_id: r.sheet_id.clone(),
                        row: r.start_row,
                        column: r.start_column,
                    };
                    if let Some(v) = self.reverse.get_mut(&a) {
                        v.remove(owner);
                        if v.is_empty() {
                            self.reverse.remove(&a);
                        }
                    }
                } else {
                    for (level, index) in cover(r.start_row, r.end_row) {
                        let key = (r.sheet_id.clone(), level, index);
                        if let Some(v) = self.buckets.get_mut(&key) {
                            v.remove(owner);
                            if v.is_empty() {
                                self.buckets.remove(&key);
                            }
                        }
                    }
                }
            }
        }
        self.point_sets.remove(owner);
        if let Some(names) = self.owner_names.remove(owner) {
            for name in names {
                if let Some(v) = self.names.get_mut(&name) {
                    v.remove(owner);
                }
            }
        }
    }
    fn insert_range(&mut self, owner: &CellAddress, r: RangeRef) {
        if r.start_row == r.end_row && r.start_column == r.end_column {
            let a = CellAddress {
                sheet_id: r.sheet_id.clone(),
                row: r.start_row,
                column: r.start_column,
            };
            self.reverse
                .entry(a.clone())
                .or_default()
                .insert(owner.clone());
            self.point_sets.entry(owner.clone()).or_default().insert(a);
        } else {
            for (level, index) in cover(r.start_row, r.end_row) {
                self.buckets
                    .entry((r.sheet_id.clone(), level, index))
                    .or_default()
                    .entry(owner.clone())
                    .or_default()
                    .push((r.start_column, r.end_column));
            }
        }
        self.ranges.entry(owner.clone()).or_default().push(r);
    }
    pub fn replace(&mut self, owner: &CellAddress, expr: &Expr) {
        self.remove(owner);
        let mut ranges = Vec::new();
        let mut names = BTreeSet::new();
        collect(expr, &mut ranges, &mut names);
        for r in ranges {
            self.insert_range(owner, r);
        }
        for n in &names {
            self.names
                .entry(n.clone())
                .or_default()
                .insert(owner.clone());
        }
        self.owner_names.insert(owner.clone(), names);
        self.point_sets.entry(owner.clone()).or_default();
    }
    pub fn add_dynamic(&mut self, owner: &CellAddress, ranges: Vec<RangeRef>) {
        for range in ranges {
            if !self.ranges.get(owner).is_some_and(|rs| rs.contains(&range)) {
                self.insert_range(owner, range);
            }
        }
    }
    pub fn affected(&self, a: &CellAddress) -> BTreeSet<CellAddress> {
        let mut out = self.reverse.get(a).cloned().unwrap_or_default();
        for level in 0..=20 {
            if let Some(bucket) = self
                .buckets
                .get(&(a.sheet_id.clone(), level, a.row >> level))
            {
                for (owner, columns) in bucket {
                    if columns
                        .iter()
                        .any(|(s, e)| a.column >= *s && a.column <= *e)
                    {
                        out.insert(owner.clone());
                    }
                }
            }
        }
        out
    }
    pub fn affected_range(&self, r: &RangeRef) -> BTreeSet<CellAddress> {
        if r.start_row == r.end_row && r.start_column == r.end_column {
            return self.affected(&CellAddress {
                sheet_id: r.sheet_id.clone(),
                row: r.start_row,
                column: r.start_column,
            });
        }
        let mut out = BTreeSet::new();
        let lo = CellAddress {
            sheet_id: r.sheet_id.clone(),
            row: r.start_row,
            column: 0,
        };
        let hi = CellAddress {
            sheet_id: r.sheet_id.clone(),
            row: r.end_row,
            column: MAX_COLUMNS - 1,
        };
        for (a, owners) in self.reverse.range(lo..=hi) {
            if r.contains(a) {
                out.extend(owners.iter().cloned());
            }
        }
        for level in 0..=20 {
            let first = (r.sheet_id.clone(), level, r.start_row >> level);
            let last = (r.sheet_id.clone(), level, r.end_row >> level);
            for (_, owners) in self.buckets.range(first..=last) {
                for (owner, columns) in owners {
                    if columns
                        .iter()
                        .any(|(s, e)| *s <= r.end_column && *e >= r.start_column)
                    {
                        out.insert(owner.clone());
                    }
                }
            }
        }
        out
    }
    pub fn named(&self, n: &str) -> BTreeSet<CellAddress> {
        self.names.get(n).cloned().unwrap_or_default()
    }
    pub fn points(&self, a: &CellAddress) -> Option<&BTreeSet<CellAddress>> {
        self.point_sets.get(a)
    }
}
fn collect(e: &Expr, ranges: &mut Vec<RangeRef>, names: &mut BTreeSet<String>) {
    match e {
        Expr::Reference(r) => ranges.push(r.clone()),
        Expr::Name(n) | Expr::Structured(n) => {
            names.insert(n.to_uppercase());
        }
        Expr::Array(rows) => {
            for e in rows.iter().flatten() {
                collect(e, ranges, names)
            }
        }
        Expr::Unary(_, e) | Expr::Spill(e) => collect(e, ranges, names),
        Expr::Binary(_, a, b) => {
            collect(a, ranges, names);
            collect(b, ranges, names)
        }
        Expr::Call(_, args) => {
            for e in args {
                collect(e, ranges, names)
            }
        }
        Expr::Invoke(e, args) => {
            collect(e, ranges, names);
            for e in args {
                collect(e, ranges, names)
            }
        }
        _ => {}
    }
}
pub fn is_volatile(e: &Expr) -> bool {
    match e {
        Expr::Call(n, args) => {
            matches!(
                n.as_str(),
                "OFFSET" | "INDIRECT" | "NOW" | "TODAY" | "RAND" | "RANDBETWEEN" | "RANDARRAY"
            ) || args.iter().any(is_volatile)
        }
        Expr::Array(r) => r.iter().flatten().any(is_volatile),
        Expr::Unary(_, e) | Expr::Spill(e) => is_volatile(e),
        Expr::Binary(_, a, b) => is_volatile(a) || is_volatile(b),
        Expr::Invoke(e, args) => is_volatile(e) || args.iter().any(is_volatile),
        _ => false,
    }
}

use kernel_core::Scalar;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum AggregateKind {
    Sum,
    Count,
    CountNumbers,
    Average,
    Min,
    Max,
    Product,
    Stdev,
    Stdevp,
    Var,
    Varp,
    DistinctCount,
}

/// Mergeable states keep the original typed numeric domain. Text is never parsed.
#[derive(Debug, Clone)]
pub(crate) struct AggregateState {
    count: u64,
    numbers: u64,
    sum: f64,
    compensation: f64,
    product: f64,
    min: f64,
    max: f64,
    mean: f64,
    m2: f64,
    error: Option<Scalar>,
    distinct: Option<HashSet<String>>,
    distinct_bytes: u64,
}
impl AggregateState {
    pub fn new(kind: AggregateKind) -> Self {
        Self {
            count: 0,
            numbers: 0,
            sum: 0.,
            compensation: 0.,
            product: 1.,
            min: f64::INFINITY,
            max: f64::NEG_INFINITY,
            mean: 0.,
            m2: 0.,
            error: None,
            distinct: (kind == AggregateKind::DistinctCount).then(HashSet::new),
            distinct_bytes: 0,
        }
    }
    pub fn add(&mut self, value: &Scalar) {
        if !matches!(value, Scalar::Null) && !matches!(value, Scalar::Text(s) if s.is_empty()) {
            self.count += 1;
            if let Some(distinct) = &mut self.distinct {
                let key = scalar_key(value);
                let bytes = key.len() as u64 + 40;
                if distinct.insert(key) {
                    self.distinct_bytes += bytes;
                }
            }
        }
        match value {
            Scalar::Number(n) => {
                self.numbers += 1;
                let compensated = n - self.compensation;
                let sum = self.sum + compensated;
                self.compensation = (sum - self.sum) - compensated;
                self.sum = sum;
                self.product *= n;
                self.min = self.min.min(*n);
                self.max = self.max.max(*n);
                let delta = n - self.mean;
                self.mean += delta / self.numbers as f64;
                self.m2 += delta * (n - self.mean);
            }
            Scalar::Error(_) if self.error.is_none() => self.error = Some(value.clone()),
            _ => {}
        }
    }
    pub fn merge(&mut self, other: &Self) {
        self.count += other.count;
        if self.error.is_none() {
            self.error = other.error.clone();
        }
        if let (Some(left), Some(right)) = (&mut self.distinct, &other.distinct) {
            for key in right {
                if left.insert(key.clone()) {
                    self.distinct_bytes += key.len() as u64 + 40;
                }
            }
        }
        if other.numbers == 0 {
            return;
        }
        let total = self.numbers + other.numbers;
        let delta = other.mean - self.mean;
        self.m2 +=
            other.m2 + delta * delta * self.numbers as f64 * other.numbers as f64 / total as f64;
        self.mean += delta * other.numbers as f64 / total as f64;
        self.numbers = total;
        let compensated = other.sum - self.compensation;
        let sum = self.sum + compensated;
        self.compensation = (sum - self.sum) - compensated;
        self.sum = sum;
        self.product *= other.product;
        self.min = self.min.min(other.min);
        self.max = self.max.max(other.max);
    }
    pub fn value(&self, kind: AggregateKind) -> Scalar {
        let number = match kind {
            AggregateKind::Count => return Scalar::Number(self.count as f64),
            AggregateKind::CountNumbers => return Scalar::Number(self.numbers as f64),
            AggregateKind::DistinctCount => {
                return Scalar::Number(self.distinct.as_ref().map_or(0, HashSet::len) as f64);
            }
            _ if self.error.is_some() => return self.error.clone().expect("checked error"),
            AggregateKind::Sum => Some(self.sum),
            AggregateKind::Average => (self.numbers > 0).then_some(self.mean),
            AggregateKind::Min => (self.numbers > 0).then_some(self.min),
            AggregateKind::Max => (self.numbers > 0).then_some(self.max),
            AggregateKind::Product => (self.numbers > 0).then_some(self.product),
            AggregateKind::Stdev => {
                (self.numbers > 1).then(|| (self.m2.max(0.) / (self.numbers - 1) as f64).sqrt())
            }
            AggregateKind::Stdevp => {
                (self.numbers > 0).then(|| (self.m2.max(0.) / self.numbers as f64).sqrt())
            }
            AggregateKind::Var => {
                (self.numbers > 1).then(|| self.m2.max(0.) / (self.numbers - 1) as f64)
            }
            AggregateKind::Varp => {
                (self.numbers > 0).then(|| self.m2.max(0.) / self.numbers as f64)
            }
        };
        match number {
            Some(n) if n.is_finite() => Scalar::Number(n),
            Some(_) => Scalar::error("#NUM!", "Pivot aggregate overflow"),
            None => Scalar::Null,
        }
    }
    pub fn bytes(&self) -> u64 {
        std::mem::size_of::<Self>() as u64 + self.distinct_bytes
    }
}
pub(crate) fn scalar_key(value: &Scalar) -> String {
    match value {
        Scalar::Null => "z".into(),
        Scalar::Boolean(b) => format!("b{b}"),
        Scalar::Number(n) => format!("n{}", if *n == 0. { 0 } else { n.to_bits() }),
        Scalar::Text(s) => format!("s{s}"),
        Scalar::Error(e) => format!("e{}", e.code),
    }
}
pub(crate) fn compare_scalar(left: &Scalar, right: &Scalar) -> std::cmp::Ordering {
    fn rank(v: &Scalar) -> u8 {
        match v {
            Scalar::Number(_) => 0,
            Scalar::Text(_) => 1,
            Scalar::Boolean(_) => 2,
            Scalar::Error(_) => 3,
            Scalar::Null => 4,
        }
    }
    rank(left)
        .cmp(&rank(right))
        .then_with(|| match (left, right) {
            (Scalar::Number(a), Scalar::Number(b)) => a.total_cmp(b),
            (Scalar::Text(a), Scalar::Text(b)) => a.cmp(b),
            (Scalar::Boolean(a), Scalar::Boolean(b)) => a.cmp(b),
            (Scalar::Error(a), Scalar::Error(b)) => a.code.cmp(&b.code),
            _ => std::cmp::Ordering::Equal,
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn typed_aggregate_and_merge() {
        let mut a = AggregateState::new(AggregateKind::Average);
        for v in [Scalar::Number(2.), Scalar::Text("900".into()), Scalar::Null] {
            a.add(&v);
        }
        let mut b = AggregateState::new(AggregateKind::Average);
        b.add(&Scalar::Number(4.));
        a.merge(&b);
        assert_eq!(a.value(AggregateKind::Average), Scalar::Number(3.));
        assert_eq!(a.value(AggregateKind::Count), Scalar::Number(3.));
        assert_eq!(a.value(AggregateKind::Var), Scalar::Number(2.));
    }
    #[test]
    fn error_propagates_only_numeric_aggregates() {
        let mut a = AggregateState::new(AggregateKind::Sum);
        a.add(&Scalar::error("#REF!", "bad"));
        assert!(matches!(a.value(AggregateKind::Sum), Scalar::Error(_)));
        assert_eq!(a.value(AggregateKind::Count), Scalar::Number(1.));
    }
}

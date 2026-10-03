import { coerceExcelNumber, normalizeExcelPrecision } from '../numeric';
import { createFormulaError, isFormulaError, type FormulaError, type FormulaValue } from '../values';

type FinancialExpression = (args: readonly number[]) => number | FormulaError;
const num = (message: string) => createFormulaError('#NUM!', message);
const zero = (message: string) => createFormulaError('#DIV/0!', message);

/** Scalar coercion and result precision remain owned by the shared numeric domain. */
function financial(expression: FinancialExpression): (args: FormulaValue[]) => FormulaValue {
  return args => {
    const numbers: number[] = [];
    for (const value of args) {
      const number = coerceExcelNumber(value);
      if (isFormulaError(number)) return number;
      numbers.push(number);
    }
    const result = expression(numbers);
    if (isFormulaError(result)) return result;
    if (!Number.isFinite(result)) return num('Financial result is not a finite real number');
    return result === 0 ? 0 : normalizeExcelPrecision(result);
  };
}

/** Periodic growth, including real integer powers for negative bases. */
function growth(rate: number, periods: number): number {
  return rate > -1 ? Math.exp(periods * Math.log1p(rate)) : Math.pow(1 + rate, periods);
}
function annuity(rate: number, periods: number): number {
  if (rate === 0) return periods;
  return rate > -1 ? Math.expm1(periods * Math.log1p(rate)) / rate : (growth(rate, periods) - 1) / rate;
}
function discountedAnnuity(rate: number, periods: number): number {
  if (rate === 0) return periods;
  return rate > -1 ? -Math.expm1(-periods * Math.log1p(rate)) / rate : (1 - growth(rate, -periods)) / rate;
}
function paymentType(type: number): FormulaError | undefined {
  return type === 0 || type === 1 ? undefined : num('Payment type must be 0 (end) or 1 (beginning)');
}

function futureValue(rate: number, periods: number, payment: number, present: number, type: number): number {
  return -(present * growth(rate, periods) + payment * (type === 1 ? 1 + rate : 1) * annuity(rate, periods));
}
function payment(rate: number, periods: number, present: number, future: number, type: number): number | FormulaError {
  const factor = type === 1 ? 1 + rate : 1;
  // Discounting keeps ordinary long positive-rate loans finite even when forward growth overflows.
  const discounted = rate > 0 && periods >= 0;
  const denominator = factor * (discounted ? discountedAnnuity(rate, periods) : annuity(rate, periods));
  if (denominator === 0) return zero('Payment requires a nonzero annuity factor');
  const numerator = discounted
    ? present + future * growth(rate, -periods)
    : future + present * growth(rate, periods);
  return -numerator / denominator;
}
function periodPayment(args: readonly number[], principal: boolean): number | FormulaError {
  const [rate = 0, period = 0, periods = 0, present = 0, future = 0, type = 0] = args;
  const typeError = paymentType(type);
  if (typeError) return typeError;
  if (period < 1 || period > periods) return num('Payment period must be between 1 and nper');
  const pmt = payment(rate, periods, present, future, type);
  if (isFormulaError(pmt)) return pmt;
  const factor = type === 1 ? 1 + rate : 1;
  const interest = rate === 0 || (type === 1 && period === 1) ? 0
    : futureValue(rate, period - 1, pmt, present, type) * rate / factor;
  return principal ? pmt - interest : interest;
}
function dollar(args: readonly number[], decimal: boolean): number | FormulaError {
  const [amount = 0, fraction = 0] = args;
  if (fraction < 0) return num('Dollar fraction cannot be negative');
  const denominator = Math.trunc(fraction);
  if (denominator === 0) return zero('Dollar fraction must be at least 1');
  const scale = 10 ** Math.ceil(Math.log10(denominator));
  const whole = Math.trunc(amount);
  return whole + (amount - whole) * (decimal ? scale / denominator : denominator / scale);
}

export const financialFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  PV: financial(([rate = 0, periods = 0, pmt = 0, future = 0, type = 0]) => {
    const error = paymentType(type);
    if (error) return error;
    if (rate === -1 && periods > 0) return zero('Present value cannot discount a zero growth factor');
    return -(future * growth(rate, -periods) + pmt * (type === 1 ? 1 + rate : 1) * discountedAnnuity(rate, periods));
  }),
  FV: financial(([rate = 0, periods = 0, pmt = 0, present = 0, type = 0]) =>
    paymentType(type) ?? futureValue(rate, periods, pmt, present, type)),
  PMT: financial(([rate = 0, periods = 0, present = 0, future = 0, type = 0]) =>
    paymentType(type) ?? payment(rate, periods, present, future, type)),
  NPER: financial(([rate = 0, pmt = 0, present = 0, future = 0, type = 0]) => {
    const error = paymentType(type);
    if (error) return error;
    if (rate === 0) return pmt === 0 ? zero('Zero-rate periods require a nonzero payment') : -(present + future) / pmt;
    if (rate <= -1) return num('NPER requires a positive periodic growth base');
    const denominator = present * rate + pmt * (type === 1 ? 1 + rate : 1);
    if (denominator === 0) return num('Cash flows do not determine a finite number of periods');
    return Math.log1p(-rate * (present + future) / denominator) / Math.log1p(rate);
  }),
  IPMT: financial(args => periodPayment(args, false)),
  PPMT: financial(args => periodPayment(args, true)),
  ISPMT: financial(([rate = 0, period = 0, periods = 0, present = 0]) =>
    periods === 0 ? zero('ISPMT nper cannot be zero') : present * rate * (period / periods - 1)),
  EFFECT: financial(([rate = 0, periods = 0]) => {
    const count = Math.trunc(periods);
    return rate <= 0 || count < 1 ? num('EFFECT requires positive rate and at least one compounding period')
      : Math.expm1(count * Math.log1p(rate / count));
  }),
  NOMINAL: financial(([rate = 0, periods = 0]) => {
    const count = Math.trunc(periods);
    return rate <= 0 || count < 1 ? num('NOMINAL requires positive rate and at least one compounding period')
      : count * Math.expm1(Math.log1p(rate) / count);
  }),
  SLN: financial(([cost = 0, salvage = 0, life = 0]) =>
    life === 0 ? zero('SLN life cannot be zero') : (cost - salvage) / life),
  DOLLARDE: financial(args => dollar(args, true)),
  DOLLARFR: financial(args => dollar(args, false)),
};

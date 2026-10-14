import type { BaseDateRule, DateRangeRule } from './types';

export interface DateRange {
  start: string;
  end: string;
}

export function resolveBaseDate(rule: BaseDateRule, currentDate = new Date(), currentValue = ''): string {
  const today = localDate(currentDate);

  if (rule === 'yesterday') {
    today.setDate(today.getDate() - 1);
  }

  return formatDate(today);
}

export function calculateDateRange(rule: DateRangeRule, baseDateValue: string): DateRange {
  const baseDate = parseDate(baseDateValue);
  const range = (start: Date, end = start): DateRange => ({ start: formatDate(start), end: formatDate(end) });

  if (rule === 'previousDay') {
    return range(addDays(baseDate, -1));
  }

  if (rule === 'nextDay') {
    return range(addDays(baseDate, 1));
  }

  if (rule === 'baseWeek') {
    return weekRange(baseDate);
  }

  if (rule === 'previousWeek') {
    return weekRange(addDays(baseDate, -7));
  }

  if (rule === 'nextWeek') {
    return weekRange(addDays(baseDate, 7));
  }

  if (rule === 'monthStartToBaseDate') {
    const start = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
    return { start: formatDate(start), end: baseDateValue };
  }

  if (rule === 'yearStartToBaseDate') {
    const start = new Date(baseDate.getFullYear(), 0, 1);
    return { start: formatDate(start), end: baseDateValue };
  }

  if (rule === 'baseMonthStart') {
    return range(monthStart(baseDate));
  }

  if (rule === 'baseMonthEnd') {
    return range(monthEnd(baseDate));
  }

  if (rule === 'baseMonth') {
    return range(monthStart(baseDate), monthEnd(baseDate));
  }

  if (rule === 'previousMonthStart') {
    return range(monthStart(addMonths(baseDate, -1)));
  }

  if (rule === 'previousMonthEnd') {
    return range(monthEnd(addMonths(baseDate, -1)));
  }

  if (rule === 'previousMonth') {
    const previous = addMonths(baseDate, -1);
    return range(monthStart(previous), monthEnd(previous));
  }

  if (rule === 'nextMonthStart') {
    return range(monthStart(addMonths(baseDate, 1)));
  }

  if (rule === 'nextMonthEnd') {
    return range(monthEnd(addMonths(baseDate, 1)));
  }

  if (rule === 'nextMonth') {
    const next = addMonths(baseDate, 1);
    return range(monthStart(next), monthEnd(next));
  }

  if (rule === 'baseMonthToNextMonthEnd') {
    return range(monthStart(baseDate), monthEnd(addMonths(baseDate, 1)));
  }

  if (rule === 'baseYearStart') {
    return range(yearStart(baseDate));
  }

  if (rule === 'baseYearEnd') {
    return range(yearEnd(baseDate));
  }

  if (rule === 'baseYear') {
    return range(yearStart(baseDate), yearEnd(baseDate));
  }

  if (rule === 'baseFirstHalf') {
    return range(new Date(baseDate.getFullYear(), 0, 1), new Date(baseDate.getFullYear(), 5, 30));
  }

  if (rule === 'baseSecondHalf') {
    return range(new Date(baseDate.getFullYear(), 6, 1), new Date(baseDate.getFullYear(), 11, 31));
  }

  if (rule === 'previousYearStart') {
    return range(yearStart(addYears(baseDate, -1)));
  }

  if (rule === 'previousYearEnd') {
    return range(yearEnd(addYears(baseDate, -1)));
  }

  if (rule === 'previousYear') {
    const previous = addYears(baseDate, -1);
    return range(yearStart(previous), yearEnd(previous));
  }

  if (rule === 'nextYearStart') {
    return range(yearStart(addYears(baseDate, 1)));
  }

  if (rule === 'nextYearEnd') {
    return range(yearEnd(addYears(baseDate, 1)));
  }

  if (rule === 'nextYear') {
    const next = addYears(baseDate, 1);
    return range(yearStart(next), yearEnd(next));
  }

  if (rule === 'sameDayPreviousYear') {
    return range(addYears(baseDate, -1));
  }

  if (rule === 'sameMonthPreviousYear') {
    const previousYear = addYears(baseDate, -1);
    return range(monthStart(previousYear), monthEnd(previousYear));
  }

  if (rule === 'previousMonthPreviousYear') {
    const previousMonthPreviousYear = addYears(addMonths(baseDate, -1), -1);
    return range(monthStart(previousMonthPreviousYear), monthEnd(previousMonthPreviousYear));
  }

  return { start: baseDateValue, end: baseDateValue };
}

export function formatDate(value: Date): string {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, '0');
  const day = String(value.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// 出力日時の表示用文字列（ローカル時刻、タイムゾーンなし）。設定!B7・出力条件メモ・文字列フィールドに使う。
export function formatLocalDateTime(value: Date): string {
  const hours = String(value.getHours()).padStart(2, '0');
  const minutes = String(value.getMinutes()).padStart(2, '0');
  const seconds = String(value.getSeconds()).padStart(2, '0');
  return `${formatDate(value)}T${hours}:${minutes}:${seconds}`;
}

// Kintoneの日時フィールドへ登録する値。REST APIはタイムゾーン（Zまたは±HH:MM）付きのISO 8601を要求する。
export function formatKintoneDateTime(value: Date): string {
  return value.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// 出力日時を書き込む先のフィールド型に合わせた値。型が分からない場合は従来どおりローカル時刻の文字列。
export function exportedAtFieldValue(value: Date, fieldType: string | undefined): string {
  if (fieldType === 'DATETIME') {
    return formatKintoneDateTime(value);
  }
  if (fieldType === 'DATE') {
    return formatDate(value);
  }
  return formatLocalDateTime(value);
}

export function parseDate(value: string): Date {
  const [year, month, day] = value.split('-').map(Number);
  return new Date(year, month - 1, day);
}

function localDate(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function addDays(value: Date, amount: number): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate() + amount);
}

function addMonths(value: Date, amount: number): Date {
  return new Date(value.getFullYear(), value.getMonth() + amount, 1);
}

function addYears(value: Date, amount: number): Date {
  const targetMonthStart = new Date(value.getFullYear() + amount, value.getMonth(), 1);
  const lastDay = monthEnd(targetMonthStart).getDate();
  return new Date(value.getFullYear() + amount, value.getMonth(), Math.min(value.getDate(), lastDay));
}

function weekRange(value: Date): DateRange {
  const start = addDays(value, -value.getDay());
  const end = addDays(start, 6);
  return { start: formatDate(start), end: formatDate(end) };
}

function monthStart(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), 1);
}

function monthEnd(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth() + 1, 0);
}

function yearStart(value: Date): Date {
  return new Date(value.getFullYear(), 0, 1);
}

function yearEnd(value: Date): Date {
  return new Date(value.getFullYear(), 11, 31);
}

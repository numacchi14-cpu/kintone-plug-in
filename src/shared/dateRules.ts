import type { BaseDateRule, DateRangeRule } from './types';

export interface DateRange {
  start: string;
  end: string;
}

export function resolveBaseDate(rule: BaseDateRule, currentDate = new Date(), currentValue = ''): string {
  if (currentValue) {
    return currentValue;
  }

  if (rule === 'manual') {
    return '';
  }

  const today = localDate(currentDate);

  if (rule === 'yesterday' || (rule === 'firstDayUsesYesterday' && today.getDate() === 1)) {
    today.setDate(today.getDate() - 1);
  }

  return formatDate(today);
}

export function calculateDateRange(rule: DateRangeRule, baseDateValue: string): DateRange {
  const baseDate = parseDate(baseDateValue);

  if (rule === 'monthStartToBaseDate') {
    const start = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
    return { start: formatDate(start), end: baseDateValue };
  }

  if (rule === 'yearStartToBaseDate') {
    const start = new Date(baseDate.getFullYear(), 0, 1);
    return { start: formatDate(start), end: baseDateValue };
  }

  if (rule === 'baseMonth') {
    const start = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
    const end = new Date(baseDate.getFullYear(), baseDate.getMonth() + 1, 0);
    return { start: formatDate(start), end: formatDate(end) };
  }

  if (rule === 'previousMonth') {
    const start = new Date(baseDate.getFullYear(), baseDate.getMonth() - 1, 1);
    const end = new Date(baseDate.getFullYear(), baseDate.getMonth(), 0);
    return { start: formatDate(start), end: formatDate(end) };
  }

  if (rule === 'nextMonth') {
    const start = new Date(baseDate.getFullYear(), baseDate.getMonth() + 1, 1);
    const end = new Date(baseDate.getFullYear(), baseDate.getMonth() + 2, 0);
    return { start: formatDate(start), end: formatDate(end) };
  }

  if (rule === 'baseMonthToNextMonthEnd') {
    const start = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
    const end = new Date(baseDate.getFullYear(), baseDate.getMonth() + 2, 0);
    return { start: formatDate(start), end: formatDate(end) };
  }

  return { start: baseDateValue, end: baseDateValue };
}

export function formatDate(value: Date): string {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, '0');
  const day = String(value.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function parseDate(value: string): Date {
  const [year, month, day] = value.split('-').map(Number);
  return new Date(year, month - 1, day);
}

function localDate(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

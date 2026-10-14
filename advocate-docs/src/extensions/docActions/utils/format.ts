/** Replaces {0}, {1}, … in a localized string. */
export function format(template: string, ...values: Array<string | number>): string {
  return template.replace(/\{(\d+)\}/g, (match, i) => (values[Number(i)] !== undefined ? String(values[Number(i)]) : match));
}

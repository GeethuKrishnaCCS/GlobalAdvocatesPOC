import * as React from 'react';
import { Checkbox, Field, Input, makeStyles, tokens } from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { IDocumentFinder } from '../hooks/useDocumentFinder';

export interface IFilterFieldsProps {
  finder: IDocumentFinder;
}

/** File types offered as filter options. */
const FILE_TYPES: ReadonlyArray<string> = ['docx', 'xlsx', 'pptx', 'pdf', 'eml', 'msg', 'txt'];

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM },
  dates: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: tokens.spacingHorizontalM },
  types: { display: 'flex', flexWrap: 'wrap', columnGap: tokens.spacingHorizontalS }
});

/** Metadata filter: document type and date (indexed columns), plus file type. */
export const FilterFields: React.FC<IFilterFieldsProps> = ({ finder }) => {
  const styles = useStyles();
  const { filter } = finder.state;
  const set = finder.actions.setFilter;

  const toggleType = (type: string, checked: boolean): void => {
    set({ ...filter, fileTypes: checked ? [...filter.fileTypes, type] : filter.fileTypes.filter(t => t !== type) });
  };

  return (
    <div className={styles.root} data-testid="filter-fields">
      <Field label={strings.FilterDocType}>
        <Input value={filter.docType} onChange={(_, d) => set({ ...filter, docType: d.value })} data-testid="filter-fields-doc-type" />
      </Field>
      <div className={styles.dates}>
        <Field label={strings.FilterDateFrom}>
          <Input type="date" value={filter.dateFrom} onChange={(_, d) => set({ ...filter, dateFrom: d.value })}
            data-testid="filter-fields-date-from" />
        </Field>
        <Field label={strings.FilterDateTo}>
          <Input type="date" value={filter.dateTo} onChange={(_, d) => set({ ...filter, dateTo: d.value })}
            data-testid="filter-fields-date-to" />
        </Field>
      </div>
      <Field label={strings.FilterFileTypes}>
        <div className={styles.types}>
          {FILE_TYPES.map(t => (
            <Checkbox key={t} label={t} checked={filter.fileTypes.indexOf(t) >= 0}
              onChange={(_, d) => toggleType(t, d.checked === true)} data-testid="filter-fields-file-type" />
          ))}
        </div>
      </Field>
    </div>
  );
};

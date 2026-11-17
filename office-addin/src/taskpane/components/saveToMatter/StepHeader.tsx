import * as React from "react";
import { Caption1, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";

export interface IStep {
  step: string;
  label: string;
}

const SAVE_STEPS: IStep[] = [
  { step: "matter", label: "1. Matter" },
  { step: "folder", label: "2. Folder" },
  { step: "details", label: "3. Details" },
];

const useStyles = makeStyles({
  root: { display: "grid", gridAutoColumns: "1fr", gridAutoFlow: "column", gap: tokens.spacingHorizontalS },
  step: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    color: tokens.colorNeutralForeground3,
    "::before": {
      content: '""',
      height: "3px",
      borderRadius: tokens.borderRadiusSmall,
      backgroundColor: tokens.colorNeutralStroke2,
    },
  },
  reached: { color: tokens.colorNeutralForeground1, "::before": { backgroundColor: tokens.colorBrandBackground } },
});

/** "1. Matter / 2. Folder / 3. Details" progress bar ("done" = all reached); Outlook panes pass their own steps. */
export const StepHeader: React.FC<{ step: string; steps?: IStep[] }> = ({ step, steps = SAVE_STEPS }) => {
  const styles = useStyles();
  const current = step === "done" ? steps.length : steps.findIndex((s) => s.step === step);
  return (
    <div className={styles.root} role="list" aria-label="Steps" data-testid="save-steps">
      {steps.map((s, index) => (
        <Caption1
          key={s.step}
          role="listitem"
          aria-current={index === current ? "step" : undefined}
          className={mergeClasses(styles.step, index <= current && styles.reached)}
          data-testid={`save-step-${s.step}`}
        >
          {s.label}
        </Caption1>
      ))}
    </div>
  );
};

import * as React from "react";
import { Caption1, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import type { SaveStep } from "../../../shared/hooks/useSaveToMatter";

const STEPS: { step: SaveStep; label: string }[] = [
  { step: "matter", label: "1. Matter" },
  { step: "folder", label: "2. Folder" },
  { step: "details", label: "3. Details" },
];

const useStyles = makeStyles({
  root: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: tokens.spacingHorizontalS },
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

/** "1. Matter / 2. Folder / 3. Details" progress bar. */
export const StepHeader: React.FC<{ step: SaveStep }> = ({ step }) => {
  const styles = useStyles();
  const current = step === "done" ? STEPS.length : STEPS.findIndex((s) => s.step === step);
  return (
    <div className={styles.root} role="list" aria-label="Steps" data-testid="save-steps">
      {STEPS.map((s, index) => (
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

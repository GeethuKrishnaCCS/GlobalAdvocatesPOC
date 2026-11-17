import { tokens } from "@fluentui/react-components";

/** How a drop target looks while dragged emails are over it (shared by matter rows, folder rows and the open folder). */
export const dropHoverStyle = {
  outline: `${tokens.strokeWidthThick} dashed ${tokens.colorBrandStroke1}`,
  outlineOffset: "-2px",
  borderRadius: tokens.borderRadiusMedium,
  backgroundColor: tokens.colorBrandBackground2,
} as const;

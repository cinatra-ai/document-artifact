// The row icon of this package's type in the Artifacts list: the file-type icon of the shared icon
// set, drawn at 17 pixels inside the application's 34 pixel cell. The cell's size and tint stay the
// application's. It reads nothing but its props and draws the same icon for every object type.
import { FileType } from "lucide-react";

import type { ArtifactRendererProps } from "../artifact-renderer-props";

export interface ListRowProps {
  artifact: Pick<ArtifactRendererProps["artifact"], "objectType">;
}

export default function ListRow(_props: ListRowProps) {
  return <FileType aria-hidden size={17} strokeWidth={2} data-row-glyph="file-type" />;
}

const structural = new Set(["Document", "Section", "Table", "TableRow", "TableCell"]);

function node(type, props = {}) {
  const { children, ...rest } = props;
  return { __docxJsx: true, type, props: rest, children: children === undefined ? [] : children };
}

function host(type) {
  return function DocxComponent(props) {
    return node(type, props);
  };
}

export const Document = host("Document");
export const Section = host("Section");
export const Paragraph = host("Paragraph");
export const Run = host("Run");
export const Text = host("Text");
export const Break = host("Break");
export const CarriageReturn = host("CarriageReturn");
export const NonBreakingSpace = host("NonBreakingSpace");
export const SoftHyphen = host("SoftHyphen");
export const NonBreakingHyphen = host("NonBreakingHyphen");
export const Image = host("Image");
export const Table = host("Table");
export const TableRow = host("TableRow");
export const TableCell = host("TableCell");
export const Header = host("Header");
export const Footer = host("Footer");
export const Hyperlink = host("Hyperlink");
export const PageNumber = host("PageNumber");
export const TotalPages = host("TotalPages");
export const List = host("List");
export const ListItem = host("ListItem");
export const Heading = host("Heading");
export const Caption = host("Caption");
export const Index = host("Index");
export const Bookmark = host("Bookmark");
export const InlineBookmark = host("InlineBookmark");
export const TableOfContents = host("TableOfContents");
export const TableOfFigures = host("TableOfFigures");
export const TableOfEntries = host("TableOfEntries");
export const TocEntry = host("TocEntry");
export const IndexEntry = host("IndexEntry");
export const Comment = host("Comment");
export const Footnote = host("Footnote");
export const Tab = host("Tab");
export const TabStop = host("TabStop");
export const Symbol = host("Symbol");
export const Bold = host("Bold");
export const Italic = host("Italic");
export const Underline = host("Underline");
export const StrikeThrough = host("StrikeThrough");
export const Superscript = host("Superscript");
export const Subscript = host("Subscript");
export const AllCaps = host("AllCaps");
export const HiddenText = host("HiddenText");
export const SpecialHiddenText = host("SpecialHiddenText");
export const DoubleStrike = host("DoubleStrike");
export const SpacedText = host("SpacedText");
export const ScaledText = host("ScaledText");
export const FitText = host("FitText");
export const BorderedText = host("BorderedText");
export const ShadedText = host("ShadedText");
export const Inserted = host("Inserted");
export const Deleted = host("Deleted");
export const MovedFrom = host("MovedFrom");
export const MovedTo = host("MovedTo");
export const PageReference = host("PageReference");
export const PositionalTab = host("PositionalTab");
export const ContentControl = host("ContentControl");
export const Field = host("Field");
export const DateField = host("DateField");
export const TimeField = host("TimeField");
export const FileNameField = host("FileNameField");
export const AuthorField = host("AuthorField");
export const TitleField = host("TitleField");
export const SubjectField = host("SubjectField");
export const SequenceField = host("SequenceField");
export const ReferenceField = host("ReferenceField");
export const MergeField = host("MergeField");
export const DocumentPropertyField = host("DocumentPropertyField");
export const FormulaField = host("FormulaField");

export const Fragment = globalThis.Symbol.for("docx-jsx.fragment");

export function jsx(type, props) {
  if (type === Fragment) return props?.children ?? [];
  if (typeof type !== "function") {
    throw new TypeError(`Unknown intrinsic JSX element: ${String(type)}`);
  }
  return type(props ?? {});
}

export const jsxs = jsx;
export const jsxDEV = jsx;

async function normalize(value, parentType) {
  value = await value;
  if (value === null || value === undefined || typeof value === "boolean") return [];
  if (Array.isArray(value)) {
    const output = [];
    for (const child of value) output.push(...await normalize(child, parentType));
    return output;
  }
  if (typeof value === "string") {
    if (structural.has(parentType) && value.trim() === "") return [];
    return [value];
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new TypeError("JSX numbers must be finite");
    return [value];
  }
  if (!value || value.__docxJsx !== true || typeof value.type !== "string") {
    throw new TypeError("A component returned a value that is not a docx-jsx node");
  }
  const children = await normalize(value.children, value.type);
  return [{ type: value.type, props: value.props ?? {}, children }];
}

export async function finalize(value, data) {
  if (typeof value === "function") value = value(data);
  const roots = await normalize(value, undefined);
  if (roots.length !== 1) throw new TypeError("The default export must produce exactly one Document");
  return { version: 1, document: roots[0] };
}

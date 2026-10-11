import {
  Bold,
  ContentControl,
  Document,
  Heading,
  Hyperlink,
  Image,
  InlineBookmark,
  List,
  ListItem,
  Paragraph,
  Run,
  Section,
  Table,
  TableCell,
  TableRow,
} from "docx-jsx";

export default (
  <Document
    createdAt="2026-08-14T00:00:00Z"
    updatedAt="2026-08-14T00:00:00Z"
    defaultFonts={{
      ascii: "Times New Roman",
      hiAnsi: "Times New Roman",
      eastAsia: "Noto Sans CJK SC",
      cs: "Times New Roman",
      hint: "eastAsia",
    }}
    defaultSize={11}
    defaultLineSpacing={{ after: 6, line: 14, lineRule: "atLeast" }}
    customProperties={{ Generator: "docx-jsx" }}
    documentVariables={{ ValidationFixture: "true" }}
    webExtensions={[
      {
        id: "7f33b723-fb58-4524-8733-dbedc4b7c095",
        referenceId: "docx-jsx-validation",
        version: "1.0.0.0",
        store: "developer",
        storeType: "Registry",
        properties: { mode: "validation" },
      },
      {
        id: "11111111-2222-3333-4444-555555555555",
        referenceId: "docx-jsx-secondary",
        version: "1.0.0.0",
        store: "OMEX",
        storeType: "Marketplace",
      },
    ]}
    customXmlItems={[
      {
        id: "06AC5857-5C65-A94A-BCEC-37356A209BC3",
        xml: "<validation><generator>docx-jsx</generator></validation>",
      },
    ]}
    styles={[
      { id: "BodyBase", name: "Body Base", type: "paragraph", paragraph: { spacingAfter: 6, hangingChars: 20 } },
      { id: "Body", name: "Body", type: "paragraph", basedOn: "BodyBase", next: "Body", link: "BodyChar", run: { fonts: { ascii: "Times New Roman", eastAsia: "Noto Sans CJK SC", hint: "eastAsia" }, size: 11 }, paragraph: { keepNext: false, keepLines: false } },
      { id: "BodyChar", name: "Body Char", type: "character", link: "Body", run: { italic: true } },
      { id: "NoLineNumbers", name: "No Line Numbers", type: "paragraph", paragraph: { suppressLineNumbers: true } },
      { id: "ValidationTable", name: "Validation Table", type: "table", table: { align: "center", indent: 6 }, cell: { verticalAlign: "center", shading: "EAF2F8", margins: { top: 1, right: 1, bottom: 1, left: 1 } } },
    ]}
  >
    <Section type="continuous" lineNumbering={{ countBy: 5, restart: "newSection", distance: 18 }} pageSize={{ width: 595.3, height: 841.9 }} margins={{ top: 72, right: 72, bottom: 72, left: 72 }} columns={2}>
      <Heading level={1}>Open XML validation fixture</Heading>
      <Paragraph style="Body">
        This document exercises <Bold>typed JSX</Bold> and package parts.
      </Paragraph>
      <Paragraph suppressLineNumbers>
        This paragraph is excluded from section line numbering.
      </Paragraph>
      <Paragraph suppressLineNumbers={false}>
        This paragraph explicitly keeps section line numbering.
      </Paragraph>
      <Paragraph style="NoLineNumbers">
        This style inherits line-number suppression.
      </Paragraph>
      <Paragraph fonts={{ ascii: "Times New Roman", eastAsia: "Noto Sans CJK SC" }}>
        <Run style="BodyChar" fonts={{ asciiTheme: "minorHAnsi", eastAsiaTheme: "minorEastAsia", hint: "eastAsia" }}>Linked character style</Run>
      </Paragraph>
      <Paragraph indentLeftChars={-100} indentRightChars={200} firstLineChars={640}>
        Character-unit indentation
      </Paragraph>
      <Paragraph indentStartChars={120} indentEndChars={80} hangingChars={20}>
        Logical character-unit indentation
      </Paragraph>
      <Paragraph style="Body">
        <Hyperlink href="https://example.com" history>
          <ContentControl alias="LinkData">bound</ContentControl>
          <InlineBookmark name="insideLink">marked</InlineBookmark>
        </Hyperlink>
      </Paragraph>
      <ContentControl alias="BlockData">
        <Paragraph style="Body">block bound</Paragraph>
      </ContentControl>
      <List
        type="ordered"
        multiLevelType="hybridMultilevel"
        levels={[
          {
            format: "decimal",
            text: "%1.",
            suffix: "space",
            legal: true,
            indentLeft: 36,
            hanging: 18,
          },
        ]}
        overrides={[{
          level: 0,
          start: 4,
          definition: { format: "lowerLetter", text: "%1)" },
        }]}
      >
        <ListItem indentStartChars={120} indentEndChars={80} hangingChars={20} suppressLineNumbers>First numbered</ListItem>
      </List>
      <Table style="ValidationTable" widthPercent={100} align="center" layout="fixed" columnWidths={[220, 220]} indent={2} margins={{ top: 1, right: 1, bottom: 1, left: 1 }} position={{ leftFromText: 2, rightFromText: 2, verticalAnchor: "text", horizontalAnchor: "margin", xAlign: "center", y: 0 }} border={{ style: "single", size: 0.5, color: "808080" }}>
        <TableRow height={30} heightRule="exact" cantSplit gridBefore={1} gridAfter={1} widthBefore={12} widthAfter={12}>
          <TableCell widthPercent={50} verticalAlign="center" margins={{ top: 2, right: 2, bottom: 2, left: 2 }} shading="D9EAF7"><Paragraph>Component</Paragraph></TableCell>
          <TableCell widthPercent={50}><Paragraph>Status</Paragraph></TableCell>
        </TableRow>
        <TableRow>
          <TableCell><Paragraph>OpenXmlValidator</Paragraph></TableCell>
          <TableCell><Paragraph>Required</Paragraph></TableCell>
        </TableRow>
      </Table>
      <Paragraph>
        <Image src="pixel.ppm" width={24} height={24} floating layoutInCell simplePosition={{ x: 25400, y: -12700 }} />
      </Paragraph>
    </Section>
  </Document>
);

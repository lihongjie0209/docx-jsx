import {
  Bold,
  Document,
  Heading,
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
    customXmlItems={[
      {
        id: "06AC5857-5C65-A94A-BCEC-37356A209BC3",
        xml: "<validation><generator>docx-jsx</generator></validation>",
      },
    ]}
    styles={[
      { id: "BodyBase", name: "Body Base", type: "paragraph", paragraph: { spacingAfter: 6 } },
      { id: "Body", name: "Body", type: "paragraph", basedOn: "BodyBase", next: "Body", link: "BodyChar", run: { fonts: { ascii: "Times New Roman", eastAsia: "Noto Sans CJK SC", hint: "eastAsia" }, size: 11 } },
      { id: "BodyChar", name: "Body Char", type: "character", link: "Body", run: { italic: true } },
      { id: "ValidationTable", name: "Validation Table", type: "table", table: { align: "center" } },
    ]}
  >
    <Section pageSize={{ width: 595.3, height: 841.9 }} margins={{ top: 72, right: 72, bottom: 72, left: 72 }}>
      <Heading level={1}>Open XML validation fixture</Heading>
      <Paragraph style="Body">
        This document exercises <Bold>typed JSX</Bold> and package parts.
      </Paragraph>
      <Paragraph fonts={{ ascii: "Times New Roman", eastAsia: "Noto Sans CJK SC" }}>
        <Run style="BodyChar" fonts={{ asciiTheme: "minorHAnsi", eastAsiaTheme: "minorEastAsia", hint: "eastAsia" }}>Linked character style</Run>
      </Paragraph>
      <Table style="ValidationTable" widthPercent={100} border={{ style: "single", size: 0.5, color: "808080" }}>
        <TableRow>
          <TableCell shading="D9EAF7"><Paragraph>Component</Paragraph></TableCell>
          <TableCell><Paragraph>Status</Paragraph></TableCell>
        </TableRow>
        <TableRow>
          <TableCell><Paragraph>OpenXmlValidator</Paragraph></TableCell>
          <TableCell><Paragraph>Required</Paragraph></TableCell>
        </TableRow>
      </Table>
    </Section>
  </Document>
);

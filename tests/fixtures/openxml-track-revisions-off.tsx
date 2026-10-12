import { Document, Section, Paragraph } from "docx-jsx";

export default (
  <Document trackRevisions={false}>
    <Section>
      <Paragraph>Revision tracking is explicitly disabled.</Paragraph>
    </Section>
  </Document>
);

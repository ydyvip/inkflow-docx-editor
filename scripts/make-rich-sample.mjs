import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  CommentRangeStart, CommentRangeEnd, CommentReference,
  Table, TableRow, TableCell, ShadingType, WidthType,
} from "docx";
import { writeFileSync } from "fs";

const doc = new Document({
  comments: {
    children: [
      { id: 1, author: "评审员", date: new Date("2026-01-01"), children: [new Paragraph("这里需要再确认一下数据来源。")] },
    ],
  },
  sections: [
    {
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("样式与批注测试")] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "居中的彩色大字", color: "C0392B", size: 36, bold: true }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun("这段包含一个"),
            new CommentRangeStart(1),
            new TextRun({ text: "被批注的短语", color: "1F5C5C", underline: {} }),
            new CommentRangeEnd(1),
            new CommentReference(1),
            new TextRun("，后面是普通文字。"),
          ],
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ shading: { type: ShadingType.CLEAR, fill: "EFEFEF" }, children: [new Paragraph("表头 A")] }),
                new TableCell({ shading: { type: ShadingType.CLEAR, fill: "D6EAF8" }, children: [new Paragraph("表头 B")] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("数据 1")] }),
                new TableCell({ children: [new Paragraph("数据 2")] }),
              ],
            }),
          ],
        }),
      ],
    },
  ],
});

const buf = await Packer.toBuffer(doc);
writeFileSync("/tmp/rich-sample.docx", buf);
console.log("wrote /tmp/rich-sample.docx", buf.length, "bytes");

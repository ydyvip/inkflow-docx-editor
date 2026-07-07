import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import { writeFileSync } from 'fs';

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun('测试文档标题')],
        }),
        new Paragraph({
          children: [
            new TextRun('这是一段普通文字，包含'),
            new TextRun({ text: '加粗', bold: true }),
            new TextRun('和'),
            new TextRun({ text: '斜体', italics: true }),
            new TextRun('内容。'),
          ],
        }),
        new Paragraph({ text: '第一条列表项', bullet: { level: 0 } }),
        new Paragraph({ text: '第二条列表项', bullet: { level: 0 } }),
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun('第二部分')],
        }),
        new Paragraph({ children: [new TextRun('结尾段落。')] }),
      ],
    },
  ],
});

const buf = await Packer.toBuffer(doc);
writeFileSync('/tmp/sample.docx', buf);
console.log('wrote /tmp/sample.docx', buf.length, 'bytes');

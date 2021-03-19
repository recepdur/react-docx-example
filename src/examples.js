import {
  Document,
  Footer,
  Header,
  PageNumber,
  PageNumberFormat,
  PageOrientation,
  Paragraph,
  TextRun,
  Bookmark,
  HeadingLevel,
  InternalHyperlink,
  PageBreak,
  Packer,
  StyleLevel,
  TableOfContents,
  File,
} from "docx";

export function ex1() {
  const doc = new Document();

  return doc;
}

export function multipleSections() {

  const doc = new Document();
  doc.addSection({
    children: [new Paragraph("Hello World")],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [new Paragraph("First Default Header on another page")],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph("Footer on another page")],
      }),
    },
    properties: {
      pageNumberStart: 1,
      pageNumberFormatType: PageNumberFormat.DECIMAL,
    },
    children: [new Paragraph("hello")],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [new Paragraph("Second Default Header on another page")],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph("Footer on another page")],
      }),
    },
    size: {
      orientation: PageOrientation.LANDSCAPE,
    },
    properties: {
      pageNumberStart: 1,
      pageNumberFormatType: PageNumberFormat.DECIMAL,
    },
    children: [new Paragraph("hello in landscape")],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: ["Page number: ", PageNumber.CURRENT],
              }),
            ],
          }),
        ],
      }),
    },
    size: {
      orientation: PageOrientation.PORTRAIT,
    },
    children: [new Paragraph("Page number in the header must be 2, because it continues from the previous section.")],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: ["Page number: ", PageNumber.CURRENT],
              }),
            ],
          }),
        ],
      }),
    },
    properties: {
      pageNumberFormatType: PageNumberFormat.UPPER_ROMAN,
      orientation: PageOrientation.PORTRAIT,
    },
    children: [
      new Paragraph(
        "Page number in the header must be III, because it continues from the previous section, but is defined as upper roman.",
      ),
    ],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: ["Page number: ", PageNumber.CURRENT],
              }),
            ],
          }),
        ],
      }),
    },
    size: {
      orientation: PageOrientation.PORTRAIT,
    },
    properties: {
      pageNumberFormatType: PageNumberFormat.DECIMAL,
      pageNumberStart: 25,
    },
    children: [
      new Paragraph("Page number in the header must be 25, because it is defined to start at 25 and to be decimal in this section."),
    ],
  });

  return doc;
}

export function bookmarks() {
  const LOREM_IPSUM =
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam mi velit, convallis convallis scelerisque nec, faucibus nec leo. Phasellus at posuere mauris, tempus dignissim velit. Integer et tortor dolor. Duis auctor efficitur mattis. Vivamus ut metus accumsan tellus auctor sollicitudin venenatis et nibh. Cras quis massa ac metus fringilla venenatis. Proin rutrum mauris purus, ut suscipit magna consectetur id. Integer consectetur sollicitudin ante, vitae faucibus neque efficitur in. Praesent ultricies nibh lectus. Mauris pharetra id odio eget iaculis. Duis dictum, risus id pellentesque rutrum, lorem quam malesuada massa, quis ullamcorper turpis urna a diam. Cras vulputate metus vel massa porta ullamcorper. Etiam porta condimentum nulla nec tristique. Sed nulla urna, pharetra non tortor sed, sollicitudin molestie diam. Maecenas enim leo, feugiat eget vehicula id, sollicitudin vitae ante.";

  const doc = new Document({
    creator: "Clippy",
    title: "Sample Document",
    description: "A brief example of using docx with bookmarks and internal hyperlinks",
  });

  doc.addSection({
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new InternalHyperlink({
                child: new TextRun({
                  text: "Click here!",
                  style: "Hyperlink",
                }),
                anchor: "myAnchorId",
              }),
            ],
          }),
        ],
      }),
    },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [
          new Bookmark({
            id: "myAnchorId",
            children: [new TextRun("Lorem Ipsum")],
          }),
        ],
      }),
      new Paragraph("\n"),
      new Paragraph(LOREM_IPSUM),
      new Paragraph({
        children: [new PageBreak()],
      }),
      new Paragraph({
        children: [
          new InternalHyperlink({
            child: new TextRun({
              text: "Anchor Text",
              style: "Hyperlink",
            }),
            anchor: "myAnchorId",
          }),
        ],
      }),
    ],
  });

  return doc;
}

export function tableOfContent() {
  const doc = new File({
    styles: {
      paragraphStyles: [
        {
          id: "MySpectacularStyle",
          name: "My Spectacular Style",
          basedOn: "Heading1",
          next: "Heading1",
          quickFormat: true,
          run: {
            italics: true,
            color: "990000",
          },
        },
      ],
    },
  });


  // WordprocessingML docs for TableOfContents can be found here:
  // http://officeopenxml.com/WPtableOfContents.php

  // Let's define the properties for generate a TOC for heading 1-5 and MySpectacularStyle,
  // making the entries be hyperlinks for the paragraph

  doc.addSection({
    children: [
      new TableOfContents("İçindekiler", {
        hyperlink: false,
        headingStyleRange: "1-5",
        stylesWithLevels: [new StyleLevel("MySpectacularStyle", 1)],
      }),
      new Paragraph({
        text: "Header #1",
        heading: HeadingLevel.HEADING_1,
        pageBreakBefore: true,
      }),
      new Paragraph("I'm a little text very nicely written.'"),
      new Paragraph({
        text: "Header #2",
        heading: HeadingLevel.HEADING_1,
        pageBreakBefore: false,
      }),
      new Paragraph("I'm a other text very nicely written.'"),
      new Paragraph({
        text: "Header #2.1",
        heading: HeadingLevel.HEADING_2,
      }),
      new Paragraph("I'm a another text very nicely written.'"),
      new Paragraph({
        text: "My Spectacular Style #1",
        style: "MySpectacularStyle",
        pageBreakBefore: false,
      }),
    ],
  });

  return doc;
}

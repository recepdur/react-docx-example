import React from "react";
import * as fs from "fs";
import { saveAs } from "file-saver";
import { multipleSections, bookmarks, tableOfContent } from "./examples";
import { Packer } from "docx";

export default function App() {
  return (
    <div>
      <h1>Hello</h1>
      <button onClick={generate}>Generate docx!</button>
    </div>
  );
  function generate() {
    var doc = tableOfContent();
    console.log(doc);
    // Packer.toBuffer(doc).then((buffer) => {
    //   fs.writeFileSync("document.docx", buffer);
    // });
    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, "document.docx");
    });
  }
}

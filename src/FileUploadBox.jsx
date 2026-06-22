import { useState } from "react";
import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import * as XLSX from "xlsx";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

export default function FileUploadBox({ onTextExtracted }) {
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [isReading, setIsReading] = useState(false);
  const [error, setError] = useState("");

  async function readTextFile(file) {
    return await file.text();
  }

  async function readDocxFile(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return result.value || "";
  }

  async function readXlsxFile(file) {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheets = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(sheet);
      sheets.push(`--- Sheet: ${sheetName} ---\n${csv}`);
    }

    return sheets.join("\n\n");
  }

  async function readPdfFile(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const pages = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      const pageText = content.items.map((item) => item.str).join(" ");
      pages.push(`--- Page ${pageNum} ---\n${pageText}`);
    }

    return pages.join("\n\n");
  }

  async function extractTextFromFile(file) {
    const name = file.name.toLowerCase();

    if (
      file.type.startsWith("text/") ||
      name.endsWith(".txt") ||
      name.endsWith(".md") ||
      name.endsWith(".csv") ||
      name.endsWith(".json")
    ) {
      return await readTextFile(file);
    }

    if (name.endsWith(".docx")) {
      return await readDocxFile(file);
    }

    if (name.endsWith(".pdf")) {
      return await readPdfFile(file);
    }

    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      return await readXlsxFile(file);
    }

    throw new Error(`Unsupported file type: ${file.name}`);
  }

  async function handleFileUpload(event) {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;

    setIsReading(true);
    setError("");

    try {
      const extractedDocs = [];

      for (const file of files) {
        const text = await extractTextFromFile(file);
        extractedDocs.push({
          name: file.name,
          text,
        });
      }

      const combinedText = extractedDocs
        .map((doc) => `Source File: ${doc.name}\n\n${doc.text}`)
        .join("\n\n==============================\n\n");

      onTextExtracted(combinedText);

      setUploadedFiles((prev) => [
        ...prev,
        ...extractedDocs.map((doc) => doc.name),
      ]);
    } catch (err) {
      console.error(err);
      setError(err.message || "Could not read one of the uploaded files.");
    } finally {
      setIsReading(false);
      event.target.value = "";
    }
  }

  return (
    <div className="file-upload-box">
      <h3>Upload Documents</h3>
      <p>
        Select PDFs, Word docs, Excel spreadsheets, text files, CSVs, or notes from your laptop.
        The app will extract the text and add it to the client notes.
      </p>

      <label className="file-upload-dropzone">
        <input
          type="file"
          multiple
          accept=".pdf,.docx,.txt,.md,.csv,.json,.xlsx,.xls,text/plain,text/markdown,text/csv,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
          onChange={handleFileUpload}
        />

        <span>{isReading ? "Reading documents..." : "Choose Documents from Laptop"}</span>
        <small>PDF, DOCX, XLSX, XLS, TXT, MD, CSV, JSON</small>
      </label>

      {uploadedFiles.length > 0 && (
        <div className="uploaded-file-list">
          <h4>Uploaded</h4>
          <ul>
            {uploadedFiles.map((file, index) => (
              <li key={`${file}-${index}`}>{file}</li>
            ))}
          </ul>
        </div>
      )}

      {error && <p className="file-upload-error">{error}</p>}
    </div>
  );
}

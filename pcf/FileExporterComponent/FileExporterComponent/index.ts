import { strToU8, zipSync } from "fflate";

import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface ProcessedDocument {
  id: string;
  fileName: string;
  fileExtension: string;
  fullFileName: string;
  mimeType: string;
  contentText: string;
  base64: string;
  sizeBytes: number;
  isValid: boolean;
  errorMessage: string;
}

interface ParseResult {
  documents: ProcessedDocument[];
  errorMessage: string;
}

interface DownloadResult {
  mode: "Disabled" | "Zip" | "Separate";
  action: "download" | "none";
  attemptedCount: number;
  downloadedCount: number;
  skippedCount: number;
  fileNames: string[];
  errors: string[];
}

type ButtonType = "primary" | "secondary" | "outline" | "text";
type FontWeight = "400" | "600" | "700";
type FontStyle = "normal" | "italic";
type TextDecoration = "none" | "underline" | "line-through";
type TextAlign = "flex-start" | "center" | "flex-end";
type VerticalAlign = "flex-start" | "center" | "flex-end";
type BorderStyle = "none" | "solid";
type DownloadMode = "Disabled" | "Zip" | "Separate";
type DisplayMode = "edit" | "disabled" | "view";

type EventContext = ComponentFramework.Context<IInputs> & {
  events?: {
    OnSelect?: () => void;
  };
};

export class FileExporterComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private static readonly outputKeys: (keyof IOutputs)[] = [
    "documentsOutputJson",
    "downloadResultJson",
    "documentCount",
    "validDocumentCount",
    "isValid",
    "errorMessage",
    "lastAction",
    "changeToken",
    "selectToken",
  ];

  private context!: ComponentFramework.Context<IInputs>;
  private notifyOutputChanged!: () => void;
  private container!: HTMLDivElement;
  private button!: HTMLButtonElement;
  private handleButtonClick = (): void => {
    this.handleClick();
  };

  private processedDocuments: ProcessedDocument[] = [];
  private isBusy = false;
  private outputs: IOutputs = {
    documentsOutputJson: "[]",
    downloadResultJson: "",
    documentCount: 0,
    validDocumentCount: 0,
    isValid: false,
    errorMessage: "",
    lastAction: "",
    changeToken: "",
    selectToken: "",
  };

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.container = container;
    this.context.mode.trackContainerResize(true);

    this.container.style.display = "flex";
    this.container.style.width = "100%";
    this.container.style.height = "100%";

    this.button = document.createElement("button");
    this.button.type = "button";
    this.button.addEventListener("click", this.handleButtonClick);
    this.container.appendChild(this.button);

    this.computeOutputs();
    this.render();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;
    this.computeOutputs();
    this.render();
  }

  public getOutputs(): IOutputs {
    return this.outputs;
  }

  public destroy(): void {
    this.button?.removeEventListener("click", this.handleButtonClick);
    this.button?.remove();
  }

  private render(): void {
    const parameters = this.context.parameters;
    const displayMode = this.getDisplayMode(parameters.displayModeInput.raw);
    const fillContainer = parameters.fillContainer.raw ?? true;
    const palette = this.getPaletteColor(parameters.colorPalette.raw);
    const buttonType = this.getButtonType(parameters.buttonType.raw);
    const buttonColors = this.resolveButtonColors(
      buttonType,
      palette,
      parameters.fontColor.raw ?? "",
      parameters.borderColor.raw ?? ""
    );
    const isDisabled = displayMode !== "edit" || this.isBusy;

    this.container.style.alignItems = fillContainer ? "stretch" : "center";
    this.container.style.justifyContent = fillContainer ? "stretch" : "flex-start";

    this.button.textContent = parameters.buttonText.raw && parameters.buttonText.raw.length > 0 ? parameters.buttonText.raw : "Export files";
    this.button.disabled = isDisabled;
    this.button.style.width = fillContainer ? "100%" : "auto";
    this.button.style.height = fillContainer ? "100%" : "auto";
    this.button.style.minWidth = "0";
    this.button.style.display = "flex";
    this.button.style.alignItems = this.getVerticalAlignment(parameters.verticalAlign.raw);
    this.button.style.justifyContent = this.getTextAlignment(parameters.textAlign.raw);
    this.button.style.background = buttonColors.backgroundColor;
    this.button.style.color = buttonColors.fontColor;
    this.button.style.borderStyle = this.getBorderStyle(parameters.borderStyle.raw);
    this.button.style.borderWidth = `${parameters.borderWidth.raw ?? 1}px`;
    this.button.style.borderColor = buttonColors.borderColor;
    this.button.style.borderRadius = `${parameters.borderRadius.raw ?? 6}px`;
    this.button.style.fontFamily = parameters.fontFamily.raw ?? "Segoe UI, sans-serif";
    this.button.style.fontSize = `${parameters.fontSize.raw ?? 14}px`;
    this.button.style.fontWeight = this.getFontWeight(parameters.fontWeight.raw);
    this.button.style.fontStyle = this.getFontStyle(parameters.fontStyle.raw);
    this.button.style.textDecoration = this.getTextDecoration(parameters.textDecoration.raw);
    this.button.style.paddingTop = `${parameters.paddingTop.raw ?? 0}px`;
    this.button.style.paddingRight = `${parameters.paddingRight.raw ?? 0}px`;
    this.button.style.paddingBottom = `${parameters.paddingBottom.raw ?? 0}px`;
    this.button.style.paddingLeft = `${parameters.paddingLeft.raw ?? 0}px`;
    this.button.style.cursor = isDisabled ? "not-allowed" : "pointer";
    this.button.style.opacity = isDisabled ? "0.65" : "1";
    this.button.style.boxSizing = "border-box";
    this.button.style.overflow = "hidden";
    this.button.style.whiteSpace = "normal";
    this.button.style.textAlign = "center";

    this.button.title = this.outputs.errorMessage && this.outputs.errorMessage.length > 0
      ? this.outputs.errorMessage
      : `Download mode: ${this.getDownloadMode(parameters.downloadMode.raw)}`;
  }

  private computeOutputs(): void {
    const parseResult = this.parseDocumentsJson(this.context.parameters.documentsJson.raw ?? "");
    this.processedDocuments = parseResult.documents;

    const documentsOutputJson = JSON.stringify(parseResult.documents);
    const documentCount = parseResult.documents.length;
    const validDocumentCount = parseResult.documents.filter((document) => document.isValid).length;
    const isValid = parseResult.errorMessage.length === 0 && documentCount > 0 && validDocumentCount === documentCount;

    const nextOutputs: IOutputs = {
      ...this.outputs,
      documentsOutputJson,
      documentCount,
      validDocumentCount,
      isValid,
      errorMessage: parseResult.errorMessage,
    };

    if (
      this.outputs.documentsOutputJson !== nextOutputs.documentsOutputJson ||
      this.outputs.documentCount !== nextOutputs.documentCount ||
      this.outputs.validDocumentCount !== nextOutputs.validDocumentCount ||
      this.outputs.isValid !== nextOutputs.isValid ||
      this.outputs.errorMessage !== nextOutputs.errorMessage
    ) {
      nextOutputs.changeToken = this.createToken();
    }

    this.updateOutputs(nextOutputs);
  }

  private parseDocumentsJson(rawDocumentsJson: string): ParseResult {
    const trimmed = rawDocumentsJson.trim();
    if (trimmed.length === 0) {
      return {
        documents: [],
        errorMessage: "Documents JSON is required.",
      };
    }

    let parsedValue: unknown;
    try {
      parsedValue = JSON.parse(trimmed);
    } catch {
      return {
        documents: [],
        errorMessage: "Documents JSON must be a valid JSON array.",
      };
    }

    if (!Array.isArray(parsedValue)) {
      return {
        documents: [],
        errorMessage: "Documents JSON must be a JSON array.",
      };
    }

    const documents = parsedValue.map((document, index) => this.normalizeDocument(document, index));
    const invalidDocuments = documents.filter((document) => !document.isValid);

    return {
      documents,
      errorMessage:
        invalidDocuments.length > 0
          ? `${invalidDocuments.length} document(s) are invalid.`
          : documents.length === 0
            ? "At least one document is required."
            : "",
    };
  }

  private normalizeDocument(document: unknown, index: number): ProcessedDocument {
    if (!document || typeof document !== "object" || Array.isArray(document)) {
      return {
        id: `doc-${index + 1}`,
        fileName: "",
        fileExtension: "",
        fullFileName: "",
        mimeType: "",
        contentText: "",
        base64: "",
        sizeBytes: 0,
        isValid: false,
        errorMessage: "Document entry must be an object.",
      };
    }

    const candidate = document as Record<string, unknown>;
    const fileName = this.sanitizeFileName(this.asTrimmedString(candidate.fileName));
    const fileExtension = this.sanitizeExtension(this.asTrimmedString(candidate.fileExtension));
    const mimeType = this.asTrimmedString(candidate.mimeType);
    const contentText = this.asString(candidate.contentText ?? candidate.content);
    const errors: string[] = [];

    if (!fileName) {
      errors.push("fileName is required.");
    }

    if (!fileExtension) {
      errors.push("fileExtension is required.");
    }

    if (!mimeType) {
      errors.push("mimeType is required.");
    }

    if (contentText.length === 0) {
      errors.push("contentText is required.");
    }

    const fullFileName = fileName && fileExtension ? `${fileName}.${fileExtension}` : "";
    const sizeBytes = new Blob([contentText]).size;

    return {
      id: this.asTrimmedString(candidate.id) || `doc-${index + 1}`,
      fileName,
      fileExtension,
      fullFileName,
      mimeType,
      contentText,
      base64: errors.length === 0 ? this.toBase64(contentText) : "",
      sizeBytes,
      isValid: errors.length === 0,
      errorMessage: errors.join(" "),
    };
  }

  private handleClick(): void {
    if (this.getDisplayMode(this.context.parameters.displayModeInput.raw) !== "edit" || this.isBusy) {
      return;
    }

    this.triggerOnSelect();

    const downloadMode = this.getDownloadMode(this.context.parameters.downloadMode.raw);
    const action = downloadMode === "Disabled" ? "none" : "download";
    const selectToken = this.createToken();

    if (action === "none") {
      const noActionResult: DownloadResult = {
        mode: downloadMode,
        action,
        attemptedCount: 0,
        downloadedCount: 0,
        skippedCount: 0,
        fileNames: [],
        errors: [],
      };

      this.updateOutputs({
        ...this.outputs,
        lastAction: action,
        selectToken,
        downloadResultJson: JSON.stringify(noActionResult),
      });
      return;
    }

    const validDocuments = this.processedDocuments.filter((document) => document.isValid);
    const skippedCount = this.processedDocuments.length - validDocuments.length;

    this.isBusy = true;
    this.render();

    let downloadResult: DownloadResult;

    try {
      if (validDocuments.length === 0) {
        downloadResult = {
          mode: downloadMode,
          action,
          attemptedCount: 0,
          downloadedCount: 0,
          skippedCount,
          fileNames: [],
          errors: ["No valid documents were available to download."],
        };
      } else if (downloadMode === "Zip") {
        downloadResult = this.downloadAsZip(validDocuments, skippedCount);
      } else {
        downloadResult = this.downloadAsSeparateFiles(validDocuments, skippedCount);
      }
    } catch (error) {
      downloadResult = {
        mode: downloadMode,
        action,
        attemptedCount: validDocuments.length,
        downloadedCount: 0,
        skippedCount,
        fileNames: validDocuments.map((document) => document.fullFileName),
        errors: [error instanceof Error ? error.message : "Download failed."],
      };
    } finally {
      this.isBusy = false;
      this.render();
    }

    this.updateOutputs({
      ...this.outputs,
      lastAction: action,
      selectToken,
      downloadResultJson: JSON.stringify(downloadResult),
    });
  }

  private downloadAsZip(validDocuments: ProcessedDocument[], skippedCount: number): DownloadResult {
    const zipEntries: Record<string, Uint8Array> = {};

    validDocuments.forEach((document) => {
      zipEntries[document.fullFileName] = strToU8(document.contentText);
    });

    const archiveBytes = zipSync(zipEntries);
    const archiveBuffer = new ArrayBuffer(archiveBytes.byteLength);
    new Uint8Array(archiveBuffer).set(archiveBytes);
    const archiveBlob = new Blob([archiveBuffer], { type: "application/zip" });
    const archiveFileName = this.normalizeArchiveFileName(this.context.parameters.archiveFileName.raw ?? "");

    this.downloadBlob(archiveBlob, archiveFileName, "application/zip");

    return {
      mode: "Zip",
      action: "download",
      attemptedCount: validDocuments.length,
      downloadedCount: validDocuments.length,
      skippedCount,
      fileNames: validDocuments.map((document) => document.fullFileName),
      errors: [],
    };
  }

  private downloadAsSeparateFiles(validDocuments: ProcessedDocument[], skippedCount: number): DownloadResult {
    validDocuments.forEach((document) => {
      this.downloadBlob(new Blob([document.contentText], { type: document.mimeType }), document.fullFileName, document.mimeType);
    });

    return {
      mode: "Separate",
      action: "download",
      attemptedCount: validDocuments.length,
      downloadedCount: validDocuments.length,
      skippedCount,
      fileNames: validDocuments.map((document) => document.fullFileName),
      errors: [],
    };
  }

  private downloadBlob(blob: Blob, fileName: string, mimeType: string): void {
    const anchor = document.createElement("a");
    const objectUrl = URL.createObjectURL(blob.slice(0, blob.size, mimeType));

    anchor.href = objectUrl;
    anchor.download = fileName;
    anchor.style.display = "none";
    this.container.appendChild(anchor);
    anchor.click();
    anchor.remove();

    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  }

  private triggerOnSelect(): void {
    const eventContext = this.context as EventContext;
    if (eventContext.events?.OnSelect) {
      eventContext.events.OnSelect();
    }
  }

  private getDisplayMode(rawValue: number | boolean | string | null | undefined): DisplayMode {
    if (rawValue === 1 || rawValue === "1" || rawValue === "disabled") {
      return "disabled";
    }

    if (rawValue === 2 || rawValue === "2" || rawValue === "view") {
      return "view";
    }

    return "edit";
  }

  private getDownloadMode(rawValue: number | string | null | undefined): DownloadMode {
    if (rawValue === 1 || rawValue === "1") {
      return "Zip";
    }

    if (rawValue === 2 || rawValue === "2") {
      return "Separate";
    }

    return "Disabled";
  }

  private getButtonType(rawValue: number | string | null | undefined): ButtonType {
    if (rawValue === 1 || rawValue === "1") {
      return "secondary";
    }

    if (rawValue === 2 || rawValue === "2") {
      return "outline";
    }

    if (rawValue === 3 || rawValue === "3") {
      return "text";
    }

    return "primary";
  }

  private getFontWeight(rawValue: number | string | null | undefined): FontWeight {
    if (rawValue === 0 || rawValue === "0") {
      return "400";
    }

    if (rawValue === 2 || rawValue === "2") {
      return "700";
    }

    return "600";
  }

  private getFontStyle(rawValue: number | string | null | undefined): FontStyle {
    return rawValue === 1 || rawValue === "1" ? "italic" : "normal";
  }

  private getTextDecoration(rawValue: number | string | null | undefined): TextDecoration {
    if (rawValue === 1 || rawValue === "1") {
      return "underline";
    }

    if (rawValue === 2 || rawValue === "2") {
      return "line-through";
    }

    return "none";
  }

  private getTextAlignment(rawValue: number | string | null | undefined): TextAlign {
    if (rawValue === 0 || rawValue === "0") {
      return "flex-start";
    }

    if (rawValue === 2 || rawValue === "2") {
      return "flex-end";
    }

    return "center";
  }

  private getVerticalAlignment(rawValue: number | string | null | undefined): VerticalAlign {
    if (rawValue === 0 || rawValue === "0") {
      return "flex-start";
    }

    if (rawValue === 2 || rawValue === "2") {
      return "flex-end";
    }

    return "center";
  }

  private getBorderStyle(rawValue: number | string | null | undefined): BorderStyle {
    return rawValue === 0 || rawValue === "0" ? "none" : "solid";
  }

  private getPaletteColor(rawValue: string | null | undefined): string {
    const trimmed = (rawValue ?? "").trim();
    return trimmed.length > 0 ? trimmed : "#0078d4";
  }

  private resolveButtonColors(buttonType: ButtonType, palette: string, rawFontColor: string, rawBorderColor: string): {
    backgroundColor: string;
    borderColor: string;
    fontColor: string;
  } {
    const fontColor = rawFontColor.trim();
    const borderColor = rawBorderColor.trim();

    if (buttonType === "secondary") {
      return {
        backgroundColor: "#f3f2f1",
        borderColor: borderColor || "#d1d1d1",
        fontColor: fontColor || "#323130",
      };
    }

    if (buttonType === "outline") {
      return {
        backgroundColor: "transparent",
        borderColor: borderColor || palette,
        fontColor: fontColor || palette,
      };
    }

    if (buttonType === "text") {
      return {
        backgroundColor: "transparent",
        borderColor: borderColor || "transparent",
        fontColor: fontColor || palette,
      };
    }

    return {
      backgroundColor: palette,
      borderColor: borderColor || palette,
      fontColor: fontColor || "#ffffff",
    };
  }

  private sanitizeFileName(value: string): string {
    const cleaned = Array.from(value)
      .filter((character) => {
        const code = character.charCodeAt(0);
        return code >= 32 && !/[<>:"/\\|?*]/.test(character);
      })
      .join("");

    return cleaned.replace(/\s+/g, " ").replace(/[. ]+$/g, "").trim();
  }

  private sanitizeExtension(value: string): string {
    return value.replace(/^\.+/, "").replace(/[^a-zA-Z0-9_-]/g, "").trim();
  }

  private normalizeArchiveFileName(value: string): string {
    const trimmed = this.sanitizeFileName(value.trim() || "documents");
    return trimmed.toLowerCase().endsWith(".zip") ? trimmed : `${trimmed}.zip`;
  }

  private toBase64(text: string): string {
    const bytes = new TextEncoder().encode(text);
    let binary = "";

    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });

    return btoa(binary);
  }

  private createToken(): string {
    return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  }

  private asString(value: unknown): string {
    if (typeof value === "string") {
      return value;
    }

    if (value === null || value === undefined) {
      return "";
    }

    if (typeof value === "number" || typeof value === "boolean") {
      return String(value);
    }

    return "";
  }

  private asTrimmedString(value: unknown): string {
    return this.asString(value).trim();
  }

  private updateOutputs(nextOutputs: IOutputs): void {
    const hasChanged = FileExporterComponent.outputKeys.some((key) => this.outputs[key] !== nextOutputs[key]);
    this.outputs = nextOutputs;

    if (hasChanged) {
      this.notifyOutputChanged();
    }
  }
}

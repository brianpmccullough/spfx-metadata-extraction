import { ListViewCommandSetContext, RowAccessor } from "@microsoft/sp-listview-extensibility";
import { IDocumentContext } from "../../models/IDocumentContext";
import { DocumentContextFactory } from "./DocumentContextFactory";

export class MetadataExtractionContext {
    private readonly _factory: DocumentContextFactory = new DocumentContextFactory();

    constructor(
        private readonly context: ListViewCommandSetContext,
        private readonly allowedFileTypes: string[]
    ) {}

    public get selectedRowCount(): number {
        return this.context.listView?.selectedRows?.length ?? 0;
    }

    private get selectedRow(): RowAccessor | undefined {
        return this.context.listView?.selectedRows?.[0];
    }

    public get isValidFileType(): boolean {
        const document = this.documentContext;
        if (!document) {
            return false;
        }
        const extension = `.${document.fileExtension?.toLowerCase()}`;
        return this.allowedFileTypes.some(
            (allowed) => allowed.toLowerCase() === extension
        );
    }

    public get isVisible(): boolean {
        return this.selectedRowCount === 1 && this.isValidFileType;
    }

    public get canExecute(): boolean {
        return this.selectedRow !== undefined && this.isValidFileType;
    }

    public get documentContext(): IDocumentContext | undefined {
        const selectedRow = this.selectedRow;
        if (selectedRow) {
            return this._factory.createFromContext(this.context, selectedRow);
        }
    }
}

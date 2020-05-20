export interface IFileUploadState {
    digest: string;
    attachments: any[];
  }
  
  export interface IDocument {
    name: string;
    value: string;
    iconName: string;
    fileType: string;
    modifiedBy: string;
    dateModified: string;
    dateModifiedValue: number;
    fileSize: string;
    fileSizeRaw: number;
  }
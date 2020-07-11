  export interface FileInterface {
    signature: number;
    signatureLen: number;
    fileExtension: string;
    validate: (buffer: Buffer) => boolean
  }
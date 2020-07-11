import { FileInterface } from './interface'

export const DOCX: FileInterface = {
  signature:0xd050,
  signatureLen:2,
  fileExtension:"docx",
  validate(buffer) { return (buffer.subarray(0,(this.signatureLen))[0] === this.signature) }
}

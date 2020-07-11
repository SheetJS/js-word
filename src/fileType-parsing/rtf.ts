import { FileInterface } from './interface'

export const RTF: FileInterface = {
  signature:0x7b,
  signatureLen:1,
  fileExtension:"rtf",
  validate(buffer) { return (buffer.subarray(0,(this.signatureLen))[0] === this.signature) }
}
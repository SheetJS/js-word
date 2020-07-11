import { FileInterface } from './interface'

export const XML: FileInterface = {
  signature:0x3c,
  signatureLen:1,
  fileExtension:"xml",
  validate(buffer) { return (buffer.subarray(0,(this.signatureLen))[0] === this.signature) }
}
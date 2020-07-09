import { FileInterface } from './interface'

export const DOC: FileInterface = {
  signature:0xd0,
  signatureLen:1,
  fileExtension:"doc",
  validate(buffer) { return (buffer.subarray(0,(this.signatureLen))[0] === this.signature) }
}

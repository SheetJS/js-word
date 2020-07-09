const path = require('path');
const CFB = require('cfb');
const {JSDOM} = require('jsdom');

/**
 * Grabs the text content of an odt file
 * 
 * @param {string} file path to .odt file
 * @return {string} text content of file
 */
function odtToText(file) {
  // Read the content.xml of the file
  const cfb = CFB.read(path.join(__dirname, file), {type: 'file'});
  const buf = CFB.find(cfb, 'content.xml').content;

  // Parse with JSDOM
  const dom = new JSDOM(buf);
  let result = "";

  // Use querySelector to grab elements from content.xml
  // We can grab any element, not just text:p
  dom.window.document.querySelectorAll("text\\:p").forEach((element) => {
    // Do something with each paragraph here
    //   TODO: compare output to word text files
    result += element.textContent + "\n";
  });

  return result;
}

module.exports = odtToText;
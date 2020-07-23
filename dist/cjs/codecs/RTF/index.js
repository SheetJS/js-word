"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
function parse_str(data) {
    // var pairs = [];
    // var seen_par = false;
    var blacklist = ["\\fonttbl", "\\rtlch", "\\fcs1", "\\af0", "\\ltrch", "\\fcs0", "\\insrsid"];
    var counter = [];
    var offsets = [];
    var para_text = [];
    var current_par = "";
    var par_start = -1;
    var last_brace = Infinity;
    var skiplength = Infinity; // keep track of skip length
    /*
        TEST SAMPLES
    */
    // var rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0\\froman Tms Rmn;}}"
    // var rtf = "{\\par The word \'93}{\\cs15\\b\\ul\\cf6 style}{\'94 is red and underlined. I used a style I called UNDERLINE.\\par }"
    // var rtf = "{\\rtlch\\fcs1 \\af31507 \\ltrch\\fcs0 \\insrsid6445377 Hello}{\\rtlch\\fcs1 \\af31507 \\ltrch\\fcs0 \\insrsid5199918 \\par }{\\rtlch\\fcs1 \\af31507 \\ltrch\\fcs0 \\insrsid6445377 There\\par }"
    // var rtf = "{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid6445377 Hello}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid3943939 \\par }{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid5116832 aaa}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid6445377 \\par }"
    var rtf = "{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid6445377 Hello}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid11279206  ba}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid3943939 \\par }{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid5116832 aaa}";
    // var rtf = "{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid6445377 Hello}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid11279206  ba}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid14709222  bok}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid3943939 \\par }{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\insrsid5116832 aaa}{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \insrsid6445377 \\par }"
    rtf = data;
    var extract_text = rtf.replace(/({|\\[A-Za-z\d]+|}?\b\w+\b)/g, function ($$, $1, idx) {
        if ($$ == "{") {
            counter.push(0);
            offsets.push(idx);
            last_brace = idx;
            return "{";
        }
        else if ($$ == "}") {
            counter.pop();
            offsets.pop();
            last_brace = -1;
            // if (seen_par) {
            //   console.log(pairs)
            // }
            if (counter.length < skiplength) {
                skiplength = Infinity; // reset if we exited
                if (par_start > -1) {
                    var par_snippet = rtf.slice(par_start + "\\par".length, idx);
                    par_start = -1;
                    current_par += par_snippet;
                }
            }
            return "}";
        }
        else if (blacklist.indexOf($$) == -1) {
            var text = $$.replace(/\\\w+ ?/g, ""); // .match(/\b\w+\b/g)
            if (text !== null) {
                para_text.push(text[0]);
            }
        }
        else if (counter.length > skiplength) {
            return ""; // skip
        }
        else {
            if (counter[counter.length - 1]++ == 0) { // counter (stack) is empty
                if (blacklist.indexOf($$) > -1) { // if word found in blacklist
                    skiplength = counter.length;
                    return ""; // skip   - found in blacklist arr
                }
                else if ($$ == "\\par") { // Control word found
                    // +1 for the length of `{`
                    current_par += rtf.slice(offsets[offsets.length - 1] + 1, idx);
                    // if (par_start > -1) {
                    //    console.log(rtf.slice(offsets.peak(), idx)) // The text between \par tags
                    // }
                    // if (!seen_par) {
                    //    seen_par = true;
                    // }
                    par_start = idx; // mark where \\par begins
                    skiplength = counter.length; // skip
                    //pairs.push($$)
                    para_text.push(current_par);
                    current_par = "";
                }
            }
            else if ($$ == "\\par") {
                current_par += rtf.slice(offsets[offsets.length - 1] + 1, idx);
                para_text.push(current_par);
                current_par = "";
            }
        }
        return $$;
    });
    return { p: [{ elts: [{ t: "s", v: para_text.join("") }] }] };
}
exports.parse_str = parse_str;
function read(data) {
    return parse_str(data.toString());
}
exports.read = read;
//# sourceMappingURL=index.js.map
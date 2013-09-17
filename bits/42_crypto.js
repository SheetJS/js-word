/* [MS-OFFCRYPTO] 2.1.4 Version */
function parse_Version(blob, length) {
	var o = {};
	o.Major = blob.read_shift(2);
	o.Minor = blob.read_shift(2);
	return o;
}
/* [MS-OFFCRYPTO] 2.3.2 Encryption Header */
function parse_EncryptionHeader(blob, length) {
	var read = blob.read_shift.bind(blob);
	var o = {};
	o.Flags = read(4);

	// Check if SizeExtra is 0x00000000
	var tmp = read(4);
	if(tmp !== 0) throw 'Unrecognized SizeExtra: ' + tmp;

	o.AlgID = read(4);
	switch(o.AlgID) {
		case 0: case 0x6801: case 0x660E: case 0x660F: case 0x6610: break;
		default: throw 'Unrecognized encryption algorithm: ' + o.AlgID;
	}
	parsenoop(blob, length-12);
	return o;
}

/* [MS-OFFCRYPTO] 2.3.3 Encryption Verifier */
function parse_EncryptionVerifier(blob, length) {
	return parsenoop(blob, length);
}
/* [MS-OFFCRYPTO] 2.3.5.1 RC4 CryptoAPI Encryption Header */
function parse_RC4CryptoHeader(blob, length) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_Version(blob, 4); length -= 4;
	if(vers.Minor != 2) throw 'unrecognized minor version code: ' + vers.Minor;
	if(vers.Major > 4 || vers.Major < 2) throw 'unrecognized major version code: ' + vers.Major;
	o.Flags = blob.read_shift(4); length -= 4;
	var sz = blob.read_shift(4); length -= 4;
	o.EncryptionHeader = parse_EncryptionHeader(blob, sz); length -= sz;
	o.EncryptionVerifier = parse_EncryptionVerifier(blob, length);
	return o;
}
/* [MS-OFFCRYPTO] 2.3.6.1 RC4 Encryption Header */
function parse_RC4Header(blob, length) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_Version(blob, 4); length -= 4;
	if(vers.Major != 1 || vers.Minor != 1) throw 'unrecognized version code ' + vers.Major + ' : ' + vers.Minor;
	o.Salt = blob.read_shift(16);
	o.EncryptedVerifier = blob.read_shift(16);
	o.EncryptedVerifierHash = blob.read_shift(16);
	return o;
}

/* 2.5.343 */
function parse_XORObfuscation(blob, length) { return parsenoop(blob, length); }
/* 2.4.117 */
function parse_FilePassHeader(blob, length, oo) {
	var o = oo || {}; o.Info = blob.read_shift(2); blob.l -= 2;
	switch(o.Info) {
		case 1: o.Data = parse_RC4Header(blob, length); break;
		case 2: case 3: case 4: o.Data = parse_RC4CryptoHeader(blob, length); break;
		default: throw 'Unrecognized encryptionInfo: ' + o.Type;
	}
	return o;
}
function parse_FilePass(blob, length) {
	var o = { Type: blob.read_shift(2) }; /* wEncryptionType */
	switch(o.Type) {
		case 0: parse_XORObfuscation(blob, length-2, o); break;
		case 1: parse_FilePassHeader(blob, length-2, o); break;
		default: throw 'Unrecognized Encryption Type ' + o.Type;
	}
	return o;
}


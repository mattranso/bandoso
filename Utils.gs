function hashPassword_(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var hash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var byte = rawHash[i];
    if (byte < 0) byte += 256;
    var byteString = byte.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    hash += byteString;
  }
  return hash;
}
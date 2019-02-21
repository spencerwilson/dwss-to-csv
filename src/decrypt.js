const Encryptor = new (require('./encryption'))();

/**
 * @param {Object} cfb
 *        A parsed CFB containing an encrypted ECMA-386 document (e.g., docx, xlsx, pptx).
 * @param {String} password
 *        Password for the given document.
 * @returns {Promise.<Buffer>}
 *        The decrypted document.
 */
module.exports = function decrypt(cfb, password) {
  return Encryptor.decryptAsync(cfb, password);
};


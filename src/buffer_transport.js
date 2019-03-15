/**
 * @param {Transport} transport
 * @returns {Transport}
 *        An object whose prototype is the given transport, but has its `log` method
 *        shadowed to queue invocations in a buffer. The queued invocations can be
 *        flushed to the underlying `log` method by calling `flush`.
 */
module.exports = function buffer(transport) {
  let buffer = [];

  return Object.create(transport, {
    log: {
      value: (...args) => buffer.push(args),
    },
    flush: {
      value: () => {
        buffer.forEach((logArgs) => transport.log(...logArgs));
        buffer = [];
      }
    },
  });
};

// Just like Python 3's `range`.
exports.range = function* range(start, end) {
  for (let i = start; i < end; i++) {
    yield i;
  }
}

exports.symbolMirror = function symbolMirror(names) {
  return names.reduce((acc, n) => {
    acc[n] = Symbol(n);
    return acc;
  }, {});
}


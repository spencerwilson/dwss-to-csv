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

function description(s) {
  return s.toString().match(/Symbol\((.*)\)/)[1];
}
exports.description = description;

exports.serialize = function serialize(mapOrSet) {
  if (mapOrSet instanceof Map) {
    const mappings = Array.from(mapOrSet).map(([k, v]) => {
      if (k.constructor === Symbol) k = description(k);
      if (v.constructor === Symbol) v = description(v);

      return `${k} => ${v.sheetName}`;
    });
    return JSON.stringify(mappings, null, 2);
  } else if (mapOrSet instanceof Set) {
    return Array.from(mapOrSet).map(v => {
      if (v.constructor === Symbol) v = description(v);
      return v;
    }).join(', ');
  }
}

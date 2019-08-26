// modified from https://github.com/puleos/object-hash
let crypto, encoder;
const isNode = typeof window === 'undefined';
if (isNode)
    crypto = require('crypto');
else {
    crypto = window.crypto;
    encoder = new TextEncoder();
}
const algorithm = isNode ? 'sha1' : 'SHA-1';

// Since buffer in browser is super slow, we will use string instead.
function BrowserStream() {
    return {
        buf: '',

        update: (b) => {
            this.buf += b;
        },

        digest: () => {
            return crypto.subtle.digest(algorithm, encoder.encode(this.buf));
        }
    };
}

const hash = async (object) => {
    let hashingStream;

    hashingStream = isNode ? crypto.createHash(algorithm) : BrowserStream();

    const hasher = typeHasher(hashingStream);
    hasher.dispatch(object);

    return hashingStream.digest();
};

function typeHasher(hashingStream, context) {
    context = context || [];
    const write = function (str) {
        return hashingStream.update(str, 'utf8');
    };

    return {
        dispatch: function (value) {
            let type = typeof value;
            if (value === null) {
                type = 'null';
            }

            //console.log("[DEBUG] Dispatch: ", value, "->", type, " -> ", "_" + type);

            return this['_' + type](value);
        },
        _object: function (object) {
            const pattern = (/\[object (.*)\]/i);
            const objString = Object.prototype.toString.call(object);
            let objType = pattern.exec(objString);
            if (!objType) { // object type did not match [object ...]
                objType = 'unknown:[' + objString + ']';
            } else {
                objType = objType[1]; // take only the class name
            }

            objType = objType.toLowerCase();

            let objectNumber;

            if ((objectNumber = context.indexOf(object)) >= 0) {
                return this.dispatch('[CIRCULAR:' + objectNumber + ']');
            } else {
                context.push(object);
            }
            if (objType !== 'object' && objType !== 'function') {
                if (this['_' + objType]) {
                    this['_' + objType](object);
                } else {
                    throw new Error('Unknown object type "' + objType + '"');
                }
            } else {
                let keys = Object.keys(object);
                keys = keys.sort();
                // Make sure to incorporate special properties, so
                // Types with different prototypes will produce
                // a different hash and objects derived from
                // different functions (`new Foo`, `new Bar`) will
                // produce different hashes.
                // We never do this for native functions since some
                // seem to break because of that.

                write('object:' + keys.length + ':');
                const self = this;
                return keys.forEach(function (key) {
                    self.dispatch(key);
                    write(':');
                    self.dispatch(object[key]);
                    write(',');
                });
            }
        },
        _array: function (arr) {
            const self = this;
            write('array:' + arr.length + ':');
            return arr.forEach(function (entry) {
                return self.dispatch(entry);
            });
        },
        _boolean: function (bool) {
            return write('bool:' + bool.toString());
        },
        _string: function (string) {
            write('string:' + string.length + ':');
            write(string.toString());
        },
        _number: function (number) {
            return write('number:' + number.toString());
        },
        _null: function () {
            return write('Null');
        },
        _undefined: function () {
            return write('Undefined');
        }
    };
}

module.exports = {
    hash
};

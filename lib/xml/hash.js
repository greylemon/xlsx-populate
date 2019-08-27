// modified from https://github.com/puleos/object-hash

/**
 * https://gist.github.com/hyamamoto/fd435505d29ebfa3d9716fd2be8d42f0
 * Returns a hash code for a string.
 * (Compatible to Java's String.hashCode())
 *
 * The hash code for a string object is computed as
 *     s[0]*31^(n-1) + s[1]*31^(n-2) + ... + s[n-1]
 * using number arithmetic, where s[i] is the i th character
 * of the given string, n is the length of the string,
 * and ^ indicates exponentiation.
 * (The hash value of the empty string is zero.)
 *
 * @param {string} s a string
 * @return {number} a hash code value for the given string.
 */
const hashString = function (s) {
    const l = s.length;
    let h = 0, i = 0;
    if (l > 0)
        while (i < l)
            h = (h << 5) - h + s.charCodeAt(i++) | 0;
    return h;
};

const hash = async (object) => {
    const hasher = xmlNodeHasher();
    return hasher.hash(object);
};

function xmlNodeHasher(context) {
    let buf;
    const write = function (str) {
        buf += str;
    };

    return {
        dispatch: function (value) {
            let type = typeof value;
            if (value === null) {
                type = 'null';
            }
            this['_' + type](value);
        },
        hash: function (value) {
            buf = '';
            context = [];
            this.dispatch(value);
            return hashString(buf);
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

const cloneDeep = o => {
    let newO;
    let i;

    if (typeof o !== 'object') return o;

    if (!o) return o;

    if (Array.isArray(o)) {
        newO = [];
        for (i = 0; i < o.length; i += 1) {
            newO[i] = cloneDeep(o[i]);
        }
        return newO;
    }

    newO = {};
    for (i in o) {
        if (o.hasOwnProperty(i)) {
            newO[i] = cloneDeep(o[i]);
        }
    }
    return newO;
};

const mapValues = (obj, field) => {
    Object.entries(obj).reduce((a, [key, value]) => {
        a[key] = value[field];
        return a;
    }, {});
};

const escapeRegExp = string => {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
};

module.exports = {
    cloneDeep, mapValues, escapeRegExp, ...require('./config'),
};

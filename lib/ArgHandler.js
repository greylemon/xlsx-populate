"use strict";

/**
 * Method argument handler. Used for overloading methods.
 * @private
 */
class ArgHandler {
    /**
     * Creates a new instance of ArgHandler.
     * @param {string} name - The method name to use in error messages
     * @param {IArguments|Array} args - The arguments
     */
    constructor(name, args) {
        this._name = name;
        this._args = args;
        this._found = false;
        this._result = undefined;
    }

    /**
     * Add a case.
     * @param {string|Array.<string>} [types] - The type or types of arguments to match this case.
     * @param {Function} handler - The function to call when this case is matched.
     * @returns {ArgHandler} The handler for chaining.
     */
    case(types, handler) {
        if (this._found) return this;
        if (arguments.length === 1) {
            if (this._args.length === 0) {
                this._result = types();
                this._found = true;
            }
        } else {
            if (!Array.isArray(types))
                types = [types];
            if (this._argsMatchTypes(types)) {
                this._result = handler.apply(null, this._args);
                this._found = true;
            }
        }
        return this;
    }

    /**
     * Handle the method arguments by checking each case in order until one matches and then call its handler.
     * @returns {undefined} The result of the handler.
     * @throws {Error} Throws if no case matches.
     */
    handle() {
        if (this._found)
            return this._result;
        throw new Error(`${this._name}: Invalid arguments.`);
    }

    /**
     * Check if the arguments match the given types.
     * @param {Array.<string>} types - The types.
     * @returns {boolean} True if matches, false otherwise.
     * @throws {Error} Throws if unknown type.
     * @private
     */
    _argsMatchTypes(types) {
        if (this._args.length !== types.length)
            return false;

        return types.every((type, i) => {
            const arg = this._args[i];
            const actualType = typeof arg;

            if (type === '*') return true;
            else if (type === 'nil') return arg == null;
            else if (type === 'string') return actualType === "string";
            else if (type === 'boolean') return actualType === "boolean";
            else if (type === 'number') return actualType === "number";
            else if (type === 'integer') return Math.trunc(arg) === arg;
            else if (type === 'function') return actualType === "function";
            else if (type === 'array') return Array.isArray(arg);
            else if (type === 'date') return arg instanceof Date;
            else if (type === 'object') return arg instanceof Object;
            else if (arg && arg.constructor && arg.constructor.name === type) return true; // ?
            else throw new Error(`Unknown type: ${type}`);
        });
    }
}

module.exports = ArgHandler;

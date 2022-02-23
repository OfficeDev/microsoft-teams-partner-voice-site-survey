export const groupBy = (list, keyGetter) => {
    const map = new Map();
    list.forEach((item) => {
        const key = keyGetter(item);
        const collection = map.get(key);
        if (!collection) {
            map.set(key, [item]);
        } else {
            collection.push(item);
        }
    });
    return map;
};

export const removeSpaceToLowercase = (str: string) => {
    return str.replace(/\s+/g, '').replace('-', '').toLowerCase();
};

export const removeSpace = (str: string) => {
    return str.replace(/\s+/g, '').replace('-', '');
};

export const encodeString = (toEncode: string) => {

    let charToEncode = toEncode.split('');
    let encodedString = "";

    for (let i = 0; i < charToEncode.length; i++) {
        const encodedChar = escape(charToEncode[i]);

        if (encodedChar.length == 3) {
            encodedString += encodedChar.replace("%", "_x00") + "_";
        }
        else if (encodedChar.length == 5) {
            encodedString += encodedChar.replace("%u", "_x") + "_";
        }
        else {
            encodedString += encodedChar;
        }
    }
    return encodedString;

};

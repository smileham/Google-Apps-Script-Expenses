function endsWith_(str, suffix) {
  return str.indexOf(suffix, str.length - suffix.length) !== -1;
}

function hash_(theString){
    var hash = 0, i, char;
    if (theString.length == 0) return hash;
    for (i = 0, l = theString.length; i < l; i++) {
        char  = theString.charCodeAt(i);
        hash  = ((hash<<5)-hash)+char;
        hash |= 0; // Convert to 32bit integer
    }
    return hash;
};

function makeLine(length) {
    var line = "";
    for (var j = 1; j <= length; j++) {
        line += "* ";
    }
    return line + "\n";
}
function buildTriangle(baseLength){
    var triangle = "";
    for (var i=1; i<=baseLength; i++){
        triangle = triangle + makeLine(i);
    }
    return triangle;
}
console.log(buildTriangle(10));
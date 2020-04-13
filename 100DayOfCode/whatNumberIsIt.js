// function whatNumberIsIt(n) {
// 	switch (n) {
// 		case Number.MAX_VALUE:
// 			console.log("Input number is Number.MAX_VALUE");
// 			break;
// 		case 5e-324:
// 			console.log("Input number is Number.MIN_VALUE");
// 			break;
// 		case Infinity:
// 			console.log("Input number is Number.POSITIVE_INFINITY");
// 			break;
// 		case -Infinity:
// 			console.log("Input number is Number.NEGATIVE_INFINITY");
// 			break;
// 		case NaN:
// 			console.log("Input number is Number.NaN");
// 		default: console.log("Input number is " + n);
// 	}
// }

// console.log(whatNumberIsIt(1 / 0));
// console.log(whatNumberIsIt(100));
// console.log(whatNumberIsIt(1.7976931348623157e+308));
// console.log(whatNumberIsIt(5e-324));
// console.log(whatNumberIsIt(-Number.MAX_VALUE * 2));
// console.log(whatNumberIsIt(NaN));
// console.log(whatNumberIsIt(Infinity + 1));


// function whatNumberIsIt(n) {
// 	switch (n) {
// 		case Number.MAX_VALUE:
// 			return "Input number is Number.MAX_VALUE";
// 			break;
// 		case 5e-324:
// 			return "Input number is Number.MIN_VALUE";
// 			break;
// 		case Infinity:
// 			return "Input number is Number.POSITIVE_INFINITY";
// 			break;
// 		case -Infinity:
// 			return "Input number is Number.NEGATIVE_INFINITY";
// 			break;
// 		case NaN:
// 			return "Input number is Number.NaN";
// 		default: return "Input number is " + n;
// 	}
// }

// console.log(whatNumberIsIt(1 / 0));
// console.log(whatNumberIsIt(100));
// console.log(whatNumberIsIt(1.7976931348623157e+308));
// console.log(whatNumberIsIt(5e-324));
// console.log(whatNumberIsIt(-Number.MAX_VALUE * 2));
// console.log(whatNumberIsIt(NaN));
// console.log(whatNumberIsIt(Infinity + 1));

function whatNumberIsIt(n) {
	if (n !== n) return "Input number is Number.NaN";
	switch (n) {
		case Number.MAX_VALUE:
			return "Input number is Number.MAX_VALUE";
			break;
		case 5e-324:
			return "Input number is Number.MIN_VALUE";
			break;
		case Infinity:
			return "Input number is Number.POSITIVE_INFINITY";
			break;
		case -Infinity:
			return "Input number is Number.NEGATIVE_INFINITY";
			break;
		//case Number.isNaN(n):
		//	return "Input number is Number.NaN";
		default: return "Input number is " + n;
	}
}
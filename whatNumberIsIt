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

whatNumberIsIt(NaN);
whatNumberIsIt(1/0);
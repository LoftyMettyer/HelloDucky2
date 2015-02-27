function replace(sExpression, sFind, sReplace) {
		//gi (global search, ignore case)
		var re = new RegExp(sFind, "gi");
		sExpression = sExpression.replace(re, sReplace);
		return (sExpression);
}

# Copyright (c) 2010-2024 openpyxl

"""
List of builtin formulae
"""

# Classic Excel functions (pre-365)
FORMULAE = (
    # Cube functions
    "CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY", "CUBERANKEDMEMBER",
    "CUBESET", "CUBESETCOUNT", "CUBEVALUE",
    # Database functions
    "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT",
    "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP",
    # Date/Time functions
    "DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS360", "EDATE", "EOMONTH",
    "HOUR", "MINUTE", "MONTH", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOW",
    "SECOND", "TIME", "TIMEVALUE", "TODAY", "WEEKDAY", "WEEKNUM", "WORKDAY",
    "WORKDAY.INTL", "YEAR", "YEARFRAC",
    # Engineering functions
    "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BIN2DEC", "BIN2HEX",
    "BIN2OCT", "COMPLEX", "CONVERT", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DELTA",
    "ERF", "ERFC", "GESTEP", "HEX2BIN", "HEX2DEC", "HEX2OCT", "IMABS",
    "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMDIV", "IMEXP",
    "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSIN",
    "IMSQRT", "IMSUB", "IMSUM", "OCT2BIN", "OCT2DEC", "OCT2HEX",
    # Financial functions
    "ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC", "COUPDAYBS", "COUPDAYS",
    "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "CUMIPMT", "CUMPRINC",
    "DB", "DDB", "DISC", "DOLLARDE", "DOLLARFR", "DURATION", "EFFECT", "FV",
    "FVSCHEDULE", "INTRATE", "IPMT", "IRR", "ISPMT", "MDURATION", "MIRR",
    "NOMINAL", "NPER", "NPV", "ODDFPRICE", "ODDFYIELD", "ODDLPRICE",
    "ODDLYIELD", "PMT", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PV",
    "RATE", "RECEIVED", "SLN", "SYD", "TBILLEQ", "TBILLPRICE", "TBILLYIELD",
    "VDB", "XIRR", "XNPV", "YIELD", "YIELDDISC", "YIELDMAT",
    # Information functions
    "CELL", "ERROR.TYPE", "INFO", "ISBLANK", "ISERR", "ISERROR", "ISEVEN",
    "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT",
    "N", "NA", "TYPE",
    # Logical functions
    "AND", "FALSE", "IF", "IFERROR", "NOT", "OR", "TRUE",
    # Lookup functions
    "ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS", "GETPIVOTDATA",
    "HLOOKUP", "HYPERLINK", "INDEX", "INDIRECT", "LOOKUP", "MATCH", "OFFSET",
    "ROW", "ROWS", "RTD", "TRANSPOSE", "VLOOKUP",
    # Math functions
    "ABS", "ACOS", "ACOSH", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH",
    "CEILING", "COMBIN", "COS", "COSH", "DEGREES", "ECMA.CEILING", "EVEN",
    "EXP", "FACT", "FACTDOUBLE", "FLOOR", "GCD", "INT", "ISO.CEILING", "LCM",
    "LN", "LOG", "LOG10", "MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND",
    "MULTINOMIAL", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS",
    "RAND", "RANDBETWEEN", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP",
    "SERIESSUM", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL", "SUM",
    "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2",
    "SUMXMY2", "TAN", "TANH", "TRUNC",
    # Statistical functions
    "AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BETADIST",
    "BETAINV", "BINOMDIST", "CHIDIST", "CHIINV", "CHITEST", "CONFIDENCE",
    "CORREL", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "COVAR",
    "CRITBINOM", "DEVSQ", "EXPONDIST", "FDIST", "FINV", "FISHER", "FISHERINV",
    "FORECAST", "FREQUENCY", "FTEST", "GAMMADIST", "GAMMAINV", "GAMMALN",
    "GEOMEAN", "GROWTH", "HARMEAN", "HYPGEOMDIST", "INTERCEPT", "KURT",
    "LARGE", "LINEST", "LOGEST", "LOGINV", "LOGNORMDIST", "MAX", "MAXA",
    "MEDIAN", "MIN", "MINA", "MODE", "NEGBINOMDIST", "NORMDIST", "NORMINV",
    "NORMSDIST", "NORMSINV", "PEARSON", "PERCENTILE", "PERCENTRANK", "PERMUT",
    "POISSON", "PROB", "QUARTILE", "RANK", "RSQ", "SKEW", "SLOPE", "SMALL",
    "STANDARDIZE", "STDEV", "STDEVA", "STDEVP", "STDEVPA", "STEYX", "TDIST",
    "TINV", "TREND", "TRIMMEAN", "TTEST", "VAR", "VARA", "VARP", "VARPA",
    "WEIBULL", "ZTEST",
    # Text functions
    "ASC", "BAHTTEXT", "CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR",
    "EXACT", "FIND", "FINDB", "FIXED", "JIS", "LEFT", "LEFTB", "LEN", "LENB",
    "LOWER", "MID", "MIDB", "PHONETIC", "PROPER", "REPLACE", "REPLACEB",
    "REPT", "RIGHT", "RIGHTB", "SEARCH", "SEARCHB", "SUBSTITUTE", "T", "TEXT",
    "TRIM", "UPPER", "VALUE",
)

# Excel 365 / Excel 2019+ functions
FORMULAE_365 = (
    # Dynamic array functions (Excel 365, 2021+)
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
    # Lambda functions (Excel 365)
    "LAMBDA", "LET", "MAKEARRAY", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL",
    "ISOMITTED",
    # Modern lookup functions
    "XLOOKUP", "XMATCH",
    # Modern text functions
    "TEXTJOIN", "CONCAT", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "VALUETOTEXT", "ARRAYTOTEXT",
    # Modern logical functions
    "SWITCH", "IFS", "XOR", "IFNA",
    # Conditional aggregate functions
    "MAXIFS", "MINIFS",
    # Date functions
    "DAYS", "ISOWEEKNUM",
    # Information functions
    "ISFORMULA", "SHEET", "SHEETS",
    # Web functions
    "WEBSERVICE", "ENCODEURL", "FILTERXML",
    # Image function
    "IMAGE",
    # Rich data functions
    "FIELDVALUE", "STOCKHISTORY",
    # Aggregation functions (Excel 365)
    "GROUPBY", "PIVOTBY", "PERCENTOF",
    # Other modern functions
    "CEILING.MATH", "FLOOR.MATH", "CEILING.PRECISE", "FLOOR.PRECISE",
    "AGGREGATE", "COMBINA", "PERMUTATIONA", "BASE", "DECIMAL",
    "ARABIC", "FORMULATEXT", "NUMBERVALUE",
    "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
    "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
    "CONFIDENCE.NORM", "CONFIDENCE.T", "COVARIANCE.P", "COVARIANCE.S",
    "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
    "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "FORECAST.LINEAR",
    "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
    "GAUSS", "HYPGEOM.DIST",
    "LOGNORM.DIST", "LOGNORM.INV",
    "MODE.MULT", "MODE.SNGL",
    "NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
    "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC",
    "PHI", "POISSON.DIST",
    "QUARTILE.EXC", "QUARTILE.INC",
    "RANK.AVG", "RANK.EQ",
    "SKEW.P",
    "STDEV.P", "STDEV.S",
    "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST",
    "VAR.P", "VAR.S",
    "WEIBULL.DIST", "Z.TEST",
    # Error handling
    "IFNA",
    # Unicode text
    "UNICHAR", "UNICODE",
    # Bitwise functions
    "BITAND", "BITOR", "BITXOR", "BITLSHIFT", "BITRSHIFT",
    # Trigonometric
    "ACOT", "ACOTH", "COT", "COTH", "CSC", "CSCH", "SEC", "SECH",
    # Engineering
    "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN",
    "ERF.PRECISE", "ERFC.PRECISE",
    # Compatibility renamed functions
    "CEILING.PRECISE", "FLOOR.PRECISE",
)

FORMULAE = frozenset(FORMULAE + FORMULAE_365)


from openpyxl.formula import Tokenizer


def validate(formula):
    """
    Utility function for checking whether a formula is syntactically correct
    """
    assert formula.startswith("=")
    formula = Tokenizer(formula)
    for t in formula.items:
        if t.type == "FUNC" and t.subtype == "OPEN":
            if not t.value.startswith("_xlfn.") and t.value[:-1] not in FORMULAE:
                raise ValueError(f"Unknown function {t.value} in {formula.formula}. The function may need a prefix")

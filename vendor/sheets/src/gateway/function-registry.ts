/// Function registry for the Excel storage contract (see formula-storage.ts).
///
/// Excel stores post-2007 function names in file formulas behind a marker:
/// `_xlfn.` for most, `_xlfn._xlws.` for the worksheet-scope dynamic-array
/// functions FILTER and SORT. The UI shows the plain name. This module is the
/// single source for which names carry which marker; future-functions.ts (the
/// vendored writer) re-exports the two name lists below.
///
/// The ACTIVE lists are the names the writer has always prefixed plus the
/// names Excel 16.0.20326 was observed to store behind a marker (ARRAYTOTEXT,
/// GROUPBY, REGEX*, TRANSLATE, ...). Names whose marker could not be confirmed
/// against Excel (COPILOT, which that build does not recognise) are registered
/// with verified:false and are never rewritten. Names that exist only in
/// Google Sheets are registered as "not-excel" so callers can warn; they are
/// never rewritten either.
///
/// CLASSIC_FUNCTION_NAMES lists the built-in functions Excel stores without a
/// marker. Together with the marked names it answers "is this identifier a
/// built-in function?", which the converter needs to tell a built-in call from
/// a LET/LAMBDA variable call, and a bare function name (`SUM`) passed as a
/// LAMBDA argument (stored `_xleta.SUM`) from a defined name.

export type ExcelAvailability =
  | "2007"
  | "2010"
  | "2013"
  | "2016"
  | "2019"
  | "2021"
  | "365"
  | "not-excel";

export type StoragePrefix = "" | "_xlfn." | "_xlfn._xlws.";

export interface FunctionSpec {
  /// Canonical uppercase name, no storage prefix.
  readonly name: string;
  readonly storage: StoragePrefix;
  readonly excel: ExcelAvailability;
  readonly volatile: boolean;
  readonly returnsArray: "never" | "maybe" | "always";
  /// Set on the functions whose arguments introduce `_xlpm.` names.
  readonly scope?: "LET" | "LAMBDA";
  /// Argument that takes a LAMBDA or a bare function name (`SUM`, stored
  /// `_xleta.SUM`): its index, or "last" when the lambda closes the list.
  readonly lambdaArg?: number | "last";
  /// True when the storage prefix is applied today. False means the prefix is
  /// not confirmed by an Excel-authored fixture: the writer leaves it alone.
  readonly verified: boolean;
}

export interface FunctionRegistry {
  get(name: string): FunctionSpec | undefined;
  /// Whether `name` is a built-in worksheet function that can be called as
  /// `name(...)`: a classic function or a marked one. Google-only names, names
  /// with an unconfirmed marker and the operator forms SINGLE and ANCHORARRAY
  /// are not built-ins.
  isBuiltin(name: string): boolean;
  /// Whether the live calculation engine implements the function. Filled by
  /// the caller from the engine's own function list; false when unknown.
  engineSupports(name: string): boolean;
}

/// Worksheet-scope dynamic-array functions: stored as `_xlfn._xlws.NAME`.
export const XLWS_FUNCTION_NAMES: readonly string[] = ["FILTER", "SORT"];

const XLFN_2010: readonly string[] = [
  "AGGREGATE",
  "BETA.DIST",
  "BETA.INV",
  "BINOM.DIST",
  "BINOM.INV",
  "CEILING.PRECISE",
  "CHISQ.DIST",
  "CHISQ.DIST.RT",
  "CHISQ.INV",
  "CHISQ.INV.RT",
  "CHISQ.TEST",
  "CONFIDENCE.NORM",
  "CONFIDENCE.T",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "ERF.PRECISE",
  "ERFC.PRECISE",
  "EXPON.DIST",
  "F.DIST",
  "F.DIST.RT",
  "F.INV",
  "F.INV.RT",
  "F.TEST",
  "FLOOR.PRECISE",
  "GAMMA.DIST",
  "GAMMA.INV",
  "GAMMALN.PRECISE",
  "HYPGEOM.DIST",
  "LOGNORM.DIST",
  "LOGNORM.INV",
  "MODE.MULT",
  "MODE.SNGL",
  "NEGBINOM.DIST",
  "NETWORKDAYS.INTL",
  "NORM.DIST",
  "NORM.INV",
  "NORM.S.DIST",
  "NORM.S.INV",
  "PERCENTILE.EXC",
  "PERCENTILE.INC",
  "PERCENTRANK.EXC",
  "PERCENTRANK.INC",
  "POISSON.DIST",
  "QUARTILE.EXC",
  "QUARTILE.INC",
  "RANK.AVG",
  "RANK.EQ",
  "STDEV.P",
  "STDEV.S",
  "T.DIST",
  "T.DIST.2T",
  "T.DIST.RT",
  "T.INV",
  "T.INV.2T",
  "T.TEST",
  "VAR.P",
  "VAR.S",
  "WEIBULL.DIST",
  "WORKDAY.INTL",
  "Z.TEST",
];

const XLFN_2013: readonly string[] = [
  "ACOT",
  "ACOTH",
  "ARABIC",
  "BASE",
  "BINOM.DIST.RANGE",
  "BITAND",
  "BITLSHIFT",
  "BITOR",
  "BITRSHIFT",
  "BITXOR",
  "CEILING.MATH",
  "COMBINA",
  "COT",
  "COTH",
  "CSC",
  "CSCH",
  "DAYS",
  "DECIMAL",
  "ENCODEURL",
  "FILTERXML",
  "FLOOR.MATH",
  "FORMULATEXT",
  "GAMMA",
  "GAUSS",
  "IFNA",
  "IMCOSH",
  "IMCOT",
  "IMCSC",
  "IMCSCH",
  "IMSEC",
  "IMSECH",
  "IMSINH",
  "IMTAN",
  "ISFORMULA",
  "ISOWEEKNUM",
  "MUNIT",
  "NUMBERVALUE",
  "PDURATION",
  "PERMUTATIONA",
  "PHI",
  "RRI",
  "SEC",
  "SECH",
  "SHEET",
  "SHEETS",
  "SKEW.P",
  "UNICHAR",
  "UNICODE",
  "WEBSERVICE",
  "XOR",
];

const XLFN_2016: readonly string[] = [
  "CONCAT",
  "FORECAST.ETS",
  "FORECAST.ETS.CONFINT",
  "FORECAST.ETS.SEASONALITY",
  "FORECAST.ETS.STAT",
  "FORECAST.LINEAR",
  "IFS",
  "MAXIFS",
  "MINIFS",
  "SWITCH",
  "TEXTJOIN",
];

const XLFN_365: readonly string[] = [
  "BYCOL",
  "BYROW",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "DROP",
  "EXPAND",
  "HSTACK",
  "IMAGE",
  "ISOMITTED",
  "LAMBDA",
  "LET",
  "MAKEARRAY",
  "MAP",
  "RANDARRAY",
  "REDUCE",
  "SCAN",
  "SEQUENCE",
  "SORTBY",
  "STOCKHISTORY",
  "TAKE",
  "TEXTAFTER",
  "TEXTBEFORE",
  "TEXTSPLIT",
  "TOCOL",
  "TOROW",
  "UNIQUE",
  "VSTACK",
  "WRAPCOLS",
  "WRAPROWS",
  "XLOOKUP",
  "XMATCH",
];

/// Functions added after the lists above, each confirmed by typing the call in
/// Excel 16.0.20326 and reading the stored formula (all store `_xlfn.`).
const XLFN_LATE: readonly string[] = [
  "ARRAYTOTEXT",
  "DETECTLANGUAGE",
  "FIELDVALUE",
  "GROUPBY",
  "PERCENTOF",
  "PIVOTBY",
  "REGEXEXTRACT",
  "REGEXREPLACE",
  "REGEXTEST",
  "TRANSLATE",
  "TRIMRANGE",
  "VALUETOTEXT",
];

/// Storage forms of the `@` implicit-intersection and `#` spill operators.
const XLFN_OPERATOR_FORMS: readonly string[] = ["SINGLE", "ANCHORARRAY"];

/// Every name the writer prefixes with `_xlfn.`, in the original list order.
/// NETWORKDAYS.INTL and WORKDAY.INTL are here for history: Excel stores them
/// without a marker but reads the marked form to the same value.
export const XLFN_FUNCTION_NAMES: readonly string[] = [
  ...XLFN_2010,
  ...XLFN_2013,
  ...XLFN_2016,
  ...XLFN_365,
  ...XLFN_LATE,
  ...XLFN_OPERATOR_FORMS,
];

/// Markers Excel 16.0.20326 could not confirm: COPILOT is not recognised by
/// that build, so what it stores is unknown. Registered so callers can warn,
/// but never prefixed.
const UNVERIFIED_PREFIX_NAMES: readonly string[] = ["COPILOT"];

/// Google Sheets functions with no Excel equivalent.
const NOT_EXCEL_NAMES: readonly string[] = [
  "ARRAYFORMULA",
  "GOOGLEFINANCE",
  "IMPORTRANGE",
  "QUERY",
  "REGEXMATCH",
  "SPARKLINE",
];

/// Built-in worksheet functions that Excel 16.0.20326 (en-US) stores without a
/// marker. Source: a candidate list (the function-name table inside EXCEL.EXE,
/// the Univer engine-formula FUNCTION_NAMES_* enums, the documented function
/// list) typed into Excel as lowercase calls and kept when Excel upper-cased the
/// name and stored it without `_xlfn.`/`_xlws.`. Excel 4 macro functions,
/// Google Sheets names and add-in functions (EUROCONVERT) are excluded because
/// a cell cannot call them. Limits: a function added by a newer Excel build is
/// missing until the list is regenerated, and JIS (Japanese editions only) is
/// absent because that build does not recognise it.
export const CLASSIC_FUNCTION_NAMES: readonly string[] = `
ABS ACCRINT ACCRINTM ACOS ACOSH ADDRESS AMORDEGRC AMORLINC AND AREAS ASC ASIN ASINH ATAN ATAN2
ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BAHTTEXT BESSELI BESSELJ BESSELK BESSELY
BETADIST BETAINV BIN2DEC BIN2HEX BIN2OCT BINOMDIST CEILING CELL CHAR CHIDIST CHIINV CHITEST
CHOOSE CLEAN CODE COLUMN COLUMNS COMBIN COMPLEX CONCATENATE CONFIDENCE CONVERT CORREL COS COSH
COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD
COVAR CRITBINOM CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET
CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS360 DB DBCS
DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DEGREES DELTA DEVSQ DGET DISC DMAX DMIN DOLLAR
DOLLARDE DOLLARFR DPRODUCT DSTDEV DSTDEVP DSUM DURATION DVAR DVARP ECMA.CEILING EDATE EFFECT
EOMONTH ERF ERFC ERROR.TYPE EVEN EXACT EXP EXPONDIST FACT FACTDOUBLE FALSE FDIST FIND FINDB
FINV FISHER FISHERINV FIXED FLOOR FORECAST FREQUENCY FTEST FV FVSCHEDULE GAMMADIST GAMMAINV
GAMMALN GCD GEOMEAN GESTEP GETPIVOTDATA GROWTH HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR
HYPERLINK HYPGEOMDIST IF IFERROR IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMDIV IMEXP IMLN
IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSIN IMSQRT IMSUB IMSUM INDEX INDIRECT INFO INT
INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER
ISO.CEILING ISODD ISPMT ISREF ISTEXT ISTHAIDIGIT KURT LARGE LCM LEFT LEFTB LEN LENB LINEST LN
LOG LOG10 LOGEST LOGINV LOGNORMDIST LOOKUP LOWER MATCH MAX MAXA MDETERM MDURATION MEDIAN MID
MIDB MIN MINA MINUTE MINVERSE MIRR MMULT MOD MODE MONTH MROUND MULTINOMIAL N NA NEGBINOMDIST
NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV
NUMBERSTRING OCT2BIN OCT2DEC OCT2HEX ODD ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR
PEARSON PERCENTILE PERCENTRANK PERMUT PHONETIC PI PMT POISSON POWER PPMT PRICE PRICEDISC
PRICEMAT PROB PRODUCT PROPER PV QUARTILE QUOTIENT RADIANS RAND RANDBETWEEN RANK RATE RECEIVED
REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDBAHTDOWN ROUNDBAHTUP ROUNDDOWN ROUNDUP ROW
ROWS RSQ RTD SEARCH SEARCHB SECOND SERIESSUM SIGN SIN SINH SKEW SLN SLOPE SMALL SQRT SQRTPI
STANDARDIZE STDEV STDEVA STDEVP STDEVPA STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT
SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SYD T TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT
THAIDAYOFWEEK THAIDIGIT THAIMONTHOFYEAR THAINUMSOUND THAINUMSTRING THAISTRINGLENGTH THAIYEAR
TIME TIMEVALUE TINV TODAY TRANSPOSE TREND TRIM TRIMMEAN TRUE TRUNC TTEST TYPE UPPER VALUE VAR
VARA VARP VARPA VDB VLOOKUP WEEKDAY WEEKNUM WEIBULL WORKDAY WORKDAY.INTL XIRR XNPV YEAR
YEARFRAC YIELD YIELDDISC YIELDMAT ZTEST
`
  .split(/\s+/)
  .filter((name) => name !== "");

/// Where each lambda-taking function expects its LAMBDA / bare function name.
const LAMBDA_ARG: Readonly<Record<string, number | "last">> = {
  BYCOL: 1,
  BYROW: 1,
  GROUPBY: 2,
  MAKEARRAY: 2,
  MAP: "last",
  PIVOTBY: 3,
  REDUCE: "last",
  SCAN: "last",
};

const BUILTIN_NAMES: ReadonlySet<string> = new Set([
  ...CLASSIC_FUNCTION_NAMES,
  ...XLFN_FUNCTION_NAMES.filter((name) => !XLFN_OPERATOR_FORMS.includes(name)),
  ...XLWS_FUNCTION_NAMES,
]);

const ALWAYS_ARRAY: ReadonlySet<string> = new Set([
  "ANCHORARRAY",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "DROP",
  "EXPAND",
  "FILTER",
  "GROUPBY",
  "HSTACK",
  "PIVOTBY",
  "RANDARRAY",
  "SEQUENCE",
  "SORT",
  "SORTBY",
  "TAKE",
  "TEXTSPLIT",
  "TOCOL",
  "TOROW",
  "TRIMRANGE",
  "UNIQUE",
  "VSTACK",
  "WRAPCOLS",
  "WRAPROWS",
]);

const MAYBE_ARRAY: ReadonlySet<string> = new Set([
  "BYCOL",
  "BYROW",
  "FIELDVALUE",
  "LAMBDA",
  "LET",
  "MAKEARRAY",
  "MAP",
  "REGEXEXTRACT",
  "REGEXREPLACE",
  "REGEXTEST",
  "SCAN",
]);

const VOLATILE: ReadonlySet<string> = new Set(["RANDARRAY"]);

function arrayKind(name: string): FunctionSpec["returnsArray"] {
  if (ALWAYS_ARRAY.has(name)) return "always";
  if (MAYBE_ARRAY.has(name)) return "maybe";
  return "never";
}

function spec(
  name: string,
  storage: StoragePrefix,
  excel: ExcelAvailability,
  verified: boolean,
): FunctionSpec {
  const base = {
    name,
    storage,
    excel,
    volatile: VOLATILE.has(name),
    returnsArray: arrayKind(name),
    verified,
  };
  if (name === "LET") return { ...base, scope: "LET" };
  if (name === "LAMBDA") return { ...base, scope: "LAMBDA" };
  const lambdaArg = LAMBDA_ARG[name];
  return lambdaArg === undefined ? base : { ...base, lambdaArg };
}

function buildSpecs(): ReadonlyMap<string, FunctionSpec> {
  const specs = new Map<string, FunctionSpec>();
  // Availability labels follow the section comments of the original list.
  for (const name of XLFN_2010) specs.set(name, spec(name, "_xlfn.", "2010", true));
  for (const name of XLFN_2013) specs.set(name, spec(name, "_xlfn.", "2013", true));
  for (const name of XLFN_2016) specs.set(name, spec(name, "_xlfn.", "2016", true));
  for (const name of XLFN_365) specs.set(name, spec(name, "_xlfn.", "365", true));
  for (const name of XLFN_LATE) specs.set(name, spec(name, "_xlfn.", "365", true));
  for (const name of XLFN_OPERATOR_FORMS) specs.set(name, spec(name, "_xlfn.", "365", true));
  for (const name of XLWS_FUNCTION_NAMES) specs.set(name, spec(name, "_xlfn._xlws.", "365", true));
  for (const name of UNVERIFIED_PREFIX_NAMES) {
    specs.set(name, spec(name, "_xlfn.", "365", false));
  }
  for (const name of NOT_EXCEL_NAMES) specs.set(name, spec(name, "", "not-excel", true));
  return specs;
}

const SPECS = buildSpecs();

/// `engineNames` are uppercase function names the live engine implements.
export function createFunctionRegistry(engineNames: ReadonlySet<string>): FunctionRegistry {
  return {
    get: (name) => SPECS.get(name.toUpperCase()),
    isBuiltin: (name) => BUILTIN_NAMES.has(name.toUpperCase()),
    engineSupports: (name) => engineNames.has(name.toUpperCase()),
  };
}

/// Registry with no engine information (engineSupports is always false).
export const DEFAULT_FUNCTION_REGISTRY: FunctionRegistry = createFunctionRegistry(new Set());

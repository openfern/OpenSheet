namespace OpenSheet.Tests.Oracle
{
    /// <summary>
    /// Formulas that are checked against Excel. They read the fixed data block written by <see cref="FormulaOracleTests"/>:
    /// A1:A8 numbers 3,1,4,1,5,9,2,6 · B1:B8 words · C1:C4 10..40 · D1:D4 x,y,x,z · E1 a date serial · F1:F4 odd texts ·
    /// H1:J4 a lookup table. Nothing here is volatile (no NOW, TODAY, RAND).
    /// </summary>
    static class FormulaBank
    {
        public static readonly string[] Formulas =
        {
            // operators and literals
            "=1+2*3^2", "=-2^2", "=10/4", "=10/0", "=\"a\"&1", "=1=1", "=(\"a\"<\"b\")", "=5%", "=2^-1", "=A1+B1", "=B1&A1", "=A1:A3", "=--TRUE", "=\"b\">\"a\"", "=7\\2",
            // math
            "=ABS(-5)", "=ABS(A1-A6)", "=ACOS(0.5)", "=ACOSH(2)", "=ASIN(0.5)", "=ASINH(1)", "=ATAN(1)", "=ATAN2(1,2)", "=ATANH(0.5)",
            "=CEILING(4.3,1)", "=CEILING(-4.3,2)", "=CEILING.MATH(4.3)", "=COMBIN(8,3)", "=COS(1)", "=COSH(1)", "=DEGREES(PI())", "=EVEN(3)",
            "=EXP(2)", "=FACT(5)", "=FACTDOUBLE(7)", "=FLOOR(4.7,1)", "=FLOOR(-4.7,2)", "=FLOOR.MATH(4.7)", "=GCD(12,18)", "=INT(-4.5)",
            "=LCM(4,6)", "=LN(10)", "=LOG(100,10)", "=LOG(8,2)", "=LOG10(1000)", "=MOD(10,3)", "=MOD(-10,3)", "=MROUND(10,3)", "=MULTINOMIAL(2,3,4)",
            "=ODD(2)", "=PI()", "=POWER(2,10)", "=PRODUCT(A1:A4)", "=QUOTIENT(10,3)", "=RADIANS(180)", "=ROUND(2.567,2)", "=ROUND(-2.5,0)",
            "=ROUND(1234.5,-2)", "=ROUNDDOWN(2.99,1)", "=ROUNDUP(2.01,1)", "=SIGN(-3)", "=SIN(1)", "=SINH(1)", "=SQRT(16)", "=SQRT(-1)",
            "=SQRTPI(2)", "=SUM(A1:A8)", "=SUM(A1:A3,10)", "=SUMIF(A1:A8,\">3\")", "=SUMIF(D1:D4,\"x\",C1:C4)", "=SUMIFS(C1:C4,D1:D4,\"x\")",
            "=SUMPRODUCT(A1:A4,C1:C4)", "=SUMPRODUCT((A1:A8>2)*(A1:A8<9))", "=SUMSQ(A1:A3)", "=SUMX2MY2(A1:A3,A4:A6)", "=SUMX2PY2(A1:A3,A4:A6)", "=SUMXMY2(A1:A3,A4:A6)",
            "=TAN(1)", "=TANH(1)", "=TRUNC(8.9)", "=TRUNC(-8.9)", "=TRUNC(3.14159,2)", "=ARABIC(\"XIV\")", "=ROMAN(14)", "=BASE(255,16)", "=DECIMAL(\"FF\",16)",
            "=SEC(1)", "=CSC(1)", "=COT(1)", "=AGGREGATE(9,6,A1:A8)", "=AGGREGATE(14,6,A1:A8,2)", "=SUBTOTAL(9,A1:A8)", "=SUBTOTAL(1,A1:A8)", "=SERIESSUM(2,0,1,A1:A3)", "=SUMIF(B1:B8,\"*a*\",A1:A8)",
            "=MAX(A1:A8)-MIN(A1:A8)", "=COMBINA(4,2)", "=ISO.CEILING(4.3)", "=SIGN(0)", "=PERMUT(5,2)", "=PERMUTATIONA(3,2)", "=RANDARRAY(1)*0", "=MMULT(A1:B2,A1:B2)",
            // statistics
            "=AVERAGE(A1:A8)", "=AVERAGEA(A1:B8)", "=AVERAGEIF(A1:A8,\">2\")", "=AVERAGEIFS(C1:C4,D1:D4,\"x\")", "=COUNT(A1:A8)", "=COUNTA(A1:B8)", "=COUNTBLANK(A1:A9)",
            "=COUNTIF(A1:A8,1)", "=COUNTIF(B1:B8,\"*a*\")", "=COUNTIF(A1:A8,\">=4\")", "=COUNTIFS(A1:A8,\">1\",A1:A8,\"<6\")", "=MAX(A1:A8)", "=MIN(A1:A8)", "=MAXA(A1:A8)", "=MINA(A1:A8)",
            "=MEDIAN(A1:A8)", "=MODE(A1:A8)", "=MODE.SNGL(A1:A8)", "=LARGE(A1:A8,2)", "=SMALL(A1:A8,2)", "=RANK(5,A1:A8)", "=RANK.EQ(5,A1:A8,1)", "=RANK.AVG(1,A1:A8)",
            "=PERCENTILE(A1:A8,0.3)", "=PERCENTILE.INC(A1:A8,0.3)", "=PERCENTILE.EXC(A1:A8,0.3)", "=PERCENTRANK(A1:A8,4)", "=QUARTILE(A1:A8,1)", "=QUARTILE.EXC(A1:A8,1)",
            "=STDEV(A1:A8)", "=STDEV.P(A1:A8)", "=STDEV.S(A1:A8)", "=STDEVP(A1:A8)", "=VAR(A1:A8)", "=VAR.P(A1:A8)", "=VAR.S(A1:A8)", "=VARP(A1:A8)", "=DEVSQ(A1:A8)", "=AVEDEV(A1:A8)",
            "=GEOMEAN(A1:A8)", "=HARMEAN(A1:A8)", "=TRIMMEAN(A1:A8,0.2)", "=CORREL(A1:A4,C1:C4)", "=COVAR(A1:A4,C1:C4)", "=COVARIANCE.S(A1:A4,C1:C4)", "=COVARIANCE.P(A1:A4,C1:C4)",
            "=SLOPE(C1:C4,A1:A4)", "=INTERCEPT(C1:C4,A1:A4)", "=RSQ(C1:C4,A1:A4)", "=FORECAST(5,C1:C4,A1:A4)", "=FORECAST.LINEAR(5,C1:C4,A1:A4)", "=STEYX(C1:C4,A1:A4)",
            "=SKEW(A1:A8)", "=KURT(A1:A8)", "=STANDARDIZE(5,3,2)", "=GAMMALN(5)", "=GAMMA(5)", "=FISHER(0.5)", "=FISHERINV(0.5)", "=MAXIFS(C1:C4,D1:D4,\"x\")", "=MINIFS(C1:C4,D1:D4,\"x\")",
            "=NORM.DIST(1,0,1,TRUE)", "=NORM.S.DIST(1,TRUE)", "=NORM.INV(0.9,0,1)", "=NORM.S.INV(0.9)", "=NORMDIST(1,0,1,FALSE)", "=NORMSDIST(1)", "=BINOM.DIST(3,10,0.5,FALSE)", "=BINOMDIST(3,10,0.5,TRUE)",
            "=POISSON.DIST(2,3,FALSE)", "=EXPON.DIST(1,2,TRUE)", "=T.DIST(1,5,TRUE)", "=T.INV.2T(0.05,10)", "=CHISQ.DIST(2,3,TRUE)", "=CHISQ.INV(0.5,3)", "=F.DIST(1,2,3,TRUE)", "=CONFIDENCE.NORM(0.05,1,100)",
            "=GAUSS(1)", "=PHI(1)", "=LOGNORM.DIST(1,0,1,TRUE)", "=WEIBULL.DIST(1,2,3,TRUE)", "=GAMMA.DIST(2,3,2,TRUE)", "=BETA.DIST(0.5,2,3,TRUE)", "=HYPGEOM.DIST(1,4,4,10,FALSE)", "=NEGBINOM.DIST(2,3,0.5,FALSE)",
            // text
            "=LEFT(B1,3)", "=RIGHT(B1,2)", "=MID(B2,2,3)", "=LEN(B1)", "=UPPER(B1)", "=LOWER(B2)", "=PROPER(\"hello wORLD\")", "=TRIM(F1)", "=CLEAN(\"a\"&CHAR(7)&\"b\")",
            "=CONCATENATE(B1,\"-\",B2)", "=CONCAT(B1:B3)", "=TEXTJOIN(\",\",TRUE,B1:B3)", "=TEXTJOIN(\"-\",FALSE,B1:B3)", "=SUBSTITUTE(F2,\"o\",\"0\")", "=SUBSTITUTE(F2,\"o\",\"0\",2)",
            "=REPLACE(B1,1,2,\"XY\")", "=FIND(\"l\",F2)", "=SEARCH(\"L\",F2)", "=FIND(\"zz\",B1)", "=SEARCH(\"?o\",F2)", "=TEXT(1234.5,\"#,##0.00\")", "=TEXT(0.256,\"0.0%\")", "=TEXT(E1,\"0000000\")",
            "=TEXT(5,\"000\")", "=TEXT(-3.14159,\"0.00\")", "=TEXT(1234567,\"#,##0\")", "=TEXT(0.5,\"# ?/?\")", "=TEXT(1234.5,\"0.00E+00\")", "=VALUE(\"123\")", "=VALUE(\"1,234.5\")", "=VALUE(\"12%\")",
            "=NUMBERVALUE(\"1.234,5\",\",\",\".\")", "=REPT(\"ab\",3)", "=EXACT(\"a\",\"A\")", "=CHAR(65)", "=CODE(\"A\")", "=UNICHAR(8364)", "=UNICODE(\"€\")", "=T(\"x\")", "=T(5)", "=N(\"5\")", "=N(TRUE)",
            "=FIXED(1234.567,1)", "=TEXTBEFORE(F2,\",\")", "=TEXTAFTER(F2,\",\")", "=LEN(F1)", "=LEFT(F2,FIND(\",\",F2)-1)", "=B1&\" \"&B2", "=UPPER(LEFT(B1,1))&MID(B1,2,99)",
            "=LENB(\"abc\")", "=ASC(\"abc\")", "=VALUETOTEXT(A1)", "=TEXTSPLIT(\"a,b\",\",\")",
            // logical and information
            "=IF(A1>2,\"y\",\"n\")", "=IF(A1>5,\"a\")", "=IFS(A1>5,\"a\",A1>2,\"b\",TRUE,\"c\")", "=AND(TRUE,FALSE)", "=AND(A1:A3>0)", "=OR(FALSE,FALSE)", "=NOT(TRUE)", "=XOR(TRUE,TRUE)",
            "=IFERROR(1/0,\"err\")", "=IFERROR(5,\"err\")", "=IFNA(NA(),\"na\")", "=SWITCH(A1,3,\"three\",1,\"one\",\"other\")", "=TRUE()", "=FALSE()",
            "=ISNUMBER(A1)", "=ISNUMBER(B1)", "=ISTEXT(B1)", "=ISBLANK(Z1)", "=ISERROR(1/0)", "=ISERR(NA())", "=ISNA(NA())", "=ISLOGICAL(TRUE)", "=ISEVEN(4)", "=ISODD(3)", "=ISREF(A1)", "=ISNONTEXT(A1)",
            "=NA()", "=TYPE(A1)", "=TYPE(B1)", "=TYPE(TRUE)", "=ERROR.TYPE(1/0)", "=ERROR.TYPE(NA())", "=N(1)", "=ROW(A5)", "=COLUMN(C1)", "=ROWS(A1:A8)", "=COLUMNS(A1:C1)", "=SHEET()", "=SHEETS()", "=ISFORMULA(A1)", "=ISOMITTED(1)",
            // lookup and reference
            "=VLOOKUP(2,H1:J4,2,FALSE)", "=VLOOKUP(2.5,H1:J4,3,TRUE)", "=VLOOKUP(9,H1:J4,2,FALSE)", "=HLOOKUP(\"x\",D1:D4,1,FALSE)", "=INDEX(A1:A8,3)", "=INDEX(H1:J4,2,3)", "=MATCH(4,A1:A8,0)",
            "=MATCH(4.5,H1:H4,1)", "=MATCH(\"cherry\",B1:B8,0)", "=MATCH(\"c*\",B1:B8,0)", "=XLOOKUP(3,H1:H4,I1:I4)", "=XLOOKUP(9,H1:H4,I1:I4,\"none\")", "=XMATCH(3,H1:H4)", "=CHOOSE(2,\"a\",\"b\",\"c\")",
            "=LOOKUP(3,H1:H4,I1:I4)", "=OFFSET(A1,2,0)", "=SUM(OFFSET(A1,0,0,3,1))", "=INDIRECT(\"A3\")", "=ADDRESS(2,3)", "=ADDRESS(2,3,4)", "=AREAS(A1:B3)", "=INDEX(A1:A8,MATCH(MAX(A1:A8),A1:A8,0))",
            "=SUM(INDEX(H1:J4,0,3))", "=ROW(A1:A3)", "=COLUMN()", "=ROW()", "=FORMULATEXT(A1)", "=HYPERLINK(\"http://x\",\"link\")", "=TRANSPOSE(A1:A3)", "=SORT(A1:A4)", "=UNIQUE(D1:D4)", "=FILTER(A1:A8,A1:A8>4)", "=SEQUENCE(3)",
            // dates and times
            "=DATE(2024,3,9)", "=DATE(2024,14,1)", "=YEAR(E1)", "=MONTH(E1)", "=DAY(E1)", "=WEEKDAY(E1)", "=WEEKDAY(E1,2)", "=WEEKNUM(E1)", "=ISOWEEKNUM(E1)", "=EDATE(E1,2)", "=EOMONTH(E1,0)", "=EOMONTH(E1,-3)",
            "=DATEVALUE(\"2024-03-09\")", "=TIME(13,30,0)", "=HOUR(0.75)", "=MINUTE(0.756)", "=SECOND(0.7567)", "=DAYS(E1+10,E1)", "=DAYS360(E1,E1+100)", "=NETWORKDAYS(E1,E1+14)", "=WORKDAY(E1,10)",
            "=DATEDIF(E1,E1+400,\"Y\")", "=DATEDIF(E1,E1+400,\"M\")", "=DATEDIF(E1,E1+400,\"D\")", "=YEARFRAC(E1,E1+180)", "=TIMEVALUE(\"13:30\")", "=E1+1", "=E1-DATE(2020,1,1)", "=NETWORKDAYS.INTL(E1,E1+14,11)", "=WORKDAY.INTL(E1,10,11)",
            // financial
            "=PMT(0.05/12,60,10000)", "=FV(0.05,10,-100)", "=PV(0.05,10,-100)", "=NPV(0.1,-100,50,60,70)", "=IRR({-100,30,40,50,60})", "=RATE(60,-200,10000)", "=NPER(0.05/12,-200,10000)", "=IPMT(0.05/12,1,60,10000)",
            "=PPMT(0.05/12,1,60,10000)", "=SLN(1000,100,5)", "=SYD(1000,100,5,1)", "=DB(1000,100,5,1)", "=DDB(1000,100,5,1)", "=EFFECT(0.05,12)", "=NOMINAL(0.05,12)", "=CUMIPMT(0.05/12,60,10000,1,12,0)", "=MIRR({-100,30,40,50,60},0.1,0.12)",
            "=XNPV(0.1,{-100,50,60},{43831,44000,44200})", "=XIRR({-100,50,60},{43831,44000,44200})", "=FVSCHEDULE(100,{0.1,0.05})", "=PDURATION(0.05,100,200)", "=RRI(10,100,200)", "=ISPMT(0.05,1,5,1000)",
            // engineering
            "=BIN2DEC(\"1010\")", "=DEC2BIN(10)", "=DEC2BIN(10,8)", "=DEC2HEX(255)", "=HEX2DEC(\"FF\")", "=OCT2DEC(\"17\")", "=DEC2OCT(64)", "=BITAND(12,10)", "=BITOR(12,10)", "=BITXOR(12,10)", "=BITLSHIFT(1,4)", "=BITRSHIFT(16,2)",
            "=DELTA(1,1)", "=GESTEP(5,3)", "=COMPLEX(3,4)", "=IMABS(\"3+4i\")", "=IMSUM(\"1+i\",\"2+3i\")", "=CONVERT(1,\"m\",\"ft\")", "=CONVERT(100,\"C\",\"F\")", "=ERF(1)", "=ERFC(1)",
            // second pass: the behaviors the documentation claims
            "=ROMAN(499,0)", "=ROMAN(499,1)", "=ROMAN(499,2)", "=ROMAN(499,3)", "=ROMAN(499,4)", "=ROMAN(1999,1)", "=ROMAN(3999,4)", "=ROMAN(0)", "=ROMAN(4000)",
            "=DEC2BIN(-1)", "=DEC2HEX(-1)", "=DEC2OCT(-8)", "=DEC2HEX(255,6)", "=DEC2BIN(512)", "=DEC2OCT(8,1)",
            "=FLOOR(4.7,-2)", "=FLOOR(-4.7,-2)", "=FLOOR(0,2)", "=FLOOR(4.7,0)", "=TRIM(\"  a  b   c \")", "=SEARCH(\"c*t\",\"a cat sat\")", "=SEARCH(\"~?\",\"a?b\")", "=CODE(\"abc\")",
            "=WORKDAY.INTL(E1,-5,\"0000011\")", "=WORKDAY.INTL(E1,10,2)", "=NETWORKDAYS.INTL(E1,E1+14,\"0000011\")", "=NETWORKDAYS.INTL(E1,E1+14,7)",
            "=XLOOKUP(2.5,H1:H4,I1:I4,\"x\",-1)", "=XLOOKUP(2.5,H1:H4,I1:I4,\"x\",1)", "=XLOOKUP(\"g*\",B1:B8,A1:A8,0,2)", "=XLOOKUP(1,A1:A8,B1:B8,\"n\",0,-1)", "=XLOOKUP(\"BANANA\",B1:B8,A1:A8)",
            "=TEXTBEFORE(F2,\"o\",2)", "=TEXTAFTER(F2,\"o\",-1)", "=TEXTBEFORE(F2,\"zz\",1,0,0,\"none\")", "=TEXTBEFORE(F2,\"O\",1,1)", "=TEXTAFTER(F2,\"d\",1,0,1)",
            "=BASE(255,2,10)", "=BASE(7,2)", "=DECIMAL(\"1010\",2)", "=AGGREGATE(4,6,A1:A8)", "=AGGREGATE(5,6,A1:A8)", "=AGGREGATE(12,6,A1:A8)", "=AGGREGATE(15,6,A1:A8,3)", "=AGGREGATE(16,6,A1:A8,0.25)", "=AGGREGATE(17,6,A1:A8,3)",
            "=CONVERT(1,\"km\",\"mi\")", "=CONVERT(32,\"F\",\"C\")", "=CONVERT(1,\"lbm\",\"kg\")", "=CONVERT(60,\"mn\",\"hr\")", "=CONVERT(1,\"m2\",\"ft2\")", "=CONVERT(1,\"gal\",\"l\")", "=CONVERT(1,\"kbyte\",\"byte\")", "=CONVERT(1,\"m\",\"kg\")",
            "=DATEDIF(E1,E1+400,\"MD\")", "=DATEDIF(E1,E1+400,\"YM\")", "=DATEDIF(E1,E1+400,\"YD\")", "=DATEDIF(E1+10,E1,\"D\")",
            "=IMDIV(\"3+4i\",\"1+2i\")", "=IMPRODUCT(\"1+i\",\"2-3i\")", "=IMSUB(\"5+3i\",\"2+i\")", "=IMREAL(\"3-4i\")", "=IMAGINARY(\"3-4i\")", "=IMCONJUGATE(\"3-4i\")", "=IMARGUMENT(\"3+4i\")",
            "=CUMPRINC(0.05/12,60,10000,1,12,0)", "=CUMIPMT(0.05/12,60,10000,1,12,1)", "=CUMPRINC(0.05/12,60,10000,1,12,1)", "=CUMIPMT(0.05/12,60,10000,13,24,0)",
            "=DB(1000,100,5,3)", "=DB(1000,100,5,6,6)", "=DDB(1000,100,5,3,1.5)", "=SYD(1000,100,5,3)", "=IPMT(0.05/12,2,60,10000,0,1)", "=PMT(0.05/12,60,10000,0,1)", "=PPMT(0.05/12,2,60,10000,0,1)", "=PPMT(0.05/12,2,60,10000)", "=IPMT(0.05/12,1,60,10000,0,1)", "=IPMT(0.05/12,12,60,10000,2000)", "=PMT(0.05/12,60,10000,5000,0)", "=FV(0.05,10,-100,0,1)", "=FV(0.05,10,-100,-500)", "=PV(0.05,10,-100,0,1)", "=NPER(0.05/12,-200,10000,0,1)", "=RATE(60,-200,10000,0,1)",
            "=BETA.DIST(0.5,2,3,FALSE)", "=BETA.INV(0.5,2,3)", "=T.INV(0.9,10)", "=T.INV(0.1,10)", "=T.DIST.2T(2,10)", "=T.DIST.RT(2,10)", "=T.DIST(1,5,FALSE)", "=CHISQ.INV.RT(0.05,3)", "=CHISQ.DIST.RT(2,3)", "=CHISQ.DIST(2,3,FALSE)",
            "=F.INV(0.9,4,6)", "=F.INV.RT(0.05,4,6)", "=F.DIST.RT(2,4,6)", "=F.DIST(1,4,6,FALSE)", "=GAMMA.INV(0.5,3,2)", "=GAMMA.DIST(2,3,2,FALSE)", "=BINOM.INV(10,0.5,0.7)", "=POISSON.DIST(2,3,TRUE)", "=EXPON.DIST(1,2,FALSE)",
            "=LOGNORM.INV(0.5,0,1)", "=LOGNORM.DIST(1,0,1,FALSE)", "=WEIBULL.DIST(1,2,3,FALSE)", "=NEGBINOM.DIST(2,3,0.5,TRUE)", "=HYPGEOM.DIST(1,4,4,10,TRUE)", "=CONFIDENCE.T(0.05,1,10)", "=NORM.INV(0.3,5,2)", "=NORM.DIST(6,5,2,FALSE)", "=NORM.S.DIST(0.5,FALSE)",
            "=SKEW.P(A1:A8)", "=COVARIANCE.S(A1:A8,A8:A1)", "=PERCENTILE.EXC(A1:A8,0.9)", "=QUARTILE.EXC(A1:A8,3)", "=RANK.AVG(1,A1:A8,1)", "=MODE.SNGL(A1:A3)", "=MULTINOMIAL(1,2,3)", "=SERIESSUM(2,1,2,A1:A3)",
            "=ERF(0.5,1.5)", "=ERFC(0)", "=GESTEP(-1)", "=BITLSHIFT(8,-2)", "=BITRSHIFT(8,-2)", "=BITAND(5,3)", "=PERMUT(4,4)", "=COMBINA(3,3)", "=UNICHAR(65)", "=UNICODE(\"ab\")", "=ISOWEEKNUM(DATE(2021,1,3))", "=ISOWEEKNUM(DATE(2020,12,31))",
            "=VALUETOTEXT(B1,1)", "=VALUETOTEXT(5,0)", "=XOR(1,0,1)", "=XOR(A1:A3>1)", "=TYPE(\"x\"&1)", "=TYPE(1/0)", "=TYPE(A1:B2)",
            // third pass: the older names and the variants that had no check at all
            "=ACOT(2)", "=ACOTH(2)", "=COTH(1)", "=CSCH(1)", "=SECH(1)", "=ARRAYTOTEXT(A1:A3,1)",
            "=BETADIST(0.5,2,3)", "=BETAINV(0.5,2,3)", "=CEILING.PRECISE(-4.3,2)", "=FLOOR.PRECISE(-4.3,2)", "=CHIDIST(2,3)", "=CHIINV(0.3,3)",
            "=CONFIDENCE(0.05,1,100)", "=CRITBINOM(10,0.5,0.4)", "=ERF.PRECISE(1)", "=ERFC.PRECISE(1)", "=EXPONDIST(1,2,TRUE)",
            "=FDIST(1,2,3)", "=FINV(0.3,2,3)", "=GAMMADIST(2,3,2,TRUE)", "=GAMMAINV(0.5,3,2)", "=GAMMALN.PRECISE(5)", "=HYPGEOMDIST(1,4,4,10)",
            "=LOGINV(0.5,0,1)", "=LOGNORMDIST(1,0,1)", "=NEGBINOMDIST(2,3,0.5)", "=NORMINV(0.9,0,1)", "=NORMSINV(0.9)", "=POISSON(2,3,FALSE)",
            "=QUARTILE.INC(A1:A8,1)", "=WEIBULL(1,2,3,TRUE)", "=ARRAYTOTEXT(A1:B2,1)", "=ARRAYTOTEXT(B1:B2,1)", "=HYPGEOMDIST(2,4,4,10)", "=NEGBINOMDIST(0,3,0.5)", "=LOGNORMDIST(2,0.5,1.5)", "=SHEETS()", "=SHEETS(A1)", "=SHEET(A1)", "=SHEET(\"S\")", "=SHEET(\"nowhere\")", "=ISFORMULA(L1)", "=ISFORMULA(A1)", "=ISFORMULA(B9)", "=ISFORMULA(1)", "=FORMULATEXT(L1)", "=FORMULATEXT(A1)", "=FORMULATEXT(1)",
            "=XNPV(0,C1:C4,H1:H4)", "=XNPV(0.1,C1:C4,H1:H4)", "=XNPV(-0.5,C1:C4,H1:H4)", "=NEGBINOM.DIST(2,2,1,TRUE)", "=NEGBINOM.DIST(2,2,0,FALSE)", "=NEGBINOM.DIST(2,2,0.5,FALSE)", "=NEGBINOMDIST(2,2,1)", "=NEGBINOMDIST(2,2,0.999)",
            "=WEEKDAY(1)", "=WEEKDAY(0)", "=WEEKDAY(60)", "=WEEKDAY(61)", "=WEEKDAY(45000,2)", "=WEEKDAY(45000,3)", "=WEEKDAY(45000,11)", "=WEEKDAY(45000,13)", "=WEEKDAY(45000,17)", "=WEEKDAY(45000,18)", "=WEEKDAY(-1)",
            "=COUNT(TRUE)", "=COUNT(\"2\",\"x\",3)", "=COUNT(K1:K4)", "=COUNT(B1:B3)", "=COUNT(NA(),1)", "=COUNT(A1:A3,TRUE,\"7\")",
            "=AND(1,TRUE)", "=AND(1,0)", "=OR(0,0)", "=OR(\"x\",1)", "=AND(\"x\")", "=OR(\"TRUE\",FALSE)", "=AND(\"FALSE\",TRUE)", "=AND(A1:A3)", "=OR(B1:B3)", "=AND(B1:B3,TRUE)", "=OR(K2:K3)", "=OR(K1:K3)", "=AND(A9)", "=AND(A9,TRUE)", "=OR(NA(),TRUE)",
            "=IFS(TRUE,\"a\")", "=IFS(2,\"num\")", "=IFS(0,1,TRUE,2)", "=IFS(FALSE,1)", "=IFS(\"x\",1)", "=IFS(A9,1,TRUE,2)", "=IFS(NA(),1)", "=IFS(A1>3,B1,A1>0,B2)",
            "=RANK(4,A1:A8)", "=RANK(4,A1:A8,1)", "=RANK(7,A1:A8)", "=RANK(1,A1:A8,1)", "=RANK(4,K1:K4)",
            "=WORKDAY(45000,5)", "=WORKDAY(45000,-5)", "=WORKDAY(45000,5,45003)", "=WORKDAY(0,2)", "=WORKDAY(-1,2)", "=WORKDAY(TRUE,2)", "=NETWORKDAYS(45000,45010)", "=NETWORKDAYS(45010,45000)", "=NETWORKDAYS(45000,45010,45003)", "=NETWORKDAYS(0,5)",
            "=PERCENTRANK(A1:A8,4)", "=PERCENTRANK(A1:A8,4.5)", "=PERCENTRANK(A1:A8,10)", "=PERCENTRANK(A1:A8,0)", "=PERCENTRANK(A1:A8,5.5,5)", "=PERCENTRANK(A1:A8,5.5,1)", "=PERCENTRANK(A1:A8,4,0)", "=PERCENTRANK.INC(A1:A8,2)", "=PERCENTRANK.EXC(A1:A8,4)", "=PERCENTRANK.EXC(A1:A8,4.5)", "=PERCENTRANK.EXC(A1:A8,1)", "=PERCENTRANK.EXC(A1:A8,9)", "=PERCENTRANK.EXC(A1:A8,9.5)", "=PERCENTRANK(2,2)", "=PERCENTRANK(K2:K3,1.5)", "=PERCENTRANK(A1:A3,\"x\")",
            "=TYPE(K1)", "=TYPE(K2)", "=TYPE(B1)", "=TYPE(A9)", "=TYPE(NA())", "=TYPE(TRUE)",
            "=EDATE(45000,2)", "=EDATE(K1,2)", "=EDATE(2,K1)", "=EOMONTH(K4,1)", "=WEEKNUM(K1)", "=WEEKNUM(K4,2)", "=DELTA(K1)", "=PERCENTILE(K1,0.5)", "=PERCENTILE(A1:A8,K4)", "=QUOTIENT(K1,2)", "=QUOTIENT(7,K4)",
            "=VAR(A1)", "=VAR.P(A1)", "=STDEV(A1:A8)", "=STDEV.S(A1)", "=STDEV.P(A1)", "=STDEVP(\"2\")", "=VAR.S(\"2\",3)", "=VARP(K2:K3)", "=STDEV(K1:K4)", "=VAR(A1:A3,TRUE)", "=VAR.P(NA())", "=STDEV.S(A9,2)",
            "=DELTA(2)", "=DELTA(1,1)", "=DELTA(TRUE)", "=DELTA(\"2\",2)", "=DELTA(A9)", "=STANDARDIZE(TRUE,2,2)", "=STANDARDIZE(2,2,0)", "=STANDARDIZE(A9,2,2)",
            "=SUMPRODUCT(A1:A4,C1:C4)", "=SUMPRODUCT(TRUE)", "=SUMPRODUCT(A1:A4,B1:B4)", "=SUMPRODUCT(A1:A4,K1:K4)", "=SUMPRODUCT(A1:A4,C1:C3)", "=SUMPRODUCT(B1:B3)", "=SUMPRODUCT(2,\"2\")", "=SUMPRODUCT(NA(),2)",
            "=MINA(A1:A3,TRUE)", "=MAXA(B1:B3)", "=MINA(A9,2)", "=MAXA(A9,-2)", "=MAXA(A1:B3)", "=MINA(A1:B3)",
            "=WORKDAY(45000,2,\"45001\")", "=WORKDAY(45000,2,TRUE)", "=NETWORKDAYS(45000,45003,K2:K3)", "=NETWORKDAYS(45000,45003,B1:B2)", "=SHEETS(2)",
            "=LOG(8,2)", "=LOG(2,0)", "=LOG(2,1)", "=LOG(-1)", "=LOG(100)", "=LOG(2,A9)", "=FORECAST(2,2,2)", "=FORECAST(3,C1:C4,A1:A4)",
            "=ISODD(TRUE)", "=MROUND(TRUE,2)", "=QUOTIENT(TRUE,2)", "=EOMONTH(TRUE,1)", "=FACTDOUBLE(TRUE)", "=YEARFRAC(TRUE,45000)", "=COMPLEX(TRUE,2)", "=DEC2BIN(TRUE)", "=WORKDAY.INTL(TRUE,2)", "=ISEVEN(TRUE)", "=ISODD(3)", "=ISEVEN(\"4\")",
            "=SLOPE(2,2)", "=SLOPE(C1:C4,A1:A4)", "=INTERCEPT(C1:C4,A1:A4)", "=INTERCEPT(2,2)", "=SLOPE(C1:C2,A1:A1)", "=SLOPE(C1:C4,K1:K4)", "=INTERCEPT(A1:A3,A1:A3)",
            "=CEILING.MATH(-4.3,1,1)", "=FLOOR.MATH(-4.7,1,1)", "=CEILING.MATH(6.5,2)", "=FLOOR.MATH(6.5,2)",
            // fourth pass: what an argument may be (text that reads as a number, a boolean, a blank cell, an error)
            "=GCD(\"12\",18)", "=LCM(\"4\",6)", "=GCD(\"x\",1)", "=SQRTPI(\"2\")", "=SQRTPI(\"50%\")", "=SQRTPI(\"1,000\")", "=SQRTPI(\" 2 \")",
            "=SQRTPI(\"\")", "=SQRTPI(TRUE)", "=SQRTPI(A9)", "=SQRTPI(1/0)", "=SEC(\"1\")", "=ACOT(\"2\")", "=ROMAN(\"14\")", "=BITAND(\"5\",3)",
            "=NORM.S.DIST(\"1\",TRUE)", "=NORM.S.DIST(1,\"TRUE\")", "=NORM.S.DIST(1,\"x\")", "=NORM.S.DIST(1,1)", "=NORM.S.DIST(1,0)", "=NORM.S.DIST(1,A9)",
            "=BINOM.DIST(3,10,0.5,\"FALSE\")", "=BINOM.DIST(3,10,0.5,1/0)", "=BINOM.DIST(3,10,0.5,2)",
            "=LENB(5)", "=LENB(A9)", "=LENB(1/0)", "=UNICODE(1.5)", "=UNICODE(A9)", "=UNICODE(1/0)", "=TEXTBEFORE(1234,3)", "=TEXTAFTER(\"a-b-c\",\"-\",2)",
            "=GEOMEAN(A1:A3,\"4\")", "=GEOMEAN(A1:A3,B1)", "=GEOMEAN(A1:A3,1/0)", "=GEOMEAN(A1:A3,TRUE)", "=HARMEAN(A1:A3,NA())", "=HARMEAN(A1:A3,\"x\")", "=AVERAGEA(A1:A3,\"4\")",
            "=AVERAGEA(B1:B3)", "=AVERAGEA(A1:A3,TRUE)", "=AVERAGEA(A1:A3,\"x\")", "=MODE(\"2\",2)", "=MODE.SNGL(A1:A3,\"2\")", "=XOR(\"TRUE\",TRUE)", "=XOR(\"2\",2)", "=FVSCHEDULE(2,\"2\")", "=XNPV(2,\"2\",2)", "=CEILING.MATH(2,2,\"2\")", "=FLOOR.MATH(2,2,\"1\")", "=POISSON(2,2,\"2\")", "=ERF(TRUE)", "=EFFECT(2,TRUE)", "=NETWORKDAYS.INTL(TRUE,2)", "=GESTEP(TRUE)", "=IFERROR(GCD(2,A9),-1)", "=TRIMMEAN(A1:A8,1)", "=TRIMMEAN(A1:A8,-0.1)", "=CORREL(A1:A4,B1:B4)", "=CORREL(A1:A4,C1:C3)",
        };
    }
}

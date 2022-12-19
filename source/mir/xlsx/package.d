module mir.xlsx;

import std.algorithm.iteration : filter, map, joiner;
import std.algorithm.mutation : reverse;
import std.algorithm.searching : all, canFind, startsWith;
import std.algorithm.sorting : sort;
import std.array : array, empty, front, replace, popFront;
import std.ascii : isDigit;
import std.conv : to;
import std.datetime : DateTime, Date, TimeOfDay;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.format : format;
import std.range : tee;
import std.regex;
import std.stdio;
import std.traits : isIntegral, isFloatingPoint, isSomeString;
import std.typecons : tuple, Nullable, nullable;
import std.utf : byChar;
import std.variant : Algebraic, visit;
import std.zip;

import dxml.dom;

///
struct Pos {
	// zero based
	size_t row;
	// zero based
	size_t col;
}

///
alias Data = Algebraic!(bool,long,double,string,DateTime,Date,TimeOfDay);

///
struct Cell {
	string loc;
	size_t row; // row[r]
	string t; // s or n, s for pointer, n for value, stored in v
	string r; // c[r]
	string v; // c.v the value or ptr
	string f; // c.f the formula
	string xmlValue;
	Pos position;
}

//
enum CellType {
	datetime,
	timeofday,
	date,
	bool_,
	double_,
	long_,
	string_
}

import std.ascii : toUpper;
///
struct Sheet {
	import std.ascii : toUpper;

	Cell[] cells;
	Cell[][] table;
	Pos maxPos;

	string toString() const @safe {
		import std.format : formattedWrite;
		import std.array : appender;
		long[] maxCol = new long[](maxPos.col + 1);
		foreach(const row; this.table) {
			foreach(const idx, Cell col; row) {
				string s = col.xmlValue;

				maxCol[idx] = maxCol[idx] < s.length ? s.length : maxCol[idx];
			}
		}
		maxCol[] += 1;

		auto app = appender!string();
		foreach(const row; this.table) {
			foreach(const idx, Cell col; row) {
				string s = col.xmlValue;
				formattedWrite(app, "%*s, ", maxCol[idx], s);
			}
			formattedWrite(app, "\n");
		}
		return app.data;
	}

	void printTable() const @safe {
		writeln(this.toString());
	}

	// Column

	Iterator!T getColumn(T)(size_t col, size_t startRow, size_t endRow) {
		auto c = this.iterateColumn!T(col, startRow, endRow);
		return typeof(return)(c.array);
	}

	private enum t = q{
	Iterator!(%1$s) getColumn%2$s(size_t col, size_t startRow, size_t endRow) @safe {
		return getColumn!(%1$s)(col, startRow, endRow);
	}
	};
	static foreach(T; ["long", "double", "string", "Date", "TimeOfDay",
			"DateTime"])
	{
		mixin(format(t, T, T[0].toUpper ~ T[1 .. $]));
	}

	ColumnUntyped iterateColumnUntyped(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(T) iterateColumn(T)(size_t col, size_t startRow, size_t endRow) {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(long) iterateColumnLong(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(double) iterateColumnDouble(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(string) iterateColumnString(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(DateTime) iterateColumnDateTime(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(Date) iterateColumnDate(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(TimeOfDay) iterateColumnTimeOfDay(size_t col, size_t startRow, size_t endRow) @safe {
		return typeof(return)(&this, col, startRow, endRow);
	}

	// Row

	Iterator!T getRow(T)(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(this.iterateRow!T(row, startColumn, endColumn).array); // TODO: why .array?
	}

	private enum t2 = q{
	Iterator!(%1$s) getRow%2$s(size_t row, size_t startColumn, size_t endColumn) @safe {
		return getRow!(%1$s)(row, startColumn, endColumn);
	}
	};
	static foreach(T; ["long", "double", "string", "Date", "TimeOfDay",
			"DateTime"])
	{
		mixin(format(t2, T, T[0].toUpper ~ T[1 .. $]));
	}

	RowUntyped iterateRowUntyped(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(T) iterateRow(T)(size_t row, size_t startColumn, size_t endColumn) @trusted /* TODO: remove @trusted when `&this` is stored in a @safe manner in `Row` */ {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(long) iterateRowLong(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(double) iterateRowDouble(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(string) iterateRowString(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(DateTime) iterateRowDateTime(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(Date) iterateRowDate(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(TimeOfDay) iterateRowTimeOfDay(size_t row, size_t startColumn, size_t endColumn) @safe {
		return typeof(return)(&this, row, startColumn, endColumn);
	}
}

struct Iterator(T) {
	T[] data;

	this(T[] data) {
		this.data = data;
	}

	@property bool empty() const pure nothrow @nogc {
		return this.data.empty;
	}

	void popFront() {
		this.data.popFront();
	}

	@property T front() {
		return this.data.front;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	// Request random access.
	inout(T)[] array() inout @safe pure nothrow @nogc {
		return data;
	}
}

///
struct RowUntyped {
	Sheet* sheet;
	const size_t row;
	const size_t startColumn;
	const size_t endColumn;
	size_t cur;

	this(Sheet* sheet, size_t row, size_t startColumn, size_t endColumn) pure nothrow @nogc @safe {
		assert(sheet.table.length == sheet.maxPos.row + 1);
		this.sheet = sheet;
		this.row = row;
		this.startColumn = startColumn;
		this.endColumn = endColumn;
		this.cur = this.startColumn;
	}

	@property bool empty() const pure nothrow @nogc @safe {
		return this.cur >= this.endColumn;
	}

	void popFront() pure nothrow @nogc @safe {
		++this.cur;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc @safe {
		return this;
	}

	@property inout(Cell) front() inout pure nothrow @nogc @safe {
		return this.sheet.table[this.row][this.cur];
	}
}

///
struct Row(T) {
	RowUntyped ru;
	T front;

	this(Sheet* sheet, size_t row, size_t startColumn, size_t endColumn) {
		this.ru = RowUntyped(sheet, row, startColumn, endColumn);
		this.read();
	}

	@property bool empty() const pure nothrow @nogc {
		return this.ru.empty;
	}

	void popFront() {
		this.ru.popFront();
		if(!this.empty) {
			this.read();
		}
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	private void read() {
		this.front = convertTo!T(this.ru.front.xmlValue);
	}
}

///
struct ColumnUntyped {
	Sheet* sheet;
	const size_t col;
	const size_t startRow;
	const size_t endRow;
	size_t cur;

	this(Sheet* sheet, size_t col, size_t startRow, size_t endRow) @safe {
		assert(sheet.table.length == sheet.maxPos.row + 1);
		this.sheet = sheet;
		this.col = col;
		this.startRow = startRow;
		this.endRow = endRow;
		this.cur = this.startRow;
	}

	@property bool empty() const pure nothrow @nogc @safe {
		return this.cur >= this.endRow;
	}

	void popFront() @safe {
		++this.cur;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc @safe {
		return this;
	}

	@property Cell front() @safe {
		return this.sheet.table[this.cur][this.col];
	}
}

///
struct Column(T) {
	ColumnUntyped cu;

	T front;

	this(Sheet* sheet, size_t col, size_t startRow, size_t endRow) {
		this.cu = ColumnUntyped(sheet, col, startRow, endRow);
		this.read();
	}

	@property bool empty() const pure nothrow @nogc {
		return this.cu.empty;
	}

	void popFront() {
		this.cu.popFront();
		if(!this.empty) {
			this.read();
		}
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	private void read() {
		this.front = convertTo!T(this.cu.front.xmlValue);
	}
}

@safe unittest {
	import std.range : isForwardRange;
	import std.meta : AliasSeq;
	static foreach(T; AliasSeq!(long,double,DateTime,TimeOfDay,Date,string)) {{
		alias C = Column!T;
		alias R = Row!T;
		alias I = Iterator!T;
		static assert(isForwardRange!C, C.stringof);
		static assert(isForwardRange!R, R.stringof);
		static assert(isForwardRange!I, I.stringof);
	}}
}

Date longToDate(long d) @safe {
	// modifed from https://www.codeproject.com/Articles/2750/
	// Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

	// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
	// leap year, but Excel/Lotus 123 think it is...
	if(d == 60) {
		return Date(1900, 2,  29);
	} else if(d < 60) {
		// Because of the 29-02-1900 bug, any serial date
		// under 60 is one off... Compensate.
		++d;
	}

	// Modified Julian to DMY calculation with an addition of 2415019
	int l = cast(int)d + 68569 + 2415019;
	int n = int(( 4 * l ) / 146097);
	l = l - int(( 146097 * n + 3 ) / 4);
	int i = int(( 4000 * ( l + 1 ) ) / 1461001);
	l = l - int(( 1461 * i ) / 4) + 31;
	int j = int(( 80 * l ) / 2447);
	int nDay = l - int(( 2447 * j ) / 80);
	l = int(j / 11);
	int nMonth = j + 2 - ( 12 * l );
	int nYear = 100 * ( n - 49 ) + i + l;
	return Date(nYear, nMonth, nDay);
}

long dateToLong(Date d) @safe {
	// modifed from https://www.codeproject.com/Articles/2750/
	// Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

	// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
	// leap year, but Excel/Lotus 123 think it is...
	if(d.day == 29 && d.month == 2 && d.year == 1900) {
		return 60;
	}

	// DMY to Modified Julian calculated with an extra subtraction of 2415019.
	long nSerialDate =
			int(( 1461 * ( d.year + 4800 + int(( d.month - 14 ) / 12) ) ) / 4) +
			int(( 367 * ( d.month - 2 - 12 *
				( ( d.month - 14 ) / 12 ) ) ) / 12) -
				int(( 3 * ( int(( d.year + 4900
				+ int(( d.month - 14 ) / 12) ) / 100) ) ) / 4) +
				d.day - 2415019 - 32075;

	if(nSerialDate < 60) {
		// Because of the 29-02-1900 bug, any serial date
		// under 60 is one off... Compensate.
		nSerialDate--;
	}

	return nSerialDate;
}

@safe unittest {
	auto ds = [ Date(1900,2,1), Date(1901, 2, 28), Date(2019, 06, 05) ];
	foreach(const d; ds) {
		long l = dateToLong(d);
		Date r = longToDate(l);
		assert(r == d, format("%s %s", r, d));
	}
}

TimeOfDay doubleToTimeOfDay(double s) @safe {
	import core.stdc.math : lround;
	double secs = (24.0 * 60.0 * 60.0) * s;

	// TODO not one-hundred my lround is needed
	int secI = to!int(lround(secs));

	return TimeOfDay(secI / 3600, (secI / 60) % 60, secI % 60);
}

double timeOfDayToDouble(TimeOfDay tod) @safe {
	long h = tod.hour * 60 * 60;
	long m = tod.minute * 60;
	long s = tod.second;
	return (h + m + s) / (24.0 * 60.0 * 60.0);
}

@safe unittest {
	auto tods = [ TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11),
		 TimeOfDay(0, 0, 0), TimeOfDay(0, 1, 0),
		 TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
	foreach(const tod; tods) {
		double d = timeOfDayToDouble(tod);
		assert(d <= 1.0, format("%s", d));
		TimeOfDay r = doubleToTimeOfDay(d);
		assert(r == tod, format("%s %s", r, tod));
	}
}

double datetimeToDouble(DateTime dt) @safe {
	double d = dateToLong(dt.date);
	double t = timeOfDayToDouble(dt.timeOfDay);
	return d + t;
}

DateTime doubleToDateTime(double d) @safe {
	long l = cast(long)d;
	Date dt = longToDate(l);
	TimeOfDay t = doubleToTimeOfDay(d - l);
	return DateTime(dt, t);
}

@safe unittest {
	auto ds = [ Date(1900,2,1), Date(1901, 2, 28), Date(2019, 06, 05) ];
	auto tods = [ TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11),
		 TimeOfDay(0, 0, 0), TimeOfDay(0, 1, 0),
		 TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
	foreach(const d; ds) {
		foreach(const tod; tods) {
			DateTime dt = DateTime(d, tod);
			double dou = datetimeToDouble(dt);

			Date rd = longToDate(cast(long)dou);
			assert(rd == d, format("%s %s", rd, d));

			double rest = dou - cast(long)dou;
			TimeOfDay rt = doubleToTimeOfDay(dou - cast(long)dou);
			assert(rt == tod, format("%s %s", rt, tod));

			DateTime r = doubleToDateTime(dou);
			assert(r == dt, format("%s %s", r, dt));
		}
	}
}

Date stringToDate(string s) @safe {
	import std.array : split;
	import std.string : indexOf;

	if(s.indexOf('/') != -1) {
		auto sp = s.split('/');
		enforce(sp.length == 3, format("[%s]", sp));
		return Date(to!int(sp[2]), to!int(sp[1]), to!int(sp[0]));
	} else {
		return longToDate(to!long(s));
	}
}

bool tryConvertTo(T,S)(S var) {
	return !(tryConvertToImpl!T(Data(var)).isNull());
}

Nullable!(T) tryConvertToImpl(T)(Data var) {
	try {
		return nullable(convertTo!T(var));
	} catch(Exception e) {
		return Nullable!T();
	}
}

T convertTo(T)(string var) @safe {
	import std.math : lround;
	static if(isSomeString!T) {
		return to!T(var);
	} else static if(is(T == bool)) {
		return var == "1";
	} else static if(isIntegral!T) {
		return to!T(var);
	} else static if(isFloatingPoint!T) {
		return to!T(var);
	} else static if(is(T == DateTime)) {
		if(var.canConvertToLong()) {
			return doubleToDateTime(to!long(var));
		} else if(var.canConvertToDouble()) {
			return doubleToDateTime(to!double(var));
		}
		enforce(false, "Can not convert '" ~ var ~ "' to a DateTime");
		assert(false, "Unreachable");
	} else static if(is(T == Date)) {
		if(var.canConvertToLong()) {
			return longToDate(to!long(var));
		} else if(var.canConvertToDouble()) {
			return longToDate(lround(to!double(var)));
		}
		return stringToDate(var);
	} else static if(is(T == TimeOfDay)) {
		double l = to!double(var);
		return doubleToTimeOfDay(l - cast(long)l);
	} else {
		static assert(false, T.stringof ~ " not supported");
	}
}

private ZipArchive readFile(in string filename) @trusted {
	enforce(exists(filename), "File with name " ~ filename ~ " does not exist");
	return new typeof(return)(read(filename));
}

struct SheetNameId {
	string name;
	int id;
	string rid;
}

string convertToString(in ubyte[] d) @trusted {
	import std.encoding;
	auto b = getBOM(d);
	switch(b.schema) {
		case BOM.none:
			return cast(string)d;
		case BOM.utf8:
			return cast(string)(d[3 .. $]);
		case BOM.utf16be: goto default;
		case BOM.utf16le: goto default;
		case BOM.utf32be: goto default;
		case BOM.utf32le: goto default;
		default:
			string ret;
			transcode(d, ret);
			return ret;
	}
}

SheetNameId[] sheetNames(in string filename) @trusted {
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable wbStr = "xl/workbook.xml";
	if(wbStr !in ams) {
		return SheetNameId[].init;
	}
	ubyte[] wb = file.expand(ams[wbStr]);
	string wbData = convertToString(wb);

	auto dom = parseDOM(wbData);
	if(dom.children.length != 1) {
		return [];
	}
	auto workbook = dom.children[0];
	string sheetName = workbook.name == "workbook"
		? "sheets" : "s:sheets";
	if(workbook.name != "workbook" && workbook.name != "s:workbook") {
		return [];
	}
	auto sheetsRng = workbook.children.filter!(c => c.name == sheetName);
	if(sheetsRng.empty) {
		return [];
	}

	return sheetsRng.front.children
		.map!(s => SheetNameId(
					s.attributes.filter!(a => a.name == "name").front.value
						.specialCharacterReplacementReverse(),
					s.attributes.filter!(a => a.name == "sheetId").front
						.value.to!int(),
					s.attributes.filter!(a => a.name == "r:id").front.value,
				)
		)
		.array
		.sort!((a, b) => a.id < b.id)
		.release;
}

@safe unittest {
	auto r = sheetNames("multitable.xlsx");
	assert(r[0].name == "wb1");
	assert(r[0].id == 1);
}

@safe unittest {
	auto r = sheetNames("sheetnames.xlsx");
	assert(r[0].name == "A & B ;", r[0].name);
	assert(r[0].id == 1);
}

struct Relationships {
	string id;
	string file;
}

Relationships[string] parseRelationships(ZipArchive za, ArchiveMember am) @trusted {
	ubyte[] d = za.expand(am);
	string relData = convertToString(d);
	auto dom = parseDOM(relData);
	assert(dom.children.length == 1);
	auto rel = dom.children[0];
	assert(rel.name == "Relationships");
	auto relRng = rel.children.filter!(c => c.name == "Relationship");
	assert(!relRng.empty);

	Relationships[string] ret;
	foreach(r; relRng) {
		Relationships tmp;
		tmp.id = r.attributes.filter!(a => a.name == "Id")
			.front.value;
		tmp.file = r.attributes.filter!(a => a.name == "Target")
			.front.value;
		ret[tmp.id] = tmp;
	}
	return ret;
}

Sheet readSheet(in string filename, in string sheetName) @safe {
	SheetNameId[] sheets = sheetNames(filename);
	auto sRng = sheets.filter!(s => s.name == sheetName);
	enforce(!sRng.empty, "No sheet with name " ~ sheetName
			~ " found in file " ~ filename);
	return readSheetImpl(filename, sRng.front.rid);
}

string eatXlPrefix(string fn) @safe {
	foreach(const p; ["xl//", "/xl/"]) {
		if(fn.startsWith(p)) {
			return fn[p.length .. $];
		}
	}
	return fn;
}

Sheet readSheetImpl(in string filename, in string rid) @trusted {
	scope(failure) {
		writefln("Failed at file '%s' and sheet '%s'", filename, rid);
	}
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable ss = "xl/sharedStrings.xml";
	string[] sharedStrings = (ss in ams)
		? readSharedEntries(file, ams[ss])
		: [];
	//logf("%s", sharedStrings);

	Relationships[string] rels = parseRelationships(file,
			ams["xl/_rels/workbook.xml.rels"]);

	Relationships* sheetRel = rid in rels;
	enforce(sheetRel !is null, format("Could not find '%s' in '%s'", rid,
				filename));
	string shrFn = eatXlPrefix(sheetRel.file);
	string fn = "xl/" ~ shrFn;
	ArchiveMember* sheet = fn in ams;
	enforce(sheet !is null, format("sheetRel.file orig '%s', fn %s not in [%s]",
				sheetRel.file, fn, ams.keys()));

	Sheet ret;
	ret.cells = insertValueIntoCell(readCells(file, *sheet), sharedStrings);
	Pos maxPos;
	foreach(ref c; ret.cells) {
		c.position = toPos(c.r);
		maxPos = elementMax(maxPos, c.position);
	}
	ret.maxPos = maxPos;
	ret.table = new Cell[][](ret.maxPos.row + 1, ret.maxPos.col + 1);
	foreach(const c; ret.cells) {
		ret.table[c.position.row][c.position.col] = c;
	}
	return ret;
}

string[] readSharedEntries(ZipArchive za, ArchiveMember am) @trusted {
	ubyte[] ss = za.expand(am);
	string ssData = convertToString(ss);
	auto dom = parseDOM(ssData);
	string[] ret;
	if(dom.type != EntityType.elementStart) {
		return ret;
	}
	assert(dom.children.length == 1);
	auto sst = dom.children[0];
	assert(sst.name == "sst");
	if(sst.type != EntityType.elementStart || sst.children.empty) {
		return ret;
	}
	auto siRng = sst.children.filter!(c => c.name == "si");
	foreach(si; siRng) {
		if(si.type != EntityType.elementStart) {
			continue;
		}
		//ret ~= extractData(si);
		string tmp;
		foreach(tORr; si.children) {
			if(tORr.name == "t" && tORr.type == EntityType.elementStart
					&& !tORr.children.empty)
			{
				//ret ~= Data(convert(tORr.children[0].text));
				ret ~= tORr.children[0].text.removeSpecialCharacter();
			} else if(tORr.name == "r") {
				foreach(r; tORr.children.filter!(r => r.name == "t")) {
					if(r.type == EntityType.elementStart && !r.children.empty) {
						tmp ~= r.children[0].text.removeSpecialCharacter();
					}
				}
			} else {
				//ret ~= Data.init;
				ret ~= "";
			}
		}
		if(!tmp.empty) {
			//ret ~= Data(convert(tmp));
			ret ~= tmp.removeSpecialCharacter();
		}
	}
	return ret;
}

string extractData(DOMEntity!string si) {
	string tmp;
	foreach(tORr; si.children) {
		if(tORr.name == "t") {
			if(!tORr.attributes.filter!(a => a.name == "xml:space").empty) {
				return "";
			} else if(tORr.type == EntityType.elementStart
					&& !tORr.children.empty)
			{
				return tORr.children[0].text;
			} else {
				return "";
			}
		} else if(tORr.name == "r") {
			foreach(r; tORr.children.filter!(r => r.name == "t")) {
				tmp ~= r.children[0].text;
			}
		}
	}
	if(!tmp.empty) {
		return tmp;
	}
	assert(false);
}

private bool canConvertToLong(in string s) @safe {
	if(s.empty) {
		return false;
	}
	return s.byChar.all!isDigit();
}

private immutable rs = r"[\+-]{0,1}[0-9][0-9]*\.[0-9]*";
private auto rgx = ctRegex!rs;

private bool canConvertToDoubleOld(in string s) @safe {
	auto cap = matchAll(s, rgx);
	return cap.empty || cap.front.hit != s ? false : true;
}

private bool canConvertToDouble(string s) pure @safe @nogc {
	if(s.startsWith('+') || s.startsWith('-')) {
		s = s[1 .. $];
	}
	if(s.empty) {
		return false;
	}

	if(s[0] < '0' || s[0] > '9') { // at least one in [0-9]
		return false;
	}

	s = s[1 .. $];

	if(s.empty) {
		return true;
	}

	while(!s.empty && s[0] >= '0' && s[0] <= '9') {
		s = s[1 .. $];
	}
	if(s.empty) {
		return true;
	}
	if(s[0] != '.') {
		return false;
	}
	s = s[1 .. $];
	if(s.empty) {
		return true;
	}

	while(!s.empty && s[0] >= '0' && s[0] <= '9') {
		s = s[1 .. $];
	}

	return s.empty;
}

@safe unittest {
	static struct Test {
		string tt;
		bool rslt;
	}
	auto tests =
		[ Test("-", false)
		, Test("0.0", true)
		, Test("-0.", true)
		, Test("-0.0", true)
		, Test("-0.a", false)
		, Test("-0.0", true)
		, Test("-1100.0", true)
		];
	foreach(const t; tests) {
		assert(canConvertToDouble(t.tt) == canConvertToDoubleOld(t.tt)
				&& canConvertToDouble(t.tt) == t.rslt
			, format("%s %s %s %s", t.tt
				, canConvertToDouble(t.tt), canConvertToDoubleOld(t.tt)
				, t.rslt));
	}
}

string removeSpecialCharacter(string s) {
	struct ToRe {
		string from;
		string to;
	}

	immutable ToRe[] toRe = [
		ToRe( "&amp;", "&"),
		ToRe( "&gt;", "<"),
		ToRe( "&lt;", ">"),
		ToRe( "&quot;", "\""),
		ToRe( "&apos;", "'")
	];

	string replaceStrings(string s) {
		import std.algorithm.searching : canFind;
		import std.array : replace;
		foreach(const tr; toRe) {
			while(canFind(s, tr.from)) {
				s = s.replace(tr.from, tr.to);
			}
		}
		return s;
	}

	return replaceStrings(s);
}

Cell[] readCells(ZipArchive za, ArchiveMember am) @trusted {
	Cell[] ret;
	ubyte[] ss = za.expand(am);
	string ssData = convertToString(ss);
	auto dom = parseDOM(ssData);
	assert(dom.children.length == 1);
	auto ws = dom.children[0];
	if(ws.name != "worksheet") {
		return ret;
	}
	auto sdRng = ws.children.filter!(c => c.name == "sheetData");
	assert(!sdRng.empty);
	if(sdRng.front.type != EntityType.elementStart) {
		return ret;
	}
	auto rows = sdRng.front.children
		.filter!(r => r.name == "row");

	foreach(ref row; rows) {
		if(row.type != EntityType.elementStart || row.children.empty) {
			continue;
		}
		foreach(ref c; row.children.filter!(r => r.name == "c")) {
			Cell tmp;
			tmp.row = row.attributes.filter!(a => a.name == "r")
				.front.value.to!size_t();
			tmp.r = c.attributes.filter!(a => a.name == "r")
				.front.value;
			auto t = c.attributes.filter!(a => a.name == "t");
			if(t.empty) {
				// we assume that no t attribute means direct number
				//writefln("Found a strange empty cell \n%s", c);
			} else {
				tmp.t = t.front.value;
			}
			if(tmp.t == "s" || tmp.t == "n") {
				if(c.type == EntityType.elementStart) {
					auto v = c.children.filter!(c => c.name == "v");
					//enforce(!v.empty, format("r %s", tmp.row));
					if(!v.empty && v.front.type == EntityType.elementStart
							&& !v.front.children.empty)
					{
						tmp.v = v.front.children[0].text;
					} else {
						tmp.v = "";
					}
				}
			} else if(tmp.t == "inlineStr") {
				auto is_ = c.children.filter!(c => c.name == "is");
				tmp.v = extractData(is_.front);
			} else if(c.type == EntityType.elementStart) {
				auto v = c.children.filter!(c => c.name == "v");
				if(!v.empty && v.front.type == EntityType.elementStart
						&& !v.front.children.empty)
				{
					tmp.v = v.front.children[0].text;
				}
			}
			if(c.type == EntityType.elementStart) {
				auto f = c.children.filter!(c => c.name == "f");
				if(!f.empty && f.front.type == EntityType.elementStart) {
					tmp.f = f.front.children[0].text;
				}
			}
			ret ~= tmp;
		}
	}
	return ret;
}

Cell[] insertValueIntoCell(Cell[] cells, string[] ss) @trusted {
	immutable excepted = ["n", "s", "b", "e", "str", "inlineStr"];
	immutable same = ["n", "e", "str", "inlineStr"];
	foreach(ref Cell c; cells) {
		assert(canFind(excepted, c.t) || c.t.empty,
				format("'%s' not in [%s]", c.t, excepted));
		if(c.t.empty) {
			//c.xmlValue = convert(c.v);
			c.xmlValue = c.v.removeSpecialCharacter();
		} else if(canFind(same, c.t)) {
			//c.xmlValue = convert(c.v);
			c.xmlValue = c.v.removeSpecialCharacter();
		} else if(c.t == "b") {
			//logf("'%s' %s", c.v, c);
			//c.xmlValue = c.v == "1";
			c.xmlValue = c.v.removeSpecialCharacter();
		} else {
			if(!c.v.empty) {
				size_t idx = to!size_t(c.v);
				//logf("'%s' %s", c.v, idx);
				//c.xmlValue = ss[idx];
				c.xmlValue = ss[idx];
			}
		}
	}
	return cells;
}

Pos toPos(in string s) @safe {
	import std.string : indexOfAny;
	import std.math : pow;
	ptrdiff_t fn = s.indexOfAny("0123456789");
	enforce(fn != -1, s);
	size_t row = to!size_t(to!long(s[fn .. $]) - 1);
	size_t col = 0;
	string colS = s[0 .. fn];
	foreach(const idx, char c; colS) {
		col = col * 26 + (c - 'A' + 1);
	}
	return Pos(row, col - 1);
}

@safe unittest {
	assert(toPos("A1").col == 0);
	assert(toPos("Z1").col == 25);
	assert(toPos("AA1").col == 26);
}

Pos elementMax(Pos a, Pos b) {
	return Pos(a.row < b.row ? b.row : a.row,
			a.col < b.col ? b.col : a.col);
}

string specialCharacterReplacement(string s) {
	return s.replace("\"", "&quot;")
		.replace("'", "&apos;")
		.replace("<", "&lt;")
		.replace(">", "&gt;")
		.replace("&", "&amp;");
}

string specialCharacterReplacementReverse(string s) {
	return s.replace("&quot;", "\"")
		.replace("&apos;", "'")
		.replace("&lt;", "<")
		.replace("&gt;", ">")
		.replace("&amp;", "&");
}

@safe unittest {
	import std.math : isClose;
	auto r = readSheet("multitable.xlsx", "wb1");
	assert(isClose(r.table[12][5].xmlValue.to!double(), 26.74),
			format("%s", r.table[12][5])
		);

	assert(isClose(r.table[13][5].xmlValue.to!double(), -26.74),
			format("%s", r.table[13][5])
		);
}

@safe unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "wb1");
	auto r = s.iterateRow!long(15, 1, 6);

	auto expected = [1, 2, 3, 4, 5];
	assert(equal(r, expected), format("%s", r));

	auto r2 = s.getRow!long(15, 1, 6);
	assert(equal(r, expected));

	auto it = s.iterateRowLong(15, 1, 6);
	assert(equal(r2, it));

	auto it2 = s.iterateRowUntyped(15, 1, 6)
		.map!(it => format("%s", it))
		.array;
}

@safe unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "wb2");
	//writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
	auto rslt = s.iterateColumn!Date(1, 1, 6);
	auto rsltUt = s.iterateColumnUntyped(1, 1, 6)
		.map!(it => format("%s", it))
		.array;
	assert(!rsltUt.empty);

	auto target = [Date(2019,5,01), Date(2016,12,27), Date(1976,7,23),
		 Date(1986,7,2), Date(2038,1,19)
	];
	assert(equal(rslt, target), format("\n%s\n%s", rslt, target));

	auto it = s.getColumn!Date(1, 1, 6);
	assert(equal(rslt, it));

	auto it2 = s.getColumnDate(1, 1, 6);
	assert(equal(rslt, it2));
}

@safe unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "Sheet3");
	writeln(s.table[0][0].xmlValue);
	assert(s.table[0][0].xmlValue.to!long(),
			format("%s", s.table[0][0].xmlValue));
	//assert(s.table[0][0].canConvertTo(CellType.bool_));
}

unittest {
	import std.file : dirEntries, SpanMode;
	import std.traits : EnumMembers;
	foreach(const de; dirEntries("xlsx_files/", "*.xlsx", SpanMode.depth)
			.filter!(a => a.name != "xlsx_files/data03.xlsx"))
	{
		//writeln(de.name);
		auto sn = sheetNames(de.name);
		foreach(const s; sn) {
			auto sheet = readSheet(de.name, s.name);
			foreach(const cell; sheet.cells) {
			}
		}
	}
}

@safe unittest {
	import std.algorithm.comparison : equal;
	auto sheet = readSheet("testworkbook.xlsx", "ws1");
	//writefln("%(%s\n%)", sheet.cells);
	//writeln(sheet.toString());
	assert(sheet.table[2][3].xmlValue.to!long() == 1337);

	auto c = sheet.getColumnLong(3, 2, 5);
	auto r = [1337, 2, 3];
	assert(equal(c, r), format("%s %s", c, sheet.toString()));

	auto c2 = sheet.getColumnString(4, 2, 5);
	string f2 = sheet.table[2][4].xmlValue;
	assert(f2 == "hello", f2);
	f2 = sheet.table[3][4].xmlValue;
	assert(f2 == "sil", f2);
	f2 = sheet.table[4][4].xmlValue;
	assert(f2 == "foo", f2);
	auto r2 = ["hello", "sil", "foo"];
	assert(equal(c2, r2), format("%s", c2));
}

@safe unittest {
	import std.math : isClose;
	auto sheet = readSheet("toto.xlsx", "Trades");
	writefln("%(%s\n%)", sheet.cells);

	auto r = sheet.getRowString(1, 0, 2).array;

	double d = to!double(r[1]);
	assert(isClose(d, 38204642.510000));
}

@safe unittest {
	auto sheet = readSheet("leading_zeros.xlsx", "Sheet1");
	auto a2 = sheet.cells.filter!(c => c.r == "A2");
	assert(!a2.empty);
	assert(a2.front.xmlValue == "0012", format("%s", a2.front));
}

@safe unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("datetimes.xlsx", "Sheet1");
	//writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
	auto rslt = s.iterateColumn!DateTime(0, 0, 2);
	assert(!rslt.empty);

	auto target =
		[ DateTime(Date(1986,1,11), TimeOfDay.init)
		, DateTime(Date(1986,7,2), TimeOfDay.init)
		];
	assert(equal(rslt, target), format("\ngot: %s\nexp: %s\ntable %s", rslt
				, target, s.toString()));
}

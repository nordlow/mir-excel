module mir.xlsx;

import core.time : Duration;
import std.algorithm.iteration : filter, map;
import std.algorithm.searching : all, canFind, startsWith;
import std.algorithm.sorting : sort;
import std.array : array, empty, front, popFront, Appender;
import std.conv : to;
import std.datetime : DateTime, Date, TimeOfDay;
import std.datetime.stopwatch : StopWatch, AutoStart;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.format : format;
import std.traits : isIntegral, isFloatingPoint, isSomeString;
import std.typecons : Nullable, nullable;
import std.zip;
import mir.algebraic : Algebraic, Variant, TypeSet;
import mir.reflection : reflectIgnore, ReflectDoc;
import mir.ion.value : IonNull;
import dxml.dom : DOMEntity, EntityType, parseDOM;
import dxml.util : decodeXML; // TODO: Replace with parseXML
import dxml.parser : parseXML; // TODO: Use instead of decodeXML

debug import std.stdio;

alias SILignore = reflectIgnore!"SIL";
alias SILdoc = ReflectDoc!"SIL";

// disabled for now for faster builds
// version = ctRegex_test;

version(mir_profileGC)
    enum runCount = 1;
else version(mir_benchmark)
    enum runCount = 10;

version(mir_benchmark)
    enum tme = true; // time me
else
    enum tme = false; // don’t time

/** Row or column offset precision.
 *
 * Excel has a limit of 1_048_576 rows and 16_384 columns per sheet so use
 * 32-bit precision now.
 *
 * Uses alias instead of sub-type for backwards compatibility in for instance ’Pos’.
 */
alias Offset = uint;

/** Row width or column height precision.
 *
 * Uses alias instead of sub-type for backwards compatibility in for instance ’Pos’.
 */
alias Length = uint;

/** Row offset, starting at 0 for top-most row.
 *
 * Defaults to uint.max to indicate not yet initialized (undefined).
 *
 * Uses sub-type instead of alias for type-safe API requiring explicit conversions.
 */
struct RowOffset {
    enum undefined =
        1_048_576; ///< Excel has a limit of 1_048_576 rows so make the 1_048_576+1:th represent undefinedness.
    enum max = undefined - 1;
    Offset value = undefined;
    alias value this;
    this(Offset value) @safe pure nothrow @nogc {
        this.value = value;
    }

    this(size_t value) @safe pure {
        this.value = value.to!Offset;
    }

    bool isDefined() const @safe pure nothrow @nogc {
        return value != undefined;
    }

    invariant(value <= undefined);
}

/** Column offset, starting at 0 for left-most column.
 *
 * Defaults to uint.max to indicate not yet initialized (undefined).
 *
 * Uses sub-type instead of alias for type-safe API requiring explicit conversions.
 */
struct ColumnOffset {
    enum undefined =
        16_384; ///< Excel has a limit of 16_384 rows so make the 16_384+1:th represent undefinedness.
    enum max = undefined - 1;
    Offset value = undefined;
    alias value this;
    this(Offset value) @safe pure nothrow @nogc {
        this.value = value;
    }

    this(size_t value) @safe pure {
        this.value = value.to!Offset;
    }

    bool isDefined() const @safe pure nothrow @nogc {
        return value != undefined;
    }

    invariant(value <= undefined);
}

alias ColOffset = ColumnOffset;

/** Row width in number of rows.
 */
struct RowWidth {
    enum undefined = RowOffset.undefined + 1;
    enum max = undefined - 1;
    Length value = undefined;
    alias value this;
    this(Offset value) @safe pure nothrow @nogc {
        this.value = value;
    }

    this(size_t value) @safe pure {
        this.value = value.to!Length;
    }

    bool isDefined() const @safe pure nothrow @nogc {
        return value != undefined;
    }

    invariant(value <= undefined);
}

/** Column height in number of columns.
 */
struct ColumnHeight {
    enum undefined = ColumnOffset.undefined + 1;
    enum max = undefined - 1;
    Length value = undefined;
    alias value this;
    this(Offset value) @safe pure nothrow @nogc {
        this.value = value;
    }

    this(size_t value) @safe pure {
        this.value = value.to!Length;
    }

    bool isDefined() const @safe pure nothrow @nogc {
        return value != undefined;
    }

    invariant(value <= undefined);
}

alias ColHeight = ColumnHeight;

/** Cell position offset, starting at [0,0] for top-left cell.
	See: https://support.socrata.com/hc/en-us/articles/115005306167-Limitations-of-Excel-and-CSV-Downloads
 */
struct Position {
    RowOffset row;
    ColumnOffset column;

    alias col = column;
    /// Position of top-left corner of sheet.
    static Position origin() @safe pure nothrow @nogc {
        return typeof(return)(RowOffset(0), ColumnOffset(0));
    }

    bool isDefined() const @safe pure nothrow @nogc {
        return row.isDefined && column.isDefined;
    }
}

/** Excel cell address in the format "A1", "A2", ....
 */
struct Address {
    string value;
    alias value this;
}

/** Excel table start position/address.
 *
 * TODO: Should we wrap this algebraic in a struct
 */
alias Start = Algebraic!(Address, Position);

/** Excel table extents as width * height.
 */
struct Extent {
    RowWidth width;
    ColumnHeight height;
}

/** Excel sheet region.
 */
struct Region {
    Start start;
    Extent extent;
}

/** Cell value with `null` state.
 *
 * The state `null` is needed because dense table cells may be uninitialized
 * when copied from the sparse representation in `Sheet.cells`.
 *
 * `IonNull` is used instead of `typeof(null)` to avoid having Value be mapped
 * to a `Result`.
 */
alias Value =
    Variant!(IonNull, bool, TimeOfDay, Date, DateTime, double, long, string);

/// ditto
alias Data = Value; // for backwards compatibility

/// Sheet cell that holds position for use in sparse storage.
struct SparseCell {
    this(string location, RowOffset row, string t, string r, string v, string formula,
         string xmlValue, Position position) @safe pure {
        this.location = location;
        this.row = row;
        this.t = t;
        this.r = r;
        this.v = v;
        this.formula = formula;
        this.xmlValue = xmlValue;
        this.position = position;
        this.value = xmlValue;
    }

    @SILignore
    string location; ///< Location.

    @SILignore
    alias loc = location;

    @SILignore
    RowOffset row; ///< Row. row[r]

    @SILignore
    string t; ///< XML Type attribute

    @SILignore
    string r; // c[r]

    @SILignore
    string v; // value or ptr

    @SILignore
    string formula; // formula

    @SILignore
    string xmlValue; ///< Value stored in cell. TODO: convert this

    @SILignore
    Position position; ///< Position of cell.

    Value value; ///< Decoded value.
}
alias Cell = SparseCell;

/** Excel table as a jagged array.
 *
 * Ref: https://en.wikipedia.org/wiki/Jagged_array
 */
struct JaggedTable {
    inout(Cell[])[] cells() inout @safe pure nothrow @nogc {
        pragma(inline, true);
        return _cells;
    }

    @SILignore
    Cell[][] _cells; /* Not to be referenced explicitly. */
    alias _cells this; /** For backwards compatibility. */

    @SILignore
    RowWidth width; // TODO: use or remove
}

/** Excel jaggedTable as a multidimensional array.
 *
 * Ref: https://en.wikipedia.org/wiki/Array_(data_type)
 */
struct DenseTable {
    inout(Cell) opIndex(
        in RowOffset rowOffset,
        in ColumnOffset columnOffset
    ) inout scope @safe pure nothrow @nogc {
        version(LDC)
            pragma(inline, true);
        return _cells[rowOffset * extent.width + columnOffset];
    }

    // Support for `x..y` notation in slicing operator for the given dimension.
    version(none)
        Offset[2] opSlice(size_t dim)(Offset start, Offset end)
                if (dim >= 0 && dim < 2)
                in(start >= 0 && end <= this.opDollar!dim) {
            pragma(inline, true);
            return [start, end];
        }

	@SILignore
    inout(Cell)[] cells() inout @safe pure nothrow @nogc {
        pragma(inline, true);
        return _cells;
    }

	@SILignore
    auto byRow() const scope @safe pure nothrow {
        struct Result { // TODO: reuse ndslice range instead?
            bool empty() const @property @safe pure nothrow @nogc {
                return _rowOffset == _denseTable.extent.height;
            }

            inout(const(Cell))[] front() inout @property
                    @safe pure nothrow @nogc {
                assert(!empty);
                return _denseTable._cells[_rowOffset * _denseTable.extent.width
                    .. (_rowOffset + 1) * _denseTable.extent.width];
            }

            void popFront() @safe pure nothrow @nogc {
                assert(!empty);
                _rowOffset += 1;
            }

        private:
            const DenseTable _denseTable;
            RowOffset _rowOffset = RowOffset(0);
        }

        return Result(this, RowOffset(0));
    }

    const(Cell)[][] rows() const return scope @safe pure nothrow {
		return byRow().array;
	}

    @SILignore
    Cell[]
        _cells; // TODO: use immutable(Cell)* _cells instead because length is same as extents.width*extents.height
    alias _cells this;
    @SILignore
    Extent extent;
    invariant {
        assert(_cells.length == extent.width * extent.height);
    }
}

/// Sheet.
struct Sheet {
    import std.ascii : toUpper;

    this(string name, immutable(Cell)[] cells,
         Extent extent) @safe pure nothrow @nogc {
        this.name = name;
        this._cells = cells;
        this.extent = extent;
    }

    @SILignore
    immutable(Cell)[]
        cells() const @property return scope @safe pure nothrow @nogc {
        return _cells;
    }

    const string name; ///< Name of sheet.

    @SILignore
    private immutable(Cell)[] _cells; ///< Cells of sheet.

    @SILignore
    inout(JaggedTable) jaggedTable() inout return scope @property
            @trusted pure nothrow { // TODO: make mutable?
        pragma(inline, true);
        if (_jaggedTable is _jaggedTable.init)
            // TODO: remove this if and when we decide if jaggedTable should be removed:
            *(cast(JaggedTable*) &_jaggedTable) = (cast() this).makeJaggedTable();
        return _jaggedTable;
    }

    @SILignore
    private
    JaggedTable makeJaggedTable() const @trusted pure nothrow {
        Cell[][] tab = new Cell[][](extent.height, extent.width);
        foreach (const ref cell; cells)
            tab[cell.position.row][cell.position.col] = cell;
        return typeof(return)(tab, extent.width);
    }

    @SILignore
    inout(DenseTable) denseTable() inout @property
            @trusted pure nothrow { // TODO: make mutable?
        pragma(inline, true);
        if (_denseTable is _denseTable.init)
            // TODO: remove this if and when we decide if densetable should be removed:
            *(cast(DenseTable*) &_denseTable) = (cast() this).makeDenseTable();
        return _denseTable;
    }
	alias table = denseTable;

    @SILignore
    private DenseTable makeDenseTable() const @trusted pure nothrow {
        auto tab = new Cell[](extent.width * extent.height);
        foreach (const ref cell; cells)
            tab[cell.position.row * extent.width + cell.position.col] = cell;
        return typeof(return)(tab, extent);
    }

    @SILignore
    private JaggedTable _jaggedTable;

    @SILignore
    private DenseTable _denseTable;

    @SILignore
    const Extent extent;

    string toString() const @property scope @safe pure {
        import std.array : appender;
        auto result = appender!(typeof(return));
        toString(result);
        return result.data[];
    }

    void toString(Sink)(ref scope Sink sink) const scope {
        import std.algorithm.comparison : max;
        scope lens = new size_t[](extent.width);
        scope tab = makeDenseTable(); // TODO: avoid this allocation
        foreach (const ref row; tab.byRow)
            foreach (const idx, const ref Cell cell; row)
                lens[idx] =
                    max(lens[idx],
                        cell.xmlValue.length); // TODO: use cell.toString

        import std.format : formattedWrite;
        foreach (const ref row; tab.byRow) {
            foreach (const idx, const ref Cell cell; row)
                sink.formattedWrite("%*s, ", lens[idx] + 1,
                                    cell.xmlValue); // TODO: use cell.toString
            sink.formattedWrite("\n");
        }
    }

@SILignore:

    void printTable() const scope @trusted {
        writeln(this); // uses toString(Sink)
    }

    // Column

    Iterator!T getColumn(T)(ColumnOffset col, RowOffset startRow,
                            RowOffset endRow) return scope @trusted {
        auto c = this.iterateColumn!T(col, startRow, endRow);
        return typeof(return)(c.array);
    }

    private enum t = q{
	Iterator!(%1$s) getColumn%2$s(ColumnOffset col, RowOffset startRow, RowOffset endRow) return scope @trusted {
		return getColumn!(%1$s)(col, startRow, endRow);
	}
	};
    static foreach (T;
        ["long", "double", "string", "Date", "TimeOfDay", "DateTime"]
    ) {
        mixin(format(t, T, T[0].toUpper ~ T[1 .. $]));
    }

    ColumnUntyped iterateColumnUntyped(ColumnOffset col, RowOffset startRow,
                                       RowOffset endRow) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(T) iterateColumn(T)(ColumnOffset col, RowOffset startRow,
                                RowOffset endRow) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(long) iterateColumnLong(ColumnOffset col, RowOffset startRow,
                                    RowOffset endRow) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(double) iterateColumnDouble(
        ColumnOffset col,
        RowOffset startRow,
        RowOffset endRow
    ) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(string) iterateColumnString(
        ColumnOffset col,
        RowOffset startRow,
        RowOffset endRow
    ) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(DateTime) iterateColumnDateTime(
        ColumnOffset col,
        RowOffset startRow,
        RowOffset endRow
    ) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(Date) iterateColumnDate(ColumnOffset col, RowOffset startRow,
                                    RowOffset endRow) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    Column!(TimeOfDay) iterateColumnTimeOfDay(
        ColumnOffset col,
        RowOffset startRow,
        RowOffset endRow
    ) return scope @trusted {
        return typeof(return)(&this, col, startRow, endRow);
    }

    // Row

    Iterator!T getRow(T)(RowOffset row, ColumnOffset startColumn,
                         ColumnOffset endColumn) return scope @trusted {
        return typeof(return)(
            this.iterateRow!T(row, startColumn, endColumn).array
        ); // TODO: why .array?
    }

    private enum t2 = q{
	Iterator!(%1$s) getRow%2$s(RowOffset row, ColumnOffset startColumn, ColumnOffset endColumn) return scope @trusted {
		return getRow!(%1$s)(row, startColumn, endColumn);
	}
	};
    static foreach (T;
        ["long", "double", "string", "Date", "TimeOfDay", "DateTime"]
    ) {
        mixin(format(t2, T, T[0].toUpper ~ T[1 .. $]));
    }

    RowUntyped iterateRowUntyped(RowOffset row, ColumnOffset startColumn,
                                 ColumnOffset endColumn) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(T) iterateRow(T)(RowOffset row, ColumnOffset startColumn,
                          ColumnOffset endColumn) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(long) iterateRowLong(RowOffset row, ColumnOffset startColumn,
                              ColumnOffset endColumn) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(double) iterateRowDouble(
        RowOffset row,
        ColumnOffset startColumn,
        ColumnOffset endColumn
    ) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(string) iterateRowString(
        RowOffset row,
        ColumnOffset startColumn,
        ColumnOffset endColumn
    ) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(DateTime) iterateRowDateTime(
        RowOffset row,
        ColumnOffset startColumn,
        ColumnOffset endColumn
    ) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(Date) iterateRowDate(RowOffset row, ColumnOffset startColumn,
                              ColumnOffset endColumn) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }

    Row!(TimeOfDay) iterateRowTimeOfDay(
        RowOffset row,
        ColumnOffset startColumn,
        ColumnOffset endColumn
    ) return scope @trusted {
        return typeof(return)(&this, row, startColumn, endColumn);
    }
}

struct Iterator(T) {
    T[] data;

    this(T[] data) {
        this.data = data;
    }

    @property
    bool empty() const pure nothrow @nogc {
        return this.data.empty;
    }

    void popFront() {
        this.data.popFront();
    }

    @property
    T front() {
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
    const RowOffset row;
    const ColumnOffset startColumn;
    const ColumnOffset endColumn;
    ColumnOffset cur;

    this(Sheet* sheet, RowOffset row, ColumnOffset startColumn,
         ColumnOffset endColumn) pure nothrow /* @nogc */ @safe {
        assert(sheet.jaggedTable.length == sheet.extent.height);
        this.sheet = sheet;
        this.row = row;
        this.startColumn = startColumn;
        this.endColumn = endColumn;
        this.cur = this.startColumn;
    }

    bool empty() const @property pure nothrow @nogc @safe {
        return this.cur >= this.endColumn;
    }

    void popFront() pure nothrow @nogc @safe {
        ++this.cur;
    }

    inout(typeof(this)) save() inout @property pure nothrow @nogc @safe {
        return this;
    }

    inout(Cell) front() inout @property pure nothrow /* @nogc */ @safe {
        return this.sheet.denseTable[row, cur];
    }
}

/// Sheet Row.
struct Row(T) {
    RowUntyped ru;
    T front;

    this(Sheet* sheet, RowOffset row, ColumnOffset startColumn,
         ColumnOffset endColumn) {
        this.ru = RowUntyped(sheet, row, startColumn, endColumn);
        this.read();
    }

    bool empty() const @property pure nothrow @nogc @safe {
        return this.ru.empty;
    }

    void popFront() {
        this.ru.popFront();
        if (!this.empty) {
            this.read();
        }
    }

    inout(typeof(this)) save() inout @property pure nothrow @nogc {
        return this;
    }

    private void read() {
        this.front = convertTo!T(this.ru.front.xmlValue);
    }
}

///
struct ColumnUntyped {
    Sheet* sheet;
    const ColumnOffset col;
    const RowOffset startRow;
    const RowOffset endRow;
    RowOffset cur;

    this(Sheet* sheet, ColumnOffset col, RowOffset startRow,
         RowOffset endRow) @safe {
        assert(sheet.jaggedTable.length == sheet.extent.height);
        this.sheet = sheet;
        this.col = col;
        this.startRow = startRow;
        this.endRow = endRow;
        this.cur = this.startRow;
    }

    bool empty() const @property pure nothrow @nogc @safe {
        return this.cur >= this.endRow;
    }

    void popFront() pure nothrow @nogc @safe {
        ++this.cur;
    }

    inout(typeof(this)) save() inout @property pure nothrow @nogc @safe {
        return this;
    }

    inout(Cell) front() inout @property pure nothrow /* @nogc */ @safe {
        return this.sheet.denseTable[cur, col];
    }
}

/// Sheet Column.
struct Column(T) {
    ColumnUntyped cu;
    T front;

    this(Sheet* sheet, ColumnOffset col, RowOffset startRow, RowOffset endRow) {
        this.cu = ColumnUntyped(sheet, col, startRow, endRow);
        this.read();
    }

    bool empty() const @property pure nothrow @nogc @safe {
        return this.cu.empty;
    }

    void popFront() {
        this.cu.popFront();
        if (!this.empty) {
            this.read();
        }
    }

    inout(typeof(this)) save() inout @property pure nothrow @nogc {
        return this;
    }

    private void read() {
        this.front = convertTo!T(this.cu.front.xmlValue);
    }
}

version(mir_test)
    @safe
    unittest {
        import std.range : isForwardRange;
        import std.meta : AliasSeq;
        static foreach (T;
            AliasSeq!(long, double, DateTime, TimeOfDay, Date, string)
        ) {
            {
                alias C = Column!T;
                alias R = Row!T;
                alias I = Iterator!T;
                static assert(isForwardRange!C, C.stringof);
                static assert(isForwardRange!R, R.stringof);
                static assert(isForwardRange!I, I.stringof);
            }
        }
    }

Date longToDate(long d) @safe {
    // modifed from https://www.codeproject.com/Articles/2750/
    // Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (d == 60) {
        return Date(1900, 2, 29);
    } else if (d < 60) {
        // Because of the 29-02-1900 bug, any serial date
        // under 60 is one off... Compensate.
        ++d;
    }

    // Modified Julian to DMY calculation with an addition of 2415019
    int l = cast(int) d + 68569 + 2415019;
    const int n = int((4 * l) / 146097);
    l = l - int((146097 * n + 3) / 4);
    const int i = int((4000 * (l + 1)) / 1461001);
    l = l - int((1461 * i) / 4) + 31;
    const int j = int((80 * l) / 2447);
    const nDay = l - int((2447 * j) / 80);
    l = int(j / 11);
    const int nMonth = j + 2 - (12 * l);
    const int nYear = 100 * (n - 49) + i + l;
    return Date(nYear, nMonth, nDay);
}

long dateToLong(Date d) @safe {
    // modifed from https://www.codeproject.com/Articles/2750/
    // Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (d.day == 29 && d.month == 2 && d.year == 1900) {
        return 60;
    }

    // DMY to Modified Julian calculated with an extra subtraction of 2415019.
    long nSerialDate = int(
            (1461 * (d.year + 4800 + int((d.month - 14) / 12))) / 4)
        + int((367 * (d.month - 2 - 12 * ((d.month - 14) / 12))) / 12)
        - int((3 * (int((d.year + 4900 + int((d.month - 14) / 12)) / 100))) / 4)
        + d.day - 2415019 - 32075;

    if (nSerialDate < 60) {
        // Because of the 29-02-1900 bug, any serial date
        // under 60 is one off... Compensate.
        nSerialDate--;
    }

    return nSerialDate;
}

version(mir_test)
    @safe
    unittest {
        auto ds = [Date(1900, 2, 1), Date(1901, 2, 28), Date(2019, 06, 05)];
        foreach (const d; ds) {
            const long l = dateToLong(d);
            const Date r = longToDate(l);
            assert(r == d, format("%s %s", r, d));
        }
    }

TimeOfDay doubleToTimeOfDay(double s) @safe {
    import core.stdc.math : lround;
    const double secs = (24.0 * 60.0 * 60.0) * s;

    // TODO not one-hundred my lround is needed
    const int secI = to!int(lround(secs));

    return TimeOfDay(secI / 3600, (secI / 60) % 60, secI % 60);
}

double timeOfDayToDouble(TimeOfDay tod) @safe {
    const long h = tod.hour * 60 * 60;
    const long m = tod.minute * 60;
    const long s = tod.second;
    return (h + m + s) / (24.0 * 60.0 * 60.0);
}

version(mir_test)
    @safe
    unittest {
        auto tods =
            [TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11), TimeOfDay(0, 0, 0),
             TimeOfDay(0, 1, 0), TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
        foreach (const tod; tods) {
            const double d = timeOfDayToDouble(tod);
            assert(d <= 1.0, format("%s", d));
            TimeOfDay r = doubleToTimeOfDay(d);
            assert(r == tod, format("%s %s", r, tod));
        }
    }

double datetimeToDouble(DateTime dt) @safe {
    const double d = dateToLong(dt.date);
    const double t = timeOfDayToDouble(dt.timeOfDay);
    return d + t;
}

DateTime doubleToDateTime(double d) @safe {
    long l = cast(long) d;
    Date dt = longToDate(l);
    TimeOfDay t = doubleToTimeOfDay(d - l);
    return DateTime(dt, t);
}

version(mir_test)
    @safe
    unittest {
        auto ds = [Date(1900, 2, 1), Date(1901, 2, 28), Date(2019, 06, 05)];
        auto tods =
            [TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11), TimeOfDay(0, 0, 0),
             TimeOfDay(0, 1, 0), TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
        foreach (const d; ds) {
            foreach (const tod; tods) {
                DateTime dt = DateTime(d, tod);
                double dou = datetimeToDouble(dt);

                Date rd = longToDate(cast(long) dou);
                assert(rd == d, format("%s %s", rd, d));

                double rest = dou - cast(long) dou;
                TimeOfDay rt = doubleToTimeOfDay(dou - cast(long) dou);
                assert(rt == tod, format("%s %s", rt, tod));

                DateTime r = doubleToDateTime(dou);
                assert(r == dt, format("%s %s", r, dt));
            }
        }
    }

Date stringToDate(string s) @safe {
    import std.array : split;
    import std.string : indexOf;

    if (s.indexOf('/') != -1) {
        auto sp = s.split('/');
        enforce(sp.length == 3, format("[%s]", sp));
        return Date(to!int(sp[2]), to!int(sp[1]), to!int(sp[0]));
    } else {
        return longToDate(to!long(s));
    }
}

bool tryConvertTo(T, S)(S var) {
    return !(tryConvertToImpl!T(Data(var)).isNull());
}

Nullable!(T) tryConvertToImpl(T)(Data var) {
    try {
        return nullable(convertTo!T(var));
    } catch (Exception e) {
        return Nullable!T();
    }
}

T convertTo(T)(string var) @safe {
    import std.math : lround;
    static if (isSomeString!T) {
        return to!T(var);
    } else static if (is(T == bool)) {
        return var == "1";
    } else static if (isIntegral!T) {
        return to!T(var);
    } else static if (isFloatingPoint!T) {
        return to!T(var);
    } else static if (is(T == DateTime)) {
        if (var.canConvertToLong()) {
            return doubleToDateTime(to!long(var));
        } else if (var.canConvertToDouble()) {
            return doubleToDateTime(to!double(var));
        }

        enforce(false, "Can not convert '" ~ var ~ "' to a DateTime");
        assert(false, "Unreachable");
    } else static if (is(T == Date)) {
        if (var.canConvertToLong()) {
            return longToDate(to!long(var));
        } else if (var.canConvertToDouble()) {
            return longToDate(lround(to!double(var)));
        }

        return stringToDate(var);
    } else static if (is(T == TimeOfDay)) {
        const double l = to!double(var);
        return doubleToTimeOfDay(l - cast(long) l);
    } else {
        static assert(false, T.stringof ~ " not supported");
    }
}

private ZipArchive readFile(in string filename) @trusted {
    enforce(exists(filename), "File with name " ~ filename ~ " does not exist");
    return new typeof(return)(read(filename));
}

private static immutable workbookXMLPath = "xl/workbook.xml";
private static immutable sharedStringXMLPath = "xl/sharedStrings.xml";
private static immutable relsXMLPath = "xl/_rels/workbook.xml.rels";

private static expandTrusted(ZipArchive za,
                             ArchiveMember de) @trusted /* TODO: pure */ {
    static if (tme)
        auto sw = StopWatch(AutoStart.yes);
    auto ret = za.expand(de);
    static if (tme)
        writeln("expand length:", ret.length, " took: ", sw.peek());
    return ret;
}

struct Relationships {
    string id;
    string file;
}

alias RelationshipsById = Relationships[string];

/// Workbook.
struct Workbook {
    @SILignore
    static typeof(this) fromFile(in string filename) @trusted {
        return typeof(return)(filename, new ZipArchive(read(filename)));
    }

    @SILignore
    static typeof(this) fromBytes(void[] buffer) @trusted {
        return typeof(return)("", new ZipArchive(buffer));
    }

    @SILignore
    package SheetNameId[] sheetNameIds() @safe /* TODO: pure */ {
        auto dom = workbookDOM();
        if (dom.children.length != 1)
            return [];

        auto workbook = dom.children[0];
        if (workbook.name != "workbook" && workbook.name != "s:workbook")
            return [];

        const sheetName = workbook.name == "workbook" ? "sheets" : "s:sheets";
        auto sheetsRng = workbook.children.filter!(c => c.name == sheetName);
        if (sheetsRng.empty)
            return [];

        return sheetsRng
            .front
            .children
            .map!(
                // TODO: optimize by using indexOf or find
                s => SheetNameId(
                    s.attributes.filter!(a => a.name == "name").front.value
                     .decodeXML(),
                    s.attributes.filter!(a => a.name == "sheetId").front.value
                     .to!int(),
                    s.attributes.filter!(a => a.name == "r:id").front.value
                ))
            .array;
    }

    /// Returns: Sheets as an eagerly evaluated and internally cached array.
    auto sheets() {
        if (_sheets is null)
            _sheets = bySheet.array;
        return _sheets;
    }

    /// Returns: Lazy range over sheets.
    @SILignore
    auto bySheet() @safe {
        return sheetNameIds.map!((const scope SheetNameId sheetNameId) {
            return extractSheet(_za, relationships, filename, sheetNameId.rid,
                                sheetNameId.name);
        });
    }

    /// Get (and cache) DOM.
    private DOMEntity!(string) workbookDOM() @safe /* TODO: pure */ {
        import dxml.parser : Config, SkipComments, SkipPI, SplitEmpty,
                             ThrowOnEntityRef;
        auto ent = workbookXMLPath in _za.directory;
        // TODO: use enforce(ent ! is null); instead?
        if (ent is null)
            return typeof(return).init;
        if (_wbDOM == _wbDOM.init) {
            auto est = _za.expandTrusted(*ent).convertToString();
            static if (tme)
                auto sw = StopWatch(AutoStart.yes);
            _wbDOM = est.parseDOM!(Config(
                SkipComments
                    .no, // TODO: change to SkipComments.yes and validate
                SkipPI.no, // TODO: change to SkipPI.yes and validate
                SplitEmpty.no, // default is ok
                ThrowOnEntityRef.yes
            ))(); // default is ok
            enforce(
                _wbDOM.children.length == 1,
                "Expected a single DOM child but got "
                    ~ _wbDOM.children.length.to!string
            );
            static if (tme)
                writeln("parseDOM length:", est.length, " took: ", sw.peek());
        }

        return _wbDOM;
    }

    @SILignore
    private RelationshipsById relationships() @safe /* TODO: pure */ {
        if (_rels is null)
            _rels = parseRelationships(_za.directory[relsXMLPath]);
        return _rels;
    }

    @SILignore
    private RelationshipsById parseRelationships(ArchiveMember am) @safe {
        auto est = _za.expandTrusted(am).convertToString();
        // import std.digest : digest;
        // import std.digest.md : MD5;
        // writeln("am.name:", am.name, " md5:", est.digest!MD5);
        auto dom = est.parseDOM();
        enforce(
            dom.children.length == 1,
            "Expected a single DOM child but got "
                ~ dom.children.length.to!string
        );

        auto rel = dom.children[0];
        enforce(
            rel.name == "Relationships",
            "Expected rel.name to be \"Relationships\" but was " ~ rel.name
        );

        typeof(return) ret;
        static if (is(typeof(ret.reserve(size_t.init)) == void)) {
            /* Use reserve() when AA gets it or `RelationshipsById` is a custom hash
			 * map. */
            ret.reserve(rel.children.length);
        }

        foreach (ref r; rel.children.filter!(c => c.name == "Relationship")) {
            Relationships tmp;
            tmp.id = r.attributes.filter!(a => a.name == "Id").front.value;
            tmp.file =
                r.attributes.filter!(a => a.name == "Target").front.value;
            ret[tmp.id] = tmp;
        }

        enforce(!ret.empty);
        return ret;
    }

    @SILignore
    private string[] sharedEntries() @safe /* TODO: pure */ {
        if (_sharedEntries is null)
            if (ArchiveMember* amPtr = sharedStringXMLPath in _za.directory)
                _sharedEntries = readSharedEntries(_za, *amPtr);
        return _sharedEntries;
    }

    @SILignore
    const string filename;

    /* TODO: remove when https://github.com/libmir/mir-core/pull/79 */
    @SILignore
    private inout(ZipArchive) _za() inout @safe pure nothrow @nogc {
        return __za;
    }

    @SILignore
    private ZipArchive __za;

    @SILignore
    private DOMEntity!string _wbDOM; ///< Workbook.

    @SILignore
    private RelationshipsById _rels;

    @SILignore
    private string[] _sharedEntries;

    @SILignore
    private const(Sheet)[] _sheets;
}

version(mir_test)
    @safe
    unittest {
        const sheets = Workbook.fromFile("test/data/50xP_sheet1.xlsx").sheets();
        assert(sheets.length == 1);
        assert(sheets[0].cells.length == 6008);
        // writeln(sheets[0]);
        // TODO: that cells are of time Data or DateTime
    }

/// benchmark reading of "50xP_sheet1.xlsx"
version(mir_benchmark)
    @safe
    unittest {
        const path = "test/data/50xP_sheet1.xlsx";
        import std.meta : AliasSeq;
        void use_sheetNamesAndreadSheet() @trusted {
            foreach (const i, const ref s; sheetNames(path)) {
                auto sheet = readSheet(path, s.name);
            }
        }

        size_t use_bySheet() @trusted {
            auto file = Workbook.fromFile(path);
            typeof(return) i = 0;
            foreach (ref sheet; file.bySheet) {
                i++;
            }

            return i;
        }

        alias funs = AliasSeq!(/* use_sheetNamesAndreadSheet, */
            use_bySheet);
        auto results = benchmarkMin!(funs)(runCount);
        foreach (const i, fun; funs) {
            writeln(fun.stringof[0 .. $ - 2], "(\"", path, "\") took ",
                    results[i]);
        }
    }

/// Sheet name, id and rid.
struct SheetNameId {
    string name;
    int id;
    string rid;
}

/// Strip BOM and convert ubyte[] to a string.
string convertToString(inout(ubyte)[] d) @trusted {
    import std.encoding : getBOM, BOM, transcode;
    const b = getBOM(d);
    switch (b.schema) {
        case BOM.none:
            return cast(string) d; // TODO: remove this cast
        case BOM.utf8:
            return cast(string) (d[3 .. $]); // TODO: remove this cast
        case BOM.utf16be:
        case BOM.utf16le:
        case BOM.utf32be:
        case BOM.utf32le:
            goto default;
        default:
            string ret;
            transcode(d, ret);
            return ret;
    }
}

/// Read sheet names stored in `filename`.
deprecated("Use `Workbook.fromFile(filename).sheetNameIds` instead")
SheetNameId[] sheetNames(in string filename) @trusted {
    return Workbook.fromFile(filename).sheetNameIds();
}

version(mir_test)
    @safe
    unittest {
        auto r = Workbook.fromFile("test/data/multitable.xlsx").sheetNameIds();
        assert(r
            == [SheetNameId("wb1", 1, "rId2"), SheetNameId("wb2", 2, "rId3"),
                SheetNameId("Sheet3", 3, "rId4")]);
    }

version(mir_test)
    @safe
    unittest {
        auto r = Workbook.fromFile("test/data/sheetnames.xlsx").sheetNameIds();
        assert(r == [SheetNameId("A & B ;", 1, "rId2")]);
    }

deprecated("Use File.relationships(ArchiveMember am) instead")
RelationshipsById parseRelationships(ZipArchive za, ArchiveMember am) @safe {
    auto dom = za.expandTrusted(am).convertToString().parseDOM();
    enforce(
        dom.children.length == 1,
        "Expected a single DOM child but got " ~ dom.children.length.to!string
    );

    auto rel = dom.children[0];
    enforce(rel.name == "Relationships",
            "Expected rel.name to be \"Relationships\" but was " ~ rel.name);

    typeof(return) ret;
    static if (is(typeof(ret.reserve(size_t.init)) == void)) {
        /* Use reserve() when AA gets it or `RelationshipsById` is a custom hash
		 * map. */
        ret.reserve(rel.children.length);
    }

    foreach (ref r; rel.children.filter!(c => c.name == "Relationship")) {
        Relationships tmp;
        tmp.id = r.attributes.filter!(a => a.name == "Id").front.value;
        tmp.file = r.attributes.filter!(a => a.name == "Target").front.value;
        ret[tmp.id] = tmp;
    }

    enforce(!ret.empty);
    return ret;
}

/// Read sheet named `sheetName` from `filename`.
Sheet readSheet(in string filename, in string sheetName) @safe {
    const SheetNameId[] sheets = Workbook.fromFile(filename).sheetNameIds();
    auto sRng = sheets.filter!(s => s.name == sheetName);
    enforce(!sRng.empty,
            "No sheet with name " ~ sheetName ~ " found in file " ~ filename);
    return readSheetImpl(filename, sRng.front.rid, sheetName);
}

string eatXlPrefix(scope return string fn) @safe pure nothrow @nogc {
    static immutable xlPrefixes = ["xl//", "/xl/"];
    foreach (const p; xlPrefixes) {
        // TODO: use fn.skipOver("xl//", "/xl/") when it’s nothrow @nogc
        if (fn.startsWith(p)) {
            return fn[p.length .. $];
        }
    }

    return fn;
}

deprecated(
    "Use Workbook.fromFile(filename).sheets().filter(sheet => sheet.name == sheetName) instead"
) Sheet readSheetImpl(in string filename, in string rid,
                      in string sheetName) @safe {
    scope(failure)
        writefln("Failed at file '%s' and sheet '%s'", filename, rid);
    auto za = readFile(filename);
    auto rels = parseRelationships(za, za.directory[relsXMLPath]);
    return extractSheet(za, rels, filename, rid, sheetName);
}

private
Sheet extractSheet(ZipArchive za, in RelationshipsById rels, in string filename,
                   in string rid, in string sheetName) @trusted {
    string[] ss; /* shared strings (table) */
    if (ArchiveMember* amPtr = sharedStringXMLPath in za.directory)
        ss = readSharedEntries(
            za, *amPtr); // TODO: cache this into File.sharedStrings

    const Relationships* sheetRel = rid
        in rels; // TODO: move this calculation to caller and pass Relationships as rels
    enforce(sheetRel !is null,
            format("Could not find '%s' in '%s'", rid, filename));
    const fn = "xl/" ~ eatXlPrefix(sheetRel.file);
    ArchiveMember* sheet = fn in za.directory;
    enforce(
        sheet !is null,
        format("sheetRel.file orig '%s', fn %s not in [%s]", sheetRel.file, fn,
               za.directory.keys())
    );

    Cell[] cells1 = readCells(za, *sheet); // hot spot!
    Cell[] cells = insertValueIntoCell(cells1, ss);

    Position maxPos = Position.origin;
    foreach (ref c; cells) {
        c.position = toPos(c.r);
        maxPos = elementMax(maxPos, c.position);
    }

    const extent =
        Extent(RowWidth(maxPos.col + 1), ColumnHeight(maxPos.row + 1));

    // debug writeln("filename:", filename, " maxPos:", maxPos, " extent:", extent);
    import std.exception : assumeUnique;
    return Sheet(sheetName, cells.assumeUnique, extent);
}

string[] readSharedEntries(ZipArchive za, ArchiveMember am) @safe {
    auto dom = za.expandTrusted(am).convertToString().parseDOM(); // TODO: cache
    if (dom.type != EntityType.elementStart)
        return typeof(return).init;
    assert(dom.children.length == 1);

    auto sst = dom.children[0];
    assert(sst.name == "sst");

    if (sst.type != EntityType.elementStart || sst.children.empty)
        return typeof(return).init;

    Appender!(typeof(return)) ret; // TODO: reserve?
    foreach (ref si; sst.children.filter!(c => c.name == "si")) {
        if (si.type != EntityType.elementStart)
            continue;
        //ret ~= extractData(si);
        string tmp;
        foreach (ref tORr; si.children) {
            if (tORr.name == "t" && tORr.type == EntityType.elementStart
                    && !tORr.children.empty) {
                //ret ~= Data(convert(tORr.children[0].text));
                ret ~= tORr.children[0].text.decodeXML;
            } else if (tORr.name == "r") {
                foreach (ref r; tORr.children.filter!(r => r.name == "t")) {
                    if (r.type == EntityType.elementStart
                            && !r.children.empty) {
                        tmp ~= r.children[0].text.decodeXML;
                    }
                }
            } else {
                //ret ~= Data.init;
                ret ~= "";
            }
        }

        if (!tmp.empty) {
            //ret ~= Data(convert(tmp));
            ret ~= tmp.decodeXML;
        }
    }

    return ret.data;
}

string extractData(DOMEntity!string si) {
    string tmp;
    foreach (ref tORr; si.children) {
        if (tORr.name == "t") {
            if (!tORr.attributes.filter!(a => a.name == "xml:space").empty) {
                return "";
            } else if (tORr.type == EntityType.elementStart
                           && !tORr.children.empty) {
                return tORr.children[0].text;
            } else {
                return "";
            }
        } else if (tORr.name == "r") {
            foreach (ref r; tORr.children.filter!(r => r.name == "t")) {
                tmp ~= r.children[0].text;
            }
        }
    }

    if (!tmp.empty) {
        return tmp;
    }

    assert(false);
}

private bool canConvertToLong(in string s) @safe pure nothrow @nogc {
    import std.utf : byChar;
    import std.ascii : isDigit;
    if (s.empty)
        return false;
    return s.byChar.all!isDigit();
}

version(ctRegex_test)
    version(unittest) {
        import std.regex : ctRegex, matchAll;
        private static immutable rs = r"[\+-]{0,1}[0-9][0-9]*\.[0-9]*";
        private static immutable rgx = ctRegex!rs;
        private bool canConvertToDoubleOld(in string s) @safe {
            auto cap = matchAll(s, rgx);
            return cap.empty || cap.front.hit != s ? false : true;
        }
    }

private bool canConvertToDouble(string s) pure @safe nothrow @nogc {
    if (s.startsWith('+', '-')) {
        s = s[1 .. $];
    }

    if (s.empty) {
        return false;
    }

    if (s[0] < '0' || s[0] > '9') { // at least one in [0-9]
        return false;
    }

    s = s[1 .. $];

    if (s.empty) {
        return true;
    }

    while (!s.empty && s[0] >= '0' && s[0] <= '9') {
        s = s[1 .. $];
    }

    if (s.empty) {
        return true;
    }

    if (s[0] != '.') {
        return false;
    }

    s = s[1 .. $];
    if (s.empty) {
        return true;
    }

    while (!s.empty && s[0] >= '0' && s[0] <= '9') {
        s = s[1 .. $];
    }

    return s.empty;
}

version(mir_test)
    @safe
    unittest {
        static struct Test {
            string tt;
            bool rslt;
        }

        auto tests = [
            Test("-", false),
            Test("0.0", true),
            Test("-0.", true),
            Test("-0.0", true),
            Test("-0.a", false),
            Test("-0.0", true),
            Test("-1100.0", true)
        ];
        foreach (const t; tests) {
            version(ctRegex_test)
                assert(
                    canConvertToDouble(t.tt) == canConvertToDoubleOld(t.tt),
                    format("%s %s %s %s", t.tt, canConvertToDouble(t.tt),
                           canConvertToDoubleOld(t.tt), t.rslt)
                );
            assert(canConvertToDouble(t.tt) == t.rslt,
                   format("%s %s %s", t.tt, canConvertToDouble(t.tt), t.rslt));
        }
    }

deprecated("Use dxml.util.decodeXML instead")
string removeSpecialCharacter(string s) {
    struct ToRe {
        string from;
        string to;
    }

    immutable ToRe[] toRe =
        [ToRe("&amp;", "&"), ToRe("&gt;", "<"), ToRe("&lt;", ">"),
         ToRe("&quot;", "\""), ToRe("&apos;", "'")];

    string replaceStrings(string s) {
        import std.algorithm.searching : canFind;
        // TODO: use substitute.array
        import std.array : replace;
        foreach (const ref tr; toRe) {
            while (canFind(s, tr.from)) {
                s = s.replace(tr.from, tr.to);
            }
        }

        return s;
    }

    return replaceStrings(s);
}

Cell[] readCells(ZipArchive za, ArchiveMember am) /* TODO: @safe */ {
    auto dom =
        za.expandTrusted(am).convertToString().parseDOM(); // TODO: cache?
    assert(dom.children.length == 1);

    auto ws = dom.children[0];
    if (ws.name != "worksheet")
        return typeof(return).init;

    auto sdRng = ws.children.filter!(c => c.name == "sheetData");
    assert(!sdRng.empty);

    if (sdRng.front.type != EntityType.elementStart)
        return typeof(return).init;

    auto rows = sdRng.front.children.filter!(r => r.name == "row");

    Appender!(typeof(return)) ret; // TODO: reserve()?
    foreach (ref row; rows) {
        if (row.type != EntityType.elementStart || row.children.empty) {
            continue;
        }

        foreach (ref c; row.children.filter!(r => r.name == "c")) {
            Cell tmp;
            tmp.row = RowOffset(row.attributes.filter!(a => a.name == "r").front
                                   .value.to!(typeof(RowOffset.value)));
            tmp.r = c.attributes.filter!(a => a.name == "r").front.value;
            auto t = c.attributes.filter!(a => a.name == "t");
            if (t.empty) {
                // we assume that no t attribute means direct number
                //writefln("Found a strange empty cell \n%s", c);
            } else {
                tmp.t = t.front.value;
            }

            if (tmp.t == "s" || tmp.t == "n") {
                if (c.type == EntityType.elementStart) {
                    auto v = c.children.filter!(c => c.name == "v");
                    //enforce(!v.empty, format("r %s", tmp.row));
                    if (!v.empty && v.front.type == EntityType.elementStart
                            && !v.front.children.empty) {
                        tmp.v = v.front.children[0].text;
                    } else {
                        tmp.v = "";
                    }
                }
            } else if (tmp.t == "inlineStr") {
                auto is_ = c.children.filter!(c => c.name == "is");
                tmp.v = extractData(is_.front);
            } else if (c.type == EntityType.elementStart) {
                auto v = c.children.filter!(c => c.name == "v");
                if (!v.empty && v.front.type == EntityType.elementStart
                        && !v.front.children.empty) {
                    tmp.v = v.front.children[0].text;
                }
            }

            if (c.type == EntityType.elementStart) {
                auto f = c.children.filter!(c => c.name == "f");
                if (!f.empty && f.front.type == EntityType.elementStart) {
                    tmp.formula = f.front.children[0].text;
                }
            }

            ret ~= tmp;
        }
    }

    return ret.data; // TODO: assumeUnique?
}

/**
 * Param: `ss` is the shared string (table)
 */
Cell[] insertValueIntoCell(Cell[] cells,
                           in string[] ss) @trusted /* TODO: pure */ {
    immutable excepted = ["f", /* formula */
                          "n", /* number */
                          "s", /* string? */
                          "d", /* date */
                          "b", /* boolean */
                          "e", /* string */
                          "str", /* string */
                          "inlineStr" /* inline string */
    ]; // TODO: what are these?
    immutable same = ["n", "e", "str", "inlineStr"]; // TODO: what are these?
    foreach (ref Cell c; cells) {
        // debug writeln("c.t:", c.t, " c.v:", c.v);
        assert(excepted.canFind(c.t) || c.t.empty,
               format("'%s' not in [%s]", c.t, excepted));
        if (c.t.empty) {
            c.xmlValue = c.v.decodeXML;
        } else if (same.canFind(c.t)) {
            c.xmlValue = c.v.decodeXML;
        } else if (c.t == "b") {
            c.xmlValue = c.v.decodeXML;
        } else if (!c.v.empty) {
            c.xmlValue = ss[c.v.to!size_t]; /* shared string table? */
        }

        switch (c.t) {
            case "b": // boolean
                if (c.xmlValue == "0")
                    c.value = false;
                else if (c.xmlValue == "1")
                    c.value = true;
                else
                    c.value = c.xmlValue;
                break;
            case "n": // number
                if (c.v.canFind("."))
                    c.value = c.v.to!double;
                else
                    c.value = c.v.to!long;
                break;
            case "d": // date
                // TODO: c.value = c.xmlValue.convertTo!DateTime;
                c.value = c.xmlValue;
                break;
            default:
                c.value = c.xmlValue;
                break;
        }
    }

    return cells;
}

Position toPos(in string s) @safe pure {
    import std.string : indexOfAny;
    import std.math : pow;
    ptrdiff_t fn = s.indexOfAny("0123456789");
    enforce(fn != -1, s);
    RowOffset row = to!RowOffset(to!int(s[fn .. $]) - 1);
    ColumnOffset col = ColumnOffset(0);
    string colS = s[0 .. fn];
    foreach (const idx, char c; colS) {
        col = col * 26 + (c - 'A' + 1);
    }

    return Position(row, ColumnOffset(col - 1));
}

version(mir_test)
    @safe
    pure unittest {
        assert(toPos("A1").col == 0);
        assert(toPos("Z1").col == 25);
        assert(toPos("AA1").col == 26);
    }

Position elementMax(Position a, Position b) @safe pure nothrow @nogc {
    import std.algorithm.comparison : max;
    return Position(max(a.row, b.row), max(a.col, b.col));
}

deprecated("This function is unused and will therefore be removed")
string specialCharacterReplacement(string s) @safe pure nothrow {
    import std.algorithm.iteration : substitute;
    import std.array : replace;
    // TODO: use substitute.array
    return s.replace("\"", "&quot;").replace("'", "&apos;").replace("<", "&lt;")
            .replace(">", "&gt;").replace("&", "&amp;");
}

version(mir_test)
    @safe
    pure nothrow unittest {
        assert("&".specialCharacterReplacement == "&amp;");
    }

deprecated("Use dxml.util.decodeXML instead")
string specialCharacterReplacementReverse(string s) @safe pure nothrow {
    import std.array : replace;
    // TODO: reuse existing Phobos function or generalize to all special characters
    return s.replace("&quot;", "\"").replace("&apos;", "'").replace("&lt;", "<")
            .replace("&gt;", ">").replace("&amp;", "&");
}

version(none) // disabled because specialCharacterReplacementReverse is deprecated
    version(mir_test)
        @safe
        pure nothrow unittest {
            assert("&quot;".specialCharacterReplacementReverse == "\"");
            assert("&apos;".specialCharacterReplacementReverse == "'");
            assert("&lt;".specialCharacterReplacementReverse == "<");
            assert("&gt;".specialCharacterReplacementReverse == ">");
            assert("&amp;".specialCharacterReplacementReverse == "&");
        }

version(mir_test)
    @safe
    unittest {
        import std.math : isClose;
        auto r = readSheet("test/data/multitable.xlsx", "wb1");
        {
            const e = r.jaggedTable[12][5];
            assert(isClose(e.xmlValue.to!double(), 26.74), format("%s", e));
        }

        {
            const e = r.jaggedTable[13][5];
            assert(isClose(e.xmlValue.to!double(), -26.74), format("%s", e));
        }

        {
            const e = r.denseTable[RowOffset(12), ColumnOffset(5)];
            assert(isClose(e.xmlValue.to!double(), 26.74), format("%s", e));
        }

        {
            const e = r.denseTable[RowOffset(13), ColumnOffset(5)];
            assert(isClose(e.xmlValue.to!double(), -26.74), format("%s", e));
        }
    }

version(mir_test)
    @safe
    unittest {
        auto s = readSheet("test/data/multitable.xlsx", "wb1");
        const expectedCells = [
            Cell("", RowOffset(3), "s", "D3", "0", "", "a",
                 Position(RowOffset(2), ColumnOffset(3))),
            Cell("", RowOffset(3), "s", "E3", "1", "", "b",
                 Position(RowOffset(2), ColumnOffset(4))),
            Cell("", RowOffset(4), "s", "D4", "2", "", "1",
                 Position(RowOffset(3), ColumnOffset(3))),
            Cell("", RowOffset(4), "s", "E4", "3", "", "\"one\"",
                 Position(RowOffset(3), ColumnOffset(4))),
            Cell("", RowOffset(5), "s", "D5", "4", "", "2",
                 Position(RowOffset(4), ColumnOffset(3))),
            Cell("", RowOffset(5), "s", "E5", "5", "", "\"two\"",
                 Position(RowOffset(4), ColumnOffset(4))),
            Cell("", RowOffset(6), "s", "D6", "6", "", "3",
                 Position(RowOffset(5), ColumnOffset(3))),
            Cell("", RowOffset(6), "s", "E6", "7", "", "\"three\"",
                 Position(RowOffset(5), ColumnOffset(4))),
            Cell("", RowOffset(7), "", "B7", "", "", "",
                 Position(RowOffset(6), ColumnOffset(1))),
            Cell("", RowOffset(7), "s", "F7", "1", "", "b",
                 Position(RowOffset(6), ColumnOffset(5))),
            Cell("", RowOffset(7), "s", "G7", "0", "", "a",
                 Position(RowOffset(6), ColumnOffset(6))),
            Cell("", RowOffset(8), "n", "C8", "0.504409722222222", "",
                 "0.504409722222222", Position(RowOffset(7), ColumnOffset(2))),
            Cell("", RowOffset(8), "s", "F8", "3", "", "\"one\"",
                 Position(RowOffset(7), ColumnOffset(5))),
            Cell("", RowOffset(8), "s", "G8", "2", "", "1",
                 Position(RowOffset(7), ColumnOffset(6))),
            Cell("", RowOffset(9), "s", "F9", "5", "", "\"two\"",
                 Position(RowOffset(8), ColumnOffset(5))),
            Cell("", RowOffset(9), "s", "G9", "4", "", "2",
                 Position(RowOffset(8), ColumnOffset(6))),
            Cell("", RowOffset(10), "s", "F10", "7", "", "\"three\"",
                 Position(RowOffset(9), ColumnOffset(5))),
            Cell("", RowOffset(10), "s", "G10", "6", "", "3", Position(RowOffset(9), ColumnOffset(6))),
            Cell("", RowOffset(11), "s", "AC11", "8", "", "Foo", Position(RowOffset(10), ColumnOffset(28))),
            Cell("", RowOffset(12), "s", "B12", "9", "", "Hello World", Position(RowOffset(11), ColumnOffset(1))),
            Cell("", RowOffset(13), "n", "E13", "13.37", "", "13.37", Position(RowOffset(12), ColumnOffset(4))),
            Cell("", RowOffset(13), "n", "F13", "26.74", "E13*2", "26.74", Position(RowOffset(12), ColumnOffset(5))),
            Cell("", RowOffset(14), "n", "F14", "-26.74", "-E13*2", "-26.74", Position(RowOffset(13), ColumnOffset(5))),
            Cell("", RowOffset(16), "n", "B16", "1", "", "1", Position(RowOffset(15), ColumnOffset(1))),
            Cell("", RowOffset(16), "n", "C16", "2", "", "2", Position(RowOffset(15), ColumnOffset(2))),
            Cell("", RowOffset(16), "n", "D16", "3", "", "3", Position(RowOffset(15), ColumnOffset(3))),
            Cell("", RowOffset(16), "n", "E16", "4", "", "4", Position(RowOffset(15), ColumnOffset(4))),
            Cell("", RowOffset(16), "n", "F16", "5", "", "5", Position(RowOffset(15), ColumnOffset(5)))
        ];
        foreach (const i, const ref cell; s.cells) {
            // compare all but last field for now as that’s subject to change
            assert(
                cell.tupleof[0 .. $ - 1] == expectedCells[i].tupleof[0 .. $ - 1]
            );
        }

        assert(s.jaggedTable.length == 16);
        foreach (const Cell[] row; s.jaggedTable)
            assert(row.length == 29);
    }

version(mir_test)
    @safe
    unittest {
        auto s = readSheet("test/data/multitable.xlsx", "wb1");

        auto r =
            s.iterateRow!long(RowOffset(15), ColumnOffset(1), ColumnOffset(6));

        auto expected = [1, 2, 3, 4, 5];
        assert(equal(r, expected), format("%s", r));

        auto r2 =
            s.getRow!long(RowOffset(15), ColumnOffset(1), ColumnOffset(6));
        assert(equal(r, expected));

        auto it =
            s.iterateRowLong(RowOffset(15), ColumnOffset(1), ColumnOffset(6));
        assert(equal(r2, it));

        auto it2 =
            s.iterateRowUntyped(RowOffset(15), ColumnOffset(1), ColumnOffset(6))
             .map!(it => format("%s", it)).array;
    }

version(mir_test)
    @safe
    unittest {
        auto s = readSheet("test/data/multitable.xlsx", "wb2");
        //writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
        auto rslt =
            s.iterateColumn!Date(ColumnOffset(1), RowOffset(1), RowOffset(6));
        auto rsltUt =
            s.iterateColumnUntyped(ColumnOffset(1), RowOffset(1), RowOffset(6))
             .map!(it => format("%s", it)).array;
        assert(!rsltUt.empty);

        auto target = [Date(2019, 5, 01), Date(2016, 12, 27), Date(1976, 7, 23),
                       Date(1986, 7, 2), Date(2038, 1, 19)];
        assert(equal(rslt, target), format("\n%s\n%s", rslt, target));

        auto it = s.getColumn!Date(ColumnOffset(1), RowOffset(1), RowOffset(6));
        assert(equal(rslt, it));

        auto it2 = s.getColumnDate(ColumnOffset(1), RowOffset(1), RowOffset(6));
        assert(equal(rslt, it2));
    }

version(mir_test)
    @safe
    unittest {
        auto s = readSheet("test/data/multitable.xlsx", "Sheet3");
        // writeln(s.jaggedTable[0][0].xmlValue);
        assert(s.jaggedTable[0][0].xmlValue.to!long(),
               format("%s", s.jaggedTable[0][0].xmlValue));
        //assert(s.jaggedTable[0][0].canConvertTo(CellType.bool_));
    }

version(mir_test)
    @system
    unittest {
        import std.file : dirEntries, SpanMode;
        import std.algorithm.searching : endsWith;
        size_t totalSheetCount;
        size_t totalCellCount;

        /* Reading the file row_col_format16.xlsx causes out of memory because one of its cells position address strings is
     * mapped to the column index 16384.
     */
        foreach (const de;
            dirEntries("test/data/xlsx_files/", "*.xlsx", SpanMode.depth)
                .filter!(a => (!a.name.endsWith(
                    "data02.xlsx", "data03.xlsx", "data04.xlsx",
                    "row_col_format16.xlsx", "row_col_format18.xlsx")))) {
            auto sn = sheetNames(de.name);
            foreach (const s; sn) {
                totalSheetCount += 1;
                auto sheet = readSheet(de.name, s.name);
                foreach (const cell; sheet.cells) {
                    totalCellCount += 1;
                }
            }
        }

        assert(totalSheetCount == 551);
        assert(totalCellCount == 4860);
    }

version(mir_test)
    @safe
    unittest {
        auto sheet = readSheet("test/data/testworkbook.xlsx", "ws1");
        //writefln("%(%s\n%)", sheet.cells);
        //writeln(sheet.toString());
        assert(sheet.jaggedTable[2][3].xmlValue.to!long() == 1337);

        auto c =
            sheet.getColumnLong(ColumnOffset(3), RowOffset(2), RowOffset(5));
        auto r = [1337, 2, 3];
        assert(equal(c, r), format("%s %s", c, sheet.toString()));

        auto c2 =
            sheet.getColumnString(ColumnOffset(4), RowOffset(2), RowOffset(5));
        string f2 = sheet.jaggedTable[2][4].xmlValue;
        assert(f2 == "hello", f2);
        f2 = sheet.jaggedTable[3][4].xmlValue;
        assert(f2 == "sil", f2);
        f2 = sheet.jaggedTable[4][4].xmlValue;
        assert(f2 == "foo", f2);
        auto r2 = ["hello", "sil", "foo"];
        assert(equal(c2, r2), format("%s", c2));
    }

version(mir_test)
    @safe
    unittest {
        import std.math : isClose;
        auto sheet = readSheet("test/data/toto.xlsx", "Trades");

        // writefln("%(%s\n%)", sheet.cells);

        auto r =
            sheet.getRowString(RowOffset(1), ColumnOffset(0), ColumnOffset(2))
                 .array;

        const double d = to!double(r[1]);
        assert(isClose(d, 38204642.510000));
    }

version(mir_test)
    @safe
    unittest {
        const sheet = readSheet("test/data/leading_zeros.xlsx", "Sheet1");
        auto a2 = sheet.cells.filter!(c => c.r == "A2");
        assert(!a2.empty);
        assert(a2.front.xmlValue == "0012", format("%s", a2.front));
    }

version(mir_test)
    @safe
    unittest {
        auto s = readSheet("test/data/datetimes.xlsx", "Sheet1");
        //writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
        // TODO: this manual patch should be removed in favor of auto-detection:
        auto rslt = s.iterateColumn!DateTime(ColumnOffset(0), RowOffset(0),
                                             RowOffset(2));
        assert(!rslt.empty);

        auto target = [DateTime(Date(1986, 1, 11), TimeOfDay.init),
                       DateTime(Date(1986, 7, 2), TimeOfDay.init)];
        assert(
            equal(rslt, target),
            format("\ngot: %s\nexp: %s\ntable %s", rslt, target, s.toString())
        );
    }

version(mir_test) {
    import std.algorithm.comparison : equal;
}

/** Variant of Phobos `benchmark` that, instead of sum all run times, returns
	minimum of all run times as that is a more stable metric.
 */
version(mir_benchmark)
    private
    Duration[funs.length] benchmarkMin(funs...)(uint n) if (funs.length >= 1) {
        import std.algorithm.comparison : min;
        Duration[funs.length] result;
        auto sw = StopWatch(AutoStart.yes);
        foreach (const i, fun; funs) {
            result[i] = Duration.max;
            foreach (const j; 0 .. n) {
                sw.reset();
                sw.start();
                fun();
                sw.stop();
                result[i] = min(result[i], sw.peek());
            }
        }

        return result;
    }

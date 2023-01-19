# TODO

- Use common mir types

- Read all XML-files to binary mir-ion and then to Tables

- Use an algebraic cell value shared with mir-csv and mir-excel

- Remove const fields

- Remove aliases

- Use mir’s version of timestamp parsing because it’s faster

- Workbook => mir Object

- Workbook => StringMap!(name, Sheet) eagerly like in JSON

- Look at how `Openpyxml._tables` is implemented to detect tables
  and use the same to decode Sheet.tables()
  For reference see https://samukweku.github.io/data-wrangling-blog/spreadsheet/python/pandas/openpyxl/2020/05/19/Access-Tables-In-Excel.html#Option-2---The-better-way-:
  Might be worthwhile looking into how multitable.xlsx is used.
- Four inner loops that converts a dense table to binary ion
  Iterate the sparse cells twice
  - First iteration: Find the most dimensions of the dense matrix
  - Second iteration: Fill in the values in the dense matrix
- IonNull and typeof(null) are serialized to the same code in binary ion

- Same wrapping as for mir-csv

- Merge `SparseCell.location` and `SparseCell.position` into one single field

- Ask Ilya if and how we should use mir.string_map. Discuss using
  mir.string_map.StringMap to represent table rows indexed via table column
  title strings.  As element type for instance in
  `Table.withHeaders(headers).rows`

- Use STAX XML parser via calls to parseXML instead of decodeXML

- Remove commented out D code

- Clean up `readSharedEntries`

- Replace row.attributes.filter!(a => a.name == "r").front with row.attributeNamed("r")

- Make `readCells` safe

- Cache parts or whole calculation of `dom` in `readCells`

- In `readCells()`, can we reserve `ret` by looking up dimensions
  somewhere in DOM. Print string passed to DOM and look into it.

- In `readSharedEntries()`, can we reserve `ret` by looking up dimensions
  somewhere in DOM. Print string passed to DOM and look into it.

- Replace calls to array append `~=` with `Appender.put()` taking a range if possible

- Clean up `readCells`

- Clean up `insertValueIntoCell`

- Avoid cast to `immutable` in `convertToString` and return `inout(ubyte)` instead

- Call `assumeUnique` at the end of `readCells` if `Cell.members` are
  `immutable`.

- TODO: 1. contruct this lazily in Sheet upon use and cache

- TODO: 2. deprecate it

- const Relationships* sheetRel = rid in rels; // TODO: move this calculation to caller and pass Relationships as rels

- Replace ret ~= tORr.children[0].text.specialCharacterReplacementReverseLazy.to!string; with
  ret.put(tORr.children[0].text.specialCharacterReplacementReverseLazy)

- Make sure all the calls `expandTrusted`, `convertToString`, `parseDOM` are
  only called once (and cached) for every ArchiveMember by moving them into
  caching members of `File`.

- Move body of `parseRelationships` into `File.parseRelationships` and make the
  module-scope public `parseRelationships` a thin wrapper on top of it. Cache
  the calculation of `parseRelationships` in an AA mapping from `ArchiveMember`
  to `RelationshipsById` if needed and store that AA in private `File` member.

- Clean up `parseRelationships`

- Replace `assert(dom.children.length == 1);` with `enforce(dom.children.length == 1);`
  and similarly for other asserts.

- Qualify some functions in std.zip as pure and then qualify D code in this
  package as pure.  Qualify `std.zip.ZipArchive` members as, at least, `pure`
  and then the functions in there that use them.

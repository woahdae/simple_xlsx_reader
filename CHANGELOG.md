### 5.1.0

* Parse sheets containing namespaces and no 'r' att (@skipchris)
* Fix Zlib error when loading from string (@myabc)
* Prevent a SimpleXlsxReader::CellLoadError (no implicit conversion of Integer
  into String) when the casted value (friendly name) is not a string (@tsdbrown)
* Accidental 25% perfarmance improvement while experimenting with namespace
  support (see #53f5a9).

### 5.0.0

* Change SimpleXlsxReader::Hyperlink to default to the visible cell value
  instead of the hyperlink URL, which in the case of mailto hyperlinks is
  surprising.
* Fix blank content when parsing docs from string (@codemole)

### 4.0.1

* Fix nil error when handling some inline strings

  Inline strings are almost exclusively used by non-Excel XLSX
  implementations, but are valid, and sometimes have nil chunks.

  Also, inline strings weren't preserving whitespace if Nokogiri is
  parsing the string in chunks, as it does when encountering escaped
  characters. Fixed.

### 4.0.0

* Fix percentage rounding errors. Previously we were dividing by 100, when we
  actually don't need to, so percentage types were 100x too small. Fixes #21.
  Major bump because workarounds might have been implemented for previous
  incorrect behavior.
* Fix small oddity in one currency format where round numbers would be cast
  to an integer instead of a float.

### 3.0.1

* Fix parsing "chunky" UTF-8 workbooks. Closes issues #39 and #45. See ce67f0d4.

### 3.0.0

* Change the way we typecast cells in the General format. This probably won't
  break anything in your app, but it's a change in behavior that theoretically
  could.

  Previously, we were treating cells using General the format as strings, when
  according to the Office XML standard, they should be treated as numbers. We
  now attempt to cast such cells as numbers, and fall back to strings if number
  casting fails.

  Thanks @jrodrigosm

### 2.0.1

* Restore ability to parse IO strings (@robbevp)
* Add Ruby 3.1 and 3.2 to CI (@taichi-ishitani)

### 2.0.0

* SPEED
  * Reimplement internals in terms of a SAX parser
  * Change `SimpleXlsxReader::Sheet#rows` to be a `RowsProxy` that streams `#each`
* Convenience - use `rows#each(headers: true)` to get header names while enumerating rows

### 1.0.5

* Support string or io input via `SimpleXlsxReader#parse` (@kalsan, @til)

### 1.0.4

* Fix Windows + RubyZip 1.2.1 bug preventing files from being read
* Add ability to parse hyperlinks
* Support files exported from Google Docs (@Strnadj)

### 1.0.3

Broken on Ruby 1.9; yanked.

### 1.0.2

* Fix Ruby 1.9.3-specific bug preventing parsing most sheets [middagj, eritiro]
* Better support for non-excel-generated xlsx files [bwlang]
  * You don't always have a numFmtId column, and that's OK
  * Sometimes 'sharedStrings.xml' can be 'sharedstrings.xml'
* Fixed parsing times very close to 12/30/1899 [Valeriy Utyaganov]
* Be more flexible with custom formats using a numFmtId < 164

### 1.0.1

* Add support for the 1904 date system [zilverline]

### 1.0.0

No changes since 1.0.0.pre. Releasing 1.0.0 since the project has seen a
few months of stability in terms of bug fix requests, and the API is not
going to change.

### 1.0.0.pre

* Handle files with blank rows [Brian Hoffman]
* Preserve seconds when casting datetimes [Rob Newbould]
* Preserve empty rows (previously would be ommitted)
* Speed up parsing by ~55%

### 0.9.8

* Rubyzip 1.0 compatability

### 0.9.7

* Fix cell parsing where cells have a type, but no content
* Add a speed test; parsing performs in linear time, but a relatively
  slow line :/

### 0.9.6

* Fix worksheet indexes when worksheets have been deleted

### 0.9.5

* Fix inlineStr support (broken by formula support commit)

### 0.9.4

* Formula support. Formulas used to cause things to blow up, now they don't!
* Support number types styled as dates. Previously, the type was honored
  above the style, which is incorrect for dates; date-numbers now parse as
  dates.
* Error-free parsing of empty sheets
* Fix custom styles w/ numFmtId == 164. Custom style types are delineated
  starting *at* numFmtId 164, not greater than 164.

### 0.9.3

* Support 1.8.7 (tests pass). Ongoing support will depend on ease.

### 0.9.2

* Support reading files written by ex. simple_xlsx_writer that don't
  specify sheet dimensions explicitly (which Excel does).

### 0.9.1

* Fixed an important parse bug that ignored empty 'Generic' cells

### 0.9.0

* Initial release. 0.9 version number is meant to reflect the near-stable
  public api, yet still prerelease status of the project.

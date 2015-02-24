### 1.0.2

* Fix Ruby 1.9.3-specific bug preventing parsing most sheets [middagj, eritiro]
* Better support for non-excel-generated xlsx files [bwlang]
  * You don't always have a numFmtId column, and that's OK
  * Sometimes 'sharedStrings.xml' can be 'sharedstrings.xml'
* Fixed parsing times very close to 12/30/1899 [Valeriy Utyaganov]

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

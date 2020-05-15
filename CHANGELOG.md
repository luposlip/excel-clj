# Change Log

## [1.4.0] = 2020-05-15
### Changed
- A slimmed & trimmed down version of the original library from Matthew. Away is the Viz dependency and the potential X11 display issue, but also the ability to create PDF documents from the Excel files.

## [1.3.2] - 2020-04-15
### Fixed
- Bug introduced in v1.3.1 where adjacent cells with width > 1 cause an 
  exception.

## [1.3.1] - 2020-04-05
### Added
- A lower-level, writer style interface for Apache POI.
- [Prototype/brainstorm](src/excel_clj/prototype.clj) of less complicated, 
  pure-data replacement for high-level API in upcoming v2 release.
### Fixed
- Bug (#3) with the way cells were being written via POI that would write cells
  out of order or mix up the style data between cells.

## [1.2.1] - 2020-04-01
### Added
- Can bind a dynamic `*n-threads*` var to set the number of threads used during 
  writing.

## [1.2.0] - 2020-08-13
### Added
- Performance improvements for large worksheets.

## [1.1.2] - 2019-06-04
### Fixed
- If the first level of the tree is a leaf, `accounting-table` doesn't walk it 
  correctly.
### Added
- Can pass through a custom `:min-leaf-depth` key to `tree` (replaces binding a 
dynamic var in earlier versions).

## [1.1.1] - 2019-06-01
### Fixed
- Total rows were not always being displayed correctly for trees

## [1.1.0] - 2019-05-28
### Added
- More flexible tree rendering/aggregation

### Changed
- Replaced lots of redundant tree code with a `walk` function

## [1.0.0] - 2019-04-17
### Added
- PDF generation
- Nicer readme, roadmap & tests

## [0.1.0] - 2019-01-15
- Pulled this code out of an accounting project I was working on as its own library.
- Already had
    - Clojure data wrapper over Apache POI.
    - Tree/table/cell specifications.
    - Excel sheet writing.


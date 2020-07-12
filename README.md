# xlsx to json w/ geocoding

Script to convert xlsx to json and geocode address field to coordinates suitable for MongoDB geospatial queries.

Developed specifically for the needs of [PlayUp](https://playup.coach). Could be adapted for other purposes.

## Dependencies
* [xlrd](https://pypi.org/project/xlrd/)
* [progress](https://pypi.org/project/progress/)

## Usage

* Add Google API Key
* Run with `python xlsxtojson.py`

## Nesting
Nesting is supported for a single level of depth. Concatenating parent and child keys with a `.` separator will result in the value being nested under the parent.

``` javascript
// Cell Heading
parent.child

// JSON Output
parent: {
    child: value
}
```
**Note**: Nesting is currently only supported for 1 level of depth, anything greater will likely lead to unexpected results.

## Arrays
Separating values with `|` will cause them to be output as an array.

``` javascript
// Cell Value
"One|Two|Three"

// JSON Output
["One", "Two", "Three"]
```

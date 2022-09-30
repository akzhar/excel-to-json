# Excel to JSON

## What is it?

Data stored as tables in Excel is converted by a VBA macro to JSON format.

## How to use it?

- The data on Excel worksheets must be stored in accordance with the [rules](#rules)
- Specify 2 options on the `Instruction` worksheet:
  - name of the file eg. `data.json`
  - path to save the file eg. `C:\data\`
- Press `EXCEL → JSON` button and thats is!

<img src="https://raw.githubusercontent.com/akzhar/vba-excel-to-json/main/demo.gif" alt="demo" title="demo" width="100%"/>

<a id="rules"></a>
## Data storage rules

<ol>
  <li><code>ROOT</code> worksheet is the place where JSON creation starts</li>
  <li>Every worksheet contains a table (set of rows), 1 row = 1 object:
    <ul>
      <li>the 1st row of the table (headers) contains <b>keys</b> of objects</li>
      <li>the other rows in the table contains <b>values</b> assosiated with the header (key) from the same column</li>
    </ul>
  </li>
  <li>1 worksheet (table) = 1 object / array of objects
    <ul>
      <li>if there is only 1 row <b>it's an object</b></li>
      <li>if there is > 1 row <b>it's an array of objects</b></li>
    </ul>
  </li>
  <li>Every cells in the tables could contain one of the following:
    <ul>
      <li>Any text value (the JSON format stores all values as strings)</li>
      <li>Array of text values:
        <ul>
          <li>use the square brackets to identify an array <code>[ ... ]</code></li>
          <li>inside the brackets the array items goes separated by comma <code>[ item1 , item2 , item3 ]</code></li>
        </ul>
      </li>
      <li>Another object / array of objects:
        <ul>
          <li>use the curly brackets to identify an object <code>{ ... }</code></li>
          <li>in fact, this is a link to another worksheet (see section 3)</li>
          <li>inside the brackets write the name of the worksheet, which rows will be converted to objects <code>{ worksheet name }</code></li>
        </ul>
      </li>
    </ul>
  </li>
  <li>Limitations:
    <ul>
      <li>put an array of text values inside another array currently <u>is not supported</u> → <code>[ item1 , [ ... ] , item3 ]</code></li>
      <li>it's OK to put an object inside an array → <code>[ { ... } ]</code></li>
    </ul>
  </li>
</ol>
XLSX = require('xlsx');
var fs = require('fs')

// Set these values. 
const INPUT_EXCEL_FILE_PATH = 'test.xlsx';
const EXCEL_SHEET_NAME = 'Sheet1';

const OUTPUT_CSV_FILE_PATH = 'test-out.csv';

// Do not touch this. 
let COLUMN_TYPE_ENUM = Object.freeze(
    {
        NUMBER: 1,
        STRING: 2,
        DEFAULT: 3
    }
);

class JSONColumnFormatter {

    // column_type: Type of column. Set it to one of the values in COLUMN_TYPE_ENUM.

    // format_value: Callback to return formatted value. [format_value (column_value, row_obj) :string].

    constructor (column_type, format_value) {
        this.column_type = column_type;
        this.format_value = format_value;
    }
}

class JSONFormatterOptions {
    // stringDelimiterCharacter: The delimiter to use for formatting strings. 
    // rowPrefix, rowSuffix: Callback with signature. [function (jsonRow): any]

    constructor (stringDelimiterCharacter, rowPrefix, rowSuffix) {
        this.stringDelimiterCharacter = stringDelimiterCharacter ? stringDelimiterCharacter : "'";
        this.rowPrefix = rowPrefix;
        this.rowSuffix = rowSuffix;
    }
}

class JSONFormatter {
    // json_column_formatter_list: column formatters.

    // json: json to format.

    // json_formatter_options: JSONFormatterOptions.

    constructor (json_column_formatter_list, json, json_formatter_options) {
        this.json_column_formatter_list = json_column_formatter_list;
        this.json = json;
        this.json_formatter_options = json_formatter_options;
    }

    // create a copy of json and return a formatted copy based on the column formatters.
    getFormattedJSON () {
        console.dir (this.json);

        let json_formatted = [];

        for (let row of this.json) {
            let row_formatted = {};
            let json_formatter_options = this.json_formatter_options;

            if (json_formatter_options.rowPrefix !== undefined) {
                let prefix = json_formatter_options.rowPrefix (row);
             
                row_formatted [prefix.key] = prefix.value;
            }

            for (let column_key of Object.keys(row)) {
                let column = row [column_key];
                
                if (this.json_column_formatter_list[column_key] !== undefined) {
                    let json_column_formatter = this.json_column_formatter_list[column_key];

                    let column_formatted = column;

                    if (json_column_formatter.format_value !== undefined) {
                        column_formatted = json_column_formatter.format_value (column, row);
                    }

                    if (json_column_formatter.column_type === COLUMN_TYPE_ENUM.STRING) {
                        column_formatted = json_formatter_options.stringDelimiterCharacter + column + json_formatter_options.stringDelimiterCharacter;
                    }

                    row_formatted[column_key] = column_formatted;
                }
            }

            if (json_formatter_options.rowSuffix !== undefined) {
                let suffix = json_formatter_options.rowSuffix (row);
             
                row_formatted [suffix.key] = suffix.value;
            }

            json_formatted.push (row_formatted);
        }

        return json_formatted;
    }
}

let workbook = XLSX.readFile(INPUT_EXCEL_FILE_PATH);
let excel_sheet = workbook.Sheets[EXCEL_SHEET_NAME];

let excel_json = XLSX.utils.sheet_to_json (excel_sheet);

let json_column_formatter_string = new JSONColumnFormatter (COLUMN_TYPE_ENUM.STRING);
let json_column_formatter_default = new JSONColumnFormatter (COLUMN_TYPE_ENUM.DEFAULT);

let json_formatter_options = new JSONFormatterOptions ("'", 
    function (row) {
        return { key: 'prefix', value: 'VALUES ( ' } ;
    }
    ,function (row) {
        return { key: 'suffix', value: ' )' } ;
    }
);

let json_formatter = new JSONFormatter (
    {
        'Compound_Name': json_column_formatter_string,
        'Product_Line': json_column_formatter_string,
        'Class': json_column_formatter_string,
        'Status': json_column_formatter_string,
        'Facility': json_column_formatter_string,
        'SAP_Number': json_column_formatter_string,
        'Ingredient': json_column_formatter_string,
        'Pass': json_column_formatter_default,
        'phr': json_column_formatter_default
    },
    excel_json,
    json_formatter_options
);

json_formatted = json_formatter.getFormattedJSON ();

let input_stream = XLSX.stream.to_csv(XLSX.utils.json_to_sheet(json_formatted), { strip: true });
input_stream.pipe(fs.createWriteStream(OUTPUT_CSV_FILE_PATH));

console.log('------ Formatted JSON ----------');

console.dir (json_formatted);

console.log('------ End ----------');
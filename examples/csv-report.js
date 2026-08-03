// csv-report.js - read, filter and write CSV with libs/csv.js, with logging
//
// Tabular data used to mean driving Excel over COM (do_in_excel +
// read_sheet_data), which needs Excel installed and starts a COM server just to
// read a few rows. libs/csv.js does it with nothing but the file system.
//
// This example writes a small sales CSV to the temp folder, reads it back as
// objects, filters and totals it, then writes a summary CSV - and uses
// libs/log.js to show levels and a log file along the way.
//
// Run via:
//   examples\run.bat csv-report.js
//   cscript.exe bin\launcher.js examples\csv-report.js
//   cscript.exe dist\launcher.js examples\csv-report.js

load("core");
load("polyfills");
load("log");
load("system");
load("csv");

// ---------------------------------------------------------------------------
// Logging setup
// ---------------------------------------------------------------------------

var fso     = new ActiveXObject("Scripting.FileSystemObject");
var tempDir = fso.GetSpecialFolder(2).Path;

var log_path = tempDir + "\\jsw_csv_report.log";
log_to_file(log_path);          // every line below also lands in this file
set_log_level("debug");         // show debug lines too; the default hides them

log_info("csv-report starting");
log_debug("temp folder is " + tempDir);

// ---------------------------------------------------------------------------
// 1. Write a source CSV
// ---------------------------------------------------------------------------

var source_path = tempDir + "\\jsw_sales.csv";

var sales = [
    { region: "North", product: "widget, large", units: "120", revenue: "2400" },
    { region: "North", product: "widget",        units: "80",  revenue: "800"  },
    { region: "South", product: 'the "big" one', units: "15",  revenue: "1500" },
    { region: "South", product: "widget",        units: "60",  revenue: "600"  },
    { region: "East",  product: "gadget",        units: "5",   revenue: "250"  }
];

// The header row comes from the object keys. Note the product names: one has a
// comma, another has quotes - csv_format quotes and escapes both, so they
// survive the round trip untouched.
write_csv_file(source_path, sales);
log_info("wrote " + sales.length + " rows to " + source_path);

// ---------------------------------------------------------------------------
// 2. Read it back as objects
// ---------------------------------------------------------------------------

var rows = read_csv_file(source_path, { headers: true });
log_info("read back " + rows.length + " rows");

log("");
log("all rows:");
for (var i = 0; i < rows.length; i++) {
    log("  " + rows[i].region + " | " + rows[i].product +
        " | units=" + rows[i].units + " | revenue=" + rows[i].revenue);
}

// ---------------------------------------------------------------------------
// 3. Filter and total
// ---------------------------------------------------------------------------

// Every CSV value is a string - there is no type guessing - so convert before
// doing arithmetic.
var big_sellers = rows.filter(function(r) { return parseInt(r.units, 10) >= 50; });
log_info("found " + big_sellers.length + " rows with 50+ units");

var totals = {};
for (var j = 0; j < rows.length; j++) {
    var region = rows[j].region;
    if (!totals[region]) { totals[region] = { region: region, units: 0, revenue: 0 }; }
    totals[region].units   += parseInt(rows[j].units, 10);
    totals[region].revenue += parseInt(rows[j].revenue, 10);
}

var summary = Object.values(totals);
log_debug("summary: " + JSON.stringify(summary));

// ---------------------------------------------------------------------------
// 4. Write the summary out
// ---------------------------------------------------------------------------

var summary_path = tempDir + "\\jsw_sales_summary.csv";

// An explicit header array fixes the column order (object key order is not
// something to rely on) and can also drop columns.
write_csv_file(summary_path, summary, { headers: ["region", "units", "revenue"] });
log_info("wrote summary to " + summary_path);

log("");
log("summary file contents:");
log(read_text_file(summary_path));

// ---------------------------------------------------------------------------
// 5. Levels
// ---------------------------------------------------------------------------

log("");
set_log_level("warn");
log_debug("this debug line is filtered out");
log_info("so is this info line");
log_warn("warnings still get through");
set_log_level("info");

log("");
log_info("log file written to " + log_path);
log_to_file(null);

// ---------------------------------------------------------------------------
// Cleanup
// ---------------------------------------------------------------------------

delete_file(source_path);
delete_file(summary_path);
delete_file(log_path);

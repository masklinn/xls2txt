#![deny(clippy::all)]
#![warn(clippy::pedantic)]
#![allow(clippy::similar_names)]

use calamine::{open_workbook_auto, DataType, Reader};
use std::error::Error;
use std::fmt::{self, Display, Debug, Formatter, Write};
use std::io;
use clap::{App, Arg, crate_name, crate_version};
use std::convert::TryFrom;

pub enum Errors {
    InvalidSeparator,
    Empty, NotFound(String),
    Csv(csv::Error),
    Spreadsheet(calamine::Error),
    CellError(calamine::CellErrorType),
}

impl Error for Errors {}
// delegate to Display so the error message is not crap
impl Debug for Errors {
    fn fmt(&self, fmt: &mut Formatter) -> fmt::Result {
        (&self as &dyn Display).fmt(fmt)
    }
}
impl Display for Errors {
    fn fmt(&self, fmt: &mut Formatter) -> fmt::Result {
        use Errors::*;
        match self {
            Empty => write!(fmt, "Empty spreadsheet"),
            NotFound(ref s) => write!(fmt, "Could not find sheet \"{}\" in spreadsheet", s),
            InvalidSeparator => write!(fmt, "A provided separator is invalid, separators need to be a single ascii chacter"),
            Csv(ref e) => write!(fmt, "{}", e),
            Spreadsheet(ref e) => write!(fmt, "{}", e),
            CellError(ref e) => write!(fmt, "Error found in cell ({:?})", e)
        }
    }
}
impl From<csv::Error> for Errors {
    fn from(err: csv::Error) -> Self {
        Self::Csv(err)
    }
}
impl From<calamine::Error> for Errors {
    fn from(err: calamine::Error) -> Self {
        Self::Spreadsheet(err)
    }
}

fn separator_to_byte(s: &str) -> Result<u8, Errors> {
    s.chars().next()
        .ok_or(Errors::InvalidSeparator)
        .and_then(|c| u8::try_from(c as u32).map_err(|_| Errors::InvalidSeparator))
}
pub fn run(n: &'static str, default_rs: &'static str, default_fs: &'static str) -> Result<(), Errors> {

    let matches = App::new(crate_name!())
        .version(crate_version!())
        .about("Converts spreadsheets to text")
        .long_about(&format!("Converts the first sheet of the spreadsheet at PATH (or <sheet> if \
requested) to {} sent to stdout.

Should be able to convert from (and automatically guess between) XLS, XLSX, XLSB and ODS.", n) as &str)
        .arg(Arg::with_name("PATH").help("Spreadsheet file path").required(true))
        .arg(Arg::with_name("sheet")
             .short("s").long("sheet")
             .default_value("1")
             .help("Name or index (1 is first) of the sheet to convert")
        )
        .arg(Arg::with_name("RS")
             .short("r").long("rs").long("record-separator")
             .takes_value(true)
             .help("Record separator (a single character)")
        )
        .arg(Arg::with_name("FS")
             .short("f").long("fs").long("field-separator")
             .takes_value(true)
             .help("Field separator (a single character)")
        )
        .get_matches();

    let rs = separator_to_byte(matches.value_of("RS").unwrap_or(default_rs))?;
    let fs = separator_to_byte(matches.value_of("FS").unwrap_or(default_fs))?;

    let mut workbook = open_workbook_auto(matches.value_of("PATH").unwrap())?;

    let sheet = matches.value_of("sheet").unwrap_or("1");
    // if sheet is a number get corresponding sheet in list
    let name = String::from(
        sheet.parse::<usize>().ok()
            .and_then(|n| workbook.sheet_names().get(n.saturating_sub(1)))
            .map_or(sheet, |s| s as &str)
    );

    let range = if let Some(Ok(r)) = workbook.worksheet_range(&name) {
        r
    } else {
        return Err(Errors::NotFound(name));
    };

    let stdout = io::stdout();
    let mut out = csv::WriterBuilder::new()
        .terminator(csv::Terminator::Any(rs))
        .delimiter(fs)
        .from_writer(stdout.lock());

    let mut contents = vec![String::new();range.width()];
    for row in range.rows() {
        for (c, cell) in row.iter().zip(contents.iter_mut()) {
            cell.clear();
            match *c {
                DataType::Error(ref e) => return Err(Errors::CellError(e.clone())),
                // don't go through fmt for strings
                DataType::String(ref s) => cell.push_str(s),
                ref rest => write!(cell, "{}", rest).expect("formatting basic types to a string should never fail"),
            };
        }
        out.write_record(&contents)?;
    }

    Ok(())
}

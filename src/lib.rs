#![deny(clippy::all)]
#![allow(clippy::similar_names)]

use calamine::{open_workbook_auto, DataType, Reader};
use clap::{builder::Arg, command};
use guard::guard;
use std::error::Error;
use std::fmt::{self, Debug, Display, Formatter, Write};
use std::io;

pub enum Errors {
    InvalidSeparator,
    MissingSeparator,
    Empty,
    NotFound(String),
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
            NotFound(s) => write!(fmt, "Could not find sheet {:?} in spreadsheet", s),
            InvalidSeparator => write!(
                fmt,
                "A provided separator is invalid, separators need to be a single ascii chacter"
            ),
            MissingSeparator => write!(fmt, "No separator found"),
            Csv(e) => write!(fmt, "{}", e),
            Spreadsheet(e) => write!(fmt, "{}", e),
            CellError(e) => write!(fmt, "Error found in cell ({:?})", e),
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
    let c = s.chars().next().ok_or(Errors::InvalidSeparator)?;
    (c as u32).try_into().map_err(|_| Errors::InvalidSeparator)
}

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
enum FormulaMode {
    CachedValue,
    IfEmpty,
    Always,
    // Evaluate
}
impl clap::ValueEnum for FormulaMode {
    fn value_variants<'a>() -> &'a [Self] {
        &[Self::CachedValue, Self::IfEmpty, Self::Always]
    }

    fn to_possible_value<'a>(&self) -> Option<clap::builder::PossibleValue> {
        match self {
            Self::CachedValue => Some(clap::builder::PossibleValue::new("cached-value")),
            Self::IfEmpty => Some(clap::builder::PossibleValue::new("if-empty")),
            Self::Always => Some(clap::builder::PossibleValue::new("always")),
        }
    }
}

pub fn run(
    n: &'static str,
    default_rs: &'static str,
    default_fs: &'static str,
) -> Result<(), Errors> {
    let help = format!(
        "\
Converts the first sheet of the spreadsheet at PATH (or <sheet> if \
requested) to {n} sent to stdout.

Should be able to convert from (and automatically guess between) XLS, XLSX, XLSB and ODS."
    );
    let matches = command!()
        .about("Converts spreadsheets to text")
        .long_about(&help)
        .arg(
            Arg::new("PATH")
                .help("Spreadsheet file path")
                .required(true),
        )
        .arg(
            Arg::new("sheet")
                .short('s')
                .long("sheet")
                .default_value("1")
                .help("Name or index (1 is first) of the sheet to convert"),
        )
        .arg(
            Arg::new("RS")
                .short('r')
                .long("rs")
                .long("record-separator")
                .default_value(default_rs)
                .help("Record separator (a single character)"),
        )
        .arg(
            Arg::new("FS")
                .short('f')
                .long("fs")
                .long("field-separator")
                .default_value(default_fs)
                .help("Field separator (a single character)"),
        )
        .arg(
            Arg::new("formula-mode")
                .long("formula")
                .value_parser(clap::builder::EnumValueParser::<FormulaMode>::new())
                .default_value("cached-value")
                .help(
                    "\
Whether and when to show formulas

- cached-value: never show formula, always display cached value, even if empty
- if-empty: show formula if cached-value is empty or absent
- always: always show formula for formula cells (ignore cached value)\
",
                ),
        )
        .get_matches();

    let rs = separator_to_byte(
        matches
            .get_one::<String>("RS")
            .ok_or(Errors::MissingSeparator)?,
    )?;
    let fs = separator_to_byte(
        matches
            .get_one::<String>("FS")
            .ok_or(Errors::MissingSeparator)?,
    )?;

    let mut workbook = open_workbook_auto(matches.get_one::<String>("PATH").unwrap())?;

    let sheet = matches.get_one::<String>("sheet").map_or("1", |s| &**s);
    // if sheet is a number get corresponding sheet in list
    let name = String::from(
        sheet
            .parse::<usize>()
            .ok()
            .and_then(|n| workbook.sheet_names().get(n.saturating_sub(1)))
            .map_or(sheet, |s| s as &str),
    );

    guard!(let Some(Ok(range)) = workbook.worksheet_range(&name) else {
        return Err(Errors::NotFound(name));
    });
    guard!(let Some((offset_j, offset_i)) = range.start() else {
        return Ok(());
    });

    let formula_mode = *matches.get_one::<FormulaMode>("formula-mode").unwrap();
    let wb = workbook
        .worksheet_formula(&name)
        .expect("we know the sheet exists");
    let formatter: Box<dyn Fn(u32, u32, DataType) -> DataType> = match wb.as_ref() {
        Ok(f) => match formula_mode {
            FormulaMode::CachedValue => Box::new(|_, _, cell| cell),
            FormulaMode::IfEmpty => Box::new(|i, j, cell| {
                let formula = f.get_value((j, i)).filter(|s| !s.is_empty());
                match cell {
                    DataType::Empty => {
                        formula.map_or(DataType::Empty, |v| DataType::String(v.to_string()))
                    }

                    DataType::String(s) if s.is_empty() => {
                        formula.map_or(DataType::Empty, |v| DataType::String(v.to_string()))
                    }

                    rest => rest,
                }
            }),
            FormulaMode::Always => Box::new(|j, i, cell| {
                f.get_value((i, j))
                    .filter(|s| !s.is_empty())
                    .map_or(cell, |s| DataType::String(s.to_string()))
            }),
        },
        Err(e) => {
            if formula_mode != FormulaMode::CachedValue {
                eprintln!("Formula parsing error: {e:?}");
            }
            Box::new(|_, _, cell| cell)
        }
    };

    let stdout = io::stdout();
    let mut out = csv::WriterBuilder::new()
        .terminator(csv::Terminator::Any(rs))
        .delimiter(fs)
        .from_writer(stdout.lock());

    let mut contents = vec![String::new(); range.width()];
    for (j, row) in range.rows().enumerate() {
        for (i, (c, cell)) in row.iter().zip(contents.iter_mut()).enumerate() {
            cell.clear();
            match formatter(i as u32 + offset_i, j as u32 + offset_j, c.clone()) {
                DataType::Error(e) => return Err(Errors::CellError(e)),
                // don't bother updating cell for empty
                DataType::Empty => (),
                // don't go through fmt for strings
                DataType::String(s) => cell.push_str(&s),
                rest => write!(cell, "{}", rest)
                    .expect("formatting basic types to a string should never fail"),
            };
        }
        out.write_record(&contents)?;
    }
    out.flush().unwrap();

    Ok(())
}

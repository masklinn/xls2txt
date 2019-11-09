use xls2txt;

fn main() -> Result<(), xls2txt::Errors> {
    xls2txt::run("comma-separated (by default) text", "\n", ",")
}

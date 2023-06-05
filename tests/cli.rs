use assert_cmd::prelude::*;
use predicates::prelude::*;
use std::error::Error;

use std::process::Command;

#[test]
fn write_csv_default() -> Result<(), Box<dyn Error>> {
    let mut cmd = Command::cargo_bin("xls2csv")?;
    cmd.arg("tests/input.xlsx");

    cmd.assert()
        .success()
        .stderr(predicate::eq(""))
        .stdout(predicate::eq(
            "\
Title,Year,Length,Label
Static Anonimity,2001,45082.73888888889,Restless Records
\"Old World Underground, Where Are You Now\",2003,0.02585648148148148,\"Last Gang Records, Everloving Records\"
Live It Out,2005,0.02840277777777778,Last Gang Records
Grow Up And Blow Away,2007,0.02719907407407407,Last Gang Records
Fantasies,2009,0.02951388888888889,Metric Music International
Synthetica,2012,0.02998842592592593,Metric Music International
Pagans in Vegas,2015,0.03428240740740741,Metric Music International
Art of Doubt,2018,0.04046296296296296,Metric Music International
Formentera,2022,0.03309027777777778,Metric Music International
",
        ));

    Ok(())
}

#[test]
fn write_txt_default() -> Result<(), Box<dyn Error>> {
    let mut cmd = Command::cargo_bin("xls2txt")?;
    cmd.arg("tests/input.xlsx");

    cmd.assert()
        .success()
        .stderr(predicate::eq(""))
        .stdout(predicate::eq("\
Title\tYear\tLength\tLabel
Static Anonimity\t2001\t45082.73888888889\tRestless Records
Old World Underground, Where Are You Now\t2003\t0.02585648148148148\tLast Gang Records, Everloving Records
Live It Out\t2005\t0.02840277777777778\tLast Gang Records
Grow Up And Blow Away\t2007\t0.02719907407407407\tLast Gang Records
Fantasies\t2009\t0.02951388888888889\tMetric Music International
Synthetica\t2012\t0.02998842592592593\tMetric Music International
Pagans in Vegas\t2015\t0.03428240740740741\tMetric Music International
Art of Doubt\t2018\t0.04046296296296296\tMetric Music International
Formentera\t2022\t0.03309027777777778\tMetric Music International
"));

    Ok(())
}

#[test]
fn write_txt_custom() -> Result<(), Box<dyn Error>> {
    let mut cmd = Command::cargo_bin("xls2txt")?;
    cmd.args(["-r", "\n"]).args(["-f", "\t"]);
    cmd.arg("tests/input.xlsx");

    cmd.assert()
        .success()
        .stderr(predicate::eq(""))
        .stdout(predicate::eq("\
Title\tYear\tLength\tLabel
Static Anonimity\t2001\t45082.73888888889\tRestless Records
Old World Underground, Where Are You Now\t2003\t0.02585648148148148\tLast Gang Records, Everloving Records
Live It Out\t2005\t0.02840277777777778\tLast Gang Records
Grow Up And Blow Away\t2007\t0.02719907407407407\tLast Gang Records
Fantasies\t2009\t0.02951388888888889\tMetric Music International
Synthetica\t2012\t0.02998842592592593\tMetric Music International
Pagans in Vegas\t2015\t0.03428240740740741\tMetric Music International
Art of Doubt\t2018\t0.04046296296296296\tMetric Music International
Formentera\t2022\t0.03309027777777778\tMetric Music International
"));

    Ok(())
}

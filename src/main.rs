use clap::Clap;
use money_manager::{GroupBy, parse_report, draw, MyCustomError};

#[derive(Clap, Debug)]
#[clap(name = "money_manager")]
struct Args {
    #[clap(short, long)]
    file: String,

    #[clap(short, long, default_value = "month")]
    group_by: String,
}

fn group_by(group: String) -> GroupBy {
    match String::from(group).as_str() {
        "year" => GroupBy::Year,
        "quarter" => GroupBy::Quarter,
        _ => GroupBy::Month,
    }
}

fn draw_images(file: String, group: String) -> Result<String, MyCustomError> {
    let data = parse_report(file, group_by(group))?;

    draw(data);
    Ok(String::from(""))
}

fn main() {
    let args = Args::parse();

    let res = draw_images(args.file, args.group_by);
    println!("res = {:?}", res);
}

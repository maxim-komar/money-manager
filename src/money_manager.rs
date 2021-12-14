use calamine::{open_workbook, DataType, Range, Reader, Xlsx, XlsxError};
use chrono::{Datelike, NaiveDate};
use plotly::common::{DashType, Line, Mode, Title};
use plotly::{ImageFormat, Layout, Plot, Scatter};
use rand::Rng;
use statistical::{mean, median};
use std::fmt;
use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

const MAX_PERIODS: usize = 12;
const FILENAME_LEN: usize = 16;

#[derive(Debug)]
pub enum MyCustomError {
    OpenError,
    OtherError,
}

impl From<XlsxError> for MyCustomError {
    fn from(_: XlsxError) -> Self {
        MyCustomError::OpenError
    }
}

impl fmt::Display for MyCustomError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            MyCustomError::OpenError => write!(f, "Can't open file"),
            MyCustomError::OtherError => write!(f, "Other error"),
        }
    }
}

#[derive(Debug)]
struct Columns {
    period: usize,
    category: usize,
    tx_type: usize,
    value: usize,
}

#[derive(Debug)]
enum TxType {
    Income,
    Outcome,
}

#[derive(Debug)]
struct Fields {
    period: NaiveDate,
    category: String,
    tx_type: TxType,
    value: f64,
}

pub enum GroupBy {
    Month,
    Quarter,
    Year,
}

type Period = String;
type Category = String;

fn by_month(date: NaiveDate) -> Period {
    format!("{:04}-{:02}", date.year(), date.month())
}

fn by_quarter(date: NaiveDate) -> Period {
    format!("{:04}-q{}", date.year(), (date.month() - 1) / 3 + 1)
}

fn by_year(date: NaiveDate) -> Period {
    format!("{:04}", date.year())
}

fn period_from_date(group_by: GroupBy) -> fn(NaiveDate) -> String {
    match group_by {
        GroupBy::Year => |date| by_year(date),
        GroupBy::Quarter => |date| by_quarter(date),
        GroupBy::Month => |date| by_month(date),
    }
}

fn read_row(columns: &Columns, row: &[DataType]) -> Result<Fields, String> {
    let mut period = None;
    if let DataType::String(s) = &row[columns.period] {
        if let Ok(date) = NaiveDate::parse_from_str(&s, "%d.%m.%Y") {
            period = Some(date);
        }
    }
    if period == None {
        return Err(format!(
            "Can't read period from {:?}",
            row[columns.period]
        ));
    }

    let mut category = None;
    if let DataType::String(s) = &row[columns.category] {
        category = Some(s);
    }
    if category == None {
        return Err(format!(
            "Can't read category from {:?}",
            row[columns.category]
        ));
    }

    let income = String::from("Доход");
    let outcome = String::from("Расход");
    let mut tx_type = None;
    if let DataType::String(s) = &row[columns.tx_type] {
        if *s == income {
            tx_type = Some(TxType::Income);
        } else if *s == outcome {
            tx_type = Some(TxType::Outcome);
        } else {
            return Err(format!(
                "Can't read transaction type from {:?}",
                row[columns.tx_type]
            ));
        }
    }

    let mut value = None;
    if let DataType::Float(f) = &row[columns.value] {
        value = Some(f);
    }
    if value == None {
        return Err(format!(
            "Can't read value from {:?}",
            row[columns.value]
        ));
    }

    Ok(Fields {
        period: period.unwrap(),
        category: category.unwrap().to_string(),
        tx_type: tx_type.unwrap(),
        value: *value.unwrap(),
    })
}

type WorksheetData = BTreeMap<Category, BTreeMap<Period, f64>>;

fn read_worksheet(
    name: String,
    range: Range<DataType>,
    group_by: fn(NaiveDate) -> Period,
) -> Result<WorksheetData, MyCustomError> {
    let period_str = "Период";
    let category_str = "Категория";
    let tx_type_str = "Доход/Расход";
    let value_str = "RUB";

    let period_dt = DataType::String(String::from(period_str));
    let category_dt = DataType::String(String::from(category_str));
    let tx_type_dt = DataType::String(String::from(tx_type_str));
    let value_dt = DataType::String(String::from(value_str));

    let mut period_pos = None;
    let mut category_pos = None;
    let mut tx_type_pos = None;
    let mut value_pos = None;

    if let Some(first_row) = range.rows().next() {
        for i in 0..first_row.len() - 1 {
            if first_row[i] == period_dt {
                period_pos = Some(i);
            } else if first_row[i] == category_dt {
                category_pos = Some(i);
            } else if first_row[i] == tx_type_dt {
                tx_type_pos = Some(i);
            } else if first_row[i] == value_dt {
                value_pos = Some(i);
            }
        }
    } else {
        //return Err(format!("Can't read first row from sheet '{}'", name));
        return Err(MyCustomError::OtherError)
    }

    if period_pos == None {
        return Err(MyCustomError::OtherError)
//        return Err(format!(
//            "Can't find column '{}' in sheet '{}'",
//            period_str, name
//        ));
    }
    if category_pos == None {
        return Err(MyCustomError::OtherError)
//        return Err(format!(
//            "Can't find column '{}' in sheet '{}'",
//            category_str, name
//        ));
    }
    if tx_type_pos == None {
        return Err(MyCustomError::OtherError)
//        return Err(format!(
//            "Can't find column '{}' in sheet '{}'",
//            tx_type_str, name
//        ));
    }
    if value_pos == None {
        return Err(MyCustomError::OtherError)
//        return Err(format!(
//            "Can't fund column '{}' in sheet '{}'",
//            value_str, name
//        ));
    }

    let columns = Columns {
        period: period_pos.unwrap(),
        category: category_pos.unwrap(),
        tx_type: tx_type_pos.unwrap(),
        value: value_pos.unwrap(),
    };

    let mut by_category: BTreeMap<Category, BTreeMap<Period, f64>> = BTreeMap::new();

    for row in range.rows() {
        if let Ok(fields) = read_row(&columns, row) {
            let period = group_by(fields.period);

            let addition = match fields.tx_type {
                TxType::Outcome => fields.value,
                TxType::Income => -fields.value,
            };

            *by_category
                .entry(fields.category)
                .or_insert(BTreeMap::new())
                .entry(period)
                .or_insert(0.0) += addition;
        }
    }

    Ok(by_category)
}

fn last_n_groups(periods: Vec<Period>, n: usize) -> Vec<Period> {
    periods
        .into_iter()
        .rev()
        .take(n + 1)
        .skip(1)
        .rev()
        .collect()
}

/*
#[derive(Debug, Clone)]
struct ChartData {
    name: String,
    y_values: Vec<f64>,
    line: Line,
}

#[derive(Debug, Clone)]
struct ImageData {
    x_values: Vec<Period>,
    charts: Vec<ChartData>,
}

fn worksheet_data_to_image_data(worksheet_data: WorksheetData, periods: Vec<Period>) -> ImageData {
    let mut chart_data = Vec::new();

    for (category, by_period) in &worksheet_data {
        let y_values: Vec<_> = periods
            .iter()
            .map(|period| match by_period.get(period) {
                Some(v) => *v,
                None => 0.0,
            })
            .collect();

        chart_data.push(ChartData {
            name: category.to_string(),
            y_values: y_values,
            line: Line::new(),
        })
    }

    ImageData {
        x_values: periods,
        charts: chart_data,
    }
}

fn filter_all_spendings(data: Vec<ChartData>) -> Vec<ChartData> {
    data.into_iter()
        .filter(|cd| median(&cd.y_values) > 0.0)
        .collect()
}

fn filter_regular_spendings(data: Vec<ChartData>) -> Vec<ChartData> {
    data.into_iter()
        .filter(|cd| {
            let mean = mean(&cd.y_values);
            let median = median(&cd.y_values);
            (median > 0.0) && (mean < 2.0 * median) && (median < 2.0 * mean)
        })
        .collect()
}

*/

/*
fn add_total(mut image_data: ImageData) -> ImageData {
    let len = image_data.x_values.len();

    let mut total: Vec<f64> = vec![0.0; len];
    for chart in &image_data.charts {
        for i in 0..len {
            total[i] += chart.y_values[i];
        }
    }

    image_data.charts.push(ChartData {
        name: String::from("Всего"),
        y_values: total,
        line: Line::new().dash(DashType::LongDashDot),
    });

    image_data
}
*/

fn fix_label(s: &String) -> String {
    s.replace(" ", "&nbsp;")
}

fn generate_random_filename(len: usize) -> String {
    const CHARSET: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let mut rng = rand::thread_rng();
    (0..len)
        .map(|_| {
            let idx = rng.gen_range(0..CHARSET.len());
            CHARSET[idx] as char
        })
        .collect()
}

/*
fn draw_image(
    image_data: &ImageData,
    title: String,
    filter: fn(Vec<ChartData>) -> Vec<ChartData>,
) -> Result<PathBuf, String> {
    let filtered_image_data = ImageData {
        x_values: image_data.x_values.clone(),
        charts: filter(image_data.charts.clone()),
    };
    let filtered_image_data = add_total(filtered_image_data);

    let mut plot = Plot::new();
    let layout = Layout::new().title(Title::new(&fix_label(&title)));
    plot.set_layout(layout);

    for data in filtered_image_data.charts {
        let mean = mean(&data.y_values);
        let label = format!("{} (avg: {}k)", data.name, (mean as i32) / 1000);
        plot.add_trace(
            Scatter::new(image_data.x_values.to_owned(), data.y_values.to_owned())
                .name(&fix_label(&label))
                .mode(Mode::LinesMarkers)
                .line(data.line),
        );
    }

    
    let mut filename = generate_random_filename(FILENAME_LEN);
    filename.push_str(".svg");
    let path = Path::new("/tmp").join(filename);

    //plot.save(&path, ImageFormat::SVG, 1400, 740, 1.0);
    plot.show();
    Ok(path)
}
*/

/*
pub fn parse_report_and_draw_images(file: String, group: GroupBy) -> Result<Vec<PathBuf>, String> {
    let group_by = period_from_date(group);
    let maybe_workbook: Result<Xlsx<_>, XlsxError> = open_workbook(file);
    match maybe_workbook {
        Err(e) => Err(e.to_string()),
        Ok(mut workbook) => {
            let mut result: Vec<Result<PathBuf, String>> = Vec::new();

            for (worksheet_name, range) in workbook.worksheets() {
                if let Ok(worksheet_data)) =
                    read_worksheet(worksheet_name, range, group_by)
                {
                    let x_values: Vec<_> = last_n_groups(periods, MAX_PERIODS);
                    let image_data = worksheet_data_to_image_data(worksheet_data, x_values);

                    result.push(draw_image(
                        &image_data,
                        String::from("Регулярные траты"),
                        filter_regular_spendings,
                    ));
                    result.push(draw_image(
                        &image_data,
                        String::from("Все траты"),
                        filter_all_spendings,
                    ));
                }
            }

            let (results, errors): (Vec<_>, Vec<_>) = result.into_iter().partition(|r| r.is_ok());

            let errors: Vec<_> = errors.into_iter().map(Result::unwrap_err).collect();

            if results.len() == 0 {
                Err(String::from("Not found sheets with report"))
            } else if errors.len() > 0 {
                Err(errors.join(";\n"))
            } else {
                Ok(results.into_iter().map(Result::unwrap).collect())
            }
        }
    }
}
*/

pub fn parse_report(file: String, group_by: GroupBy) -> Result<Vec<WorksheetData>, MyCustomError> {
    let mut workbook: Xlsx<_> = open_workbook(file)?;
    let group_by_fn = period_from_date(group_by);
        
    workbook.worksheets()
        .into_iter()
        .map(|(name, range)| read_worksheet(name, range, group_by_fn))
        .collect()
}

fn worksheet_data_to_periods(data: &Vec<WorksheetData>) -> Vec<Period> {
    let mut periods: BTreeSet<Period> = BTreeSet::new();

    for worksheet_data in data {
        for (_cat, by_cat) in worksheet_data {
            for (period, _by_period) in by_cat {
                periods.insert(period.clone());
            }
        }
    }

    periods.into_iter().collect()
}

fn y(by_cat: &BTreeMap<Period, f64>, periods: &Vec<Period>) -> Vec<f64> {
    periods
        .iter()
        .map(|p| by_cat.get(p).unwrap_or(&0.0).clone())
        .collect()
}

fn is_spending_category(y_values: &Vec<f64>) -> bool {
    median(y_values) > 0.0
}

fn plot(title: String, worksheet_data: &WorksheetData, periods: &Vec<Period>) -> Plot {
    let mut plot = Plot::new();
    plot.set_layout(
        Layout::new()
            .title(Title::new(&fix_label(&title)))
    );

    let mut y_total : Vec<f64> = Vec::new();
    for _ in periods.iter() {
        y_total.push(0.0);
    }

    for (cat, by_cat) in worksheet_data {
        let y_values = y(by_cat, &periods);
        if is_spending_category(&y_values) {
            for it in y_values.iter().zip(y_total.iter_mut()) {
                let (v, t) = it;
                *t += *v;
            }
            let label = format!("{} (avg: {}k)", cat, (mean(&y_values) as i32) / 1000);
            plot.add_trace(
                Scatter::new(periods.to_owned(), y_values.to_owned())
                    .name(&fix_label(&label))
                    .mode(Mode::LinesMarkers)
                    .line(Line::new()),
            );
        }
    }

    let label = format!("Всего (avg: {}k)", (mean(&y_total) as i32) / 1000);
    plot.add_trace(
        Scatter::new(periods.to_owned(), y_total.to_owned())
            .name(&fix_label(&label))
            .mode(Mode::LinesMarkers)
            .line(Line::new()),
    );

    plot
}

pub fn draw(data: Vec<WorksheetData>) {
    let periods = last_n_groups(worksheet_data_to_periods(&data), MAX_PERIODS);

    for worksheet_data in data {
        let title = String::from("Все траты");
        let plot = plot(title, &worksheet_data, &periods);
        plot.show();
    }
}

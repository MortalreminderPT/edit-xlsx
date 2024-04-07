# edit-xlsx

[<img alt="github" src="https://img.shields.io/badge/github-MortalreminderPT/edit--xlsx-8da0cb?style=for-the-badge&labelColor=555555&logo=github" height="20">](https://github.com/MortalreminderPT/edit-xlsx)
[<img alt="crates.io" src="https://img.shields.io/crates/v/edit-xlsx.svg?style=for-the-badge&color=fc8d62&logo=rust" height="20">](https://crates.io/crates/edit-xlsx)
[<img alt="docs.rs" src="https://img.shields.io/badge/docs.rs-edit--xlsx-66c2a5?style=for-the-badge&labelColor=555555&logo=docs.rs" height="20">](https://docs.rs/edit-xlsx)

Welcome to Edit-xlsx, your quick and easy-to-use Rust library for Excel file editing. Whether you're a developer working on a project that involves manipulating Excel files or a business user streamlining data workflows, Edit-xlsx is designed to make Excel editing a breeze.

## Features

- **Formula Editing:** Easily manipulate and customize formulas in Excel sheets.
- **Cell Text Editing:** Edit the content of individual cells, including inserting, modifying, or deleting text.
- **Background Setting:** Set and adjust background colors for cells or ranges.
- **Image Insertion:** Seamlessly insert images into your Excel files.
- **Format Setting:** Apply various formatting options to cells, such as font styles, sizes, and text alignments.
- **Cell Merging:** Merge cells to create visually appealing layouts.
- **Worksheet Editing:** Edit and manage worksheets with ease.
- **Pane Manipulation:** Control and customize panes for a better viewing experience.

## Getting Started

Getting started with Edit-xlsx is a straightforward process. Add the library to your Rust project, and you can instantly enjoy the convenience of simplified Excel editing.

### Installation
To use Edit-xlsx in your Rust project, add the following to your Cargo.toml file:

```toml
[dependencies]
edit-xlsx = "0.3.9"
```

## Notice

If you encounter any issues or have questions while using Edit-xlsx, please don't hesitate to reach out. Feel free to create an issue on our issue tracker. Your feedback is valuable, and we are here to assist you!

## Usage

A simple example of usage is shown below, and you can see more examples in the [example](https://github.com/MortalreminderPT/edit-xlsx/tree/dev-0.3.0/examples) directory

```rust
fn main() -> WorkbookResult<()> {
    // Create a new workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.get_worksheet(1)?;
    // write some text
    WorkSheet::write(worksheet, "A1", "Hello")?;
    worksheet.write("B1", "World")?;
    worksheet.write("C1", "Rust")?;
    // Adjust font size
    let big = Format::default().set_size(32);
    worksheet.write_with_format("B1", "big text", &big)?;
    // Change font color
    let red = Format::default().set_color(FormatColor::RGB("00FF7777"));
    worksheet.write_with_format("C1", "red text", &red)?;
    // Change the font style
    let bold = red.set_bold();
    worksheet.write_with_format("D1", "red bold text", &bold)?;
    // Change font
    let font = Format::default().set_font("华文行楷");
    worksheet.write_with_format("E1", "你好", &font)?;
    // adjust the text align
    let left_top = Format::default().set_align(FormatAlignType::Left).set_align(FormatAlignType::Top);
    worksheet.write_with_format("A2", "left top", &left_top)?;
    // add borders
    let thin_border = Format::default().set_border(FormatBorderType::Thin);
    worksheet.write_with_format("B2", "bordered text", &thin_border)?;
    // add background
    let red_background = Format::default().set_background_color(FormatColor::RGB("00FF7777"));
    worksheet.write_with_format("C2", "red", &red_background)?;
    // add a number
    worksheet.write("D2", std::f64::consts::PI)?;
    // add a new worksheet and set a tab color
    let worksheet = workbook.add_worksheet_by_name("Other examples")?;
    worksheet.set_tab_color(&FormatColor::RGB("00FF9900")); // Orange
    // Set a background.
    worksheet.set_background("examples/pics/ferris.png")?;
    // Create a format to use in the merged range.
    let merge_format = Format::default()
        .set_bold()
        .set_border(FormatBorderType::Double)
        .set_align(FormatAlignType::Center)
        .set_align(FormatAlignType::VerticalCenter)
        .set_background_color(FormatColor::RGB("00ffff00"));
    // Merge cells.
    worksheet.merge_range_with_format("A1:C3", "Merged Range", &merge_format)?;
    // Add an image
    worksheet.insert_image("A4:C10", &"./examples/pics/rust.png");
    workbook.save_as("examples/hello_world.xlsx")?;
    Ok(())
}
```

## License
This library is licensed under the [MIT License](https://opensource.org/license/mit).

Feel free to reach out if you have any questions or encounter any issues. Happy coding with Edit-xlsx!
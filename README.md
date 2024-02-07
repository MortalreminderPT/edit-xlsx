# edit-xlsx

The `edit-xlsx` library is a Rust library for editing Excel files in the xlsx format.

## Notice

When you're having trouble with `edit-xlsx`, please don't hesitate to submit an issue to give feedback on the problem you're having!

## Examples

A simple example of usage is shown below, and you can see more examples in the [example](https://github.com/MortalreminderPT/edit-xlsx/tree/dev-0.2.0/examples) directory

```rust
fn main() -> WorkbookResult<()> {
    let mut workbook = Workbook::from_path("examples/xlsx/edit_cell.xlsx");
    let mut worksheet = workbook.get_worksheet(1)?;
    // write some text
    worksheet.write(1, 1, "Hello")?;
    worksheet.write(1, 2, "World")?;
    worksheet.write(1, 3, "Rust")?;
    // Adjust font size
    let big = Format::new().set_size(32);
    worksheet.write_with_format(2, 1, "big text", &big)?;
    // Change font color
    let red = Format::new().set_color(Color::RGB("00FF7777"));
    worksheet.write_with_format(3, 1, "red text", &red)?;
    // Change the font style
    let bold = red.set_bold();
    worksheet.write_with_format(4, 1, "red bold text", &bold)?;
    // add a new worksheet
    let mut worksheet = workbook.add_worksheet().unwrap();
    // adjust the text align
    let left_top = Format::new().set_align(FormatAlign::Left).set_align(FormatAlign::Top);
    worksheet.write_with_format(1, 1, "left top", &left_top)?;
    // add borders
    let thin_border = Format::new().set_border(FormatBorder::Thin);
    worksheet.write_with_format(2, 1, "bordered text", &thin_border)?;
    // add background
    let red_background = Format::new().set_background_color(Color::RGB("00FF7777"));
    worksheet.write_with_format(3, 1, "red", &red_background)?;
    // add a number
    worksheet.write(4, 1, 3.14159)?;
    workbook.save_as("examples/output/edit_cell.xlsx")?;
    Ok(())
}
```
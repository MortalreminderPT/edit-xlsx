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
    workbook.save()?;
    Ok(())
}
```
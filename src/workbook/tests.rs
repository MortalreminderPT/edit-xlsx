#[cfg(test)]
mod tests {
    use crate::workbook::Workbook;

    #[test]
    fn test_from() {
        let wb = Workbook::from_path("D:/Github/edit-xlsx/2024-01-16 - 副本1.xlsx");
    }

    #[test]
    fn test_sheet() {
        let mut workbook = Workbook::from_path("2024-01-16 - 副本).xlsx");
        let worksheet = workbook.get_mut_sheet(0).unwrap();
        worksheet.write(0, 0, "Hello");
        workbook.save("2024-01-16 - 副本).xlsx").unwrap();
    }
}
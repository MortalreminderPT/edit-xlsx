pub(crate) fn to_col_name(mut col_num: u32) -> String {
    let mut col = String::new();
    while col_num > 0 {
        let pop = (col_num - 1) % 26;
        col_num = (col_num - 1) / 26;
        col.push(char::from_u32(65 + pop).unwrap());
    }
    col.chars().rev().collect::<String>()
}

pub(crate) fn to_col(col: &str) -> u32 {
    let mut col = col.as_bytes();
    let mut num = 0;
    while col.len() > 0 {
        if col[0] > 64 && col[0] < 91 {
            num *= 26;
            num += (col[0] - 64) as u32;
        }
        col = &col[1..];
    }
    num
}

pub(crate) fn to_ref(row: u32, col: u32) -> String {
    format!("{}{}", to_col_name(col), row)
}

pub(crate) fn to_loc(loc_ref: &str) -> (u32, u32) {
    let row = loc_ref.chars().filter(|&c| { c >= '0' && c <= '9' }).collect::<String>();
    let col = loc_ref.chars().filter(|&c| { c >= 'A' && c <= 'Z' }).collect::<String>();
    (row.parse().unwrap(), to_col(&col))
}

#[test]
fn test_col () {
    for i in 1..5_000_000 {
        let s = to_col_name(i);
        let r = to_col(&s);
        assert_eq!(i, r)
    }
}

#[test]
fn test_to_loc() {
    for row in 1..1000 {
        for col in 1..1000 {
            let loc_ref = to_ref(row, col);
            let (r, c) = to_loc(&loc_ref);
            assert_eq!(r, row);
            assert_eq!(c, col);
        }
    }
}
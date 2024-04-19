pub fn tab_header(widths: &Vec<f64>) -> String {
    let mut out = "[cols=\"".to_owned();

    for w in widths.iter().enumerate() {
        let w100 = w.1 * 100.0;
        let w100 = w100.round();
        let w100 = w100 as u32;
        out += &format!("{}", w100);
        if w.0 < widths.len() - 1 {
            out += ", ";
        };
    }
    out + "\"]\r"
}

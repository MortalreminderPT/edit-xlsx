fn hello() {
    println!("hello");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hello() {
        hello();
    }
}
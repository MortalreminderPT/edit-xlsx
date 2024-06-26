use std::ops::Add;
use crate::api::format::{FormatFont, FormatColor};

#[derive(Clone, Debug, Default, PartialEq)]
pub struct RichText {
    pub words: Vec<Word>
}

impl RichText {
    fn new() -> RichText {
        RichText {
            words: vec![],
        }
    }

    pub fn new_word(text: &str, font: &FormatFont) -> RichText {
        RichText {
            words: vec![Word::new(text, font)],
        }
    }
}

impl Add<Word> for RichText {
    type Output = RichText;

    fn add(mut self, rhs: Word) -> Self::Output {
        self.words.push(rhs);
        self
    }
}

impl Add<&RichText> for RichText {
    type Output = RichText;

    fn add(mut self, rhs: &RichText) -> Self::Output {
        self.words.extend(rhs.words.clone());
        self
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Word {
    pub text: String,
    pub font: Option<FormatFont>,
}

impl Word {
    pub fn new(text: &str, font: &FormatFont) -> Word {
        Word {
            text: text.to_string(),
            font: None,
        }
    }
}


#[test]
fn test() {
    let mut font1 = FormatFont::default();
    font1.bold = true;
    let mut font2 = FormatFont::default();
    font2.italic = true;
    let mut rich_text = RichText::new_word("a", &font1);
    let word = Word::new("b", &font2);
    rich_text = rich_text + word;
    println!("{:?}", rich_text);
}
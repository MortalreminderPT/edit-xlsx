use serde::{Deserialize, Serialize};
use crate::xml::style::font::Font;
use crate::api::cell::rich_text::{RichText as ApiRichText, Word};
use crate::xml::common::FromFormat;

#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub(crate) struct InlineString {
    #[serde(rename = "r", skip_serializing_if = "Vec::is_empty")]
    rich_texts: Vec<RichText>
}

#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub(crate) struct RichText {
    #[serde(rename = "rPr", skip_serializing_if = "Option::is_none")]
    font: Option<Font>,
    #[serde(rename = "t", skip_serializing_if = "String::is_empty")]
    text: String,
}

impl FromFormat<ApiRichText> for InlineString {
    fn set_attrs_by_format(&mut self, api_rich_text: &ApiRichText) {
        self.rich_texts = api_rich_text.words
            .iter()
            .map(|w| RichText::from_format(w))
            .collect();
    }

    fn set_format(&self, api_rich_text: &mut ApiRichText) {
        api_rich_text.words = self.rich_texts
            .iter()
            .map(|r|r.get_format())
            .collect();
    }
}

impl FromFormat<Word> for RichText {
    fn set_attrs_by_format(&mut self, word: &Word) {
        self.text = word.text.clone();
        self.font = Some(Font::from_format(&word.font));
    }

    fn set_format(&self, word: &mut Word) {
        word.text = self.text.clone();
        word.font = self.font.clone().unwrap().get_format();
    }
}
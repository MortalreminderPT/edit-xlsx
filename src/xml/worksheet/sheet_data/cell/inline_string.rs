use serde::{Deserialize, Serialize};
use crate::xml::style::font::Font;
use crate::api::cell::rich_text::{RichText as ApiRichText, Word};
use crate::xml::common::FromFormat;
use crate::xml::worksheet::sheet_data::cell::text::Text;

#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub(crate) struct InlineString {
    #[serde(rename = "r", skip_serializing_if = "Vec::is_empty")]
    rich_texts: Vec<RichText>
}

#[derive(Debug, Clone, Default, Deserialize, Serialize)]
pub(crate) struct RichText {
    #[serde(rename = "rPr", skip_serializing_if = "Option::is_none")]
    font: Option<Font>,
    #[serde(rename = "t", skip_serializing_if = "Option::is_none")]
    text: Option<Text>,
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
        self.text = Some(Text::new_with_space(&word.text));
        self.font = Some(Font::from_rich_font_format(&word.font));
    }

    fn set_format(&self, word: &mut Word) {
        word.text = self.text.clone().unwrap().text;
        word.font = self.font.clone().unwrap().get_rich_font_format();
    }
}
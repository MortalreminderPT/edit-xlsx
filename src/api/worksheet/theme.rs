use crate::WorkSheet;
use crate::api::theme::Theme;

impl WorkSheet {
    pub fn get_theme(&self, theme_id: u32) -> Theme {
        let binding = self.themes.borrow();
        let theme = binding.themes.get(theme_id as usize).unwrap();
        theme.to_api_theme()
    }
}
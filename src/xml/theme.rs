use serde::{Deserialize, Serialize};
use crate::xml::style::color::Color;
use crate::api::theme::Theme as ApiTheme;
use crate::FormatColor;
use crate::xml::common::FromFormat;

#[derive(Debug, Default)]
pub(crate) struct Themes {
    pub(crate) themes: Vec<Theme>
}

impl Themes {
    pub(crate) fn add_theme(&mut self, theme: Theme) {
        self.themes.push(theme);
    }
}

#[derive(Debug, Default, Serialize, Deserialize)]
#[serde(rename(serialize = "a:theme", deserialize = "theme"))]
pub(crate) struct Theme {
    #[serde(rename(serialize = "a:themeElements", deserialize = "themeElements"), default)]
    theme_elements: ThemeElements,
}

impl Theme {
    pub(crate) fn to_api_theme(&self) -> ApiTheme {
        let clr = &self.theme_elements.clr_theme.clr;
        let colors_rgb = clr.iter()
            .map(|clr| clr.to_color())
            .map(|c| c.get_format())
            .collect::<Vec<FormatColor>>();
        ApiTheme::new(colors_rgb)
    }

    pub(crate) fn get_color_rgb(&self, color_theme: u32) -> Color {
        let clr = &self.theme_elements.clr_theme.clr;
        if let Some(clr) = clr.get(color_theme as usize) {
            clr.to_color()
        } else {
            Color::default()
        }
    }
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct ThemeElements {
    #[serde(rename(serialize = "a:clrScheme", deserialize = "clrScheme"), default)]
    clr_theme: ClrTheme,
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct ClrTheme {
    #[serde(rename = "$value")]
    clr: Vec<Clr>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct Clr {
    #[serde(rename(serialize = "a:sysClr", deserialize = "sysClr"), default)]
    sys_clr: Option<SysClr>,
    #[serde(rename(serialize = "a:srgbClr", deserialize = "srgbClr"), default)]
    srgb_clr: Option<SrgbClr>,
}

impl Clr {
    fn to_color(&self) -> Color {
        match self {
            Clr {
                sys_clr: Some(sys_clr),
                srgb_clr: None,
            } => {
                sys_clr.to_color_rgb()
            }
            Clr {
                sys_clr: None,
                srgb_clr: Some(srgb_clr),
            } => {
                srgb_clr.to_color_rgb()
            }
            _ => {
                Color::default()
            }
        }
    }
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct SysClr {
    #[serde(rename = "@val", default)]
    val: String,
    #[serde(rename = "@lastClr", default)]
    last_clr: String,
}

impl SysClr {
    fn to_color_rgb(&self) -> Color {
        Color::from_rgb_hex(&self.last_clr)
    }
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct SrgbClr {
    #[serde(rename = "@val", default)]
    val: String,
}

impl SrgbClr {
    fn to_color_rgb(&self) -> Color {
        Color::from_rgb_hex(&self.val)
    }
}
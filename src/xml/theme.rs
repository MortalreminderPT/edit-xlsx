use serde::{Deserialize, Serialize};

#[derive(Debug, Default)]
pub(crate) struct Themes {
    themes: Vec<Theme>
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

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct SysClr {
    #[serde(rename = "@val", default)]
    val: String,
    #[serde(rename = "@lastClr", default)]
    last_clr: String,
}

#[derive(Debug, Default, Serialize, Deserialize)]
pub(crate) struct SrgbClr {
    #[serde(rename = "@val", default)]
    val: String,
}
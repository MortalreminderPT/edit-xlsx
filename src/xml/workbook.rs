use std::io;
use std::path::Path;
use quick_xml::{de, se};
use serde::{Deserialize, Serialize};
use crate::file::{XlsxFileReader, XlsxFileType, XlsxFileWriter};
use crate::result::{SheetError, WorkbookError};
use crate::WorkbookResult;
use crate::xml::common::{ExtLst, XmlnsAttrs};
use crate::xml::manage::Io;

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename="workbook")]
pub(crate) struct Workbook {
    #[serde(flatten)]
    xmlns_attrs: XmlnsAttrs,
    #[serde(rename = "fileVersion")]
    file_version: FileVersion,
    #[serde(rename = "fileSharing", skip_serializing_if = "Option::is_none")]
    pub(crate) file_sharing: Option<FileSharing>,
    #[serde(rename = "workbookPr")]
    workbook_pr: WorkbookPr,
    #[serde(rename(serialize = "xr:revisionPtr", deserialize = "revisionPtr"), skip_serializing_if = "Option::is_none")]
    xr_revision_ptr: Option<XrRevisionPtr>,
    #[serde(rename = "bookViews")]
    pub(crate) book_views: BookViews,
    #[serde(rename = "sheets")]
    pub(crate) sheets: Sheets,
    #[serde(rename = "calcPr")]
    calc_pr: CalcPr,
    #[serde(rename = "extLst", skip_serializing_if = "Option::is_none")]
    ext_lst: Option<ExtLst>,
}

impl Workbook {
    pub(crate) fn add_worksheet(&mut self) -> WorkbookResult<(u32, String)> {
        let id = 1 + self.sheets.sheets.iter().max_by_key(|s| { s.sheet_id }).unwrap().sheet_id;
        let name = format!("Sheet{id}");
        if let Some(_) = self.sheets.sheets.iter().find(|s| { s.name == name }) {
            return Err(WorkbookError::SheetError(SheetError::DuplicatedSheets));
        }
        self.sheets.sheets.push(Sheet::by_name(id, &name));
        Ok((id, name))
    }

    pub(crate) fn add_worksheet_by_name(&mut self, name: &str) -> WorkbookResult<u32> {
        let id = 1 + self.sheets.sheets.iter().max_by_key(|s| { s.sheet_id }).unwrap().sheet_id;
        if let Some(_) = self.sheets.sheets.iter().find(|s| { s.name == name }) {
            return Err(WorkbookError::SheetError(SheetError::DuplicatedSheets));
        }
        self.sheets.sheets.push(Sheet::by_name(id, name));
        Ok(id)
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct FileVersion {
    #[serde(rename = "@appName")]
    app_name: String,
    #[serde(rename = "@lastEdited", skip_serializing_if = "Option::is_none")]
    last_edited: Option<u32>,
    #[serde(rename = "@lowestEdited", skip_serializing_if = "Option::is_none")]
    lowest_edited: Option<u32>,
    #[serde(rename = "@rupBuild", skip_serializing_if = "Option::is_none")]
    rup_build: Option<String>,
}

impl Default for FileVersion {
    fn default() -> Self {
        FileVersion {
            app_name: "xl".to_string(),
            last_edited: None,
            lowest_edited: None,
            rup_build: Some(String::from("14420")),
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct FileSharing {
    #[serde(rename = "@readOnlyRecommended")]
    pub(crate) read_only_recommended: u8,
}

impl Default for FileSharing {
    fn default() -> Self {
        Self {
            read_only_recommended: 0,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct WorkbookPr {
    #[serde(rename = "@filterPrivacy", skip_serializing_if = "Option::is_none")]
    filter_privacy: Option<u32>,
    #[serde(rename = "@defaultThemeVersion")]
    default_theme_version: String,
}

impl Default for WorkbookPr {
    fn default() -> Self {
        WorkbookPr {
            filter_privacy: None,
            default_theme_version: String::from("164011"),
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct BookViews {
    #[serde(rename = "workbookView")]
    pub(crate) book_views: Vec<WorkbookView>
}

impl Default for BookViews {
    fn default() -> Self {
        BookViews {
            book_views: vec![Default::default()],
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct WorkbookView {
    #[serde(rename = "@xWindow")]
    x_window: u32,
    #[serde(rename = "@yWindow")]
    y_window: u32,
    #[serde(rename = "@windowWidth")]
    pub(crate) window_width: u32,
    #[serde(rename = "@windowHeight")]
    pub(crate) window_height: u32,
    #[serde(rename = "@tabRatio", skip_serializing_if = "Option::is_none")]
    pub(crate) tab_ratio: Option<u32>,
    #[serde(rename = "@activeTab", skip_serializing_if = "Option::is_none")]
    pub(crate) active_tab: Option<u32>
}

impl Default for WorkbookView {
    fn default() -> Self {
        WorkbookView {
            x_window: 0,
            y_window: 0,
            window_width: 22260,
            window_height: 12645,
            tab_ratio: None,
            active_tab: None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Sheets {
    #[serde(rename = "sheet")]
    pub(crate) sheets: Vec<Sheet>
}

impl Default for Sheets {
    fn default() -> Self {
        Sheets {
            sheets: vec![],
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Sheet {
    #[serde(rename = "@name")]
    pub(crate) name: String,
    #[serde(rename = "@sheetId")]
    pub(crate) sheet_id: u32,
    #[serde(rename(serialize = "@r:id", deserialize = "@id"))]
    pub(crate) r_id: String,
    #[serde(rename = "@state", skip_serializing_if = "Option::is_none")]
    pub(crate) state: Option<String>,
}

impl Default for Sheet {
    fn default() -> Sheet {
        Sheet {
            name: format!("sheet1"),
            sheet_id: 1,
            r_id: format!("rId1"),
            state: None,
        }
    }
}

impl Sheet {
    pub(crate) fn by_id(id: u32) -> Sheet {
        Sheet {
            name: format!("Sheet{id}"),
            sheet_id: id,
            r_id: format!("rId{id}"),
            state: None,
        }
    }

    pub(crate) fn by_name(id: u32, name: &str) -> Sheet {
        Sheet {
            name: String::from(name),
            sheet_id: id,
            r_id: format!("rId{id}"),
            state: None,
        }
    }

    pub(crate) fn change_id(&mut self, id: u32) {
        self.sheet_id = id;
        self.r_id = format!("rId{id}");
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct CalcPr {
    #[serde(rename = "@calcId")]
    calc_id: String,
}

impl Default for CalcPr {
    fn default() -> Self {
        CalcPr {
            calc_id: String::from("162913"),
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct XrRevisionPtr {
    #[serde(rename = "@revIDLastSave", skip_serializing_if = "Option::is_none")]
    rev_id_last_save: Option<u32>,
    #[serde(rename = "@documentId", skip_serializing_if = "Option::is_none")]
    document_id: Option<String>,
    #[serde(rename(serialize = "@xr6:coauthVersionLast", deserialize = "@coauthVersionLast"), skip_serializing_if = "Option::is_none")]
    xr6_coauth_version_last: Option<u32>,
    #[serde(rename(serialize = "@xr6:coauthVersionMax", deserialize = "@coauthVersionMax"), skip_serializing_if = "Option::is_none")]
    xr6_coauth_version_max: Option<u32>,
    #[serde(rename(serialize = "@xr10:uidLastSave", deserialize = "@uidLastSave"), skip_serializing_if = "Option::is_none")]
    xr10_uid_last_save: Option<String>,
}

impl Default for Workbook {
    fn default() -> Self {
        Workbook {
            xmlns_attrs: XmlnsAttrs::workbook_default(),
            file_version: Default::default(),
            file_sharing: None,
            workbook_pr: Default::default(),
            xr_revision_ptr: None,
            book_views: Default::default(),
            sheets: Default::default(),
            calc_pr: Default::default(),
            ext_lst: Some(Default::default()),
        }
    }
}

impl Io<Workbook> for Workbook {
    fn from_path<P: AsRef<Path>>(file_path: P) -> io::Result<Workbook> {
        let mut file = XlsxFileReader::from_path(file_path, XlsxFileType::WorkbookFile)?;
        let mut xml = String::new();
        file.read_to_string(&mut xml).unwrap();
        let work_book = de::from_str(&xml).unwrap();
        Ok(work_book)
    }

    fn save<P: AsRef<Path>>(&mut self, file_path: P) {
        let xml = se::to_string_with_root("workbook", &self).unwrap();
        let xml = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}", xml);
        let mut file = XlsxFileWriter::from_path(file_path, XlsxFileType::WorkbookFile).unwrap();
        file.write_all(xml.as_ref()).unwrap();
    }
}
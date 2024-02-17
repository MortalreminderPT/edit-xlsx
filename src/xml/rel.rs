use serde::{Deserialize, Serialize};

const SHEET_TYPE_STRING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
#[derive(Debug, Deserialize, Serialize)]
pub(crate) struct Relationships {
    #[serde(rename = "@xmlns")]
    xmlns: String,
    #[serde(rename = "Relationship")]
    relationship: Vec<RelationShip>,
    #[serde(skip)]
    last_r_id: Option<u32>,
    #[serde(skip)]
    r_id_offset: Option<u32>,
}

impl Relationships {
    fn add_worksheet(&mut self) -> u32 {
        let last_r_id = match self.last_r_id {
            Some(last_r_id) => last_r_id,
            None => {
                self.relationship
                    .iter()
                    .filter(|rel| { rel.target.starts_with("worksheets") })
                    .map(|r| r.id[3..].parse::<u32>().unwrap())
                    .max()
                    .unwrap_or(0)
            }
        };
        let id = last_r_id + 1;
        self.relationship.push(
            RelationShip::new_sheet(id)
        );
        id
    }
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
struct RelationShip {
    #[serde(rename = "@Id")]
    id: String,
    #[serde(rename = "@Type")]
    rel_type: String,
    #[serde(rename = "@Target")]
    target: String,
}

impl RelationShip {
    fn new_sheet(id: u32) -> RelationShip {
        RelationShip {
            id: format!("rId{id}"),
            rel_type: String::from(SHEET_TYPE_STRING),
            target: format!("worksheets/sheet{id}.xml"),
        }
    }
}
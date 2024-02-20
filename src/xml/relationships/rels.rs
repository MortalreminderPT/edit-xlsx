use crate::xml::relationships::RelationShip;

enum RelsType {
    Sheet,
    Theme,
    Style,
}

struct Rels {
    rels_type: RelsType,
    rels: Vec<RelationShip>,
}
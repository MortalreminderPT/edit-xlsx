use crate::xml::relationship::RelationShip;

enum RelsType {
    Sheet,
    Theme,
    Style,
}

struct Rels {
    rels_type: RelsType,
    rels: Vec<RelationShip>,
}
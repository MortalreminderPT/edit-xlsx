pub(crate) struct Relationships {
    xmlns: String,
    relationship: Vec<RelationShip>
}

struct RelationShip {
    id: String,
    rel_type: String,
    target: String,
}

impl Relationships {

}